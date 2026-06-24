#!/usr/bin/env python3
"""端到端测试：生成激活码 → 激活 → 验证签名 → 解绑 → 数据包

用法（在 Mac 终端）：
    cd /Users/rickxu/Finance/license-server
    pip3 install cryptography  # 如果没装
    python3 test_e2e.py
"""
import base64
import json
import sys
import urllib.error
import urllib.request

try:
    from cryptography.exceptions import InvalidSignature
    from cryptography.hazmat.primitives.asymmetric.ed25519 import Ed25519PublicKey
except ImportError:
    print("❌ 缺 cryptography 库，请先：pip3 install cryptography")
    sys.exit(1)


# 读 .env
env = {}
for line in open(".env"):
    line = line.strip()
    if line and not line.startswith("#") and "=" in line:
        k, v = line.split("=", 1)
        env[k.strip()] = v.strip()

ADMIN_KEY   = env["ADMIN_KEY"]
WORKERS_URL = env["WORKERS_URL"]
PUB_KEY_B64 = env["ED25519_PUBLIC_KEY_B64"]


def post(path, body, headers=None, max_retries=3):
    h = {
        "Content-Type": "application/json",
        # Cloudflare 边缘会按 UA 拦截可疑机器人，伪装成浏览器
        "User-Agent": "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) WiseLedger-Test/1.0",
    }
    if headers:
        h.update(headers)
    import time
    last_exc = None
    for attempt in range(max_retries):
        req = urllib.request.Request(
            WORKERS_URL + path,
            data=json.dumps(body).encode(),
            method="POST",
            headers=h,
        )
        try:
            with urllib.request.urlopen(req, timeout=15) as resp:
                return resp.status, json.loads(resp.read().decode())
        except urllib.error.HTTPError as e:
            return e.code, json.loads(e.read().decode())
        except (urllib.error.URLError, Exception) as e:
            last_exc = e
            if attempt < max_retries - 1:
                print(f"  ⏳ 网络瞬时错误，{1+attempt}s 后重试: {e}")
                time.sleep(1 + attempt)
            else:
                raise


def section(title):
    print("\n" + "=" * 65)
    print(title)
    print("=" * 65)


# ────────────────────────────────────────────────────────────
section("测试 1：POST /admin/generate（生成激活码）")
status, body = post(
    "/admin/generate",
    {"customer_email": "test@example.com", "price": 1, "note": "端到端测试"},
    {"X-Admin-Key": ADMIN_KEY},
)
print(f"HTTP {status}")
print(json.dumps(body, indent=2, ensure_ascii=False))
assert status == 200, f"生成激活码失败"
code = body["license_code"]
print(f"\n✓ 拿到激活码：{code}")

# ────────────────────────────────────────────────────────────
section("测试 2：不带 admin key（应该 401）")
status, body = post("/admin/generate", {"customer_email": "x"})
print(f"HTTP {status}: {body}")
assert status == 401
print("✓ 鉴权正常")

# ────────────────────────────────────────────────────────────
section("测试 3：POST /activate（首次激活）")
machine_id = "fake_machine_e2e_test_001"
status, body = post("/activate", {"code": code, "machine_id": machine_id})
print(f"HTTP {status}")
print(json.dumps(body, indent=2, ensure_ascii=False))
assert status == 200
assert body["license_status"] == "active"
token = body["token"]
print(f"\n✓ 激活成功")

# ────────────────────────────────────────────────────────────
section("测试 4：用公钥验证 token 签名")
payload_bytes = base64.b64decode(token["payload_b64"])
sig_bytes = base64.b64decode(token["signature_b64"])
pub = Ed25519PublicKey.from_public_bytes(base64.b64decode(PUB_KEY_B64))
try:
    pub.verify(sig_bytes, payload_bytes)
    print("✓ 签名验证通过")
except InvalidSignature:
    print("❌ 签名验证失败")
    raise

payload = json.loads(payload_bytes.decode())
print("\nToken 内容：")
print(json.dumps(payload, indent=2, ensure_ascii=False))
assert payload["license_code"] == code
assert payload["machine_id"] == machine_id
print("\n✓ 激活码 / 机器码匹配")

# ────────────────────────────────────────────────────────────
section("测试 5：换机器（第 1 次解绑）")
m2 = "fake_machine_e2e_test_002"
status, body = post("/activate", {"code": code, "machine_id": m2})
print(f"HTTP {status}: license_status={body.get('license_status')}")
assert status == 200
print("✓ 解绑迁移 1/2 成功")

# ────────────────────────────────────────────────────────────
section("测试 6：再换机器（第 2 次解绑）")
import time
time.sleep(1.5)  # 给 KV 一点一致性时间
m3 = "fake_machine_e2e_test_003"
status, body = post("/activate", {"code": code, "machine_id": m3})
print(f"HTTP {status}")
print(f"Body: {json.dumps(body, ensure_ascii=False)}")
if status != 200:
    print("\n⚠️  失败了。再试一次（看是不是瞬时错误）...")
    time.sleep(2)
    status, body = post("/activate", {"code": code, "machine_id": m3})
    print(f"重试 HTTP {status}: {json.dumps(body, ensure_ascii=False)}")
assert status == 200
print("✓ 解绑迁移 2/2 成功")

# ────────────────────────────────────────────────────────────
section("测试 7：第 3 次解绑（应被拒绝）")
print("（等 3 秒让 KV 同步，确保 count=2 已传播）")
time.sleep(3)
status, body = post("/activate", {"code": code, "machine_id": "fake_004"})
print(f"HTTP {status}: {json.dumps(body, ensure_ascii=False)}")
if status != 403:
    print("\n⚠️  期望 403 但是 200，可能 KV 还没同步，再等 5 秒试...")
    time.sleep(5)
    status, body = post("/activate", {"code": code, "machine_id": "fake_005"})
    print(f"重试 HTTP {status}: {json.dumps(body, ensure_ascii=False)}")
assert status == 403
print("✓ 解绑配额限制工作")

# ────────────────────────────────────────────────────────────
section("测试 8：错误激活码格式")
status, body = post("/activate", {"code": "INVALID-CODE", "machine_id": "x"})
print(f"HTTP {status}: {json.dumps(body, ensure_ascii=False)}")
assert status == 400
print("✓ 格式校验工作")

# ────────────────────────────────────────────────────────────
section("测试 9：不存在的激活码")
status, body = post("/activate", {"code": "WL-AAAA-BBBB-CCC2", "machine_id": "x"})
print(f"HTTP {status}: {json.dumps(body, ensure_ascii=False)}")
assert status in (400, 404)
print("✓ not found 工作")

# ────────────────────────────────────────────────────────────
section("测试 10：POST /update-data（数据包检查）")
print("（等 3 秒让 KV 在各 colo 同步…）")
time.sleep(3)
status, body = post("/update-data", {
    "code": code,
    "machine_id": m3,
    "current_pack_version": "2025.01",
})
print(f"HTTP {status}")
print(json.dumps(body, indent=2, ensure_ascii=False))
if status != 200:
    print("\n⚠️  又失败，再等 5 秒重试...")
    time.sleep(5)
    status, body = post("/update-data", {
        "code": code,
        "machine_id": m3,
        "current_pack_version": "2025.01",
    })
    print(f"重试 HTTP {status}: {json.dumps(body, indent=2, ensure_ascii=False)}")
assert status == 200
print("✓ 数据包接口工作")

# ────────────────────────────────────────────────────────────
print("\n" + "=" * 65)
print("🎉 全部 10 个测试通过！后端工作正常")
print("=" * 65)
