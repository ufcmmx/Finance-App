#!/usr/bin/env python3
"""调试：直接看 Worker 返回啥"""
import json
import urllib.error
import urllib.request

# 读 .env
env = {}
for line in open(".env"):
    line = line.strip()
    if line and not line.startswith("#") and "=" in line:
        k, v = line.split("=", 1)
        env[k.strip()] = v.strip()

ADMIN_KEY   = env["ADMIN_KEY"]
WORKERS_URL = env["WORKERS_URL"]


def raw_post(path, body_dict, headers=None):
    h = {"Content-Type": "application/json"}
    if headers:
        h.update(headers)
    req = urllib.request.Request(
        WORKERS_URL + path,
        data=json.dumps(body_dict).encode(),
        method="POST",
        headers=h,
    )
    try:
        with urllib.request.urlopen(req, timeout=15) as resp:
            return resp.status, dict(resp.headers), resp.read().decode("utf-8", errors="replace")
    except urllib.error.HTTPError as e:
        return e.code, dict(e.headers), e.read().decode("utf-8", errors="replace")
    except Exception as e:
        return None, {}, f"<连接异常: {e}>"


print("Worker URL:", WORKERS_URL)
print()

print("=" * 65)
print("测试 1：根路径 GET（看部署有没有起来）")
print("=" * 65)
try:
    req = urllib.request.Request(WORKERS_URL + "/", method="GET")
    with urllib.request.urlopen(req, timeout=10) as resp:
        print(f"HTTP {resp.status}")
        print(resp.read().decode())
except urllib.error.HTTPError as e:
    print(f"HTTP {e.code}")
    print(e.read().decode("utf-8", errors="replace"))
except Exception as e:
    print(f"连接失败: {e}")

print()
print("=" * 65)
print("测试 2：POST /admin/generate（带 admin key）")
print("=" * 65)
status, headers, body = raw_post(
    "/admin/generate",
    {"customer_email": "test@example.com", "price": 1},
    {"X-Admin-Key": ADMIN_KEY},
)
print(f"HTTP {status}")
print(f"Headers: {headers}")
print(f"Body:\n{body}")

print()
print("=" * 65)
print("测试 3：POST /admin/generate（不带 admin key，应该 401）")
print("=" * 65)
status, headers, body = raw_post("/admin/generate", {"customer_email": "x"})
print(f"HTTP {status}")
print(f"Body:\n{body}")
