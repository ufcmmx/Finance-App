"""license.py — 客户端激活授权模块

职责：
1. 取机器特征码（Windows: CPU + 主板 + 硬盘；Mac: 开发用固定值）
2. 调用 Workers /activate 完成激活
3. 用嵌入的 Ed25519 公钥验本地 token 签名
4. 启动时检查激活状态
5. 调用 /update-data 检查年度数据包

数据存放：
  Windows: %APPDATA%/WiseLedger/license.json
  macOS:   ~/Library/Application Support/WiseLedger/license.json （开发用）

Mac 开发：设环境变量 WL_DEV=1 可用固定假 machine_id 调试。

将来用 Cython 编译为 .pyd 以增加反编译难度。
"""
from __future__ import annotations   # 让 dict | None 等新语法在 Py3.9 也能用
import base64
import hashlib
import json
import os
import platform
import re
import subprocess
import sys
import urllib.error
import urllib.request
from pathlib import Path

# ─── 嵌入式常量（不敏感，公开即可） ─────────────────────────────────
WORKERS_URL = "https://license.wisdompluscn.com"

# Ed25519 公钥（base64 编码的 32 字节裸密钥）— 仅能验签，不能签名
ED25519_PUBLIC_KEY_B64 = "mQlP5F1dYSjWWR7+QoL/jO1v9n2/S4AiIiyQeBfScQ0="

# 客户端 User-Agent（绕过 Cloudflare 默认 Python UA 拦截）
USER_AGENT = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) WiseLedger/1.0"

# 数据包当前版本（升级软件时跟着改）
CURRENT_DATA_PACK_VERSION = "2026.01"


# ─── License 状态常量 ────────────────────────────────────────────
# is_activated() 返回的 status 枚举
STATUS_ACTIVE       = "active"        # 正常使用中
STATUS_TRIAL        = "trial"         # 试用中
STATUS_GRACE        = "grace"         # 订阅已过期但在宽限期
STATUS_READONLY     = "readonly"      # 订阅已过期且超过宽限期（只读）
STATUS_EXPIRED      = "expired"       # 试用/订阅已过期（无宽限期）
STATUS_NOT_ACTIVE   = "not_activated" # 未激活


# ─── 工具：错误码 ─────────────────────────────────────────────────
class LicenseError(Exception):
    """激活相关错误的基类，msg 中应包含用户可读说明"""


# ─── 工具：ISO 时间处理 ─────────────────────────────────────────
def _parse_iso(s: str):
    """把 ISO 8601 字符串解析成 datetime（UTC 时区）"""
    from datetime import datetime, timezone
    if not s:
        return None
    # Python 3.11+ 支持 Z 后缀，3.10 及以下要替换
    s = s.replace("Z", "+00:00")
    try:
        dt = datetime.fromisoformat(s)
        if dt.tzinfo is None:
            dt = dt.replace(tzinfo=timezone.utc)
        return dt
    except Exception:
        return None


def _now_utc():
    from datetime import datetime, timezone
    return datetime.now(timezone.utc)


# ─── 1. 机器特征码 ─────────────────────────────────────────────────
def _run_wmic(field_cmd: str) -> str:
    """跑一个 wmic 命令，返回去除空白后的第一个非空行（去掉表头）"""
    try:
        out = subprocess.check_output(
            field_cmd, shell=True, stderr=subprocess.DEVNULL, timeout=8
        ).decode("utf-8", errors="ignore")
    except Exception:
        return ""
    lines = [ln.strip() for ln in out.splitlines() if ln.strip()]
    # 第一行通常是表头，第二行是值
    return lines[1] if len(lines) >= 2 else ""


def _get_windows_machine_id() -> str:
    """Windows: CPU ProcessorId + 主板 UUID + 系统盘物理 SN，三者拼接 SHA256"""
    cpu_id = _run_wmic("wmic cpu get ProcessorId")
    mb_uuid = _run_wmic("wmic csproduct get UUID")
    disk_sn = _run_wmic("wmic diskdrive get SerialNumber")
    combined = f"{cpu_id}|{mb_uuid}|{disk_sn}"
    return hashlib.sha256(combined.encode()).hexdigest()


def _get_macos_dev_machine_id() -> str:
    """macOS 开发用：固定 SHA256，方便本地测试激活流程"""
    return hashlib.sha256(b"WL_MACOS_DEV_FIXED").hexdigest()


def get_machine_id() -> str:
    """统一入口：返回当前机器特征码 (SHA256 hex)。

    Mac 开发时设 WL_DEV=1 走固定 ID。
    """
    if os.environ.get("WL_DEV"):
        return _get_macos_dev_machine_id()
    system = platform.system()
    if system == "Windows":
        return _get_windows_machine_id()
    if system == "Darwin":
        # 没设 WL_DEV 的 mac 也走 dev id（避免崩，但不应该出现在生产）
        return _get_macos_dev_machine_id()
    raise LicenseError(f"不支持的操作系统：{system}")


# ─── 2. Token 本地存储 ────────────────────────────────────────────
def _get_token_path() -> Path:
    """跨平台返回 token 文件路径，目录不存在自动创建"""
    if platform.system() == "Windows":
        base = os.environ.get("APPDATA") or os.path.expanduser("~")
        d = Path(base) / "WiseLedger"
    else:  # macOS / Linux
        d = Path.home() / "Library" / "Application Support" / "WiseLedger"
    d.mkdir(parents=True, exist_ok=True)
    return d / "license.json"


def load_local_token() -> dict | None:
    """读本地 token，文件不存在返回 None"""
    path = _get_token_path()
    if not path.exists():
        return None
    try:
        with open(path, "r", encoding="utf-8") as f:
            return json.load(f)
    except Exception:
        return None


def save_local_token(token_envelope: dict) -> None:
    """token_envelope = {"payload_b64": ..., "signature_b64": ...}"""
    path = _get_token_path()
    with open(path, "w", encoding="utf-8") as f:
        json.dump(token_envelope, f, ensure_ascii=False, indent=2)


def remove_local_token() -> None:
    path = _get_token_path()
    if path.exists():
        path.unlink()


# ─── 3. Ed25519 验签 ──────────────────────────────────────────────
def _verify_signature(payload_bytes: bytes, signature_bytes: bytes) -> bool:
    """用嵌入的公钥验 Ed25519 签名，签名有效返回 True"""
    try:
        from cryptography.hazmat.primitives.asymmetric.ed25519 import (
            Ed25519PublicKey,
        )
        from cryptography.exceptions import InvalidSignature

        pub = Ed25519PublicKey.from_public_bytes(
            base64.b64decode(ED25519_PUBLIC_KEY_B64)
        )
        try:
            pub.verify(signature_bytes, payload_bytes)
            return True
        except InvalidSignature:
            return False
    except ImportError:
        raise LicenseError("缺少 cryptography 库，请先 pip install cryptography")


def _parse_and_verify_token(envelope: dict) -> dict | None:
    """从 envelope 解出 payload 并验签，验签通过返回 payload dict，否则 None"""
    try:
        payload_bytes = base64.b64decode(envelope["payload_b64"])
        sig_bytes = base64.b64decode(envelope["signature_b64"])
    except Exception:
        return None
    if not _verify_signature(payload_bytes, sig_bytes):
        return None
    try:
        return json.loads(payload_bytes.decode("utf-8"))
    except Exception:
        return None


# ─── 4. 激活状态检查 ──────────────────────────────────────────────
def get_license_status() -> dict:
    """完整版本的授权状态查询，返回所有 UI 决策所需信息。

    返回 dict 字段：
      status:            STATUS_* 常量之一
      message:           人类可读描述
      type:              None | 'trial' | 'annual' | 'permanent'
      license_code:      激活码（或 trial display code）
      expires_at:        ISO 时间字符串 或 None（永久）
      activated_at:      ISO 时间字符串
      days_remaining:    到期前剩余天数（负值表示已过期，None 表示永久）
      grace_days_left:   宽限期剩余天数（仅 annual 到期后有意义）
      readonly:          bool 是否只读模式
    """
    empty = {
        "status": STATUS_NOT_ACTIVE, "message": "未激活",
        "type": None, "license_code": None,
        "expires_at": None, "activated_at": None,
        "days_remaining": None, "grace_days_left": 0,
        "readonly": False,
    }

    envelope = load_local_token()
    if not envelope:
        return empty

    payload = _parse_and_verify_token(envelope)
    if not payload:
        return {**empty, "message": "本地激活信息已损坏或被篡改"}

    # 机器码匹配
    if payload.get("machine_id") != get_machine_id():
        return {**empty,
                "message": "激活信息与当前电脑不匹配（可能更换了硬件或复制了软件）"}

    lic_type = payload.get("type", "permanent")
    expires_at_str = payload.get("expires_at")
    grace_days = int(payload.get("grace_days", 0))
    code = payload.get("license_code", "")

    base = {
        "type": lic_type,
        "license_code": code,
        "activated_at": payload.get("activated_at"),
        "expires_at": expires_at_str,
        "grace_days_left": 0,
        "readonly": False,
    }

    # 永久版：永远有效
    if lic_type == "permanent" or expires_at_str is None:
        return {**base, "status": STATUS_ACTIVE, "days_remaining": None,
                "message": f"永久版：{code}"}

    # trial / annual：需要判断是否过期
    expires_at = _parse_iso(expires_at_str)
    if expires_at is None:
        return {**base, "status": STATUS_ACTIVE, "days_remaining": None,
                "message": f"已激活：{code}"}

    now = _now_utc()
    delta_days = (expires_at - now).total_seconds() / 86400

    if delta_days > 0:
        # 未到期
        status = STATUS_TRIAL if lic_type == "trial" else STATUS_ACTIVE
        msg_prefix = "试用中" if lic_type == "trial" else "订阅版"
        return {**base, "status": status, "days_remaining": int(delta_days),
                "message": f"{msg_prefix}：剩余 {int(delta_days)} 天"}

    # 已到期
    if lic_type == "trial":
        # 试用无宽限
        return {**base, "status": STATUS_EXPIRED, "days_remaining": int(delta_days),
                "readonly": True,
                "message": f"试用已到期（{-int(delta_days)} 天前）"}

    # 订阅到期 —— 检查宽限期
    grace_ends_days = delta_days + grace_days
    if grace_ends_days > 0:
        # 在宽限期内
        return {**base, "status": STATUS_GRACE,
                "days_remaining": int(delta_days),
                "grace_days_left": int(grace_ends_days),
                "message": f"订阅已过期，宽限期剩余 {int(grace_ends_days)} 天"}

    # 宽限期也过 → 只读
    return {**base, "status": STATUS_READONLY,
            "days_remaining": int(delta_days),
            "grace_days_left": 0, "readonly": True,
            "message": "订阅已过期，进入只读模式"}


def is_activated() -> tuple[bool, str]:
    """向后兼容的简单接口：返回 (是否可正常使用, 描述)。

    宽限期算"可用"，只读算"不可用（需重新激活）"。
    """
    s = get_license_status()
    if s["status"] in (STATUS_ACTIVE, STATUS_TRIAL, STATUS_GRACE):
        return True, s["message"]
    return False, s["message"]


def is_readonly_mode() -> bool:
    """便捷函数：当前是否应进入全局只读模式"""
    return get_license_status()["readonly"]


def get_activated_license_code() -> str | None:
    """返回当前激活码（如果已激活），否则 None"""
    envelope = load_local_token()
    if not envelope:
        return None
    payload = _parse_and_verify_token(envelope)
    return payload.get("license_code") if payload else None


# ─── 5. HTTP 调用 Workers ─────────────────────────────────────────
def _post(path: str, body: dict, headers: dict | None = None,
          timeout: int = 15) -> tuple[int, dict]:
    h = {"Content-Type": "application/json", "User-Agent": USER_AGENT}
    if headers:
        h.update(headers)
    req = urllib.request.Request(
        WORKERS_URL + path,
        data=json.dumps(body).encode("utf-8"),
        method="POST",
        headers=h,
    )
    try:
        with urllib.request.urlopen(req, timeout=timeout) as resp:
            return resp.status, json.loads(resp.read().decode("utf-8"))
    except urllib.error.HTTPError as e:
        try:
            return e.code, json.loads(e.read().decode("utf-8"))
        except Exception:
            return e.code, {"error": "服务器返回了非 JSON 错误"}
    except urllib.error.URLError as e:
        raise LicenseError(f"网络错误：{e.reason}")
    except Exception as e:
        raise LicenseError(f"请求失败：{e}")


# ─── 6. 激活 ──────────────────────────────────────────────────────
_CODE_RE = re.compile(r"^WL-[A-Z2-9]{4}-[A-Z2-9]{4}-[A-Z2-9]{4}$")
_ALPHABET = "ABCDEFGHJKLMNPQRSTUVWXYZ23456789"


def normalize_license_code(raw: str) -> str:
    """把用户输入规范化：去空格、小写转大写"""
    return raw.replace(" ", "").replace("　", "").upper()


def is_valid_format(code: str) -> bool:
    """格式 + 校验位本地预检（避免无效请求打到服务器）"""
    if not _CODE_RE.match(code):
        return False
    body = code[3:].replace("-", "")  # 12 chars
    s = sum(_ALPHABET.index(c) for c in body[:11])
    return _ALPHABET[s % 32] == body[11]


def activate(license_code: str) -> tuple[bool, str]:
    """调用 Workers /activate 完成激活。

    返回 (成功?, 消息)。成功时 token 已自动保存到本地。
    """
    code = normalize_license_code(license_code)
    if not is_valid_format(code):
        return False, "激活码格式错误，请检查是否输入完整（含校验位）"

    try:
        machine_id = get_machine_id()
    except LicenseError as e:
        return False, str(e)

    status, body = _post("/activate", {
        "code": code,
        "machine_id": machine_id,
    })

    if status == 200 and "token" in body:
        save_local_token(body["token"])
        return True, "激活成功"

    err_msg = body.get("error", "未知错误") if isinstance(body, dict) else str(body)
    if status == 404:
        return False, f"激活码不存在：{err_msg}"
    if status == 403:
        if "revoked" in err_msg.lower():
            return False, "此激活码已被吊销，请联系客服"
        if "unbind" in err_msg.lower() or "limit" in err_msg.lower():
            return False, "本年度解绑次数已用尽（每年限 2 次），请联系客服"
        return False, f"激活被拒绝：{err_msg}"
    if status == 400:
        return False, f"请求错误：{err_msg}"
    return False, f"服务器返回 HTTP {status}：{err_msg}"


# ─── 6.5 免费试用 ─────────────────────────────────────────────────
def request_trial() -> tuple[bool, str]:
    """向服务器申请 7 天免费试用。

    返回 (成功?, 消息)。成功时 token 已自动保存到本地。
    每台电脑仅可申请一次；再次申请会被服务器拒绝。
    """
    try:
        machine_id = get_machine_id()
    except LicenseError as e:
        return False, str(e)

    status, body = _post("/trial", {"machine_id": machine_id})

    if status == 200 and "token" in body:
        save_local_token(body["token"])
        return True, "试用已激活，有效期 7 天"

    err_msg = body.get("error", "未知错误") if isinstance(body, dict) else str(body)
    if status == 403 and "already used" in err_msg.lower():
        return False, "此电脑已使用过试用（每台电脑仅可试用一次），请购买正式激活码"
    return False, f"申请试用失败：{err_msg}"


# ─── 7. 数据包更新 ────────────────────────────────────────────────
def check_data_update() -> dict | None:
    """检查并下载年度数据包更新。已是最新返回 None。"""
    envelope = load_local_token()
    if not envelope:
        raise LicenseError("尚未激活")
    payload = _parse_and_verify_token(envelope)
    if not payload:
        raise LicenseError("本地 token 无效")

    status, body = _post("/update-data", {
        "code": payload["license_code"],
        "machine_id": payload["machine_id"],
        "current_pack_version": CURRENT_DATA_PACK_VERSION,
    })

    if status != 200:
        err = body.get("error", str(status)) if isinstance(body, dict) else str(body)
        raise LicenseError(f"更新检查失败：{err}")

    if body.get("up_to_date"):
        return None  # 已是最新

    return body.get("data") or {}


# ─── 调试入口：python license.py ──────────────────────────────────
if __name__ == "__main__":
    print("=== WiseLedger License 调试 ===")
    print(f"系统：{platform.system()}")
    print(f"开发模式：{bool(os.environ.get('WL_DEV'))}")
    print(f"机器码：{get_machine_id()}")
    print(f"Token 路径：{_get_token_path()}")
    activated, msg = is_activated()
    print(f"激活状态：{msg}")
    if len(sys.argv) > 1:
        cmd = sys.argv[1]
        if cmd == "activate" and len(sys.argv) > 2:
            ok, m = activate(sys.argv[2])
            print(f"\n激活结果：{'✓' if ok else '✗'} {m}")
        elif cmd == "logout":
            remove_local_token()
            print("已删除本地激活信息")
        elif cmd == "update":
            try:
                data = check_data_update()
                if data is None:
                    print("已是最新版本")
                else:
                    print(f"获得新数据：{json.dumps(data, indent=2, ensure_ascii=False)}")
            except LicenseError as e:
                print(f"错误：{e}")
