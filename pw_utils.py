"""pw_utils.py — 密码哈希工具

使用 Argon2id 算法替代原 SHA-256（无盐）方案。
兼容旧版 SHA-256 哈希，登录时自动透明迁移，用户无感知。

使用方式：
    from pw_utils import hash_pw, verify_pw

    # 新建/修改密码
    stored = hash_pw("my_password")

    # 验证密码（自动处理新旧格式）
    valid, needs_rehash = verify_pw("my_password", stored)
    if valid and needs_rehash:
        # 旧格式，登录后更新数据库
        conn.execute("UPDATE users SET password_hash=? WHERE id=?",
                     (hash_pw("my_password"), user_id))
"""
import hashlib
from argon2 import PasswordHasher
from argon2.exceptions import VerifyMismatchError, VerificationError, InvalidHashError

# Argon2id 参数（OWASP 推荐最低配置）
_ph = PasswordHasher(
    time_cost=2,        # 迭代次数
    memory_cost=65536,  # 64 MB 内存
    parallelism=2,      # 并行度
    hash_len=32,
    salt_len=16,
)


def hash_pw(password: str) -> str:
    """生成 Argon2id 哈希，用于新建用户或修改密码。"""
    return _ph.hash(password)


def _is_legacy_sha256(h: str) -> bool:
    """判断是否为旧版 SHA-256 哈希（64位十六进制字符串）。"""
    return len(h) == 64 and all(c in "0123456789abcdef" for c in h.lower())


def _legacy_hash(password: str) -> str:
    return hashlib.sha256(password.encode()).hexdigest()


def verify_pw(password: str, stored_hash: str) -> tuple[bool, bool]:
    """验证密码，返回 (is_valid, needs_rehash)。

    needs_rehash=True 表示当前存储的是旧版 SHA-256 格式，
    调用方应在验证通过后将密码升级为 Argon2 存回数据库。
    """
    if _is_legacy_sha256(stored_hash):
        valid = _legacy_hash(password) == stored_hash
        return valid, valid   # 验证通过时 needs_rehash=True，触发迁移
    try:
        _ph.verify(stored_hash, password)
        needs_rehash = _ph.check_needs_rehash(stored_hash)
        return True, needs_rehash
    except (VerifyMismatchError, VerificationError, InvalidHashError):
        return False, False
