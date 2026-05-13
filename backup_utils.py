"""backup_utils.py — 加密备份与恢复工具

备份文件格式（.zyac）：
    [4B  魔数      b'ZYAC'               ]
    [16B salt      随机，用于 PBKDF2 派生 ]
    [12B nonce     随机，AES-GCM IV      ]
    [nB  密文      AES-256-GCM 加密内容  ]
    [16B tag       GCM 认证标签（含在密文末尾，cryptography 库自动处理）]

密钥派生：PBKDF2-HMAC-SHA256，600,000 次迭代（OWASP 2023 推荐）
加密算法：AES-256-GCM（同时提供加密和完整性校验）
"""
import os
from cryptography.hazmat.primitives.kdf.pbkdf2 import PBKDF2HMAC
from cryptography.hazmat.primitives import hashes
from cryptography.hazmat.primitives.ciphers.aead import AESGCM

MAGIC = b"ZYAC"
PBKDF2_ITERATIONS = 600_000


def _derive_key(password: str, salt: bytes) -> bytes:
    kdf = PBKDF2HMAC(
        algorithm=hashes.SHA256(),
        length=32,
        salt=salt,
        iterations=PBKDF2_ITERATIONS,
    )
    return kdf.derive(password.encode("utf-8"))


def encrypt_backup(db_path: str, dest_path: str, password: str) -> None:
    """读取 db_path 的数据库文件，加密后写入 dest_path。"""
    with open(db_path, "rb") as f:
        data = f.read()
    salt  = os.urandom(16)
    nonce = os.urandom(12)
    key   = _derive_key(password, salt)
    ciphertext = AESGCM(key).encrypt(nonce, data, None)
    with open(dest_path, "wb") as f:
        f.write(MAGIC)
        f.write(salt)
        f.write(nonce)
        f.write(ciphertext)


def decrypt_backup(backup_path: str, dest_path: str, password: str) -> None:
    """读取 backup_path 的备份文件，解密后写入 dest_path。

    密码错误或文件损坏时抛出 ValueError。
    """
    with open(backup_path, "rb") as f:
        raw = f.read()
    if len(raw) < 32 or raw[:4] != MAGIC:
        raise ValueError("不是有效的智一会计备份文件（文件头不匹配）")
    salt      = raw[4:20]
    nonce     = raw[20:32]
    ciphertext = raw[32:]
    key = _derive_key(password, salt)
    try:
        data = AESGCM(key).decrypt(nonce, ciphertext, None)
    except Exception:
        raise ValueError("密码错误或备份文件已损坏，无法解密")
    with open(dest_path, "wb") as f:
        f.write(data)
