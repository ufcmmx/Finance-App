"""kr_utils.py — 系统凭据管理器（keyring）工具函数

放在根目录，供 pages/system.py 和 auto_backup.py 共同导入，避免循环依赖。
Windows 底层使用 Credential Manager，macOS 使用 Keychain，均与系统账户绑定。
"""

_KR_SERVICE = "智一会计"
_KR_ACCOUNT = "backup_password"


def kr_get() -> str | None:
    """从系统凭据管理器读取备份密码，失败或未设置返回 None。"""
    try:
        import keyring
        return keyring.get_password(_KR_SERVICE, _KR_ACCOUNT)
    except Exception:
        return None


def kr_set(pw: str) -> bool:
    """将备份密码存入系统凭据管理器，成功返回 True。"""
    try:
        import keyring
        keyring.set_password(_KR_SERVICE, _KR_ACCOUNT, pw)
        return True
    except Exception:
        return False


def kr_available() -> bool:
    """检测 keyring 后端是否可用（PyInstaller 打包后偶尔缺失）。"""
    try:
        import keyring
        keyring.get_password(_KR_SERVICE, "__probe__")
        return True
    except Exception:
        return False
