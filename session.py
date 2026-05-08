"""session.py — 全局登录会话与权限管理"""

# ── 角色权限定义 ──────────────────────────────────────────────
# 每个权限字符串对应一个具体操作
# superadmin 拥有所有权限，用 'all' 标记，检查时特殊处理

ROLE_PERMISSIONS: dict[str, set[str]] = {
    "superadmin": {"all"},   # 检查时直接返回 True
    "admin": {
        "client.view",        # 查看客户列表
        "client.manage",      # 新建/编辑/删除客户
        "account.view",       # 查看科目
        "account.manage",     # 管理科目
        "opening.view",       # 查看期初
        "opening.manage",     # 编辑期初
        "voucher.view",       # 查看凭证
        "voucher.create",     # 新增凭证
        "voucher.edit",       # 编辑凭证
        "voucher.delete",     # 删除凭证（仅待审核）
        "voucher.approve",    # 审核/撤销审核凭证
        "settle.manage",      # 期末结账/反结账
        "report.view",        # 查看报表
        "report.export",      # 导出报表
        "audit.view",         # 查看审计日志
        # 注意：admin 没有 system.manage，用户管理只有 superadmin
    },
    "accountant": {
        "client.view",
        "account.view",
        "opening.view",
        "voucher.view",
        "voucher.create",
        "voucher.edit",
        "report.view",
        "report.export",
    },
    "readonly": {
        "client.view",
        "account.view",
        "opening.view",
        "voucher.view",
        "report.view",
    },
}

# 角色显示名
ROLE_LABELS: dict[str, str] = {
    "superadmin": "超级管理员",
    "admin":      "管理员",
    "accountant": "会计",
    "readonly":   "只读",
}


class AppSession:
    """全局单例：保存当前登录用户信息，提供权限检查方法。"""

    _user: dict | None = None   # {"id":1, "username":"admin",
                                #  "display_name":"超级管理员", "role":"superadmin"}

    # ── 登录 / 登出 ──────────────────────────────────────────
    @classmethod
    def login(cls, user: dict) -> None:
        cls._user = user

    @classmethod
    def logout(cls) -> None:
        cls._user = None

    # ── 当前用户信息 ─────────────────────────────────────────
    @classmethod
    def get(cls) -> dict | None:
        return cls._user

    @classmethod
    def is_logged_in(cls) -> bool:
        return cls._user is not None

    @classmethod
    def display_name(cls) -> str:
        return cls._user["display_name"] if cls._user else ""

    @classmethod
    def role(cls) -> str:
        return cls._user["role"] if cls._user else ""

    @classmethod
    def role_label(cls) -> str:
        return ROLE_LABELS.get(cls.role(), "")

    # ── 权限检查 ─────────────────────────────────────────────
    @classmethod
    def has_perm(cls, perm: str) -> bool:
        """检查当前用户是否拥有指定权限。未登录返回 False。"""
        if not cls._user:
            return False
        role = cls._user.get("role", "")
        perms = ROLE_PERMISSIONS.get(role, set())
        if "all" in perms:      # superadmin
            return True
        return perm in perms

    @classmethod
    def can_access_client(cls, client_id: int) -> bool:
        """
        检查当前用户是否有权访问指定客户账套。
        superadmin / admin 可访问所有客户；
        accountant / readonly 只能访问 user_client_access 授权的客户。
        """
        if not cls._user:
            return False
        role = cls._user.get("role", "")
        if role in ("superadmin", "admin"):
            return True
        # 查数据库授权表
        try:
            from db import get_db
            conn = get_db()
            c = conn.cursor()
            c.execute("""SELECT id FROM user_client_access
                         WHERE user_id=? AND client_id=?""",
                      (cls._user["id"], client_id))
            result = c.fetchone()
            conn.close()
            return result is not None
        except Exception:
            return False
