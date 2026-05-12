"""login_dialog.py — 登录窗口"""
from PySide6.QtWidgets import (
    QDialog, QVBoxLayout, QHBoxLayout, QLabel,
    QLineEdit, QPushButton, QMessageBox, QFrame
)
from PySide6.QtCore import Qt, QTimer
from PySide6.QtGui import QFont, QKeyEvent

from db import get_db
from session import AppSession
from pw_utils import hash_pw, verify_pw


class LoginDialog(QDialog):
    """登录窗口 — 验证用户名密码，写入 AppSession"""

    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("智一会计 · 登录")
        self.setFixedSize(400, 340)
        self.setWindowFlags(Qt.Dialog | Qt.WindowTitleHint)
        self._build()

    def _build(self):
        root = QVBoxLayout(self)
        root.setContentsMargins(0, 0, 0, 0)
        root.setSpacing(0)

        # ── 顶部深色标题栏 ──
        header = QFrame()
        header.setStyleSheet("background:#1c2340;")
        header.setFixedHeight(100)
        hl = QVBoxLayout(header)
        hl.setContentsMargins(32, 20, 32, 16)
        title = QLabel("智一会计")
        title.setStyleSheet("color:#ff8c00;font-size:22px;font-weight:bold;")
        sub = QLabel("本地专业版  —  请登录")
        sub.setStyleSheet("color:#8b93ae;font-size:12px;")
        hl.addWidget(title)
        hl.addWidget(sub)
        root.addWidget(header)

        # ── 表单区 ──
        body = QFrame()
        body.setStyleSheet("background:#f0f2f5;")
        bl = QVBoxLayout(body)
        bl.setContentsMargins(40, 28, 40, 28)
        bl.setSpacing(0)

        def _add_field(label_text: str, widget: QLineEdit):
            layout = QVBoxLayout()
            layout.setContentsMargins(0, 0, 0, 0)
            layout.setSpacing(4)
            layout.addWidget(self._field_label(label_text))
            layout.addWidget(widget)
            bl.addLayout(layout)
            bl.addSpacing(18)

        # 用户名
        self.f_user = QLineEdit()
        self.f_user.setPlaceholderText("请输入用户名")
        self.f_user.setFixedHeight(30)
        self.f_user.setStyleSheet(self._input_style())
        _add_field("用户名", self.f_user)

        # 密码
        self.f_pass = QLineEdit()
        self.f_pass.setPlaceholderText("请输入密码")
        self.f_pass.setEchoMode(QLineEdit.Password)
        self.f_pass.setFixedHeight(30)
        self.f_pass.setStyleSheet(self._input_style())
        self.f_pass.returnPressed.connect(self._do_login)
        _add_field("密码", self.f_pass)

        # 错误提示（平时隐藏）
        self.err_lbl = QLabel("")
        self.err_lbl.setStyleSheet(
            "color:#ff4d4f;font-size:12px;padding:4px 0;")
        self.err_lbl.setVisible(False)
        bl.addWidget(self.err_lbl)

        bl.addSpacing(6)

        # 登录按钮
        self.btn_login = QPushButton("登  录")
        self.btn_login.setFixedHeight(40)
        self.btn_login.setStyleSheet("""
            QPushButton {
                background:#3d6fdb; color:#fff; border:none;
                border-radius:6px; font-size:14px; font-weight:bold;
            }
            QPushButton:hover { background:#2d5dc8; }
            QPushButton:disabled { background:#b0bec5; }
        """)
        self.btn_login.clicked.connect(self._do_login)
        bl.addWidget(self.btn_login)

        # 默认账号提示（首次使用）
        hint = QLabel("默认账号：admin  /  默认密码：admin123")
        hint.setStyleSheet("color:#aaa;font-size:11px;")
        hint.setAlignment(Qt.AlignCenter)
        bl.addWidget(hint)

        root.addWidget(body)

        # 焦点默认在用户名
        QTimer.singleShot(100, self.f_user.setFocus)

    def _field_label(self, text: str) -> QLabel:
        lbl = QLabel(text)
        lbl.setStyleSheet("color:#555;font-size:13px;font-weight:bold;")
        return lbl

    def _input_style(self) -> str:
        return """
            QLineEdit {
                background:#fff; border:1px solid #bfc7d3;
                border-radius:6px; padding:2px 10px; font-size:13px;
            }
            QLineEdit:focus { border:1.5px solid #3d6fdb; }
        """

    def _do_login(self):
        username = self.f_user.text().strip()
        password = self.f_pass.text()

        if not username or not password:
            self._show_error("用户名和密码不能为空")
            return

        self.btn_login.setEnabled(False)
        self.btn_login.setText("验证中…")

        try:
            conn = get_db()
            c = conn.cursor()
            c.execute("""SELECT id, username, display_name, role, is_active, password_hash
                         FROM users WHERE username=?""", (username,))
            user = c.fetchone()

            if not user:
                self._show_error("用户名不存在")
                return

            if not user["is_active"]:
                self._show_error("该账号已被停用，请联系管理员")
                return

            valid, needs_rehash = verify_pw(password, user["password_hash"])
            if not valid:
                self._show_error("密码错误")
                return

            # 旧版 SHA-256 哈希透明迁移为 Argon2
            if needs_rehash:
                c.execute("UPDATE users SET password_hash=? WHERE id=?",
                          (hash_pw(password), user["id"]))

            # 更新最后登录时间
            c.execute("UPDATE users SET last_login=CURRENT_TIMESTAMP WHERE id=?",
                      (user["id"],))
            conn.commit()
            conn.close()

            # 写入全局会话
            AppSession.login({
                "id":           user["id"],
                "username":     user["username"],
                "display_name": user["display_name"],
                "role":         user["role"],
            })

            self.accept()

        except Exception as e:
            self._show_error(f"登录失败：{e}")
        finally:
            self.btn_login.setEnabled(True)
            self.btn_login.setText("登  录")

    def _show_error(self, msg: str):
        self.err_lbl.setText(f"⚠ {msg}")
        self.err_lbl.setVisible(True)
        self.f_pass.clear()
        self.f_pass.setFocus()