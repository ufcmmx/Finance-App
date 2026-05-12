"""pages/system.py — 系统管理页（用户管理、客户授权）"""
from pw_utils import hash_pw, verify_pw
from PySide6.QtWidgets import *
from PySide6.QtCore import Qt
from PySide6.QtGui import QColor, QFont

from db import get_db, log_action
from session import AppSession, ROLE_LABELS
from utils import lbl, card, sep





# ── 新增/编辑用户对话框 ────────────────────────────────────────────────────
class UserDialog(QDialog):
    def __init__(self, parent=None, user: dict = None):
        super().__init__(parent)
        self.user = user  # None = 新增，dict = 编辑
        self.setWindowTitle("编辑用户" if user else "新增用户")
        self.setFixedWidth(420)
        self._build()
        if user:
            self._load()

    def _build(self):
        L = QVBoxLayout(self)
        L.setContentsMargins(28, 24, 28, 24)
        L.setSpacing(14)
        L.addWidget(lbl("用户信息", bold=True, size=14))

        F = QFormLayout()
        F.setSpacing(10)
        F.setLabelAlignment(Qt.AlignRight)

        self.f_username = QLineEdit()
        self.f_username.setPlaceholderText("登录用户名（不可重复）")
        self.f_display  = QLineEdit()
        self.f_display.setPlaceholderText("显示名称，如：张三")

        self.f_role = QComboBox()
        for role_key, role_label in ROLE_LABELS.items():
            self.f_role.addItem(role_label, role_key)

        self.f_pw1 = QLineEdit()
        self.f_pw1.setEchoMode(QLineEdit.Password)
        self.f_pw1.setPlaceholderText("留空则不修改密码" if self.user else "请输入密码")
        self.f_pw2 = QLineEdit()
        self.f_pw2.setEchoMode(QLineEdit.Password)
        self.f_pw2.setPlaceholderText("再次输入密码")

        self.f_active = QCheckBox("账号启用")
        self.f_active.setChecked(True)

        F.addRow("用户名 *", self.f_username)
        F.addRow("显示名称 *", self.f_display)
        F.addRow("角色 *", self.f_role)
        F.addRow("密码", self.f_pw1)
        F.addRow("确认密码", self.f_pw2)
        F.addRow("状态", self.f_active)
        L.addLayout(F)

        # 禁止修改自己的角色和状态（防止把自己降权/停用）
        if self.user and self.user["id"] == AppSession.get()["id"]:
            self.f_role.setEnabled(False)
            self.f_active.setEnabled(False)
            note = QLabel("⚠ 不能修改自己的角色和启用状态")
            note.setStyleSheet("color:#fa8c16;font-size:12px;")
            L.addWidget(note)

        row = QHBoxLayout()
        row.addStretch()
        bc = QPushButton("取消"); bc.setObjectName("btn_gray")
        bs = QPushButton("保 存"); bs.setObjectName("btn_primary")
        bc.clicked.connect(self.reject)
        bs.clicked.connect(self._save)
        row.addWidget(bc); row.addWidget(bs)
        L.addLayout(row)

    def _load(self):
        u = self.user
        self.f_username.setText(u["username"])
        self.f_username.setReadOnly(True)  # 用户名建立后不可改
        self.f_username.setStyleSheet("background:#f5f5f5;color:#999;")
        self.f_display.setText(u["display_name"])
        for i in range(self.f_role.count()):
            if self.f_role.itemData(i) == u["role"]:
                self.f_role.setCurrentIndex(i)
                break
        self.f_active.setChecked(bool(u["is_active"]))

    def _save(self):
        username = self.f_username.text().strip()
        display  = self.f_display.text().strip()
        role     = self.f_role.currentData()
        pw1      = self.f_pw1.text()
        pw2      = self.f_pw2.text()
        active   = 1 if self.f_active.isChecked() else 0

        if not username:
            QMessageBox.warning(self, "提示", "用户名不能为空"); return
        if not display:
            QMessageBox.warning(self, "提示", "显示名称不能为空"); return
        if not self.user and not pw1:
            QMessageBox.warning(self, "提示", "新用户必须设置密码"); return
        if pw1 and pw1 != pw2:
            QMessageBox.warning(self, "提示", "两次输入的密码不一致"); return
        if pw1 and len(pw1) < 6:
            QMessageBox.warning(self, "提示", "密码长度不能少于6位"); return

        conn = get_db(); c = conn.cursor()
        try:
            if self.user:
                if pw1:
                    c.execute("""UPDATE users SET display_name=?, role=?,
                                 is_active=?, password_hash=? WHERE id=?""",
                              (display, role, active, hash_pw(pw1), self.user["id"]))
                else:
                    c.execute("""UPDATE users SET display_name=?, role=?,
                                 is_active=? WHERE id=?""",
                              (display, role, active, self.user["id"]))
                log_action(conn, None, "编辑用户", "user", self.user["id"],
                           f"用户:{username} 角色:{role}")
            else:
                c.execute("""INSERT INTO users(username, password_hash, display_name, role, is_active)
                             VALUES(?,?,?,?,?)""",
                          (username, hash_pw(pw1), display, role, active))
                log_action(conn, None, "新增用户", "user", c.lastrowid,
                           f"用户:{username} 角色:{role}")
            conn.commit()
            self.accept()
        except Exception as e:
            conn.rollback()
            QMessageBox.warning(self, "保存失败", str(e))
        finally:
            conn.close()


# ── 客户授权对话框 ─────────────────────────────────────────────────────────
class ClientAccessDialog(QDialog):
    """为某个用户分配可访问的客户账套"""
    def __init__(self, parent=None, user: dict = None):
        super().__init__(parent)
        self.user = user
        self.setWindowTitle(f"客户授权 — {user['display_name']}")
        self.setMinimumSize(460, 500)
        self._build()
        self._load()

    def _build(self):
        L = QVBoxLayout(self)
        L.setContentsMargins(24, 20, 24, 20)
        L.setSpacing(12)

        title = lbl(f"为【{self.user['display_name']}】分配可访问的客户", bold=True, size=14)
        L.addWidget(title)

        note = QLabel("  superadmin / admin 可访问全部客户，无需单独授权。\n"
                      "  会计 / 只读 角色只能访问已勾选的客户。")
        note.setStyleSheet("background:#f6f8ff;color:#555;border-radius:5px;"
                           "padding:8px 12px;font-size:12px;")
        note.setWordWrap(True)
        L.addWidget(note)

        # 全选/取消
        chk_row = QHBoxLayout()
        b_all  = QPushButton("全选");  b_all.setObjectName("btn_outline")
        b_none = QPushButton("取消全选"); b_none.setObjectName("btn_gray")
        b_all.clicked.connect(lambda: self._check_all(True))
        b_none.clicked.connect(lambda: self._check_all(False))
        chk_row.addWidget(b_all); chk_row.addWidget(b_none); chk_row.addStretch()
        L.addLayout(chk_row)

        # 客户列表
        f = card()
        vl = QVBoxLayout(f); vl.setContentsMargins(0, 0, 0, 0)
        self.client_list = QListWidget()
        self.client_list.setStyleSheet(
            "QListWidget{border:none;}"
            "QListWidget::item{padding:8px 14px;border-bottom:1px solid #f0f2f5;}"
            "QListWidget::item:selected{background:#e6f0ff;}")
        vl.addWidget(self.client_list)
        L.addWidget(f)

        row = QHBoxLayout(); row.addStretch()
        bc = QPushButton("取消"); bc.setObjectName("btn_gray")
        bs = QPushButton("保 存授权"); bs.setObjectName("btn_primary")
        bc.clicked.connect(self.reject)
        bs.clicked.connect(self._save)
        row.addWidget(bc); row.addWidget(bs)
        L.addLayout(row)

    def _load(self):
        conn = get_db(); c = conn.cursor()
        c.execute("SELECT id, name FROM clients ORDER BY name")
        all_clients = c.fetchall()
        c.execute("SELECT client_id FROM user_client_access WHERE user_id=?",
                  (self.user["id"],))
        granted = {r[0] for r in c.fetchall()}
        conn.close()

        self.client_list.clear()
        for cl in all_clients:
            item = QListWidgetItem(cl["name"])
            item.setData(Qt.UserRole, cl["id"])
            item.setFlags(item.flags() | Qt.ItemIsUserCheckable)
            item.setCheckState(Qt.Checked if cl["id"] in granted else Qt.Unchecked)
            self.client_list.addItem(item)

    def _check_all(self, checked: bool):
        state = Qt.Checked if checked else Qt.Unchecked
        for i in range(self.client_list.count()):
            self.client_list.item(i).setCheckState(state)

    def _save(self):
        checked_ids = []
        for i in range(self.client_list.count()):
            item = self.client_list.item(i)
            if item.checkState() == Qt.Checked:
                checked_ids.append(item.data(Qt.UserRole))

        conn = get_db(); c = conn.cursor()
        try:
            c.execute("DELETE FROM user_client_access WHERE user_id=?",
                      (self.user["id"],))
            for cid in checked_ids:
                c.execute("INSERT OR IGNORE INTO user_client_access(user_id, client_id)"
                          " VALUES(?,?)", (self.user["id"], cid))
            log_action(conn, None, "更新客户授权", "user", self.user["id"],
                       f"授权客户数:{len(checked_ids)}")
            conn.commit()
            QMessageBox.information(self, "成功",
                                    f"已为【{self.user['display_name']}】授权 {len(checked_ids)} 个客户")
            self.accept()
        except Exception as e:
            conn.rollback()
            QMessageBox.warning(self, "保存失败", str(e))
        finally:
            conn.close()


# ── 修改密码对话框 ─────────────────────────────────────────────────────────
class ChangePasswordDialog(QDialog):
    """用户修改自己的密码"""
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("修改密码")
        self.setFixedWidth(380)
        self._build()

    def _build(self):
        L = QVBoxLayout(self)
        L.setContentsMargins(28, 24, 28, 24)
        L.setSpacing(14)
        L.addWidget(lbl("修改登录密码", bold=True, size=14))

        F = QFormLayout(); F.setSpacing(10); F.setLabelAlignment(Qt.AlignRight)
        self.f_old = QLineEdit(); self.f_old.setEchoMode(QLineEdit.Password)
        self.f_old.setPlaceholderText("当前密码")
        self.f_new1 = QLineEdit(); self.f_new1.setEchoMode(QLineEdit.Password)
        self.f_new1.setPlaceholderText("新密码（至少6位）")
        self.f_new2 = QLineEdit(); self.f_new2.setEchoMode(QLineEdit.Password)
        self.f_new2.setPlaceholderText("再次输入新密码")
        F.addRow("当前密码", self.f_old)
        F.addRow("新密码",   self.f_new1)
        F.addRow("确认新密码", self.f_new2)
        L.addLayout(F)

        row = QHBoxLayout(); row.addStretch()
        bc = QPushButton("取消"); bc.setObjectName("btn_gray")
        bs = QPushButton("确认修改"); bs.setObjectName("btn_primary")
        bc.clicked.connect(self.reject)
        bs.clicked.connect(self._save)
        row.addWidget(bc); row.addWidget(bs)
        L.addLayout(row)

    def _save(self):
        old  = self.f_old.text()
        new1 = self.f_new1.text()
        new2 = self.f_new2.text()
        if not old or not new1:
            QMessageBox.warning(self, "提示", "请填写所有字段"); return
        if new1 != new2:
            QMessageBox.warning(self, "提示", "两次输入的新密码不一致"); return
        if len(new1) < 6:
            QMessageBox.warning(self, "提示", "新密码不能少于6位"); return

        uid = AppSession.get()["id"]
        conn = get_db(); c = conn.cursor()
        c.execute("SELECT password_hash FROM users WHERE id=?", (uid,))
        row = c.fetchone()
        valid, _ = verify_pw(old, row["password_hash"]) if row else (False, False)
        if not row or not valid:
            conn.close()
            QMessageBox.warning(self, "错误", "当前密码不正确"); return

        c.execute("UPDATE users SET password_hash=? WHERE id=?",
                  (hash_pw(new1), uid))
        log_action(conn, None, "修改密码", "user", uid, "用户自行修改密码")
        conn.commit(); conn.close()
        QMessageBox.information(self, "成功", "密码已修改，下次登录生效")
        self.accept()


# ── 主页面 ─────────────────────────────────────────────────────────────────
class SystemPage(QWidget):
    def __init__(self, parent=None):
        super().__init__(parent)
        self._build()

    def _build(self):
        L = QVBoxLayout(self)
        L.setContentsMargins(0, 0, 0, 0)
        L.setSpacing(0)

        # Tab 栏
        tab_bar = QWidget()
        tab_bar.setStyleSheet("background:#fff;border-bottom:1px solid #e8ecf2;")
        tl = QHBoxLayout(tab_bar)
        tl.setContentsMargins(16, 0, 16, 0)
        tl.setSpacing(0)
        self._tabs = []
        for name in ["用户管理", "修改密码"]:
            b = QPushButton(name)
            b.setStyleSheet("""QPushButton{background:transparent;color:#888;border:none;
                padding:12px 16px;border-bottom:2px solid transparent;}
                QPushButton:hover{color:#3d6fdb;}
                QPushButton[active=true]{color:#3d6fdb;border-bottom:2px solid #3d6fdb;}""")
            b.clicked.connect(lambda _, n=name: self._switch(n))
            tl.addWidget(b); self._tabs.append(b)
        tl.addStretch()
        L.addWidget(tab_bar)

        self.stack = QStackedWidget()
        self._build_user_mgmt()
        self._build_change_pw()
        L.addWidget(self.stack)
        self._switch("用户管理")

    def _switch(self, name):
        mapping = {"用户管理": 0, "修改密码": 1}
        # 非 superadmin 隐藏用户管理 tab
        if name == "用户管理" and not AppSession.has_perm("system.manage"):
            name = "修改密码"
        self.stack.setCurrentIndex(mapping[name])
        for b in self._tabs:
            is_active = b.text() == name
            b.setProperty("active", "true" if is_active else "false")
            b.style().unpolish(b); b.style().polish(b)
            # 隐藏没有权限的 tab
            if b.text() == "用户管理":
                b.setVisible(AppSession.has_perm("system.manage"))
        if name == "用户管理":
            self._load_users()

    def refresh_after_login(self):
        """登录后调用，重新刷新 SystemPage 的状态"""
        self._switch("用户管理")

    # ── Tab 1：用户管理 ──
    def _build_user_mgmt(self):
        w = QWidget()
        L = QVBoxLayout(w)
        L.setContentsMargins(24, 16, 24, 16)
        L.setSpacing(12)

        hdr = QHBoxLayout()
        hdr.addWidget(lbl("用户管理", bold=True, size=15))
        hdr.addStretch()
        b_add = QPushButton("＋ 新增用户"); b_add.setObjectName("btn_primary")
        b_add.clicked.connect(self._add_user)
        hdr.addWidget(b_add)
        L.addLayout(hdr)

        note = QLabel("  用户管理仅超级管理员可见。"
                      "会计 / 只读 角色需要单独授权才能访问对应客户账套。")
        note.setStyleSheet("background:#f6f8ff;color:#555;border-radius:5px;"
                           "padding:8px 12px;font-size:12px;")
        note.setWordWrap(True)
        L.addWidget(note)

        f = card()
        vl = QVBoxLayout(f); vl.setContentsMargins(0, 0, 0, 0)
        self.user_tbl = QTableWidget()
        self.user_tbl.setEditTriggers(QTableWidget.NoEditTriggers)
        self.user_tbl.setSelectionBehavior(QTableWidget.SelectRows)
        self.user_tbl.setShowGrid(False)
        self.user_tbl.verticalHeader().setVisible(False)
        self.user_tbl.setColumnCount(6)
        self.user_tbl.setHorizontalHeaderLabels(
            ["用户名", "显示名称", "角色", "状态", "最后登录", "操作"])
        hh = self.user_tbl.horizontalHeader()
        hh.setSectionResizeMode(QHeaderView.Interactive)
        hh.setSectionResizeMode(1, QHeaderView.Stretch)
        self.user_tbl.setColumnWidth(0, 110)
        self.user_tbl.setColumnWidth(2, 100)
        self.user_tbl.setColumnWidth(3, 70)
        self.user_tbl.setColumnWidth(4, 140)
        self.user_tbl.setColumnWidth(5, 260)
        vl.addWidget(self.user_tbl)
        L.addWidget(f)
        self.stack.addWidget(w)

    def _load_users(self):
        conn = get_db(); c = conn.cursor()
        c.execute("""SELECT id, username, display_name, role, is_active, last_login
                     FROM users ORDER BY id""")
        users = c.fetchall(); conn.close()

        self.user_tbl.setRowCount(len(users))
        cur_uid = AppSession.get()["id"]

        for i, u in enumerate(users):
            self.user_tbl.setRowHeight(i, 46)
            role_label = ROLE_LABELS.get(u["role"], u["role"])
            role_color = {"superadmin": "#3d6fdb", "admin": "#52c41a",
                          "accountant": "#fa8c16", "readonly": "#888"}.get(u["role"], "#555")
            active_text = "启用" if u["is_active"] else "停用"
            active_color = "#52c41a" if u["is_active"] else "#ff4d4f"
            last_login = (u["last_login"] or "从未登录")[:16]

            for j, (val, align, color) in enumerate([
                (u["username"],   Qt.AlignCenter,              None),
                (u["display_name"], Qt.AlignLeft|Qt.AlignVCenter, None),
                (role_label,      Qt.AlignCenter,              role_color),
                (active_text,     Qt.AlignCenter,              active_color),
                (last_login,      Qt.AlignCenter,              "#888"),
            ]):
                it = QTableWidgetItem(val)
                it.setTextAlignment(align)
                if color:
                    it.setForeground(QColor(color))
                # 当前登录用户高亮
                if u["id"] == cur_uid:
                    it.setBackground(QColor("#f0f7ff"))
                self.user_tbl.setItem(i, j, it)

            # 操作按钮
            bw = QWidget(); bl = QHBoxLayout(bw)
            bl.setContentsMargins(6, 3, 6, 3); bl.setSpacing(6)

            b_edit = QPushButton("编辑"); b_edit.setObjectName("btn_outline")
            b_edit.setFixedSize(60, 28)
            b_edit.clicked.connect(lambda _, ud=dict(u): self._edit_user(ud))

            b_auth = QPushButton("客户授权"); b_auth.setObjectName("btn_outline")
            b_auth.setFixedSize(76, 28)
            # superadmin/admin 不需要客户授权
            b_auth.setEnabled(u["role"] in ("accountant", "readonly"))
            b_auth.clicked.connect(lambda _, ud=dict(u): self._edit_access(ud))

            b_del = QPushButton("删除"); b_del.setObjectName("btn_red")
            b_del.setFixedSize(56, 28)
            # 不能删除自己
            b_del.setEnabled(u["id"] != cur_uid)
            b_del.clicked.connect(lambda _, uid=u["id"], uname=u["username"]:
                                  self._del_user(uid, uname))

            bl.addWidget(b_edit); bl.addWidget(b_auth)
            bl.addWidget(b_del); bl.addStretch()
            self.user_tbl.setCellWidget(i, 5, bw)

    def _add_user(self):
        d = UserDialog(self)
        if d.exec():
            self._load_users()

    def _edit_user(self, user: dict):
        d = UserDialog(self, user=user)
        if d.exec():
            self._load_users()

    def _edit_access(self, user: dict):
        d = ClientAccessDialog(self, user=user)
        d.exec()

    def _del_user(self, uid: int, username: str):
        if QMessageBox.question(
                self, "确认删除",
                f"确认删除用户【{username}】？\n该用户的所有客户授权也会同步删除。",
                QMessageBox.Yes | QMessageBox.No) != QMessageBox.Yes:
            return
        conn = get_db()
        conn.execute("DELETE FROM users WHERE id=?", (uid,))
        log_action(conn, None, "删除用户", "user", uid, f"用户:{username}")
        conn.commit(); conn.close()
        self._load_users()

    # ── Tab 2：修改密码 ──
    def _build_change_pw(self):
        w = QWidget()
        L = QVBoxLayout(w)
        L.setContentsMargins(24, 16, 24, 16)
        L.setSpacing(16)
        L.addWidget(lbl("账号设置", bold=True, size=15))

        # 当前用户信息卡片
        info_frame = QFrame()
        info_frame.setStyleSheet(
            "QFrame{background:#fff;border:1px solid #e4e8f0;border-radius:8px;}")
        il = QHBoxLayout(info_frame)
        il.setContentsMargins(20, 16, 20, 16)
        il.setSpacing(20)
        user = AppSession.get() or {}
        il.addWidget(lbl("当前用户：", color="#888"))
        il.addWidget(lbl(user.get("display_name", ""), bold=True))
        il.addWidget(lbl("角色：", color="#888"))
        role_label = ROLE_LABELS.get(user.get("role", ""), "")
        il.addWidget(lbl(role_label, color="#3d6fdb", bold=True))
        il.addStretch()
        L.addWidget(info_frame)

        b_pw = QPushButton("修改登录密码")
        b_pw.setObjectName("btn_outline")
        b_pw.setFixedWidth(140)
        b_pw.clicked.connect(lambda: ChangePasswordDialog(self).exec())
        L.addWidget(b_pw)
        L.addStretch()
        self.stack.addWidget(w)