"""pages/system.py — 系统管理页（用户管理、客户授权）"""
import os
from pw_utils import hash_pw, verify_pw
from PySide6.QtWidgets import *
from PySide6.QtCore import Qt
from PySide6.QtGui import QColor, QFont

from db import get_db, DB_PATH, log_action, get_setting, set_setting
from pw_utils import hash_pw, verify_pw
from backup_utils import encrypt_backup, decrypt_backup
from session import AppSession, ROLE_LABELS
from utils import lbl, card, sep

from datetime import datetime

from kr_utils import kr_get, kr_set, kr_available




# ── 新增/编辑用户对话框 ────────────────────────────────────────────────────
class UserDialog(QDialog):
    def __init__(self, parent=None, user: dict | None = None):
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
        F.setLabelAlignment(Qt.AlignmentFlag.AlignRight)

        self.f_username = QLineEdit()
        self.f_username.setPlaceholderText("登录用户名（不可重复）")
        self.f_display  = QLineEdit()
        self.f_display.setPlaceholderText("显示名称，如：张三")

        self.f_role = QComboBox()
        for role_key, role_label in ROLE_LABELS.items():
            self.f_role.addItem(role_label, role_key)

        self.f_pw1 = QLineEdit()
        self.f_pw1.setEchoMode(QLineEdit.EchoMode.Password)
        self.f_pw1.setPlaceholderText("留空则不修改密码" if self.user else "请输入密码")
        self.f_pw2 = QLineEdit()
        self.f_pw2.setEchoMode(QLineEdit.EchoMode.Password)
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
        _cur = AppSession.get()
        if self.user and _cur and self.user["id"] == _cur["id"]:
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
        if u is None:
            return
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
                log_action(conn, None, "新增用户", "user", str(c.lastrowid or 0),
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
    def __init__(self, parent=None, user: dict | None = None):
        super().__init__(parent)
        self.user = user
        _dn = user['display_name'] if user else ''
        self.setWindowTitle(f"客户授权 — {_dn}")
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
            item.setData(Qt.ItemDataRole.UserRole, cl["id"])
            item.setFlags(item.flags() | Qt.ItemFlag.ItemIsUserCheckable)
            item.setCheckState(Qt.CheckState.Checked if cl["id"] in granted else Qt.CheckState.Unchecked)
            self.client_list.addItem(item)

    def _check_all(self, checked: bool):
        state = Qt.CheckState.Checked if checked else Qt.CheckState.Unchecked
        for i in range(self.client_list.count()):
            self.client_list.item(i).setCheckState(state)

    def _save(self):
        checked_ids = []
        for i in range(self.client_list.count()):
            item = self.client_list.item(i)
            if item.checkState() == Qt.CheckState.Checked:
                checked_ids.append(item.data(Qt.ItemDataRole.UserRole))

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

        F = QFormLayout(); F.setSpacing(10); F.setLabelAlignment(Qt.AlignmentFlag.AlignRight)
        self.f_old = QLineEdit(); self.f_old.setEchoMode(QLineEdit.EchoMode.Password)
        self.f_old.setPlaceholderText("当前密码")
        self.f_new1 = QLineEdit(); self.f_new1.setEchoMode(QLineEdit.EchoMode.Password)
        self.f_new1.setPlaceholderText("新密码（至少6位）")
        self.f_new2 = QLineEdit(); self.f_new2.setEchoMode(QLineEdit.EchoMode.Password)
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

        _sess = AppSession.get()
        if _sess is None:
            return
        uid = _sess["id"]
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
        for name in ["用户管理", "修改密码", "数据备份"]:
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
        self._build_backup()
        L.addWidget(self.stack)
        self._switch("用户管理")

    def _switch(self, name):
        mapping = {"用户管理": 0, "修改密码": 1, "数据备份": 2}
        is_superadmin = AppSession.has_perm("system.manage")
        if name in ("用户管理", "数据备份") and not is_superadmin:
            name = "修改密码"
        self.stack.setCurrentIndex(mapping[name])
        for b in self._tabs:
            is_active = b.text() == name
            b.setProperty("active", "true" if is_active else "false")
            b.style().unpolish(b); b.style().polish(b)
            if b.text() in ("用户管理", "数据备份"):
                b.setVisible(is_superadmin)
        if name == "用户管理":
            self._load_users()
        elif name == "修改密码":
            self._refresh_change_pw()
        elif name == "数据备份":
            self._refresh_last_backup()
            self._refresh_pw_status()
            self._refresh_auto_backup_ui()

    def refresh_after_login(self):
        """登录后调用，重新刷新 SystemPage 的状态"""
        self._refresh_change_pw()
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
        hh.setStretchLastSection(False)
        self.user_tbl.setHorizontalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAsNeeded)
        self.user_tbl.setColumnWidth(0, 110)
        self.user_tbl.setColumnWidth(1, 160)   # 显示名称，可拖拽调整
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
        _cur_sess = AppSession.get()
        cur_uid = _cur_sess["id"] if _cur_sess else -1

        for i, u in enumerate(users):
            self.user_tbl.setRowHeight(i, 46)
            role_label = ROLE_LABELS.get(u["role"], u["role"])
            role_color = {"superadmin": "#3d6fdb", "admin": "#52c41a",
                          "accountant": "#fa8c16", "readonly": "#888"}.get(u["role"], "#555")
            active_text = "启用" if u["is_active"] else "停用"
            active_color = "#52c41a" if u["is_active"] else "#ff4d4f"
            last_login = (u["last_login"] or "从未登录")[:16]

            for j, (val, align, color) in enumerate([
                (u["username"],   Qt.AlignmentFlag.AlignCenter,              None),
                (u["display_name"], Qt.AlignmentFlag.AlignLeft|Qt.AlignmentFlag.AlignVCenter, None),
                (role_label,      Qt.AlignmentFlag.AlignCenter,              role_color),
                (active_text,     Qt.AlignmentFlag.AlignCenter,              active_color),
                (last_login,      Qt.AlignmentFlag.AlignCenter,              "#888"),
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
        il.addWidget(lbl("当前用户：", color="#888"))
        self.lbl_cur_user = lbl("", bold=True)
        il.addWidget(self.lbl_cur_user)
        il.addWidget(lbl("角色：", color="#888"))
        self.lbl_cur_role = lbl("", color="#3d6fdb", bold=True)
        il.addWidget(self.lbl_cur_role)
        il.addStretch()
        L.addWidget(info_frame)

        b_pw = QPushButton("修改登录密码")
        b_pw.setObjectName("btn_outline")
        b_pw.setFixedWidth(140)
        b_pw.clicked.connect(lambda: ChangePasswordDialog(self).exec())
        L.addWidget(b_pw)
        L.addStretch()
        self.stack.addWidget(w)

    def _refresh_change_pw(self):
        """切换到「修改密码」Tab 时，从 AppSession 动态读取当前用户信息并更新标签。"""
        user = AppSession.get() or {}
        self.lbl_cur_user.setText(user.get("display_name", ""))
        self.lbl_cur_role.setText(ROLE_LABELS.get(user.get("role", ""), ""))

    # ── Tab 3：数据备份（仅超级管理员）──────────────────────────────────────
    def _build_backup(self):
        # 外层套 QScrollArea，内容过多时可滚动
        outer = QWidget()
        outer_l = QVBoxLayout(outer)
        outer_l.setContentsMargins(0, 0, 0, 0)

        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        scroll.setFrameShape(QFrame.NoFrame)
        scroll.setHorizontalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAlwaysOff)

        w = QWidget()
        L = QVBoxLayout(w)
        L.setContentsMargins(24, 16, 24, 24)
        L.setSpacing(16)

        # ── 页标题 ──
        title_lbl = QLabel("数据备份与恢复")
        title_lbl.setStyleSheet("font-size:15px;font-weight:bold;color:#1a1a2e;")
        L.addWidget(title_lbl)

        intro = QLabel(
            "备份文件包含所有账套、凭证、用户数据，使用 AES-256-GCM 加密保护。\n"
            "备份密码独立于登录密码，由系统凭据管理器安全保管，无需每次手动输入。"
        )
        intro.setWordWrap(True)
        intro.setStyleSheet(
            "color:#555;font-size:12px;background:#f6f8fc;"
            "border-radius:6px;padding:10px 14px;")
        L.addWidget(intro)

        self.last_backup_lbl = QLabel("最近备份：查询中…")
        self.last_backup_lbl.setStyleSheet("color:#888;font-size:12px;")
        L.addWidget(self.last_backup_lbl)

        # ── 备份密码管理 ──────────────────────────────────
        pf = QFrame()
        pf.setStyleSheet("QFrame{background:#fff;border:1px solid #e4e8f0;border-radius:8px;}")
        pl = QVBoxLayout(pf)
        pl.setContentsMargins(20, 16, 20, 18)
        pl.setSpacing(10)

        ph = QHBoxLayout()
        pw_title = QLabel("备份密码")
        pw_title.setStyleSheet("font-size:13px;font-weight:bold;color:#1a1a2e;")
        ph.addWidget(pw_title)
        ph.addStretch()
        self.lbl_pw_status = QLabel()
        ph.addWidget(self.lbl_pw_status)
        pl.addLayout(ph)

        pw_note = QLabel(
            "密码存储于操作系统凭据管理器\n"
            "（Windows Credential Manager / macOS Keychain），\n"
            "与当前系统账户绑定，其他用户无法读取。\n"
            "⚠ 修改密码只影响此后生成的新备份，\n"
            "   已有备份文件仍需使用创建时的密码恢复。"
        )
        pw_note.setWordWrap(True)
        pw_note.setStyleSheet("color:#888;font-size:12px;")
        pl.addWidget(pw_note)

        b_set_pw = QPushButton("设置 / 修改备份密码…")
        b_set_pw.setFixedWidth(180)
        b_set_pw.setStyleSheet(
            "QPushButton{background:#fff;color:#3d6fdb;border:1px solid #3d6fdb;"
            "border-radius:4px;padding:6px 12px;font-size:13px;}"
            "QPushButton:hover{background:#e6f0ff;}")
        b_set_pw.clicked.connect(self._set_backup_password)
        pl.addWidget(b_set_pw)
        L.addWidget(pf)

        # ── 月末自动备份 ──────────────────────────────────
        af = QFrame()
        af.setStyleSheet("QFrame{background:#fff;border:1px solid #e4e8f0;border-radius:8px;}")
        al = QVBoxLayout(af)
        al.setContentsMargins(20, 16, 20, 18)
        al.setSpacing(10)

        ah = QHBoxLayout()
        auto_title = QLabel("月末自动备份")
        auto_title.setStyleSheet("font-size:13px;font-weight:bold;color:#1a1a2e;")
        ah.addWidget(auto_title)
        ah.addStretch()
        self.chk_auto = QCheckBox("启用")
        self.chk_auto.setStyleSheet("font-size:13px;color:#333;")
        self.chk_auto.setChecked(get_setting("auto_backup_enabled", "0") == "1")
        ah.addWidget(self.chk_auto)
        al.addLayout(ah)

        auto_note = QLabel(
            "每月最后一天自动备份一次。若当天未登录软件，\n"
            "下次登录时自动补跑（最多追溯 6 个月，每次补跑一个月份）。"
        )
        auto_note.setWordWrap(True)
        auto_note.setStyleSheet("color:#888;font-size:12px;")
        al.addWidget(auto_note)

        dir_row = QHBoxLayout()
        dir_lbl = QLabel("保存目录：")
        dir_lbl.setStyleSheet("color:#333;font-size:12px;")
        dir_lbl.setFixedWidth(72)
        dir_row.addWidget(dir_lbl)
        self.edit_auto_path = QLineEdit()
        self.edit_auto_path.setPlaceholderText("留空则使用程序目录下的 backups/ 子目录")
        self.edit_auto_path.setText(get_setting("auto_backup_path", ""))
        self.edit_auto_path.setStyleSheet(
            "border:1px solid #d9d9d9;border-radius:4px;"
            "padding:4px 8px;font-size:12px;background:#fff;color:#333;")
        dir_row.addWidget(self.edit_auto_path)
        b_browse = QPushButton("浏览…")
        b_browse.setFixedWidth(60)
        b_browse.setStyleSheet(
            "QPushButton{background:#fff;color:#3d6fdb;border:1px solid #3d6fdb;"
            "border-radius:4px;padding:4px 8px;font-size:12px;}"
            "QPushButton:hover{background:#e6f0ff;}")
        b_browse.clicked.connect(self._browse_auto_path)
        dir_row.addWidget(b_browse)
        al.addLayout(dir_row)

        self.lbl_last_auto = QLabel("最近自动备份：查询中…")
        self.lbl_last_auto.setStyleSheet("color:#888;font-size:12px;")
        al.addWidget(self.lbl_last_auto)

        b_save_auto = QPushButton("保存自动备份设置")
        b_save_auto.setFixedWidth(150)
        b_save_auto.setStyleSheet(
            "QPushButton{background:#3d6fdb;color:#fff;border:none;"
            "border-radius:4px;padding:6px 14px;font-size:13px;font-weight:bold;}"
            "QPushButton:hover{background:#2d5bc4;}")
        b_save_auto.clicked.connect(self._save_auto_backup_settings)
        al.addWidget(b_save_auto)
        L.addWidget(af)

        # ── 立即备份 ──────────────────────────────────────
        bf = QFrame()
        bf.setStyleSheet("QFrame{background:#fff;border:1px solid #e4e8f0;border-radius:8px;}")
        bfl = QVBoxLayout(bf)
        bfl.setContentsMargins(20, 16, 20, 18)
        bfl.setSpacing(8)
        bk_title = QLabel("立即备份")
        bk_title.setStyleSheet("font-size:13px;font-weight:bold;color:#1a1a2e;")
        bfl.addWidget(bk_title)
        bk_note = QLabel("将当前数据库加密备份到指定位置，建议定期备份并存储到外部设备或云盘。")
        bk_note.setWordWrap(True)
        bk_note.setStyleSheet("color:#555;font-size:12px;")
        bfl.addWidget(bk_note)
        b_bk = QPushButton("选择位置并备份…")
        b_bk.setFixedWidth(160)
        b_bk.setStyleSheet(
            "QPushButton{background:#3d6fdb;color:#fff;border:none;"
            "border-radius:4px;padding:6px 14px;font-size:13px;font-weight:bold;}"
            "QPushButton:hover{background:#2d5bc4;}")
        b_bk.clicked.connect(self._do_backup)
        bfl.addWidget(b_bk)
        L.addWidget(bf)

        # ── 从备份恢复 ────────────────────────────────────
        rf = QFrame()
        rf.setStyleSheet("QFrame{background:#fff;border:1px solid #f5c6cb;border-radius:8px;}")
        rfl = QVBoxLayout(rf)
        rfl.setContentsMargins(20, 16, 20, 18)
        rfl.setSpacing(8)
        rs_title = QLabel("从备份恢复")
        rs_title.setStyleSheet("font-size:13px;font-weight:bold;color:#c0392b;")
        rfl.addWidget(rs_title)
        w2 = QLabel("⚠ 恢复操作将完全覆盖当前数据库，所有未备份的数据将丢失，操作不可撤销。")
        w2.setWordWrap(True)
        w2.setStyleSheet("color:#c0392b;font-size:12px;")
        rfl.addWidget(w2)
        restore_note = QLabel(
            "同一台电脑上的备份将自动使用已保存的密码恢复；\n"
            "换电脑恢复时，程序将提示手动输入备份创建时所用的密码。"
        )
        restore_note.setWordWrap(True)
        restore_note.setStyleSheet("color:#888;font-size:12px;")
        rfl.addWidget(restore_note)
        b_rs = QPushButton("选择备份文件并恢复…")
        b_rs.setFixedWidth(190)
        b_rs.setStyleSheet(
            "QPushButton{background:#ff4d4f;color:#fff;border:none;"
            "border-radius:4px;padding:6px 14px;font-size:13px;font-weight:bold;}"
            "QPushButton:hover{background:#cf1322;}")
        b_rs.clicked.connect(self._do_restore)
        rfl.addWidget(b_rs)
        L.addWidget(rf)

        L.addStretch()
        scroll.setWidget(w)
        outer_l.addWidget(scroll)
        self.stack.addWidget(outer)
        self._refresh_last_backup()
        self._refresh_pw_status()
        self._refresh_auto_backup_ui()

    def _refresh_last_backup(self):
        """查询最近一次备份时间并更新标签。"""
        try:
            conn = get_db(); c = conn.cursor()
            c.execute("""SELECT created_at, operator FROM audit_log
                         WHERE client_id=0 AND action='数据备份'
                         ORDER BY id DESC LIMIT 1""")
            row = c.fetchone(); conn.close()
            if row:
                self.last_backup_lbl.setText(
                    f"最近备份：{row['created_at'][:16]}  操作人：{row['operator']}")
                self.last_backup_lbl.setStyleSheet("color:#389e0d;font-size:12px;")
            else:
                self.last_backup_lbl.setText("最近备份：尚未进行过备份")
                self.last_backup_lbl.setStyleSheet("color:#cf1322;font-size:12px;")
        except Exception:
            self.last_backup_lbl.setText("最近备份：查询失败")

    def _refresh_pw_status(self):
        """更新备份密码状态标签。"""
        if not kr_available():
            self.lbl_pw_status.setText("⚠ 凭据管理器不可用")
            self.lbl_pw_status.setStyleSheet("color:#fa8c16;font-size:12px;")
        elif kr_get():
            self.lbl_pw_status.setText("● 已设置")
            self.lbl_pw_status.setStyleSheet("color:#389e0d;font-size:12px;font-weight:bold;")
        else:
            self.lbl_pw_status.setText("● 未设置")
            self.lbl_pw_status.setStyleSheet("color:#cf1322;font-size:12px;font-weight:bold;")

    def _refresh_auto_backup_ui(self):
        """刷新自动备份区：复选框状态、目录、最近自动备份时间。"""
        self.chk_auto.setChecked(get_setting("auto_backup_enabled", "0") == "1")
        self.edit_auto_path.setText(get_setting("auto_backup_path", ""))
        try:
            conn = get_db(); c = conn.cursor()
            c.execute("""SELECT created_at, operator, target_id FROM audit_log
                         WHERE client_id=0 AND action='自动备份'
                         ORDER BY id DESC LIMIT 1""")
            row = c.fetchone(); conn.close()
            if row:
                self.lbl_last_auto.setText(
                    f"最近自动备份：{row['created_at'][:16]}  "
                    f"月份：{row['target_id']}  操作人：{row['operator']}")
                self.lbl_last_auto.setStyleSheet("color:#389e0d;font-size:12px;")
            else:
                self.lbl_last_auto.setText("最近自动备份：尚未执行过")
                self.lbl_last_auto.setStyleSheet("color:#888;font-size:12px;")
        except Exception:
            pass

    def _browse_auto_path(self):
        """弹出目录选择器，将选中路径填入自动备份目录框。"""
        d = QFileDialog.getExistingDirectory(self, "选择自动备份保存目录",
                                             self.edit_auto_path.text() or "")
        if d:
            self.edit_auto_path.setText(d)

    def _save_auto_backup_settings(self):
        """保存自动备份开关和目录到 settings 表。"""
        enabled = "1" if self.chk_auto.isChecked() else "0"
        path    = self.edit_auto_path.text().strip()
        set_setting("auto_backup_enabled", enabled)
        set_setting("auto_backup_path",    path)
        conn = get_db()
        log_action(conn, 0, "修改自动备份设置", "system", 0,
                   f"启用:{enabled} 目录:{path or '(默认)'}")
        conn.commit(); conn.close()
        QMessageBox.information(self, "已保存", "自动备份设置已保存。")

    def _set_backup_password(self):
        """设置或修改备份密码，存入系统凭据管理器。"""
        if not kr_available():
            QMessageBox.warning(
                self, "不可用",
                "当前环境无法访问系统凭据管理器。\n"
                "请确认已安装 keyring 库，或手动输入密码进行备份。")
            return

        existing = kr_get()
        title = "修改备份密码" if existing else "设置备份密码"

        # 修改时需先验证旧密码
        if existing:
            old, ok = QInputDialog.getText(
                self, title, "请输入当前备份密码以验证身份：", QLineEdit.EchoMode.Password)
            if not ok:
                return
            if old != existing:
                QMessageBox.warning(self, "验证失败", "当前备份密码不正确。")
                return

        pw1, ok1 = QInputDialog.getText(
            self, title, "请输入新备份密码（至少8位）：", QLineEdit.EchoMode.Password)
        if not ok1 or not pw1:
            return
        if len(pw1) < 8:
            QMessageBox.warning(self, "密码太短", "备份密码至少需要 8 位。")
            return
        pw2, ok2 = QInputDialog.getText(
            self, title, "请再次输入新备份密码：", QLineEdit.EchoMode.Password)
        if not ok2 or pw1 != pw2:
            QMessageBox.warning(self, "密码不一致", "两次输入的密码不一致，请重新操作。")
            return

        if kr_set(pw1):
            conn = get_db()
            action = "修改备份密码" if existing else "设置备份密码"
            log_action(conn, 0, action, "system", 0, "备份密码已更新至系统凭据管理器")
            conn.commit(); conn.close()
            self._refresh_pw_status()
            msg = (
                "备份密码已修改并保存至系统凭据管理器。\n\n"
                "⚠ 已有备份文件仍需使用原密码恢复，请妥善保管历史密码记录。"
                if existing else
                "备份密码已设置并保存至系统凭据管理器。\n"
                "今后备份时将自动使用此密码，无需每次手动输入。"
            )
            QMessageBox.information(self, "成功", msg)
        else:
            QMessageBox.critical(self, "保存失败",
                                 "密码无法存入系统凭据管理器，请检查系统权限。")

    def _do_backup(self):
        """执行备份：优先使用 keyring 中的密码，未设置则引导先设置。"""
        pw = kr_get()
        if not pw:
            # 首次使用，引导设置密码
            ret = QMessageBox.information(
                self, "尚未设置备份密码",
                "首次备份需要先设置备份密码。\n\n"
                "密码将安全保存在系统凭据管理器中，今后备份无需重复输入。\n\n"
                "点击「确定」立即设置。",
                QMessageBox.Ok | QMessageBox.Cancel, QMessageBox.Ok)
            if ret != QMessageBox.Ok:
                return
            self._set_backup_password()
            pw = kr_get()
            if not pw:
                return  # 用户取消了设置

        ts = datetime.now().strftime("%Y%m%d_%H%M%S")
        dest, _ = QFileDialog.getSaveFileName(
            self, "选择备份保存位置", f"WiseLedger备份_{ts}.zyac",
            "WiseLedger备份文件 (*.zyac)")
        if not dest:
            return
        try:
            encrypt_backup(DB_PATH, dest, pw)
            conn = get_db()
            log_action(conn, 0, "数据备份", "system", 0, f"备份文件：{dest}")
            conn.commit(); conn.close()
            self._refresh_last_backup()
            QMessageBox.information(self, "备份成功", f"备份已保存至：\n{dest}")
        except Exception as e:
            QMessageBox.critical(self, "备份失败", f"备份过程中发生错误：\n{e}")

    def _do_restore(self):
        """执行恢复：先用 keyring 密码尝试，失败则提示手动输入（换机场景）。"""
        reply = QMessageBox.warning(
            self, "⚠ 确认恢复",
            "恢复操作将完全覆盖当前数据库！\n\n"
            "所有未备份的数据（凭证、用户、账套）将永久丢失。\n\n"
            "确认要继续吗？",
            QMessageBox.Yes | QMessageBox.No, QMessageBox.No)
        if reply != QMessageBox.Yes:
            return

        src, _ = QFileDialog.getOpenFileName(
            self, "选择备份文件", "", "WiseLedger备份文件 (*.zyac)")
        if not src:
            return

        reply2 = QMessageBox.critical(
            self, "最终确认",
            "即将覆盖当前数据库，此操作不可撤销。\n\n确认执行恢复？",
            QMessageBox.Yes | QMessageBox.No, QMessageBox.No)
        if reply2 != QMessageBox.Yes:
            return

        # 优先用 keyring 密码，失败再让用户手动输入
        saved_pw = kr_get()
        candidates = [saved_pw] if saved_pw else []

        tmp_path = DB_PATH + ".restore_tmp"

        def _try_restore(pw: str) -> bool:
            """尝试用指定密码解密，成功返回 True，密码错误返回 False，其他异常上抛。"""
            try:
                decrypt_backup(src, tmp_path, pw)
                return True
            except ValueError:
                return False

        pw_used = None
        for pw in candidates:
            if _try_restore(pw):
                pw_used = pw
                break

        if pw_used is None:
            # keyring 密码不可用或不存在，提示手动输入（换机恢复场景）
            hint = (
                "未能用已保存的密码解密该备份文件。\n"
                "这通常发生在换了电脑或密码已被修改的情况下。\n\n"
                "请手动输入该备份文件创建时所用的密码："
                if saved_pw else
                "当前设备尚未保存备份密码，请手动输入该备份文件的密码："
            )
            pw_manual, ok = QInputDialog.getText(self, "输入备份密码", hint, QLineEdit.EchoMode.Password)
            if not ok or not pw_manual:
                if os.path.exists(tmp_path):
                    os.remove(tmp_path)
                return
            if not _try_restore(pw_manual):
                # 手动输入也错了 → 记录日志并提示
                conn = get_db()
                user = AppSession.get()
                log_action(conn, 0, "恢复备份失败", "system", 0,
                           f"密码错误或文件损坏，文件：{src}",
                           operator=user["username"] if user else "未知")
                conn.commit(); conn.close()
                QMessageBox.critical(self, "恢复失败",
                                     "密码错误或备份文件已损坏，恢复已取消。")
                if os.path.exists(tmp_path):
                    os.remove(tmp_path)
                return
            pw_used = pw_manual

        # 解密成功，替换数据库
        try:
            conn = get_db()
            log_action(conn, 0, "数据恢复", "system", 0, f"从备份恢复：{src}")
            conn.commit(); conn.close()
            if os.path.exists(DB_PATH):
                os.replace(DB_PATH, DB_PATH + ".pre_restore_bak")
            os.replace(tmp_path, DB_PATH)
            QMessageBox.information(
                self, "恢复成功",
                "数据库已恢复成功。\n\n"
                "请关闭并重新启动软件以使恢复生效。\n"
                "（原数据库已保存为 accounting.db.pre_restore_bak）"
            )
        except Exception as e:
            QMessageBox.critical(self, "恢复失败", f"恢复过程中发生错误：\n{e}")
            if os.path.exists(tmp_path):
                os.remove(tmp_path)