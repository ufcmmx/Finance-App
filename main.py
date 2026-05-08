"""main.py — 应用入口，仅包含 MainWindow 和 main()"""
import sys, os

# ── Windows 打包必须：防止 multiprocessing 产生子进程控制台窗口 ──
if getattr(sys, 'frozen', False):
    import multiprocessing
    multiprocessing.freeze_support()

_here = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, _here)
from datetime import datetime

from PySide6.QtWidgets import *
from PySide6.QtCore import Qt, QDate, Signal, QTimer
from PySide6.QtGui import QColor, QFont, QBrush, QPalette

from db import init_db, get_db, log_action, STANDARD_ACCOUNTS_SMALL
from utils import SS, lbl
from session import AppSession, ROLE_LABELS
from login_dialog import LoginDialog
from dialogs import ImportAccountSetDialog
from pages.client  import ClientPage
from pages.opening import OpeningBalancePage
from pages.voucher import VoucherPage
from pages.account import AccountPage
from pages.settle  import SettlePage
from pages.report  import ReportPage
from pages.audit   import AuditPage
from pages.system  import SystemPage

# 导航菜单：名称 → (stack索引, 所需权限)
_NAV_ITEMS = [
    ("客户管理",    0, "client.view"),
    ("科目管理",    1, "account.manage"),
    ("科目期初",    2, "opening.manage"),
    ("记账（凭证）", 3, "voucher.view"),
    ("期末结账",    4, "settle.manage"),
    ("财务报表",    5, "report.view"),
    ("审计日志",    6, "audit.view"),
    ("系统管理",    7, "system.manage"),
]

class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("智一会计 · 本地版")
        self.setMinimumSize(1150, 720)
        self._cur_client = None; self._cur_name = ""; self._cur_period = ""
        self._build()

    def _build(self):
        root = QWidget(); root.setObjectName("root")
        self.setCentralWidget(root)
        row = QHBoxLayout(root); row.setSpacing(0); row.setContentsMargins(0,0,0,0)

        # ── Sidebar ──
        sb = QWidget(); sb.setObjectName("sidebar")
        sb.setFixedWidth(196)
        sl = QVBoxLayout(sb); sl.setContentsMargins(0,0,0,0); sl.setSpacing(0)
        logo = QLabel("智一会计"); logo.setObjectName("logo")
        logo.setStyleSheet("color:#fff;font-size:18px;font-weight:bold;padding:22px 20px 4px 20px;")
        sub = QLabel("本地专业版"); sub.setObjectName("subt")
        sub.setStyleSheet("color:#445;font-size:11px;padding:0 20px 14px 20px;")
        sl.addWidget(logo); sl.addWidget(sub)
        div = QFrame(); div.setFrameShape(QFrame.HLine)
        div.setStyleSheet("background:#2a3255;max-height:1px;margin:0 16px 8px 16px;")
        sl.addWidget(div)

        # 导航按钮 — 按权限决定是否显示
        self._nav_btns = []
        for name, _, perm in _NAV_ITEMS:
            b = QPushButton(name); b.setObjectName("nav"); b.setProperty("active","false")
            b.clicked.connect(lambda _,n=name: self._nav(n))
            b.setVisible(AppSession.has_perm(perm))
            sl.addWidget(b); self._nav_btns.append(b)

        sl.addStretch()

        # 当前客户信息
        self._client_info = QLabel(""); self._client_info.setWordWrap(True)
        self._client_info.setStyleSheet("color:#556;font-size:11px;padding:6px 16px;")
        sl.addWidget(self._client_info)

        # 底部：登录用户信息 + 退出
        user_bar = QWidget()
        user_bar.setStyleSheet("background:#151b30;")
        ul = QVBoxLayout(user_bar); ul.setContentsMargins(14,8,14,10); ul.setSpacing(2)
        self._user_lbl = QLabel(
            f"{AppSession.display_name()}  [{AppSession.role_label()}]")
        self._user_lbl.setStyleSheet("color:#8b93ae;font-size:11px;")
        self._user_lbl.setWordWrap(True)
        b_logout = QPushButton("退出登录")
        b_logout.setStyleSheet("""QPushButton{background:transparent;color:#556;
            border:none;font-size:11px;text-align:left;padding:2px 0;}
            QPushButton:hover{color:#e05252;}""")
        b_logout.clicked.connect(self._logout)
        ul.addWidget(self._user_lbl); ul.addWidget(b_logout)
        sl.addWidget(user_bar)
        row.addWidget(sb)

        # Content
        self.stack = QStackedWidget(); row.addWidget(self.stack)
        self.pg_clients  = ClientPage(self.stack)
        self.pg_opening  = OpeningBalancePage(self.stack)
        self.pg_vouchers = VoucherPage(self.stack)
        self.pg_accounts = AccountPage(self.stack)
        self.pg_settle = SettlePage(self.stack)
        self.pg_reports = ReportPage(self.stack)
        self.pg_audit = AuditPage(self.stack)
        self.pg_system = SystemPage(self.stack)
        for pg in [self.pg_clients, self.pg_accounts, self.pg_opening,
                   self.pg_vouchers, self.pg_settle, self.pg_reports,
                   self.pg_audit, self.pg_system]:
            self.stack.addWidget(pg)
        self.pg_clients.client_opened.connect(self._open_client)
        self.pg_settle.carryforward_done.connect(self._on_carryforward_done)
        self._nav("客户管理")

    def _on_carryforward_done(self):
        """After carryforward, switch to voucher page and refresh so user can see new vouchers."""
        self.pg_vouchers._switch_tab("查凭证")
        self._nav("记账（凭证）")

    def _nav(self, name):
        mapping = {item[0]: item[1] for item in _NAV_ITEMS}
        perm_map = {item[0]: item[2] for item in _NAV_ITEMS}
        if name not in mapping:
            return
        # 二次权限校验（防绕过）
        if not AppSession.has_perm(perm_map[name]):
            QMessageBox.warning(self, "无权限", f"您没有访问【{name}】的权限")
            return
        self.stack.setCurrentIndex(mapping[name])
        for b in self._nav_btns:
            b.setProperty("active","true" if b.text()==name else "false")
            b.style().unpolish(b); b.style().polish(b)
        if name=="客户管理": self.pg_clients.load()

    def _open_client(self, client_id, name, code):
        # 会计角色检查客户授权
        if not AppSession.can_access_client(client_id):
            QMessageBox.warning(self, "无权限", f"您没有访问客户【{name}】的权限")
            return
        self._cur_client = client_id; self._cur_name = name
        now = datetime.now()
        self._cur_period = f"{now.year}-{now.month:02d}"
        self._client_info.setText(f"当前客户:\n{name}\n({code})")
        self.pg_vouchers.set_client(client_id, name, self._cur_period)
        self.pg_accounts.set_client(client_id)
        self.pg_opening.set_client(client_id, name, self._cur_period)
        self.pg_settle.set_client(client_id, name, self._cur_period)
        self.pg_reports.set_client(client_id, name, self._cur_period)
        self.pg_audit.set_client(client_id)
        conn = get_db()
        log_action(conn, client_id, "打开账套", "client", client_id, f"客户: {name}")
        conn.commit(); conn.close()
        self._nav("记账（凭证）")

    def _logout(self):
        reply = QMessageBox.question(self, "退出登录", "确认退出登录？",
                                     QMessageBox.Yes | QMessageBox.No)
        if reply != QMessageBox.Yes:
            return
        AppSession.logout()
        self.close()
        # 重新显示登录窗口
        from login_dialog import LoginDialog
        dlg = LoginDialog()
        if dlg.exec() == QDialog.Accepted:
            w = MainWindow()
            w.show()
            # 把新窗口引用存到 app，防止被 GC
            QApplication.instance()._main_window = w
        else:
            QApplication.instance().quit()


def main():
    import traceback
    _log_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "startup.log")
    _is_frozen = getattr(sys, 'frozen', False)
    def _wlog(msg):
        try:
            with open(_log_path, "a", encoding="utf-8") as _f:
                _f.write(msg + "\n")
            # 打包后不写 stderr，避免 Windows 弹出控制台窗口
            if not _is_frozen:
                print(msg, file=sys.stderr, flush=True)
        except Exception:
            pass
    from datetime import datetime as _dt
    _wlog(f"\n=== 启动 {_dt.now()} ===")
    try:
        _wlog("step 1: init_db")
        init_db()
        _wlog("step 2: QApplication")
        app = QApplication(sys.argv)
        _wlog("step 3: setStyleSheet")
        app.setStyleSheet(SS)

        # ── 登录 ──
        _wlog("step 4: LoginDialog")
        login_dlg = LoginDialog()
        if login_dlg.exec() != QDialog.Accepted:
            _wlog("用户取消登录，退出")
            sys.exit(0)

        _wlog("step 5: MainWindow()")
        w = MainWindow()
        app._main_window = w   # 防止被 GC
        _wlog("step 6: w.show()")
        w.show()
        _wlog("step 7: entering event loop")
        sys.exit(app.exec())
    except Exception:
        tb = traceback.format_exc()
        _wlog("EXCEPTION: " + tb)
        try:
            _app = QApplication.instance() or QApplication(sys.argv)
            QMessageBox.critical(None, "启动错误", tb[:2000])
        except Exception as e2:
            _wlog("Dialog also failed: " + str(e2))
        sys.exit(1)

if __name__ == "__main__":
    main()