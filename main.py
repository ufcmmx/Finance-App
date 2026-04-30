"""main.py — 应用入口，仅包含 MainWindow 和 main()"""
import sys, os

# ── 启动前清理旧的字节码缓存，防止 .pyc 不一致导致静默崩溃 ──
# ── 启动前清理旧的字节码缓存，防止 .pyc 不一致导致静默崩溃 ──
_here = os.path.dirname(os.path.abspath(__file__))
_pycache = os.path.join(_here, '__pycache__')
if os.path.isdir(_pycache):
    import shutil, glob
    for _f in glob.glob(os.path.join(_pycache, 'main.*.pyc')):
        try: os.remove(_f)
        except: pass


sys.path.insert(0, _here)
from datetime import datetime

from PySide6.QtWidgets import *
from PySide6.QtCore import Qt, QDate, Signal, QTimer
from PySide6.QtGui import QColor, QFont, QBrush, QPalette

from db import init_db, get_db, log_action, STANDARD_ACCOUNTS_SMALL
from utils import SS, lbl
from dialogs import ImportAccountSetDialog
from pages.client  import ClientPage
from pages.voucher import VoucherPage
from pages.account import AccountPage
from pages.settle  import SettlePage
from pages.report  import ReportPage
from pages.audit   import AuditPage
from pages.system  import SystemPage

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

        # Sidebar
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
        self._nav_btns = []
        for name in ["客户管理","科目管理","记账（凭证）","期末结账","财务报表","审计日志","系统管理"]:
            b = QPushButton(name); b.setObjectName("nav"); b.setProperty("active","false")
            b.clicked.connect(lambda _,n=name: self._nav(n))
            sl.addWidget(b); self._nav_btns.append(b)
        sl.addStretch()
        self._client_info = QLabel(""); self._client_info.setWordWrap(True)
        self._client_info.setStyleSheet("color:#556;font-size:11px;padding:10px 16px;")
        sl.addWidget(self._client_info)
        row.addWidget(sb)

        # Content
        self.stack = QStackedWidget(); row.addWidget(self.stack)
        self.pg_clients = ClientPage()
        self.pg_vouchers = VoucherPage()
        self.pg_accounts = AccountPage()
        self.pg_settle = SettlePage()
        self.pg_reports = ReportPage()
        self.pg_audit = AuditPage()
        self.pg_system = SystemPage()
        for pg in [self.pg_clients, self.pg_accounts, self.pg_vouchers,
                   self.pg_settle, self.pg_reports, self.pg_audit, self.pg_system]:
            self.stack.addWidget(pg)
        self.pg_clients.client_opened.connect(self._open_client)
        self.pg_settle.carryforward_done.connect(self._on_carryforward_done)
        self._nav("客户管理")

    def _on_carryforward_done(self):
        """After carryforward, switch to voucher page and refresh so user can see new vouchers."""
        self.pg_vouchers._switch_tab("查凭证")
        self._nav("记账（凭证）")

    def _nav(self, name):
        mapping = {"客户管理":0,"科目管理":1,"记账（凭证）":2,"期末结账":3,
                   "财务报表":4,"审计日志":5,"系统管理":6}
        self.stack.setCurrentIndex(mapping[name])
        for b in self._nav_btns:
            b.setProperty("active","true" if b.text()==name else "false")
            b.style().unpolish(b); b.style().polish(b)
        if name=="客户管理": self.pg_clients.load()

    def _open_client(self, client_id, name, code):
        self._cur_client = client_id; self._cur_name = name
        now = datetime.now()
        self._cur_period = f"{now.year}-{now.month:02d}"
        self._client_info.setText(f"当前客户:\n{name}\n({code})")
        self.pg_vouchers.set_client(client_id, name, self._cur_period)
        self.pg_accounts.set_client(client_id)
        self.pg_settle.set_client(client_id, name, self._cur_period)
        self.pg_reports.set_client(client_id, name, self._cur_period)
        self.pg_audit.set_client(client_id)
        # Log client open
        conn = get_db()
        log_action(conn, client_id, "打开账套", "client", client_id, f"客户: {name}")
        conn.commit(); conn.close()
        self._nav("记账（凭证）")


def main():
    import traceback, atexit
    # Write log file next to the script for silent-crash debugging
    _log_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "startup.log")
    def _wlog(msg):
        try:
            with open(_log_path, "a", encoding="utf-8") as _f:
                _f.write(msg + "\n")
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
        _wlog("step 4: MainWindow()")
        w = MainWindow()
        _wlog("step 5: w.show()")
        w.show()
        _wlog("step 6: entering event loop")
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
