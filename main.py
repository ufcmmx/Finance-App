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

_PAGE_FACTORIES = {
    "客户管理": ClientPage,
    "科目管理": AccountPage,
    "科目期初": OpeningBalancePage,
    "记账（凭证）": VoucherPage,
    "期末结账": SettlePage,
    "财务报表": ReportPage,
    "审计日志": AuditPage,
    "系统管理": SystemPage,
}

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

        # 导航按钮 — 初始全部隐藏，登录后 _refresh_for_login() 按权限显示
        self._nav_btns = []
        for name, _, perm in _NAV_ITEMS:
            b = QPushButton(name); b.setObjectName("nav"); b.setProperty("active","false")
            b.clicked.connect(lambda _,n=name: self._nav(n))
            b.setVisible(False)
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
        self._user_lbl = QLabel("")
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

        # Content — 主页面全部按需创建，避免登录完成前触发 Windows 原生弹层闪窗
        self.stack = QStackedWidget(); row.addWidget(self.stack)
        self._pages = {}
        self._page_placeholders = {}
        for name, idx, _ in _NAV_ITEMS:
            placeholder = QWidget()
            self._page_placeholders[name] = placeholder
            self.stack.insertWidget(idx, placeholder)

    def closeEvent(self, event):
        """主窗口关闭时退出整个 app"""
        QApplication.instance().quit()
        event.accept()

    def _refresh_for_login(self):
        """登录成功后：按权限显示导航、更新用户名、跳转首页"""
        for b, (_, __, perm) in zip(self._nav_btns, _NAV_ITEMS):
            b.setVisible(AppSession.has_perm(perm))
        self._user_lbl.setText(
            f"{AppSession.display_name()}  [{AppSession.role_label()}]")
        self._nav("客户管理")

    def _reset_after_logout(self):
        """清理登录态相关界面，避免重新登录前短暂露出旧界面。"""
        self._cur_client = None
        self._cur_name = ""
        self._cur_period = ""
        self._client_info.setText("")
        for b in self._nav_btns:
            b.setVisible(False)
            b.setProperty("active", "false")
            b.style().unpolish(b)
            b.style().polish(b)
        self._user_lbl.setText("")
        self._client_info.setText("")
        self.stack.setCurrentIndex(0)

    def _on_carryforward_done(self):
        """After carryforward, switch to voucher page and refresh so user can see new vouchers."""
        voucher_page = self._ensure_page("记账（凭证）")
        voucher_page._switch_tab("查凭证")
        self._nav("记账（凭证）")

    def _ensure_page(self, name):
        """首次访问时再创建页面，避免登录完成时批量注册顶层原生子窗。"""
        page = self._pages.get(name)
        if page is not None:
            return page

        page = _PAGE_FACTORIES[name]()
        idx = next(item[1] for item in _NAV_ITEMS if item[0] == name)
        placeholder = self._page_placeholders.pop(name, None)
        if placeholder is not None:
            self.stack.removeWidget(placeholder)
            placeholder.deleteLater()
        self.stack.insertWidget(idx, page)
        self._pages[name] = page

        if name == "客户管理":
            page.client_opened.connect(self._open_client)
        elif name == "期末结账":
            page.carryforward_done.connect(self._on_carryforward_done)

        if self._cur_client is not None:
            if name == "科目管理":
                page.set_client(self._cur_client)
            elif name in ("科目期初", "期末结账", "财务报表"):
                page.set_client(self._cur_client, self._cur_name, self._cur_period)
            elif name == "记账（凭证）":
                page.set_client(self._cur_client, self._cur_name, self._cur_period)
            elif name == "审计日志":
                page.set_client(self._cur_client)

        return page

    def _nav(self, name):
        mapping = {item[0]: item[1] for item in _NAV_ITEMS}
        perm_map = {item[0]: item[2] for item in _NAV_ITEMS}
        if name not in mapping:
            return
        # 二次权限校验（防绕过）
        if not AppSession.has_perm(perm_map[name]):
            QMessageBox.warning(self, "无权限", f"您没有访问【{name}】的权限")
            return
        page = self._ensure_page(name)
        self.stack.setCurrentIndex(mapping[name])
        for b in self._nav_btns:
            b.setProperty("active","true" if b.text()==name else "false")
            b.style().unpolish(b); b.style().polish(b)
        if name == "客户管理":
            page.load()

    def _open_client(self, client_id, name, code):
        # 会计角色检查客户授权
        if not AppSession.can_access_client(client_id):
            QMessageBox.warning(self, "无权限", f"您没有访问客户【{name}】的权限")
            return
        self._cur_client = client_id; self._cur_name = name
        now = datetime.now()
        self._cur_period = f"{now.year}-{now.month:02d}"
        self._client_info.setText(f"当前客户:\n{name}\n({code})")
        for page_name, page in self._pages.items():
            if page_name == "科目管理":
                page.set_client(client_id)
            elif page_name in ("科目期初", "记账（凭证）", "期末结账", "财务报表"):
                page.set_client(client_id, name, self._cur_period)
            elif page_name == "审计日志":
                page.set_client(client_id)
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
        self._reset_after_logout()
        self.hide()
        # 重新登录时不绑定到主窗口，避免 Windows 在父子窗口切换时闪出底层窗口
        dlg = LoginDialog()
        self._relogin_dlg = dlg   # 防止被 GC
        def _on_done(result):
            self._relogin_dlg = None
            if result == QDialog.Accepted:
                self._refresh_for_login()
                self.show()
                self.raise_()
                self.activateWindow()
            else:
                QApplication.instance().quit()
        dlg.finished.connect(_on_done)
        dlg.open()


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

        # ── 先建主窗口但保持隐藏，登录成功后再显示，避免 Windows 闪出底层窗口 ──
        _wlog("step 4: MainWindow()")
        w = MainWindow()
        app._main_window = w

        _wlog("step 6: LoginDialog.open()")
        login_dlg = LoginDialog()
        app._login_dlg = login_dlg   # 防止被 GC

        def _on_login_done(result):
            if result != QDialog.Accepted:
                _wlog("用户取消登录，退出")
                app.quit()
                return
            _wlog("step 7: _refresh_for_login")
            w._refresh_for_login()
            _wlog("step 8: w.show()")
            w.show()
            w.raise_()
            w.activateWindow()

        login_dlg.finished.connect(_on_login_done)
        login_dlg.open()   # 返回即进入 app.exec()，无局部事件循环

        _wlog("step 9: entering event loop")
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
