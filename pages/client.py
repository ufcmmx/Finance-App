"""pages/client.py — ClientPage — 客户账套管理"""
from datetime import datetime
from PySide6.QtWidgets import *
from PySide6.QtCore import Qt, QDate, Signal, QTimer
from PySide6.QtGui import QColor, QFont, QBrush, QPalette

from db import get_db, log_action
from utils import lbl, sep, card, fmt_amt, NoScrollSpinBox, NoScrollDoubleSpinBox
from dialogs import ClientDialog, ImportAccountSetDialog
from session import AppSession
# openpyxl imported lazily inside each export function

# 客户类型 → (胶囊底色, 胶囊文字色)
_TYPE_COLORS = {
    "小规模纳税人": ("#e8f1fd", "#1e6fb8"),  # 浅蓝
    "一般纳税人":   ("#e8f7ee", "#1e7a3c"),  # 浅绿
    "其他":         ("#f0f2f5", "#6b7280"),  # 浅灰
}
_TYPE_DEFAULT = ("#f0f2f5", "#9ca3af")


def _type_pill(text: str) -> QWidget:
    """生成居中的彩色胶囊标签。"""
    wrap = QWidget()
    lay = QHBoxLayout(wrap); lay.setContentsMargins(6, 0, 6, 0)
    if text:
        bg, fg = _TYPE_COLORS.get(text, _TYPE_DEFAULT)
        pill = QLabel(text)
        pill.setStyleSheet(
            f"background:{bg}; color:{fg}; border-radius:10px;"
            f"padding:2px 10px; font-size:12px; font-weight:600;"
        )
        pill.setAlignment(Qt.AlignmentFlag.AlignCenter)
        lay.addWidget(pill, alignment=Qt.AlignmentFlag.AlignCenter)
    else:
        dash = QLabel("—"); dash.setStyleSheet("color:#cbd0d9;")
        dash.setAlignment(Qt.AlignmentFlag.AlignCenter)
        lay.addWidget(dash, alignment=Qt.AlignmentFlag.AlignCenter)
    return wrap


class ClientPage(QWidget):
    client_opened = Signal(int, str, str)

    def __init__(self, parent=None):
        super().__init__(parent)
        L = QVBoxLayout(self); L.setContentsMargins(24,20,24,20); L.setSpacing(14)
        hdr = QHBoxLayout()
        hdr.addWidget(lbl("客户列表", bold=True, size=18)); hdr.addStretch()
        self.b_imp = QPushButton("导入账套"); self.b_imp.setObjectName("btn_outline")
        self.b_imp.clicked.connect(self._import_account_set)
        self.b_imp.setVisible(AppSession.has_perm("client.manage"))
        self.b_add = QPushButton("＋ 新建客户"); self.b_add.setObjectName("btn_primary"); self.b_add.clicked.connect(self._add)
        self.b_add.setVisible(AppSession.has_perm("client.manage"))
        hdr.addWidget(self.b_imp); hdr.addWidget(self.b_add); L.addLayout(hdr)
        self.search = QLineEdit(); self.search.setPlaceholderText("搜索客户名称或助记码...")
        self.search.textChanged.connect(self.load)
        L.addWidget(self.search)
        f = card(); vl = QVBoxLayout(f); vl.setContentsMargins(0,0,0,0)
        self.tbl = QTableWidget(); self.tbl.setEditTriggers(QTableWidget.NoEditTriggers)
        self.tbl.setSelectionBehavior(QTableWidget.SelectRows); self.tbl.setShowGrid(False)
        self.tbl.verticalHeader().setVisible(False)
        self.tbl.setColumnCount(7)
        self.tbl.setHorizontalHeaderLabels(["序号","客户名称","助记码","客户类型","税号","联系人","操作"])
        hh = self.tbl.horizontalHeader()
        hh.setSectionResizeMode(QHeaderView.Interactive)   # 全部列均可拖拽调整宽度
        hh.setMinimumSectionSize(40)
        hh.setStretchLastSection(True)   # 操作列吸收剩余宽度，避免按钮被截断
        hh.setFixedHeight(44)
        self.tbl.setColumnWidth(0, 52); self.tbl.setColumnWidth(1, 200)
        self.tbl.setColumnWidth(2, 80); self.tbl.setColumnWidth(3, 135)
        self.tbl.setColumnWidth(4, 140); self.tbl.setColumnWidth(5, 90)
        # 第 6 列（操作）由 stretchLastSection 自动占满剩余
        self.tbl.setHorizontalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAsNeeded)
        vl.addWidget(self.tbl); L.addWidget(f)

    def load(self):
        kw = self.search.text().strip()
        conn = get_db(); c = conn.cursor()
        user = AppSession.get()
        role = user["role"] if user else ""

        # 更新顶部按钮的可见性
        can_manage = AppSession.has_perm("client.manage")
        self.b_imp.setVisible(can_manage)
        self.b_add.setVisible(can_manage)

        if role in ("superadmin", "admin"):
            # 管理员以上看全部客户
            if kw:
                c.execute("SELECT * FROM clients WHERE name LIKE ? OR short_code LIKE ? ORDER BY id",
                          (f"%{kw}%", f"%{kw}%"))
            else:
                c.execute("SELECT * FROM clients ORDER BY id")
        else:
            # 会计/只读只看授权的客户
            uid = user["id"] if user else -1
            if kw:
                c.execute("""SELECT cl.* FROM clients cl
                             JOIN user_client_access uca ON uca.client_id=cl.id
                             WHERE uca.user_id=?
                             AND (cl.name LIKE ? OR cl.short_code LIKE ?)
                             ORDER BY cl.id""",
                          (uid, f"%{kw}%", f"%{kw}%"))
            else:
                c.execute("""SELECT cl.* FROM clients cl
                             JOIN user_client_access uca ON uca.client_id=cl.id
                             WHERE uca.user_id=? ORDER BY cl.id""", (uid,))

        rows = c.fetchall(); conn.close()
        self.tbl.setRowCount(len(rows))
        for i,r in enumerate(rows):
            self.tbl.setRowHeight(i,56)
            # Index badge — 恢复原色（类型已通过 col 3 的彩色胶囊展示）
            badge = QLabel(f"  {r['id']:02d}  ")
            badge.setStyleSheet("background:#f0f4ff;color:#3d6fdb;border-radius:4px;font-size:11px;")
            badge.setAlignment(Qt.AlignmentFlag.AlignCenter)
            self.tbl.setCellWidget(i,0,badge)
            # 客户名称 / 助记码
            for j,v in enumerate([r['name'],r['short_code'] or ''],1):
                it = QTableWidgetItem(v); it.setTextAlignment(Qt.AlignmentFlag.AlignCenter)
                it.setData(Qt.ItemDataRole.UserRole, r['id']); self.tbl.setItem(i,j,it)
            # 客户类型 — 彩色胶囊
            self.tbl.setCellWidget(i, 3, _type_pill(r['client_type'] or ''))
            # 税号 / 联系人
            for j,v in enumerate([r['tax_id'] or '', r['contact'] or ''], 4):
                it = QTableWidgetItem(v); it.setTextAlignment(Qt.AlignmentFlag.AlignCenter)
                it.setData(Qt.ItemDataRole.UserRole, r['id']); self.tbl.setItem(i,j,it)
            # Buttons — 三个按钮统一尺寸；走全局 padding（已收紧到 6px 12px）
            bw = QWidget()
            bw.setObjectName("btnRow"); bw.setStyleSheet("#btnRow { background:#ffffff; }")
            bl = QHBoxLayout(bw); bl.setContentsMargins(8,4,8,4); bl.setSpacing(8)
            b1 = QPushButton("进账簿"); b1.setObjectName("btn_primary")
            b1.setFixedSize(76, 30)
            b2 = QPushButton("编辑"); b2.setObjectName("btn_outline")
            b2.setFixedSize(60, 30)
            b2.setVisible(can_manage)
            b3 = QPushButton("删除"); b3.setObjectName("btn_red")
            b3.setFixedSize(60, 30)
            b3.setVisible(can_manage)
            b1.clicked.connect(lambda _,rr=r: self.client_opened.emit(rr['id'],rr['name'],rr['short_code'] or ''))
            b2.clicked.connect(lambda _,rr=r: self._edit(rr))
            b3.clicked.connect(lambda _,rr=r: self._del(rr))
            bl.addWidget(b1); bl.addWidget(b2); bl.addWidget(b3); bl.addStretch()
            self.tbl.setCellWidget(i,6,bw)

    def _import_account_set(self):
        d = ImportAccountSetDialog(self)
        d.exec(); self.load()

    def _add(self):
        d = ClientDialog(self)
        if d.exec(): self.load()

    def _edit(self,r):
        d = ClientDialog(self, r)
        if d.exec(): self.load()

    def _del(self,r):
        if QMessageBox.question(self,"确认",f"删除 [{r['name']}]？所有账目数据一并删除。",
                                QMessageBox.Yes|QMessageBox.No) == QMessageBox.Yes:
            conn = get_db()
            try:
                client_id = r['id']
                # Delete dependent rows explicitly because most FKs are NO ACTION.
                conn.execute("DELETE FROM voucher_entries WHERE voucher_id IN (SELECT id FROM vouchers WHERE client_id=?)",
                             (client_id,))
                conn.execute("DELETE FROM voucher_templates WHERE client_id=?", (client_id,))
                conn.execute("DELETE FROM bank_statements WHERE client_id=?", (client_id,))
                conn.execute("DELETE FROM account_aux_config WHERE client_id=?", (client_id,))
                conn.execute("DELETE FROM aux_items WHERE client_id=?", (client_id,))
                conn.execute("DELETE FROM aux_dimensions WHERE client_id=?", (client_id,))
                conn.execute("DELETE FROM periods WHERE client_id=?", (client_id,))
                conn.execute("DELETE FROM audit_log WHERE client_id=?", (client_id,))
                conn.execute("DELETE FROM vouchers WHERE client_id=?", (client_id,))
                conn.execute("DELETE FROM accounts WHERE client_id=?", (client_id,))
                conn.execute("DELETE FROM clients WHERE id=?", (client_id,))
                conn.commit()
            except Exception:
                conn.rollback()
                raise
            finally:
                conn.close()
            self.load()


