"""pages/client.py — ClientPage — 客户账套管理"""
from datetime import datetime
from PySide6.QtWidgets import *
from PySide6.QtCore import Qt, QDate, Signal, QTimer
from PySide6.QtGui import QColor, QFont, QBrush, QPalette

from db import get_db, log_action
from utils import lbl, sep, card, fmt_amt, NoScrollSpinBox, NoScrollDoubleSpinBox
from dialogs import ClientDialog, ImportAccountSetDialog
# openpyxl imported lazily inside each export function

class ClientPage(QWidget):
    client_opened = Signal(int, str, str)

    def __init__(self):
        super().__init__()
        L = QVBoxLayout(self); L.setContentsMargins(24,20,24,20); L.setSpacing(14)
        hdr = QHBoxLayout()
        hdr.addWidget(lbl("客户列表", bold=True, size=18)); hdr.addStretch()
        b_imp = QPushButton("导入账套"); b_imp.setObjectName("btn_outline")
        b_imp.clicked.connect(self._import_account_set)
        b = QPushButton("＋ 新建客户"); b.setObjectName("btn_primary"); b.clicked.connect(self._add)
        hdr.addWidget(b_imp); hdr.addWidget(b); L.addLayout(hdr)
        self.search = QLineEdit(); self.search.setPlaceholderText("搜索客户名称或助记码...")
        self.search.textChanged.connect(self.load)
        L.addWidget(self.search)
        f = card(); vl = QVBoxLayout(f); vl.setContentsMargins(0,0,0,0)
        self.tbl = QTableWidget(); self.tbl.setEditTriggers(QTableWidget.NoEditTriggers)
        self.tbl.setSelectionBehavior(QTableWidget.SelectRows); self.tbl.setShowGrid(False)
        self.tbl.verticalHeader().setVisible(False)
        self.tbl.setColumnCount(7)
        self.tbl.setHorizontalHeaderLabels(["","客户名称","助记码","客户类型","税号","联系人","操作"])
        hh = self.tbl.horizontalHeader()
        hh.setSectionResizeMode(QHeaderView.Interactive)   # all columns user-draggable
        hh.setSectionResizeMode(1, QHeaderView.Stretch)    # name column stretches
        hh.setMinimumSectionSize(40)
        self.tbl.setColumnWidth(0, 44); self.tbl.setColumnWidth(2, 80)
        self.tbl.setColumnWidth(3, 110); self.tbl.setColumnWidth(4, 140)
        self.tbl.setColumnWidth(5, 90); self.tbl.setColumnWidth(6, 300)
        vl.addWidget(self.tbl); L.addWidget(f)

    def load(self):
        kw = self.search.text().strip()
        conn = get_db(); c = conn.cursor()
        if kw:
            c.execute("SELECT * FROM clients WHERE name LIKE ? OR short_code LIKE ? ORDER BY id",
                      (f"%{kw}%",f"%{kw}%"))
        else:
            c.execute("SELECT * FROM clients ORDER BY id")
        rows = c.fetchall(); conn.close()
        self.tbl.setRowCount(len(rows))
        for i,r in enumerate(rows):
            self.tbl.setRowHeight(i,50)
            # Index badge
            badge = QLabel(f"  {r['id']:02d}  ")
            badge.setStyleSheet("background:#f0f4ff;color:#3d6fdb;border-radius:4px;font-size:11px;")
            badge.setAlignment(Qt.AlignCenter)
            self.tbl.setCellWidget(i,0,badge)
            for j,v in enumerate([r['name'],r['short_code'] or '',r['client_type'] or '',
                                   r['tax_id'] or '',r['contact'] or ''],1):
                it = QTableWidgetItem(v); it.setTextAlignment(Qt.AlignCenter)
                it.setData(Qt.UserRole, r['id']); self.tbl.setItem(i,j,it)
            # Buttons
            bw = QWidget()
            bw.setObjectName("btnRow"); bw.setStyleSheet("#btnRow { background:#ffffff; }")
            bl = QHBoxLayout(bw); bl.setContentsMargins(8,4,8,4); bl.setSpacing(8)
            b1 = QPushButton("进账簿"); b1.setObjectName("btn_primary")
            b1.setFixedSize(94, 30)
            b2 = QPushButton("编辑"); b2.setObjectName("btn_outline")
            b2.setFixedSize(68, 30)
            b3 = QPushButton("删除"); b3.setObjectName("btn_red")
            b3.setFixedSize(68, 30)
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


