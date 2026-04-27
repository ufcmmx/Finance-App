"""dialogs/client_dialogs.py — 客户相关对话框"""
from datetime import datetime
from PySide6.QtWidgets import *
from PySide6.QtCore import Qt, QDate, QTimer
from PySide6.QtGui import QColor, QFont, QPalette

from db import get_db, seed_client_accounts, log_action, VOUCHER_TEMPLATES
from utils import (lbl, sep, card, fmt_amt, cn_amount,
                   NoScrollSpinBox, NoScrollDoubleSpinBox,
                   infer_account_type_direction)

# openpyxl imported lazily inside each export function

class ClientDialog(QDialog):
    def __init__(self, parent=None, client=None):
        super().__init__(parent)
        self.client = client
        self.setWindowTitle("新建客户" if not client else "编辑客户")
        self.setMinimumWidth(400)
        self._build()
        if client: self._load()

    def _build(self):
        L = QVBoxLayout(self); L.setContentsMargins(24,24,24,24); L.setSpacing(14)
        L.addWidget(lbl("客户信息", bold=True, size=15))
        F = QFormLayout(); F.setSpacing(10); F.setLabelAlignment(Qt.AlignRight)
        self.name = QLineEdit(); self.name.setPlaceholderText("公司全称（必填）")
        self.code = QLineEdit(); self.code.setPlaceholderText("如 ZY")
        self.typ = QComboBox(); self.typ.addItems(["小规模纳税人","一般纳税人","其他"])
        self.taxid = QLineEdit(); self.taxid.setPlaceholderText("统一社会信用代码")
        self.contact = QLineEdit(); self.phone = QLineEdit()
        F.addRow("公司名称 *", self.name); F.addRow("助记码", self.code)
        F.addRow("客户类型", self.typ);   F.addRow("税号", self.taxid)
        F.addRow("联系人", self.contact); F.addRow("电话", self.phone)
        L.addLayout(F)
        row = QHBoxLayout(); row.addStretch()
        b_cancel = QPushButton("取消"); b_cancel.setObjectName("btn_gray")
        b_save = QPushButton("保 存"); b_save.setObjectName("btn_primary")
        b_cancel.clicked.connect(self.reject); b_save.clicked.connect(self._save)
        row.addWidget(b_cancel); row.addWidget(b_save); L.addLayout(row)

    def _load(self):
        c = self.client
        self.name.setText(c["name"] or ""); self.code.setText(c["short_code"] or "")
        self.taxid.setText(c["tax_id"] or ""); self.contact.setText(c["contact"] or "")
        self.phone.setText(c["phone"] or "")
        idx = self.typ.findText(c["client_type"] or ""); self.typ.setCurrentIndex(max(0,idx))

    def _save(self):
        n = self.name.text().strip()
        if not n: QMessageBox.warning(self,"提示","公司名称不能为空"); return
        conn = get_db(); c = conn.cursor()
        d = (n, self.code.text().strip(), self.typ.currentText(),
             self.taxid.text().strip(), self.contact.text().strip(), self.phone.text().strip())
        if self.client:
            c.execute("UPDATE clients SET name=?,short_code=?,client_type=?,tax_id=?,contact=?,phone=? WHERE id=?",
                      d+(self.client["id"],))
        else:
            c.execute("INSERT INTO clients(name,short_code,client_type,tax_id,contact,phone) VALUES(?,?,?,?,?,?)",d)
            cid = c.lastrowid
            seed_client_accounts(cid, conn)   # reuse same connection — no lock
        conn.commit(); conn.close(); self.accept()


class AccountInitDialog(QDialog):
    """科目期初设置"""
    def __init__(self, parent, client_id, period):
        super().__init__(parent)
        self.client_id = client_id; self.period = period
        self.setWindowTitle("科目期初余额"); self.setMinimumSize(680, 500)
        self._build()

    def _build(self):
        L = QVBoxLayout(self); L.setContentsMargins(16,16,16,16); L.setSpacing(10)
        L.addWidget(lbl("科目期初余额设置", bold=True, size=14))
        L.addWidget(lbl(f"期间：{self.period}  （仅显示末级科目）", color="#888"))
        self.table = QTableWidget()
        self.table.setColumnCount(4)
        self.table.setHorizontalHeaderLabels(["科目编号","科目名称","期初借方","期初贷方"])
        hh = self.table.horizontalHeader()
        hh.setSectionResizeMode(QHeaderView.Interactive)
        hh.setSectionResizeMode(1, QHeaderView.Stretch)
        hh.setMinimumSectionSize(80)
        self.table.setColumnWidth(0, 110)
        self.table.setColumnWidth(2, 160)
        self.table.setColumnWidth(3, 160)
        self.table.verticalHeader().setVisible(False)
        self._load()
        L.addWidget(self.table)
        row = QHBoxLayout(); row.addStretch()
        bs = QPushButton("保存期初"); bs.setObjectName("btn_primary")
        bc = QPushButton("关闭"); bc.setObjectName("btn_gray")
        bs.clicked.connect(self._save); bc.clicked.connect(self.accept)
        row.addWidget(bc); row.addWidget(bs); L.addLayout(row)

    def _load(self):
        conn = get_db(); c = conn.cursor()
        c.execute("SELECT id,code,name,opening_debit,opening_credit FROM accounts WHERE client_id=? ORDER BY code",
                  (self.client_id,))
        rows = c.fetchall(); conn.close()
        self.table.setRowCount(len(rows))
        self._ids = []
        for i,r in enumerate(rows):
            self.table.setRowHeight(i, 40)          # 行高足够容纳输入框
            self._ids.append(r["id"])
            code_it = QTableWidgetItem(r["code"])
            code_it.setForeground(QColor("#3d6fdb"))
            name_it = QTableWidgetItem(r["name"])
            self.table.setItem(i,0,code_it)
            self.table.setItem(i,1,name_it)
            # Spinbox with explicit minimum size so numbers are readable
            def make_spin(val):
                sp = NoScrollDoubleSpinBox()
                sp.setRange(0, 9999999999)
                sp.setDecimals(2)
                sp.setValue(val or 0)
                sp.setMinimumHeight(32)
                sp.setMinimumWidth(140)
                sp.setAlignment(Qt.AlignRight)
                sp.setStyleSheet("QDoubleSpinBox{padding:4px 8px;font-size:13px;}")
                return sp
            d_spin  = make_spin(r["opening_debit"])
            cr_spin = make_spin(r["opening_credit"])
            self.table.setCellWidget(i,2,d_spin)
            self.table.setCellWidget(i,3,cr_spin)

    def _save(self):
        conn = get_db(); c = conn.cursor()
        for i,aid in enumerate(self._ids):
            d = self.table.cellWidget(i,2).value()
            cr = self.table.cellWidget(i,3).value()
            c.execute("UPDATE accounts SET opening_debit=?,opening_credit=? WHERE id=?", (d,cr,aid))
        conn.commit(); conn.close()
        QMessageBox.information(self,"成功","期初余额已保存")


