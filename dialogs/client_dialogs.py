"""dialogs/client_dialogs.py — 客户相关对话框"""
from datetime import datetime
from PySide6.QtWidgets import *
from PySide6.QtCore import Qt, QDate, QTimer
from PySide6.QtGui import QColor, QFont, QPalette

from db import get_db, seed_client_accounts, log_action, VOUCHER_TEMPLATES, STANDARD_ACCOUNTS_SMALL
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
        self.std = QComboBox(); self.std.addItems(["企业会计准则","小企业会计制度"])
        self.taxid = QLineEdit(); self.taxid.setPlaceholderText("统一社会信用代码")
        self.contact = QLineEdit(); self.phone = QLineEdit()
        F.addRow("公司名称 *", self.name); F.addRow("助记码", self.code)
        F.addRow("客户类型", self.typ);   F.addRow("会计制度", self.std)
        F.addRow("税号", self.taxid)
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
        try:
            idx2 = self.std.findText(c["accounting_std"] or "企业会计准则")
            self.std.setCurrentIndex(max(0, idx2))
        except Exception:
            pass

    def _save(self):
        n = self.name.text().strip()
        if not n: QMessageBox.warning(self,"提示","公司名称不能为空"); return
        conn = get_db(); c = conn.cursor()
        std = self.std.currentText()
        d = (n, self.code.text().strip(), self.typ.currentText(), std,
             self.taxid.text().strip(), self.contact.text().strip(), self.phone.text().strip())
        if self.client:
            c.execute("UPDATE clients SET name=?,short_code=?,client_type=?,accounting_std=?,tax_id=?,contact=?,phone=? WHERE id=?",
                      d+(self.client["id"],))
        else:
            c.execute("INSERT INTO clients(name,short_code,client_type,accounting_std,tax_id,contact,phone) VALUES(?,?,?,?,?,?,?)",d)
            cid = c.lastrowid
            seed_client_accounts(cid, conn, accounting_std=std)
        conn.commit(); conn.close(); self.accept()


class AccountInitDialog(QDialog):
    """科目期初设置 — 只允许编辑末级科目，保存后自动汇总到上级科目"""
    def __init__(self, parent, client_id, period):
        super().__init__(parent)
        self.client_id = client_id; self.period = period
        self.setWindowTitle("科目期初余额"); self.setMinimumSize(680, 520)
        self._build()

    def _build(self):
        L = QVBoxLayout(self); L.setContentsMargins(16,16,16,16); L.setSpacing(10)
        L.addWidget(lbl("科目期初余额设置", bold=True, size=14))
        hint = QLabel("  只需填写末级科目期初余额，保存时自动汇总到上级科目。上级科目显示为灰色，不可直接编辑。")
        hint.setStyleSheet("background:#f6f8ff;color:#444;border-radius:5px;padding:6px 10px;font-size:12px;")
        hint.setWordWrap(True); L.addWidget(hint)
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

    @staticmethod
    def _rollup(acct_map):
        """从末级科目向上累加期初余额，返回 {code: (debit, credit)} 汇总字典。"""
        totals = {code: [a['opening_debit'] or 0, a['opening_credit'] or 0]
                  for code, a in acct_map.items()}
        all_codes = set(acct_map.keys())
        # Identify leaf codes (no child exists)
        leaf_codes = {c for c in all_codes
                      if not any(o != c and (o.startswith(c+'.')) for o in all_codes)}
        # Zero out parent codes so we start fresh from leaves
        for code in all_codes - leaf_codes:
            totals[code] = [0.0, 0.0]
        # Bubble up: for every leaf, add to each ancestor
        for code in leaf_codes:
            parts = code.split('.')
            d, cr = totals[code]
            for depth in range(1, len(parts)):
                parent = '.'.join(parts[:depth])
                if parent in totals:
                    totals[parent][0] += d
                    totals[parent][1] += cr
        return totals

    def _load(self):
        conn = get_db(); c = conn.cursor()
        c.execute("SELECT id,code,name,opening_debit,opening_credit FROM accounts WHERE client_id=? ORDER BY code",
                  (self.client_id,))
        rows = c.fetchall(); conn.close()

        # Build maps
        all_codes = {r['code'] for r in rows}
        leaf_codes = {c for c in all_codes
                      if not any(o != c and o.startswith(c+'.') for o in all_codes)}

        self.table.setRowCount(len(rows))
        self._ids = []; self._leaf_flags = []
        for i, r in enumerate(rows):
            self.table.setRowHeight(i, 40)
            self._ids.append(r["id"])
            is_leaf = r['code'] in leaf_codes
            self._leaf_flags.append(is_leaf)

            code_it = QTableWidgetItem(r["code"])
            name_it = QTableWidgetItem(r["name"])
            if is_leaf:
                code_it.setForeground(QColor("#3d6fdb"))
            else:
                code_it.setForeground(QColor("#aaa"))
                name_it.setForeground(QColor("#aaa"))
                name_it.setFont(QFont("", weight=QFont.Bold))

            self.table.setItem(i, 0, code_it)
            self.table.setItem(i, 1, name_it)

            def make_spin(val, editable):
                sp = NoScrollDoubleSpinBox()
                sp.setRange(0, 9999999999)
                sp.setDecimals(2)
                sp.setValue(val or 0)
                sp.setMinimumHeight(32)
                sp.setMinimumWidth(140)
                sp.setAlignment(Qt.AlignRight)
                if editable:
                    sp.setStyleSheet("QDoubleSpinBox{padding:4px 8px;font-size:13px;}")
                else:
                    sp.setReadOnly(True)
                    sp.setStyleSheet(
                        "QDoubleSpinBox{padding:4px 8px;font-size:13px;"
                        "background:#f5f7fa;color:#aaa;border:1px solid #e8ecf2;}")
                return sp

            d_spin  = make_spin(r["opening_debit"],  is_leaf)
            cr_spin = make_spin(r["opening_credit"], is_leaf)

            # Connect leaf spinboxes to auto-refresh parent totals
            if is_leaf:
                d_spin.valueChanged.connect(self._refresh_parents)
                cr_spin.valueChanged.connect(self._refresh_parents)

            self.table.setCellWidget(i, 2, d_spin)
            self.table.setCellWidget(i, 3, cr_spin)

        self._refresh_parents()

    def _refresh_parents(self):
        """Recompute and display rolled-up totals for all parent rows."""
        # Collect current leaf values
        leaf_vals = {}
        for i, (aid, is_leaf) in enumerate(zip(self._ids, self._leaf_flags)):
            if not is_leaf: continue
            code_item = self.table.item(i, 0)
            if not code_item: continue
            dw = self.table.cellWidget(i, 2)
            cw = self.table.cellWidget(i, 3)
            leaf_vals[code_item.text()] = (
                dw.value() if dw else 0,
                cw.value() if cw else 0
            )

        # Bubble up to parents
        parent_totals = {}
        all_leaf_codes = set(leaf_vals.keys())
        for code, (d, cr) in leaf_vals.items():
            parts = code.split('.')
            for depth in range(1, len(parts)):
                parent = '.'.join(parts[:depth])
                if parent not in parent_totals:
                    parent_totals[parent] = [0.0, 0.0]
                parent_totals[parent][0] += d
                parent_totals[parent][1] += cr

        # Update parent rows in table
        for i, (aid, is_leaf) in enumerate(zip(self._ids, self._leaf_flags)):
            if is_leaf: continue
            code_item = self.table.item(i, 0)
            if not code_item: continue
            code = code_item.text()
            d_total, cr_total = parent_totals.get(code, (0.0, 0.0))
            dw = self.table.cellWidget(i, 2)
            cw = self.table.cellWidget(i, 3)
            if dw: dw.blockSignals(True); dw.setValue(d_total); dw.blockSignals(False)
            if cw: cw.blockSignals(True); cw.setValue(cr_total); cw.blockSignals(False)

    def _save(self):
        conn = get_db(); c = conn.cursor()
        # Save all rows (both leaf and parent — parents have already been rolled up in UI)
        for i, aid in enumerate(self._ids):
            dw = self.table.cellWidget(i, 2)
            cw = self.table.cellWidget(i, 3)
            d  = dw.value() if dw else 0
            cr = cw.value() if cw else 0
            c.execute("UPDATE accounts SET opening_debit=?,opening_credit=? WHERE id=?",
                      (d, cr, aid))
        conn.commit(); conn.close()
        QMessageBox.information(self, "成功", "期初余额已保存（末级科目已自动汇总到上级）")