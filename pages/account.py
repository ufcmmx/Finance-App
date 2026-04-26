"""pages/account.py — AccountPage — 会计科目管理"""
from datetime import datetime
from PySide6.QtWidgets import *
from PySide6.QtCore import Qt, QDate, Signal, QTimer
from PySide6.QtGui import QColor, QFont, QBrush, QPalette

from db import get_db, log_action
from utils import lbl, sep, card, fmt_amt, NoScrollSpinBox, NoScrollDoubleSpinBox
from dialogs import AccountEditDialog, ImportExcelDialog
# openpyxl imported lazily inside each export function

class AccountPage(QWidget):
    """科目管理 — 查看/新增/编辑/删除二三级科目"""

    def __init__(self):
        super().__init__()
        self.client_id = None
        L = QVBoxLayout(self); L.setContentsMargins(24,20,24,20); L.setSpacing(14)

        hdr = QHBoxLayout()
        hdr.addWidget(lbl("会计科目管理", bold=True, size=18)); hdr.addStretch()
        b_add = QPushButton("＋ 新增科目"); b_add.setObjectName("btn_primary")
        b_add.clicked.connect(self._add)
        b_imp = QPushButton("从Excel导入历史数据"); b_imp.setObjectName("btn_outline")
        b_imp.clicked.connect(self._import_excel)
        hdr.addWidget(b_imp); hdr.addWidget(b_add); L.addLayout(hdr)

        info = QLabel("  提示：可新增子科目（如 6602.100 管理费用-其他）。系统默认科目不可删除，但可冻结或添加子科目。")
        info.setStyleSheet("background:#fffbe6;color:#ad6800;border-radius:6px;padding:8px 12px;font-size:12px;")
        L.addWidget(info)

        # Filter
        fr = QHBoxLayout(); fr.setSpacing(10)
        self.search_acct = QLineEdit(); self.search_acct.setPlaceholderText("搜索科目编号或名称...")
        self.search_acct.textChanged.connect(self.load)
        self.type_filter = QComboBox()
        self.type_filter.addItems(["全部类型","资产","负债","所有者权益","成本","收入","费用"])
        self.type_filter.currentIndexChanged.connect(self.load)
        fr.addWidget(self.search_acct); fr.addWidget(self.type_filter); fr.addStretch()
        L.addLayout(fr)

        f = card(); vl = QVBoxLayout(f); vl.setContentsMargins(0,0,0,0)
        self.tbl = QTableWidget(); self.tbl.setEditTriggers(QTableWidget.NoEditTriggers)
        self.tbl.setSelectionBehavior(QTableWidget.SelectRows); self.tbl.setShowGrid(False)
        self.tbl.verticalHeader().setVisible(False)
        self.tbl.setColumnCount(5)
        self.tbl.setHorizontalHeaderLabels(["科目编号","科目名称","类型","方向","操作"])
        hh = self.tbl.horizontalHeader()
        hh.setSectionResizeMode(QHeaderView.Interactive)
        hh.setStretchLastSection(True)
        hh.setMinimumSectionSize(55)
        self.tbl.setColumnWidth(0,120); self.tbl.setColumnWidth(1,240)
        self.tbl.setColumnWidth(2,80);  self.tbl.setColumnWidth(3,55)
        self.tbl.setColumnWidth(4,350)
        self.tbl.setHorizontalScrollMode(QTableWidget.ScrollPerPixel)
        vl.addWidget(self.tbl); L.addWidget(f)

    def set_client(self, client_id):
        self.client_id = client_id
        self.load()

    def load(self):
        if not self.client_id: return
        kw = self.search_acct.text().strip()
        tf = self.type_filter.currentText()
        conn = get_db(); c = conn.cursor()
        sql = "SELECT * FROM accounts WHERE client_id=?"
        params = [self.client_id]
        if kw: sql += " AND (code LIKE ? OR name LIKE ?)"; params += [f"%{kw}%",f"%{kw}%"]
        if tf != "全部类型": sql += " AND type=?"; params.append(tf)
        sql += " ORDER BY code"
        c.execute(sql, params)
        rows = c.fetchall()
        
        # Check which accounts have been used
        used_accounts = set()
        c.execute("SELECT DISTINCT account_code FROM voucher_entries")
        for row in c.fetchall():
            used_accounts.add(row['account_code'])
        conn.close()

        self.tbl.setRowCount(len(rows))
        type_colors = {"资产":"#3d6fdb","负债":"#e05252","所有者权益":"#722ed1",
                       "成本":"#fa8c16","收入":"#52c41a","费用":"#eb5757"}
        for i,r in enumerate(rows):
            self.tbl.setRowHeight(i,68)
            level = r["level"] or 1
            indent = "    " * (level-1)
            code_it = QTableWidgetItem(r["code"])
            code_it.setForeground(QColor("#3d6fdb")); code_it.setTextAlignment(Qt.AlignCenter)
            name_it = QTableWidgetItem(indent + r["name"])
            if level == 1: name_it.setFont(QFont("",weight=QFont.Bold))
            
            # If account is frozen, show gray text
            try:
                is_frozen = r['is_frozen']
            except (KeyError, IndexError):
                is_frozen = 0
            if is_frozen:
                code_it.setForeground(QColor("#ccc"))
                name_it.setForeground(QColor("#ccc"))
            
            type_it = QTableWidgetItem(r["type"] or "")
            type_it.setForeground(QColor(type_colors.get(r["type"],"#888")))
            type_it.setTextAlignment(Qt.AlignCenter)
            dir_it = QTableWidgetItem(r["direction"] or "借"); dir_it.setTextAlignment(Qt.AlignCenter)
            for j,it in enumerate([code_it, name_it, type_it, dir_it]):
                self.tbl.setItem(i,j,it)

            bw = QWidget()
            bw.setStyleSheet("QWidget { background:#ffffff; }")
            bl = QHBoxLayout(bw); bl.setContentsMargins(8,10,8,10); bl.setSpacing(12)
            
            # Button style to ensure text is visible on Windows
            outline_style = "color:#3d6fdb; border:1px solid #3d6fdb; background:transparent; border-radius:4px; padding:6px 14px; font-size:12px; font-weight:bold;"
            red_style = "color:#fff; background:#ff4d4f; border:none; border-radius:4px; padding:6px 14px; font-size:12px; font-weight:bold;"
            
            # If frozen, show frozen status
            if is_frozen:
                frozen_lbl = lbl("已冻结", color="#ccc", bold=True)
                bl.addWidget(frozen_lbl)
                bl.addStretch()
                self.tbl.setCellWidget(i,4,bw)
                continue
            
            b_sub = QPushButton("＋ 子科目")
            b_sub.setMinimumWidth(110); b_sub.setMaximumWidth(120); b_sub.setStyleSheet(outline_style)
            b_sub.clicked.connect(lambda _,rr=r: self._add_sub(rr))
            b_ed = QPushButton("✏ 编辑")
            b_ed.setMinimumWidth(90); b_ed.setMaximumWidth(100); b_ed.setStyleSheet(outline_style)
            b_ed.clicked.connect(lambda _,rr=r: self._edit(rr))
            bl.addWidget(b_sub); bl.addWidget(b_ed)
            
            if level > 1:
                is_used = r['code'] in used_accounts
                if is_used:
                    # Account has been used, show freeze button instead of delete
                    b_freeze = QPushButton("❄ 冻结")
                    b_freeze.setMinimumWidth(90); b_freeze.setMaximumWidth(100); b_freeze.setStyleSheet(outline_style)
                    b_freeze.setToolTip("冻结此科目，不再允许使用")
                    b_freeze.clicked.connect(lambda _,rid=r["id"]: self._freeze(rid))
                    bl.addWidget(b_freeze)
                else:
                    # Account not used, show delete button
                    b_del = QPushButton("🗑 删除")
                    b_del.setMinimumWidth(90); b_del.setMaximumWidth(100); b_del.setStyleSheet(red_style)
                    b_del.setToolTip("删除此科目")
                    b_del.clicked.connect(lambda _,rid=r["id"]: self._del(rid))
                    bl.addWidget(b_del)
            bl.addStretch()
            self.tbl.setCellWidget(i,4,bw)

    def _add(self):
        dlg = AccountEditDialog(self, self.client_id)
        if dlg.exec(): self.load()

    def _add_sub(self, parent_acct):
        # Check if parent account has been used (has voucher entries)
        conn = get_db(); c = conn.cursor()
        c.execute("SELECT COUNT(*) FROM voucher_entries WHERE account_code=?", (parent_acct['code'],))
        used_count = c.fetchone()[0]
        conn.close()
        
        if used_count > 0:
            QMessageBox.warning(self, "无法添加下级科目", 
                f"上级科目 【{parent_acct['code']} {parent_acct['name']}】 已有凭证使用，不允许添加下级科目。")
            return
        
        dlg = AccountEditDialog(self, self.client_id, parent_acct=parent_acct)
        if dlg.exec(): self.load()

    def _edit(self, r):
        dlg = AccountEditDialog(self, self.client_id, account=r)
        if dlg.exec(): self.load()

    def _freeze(self, aid):
        """Freeze an account to prevent further use"""
        conn = get_db(); c = conn.cursor()
        c.execute("SELECT code, name, is_default FROM accounts WHERE id=?", (aid,))
        acct = c.fetchone()
        conn.close()
        
        if acct:
            is_def = acct['is_default'] if 'is_default' in acct.keys() else 0
            extra = "\n（系统默认科目冻结后仍保留，不可删除）" if is_def else ""
            # Second confirmation for freezing
            reply = QMessageBox.question(self, "冻结确认", 
                f"确认冻结科目 【{acct['code']} {acct['name']}】 吗？\n\n冻结后将不再允许使用此科目。{extra}",
                QMessageBox.Yes | QMessageBox.No)
            
            if reply == QMessageBox.Yes:
                conn = get_db()
                conn.execute("UPDATE accounts SET is_frozen=1 WHERE id=?", (aid,))
                conn.commit(); conn.close()
                QMessageBox.information(self, "成功", f"科目 【{acct['code']} {acct['name']}】 已冻结。")
                self.load()

    def _del(self, aid):
        conn = get_db(); c = conn.cursor()
        c.execute("SELECT code, name, is_default FROM accounts WHERE id=?", (aid,))
        acct = c.fetchone()
        if not acct: conn.close(); return
        
        acct_code, acct_name = acct['code'], acct['name']
        try:
            is_def = acct['is_default']
        except (KeyError, IndexError):
            is_def = 0
        
        if is_def:
            conn.close()
            QMessageBox.warning(self, "系统默认科目", f"科目【{acct_code} {acct_name}】是系统默认科目，不可删除。如不需要使用，可以点击【冻结科目】。")
            return
        
        # Check if account has been used
        c.execute("SELECT COUNT(*) FROM voucher_entries WHERE account_code=?", (acct_code,))
        used_count = c.fetchone()[0]
        conn.close()
        
        if used_count > 0:
            # Account has been used, offer freeze or delete confirmation
            msg_box = QMessageBox(QMessageBox.Warning, "科目已使用", 
                f"科目 【{acct_code} {acct_name}】 已有凭证使用。\n\n选择操作：")
            btn_freeze = msg_box.addButton("冻结科目", QMessageBox.AcceptRole)
            btn_cancel = msg_box.addButton("取消", QMessageBox.RejectRole)
            msg_box.exec()
            
            if msg_box.clickedButton() == btn_freeze:
                # Freeze the account
                conn = get_db()
                conn.execute("UPDATE accounts SET is_frozen=1 WHERE id=?", (aid,))
                conn.commit(); conn.close()
                QMessageBox.information(self, "成功", f"科目 【{acct_code} {acct_name}】 已冻结，不再允许使用。")
                self.load()
            return
        
        # Account not used, show delete confirmation dialog
        reply = QMessageBox.question(self, "二次确认", 
            f"确认删除科目 【{acct_code} {acct_name}】 吗？\n\n此操作不可撤销。",
            QMessageBox.Yes | QMessageBox.No)
        
        if reply == QMessageBox.Yes:
            conn = get_db()
            conn.execute("DELETE FROM accounts WHERE id=?", (aid,))
            conn.commit(); conn.close()
            QMessageBox.information(self, "成功", f"科目 【{acct_code} {acct_name}】 已删除。")
            self.load()

    def _import_excel(self):
        """Import historical vouchers from Excel."""
        if not self.client_id:
            QMessageBox.information(self,"提示","请先选择客户"); return
        dlg = ImportExcelDialog(self, self.client_id)
        dlg.exec()
        self.load()


