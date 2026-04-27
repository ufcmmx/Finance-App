"""pages/settle.py — SettlePage — 期末结账"""
from datetime import datetime
from PySide6.QtWidgets import *
from PySide6.QtCore import Qt, QDate, Signal, QTimer
from PySide6.QtGui import QColor, QFont, QBrush, QPalette

from db import get_db, log_action
from utils import lbl, sep, card, fmt_amt, NoScrollSpinBox, NoScrollDoubleSpinBox

# openpyxl imported lazily inside each export function

class SettlePage(QWidget):
    """期末结账"""
    carryforward_done = Signal()   # emitted after vouchers created

    def __init__(self):
        super().__init__()
        self.client_id = None; self.client_name = ""; self.period = ""
        outer = QVBoxLayout(self)
        outer.setContentsMargins(0,0,0,0)
        outer.setSpacing(0)
        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        scroll.setFrameShape(QFrame.NoFrame)
        scroll.setHorizontalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
        outer.addWidget(scroll)

        content = QWidget()
        scroll.setWidget(content)
        L = QVBoxLayout(content); L.setContentsMargins(24,20,24,20); L.setSpacing(14)

        # Step indicator
        step_row = QHBoxLayout(); step_row.setSpacing(0)
        s1 = self._step_box("1","期末结转","active"); s2 = self._step_box("2","结账检测","")
        step_row.addWidget(s1); step_row.addWidget(lbl("  ➔  ", color="#bbb"))
        step_row.addWidget(s2); step_row.addStretch()
        L.addLayout(step_row)

        # Period / client row
        pr = QHBoxLayout()
        self.period_combo = QComboBox()
        self.period_combo.setMinimumWidth(130)
        self.period_combo.currentIndexChanged.connect(self._on_period_change)
        self.client_lbl = lbl("请先从客户管理进入账簿", color="#888")
        self.status_lbl = lbl("", bold=True)   # 显示已结账/未结账状态
        self.do_btn = QPushButton("生成结转凭证"); self.do_btn.setObjectName("btn_primary")
        self.do_btn.clicked.connect(self._do_carryforward)
        self.close_btn = QPushButton("结账封账"); self.close_btn.setObjectName("btn_primary")
        self.close_btn.clicked.connect(self._close_period)
        self.reopen_btn = QPushButton("反结账"); self.reopen_btn.setObjectName("btn_red")
        self.reopen_btn.clicked.connect(self._reopen_period)
        pr.addWidget(lbl("结账期间：")); pr.addWidget(self.period_combo)
        pr.addSpacing(10); pr.addWidget(self.client_lbl)
        pr.addSpacing(12); pr.addWidget(self.status_lbl)
        pr.addStretch()
        pr.addWidget(self.do_btn); pr.addWidget(self.close_btn); pr.addWidget(self.reopen_btn)
        L.addLayout(pr)
        L.addWidget(sep())

        # Carry cards
        cards_row = QHBoxLayout(); cards_row.setSpacing(16)
        self.card_income  = self._carry_card("结转本期损益(收入)",  "0.00", "#3d6fdb")
        self.card_expense = self._carry_card("结转本期损益(成本费用)","0.00", "#e05252")
        cards_row.addWidget(self.card_income); cards_row.addWidget(self.card_expense)
        cards_row.addStretch(); L.addLayout(cards_row)

        # ── New: period activity summary table ──
        L.addWidget(lbl("本期收入/费用科目发生额（仅显示已审核凭证）", bold=True, size=13))
        hint = QLabel("  结转的前提：凭证中需要有 6001-6899 收入/费用科目，且凭证状态为【已审核】。")
        hint.setStyleSheet("color:#ad6800;background:#fffbe6;border-radius:5px;padding:6px 10px;font-size:12px;")
        L.addWidget(hint)
        f = card(); vl2 = QVBoxLayout(f); vl2.setContentsMargins(0,0,0,0)
        self.activity_tbl = QTableWidget()
        self.activity_tbl.setEditTriggers(QTableWidget.NoEditTriggers)
        self.activity_tbl.setShowGrid(False); self.activity_tbl.verticalHeader().setVisible(False)
        self.activity_tbl.setVerticalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
        self.activity_tbl.setColumnCount(5)
        self.activity_tbl.setHorizontalHeaderLabels(["科目编号","科目名称","类型","本期借方","本期贷方"])
        self.activity_tbl.horizontalHeader().setSectionResizeMode(1, QHeaderView.Stretch)
        self.activity_tbl.setColumnWidth(0,90); self.activity_tbl.setColumnWidth(2,70)
        self.activity_tbl.setColumnWidth(3,110); self.activity_tbl.setColumnWidth(4,110)
        self.activity_tbl.setMaximumHeight(200)
        vl2.addWidget(self.activity_tbl); L.addWidget(f)
        L.addWidget(sep())

        # Check list
        L.addWidget(lbl("结账检测", bold=True, size=14))
        self.check_list = QTableWidget(); self.check_list.setColumnCount(3)
        self.check_list.setHorizontalHeaderLabels(["序号","检测项目","状态"])
        self.check_list.horizontalHeader().setSectionResizeMode(1, QHeaderView.Stretch)
        self.check_list.verticalHeader().setVisible(False); self.check_list.setShowGrid(False)
        self.check_list.setVerticalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
        L.addWidget(self.check_list); L.addStretch()

    def _step_box(self, num, text, state):
        w = QFrame()
        color = "#3d6fdb" if state=="active" else "#ddd"
        w.setStyleSheet(f"background:{color};border-radius:8px;padding:4px;")
        vl = QHBoxLayout(w); vl.setContentsMargins(16,10,16,10)
        n = lbl(num, bold=True, color="#fff" if state=="active" else "#999", size=18)
        t = lbl(text, color="#fff" if state=="active" else "#999")
        vl.addWidget(n); vl.addWidget(t); return w

    def _carry_card(self, title, amount, color):
        f = QFrame()
        f.setStyleSheet("background:#fff;border-radius:10px;border:1px solid #e4e8f0;")
        f.setFixedWidth(260)
        vl = QVBoxLayout(f); vl.setContentsMargins(20,16,20,16); vl.setSpacing(8)
        # Checkbox + smart label
        row = QHBoxLayout()
        cb = QCheckBox(); cb.setChecked(True)
        smart = lbl("智能生成", color="#52c41a", size=11)
        row.addWidget(cb); row.addStretch(); row.addWidget(smart)
        vl.addLayout(row)
        icon = lbl("⟳", color=color, size=28); icon.setAlignment(Qt.AlignCenter)
        vl.addWidget(icon)
        t = lbl(title, color="#555"); t.setAlignment(Qt.AlignCenter); vl.addWidget(t)
        a = lbl(f"金额：{amount}", bold=True, color=color, size=14); a.setAlignment(Qt.AlignCenter)
        vl.addWidget(a)
        f._amount_lbl = a; f._cb = cb
        return f

    def set_client(self, client_id, client_name, period):
        self.client_id = client_id; self.client_name = client_name; self.period = period
        self.client_lbl.setText(f"【{client_name}】")
        self._refresh_period_options(period)
        self._refresh_period_view()

    def _refresh_period_options(self, selected_period=None):
        self.period_combo.blockSignals(True)
        self.period_combo.clear()
        if not self.client_id:
            self.period_combo.addItem("请先从客户管理进入账簿", "")
            self.period_combo.setEnabled(False)
            self.period_combo.blockSignals(False)
            return

        periods = set()
        now = datetime.now()
        year, month = now.year, now.month
        for _ in range(36):
            periods.add(f"{year}-{month:02d}")
            month -= 1
            if month == 0:
                year -= 1
                month = 12

        conn = get_db(); c = conn.cursor()
        c.execute("""SELECT period FROM vouchers WHERE client_id=?
                     UNION
                     SELECT period FROM periods WHERE client_id=?""",
                  (self.client_id, self.client_id))
        periods.update(row["period"] for row in c.fetchall() if row["period"])
        conn.close()

        target = selected_period or self.period
        if target:
            periods.add(target)

        for p in sorted(periods, reverse=True):
            self.period_combo.addItem(f"{p[:4]}年{p[5:]}期", p)

        idx = self.period_combo.findData(target)
        if idx < 0 and self.period_combo.count():
            idx = 0
        if idx >= 0:
            self.period_combo.setCurrentIndex(idx)
            self.period = self.period_combo.itemData(idx)
        self.period_combo.setEnabled(self.period_combo.count() > 0)
        self.period_combo.blockSignals(False)

    def _on_period_change(self):
        period = self.period_combo.currentData()
        if not period:
            return
        self.period = period
        self._refresh_period_view()

    def _refresh_period_view(self):
        if not self.client_id or not self.period:
            return
        self._refresh_lock_state()
        self._refresh_carry_amounts()
        self._load_activity()
        self._run_checks()

    def _fit_table_height(self, table, extra=8):
        height = table.horizontalHeader().height()
        for row in range(table.rowCount()):
            height += table.rowHeight(row)
        if table.rowCount() == 0:
            height += table.verticalHeader().defaultSectionSize()
        table.setMinimumHeight(height + extra)
        table.setMaximumHeight(height + extra)

    def _is_period_closed(self):
        """Check if current period is closed."""
        if not self.client_id or not self.period: return False
        conn = get_db(); c = conn.cursor()
        c.execute("SELECT is_closed FROM periods WHERE client_id=? AND period=?",
                  (self.client_id, self.period))
        row = c.fetchone(); conn.close()
        return bool(row and row["is_closed"])

    def _refresh_lock_state(self):
        """Update UI buttons and status label based on period close state."""
        closed = self._is_period_closed()
        if closed:
            self.status_lbl.setText("🔒 已结账封账")
            self.status_lbl.setStyleSheet("color:#ff4d4f;font-weight:bold;")
            self.do_btn.setEnabled(False)
            self.close_btn.setVisible(False)
            self.reopen_btn.setVisible(True)
        else:
            self.status_lbl.setText("○ 未结账")
            self.status_lbl.setStyleSheet("color:#888;")
            self.do_btn.setEnabled(True)
            self.close_btn.setVisible(True)
            self.reopen_btn.setVisible(False)

    def _close_period(self):
        """结账封账：检查无待审核凭证后锁定期间。"""
        if not self.client_id: return
        conn = get_db(); c = conn.cursor()
        # 检查待审核凭证
        c.execute("SELECT COUNT(*) FROM vouchers WHERE client_id=? AND period=? AND status='待审核'",
                  (self.client_id, self.period))
        pending = c.fetchone()[0]
        if pending > 0:
            conn.close()
            QMessageBox.warning(self, "无法结账",
                f"本期还有 {pending} 张【待审核】凭证，请先全部审核后再结账封账。")
            return
        # 检查借贷不平凭证（额外保障）
        c.execute("""SELECT v.voucher_no,
            ABS(SUM(e.debit)-SUM(e.credit)) AS diff
            FROM vouchers v JOIN voucher_entries e ON e.voucher_id=v.id
            WHERE v.client_id=? AND v.period=?
            GROUP BY v.id HAVING diff > 0.005""", (self.client_id, self.period))
        unbal = c.fetchall()
        if unbal:
            conn.close()
            nos = "、".join(r["voucher_no"] for r in unbal[:5])
            QMessageBox.warning(self, "无法结账",
                f"以下凭证借贷不平衡，请先修正：{nos}")
            return
        if QMessageBox.question(self, "确认结账封账",
                f"确认对期间【{self.period}】进行结账封账？\n\n封账后该期间凭证将无法新增或修改，如需修改请先反结账。",
                QMessageBox.Yes | QMessageBox.No) != QMessageBox.Yes:
            conn.close(); return
        # 写入或更新 periods 表
        c.execute("INSERT OR REPLACE INTO periods(client_id,period,is_closed,closed_at) VALUES(?,?,1,datetime('now'))",
                  (self.client_id, self.period))
        log_action(conn, self.client_id, "期间结账封账", "period", self.period,
                   f"期间 {self.period} 封账")
        conn.commit(); conn.close()
        self._refresh_lock_state()
        self._run_checks()
        QMessageBox.information(self, "结账成功", f"期间【{self.period}】已结账封账。")

    def _reopen_period(self):
        """反结账：解除期间锁定。"""
        if not self.client_id: return
        conn = get_db(); c = conn.cursor()
        c.execute("""SELECT period FROM periods
                     WHERE client_id=? AND is_closed=1 AND period>=?
                     ORDER BY period""",
                  (self.client_id, self.period))
        closed_periods = [row["period"] for row in c.fetchall()]
        if not closed_periods:
            conn.close()
            QMessageBox.information(self, "提示", f"期间【{self.period}】当前未封账，无需反结账。")
            self._refresh_lock_state()
            self._run_checks()
            return

        msg = QMessageBox(self)
        msg.setWindowTitle("确认反结账")
        msg.setIcon(QMessageBox.Warning)
        msg.setText(f"请选择期间【{self.period}】的反结账方式。")
        if len(closed_periods) > 1:
            msg.setInformativeText(
                f"当前期间之后还有 {len(closed_periods) - 1} 个已结账期间："
                f"{'、'.join(closed_periods[1:4])}"
                f"{' 等' if len(closed_periods) > 4 else ''}\n\n"
                "你可以只反当前期间，也可以连同后续已结账期间一起反结账。")
        else:
            msg.setInformativeText("当前期间之后没有其他已结账期间。")
        current_btn = msg.addButton("只反当前期间", QMessageBox.AcceptRole)
        cascade_btn = None
        if len(closed_periods) > 1:
            cascade_btn = msg.addButton("反当前及后续期间", QMessageBox.DestructiveRole)
        cancel_btn = msg.addButton("取消", QMessageBox.RejectRole)
        msg.setDefaultButton(current_btn)
        msg.exec()

        clicked = msg.clickedButton()
        if clicked == cancel_btn or clicked is None:
            conn.close()
            return

        if clicked == cascade_btn:
            target_periods = closed_periods
            c.execute("""UPDATE periods
                         SET is_closed=0, closed_at=NULL
                         WHERE client_id=? AND is_closed=1 AND period>=?""",
                      (self.client_id, self.period))
            detail = f"期间 {self.period} 反结账，并级联解除后续 {len(target_periods) - 1} 个期间封账"
            success_msg = (
                f"期间【{self.period}】及后续共 {len(target_periods)} 个已结账期间"
                "已解除封账，可重新修改凭证。")
        else:
            target_periods = [self.period]
            c.execute("""UPDATE periods
                         SET is_closed=0, closed_at=NULL
                         WHERE client_id=? AND period=?""",
                      (self.client_id, self.period))
            detail = f"期间 {self.period} 反结账"
            success_msg = f"期间【{self.period}】已解除封账，可重新修改凭证。"

        log_action(conn, self.client_id, "期间反结账", "period", self.period, detail)
        conn.commit(); conn.close()
        self._refresh_lock_state()
        self._run_checks()
        QMessageBox.information(self, "反结账成功", success_msg)

    def _refresh_carry_amounts(self):
        if not self.client_id: return
        conn = get_db(); c = conn.cursor()
        # Only count APPROVED vouchers
        c.execute("""SELECT SUM(e.credit)-SUM(e.debit) FROM voucher_entries e
            JOIN vouchers v ON v.id=e.voucher_id
            WHERE v.client_id=? AND v.period=? AND v.status='已审核'
            AND (e.account_code >= '6001' AND e.account_code < '6400')""",
                  (self.client_id, self.period))
        income = c.fetchone()[0] or 0
        c.execute("""SELECT SUM(e.debit)-SUM(e.credit) FROM voucher_entries e
            JOIN vouchers v ON v.id=e.voucher_id
            WHERE v.client_id=? AND v.period=? AND v.status='已审核'
            AND (e.account_code >= '6400' AND e.account_code < '7000')""",
                  (self.client_id, self.period))
        expense = c.fetchone()[0] or 0
        conn.close()
        self.card_income._amount_lbl.setText(f"金额：{income:,.2f}")
        self.card_expense._amount_lbl.setText(f"金额：{expense:,.2f}")
        self._income_amt = income; self._expense_amt = expense

    def _load_activity(self):
        """Show all 5xxx account activity this period so user can verify before carryforward."""
        if not self.client_id: return
        conn = get_db(); c = conn.cursor()
        # All income+expense accounts with any activity, approved vouchers only
        c.execute("""SELECT e.account_code, e.account_name,
            CASE WHEN (e.account_code >= '6001' AND e.account_code < '6400')
                 THEN '收入' ELSE '费用' END as cat,
            SUM(e.debit) td, SUM(e.credit) tc
            FROM voucher_entries e JOIN vouchers v ON v.id=e.voucher_id
            WHERE v.client_id=? AND v.period=? AND v.status='已审核'
            AND (e.account_code >= '6001' AND e.account_code < '7000')
            GROUP BY e.account_code ORDER BY e.account_code""",
                  (self.client_id, self.period))
        rows = c.fetchall()
        # Also check unapproved counts
        c.execute("""SELECT COUNT(*) FROM vouchers
            WHERE client_id=? AND period=? AND status='待审核'""",
                  (self.client_id, self.period))
        pending = c.fetchone()[0]
        conn.close()

        self.activity_tbl.setRowCount(len(rows))
        if not rows:
            self.activity_tbl.setRowCount(1)
            msg = f"本期已审核凭证中无收入/费用科目（6001-6899）发生额。"
            if pending:
                msg += f"  ⚠ 有 {pending} 张凭证【待审核】，请先审核后再结转。"
            it = QTableWidgetItem(msg)
            it.setForeground(QColor("#ad6800"))
            self.activity_tbl.setItem(0, 0, it)
            self.activity_tbl.setSpan(0, 0, 1, 5)
            self.activity_tbl.setRowHeight(0, 42)
            self._fit_table_height(self.activity_tbl)
            return

        for i, r in enumerate(rows):
            self.activity_tbl.setRowHeight(i, 34)
            cat_color = "#3d6fdb" if r[2] == "收入" else "#e05252"
            for j, val in enumerate([r[0], r[1], r[2], fmt_amt(r[3]), fmt_amt(r[4])]):
                it = QTableWidgetItem(val)
                it.setTextAlignment(Qt.AlignCenter if j != 1 else Qt.AlignLeft | Qt.AlignVCenter)
                if j == 2: it.setForeground(QColor(cat_color))
                if j == 3 and r[3]: it.setForeground(QColor("#3d6fdb"))
                if j == 4 and r[4]: it.setForeground(QColor("#e05252"))
                self.activity_tbl.setItem(i, j, it)

        if pending:
            warn = QTableWidgetItem(f"  ⚠ 另有 {pending} 张凭证【待审核】未计入，请先审核。")
            warn.setForeground(QColor("#ff4d4f"))
            row = len(rows)
            self.activity_tbl.setRowCount(row + 1)
            self.activity_tbl.setRowHeight(row, 38)
            self.activity_tbl.setItem(row, 0, warn)
            self.activity_tbl.setSpan(row, 0, 1, 5)
        self._fit_table_height(self.activity_tbl)

    def _do_carryforward(self):
        if not self.client_id: return
        # 已结账期间不能结转
        if self._is_period_closed():
            QMessageBox.warning(self, "期间已封账", "该期间已结账封账，请先反结账后再操作。"); return
        # 有待审核凭证不能结转
        conn = get_db(); c = conn.cursor()
        c.execute("SELECT COUNT(*) FROM vouchers WHERE client_id=? AND period=? AND status='待审核'",
                  (self.client_id, self.period))
        pending = c.fetchone()[0]
        if pending > 0:
            conn.close()
            QMessageBox.warning(self, "存在待审核凭证",
                f"本期有 {pending} 张凭证尚未审核，请先全部审核通过后再执行结转。"); return

        # 检查是否已存在结转凭证，如有则先删除再重新生成
        c.execute("""SELECT id, voucher_no FROM vouchers
            WHERE client_id=? AND period=? AND note IN ('结转收入','结转费用')""",
                  (self.client_id, self.period))
        old_carry = c.fetchall()
        if old_carry:
            old_nos = "、".join(r["voucher_no"] for r in old_carry)
            reply = QMessageBox.question(self, "已存在结转凭证",
                f"本期已有结转凭证：{old_nos}\n\n"
                "是否删除旧结转凭证并重新生成？",
                QMessageBox.Yes | QMessageBox.No)
            if reply != QMessageBox.Yes:
                conn.close(); return
            for r in old_carry:
                c.execute("DELETE FROM vouchers WHERE id=?", (r["id"],))
            log_action(conn, self.client_id, "删除旧结转凭证", "settle", self.period,
                       f"删除 {len(old_carry)} 张旧结转凭证: {old_nos}")

        def next_vno():
            c.execute("SELECT COUNT(*) FROM vouchers WHERE client_id=? AND period=?",
                      (self.client_id, self.period))
            return f"记-{c.fetchone()[0]+1:03d}"

        # Determine profit account once: use 4103 if exists (6xxx体系), else 3103
        c.execute("SELECT code FROM accounts WHERE client_id=? AND code='4103'", (self.client_id,))
        profit_code = "4103" if c.fetchone() else "3103"
        profit_name = "本年利润"

        generated = []
        # Income carry: debit income accounts, credit 本年利润
        if self.card_income._cb.isChecked() and abs(self._income_amt) > 0.005:
            # Collect all income account balances for this period (5xxx and 6xxx)
            c.execute("""SELECT e.account_code,e.account_name,
                SUM(e.credit)-SUM(e.debit) AS net
                FROM voucher_entries e JOIN vouchers v ON v.id=e.voucher_id
                WHERE v.client_id=? AND v.period=?
                AND (e.account_code >= '6001' AND e.account_code < '6400')
                GROUP BY e.account_code,e.account_name HAVING net>0.005""",
                (self.client_id, self.period))
            income_rows = c.fetchall()
            if income_rows:
                vno = next_vno()
                date = self.period + "-28"
                c.execute("INSERT INTO vouchers(client_id,period,voucher_no,date,status,note) VALUES(?,?,?,?,?,?)",
                          (self.client_id,self.period,vno,date,"已审核","结转收入"))
                vid = c.lastrowid
                for ln,(code,name,net) in enumerate(income_rows,1):
                    c.execute("INSERT INTO voucher_entries(voucher_id,line_no,summary,account_code,account_name,debit,credit) VALUES(?,?,?,?,?,?,?)",
                              (vid,ln,"结转本期损益",code,name,net,0))
                total_income = sum(r[2] for r in income_rows)
                c.execute("INSERT INTO voucher_entries(voucher_id,line_no,summary,account_code,account_name,debit,credit) VALUES(?,?,?,?,?,?,?)",
                          (vid,len(income_rows)+1,"结转本期损益",profit_code,profit_name,0,total_income))
                generated.append(f"{vno}（结转收入 {total_income:,.2f}）")

        # Expense carry: credit expense accounts, debit 本年利润
        if self.card_expense._cb.isChecked() and abs(self._expense_amt) > 0.005:
            c.execute("""SELECT e.account_code,e.account_name,
                SUM(e.debit)-SUM(e.credit) AS net
                FROM voucher_entries e JOIN vouchers v ON v.id=e.voucher_id
                WHERE v.client_id=? AND v.period=?
                AND (e.account_code >= '6400' AND e.account_code < '7000')
                GROUP BY e.account_code,e.account_name HAVING net>0.005""",
                (self.client_id, self.period))
            expense_rows = c.fetchall()
            if expense_rows:
                vno = next_vno()
                date = self.period + "-28"
                c.execute("INSERT INTO vouchers(client_id,period,voucher_no,date,status,note) VALUES(?,?,?,?,?,?)",
                          (self.client_id,self.period,vno,date,"已审核","结转费用"))
                vid = c.lastrowid
                total_expense = sum(r[2] for r in expense_rows)
                c.execute("INSERT INTO voucher_entries(voucher_id,line_no,summary,account_code,account_name,debit,credit) VALUES(?,?,?,?,?,?,?)",
                          (vid,1,"结转本期损益",profit_code,profit_name,total_expense,0))
                for ln,(code,name,net) in enumerate(expense_rows,2):
                    c.execute("INSERT INTO voucher_entries(voucher_id,line_no,summary,account_code,account_name,debit,credit) VALUES(?,?,?,?,?,?,?)",
                              (vid,ln,"结转本期损益",code,name,0,net))
                generated.append(f"{vno}（结转费用 {total_expense:,.2f}）")

        if generated:
            log_action(conn, self.client_id, "期末结转", "settle", self.period,
                       f"生成{len(generated)}张结转凭证: {'; '.join(generated)}")
        conn.commit(); conn.close()

        if generated:
            detail = "\n".join(generated)
            QMessageBox.information(self,"✓ 结转凭证已生成",
                f"已生成 {len(generated)} 张结转凭证（状态：已审核），保存在本期凭证列表中：\n\n{detail}\n\n请到【记账（凭证）→ 查凭证】查看。")
            self.carryforward_done.emit()   # notify VoucherPage to refresh
        else:
            QMessageBox.information(self,"提示","本期无需结转（收入和费用均为零）。")
        self._refresh_carry_amounts(); self._load_activity(); self._run_checks()

    def _run_checks(self):
        if not self.client_id:
            return
        conn = get_db(); c = conn.cursor()

        # 1. 待审核凭证数
        c.execute("SELECT COUNT(*) FROM vouchers WHERE client_id=? AND period=? AND status='待审核'",
                  (self.client_id, self.period))
        pending = c.fetchone()[0]

        # 2. 借贷不平凭证数
        c.execute("""SELECT COUNT(*) FROM (
            SELECT v.id FROM vouchers v JOIN voucher_entries e ON e.voucher_id=v.id
            WHERE v.client_id=? AND v.period=?
            GROUP BY v.id HAVING ABS(SUM(e.debit)-SUM(e.credit)) > 0.005
        )""", (self.client_id, self.period))
        unbalanced = c.fetchone()[0]

        # 3. 是否已结账封账
        c.execute("SELECT is_closed FROM periods WHERE client_id=? AND period=?",
                  (self.client_id, self.period))
        row = c.fetchone()
        is_closed = bool(row and row["is_closed"])

        # 4. 结转凭证是否存在
        c.execute("SELECT COUNT(*) FROM vouchers WHERE client_id=? AND period=? AND note IN ('结转收入','结转费用')",
                  (self.client_id, self.period))
        carried = c.fetchone()[0]
        conn.close()

        if carried > 0:
            carry_status = "已完成"
        elif abs(getattr(self, "_income_amt", 0)) <= 0.005 and abs(getattr(self, "_expense_amt", 0)) <= 0.005:
            carry_status = "无需结转"
        else:
            carry_status = "未结转"

        checks = [
            ("01", "待审核凭证",   "通过" if pending == 0    else f"风险：{pending}张待审核"),
            ("02", "借贷平衡",     "通过" if unbalanced == 0  else f"风险：{unbalanced}张不平"),
            ("03", "期末结转",     carry_status),
            ("04", "期间封账",     "已封账" if is_closed      else "未封账"),
        ]

        self.check_list.setRowCount(len(checks))
        for i, (no, name, status) in enumerate(checks):
            self.check_list.setRowHeight(i, 40)
            for j, v in enumerate([no, name]):
                it = QTableWidgetItem(v); it.setTextAlignment(Qt.AlignCenter)
                self.check_list.setItem(i, j, it)
            is_ok = "风险" not in status
            icon = "✓" if is_ok else "✗"
            color = "#52c41a" if is_ok else "#ff4d4f"
            if status in ("未结转", "未封账", "无需结转"):
                color = "#fa8c16"; icon = "○"
            s_w = QLabel(f"  {icon}  {status}  ")
            s_w.setStyleSheet(f"color:{color};font-weight:bold;")
            self.check_list.setCellWidget(i, 2, s_w)
        self._fit_table_height(self.check_list)


