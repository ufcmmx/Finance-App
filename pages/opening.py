"""pages/opening.py — OpeningBalancePage — 科目期初余额（独立模块）"""
from datetime import datetime
from PySide6.QtWidgets import *
from PySide6.QtCore import Qt, QDate, Signal, QTimer
from PySide6.QtGui import QColor, QFont, QBrush, QPalette

from db import get_db, log_action
from utils import lbl, sep, card, fmt_amt, NoScrollSpinBox, NoScrollDoubleSpinBox
from dialogs import AccountInitDialog


class OpeningBalancePage(QWidget):
    """科目期初余额 — 独立页面
    允许操作的期间：
      1. 账套建立期间（accounts 表中有数据的最早 period，或无凭证时任意期间）
      2. 每年第一个期间（XX-01）
    """

    def __init__(self):
        super().__init__()
        self.client_id = None
        self.client_name = ""
        self._preview_display = []
        self._build()

    def _build(self):
        L = QVBoxLayout(self)
        L.setContentsMargins(24, 20, 24, 20)
        L.setSpacing(14)

        # ── 标题 ──
        hdr = QHBoxLayout()
        hdr.addWidget(lbl("科目期初余额", bold=True, size=18))
        hdr.addStretch()
        L.addLayout(hdr)

        # ── 说明 ──
        hint = QLabel(
            "  期初余额只允许在以下两种情况下录入：\n"
            "  ① 账套建立的起始期间（首次建账）\n"
            "  ② 每年的第一个期间（1月，用于年初结转）")
        hint.setStyleSheet(
            "background:#f6f8ff;color:#444;border-radius:6px;"
            "padding:10px 14px;font-size:12px;")
        hint.setWordWrap(True)
        L.addWidget(hint)

        # ── 期间选择 ──
        pr = QHBoxLayout(); pr.setSpacing(10)
        pr.addWidget(lbl("选择期间："))
        self.period_combo = QComboBox()
        self.period_combo.setMinimumWidth(160)
        self.period_combo.currentIndexChanged.connect(self._on_period_change)
        pr.addWidget(self.period_combo)
        self.status_lbl = lbl("", color="#888")
        pr.addWidget(self.status_lbl)
        pr.addStretch()
        self.edit_btn = QPushButton("录入/修改期初余额")
        self.edit_btn.setObjectName("btn_primary")
        self.edit_btn.setMinimumWidth(160)
        self.edit_btn.clicked.connect(self._open_dialog)
        pr.addWidget(self.edit_btn)
        L.addLayout(pr)
        L.addWidget(sep())

        # ── 当前期初余额预览表 ──
        L.addWidget(lbl("当前期初余额（末级科目）", bold=True, size=13))
        f = card(); vl = QVBoxLayout(f); vl.setContentsMargins(0, 0, 0, 0)
        f.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Expanding)
        self.tbl = QTableWidget()
        self.tbl.setEditTriggers(QTableWidget.NoEditTriggers)
        self.tbl.setShowGrid(True)
        self.tbl.verticalHeader().setVisible(False)
        self.tbl.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Expanding)
        self.tbl.setColumnCount(4)
        self.tbl.setHorizontalHeaderLabels(["科目编号", "科目名称", "期初借方", "期初贷方"])
        hh = self.tbl.horizontalHeader()
        hh.setSectionResizeMode(QHeaderView.Interactive)
        hh.setSectionResizeMode(1, QHeaderView.Stretch)
        self.tbl.setColumnWidth(0, 130)
        self.tbl.setColumnWidth(2, 130)
        self.tbl.setColumnWidth(3, 130)
        self.tbl.setVerticalScrollMode(QAbstractItemView.ScrollPerPixel)
        vl.addWidget(self.tbl)
        L.addWidget(f, 1)

        self._set_ui_enabled(False)

    def set_client(self, client_id, client_name, period):
        self.client_id = client_id
        self.client_name = client_name
        self._refresh_periods()
        # Try to select the passed period if allowed, else first allowed
        found = False
        for i in range(self.period_combo.count()):
            if self.period_combo.itemData(i) == period:
                self.period_combo.setCurrentIndex(i)
                found = True
                break
        if not found and self.period_combo.count():
            self.period_combo.setCurrentIndex(0)
        self._load_preview()

    def _allowed_periods(self):
        """返回允许录入期初的期间列表。"""
        if not self.client_id:
            return []

        conn = get_db(); c = conn.cursor()

        # 找到最早有凭证的期间（账套起始期间）
        c.execute(
            "SELECT MIN(period) FROM vouchers WHERE client_id=?",
            (self.client_id,))
        row = c.fetchone()
        earliest_voucher = row[0] if row and row[0] else None

        # 找到 accounts 创建时间 — 用 clients 表的 created_at 推断起始期间
        c.execute("SELECT created_at FROM clients WHERE id=?", (self.client_id,))
        row = c.fetchone()
        conn.close()

        allowed = set()
        now = datetime.now()

        # 规则1：每年1月
        for y in range(2018, now.year + 2):
            allowed.add(f"{y}-01")

        # 规则2：账套起始期间（最早凭证所在期间或当前期间）
        if earliest_voucher:
            allowed.add(earliest_voucher)
        else:
            # 无凭证时，允许当前期间（新账套）
            allowed.add(f"{now.year}-{now.month:02d}")

        return sorted(allowed, reverse=True)

    def _refresh_periods(self):
        self.period_combo.blockSignals(True)
        self.period_combo.clear()
        periods = self._allowed_periods()
        for p in periods:
            y, m = p.split('-')
            label = f"{y}年{m}期"
            if m == "01":
                label += "（年初）"
            elif p == periods[-1] if periods else False:
                label += "（建账期）"
            self.period_combo.addItem(label, p)
        self.period_combo.blockSignals(False)
        self._set_ui_enabled(bool(periods))

    def _on_period_change(self):
        self._load_preview()

    def _set_ui_enabled(self, enabled):
        self.edit_btn.setEnabled(enabled)
        self.period_combo.setEnabled(enabled)
        if not enabled:
            self.status_lbl.setText("请先从客户管理进入账簿")
            self.status_lbl.setStyleSheet("color:#aaa;")

    def _load_preview(self):
        """显示当前期间的期初余额（仅末级科目，非零行）。"""
        self.tbl.setRowCount(0)
        if not self.client_id:
            return

        conn = get_db(); c = conn.cursor()
        c.execute(
            "SELECT code, name, opening_debit, opening_credit FROM accounts "
            "WHERE client_id=? ORDER BY code",
            (self.client_id,))
        rows = c.fetchall()
        conn.close()

        all_codes = {r['code'] for r in rows}
        leaf_codes = {r['code'] for r in rows
                      if not any(o != r['code'] and o.startswith(r['code'] + '.')
                                 for o in all_codes)}

        display = [r for r in rows
                   if r['code'] in leaf_codes
                   and ((r['opening_debit'] or 0) != 0 or (r['opening_credit'] or 0) != 0)]

        period = self.period_combo.currentData() or ""
        if display:
            self.status_lbl.setText(
                f"期间 {period}  共 {len(display)} 个末级科目有期初余额")
            self.status_lbl.setStyleSheet("color:#52c41a;")
        else:
            self.status_lbl.setText(f"期间 {period}  尚未录入期初余额")
            self.status_lbl.setStyleSheet("color:#fa8c16;")

        self._preview_display = list(display)
        self._render_preview_table()

    def _render_preview_table(self):
        display = list(getattr(self, "_preview_display", []))
        row_h = 34
        data_rows = len(display)
        extra_rows = 0
        viewport_h = max(0, self.tbl.viewport().height())
        if viewport_h > 0:
            extra_rows = max(0, viewport_h // row_h - data_rows)

        self.tbl.setRowCount(data_rows + extra_rows)
        for i, r in enumerate(display):
            self.tbl.setRowHeight(i, row_h)
            od = r['opening_debit'] or 0
            oc = r['opening_credit'] or 0
            code_it = QTableWidgetItem(r['code'])
            code_it.setForeground(QColor("#3d6fdb"))
            name_it = QTableWidgetItem(r['name'])
            d_it = QTableWidgetItem(fmt_amt(od))
            d_it.setTextAlignment(Qt.AlignRight | Qt.AlignVCenter)
            if od: d_it.setForeground(QColor("#3d6fdb"))
            c_it = QTableWidgetItem(fmt_amt(oc))
            c_it.setTextAlignment(Qt.AlignRight | Qt.AlignVCenter)
            if oc: c_it.setForeground(QColor("#e05252"))
            for j, it in enumerate([code_it, name_it, d_it, c_it]):
                self.tbl.setItem(i, j, it)

        for i in range(data_rows, data_rows + extra_rows):
            self.tbl.setRowHeight(i, row_h)
            for j in range(self.tbl.columnCount()):
                it = QTableWidgetItem("")
                it.setFlags(Qt.ItemIsEnabled)
                self.tbl.setItem(i, j, it)

    def resizeEvent(self, event):
        super().resizeEvent(event)
        if getattr(self, "_preview_display", None) is not None:
            QTimer.singleShot(0, self._render_preview_table)

    def _open_dialog(self):
        if not self.client_id:
            QMessageBox.information(self, "提示", "请先从客户管理进入账簿")
            return
        period = self.period_combo.currentData()
        if not period:
            return
        # Check if this period is allowed
        allowed = self._allowed_periods()
        if period not in allowed:
            QMessageBox.warning(
                self, "不允许操作",
                f"期间【{period}】不允许录入期初余额。\n\n"
                "只有以下期间允许录入：\n"
                "· 账套建立的起始期间\n"
                "· 每年1月（年初期间）")
            return
        d = AccountInitDialog(self, self.client_id, period)
        d.exec()
        self._load_preview()
