"""dialogs/client_dialogs.py — 客户相关对话框"""
from datetime import datetime
from PySide6.QtWidgets import *
from PySide6.QtCore import Qt, QDate, QTimer, QEvent, Signal
from PySide6.QtGui import QColor, QFont, QPalette, QKeySequence, QGuiApplication

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
        F = QFormLayout(); F.setSpacing(10); F.setLabelAlignment(Qt.AlignmentFlag.AlignRight)
        self.name = QLineEdit(); self.name.setPlaceholderText("公司全称（必填）")
        self.code = QLineEdit(); self.code.setPlaceholderText("如 ZY")
        self.typ = QComboBox(); self.typ.addItems(["小规模纳税人","一般纳税人","其他"])
        self.std = QComboBox(); self.std.addItems(["企业会计准则","小企业会计制度"])
        self.start_period = QComboBox()
        now = datetime.now()
        cur_p = f"{now.year}-{now.month:02d}"
        for y in range(now.year + 1, 2018 - 1, -1):
            for m in range(12, 0, -1):
                p = f"{y}-{m:02d}"
                lbl_text = f"{y}年{m:02d}期"
                if p == cur_p:
                    lbl_text += "（当前月）"
                self.start_period.addItem(lbl_text, p)
        for i in range(self.start_period.count()):
            if self.start_period.itemData(i) == cur_p:
                self.start_period.setCurrentIndex(i)
                break
        self.taxid = QLineEdit(); self.taxid.setPlaceholderText("统一社会信用代码")
        self.contact = QLineEdit(); self.phone = QLineEdit()
        F.addRow("公司名称 *", self.name); F.addRow("助记码", self.code)
        F.addRow("客户类型", self.typ);   F.addRow("会计制度", self.std)
        F.addRow("启用年月", self.start_period)
        self.sp_hint = QLabel(""); self.sp_hint.setWordWrap(True)
        self.sp_hint.setStyleSheet("color:#fa8c16;font-size:12px;")
        F.addRow("", self.sp_hint)
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
        try:
            if "start_period" in c.keys() and c["start_period"]:
                idx_sp = self.start_period.findData(c["start_period"])
                if idx_sp >= 0:
                    self.start_period.setCurrentIndex(idx_sp)
        except Exception:
            pass
        reason = self._start_period_locked_reason()
        if reason:
            self.start_period.setEnabled(False)
            self.sp_hint.setText(f"{reason}，启用年月不可修改。")

    def _start_period_locked_reason(self):
        """账套已有凭证或已结账期间后，启用年月不可修改（与用友/金蝶一致：启用之后不能再改启用期间）。
        可以修改时返回空串；新建客户时恒为空串。"""
        if not self.client:
            return ""
        conn = get_db()
        try:
            cid = self.client["id"]
            n_v = conn.execute("SELECT COUNT(*) FROM vouchers WHERE client_id=?", (cid,)).fetchone()[0]
            n_cl = conn.execute("SELECT COUNT(*) FROM periods WHERE client_id=? AND is_closed=1", (cid,)).fetchone()[0]
        finally:
            conn.close()
        if n_v:
            return f"该账套已录入 {n_v} 张凭证"
        if n_cl:
            return f"该账套已有 {n_cl} 个已结账期间"
        return ""

    def _save(self):
        n = self.name.text().strip()
        if not n: QMessageBox.warning(self,"提示","公司名称不能为空"); return
        conn = get_db(); c = conn.cursor()
        std = self.std.currentText()
        sp = self.start_period.currentData()
        d = (n, self.code.text().strip(), self.typ.currentText(), std,
             self.taxid.text().strip(), self.contact.text().strip(), self.phone.text().strip(), sp)
        if self.client:
            old_sp = self.client["start_period"] if "start_period" in self.client.keys() else None
            cid = self.client["id"]
            if self._start_period_locked_reason():
                # 已锁定：不写启用年月（后端兜底，不依赖界面上控件是否被禁用）
                c.execute("UPDATE clients SET name=?,short_code=?,client_type=?,accounting_std=?,tax_id=?,contact=?,phone=? WHERE id=?",
                          d[:-1]+(cid,))
                sp = old_sp or "未设置"
            else:
                c.execute("UPDATE clients SET name=?,short_code=?,client_type=?,accounting_std=?,tax_id=?,contact=?,phone=?,start_period=? WHERE id=?",
                          d+(cid,))
            if old_sp and old_sp != sp:
                detail = f"客户【{n}】信息更新。启用年月由【{old_sp}】变更为【{sp}】"
            else:
                detail = f"客户【{n}】信息更新。启用年月: {sp}, 会计制度: {std}"
            log_action(conn, cid, "编辑客户", "client", cid, detail)
        else:
            c.execute("INSERT INTO clients(name,short_code,client_type,accounting_std,tax_id,contact,phone,start_period) VALUES(?,?,?,?,?,?,?,?)",d)
            cid = c.lastrowid
            seed_client_accounts(cid, conn, accounting_std=std)
            log_action(conn, cid, "新建客户", "client", cid,
                       f"新建客户【{n}】，启用年月: {sp}, 会计制度: {std}")
        conn.commit(); conn.close(); self.accept()


_FULL_TO_HALF = str.maketrans({
    '０': '0', '１': '1', '２': '2', '３': '3', '４': '4',
    '５': '5', '６': '6', '７': '7', '８': '8', '９': '9',
    '．': '.', '。': '.', '·': '.', '・': '.', '﹒': '.',
    '，': ',', '、': ',',
})


class AccountOpeningSpinBox(QLineEdit):
    """专用于科目期初余额录入的金额输入框（继承 QLineEdit 避免嵌套 QDoubleSpinBox 样式塌陷）：
    1. 彻底解决文本不可见：直接单层控件，杜绝 Windows 下嵌套内层 QLineEdit padding 叠加导致高度被裁切（0.00 看不见）
    2. 彻底解决 10 倍扩散与覆盖问题：点击或聚焦时全选现有文本；失焦重新聚焦或选区折叠时光标任意位置键入，自动全量覆盖旧值
    3. 中文标点与全角清洗：支持中文句号(。)、中圆点(·)、全角点(．)及千分位逗号自动转换并识别为小数点
    4. 回车/Tab 跳转：按回车键自动完成当前格式化并跳转到下一个可录入科目，非末级科目设为 NoFocus 自动跳过
    5. 失焦自动补齐：失焦时规范格式化为两位小数（例如 22 -> 22.00，空文本 -> 0.00）
    6. 粘贴安全过滤：非法内容（负数、文本、nan、inf、双点等）自动拒绝并保留原值；支持带货币符号或千分位的合法金额
    """
    valueChanged = Signal(float)

    def __init__(self, editable=True, parent=None):
        self._value = 0.0
        self._fresh_input = True
        super().__init__(parent)
        self.setAttribute(Qt.WidgetAttribute.WA_InputMethodEnabled, False)
        self.setAlignment(Qt.AlignmentFlag.AlignRight)
        self.setText("0.00")
        if editable:
            self.setStyleSheet(
                "QLineEdit { background:#ffffff; color:#1e2130; "
                "border:1px solid #d9d9d9; border-radius:4px; "
                "padding:2px 8px; font-family:'Microsoft YaHei','PingFang SC',sans-serif; "
                "font-size:13px; font-weight:500; } "
                "QLineEdit:focus { border:1.5px solid #3d6fdb; }"
            )
            self.textEdited.connect(self._on_text_edited)
        else:
            self.setReadOnly(True)
            self.setFocusPolicy(Qt.FocusPolicy.NoFocus)
            self.setStyleSheet(
                "QLineEdit { background:#f5f7fa; color:#8c93a0; "
                "border:1px solid #e8ecf2; border-radius:4px; "
                "padding:2px 8px; font-family:'Microsoft YaHei','PingFang SC',sans-serif; "
                "font-size:13px; font-weight:bold; }"
            )

    def setRange(self, a, b):
        pass

    def setDecimals(self, d):
        pass

    def value(self) -> float:
        return getattr(self, "_value", 0.0)

    def setText(self, t: str):
        super().setText(t)
        self._fresh_input = True
        try:
            clean = str(t).strip().translate(_FULL_TO_HALF).replace(",", "")
            self._value = round(float(clean), 2)
        except ValueError:
            self._value = 0.0

    def setValue(self, val: float):
        self._value = round(float(val or 0.0), 2)
        self._fresh_input = True
        super().setText(f"{self._value:.2f}")

    def cleanText(self) -> str:
        return self.text().strip()

    def _on_text_edited(self, text: str):
        clean = text.translate(_FULL_TO_HALF).replace(",", "")
        filtered = ""
        has_dot = False
        dot_count = 0
        for ch in clean:
            if ch.isdigit():
                if has_dot:
                    if dot_count < 2:
                        filtered += ch
                        dot_count += 1
                else:
                    filtered += ch
            elif ch == "." and not has_dot:
                filtered += ch
                has_dot = True

        if filtered != text:
            pos = self.cursorPosition()
            super().setText(filtered)
            self.setCursorPosition(min(pos, len(filtered)))
            clean = filtered

        try:
            v = round(float(clean), 2)
        except ValueError:
            v = 0.0
        self._value = v
        self.valueChanged.emit(v)

    def keyPressEvent(self, event):
        if self.isReadOnly():
            event.ignore()
            return

        key = event.key()
        if key in (Qt.Key.Key_Return, Qt.Key.Key_Enter):
            self.focusNextChild()
            event.accept()
            return

        if (key == Qt.Key.Key_V and (event.modifiers() & Qt.KeyboardModifier.ControlModifier)) or \
           (key == Qt.Key.Key_Insert and (event.modifiers() & Qt.KeyboardModifier.ShiftModifier)):
            self.paste()
            event.accept()
            return

        if key == Qt.Key.Key_Delete:
            if self.hasSelectedText() or self._fresh_input:
                self.clear()
                self._value = 0.0
                self._fresh_input = False
                self.valueChanged.emit(0.0)
                event.accept()
                return
            super().keyPressEvent(event)
            return

        if key == Qt.Key.Key_Backspace:
            if self.hasSelectedText() or (self._fresh_input and self.value() != 0.0):
                self.clear()
                self._value = 0.0
                self._fresh_input = False
                self.valueChanged.emit(0.0)
                event.accept()
                return
            super().keyPressEvent(event)
            return

        raw_text = event.text()
        if raw_text:
            mapped = raw_text.translate(_FULL_TO_HALF)
            is_digit_or_dot = any(c.isdigit() or c == '.' for c in mapped)
            if is_digit_or_dot:
                if self._fresh_input or self.text().strip() in ("0.00", "0.0", "0", ""):
                    self._fresh_input = False
                    super().setText("")
                    insert_txt = "".join(c for c in mapped if c.isdigit() or c == '.')
                    if insert_txt:
                        self.insert(insert_txt)
                        self._on_text_edited(self.text())
                    event.accept()
                    return
                elif mapped != raw_text:
                    insert_txt = "".join(c for c in mapped if c.isdigit() or c == '.')
                    if insert_txt:
                        self.insert(insert_txt)
                        self._on_text_edited(self.text())
                    event.accept()
                    return
                else:
                    self._fresh_input = False

        super().keyPressEvent(event)

    def inputMethodEvent(self, event):
        if self.isReadOnly():
            event.ignore()
            return
        c_text = event.commitString()
        if c_text:
            mapped = c_text.translate(_FULL_TO_HALF)
            if self._fresh_input or self.text().strip() in ("0.00", "0.0", "0", ""):
                self._fresh_input = False
                super().setText("")
            filtered = ""
            has_dot = "." in self.text()
            dot_count = len(self.text().split(".")[1]) if has_dot else 0
            for ch in mapped:
                if ch.isdigit():
                    if has_dot:
                        if dot_count < 2:
                            filtered += ch
                            dot_count += 1
                    else:
                        filtered += ch
                elif ch == "." and not has_dot:
                    filtered += ch
                    has_dot = True
            if filtered:
                self.insert(filtered)
                self._on_text_edited(self.text())
            event.accept()
            return
        super().inputMethodEvent(event)

    def paste(self):
        if self.isReadOnly():
            return
        clip = QGuiApplication.clipboard().text()
        if not clip:
            return
        t = clip.strip().translate(_FULL_TO_HALF)
        for pfx in ("￥", "¥", "$"):
            if t.startswith(pfx):
                t = t[len(pfx):].strip()
        t = t.replace(",", "").strip()
        import re
        if not re.match(r'^\d+(\.\d+)?$', t):
            return
        try:
            val = round(float(t), 2)
            self._fresh_input = True
            self.setValue(val)
            self.valueChanged.emit(val)
        except ValueError:
            pass

    def focusInEvent(self, event):
        if event:
            super().focusInEvent(event)
        self._fresh_input = True
        QTimer.singleShot(0, self.selectAll)

    def mousePressEvent(self, event):
        had_focus = self.hasFocus()
        if event:
            super().mousePressEvent(event)
        self._fresh_input = True
        if not had_focus:
            self.selectAll()

    def focusOutEvent(self, event):
        if event:
            super().focusOutEvent(event)
        self._fresh_input = True
        v = self.value()
        super().setText(f"{v:.2f}")


class AuxOpeningDialog(QDialog):
    """辅助核算期初录入对话框：针对已设置辅助核算的末级科目录入各辅助项的期初余额"""
    def __init__(self, parent, client_id, acct_code, acct_name, dim_id, dim_name, existing_data=None):
        super().__init__(parent)
        self.client_id = client_id
        self.acct_code = acct_code
        self.dim_id = dim_id
        self.existing_data = existing_data or {}
        self.setWindowTitle(f"辅助核算期初录入 — {acct_code} {acct_name}（{dim_name}）")
        self.setMinimumSize(560, 420)
        self.result_data = {}
        self.total_debit = 0.0
        self.total_credit = 0.0
        self._build()

    def _build(self):
        L = QVBoxLayout(self); L.setContentsMargins(16, 16, 16, 16); L.setSpacing(10)
        hint = QLabel(f"  请录入【{self.acct_code}】下各辅助项的期初余额，保存后自动汇总回填至该科目主行。")
        hint.setStyleSheet("background:#f6f8ff;color:#444;border-radius:5px;padding:8px 12px;font-size:12px;")
        L.addWidget(hint)

        self.table = QTableWidget()
        self.table.setColumnCount(3)
        self.table.setHorizontalHeaderLabels(["辅助项目", "期初借方", "期初贷方"])
        hh = self.table.horizontalHeader()
        hh.setSectionResizeMode(0, QHeaderView.Stretch)
        self.table.setColumnWidth(1, 150)
        self.table.setColumnWidth(2, 150)
        self.table.verticalHeader().setVisible(False)

        conn = get_db(); c = conn.cursor()
        c.execute("SELECT id, name, code FROM aux_items WHERE client_id=? AND dimension_id=? AND is_active=1 ORDER BY id",
                  (self.client_id, self.dim_id))
        items = c.fetchall(); conn.close()

        self.table.setRowCount(len(items))
        self._rows = []
        for i, it in enumerate(items):
            self.table.setRowHeight(i, 38)
            item_id = it["id"]
            code_str = f" [{it['code']}]" if it["code"] else ""
            name_item = QTableWidgetItem(f"{it['name']}{code_str}")
            name_item.setFlags(Qt.ItemFlag.ItemIsEnabled | Qt.ItemFlag.ItemIsSelectable)
            self.table.setItem(i, 0, name_item)

            d_val, cr_val = self.existing_data.get(item_id, (0.0, 0.0))
            d_spin = AccountOpeningSpinBox(editable=True)
            cr_spin = AccountOpeningSpinBox(editable=True)
            d_spin.setValue(d_val)
            cr_spin.setValue(cr_val)

            d_spin.valueChanged.connect(self._recalc_totals)
            cr_spin.valueChanged.connect(self._recalc_totals)

            self.table.setCellWidget(i, 1, d_spin)
            self.table.setCellWidget(i, 2, cr_spin)
            self._rows.append((item_id, d_spin, cr_spin))

        L.addWidget(self.table)

        self.total_lbl = QLabel()
        self.total_lbl.setStyleSheet("font-size:13px;padding:6px 12px;background:#f0f5ff;color:#1e2130;border-radius:4px;font-weight:bold;")
        L.addWidget(self.total_lbl)
        self._recalc_totals()

        row = QHBoxLayout(); row.addStretch()
        bc = QPushButton("取消"); bc.clicked.connect(self.reject)
        bs = QPushButton("确定"); bs.setObjectName("btn_primary"); bs.clicked.connect(self._on_confirm)
        row.addWidget(bc); row.addWidget(bs)
        L.addLayout(row)

    def _recalc_totals(self):
        td = sum(d_spin.value() for _, d_spin, _ in self._rows)
        tc = sum(cr_spin.value() for _, _, cr_spin in self._rows)
        self.total_debit = round(td, 2)
        self.total_credit = round(tc, 2)
        self.total_lbl.setText(f"借方合计: ¥ {self.total_debit:,.2f}    贷方合计: ¥ {self.total_credit:,.2f}")

    def _on_confirm(self):
        self._recalc_totals()
        self.result_data = {
            item_id: (d_spin.value(), cr_spin.value())
            for item_id, d_spin, cr_spin in self._rows
            if d_spin.value() > 0 or cr_spin.value() > 0
        }
        self.accept()


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
        self.table.setStyleSheet("QTableWidget::item { padding: 1px 2px; }")
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
        # ── 底部借贷平衡状态栏（实时反馈）──
        self.balance_lbl = QLabel("借方合计 ¥ 0.00　　贷方合计 ¥ 0.00　　差额 ¥ 0.00　　✓ 借贷平衡")
        self.balance_lbl.setStyleSheet(
            "font-size:13px;padding:8px 12px;border-radius:5px;"
            "background:#f6ffed;color:#1a5028;font-weight:bold;")
        L.addWidget(self.balance_lbl)
        self._refresh_parents()   # 初始化时再算一次确保状态栏与表格一致
        row = QHBoxLayout(); row.addStretch()
        bs = QPushButton("保存期初"); bs.setObjectName("btn_primary")
        bc = QPushButton("关闭"); bc.setObjectName("btn_gray")
        bs.setAutoDefault(False); bc.setAutoDefault(False)
        bs.clicked.connect(self._save); bc.clicked.connect(self.accept)
        row.addWidget(bc); row.addWidget(bs); L.addLayout(row)

    @staticmethod
    def _rollup(acct_map):
        """从末级科目向上累加期初余额，返回 {code: (debit, credit)} 汇总字典。
        非末级科目（汇总科目）借贷相抵轧差，只在较大的一方保留净额。"""
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

        # 借贷相抵轧差：非末级科目只在较大的一方保留净余额
        for code in all_codes - leaf_codes:
            d, cr = totals[code]
            if d > cr:
                totals[code] = [round(d - cr, 2), 0.0]
            elif cr > d:
                totals[code] = [0.0, round(cr - d, 2)]
            else:
                totals[code] = [0.0, 0.0]

        return totals

    def _load(self):
        conn = get_db(); c = conn.cursor()
        c.execute("SELECT id,code,name,opening_debit,opening_credit FROM accounts WHERE client_id=? ORDER BY code",
                  (self.client_id,))
        rows = c.fetchall()

        # 加载账套科目辅助核算配置
        c.execute("""
            SELECT ac.account_code, ac.dimension_id, ad.name AS dim_name
            FROM account_aux_config ac
            JOIN aux_dimensions ad ON ad.id = ac.dimension_id
            WHERE ac.client_id=?
        """, (self.client_id,))
        self._aux_config = {r["account_code"]: (r["dimension_id"], r["dim_name"]) for r in c.fetchall()}

        # 加载已有辅助核算期初余额
        self._aux_openings = {}
        try:
            c.execute("""
                SELECT account_code, aux_item_id, debit, credit
                FROM account_opening_aux
                WHERE client_id=?
            """, (self.client_id,))
            for r in c.fetchall():
                self._aux_openings.setdefault(r["account_code"], {})[r["aux_item_id"]] = (r["debit"] or 0.0, r["credit"] or 0.0)
        except Exception:
            pass
        conn.close()

        # Build maps
        all_codes = {r['code'] for r in rows}
        leaf_codes = {c for c in all_codes
                      if not any(o != c and o.startswith(c+'.') for o in all_codes)}

        self.table.setRowCount(len(rows))
        self._ids = []; self._leaf_flags = []
        for i, r in enumerate(rows):
            self.table.setRowHeight(i, 42)
            self._ids.append(r["id"])
            code = r['code']
            is_leaf = code in leaf_codes
            self._leaf_flags.append(is_leaf)

            has_aux = code in self._aux_config
            dim_name = self._aux_config[code][1] if has_aux else ""

            code_it = QTableWidgetItem(code)
            display_name = f"{r['name']} [{dim_name}]" if has_aux else r["name"]
            name_it = QTableWidgetItem(display_name)
            if is_leaf:
                code_it.setForeground(QColor("#3d6fdb"))
                if has_aux:
                    name_it.setForeground(QColor("#1890ff"))
                    name_it.setToolTip(f"该科目已启用【{dim_name}】辅助核算，点击借/贷输入框可录入明细")
            else:
                code_it.setForeground(QColor("#aaa"))
                name_it.setForeground(QColor("#aaa"))
                f_bold = self.table.font()
                f_bold.setBold(True)
                name_it.setFont(f_bold)

            self.table.setItem(i, 0, code_it)
            self.table.setItem(i, 1, name_it)

            d_spin  = AccountOpeningSpinBox(editable=is_leaf)
            cr_spin = AccountOpeningSpinBox(editable=is_leaf)
            d_spin.setValue(r["opening_debit"] or 0)
            cr_spin.setValue(r["opening_credit"] or 0)

            if is_leaf:
                if has_aux:
                    # 辅助核算科目：点击时弹出辅助核算期初录入弹窗
                    d_spin.setToolTip("点击录入辅助核算期初明细")
                    cr_spin.setToolTip("点击录入辅助核算期初明细")
                    d_spin.mousePressEvent = lambda e, row_idx=i, ac=code: self._open_aux_dialog(row_idx, ac)
                    cr_spin.mousePressEvent = lambda e, row_idx=i, ac=code: self._open_aux_dialog(row_idx, ac)
                else:
                    d_spin.valueChanged.connect(self._refresh_parents)
                    cr_spin.valueChanged.connect(self._refresh_parents)

            self.table.setCellWidget(i, 2, d_spin)
            self.table.setCellWidget(i, 3, cr_spin)

        self._refresh_parents()

    def _open_aux_dialog(self, row: int, code: str):
        """弹出辅助核算期初明细录入弹窗，确认后回填并汇总"""
        dim_info = self._aux_config.get(code)
        if not dim_info:
            return
        dim_id, dim_name = dim_info
        name_item = self.table.item(row, 1)
        acct_name = name_item.text().split(" [")[0] if name_item else ""
        existing = self._aux_openings.get(code, {})

        dlg = AuxOpeningDialog(self, self.client_id, code, acct_name, dim_id, dim_name, existing)
        if dlg.exec():
            self._aux_openings[code] = dlg.result_data
            dw = self.table.cellWidget(row, 2)
            cw = self.table.cellWidget(row, 3)
            if dw:
                dw.blockSignals(True)
                dw.setValue(dlg.total_debit)
                dw.blockSignals(False)
            if cw:
                cw.blockSignals(True)
                cw.setValue(dlg.total_credit)
                cw.blockSignals(False)
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
        for code, (d, cr) in leaf_vals.items():
            parts = code.split('.')
            for depth in range(1, len(parts)):
                parent = '.'.join(parts[:depth])
                if parent not in parent_totals:
                    parent_totals[parent] = [0.0, 0.0]
                parent_totals[parent][0] += d
                parent_totals[parent][1] += cr

        # 借贷相抵轧差：上级汇总科目只在较大的一方显示净余额
        for p in parent_totals:
            d, cr = parent_totals[p]
            if d > cr:
                parent_totals[p] = [round(d - cr, 2), 0.0]
            elif cr > d:
                parent_totals[p] = [0.0, round(cr - d, 2)]
            else:
                parent_totals[p] = [0.0, 0.0]

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

        # ── 更新底部借贷平衡状态栏 ──
        if hasattr(self, "balance_lbl"):
            total_d = sum(d for d, cr in leaf_vals.values())
            total_c = sum(cr for d, cr in leaf_vals.values())
            diff = total_d - total_c
            if abs(diff) < 0.01:
                self.balance_lbl.setText(
                    f"借方合计 ¥ {total_d:,.2f}　　贷方合计 ¥ {total_c:,.2f}　　"
                    f"差额 ¥ 0.00　　✓ 借贷平衡")
                self.balance_lbl.setStyleSheet(
                    "font-size:13px;padding:8px 12px;border-radius:5px;"
                    "background:#f6ffed;color:#1a5028;font-weight:bold;")
            else:
                self.balance_lbl.setText(
                    f"借方合计 ¥ {total_d:,.2f}　　贷方合计 ¥ {total_c:,.2f}　　"
                    f"差额 ¥ {abs(diff):,.2f}　　✗ 借贷不平")
                self.balance_lbl.setStyleSheet(
                    "font-size:13px;padding:8px 12px;border-radius:5px;"
                    "background:#fff1f0;color:#b03020;font-weight:bold;")

    def _save(self):
        # ── 保存前借贷平衡校验（仅末级科目参与合计）──
        total_d = 0.0
        total_c = 0.0
        for i, is_leaf in enumerate(self._leaf_flags):
            if not is_leaf:
                continue
            dw = self.table.cellWidget(i, 2)
            cw = self.table.cellWidget(i, 3)
            if dw: total_d += dw.value()
            if cw: total_c += cw.value()
        diff = abs(total_d - total_c)
        if diff >= 0.01:
            reply = QMessageBox.warning(
                self, "借贷不平衡",
                f"借方合计 ¥ {total_d:,.2f}\n"
                f"贷方合计 ¥ {total_c:,.2f}\n"
                f"差额 ¥ {diff:,.2f}\n\n"
                "期初余额借贷不平衡可能导致账套数据从一开始就有错。\n"
                "是否仍要保存？",
                QMessageBox.Yes | QMessageBox.No,
                QMessageBox.No)
            if reply != QMessageBox.Yes:
                return

        conn = get_db(); c = conn.cursor()
        # Save all rows (both leaf and parent — parents have already been rolled up in UI)
        for i, aid in enumerate(self._ids):
            dw = self.table.cellWidget(i, 2)
            cw = self.table.cellWidget(i, 3)
            d  = dw.value() if dw else 0
            cr = cw.value() if cw else 0
            c.execute("UPDATE accounts SET opening_debit=?,opening_credit=? WHERE id=?",
                      (d, cr, aid))

        # 保存辅助核算期初明细
        if hasattr(self, "_aux_openings") and self._aux_openings:
            for code, items in self._aux_openings.items():
                dim_info = self._aux_config.get(code)
                if not dim_info:
                    continue
                dim_id = dim_info[0]
                c.execute("DELETE FROM account_opening_aux WHERE client_id=? AND account_code=?",
                          (self.client_id, code))
                for item_id, (d_val, cr_val) in items.items():
                    if d_val > 0 or cr_val > 0:
                        c.execute("""
                            INSERT INTO account_opening_aux(client_id, account_code, dimension_id, aux_item_id, debit, credit)
                            VALUES (?, ?, ?, ?, ?, ?)
                        """, (self.client_id, code, dim_id, item_id, d_val, cr_val))

        if self.period:
            cur = c.execute("SELECT start_period FROM clients WHERE id=?", (self.client_id,)).fetchone()
            cur_sp = cur[0] if cur and cur[0] else ""
            if not cur_sp:
                c.execute("UPDATE clients SET start_period=? WHERE id=?", (self.period, self.client_id))
            elif self.period < cur_sp:
                # 建账期前移：更新并记录审计日志
                c.execute("UPDATE clients SET start_period=? WHERE id=?", (self.period, self.client_id))
                log_action(conn, self.client_id, "调整建账期", "client", self.client_id,
                           f"建账期由【{cur_sp}】变更为【{self.period}】")

        log_action(conn, self.client_id, "保存期初余额", "client", self.client_id,
                   f"期间: {self.period}, 录入/修改科目期初余额")
        conn.commit(); conn.close()
        QMessageBox.information(self, "成功", "期初余额已保存（末级科目已自动汇总到上级）")