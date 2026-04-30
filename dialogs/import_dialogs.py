"""dialogs/import_dialogs.py — 账套导入向导"""
from datetime import datetime
from PySide6.QtWidgets import *
from PySide6.QtCore import Qt, QDate, QTimer
from PySide6.QtGui import QColor, QFont, QPalette

from db import get_db, seed_client_accounts, log_action, VOUCHER_TEMPLATES
from utils import (lbl, sep, card, fmt_amt, cn_amount,
                   process_aux_from_code,
                   NoScrollSpinBox, NoScrollDoubleSpinBox,
                   infer_account_type_direction)

# openpyxl imported lazily inside each export function

class ImportAccountSetDialog(QDialog):
    """账套导入向导 — 四步骤，支持科目余额表/凭证/银行日记账"""

    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("导入账套")
        self.setMinimumSize(860, 640)
        self.resize(960, 720)
        self._files = {}        # key -> path
        self._client_id = None
        self._build()

    def showEvent(self, event):
        super().showEvent(event)
        self._fit_to_screen()

    def _fit_to_screen(self):
        # macOS 下模态对话框有时会按默认几何打开，导致底部按钮栏超出可视区。
        screen = self.windowHandle().screen() if self.windowHandle() else QApplication.primaryScreen()
        if not screen:
            return
        avail = screen.availableGeometry()
        max_width = max(520, avail.width() - 80)
        max_height = max(420, avail.height() - 80)
        width = min(max(self.width(), 860), max_width)
        height = min(max(self.height(), 640), max_height)
        self.resize(width, height)
        x = avail.x() + (avail.width() - width) // 2
        y = avail.y() + (avail.height() - height) // 2
        self.move(max(avail.x(), x), max(avail.y(), y))

    # ─────────────────────────── UI 骨架 ───────────────────────────
    def _build(self):
        L = QVBoxLayout(self); L.setContentsMargins(0,0,0,0); L.setSpacing(0)

        # 顶部标题
        title_bar = QWidget()
        title_bar.setStyleSheet("background:#1c2340;")
        tl = QHBoxLayout(title_bar); tl.setContentsMargins(24,16,24,16)
        tl.addWidget(lbl("📥  账套导入向导", bold=True, color="#fff", size=16))
        tl.addSpacing(12)
        tl.addWidget(lbl("从用友 / 金蝶导出文件一键建立账套", color="#8b93ae", size=12))
        tl.addStretch(); L.addWidget(title_bar)

        # 左侧步骤导航 + 右侧内容
        body = QWidget(); bl = QHBoxLayout(body)
        bl.setContentsMargins(0,0,0,0); bl.setSpacing(0)

        nav = QWidget(); nav.setFixedWidth(176)
        nav.setStyleSheet("background:#f7f9fc;border-right:1px solid #e4e8f0;")
        nl = QVBoxLayout(nav); nl.setContentsMargins(0,14,0,14); nl.setSpacing(2)
        self._step_btns = []
        for num, text in [("1","客户信息"),("2","选择文件"),("3","预览确认"),("4","导入完成")]:
            sb = QPushButton(f"  {num}   {text}")
            sb.setStyleSheet("""QPushButton{background:transparent;color:#888;border:none;
                text-align:left;padding:11px 14px;border-left:3px solid transparent;}
                QPushButton[active=true]{background:#e6f0ff;color:#3d6fdb;
                border-left:3px solid #3d6fdb;font-weight:bold;}""")
            sb.setProperty("active","false"); nl.addWidget(sb)
            self._step_btns.append(sb)
        nl.addStretch(); bl.addWidget(nav)

        self.right_scroll = QScrollArea()
        self.right_scroll.setWidgetResizable(True)
        self.right_scroll.setFrameShape(QFrame.NoFrame)
        self.right_scroll.setHorizontalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
        self.right = QStackedWidget()
        self.right_scroll.setWidget(self.right)
        bl.addWidget(self.right_scroll)
        # stretch=1 让 body 撑满剩余空间，foot 保持固定高度不被挤走
        L.addWidget(body, 1)

        # 底部按钮栏 — 固定高度，确保始终可见
        foot = QWidget()
        foot.setObjectName("import_foot")
        foot.setMinimumHeight(76)
        foot.setMaximumHeight(76)
        fl = QHBoxLayout(foot)
        fl.setContentsMargins(20, 14, 20, 14)
        fl.setSpacing(10)
        fl.addStretch()
        self.btn_back = QPushButton("← 上一步"); self.btn_back.setObjectName("btn_gray")
        self.btn_back.setMinimumHeight(36)
        self.btn_back.setMinimumWidth(120)
        self.btn_back.clicked.connect(self._prev_step)
        self.btn_next = QPushButton("下一步 →"); self.btn_next.setObjectName("btn_primary")
        self.btn_next.setMinimumHeight(36)
        self.btn_next.setMinimumWidth(120)
        self.btn_next.clicked.connect(self._next_step)
        self.btn_close = QPushButton("关 闭"); self.btn_close.setObjectName("btn_gray")
        self.btn_close.setMinimumHeight(36)
        self.btn_close.setMinimumWidth(96)
        self.btn_close.clicked.connect(self.accept); self.btn_close.setVisible(False)
        fl.addWidget(self.btn_back); fl.addWidget(self.btn_next); fl.addWidget(self.btn_close)
        L.addWidget(foot, 0)   # stretch=0，高度由 setFixedHeight 决定

        self._build_step1(); self._build_step2(); self._build_step3(); self._build_step4()
        self._goto_step(0)

    # ─────────────────── Step 1 · 客户信息 ───────────────────
    def _build_step1(self):
        w = QWidget(); L = QVBoxLayout(w); L.setContentsMargins(32,24,32,24); L.setSpacing(14)
        L.addWidget(lbl("填写客户（账套）基本信息", bold=True, size=14))
        hint = QLabel("  新建账套，或选择已有客户追加导入。同名账套自动追加，不会重复创建。")
        hint.setStyleSheet("background:#f6f8ff;color:#444;border-radius:6px;"
                           "padding:8px 12px;font-size:12px;")
        hint.setWordWrap(True); L.addWidget(hint)

        mode_row = QHBoxLayout()
        self.mode_new   = QPushButton("✦ 新建账套")
        self.mode_exist = QPushButton("◈ 追加到已有账套")
        for b in [self.mode_new, self.mode_exist]:
            b.setCheckable(True)
            b.setStyleSheet("""QPushButton{background:#f5f7fa;color:#555;
                border:1px solid #d9d9d9;border-radius:6px;padding:8px 18px;}
                QPushButton:checked{background:#3d6fdb;color:#fff;border-color:#3d6fdb;}""")
        self.mode_new.setChecked(True)
        self.mode_new.clicked.connect(lambda: self._toggle_mode(True))
        self.mode_exist.clicked.connect(lambda: self._toggle_mode(False))
        mode_row.addWidget(self.mode_new); mode_row.addWidget(self.mode_exist)
        mode_row.addStretch(); L.addLayout(mode_row)

        # 新建表单
        self.form_new = QWidget(); fn = QFormLayout(self.form_new)
        fn.setSpacing(10); fn.setLabelAlignment(Qt.AlignRight)
        self.f_name  = QLineEdit(); self.f_name.setPlaceholderText("公司全称（必填）")
        self.f_code  = QLineEdit(); self.f_code.setPlaceholderText("助记码，如 ZY")
        self.f_type  = QComboBox()
        self.f_type.addItems(["小规模纳税人","一般纳税人","其他"])
        self.f_taxid = QLineEdit(); self.f_taxid.setPlaceholderText("统一社会信用代码")
        fn.addRow("公司名称 *", self.f_name); fn.addRow("助记码", self.f_code)
        fn.addRow("客户类型", self.f_type);   fn.addRow("税号",    self.f_taxid)
        L.addWidget(self.form_new)

        # 已有账套
        self.form_exist = QWidget(); fe = QFormLayout(self.form_exist)
        fe.setSpacing(10); fe.setLabelAlignment(Qt.AlignRight)
        self.f_exist_combo = QComboBox(); self.f_exist_combo.setMinimumWidth(280)
        fe.addRow("选择已有账套", self.f_exist_combo)
        warn = QLabel("  ⚠ 追加模式：同期间同凭证号自动跳过，不会重复写入。")
        warn.setStyleSheet("color:#ad6800;font-size:12px;")
        fe.addRow("", warn)
        self.form_exist.setVisible(False); L.addWidget(self.form_exist)
        L.addStretch()
        self.right.addWidget(w)
        self._reload_exist_clients()

    def _reload_exist_clients(self):
        self.f_exist_combo.clear()
        conn = get_db(); c = conn.cursor()
        c.execute("SELECT id,name FROM clients ORDER BY name")
        for r in c.fetchall():
            self.f_exist_combo.addItem(r["name"], r["id"])
        conn.close()

    def _toggle_mode(self, is_new):
        self.mode_new.setChecked(is_new); self.mode_exist.setChecked(not is_new)
        self.form_new.setVisible(is_new); self.form_exist.setVisible(not is_new)

    # ─────────────────── Step 2 · 选择文件 ───────────────────
    def _build_step2(self):
        w = QWidget(); L = QVBoxLayout(w); L.setContentsMargins(32,24,32,24); L.setSpacing(10)
        L.addWidget(lbl("选择要导入的账务文件", bold=True, size=14))
        hint = QLabel("  支持用友T3/T6/U8、金蝶KIS/K3导出的 XLS/XLSX 文件。"
                       "所有类型均可选，选得越多导入越完整。")
        hint.setStyleSheet("background:#f6f8ff;color:#444;border-radius:6px;"
                           "padding:8px 12px;font-size:12px;")
        hint.setWordWrap(True); L.addWidget(hint)

        file_types = [
            ("balance", "📊 科目余额表",
             "含科目编号、名称、期初余额 — 自动建科目并写期初数（必选）"),
            ("voucher", "📝 记账凭证",
             "全部凭证分录，自动识别用友 / 通用模板格式"),
            ("bank",    "🏦 银行存款日记账",
             "银行流水逐笔明细，写入 bank_statements 表供对账"),
            ("ledger",  "📒 总账 / 明细账",
             "按科目汇总的发生额与余额，用于核对"),
            ("income",  "💹 利润表",
             "收入费用汇总，仅用于人工核对，不写入账套"),
            ("bs",      "🏛 资产负债表",
             "资产负债期末数，仅用于人工核对，不写入账套"),
        ]
        self._file_path_lbls = {}
        for key, label, desc in file_types:
            row_w = QFrame()
            row_w.setStyleSheet("QFrame{background:#fff;border:1px solid #e8ecf2;"
                                "border-radius:8px;}")
            rl = QHBoxLayout(row_w); rl.setContentsMargins(16,10,16,10); rl.setSpacing(10)
            vv = QVBoxLayout()
            vv.addWidget(lbl(label, bold=True))
            vv.addWidget(lbl(desc, color="#888", size=12))
            rl.addLayout(vv); rl.addStretch()
            path_lbl = QLabel("未选择")
            path_lbl.setStyleSheet("color:#bbb;font-size:11px;min-width:160px;max-width:220px;")
            b_sel = QPushButton("选择"); b_sel.setObjectName("btn_outline"); b_sel.setFixedWidth(72)
            b_clr = QPushButton("✕");   b_clr.setObjectName("btn_gray");    b_clr.setFixedSize(26,26)
            b_sel.clicked.connect(lambda _,k=key,pl=path_lbl: self._pick_file(k, pl))
            b_clr.clicked.connect(lambda _,k=key,pl=path_lbl: self._clear_file(k, pl))
            rl.addWidget(path_lbl); rl.addWidget(b_sel); rl.addWidget(b_clr)
            self._file_path_lbls[key] = path_lbl
            L.addWidget(row_w)
        L.addStretch()
        self.right.addWidget(w)

    def _pick_file(self, key, lbl_w):
        path, _ = QFileDialog.getOpenFileName(
            self, "选择文件", "", "Excel 文件 (*.xls *.xlsx)")
        if not path: return
        self._files[key] = path
        import os; lbl_w.setText(os.path.basename(path))
        lbl_w.setStyleSheet("color:#3d6fdb;font-size:11px;")

    def _clear_file(self, key, lbl_w):
        self._files.pop(key, None)
        lbl_w.setText("未选择"); lbl_w.setStyleSheet("color:#bbb;font-size:11px;")

    # ─────────────────── Step 3 · 预览确认 ───────────────────
    def _build_step3(self):
        w = QWidget(); L = QVBoxLayout(w); L.setContentsMargins(32,24,32,24); L.setSpacing(10)
        L.addWidget(lbl("预览与确认", bold=True, size=14))
        self.preview_info = QLabel("")
        self.preview_info.setStyleSheet("background:#f6f8ff;color:#333;border-radius:6px;"
                                        "padding:10px 14px;font-size:12px;")
        self.preview_info.setWordWrap(True); L.addWidget(self.preview_info)
        L.addWidget(lbl("数据预览（首个文件前 20 行）:", bold=True, size=12))
        self.preview_tbl = QTableWidget()
        self.preview_tbl.setEditTriggers(QTableWidget.NoEditTriggers)
        self.preview_tbl.setShowGrid(True)
        self.preview_tbl.verticalHeader().setVisible(False)
        L.addWidget(self.preview_tbl)
        self.right.addWidget(w)

    def _refresh_preview(self):
        lines = []
        if self.mode_new.isChecked():
            lines.append(f"新建账套：{self.f_name.text().strip() or '（未填）'}")
        else:
            lines.append(f"追加到：{self.f_exist_combo.currentText()}")
        labels = {"balance":"科目余额表","voucher":"记账凭证","bank":"银行日记账",
                  "ledger":"总账","income":"利润表","bs":"资产负债表"}
        for k, v in labels.items():
            import os
            mark = f"✓  {os.path.basename(self._files[k])}" if k in self._files else "○  跳过"
            lines.append(f"{v}：{mark}")
        self.preview_info.setText("\n".join(lines))

        first = next((k for k in ["balance","voucher","bank","ledger"] if k in self._files), None)
        if not first:
            self.preview_tbl.setRowCount(0); self.preview_tbl.setColumnCount(0); return
        try:
            df = self._read_df(self._files[first])
            df = df.iloc[:20]
            cols = min(df.shape[1], 10)
            self.preview_tbl.setColumnCount(cols); self.preview_tbl.setRowCount(len(df))
            for ri, row in df.iterrows():
                self.preview_tbl.setRowHeight(ri, 26)
                for ci in range(cols):
                    self.preview_tbl.setItem(ri, ci, QTableWidgetItem(str(row.iloc[ci])))
            self.preview_tbl.horizontalHeader().setSectionResizeMode(
                QHeaderView.ResizeToContents)
        except Exception as e:
            self.preview_tbl.setRowCount(1); self.preview_tbl.setColumnCount(1)
            self.preview_tbl.setItem(0, 0, QTableWidgetItem(f"预览失败：{e}"))

    # ─────────────────── Step 4 · 导入完成 ───────────────────
    def _build_step4(self):
        w = QWidget(); L = QVBoxLayout(w); L.setContentsMargins(32,24,32,24); L.setSpacing(10)
        self.result_icon  = lbl("⏳", size=40, color="#3d6fdb")
        self.result_icon.setAlignment(Qt.AlignCenter)
        self.result_title = lbl("正在导入…", bold=True, size=15)
        self.result_title.setAlignment(Qt.AlignCenter)
        L.addStretch()
        L.addWidget(self.result_icon); L.addWidget(self.result_title); L.addSpacing(8)
        self.log_box = QTextEdit(); self.log_box.setReadOnly(True)
        self.log_box.setStyleSheet(
            "font-size:12px;font-family:monospace;background:#fafafa;border-radius:6px;")
        L.addWidget(self.log_box); L.addStretch()
        self.right.addWidget(w)

    # ─────────────────── 步骤控制 ───────────────────
    def _goto_step(self, idx):
        self._cur_step = idx
        self.right.setCurrentIndex(idx)
        for i, b in enumerate(self._step_btns):
            b.setProperty("active", "true" if i == idx else "false")
            b.style().unpolish(b); b.style().polish(b)
        self.btn_back.setVisible(0 < idx < 3)
        self.btn_next.setVisible(idx < 3)
        self.btn_close.setVisible(idx == 3)
        if idx == 2: self._refresh_preview()
        if idx == 3: QTimer.singleShot(100, self._do_import)

    def _prev_step(self):
        if self._cur_step > 0: self._goto_step(self._cur_step - 1)

    def _next_step(self):
        if self._cur_step == 0:
            if self.mode_new.isChecked() and not self.f_name.text().strip():
                QMessageBox.warning(self, "提示", "请填写公司名称"); return
        if self._cur_step == 1 and not self._files:
            QMessageBox.information(self, "提示", "请至少选择一个文件"); return
        self._goto_step(self._cur_step + 1)

    # ─────────────────── 工具方法 ───────────────────
    def _read_df(self, path):
        import pandas as pd
        engine = "xlrd" if path.endswith(".xls") else "openpyxl"
        try:
            return pd.read_excel(path, header=None, dtype=str,
                                 engine=engine).fillna("")
        except Exception:
            import xlrd
            wb = xlrd.open_workbook(path, ignore_workbook_corruption=True)
            ws = wb.sheet_by_index(0)
            data = [[str(ws.cell_value(r, c)) for c in range(ws.ncols)]
                    for r in range(ws.nrows)]
            import pandas as pd
            return pd.DataFrame(data).fillna("")

    def _flt(self, v):
        try: return float(str(v).replace(",", "").strip())
        except: return 0.0

    def _log(self, text):
        self.log_box.append(text)
        QApplication.processEvents()

    # ─────────────────── 核心导入 ───────────────────
    def _do_import(self):
        log = self._log
        log("🚀 开始导入账套…")
        conn = get_db(); c = conn.cursor()
        ok_total = err_total = 0
        try:
            # ── 创建 / 获取客户 ──
            if self.mode_new.isChecked():
                name  = self.f_name.text().strip()
                c.execute("SELECT id FROM clients WHERE name=?", (name,))
                row = c.fetchone()
                if row:
                    self._client_id = row[0]
                    log(f"ℹ 账套已存在，追加数据到：{name}")
                else:
                    c.execute(
                        "INSERT INTO clients(name,short_code,client_type,tax_id)"
                        " VALUES(?,?,?,?)",
                        (name, self.f_code.text().strip(),
                         self.f_type.currentText(),
                         self.f_taxid.text().strip()))
                    self._client_id = c.lastrowid
                    seed_client_accounts(self._client_id, conn)
                    conn.commit()
                    log(f"✓ 新建账套：{name}（ID={self._client_id}）")
            else:
                self._client_id = self.f_exist_combo.currentData()
                log(f"ℹ 追加到已有账套 ID={self._client_id}")

            # ── 1. 科目余额表 ──
            if "balance" in self._files:
                log("\n─── 导入科目余额表 ───")
                ok, err = self._imp_balance(conn, c)
                ok_total += ok; err_total += err

            # ── 2. 记账凭证 ──
            if "voucher" in self._files:
                log("\n─── 导入记账凭证 ───")
                ok, err = self._imp_voucher(conn, c)
                ok_total += ok; err_total += err

            # ── 3. 银行日记账 ──
            if "bank" in self._files:
                log("\n─── 导入银行日记账 ───")
                ok, err = self._imp_bank(conn, c)
                ok_total += ok; err_total += err

            # ── 4. 总账（补充科目期初） ──
            if "ledger" in self._files:
                log("\n─── 导入总账（补充期初） ───")
                ok, err = self._imp_ledger(conn, c)
                ok_total += ok; err_total += err

            # ── 5. 报表类 — 仅提示 ──
            for k, v in [("income","利润表"),("bs","资产负债表")]:
                if k in self._files:
                    log(f"\n○ {v} 已选择 — 仅供人工核对，不写入账套")

            log_action(conn, self._client_id, "账套导入", "import",
                       str(self._client_id),
                       f"文件:{list(self._files.keys())} 成功:{ok_total} 失败:{err_total}")
            conn.commit()
            log(f"\n✅ 导入完成！成功 {ok_total} 项，跳过/失败 {err_total} 项。")
            self.result_icon.setText("✅")
            self.result_title.setText("导入成功！")
        except Exception as e:
            conn.rollback()
            import traceback
            log(f"\n❌ 导入异常：{e}\n{traceback.format_exc()}")
            self.result_icon.setText("❌")
            self.result_title.setText("导入失败，请查看日志")
        finally:
            conn.close()

    # ── 科目余额表 ──
    def _imp_balance(self, conn, c):
        import re
        df = self._read_df(self._files["balance"])
        ok = err = 0
        for ri in range(len(df)):
            row = df.iloc[ri]
            code = str(row.iloc[1]).strip()
            name = str(row.iloc[2]).strip() if df.shape[1] > 2 else ""
            if not code or not name: continue
            if not re.match(r"^\d[\d._]*$", code): continue
            if len(code) < 4 or not code[:4].isdigit(): continue
            od = self._flt(row.iloc[3]) if df.shape[1] > 3 else 0
            oc = self._flt(row.iloc[4]) if df.shape[1] > 4 else 0
            # 推断类型 & 方向
            atype, direction = infer_account_type_direction(code, name)
            normalized = code.replace("_",".")
            level  = normalized.count(".") + 1
            parent = (code.rsplit(".",1)[0] if "." in code
                      else code.rsplit("_",1)[0] if "_" in code else None)
            c.execute("SELECT id FROM accounts WHERE client_id=? AND code=?",
                      (self._client_id, code))
            ex = c.fetchone()
            if ex:
                c.execute("UPDATE accounts SET name=?,full_name=?,type=?,direction=?,"
                          "opening_debit=?,opening_credit=? WHERE id=?",
                          (name, name, atype, direction, od, oc, ex[0]))
            else:
                c.execute("""INSERT INTO accounts(client_id,code,name,full_name,type,
                    direction,parent_code,level,opening_debit,opening_credit)
                    VALUES(?,?,?,?,?,?,?,?,?,?)""",
                    (self._client_id,code,name,name,atype,direction,parent,level,od,oc))
            ok += 1
            if ok <= 5 or ok % 30 == 0:
                self._log(f"  {code}  {name}  借={od:.2f}  贷={oc:.2f}")

        # ── 第二遍：处理含 _ 的辅助核算科目 ──
        aux_created = 0
        for ri in range(len(df)):
            row = df.iloc[ri]
            code = str(row.iloc[1]).strip()
            name = str(row.iloc[2]).strip() if df.shape[1] > 2 else ""
            if '_' not in code: continue
            if not re.match(r"^\d[\d._]*$", code): continue
            if process_aux_from_code(c, self._client_id, code, name):
                aux_created += 1
        if aux_created:
            self._log(f"  → 同步创建辅助核算对象 {aux_created} 个")
        self._log(f"  → 处理科目 {ok} 个")
        return ok, err

    # ── 记账凭证 ──
    def _imp_voucher(self, conn, c):
        import re
        df = self._read_df(self._files["voucher"])
        ok = skip = err = 0

        # 检测格式
        is_yonyou = any(
            "凭证字号" in str(df.iloc[r, 1] if df.shape[1] > 1 else "")
            for r in range(min(10, len(df))))

        title = str(df.iloc[0, 1]) if df.shape[1] > 1 else ""
        pm = re.search(r"(\d{4})年(\d{2})期", title)
        default_period = f"{pm.group(1)}-{pm.group(2)}" if pm else "2026-01"

        if is_yonyou:
            cur_vno = None; cur_date = ""; entries = []

            def flush(vno, date, ents):
                nonlocal ok, skip, err
                if not vno or not ents: return
                c.execute("SELECT id FROM vouchers WHERE client_id=? AND period=? AND voucher_no=?",
                          (self._client_id, default_period, vno))
                if c.fetchone(): skip += 1; return
                td = sum(e[3] for e in ents); tc = sum(e[4] for e in ents)
                if abs(td - tc) > 0.01: err += 1; return
                c.execute("INSERT INTO vouchers(client_id,period,voucher_no,date,status)"
                          " VALUES(?,?,?,?,?)",
                          (self._client_id, default_period, vno, date, "已审核"))
                vid = c.lastrowid
                for ln, ent in enumerate(ents, 1):
                    c.execute("INSERT INTO voucher_entries(voucher_id,line_no,summary,"
                              "account_code,account_name,debit,credit) VALUES(?,?,?,?,?,?,?)",
                              (vid, ln) + ent)
                ok += 1
                if ok <= 5 or ok % 20 == 0:
                    self._log(f"  ✓ {vno}  {len(ents)} 行  借={td:.2f}")

            for ri in range(len(df)):
                row = df.iloc[ri]
                cell1 = str(row.iloc[1] if df.shape[1] > 1 else "").strip()
                if "凭证字号" in cell1:
                    flush(cur_vno, cur_date, entries); entries = []
                    dm = re.search(r"日期:(\S+)", cell1)
                    nm = re.search(r"凭证字号:(\S+)", cell1)
                    cur_date = dm.group(1) if dm else default_period + "-28"
                    cur_vno  = nm.group(1).split()[0] if nm else None
                elif not cell1 or cell1.startswith("合计"): continue
                else:
                    af = str(row.iloc[2] if df.shape[1] > 2 else "").strip()
                    if not af: continue
                    parts = af.split(" ", 1)
                    # Preserve auxiliary-account separators from source files.
                    code = parts[0]
                    aname = parts[1] if len(parts) > 1 else af
                    d  = self._flt(row.iloc[3]) if df.shape[1] > 3 else 0
                    cr = self._flt(row.iloc[4]) if df.shape[1] > 4 else 0
                    if d == 0 and cr == 0: continue
                    entries.append((cell1, code, aname, d, cr))
            flush(cur_vno, cur_date, entries)
        else:
            # 通用模板格式
            from collections import OrderedDict
            vouchers = OrderedDict()
            n = df.shape[1]
            for ri in range(1, len(df)):
                row = df.iloc[ri]
                def gcol(i, default=""):
                    try:
                        v = str(row.iloc[i]).strip() if i < n else default
                        return v if v not in ("nan","") else default
                    except: return default
                def gcolf(i):
                    try:
                        v = row.iloc[i] if i < n else 0
                        return self._flt(v) if v and str(v) not in ("nan","") else 0
                    except: return 0
                period = gcol(0); vno = gcol(1); date = gcol(2)
                summ   = gcol(3); code = gcol(4); aname = gcol(5)
                d = gcolf(6); cr = gcolf(7)
                if not period or not vno or not code: continue
                key = (period, vno)
                if key not in vouchers:
                    vouchers[key] = {"period":period,"vno":vno,"date":date,"entries":[]}
                vouchers[key]["entries"].append((summ, code, aname, d, cr))
            for (period, vno), v in vouchers.items():
                c.execute("SELECT id FROM vouchers WHERE client_id=? AND period=? AND voucher_no=?",
                          (self._client_id, period, vno))
                if c.fetchone(): skip += 1; continue
                ents = v["entries"]
                td = sum(e[3] for e in ents); tc = sum(e[4] for e in ents)
                if abs(td - tc) > 0.01: err += 1; continue
                c.execute("INSERT INTO vouchers(client_id,period,voucher_no,date,status)"
                          " VALUES(?,?,?,?,?)",
                          (self._client_id, period, vno, v["date"], "已审核"))
                vid = c.lastrowid
                for ln, ent in enumerate(ents, 1):
                    c.execute("INSERT INTO voucher_entries(voucher_id,line_no,summary,"
                              "account_code,account_name,debit,credit) VALUES(?,?,?,?,?,?,?)",
                              (vid, ln) + ent)
                ok += 1

        self._log(f"  → 凭证导入 {ok} 张，跳过 {skip}，失败 {err}")
        return ok, skip + err

    # ── 银行日记账 → bank_statements ──
    def _imp_bank(self, conn, c):
        import re
        df = self._read_df(self._files["bank"])
        ok = skip = err = 0

        # 自动检测数据列布局
        # 用友明细账格式：col1=科目, col2=日期, col3=凭证号, col4=摘要,
        #                 col5=借方, col6=贷方, col8=余额
        # 通用银行流水格式：col0=日期, col1=摘要, col2=借/收入, col3=贷/支出, col4=余额

        # 先扫描找到第一个有 YYYY-MM-DD 格式日期的列和行
        date_col = None; data_start = 0
        for ri in range(min(15, len(df))):
            for ci in range(min(df.shape[1], 6)):
                v = str(df.iloc[ri, ci]).strip()
                if re.match(r"\d{4}[-/]\d{2}[-/]\d{2}", v):
                    date_col = ci; data_start = ri; break
            if date_col is not None: break

        if date_col is None:
            self._log("  ✗ 未能识别日期列，请确认文件格式"); return 0, 1

        # 根据 date_col 位置判断布局
        if date_col == 2:
            # 用友明细账：日期在 col2
            acct_col=1; vno_col=3; summ_col=4; d_col=5; cr_col=6; bal_col=8
        elif date_col == 0:
            # 通用银行流水：日期在 col0
            acct_col=None; vno_col=None; summ_col=1; d_col=2; cr_col=3; bal_col=4
        else:
            acct_col=None; vno_col=None; summ_col=date_col+1
            d_col=date_col+2; cr_col=date_col+3; bal_col=date_col+4

        def gcol(row, ci, default=""):
            try:
                v = str(row.iloc[ci]).strip() if ci is not None and ci < len(row) else default
                return v if v not in ("nan","") else default
            except: return default

        for ri in range(data_start, len(df)):
            row = df.iloc[ri]
            raw_date = gcol(row, date_col)
            # 标准化日期
            raw_date = raw_date.replace("/","-")
            if not re.match(r"\d{4}-\d{2}-\d{2}", raw_date): continue
            summary = gcol(row, summ_col)
            if summary in ("期初余额","本月合计","本年累计","合计",""): continue
            d  = self._flt(gcol(row, d_col))
            cr = self._flt(gcol(row, cr_col))
            bal = self._flt(gcol(row, bal_col)) if bal_col and bal_col < df.shape[1] else None
            if d == 0 and cr == 0: continue
            vno  = gcol(row, vno_col) if vno_col else ""
            # 科目信息
            acct_raw  = gcol(row, acct_col) if acct_col else "1002"
            parts = acct_raw.split(" ", 1)
            acct_code = parts[0] if re.match(r"^\d[\d.]*$", parts[0]) else "1002"
            acct_name = parts[1] if len(parts) > 1 else "银行存款"
            # 去重检查（日期+摘要+借+贷）
            c.execute("""SELECT id FROM bank_statements
                WHERE client_id=? AND date=? AND description=?
                AND debit=? AND credit=?""",
                (self._client_id, raw_date, summary, d, cr))
            if c.fetchone(): skip += 1; continue
            c.execute("""INSERT INTO bank_statements
                (client_id,account_code,account_name,date,voucher_no,
                 description,debit,credit,balance,source)
                VALUES(?,?,?,?,?,?,?,?,?,?)""",
                (self._client_id, acct_code, acct_name, raw_date,
                 vno, summary, d, cr, bal, "import"))
            ok += 1
            if ok <= 5 or ok % 50 == 0:
                self._log(f"  ✓ {raw_date}  {summary[:18]}  "
                          f"借={d:.2f}  贷={cr:.2f}")

        self._log(f"  → 银行流水导入 {ok} 条，跳过重复 {skip} 条")
        return ok, skip

    # ── 总账（补充期初） ──
    def _imp_ledger(self, conn, c):
        import re
        df = self._read_df(self._files["ledger"])
        ok = 0
        for ri in range(len(df)):
            row = df.iloc[ri]
            code = str(row.iloc[1]).strip() if df.shape[1] > 1 else ""
            name = str(row.iloc[2]).strip() if df.shape[1] > 2 else ""
            if not code or not name: continue
            if not re.match(r"^\d[\d._]*$", code): continue
            if len(code) < 4 or not code[:4].isdigit(): continue
            od = self._flt(row.iloc[3]) if df.shape[1] > 3 else 0
            oc = self._flt(row.iloc[4]) if df.shape[1] > 4 else 0
            if od == 0 and oc == 0: continue
            c.execute("SELECT id,opening_debit,opening_credit FROM accounts"
                      " WHERE client_id=? AND code=?", (self._client_id, code))
            ex = c.fetchone()
            if ex and ex["opening_debit"] == 0 and ex["opening_credit"] == 0:
                c.execute("UPDATE accounts SET opening_debit=?,opening_credit=? WHERE id=?",
                          (od, oc, ex["id"]))
                ok += 1
        self._log(f"  → 补充科目期初 {ok} 个")
        return ok, 0