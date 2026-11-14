"""tests/helpers.py — 测试数据构造 + 对页面私有接口的适配层

约定：测试文件不直接调用页面的私有方法，一律经过这里。
将来页面重构、方法改名，只需要改这一个文件。
"""
import os
from datetime import datetime
from PySide6.QtCore import Qt, QPoint, QEvent
from PySide6.QtGui import QKeyEvent, QInputMethodEvent
from PySide6.QtTest import QTest
from PySide6.QtWidgets import QApplication, QWidget, QVBoxLayout, QPushButton

APP = QApplication.instance()
PERIOD = "2026-10"

# 两套制度下与本项目科目表一致的科目编码（来自 STANDARD_ACCOUNTS / STANDARD_ACCOUNTS_SMALL）
CODES = {
    "企业会计准则":   dict(rev="6001", cost="6401", tax="6403", sell="6601", mgmt="6602", fin="6603",
                           inv="6111", nin="6301", nout="6711", inc="6801", profit="4103", lo="6001", hi="7000"),
    "小企业会计制度": dict(rev="5001", cost="5401", tax="5403", sell="5501", mgmt="5502", fin="5503",
                           inv="5111", nin="5301", nout="5601", inc="5701", profit="3103", lo="5001", hi="6000"),
}


def edit_of(w):
    """金额框内部真正接收键盘/鼠标的输入控件：QDoubleSpinBox 子类取 lineEdit()，QLineEdit 子类就是它自己。"""
    return w.lineEdit() if hasattr(w, "lineEdit") else w


def spin_events():
    """让事件循环转两圈：真人点击时，按下与松开之间定时器会触发。"""
    APP.processEvents(); APP.processEvents()


# ───────────────────────── 数据库环境 ─────────────────────────
class Env:
    def __init__(self, db, tmp_path):
        self.db = db
        self.tmp = tmp_path
        self._compat()

    def _compat(self):
        """仅为兼容较旧的 db.py：真实项目已具备 start_period 与 get_client_start_period 时这里什么都不做。"""
        db = self.db
        conn = db.get_db()
        cols = [r[1] for r in conn.execute("PRAGMA table_info(clients)")]
        if "start_period" not in cols:
            conn.execute("ALTER TABLE clients ADD COLUMN start_period TEXT"); conn.commit()
        conn.close()
        if not hasattr(db, "get_client_start_period"):
            def _gcsp(conn, client_id):
                r = conn.execute("select start_period from clients where id=?", (client_id,)).fetchone()
                if r and r[0]: return r[0]
                r = conn.execute("select min(period) from vouchers where client_id=?", (client_id,)).fetchone()
                if r and r[0]: return r[0]
                return datetime.now().strftime("%Y-%m")
            db.get_client_start_period = _gcsp

    # ── 构造数据 ──
    def client(self, std, name=None, start="2026-01"):
        name = name or f"测试-{std}"
        conn = self.db.get_db(); c = conn.cursor()
        cols = [r[1] for r in c.execute("PRAGMA table_info(clients)")]
        f = {"name": name, "accounting_std": std}
        if "code" in cols: f["code"] = name
        if "start_period" in cols: f["start_period"] = start
        c.execute(f"INSERT INTO clients({','.join(f)}) VALUES({','.join('?' * len(f))})", list(f.values()))
        cid = c.lastrowid
        self.db.seed_client_accounts(cid, conn, std)
        conn.commit(); conn.close()
        return cid

    def codes_of(self, cid):
        conn = self.db.get_db(); r = {x[0] for x in conn.execute("select code from accounts where client_id=?", (cid,))}; conn.close(); return r

    def leaf(self, cid, code):
        """系统不允许直接往有下级的科目记账：返回该科目下第一个末级科目（自身已是末级则返回自身）。"""
        codes = self.codes_of(cid)
        if not any(o != code and o.startswith(code + ".") for o in codes):
            return code
        return sorted(o for o in codes if o.startswith(code + ".") and not any(x != o and x.startswith(o + ".") for x in codes))[0]

    def post(self, cid, period, entries, status="已审核", note="", no=None, to_leaf=True):
        """entries: [(科目, 借, 贷), ...]；返回凭证 id"""
        conn = self.db.get_db(); c = conn.cursor()
        no = no or f"记-{(c.execute('select count(*) from vouchers where client_id=? and period=?', (cid, period)).fetchone()[0] + 1):03d}"
        c.execute("insert into vouchers(client_id,period,voucher_no,date,status,note) values(?,?,?,?,?,?)",
                  (cid, period, no, period + "-15", status, note))
        vid = c.lastrowid
        for i, (code, dr, cr) in enumerate(entries, 1):
            code = self.leaf(cid, code) if to_leaf else code
            nm = (c.execute("select name from accounts where client_id=? and code=?", (cid, code)).fetchone() or ["?"])[0]
            c.execute("insert into voucher_entries(voucher_id,line_no,summary,account_code,account_name,debit,credit) values(?,?,?,?,?,?,?)",
                      (vid, i, "测试", code, nm, dr, cr))
        conn.commit(); conn.close()
        return vid

    def standard_month(self, cid, std, period=PERIOD):
        """一个月的标准业务：收入10000、成本4000、税金100、销售500、管理2000、财务50、
        投资收益+300、营业外收200、营业外支50、所得税400、短期投资/交易性金融资产5000。
        期望：营业利润3650、利润总额3800、净利润3400。"""
        K = CODES[std]
        for e in ([("1002", 10000, 0), (K["rev"], 0, 10000)], [(K["cost"], 4000, 0), ("1002", 0, 4000)],
                  [(K["tax"], 100, 0), ("1002", 0, 100)], [(K["sell"], 500, 0), ("1002", 0, 500)],
                  [(K["mgmt"], 2000, 0), ("1002", 0, 2000)], [(K["fin"], 50, 0), ("1002", 0, 50)],
                  [("1002", 300, 0), (K["inv"], 0, 300)], [("1002", 200, 0), (K["nin"], 0, 200)],
                  [(K["nout"], 50, 0), ("1002", 0, 50)], [(K["inc"], 400, 0), ("1002", 0, 400)],
                  [("1101", 5000, 0), ("1002", 0, 5000)]):
            self.post(cid, period, list(e))
        return K

    # ── 查询 ──
    def audits(self, action):
        c = self.db.get_db(); r = [tuple(x) for x in c.execute(
            "select client_id,operator,target_type,target_id,detail from audit_log where action=?", (action,))]
        c.close(); return r

    def start_period(self, cid):
        c = self.db.get_db(); r = c.execute("select start_period from clients where id=?", (cid,)).fetchone()[0]; c.close(); return r

    def opening(self, cid):
        c = self.db.get_db()
        r = {x[0]: (x[1], x[2]) for x in c.execute(
            "select code,opening_debit,opening_credit from accounts where client_id=? and (opening_debit<>0 or opening_credit<>0)", (cid,))}
        c.close(); return r


# ───────────────────────── 表格读取 ─────────────────────────
def dump_table(tbl):
    return [[(tbl.item(r, c).text() if tbl.item(r, c) else "") for c in range(tbl.columnCount())] for r in range(tbl.rowCount())]

def num(s):
    s = (s or "").replace(",", "").strip()
    try: return float(s)
    except ValueError: return None

def find_row(rows, label):
    for r in rows:
        if label in r[0].replace("\n", ""): return r
    return None


# ───────────────────────── 页面适配 ─────────────────────────
def carry_forward(cid, period=PERIOD):
    import pages.settle as S
    pg = S.SettlePage(); pg.set_client(cid, "X", period); pg._refresh_carry_amounts()
    pg.card_income._cb.setChecked(True); pg.card_expense._cb.setChecked(True); pg._do_carryforward()
    return pg

def income_rows(cid, period=PERIOD):
    import pages.report as R
    pg = R.ReportPage(); pg.set_client(cid, "X", period); pg._load_income(); APP.processEvents()
    return dump_table(pg.inc_tbl)

def balance_sheet_rows(cid, period=PERIOD):
    import pages.report as R
    pg = R.ReportPage(); pg.set_client(cid, "X", period); pg._load_balance(); APP.processEvents()
    return dump_table(pg.bs_tbl)

def bs_totals(rows):
    """返回 (资产总计, 负债和所有者权益总计)"""
    r = find_row(rows, "资产总计")
    return (num(r[2]) or 0.0, num(r[6]) or 0.0) if r else (None, None)

def voucher_page(cid, period=PERIOD):
    import pages.voucher as V
    pg = V.VoucherPage(); pg.set_client(cid, "X", period)
    for combo in (pg.bal_start_period, pg.bal_end_period):
        combo.setCurrentIndex(combo.findData(period))
    pg.period_combo.setCurrentIndex(pg.period_combo.findData(period))
    return pg

def balance_table(pg):
    pg._load_balance(); APP.processEvents(); t = pg.bal_tbl; out = {}
    for r in range(t.rowCount() - 1):
        it = t.item(r, 0)
        if it and it.text():
            out[it.text()] = [float((t.item(r, c).text() or "0").replace(",", "") or 0) for c in range(2, 8)]
    return out


# ───────────────────────── 金额框交互 ─────────────────────────
def set_style(style):
    import utils
    APP.setStyle(style); APP.setStyleSheet(utils.SS)

def make_spin(cls, pre=0.0, click=None, editable=True, focused_first=False, size=(140, 32)):
    """在一个容器里创建金额框。默认由别的按钮持有焦点，金额框处于「未聚焦」状态（模拟用户首次点击）。
    click: None / 'left' / 'mid' / 'right'。点击采用真人时序：按下 → 事件循环 → 松开。"""
    box = QWidget(); lay = QVBoxLayout(box); btn = QPushButton("x")
    w = cls(editable=editable) if cls.__name__ == "AccountOpeningSpinBox" else cls()
    if cls.__name__ != "AccountOpeningSpinBox":
        w.setRange(0, 9999999999); w.setDecimals(2); w.setAlignment(Qt.AlignRight)
    lay.addWidget(btn); lay.addWidget(w)
    w.setValue(pre); w.setFixedSize(*size); box.resize(size[0] + 80, 140); box.show(); spin_events()
    (w if focused_first else btn).setFocus(); spin_events()
    le = edit_of(w)
    assert le.width() > 40, (f"金额框内部输入区宽度只有 {le.width()}px（控件宽 {size[0]}px）——数字和光标都会看不见。"
                             "这是样式表与当前 Qt 版本不兼容导致的，见 test_00_env_canary.py")
    if click:
        x = {"left": 4, "mid": le.width() // 2, "right": le.width() - 4}[click]
        press_release(le, QPoint(x, le.height() // 2))
    w._box = box        # 防止被垃圾回收
    return w, le

def press_release(le, p):
    QTest.mousePress(le, Qt.LeftButton, Qt.NoModifier, p); spin_events()
    QTest.mouseRelease(le, Qt.LeftButton, Qt.NoModifier, p); spin_events()

def key(le, ch):
    if ord(ch) < 128:
        QTest.keyClicks(le, ch)
    else:   # QTest 不能发送非 ASCII 字符：手工构造按键事件（全角数字、中文标点）
        QApplication.sendEvent(le, QKeyEvent(QEvent.KeyPress, 0, Qt.NoModifier, ch))
        QApplication.sendEvent(le, QKeyEvent(QEvent.KeyRelease, 0, Qt.NoModifier, ch))
    APP.processEvents()

def type_text(le, s):
    for ch in s: key(le, ch)

def ime_commit(le, s):
    """模拟输入法提交文本（Windows 中文输入法下数字/标点可能以此方式到达，而不是 KeyPress）"""
    ev = QInputMethodEvent(); ev.setCommitString(s); QApplication.sendEvent(le, ev); APP.processEvents()

def ime_type(le, s):
    for ch in s: ime_commit(le, ch)

def paste_clipboard(le, text):
    QApplication.clipboard().setText(text); QTest.keyClick(le, Qt.Key_V, Qt.ControlModifier); APP.processEvents()


def delete_client(cid, name="测试客户"):
    """走页面真实的删除流程（确认框已被 conftest 替换为自动点「是」）。"""
    import pages.client as C
    pg = C.ClientPage(); pg._del({"id": cid, "name": name})
