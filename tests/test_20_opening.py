"""L3 科目期初：录入弹窗 + 期初页面 + 建账期(start_period)规则。两套会计制度都跑。"""
import pytest
from datetime import datetime
from helpers import press_release, type_text, spin_events, edit_of, APP
from PySide6.QtCore import QPoint
from dialogs.client_dialogs import AccountInitDialog

DEEP = [("1002.001", "银行存款-工行", "1002", 2), ("1002.001.001", "工行基本户", "1002.001", 3),
        ("1002.001.002", "工行一般户", "1002.001", 3), ("1002.002", "银行存款-建行", "1002", 2)]


def add_deep_accounts(env, cid):
    conn = env.db.get_db()
    for code, name, parent, lvl in DEEP:
        conn.execute("insert or ignore into accounts(client_id,code,name,full_name,type,direction,parent_code,level,is_default) "
                     "values(?,?,?,?,?,?,?,?,0)", (cid, code, name, name, "资产", "借", parent, lvl))
    conn.commit(); conn.close()


def open_dialog(cid, period="2026-04"):
    d = AccountInitDialog(None, cid, period); d.show(); spin_events()
    return d, {d.table.item(i, 0).text(): i for i in range(d.table.rowCount())}


def enter(d, row, col, text):
    """像用户一样：点击格子 → 敲键盘。"""
    w = d.table.cellWidget(row, col); le = edit_of(w)
    assert le.width() > 40, "金额格内部输入区被压扁（样式表/Qt 版本问题）——见 test_00_env_canary.py"
    w.setFocus(); spin_events()
    press_release(le, QPoint(le.width() - 4, le.height() // 2))
    type_text(le, text)


def first_leaf(d, rm, prefix):
    return next(c for c, i in rm.items() if c.startswith(prefix) and d._leaf_flags[i])


@pytest.fixture
def dlg(env, std):
    cid = env.client(std, start="2026-04")
    add_deep_accounts(env, cid)
    d, rm = open_dialog(cid)
    yield env, cid, d, rm
    d.close()


@pytest.mark.gui
def test_leaf_and_parent_flags(dlg):
    env, cid, d, rm = dlg
    flags = {c: d._leaf_flags[i] for c, i in rm.items()}
    assert not flags["1002"] and not flags["1002.001"], "有下级的科目必须是父级（不可录入）"
    assert flags["1002.001.001"] and flags["1002.002"], "末级科目必须可录入"
    assert d.table.cellWidget(rm["1002"], 2).isReadOnly()
    assert not d.table.cellWidget(rm["1002.001.001"], 2).isReadOnly()


@pytest.mark.gui
def test_keyboard_entry_rolls_up_to_parents_and_balances(dlg):
    env, cid, d, rm = dlg
    cr = first_leaf(d, rm, "2")
    enter(d, rm["1002.001.001"], 2, "1234.56")
    enter(d, rm["1002.001.002"], 2, "0.5")            # 小数不能被吞
    enter(d, rm["1002.002"], 2, "22")                 # 不能被乘 10
    enter(d, rm[cr], 3, "1257.06")
    v = lambda code, col: d.table.cellWidget(rm[code], col).value()
    assert v("1002.001.001", 2) == pytest.approx(1234.56)
    assert v("1002.001.002", 2) == pytest.approx(0.5)
    assert v("1002.002", 2) == pytest.approx(22)
    assert v("1002.001", 2) == pytest.approx(1235.06), "二级父级应等于下级之和"
    assert v("1002", 2) == pytest.approx(1257.06), "一级父级应等于全部末级之和"
    txt = d.balance_lbl.text()
    assert "借贷平衡" in txt and "1,257.06" in txt, txt        # 只统计末级，父级不能重复计入


@pytest.mark.gui
def test_save_persists_and_reopens_identically(dlg, msgs):
    env, cid, d, rm = dlg
    cr = first_leaf(d, rm, "2")
    enter(d, rm["1002.001.001"], 2, "1234.56"); enter(d, rm["1002.001.002"], 2, "0.5")
    enter(d, rm["1002.002"], 2, "22"); enter(d, rm[cr], 3, "1257.06")
    d._save()
    saved = env.opening(cid)
    assert saved["1002.001.001"][0] == pytest.approx(1234.56)
    assert saved["1002.001.002"][0] == pytest.approx(0.5)
    assert saved["1002"][0] == pytest.approx(1257.06), "父级汇总值应一并写入"
    assert saved[cr][1] == pytest.approx(1257.06)
    assert env.audits("保存期初余额"), "保存期初必须写审计日志"
    d2, rm2 = open_dialog(cid)
    assert d2.table.cellWidget(rm2["1002.001.002"], 2).value() == pytest.approx(0.5)
    d2.close()


# ───────────── 建账期 start_period 规则 ─────────────
# 规则：建账期只允许前移（取较早者）——选更早的月份建账可以，但在「次年1月(年初)」保存不能把建账期推后。
START_CASES = [
    ("A 次年1月(年初)保存，不得推后", "2026-04", "2027-01", "2026-04"),
    ("B 建账月早于创建月，应前移",     "2026-11", "2026-03", "2026-03"),
    ("C 同月再次保存，不变",           "2026-04", "2026-04", "2026-04"),
    ("D 晚于创建月的普通月份，不变",   "2026-04", "2026-06", "2026-04"),
]

@pytest.mark.gui
@pytest.mark.parametrize("desc,created,saved_at,expect", START_CASES, ids=[c[0] for c in START_CASES])
def test_start_period_only_moves_forward_in_time(env, desc, created, saved_at, expect):
    cid = env.client("小企业会计制度", start=created)
    d, rm = open_dialog(cid, saved_at); d._save(); d.close()
    assert env.start_period(cid) == expect


@pytest.mark.gui
def test_start_period_move_is_audited(env):
    cid = env.client("小企业会计制度", start="2026-11")
    d, rm = open_dialog(cid, "2026-03"); d._save(); d.close()
    logs = env.audits("调整建账期")
    assert logs and "2026-11" in logs[0][-1] and "2026-03" in logs[0][-1], "建账期前移必须写审计日志（含前后值）"


# ───────────── 期初页面：可选期间与标签 ─────────────
@pytest.fixture
def page(env):
    import pages.opening as O
    now = datetime.now()
    def make(cid):
        pg = O.OpeningBalancePage(); pg.set_client(cid, "测试", f"{now.year}-{now.month:02d}"); APP.processEvents()
        return pg, {pg.period_combo.itemData(i): pg.period_combo.itemText(i) for i in range(pg.period_combo.count())}
    return make


@pytest.mark.gui
def test_new_client_any_month_selectable_and_defaults_to_start_period(env, page):
    pg, items = page(env.client("小企业会计制度", start="2026-04"))
    assert pg.period_combo.currentData() == "2026-04"
    assert "建账期" in items["2026-04"] and "年初" in items["2026-01"]
    assert "当前月" in items.get(f"{datetime.now().year}-{datetime.now().month:02d}", "")
    assert all(f"{y}-{m:02d}" in items for y in (2018, 2022, 2026) for m in (1, 6, 12))


@pytest.mark.gui
def test_client_with_vouchers_cannot_start_after_earliest_voucher(env, page):
    cid = env.client("小企业会计制度", start="2026-04")
    env.post(cid, "2026-04", [("1002", 10, 0), ("5001", 0, 10)])
    pg, items = page(cid)
    assert "2026-04" in items and "2026-02" in items, "最早凭证期及更早的月份可作建账期"
    assert "2026-06" not in items, "晚于最早凭证的非 1 月月份不可作建账期"
    assert "2027-01" in items, "每年 1 月（年初结转）始终可选"


@pytest.mark.gui
def test_preview_balance_bar(env, page):
    cid = env.client("企业会计准则", start="2026-04")
    conn = env.db.get_db()
    conn.execute("update accounts set opening_debit=1000 where client_id=? and code='1002'", (cid,))
    conn.execute("update accounts set opening_credit=1000 where client_id=? and code='4001'", (cid,)); conn.commit(); conn.close()
    pg, _ = page(cid); pg._load_preview()
    assert "借贷平衡" in pg.balance_lbl.text()
