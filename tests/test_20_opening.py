"""L3 科目期初：录入弹窗 + 期初页面 + 建账期(start_period)规则。两套会计制度都跑。"""
import pytest
from datetime import datetime
from helpers import press_release, type_text, spin_events, edit_of, APP
from PySide6.QtCore import QPoint, Qt
from PySide6.QtTest import QTest
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


def commit_focus(d):
    """让当前持有焦点的格子失去焦点——真人点「保存」按钮时就会发生这件事。
    有的金额框只在失焦时才提交数值，测试若不模拟这一步，读到的永远是 0（这会是测试的误报，不是缺陷）。"""
    fw = d.focusWidget()
    if fw is not None:
        fw.clearFocus()
    spin_events()


def cell_state(d, w, le, before="", typed=""):
    """现场快照：金额格的真实类型、内部输入控件、文本与数值、焦点——失败信息里带上它，不用再手工诊断。"""
    fw = d.focusWidget()
    return (f"[现场] 格子类型={[c.__name__ for c in type(w).__mro__[:4]]} 内部控件={type(le).__name__}({le.width()}x{le.height()}) "
            f"键入前文本={before!r} 键入后文本={typed!r} 当前文本={le.text()!r} value()={w.value() if hasattr(w, 'value') else 'n/a'} "
            f"焦点:格子={w.hasFocus()} 内部={le.hasFocus()} 弹窗焦点控件={type(fw).__name__ if fw else None} "
            f"弹窗已激活={d.isActiveWindow()} 只读={getattr(w, 'isReadOnly', lambda: None)()}")


def enter(d, row, col, text):
    """像用户一样：点击格子 → 敲键盘 → 按 Tab 离开格子。键入没有生效时，直接抛出带现场快照的错误。"""
    w = d.table.cellWidget(row, col); le = edit_of(w); before = le.text()
    assert le.width() > 40, "金额格内部输入区被压扁（样式表/Qt 版本问题）——见 test_00_env_canary.py"
    w.setFocus(); spin_events()
    press_release(le, QPoint(le.width() - 4, le.height() // 2))
    type_text(le, text)
    typed = le.text()
    QTest.keyClick(le, Qt.Key_Tab); spin_events()
    if hasattr(w, "value") and abs(w.value()) < 1e-9 and any(ch in "123456789" for ch in text):
        raise AssertionError(f"在弹窗的金额格里键入 {text!r} 并按 Tab 离开后，数值仍为 0。" + cell_state(d, w, le, before, typed))


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
    commit_focus(d)
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


@pytest.mark.gui
def test_typing_then_saving_without_leaving_the_cell_still_saves(env, std):
    """真人最常见的操作：敲完最后一个数，不按 Tab，直接点「保存」。最后一个格子的数也必须保存下来。"""
    cid = env.client(std, start="2026-04"); add_deep_accounts(env, cid)
    d, rm = open_dialog(cid)
    w = d.table.cellWidget(rm["1002.002"], 2); le = edit_of(w)
    w.setFocus(); spin_events()
    press_release(le, QPoint(le.width() - 4, le.height() // 2)); type_text(le, "1234.56")   # 不按 Tab
    typed = le.text()
    commit_focus(d)                                                                             # = 点击「保存」按钮时焦点离开
    snap = cell_state(d, w, le, "", typed)
    d._save(); d.close()
    assert env.opening(cid).get("1002.002", (0, 0))[0] == pytest.approx(1234.56), "敲完最后一个数直接保存，数字丢失。" + snap


@pytest.mark.gui
def test_parent_rollup_debit_credit_netting(env):
    """验证用户反馈 Bug B：三级及多级科目同时有借贷余额时，二级和一级上级汇总科目必须借贷相抵轧差，只在较大的一方留余额。"""
    cid = env.client("企业会计准则", start="2026-04")
    conn = env.db.get_db()
    # 建立三级科目结构：2221(一级) -> 2221.001(二级) -> 2221.001.001(三级进项), 2221.001.002(三级销项)
    conn.execute("INSERT OR IGNORE INTO accounts(client_id,code,name,type,direction) VALUES(?,?,?,?,?)",
                 (cid, "2221", "应交税费", "负债", "贷"))
    conn.execute("INSERT OR IGNORE INTO accounts(client_id,code,name,type,direction) VALUES(?,?,?,?,?)",
                 (cid, "2221.001", "应交增值税", "负债", "贷"))
    conn.execute("INSERT OR IGNORE INTO accounts(client_id,code,name,type,direction) VALUES(?,?,?,?,?)",
                 (cid, "2221.001.001", "进项税额", "负债", "借"))
    conn.execute("INSERT OR IGNORE INTO accounts(client_id,code,name,type,direction) VALUES(?,?,?,?,?)",
                 (cid, "2221.001.002", "销项税额", "负债", "贷"))
    conn.commit(); conn.close()

    d, rm = open_dialog(cid)
    enter(d, rm["2221.001.001"], 2, "10000.00")  # 进项税借方 10,000
    enter(d, rm["2221.001.002"], 3, "6000.00")   # 销项税贷方 6,000

    v = lambda code, col: d.table.cellWidget(rm[code], col).value()
    # 二级应交增值税：借 10000 - 贷 6000 = 净借方 4000，贷方为 0
    assert v("2221.001", 2) == pytest.approx(4000.0), "二级汇总科目借贷相抵后，借方应为净额 4000"
    assert v("2221.001", 3) == pytest.approx(0.0), "二级汇总科目贷方应轧差为 0"

    # 一级应交税费：同理净借方 4000，贷方为 0（不能两边同时挂数）
    assert v("2221", 2) == pytest.approx(4000.0), "一级汇总科目借贷相抵后，借方应为净额 4000"
    assert v("2221", 3) == pytest.approx(0.0), "一级汇总科目贷方应轧差为 0"
    d.close()


@pytest.mark.gui
def test_aux_opening_dialog_and_persistence(env):
    """验证用户反馈 Bug A：已配置辅助核算的科目在期初录入时识别辅助项，录入后自动汇总回填并持久化到 account_opening_aux。"""
    cid = env.client("小企业会计制度", start="2026-04")
    conn = env.db.get_db()
    # 创建辅助维度与项目：个人/员工（张三、李四）
    conn.execute("INSERT INTO aux_dimensions(client_id, name, code) VALUES(?, '员工', 'EMP')", (cid,))
    dim_id = conn.execute("SELECT id FROM aux_dimensions WHERE client_id=?", (cid,)).fetchone()[0]
    conn.execute("INSERT INTO aux_items(client_id, dimension_id, name, code) VALUES(?, ?, '张三', '001')", (cid, dim_id))
    conn.execute("INSERT INTO aux_items(client_id, dimension_id, name, code) VALUES(?, ?, '李四', '002')", (cid, dim_id))
    i1 = conn.execute("SELECT id FROM aux_items WHERE client_id=? AND name='张三'", (cid,)).fetchone()[0]
    i2 = conn.execute("SELECT id FROM aux_items WHERE client_id=? AND name='李四'", (cid,)).fetchone()[0]
    # 配置科目 2241 (其他应付款) 挂接该维度
    conn.execute("INSERT INTO account_aux_config(client_id, account_code, dimension_id) VALUES(?, '2241', ?)",
                 (cid, dim_id))
    conn.commit(); conn.close()

    d, rm = open_dialog(cid)
    # 验证科目名称带上了辅助核算标签
    name_txt = d.table.item(rm["2241"], 1).text()
    assert "员工" in name_txt, f"期初科目名称应展示辅助核算维度标签，实际得到: {name_txt}"

    # 模拟辅助核算录入：张三贷方 5000，李四贷方 3000
    d._aux_openings["2241"] = {i1: (0.0, 5000.0), i2: (0.0, 3000.0)}
    dw = d.table.cellWidget(rm["2241"], 2)
    cw = d.table.cellWidget(rm["2241"], 3)
    cw.setValue(8000.0)
    d._refresh_parents()

    # 平衡保存（借方 1002 录入 8000）
    enter(d, rm["1002"], 2, "8000.00")
    d._save(); d.close()

    # 验证 account_opening_aux 持久化数据
    conn = env.db.get_db()
    rows = conn.execute("""
        SELECT ai.name, ao.debit, ao.credit
        FROM account_opening_aux ao
        JOIN aux_items ai ON ai.id = ao.aux_item_id
        WHERE ao.client_id=? AND ao.account_code='2241'
        ORDER BY ai.name
    """, (cid,)).fetchall()
    conn.close()

    assert len(rows) == 2, f"应持久化 2 条辅助核算明细，实际: {len(rows)}"
    assert rows[0]["name"] == "张三" and rows[0]["credit"] == 5000.0
    assert rows[1]["name"] == "李四" and rows[1]["credit"] == 3000.0

    # 验证重新打开弹窗时，数据成功回显
    d2, rm2 = open_dialog(cid)
    assert d2.table.cellWidget(rm2["2241"], 3).value() == pytest.approx(8000.0)
    assert "2241" in d2._aux_openings
    assert d2._aux_openings["2241"][i1] == (0.0, 5000.0)
    assert d2._aux_openings["2241"][i2] == (0.0, 3000.0)
    d2.close()

