"""L2/L3 凭证页：科目余额表口径、导出一致性、封账只读、按钮可见性。"""
import openpyxl
import pytest
from PySide6.QtWidgets import QPushButton
from helpers import (CODES, PERIOD, carry_forward, income_rows, voucher_page, balance_table, find_row, num, APP)


def buttons_of_rows(pg):
    return [[b.text() for b in pg.v_tbl.cellWidget(r, 6).findChildren(QPushButton)]
            for r in range(pg.v_tbl.rowCount()) if pg.v_tbl.cellWidget(r, 6)]


@pytest.mark.gui
def test_balance_table_excludes_carry_vouchers_and_matches_income_statement(env, std):
    """科目余额表的本期发生额不能被结转凭证抵消；收入科目净额应等于利润表营业收入。"""
    cid = env.client(std); K = env.standard_month(cid, std); carry_forward(cid)
    bt = balance_table(voucher_page(cid))
    inv, rev = env.leaf(cid, K["inv"]), env.leaf(cid, K["rev"])
    assert bt[inv][2] == pytest.approx(0) and bt[inv][3] == pytest.approx(300), f"投资收益本期借/贷 {bt[inv][2:4]}，应为 0/300（结转凭证被计入了发生额）"
    assert bt[rev][3] == pytest.approx(10000)
    assert num(find_row(income_rows(cid), "营业收入")[2]) == pytest.approx(bt[rev][3] - bt[rev][2])


@pytest.mark.gui
def test_balance_export_matches_screen_cell_by_cell(env, std, files, tmp_path):
    """导出的 Excel 必须与界面逐科目逐列一致；曾因合并标题行 + column_letter 导致导出直接崩溃。"""
    cid = env.client(std); env.standard_month(cid, std); carry_forward(cid)
    pg = voucher_page(cid); ui = balance_table(pg)
    files.save = str(tmp_path / "bal.xlsx")
    pg._export_balance()
    ws = openpyxl.load_workbook(files.save).active
    xl = {str(r[0]): [float(x or 0) for x in r[2:8]] for r in ws.iter_rows(min_row=3, values_only=True)
          if r[0] and r[1] and "合计" not in str(r[1])}
    missing = [c for c in ui if c not in xl]
    diff = [c for c in ui if c in xl and any(abs(a - b) > 0.005 for a, b in zip(ui[c], xl[c]))]
    assert not missing and not diff, f"导出与界面不一致：缺失 {missing[:5]} 差异 {diff[:5]}"


@pytest.mark.gui
def test_locked_period_shows_view_button_only_and_readonly_dialog(env, std):
    cid = env.client(std); env.post(cid, PERIOD, [("1002", 10, 0), (CODES[std]["rev"], 0, 10)])
    conn = env.db.get_db(); conn.execute("insert into periods(client_id,period,is_closed) values(?,?,1)", (cid, PERIOD)); conn.commit(); conn.close()
    pg = voucher_page(cid); pg._load_vouchers(); APP.processEvents()
    rows = buttons_of_rows(pg)
    assert rows and all(r == ["查看"] for r in rows), f"封账期间操作列应只有「查看」，实际 {rows}"
    from dialogs.voucher_dialogs import VoucherDialog
    vid = env.db.get_db().execute("select id from vouchers where client_id=? limit 1", (cid,)).fetchone()[0]
    d = VoucherDialog(None, cid, PERIOD, vid, readonly=True); d.show(); APP.processEvents()
    texts = [b.text() for b in d.findChildren(QPushButton)]
    assert d.windowTitle() == "查看凭证"
    assert not any(k in t for t in texts for k in ("保存", "保 存", "模板", "增行", "删行")), texts
    assert all(not d.table.cellWidget(r, c).isEnabled() for r in range(d.table.rowCount())
               for c in range(d.table.columnCount()) if d.table.cellWidget(r, c)), "只读弹窗里仍有可编辑控件"
    assert not d.date_edit.isEnabled() and not d.attach_spin.isEnabled()


@pytest.mark.gui
def test_open_period_keeps_edit_and_approve_buttons(env, std):
    cid = env.client(std)
    env.post(cid, PERIOD, [("1002", 10, 0), (CODES[std]["rev"], 0, 10)], status="待审核")
    env.post(cid, PERIOD, [("1002", 20, 0), (CODES[std]["rev"], 0, 20)], status="已审核")
    pg = voucher_page(cid); pg._load_vouchers(); APP.processEvents()
    rows = buttons_of_rows(pg)
    assert any({"✓ 审核通过", "✗ 拒绝", "编辑", "删除"} <= set(r) for r in rows), rows
    assert any("↩ 撤销审核" in r and "查看" not in r for r in rows), rows


@pytest.mark.gui
def test_unbalanced_voucher_cannot_be_approved_and_locked_period_blocks_status_change(env):
    cid = env.client("企业会计准则")
    ok = env.post(cid, PERIOD, [("1002", 11300, 0), ("6001", 0, 10000), ("2221.001.001", 0, 1300)], status="待审核", to_leaf=False)
    bad = env.post(cid, PERIOD, [("1002", 5000, 0), ("6001", 0, 4000)], status="待审核", to_leaf=False)
    pg = voucher_page(cid)
    status = lambda v: env.db.get_db().execute("select status from vouchers where id=?", (v,)).fetchone()[0]
    pg._set_voucher_status(ok, "已审核");  assert status(ok) == "已审核"
    pg._set_voucher_status(bad, "已审核"); assert status(bad) == "待审核", "借贷不平的凭证不应能被审核"
    conn = env.db.get_db(); conn.execute("insert into periods(client_id,period,is_closed) values(?,?,1)", (cid, PERIOD)); conn.commit(); conn.close()
    pg._set_voucher_status(ok, "待审核");  assert status(ok) == "已审核", "封账期间不应能撤销审核"
