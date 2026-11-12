"""L2 期末结转 —— 两套会计制度。

覆盖：收入科目出现借方余额（投资亏损）、费用科目出现贷方余额（费用冲回、利息收入大于利息支出）、
结转凭证借贷平衡、损益科目全部结平、本年利润 = 利润。
"""
import pytest
from helpers import CODES, PERIOD, carry_forward


def pl_net_after(env, cid, K, period=PERIOD):
    """结转后，损益区间内每个科目的 贷-借 净额（应全部为 0）。"""
    conn = env.db.get_db()
    rows = conn.execute("""select e.account_code, round(sum(e.credit)-sum(e.debit),2) from voucher_entries e
        join vouchers v on v.id=e.voucher_id where v.client_id=? and v.period=? and e.account_code>=? and e.account_code<?
        group by e.account_code""", (cid, period, K["lo"], K["hi"])).fetchall()
    profit = conn.execute("""select round(sum(credit)-sum(debit),2) from voucher_entries e join vouchers v on v.id=e.voucher_id
        where v.client_id=? and v.period=? and e.account_code=?""", (cid, period, K["profit"])).fetchone()[0] or 0
    unbalanced = []
    for vid, no, note in conn.execute("select id,voucher_no,note from vouchers where client_id=? and note like '结转%'", (cid,)):
        d, c = conn.execute("select sum(debit),sum(credit) from voucher_entries where voucher_id=?", (vid,)).fetchone()
        if abs((d or 0) - (c or 0)) > 0.005: unbalanced.append(no)
    notes = [r[0] for r in conn.execute("select note from vouchers where client_id=? and note like '结转%' order by id", (cid,))]
    conn.close()
    return {r[0]: r[1] for r in rows if abs(r[1]) > 0.005}, profit, unbalanced, notes


@pytest.mark.gui
def test_carry_forward_with_investment_loss_and_expense_reversal(env, std):
    """收入10000+营业外收40；投资亏损300（收入类科目借方余额）；管理费用2000-冲回50；营业外支出100。
    利润 = 10000+40-300-(2000-50)-100 = 7690。"""
    cid = env.client(std); K = CODES[std]
    env.post(cid, PERIOD, [("1002", 10000, 0), (K["rev"], 0, 10000)])
    env.post(cid, PERIOD, [(K["inv"], 300, 0), ("1002", 0, 300)])
    env.post(cid, PERIOD, [(K["mgmt"], 2000, 0), ("1002", 0, 2000)])
    env.post(cid, PERIOD, [("1002", 50, 0), (K["mgmt"], 0, 50)])
    env.post(cid, PERIOD, [("1002", 40, 0), (K["nin"], 0, 40)])
    env.post(cid, PERIOD, [(K["nout"], 100, 0), ("1002", 0, 100)])
    carry_forward(cid)
    left, profit, unbalanced, notes = pl_net_after(env, cid, K)
    assert left == {}, f"结转后仍有损益科目未结平：{left}（投资亏损等借方余额的收入类科目被漏掉）"
    assert not unbalanced, f"结转凭证借贷不平：{unbalanced}"
    assert profit == pytest.approx(7690.0), f"本年利润 {profit}，应为 7690"
    assert notes == ["结转收入", "结转费用"]


@pytest.mark.gui
def test_carry_forward_expense_account_with_net_credit_balance(env, std):
    """财务费用整体是贷方余额（利息收入 > 利息支出）也必须结转。"""
    cid = env.client(std); K = CODES[std]
    env.post(cid, PERIOD, [("1002", 10000, 0), (K["rev"], 0, 10000)])
    env.post(cid, PERIOD, [(K["mgmt"], 2000, 0), ("1002", 0, 2000)])
    env.post(cid, PERIOD, [("1002", 80, 0), (K["fin"], 0, 80)])
    carry_forward(cid)
    left, profit, unbalanced, _ = pl_net_after(env, cid, K)
    assert left == {} and not unbalanced
    assert profit == pytest.approx(8080.0)


@pytest.mark.gui
def test_carry_forward_is_not_run_when_nothing_to_carry(env, std):
    cid = env.client(std)
    carry_forward(cid)
    conn = env.db.get_db(); n = conn.execute("select count(*) from vouchers where client_id=?", (cid,)).fetchone()[0]; conn.close()
    assert n == 0, "没有损益发生额时不应生成结转凭证"
