"""L2 财务报表 —— 利润表、资产负债表（两套制度）+ 资产负债表取数覆盖性。"""
import pytest
from helpers import (CODES, PERIOD, carry_forward, income_rows, balance_sheet_rows, bs_totals, find_row, num)

EXPECT = {"营业收入": 10000.0, "营业利润": 3650.0, "利润总额": 3800.0, "净利润": 3400.0}   # 见 Env.standard_month


def col(rows, label):
    r = find_row(rows, label)
    return (num(r[2]) or 0.0) if r else None


@pytest.mark.gui
def test_income_statement_numbers_before_and_after_carry_forward(env, std):
    """利润总额 = 营业利润 + 营业外收入 − 营业外支出（曾被写成加）；结转凭证不得冲抵利润表。"""
    cid = env.client(std); env.standard_month(cid, std)
    before = {k: col(income_rows(cid), k) for k in EXPECT}
    carry_forward(cid)
    after = {k: col(income_rows(cid), k) for k in EXPECT}
    for k, v in EXPECT.items():
        assert before[k] == pytest.approx(v), f"结转前 {k}={before[k]}，应为 {v}"
        assert after[k] == pytest.approx(v), f"结转后 {k}={after[k]}，应为 {v}（结转凭证被计入了利润表？）"


@pytest.mark.gui
def test_income_statement_layout_per_standard(env, std):
    cid = env.client(std); env.standard_month(cid, std)
    rows = income_rows(cid); labels = "".join(r[0].replace("\n", "") for r in rows)
    if std == "小企业会计制度":
        assert len(rows) == 13, f"小企业利润表应为 13 行，实际 {len(rows)}"
        for bad in ("研发费用", "其他综合收益", "每股收益", "资产减值损失", "公允价值变动"):
            assert bad not in labels, f"小企业利润表不应出现「{bad}」"
        assert "营业税金及附加" in labels and "投资收益" in labels
        for must in ("销售费用", "管理费用", "财务费用"):
            assert col(rows, must) not in (None, 0.0), f"小企业利润表「{must}」没有取到数"
    else:
        assert len(rows) == 37, f"企业准则利润表应为 37 行，实际 {len(rows)}"


@pytest.mark.gui
def test_balance_sheet_balances_with_trading_assets(env, std):
    """1101 = 小企业「短期投资」/ 企业「交易性金融资产」，必须被取数并计入合计，否则报表不平。"""
    cid = env.client(std); env.standard_month(cid, std); carry_forward(cid)
    rows = balance_sheet_rows(cid)
    a, le = bs_totals(rows)
    assert a == pytest.approx(le), f"资产总计 {a:,.2f} ≠ 负债和所有者权益总计 {le:,.2f}"
    if std == "小企业会计制度":
        r = find_row(rows, "短期投资")
        assert r is not None and any(num(c) == 5000.0 for c in r), f"小企业资产负债表「短期投资」行：{r}"
        assert find_row(rows, "衍生金融资产") is None, "小企业无「衍生金融资产」"
    else:
        r = find_row(rows, "以公允价值计量且其变动计入当期损益的金融资产")
        assert r is not None and any(num(c) == 5000.0 for c in r), f"企业准则行次 2：{r}"


# ───────────── 资产负债表取数覆盖性：每个一级资产/负债/权益科目有余额时，报表必须仍然平衡 ─────────────
# 做法：对每个一级科目取其第一个末级科目，记一张借贷平衡的凭证（对手科目为银行存款），报表不平 = 该科目没被取数。
from known_issues import BS_GAPS


def _level1_leaf_accounts(std):
    import db
    rows = db.STANDARD_ACCOUNTS if std == "企业会计准则" else db.STANDARD_ACCOUNTS_SMALL
    codes = {r[0] for r in rows}
    leaf = lambda c: not any(o != c and o.startswith(c + ".") for o in codes)
    out = []
    for r in rows:
        code, name, direction = r[0], r[1], r[4]
        if "." in code or code[0] not in "1234":
            continue
        cand = sorted(x[0] for x in rows if (x[0] == code or x[0].startswith(code + ".")) and leaf(x[0]))
        if cand:
            out.append((code, name, direction, cand[0]))
    return out


def _coverage_params():
    params = []
    for std in ("企业会计准则", "小企业会计制度"):
        for code, name, direction, leaf_code in _level1_leaf_accounts(std):
            marks = []
            if (std, code) in BS_GAPS:
                marks.append(pytest.mark.xfail(reason=f"已知缺口：{BS_GAPS[(std, code)]}", strict=False))
            params.append(pytest.param(std, code, name, direction, leaf_code, marks=marks, id=f"{std[:2]}-{code}{name}"))
    return params


@pytest.mark.gui
@pytest.mark.parametrize("std,code,name,direction,leaf_code", _coverage_params())
def test_balance_sheet_covers_account(env, std, code, name, direction, leaf_code):
    cid = env.client(std)
    plug = "1001" if code == "1002" else "1002"
    entries = [(leaf_code, 100, 0), (plug, 0, 100)] if direction == "借" else [(plug, 100, 0), (leaf_code, 0, 100)]
    env.post(cid, PERIOD, entries, to_leaf=False)
    a, le = bs_totals(balance_sheet_rows(cid))
    assert a is not None and abs(a - le) < 0.005, f"{code} {name} 有余额时资产负债表不平：资产 {a} / 负债+权益 {le}（该科目没有被取数）"
