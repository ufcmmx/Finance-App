"""L4 审计日志：不可删改（财会〔2024〕12号）。"""
import pytest
from PySide6.QtWidgets import QPushButton
from helpers import APP, delete_client, CODES, PERIOD


@pytest.mark.gui
def test_audit_page_has_no_clear_and_still_exports(env):
    import pages.audit as AU
    cid = env.client("企业会计准则")
    c = env.db.get_db(); env.db.log_action(c, cid, "新增凭证", "voucher", 1, "d"); c.commit(); c.close()
    pg = AU.AuditPage(); pg.set_client(cid); pg.show(); APP.processEvents()
    texts = [b.text() for b in pg.findChildren(QPushButton)]
    assert not any("清空" in t or "删除" in t for t in texts), f"审计日志页不应有清空/删除按钮：{texts}"
    assert not hasattr(pg, "_clear"), "后端删除入口 _clear 应已移除"
    assert any("导出" in t for t in texts) and pg.tbl.rowCount() >= 1


def test_no_code_path_deletes_audit_log():
    """静态检查：源码里不得出现对 audit_log 的 DELETE / DROP / UPDATE。"""
    import re
    from conftest import ROOT
    hits = []
    for f in list(ROOT.glob("*.py")) + list((ROOT / "pages").glob("*.py")) + list((ROOT / "dialogs").glob("*.py")):
        for i, line in enumerate(f.read_text(encoding="utf-8", errors="ignore").splitlines(), 1):
            if re.search(r"(delete\s+from|drop\s+table|update)\s+audit_log", line, re.I):
                hits.append(f"{f.name}:{i}: {line.strip()}")
    assert not hits, "发现修改/删除 audit_log 的语句：\n  " + "\n  ".join(hits)


@pytest.mark.gui
def test_deleting_a_client_keeps_its_audit_log_and_records_the_deletion(env, std):
    """删除客户会清掉账务数据，但绝不能清掉它的审计日志；并且要留下一条说明删了什么的日志。"""
    cid = env.client(std, name="待删客户")
    env.post(cid, PERIOD, [("1002", 10, 0), (CODES[std]["rev"], 0, 10)])
    c = env.db.get_db()
    for act in ("新增凭证", "审核凭证", "期末结转"):
        env.db.log_action(c, cid, act, "voucher", 1, "删除前产生的日志")
    c.commit()
    before = c.execute("select count(*) from audit_log where client_id=?", (cid,)).fetchone()[0]
    c.close()
    assert before >= 3
    delete_client(cid, "待删客户")
    c = env.db.get_db()
    gone_data = c.execute("select count(*) from vouchers where client_id=?", (cid,)).fetchone()[0]
    kept = c.execute("select count(*) from audit_log where client_id=?", (cid,)).fetchone()[0]
    rec = c.execute("select detail from audit_log where action='删除客户' and target_id=?", (str(cid),)).fetchall()
    c.close()
    assert gone_data == 0, "客户的凭证应已删除"
    assert kept >= before, f"删除客户后该客户的审计日志从 {before} 条变成了 {kept} 条——审计日志被删了"
    assert rec, "删除客户必须留下一条「删除客户」日志"
    assert "待删客户" in rec[0][0] and "凭证 1 张" in rec[0][0], f"删除日志应记录客户名和删除的凭证数：{rec[0][0]}"
