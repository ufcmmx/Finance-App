"""L4 审计日志：不可删改（财会〔2024〕12号）。"""
import pytest
from PySide6.QtWidgets import QPushButton
from helpers import APP


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
