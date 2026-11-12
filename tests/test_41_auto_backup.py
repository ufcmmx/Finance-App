"""L4 月末自动备份：日期补跑逻辑（含闰年、跨年）+ 端到端。只写临时目录，使用内存假 keyring。"""
import glob, os
from datetime import date
import pytest
import auto_backup as AB
import kr_utils
from conftest import _MemKeyring


def fake_today(monkeypatch, y, m, d):
    class FD(date):
        @classmethod
        def today(cls): return date(y, m, d)
    monkeypatch.setattr(AB, "date", FD)


def mark(env, ym):
    c = env.db.get_db(); env.db.log_action(c, 0, "自动备份", "system", ym, "t"); c.commit(); c.close()


CASES = [
    ("11月6日，无记录 → 补跑上月", (2026, 11, 6), [], "2026-10"),
    ("10月31日(月末最后一天) → 当月", (2026, 10, 31), [], "2026-10"),
    ("10月30日(非最后一天) → 上月", (2026, 10, 30), [], "2026-09"),
    ("上月已备份 → 继续补更早一个月", (2026, 11, 6), ["2026-10"], "2026-09"),
    ("近6个月全部已备份 → 无需备份", (2026, 11, 6), ["2026-10", "2026-09", "2026-08", "2026-07", "2026-06"], None),
    ("跨年：2027-01-15 → 补 2026-12", (2027, 1, 15), [], "2026-12"),
    ("闰年2月29日 → 当月", (2028, 2, 29), [], "2028-02"),
    ("闰年3月1日 → 补 2028-02", (2028, 3, 1), [], "2028-02"),
    ("平年2月28日 → 当月", (2027, 2, 28), [], "2027-02"),
]

@pytest.mark.parametrize("desc,today,marks,want", CASES, ids=[c[0] for c in CASES])
def test_target_month(env, monkeypatch, desc, today, marks, want):
    fake_today(monkeypatch, *today)
    for k in marks: mark(env, k)
    assert AB.get_target_month() == want


def _setup(env, enabled=True, pw=True, path=None, tmp=None):
    _MemKeyring.store.clear()          # 每次 setup 都从干净的凭据状态开始
    if pw: kr_utils.kr_set("Auto#Pass2026")
    env.db.set_setting("auto_backup_enabled", "1" if enabled else "0")
    p = path if path is not None else str(tmp)
    env.db.set_setting("auto_backup_path", p); return p


def test_runs_and_backfills_one_month_per_login(env, monkeypatch, tmp_path):
    fake_today(monkeypatch, 2026, 11, 6); p = _setup(env, tmp=tmp_path / "bk")
    AB.check_and_run(None)
    fs = glob.glob(p + "/*.zyac")
    assert len(fs) == 1 and "2026-10" in os.path.basename(fs[0])
    rec = env.audits("自动备份"); assert rec and rec[0][0] == 0 and rec[0][3] == "2026-10"
    AB.check_and_run(None)
    assert sorted(r[3] for r in env.audits("自动备份")) == ["2026-09", "2026-10"], "每次登录只补跑一个月份"


def test_disabled_or_no_password_is_silent(env, monkeypatch, tmp_path):
    fake_today(monkeypatch, 2026, 11, 6)
    p = _setup(env, enabled=False, tmp=tmp_path / "a"); AB.check_and_run(None)
    assert not glob.glob(p + "/*.zyac") and not env.audits("自动备份")
    p = _setup(env, pw=False, tmp=tmp_path / "b"); AB.check_and_run(None)
    assert not glob.glob(p + "/*.zyac") and not env.audits("自动备份")


def test_uncreatable_directory_logs_failure_without_raising(env, monkeypatch, tmp_path):
    fake_today(monkeypatch, 2026, 11, 6)
    blocker = tmp_path / "file_not_dir"; blocker.write_text("x")
    _setup(env, path=str(blocker / "sub"))       # 目录的上级是一个文件：无论 Windows/macOS 都不可创建
    AB.check_and_run(None)
    assert len(env.audits("自动备份失败")) == 1 and not env.audits("自动备份")


def test_empty_path_defaults_to_backups_folder_next_to_db(env, monkeypatch, tmp_path):
    fake_today(monkeypatch, 2026, 11, 6); _setup(env, path="")
    AB.check_and_run(None)
    assert len(glob.glob(os.path.join(os.path.dirname(env.db.DB_PATH), "backups", "*.zyac"))) == 1
