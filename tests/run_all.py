#!/usr/bin/env python3
"""tests/run_all.py — 一键运行全部测试并输出中文报告（Windows / macOS / CI 通用）

用法：
  python tests/run_all.py                 # 默认：offscreen，全部测试
  python tests/run_all.py --fast          # 跳过 139 项资产负债表逐科目覆盖测试（约省 25 秒）
  python tests/run_all.py --gui           # 使用真实窗口系统（Windows 上建议再跑一遍，样式/DPI 更真实）
  python tests/run_all.py -k spinbox      # 只跑名称匹配的测试
  python tests/run_all.py --root D:\\wiseledger   # 指定项目根目录

退出码：0 = 无失败；1 = 有失败；2 = 环境问题（无法收集/运行）。
报告：tests/report/summary.md 和 tests/report/results.json（每次覆盖）。
"""
import argparse, json, os, pathlib, sys, time

HERE = pathlib.Path(__file__).resolve().parent
LAYERS = [
    ("test_00", "L0 环境与兼容性金丝雀", "PySide6/Qt 版本固定、样式表×Qt 版本下金额框是否可用"),
    ("test_10", "L1 金额框交互",         "真人点击时序、输入法/全角/粘贴、选区丢失"),
    ("test_20", "L3 科目期初",           "录入弹窗汇总、保存重开、建账期规则、期间选择"),
    ("test_30", "L2 期末结转",           "收入借方余额/费用贷方余额、结平、本年利润"),
    ("test_31", "L2 财务报表",           "利润表、资产负债表、逐科目取数覆盖"),
    ("test_32", "L2/L3 凭证与余额表",     "余额表口径、导出一致、封账只读、按钮可见性"),
    ("test_40", "L4 系统管理",           "修改密码页、备份页布局、keyring、备份/恢复"),
    ("test_41", "L4 自动备份",           "补跑日期逻辑、端到端"),
    ("test_42", "L4 审计日志",           "不可删改"),
]


class Collector:
    def __init__(self): self.rows = []
    def pytest_runtest_logreport(self, report):
        if report.when == "call" or (report.when == "setup" and report.outcome in ("failed", "skipped")):
            wasxfail = hasattr(report, "wasxfail")
            if report.outcome == "passed": kind = "xpass" if wasxfail else "passed"
            elif report.outcome == "failed": kind = "failed"
            else: kind = "xfailed" if wasxfail else "skipped"
            msg = ""
            if report.outcome == "failed" and report.longrepr is not None:
                lr = str(getattr(report.longrepr, "reprcrash", None) or report.longrepr)
                msg = lr.strip().splitlines()[-1][:240] if lr.strip() else ""
            elif kind in ("skipped", "xfailed"):
                msg = str(getattr(report, "wasxfail", "") or (report.longrepr[2] if isinstance(report.longrepr, tuple) else ""))[:160]
            self.rows.append({"id": report.nodeid, "kind": kind, "msg": msg, "sec": round(report.duration, 3)})


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--root"); ap.add_argument("--gui", action="store_true"); ap.add_argument("--fast", action="store_true")
    ap.add_argument("-k", dest="kexpr"); ap.add_argument("--styles")
    a = ap.parse_args()
    if a.root: os.environ["WL_PROJECT_ROOT"] = a.root
    if a.gui: os.environ["WL_TEST_GUI"] = "1"
    if a.styles: os.environ["WL_TEST_STYLES"] = a.styles
    try:
        import pytest
    except ImportError:
        print("缺少 pytest：请先 pip install pytest"); return 2
    args = [str(HERE), "-q", "-p", "no:cacheprovider", "--no-header", "-x" if False else "-q"]
    kx = [a.kexpr] if a.kexpr else []
    if a.fast: kx.append("not covers_account")
    if kx: args += ["-k", " and ".join(f"({k})" for k in kx)]
    col = Collector(); t0 = time.time()
    rc = pytest.main(args, plugins=[col])
    dur = time.time() - t0
    if rc in (2, 3, 4) and not col.rows:
        print("测试无法运行（收集阶段出错）。请检查 WL_PROJECT_ROOT 与依赖。"); return 2

    import PySide6
    from PySide6.QtCore import qVersion
    cnt = lambda rows, k: sum(1 for r in rows if r["kind"] == k)
    out = []
    out.append("# WiseLedger 测试报告\n")
    out.append(f"- 时间：{time.strftime('%Y-%m-%d %H:%M:%S')}　耗时 {dur:.1f}s　Python {sys.version.split()[0]}　PySide6 {PySide6.__version__}　Qt {qVersion()}")
    out.append(f"- 平台插件：{os.environ.get('QT_QPA_PLATFORM', '(系统默认，真实窗口)')}　项目根：{os.environ.get('WL_PROJECT_ROOT', '(上一级)')}\n")
    out.append("| 层级 | 覆盖 | 通过 | 失败 | 已知问题(xfail) | 跳过 | 结论 |\n|---|---|---|---|---|---|---|")
    for prefix, name, desc in LAYERS:
        rows = [r for r in col.rows if f"/{prefix}" in r["id"] or r["id"].startswith(f"tests/{prefix}") or f"{prefix}_" in r["id"]]
        if not rows: continue
        f = cnt(rows, "failed")
        verdict = "❌ 失败" if f else ("✅ 通过" + ("（有已知问题）" if cnt(rows, "xfailed") else ""))
        out.append(f"| {name} | {desc} | {cnt(rows,'passed')+cnt(rows,'xpass')} | {f} | {cnt(rows,'xfailed')} | {cnt(rows,'skipped')} | {verdict} |")
    total_f = cnt(col.rows, "failed")
    out.append(f"\n**合计：通过 {cnt(col.rows,'passed')}　失败 {total_f}　已知问题 {cnt(col.rows,'xfailed')}　意外通过(XPASS) {cnt(col.rows,'xpass')}　跳过 {cnt(col.rows,'skipped')}**\n")
    fails = [r for r in col.rows if r["kind"] == "failed"]
    if fails:
        out.append("## 失败明细（前 30 条）\n")
        for r in fails[:30]: out.append(f"- `{r['id']}`\n  - {r['msg']}")
        if any("test_00_env_canary" in r["id"] for r in fails):
            out.append("\n> ⚠️ **L0 金丝雀失败**：当前 Qt 版本与样式表不兼容，金额框内部输入区被压扁，界面上数字会看不见。"
                       "先解决 L0（固定/回退 PySide6 版本，或修复样式表），其余层级的失败多半是它的连带结果。详见 TESTING.md 第 4 节。")
    xp = [r for r in col.rows if r["kind"] == "xpass"]
    if xp:
        out.append("\n## 意外通过（XPASS）——这些已知问题可能已被修复，请从 known_issues.py / xfail 标记中删除：\n")
        for r in xp[:40]: out.append(f"- `{r['id']}`")
    sk = [r for r in col.rows if r["kind"] == "skipped"]
    if sk:
        out.append("\n## 跳过项（需人工或真实环境）\n")
        for r in sk[:15]: out.append(f"- `{r['id']}`：{r['msg']}")
    text = "\n".join(out)
    rep = HERE / "report"; rep.mkdir(exist_ok=True)
    (rep / "summary.md").write_text(text, encoding="utf-8")
    (rep / "results.json").write_text(json.dumps({"rows": col.rows, "failed": total_f, "qt": qVersion(), "pyside6": PySide6.__version__}, ensure_ascii=False, indent=1), encoding="utf-8")
    print("\n" + text)
    return 1 if total_f else 0


if __name__ == "__main__":
    sys.exit(main())
