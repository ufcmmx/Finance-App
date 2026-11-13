#!/usr/bin/env python3
"""tests/compare_runs.py — 对比两台机器（或两次运行）的测试结果

用法：  python tests/compare_runs.py results_Darwin.json results_Windows.json

先回答「两边用的是不是同一份代码」，再对比测试结果。代码不同时，结果不能直接比较。
"""
import json, sys


def load(p):
    d = json.load(open(p, encoding="utf-8"))
    return d, {r["id"]: r for r in d["rows"]}


def main():
    if len(sys.argv) != 3:
        print(__doc__); return 2
    (a, ra), (b, rb) = load(sys.argv[1]), load(sys.argv[2])
    ea, eb = a.get("env") or {}, b.get("env") or {}
    print(f"{'':<14}{'A: ' + sys.argv[1]:<44}B: {sys.argv[2]}")
    row = lambda k, fa, fb: print(f"{k:<14}{str(fa):<44}{fb}")
    row("系统", ea.get("system", "?"), eb.get("system", "?"))
    row("PySide6 / Qt", f"{a.get('pyside6')} / {a.get('qt')}", f"{b.get('pyside6')} / {b.get('qt')}")
    row("git 提交", ea.get("git_commit", "?"), eb.get("git_commit", "?"))
    row("未提交改动", len(ea.get("git_dirty", [])), len(eb.get("git_dirty", [])))
    fa, fb = ea.get("fingerprint", {}), eb.get("fingerprint", {})
    for g, label in (("app", "应用代码指纹"), ("tests", "测试脚本指纹"), ("ci", "CI 配置指纹")):
        same = "✅ 相同" if fa.get(g) and fa.get(g) == fb.get(g) else "❌ 不同"
        row(label, fa.get(g, "(旧版结果，无指纹)"), f"{fb.get(g, '(旧版结果，无指纹)')}   {same}")

    same_code = bool(fa) and fa == fb
    if not fa or not fb:
        print("\n⚠️ 至少有一份结果来自旧版运行器，没有代码指纹——无法确认代码是否相同。请两边都用最新的 run_all.py 重跑。")
    elif not same_code:
        fx, fy = ea.get("files", {}), eb.get("files", {})
        diff = sorted(k for k in set(fx) | set(fy) if fx.get(k) != fy.get(k))
        print(f"\n❌ 两边的代码不同，共 {len(diff)} 个文件不一致（A 缺失 = 只在 B 有；B 缺失 = 只在 A 有）：")
        for k in diff[:40]:
            print(f"   {k:<44} {'A缺失' if k not in fx else ('B缺失' if k not in fy else '内容不同')}")
        if len(diff) > 40: print(f"   …另有 {len(diff) - 40} 个")
        print("\n→ 先让两边的代码一致（git pull 到同一个提交），再比较测试结果。下面的结果对比仅供参考。")
    else:
        print("\n✅ 两边代码完全一致（换行符已统一比较）。下面的差异就是平台/环境本身造成的。")

    fail = lambda r: {i for i, x in r.items() if x["kind"] == "failed"}
    only_a, only_b, both = fail(ra) - fail(rb), fail(rb) - fail(ra), fail(ra) & fail(rb)
    print(f"\n失败：A {len(fail(ra))} 项　B {len(fail(rb))} 项　两边都失败 {len(both)} 项")
    for name, ids in (("只在 A 失败", only_a), ("只在 B 失败", only_b)):
        if ids:
            print(f"\n{name}（{len(ids)} 项）：")
            for i in sorted(ids)[:25]:
                print(f"  - {i.replace('tests/', '')}")
            if len(ids) > 25: print(f"  …另有 {len(ids) - 25} 项")
    return 0 if same_code else 1


if __name__ == "__main__":
    sys.exit(main())
