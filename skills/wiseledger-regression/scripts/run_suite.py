#!/usr/bin/env python3
"""定位项目根目录并运行 tests/run_all.py，参数原样透传。Windows / macOS / CI 通用。"""
import os, pathlib, subprocess, sys

def find_root(start):
    p = pathlib.Path(start).resolve()
    for d in [p, *p.parents]:
        if (d / "tests" / "run_all.py").exists():
            return d
    return None

root = find_root(os.environ.get("WL_PROJECT_ROOT") or os.getcwd()) or find_root(__file__)
if not root:
    print("找不到 tests/run_all.py：请在项目根目录下运行，或设置 WL_PROJECT_ROOT"); sys.exit(2)
venv_py = next((str(root / p) for p in (".venv/bin/python", ".venv/Scripts/python.exe", "venv/bin/python", "venv/Scripts/python.exe")
                if (root / p).exists()), sys.executable)
env = dict(os.environ, WL_PROJECT_ROOT=str(root))
rc = subprocess.call([venv_py, str(root / "tests" / "run_all.py"), *sys.argv[1:]], cwd=str(root), env=env)
print(f"\n报告：{root / 'tests' / 'report' / 'summary.md'}")
sys.exit(rc)
