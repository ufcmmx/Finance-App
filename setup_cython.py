"""setup_cython.py — Cython 编译脚本

把 license.py 编译为原生扩展（Windows: .pyd，macOS/Linux: .so）。
反编译比 .pyc 困难数个量级，是防破解的核心步骤。

用法：
    # 本地测试编译（产生 .so）
    python setup_cython.py build_ext --inplace

    # CI 打包时：编译 + 自动删除源 .py（防止源码也被 PyInstaller 打进 exe）
    python setup_cython.py build_ext --inplace --remove-py

注意：
    1. Cython 编译产物是平台相关的（Mac 的 .so 不能在 Windows 跑）
    2. 本地编译产物不要 commit（.gitignore 已配置）
    3. license.py 的内容不会被脚本修改，只读取
"""
import os
import shutil
import sys
from pathlib import Path

# Windows CI 默认是 cp1252 编码，print 中文会崩溃；强制 UTF-8 输出
if hasattr(sys.stdout, "reconfigure"):
    try:
        sys.stdout.reconfigure(encoding="utf-8", errors="replace")
        sys.stderr.reconfigure(encoding="utf-8", errors="replace")
    except Exception:
        pass

# ─── 需要 Cython 编译的模块（只编译防破解核心）────
MODULES_TO_COMPILE = [
    "license.py",
    # 后续可以加 pw_utils.py / backup_utils.py 等
]

# ─── 参数解析（在 import setuptools 前处理掉自定义参数）────
REMOVE_PY_AFTER = "--remove-py" in sys.argv
if REMOVE_PY_AFTER:
    sys.argv.remove("--remove-py")


def main():
    try:
        from setuptools import setup
        from Cython.Build import cythonize
    except ImportError:
        print("❌ 缺 cython 或 setuptools，请先：")
        print("   pip install cython setuptools")
        sys.exit(1)

    here = Path(__file__).parent.resolve()
    src_files = [str(here / m) for m in MODULES_TO_COMPILE if (here / m).exists()]
    if not src_files:
        print("❌ 没找到任何要编译的 .py 文件")
        sys.exit(1)

    print(f"将编译以下模块: {src_files}")

    setup(
        ext_modules=cythonize(
            src_files,
            compiler_directives={
                "language_level": "3",
                "always_allow_keywords": True,
            },
        ),
    )

    # 编译完成后清理临时 .c 文件 + 可选删除源 .py
    for module in MODULES_TO_COMPILE:
        py_path = here / module
        c_path = here / module.replace(".py", ".c")
        if c_path.exists():
            c_path.unlink()
            print(f"清理临时文件: {c_path.name}")
        if REMOVE_PY_AFTER and py_path.exists():
            py_path.unlink()
            print(f"⚠️  已删除源文件: {py_path.name}（CI 打包模式）")

    # 清理 build/ 临时目录（保持工作区干净）
    build_dir = here / "build"
    if build_dir.exists():
        shutil.rmtree(build_dir)
        print(f"清理构建临时目录: build/")

    print("\n✓ Cython 编译完成")
    print("生成的扩展文件：")
    for f in here.glob("license*.so"):
        print(f"  {f.name}")
    for f in here.glob("license*.pyd"):
        print(f"  {f.name}")


if __name__ == "__main__":
    main()
