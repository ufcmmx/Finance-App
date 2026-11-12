"""L0 环境与兼容性「金丝雀」测试 — 每次发版前、每次升级 PySide6/Qt 后必须先过这一层。

背景：Qt 6.12.0 起，全局样式表里的 `qproperty-buttonSymbols: NoButtons` 与带边框的样式表组合，
会把 QDoubleSpinBox 内部输入区压成 1 像素宽——数字、光标都看不见，点击也落不到输入框，
表现为「格子里的 0.00 显示不出来、输入的数字被乘 10 且看不见」。
6.6 ~ 6.11 正常。这一层测试就是为了在打包之前发现这类「框架版本 × 样式」问题。
"""
import re, pathlib
import pytest
from PySide6.QtCore import Qt, qVersion
from conftest import ROOT, STYLES
from helpers import make_spin, set_style, spin_events, edit_of, APP


def _spin_classes():
    import utils
    from dialogs.client_dialogs import AccountOpeningSpinBox
    return [("NoScrollDoubleSpinBox", utils.NoScrollDoubleSpinBox),
            ("NoScrollSpinBox", utils.NoScrollSpinBox),
            ("AccountOpeningSpinBox", AccountOpeningSpinBox)]


# ── 依赖版本必须固定（这次事故的根源：CI 里 pip install PySide6 没有固定版本，装到了有缺陷的最新版）──
def test_requirements_pin_pyside6():
    f = ROOT / "requirements.txt"
    if not f.exists():
        pytest.skip("项目根目录没有 requirements.txt（建议新增并固定 PySide6 版本，见 TESTING.md 第 3 节）")
    txt = f.read_text(encoding="utf-8")
    m = re.search(r"^\s*PySide6(?:-Essentials|-Addons)?\s*==\s*([\d.]+)", txt, re.M | re.I)
    assert m, "requirements.txt 里的 PySide6 没有用 == 固定版本"


def test_running_pyside6_matches_requirements():
    import PySide6
    f = ROOT / "requirements.txt"
    if not f.exists():
        pytest.skip("没有 requirements.txt")
    m = re.search(r"^\s*PySide6\s*==\s*([\d.]+)", f.read_text(encoding="utf-8"), re.M | re.I)
    if not m:
        pytest.skip("requirements.txt 未固定 PySide6")
    assert PySide6.__version__.startswith(m.group(1)), \
        f"当前运行的 PySide6 是 {PySide6.__version__}，requirements.txt 固定的是 {m.group(1)}——本地、CI、打包用的版本不一致"


def test_ci_workflow_pins_pyside6():
    wf = ROOT / ".github" / "workflows"
    if not wf.exists():
        pytest.skip("没有 .github/workflows")
    bad = []
    for f in wf.glob("*.y*ml"):
        for i, line in enumerate(f.read_text(encoding="utf-8").splitlines(), 1):
            if re.search(r"pip install", line) and re.search(r"\bPySide6\b", line) and not re.search(r"PySide6[\w-]*==", line):
                bad.append(f"{f.name}:{i}: {line.strip()}")
    assert not bad, "CI 里安装 PySide6 没有固定版本（会随上游发布悄悄变化）：\n  " + "\n  ".join(bad)


# ── 核心金丝雀：带生产样式表时，金额框内部输入区不能被压扁 ──
@pytest.mark.gui
@pytest.mark.parametrize("style", STYLES)
@pytest.mark.parametrize("cls_name,cls", _spin_classes(), ids=[c[0] for c in _spin_classes()])
@pytest.mark.parametrize("width", [90, 140, 200])
def test_spinbox_edit_area_not_collapsed(style, cls_name, cls, width):
    set_style(style)
    for editable in ([True, False] if cls_name == "AccountOpeningSpinBox" else [True]):
        box_w = make_spin_noassert(cls, editable, width)
        assert box_w >= width * 0.6, (
            f"[Qt {qVersion()} / 样式 {style}] {cls_name}(editable={editable}) 控件宽 {width}px，内部输入区只有 {box_w}px。"
            "数字和光标会看不见。检查全局样式表里的 spinbox 规则（qproperty-buttonSymbols:NoButtons）。")


def make_spin_noassert(cls, editable, width):
    from PySide6.QtWidgets import QWidget, QVBoxLayout
    box = QWidget(); lay = QVBoxLayout(box)
    w = cls(editable=editable) if cls.__name__ == "AccountOpeningSpinBox" else cls()
    lay.addWidget(w); w.setValue(1234.5); w.setFixedSize(width, 32); box.resize(width + 40, 80); box.show(); spin_events()
    res = edit_of(w).geometry().width(); box.close(); return res


# ── 渲染检查：不只看宽度，直接看像素里有没有字 ──
@pytest.mark.gui
@pytest.mark.parametrize("style", STYLES)
@pytest.mark.parametrize("editable", [True, False], ids=["可编辑", "只读"])
def test_spinbox_text_is_actually_painted(style, editable):
    from dialogs.client_dialogs import AccountOpeningSpinBox
    set_style(style)
    w, le = make_spin(AccountOpeningSpinBox, pre=1234.5, editable=editable, size=(140, 32))
    img = w.grab().toImage()
    dark = 0
    for x in range(8, img.width() - 8):
        for y in range(6, img.height() - 6):
            c = img.pixelColor(x, y)
            if (c.red() + c.green() + c.blue()) / 3 < 190:
                dark += 1
    assert dark >= 15, f"[Qt {qVersion()} / 样式 {style}] 控件里几乎没有文字像素（{dark} 个）——数字没有画出来"
