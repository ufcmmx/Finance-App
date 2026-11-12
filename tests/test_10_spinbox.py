"""L1 金额框（AccountOpeningSpinBox）交互测试 —— 带生产样式表、真人点击时序。

要点：所有「点击」都是 按下 → 事件循环转一圈 → 松开（QTest.mouseClick 把两者同步发出，会掩盖定时器相关的问题）。
"""
import pytest
from PySide6.QtCore import Qt, QPoint
from PySide6.QtTest import QTest
from helpers import (edit_of, make_spin, key, type_text, ime_commit, ime_type, paste_clipboard, press_release,
                     spin_events, APP)
from dialogs.client_dialogs import AccountOpeningSpinBox as Spin


def lose_selection(le, where):
    """模拟「点击后选区被平台折叠、光标移到某处」，且用户没有按任何导航键。"""
    le.deselect(); le.setCursorPosition(0 if where == "front" else len(le.text())); spin_events()


# ───────────── 半角输入：首次点击 → 输入 ─────────────
FIRST_CLICK = [
    (0, "left", "22", 22), (0, "right", "22", 22), (0, "mid", "10", 10), (0, "right", "1234.56", 1234.56),
    (0, "right", "0.5", 0.5), (0, "right", "0.05", 0.05), (0, "mid", "0.99", 0.99), (0, "right", ".5", 0.5),
    (22, "mid", "30", 30), (1500, "right", "30", 30), (1500, "left", "30", 30),
    (1500, "right", "0.5", 0.5), (100, "right", "0.3", 0.3), (9999.99, "mid", "1", 1),
]

@pytest.mark.gui
@pytest.mark.parametrize("pre,pos,text,want", FIRST_CLICK, ids=[f"原值{p:g}-点{pos}-输入{t}" for p, pos, t, _ in FIRST_CLICK])
def test_first_click_then_type_replaces_value(pre, pos, text, want):
    """回归：输入 22 不能变 220/2210，0.5 不能变 5，已有金额必须被整体覆盖。"""
    w, le = make_spin(Spin, pre, click=pos)
    type_text(le, text)
    assert w.value() == pytest.approx(want), f"得到 {w.value():g}，文本 {le.text()!r}"


@pytest.mark.gui
def test_click_selects_whole_text():
    w, le = make_spin(Spin, 1234.5, click="right")
    assert le.hasSelectedText() and le.selectedText() == le.text()


@pytest.mark.gui
def test_backspace_clears_then_decimal_typing_works():
    w, le = make_spin(Spin, 100, click="mid")
    QTest.keyClick(le, Qt.Key_Backspace); spin_events()
    type_text(le, "0.3")
    assert w.value() == pytest.approx(0.3)


@pytest.mark.gui
def test_delete_clears_to_zero():
    w, le = make_spin(Spin, 88, click="mid")
    QTest.keyClick(le, Qt.Key_Delete); spin_events()
    assert w.value() == 0


@pytest.mark.gui
@pytest.mark.parametrize("nav", [Qt.Key_Home, Qt.Key_End], ids=["光标在最前", "光标在末尾"])
def test_zero_cell_cursor_anywhere_still_replaces(nav):
    """原 bug 的两种形态：光标在末尾→输入被吞；光标在最前→22 变 220。"""
    w, le = make_spin(Spin, 0.0, focused_first=True)
    le.deselect(); QTest.keyClick(le, nav); spin_events()
    type_text(le, "22")
    assert w.value() == pytest.approx(22)


@pytest.mark.gui
def test_tab_into_nonzero_cell_then_type_replaces():
    from PySide6.QtWidgets import QWidget, QVBoxLayout
    box = QWidget(); lay = QVBoxLayout(box); a, b = Spin(editable=True), Spin(editable=True)
    a.setValue(5); b.setValue(1500)
    for s in (a, b): s.setFixedSize(140, 32); lay.addWidget(s)
    box.show(); spin_events(); a.setFocus(); spin_events()
    QTest.keyClick(edit_of(a), Qt.Key_Tab); spin_events()
    type_text(edit_of(b), "30")
    assert b.value() == pytest.approx(30)


@pytest.mark.gui
def test_readonly_cell_ignores_keys_and_paste():
    w, le = make_spin(Spin, 5.0, editable=False, click="right")
    type_text(le, "7"); paste_clipboard(le, "999")
    assert w.value() == 5.0 and le.text() == "5.00"


# ───────────── 选区丢失（有的平台在鼠标释放时折叠选区）：点击后的第一次输入仍应「覆盖」 ─────────────
LOST = [(10, "front", "22", 22), (10, "end", "22", 22), (0, "front", "22", 22), (0, "end", "22", 22),
        (1500, "end", "0.5", 0.5), (88, "front", "1.5", 1.5)]

@pytest.mark.gui
@pytest.mark.parametrize("pre,where,text,want", LOST, ids=[f"原值{p:g}-光标{w}-键入{t}" for p, w, t, _ in LOST])
def test_selection_lost_keyboard_still_replaces(pre, where, text, want):
    w, le = make_spin(Spin, pre, click="mid"); lose_selection(le, where)
    type_text(le, text)
    assert w.value() == pytest.approx(want), f"得到 {w.value():g}"


# ───────────── 全角 / 中文标点（按键路径）─────────────
FULLWIDTH = [("１２", 12), ("１２．５", 12.5), ("０．５", 0.5), ("１２3．5", 123.5), ("12。5", 12.5), ("1，234", 1234)]

@pytest.mark.gui
@pytest.mark.parametrize("text,want", FULLWIDTH)
def test_fullwidth_and_chinese_punctuation_keys(text, want):
    """中文标点状态下 '.' 键产生的是「。」；丢掉它会让 12.5 静默变成 125。"""
    w, le = make_spin(Spin, 0.0, click="right")
    type_text(le, text)
    assert w.value() == pytest.approx(want), f"得到 {w.value():g}"


# ───────────── 输入法提交事件（Windows 中文输入法下数字/标点可能以此方式到达）─────────────
IME = [(0, "click", "22", 22), (10, "click", "22", 22), (0, "click", "1.5", 1.5), (0, "click", "1。5", 1.5),
       (0, "click", "１２．５", 12.5), (0, "click", "0.05", 0.05),
       (0, "lostfront", "22", 22), (0, "lostend", "22", 22), (10, "lostfront", "22", 22), (10, "lostend", "22", 22)]

@pytest.mark.gui
@pytest.mark.parametrize("pre,mode,text,want", IME, ids=[f"原值{p}-{m}-提交{t}" for p, m, t, _ in IME])
def test_input_method_commit_events(pre, mode, text, want):
    w, le = make_spin(Spin, pre, click="mid")
    if mode.startswith("lost"):
        lose_selection(le, "front" if mode == "lostfront" else "end")
    ime_type(le, text)
    assert w.value() == pytest.approx(want), f"得到 {w.value():g}，文本 {le.text()!r}"


@pytest.mark.gui
def test_input_method_is_disabled_on_amount_box():
    """金额框不应启用输入法（Windows 中文输入法下 '.' 才不会被转成「。」）。"""
    w, le = make_spin(Spin, 0.0)
    assert not w.testAttribute(Qt.WA_InputMethodEnabled)
    assert not le.testAttribute(Qt.WA_InputMethodEnabled)


# ───────────── 粘贴 ─────────────
# 当前设计：Ctrl+V 走 Qt 默认粘贴——能识别的数字照常录入，不认识的内容直接忽略并保持原值（安全）。
PASTE_OK = [("1,234.50", 1234.5), (" 88 ", 88)]
# 增强项（非缺陷）：全角、带货币符号、超过两位小数的内容目前被忽略而不是被转换。想支持时去掉 xfail 即可。
PASTE_NICE = [("１２３．５", 123.5), ("￥1,234.50", 1234.5), ("¥500", 500), ("1234.567", 1234.57)]
# 必须被拒绝并保持原值的内容。（注：Qt 允许科学计数法 1e3，现实中几乎不会出现，不在此列）
PASTE_BAD = ["-100", "abc", "nan", "inf", "1_000", "(500)", "12abc", "1.2.3"]

@pytest.mark.gui
@pytest.mark.parametrize("text,want", PASTE_OK)
def test_paste_valid_amounts(text, want):
    w, le = make_spin(Spin, 7.0, click="right")
    paste_clipboard(le, text)
    assert w.value() == pytest.approx(want)

@pytest.mark.gui
@pytest.mark.xfail(reason="增强项：当前走 Qt 默认粘贴，全角/货币符号/超位小数会被忽略（安全，非缺陷）", strict=False)
@pytest.mark.parametrize("text,want", PASTE_NICE)
def test_paste_enhancements(text, want):
    w, le = make_spin(Spin, 7.0, click="right")
    paste_clipboard(le, text)
    assert w.value() == pytest.approx(want)

@pytest.mark.gui
@pytest.mark.parametrize("text", PASTE_BAD)
def test_paste_invalid_text_keeps_original_value(text):
    """无法识别的内容必须保持原值：nan/inf 不能变成 99 亿，负数不能静默变 0。"""
    w, le = make_spin(Spin, 7.0, click="right")
    paste_clipboard(le, text)
    assert w.value() == 7.0, f"粘贴 {text!r} 后变成了 {w.value()!r}"


# ───────────── 调试日志（排查 Windows 现场问题用）：默认不写文件 ─────────────
@pytest.mark.gui
def test_debug_log_off_by_default(tmp_path):
    import dialogs.client_dialogs as cd
    assert getattr(cd, "_INPUT_DEBUG", False) is False, "未设置 ZYAC_DEBUG_INPUT 时不应开启输入调试日志"
