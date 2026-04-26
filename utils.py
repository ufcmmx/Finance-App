"""utils.py — 全局样式、UI辅助函数、控件子类、会计工具函数"""
from PySide6.QtWidgets import (QLabel, QFrame, QVBoxLayout, QSpinBox,
                                QDoubleSpinBox)
from PySide6.QtCore import Qt
from PySide6.QtGui import QColor, QFont

SS = """
* { font-family:'Microsoft YaHei','PingFang SC',sans-serif; font-size:13px; color:#1e2130; }
QMainWindow,QWidget#root { background:#f0f2f5; }
/* Sidebar */
QWidget#sidebar { background:#1c2340; }
QPushButton#nav { background:transparent; color:#8b93ae; border:none;
    text-align:left; padding:11px 20px; border-radius:0; }
QPushButton#nav:hover { background:#252d4a; color:#fff; }
QPushButton#nav[active=true] { background:#2d3760; color:#fff;
    border-left:3px solid #4e7df4; padding-left:17px; }
QLabel#logo { color:#fff; font-size:17px; font-weight:bold;
    padding:22px 20px 6px 20px; }
QLabel#subt { color:#4a5578; font-size:11px; padding:0 20px 14px 20px; }
/* Cards */
QFrame#card { background:#fff; border-radius:10px; border:1px solid #e4e8f0; }
QWidget#import_foot { background:#fff; border-top:1px solid #e4e8f0; }
/* Buttons */
QPushButton#btn_primary { background:#3d6fdb; color:#fff; border:none;
    border-radius:6px; padding:7px 18px; font-weight:bold; }
QPushButton#btn_primary:hover { background:#2d5dc8; }
QPushButton#btn_red { background:#ff4d4f; color:#fff; border:none;
    border-radius:6px; padding:7px 14px; }
QPushButton#btn_red:hover { background:#e63b3d; }
QPushButton#btn_green { background:#52c41a; color:#fff; border:none;
    border-radius:6px; padding:7px 14px; }
QPushButton#btn_outline { background:transparent; color:#3d6fdb;
    border:1px solid #3d6fdb; border-radius:6px; padding:7px 14px; }
QPushButton#btn_outline:hover { background:#eef3ff; }
QPushButton#btn_gray { background:#f5f5f5; color:#666; border:1px solid #d9d9d9;
    border-radius:6px; padding:7px 14px; }
QPushButton#btn_gray:hover { background:#e8e8e8; }
/* Inputs */
QLineEdit,QDateEdit,QComboBox,QDoubleSpinBox,QSpinBox,QTextEdit {
    background:#fff; border:1px solid #d9d9d9; border-radius:5px;
    padding:6px 10px; }
QSpinBox,QDoubleSpinBox { qproperty-buttonSymbols: NoButtons; }
QLineEdit:focus,QDateEdit:focus,QDoubleSpinBox:focus { border:1.5px solid #3d6fdb; }
QComboBox::drop-down { border:none; width:22px; }
/* Tables */
QTableWidget { background:#fff; border:none; gridline-color:#f0f2f5;
    selection-background-color:#e6f0ff; selection-color:#1e2130; }
QTableWidget::item { padding:8px 10px; }
QHeaderView::section { background:#fafafa; color:#8b93ae; border:none;
    border-bottom:1px solid #e8ecf2; padding:8px 10px;
    font-size:12px; font-weight:bold; }
/* Tabs */
QTabBar::tab { background:transparent; color:#8b93ae; padding:9px 18px;
    border:none; border-bottom:2px solid transparent; }
QTabBar::tab:selected { color:#3d6fdb; border-bottom:2px solid #3d6fdb; }
QTabWidget::pane { border:none; }
/* Scrollbar */
QScrollBar:vertical { width:5px; background:transparent; }
QScrollBar::handle:vertical { background:#dde1ea; border-radius:2px; min-height:24px; }
QScrollBar::add-line:vertical,QScrollBar::sub-line:vertical { height:0; }
QScrollBar:horizontal { height:5px; background:transparent; }
QScrollBar::handle:horizontal { background:#dde1ea; border-radius:2px; }
QScrollBar::add-line:horizontal,QScrollBar::sub-line:horizontal { width:0; }
"""

def lbl(text, bold=False, color=None, size=None):
    w = QLabel(text)
    st = ""
    if bold: st += "font-weight:bold;"
    if color: st += f"color:{color};"
    if size: st += f"font-size:{size}px;"
    if st: w.setStyleSheet(st)
    return w

def sep():
    f = QFrame(); f.setFrameShape(QFrame.HLine)
    f.setStyleSheet("color:#e8ecf2; margin:4px 0;"); return f

def card(widget=None):
    f = QFrame(); f.setObjectName("card")
    if widget:
        vl = QVBoxLayout(f); vl.setContentsMargins(0,0,0,0)
        vl.addWidget(widget)
    return f

class NoScrollSpinBox(QSpinBox):
    """QSpinBox that ignores mouse-wheel events to prevent accidental value changes."""
    def wheelEvent(self, event):
        event.ignore()

class NoScrollDoubleSpinBox(QDoubleSpinBox):
    """QDoubleSpinBox that ignores mouse-wheel events to prevent accidental value changes."""
    def wheelEvent(self, event):
        event.ignore()


def fmt_amt(v):
    if v == 0 or v is None: return ""
    return f"{v:,.2f}"

# Contra-asset account root codes (credit-normal)
_CONTRA_ASSET_ROOTS = {
    "1231","1471","1472","1482","1502","1512",
    "1602","1603","1608","1609","1622","1632","1702","1703"
}

def infer_account_type_direction(code, name=""):
    """Infer account type and normal-balance direction from code and name.
    Matches the standard chart of accounts template (4xxx=equity, 5xxx=cost, 6xxx=income/expense).
    """
    code = (code or "").strip()
    name = (name or "").strip()
    if not code:
        return "资产", "借"

    prefix = code[0]
    parts = code.split(".")
    root4 = code[:4] if len(code) >= 4 else code

    if prefix == "1":
        # Check contra-asset (credit normal): traverse ancestors
        for i in range(len(parts), 0, -1):
            ancestor = ".".join(parts[:i])
            if ancestor in _CONTRA_ASSET_ROOTS:
                return "资产", "贷"
        return "资产", "借"

    if prefix == "2":
        # 2702 未确认融资费用 is a contra-liability (debit normal)
        if code.startswith("2702"):
            return "负债", "借"
        return "负债", "贷"

    if prefix == "3":
        # In this template 3xxx are special clearing/hedging accounts (asset-like)
        return "资产", "借"

    if prefix == "4":
        if code.startswith("4201"):   # 库存股 — debit normal
            return "所有者权益", "借"
        return "所有者权益", "贷"

    if prefix == "5":
        if code.startswith("5402"):   # 工程结算 — credit normal
            return "成本", "贷"
        return "成本", "借"

    if prefix == "6":
        try:
            top4 = int(code[:4])
        except ValueError:
            top4 = 9999
        if top4 <= 6301:              # 6001-6301 are income/gain accounts
            return "收入", "贷"
        return "费用", "借"           # 6401+ are expense accounts

    return "资产", "借"

def cn_amount(n):
    units = ["", "拾", "佰", "仟", "万", "拾万", "佰万", "仟万", "亿"]
    digits = "零壹贰叁肆伍陆柒捌玖"
    if n == 0: return "零元整"
    sign = "负" if n < 0 else ""
    n = abs(round(n, 2))
    i_part = int(n); d_part = round((n - i_part) * 100)
    fen = digits[d_part % 10]; jiao = digits[d_part // 10]
    result = ""
    s = str(i_part); idx = 0
    for ch in reversed(s):
        d = int(ch)
        result = (digits[d] + units[idx] if d != 0 else "零") + result
        idx += 1
    result = result.rstrip("零") or "零"
    if d_part == 0: result += "元整"
    elif d_part % 10 == 0: result += f"元{jiao}角"
    else: result += f"元{jiao}角{fen}分"
    return sign + result

