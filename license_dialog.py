"""license_dialog.py — 激活窗口（首次启动未激活时弹）"""
from __future__ import annotations
import os
import sys
from pathlib import Path

from PySide6.QtCore import Qt, QTimer
from PySide6.QtGui import QFont
from PySide6.QtWidgets import (
    QCheckBox, QDialog, QFrame, QHBoxLayout, QLabel, QLineEdit, QMessageBox,
    QPushButton, QTextBrowser, QVBoxLayout,
)

import license as lic


def _find_eula_path() -> Path | None:
    """定位 用户许可协议.md 文件；兼容开发环境和 PyInstaller 打包环境"""
    if hasattr(sys, "_MEIPASS"):
        # PyInstaller 打包运行时：资源解压到 sys._MEIPASS
        base = Path(sys._MEIPASS)
    else:
        base = Path(__file__).parent
    p = base / "用户许可协议.md"
    return p if p.exists() else None


class EulaViewerDialog(QDialog):
    """协议全文查看窗口"""
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("最终用户许可协议")
        self.resize(720, 560)
        self.setWindowFlags(
            Qt.Dialog | Qt.WindowTitleHint
            | Qt.WindowSystemMenuHint | Qt.WindowCloseButtonHint
        )
        L = QVBoxLayout(self)
        L.setContentsMargins(0, 0, 0, 0)
        L.setSpacing(0)

        browser = QTextBrowser()
        browser.setStyleSheet(
            "QTextBrowser { background:#fafbfc; border:none;"
            " padding:16px 24px; font-size:13px; line-height:1.6; }"
        )
        eula_path = _find_eula_path()
        if eula_path:
            with open(eula_path, "r", encoding="utf-8") as f:
                browser.setMarkdown(f.read())
        else:
            browser.setPlainText("协议文件缺失。请联系客服。")
        L.addWidget(browser, 1)

        # 底部关闭按钮 — 用 objectName 限定样式，避免污染子按钮
        btn_row = QFrame()
        btn_row.setObjectName("eulaFooter")
        btn_row.setStyleSheet(
            "#eulaFooter { background:#f0f2f5; border-top:1px solid #e4e8f0; }"
        )
        bl = QHBoxLayout(btn_row)
        bl.setContentsMargins(16, 12, 16, 12)
        bl.addStretch()
        btn_close = QPushButton("关闭")
        btn_close.setObjectName("btn_primary")
        btn_close.setFixedSize(80, 34)
        btn_close.clicked.connect(self.accept)
        bl.addWidget(btn_close)
        L.addWidget(btn_row)


class LicenseDialog(QDialog):
    """激活码输入对话框 — 风格跟 LoginDialog 一致"""

    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("智一盈小账 · WiseLedger")
        self.setFixedSize(440, 520)
        # 关闭按钮、标题栏（Windows 兼容）
        self.setWindowFlags(
            Qt.Dialog
            | Qt.WindowTitleHint
            | Qt.WindowSystemMenuHint
            | Qt.WindowCloseButtonHint
        )
        self._build()

    def closeEvent(self, event):
        """点 X 关闭 = 拒绝（主程序会退出）"""
        super().closeEvent(event)

    def _build(self):
        root = QVBoxLayout(self)
        root.setContentsMargins(0, 0, 0, 0)
        root.setSpacing(0)

        # ── 顶部深色标题栏 ──
        header = QFrame()
        header.setStyleSheet("background:#1c2340;")
        header.setFixedHeight(100)
        hl = QVBoxLayout(header)
        hl.setContentsMargins(32, 20, 32, 16)
        title = QLabel("智一盈小账 · WiseLedger")
        title.setStyleSheet("color:#ff8c00;font-size:18px;font-weight:bold;")
        sub = QLabel("请输入激活码以激活软件")
        sub.setStyleSheet("color:#8b93ae;font-size:12px;")
        hl.addWidget(title)
        hl.addWidget(sub)
        root.addWidget(header)

        # ── 表单区 ──
        body = QFrame()
        body.setStyleSheet("background:#f0f2f5;")
        bl = QVBoxLayout(body)
        bl.setContentsMargins(40, 24, 40, 24)
        bl.setSpacing(0)

        # 说明
        info = QLabel(
            "激活码格式：WL-XXXX-XXXX-XXXX\n"
            "每个激活码绑定一台电脑。"
        )
        info.setStyleSheet("color:#666;font-size:12px;padding:0 0 10px 0;")
        info.setWordWrap(True)
        bl.addWidget(info)

        # 字段标签
        lbl_code = QLabel("激活码")
        lbl_code.setStyleSheet("color:#555;font-size:13px;font-weight:bold;")
        bl.addWidget(lbl_code)
        bl.addSpacing(4)

        # 激活码输入
        self.f_code = QLineEdit()
        self.f_code.setPlaceholderText("WL-XXXX-XXXX-XXXX")
        self.f_code.setFixedHeight(34)
        f = QFont("Courier New, Menlo, Consolas, monospace")
        f.setPointSize(13)
        self.f_code.setFont(f)
        self.f_code.setStyleSheet("""
            QLineEdit {
                background:#fff; border:1px solid #bfc7d3;
                border-radius:6px; padding:2px 10px;
                letter-spacing:1px;
            }
            QLineEdit:focus { border:1.5px solid #3d6fdb; }
        """)
        self.f_code.returnPressed.connect(self._do_activate)
        # 自动大写转换
        self.f_code.textChanged.connect(self._normalize_input)
        bl.addWidget(self.f_code)
        bl.addSpacing(10)

        # 错误/状态提示（平时隐藏）
        self.status_lbl = QLabel("")
        self.status_lbl.setStyleSheet("color:#ff4d4f;font-size:12px;padding:4px 0;")
        self.status_lbl.setWordWrap(True)
        self.status_lbl.setVisible(False)
        bl.addWidget(self.status_lbl)

        bl.addStretch()

        # ── EULA 同意区（未勾选时激活按钮禁用） ──
        eula_row = QHBoxLayout(); eula_row.setSpacing(4)
        self.chk_eula = QCheckBox("我已阅读并同意")
        self.chk_eula.setStyleSheet("color:#555;font-size:12px;")
        self.chk_eula.toggled.connect(self._on_eula_toggled)
        self.btn_view_eula = QPushButton("《最终用户许可协议》")
        self.btn_view_eula.setCursor(Qt.CursorShape.PointingHandCursor)
        self.btn_view_eula.setStyleSheet("""
            QPushButton {
                background:transparent; border:none;
                color:#3d6fdb; padding:0; font-size:12px;
                text-decoration:underline;
            }
            QPushButton:hover { color:#2d5dc8; }
        """)
        self.btn_view_eula.clicked.connect(self._show_eula)
        eula_row.addWidget(self.chk_eula)
        eula_row.addWidget(self.btn_view_eula)
        eula_row.addStretch()
        bl.addLayout(eula_row)
        bl.addSpacing(10)

        # 按钮区（激活 + 取消）
        btn_row = QHBoxLayout()
        btn_row.setSpacing(10)

        self.btn_cancel = QPushButton("暂不激活")
        self.btn_cancel.setFixedHeight(38)
        self.btn_cancel.setStyleSheet("""
            QPushButton {
                background:#f5f5f5; color:#666;
                border:1px solid #d9d9d9; border-radius:6px;
                font-size:13px;
            }
            QPushButton:hover { background:#e8e8e8; }
        """)
        self.btn_cancel.clicked.connect(self.reject)

        self.btn_activate = QPushButton("激  活")
        self.btn_activate.setFixedHeight(38)
        self.btn_activate.setStyleSheet("""
            QPushButton {
                background:#3d6fdb; color:#fff; border:none;
                border-radius:6px; font-size:14px; font-weight:bold;
            }
            QPushButton:hover { background:#2d5dc8; }
        """)
        self.btn_activate.clicked.connect(self._do_activate)

        btn_row.addWidget(self.btn_cancel, 1)
        btn_row.addWidget(self.btn_activate, 2)
        bl.addLayout(btn_row)

        # 分隔线 + 免费试用按钮
        bl.addSpacing(14)
        sep_row = QHBoxLayout()
        line_l = QFrame(); line_l.setFrameShape(QFrame.HLine)
        line_l.setStyleSheet("background:#dcdfe6;")
        line_l.setFixedHeight(1)
        line_r = QFrame(); line_r.setFrameShape(QFrame.HLine)
        line_r.setStyleSheet("background:#dcdfe6;")
        line_r.setFixedHeight(1)
        or_lbl = QLabel("还没购买？")
        or_lbl.setStyleSheet("color:#888;font-size:11px;padding:0 8px;")
        sep_row.addWidget(line_l, 1); sep_row.addWidget(or_lbl); sep_row.addWidget(line_r, 1)
        bl.addLayout(sep_row)
        bl.addSpacing(8)

        self.btn_trial = QPushButton("🎁  免费试用 7 天")
        self.btn_trial.setFixedHeight(36)
        self.btn_trial.setStyleSheet("""
            QPushButton {
                background:transparent; color:#ff8c00;
                border:1.5px dashed #ff8c00; border-radius:6px;
                font-size:13px; font-weight:bold;
            }
            QPushButton:hover { background:#fff5e6; }
        """)
        self.btn_trial.clicked.connect(self._do_trial)
        bl.addWidget(self.btn_trial)

        # 客服提示
        hint = QLabel("如有问题，请联系客服微信：xxx-xxx-xxx")
        hint.setStyleSheet("color:#aaa;font-size:11px;padding-top:10px;")
        hint.setAlignment(Qt.AlignmentFlag.AlignCenter)
        bl.addWidget(hint)

        root.addWidget(body)

        # 焦点
        QTimer.singleShot(100, self.f_code.setFocus)

    def _normalize_input(self, text: str):
        """实时把输入大写化"""
        normalized = lic.normalize_license_code(text)
        if normalized != text:
            cursor_pos = self.f_code.cursorPosition()
            self.f_code.blockSignals(True)
            self.f_code.setText(normalized)
            self.f_code.setCursorPosition(cursor_pos)
            self.f_code.blockSignals(False)

    def _on_eula_toggled(self):
        """勾选协议后，如果之前有"请勾选"错误提示，隐藏它"""
        if self.chk_eula.isChecked() and self.status_lbl.isVisible():
            # 只在错误消息是"请先同意协议"时才隐藏（避免误消其他状态）
            if "同意" in self.status_lbl.text():
                self.status_lbl.setVisible(False)

    def _require_eula_or_prompt(self) -> bool:
        """检查是否已勾选协议；未勾选则提示 + 高亮 checkbox 后返回 False"""
        if self.chk_eula.isChecked():
            return True
        self._set_status("请先勾选下方【我已阅读并同意《最终用户许可协议》】")
        # 让 checkbox 短暂闪一下红色以吸引注意
        self.chk_eula.setStyleSheet(
            "color:#ff4d4f;font-size:12px;font-weight:bold;"
        )
        QTimer.singleShot(1500, lambda: self.chk_eula.setStyleSheet(
            "color:#555;font-size:12px;"
        ))
        return False

    def _show_eula(self):
        """弹出协议查看窗口"""
        dlg = EulaViewerDialog(self)
        dlg.exec()

    def _do_trial(self):
        """申请免费试用"""
        if not self._require_eula_or_prompt():
            return

        self.btn_trial.setEnabled(False)
        self.btn_trial.setText("申请中…")
        self.btn_trial.repaint()

        try:
            ok, msg = lic.request_trial()
        except Exception as e:
            ok, msg = False, f"未预期错误：{e}"

        self.btn_trial.setEnabled(True)
        self.btn_trial.setText("🎁  免费试用 7 天")

        if ok:
            self._set_status(msg + "（试用 7 天开始）", error=False)
            QTimer.singleShot(700, self.accept)
        else:
            self._set_status(msg)

    def _set_status(self, text: str, error: bool = True):
        color = "#ff4d4f" if error else "#52c41a"
        self.status_lbl.setStyleSheet(
            f"color:{color};font-size:12px;padding:4px 0;"
        )
        self.status_lbl.setText(("⚠ " if error else "✓ ") + text)
        self.status_lbl.setVisible(True)

    def _do_activate(self):
        if not self._require_eula_or_prompt():
            return
        code = self.f_code.text().strip()
        if not code:
            self._set_status("请输入激活码")
            return
        if not lic.is_valid_format(code):
            self._set_status("激活码格式错误（应为 WL-XXXX-XXXX-XXXX）")
            return

        self.btn_activate.setEnabled(False)
        self.btn_activate.setText("激活中…")
        self.status_lbl.setVisible(False)
        # 让 UI 立刻刷新
        self.btn_activate.repaint()

        try:
            ok, msg = lic.activate(code)
        except Exception as e:
            ok, msg = False, f"未预期错误：{e}"

        self.btn_activate.setEnabled(True)
        self.btn_activate.setText("激  活")

        if ok:
            self._set_status(msg, error=False)
            # 0.5 秒后自动关闭
            QTimer.singleShot(600, self.accept)
        else:
            self._set_status(msg)
            self.f_code.setFocus()
            self.f_code.selectAll()
