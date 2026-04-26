"""pages/system.py — SystemPage — 系统管理（备份/恢复/关于）"""
from datetime import datetime
from PySide6.QtWidgets import *
from PySide6.QtCore import Qt, QDate, Signal, QTimer
from PySide6.QtGui import QColor, QFont, QBrush, QPalette

from db import get_db, log_action
from utils import lbl, sep, card, fmt_amt, NoScrollSpinBox, NoScrollDoubleSpinBox

# openpyxl imported lazily inside each export function

class SystemPage(QWidget):
    """系统管理：数据备份/恢复、关于"""

    def __init__(self):
        super().__init__()
        L = QVBoxLayout(self); L.setContentsMargins(40,32,40,32); L.setSpacing(24)
        L.addWidget(lbl("系统管理", bold=True, size=18))

        # ── 数据备份 ──
        grp1 = QFrame(); grp1.setObjectName("card")
        g1 = QVBoxLayout(grp1); g1.setContentsMargins(24,20,24,20); g1.setSpacing(12)
        g1.addWidget(lbl("数据备份", bold=True, size=14))
        g1.addWidget(lbl(
            f"数据库位置：{self._db_path()}\n"
            "备份会将当前数据库完整复制到你指定的位置，建议定期备份到云盘或移动硬盘。",
            color="#666"))
        row1 = QHBoxLayout(); row1.setSpacing(12)
        b_backup = QPushButton("📦 立即备份"); b_backup.setObjectName("btn_primary")
        b_backup.setFixedWidth(140); b_backup.clicked.connect(self._backup)
        row1.addWidget(b_backup); row1.addStretch()
        g1.addLayout(row1)
        self.backup_log = QLabel(""); self.backup_log.setStyleSheet("color:#52c41a;font-size:12px;")
        g1.addWidget(self.backup_log)
        L.addWidget(grp1)

        # ── 数据恢复 ──
        grp2 = QFrame(); grp2.setObjectName("card")
        g2 = QVBoxLayout(grp2); g2.setContentsMargins(24,20,24,20); g2.setSpacing(12)
        g2.addWidget(lbl("数据恢复", bold=True, size=14))
        warn = QLabel("⚠ 恢复将覆盖当前全部数据，操作不可撤销。请确保已备份当前数据！")
        warn.setStyleSheet("background:#fff7e6;color:#d46b08;border-radius:6px;padding:8px 12px;font-size:12px;")
        warn.setWordWrap(True); g2.addWidget(warn)
        row2 = QHBoxLayout(); row2.setSpacing(12)
        b_restore = QPushButton("📂 从备份文件恢复"); b_restore.setObjectName("btn_red")
        b_restore.setFixedWidth(180); b_restore.clicked.connect(self._restore)
        row2.addWidget(b_restore); row2.addStretch()
        g2.addLayout(row2)
        self.restore_log = QLabel(""); self.restore_log.setStyleSheet("color:#666;font-size:12px;")
        g2.addWidget(self.restore_log)
        L.addWidget(grp2)

        # ── 关于 ──
        grp3 = QFrame(); grp3.setObjectName("card")
        g3 = QVBoxLayout(grp3); g3.setContentsMargins(24,20,24,20); g3.setSpacing(6)
        g3.addWidget(lbl("关于 智一会计", bold=True, size=14))
        g3.addWidget(lbl("版本：1.0.0    企业会计准则（2006）    支持 macOS / Windows", color="#666"))
        g3.addWidget(lbl(f"数据目录：{self._db_path()}", color="#aaa", size=11))
        L.addWidget(grp3)
        L.addStretch()

    def _db_path(self):
        from db import DB_PATH
        return DB_PATH

    def _backup(self):
        import shutil, datetime
        default = f"智一会计备份_{datetime.datetime.now().strftime('%Y%m%d_%H%M%S')}.db"
        path, _ = QFileDialog.getSaveFileName(self, "选择备份位置", default, "数据库文件(*.db)")
        if not path: return
        try:
            # Use SQLite backup API for a safe online backup
            import sqlite3
            from db import DB_PATH
            src = sqlite3.connect(DB_PATH)
            dst = sqlite3.connect(path)
            src.backup(dst)
            src.close(); dst.close()
            self.backup_log.setText(f"✓ 备份成功：{path}")
            self.backup_log.setStyleSheet("color:#52c41a;font-size:12px;")
        except Exception as e:
            self.backup_log.setText(f"✗ 备份失败：{e}")
            self.backup_log.setStyleSheet("color:#ff4d4f;font-size:12px;")

    def _restore(self):
        path, _ = QFileDialog.getOpenFileName(self, "选择备份文件", "", "数据库文件(*.db)")
        if not path: return
        if QMessageBox.question(self, "确认恢复",
                "恢复将覆盖当前所有数据，此操作不可撤销！\n\n确定要从所选备份文件恢复吗？",
                QMessageBox.Yes | QMessageBox.No) != QMessageBox.Yes:
            return
        try:
            import sqlite3
            from db import DB_PATH
            # Validate: check it's a valid SQLite db
            test = sqlite3.connect(path)
            test.execute("SELECT name FROM sqlite_master LIMIT 1")
            test.close()
            # Backup current first
            import shutil, datetime
            auto_bak = DB_PATH + f".autobak_{datetime.datetime.now().strftime('%Y%m%d%H%M%S')}"
            shutil.copy2(DB_PATH, auto_bak)
            # Restore
            src = sqlite3.connect(path)
            dst = sqlite3.connect(DB_PATH)
            src.backup(dst)
            src.close(); dst.close()
            self.restore_log.setText(f"✓ 恢复成功！原数据已自动备份至：{auto_bak}\n请重启应用使数据生效。")
            self.restore_log.setStyleSheet("color:#52c41a;font-size:12px;")
            QMessageBox.information(self, "恢复成功",
                "数据恢复成功！\n请关闭并重新启动应用以加载恢复的数据。")
        except Exception as e:
            self.restore_log.setText(f"✗ 恢复失败：{e}")
            self.restore_log.setStyleSheet("color:#ff4d4f;font-size:12px;")


