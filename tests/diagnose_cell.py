#!/usr/bin/env python3
"""tests/diagnose_cell.py — 诊断：期初录入弹窗里敲数字为什么没有生效

用法（在项目根目录）：  python tests/diagnose_cell.py
把完整输出发给我。它会打印：弹窗里金额格的真实类型、内部输入控件、点击与键入前后的状态。
"""
import os, sys, pathlib, tempfile
ROOT = pathlib.Path(os.environ.get("WL_PROJECT_ROOT") or pathlib.Path(__file__).resolve().parents[1])
sys.path.insert(0, str(ROOT)); sys.path.insert(0, str(ROOT / "tests"))
os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")
from PySide6.QtWidgets import QApplication
app = QApplication(sys.argv[:1])
import utils, db
app.setStyleSheet(utils.SS)
db.DB_PATH = os.path.join(tempfile.mkdtemp(), "diag.db"); db.init_db()
conn = db.get_db(); c = conn.cursor()
cols = [r[1] for r in c.execute("PRAGMA table_info(clients)")]
f = {"name": "诊断", "accounting_std": "企业会计准则"}
if "code" in cols: f["code"] = "diag"
if "short_code" in cols: f["short_code"] = "diag"
if "start_period" in cols: f["start_period"] = "2026-04"
c.execute(f"INSERT INTO clients({','.join(f)}) VALUES({','.join('?' * len(f))})", list(f.values())); cid = c.lastrowid
db.seed_client_accounts(cid, conn, "企业会计准则"); conn.commit(); conn.close()

from PySide6.QtCore import QPoint, Qt
from PySide6.QtTest import QTest
from dialogs.client_dialogs import AccountInitDialog
d = AccountInitDialog(None, cid, "2026-04"); d.show(); app.processEvents()
row = next(i for i in range(d.table.rowCount()) if d._leaf_flags[i])
code = d.table.item(row, 0).text(); w = d.table.cellWidget(row, 2)
print(f"首个末级科目: {code}（行 {row}）")
print("金额格类型:", type(w).__mro__[:4])
print("有 lineEdit()?", hasattr(w, "lineEdit"), "| 有 setText?", hasattr(w, "setText"), "| 只读?", getattr(w, "isReadOnly", lambda: "n/a")())
le = w.lineEdit() if hasattr(w, "lineEdit") else w
print(f"内部输入控件: {type(le).__name__}  宽{le.width()} 高{le.height()}  初始文本={le.text()!r}  value()={w.value() if hasattr(w,'value') else 'n/a'}")
w.setFocus(); app.processEvents()
for step, act in [("鼠标按下/松开", lambda: (QTest.mousePress(le, Qt.LeftButton, Qt.NoModifier, QPoint(le.width()-4, le.height()//2)),
                                          app.processEvents(), app.processEvents(),
                                          QTest.mouseRelease(le, Qt.LeftButton, Qt.NoModifier, QPoint(le.width()-4, le.height()//2)),
                                          app.processEvents(), app.processEvents())),
                  ("键入 1", lambda: (QTest.keyClicks(le, "1"), app.processEvents())),
                  ("键入 2", lambda: (QTest.keyClicks(le, "2"), app.processEvents())),
                  ("键入 .", lambda: (QTest.keyClicks(le, "."), app.processEvents())),
                  ("键入 5", lambda: (QTest.keyClicks(le, "5"), app.processEvents())),
                  ("Tab 离开(提交)", lambda: (QTest.keyClick(le, Qt.Key_Tab), app.processEvents()))]:
    act()
    print(f"  {step:<14} 文本={le.text()!r:<10} 选中={le.selectedText()!r:<8} 有焦点={le.hasFocus()!s:<5} value()={w.value() if hasattr(w,'value') else 'n/a'}")
print("\n结论提示：若「键入」后文本变化但 value() 始终为 0，说明数值只在失焦/提交时更新——测试需要在读取前提交；"
      "若文本本身没变化，说明按键没有到达输入控件。")
