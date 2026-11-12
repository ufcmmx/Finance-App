---
name: wiseledger-regression
description: >-
  Run and interpret the WiseLedger (智一盈小账) automated regression suite (pytest, PySide6, SQLite).
  Use whenever the user changed code and wants it verified, is about to commit / tag / release,
  asks for a "完整测试 / 回归测试 / 发版前检查", reports a UI problem that only appears on Windows
  (digits invisible, input multiplied by 10, decimal point ignored, wrong report totals, carry-forward
  or balance-sheet mismatches), or upgrades PySide6/Qt. Also use to explain a failing test.
  The suite lives in tests/ of the repo; this skill runs it and reports ONLY what actually happened.
---

# WiseLedger 回归测试

## 怎么跑

1. 找到项目根目录（含 `tests/run_all.py` 的目录）。没有 pytest 就先 `pip install -r requirements.txt`。
2. 运行（任选其一，用脚本最省事）：
   - `python skills/wiseledger-regression/scripts/run_suite.py` — 自动定位项目根并运行
   - 加参数透传给 `tests/run_all.py`：`--fast`（跳过 139 项逐科目覆盖）、`--gui`（真实窗口系统，Windows 建议再跑一次）、`-k 关键字`
3. 读取 `tests/report/summary.md`，按下面的规范汇报。

## 汇报规范（最重要）

- **只汇报本次运行的真实输出**：每一层的通过/失败/已知问题/跳过数量，直接取自 summary.md。
  **严禁**使用预先写好的「全部 ✅」表格，严禁写没有运行过的项目。
- 有失败：列出失败用例 id + 一行原因，**先查代码，不要改测试让它变绿**。
  只有确认是测试本身写错时才能改测试，并且要说明为什么。
- **L0 金丝雀（test_00）红**：这是 Qt 版本 × 样式表不兼容（Qt 6.12.0 起金额框内部输入区被压成 1px，
  数字看不见、点击落不到输入框）。先解决 L0；其它层的失败多半是连带结果。
  处理：固定 PySide6 到 ≤ 6.11，或应用 `patches/utils_stylesheet_qt612.patch`。详见 TESTING.md 第 4 节。
- **XPASS（意外通过）**：说明某个已知问题可能已被修复——提醒用户把它从 `tests/known_issues.py` 或 xfail 标记里删除。
- **跳过项**：说明原因（例如沙盒里没有真实加密实现、没有 requirements.txt），不要当作通过。
- 自动化通过 ≠ 可以发布。结尾必须提醒：发布前还要完成 TESTING.md 第 6 节的 **Windows 真机手工清单**
  （真实输入法、凭据管理器、覆盖升级、安装包来源），这些自动化覆盖不到。

## 常见问题的定位

| 用户描述 | 先看 |
|---|---|
| 格子里 0.00 显示不出来 / 输入的数字看不见 / 乘 10 | L0 → 再看 PySide6 版本：`pip show PySide6` |
| Mac 正常、Windows 不正常 | 两台机器的 PySide6 版本是否一致（CI 是否固定了版本） |
| 重新发版后问题依旧 | 安装包是不是旧 commit 构建的（tag 没删/没移动）：对比 Actions 运行页的 commit 与 `git rev-parse` |
| 报表数字不对 | `test_31_reports`、`test_32_voucher`；用 `env.standard_month` 的期望值（营业利润 3650、利润总额 3800、净利润 3400）对照 |
| 结账后利润表全 0 | 结转凭证是否被计入报表口径：`test_30_settle` + `test_31_reports` |

## 不要做的事

- 不要在测试里重新实现被测逻辑（例如自己求和再断言），那样测不出任何东西。
- 不要关闭样式表、改用 Fusion 样式来「让测试通过」——这次事故正是样式表 × Qt 版本造成的。
- 不要把 `xfail` 当成通过；它是「已知缺陷」，要在报告里单独列出。
