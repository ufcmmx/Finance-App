# WiseLedger 测试报告

- 时间：2026-11-11 13:23:14　耗时 37.2s　Python 3.12.13　PySide6 6.11.0　Qt 6.11.0
- 平台插件：offscreen　项目根：(上一级)

| 层级 | 覆盖 | 通过 | 失败 | 已知问题(xfail) | 跳过 | 结论 |
|---|---|---|---|---|---|---|
| L0 环境与兼容性金丝雀 | PySide6/Qt 版本固定、样式表×Qt 版本下金额框是否可用 | 34 | 2 | 0 | 0 | ❌ 失败 |
| L1 金额框交互 | 真人点击时序、输入法/全角/粘贴、选区丢失 | 37 | 21 | 1 | 0 | ❌ 失败 |
| L3 科目期初 | 录入弹窗汇总、保存重开、建账期规则、期间选择 | 8 | 6 | 0 | 0 | ❌ 失败 |
| L2 期末结转 | 收入借方余额/费用贷方余额、结平、本年利润 | 6 | 0 | 0 | 0 | ✅ 通过 |
| L2 财务报表 | 利润表、资产负债表、逐科目取数覆盖 | 138 | 0 | 7 | 0 | ✅ 通过（有已知问题） |
| L2/L3 凭证与余额表 | 余额表口径、导出一致、封账只读、按钮可见性 | 9 | 0 | 0 | 0 | ✅ 通过 |
| L4 系统管理 | 修改密码页、备份页布局、keyring、备份/恢复 | 10 | 1 | 0 | 0 | ❌ 失败 |
| L4 自动备份 | 补跑日期逻辑、端到端 | 13 | 0 | 0 | 0 | ✅ 通过 |
| L4 审计日志 | 不可删改 | 1 | 1 | 0 | 0 | ❌ 失败 |

**合计：通过 226　失败 31　已知问题 8　意外通过(XPASS) 30　跳过 0**

## 失败明细（前 30 条）

- `tests/test_00_env_canary.py::test_running_pyside6_matches_requirements`
  - /Users/rickxu/Finance/tests/test_00_env_canary.py:41: AssertionError: 当前运行的 PySide6 是 6.11.0，requirements.txt 固定的是 6.8.3——本地、CI、打包用的版本不一致
- `tests/test_00_env_canary.py::test_ci_workflow_pins_pyside6`
  - /Users/rickxu/Finance/tests/test_00_env_canary.py:54: AssertionError: CI 里安装 PySide6 没有固定版本（会随上游发布悄悄变化）：
- `tests/test_10_spinbox.py::test_zero_cell_cursor_anywhere_still_replaces[光标在最前]`
  - /Users/rickxu/Finance/tests/test_10_spinbox.py:63: assert 220.0 == 22 ± 2.2e-05
- `tests/test_10_spinbox.py::test_zero_cell_cursor_anywhere_still_replaces[光标在末尾]`
  - /Users/rickxu/Finance/tests/test_10_spinbox.py:63: assert 0.0 == 22 ± 2.2e-05
- `tests/test_10_spinbox.py::test_selection_lost_keyboard_still_replaces[原值10-光标front-键入22]`
  - /Users/rickxu/Finance/tests/test_10_spinbox.py:94: AssertionError: 得到 2210
- `tests/test_10_spinbox.py::test_selection_lost_keyboard_still_replaces[原值10-光标end-键入22]`
  - /Users/rickxu/Finance/tests/test_10_spinbox.py:94: AssertionError: 得到 10
- `tests/test_10_spinbox.py::test_selection_lost_keyboard_still_replaces[原值0-光标front-键入22]`
  - /Users/rickxu/Finance/tests/test_10_spinbox.py:94: AssertionError: 得到 220
- `tests/test_10_spinbox.py::test_selection_lost_keyboard_still_replaces[原值0-光标end-键入22]`
  - /Users/rickxu/Finance/tests/test_10_spinbox.py:94: AssertionError: 得到 0
- `tests/test_10_spinbox.py::test_selection_lost_keyboard_still_replaces[原值1500-光标end-键入0.5]`
  - /Users/rickxu/Finance/tests/test_10_spinbox.py:94: AssertionError: 得到 1500
- `tests/test_10_spinbox.py::test_selection_lost_keyboard_still_replaces[原值88-光标front-键入1.5]`
  - /Users/rickxu/Finance/tests/test_10_spinbox.py:94: AssertionError: 得到 1.58
- `tests/test_10_spinbox.py::test_input_method_commit_events[原值0-lostfront-提交22]`
  - /Users/rickxu/Finance/tests/test_10_spinbox.py:121: AssertionError: 得到 220，文本 '220.00'
- `tests/test_10_spinbox.py::test_input_method_commit_events[原值0-lostend-提交22]`
  - /Users/rickxu/Finance/tests/test_10_spinbox.py:121: AssertionError: 得到 0，文本 '0.00'
- `tests/test_10_spinbox.py::test_input_method_commit_events[原值10-lostfront-提交22]`
  - /Users/rickxu/Finance/tests/test_10_spinbox.py:121: AssertionError: 得到 2210，文本 '2210.00'
- `tests/test_10_spinbox.py::test_input_method_commit_events[原值10-lostend-提交22]`
  - /Users/rickxu/Finance/tests/test_10_spinbox.py:121: AssertionError: 得到 10，文本 '10.00'
- `tests/test_10_spinbox.py::test_input_method_is_disabled_on_amount_box`
  - /Users/rickxu/Finance/tests/test_10_spinbox.py:128: assert not True
- `tests/test_10_spinbox.py::test_paste_invalid_text_keeps_original_value[-100]`
  - /Users/rickxu/Finance/tests/test_10_spinbox.py:161: AssertionError: 粘贴 '-100' 后变成了 100.0
- `tests/test_10_spinbox.py::test_paste_invalid_text_keeps_original_value[abc]`
  - /Users/rickxu/Finance/tests/test_10_spinbox.py:161: AssertionError: 粘贴 'abc' 后变成了 0.0
- `tests/test_10_spinbox.py::test_paste_invalid_text_keeps_original_value[nan]`
  - /Users/rickxu/Finance/tests/test_10_spinbox.py:161: AssertionError: 粘贴 'nan' 后变成了 0.0
- `tests/test_10_spinbox.py::test_paste_invalid_text_keeps_original_value[inf]`
  - /Users/rickxu/Finance/tests/test_10_spinbox.py:161: AssertionError: 粘贴 'inf' 后变成了 0.0
- `tests/test_10_spinbox.py::test_paste_invalid_text_keeps_original_value[1_000]`
  - /Users/rickxu/Finance/tests/test_10_spinbox.py:161: AssertionError: 粘贴 '1_000' 后变成了 1000.0
- `tests/test_10_spinbox.py::test_paste_invalid_text_keeps_original_value[(500)]`
  - /Users/rickxu/Finance/tests/test_10_spinbox.py:161: AssertionError: 粘贴 '(500)' 后变成了 500.0
- `tests/test_10_spinbox.py::test_paste_invalid_text_keeps_original_value[12abc]`
  - /Users/rickxu/Finance/tests/test_10_spinbox.py:161: AssertionError: 粘贴 '12abc' 后变成了 12.0
- `tests/test_10_spinbox.py::test_paste_invalid_text_keeps_original_value[1.2.3]`
  - /Users/rickxu/Finance/tests/test_10_spinbox.py:161: AssertionError: 粘贴 '1.2.3' 后变成了 1.23
- `tests/test_20_opening.py::test_keyboard_entry_rolls_up_to_parents_and_balances[企业准则]`
  - /Users/rickxu/Finance/tests/test_20_opening.py:66: assert 0.0 == 1234.56 ± 0.00123456
- `tests/test_20_opening.py::test_keyboard_entry_rolls_up_to_parents_and_balances[小企业]`
  - /Users/rickxu/Finance/tests/test_20_opening.py:66: assert 0.0 == 1234.56 ± 0.00123456
- `tests/test_20_opening.py::test_save_persists_and_reopens_identically[企业准则]`
  - /Users/rickxu/Finance/tests/test_20_opening.py:83: KeyError: '1002.001.001'
- `tests/test_20_opening.py::test_save_persists_and_reopens_identically[小企业]`
  - /Users/rickxu/Finance/tests/test_20_opening.py:83: KeyError: '1002.001.001'
- `tests/test_20_opening.py::test_start_period_only_moves_forward_in_time[B 建账月早于创建月，应前移]`
  - /Users/rickxu/Finance/tests/test_20_opening.py:107: AssertionError: assert '2026-11' == '2026-03'
- `tests/test_20_opening.py::test_start_period_move_is_audited`
  - /Users/rickxu/Finance/tests/test_20_opening.py:115: AssertionError: 建账期前移必须写审计日志（含前后值）
- `tests/test_40_system.py::test_restore_replaces_db_keeps_previous_copy_and_wrong_password_is_audited`
  - /Users/rickxu/Finance/tests/test_40_system.py:144: AssertionError: 恢复前应把原库另存为 .pre_restore_bak

> ⚠️ **L0 金丝雀失败**：当前 Qt 版本与样式表不兼容，金额框内部输入区被压扁，界面上数字会看不见。先解决 L0（固定/回退 PySide6 版本，或修复样式表），其余层级的失败多半是它的连带结果。详见 TESTING.md 第 4 节。

## 意外通过（XPASS）——这些已知问题可能已被修复，请从 known_issues.py / xfail 标记中删除：

- `tests/test_10_spinbox.py::test_paste_enhancements[１２３．５-123.5]`
- `tests/test_10_spinbox.py::test_paste_enhancements[￥1,234.50-1234.5]`
- `tests/test_10_spinbox.py::test_paste_enhancements[¥500-500]`
- `tests/test_31_reports.py::test_balance_sheet_covers_account[企业-1321代理业务资产]`
- `tests/test_31_reports.py::test_balance_sheet_covers_account[企业-1481持有待售资产]`
- `tests/test_31_reports.py::test_balance_sheet_covers_account[企业-1482持有待售资产减值准备]`
- `tests/test_31_reports.py::test_balance_sheet_covers_account[企业-1512长期股权投资减值准备]`
- `tests/test_31_reports.py::test_balance_sheet_covers_account[企业-1531长期应收款]`
- `tests/test_31_reports.py::test_balance_sheet_covers_account[企业-1605工程物资]`
- `tests/test_31_reports.py::test_balance_sheet_covers_account[企业-1606固定资产清理]`
- `tests/test_31_reports.py::test_balance_sheet_covers_account[企业-1608在建工程减值准备]`
- `tests/test_31_reports.py::test_balance_sheet_covers_account[企业-1609工程物资减值准备]`
- `tests/test_31_reports.py::test_balance_sheet_covers_account[企业-1621生产性生物资产]`
- `tests/test_31_reports.py::test_balance_sheet_covers_account[企业-1622生产性生物资产累计折旧]`
- `tests/test_31_reports.py::test_balance_sheet_covers_account[企业-1631油气资产]`
- `tests/test_31_reports.py::test_balance_sheet_covers_account[企业-1632累计折耗]`
- `tests/test_31_reports.py::test_balance_sheet_covers_account[企业-1711商誉]`
- `tests/test_31_reports.py::test_balance_sheet_covers_account[企业-2245持有待售负债]`
- `tests/test_31_reports.py::test_balance_sheet_covers_account[企业-2314代理业务负债]`
- `tests/test_31_reports.py::test_balance_sheet_covers_account[企业-2315受托代销商品款]`
- `tests/test_31_reports.py::test_balance_sheet_covers_account[企业-2401递延收益]`
- `tests/test_31_reports.py::test_balance_sheet_covers_account[企业-2702未确认融资费用]`
- `tests/test_31_reports.py::test_balance_sheet_covers_account[企业-2711专项应付款]`
- `tests/test_31_reports.py::test_balance_sheet_covers_account[企业-4301专项储备]`
- `tests/test_31_reports.py::test_balance_sheet_covers_account[小企-1605工程物资]`
- `tests/test_31_reports.py::test_balance_sheet_covers_account[小企-1606固定资产清理]`
- `tests/test_31_reports.py::test_balance_sheet_covers_account[小企-2401递延收益]`
- `tests/test_31_reports.py::test_balance_sheet_covers_account[小企-2601长期应付款]`
- `tests/test_31_reports.py::test_balance_sheet_covers_account[小企-4301研发支出]`
- `tests/test_40_system.py::test_restore_leaves_audit_trail_in_the_live_database`