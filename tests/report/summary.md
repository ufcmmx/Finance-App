# WiseLedger 测试报告

- 时间：2026-11-17 21:36:36　耗时 21.8s　Python 3.12.13　PySide6 6.11.0　Qt 6.11.0
- 平台插件：offscreen　项目根：/Users/rickxu/Finance
- 机器：Darwin / isssxmbp.local　git：feature/update-report @ 4d1e629　代码指纹：app `b20efa6604b0` / tests `de21e34820ea` / ci `88d1f331e911`
- ⚠️ **有 14 个未提交的改动**（ESTING.md, db.py, dialogs/client_dialogs.py, pages/opening.py, patches/client_keep_audit_log.patch…）——这份结果不对应任何一个提交，无法与另一台机器的结果严格比较。

| 层级 | 覆盖 | 通过 | 失败 | 已知问题(xfail) | 跳过 | 结论 |
|---|---|---|---|---|---|---|
| L0 环境与兼容性金丝雀 | PySide6/Qt 版本固定、样式表×Qt 版本下金额框是否可用 | 36 | 0 | 0 | 0 | ✅ 通过 |
| L1 金额框交互 | 真人点击时序、输入法/全角/粘贴、选区丢失 | 58 | 0 | 0 | 0 | ✅ 通过 |
| L3 科目期初 | 录入弹窗汇总、保存重开、建账期规则、期间选择 | 18 | 0 | 0 | 0 | ✅ 通过 |
| L2 期末结转 | 收入借方余额/费用贷方余额、结平、本年利润 | 6 | 0 | 0 | 0 | ✅ 通过 |
| L2 财务报表 | 利润表、资产负债表、逐科目取数覆盖 | 138 | 0 | 7 | 0 | ✅ 通过（有已知问题） |
| L2/L3 凭证与余额表 | 余额表口径、导出一致、封账只读、按钮可见性 | 9 | 0 | 0 | 0 | ✅ 通过 |
| L4 系统管理 | 修改密码页、备份页布局、keyring、备份/恢复 | 11 | 0 | 0 | 0 | ✅ 通过 |
| L4 自动备份 | 补跑日期逻辑、端到端 | 13 | 0 | 0 | 0 | ✅ 通过 |
| L4 审计日志 | 不可删改 | 4 | 0 | 0 | 0 | ✅ 通过 |

**合计：通过 269　失败 0　已知问题 7　意外通过(XPASS) 24　跳过 0**


## 意外通过（XPASS）——这些已知问题可能已被修复，请从 known_issues.py / xfail 标记中删除：

- `tests/test_10_spinbox.py::test_zero_cell_cursor_anywhere_still_replaces[光标在最前]`
- `tests/test_10_spinbox.py::test_zero_cell_cursor_anywhere_still_replaces[光标在末尾]`
- `tests/test_10_spinbox.py::test_selection_lost_keyboard_still_replaces[原值10-光标front-键入22]`
- `tests/test_10_spinbox.py::test_selection_lost_keyboard_still_replaces[原值10-光标end-键入22]`
- `tests/test_10_spinbox.py::test_selection_lost_keyboard_still_replaces[原值0-光标front-键入22]`
- `tests/test_10_spinbox.py::test_selection_lost_keyboard_still_replaces[原值0-光标end-键入22]`
- `tests/test_10_spinbox.py::test_selection_lost_keyboard_still_replaces[原值1500-光标end-键入0.5]`
- `tests/test_10_spinbox.py::test_selection_lost_keyboard_still_replaces[原值88-光标front-键入1.5]`
- `tests/test_10_spinbox.py::test_input_method_commit_events[原值0-lostfront-提交22]`
- `tests/test_10_spinbox.py::test_input_method_commit_events[原值0-lostend-提交22]`
- `tests/test_10_spinbox.py::test_input_method_commit_events[原值10-lostfront-提交22]`
- `tests/test_10_spinbox.py::test_input_method_commit_events[原值10-lostend-提交22]`
- `tests/test_10_spinbox.py::test_paste_enhancements[１２３．５-123.5]`
- `tests/test_10_spinbox.py::test_paste_enhancements[￥1,234.50-1234.5]`
- `tests/test_10_spinbox.py::test_paste_enhancements[¥500-500]`
- `tests/test_10_spinbox.py::test_paste_enhancements[1234.567-1234.57]`
- `tests/test_10_spinbox.py::test_paste_invalid_text_keeps_original_value[-100]`
- `tests/test_10_spinbox.py::test_paste_invalid_text_keeps_original_value[abc]`
- `tests/test_10_spinbox.py::test_paste_invalid_text_keeps_original_value[nan]`
- `tests/test_10_spinbox.py::test_paste_invalid_text_keeps_original_value[inf]`
- `tests/test_10_spinbox.py::test_paste_invalid_text_keeps_original_value[1_000]`
- `tests/test_10_spinbox.py::test_paste_invalid_text_keeps_original_value[(500)]`
- `tests/test_10_spinbox.py::test_paste_invalid_text_keeps_original_value[12abc]`
- `tests/test_10_spinbox.py::test_paste_invalid_text_keeps_original_value[1.2.3]`