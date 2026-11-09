---
name: wiseledger-testing
description: >-
  Comprehensive full-software regression and automated verification test runner for WiseLedger (智一盈小账).
  Use this skill whenever code modifications have been made to verify full software integrity,
  including Python compilation, password security, dual-standard chart of accounts seeding,
  opening balance rollup, voucher lifecycle & locking, period-end profit/loss carryforward,
  financial report articulation, AES-256-GCM backup/restore audit logging, PySide6 GUI input controls,
  Windows DirectWrite font compatibility, and RBAC permission isolation.
---

# WiseLedger 全功能综合自动化测试规程 (WiseLedger Testing Skill)

本 Skill 用于在每次完成代码开发、Bug 修复或功能重构后，协助进行**软件全功能、全流程、多维度**的自动化回归测试，确保交付的代码在 macOS 与 Windows 上均能稳定可靠运行，杜绝“改动此模块却破坏彼功能”的回归风险。

## 适用场景

- 每次修改代码后（包括 `main.py`、`pages/*.py`、`dialogs/*.py`、`utils.py`、`db.py`、`session.py`、`backup_utils.py` 等）。
- 提交 Git Commit 或打 Release Tag 之前。
- 用户询问“帮我做一次完整测试”或排查疑难兼容性问题时。

---

## 12 大核心功能测试矩阵

1. **全库语法与编译检查 (Syntax & Import Integrity)**
   - 遍历项目中所有 Python 模块执行 `compileall` 编译。
   - 杜绝语法错误、悬挂引用与坏模块导入。

2. **口令安全与哈希认证 (Password Security & Hash Auth)**
   - 验证口令哈希算法安全性及动态盐值（Salt）唯一性。
   - 验证正确密码认证通过，错误密码被严格拦截。

3. **双准则账套初始化及科目体系 (Dual-Standard Chart of Accounts)**
   - 验证新客户账套创建时，企业会计准则及小企业会计制度标准科目的全量初始化植入。
   - 验证资产、负债、所有者权益、成本、损益五大要素科目分类完整无缺失。

4. **期初余额录入与层级递归汇总 (Opening Balance Rollup)**
   - 验证子级科目（如 `1002.01`、`1002.02`）录入期初余额后，通过 `_rollup` 递归向上一级科目汇总。
   - 验证借贷方期初余额的汇总与试算平衡。

5. **记账凭证全生命周期与平衡校验 (Voucher Lifecycle & Validation)**
   - **借贷平衡**：验证借贷不平凭证无法保存与审核。
   - **状态流转**：验证凭证审核、反审核机制。
   - **安全封账拦截**：已封账期间禁止新增、修改、反审核或删除凭证。

6. **期末损益结转与结账封账机制 (Period-End Closing & Carryforward)**
   - 验证收入类（6001/5001等）与费用类（6602/5602等）损益科目全自动结转至 `本年利润`。
   - 验证期间结账封账流程（`is_closed=1`）及反结账（`反结账`）状态回滚。

7. **财务报表勾稽关系验证 (Financial Report Articulation)**
   - 验证科目余额表三平（期初平、发生平、期末平）。
   - 验证资产负债表与利润表勾稽（净利润变动额与未分配利润变动对应）。

8. **双准则资产负债表全科目取数与平衡 (Balance Sheet Dual-Standard Balancing)**
   - **企业会计准则**：包含 `1605 工程物资`、`1606 固定资产清理`、`1711 商誉`、`2401 递延收益`、`4301 专项储备` 等全口径科目下，资产负债表 `资产总计 == 负债和所有者权益总计`，差额严格为 0。
   - **小企业会计制度**：包含 `2601 长期应付款`、`2701 预计负债`、`1501 长期股权投资` 等特殊映射科目下，资产负债表严格平衡。

9. **AES-256-GCM 备份还原与现行库审计留痕 (Backup Decryption & Active DB Audit Trail)**
   - 验证数据库使用 AES-256-GCM 加密导出与还原。
   - 验证错误口令解密防范拦截。
   - 验证解密替换时序后，现行活跃库能够成功写入并检索到 `数据恢复` 审计日志。

10. **GUI 控件与 Windows 兼容性 (GUI Controls & IME Compatibility)**
    - **高度防塌陷**：验证表格单元格内的输入框在 Windows 样式下高度 $\ge 35\text{px}$，避免文字被裁切。
    - **中文输入法自动纠偏**：中文句号（`。`）、间隔号（`·`）等毫秒级自动转为英文小数点 `.`。
    - **全角与千分位清洗**：全角数字（`１２３。４５`）与千分位逗号（`1,234.50`）正常识别，失焦自动格式化为两位小数（`123.45`）。

11. **字体引擎 DirectWrite 兼容性 (Font Engine DirectWrite Compatibility)**
    - 验证应用全局默认字体已显式配置为现代矢量字体（`Microsoft YaHei`, `PingFang SC`, `Segoe UI`, `sans-serif`），严禁回退至古代点阵字体 `MS Sans Serif`，避免 Windows DirectWrite 引发数秒卡顿。

12. **RBAC 权限体系与用户会话隔离 (RBAC & User Session Security)**
    - 验证超级管理员、管理员、会计员与只读模式的角色权限隔离矩阵。

---

## 执行测试命令

使用项目专用的虚拟解释器运行测试套件：

```bash
/Users/rickxu/finance/.venv/bin/python /Users/rickxu/finance/tests/run_tests.py
```

或使用本 Skill 提供的辅助脚本：

```bash
bash .agents/skills/wiseledger-testing/scripts/run_test_suite.sh
```

---

## 结果判定与反馈规范

测试完成后，应向用户输出清晰的结构化测试矩阵：

| 测试模块 | 覆盖功能点 | 状态 | 说明 |
|---|---|---|---|
| Phase 1: 语法编译 | 全库 Python 代码编译 | ✅ 通过 | 0 语法错误 |
| Phase 2: 口令安全 | Argon2id/PBKDF2 哈希与认证 | ✅ 通过 | 动态盐值唯一，错误拦截有效 |
| Phase 3: 账套与科目 | 双准则科目体系全量预置 | ✅ 通过 | 资产/负债/权益/成本/损益完整覆盖 |
| Phase 4: 期初与汇总 | 递归父级汇总与借贷校验 | ✅ 通过 | 子级科目变动向上自动联动 |
| Phase 5: 凭证生命周期 | 借贷平衡、审核、封账拦截 | ✅ 通过 | 不平凭证拦截，封账保护生效 |
| Phase 6: 损益结转与封账 | 自动结转至本年利润及封账 | ✅ 通过 | 结转凭证生成正确，反结账正常 |
| Phase 7: 报表勾稽 | 科目余额表三平与报表勾稽 | ✅ 通过 | 资产总计与权益总计差额 ¥0.00 |
| Phase 8: 双准则资产负债表 | 25个企业科目 + 8个小企业科目 | ✅ 通过 | 企业与小企业全口径科目双向平衡 |
| Phase 9: 备份加密与恢复审计 | AES-256-GCM 加密与活跃库审计日志 | ✅ 通过 | 现行库留痕确证，密码错误拦截 |
| Phase 10: GUI 控件与输入法 | 高度防塌陷、中文标点/全角转换 | ✅ 通过 | 高度 $\ge 35\text{px}$，中文句号转点 |
| Phase 11: 字体引擎配置 | 现代矢量字体链配置 | ✅ 通过 | 杜绝 MS Sans Serif 点阵字体卡顿 |
| Phase 12: RBAC 会话安全 | 超级管理员/管理员/会计/只读模式 | ✅ 通过 | 权限矩阵隔离无越权风险 |
