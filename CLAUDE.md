# 智一盈小账 · WiseLedger — 项目说明

## 技术栈
Python 3.12 / PySide6 / SQLite / PyInstaller

## Git 操作规范
- **git 提交、建分支等操作一律在 VS Code Source Control 面板完成**
- 每个功能新建独立分支，完成后 merge 到 main
- GitHub Actions 自动打包 Windows exe（push 到对应分支触发，产物在 Actions → Artifacts）

## 项目结构
- `main.py` — 程序入口（MainWindow + 登录流程）
- `pages/` — 8个业务页面
  - `client.py` — 客户管理
  - `account.py` — 科目管理
  - `opening.py` — 科目期初
  - `voucher.py` — 记账凭证
  - `settle.py` — 期末结账
  - `report.py` — 财务报表（含 _PrintDialog 一体化打印窗口）
  - `audit.py` — 审计日志
  - `system.py` — 系统管理（用户管理 / 修改密码 / 数据备份 / 账套导出）
- `dialogs/` — 弹窗
- `print_utils.py` — 报表 PDF 导出（reportlab，供 report.py 调用）
- `db.py` — 数据库（SQLite）
- `session.py` — 登录会话 + RBAC
- `backup_utils.py` / `auto_backup.py` / `kr_utils.py` — 备份相关

## 数据库表结构（db.py）
- `clients` — 账套基本信息（id, name, short_code, tax_id, client_type, accounting_std, contact, phone, email）
- `accounts` — 科目表（client_id, code, name, type, direction, parent_code, level, is_leaf, opening_debit, opening_credit）
- `vouchers` — 凭证头（client_id, period, voucher_no, date, preparer, status, note）
- `voucher_entries` — 凭证明细（voucher_id, line_no, summary, account_code, account_name, debit, credit）
- `periods` — 期间结账状态（client_id, period, is_closed）
- `audit_log` — 审计日志
- `users` — 用户（username, password_hash, role: superadmin/admin/accountant/readonly）
- `user_client_access` — 用户与账套的授权关系
- `settings` — 系统设置 key/value
- `bank_statements` — 银行对账单
- `aux_dimensions` / `aux_items` / `account_aux_config` / `voucher_entry_aux` — 辅助核算

## pages/system.py 结构
SystemPage 有 4 个 Tab（mapping 在 `_switch()` 方法）：
- Tab 0：用户管理（`_build_user_mgmt()`，仅 superadmin）
- Tab 1：修改密码（`_build_change_pw()`，所有用户）
- Tab 2：数据备份（`_build_backup()`，仅 superadmin）
- Tab 3：账套导出（`_build_export()`，admin 及以上）
  - 选择客户 + 期间范围（动态加载已有凭证期间）
  - 导出 ZIP，含：账套基本信息、科目表与期初余额、凭证汇总、科目余额表、财务报表
  - 记账凭证按用友样式导出，每期单独一个 xlsx（命名：记账凭证(YYYY年MM期).xlsx）

## pages/report.py 关键结构
- `_PrintDialog` 类（文件顶部）— 一体化打印窗口
- `ReportPage._render_pdf(tab_name, path)` — 统一 PDF 渲染入口
- `ReportPage._build_print_html(tab)` — 报表转打印用 HTML
- 下载按钮：`make_export_button()` → 弹出 Excel/PDF 选择菜单（utils.py 通用函数）
- 打印按钮：`_print_report()` → 打开 `_PrintDialog`

## utils.py 通用工具
- `make_export_button(actions, label)` — 带下拉菜单的导出按钮，report.py / voucher.py 共用
- 页面权限按钮需在 `refresh_after_login()` 里重新设置 setVisible，避免登录前构建时权限判断失效

## 代码规范
- PySide6 枚举使用完整命名空间：`Qt.AlignmentFlag.AlignCenter`，不用 `Qt.AlignCenter`
- 可空 dict 参数写法：`user: dict | None = None`
- Pylance 配置见 `pyrightconfig.json`
- 页面按需创建（lazy loading），避免 Windows 启动闪窗
