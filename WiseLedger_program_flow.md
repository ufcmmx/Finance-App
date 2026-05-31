# 智一盈小账程序流程图

本文档基于当前代码结构整理，重点覆盖程序启动、登录鉴权、账套选择、凭证处理、期末结账、报表与审计日志的数据流。

## 1. 程序主流程

```mermaid
flowchart TD
    A[启动 main.py] --> B[创建 QApplication 并设置样式/图标]
    B --> C[显示 Splash 启动画面]
    C --> D[init_db 初始化 SQLite 表结构/迁移/默认管理员]
    D --> E[创建 MainWindow]
    E --> F[构建侧边栏与 QStackedWidget 页面]
    F --> G[延迟导入并实例化业务页面]
    G --> H[显示主窗口]
    H --> I[弹出 LoginDialog]

    I --> J{用户名/密码有效?}
    J -- 否 --> K[记录登录失败审计日志并提示]
    K --> I
    J -- 是 --> L[写入 AppSession]
    L --> M[_refresh_for_login]

    M --> N[按角色权限显示导航菜单]
    N --> O[加载客户列表与系统状态]
    O --> P[检查并补跑月末自动备份]
    P --> Q[进入客户管理页]

    Q --> R{选择客户账套?}
    R -- 否 --> Q
    R -- 是 --> S[校验客户访问权限]
    S --> T[设置当前 client_id / 当前期间]
    T --> U[把客户上下文注入凭证/科目/期初/结账/报表/审计页面]
    U --> V[记录打开账套审计日志]
    V --> W[跳转到记账（凭证）页]
```

## 2. 财务业务闭环

```mermaid
flowchart TD
    A[客户管理] --> B{新建/导入账套}
    B --> C[创建客户信息]
    C --> D[初始化标准会计科目]
    D --> E[科目管理]
    E --> F[维护科目/辅助核算绑定]
    F --> G[科目期初]
    G --> H{是否为建账期或每年 1 月?}
    H -- 否 --> I[禁止录入期初]
    H -- 是 --> J[录入/修改末级科目期初余额]

    J --> K[记账（凭证）]
    F --> K
    K --> L[新增或编辑凭证]
    L --> M{期间是否已封账?}
    M -- 是 --> N[禁止新增/修改/删除/审核]
    M -- 否 --> O[校验分录]

    O --> P{校验通过?}
    P -- 否 --> Q[提示修正: 科目/辅助核算/借贷平衡]
    Q --> L
    P -- 是 --> R[写入 vouchers / voucher_entries / voucher_entry_aux]
    R --> S[记录新增或编辑凭证审计日志]
    S --> T[凭证状态: 待审核]

    T --> U{审核操作}
    U -- 拒绝 --> V[状态改为已拒绝并记日志]
    U -- 通过 --> W{借贷是否平衡?}
    W -- 否 --> Q
    W -- 是 --> X[状态改为已审核并记日志]

    X --> Y[财务报表]
    Y --> Z[仅汇总已审核凭证生成报表/导出 Excel]

    X --> AA[期末结账]
    AA --> AB{有待审核凭证?}
    AB -- 是 --> AC[阻止结转/封账]
    AB -- 否 --> AD[生成收入/费用结转凭证]
    AD --> AE[结转凭证状态为已审核]
    AE --> AF{封账检测通过?}
    AF -- 否 --> AC
    AF -- 是 --> AG[写入 periods.is_closed=1]
    AG --> AH[期间封账]

    AH --> AI{需要修改已封账期间?}
    AI -- 是 --> AJ[反结账: periods.is_closed=0]
    AJ --> K
    AI -- 否 --> AK[审计日志查询/导出]
```

## 3. 核心数据表关系

```mermaid
flowchart LR
    clients[clients 客户/账套] --> accounts[accounts 会计科目与期初]
    clients --> vouchers[vouchers 凭证主表]
    vouchers --> voucher_entries[voucher_entries 凭证分录]
    voucher_entries --> voucher_entry_aux[voucher_entry_aux 分录辅助核算]
    clients --> periods[periods 期间封账状态]
    clients --> aux_dimensions[aux_dimensions 辅助核算维度]
    aux_dimensions --> aux_items[aux_items 辅助核算对象]
    accounts --> account_aux_config[account_aux_config 科目辅助核算配置]
    users[users 用户] --> user_client_access[user_client_access 客户授权]
    users --> audit_log[audit_log 操作审计]
    clients --> audit_log
    settings[settings 系统设置] --> backup[手动/自动加密备份]
```

## 4. 代码依据

- `main.py`：`_NAV_ITEMS` 定义主菜单与权限；`MainWindow._build()` 创建页面栈；`_refresh_for_login()` 登录后刷新导航和自动备份；`_open_client()` 把客户上下文注入各业务页面。
- `login_dialog.py`：`LoginDialog._do_login()` 完成用户查询、密码校验、旧哈希迁移、登录日志和 `AppSession.login()`。
- `session.py`：`ROLE_PERMISSIONS` 和 `AppSession` 负责角色权限与客户访问控制。
- `db.py`：`init_db()` 创建核心表；`log_action()` 统一写审计日志；`seed_client_accounts()` 为账套初始化标准科目。
- `dialogs/voucher_dialogs.py`：`VoucherDialog._save()` 校验凭证分录、辅助核算、借贷平衡，并写入凭证表/分录表。
- `pages/voucher.py`：`VoucherPage` 处理凭证新增、编辑、删除、审核、期间封账限制与凭证查询。
- `pages/settle.py`：`SettlePage` 处理期末结转、封账检测、期间封账和反结账。
- `pages/report.py`：报表查询只汇总 `status='已审核'` 的凭证数据。
- `pages/audit.py`：审计日志页面按客户、日期和操作类型查询/导出关键操作。

