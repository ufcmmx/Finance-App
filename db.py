import sqlite3, os, sys

# 数据库存放在程序同目录（开发时在脚本旁，打包后在 exe 旁）
if getattr(sys, 'frozen', False):
    # PyInstaller 打包后，exe 所在目录
    _APP_DIR = os.path.dirname(sys.executable)
else:
    # 开发时，main.py / db.py 所在目录
    _APP_DIR = os.path.dirname(os.path.abspath(__file__))

DB_PATH = os.path.join(_APP_DIR, "accounting.db")

def get_db():
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    conn.execute("PRAGMA foreign_keys=ON")
    return conn

def init_db():
    conn = get_db(); c = conn.cursor()
    c.executescript("""
    CREATE TABLE IF NOT EXISTS clients (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        name TEXT NOT NULL,
        short_code TEXT,
        tax_id TEXT,
        client_type TEXT DEFAULT '小规模纳税人',
        contact TEXT, phone TEXT, email TEXT,
        created_at TEXT DEFAULT CURRENT_TIMESTAMP
    );
    CREATE TABLE IF NOT EXISTS accounts (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        client_id INTEGER NOT NULL,
        code TEXT NOT NULL,
        name TEXT NOT NULL,
        full_name TEXT NOT NULL,
        type TEXT NOT NULL,
        direction TEXT NOT NULL DEFAULT '借',
        parent_code TEXT,
        level INTEGER DEFAULT 1,
        is_leaf INTEGER DEFAULT 1,
        is_frozen INTEGER DEFAULT 0,
        opening_debit REAL DEFAULT 0,
        opening_credit REAL DEFAULT 0,
        UNIQUE(client_id, code),
        FOREIGN KEY(client_id) REFERENCES clients(id)
    );
    CREATE TABLE IF NOT EXISTS vouchers (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        client_id INTEGER NOT NULL,
        period TEXT NOT NULL,
        voucher_no TEXT NOT NULL,
        date TEXT NOT NULL,
        preparer TEXT DEFAULT '未来',
        attachment_count INTEGER DEFAULT 0,
        status TEXT DEFAULT '待审核',
        note TEXT,
        created_at TEXT DEFAULT CURRENT_TIMESTAMP,
        FOREIGN KEY(client_id) REFERENCES clients(id)
    );
    CREATE TABLE IF NOT EXISTS voucher_entries (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        voucher_id INTEGER NOT NULL,
        line_no INTEGER NOT NULL,
        summary TEXT,
        account_code TEXT,
        account_name TEXT,
        debit REAL DEFAULT 0,
        credit REAL DEFAULT 0,
        FOREIGN KEY(voucher_id) REFERENCES vouchers(id) ON DELETE CASCADE
    );
    CREATE TABLE IF NOT EXISTS periods (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        client_id INTEGER NOT NULL,
        period TEXT NOT NULL,
        is_closed INTEGER DEFAULT 0,
        closed_at TEXT,
        UNIQUE(client_id, period),
        FOREIGN KEY(client_id) REFERENCES clients(id)
    );
    CREATE TABLE IF NOT EXISTS audit_log (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        client_id INTEGER,
        operator TEXT DEFAULT '未来',
        action TEXT NOT NULL,
        target_type TEXT,
        target_id TEXT,
        detail TEXT,
        created_at TEXT DEFAULT CURRENT_TIMESTAMP
    );
    CREATE TABLE IF NOT EXISTS aux_dimensions (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        client_id INTEGER NOT NULL,
        name TEXT NOT NULL,
        code TEXT,
        sort_order INTEGER DEFAULT 0,
        UNIQUE(client_id, name),
        FOREIGN KEY(client_id) REFERENCES clients(id)
    );
    CREATE TABLE IF NOT EXISTS aux_items (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        client_id INTEGER NOT NULL,
        dimension_id INTEGER NOT NULL,
        name TEXT NOT NULL,
        code TEXT,
        contact TEXT,
        phone TEXT,
        is_active INTEGER DEFAULT 1,
        FOREIGN KEY(client_id) REFERENCES clients(id),
        FOREIGN KEY(dimension_id) REFERENCES aux_dimensions(id)
    );
    CREATE TABLE IF NOT EXISTS account_aux_config (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        client_id INTEGER NOT NULL,
        account_code TEXT NOT NULL,
        dimension_id INTEGER NOT NULL,
        is_required INTEGER DEFAULT 0,
        UNIQUE(client_id, account_code, dimension_id),
        FOREIGN KEY(client_id) REFERENCES clients(id),
        FOREIGN KEY(dimension_id) REFERENCES aux_dimensions(id)
    );
    CREATE TABLE IF NOT EXISTS voucher_entry_aux (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        entry_id INTEGER NOT NULL,
        dimension_id INTEGER NOT NULL,
        aux_item_id INTEGER,
        aux_item_name TEXT,
        FOREIGN KEY(entry_id) REFERENCES voucher_entries(id) ON DELETE CASCADE,
        FOREIGN KEY(dimension_id) REFERENCES aux_dimensions(id)
    );
    CREATE TABLE IF NOT EXISTS voucher_templates (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        client_id INTEGER,
        name TEXT NOT NULL,
        entries TEXT NOT NULL,
        created_at TEXT DEFAULT CURRENT_TIMESTAMP
    );
    """)
    conn.commit()
    # Run migrations
    _migrate_db(conn)
    conn.close()

def _migrate_db(conn):
    """Apply schema migrations"""
    c = conn.cursor()
    # Check if is_frozen column exists in accounts table
    c.execute("PRAGMA table_info(accounts)")
    columns = {row[1] for row in c.fetchall()}
    if 'is_frozen' not in columns:
        c.execute("ALTER TABLE accounts ADD COLUMN is_frozen INTEGER DEFAULT 0")
        conn.commit()

# 企业会计准则（2006）完整科目表
STANDARD_ACCOUNTS = [
    # ── 资产 ──
    ("1001","库存现金","库存现金","资产","借","",1),
    ("1002","银行存款","银行存款","资产","借","",1),
    ("1012","其他货币资金","其他货币资金","资产","借","",1),
    ("1101","交易性金融资产","交易性金融资产","资产","借","",1),
    ("1121","应收票据","应收票据","资产","借","",1),
    ("1122","应收账款","应收账款","资产","借","",1),
    ("1123","预付账款","预付账款","资产","借","",1),
    ("1131","应收股利","应收股利","资产","借","",1),
    ("1132","应收利息","应收利息","资产","借","",1),
    ("1221","其他应收款","其他应收款","资产","借","",1),
    ("1231","坏账准备","坏账准备","资产","贷","",1),
    ("1401","材料采购","材料采购","资产","借","",1),
    ("1402","在途物资","在途物资","资产","借","",1),
    ("1403","原材料","原材料","资产","借","",1),
    ("1404","材料成本差异","材料成本差异","资产","借","",1),
    ("1405","库存商品","库存商品","资产","借","",1),
    ("1408","委托加工物资","委托加工物资","资产","借","",1),
    ("1411","周转材料","周转材料","资产","借","",1),
    ("1461","待摊费用","待摊费用","资产","借","",1),
    ("1511","持有至到期投资","持有至到期投资","资产","借","",1),
    ("1521","可供出售金融资产","可供出售金融资产","资产","借","",1),
    ("1531","长期应收款","长期应收款","资产","借","",1),
    ("1601","固定资产","固定资产","资产","借","",1),
    ("1602","累计折旧","累计折旧","资产","贷","",1),
    ("1603","固定资产减值准备","固定资产减值准备","资产","贷","",1),
    ("1604","在建工程","在建工程","资产","借","",1),
    ("1605","工程物资","工程物资","资产","借","",1),
    ("1606","固定资产清理","固定资产清理","资产","借","",1),
    ("1701","无形资产","无形资产","资产","借","",1),
    ("1702","累计摊销","累计摊销","资产","贷","",1),
    ("1801","长期股权投资","长期股权投资","资产","借","",1),
    ("1811","长期债权投资","长期债权投资","资产","借","",1),
    ("1901","长期待摊费用","长期待摊费用","资产","借","",1),
    ("1911","递延所得税资产","递延所得税资产","资产","借","",1),
    # ── 负债 ──
    ("2001","短期借款","短期借款","负债","贷","",1),
    ("2002","存入保证金","存入保证金","负债","贷","",1),
    ("2101","交易性金融负债","交易性金融负债","负债","贷","",1),
    ("2201","应付票据","应付票据","负债","贷","",1),
    ("2202","应付账款","应付账款","负债","贷","",1),
    ("2203","预收账款","预收账款","负债","贷","",1),
    ("2211","应付职工薪酬","应付职工薪酬","负债","贷","",1),
    ("2221","应交税费","应交税费","负债","贷","",1),
    ("2231","应付利息","应付利息","负债","贷","",1),
    ("2232","应付股利","应付股利","负债","贷","",1),
    ("2241","其他应付款","其他应付款","负债","贷","",1),
    ("2401","递延收益","递延收益","负债","贷","",1),
    ("2501","长期借款","长期借款","负债","贷","",1),
    ("2502","应付债券","应付债券","负债","贷","",1),
    ("2601","长期应付款","长期应付款","负债","贷","",1),
    ("2701","预计负债","预计负债","负债","贷","",1),
    ("2901","递延所得税负债","递延所得税负债","负债","贷","",1),
    # ── 所有者权益 ──
    ("3001","实收资本","实收资本","所有者权益","贷","",1),
    ("3002","资本公积","资本公积","所有者权益","贷","",1),
    ("3101","盈余公积","盈余公积","所有者权益","贷","",1),
    ("3102","一般风险准备","一般风险准备","所有者权益","贷","",1),
    ("3103","本年利润","本年利润","所有者权益","贷","",1),
    ("3104","利润分配","利润分配","所有者权益","贷","",1),
    ("3201","库存股","库存股","所有者权益","借","",1),
    # ── 成本 ──
    ("4001","生产成本","生产成本","成本","借","",1),
    ("4101","制造费用","制造费用","成本","借","",1),
    ("4301","研发支出","研发支出","成本","借","",1),
    ("4401","工程施工","工程施工","成本","借","",1),
    ("4402","工程结算","工程结算","成本","贷","",1),
    ("4403","机械作业","机械作业","成本","借","",1),
    # ── 收入 ──
    ("5001","主营业务收入","主营业务收入","收入","贷","",1),
    ("5051","其他业务收入","其他业务收入","收入","贷","",1),
    ("5111","投资收益","投资收益","收入","贷","",1),
    ("5121","公允价值变动损益","公允价值变动损益","收入","贷","",1),
    ("5211","汇兑损益","汇兑损益","收入","贷","",1),
    ("5301","营业外收入","营业外收入","收入","贷","",1),
    # ── 费用 ──
    ("5401","主营业务成本","主营业务成本","费用","借","",1),
    ("5402","其他业务成本","其他业务成本","费用","借","",1),
    ("5403","税金及附加","税金及附加","费用","借","",1),
    ("5501","销售费用","销售费用","费用","借","",1),
    ("5502","管理费用","管理费用","费用","借","",1),
    ("5503","财务费用","财务费用","费用","借","",1),
    ("5601","营业外支出","营业外支出","费用","借","",1),
    ("5701","所得税费用","所得税费用","费用","借","",1),
    ("5801","以前年度损益调整","以前年度损益调整","费用","借","",1),
    # ── 6xxx 科目体系（企业会计准则新准则，与5xxx并存） ──
    # ── 所有者权益（4xxx，部分软件采用此编码） ──
    ("4001","实收资本","实收资本","所有者权益","贷","",1),
    ("4002","资本公积","资本公积","所有者权益","贷","",1),
    ("4101","盈余公积","盈余公积","所有者权益","贷","",1),
    ("4103","本年利润","本年利润","所有者权益","贷","",1),
    ("4104","利润分配","利润分配","所有者权益","贷","",1),
    ("4201","库存股","库存股","所有者权益","借","",1),
    # ── 6xxx 收入 ──
    ("6001","主营业务收入","主营业务收入","收入","贷","",1),
    ("6002","其他业务收入","其他业务收入","收入","贷","",1),
    ("6051","其他业务收入","其他业务收入","收入","贷","",1),
    ("6111","投资收益","投资收益","收入","贷","",1),
    ("6121","公允价值变动损益","公允价值变动损益","收入","贷","",1),
    ("6301","营业外收入","营业外收入","收入","贷","",1),
    # ── 6xxx 成本 ──
    ("6401","主营业务成本","主营业务成本","费用","借","",1),
    ("6402","其他业务成本","其他业务成本","费用","借","",1),
    ("6403","税金及附加","税金及附加","费用","借","",1),
    # ── 6xxx 费用 ──
    ("6601","销售费用","销售费用","费用","借","",1),
    ("6602","管理费用","管理费用","费用","借","",1),
    ("6603","财务费用","财务费用","费用","借","",1),
    ("6604","研发费用","研发费用","费用","借","",1),
    ("6711","营业外支出","营业外支出","费用","借","",1),
    ("6801","所得税费用","所得税费用","费用","借","",1),
    ("6901","以前年度损益调整","以前年度损益调整","费用","借","",1),
]

def log_action(conn, client_id, action, target_type="", target_id="", detail="", operator="未来"):
    """Write one audit log entry using an existing open connection."""
    conn.execute(
        "INSERT INTO audit_log(client_id,operator,action,target_type,target_id,detail) VALUES(?,?,?,?,?,?)",
        (client_id, operator, action, target_type, str(target_id), detail)
    )

VOUCHER_TEMPLATES = [
    ("计提工资", [("计提工资","5502","管理费用",0,0),("计提工资","2211","应付职工薪酬",0,0)]),
    ("发放工资", [("发放工资","2211","应付职工薪酬",0,0),("发放工资","1002","银行存款",0,0)]),
    ("现金存款", [("现金存入银行","1002","银行存款",0,0),("现金存入银行","1001","库存现金",0,0)]),
    ("银行取现", [("银行取现","1001","库存现金",0,0),("银行取现","1002","银行存款",0,0)]),
    ("采购付款", [("采购付款","1403","原材料",0,0),("采购付款","1002","银行存款",0,0)]),
    ("销售收款", [("收到货款","1002","银行存款",0,0),("收到货款","5001","主营业务收入",0,0)]),
    ("缴纳增值税", [("缴纳增值税","2221","应交税费",0,0),("缴纳增值税","1002","银行存款",0,0)]),
]

def seed_client_accounts(client_id, conn=None):
    """Insert standard accounts for a new client.
    Pass an existing open connection to avoid SQLite 'database is locked' errors.
    If no connection is given, a new one is opened and closed automatically.
    """
    own_conn = conn is None
    if own_conn:
        conn = get_db()
    c = conn.cursor()
    c.execute("SELECT COUNT(*) FROM accounts WHERE client_id=?", (client_id,))
    if c.fetchone()[0] > 0:
        if own_conn: conn.close()
        return
    for code, name, full, typ, direction, parent, level in STANDARD_ACCOUNTS:
        c.execute(
            "INSERT OR IGNORE INTO accounts(client_id,code,name,full_name,type,direction,parent_code,level) VALUES(?,?,?,?,?,?,?,?)",
            (client_id, code, name, full, typ, direction, parent or None, level)
        )
    if own_conn:
        conn.commit(); conn.close()
