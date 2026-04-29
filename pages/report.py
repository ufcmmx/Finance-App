"""pages/report.py — ReportPage — 财务报表"""
from datetime import datetime
from PySide6.QtWidgets import *
from PySide6.QtCore import Qt, QDate, Signal, QTimer
from PySide6.QtGui import QColor, QFont, QBrush, QPalette

from db import get_db, log_action
from utils import lbl, sep, card, fmt_amt, NoScrollSpinBox, NoScrollDoubleSpinBox

# openpyxl imported lazily inside each export function

class ReportPage(QWidget):
    """财务报表 — 资产负债表 + 利润表"""

    def __init__(self):
        super().__init__()
        self.client_id = None; self.period = ""
        L = QVBoxLayout(self); L.setContentsMargins(0,0,0,0); L.setSpacing(0)
        # Top tabs
        tb = QWidget(); tb.setStyleSheet("background:#fff;border-bottom:1px solid #e8ecf2;")
        tl = QHBoxLayout(tb); tl.setContentsMargins(16,0,16,0); tl.setSpacing(0)
        self._rtabs = []
        for n in ["资产负债表","利润表","所有者权益变动表","现金流量表","收支统计表"]:
            b = QPushButton(n); b.setStyleSheet("""QPushButton{background:transparent;color:#888;
                border:none;padding:12px 16px;border-bottom:2px solid transparent;}
                QPushButton:hover{color:#3d6fdb;}
                QPushButton[active=true]{color:#3d6fdb;border-bottom:2px solid #3d6fdb;}""")
            b.clicked.connect(lambda _,nn=n:self._switch(nn)); tl.addWidget(b); self._rtabs.append(b)
        tl.addStretch()
        # Period selector
        pr = QHBoxLayout(); pr.setSpacing(8)
        pr.addWidget(lbl("报告期间:", color="#666"))
        self.rep_start_period = QComboBox(); self.rep_start_period.setMinimumWidth(100)
        self.rep_end_period = QComboBox(); self.rep_end_period.setMinimumWidth(100)
        pr.addWidget(self.rep_start_period); pr.addWidget(lbl("至", color="#666")); pr.addWidget(self.rep_end_period)
        b_refresh = QPushButton("刷新"); b_refresh.setObjectName("btn_primary"); b_refresh.clicked.connect(self._refresh_reports)
        pr.addWidget(b_refresh); pr.addStretch()
        tl.addLayout(pr)
        self.period_lbl = lbl("", color="#888"); tl.addWidget(self.period_lbl)
        b_dl = QPushButton(" ↓ 下载"); b_dl.setObjectName("btn_outline"); b_dl.clicked.connect(self._export)
        tl.addSpacing(12); tl.addWidget(b_dl)
        L.addWidget(tb)
        self.stack = QStackedWidget(); L.addWidget(self.stack)
        self._build_balance(); self._build_income(); self._build_equity(); self._build_cf_stmt(); self._build_cashflow()
        self._switch("资产负债表")

    def _refresh_reports(self):
        """Refresh current report with selected period range"""
        current_tab = None
        for b in self._rtabs:
            if b.property("active") == "true":
                current_tab = b.text()
                break
        if current_tab:
            self._switch(current_tab)

    def _build_placeholder(self, name):
        w = QWidget(); vl = QVBoxLayout(w)
        vl.addStretch(); vl.addWidget(lbl(f"{name}（生成后显示）", color="#bbb", size=16))
        vl.addStretch(); self.stack.addWidget(w)

    def _switch(self, name):
        mapping = {"资产负债表":0,"利润表":1,"所有者权益变动表":2,"现金流量表":3,"收支统计表":4}
        for b in self._rtabs:
            b.setProperty("active","true" if b.text()==name else "false")
            b.style().unpolish(b); b.style().polish(b)
        if name in mapping:
            self.stack.setCurrentIndex(mapping[name])
            if name=="资产负债表": self._load_balance()
            elif name=="利润表": self._load_income()
            elif name=="所有者权益变动表": self._load_equity()
            elif name=="现金流量表": self._load_cf_stmt()
            elif name=="收支统计表": self._load_cashflow()

    def _make_report_table(self, cols, col_widths=None):
        t = QTableWidget(); t.setColumnCount(len(cols))
        t.setHorizontalHeaderLabels(cols)
        t.verticalHeader().setVisible(False); t.setShowGrid(True)
        t.setEditTriggers(QTableWidget.NoEditTriggers)
        if col_widths:
            for i,w in enumerate(col_widths):
                if w == -1: t.horizontalHeader().setSectionResizeMode(i, QHeaderView.Stretch)
                else: t.setColumnWidth(i,w)
        return t

    def _build_balance(self):
        w = QWidget(); L = QVBoxLayout(w); L.setContentsMargins(20,14,20,14)
        self.bs_tbl = self._make_report_table(
            ["资产项目","行次","期末金额","年初金额","负债和所有者权益","行次","期末金额","年初金额"],
            [-1,40,110,110,-1,40,110,110])
        L.addWidget(self.bs_tbl); self.stack.addWidget(w)

    def _load_balance(self):
        if not self.client_id: return
        end_period = self.rep_end_period.currentData()
        if not end_period: return
        year = end_period[:4]
        year_start = f"{year}-01"   # 本年第一期
        conn = get_db(); c = conn.cursor()

        # 期末：截止所选期间的累计发生额
        c.execute("""SELECT e.account_code, SUM(e.debit)-SUM(e.credit) net
            FROM voucher_entries e JOIN vouchers v ON v.id=e.voucher_id
            WHERE v.client_id=? AND v.period<=? AND v.status='已审核'
            GROUP BY e.account_code""", (self.client_id, end_period))
        mv = {r[0]: r[1] or 0 for r in c.fetchall()}

        # 年初：上年年末 = 期初余额 + 本年第一期之前的凭证发生额
        # 若所选期间就是01期，则年初 = 纯期初余额（凭证发生额=0）
        c.execute("""SELECT e.account_code, SUM(e.debit)-SUM(e.credit) net
            FROM voucher_entries e JOIN vouchers v ON v.id=e.voucher_id
            WHERE v.client_id=? AND v.period<? AND v.status='已审核'
            GROUP BY e.account_code""", (self.client_id, year_start))
        mv_ys = {r[0]: r[1] or 0 for r in c.fetchall()}

        c.execute("SELECT code,opening_debit,opening_credit,direction FROM accounts WHERE client_id=?",
                  (self.client_id,))
        accts = {r['code']: r for r in c.fetchall()}
        conn.close()

        # 预计算末级科目：没有任何子科目的科目
        all_codes = set(accts.keys())
        leaf_codes = {
            code for code in all_codes
            if not any(
                other != code and (other.startswith(code+".") or other.startswith(code+"_"))
                for other in all_codes
            )
        }

        def _bal_with_mv(code_prefix_list, movements):
            """通用余额计算：末级科目取期初+发生额，父科目只取发生额"""
            total = 0
            for code, a in accts.items():
                if not any(code == p or code.startswith(p+".") or code.startswith(p+"_")
                           for p in code_prefix_list):
                    continue
                net_mv = movements.get(code, 0)
                if code in leaf_codes:
                    od = a['opening_debit'] or 0; oc = a['opening_credit'] or 0
                    if a['direction'] == '借':
                        total += (od - oc) + net_mv
                    else:
                        total += (oc - od) - net_mv
                else:
                    if a['direction'] == '借':
                        total += net_mv
                    else:
                        total -= net_mv
            return total

        def bal(code_prefix_list):
            """期末余额"""
            return _bal_with_mv(code_prefix_list, mv)

        def bal_ys(code_prefix_list):
            """年初余额（上年年末 = 期初 + 本年首期前发生额）"""
            return _bal_with_mv(code_prefix_list, mv_ys)

        # ── 资产方 ──
        cash      = bal(["1001","1002","1012"])
        notes_rec = bal(["1121"])
        acct_rec  = bal(["1122"])
        _prepay_raw = bal(["1123"])
        _advrec_raw = bal(["2203"])
        # 预付账款贷方余额 → 重分类为预收账款（确保资产负债表两边同步）
        prepay  = max(0.0, _prepay_raw)  + max(0.0, -_advrec_raw)
        int_rec   = bal(["1132"])
        div_rec   = bal(["1131"])
        oth_rec   = bal(["1221"])

        # ── 年初余额（同结构，使用 bal_ys） ──
        notes_rec_y = bal_ys(["1121"])
        acct_rec_y  = bal_ys(["1122"])
        _prepay_y   = bal_ys(["1123"])
        _advrec_y   = bal_ys(["2203"])
        prepay_y    = max(0.0, _prepay_y) + max(0.0, -_advrec_y)
        int_rec_y   = bal_ys(["1132"])
        div_rec_y   = bal_ys(["1131"])
        oth_rec_y   = bal_ys(["1221"])
        cash_y      = bal_ys(["1001","1002","1012"])
        inventory_y = (bal_ys(["1401","1402","1403","1404","1405","1406","1407","1408","1409","1411","1415","1421"])
                      - abs(bal_ys(["1471","1472"])))
        prepd_exp_y = bal_ys(["1901"])
        fa_y        = bal_ys(["1601"]) - abs(bal_ys(["1602"])) - abs(bal_ys(["1603"]))
        wip_y       = bal_ys(["1604"])
        intangible_y= bal_ys(["1701"]) - abs(bal_ys(["1702"])) - abs(bal_ys(["1703"]))
        lt_prepaid_y= bal_ys(["1801"])
        deferred_a_y= bal_ys(["1811"])
        avail_sale_y   = bal_ys(["1503"])
        held_to_mat_y  = bal_ys(["1501"]) - abs(bal_ys(["1502"]))
        lt_eq_invest_y = bal_ys(["1511"])
        invest_prop_y  = bal_ys(["1521"])
        lt_equity_y    = avail_sale_y + held_to_mat_y + lt_eq_invest_y + invest_prop_y
        cur_asset_y  = cash_y+notes_rec_y+acct_rec_y+prepay_y+int_rec_y+div_rec_y+oth_rec_y+inventory_y+prepd_exp_y
        noncur_asset_y = fa_y+wip_y+intangible_y+lt_prepaid_y+lt_equity_y+deferred_a_y
        total_asset_y  = cur_asset_y + noncur_asset_y

        st_loan_y   = bal_ys(["2001"]); notes_pay_y = bal_ys(["2201"])
        acct_pay_y  = bal_ys(["2202"]); adv_rec_y   = max(0.0, _advrec_y) + max(0.0, -_prepay_y)
        emp_pay_y   = bal_ys(["2211"]); tax_pay_y   = bal_ys(["2221"])
        int_pay_y   = bal_ys(["2231"]); div_pay_y   = bal_ys(["2232"])
        oth_pay_y   = bal_ys(["2241"])
        cur_liab_y  = st_loan_y+notes_pay_y+acct_pay_y+adv_rec_y+emp_pay_y+tax_pay_y+int_pay_y+div_pay_y+oth_pay_y
        lt_loan_y   = bal_ys(["2501"]); bonds_pay_y = bal_ys(["2502"])
        lt_payable_y= bal_ys(["2701"]); est_liab_y  = bal_ys(["2801"])
        deferred_l_y= bal_ys(["2901"])
        noncur_liab_y = lt_loan_y+bonds_pay_y+lt_payable_y+est_liab_y+deferred_l_y
        total_liab_y  = cur_liab_y + noncur_liab_y
        cap_y     = bal_ys(["4001"]); cap_res_y = bal_ys(["4002"])
        surp_res_y= bal_ys(["4101"])
        profit_y  = bal_ys(["4103"]) + bal_ys(["4104"])
        tsy_y     = bal_ys(["4201"])
        total_equity_y = cap_y + cap_res_y + surp_res_y + profit_y - tsy_y
        total_le_y     = total_liab_y + total_equity_y
        # 存货 = 各存货科目合计 - 存货跌价准备 - 消耗性生物资产跌价准备
        inventory = (bal(["1401","1402","1403","1404","1405","1406","1407","1408","1409","1411","1415","1421"])
                     - abs(bal(["1471","1472"])))
        prepd_exp = bal(["1901"])   # 待处理财产损溢
        fa        = bal(["1601"]) - abs(bal(["1602"])) - abs(bal(["1603"]))
        wip       = bal(["1604"])
        intangible= bal(["1701"]) - abs(bal(["1702"])) - abs(bal(["1703"]))
        lt_prepaid= bal(["1801"])   # 长期待摊费用
        deferred_a= bal(["1811"])   # 递延所得税资产
        avail_sale  = bal(["1503"])                    # 可供出售金融资产
        held_to_mat = bal(["1501"]) - abs(bal(["1502"]))  # 持有至到期投资净额
        lt_eq_invest= bal(["1511"])                    # 长期股权投资
        invest_prop = bal(["1521"])                    # 投资性房地产
        lt_equity   = avail_sale + held_to_mat + lt_eq_invest + invest_prop
        cur_asset = cash+notes_rec+acct_rec+prepay+int_rec+div_rec+oth_rec+inventory+prepd_exp
        noncur_asset = fa+wip+intangible+lt_prepaid+lt_equity+deferred_a
        total_asset = cur_asset + noncur_asset

        # ── 负债方 ──
        st_loan   = bal(["2001"])
        notes_pay = bal(["2201"])
        acct_pay  = bal(["2202"])
        adv_rec   = max(0.0, _advrec_raw) + max(0.0, -_prepay_raw)
        emp_pay   = bal(["2211"])
        tax_pay   = bal(["2221"])
        int_pay   = bal(["2231"])
        div_pay   = bal(["2232"])
        oth_pay   = bal(["2241"])
        cur_liab  = st_loan+notes_pay+acct_pay+adv_rec+emp_pay+tax_pay+int_pay+div_pay+oth_pay
        lt_loan   = bal(["2501"])
        bonds_pay = bal(["2502"])
        lt_payable= bal(["2701"])   # 长期应付款
        est_liab  = bal(["2801"])   # 预计负债
        deferred_l= bal(["2901"])   # 递延所得税负债
        noncur_liab = lt_loan+bonds_pay+lt_payable+est_liab+deferred_l
        total_liab = cur_liab + noncur_liab

        # ── 所有者权益 ──
        # 兼容3xxx（标准）和4xxx（部分软件）两套所有者权益科目体系
        cap       = bal(["4001"])
        cap_res   = bal(["4002"])
        surp_res  = bal(["4101"])
        profit    = bal(["4103"]) + bal(["4104"])
        tsy_stock = bal(["4201"])
        total_equity = cap + cap_res + surp_res + profit - tsy_stock
        total_le     = total_liab + total_equity

        def R(label, rowno, left_val, right_label="", right_rowno="", right_val=None,
              is_header=False, is_total=False,
              left_ys=None, right_ys=None):
            return (label, rowno, left_val, right_label, right_rowno, right_val,
                    is_header, is_total, left_ys, right_ys)

        rows = [
            R("流动资产：","","",  "流动负债：","","",          True),
            R("货币资金","1",cash,            "短期借款","34",st_loan,         left_ys=cash_y,        right_ys=st_loan_y),
            R("以公允价值计量且其变动\n计入当期损益的金融资产","2",0, "以公允价值计量且其变动\n计入当期损益的金融负债","35",0),
            R("衍生金融资产","3",0,            "衍生金融负债","36",0),
            R("应收票据","4",notes_rec,        "应付票据","37",notes_pay,       left_ys=notes_rec_y,   right_ys=notes_pay_y),
            R("应收账款","5",acct_rec,         "应付账款","38",acct_pay,        left_ys=acct_rec_y,    right_ys=acct_pay_y),
            R("预付款项","6",prepay,           "预收款项","39",adv_rec,         left_ys=prepay_y,      right_ys=adv_rec_y),
            R("应收利息","7",int_rec,          "应付职工薪酬","40",emp_pay,     left_ys=int_rec_y,     right_ys=emp_pay_y),
            R("应收股利","8",div_rec,          "应交税费","41",tax_pay,         left_ys=div_rec_y,     right_ys=tax_pay_y),
            R("其他应收款","9",oth_rec,         "应付利息","42",int_pay,        left_ys=oth_rec_y,     right_ys=int_pay_y),
            R("存货","10",inventory,           "应付股利","43",div_pay,         left_ys=inventory_y,   right_ys=div_pay_y),
            R("持有待售资产","11",0,            "其他应付款","44",oth_pay,                              right_ys=oth_pay_y),
            R("一年内到期的非流动资产","12",0,  "持有待售负债","45",0),
            R("其他流动资产","13",prepd_exp,    "一年内到期的非流动负债","46",0, left_ys=prepd_exp_y),
            R("流动资产合计","14",cur_asset,   "其他流动负债","47",0,   False,True, left_ys=cur_asset_y, right_ys=cur_liab_y),
            R("非流动资产：","","",            "流动负债合计","48",cur_liab,True,True,                  right_ys=cur_liab_y),
            R("可供出售金融资产","15",avail_sale,   "非流动负债：","","",   False,False, left_ys=avail_sale_y),
            R("持有至到期投资","16",held_to_mat,    "长期借款","49",lt_loan, left_ys=held_to_mat_y,  right_ys=lt_loan_y),
            R("长期应收款","17",0,                  "应付债券","50",bonds_pay),
            R("长期股权投资","18",lt_eq_invest,     "其中：优先股","51",0,  left_ys=lt_eq_invest_y),
            R("投资性房地产","19",invest_prop,       "永续债","52",0,        left_ys=invest_prop_y),
            R("固定资产","20",fa,              "长期应付款","53",lt_payable,    left_ys=fa_y,           right_ys=lt_payable_y),
            R("在建工程","21",wip,             "专项应付款","54",0,             left_ys=wip_y),
            R("工程物资","22",0,               "预计负债","55",est_liab,                               right_ys=est_liab_y),
            R("固定资产清理","23",0,            "递延收益","56",0),
            R("生产性生物资产","24",0,          "递延所得税负债","57",deferred_l,                       right_ys=deferred_l_y),
            R("油气资产","25",0,               "其他非流动负债","58",0),
            R("无形资产","26",intangible,       "非流动负债合计","59",noncur_liab, False,True, left_ys=intangible_y, right_ys=noncur_liab_y),
            R("开发支出","27",0,               "负债合计","60",total_liab,    False,True,               right_ys=total_liab_y),
            R("商誉","28",0,                   "所有者权益（或股东权益）：","","",True),
            R("长期待摊费用","29",lt_prepaid,   "实收资本（或股本）","61",cap,  left_ys=lt_prepaid_y,   right_ys=cap_y),
            R("递延所得税资产","30",deferred_a, "其他权益工具","62",0,          left_ys=deferred_a_y),
            R("其他非流动资产","31",0,          "其中：优先股","63",0),
            R("非流动资产合计","32",noncur_asset,"永续债","64",0,          False,True, left_ys=noncur_asset_y),
            R("","","",                        "资本公积","65",cap_res,                                 right_ys=cap_res_y),
            R("","","",                        "减：库存股","66",tsy_stock,                             right_ys=tsy_y),
            R("","","",                        "其他综合收益","67",0),
            R("","","",                        "盈余公积","68",surp_res,                                right_ys=surp_res_y),
            R("","","",                        "未分配利润","69",profit,                                right_ys=profit_y),
            R("","","",                        "所有者权益合计","70",total_equity, False,True,          right_ys=total_equity_y),
            R("资产总计","33",total_asset,     "负债和所有者权益总计","71",total_le,False,True, left_ys=total_asset_y, right_ys=total_le_y),
        ]

        self.bs_tbl.setRowCount(len(rows))
        for i,(l_name,l_row,l_val,r_name,r_row,r_val,is_hdr,is_tot,l_ys,r_ys) in enumerate(rows):
            self.bs_tbl.setRowHeight(i, 32)
            # Left
            for j,(text,align) in enumerate([
                (l_name, Qt.AlignLeft|Qt.AlignVCenter),
                (str(l_row), Qt.AlignCenter),
                (fmt_amt(l_val) if isinstance(l_val,(int,float)) else "", Qt.AlignRight|Qt.AlignVCenter),
                (fmt_amt(l_ys) if isinstance(l_ys,(int,float)) else "", Qt.AlignRight|Qt.AlignVCenter),
            ]):
                it = QTableWidgetItem(text); it.setTextAlignment(align)
                if is_hdr or is_tot:
                    it.setBackground(QColor("#f0f4ff" if is_hdr else "#f5f7fa"))
                    if is_tot: it.setFont(QFont("",weight=QFont.Bold))
                if j==0 and is_hdr: it.setForeground(QColor("#3d6fdb"))
                if j==2 and isinstance(l_val,(int,float)) and l_val<0:
                    it.setForeground(QColor("#e05252"))
                if j==3 and isinstance(l_ys,(int,float)) and l_ys<0:
                    it.setForeground(QColor("#e05252"))
                self.bs_tbl.setItem(i,j,it)
            # Right
            for j,(text,align) in enumerate([
                (r_name, Qt.AlignLeft|Qt.AlignVCenter),
                (str(r_row), Qt.AlignCenter),
                (fmt_amt(r_val) if isinstance(r_val,(int,float)) else "", Qt.AlignRight|Qt.AlignVCenter),
                (fmt_amt(r_ys) if isinstance(r_ys,(int,float)) else "", Qt.AlignRight|Qt.AlignVCenter),
            ],4):
                it = QTableWidgetItem(text); it.setTextAlignment(align)
                if is_hdr or is_tot:
                    it.setBackground(QColor("#f0f4ff" if is_hdr else "#f5f7fa"))
                    if is_tot: it.setFont(QFont("",weight=QFont.Bold))
                if j==4 and is_hdr: it.setForeground(QColor("#3d6fdb"))
                if j==6 and isinstance(r_val,(int,float)) and r_val<0:
                    it.setForeground(QColor("#e05252"))
                if j==7 and isinstance(r_ys,(int,float)) and r_ys<0:
                    it.setForeground(QColor("#e05252"))
                self.bs_tbl.setItem(i,j,it)

    def _build_income(self):
        w = QWidget(); L = QVBoxLayout(w); L.setContentsMargins(20,14,20,14)
        self.inc_tbl = self._make_report_table(
            ["项目","行次","本期金额","本年累计"],[-1,40,160,160])
        L.addWidget(self.inc_tbl); self.stack.addWidget(w)

    def _load_income(self):
        if not self.client_id: return
        start_period = self.rep_start_period.currentData()
        end_period = self.rep_end_period.currentData()
        if not start_period or not end_period: return
        conn = get_db(); c = conn.cursor()

        def fetch_period(period_filter):
            c.execute("""SELECT e.account_code, SUM(e.credit)-SUM(e.debit) net
                FROM voucher_entries e JOIN vouchers v ON v.id=e.voucher_id
                WHERE v.client_id=? AND v.period"""+period_filter+""" AND v.status='已审核'
                GROUP BY e.account_code""", (self.client_id,))
            return {r[0]: r[1] or 0 for r in c.fetchall()}

        # Current period: from start to end
        cur = fetch_period(f">='{start_period}' AND v.period<='{end_period}'")
        year = end_period[:4]
        # Year-to-date: from year start to end
        ytd = fetch_period(f" LIKE '{year}%' AND v.period<='{end_period}'")
        conn.close()

        def g(codes, d=None):
            """Sum credit-minus-debit net for all accounts matching any prefix in codes list."""
            if d is None: d = cur
            if isinstance(codes, str): codes = [codes]
            total = 0
            for acct_code, val in d.items():
                for code in codes:
                    if acct_code == code or acct_code.startswith(code+".") or acct_code.startswith(code+"_"):
                        total += val
                        break
            return total
        def gy(codes): return g(codes, ytd)

        # Try 6xxx first, fall back to 5xxx
        use_6xxx = bool(g(["6001","6401","6601","6602"]))  # 检测6xxx科目体系

        if use_6xxx:
            # 6xxx科目体系（用友/金蝶新版）
            rev      = g(["6001","6051"])   # 主营业务收入+其他业务收入         # 主营+其他业务收入（贷方余额为正）
            cost_n   = -g(["6401","6402"])               # 主营+其他业务成本（借方为正，取负得正数成本）
            tax      = -g(["6403"])                      # 税金及附加
            sell     = -g(["6601"])                      # 销售费用
            mgmt     = -g(["6602"])                      # 管理费用
            rnd      = -g(["6604"])                      # 研发费用
            fin_net  = g(["6603"])                       # 财务费用净额（正=净收益，负=净支出）
            inv_g    = g(["6111"])                       # 投资收益
            fv_g     = g(["6101"])   # 公允价值变动损益                       # 公允价值变动
            asset_d  = g(["6301"])                       # 营业外收入（此处作资产处置收益）
            op_profit = rev - cost_n - tax - sell - mgmt - rnd + fin_net + inv_g + fv_g
            nop_inc   = g(["6301"])                      # 营业外收入
            nop_exp   = -g(["6711"])                     # 营业外支出
            tax_exp   = -g(["6801"])                     # 所得税费用
            # YTD
            rev_y    = gy(["6001","6051"])
            cost_y   = -gy(["6401","6402"])
            sell_y   = -gy(["6601"]); mgmt_y = -gy(["6602"])
            fin_y    = gy(["6603"]); inv_y = gy(["6111"])
            nop_y    = gy(["6301"]); nopx_y = -gy(["6711"])
            tax_y    = -gy(["6801"])
            op_y     = rev_y - cost_y - gy(["6403"]) - sell_y - mgmt_y + fin_y + inv_y
            net_y    = op_y + nop_y + nopx_y + tax_y
        else:
            # 未检测到6xxx科目凭证，所有值置零
            rev = cost_n = tax = sell = mgmt = rnd = fin_net = 0
            inv_g = fv_g = asset_d = nop_inc = nop_exp = tax_exp = 0
            op_profit = rev_y = cost_y = sell_y = mgmt_y = 0
            fin_y = inv_y = nop_y = nopx_y = tax_y = op_y = net_y = 0

        total_profit = op_profit + nop_inc + nop_exp
        net_profit   = total_profit + tax_exp

        rows_data = [
            ("一、营业收入",           "1",  rev,           rev_y,    True),
            ("  减：营业成本",          "2",  cost_n,        cost_y,   False),
            ("      税金及附加",        "3",  tax,           0,        False),
            ("      销售费用",          "4",  sell,          sell_y if use_6xxx else 0, False),
            ("      管理费用",          "5",  mgmt,          mgmt_y if use_6xxx else 0, False),
            ("      研发费用",          "6",  rnd,           0,        False),
            ("  加：财务费用（收益以-号填列）","7", fin_net, fin_y if use_6xxx else 0, False),
            ("      投资收益",          "8",  inv_g,         inv_y if use_6xxx else 0, False),
            ("      公允价值变动收益",   "9",  fv_g,          0,        False),
            ("      资产处置收益",       "9a", asset_d,       0,        False),
            ("二、营业利润（亏损）",     "10", op_profit,     op_y,     True),
            ("  加：营业外收入",         "11", nop_inc,       nop_y if use_6xxx else 0, False),
            ("  减：营业外支出",         "12", nop_exp,       nopx_y if use_6xxx else 0, False),
            ("三、利润总额（亏损总额）", "13", total_profit,  0,        True),
            ("  减：所得税费用",         "14", tax_exp,       tax_y if use_6xxx else 0, False),
            ("四、净利润（净亏损）",     "15", net_profit,    net_y,    True),
            ("  其中：归属于母公司股东的净利润","16", net_profit, 0, False),
            ("        少数股东损益",    "17", 0,             0,        False),
            ("五、其他综合收益的税后净额","18", 0,            0,        True),
            ("六、综合收益总额",         "19", net_profit,   0,        True),
            ("  其中：归属于母公司股东的综合收益","20", net_profit, 0, False),
            ("        归属于少数股东的综合收益","21", 0,      0,       False),
            ("七、每股收益","","","",True),
            ("  基本每股收益",           "22", 0,             0,        False),
            ("  稀释每股收益",           "23", 0,             0,        False),
        ]

        self.inc_tbl.setRowCount(len(rows_data))
        for i,row_item in enumerate(rows_data):
            name = row_item[0]; rowno = row_item[1]
            cur_v = row_item[2]; ytd_v = row_item[3]
            is_key = row_item[4] if len(row_item)>4 else False
            self.inc_tbl.setRowHeight(i, 34)
            bg = QColor("#f0f4ff") if is_key else None
            for j,v in enumerate([name, str(rowno) if rowno else "",
                                   fmt_amt(cur_v) if isinstance(cur_v,(int,float)) else "",
                                   fmt_amt(ytd_v) if isinstance(ytd_v,(int,float)) else ""]):
                it = QTableWidgetItem(v)
                it.setTextAlignment(Qt.AlignLeft|Qt.AlignVCenter if j==0 else Qt.AlignCenter if j==1 else Qt.AlignRight|Qt.AlignVCenter)
                if is_key:
                    it.setFont(QFont("",weight=QFont.Bold))
                    if bg: it.setBackground(bg)
                if j>=2 and isinstance(cur_v,(int,float)):
                    val = cur_v if j==2 else ytd_v
                    if val and val < 0: it.setForeground(QColor("#ff4d4f"))
                self.inc_tbl.setItem(i,j,it)


    def _build_equity(self):
        w = QWidget(); L = QVBoxLayout(w); L.setContentsMargins(20,14,20,14); L.setSpacing(8)
        title_row = QHBoxLayout()
        title_row.addWidget(lbl("所有者权益变动表", bold=True, size=15))
        title_row.addStretch()
        title_row.addWidget(lbl("（企业会计准则格式）", color="#888", size=12))
        L.addLayout(title_row)
        L.addWidget(lbl("单位：元", color="#aaa", size=11))

        self.eq_tbl = self._make_report_table(
            ["项目",
             "实收资本(股本)",
             "资本公积",
             "其他综合收益",
             "盈余公积",
             "未分配利润",
             "合计"],
            [-1, 110, 110, 100, 100, 110, 110]
        )
        self.eq_tbl.horizontalHeader().setDefaultAlignment(Qt.AlignCenter)
        self.eq_tbl.setWordWrap(True)
        self.eq_tbl.horizontalHeader().setSectionResizeMode(QHeaderView.Interactive)
        self.eq_tbl.horizontalHeader().setSectionResizeMode(0, QHeaderView.Stretch)
        L.addWidget(self.eq_tbl)
        self.stack.addWidget(w)

    def _load_equity(self):
        if not self.client_id: return
        end_period = self.rep_end_period.currentData()
        if not end_period: return
        conn = get_db(); c = conn.cursor()
        year = end_period[:4]

        # Fetch year-to-date balances from voucher entries (approved only)
        c.execute("""SELECT e.account_code, SUM(e.debit)-SUM(e.credit) net
            FROM voucher_entries e JOIN vouchers v ON v.id=e.voucher_id
            WHERE v.client_id=? AND v.period LIKE ? AND v.status='已审核'
            GROUP BY e.account_code""", (self.client_id, f"{year}%"))
        ytd = {r[0]: -(r[1] or 0) for r in c.fetchall()}  # credit-normal for equity

        # Opening balances from accounts table
        c.execute("SELECT code, opening_credit FROM accounts WHERE client_id=?", (self.client_id,))
        opening = {r[0]: r[1] or 0 for r in c.fetchall()}
        conn.close()

        def op(code):
            # Support both 3xxx and 4xxx equity accounts
            alt = {"3001":"4001","3002":"4002","3101":"4101","3103":"4103","3104":"4104","3201":"4201"}
            return opening.get(code, 0) or opening.get(alt.get(code,""), 0)
        def mv(code):
            alt = {"3001":"4001","3002":"4002","3101":"4101","3103":"4103","3104":"4104","3201":"4201"}
            return ytd.get(code, 0) or ytd.get(alt.get(code,""), 0)

        cap_op  = op("3001"); cap_mv  = mv("3001")
        cprs_op = op("3002"); cprs_mv = mv("3002")
        oci_op  = 0;          oci_mv  = 0          # 其他综合收益（暂无专用科目）
        surp_op = op("3101"); surp_mv = mv("3101")
        re_op   = op("3103") + op("3104")
        re_mv   = mv("3103") + mv("3104")          # 本年利润 + 利润分配

        def row_data(label, c1, c2, c3, c4, c5, bold=False, bg=None):
            total = c1+c2+c3+c4+c5
            return (label, c1, c2, c3, c4, c5, total, bold, bg)

        rows = [
            row_data("一、上年年末余额",    cap_op,  cprs_op, oci_op,  surp_op, re_op,  bold=True,  bg="#f0f4ff"),
            row_data("  加：会计政策变更",  0, 0, 0, 0, 0),
            row_data("     前期差错更正",   0, 0, 0, 0, 0),
            row_data("二、本年年初余额",    cap_op,  cprs_op, oci_op,  surp_op, re_op,  bold=True,  bg="#f0f4ff"),
            row_data("三、本年增减变动",    cap_mv,  cprs_mv, oci_mv,  surp_mv, re_mv,  bold=True,  bg="#fafafa"),
            row_data("  (一)综合收益总额",  0,       0,       oci_mv,  0,       re_mv),
            row_data("  (二)所有者投入",    cap_mv,  cprs_mv, 0,       0,       0),
            row_data("  (三)利润分配",      0,       0,       0,       surp_mv, re_mv - re_mv),
            row_data("四、本年年末余额",
                     cap_op+cap_mv, cprs_op+cprs_mv, oci_op+oci_mv,
                     surp_op+surp_mv, re_op+re_mv,    bold=True, bg="#e6f0ff"),
        ]

        self.eq_tbl.setRowCount(len(rows))
        for i, (label,c1,c2,c3,c4,c5,total,bold,bg) in enumerate(rows):
            self.eq_tbl.setRowHeight(i, 38)
            vals = [label, c1, c2, c3, c4, c5, total]
            for j, v in enumerate(vals):
                text = v if j == 0 else (fmt_amt(v) if v else "")
                it = QTableWidgetItem(text)
                it.setTextAlignment(Qt.AlignLeft|Qt.AlignVCenter if j==0 else Qt.AlignRight|Qt.AlignVCenter)
                if bold: it.setFont(QFont("", weight=QFont.Bold))
                if bg:   it.setBackground(QColor(bg))
                if j > 0 and isinstance(v, float) and v < 0:
                    it.setForeground(QColor("#e05252"))
                self.eq_tbl.setItem(i, j, it)

    def _build_cf_stmt(self):
        w = QWidget(); L = QVBoxLayout(w); L.setContentsMargins(20,14,20,14); L.setSpacing(8)
        hdr = QHBoxLayout()
        hdr.addWidget(lbl("现金流量表", bold=True, size=15)); hdr.addStretch()
        b_dl = QPushButton("导出Excel"); b_dl.setObjectName("btn_outline")
        b_dl.clicked.connect(self._export_cf_stmt)
        hdr.addWidget(b_dl); L.addLayout(hdr)
        L.addWidget(lbl("（采用直接法，现金及现金等价物 = 库存现金+银行存款+其他货币资金）",
                         color="#888", size=12))
        self.cf_stmt_tbl = self._make_report_table(
            ["项目", "行次", "本期金额", "本年累计金额"],
            [-1, 40, 140, 140])
        L.addWidget(self.cf_stmt_tbl)
        self.stack.addWidget(w)

    def _get_cash_balance(self, c, client_id, period_end):
        """期末现金余额 = 期初 + 本年至今净发生额"""
        # Opening balance from accounts
        c.execute("""SELECT SUM(opening_debit - opening_credit) FROM accounts
            WHERE client_id=? AND (code='1001' OR code LIKE '1001.%' OR code LIKE '1001_%'
              OR code='1002' OR code LIKE '1002.%' OR code LIKE '1002_%'
              OR code='1012' OR code LIKE '1012.%' OR code LIKE '1012_%')""",
            (client_id,))
        opening = c.fetchone()[0] or 0
        if not period_end:
            return opening
        year = period_end[:4]
        c.execute("""SELECT SUM(e.debit) - SUM(e.credit) FROM voucher_entries e
            JOIN vouchers v ON v.id=e.voucher_id
            WHERE v.client_id=? AND v.period<=? AND v.period LIKE ? AND v.status='已审核'
            AND (e.account_code='1001' OR e.account_code LIKE '1001.%' OR e.account_code LIKE '1001_%'
              OR e.account_code='1002' OR e.account_code LIKE '1002.%' OR e.account_code LIKE '1002_%'
              OR e.account_code='1012' OR e.account_code LIKE '1012.%' OR e.account_code LIKE '1012_%')""",
            (client_id, period_end, f"{year}%"))
        ytd_net = c.fetchone()[0] or 0
        return opening + ytd_net

    def _compute_cf(self, c, client_id, start_period, end_period):
        """
        Compute cash flow by analyzing cash account counterparts in vouchers.
        Returns dict: row_key -> amount (positive = inflow, negative = outflow shown as positive)
        Two dicts returned: current_period and ytd.
        """
        year = end_period[:4]

        def _analyze(p_start, p_end):
            """Analyze cash flows for a period range."""
            # Get all voucher IDs with cash account entries in range
            c.execute("""SELECT DISTINCT v.id FROM vouchers v
                JOIN voucher_entries e ON e.voucher_id=v.id
                WHERE v.client_id=? AND v.period>=? AND v.period<=? AND v.status='已审核'
                AND (e.account_code='1001' OR e.account_code LIKE '1001.%' OR e.account_code LIKE '1001_%'
                  OR e.account_code='1002' OR e.account_code LIKE '1002.%' OR e.account_code LIKE '1002_%'
                  OR e.account_code='1012' OR e.account_code LIKE '1012.%' OR e.account_code LIKE '1012_%')""",
                (client_id, p_start, p_end))
            vids = [r[0] for r in c.fetchall()]

            rows = {}  # row_number -> amount

            def add(key, amt):
                rows[key] = rows.get(key, 0) + amt

            for vid in vids:
                c.execute("SELECT account_code, debit, credit FROM voucher_entries WHERE voucher_id=?", (vid,))
                entries = c.fetchall()

                cash_in = 0; cash_out = 0
                non_cash = []
                for e in entries:
                    code = e[0] or ""
                    d = e[1] or 0; cr = e[2] or 0
                    if (code == '1001' or code.startswith('1001.') or code.startswith('1001_') or
                        code == '1002' or code.startswith('1002.') or code.startswith('1002_') or
                        code == '1012' or code.startswith('1012.') or code.startswith('1012_')):
                        cash_in += d; cash_out += cr
                    else:
                        non_cash.append((code, d, cr))

                # Classify inflows (cash debited)
                if cash_in > 0:
                    for code, d, cr in non_cash:
                        amt = cr  # credit side = source of cash
                        if amt <= 0: continue
                        p = code[:4]
                        # Revenue accounts → 销售商品收到现金
                        if (code.startswith('6001') or code.startswith('6002') or
                            code.startswith('6051') or code.startswith('5001') or
                            code.startswith('5051') or code.startswith('1122') or
                            code.startswith('2203')):
                            add('r1', amt)
                        elif code.startswith('2221') or code.startswith('1321'):
                            add('r2', amt)  # 税费返还
                        elif code.startswith('6301') or code.startswith('5301'):
                            add('r3', amt)  # 其他经营收入
                        elif (code.startswith('6111') or code.startswith('5111') or
                              code.startswith('1511') or code.startswith('1521') or
                              code.startswith('1131') or code.startswith('1132')):
                            add('r12', amt)  # 取得投资收益
                        elif code.startswith('1601') or code.startswith('1604'):
                            add('r13', amt)  # 处置固定资产
                        elif code.startswith('2001') or code.startswith('2501'):
                            add('r24', amt)  # 取得借款
                        elif code.startswith('3001') or code.startswith('4001'):
                            add('r23', amt)  # 吸收投资
                        else:
                            add('r3', amt)   # 其他经营收入

                # Classify outflows (cash credited)
                if cash_out > 0:
                    for code, d, cr in non_cash:
                        amt = d  # debit side = destination of cash
                        if amt <= 0: continue
                        if (code.startswith('1403') or code.startswith('1401') or
                            code.startswith('1405') or code.startswith('6401') or
                            code.startswith('6402') or code.startswith('5401') or
                            code.startswith('5402') or code.startswith('2202') or
                            code.startswith('1221')):
                            add('r5', amt)   # 购买商品
                        elif code.startswith('2211'):
                            add('r6', amt)   # 支付员工
                        elif code.startswith('2221') or code.startswith('2231'):
                            add('r7', amt)   # 支付税费
                        elif (code.startswith('6601') or code.startswith('6602') or
                              code.startswith('6603') or code.startswith('5501') or
                              code.startswith('5502') or code.startswith('5503') or
                              code.startswith('2241') or code.startswith('1461')):
                            add('r8', amt)   # 其他经营支出
                        elif (code.startswith('1601') or code.startswith('1604') or
                              code.startswith('1605') or code.startswith('1701')):
                            add('r17', amt)  # 购建固定资产
                        elif (code.startswith('1801') or code.startswith('1511') or
                              code.startswith('1521') or code.startswith('1531')):
                            add('r18', amt)  # 投资支出
                        elif code.startswith('2001') or code.startswith('2501'):
                            add('r27', amt)  # 偿还借款
                        elif (code.startswith('3104') or code.startswith('4104') or
                              code.startswith('2232')):
                            add('r28', amt)  # 分配股利
                        else:
                            add('r8', amt)   # 其他经营支出

            return rows

        cur = _analyze(start_period, end_period)
        ytd = _analyze(f"{year}-01", end_period)
        return cur, ytd

    def _load_cf_stmt(self):
        if not self.client_id: return
        start_period = self.rep_start_period.currentData()
        end_period   = self.rep_end_period.currentData()
        if not start_period or not end_period: return
        year = end_period[:4]

        conn = get_db(); c = conn.cursor()

        # Cash balances
        cash_end  = self._get_cash_balance(c, self.client_id, end_period)
        cash_beg  = self._get_cash_balance(c, self.client_id,
                        f"{year}-01" if start_period[:4] == year else start_period)
        cash_open = self._get_cash_balance(c, self.client_id, None)  # opening from accounts

        # Compute cash flow amounts
        cur, ytd = self._compute_cf(c, self.client_id, start_period, end_period)

        # Subtotals
        def g(d, *keys): return sum(d.get(k, 0) for k in keys)

        # Current period
        ci  = g(cur,'r1','r2','r3')       # 经营流入
        co  = g(cur,'r5','r6','r7','r8')  # 经营流出
        cn  = ci - co                      # 经营净额
        ii  = g(cur,'r11','r12','r13','r14','r15')
        io_ = g(cur,'r17','r18','r19','r20')
        inv_n = ii - io_
        fi  = g(cur,'r23','r24','r25')
        fo  = g(cur,'r27','r28','r29')
        fin_n = fi - fo
        net_cur = cn + inv_n + fin_n

        # YTD
        ci_y  = g(ytd,'r1','r2','r3')
        co_y  = g(ytd,'r5','r6','r7','r8')
        cn_y  = ci_y - co_y
        ii_y  = g(ytd,'r11','r12','r13','r14','r15')
        io_y  = g(ytd,'r17','r18','r19','r20')
        inv_ny = ii_y - io_y
        fi_y  = g(ytd,'r23','r24','r25')
        fo_y  = g(ytd,'r27','r28','r29')
        fin_ny = fi_y - fo_y
        net_ytd = cn_y + inv_ny + fin_ny

        # Net profit for supplementary
        c.execute("""SELECT e.account_code, SUM(e.credit)-SUM(e.debit) net
            FROM voucher_entries e JOIN vouchers v ON v.id=e.voucher_id
            WHERE v.client_id=? AND v.period>=? AND v.period<=? AND v.status='已审核'
            GROUP BY e.account_code""", (self.client_id, f"{year}-01", end_period))
        mv_ytd = {r[0]: r[1] or 0 for r in c.fetchall()}
        c.execute("""SELECT e.account_code, SUM(e.credit)-SUM(e.debit) net
            FROM voucher_entries e JOIN vouchers v ON v.id=e.voucher_id
            WHERE v.client_id=? AND v.period>=? AND v.period<=? AND v.status='已审核'
            GROUP BY e.account_code""", (self.client_id, start_period, end_period))
        mv_cur = {r[0]: r[1] or 0 for r in c.fetchall()}
        conn.close()

        def net_profit(mv):
            use6 = any(k.startswith('6') for k in mv)
            if use6:
                rev  = sum(v for k,v in mv.items() if k[:4]<'6400' and k[0]=='6')
                cost = -sum(v for k,v in mv.items() if k[:4]>='6400' and k[0]=='6')
                return rev + cost
            else:
                rev  = sum(v for k,v in mv.items() if k[0]=='5' and k[:4]<'5400')
                cost = -sum(v for k,v in mv.items() if k[0]=='5' and k[:4]>='5400')
                return rev + cost

        np_cur = net_profit(mv_cur)
        np_ytd = net_profit(mv_ytd)

        # AR/AP changes for supplementary (ytd)
        def bal_chg(mv, codes):
            total = 0
            for k, v in mv.items():
                for code in codes:
                    if k == code or k.startswith(code+'.') or k.startswith(code+'_'):
                        total += v; break
            return total
        ar_chg  = -bal_chg(mv_ytd, ['1122','1123','1131','1132','1221'])
        ap_chg  =  bal_chg(mv_ytd, ['2202','2203','2211','2241'])

        # ── Build table rows ──
        BOLD_BG = "#f0f4ff"; HDR_BG = "#e6ecf8"

        def R(label, rowno, cur_val, ytd_val, style="normal"):
            return (label, str(rowno) if rowno else "", cur_val, ytd_val, style)

        rows = [
            # ── 经营活动 ──
            R("一、经营活动产生的现金流量：",  "", None, None, "header"),
            R("  销售商品、提供劳务收到的现金","1", cur.get('r1',0), ytd.get('r1',0)),
            R("  收到的税费返还",              "2", cur.get('r2',0), ytd.get('r2',0)),
            R("  收到的其他与经营活动有关的现金","3",cur.get('r3',0),ytd.get('r3',0)),
            R("  经营活动现金流入小计",         "4", ci,   ci_y,  "subtotal"),
            R("  购买商品、接受劳务支付的现金", "5", cur.get('r5',0), ytd.get('r5',0)),
            R("  支付给职工以及为职工支付的现金","6",cur.get('r6',0),ytd.get('r6',0)),
            R("  支付的各项税费",               "7", cur.get('r7',0), ytd.get('r7',0)),
            R("  支付的其他与经营活动有关的现金","8",cur.get('r8',0),ytd.get('r8',0)),
            R("  经营活动现金流出小计",          "9", co,   co_y,  "subtotal"),
            R("  经营活动产生的现金流量净额",   "10", cn,   cn_y,  "total"),
            # ── 投资活动 ──
            R("二、投资活动产生的现金流量：",   "", None, None,   "header"),
            R("  收回投资收到的现金",           "11", cur.get('r11',0), ytd.get('r11',0)),
            R("  取得投资收益收到的现金",        "12", cur.get('r12',0), ytd.get('r12',0)),
            R("  处置固定资产收回的现金净额",   "13", cur.get('r13',0), ytd.get('r13',0)),
            R("  处置子公司收到的现金净额",     "14", cur.get('r14',0), ytd.get('r14',0)),
            R("  收到的其他与投资活动有关的现金","15",cur.get('r15',0),ytd.get('r15',0)),
            R("  投资活动现金流入小计",         "16", ii,   ii_y,  "subtotal"),
            R("  购建固定资产支付的现金",        "17", cur.get('r17',0), ytd.get('r17',0)),
            R("  投资支付的现金",               "18", cur.get('r18',0), ytd.get('r18',0)),
            R("  取得子公司支付的现金净额",     "19", cur.get('r19',0), ytd.get('r19',0)),
            R("  支付的其他与投资活动有关的现金","20",cur.get('r20',0),ytd.get('r20',0)),
            R("  投资活动现金流出小计",         "21", io_,  io_y,  "subtotal"),
            R("  投资活动产生的现金流量净额",   "22", inv_n,inv_ny,"total"),
            # ── 筹资活动 ──
            R("三、筹资活动产生的现金流量：",   "", None, None,   "header"),
            R("  吸收投资收到的现金",           "23", cur.get('r23',0), ytd.get('r23',0)),
            R("  取得借款收到的现金",           "24", cur.get('r24',0), ytd.get('r24',0)),
            R("  收到的其他与筹资活动有关的现金","25",cur.get('r25',0),ytd.get('r25',0)),
            R("  筹资活动现金流入小计",         "26", fi,   fi_y,  "subtotal"),
            R("  偿还债务支付的现金",           "27", cur.get('r27',0), ytd.get('r27',0)),
            R("  分配股利或偿付利息支付的现金", "28", cur.get('r28',0), ytd.get('r28',0)),
            R("  支付的其他与筹资活动有关的现金","29",cur.get('r29',0),ytd.get('r29',0)),
            R("  筹资活动现金流出小计",         "30", fo,   fo_y,  "subtotal"),
            R("  筹资活动产生的现金流量净额",   "31", fin_n,fin_ny,"total"),
            R("四、汇率变动对现金及现金等价物的影响","32",0,0),
            R("五、现金及现金等价物净增加额",   "33", net_cur, net_ytd, "total"),
            R("  加：期初现金及现金等价物余额", "34", cash_open, cash_open),
            R("六、期末现金及现金等价物余额",   "35", cash_end, cash_end, "total"),
            # ── 补充资料分隔 ──
            R("━━━━ 补充资料 ━━━━",            "",  None, None,  "header"),
            R("一、将净利润调节为经营活动现金流量：","", None, None, "header"),
            R("  净利润",                       "1",  np_cur, np_ytd),
            R("  加：资产减值准备",             "2",  0, 0),
            R("  固定资产折旧",                 "3",  0, 0),
            R("  无形资产摊销",                 "4",  0, 0),
            R("  长期待摊费用摊销",             "5",  0, 0),
            R("  处置固定资产损失（收益-）",    "6",  0, 0),
            R("  公允价值变动损失（收益-）",    "8",  0, 0),
            R("  财务费用（收益-）",            "9",  0, 0),
            R("  投资损失（收益-）",            "10", 0, 0),
            R("  经营性应收项目的减少（增加-）","14", 0, ar_chg),
            R("  经营性应付项目的增加（减少-）","15", 0, ap_chg),
            R("  其他",                         "16", 0, 0),
            R("  经营活动产生的现金流量净额",   "17", cn, cn_y, "total"),
            R("三、现金及现金等价物净变动情况：","", None, None, "header"),
            R("  现金的期末余额",               "21", cash_end, cash_end),
            R("  减：现金的期初余额",           "22", cash_open, cash_open),
            R("  现金及现金等价物净增加额",     "25", cash_end - cash_open, cash_end - cash_open, "total"),
        ]

        self.cf_stmt_tbl.setRowCount(len(rows))
        for i, (label, rowno, cur_v, ytd_v, style) in enumerate(rows):
            self.cf_stmt_tbl.setRowHeight(i, 32)
            is_hdr    = (style == "header")
            is_sub    = (style == "subtotal")
            is_tot    = (style == "total")
            bg = QColor(HDR_BG) if is_hdr else QColor(BOLD_BG) if is_sub or is_tot else None

            for j, (text, align) in enumerate([
                (label,  Qt.AlignLeft|Qt.AlignVCenter),
                (rowno,  Qt.AlignCenter),
                (fmt_amt(cur_v) if isinstance(cur_v, (int,float)) else "",
                         Qt.AlignRight|Qt.AlignVCenter),
                (fmt_amt(ytd_v) if isinstance(ytd_v, (int,float)) else "",
                         Qt.AlignRight|Qt.AlignVCenter),
            ]):
                it = QTableWidgetItem(text); it.setTextAlignment(align)
                if is_hdr:
                    it.setBackground(QColor(HDR_BG))
                    if j == 0: it.setForeground(QColor("#3d6fdb"))
                    it.setFont(QFont("", weight=QFont.Bold))
                elif is_sub or is_tot:
                    it.setBackground(QColor(BOLD_BG))
                    it.setFont(QFont("", weight=QFont.Bold))
                if j >= 2 and isinstance(cur_v if j==2 else ytd_v, (int,float)):
                    val = cur_v if j == 2 else ytd_v
                    if val and val < 0:
                        it.setForeground(QColor("#ff4d4f"))
                self.cf_stmt_tbl.setItem(i, j, it)

    def _export_cf_stmt(self):
        if not self.client_id: return
        import openpyxl
        from openpyxl.styles import Font as XFont, Alignment, PatternFill, Border, Side

        end_period = self.rep_end_period.currentData() or self.period
        path, _ = QFileDialog.getSaveFileName(self, "保存",
            f"现金流量表_{end_period}.xlsx", "Excel(*.xlsx)")
        if not path: return
        wb = openpyxl.Workbook(); ws = wb.active; ws.title = "现金流量表"
        hdrs = ["项目","行次","本期金额","本年累计金额"]
        fill_hdr = PatternFill("solid", fgColor="1C2340")
        for ci, h in enumerate(hdrs, 1):
            cell = ws.cell(1, ci, h)
            cell.font = XFont(bold=True, color="FFFFFF"); cell.fill = fill_hdr
            cell.alignment = Alignment(horizontal="center")
        for ri in range(self.cf_stmt_tbl.rowCount()):
            row_vals = []
            for ci in range(4):
                it = self.cf_stmt_tbl.item(ri, ci)
                row_vals.append(it.text() if it else "")
            ws.append(row_vals)
        ws.column_dimensions['A'].width = 45
        for col in ['B','C','D']: ws.column_dimensions[col].width = 16
        wb.save(path); QMessageBox.information(self, "成功", f"已导出:\n{path}")

    def _build_cashflow(self):
        w = QWidget(); L = QVBoxLayout(w); L.setContentsMargins(20,14,20,14); L.setSpacing(8)
        L.addWidget(lbl("收支统计表（本期科目发生额汇总）", bold=True, size=15))
        L.addWidget(lbl("按资产/负债/收入/费用分类展示本期所有科目的借贷发生额", color="#888", size=12))
        self.cf_tbl = self._make_report_table(
            ["科目编号","科目名称","类型","本期借方","本期贷方","净额"],
            [90,-1,70,110,110,110])
        self.cf_tbl.horizontalHeader().setSectionResizeMode(QHeaderView.Interactive)
        self.cf_tbl.horizontalHeader().setSectionResizeMode(1, QHeaderView.Stretch)
        L.addWidget(self.cf_tbl); self.stack.addWidget(w)

    def _load_cashflow(self):
        if not self.client_id: return
        start_period = self.rep_start_period.currentData()
        end_period = self.rep_end_period.currentData()
        if not start_period or not end_period: return
        conn = get_db(); c = conn.cursor()
        c.execute("""SELECT e.account_code, e.account_name,
            SUM(e.debit) td, SUM(e.credit) tc
            FROM voucher_entries e JOIN vouchers v ON v.id=e.voucher_id
            WHERE v.client_id=? AND v.period>=? AND v.period<=? AND v.status='已审核'
            GROUP BY e.account_code ORDER BY e.account_code""",
            (self.client_id, start_period, end_period))
        entries = c.fetchall()
        # Get account types
        c.execute("SELECT code,type FROM accounts WHERE client_id=?",(self.client_id,))
        acct_types = {r[0]:r[1] for r in c.fetchall()}
        conn.close()

        type_colors = {"资产":"#3d6fdb","负债":"#e05252","所有者权益":"#722ed1",
                       "成本":"#fa8c16","收入":"#52c41a","费用":"#eb5757"}

        self.cf_tbl.setRowCount(len(entries))
        td_total = tc_total = 0
        for i,r in enumerate(entries):
            self.cf_tbl.setRowHeight(i,34)
            d=r['td'] or 0; cr=r['tc'] or 0; net=d-cr
            td_total+=d; tc_total+=cr
            atype = acct_types.get(r['account_code'],"")
            tcolor = type_colors.get(atype,"#555")
            vals = [r['account_code'],r['account_name'] or "",atype,fmt_amt(d),fmt_amt(cr),fmt_amt(net)]
            for j,v in enumerate(vals):
                it = QTableWidgetItem(v)
                it.setTextAlignment(Qt.AlignCenter if j!=1 else Qt.AlignLeft|Qt.AlignVCenter)
                if j==2: it.setForeground(QColor(tcolor))
                if j==5: it.setForeground(QColor("#3d6fdb") if net>0 else QColor("#ff4d4f") if net<0 else QColor("#888"))
                self.cf_tbl.setItem(i,j,it)
        # Add totals row
        n = len(entries)
        self.cf_tbl.setRowCount(n+1)
        self.cf_tbl.setRowHeight(n,38)
        for j,v in enumerate(["","合计","",fmt_amt(td_total),fmt_amt(tc_total),fmt_amt(td_total-tc_total)]):
            it = QTableWidgetItem(v); it.setTextAlignment(Qt.AlignCenter if j!=1 else Qt.AlignLeft|Qt.AlignVCenter)
            it.setFont(QFont("",weight=QFont.Bold)); it.setBackground(QColor("#f5f7fa"))
            self.cf_tbl.setItem(n,j,it)

    def set_client(self, client_id, client_name, period):
        self.client_id = client_id; self.period = period
        self.period_lbl.setText(f"【{client_name}】{period}")
        # Initialize period ranges
        self.rep_start_period.clear()
        self.rep_end_period.clear()
        now = datetime.now()
        periods = []
        for y in range(now.year, 2018-1, -1):
            for m in range(12,0,-1):
                period_str = f"{y}-{m:02d}"
                display_str = f"{y}年{m:02d}期"
                periods.append((period_str, display_str))

        for period_str, display_str in periods:
            self.rep_start_period.addItem(display_str, period_str)
            self.rep_end_period.addItem(display_str, period_str)

        # Set default to current period
        current_period = f"{now.year}-{now.month:02d}"
        for i in range(self.rep_start_period.count()):
            if self.rep_start_period.itemData(i) == current_period:
                self.rep_start_period.setCurrentIndex(i)
                self.rep_end_period.setCurrentIndex(i)
                break

        idx = self.stack.currentIndex()
        if idx==0: self._load_balance()
        elif idx==1: self._load_income()
        elif idx==2: self._load_equity()
        elif idx==3: self._load_cf_stmt()
        elif idx==4: self._load_cashflow()

    def _export(self):
        if not self.client_id: return
        import openpyxl
        from openpyxl.styles import Font as XFont, Alignment, PatternFill, Border, Side
        end_period = self.rep_end_period.currentData()
        if not end_period: return
        path,_=QFileDialog.getSaveFileName(self,"保存",f"财务报表_{end_period}.xlsx","Excel(*.xlsx)")
        if not path: return
        wb = openpyxl.Workbook()
        # Income sheet
        ws = wb.active; ws.title="利润表"
        ws.append(["项目","行次","本期金额","本年累计"])
        conn = get_db(); c = conn.cursor()
        start_period = self.rep_start_period.currentData()
        end_period = self.rep_end_period.currentData()
        if not start_period or not end_period:
            QMessageBox.warning(self, "错误", "请选择报告期间")
            return
        c.execute("""SELECT e.account_code,SUM(e.credit)-SUM(e.debit) FROM voucher_entries e
            JOIN vouchers v ON v.id=e.voucher_id WHERE v.client_id=? AND v.period>=? AND v.period<=? GROUP BY e.account_code""",
                  (self.client_id, start_period, end_period))
        cur = {r[0]:r[1] or 0 for r in c.fetchall()}
        def g(code):
            total = 0
            for k, v in cur.items():
                if k == code or k.startswith(code+".") or k.startswith(code+"_"):
                    total += (v or 0)
            return total
        use_6 = bool(g("6001") or g("6401"))
        if use_6:
            income = g("6001") + g("6051")
            cost   = -(g("6401") + g("6402"))
            ops    = income + cost - abs(g("6403")) - abs(g("6601")) - abs(g("6602")) + g("6603") - abs(g("6604"))
            net    = ops + g("6301") - abs(g("6711")) - abs(g("6801"))
        else:
            income = g("5001") + g("5051"); cost = -g("5401") - g("5402")
            ops    = income + cost - abs(g("5501")) - abs(g("5502")) + g("5503")
            net    = ops + g("5301") - abs(g("5601")) - abs(g("5701"))
        ws.append(["营业收入","1",income,""]); ws.append(["营业成本","2",cost,""])
        ws.append(["营业利润","10",ops,""])
        ws.append(["净利润","17",net,""])
        conn.close()
        for col in ws.columns: ws.column_dimensions[col[0].column_letter].width=20
        wb.save(path); QMessageBox.information(self,"成功",f"报表已导出:\n{path}")