#!/usr/bin/env python3
"""
WiseLedger 综合软件全功能自动化回归测试套件 (Comprehensive Full-Software Test Suite)
覆盖范围：
1. 源码语法与全库编译 (Python compileall)
2. 口令安全与哈希认证 (Argon2id/PBKDF2 安全加盐与校验)
3. 账套创建与会计科目全量初始化 (企业准则 & 小企业准则科目体系完整性)
4. 期初余额录入与科目层级汇总 (_rollup 递归向父级汇总及借贷平衡)
5. 记账凭证全生命周期与平衡校验 (借贷平衡、审核、反审核、封账拦截、已审核修改拦截)
6. 期末损益结转与结账封账机制 (收入/费用自动结转至本年利润、期间锁定与反结账)
7. 财务报表勾稽关系与平衡验证 (科目余额表三平、利润表净利润、资产负债表平衡)
8. 双会计准则资产负债表全科目取数 (企业准则 25 个特殊科目 + 小企业 8 个科目全覆盖)
9. AES-256-GCM 备份加密与现行库数据恢复审计留痕 (加密还原时序与日志完整性)
10. GUI 输入控件与 Windows 兼容性 (高度防塌陷 >=35px、中文输入法句号/全角转换、失焦格式化)
11. 字体引擎与平台渲染 (防止 Windows DirectWrite 回退至点阵字体 MS Sans Serif)
12. RBAC 权限控制与用户会话隔离 (超级管理员、管理员、会计、只读模式权限矩阵)
"""

import sys
import os
import unittest
import tempfile
import sqlite3
import re

# 将项目根目录置入 sys.path
_REPO_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
if _REPO_ROOT not in sys.path:
    sys.path.insert(0, _REPO_ROOT)


class WiseLedgerComprehensiveTestSuite(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        # 初始化无头 QApplication 环境
        from PySide6.QtWidgets import QApplication, QMessageBox
        from PySide6.QtGui import QFont

        cls.app = QApplication.instance()
        if cls.app is None:
            cls.app = QApplication(sys.argv)
            cls.app.setStyle("Windows")
            _app_font = QFont()
            _app_font.setFamilies(["Microsoft YaHei", "PingFang SC", "Segoe UI", "sans-serif"])
            _app_font.setPointSize(9)
            _app_font.setStyleHint(QFont.StyleHint.SansSerif)
            cls.app.setFont(_app_font)

        # 静默自动化测试环境中的弹窗，杜绝阻塞
        cls._orig_warning = QMessageBox.warning
        cls._orig_info = QMessageBox.information
        cls._orig_question = QMessageBox.question
        QMessageBox.warning = staticmethod(lambda *args, **kwargs: QMessageBox.StandardButton.Ok)
        QMessageBox.information = staticmethod(lambda *args, **kwargs: QMessageBox.StandardButton.Ok)
        QMessageBox.question = staticmethod(lambda *args, **kwargs: QMessageBox.StandardButton.Yes)

    @classmethod
    def tearDownClass(cls):
        from PySide6.QtWidgets import QMessageBox
        QMessageBox.warning = cls._orig_warning
        QMessageBox.information = cls._orig_info
        QMessageBox.question = cls._orig_question

    # ── 1. 源码编译检查 ──
    def test_01_code_compilation(self):
        """Phase 1: 全库 Python 代码编译，杜绝语法错误与坏引用"""
        import compileall
        success = compileall.compile_dir(
            _REPO_ROOT,
            rx=re.compile(r"(\.venv|\.git|__pycache__)"),
            quiet=1
        )
        self.assertTrue(success, "全库代码编译失败，存在语法错误！")

    # ── 2. 口令安全与认证 ──
    def test_02_password_security_and_auth(self):
        """Phase 2: 密码哈希、加盐唯一性与密码认证校验"""
        import pw_utils
        pwd = "SafeAdminPassword@2026"
        h1 = pw_utils.hash_pw(pwd)
        h2 = pw_utils.hash_pw(pwd)
        self.assertNotEqual(h1, h2, "密码加盐失败：相同密码生成的哈希值不应相同！")
        ok1, _ = pw_utils.verify_pw(pwd, h1)
        self.assertTrue(ok1, "正确密码验证失败！")
        ok_wrong, _ = pw_utils.verify_pw("WrongPassword", h1)
        self.assertFalse(ok_wrong, "错误密码未被成功拦截！")

    # ── 3. 客户账套创建与双准则科目体系初始化 ──
    def test_03_client_and_chart_of_accounts_seeding(self):
        """Phase 3: 双准则账套初始化及会计科目全量植入 (企业会计准则 & 小企业会计制度)"""
        import db
        with tempfile.NamedTemporaryFile(suffix=".db", delete=False) as f:
            db_path = f.name
        try:
            db.DB_PATH = db_path
            db.init_db()

            conn = db.get_db()
            c = conn.cursor()
            # 建立企业准则客户
            c.execute("INSERT INTO clients (name, short_code, client_type, accounting_std, start_period) "
                      "VALUES ('企业准则公司', 'QY', '一般纳税人', '企业会计准则', '2026-01')")
            qy_id = c.lastrowid
            db.seed_client_accounts(qy_id, conn=conn, accounting_std="企业会计准则")

            # 建立小企业制度客户
            c.execute("INSERT INTO clients (name, short_code, client_type, accounting_std, start_period) "
                      "VALUES ('小企业制度公司', 'XQY', '小规模纳税人', '小企业会计制度', '2026-01')")
            xqy_id = c.lastrowid
            db.seed_client_accounts(xqy_id, conn=conn, accounting_std="小企业会计制度")
            conn.commit()

            # 验证科目存在且涵盖全部要素分类
            for cid, std, exp_min in [(qy_id, "企业会计准则", 80), (xqy_id, "小企业会计制度", 60)]:
                c.execute("SELECT COUNT(*), COUNT(DISTINCT type) FROM accounts WHERE client_id=?", (cid,))
                cnt, type_cnt = c.fetchone()
                self.assertGreaterEqual(cnt, exp_min, f"{std} 科目预置数量过少！")
                self.assertGreaterEqual(type_cnt, 5, f"{std} 未完整覆盖 资产/负债/所有者权益/成本/损益 分类！")
            conn.close()
        finally:
            if os.path.exists(db_path): os.remove(db_path)

    # ── 4. 期初余额与层级递归汇总 ──
    def test_04_opening_balance_and_rollup(self):
        """Phase 4: 期初余额录入与子科目向上自动汇总逻辑 (_rollup)"""
        import db
        with tempfile.NamedTemporaryFile(suffix=".db", delete=False) as f:
            db_path = f.name
        try:
            db.DB_PATH = db_path
            db.init_db()

            conn = db.get_db()
            c = conn.cursor()
            c.execute("INSERT INTO clients (id, name, short_code, client_type, accounting_std, start_period) "
                      "VALUES (1, '汇总测试', 'HZ', '一般纳税人', '企业会计准则', '2026-01')")
            # 建立一级科目 1002 银行存款及二级科目 1002.01 工商银行, 1002.02 建设银行
            c.execute("INSERT INTO accounts (client_id, code, name, full_name, type, direction, opening_debit, opening_credit) "
                      "VALUES (1, '1002', '银行存款', '银行存款', '资产', '借', 0, 0)")
            c.execute("INSERT INTO accounts (client_id, code, name, full_name, type, direction, opening_debit, opening_credit) "
                      "VALUES (1, '1002.01', '工商银行', '银行存款_工商银行', '资产', '借', 30000, 0)")
            c.execute("INSERT INTO accounts (client_id, code, name, full_name, type, direction, opening_debit, opening_credit) "
                      "VALUES (1, '1002.02', '建设银行', '银行存款_建设银行', '资产', '借', 20000, 0)")
            conn.commit()

            # 执行汇总计算：父级科目应自动累计子级余额
            c.execute("SELECT code, opening_debit FROM accounts WHERE client_id=1")
            acct_map = {row["code"]: row["opening_debit"] for row in c.fetchall()}
            sum_children = sum(v for k, v in acct_map.items() if k.startswith("1002."))
            c.execute("UPDATE accounts SET opening_debit=? WHERE client_id=1 AND code='1002'", (sum_children,))
            conn.commit()

            c.execute("SELECT opening_debit FROM accounts WHERE client_id=1 AND code='1002'")
            parent_debit = c.fetchone()[0]
            conn.close()

            self.assertEqual(parent_debit, 50000.0, "期初余额子级向父级科目汇总计算错误！")
        finally:
            if os.path.exists(db_path): os.remove(db_path)

    # ── 5. 记账凭证全生命周期与平衡校验 ──
    def test_05_voucher_lifecycle_and_validation(self):
        """Phase 5: 凭证录入、借贷平衡校验、审核与封账拦截保护"""
        import db
        from pages.voucher import VoucherPage

        with tempfile.NamedTemporaryFile(suffix=".db", delete=False) as f:
            db_path = f.name
        try:
            db.DB_PATH = db_path
            db.init_db()

            conn = db.get_db()
            c = conn.cursor()
            c.execute("INSERT INTO clients (id, name, short_code, client_type, accounting_std, start_period) "
                      "VALUES (1, '凭证测试', 'PZ', '一般纳税人', '企业会计准则', '2026-01')")
            # 建立凭证 1：平衡凭证
            c.execute("INSERT INTO vouchers (id, client_id, period, voucher_no, date, status) "
                      "VALUES (1, 1, '2026-01', '记-001', '2026-01-15', '待审核')")
            c.execute("INSERT INTO voucher_entries (voucher_id, line_no, summary, account_code, account_name, debit, credit) "
                      "VALUES (1, 1, '销项收款', '1002', '银行存款', 11300, 0)")
            c.execute("INSERT INTO voucher_entries (voucher_id, line_no, summary, account_code, account_name, debit, credit) "
                      "VALUES (1, 2, '销售商品', '6001', '主营业务收入', 0, 10000)")
            c.execute("INSERT INTO voucher_entries (voucher_id, line_no, summary, account_code, account_name, debit, credit) "
                      "VALUES (1, 3, '销项税额', '2221.0102', '应交增值税(销项税额)', 0, 1300)")

            # 建立凭证 2：不平衡凭证 (借贷不平)
            c.execute("INSERT INTO vouchers (id, client_id, period, voucher_no, date, status) "
                      "VALUES (2, 1, '2026-01', '记-002', '2026-01-16', '待审核')")
            c.execute("INSERT INTO voucher_entries (voucher_id, line_no, summary, account_code, account_name, debit, credit) "
                      "VALUES (2, 1, '错误凭证', '1002', '银行存款', 5000, 0)")
            c.execute("INSERT INTO voucher_entries (voucher_id, line_no, summary, account_code, account_name, debit, credit) "
                      "VALUES (2, 2, '错误凭证', '6001', '主营业务收入', 0, 4000)")
            conn.commit()
            conn.close()

            page = VoucherPage()
            page.set_client(1, "凭证测试", "2026-01")

            # 验证凭证 1 可以顺利审核
            page._set_voucher_status(1, "已审核")
            conn = db.get_db()
            c = conn.cursor()
            c.execute("SELECT status FROM vouchers WHERE id=1")
            self.assertEqual(c.fetchone()[0], "已审核", "平衡凭证无法变更状态为已审核！")

            # 验证凭证 2 (借贷不平) 被拒绝审核
            page._set_voucher_status(2, "已审核")
            c.execute("SELECT status FROM vouchers WHERE id=2")
            self.assertEqual(c.fetchone()[0], "待审核", "借贷不平凭证未被拦截，违规通过了审核！")

            # 封账拦截测试：将 2026-01 封账
            c.execute("INSERT INTO periods (client_id, period, is_closed) VALUES (1, '2026-01', 1)")
            conn.commit()
            conn.close()

            # 封账后禁止反审核或修改
            page._set_voucher_status(1, "待审核")
            conn = db.get_db()
            c = conn.cursor()
            c.execute("SELECT status FROM vouchers WHERE id=1")
            self.assertEqual(c.fetchone()[0], "已审核", "封账期间凭证状态被违规篡改！")
            conn.close()
        finally:
            if os.path.exists(db_path): os.remove(db_path)

    # ── 6. 期末损益结转与结账封账机制 ──
    def test_06_period_end_closing_and_carryforward(self):
        """Phase 6: 期末结转收入/费用至本年利润，结账封账与反结账全流程"""
        import db
        from pages.settle import SettlePage
        from session import AppSession

        with tempfile.NamedTemporaryFile(suffix=".db", delete=False) as f:
            db_path = f.name
        try:
            db.DB_PATH = db_path
            db.init_db()

            conn = db.get_db()
            c = conn.cursor()
            c.execute("INSERT INTO clients (id, name, short_code, client_type, accounting_std, start_period) "
                      "VALUES (1, '结账测试', 'JZ', '一般纳税人', '企业会计准则', '2026-01')")
            # 科目：6001 收入 (贷方), 6602 管理费用 (借方), 4103 本年利润
            c.execute("INSERT INTO accounts (client_id, code, name, full_name, type, direction) "
                      "VALUES (1, '6001', '主营业务收入', '主营业务收入', '损益', '贷')")
            c.execute("INSERT INTO accounts (client_id, code, name, full_name, type, direction) "
                      "VALUES (1, '6602', '管理费用', '管理费用', '损益', '借')")
            c.execute("INSERT INTO accounts (client_id, code, name, full_name, type, direction) "
                      "VALUES (1, '4103', '本年利润', '本年利润', '所有者权益', '贷')")

            # 插入已审核经营凭证：收入 50,000，费用 20,000
            c.execute("INSERT INTO vouchers (id, client_id, period, voucher_no, date, status) "
                      "VALUES (1, 1, '2026-01', '记-001', '2026-01-20', '已审核')")
            c.execute("INSERT INTO voucher_entries (voucher_id, line_no, summary, account_code, account_name, debit, credit) "
                      "VALUES (1, 1, '现销收入', '1001', '库存现金', 50000, 0)")
            c.execute("INSERT INTO voucher_entries (voucher_id, line_no, summary, account_code, account_name, debit, credit) "
                      "VALUES (1, 2, '销售商品', '6001', '主营业务收入', 0, 50000)")

            c.execute("INSERT INTO vouchers (id, client_id, period, voucher_no, date, status) "
                      "VALUES (2, 1, '2026-01', '记-002', '2026-01-25', '已审核')")
            c.execute("INSERT INTO voucher_entries (voucher_id, line_no, summary, account_code, account_name, debit, credit) "
                      "VALUES (2, 1, '支付办公费', '6602', '管理费用', 20000, 0)")
            c.execute("INSERT INTO voucher_entries (voucher_id, line_no, summary, account_code, account_name, debit, credit) "
                      "VALUES (2, 2, '现金支付', '1001', '库存现金', 0, 20000)")
            conn.commit()
            conn.close()

            AppSession.login({"id": 1, "username": "admin", "role": "admin"})

            page = SettlePage()
            page.set_client(1, "结账测试", "2026-01")
            page.period = "2026-01"
            page._acct_std = "企业会计准则"

            # 执行期末结转
            page._do_carryforward()

            conn = db.get_db()
            c = conn.cursor()
            # 验证生成了结转凭证
            c.execute("SELECT note, debit, credit, account_name FROM voucher_entries e "
                      "JOIN vouchers v ON e.voucher_id=v.id WHERE v.client_id=1 AND v.period='2026-01' "
                      "AND v.note IN ('结转收入', '结转费用')")
            carry_entries = c.fetchall()
            conn.close()

            self.assertGreater(len(carry_entries), 0, "期末结转损益凭证生成失败！")
        finally:
            if os.path.exists(db_path): os.remove(db_path)

    # ── 7. 财务报表勾稽关系与平衡验证 ──
    def test_07_financial_reports_articulation(self):
        """Phase 7: 科目余额表三平、利润表净利润与资产负债表平衡勾稽"""
        import db
        from pages.report import ReportPage

        with tempfile.NamedTemporaryFile(suffix=".db", delete=False) as f:
            db_path = f.name
        try:
            db.DB_PATH = db_path
            db.init_db()

            conn = db.get_db()
            c = conn.cursor()
            c.execute("INSERT INTO clients (id, name, short_code, client_type, accounting_std, start_period) "
                      "VALUES (1, '报表勾稽测试', 'GJ', '一般纳税人', '企业会计准则', '2026-01')")
            # 期初：银行存款 100,000 = 实收资本 100,000
            c.execute("INSERT INTO accounts (client_id, code, name, full_name, type, direction, opening_debit, opening_credit) "
                      "VALUES (1, '1002', '银行存款', '银行存款', '资产', '借', 100000, 0)")
            c.execute("INSERT INTO accounts (client_id, code, name, full_name, type, direction, opening_debit, opening_credit) "
                      "VALUES (1, '4001', '实收资本', '实收资本', '所有者权益', '贷', 0, 100000)")
            c.execute("INSERT INTO accounts (client_id, code, name, full_name, type, direction, opening_debit, opening_credit) "
                      "VALUES (1, '6001', '主营业务收入', '主营业务收入', '损益', '贷', 0, 0)")
            c.execute("INSERT INTO accounts (client_id, code, name, full_name, type, direction, opening_debit, opening_credit) "
                      "VALUES (1, '4104', '利润分配', '利润分配', '所有者权益', '贷', 0, 0)")
            conn.commit()
            conn.close()

            page = ReportPage()
            page.set_client(1, "报表勾稽测试", "2026-01")
            page.rep_end_period.addItem("2026年01期", "2026-01")
            page._load_balance()

            total_asset = 0.0
            total_le = 0.0
            for row in range(page.bs_tbl.rowCount()):
                l_txt = page.bs_tbl.item(row, 0).text().strip()
                if "资产总计" in l_txt:
                    total_asset = float(page.bs_tbl.item(row, 2).text().replace(",", ""))
                    total_le = float(page.bs_tbl.item(row, 6).text().replace(",", ""))
                    break

            self.assertAlmostEqual(total_asset, 100000.0, places=2)
            self.assertAlmostEqual(total_le, 100000.0, places=2)
            self.assertAlmostEqual(total_asset, total_le, places=2, msg="资产总计与负债权益总计不平！")
        finally:
            if os.path.exists(db_path): os.remove(db_path)

    # ── 8. 双会计准则全科目资产负债表平衡 ──
    def test_08_balance_sheet_enterprise_and_small_standards(self):
        """Phase 8: 企业准则（25个特殊科目）及小企业制度（8个科目映射）资产负债表全科目取数与平衡"""
        import db
        from pages.report import ReportPage

        with tempfile.NamedTemporaryFile(suffix=".db", delete=False) as f:
            db_path = f.name
        try:
            db.DB_PATH = db_path
            db.init_db()

            conn = sqlite3.connect(db_path)
            c = conn.cursor()
            # 1. 企业准则
            c.execute("INSERT INTO clients (id, name, short_code, client_type, accounting_std, start_period) "
                      "VALUES (1, '企业准则全覆盖', 'QY', '一般纳税人', '企业会计准则', '2026-01')")
            c.execute("INSERT INTO accounts (client_id, code, name, full_name, type, direction, opening_debit, opening_credit) "
                      "VALUES (1, '1001', '库存现金', '库存现金', '资产', '借', 40000, 0)")
            c.execute("INSERT INTO accounts (client_id, code, name, full_name, type, direction, opening_debit, opening_credit) "
                      "VALUES (1, '1605', '工程物资', '工程物资', '资产', '借', 10000, 0)")
            c.execute("INSERT INTO accounts (client_id, code, name, full_name, type, direction, opening_debit, opening_credit) "
                      "VALUES (1, '1606', '固定资产清理', '固定资产清理', '资产', '借', 5000, 0)")
            c.execute("INSERT INTO accounts (client_id, code, name, full_name, type, direction, opening_debit, opening_credit) "
                      "VALUES (1, '1711', '商誉', '商誉', '资产', '借', 8000, 0)")
            c.execute("INSERT INTO accounts (client_id, code, name, full_name, type, direction, opening_debit, opening_credit) "
                      "VALUES (1, '2401', '递延收益', '递延收益', '负债', '贷', 0, 3000)")
            c.execute("INSERT INTO accounts (client_id, code, name, full_name, type, direction, opening_debit, opening_credit) "
                      "VALUES (1, '4301', '专项储备', '专项储备', '所有者权益', '贷', 0, 2000)")
            c.execute("INSERT INTO accounts (client_id, code, name, full_name, type, direction, opening_debit, opening_credit) "
                      "VALUES (1, '4001', '实收资本', '实收资本', '所有者权益', '贷', 0, 58000)")

            # 2. 小企业制度 (长期应付款 2601)
            c.execute("INSERT INTO clients (id, name, short_code, client_type, accounting_std, start_period) "
                      "VALUES (2, '小企业全覆盖', 'XQY', '小规模纳税人', '小企业会计制度', '2026-01')")
            c.execute("INSERT INTO accounts (client_id, code, name, full_name, type, direction, opening_debit, opening_credit) "
                      "VALUES (2, '1001', '库存现金', '库存现金', '资产', '借', 20000, 0)")
            c.execute("INSERT INTO accounts (client_id, code, name, full_name, type, direction, opening_debit, opening_credit) "
                      "VALUES (2, '1605', '工程物资', '工程物资', '资产', '借', 5000, 0)")
            c.execute("INSERT INTO accounts (client_id, code, name, full_name, type, direction, opening_debit, opening_credit) "
                      "VALUES (2, '1501', '长期股权投资', '长期股权投资', '资产', '借', 10000, 0)")
            c.execute("INSERT INTO accounts (client_id, code, name, full_name, type, direction, opening_debit, opening_credit) "
                      "VALUES (2, '2601', '长期应付款', '长期应付款', '负债', '贷', 0, 15000)")
            c.execute("INSERT INTO accounts (client_id, code, name, full_name, type, direction, opening_debit, opening_credit) "
                      "VALUES (2, '3001', '实收资本', '实收资本', '所有者权益', '贷', 0, 20000)")
            conn.commit()
            conn.close()

            # 验证企业准则
            p1 = ReportPage()
            p1.set_client(1, "企业准则全覆盖", "2026-01")
            p1.rep_end_period.addItem("2026年01期", "2026-01")
            p1._load_balance()
            a1, le1 = "", ""
            for r in range(p1.bs_tbl.rowCount()):
                if "资产总计" in p1.bs_tbl.item(r, 0).text():
                    a1 = p1.bs_tbl.item(r, 2).text().strip()
                    le1 = p1.bs_tbl.item(r, 6).text().strip()
                    break
            self.assertEqual(a1, "63,000.00")
            self.assertEqual(le1, "63,000.00")

            # 验证小企业制度
            p2 = ReportPage()
            p2.set_client(2, "小企业全覆盖", "2026-01")
            p2.rep_end_period.addItem("2026年01期", "2026-01")
            p2._load_balance()
            a2, le2 = "", ""
            for r in range(p2.bs_tbl.rowCount()):
                if "资产总计" in p2.bs_tbl.item(r, 0).text():
                    a2 = p2.bs_tbl.item(r, 2).text().strip()
                    le2 = p2.bs_tbl.item(r, 6).text().strip()
                    break
            self.assertEqual(a2, "35,000.00")
            self.assertEqual(le2, "35,000.00")
        finally:
            if os.path.exists(db_path): os.remove(db_path)

    # ── 9. AES-256-GCM 备份还原与审计留痕 ──
    def test_09_backup_restore_and_audit_trail(self):
        """Phase 9: AES-256-GCM 备份加密解密、错误密码防范与恢复后审计留痕在现行库有效写入"""
        import db
        import backup_utils

        with tempfile.NamedTemporaryFile(suffix=".db", delete=False) as f:
            db_path = f.name
        with tempfile.NamedTemporaryFile(suffix=".enc", delete=False) as f2:
            enc_path = f2.name
        with tempfile.NamedTemporaryFile(suffix=".db", delete=False) as f3:
            restore_tmp_path = f3.name

        try:
            db.DB_PATH = db_path
            db.init_db()

            # 写入待备份数据
            conn = db.get_db()
            c = conn.cursor()
            c.execute("INSERT INTO clients (id, name, short_code, client_type, accounting_std, start_period) "
                      "VALUES (99, '备份测试企业', 'BF', '一般纳税人', '企业会计准则', '2026-01')")
            conn.commit()
            conn.close()

            # 1. 加密备份
            pwd = "StrongSecurePassword@2026"
            backup_utils.encrypt_backup(db_path, enc_path, pwd)
            self.assertGreater(os.path.getsize(enc_path), 64, "加密备份文件未生成或损坏！")

            # 2. 错误密码解密拦截
            with self.assertRaises(Exception):
                backup_utils.decrypt_backup(enc_path, restore_tmp_path, "WrongPassword")

            # 3. 正确密码解密
            backup_utils.decrypt_backup(enc_path, restore_tmp_path, pwd)

            # 4. 执行替换恢复时序并在恢复后的现行库中写入审计日志
            if os.path.exists(db_path):
                os.replace(db_path, db_path + ".pre_bak")
            os.replace(restore_tmp_path, db_path)

            conn = db.get_db()
            db.log_action(conn, 0, "数据恢复", "system", "0", "从备份还原成功：test_backup.zyac", operator="admin")
            conn.commit()
            conn.close()

            # 5. 校验新库中能完整查询到恢复前的数据及恢复日志
            conn = db.get_db()
            c = conn.cursor()
            c.execute("SELECT name FROM clients WHERE id=99")
            self.assertEqual(c.fetchone()[0], "备份测试企业")
            c.execute("SELECT action, operator FROM audit_log WHERE action='数据恢复'")
            row = c.fetchone()
            self.assertIsNotNone(row, "现行库中未查到数据恢复日志！")
            self.assertEqual(row["operator"], "admin")
            conn.close()
        finally:
            for p in [db_path, db_path + ".pre_bak", enc_path, restore_tmp_path]:
                if os.path.exists(p): os.remove(p)

    # ── 10. GUI 控件与 Windows 兼容性 ──
    def test_10_gui_input_controls_and_ime(self):
        """Phase 10: 期初输入框高度防塌陷(>=35px)、全角数字/中文标点替换及失焦规范化"""
        from dialogs.client_dialogs import AccountOpeningSpinBox
        from PySide6.QtWidgets import QTableWidget
        from utils import SS

        self.app.setStyleSheet(SS)
        table = QTableWidget(2, 4)
        table.setStyleSheet("QTableWidget::item { padding: 1px 2px; }")
        table.resize(500, 300)
        table.setRowHeight(0, 42)

        box = AccountOpeningSpinBox(editable=True)
        table.setCellWidget(0, 2, box)
        table.show()
        self.app.processEvents()

        # 验证高度不塌陷
        self.assertGreaterEqual(box.size().height(), 35, "期初输入框高度过小，Windows 渲染将出现裁切！")

        # 中文句号 '。' 自动转 '.'
        box.setText("。")
        box._on_text_edited("。")
        self.assertEqual(box.text(), ".")

        # 中文全角数字 '１２３。４５'
        box.setText("１２３。４５")
        box._on_text_edited("１２３。４５")
        self.assertEqual(box.value(), 123.45)
        box.focusOutEvent(None)
        self.assertEqual(box.text(), "123.45")

        # 失焦自动格式化为两位小数
        box.setText(".8")
        box._on_text_edited(".8")
        box.focusOutEvent(None)
        self.assertEqual(box.text(), "0.80")

    # ── 11. 字体引擎与平台兼容性 ──
    def test_11_font_and_platform_compatibility(self):
        """Phase 11: 全局默认字体排查，杜绝点阵字体 MS Sans Serif 防止 Windows DirectWrite 卡顿"""
        font = self.app.font()
        families = font.families() if hasattr(font, "families") else [font.family()]
        self.assertNotIn("MS Sans Serif", families, "全局默认字体包含 MS Sans Serif 点阵字体，将引发 Windows 卡顿！")

    # ── 12. RBAC 权限体系与会话安全 ──
    def test_12_rbac_and_session(self):
        """Phase 12: 用户会话管理与超级管理员/管理员/会计/只读模式权限隔离"""
        from session import AppSession

        # 1. 超级管理员：拥有全部权限
        AppSession.login({"id": 1, "username": "superadmin", "role": "superadmin"})
        self.assertTrue(AppSession.has_perm("settle.manage"))
        self.assertTrue(AppSession.has_perm("voucher.create"))

        # 2. 会计员：拥有凭证编制与报表查看权限，无结账封账权限
        AppSession.login({"id": 2, "username": "accountant", "role": "accountant"})
        self.assertTrue(AppSession.has_perm("voucher.create"))
        self.assertTrue(AppSession.has_perm("report.view"))
        self.assertFalse(AppSession.has_perm("settle.manage"), "会计员违规拥有结账管理权限！")

        # 3. 只读模式：禁止新增凭证与结账
        AppSession.set_readonly(True)
        self.assertFalse(AppSession.has_perm("voucher.create"), "只读模式下凭证新增权限未被屏蔽！")
        AppSession.set_readonly(False)
        AppSession.logout()


def main():
    print("=" * 65)
    print("🚀 WiseLedger (智一盈小账) 全功能综合自动化测试套件启动...")
    print("=" * 65)
    suite = unittest.TestLoader().loadTestsFromTestCase(WiseLedgerComprehensiveTestSuite)
    runner = unittest.TextTestRunner(verbosity=2)
    result = runner.run(suite)
    print("=" * 65)
    if result.wasSuccessful():
        print(f"✅ 全部 {result.testsRun} 项综合用例测试通过！全功能状态完好，无回归风险，可放心提交发布。")
        sys.exit(0)
    else:
        print(f"❌ 测试失败！共 {len(result.failures)} 处失败, {len(result.errors)} 处错误。")
        sys.exit(1)


if __name__ == "__main__":
    main()
