"""dialogs/__init__.py — 统一导出所有对话框，保持向后兼容"""
from dialogs.client_dialogs  import ClientDialog, AccountInitDialog
from dialogs.voucher_dialogs import VoucherDialog, AuxItemDialog, AuxPage
from dialogs.import_dialogs  import ImportAccountSetDialog
from dialogs.account_dialogs import AccountEditDialog, ImportExcelDialog

__all__ = [
    "ClientDialog", "AccountInitDialog",
    "VoucherDialog", "AuxItemDialog", "AuxPage",
    "ImportAccountSetDialog",
    "AccountEditDialog", "ImportExcelDialog",
]
