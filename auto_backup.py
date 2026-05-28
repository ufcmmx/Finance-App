"""auto_backup.py — 月末自动备份（补跑机制）

每次登录时调用 check_and_run()，检查是否有尚未备份的历史月份，
有则静默执行并写审计日志。用户不需要在月末最后一天开启软件。
"""

import os
import calendar
from datetime import date, datetime


# ── 内部工具 ──────────────────────────────────────────────────────────────

def _last_day_of_month(d: date) -> date:
    """返回指定日期所在月份的最后一天。"""
    return d.replace(day=calendar.monthrange(d.year, d.month)[1])


def _has_auto_backup(ym: str) -> bool:
    """检查 audit_log 中是否已有指定月份（YYYY-MM）的自动备份记录。"""
    try:
        from db import get_db
        conn = get_db(); c = conn.cursor()
        c.execute(
            "SELECT 1 FROM audit_log "
            "WHERE client_id=0 AND action='自动备份' AND target_id=? LIMIT 1",
            (ym,))
        found = c.fetchone() is not None
        conn.close()
        return found
    except Exception:
        return False


# ── 公开接口 ──────────────────────────────────────────────────────────────

def get_target_month() -> str | None:
    """
    返回最近一个需要自动备份的月份字符串（YYYY-MM），或 None（无需备份）。

    规则（往前最多检查 6 个月，取最近一个）：
      - 该月已结束（今天已进入下一个月），或今天恰好是该月最后一天
      - 且 audit_log 中尚无该月的「自动备份」记录

    只返回最近一个，避免长时间未登录时一次创建大量备份文件。
    下次登录再补跑更早的月份。
    """
    today = date.today()

    for i in range(6):
        if i == 0:
            # 当月：仅今天是最后一天时才算到期
            if today != _last_day_of_month(today):
                continue
            ym = today.strftime("%Y-%m")
        else:
            # 往前推 i 个月
            year, month = today.year, today.month - i
            while month <= 0:
                month += 12
                year -= 1
            ym = f"{year}-{month:02d}"

        if not _has_auto_backup(ym):
            return ym

    return None


def check_and_run(parent=None) -> None:
    """
    登录后调用的统一入口。

    流程：
      1. 检查自动备份是否已启用（settings 表）
      2. 检查备份密码是否已设置（keyring）
      3. 找到最近一个需要补跑的月份
      4. 创建备份目录（若不存在）
      5. 静默加密备份 → 写审计日志 → 弹一次完成提示
      6. 任何步骤失败只写审计日志，不弹错误框打扰用户
    """
    from db import get_db, DB_PATH, log_action, get_setting
    from backup_utils import encrypt_backup
    from kr_utils import kr_get

    # ── 前置检查 ──
    if get_setting("auto_backup_enabled", "0") != "1":
        return

    pw = kr_get()
    if not pw:
        return  # 密码未设置，静默跳过

    ym = get_target_month()
    if not ym:
        return  # 所有近期月份均已备份

    # ── 确定备份目录 ──
    backup_dir = get_setting("auto_backup_path", "")
    if not backup_dir:
        # 默认：数据库文件同级的 backups/ 子目录
        backup_dir = os.path.join(os.path.dirname(DB_PATH), "backups")

    try:
        os.makedirs(backup_dir, exist_ok=True)
    except OSError as e:
        _log_failure(ym, f"无法创建备份目录 {backup_dir}：{e}")
        return

    # ── 执行备份 ──
    ts = datetime.now().strftime("%Y%m%d_%H%M%S")
    dest = os.path.join(backup_dir, f"WiseLedger自动备份_{ym}_{ts}.zyac")

    try:
        encrypt_backup(DB_PATH, dest, pw)

        conn = get_db()
        log_action(conn, 0, "自动备份", "system", ym,
                   f"月末自动备份，覆盖月份：{ym}，文件：{dest}")
        conn.commit()
        conn.close()

        # 非阻塞完成提示（主窗口已可见，不会造成闪烁）
        if parent:
            from PySide6.QtWidgets import QMessageBox
            QMessageBox.information(
                parent, "自动备份完成",
                f"已自动完成 {ym} 月末数据备份。\n\n"
                f"保存位置：\n{dest}"
            )

    except Exception as e:
        _log_failure(ym, str(e))


def _log_failure(ym: str, detail: str) -> None:
    """静默写入自动备份失败日志，不向用户弹框。"""
    try:
        from db import get_db, log_action
        conn = get_db()
        log_action(conn, 0, "自动备份失败", "system", ym, detail)
        conn.commit()
        conn.close()
    except Exception:
        pass
