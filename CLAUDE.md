# 智一盈小账 · WiseLedger — 项目说明

## 技术栈
Python 3.12 / PySide6 / SQLite / PyInstaller

## Git 操作规范
- **git 提交、建分支等操作一律在 GitHub Desktop 完成**
- VS Code 是 click 级别，无法输入提交信息，不要用它做 git 操作

## 项目结构
- `main.py` — 程序入口
- `pages/` — 8个业务页面
- `dialogs/` — 弹窗
- `print_utils.py` — 报表 PDF 导出（reportlab）
- `db.py` — 数据库（SQLite）
- `session.py` — 登录会话 + RBAC
- `backup_utils.py` / `auto_backup.py` / `kr_utils.py` — 备份相关

## 代码规范
- PySide6 枚举使用完整命名空间：`Qt.AlignmentFlag.AlignCenter`，不用 `Qt.AlignCenter`
- 可空 dict 参数写法：`user: dict | None = None`
- Pylance 配置见 `pyrightconfig.json`
