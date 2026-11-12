"""tests/conftest.py — 全局夹具

安全原则（写在最前面，改动前务必保留）：
  1. 每个测试使用独立的临时数据库，绝不读写真实的 accounting.db（DB_PATH 在所有已导入模块里同步替换）。
  2. 全程使用内存里的假 keyring，绝不读写系统真实的凭据管理器 / Keychain。
  3. 所有 QMessageBox / QInputDialog / QFileDialog 都被替换为不阻塞的桩，并记录弹出内容，供断言。

环境变量：
  WL_PROJECT_ROOT   项目根目录（默认：tests 的上一级）
  WL_TEST_GUI=1     使用真实窗口系统（默认 offscreen，不弹窗口）。在 Windows 上建议两种都跑一遍
  WL_TEST_STYLES    逗号分隔的 Qt 样式名，默认取 Windows,Fusion,windowsvista,windows11,macOS 中本机存在的
"""
import os, sys, pathlib
import pytest

ROOT = pathlib.Path(os.environ.get("WL_PROJECT_ROOT") or pathlib.Path(__file__).resolve().parents[1]).resolve()
for p in (str(ROOT), str(ROOT / "tests")):
    if p not in sys.path:
        sys.path.insert(0, p)
if not os.environ.get("WL_TEST_GUI"):
    os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")

from PySide6.QtWidgets import QApplication, QMessageBox, QInputDialog, QFileDialog, QStyleFactory

_app = QApplication.instance() or QApplication(sys.argv[:1])

def _styles():
    want = os.environ.get("WL_TEST_STYLES")
    names = [s.strip() for s in want.split(",")] if want else ["Windows", "Fusion", "windowsvista", "windows11", "macOS"]
    have = {k.lower(): k for k in QStyleFactory.keys()}
    out = []
    for n in names:
        if n.lower() in have and have[n.lower()] not in out:
            out.append(have[n.lower()])
    return out or [QStyleFactory.keys()[0]]

STYLES = _styles()


def pytest_report_header(config):
    import PySide6
    from PySide6.QtCore import qVersion
    return [f"WiseLedger 测试 | 项目根目录: {ROOT}",
            f"Python {sys.version.split()[0]} | PySide6 {PySide6.__version__} | Qt {qVersion()} | 平台插件 {os.environ.get('QT_QPA_PLATFORM', '(系统默认)')}",
            f"参与测试的 Qt 样式: {STYLES}"]


# ───────────────────────── 全局安全夹具 ─────────────────────────
@pytest.fixture(scope="session", autouse=True)
def qapp():
    import utils
    _app.setStyleSheet(utils.SS)          # 与 main.py 一致：直接使用生产样式表
    return _app


class _MemKeyring:
    """内存假 keyring：保证测试不会碰真实凭据管理器。"""
    store = {}


@pytest.fixture(scope="session", autouse=True)
def _fake_keyring():
    try:
        import keyring
        from keyring.backend import KeyringBackend
    except Exception:
        yield None
        return
    class Mem(KeyringBackend):
        priority = 99
        def get_password(self, s, u): return _MemKeyring.store.get((s, u))
        def set_password(self, s, u, p): _MemKeyring.store[(s, u)] = p
        def delete_password(self, s, u): _MemKeyring.store.pop((s, u), None)
    old = keyring.get_keyring()
    keyring.set_keyring(Mem())
    yield Mem
    keyring.set_keyring(old)


@pytest.fixture(autouse=True)
def _clean_keyring():
    _MemKeyring.store.clear()
    yield


class Msgs(list):
    """记录所有被拦截的消息框：(类型, 标题, 内容)"""
    def titles(self): return [m[1] for m in self]
    def has(self, kw): return any(kw in (m[1] + m[2]) for m in self)


@pytest.fixture(autouse=True)
def msgs(monkeypatch):
    rec = Msgs()
    def mk(kind, ret):
        def f(*a, **k):
            title = a[1] if len(a) > 1 else k.get("title", "")
            text = a[2] if len(a) > 2 else k.get("text", "")
            rec.append((kind, str(title), str(text)))
            return ret
        return staticmethod(f)
    monkeypatch.setattr(QMessageBox, "information", mk("info", QMessageBox.Ok))
    monkeypatch.setattr(QMessageBox, "warning", mk("warn", QMessageBox.Yes))
    monkeypatch.setattr(QMessageBox, "critical", mk("crit", QMessageBox.Yes))
    monkeypatch.setattr(QMessageBox, "question", mk("ask", QMessageBox.Yes))
    return rec


@pytest.fixture
def inputs(monkeypatch):
    """QInputDialog.getText 的应答队列：inputs.extend([("密码", True), ...])；队列空时视为用户取消。"""
    q = []
    monkeypatch.setattr(QInputDialog, "getText", staticmethod(lambda *a, **k: q.pop(0) if q else ("", False)))
    return q


@pytest.fixture
def files(monkeypatch):
    """QFileDialog 应答：files.save = 路径；files.open = 路径"""
    class F: save = ""; open = ""; dir = ""
    f = F()
    monkeypatch.setattr(QFileDialog, "getSaveFileName", staticmethod(lambda *a, **k: (f.save, "")))
    monkeypatch.setattr(QFileDialog, "getOpenFileName", staticmethod(lambda *a, **k: (f.open, "")))
    monkeypatch.setattr(QFileDialog, "getExistingDirectory", staticmethod(lambda *a, **k: f.dir))
    return f


@pytest.fixture(autouse=True)
def login(monkeypatch):
    """模拟已登录超级管理员。不依赖 session.py 的登录接口，只替换页面实际调用的几个方法。"""
    import session
    A = session.AppSession
    user = {"id": 1, "username": "admin", "display_name": "超级管理员", "role": "superadmin"}
    for name, fn in [("get", lambda cls: user), ("has_perm", lambda cls, p: True),
                     ("can_access_client", lambda cls, c: True),
                     ("display_name", lambda cls: user["display_name"]),
                     ("role_label", lambda cls: "超级管理员")]:
        monkeypatch.setattr(A, name, classmethod(fn), raising=False)
    return user


@pytest.fixture
def logged_out(monkeypatch):
    import session
    A = session.AppSession
    for name, fn in [("get", lambda cls: None), ("has_perm", lambda cls, p: False),
                     ("can_access_client", lambda cls, c: False),
                     ("display_name", lambda cls: ""), ("role_label", lambda cls: "")]:
        monkeypatch.setattr(A, name, classmethod(fn), raising=False)


@pytest.fixture
def env(tmp_path, monkeypatch, qapp):
    """独立的临时数据库环境。返回 helpers.Env。"""
    import db
    from helpers import Env
    old = db.DB_PATH
    new = str(tmp_path / "test_accounting.db")
    assert new != old
    for m in list(sys.modules.values()):          # 其它模块 from db import DB_PATH 得到的是旧值，必须一并替换
        try:
            if getattr(m, "DB_PATH", None) == old:
                monkeypatch.setattr(m, "DB_PATH", new)
        except Exception:
            pass
    monkeypatch.setattr(db, "DB_PATH", new)
    db.init_db()
    return Env(db, tmp_path)


STD_NAMES = ["企业会计准则", "小企业会计制度"]

@pytest.fixture(params=STD_NAMES, ids=["企业准则", "小企业"])
def std(request):
    return request.param
