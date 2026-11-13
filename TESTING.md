# WiseLedger 测试方案

> 目标：每次改动后，**用一条命令**判断「能不能发布」；并且在打包之前，就能发现像
> 「Qt 6.12 让金额框看不见数字」这类只在特定 Windows/Qt 版本上出现的问题。

## 1. 原则

1. **测试必须能失败**。每条用例都经过「缺陷注入」验证：把已知 bug 重新放回去，测试必须变红（见第 9 节）。
2. **测真实代码，不在测试里重写业务逻辑**。例如「父级汇总」要通过 `AccountInitDialog` 录入、读取父级控件的值来验证，而不是在测试里自己求和再写库。
3. **带生产样式表、用真人点击时序**。样式表与 Qt 版本的组合、按下→松开之间的定时器，都是这次事故的根源。
4. **绝不碰真实数据**：临时数据库 + 内存假 keyring + 非阻塞对话框（见 `tests/conftest.py` 开头）。
5. **报告必须来自真实运行结果**，不允许预先写好「✅ 通过」的模板。

## 2. 分层

| 层 | 文件 | 内容 | 自动化 |
|---|---|---|---|
| L0 环境/兼容性金丝雀 | `test_00_env_canary.py` | PySide6 版本固定；各 Qt 样式下金额框内部输入区宽度、文字是否真的画出来 | ✅ |
| L1 控件交互 | `test_10_spinbox.py` | 首次点击覆盖、小数、清空、Tab、选区丢失、输入法提交事件、全角/中文句号、粘贴 | ✅ |
| L2 业务逻辑 | `test_30_settle.py` `test_31_reports.py` `test_32_voucher.py` | 结转（含投资亏损/费用冲回）、利润表、资产负债表、**139 个一级科目逐个取数覆盖**、余额表口径、导出一致 | ✅ |
| L3 页面集成 | `test_20_opening.py` `test_32_voucher.py` | 期初弹窗（父级汇总、保存重开）、建账期规则、封账只读、按钮可见性 | ✅ |
| L4 系统功能 | `test_40_system.py` `test_41_auto_backup.py` `test_42_audit.py` | 修改密码页、备份页布局、keyring、备份/恢复、自动备份补跑日期、审计不可删改 | ✅ |
| L5 打包与真机 | 第 6 节清单 | exe 启动、真实输入法、凭据管理器、覆盖升级、安装包来源 | ❌ 人工 |

两套会计制度（企业准则 / 小企业）在 L2、L3 里都参数化跑。

## 3. 固定依赖版本（先做这一步）

这次问题的根源：CI 里 `pip install PySide6` 没有固定版本，某次构建悄悄装到了有缺陷的 Qt 6.12.0。

1. 项目根目录放 `requirements.txt`（已提供样例），**PySide6 用 `==` 固定**，版本取你 Mac 上一直正常的那个。
2. 本地、CI、打包都用 `pip install -r requirements.txt`。
3. `test_00` 会检查：requirements 里是否固定了版本、当前运行版本是否一致、CI 里 `pip install PySide6` 是否带 `==`。

## 4. L0 金丝雀失败了怎么办

现象：`test_spinbox_edit_area_not_collapsed` / `test_spinbox_text_is_actually_painted` 失败，
报告里提示「内部输入区只有 1px」。

原因：Qt 6.12.0 起，样式表里的 `qproperty-buttonSymbols: NoButtons` 与带边框的样式表组合，会把 `QDoubleSpinBox`
内部输入区算成 1 像素。6.6–6.11 正常。

两种处理，任选其一：
- **回退/固定 PySide6 到 ≤ 6.11**（最简单，外观不变）。
- 应用 `patches/utils_stylesheet_qt612.patch`：把那一行改成「保留按钮但宽度为 0」，在 6.8.3 / 6.11 / 6.12 上都已验证。

L0 红的时候，其它层的失败多半是它的连带结果，**先解决 L0**。

## 5. 运行

```bash
pip install -r requirements.txt
python tests/run_all.py            # 全部（约 35 秒）
python tests/run_all.py --fast     # 跳过逐科目覆盖（约 15 秒）
python tests/run_all.py --gui       # 真实窗口系统（Windows 上请再跑一次：真实样式/DPI）
python tests/run_all.py -k spinbox  # 只跑某一类
pytest tests/test_30_settle.py -q   # 也可以直接用 pytest
```

环境变量：`WL_PROJECT_ROOT`（项目根）、`WL_TEST_GUI=1`、`WL_TEST_STYLES=Windows,Fusion,windows11`。
输出：终端 + `tests/report/summary.md` + `tests/report/results.json`。退出码 0=无失败，1=有失败，2=环境问题。

### CI（发布前必须通过）

```yaml
  test:
    runs-on: windows-latest
    steps:
      - uses: actions/checkout@v4
      - uses: actions/setup-python@v5
        with: { python-version: '3.12' }
      - run: pip install -r requirements.txt
      - run: python tests/run_all.py --gui
  build:
    needs: test          # 测试不过，不打包、不上传 R2
    runs-on: windows-latest
    # ……原有步骤；pip install 同样改用 -r requirements.txt
```

## 6. L5：Windows 真机手工清单（自动化覆盖不到）

每次发版前，在**一台干净的 Windows 机器**上用**下载的安装包**（不是源码）逐项确认：

| # | 项 | 通过标准 |
|---|---|---|
| 1 | 安装包来源 | Actions 运行页显示的 commit = 你要发布的 commit；tag 指向同一 commit（`git rev-parse vX^{commit}`） |
| 2 | 科目期初：零值格 | 格里能看到 `0.00`；点击后敲 `22` → `22`；敲 `1.5` → `1.5`；敲 `0.05` → `0.05` |
| 3 | 科目期初：已有金额 | 点击后敲新数字，整体覆盖，不拼接 |
| 4 | **中文输入法** | 微软拼音/搜狗，中文标点状态下敲 `12.5` 得 `12.5`（不是 `125`） |
| 5 | 125% / 150% 显示缩放 | 数字仍清晰可见，不被裁切 |
| 6 | 凭证录入的金额框 | 同 2、3（同一份全局样式表影响所有 spinbox） |
| 7 | 备份密码 | 设置后，在「凭据管理器 → Windows 凭据」能看到「智一会计」；重启后备份无需再输密码 |
| 8 | 备份 → 恢复 | 恢复成功；现行库里有「数据恢复」日志（已知问题见第 7 节） |
| 9 | 覆盖升级 | 用新版本覆盖旧目录（保留 accounting.db）后启动，数据完整 |
| 10 | 排查工具 | `set ZYAC_DEBUG_INPUT=1` 后运行，程序目录出现 `input_debug.log`，首行含 build 标记 |

## 7. 已知问题（测试里标为 xfail：不阻塞发布，但每次运行都会显示）

| 项 | 位置 | 处理 |
|---|---|---|
| 资产负债表还有 7 个一级科目不取数：企业 `1231 坏账准备` `1322 受托代销商品` `1532 未实现融资收益` `1623 公益性生物资产`；小企业 `1231 坏账准备` `4001 生产成本` `4101 制造费用` | `tests/known_issues.py` | 其中 `4001/4101` 应计入存货，制造业客户会用到，优先补；`1231` 是备抵科目，先用「应收账款 + 坏账准备」成对数据复核是否真是缺口 |
| 粘贴 `1234.567`（超过两位小数）被忽略而不是取整 | `test_10_spinbox.py` | 增强项，非缺陷 |

> 2026-11 第一次在真实仓库运行：此前登记的 33 个资产负债表缺口已修复 26 个，「恢复备份后现行库无审计日志」已修复，
> 全角/货币符号粘贴已支持。对应的 XPASS 标记请从测试里删除，让它们变成硬性回归保护。

## 8. 发布门禁

1. L0–L4 **无失败**（xfail 允许；**XPASS 必须清理**）。
2. 第 6 节清单全部勾选，并记录机器、系统版本、PySide6 版本、安装包文件名。
3. CI 中 `test` 通过后才执行打包与上传。

## 9. 测试有效性证明（缺陷注入）

把对话里发现的缺陷逐个重新放回项目，对应测试必须变红：

| 注入的缺陷 | 抓住它的测试 |
|---|---|
| 投资亏损（收入类科目借方余额）漏结转 | `test_30_settle` |
| 利润总额误把营业外支出相加 | `test_31_reports` |
| 短期投资/交易性金融资产不取数 | `test_31_reports` |
| 科目余额表导出崩溃（合并单元格 `column_letter`） | `test_32_voucher` |
| 封账期间仍显示审核类按钮 | `test_32_voucher` |
| 期初弹窗父级汇总被破坏 | `test_20_opening` |
| 金额框没有输入法/全角处理 | `test_10_spinbox` |
| 审计日志页出现清空按钮 | `test_42_audit` |
| 卡片内标签带多余边框 | `test_40_system` |
| **Qt 6.12 样式表压扁金额框** | `test_00_env_canary` |

## 10. 新增测试 / 适配

- 对页面**私有方法**的调用全部集中在 `tests/helpers.py`。页面重构改名时，只改这一个文件。
- 新增会计逻辑测试时，用 `std` 夹具，两套制度各跑一遍；用 `env.standard_month()` 造数据。
- 新增已知问题：在对应用例上加 `@pytest.mark.xfail(reason=..., strict=False)`，并登记到第 7 节。

## 11. 目前没有自动化覆盖的部分（需要补）

- RBAC 权限矩阵：需要读 `session.py` 里角色与权限的定义，才能写出不重复实现的断言。
- 账套导入向导（用友/金蝶 XLS）、凭证录入弹窗的完整流程、PDF 导出。
- PyInstaller 产物的冒烟测试（启动后自动打开主窗口并退出）。
- 现金流量表、所有者权益变动表。

## 12. 如何保证 Mac 和 Windows 用的是同一份代码

**原则：git 是唯一的代码来源；Windows 上只读，不在那边改代码，也不手工拷文件。**
需要改代码时：在 Mac 上改 → 提交 → 推送 → Windows `git pull`。

### Windows 一次性设置（只做一次）

1. 安装 Git for Windows。
2. **不要在旧文件夹里补救**，在一个新位置克隆：
   ```
   git clone <你的仓库地址> D:\finance_win
   cd D:\finance_win
   python -m venv .venv
   .venv\Scripts\activate
   pip install -r requirements.txt
   ```
3. 旧文件夹改名保留（如 `finance_old`），先别删。里面如果有只存在于本地、没提交到 git 的文件
   （例如自己的 `accounting.db`、授权/配置文件），手工拷到新文件夹对应位置。
   在 Mac 上执行 `git status --ignored` 可以看到哪些文件没有被 git 管理。
4. 仓库根目录放一个 `.gitattributes`，让两边换行符一致：
   ```
   *.py   text eol=lf
   *.yml  text eol=lf
   *.md   text eol=lf
   *.txt  text eol=lf
   ```
5. `.gitignore` 里确认有这些（不要提交运行产物）：
   ```
   .venv/
   __pycache__/
   tests/report/
   build/
   dist/
   startup.log
   input_debug.log
   ```

### 每次测试的固定动作

| 机器 | 命令 |
|---|---|
| Mac | `git add -A && git commit -m "..." && git push` |
| Windows | `git pull` → `python tests\run_all.py --gui` |
| 任一台 | 报告头部会打印 **git 提交** 和 **代码指纹**；有未提交改动时会显示 ⚠️ |

### 对比两台机器的结果

每台机器运行后，`tests/report/` 里会生成 `results_<系统>.json`（Mac=`results_Darwin.json`，
Windows=`results_Windows.json`），两台互不覆盖。把两份放到一起：

```
python tests/compare_runs.py results_Darwin.json results_Windows.json
```

它先回答「代码是不是同一份」（逐文件指出哪个不同；换行符差异不算不同），再列出
**只在某一台失败**的测试——那些就是平台本身造成的问题。
代码不同时，结果不能直接比较，先让两边 `git pull` 到同一个提交。

### 关于「report 会不会被覆盖」

- **`tests/report/`（测试结果目录）**：每次运行都会重写，但现在按系统分开命名，Mac 和 Windows 不会互相覆盖。
  它不进 git。想长期留存就自己复制一份。
- **`pages/report.py`（应用代码）**：解压测试压缩包**不会碰它**（压缩包里只有 `tests/`、`pytest.ini`、`TESTING.md`、
  `requirements.txt`、`patches/`、`skills/`）。`git pull` 会把它更新成已提交的版本，这正是目的；
  如果 Windows 上有未提交的本地改动，git 会**拒绝覆盖并报错**，不会静默丢失。
