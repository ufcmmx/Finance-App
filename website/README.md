# WiseLedger 官网

静态落地页，部署到 Cloudflare Pages，绑定到 `wisdompluscn.com`。

## 本地预览

```bash
cd /Users/rickxu/Finance/website
python3 -m http.server 8000
# 浏览器打开 http://localhost:8000
```

## 部署到 Cloudflare Pages

**方式 A：手动拖拽上传（最快，5 分钟搞定）**

1. Cloudflare Dashboard → **计算** → **Workers 和 Pages**
2. 点 **创建应用程序** → **Pages** → **上传资产**
3. 项目名：`wisdompluscn`
4. 拖拽 `website/` 整个目录（不是 zip）到上传区
5. 部署完成后 → **自定义域** → 添加：
   - `wisdompluscn.com`（根域名）
   - `www.wisdompluscn.com`（可选）

**方式 B：从 GitHub 自动部署（推荐长期）**

1. 把 `website/` 目录 push 到一个 GitHub 仓库（可以就用 Finance 仓库或新建一个）
2. Cloudflare Dashboard → Pages → **创建应用程序** → **连接到 Git**
3. 选择仓库和分支
4. **构建设置**：
   - 框架预设：**None**
   - 构建命令：（留空）
   - 构建输出目录：`website`（如果 index.html 在子目录）或者 `/`
5. 保存并部署
6. 绑定自定义域名

以后 push 代码，网站自动更新。

## 目录结构

```
website/
├── index.html              主页（单页）
├── README.md               本文件
└── assets/
    ├── logo-256.png        Logo 256×256
    └── ui/                 产品截图
        ├── 01-clients-login.png
        ├── 02-vouchers.png
        ├── 03-report.png
        ├── 04-audit.png
        ├── 05-backup.png
        └── 06-opening.png
```

## 更新内容

**改文案：** 直接编辑 `index.html`
- 客服微信号：搜索 "（联系时提供）" 替换
- 价格：搜索 "¥399" / "¥599" 替换
- 邮箱：搜索 "2928314561@qq.com" 替换

**换截图：** 替换 `assets/ui/` 里对应文件（保持文件名不变即可）

**改 logo：** 替换 `assets/logo-256.png`

## 技术栈

- 纯静态 HTML + Tailwind CSS (via CDN)
- 无构建步骤、无框架依赖
- 单页面，所有内容在 index.html 里
