# WiseLedger License Server

Cloudflare Workers + KV 实现的激活授权后端。

## 项目结构

```
license-server/
├── .env                  # 敏感配置（admin key / 签名密钥），不进 git
├── .gitignore
├── public_key.b64        # 客户端嵌入用的 Ed25519 公钥（base64）
├── wrangler.toml         # Cloudflare Workers 配置
├── src/
│   └── index.js          # Worker 后端代码
├── gen_license.py        # 本地工具：生成激活码
└── README.md             # 本文件
```

## 首次部署步骤（一次性）

```bash
# 1. 装 wrangler
npm install -g wrangler

# 2. 登录 Cloudflare
wrangler login

# 3. 进项目目录
cd /Users/rickxu/Finance/license-server

# 4. 创建 KV 命名空间
wrangler kv namespace create LICENSES
# 复制输出的 id，替换 wrangler.toml 里的 PLACEHOLDER_REPLACE_AFTER_KV_CREATE

# 5. 上传两个 Secret（从 .env 复制）
wrangler secret put ADMIN_KEY
# 粘贴 .env 里的 ADMIN_KEY

wrangler secret put ED25519_PRIVATE_KEY_B64
# 粘贴 .env 里的 ED25519_PRIVATE_KEY_B64

# 6. 部署
wrangler deploy
# 部署成功会显示访问 URL，类似：
# https://wiseledger-license.xxxx.workers.dev

# 7. 把这个 URL 写到 .env 的 WORKERS_URL
```

## 卖货生成激活码

```bash
cd /Users/rickxu/Finance/license-server
python3 gen_license.py --email customer@example.com --price 399
```

输出：

```
✓ 激活码已生成：WL-XK7Q-M2N4-P9H1
...
```

## 接口约定

### POST /admin/generate

生成新激活码。**仅你自己用。**

```
Headers: X-Admin-Key: <ADMIN_KEY>
Body:
{
  "customer_email": "user@example.com",
  "price": 399,
  "note": "optional"
}

Response:
{
  "license_code": "WL-XXXX-XXXX-XXXX",
  "customer_email": "...",
  "sold_at": "2026-06-22T..."
}
```

### POST /activate

客户端首次激活。

```
Body:
{
  "code": "WL-XXXX-XXXX-XXXX",
  "machine_id": "<sha256 hex>"
}

Response:
{
  "token": {
    "payload_b64": "...",
    "signature_b64": "..."
  },
  "license_status": "active"
}
```

token 解码后是 JSON：

```json
{
  "license_code": "WL-...",
  "machine_id": "...",
  "activated_at": "...",
  "type": "permanent",
  "issued_at": "..."
}
```

客户端用 `public_key.b64` 里的公钥验签。

### POST /update-data

年度数据包更新（税率 / 科目库等）。

```
Body:
{
  "code": "WL-XXXX-XXXX-XXXX",
  "machine_id": "...",
  "current_pack_version": "2026.01"
}

Response (已是最新):
{ "up_to_date": true, "latest_version": "2026.01" }

Response (有新版本):
{
  "up_to_date": false,
  "latest_version": "2026.06",
  "data": { ... }
}
```

## 测试

部署完用 curl 跑一遍：

```bash
# 1. 生成一个测试激活码
curl -X POST https://wiseledger-license.xxxx.workers.dev/admin/generate \
  -H "X-Admin-Key: <ADMIN_KEY>" \
  -H "Content-Type: application/json" \
  -d '{"customer_email":"test@test.com","price":1}'
# 输出：{"license_code":"WL-XXXX-XXXX-XXXX",...}

# 2. 模拟激活
curl -X POST https://wiseledger-license.xxxx.workers.dev/activate \
  -H "Content-Type: application/json" \
  -d '{"code":"WL-XXXX-XXXX-XXXX","machine_id":"fake_machine_001"}'
# 输出：{"token":{"payload_b64":"...","signature_b64":"..."},...}
```

## 安全提醒

- `.env` 永远不要提交到 git（已在 .gitignore）
- `ADMIN_KEY` 一旦泄露立刻 `wrangler secret put` 重新生成
- 私钥 `ED25519_PRIVATE_KEY_B64` 仅 Workers 用，本地仅备份
- 公钥 `public_key.b64` 嵌入客户端，泄露无所谓（只能验签不能签名）
