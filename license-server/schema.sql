-- WiseLedger 下载埋点 D1 表结构
-- 部署命令：
--   wrangler d1 execute wiseledger-stats --remote --file=./schema.sql

CREATE TABLE IF NOT EXISTS downloads (
  id         INTEGER PRIMARY KEY AUTOINCREMENT,
  created_at TEXT    NOT NULL,           -- ISO 8601 UTC
  source     TEXT    DEFAULT ''          -- 来源标记（hero/nav/footer 等，可选）
);

CREATE INDEX IF NOT EXISTS idx_downloads_created_at ON downloads(created_at);
