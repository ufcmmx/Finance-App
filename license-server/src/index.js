/**
 * WiseLedger License Server — Cloudflare Workers
 *
 * 接口：
 *   POST /admin/generate   生成新激活码（需 X-Admin-Key header）
 *   POST /activate         客户端首次激活，返回签名 token
 *   POST /trial            申请 7 天试用
 *   POST /update-data      检查/下载年度数据包（税率/科目库）
 *   POST /download-log     网站下载按钮点击埋点（KV 计数 + D1 日志）
 *   GET  /stats            聚合统计（需 X-Admin-Key header）
 *
 * 存储：
 *   Cloudflare KV (binding: LICENSES)
 *     key: 激活码 (e.g., "WL-A3F2-B891-K7H3")
 *     val: JSON { customer, machine_id, sold_at, activated_at, status, ... }
 *   Cloudflare KV (同一 binding)
 *     key: "download:total"       val: "123"                累计总数
 *     key: "download:YYYY-MM-DD"  val: "17"                 每日计数
 *   Cloudflare D1 (binding: STATS)
 *     table: downloads(id, created_at)  每次点击一行
 *
 * 密钥：
 *   ADMIN_KEY                环境变量 (Workers Secret)
 *   ED25519_PRIVATE_KEY_B64  环境变量 (Workers Secret)
 */

// 32 字母数字字符集（去掉容易混淆的 I, O, 1, 0）
const ALPHABET = 'ABCDEFGHJKLMNPQRSTUVWXYZ23456789';

// ────────────────────────────────────────────────────────────────
// 激活码生成 + 校验位
// ────────────────────────────────────────────────────────────────
function generateLicenseCode() {
  // 11 位随机 + 1 位校验，格式 WL-XXXX-XXXX-XXXX
  const random = new Uint8Array(11);
  crypto.getRandomValues(random);
  let body = '';
  for (let i = 0; i < 11; i++) {
    body += ALPHABET[random[i] % 32];
  }
  body += checksum(body);
  return `WL-${body.slice(0, 4)}-${body.slice(4, 8)}-${body.slice(8, 12)}`;
}

function checksum(body11) {
  let sum = 0;
  for (const ch of body11) sum += ALPHABET.indexOf(ch);
  return ALPHABET[sum % 32];
}

function validateLicenseFormat(code) {
  // WL-XXXX-XXXX-XXXX
  if (!/^WL-[A-Z2-9]{4}-[A-Z2-9]{4}-[A-Z2-9]{4}$/.test(code)) return false;
  const body = code.slice(3).replace(/-/g, ''); // 12 chars
  return checksum(body.slice(0, 11)) === body[11];
}

// ────────────────────────────────────────────────────────────────
// Ed25519 签名
// ────────────────────────────────────────────────────────────────
function b64decode(b64) {
  return Uint8Array.from(atob(b64), c => c.charCodeAt(0));
}
function b64encode(bytes) {
  return btoa(String.fromCharCode(...bytes));
}

async function importPrivateKey(privB64) {
  // Workers SubtleCrypto 不接受 raw 格式的 Ed25519 私钥（raw 被当公钥）
  // 必须用 PKCS8。手动给 32 字节裸私钥包 PKCS8 ASN.1 头。
  const raw = b64decode(privB64);
  if (raw.length !== 32) {
    throw new Error(`Expected 32-byte Ed25519 private key, got ${raw.length}`);
  }
  // PKCS8 prefix for Ed25519: SEQUENCE + version + AlgID(Ed25519 OID) + OCTET STRING
  const PKCS8_PREFIX = new Uint8Array([
    0x30, 0x2e,             // SEQUENCE, length 46
    0x02, 0x01, 0x00,       // INTEGER version 0
    0x30, 0x05,             // SEQUENCE, length 5
    0x06, 0x03, 0x2b, 0x65, 0x70, // OID 1.3.101.112 (Ed25519)
    0x04, 0x22,             // OCTET STRING, length 34
    0x04, 0x20,             // inner OCTET STRING, length 32
  ]);
  const pkcs8 = new Uint8Array(PKCS8_PREFIX.length + 32);
  pkcs8.set(PKCS8_PREFIX);
  pkcs8.set(raw, PKCS8_PREFIX.length);
  return await crypto.subtle.importKey(
    'pkcs8',
    pkcs8,
    { name: 'Ed25519' },
    false,
    ['sign']
  );
}

async function signToken(payload, privKey) {
  const payloadJson = JSON.stringify(payload);
  const payloadBytes = new TextEncoder().encode(payloadJson);
  const sigBytes = new Uint8Array(
    await crypto.subtle.sign('Ed25519', privKey, payloadBytes)
  );
  return {
    payload_b64: b64encode(payloadBytes),
    signature_b64: b64encode(sigBytes),
  };
}

// ────────────────────────────────────────────────────────────────
// 工具函数
// ────────────────────────────────────────────────────────────────
function jsonResponse(obj, status = 200) {
  return new Response(JSON.stringify(obj), {
    status,
    headers: { 'Content-Type': 'application/json; charset=utf-8' },
  });
}

function errorResponse(message, status = 400) {
  return jsonResponse({ error: message }, status);
}

// ────────────────────────────────────────────────────────────────
// 接口 1：POST /admin/generate
// 生成新激活码，由你卖货时手动调用
// 支持类型：annual (¥99/年 订阅版) | permanent (¥399 永久版)
// ────────────────────────────────────────────────────────────────
const VALID_TYPES = ['annual', 'permanent'];
const ANNUAL_DAYS = 365;
const GRACE_DAYS = 7;
const TRIAL_DAYS = 7;

async function handleAdminGenerate(request, env) {
  // 鉴权
  const providedKey = request.headers.get('X-Admin-Key');
  if (providedKey !== env.ADMIN_KEY) {
    return errorResponse('Unauthorized', 401);
  }

  const body = await request.json().catch(() => ({}));
  const {
    customer_email = '',
    price = 0,
    note = '',
    type = 'permanent',      // 默认永久版
  } = body;

  if (!VALID_TYPES.includes(type)) {
    return errorResponse(`Invalid type: ${type}. Must be one of: ${VALID_TYPES.join(', ')}`);
  }

  // 生成激活码（重试直到唯一）
  let code;
  for (let i = 0; i < 10; i++) {
    code = generateLicenseCode();
    const existing = await env.LICENSES.get(code);
    if (!existing) break;
  }

  const record = {
    customer_email,
    price,
    note,
    type,                       // annual | permanent
    sold_at: new Date().toISOString(),
    activated_at: null,
    machine_id: null,
    expires_at: null,           // 激活时才计算
    status: 'unused',           // unused | active | revoked
    unbind_count_this_year: 0,
    unbind_year: new Date().getUTCFullYear(),
  };

  await env.LICENSES.put(code, JSON.stringify(record));

  return jsonResponse({
    license_code: code,
    customer_email,
    type,
    sold_at: record.sold_at,
  });
}

// ────────────────────────────────────────────────────────────────
// 接口 2：POST /activate
// 客户端首次激活
// 请求：{ "code": "WL-...", "machine_id": "<sha256 hex>" }
// 返回：{ "token": { "payload_b64": "...", "signature_b64": "..." } }
// ────────────────────────────────────────────────────────────────
function addDays(iso, days) {
  const d = new Date(iso);
  d.setUTCDate(d.getUTCDate() + days);
  return d.toISOString();
}

async function handleActivate(request, env) {
  const body = await request.json().catch(() => ({}));
  const { code, machine_id } = body;

  if (!code || !machine_id) {
    return errorResponse('Missing code or machine_id');
  }
  if (!validateLicenseFormat(code)) {
    return errorResponse('Invalid license code format');
  }

  const raw = await env.LICENSES.get(code);
  if (!raw) {
    return errorResponse('License code not found', 404);
  }
  const record = JSON.parse(raw);

  if (record.status === 'revoked') {
    return errorResponse('License has been revoked', 403);
  }

  // 兼容旧记录（没有 type 字段）
  const licType = record.type || 'permanent';

  // 已激活过的情况
  if (record.status === 'active') {
    if (record.machine_id === machine_id) {
      // 同一台机器重激活 → 重新签发 token（不重置 expires_at 和 activated_at）
    } else {
      // 不同机器 → 检查解绑配额
      const currentYear = new Date().getUTCFullYear();
      if (record.unbind_year !== currentYear) {
        record.unbind_year = currentYear;
        record.unbind_count_this_year = 0;
      }
      if (record.unbind_count_this_year >= 2) {
        return errorResponse(
          'Annual unbind limit reached (2/year)',
          403
        );
      }
      record.unbind_count_this_year += 1;
      record.machine_id = machine_id;
      // 注意：换机不重置 activated_at 和 expires_at
    }
  } else {
    // 首次激活
    record.status = 'active';
    record.activated_at = new Date().toISOString();
    record.machine_id = machine_id;
    // 订阅版计算到期时间
    if (licType === 'annual') {
      record.expires_at = addDays(record.activated_at, ANNUAL_DAYS);
    } else {
      record.expires_at = null;   // 永久版
    }
  }

  await env.LICENSES.put(code, JSON.stringify(record));

  // 生成签名 token
  const privKey = await importPrivateKey(env.ED25519_PRIVATE_KEY_B64);
  const payload = {
    license_code: code,
    machine_id,
    type: licType,
    activated_at: record.activated_at,
    expires_at: record.expires_at,        // null 表示永久
    grace_days: licType === 'annual' ? GRACE_DAYS : 0,
    issued_at: new Date().toISOString(),
  };
  const token = await signToken(payload, privKey);

  return jsonResponse({ token, license_status: record.status, type: licType });
}

// ────────────────────────────────────────────────────────────────
// 接口 4：POST /trial
// 客户端申请免费试用 7 天
// 请求：{ "machine_id": "<sha256 hex>" }
// 返回：签名 token（type: trial, expires_at: now + 7 天）
// 每台机器仅可试用一次（用 KV key "trial:<machine_id>" 记录）
// ────────────────────────────────────────────────────────────────
async function handleTrial(request, env) {
  const body = await request.json().catch(() => ({}));
  const { machine_id } = body;

  if (!machine_id) {
    return errorResponse('Missing machine_id');
  }

  const trialKey = `trial:${machine_id}`;
  const existing = await env.LICENSES.get(trialKey);
  if (existing) {
    const prev = JSON.parse(existing);
    return errorResponse(
      `Trial already used (started at ${prev.started_at})`,
      403
    );
  }

  const startedAt = new Date().toISOString();
  const expiresAt = addDays(startedAt, TRIAL_DAYS);

  // 生成一个仅用于展示的伪激活码（不写入 LICENSES 主表）
  const displayCode = `WL-TRIAL-${machine_id.slice(0, 4).toUpperCase()}-${machine_id.slice(4, 8).toUpperCase()}`;

  const trialRecord = {
    machine_id,
    started_at: startedAt,
    expires_at: expiresAt,
    display_code: displayCode,
  };
  await env.LICENSES.put(trialKey, JSON.stringify(trialRecord));

  // 生成签名 token
  const privKey = await importPrivateKey(env.ED25519_PRIVATE_KEY_B64);
  const payload = {
    license_code: displayCode,
    machine_id,
    type: 'trial',
    activated_at: startedAt,
    expires_at: expiresAt,
    grace_days: 0,               // 试用无宽限期
    issued_at: startedAt,
  };
  const token = await signToken(payload, privKey);

  return jsonResponse({
    token,
    license_status: 'trial',
    type: 'trial',
    expires_at: expiresAt,
  });
}

// ────────────────────────────────────────────────────────────────
// 接口 5：POST /download-log
// 网站下载按钮点击埋点。只记聚合数，不记 IP/UA。
// 请求：{} （空 body 即可，也可选传 { source: "hero" } 标记来源）
// 返回：{ ok: true }
// ────────────────────────────────────────────────────────────────
async function handleDownloadLog(request, env) {
  const body = await request.json().catch(() => ({}));
  const source = typeof body.source === 'string' ? body.source.slice(0, 32) : '';

  const now = new Date();
  const dayKey = `download:${now.toISOString().slice(0, 10)}`;   // YYYY-MM-DD
  const totalKey = 'download:total';

  // KV 计数（原子性差，但埋点场景够用；并发极高时可切 D1 SUM）
  const [dayVal, totalVal] = await Promise.all([
    env.LICENSES.get(dayKey),
    env.LICENSES.get(totalKey),
  ]);
  const nextDay = (parseInt(dayVal || '0', 10) + 1).toString();
  const nextTotal = (parseInt(totalVal || '0', 10) + 1).toString();

  const kvWrites = [
    env.LICENSES.put(dayKey, nextDay),
    env.LICENSES.put(totalKey, nextTotal),
  ];

  // D1 详细日志（如果绑定了）
  if (env.STATS) {
    kvWrites.push(
      env.STATS.prepare('INSERT INTO downloads (created_at, source) VALUES (?, ?)')
        .bind(now.toISOString(), source)
        .run()
        .catch(err => console.error('D1 insert failed:', err))
    );
  }

  await Promise.all(kvWrites);

  return jsonResponse({ ok: true });
}

// ────────────────────────────────────────────────────────────────
// 接口 6：GET /stats
// 聚合统计（下载数 + 试用数 + 激活数），需 X-Admin-Key 鉴权
// 返回 JSON，浏览器直接打开即可看
// ────────────────────────────────────────────────────────────────
async function handleStats(request, env) {
  const providedKey = request.headers.get('X-Admin-Key') || new URL(request.url).searchParams.get('key');
  if (providedKey !== env.ADMIN_KEY) {
    return errorResponse('Unauthorized', 401);
  }

  const today = new Date();
  const todayStr = today.toISOString().slice(0, 10);

  // 累计总数（KV）
  const total = parseInt((await env.LICENSES.get('download:total')) || '0', 10);
  const todayCount = parseInt((await env.LICENSES.get(`download:${todayStr}`)) || '0', 10);

  // 最近 7 天 / 30 天
  const dailyKeys = [];
  for (let i = 0; i < 30; i++) {
    const d = new Date(today);
    d.setUTCDate(d.getUTCDate() - i);
    dailyKeys.push(`download:${d.toISOString().slice(0, 10)}`);
  }
  const dailyVals = await Promise.all(dailyKeys.map(k => env.LICENSES.get(k)));
  const daily = dailyKeys.map((k, i) => ({
    date: k.slice('download:'.length),
    count: parseInt(dailyVals[i] || '0', 10),
  }));
  const last7 = daily.slice(0, 7).reduce((s, d) => s + d.count, 0);
  const last30 = daily.reduce((s, d) => s + d.count, 0);

  // D1 详细（如果有）
  let d1_total = null;
  if (env.STATS) {
    try {
      const r = await env.STATS.prepare('SELECT COUNT(*) AS n FROM downloads').first();
      d1_total = r?.n ?? null;
    } catch (err) {
      d1_total = `error: ${err.message}`;
    }
  }

  return jsonResponse({
    generated_at: new Date().toISOString(),
    downloads: {
      total,           // KV 累计
      today: todayCount,
      last_7_days: last7,
      last_30_days: last30,
      d1_total,        // D1 累计（可与 KV 对账）
      daily,           // 最近 30 天每日数据
    },
  });
}

// ────────────────────────────────────────────────────────────────
// 接口 3：POST /update-data
// 年度数据包（税率/科目库）更新
// 暂时只返回当前可用版本，真正的数据可以放到 R2 或硬编码
// ────────────────────────────────────────────────────────────────
async function handleUpdateData(request, env) {
  const body = await request.json().catch(() => ({}));
  const { code, machine_id, current_pack_version } = body;

  if (!code) return errorResponse('Missing code');

  const raw = await env.LICENSES.get(code);
  if (!raw) return errorResponse('License code not found', 404);
  const record = JSON.parse(raw);

  if (record.status !== 'active') {
    return errorResponse('License not active', 403);
  }
  if (record.machine_id !== machine_id) {
    return errorResponse('Machine ID mismatch', 403);
  }

  // 当前数据包版本（你将来更新时改这个）
  const latest = '2026.01';

  if (current_pack_version === latest) {
    return jsonResponse({
      up_to_date: true,
      latest_version: latest,
    });
  }

  // 返回新数据包内容（这里先返回占位，实际可以放税率表等）
  return jsonResponse({
    up_to_date: false,
    latest_version: latest,
    data: {
      vat_rates: {
        general: 0.13,
        small_scale: 0.01,
      },
      // 之后扩展：科目模板、银行清算号等
    },
  });
}

// ────────────────────────────────────────────────────────────────
// 路由
// ────────────────────────────────────────────────────────────────
export default {
  async fetch(request, env, ctx) {
    const url = new URL(request.url);

    // CORS 简单处理（开发期间方便测试）
    if (request.method === 'OPTIONS') {
      return new Response(null, {
        headers: {
          'Access-Control-Allow-Origin': '*',
          'Access-Control-Allow-Methods': 'GET, POST, OPTIONS',
          'Access-Control-Allow-Headers': 'Content-Type, X-Admin-Key',
        },
      });
    }

    try {
      // GET 路由（目前只有 /stats）
      if (request.method === 'GET') {
        if (url.pathname === '/stats') {
          const resp = await handleStats(request, env);
          resp.headers.set('Access-Control-Allow-Origin', '*');
          return resp;
        }
        return errorResponse('Not found', 404);
      }

      if (request.method !== 'POST') {
        return errorResponse('Method not allowed', 405);
      }

      if (url.pathname === '/admin/generate') {
        return await handleAdminGenerate(request, env);
      }
      if (url.pathname === '/activate') {
        return await handleActivate(request, env);
      }
      if (url.pathname === '/trial') {
        return await handleTrial(request, env);
      }
      if (url.pathname === '/update-data') {
        return await handleUpdateData(request, env);
      }
      if (url.pathname === '/download-log') {
        const resp = await handleDownloadLog(request, env);
        resp.headers.set('Access-Control-Allow-Origin', '*');
        return resp;
      }
      return errorResponse('Not found', 404);
    } catch (err) {
      console.error(err);
      return errorResponse(`Internal error: ${err.message}`, 500);
    }
  },
};
