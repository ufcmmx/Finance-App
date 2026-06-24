/**
 * WiseLedger License Server — Cloudflare Workers
 *
 * 三个接口：
 *   POST /admin/generate   生成新激活码（需 X-Admin-Key header）
 *   POST /activate         客户端首次激活，返回签名 token
 *   POST /update-data      检查/下载年度数据包（税率/科目库）
 *
 * 存储：
 *   Cloudflare KV (binding: LICENSES)
 *   key: 激活码 (e.g., "WL-A3F2-B891-K7H3")
 *   val: JSON { customer, machine_id, sold_at, activated_at, status, ... }
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
// ────────────────────────────────────────────────────────────────
async function handleAdminGenerate(request, env) {
  // 鉴权
  const providedKey = request.headers.get('X-Admin-Key');
  if (providedKey !== env.ADMIN_KEY) {
    return errorResponse('Unauthorized', 401);
  }

  const body = await request.json().catch(() => ({}));
  const { customer_email = '', price = 0, note = '' } = body;

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
    sold_at: new Date().toISOString(),
    activated_at: null,
    machine_id: null,
    status: 'unused',          // unused | active | revoked
    unbind_count_this_year: 0,
    unbind_year: new Date().getUTCFullYear(),
  };

  await env.LICENSES.put(code, JSON.stringify(record));

  return jsonResponse({
    license_code: code,
    customer_email,
    sold_at: record.sold_at,
  });
}

// ────────────────────────────────────────────────────────────────
// 接口 2：POST /activate
// 客户端首次激活
// 请求：{ "code": "WL-...", "machine_id": "<sha256 hex>" }
// 返回：{ "token": { "payload_b64": "...", "signature_b64": "..." } }
// ────────────────────────────────────────────────────────────────
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

  // 已激活过的情况
  if (record.status === 'active') {
    if (record.machine_id === machine_id) {
      // 同一台机器重激活 → 重新签发 token
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
    }
  } else {
    // 首次激活
    record.status = 'active';
    record.activated_at = new Date().toISOString();
    record.machine_id = machine_id;
  }

  await env.LICENSES.put(code, JSON.stringify(record));

  // 生成签名 token
  const privKey = await importPrivateKey(env.ED25519_PRIVATE_KEY_B64);
  const payload = {
    license_code: code,
    machine_id,
    activated_at: record.activated_at,
    type: 'permanent',
    issued_at: new Date().toISOString(),
  };
  const token = await signToken(payload, privKey);

  return jsonResponse({ token, license_status: record.status });
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
          'Access-Control-Allow-Methods': 'POST, OPTIONS',
          'Access-Control-Allow-Headers': 'Content-Type, X-Admin-Key',
        },
      });
    }

    if (request.method !== 'POST') {
      return errorResponse('Method not allowed', 405);
    }

    try {
      if (url.pathname === '/admin/generate') {
        return await handleAdminGenerate(request, env);
      }
      if (url.pathname === '/activate') {
        return await handleActivate(request, env);
      }
      if (url.pathname === '/update-data') {
        return await handleUpdateData(request, env);
      }
      return errorResponse('Not found', 404);
    } catch (err) {
      console.error(err);
      return errorResponse(`Internal error: ${err.message}`, 500);
    }
  },
};
