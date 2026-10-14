// Cloudflare Pages Function：账号桥（云端版）
//
// 目的：让**部署到 Cloudflare 的** App 也能零人工执行「数据报告回退」——
// 回退第 1 步（jcbg/检测结果报告单）要以 btc-gxz 身份、第 2 步（jcdg/检测底稿）以 btc-wcy 身份。
// 浏览器不持有这些账号的令牌，本地版由 Python 账号桥提供；云端没有本机，故把
// 「账号池 + 验证码 OCR + RSA 加密 + 登录」整体移植到本 Function（见 js/captcha-core.mjs）。
//
// 路由（与本地 server.mjs 的同源路由保持一致，网页端无需区分环境）：
//   GET /api/bridge/ping                 → { ok, autoStart, up, cloud }
//   GET /api/bridge/token/<账号>          → { token, orgId, ... }
//
// 安全：
//   · 调用方必须带自己的平台会话令牌（Authorization 头），本接口会用它调 /auth/info 校验；
//     校验不通过一律 401 —— 避免这个公开接口被任意人拿去换 btc-gxz/btc-wcy 的令牌。
//   · 账号口令只来自 Pages 加密变量 BRIDGE_ACCOUNTS 或 KV bridge:accounts，**不进前端代码**。
//   · 令牌与验证码用量存 KV（binding: SAMPLING_RECORDS）。
//
// 平台访问：复用现成的同源代理 /api/platform（cloudflare:sockets 原始 TCP），
// 所以这里的 apiBase 就是 /api/platform。
import { loginAccount } from "./_captcha-core.mjs";
import { CAPTCHA_TEMPLATES } from "./_captcha-templates.mjs";

const TOKEN_PREFIX = "bridge:token:";
const BUDGET_PREFIX = "bridge:captcha:";
const KV_ACCOUNTS_KEY = "bridge:accounts";
const CONF_THRESHOLD = 0.25;
const MAX_TRIES = 4;

const CORS = {
  "Access-Control-Allow-Origin": "*",
  "Access-Control-Allow-Methods": "GET, OPTIONS",
  "Access-Control-Allow-Headers": "Authorization, organizationId, Content-Type",
  "Access-Control-Max-Age": "86400",
};

function json(obj, status = 200) {
  return new Response(JSON.stringify(obj), {
    status,
    headers: { "Content-Type": "application/json; charset=utf-8", "Cache-Control": "no-store", ...CORS },
  });
}

function today() {
  return new Date().toISOString().slice(0, 10);
}

async function readAccounts(env) {
  // 1) Pages 加密变量 BRIDGE_ACCOUNTS（推荐）：JSON 字符串 [{username,password,orgId,label}]
  if (env && env.BRIDGE_ACCOUNTS) {
    try {
      const a = JSON.parse(env.BRIDGE_ACCOUNTS);
      if (Array.isArray(a)) return a;
    } catch (e) { /* 落到 KV */ }
  }
  // 2) KV bridge:accounts
  if (env && env.SAMPLING_RECORDS) {
    try {
      const raw = await env.SAMPLING_RECORDS.get(KV_ACCOUNTS_KEY);
      if (raw) {
        const a = JSON.parse(raw);
        if (Array.isArray(a)) return a;
      }
    } catch (e) { /* ignore */ }
  }
  return [];
}

function pickAccount(list, username) {
  const u = String(username || "").toLowerCase();
  return list.find((a) => String(a && a.username || "").toLowerCase() === u) || null;
}

/** 用给定令牌探测平台会话是否有效 */
async function sessionValid(base, token, orgId) {
  if (!token) return false;
  try {
    const r = await fetch(base + "/auth/info", {
      headers: { Authorization: token, organizationId: String(orgId || 1) },
    });
    return r.status === 200;
  } catch (e) {
    return false;
  }
}

export async function onRequest(context) {
  const { request, env, params } = context;
  if (request.method === "OPTIONS") return new Response(null, { status: 204, headers: CORS });

  const segs = Array.isArray(params && params.path) ? params.path : [];
  const head = segs[0] || "";
  const base = new URL("/api/platform", request.url).href;

  if (head === "ping") {
    return json({ ok: true, autoStart: false, up: true, cloud: true });
  }

  if (head !== "token") {
    return json({ ok: false, message: "未知路径 /api/bridge/" + segs.join("/") }, 404);
  }

  const username = segs[1];
  if (!username) return json({ ok: false, message: "缺少账号名" }, 400);

  // —— 门禁：调用方必须带一个**有效的平台会话令牌** ——
  const callerToken = request.headers.get("authorization") || request.headers.get("Authorization") || "";
  if (!callerToken) {
    return json({ ok: false, message: "缺少 Authorization（本接口仅对已登录平台的用户开放）" }, 401);
  }
  if (!(await sessionValid(base, callerToken, request.headers.get("organizationid") || 1))) {
    return json({ ok: false, message: "调用方平台会话无效或已过期，请重新登录后再试" }, 401);
  }

  const kv = env && env.SAMPLING_RECORDS;
  const cacheKey = TOKEN_PREFIX + username;

  // 1) 缓存命中且仍有效 → 直接用（不消耗验证码）
  if (kv) {
    try {
      const raw = await kv.get(cacheKey);
      if (raw) {
        const t = JSON.parse(raw);
        if (t && t.token && await sessionValid(base, t.token, t.orgId)) {
          return json({ ok: true, token: t.token, orgId: t.orgId, user: t.user || null, via: t.via || null, cached: true });
        }
      }
    } catch (e) { /* 忽略坏缓存，走重登 */ }
  }

  // 2) 惰性登录
  const accounts = await readAccounts(env);
  if (!accounts.length) {
    return json({
      ok: false,
      message: "云端未配置账号：请在 Pages 设置里加加密变量 BRIDGE_ACCOUNTS（JSON: "
        + '[{"username":"btc-gxz","password":"…","orgId":1},{"username":"btc-wcy","password":"…","orgId":1}]）',
    }, 503);
  }
  const acc = pickAccount(accounts, username);
  if (!acc) return json({ ok: false, message: "账号不在云端配置中：" + username }, 404);

  const logs = [];
  const got = await loginAccount({
    apiBase: base,
    username: acc.username,
    password: acc.password,
    templates: CAPTCHA_TEMPLATES,
    confThreshold: CONF_THRESHOLD,
    maxTries: MAX_TRIES,
    onLog: (m) => logs.push(m),
    onCaptcha: async () => {
      if (!kv) return;
      try {
        const k = BUDGET_PREFIX + today();
        const n = Number(await kv.get(k) || 0) + 1;
        await kv.put(k, String(n));
      } catch (e) { /* 用量统计失败不影响登录 */ }
    },
  });

  if (!got) {
    return json({ ok: false, message: "登录失败（已尝试 " + MAX_TRIES + " 张验证码）", logs }, 502);
  }

  const rec = { token: got.token, orgId: acc.orgId || 1, user: got.user, via: got.via, conf: got.conf, at: new Date().toISOString() };
  if (kv) {
    try { await kv.put(cacheKey, JSON.stringify(rec)); } catch (e) { /* 缓存失败仍可用 */ }
  }
  return json({ ok: true, token: rec.token, orgId: rec.orgId, user: rec.user, via: rec.via, conf: rec.conf, cached: false, logs });
}
