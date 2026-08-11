// Cloudflare Pages Function：现场调查上传页的团队配置云端存取（存于 KV，跨电脑读取）
// 需要绑定 KV 命名空间，binding 名称：SAMPLING_RECORDS（key 前缀 team:，按登录账号区分）

function json(code, obj) {
  return new Response(JSON.stringify(obj), {
    status: code,
    headers: { "Content-Type": "application/json", "Cache-Control": "no-store" },
  });
}

const MAX_CONFIG = 64 * 1024; // 单个账号配置上限 64KB

function kv(context) {
  return context.env && context.env.SAMPLING_RECORDS;
}

function keyFor(account) {
  return "team:" + String(account || "").trim();
}

export async function onRequestGet(context) {
  const kvStore = kv(context);
  if (!kvStore) return json(500, { error: "未绑定 KV 命名空间 SAMPLING_RECORDS" });
  const account = new URL(context.request.url).searchParams.get("account") || "";
  if (!account) return json(400, { error: "缺少 account 参数" });
  try {
    const raw = await kvStore.get(keyFor(account));
    if (!raw) return json(200, {});
    const obj = JSON.parse(raw);
    return json(200, obj && typeof obj === "object" && !Array.isArray(obj) ? obj : {});
  } catch (e) {
    return json(500, { error: String((e && e.message) || e) });
  }
}

export async function onRequestPut(context) {
  const kvStore = kv(context);
  if (!kvStore) return json(500, { error: "未绑定 KV 命名空间 SAMPLING_RECORDS" });
  try {
    const body = await context.request.json();
    const account = String(body.account || "").trim();
    const config = body.config;
    if (!account) return json(400, { error: "缺少 account 参数" });
    if (!config || typeof config !== "object" || Array.isArray(config)) {
      return json(400, { error: "config 格式错误" });
    }
    const text = JSON.stringify(config);
    if (text.length > MAX_CONFIG) return json(413, { error: "配置过大" });
    await kvStore.put(keyFor(account), text);
    return json(200, { ok: true });
  } catch (e) {
    return json(500, { error: String((e && e.message) || e) });
  }
}

export async function onRequestPost(context) {
  return onRequestPut(context);
}
