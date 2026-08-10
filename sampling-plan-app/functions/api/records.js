// Cloudflare Pages Function：数据记录读写（存于 Cloudflare KV）
// 需在 Pages 项目中绑定 KV 命名空间，binding 名称：SAMPLING_RECORDS

const KEY = "records";

function json(code, obj) {
  return new Response(JSON.stringify(obj), {
    status: code,
    headers: { "Content-Type": "application/json", "Cache-Control": "no-store" },
  });
}

export async function onRequestGet(context) {
  const kv = context.env && context.env.SAMPLING_RECORDS;
  if (!kv) return json(500, { error: "未绑定 KV 命名空间 SAMPLING_RECORDS" });
  try {
    const raw = await kv.get(KEY);
    const list = raw ? JSON.parse(raw) : [];
    return json(200, Array.isArray(list) ? list : []);
  } catch (e) {
    return json(500, { error: String((e && e.message) || e) });
  }
}

export async function onRequestPut(context) {
  const kv = context.env && context.env.SAMPLING_RECORDS;
  if (!kv) return json(500, { error: "未绑定 KV 命名空间 SAMPLING_RECORDS" });
  try {
    const body = await context.request.json();
    if (!Array.isArray(body)) return json(400, { error: "数据格式错误" });
    await kv.put(KEY, JSON.stringify(body));
    return json(200, { ok: true, count: body.length });
  } catch (e) {
    return json(500, { error: String((e && e.message) || e) });
  }
}

export async function onRequestPost(context) {
  return onRequestPut(context);
}
