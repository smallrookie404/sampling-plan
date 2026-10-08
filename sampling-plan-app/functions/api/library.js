// Cloudflare Pages Function：参考库（危害因素库 + 检测项目参考）读写（存于 Cloudflare KV）
// 需要绑定 KV 命名空间，binding 名称：SAMPLING_RECORDS

const KEY = "library";

// 读取响应头：max-age=0 + SWR——浏览器立即用缓存渲染并后台再验证；
// 应用端按 updatedAt 比较同步，缓存旧值不会造成数据回退。写入仍 no-store。
const READ_CACHE = "public, max-age=0, stale-while-revalidate=30";

function jsonCache(code, obj) {
  return new Response(JSON.stringify(obj), {
    status: code,
    headers: { "Content-Type": "application/json", "Cache-Control": READ_CACHE },
  });
}

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
    if (!raw) return jsonCache(200, {});
    const obj = JSON.parse(raw);
    return jsonCache(200, obj && typeof obj === "object" && !Array.isArray(obj) ? obj : {});
  } catch (e) {
    return json(500, { error: String((e && e.message) || e) });
  }
}

export async function onRequestPut(context) {
  const kv = context.env && context.env.SAMPLING_RECORDS;
  if (!kv) return json(500, { error: "未绑定 KV 命名空间 SAMPLING_RECORDS" });
  try {
    const body = await context.request.json();
    if (!body || typeof body !== "object" || Array.isArray(body)) {
      return json(400, { error: "数据格式错误" });
    }
    await kv.put(KEY, JSON.stringify(body));
    return json(200, { ok: true });
  } catch (e) {
    return json(500, { error: String((e && e.message) || e) });
  }
}

export async function onRequestPost(context) {
  return onRequestPut(context);
}
