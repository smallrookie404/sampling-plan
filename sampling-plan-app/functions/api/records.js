// Cloudflare Pages Function：数据记录读写（存于 Cloudflare KV，按记录分 key）
// 需在 Pages 项目中绑定 KV 命名空间，binding 名称：SAMPLING_RECORDS
//
// 结构（与本地 server.mjs 保持一致）：
//   records-index        → 索引 [{id,name,createdAt,updatedAt}, ...]
//   record:<id>          → 单条完整记录（含 rows）
// 接口：
//   GET    /api/records          → 索引
//   GET    /api/records?id=x     → 单条完整记录
//   PUT    /api/records?id=x     → 保存单条（body 为该记录对象），并更新索引
//   DELETE /api/records?id=x     → 删除单条并更新索引

const INDEX_KEY = "records-index";
const recKey = (id) => "record:" + id;

function json(code, obj) {
  return new Response(JSON.stringify(obj), {
    status: code,
    headers: { "Content-Type": "application/json", "Cache-Control": "no-store" },
  });
}

async function readIndex(kv) {
  const raw = await kv.get(INDEX_KEY);
  const list = raw ? JSON.parse(raw) : [];
  return Array.isArray(list) ? list : [];
}

async function writeIndex(kv, index) {
  await kv.put(INDEX_KEY, JSON.stringify(index));
}

// 旧格式兼容：KV 里可能遗留整列表的 "records" key（元素含 rows）→ 拆分为分 key 结构
async function migrateLegacy(kv) {
  try {
    const raw = await kv.get("records");
    if (!raw) return;
    const list = JSON.parse(raw);
    if (!Array.isArray(list) || !list.some((r) => r && Array.isArray(r.rows))) return;
    const index = [];
    for (const rec of list) {
      if (!rec || !rec.id) continue;
      await kv.put(recKey(rec.id), JSON.stringify(rec));
      index.push({ id: rec.id, name: rec.name || "", createdAt: rec.createdAt || "", updatedAt: rec.updatedAt || "" });
    }
    await writeIndex(kv, index);
    await kv.delete("records");
  } catch {}
}

export async function onRequestGet(context) {
  const kv = context.env && context.env.SAMPLING_RECORDS;
  if (!kv) return json(500, { error: "未绑定 KV 命名空间 SAMPLING_RECORDS" });
  try {
    await migrateLegacy(kv);
    const id = new URL(context.request.url).searchParams.get("id");
    if (id) {
      const raw = await kv.get(recKey(id));
      if (!raw) return json(404, { error: "记录不存在" });
      return json(200, JSON.parse(raw));
    }
    return json(200, await readIndex(kv));
  } catch (e) {
    return json(500, { error: String((e && e.message) || e) });
  }
}

async function saveRecord(context, kv) {
  const id = new URL(context.request.url).searchParams.get("id");
  if (!id) return json(400, { error: "缺少记录 id 参数" });
  const body = await context.request.json();
  if (!body || typeof body !== "object" || body.id !== id) {
    return json(400, { error: "数据格式错误（需含与 id 参数一致的 id 字段）" });
  }
  await kv.put(recKey(id), JSON.stringify(body));
  const index = await readIndex(kv);
  const meta = { id: body.id, name: body.name || "", createdAt: body.createdAt || "", updatedAt: body.updatedAt || "" };
  const pos = index.findIndex((r) => r.id === id);
  if (pos >= 0) index[pos] = meta;
  else index.push(meta);
  await writeIndex(kv, index);
  return json(200, { ok: true, id });
}

export async function onRequestPut(context) {
  const kv = context.env && context.env.SAMPLING_RECORDS;
  if (!kv) return json(500, { error: "未绑定 KV 命名空间 SAMPLING_RECORDS" });
  try {
    await migrateLegacy(kv);
    return await saveRecord(context, kv);
  } catch (e) {
    return json(500, { error: String((e && e.message) || e) });
  }
}

export async function onRequestPost(context) {
  return onRequestPut(context);
}

export async function onRequestDelete(context) {
  const kv = context.env && context.env.SAMPLING_RECORDS;
  if (!kv) return json(500, { error: "未绑定 KV 命名空间 SAMPLING_RECORDS" });
  try {
    const id = new URL(context.request.url).searchParams.get("id");
    if (!id) return json(400, { error: "缺少记录 id 参数" });
    await kv.delete(recKey(id));
    await writeIndex(kv, (await readIndex(kv)).filter((r) => r.id !== id));
    return json(200, { ok: true });
  } catch (e) {
    return json(500, { error: String((e && e.message) || e) });
  }
}
