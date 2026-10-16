// Cloudflare Pages Functions 逻辑验证（本地用模拟 KV 测试，无需真实账号）
//
// 说明：records.js 早已从「整表数组存一个 KV key」改为
//   records-index → 索引（不含 rows）
//   record:<id>  → 单条完整记录
// 本测试按**当前**接口契约断言（旧版断言「GET 返回数组」自 6662b3e 起已失效）。
// GET/PUT/DELETE 均需带 `?id=` query，故这里用 URL 构造 request。
import fs from "node:fs";

// 以 data URL 方式把函数源码作为 ESM 加载，避免影响仓库内其他 CommonJS 文件
async function loadFunction(rel) {
  const code = fs.readFileSync(new URL(rel, import.meta.url), "utf8");
  return import("data:text/javascript;base64," + Buffer.from(code).toString("base64"));
}

const healthMod = await loadFunction("../functions/api/health.js");
const recordsMod = await loadFunction("../functions/api/records.js");
const libraryMod = await loadFunction("../functions/api/library.js");
const { onRequestGet: healthGet } = healthMod;
const { onRequestGet, onRequestPut, onRequestDelete } = recordsMod;
const {
  onRequestGet: libraryGet,
  onRequestPut: libraryPut,
} = libraryMod;

// 模拟 KV：Map 存储，支持 get / put / delete
function fakeKV() {
  const store = new Map();
  return {
    async get(k) { return store.has(k) ? store.get(k) : null; },
    async put(k, v) { store.set(k, v); },
    async delete(k) { store.delete(k); },
  };
}

// 构造带 ?id= 的最小 request 桩
function reqWithId(id, body) {
  const r = { url: "https://x/api/records" + (id ? "?id=" + encodeURIComponent(id) : "") };
  if (body !== undefined) r.json = async () => body;
  return r;
}

const kv = fakeKV();
const env = { SAMPLING_RECORDS: kv };

// 健康检查
const health = await healthGet();
if (health.status !== 200) throw new Error("health 失败: " + health.status);

// 初始为空：GET 无 id → 返回空索引数组
let res = await onRequestGet({ env, request: reqWithId() });
let list = await res.json();
if (res.status !== 200 || !Array.isArray(list) || list.length !== 0) throw new Error("初始 GET 索引不符");

// 写入单条（body 需含与 id 参数一致的 id）
res = await onRequestPut({ env, request: reqWithId("1", { id: "1", name: "云测试记录", rows: [] }) });
if (res.status !== 200) throw new Error("PUT 失败: " + res.status);

// 读索引：应含 1 条元信息（且不含 rows）
res = await onRequestGet({ env, request: reqWithId() });
list = await res.json();
if (list.length !== 1 || list[0].name !== "云测试记录" || "rows" in list[0]) {
  throw new Error("索引往返失败");
}

// 读单条完整记录
res = await onRequestGet({ env, request: reqWithId("1") });
const rec = await res.json();
if (res.status !== 200 || rec.name !== "云测试记录" || !Array.isArray(rec.rows)) {
  throw new Error("单条读取失败");
}

// 删除单条 → 索引应为空
res = await onRequestDelete({ env, request: reqWithId("1") });
if (res.status !== 200) throw new Error("DELETE 失败: " + res.status);
res = await onRequestGet({ env, request: reqWithId() });
list = await res.json();
if (list.length !== 0) throw new Error("删除后索引未清空");

// PUT 缺 id → 400
res = await onRequestPut({ env, request: reqWithId("", { name: "x" }) });
if (res.status !== 400) throw new Error("缺少 id 应返回 400");

// PUT 的 body.id 与 query id 不一致 → 400
res = await onRequestPut({ env, request: reqWithId("2", { id: "3", name: "x" }) });
if (res.status !== 400) throw new Error("id 不一致应返回 400");

// 未绑定 KV → 500，且提示清晰
res = await onRequestGet({ env: {}, request: reqWithId() });
if (res.status !== 500) throw new Error("未绑定 KV 应返回 500");
res = await onRequestPut({ env: {}, request: reqWithId("1", { id: "1" }) });
if (res.status !== 500) throw new Error("未绑定 KV PUT 应返回 500");

// ---- 参考库（危害因素库）----

// 初始为空
res = await libraryGet({ env });
let lib = await res.json();
if (res.status !== 200 || typeof lib !== "object" || Object.keys(lib).length !== 0) {
  throw new Error("参考库初始 GET 不符");
}

// 写入带时间戳的参考库
const sampleLib = {
  updatedAt: "2026-08-10T04:00:00.000Z",
  hazardFactors: [{ rec: "苯", name: "苯", dust: "", qual: "否", outsource: "否" }],
};
res = await libraryPut({ env, request: { json: async () => sampleLib } });
if (res.status !== 200) throw new Error("参考库 PUT 失败: " + res.status);

// 读回（KV 往返，保留 updatedAt）
res = await libraryGet({ env });
lib = await res.json();
if (lib.updatedAt !== sampleLib.updatedAt || lib.hazardFactors.length !== 1) {
  throw new Error("参考库 KV 往返失败");
}

// 数组等非法格式 → 400
res = await libraryPut({ env, request: { json: async () => [1, 2, 3] } });
if (res.status !== 400) throw new Error("参考库非法格式应返回 400");

// 未绑定 KV → 500
res = await libraryGet({ env: {} });
if (res.status !== 500) throw new Error("参考库未绑定 KV GET 应返回 500");
res = await libraryPut({ env: {}, request: { json: async () => sampleLib } });
if (res.status !== 500) throw new Error("参考库未绑定 KV PUT 应返回 500");

console.log("Cloudflare Pages Functions 逻辑校验通过 ✔");
