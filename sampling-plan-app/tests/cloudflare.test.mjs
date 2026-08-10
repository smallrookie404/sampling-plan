// Cloudflare Pages Functions 逻辑验证（本地用模拟 KV 测试，无需真实账号）
import fs from "node:fs";

// 以 data URL 方式把函数源码作为 ESM 加载，避免影响仓库内其他 CommonJS 文件
async function loadFunction(rel) {
  const code = fs.readFileSync(new URL(rel, import.meta.url), "utf8");
  return import("data:text/javascript;base64," + Buffer.from(code).toString("base64"));
}

const healthMod = await loadFunction("../functions/api/health.js");
const recordsMod = await loadFunction("../functions/api/records.js");
const { onRequestGet: healthGet } = healthMod;
const { onRequestGet, onRequestPut } = recordsMod;

function fakeKV() {
  let store = null;
  return {
    async get() {
      return store;
    },
    async put(_k, v) {
      store = v;
    },
  };
}

const kv = fakeKV();
const env = { SAMPLING_RECORDS: kv };

// 健康检查
const health = await healthGet();
if (health.status !== 200) throw new Error("health 失败: " + health.status);

// 初始为空
let res = await onRequestGet({ env });
let list = await res.json();
if (res.status !== 200 || !Array.isArray(list) || list.length !== 0) throw new Error("初始 GET 不符");

// 写入
res = await onRequestPut({ env, request: { json: async () => [{ id: "1", name: "云测试记录", rows: [] }] } });
if (res.status !== 200) throw new Error("PUT 失败: " + res.status);

// 读回（KV 往返）
res = await onRequestGet({ env });
list = await res.json();
if (list.length !== 1 || list[0].name !== "云测试记录") throw new Error("KV 往返失败");

// 非法格式 → 400
res = await onRequestPut({ env, request: { json: async () => ({ a: 1 }) } });
if (res.status !== 400) throw new Error("非法格式应返回 400");

// 未绑定 KV → 500，且提示清晰
res = await onRequestGet({ env: {} });
if (res.status !== 500) throw new Error("未绑定 KV 应返回 500");
res = await onRequestPut({ env: {}, request: { json: async () => [] } });
if (res.status !== 500) throw new Error("未绑定 KV PUT 应返回 500");

console.log("Cloudflare Pages Functions 逻辑校验通过 ✔");
