/*
 * 云端账号桥（functions/api/bridge/[[path]].js）逻辑验证
 *
 * 用假的 KV + 假的平台 fetch 跑通全流程；验证码用 captcha-ocr/samples/0000.png 真图，
 * 因此 OCR（PNG 解码 + 分割 + kNN）与 RSA 加密都会被真实执行。
 * 需要 samples/（未入库）；缺失时自动跳过。
 */
import fs from "node:fs";
import path from "node:path";
import { fileURLToPath, pathToFileURL } from "node:url";

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const APP_DIR = path.resolve(__dirname, "..");
const FN = path.join(APP_DIR, "functions", "api", "bridge", "[[path]].js");
const SAMPLE = path.resolve(APP_DIR, "..", "captcha-ocr", "samples", "0000.png");

const results = [];
const check = (name, ok, extra) => results.push({ name, ok, extra });

if (!fs.existsSync(SAMPLE)) {
  console.log("未找到 captcha-ocr/samples/0000.png，跳过云端账号桥测试");
  process.exit(0);
}

// 把同目录相对 import 改写成绝对 file:// URL，再用 data URL 加载
// （本仓库无 package.json，`.js` 默认按 CJS 解析，故不能用 file:// 直接 import）
function loadAsModule(file) {
  let code = fs.readFileSync(file, "utf8");
  code = code.replace(/"\.\/(_captcha-[a-z-]+\.mjs)"/g, (_m, f) =>
    JSON.stringify(pathToFileURL(path.join(APP_DIR, "functions", "api", "bridge", f)).href));
  return import("data:text/javascript;base64," + Buffer.from(code).toString("base64"));
}

// ---- 假平台 ----
const SAMPLE_B64 = fs.readFileSync(SAMPLE).toString("base64");
const CALLER_TOKEN = "Bearer CALLER-SESSION";
const ISSUED_TOKEN = "Bearer ISSUED-GXZ";
const platform = { code: 0, login: 0, info: 0 };

globalThis.fetch = async (url, init = {}) => {
  const u = String(url);
  const hdr = new Headers(init.headers || {});
  const auth = hdr.get("authorization") || "";
  const ok = (obj, status = 200) => new Response(JSON.stringify(obj), { status, headers: { "Content-Type": "application/json" } });
  if (u.endsWith("/auth/code")) {
    platform.code++;
    return ok({ uuid: "u-" + platform.code, img: "data:image/png;base64," + SAMPLE_B64 });
  }
  if (u.endsWith("/auth/login")) {
    platform.login++;
    const body = JSON.parse(init.body);
    if (body.code === 35) return ok({ user: { userName: "高循洲" }, token: ISSUED_TOKEN });  // 0000.png = 7×5=35
    return ok({ message: "验证码错误" }, 400);
  }
  if (u.endsWith("/auth/info")) {
    platform.info++;
    if (auth === CALLER_TOKEN || auth === ISSUED_TOKEN) return ok({ user: { userName: "x" } });
    return ok({ message: "找不到当前登录的信息" }, 401);
  }
  return ok({ message: "not found " + u }, 404);
};

// ---- 假 KV ----
const store = new Map();
const kv = {
  async get(k) { return store.has(k) ? store.get(k) : null; },
  async put(k, v) { store.set(k, v); },
};

const mod = await loadAsModule(FN);
const env = {
  SAMPLING_RECORDS: kv,
  BRIDGE_ACCOUNTS: JSON.stringify([{ username: "btc-gxz", password: "dummy-pass", orgId: 1 }]),
};
const ctx = (path, headers = {}, method = "GET") => ({
  request: new Request("https://sampling-plan.pages.dev/api/bridge/" + path, { method, headers }),
  env,
  params: { path: path.split("/") },
});

// 1) ping
let r = await mod.onRequest(ctx("ping"));
let j = await r.json();
check("ping 返回 ok/cloud", r.status === 200 && j.ok === true && j.cloud === true, JSON.stringify(j));

// 2) 无 Authorization → 401（公开接口不能被任意人拿去换令牌）
r = await mod.onRequest(ctx("token/btc-gxz"));
check("无 Authorization → 401", r.status === 401, "HTTP " + r.status);

// 3) 会话无效 → 401
r = await mod.onRequest(ctx("token/btc-gxz", { Authorization: "Bearer BAD" }));
check("会话无效 → 401", r.status === 401, "HTTP " + r.status);

// 4) 会话有效 → 真跑 OCR + 登录，签发令牌
store.clear(); platform.login = 0; platform.code = 0;
r = await mod.onRequest(ctx("token/btc-gxz", { Authorization: CALLER_TOKEN }));
j = await r.json();
check("有效会话 → 签发令牌（真跑 OCR）",
  r.status === 200 && j.ok === true && j.token === ISSUED_TOKEN,
  "HTTP " + r.status + " ／ " + JSON.stringify({ ok: j.ok, via: j.via, conf: j.conf, cached: j.cached, logs: j.logs }));
check("OCR 识别正确并成功登录（1 次登录）", platform.login === 1, "login=" + platform.login + " code=" + platform.code);

// 5) 再次调用 → 命中 KV 缓存，不再登录
platform.login = 0; platform.code = 0;
r = await mod.onRequest(ctx("token/btc-gxz", { Authorization: CALLER_TOKEN }));
j = await r.json();
check("二次调用命中缓存（0 次登录 / 0 张验证码）",
  j.cached === true && j.token === ISSUED_TOKEN && platform.login === 0 && platform.code === 0,
  "cached=" + j.cached + " login=" + platform.login + " code=" + platform.code);

// 6) 缓存里的令牌失效 → 自动重登
store.set("bridge:token:btc-gxz", JSON.stringify({ token: "Bearer DEAD", orgId: 1 }));
platform.login = 0;
r = await mod.onRequest(ctx("token/btc-gxz", { Authorization: CALLER_TOKEN }));
j = await r.json();
check("缓存令牌失效 → 自动重登", j.ok === true && j.token === ISSUED_TOKEN && j.cached === false && platform.login === 1,
  "cached=" + j.cached + " login=" + platform.login);

// 7) 未配置账号 → 503 且给出配置指引
const r2 = await mod.onRequest({ ...ctx("token/btc-wcy", { Authorization: CALLER_TOKEN }), env: { SAMPLING_RECORDS: kv } });
j = await r2.json();
check("未配置账号 → 503 并提示 BRIDGE_ACCOUNTS", r2.status === 503 && /BRIDGE_ACCOUNTS/.test(j.message), "HTTP " + r2.status);

// 8) 验证码用量计数写入 KV
check("验证码用量写入 KV（bridge:captcha:<日期>）",
  Number(store.get("bridge:captcha:" + new Date().toISOString().slice(0, 10)) || 0) > 0,
  String(store.get("bridge:captcha:" + new Date().toISOString().slice(0, 10))));

// 9) 未知路径
r = await mod.onRequest(ctx("nope"));
check("未知路径 → 404", r.status === 404, "HTTP " + r.status);

console.log("\n=== 云端账号桥（Pages Function）===");
let pass = 0;
for (const x of results) {
  console.log((x.ok ? "✅" : "❌") + " " + x.name + (x.extra ? "  → " + x.extra : ""));
  if (x.ok) pass++;
}
console.log("\n" + pass + "/" + results.length + " 通过");
process.exit(pass === results.length ? 0 : 1);
