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
const platform = { code: 0, login: 0, info: 0, brokenCode: false, rejectLogin: false };

globalThis.fetch = async (url, init = {}) => {
  const u = String(url);
  const hdr = new Headers(init.headers || {});
  const auth = hdr.get("authorization") || "";
  const ok = (obj, status = 200) => new Response(JSON.stringify(obj), { status, headers: { "Content-Type": "application/json" } });
  if (u.endsWith("/auth/code")) {
    platform.code++;
    if (platform.brokenCode) return ok({ message: "平台取验证码异常" }, 500);
    return ok({ uuid: "u-" + platform.code, img: "data:image/png;base64," + SAMPLE_B64 });
  }
  if (u.endsWith("/auth/login")) {
    platform.login++;
    const body = JSON.parse(init.body);
    if (!platform.rejectLogin && body.code === 35) {
      return ok({ user: { userName: "高循洲", userCode: "BTC-GXZ", organizationId: 7 }, token: ISSUED_TOKEN });
    }
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
// POST 版本（带 JSON body）
const ctxPost = (path, body, headers = {}) => ({
  request: new Request("https://sampling-plan.pages.dev/api/bridge/" + path, {
    method: "POST",
    headers: { "Content-Type": "application/json", ...headers },
    body: typeof body === "string" ? body : JSON.stringify(body),
  }),
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

// 9) 登录页验证码自动填写：/code 只取图+OCR，不登录、不需会话
const budgetKey = "bridge:captcha:" + new Date().toISOString().slice(0, 10);
platform.login = 0; platform.code = 0;
r = await mod.onRequest(ctx("code"));
j = await r.json();
check("/code 无需 Authorization 即返回 uuid+code",
  r.status === 200 && j.ok === true && !!j.uuid && j.code === "35",
  "HTTP " + r.status + " ／ " + JSON.stringify({ uuid: j.uuid, code: j.code, conf: j.conf, expr: j.expr }));
check("/code 回传图片供用户核对", typeof j.img === "string" && j.img.startsWith("data:image/png;base64,"));
check("/code 不做登录（login=0）", platform.login === 0, "login=" + platform.login);

// 10) /code 不涉及账号：即使未配置 BRIDGE_ACCOUNTS 也能用
const r3 = await mod.onRequest({ ...ctx("code"), env: { SAMPLING_RECORDS: kv } });
const j3 = await r3.json();
check("/code 未配置账号时仍可用（不读账号池）", r3.status === 200 && j3.ok === true && j3.code === "35", "HTTP " + r3.status);

// 11) /code 计入每日用量
check("/code 计入验证码用量（bridge:captcha:<日期>）", Number(store.get(budgetKey) || 0) > 0, String(store.get(budgetKey)));

// 12) /code 跨站调用 → 403
r = await mod.onRequest(ctx("code", { Origin: "https://evil.example" }));
check("/code 跨站 Origin → 403", r.status === 403, "HTTP " + r.status);

// 13) /code 达到每日上限 → 429，且提示改手输
store.set(budgetKey, "300");
r = await mod.onRequest(ctx("code"));
j = await r.json();
check("/code 达到每日上限 → 429 并提示手动输入",
  r.status === 429 && /手动输入/.test(j.message || ""), "HTTP " + r.status + " ／ " + j.message);
store.delete(budgetKey);

// 14) 平台取验证码一直异常 → /code 返回 502（前端据此回退手输）
platform.brokenCode = true; platform.code = 0;
r = await mod.onRequest(ctx("code"));
j = await r.json();
check("/code 识别失败 → 502 并提示手动输入",
  r.status === 502 && /手动输入/.test(j.message || ""), "HTTP " + r.status + " ／ " + j.message);
check("/code 失败时按 CODE_MAX_TRIES 换图重试（3 次）", platform.code === 3, "code=" + platform.code);
platform.brokenCode = false;

// 15) 登录页云端代登录：POST /login（账号密码来自浏览器，验证码全程在云端）
const loginKey = "bridge:login:" + new Date().toISOString().slice(0, 10);
platform.login = 0; platform.code = 0;
r = await mod.onRequest(ctxPost("login", { username: "btc-wf", password: "pw-123" }));
j = await r.json();
check("/login 无需 Authorization 即完成登录",
  r.status === 200 && j.ok === true && j.token === ISSUED_TOKEN,
  "HTTP " + r.status + " ／ " + JSON.stringify({ ok: j.ok, via: j.via, conf: j.conf }));
check("/login 回传用户与机构（供前端解析组织）",
  !!(j.user && j.user.userName === "高循洲") && j.orgId === 7,
  JSON.stringify({ user: j.user, orgId: j.orgId }));
check("/login 真跑 OCR 且只登录 1 次", platform.login === 1 && platform.code === 1,
  "login=" + platform.login + " code=" + platform.code);
check("/login 不回传验证码答案（uuid/code 都不出现）",
  !("code" in j) && !("uuid" in j), JSON.stringify(Object.keys(j)));
check("/login 不写令牌缓存到 KV（凭据换来的令牌不持久化）", !store.has("bridge:token:btc-wf"));
check("/login 计入登录用量（bridge:login:<日期>）", Number(store.get(loginKey) || 0) === 1, String(store.get(loginKey)));

// 16) /login 参数与协议校验
r = await mod.onRequest(ctxPost("login", { username: "btc-wf" }));
check("/login 缺密码 → 400", r.status === 400, "HTTP " + r.status);
r = await mod.onRequest(ctxPost("login", "not-json"));
check("/login 请求体非 JSON → 400", r.status === 400, "HTTP " + r.status);
r = await mod.onRequest(ctxPost("login", { username: "a", password: "b" }, { Origin: "https://evil.example" }));
check("/login 跨站 Origin → 403", r.status === 403, "HTTP " + r.status);
r = await mod.onRequest(ctx("login"));
check("/login 用 GET → 405", r.status === 405, "HTTP " + r.status);

// 17) 平台一直拒绝（OCR 猜错/密码错）→ 502，并提示可改用验证码登录
platform.rejectLogin = true; platform.login = 0; platform.code = 0;
r = await mod.onRequest(ctxPost("login", { username: "btc-wf", password: "pw" }));
j = await r.json();
check("/login 平台拒绝 → 502 且提示改用验证码登录",
  r.status === 502 && /改用验证码登录/.test(j.message || ""), "HTTP " + r.status + " ／ " + j.message);
check("/login 失败按 MAX_TRIES 重试（4 次换图）", platform.login === 4 && platform.code === 4,
  "login=" + platform.login + " code=" + platform.code);
platform.rejectLogin = false;

// 18) /login 每日上限 → 429
store.set(loginKey, "200");
r = await mod.onRequest(ctxPost("login", { username: "btc-wf", password: "pw" }));
j = await r.json();
check("/login 达每日上限 → 429 并提示改用验证码登录",
  r.status === 429 && /改用验证码登录/.test(j.message || ""), "HTTP " + r.status + " ／ " + j.message);
store.delete(loginKey);

// 19) 未知路径
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
