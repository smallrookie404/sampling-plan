// 真实登录验证：用 JS 移植版（PNG解码+kNN+RSA+登录）直接登平台，看能否拿到 token
import fs from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";
import { loginAccount, b64ToBytes, decodePng, recognize, rsaEncrypt } from "../functions/api/bridge/_captcha-core.mjs";
import { CAPTCHA_TEMPLATES } from "../functions/api/bridge/_captcha-templates.mjs";

const API = "http://223.93.144.122:27800";
const __dirname = path.dirname(fileURLToPath(import.meta.url));
const accounts = JSON.parse(fs.readFileSync(path.resolve(__dirname, "..", "..", "captcha-ocr", "accounts.json"), "utf8"));

const who = process.argv[2] || "btc-gxz";
const acc = accounts.find((a) => a.username.toLowerCase() === who.toLowerCase());
if (!acc) { console.error("账号不在 accounts.json：" + who); process.exit(1); }

console.log("=== 取一张验证码并识别 ===");
const r = await fetch(API + "/auth/code");
const code = await r.json();
const dec = await decodePng(b64ToBytes(String(code.img).split(",").pop()));
const rec = recognize(dec, CAPTCHA_TEMPLATES);
console.log("  尺寸", dec.w + "x" + dec.h, "｜ 识别", rec ? rec.expr + "=" + rec.answer : "失败", "｜ conf", rec ? rec.conf.toFixed(2) : "-");
console.log("  密文长度", b64ToBytes(rsaEncrypt(acc.password)).length, "字节");

console.log("\n=== loginAccount（最多 4 张，置信度门限 0.25）===");
const got = await loginAccount({
  apiBase: API, username: acc.username, password: acc.password,
  templates: CAPTCHA_TEMPLATES, confThreshold: 0.25, maxTries: 4,
  onLog: (m) => console.log("  · " + m),
});
if (!got) { console.log("\n❌ 登录失败"); process.exit(1); }
console.log("\n✅ 登录成功：", got.user, "｜ 算式", got.via, "｜ conf", got.conf);
console.log("  token 前缀：", got.token.slice(0, 24) + "…");

console.log("\n=== 用该 token 校验 /auth/info ===");
const info = await fetch(API + "/auth/info", {
  headers: { Authorization: got.token, organizationId: String(acc.orgId || 1) },
});
const j = await info.json();
console.log("  HTTP", info.status, "→", (j.user || {}).userName, (j.user || {}).userCode);
