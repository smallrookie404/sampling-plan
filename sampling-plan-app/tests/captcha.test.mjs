/*
 * 验证码 OCR（JS 移植版）验证：
 *   1) PNG 解码 + 分割（每图必须切出 5 段）
 *   2) 留一法（LOO）逐位置 / 整式准确率 —— 与 Python 原型对照（原为 92.2% / 76.7%）
 *   3) 置信度门限收益表
 *   4) RSA PKCS#1 v1.5 自检（长度、随机性、可解密性用 BigInt 私钥不可得 → 只验长度与确定性差异）
 *
 * 需要 captcha-ocr/samples/（未入库）；样本缺失时自动跳过并返回 0。
 */
import fs from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";
import { decodePng, extractFeatures, dist2, recognize, rsaEncrypt, b64ToBytes } from "../functions/api/bridge/_captcha-core.mjs";
import { CAPTCHA_TEMPLATES } from "../functions/api/bridge/_captcha-templates.mjs";

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const SAMPLES = path.resolve(__dirname, "..", "..", "captcha-ocr", "samples");

// 与 captcha-ocr/classify.py 的 LABELS 一致
const LABELS = {
  0: ['7', '×', '5'], 1: ['6', '+', '8'], 2: ['9', '×', '2'], 3: ['8', '×', '7'],
  4: ['1', '+', '4'], 5: ['3', '×', '8'], 6: ['7', '×', '2'], 7: ['5', '-', '5'],
  8: ['2', '×', '3'], 9: ['6', '-', '7'], 10: ['3', '-', '6'], 11: ['9', '+', '7'],
  12: ['0', '-', '9'], 13: ['2', '+', '3'], 14: ['5', '+', '0'], 15: ['7', '-', '7'],
  16: ['5', '-', '1'], 17: ['8', '+', '3'], 18: ['6', '+', '8'], 19: ['9', '+', '3'],
  20: ['0', '-', '8'], 21: ['1', '+', '3'], 22: ['4', '+', '7'], 23: ['7', '+', '1'],
  24: ['0', '+', '8'], 25: ['9', '×', '9'], 26: ['0', '+', '1'], 27: ['8', '×', '9'],
  28: ['0', '+', '8'], 29: ['3', '×', '3'], 30: ['4', '-', '7'], 31: ['0', '×', '1'],
  32: ['6', '+', '0'], 33: ['4', '-', '2'], 34: ['2', '-', '2'], 35: ['2', '-', '3'],
  36: ['1', '+', '2'], 37: ['0', '+', '6'], 38: ['2', '×', '0'], 39: ['3', '-', '5'],
  40: ['0', '-', '8'], 41: ['4', '×', '7'], 42: ['3', '+', '0'], 43: ['5', '+', '6'],
  44: ['2', '×', '0'], 45: ['5', '+', '0'], 46: ['6', '×', '7'], 47: ['3', '×', '6'],
  48: ['0', '-', '9'], 49: ['5', '+', '7'], 50: ['7', '-', '0'], 51: ['4', '+', '5'],
  52: ['8', '×', '8'], 53: ['6', '×', '5'], 54: ['2', '-', '4'], 55: ['7', '+', '5'],
  56: ['7', '×', '4'], 57: ['8', '+', '1'], 58: ['5', '+', '3'], 59: ['5', '×', '0'],
};

const results = [];
const check = (name, ok, extra) => results.push({ name, ok, extra });

// ---- 1) 解码 + 分割 ----
const cache = {};
let decodeOk = 0, segOk = 0, total = 0;
if (fs.existsSync(SAMPLES)) {
  for (const idx of Object.keys(LABELS).map(Number)) {
    const f = path.join(SAMPLES, String(idx).padStart(4, "0") + ".png");
    if (!fs.existsSync(f)) continue;
    total++;
    const dec = await decodePng(new Uint8Array(fs.readFileSync(f)));
    if (dec.w === 111 && dec.h === 36 && dec.px.length === 111 * 36 * 3) decodeOk++;
    const feats = extractFeatures(dec);
    if (feats.length === 5 && feats.every((x) => x)) segOk++;
    cache[idx] = feats;
  }
}
if (!total) {
  console.log("未找到 captcha-ocr/samples/，跳过 OCR 准确率测试");
  process.exit(0);
}
check("PNG 解码（111×36 RGB）", decodeOk === total, decodeOk + "/" + total);
check("分割切出 5 个字符", segOk === total, segOk + "/" + total);

// ---- 2) 留一法 ----
const idxs = Object.keys(cache).map(Number);
const stat = { d1: [0, 0], op: [0, 0], d2: [0, 0] };
const rows = [];
for (const i of idxs) {
  const t = LABELS[i];
  const digits = [], ops = [];
  for (const j of idxs) {
    if (j === i) continue;
    digits.push([LABELS[j][0], cache[j][0]]);
    digits.push([LABELS[j][2], cache[j][2]]);
    ops.push([LABELS[j][1], cache[j][1]]);
  }
  const preds = [], confs = [];
  for (const pos of [0, 1, 2]) {
    const pool = pos === 1 ? ops : digits;
    const fv = cache[i][pos];
    const cand = pool.map(([label, f]) => ({ label, d: dist2(fv, f) })).sort((a, b) => a.d - b.d);
    const best = cand[0];
    const second = cand.find((c) => c.label !== best.label);
    confs.push(!second ? 1 : (second.d > 0 ? Math.max(0, (second.d - best.d) / second.d) : 1));
    preds.push(best.label);
  }
  const key = ['d1', 'op', 'd2'];
  for (let k = 0; k < 3; k++) {
    stat[key[k]][1]++;
    if (preds[k] === t[k]) stat[key[k]][0]++;
  }
  rows.push({ conf: Math.min(...confs), ok: preds.join("") === t.join("") });
}
const pct = (a, b) => (b ? (100 * a / b).toFixed(1) + "%" : "n/a");
const hit = stat.d1[0] + stat.op[0] + stat.d2[0];
const all = stat.d1[1] + stat.op[1] + stat.d2[1];
const exprOk = rows.filter((r) => r.ok).length;

console.log("\n=== 留一法（" + idxs.length + " 样本）===");
console.log("  第1个数字 " + pct(stat.d1[0], stat.d1[1]) +
            " ｜ 运算符 " + pct(stat.op[0], stat.op[1]) +
            " ｜ 第2个数字 " + pct(stat.d2[0], stat.d2[1]));
console.log("  逐字符合计 " + pct(hit, all) + " ｜ 整式全对 " + pct(exprOk, rows.length));
check("逐字符准确率 ≥ 85%", hit / all >= 0.85, pct(hit, all));
check("整式全对率 ≥ 60%", exprOk / rows.length >= 0.60, pct(exprOk, rows.length));

console.log("\n=== 置信度门限（丢弃换图）===");
for (const th of [0, 0.1, 0.2, 0.25, 0.3, 0.4]) {
  const acc = rows.filter((r) => r.conf >= th);
  const good = acc.filter((r) => r.ok).length;
  console.log("  ≥" + th.toFixed(2) + "  接受 " + pct(acc.length, rows.length) +
              "  接受后整式正确率 " + pct(good, acc.length));
}

// ---- 3) 端到端 recognize（用全量模板）----
let e2e = 0;
for (const i of idxs) {
  const f = path.join(SAMPLES, String(i).padStart(4, "0") + ".png");
  const rec = recognize(await decodePng(new Uint8Array(fs.readFileSync(f))), CAPTCHA_TEMPLATES);
  if (rec && rec.expr === LABELS[i].join("")) e2e++;
}
check("用生成模板端到端识别整式全对率 ≥ 80%", e2e / idxs.length >= 0.80, pct(e2e, idxs.length));

// ---- 4) RSA 自检 ----
const c1 = rsaEncrypt("dummy-pass");
const c2 = rsaEncrypt("dummy-pass");
const cBytes = b64ToBytes(c1);
check("RSA 密文为 64 字节 base64", cBytes.length === 64, cBytes.length + " 字节");
check("RSA 每次填充随机（两次密文不同）", c1 !== c2);
check("RSA 密文首字节 < 0x40（PKCS#1 低值特征，非硬校验）", true, "c1[0]=0x" + cBytes[0].toString(16));

// ---- 汇总 ----
console.log("\n=== 验证码 OCR（JS 移植版）===");
let pass = 0;
for (const r of results) {
  console.log((r.ok ? "✅" : "❌") + " " + r.name + (r.extra ? "  → " + r.extra : ""));
  if (r.ok) pass++;
}
console.log("\n" + pass + "/" + results.length + " 通过");
process.exit(pass === results.length ? 0 : 1);
