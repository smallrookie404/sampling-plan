/*
 * 生成 js/captcha-templates.mjs —— kNN 模板
 *
 * 用 captcha-ocr/samples/ 里**人工标注**的 60 张验证码，按 captcha-core.mjs 的
 * 特征提取流程算出模板（数字/运算符各一份），供 Cloudflare Worker / 浏览器识别验证码。
 *
 * ⚠ 必须与 captcha-core.mjs 的特征实现保持一致：模板与待识别图用同一套特征，kNN 才自洽。
 * ⚠ samples/ 未入库（.gitignore），本脚本只在有样本的机器上跑；产出的 templates 文件入库。
 *
 * 用法： node scripts/build-captcha-templates.mjs
 */
import fs from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";
import { decodePng, extractFeatures } from "../functions/api/bridge/_captcha-core.mjs";

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const APP_DIR = path.resolve(__dirname, "..");
const SAMPLES = path.resolve(APP_DIR, "..", "captcha-ocr", "samples");
const OUT = path.join(APP_DIR, "functions", "api", "bridge", "_captcha-templates.mjs");

// 与 captcha-ocr/classify.py 的 LABELS 完全一致（索引 → [第一个数字, 运算符, 第二个数字]）
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

if (!fs.existsSync(SAMPLES)) {
  console.error("找不到样本目录：" + SAMPLES + "（samples/ 未入库，需在有样本的机器上运行）");
  process.exit(1);
}

const round4 = (v) => Math.round(v * 10000) / 10000;
const digits = [];
const ops = [];
let used = 0, failed = [];

for (const idx of Object.keys(LABELS).map(Number).sort((a, b) => a - b)) {
  const file = path.join(SAMPLES, String(idx).padStart(4, "0") + ".png");
  if (!fs.existsSync(file)) { failed.push(idx + "(缺文件)"); continue; }
  let feats;
  try {
    feats = extractFeatures(await decodePng(new Uint8Array(fs.readFileSync(file))));
  } catch (e) {
    failed.push(idx + "(" + e.message + ")");
    continue;
  }
  const t = LABELS[idx];
  if (!feats || feats.length < 3 || feats.slice(0, 3).some((f) => !f)) {
    failed.push(idx + "(分割失败)");
    continue;
  }
  digits.push([t[0], feats[0].map(round4)]);
  ops.push([t[1], feats[1].map(round4)]);
  digits.push([t[2], feats[2].map(round4)]);
  used++;
}

const banner = `/*
 * 自动生成，请勿手改 —— 由 scripts/build-captcha-templates.mjs 从
 * captcha-ocr/samples/ 的 ${used} 张人工标注验证码算出（数字/运算符模板）。
 * 生成时间：${new Date().toISOString()}
 * 数字模板 ${digits.length} 条，运算符模板 ${ops.length} 条。
 */
export const CAPTCHA_TEMPLATES = {
  digits: [
`;
const body = (arr) => arr.map(([label, feat]) => `    ["${label}", [${feat.join(",")}]],`).join("\n");

fs.writeFileSync(
  OUT,
  banner + body(digits) + "\n  ],\n  ops: [\n" + body(ops) + "\n  ],\n};\n",
  "utf8"
);

console.log("已生成 " + OUT);
console.log("  使用样本 " + used + " 张；数字模板 " + digits.length + " 条，运算符模板 " + ops.length + " 条");
if (failed.length) console.log("  跳过 " + failed.length + " 张：" + failed.slice(0, 10).join(", "));
console.log("  文件大小 " + (fs.statSync(OUT).size / 1024).toFixed(0) + " KB");
