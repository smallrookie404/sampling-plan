// 真实浏览器冒烟：SamplingUpload 的噪声信息 API（noiseInfoFromXml / getCachedNoiseInfo）。
//
// 为什么需要单独一条浏览器测试：这两个函数依赖**原生 DOMParser**（Node 里没有可靠实现，
// vm 沙箱里只能打桩），Node 侧的 pjbg.noisereuse.test.mjs 只验证了 blf.js 的接线，
// 验证不了「真的能把噪声表从 document.xml 里解析出来」。
//
// 覆盖：
//   A. noiseInfoFromXml 能从 document.xml 字符串解析出噪声表，并按 LEX 最大值正确分组
//      （< 80 → 非噪声；≥ 80 → 噪声车间；≥ 85 → 超标）；
//   B. 「岗位/工种/测量点/对象」按首个「/」拆出「测量点/对象」，岗位取「/」前部分；
//   C. nonNoiseText 格式为「车间岗位；车间岗位」，与页签「筛选结果」一致；
//   D. getCachedNoiseInfo 在未解析过的 buffer 上返回 null（只读，不触发 IO）；
//   E. 找不到噪声表时返回 ok=false 且各列表为空（不抛异常）。
//
// 用法：node tests/pjbg.noiseapi.smoke.mjs
import { chromium } from "playwright-core";
import fs from "node:fs";
import path from "node:path";
import { startServer } from "../server.mjs";
import { findBrowser } from "./_browser.mjs";

const CHROME = findBrowser();
if (!CHROME) { console.error("未找到 Chrome/Edge，跳过该冒烟测试"); process.exit(0); }

const TEST_DIR = path.resolve("data/_test-noiseapi");
const TEST_DATA = path.join(TEST_DIR, "records.json");
fs.mkdirSync(TEST_DIR, { recursive: true });

const server = await startServer({ port: 0, dataFile: TEST_DATA });
const PORT = server.address().port;
const APP_URL = `http://127.0.0.1:${PORT}/`;
const browser = await chromium.launch({ executablePath: CHROME, headless: true });
process.on("exit", () => {
  try { server.close(); } catch {}
  try { fs.rmSync(TEST_DIR, { recursive: true, force: true }); } catch {}
  try { if (typeof browser !== "undefined" && browser) browser.process()?.kill(); } catch {}
});

const page = await browser.newPage({ viewport: { width: 1600, height: 900 } });
page.on("dialog", (d) => d.accept());
const pageErrors = [];
page.on("pageerror", (e) => pageErrors.push(e.message));
await page.addInitScript(() => {
  sessionStorage.setItem("xcdc_session_v1", JSON.stringify({
    token: "test-token", orgId: "1",
    userInfo: { userName: "测试账号", userCode: "TEST", organization: { organizationId: "1" } },
  }));
});

let pass = 0, fail = 0;
const check = (cond, name, extra) => {
  if (cond) { pass++; console.log(`  ✔ ${name}`); }
  else { fail++; console.error(`  ✘ ${name}${extra ? " :: " + extra : ""}`); }
};

// ---- 构造 document.xml（噪声表在「噪声检测结果分析与评价」标题段落之后）----
const W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"';
const esc = (s) => String(s).replace(/&/g, "&amp;").replace(/</g, "&lt;");
const p = (t) => "<w:p><w:r><w:t xml:space=\"preserve\">" + esc(t) + "</w:t></w:r></w:p>";
const row = (cells) => "<w:tr>" + cells.map((c) => "<w:tc>" + p(c) + "</w:tc>").join("") + "</w:tr>";

const NOISE_HDR = ["测量编号", "单元/工作场所", "岗位/工种/测量点/对象", "接触时间", "测量起止时间",
  "LEX,8h", "接触限值", "单项结论"];
const NOISE_ROWS = [
  // 打磨岗位 75 dB → 非噪声（< 80）
  ["1", "车间A", "打磨岗位/1#点", "4h", "08:00~12:00", "75.0", "85", "合格"],
  // 装配岗位 88 dB → 噪声车间（≥ 80）且超标（≥ 85）
  ["2", "车间A", "装配岗位/2#点", "4h", "08:00~12:00", "88.0", "85", "超标"],
  // 巡检岗位 78 dB → 非噪声
  ["3", "车间B", "巡检岗位/3#点", "4h", "08:00~12:00", "78.0", "85", "合格"],
];

const docXml =
  '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
  + "<w:document " + W + "><w:body>"
  + p("7.3 噪声检测结果分析与评价")
  + "<w:tbl>" + row(NOISE_HDR) + NOISE_ROWS.map(row).join("") + "</w:tbl>"
  + p("7.4 结束")
  + "</w:body></w:document>";

// 没有噪声表的文档（用于 E）
const docXmlNoNoise =
  '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
  + "<w:document " + W + "><w:body>" + p("1 概述") + p("正文内容") + "</w:body></w:document>";

await page.goto(APP_URL, { waitUntil: "domcontentloaded" });
await page.waitForTimeout(600);
await page.evaluate(() => document.querySelector("#xcdc-login")?.remove());

const hasApi = await page.evaluate(() => !!(window.SamplingUpload && typeof window.SamplingUpload.noiseInfoFromXml === "function"));
check(hasApi, "SamplingUpload.noiseInfoFromXml 已暴露");

const info = await page.evaluate((xml) => {
  const r = window.SamplingUpload.noiseInfoFromXml(xml);
  return {
    ok: r.ok,
    nonNoiseGroups: r.nonNoiseGroups,
    nonNoisePairs: r.nonNoisePairs,
    nonNoiseText: r.nonNoiseText,
    noiseWorkshopGroups: r.noiseWorkshopGroups,
  };
}, docXml);

console.log("\n[A] 分组结果");
console.log("   ->", JSON.stringify(info));
check(info.ok === true, "解析成功（ok=true）");
check(JSON.stringify(info.nonNoiseGroups) === JSON.stringify([
  { unit: "车间A", jobs: ["打磨岗位"] },
  { unit: "车间B", jobs: ["巡检岗位"] },
]), "非噪声车间岗位 = 车间A打磨岗位 + 车间B巡检岗位", JSON.stringify(info.nonNoiseGroups));
check(JSON.stringify(info.noiseWorkshopGroups) === JSON.stringify([
  { unit: "车间A", jobs: ["装配岗位"] },
]), "噪声车间岗位 = 车间A装配岗位", JSON.stringify(info.noiseWorkshopGroups));

console.log("\n[B] 岗位按「/」拆分");
check(info.nonNoisePairs.length === 2, "pairs 共 2 条", JSON.stringify(info.nonNoisePairs));
check(info.nonNoisePairs[0] && info.nonNoisePairs[0].job === "打磨岗位",
  "岗位取「/」前的部分（打磨岗位，不含 1#点）", JSON.stringify(info.nonNoisePairs[0]));

console.log("\n[C] 文本格式");
check(info.nonNoiseText === "车间A打磨岗位；车间B巡检岗位",
  "nonNoiseText = 车间A打磨岗位；车间B巡检岗位", info.nonNoiseText);

console.log("\n[D] getCachedNoiseInfo 只读");
const cached = await page.evaluate(() => {
  const fakeBuf = new ArrayBuffer(8);
  const r = window.SamplingUpload.getCachedNoiseInfo(fakeBuf);
  return { isNull: r === null, isUndefined: r === undefined };
});
check(cached.isNull === true, "未解析过的 buffer → 返回 null（不触发 IO）", JSON.stringify(cached));

console.log("\n[E] 无噪声表 → ok=false，不抛异常");
const bad = await page.evaluate((xml) => {
  try {
    const r = window.SamplingUpload.noiseInfoFromXml(xml);
    return { threw: false, ok: r.ok, n: (r.nonNoiseGroups || []).length, w: (r.noiseWorkshopGroups || []).length };
  } catch (e) { return { threw: true, msg: String(e && e.message) }; }
}, docXmlNoNoise);
check(bad.threw === false, "不抛异常");
check(bad.ok === false && bad.n === 0 && bad.w === 0, "ok=false 且各列表为空", JSON.stringify(bad));

console.log("\n[F] 页面无 JS 错误");
check(pageErrors.length === 0, "无页面错误", pageErrors.join(" | "));

console.log(`\n结果：${pass} 通过 / ${fail} 失败`);
await browser.close();
process.exit(fail ? 1 : 0);
