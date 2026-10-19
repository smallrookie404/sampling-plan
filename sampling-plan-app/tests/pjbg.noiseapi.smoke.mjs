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

// ---- 多点位挤在同一单元格 + 有「测量结果[dB(A)]」列的噪声表（用于 G）----
// 每个点位固定 3 个测量值：第 1 个点位=前 3 行，第 2 个=第 4~6 行…
const cellLines = (lines) => "<w:tc>" + lines.map((t, i) =>
  (i ? "<w:r><w:br/></w:r>" : "") + "<w:r><w:t xml:space=\"preserve\">" + esc(t) + "</w:t></w:r>"
).join("") + "</w:tc>";
const rowCells = (cells) => "<w:tr>" + cells.map((c) =>
  (Array.isArray(c) ? cellLines(c) : "<w:tc>" + p(c) + "</w:tc>")).join("") + "</w:tr>";

const PTS_HDR = ["测量编号", "单元/工作场所", "岗位/工种/测量点/对象", "接触时间",
  "测量结果[dB(A)]", "LEX,8h", "LEX,W", "接触限值", "单项结论"];
const PTS_ROWS = [
  // 针车岗位：两点位挤在一个单元格 → 针车1 三个值都 ≥80（噪声点位），针车2 三个值都 <80（低噪点位）
  ["1", "车间A", ["针车岗位/针车1", "针车岗位/针车2"], "7h/d",
    ["84.2", "83.6", "83.1", "78.1", "77.7", "77.4"], "-", "83.9", "85", "符合"],
  // 压机岗位：压机1 全 <80（低噪），压机2 全 ≥80（噪声）→ 只删压机1
  ["2", "车间A", ["压机岗位/压机1", "压机岗位/压机2"], "7h/d",
    ["79.5", "79.1", "78.8", "80.2", "80.5", "80.1"], "-", "80.5", "85", "符合"],
  // 装配岗位：LEX ≥ 80 但两个点位的测量值都全 <80 → 没有「3 值全 ≥80」的点位 → 不判定（前提不成立）
  ["3", "车间C", ["装配岗位/装配1", "装配岗位/装配2"], "7h/d",
    ["79.0", "78.0", "77.0", "76.0", "75.0", "74.0"], "-", "80.5", "85", "符合"],
  // 打磨岗位：整岗位 LEX < 80 → 走②的整岗位规则，不进低噪点位列表
  ["4", "车间A", "打磨岗位/打磨", "4h", ["75.0", "74.5", "74.8"], "-", "74.8", "85", "符合"],
  // ★ 冲压岗位：两点位**各占一行**（不是挤在一格）→ 单点位单行不参与点位级判断
  //   （冲压1 84.9 ≥ 80，冲压2 78.0 < 80；若误按点位判断会把冲压2 判成低噪点位）
  ["5", "车间D", "冲压岗位/冲压1", "7h/d", ["84.9", "84.3", "84.1"], "-", "84.4", "85", "符合"],
  ["6", "车间D", "冲压岗位/冲压2", "7h/d", ["78.0", "77.6", "77.3"], "-", "77.6", "85", "符合"],
];
const docXmlPoints =
  '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
  + "<w:document " + W + "><w:body>"
  + p("7.3 噪声检测结果分析与评价")
  + "<w:tbl>" + rowCells(PTS_HDR) + PTS_ROWS.map(rowCells).join("") + "</w:tbl>"
  + p("7.4 结束")
  + "</w:body></w:document>";

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

console.log("\n[G] 行内多点位：按「每点位 3 个测量值」切分，3 值全 < 80 的点位进 quietPointPairs");
{
  const r = await page.evaluate((xml) => {
    const info = window.SamplingUpload.noiseInfoFromXml(xml);
    return {
      ok: info.ok,
      quiet: info.quietPointPairs,
      nonNoise: info.nonNoiseGroups,
      noise: info.noiseWorkshopGroups,
    };
  }, docXmlPoints);
  console.log("   ->", JSON.stringify(r));
  check(r.ok === true, "解析成功（ok=true）");
  check(JSON.stringify(r.nonNoise) === JSON.stringify([
    { unit: "车间A", jobs: ["打磨岗位"] },
  ]), "非噪声车间岗位仍只含打磨岗位（低噪点位不影响整岗位归类）", JSON.stringify(r.nonNoise));
  check(JSON.stringify(r.noise) === JSON.stringify([
    { unit: "车间A", jobs: ["针车岗位", "压机岗位"] },
    { unit: "车间C", jobs: ["装配岗位"] },
    { unit: "车间D", jobs: ["冲压岗位"] },
  ]), "噪声车间岗位 = 车间A 针车/压机 + 车间C 装配 + 车间D 冲压", JSON.stringify(r.noise));
  check(JSON.stringify(r.quiet) === JSON.stringify([
    { unit: "车间A", job: "针车岗位", point: "针车2" },
    { unit: "车间A", job: "压机岗位", point: "压机1" },
  ]), "低噪点位 = 车间A针车岗位针车2 + 车间A压机岗位压机1", JSON.stringify(r.quiet));
  // ★ 守门：冲压岗位两点位各占一行（不是「多行挤在一格」）→ 即使冲压2 三个值都 <80 也不判定
  check(!(r.quiet || []).some((q) => q.job === "冲压岗位"),
    "单点位单行不参与点位级判断（冲压岗位无低噪点位）", JSON.stringify(r.quiet));
}

console.log("\n[F] 页面无 JS 错误");
check(pageErrors.length === 0, "无页面错误", pageErrors.join(" | "));

console.log(`\n结果：${pass} 通过 / ${fail} 失败`);
await browser.close();
process.exit(fail ? 1 : 0);
