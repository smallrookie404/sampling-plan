/*
 * 主表大表性能实测（非正式测试，文件名以下划线开头）。
 * 用真实 Chrome 打开页面 → 上传一份 N 行的 xlsx → 测：
 *   ① 导入耗时（上传 → 表格出现 N 行）
 *   ② 滚动帧耗时（每帧推进 90px，连续 150 帧）
 *   ③ 编辑延迟（改一个输入框 → 依赖列文本更新）
 *
 * 用法：node tests/_perf_grid.mjs [行数]
 */
import fs from "node:fs";
import os from "node:os";
import path from "node:path";
import vm from "node:vm";
import { fileURLToPath } from "node:url";
import { createRequire } from "node:module";
import { chromium } from "playwright-core";
import { startServer } from "../server.mjs";
import { findBrowser } from "./_browser.mjs";

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const APP = path.join(__dirname, "..");
const require = createRequire(import.meta.url);
const X = require(path.join(APP, "js/xlsxio.js"));

const ROWS = Number(process.argv[2] || 3000);
const PORT = 8123;

// ---------- 造一份 N 行的 xlsx（结构与导出文件一致：表头 + 数据行） ----------
const libCtx = {};
vm.createContext(libCtx);
vm.runInContext(fs.readFileSync(path.join(APP, "js/data-core.js"), "utf8"), libCtx);
const headers = libCtx.SamplingData.mainHeaders;

const SHOPS = ["一车间", "二车间", "动力车间", "公用工程站"];
const POSTS = ["反应工", "离心工", "包装工", "巡检工", "班长"];
const POINTS = ["投料口", "反应釜区", "离心机旁", "包装线", "泵房", "配电室"];
const HAZ = ["噪声", "甲苯", "粉尘", "高温", "氯化氢", "工频电场"];
const rows = [headers];
for (let i = 0; i < ROWS; i++) {
  const r = new Array(headers.length).fill("");
  r[0] = SHOPS[Math.floor(i / 700) % SHOPS.length];
  r[1] = POSTS[Math.floor(i / 40) % POSTS.length];
  r[2] = POINTS[i % POINTS.length];
  r[3] = HAZ[i % HAZ.length];
  r[4] = "2";
  r[5] = "8";
  r[6] = "5";
  r[7] = String(1 + (i % 3));
  r[8] = "1";
  r[18] = "三班两运转";
  rows.push(r);
}
const buf = await X.writeWorkbook([{ name: "测点布局情况调查", rows }]);

// ---------- 起服务 + 打开页面 ----------
const server = await startServer({ port: PORT, dataFile: path.join(os.tmpdir(), `perf-records-${process.pid}.json`) });
const exe = findBrowser();
if (!exe) {
  console.error("未找到 Chrome/Edge，跳过");
  process.exit(0);
}
const HEADFUL = process.env.HEADFUL === "1";
const browser = await chromium.launch({ executablePath: exe, headless: !HEADFUL, args: HEADFUL ? [] : ["--disable-gpu"] });
const page = await browser.newPage({ viewport: { width: 1600, height: 900 } });
const errors = [];
page.on("pageerror", (e) => errors.push(String(e.message)));
page.on("dialog", (d) => d.accept().catch(() => {}));

await page.goto(`http://127.0.0.1:${PORT}/index.html`, { waitUntil: "domcontentloaded" });
await page.waitForSelector("#grid-body tr[data-r]", { timeout: 30000 });

// ---------- ① 导入 ----------
const tImport0 = Date.now();
await page.setInputFiles("#file-input", {
  name: "perf.xlsx",
  mimeType: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
  buffer: Buffer.from(buf),
});
await page.waitForFunction(
  (n) => (document.getElementById("grid-status")?.textContent || "").includes(`共 ${n} 行`),
  ROWS,
  { timeout: 180000 }
);
const tImport = Date.now() - tImport0;

// ---------- ② 滚动帧耗时 ----------
await page.evaluate(() => {
  document.getElementById("grid-wrap").scrollTop = 0;
});
// 先点一个单元格：让「布防框(arm)」存在——真实用户几乎总会先点击，
// 而 syncArmPosition 会在滚动帧里强制读取几何（强制同步布局）。
const armed = await page.evaluate(() => {
  const el = document.querySelector('td[data-r="2"][data-c="A"] input');
  if (!el) return false;
  el.dispatchEvent(new MouseEvent("mousedown", { bubbles: true, clientX: 1, clientY: 1 }));
  el.dispatchEvent(new MouseEvent("click", { bubbles: true, clientX: 1, clientY: 1 }));
  el.focus();
  return true;
});
await page.waitForTimeout(100);
await page.evaluate(() => {
  window.__frames = [];
  let last = performance.now();
  window.__recording = true;
  const tick = (t) => {
    window.__frames.push(t - last);
    last = t;
    if (window.__recording) requestAnimationFrame(tick);
  };
  requestAnimationFrame(tick);
});
await page.evaluate(async () => {
  const wrap = document.getElementById("grid-wrap");
  for (let i = 0; i < 150; i++) {
    wrap.scrollTop = wrap.scrollTop + 90;
    await new Promise((r) => requestAnimationFrame(r));
  }
});
const scroll = await page.evaluate(() => {
  window.__recording = false;
  const f = window.__frames.slice(3).sort((a, b) => a - b);
  const q = (p) => Math.round(f[Math.floor(f.length * p)] ?? 0);
  return {
    frames: f.length,
    median: q(0.5),
    p90: q(0.9),
    max: Math.round(f[f.length - 1] ?? 0),
    over33: f.filter((x) => x > 33).length,
    over50: f.filter((x) => x > 50).length,
  };
});

// ---------- ③ 编辑延迟（改「人数」→ 等「每班最大人数」列变化） ----------
await page.evaluate(() => {
  document.getElementById("grid-wrap").scrollTop = 0;
});
await page.waitForTimeout(300);
const edit = await page.evaluate(async () => {
  const inp = document.querySelector('td[data-r="3"][data-c="H"] input');
  const cell = document.querySelector('td[data-r="3"][data-c="AB"] .computed-text');
  if (!inp || !cell) return { error: "找不到目标单元格" };
  const before = cell.textContent;
  const t0 = performance.now();
  inp.value = before === "9" ? "5" : "9";
  inp.dispatchEvent(new Event("input", { bubbles: true }));
  const changed = await new Promise((resolve) => {
    const started = performance.now();
    const check = () => {
      if (cell.textContent !== before) resolve(true);
      else if (performance.now() - started > 3000) resolve(false);
      else requestAnimationFrame(check);
    };
    requestAnimationFrame(check);
  });
  return { ms: Math.round(performance.now() - t0), changed, before, after: cell.textContent };
});

// ---------- ④ 观察一次「选中整行」（会整窗重渲染） ----------
const rowClick = await page.evaluate(async () => {
  const td = document.querySelector('td.rowno[data-r="40"]');
  if (!td) return { error: "找不到行号" };
  const t0 = performance.now();
  td.dispatchEvent(new MouseEvent("click", { bubbles: true, clientX: 10, clientY: 10 }));
  await new Promise((r) => requestAnimationFrame(r));
  await new Promise((r) => requestAnimationFrame(r));
  return { ms: Math.round(performance.now() - t0) };
});

// ---------- ⑤ sticky 列是否仍锁定（横向滚动后 A 列应贴在行号右侧） ----------
const sticky = await page.evaluate(async () => {
  const wrap = document.getElementById("grid-wrap");
  wrap.scrollLeft = 900;
  await new Promise((r) => requestAnimationFrame(r));
  await new Promise((r) => requestAnimationFrame(r));
  const base = wrap.getBoundingClientRect().left;
  const tdA = document.querySelector("#main-grid tbody td.sticky-col-a");
  const thA = document.querySelector("#main-grid thead th.sticky-col-a");
  const out = {
    tdA: tdA ? Math.round(tdA.getBoundingClientRect().left - base) : null,
    thA: thA ? Math.round(thA.getBoundingClientRect().left - base) : null,
  };
  wrap.scrollLeft = 0;
  return out;
});

console.log(`行数 ${ROWS} · 浏览器 ${path.basename(exe)}（${HEADFUL ? "有头/真实 GPU" : "headless"}）`);
console.log(`  导入总耗时        ${tImport} ms`);
console.log(`  布防框            ${armed ? "已建立（真实场景）" : "未建立"}`);
console.log(`  滚动帧耗时        中位 ${scroll.median}ms · p90 ${scroll.p90}ms · 最大 ${scroll.max}ms · >33ms ${scroll.over33} 帧 / >50ms ${scroll.over50} 帧（共 ${scroll.frames} 帧）`);
console.log(`  编辑延迟          ${edit.error || `${edit.ms} ms（列值 ${edit.before} → ${edit.after}，更新${edit.changed ? "成功" : "超时"}）`}`);
console.log(`  点行号重渲染      ${rowClick.error || rowClick.ms + " ms"}`);
console.log(`  sticky 列（横滚后）A 列 ${sticky.tdA}px / 表头 ${sticky.thA}px（应为 42px 左右）`);
if (errors.length) console.log(`  页面报错          ${errors.slice(0, 3).join(" | ")}`);

await browser.close();
server.close();
