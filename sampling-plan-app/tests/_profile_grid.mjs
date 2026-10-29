/*
 * 主表渲染热点剖析（非正式测试）。
 * 用 CDP Profiler 抓「点行号触发整窗重渲染」与「滚动」期间的 JS 采样，
 * 按自耗时排序输出热点函数，用于定位滚动/编辑卡顿的真正来源。
 *
 * 用法：node tests/_profile_grid.mjs [行数] [场景: rowclick|scroll]
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
const SCENE = process.argv[3] || "rowclick";
const PORT = 8124;

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
  r[4] = "2"; r[5] = "8"; r[6] = "5"; r[7] = String(1 + (i % 3)); r[8] = "1";
  r[18] = "三班两运转";
  rows.push(r);
}
const buf = await X.writeWorkbook([{ name: "测点布局情况调查", rows }]);

const server = await startServer({ port: PORT, dataFile: path.join(os.tmpdir(), `prof-records-${process.pid}.json`) });
const exe = findBrowser();
const browser = await chromium.launch({ executablePath: exe, headless: true, args: ["--disable-gpu"] });
const page = await browser.newPage({ viewport: { width: 1600, height: 900 } });
page.on("dialog", (d) => d.accept().catch(() => {}));
await page.goto(`http://127.0.0.1:${PORT}/index.html`, { waitUntil: "domcontentloaded" });
await page.waitForSelector("#grid-body tr[data-r]");
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
await page.evaluate(() => { document.getElementById("grid-wrap").scrollTop = 0; });
await page.waitForTimeout(500);

const client = await page.context().newCDPSession(page);
await client.send("Profiler.enable");
await client.send("Profiler.setSamplingInterval", { interval: 100 });
await client.send("Profiler.start");

if (SCENE === "rowclick") {
  await page.evaluate(async () => {
    for (let k = 0; k < 10; k++) {
      const td = document.querySelector(`td.rowno[data-r="${20 + k}"]`);
      if (td) td.dispatchEvent(new MouseEvent("click", { bubbles: true, clientX: 10, clientY: 10 }));
      await new Promise((r) => requestAnimationFrame(r));
    }
  });
} else {
  await page.evaluate(async () => {
    const wrap = document.getElementById("grid-wrap");
    for (let i = 0; i < 120; i++) {
      wrap.scrollTop = wrap.scrollTop + 90;
      await new Promise((r) => requestAnimationFrame(r));
    }
  });
}

const { profile } = await client.send("Profiler.stop");

// 统计每个函数的自耗时（samples 次数 × 采样间隔）
const byId = new Map(profile.nodes.map((n) => [n.id, n]));
const self = new Map();
for (const id of profile.samples) self.set(id, (self.get(id) || 0) + 1);
const total = profile.samples.length;
const rowsOut = [];
for (const [id, cnt] of self) {
  const n = byId.get(id);
  if (!n) continue;
  const cf = n.callFrame;
  const name = cf.functionName || "(anonymous)";
  const url = (cf.url || "").split("/").pop();
  rowsOut.push({ name, url, line: cf.lineNumber + 1, cnt, pct: (cnt / total) * 100 });
}
rowsOut.sort((a, b) => b.cnt - a.cnt);
console.log(`行数 ${ROWS} · 场景 ${SCENE} · 采样 ${total} 次`);
for (const r of rowsOut.slice(0, 18)) {
  console.log(`  ${r.pct.toFixed(1).padStart(5)}%  ${(r.name + " ").padEnd(26)} ${r.url}:${r.line}`);
}

await browser.close();
server.close();
