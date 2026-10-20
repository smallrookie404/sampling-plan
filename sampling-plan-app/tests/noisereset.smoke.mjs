// 浏览器冒烟：切换项目时清空「上一个报告」的噪声筛选结果。
//
// 需求（2026-10-20）：噪声数据结果页签里的「非噪声车间岗位 / 噪声车间岗位 / 噪声超标车间岗位」
// 三块筛选结果（以及噪声表格、状态行），在导出/查看其他项目的数据时不能残留上一个报告的内容。
//
// 覆盖：
//   A. resetNoiseState() 直接调用：三块筛选结果 + 表格 + 状态行全部清空、页签收起
//   B. 选中新项目（点搜索建议）→ 自动清空上一份噪声结果
//   C. 重新搜索（输入关键字触发搜索）→ 已选项目作废，噪声结果一并清空
//   D. 清空后「非噪声岗位同步」不再拿旧数据（提示先加载）
//
// 用法：node tests/noisereset.smoke.mjs
import { chromium } from "playwright-core";
import fs from "node:fs";
import path from "node:path";
import { startServer } from "../server.mjs";
import { findBrowser } from "./_browser.mjs";

const CHROME = findBrowser();
if (!CHROME) { console.error("未找到 Chrome/Edge，跳过该冒烟测试"); process.exit(0); }

const TEST_DIR = path.resolve("data/_test-noisereset");
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

const page = await browser.newPage({ viewport: { width: 1400, height: 900 } });
const dialogs = [];
page.on("dialog", (d) => { dialogs.push(d.message()); d.accept(); });
const pageErrors = [];
page.on("pageerror", (e) => pageErrors.push(e.message));
await page.addInitScript(() => {
  sessionStorage.setItem("xcdc_session_v1", JSON.stringify({
    token: "test-token", orgId: "1",
    userInfo: { userName: "测试账号", userCode: "TEST", organization: { organizationId: "1" } },
  }));
  const origFetch = window.fetch.bind(window);
  const json = (obj) => Promise.resolve(new Response(JSON.stringify(obj), {
    status: 200, headers: { "Content-Type": "application/json" },
  }));
  window.fetch = function (input, init) {
    const url = typeof input === "string" ? input : (input && input.url) || "";
    if (url.indexOf("223.93.144.122") >= 0) {
      if (url.indexOf("/auth/info") >= 0) {
        return json({ code: "200", user: { id: 49, userName: "测试账号", userCode: "TEST", organization: { organizationId: 1 } } });
      }
      // E 段用：让噪声数据加载卡在「加载中」（请求永不返回），验证加载期间不残留旧数据
      if (window.__hangPjbg && url.indexOf("/api/pjbgModule") >= 0) {
        return new Promise(function () {});
      }
      if (url.indexOf("/api/reportData/findList") >= 0) {
        return json({ code: "200", message: "查询成功！", body: { records: [
          { id: 6205, code: "BTC26-SZJC0959", belongInspect: 1450, belongInspectName: "甲单位" },
          { id: 6300, code: "BTC26-SZJC1009", belongInspect: 1600, belongInspectName: "乙单位" },
        ] } });
      }
      return json({ code: "200", body: { records: [], content: [] } });
    }
    return origFetch(input, init);
  };
});

await page.goto(APP_URL, { waitUntil: "load" });
await page.waitForFunction(() => window.SamplingUpload && window.SamplingUpload.resetNoiseState, null, { timeout: 20000 });
await page.evaluate(() => window.SamplingUpload.show());
await page.waitForTimeout(300);

let pass = 0, fail = 0;
const check = (cond, name, extra) => {
  if (cond) { pass++; console.log(`  ✔ ${name}`); }
  else { fail++; console.error(`  ✘ ${name}${extra ? " :: " + extra : ""}`); }
};
// 造一份「上一个报告」的噪声结果（三块筛选 + 表格 + 状态行）
const seedNoise = () => page.evaluate(() => {
  const sum = document.getElementById('noise-summary');
  sum.innerHTML = '<div class="noise-summary-grid">'
    + '<div class="noise-summary-item ok"><div class="nsi-body">旧车间、旧岗位</div></div>'
    + '<div class="noise-summary-item noise"><div class="nsi-body">旧车间、噪声岗位</div></div>'
    + '<div class="noise-summary-item over"><div class="nsi-body">旧车间、超标岗位</div></div></div>';
  sum.classList.remove('hidden');
  document.getElementById('noise-table-box').innerHTML = '<table><tr><td>旧报告数据</td></tr></table>';
  document.getElementById('noise-status').textContent = '项目 BTC26-SZJC0959 · 共 99 行';
  document.getElementById('pjbg-status').textContent = '旧报告已生成';
});
const snapNoise = () => page.evaluate(() => {
  const sum = document.getElementById('noise-summary');
  const bodies = Array.from(sum.querySelectorAll('.nsi-body')).map((x) => x.textContent.trim());
  return {
    summaryHtml: sum.innerHTML,
    summaryHidden: sum.classList.contains('hidden'),
    items: sum.querySelectorAll('.noise-summary-item').length,
    titles: Array.from(sum.querySelectorAll('.nsi-title')).map((x) => x.textContent.trim()),
    bodies: bodies,
    tableHtml: document.getElementById('noise-table-box').innerHTML,
    status: document.getElementById('noise-status').textContent,
    pjbg: document.getElementById('pjbg-status').textContent,
  };
});
// 「页签还留着、内容清空」= 三块卡片仍在，正文全部为「（无）」
const isEmptySkeleton = (s) => s.items === 3 && s.summaryHidden === false
  && s.bodies.length === 3 && s.bodies.every((b) => b === '（无）')
  && s.titles.join(',') === '非噪声车间岗位,噪声车间岗位,噪声超标车间岗位';

// ---- A. 直接调用 resetNoiseState ----
await seedNoise();
const a0 = await snapNoise();
check(a0.summaryHtml.length > 0 && !a0.summaryHidden, "A1 造出「上一个报告」的筛选结果", a0.summaryHtml.length);
await page.evaluate(() => window.SamplingUpload.resetNoiseState());
const a1 = await snapNoise();
check(isEmptySkeleton(a1), "A2 三块页签仍在、内容已清空（（无））", JSON.stringify(a1));
check(a1.tableHtml === '' && a1.status === '' && a1.pjbg === '', "A3 噪声表格 / 状态行 / 报告状态一并清空", JSON.stringify(a1));

// ---- B. 选中新项目 → 自动清空 ----
await seedNoise();
await page.evaluate(() => {
  const code = document.getElementById('projectCode');
  code.value = "0959";
  code.dispatchEvent(new Event("input", { bubbles: true }));
});
await page.waitForFunction(() => document.querySelectorAll(".suggest-item").length > 0, null, { timeout: 15000 });
await page.evaluate(() => document.querySelector(".suggest-item").click());
await page.waitForTimeout(200);
const b1 = await snapNoise();
check(isEmptySkeleton(b1), "B1 选中项目后页签仍在、内容已清空（不残留上一家）", JSON.stringify(b1));
check(b1.tableHtml === '', "B2 选中项目后噪声表格已清空", JSON.stringify(b1.tableHtml));
const b2 = await page.evaluate(() => document.getElementById('projectInfo').textContent);
check(b2.indexOf("0959") >= 0, "B3 项目仍正常选中", b2);

// ---- C. 重新搜索 → 清空 ----
await seedNoise();
await page.evaluate(() => {
  const code = document.getElementById('projectCode');
  code.value = "1009";
  code.dispatchEvent(new Event("input", { bubbles: true }));
});
await page.waitForTimeout(600);
const c1 = await snapNoise();
check(isEmptySkeleton(c1), "C1 重新搜索后页签仍在、内容已清空", JSON.stringify(c1));
check(c1.tableHtml === '', "C2 重新搜索后噪声表格已清空", JSON.stringify(c1.tableHtml));

// ---- D. 清空后「非噪声岗位同步」不再用旧数据 ----
dialogs.length = 0;
await page.evaluate(() => document.getElementById('noise-sync').click());
await page.waitForTimeout(200);
check(dialogs.some((m) => m.indexOf('请先加载噪声数据结果') >= 0),
  "D1 清空后点「非噪声岗位同步」→ 提示先加载（不用旧缓存）", JSON.stringify(dialogs));

// ---- E. 加载过程中不残留（点开页签 → 请求挂起 → 三块/表格已清空）----
// C 段重新搜索后已无选中项目，这里先重新选一个
await page.evaluate(() => {
  const code = document.getElementById('projectCode');
  code.value = "1009";
  code.dispatchEvent(new Event("input", { bubbles: true }));
});
// 等防抖搜索真正跑完（搜索一开始会清空已选项目，太早点建议会被随后的搜索清掉）
await page.waitForTimeout(900);
await page.waitForFunction(() => document.querySelectorAll(".suggest-item").length > 0, null, { timeout: 15000 });
await page.evaluate(() => document.querySelector(".suggest-item").click());
await page.waitForFunction(() => document.getElementById("projectInfo").textContent.indexOf("已选择项目") >= 0, null, { timeout: 10000 });
await page.evaluate(() => { window.__hangPjbg = true; });
await seedNoise();
const e0 = await snapNoise();
check(e0.summaryHtml.length > 0, "E0 加载前先造出上一家的三块数据", e0.summaryHtml.length);
await page.evaluate(() => document.querySelector('.tab[data-tab="noise"]').click());
await page.waitForTimeout(400);
const e1 = await snapNoise();
check(isEmptySkeleton(e1), "E1 加载中三块页签仍在、内容已清空（不残留上一家）", JSON.stringify(e1));
check(e1.tableHtml === '', "E2 加载中噪声表格已清空", JSON.stringify(e1.tableHtml));
check(e1.status.indexOf('加载中') >= 0, "E3 状态行显示「加载中…」（确实在加载中）", JSON.stringify({ status: e1.status, dialogs: dialogs }));
await page.evaluate(() => { window.__hangPjbg = false; });

check(pageErrors.length === 0, "H1 页面无未捕获异常", pageErrors.join(" | "));

console.log(`\nnoisereset.smoke: ${pass} 通过, ${fail} 失败`);
try { await browser.close(); } catch {}
server.close();
try { fs.rmSync(TEST_DIR, { recursive: true, force: true }); } catch {}
process.exit(fail > 0 ? 1 : 0);
