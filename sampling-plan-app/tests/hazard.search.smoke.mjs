// 浏览器冒烟：危害因素库「搜索后」的单元格编辑（回归 2026-10-21 的 bug）。
//
// 问题：搜索/列筛选后双击单元格进不了编辑态。
// 根因：窗口化渲染的行位置按「显示顺序」排，而 hazSetCur/hazRevealCell 用**绝对下标**算滚动位置，
//       筛选后算到很远处 → 误判「当前格不可见」→ 无谓 renderHazard() 把 DOM 换掉，
//       双击事件里抓到的 td 变成游离节点 → input.focus() 失效，编辑态没进去。
// 修复：新增 hazDisplayPos()，hazSetCur/hazRevealCell 都按显示位置算；双击处重取 td。
//
// 覆盖：
//   A. 搜索后双击 → 进入编辑（可键入、Esc 还原）
//   B. 搜索后 F2 → 进入编辑
//   C. 搜索后方向键按「显示顺序」移动（不是绝对下标 +1）
//   D. 搜索后 Delete 清空当前格
//   E. 清空搜索后双击仍可编辑（不回归）
//
// 用法：node tests/hazard.search.smoke.mjs
import { chromium } from "playwright-core";
import fs from "node:fs";
import path from "node:path";
import { startServer } from "../server.mjs";
import { findBrowser } from "./_browser.mjs";

const CHROME = findBrowser();
if (!CHROME) { console.error("未找到 Chrome/Edge，跳过该冒烟测试"); process.exit(0); }

const TEST_DIR = path.resolve("data/_test-hazard-search");
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
page.on("dialog", (d) => d.accept().catch(() => {}));
const pageErrors = [];
page.on("pageerror", (e) => pageErrors.push(e.message));
await page.addInitScript(() => {
  sessionStorage.setItem("xcdc_session_v1", JSON.stringify({
    token: "test-token", orgId: "1",
    userInfo: { userName: "测试账号", userCode: "TEST", organization: { organizationId: "1" } },
  }));
  // 本测试会改单元格（键入/Delete）→ 拦掉参考库的**写**请求，避免把改动落到 data/library.json
  // （server.mjs 的 /api/library 写入固定写到 appDir/data/library.json，与测试数据目录无关）
  const origFetch = window.fetch.bind(window);
  window.fetch = function (input, init) {
    const url = typeof input === "string" ? input : (input && input.url) || "";
    const method = String((init && init.method) || "GET").toUpperCase();
    if (url.indexOf("/api/library") >= 0 && method !== "GET") {
      return Promise.resolve(new Response(JSON.stringify({ ok: true }), {
        status: 200, headers: { "Content-Type": "application/json" },
      }));
    }
    return origFetch(input, init);
  };
});

await page.goto(APP_URL, { waitUntil: "domcontentloaded" });
await page.waitForTimeout(600);
await page.evaluate(() => document.querySelector("#xcdc-login")?.remove());
await page.click('.tab[data-tab="hazard"]');
await page.waitForSelector("#hazard-body tr[data-h]", { timeout: 8000 });
await page.evaluate(() => { const w = document.querySelector(".hazard-wrap"); if (w) w.scrollTop = 0; });
await page.waitForTimeout(200);

let pass = 0, fail = 0;
const check = (cond, name, extra) => {
  if (cond) { pass++; console.log(`  ✔ ${name}`); }
  else { fail++; console.error(`  ✘ ${name}${extra ? " :: " + extra : ""}`); }
};

const visibleRows = () => page.evaluate(() =>
  Array.from(document.querySelectorAll("#hazard-body tr[data-h]")).map((tr) => Number(tr.dataset.h)));
const cellCenter = (r, c) => page.evaluate(([rr, cc]) => {
  const tr = document.querySelector(`#hazard-body tr[data-h="${rr}"]`);
  const td = tr && tr.querySelector(`td[data-hc="${cc}"]`);
  const inp = td && td.querySelector("input");
  if (!inp) return null;
  const b = inp.getBoundingClientRect();
  return { x: b.left + b.width / 2, y: b.top + b.height / 2, val: inp.value };
}, [r, c]);
const cellVal = (r, c) => page.evaluate(([rr, cc]) => {
  const tr = document.querySelector(`#hazard-body tr[data-h="${rr}"]`);
  const inp = tr && tr.querySelector(`td[data-hc="${cc}"] input`);
  return inp ? inp.value : null;
}, [r, c]);
const active = () => page.evaluate(() => ({
  tag: document.activeElement && document.activeElement.tagName,
  val: document.activeElement && document.activeElement.value,
  inHazard: !!(document.activeElement && document.getElementById("hazard-body").contains(document.activeElement)),
  cur: (() => {
    const td = document.querySelector("#hazard-body td.cur");
    return td ? Number(td.closest("tr").dataset.h) + "/" + Number(td.dataset.hc) : null;
  })(),
}));

// ---- 搜索 ----
await page.fill("#hazard-search", "甲");
await page.waitForTimeout(450);
const rows = await visibleRows();
const total = await page.evaluate(() => Number((document.getElementById("hazard-status").textContent.match(/共 (\d+) 条/) || [])[1] || 0));
check(rows.length > 0 && rows.length < total, "S1 搜索后只显示匹配行（少于总数）", JSON.stringify({ shown: rows.length, total }));
const r0 = rows[0];
const name0 = await cellVal(r0, 1);
check(!!name0 && name0.indexOf("甲") >= 0, "S2 首行确实命中关键字", JSON.stringify({ r0, name0 }));

// ---- A. 双击进入编辑 ----
const a0 = await cellCenter(r0, 1);
await page.mouse.dblclick(a0.x, a0.y);
const a1 = await active();
check(a1.tag === "INPUT" && a1.inHazard, "A1 搜索后双击 → 单元格进入编辑（input 获得焦点）", JSON.stringify(a1));
await page.keyboard.type("测");
const a2 = await cellVal(r0, 1);
check(a2 === name0 + "测", "A2 编辑态可直接键入（追加到原值后）", JSON.stringify({ before: name0, after: a2 }));
await page.keyboard.press("Escape");
const a3 = await cellVal(r0, 1);
check(a3 === name0, "A3 Esc 还原原值", JSON.stringify({ a3, name0 }));

// ---- B. F2 进入编辑 ----
const b0 = await cellCenter(r0, 2);
await page.mouse.click(b0.x, b0.y);
await page.keyboard.press("F2");
const b1 = await active();
check(b1.tag === "INPUT" && b1.inHazard && b1.cur === r0 + "/2", "B1 搜索后 F2 → 进入编辑（当前格不变）", JSON.stringify(b1));
await page.keyboard.press("Escape");

// ---- C. 方向键按显示顺序移动 ----
const c0 = await cellCenter(r0, 1);
await page.mouse.click(c0.x, c0.y);
await page.keyboard.press("ArrowDown");
const c1 = await active();
const expectNext = rows[1];
check(c1.cur === expectNext + "/1", "C1 方向键下移按「显示顺序」到下一筛选行", JSON.stringify({ cur: c1.cur, expect: expectNext + "/1" }));
check(expectNext !== r0 + 1, "C2 该用例本身有区分度（显示下一行 ≠ 绝对下标+1）", JSON.stringify({ r0, expectNext }));

// ---- D. Delete 清空当前格 ----
const d0 = await cellCenter(r0, 3);
await page.mouse.click(d0.x, d0.y);
await page.keyboard.press("Delete");
const d1 = await cellVal(r0, 3);
check(d1 === "", "D1 搜索后 Delete 清空当前格", JSON.stringify({ d1 }));

// ---- E. 清空搜索后仍可双击编辑 ----
await page.fill("#hazard-search", "");
await page.waitForTimeout(450);
const rows2 = await visibleRows();
const e0 = await cellCenter(rows2[0], 1);
await page.mouse.dblclick(e0.x, e0.y);
const e1 = await active();
check(e1.tag === "INPUT" && e1.inHazard, "E1 清空搜索后双击仍进入编辑", JSON.stringify(e1));
await page.keyboard.press("Escape");

check(pageErrors.length === 0, "H1 页面无未捕获异常", pageErrors.join(" | "));

console.log(`\nhazard.search.smoke: ${pass} 通过, ${fail} 失败`);
try { await browser.close(); } catch {}
server.close();
try { fs.rmSync(TEST_DIR, { recursive: true, force: true }); } catch {}
process.exit(fail > 0 ? 1 : 0);
