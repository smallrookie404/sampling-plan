// 浏览器冒烟：数据上传「项目组成员」人员选项框（2026-10-20 用户需求）。
//
// 需求：项目组成员下方的人员列表**不要一直显示**，只有搜索框里输入内容时才出现选项框。
//
// 覆盖：
//   A. 默认收起：未输入时 #memberList 不显示（display:none），但人员已渲染
//   B. 输入关键字 → 选项框出现，且只显示匹配项
//   C. 无匹配 → 选项框仍在，显示「无匹配人员」
//   D. 清空搜索 → 收起
//   E. Esc → 清空并收起；点选项框外部 → 收起；点选项框内部（勾选）→ 保持展开
//   F. 选项框不遮挡「确认成员」按钮（按钮可点，且勾选结果能确认）
//
// 用法：node tests/memberpicker.smoke.mjs
import { chromium } from "playwright-core";
import fs from "node:fs";
import path from "node:path";
import { startServer } from "../server.mjs";
import { findBrowser } from "./_browser.mjs";

const CHROME = findBrowser();
if (!CHROME) { console.error("未找到 Chrome/Edge，跳过该冒烟测试"); process.exit(0); }

const TEST_DIR = path.resolve("data/_test-memberpicker");
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

const page = await browser.newPage({ viewport: { width: 1200, height: 1200 } });
page.on("dialog", (d) => d.accept());
const pageErrors = [];
page.on("pageerror", (e) => pageErrors.push(e.message));
await page.addInitScript(() => {
  sessionStorage.setItem("xcdc_session_v1", JSON.stringify({
    token: "test-token", orgId: "1",
    userInfo: { userName: "王飞", userCode: "btc-wf", organization: { organizationId: "1" } },
  }));
  const origFetch = window.fetch.bind(window);
  const json = (obj) => Promise.resolve(new Response(JSON.stringify(obj), {
    status: 200, headers: { "Content-Type": "application/json" },
  }));
  window.fetch = function (input, init) {
    const url = typeof input === "string" ? input : (input && input.url) || "";
    if (url.indexOf("/api/users") >= 0) {
      return json({ code: "200", body: { content: [
        { id: 1, userName: "王飞", userCode: "btc-wf" },
        { id: 2, userName: "金云龚", userCode: "btc-jyg" },
        { id: 3, userName: "王慈媛", userCode: "btc-wcy" },
        { id: 4, userName: "周海江", userCode: "btc-zhj" },
      ] } });
    }
    if (url.indexOf("223.93.144.122") >= 0) {
      if (url.indexOf("/auth/info") >= 0) {
        return json({ code: "200", user: { id: 49, userName: "王飞", userCode: "btc-wf", organization: { organizationId: 1 } } });
      }
      return json({ code: "200", body: { records: [], content: [] } });
    }
    return origFetch(input, init);
  };
});

await page.goto(APP_URL, { waitUntil: "load" });
await page.waitForFunction(() => window.SamplingUpload && window.SamplingUpload.show, null, { timeout: 20000 });
await page.evaluate(() => window.SamplingUpload.show());
await page.waitForFunction(() => document.querySelectorAll("#memberList .member-item").length > 0, null, { timeout: 20000 });

let pass = 0, fail = 0;
const check = (cond, name, extra) => {
  if (cond) { pass++; console.log(`  ✔ ${name}`); }
  else { fail++; console.error(`  ✘ ${name}${extra ? " :: " + extra : ""}`); }
};
const state = () => page.evaluate(() => {
  const l = document.getElementById("memberList");
  const items = Array.from(l.querySelectorAll(".member-item"));
  return {
    open: l.classList.contains("open"),
    display: getComputedStyle(l).display,
    total: items.length,
    visible: items.filter((x) => !x.classList.contains("hidden")).map((x) => x.textContent.trim()),
    empty: !!l.querySelector(".member-empty"),
  };
});
const type = (v) => page.evaluate((val) => {
  const s = document.getElementById("memberSearch");
  s.value = val;
  s.dispatchEvent(new Event("input", { bubbles: true }));
}, v);

// ---- A. 默认收起 ----
const a1 = await state();
check(a1.total === 4, "A1 人员已渲染（4 人）", JSON.stringify(a1));
check(a1.open === false && a1.display === "none", "A2 未搜索时人员列表不显示", JSON.stringify(a1));

// ---- A2b. 布局：「项目组成员」+「确认成员」在搜索框上方 ----
const layout = await page.evaluate(() => {
  const r = (sel) => { const e = document.querySelector(sel); return e ? e.getBoundingClientRect() : null; };
  const head = r(".member-head"), btn = r("#btnMemberConfirm"), input = r("#memberSearch"), label = r(".member-head label");
  const btnEl = document.getElementById("btnMemberConfirm");
  return {
    btnAboveInput: !!btn && !!input && btn.bottom <= input.top + 1,
    sameRow: !!head && !!btn && head.top <= btn.top && btn.bottom <= head.bottom + 1,
    labelLeftOfBtn: !!label && !!btn && label.right < btn.left,
    btnInHead: !!btnEl && btnEl.parentElement === document.querySelector(".member-head"),
  };
});
check(layout.btnAboveInput, "A5 「确认成员」按钮在搜索框上方", JSON.stringify(layout));
check(layout.sameRow && layout.labelLeftOfBtn, "A6 「项目组成员」与「确认成员」同一行（左标题 / 右按钮）", JSON.stringify(layout));
check(layout.btnInHead, "A7 按钮位于标题行容器内", JSON.stringify(layout));

// ---- A3/A4. 点搜索框即展开（2026-10-20 追加需求）----
await page.click("#memberSearch");
const a3 = await state();
check(a3.open === true && a3.display !== "none", "A3 点搜索框 → 选项框出现", JSON.stringify(a3));
check(a3.visible.length === 4, "A4 未输入关键字时显示全部人员（4 人）", JSON.stringify(a3.visible));

// ---- B. 输入即过滤，且只显示匹配项 ----
await type("王");
const b1 = await state();
check(b1.open === true && b1.display !== "none", "B1 输入关键字 → 选项框保持展开", JSON.stringify(b1));
check(b1.visible.length === 2 && b1.visible.join(",").indexOf("金云龚") < 0,
  "B2 只显示匹配项（王飞/王慈媛）", JSON.stringify(b1.visible));

await type("btc-jyg");
const b3 = await state();
check(b3.visible.length === 1 && b3.visible[0].indexOf("金云龚") >= 0, "B3 支持按工号搜索", JSON.stringify(b3.visible));

// ---- C. 无匹配 ----
await type("zzz");
const c1 = await state();
check(c1.open === true && c1.visible.length === 0 && c1.empty === true, "C1 无匹配 → 选项框仍显示「无匹配人员」", JSON.stringify(c1));

// ---- D. 清空搜索（仍在框内）→ 回到全部；点外部才收起 ----
await type("");
const d1 = await state();
check(d1.open === true && d1.visible.length === 4, "D1 清空搜索（仍聚焦）→ 保持展开并显示全部人员", JSON.stringify(d1));
check(d1.empty === false, "D2 清空后移除「无匹配人员」提示", JSON.stringify(d1));
await page.click("#projectInfo", { force: true });
const d3 = await state();
check(d3.open === false && d3.display === "none", "D3 点搜索框外部 → 收起", JSON.stringify(d3));

// ---- E. Esc / 点外部 / 点内部 ----
await page.click("#memberSearch");
await page.evaluate(() => document.getElementById("memberSearch").dispatchEvent(
  new KeyboardEvent("keydown", { key: "Escape", bubbles: true })));
const e1 = await page.evaluate(() => ({
  value: document.getElementById("memberSearch").value,
  open: document.getElementById("memberList").classList.contains("open"),
}));
check(e1.value === "" && e1.open === false, "E1 Esc → 清空搜索并收起", JSON.stringify(e1));

await type("王");
await page.click("#projectInfo", { force: true });
const e2 = await state();
check(e2.open === false, "E2 点选项框外部 → 收起", JSON.stringify({ open: e2.open }));

await type("王");
await page.click("#memberList .member-item:not(.hidden) input");
const e3 = await page.evaluate(() => ({
  open: document.getElementById("memberList").classList.contains("open"),
  checked: Array.from(document.querySelectorAll("#memberList input:checked")).map((x) => x.value),
}));
check(e3.open === true && e3.checked.length === 1, "E3 点选项框内部勾选 → 保持展开且已勾选", JSON.stringify(e3));

// ---- F. 不遮挡「确认成员」，勾选可确认 ----
const f1 = await page.evaluate(() => {
  const btn = document.getElementById("btnMemberConfirm");
  const r = btn.getBoundingClientRect();
  const hit = document.elementFromPoint(r.left + r.width / 2, r.top + r.height / 2);
  return { hitIsBtn: !!(hit && (hit === btn || btn.contains(hit))), visible: r.width > 0 && r.height > 0 };
});
check(f1.visible && f1.hitIsBtn, "F1 选项框展开时「确认成员」按钮未被遮挡", JSON.stringify(f1));
await page.click("#btnMemberConfirm");
await page.waitForFunction(() => document.getElementById("memberSearch").value === "", null, { timeout: 10000 });
const f2 = await page.evaluate(() => ({
  open: document.getElementById("memberList").classList.contains("open"),
  tags: document.getElementById("memberTags").textContent.trim(),
  count: document.getElementById("memberCount").textContent,
}));
check(f2.open === false && f2.tags.indexOf("王飞") >= 0, "F2 确认成员后选项框收起、成员进入标签", JSON.stringify(f2));
check(f2.count.indexOf("1") >= 0, "F3 已勾选人数同步", f2.count);

check(pageErrors.length === 0, "H1 页面无未捕获异常", pageErrors.join(" | "));

console.log(`\nmemberpicker.smoke: ${pass} 通过, ${fail} 失败`);
try { await browser.close(); } catch {}
server.close();
try { fs.rmSync(TEST_DIR, { recursive: true, force: true }); } catch {}
process.exit(fail > 0 ? 1 : 0);
