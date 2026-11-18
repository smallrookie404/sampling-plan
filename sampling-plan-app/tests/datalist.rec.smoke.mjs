// 浏览器冒烟：录入区「接害因素」下拉只取危害因素库的「识别」列(rec)。
//
// 需求（2026-11-18 用户原话）：
//   「测点布局调查的录入区中的接害因素下拉内容，只取危害因素库中的识别列内容」
//
// 背景：rebuildDatalist() 原先把 h.rec 与 h.name（识别 + 系统名称）一起塞进 #hazard-dl，
//   但计算侧「接害因素」是按 rec 逐字匹配（logic.js findRec ↔ VLOOKUP(接害因素,A:B,2,FALSE)），
//   选中系统名称（如 安妥 / 1,3-丁二烯）反而匹配不到、AN 为空。现在候选只有 rec。
//
// 覆盖：
//   A. 候选集合 == 参考库 rec 去重后的集合（逐项一致、无重复、无空值）
//   B. 「只在系统名称里出现」的值不出现在候选里（安妥 / 1,3-丁二烯 这类）
//   C. D 列（接害因素）输入框确实挂了 list="hazard-dl"，其它联想列各用各的
//   D. 编辑危害因素库的「识别」列 → 候选跟着变（新值进、旧值出），说明下拉来自库而非硬编码
//   E. 页面无未捕获异常
//
// 用法：node tests/datalist.rec.smoke.mjs
import { chromium } from "playwright-core";
import fs from "node:fs";
import path from "node:path";
import { startServer } from "../server.mjs";
import { findBrowser } from "./_browser.mjs";

const CHROME = findBrowser();
if (!CHROME) { console.error("未找到 Chrome/Edge，跳过该冒烟测试"); process.exit(0); }

const APP_DIR = process.cwd();
const LIB_PATH = path.join(APP_DIR, "data", "library.json");
const lib = JSON.parse(fs.readFileSync(LIB_PATH, "utf8"));
const all = lib.hazardFactors || [];
const recSet = new Set(all.map((h) => h.rec).filter(Boolean));
const nameOnly = [...new Set(all.map((h) => h.name).filter((n) => n && !recSet.has(n)))];

const TEST_DIR = path.resolve("data/_test-datalist-rec");
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
page.on("dialog", (d) => d.accept().catch(() => {}));
const pageErrors = [];
page.on("pageerror", (e) => pageErrors.push(e.message));
await page.addInitScript(() => {
  try { sessionStorage.setItem("xcdc_guest_v1", "1"); } catch (e) {}
  // 本用例会改「识别」列 → 拦掉参考库的写请求，避免把改动落到 data/library.json
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

let pass = 0, fail = 0;
function check(cond, name, detail) {
  if (cond) { pass++; console.log("  ✓ " + name); }
  else { fail++; console.log("  ✗ " + name + (detail === undefined ? "" : "  → " + detail)); }
}

await page.goto(APP_URL, { waitUntil: "load", timeout: 20000 });
await page.waitForFunction(() => document.querySelectorAll("#hazard-dl option").length > 100, null, { timeout: 20000 });
await page.waitForTimeout(300);

const dlValues = () => page.evaluate(() =>
  Array.from(document.querySelectorAll("#hazard-dl option")).map((o) => o.value));

console.log("[A] 候选集合 = 参考库「识别」列");
const vals = await dlValues();
const vset = new Set(vals);
check(vals.length === vset.size, "A1 候选无重复", `count=${vals.length} unique=${vset.size}`);
check(vset.size === recSet.size, "A2 候选条数 = 识别列去重条数", `dom=${vset.size} lib=${recSet.size}`);
const missing = [...recSet].filter((v) => !vset.has(v));
const extra = [...vset].filter((v) => !recSet.has(v));
check(missing.length === 0, "A3 识别列的值全部在候选里", missing.slice(0, 5).join(" / "));
check(extra.length === 0, "A4 候选里没有识别列之外的值", extra.slice(0, 5).join(" / "));
check(!vals.some((v) => !v || !v.trim()), "A5 没有空候选");

console.log("[B] 只在「系统名称」里出现的值不该进候选");
check(nameOnly.length > 0, "B0 库里确实存在 rec≠name 的条目（用例有效）", `nameOnly=${nameOnly.length}`);
const leaked = nameOnly.filter((n) => vset.has(n));
check(leaked.length === 0, "B1 系统名称独有值没有混进候选", leaked.slice(0, 5).join(" / "));

console.log("[C] 只有接害因素列挂这个下拉");
const lists = await page.evaluate(() => {
  // 表格是虚拟滚动的，首个子节点是 row-spacer，取第一条真实数据行
  const tr = document.querySelector('#grid-body tr[data-r="0"]');
  const out = {};
  if (tr) tr.querySelectorAll("input[list]").forEach((inp) => { out[inp.dataset.c] = inp.getAttribute("list"); });
  return out;
});
check(lists.D === "hazard-dl", "C1 D 列（接害因素）用 hazard-dl", JSON.stringify(lists));
check(lists.L === "zuoye-dl" && lists.N === "caiyang-dl" && lists.O === "gangwei-dl", "C2 其它联想列各用各的下拉", JSON.stringify(lists));

console.log("[D] 改「识别」列 → 候选同步");
const oldRec = vals[0];
await page.evaluate(() => document.querySelector('.tab[data-tab="hazard"]').click());
await page.waitForTimeout(500);
const box = await page.evaluate(() => {
  const inp = document.querySelector('#hazard-body tr[data-h="0"] td[data-hc="0"] input');
  if (!inp) return null;
  const b = inp.getBoundingClientRect();
  return { x: b.left + b.width / 2, y: b.top + b.height / 2, val: inp.value };
});
check(!!box && box.val === oldRec, "D0 危害因素库首行识别列 = 候选首项（同一来源）", JSON.stringify({ box, oldRec }));
await page.mouse.click(box.x, box.y);
await page.mouse.dblclick(box.x, box.y);
await page.keyboard.press("Control+a");
await page.keyboard.type("ZZ测试因素");
await page.keyboard.press("Tab");
await page.waitForTimeout(800);
const vals2 = await dlValues();
check(vals2.includes("ZZ测试因素"), "D1 新识别值进了候选", `has=${vals2.includes("ZZ测试因素")}`);
check(!vals2.includes(oldRec), "D2 旧识别值从候选消失", `still=${vals2.includes(oldRec)}`);

console.log("[E] 异常");
check(pageErrors.length === 0, "E1 页面无未捕获异常", pageErrors.join(" | "));

console.log(`\ndatalist.rec.smoke: ${pass} 通过, ${fail} 失败`);
try { await browser.close(); } catch {}
server.close();
try { fs.rmSync(TEST_DIR, { recursive: true, force: true }); } catch {}
process.exit(fail > 0 ? 1 : 0);
