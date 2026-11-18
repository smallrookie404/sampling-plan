// 浏览器冒烟：横向滚动条必须贴在浏览器窗口底部（不能掉到页面最下面）。
//
// 背景（2026-11-18 用户原话）：
//   「测点布局调查中，有些浏览器的左右滚动条在最下面，没有固定在浏览器下方位置」
//   主表 / 调查表 / 危害因素库的滚动容器都是 flex:1 + min-height:0，靠 flex 收缩把底边压在窗口
//   底部。个别浏览器（360 极速/兼容模式、部分老内核）对「flex 子项按剩余高度收缩」支持不到位，
//   容器会被整表内容撑高 → 横向滚动条掉到页面最底部，要先把页面滚到底才能拖。
//   修法：app.js 的 fitActiveScrollWrap() 按视口实测剩余高度，用 max-height 兜底钳住容器。
//
// 覆盖：
//   A. 主表：容器底边 = 窗口底部 - 12px（.tab-panel 的 padding-bottom）、页面本身不纵向滚动
//   B. 兜底生效：模拟 flex 收缩失效（body 变 block、去掉 max-height）→ 容器被内容撑高；
//      再触发一次重算 → 又贴回窗口底部
//   C. 切换页签（调查表 / 危害因素库 / 测点布局调查）后，当前页签的容器同样贴底
//   D. 视口缩小 / 放大后重新贴底（1280x560、1600x900）
//   E. 全屏降级模式（body.fullscreen-mode）下容器同样贴底
//   F. 页面无未捕获异常
//
// 用法：node tests/layout.hscroll.smoke.mjs
import { chromium } from "playwright-core";
import fs from "node:fs";
import path from "node:path";
import { startServer } from "../server.mjs";
import { findBrowser } from "./_browser.mjs";

const CHROME = findBrowser();
if (!CHROME) { console.error("未找到 Chrome/Edge，跳过该冒烟测试"); process.exit(0); }

const TEST_DIR = path.resolve("data/_test-layout-hscroll");
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

const page = await browser.newPage({ viewport: { width: 1366, height: 768 } });
page.on("dialog", (d) => d.accept().catch(() => {}));
const pageErrors = [];
page.on("pageerror", (e) => pageErrors.push(e.message));
// 以访客身份进入，跳过登录遮罩（本用例只验证布局）
await page.addInitScript(() => { try { sessionStorage.setItem("xcdc_guest_v1", "1"); } catch (e) {} });

let pass = 0, fail = 0;
function check(cond, name, detail) {
  if (cond) { pass++; console.log("  ✓ " + name); }
  else { fail++; console.log("  ✗ " + name + (detail === undefined ? "" : "  → " + detail)); }
}

// 量一次当前页签的滚动容器：底边、高度、页面是否纵向滚动、是否有横向滚动条
const SNAP = `(() => {
  const panel = document.querySelector('.tab-panel.active');
  const wrap = panel && panel.querySelector('.grid-wrap, .hazard-wrap');
  if (!wrap) return null;
  const r = wrap.getBoundingClientRect();
  return {
    panel: panel.id,
    innerH: window.innerHeight,
    wrapBottom: Math.round(r.bottom),
    wrapH: Math.round(r.height),
    maxHeight: wrap.style.maxHeight || '',
    docScrollH: document.documentElement.scrollHeight,
    docClientH: document.documentElement.clientHeight,
    hasHScroll: wrap.scrollWidth > wrap.clientWidth + 1,
    rows: document.querySelectorAll('#grid-body tr').length,
  };
})()`;
const snap = () => page.evaluate(SNAP);
// 触发一次重算（等价于窗口尺寸变化）
const refit = async () => {
  await page.evaluate(() => window.dispatchEvent(new Event("resize")));
  await page.waitForTimeout(220);
};

await page.goto(APP_URL, { waitUntil: "load", timeout: 20000 });
await page.waitForFunction(() => document.querySelectorAll("#grid-body tr").length > 0, null, { timeout: 20000 });
await page.waitForTimeout(500);

console.log("[A] 主表默认视口");
const a = await snap();
check(!!a && a.panel === "tab-main", "A1 当前在「测点布局调查」页签", JSON.stringify(a));
check(a && a.wrapBottom <= a.innerH, "A2 滚动容器底边在窗口内（滚动条不会掉到页面外）", JSON.stringify({ bottom: a && a.wrapBottom, innerH: a && a.innerH }));
check(a && a.wrapBottom === a.innerH - 12, "A3 容器底边 = 窗口底部 - 12px（正好贴住窗口下方）", JSON.stringify({ bottom: a && a.wrapBottom, expect: a && a.innerH - 12 }));
check(a && a.docScrollH <= a.docClientH, "A4 页面本身不纵向滚动（整屏只有表格内部滚动）", JSON.stringify({ sh: a && a.docScrollH, ch: a && a.docClientH }));
check(a && a.hasHScroll, "A5 主表确实存在横向滚动条（表格比窗口宽）", JSON.stringify(a));

console.log("[B] 模拟 flex 收缩失效（老内核）时的兜底");
// 去掉钳制 + body 变 block：正是那类浏览器下的错误形态
await page.evaluate(() => {
  const w = document.querySelector(".tab-panel.active .grid-wrap");
  w.style.maxHeight = "";
  delete w.dataset.fitMax;
  document.body.style.display = "block";
});
await page.waitForTimeout(300);
const b1 = await snap();
check(b1 && b1.wrapH > b1.innerH, "B1 去掉钳制后容器被整表内容撑高（复现「滚动条在页面最下面」）", JSON.stringify({ h: b1 && b1.wrapH, innerH: b1 && b1.innerH }));
await refit();
const b2 = await snap();
check(b2 && b2.wrapBottom <= b2.innerH && b2.maxHeight !== "", "B2 触发一次重算后容器又被钳回窗口底部", JSON.stringify(b2));
await page.evaluate(() => { document.body.style.display = ""; });
await refit();

console.log("[C] 切换页签");
for (const tab of ["survey", "hazard", "main"]) {
  await page.evaluate((t) => document.querySelector(`.tab[data-tab="${t}"]`).click(), tab);
  await page.waitForTimeout(450);
  const c = await snap();
  check(c && c.wrapBottom <= c.innerH, `C ${tab} 页签容器底边在窗口内`, JSON.stringify(c));
  check(c && c.wrapBottom === c.innerH - 12, `C ${tab} 页签容器正好贴住窗口底部`, JSON.stringify({ bottom: c && c.wrapBottom, expect: c && c.innerH - 12 }));
}

console.log("[D] 视口缩放");
for (const vp of [{ width: 1280, height: 560 }, { width: 1600, height: 900 }, { width: 1366, height: 768 }]) {
  await page.setViewportSize(vp);
  await page.waitForTimeout(450);
  const d = await snap();
  check(d && d.wrapBottom <= d.innerH && d.wrapBottom === d.innerH - 12, `D ${vp.width}x${vp.height} 贴住窗口底部`, JSON.stringify({ bottom: d && d.wrapBottom, expect: d && d.innerH - 12 }));
}

console.log("[E] 全屏降级模式");
await page.evaluate(() => { document.body.classList.add("fullscreen-mode"); window.dispatchEvent(new Event("resize")); });
await page.waitForTimeout(400);
const e = await snap();
check(e && e.wrapBottom <= e.innerH && e.wrapBottom === e.innerH - 12, "E1 全屏模式下容器贴住窗口底部", JSON.stringify({ bottom: e && e.wrapBottom, expect: e && e.innerH - 12 }));
await page.evaluate(() => { document.body.classList.remove("fullscreen-mode"); window.dispatchEvent(new Event("resize")); });
await page.waitForTimeout(300);

console.log("[F] 异常");
check(pageErrors.length === 0, "F1 页面无未捕获异常", pageErrors.join(" | "));

console.log(`\nlayout.hscroll.smoke: ${pass} 通过, ${fail} 失败`);
try { await browser.close(); } catch {}
server.close();
try { fs.rmSync(TEST_DIR, { recursive: true, force: true }); } catch {}
process.exit(fail > 0 ? 1 : 0);
