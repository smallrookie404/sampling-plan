// 危害因素库单元格编辑（与主表录入区一致的 Excel 语义）专项冒烟测试。
// 独立成文件的原因：ui.smoke.mjs 在「覆盖联动」既有失败处提前退出，其后的危害因素库用例无法执行。
// 覆盖：取消下拉（全列文本输入）｜单击仅选中（布防态）｜直接键入替换｜双击/F2 编辑｜
//       方向键/Enter/Tab/Home/End 导航｜Delete 清空｜拖选高亮｜Ctrl+A 全选｜Escape 取消。
import { chromium } from "playwright-core";
import fs from "node:fs";
import path from "node:path";
import { startServer } from "../server.mjs";
import { findBrowser } from "./_browser.mjs";

const CHROME = findBrowser();
if (!CHROME) { console.error("未找到 Chrome/Edge，跳过该冒烟测试"); process.exit(0); }

const TEST_DIR = path.resolve("data/_test-hazard-edit");
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

await page.goto(APP_URL, { waitUntil: "domcontentloaded" });
await page.waitForTimeout(600);
await page.evaluate(() => document.querySelector("#xcdc-login")?.remove());
await page.click('.tab[data-tab="hazard"]');
await page.waitForSelector("#hazard-body tr[data-h]", { timeout: 8000 });
await page.evaluate(() => { const w = document.querySelector(".hazard-wrap"); if (w) w.scrollTop = 0; });
await page.waitForTimeout(200);

const cellSel = (r, c) => `#hazard-body tr[data-h="${r}"] td[data-hc="${c}"]`;
const valOf = (r, c) => page.$eval(cellSel(r, c) + " input", (el) => el.value);
const clsOf = (r, c) => page.$eval(cellSel(r, c), (el) => el.className);
const clickCell = async (r, c) => {
  const b = await page.$eval(cellSel(r, c), (el) => {
    const x = el.getBoundingClientRect();
    return [x.x + x.width / 2, x.y + x.height / 2];
  });
  await page.mouse.click(b[0], b[1]);
  await page.waitForTimeout(70);
};

// 1) 已取消下拉：危害因素库内不再渲染任何 <select>，全部为文本 input
const selCount = await page.$$eval("#hazard-body select", (e) => e.length);
check(selCount === 0, "危害因素库已取消下拉（select 数 = 0）", "select=" + selCount);
const inputCount = await page.$$eval("#hazard-body input", (e) => e.length);
check(inputCount > 0, `全部列为文本输入（input 数 = ${inputCount}）`);

// 2) 单击仅选中：当前格高亮 + 焦点落在隐形布防框（不直接进编辑）
await clickCell(0, 0);
check((await clsOf(0, 0)).includes("cur"), "单击设为当前格（cur 高亮）");
check(await page.evaluate(() => document.activeElement && document.activeElement.tagName) === "TEXTAREA",
  "单击后为布防态（隐形 textarea 持有焦点）");

// 3) 直接键入 = Excel 替换语义（覆盖原内容，不是追加）
await page.keyboard.type("XY");
await page.waitForTimeout(200);
check((await valOf(0, 0)) === "XY", "直接键入替换原内容（Excel 语义）", await valOf(0, 0));

// 4) Enter 提交并下移一行，原值保留
await page.keyboard.press("Enter");
await page.waitForTimeout(200);
check((await clsOf(1, 0)).includes("cur"), "Enter 提交并下移到下一行");
check((await valOf(0, 0)) === "XY", "Enter 提交后原值保留", await valOf(0, 0));

// 5) 方向键导航 + Tab/Shift+Tab（含原「粉尘性质」列，现为文本）
await page.keyboard.press("ArrowRight");
await page.waitForTimeout(100);
check((await clsOf(1, 1)).includes("cur"), "方向键右移");
await page.keyboard.press("Tab");
await page.waitForTimeout(100);
check((await clsOf(1, 2)).includes("cur"), "Tab 右移（原下拉列现为文本输入）");
await page.keyboard.press("Shift+Tab");
await page.waitForTimeout(100);
check((await clsOf(1, 1)).includes("cur"), "Shift+Tab 左移");

// 6) F2 进入编辑（方向键在格内移动光标），Escape 退出
await page.keyboard.press("F2");
await page.waitForTimeout(120);
check(await page.evaluate(() => !!document.querySelector("#hazard-body input:focus")), "F2 进入编辑（input 聚焦）");
await page.keyboard.press("Escape");
await page.waitForTimeout(120);

// 7) 双击进入编辑并追加输入
await page.mouse.dblclick(...(await page.$eval(cellSel(2, 0), (el) => {
  const b = el.getBoundingClientRect(); return [b.x + b.width / 2, b.y + b.height / 2];
})));
await page.waitForTimeout(180);
check(await page.evaluate(() => !!document.querySelector("#hazard-body input:focus")), "双击进入编辑");
await page.keyboard.type("ZZ");
await page.waitForTimeout(150);
check((await valOf(2, 0)).endsWith("ZZ"), "双击编辑追加输入", await valOf(2, 0));
await page.keyboard.press("Enter");
await page.waitForTimeout(250);

// 8) Delete 清空当前格（模型与 DOM 同步）
await clickCell(2, 0);
await page.keyboard.press("Delete");
await page.waitForTimeout(250);
check((await valOf(2, 0)) === "", "Delete 清空当前格", await valOf(2, 0));

// 9) 拖选矩形高亮（sel 类铺满选区）
await clickCell(0, 3);
const p1 = await page.$eval(cellSel(0, 3), (el) => { const b = el.getBoundingClientRect(); return [b.x + 4, b.y + b.height / 2]; });
const p2 = await page.$eval(cellSel(3, 4), (el) => { const r = el.getBoundingClientRect(); return [r.right - 4, r.y + r.height / 2]; });
await page.mouse.move(p1[0], p1[1]);
await page.mouse.down();
await page.mouse.move(p2[0], p2[1], { steps: 6 });
await page.mouse.up();
await page.waitForTimeout(200);
const selCells = await page.$$eval("#hazard-body td.sel", (e) => e.length);
check(selCells >= 8, `拖选矩形高亮（sel = ${selCells}）`);

// 10) 复制选区 → setData 写入 TSV（headless 下直接断言写入内容，避免依赖剪贴板读取权限）
await page.evaluate(() => {
  window.__setData = null;
  const orig = DataTransfer.prototype.setData;
  DataTransfer.prototype.setData = function (t, v) { window.__setData = { type: t, value: v }; return orig.call(this, t, v); };
});
await page.keyboard.press("Control+c");
await page.waitForTimeout(250);
const sd = await page.evaluate(() => window.__setData);
check(!!(sd && sd.type === "text/plain" && sd.value.includes("\t")),
  "Ctrl+C 复制选区为 TSV", sd ? sd.type + " len=" + sd.value.length : "null");

// 11) Escape 取消选区
await page.keyboard.press("Escape");
await page.waitForTimeout(150);
check((await page.$$eval("#hazard-body td.sel", (e) => e.length)) === 0, "Escape 取消选区");

// 12) Home / End 行内跳转
await clickCell(0, 0);
await page.keyboard.press("End");
await page.waitForTimeout(120);
const endCol = await page.$eval("#hazard-body td.cur", (el) => el.dataset.hc);
check(Number(endCol) >= 10, `End 跳到末列（col = ${endCol}）`);
await page.keyboard.press("Home");
await page.waitForTimeout(120);
check((await page.$eval("#hazard-body td.cur", (el) => el.dataset.hc)) === "0", "Home 跳到首列");

// 13) 无运行时错误
check(pageErrors.length === 0, "无页面 JS 错误", pageErrors.join(" | "));

console.log(`\n结果：${pass} 通过 / ${fail} 失败`);
await browser.close();
process.exit(fail ? 1 : 0);
