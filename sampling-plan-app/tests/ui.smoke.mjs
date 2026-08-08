import { chromium } from "playwright";
import fs from "node:fs";
import path from "node:path";
import { createRequire } from "node:module";
import { startServer } from "../server.mjs";

const require = createRequire(import.meta.url);
const xlsxio = require("../js/xlsxio.js");

const CHROME = "C:/Program Files/Google/Chrome/Application/chrome.exe";
const SHOT_DIR = "tests/shots";
const TEST_DATA = path.resolve("data/test-records.json");
fs.mkdirSync(SHOT_DIR, { recursive: true });

const server = await startServer({ port: 0, dataFile: TEST_DATA });
const PORT = server.address().port;
const APP_URL = `http://127.0.0.1:${PORT}/`;
process.on("exit", () => {
  try { server.close(); } catch {}
  try { fs.rmSync(TEST_DATA, { force: true }); } catch {}
});

const browser = await chromium.launch({ executablePath: CHROME, headless: true });
const page = await browser.newPage({ viewport: { width: 1600, height: 900 } });
page.on("dialog", (d) => d.accept());

const errors = [];
page.on("console", (msg) => {
  if (msg.type() === "error") errors.push("console: " + msg.text());
});
page.on("pageerror", (err) => errors.push("pageerror: " + err.message));

await page.goto(APP_URL, { waitUntil: "load" });
await page.waitForSelector("#main-grid tbody tr[data-r]", { timeout: 15000 });
const noticeHidden = await page.$eval("#storage-notice", (el) => el.classList.contains("hidden"));
if (!noticeHidden) throw new Error("服务模式下不应显示临时模式提示条");

const val = async (r, c) =>
  await page.$eval(`#main-grid tr[data-r="${r}"] td[data-c="${c}"] input`, (el) => el.value);
const selVal = async (r, c) =>
  await page.$eval(`#main-grid tr[data-r="${r}"] td[data-c="${c}"] select`, (el) => el.value);

// 1) 计算值校验
const checks = [
  [0, "W", "3号车间"], [0, "X", "操作工"], [0, "AN", "二氧化钛粉尘"],
  [0, "AU", "总尘"], [0, "BB", "3"], [0, "AM", "操作设备投料作业"],
  [7, "AN", "噪声"], [7, "BD", "设备运行"], [7, "AZ", "1"], [7, "AS", "短时间"],
];
for (const [r, c, exp] of checks) {
  const got = await val(r, c);
  if (got !== exp) throw new Error(`单元格 r${r} ${c} 期望 ${exp} 实际 ${got}`);
}
const status = await page.textContent("#grid-status");
if (!status.includes("错误 0 处")) throw new Error("初始状态应有 0 错误: " + status);
console.log("计算值与状态校验通过 ✔");

await page.screenshot({ path: SHOT_DIR + "/1-main.png" });

// 2) 输入校验联动：把第 2 行班制改成非法值
await page.selectOption(`#main-grid tr[data-r="1"] td[data-c="R"] select`, { label: "单班" });
await page.evaluate(() => {
  const sel = document.querySelector(`#main-grid tr[data-r="1"] td[data-c="R"] select`);
  // 模拟手动输入非法值：直接写入 select 不可能，改用 input 的 R 不可行；改为通过引擎验证
  window.__testBad = true;
});
// 改为校验 AN 找不到：把第 2 行接害因素改为不存在的因素
await page.fill(`#main-grid tr[data-r="1"] td[data-c="D"] input`, "不存在的因素XYZ");
await page.waitForTimeout(150);
const status2 = await page.textContent("#grid-status");
if (!status2.includes("错误 1 处")) throw new Error("非法因素应产生 1 处错误: " + status2);
const errCell = await page.$(`#main-grid tr[data-r="1"] td[data-c="AN"].error`);
if (!errCell) throw new Error("AN 错误单元格未标红");
console.log("错误联动校验通过 ✔");

// 恢复第 2 行
await page.fill(`#main-grid tr[data-r="1"] td[data-c="D"] input`, "二氧化钛粉尘(总尘)");
await page.waitForTimeout(150);

// 3) 覆盖联动：第 2 行检测方式覆盖为个体
await page.selectOption(`#main-grid tr[data-r="1"] td[data-c="AR"] select`, { label: "个体" });
await page.waitForTimeout(150);
const as1 = await page.$eval(`#main-grid tr[data-r="1"] td[data-c="AS"] input`, (el) => el.value);
const ak1 = await page.$eval(`#main-grid tr[data-r="1"] td[data-c="AK"] input`, (el) => el.value);
if (as1 !== "长时间" || ak1 !== "采样对象") throw new Error(`覆盖联动失败 AS=${as1} AK=${ak1}`);
console.log("覆盖联动校验通过 ✔");

// 4) 危害因素页
await page.click('.tab[data-tab="hazard"]');
await page.waitForSelector("#hazard-body tr", { timeout: 5000 });
const hazardRows = await page.$$eval("#hazard-body tr", (trs) => trs.length);
const hazardStatus = await page.textContent("#hazard-status");
if (hazardRows < 600) throw new Error("危害因素行数过少: " + hazardRows);
console.log("危害因素页:", hazardStatus, "✔");
await page.screenshot({ path: SHOT_DIR + "/2-hazard.png" });

// 5) 检测项目页
await page.click('.tab[data-tab="items"]');
await page.waitForSelector("#items-list .item");
const itemCount = await page.$$eval("#items-list .item", (els) => els.length);
const itemStatus = await page.textContent("#items-status");
if (itemCount < 400) throw new Error("检测项目显示过少: " + itemCount);
console.log("检测项目页:", itemStatus, "✔");
await page.screenshot({ path: SHOT_DIR + "/3-items.png" });

// 6) 浏览器内导出 Excel
await page.click('.tab[data-tab="main"]');
await page.waitForSelector("#main-grid");
const [download] = await Promise.all([
  page.waitForEvent("download", { timeout: 15000 }),
  page.click("#btn-export"),
]);
// 导出前恢复 AR 覆盖，避免 confirm 干扰（无错误时直接导出）
const xlsxPath = SHOT_DIR + "/exported.xlsx";
await download.saveAs(xlsxPath);
const size = fs.statSync(xlsxPath).size;
if (size < 1000) throw new Error("导出文件过小: " + size);
console.log("浏览器导出成功，字节数:", size, "✔");

// 验证导出文件只含自动计算区（单张工作表、39 列）
const xlSheets = await xlsxio.readWorkbook(fs.readFileSync(xlsxPath));
if (xlSheets.length !== 1 || xlSheets[0].name !== "测点布局情况调查") {
  throw new Error("导出应只有自动计算区单张工作表");
}
const xlArr = xlsxio.sheetToArray(xlSheets[0]);
if (xlArr[0].length !== 39 || String(xlArr[0][0]) !== "*单元/工作场所") {
  throw new Error("导出表头应为自动计算区 39 列");
}
if (String(xlArr[1][17]) !== "二氧化钛粉尘") throw new Error("导出首行 AN 值不符");
console.log("导出内容校验（仅自动计算区）通过 ✔");

// CSV 导出同样仅含自动计算区
const [csvDownload] = await Promise.all([
  page.waitForEvent("download", { timeout: 15000 }),
  page.click("#btn-csv"),
]);
const csvPath = SHOT_DIR + "/exported.csv";
await csvDownload.saveAs(csvPath);
const csvText = fs.readFileSync(csvPath, "utf8");
const csvHead = csvText.replace(/^\ufeff/, "").split("\r\n")[0].split(",");
if (csvHead.length !== 39 || csvHead[0] !== "*单元/工作场所" || csvHead.includes("车间")) {
  throw new Error("CSV 应只含自动计算区 39 列");
}
console.log("CSV 导出校验（仅自动计算区）通过 ✔");

// ---------- 7) 行操作支持自定义行数 ----------
const statusText = async () => page.textContent("#grid-status");
const rowCountOf = async () => {
  const m = (await statusText()).match(/共 (\d+) 行/);
  return m ? Number(m[1]) : -1;
};
async function doRowOp(btnId, count) {
  await page.click(btnId);
  await page.waitForSelector("#prompt-modal:not(.hidden)");
  await page.fill("#prompt-input", String(count));
  await page.click("#prompt-ok");
  await page.waitForSelector("#prompt-modal.hidden", { state: "attached" });
  await page.waitForTimeout(120);
}
let n0 = await rowCountOf(); // 8
await doRowOp("#btn-add", 3);
if ((await rowCountOf()) !== n0 + 3) throw new Error("新增自定义行数无效");
await page.click('#main-grid tr[data-r="0"] td.rowno');
await doRowOp("#btn-copy", 2);
if ((await rowCountOf()) !== n0 + 5) throw new Error("复制自定义行数无效");
await doRowOp("#btn-insert", 2);
if ((await rowCountOf()) !== n0 + 7) throw new Error("插入自定义行数无效");
await doRowOp("#btn-del", 3);
if ((await rowCountOf()) !== n0 + 4) throw new Error("删除自定义行数无效");
console.log("行操作自定义行数校验通过 ✔");

// ---------- 8) 数据记录：保存 / 搜索 / 调用 / 持久化 / 删除 ----------
await page.click("#btn-save");
await page.waitForSelector("#prompt-modal:not(.hidden)");
await page.fill("#prompt-input", "测试记录A");
await page.click("#prompt-ok");
await page.waitForSelector("#prompt-modal.hidden", { state: "attached" });
await page.waitForTimeout(200);

// 验证数据已写入项目文件夹中的记录文件（而非浏览器数据库）
const savedFile = JSON.parse(fs.readFileSync(TEST_DATA, "utf8"));
if (!Array.isArray(savedFile) || !savedFile.some((r) => r.name === "测试记录A")) {
  throw new Error("数据未写入项目文件夹 data/test-records.json");
}
console.log("数据文件落盘校验通过 ✔（记录已写入项目文件夹 JSON 文件）");

await page.click("#btn-db");
await page.waitForSelector("#db-modal:not(.hidden)");
await page.fill("#db-search", "测试记录A");
await page.waitForTimeout(150);
let dbItems = await page.$$eval("#db-list .db-item", (els) => els.length);
if (dbItems !== 1) throw new Error("数据库搜索未找到记录");
const metaText = await page.textContent("#db-list .db-item .db-item-meta");
const savedRows = Number(metaText.match(/(\d+) 行/)[1]);
await page.click('#db-list .db-item button[data-act="load"]');
await page.waitForSelector("#db-modal.hidden", { state: "attached" });
await page.waitForTimeout(200);
if (await val(0, "AN") !== "二氧化钛粉尘") throw new Error("调用后 AN 值不符");
if ((await rowCountOf()) !== savedRows) throw new Error("调用后行数不符");

// 刷新页面：记录仍在数据库中，可再次调用
await page.reload({ waitUntil: "load" });
await page.waitForSelector("#main-grid tbody tr[data-r]");
await page.click("#btn-db");
await page.waitForSelector("#db-modal:not(.hidden)");
await page.fill("#db-search", "测试记录A");
await page.waitForTimeout(150);
await page.click('#db-list .db-item button[data-act="load"]');
await page.waitForSelector("#db-modal.hidden", { state: "attached" });
await page.waitForTimeout(200);
if ((await rowCountOf()) !== savedRows) throw new Error("刷新后调用失败");

// 清理：删除测试记录
await page.click("#btn-db");
await page.waitForSelector("#db-modal:not(.hidden)");
await page.fill("#db-search", "测试记录A");
await page.waitForTimeout(150);
await page.click('#db-list .db-item button[data-act="del"]');
await page.waitForTimeout(200);
dbItems = await page.$$eval("#db-list .db-item", (els) => els.length);
if (dbItems !== 0) throw new Error("删除记录失败");
const afterDelete = JSON.parse(fs.readFileSync(TEST_DATA, "utf8"));
if (afterDelete.some((r) => r.name === "测试记录A")) throw new Error("删除后记录文件仍残留数据");
console.log("数据记录 保存/搜索/调用/持久化/删除 校验通过 ✔");

// ---------- 9) GitHub 配置界面 ----------
await page.click("#db-close");
await page.waitForSelector("#db-modal.hidden", { state: "attached" });
await page.click("#btn-gh");
await page.waitForSelector("#gh-modal:not(.hidden)");
await page.fill("#gh-repo", "testowner/sampling-plan");
await page.fill("#gh-token", "github_pat_TEST");
await page.click("#gh-save");
await page.waitForTimeout(800);
const ghClass = await page.getAttribute("#gh-modal", "class");
const ghMsg = await page.textContent("#gh-msg");
const ghStored = await page.evaluate(() => localStorage.getItem("samplingPlanGithubConfig_v1"));
if (!ghClass.includes("hidden")) {
  throw new Error(`GH 弹窗未关闭 class=${ghClass} msg=${ghMsg} stored=${ghStored} 浏览器错误=${JSON.stringify(errors)}`);
}
// 服务模式下存储模式不应被切换，提示条保持隐藏
if (!(await page.$eval("#storage-notice", (el) => el.classList.contains("hidden")))) {
  throw new Error("服务模式下不应显示存储提示条");
}
// 清除配置
await page.click("#btn-gh");
await page.waitForSelector("#gh-modal:not(.hidden)");
await page.click("#gh-clear");
await page.waitForTimeout(200);
const ghMsg2 = await page.textContent("#gh-msg");
if (!ghMsg2.includes("已清除")) throw new Error("清除配置失败: " + ghMsg2);
const ghCleared = await page.evaluate(() => localStorage.getItem("samplingPlanGithubConfig_v1"));
if (ghCleared !== null) throw new Error("GitHub 配置未清除");
await page.click("#gh-close");
await page.waitForSelector("#gh-modal.hidden", { state: "attached" });
await page.waitForTimeout(150);
console.log("GitHub 配置界面校验通过 ✔");

await page.screenshot({ path: SHOT_DIR + "/4-main-after-export.png" });

if (errors.length) {
  console.log("浏览器错误:", errors);
  process.exit(1);
}
console.log("UI 冒烟测试全部通过 ✔");
await browser.close();
server.closeAllConnections();
await server.close();
fs.rmSync(TEST_DATA, { force: true });
