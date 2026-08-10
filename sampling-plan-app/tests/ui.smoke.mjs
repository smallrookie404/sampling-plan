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
  if (msg.type() === "error") {
    // 测试环境无外网：GitHub 预同步请求被沙箱拦截属预期，不视为失败
    if (msg.text().includes("ERR_NETWORK_ACCESS_DENIED")) return;
    errors.push("console: " + msg.text());
  }
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

// 列宽自适应：内容变长后列宽应随之变宽
const wBefore = await page.$eval("#grid-cols col:nth-child(5)", (el) => el.style.width);
await page.fill(
  '#main-grid tr[data-r="0"] td[data-c="D"] input',
  "这是一个非常非常长的接害因素名称用来测试列宽自适应功能是否正常工作"
);
await page.waitForTimeout(900);
const wAfter = await page.$eval("#grid-cols col:nth-child(5)", (el) => el.style.width);
if (parseFloat(wAfter) <= parseFloat(wBefore)) {
  throw new Error(`列宽未随内容自适应: ${wBefore} -> ${wAfter}`);
}
await page.fill('#main-grid tr[data-r="0"] td[data-c="D"] input', "二氧化钛粉尘(总尘)");
await page.waitForTimeout(900);
console.log("列宽自适应校验通过 ✔");

// 录入区：区域选择 / 复制 / 粘贴
await page.evaluate(() => {
  const b = document.querySelector('#main-grid tr[data-r="0"] td[data-c="B"]');
  b.dispatchEvent(new MouseEvent("mousedown", { bubbles: true }));
  const d = document.querySelector('#main-grid tr[data-r="0"] td[data-c="D"]');
  d.dispatchEvent(new MouseEvent("mousedown", { bubbles: true, shiftKey: true }));
});
const selCount = await page.$$eval('#main-grid tr[data-r="0"] td.sel', (els) => els.length);
if (selCount !== 3) throw new Error("区域选择失败: " + selCount);
const copied = await page.evaluate(() => {
  const ev = new ClipboardEvent("copy", { clipboardData: new DataTransfer(), bubbles: true, cancelable: true });
  document.querySelector('#main-grid tr[data-r="0"] td[data-c="D"]').dispatchEvent(ev);
  return ev.clipboardData.getData("text/plain");
});
if (copied !== "操作工\t投料\t二氧化钛粉尘(总尘)") throw new Error("复制内容不符: " + JSON.stringify(copied));
await page.evaluate(() => {
  const ev = new ClipboardEvent("paste", { clipboardData: new DataTransfer(), bubbles: true, cancelable: true });
  ev.clipboardData.setData("text/plain", "粘贴值1\t粘贴值2\t粘贴值3");
  document.querySelector('#main-grid tr[data-r="0"] td[data-c="D"]').dispatchEvent(ev);
});
await page.waitForTimeout(300);
if ((await val(0, "B")) !== "粘贴值1" || (await val(0, "C")) !== "粘贴值2" || (await val(0, "D")) !== "粘贴值3") {
  throw new Error("粘贴未生效");
}
// 还原第 1 行
await page.fill('#main-grid tr[data-r="0"] td[data-c="B"] input', "操作工");
await page.fill('#main-grid tr[data-r="0"] td[data-c="C"] input', "投料");
await page.fill('#main-grid tr[data-r="0"] td[data-c="D"] input', "二氧化钛粉尘(总尘)");
await page.waitForTimeout(300);
console.log("录入区 选择/复制/粘贴 校验通过 ✔");

// 粘贴填充：复制单个内容 → 整片选中区域全部填入
await page.evaluate(() => {
  const b0 = document.querySelector('#main-grid tr[data-r="0"] td[data-c="B"]');
  b0.dispatchEvent(new MouseEvent("mousedown", { bubbles: true }));
  const c1 = document.querySelector('#main-grid tr[data-r="1"] td[data-c="C"]');
  c1.dispatchEvent(new MouseEvent("mousedown", { bubbles: true, shiftKey: true }));
});
await page.evaluate(() => {
  const ev = new ClipboardEvent("paste", { clipboardData: new DataTransfer(), bubbles: true, cancelable: true });
  ev.clipboardData.setData("text/plain", "填充值");
  document.querySelector('#main-grid tr[data-r="1"] td[data-c="C"]').dispatchEvent(ev);
});
await page.waitForTimeout(300);
if (
  (await val(0, "B")) !== "填充值" ||
  (await val(0, "C")) !== "填充值" ||
  (await val(1, "B")) !== "填充值" ||
  (await val(1, "C")) !== "填充值"
) {
  throw new Error("单值整片填充未生效");
}
await page.fill('#main-grid tr[data-r="0"] td[data-c="B"] input', "操作工");
await page.fill('#main-grid tr[data-r="0"] td[data-c="C"] input', "投料");
await page.fill('#main-grid tr[data-r="1"] td[data-c="B"] input', "操作工");
await page.fill('#main-grid tr[data-r="1"] td[data-c="C"] input', "投料");
await page.waitForTimeout(300);

// 粘贴平铺：多格内容按行列规律重复填充到选中区域
await page.evaluate(() => {
  const b0 = document.querySelector('#main-grid tr[data-r="0"] td[data-c="B"]');
  b0.dispatchEvent(new MouseEvent("mousedown", { bubbles: true }));
  const b3 = document.querySelector('#main-grid tr[data-r="3"] td[data-c="B"]');
  b3.dispatchEvent(new MouseEvent("mousedown", { bubbles: true, shiftKey: true }));
});
await page.evaluate(() => {
  const ev = new ClipboardEvent("paste", { clipboardData: new DataTransfer(), bubbles: true, cancelable: true });
  ev.clipboardData.setData("text/plain", "甲\n乙");
  document.querySelector('#main-grid tr[data-r="3"] td[data-c="B"]').dispatchEvent(ev);
});
await page.waitForTimeout(300);
const bvals = [];
for (let i = 0; i <= 3; i++) bvals.push(await val(i, "B"));
if (bvals.join(",") !== "甲,乙,甲,乙") throw new Error("平铺填充未生效: " + bvals.join(","));
for (let i = 0; i <= 2; i++) {
  await page.fill(`#main-grid tr[data-r="${i}"] td[data-c="B"] input`, "操作工");
}
await page.fill('#main-grid tr[data-r="3"] td[data-c="B"] input', "");
await page.waitForTimeout(300);
console.log("录入区 整片填充/平铺粘贴 校验通过 ✔");

// 粘贴后无需重新点击，可直接 Delete 清空粘贴区域
await page.evaluate(() => {
  const b0 = document.querySelector('#main-grid tr[data-r="0"] td[data-c="B"]');
  b0.dispatchEvent(new MouseEvent("mousedown", { bubbles: true }));
  const d0 = document.querySelector('#main-grid tr[data-r="0"] td[data-c="D"]');
  d0.dispatchEvent(new MouseEvent("mousedown", { bubbles: true, shiftKey: true }));
});
await page.evaluate(() => {
  const ev = new ClipboardEvent("paste", { clipboardData: new DataTransfer(), bubbles: true, cancelable: true });
  ev.clipboardData.setData("text/plain", "待删A\t待删B\t待删C");
  document.querySelector('#main-grid tr[data-r="0"] td[data-c="D"]').dispatchEvent(ev);
});
await page.waitForTimeout(300);
if ((await val(0, "B")) !== "待删A") throw new Error("Delete 测试前置粘贴未生效");
await page.keyboard.press("Delete");
await page.waitForTimeout(300);
if ((await val(0, "B")) !== "" || (await val(0, "C")) !== "" || (await val(0, "D")) !== "") {
  throw new Error("粘贴后未重新点击，Delete 应直接清空粘贴区域");
}
await page.fill('#main-grid tr[data-r="0"] td[data-c="B"] input', "操作工");
await page.fill('#main-grid tr[data-r="0"] td[data-c="C"] input', "投料");
await page.fill('#main-grid tr[data-r="0"] td[data-c="D"] input', "二氧化钛粉尘(总尘)");
await page.waitForTimeout(300);
console.log("粘贴后立即 Delete 清空校验通过 ✔");

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
const hazardStatus = await page.textContent("#hazard-status");
const hazardCount = Number((hazardStatus.match(/共 (\d+) 条/) || [])[1] || 0);
if (hazardCount < 100) throw new Error("危害因素数量异常: " + hazardStatus);
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
if (!(await page.$eval("#sync-status", (el) => el.classList.contains("hidden")))) {
  throw new Error("未配置 GitHub 时不应显示同步状态条");
}
await page.click("#btn-gh");
await page.waitForSelector("#gh-modal:not(.hidden)");
const hint0 = await page.textContent("#gh-saved-hint");
if (!hint0.includes("未找到")) throw new Error("未配置时应提示未找到 Token");
// Token 显示/隐藏开关
if ((await page.getAttribute("#gh-token", "type")) !== "password") throw new Error("Token 应为密码类型");
await page.check("#gh-show-token");
if ((await page.getAttribute("#gh-token", "type")) !== "text") throw new Error("显示 Token 开关无效");
await page.uncheck("#gh-show-token");
if ((await page.getAttribute("#gh-token", "type")) !== "password") throw new Error("隐藏 Token 开关无效");
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
const hint1 = await page.textContent("#gh-saved-hint");
if (!hint1.includes("已保存") || !hint1.includes("t_TEST")) {
  throw new Error("保存后应显示已保存 Token 状态: " + hint1);
}
await page.click("#gh-clear");
await page.waitForTimeout(200);
const ghMsg2 = await page.textContent("#gh-msg");
if (!ghMsg2.includes("已清除")) throw new Error("清除配置失败: " + ghMsg2);
const ghCleared = await page.evaluate(() => localStorage.getItem("samplingPlanGithubConfig_v1"));
if (ghCleared !== null) throw new Error("GitHub 配置未清除");
await page.click("#gh-close");
await page.waitForSelector("#gh-modal.hidden", { state: "attached" });
await page.waitForTimeout(150);
if (!(await page.$eval("#sync-status", (el) => el.classList.contains("hidden")))) {
  throw new Error("清除配置后同步状态条应隐藏");
}
console.log("GitHub 配置界面校验通过 ✔");

// ---------- 10) 一键清空录入区 ----------
const rowsBeforeClear = await rowCountOf();
await page.click("#btn-clear-input");
await page.waitForTimeout(300);
if ((await val(0, "A")) !== "" || (await val(0, "D")) !== "") throw new Error("清空录入区未生效");
if ((await rowCountOf()) !== rowsBeforeClear) throw new Error("清空录入区不应改变行数");
console.log("一键清空录入区校验通过 ✔");

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
