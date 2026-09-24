import { chromium } from "playwright-core";
import http from "node:http";
import fs from "node:fs";
import path from "node:path";

const ROOT = "G:/AtomCode/采样计划/sampling-plan-app";
const MIME = { ".html": "text/html", ".js": "text/javascript", ".css": "text/css", ".json": "application/json" };
const srv = http.createServer((req, res) => {
  let p = decodeURIComponent(req.url.split("?")[0]);
  if (p === "/") p = "/index.html";
  const f = path.join(ROOT, p);
  try {
    const data = fs.readFileSync(f);
    res.writeHead(200, { "Content-Type": MIME[path.extname(f)] || "application/octet-stream" });
    res.end(data);
  } catch {
    res.writeHead(404); res.end("nf");
  }
});
await new Promise((r) => srv.listen(0, r));
const port = srv.address().port;

const browser = await chromium.launch({ executablePath: "C:/Program Files/Google/Chrome/Application/chrome.exe", headless: true });
const page = await browser.newPage();
const errors = [];
page.on("pageerror", (e) => errors.push("PAGEERROR: " + e.message));
page.on("console", (m) => { if (m.type() === "error") errors.push("CONSOLE: " + m.text()); });

await page.goto(`http://127.0.0.1:${port}/index.html`, { waitUntil: "load" });
await page.waitForTimeout(1200);
await page.evaluate(() => {
  const ov = document.getElementById("xcdc-login");
  if (ov) ov.remove();
});
await page.click('.tab[data-tab="survey"]');
await page.waitForTimeout(400);

// 1) 单击选中第一个单元格
const cell = await page.evaluate(() => {
  const td = document.querySelector("#survey-body tr[data-r] td[data-c]");
  const r = td.getBoundingClientRect();
  return { x: r.x + r.width / 2, y: r.y + r.height / 2 };
});
await page.mouse.click(cell.x, cell.y);
await page.waitForTimeout(200);
const sel = await page.evaluate(() => ({
  cur: document.querySelectorAll("#survey-body td.cur").length,
  sel: document.querySelectorAll("#survey-body td.sel").length,
}));
console.log("选中:", JSON.stringify(sel));

// 2) 直接打字 → 进入编辑并替换
await page.keyboard.type("甲");
await page.waitForTimeout(200);
const editing = await page.evaluate(() => {
  const input = document.querySelector("#survey-body td.cur input, #survey-body td.sel input");
  return { hasInput: !!input, value: input ? input.value : null, focused: input && document.activeElement === input };
});
console.log("打字进入编辑:", JSON.stringify(editing));

// 3) Enter 提交并下移
await page.keyboard.press("Enter");
await page.waitForTimeout(200);
const after = await page.evaluate(() => {
  const stored = JSON.parse(localStorage.getItem("samplingPlanSurvey_v1") || "{}");
  return {
    firstCell: stored.process && stored.process[0][0],
    tdText: document.querySelector('#survey-body tr[data-r="0"] td[data-c="0"]').textContent,
    inputLeft: document.querySelectorAll("#survey-body input").length,
  };
});
console.log("提交后:", JSON.stringify(after));

// 4) 双击进入编辑、Esc 取消
await page.mouse.dblclick(cell.x, cell.y);
await page.waitForTimeout(200);
const dbl = await page.evaluate(() => !!document.querySelector("#survey-body td input"));
console.log("双击进入编辑:", dbl);
await page.keyboard.press("Escape");
await page.waitForTimeout(200);
const esc = await page.evaluate(() => ({
  inputLeft: document.querySelectorAll("#survey-body input").length,
  tdText: document.querySelector('#survey-body tr[data-r="0"] td[data-c="0"]').textContent,
}));
console.log("Esc 还原:", JSON.stringify(esc));

console.log("errors:", errors.filter((e) => !e.includes("404")).slice(0, 5));
await browser.close();
srv.close();
