import { chromium } from "playwright-core";
import http from "node:http";
import fs from "node:fs";
import path from "node:path";

const ROOT = process.cwd();
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
await page.waitForTimeout(1500);

const info = await page.evaluate(() => ({
  hasSurveySheets: !!window.SurveySheets,
  tabsHtml: document.getElementById("survey-tabs") ? document.getElementById("survey-tabs").innerHTML.slice(0, 300) : "NO #survey-tabs",
  panelExists: !!document.getElementById("tab-survey"),
  panelClass: document.getElementById("tab-survey") ? document.getElementById("tab-survey").className : "",
  tabBtn: !!document.querySelector('.tab[data-tab="survey"]'),
}));
console.log(JSON.stringify(info, null, 2));

// 点击页签再看
await page.click('.tab[data-tab="survey"]').catch(() => {});
await page.waitForTimeout(600);
const info2 = await page.evaluate(() => ({
  panelClass: document.getElementById("tab-survey")?.className,
  tabsHtml: document.getElementById("survey-tabs")?.innerHTML.slice(0, 300),
  stabCount: document.querySelectorAll("#survey-tabs .stab").length,
  bodyRows: document.querySelectorAll("#survey-body tr").length,
}));
console.log(JSON.stringify(info2, null, 2));
console.log("errors:", errors.slice(0, 8));
await browser.close();
srv.close();
