// 「数据报告处理」页签冒烟测试：页签可切换、两个按钮存在、空态提示正确
import { chromium } from "playwright-core";
import fs from "node:fs";
import path from "node:path";
import { startServer } from "../server.mjs";

const CANDIDATES = [
  "C:/Program Files/Google/Chrome/Application/chrome.exe",
  "C:/Program Files (x86)/Microsoft/Edge/Application/msedge.exe",
  "C:/Program Files/Microsoft/Edge/Application/msedge.exe",
];
const CHROME = CANDIDATES.find((p) => fs.existsSync(p));
if (!CHROME) { console.error("未找到 Chrome/Edge，跳过该冒烟测试"); process.exit(0); }
const SHOT_DIR = "tests/shots";
// 用独立临时目录：server.mjs 会在 dataFile 同级建 records/，避免污染真实的 data/records/
const TEST_DIR = path.resolve("data/_test-report");
const TEST_DATA = path.join(TEST_DIR, "records.json");
fs.mkdirSync(SHOT_DIR, { recursive: true });

const server = await startServer({ port: 0, dataFile: TEST_DATA });
const PORT = server.address().port;
const APP_URL = `http://127.0.0.1:${PORT}/`;
let browser;
process.on("exit", () => {
  try { server.close(); } catch {}
  try { fs.rmSync(TEST_DIR, { recursive: true, force: true }); } catch {}
  try { if (browser) browser.process()?.kill(); } catch {}
});

browser = await chromium.launch({ executablePath: CHROME, headless: true });
const page = await browser.newPage({ viewport: { width: 1600, height: 900 } });
const dialogs = [];
page.on("dialog", (d) => { dialogs.push(d.message()); d.dismiss(); });
await page.addInitScript(() => {
  sessionStorage.setItem("xcdc_session_v1", JSON.stringify({
    token: "test-token", orgId: "1",
    userInfo: { userName: "测试账号", userCode: "TEST", organization: { organizationId: "1" } },
  }));
});

const errors = [];
const IGNORE = /401|Failed to load resource|auth\/info/i;   // 假 token 会被平台拒绝，属测试环境预期
page.on("console", (m) => { if (m.type() === "error" && !IGNORE.test(m.text())) errors.push(m.text()); });
page.on("pageerror", (e) => { if (!IGNORE.test(e.message)) errors.push("pageerror: " + e.message); });

// 平台接口 mock：必须在 goto 之前注册。
// 否则启动期的 /auth/info 401 会触发「登录失效」把会话清掉（token=null），
// 后面点「回退」只会走 showLogin() 而不会发出任何请求。
const seen = { findList: "", jcbg: "", reset: [] };
const j = (body) => ({ status: 200, contentType: "application/json", body: JSON.stringify(body) });
await page.route("**/auth/info*", (route) => route.fulfill(j({ user: { userName: "测试账号" } })));
await page.route("**/api/reportData/findList*", (route) => {
  seen.findList = route.request().url();
  return route.fulfill(j({ code: "200", body: { total: 1, records: [
    { id: 6205, code: "BTC26-SZJC0959", belongInspectName: "浙江永贵电器股份有限公司", status: "6" }] } }));
});
await page.route("**/api/projectJcbgLastAffix*", (route) => {
  seen.jcbg = route.request().url();
  return route.fulfill(j({ code: "200", body: { total: 1, records: [
    { id: 6205, code: "BTC26-SZJC0959", belongInspectName: "浙江永贵电器股份有限公司", status: "6", isDel: "1" }] } }));
});
// 选中项目后会连带请求这几个接口；不 mock 就会 401 → 同样清掉会话
await page.route("**/api/users*", (route) => route.fulfill(j({ code: "200", body: { content: [] } })));
await page.route("**/api/allocation/projectHuman*", (route) => route.fulfill(j({ code: "200", body: { records: [] } })));
await page.route("**/api/projectXcdcUser*", (route) => route.fulfill(j({ code: "200", body: { records: [] } })));
// 账号桥 + 重置接口：平台实测返回 body:false 表示「本次无需变更」，仍是成功
// ping 直接返回 up:true，避免测试里走到「真实自动拉起」分支（那会等最多 12s 导致超时）
await page.route("**/api/bridge/ping", (route) =>
  route.fulfill(j({ ok: true, autoStart: true, up: true })));
await page.route("**/token/*", (route) => route.fulfill(j({ token: "bridge-token", orgId: "1" })));

// 审核链（4 步）用到的接口。注意各 pattern 互不重叠（`*` 不跨 `/`），顺序无关。
await page.route("**/api/projectJcdgLastAffixNew/submitJcdgReview*", (route) =>
  route.fulfill(j({ code: "200", body: { prjCode: "BTC26-SZJC0959", msg: "提交成功" } })));
await page.route("**/api/projectJcdgLastAffixNew/auditAll*", (route) =>
  route.fulfill(j({ code: "200", body: [{ prjCode: "BTC26-SZJC0959", msg: "通过" }] })));
await page.route("**/api/projectJcbgLastAffixNew/auditAll*", (route) =>
  route.fulfill(j({ code: "200", body: [{ prjCode: "BTC26-SZJC0959", msg: "通过" }] })));
await page.route("**/api/projectJcdgLastAffix*", (route) =>
  route.fulfill(j({ code: "200", body: { total: 1, records: [
    { id: 6205, code: "BTC26-SZJC0959", status: "0", isDel: "0" }] } })));
await page.route("**/projectXccyLastAffix/resetProjectApprovalProcess*", (route) =>
  route.fulfill(j({ code: "200", message: "重置成功！", body: false })));
page.on("request", (r) => { if (/resetProjectApprovalProcess/.test(r.url())) seen.reset.push(r.url()); });

await page.goto(APP_URL, { waitUntil: "domcontentloaded" });
await page.waitForTimeout(1200);
// 登录遮罩 z-index:1000 会拦截点击；假 token 会让启动流程重新弹出它，直接移除该元素
await page.evaluate(() => {
  const kill = () => { const el = document.getElementById("xcdc-login"); if (el) el.remove(); };
  kill();
  new MutationObserver(kill).observe(document.body, { childList: true, subtree: true });
});
await page.waitForTimeout(300);

const results = [];
const check = (name, ok, extra) => { results.push({ name, ok, extra }); };

// 1) 页签存在且在「噪声数据结果」之后
const tabLabels = await page.$$eval("nav.tabs .tab", (els) => els.map((e) => e.textContent.trim()));
check("页签顺序含「数据报告处理」且在噪声之后",
  tabLabels.indexOf("数据报告处理") === tabLabels.indexOf("噪声数据结果") + 1, tabLabels.join(" | "));

// 2) 点击页签 → 面板切换
await page.click('nav.tabs .tab[data-tab="report"]');
await page.waitForTimeout(300);
const panelActive = await page.$eval("#tab-report", (el) => el.classList.contains("active"));
check("点击后 #tab-report 激活", panelActive);

// 3) 两个按钮存在且可见
const resetVisible = await page.isVisible("#report-reset");
const auditVisible = await page.isVisible("#report-audit");
check("「数据报告回退」按钮可见", resetVisible);
check("「数据报告审核」按钮可见", auditVisible);

// 4) 未选项目时显示空态提示
const emptyText = await page.$eval("#report-empty", (el) => el.textContent.trim());
check("未选项目时显示空态提示", /尚未在「数据上传」中选择项目/.test(emptyText), emptyText.slice(0, 40));

// 5) 未选项目时点「回退」→ 弹提示且不发请求
let requested = false;
page.on("request", (r) => { if (/resetProjectApprovalProcess|auditAll/.test(r.url())) requested = true; });
await page.click("#report-reset");
await page.waitForTimeout(400);
check("未选项目点回退会提示且不请求接口", dialogs.length > 0 && !requested,
  "dialogs=" + dialogs.length + " requested=" + requested);

// 6) 选中项目后点「回退」→ 必须**先在「已完结」项目里定位**
//    （GET /api/projectJcbgLastAffix?isDel=1；回退只对已完结项目有效，用 findList 搜出来的会「找不到」）
//    接口 mock 已在 goto 之前注册（见上）。
await page.evaluate(() => {
  document.getElementById("xcdc-upload").classList.remove("hidden");
  const code = document.getElementById("projectCode");
  code.value = "0959";
  code.dispatchEvent(new Event("input", { bubbles: true }));
});
await page.waitForTimeout(1200);
const items = await page.$$eval("#suggest .suggest-item", (els) => els.map((e) => e.textContent.trim()));
if (items.length) await page.click("#suggest .suggest-item");
await page.waitForTimeout(400);
const projInfo = await page.$eval("#projectInfo", (e) => e.textContent.trim());
check("「数据上传」搜索能列出并选中项目", items.length > 0 && /BTC26-SZJC0959/.test(projInfo),
  items.join(" | ") + " ／ 已选=" + projInfo);

// 关掉上传遮罩（它会拦截页签点击）→ 切「数据报告处理」→ 点回退
await page.evaluate(() => { document.getElementById("xcdc-upload").classList.add("hidden"); });
await page.click('nav.tabs .tab[data-tab="report"]');
await page.waitForTimeout(200);
await page.click("#report-reset");
await page.waitForTimeout(1500);
const rstat = await page.$eval("#report-status", (e) => e.textContent.trim());

check("回退前先在「已完结」项目中定位（isDel=1 + 项目编号）",
  /projectJcbgLastAffix\?[^#]*isDel=1/.test(seen.jcbg) && /code=BTC26-SZJC0959/.test(seen.jcbg),
  seen.jcbg || "（未发出请求）");
check("定位成功并提示；三重确认被取消时不发重置请求",
  /已在「已完结」项目中定位到/.test(rstat) && seen.reset.length === 0,
  "status=" + rstat + " ／ reset=" + seen.reset.length);

// 7) 三重确认通过后：平台返回 {code:'200', body:false}（本次无需变更）**不能**误报失败
await page.evaluate(() => { window.confirm = () => true; });
const dlgBefore = dialogs.length;
await page.click("#report-reset");
await page.waitForTimeout(2000);
const rstat2 = await page.$eval("#report-status", (e) => e.textContent.trim());
const newDlg = dialogs.slice(dlgBefore);
check("body:false 视为成功（两步都跑完，不误报「重置失败」）",
  seen.reset.length === 2 && /两步回退全部完成/.test(rstat2)
    && newDlg.some((m) => /回退成功/.test(m)) && !newDlg.some((m) => /失败/.test(m)),
  "reset次数=" + seen.reset.length + " ／ status=" + rstat2 + " ／ 弹窗=" + JSON.stringify(newDlg).slice(0, 160));

// 8) 「数据报告审核」= 4 步跨账号审批链
//    ① btc-jyg 提交审复核（检测底稿） ② btc-wcy 检测底稿通过
//    ③ btc-zhj 检测结果报告单通过    ④ btc-gxz 检测结果报告单通过
const calls = [];      // 审批类请求（按发生顺序）
const tokenCalls = []; // 账号桥取令牌的账号（按发生顺序）
page.on("request", (r) => {
  const u = r.url();
  if (/submitJcdgReview|auditAll/.test(u)) calls.push(u.replace(/^https?:\/\/[^/]+/, "").split("?")[0]);
  const m = /\/token\/([^/?#]+)/.exec(u);
  if (m) tokenCalls.push(decodeURIComponent(m[1]));
});
const dlgBefore2 = dialogs.length;
// 缩短轮询等待时间：mock 固定响应不会变，避免等满 30s
await page.evaluate(() => { window.__AUDIT_POLL_OPTS__ = { maxWaitMs: 400, intervalMs: 100 }; });
await page.click("#report-audit");
await page.waitForTimeout(8000);
const rstat3 = await page.$eval("#report-status", (e) => e.textContent.trim());
const newDlg2 = dialogs.slice(dlgBefore2);
const expectCalls = [
  "/api/projectJcdgLastAffixNew/submitJcdgReview",
  "/api/projectJcdgLastAffixNew/auditAll",
  "/api/projectJcbgLastAffixNew/auditAll",
  "/api/projectJcbgLastAffixNew/auditAll",
];
check("审核按 4 步链顺序执行（提交审复核 → jcdg通过 → jcbg通过 ×2）",
  JSON.stringify(calls) === JSON.stringify(expectCalls),
  JSON.stringify(calls));
check("4 步分别用各自账号取令牌（btc-jyg → btc-wcy → btc-zhj → btc-gxz）",
  JSON.stringify(tokenCalls.slice(-4)) === JSON.stringify(["btc-jyg", "btc-wcy", "btc-zhj", "btc-gxz"]),
  JSON.stringify(tokenCalls.slice(-4)));
check("审核链全部完成且无失败弹窗",
  /审核链全部完成/.test(rstat3) && !newDlg2.some((m) => /失败/.test(m)),
  "status=" + rstat3 + " ／ 弹窗=" + JSON.stringify(newDlg2).slice(0, 140));

await page.screenshot({ path: path.join(SHOT_DIR, "report-tab.png") });
check("无控制台错误", errors.length === 0, errors.slice(0, 2).join(" ; "));

await browser.close();

console.log("\n=== 数据报告处理页签冒烟 ===");
let pass = 0;
for (const r of results) {
  console.log((r.ok ? "✅" : "❌") + " " + r.name + (r.extra ? "  → " + r.extra : ""));
  if (r.ok) pass++;
}
console.log(`\n${pass}/${results.length} 通过`);
process.exit(pass === results.length ? 0 : 1);
