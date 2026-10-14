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
// ⚠ 实测要点（2026-10-15）：通过类是 `audit`（**单个对象**），不是 auditAll（数组 → 400）；
//   `type` 必须取 getAffix 的当前节点号（不是列表 status）；jcbg 的 type 须为数字。
await page.route("**/api/projectJcdgLastAffixNew/submitJcdgReview*", (route) =>
  route.fulfill(j({ code: "200", body: { prjCode: "BTC26-SZJC0959", msg: "检测底稿提交成功！" } })));
// ★ 检测报告单的「提交审复核」是**独立接口**（前端「检测报告单」页的按钮用的就是它）
await page.route("**/api/reportFile/submitJcbgReview*", (route) =>
  route.fulfill(j({ code: "200", body: { prjCode: "BTC26-SZJC0959", msg: "检测报告提交成功" } })));
await page.route("**/api/projectJcdgLastAffixNew/audit*", (route) =>
  route.fulfill(j({ code: "200", message: "审批成功！", body: "审批成功！" })));
await page.route("**/api/projectJcbgLastAffixNew/audit*", (route) =>
  route.fulfill(j({ code: "200", message: "审批成功！", body: "审批成功！" })));
// getAffix 返回「当前节点号」：mock 成状态机——每被 audit 通过一次就推进到下一节点
//   jcdg：初始无记录(未提交) → 调 submitJcdgReview 后 type=0(理化复核) → 通过后无记录(流程完结)
//   jcbg：初始无记录(未提交) → 调 reportFile/submitJcbgReview 后 type=2(理化审核) → 通过后 type=6 → 再通过 isOver=1
//   ⚠ 关键：jcbg 只有在**调用了独立提交接口**之后才出现待审记录——用来锁定「第①步必须补交检测报告单」
const affixState = { jcdg: null, jcbg: null, jcbgOver: false };
// 先注册「列表」路由，再注册「getAffix」路由（Playwright 后注册者优先，保证 getAffix 命中）
await page.route("**/api/projectJcdgLastAffix*", (route) =>
  route.fulfill(j({ code: "200", body: { total: 1, records: [
    { id: 6205, code: "BTC26-SZJC0959", status: "0", isDel: "0" }] } })));
await page.route("**/api/projectJcdgLastAffixNew/getAffix*", (route) =>
  route.fulfill(j({ code: "200", body: affixState.jcdg == null ? [] : [{ id: 6255, type: affixState.jcdg, isOver: "0" }] })));
await page.route("**/api/projectJcbgLastAffix/getAffix*", (route) =>
  route.fulfill(j({ code: "200", body: affixState.jcbg == null ? [] : [{
    id: 6198, type: affixState.jcbg, isOver: affixState.jcbgOver ? "1" : "0",
    // 「审批完结检测结果报告单」PDF 路径（签发版）——供「查看审批完结检测结果报告单」用
    jcbgAffix: "/src/ohit/jcbgReport/913310020555095418/jcbg/BTC26-SZJC0959-x/review/BTC26-SZJC0959_检测报告_签发.pdf",
  }] })));
// audit 被调用后推进 mock 节点状态（供下一次 getAffix 读到新节点）
page.on("request", (r) => {
  const u = r.url();
  if (r.method() === "POST" && /submitJcdgReview/.test(u)) { affixState.jcdg = "0"; }
  else if (r.method() === "POST" && /reportFile\/submitJcbgReview/.test(u)) { affixState.jcbg = "2"; }
  else if (r.method() === "POST" && /projectJcdgLastAffixNew\/audit$/.test(u)) { affixState.jcdg = null; }
  else if (r.method() === "POST" && /projectJcbgLastAffixNew\/audit$/.test(u)) {
    if (affixState.jcbg === "2") affixState.jcbg = "6"; else { affixState.jcbgOver = true; }
  }
});
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

// 3) 只剩「数据报告回退」「数据报告审核」两个按钮；「查看…报告单」按钮已移除
const resetVisible = await page.isVisible("#report-reset");
const auditVisible = await page.isVisible("#report-audit");
const viewFinalExists = await page.$("#report-view-final");
check("「数据报告回退」按钮可见", resetVisible);
check("「数据报告审核」按钮可见", auditVisible);
check("「查看…报告单」按钮已移除（审核成功自动显示）", viewFinalExists === null);

// 4) 未选项目时显示空态提示
const emptyText = await page.$eval("#report-empty", (el) => el.textContent.trim());
check("未选项目时显示空态提示", /尚未在「数据上传」中选择项目/.test(emptyText), emptyText.slice(0, 40));

// 5) 未选项目时点「回退」→ 弹提示且不发请求
let requested = false;
page.on("request", (r) => { if (/resetProjectApprovalProcess|\/audit$|submitJcdgReview/.test(r.url())) requested = true; });
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
await page.waitForTimeout(2500);
const rstat2 = await page.$eval("#report-status", (e) => e.textContent.trim());
const newDlg = dialogs.slice(dlgBefore);
// ⚠ 回退成功后状态栏会被「正在加载/已显示 检测结果报告单」覆盖，故用弹窗判断回退本身是否成功
check("body:false 视为成功（两步都跑完，不误报「重置失败」）",
  seen.reset.length === 2
    && newDlg.some((m) => /回退成功/.test(m)) && !newDlg.some((m) => /失败/.test(m)),
  "reset次数=" + seen.reset.length + " ／ status=" + rstat2 + " ／ 弹窗=" + JSON.stringify(newDlg).slice(0, 160));

// 8) 「数据报告审核」= 4 步跨账号审批链
//    ① btc-jyg 提交审复核（检测底稿） ② btc-wcy 理化复核通过
//    ③ btc-zhj 报告单理化审核通过     ④ btc-gxz 报告单理化签发通过
const calls = [];      // 审批类请求（按发生顺序）
const tokenCalls = []; // 账号桥取令牌的账号（按发生顺序）
const auditReqBodies = []; // 通过类请求的 body（解析后对象，附 __url）
const affixHits = [];  // getAffix 请求（按发生顺序，含 jcdg/jcbg 标识）——用于验证提交后「双流程」等待
page.on("request", (r) => {
  const u = r.url();
  if (/submitJcdgReview|reportFile\/submitJcbgReview|projectJcdgLastAffixNew\/audit$|projectJcbgLastAffixNew\/audit$/.test(u))
    calls.push(u.replace(/^https?:\/\/[^/]+/, "").split("?")[0]);
  const m = /\/token\/([^/?#]+)/.exec(u);
  if (m) tokenCalls.push(decodeURIComponent(m[1]));
  if (/projectJcdgLastAffixNew\/getAffix/.test(u)) affixHits.push("jcdg");
  else if (/projectJcbgLastAffix\/getAffix/.test(u)) affixHits.push("jcbg");
  if (r.method() === "POST" && /projectJcdgLastAffixNew\/audit$|projectJcbgLastAffixNew\/audit$/.test(u)) {
    try { const body = JSON.parse(r.postData() || "null"); if (body && typeof body === "object") body.__url = u.replace(/^https?:\/\/[^/]+/, "").split("?")[0]; auditReqBodies.push(body); } catch (e) {}
  }
});
const dlgBefore2 = dialogs.length;
// 缩短轮询等待时间：mock 固定响应不会变，避免等满 30s
await page.evaluate(() => { window.__AUDIT_POLL_OPTS__ = { maxWaitMs: 400, intervalMs: 100 }; });
// ★ 审核成功后会自动加载「审批完结检测结果报告单」PDF：
//   mock 报告 PDF 路径的下载请求返回一个最小 PDF；劫持 window.open 以确认**未被调用**（内嵌而非新开标签页）
await page.route("**/src/ohit/**", (route) =>
  route.fulfill({ status: 200, contentType: "application/pdf", body: "%PDF-1.4\n%mock\n" }));
await page.evaluate(() => {
  window.__openedUrls = [];
  window.open = (u) => { window.__openedUrls.push(String(u)); return null; };
});
await page.click("#report-audit");
await page.waitForTimeout(8000);
const rstat3 = await page.$eval("#report-status", (e) => e.textContent.trim());
const newDlg2 = dialogs.slice(dlgBefore2);
const expectCalls = [
  "/api/projectJcdgLastAffixNew/submitJcdgReview",
  "/api/reportFile/submitJcbgReview",
  "/api/projectJcdgLastAffixNew/audit",
  "/api/projectJcbgLastAffixNew/audit",
  "/api/projectJcbgLastAffixNew/audit",
];
check("审核按 4 步链顺序执行（①同时提交底稿+报告单 → ②jcdg理化复核 → ③jcbg理化审核 → ④jcbg理化签发）",
  JSON.stringify(calls) === JSON.stringify(expectCalls),
  JSON.stringify(calls));
// ★ 关键：第①步必须**先提交检测底稿、再提交检测报告单**（独立接口 /api/reportFile/submitJcbgReview），
//   否则 jcbg 一直没有待审记录 → 后续 jcbg 审批必然报「该流程已无待审的检测记录」。
const idxJcdgSubmit = calls.indexOf("/api/projectJcdgLastAffixNew/submitJcdgReview");
const idxJcbgSubmit = calls.indexOf("/api/reportFile/submitJcbgReview");
const idxFirstJcbgAudit = calls.indexOf("/api/projectJcbgLastAffixNew/audit");
check("第①步补交了「检测报告单」（独立接口 reportFile/submitJcbgReview），且在任何 jcbg 审批之前",
  idxJcdgSubmit === 0 && idxJcbgSubmit === 1 && idxFirstJcbgAudit > idxJcbgSubmit,
  "idxJcdgSubmit=" + idxJcdgSubmit + " idxJcbgSubmit=" + idxJcbgSubmit + " idxFirstJcbgAudit=" + idxFirstJcbgAudit);
check("4 步分别用各自账号取令牌（btc-jyg → btc-wcy → btc-zhj → btc-gxz）",
  JSON.stringify(tokenCalls.slice(-4)) === JSON.stringify(["btc-jyg", "btc-wcy", "btc-zhj", "btc-gxz"]),
  JSON.stringify(tokenCalls.slice(-4)));
// 通过类请求体：单个对象（非数组），且 jcbg 的 type 为**数字**（平台要求 int，字符串会 400）
const auditBodies = auditReqBodies; // 见上方 request 监听收集
const jcbgBody = auditBodies.find((b) => b && b.__url && /projectJcbgLastAffixNew\/audit$/.test(b.__url));
check("通过类请求体为单个对象（非数组）且 jcbg.type 为数字",
  jcbgBody && !Array.isArray(jcbgBody) && typeof jcbgBody.type === "number",
  jcbgBody ? JSON.stringify(jcbgBody) : "（未捕获到 jcbg 请求体）");
// ★ 提交审复核后必须等「检测底稿 + 检测报告单」**两条** getAffix 就绪再往下走：
//   即在任何 jcbg 审批之前，序列里应已同时出现 jcdg 与 jcbg 的 getAffix 读取（第一步的双流程等待）。
const firstJcbgAuditIdx = calls.indexOf("/api/projectJcbgLastAffixNew/audit");
const hitsBeforeFirstJcbgAudit = affixHits.slice(0, affixHits.length); // affix 序列（与 calls 交错，独立采集）
const hasJcdgBefore = affixHits.indexOf("jcdg") >= 0;
const hasJcbgBefore = affixHits.indexOf("jcbg") >= 0;
check("提交后等待「检测底稿 + 检测报告单」两条流程都就绪（双流程 getAffix 均已读取）",
  hasJcdgBefore && hasJcbgBefore && firstJcbgAuditIdx >= 0,
  "affix序列=" + JSON.stringify(hitsBeforeFirstJcbgAudit.slice(0, 6)));
// 审核链完成的标志：完成弹窗出现 且 无失败弹窗（状态栏随后会被「已显示报告 PDF」覆盖，故改用弹窗判断）
check("审核链全部完成且无失败弹窗",
  newDlg2.some((m) => /数据报告审核完成/.test(m)) && !newDlg2.some((m) => /失败/.test(m)),
  "status=" + rstat3 + " ／ 弹窗=" + JSON.stringify(newDlg2).slice(0, 140));

// 9) 审核成功后：应**自动**在页面下方内嵌 iframe 显示该项目的「审批完结检测结果报告单」PDF
//    （原「查看…报告单」按钮已移除；mock PDF 路由与 window.open 劫持已在上面注册）
const openedUrls = await page.evaluate(() => window.__openedUrls || []);
const newDlg3 = newDlg2;  // 审核段的弹窗
const pdfBoxVisible = await page.isVisible("#report-pdf-box");
const pdfIframeSrc = await page.$eval("#report-pdf-frame-wrap iframe", (el) => el.getAttribute("src")).catch(() => "");
const rstatAuditEnd = rstat3;
check("审核成功后自动在页面下方内嵌 iframe 显示 PDF（不新开标签页）",
  pdfBoxVisible && pdfIframeSrc.startsWith("blob:") && openedUrls.length === 0
    && /已在下方显示.*检测结果报告单/.test(rstatAuditEnd) && !newDlg3.some((m) => /失败/.test(m)),
  "boxVisible=" + pdfBoxVisible + " iframeSrc=" + String(pdfIframeSrc).slice(0, 24)
    + " opened=" + JSON.stringify(openedUrls) + " ／ status=" + rstatAuditEnd.slice(0, 70));

// 10) 打开 PDF 后：整个页面（PDF 之外）应可纵向滚动，且能把 PDF 顶部一路滚到浏览器顶部
const scrollInfo = await page.evaluate(() => {
  const box = document.getElementById("report-pdf-box");
  const before = box ? Math.round(box.getBoundingClientRect().top) : null;
  window.scrollTo(0, document.documentElement.scrollHeight);
  const after = box ? Math.round(box.getBoundingClientRect().top) : null;
  window.scrollTo(0, 0);
  return {
    hasClass: document.body.classList.contains("pdf-open"),
    bodyOverflowY: getComputedStyle(document.body).overflowY,
    scrollable: document.documentElement.scrollHeight > window.innerHeight,
    topBefore: before,
    topAfter: after,
    docH: document.documentElement.scrollHeight,
    winH: window.innerHeight,
  };
});
check("打开 PDF 后整页可纵向滚动（body.pdf-open + 文档高于视口）",
  scrollInfo.hasClass && scrollInfo.scrollable,
  JSON.stringify(scrollInfo));
// 滚动到底后，PDF 预览框顶部应能到达（或越过）浏览器视口顶部
check("滚动条足够长：滚到底能把 PDF 顶部带到浏览器顶部",
  scrollInfo.topAfter !== null && scrollInfo.topAfter <= 4,
  "滚动前top=" + scrollInfo.topBefore + " 滚动到底top=" + scrollInfo.topAfter);

// 11) 「数据报告回退」成功后：与「数据报告审核」同一套界面/逻辑，自动在下方内嵌显示「检测结果报告单」PDF
//     此时项目已审核完结（jcbg isOver=1），回退应能取到报告单 PDF 并内嵌显示
await page.evaluate(() => {
  // 先清掉上一次的预览（模拟回退前状态），并隐藏已完结预览
  if (window.SamplingUpload && window.SamplingUpload.hideReportPdf) window.SamplingUpload.hideReportPdf();
});
const dlgBeforeReset = dialogs.length;
const openedBeforeReset = (await page.evaluate(() => window.__openedUrls || [])).length;
await page.click("#report-reset");
await page.waitForTimeout(2500);
const rstatReset = await page.$eval("#report-status", (e) => e.textContent.trim());
const resetDlg = dialogs.slice(dlgBeforeReset);
const resetBoxVisible = await page.isVisible("#report-pdf-box");
const resetIframeSrc = await page.$eval("#report-pdf-frame-wrap iframe", (el) => el.getAttribute("src")).catch(() => "");
const openedAfterReset = (await page.evaluate(() => window.__openedUrls || [])).length;
check("回退成功后自动在下方内嵌 iframe 显示检测结果报告单（界面/逻辑同审核）",
  resetBoxVisible && resetIframeSrc.startsWith("blob:")
    && openedAfterReset === openedBeforeReset
    && /已在下方显示回退后检测结果报告单/.test(rstatReset)
    && resetDlg.some((m) => /回退成功/.test(m)) && !resetDlg.some((m) => /失败/.test(m)),
  "boxVisible=" + resetBoxVisible + " iframeSrc=" + String(resetIframeSrc).slice(0, 24)
    + " openedDiff=" + (openedAfterReset - openedBeforeReset) + " ／ status=" + rstatReset.slice(0, 70));

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
