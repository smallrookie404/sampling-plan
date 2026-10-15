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
const seen = { findList: "", jcbg: "", jcbgYl: "", reset: [], getAffixJcbg: "" };
const j = (body) => ({ status: 200, contentType: "application/json", body: JSON.stringify(body) });
await page.route("**/auth/info*", (route) => route.fulfill(j({ user: { userName: "测试账号" } })));
await page.route("**/api/reportData/findList*", (route) => {
  seen.findList = route.request().url();
  return route.fulfill(j({ code: "200", body: { total: 1, records: [
    { id: 6205, code: "BTC26-SZJC0959", belongInspectName: "浙江永贵电器股份有限公司", status: "6" }] } }));
});
// ★ 列表接口 projectJcbgLastAffix = 平台「流程审批 → 检测结果报告单」列表的数据源。
//   它同时被两处使用：
//     a) 回退前的「在已完结项目中定位」（isDel=1）
//     b) 「数据报告状态」查询（先 isDel=0 未完结，再 isDel=1 已完结）
//   所以 mock 必须**按 isDel 分派**，否则状态查询会拿到错误分支。
//   listState 由测试流程推动：
//     · 初始 {isDel:'1', status:'6'} → 项目「已完结」（回退只对已完结项目有效）
//     · 点「回退」成功后 → {isDel:'0', status:'0'} → 变「待理化复核」（回退把它打回初始节点）
//     · 点「审核」成功后 → {isDel:'1', status:'6'} → 重新「已完结」
//   （状态展示断言在步骤 6.1，那时项目已是「已完结」，故断言的是「已完结」徽标。）
const listState = { isDel: "1", status: "6" };
await page.route("**/api/projectJcbgLastAffix*", (route) => {
  const u = route.request().url();
  // getAffix 是同前缀的不同路径，交给后注册的专属路由处理（Playwright 后注册优先），这里只兜列表
  if (/\/getAffix/.test(u)) return route.fallback();
  seen.jcbg = u;
  const isDel = /[?&]isDel=1(&|$)/.test(u) ? "1" : "0";
  const hit = isDel === listState.isDel;
  return route.fulfill(j({ code: "200", body: { total: hit ? 1 : 0, records: hit ? [
    { id: 6205, code: "BTC26-SZJC0959", belongInspectName: "浙江永贵电器股份有限公司",
      status: listState.status, isDel: listState.isDel, jcdgFlag: "1" }] : [] } }));
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
await page.route("**/api/projectJcbgLastAffix/getAffix*", (route) => {
  seen.getAffixJcbg = route.request().url();
  return route.fulfill(j({ code: "200", body: affixState.jcbg == null ? [] : [{
    id: 6198, type: affixState.jcbg, isOver: affixState.jcbgOver ? "1" : "0",
    // 「审批完结检测结果报告单」PDF 路径（签发版）——供「查看审批完结检测结果报告单」用
    jcbgAffix: "/src/ohit/jcbgReport/913310020555095418/jcbg/BTC26-SZJC0959-x/review/BTC26-SZJC0959_检测报告_签发.pdf",
  }] }));
});
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
// 回退成功 → 项目被打回未完结（isDel=0，节点回到 '0' 待理化复核）；供后续「数据报告状态」断言
page.on("request", (r) => {
  if (r.method() === "POST" && /resetProjectApprovalProcess/.test(r.url())) { listState.isDel = "0"; listState.status = "0"; }
});
// 审核链走完（jcbg 签发通过）→ 项目重新变已完结
page.on("request", (r) => {
  if (r.method() === "POST" && /projectJcbgLastAffixNew\/audit$/.test(r.url())) { listState.isDel = "1"; listState.status = "6"; }
});
// ★ 检测结果报告单**本体**接口（平台「检测结果报告单」页用的就是它）：
//   GET /api/jcbgReport/jcbgYl?projectId=&organizationId= → body = 相对路径
//   ⚠ 数据报告回退会清空 jcbg 审批链记录（getAffix 变空），此时报告只能从这里取 —— 回退后预览的来源。
await page.route("**/api/jcbgReport/jcbgYl*", (route) => {
  seen.jcbgYl = route.request().url();
  return route.fulfill(j({ code: "200", message: "success",
    body: "/src/ohit/jcbgReport/913310020555095418/jcbg/BTC26-SZJC0959-x/BTC26-SZJC0959_检测报告.pdf" }));
});
// ★ 报告 PDF 本体下载（路径形如 /src/ohit/**）：返回最小 PDF；
//   内嵌 iframe 用的是 blob URL，故这里只需保证 fetch 成功。
//   ⚠ 必须在 goto 之前注册 —— 「进页签自动浏览当前状态报告单」发生在很早的步骤，晚注册会 404。
await page.route("**/src/ohit/**", (route) =>
  route.fulfill({ status: 200, contentType: "application/pdf", body: "%PDF-1.4\n%mock\n" }));

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

// 6.1) ★ 选中项目后，「数据报告处理」页应显示该项目的**数据报告项目状态**
//   数据源 = GET /api/projectJcbgLastAffix?isDel=<0|1>&code=<编号>（平台「检测结果报告单」列表的数据源）
//   文案 = 平台 statusFn(status, isDel)：isDel='1' → 「已完结」；否则按 status 映射
//   此时 listState 为 {isDel:'1', status:'6'}（项目已审核完结）→ 徽标应为「已完结」。
await page.evaluate(() => { document.getElementById("xcdc-upload").classList.add("hidden"); });
await page.click('nav.tabs .tab[data-tab="report"]');
await page.waitForTimeout(200);
await page.waitForFunction(() => {
  const b = document.getElementById("report-state-badge");
  return b && b.textContent.trim() && b.textContent.trim() !== "查询中…";
}, null, { timeout: 15000 }).catch(() => {});
const stateVisible = await page.isVisible("#report-state");
const stateBadge = await page.$eval("#report-state-badge", (e) => e.textContent.trim()).catch(() => "");
const stateBadgeCls = await page.$eval("#report-state-badge", (e) => e.className).catch(() => "");
const stateDetail = await page.$eval("#report-state-detail", (e) => e.textContent.trim()).catch(() => "");
check("选中项目后显示「数据报告状态」（已完结 → 已完结徽标）",
  stateVisible && stateBadge === "已完结" && /done/.test(stateBadgeCls)
    && stateDetail === "",   // ★ 明细不再暴露 status/isDel 等技术字段
  "badge=" + stateBadge + " cls=" + stateBadgeCls + " detail=" + stateDetail);
// 状态文案映射（纯函数）：已完结优先、各节点、未知回落
const mapCases = await page.evaluate(() => {
  const f = window.SamplingUpload && window.SamplingUpload.projectStatusText;
  if (typeof f !== "function") return null;
  return {
    done: f("6", "1"),        // isDel=1 → 已完结（优先于 status）
    s0: f("0", "0"), s2: f("2", "0"), s6: f("6", "0"), s9: f("9", "0"), s10: f("10", "0"),
    nil: f(null, "0"),        // 无 status → 未开始
    unknown: f("99", "0"),    // 未知节点 → 回落「节点99」
  };
});
check("项目状态映射与平台 statusFn 一致（isDel=1 优先 → 已完结）",
  mapCases && mapCases.done === "已完结" && mapCases.s0 === "待理化复核" && mapCases.s2 === "待理化审核"
    && mapCases.s6 === "待理化签发" && mapCases.s9 === "待第三方审核" && mapCases.s10 === "待签章"
    && mapCases.nil === "未开始" && mapCases.unknown === "节点99",
  JSON.stringify(mapCases));

// 6.1b) ★ 选中项目进入「数据报告处理」页时，应**默认自动弹出当前状态下的检测结果报告单浏览**
//     —— 来源优先取「审批记录的附件」= GET /api/projectJcbgLastAffix/getAffix 的 body[0].jcbgAffix
//        （与人工在平台「流程审批 → 检测结果报告单 → 附件」看到的完全一致；**未完结也照样显示当前状态版**）。
//     这里把 affixState.jcbg 置为未完结的「理化审核(2)」以模拟真实未完结项目（如 1009 已提交审复核），
//     断言：① 自动发出了 getAffix 请求；② 页面下方内嵌 iframe 显示 PDF；③ 状态栏标注「当前状态」+节点名；④ 未弹「尚未审批完结」提示。
seen.getAffixJcbg = "";
affixState.jcbg = "2";
affixState.jcbgOver = false;
await page.click('nav.tabs .tab[data-tab="report"]');
await page.waitForFunction(() => {
  const box = document.getElementById("report-pdf-box");
  return box && !box.classList.contains("hidden");
}, null, { timeout: 15000 }).catch(() => {});
await page.waitForTimeout(600);
const autoBoxVisible = await page.isVisible("#report-pdf-box");
const autoIframeSrc = await page.$eval("#report-pdf-frame-wrap iframe", (el) => el.getAttribute("src")).catch(() => "");
const autoStatus = await page.$eval("#report-status", (e) => e.textContent.trim()).catch(() => "");
const autoPendingDlg = dialogs.some((m) => /尚未审批完结/.test(m));
check("进「数据报告处理」默认自动显示当下流程中的数据报告（取审批记录附件，未完结也显示）",
  autoBoxVisible && autoIframeSrc.startsWith("blob:")
    && /getAffix/.test(seen.getAffixJcbg || "")
    && /已显示该项目当下流程中的数据报告/.test(autoStatus)
    && /节点 理化审核/.test(autoStatus) && /未完结/.test(autoStatus)
    && !autoPendingDlg,
  "box=" + autoBoxVisible + " src=" + String(autoIframeSrc).slice(0, 22)
    + " ／ status=" + autoStatus.slice(0, 90));
// ★ 自动浏览**不应**滚动主页面（用户反馈「进入数据报告浏览时滚动条突然下滑」）：
//   自动浏览（_auto）不调 scrollIntoView；此处页面应仍停在顶部附近、且未发生明显滚动跳变。
const autoScrollY = await page.evaluate(() => Math.round(window.scrollY));
check("自动浏览当前状态报告时不滚动主页面（滚动条不突跳）",
  autoScrollY <= 8,
  "scrollY=" + autoScrollY);
// 复位：后续步骤 8 依赖「jcbg 初始无记录，只有调 submitJcbgReview 才出现待审记录」来锁定第①步修复
affixState.jcbg = null;
affixState.jcbgOver = false;

// 6.2) 选中项目后点「回退」→ 必须**先在「已完结」项目里定位**
//      （GET /api/projectJcbgLastAffix?isDel=1；回退只对已完结项目有效，用 findList 搜出来的会「找不到」）
//      接口 mock 已在 goto 之前注册（见上）。
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
//    ⚠ 上一步回退会连带触发一次 PDF 加载；须等按钮重新可用再点审核。
await page.waitForFunction(() => {
  const b = document.getElementById("report-audit");
  return b && !b.disabled;
}, null, { timeout: 20000 });
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
// ★ 审核成功后会自动加载「审批完结检测结果报告单」PDF（PDF 下载路由已在文件前部统一注册）。
//   劫持 window.open 以确认**未被调用**（内嵌而非新开标签页）
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
// ★ 账号池共享令牌缓存（远程 0c68166）：4 个池内账号的令牌统一从缓存取，
//   冷启动由 warmAccountPool() 经桥补齐、业务调用前经 /auth/info 探针校验新鲜度。
//   故不再"每步一次 /token"，改为断言「4 个池账号的令牌都已在共享缓存中就绪」。
const poolTokens = await page.evaluate(() => window.__POOL_TEST__ && window.__POOL_TEST__.poolTokens ? Object.keys(window.__POOL_TEST__.poolTokens) : null);
const needAccounts = ["btc-gxz", "btc-wcy", "btc-jyg", "btc-zhj"];
check("4 个池内账号的令牌均在共享缓存中就绪（账号池共享令牌缓存）",
  Array.isArray(poolTokens) && needAccounts.every((a) => poolTokens.indexOf(a) !== -1),
  "poolTokens=" + JSON.stringify(poolTokens) + " bridgeCalls=" + JSON.stringify(tokenCalls.slice(0, 8)));
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

// 9) 审核成功后：应**自动**在页面下方内嵌 iframe 显示该项目当下流程中的数据报告 PDF
//    （原「查看…报告单」按钮已移除；mock PDF 路由与 window.open 劫持已在上面注册）
const openedUrls = await page.evaluate(() => window.__openedUrls || []);
const newDlg3 = newDlg2;  // 审核段的弹窗
const pdfBoxVisible = await page.isVisible("#report-pdf-box");
const pdfIframeSrc = await page.$eval("#report-pdf-frame-wrap iframe", (el) => el.getAttribute("src")).catch(() => "");
const rstatAuditEnd = rstat3;
check("审核成功后自动在页面下方内嵌 iframe 显示当下流程中的数据报告（不新开标签页）",
  pdfBoxVisible && pdfIframeSrc.startsWith("blob:") && openedUrls.length === 0
    && /已显示该项目当下流程中的数据报告/.test(rstatAuditEnd) && !newDlg3.some((m) => /失败/.test(m)),
  "boxVisible=" + pdfBoxVisible + " iframeSrc=" + String(pdfIframeSrc).slice(0, 24)
    + " opened=" + JSON.stringify(openedUrls) + " ／ status=" + rstatAuditEnd.slice(0, 70));

// 9.5) ★ 断点续审（planAuditSteps 纯函数）：按两个流程的当下节点快照算出「还需执行哪些步骤」
//   节点语义：jcdg 只有 "0"（理化复核）；jcbg 有 "2"（理化审核）→"6"（理化签发）。
//   AUDIT_STEPS 下标：0=提交 1=jcdg理化复核 2=jcbg理化审核 3=jcbg理化签发。
const planCases = await page.evaluate(() => {
  const f = window.SamplingUpload && window.SamplingUpload.planAuditSteps;
  if (typeof f !== "function") return null;
  const n = (t) => (t === null ? null : { type: t, isOver: "0" });
  return {
    fresh: f({ jcdg: null, jcbg: null }),                        // 未提交 → 全 4 步
    midJcbg6: f({ jcdg: null, jcbg: n("6") }),                   // 底稿已走完，报告单待签发 → 只 3
    midJcdg0: f({ jcdg: n("0"), jcbg: n("2") }),                 // 底稿待复核 + 报告单待审核 → 1,2
    half: f({ jcdg: null, jcbg: n("2") }),                       // 底稿已走完，报告单待审核 → 只 2
    unknownJcdg: f({ jcdg: n("4"), jcbg: n("2") }),              // 底稿在「职卫复核」（链外）→ 中止
    unknownJcbg: f({ jcdg: n("0"), jcbg: n("9") }),              // 报告单在「第三方审核」（链外）→ 中止
  };
});
check("断点续审规划：仅在链路内推进，跳过已完成步骤；链外节点中止",
  planCases
    && JSON.stringify(planCases.fresh.indices) === "[0,1,2,3]" && planCases.fresh.unknowns.length === 0
    && JSON.stringify(planCases.midJcbg6.indices) === "[3]" && planCases.midJcbg6.unknowns.length === 0
    && JSON.stringify(planCases.midJcdg0.indices) === "[1,2]" && planCases.midJcdg0.unknowns.length === 0
    && JSON.stringify(planCases.half.indices) === "[2]" && planCases.half.unknowns.length === 0
    && planCases.unknownJcdg.indices.length === 0 && planCases.unknownJcdg.unknowns.length === 1
    && planCases.unknownJcbg.indices.length === 0 && planCases.unknownJcbg.unknowns.length === 1,
  JSON.stringify(planCases));

// 9.6) ★ 断点续审（端到端）：项目**已在链路中途**（底稿待理化复核 / 报告单待理化签发），
//   点「数据报告审核」应**只执行未完成的步骤**（② jcdg 理化复核 + ④ jcbg 理化签发），
//   **不重复调提交接口、不调已完成节点的审批**。
const dlgBeforeResume = dialogs.length;
const callsResume = [];
page.on("request", (r) => {
  const u = r.url();
  if (/submitJcdgReview|reportFile\/submitJcbgReview|projectJcdgLastAffixNew\/audit$|projectJcbgLastAffixNew\/audit$/.test(u))
    callsResume.push(u.replace(/^https?:\/\/[^/]+/, "").split("?")[0]);
});
// 造出「项目已在链中途」的平台状态：底稿停在理化复核(0)、报告单停在理化签发(6)
affixState.jcdg = "0";
affixState.jcbg = "6";
affixState.jcbgOver = false;
await page.click("#report-audit");
await page.waitForTimeout(6000);
const newDlgResume = dialogs.slice(dlgBeforeResume);
const expectResume = [
  "/api/projectJcdgLastAffixNew/audit",   // ② 底稿 理化复核（仅此一步，且不重复提交）
  "/api/projectJcbgLastAffixNew/audit",   // ④ 报告单 理化签发
];
check("断点续审端到端：链中途只执行未完成步骤（不重复提交、不重审已完成节点）",
  JSON.stringify(callsResume) === JSON.stringify(expectResume),
  JSON.stringify(callsResume));
check("断点续审完成弹窗标注「跳过已完成」且无失败",
  newDlgResume.some((m) => /数据报告审核完成/.test(m) && /跳过/.test(m))
    && !newDlgResume.some((m) => /失败/.test(m)),
  JSON.stringify(newDlgResume).slice(0, 200));
// 复位：后续不受影响
affixState.jcdg = null;
affixState.jcbg = null;
affixState.jcbgOver = false;
await page.waitForFunction(() => {
  const b = document.getElementById("report-audit");
  return b && !b.disabled;
}, null, { timeout: 20000 });

// 10) 打开 PDF 后：整个页面（PDF 之外）应可纵向滚动，且滚动条下拉区域只比 PDF 内嵌窗口略长
//     （用户反馈：「主页面的滚动条下拉区域太长，只要超过 pdf 浏览的内嵌窗口一点就行」）
const scrollInfo = await page.evaluate(() => {
  const box = document.getElementById("report-pdf-box");
  const before = box ? Math.round(box.getBoundingClientRect().top) : null;
  window.scrollTo(0, document.documentElement.scrollHeight);
  const after = box ? Math.round(box.getBoundingClientRect().top) : null;
  const over = box ? Math.round(box.getBoundingClientRect().bottom) : null;
  window.scrollTo(0, 0);
  return {
    hasClass: document.body.classList.contains("pdf-open"),
    bodyOverflowY: getComputedStyle(document.body).overflowY,
    scrollable: document.documentElement.scrollHeight > window.innerHeight,
    maxScroll: document.documentElement.scrollHeight - window.innerHeight,
    topBefore: before,
    topAfter: after,
    // 滚到底时 PDF 框底部相对视口的偏移：正值表示框底部已进入视口（未被滚出），
    // 反过来说 maxScroll 只比「框高」略大 → 下拉区域仅超出预览窗口一点点
    boxBottomAfter: over,
    docH: document.documentElement.scrollHeight,
    winH: window.innerHeight,
  };
});
check("打开 PDF 后整页可纵向滚动（body.pdf-open + 文档高于视口）",
  scrollInfo.hasClass && scrollInfo.scrollable,
  JSON.stringify(scrollInfo));
// 滚到底时 PDF 框底部仍留在视口内（未被完全滚出）→ 下拉区域只比内嵌窗口略长
check("滚动条下拉区域仅略超出 PDF 内嵌窗口（滚到底 PDF 框底仍在视口内，多余空白 < 120px）",
  scrollInfo.boxBottomAfter !== null && scrollInfo.boxBottomAfter > -120 && scrollInfo.maxScroll < scrollInfo.winH,
  "maxScroll=" + scrollInfo.maxScroll + " winH=" + scrollInfo.winH
    + " 滚到底框底top=" + scrollInfo.boxBottomAfter);

// 11) 「数据报告回退」成功后：与「数据报告审核」同一套界面/逻辑，自动在下方内嵌显示「检测结果报告单」PDF
//     ★ 关键：回退会**清空 jcbg 审批链记录**（getAffix 变空），此时报告只能从「检测结果报告单本体」
//       接口 GET /api/jcbgReport/jcbgYl?projectId=&organizationId= 取 —— 这正是「回退后加载失败：
//       该项目的「检测结果报告单」流程无记录」的根因修复点。
//     ⚠ 第 7 步的回退会连带触发一次 PDF 加载；须等按钮重新可用，否则第 11 步点击会命中 disabled。
await page.waitForFunction(() => {
  const b = document.getElementById("report-reset");
  return b && !b.disabled;
}, null, { timeout: 20000 });
await page.evaluate(() => {
  // 先清掉上一次的预览（模拟回退前状态），并隐藏已完结预览
  if (window.SamplingUpload && window.SamplingUpload.hideReportPdf) window.SamplingUpload.hideReportPdf();
});
// 模拟回退效果：清空 jcbg 审批链记录（getAffix → []）
affixState.jcbg = null;
affixState.jcbgOver = false;
seen.jcbgYl = "";
const dlgBeforeReset = dialogs.length;
const openedBeforeReset = (await page.evaluate(() => window.__openedUrls || [])).length;
await page.click("#report-reset");
await page.waitForTimeout(2500);
const rstatReset = await page.$eval("#report-status", (e) => e.textContent.trim());
const resetDlg = dialogs.slice(dlgBeforeReset);
const resetBoxVisible = await page.isVisible("#report-pdf-box");
const resetIframeSrc = await page.$eval("#report-pdf-frame-wrap iframe", (el) => el.getAttribute("src")).catch(() => "");
const openedAfterReset = (await page.evaluate(() => window.__openedUrls || [])).length;
check("回退后改从「检测结果报告单本体」接口取报告（getAffix 已空，仍能加载）",
  /jcbgReport\/jcbgYl\?[^#]*projectId=/.test(seen.jcbgYl || ""),
  seen.jcbgYl || "（未发出 jcbgYl 请求）");
check("回退成功后自动在下方内嵌 iframe 显示当下流程中的数据报告（界面/逻辑同审核）",
  resetBoxVisible && resetIframeSrc.startsWith("blob:")
    && openedAfterReset === openedBeforeReset
    && /已显示该项目当下流程中的数据报告/.test(rstatReset)
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
