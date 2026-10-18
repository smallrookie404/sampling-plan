// 浏览器冒烟：企业信息（在职人数/接害人数）上传链路。
//
// 背景（2026-10-18 上游逆向 + 真机验证）：
//   平台「现场调查 - 企业信息」页签字段 enterpriseNum=在职人数、hazardNum=接害人数；
//   读 GET /api/inspectedUnitInfo?id=<belongInspect>；写 PUT /api/inspectedUnitInfo/edit（整对象回写）。
//
// 覆盖：
//   A. 页面存在 #entOnJobNum / #entHazardNum 输入框（number 类型）
//   B. buildEntInfoPatch：空值跳过 / 非法值报错 / 只填一侧只出该字段 / 数字串转 Number
//   C. applyEntInfoOnUpload 正常路径：GET 读出 → PUT 回写
//      （enterpriseNum/hazardNum 为数字、数组字段按平台口径拼回字符串、id 原样保留、onSaved 回调触发）
//   D. 两个都空 → 返回 '' 且零请求（不碰平台）
//   E. 非法值 / 缺 belongInspect / GET 失败 / GET 空记录 / PUT 失败 → 各自返回可读说明且流程不抛异常
//
// 用法：node tests/entinfo.smoke.mjs
import { chromium } from "playwright-core";
import fs from "node:fs";
import path from "node:path";
import { startServer } from "../server.mjs";
import { findBrowser } from "./_browser.mjs";

const CHROME = findBrowser();
if (!CHROME) { console.error("未找到 Chrome/Edge，跳过该冒烟测试"); process.exit(0); }

const TEST_DIR = path.resolve("data/_test-entinfo");
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
  // 打桩平台请求：/api/inspectedUnitInfo 走预制应答；其余平台请求一律快速 401（避免拖慢测试）
  window.__ent = { calls: [], getMode: "ok", putMode: "ok" };
  const origFetch = window.fetch.bind(window);
  const json = (obj, status) => Promise.resolve(new Response(JSON.stringify(obj), {
    status: status || 200, headers: { "Content-Type": "application/json" },
  }));
  window.fetch = function (input, init) {
    const url = typeof input === "string" ? input : (input && input.url) || "";
    const method = String((init && init.method) || (input && input.method) || "GET").toUpperCase();
    if (url.indexOf("/api/inspectedUnitInfo") >= 0) {
      window.__ent.calls.push({ url: url, method: method, body: init && init.body ? String(init.body) : null });
      if (method === "PUT") {
        if (window.__ent.putMode === "fail") return json({ code: "500", message: "服务异常" });
        return json({ code: "200", message: "修改成功！", body: true });
      }
      if (window.__ent.getMode === "reject") return Promise.reject(new TypeError("network down"));
      if (window.__ent.getMode === "empty") return json({ code: "200", message: "ok", body: { records: [] } });
      return json({ code: "200", message: "ok", body: { records: [{
        id: 1450, name: "测试单位", code: "TESTCODE",
        enterpriseNum: 410, hazardNum: 53, hazardTotalTrain: 0,
        area: ["浙江省", "台州市"], belongIndustry: ["C", "38"], registrant: ["张三"],
        jsfwField: ["fw1", "fw2"],
        inspectedHelp: { newPneumDisease: ["尘肺"], newHxDisease: ["化学物"], meetStandardDetails: ["m1", "m2"] },
      }] } });
    }
    if (url.indexOf("223.93.144.122") >= 0) {
      // 其余平台请求：保持会话可用（/auth/info 返回用户），项目搜索返回固定项目，避免真实联网
      if (url.indexOf("/auth/info") >= 0) {
        return json({ code: "200", user: { id: 49, userName: "王飞", userCode: "btc-wf", organization: { organizationId: 1 } } });
      }
      if (url.indexOf("/api/reportData/findList") >= 0) {
        return json({ code: "200", message: "查询成功！", body: { records: [
          { id: 6205, code: "BTC26-SZJC0959", belongInspect: 1450, belongInspectName: "测试单位" },
        ] } });
      }
      return json({ code: "200", body: { records: [], content: [] } });
    }
    return origFetch(input, init);
  };
});

await page.goto(APP_URL, { waitUntil: "load" });
await page.waitForFunction(() => window.SamplingUpload && window.SamplingUpload.applyEntInfoOnUpload, null, { timeout: 20000 });

let pass = 0, fail = 0;
const check = (cond, name, extra) => {
  if (cond) { pass++; console.log(`  ✔ ${name}`); }
  else { fail++; console.error(`  ✘ ${name}${extra ? " :: " + extra : ""}`); }
};
const resetCalls = () => page.evaluate(() => { window.__ent.calls.length = 0; });
const calls = () => page.evaluate(() => window.__ent.calls);
const setMode = (getMode, putMode) => page.evaluate(([g, p]) => { window.__ent.getMode = g; window.__ent.putMode = p; }, [getMode, putMode]);

// ---- A. DOM ----
const dom = await page.evaluate(() => {
  const a = document.getElementById("entOnJobNum");
  const b = document.getElementById("entHazardNum");
  const wrapB = b && b.closest(".tb-field");
  const lastTab = document.querySelector("#survey-tabs .stab:last-child");
  return {
    hasA: !!a, hasB: !!b, typeA: a && a.type, typeB: b && b.type,
    aInMainLeft: !!(a && a.closest("#tab-main .grid-toolbar .left")),
    bInSurveyLeft: !!(b && b.closest("#tab-survey .grid-toolbar .left")),
    bRightOfPpeTab: !!(wrapB && wrapB.previousElementSibling && wrapB.previousElementSibling.id === "survey-tabs"),
    lastTabText: lastTab ? lastTab.textContent.trim() : "",
    aInUpload: !!(a && a.closest("#xcdc-upload")),
    bInUpload: !!(b && b.closest("#xcdc-upload")),
    phA: a ? a.getAttribute("placeholder") : "?",
    phB: b ? b.getAttribute("placeholder") : "?",
  };
});
check(dom.hasA && dom.hasB, "A1 页面存在在职人数/接害人数输入框", JSON.stringify(dom));
check(dom.typeA === "number" && dom.typeB === "number", "A2 两个输入框均为 number 类型", JSON.stringify(dom));
check(dom.aInMainLeft, "A3 在职人数位于「测点布局调查」工具栏左侧（录入区同一行）", JSON.stringify(dom));
check(dom.bInSurveyLeft && dom.bRightOfPpeTab && dom.lastTabText === "个体防护调查", "A4 接害人数紧跟在「个体防护调查」子页签右侧", JSON.stringify(dom));
check(!dom.aInUpload && !dom.bInUpload, "A5 两个输入框已移出上传遮罩（在上传视图外填写）", JSON.stringify(dom));
check(!dom.phA && !dom.phB, "A8 空值时不显示占位横线（无 placeholder）", JSON.stringify({ a: dom.phA, b: dom.phB }));
const spin = await page.evaluate(() => {
  const el = document.getElementById("entOnJobNum");
  const cs = getComputedStyle(el, "::-webkit-inner-spin-button");
  return { appearance: cs.appearance || cs.webkitAppearance, cssText: Array.from(document.styleSheets).map((s) => {
    try { return Array.from(s.cssRules).map((r) => r.cssText).join("\n"); } catch (e) { return ""; }
  }).join("\n").indexOf("-webkit-inner-spin-button") >= 0 };
});
check(spin.appearance === "none" || spin.cssText, "A9 数字框右侧上下箭头已隐藏", JSON.stringify(spin));

// A6/A7 真实键入：确认表格的全局键盘处理不会拦截工具栏输入框
await page.click("#entOnJobNum");
await page.fill("#entOnJobNum", "");
await page.type("#entOnJobNum", "410");
const typedMain = await page.evaluate(() => document.getElementById("entOnJobNum").value);
check(typedMain === "410", "A6 在职人数可正常键入（不被表格键盘拦截）", typedMain);
await page.click('.tab[data-tab="survey"]');
await page.click("#entHazardNum");
await page.fill("#entHazardNum", "");
await page.type("#entHazardNum", "53");
const typedSurvey = await page.evaluate(() => document.getElementById("entHazardNum").value);
check(typedSurvey === "53", "A7 接害人数可正常键入（不被表格键盘拦截）", typedSurvey);
await page.click('.tab[data-tab="main"]');

// ---- B. buildEntInfoPatch 纯函数 ----
const patchCases = await page.evaluate(() => {
  const f = window.SamplingUpload.buildEntInfoPatch;
  return {
    empty: f({ onJobNum: "", hazardNum: "  " }),
    blank: f({}),
    bad: f({ onJobNum: "12x", hazardNum: "" }),
    bad2: f({ onJobNum: "", hazardNum: "-3" }),
    one: f({ onJobNum: "500", hazardNum: "" }),
    both: f({ onJobNum: "500", hazardNum: "60" }),
  };
});
check(patchCases.empty.empty === true && patchCases.blank.empty === true, "B1 两个都空 → empty:true（跳过写入）");
check(patchCases.bad.error === "在职人数须为非负整数", "B2 在职人数非数字 → 报错", JSON.stringify(patchCases.bad));
check(patchCases.bad2.error === "接害人数须为非负整数", "B3 接害人数非数字 → 报错", JSON.stringify(patchCases.bad2));
check(patchCases.one.patch.enterpriseNum === 500 && patchCases.one.patch.hazardNum === undefined, "B4 只填在职人数 → 只出 enterpriseNum 且为数字", JSON.stringify(patchCases.one));
check(patchCases.both.patch.enterpriseNum === 500 && patchCases.both.patch.hazardNum === 60, "B5 两个都填 → 两个字段均为数字", JSON.stringify(patchCases.both));

// ---- C. 正常路径 ----
await resetCalls();
const c1 = await page.evaluate(async () => {
  let saved = null;
  const msg = await window.SamplingUpload.applyEntInfoOnUpload(
    "tok", "1", { id: 6205, belongInspect: 1450 }, { onJobNum: "500", hazardNum: "60" },
    (patch) => { saved = patch; }
  );
  return { msg: msg, saved: saved };
});
check(c1.msg.indexOf("企业信息已更新") === 0 && c1.msg.indexOf("500") >= 0 && c1.msg.indexOf("60") >= 0, "C1 返回成功说明（含写入值）", c1.msg);
check(c1.saved && c1.saved.enterpriseNum === 500 && c1.saved.hazardNum === 60, "C2 onSaved 回调收到补丁", JSON.stringify(c1.saved));
const cCalls = await calls();
check(cCalls.length === 2 && cCalls[0].method === "GET" && cCalls[1].method === "PUT", "C3 先 GET 后 PUT 各一次", JSON.stringify(cCalls.map(c => c.method)));
check(/id=1450/.test(cCalls[0].url), "C4 GET 带 id=<belongInspect>", cCalls[0].url);
const putBody = JSON.parse(cCalls[1].body || "{}");
check(putBody.enterpriseNum === 500 && putBody.hazardNum === 60, "C5 PUT 体含两个人数且为数字", JSON.stringify({ e: putBody.enterpriseNum, h: putBody.hazardNum }));
check(putBody.id === 1450 && putBody.name === "测试单位", "C6 PUT 体保留原记录 id/名称（整对象回写）", JSON.stringify({ id: putBody.id, name: putBody.name }));
check(putBody.area === "浙江省/台州市" && putBody.belongIndustry === "C/38" && putBody.registrant === "张三", "C7 数组字段按 '/' 拼回", JSON.stringify({ a: putBody.area, i: putBody.belongIndustry, r: putBody.registrant }));
check(putBody.jsfwField === "fw1,fw2", "C8 jsfwField 按 ',' 拼回", putBody.jsfwField);
check(putBody.inspectedHelp && putBody.inspectedHelp.newPneumDisease === "尘肺" && putBody.inspectedHelp.meetStandardDetails === "m1/m2", "C9 inspectedHelp 数组字段拼回", JSON.stringify(putBody.inspectedHelp));

// ---- D. 两个都空 → 零请求 ----
await resetCalls();
const d1 = await page.evaluate(() => window.SamplingUpload.applyEntInfoOnUpload("tok", "1", { id: 6205, belongInspect: 1450 }, { onJobNum: "", hazardNum: "" }));
const dCalls = await calls();
check(d1 === "" && dCalls.length === 0, "D1 都为空 → 返回 '' 且不请求平台");

// ---- E. 异常分支 ----
await resetCalls();
const e1 = await page.evaluate(() => window.SamplingUpload.applyEntInfoOnUpload("tok", "1", { id: 6205, belongInspect: 1450 }, { onJobNum: "12x", hazardNum: "" }));
check(e1 === "企业信息未写入：在职人数须为非负整数", "E1 非法值 → 说明文案", e1);

const e2 = await page.evaluate(() => window.SamplingUpload.applyEntInfoOnUpload("tok", "1", { id: 6205 }, { onJobNum: "10", hazardNum: "" }));
check(e2.indexOf("belongInspect") >= 0, "E2 缺 belongInspect → 说明文案", e2);

await setMode("reject", "ok");
const e3 = await page.evaluate(() => window.SamplingUpload.applyEntInfoOnUpload("tok", "1", { id: 6205, belongInspect: 1450 }, { onJobNum: "10", hazardNum: "" }));
check(e3.indexOf("企业信息未写入") === 0 && e3.indexOf("network down") >= 0, "E3 GET 网络失败 → 说明文案不抛异常", e3);

await setMode("empty", "ok");
const e4 = await page.evaluate(() => window.SamplingUpload.applyEntInfoOnUpload("tok", "1", { id: 6205, belongInspect: 1450 }, { onJobNum: "10", hazardNum: "" }));
check(e4.indexOf("平台无该受检单位记录") >= 0, "E4 GET 无记录 → 说明文案", e4);

await setMode("ok", "fail");
const e5 = await page.evaluate(() => window.SamplingUpload.applyEntInfoOnUpload("tok", "1", { id: 6205, belongInspect: 1450 }, { onJobNum: "10", hazardNum: "" }));
check(e5.indexOf("企业信息未写入") === 0 && e5.indexOf("服务异常") >= 0, "E5 PUT 业务失败 → 说明文案", e5);

const errCount = pageErrors.length;

// ---- G. 选中项目 → 工具栏输入框预填平台现值 ----
await setMode("ok", "ok");
await page.evaluate(() => {
  const code = document.getElementById("projectCode");
  code.value = "0959";
  code.dispatchEvent(new Event("input", { bubbles: true }));
});
await page.waitForFunction(() => document.querySelectorAll(".suggest-item").length > 0, null, { timeout: 15000 });
await page.evaluate(() => document.querySelector(".suggest-item").click());
await page.waitForFunction(() => document.getElementById("entOnJobNum").value === "410", null, { timeout: 15000 });
const prefill = await page.evaluate(() => ({
  onJob: document.getElementById("entOnJobNum").value,
  hazard: document.getElementById("entHazardNum").value,
  cache: localStorage.getItem("xcdc_entinfo_v1"),
}));
check(prefill.onJob === "410" && prefill.hazard === "53", "G1 选中项目后自动预填平台现值（410/53）", JSON.stringify(prefill));
let cacheObj = {};
try { cacheObj = JSON.parse(prefill.cache || "{}"); } catch (e) {}
check(cacheObj["6205"] && cacheObj["6205"].onJobNum === "410" && cacheObj["6205"].hazardNum === "53", "G2 预填值写入本地缓存 xcdc_entinfo_v1", prefill.cache);

// ---- I. 上传分工：主表格只带在职人数，调查表只带接害人数 ----
const split = await page.evaluate(() => {
  const f = window.SamplingUpload.entPayloadForUpload;
  return {
    main: f("main", { onJobNum: "410", hazardNum: "53" }),
    survey: f("survey", { onJobNum: "410", hazardNum: "53" }),
  };
});
check(split.main.onJobNum === "410" && split.main.hazardNum === "", "I1 主表格上传只带在职人数", JSON.stringify(split.main));
check(split.survey.onJobNum === "" && split.survey.hazardNum === "53", "I2 调查表上传只带接害人数", JSON.stringify(split.survey));

// ---- J. 只写一侧时，另一侧保持平台原值（不被清空） ----
await setMode("ok", "ok");
await resetCalls();
const jMsg1 = await page.evaluate(() => window.SamplingUpload.applyEntInfoOnUpload(
  "tok", "1", { id: 6205, belongInspect: 1450 },
  window.SamplingUpload.entPayloadForUpload("main", { onJobNum: "500", hazardNum: "99" })
));
const jb1 = JSON.parse(((await calls())[1] || {}).body || "{}");
check(jb1.enterpriseNum === 500 && jb1.hazardNum === 53, "J1 主表格上传：只改在职人数，接害保持平台原值 53", JSON.stringify({ e: jb1.enterpriseNum, h: jb1.hazardNum }));
check(jMsg1.indexOf("接害 未改") >= 0, "J2 主表格上传提示「接害 未改」", jMsg1);

await resetCalls();
const jMsg2 = await page.evaluate(() => window.SamplingUpload.applyEntInfoOnUpload(
  "tok", "1", { id: 6205, belongInspect: 1450 },
  window.SamplingUpload.entPayloadForUpload("survey", { onJobNum: "999", hazardNum: "60" })
));
const jb2 = JSON.parse(((await calls())[1] || {}).body || "{}");
check(jb2.enterpriseNum === 410 && jb2.hazardNum === 60, "J3 调查表上传：只改接害人数，在职保持平台原值 410", JSON.stringify({ e: jb2.enterpriseNum, h: jb2.hazardNum }));
check(jMsg2.indexOf("在职 未改") >= 0, "J4 调查表上传提示「在职 未改」", jMsg2);

// ---- K. 上传成功后清空「本次写入」的输入框 + 其本地缓存值（2026-10-18 需求）----
// 语义（用户确认）：只清本次真正写进平台的那一侧；另一侧保留（另一个文件还没上传）；
//               缓存也只清那一侧；patch 为空（没写入任何字段）时一律不清。
const seedK = () => page.evaluate(() => {
  document.getElementById("entOnJobNum").value = "410";
  document.getElementById("entHazardNum").value = "53";
  localStorage.setItem("xcdc_entinfo_v1", JSON.stringify({ "6205": { onJobNum: "410", hazardNum: "53" } }));
});
const snapK = () => page.evaluate(() => {
  let c = {};
  try { c = JSON.parse(localStorage.getItem("xcdc_entinfo_v1") || "{}"); } catch (e) {}
  return {
    onJob: document.getElementById("entOnJobNum").value,
    hazard: document.getElementById("entHazardNum").value,
    cache: c["6205"] || {},
  };
});

await seedK();
const k1 = await page.evaluate(() => window.SamplingUpload.clearEntInputsAfterUpload({ enterpriseNum: 500 }, 6205));
const s1 = await snapK();
check(k1 === "在职人数", "K1 主表格上传（只写在职人数）→ 返回清空名称「在职人数」", k1);
check(s1.onJob === "", "K2 在职人数输入框已清空", JSON.stringify(s1));
check(s1.hazard === "53", "K3 接害人数输入框保留（另一个文件还没上传）", JSON.stringify(s1));
check(s1.cache.onJobNum === "" && s1.cache.hazardNum === "53", "K4 缓存只清写入的那一侧", JSON.stringify(s1.cache));

await seedK();
const k2 = await page.evaluate(() => window.SamplingUpload.clearEntInputsAfterUpload({ hazardNum: 60 }, 6205));
const s2 = await snapK();
check(k2 === "接害人数", "K5 调查表上传（只写接害人数）→ 返回「接害人数」", k2);
check(s2.hazard === "" && s2.onJob === "410", "K6 接害人数清空、在职人数保留", JSON.stringify(s2));
check(s2.cache.hazardNum === "" && s2.cache.onJobNum === "410", "K7 缓存只清接害人数", JSON.stringify(s2.cache));

await seedK();
const k3 = await page.evaluate(() => window.SamplingUpload.clearEntInputsAfterUpload({}, 6205));
const s3 = await snapK();
check(k3 === "", "K8 未写入任何字段（patch 为空）→ 不清空", k3);
check(s3.onJob === "410" && s3.hazard === "53", "K9 两个输入框都保持原值", JSON.stringify(s3));
check(s3.cache.onJobNum === "410" && s3.cache.hazardNum === "53", "K10 缓存不变", JSON.stringify(s3.cache));

// K11：真实链路 —— applyEntInfoOnUpload 写入成功后，onSaved 拿到的 patch 正好是「写入的那一侧」，
//      即 runUploadFlow 里用同一个 patch 调 clearEntInputsAfterUpload 一定只清对应输入框。
await setMode("ok", "ok");
await resetCalls();
await seedK();
const k11 = await page.evaluate(async () => {
  let saved = null;
  const msg = await window.SamplingUpload.applyEntInfoOnUpload(
    "tok", "1", { id: 6205, belongInspect: 1450 },
    window.SamplingUpload.entPayloadForUpload("main", { onJobNum: "410", hazardNum: "53" }),
    (patch) => { saved = patch; }
  );
  const cleared = saved ? window.SamplingUpload.clearEntInputsAfterUpload(saved, 6205) : "";
  return {
    msg: msg,
    saved: saved,
    cleared: cleared,
    onJob: document.getElementById("entOnJobNum").value,
    hazard: document.getElementById("entHazardNum").value,
  };
});
check(k11.saved && k11.saved.enterpriseNum === 410 && k11.saved.hazardNum === undefined,
  "K11 主表格上传：onSaved 只收到 enterpriseNum（不含 hazardNum）", JSON.stringify(k11.saved));
check(k11.cleared === "在职人数" && k11.onJob === "" && k11.hazard === "53",
  "K12 主表格上传后只清空在职人数", JSON.stringify(k11));
check(k11.msg.indexOf("在职 410 人") >= 0, "K13 提示文案仍报告写入值（不受清空影响）", k11.msg);

check(errCount === 0, "H1 页面无未捕获异常", pageErrors.join(" | "));

console.log(`\nentinfo.smoke: ${pass} 通过, ${fail} 失败`);
try { await browser.close(); } catch {}
server.close();
try { fs.rmSync(TEST_DIR, { recursive: true, force: true }); } catch {}
process.exit(fail > 0 ? 1 : 0);
