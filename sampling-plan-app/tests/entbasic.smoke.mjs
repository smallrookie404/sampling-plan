// 浏览器冒烟：「企业基本信息上传」勾选项（调查表上传时自动补全平台企业信息）。
//
// 需求（2026-10-20 用户原话）：
//   在「数据上传」页签「上传调查表」按钮下方加一个「企业基本信息上传」勾选项；勾选后上传调查表时：
//   ① 职卫管理部门为空 → 填「办公室」；
//   ② 职卫负责人 / 职卫负责人电话 为空 → 抄「联系人」/「联系电话」；
//   ③ 危害风险「一般」→ 职卫兼职人数空则填 1；「严重」→ 职卫专职人数空则填 1。
//
// 覆盖：
//   A. DOM：#chkEntBasicUpload 存在、位于「上传调查表」按钮下方、默认未勾选、勾选状态本地记忆
//   B. buildEntBasicPatch 纯函数：三条规则 / 只补空不覆盖 / 联系人空无法补 / 危害风险未选跳过 /
//      计数 0 视为空（zeroAsEmpty 可关）
//   C. applyEntInfoOnUpload + extraBuilder：合并成**一次** PUT（不连写两次），文案含补全说明
//   D. 输入框都为空但勾选了 → 仍然读一次平台并写补全字段
//   E. 没勾选且输入框都为空 → 零请求
//   F. 端到端：勾选 → 点「上传调查表（6 合一）」→ PUT 体含 4 个补全字段
//   G. 端到端：未勾选 → 点「上传调查表」→ PUT 体不含补全字段
//
// 用法：node tests/entbasic.smoke.mjs
import { chromium } from "playwright-core";
import fs from "node:fs";
import path from "node:path";
import { startServer } from "../server.mjs";
import { findBrowser } from "./_browser.mjs";

const CHROME = findBrowser();
if (!CHROME) { console.error("未找到 Chrome/Edge，跳过该冒烟测试"); process.exit(0); }

const TEST_DIR = path.resolve("data/_test-entbasic");
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
  // 平台打桩：企业信息 GET 返回「有联系人、缺职卫字段、危害风险一般」的记录，PUT 记录请求体
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
      if (window.__ent.getMode === "empty") return json({ code: "200", message: "ok", body: { records: [] } });
      return json({ code: "200", message: "ok", body: { records: [{
        id: 1450, name: "测试单位",
        // 三条规则都要触发：职卫三件套全空、危害风险=一般、兼职人数 0（视为空）
        occupationalOrg: null,
        connectPerson: "联系人甲", connectTel: "13800000000",
        occupationalPerson: "", occupationalPersonTel: null,
        occupationalPartNum: 0, occupationalFullNum: 2,
        hazardRisk: "1",
        enterpriseNum: 410, hazardNum: 53,
        area: ["浙江省", "台州市"], belongIndustry: ["C", "38"], registrant: ["张三"],
        jsfwField: ["fw1", "fw2"],
        inspectedHelp: { newPneumDisease: ["尘肺"], newHxDisease: ["化学物"], meetStandardDetails: ["m1", "m2"] },
      }] } });
    }
    if (url.indexOf("223.93.144.122") >= 0) {
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
await page.waitForFunction(() => window.SamplingUpload && window.SamplingUpload.buildEntBasicPatch, null, { timeout: 20000 });

let pass = 0, fail = 0;
const check = (cond, name, extra) => {
  if (cond) { pass++; console.log(`  ✔ ${name}`); }
  else { fail++; console.error(`  ✘ ${name}${extra ? " :: " + extra : ""}`); }
};
const resetCalls = () => page.evaluate(() => { window.__ent.calls.length = 0; });
const calls = () => page.evaluate(() => window.__ent.calls);
// 上传视图默认带 .hidden（未点「数据上传」时不可见）→ 用脚本设值/点击，绕开可见性限制
const setChk = (v) => page.evaluate((val) => {
  const cb = document.getElementById("chkEntBasicUpload");
  cb.checked = val;
  cb.dispatchEvent(new Event("change", { bubbles: true }));
}, v);
const clickUploadSurvey = () => page.evaluate(() => document.getElementById("btnUploadSurvey").click());

// ---- A. DOM ----
const dom = await page.evaluate(() => {
  const cb = document.getElementById("chkEntBasicUpload");
  const btn = document.getElementById("btnUploadSurvey");
  const hint = document.querySelector(".chk-hint");
  const label = cb ? cb.closest("label") : null;
  return {
    has: !!cb,
    type: cb && cb.type,
    inUpload: !!(cb && cb.closest("#xcdc-upload")),
    afterBtn: !!(cb && btn && (btn.compareDocumentPosition(cb) & Node.DOCUMENT_POSITION_FOLLOWING)),
    labelText: label ? label.textContent.trim() : "",
    hasHint: !!hint,
    checked: cb && cb.checked,
  };
});
check(dom.has && dom.type === "checkbox", "A1 存在「企业基本信息上传」勾选框", JSON.stringify(dom));
check(dom.inUpload, "A2 勾选框在数据上传视图内", JSON.stringify(dom));
check(dom.afterBtn, "A3 勾选框位于「上传调查表」按钮下方", JSON.stringify(dom));
check(dom.labelText.indexOf("企业基本信息上传") >= 0, "A4 勾选框文案为「企业基本信息上传」", dom.labelText);
check(dom.hasHint, "A5 勾选框下方有规则说明");
check(dom.checked === false, "A6 默认未勾选（localStorage 无记忆时）", String(dom.checked));

await setChk(true);
const persisted = await page.evaluate(() => ({
  ls: localStorage.getItem("xcdc_entbasic_v1"),
  api: window.SamplingUpload.entBasicChecked(),
}));
check(persisted.ls === "1" && persisted.api === true, "A7 勾选后写入本地记忆且 entBasicChecked()=true", JSON.stringify(persisted));
await setChk(false);
const persisted2 = await page.evaluate(() => ({
  ls: localStorage.getItem("xcdc_entbasic_v1"),
  api: window.SamplingUpload.entBasicChecked(),
}));
check(persisted2.ls === "0" && persisted2.api === false, "A8 取消勾选后同步本地记忆", JSON.stringify(persisted2));

// 记忆恢复：刷新后仍保持「未勾选」；手动置 1 再刷新 → 勾选
await page.evaluate(() => localStorage.setItem("xcdc_entbasic_v1", "1"));
await page.reload({ waitUntil: "load" });
await page.waitForFunction(() => window.SamplingUpload && window.SamplingUpload.entBasicChecked, null, { timeout: 20000 });
const restored = await page.evaluate(() => document.getElementById("chkEntBasicUpload").checked);
check(restored === true, "A9 刷新后恢复上次勾选状态", String(restored));

// ---- B. buildEntBasicPatch 纯函数 ----
const cases = await page.evaluate(() => {
  const f = window.SamplingUpload.buildEntBasicPatch;
  const base = {
    occupationalOrg: null, connectPerson: "张三", connectTel: "139",
    occupationalPerson: "", occupationalPersonTel: null,
    occupationalPartNum: 0, occupationalFullNum: 0, hazardRisk: "1",
  };
  const pick = (r, opts) => f(r, opts);
  return {
    general: pick(base),
    serious: pick(Object.assign({}, base, { hazardRisk: "0" })),
    filled: pick({ occupationalOrg: "EHS部", connectPerson: "张三", connectTel: "139", occupationalPerson: "李四", occupationalPersonTel: "138", occupationalPartNum: 3, occupationalFullNum: 2, hazardRisk: "0" }),
    noConnect: pick({ occupationalOrg: "办公室", connectPerson: "", connectTel: "", occupationalPerson: "", occupationalPersonTel: "", occupationalPartNum: 1, occupationalFullNum: 1, hazardRisk: "1" }),
    noRisk: pick(Object.assign({}, base, { hazardRisk: "" })),
    zeroKept: pick(Object.assign({}, base, { hazardRisk: "0" }), { zeroAsEmpty: false }),
    empty: pick({}),
  };
});
check(cases.general.patch.occupationalOrg === "办公室", "B1 职卫管理部门空 → 办公室", JSON.stringify(cases.general.patch));
check(cases.general.patch.occupationalPerson === "张三" && cases.general.patch.occupationalPersonTel === "139",
  "B2 职卫负责人/电话空 → 抄联系人/联系电话", JSON.stringify(cases.general.patch));
check(cases.general.patch.occupationalPartNum === 1 && cases.general.patch.occupationalFullNum === undefined,
  "B3 危害风险一般 → 兼职人数空补 1（不动专职人数）", JSON.stringify(cases.general.patch));
check(cases.general.notes.length >= 3 && cases.general.notes.join("").indexOf("办公室") >= 0, "B4 返回补全说明 notes", JSON.stringify(cases.general.notes));
check(cases.serious.patch.occupationalFullNum === 1 && cases.serious.patch.occupationalPartNum === undefined,
  "B5 危害风险严重 → 专职人数空补 1（不动兼职人数）", JSON.stringify(cases.serious.patch));
check(Object.keys(cases.filled.patch).length === 0, "B6 已有值一律不覆盖（patch 为空）", JSON.stringify(cases.filled.patch));
check(cases.noConnect.patch.occupationalPerson === undefined && cases.noConnect.patch.occupationalPersonTel === undefined,
  "B7 联系人/联系电话也空 → 不写入空值", JSON.stringify(cases.noConnect.patch));
check(cases.noConnect.notes.join("").indexOf("无法补全") >= 0, "B8 无法补全时给出说明", JSON.stringify(cases.noConnect.notes));
check(cases.noRisk.patch.occupationalPartNum === undefined && cases.noRisk.patch.occupationalFullNum === undefined,
  "B9 危害风险未选 → 跳过人数补全", JSON.stringify(cases.noRisk.patch));
check(cases.noRisk.notes.join("").indexOf("危害风险未选") >= 0, "B10 危害风险未选有提示", JSON.stringify(cases.noRisk.notes));
check(cases.zeroKept.patch.occupationalFullNum === undefined, "B11 zeroAsEmpty=false 时 0 人不再视为空", JSON.stringify(cases.zeroKept.patch));
check(cases.empty.patch.occupationalOrg === "办公室", "B12 空记录（{}）也能补职卫管理部门", JSON.stringify(cases.empty.patch));

// ---- C. 合并成一次 PUT ----
await resetCalls();
const c1 = await page.evaluate(async () => {
  const msg = await window.SamplingUpload.applyEntInfoOnUpload(
    "tok", "1", { id: 6205, belongInspect: 1450 },
    window.SamplingUpload.entPayloadForUpload("survey", { onJobNum: "", hazardNum: "60" }),
    null,
    (rec) => window.SamplingUpload.buildEntBasicPatch(rec)
  );
  return msg;
});
const cCalls = await calls();
check(cCalls.length === 2 && cCalls[0].method === "GET" && cCalls[1].method === "PUT",
  "C1 勾选时仍只有一次 GET + 一次 PUT（合并写入）", JSON.stringify(cCalls.map((c) => c.method)));
const cBody = JSON.parse((cCalls[1] || {}).body || "{}");
check(cBody.hazardNum === 60, "C2 输入框填的接害人数照常写入", JSON.stringify({ h: cBody.hazardNum }));
check(cBody.occupationalOrg === "办公室" && cBody.occupationalPerson === "联系人甲" &&
  cBody.occupationalPersonTel === "13800000000" && cBody.occupationalPartNum === 1,
  "C3 PUT 体同时含 4 个补全字段", JSON.stringify({
    o: cBody.occupationalOrg, p: cBody.occupationalPerson, t: cBody.occupationalPersonTel, n: cBody.occupationalPartNum,
  }));
check(cBody.occupationalFullNum === 2, "C4 未命中的专职人数保持平台原值 2", String(cBody.occupationalFullNum));
check(cBody.id === 1450 && cBody.name === "测试单位", "C5 整对象回写保留 id/名称", JSON.stringify({ id: cBody.id, n: cBody.name }));
check(c1.indexOf("企业信息已更新") === 0 && c1.indexOf("接害 60 人") >= 0, "C6 成功文案含写入值", c1);
check(c1.indexOf("企业基本信息补全") >= 0 && c1.indexOf("办公室") >= 0, "C7 成功文案含补全说明", c1);

// ---- D. 输入框都为空 + 勾选 → 仍然写补全字段 ----
await resetCalls();
const d1 = await page.evaluate(async () => window.SamplingUpload.applyEntInfoOnUpload(
  "tok", "1", { id: 6205, belongInspect: 1450 }, { onJobNum: "", hazardNum: "" }, null,
  (rec) => window.SamplingUpload.buildEntBasicPatch(rec)
));
const dCalls = await calls();
const dBody = JSON.parse((dCalls[1] || {}).body || "{}");
check(dCalls.length === 2 && dBody.occupationalOrg === "办公室", "D1 输入框都空但勾选 → 仍写补全字段", JSON.stringify(dCalls.map((c) => c.method)));
check(dBody.enterpriseNum === 410 && dBody.hazardNum === 53, "D2 未填的人数保持平台原值", JSON.stringify({ e: dBody.enterpriseNum, h: dBody.hazardNum }));
check(d1.indexOf("在职 未改") >= 0 && d1.indexOf("接害 未改") >= 0, "D3 文案标注「未改」", d1);

// ---- E. 没勾选 + 输入框都空 → 零请求 ----
await resetCalls();
const e1 = await page.evaluate(() => window.SamplingUpload.applyEntInfoOnUpload(
  "tok", "1", { id: 6205, belongInspect: 1450 }, { onJobNum: "", hazardNum: "" }, null, null));
const eCalls = await calls();
check(e1 === "" && eCalls.length === 0, "E1 未勾选且都空 → 不请求平台", JSON.stringify({ msg: e1, n: eCalls.length }));

// 无需补全（平台值都齐）→ 不写
await resetCalls();
const e2 = await page.evaluate(async () => window.SamplingUpload.applyEntInfoOnUpload(
  "tok", "1", { id: 6205, belongInspect: 1450 }, { onJobNum: "", hazardNum: "" }, null,
  () => ({ patch: {}, notes: ["测试：无需补全"] })
));
const e2Calls = await calls();
check(e2.indexOf("未改动") >= 0 && e2Calls.length === 1, "E2 无字段可补 → 只读不写并说明", JSON.stringify({ msg: e2, n: e2Calls.length }));

// ---- F/G. 端到端：点「上传调查表（6 合一）」 ----
const selectProject = async () => {
  await page.evaluate(() => {
    const code = document.getElementById("projectCode");
    code.value = "0959";
    code.dispatchEvent(new Event("input", { bubbles: true }));
  });
  await page.waitForFunction(() => document.querySelectorAll(".suggest-item").length > 0, null, { timeout: 15000 });
  await page.evaluate(() => document.querySelector(".suggest-item").click());
  await page.waitForFunction(() => document.getElementById("entOnJobNum").value === "410", null, { timeout: 15000 });
};
await selectProject();

// G：未勾选
await setChk(false);
await resetCalls();
await clickUploadSurvey();
await page.waitForFunction(() => (document.getElementById("log") || {}).textContent.indexOf("企业信息") >= 0, null, { timeout: 60000 });
const gCalls = (await calls()).filter((c) => c.method === "PUT");
const gBody = JSON.parse((gCalls[0] || {}).body || "{}");
check(gCalls.length === 1, "G1 未勾选 → 调查表上传只写一次企业信息", JSON.stringify(gCalls.map((c) => c.method)));
// 未勾选时整对象回写仍会带上平台原值（occupationalOrg=null / occupationalPerson=""），
// 所以断言的是「没有被补成办公室/联系人」而不是「字段不存在」
check(gBody.occupationalOrg !== "办公室" && gBody.occupationalPerson !== "联系人甲",
  "G2 未勾选 → PUT 体不含补全后的值", JSON.stringify({ o: gBody.occupationalOrg, p: gBody.occupationalPerson }));

// F：勾选（此时接害人数输入框已被上一次上传清空，正好覆盖「只勾选也能写」）
await setChk(true);
await resetCalls();
await clickUploadSurvey();
await page.waitForFunction(() => (document.getElementById("log") || {}).textContent.indexOf("企业基本信息补全") >= 0, null, { timeout: 60000 });
const fCalls = (await calls()).filter((c) => c.method === "PUT");
const fBody = JSON.parse((fCalls[0] || {}).body || "{}");
check(fCalls.length === 1, "F1 勾选 → 仍只写一次企业信息", String(fCalls.length));
check(fBody.occupationalOrg === "办公室" && fBody.occupationalPerson === "联系人甲" &&
  fBody.occupationalPersonTel === "13800000000" && fBody.occupationalPartNum === 1,
  "F2 勾选后 PUT 体含 4 个补全字段", JSON.stringify({
    o: fBody.occupationalOrg, p: fBody.occupationalPerson, t: fBody.occupationalPersonTel, n: fBody.occupationalPartNum,
  }));
const logText = await page.evaluate(() => document.getElementById("log").textContent);
check(logText.indexOf("企业基本信息补全") >= 0 && logText.indexOf("办公室") >= 0, "F3 日志打印补全说明");

check(pageErrors.length === 0, "H1 页面无未捕获异常", pageErrors.join(" | "));

console.log(`\nentbasic.smoke: ${pass} 通过, ${fail} 失败`);
try { await browser.close(); } catch {}
server.close();
try { fs.rmSync(TEST_DIR, { recursive: true, force: true }); } catch {}
process.exit(fail > 0 ? 1 : 0);
