// 回归测试：后处理③「正文措辞替换」——跨 run 替换不能重复插入。
//
// 背景（2026-10-21，测试.docx 段落 #437）：
//   报告里 7.1 标题导出成「7.1物理因素检测结果分析与评价物理因素检测结果分析与评价」（重复两遍）。
//   根因在 spliceParaRange：替换区间 [start,end) **跨越多个 run** 时，原实现让每个相交的 run 都
//   执行 `slice(0,a) + to + slice(b)`，于是 to 被插入 N 次（N = 相交 run 数）。
//   平台原文是「7.1」「物理有害因素」「检测结果分析与评价」三个 run，
//   替换「物理有害因素检测结果分析与评价」→「物理因素检测结果分析与评价」时，后两个 run 各插一次 → 重复。
//
// 覆盖：
//   A. 跨 2 个 run → to 只出现 1 次
//   B. 跨 3 个 run → to 只出现 1 次
//   C. 单 run 内 → 正常替换（不回归）
//   D. 区间在中间（前后都有文字）→ 前后缀都保留，to 只 1 次
//   E. 另两条规则（接触限值）跨 run 也正确
//   F. 未命中的段落原样返回
//
// 用法：node tests/pjbg.textfix.test.mjs
import fs from "node:fs";
import path from "node:path";
import vm from "node:vm";
import { fileURLToPath } from "node:url";

const here = path.dirname(fileURLToPath(import.meta.url));
const APP = path.resolve(here, "..");
const blfSrc = fs.readFileSync(path.join(APP, "js/blf.js"), "utf8");

let pass = 0, fail = 0;
function check(name, cond, extra) {
  if (cond) { pass++; console.log("  ok  " + name); }
  else { fail++; console.log("  FAIL " + name + (extra ? " :: " + extra : "")); }
}

const noop = () => {};
const fakeEl = { style: {}, classList: { add: noop, remove: noop, toggle: noop }, addEventListener: noop,
  appendChild: noop, removeChild: noop, setAttribute: noop, getAttribute: () => null, hasAttribute: () => false,
  querySelector: () => null, querySelectorAll: () => [], innerHTML: "", textContent: "", disabled: false, click: noop };
const documentStub = { readyState: "complete", getElementById: () => null, querySelector: () => null,
  querySelectorAll: () => [], createElement: () => Object.assign({}, fakeEl), addEventListener: noop,
  body: Object.assign({}, fakeEl), head: Object.assign({}, fakeEl), documentElement: Object.assign({}, fakeEl) };
const sandbox = { console, document: documentStub,
  DOMParser: class { parseFromString() { return { getElementsByTagNameNS: () => [] }; } },
  setTimeout, clearTimeout, fetch: () => Promise.reject(new Error("no fetch")),
  URL: { createObjectURL: () => "blob:x", revokeObjectURL: noop } };
sandbox.self = sandbox; sandbox.window = sandbox;
vm.createContext(sandbox);
vm.runInContext(blfSrc, sandbox, { filename: "blf.js" });
const blf = sandbox.SamplingBlf;

const esc = (s) => String(s).replace(/&/g, "&amp;").replace(/</g, "&lt;");
// runs: 字符串数组，每个一个 <w:r><w:t>
function p(runs) {
  return "<w:p>" + runs.map((t) => '<w:r><w:t xml:space="preserve">' + esc(t) + "</w:t></w:r>").join("") + "</w:p>";
}
function textOf(xml) {
  const ts = xml.match(/<w:t(?:\s[^>]*)?>[\s\S]*?<\/w:t>/g) || [];
  return ts.map((t) => t.replace(/<[^>]*>/g, "")).join("");
}
const count = (s, sub) => s.split(sub).length - 1;

const FROM = "物理有害因素检测结果分析与评价";
const TO = "物理因素检测结果分析与评价";

// ---- A. 跨 2 个 run ----
{
  const xml = p(["7.1", "物理有害因素", "检测结果分析与评价"]);
  const r = blf.fixReportText(xml);
  const txt = textOf(r.xml);
  check("A1 跨 2 run：结果为「7.1」+ 新标题", txt === "7.1" + TO, JSON.stringify(txt));
  check("A2 跨 2 run：新标题只出现 1 次", count(txt, TO) === 1, "次数=" + count(txt, TO));
}

// ---- B. 跨 3 个 run ----
{
  const xml = p(["7.1", "物理有害", "因素检测", "结果分析与评价"]);
  const r = blf.fixReportText(xml);
  const txt = textOf(r.xml);
  check("B1 跨 3 run：结果正确", txt === "7.1" + TO, JSON.stringify(txt));
  check("B2 跨 3 run：新标题只出现 1 次", count(txt, TO) === 1, "次数=" + count(txt, TO));
}

// ---- C. 单个 run ----
{
  const xml = p(["7.1" + FROM]);
  const r = blf.fixReportText(xml);
  const txt = textOf(r.xml);
  check("C1 单 run：正常替换", txt === "7.1" + TO, JSON.stringify(txt));
  check("C2 单 run：只出现 1 次", count(txt, TO) === 1);
}

// ---- D. 区间在中间（前后都有文字）----
{
  const xml = p(["前言", "物理有害因素", "检测结果分析与评价", "后记"]);
  const r = blf.fixReportText(xml);
  const txt = textOf(r.xml);
  check("D1 前后缀都保留", txt === "前言" + TO + "后记", JSON.stringify(txt));
  check("D2 新标题只出现 1 次", count(txt, TO) === 1);
}

// ---- E. 另一条规则（接触限值）跨 run ----
{
  const xml = p(["工作场所空气中", "粉尘及化学有害因素", "职业接触限值"]);
  const r = blf.fixReportText(xml);
  const txt = textOf(r.xml);
  const want = "工作场所空气中化学有害因素职业接触限值";
  check("E1 接触限值规则跨 run 正确", txt === want, JSON.stringify(txt));
  check("E2 未重复", count(txt, "工作场所空气中化学有害因素职业接触限值") === 1);
}

// ---- F. 未命中的段落原样返回 ----
{
  const xml = p(["这一段没有任何需要替换的说法。"]);
  const r = blf.fixReportText(xml);
  check("F1 未命中段落原样", r.xml === xml && r.changed === 0, "changed=" + r.changed);
}

console.log("\npjbg.textfix: " + pass + " 通过, " + fail + " 失败");
process.exit(fail ? 1 : 0);
