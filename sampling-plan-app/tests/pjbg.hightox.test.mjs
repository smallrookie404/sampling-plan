// 回归测试：检评报告后处理①「存在高毒物品」列取值规整。
//
// 背景（2026-10-21，BTC26-SZJC1095 实测）：
//   「检测结果与评价结论表」里 生产车间(1F) / 打包岗位 / 打包 这一行，平台「检测报告编辑」
//   里显示的是「否」，但平台导出的 docx 里该单元格却是占位横线「-」。
//   该列是二值属性（是/否），不存在「不适用」，故后处理把占位横线也按「否」处理。
//
// 覆盖：
//   A. 0 / 1 → 否 / 是（原有行为，不回归）
//   B. 半角横线「-」→ 否（本次修复，复现 1095 场景）
//   C. 全角「－」/ 破折号「—」→ 否
//   D. 已经是「否」「是」→ 原样不动（不重复处理）
//   E. 空单元格 → 不动（没有 <w:t> 内容，无从判断）
//   F. 表头不含「存在高毒物品」的表 → 完全不动
//   G. 表头被拆成多个 run（「存在高」+「毒物品」）→ 仍能识别
//   H. 表头含软换行（「存在高\n毒物品」）→ 仍能识别
//
// 用法：node tests/pjbg.hightox.test.mjs
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
if (!blf || typeof blf.fixHighToxColumn !== "function") {
  console.error("SamplingBlf.fixHighToxColumn 未导出，测试无法进行");
  process.exit(1);
}

// ---- 最小 XML 夹具 ----
function p(text, runs) {
  const parts = runs || [text];
  const rs = parts.map((t) => '<w:r><w:t xml:space="preserve">' + String(t).replace(/&/g, "&amp;").replace(/</g, "&lt;") + "</w:t></w:r>").join("");
  return "<w:p>" + rs + "</w:p>";
}
function pLines(lines) {
  let s = "";
  lines.forEach((ln, i) => {
    if (i > 0) s += "<w:r><w:br/></w:r>";
    s += '<w:r><w:t xml:space="preserve">' + String(ln).replace(/&/g, "&amp;").replace(/</g, "&lt;") + "</w:t></w:r>";
  });
  return "<w:p>" + s + "</w:p>";
}
function tbl(headers, rows) {
  const hc = headers.map((h) => "<w:tc>" + p(h) + "</w:tc>").join("");
  const rc = rows.map((r) => "<w:tr>" + r.map((c) => "<w:tc>" + p(c) + "</w:tc>").join("") + "</w:tr>").join("");
  return "<w:tbl><w:tr>" + hc + "</w:tr>" + rc + "</w:tbl>";
}
// 取某张表第 col 列各行的文本（用与实现一致的口径：拼接所有 <w:t>）
function colTexts(xml, col) {
  const out = [];
  const rows = xml.match(/<w:tr(?:\s[^>]*)?>[\s\S]*?<\/w:tr>/g) || [];
  for (let i = 1; i < rows.length; i++) {
    const cells = rows[i].match(/<w:tc(?:\s[^>]*)?>[\s\S]*?<\/w:tc>/g) || [];
    if (col >= cells.length) { out.push(null); continue; }
    const ts = cells[col].match(/<w:t(?:\s[^>]*)?>[\s\S]*?<\/w:t>/g) || [];
    out.push(ts.map((t) => t.replace(/<[^>]*>/g, "")).join(""));
  }
  return out;
}

const H = ["序号", "单元/工作场所", "岗位（工种）", "点位", "职业病危害因素", "存在高毒物品"];
const COL = 5;   // 「存在高毒物品」列下标

// ---- A. 0 / 1 → 否 / 是 ----
const xmlA = tbl(H, [
  ["1", "生产车间(1F)", "打包", "打包", "其他粉尘", "0"],
  ["2", "生产车间(1F)", "投料", "投料", "甲苯", "1"],
]);
const rA = blf.fixHighToxColumn(xmlA);
const aCols = colTexts(rA.xml, COL);
check("A1 0 → 否、1 → 是", rA.changed === 2 && aCols[0] === "否" && aCols[1] === "是",
  JSON.stringify({ changed: rA.changed, cols: aCols }));

// ---- B. 半角横线「-」→ 否（BTC26-SZJC1095 实况）----
const xmlB = tbl(H, [
  ["1", "生产车间(1F)", "打包岗位", "打包", "其他粉尘", "-"],
]);
const rB = blf.fixHighToxColumn(xmlB);
check("B1 占位横线「-」→ 否", rB.changed === 1 && colTexts(rB.xml, COL)[0] === "否",
  JSON.stringify({ changed: rB.changed, cols: colTexts(rB.xml, COL) }));

// ---- C. 全角「－」/ 破折号「—」/「–」→ 否 ----
const xmlC = tbl(H, [
  ["1", "a", "b", "c", "d", "－"],
  ["2", "a", "b", "c", "d", "—"],
  ["3", "a", "b", "c", "d", "–"],
]);
const rC = blf.fixHighToxColumn(xmlC);
const cCols = colTexts(rC.xml, COL);
check("C1 全角横线/破折号 → 否", rC.changed === 3 && cCols.every((v) => v === "否"),
  JSON.stringify({ changed: rC.changed, cols: cCols }));

// ---- D. 已是「否」「是」→ 不动 ----
const xmlD = tbl(H, [
  ["1", "a", "b", "c", "d", "否"],
  ["2", "a", "b", "c", "d", "是"],
]);
const rD = blf.fixHighToxColumn(xmlD);
check("D1 已是「否/是」不重复处理", rD.changed === 0 && rD.xml === xmlD,
  "changed=" + rD.changed);

// ---- E. 空单元格 → 不动 ----
const xmlE = tbl(H, [["1", "a", "b", "c", "d", ""]]);
const rE = blf.fixHighToxColumn(xmlE);
check("E1 空单元格不动", rE.changed === 0 && colTexts(rE.xml, COL)[0] === "",
  JSON.stringify({ changed: rE.changed, cols: colTexts(rE.xml, COL) }));

// ---- F. 非目标表 → 完全不动 ----
const xmlF = tbl(["序号", "名称", "数量"], [["1", "x", "-"]]);
const rF = blf.fixHighToxColumn(xmlF);
check("F1 表头不含「存在高毒物品」→ 不动", rF.changed === 0 && rF.xml === xmlF,
  "changed=" + rF.changed);

// ---- G. 表头被拆成多个 run ----
const xmlG = '<w:tbl><w:tr><w:tc>' + p("存在高毒物品", ["存在高", "毒物品"]) + "</w:tc></w:tr>" +
  "<w:tr><w:tc>" + p("-") + "</w:tc></w:tr></w:tbl>";
const rG = blf.fixHighToxColumn(xmlG);
check("G1 表头被拆成多个 run 仍能识别", rG.changed === 1 && colTexts(rG.xml, 0)[0] === "否",
  JSON.stringify({ changed: rG.changed, cols: colTexts(rG.xml, 0) }));

// ---- H. 表头含软换行 ----
const xmlH = '<w:tbl><w:tr><w:tc>' + pLines(["存在高", "毒物品"]) + "</w:tc></w:tr>" +
  "<w:tr><w:tc>" + p("0") + "</w:tc></w:tr></w:tbl>";
const rH = blf.fixHighToxColumn(xmlH);
check("H1 表头含软换行仍能识别", rH.changed === 1 && colTexts(rH.xml, 0)[0] === "否",
  JSON.stringify({ changed: rH.changed, cols: colTexts(rH.xml, 0) }));

console.log("\npjbg.hightox: " + pass + " 通过, " + fail + " 失败");
process.exit(fail ? 1 : 0);
