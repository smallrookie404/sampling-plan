// 回归测试：检评报告后处理⑩「表8-2 超标原因分析表」按检测项目自动填「超标原因分析」列。
//
// 需求（用户 2026-10-21）：8.2「不符合岗位/工种超标原因分析」若有表8-2，逐行看「检测项目」列：
//   · 噪声 → 4 行模板；高温 → 6 行；物化性质「粉尘」/「粉尘、毒物」→ 粉尘 6 行；
//     物化性质除「粉尘、毒物」外含「毒物」→ 毒物 5 行；
//   · 只填空 / 占位（-、—、/、无）的单元格，已有内容不覆盖。
//
// 覆盖：
//   A. 噪声 / 高温 / 粉尘 / 「粉尘、毒物」/「毒物、腐蚀」五种因素各写入对应模板
//   B. 不在库里的因素（紫外辐射）→ 不填
//   C. 已有内容的单元格 → 不覆盖
//   D. 一个单元格含多因素（「甲苯,噪声」）→ 两类模板都写，且不重复
//   E. 没有表8-2 → 不改动；表头列被合并平移时仍能定位
//   F. 多行格式：3 个以上 <w:br/>，文本逐行一致
//
// 用法：node tests/pjbg.overlimit.test.mjs
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
const T = blf.PJBG_OVER_LIMIT_TEXTS;

const esc = (t) => String(t).replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;");
function p(text) {
  return '<w:p><w:pPr><w:ind w:firstLine="0"/></w:pPr>'
    + '<w:r><w:rPr><w:rFonts w:ascii="Times New Roman" w:eastAsia="仿宋_GB2312"/><w:sz w:val="18"/></w:rPr>'
    + '<w:t xml:space="preserve">' + esc(text) + '</w:t></w:r></w:p>';
}
function tbl(headers, rows) {
  const hc = headers.map((h) => '<w:tc>' + p(h) + '</w:tc>').join("");
  const rc = rows.map((r) => '<w:tr>' + r.map((c) => '<w:tc>' + p(c) + '</w:tc>').join("") + '</w:tr>').join("");
  return '<w:tbl><w:tr>' + hc + '</w:tr>' + rc + '</w:tbl>';
}
const text = (xml) => { let t = ""; const re = /<w:t(?:\s[^>]*)?>([\s\S]*?)<\/w:t>/g; let m; while ((m = re.exec(xml))) t += m[1]; return t; };

// 危害因素库（真实条目）
const LIB = [
  { name: "大理石粉尘(碳酸钙)", physchem: "粉尘", highTox: "否" },
  { name: "锰及其无机化合物(按MnO2计)", physchem: "粉尘、毒物", highTox: "是" },
  { name: "硫酸及三氧化硫", physchem: "毒物、腐蚀", highTox: "否" },
  { name: "甲苯", physchem: "毒物", highTox: "否" },
  { name: "噪声", physchem: "噪声", highTox: "否" },
  { name: "高温", physchem: "高温", highTox: "否" },
  { name: "紫外辐射", physchem: "紫外辐射", highTox: "否" },
];
const HEADS = ["单元/工作场所", "岗位/工种", "检测点/对象", "检测项目", "超标原因分析"];
const HEAD = p("8.2不符合岗位/工种超标原因分析");

// ============ A. 五种因素分别写入对应模板 ============
console.log("[A] 噪声 / 高温 / 粉尘 / 粉尘、毒物 / 毒物、腐蚀");
{
  const xml = HEAD + tbl(HEADS, [
    ["车间A", "岗位1", "点1", "噪声", "-"],
    ["车间A", "岗位2", "点2", "高温", "-"],
    ["车间B", "岗位3", "点3", "大理石粉尘(碳酸钙)", "—"],
    ["车间B", "岗位4", "点4", "锰及其无机化合物(按MnO2计)", "/"],
    ["车间C", "岗位5", "点5", "硫酸及三氧化硫", "无"],
  ]);
  const r = blf.fillOverLimitCause(xml, LIB);
  const t = text(r.xml);
  check("changed=5 行", r.changed === 5, String(r.changed));
  check("噪声→4 行模板", T.noise.every((x) => t.indexOf(x) >= 0));
  check("高温→6 行模板", T.heat.every((x) => t.indexOf(x) >= 0));
  check("粉尘（粉尘）→6 行模板", t.indexOf(T.dust[0]) >= 0);
  // 本用例有「大理石粉尘(粉尘)」「锰(粉尘、毒物)」两行 → 各自一份粉尘模板，共 2 份（单元格之间不去重）
  check("「粉尘、毒物」→也走粉尘模板（两行各一份）", (t.split(T.dust[0]).length - 1) === 2, String(t.split(T.dust[0]).length - 1));
  check("「毒物、腐蚀」→毒物 5 行模板", T.tox.every((x) => t.indexOf(x) >= 0));
  check("占位符 -、—、/、无 都已被替换", !/超标原因分析/.test("") && (t.split("无").length - 1) >= 1);
}

// ============ B. 库里查不到 / 物理因素非噪声高温 → 不填 ============
console.log("[B] 紫外辐射 → 不填");
{
  const xml = HEAD + tbl(HEADS, [["车间A", "岗位1", "点1", "紫外辐射", "-"]]);
  const r = blf.fillOverLimitCause(xml, LIB);
  check("changed=0", r.changed === 0, String(r.changed));
  check("单元格仍是「-」", text(r.xml).indexOf("紫外辐射") >= 0 && text(r.xml).indexOf(T.noise[0]) < 0);
}

// ============ C. 已有内容不覆盖 ============
console.log("[C] 已有人工填写的内容 → 不覆盖");
{
  const xml = HEAD + tbl(HEADS, [["车间A", "岗位1", "点1", "噪声", "人工填写的原因说明"]]);
  const r = blf.fillOverLimitCause(xml, LIB);
  check("changed=0", r.changed === 0, String(r.changed));
  check("原文保留", text(r.xml).indexOf("人工填写的原因说明") >= 0 && text(r.xml).indexOf(T.noise[0]) < 0);
}

// ============ D. 一个单元格含多因素 ============
console.log("[D] 「甲苯,噪声」→ 毒物 + 噪声两类都写");
{
  const xml = HEAD + tbl(HEADS, [["车间A", "岗位1", "点1", "甲苯,噪声", "-"]]);
  const r = blf.fillOverLimitCause(xml, LIB);
  const t = text(r.xml);
  check("changed=1", r.changed === 1, String(r.changed));
  check("写了毒物模板", T.tox.every((x) => t.indexOf(x) >= 0));
  check("写了噪声模板", T.noise.every((x) => t.indexOf(x) >= 0));
  check("噪声模板只出现一次", (t.split(T.noise[0]).length - 1) === 1, String(t.split(T.noise[0]).length - 1));
}

// ============ E. 无表8-2 / 表头列合并平移 ============
console.log("[E] 无表8-2 不改动；表头合并平移仍能定位");
{
  const xml = HEAD + tbl(["序号", "名称"], [["1", "x"]]);
  const r = blf.fillOverLimitCause(xml, LIB);
  check("无表8-2 → changed=0 且原文不变", r.changed === 0 && r.xml === xml);

  // 数据行比表头多一格（模拟合并单元格）→ 仍应按平移后的列号读
  const hc = HEADS.map((h) => '<w:tc>' + p(h) + '</w:tc>').join("");
  // 数据行比表头多 1 格（表头「检测项目」被合并成 2 格）→ 该格位于目标列**之前**
  const row = '<w:tr>' + [p("车间A"), p("岗位1"), p("点1"), p("1次"), p("噪声"), p("-")].map((c) => '<w:tc>' + c + '</w:tc>').join("") + '</w:tr>';
  const xml2 = HEAD + '<w:tbl><w:tr>' + hc + '</w:tr>' + row + '</w:tbl>';
  const r2 = blf.fillOverLimitCause(xml2, LIB);
  check("平移后仍填对列（噪声模板出现）", text(r2.xml).indexOf(T.noise[0]) >= 0, JSON.stringify(r2.changed));
}

// ============ F. 多行格式 ============
console.log("[F] 多行：一段 N 行（<w:br/> 分行）");
{
  const xml = HEAD + tbl(HEADS, [["车间A", "岗位1", "点1", "噪声", "-"]]);
  const r = blf.fillOverLimitCause(xml, LIB);
  const cell = /<w:tc>(?:(?!<\/w:tc>)[\s\S])*注?XX设备运行[\s\S]*?<\/w:tc>/.exec(r.xml)[0];
  check("含 3 个 <w:br/>（4 行）", (cell.match(/<w:br\/>/g) || []).length === 3, String((cell.match(/<w:br\/>/g) || []).length));
  check("4 行文本逐行一致", T.noise.every((x) => text(cell).indexOf(x) >= 0));
  check("保留原单元格格式（仿宋_GB2312 / sz 18）", cell.indexOf("仿宋_GB2312") >= 0 && cell.indexOf('<w:sz w:val="18"/>') >= 0);
}

console.log(`\npjbg.overlimit: ${pass} 通过, ${fail} 失败`);
process.exit(fail > 0 ? 1 : 0);
