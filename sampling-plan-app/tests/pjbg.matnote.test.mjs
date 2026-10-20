// 回归测试：检评报告后处理⑨「表3-3 原辅材料情况」下方插入「注1~注4」（整段标红）。
//
// 需求（用户 2026-10-21）：在表3-3 表格下面增加一段（4 行），说明「漆料主要成分」的取值口径，文字标红。
//
// 覆盖：
//   A. 插入位置：紧跟在表3-3 的 </w:tbl> 之后、下一段（3.4.2产品情况）之前
//   B. 文本：4 条注逐字一致，用 <w:br/> 分成 4 行（一段 4 行）
//   C. 标红：新段落每个 run 的 rPr 都带 <w:color w:val="FF0000"/>
//   D. 格式继承：pPr / rPr 取自「表3-3」标题前一段（仿宋_GB2312、sz 21），不是标题段本身
//   E. 幂等：同一文档再跑一次 → changed=0，且只有一份注
//   F. 找不到「表3-3」标题 → 按表头（物料名称 + 主要成分 + 年用量）兜底定位
//   G. 没有原辅材料表 → 不改动
//
// 用法：node tests/pjbg.matnote.test.mjs
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
const NOTES = blf.PJBG_MATERIAL_NOTES;

// ---- 夹具 ----
const esc = (t) => String(t).replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;");
// 正文段（仿宋_GB2312 / sz 21，首行缩进 2 字符）——用于验证格式继承
function bodyP(text) {
  return '<w:p><w:pPr><w:ind w:firstLine="420"/></w:pPr>'
    + '<w:r><w:rPr><w:rFonts w:ascii="仿宋_GB2312" w:hAnsi="仿宋_GB2312" w:eastAsia="仿宋_GB2312" w:hint="eastAsia"/><w:sz w:val="21"/><w:szCs w:val="21"/></w:rPr>'
    + '<w:t xml:space="preserve">' + esc(text) + '</w:t></w:r></w:p>';
}
// 表格标题段（黑体、居中）——不应被当作格式模板
function capP(text) {
  return '<w:p><w:pPr><w:jc w:val="center"/></w:pPr>'
    + '<w:r><w:rPr><w:rFonts w:ascii="黑体" w:eastAsia="黑体"/><w:sz w:val="24"/></w:rPr>'
    + '<w:t>' + esc(text) + '</w:t></w:r></w:p>';
}
function tbl(headers, rows) {
  const hc = headers.map((h) => '<w:tc>' + bodyP(h) + '</w:tc>').join("");
  const rc = rows.map((r) => '<w:tr>' + r.map((c) => '<w:tc>' + bodyP(c) + '</w:tc>').join("") + '</w:tr>').join("");
  return '<w:tbl><w:tr>' + hc + '</w:tr>' + rc + '</w:tbl>';
}
const MAT_TABLE = tbl(["工作场所", "物料名称", "主要成分", "物理状态", "年用量"],
  [["生产车间", "漆料", "甲苯、二甲苯", "液态", "5t"], ["生产车间", "清洗剂", "丙酮", "液态", "1t"]]);
const text = (xml) => { let t = ""; const re = /<w:t(?:\s[^>]*)?>([\s\S]*?)<\/w:t>/g; let m; while ((m = re.exec(xml))) t += m[1]; return t; };
const count = (xml, s) => (text(xml).split(s).length - 1);

// ============ A/B/C/D. 正常插入 ============
console.log("[A-D] 正常插入（位置 / 4 行 / 标红 / 格式继承）");
{
  const xml = bodyP("3.4原辅材料及产品情况")
    + bodyP("3.4.1原辅材料情况")
    + bodyP("用人单位原辅材料情况详见表3-3。")
    + capP("表3-3 原辅材料情况")
    + MAT_TABLE
    + bodyP("3.4.2产品情况");
  const r = blf.addMaterialNote(xml);
  check("changed=1", r.changed === 1, String(r.changed));

  // A. 位置：注段落在 </w:tbl> 之后、3.4.2 之前
  const tblEnd = r.xml.indexOf("</w:tbl>") + "</w:tbl>".length;
  const noteIdx = r.xml.indexOf("注1：本次检测对用人单位漆料");
  const nextIdx = r.xml.indexOf("3.4.2产品情况");
  check("注段落在表格之后", noteIdx > tblEnd);
  check("注段落在 3.4.2 之前", noteIdx < nextIdx);
  const between = r.xml.slice(tblEnd, nextIdx);
  // 表格之后到「3.4.2」文字之前：只有「新插入的注段落 + 3.4.2 那段」两个 <w:p>
  check("表格与 3.4.2 之间只有「注段落 + 3.4.2 段」两段", (between.match(/<w:p>/g) || []).length === 2, between.slice(0, 80));

  // B. 文本：4 条逐字一致，用 3 个 <w:br/> 分 4 行
  NOTES.forEach((n, i) => check("注" + (i + 1) + " 文本逐字一致", text(r.xml).indexOf(n) >= 0, n.slice(0, 30)));
  const notePara = r.xml.slice(r.xml.lastIndexOf("<w:p>", noteIdx), r.xml.indexOf("</w:p>", noteIdx) + 6);
  check("注段落含 3 个 <w:br/>（4 行）", (notePara.match(/<w:br\/>/g) || []).length === 3, String((notePara.match(/<w:br\/>/g) || []).length));
  check("注段落含 4 个 <w:t>", (notePara.match(/<w:t[ >]/g) || []).length === 4, String((notePara.match(/<w:t[ >]/g) || []).length));

  // C. 标红：每个 run 都有 FF0000
  const runs = notePara.match(/<w:r>[\s\S]*?<\/w:r>/g) || [];
  const allRed = runs.length > 0 && runs.every((x) => x.indexOf('w:val="FF0000"') >= 0);
  check("注段落每个 run 都标红（FF0000）", allRed, JSON.stringify(runs.map((x) => (x.match(/<w:color[^>]*\/>/) || [""])[0])));

  // D. 格式继承自「标题前一段」正文（仿宋_GB2312 / sz 21 / 首行缩进），不是黑体标题
  check("字体继承正文（仿宋_GB2312）", notePara.indexOf("仿宋_GB2312") >= 0, notePara.slice(0, 160));
  check("字号继承正文（sz 21）", notePara.indexOf('<w:sz w:val="21"/>') >= 0);
  check("段落缩进继承正文（firstLine 420）", notePara.indexOf('w:firstLine="420"') >= 0);
  check("未误用标题的黑体/24 号", notePara.indexOf("黑体") < 0 && notePara.indexOf('<w:sz w:val="24"/>') < 0);
}

// ============ E. 幂等 ============
console.log("[E] 幂等：重复调用不叠加");
{
  const xml = capP("表3-3 原辅材料情况") + MAT_TABLE + bodyP("3.4.2产品情况");
  const r1 = blf.addMaterialNote(xml);
  const r2 = blf.addMaterialNote(r1.xml);
  check("第二次 changed=0", r2.changed === 0, String(r2.changed));
  check("全文只有一份「注1：」", count(r2.xml, "注1：本次检测对用人单位漆料") === 1, String(count(r2.xml, "注1：")));
}

// ============ F. 无标题 → 按表头兜底 ============
console.log("[F] 找不到「表3-3」标题 → 按表头兜底");
{
  const xml = bodyP("3.4.1原辅材料情况") + MAT_TABLE + bodyP("3.4.2产品情况");
  const r = blf.addMaterialNote(xml);
  check("兜底也能插入", r.changed === 1 && text(r.xml).indexOf(NOTES[0]) >= 0, String(r.changed));
  const tblEnd = r.xml.indexOf("</w:tbl>") + "</w:tbl>".length;
  check("兜底插入位置在表后", r.xml.indexOf("注1：本次检测") > tblEnd);
}

// ============ G. 没有原辅材料表 → 不动 ============
console.log("[G] 没有原辅材料表 → 不改动");
{
  const xml = bodyP("3.4.1原辅材料情况") + tbl(["序号", "名称"], [["1", "x"]]) + bodyP("3.4.2产品情况");
  const r = blf.addMaterialNote(xml);
  check("changed=0 且原文不变", r.changed === 0 && r.xml === xml);
}

console.log(`\npjbg.matnote: ${pass} 通过, ${fail} 失败`);
process.exit(fail > 0 ? 1 : 0);
