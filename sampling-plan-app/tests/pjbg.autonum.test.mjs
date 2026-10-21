// 回归测试：检评报告后处理⑪「1.1 / 1.2 章节的列表项改成 Word 自动编号」（2026-10-21）。
//
// 规则：
//   · 1.1（法律、法规、规章）与 1.2（技术规范与标准）里的「(1)(2)(3)…」列表项，原来是
//     「一个段落 + <w:br/> 软换行」里**文字写死**的编号 → 改成 Word 原生自动编号：
//     每条拆成独立段落，段落格式（字体/字号/行距/首行缩进/对齐）沿用原段落，编号「(N)」由 Word 生成；
//   · 每个原段落（= 一个列表）一个独立 numId，各自从 1 开始；
//   · 编号定义写入 word/numbering.xml，并补 [Content_Types].xml override 与 document.xml.rels 关系；
//   · 1.3 之后、表格内的列表**不动**。
//
// 用法：node tests/pjbg.autonum.test.mjs
import fs from "node:fs";
import path from "node:path";
import vm from "node:vm";
import { fileURLToPath } from "node:url";
import JSZip from "jszip";

const here = path.dirname(fileURLToPath(import.meta.url));
const APP = path.resolve(here, "..");
const blfSrc = fs.readFileSync(path.join(APP, "js/blf.js"), "utf8");

let pass = 0, fail = 0;
function check(name, cond, extra) {
  if (cond) { pass++; console.log("  ok  " + name); }
  else { fail++; console.log("  FAIL " + name + (extra ? " :: " + extra : "")); }
}

const noop = () => {};
const fakeEl = () => ({ style: {}, classList: { add: noop, remove: noop, toggle: noop, contains: () => false },
  addEventListener: noop, appendChild: noop, removeChild: noop, setAttribute: noop, getAttribute: () => null,
  hasAttribute: () => false, querySelector: () => null, querySelectorAll: () => [], innerHTML: "", textContent: "",
  disabled: false, click: noop, children: [], parentNode: null, dataset: {} });
function makeSandbox() {
  const sb = {
    console: { log: noop, info: noop, warn: noop, error: noop },
    document: { readyState: "complete", getElementById: fakeEl, querySelector: fakeEl, querySelectorAll: () => [],
      createElement: fakeEl, addEventListener: noop, body: fakeEl(), head: fakeEl(), documentElement: fakeEl() },
    DOMParser: class { parseFromString() { return { getElementsByTagNameNS: () => [] }; } },
    setTimeout, clearTimeout, fetch: () => Promise.reject(new Error("no fetch")),
    URL: { createObjectURL: () => "blob:x", revokeObjectURL: noop },
    JSZip, TextDecoder, TextEncoder, CompressionStream,
  };
  sb.self = sb; sb.window = sb;
  vm.createContext(sb);
  vm.runInContext(blfSrc, sb, { filename: "blf.js" });
  return sb.SamplingBlf;
}
const blf = makeSandbox();

// ---- 夹具 ----
const esc = (s) => String(s).replace(/&/g, "&amp;").replace(/</g, "&lt;");
const PPR = '<w:pPr><w:keepNext w:val="0" /><w:keepLines w:val="0" /><w:pageBreakBefore w:val="0" />'
  + '<w:widowControl w:val="0" /><w:suppressLineNumbers w:val="0" /><w:kinsoku /><w:overflowPunct />'
  + '<w:spacing w:line="490" w:lineRule="exact" /><w:ind w:firstLine="420" w:firstLineChars="150" />'
  + '<w:jc w:val="both" /><w:textAlignment w:val="auto" /><w:rPr><w:rFonts w:ascii="Times New Roman"'
  + ' w:eastAsia="仿宋_GB2312" w:hAnsi="Times New Roman" w:cs="Times New Roman" w:hint="default" />'
  + '<w:kern w:val="2" /><w:sz w:val="28" /><w:szCs w:val="28" /></w:rPr></w:pPr>';
const RPR = '<w:rPr><w:rFonts w:ascii="Times New Roman" w:eastAsia="仿宋_GB2312" w:hAnsi="Times New Roman"'
  + ' w:cs="Times New Roman" w:hint="default" /><w:kern w:val="2" /><w:sz w:val="28" /><w:szCs w:val="28" /></w:rPr>';
const p = (text, pPr) => "<w:p>" + (pPr || "") + '<w:r>' + RPR + '<w:t xml:space="preserve">' + esc(text) + "</w:t></w:r></w:p>";
// 多行列表段落：run 内用 <w:br/><w:tab/> 分行（真实报告形态）
function listPara(items) {
  let s = "";
  items.forEach((it, i) => {
    if (i > 0) s += "<w:br /><w:tab />";
    s += '<w:t xml:space="preserve">' + esc(it) + "</w:t>";
  });
  return "<w:p>" + PPR + "<w:r>" + RPR + s + "</w:r></w:p>";
}
const tblWithList = '<w:tbl><w:tr><w:tc>' + listPara(["(1)表内条目一", "(2)表内条目二"]) + "</w:tc></w:tr></w:tbl>";

function buildXml() {
  return "<w:body>"
    + p("1检测与评价依据")
    + p("1.1法律、法规、规章")
    + listPara(["(1)《中华人民共和国职业病防治法》；", "(2)《中华人民共和国尘肺病防治条例》；", "(3)《使用有毒物品作业场所劳动保护条例》。"])
    + p("1.2技术规范与标准")
    + p("1.2.1现场采样与测量依据")
    + listPara(["(1)《工作场所空气中有害物质监测的采样规范》(GBZ 159-2004)；", "(2)《工作场所物理因素测量第8部分：噪声》(GBZ/T 189.8-2007)。"])
    + p("1.2.2实验室检测依据")
    + listPara(["(1)《工作场所空气中粉尘测定标准 第1部分：总粉尘浓度》(GBZ/T 192.1-2025)。"])
    + p("1.3质量控制")
    + listPara(["(1)1.3 之后的列表不应被改", "(2)第二条"])
    + tblWithList
    + "</w:body>";
}
const allParas = (xml) => xml.match(/<w:p(?:\s[^>]*)?>[\s\S]*?<\/w:p>/g) || [];
const pText = (x) => (x.match(/<w:t(?:\s[^>]*)?>([\s\S]*?)<\/w:t>/g) || [])
  .map((z) => z.replace(/<[^>]+>/g, "")).join("");
const numIdOf = (x) => {
  const i = x.indexOf('<w:numId w:val="');
  if (i < 0) return "";
  return x.slice(i + 16, x.indexOf('"', i + 16));
};

// ============ A. 1.1 列表拆分 + 格式保留 ============
console.log("[A] 1.1 列表：3 条拆成 3 段自动编号，格式沿用");
{
  const xml = buildXml();
  const r = blf.autoNumberItemLists(xml);
  check("改了 3 个列表（1.1 / 1.2.1 / 1.2.2）", r.changed === 3, String(r.changed));
  check("共 6 条", r.items === 6, String(r.items));
  const ps = allParas(r.xml);
  const first = ps.find((x) => pText(x).indexOf("中华人民共和国职业病防治法") >= 0);
  const second = ps.find((x) => pText(x).indexOf("尘肺病防治条例") >= 0);
  check("每条独立成段", !!first && !!second && first !== second);
  check("文字里的「(1)」已去掉", pText(first).indexOf("(1)") < 0, pText(first).slice(0, 30));
  check("编号 1.1 = numId 90", numIdOf(first) === "90", numIdOf(first));
  check("同列表后续条目同 numId", numIdOf(second) === "90", numIdOf(second));
  // 格式沿用
  check("保留行距 490 exact", first.indexOf('<w:spacing w:line="490" w:lineRule="exact" />') >= 0);
  check("保留首行缩进 420", first.indexOf('w:firstLine="420"') >= 0);
  check("保留两端对齐", first.indexOf('<w:jc w:val="both" />') >= 0);
  check("保留字体 仿宋_GB2312 / sz28",
    first.indexOf('w:eastAsia="仿宋_GB2312"') >= 0 && first.indexOf('<w:sz w:val="28" />') >= 0);
  check("numPr 在 widowControl 之后、suppressLineNumbers 之前",
    /<w:widowControl w:val="0" \/><w:numPr>[\s\S]*?<\/w:numPr><w:suppressLineNumbers/.test(first));
  check("不再有 <w:br/>（每条独立段落）", first.indexOf("<w:br") < 0);
}

// ============ B. 每个列表独立编号 ============
console.log("[B] 1.2.1 / 1.2.2 各自独立 numId（都从 1 开始）");
{
  const r = blf.autoNumberItemLists(buildXml());
  const ps = allParas(r.xml);
  const a = ps.find((x) => pText(x).indexOf("GBZ 159-2004") >= 0);
  const b = ps.find((x) => pText(x).indexOf("GBZ/T 192.1-2025") >= 0);
  check("1.2.1 = numId 91", numIdOf(a) === "91", numIdOf(a));
  check("1.2.2 = numId 92", numIdOf(b) === "92", numIdOf(b));
  check("numIds 顺序 90/91/92", JSON.stringify(r.numIds) === "[90,91,92]", JSON.stringify(r.numIds));
}

// ============ C. 1.3 之后不动 ============
console.log("[C] 1.3 之后的列表保持原样");
{
  const r = blf.autoNumberItemLists(buildXml());
  const ps = allParas(r.xml);
  const tail = ps.find((x) => pText(x).indexOf("1.3 之后的列表不应被改") >= 0);
  check("仍是带「(1)」的原段落", !!tail && pText(tail).indexOf("(1)1.3") >= 0, tail && pText(tail).slice(0, 20));
  check("没有 numPr", tail && numIdOf(tail) === "", tail && numIdOf(tail));
}

// ============ D. 表格内列表不动 ============
console.log("[D] 表格内的「(1)」段落不动");
{
  const r = blf.autoNumberItemLists(buildXml());
  const inTbl = allParas(r.xml).filter((x) => pText(x).indexOf("表内条目") >= 0);
  check("表内 2 段仍是原样（同一段落）", inTbl.length === 1 && pText(inTbl[0]).indexOf("(1)表内条目一") >= 0,
    String(inTbl.length));
}

// ============ E. numbering.xml 结构 ============
console.log("[E] buildNumberingXml");
{
  const nx = blf.buildNumberingXml([90, 91, 92]);
  check("含 abstractNum 90", nx.indexOf('<w:abstractNum w:abstractNumId="90">') >= 0);
  check("编号格式「(%1)」", nx.indexOf('<w:lvlText w:val="(%1)"/>') >= 0);
  check("编号后不加空格/制表符（suff=nothing）", nx.indexOf('<w:suff w:val="nothing"/>') >= 0);
  check("3 个 w:num 且都指向 abstractNum 90",
    (nx.match(/<w:num w:numId="\d+">/g) || []).length === 3 && (nx.match(/<w:abstractNumId w:val="90"\/>/g) || []).length === 3);
  // ★ 多个 numId 共用 abstractNum 时，Word 默认接着上一个编号 → 必须 startOverride=1 才重新从 1 开始
  check("每个 num 都带 lvlOverride/startOverride=1（各自重新编号）",
    (nx.match(/<w:lvlOverride w:ilvl="0"><w:startOverride w:val="1"\/><\/w:lvlOverride>/g) || []).length === 3,
    String((nx.match(/<w:startOverride/g) || []).length));
  check("abstractNum 在 num 之前", nx.indexOf('<w:abstractNum') < nx.indexOf('<w:num '));
  check("编号字体与正文一致（Times New Roman / 仿宋_GB2312 / sz28）",
    nx.indexOf('w:ascii="Times New Roman"') >= 0 && nx.indexOf('w:eastAsia="仿宋_GB2312"') >= 0);
}

// ============ F. 真实 0959 样例 ============
const SAMPLE = "C:/Users/54606/OneDrive/Desktop/BTC26-SZJC0959-浙江永贵电器股份有限公司_评价报告_49.docx";
if (fs.existsSync(SAMPLE)) {
  console.log("\n[F] 真实 0959 报告：4 个列表 / 38 项");
  const xml = await (await JSZip.loadAsync(fs.readFileSync(SAMPLE))).file("word/document.xml").async("string");
  const r = blf.autoNumberItemLists(xml);
  console.log("   ->", JSON.stringify({ changed: r.changed, items: r.items, numIds: r.numIds }));
  check("4 个列表", r.changed === 4, String(r.changed));
  check("38 条", r.items === 38, String(r.items));
  check("numIds = 90..93", JSON.stringify(r.numIds) === "[90,91,92,93]", JSON.stringify(r.numIds));
  const ps = allParas(r.xml);
  const first = ps.find((x) => numIdOf(x) === "90");
  check("1.1 首条已是自动编号且无文字序号", !!first && pText(first).indexOf("(1)") !== 0, first && pText(first).slice(0, 24));
} else {
  console.log("\n[F] 跳过：未找到 0959 样例（" + SAMPLE + "）");
}

// ============ G. patchPjbgDocx 端到端：三个部件都写进包里 ============
console.log("\n[G] patchPjbgDocx：numbering.xml / Content_Types / rels 同时写入");
{
  const docXml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    + '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
    + buildXml() + "</w:document>";
  const zip = new JSZip();
  zip.file("[Content_Types].xml", '<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
    + '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml" />'
    + '<Default Extension="xml" ContentType="application/xml" />'
    + '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml" /></Types>');
  zip.file("word/_rels/document.xml.rels", '<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
    + '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml" /></Relationships>');
  zip.file("word/styles.xml", '<?xml version="1.0"?><w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"/>');
  zip.file("word/document.xml", docXml);
  const buf = await zip.generateAsync({ type: "arraybuffer", compression: "DEFLATE" });
  const out = await blf.patchPjbgDocx(buf);
  const zo = await JSZip.loadAsync(out);
  const num = zo.file("word/numbering.xml");
  check("新增 word/numbering.xml", !!num);
  const numXml = num ? await num.async("string") : "";
  check("numbering.xml 含 numId 90/91/92", numXml.indexOf('w:numId="90"') >= 0 && numXml.indexOf('w:numId="92"') >= 0);
  const ct = await zo.file("[Content_Types].xml").async("string");
  check("Content_Types 补了 numbering override", ct.indexOf("/word/numbering.xml") >= 0, ct.slice(-160));
  const rels = await zo.file("word/_rels/document.xml.rels").async("string");
  check("rels 补了 numbering 关系", rels.indexOf("relationships/numbering") >= 0, rels.slice(-200));
  check("原 styles.xml 仍在", !!zo.file("word/styles.xml"));
  const d2 = await zo.file("word/document.xml").async("string");
  const refIds = new Set([...d2.matchAll(/<w:numId w:val="(\d+)"\/>/g)].map((m) => m[1]));
  const defIds = new Set([...numXml.matchAll(/<w:num w:numId="(\d+)">/g)].map((m) => m[1]));
  check("document 引用的 numId 都有定义", [...refIds].every((i) => defIds.has(i)), JSON.stringify([...refIds]));
}

console.log(`\n结果：${pass} 通过 / ${fail} 失败`);
process.exit(fail ? 1 : 0);
