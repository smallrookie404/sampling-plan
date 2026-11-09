// 回归测试：⑭ 检评报告「后处理完毕后全文统一字体 仿宋_GB2312」（用户 2026-11-09）。
//
// 需求边界：
//   · 位置：patchPjbgDocx 管线 ①~⑬ 全部跑完之后、重打包之前（新建/改写的段落也要统一）；
//   · 范围：word/document.xml（正文 / 表格 / 标题 / 段落标记）+ word/numbering.xml
//     （1.1/1.2 自动编号的「编号字体」，用户 2026-11-09 追加要求）；
//     页眉页脚、styles.xml 一律不动（用户明确「页眉页脚不用管」）；
//   · 只动字体名：字号 / 加粗 / 颜色 / 行距 / 缩进 / hint 等其它格式原样保留。
//
// 覆盖：
//   A. 纯函数 forceFontFangSong：各种 rPr / rFonts / run 形态、theme 属性、rStyle 顺序、幂等、
//      异常输入、numbering.xml 编号 rPr
//   B. 端到端 patchPjbgDocx：快速重打包 + JSZip 回退两条路径，校验 document.xml 生效、
//      页眉页脚 / styles.xml 逐字节不变
//   D. 1.1/1.2 自动编号：⑪ 新建 numbering.xml / 模板自带 numbering.xml 两种场景 × 两条重打包路径，
//      校验编号 rFonts 四槽统一、编号字号与 lvl 结构保留、页眉页脚仍不变
//
// 用法：node tests/pjbg.font.test.mjs
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
const fakeEl = { style: {}, classList: { add: noop, remove: noop, toggle: noop }, addEventListener: noop,
  appendChild: noop, removeChild: noop, setAttribute: noop, getAttribute: () => null, hasAttribute: () => false,
  querySelector: () => null, querySelectorAll: () => [], innerHTML: "", textContent: "", disabled: false, click: noop };
const documentStub = { readyState: "complete", getElementById: () => null, querySelector: () => null,
  querySelectorAll: () => [], createElement: () => Object.assign({}, fakeEl), addEventListener: noop,
  body: Object.assign({}, fakeEl), head: Object.assign({}, fakeEl), documentElement: Object.assign({}, fakeEl) };
const LIB = JSON.parse(fs.readFileSync(path.join(APP, "data/library.json"), "utf8"));
function makeBlf(withCompressionStream) {
  const sb = { console, document: documentStub,
    DOMParser: class { parseFromString() { return { getElementsByTagNameNS: () => [] }; } },
    setTimeout, clearTimeout, fetch: () => Promise.reject(new Error("no fetch")),
    URL: { createObjectURL: () => "blob:x", revokeObjectURL: noop }, JSZip,
    TextEncoder, TextDecoder };
  if (withCompressionStream && typeof CompressionStream === "function") sb.CompressionStream = CompressionStream;
  sb.self = sb; sb.window = sb;
  sb.SamplingApp = { librarySnapshot: () => LIB };
  vm.createContext(sb);
  vm.runInContext(blfSrc, sb, { filename: "blf.js" });
  return sb.SamplingBlf;
}
const blf = makeBlf(true);
const F = blf.PJBG_FONT;
const RF = 'w:ascii="' + F + '" w:hAnsi="' + F + '" w:eastAsia="' + F + '" w:cs="' + F + '"';

// vm 沙箱里产生的 ArrayBuffer / Uint8Array 属于另一个 realm，JSZip 的 instanceof 检查会失败；
// 统一转成 Node realm 的 Buffer 再交给 JSZip / 逐字节比较。
function toNodeBytes(x) {
  if (x == null) return x;
  if (Buffer.isBuffer(x)) return x;
  if (typeof x === "string") return Buffer.from(x);
  if (x.byteLength != null) {
    const view = x.buffer ? new Uint8Array(x.buffer, x.byteOffset || 0, x.byteLength) : new Uint8Array(x);
    return Buffer.from(view);
  }
  return Buffer.from(x);
}
function bytesEqual(a, b) {
  const x = toNodeBytes(a), y = toNodeBytes(b);
  if (!x || !y || x.length !== y.length) return false;
  return x.equals(y);
}

// ============ A. 纯函数：各种 XML 形态 ============
console.log("[A] forceFontFangSong 纯函数");
{
  check("钩子暴露且字体名为 仿宋_GB2312", F === "仿宋_GB2312" && typeof blf.forceFontFangSong === "function", String(F));

  const x1 = '<w:r><w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman" w:eastAsia="宋体" w:hint="eastAsia"/><w:sz w:val="28"/></w:rPr><w:t>正文</w:t></w:r>';
  const r1 = blf.forceFontFangSong(x1);
  check("已有 rFonts：四个字体槽统一为 仿宋_GB2312", r1.xml.indexOf(RF) >= 0, r1.xml.slice(0, 200));
  check("已有 rFonts：旧字体被清除", r1.xml.indexOf("Times New Roman") < 0 && r1.xml.indexOf("宋体") < 0);
  check("保留 w:hint", r1.xml.indexOf('w:hint="eastAsia"') >= 0);
  check("保留字号 sz 28（其它格式不动）", r1.xml.indexOf('<w:sz w:val="28"/>') >= 0);
  check("正文文字不被破坏", r1.xml.indexOf("<w:t>正文</w:t>") >= 0);
  check("changed > 0", r1.changed > 0, String(r1.changed));

  const x2 = '<w:r><w:rPr><w:rFonts w:asciiTheme="minorHAnsi" w:hAnsiTheme="minorHAnsi" w:eastAsiaTheme="minorEastAsia" w:cstheme="minorBidi"/><w:b/><w:color w:val="FF0000"/><w:spacing w:val="0"/></w:rPr><w:t>标红</w:t></w:r>';
  const r2 = blf.forceFontFangSong(x2);
  check("主题字体属性被移除", !/Theme=|cstheme=/.test(r2.xml), r2.xml.slice(0, 200));
  check("主题字体也统一为 仿宋_GB2312", r2.xml.indexOf(RF) >= 0);
  check("保留加粗 / 红色 / 字距", r2.xml.indexOf("<w:b/>") >= 0 && r2.xml.indexOf('w:val="FF0000"') >= 0 && r2.xml.indexOf('<w:spacing w:val="0"/>') >= 0);

  const x3 = '<w:r><w:rPr><w:rStyle w:val="a5"/><w:b/><w:sz w:val="21"/></w:rPr><w:t>无字体定义</w:t></w:r>';
  const r3 = blf.forceFontFangSong(x3);
  check("无 rFonts 的 rPr：补上 rFonts", r3.xml.indexOf(RF) >= 0, r3.xml.slice(0, 200));
  check("rStyle 仍排在 rFonts 之前（OOXML 顺序）", r3.xml.indexOf('w:rStyle') < r3.xml.indexOf('w:rFonts'));
  check("rFonts 排在 b / sz 之前", r3.xml.indexOf("w:rFonts") < r3.xml.indexOf("<w:b/>"));

  const r4 = blf.forceFontFangSong('<w:r><w:rPr/><w:t>自闭合</w:t></w:r>');
  check("自闭合 <w:rPr/> 展开并补字体", r4.xml.indexOf(RF) >= 0 && r4.xml.indexOf("</w:rPr>") >= 0, r4.xml.slice(0, 200));

  const r5 = blf.forceFontFangSong('<w:p><w:r><w:t>裸 run</w:t></w:r></w:p>');
  check("无 rPr 的 <w:r> 补一个 rPr", r5.xml.indexOf(RF) >= 0, r5.xml.slice(0, 200));
  check("补的 rPr 在 <w:t> 之前（run 首个子元素）", r5.xml.indexOf("<w:rPr>") < r5.xml.indexOf("<w:t>"));
  check("一个 run 只有一个 rPr", (r5.xml.match(/<w:rPr>/g) || []).length === 1);

  const r6 = blf.forceFontFangSong('<w:p><w:pPr><w:rPr><w:rFonts w:eastAsia="黑体"/></w:rPr></w:pPr><w:r><w:rPr><w:rFonts w:eastAsia="宋体"/></w:rPr><w:t>段落标记</w:t></w:r></w:p>');
  check("段落标记 pPr/rPr 也统一", (r6.xml.match(new RegExp('w:eastAsia="' + F + '"', 'g')) || []).length === 2, r6.xml.slice(0, 240));
  check("段落标记旧字体清除", r6.xml.indexOf("黑体") < 0 && r6.xml.indexOf("宋体") < 0);

  const r7 = blf.forceFontFangSong('<w:r><w:rPr><w:rFonts w:ascii="Calibri" w:eastAsia="黑体"></w:rFonts><w:sz w:val="18"/></w:rPr><w:t>非自闭合</w:t></w:r>');
  check("非自闭合 <w:rFonts></w:rFonts> 也统一", r7.xml.indexOf(RF) >= 0 && r7.xml.indexOf("Calibri") < 0, r7.xml.slice(0, 220));

  const r8 = blf.forceFontFangSong('<w:tbl><w:tr><w:tc><w:p><w:r><w:rPr><w:rFonts w:ascii="Arial" w:eastAsia="微软雅黑"/></w:rPr><w:t>表格内</w:t></w:r></w:p></w:tc></w:tr></w:tbl>');
  check("表格内 run 也统一", r8.xml.indexOf(RF) >= 0 && r8.xml.indexOf("微软雅黑") < 0);

  const once = blf.forceFontFangSong(x1).xml;
  const twice = blf.forceFontFangSong(once).xml;
  check("幂等：重复执行输出不变", once === twice);
  check("空串 / 非字符串原样返回", blf.forceFontFangSong("").changed === 0
    && blf.forceFontFangSong(null).xml === null && blf.forceFontFangSong(undefined).changed === 0);

  const r9 = blf.forceFontFangSong('<w:p><w:pPr><w:jc w:val="both"/><w:spacing w:line="320" w:lineRule="exact"/></w:pPr></w:p>');
  check("没有 rPr / run 的段落：不凭空造字体", r9.xml === '<w:p><w:pPr><w:jc w:val="both"/><w:spacing w:line="320" w:lineRule="exact"/></w:pPr></w:p>');

  const nested = '<w:r><w:rPr><w:rFonts w:eastAsia="宋体"/></w:rPr><w:drawing><w:txbxContent><w:p><w:r><w:t>文本框内文字</w:t></w:r></w:p></w:txbxContent></w:drawing></w:r>';
  const r10 = blf.forceFontFangSong(nested);
  check("嵌套文本框：结构不被破坏", (r10.xml.match(/<w:r>/g) || []).length === 2 && r10.xml.indexOf("</w:drawing>") >= 0, r10.xml.slice(0, 240));
  check("嵌套文本框：内外层字体都统一", r10.xml.indexOf("宋体") < 0 && (r10.xml.match(/w:rFonts/g) || []).length >= 2
    && r10.xml.indexOf("文本框内文字") >= 0);

  // 1.1/1.2 自动编号：编号字体由 numbering.xml 的 <w:lvl><w:rPr><w:rFonts> 决定
  const numFragment = '<w:numbering xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
    + '<w:abstractNum w:abstractNumId="90"><w:lvl w:ilvl="0"><w:start w:val="1"/>'
    + '<w:numFmt w:val="decimal"/><w:lvlText w:val="(%1)"/>'
    + '<w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman" w:cs="Times New Roman" w:eastAsia="仿宋_GB2312"/>'
    + '<w:sz w:val="28"/><w:szCs w:val="28"/></w:rPr></w:lvl></w:abstractNum></w:numbering>';
  const r11 = blf.forceFontFangSong(numFragment);
  check("numbering.xml 编号 rFonts 四槽统一", r11.xml.indexOf(RF) >= 0 && r11.xml.indexOf("Times New Roman") < 0, r11.xml.slice(0, 260));
  check("numbering.xml 保留编号字号 / lvl 结构", r11.xml.indexOf('<w:sz w:val="28"/>') >= 0
    && r11.xml.indexOf('<w:lvlText w:val="(%1)"/>') >= 0 && r11.xml.indexOf("</w:lvl>") >= 0);
}

// ============ B. 端到端：只改 document.xml，页眉页脚 / styles.xml 不动 ============
const DOC_HEAD = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
  + '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>';
const DOC_TAIL = '</w:body></w:document>';
const DOC_XML = DOC_HEAD
  + '<w:p><w:pPr><w:rPr><w:rFonts w:ascii="Times New Roman" w:eastAsia="黑体"/></w:rPr></w:pPr>'
  + '<w:r><w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman" w:eastAsia="宋体" w:cs="Times New Roman" w:hint="eastAsia"/><w:b/><w:sz w:val="28"/><w:color w:val="FF0000"/></w:rPr><w:t>检测与评价结论</w:t></w:r></w:p>'
  + '<w:p><w:r><w:t>裸 run 无字体定义</w:t></w:r></w:p>'
  + '<w:tbl><w:tr><w:tc><w:p><w:r><w:rPr><w:rFonts w:ascii="Calibri" w:eastAsia="微软雅黑"/><w:sz w:val="18"/></w:rPr><w:t>表格内文字</w:t></w:r></w:p></w:tc></w:tr></w:tbl>'
  + DOC_TAIL;
const HEADER_XML = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
  + '<w:hdr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:p><w:r><w:rPr><w:rFonts w:ascii="Calibri" w:eastAsia="黑体"/></w:rPr><w:t>第 1 页页眉</w:t></w:r></w:p></w:hdr>';
const FOOTER_XML = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
  + '<w:ftr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:p><w:r><w:rPr><w:rFonts w:ascii="Arial" w:eastAsia="宋体"/></w:rPr><w:t>页脚</w:t></w:r></w:p></w:ftr>';
const STYLES_XML = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
  + '<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="Calibri" w:eastAsia="宋体" w:cs="Times New Roman"/></w:rPr></w:rPrDefault></w:docDefaults></w:styles>';

async function buildDocx() {
  const zip = new JSZip();
  zip.file("[Content_Types].xml", '<?xml version="1.0"?><Types/>');
  zip.file("word/document.xml", DOC_XML);
  zip.file("word/header1.xml", HEADER_XML);
  zip.file("word/footer1.xml", FOOTER_XML);
  zip.file("word/styles.xml", STYLES_XML);
  zip.file("word/_rels/document.xml.rels", '<?xml version="1.0"?><Relationships/>');
  zip.file("docProps/app.xml", "<app/>");
  return await zip.generateAsync({ type: "arraybuffer", compression: "DEFLATE" });
}

for (const withCS of [true, false]) {
  const tag = withCS ? "快速重打包" : "JSZip 回退";
  console.log("[B] 端到端 patchPjbgDocx（" + tag + "）");
  const b = makeBlf(withCS);
  const src = await buildDocx();
  const out = toNodeBytes(await b.patchPjbgDocx(src));
  let z, o;
  try { z = await JSZip.loadAsync(out); o = await JSZip.loadAsync(src); }
  catch (e) { check(tag + "：输出可被 JSZip 解析", false, String(e)); continue; }
  check(tag + "：输出可被 JSZip 解析", true);
  const doc = await z.file("word/document.xml").async("string");
  check(tag + "：document.xml 全文统一为 仿宋_GB2312", doc.indexOf(RF) >= 0);
  check(tag + "：旧字体全部清除", !/Times New Roman|宋体|黑体|微软雅黑|Calibri|Arial/.test(doc), doc.slice(0, 240));
  check(tag + "：裸 run 也补上字体", doc.indexOf("裸 run 无字体定义") >= 0 && (doc.match(/w:rFonts/g) || []).length >= 4);
  check(tag + "：字号 / 加粗 / 红色等格式保留", doc.indexOf('<w:sz w:val="28"/>') >= 0
    && doc.indexOf("<w:b/>") >= 0 && doc.indexOf('w:val="FF0000"') >= 0 && doc.indexOf('<w:sz w:val="18"/>') >= 0);
  check(tag + "：正文文字完整", doc.indexOf("检测与评价结论") >= 0 && doc.indexOf("表格内文字") >= 0);

  const hdr = await z.file("word/header1.xml").async("uint8array");
  const ftr = await z.file("word/footer1.xml").async("uint8array");
  const sty = await z.file("word/styles.xml").async("uint8array");
  const hdr0 = await o.file("word/header1.xml").async("uint8array");
  const ftr0 = await o.file("word/footer1.xml").async("uint8array");
  const sty0 = await o.file("word/styles.xml").async("uint8array");
  check(tag + "：页眉逐字节不变", bytesEqual(hdr, hdr0));
  check(tag + "：页脚逐字节不变", bytesEqual(ftr, ftr0));
  check(tag + "：styles.xml 逐字节不变", bytesEqual(sty, sty0));
  check(tag + "：页眉仍保留原字体（未被误改）", (await z.file("word/header1.xml").async("string")).indexOf("黑体") >= 0);
  const names = (zz) => Object.keys(zz.files).filter((n) => !zz.files[n].dir).sort();
  check(tag + "：部件清单一致", JSON.stringify(names(z)) === JSON.stringify(names(o)));

  const again = toNodeBytes(await b.patchPjbgDocx(toNodeBytes(await b.patchPjbgDocx(src))));
  const z2 = await JSZip.loadAsync(again);
  check(tag + "：重复导出幂等（document.xml 不再变化）",
    (await z2.file("word/document.xml").async("string")) === doc);
}

// ============ D. 1.1/1.2 自动编号字体（word/numbering.xml） ============
// 编号字体不在 document.xml 里，由 numbering.xml 的 <w:lvl><w:rPr><w:rFonts> 决定。
// 两种来源都要覆盖：⑪ 本次新建（在 numParts 里）/ 模板自带（⑪ 跳过，需从 zip 读出后替换）。
const LIST_PPR = '<w:pPr><w:spacing w:line="490" w:lineRule="exact"/><w:ind w:firstLine="420"/>'
  + '<w:jc w:val="both"/><w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman" w:eastAsia="仿宋_GB2312"/>'
  + '<w:sz w:val="28"/></w:rPr></w:pPr>';
const LIST_RPR = '<w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman" w:eastAsia="仿宋_GB2312"/>'
  + '<w:sz w:val="28"/></w:rPr>';
function listPara(items) {
  let s = "";
  items.forEach((it, i) => {
    if (i > 0) s += "<w:br /><w:tab />";
    s += '<w:t xml:space="preserve">' + it + "</w:t>";
  });
  return "<w:p>" + LIST_PPR + "<w:r>" + LIST_RPR + s + "</w:r></w:p>";
}
const NUM_DOC_XML = DOC_HEAD
  + '<w:p><w:r><w:t>1检测与评价依据</w:t></w:r></w:p>'
  + '<w:p><w:r><w:t>1.1法律、法规、规章</w:t></w:r></w:p>'
  + listPara(["(1)《中华人民共和国职业病防治法》；", "(2)《中华人民共和国尘肺病防治条例》；"])
  + '<w:p><w:r><w:t>1.3质量控制</w:t></w:r></w:p>'
  + DOC_TAIL;
const CT_XML = '<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
  + '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml" />'
  + '<Default Extension="xml" ContentType="application/xml" />'
  + '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml" /></Types>';
const RELS_XML = '<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
  + '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml" /></Relationships>';
const EXISTING_NUMBERING = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
  + '<w:numbering xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
  + '<w:abstractNum w:abstractNumId="0"><w:lvl w:ilvl="0"><w:start w:val="1"/>'
  + '<w:numFmt w:val="decimal"/><w:lvlText w:val="(%1)"/>'
  + '<w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman" w:cs="Times New Roman" w:eastAsia="仿宋_GB2312"/>'
  + '<w:sz w:val="28"/></w:rPr></w:lvl></w:abstractNum>'
  + '<w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num></w:numbering>';

async function buildNumDocx(withExisting) {
  const z = new JSZip();
  z.file("[Content_Types].xml", CT_XML);
  z.file("word/_rels/document.xml.rels", RELS_XML);
  z.file("word/document.xml", NUM_DOC_XML);
  z.file("word/header1.xml", HEADER_XML);
  z.file("word/footer1.xml", FOOTER_XML);
  z.file("word/styles.xml", STYLES_XML);
  if (withExisting) z.file("word/numbering.xml", EXISTING_NUMBERING);
  return await z.generateAsync({ type: "arraybuffer", compression: "DEFLATE" });
}

for (const withCS of [true, false]) {
  for (const withExisting of [false, true]) {
    const tag = (withCS ? "快速重打包" : "JSZip 回退") + " / " + (withExisting ? "模板自带 numbering" : "⑪ 新建 numbering");
    console.log("[D] 1.1/1.2 自动编号字体（" + tag + "）");
    const b = makeBlf(withCS);
    const src = await buildNumDocx(withExisting);
    const out = toNodeBytes(await b.patchPjbgDocx(src));
    let z, o;
    try { z = await JSZip.loadAsync(out); o = await JSZip.loadAsync(src); }
    catch (e) { check(tag + "：输出可被 JSZip 解析", false, String(e)); continue; }
    const nf = z.file("word/numbering.xml");
    check(tag + "：存在 word/numbering.xml", !!nf);
    const numXml = nf ? await nf.async("string") : "";
    check(tag + "：编号 rFonts 四槽统一为 仿宋_GB2312", numXml.indexOf(RF) >= 0, numXml.slice(0, 260));
    check(tag + "：编号旧字体 Times New Roman 已清除", numXml.indexOf("Times New Roman") < 0);
    check(tag + "：编号字号 / 编号格式 / lvl 结构保留", numXml.indexOf('<w:sz w:val="28"/>') >= 0
      && numXml.indexOf('<w:lvlText w:val="(%1)"/>') >= 0 && numXml.indexOf("</w:lvl>") >= 0);
    check(tag + "：numId 定义仍在", /w:numId="\d+"/.test(numXml));
    const doc = await z.file("word/document.xml").async("string");
    check(tag + "：document.xml 列表段落字体也统一", doc.indexOf(RF) >= 0 && doc.indexOf("Times New Roman") < 0);
    if (!withExisting) {
      check(tag + "：⑪ 自动编号生效（document 引用 numId 90）", doc.indexOf('w:numId w:val="90"') >= 0);
      check(tag + "：[Content_Types] 补了 numbering override",
        (await z.file("[Content_Types].xml").async("string")).indexOf("/word/numbering.xml") >= 0);
    }
    check(tag + "：页眉逐字节不变", bytesEqual(
      await z.file("word/header1.xml").async("uint8array"), await o.file("word/header1.xml").async("uint8array")));
    check(tag + "：页脚逐字节不变", bytesEqual(
      await z.file("word/footer1.xml").async("uint8array"), await o.file("word/footer1.xml").async("uint8array")));
    check(tag + "：styles.xml 逐字节不变", bytesEqual(
      await z.file("word/styles.xml").async("uint8array"), await o.file("word/styles.xml").async("uint8array")));
  }
}

// ============ C. 大文档性能（6MB 级，仅记录，不做硬断言） ============
console.log("[C] 大文档性能");
{
  const unit = '<w:p><w:r><w:rPr><w:rFonts w:ascii="Times New Roman" w:eastAsia="宋体"/><w:sz w:val="21"/></w:rPr><w:t>检测数据</w:t></w:r></w:p>';
  const big = DOC_HEAD + unit.repeat(60000) + DOC_TAIL; // ≈ 8MB
  const t0 = Date.now();
  const r = blf.forceFontFangSong(big);
  const ms = Date.now() - t0;
  console.log("  info  输入 " + (big.length / 1024 / 1024).toFixed(1) + " MB，处理 " + ms + " ms，改动 " + r.changed + " 处");
  check("大文档处理完成且全部统一", r.xml.indexOf(RF) >= 0 && r.xml.indexOf("Times New Roman") < 0);
  check("大文档耗时可接受（< 3s）", ms < 3000, ms + "ms");
}

console.log("\n结果：" + pass + " 通过 / " + fail + " 失败");
process.exit(fail ? 1 : 0);
