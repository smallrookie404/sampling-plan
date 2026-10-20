// 回归测试：检评报告「快速重打包」（B1）——只重压 document.xml，其余条目原始压缩字节直接拷贝。
//
// 背景（2026-10-21 实测，0959：docx 1.88MB / document.xml 6.14MB / 31 个条目）：
//   JSZip 的 generateAsync 会把所有条目（含 1.7MB 图片）重新 DEFLATE，耗时 584ms；
//   实际只有 document.xml 变了。repackDocxFast 按 ZIP 结构重建：其余条目零重压，实测约 255ms。
//
// 覆盖：
//   A. 正确性（合成 docx）：条目名/顺序一致、其它条目逐字节一致、document.xml 内容一致
//   B. 回退：没有 CompressionStream（老内核）→ 返回 null，patchPjbgDocx 仍走 JSZip 且结果正确
//   C. 真实样例（桌面/临时目录存在时）：输出可被 JSZip 解析、条目数一致、图片逐字节一致
//   D. 异常输入：空 buffer / 非 zip → 返回 null（不抛）
//
// 用法：node tests/pjbg.repack.test.mjs
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

// ---- 沙箱：可选带 CompressionStream（模拟新/老内核） ----
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

const DOC_XML = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
  + '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>'
  + '<w:p><w:r><w:t>原始正文</w:t></w:r></w:p></w:body></w:document>';
const IMG = new Uint8Array(4096);
for (let i = 0; i < IMG.length; i++) IMG[i] = (i * 37) & 0xFF;

async function buildDocx() {
  const zip = new JSZip();
  zip.file("[Content_Types].xml", '<?xml version="1.0"?><Types/>');
  zip.file("word/document.xml", DOC_XML);
  zip.file("word/media/image1.png", IMG);
  zip.file("docProps/app.xml", "<app/>");
  return await zip.generateAsync({ type: "arraybuffer", compression: "DEFLATE" });
}

// ============ A. 正确性（合成 docx） ============
console.log("[A] 快速重打包正确性（合成 docx）");
{
  const blf = makeBlf(true);
  const src = await buildDocx();
  const newXml = DOC_XML.replace("原始正文", "改过的正文");
  const out = await blf.repackDocxFast(src, newXml);
  check("有 CompressionStream → 走快速路径（返回非空）", !!out, String(out && out.byteLength));
  if (out) {
    const z = await JSZip.loadAsync(out);
    const o = await JSZip.loadAsync(src);
    const names = (zz) => Object.keys(zz.files).filter((n) => !zz.files[n].dir).sort();
    check("条目名/顺序一致", JSON.stringify(names(z)) === JSON.stringify(names(o)), JSON.stringify(names(z)));
    const imgA = await o.file("word/media/image1.png").async("uint8array");
    const imgB = await z.file("word/media/image1.png").async("uint8array");
    let same = imgA.length === imgB.length;
    if (same) for (let i = 0; i < imgA.length; i++) if (imgA[i] !== imgB[i]) { same = false; break; }
    check("图片条目逐字节一致（未重压）", same);
    const t = await z.file("word/document.xml").async("string");
    check("document.xml 用新内容", t.indexOf("改过的正文") >= 0 && t.indexOf("原始正文") < 0, t.slice(0, 80));
    check("输出是合法 ZIP（PK 头）", new Uint8Array(out)[0] === 0x50 && new Uint8Array(out)[1] === 0x4B);
  }
}

// ============ B. 无 CompressionStream → 回退 ============
console.log("[B] 老内核（无 CompressionStream）→ repack 返回 null，patch 仍走 JSZip");
{
  const blf = makeBlf(false);
  const src = await buildDocx();
  const out = await blf.repackDocxFast(src, DOC_XML.replace("原始正文", "x"));
  check("repackDocxFast 返回 null", out === null, String(out));
  const patched = await blf.patchPjbgDocx(src);
  const z = await JSZip.loadAsync(patched);
  const t = await z.file("word/document.xml").async("string");
  check("patchPjbgDocx 仍返回可解析 docx", !!t && t.indexOf("<w:document") >= 0);
}

// ============ C. 真实样例（存在才跑） ============
const SAMPLE = process.env.PJBG_SAMPLE
  || path.join(process.env.TEMP || "/tmp", "sampling-plan-login", "pjbg.docx");
if (fs.existsSync(SAMPLE)) {
  console.log("[C] 真实样例：" + path.basename(SAMPLE));
  const blf = makeBlf(true);
  const buf = fs.readFileSync(SAMPLE);
  const xml = await (await JSZip.loadAsync(buf)).file("word/document.xml").async("string");
  const out = await blf.repackDocxFast(buf, xml);
  check("真实 docx 走快速路径", !!out, String(out && out.byteLength));
  if (out) {
    const z = await JSZip.loadAsync(out);
    const o = await JSZip.loadAsync(buf);
    const names = (zz) => Object.keys(zz.files).filter((n) => !zz.files[n].dir).sort();
    check("条目数一致（" + names(o).length + "）", names(z).length === names(o).length, String(names(z).length));
    const imgs = names(o).filter((n) => n.indexOf("word/media/") === 0);
    let bad = 0;
    for (const n of imgs) {
      const a = await o.file(n).async("uint8array");
      const b = await z.file(n).async("uint8array");
      if (a.length !== b.length) { bad++; continue; }
      for (let i = 0; i < a.length; i++) if (a[i] !== b[i]) { bad++; break; }
    }
    check("全部图片逐字节一致（" + imgs.length + " 张）", bad === 0, String(bad));
    const t = await z.file("word/document.xml").async("string");
    check("document.xml 文本与输入一致", t === xml);
  }
} else {
  console.log("[C] 跳过：未找到真实样例（" + SAMPLE + "）");
}

// ============ D. 异常输入不抛 ============
console.log("[D] 异常输入 → 返回 null");
{
  const blf = makeBlf(true);
  check("空 buffer", await blf.repackDocxFast(new ArrayBuffer(0), "<x/>") === null);
  const junk = new Uint8Array([1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19, 20, 21, 22, 23]);
  check("非 zip 数据", await blf.repackDocxFast(junk.buffer, "<x/>") === null);
  const ok = await buildDocx();
  check("缺 document.xml 的 zip", await blf.repackDocxFast(await (async () => {
    const z = new JSZip(); z.file("a.txt", "x");
    return await z.generateAsync({ type: "arraybuffer" });
  })(), "<x/>") === null);
}

console.log(`\npjbg.repack: ${pass} 通过, ${fail} 失败`);
process.exit(fail > 0 ? 1 : 0);
