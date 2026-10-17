// 性能基准：measure patchPjbgDocx 端到端耗时（含噪声表解析次数）。
// 用法：node tests/_bench_pjbg.mjs
import fs from "node:fs";
import path from "node:path";
import vm from "node:vm";
import { fileURLToPath } from "node:url";
import JSZip from "jszip";

const here = path.dirname(fileURLToPath(import.meta.url));
const APP = path.resolve(here, "..");
const uploadSrc = fs.readFileSync(path.join(APP, "js/upload.js"), "utf8");
const utilSrc = fs.readFileSync(path.join(APP, "js/util.js"), "utf8");
const blfSrc = fs.readFileSync(path.join(APP, "js/blf.js"), "utf8");

const noop = () => {};
const fakeEl = { style: {}, classList: { add: noop, remove: noop, toggle: noop }, addEventListener: noop,
  appendChild: noop, removeChild: noop, setAttribute: noop, getAttribute: () => null, hasAttribute: () => false,
  querySelector: () => null, querySelectorAll: () => [], innerHTML: "", textContent: "", disabled: false,
  click: noop, focus: noop, blur: noop, dispatchEvent: noop, dataset: {} };
// 每个 getElementById 返回独立 stub（upload.js 顶层会注册监听器，不能为 null）
function mkEl() {
  const el = Object.assign({}, fakeEl);
  el.classList = { add: noop, remove: noop, toggle: noop, contains: () => false };
  el.style = {};
  el.children = [];
  el.parentNode = null;
  return el;
}
const documentStub = {
  readyState: "complete",
  getElementById: () => mkEl(), querySelector: () => mkEl(), querySelectorAll: () => [],
  createElement: () => mkEl(), addEventListener: noop,
  body: mkEl(), head: mkEl(), documentElement: mkEl(),
};

// DOMParser 用轻量实现（够解析我们的夹具 XML）：
// 由于 Node 无原生 DOMParser，用最小 xml 解析器替代会太脆。
// 改为：真实场景下 patchPjbgDocx 只用 DOMParser 在 upload.js 的 parseTableAfterHeadingFromDocx 里。
// 这里为了测真实代价，用一个基于正则的实现模拟「扫全部 <w:tc>」的等价工作量。

const sandbox = {
  console, document: documentStub,
  DOMParser: class { parseFromString() { return { getElementsByTagNameNS: () => [] }; } },
  setTimeout, clearTimeout,
  fetch: () => Promise.reject(new Error("no fetch")),
  URL: { createObjectURL: () => "blob:x", revokeObjectURL: noop },
  JSZip, Promise, Map, Array, Object, String, Number, Math, JSON, RegExp, Error, isNaN, parseInt, parseFloat,
};
sandbox.self = sandbox; sandbox.window = sandbox;
vm.createContext(sandbox);
vm.runInContext(utilSrc, sandbox, { filename: "util.js" });
vm.runInContext(uploadSrc, sandbox, { filename: "upload.js" });
vm.runInContext(blfSrc, sandbox, { filename: "blf.js" });

const blf = sandbox.SamplingBlf;
const up = sandbox.SamplingUpload;

// ---- 构造夹具 docx：一个「噪声检测结果分析与评价」大表 + 若干段落表 ----
function esc(s) { return String(s).replace(/&/g, "&amp;").replace(/</g, "&lt;"); }
function p(text, runs) {
  const parts = runs || [text];
  return "<w:p>" + parts.map((t) => '<w:r><w:t xml:space="preserve">' + esc(t) + "</w:t></w:r>").join("") + "</w:p>";
}
function tbl(headers, rows) {
  const hc = headers.map((h) => "<w:tc>" + p(h) + "</w:tc>").join("");
  const rc = rows.map((r) => "<w:tr>" + r.map((c) => "<w:tc>" + p(c) + "</w:tc>").join("") + "</w:tr>").join("");
  return "<w:tbl><w:tr>" + hc + "</w:tr>" + rc + "</w:tbl>";
}

const NOISE_HDR = ["测量编号", "单元/工作场所", "岗位/工种/测量点/对象", "接触时间", "噪声性质",
  "测量范围", "测量结果", "LEX,8h", "LEX,W", "接触限值", "单项结论"];
const noiseRows = [];
for (let i = 1; i <= 240; i++) {
  const unit = "车间" + (1 + (i % 12));
  const job = "岗位" + (1 + (i % 30)) + (i % 4 === 0 ? "/点A" : "");
  const lex = String(70 + (i % 22));
  noiseRows.push([String(i), unit, job, "4h", "稳态", "70~95", lex, lex, "", "85", Number(lex) >= 85 ? "超标" : "合格"]);
}
// 若干「确认表」/「评价依据」表，制造真实文档体积
const confirmHdr = ["序号", "职业病危害因素", "检测项目", "备注"];
const confirmRows = Array.from({ length: 40 }, (_, i) => [String(i + 1), "因素" + i, "项目" + i, ""]);
const evalItems = Array.from({ length: 60 }, (_, i) => "（" + (i + 1) + "）《标准" + i + "》；");

let bodyXml = p("7.3 噪声检测结果分析与评价");
bodyXml += tbl(NOISE_HDR, noiseRows);
bodyXml += p("1.2.3评价依据");
for (const it of evalItems) bodyXml += p(it, ["（ " + (evalItems.indexOf(it) + 1) + "）", it.slice(it.indexOf("）") + 1)]);
bodyXml += p("表4-1 职业病危害因素");
bodyXml += tbl(confirmHdr, confirmRows);
// 再塞大量无关段落，模拟真实文档（几十页）
for (let i = 0; i < 800; i++) bodyXml += p("这是正文章节第" + i + "段，用于模拟真实报告体量，包含若干描述性文字与说明内容。");
const docXml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
  + '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
  + "<w:body>" + bodyXml + "</w:body></w:document>";

async function buildDocx() {
  const zip = new JSZip();
  zip.file("[Content_Types].xml", '<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"/>');
  zip.file("word/document.xml", docXml);
  return await zip.generateAsync({ type: "arraybuffer", compression: "DEFLATE" });
}

const buf = await buildDocx();
console.log("夹具大小:", (buf.byteLength / 1024).toFixed(1), "KB ; document.xml:", (docXml.length / 1024).toFixed(1), "KB");

// 打点统计「噪声表解析」次数：包一层 SamplingUpload 的噪声取数函数
let parseCount = 0;
const origParse = up.parseNoiseTable;
// upload.js 内部调用的是闭包内 parsePjbgNoiseTableFromDocx，无法直接拦截；
// 用代理 SamplingUpload.*FromBuf 计数来近似「外部调用次数」。

async function runOnce(label) {
  const t0 = performance.now();
  const out = await blf.patchPjbgDocx(buf);
  const t1 = performance.now();
  console.log("  " + label + ": " + (t1 - t0).toFixed(0) + " ms  (输出 " + (out.byteLength / 1024).toFixed(1) + " KB)");
  return t1 - t0;
}

console.log("\n== 冷启动（含 jszip 首次加载/预热）==");
await runOnce("run1");
console.log("\n== 热态（重复 5 次取均值）==");
let sum = 0;
for (let i = 0; i < 5; i++) sum += await runOnce("run" + (i + 1));
console.log("  平均: " + (sum / 5).toFixed(0) + " ms");
