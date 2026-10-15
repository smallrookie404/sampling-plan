// 回归测试：非噪声车间岗位筛选结果为空时，注段落 **A 段必须保留**（只删 B 段）。
// 起因：1009 项目实测报告（真实结构）在筛选为空时出现「A 段、B 段整段消失」的 bug。
// 本测试直接跑真实 1009 检评报告 docx（若存在），断言空筛选场景下：
//   ① document.xml 里仍能找到「注：用人单位职工总数…」段；
//   ② 该段带实时人数；
//   ③ 不含 B 段首句「接触的噪声经过验证性检测」。
// 用法：node tests/pjbg.note1009.test.mjs [1009报告.docx]
import fs from "node:fs";
import path from "node:path";
import vm from "node:vm";
import { fileURLToPath } from "node:url";
import JSZip from "jszip";

const here = path.dirname(fileURLToPath(import.meta.url));
const APP = path.resolve(here, "..");
const blfSrc = fs.readFileSync(path.join(APP, "js/blf.js"), "utf8");
const SAMPLE = process.argv[2] || process.env.PJBG_1009
  || path.join(process.env.TEMP || "/tmp", "sampling-plan-login", "pjbg-1009.docx");

let pass = 0, fail = 0;
function check(name, cond, extra) {
  if (cond) { pass++; console.log("  ok  " + name); }
  else { fail++; console.log("  FAIL " + name + (extra ? " :: " + extra : "")); }
}

if (!fs.existsSync(SAMPLE)) {
  console.log("跳过：未找到 1009 报告样本 " + SAMPLE);
  process.exit(0);
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
  URL: { createObjectURL: () => "blob:x", revokeObjectURL: noop }, JSZip };
sandbox.self = sandbox; sandbox.window = sandbox;
// 模拟真实浏览器：非噪声车间岗位筛选结果为空 → 文本为 ''
sandbox.SamplingUpload = {
  noiseNonNoiseFromBuf: async () => [],
  noiseNonNoiseTextFromBuf: async () => "",
};
vm.createContext(sandbox);
vm.runInContext(blfSrc, sandbox, { filename: "blf.js" });
const blf = sandbox.SamplingBlf;

const buf = fs.readFileSync(SAMPLE);
const zip = await JSZip.loadAsync(buf);
const xml = await zip.file("word/document.xml").async("string");

console.log("[1009] 空筛选 → A 段必须保留");
const r = blf.fixNoteAndContactCount(xml, "");
check("有改动", r.changed > 0, String(r.changed));
check("jobs 为空", r.jobs === "", JSON.stringify(r.jobs));
check("只产出 1 段（NoteTexts）", Array.isArray(r.noteTexts) && r.noteTexts.length === 1,
  r.noteTexts && String(r.noteTexts.length));
check("A 段以「注：用人单位职工总数」开头", r.noteTexts && r.noteTexts[0] && r.noteTexts[0].startsWith("注：用人单位职工总数"),
  r.noteTexts && r.noteTexts[0] && r.noteTexts[0].slice(0, 20));
check("A 段含「接触职业病危害人数」", r.noteTexts && r.noteTexts[0] && r.noteTexts[0].includes("接触职业病危害人数"));

// 处理后 xml 里必须还能找到「注：用人单位职工总数」这段
check("处理后 xml 仍含 A 段文本", r.xml.includes("注：用人单位职工总数"));
check("处理后 xml 已无 B 段首句", !r.xml.includes("接触的噪声经过验证性检测"));
check("处理后 xml 已无「注：无」", !r.xml.includes("注：无"));

// 整链路
console.log("[1009] patchPjbgDocx 整链路");
const patched = await blf.patchPjbgDocx(buf);
const z2 = await JSZip.loadAsync(patched);
const xml2 = await z2.file("word/document.xml").async("string");
check("重打包后仍含 A 段文本", xml2.includes("注：用人单位职工总数"));
check("重打包后已无 B 段首句", !xml2.includes("接触的噪声经过验证性检测"));
// 统计「注：用人单位职工总数」出现次数必须恰好 1 次（不能重复替换）
const cnt = (xml2.match(/注：用人单位职工总数/g) || []).length;
check("A 段只出现 1 次（无重复插入）", cnt === 1, String(cnt));
// 整段没消失：A 段结尾句必须在
check("重打包后含 A 段结尾「12:30～17:00。」", xml2.includes("12:30～17:00。"));

console.log("\n结果：" + pass + " 通过 / " + fail + " 失败");
process.exit(fail ? 1 : 0);
