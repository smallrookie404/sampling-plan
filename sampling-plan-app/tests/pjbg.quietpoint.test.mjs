// 回归测试：检评报告后处理②c「噪声岗位内低噪点位单独删掉『噪声』危害」。
//
// 背景（用户要求 2026-10-19）：
//   某「车间+岗位」的噪声 LEX 最大值 ≥ 80 dB（属噪声作业岗位，② 不整体删噪声）时，
//   噪声表里该岗位常把多个点位挤在同一单元格内（单元格多行），「测量结果[dB(A)」同样多行，
//   每个点位固定 3 个测量值（第 1 个点位=前 3 行，第 2 个=4~6 行…）。
//   逐点位判断：3 个测量值**全部 < 80 dB** → 该点位不接触噪声 → 表4-1 里该点位那一行删掉「噪声」。
//   前提：同岗位至少有一个点位 3 个测量值全部 ≥ 80 dB（否则整岗位走②，不能删光）。
//
// 覆盖：
//   A. 单点位一行的形态：只有低噪点位的行被删，噪声点位的行保留；
//   B. 「工作地点及内容」多行（多个点位共用一行）时：全部低噪才删，否则整行保留；
//   C. 表5-1 检测项目确认表（无「工作地点…」列）不被误改；
//   D. 空 quietPoints → 原样不动；
//   E. patchPjbgDocx 接线：② 与 ②c 同时生效（老接口无 quietPointPairs 时不报错）。
//
// 用法：node tests/pjbg.quietpoint.test.mjs
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
const documentStub = { readyState: "complete", getElementById: fakeEl, querySelector: fakeEl,
  querySelectorAll: () => [], createElement: fakeEl, addEventListener: noop,
  body: fakeEl(), head: fakeEl(), documentElement: fakeEl() };

function makeSandbox(uploadStub) {
  const sandbox = {
    console: { log: noop, info: noop, warn: noop, error: noop },
    document: documentStub,
    DOMParser: class { parseFromString() { return { getElementsByTagNameNS: () => [] }; } },
    setTimeout, clearTimeout, fetch: () => Promise.reject(new Error("no fetch")),
    URL: { createObjectURL: () => "blob:x", revokeObjectURL: noop },
    JSZip, TextDecoder,
  };
  sandbox.self = sandbox; sandbox.window = sandbox;
  if (uploadStub) sandbox.SamplingUpload = uploadStub;
  vm.createContext(sandbox);
  vm.runInContext(blfSrc, sandbox, { filename: "blf.js" });
  return sandbox.SamplingBlf;
}

// ---- XML 夹具 ----
const esc = (s) => String(s).replace(/&/g, "&amp;").replace(/</g, "&lt;");
function p(text) {
  return "<w:p><w:r><w:t xml:space=\"preserve\">" + esc(text) + "</w:t></w:r></w:p>";
}
// 多行单元格：行间用 <w:br/>（真实报告形态）
function pLines(lines) {
  let s = "";
  lines.forEach((ln, i) => {
    if (i > 0) s += "<w:r><w:br/></w:r>";
    s += "<w:r><w:t xml:space=\"preserve\">" + esc(ln) + "</w:t></w:r>";
  });
  return "<w:p>" + s + "</w:p>";
}
const tc = (text) => "<w:tc>" + p(text) + "</w:tc>";
const tcLines = (lines) => "<w:tc>" + pLines(lines) + "</w:tc>";
const tr = (cells) => "<w:tr>" + cells.join("") + "</w:tr>";
const tbl = (headCells, rows) => "<w:tbl>" + tr(headCells) + rows.join("") + "</w:tbl>";

// 表4-1：表头 13 格；数据行 14 格（「人数」拆成 总数 / 数/班）→ 列号需按差值平移
const H41 = ["单元/工作场所", "岗位（工种）", "人数", "工作班制", "工作时间", "作业类型", "作业方式",
  "接触类型", "工作地点及内容", "职业病危害因素", "来源", "接触时间", "体力劳动强度"];
function row41(unit, job, place, hazard) {
  // place 可以是字符串（单行）或数组（多行）
  const placeCell = Array.isArray(place) ? tcLines(place) : tc(place);
  return tr([tc(unit), tc(job), tc("1"), tc("1"), tc("长白班"), tc("9h/d\n6d/w"), tc("固定作业"),
    tc("半手工作业"), tc("稳定"), placeCell, tc(hazard), tc("设备运行"), tc("7h/d\n6d/w"), tc("-")]);
}
const T41_HEAD = H41.map(tc);

// 表5-1 检测项目确认表（也有「职业病危害因素」列，但没有「工作地点…」列）
const H51 = ["单元/工作场所", "岗位/工种/采样/测量点/对象", "职业病危害因素", "检测项目", "情况说明"];
function row51(unit, jobPoint, hazard) {
  return tr([tc(unit), tc(jobPoint), tc(hazard), tc(hazard), tc("-")]);
}

// 取表4-1「职业病危害因素」列文本（用 blf 自己的实现，顺带覆盖列号平移）
function hazardTexts(blf, xml) {
  const r = blf.collectConfirmColTexts(xml, ["职业病危害因素"]);
  return (r && r.texts) || [];
}
function allText(xml) {
  const out = [];
  const re = /<w:t(?:\s[^>]*)?>([\s\S]*?)<\/w:t>/g;
  let m; while ((m = re.exec(xml))) out.push(m[1]);
  return out.join("");
}

const blf = makeSandbox(null);

// ============ A. 单点位一行：只删低噪点位的行 ============
console.log("[A] 单点位一行：针车2（低噪）删噪声，针车1（噪声）保留");
{
  const xml = tbl(T41_HEAD, [
    row41("生产车间(2F)", "针车岗位", "针车1/操作设备针车作业", "噪声"),
    row41("生产车间(2F)", "针车岗位", "针车2/操作设备针车作业", "噪声"),
    row41("生产车间(2F)", "套邦岗位", "套邦/操作设备套邦作业", "噪声,电焊烟尘"),
    row41("生产车间(2F)", "刷胶岗位", "刷胶/操作设备刷胶作业", "噪声"),
  ]);
  const quiet = [
    { unit: "生产车间(2F)", job: "针车岗位", point: "针车2" },
    { unit: "生产车间(2F)", job: "套邦岗位", point: "套邦" },
  ];
  const r = blf.fixQuietPointHazards(xml, quiet);
  const texts = hazardTexts(blf, r.xml);
  check("改动 2 处", r.changed === 2, String(r.changed));
  check("针车1（噪声点位）保留「噪声」", texts[0] === "噪声", JSON.stringify(texts));
  check("针车2（低噪点位）删掉「噪声」→ 填 -", texts[1] === "-", JSON.stringify(texts));
  check("套邦（低噪点位）只删噪声，电焊烟尘保留", texts[2] === "电焊烟尘", JSON.stringify(texts));
  check("刷胶（不在低噪列表）不动", texts[3] === "噪声", JSON.stringify(texts));
  check("表结构完整（行数不变）", (r.xml.match(/<w:tr>/g) || []).length === 5,
    String((r.xml.match(/<w:tr>/g) || []).length));
}

// ============ B. 「工作地点及内容」多行：多个点位共用一行 ============
console.log("[B] 同一行含多个点位：全部低噪才删");
{
  const build = () => tbl(T41_HEAD, [
    row41("生产车间(2F)", "针车岗位", ["针车1/操作设备针车作业", "针车2/操作设备针车作业"], "噪声"),
    row41("生产车间(2F)", "针车岗位", ["针车3/操作设备针车作业", "针车4/操作设备针车作业"], "噪声"),
  ]);
  // 只命中 针车2 → 第 1 行仍有针车1（噪声），整行保留；第 2 行两点位都低噪 → 删
  const r1 = blf.fixQuietPointHazards(build(), [
    { unit: "生产车间(2F)", job: "针车岗位", point: "针车2" },
    { unit: "生产车间(2F)", job: "针车岗位", point: "针车3" },
    { unit: "生产车间(2F)", job: "针车岗位", point: "针车4" },
  ]);
  const t1 = hazardTexts(blf, r1.xml);
  check("部分点位低噪 → 该行保留噪声", t1[0] === "噪声", JSON.stringify(t1));
  check("全部点位低噪 → 该行删噪声", t1[1] === "-", JSON.stringify(t1));
  check("changed = 1", r1.changed === 1, String(r1.changed));
}

// ============ C. 表5-1（无「工作地点」列）不被误改 ============
console.log("[C] 表5-1 检测项目确认表不受影响");
{
  const xml = tbl(H51.map(tc), [
    row51("生产车间(2F)", "针车岗位/针车2", "噪声"),
    row51("生产车间(2F)", "套邦岗位/套邦", "噪声"),
  ]);
  const r = blf.fixQuietPointHazards(xml, [
    { unit: "生产车间(2F)", job: "针车岗位", point: "针车2" },
  ]);
  check("changed === 0", r.changed === 0, String(r.changed));
  check("XML 完全未变", r.xml === xml);
}

// ============ D. 空列表 / 未命中 → 原样 ============
console.log("[D] 空列表与未命中");
{
  const xml = tbl(T41_HEAD, [row41("车间A", "岗位A", "点1/作业", "噪声")]);
  check("空列表 → 原样", blf.fixQuietPointHazards(xml, []).xml === xml);
  check("null → 原样", blf.fixQuietPointHazards(xml, null).xml === xml);
  const r = blf.fixQuietPointHazards(xml, [{ unit: "车间B", job: "岗位A", point: "点1" }]);
  check("车间不匹配 → 原样", r.xml === xml && r.changed === 0);
  const r2 = blf.fixQuietPointHazards(xml, [{ unit: "车间A", job: "岗位A", point: "点2" }]);
  check("点位不匹配 → 原样", r2.xml === xml && r2.changed === 0);
  // 已没有「噪声」的行不重复计数
  const xml2 = tbl(T41_HEAD, [row41("车间A", "岗位A", "点1/作业", "电焊烟尘")]);
  const r3 = blf.fixQuietPointHazards(xml2, [{ unit: "车间A", job: "岗位A", point: "点1" }]);
  check("危害列无噪声 → 不改", r3.xml === xml2 && r3.changed === 0);
}

// ============ E. patchPjbgDocx 接线 ============
console.log("[E] patchPjbgDocx：② 与 ②c 同时生效");
{
  const body = p("表4-1 岗位（工种）设置、劳动者作业情况及职业病危害因素识别情况")
    + tbl(T41_HEAD, [
      row41("生产车间(2F)", "针车岗位", "针车1/操作设备针车作业", "噪声"),
      row41("生产车间(2F)", "针车岗位", "针车2/操作设备针车作业", "噪声"),
      row41("生产车间(2F)", "打磨岗位", "打磨/操作设备打磨作业", "噪声,电焊烟尘"),
    ]);
  const xmlDoc = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    + '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>'
    + body + "</w:body></w:document>";
  const zip = new JSZip();
  zip.file("word/document.xml", xmlDoc);
  const buf = await zip.generateAsync({ type: "arraybuffer", compression: "DEFLATE" });

  const INFO = {
    rows: [], ok: true,
    nonNoiseGroups: [{ unit: "生产车间(2F)", jobs: ["打磨岗位"] }],
    nonNoisePairs: [{ unit: "生产车间(2F)", job: "打磨岗位" }],
    nonNoiseText: "生产车间(2F)打磨岗位",
    noiseWorkshopGroups: [{ unit: "生产车间(2F)", jobs: ["针车岗位"] }],
    quietPointPairs: [{ unit: "生产车间(2F)", job: "针车岗位", point: "针车2" }],
  };
  const blf2 = makeSandbox({ getCachedNoiseInfo: () => INFO });
  const out = await blf2.patchPjbgDocx(buf);
  const outZip = await JSZip.loadAsync(out);
  const outXml = await outZip.file("word/document.xml").async("string");
  const texts = hazardTexts(blf2, outXml);
  check("② 打磨岗位（非噪声）删掉噪声，电焊烟尘保留", texts[2] === "电焊烟尘", JSON.stringify(texts));
  check("②c 针车2（低噪点位）删掉噪声", texts[1] === "-", JSON.stringify(texts));
  check("②c 针车1（噪声点位）保留噪声", texts[0] === "噪声", JSON.stringify(texts));

  // 老版 upload.js（无 quietPointPairs）→ 不报错、② 仍生效
  const OLD = { ...INFO };
  delete OLD.quietPointPairs;
  const blf3 = makeSandbox({ getCachedNoiseInfo: () => OLD });
  const out3 = await blf3.patchPjbgDocx(buf);
  const out3Xml = await (await JSZip.loadAsync(out3)).file("word/document.xml").async("string");
  const texts3 = hazardTexts(blf3, out3Xml);
  check("无 quietPointPairs → 针车2 保留噪声（向后兼容）", texts3[1] === "噪声", JSON.stringify(texts3));
  check("无 quietPointPairs → ② 仍删打磨岗位噪声", texts3[2] === "电焊烟尘", JSON.stringify(texts3));
  check("输出仍是合法 document.xml", allText(outXml).indexOf("表4-1") >= 0);
}

console.log(`\n结果：${pass} 通过 / ${fail} 失败`);
process.exit(fail ? 1 : 0);
