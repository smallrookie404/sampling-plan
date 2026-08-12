import fs from "node:fs";
import path from "node:path";
import { createRequire } from "node:module";
import JSZip from "jszip";

const require = createRequire(import.meta.url);
const data = require("../js/data.js");
const logic = require("../js/logic.js");

const SRC = "G:/codex工作台/采样计划/0.系统测点布局调查2025.12.19（6）.xlsx";

// ---------- 读取原表（输入 A:U 与缓存计算 W:BI） ----------
function unescapeXml(s) {
  return s
    .replace(/&lt;/g, "<").replace(/&gt;/g, ">").replace(/&quot;/g, '"')
    .replace(/&apos;/g, "'")
    .replace(/&#(\d+);/g, (_, d) => String.fromCodePoint(Number(d)))
    .replace(/&amp;/g, "&");
}
function parseSharedStrings(xml) {
  return [...xml.matchAll(/<si>(.*?)<\/si>/gs)].map((m) =>
    unescapeXml([...m[1].matchAll(/<t(?: [^>]*)?>(.*?)<\/t>/gs)].map((t) => t[1]).join(""))
  );
}
function colToIndex(ref) {
  const mm = /^([A-Z]+)(\d+)$/.exec(ref);
  let c = 0;
  for (const ch of mm[1]) c = c * 26 + (ch.charCodeAt(0) - 64);
  return { col: c - 1, row: Number(mm[2]) - 1 };
}
function parseSheet(xml, ss) {
  const grid = new Map();
  for (const m of xml.matchAll(/<c\s+([^>]*?)(?:\/>|>(.*?)<\/c>)/gs)) {
    const attrs = m[1], body = m[2] || "";
    const ref = /r="([A-Z]+\d+)"/.exec(attrs);
    if (!ref) continue;
    const { col, row } = colToIndex(ref[1]);
    const t = /t="([^"]+)"/.exec(attrs);
    const type = t ? t[1] : "";
    const v = /<v>(.*?)<\/v>/s.exec(body);
    let val = null;
    if (type === "inlineStr") {
      const is = /<is>(.*?)<\/is>/s.exec(body);
      if (is) val = unescapeXml([...is[1].matchAll(/<t(?: [^>]*)?>(.*?)<\/t>/gs)].map((x) => x[1]).join(""));
    } else if (type === "s" && v) val = ss[Number(v[1])] ?? null;
    else if (v) {
      val = type === "str" || type === "e" ? unescapeXml(v[1]) : v[1];
      if (val !== "" && type !== "str" && type !== "e" && !isNaN(Number(val))) val = Number(val);
    }
    grid.set(`${row}|${col}`, val);
  }
  return grid;
}

const zip = await JSZip.loadAsync(await fs.promises.readFile(SRC));
const wbXml = await zip.file("xl/workbook.xml").async("string");
const relsXml = await zip.file("xl/_rels/workbook.xml.rels").async("string");
const sheetMap = {};
for (const m of wbXml.matchAll(/<sheet\s+([^>]*?)\/>/gs)) {
  const name = /name="([^"]+)"/.exec(m[1])?.[1];
  const rid = /r:id="([^"]+)"/.exec(m[1])?.[1];
  if (name && rid) sheetMap[name] = rid;
}
const relFile = {};
for (const m of relsXml.matchAll(/<Relationship\s+([^>]*?)\/>/gs)) {
  const id = /Id="([^"]+)"/.exec(m[1])?.[1];
  const target = /Target="([^"]+)"/.exec(m[1])?.[1];
  if (id && target) relFile[id] = "xl/" + target.replace(/^\//, "");
}
const ssXml = await zip.file("xl/sharedStrings.xml").async("string");
const ss = parseSharedStrings(ssXml);
const mainXml = await zip.file(relFile[sheetMap["测点布局情况调查"]]).async("string");
const grid = parseSheet(mainXml, ss);

const cell = (row, col) => {
  const v = grid.get(`${row}|${col}`);
  return v === null || v === undefined ? "" : v;
};

const INPUT_COLS = "ABCDEFGHIJKLMNOPQRSTU".split("");
const COMPUTED_COLS = [
  "W", "X", "Y", "Z", "AA", "AB", "AC", "AD", "AE", "AF", "AG", "AH", "AI",
  "AJ", "AK", "AL", "AM", "AN", "AO", "AP", "AQ", "AR", "AS", "AT", "AU",
  "AV", "AW", "AX", "AY", "AZ", "BA", "BB", "BC", "BD", "BE", "BF", "BG",
  "BH", "BI",
];
const colIndex = (letter) => {
  let c = 0;
  for (const ch of letter) c = c * 26 + (ch.charCodeAt(0) - 64);
  return c - 1;
};

const rows = [];
for (let r = 1; r < 9; r++) {
  const input = {};
  for (const c of INPUT_COLS) input[c] = cell(r, colIndex(c));
  const manual = { AI: "", AJ: "", AX: "", BE: "", BF: "", BG: "", BH: "" };
  for (const c of ["AI", "AJ", "AX", "BE", "BF", "BG", "BH"]) manual[c] = cell(r, colIndex(c));
  rows.push({ input, manual, overridden: {}, values: {}, errors: {} });
}

const expected = rows.map((_, i) => {
  const r = i + 1;
  const obj = {};
  for (const c of COMPUTED_COLS) obj[c] = cell(r, colIndex(c));
  return obj;
});

// ---------- 引擎计算 ----------
logic.computeRows(rows, data);

// ---------- 对比 ----------
const norm = (v) => {
  if (v === "" || v === null || v === undefined) return "";
  if (typeof v === "number") return String(v);
  const n = Number(v);
  return !isNaN(n) && String(v).trim() !== "" ? String(n) : String(v);
};

let pass = 0, fail = 0;
const failures = [];
for (let i = 0; i < rows.length; i++) {
  for (const c of COMPUTED_COLS) {
    const got = norm(rows[i].values[c]);
    const exp = norm(expected[i][c]);
    if (got !== exp) {
      fail++;
      failures.push(`行${i + 2} ${c}: 引擎=${got} Excel缓存=${exp}`);
    } else pass++;
  }
}

console.log(`对比单元格: 通过 ${pass} / 失败 ${fail}`);
if (failures.length) {
  console.log("失败明细:");
  for (const f of failures.slice(0, 60)) console.log("  " + f);
  process.exit(1);
}

// ---------- 附加单元检查 ----------
const h = logic.findHazardByRec(data, "二氧化钛粉尘(总尘)");
if (!h || h.name !== "二氧化钛粉尘") throw new Error("findHazardByRec 失败");
const h2 = logic.findHazardByName(data, "二氧化钛粉尘");
if (!h2 || h2.dust !== "总尘") throw new Error("findHazardByName 失败");

const badRow = {
  input: { A: "1车间", B: "操作工", C: "投料", D: "噪声", R: "白班", S: "随便", U: "五" },
  manual: {}, overridden: {}, values: {}, errors: {},
};
logic.computeRows([badRow], data);
const errs = badRow.errors;
if (!errs.R || !errs.S || !errs.U || !errs.AE) throw new Error("校验逻辑未覆盖 R/S/U/AE");
if (!errs.AN || errs.AN !== "接害因素在危害因素库中未找到，请检查名称或补充库表") {
  // D=噪声 实际能查到，故 AN 不应报错
  if (errs.AN) throw new Error("AN 误报错误");
}

const dup = logic.findDuplicates(data.detectionItems);
if (dup.length) throw new Error("检测项目不应有重复: " + dup.join(","));

// 备注列（BI）：默认自动生成，手动编辑后保留手动值
if (!Array.isArray(logic.TEXT_OVERRIDE_COLS) || !logic.TEXT_OVERRIDE_COLS.includes("BI")) {
  throw new Error("TEXT_OVERRIDE_COLS 应包含 BI");
}
const remarkRow = {
  input: { A: "1车间", B: "操作工", C: "投料", D: "碘" },
  manual: {}, overridden: {}, values: {}, errors: {},
};
logic.computeRows([remarkRow], data);
if (remarkRow.values.BI !== "无合适外包机构") {
  throw new Error("备注列未自动生成: " + JSON.stringify(remarkRow.values.BI));
}
remarkRow.overridden.BI = true;
remarkRow.values.BI = "手动备注";
logic.computeRows([remarkRow], data);
if (remarkRow.values.BI !== "手动备注") {
  throw new Error("备注列手动值被覆盖: " + JSON.stringify(remarkRow.values.BI));
}

// Unicode 安全 base64 往返（GitHub API 数据读写用）
const demoJson = JSON.stringify({ name: "测试记录", rows: [{ A: "3号车间", D: "二氧化钛粉尘(总尘)" }] });
if (logic.decodeUnicodeBase64(logic.encodeUnicodeBase64(demoJson)) !== demoJson) {
  throw new Error("Unicode base64 往返失败");
}
console.log("Unicode base64 往返校验通过 ✔");

console.log("引擎校验：全部通过 ✔");
