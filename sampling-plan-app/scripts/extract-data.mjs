import fs from "node:fs";
import path from "node:path";
import JSZip from "jszip";

const SRC = "G:/codex工作台/采样计划/0.系统测点布局调查2025.12.19（6）.xlsx";
const OUT = path.resolve("js/data.js");

function unescapeXml(s) {
  return s
    .replace(/&lt;/g, "<")
    .replace(/&gt;/g, ">")
    .replace(/&quot;/g, '"')
    .replace(/&apos;/g, "'")
    .replace(/&#(\d+);/g, (_, d) => String.fromCodePoint(Number(d)))
    .replace(/&#x([0-9a-fA-F]+);/g, (_, d) => String.fromCodePoint(parseInt(d, 16)))
    .replace(/&amp;/g, "&");
}

function escapeXml(s) {
  return String(s)
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&apos;");
}

function parseSharedStrings(xml) {
  const out = [];
  const siRe = /<si>(.*?)<\/si>/gs;
  let m;
  while ((m = siRe.exec(xml))) {
    const inner = m[1];
    const text = unescapeXml(
      [...inner.matchAll(/<t(?: [^>]*)?>(.*?)<\/t>/gs)]
        .map((t) => t[1])
        .join("")
    );
    out.push(text);
  }
  return out;
}

function colToIndex(ref) {
  const mm = /^([A-Z]+)(\d+)$/.exec(ref);
  if (!mm) return null;
  let c = 0;
  for (const ch of mm[1]) c = c * 26 + (ch.charCodeAt(0) - 64);
  return { col: c - 1, row: Number(mm[2]) - 1 };
}

function parseSheet(xml, sharedStrings) {
  const grid = new Map();
  const cellRe = /<c\s+([^>]*?)(?:\/>|>(.*?)<\/c>)/gs;
  let m;
  while ((m = cellRe.exec(xml))) {
    const attrs = m[1];
    const body = m[2] || "";
    const ref = /r="([A-Z]+\d+)"/.exec(attrs);
    if (!ref) continue;
    const { col, row } = colToIndex(ref[1]);
    const t = /t="([^"]+)"/.exec(attrs);
    const type = t ? t[1] : "";
    const v = /<v>(.*?)<\/v>/s.exec(body);
    let val = null;
    if (type === "inlineStr") {
      const is = /<is>(.*?)<\/is>/s.exec(body);
      if (is) {
        val = unescapeXml(
          [...is[1].matchAll(/<t(?: [^>]*)?>(.*?)<\/t>/gs)]
            .map((x) => x[1])
            .join("")
        );
      }
    } else if (type === "s" && v) {
      val = sharedStrings[Number(v[1])] ?? null;
    } else if (v) {
      const raw = v[1];
      val = type === "str" || type === "e" ? unescapeXml(raw) : raw;
      if (val !== null && val !== "" && type !== "str" && type !== "e" && !isNaN(Number(val))) {
        val = Number(val);
      }
    }
    grid.set(`${row}|${col}`, val);
  }
  return grid;
}

function cell(grid, row, col) {
  const v = grid.get(`${row}|${col}`);
  return v === null || v === undefined ? "" : v;
}

const zip = await JSZip.loadAsync(await fs.promises.readFile(SRC));

// workbook.xml -> sheet name -> file
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

function sheetByName(name) {
  const file = relFile[sheetMap[name]];
  return zip.file(file).async("string");
}

const ssXml = await zip.file("xl/sharedStrings.xml").async("string");
const sharedStrings = parseSharedStrings(ssXml);

// --- 危害因素 ---
const hazardXml = await sheetByName("危害因素");
const hazardGrid = parseSheet(hazardXml, sharedStrings);
const HAZARD_HEADERS = ["识别", "系统名称", "粉尘性质", "定性分析", "委外检测", "计算TWA", "计算STEL", "计算CPE", "计算MAC", "结果保留位数", "存在高毒物品", "不检测原因说明"];
const hazardFactors = [];
for (let r = 1; r < 612; r++) {
  const rec = cell(hazardGrid, r, 0);
  if (rec === "") continue;
  hazardFactors.push({
    rec,
    name: cell(hazardGrid, r, 1),
    dust: cell(hazardGrid, r, 2),
    qual: cell(hazardGrid, r, 3),
    outsource: cell(hazardGrid, r, 4),
    twa: cell(hazardGrid, r, 5),
    stel: cell(hazardGrid, r, 6),
    cpe: cell(hazardGrid, r, 7),
    mac: cell(hazardGrid, r, 8),
    digits: cell(hazardGrid, r, 9),
    highTox: cell(hazardGrid, r, 10),
    noTestReason: cell(hazardGrid, r, 11),
  });
}

// --- 检测项目 ---
const itemXml = await sheetByName("检测项目");
const itemGrid = parseSheet(itemXml, sharedStrings);
const detectionItems = [];
for (let r = 1; r < 445; r++) {
  const v = cell(itemGrid, r, 0);
  if (v !== "") detectionItems.push(v);
}

// --- 主表表头（导出对齐用） ---
const mainXml = await sheetByName("测点布局情况调查");
const mainGrid = parseSheet(mainXml, sharedStrings);
const mainHeaders = [];
for (let c = 0; c < 61; c++) {
  if (c === 21) {
    mainHeaders.push(""); // V 列空表头
    continue;
  }
  mainHeaders.push(cell(mainGrid, 0, c));
}
// 自动计算区第 29 列（AC 列）表头规范化为“*作业方式”，与自动区其他列风格一致
if (mainHeaders[28] === "作业方式") mainHeaders[28] = "*作业方式";

const data = { hazardFactors, detectionItems, mainHeaders };
const body =
  "/* 由 scripts/extract-data.mjs 从原 Excel 自动生成，请勿手工编辑 */\n" +
  "(function (root, factory) {\n" +
  "  if (typeof module === 'object' && module.exports) module.exports = factory();\n" +
  "  else root.SamplingData = factory();\n" +
  "})(typeof self !== 'undefined' ? self : this, function () {\n" +
  "  return " + JSON.stringify(data) + ";\n" +
  "});\n";

fs.writeFileSync(OUT, body, "utf8");
console.log(
  `OK hazard=${hazardFactors.length} items=${detectionItems.length} headers=${mainHeaders.length}`
);
console.log("mainHeaders:", JSON.stringify(mainHeaders.slice(0, 24)));
