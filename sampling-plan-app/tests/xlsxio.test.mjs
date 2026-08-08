import fs from "node:fs";
import { createRequire } from "node:module";
import { FileBlob, SpreadsheetFile } from "@oai/artifact-tool";

const require = createRequire(import.meta.url);
const data = require("../js/data.js");
const logic = require("../js/logic.js");
const xlsxio = require("../js/xlsxio.js");

const INPUT_COLS = logic.INPUT_COLS;
const MANUAL_COLS = logic.MANUAL_COLS;
const OVERRIDE_COLS = logic.OVERRIDE_COLS;
const COMPUTED_COLS = logic.COMPUTED_COLS;
const HEADERS = data.mainHeaders;
const colIdx = (letter) => {
  let c = 0;
  for (const ch of letter) c = c * 26 + (ch.charCodeAt(0) - 64);
  return c - 1;
};
const NUM_COLS = new Set(["E", "F", "G", "H", "I", "AA", "AB", "AG", "AH", "AY", "AZ", "BA", "BB"]);
const COMPUTED_HEADERS = COMPUTED_COLS.map((c) => HEADERS[colIdx(c)]);
const cIdx = (letter) => COMPUTED_COLS.indexOf(letter); // 自动计算区内的列索引

function blankRow() {
  const input = {};
  for (const c of INPUT_COLS) input[c] = "";
  const manual = {};
  for (const c of MANUAL_COLS) manual[c] = "";
  return { input, manual, overridden: {}, values: {}, errors: {} };
}

const samples = [
  { A: "3号车间", B: "操作工", C: "投料", D: "二氧化钛粉尘(总尘)", E: "4", F: "8", G: "6", H: "4", I: "1", P: "是" },
  { A: "3号车间", B: "操作工", C: "投料", D: "硫酸钡", P: "是" },
  { A: "3号车间", B: "操作工", C: "投料", D: "丙烯酸", P: "是" },
  { D: "苯乙烯", P: "是" },
  { A: "3号车间", B: "操作工", C: "投料", D: "噪声", P: "是" },
  { A: "3号车间", B: "操作工", C: "放料", D: "丙烯酸", E: "1", P: "是" },
  { D: "苯乙烯", P: "是" },
  { A: "3号车间", B: "操作工", C: "放料", D: "噪声", P: "是" },
];
const rows = samples.map((s) => {
  const r = blankRow();
  for (const k of Object.keys(s)) r.input[k] = s[k];
  return r;
});
// 制造一处覆盖：第 4 行检测方式覆盖为个体
rows[3].values.AR = "个体";
rows[3].overridden.AR = true;
logic.computeRows(rows, data);

// ---------- 导出：仅自动计算区（W~BI 39 列） ----------
const mainRows = [COMPUTED_HEADERS];
for (const r of rows) {
  const line = [];
  for (const c of COMPUTED_COLS) {
    if (MANUAL_COLS.includes(c)) line.push(r.manual[c]);
    else line.push(NUM_COLS.has(c) ? logic.toNum(r.values[c]) : r.values[c]);
  }
  mainRows.push(line);
}
const bytes = await xlsxio.writeWorkbook({ mainRows, mainWidths: Array(39).fill(10) });
await fs.promises.writeFile("tests/out.xlsx", bytes);
console.log("导出字节数:", bytes.length);

// ---------- 用 Excel 引擎验证导出文件 ----------
const input = await FileBlob.load("tests/out.xlsx");
const wb = await SpreadsheetFile.importXlsx(input);
const names = wb.worksheets.items.map((s) => s.name);
console.log("工作表:", names.join(" / "));
if (names.join() !== "测点布局情况调查") throw new Error("应只导出自动计算区单张工作表");

const main = wb.worksheets.getItem("测点布局情况调查");
const hdr = (await main.getRange("A1:AM1").values)[0];
if (hdr.length !== 39) throw new Error("导出表头应为 39 列: " + hdr.length);
if (String(hdr[0]) !== "*单元/工作场所" || String(hdr[cIdx("AN")]) !== "*检测项目") throw new Error("导出表头不符");

const r2 = (await main.getRange("A2:AM2").values)[0];
const r5 = (await main.getRange("A5:AM5").values)[0];
const r9 = (await main.getRange("A9:AM9").values)[0];
const norm = (v) => (typeof v === "number" ? String(v) : String(v ?? ""));
const check = (row, col, expect, label) => {
  if (norm(row[cIdx(col)]) !== norm(expect)) throw new Error(`${label}：期望 ${expect}，实际 ${row[cIdx(col)]}`);
};
check(r2, "W", "3号车间", "W 单元");
check(r2, "AN", "二氧化钛粉尘", "AN 检测项目");
check(r2, "AU", "总尘", "AU 粉尘性质");
check(r2, "BB", 3, "BB 每天样品数");
check(r2, "AB", 4, "AB 每班最大人数");
check(r2, "AM", "操作设备投料作业", "AM 工作内容");
check(r9, "AN", "噪声", "AN9");
check(r9, "BD", "设备运行", "BD9 危害因素来源");
check(r9, "BB", 3, "BB9");
check(r9, "AZ", 1, "AZ9 日接触时长");
check(r5, "AR", "个体", "AR5 覆盖值");
check(r5, "AS", "长时间", "AS5 采样时间类型(覆盖联动)");
check(r5, "AK", "采样对象", "AK5 采样类别(覆盖联动)");
console.log("artifact-tool 校验通过 ✔");

// ---------- 回读（导出→导入）往返一致性 ----------
const sheets = await xlsxio.readWorkbook(bytes);
if (sheets.length !== 1 || sheets[0].name !== "测点布局情况调查") throw new Error("回读工作表不符");
const mainArr = xlsxio.sheetToArray(sheets[0]);
if (mainArr.length !== rows.length + 1) throw new Error("回读行数不符");

// 模拟“仅自动计算区”导入：无录入区表头 → 全部计算列静态保留
const idx = {};
mainArr[0].forEach((h, i) => { if (h !== "" && h !== null && h !== undefined) idx[String(h).trim()] = i; });
const hasInput = ["车间", "接害因素", "接触时间h/d", "人数"].some((h) => h in idx);
if (hasInput) throw new Error("自动计算区文件不应包含录入区表头");

const importedRows = mainArr.slice(1).map((r) => {
  const row = blankRow();
  const vals = {};
  for (const c of MANUAL_COLS) {
    const h = COMPUTED_HEADERS[cIdx(c)];
    if (h in idx) row.manual[c] = String(r[idx[h]] ?? "");
  }
  for (const c of COMPUTED_COLS) {
    const h = COMPUTED_HEADERS[cIdx(c)];
    if (h in idx) vals[c] = String(r[idx[h]] ?? "");
  }
  row.__vals = vals;
  return row;
});
logic.computeRows(importedRows, data);
importedRows.forEach((row, i) => {
  for (const c of COMPUTED_COLS) {
    const iv = row.__vals[c];
    if (iv !== undefined && iv !== "" && iv !== row.values[c]) {
      row.values[c] = iv;
      row.overridden[c] = true;
    } else {
      delete row.overridden[c];
    }
  }
  delete row.__vals;
});
logic.computeRows(importedRows, data);

for (let i = 0; i < rows.length; i++) {
  for (const c of COMPUTED_COLS) {
    const a = String(importedRows[i].values[c] ?? "");
    const b = String(rows[i].values[c] ?? "");
    if (a !== b) throw new Error(`往返不一致 行${i + 1} ${c}: ${a} vs ${b}`);
  }
}
console.log("往返一致性校验通过 ✔（8 行 × 39 列）");

// ---------- 兼容性：完整结构文件（含录入区）导入路径 ----------
const fullRows = [HEADERS];
for (const r of rows) {
  const line = [];
  for (const c of [...INPUT_COLS, "V", ...COMPUTED_COLS]) {
    if (c === "V") { line.push(""); continue; }
    if (INPUT_COLS.includes(c)) line.push(NUM_COLS.has(c) ? logic.toNum(r.input[c]) : r.input[c]);
    else if (MANUAL_COLS.includes(c)) line.push(r.manual[c]);
    else line.push(NUM_COLS.has(c) ? logic.toNum(r.values[c]) : r.values[c]);
  }
  fullRows.push(line);
}
const fullBytes = await xlsxio.writeWorkbook({
  hazardRows: [["识别"], ["测试"]],
  mainRows: fullRows,
  itemRows: [["系统内检测项目名参照表"], ["测试项"]],
});
const fullSheets = await xlsxio.readWorkbook(fullBytes);
if (fullSheets.map((s) => s.name).join() !== "危害因素,测点布局情况调查,检测项目") {
  throw new Error("完整结构导出失败");
}
console.log("完整三表结构兼容导出通过 ✔");
