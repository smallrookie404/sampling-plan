/*
 * xlsx 读写耗基准（非正式测试，文件名以下划线开头）。
 * 导入第一步是解析工作簿：这里测「生成 + 解析」3000 行的耗时，
 * 判断解析环节是否是导入卡顿的瓶颈。
 *
 * 用法：node tests/_bench_xlsx.mjs [行数]
 */
import fs from "node:fs";
import os from "node:os";
import path from "node:path";
import vm from "node:vm";
import { fileURLToPath } from "node:url";
import { createRequire } from "node:module";

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const APP = path.join(__dirname, "..");
const require = createRequire(import.meta.url);
const X = require(path.join(APP, "js/xlsxio.js"));

const libCtx = {};
vm.createContext(libCtx);
vm.runInContext(fs.readFileSync(path.join(APP, "js/data-core.js"), "utf8"), libCtx);
const headers = libCtx.SamplingData.mainHeaders;

const N = Number(process.argv[2] || 3000);
const now = () => Number(process.hrtime.bigint() / 1000n) / 1000;

const rows = [headers];
for (let i = 0; i < N; i++) {
  const row = [];
  for (let c = 0; c < headers.length; c++) {
    // 模拟真实数据：文字列填中文、数字列填数字、其余空
    if (c < 8) row.push(["一车间", "操作工", "投料口", "噪声", "2", "8", "5", "3"][c]);
    else row.push(c % 7 === 0 ? String(i % 20) : "");
  }
  rows.push(row);
}

let t0 = now();
const bytes = await X.writeWorkbook([{ name: "测点布局情况调查", rows }]);
const tWrite = now() - t0;

const file = path.join(os.tmpdir(), `bench_import_${process.pid}.xlsx`);
fs.writeFileSync(file, bytes);

t0 = now();
const sheets = await X.readWorkbook(fs.readFileSync(file));
const tZip = now() - t0;

t0 = now();
const arr = X.sheetToArray(sheets[0]);
const tArr = now() - t0;

fs.unlinkSync(file);

// 往返一致性：写出的每一格读回来必须与原始值相同（空值读回空串）
let mismatch = 0;
for (let r = 0; r < rows.length && r < arr.length; r++) {
  for (let c = 0; c < headers.length; c++) {
    const a = rows[r][c] === null || rows[r][c] === undefined ? "" : String(rows[r][c]);
    const b = arr[r][c] === null || arr[r][c] === undefined ? "" : String(arr[r][c]);
    if (a !== b) {
      mismatch++;
      if (mismatch <= 5) console.log(`  ✗ 往返不一致 r${r} c${c}: 原「${a}」读回「${b}」`);
    }
  }
}

console.log(`行数 ${N} · 列 ${headers.length} · 文件 ${(bytes.length / 1024).toFixed(0)} KB`);
console.log(`  生成 xlsx      ${tWrite.toFixed(0)} ms`);
console.log(`  解析 xlsx      ${tZip.toFixed(0)} ms`);
console.log(`  转二维数组     ${tArr.toFixed(0)} ms  （${arr.length} 行）`);
console.log(`  往返一致性     ${mismatch === 0 ? "全部一致 ✔" : mismatch + " 处不一致 ✗"}`);
process.exit(mismatch === 0 ? 0 : 1);
