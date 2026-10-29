/*
 * 导入性能基准（非正式测试，文件名以下划线开头）。
 * 从 js/app.js 源码中提取「接害因素解析」相关函数，模拟一次大批量导入的纯 JS 阶段，
 * 打印各阶段耗时，用于优化前后对比。
 *
 * 用法：node tests/_bench_import.mjs [行数]
 */
import fs from "node:fs";
import path from "node:path";
import vm from "node:vm";
import { fileURLToPath } from "node:url";
import { createRequire } from "node:module";

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const APP = path.join(__dirname, "..");
const require = createRequire(import.meta.url);
const L = require(path.join(APP, "js/logic.js"));

// 内置危害因素库
const libCtx = { module: undefined, exports: undefined };
vm.createContext(libCtx);
vm.runInContext(fs.readFileSync(path.join(APP, "js/data-library.js"), "utf8"), libCtx);
const hazardFactors = libCtx.SamplingLibrary.hazardFactors;

// ---------- 从 app.js 提取函数（按函数名 + 花括号平衡） ----------
const appSrc = fs.readFileSync(path.join(APP, "js/app.js"), "utf8");
function grab(name) {
  const key = "function " + name + "(";
  const at = appSrc.indexOf(key);
  if (at < 0) throw new Error("未找到函数：" + name);
  const open = appSrc.indexOf("{", at);
  let depth = 0;
  for (let i = open; i < appSrc.length; i++) {
    const ch = appSrc[i];
    if (ch === "{") depth++;
    else if (ch === "}") {
      depth--;
      if (depth === 0) return appSrc.slice(at, i + 1);
    }
  }
  throw new Error("花括号不平衡：" + name);
}

// 提取单行常量声明（如 const HAZARD_KEY_RE = /.../;）
function grabConst(name) {
  const key = "const " + name + " =";
  const at = appSrc.indexOf(key);
  if (at < 0) throw new Error("未找到常量：" + name);
  const end = appSrc.indexOf(";", at);
  if (end < 0) throw new Error("常量未闭合：" + name);
  return appSrc.slice(at, end + 1);
}

const sandbox = { hazardFactors, console };
vm.createContext(sandbox);
vm.runInContext(
  [
    grabConst("HAZARD_KEY_MAP"),
    grabConst("HAZARD_KEY_RE"),
    grab("normHazardKey"),
    grab("buildHazardIndex"),
    grab("hazardDustKey"),
    grab("resolveHazardRec"),
    // 供基准调用
    "globalThis.__norm = normHazardKey;",
    "globalThis.__index = buildHazardIndex;",
    "globalThis.__resolve = resolveHazardRec;",
  ].join("\n"),
  sandbox
);
const resolveHazardRec = sandbox.__resolve;
const buildHazardIndex = sandbox.__index;

// ---------- 构造接近真实的数据 ----------
const N = Number(process.argv[2] || 3000);
// 取库里 40 个不同危害因素循环使用（真实表里通常几十种，粉尘类占比高）
const sampleRecs = hazardFactors.filter((h) => h.name === "噪声" || h.name === "粉尘" || h.rec.includes("粉尘")).slice(0, 20);
const recs = (sampleRecs.length ? sampleRecs : hazardFactors.slice(0, 20)).map((h) => h.rec);
const recNames = (sampleRecs.length ? sampleRecs : hazardFactors.slice(0, 20)).map((h) => h.name);
const dusts = ["总尘", "呼尘"];
const blankRow = () => {
  const input = {};
  for (const c of L.INPUT_COLS) input[c] = "";
  const manual = {};
  for (const c of L.MANUAL_COLS) manual[c] = "";
  return { input, manual, overridden: {}, values: {}, errors: {} };
};

const t = {};
const now = () => Number(process.hrtime.bigint() / 1000n) / 1000; // ms
let t0 = now();

const rows = [];
for (let i = 0; i < N; i++) {
  const row = blankRow();
  row.input.A = "车间" + Math.floor(i / 300);
  row.input.B = "岗位" + Math.floor(i / 30);
  row.input.C = "点位" + (i % 30);
  row.input.D = recNames[i % recNames.length];
  row.input.E = "2";
  row.input.F = "8";
  row.input.G = "5";
  row.input.H = "3";
  row.input.I = "1";
  row.input.S = "三班两运转";
  rows.push(row);
}
t.构建行 = now() - t0;

t0 = now();
const hIndex = buildHazardIndex();
for (const row of rows) {
  if (row.input.D) {
    row.input.D = resolveHazardRec(row.input.D, dusts[Number(row.input.A.slice(2)) % 2], hIndex);
  }
}
t.解析危害因素 = now() - t0;

t0 = now();
L.computeRows(rows, { hazardFactors });
t.首次计算 = now() - t0;

t0 = now();
L.computeRows(rows, { hazardFactors });
t.二次计算 = now() - t0;

t0 = now();
let acc = 0;
for (const row of rows) acc += Object.keys(row.values).length + Object.keys(row.errors).length;
t.汇总检查 = now() - t0;

const total = Object.values(t).reduce((a, b) => a + b, 0);
console.log(`行数 ${N} · 库 ${hazardFactors.length} 条`);
for (const [k, v] of Object.entries(t)) console.log(`  ${k.padEnd(8, "　")} ${v.toFixed(1)} ms`);
console.log(`  ${"合计".padEnd(8, "　")} ${total.toFixed(1)} ms`);
