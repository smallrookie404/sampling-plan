/*
 * 计算引擎语义等价校验（非正式测试，文件名以下划线开头）。
 * 用随机数据同时跑「git HEAD 版本」与「工作区版本」的 computeRows，
 * 逐行逐列比较 values / errors，确保性能优化没有改变任何计算结果。
 *
 * 用法：node tests/_diff_logic.mjs [样本行数]
 */
import fs from "node:fs";
import os from "node:os";
import path from "node:path";
import vm from "node:vm";
import { execSync } from "node:child_process";
import { fileURLToPath } from "node:url";
import { createRequire } from "node:module";

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const APP = path.join(__dirname, "..");
const require = createRequire(import.meta.url);

const newLogic = require(path.join(APP, "js/logic.js"));

// 旧版：从 git HEAD 取出到临时文件后 require
const oldSrc = execSync("git show HEAD:sampling-plan-app/js/logic.js", { cwd: APP, encoding: "utf8", maxBuffer: 1 << 26 });
const tmp = path.join(os.tmpdir(), "logic_head_" + process.pid + ".cjs");
fs.writeFileSync(tmp, oldSrc, "utf8");
const oldLogic = require(tmp);
fs.unlinkSync(tmp);

// 内置危害因素库
const libCtx = {};
vm.createContext(libCtx);
vm.runInContext(fs.readFileSync(path.join(APP, "js/data-library.js"), "utf8"), libCtx);
const hazardFactors = libCtx.SamplingLibrary.hazardFactors;

// 确定性伪随机（便于复现）
let seed = 20261029;
const rnd = () => {
  seed = (seed * 1103515245 + 12345) & 0x7fffffff;
  return seed / 0x7fffffff;
};
const pick = (arr) => arr[Math.floor(rnd() * arr.length) % arr.length];
const chance = (p) => rnd() < p;

const N = Number(process.argv[2] || 3000);
const WORDS = ["甲苯", "噪声", "高温", "粉尘", "电焊烟尘", "游离二氧化硅", "手传振动", "工频电场", "紫外辐射", "氯化氢"];
const SHOPS = ["一车间", "二车间", "动力车间"];
const POSTS = ["操作工", "班长", "巡检工", "包装工"];
const POINTS = ["投料口", "反应釜区", "成品包装", "泵房", "配电室"];
const SHIFTS = ["", "单班", "轮班"];
const SHIFTS2 = ["", "两班两运转", "三班两运转", "三班三运转", "四班三运转", "五班三运转", "长白班", "其他"];
const LDS = ["", "Ⅰ", "Ⅱ", "Ⅲ", "Ⅳ"];
const GANGWEI = ["", "固定", "流动", "巡检"];
const YESNO = ["", "是", "否"];
const NORM = ["", "定点", "个体"];
const WORK = ["", "手工作业", "半手工作业", "自动化作业"];

function makeRows(n) {
  const rows = [];
  for (let i = 0; i < n; i++) {
    const input = {};
    for (const c of newLogic.INPUT_COLS) input[c] = "";
    const manual = {};
    for (const c of newLogic.MANUAL_COLS) manual[c] = "";
    // 前几列按「只保留变化处」习惯稀疏填写，贴近真实导入数据
    if (chance(0.25) || i === 0) input.A = pick(SHOPS);
    if (chance(0.3)) input.B = pick(POSTS);
    if (chance(0.35)) input.C = pick(POINTS);
    if (chance(0.9)) input.D = pick(WORDS);
    if (chance(0.4)) input.E = String(pick(["0.5", "1", "2", "4", ""]));
    if (chance(0.3)) input.F = String(pick(["4", "8", "12", ""]));
    if (chance(0.3)) input.G = String(pick(["4", "5", "6", ""]));
    if (chance(0.3)) input.H = String(pick(["1", "2", "3", "10", ""]));
    if (chance(0.3)) input.I = pick(["1", "2", "3", ""]);
    if (chance(0.2)) input.J = pick(["巡检设备", "投料作业", ""]);
    if (chance(0.2)) input.K = pick(["原辅物料", "设备运行", ""]);
    if (chance(0.2)) input.L = pick(WORK);
    if (chance(0.15)) input.M = pick(YESNO);
    if (chance(0.15)) input.N = pick(NORM);
    if (chance(0.15)) input.O = pick(GANGWEI);
    if (chance(0.15)) input.P = pick(YESNO);
    if (chance(0.1)) input.Q = pick(YESNO);
    if (chance(0.3)) input.R = pick(SHIFTS);
    if (chance(0.3)) input.S = pick(SHIFTS2);
    if (chance(0.1)) input.T = pick(["", "1", "2"]);
    if (chance(0.1)) input.U = pick(LDS);
    if (chance(0.1)) input.V = pick(["", "已停产", "备注内容"]);
    if (chance(0.05)) input.D = "库里没有的因素"; // 未命中分支
    if (chance(0.05)) input.D = "滑石粉尘(游离SiO2含量＜10%)";
    if (chance(0.05)) input.Y = ""; // 录入区无 Y，走计算区覆盖
    const overridden = {};
    const values = {};
    if (chance(0.08)) { overridden.Y = true; values.Y = pick(["固定", "流动", "巡检"]); }
    if (chance(0.08)) { overridden.Z = true; values.Z = pick(["浓度（或强度）相对稳定", "浓度（或强度）相对不稳定"]); }
    if (chance(0.08)) { overridden.AO = true; values.AO = pick(["是", "否"]); }
    if (chance(0.08)) { overridden.AR = true; values.AR = pick(["定点", "个体"]); }
    if (chance(0.08)) { overridden.BI = true; values.BI = "手工备注"; }
    if (chance(0.1)) { overridden.W = true; values.W = "旧车间"; } // 任意列可被覆盖
    for (const c of newLogic.MANUAL_COLS) if (chance(0.1)) manual[c] = pick(["1", "2", "是", "否", ""]);
    rows.push({ input, manual, overridden, values, errors: {} });
  }
  return rows;
}

const base = makeRows(N);
const clone = (rows) => rows.map((r) => ({
  input: { ...r.input },
  manual: { ...r.manual },
  overridden: { ...r.overridden },
  values: { ...r.values },
  errors: {},
}));

const a = clone(base);
const b = clone(base);
oldLogic.computeRows(a, { hazardFactors });
newLogic.computeRows(b, { hazardFactors });

let diff = 0;
const samples = [];
const cols = new Set();
for (let i = 0; i < a.length; i++) {
  for (const key of new Set([...Object.keys(a[i].values), ...Object.keys(b[i].values)])) {
    if (String(a[i].values[key] ?? "") !== String(b[i].values[key] ?? "")) {
      diff++;
      cols.add("values." + key);
      if (samples.length < 8) samples.push(`行${i} values.${key}: 旧「${a[i].values[key]}」 新「${b[i].values[key]}」`);
    }
  }
  for (const key of new Set([...Object.keys(a[i].errors), ...Object.keys(b[i].errors)])) {
    if (String(a[i].errors[key] ?? "") !== String(b[i].errors[key] ?? "")) {
      diff++;
      cols.add("errors." + key);
      if (samples.length < 8) samples.push(`行${i} errors.${key}: 旧「${a[i].errors[key]}」 新「${b[i].errors[key]}」`);
    }
  }
}

console.log(`样本 ${N} 行 · 差异 ${diff} 处` + (cols.size ? ` · 涉及列：${[...cols].join(",")}` : ""));
for (const s of samples) console.log("  " + s);
process.exit(diff === 0 ? 0 : 1);
