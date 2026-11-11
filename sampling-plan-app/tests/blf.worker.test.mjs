// 回归测试：网站数据导入时，自动计算区的「流量」「采样时长」不导入内容（用户 2026-11-11）。
//
// 需求边界：
//   · 「网站数据导入」（采样方案导入）走 blf.js workerTo39：上游「劳动定员」xlsx（43 列）
//     → 导入模板主表 39 列（W..BI）；
//   · 其中「流量」(BF) / 「采样时长」(BG) 属 logic.js MANUAL_COLS，导入时若有值会被当作
//     本行手工值锁定，不再随参数联动 —— 故**只清空这两列的单元格内容**；
//   · 表头与 39 列结构保持不动（importWorkbook 按表头对齐，列序不能少）；
//   · 其它列（含同为手工列的 是否接害 BH、备注 BI）照常导入。
//
// 覆盖：
//   A. workerTo39：两列清空 / 其它列原样 / 数字与 0 值 / 空行跳过 / 不修改入参
//   B. 上游列多于 39 列（真实 43 列）时按表头名对齐，仍只清这两列
//   C. 上游缺这两列表头时不报错、其它列不受影响
//   D. buildTemplateSheets 端到端：主表 sheet 的 BF/BG 单元格为空，表头仍是 39 列
//
// 用法：node tests/blf.worker.test.mjs
import fs from "node:fs";
import path from "node:path";
import vm from "node:vm";
import { fileURLToPath } from "node:url";

const here = path.dirname(fileURLToPath(import.meta.url));
const APP = path.resolve(here, "..");

let pass = 0, fail = 0;
function check(name, cond, extra) {
  if (cond) { pass++; console.log("  ok  " + name); }
  else { fail++; console.log("  FAIL " + name + (extra ? " :: " + extra : "")); }
}

const noop = () => {};
const fakeEl = { style: {}, classList: { add: noop, remove: noop, toggle: noop, addEventListener: noop },
  appendChild: noop, removeChild: noop, setAttribute: noop, getAttribute: () => null, hasAttribute: () => false,
  querySelector: () => null, querySelectorAll: () => [], innerHTML: "", textContent: "", disabled: false, click: noop };
const documentStub = { readyState: "complete", getElementById: () => null, querySelector: () => null,
  querySelectorAll: () => [], createElement: () => Object.assign({}, fakeEl), addEventListener: noop,
  body: Object.assign({}, fakeEl), head: Object.assign({}, fakeEl), documentElement: Object.assign({}, fakeEl) };

function loadBlf() {
  const blfSrc = fs.readFileSync(path.join(APP, "js/blf.js"), "utf8");
  const sandbox = { console, document: documentStub, setTimeout, clearTimeout,
    fetch: () => Promise.reject(new Error("no fetch")),
    URL: { createObjectURL: () => "blob:x", revokeObjectURL: noop }, TextEncoder, TextDecoder };
  sandbox.self = sandbox; sandbox.window = sandbox;
  sandbox.SamplingApp = { librarySnapshot: () => ({}) };
  vm.createContext(sandbox);
  vm.runInContext(blfSrc, sandbox, { filename: "blf.js" });
  return sandbox.SamplingBlf;
}

const blf = loadBlf();

// 表头直接取自 buildTemplateSheets 的主表（避免在测试里重复抄 39 列）
function workerHeaders() {
  const sheets = blf.buildTemplateSheets({ worker39: [] }, { worker: true, survey: false, dicts: false });
  return sheets[0].rows[0].slice();
}
const HDR = workerHeaders();
const IX_FLOW = HDR.indexOf("流量");
const IX_DUR = HDR.indexOf("采样时长");
const IX_HARM = HDR.indexOf("是否接害");
const IX_NOTE = HDR.indexOf("备注");

console.log("== A. workerTo39 基本行为 ==");
check("表头 39 列，含「流量」「采样时长」", HDR.length === 39 && IX_FLOW >= 0 && IX_DUR >= 0,
  "len=" + HDR.length + " flow=" + IX_FLOW + " dur=" + IX_DUR);

// 每列填 C<索引>，流量 / 采样时长填可识别的大值
const rowVals = HDR.map((_, i) => "C" + i);
rowVals[IX_FLOW] = "20.0";
rowVals[IX_DUR] = "480";
rowVals[IX_HARM] = "是";
rowVals[IX_NOTE] = "平台备注";
const aoa = [HDR.slice(), rowVals.slice()];
const before = JSON.stringify(aoa);
const out = blf.workerTo39(aoa);

check("输出行数 1 行、宽度 39", out.length === 1 && out[0].length === 39);
check("流量列内容被清空", out[0][IX_FLOW] === "", JSON.stringify(out[0][IX_FLOW]));
check("采样时长列内容被清空", out[0][IX_DUR] === "", JSON.stringify(out[0][IX_DUR]));
check("其它列原样保留", out[0][0] === "C0" && out[0][34] === "C34" && out[0][33] === "C33",
  JSON.stringify([out[0][0], out[0][34], out[0][33]]));
check("同为手工列的「是否接害」照常导入", out[0][IX_HARM] === "是", JSON.stringify(out[0][IX_HARM]));
check("「备注」照常导入", out[0][IX_NOTE] === "平台备注", JSON.stringify(out[0][IX_NOTE]));
check("不修改传入数组（无副作用）", JSON.stringify(aoa) === before);

// 数值 0 / 空字符串
const zeroRow = rowVals.slice();
zeroRow[IX_FLOW] = 0;
zeroRow[IX_DUR] = "";
const outZero = blf.workerTo39([HDR.slice(), zeroRow.slice()]);
check("流量为 0 时同样不导入", outZero[0][IX_FLOW] === "");
check("采样时长空值时保持空", outZero[0][IX_DUR] === "");

// 空行跳过
const emptyRow = HDR.map(() => "");
const outEmpty = blf.workerTo39([HDR.slice(), emptyRow.slice(), rowVals.slice()]);
check("纯空行被跳过", outEmpty.length === 1);
check("无表头 / 空输入返回空数组",
  blf.workerTo39([]).length === 0 && blf.workerTo39(null).length === 0);

console.log("== B. 上游 43 列（多出无关列）按表头对齐 ==");
const extra = ["项目名称", "企业名称", "备注2", "流程状态"];
const hdr43 = HDR.slice(0, 5).concat(extra, HDR.slice(5));
const row43 = hdr43.map((h) => "X:" + h);
const at = (h) => hdr43.indexOf(h);
row43[at("流量")] = "3.5";
row43[at("采样时长")] = "120";
row43[at("*单元/工作场所")] = "一车间";
const out43 = blf.workerTo39([hdr43.slice(), row43.slice()]);
check("43 列输入 → 仍输出 39 列", out43.length === 1 && out43[0].length === 39);
check("43 列下流量仍清空", out43[0][IX_FLOW] === "");
check("43 列下采样时长仍清空", out43[0][IX_DUR] === "");
check("43 列下单元列对齐正确", out43[0][0] === "一车间", JSON.stringify(out43[0][0]));
check("43 列下岗位列对齐正确", out43[0][1] === "X:*岗位/工种", JSON.stringify(out43[0][1]));

console.log("== C. 上游缺这两列表头 ==");
const keep = HDR.map((h, i) => i).filter((i) => i !== IX_FLOW && i !== IX_DUR);
const hdrNo = keep.map((i) => HDR[i]);
const rowNo = keep.map((i) => "V" + i);
const outNo = blf.workerTo39([hdrNo.slice(), rowNo.slice()]);
check("缺列时输出仍为 39 列（缺的列补空）", outNo[0].length === 39);
check("缺列时流量列为空", outNo[0][IX_FLOW] === "");
check("缺列时采样时长列为空", outNo[0][IX_DUR] === "");
check("缺列时其它列正确", outNo[0][0] === "V0" && outNo[0][38] === "V38");

console.log("== D. buildTemplateSheets 端到端 ==");
const up39 = [rowVals.slice(), rowVals.slice()];
const sheets = blf.buildTemplateSheets({ worker39: up39 }, { worker: true, survey: false, dicts: false });
const main = sheets[0];
check("主表 sheet 存在且表头 39 列", !!main && main.rows[0].length === 39);
check("端到端：数据行流量为空", main.rows[1][IX_FLOW] === "" && main.rows[2][IX_FLOW] === "");
check("端到端：数据行采样时长为空", main.rows[1][IX_DUR] === "" && main.rows[2][IX_DUR] === "");
check("端到端：其它列仍有值", main.rows[1][0] === "C0" && main.rows[2][34] === "C34");

console.log("\n结果： " + pass + " 通过 / " + fail + " 失败");
if (fail) process.exit(1);
