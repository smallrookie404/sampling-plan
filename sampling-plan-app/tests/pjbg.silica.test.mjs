// 回归测试：表4-1「职业病危害因素」列删除「游离二氧化硅」（用户 2026-11-12）。
//
// 需求边界：
//   · 「游离二氧化硅」不是职业病危害因素（是粉尘的组成检测项），该列里两种形态都去掉：
//       ① 单独一项：`二甲苯,游离二氧化硅,噪声` → `二甲苯,噪声`
//       ② 括号说明：`滑石粉尘(游离SiO2含量<10%)` → `滑石粉尘`
//   · **只删这一项，不能连累后面的因素**（用户明确：之前把「）」及后面的内容都删了）；
//   · 只动表4-1（表头同时含「职业病危害因素」+「工作地点」列）；表5-1 检测项目确认表不动；
//   · 括号内不以「游离」开头的（如「矽尘(10%≤游离SiO2含量≤50%)」）保留，含量区间是分类依据。
//
// 用法：node tests/pjbg.silica.test.mjs
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
const fakeEl = { style: {}, classList: { add: noop, remove: noop, toggle: noop }, addEventListener: noop,
  appendChild: noop, removeChild: noop, setAttribute: noop, getAttribute: () => null, hasAttribute: () => false,
  querySelector: () => null, querySelectorAll: () => [], innerHTML: "", textContent: "", disabled: false, click: noop };
const documentStub = { readyState: "complete", getElementById: () => null, querySelector: () => null,
  querySelectorAll: () => [], createElement: () => Object.assign({}, fakeEl), addEventListener: noop,
  body: Object.assign({}, fakeEl), head: Object.assign({}, fakeEl), documentElement: Object.assign({}, fakeEl) };

function loadBlf() {
  const src = fs.readFileSync(path.join(APP, "js/blf.js"), "utf8");
  const sandbox = { console, document: documentStub, setTimeout, clearTimeout,
    fetch: () => Promise.reject(new Error("no fetch")),
    URL: { createObjectURL: () => "blob:x", revokeObjectURL: noop }, TextEncoder, TextDecoder };
  sandbox.self = sandbox; sandbox.window = sandbox;
  sandbox.SamplingApp = { librarySnapshot: () => ({}) };
  vm.createContext(sandbox);
  vm.runInContext(src, sandbox, { filename: "blf.js" });
  return sandbox.SamplingBlf;
}

const blf = loadBlf();
const strip = blf.stripFreeSilicaText;

console.log("== A. 纯文本：括号说明 ==");
check("括号内含「游离SiO2含量<10%」→ 只留粉尘名",
  strip("二甲苯,滑石粉尘(游离SiO2含量<10%),乙苯,乙酸丁酯,游离二氧化硅,噪声")
    === "二甲苯,滑石粉尘,乙苯,乙酸丁酯,噪声",
  JSON.stringify(strip("二甲苯,滑石粉尘(游离SiO2含量<10%),乙苯,乙酸丁酯,游离二氧化硅,噪声")));
check("全角括号 + 空格差异",
  strip("二甲苯，滑石粉尘（游离 SiO2 含量<10%），噪声") === "二甲苯，滑石粉尘，噪声",
  JSON.stringify(strip("二甲苯，滑石粉尘（游离 SiO2 含量<10%），噪声")));
check("写法「游离二氧化硅含量<10%」也识别",
  strip("滑石粉尘(游离二氧化硅含量<10%),噪声") === "滑石粉尘,噪声",
  JSON.stringify(strip("滑石粉尘(游离二氧化硅含量<10%),噪声")));
check("小写 sio2 也识别",
  strip("滑石粉尘(游离 sio2 含量<10%)") === "滑石粉尘");
check("括号内有全角空格 / 多余空格",
  strip("滑石粉尘 ( 游离　SiO2 含量<10% ) ,噪声") === "滑石粉尘,噪声",
  JSON.stringify(strip("滑石粉尘 ( 游离　SiO2 含量<10% ) ,噪声")));

console.log("== B. 纯文本：独立一项 ==");
check("中间一项删除，后面因素保留",
  strip("二甲苯,游离二氧化硅,噪声") === "二甲苯,噪声",
  JSON.stringify(strip("二甲苯,游离二氧化硅,噪声")));
check("末尾一项删除，不留尾逗号",
  strip("二甲苯,游离二氧化硅") === "二甲苯",
  JSON.stringify(strip("二甲苯,游离二氧化硅")));
check("只有这一项 → 占位「-」", strip("游离二氧化硅") === "-", JSON.stringify(strip("游离二氧化硅")));
check("前面有空格 / 全角逗号",
  strip("二甲苯， 游离二氧化硅 ，噪声") === "二甲苯，噪声",
  JSON.stringify(strip("二甲苯， 游离二氧化硅 ，噪声")));

console.log("== C. 不该动的形态 ==");
check("「矽尘(10%≤游离SiO2含量≤50%)」括号保留（区间是分类依据）",
  strip("矽尘(10%≤游离SiO2含量≤50%),游离二氧化硅,噪声") === "矽尘(10%≤游离SiO2含量≤50%),噪声",
  JSON.stringify(strip("矽尘(10%≤游离SiO2含量≤50%),游离二氧化硅,噪声")));
check("完全不含游离 → null（不改）", strip("二甲苯,乙苯,噪声") === null);
check("含「游离」但非目标 → null", strip("二甲苯,游离态物质,噪声") === null);
check("空值 / null → null", strip("") === null && strip(null) === null && strip(undefined) === null);

console.log("== D. XML：表4-1 端到端 ==");
const H = ["单元/工作场所", "岗位（工种）", "工作地点及内容", "职业病危害因素", "人数"];
function run(t) { return `<w:r><w:t xml:space="preserve">${t}</w:t></w:r>`; }
function tc(inner) { return `<w:tc><w:p>${inner}</w:p></w:tc>`; }
function row(cells) { return `<w:tr>${cells.map(tc).join("")}</w:tr>`; }
function table(heads, rows) { return `<w:tbl>${row(heads.map(run))}${rows.map((r) => row(r.map(run))).join("")}</w:tbl>`; }

const hazardText = "二甲苯,滑石粉尘(游离SiO2含量&lt;10%),乙苯,乙酸丁酯,游离二氧化硅,噪声";
const xml41 = '<w:document><w:body>' + table(H, [
  ["一车间", "刷腻子/人工进行刷腻子作业", "刷腻子", hazardText, "2"],
  ["一车间", "喷胶区/操作设备进行喷胶作业", "喷胶区", "丙酮,二甲苯,噪声", "1"],
]) + '</w:body></w:document>';

const r1 = blf.stripFreeSilica(xml41);
check("表4-1：命中行被改写", r1.changed === 1, "changed=" + r1.changed);
check("表4-1：结果为「二甲苯,滑石粉尘,乙苯,乙酸丁酯,噪声」",
  r1.xml.indexOf(">二甲苯,滑石粉尘,乙苯,乙酸丁酯,噪声<") >= 0, r1.xml.slice(r1.xml.indexOf("二甲苯"), r1.xml.indexOf("二甲苯") + 80));
check("表4-1：后面的因素（乙苯/乙酸丁酯/噪声）都在", r1.xml.indexOf("乙酸丁酯") >= 0 && r1.xml.indexOf("噪声") >= 0);
check("表4-1：不再含「游离」", r1.xml.indexOf("游离") < 0);
check("表4-1：未命中行原样", r1.xml.indexOf(">丙酮,二甲苯,噪声<") >= 0);
check("表4-1：XML 结构完整（<w:t> 数量不变）",
  (xml41.match(/<w:t[\s>]/g) || []).length === (r1.xml.match(/<w:t[\s>]/g) || []).length);
check("表4-1：未出现裸 < 号（&lt; 未被破坏）", r1.xml.indexOf("含量<10%") < 0);

const r1b = blf.stripFreeSilica(r1.xml);
check("幂等：第二次不再改动", r1b.changed === 0 && r1b.xml === r1.xml);

console.log("== E. XML：表5-1（无「工作地点」列）不动 ==");
const xml51 = '<w:document><w:body>' + table(
  ["检测项目", "职业病危害因素", "备注"],
  [["游离二氧化硅", "游离二氧化硅", "—"], ["噪声", "噪声", "—"]],
) + '</w:body></w:document>';
const r2 = blf.stripFreeSilica(xml51);
check("表5-1：changed=0", r2.changed === 0, "changed=" + r2.changed);
check("表5-1：内容逐字不变", r2.xml === xml51);

console.log("== F. XML：跨 run / 多段落 ==");
const crossRun = '<w:document><w:body><w:tbl>' +
  row(H.map(run)) +
  '<w:tr>' + tc(run("一车间")) + tc(run("刷腻子/人工进行刷腻子作业")) + tc(run("刷腻子")) +
  tc(`<w:r><w:t xml:space="preserve">二甲苯,滑石粉尘(游离</w:t></w:r><w:r><w:t xml:space="preserve">SiO2含量&lt;10%),游离二氧化硅,噪声</w:t></w:r>`) +
  tc(run("2")) + '</w:tr></w:tbl></w:body></w:document>';
const r3 = blf.stripFreeSilica(crossRun);
check("跨 run：仍能删除括号说明与独立项", r3.xml.indexOf(">二甲苯,滑石粉尘,噪声<") >= 0,
  r3.xml.slice(r3.xml.indexOf("二甲苯"), r3.xml.indexOf("二甲苯") + 60));
check("跨 run：段落数不变", (crossRun.match(/<w:p>/g) || []).length === (r3.xml.match(/<w:p>/g) || []).length);

const multiPara = '<w:document><w:body><w:tbl>' +
  row(H.map(run)) +
  '<w:tr>' + tc(run("一车间")) + tc(run("刷腻子/人工进行刷腻子作业")) + tc(run("刷腻子")) +
  `<w:tc><w:p>${run("二甲苯,乙苯")}</w:p><w:p>${run("游离二氧化硅")}</w:p></w:tc>` +
  tc(run("2")) + '</w:tr></w:tbl></w:body></w:document>';
const r4 = blf.stripFreeSilica(multiPara);
check("多段落：只改含游离的那一段", r4.changed === 1, "changed=" + r4.changed);
check("多段落：第一段保持原样", r4.xml.indexOf(">二甲苯,乙苯<") >= 0);
check("多段落：第二段变为「-」", r4.xml.indexOf(">-<") >= 0, r4.xml.slice(r4.xml.indexOf("二甲苯"), r4.xml.indexOf("二甲苯") + 120));

console.log("\n结果： " + pass + " 通过 / " + fail + " 失败");
if (fail) process.exit(1);
