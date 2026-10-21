// 回归测试：⑪「无腐蚀」裁剪——真实报告（BTC26-SZJC0986）触发链。
//
// 背景（2026-10-21 实测 0986 报告）：
//   9.2 第（4）行是一条长句，粉尘/混合性粉(烟)尘/有毒物品/腐蚀/噪声/高温/紫外 七类警示句挤在同一行。
//   ⑪ 无腐蚀时应删掉「能引起职业性灼伤或腐蚀的化学品工作场所，设置“当心腐蚀”…」整句。
//   但第一版修复只处理了「分句里没有句号」的情形，真实报告里漏掉了：
//     · dust 规则先删掉「如产生粉尘的工作场所设置…」「对皮肤有刺激性…」两个分句；
//     · mixed 规则再删「产生含有有毒物质的混合性粉(烟)尘…」时**连分隔符一起吃掉**；
//     · 于是「…指令标识。」与「能引起…工作场所」粘成**同一分句**（中间只剩句号）；
//     · corrode 命中时走的是「保留句号前引导句」分支，主语被删、谓语「设置“当心腐蚀”…」残留。
//   修复：两个分支都要检查「紧随的下一分句是否以『设置』开头」→ 是则一并丢弃。
//
// 夹具直接用 0986 报告的原文（含「。如产生粉尘…」引导句），比人造夹具更贴近真实。
// 用法：node tests/pjbg.corrode.test.mjs
import fs from "node:fs";
import path from "node:path";
import vm from "node:vm";
import { fileURLToPath } from "node:url";

const here = path.dirname(fileURLToPath(import.meta.url));
const APP = path.resolve(here, "..");
const blfSrc = fs.readFileSync(path.join(APP, "js/blf.js"), "utf8");

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
const sandbox = { console, document: documentStub,
  DOMParser: class { parseFromString() { return { getElementsByTagNameNS: () => [] }; } },
  setTimeout, clearTimeout, fetch: () => Promise.reject(new Error("no fetch")),
  URL: { createObjectURL: () => "blob:x", revokeObjectURL: noop } };
sandbox.self = sandbox; sandbox.window = sandbox;
vm.createContext(sandbox);
vm.runInContext(blfSrc, sandbox, { filename: "blf.js" });
const blf = sandbox.SamplingBlf;

// ---- 最小 XML 夹具 ----
const esc = (s) => String(s).replace(/&/g, "&amp;").replace(/</g, "&lt;");
function p(text) { return '<w:p><w:r><w:t xml:space="preserve">' + esc(text) + "</w:t></w:r></w:p>"; }
function pLines(lines) {
  let s = "";
  lines.forEach((ln, i) => {
    if (i > 0) s += "<w:r><w:br/></w:r>";
    s += '<w:r><w:t xml:space="preserve">' + esc(ln) + "</w:t></w:r>";
  });
  return "<w:p>" + s + "</w:p>";
}
function tbl(headers, rows) {
  const hc = headers.map((h) => "<w:tc>" + p(h) + "</w:tc>").join("");
  const rc = rows.map((r) => "<w:tr>" + r.map((c) => "<w:tc>" + p(c) + "</w:tc>").join("") + "</w:tr>").join("");
  return "<w:tbl><w:tr>" + hc + "</w:tr>" + rc + "</w:tbl>";
}

// ★ 0986 报告 9.2 第（4）行原文（逐字照抄，含「。如产生粉尘…」引导句与半角/全角分号混用）
const REAL_LINE4 =
  '（4）用人单位应按《工作场所职业病危害警示标识》和《用人单位职业病危害告知与警示标识管理规范》等相关文件的要求，' +
  '在生产车间各岗位的醒目位置分别设置相应的警示标识和指令标识。' +
  '如产生粉尘的工作场所设置“注意防尘”“戴防尘口罩”“注意通风”等警示标识，' +
  '对皮肤有刺激性或经皮肤吸收的粉尘工作场所还应设置“穿防护服”“戴防护手套”“戴防护眼镜”等警示标识，' +
  '产生含有有毒物质的混合性粉(烟)尘的工作场所应设置“戴防尘毒口罩”等警示标识; ' +
  '有毒物品工作场所设置“当心中毒”“穿防护服”“戴防毒面具”“戴防护手套”“戴防护眼镜”“注意通风”等警示标识；' +
  '能引起职业性灼伤或腐蚀的化学品工作场所，设置“当心腐蚀”“当心灼伤”“穿防护服”“戴防护手套”“戴防护眼镜”“戴防毒面具”等警示标识; ' +
  '产生噪声的工作场所设置“噪声有害”“戴护耳器”等警示标识; ' +
  '高温工作场所设置“当心中暑”“注意高温”“注意通风”等警示标识; ' +
  '产生紫外辐射的工作场所设置“弧光”“戴防护眼镜”等警示标识。';

const K = {
  dust: "产生粉尘的工作场所设置",
  dust2: "对皮肤有刺激性或经皮肤吸收的粉尘工作场所",
  mixed: "产生含有有毒物质的混合性粉(烟)尘的工作场所应设置",
  tox: "有毒物品工作场所设置",
  corrode: "能引起职业性灼伤或腐蚀的化学品工作场所",
  corrodeVerb: "设置“当心腐蚀”",
  noise: "产生噪声的工作场所设置",
  heat: "高温工作场所设置",
  uv: "产生紫外辐射的工作场所设置",
};

// 危害因素库（与真实库同名条目的 physchem 一致）
const LIB = [
  { name: "大理石粉尘(碳酸钙)", physchem: "粉尘", highTox: "否", dust: "总尘" },
  { name: "锰及其无机化合物(按MnO2计)", physchem: "粉尘、毒物", highTox: "是", dust: "" },
  { name: "硫酸及三氧化硫", physchem: "毒物、腐蚀", highTox: "否", dust: "" },
  { name: "氨", physchem: "毒物", highTox: "是", dust: "" },
  { name: "高温", physchem: "高温", highTox: "否", dust: "" },
  { name: "紫外辐射", physchem: "紫外辐射", highTox: "否", dust: "" },
  { name: "噪声", physchem: "噪声", highTox: "否", dust: "" },
];

function buildXml(o) {
  let s = p("9.2职业卫生管理");
  s += pLines([REAL_LINE4, "（5）用人单位应在生产、使用有毒物品的工作场所设置黄色区域警示线。"]);
  // 检测项目确认表
  s += tbl(["序号", "职业病危害因素", "存在高毒物品"],
    o.factors.map((f, i) => [String(i + 1), f, "否"]));
  // 表4-1（判 ⑤⑫ 噪声用）：第 2 行是子表头
  s += tbl(["单元/工作场所", "岗位（工种）", "人数", "工作班制", "职业病危害因素", "来源", "接触时间"],
    [["", "", "总数", "数/班", "", "", ""], ["车间A", "岗位甲", "3", "长白班", o.t41, "原辅物料", "7h/d"]]);
  return "<w:body>" + s + "</w:body>";
}
const run = (o) => blf.fixConditionalDeletions(buildXml(o), { noiseWorkshopEmpty: false, hazardFactors: LIB });
const has = (xml, s) => xml.indexOf(s) >= 0;

// ============ [1] 只有噪声（0986 实况）============
console.log("[1] 只有噪声 → 五类警示句全删、噪声句保留");
{
  const r = run({ factors: ["噪声"], t41: "噪声" });
  check("粉尘句已删", !has(r.xml, K.dust));
  check("粉尘第二分句已删", !has(r.xml, K.dust2));
  check("混合性粉(烟)尘句已删", !has(r.xml, K.mixed));
  check("有毒物品句已删", !has(r.xml, K.tox));
  check("腐蚀句主语已删", !has(r.xml, K.corrode));
  check("★ 腐蚀句谓语也删（不残留「设置“当心腐蚀”…」）", !has(r.xml, K.corrodeVerb),
    "残留了谓语分句");
  check("高温句已删", !has(r.xml, K.heat));
  check("紫外句已删", !has(r.xml, K.uv));
  check("噪声句保留（项目有噪声）", has(r.xml, K.noise));
  check("引导句保留", has(r.xml, "分别设置相应的警示标识和指令标识"));
  // 标点收尾：不应出现「。；」「；；」「，，」这类删分句留下的残迹
  check("无「。;」残留", !/。[；;]/.test(r.xml));
  check("无「；；」残留", !/[；;]\s*[；;]/.test(r.xml));
  check("无「，，」残留", !/[，,]\s*[，,]/.test(r.xml));
}

// ============ [2] 有腐蚀（硫酸及三氧化硫，physchem=毒物、腐蚀）============
console.log("[2] 有腐蚀 → 腐蚀句完整保留");
{
  const r = run({ factors: ["硫酸及三氧化硫", "噪声"], t41: "噪声" });
  check("腐蚀句主语保留", has(r.xml, K.corrode));
  check("腐蚀句谓语保留", has(r.xml, K.corrodeVerb));
}

// ============ [3] 有高温 ============
console.log("[3] 有高温 → 高温句保留");
{
  const r = run({ factors: ["高温", "噪声"], t41: "噪声" });
  check("高温句保留", has(r.xml, K.heat));
  check("腐蚀句仍删（无腐蚀）", !has(r.xml, K.corrode) && !has(r.xml, K.corrodeVerb));
}

console.log("\npjbg.corrode: " + pass + " 通过, " + fail + " 失败");
process.exit(fail ? 1 : 0);
