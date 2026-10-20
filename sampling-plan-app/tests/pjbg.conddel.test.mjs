// 回归测试：检评报告后处理⑥「按条件删除评价依据/警示标识条目 + 章节内序号重排」。
//
// ★ 2026-10-21 口径全面改写（用户逐条指定，14 条）：
//   ① 无「粉尘」（「粉尘、毒物」不算）→ 删 1.1《中华人民共和国尘肺病防治条例》
//   ② 无「毒物」（「粉尘、毒物」不算）→ 删 1.1《使用有毒物品作业场所劳动保护条例》
//   ③ 无高毒物品（判读同旧版）        → 删 1.1《高毒物品目录》
//   ④ 既无「毒物」也无「粉尘」        → 删 1.2.3《呼吸防护装备的选择、使用和维护》(GB 18664-2025)
//   ⑤ 表8-1 检测项目列无「噪声」      → 删 1.2.3《听力防护装备的选择、使用和维护》(GB 23466-2025)
//   ⑥ 无「毒物」（「粉尘、毒物」不算） → 删 1.2.3《呼吸防护 自吸过滤式防毒面具》(GB 2890-2022)
//   ⑦ 无「粉尘」或「粉尘、毒物」      → 删 1.2.3《呼吸防护 自吸过滤式防颗粒物呼吸器》(GB 2626-2019)
//   ⑧ 同上                            → 删 9.2「产生粉尘的工作场所设置…」分句
//   ⑨ 无「粉尘、毒物」                → 删「产生含有有毒物质的混合性粉(烟)尘…」分句
//   ⑩ 无「毒物」（「粉尘、毒物」不算） → 删「有毒物品工作场所设置…」分句
//   ⑪ 无「腐蚀」                      → 删「能引起职业性灼伤或腐蚀的化学品工作场所…」分句
//   ⑫ 表8-1 检测项目列无「噪声」      → 删「产生噪声的工作场所设置…」分句
//   ⑬ 无「高温」                      → 删「高温工作场所设置…」分句
//   ⑭ 无「紫外辐射」                  → 删「产生紫外辐射的工作场所设置…」分句
//
// 夹具按真实报告形态搭：1.1 / 1.2.3 / 9.2 都是「一个段落 + <w:br/> 软换行分行」，
// 9.2 第（4）行把 粉尘/混合/有毒物品/腐蚀/噪声/高温/紫外 七类分句挤在同一行（与 0959 实测一致）。
// 用法：node tests/pjbg.conddel.test.mjs
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

// ---- XML 夹具工具 ----
const esc = (t) => String(t).replace(/&/g, "&amp;").replace(/</g, "&lt;");
function p(text, runs) {
  const parts = runs || [text];
  return '<w:p>' + parts.map((t) => '<w:r><w:t xml:space="preserve">' + esc(t) + '</w:t></w:r>').join("") + '</w:p>';
}
// 一个段落内多行（行间 <w:br/>，真实报告形态）
function pLines(lines) {
  let s = "";
  lines.forEach((ln, i) => {
    if (i > 0) s += "<w:r><w:br/></w:r>";
    s += '<w:r><w:t xml:space="preserve">' + esc(ln) + "</w:t></w:r>";
  });
  return "<w:p>" + s + "</w:p>";
}
function tbl(headers, rows) {
  const hc = headers.map((h) => '<w:tc>' + p(h) + '</w:tc>').join("");
  const rc = rows.map((r) => '<w:tr>' + r.map((c) => '<w:tc>' + p(c) + '</w:tc>').join("") + '</w:tr>').join("");
  return '<w:tbl><w:tr>' + hc + '</w:tr>' + rc + '</w:tbl>';
}
function allText(xml) {
  let t = "";
  const re = /<w:t(?:\s[^>]*)?>([\s\S]*?)<\/w:t>/g;
  let m; while ((m = re.exec(xml))) t += m[1];
  return t;
}
const has = (xml, s) => allText(xml).indexOf(s) >= 0;

// ---- 报告正文夹具 ----
const LAW_ITEMS = [   // 1.1 法律、法规、规章
  '（1）《中华人民共和国职业病防治法》（中华人民共和国主席令第24号，2018年12月29日修订）；',
  '（2）《使用有毒物品作业场所劳动保护条例》（国务院令〔2002〕352号发布，国发〔2023〕20号调整、国务院令〔2024〕797号修正）；',
  '（3）《中华人民共和国尘肺病防治条例》（国发〔1987〕105号）；',
  '（4）《高毒物品目录》（卫法监发（2003）142号）；',
];
const EVAL_ITEMS = [  // 1.2.3 评价依据
  '（1）《呼吸防护装备的选择、使用和维护》(GB 18664-2025)；',
  '（2）《听力防护装备的选择、使用和维护》(GB 23466-2025)；',
  '（3）《呼吸防护 自吸过滤式防颗粒物呼吸器》(GB 2626-2019)；',
  '（4）《呼吸防护 自吸过滤式防毒面具》(GB 2890-2022)；',
];
const KW = {
  dustLaw: "尘肺病防治条例",
  toxLaw: "使用有毒物品作业场所劳动保护条例",
  highTox: "高毒物品目录",
  resp18664: "GB 18664-2025",
  hearing23466: "GB 23466-2025",
  resp2626: "GB 2626-2019",
  mask2890: "GB 2890-2022",
  dustClause: "产生粉尘的工作场所设置",
  dustClause2: "对皮肤有刺激性或经皮肤吸收的粉尘工作场所",
  mixedClause: "产生含有有毒物质的混合性粉(烟)尘的工作场所应设置",
  toxClause: "有毒物品工作场所设置",
  corrodeClause: "能引起职业性灼伤或腐蚀的化学品工作场所",
  noiseClause: "产生噪声的工作场所设置",
  heatClause: "高温工作场所设置",
  uvClause: "产生紫外辐射的工作场所设置",
};
// 9.2 第（4）行：七类警示句挤在同一行（与 0959 实测一致）
const MGMT_LINE4 = '（4）用人单位应按《工作场所职业病危害警示标识》等相关文件的要求，在生产车间各岗位的醒目位置分别设置相应的警示标识和指令标识。'
  + '产生粉尘的工作场所设置“注意防尘”“戴防尘口罩”“注意通风”等警示标识，对皮肤有刺激性或经皮肤吸收的粉尘工作场所还应设置“穿防护服”“戴防护手套”“戴防护眼镜”等警示标识，'
  + '产生含有有毒物质的混合性粉(烟)尘的工作场所应设置“戴防尘毒口罩”等警示标识; '
  + '有毒物品工作场所设置“当心中毒”“穿防护服”“戴防毒面具”“戴防护手套”“戴防护眼镜”“注意通风”等警示标识；'
  + '能引起职业性灼伤或腐蚀的化学品工作场所，设置“当心腐蚀”“当心灼伤”“穿防护服”“戴防护手套”“戴防护眼镜”“戴防毒面具”等警示标识; '
  + '产生噪声的工作场所设置“噪声有害”“戴护耳器”等警示标识; '
  + '高温工作场所设置“当心中暑”“注意高温”“注意通风”等警示标识; '
  + '产生紫外辐射的工作场所设置“当心弧光”“戴防护眼镜”等警示标识。';
const MGMT_LINES = [
  '（1）用人单位应做好个人防护用品选择、使用、维护、更换工作。',
  MGMT_LINE4,
  '（5）用人单位应在生产、使用有毒物品的工作场所设置黄色区域警示线。',
];

function buildConfirm(factors) {
  return tbl(["序号", "职业病危害因素", "存在高毒物品"],
    factors.map((f, i) => [String(i + 1), f.name, f.highTox]));
}
// 表4-1「岗位（工种）设置、劳动者作业情况及职业病危害因素识别情况」
//   第 2 行是子表头（总数 / 数/班），与真实报告一致
function buildT41(hazards) {
  const rows = hazards.map((h) => ["车间A", "岗位甲", "3", "长白班", h, "原辅物料", "7h/d"]);
  return tbl(["单元/工作场所", "岗位（工种）", "人数", "工作班制", "职业病危害因素", "来源", "接触时间"],
    [["", "", "总数", "数/班", "", "", ""]].concat(rows));
}
// 表8-1（仅用于验证「不再按它判断噪声」）
function buildT81(items) {
  return tbl(["单元/工作场所", "岗位/工种", "作业人数", "检测点/对象", "检测项目", "检测结果", "评价结论"],
    items.map((it) => ["车间A", "岗位甲", "3", "1号点", it, "合格", "符合"]));
}
// 危害因素库（取真实库里的典型条目）
const LIB = [
  { name: "大理石粉尘(碳酸钙)", physchem: "粉尘", highTox: "否", dust: "总尘" },
  { name: "锰及其无机化合物(按MnO2计)", physchem: "粉尘、毒物", highTox: "是", dust: "" },
  { name: "硫酸及三氧化硫", physchem: "毒物、腐蚀", highTox: "否", dust: "" },
  { name: "氨", physchem: "毒物", highTox: "是", dust: "" },
  { name: "苯", physchem: "毒物", highTox: "是", dust: "" },
  { name: "高温", physchem: "高温", highTox: "否", dust: "" },
  { name: "紫外辐射", physchem: "紫外辐射", highTox: "否", dust: "" },
  { name: "噪声", physchem: "", highTox: "否", dust: "" },
];
function buildXml(opt) {
  const o = opt || {};
  let s = p("1.1法律、法规、规章") + pLines(LAW_ITEMS);
  s += p("1.2.3评价依据") + pLines(EVAL_ITEMS);
  s += p("9.2职业卫生管理") + pLines(MGMT_LINES);
  s += buildConfirm(o.factors || []);
  if (o.t41) s += buildT41(o.t41);
  if (o.t81) s += buildT81(o.t81);
  return s;
}
const run = (opt) => blf.fixConditionalDeletions(buildXml(opt), {
  noiseWorkshopEmpty: opt && opt.noiseWorkshopEmpty === true,
  hazardFactors: LIB,
});

// ============ [1] 只有毒物（无粉尘 / 无腐蚀 / 无高温 / 无紫外）+ 表8-1 无噪声 ============
console.log("[1] 只有毒物 + 表8-1 无噪声");
{
  const r = run({ factors: [{ name: "氨", highTox: "是" }, { name: "苯", highTox: "是" }], t41: ["氨,苯"] });
  check("① 无粉尘→删尘肺病条例", !has(r.xml, KW.dustLaw));
  check("② 有毒物→保留使用有毒物品条例", has(r.xml, KW.toxLaw));
  check("③ 有高毒→保留高毒目录", has(r.xml, KW.highTox));
  check("④ 有毒物→保留 GB18664", has(r.xml, KW.resp18664));
  check("⑤ 表4-1 无噪声→删 GB23466", !has(r.xml, KW.hearing23466));
  check("⑥ 有毒物→保留 GB2890", has(r.xml, KW.mask2890));
  check("⑦ 无粉尘→删 GB2626", !has(r.xml, KW.resp2626));
  check("⑧ 无粉尘→删粉尘警示句（含第二分句）", !has(r.xml, KW.dustClause) && !has(r.xml, KW.dustClause2));
  check("⑨ 无「粉尘、毒物」→删混合性粉(烟)尘句", !has(r.xml, KW.mixedClause));
  check("⑩ 有毒物→保留有毒物品警示句", has(r.xml, KW.toxClause));
  check("⑪ 无腐蚀→删腐蚀类警示句", !has(r.xml, KW.corrodeClause));
  check("⑫ 表4-1 无噪声→删噪声警示句", !has(r.xml, KW.noiseClause));
  check("⑬ 无高温→删高温警示句", !has(r.xml, KW.heatClause));
  check("⑭ 无紫外辐射→删紫外辐射警示句", !has(r.xml, KW.uvClause));
  check("有删除且未报错", r.changed > 0, String(r.changed));
  check("其它正文保留（黄色区域警示线仍在）", has(r.xml, "黄色区域警示线"));
}

// ============ [2] 只有粉尘 + 表8-1 有噪声 ============
console.log("[2] 只有粉尘 + 表8-1 有噪声");
{
  const r = run({ factors: [{ name: "大理石粉尘(碳酸钙)", highTox: "否" }], t41: ["噪声,大理石粉尘(碳酸钙)"] });
  check("① 有粉尘→保留尘肺病条例", has(r.xml, KW.dustLaw));
  check("② 无毒物→删使用有毒物品条例", !has(r.xml, KW.toxLaw));
  check("③ 无高毒→删高毒目录", !has(r.xml, KW.highTox));
  check("④ 有粉尘→保留 GB18664", has(r.xml, KW.resp18664));
  check("⑤ 表4-1 有噪声→保留 GB23466", has(r.xml, KW.hearing23466));
  check("⑥ 无毒物→删 GB2890", !has(r.xml, KW.mask2890));
  check("⑦ 有粉尘→保留 GB2626", has(r.xml, KW.resp2626));
  check("⑧ 有粉尘→保留粉尘警示句", has(r.xml, KW.dustClause) && has(r.xml, KW.dustClause2));
  check("⑨ 无「粉尘、毒物」→删混合句", !has(r.xml, KW.mixedClause));
  check("⑩ 无毒物→删有毒物品句", !has(r.xml, KW.toxClause));
  check("⑪ 无腐蚀→删腐蚀句", !has(r.xml, KW.corrodeClause));
  check("⑫ 表4-1 有噪声→保留噪声句", has(r.xml, KW.noiseClause));
  check("⑬ 无高温→删高温句", !has(r.xml, KW.heatClause));
  check("⑭ 无紫外→删紫外句", !has(r.xml, KW.uvClause));
}

// ============ [3] 只有「粉尘、毒物」（★ 关键区分：不算「粉尘」也不算「毒物」） ============
console.log("[3] 只有「粉尘、毒物」");
{
  const r = run({ factors: [{ name: "锰及其无机化合物(按MnO2计)", highTox: "是" }], t41: ["锰及其无机化合物(按MnO2计)"] });
  check("① 「粉尘、毒物」不算粉尘→删尘肺病条例", !has(r.xml, KW.dustLaw));
  check("② 「粉尘、毒物」不算毒物→删使用有毒物品条例", !has(r.xml, KW.toxLaw));
  check("③ 有高毒→保留高毒目录", has(r.xml, KW.highTox));
  check("④ 有粉尘/毒物→保留 GB18664", has(r.xml, KW.resp18664));
  check("⑥ 「粉尘、毒物」不算毒物→删 GB2890", !has(r.xml, KW.mask2890));
  check("⑦ 含「粉尘」→保留 GB2626", has(r.xml, KW.resp2626));
  check("⑧ 含「粉尘」→保留粉尘警示句", has(r.xml, KW.dustClause));
  check("⑨ 有「粉尘、毒物」→保留混合性粉(烟)尘句", has(r.xml, KW.mixedClause));
  check("⑩ 不算毒物→删有毒物品句", !has(r.xml, KW.toxClause));
  check("⑤ 表4-1 无噪声→删 GB23466", !has(r.xml, KW.hearing23466));
}

// ============ [4] 毒物 + 腐蚀 ============
console.log("[4] 毒物 + 腐蚀");
{
  const r = run({ factors: [{ name: "硫酸及三氧化硫", highTox: "否" }], t41: [] });
  check("⑪ 有腐蚀→保留腐蚀类警示句", has(r.xml, KW.corrodeClause));
  check("② 有毒物→保留使用有毒物品条例", has(r.xml, KW.toxLaw));
  check("⑥ 有毒物→保留 GB2890", has(r.xml, KW.mask2890));
  check("⑨ 无「粉尘、毒物」→删混合句", !has(r.xml, KW.mixedClause));
  check("⑬⑭ 无高温/紫外→删两句", !has(r.xml, KW.heatClause) && !has(r.xml, KW.uvClause));
}

// ============ [5] 全齐 → 零改动 ============
console.log("[5] 六类因素全齐 + 表8-1 有噪声 → 原样不动");
{
  const r = run({
    factors: [
      { name: "大理石粉尘(碳酸钙)", highTox: "否" },
      { name: "氨", highTox: "是" },
      { name: "锰及其无机化合物(按MnO2计)", highTox: "是" },
      { name: "硫酸及三氧化硫", highTox: "否" },
      { name: "高温", highTox: "否" },
      { name: "紫外辐射", highTox: "否" },
      { name: "噪声", highTox: "否" },
    ],
    t41: ["噪声,大理石粉尘(碳酸钙),氨,锰及其无机化合物(按MnO2计),硫酸及三氧化硫,高温,紫外辐射"],
  });
  check("零改动", r.changed === 0, String(r.changed));
  ["dustLaw", "toxLaw", "highTox", "resp18664", "hearing23466", "mask2890", "resp2626",
    "dustClause", "mixedClause", "toxClause", "corrodeClause", "noiseClause", "heatClause", "uvClause"]
    .forEach((k) => check("保留 " + k, has(r.xml, KW[k])));
}

// ============ [6] 缺数据（因素不在库）→ ①-④⑥-⑪⑬⑭ 全不删；⑤⑫ 仍按表4-1 ============
console.log("[6] 缺数据 → 危害因素相关规则全部不删，噪声仍按表4-1");
{
  const r = run({ factors: [{ name: "某不存在的因素XYZ", highTox: "否" }], t41: ["某不存在的因素XYZ"] });
  ["dustLaw", "toxLaw", "highTox", "resp18664", "mask2890", "resp2626",
    "dustClause", "mixedClause", "toxClause", "corrodeClause", "heatClause", "uvClause"]
    .forEach((k) => check("缺数据→不删 " + k, has(r.xml, KW[k])));
  check("缺数据→仍按表4-1 删 GB23466", !has(r.xml, KW.hearing23466));
  check("缺数据→仍按表4-1 删噪声句", !has(r.xml, KW.noiseClause));
  check("原因里写明缺数据", (r.reasons || []).some((x) => x.indexOf("找不到") >= 0), JSON.stringify(r.reasons));
}

// ============ [7] 没有表4-1 → 退回旧口径 noiseWorkshopEmpty ============
console.log("[7] 无表4-1 → 退回「噪声车间岗位是否为空」");
{
  const r1 = run({ factors: [{ name: "氨", highTox: "是" }], t41: null, noiseWorkshopEmpty: true });
  check("旧口径：噪声为空→删 GB23466", !has(r1.xml, KW.hearing23466));
  check("旧口径：噪声为空→删噪声句", !has(r1.xml, KW.noiseClause));
  const r2 = run({ factors: [{ name: "氨", highTox: "是" }], t41: null, noiseWorkshopEmpty: false });
  check("旧口径：噪声非空→保留 GB23466", has(r2.xml, KW.hearing23466));
  check("旧口径：噪声非空→保留噪声句", has(r2.xml, KW.noiseClause));
}

// ============ [8] 章节内序号重排 ============
console.log("[8] 1.1 删两条后序号重排");
{
  const r = run({ factors: [{ name: "氨", highTox: "是" }], t41: [] });
  const txt = allText(r.xml);
  check("重排后（1）仍是职业病防治法", txt.indexOf("（1）《中华人民共和国职业病防治法》") >= 0, txt.slice(0, 120));
  check("不再出现（4）《高毒物品目录》编号", txt.indexOf("（4）《高毒物品目录》") < 0);
}

// ============ [9] 段落级兜底：某句单独成段时整段删 ============
console.log("[9] 单句单独成段 → 整段删");
{
  const xml = p("1.2.3评价依据")
    + p("（1）《呼吸防护 自吸过滤式防颗粒物呼吸器》(GB 2626-2019)；")
    + p("（2）《呼吸防护 自吸过滤式防毒面具》(GB 2890-2022)；")
    + buildConfirm([{ name: "氨", highTox: "是" }])
    + buildT81(["噪声"]);
  const r = blf.fixConditionalDeletions(xml, { noiseWorkshopEmpty: false, hazardFactors: LIB });
  check("无粉尘→整段删 GB2626", !has(r.xml, KW.resp2626));
  check("有毒物→保留 GB2890", has(r.xml, KW.mask2890));
  check("序号重排为（1）", allText(r.xml).indexOf("（1）《呼吸防护 自吸过滤式防毒面具》") >= 0, allText(r.xml));
}

// ============ [10] 表4-1 与表8-1 结论相反 → 以表4-1 为准 ============
console.log("[10] 噪声判定只看表4-1（表8-1 检测项目列不参与）");
{
  // 表4-1 无噪声，但表8-1 检测项目有噪声 → 仍应删
  const a = run({ factors: [{ name: "氨", highTox: "是" }], t41: ["氨"], t81: ["噪声"] });
  check("表4-1 无噪声→删 GB23466（表8-1 有噪声不生效）", !has(a.xml, KW.hearing23466));
  check("表4-1 无噪声→删噪声警示句", !has(a.xml, KW.noiseClause));
  // 表4-1 有噪声，表8-1 没有 → 应保留
  const b = run({ factors: [{ name: "氨", highTox: "是" }], t41: ["噪声,氨"], t81: ["总粉尘"] });
  check("表4-1 有噪声→保留 GB23466（表8-1 无噪声不生效）", has(b.xml, KW.hearing23466));
  check("表4-1 有噪声→保留噪声警示句", has(b.xml, KW.noiseClause));
}

// ============ [11] 处理后（②/②c 已把噪声删光）→ 表4-1 危害因素列不再有噪声 → 删 ============
console.log("[11] 处理后表4-1 已无噪声 → 删 GB23466 + 噪声句");
{
  // 模拟 ②/②c 处理后的形态：危害因素列里噪声已被删掉，只剩化学因素
  const r = run({ factors: [{ name: "氨", highTox: "是" }], t41: ["氨,苯", "苯"] });
  check("处理后无噪声→删 GB23466", !has(r.xml, KW.hearing23466));
  check("处理后无噪声→删噪声句", !has(r.xml, KW.noiseClause));
  check("原因写明「表4-1…处理后」", (r.reasons || []).some((x) => x.indexOf("表4-1") >= 0), JSON.stringify(r.reasons));
}

console.log(`\npjbg.conddel: ${pass} 通过, ${fail} 失败`);
process.exit(fail > 0 ? 1 : 0);
