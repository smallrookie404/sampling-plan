// 回归测试：检评报告后处理⑥「按条件删除评价依据/警示标识条目 + 章节内序号重排」。
// 用最小 XML 夹具（不依赖真实报告）覆盖：
//   A. 噪声车间岗位为空 → 删 1.2.3 评价依据下的 GB 23466-2025 条目、9.2 下的噪声警示句；
//   B. 物化性质全部不含「毒物」→ 删毒物组；全部不含「粉尘」→ 删粉尘组；
//      粉尘毒物俱无 → 再删 GB 18664-2025；
//   C. 存在高毒物品全部不是「是」→ 删《高毒物品目录》；
//   D. 缺数据（因素在库中找不到）→ 对应组**不删**（保守）；
//   E. 章节内「（1）（2）…」序号重排；
//   F. 无任何条件命中 → 原样不动。
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

// ---- 构造 docx 段落 XML 夹具 ----
// 把一段文本拆成多个 run（模拟报告里文字被拆散的情况）
function p(text, runs) {
  const parts = runs || [text];
  const rs = parts.map((t) => '<w:r><w:t xml:space="preserve">' + t.replace(/&/g, "&amp;").replace(/</g, "&lt;") + '</w:t></w:r>').join("");
  return '<w:p>' + rs + '</w:p>';
}
function tbl(headers, rows) {
  const hc = headers.map((h) => '<w:tc>' + p(h) + '</w:tc>').join("");
  const rc = rows.map((r) => '<w:tr>' + r.map((c) => '<w:tc>' + p(c) + '</w:tc>').join("") + '</w:tr>').join("");
  return '<w:tbl><w:tr>' + hc + '</w:tr>' + rc + '</w:tbl>';
}

// 评价依据章节（1.2.3）——条目带「（N）」序号
function buildEvalSection(items) {
  let s = p("1.2.3评价依据");
  items.forEach((it, i) => { s += p("（" + (i + 1) + "）" + it, ["（" + (i + 1) + "）", it]); });
  return s;
}

const EVAL_ITEMS_FULL = [
  "《中华人民共和国职业病防治法》；",
  "《使用有毒物品作业场所劳动保护条例》（国务院令〔2002〕352号发布，国发〔2023〕20号调整、国务院令〔2024〕797号修正）；",
  "《高毒物品目录》（卫法监发（2003）142号）；",
  "《听力防护装备的选择、使用和维护》(GB 23466-2025)；",
  "《呼吸防护装备的选择、使用和维护》(GB 18664-2025)；",
  "《中华人民共和国尘肺病防治条例》（国发〔1987〕105号）；",
  "《呼吸防护 自吸过滤式防颗粒物呼吸器》(GB 2626-2019)；",
  "《呼吸防护 自吸过滤式防毒面具》(GB 2890-2022)；",
];

function buildMgmtSection() {
  let s = p("9.2职业卫生管理");
  s += p("（1）产生噪声的工作场所设置“噪声有害”“戴护耳器”等警示标识; ");
  s += p("（2）产生粉尘的工作场所设置“注意防尘”“戴防尘口罩”“注意通风”等警示标识，对皮肤有刺激性或经皮肤吸收的粉尘工作场所还应设置“穿防护服”“戴防护手套”“戴防护眼镜”等警示标识，产生含有有毒物质的混合性粉(烟)尘的工作场所应设置“戴防尘毒口罩”等警示标识; ");
  s += p("（3）产生含有有毒物质的混合性粉(烟)尘的工作场所应设置“戴防尘毒口罩”等警示标识; 有毒物品工作场所设置“当心中毒”“穿防护服”“戴防毒面具”“戴防护手套”“戴防护眼镜”“注意通风”等警示标识；能引起职业性灼伤或腐蚀的化学品工作场所，设置“当心腐蚀”“当心灼伤”“穿防护服”“戴防护手套”“戴防护眼镜”“戴防毒面具”等警示标识; ");
  s += p("（4）为劳动者提供符合要求的个人防护用品。");
  return s;
}

function buildConfirmTable(factors) {
  return tbl(["序号", "职业病危害因素", "存在高毒物品"], factors.map((f, i) => [String(i + 1), f.name, f.highTox]));
}

const HAZLIB = [
  { name: "电焊烟尘", physchem: "粉尘", highTox: "否" },
  { name: "氨", physchem: "毒物", highTox: "是" },
  { name: "噪声", physchem: "", highTox: "否" },
  { name: "二氧化氮", physchem: "粉尘、毒物", highTox: "否" },
  { name: "苯", physchem: "毒物", highTox: "是" },
];

function buildXml(factors) {
  return buildEvalSection(EVAL_ITEMS_FULL) + buildMgmtSection() + buildConfirmTable(factors);
}
const countOf = (xml, s) => (xml.split(s).length - 1);
// 把 XML 里的所有 <w:t> 抽出来拼成「可见文本」（跨 run 合并），供断言使用
function allText(xml) {
  let t = "";
  const re = /<w:t(?:\s[^>]*)?>([\s\S]*?)<\/w:t>/g;
  let m; while ((m = re.exec(xml))) t += m[1];
  return t;
}

// ============ 用例 1：噪声为空 + 无粉尘无高毒（只有毒物） ============
console.log("[1] 噪声为空 + 物化性质只有毒物（无粉尘）");
{
  const xml = buildXml([{ name: "氨", highTox: "是" }, { name: "苯", highTox: "是" }]);
  const r = blf.fixConditionalDeletions(xml, { noiseWorkshopEmpty: true, hazardFactors: HAZLIB });
  check("有删除", r.changed > 0, String(r.changed));
  check("删了噪声评价依据(GB23466)", countOf(r.xml, "GB 23466-2025") === 0);
  check("删了噪声警示句", !r.xml.includes("产生噪声的工作场所设置"));
  check("删了尘肺病条例(无粉尘)", countOf(r.xml, "尘肺病防治条例") === 0);
  check("删了防颗粒物呼吸器(GB2626,无粉尘)", countOf(r.xml, "GB 2626-2019") === 0);
  check("删了产生粉尘警示句(无粉尘)", !r.xml.includes("产生粉尘的工作场所设置"));
  check("保留了毒物组(有毒物)", countOf(r.xml, "使用有毒物品作业场所劳动保护条例") === 1);
  check("保留了防毒面具(GB2890,有毒物)", countOf(r.xml, "GB 2890-2022") === 1);
  check("未删GB18664(有毒物)", countOf(r.xml, "GB 18664-2025") === 1);
  // 序号重排：评价依据原 8 条，删 3→ 剩 5 条，编号应为 1..5
  const txt = allText(r.xml);
  check("序号重排为 1..5", txt.includes("（1）《中华人民共和国职业病防治法》") && txt.includes("（5）") && !txt.includes("（6）"));
}

// ============ 用例 2：无粉尘无毒物（噪声非空）→ 再删 GB18664 ============
console.log("[2] 无粉尘无毒物 + 噪声非空");
{
  const xml = buildXml([{ name: "噪声", highTox: "否" }]);
  const r = blf.fixConditionalDeletions(xml, { noiseWorkshopEmpty: false, hazardFactors: HAZLIB });
  check("有删除", r.changed > 0, String(r.changed));
  check("噪声条目保留(噪声非空)", countOf(r.xml, "GB 23466-2025") === 1);
  check("噪声警示句保留", r.xml.includes("产生噪声的工作场所设置"));
  check("删了GB18664(粉尘毒物俱无)", countOf(r.xml, "GB 18664-2025") === 0);
  check("删了毒物组", countOf(r.xml, "使用有毒物品作业场所劳动保护条例") === 0);
  check("删了粉尘组", countOf(r.xml, "尘肺病防治条例") === 0);
}

// ============ 用例 3：有毒物有粉尘 + 有高毒 → 基本不删法规 ============
console.log("[3] 有毒物有粉尘 + 有高毒");
{
  const xml = buildXml([{ name: "二氧化氮", highTox: "否" }, { name: "氨", highTox: "是" }]);
  const r = blf.fixConditionalDeletions(xml, { noiseWorkshopEmpty: false, hazardFactors: HAZLIB });
  check("保留毒物组", countOf(r.xml, "使用有毒物品作业场所劳动保护条例") === 1);
  check("保留粉尘组", countOf(r.xml, "尘肺病防治条例") === 1);
  check("保留GB18664", countOf(r.xml, "GB 18664-2025") === 1);
  check("保留高毒目录(有高毒)", countOf(r.xml, "高毒物品目录") === 1);
  check("保留防毒面具", countOf(r.xml, "GB 2890-2022") === 1);
}

// ============ 用例 4：缺数据（因素不在库）→ 对应组不删 ============
console.log("[4] 缺数据 → 毒物/粉尘/高毒组均不删");
{
  const xml = buildXml([{ name: "某不存在的因素XYZ", highTox: "否" }]);
  const r = blf.fixConditionalDeletions(xml, { noiseWorkshopEmpty: false, hazardFactors: HAZLIB });
  check("缺数据→不删毒物组", countOf(r.xml, "使用有毒物品作业场所劳动保护条例") === 1);
  check("缺数据→不删粉尘组", countOf(r.xml, "尘肺病防治条例") === 1);
  check("缺数据→不删GB18664", countOf(r.xml, "GB 18664-2025") === 1);
  check("缺数据→不删高毒目录", countOf(r.xml, "高毒物品目录") === 1);
  check("无改动", r.changed === 0, String(r.changed));
}

// ============ 用例 5：无任何条件命中 → 原样不动 ============
console.log("[5] 噪声非空 + 有粉尘有毒物 + 有高毒 → 原样");
{
  const xml = buildXml([{ name: "二氧化氮", highTox: "否" }, { name: "苯", highTox: "是" }]);
  const r = blf.fixConditionalDeletions(xml, { noiseWorkshopEmpty: false, hazardFactors: HAZLIB });
  check("changed===0", r.changed === 0, String(r.changed));
  check("XML 完全未变", r.xml === xml);
}

// ============ 用例 6：collectConfirmColTexts 单独校验 ============
console.log("[6] collectConfirmColTexts / lookupHazardField");
{
  const xml = buildConfirmTable([{ name: "电焊烟尘", highTox: "否" }, { name: "氨", highTox: "是" }]);
  const cf = blf.collectConfirmColTexts(xml, ["职业病危害因素"]);
  check("找到确认表", cf.found === true);
  check("取到 2 个因素", cf.texts.length === 2, JSON.stringify(cf.texts));
  const lk = blf.lookupHazardField(HAZLIB, cf.texts, "physchem");
  check("物化性质可判定", lk.ok === true, JSON.stringify(lk));
  check("值为 粉尘 / 毒物", lk.values.join("|") === "粉尘|毒物", JSON.stringify(lk.values));
}

// ============ 用例 7：标题识别不误伤法规条目（「1.《xx》」不算标题） ============
console.log("[7] 数字+书名号条目不作为章节标题");
{
  let xml = p("1.2.3评价依据");
  xml += p("1.《中华人民共和国职业病防治法》；");       // 数字开头条目，不应被当标题
  xml += p("2.《中华人民共和国尘肺病防治条例》（国发〔1987〕105号）；");
  xml += p("3.《高毒物品目录》（卫法监发（2003）142号）；");
  const r = blf.deletePjbgParas(xml, ["《中华人民共和国尘肺病防治条例》"]);
  check("删掉含尘肺病条例的条目", countOf(r.xml, "尘肺病防治条例") === 0, String(r.changed));
  check("保留职业病防治法条目", allText(r.xml).includes("中华人民共和国职业病防治法"));
}

console.log(`\n结果：${pass} 通过 / ${fail} 失败`);
process.exit(fail ? 1 : 0);
