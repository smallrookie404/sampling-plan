// 回归测试：检评报告后处理⑦「两班制作息说明句标红」。
// 覆盖：真实段落形态（单 run 单 <w:t>）切分上色、同段前缀不上色、w:color 插入位置、
//       无锚点不动、已有颜色不重复、句尾无句号、跨 run 时不破坏结构。
// 用法：node tests/pjbg.shiftred.test.mjs
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

const RPR = '<w:rPr><w:rFonts w:eastAsia="仿宋_GB2312"/><w:b w:val="0"/><w:spacing w:val="0"/><w:kern w:val="2"/><w:sz w:val="21"/><w:lang w:val="en-US"/></w:rPr>';
const run = (t) => '<w:r>' + RPR + '<w:t xml:space="preserve">' + t + '</w:t></w:r>';
const para = (inner) => '<w:p><w:pPr>' + RPR + '</w:pPr>' + inner + '</w:p>';

const SENT = "打磨岗位实行两班制，上班时间为白班8:00～20:00、晚班20:00～8:00（次日），其余各岗位均实行白班制，上班时间为8:00～11:30、12:30～17:00。";
const PREFIX = "注：用人单位职工总数38人，接触职业病危害人数25人。";

// 取段落里每个 run 的 { text, hasColor, xml }
function runsOf(pXml) {
  return (pXml.match(/<w:r(?:\s[^>]*)?>[\s\S]*?<\/w:r>/g) || []).map((x) => ({
    text: (x.match(/<w:t(?:\s[^>]*)?>([\s\S]*?)<\/w:t>/) || [])[1] || "",
    hasColor: /<w:color w:val="FF0000"\/>/.test(x),
    xml: x,
  }));
}

// ============ 1. 真实形态：单 run 单 <w:t>，前缀 + 目标句 ============
console.log("[1] 真实段落（单 run）→ 切成「前缀 + 红色句」");
{
  const xml = para(run(PREFIX + SENT));
  const r = blf.markShiftNoteRed(xml);
  check("changed===1", r.changed === 1, String(r.changed));
  const rs = runsOf(r.xml);
  check("切成 2 个 run", rs.length === 2, String(rs.length));
  check("前段文本 = 注：…25人。", rs[0] && rs[0].text === PREFIX, rs[0] && rs[0].text);
  check("前段**不上色**", rs[0] && !rs[0].hasColor);
  check("★ 红色段文本 = 整句（含句号）", rs[1] && rs[1].text === SENT, rs[1] && rs[1].text);
  check("★ 红色段带 FF0000", rs[1] && rs[1].hasColor);
  check("段落文字整体不变", rs.map((x) => x.text).join("") === PREFIX + SENT);
}

// ============ 2. w:color 必须插在 w:spacing / w:sz / w:lang 之前 ============
console.log("[2] w:color 的插入位置符合 OOXML 顺序");
{
  const r = blf.markShiftNoteRed(para(run(PREFIX + SENT)));
  const red = (r.xml.match(/<w:r>(?:(?!<\/w:r>)[\s\S])*?<w:color w:val="FF0000"\/>[\s\S]*?<\/w:r>/) || [""])[0];
  const iColor = red.indexOf('<w:color');
  const iSpacing = red.indexOf('<w:spacing');
  const iSz = red.indexOf('<w:sz ');
  const iLang = red.indexOf('<w:lang');
  check("color 在 spacing 之前", iColor >= 0 && iSpacing > iColor, "color@" + iColor + " spacing@" + iSpacing);
  check("color 在 sz 之前", iSz > iColor, "sz@" + iSz);
  check("color 在 lang 之前", iLang > iColor, "lang@" + iLang);
}

// ============ 3. 无锚点 → 原样不动 ============
console.log("[3] 没有「实行两班制」→ 不动");
{
  const xml = para(run("注：用人单位职工总数10人。"));
  const r = blf.markShiftNoteRed(xml);
  check("changed===0", r.changed === 0, String(r.changed));
  check("XML 完全未变", r.xml === xml);
}

// ============ 4. 已有颜色 → 不重复添加 ============
console.log("[4] run 已有 w:color → 替换而不是叠加");
{
  const xml = para('<w:r><w:rPr><w:color w:val="0000FF"/><w:sz w:val="21"/></w:rPr><w:t xml:space="preserve">' + SENT + '</w:t></w:r>');
  const r = blf.markShiftNoteRed(xml);
  check("changed===1", r.changed === 1, String(r.changed));
  check("只剩一个 w:color", (r.xml.match(/<w:color /g) || []).length === 1, String((r.xml.match(/<w:color /g) || []).length));
  check("颜色已变为 FF0000", /<w:color w:val="FF0000"\/>/.test(r.xml));
  check("旧颜色 0000FF 已消失", r.xml.indexOf("0000FF") < 0);
}

// ============ 5. 目标句在段末且没有句号 → 一直到段尾 ============
console.log("[5] 句尾无「。」→ 上色到段尾");
{
  const tail = "打磨岗位实行两班制，上班时间为8:00～20:00";
  const r = blf.markShiftNoteRed(para(run(PREFIX + tail)));
  const rs = runsOf(r.xml);
  check("changed===1", r.changed === 1, String(r.changed));
  check("红色段 = 从「打磨岗位」到段尾", rs.some((x) => x.hasColor && x.text === tail), JSON.stringify(rs.map((x) => x.text)));
  check("前缀仍不上色", rs.some((x) => !x.hasColor && x.text === PREFIX));
}

// ============ 6. 句子前面没有句号 → 从段首开始上色 ============
console.log("[6] 段首就是该句（前面无句号）→ 从 0 开始");
{
  const r = blf.markShiftNoteRed(para(run(SENT)));
  const rs = runsOf(r.xml);
  check("整段一个红色 run", rs.length === 1 && rs[0].hasColor && rs[0].text === SENT, JSON.stringify(rs.map((x) => x.text)));
}

// ============ 7. 跨多个 run → 各 run 覆盖部分分别上色，不破坏结构 ============
console.log("[7] 目标跨多个 run → 分段上色且文本不变");
{
  const a = "打磨岗位实行两班制，";
  const b = "上班时间为白班8:00～20:00。";
  const xml = para(run(PREFIX) + run(a) + run(b));
  const r = blf.markShiftNoteRed(xml);
  const rs = runsOf(r.xml);
  check("changed===1", r.changed === 1, String(r.changed));
  check("整体文本不变", rs.map((x) => x.text).join("") === PREFIX + a + b, JSON.stringify(rs.map((x) => x.text)));
  check("含「实行两班制」的 run 已上色", rs.some((x) => x.hasColor && x.text.indexOf("实行两班制") >= 0));
  check("前缀 run 未上色", rs.some((x) => !x.hasColor && x.text === PREFIX));
}

// ============ 8. 复杂 run（含 <w:br/>）→ 放过不破坏 ============
console.log("[8] 含软换行的 run → 原样放过（宁可不标红也不破坏结构）");
{
  const xml = para('<w:r>' + RPR + '<w:t xml:space="preserve">' + PREFIX + SENT + '</w:t><w:br/></w:r>');
  const r = blf.markShiftNoteRed(xml);
  check("changed===0（跳过）", r.changed === 0, String(r.changed));
  check("XML 未变", r.xml === xml);
}

console.log(`\n结果：${pass} 通过 / ${fail} 失败`);
process.exit(fail ? 1 : 0);
