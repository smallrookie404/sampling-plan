// 回归测试：接触职业病危害人数改为「表4-1『人数 / 总数』按岗位去重合计」（2026-10-19 口径变更）。
//
// 新口径（用户要求）：
//   · 取「表4-1 岗位（工种）设置、劳动者作业情况及职业病危害因素识别情况」里
//     「人数」列下「总数」子列的数字；
//   · 同一岗位的多个点位行会把**同一个人数**重复写在每行（包装工 20/20/20、压机工 30×4…）
//     → 必须按「车间+岗位」去重后**只计一次**（1121 样例：逐格相加 231、去重才是 77）；
//   · 某岗位的所有行里「职业病危害因素」列**全部为空或「-」**→ 该岗位人数不纳入统计；
//   · 该步在 ②/②c 删完噪声之后执行（删空填「-」的岗位自然被排除）；
//   · 表4-1 找不到时退回旧的表8-1「作业人数」去重合计（兜底）。
//
// 覆盖：A 去重 + 整组为「-」排除；B 表3-1 与注段落写入新值；C 无表4-1 回退表8-1；
//       D 纵向合并续行（人数格为空）归入同组；E 「—」「无」「/」都算空；
//       F 真实 1121 样例（桌面存在时才跑）：77 人，与报告表3-1 原值一致。
//
// 用法：node tests/pjbg.contactcount.test.mjs
import fs from "node:fs";
import path from "node:path";
import vm from "node:vm";
import { fileURLToPath } from "node:url";
import JSZip from "jszip";

const here = path.dirname(fileURLToPath(import.meta.url));
const APP = path.resolve(here, "..");
const blfSrc = fs.readFileSync(path.join(APP, "js/blf.js"), "utf8");

let pass = 0, fail = 0;
function check(name, cond, extra) {
  if (cond) { pass++; console.log("  ok  " + name); }
  else { fail++; console.log("  FAIL " + name + (extra ? " :: " + extra : "")); }
}

const noop = () => {};
const fakeEl = () => ({ style: {}, classList: { add: noop, remove: noop, toggle: noop, contains: () => false },
  addEventListener: noop, appendChild: noop, removeChild: noop, setAttribute: noop, getAttribute: () => null,
  hasAttribute: () => false, querySelector: () => null, querySelectorAll: () => [], innerHTML: "", textContent: "",
  disabled: false, click: noop, children: [], parentNode: null, dataset: {} });
const documentStub = { readyState: "complete", getElementById: fakeEl, querySelector: fakeEl,
  querySelectorAll: () => [], createElement: fakeEl, addEventListener: noop,
  body: fakeEl(), head: fakeEl(), documentElement: fakeEl() };

const sandbox = {
  console: { log: noop, info: noop, warn: noop, error: noop },
  document: documentStub,
  DOMParser: class { parseFromString() { return { getElementsByTagNameNS: () => [] }; } },
  setTimeout, clearTimeout, fetch: () => Promise.reject(new Error("no fetch")),
  URL: { createObjectURL: () => "blob:x", revokeObjectURL: noop },
  JSZip, TextDecoder,
};
sandbox.self = sandbox; sandbox.window = sandbox;
vm.createContext(sandbox);
vm.runInContext(blfSrc, sandbox, { filename: "blf.js" });
const blf = sandbox.SamplingBlf;

// ---- XML 夹具 ----
const esc = (s) => String(s).replace(/&/g, "&amp;").replace(/</g, "&lt;");
const p = (text) => "<w:p><w:r><w:t xml:space=\"preserve\">" + esc(text) + "</w:t></w:r></w:p>";
const tc = (text) => "<w:tc>" + p(text) + "</w:tc>";
const tr = (cells) => "<w:tr>" + cells.map(tc).join("") + "</w:tr>";
const tbl = (headCells, rows) => "<w:tbl>" + tr(headCells) + rows.map((r) => tr(r)).join("") + "</w:tbl>";

// 表4-1：表头 13 格；数据行 14 格（「人数」拆成 总数 / 数/班）
const H41 = ["单元/工作场所", "岗位（工种）", "人数", "工作班制", "工作时间", "作业类型", "作业方式",
  "接触类型", "工作地点及内容", "职业病危害因素", "来源", "接触时间", "体力劳动强度"];
// 数据行：[unit, job, 总数, 数/班, 工作班制, 工作时间, 作业类型, 作业方式, 接触类型, 工作地点, 危害因素, 来源, 接触时间, 体力]
function row41(unit, job, total, place, hazard) {
  return [unit, job, String(total), String(total), "长白班", "9h/d", "固定作业", "半手工作业", "稳定",
    place, hazard, "设备运行", "7h/d", "Ⅱ"];
}
const SUB_ROW = ["", "", "总数", "数/班", "", "", "", "", "", "", "", "", "", ""];

// 表3-1（键值对表）
const T31 = [
  ["在册职工人数", "300人", "劳务派遣人员数量", "0人"],
  ["接触职业病危害人数", "0人", "法定代表人/负责人", "-"],
];

// 表8-1（兜底口径用）
const H81 = ["单元/工作场所", "岗位/工种", "作业人数", "检测点/对象", "检测项目", "检测结果", "补救措施", "评价结论", "存在高毒物品"];

function buildXml(opts) {
  const o = opts || {};
  let s = p("表4-1 岗位（工种）设置、劳动者作业情况及职业病危害因素识别情况");
  if (o.t41 !== false) {
    const rows = [SUB_ROW].concat(o.rows || []);
    s += tbl(H41, rows);
  }
  s += p(o.note || "注：用人单位职工总数0人，接触职业病危害人数0人。");
  if (o.t81) s += p("表8-1 检测结果与评价结论") + tbl(H81, o.t81);
  s += p("表3-1 用人单位基本情况") + tbl(["项目", "内容", "项目", "内容"], T31);
  return '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    + '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>'
    + s + "</w:body></w:document>";
}

// ============ A. 按岗位去重 + 整组「-」排除 ============
console.log("[A] 表4-1「人数/总数」按岗位去重合计（整组为「-」不纳入）");
{
  const rows = [
    row41("包装车间", "包装工", 20, "包装1/作业", "-"),
    row41("包装车间", "包装工", 20, "包装2/作业", "-"),      // 整组「-」→ 20 不计
    row41("注塑车间", "注塑工", 20, "注塑1/作业", "高温,噪声"),
    row41("注塑车间", "注塑工", 20, "注塑2/作业", "高温"),    // 同岗位重复人数 → 只计一次 20
    row41("注塑车间", "粉碎工", 1, "粉碎/作业", "聚氯乙烯粉尘,噪声"),
    row41("注塑车间", "修边工", 8, "修边1/作业", "-"),
    row41("注塑车间", "修边工", 8, "修边2/作业", ""),        // 全空 → 8 不计
  ];
  const xml = buildXml({ rows: rows });
  const r = blf.sumExposedFromT41(xml);
  console.log("   ->", JSON.stringify(r));
  check("能定位表4-1", !!r);
  check("合计 = 21（20 注塑工 + 1 粉碎工）", r && r.total === 21, r && String(r.total));
  check("岗位去重后 4 个（包装工/注塑工/粉碎工/修边工）", r && r.groups === 4, r && String(r.groups));
  check("纳入 2 个、排除 2 个", r && r.counted === 2 && r.skipped === 2,
    r && (r.counted + "/" + r.skipped));
  // 反证：逐格相加会得到 69（20*3 + 1 + 8*2），说明去重是必要的
  check("不是逐格相加（69）", r && r.total !== 69, r && String(r.total));
}

// ============ B. 表3-1 + 注段落写入新值 ============
console.log("[B] fixNoteAndContactCount：表3-1 与注段落都写入 21人");
{
  const xml = buildXml({ rows: [row41("注塑车间", "注塑工", 20, "注塑1/作业", "高温,噪声"),
    row41("注塑车间", "粉碎工", 1, "粉碎/作业", "噪声")] });
  const r = blf.fixNoteAndContactCount(xml, "");
  check("来源 = 表4-1", r.totalFrom === "t41", r.totalFrom);
  check("total = 21", r.total === 21, String(r.total));
  check("contactVal = 21人", r.contactVal === "21人", r.contactVal);
  check("注段落 A 段写入 21人", r.noteTexts[0].indexOf("接触职业病危害人数21人") >= 0, r.noteTexts[0].slice(0, 60));
  check("表3-1 右邻 = 21人", /接触职业病危害人数<\/w:t><\/w:r><\/w:p><\/w:tc><w:tc>(?:<w:p><w:r><w:t[^>]*>)?21人/.test(r.xml)
    || r.xml.indexOf(">21人<") >= 0, "未找到 21人 单元格");
}

// ============ C. 无表4-1 → 回退表8-1 ============
console.log("[C] 没有表4-1 → 回退表8-1 作业人数去重合计");
{
  const xml = buildXml({
    t41: false,
    t81: [
      ["车间A", "岗位A", "5", "点1", "噪声", "符合", "-", "-", "否"],
      ["车间A", "岗位A", "5", "点1", "高温", "符合", "-", "-", "否"],   // 同岗位去重
      ["车间A", "岗位B", "3", "点2", "噪声", "符合", "-", "-", "否"],
    ],
  });
  const r = blf.fixNoteAndContactCount(xml, "");
  check("来源 = 表8-1（兜底）", r.totalFrom === "t81", r.totalFrom);
  check("total = 8（5 + 3）", r.total === 8, String(r.total));
}

// ============ D. 纵向合并续行（人数格为空）归入同组 ============
console.log("[D] 人数格为空但车间/岗位相同的续行 → 并入同组（危害因素按「全部行」判定）");
{
  const rows = [
    row41("车间A", "岗位A", 6, "点1/作业", "-"),
    row41("车间A", "岗位A", "", "点2/作业", "高温"),      // 续行：人数格为空，但危害因素有内容 → 整组纳入
    row41("车间A", "岗位B", 2, "点3/作业", "-"),          // 整组只有一个「-」→ 不纳入
  ];
  const r = blf.sumExposedFromT41(buildXml({ rows: rows }));
  console.log("   ->", JSON.stringify(r));
  check("合计 = 6（岗位A 的续行有危害因素 → 纳入；岗位B 排除）", r && r.total === 6, r && String(r.total));
  check("2 个岗位、纳入 1、排除 1", r && r.groups === 2 && r.counted === 1 && r.skipped === 1,
    r && JSON.stringify(r));
}

// ============ E. 「—」「无」「/」也算空 ============
console.log("[E] 危害因素列为「—」「无」「/」都视为空");
{
  const rows = [
    row41("车间A", "岗位A", 9, "点1/作业", "—"),
    row41("车间A", "岗位B", 4, "点2/作业", "无"),
    row41("车间A", "岗位C", 3, "点3/作业", "/"),
    row41("车间A", "岗位D", 2, "点4/作业", "噪声"),
  ];
  const r = blf.sumExposedFromT41(buildXml({ rows: rows }));
  check("只纳入 2（岗位D）", r && r.total === 2, r && String(r.total));
  check("排除 3 个岗位", r && r.skipped === 3, r && String(r.skipped));
}

// ============ F. 真实 1121 样例（桌面存在时才跑） ============
const SAMPLE = "C:/Users/54606/OneDrive/Desktop/BTC26-SZJC1121-浙江三门维艾尔工业有限公司_评价报告_49.docx";
if (fs.existsSync(SAMPLE)) {
  console.log("\n[F] 真实 1121 报告：合计 77 人（与报告表3-1 原值一致）");
  const xml = await (await JSZip.loadAsync(fs.readFileSync(SAMPLE))).file("word/document.xml").async("string");
  const t41 = blf.sumExposedFromT41(xml);
  console.log("   ->", JSON.stringify(t41));
  check("15 个岗位、纳入 11、排除 4（包装工/缝纫工/修边工/切片工 全为「-」）",
    t41 && t41.groups === 15 && t41.counted === 11 && t41.skipped === 4, t41 && JSON.stringify(t41));
  check("合计 = 77 人", t41 && t41.total === 77, t41 && String(t41.total));
  const basic = blf.readBasicInfoTable(xml);
  check("与表3-1「接触职业病危害人数」原值 77人 一致", basic && basic.contactText === "77人", basic && basic.contactText);
} else {
  console.log("\n[F] 跳过：未找到 1121 样例（" + SAMPLE + "）");
}

// ============ G. 真实 0959 样例（桌面存在时才跑） ============
// 用户 2026-10-19 反馈：这份报告的接害人数应为 87（旧表8-1 口径与报告原值都是 84，偏低 3）。
const SAMPLE2 = "C:/Users/54606/OneDrive/Desktop/BTC26-SZJC0959-浙江永贵电器股份有限公司_评价报告_49.docx";
if (fs.existsSync(SAMPLE2)) {
  console.log("\n[G] 真实 0959 报告：新口径 87 人（旧表8-1 = 报告原值 = 84，偏低 3）");
  const xml = await (await JSZip.loadAsync(fs.readFileSync(SAMPLE2))).file("word/document.xml").async("string");
  const t41 = blf.sumExposedFromT41(xml);
  const t81 = blf.sumExposedWorkers(xml);
  console.log("   -> 新:", JSON.stringify(t41), " 旧:", JSON.stringify(t81));
  check("新口径 = 87 人", t41 && t41.total === 87, t41 && String(t41.total));
  check("旧口径（表8-1）= 84 人（对照）", t81 && t81.total === 84, t81 && String(t81.total));
  check("29 个岗位、纳入 22、排除 7", t41 && t41.groups === 29 && t41.counted === 22 && t41.skipped === 7,
    t41 && JSON.stringify(t41));
  // 冲床岗位：剪板行为「-」、冲床行有「噪声」→ 整组仍有内容 → 必须计入 3 人（正是 84→87 的差额）
  check("冲床岗位计入（剪板 - 但冲床行有噪声）", t41 && t41.total - (t81 ? t81.total : 0) === 3,
    String((t41 ? t41.total : 0) - (t81 ? t81.total : 0)));
} else {
  console.log("\n[G] 跳过：未找到 0959 样例（" + SAMPLE2 + "）");
}

console.log(`\n结果：${pass} 通过 / ${fail} 失败`);
process.exit(fail ? 1 : 0);
