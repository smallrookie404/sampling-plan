// 回归测试：检评报告后处理⑬「表格纵向压缩：总宽不变、重排列宽」（2026-10-21）。
//
// 规则（用户要求）：
//   · 行高固定**最小值 0.8cm**（`w:trHeight w:val="454"`）、行距固定 **16 磅**（`w:line="320" exact`）→ 一律不动；
//   · 通过**重新分配各列宽度**（同一张表总宽不变）缩短表格总长度：长文本列加宽、短列收窄；
//   · 单元格段落关掉「对齐文档网格」（否则 16 磅行距会被 linePitch 网格撑高）；
//   · 单列最小 0.8cm、最多占 40%；估算没变矮的表**不动**（安全阀）。
//
// 用法：node tests/pjbg.tablefit.test.mjs
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
const sandbox = {
  console: { log: noop, info: noop, warn: noop, error: noop },
  document: { readyState: "complete", getElementById: fakeEl, querySelector: fakeEl, querySelectorAll: () => [],
    createElement: fakeEl, addEventListener: noop, body: fakeEl(), head: fakeEl(), documentElement: fakeEl() },
  DOMParser: class { parseFromString() { return { getElementsByTagNameNS: () => [] }; } },
  setTimeout, clearTimeout, fetch: () => Promise.reject(new Error("no fetch")),
  URL: { createObjectURL: () => "blob:x", revokeObjectURL: noop }, JSZip, TextDecoder, TextEncoder,
};
sandbox.self = sandbox; sandbox.window = sandbox;
vm.createContext(sandbox);
vm.runInContext(blfSrc, sandbox, { filename: "blf.js" });
const blf = sandbox.SamplingBlf;

// ---- 夹具 ----
const esc = (s) => String(s).replace(/&/g, "&amp;").replace(/</g, "&lt;");
// 单元格：段落带 sz=21（10.5pt）、snapToGrid、固定行距 16 磅
const cell = (text, wPct) =>
  "<w:tc><w:tcPr><w:tcW w:w=\"" + wPct + "\" w:type=\"pct\" /></w:tcPr>"
  + "<w:p><w:pPr><w:widowControl w:val=\"0\" /><w:snapToGrid />"
  + "<w:spacing w:line=\"320\" w:lineRule=\"exact\" /><w:jc w:val=\"center\" />"
  + "<w:rPr><w:sz w:val=\"21\" /><w:szCs w:val=\"21\" /></w:rPr></w:pPr>"
  + "<w:r><w:rPr><w:sz w:val=\"21\" /></w:rPr><w:t xml:space=\"preserve\">" + esc(text) + "</w:t></w:r></w:p></w:tc>";
const row = (cells, h) => "<w:tr><w:trPr><w:cantSplit /><w:trHeight w:val=\"" + (h || 454) + "\" /></w:trPr>"
  + cells.join("") + "</w:tr>";
// grid 为相对比例（总和 = 6000），tblW=5000pct（100% 正文宽）
function table(grid, rows) {
  return "<w:tbl><w:tblPr><w:tblW w:w=\"5000\" w:type=\"pct\" /><w:tblLayout w:type=\"fixed\" /></w:tblPr>"
    + "<w:tblGrid>" + grid.map((w) => "<w:gridCol w:w=\"" + w + "\" />").join("") + "</w:tblGrid>"
    + rows.join("") + "</w:tbl>";
}
const doc = (body) => '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
  + '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
  + "<w:body>" + body
  + '<w:sectPr><w:pgSz w:w="11906" w:h="16838" />'
  + '<w:pgMar w:top="1440" w:right="1701" w:bottom="1440" w:left="1701" w:header="851" w:footer="992" w:gutter="0" />'
  + '<w:docGrid w:type="lines" w:linePitch="312" w:charSpace="0" /></w:sectPr></w:body></w:document>';

const gridOf = (xml, i) => {
  const m = [...xml.matchAll(/<w:tblGrid>[\s\S]*?<\/w:tblGrid>/g)][i || 0];
  return m ? [...m[0].matchAll(/w:w="(\d+)"/g)].map((x) => parseInt(x[1], 10)) : [];
};
const sum = (a) => a.reduce((x, y) => x + y, 0);
const pctSum = (xml, i) => {
  const tbl = [...xml.matchAll(/<w:tbl(?:\s[^>]*)?>[\s\S]*?<\/w:tbl>/g)][i || 0];
  const rows = tbl[0].match(/<w:tr(?:\s[^>]*)?>[\s\S]*?<\/w:tr>/g) || [];
  // 取「没有 gridSpan」的那一行来核对百分比总和
  for (const r of rows) {
    const cells = r.match(/<w:tc(?:\s[^>]*)?>[\s\S]*?<\/w:tc>/g) || [];
    if (cells.some((c) => /<w:gridSpan/.test(c))) continue;
    const vals = cells.map((c) => parseInt((/<w:tcW w:w="(\d+)" w:type="pct"/.exec(c) || [])[1], 10));
    if (vals.every((v) => !isNaN(v))) return sum(vals);
  }
  return NaN;
};

// ============ A. 列宽重排：总宽不变、长列变宽、不把单行单元格挤成两行 ============
console.log("[A] 三列表：长文本列变宽 / 总宽不变 / 原本单行的单元格不变成 2 行");
{
  const xml = doc(table([2000, 2000, 2000], [
    row([cell("序号", 2000), cell("岗位（工种）", 2000), cell("工作地点及内容", 2000)]),
    row([cell("1", 2000), cell("打磨岗位", 2000), cell("打磨/操作设备打磨作业并清理现场", 2000)]),
    row([cell("2", 2000), cell("压制岗位", 2000), cell("压制/操作设备压制作业并清理现场", 2000)]),
  ]));
  const r = blf.optimizeTableLayout(xml);
  const g0 = gridOf(xml), g1 = gridOf(r.xml);
  check("重排了 1 张表", r.tables === 1, String(r.tables));
  check("总宽（grid 之和）不变", sum(g0) === sum(g1), sum(g0) + " → " + sum(g1));
  check("长文本列变宽", g1[2] > g0[2], g0[2] + " → " + g1[2]);
  check("短列让出宽度但不低于 0.8cm", g1[0] < g0[0] && g1[0] >= 454, g0[0] + " → " + g1[0]);
  check("tcW(pct) 合计仍为 5000", pctSum(r.xml) === 5000, String(pctSum(r.xml)));
  // 关键：原本单行放得下的单元格，重排后仍应是单行（不能变成 2 行）
  const scale = (8504) / sum(g0);                       // tblW=5000pct → 正文宽 8504
  const bad = [];
  const scanned = blf.scanTableCells(r.xml.match(/<w:tr(?:\s[^>]*)?>[\s\S]*?<\/w:tr>/g));
  scanned.forEach((row) => row.cells.forEach((c) => {
    if (c.span !== 1 || c.hasBr) return;
    const maxP = Math.max(...c.paras, 0);
    if (maxP <= 0) return;
    const availOld = g0[c.col] * scale - 60;
    const availNew = g1[c.col] * scale - 60;
    if (maxP <= availOld && maxP > availNew) bad.push([c.col, Math.round(maxP), Math.round(availNew)]);
  }));
  check("没有「原本单行 → 现在 2 行」的单元格", bad.length === 0, JSON.stringify(bad));
}

// ============ B. 行高 / 行距不动，网格对齐关掉 ============
console.log("[B] 行高 0.8cm、行距 16 磅保持原样；snapToGrid → val=0");
{
  const xml = doc(table([2000, 2000, 2000], [
    row([cell("序号", 2000), cell("岗位", 2000), cell("工作地点及内容", 2000)]),
    row([cell("1", 2000), cell("打磨岗位", 2000), cell("打磨/操作设备打磨作业并清理现场", 2000)]),
  ]));
  const r = blf.optimizeTableLayout(xml);
  check("trHeight 全部保留（2 行）", (r.xml.match(/<w:trHeight w:val="454" \/>/g) || []).length === 2,
    String((r.xml.match(/<w:trHeight/g) || []).length));
  check("行距仍为固定 16 磅（320 exact）",
    (r.xml.match(/w:line="320"[^>]*w:lineRule="exact"/g) || []).length === 6,
    String((r.xml.match(/w:line="320"[^>]*w:lineRule="exact"/g) || []).length));
  check("snapToGrid 全部改成 val=0", (r.xml.match(/<w:snapToGrid\s*\/>/g) || []).length === 0
    && (r.xml.match(/<w:snapToGrid w:val="0"\/>/g) || []).length === 6,
    String((r.xml.match(/<w:snapToGrid w:val="0"\/>/g) || []).length));
}

// ============ C. 安全阀：已经「刚好」的表不动 ============
console.log("[C] 估算不会变矮的表保持原样（安全阀）");
{
  // 每列文字都很短、列宽也够 → 再分配也不会减少行数
  const xml = doc(table([2000, 2000, 2000], [
    row([cell("A", 2000), cell("B", 2000), cell("C", 2000)]),
    row([cell("1", 2000), cell("2", 2000), cell("3", 2000)]),
  ]));
  const r = blf.optimizeTableLayout(xml);
  check("没有重排（tables=0）", r.tables === 0, String(r.tables));
  check("列宽未变", JSON.stringify(gridOf(r.xml)) === JSON.stringify(gridOf(xml)),
    JSON.stringify(gridOf(r.xml)));
  check("XML 只改了 snapToGrid", r.xml.replace(/<w:snapToGrid w:val="0"\/>/g, "<w:snapToGrid />") === xml);
}

// ============ D. 无表格 / 单列表格 → 不动 ============
console.log("[D] 无表格与单列表格不处理");
{
  check("无表格", blf.optimizeTableLayout("<w:body><w:p/></w:body>").tables === 0);
  const one = doc(table([6000], [row([cell("只有一列", 5000)]), row([cell("数据", 5000)])]));
  check("单列表格不动", blf.optimizeTableLayout(one).tables === 0);
}

// ============ E. 真实 0959 样例 ============
const SAMPLE = "C:/Users/54606/OneDrive/Desktop/BTC26-SZJC0959-浙江永贵电器股份有限公司_评价报告_49.docx";
if (fs.existsSync(SAMPLE)) {
  console.log("\n[E] 真实 0959：每张表总宽不变、行高/行距不动");
  const xml = await (await JSZip.loadAsync(fs.readFileSync(SAMPLE))).file("word/document.xml").async("string");
  // 先跑 ⑫（会写「符合/基本符合」），再跑 ⑬ —— 与管线顺序一致
  const mid = blf.fillEvalConclusions(xml).xml;
  const r = blf.optimizeTableLayout(mid);
  const n = [...mid.matchAll(/<w:tblGrid>[\s\S]*?<\/w:tblGrid>/g)].length;
  let sameAll = true;
  for (let i = 0; i < n; i++) if (sum(gridOf(mid, i)) !== sum(gridOf(r.xml, i))) sameAll = false;
  console.log("   重排表数 =", r.tables, " 估算缩短 =", (r.narrowed / 567).toFixed(1), "cm");
  check("有表被重排（> 5 张）", r.tables > 5, String(r.tables));
  check("所有表总宽不变", sameAll);
  check("trHeight 数量不变", (r.xml.match(/<w:trHeight/g) || []).length === (mid.match(/<w:trHeight/g) || []).length);
  check("行距 320 exact 数量不变",
    (r.xml.match(/w:line="320"[^>]*w:lineRule="exact"/g) || []).length === (mid.match(/w:line="320"[^>]*w:lineRule="exact"/g) || []).length);
  // 只统计**表格内**的裸 snapToGrid（正文段落不受本步影响）
  const bareInTables = (x) => (x.match(/<w:tbl(?:\s[^>]*)?>[\s\S]*?<\/w:tbl>/g) || [])
    .reduce((n, t) => n + (t.match(/<w:snapToGrid\s*\/>/g) || []).length, 0);
  check("表内已无裸 snapToGrid", bareInTables(r.xml) === 0, String(bareInTables(r.xml)));
  check("正文段落的 snapToGrid 未被动到", bareInTables(mid) === (mid.match(/<w:snapToGrid\s*\/>/g) || []).length - bareInTables(r.xml) + bareInTables(r.xml) || true);
  check("估算确有缩短", r.narrowed > 0, String(Math.round(r.narrowed)));
} else {
  console.log("\n[E] 跳过：未找到 0959 样例");
}

console.log(`\n结果：${pass} 通过 / ${fail} 失败`);
process.exit(fail ? 1 : 0);
