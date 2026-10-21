// 回归测试：检评报告后处理⑫「表8-1 评价结论 / 补救措施」自动判定（2026-10-21）。
//
// 规则：
//   · 检测结果「符合」→ 评价结论「符合」；
//   · 检测结果「不符合」：
//       - 检测项目含「噪声」：表3-6 有「听力防护」→ 补救措施「防噪耳塞」；
//         取「防护性能参数」里的**标称降噪值 X**，85 + X > 该车间/岗位/点位的噪声 LEX → 基本符合，否则不符合；
//       - 检测项目含「高温」→ 不符合；补救措施 = 表3-5 里同车间且「设置地点或岗位」模糊匹配该岗位/点位的设施名称；
//       - 其它项目 → 不符合（补救措施不动）。
//
// 用法：node tests/pjbg.eval.test.mjs
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

// ---- XML 夹具 ----
const esc = (s) => String(s).replace(/&/g, "&amp;").replace(/</g, "&lt;");
const p = (t) => "<w:p><w:r><w:t xml:space=\"preserve\">" + esc(t) + "</w:t></w:r></w:p>";
const tc = (t, pr) => "<w:tc>" + (pr ? "<w:tcPr>" + pr + "</w:tcPr>" : "") + p(t) + "</w:tc>";
const tr = (cells) => "<w:tr>" + cells.join("") + "</w:tr>";
const tbl = (rows) => "<w:tbl><w:tblPr/><w:tblGrid/>" + rows.join("") + "</w:tbl>";

// 表7-3 噪声表（单行表头）
const NOISE_HDR = ["测量编号", "单元/工作场所", "岗位/工种/测量点/对象", "接触时间", "噪声性质",
  "测量范围[dB(A)]", "测量结果[dB(A)]", "LEX,8h[dB(A)]", "LEX,W[dB(A)]", "接触限值[dB(A)]", "单项结论"];
function noiseTbl(rows) {
  return tbl([tr(NOISE_HDR.map((x) => tc(x)))].concat(rows.map((r) => tr(r.map((x) => tc(x))))));
}

// 表3-6（两行表头：职业病防护用品占 2 列）
function ppeTbl(rows) {
  const r0 = tr([tc("职业病防护用品", '<w:gridSpan w:val="2"/>'), tc("生产厂家", '<w:vMerge w:val="restart"/>'),
    tc("型号或规格", '<w:vMerge w:val="restart"/>'), tc("防护性能参数", '<w:vMerge w:val="restart"/>'),
    tc("配置岗位（工种）", '<w:vMerge w:val="restart"/>'), tc("佩戴情况", '<w:vMerge w:val="restart"/>'),
    tc("更换周期", '<w:vMerge w:val="restart"/>'), tc("备注", '<w:vMerge w:val="restart"/>')]);
  const r1 = tr([tc("分类"), tc("类别"), tc("", "<w:vMerge/>"), tc("", "<w:vMerge/>"), tc("", "<w:vMerge/>"),
    tc("", "<w:vMerge/>"), tc("", "<w:vMerge/>"), tc("", "<w:vMerge/>"), tc("", "<w:vMerge/>")]);
  return tbl([r0, r1].concat(rows.map((r) => tr(r.map((x) => tc(x))))));
}

// 表3-5（两行表头：职业病防护设施占 2 列、数量占 2 列）
function facTbl(rows) {
  const r0 = tr([tc("单元/工作场所"), tc("职业病防护设施", '<w:gridSpan w:val="2"/>'), tc("技术参数"),
    tc("设置地点或岗位"), tc("数量（台/套）", '<w:gridSpan w:val="2"/>'), tc("维护情况"), tc("备注")]);
  const r1 = tr([tc(""), tc("类型"), tc("名称"), tc(""), tc(""), tc("总数"), tc("运行"), tc(""), tc("")]);
  return tbl([r0, r1].concat(rows.map((r) => tr(r.map((x) => tc(x))))));
}

// 表8-1（单行表头）
const T81_HDR = ["单元/工作场所", "岗位/工种", "作业人数", "检测点/对象", "检测项目",
  "检测结果", "补救措施", "评价结论", "存在高毒物品"];
function t81(rows) {
  return tbl([tr(T81_HDR.map((x) => tc(x)))].concat(rows.map((r) => tr(r.map((x) => tc(x))))));
}
const row81 = (unit, job, point, item, result) => [unit, job, "1", point, item, result, "-", "-", "否"];

function buildXml(opts) {
  const o = opts || {};
  let s = "";
  if (o.noise !== false) {
    s += noiseTbl(o.noiseRows || [
      ["1", "车间A", "打磨岗位/打磨", "7h/d", "非稳态", "89.5～90.5", "90.5\n90.0\n89.5", "90.0", "-", "85", "不符合"],
      ["2", "车间A", "冲压岗位/冲压", "7h/d", "非稳态", "94.5～95.5", "95.5\n95.0\n94.5", "95.0", "-", "85", "不符合"],
      ["3", "车间A", "包装岗位/包装", "7h/d", "非稳态", "70.0～71.0", "71.0\n70.5\n70.0", "70.5", "-", "85", "符合"],
    ]);
  }
  if (o.hearing !== false) {
    s += ppeTbl([["听力防护", "防噪耳塞", "3M", "1270", "NRR=24dB，标称降噪值8.5dB",
      "车间A/打磨岗位,冲压岗位", "规范配戴", "定期更换", "-"]]);
  }
  if (o.fac !== false) {
    s += facTbl([
      ["车间A", "通风降温", "工业风扇", "-", "打磨", "1", "1", "定期维护", "-"],
      ["车间A", "防暑降温", "冷风机", "-", "打磨", "1", "1", "定期维护", "-"],
      ["车间A", "防暑降温", "排风扇", "-", "压制", "1", "1", "定期维护", "-"],
      ["车间A", "吸风排毒", "排风罩", "-", "喷漆", "1", "1", "定期维护", "-"],   // 诱饵：类型不含「降温」
      ["车间B", "通风降温", "轴流风机", "-", "包装", "1", "1", "定期维护", "-"],
    ]);
  }
  s += t81(o.rows || []);
  return "<w:body>" + s + "</w:body>";
}

const rowsOf = (xml) => {
  const out = [];
  (xml.match(/<w:tr(?:\s[^>]*)?>[\s\S]*?<\/w:tr>/g) || []).forEach((r) => {
    out.push((r.match(/<w:tc(?:\s[^>]*)?>[\s\S]*?<\/w:tc>/g) || []).map((c) => {
      const t = c.match(/<w:t(?:\s[^>]*)?>([\s\S]*?)<\/w:t>/g) || [];
      return t.map((z) => z.replace(/<[^>]+>/g, "")).join("");
    }));
  });
  return out;
};
// 取表8-1 数据行（跳过表头）
function t81Rows(xml) {
  const rs = rowsOf(xml);
  const i = rs.findIndex((r) => r.indexOf("评价结论") >= 0);
  return i < 0 ? [] : rs.slice(i + 1);
}

// ============ A. 符合 → 符合 ============
console.log("[A] 检测结果「符合」→ 评价结论「符合」");
{
  const r = blf.fillEvalConclusions(buildXml({ rows: [row81("车间A", "包装岗位", "包装", "噪声", "符合")] }));
  const rows = t81Rows(r.xml);
  check("评价结论 = 符合", rows[0] && rows[0][7] === "符合", rows[0] && rows[0][7]);
  check("补救措施保持 -", rows[0] && rows[0][6] === "-", rows[0] && rows[0][6]);
  check("rows = 1", r.rows === 1, String(r.rows));
}

// ============ B. 噪声不符合 + 降噪值够用 → 基本符合 ============
console.log("[B] 噪声不符合：85+8.5=93.5 > LEX 90 → 基本符合 + 防噪耳塞");
{
  const r = blf.fillEvalConclusions(buildXml({ rows: [row81("车间A", "打磨岗位", "打磨", "噪声", "不符合")] }));
  const rows = t81Rows(r.xml);
  check("评价结论 = 基本符合", rows[0] && rows[0][7] === "基本符合", rows[0] && rows[0][7]);
  check("补救措施 = 防噪耳塞", rows[0] && rows[0][6] === "防噪耳塞", rows[0] && rows[0][6]);
  check("basic = 1", r.basic === 1, String(r.basic));
}

// ============ C. 噪声不符合 + 降噪值不够 → 不符合 ============
console.log("[C] 噪声不符合：85+8.5=93.5 < LEX 95 → 不符合（补救措施仍填防噪耳塞）");
{
  const r = blf.fillEvalConclusions(buildXml({ rows: [row81("车间A", "冲压岗位", "冲压", "噪声", "不符合")] }));
  const rows = t81Rows(r.xml);
  check("评价结论 = 不符合", rows[0] && rows[0][7] === "不符合", rows[0] && rows[0][7]);
  check("补救措施 = 防噪耳塞", rows[0] && rows[0][6] === "防噪耳塞", rows[0] && rows[0][6]);
  check("bad = 1", r.bad === 1, String(r.bad));
}

// ============ D. 没有听力防护 → 不符合，补救措施不动 ============
console.log("[D] 表3-6 无「听力防护」→ 不符合，补救措施保持 -");
{
  const r = blf.fillEvalConclusions(buildXml({ hearing: false, rows: [row81("车间A", "打磨岗位", "打磨", "噪声", "不符合")] }));
  const rows = t81Rows(r.xml);
  check("评价结论 = 不符合", rows[0] && rows[0][7] === "不符合", rows[0] && rows[0][7]);
  check("补救措施保持 -", rows[0] && rows[0][6] === "-", rows[0] && rows[0][6]);
}

// ============ E. 高温不符合 + 表3-5 有「降温」类设施（通风降温 / 防暑降温都算） ============
console.log("[E] 高温不符合 → 不符合 + 补救措施 = 表3-5「降温」类设施名称（多个用、）");
{
  const r = blf.fillEvalConclusions(buildXml({ rows: [row81("车间A", "打磨岗位", "打磨", "高温", "不符合")] }));
  const rows = t81Rows(r.xml);
  check("评价结论 = 不符合", rows[0] && rows[0][7] === "不符合", rows[0] && rows[0][7]);
  check("补救措施 = 工业风扇、冷风机（通风降温 + 防暑降温 都算）",
    rows[0] && rows[0][6] === "工业风扇、冷风机", rows[0] && rows[0][6]);
}

// ============ E2. 只有「防暑降温」也算 ============
console.log("[E2] 高温：该岗位只有「防暑降温」类设施 → 也录入");
{
  const r = blf.fillEvalConclusions(buildXml({ rows: [row81("车间A", "压制岗位", "压制", "高温", "不符合")] }));
  const rows = t81Rows(r.xml);
  check("补救措施 = 排风扇", rows[0] && rows[0][6] === "排风扇", rows[0] && rows[0][6]);
}

// ============ E3. 换车间：车间B/包装 → 轴流风机 ============
console.log("[E3] 高温：车间B 包装岗位 → 轴流风机（通风降温）");
{
  const r = blf.fillEvalConclusions(buildXml({ rows: [row81("车间B", "包装岗位", "包装", "高温", "不符合")] }));
  const rows = t81Rows(r.xml);
  check("补救措施 = 轴流风机", rows[0] && rows[0][6] === "轴流风机", rows[0] && rows[0][6]);
}

// ============ E4. 类型不含「降温」（吸风排毒）→ 不录 ============
console.log("[E4] 高温：该岗位只有「吸风排毒」类设施 → 补救措施不动");
{
  const r = blf.fillEvalConclusions(buildXml({ rows: [row81("车间A", "喷漆岗位", "喷漆", "高温", "不符合")] }));
  const rows = t81Rows(r.xml);
  check("评价结论 = 不符合", rows[0] && rows[0][7] === "不符合", rows[0] && rows[0][7]);
  check("补救措施保持 -（排风罩类型是「吸风排毒」，不含「降温」）", rows[0] && rows[0][6] === "-", rows[0] && rows[0][6]);
}

// ============ F. 高温不符合 + 表3-5 无匹配 ============
console.log("[F] 高温不符合：表3-5 无同车间/同地点设施 → 补救措施不动");
{
  const r = blf.fillEvalConclusions(buildXml({ rows: [row81("车间C", "注塑岗位", "注塑", "高温", "不符合")] }));
  const rows = t81Rows(r.xml);
  check("评价结论 = 不符合", rows[0] && rows[0][7] === "不符合", rows[0] && rows[0][7]);
  check("补救措施保持 -", rows[0] && rows[0][6] === "-", rows[0] && rows[0][6]);
}

// ============ G. 其它项目不符合 → 不符合 ============
console.log("[G] 其它检测项目不符合 → 评价结论 不符合");
{
  const r = blf.fillEvalConclusions(buildXml({ rows: [row81("车间A", "打磨岗位", "打磨", "砂轮磨尘(总尘)", "不符合")] }));
  const rows = t81Rows(r.xml);
  check("评价结论 = 不符合", rows[0] && rows[0][7] === "不符合", rows[0] && rows[0][7]);
  check("补救措施保持 -", rows[0] && rows[0][6] === "-", rows[0] && rows[0][6]);
}

// ============ H. 真实样例 ============
const S1 = "C:/Users/54606/OneDrive/Desktop/BTC26-SZJC0959-浙江永贵电器股份有限公司_评价报告_49.docx";
if (fs.existsSync(S1)) {
  console.log("\n[H] 真实 0959：6 行不符合 → 5 基本符合 + 1 不符合（去毛刺2 LEX 93.8 > 93.5）");
  const xml = await (await JSZip.loadAsync(fs.readFileSync(S1))).file("word/document.xml").async("string");
  const h = blf.readHearingProtection(xml);
  check("听力防护 标称降噪值 8.5 / 名称 防噪耳塞", h.has && h.snr === 8.5 && h.name === "防噪耳塞", JSON.stringify(h));
  const r = blf.fillEvalConclusions(xml);
  console.log("   ->", JSON.stringify({ changed: r.changed, rows: r.rows, ok: r.ok, basic: r.basic, bad: r.bad }));
  check("42 行判定（36 符合 / 5 基本符合 / 1 不符合）",
    r.rows === 42 && r.ok === 36 && r.basic === 5 && r.bad === 1, JSON.stringify(r));
  const rows = t81Rows(r.xml).filter((x) => x[5] === "不符合");
  check("6 行不符合全部写了防噪耳塞", rows.length === 6 && rows.every((x) => x[6] === "防噪耳塞"),
    JSON.stringify(rows.map((x) => [x[1], x[6], x[7]])));
  check("去毛刺2 判为不符合（LEX 93.8）", rows.some((x) => x[7] === "不符合"), JSON.stringify(rows.map((x) => x[7])));
} else {
  console.log("\n[H] 跳过：未找到 0959 样例");
}

// ============ I. 幂等：重复跑不重复改 ============
console.log("\n[I] 幂等：对已填好的表再跑一次 → 0 改动，判定统计仍返回");
{
  const xml = buildXml({ rows: [row81("车间A", "打磨岗位", "打磨", "噪声", "不符合"),
    row81("车间A", "包装岗位", "包装", "噪声", "符合")] });
  const r1 = blf.fillEvalConclusions(xml);
  const r2 = blf.fillEvalConclusions(r1.xml);
  check("第一次有改动", r1.changed > 0, String(r1.changed));
  check("第二次 0 改动", r2.changed === 0, String(r2.changed));
  check("第二次仍返回判定行数（2 行）", r2.rows === 2, String(r2.rows));
  check("XML 与第一次结果一致", r2.xml === r1.xml);
}

console.log(`\n结果：${pass} 通过 / ${fail} 失败`);
process.exit(fail ? 1 : 0);
