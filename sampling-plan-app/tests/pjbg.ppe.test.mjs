// 回归测试：检评报告后处理⑧「表3-6 个人使用的职业病防护用品 —— 同规格行合并」。
//
// 规则（用户 2026-10-20）：
//   · 多行的「分类 / 类别 / 生产厂家 / 型号或规格 / 防护性能参数」五列内容分别相同时 → 合并成一行；
//   · 「配置岗位（工种）」：「/」前的部分（车间）提取在一起、去重后用「,」连接；「/」后的部分（岗位）
//     同样处理；两者之间用「/」连接；
//   · 其余列（佩戴情况 / 更换周期 / 备注）合并相同值去重。
//
// 覆盖：A 两行同规格合并（含配置岗位前后段拼接）；B 三行合并 + 重复岗位去重；C 型号不同不合并；
//       D 配置岗位为「/」时保持原样；E 非表3-6 的表不动；F 真实 0959 样例（8 行 → 5 行）。
//
// 用法：node tests/pjbg.ppe.test.mjs
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
  URL: { createObjectURL: () => "blob:x", revokeObjectURL: noop }, JSZip, TextDecoder,
};
sandbox.self = sandbox; sandbox.window = sandbox;
vm.createContext(sandbox);
vm.runInContext(blfSrc, sandbox, { filename: "blf.js" });
const blf = sandbox.SamplingBlf;

// ---- 夹具 ----
const esc = (s) => String(s).replace(/&/g, "&amp;").replace(/</g, "&lt;");
const p = (t) => "<w:p><w:r><w:t xml:space=\"preserve\">" + esc(t) + "</w:t></w:r></w:p>";
const tc = (t, pr) => "<w:tc>" + (pr ? "<w:tcPr>" + pr + "</w:tcPr>" : "") + p(t) + "</w:tc>";
const tr = (cells) => "<w:tr>" + cells.join("") + "</w:tr>";
const tbl = (rows) => "<w:tbl><w:tblPr/><w:tblGrid/>" + rows.join("") + "</w:tbl>";

// 表3-6 表头：第 1 行 8 格（「职业病防护用品」占 2 列），第 2 行 9 格（分类/类别 + 7 个纵向合并续格）
function ppeHeader() {
  const r0 = "<w:tr>"
    + tc("职业病防护用品", '<w:gridSpan w:val="2"/>')
    + tc("生产厂家", '<w:vMerge w:val="restart"/>')
    + tc("型号或规格", '<w:vMerge w:val="restart"/>')
    + tc("防护性能参数", '<w:vMerge w:val="restart"/>')
    + tc("配置岗位（工种）", '<w:vMerge w:val="restart"/>')
    + tc("佩戴情况", '<w:vMerge w:val="restart"/>')
    + tc("更换周期", '<w:vMerge w:val="restart"/>')
    + tc("备注", '<w:vMerge w:val="restart"/>')
    + "</w:tr>";
  const cont = '<w:vMerge/>';
  const r1 = "<w:tr>" + tc("分类") + tc("类别") + tc("", cont) + tc("", cont) + tc("", cont)
    + tc("", cont) + tc("", cont) + tc("", cont) + tc("", cont) + "</w:tr>";
  return [r0, r1];
}
// 数据行：分类 / 类别 / 生产厂家 / 型号或规格 / 防护性能参数 / 配置岗位 / 佩戴情况 / 更换周期 / 备注
function ppeRow(cat, type, maker, model, param, job, wear, cycle, note) {
  return tr([cat, type, maker, model, param, job, wear || "规范配戴", cycle || "定期更换", note || "-"].map((x) => tc(x)));
}
const body = (rows) => tbl(ppeHeader().concat(rows));

function tableRowsText(xml) {
  const rows = xml.match(/<w:tr(?:\s[^>]*)?>[\s\S]*?<\/w:tr>/g) || [];
  return rows.map((r) => (r.match(/<w:tc(?:\s[^>]*)?>[\s\S]*?<\/w:tc>/g) || []).map((c) => {
    const t = c.match(/<w:t(?:\s[^>]*)?>([\s\S]*?)<\/w:t>/g) || [];
    return t.map((z) => z.replace(/<[^>]+>/g, "")).join("");
  }));
}

// ============ A. 两行同规格 → 合并 ============
console.log("[A] 两行同规格合并（配置岗位前后段各自拼接）");
{
  const xml = body([
    ppeRow("听力防护", "防噪耳塞", "3M", "1270", "NRR=24dB", "连接器车间/针孔件制造岗位"),
    ppeRow("听力防护", "防噪耳塞", "3M", "1270", "NRR=24dB", "金工成型车间/打标岗位,压制岗位"),
  ]);
  const r = blf.mergePpeRows(xml);
  const rows = tableRowsText(r.xml);
  check("合并 1 行", r.changed === 1, String(r.changed));
  check("表体 3 行（2 表头 + 1 数据）", rows.length === 3, String(rows.length));
  check("配置岗位 = 车间,车间/岗位,岗位",
    rows[2][5] === "连接器车间,金工成型车间/针孔件制造岗位,打标岗位,压制岗位", rows[2][5]);
  check("分类/类别/厂家/型号/参数保留", rows[2].slice(0, 5).join("|") === "听力防护|防噪耳塞|3M|1270|NRR=24dB", rows[2].join("|"));
  check("佩戴情况/更换周期/备注 去重", rows[2][6] === "规范配戴" && rows[2][7] === "定期更换" && rows[2][8] === "-",
    JSON.stringify(rows[2].slice(6)));
}

// ============ B. 三行合并 + 重复岗位去重 ============
console.log("[B] 三行合并 + 重复岗位/车间去重");
{
  const xml = body([
    ppeRow("呼吸防护", "防尘口罩", "-", "-", "KN95", "车间A/台钻岗位"),
    ppeRow("呼吸防护", "防尘口罩", "-", "-", "KN95", "车间B/台钻岗位,打磨岗位"),
    ppeRow("呼吸防护", "防尘口罩", "-", "-", "KN95", "车间A/普车岗位"),
  ]);
  const r = blf.mergePpeRows(xml);
  const rows = tableRowsText(r.xml);
  check("合并 2 行", r.changed === 2, String(r.changed));
  check("只剩 1 个数据行", rows.length === 3, String(rows.length));
  check("车间去重保序（车间A 只出现一次）", rows[2][5] === "车间A,车间B/台钻岗位,打磨岗位,普车岗位", rows[2][5]);
}

// ============ C. 型号不同 → 不合并 ============
console.log("[C] 型号或规格不同 → 不合并");
{
  const xml = body([
    ppeRow("呼吸防护", "防尘口罩", "-", "KN95", "APF=10", "车间A/岗位1"),
    ppeRow("呼吸防护", "防尘口罩", "-", "KN90", "APF=10", "车间B/岗位2"),
  ]);
  const r = blf.mergePpeRows(xml);
  check("无合并", r.changed === 0, String(r.changed));
  check("XML 未变", r.xml === xml);
}

// ============ D. 配置岗位为「/」→ 合并后保持「/」 ============
console.log("[D] 配置岗位为「/」：合并后仍是「/」");
{
  const xml = body([
    ppeRow("呼吸防护", "防颗粒物呼吸器", "创维邦", "配KN95滤棉", "KP95", "/"),
    ppeRow("呼吸防护", "防颗粒物呼吸器", "创维邦", "配KN95滤棉", "KP95", "/"),
  ]);
  const r = blf.mergePpeRows(xml);
  const rows = tableRowsText(r.xml);
  check("合并 1 行", r.changed === 1, String(r.changed));
  check("配置岗位仍是「/」", rows[2][5] === "/", rows[2][5]);
}

// ============ E. 非表3-6 的表不动 ============
console.log("[E] 不含「生产厂家/防护性能」的表不动");
{
  const other = tbl([
    tr([tc("单元/工作场所"), tc("岗位（工种）"), tc("人数"), tc("配置岗位（工种）")]),
    tr([tc("车间A"), tc("岗位A"), tc("3"), tc("车间A/岗位A")]),
    tr([tc("车间A"), tc("岗位A"), tc("3"), tc("车间A/岗位A")]),
  ]);
  const r = blf.mergePpeRows(other);
  check("无合并", r.changed === 0 && r.xml === other);
}

// ============ F. 真实 0959 样例 ============
const SAMPLE = "C:/Users/54606/OneDrive/Desktop/BTC26-SZJC0959-浙江永贵电器股份有限公司_评价报告_49.docx";
if (fs.existsSync(SAMPLE)) {
  console.log("\n[F] 真实 0959 报告：表3-6 由 8 行合并为 5 行");
  const xml = await (await JSZip.loadAsync(fs.readFileSync(SAMPLE))).file("word/document.xml").async("string");
  const r = blf.mergePpeRows(xml);
  check("合并 3 行", r.changed === 3, String(r.changed));
  // 取表3-6
  const tbls = r.xml.match(/<w:tbl(?:\s[^>]*)?>[\s\S]*?<\/w:tbl>/g) || [];
  const t36 = tbls.find((t) => t.indexOf("配置岗位") >= 0 && t.indexOf("生产厂家") >= 0);
  const rows = tableRowsText(t36 || "");
  console.log("   行数 =", rows.length);
  check("表体 5 行", rows.length === 5, String(rows.length));
  check("听力防护行合并了 2 个车间",
    rows[2] && rows[2][5] === "连接器车间,金工成型车间/针孔件制造岗位,打标岗位,压制岗位,打磨岗位,数控岗位,加工中心岗位,普车岗位,下料清洗岗位,压铸岗位,冲床岗位,抛丸岗位,喷砂岗位,去毛刺岗位",
    rows[2] && rows[2][5]);
  check("呼吸防护行合并了 3 个车间、岗位去重（台钻岗位只 1 次）",
    rows[3] && rows[3][5].indexOf("连接器车间,金工成型车间,减振器车间/") === 0
      && (rows[3][5].match(/台钻岗位/g) || []).length === 1, rows[3] && rows[3][5]);
  check("眼面防护行未动", rows[4] && rows[4][5] === "金工成型车间/电焊打磨岗位", rows[4] && rows[4][5]);
} else {
  console.log("\n[F] 跳过：未找到 0959 样例（" + SAMPLE + "）");
}

console.log(`\n结果：${pass} 通过 / ${fail} 失败`);
process.exit(fail ? 1 : 0);
