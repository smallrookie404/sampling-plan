import fs from "node:fs";
import path from "node:path";
import vm from "node:vm";
import { fileURLToPath } from "node:url";
import JSZip from "jszip";

const here = path.dirname(fileURLToPath(import.meta.url));
const APP = path.resolve(here, "..");
const blfSrc = fs.readFileSync(path.join(APP, "js/blf.js"), "utf8");
const SAMPLE = process.argv[2] || process.env.PJBG_SAMPLE
  || path.join(process.env.TEMP || "/tmp", "sampling-plan-login", "pjbg.docx");

let pass = 0, fail = 0;
function check(name, cond, extra) {
  if (cond) { pass++; console.log("  ok  " + name); }
  else { fail++; console.log("  FAIL " + name + (extra ? " :: " + extra : "")); }
}

// ---- 最小 DOM 桩：让 blf.js 顶层能加载（bind() 会找不到元素，但 $ 返回 null 已容错） ----
const noop = () => {};
const fakeEl = { style: {}, classList: { add: noop, remove: noop, toggle: noop }, addEventListener: noop,
  appendChild: noop, removeChild: noop, setAttribute: noop, getAttribute: () => null, hasAttribute: () => false,
  querySelector: () => null, querySelectorAll: () => [], innerHTML: "", textContent: "", disabled: false, click: noop };
const documentStub = {
  readyState: "complete",
  getElementById: () => null,
  querySelector: () => null,
  querySelectorAll: () => [],
  createElement: () => Object.assign({}, fakeEl),
  addEventListener: noop,
  body: Object.assign({}, fakeEl),
  head: Object.assign({}, fakeEl),
  documentElement: Object.assign({}, fakeEl),
};
const sandbox = {
  console,
  document: documentStub,
  DOMParser: class { parseFromString() { return { getElementsByTagNameNS: () => [], }; } },
  setTimeout, clearTimeout,
  fetch: () => Promise.reject(new Error("no fetch")),
  URL: { createObjectURL: () => "blob:x", revokeObjectURL: noop },
  JSZip: JSZip,   // 让 ensureJSZip() 直接命中，避免走 loadScript（桩里不会触发 onload）
};
sandbox.self = sandbox;
sandbox.window = sandbox;
vm.createContext(sandbox);
vm.runInContext(blfSrc, sandbox, { filename: "blf.js" });

const blf = sandbox.SamplingBlf;
check("SamplingBlf 钩子已暴露", !!blf && typeof blf.patchPjbgDocx === "function");

// ---- 解出 document.xml，跑纯函数 ----
const buf = fs.readFileSync(SAMPLE);
const zip = await JSZip.loadAsync(buf);
const xml = await zip.file("word/document.xml").async("string");

console.log("\n[1] sumExposedWorkers（表8-1 作业人数去重合计）");
const sum = blf.sumExposedWorkers(xml);
console.log("   ->", JSON.stringify(sum));
check("能定位表8-1", !!sum);
check("去重合计为 91", sum && sum.total === 91, sum && String(sum.total));
check("去重后 20 个岗位", sum && sum.groups === 20, sum && String(sum.groups));

console.log("\n[2] readBasicInfoTable（表3-1 两个人数）");
const basic = blf.readBasicInfoTable(xml);
console.log("   ->", JSON.stringify({ staffText: basic && basic.staffText, contactText: basic && basic.contactText }));
check("在册职工人数 = 410人", basic && basic.staffText === "410人", basic && basic.staffText);
check("接触职业病危害人数原值 = 53人", basic && basic.contactText === "53人", basic && basic.contactText);

console.log("\n[3a] fixNoteAndContactCount — 非噪声岗位非空（拆两段）");
const JOBS = "金工成型车间打磨岗位、焊接岗位；连接器车间总装岗位";
const r = blf.fixNoteAndContactCount(xml, JOBS);
check("有改动", r.changed > 0, String(r.changed));
check("staffVal=410人", r.staffVal === "410人", r.staffVal);
check("contactVal=91人", r.contactVal === "91人", r.contactVal);
check("表8-1 合计=91", r.total === 91, String(r.total));
check("jobs 原样透传", r.jobs === JOBS, r.jobs);

const outXml = r.xml;
// 抽出表4-1 之后的注段落（可能有 2 段）
function noteParas(x) {
  const tbls = [...x.matchAll(/<w:tbl(?:\s[^>]*)?>[\s\S]*?<\/w:tbl>/g)];
  let t41 = null;
  for (const m of tbls) {
    const h = (m[0].match(/<w:tr(?:\s[^>]*)?>[\s\S]*?<\/w:tr>/) || [""])[0];
    if (h.includes("岗位（工种）") && h.includes("职业病危害因素")) { t41 = m; break; }
  }
  if (!t41) return [];
  const tail = x.slice(t41.index + t41[0].length);
  const out = [];
  const re = /<w:p(?:\s[^>]*)?>[\s\S]*?<\/w:p>/g;
  let m;
  while ((m = re.exec(tail))) {
    const t = (m[0].match(/<w:t(?:\s[^>]*)?>([\s\S]*?)<\/w:t>/g) || [])
      .map((z) => z.replace(/<[^>]+>/g, "")).join("");
    if (!t) continue;
    if (t.indexOf("注：") >= 0 || t.indexOf("接触的噪声经过验证性检测") >= 0) { out.push(t); }
    else break;
  }
  return out;
}
const paras = noteParas(outXml);
console.log("   段落数 =", paras.length);
paras.forEach((p, i) => console.log("   [" + i + "]", p.slice(0, 60) + (p.length > 60 ? "…" : "")));
check("注段落拆成 2 段", paras.length === 2, String(paras.length));
check("第1段含「职工总数410人」", paras[0] && paras[0].includes("职工总数410人"));
check("第1段含「接触职业病危害人数91人」", paras[0] && paras[0].includes("接触职业病危害人数91人"));
check("第1段含「12:30～17:00。」结尾", paras[0] && paras[0].trim().endsWith("12:30～17:00。"), paras[0] && paras[0].slice(-20));
check("第1段不含「M车间A岗位」", paras[0] && !paras[0].includes("M车间A岗位"));
check("第1段不含「接触的噪声经过验证性检测」", paras[0] && !paras[0].includes("接触的噪声经过验证性检测"));
check("第2段以岗位文本开头", paras[1] && paras[1].startsWith(JOBS), paras[1] && paras[1].slice(0, 40));
check("第2段含 GBZ/T224-2010", paras[1] && paras[1].includes("GBZ/T224-2010"));
check("第2段含「当天工况噪声检测结果小于80dB(A)」", paras[1] && paras[1].includes("当天工况噪声检测结果小于80dB(A)"));
check("第2段不含旧占位「M车间A岗位」", paras[1] && !paras[1].includes("M车间A岗位"));
check("旧「注：无」已不存在", !/注：无/.test(outXml));

// 表3-1 接触职业病危害人数单元格 = 91人
function basicRows(x) {
  const tbls = x.match(/<w:tbl(?:\s[^>]*)?>[\s\S]*?<\/w:tbl>/g) || [];
  for (const t of tbls) {
    if (t.includes("接触职业病危害人数") && t.includes("在册职工人数")) {
      return (t.match(/<w:tr(?:\s[^>]*)?>[\s\S]*?<\/w:tr>/g) || []).map((row) =>
        (row.match(/<w:tc(?:\s[^>]*)?>[\s\S]*?<\/w:tc>/g) || []).map((c) =>
          (c.match(/<w:t(?:\s[^>]*)?>[\s\S]*?<\/w:t>/g) || []).map((x) => x.replace(/<[^>]+>/g, "")).join("")));
    }
  }
  return [];
}
const rows = basicRows(outXml);
const contactRow = rows.find((r) => r.includes("接触职业病危害人数")) || [];
const staffRow = rows.find((r) => r.includes("在册职工人数")) || [];
check("表3-1 接触职业病危害人数右邻 = 91人", contactRow[1] === "91人", JSON.stringify(contactRow));
check("表3-1 在册职工人数右邻 保持 410人", staffRow[1] === "410人", JSON.stringify(staffRow));

console.log("\n[3b] 非噪声岗位为空 → 第二段整段删除");
const rEmpty = blf.fixNoteAndContactCount(xml, "");
check("有改动", rEmpty.changed > 0, String(rEmpty.changed));
check("jobs 为空", rEmpty.jobs === "", JSON.stringify(rEmpty.jobs));
const parasEmpty = noteParas(rEmpty.xml);
console.log("   段落数 =", parasEmpty.length);
parasEmpty.forEach((p, i) => console.log("   [" + i + "]", p.slice(0, 50) + (p.length > 50 ? "…" : "")));
check("只剩 1 段（第二段已删除）", parasEmpty.length === 1, String(parasEmpty.length));
check("保留段含「职工总数410人」", parasEmpty[0] && parasEmpty[0].includes("职工总数410人"));
check("已无「接触的噪声经过验证性检测」", !rEmpty.xml.includes("接触的噪声经过验证性检测"));
check("已无「非噪声作业岗位」", !rEmpty.xml.includes("非噪声作业岗位"));
check("表3-1 接触人数仍写入 91人", (basicRows(rEmpty.xml).find((x) => x.includes("接触职业病危害人数")) || [])[1] === "91人");

console.log("\n[4] patchPjbgDocx（整链路；沙箱无 SamplingUpload → 第二段删除分支）");
const patched = await blf.patchPjbgDocx(buf);
check("返回新 buffer", patched !== buf && patched.byteLength > 0, String(patched.byteLength));
const z2 = await JSZip.loadAsync(patched);
const xml2 = await z2.file("word/document.xml").async("string");
check("重打包后仍含「职工总数410人」", xml2.includes("职工总数410人"));
check("重打包后仍含「接触职业病危害人数91人」", xml2.includes("接触职业病危害人数91人"));
check("其它文件仍在（styles.xml）", !!z2.file("word/styles.xml"));

// ---- [5] fixNitrogenOxide：表4-1 职业病危害因素列 氮氧化物措辞精简 ----
console.log("\n[5] fixNitrogenOxide（表4-1 危害列「氮氧化物」精简）");
const NOX_OLD = "氮氧化物(一氧化氮和二氧化氮):氮氧化物(以二氧化氮计)";
const NOX_NEW = "氮氧化物(一氧化氮和二氧化氮)";
// 5a 混在逗号清单里（真实 1009 报告形态，单 run）
const mkTbl = (cellXml) =>
  "<w:tbl><w:tr>"
  + "<w:tc><w:p><w:r><w:t>单元/工作场所</w:t></w:r></w:p></w:tc>"
  + "<w:tc><w:p><w:r><w:t>岗位（工种）</w:t></w:r></w:p></w:tc>"
  + "<w:tc><w:p><w:r><w:t>职业病危害因素</w:t></w:r></w:p></w:tc>"
  + "</w:tr><w:tr>"
  + "<w:tc><w:p><w:r><w:t>喷漆车间</w:t></w:r></w:p></w:tc>"
  + "<w:tc><w:p><w:r><w:t>喷漆岗位</w:t></w:r></w:p></w:tc>"
  + "<w:tc>" + cellXml + "</w:tc>"
  + "</w:tr></w:tbl>";
const cellSingle = "<w:p><w:r><w:t>" + NOX_OLD + ",电焊烟尘,一氧化碳,噪声</w:t></w:r></w:p>";
const r5a = blf.fixNitrogenOxide(mkTbl(cellSingle));
check("单 run 清单：有改动", r5a.changed === 1, String(r5a.changed));
check("旧短语已消失", !r5a.xml.includes(NOX_OLD));
check("新短语已出现", r5a.xml.includes(NOX_NEW));
check("清单其余项保留（电焊烟尘/一氧化碳/噪声）",
  r5a.xml.includes(NOX_NEW + ",电焊烟尘,一氧化碳,噪声"), r5a.xml.slice(r5a.xml.indexOf("氮氧化物"), r5a.xml.indexOf("氮氧化物") + 60));
// 5b 独占整格
const r5b = blf.fixNitrogenOxide(mkTbl("<w:p><w:r><w:t>" + NOX_OLD + "</w:t></w:r></w:p>"));
check("独占整格：替换为纯新短语", r5b.xml.includes("<w:t>" + NOX_NEW + "</w:t>"), String(r5b.changed));
// 5c 跨 run 拆分
const cellSplit = "<w:p><w:r><w:t>氮氧化物(一氧化氮</w:t></w:r><w:r><w:t>和二氧化氮):氮氧化物(以</w:t></w:r><w:r><w:t>二氧化氮计),紫外辐射</w:t></w:r></w:p>";
const r5c = blf.fixNitrogenOxide(mkTbl(cellSplit));
check("跨 run 拆分：有改动", r5c.changed === 1, String(r5c.changed));
check("跨 run 拆分：旧短语消失", !r5c.xml.includes(NOX_OLD));
check("跨 run 拆分：新短语出现", r5c.xml.includes(NOX_NEW));
// 5d 全角括号/冒号变体
const FULL_OLD = "氮氧化物（一氧化氮和二氧化氮）：氮氧化物（以二氧化氮计）";
const FULL_NEW = "氮氧化物（一氧化氮和二氧化氮）";
const r5d = blf.fixNitrogenOxide(mkTbl("<w:p><w:r><w:t>" + FULL_OLD + ",噪声</w:t></w:r></w:p>"));
check("全角变体：有改动", r5d.changed === 1, String(r5d.changed));
check("全角变体：换为全角新短语", r5d.xml.includes(FULL_NEW + ",噪声"));
// 5e 无命中 → 不改
const r5e = blf.fixNitrogenOxide(mkTbl("<w:p><w:r><w:t>电焊烟尘,一氧化碳</w:t></w:r></w:p>"));
check("无命中：changed=0", r5e.changed === 0, String(r5e.changed));

// ---- [6] 真实样例整链路：氮氧化物短语不应残留旧形态 ----
console.log("\n[6] 整链路：氮氧化物旧形态不残留");
check("重打包后无旧氮氧化物短语", !xml2.includes(NOX_OLD), String((xml2.match(new RegExp(NOX_OLD.replace(/[()]/g, "\\$&"), "g")) || []).length));

console.log("\n结果：" + pass + " 通过 / " + fail + " 失败");
process.exit(fail ? 1 : 0);
