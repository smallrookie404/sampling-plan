// 回归测试：检评报告后处理「噪声表只解析一次」的复用逻辑（2026-10-18 性能重构）。
//
// 验证目标（不依赖真实报告 / 真实 DOMParser）：
//   A. patchPjbgDocx 只调用**一次**噪声解析，且 ②⑤⑥ 三步共用同一结果；
//   B. 优先级正确：getCachedNoiseInfo（缓存）> noiseInfoFromXml（复用已解压 xml）> noiseInfoFromBuf（兜底）；
//   C. 老接口（noiseNonNoiseFromBuf / noiseWorkshopGroupsFromBuf）在**新接口可用时不被调用**
//      —— 这正是「同一份报告被解析 3 次」的根因，必须证明已被绕开；
//   D. ② 按 nonNoisePairs 删掉确认表里的「噪声」；⑥ 在噪声车间岗位为空时删噪声类条目；
//   E. 新接口全部缺失时能退化到老接口（向后兼容）。
//
// 用法：node tests/pjbg.noisereuse.test.mjs
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

// 每个用例都新建沙箱（SamplingUpload 桩不同）
function makeSandbox(uploadStub) {
  const sandbox = {
    console: { log: noop, info: noop, warn: noop, error: noop },
    document: documentStub,
    DOMParser: class { parseFromString() { return { getElementsByTagNameNS: () => [] }; } },
    setTimeout, clearTimeout, fetch: () => Promise.reject(new Error("no fetch")),
    URL: { createObjectURL: () => "blob:x", revokeObjectURL: noop },
    JSZip, TextDecoder,
  };
  sandbox.self = sandbox;
  sandbox.window = sandbox;
  if (uploadStub) sandbox.SamplingUpload = uploadStub;
  vm.createContext(sandbox);
  vm.runInContext(blfSrc, sandbox, { filename: "blf.js" });
  return sandbox.SamplingBlf;
}

// ---- 夹具 XML ----
function esc(s) { return String(s).replace(/&/g, "&amp;").replace(/</g, "&lt;"); }
function p(text, runs) {
  return "<w:p>" + (runs || [text]).map((t) => '<w:r><w:t xml:space="preserve">' + esc(t) + "</w:t></w:r>").join("") + "</w:p>";
}
function tbl(headers, rows) {
  const hc = headers.map((h) => "<w:tc>" + p(h) + "</w:tc>").join("");
  const rc = rows.map((r) => "<w:tr>" + r.map((c) => "<w:tc>" + p(c) + "</w:tc>").join("") + "</w:tr>").join("");
  return "<w:tbl><w:tr>" + hc + "</w:tr>" + rc + "</w:tbl>";
}

// 确认表：单元/工作场所 | 岗位/工种 | 职业病危害因素 | 检测项目 | 备注
// ⚠ 危害因素列的多个因素用**半角逗号**分隔（dropNoiseItem 只按 ,/， 拆分，不认 、）
const CONFIRM_HDR = ["单元/工作场所", "岗位/工种", "职业病危害因素", "检测项目", "备注"];
const CONFIRM_ROWS = [
  ["M车间", "打磨岗位", "噪声,电焊烟尘", "噪声", ""],
  ["M车间", "装配岗位", "噪声,紫外辐射", "噪声", ""],
];
const NOISE_ITEM = '《听力防护装备的选择、使用和维护》(GB 23466-2025)';

// 取确认表「职业病危害因素」列的各单元格文本（用 blf 自己的实现，顺带覆盖它）
function hazardColTexts(blf, xml) {
  const r = blf.collectConfirmColTexts(xml, ['职业病危害因素']);
  return (r && r.texts) || [];
}

function buildBodyXml() {
  let s = p("1.2.3评价依据");
  s += p("（1）《中华人民共和国职业病防治法》；", ["（1）", "《中华人民共和国职业病防治法》；"]);
  s += p("（2）" + NOISE_ITEM + "；", ["（2）", NOISE_ITEM, "；"]);
  s += p("（3）《工作场所职业病危害因素检测工作规范》；", ["（3）", "《工作场所职业病危害因素检测工作规范》；"]);
  s += p("表4-1 岗位（工种）设置、劳动者作业情况及职业病危害因素识别情况");
  s += tbl(CONFIRM_HDR, CONFIRM_ROWS);
  return s;
}

async function buildDocx(bodyXml) {
  const xml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    + '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
    + "<w:body>" + bodyXml + "</w:body></w:document>";
  const zip = new JSZip();
  zip.file("[Content_Types].xml", '<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"/>');
  zip.file("word/document.xml", xml);
  return await zip.generateAsync({ type: "arraybuffer", compression: "DEFLATE" });
}

async function readXml(buf) {
  const zip = await JSZip.loadAsync(buf);
  return await zip.file("word/document.xml").async("string");
}

// 把 <w:t> 内容拼起来（跨 run 断言用）
function allText(xml) {
  const out = [];
  const re = /<w:t(?:\s[^>]*)?>([\s\S]*?)<\/w:t>/g;
  let m;
  while ((m = re.exec(xml))) out.push(m[1]);
  return out.join("");
}

// ---- 信息对象：打磨岗位是非噪声，装配岗位是噪声（不在 nonNoise 里） ----
const INFO = {
  rows: [], ok: true,
  nonNoiseGroups: [{ unit: "M车间", jobs: ["打磨岗位"] }],
  nonNoisePairs: [{ unit: "M车间", job: "打磨岗位" }],
  nonNoiseText: "M车间打磨岗位",
  noiseWorkshopGroups: [],      // 空 → ⑥ 删噪声类条目
};

(async () => {
  const buf = await buildDocx(buildBodyXml());

  // ================= A. 新接口：只解析一次，三步共用 =================
  console.log("[A] 新接口（noiseInfoFromXml）：只解析一次 + ②⑥ 生效");
  {
    const calls = { fromXml: 0, fromBuf: 0, cached: 0, oldPairs: 0, oldWorkshop: 0, oldText: 0 };
    const stub = {
      getCachedNoiseInfo() { calls.cached++; return null; },
      noiseInfoFromXml() { calls.fromXml++; return INFO; },
      async noiseInfoFromBuf() { calls.fromBuf++; return INFO; },
      // 老接口：一旦被调用就说明绕开失败 → 计数即可，同时抛错以暴露问题
      async noiseNonNoiseFromBuf() { calls.oldPairs++; throw new Error("不应调用 noiseNonNoiseFromBuf"); },
      async noiseNonNoiseTextFromBuf() { calls.oldText++; throw new Error("不应调用 noiseNonNoiseTextFromBuf"); },
      async noiseWorkshopGroupsFromBuf() { calls.oldWorkshop++; throw new Error("不应调用 noiseWorkshopGroupsFromBuf"); },
    };
    const blf = makeSandbox(stub);
    const out = await blf.patchPjbgDocx(buf);
    const xml = await readXml(out);
    const text = allText(xml);

    check("getCachedNoiseInfo 被询问 1 次", calls.cached === 1, "实际 " + calls.cached);
    check("noiseInfoFromXml 只调用 1 次（不是 3 次）", calls.fromXml === 1, "实际 " + calls.fromXml);
    check("noiseInfoFromBuf 未被调用（已有 xml，不必再解压）", calls.fromBuf === 0, "实际 " + calls.fromBuf);
    check("老接口 noiseNonNoiseFromBuf 未被调用", calls.oldPairs === 0, "实际 " + calls.oldPairs);
    check("老接口 noiseNonNoiseTextFromBuf 未被调用", calls.oldText === 0, "实际 " + calls.oldText);
    check("老接口 noiseWorkshopGroupsFromBuf 未被调用", calls.oldWorkshop === 0, "实际 " + calls.oldWorkshop);

    const haz = hazardColTexts(blf, xml);
    check("② 打磨岗位（非噪声）的「噪声」已移除", haz[0] && haz[0].indexOf('噪声') < 0, JSON.stringify(haz));
    check("② 打磨岗位其它危害因素保留（电焊烟尘）", haz[0] && haz[0].indexOf('电焊烟尘') >= 0, JSON.stringify(haz));
    check("② 装配岗位（噪声）的「噪声」保留", haz[1] && haz[1].indexOf('噪声') >= 0, JSON.stringify(haz));
    check("⑥ 噪声车间岗位为空 → 删除 GB 23466-2025 条目", text.indexOf("GB 23466-2025") < 0);
    check("⑥ 未误删其它评价依据条目", text.indexOf("《中华人民共和国职业病防治法》") >= 0);
  }

  // ================= B. 缓存命中：零解析 =================
  console.log("\n[B] 缓存命中（getCachedNoiseInfo 返回数据）：不再解析");
  {
    const calls = { fromXml: 0, fromBuf: 0, cached: 0 };
    const stub = {
      getCachedNoiseInfo() { calls.cached++; return INFO; },
      noiseInfoFromXml() { calls.fromXml++; return INFO; },
      async noiseInfoFromBuf() { calls.fromBuf++; return INFO; },
    };
    const blf = makeSandbox(stub);
    const out = await blf.patchPjbgDocx(buf);
    const xml = await readXml(out);
    check("缓存命中时不再调用 noiseInfoFromXml", calls.fromXml === 0, "实际 " + calls.fromXml);
    check("缓存命中时不再调用 noiseInfoFromBuf", calls.fromBuf === 0, "实际 " + calls.fromBuf);
    const hazB = hazardColTexts(blf, xml);
    check("缓存命中仍正确删除「噪声」", hazB[0] && hazB[0].indexOf('噪声') < 0, JSON.stringify(hazB));
  }

  // ================= C. 只有老接口：向后兼容 =================
  console.log("\n[C] 只有老接口（页面缓存了旧版 upload.js）：退化可用");
  {
    const calls = { oldPairs: 0, oldText: 0, oldWorkshop: 0 };
    const stub = {
      async noiseNonNoiseFromBuf() { calls.oldPairs++; return [{ unit: "M车间", job: "打磨岗位" }]; },
      async noiseNonNoiseTextFromBuf() { calls.oldText++; return "M车间打磨岗位"; },
      async noiseWorkshopGroupsFromBuf() { calls.oldWorkshop++; return []; },
    };
    const blf = makeSandbox(stub);
    const out = await blf.patchPjbgDocx(buf);
    const xml = await readXml(out);
    const text = allText(xml);
    check("退化路径：调用老接口取 pairs", calls.oldPairs === 1, "实际 " + calls.oldPairs);
    check("退化路径：调用老接口取文本", calls.oldText === 1, "实际 " + calls.oldText);
    check("退化路径：调用老接口取噪声车间分组", calls.oldWorkshop === 1, "实际 " + calls.oldWorkshop);
    const hazC = hazardColTexts(blf, xml);
    check("退化路径：② 仍删掉「噪声」", hazC[0] && hazC[0].indexOf('噪声') < 0, JSON.stringify(hazC));
    check("退化路径：⑥ 仍删 GB 23466-2025", text.indexOf("GB 23466-2025") < 0);
  }

  // ================= D. 完全没有 SamplingUpload：不删噪声类条目 =================
  console.log("\n[D] 无 SamplingUpload（Node 沙箱形态）：噪声类条目不删");
  {
    const blf = makeSandbox(null);
    const out = await blf.patchPjbgDocx(buf);
    const xml = await readXml(out);
    check("无上传模块时不误删 GB 23466-2025", allText(xml).indexOf("GB 23466-2025") >= 0);
  }

  // ================= E. ok=false（解析失败）：不删「噪声」但保持既有语义 =================
  console.log("\n[E] 解析失败（ok=false）：② 不删噪声，⑥ 仍按「为空」处理（与旧实现一致）");
  {
    const BAD = { rows: [], ok: false, nonNoiseGroups: [], nonNoisePairs: [], nonNoiseText: "", noiseWorkshopGroups: [] };
    const stub = {
      getCachedNoiseInfo() { return null; },
      noiseInfoFromXml() { return BAD; },
    };
    const blf = makeSandbox(stub);
    const out = await blf.patchPjbgDocx(buf);
    const xml = await readXml(out);
    const hazE = hazardColTexts(blf, xml);
    check("解析失败：确认表里的「噪声」保留", hazE[0] && hazE[0].indexOf('噪声') >= 0, JSON.stringify(hazE));
    check("解析失败：⑥ 仍删噪声类条目（旧语义：空 = 删）", allText(xml).indexOf("GB 23466-2025") < 0);
  }

  console.log("\n结果：" + pass + " 通过 / " + fail + " 失败");
  if (fail) process.exitCode = 1;
})();
