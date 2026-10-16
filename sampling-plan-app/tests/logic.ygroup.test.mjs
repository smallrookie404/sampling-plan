// 岗位性质（Y）按「车间+岗位」组填充的回归测试。
// 需求：录入区岗位性质（O）有内容时，计算区同车间、同岗位的所有行的 Y 均改为该内容
// （参照「工作内容 J→AM」的分组填充逻辑，但分组键不含点位）。
// 纯 Node，无浏览器依赖。
import { createRequire } from "node:module";
import path from "node:path";
import { fileURLToPath } from "node:url";

const require = createRequire(import.meta.url);
const APP_DIR = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "..");
const L = require(path.join(APP_DIR, "js", "logic.js"));

let pass = 0, fail = 0;
function ok(name, cond, extra) {
  if (cond) { pass++; console.log("  ok  " + name + (extra ? "  → " + extra : "")); }
  else { fail++; console.log("  FAIL " + name + (extra ? "  → " + extra : "")); }
}
const mkRow = (input, overridden) => ({
  input, manual: {}, overridden: overridden || {}, values: {}, errors: {},
});
const compute = (rows) => L.computeRows(rows, { hazardFactors: [] }).map((r) => r.values.Y);
const eq = (a, b) => JSON.stringify(a) === JSON.stringify(b);

console.log("[1] 基础分组填充：同车间+同岗位跨点位，O 只在某行有值 → 整组统一");
{
  const rows = [
    mkRow({ A: "车间1", B: "岗位P", C: "点位1", O: "" }),
    mkRow({ A: "车间1", B: "岗位P", C: "点位2", O: "流动" }),
    mkRow({ A: "车间1", B: "岗位P", C: "点位3", O: "" }),
  ];
  ok("整组变为「流动」", eq(compute(rows), ["流动", "流动", "流动"]), JSON.stringify(compute(rows)));
}

console.log("[2] 分组边界：不同车间 / 不同岗位互不影响");
{
  const rows = [
    mkRow({ A: "车间1", B: "岗位P", C: "点位1", O: "流动" }),
    mkRow({ A: "车间2", B: "岗位P", C: "点位1", O: "" }), // 不同车间 → 仍固定
    mkRow({ A: "车间1", B: "岗位Q", C: "点位1", O: "" }), // 不同岗位 → 仍固定
  ];
  const got = compute(rows);
  ok("同车间同岗位才填充", eq(got, ["流动", "固定", "固定"]), JSON.stringify(got));
}

console.log("[3] 组内多个非空值时取第一个非空（组内最先出现）");
{
  const rows = [
    mkRow({ A: "车间1", B: "岗位Q", C: "点位1", O: "巡检" }),
    mkRow({ A: "车间1", B: "岗位Q", C: "点位2", O: "固定" }),
  ];
  const got = compute(rows);
  ok("取第一个非空「巡检」", eq(got, ["巡检", "巡检"]), JSON.stringify(got));
}

console.log("[4] 手工覆盖（overridden.Y）优先于分组填充");
{
  const rows = [
    mkRow({ A: "车间1", B: "岗位P", C: "点位1", O: "" }),
    mkRow({ A: "车间1", B: "岗位P", C: "点位2", O: "流动" }),
  ];
  rows[0].overridden = { Y: true };
  rows[0].values.Y = "巡检";
  const got = compute(rows);
  ok("覆盖行保持「巡检」", got[0] === "巡检", got[0]);
  ok("未覆盖行仍按组填「流动」", got[1] === "流动", got[1]);
}

console.log("[5] 整组 O 全空 → 默认「固定」");
{
  const rows = [mkRow({ A: "车间Z", B: "岗位Z", C: "点位1", O: "" })];
  ok("默认固定", compute(rows)[0] === "固定", compute(rows)[0]);
}

console.log("[6] 单行组（无同组其他行）行为不变");
{
  const rows = [mkRow({ A: "车间1", B: "岗位P", C: "点位1", O: "巡检" })];
  ok("单行 O=巡检 → 巡检", compute(rows)[0] === "巡检", compute(rows)[0]);
}

console.log("\n结果：" + pass + " 通过 / " + fail + " 失败");
process.exit(fail ? 1 : 0);
