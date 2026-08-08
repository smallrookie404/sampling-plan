/*
 * 计算引擎：与《0.系统测点布局调查》Excel 公式逐列对应的移植版本。
 * 纯函数、无 DOM 依赖，可在浏览器与 Node 中复用。
 */
(function (root, factory) {
  if (typeof module === "object" && module.exports) module.exports = factory();
  else root.SamplingLogic = factory();
})(typeof self !== "undefined" ? self : this, function () {
  "use strict";

  // ---------- 常量（与 Excel 数据验证一致） ----------
  const BANZHI_R = ["单班", "轮班"];
  const BANZHI_S = [
    "两班两运转",
    "三班两运转",
    "三班三运转",
    "三班四运转",
    "四班三运转",
    "五班三运转",
    "五班四运转",
    "长夜班",
    "其他",
  ];
  // 自动计算列 AE（岗位班制数）额外允许“长白班”（Excel 条件格式规则）
  const BANZHI_AE = ["长白班", ...BANZHI_S];
  const GANGWEI_XZ = ["固定", "流动", "巡检"];
  const JIECHU_LX = ["浓度（或强度）相对不稳定", "浓度（或强度）相对稳定"];
  const SHI_FOU = ["是", "否"];
  const JIANCE_FS = ["定点", "个体"];
  const TILI_LD = ["Ⅰ", "Ⅱ", "Ⅲ", "Ⅳ"];

  // 列定义：A..BI
  const INPUT_COLS = "ABCDEFGHIJKLMNOPQRSTU".split("");
  const COMPUTED_COLS = [
    "W", "X", "Y", "Z", "AA", "AB", "AC", "AD", "AE", "AF", "AG", "AH", "AI",
    "AJ", "AK", "AL", "AM", "AN", "AO", "AP", "AQ", "AR", "AS", "AT", "AU",
    "AV", "AW", "AX", "AY", "AZ", "BA", "BB", "BC", "BD", "BE", "BF", "BG",
    "BH", "BI",
  ];

  // 自动计算区内仍由用户手工填写的列
  const MANUAL_COLS = ["AI", "AJ", "AX", "BE", "BF", "BG", "BH"];
  // 表中有下拉、可覆盖自动值的列
  const OVERRIDE_COLS = ["Y", "Z", "AO", "AR"];

  // ---------- 工具函数 ----------
  function toNum(v) {
    if (v === null || v === undefined || v === "") return "";
    const n = typeof v === "number" ? v : parseFloat(String(v));
    return isNaN(n) ? "" : n;
  }

  function ceilDiv(a, b) {
    const n = toNum(a);
    const d = toNum(b) || 1;
    if (n === "") return "";
    return Math.ceil(n / d);
  }

  // 向前填充：空白继承上一个非空值（对应 Excel SCAN+LAMBDA 的前向填充）
  function fillDown(arr) {
    let prev = "";
    return arr.map((v) => {
      if (v !== null && v !== undefined && v !== "") {
        prev = v;
        return v;
      }
      return prev;
    });
  }

  function str(v) {
    return v === null || v === undefined ? "" : String(v);
  }

  function contains(needle, hay) {
    return str(hay).includes(str(needle));
  }

  // 危害因素表查找：按“识别”列精确匹配（对应 VLOOKUP(接害因素, A:B, 2, FALSE)）
  function findHazardByRec(ref, rec) {
    const key = str(rec);
    if (!key) return null;
    return ref.hazardFactors.find((h) => str(h.rec) === key) || null;
  }

  // 危害因素表查找：按“系统名称”列精确匹配（对应 XLOOKUP/VLOOKUP(AN, B:L, ...)）
  function findHazardByName(ref, name) {
    const key = str(name);
    if (!key) return null;
    return ref.hazardFactors.find((h) => str(h.name) === key) || null;
  }

  // ---------- 计算入口 ----------
  /*
   * rows: 行数组，每行结构
   *   {
   *     input: { A..U },           // 用户录入区
   *     manual: { AI,AJ,AX,BE,BF,BG,BH }, // 手工填写列
   *     overridden: { Y:true, Z:true, AO:true, AR:true }, // 用户覆盖过的列
   *     values: { W..BI },         // 引擎计算结果（含手工列合并后的最终值）
   *     errors: { col: 说明 }       // 校验错误
   *   }
   */
  function computeRows(rows, ref) {
    const n = rows.length;
    const last = n - 1;

    // 提取各输入列
    const col = (c) => rows.map((r, i) => (i <= last ? r.input[c] ?? "" : ""));
    const A = col("A"), B = col("B"), C = col("C"), D = col("D"), E = col("E");
    const F = col("F"), G = col("G"), H = col("H"), I = col("I"), L = col("L");
    const M = col("M"), N = col("N"), O = col("O"), P = col("P"), Q = col("Q");
    const R = col("R"), S = col("S"), T = col("T"), U = col("U");

    // 1) 基础下填/默认列（对应 W/X/Y/Z/AA/AC/AG/AH/AY/AZ/AL/AQ/AR/BC）
    const W = fillDown(A);
    const X = fillDown(B);
    const Ybase = O.map((v) => (v !== "" ? v : "固定"));
    const Zbase = A.map(() => "浓度（或强度）相对稳定");
    const AA = fillDown(H).map(toNum);
    const AC = L.map((v) => (v !== "" ? v : "半手工作业"));
    const AG = fillDown(G);
    const AH = fillDown(F).map(toNum);
    const AL = fillDown(C);
    const AQ = Q.map((v) => (v !== "" ? v : "否"));
    const ARbase = N.map((v) => (v !== "" ? v : "定点"));
    const AY = fillDown(G);
    const AZ = fillDown(E).map(toNum);
    const BC = P.map((v) => (v !== "" ? v : "否"));
    const I_filled = fillDown(I); // BA 使用的“检测天数”下填值

    // 2) 组级计算（按 车间|岗位 分组）
    // AB 每班最大人数
    const baseDivisor = S.map((s, i) => {
      const check = contains("其他", s) ? T[i] : s;
      const ck = str(check);
      if (contains("五班", ck)) return 5;
      if (contains("四班", ck)) return 4;
      if (contains("三班", ck)) return 3;
      if (contains("两班", ck)) return 2;
      return 1;
    });
    const groupDivisor = {};
    const groupId = W.map((w, i) => str(w) + "|" + str(X[i]));
    for (let i = 0; i < n; i++) {
      const g = groupId[i];
      if (!(g in groupDivisor) || baseDivisor[i] > groupDivisor[g]) groupDivisor[g] = baseDivisor[i];
    }
    const AB = AA.map((v, i) => {
      const d = groupDivisor[groupId[i]] ?? 1;
      return v === "" ? "" : ceilDiv(v, d);
    });

    // AD/AE/AF 班制：组内回填，无则单班/长白班
    const AD = R.map((v, i) => {
      if (v !== "") return v;
      const g = groupId[i];
      const other = R.some((rv, j) => j !== i && groupId[j] === g && contains("轮班", rv));
      return other ? "轮班" : "单班";
    });
    const AE = S.map((v, i) => {
      if (v !== "") return v;
      const g = groupId[i];
      const hit = S.find((sv, j) => groupId[j] === g && sv !== "");
      return hit === undefined || hit === "" ? "长白班" : hit;
    });
    const AF = T.map((v, i) => {
      if (v !== "") return v;
      const g = groupId[i];
      const hit = T.find((tv, j) => groupId[j] === g && tv !== "");
      return hit ?? "";
    });

    // AV 噪声源：噪声+定点 时，同车间同岗位存在“个体”噪声采样 → 是
    const AR = rows.map((r, i) => (r.overridden && r.overridden.AR ? r.values.AR : ARbase[i]));
    const ANpre = D.map((d) => {
      const h = findHazardByRec(ref, d);
      return h ? h.name : "";
    });
    const AV = ANpre.map((an, i) => {
      if (an !== "噪声" || AR[i] !== "定点") return "否";
      const g = groupId[i];
      return ANpre.some((an2, j) => groupId[j] === g && an2 === "噪声" && AR[j] === "个体") ? "是" : "否";
    });

    // 3) 危害因素表查找列
    const AN = ANpre;
    const AP = D.map((d) => {
      const h = findHazardByRec(ref, d);
      return h ? h.qual : "";
    });
    const AT = D.map((d) => {
      const h = findHazardByRec(ref, d);
      return h ? h.outsource : "";
    });
    const AU = D.map((d) => {
      const h = findHazardByRec(ref, d);
      return h ? h.dust : "";
    });
    const BI = AN.map((an) => {
      const h = findHazardByName(ref, an);
      return h ? h.noTestReason : "";
    });
    const AObase = M.map((v, i) => {
      if (v !== "") return v;
      const h = findHazardByName(ref, AN[i]);
      return h && h.noTestReason !== "" ? "否" : "是";
    });
    const AO = rows.map((r, i) => (r.overridden && r.overridden.AO ? r.values.AO : AObase[i]));

    // 4) 派生列
    const AK = AR.map((v) => (v === "定点" ? "采样点" : "采样对象"));
    const AM = AL.map((s) => {
      const text = str(s);
      const posQ = text.indexOf("区");
      const digits = [0, 1, 2, 3, 4, 5, 6, 7, 8, 9].map((d) => text.indexOf(String(d))).filter((p) => p >= 0);
      const earliest = [posQ, ...digits].filter((p) => p >= 0).sort((a, b) => a - b)[0];
      const base = earliest === undefined || earliest === text.length ? text : text.slice(0, earliest);
      return "操作设备" + base + "作业";
    });
    const AS = AR.map((v) => (v === "定点" ? "短时间" : v === "个体" ? "长时间" : ""));
    const AW = AN.map((an, i) => (an === "高温" ? (U[i] !== "" ? U[i] : "Ⅱ") : ""));
    const BA = AN.map((an, i) => {
      const special =
        ["游离二氧化硅", "噪声", "高温", "工频电场", "手传振动", "高频电磁场", "紫外辐射"].includes(an) ||
        contains("有机组分定性", an);
      return special ? 1 : I_filled[i];
    });
    const BB = AN.map((an, i) => {
      if (contains("有机组分定性", an) || AR[i] === "个体" || ["游离二氧化硅", "工频电场", "高频电磁场"].includes(an)) return 1;
      if (["噪声", "高温", "手传振动"].includes(an)) return 3;
      if (an === "紫外辐射") return 1;
      const az = AZ[i];
      if (az === 0.5) return 2;
      if (az === 0.25) return 1;
      return 3;
    });
    const BD = AN.map((an) => (["噪声", "高温", "紫外辐射"].includes(an) ? "设备运行" : "原辅物料"));

    // 5) 组装结果
    for (let i = 0; i < n; i++) {
      const r = rows[i];
      const vals = r.values;
      // 任意自动列均可被覆盖（导入“仅自动计算区”文件时用于保留静态结果）
      const keep = (c) => r.overridden && r.overridden[c] === true;
      const put = (c, v) => {
        if (!keep(c)) vals[c] = v;
      };
      put("W", W[i]);
      put("X", X[i]);
      put("Y", Ybase[i]);
      put("Z", Zbase[i]);
      put("AA", AA[i]);
      put("AB", AB[i]);
      put("AC", AC[i]);
      put("AD", AD[i]);
      put("AE", AE[i]);
      put("AF", AF[i]);
      put("AG", AG[i]);
      put("AH", AH[i]);
      put("AK", AK[i]);
      put("AL", AL[i]);
      put("AM", AM[i]);
      put("AN", AN[i]);
      put("AO", AO[i]);
      put("AP", AP[i]);
      put("AQ", AQ[i]);
      put("AR", AR[i]);
      put("AS", AS[i]);
      put("AT", AT[i]);
      put("AU", AU[i]);
      put("AV", AV[i]);
      put("AW", AW[i]);
      put("AY", AY[i]);
      put("AZ", AZ[i]);
      put("BA", BA[i]);
      put("BB", BB[i]);
      put("BC", BC[i]);
      put("BD", BD[i]);
      put("BI", BI[i]);
      // 手工列
      for (const mc of MANUAL_COLS) vals[mc] = r.manual[mc] ?? "";
      // 校验
      r.errors = validateRow(r, ref);
    }
    return rows;
  }

  // ---------- 校验（对应 Excel 条件格式规则） ----------
  function validateRow(r, ref) {
    const errs = {};
    const input = r.input, vals = r.values;
    const check = (col, val, allowed, label) => {
      if (val !== "" && !allowed.includes(str(val))) errs[col] = `${label}“${val}”无效，应为：${allowed.join(" / ")}`;
    };
    check("R", input.R, BANZHI_R, "岗位工作班制");
    check("S", input.S, BANZHI_S, "岗位班制数");
    check("U", input.U, TILI_LD, "体力劳动强度");
    check("Y", vals.Y, GANGWEI_XZ, "岗位性质");
    check("Z", vals.Z, JIECHU_LX, "接触类型");
    check("AO", vals.AO, SHI_FOU, "是否采样/测量");
    check("AR", vals.AR, JIANCE_FS, "检测方式");
    check("AD", vals.AD, BANZHI_R, "岗位工作班制（自动）");
    check("AE", vals.AE, BANZHI_AE, "岗位班制数（自动）");
    if (input.D !== "" && vals.AN === "") errs.AN = "接害因素在危害因素库中未找到，请检查名称或补充库表";
    return errs;
  }

  function countErrors(rows) {
    let total = 0;
    const byCol = {};
    for (const r of rows) {
      for (const k of Object.keys(r.errors || {})) {
        total++;
        byCol[k] = (byCol[k] || 0) + 1;
      }
    }
    return { total, byCol };
  }

  // 行状态快照：仅保留可序列化字段，用于保存到本机数据库
  function snapshotRows(rows) {
    return rows.map((r) => ({
      input: { ...(r.input || {}) },
      manual: { ...(r.manual || {}) },
      overridden: { ...(r.overridden || {}) },
      values: { ...(r.values || {}) },
    }));
  }

  // 从快照恢复行对象（深层复制，避免与保存时的引用共享）
  function restoreRows(snap) {
    return snap.map((s) => {
      const input = {};
      for (const c of INPUT_COLS) input[c] = s.input?.[c] ?? "";
      const manual = {};
      for (const c of MANUAL_COLS) manual[c] = s.manual?.[c] ?? "";
      return {
        input,
        manual,
        overridden: { ...(s.overridden || {}) },
        values: { ...(s.values || {}) },
        errors: {},
      };
    });
  }

  // 统计参考库中检测项目重复项
  function findDuplicates(items) {
    const seen = new Map();
    for (const it of items) {
      const k = str(it);
      seen.set(k, (seen.get(k) || 0) + 1);
    }
    return [...seen.entries()].filter(([, c]) => c > 1).map(([k]) => k);
  }

  return {
    BANZHI_R,
    BANZHI_S,
    BANZHI_AE,
    GANGWEI_XZ,
    JIECHU_LX,
    SHI_FOU,
    JIANCE_FS,
    TILI_LD,
    INPUT_COLS,
    COMPUTED_COLS,
    MANUAL_COLS,
    OVERRIDE_COLS,
    fillDown,
    toNum,
    computeRows,
    validateRow,
    countErrors,
    snapshotRows,
    restoreRows,
    findDuplicates,
    findHazardByRec,
    findHazardByName,
  };
});
