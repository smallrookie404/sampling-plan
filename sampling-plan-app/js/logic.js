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
    "长白班",
    "其他",
  ];
  // 自动计算列 AE（岗位班制数）允许值与录入列 S 一致（含「长白班」）
  const BANZHI_AE = [...BANZHI_S];
  const GANGWEI_XZ = ["固定", "流动", "巡检"];
  const JIECHU_LX = ["浓度（或强度）相对不稳定", "浓度（或强度）相对稳定"];
  const SHI_FOU = ["是", "否"];
  const JIANCE_FS = ["定点", "个体"];
  const TILI_LD = ["Ⅰ", "Ⅱ", "Ⅲ", "Ⅳ"];

  // 列定义：A..BI
  // A..U 原录入列，V 为新增录入列「备注」（显示在“是否采样”后面）
  const INPUT_COLS = "ABCDEFGHIJKLMNOPQRSTUV".split("");
  const COMPUTED_COLS = [
    "W", "X", "Y", "Z", "AA", "AB", "AC", "AD", "AE", "AF", "AG", "AH", "AI",
    "AJ", "AK", "AL", "AM", "AN", "AO", "AP", "AQ", "AR", "AS", "AT", "AU",
    "AV", "AW", "AX", "AY", "AZ", "BA", "BB", "BC", "BD", "BE", "BF", "BG",
    "BH", "BI",
  ];

  // 自动计算区内仍由用户手工填写的列
  const MANUAL_COLS = ["AI", "AJ", "AX", "BF", "BG", "BH"];
  // 表中有下拉、可覆盖自动值的列
  const OVERRIDE_COLS = ["Y", "Z", "AO", "AR"];
  // 自动计算区内默认自动生成、但允许手动编辑覆盖的列（编辑后保留手动值）
  const TEXT_OVERRIDE_COLS = ["BI"];

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

  // Unicode 安全的 base64（浏览器与 Node 通用），用于 GitHub API 读写 JSON
  function encodeUnicodeBase64(str) {
    const bytes = new TextEncoder().encode(String(str));
    let bin = "";
    for (let i = 0; i < bytes.length; i++) bin += String.fromCharCode(bytes[i]);
    return btoa(bin);
  }

  function decodeUnicodeBase64(b64) {
    const bin = atob(String(b64));
    const bytes = new Uint8Array(bin.length);
    for (let i = 0; i < bin.length; i++) bytes[i] = bin.charCodeAt(i);
    return new TextDecoder().decode(bytes);
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
   *     manual: { AI,AJ,AX,BF,BG,BH }, // 手工填写列
   *     overridden: { Y:true, Z:true, AO:true, AR:true }, // 用户覆盖过的列
   *     values: { W..BI },         // 引擎计算结果（含手工列合并后的最终值）
   *     errors: { col: 说明 }       // 校验错误
   *   }
   */
  // 单行「操作设备…作业」文本（AM 的默认推导）：取「区」或首个数字之前的部分
  function opBaseText(text) {
    const t = str(text);
    let earliest = -1;
    const q = t.indexOf("区");
    if (q >= 0) earliest = q;
    for (let d = 0; d < 10; d++) {
      const p = t.indexOf(String(d));
      if (p >= 0 && (earliest < 0 || p < earliest)) earliest = p;
    }
    const base = earliest < 0 || earliest === t.length ? t : t.slice(0, earliest);
    return "操作设备" + base + "作业";
  }

  /*
   * 计算入口。这是全表重算：行数上万时每次编辑都会调用，
   * 因此刻意不用「每列一个 map 数组」的写法（旧实现每次重算要分配 40 多个 n 长数组、
   * 约 7n 次分组键字符串拼接），改为两遍循环：
   *   ① 下填列 + 「车间|岗位(|点位)」分组聚合（一遍扫描，结果落在 Map/Set 里）；
   *   ② 逐行组装 values 并校验。
   * 语义与旧版逐列实现完全一致（逐列比对见 tests/_diff_logic.mjs）。
   */
  function computeRows(rows, ref) {
    const n = rows.length;
    if (!n) return rows;

    // 危害因素查找索引：一次构建，逐行 O(1) 查询（语义与原 find 一致：保留首个匹配）
    const hazardRecMap = new Map();
    const hazardNameMap = new Map();
    for (const h of ref.hazardFactors) {
      const rec = str(h.rec);
      if (rec && !hazardRecMap.has(rec)) hazardRecMap.set(rec, h);
      const name = str(h.name);
      if (name && !hazardNameMap.has(name)) hazardNameMap.set(name, h);
    }
    const findRec = (d) => {
      const k = str(d);
      return k ? hazardRecMap.get(k) || null : null;
    };
    const findName = (nm) => {
      const k = str(nm);
      return k ? hazardNameMap.get(k) || null : null;
    };

    // ---------- 第一遍：下填列 + 分组聚合 ----------
    // 下填结果按列存数组（第二遍组装要用）：W=A X=B AL=C AG=G AH=F AZ=E AA=H BA=I
    const W = new Array(n), X = new Array(n), AL = new Array(n);
    const AG = new Array(n), AH = new Array(n), AZ = new Array(n);
    const AA = new Array(n), Ifilled = new Array(n);
    const keyWH = new Array(n);   // 车间|岗位
    const keyWHP = new Array(n);  // 车间|岗位|点位

    const groupDivisor = new Map();     // AB：车间|岗位 → 最大班数
    const groupHasRounds = new Set();   // AD：组内是否有「轮班」
    const groupSHit = new Map();        // AE：组内首个非空 S
    const groupTHit = new Map();        // AF：组内首个非空 T
    const groupJ = new Map();           // AM：组内首个非空 J
    const groupO = new Map();           // Y：组内首个非空 O
    const groupL = new Map();           // AC：组内首个非空 L
    const groupK = new Map();           // BE/BD：组内首个非空 K
    const groupHasNoiseInd = new Set(); // AV：组内存在「噪声 + 个体」

    let prevA = "", prevB = "", prevC = "", prevE = "", prevF = "", prevG = "", prevH = "", prevI = "";
    for (let i = 0; i < n; i++) {
      const r = rows[i];
      const inp = r.input;

      const a = inp.A ?? ""; if (a !== "") prevA = a;
      const b = inp.B ?? ""; if (b !== "") prevB = b;
      const c = inp.C ?? ""; if (c !== "") prevC = c;
      const e = inp.E ?? ""; if (e !== "") prevE = e;
      const f = inp.F ?? ""; if (f !== "") prevF = f;
      const g = inp.G ?? ""; if (g !== "") prevG = g;
      const hh = inp.H ?? ""; if (hh !== "") prevH = hh;
      const ii = inp.I ?? ""; if (ii !== "") prevI = ii;

      W[i] = prevA; X[i] = prevB; AL[i] = prevC;
      AG[i] = prevG;
      AH[i] = toNum(prevF);
      AZ[i] = toNum(prevE);
      AA[i] = toNum(prevH);
      Ifilled[i] = prevI;

      const k2 = prevA + "\u0001" + prevB;
      const k3 = k2 + "\u0001" + prevC;
      keyWH[i] = k2;
      keyWHP[i] = k3;

      // AB 的班数除数：组内取最大值
      const s = inp.S ?? "";
      const check = contains("其他", s) ? (inp.T ?? "") : s;
      let div = 1;
      if (contains("五班", check)) div = 5;
      else if (contains("四班", check)) div = 4;
      else if (contains("三班", check)) div = 3;
      else if (contains("两班", check)) div = 2;
      const gd = groupDivisor.get(k2);
      if (gd === undefined || div > gd) groupDivisor.set(k2, div);

      if (contains("轮班", inp.R)) groupHasRounds.add(k2);
      if (s !== "" && !groupSHit.has(k2)) groupSHit.set(k2, s);
      const t = inp.T ?? "";
      if (t !== "" && !groupTHit.has(k2)) groupTHit.set(k2, t);
      const j = inp.J ?? "";
      if (j !== "" && !groupJ.has(k3)) groupJ.set(k3, j);
      const o = inp.O ?? "";
      if (o !== "" && !groupO.has(k2)) groupO.set(k2, o);
      const l = inp.L ?? "";
      if (l !== "" && !groupL.has(k3)) groupL.set(k3, l);
      const k = inp.K ?? "";
      if (k !== "" && !groupK.has(k3)) groupK.set(k3, k);

      // AV 预扫描：本组是否已有「噪声 + 个体」
      const ov = r.overridden;
      const rec0 = findRec(inp.D);
      const an0 = rec0 ? rec0.name : "";
      const ar0 = ov && ov.AR ? r.values.AR : (inp.N !== "" ? inp.N : "定点");
      if (an0 === "噪声" && ar0 === "个体") groupHasNoiseInd.add(k2);
    }

    // ---------- 第二遍：组装结果 ----------
    for (let i = 0; i < n; i++) {
      const r = rows[i];
      const inp = r.input;
      const vals = r.values;
      const ov = r.overridden;
      const k2 = keyWH[i];
      const k3 = keyWHP[i];
      // 录入区取值：缺键统一按空串（与旧实现的 r.input[c] ?? "" 一致）
      const vR = inp.R ?? "", vS = inp.S ?? "", vT = inp.T ?? "", vO = inp.O ?? "";
      const vL = inp.L ?? "", vM = inp.M ?? "", vN = inp.N ?? "", vP = inp.P ?? "";
      const vQ = inp.Q ?? "", vU = inp.U ?? "";

      const rec = findRec(inp.D);
      const an = rec ? rec.name : "";
      const anName = findName(an);
      const ar = ov && ov.AR ? vals.AR : vN !== "" ? vN : "定点";

      // 1) 基础下填/默认列
      if (ov == null || ov.W !== true) vals.W = W[i];
      if (ov == null || ov.X !== true) vals.X = X[i];
      if (ov == null || ov.Y !== true) {
        if (groupO.has(k2)) vals.Y = groupO.get(k2);
        else vals.Y = vO !== "" ? vO : "固定";
      }
      if (ov == null || ov.Z !== true) vals.Z = "浓度（或强度）相对稳定";
      if (ov == null || ov.AA !== true) vals.AA = AA[i];
      if (ov == null || ov.AB !== true) {
        const d = groupDivisor.get(k2) ?? 1;
        vals.AB = AA[i] === "" ? "" : ceilDiv(AA[i], d);
      }
      if (ov == null || ov.AC !== true) {
        if (groupL.has(k3)) vals.AC = groupL.get(k3);
        else vals.AC = vL !== "" ? vL : "半手工作业";
      }
      if (ov == null || ov.AD !== true) vals.AD = vR !== "" ? vR : groupHasRounds.has(k2) ? "轮班" : "单班";
      if (ov == null || ov.AE !== true) vals.AE = vS !== "" ? vS : groupSHit.has(k2) ? groupSHit.get(k2) : "长白班";
      if (ov == null || ov.AF !== true) vals.AF = vT !== "" ? vT : (groupTHit.get(k2) ?? "");
      if (ov == null || ov.AG !== true) vals.AG = AG[i];
      if (ov == null || ov.AH !== true) vals.AH = AH[i];

      // 2) 危害因素查表列
      if (ov == null || ov.AK !== true) vals.AK = ar === "定点" ? "采样点" : "采样对象";
      if (ov == null || ov.AL !== true) vals.AL = AL[i];
      if (ov == null || ov.AM !== true) vals.AM = groupJ.has(k3) ? groupJ.get(k3) : opBaseText(AL[i]);
      if (ov == null || ov.AN !== true) vals.AN = an;
      if (ov == null || ov.AO !== true) {
        if (vM !== "") vals.AO = vM;
        else vals.AO = anName && anName.noTestReason !== "" ? "否" : "是";
      }
      if (ov == null || ov.AP !== true) vals.AP = rec ? rec.qual : "";
      if (ov == null || ov.AQ !== true) vals.AQ = vQ !== "" ? vQ : "否";
      if (ov == null || ov.AR !== true) vals.AR = ar;
      if (ov == null || ov.AS !== true) vals.AS = ar === "定点" ? "短时间" : ar === "个体" ? "长时间" : "";
      if (ov == null || ov.AT !== true) vals.AT = rec ? rec.outsource : "";
      if (ov == null || ov.AU !== true) vals.AU = rec ? rec.dust : "";
      if (ov == null || ov.AV !== true) {
        if (an !== "噪声" || ar !== "定点") vals.AV = "否";
        else vals.AV = groupHasNoiseInd.has(k2) ? "是" : "否";
      }
      if (ov == null || ov.AW !== true) vals.AW = an === "高温" ? (vU !== "" ? vU : "Ⅱ") : "";
      if (ov == null || ov.AY !== true) vals.AY = AG[i];
      if (ov == null || ov.AZ !== true) vals.AZ = AZ[i];

      // 3) 派生列
      if (ov == null || ov.BA !== true) {
        const special =
          an === "游离二氧化硅" || an === "噪声" || an === "高温" || an === "工频电场" ||
          an === "手传振动" || an === "高频电磁场" || an === "紫外辐射" || contains("有机组分定性", an);
        vals.BA = special ? 1 : Ifilled[i];
      }
      if (ov == null || ov.BB !== true) {
        if (contains("有机组分定性", an) || ar === "个体") vals.BB = 1;
        else if (an === "游离二氧化硅" || an === "工频电场" || an === "高频电磁场" || an === "激光辐射") vals.BB = 1;
        else if (an === "噪声" || an === "高温" || an === "手传振动") vals.BB = 3;
        else if (an === "紫外辐射") vals.BB = 1;
        else if (AZ[i] === 0.5) vals.BB = 2;
        else if (AZ[i] === 0.25) vals.BB = 1;
        else vals.BB = 3;
      }
      if (ov == null || ov.BC !== true) vals.BC = vP !== "" ? vP : "否";
      if (ov == null || ov.BD !== true) {
        if (groupK.has(k3)) vals.BD = "";
        else vals.BD = an === "噪声" || an === "高温" || an === "紫外辐射" ? "设备运行" : "原辅物料";
      }
      if (ov == null || ov.BE !== true) vals.BE = groupK.get(k3) ?? "";
      if (ov == null || ov.BI !== true) vals.BI = anName ? anName.noTestReason : "";

      // 4) 手工列
      for (let mi = 0; mi < MANUAL_COLS.length; mi++) {
        const mc = MANUAL_COLS[mi];
        vals[mc] = r.manual[mc] ?? "";
      }
      // 录入区备注列（V）有内容时，自动计算区备注列（BI）填充该内容；为空则按原逻辑
      const remark = inp.V;
      if (remark !== undefined && String(remark).trim() !== "") {
        vals.BI = String(remark);
      }
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
    TEXT_OVERRIDE_COLS,
    fillDown,
    toNum,
    computeRows,
    validateRow,
    countErrors,
    snapshotRows,
    restoreRows,
    encodeUnicodeBase64,
    decodeUnicodeBase64,
    findHazardByRec,
    findHazardByName,
  };
});
