/* =====================================================================
 * 「网站数据导入」按钮 —— 对齐内网 192.168.22.73:8000「资料下载 → 生成2个xlsx」
 *
 *   测点布局草稿.xlsx（6 页）：危害因素 / 测点布局情况调查 / 劳动者工作日写实调查 /
 *                              噪声数据 / 下拉内容 / 检测项目
 *   导入模板.xlsx（20 页）  ：劳动定员和职业病危害因素接触情况调查 + 6 张现场调查表 +
 *                              13 张字典页
 *
 * 内容来源：按「数据上传」中**选中的项目编号**，到上游职卫系统
 *   （223.93.144.122:27800，前端 :27900）取真实数据填充：
 *   - 劳动定员（39 列） ← POST api/newEvaluationUnit/export331（返回 xlsx）
 *   - 现场调查 6 表     ← GET api/productionProcess / eqpLayoutInfo / rawMaterialsInfo /
 *                          mainProduct / occupationalProduct(occType=0|1)
 *   - 危害因素 / 检测项目 / 下拉内容 ← 参考库 + js/blf-data.js（内网模板提取的静态字典）
 * 未登录或未选项目时退化为「用当前网页数据生成」。
 * ===================================================================== */
(function (root) {
  'use strict';

  const $ = (id) => document.getElementById(id);

  // ---------------- 列 / 表头定义（与内网文件逐列对齐） ----------------

  // 导入模板 第1页 / 草稿映射区：39 列（= 主表自动计算区 W..BI 同名）
  const WORKER_HEADERS = [
    '*单元/工作场所', '*岗位/工种', '*岗位性质', '*接触类型', '*岗位总工作人数', '*每班最大人数',
    '*作业方式', '*岗位工作班制', '*岗位班制数', '岗位其他班制数', '*岗位周工作天数(d)', '*岗位日工作时长(h)',
    '是否噪声作业岗位', '是否高温作业岗位', '*采样/测量类别', '*点位/采样对象', '工作内容', '*检测项目',
    '*是否采样/测量', '*有机定性分析', '*排除性检测', '*检测方式', '*采样时间类型', '*委外检测',
    '*粉尘性质', '噪声源', '体力劳动强度', '*采样时机', '周工作天数(d)', '日接触时长(h)',
    '*检测天数', '*每天样品数', '*是否合岗', '危害因素来源', '危害因素其他来源', '流量',
    '采样时长', '是否接害', '备注',
  ];

  // 现场调查 6 张表（与调查表页签、导入模板一致）
  const SURVEY_SHEETS = [
    { key: 'process', name: '生产工艺调查', headers: ['*工艺名称', '*单元/工作场所', '*工艺描述'] },
    { key: 'equipment', name: '设备设施调查', headers: ['*设备名称', '*设备型号及规格', '*设备总数', '*设备运行数', '单元/工作场所', '操作岗位(工种)', '工作地点', '*设施布局', '备注'] },
    { key: 'material', name: '原辅物料调查', headers: ['*物料名称', '单元/工作场所', '使用岗位(工种)', '使用地点', '*物理状态', '储存量', '*年用量', '规格', '*主要成分', '*储存方式', '*加药/投料方式', '*运输方式', '*装卸方式', '*装卸周期', '备注'] },
    { key: 'product', name: '主要产品调查', headers: ['*类型', '*产品名称', '单元/工作场所', '影响岗位(工种)', '*产量/产值', '*物理状态', '*主要成分', '*包装方式', '性状', '储存方式'] },
    { key: 'protect', name: '职业防护调查', headers: ['*防护类型', '*防护设施名称', '单元/工作场所', '设置岗位(工种)', '设置地点', '*设施总数', '*设施运行数', '*设施运行情况', '维修情况', '技术参数', '备注'] },
    { key: 'ppe', name: '个体防护调查', headers: ['*防护用品分类', '*防护用品类别', '单元/工作场所', '配置岗位(工种)', '*型号或规格', '*生产厂家', '*佩戴情况', '*更换周期', '*防护性能参数', '备注'] },
  ];

  // 测点布局草稿 - 测点布局情况调查：A..Y 录入块（Z 为提示列）+ AA..BM 映射区（39 列）
  const DRAFT_LEFT = [
    '*单元/工作场所', '*岗位/工种', '*点位/采样对象', '*检测项目', '日接触时长(h)', '周工作天数(d)',
    '*岗位日工作时长(h)', '*岗位周工作天数(d)', '岗位总工作人数', '*岗位工作班制', '*岗位班制数', '岗位其他班制数',
    '工作内容', '作业方式', '*岗位性质', '*是否采样/测量', '*检测方式', '*是否合岗', '*检测天数', '*排除性检测',
    '*采样时机', '危害因素来源', '危害因素其他来源', '体力劳动强度', '备注',
  ];
  const DRAFT_HINT = '红字列可不填，蓝字列特殊情况选填。';
  const DRAFT_RIGHT = WORKER_HEADERS.slice();
  // A..Y 各列取值 → 39 列（A..AM）中的下标
  const LEFT_IDX = [0, 1, 15, 17, 29, 28, 11, 10, 4, 7, 8, 9, 16, 6, 2, 18, 21, 32, 30, 20, 27, 33, 34, 26, 38];
  // 主表数据路径：A..Y 取值来源（主表自动计算区列字母）+ AA..BM 取值来源
  const LEFT_SRC = ['W', 'X', 'AL', 'AN', 'AZ', 'AY', 'AH', 'AG', 'AA', 'AD', 'AE', 'AF', 'AM', 'AC', 'Y', 'AO', 'AR', 'BC', 'BA', 'AQ', 'AX', 'BD', 'BE', 'AW', 'BI'];
  const RIGHT_SRC = ['W', 'X', 'Y', 'Z', 'AA', 'AB', 'AC', 'AD', 'AE', 'AF', 'AG', 'AH', 'AI', 'AJ', 'AK', 'AL', 'AM', 'AN', 'AO', 'AP', 'AQ', 'AR', 'AS', 'AT', 'AU', 'AV', 'AW', 'AX', 'AY', 'AZ', 'BA', 'BB', 'BC', 'BD', 'BE', 'BF', 'BG', 'BH', 'BI'];

  // 危害因素页
  const HAZARD_HEADERS = ['识别', '职业卫生检测管理系统', '粉尘性质', '定性分析', '委外检测', '*计算TWA', '*计算STEL', '*计算CPE', '*计算MAC', '结果保留位数', '存在高毒物品', '不检测原因说明'];
  const HAZARD_KEYS = ['rec', 'name', 'dust', 'qual', 'outsource', 'twa', 'stel', 'cpe', 'mac', 'digits', 'highTox', 'noTestReason'];

  // 劳动者工作日写实调查
  const DAILY_HEADERS = ['*职业病危害因素', '车间', '岗位', '工作地点', '*岗位作业总人数', '岗位接害总人数', '每班最大人数', '体力劳动强度', '工作班制', '班制数', '其他班制数', '作业类型', '作业方式', '工作时间', '接触时间', '来源', '危害因素其他来源', '工作内容'];

  // 测点布局情况调查 AA..BM 数组公式（内网同款；置 true 则写入公式而非算好的值）
  const DRAFT_USE_FORMULAS = false;
  const DRAFT_FORMULAS = {
    AA: '_xlfn.SCAN("",A2:INDEX(A:A,MATCH("*",D:D,-1)),_xlfn.LAMBDA(_xlpm.prev,_xlpm.curr,IF(_xlpm.curr<>"",_xlpm.curr,_xlpm.prev)))',
    AB: '_xlfn.SCAN("",B2:INDEX(B:B,MATCH("*",D:D,-1)),_xlfn.LAMBDA(_xlpm.prev,_xlpm.curr,IF(_xlpm.curr<>"",_xlpm.curr,_xlpm.prev)))',
    AC: '_xlfn.SCAN("",O2:INDEX(O:O,MATCH("*",D:D,-1)),_xlfn.LAMBDA(_xlpm.prev,_xlpm.curr,IF(_xlpm.curr<>"",_xlpm.curr,"固定")))',
    AD: '_xlfn.SCAN("",C2:INDEX(C:C,MATCH("*",D:D,-1)),_xlfn.LAMBDA(_xlpm.prev,_xlpm.curr,"浓度（或强度）相对稳定"))',
    AE: '_xlfn.SCAN("",I2:INDEX(I:I,MATCH("*",D:D,-1)),_xlfn.LAMBDA(_xlpm.prev,_xlpm.curr,IF(_xlpm.curr<>"",_xlpm.curr,_xlpm.prev)))',
    AF: '_xlfn.LET(_xlpm.lastRow,MAX((D:D<>"")*ROW(D:D)),_xlpm.dataRange,AE2:INDEX(AE:AE,_xlpm.lastRow),_xlpm.acRange,AI2:INDEX(AI:AI,_xlpm.lastRow),_xlpm.divisor,_xlfn.SWITCH(_xlpm.acRange,"两班两运转",2,"三班两运转",3,"三班三运转",3,"四班三运转",4,"五班三运转",5,"五班四运转",5,1),_xlpm.divided,_xlpm.dataRange/_xlpm.divisor,IF(MOD(_xlpm.dataRange,_xlpm.divisor)=0,_xlpm.divided,INT(_xlpm.divided)+1))',
    AG: '_xlfn.SCAN("",N2:INDEX(N:N,MATCH("*",D:D,-1)),_xlfn.LAMBDA(_xlpm.prev,_xlpm.curr,IF(_xlpm.curr<>"",_xlpm.curr,_xlpm.prev)))',
    AH: '_xlfn.SCAN("",J2:INDEX(J:J,MATCH("*",D:D,-1)),_xlfn.LAMBDA(_xlpm.prev,_xlpm.curr,IF(_xlpm.curr<>"",_xlpm.curr,_xlpm.prev)))',
    AI: '_xlfn.SCAN("",K2:INDEX(K:K,MATCH("*",D:D,-1)),_xlfn.LAMBDA(_xlpm.prev,_xlpm.curr,IF(_xlpm.curr<>"",_xlpm.curr,_xlpm.prev)))',
    AJ: '_xlfn.SCAN("",L2:INDEX(L:L,MATCH("*",D:D,-1)),_xlfn.LAMBDA(_xlpm.prev,_xlpm.curr,IF(_xlpm.curr<>"",_xlpm.curr,"")))',
    AK: '_xlfn.SCAN("",H2:INDEX(H:H,MATCH("*",D:D,-1)),_xlfn.LAMBDA(_xlpm.prev,_xlpm.curr,IF(_xlpm.curr<>"",_xlpm.curr,_xlpm.prev)))',
    AL: '_xlfn.SCAN("",G2:INDEX(G:G,MATCH("*",D:D,-1)),_xlfn.LAMBDA(_xlpm.prev,_xlpm.curr,IF(_xlpm.curr<>"",_xlpm.curr,_xlpm.prev)))',
    AO: '_xlfn.SCAN("",C2:INDEX(C:C,MATCH("*",D:D,-1)),_xlfn.LAMBDA(_xlpm.prev,_xlpm.curr,IF(INDEX(AV:AV,ROW(_xlpm.curr))="定点","采样点","采样对象")))',
    AP: '_xlfn.SCAN("",C2:INDEX(C:C,MATCH("*",D:D,-1)),_xlfn.LAMBDA(_xlpm.prev,_xlpm.curr,IF(_xlpm.curr<>"",_xlpm.curr,_xlpm.prev)))',
    AQ: '_xlfn.SCAN("",M2:INDEX(M:M,MATCH("*",D:D,-1)),_xlfn.LAMBDA(_xlpm.prev,_xlpm.curr,IF(_xlpm.curr<>"",_xlpm.curr,_xlpm.prev)))',
    AR: 'VLOOKUP(D2:INDEX(D:D,MATCH("*",D:D,-1)),危害因素!A:B,2,FALSE)',
    AS: '_xlfn.SCAN("",P2:INDEX(P:P,MATCH("*",D:D,-1)),_xlfn.LAMBDA(_xlpm.prev,_xlpm.curr,IF(_xlpm.curr<>"",_xlpm.curr,"是")))',
    AT: 'IFERROR(VLOOKUP(D2:INDEX(D:D,MATCH("*",D:D,-1)),危害因素!$A:$J,4,FALSE)&"","")',
    AU: '_xlfn.SCAN("",T2:INDEX(T:T,MATCH("*",D:D,-1)),_xlfn.LAMBDA(_xlpm.prev,_xlpm.curr,IF(_xlpm.curr<>"",_xlpm.curr,"否")))',
    AV: '_xlfn.SCAN("",Q2:INDEX(Q:Q,MATCH("*",D:D,-1)),_xlfn.LAMBDA(_xlpm.prev,_xlpm.curr,IF(_xlpm.curr<>"",_xlpm.curr,"定点")))',
    AW: '_xlfn.LET(_xlpm.lastRow,MAX(2,MATCH("*",D:D,-1)),_xlpm.amRange,AV2:INDEX(AV:AV,_xlpm.lastRow),IF(_xlpm.amRange="定点","短时间",IF(_xlpm.amRange="个体","长时间","")))',
    AX: 'IFERROR(VLOOKUP(D2:INDEX(D:D,MATCH("*",D:D,-1)),危害因素!$A:$J,5,FALSE)&"","")',
    AY: 'IFERROR(VLOOKUP(D2:INDEX(D:D,MATCH("*",D:D,-1)),危害因素!$A:$J,3,FALSE)&"","")',
    AZ: '_xlfn.LET(_xlpm.lastRow,MATCH("*",D:D,-1),_xlpm.dataRange,_xlfn.SEQUENCE(_xlpm.lastRow-1,1,2),_xlpm.pCol,INDEX(AA:AA,_xlpm.dataRange),_xlpm.qCol,INDEX(AB:AB,_xlpm.dataRange),_xlpm.aeCol,INDEX(AR:AR,_xlpm.dataRange),_xlpm.aiCol,INDEX(AV:AV,_xlpm.dataRange),_xlpm.result,_xlfn.MAP(_xlpm.aeCol,_xlpm.aiCol,_xlpm.pCol,_xlpm.qCol,_xlpm.dataRange,_xlfn.LAMBDA(_xlpm.ae,_xlpm.ai,_xlpm.p,_xlpm.q,_xlpm.r,IF(AND(_xlpm.ae="噪声",_xlpm.ai="定点"),_xlfn.LET(_xlpm.matchingRows,_xlfn._xlws.FILTER(_xlpm.dataRange,(_xlpm.pCol=_xlpm.p)*(_xlpm.qCol=_xlpm.q)*(_xlpm.aeCol="噪声")),IF(COUNT(_xlfn._xlws.FILTER(_xlpm.matchingRows,INDEX(AV:AV,_xlpm.matchingRows)="个体"))>0,"是","否")),"否"))),_xlpm.result)',
    BA: '_xlfn.SCAN("",X2:INDEX(X:X,MATCH("*",D:D,-1)),_xlfn.LAMBDA(_xlpm.prev,_xlpm.curr,IF(_xlpm.curr<>"",_xlpm.curr,"")))',
    BB: '_xlfn.SCAN("",U2:INDEX(U:U,MATCH("*",D:D,-1)),_xlfn.LAMBDA(_xlpm.prev,_xlpm.curr,IF(_xlpm.curr<>"",_xlpm.curr,_xlpm.prev)))',
    BC: '_xlfn.SCAN("",F2:INDEX(F:F,MATCH("*",D:D,-1)),_xlfn.LAMBDA(_xlpm.prev,_xlpm.curr,IF(_xlpm.curr<>"",_xlpm.curr,_xlpm.prev)))',
    BD: '_xlfn.SCAN("",E2:INDEX(E:E,MATCH("*",D:D,-1)),_xlfn.LAMBDA(_xlpm.prev,_xlpm.curr,IF(_xlpm.curr<>"",_xlpm.curr,_xlpm.prev)))',
    BE: '_xlfn.LET(_xlpm.lastRow,MATCH("*",D:D,-1),_xlpm.anRange,AR2:INDEX(AR:AR,_xlpm.lastRow),_xlpm.iRange,S2:INDEX(S:S,_xlpm.lastRow),_xlpm.scanResult,_xlfn.SCAN("",_xlpm.iRange,_xlfn.LAMBDA(_xlpm.prev,_xlpm.curr,IF(_xlpm.curr<>"",_xlpm.curr,_xlpm.prev))),_xlfn.MAP(_xlpm.anRange,_xlpm.scanResult,_xlfn.LAMBDA(_xlpm.a,_xlpm.s,IF(OR(_xlpm.a="游离二氧化硅",_xlpm.a="噪声",_xlpm.a="高温",_xlpm.a="工频电场",_xlpm.a="手传振动",_xlpm.a="高频电磁场",_xlpm.a="紫外辐射",ISNUMBER(SEARCH("有机组分定性",_xlpm.a))),1,_xlpm.s))))',
    BF: '_xlfn.LET(_xlpm.lookupRange,D2:INDEX(D:D,MATCH("*",D:D,-1)),_xlpm.resultRows,ROWS(VLOOKUP(_xlpm.lookupRange,危害因素!A:B,2,FALSE)),_xlpm.rowSeq,_xlfn.SEQUENCE(_xlpm.resultRows),IF(_xlpm.rowSeq,_xlfn.MAP(_xlpm.rowSeq,_xlfn.LAMBDA(_xlpm.r,IF(OR(ISNUMBER(SEARCH("有机组分定性",INDEX(AR:AR,_xlpm.r+1))),INDEX(AV:AV,_xlpm.r+1)="个体",INDEX(AR:AR,_xlpm.r+1)="游离二氧化硅",INDEX(AR:AR,_xlpm.r+1)="工频电场",INDEX(AR:AR,_xlpm.r+1)="高频电磁场",INDEX(AR:AR,_xlpm.r+1)="激光辐射"),1,IF(OR(INDEX(AR:AR,_xlpm.r+1)="噪声",INDEX(AR:AR,_xlpm.r+1)="高温",INDEX(AR:AR,_xlpm.r+1)="手传振动"),3,IF(INDEX(AR:AR,_xlpm.r+1)="紫外辐射",1,IF(INDEX(BD:BD,_xlpm.r+1)=0.5,2,IF(INDEX(BD:BD,_xlpm.r+1)=0.25,1,3)))))))))',
    BG: '_xlfn.SCAN("",R2:INDEX(R:R,MATCH("*",D:D,-1)),_xlfn.LAMBDA(_xlpm.prev,_xlpm.curr,IF(_xlpm.curr<>"",_xlpm.curr,"否")))',
    BH: '_xlfn.SCAN("",V2:INDEX(V:V,MATCH("*",D:D,-1)),_xlfn.LAMBDA(_xlpm.prev,_xlpm.curr,IF(_xlpm.curr<>"",_xlpm.curr,_xlpm.prev)))',
    BI: '_xlfn.SCAN("",W2:INDEX(W:W,MATCH("*",D:D,-1)),_xlfn.LAMBDA(_xlpm.prev,_xlpm.curr,IF(_xlpm.curr<>"",_xlpm.curr,"")))',
    BM: '_xlfn.SCAN("",Y2:INDEX(Y:Y,MATCH("*",D:D,-1)),_xlfn.LAMBDA(_xlpm.prev,_xlpm.curr,IF(_xlpm.curr<>"",_xlpm.curr,"")))',
  };
  const COL_LETTERS = (function () {
    const out = [];
    for (let i = 0; i < 60; i++) {
      let s = '', n = i + 1;
      while (n > 0) { const r = (n - 1) % 26; s = String.fromCharCode(65 + r) + s; n = Math.floor((n - 1) / 26); }
      out.push(s);
    }
    return out;
  })();

  // ---------------- 依赖懒加载 ----------------

  function loadScript(src) {
    return new Promise((resolve, reject) => {
      const s = document.createElement('script');
      s.src = src;
      s.onload = resolve;
      s.onerror = () => reject(new Error('加载失败：' + src));
      document.head.appendChild(s);
    });
  }

  let xlsxPromise = null;
  function ensureXlsx() {
    if (root.SamplingXlsx) return Promise.resolve(root.SamplingXlsx);
    if (!xlsxPromise) {
      xlsxPromise = (root.SamplingApp && root.SamplingApp.loadScript
        ? root.SamplingApp.loadScript('js/jszip.min.js').then(() => root.SamplingApp.loadScript('js/xlsxio.js'))
        : Promise.all([loadScript('js/jszip.min.js'), loadScript('js/xlsxio.js')])
      ).then(() => root.SamplingXlsx);
    }
    return xlsxPromise;
  }

  function data() { return root.SamplingBlfData || {}; }

  function widthArray(obj, colCount) {
    if (!obj) return null;
    const arr = [];
    for (let i = 0; i < colCount; i++) arr.push(obj[COL_LETTERS[i]] || 9);
    return arr;
  }

  function txt(v) { return v === null || v === undefined ? '' : v; }

  // ---------------- 本地数据采集 ----------------

  function appRows() {
    const app = root.SamplingApp;
    if (!app || typeof app.rowsSnapshot !== 'function') return [];
    return app.rowsSnapshot() || [];
  }

  function library() {
    const app = root.SamplingApp;
    if (app && typeof app.librarySnapshot === 'function') {
      const lib = app.librarySnapshot();
      if (lib && (lib.hazardFactors || lib.detectionItems)) return lib;
    }
    const fallback = root.SamplingLibrary || {};
    return { hazardFactors: fallback.hazardFactors || [], detectionItems: fallback.detectionItems || [] };
  }

  function surveyData() {
    try {
      if (root.SurveySheets && typeof root.SurveySheets.getData === 'function') return root.SurveySheets.getData() || {};
    } catch (e) { /* ignore */ }
    return {};
  }

  // ---------------- 上游取数（223.93.144.122:27800） ----------------

  function upstreamCtx() {
    const u = root.SamplingUpload;
    if (u && typeof u.getUpstreamContext === 'function') {
      try { return u.getUpstreamContext(); } catch (e) { return null; }
    }
    return null;
  }

  async function upRequest(c, method, path, body) {
    const headers = {};
    if (c.token) headers['Authorization'] = c.token;
    if (c.orgId) headers['organizationId'] = String(c.orgId);
    const init = { method: method, headers: headers };
    if (body !== undefined) {
      headers['Content-Type'] = 'application/json';
      init.body = JSON.stringify(body);
    }
    const r = await fetch(c.apiBase + path, init);
    if (!r.ok) throw new Error('HTTP ' + r.status);
    return r;
  }

  // 兼容上游多种列表返回结构
  function pickRows(j) {
    const cands = [j, j && j.body, j && j.data, j && j.rows, j && j.records];
    for (const c of cands) {
      if (Array.isArray(c)) return c;
      if (c && typeof c === 'object') {
        for (const k of ['records', 'rows', 'list', 'data']) if (Array.isArray(c[k])) return c[k];
      }
    }
    return [];
  }

  // 现场调查 6 表：上游字段 → 导入模板列（字段名取自上游职卫系统）
  const SURVEY_MAP = {
    process: (r) => [r.processName, r.belongUnitName, r.processDescribe],
    equipment: (r) => [r.eqpName, r.eqpType, r.eqpNum, r.eqpUse, r.belongUnitName, r.belongJobName, r.belongSiteName, r.layout, r.remark],
    material: (r) => [r.materialName, r.belongUnitName, r.belongJobName, r.belongSiteName, r.materialForm, r.storageCapacity, r.wastage, r.specifications, r.mainBasis, r.storageMethod, r.feedingMethod, r.transportMethod, r.loadMethod, r.loadCycle, r.reamark],
    product: (r) => [r.type, r.name, r.belongUnitName, r.belongJobName, r.number, r.physicalForm, r.mainBasis, r.packMethod, r.xz, r.storageMode],
    protect: (r) => [r.eqpType, r.name, r.belongUnitName, r.belongJobName, r.installationPosition, r.totalNum, r.useNum, r.eqpDescribe, r.repairSituation, r.facilityParameters, r.notes],
    ppe: (r) => [r.classification || r.category, (r.name || '') + (r.protectModel ? ',' + r.protectModel : ''), r.belongUnitName, r.belongJobName, r.gtSpecifications, r.manufacturer, r.eqpDescribe, r.replacementCycle, r.protectionParameters, r.notes],
  };

  // 上游为空时内网会补的占位值（列下标从 0 起；对齐内网样例输出）
  const SURVEY_EMPTY = {
    material: { 8: '/', 13: '按需' },
    product: { 0: '产品', 6: '/' },
    protect: { 0: '防毒', 5: '/', 6: '/', 8: '日常维护' },
    ppe: { 0: '呼吸防护', 4: '/', 5: '/', 7: '按需更换', 8: '/' },
  };

  // [键, 上游资源, occ 过滤值]；occupationalProduct 上游不区分，取回后按行内 occupationalType 过滤
  const SURVEY_RES = [
    ['process', 'productionProcess', null],
    ['equipment', 'eqpLayoutInfo', null],
    ['material', 'rawMaterialsInfo', null],
    ['product', 'projectInspect/mainProduct/select', null],
    ['protect', 'occupationalProduct', '0'],
    ['ppe', 'occupationalProduct', '1'],
  ];

  function normHeader(s) { return String(s === null || s === undefined ? '' : s).replace(/[*\s]/g, ''); }

  // 上游导出的「劳动定员」xlsx（43 列）→ 本表 39 列（按表头名对齐）
  function workerTo39(aoa) {
    if (!aoa || !aoa.length) return [];
    const hdr = (aoa[0] || []).map(normHeader);
    const idx = WORKER_HEADERS.map((h) => hdr.indexOf(normHeader(h)));
    const rows = [];
    for (let i = 1; i < aoa.length; i++) {
      const r = aoa[i] || [];
      if (!r.some((v) => v !== '' && v !== null && v !== undefined)) continue;
      rows.push(idx.map((k) => (k >= 0 ? txt(r[k]) : '')));
    }
    return rows;
  }

  async function fetchUpstream(c) {
    const pid = c.project.id;
    const out = { worker: null, worker39: null, survey: {}, errors: [] };
    // 1) 劳动定员：上游 export331 直接返回 xlsx
    try {
      const r = await upRequest(c, 'POST', '/api/newEvaluationUnit/export331', { belongProject: pid });
      const blob = await r.blob();
      const X = await ensureXlsx();
      const sheets = await X.readWorkbook(blob);
      if (sheets && sheets.length) {
        out.worker = X.sheetToArray(sheets[0]);
        out.worker39 = workerTo39(out.worker);
      }
    } catch (e) {
      out.errors.push('劳动定员导出失败：' + e.message);
    }
    // 2) 现场调查 6 表
    for (const [key, res, occ] of SURVEY_RES) {
      try {
        let rows;
        if (res === 'occupationalProduct') {
          // 职业防护(0) / 个体防护(1) 上游是同一张表，靠 occupationalType 区分：
          // 带上参数请求，再按行内字段兜底过滤（参数被忽略时也能拆开）
          const r = await upRequest(
            c, 'GET',
            '/api/' + res + '?belongProject=' + encodeURIComponent(pid) +
            '&occupationalType=' + occ + '&pageNumber=1&pageSize=-1'
          );
          const all = pickRows(await r.json());
          const hasType = all.some((x) => txt(x.occupationalType) !== '');
          rows = hasType ? all.filter((x) => String(txt(x.occupationalType)) === String(occ)) : all;
        } else {
          const q = '?belongProject=' + encodeURIComponent(pid) + '&pageNumber=1&pageSize=-1';
          const r = await upRequest(c, 'GET', '/api/' + res + q);
          rows = pickRows(await r.json());
        }
        out.survey[key] = rows;
      } catch (e) {
        out.survey[key] = null;
        out.errors.push(res + ' 取数失败：' + e.message);
      }
    }
    return out;
  }

  // ---------------- 组装：测点布局草稿 ----------------

  function buildDraftSheets(up) {
    const lib = library();
    const D = data();

    // 1) 危害因素
    const hazards = (lib.hazardFactors || []).filter((h) => h && typeof h === 'object');
    const hazardRows = [HAZARD_HEADERS].concat(hazards.map((h) => HAZARD_KEYS.map((k) => txt(h[k]))));

    // 2) 测点布局情况调查
    const draftHeader = DRAFT_LEFT.concat([DRAFT_HINT], DRAFT_RIGHT);
    const draftBody = [];
    const useUp = !!(up && up.worker39 && up.worker39.length);
    if (useUp) {
      for (const r39 of up.worker39) {
        draftBody.push(LEFT_IDX.map((i) => txt(r39[i])).concat([''], r39.map(txt)));
      }
    } else {
      for (const r of appRows()) {
        const input = r.input || {};
        const vals = r.values || {};
        const left = LEFT_SRC.map((c) => txt(vals[c] !== undefined && vals[c] !== '' ? vals[c] : input[c]));
        draftBody.push(left.concat([''], RIGHT_SRC.map((c) => txt(vals[c]))));
      }
    }
    // 可选取回内网同款数组公式（DRAFT_USE_FORMULAS）
    if (DRAFT_USE_FORMULAS && draftBody.length) {
      const lastRow = draftBody.length + 1;
      DRAFT_RIGHT.forEach((_, ci) => {
        const col = COL_LETTERS[26 + ci];
        const f = DRAFT_FORMULAS[col];
        if (f) draftBody[0][26 + ci] = { f: f, ref: col + '2:' + col + lastRow };
      });
    }

    // 3) 劳动者工作日写实调查（表头 + 主表同源写实行；上游模式下留空）
    const dailyRows = [DAILY_HEADERS];
    if (!useUp) {
      for (const r of appRows()) {
        const vals = r.values || {};
        dailyRows.push([
          txt(vals['AN']), txt(vals['W']), txt(vals['X']), txt(vals['AL']),
          txt(vals['AA']), '', txt(vals['AB']), txt(vals['AW']),
          txt(vals['AD']), txt(vals['AE']), txt(vals['AF']),
          '', txt(vals['AC']), '', txt(vals['AZ']), txt(vals['BD']),
          txt(vals['BE']), txt(vals['AM']),
        ]);
      }
    }

    // 4) 噪声数据（内网为空页）
    const noiseRows = [['']];

    // 5) 下拉内容
    const dd = (D.dropdown || []).map((row) => row.map((v) => txt(v)));

    // 6) 检测项目
    const items = (lib.detectionItems || []).map((v) => txt(v));
    const itemRows = [['系统内检测项目名参照表']].concat(items.map((v) => [v]));

    const widthsDraft = D.widthsDraft || {};
    const hazardCount = hazardRows.length - 1;

    return [
      { name: '危害因素', rows: hazardRows, widths: widthArray(widthsDraft['危害因素'], 12) },
      {
        name: '测点布局情况调查',
        rows: [draftHeader].concat(draftBody),
        widths: widthArray(widthsDraft['测点布局情况调查'], DRAFT_LEFT.length + 1 + DRAFT_RIGHT.length),
        freeze: 1,
        dataValidations: [
          { sqref: 'D2:D1048576', formula1: '危害因素!$A$2:$A$' + (hazardCount + 1) },
          { sqref: 'J2:J1048576', formula1: '下拉内容!$D$2:$D$3' },
          { sqref: 'K2:K1048576', formula1: '下拉内容!$E$2:$E$13' },
          { sqref: 'N2:N1048576', formula1: '下拉内容!$C$2:$C$4' },
          { sqref: 'O2:O1048576', formula1: '下拉内容!$A$2:$A$4' },
          { sqref: 'P2:P1048576', formula1: '下拉内容!$G$2:$G$3' },
          { sqref: 'Q2:Q1048576', formula1: '下拉内容!$J$2:$J$3' },
          { sqref: 'R2:R1048576', formula1: '下拉内容!$Q$2:$Q$3' },
          { sqref: 'S2:S1048576', formula1: '下拉内容!$P$2:$P$3' },
          { sqref: 'T2:T1048576', formula1: '下拉内容!$I$2:$I$3' },
          { sqref: 'V2:V1048576', formula1: '下拉内容!$R$2:$R$7' },
          { sqref: 'X2:X1048576', formula1: '下拉内容!$N$2:$N$5' },
        ],
      },
      { name: '劳动者工作日写实调查', rows: dailyRows, widths: widthArray(widthsDraft['劳动者工作日写实调查'], 18) },
      { name: '噪声数据', rows: noiseRows },
      { name: '下拉内容', rows: dd.length ? dd : [['']] },
      { name: '检测项目', rows: itemRows, widths: widthArray(widthsDraft['检测项目'], 1) },
    ];
  }

  // ---------------- 组装：导入模板 ----------------

  function buildTemplateSheets(up) {
    const D = data();
    const widthsTpl = D.widthsTemplate || {};
    const localSv = surveyData();
    const useUp = !!(up && (up.worker39 || up.survey));

    const sheets = [];
    // 1) 劳动定员和职业病危害因素接触情况调查
    const workerBody = useUp && up.worker39 ? up.worker39 : [];
    sheets.push({
      name: '劳动定员和职业病危害因素接触情况调查',
      rows: [WORKER_HEADERS].concat(workerBody.map((r) => WORKER_HEADERS.map((_, i) => txt(r[i])))),
      widths: widthArray(widthsTpl['劳动定员和职业病危害因素接触情况调查'], WORKER_HEADERS.length),
      freeze: 1,
      dataValidations: [
        { sqref: 'M2:M1048576', formula1: '"是,否"' },
        { sqref: 'N2:N1048576', formula1: '"是,否"' },
        { sqref: 'AL2:AL1048576', formula1: '"是,否"' },
      ],
    });
    // 2) 现场调查 6 张表
    for (const sh of SURVEY_SHEETS) {
      let body;
      if (useUp) {
        const src = up.survey ? up.survey[sh.key] : null;
        const mapper = SURVEY_MAP[sh.key];
        const empty = SURVEY_EMPTY[sh.key] || {};
        body = (src || []).map((r) => {
          const row = (mapper ? mapper(r) : []).slice(0, sh.headers.length).map(txt);
          while (row.length < sh.headers.length) row.push('');
          for (const k in empty) if (row[k] === '' || row[k] === null || row[k] === undefined) row[k] = empty[k];
          return row;
        });
      } else {
        body = (localSv[sh.key] || []).map((row) => sh.headers.map((_, i) => txt(row[i])));
      }
      sheets.push({
        name: sh.name,
        rows: [sh.headers].concat(body),
        widths: widthArray(widthsTpl[sh.name], sh.headers.length),
        freeze: 1,
      });
    }
    // 3) 13 张字典页
    for (const d of D.dicts || []) {
      const rows = (d.rows || []).map((r) => r.map((v) => txt(v)));
      sheets.push({ name: d.name, rows: rows.length ? rows : [['']], widths: widthArray(widthsTpl[d.name], (rows[0] || [1]).length) });
    }
    return sheets;
  }

  // ---------------- 入口 ----------------

  function setStatus(msg, ok) {
    const el = $('blf-status');
    if (!el) return;
    el.textContent = msg || '';
    el.style.color = ok === false ? '#cf1322' : ok === true ? '#15803d' : '#64748b';
  }

  function recordName() {
    try {
      const el = document.querySelector('#db-list .db-item.active');
      if (el) {
        const t = (el.getAttribute('data-name') || el.textContent || '').trim();
        if (t) return t.split('\n')[0].trim();
      }
    } catch (e) { /* ignore */ }
    return '';
  }

  async function generate() {
    const btn = $('btn-blf');
    const old = btn ? btn.textContent : '';
    if (btn) { btn.disabled = true; btn.textContent = '生成中…'; }

    const ctx = upstreamCtx();
    let up = null;
    let prefix = '';
    try {
      if (ctx && ctx.project) {
        if (!ctx.loggedIn) {
          setStatus('未登录，无法取上游数据', false);
          alert('尚未登录平台，无法按项目取数。\n请先点「数据上传」登录并选择项目，再回来生成。');
          return;
        }
        setStatus('正在按项目 ' + (ctx.project.code || ctx.project.id) + ' 取上游数据…');
        up = await fetchUpstream(ctx);
        prefix = String(ctx.project.code || '').replace(/[\\/:*?"<>|]/g, '_');
        if (prefix) prefix += '-';
        if (up.errors && up.errors.length) console.warn('[网站数据导入] 部分取数失败：', up.errors);
      } else {
        setStatus('未选择项目：将用当前网页数据生成');
        const nm = recordName().replace(/[\\/:*?"<>|]/g, '_').trim();
        prefix = nm ? nm + '-' : '';
      }

      const X = await ensureXlsx();

      // 测点布局草稿
      const draftBytes = await X.writeWorkbook(buildDraftSheets(up));
      X.downloadBlob(
        new Blob([draftBytes], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' }),
        prefix + '测点布局草稿.xlsx'
      );

      // 导入模板
      const tplBytes = await X.writeWorkbook(buildTemplateSheets(up));
      X.downloadBlob(
        new Blob([tplBytes], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' }),
        prefix + '导入模板.xlsx'
      );

      if (up) {
        const n39 = up.worker39 ? up.worker39.length : 0;
        const nSv = Object.keys(up.survey || {}).reduce((a, k) => a + ((up.survey[k] || []).length), 0);
        const bad = !!(up.errors && up.errors.length);
        setStatus('已生成：定员 ' + n39 + ' 行 / 调查表 ' + nSv + ' 行' + (bad ? '（部分接口失败，见控制台）' : ''), !bad);
      } else {
        setStatus('已生成：测点布局草稿 + 导入模板（网页数据）', true);
      }
    } catch (e) {
      console.error(e);
      setStatus('生成失败：' + (e && e.message ? e.message : e), false);
      alert('网站数据导入失败：' + (e && e.message ? e.message : e));
    } finally {
      if (btn) { btn.disabled = false; btn.textContent = old || '网站数据导入'; }
    }
  }

  root.SamplingBlf = { generate, buildDraftSheets, buildTemplateSheets, fetchUpstream, workerTo39 };

  // 绑定按钮（主工具栏 + 全屏工具栏）
  function bind() {
    ['btn-blf', 'fs-blf'].forEach((id) => {
      const b = $(id);
      if (b && !b.__blfBound) {
        b.__blfBound = true;
        b.addEventListener('click', generate);
      }
    });
  }
  if (document.readyState === 'loading') document.addEventListener('DOMContentLoaded', bind);
  else bind();
})(typeof self !== 'undefined' ? self : this);
