/* =====================================================================
 * 「网站数据导入」按钮 —— 对齐内网 192.168.22.73:8000「资料下载 → 生成2个xlsx」
 *   取其「导入模板.xlsx」（20 页）的格式：
 *     劳动定员和职业病危害因素接触情况调查 + 6 张现场调查表 + 13 张字典页
 *
 * 内容来源：按「数据上传」中**选中的项目编号**，到上游职卫系统
 *   （223.93.144.122:27800，前端 :27900）取真实数据填充：
 *   - 劳动定员（39 列） ← POST api/newEvaluationUnit/export331（返回 xlsx）
 *   - 现场调查 6 表     ← GET api/productionProcess / eqpLayoutInfo / rawMaterialsInfo /
 *                          mainProduct / occupationalProduct(occType=0|1)
 *   - 字典页 ← js/blf-data.js（内网模板提取的静态字典）
 * 未登录或未选择项目时，提示先到「数据上传」查询并选中项目，不再用当前网页数据兜底。
 *
 * 生成后**不下载文件**，而是通过网页既有的「导入 Excel」逻辑
 * （SamplingApp.importWorkbook）直接导入到页面：主表格 + 调查表页签。
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
  // 统一走共享工具 js/util.js；本文件在 Node 测试里会被单独加载（无 SamplingUtil），
  // 故保留一份等价的最小兜底，保证独立可运行。
  const Util = root.SamplingUtil || null;

  function loadScript(src) {
    if (Util && Util.loadScript) return Util.loadScript(src);
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
        : loadScript('js/jszip.min.js').then(() => loadScript('js/xlsxio.js'))
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

  function surveyData() {
    try {
      if (root.SurveySheets && typeof root.SurveySheets.getData === 'function') return root.SurveySheets.getData() || {};
    } catch (e) { /* ignore */ }
    return {};
  }

  // ---------------- 上游取数（223.93.144.122:27800） ----------------

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
    // 登录失效（401）：交给统一处理（提示 + 引导重新登录）
    if (r.status === 401) {
      const u = root.SamplingUpload;
      if (u && typeof u.notifyAuthExpired === 'function') u.notifyAuthExpired();
    }
    if (!r.ok) throw new Error('HTTP ' + r.status);
    return r;
  }

  // 兼容上游多种列表返回结构 / 字段改名 —— 统一走共享工具（js/util.js），
  // 单独加载时用等价本地实现兜底。
  function pickRows(j) {
    if (Util && Util.pickRows) return Util.pickRows(j);
    if (Array.isArray(j)) return j;
    const cands = [j, j && j.body, j && j.data, j && j.rows, j && j.records];
    for (const c of cands) {
      if (Array.isArray(c)) return c;
      if (c && typeof c === 'object') {
        for (const k of ['records', 'rows', 'list', 'data']) if (Array.isArray(c[k])) return c[k];
      }
    }
    return [];
  }

  // 上游字段在新旧版本中会改名（如 belongUnitName ↔ belongUnit、facilityOperations ↔ useNum）：
  // 按候选顺序取第一个非空值，做到新旧版兼容
  function pick(r, keys) {
    if (Util && Util.pickField) return Util.pickField(r, keys);
    for (const k of keys) {
      const v = r[k];
      if (v !== undefined && v !== null && v !== '') return v;
    }
    return '';
  }

  // 产品类型字典（上游「主要产品」的 type 存的是字典 id；取值来自平台前端固定字典）
  const PRODUCT_TYPE = { '0': '产品', '1': '中间产品', '2': '副产品', '3': '联产品' };

  // 「佩戴情况」：新版是文本字段 wearingSituation；旧版是编码字段 adorn（1/2/3）
  const ADORN_TEXT = { '1': '有效', '2': '部分有', '3': '无' };
  function wearingSituation(r) {
    if (txt(r.wearingSituation) !== '') return r.wearingSituation;
    const code = String(txt(r.adorn));
    return ADORN_TEXT[code] || txt(r.adorn);
  }

  // 现场调查 6 表：上游字段 → 导入模板列（字段名取自上游职卫系统，含新旧版候选）
  const SURVEY_MAP = {
    process: (r) => [r.processName, pick(r, ['belongUnitName', 'belongUnit']), r.processDescribe],
    equipment: (r) => [r.eqpName, r.eqpType, r.eqpNum, r.eqpUse, pick(r, ['belongUnitName', 'belongUnit']), pick(r, ['belongJobName', 'belongJob']), r.belongSiteName, r.layout, pick(r, ['remark', 'reamark', 'notes'])],
    material: (r) => [r.materialName, pick(r, ['belongUnitName', 'belongUnit']), pick(r, ['belongJobName', 'belongJob']), r.belongSiteName, r.materialForm, r.storageCapacity, r.wastage, r.specifications, r.mainBasis, r.storageMethod, r.feedingMethod, r.transportMethod, r.loadMethod, r.loadCycle, pick(r, ['reamark', 'remark', 'notes'])],
    // 类型：type 是字典 id，需转成名称（否则会显示成 0/1/2/3）
    product: (r) => [PRODUCT_TYPE[String(txt(r.type))] || txt(r.type), r.name, pick(r, ['belongUnitName', 'belongUnit']), pick(r, ['belongJobName', 'belongJob']), r.number, r.physicalForm, r.mainBasis, r.packMethod, r.xz, r.storageMode],
    // 职业防护：设施运行数 = facilityOperations（新版）/ useNum（旧版）
    protect: (r) => [r.eqpType, r.name, pick(r, ['belongUnitName', 'belongUnit']), pick(r, ['belongJobName', 'belongJob']), r.installationPosition, r.totalNum, pick(r, ['facilityOperations', 'useNum']), r.eqpDescribe, r.repairSituation, r.facilityParameters, pick(r, ['notes', 'reamark', 'remark'])],
    // 个体防护：佩戴情况 = wearingSituation / adorn，型号或规格 = protectModel
    ppe: (r) => [pick(r, ['classification', 'category']), r.name, pick(r, ['belongUnitName', 'belongUnit']), pick(r, ['belongJobName', 'belongJob']), r.protectModel, r.manufacturer, wearingSituation(r), r.replacementCycle, r.protectionParameters, pick(r, ['notes', 'reamark', 'remark'])],
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

  // 取上游数据。opts.worker / opts.survey 可只取其中一部分：
  //   · 采样方案导入 → { worker: true,  survey: false }
  //   · 调查表导入 → { worker: false, survey: true }
  //   不传 opts 时两部分都取（保持旧行为）。
  async function fetchUpstream(c, opts) {
    const wantWorker = !opts || opts.worker !== false;
    const wantSurvey = !opts || opts.survey !== false;
    const pid = c.project.id;
    const out = { worker: null, worker39: null, survey: {}, processPictures: [], errors: [] };
    // 1) 劳动定员：上游 export331 直接返回 xlsx（主表格数据来源）
    if (wantWorker) {
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
    }
    // 2) 现场调查 6 表（调查表数据来源）
    if (wantSurvey) {
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
          // 生产工艺：顺带记下每条记录的工艺图路径（processPicturePath），导入时写入「生产工艺调查」的工艺图槽位
          if (key === 'process') {
            out.processPictures = rows
              .map((r) => ({ name: String(txt(r.processName)).trim(), path: String(txt(r.processPicturePath)).trim() }))
              .filter((x) => x.name && x.path);
          }
        } catch (e) {
          out.survey[key] = null;
          out.errors.push(res + ' 取数失败：' + e.message);
        }
      }
    }
    return out;
  }

  // 下载平台的工艺图（processPicturePath），压缩后写入「生产工艺调查」对应工艺名称的槽位。
  // 返回 { total, ok, skip, fail }；已有图片也覆盖（以平台为准）。
  async function importProcessPictures(c, list) {
    const res = { total: (list || []).length, ok: 0, skip: 0, fail: 0 };
    if (!res.total) return res;
    const sv = root.SurveySheets;
    if (!sv || typeof sv.importProcessImages !== 'function') return res;
    const items = [];
    for (const it of list) {
      try {
        const enc = String(it.path).split('/').map(encodeURIComponent).join('/');
        const r = await fetch(c.apiBase + enc, { headers: c.token ? { Authorization: c.token } : {} });
        if (!r.ok) { res.fail++; continue; }
        items.push({ name: it.name, blob: await r.blob() });
      } catch (e) { res.fail++; }
    }
    const r2 = await sv.importProcessImages(items, true);
    res.ok = r2.ok; res.skip += r2.skip; res.fail += r2.fail;
    return res;
  }

  // ---------------- 工艺、点位图保存（直连平台接口，秒级） ----------------
  // 两组图片来源（都直接取，**不生成报告**）：
  //   ① 工艺流程图      ← api/productionProcess        → processPicturePath
  //   ② 采样/测量布点图  ← api/samplingPlan/projectLayout → layoutPicturePath
  // 实测：报告「3.2.1 生产工艺」与「附录2 采样/测量布点图」里嵌的就是这两组图（字节完全一致），
  //       而生成一份检评报告要 ~21s，直连取图只要 ~1s。

  let jszipPromise = null;
  function ensureJSZip() {
    if (root.JSZip) return Promise.resolve(root.JSZip);
    if (!jszipPromise) {
      jszipPromise = (root.SamplingApp && root.SamplingApp.loadScript
        ? root.SamplingApp.loadScript('js/jszip.min.js')
        : loadScript('js/jszip.min.js')
      ).then(() => {
        if (!root.JSZip) throw new Error('jszip 加载失败');
        return root.JSZip;
      });
    }
    return jszipPromise;
  }

  // 取 zip 条目文本：优先「uint8array + 原生 TextDecoder」，退化到 JSZip 自带的 async('string')。
  // ★ 性能（2026-10-18）：JSZip 的 async('string') 走它自己的 JS 版 UTF-8 解码器（分块 worker），
  //   在 6MB 的 document.xml 上明显慢于原生 TextDecoder（实测 ~162ms → ~130ms）。
  //   注意：Node 的 vm 沙箱里可能没有 TextDecoder 全局，故必须做存在性判断后回退。
  async function readZipEntryText(file) {
    if (typeof TextDecoder !== 'undefined') {
      try {
        const u8 = await file.async('uint8array');
        return new TextDecoder('utf-8').decode(u8);
      } catch (e) { /* 回退 */ }
    }
    return await file.async('string');
  }

  function saveBlobAs(blob, name) {
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = name;
    document.body.appendChild(a);
    a.click();
    setTimeout(() => { URL.revokeObjectURL(url); a.remove(); }, 5000);
  }

  // ---------------- 生成检评报告（Word） ----------------
  // 链路（与平台「检评报告」页一致）：
  //   ① 模板列表  GET /api/pjbgModule?pageNumber=1&pageSize=-1&organizationId=<orgId>
  //   ② 生成报告  GET /api/jpbgReportData/jpbgYL?projectId=&moduleId=&organizationId=  → body = PDF 路径
  //   ③ 同路径把 .pdf 换成 .docx 下载（平台 Word 与 PDF 同目录同名）
  // ⚠ 平台生成一次约 20s 且**无缓存**，调用期间要把按钮禁用并给出进度提示。
  async function genPjbgDocx(c) {
    const tr = await upRequest(c, 'GET', '/api/pjbgModule?pageNumber=1&pageSize=-1&organizationId=' + encodeURIComponent(c.orgId || ''));
    const recs = pickRows(await tr.json());
    if (!recs.length) throw new Error('未获取到检评报告模板，请先在平台创建模板');
    let mod = null;
    for (const x of recs) { if (/检评/.test(String(x.name || ''))) { mod = x; break; } }
    if (!mod) mod = recs[0];
    const rr = await upRequest(
      c, 'GET',
      '/api/jpbgReportData/jpbgYL?projectId=' + encodeURIComponent(c.project.id) +
        '&moduleId=' + encodeURIComponent(mod.id) + '&organizationId=' + encodeURIComponent(c.orgId || '')
    );
    const j = await rr.json();
    const body = j && j.body;
    if (!body) throw new Error('生成检评报告失败：' + ((j && j.message) || '无返回'));
    const enc = String(body).replace(/\.pdf$/i, '.docx').split('/').map(encodeURIComponent).join('/');
    const resp = await fetch(c.apiBase + enc, { headers: c.token ? { Authorization: c.token } : {} });
    if (!resp.ok) throw new Error('下载报告失败(HTTP ' + resp.status + ')');
    // 文件名与平台「直接导出」完全一致：平台用的是 wordUrl.substring(lastIndexOf('/')+1)，
    // 即 body 路径的最后一段（把 .pdf 换成 .docx），例如
    // BTC26-SZJC0959-浙江永贵电器股份有限公司_评价报告_49.docx
    const fileName = String(body).split('/').pop().replace(/\.pdf$/i, '.docx');
    const raw = await resp.arrayBuffer();
    // 报告里「存在高毒物品」列平台有时输出 0/1 → 按字典换成 否/是
    const buf = await patchPjbgDocx(raw);
    return { buf: buf, modName: mod.name || '', fileName: fileName };
  }

  // ---------------- 报告后处理：「存在高毒物品」列 0/1 → 否/是 ----------------
  // 平台模板会把该列的原始标记直接印出来（0=否、1=是），需按字典换成中文再交付。
  const HIGHTOX_HEAD = '存在高毒物品';

  // ---------------- 报告后处理：共用工具 ----------------
  // ★ 段落正则只在这里定义一份：① ~ ⑦ 所有「逐段处理」的步骤都走 `mapParagraphs` / `PARA_SRC`，
  //   避免同一条正则在各处重复书写、改一处漏一处。
  const PARA_SRC = '<w:p(?:\\s[^>]*)?>[\\s\\S]*?<\\/w:p>';
  const TBL_SRC = '<w:tbl(?:\\s[^>]*)?>[\\s\\S]*?<\\/w:tbl>';
  function mapParagraphs(xml, cb) { return xml.replace(new RegExp(PARA_SRC, 'g'), cb); }
  function mapTables(xml, cb) { return xml.replace(new RegExp(TBL_SRC, 'g'), cb); }
  // ★ 性能（2026-10-21）：同一步里常要连续查好几张表（表3-1 / 表4-1 / 表8-1 …），
  //   每次 tablesOf 都要在 6MB XML 上跑一遍全局正则并生成 21 张表的字符串（约 9ms）。
  //   这里按「同一个 xml 字符串引用」缓存一份（字符串不可变，引用相同 = 内容相同）：
  //   步骤之间 xml 会变，缓存自然失效，不会拿到过期结果。
  let TBL_CACHE_XML = null;
  let TBL_CACHE_LIST = null;
  function tablesOf(xml) {
    if (xml === TBL_CACHE_XML && TBL_CACHE_LIST) return TBL_CACHE_LIST;
    const list = xml.match(new RegExp(TBL_SRC, 'g')) || [];
    TBL_CACHE_XML = xml;
    TBL_CACHE_LIST = list;
    return list;
  }
  function eachParagraph(xml, cb) {                    // 需要「匹配位置」时用
    const re = new RegExp(PARA_SRC, 'g');
    let m;
    while ((m = re.exec(xml))) cb(m[0], m.index);
  }

  // ★ 只处理「包含指定字符」的段落，避免为了少数几个段落而全篇扫描。
  //   needles 必须是**单字**（或至少是绝不会被 run 拆开的片段）——因为目标句可能被拆到多个 <w:t> 里，
  //   但其中每个汉字仍原样出现在 XML 中，用单字做「必要预判」是安全的。
  //   实现：indexOf 定位 → 向外扩到 <w:p …> … </w:p> 边界 → 去重 → **从后往前**替换（偏移不失效）。
  function paraStartBefore(xml, pos) {
    let s = xml.lastIndexOf('<w:p', pos);
    while (s >= 0) {
      const c = xml.charAt(s + 4);                     // 排除 <w:pPr> 之类
      if (c === '>' || c === ' ' || c === '\t' || c === '\n' || c === '\r') return s;
      s = xml.lastIndexOf('<w:p', s - 1);
    }
    return -1;
  }
  function mapParasContaining(xml, needles, cb) {
    const list = (needles || []).filter(Boolean);
    if (!list.length) return xml;
    const spans = [];
    const seen = new Set();
    for (const nd of list) {
      let i = xml.indexOf(nd);
      while (i >= 0) {
        const s = paraStartBefore(xml, i);
        const e = xml.indexOf('</w:p>', i);
        if (s >= 0 && e >= 0) {
          const end = e + '</w:p>'.length;
          const key = s + ':' + end;
          if (!seen.has(key)) { seen.add(key); spans.push({ s: s, e: end }); }
        }
        i = xml.indexOf(nd, i + nd.length);
      }
    }
    if (!spans.length) return xml;
    spans.sort(function (a, b) { return b.s - a.s; });   // 从后往前，前面的偏移仍然有效
    let out = xml;
    for (const sp of spans) {
      const p = out.slice(sp.s, sp.e);
      const np = cb(p);
      if (np !== p) out = out.slice(0, sp.s) + np + out.slice(sp.e);
    }
    return out;
  }

  // 取一个 <w:tc>…</w:tc> 里的可见文本。
  // <w:br/> 的两种写法（与 <w:t> 并列 / 写在 <w:t> 内部）都还原成换行。
  //
  // ★ 性能（2026-10-17 重构）：原来是「先把所有非 <w:t> 标签用负向前瞻正则剥掉」，
  //   在 7.7MB 文档 / 6000+ 段落上是最贵的一步（`/<(?!\/?w:t(?:\s|>))[^>]*>/g`）。
  //   现在改为**一次正向扫描**，只匹配 `<w:t>` / `<w:br/>` / `<w:tab/>` 三类，输出完全等价。
  const CELL_TOK_RE = /<w:t(?:\s[^>]*)?>([\s\S]*?)<\/w:t>|<w:br(?:\s[^>]*)?\/?>|<w:tab\s*\/?>/g;
  function cellText(tcXml) {
    const s = String(tcXml === null || tcXml === undefined ? '' : tcXml);
    if (s.indexOf('<w:t') < 0) return '';          // 没有文本节点 → 快速返回（图片/空单元格）
    CELL_TOK_RE.lastIndex = 0;
    let out = '';
    let m;
    while ((m = CELL_TOK_RE.exec(s))) {
      if (m[1] !== undefined) {
        let v = m[1];
        // 少数报告把 <w:br/> / <w:tab/> 写进了 <w:t> 内部，先就地还原
        if (v.indexOf('<') >= 0) {
          v = v.replace(/<w:br(?:\s[^>]*)?\/?>/g, '\u0001').replace(/<w:tab\s*\/?>/g, '\t');
        }
        out += v;
      } else if (m[0].indexOf('<w:br') === 0) {
        out += '\u0001';
      } else {
        out += '\t';
      }
    }
    return out
      .replace(/\u0001/g, '\n')
      .replace(/&lt;/g, '<').replace(/&gt;/g, '>').replace(/&quot;/g, '"')
      .replace(/&apos;/g, "'").replace(/&amp;/g, '&')
      .trim();
  }

  // 把段落里所有 <w:t> 的内容按顺序拼起来（**不 trim、不处理换行**）。
  // 与 cellText 的区别：这里的字符偏移与 <w:t> 内容**一一对应**，可直接用于 spliceParaRange。
  const PARA_T_RE = /<w:t(?:\s[^>]*)?>([\s\S]*?)<\/w:t>/g;
  function paraTextConcat(pXml) {
    PARA_T_RE.lastIndex = 0;
    let s = '';
    let m;
    while ((m = PARA_T_RE.exec(pXml))) s += (m[1] || '');
    return s;
  }
  // 把单元格/段落文本改成 text：第一个 <w:t> 承载内容，其余清空（保留段落结构）
  function setCellText(tcXml, text) {
    let first = true;
    return tcXml.replace(/<w:t(?:\s[^>]*)?>[\s\S]*?<\/w:t>/g, function () {
      if (first) { first = false; return '<w:t xml:space="preserve">' + text + '</w:t>'; }
      return '<w:t xml:space="preserve"></w:t>';
    });
  }

  // 在 document.xml 里找表头含「高毒物品」的表，把该列的值规整成 是 / 否：
  //   · 平台模板有时把原始标记直接印出来（0=否、1=是）→ 按字典换成中文；
  //   · 平台导出偶尔把「否」印成占位横线「-」（实测 BTC26-SZJC1095「生产车间(1F) / 打包岗位 / 打包」
  //     一行：平台「检测报告编辑」里明明是「否」，导出的 docx 却是「-」）→ 该列是二值属性，
  //     不存在「不适用」的情形，占位横线一律按「否」处理。
  // 表头可能被拆成多个 run / 带软换行（如「存在高\n毒物品」），故用**包含**匹配而非全等。
  // 返回 { xml, changed }。
  const HIGHTOX_PLACEHOLDER_NO = ['-', '－', '—', '——', '–'];  // 半角/全角横线、破折号
  function fixHighToxColumn(xml) {
    // ★ 快速早退：整篇没有「高毒」/「毒物品」就不必逐表扫描。
    //   ⚠ 不能用完整表头「存在高毒物品」做判据——表头常被拆成多个 run 或带软换行
    //     （如「存在高」+<w:br/>+「毒物品」），完整串根本不存在，会导致**整列漏处理**。
    if (xml.indexOf('高毒') < 0 && xml.indexOf('毒物品') < 0) return { xml: xml, changed: 0 };
    let changed = 0;
    const out = mapTables(xml, function (tbl) {
      // ★ 性能：先只取**第一行（表头）**判断是不是目标表——避免为 20 张无关表解析全部行。
      const firstRow = /<w:tr(?:\s[^>]*)?>[\s\S]*?<\/w:tr>/.exec(tbl);
      if (!firstRow) return tbl;
      const headCells = firstRow[0].match(/<w:tc(?:\s[^>]*)?>[\s\S]*?<\/w:tc>/g) || [];
      let col = -1;
      for (let i = 0; i < headCells.length; i++) {
        // 去空白后比较：表头若带软换行，cellText 会返回「存在高\n毒物品」，直接 indexOf 匹配不到
        if (cellText(headCells[i]).replace(/\s+/g, '').indexOf(HIGHTOX_HEAD) >= 0) { col = i; break; }
      }
      if (col < 0) return tbl;   // 不是这张表
      const rows = tbl.match(/<w:tr(?:\s[^>]*)?>[\s\S]*?<\/w:tr>/g) || [];
      if (rows.length < 2) return tbl;
      let newTbl = tbl;
      for (let ri = 1; ri < rows.length; ri++) {
        const cells = rows[ri].match(/<w:tc(?:\s[^>]*)?>[\s\S]*?<\/w:tc>/g) || [];
        if (col >= cells.length) continue;
        const t = cellText(cells[col]);
        let target = null;
        if (t === '1') target = '是';
        else if (t === '0' || HIGHTOX_PLACEHOLDER_NO.indexOf(t) >= 0) target = '否';
        if (!target) continue;
        const newRow = rows[ri].replace(cells[col], setCellText(cells[col], target));
        newTbl = newTbl.replace(rows[ri], newRow);
        changed++;
      }
      return newTbl;
    });
    return { xml: out, changed: changed };
  }

  // ---------------- 报告后处理②：非噪声岗位删掉「噪声」危害 ----------------
  // 针对「表4-1 岗位（工种）设置、劳动者作业情况及职业病危害因素识别情况」的
  // 「职业病危害因素」列：若该行的「单元/工作场所 + 岗位（工种）」属于**非噪声岗位**
  // （即噪声筛选结果里 LEX 最大值 < 80 dB 的那批），就把单元格里的「噪声」项删掉；
  // 同一岗位的多个点位行都会处理。
  const HAZARD_COL_HEAD = '职业病危害因素';
  const NOISE_HAZARD = '噪声';

  // 去空白比较键（共享工具）；单独加载时兜底
  function normKey(s) {
    if (Util && Util.stripWs) return Util.stripWs(s);
    return String(s === null || s === undefined ? '' : s).replace(/\s+/g, '');
  }

  // 从「其他粉尘,噪声,紫外辐射」里删掉「噪声」项；没删到返回 null。
  // 删空（该岗位只剩噪声）时填「-」，与报告其它空列写法一致。
  function dropNoiseItem(text) {
    const raw = String(text === null || text === undefined ? '' : text);
    const items = raw.split(/[,，]/);
    const kept = items
      .map(function (x) { return x.trim(); })
      .filter(function (x) { return x !== '' && x !== NOISE_HAZARD; });
    if (kept.length === items.filter(function (x) { return x.trim() !== ''; }).length) return null;
    return kept.length ? kept.join(',') : '-';
  }

  // nonNoise: [{ unit, job }]
  function fixNonNoiseHazards(xml, nonNoise) {
    if (!nonNoise || !nonNoise.length) return { xml: xml, changed: 0 };
    if (xml.indexOf(HAZARD_COL_HEAD) < 0) return { xml: xml, changed: 0 };
    const set = {};
    nonNoise.forEach(function (p) { set[normKey(p.unit) + '\u0001' + normKey(p.job)] = 1; });
    let changed = 0;
    const out = mapTables(xml, function (tbl) {
      const rows = tbl.match(/<w:tr(?:\s[^>]*)?>[\s\S]*?<\/w:tr>/g) || [];
      if (rows.length < 2) return tbl;
      const headCells = rows[0].match(/<w:tc(?:\s[^>]*)?>[\s\S]*?<\/w:tc>/g) || [];
      let hIdx = -1;
      for (let i = 0; i < headCells.length; i++) {
        if (cellText(headCells[i]).indexOf(HAZARD_COL_HEAD) >= 0) { hIdx = i; break; }
      }
      if (hIdx < 0) return tbl;   // 不是这张表
      let newTbl = tbl;
      for (let ri = 1; ri < rows.length; ri++) {
        const cells = rows[ri].match(/<w:tc(?:\s[^>]*)?>[\s\S]*?<\/w:tc>/g) || [];
        if (cells.length < 3) continue;
        const unit = cellText(cells[0]);
        const job = cellText(cells[1]);
        if (!unit && !job) continue;   // 表头续行（只有「总数 / 数/班」那行）
        if (!set[normKey(unit) + '\u0001' + normKey(job)]) continue;
        // 表头可能因合并单元格比数据行少若干格（本项目「人数」表头 1 格 / 数据行 2 格），按差值平移
        const ci = hIdx + (cells.length - headCells.length);
        if (ci < 0 || ci >= cells.length) continue;
        const text = cellText(cells[ci]);
        if (text.indexOf(NOISE_HAZARD) < 0) continue;   // 兜底：列对不上就不动
        const next = dropNoiseItem(text);
        if (next === null) continue;
        const newRow = rows[ri].replace(cells[ci], setCellText(cells[ci], next));
        newTbl = newTbl.replace(rows[ri], newRow);
        changed++;
      }
      return newTbl;
    });
    return { xml: out, changed: changed };
  }

  // ---------------- 报告后处理②c：噪声岗位内「低噪点位」单独删掉「噪声」危害 ----------------
  // 用户要求（2026-10-19）：某「车间+岗位」的噪声 LEX 最大值 ≥ 80 dB（属噪声作业岗位，② 不整体删噪声）时，
  // 噪声表里该岗位把多个点位挤在**同一个单元格**里（单元格多行），「测量结果[dB(A)」同样多行，
  // 且每个点位固定 3 个测量值、按点位顺序排列（第 1 个点位 = 前 3 行，第 2 个 = 第 4~6 行…）。
  // 逐点位判断：某点位 3 个测量值**全部 < 80 dB** → 该点位实际不接触噪声 →
  // 表4-1 里「车间+岗位+点位」那一行的「职业病危害因素」列也把「噪声」删掉。
  //
  // ⚠ 仅「多行挤在一格」形态才判定：单点位单行的行不产生 quietPointPairs（由 upload.js 侧保证）。
  // 前提（同样由 upload.js 侧算好并保证）：同岗位**至少有一个点位 3 个测量值全部 ≥ 80 dB**，
  //   否则整个岗位的噪声会被删光——那种情况本就该走 ② 的整岗位规则。
  // 匹配口径：
  //   · 只在**表4-1**（表头同时含「职业病危害因素」和「工作地点…」列）上动手，表5-1 检测项目确认表不碰；
  //   · 行键 = 单元/工作场所(第1列) + 岗位（工种）(第2列) + 点位；
  //     表4-1 的点位在「工作地点及内容」列里、取首个「/」前的部分（如「针车2/操作设备针车作业」→ 针车2）；
  //   · 该行「工作地点及内容」可能也是多行（多个点位共用一行/一个危害因素单元格），
  //     此时只有**全部点位**都是低噪点位才删（否则该行仍需保留噪声）。
  const WORKPLACE_COL_HEAD = '工作地点';   // 表4-1「工作地点及内容」列（点位在首个「/」前）

  // 单元格文本按软换行拆行（去掉空行）；点位取首个「/」（兼容全角「／」）前的部分
  function cellLines(text) {
    return String(text === null || text === undefined ? '' : text)
      .split('\n')
      .map(function (s) { return s.trim(); })
      .filter(function (s) { return s !== ''; });
  }
  function beforeSlashText(text) {
    const s = String(text === null || text === undefined ? '' : text);
    let i = s.indexOf('/');
    const j = s.indexOf('／');
    if (i < 0 || (j >= 0 && j < i)) i = j;
    return (i < 0 ? s : s.slice(0, i)).trim();
  }

  // quietPoints: [{ unit, job, point }]
  function fixQuietPointHazards(xml, quietPoints) {
    if (!quietPoints || !quietPoints.length) return { xml: xml, changed: 0 };
    if (xml.indexOf(HAZARD_COL_HEAD) < 0) return { xml: xml, changed: 0 };
    const set = {};
    quietPoints.forEach(function (p) {
      set[normKey(p.unit) + '\u0001' + normKey(p.job) + '\u0001' + normKey(p.point)] = 1;
    });
    let changed = 0;
    const out = mapTables(xml, function (tbl) {
      const rows = tbl.match(/<w:tr(?:\s[^>]*)?>[\s\S]*?<\/w:tr>/g) || [];
      if (rows.length < 2) return tbl;
      const headCells = rows[0].match(/<w:tc(?:\s[^>]*)?>[\s\S]*?<\/w:tc>/g) || [];
      let hIdx = -1;
      let pIdx = -1;
      for (let i = 0; i < headCells.length; i++) {
        const t = cellText(headCells[i]);
        if (hIdx < 0 && t.indexOf(HAZARD_COL_HEAD) >= 0) hIdx = i;
        if (pIdx < 0 && t.indexOf(WORKPLACE_COL_HEAD) >= 0) pIdx = i;
      }
      // 只处理表4-1（有「工作地点及内容」列）；表5-1 检测项目确认表没有该列 → 不动
      if (hIdx < 0 || pIdx < 0) return tbl;
      let newTbl = tbl;
      for (let ri = 1; ri < rows.length; ri++) {
        const cells = rows[ri].match(/<w:tc(?:\s[^>]*)?>[\s\S]*?<\/w:tc>/g) || [];
        if (cells.length < 3) continue;
        const unit = cellText(cells[0]);
        const job = cellText(cells[1]);
        if (!unit && !job) continue;   // 表头续行（只有「总数 / 数/班」那行）
        // 表头可能因合并单元格比数据行少若干格（本项目「人数」表头 1 格 / 数据行 2 格），按差值平移
        const shift = cells.length - headCells.length;
        const pi = pIdx + shift;
        const ci = hIdx + shift;
        if (pi < 0 || pi >= cells.length || ci < 0 || ci >= cells.length) continue;
        const pts = cellLines(cellText(cells[pi]))
          .map(beforeSlashText)
          .filter(function (s) { return s !== ''; });
        if (!pts.length) continue;
        const base = normKey(unit) + '\u0001' + normKey(job) + '\u0001';
        const allQuiet = pts.every(function (pt) { return !!set[base + normKey(pt)]; });
        if (!allQuiet) continue;
        const text = cellText(cells[ci]);
        if (text.indexOf(NOISE_HAZARD) < 0) continue;   // 兜底：列对不上 / 已删过
        const next = dropNoiseItem(text);
        if (next === null) continue;
        const newRow = rows[ri].replace(cells[ci], setCellText(cells[ci], next));
        newTbl = newTbl.replace(rows[ri], newRow);
        changed++;
      }
      return newTbl;
    });
    return { xml: out, changed: changed };
  }

  // ---------------- 报告后处理②b：表4-1 职业病危害因素列「氮氧化物」措辞精简 ----------------
  // 用户要求（2026-10-16）：表4-1「职业病危害因素」列里若出现
  //   氮氧化物(一氧化氮和二氧化氮):氮氧化物(以二氧化氮计)
  // 则替换为
  //   氮氧化物(一氧化氮和二氧化氮)
  // （即去掉冒号及其后的「氮氧化物(以二氧化氮计)」补充说明）。
  // ⚠ 该短语可能① 独占整个单元格首段，也可能② 混在「…,氮氧化物(一氧化氮和二氧化氮):氮氧化物(以二氧化氮计),电焊烟尘,…」
  //    这种逗号分隔的清单里（真实 1009 报告即如此），故按**子串替换**处理，不做整格覆盖。
  // ⚠ 平台模板可能用**半角**「():」或**全角**「（）：」；两种都替换。
  const NOX_FROM_HALF = '氮氧化物(一氧化氮和二氧化氮):氮氧化物(以二氧化氮计)';
  const NOX_TO = '氮氧化物(一氧化氮和二氧化氮)';
  const NOX_FROM_FULL = '氮氧化物（一氧化氮和二氧化氮）：氮氧化物（以二氧化氮计）';
  const NOX_TO_FULL = '氮氧化物（一氧化氮和二氧化氮）';
  const NOX_FIXES = [
    { from: NOX_FROM_HALF, to: NOX_TO },
    { from: NOX_FROM_FULL, to: NOX_TO_FULL },
  ];

  // 针对表4-1 的「职业病危害因素」列（定位方式同 fixNonNoiseHazards：表头含该列名），
  // 对每个数据行的该列单元格做措辞替换（支持跨 run）。返回 { xml, changed }。
  function fixNitrogenOxide(xml) {
    if (xml.indexOf('氮氧化物') < 0) return { xml: xml, changed: 0 };
    let changed = 0;
    const out = mapTables(xml, function (tbl) {
      const rows = tbl.match(/<w:tr(?:\s[^>]*)?>[\s\S]*?<\/w:tr>/g) || [];
      if (rows.length < 2) return tbl;
      const headCells = rows[0].match(/<w:tc(?:\s[^>]*)?>[\s\S]*?<\/w:tc>/g) || [];
      let hIdx = -1;
      for (let i = 0; i < headCells.length; i++) {
        if (cellText(headCells[i]).indexOf(HAZARD_COL_HEAD) >= 0) { hIdx = i; break; }
      }
      if (hIdx < 0) return tbl;   // 不是表4-1
      let newTbl = tbl;
      for (let ri = 1; ri < rows.length; ri++) {
        const cells = rows[ri].match(/<w:tc(?:\s[^>]*)?>[\s\S]*?<\/w:tc>/g) || [];
        if (cells.length < 3) continue;
        const ci = hIdx + (cells.length - headCells.length);   // 合并单元格导致的列号平移
        if (ci < 0 || ci >= cells.length) continue;
        const cell = cells[ci];
        const text = cellText(cell);
        if (text.indexOf('氮氧化物') < 0) continue;            // 兜底：列对不上就不动
        let newCell = cell;
        for (const f of NOX_FIXES) {
          // 先试单元格内跨 run 替换（保留格式）；再退化为直接对该单元格 XML 做替换
          const np = replaceAcrossRuns(newCell, f.from, f.to);
          if (np !== null) { newCell = np; changed++; continue; }
          if (newCell.indexOf(f.from) >= 0) { newCell = newCell.split(f.from).join(f.to); changed++; }
        }
        if (newCell === cell) continue;
        const newRow = rows[ri].replace(cell, newCell);
        newTbl = newTbl.replace(rows[ri], newRow);
      }
      return newTbl;
    });
    return { xml: out, changed: changed };
  }

  // ---------------- 报告后处理⑧：表3-6 个人使用的职业病防护用品「同规格行合并」 ----------------
  // 用户要求（2026-10-20）：表3-6「个人使用的职业病防护用品配置及使用情况」里，若多行的
  //   「分类 / 类别 / 生产厂家 / 型号或规格 / 防护性能参数」五列内容**分别相同**，就把这些行合并成一行：
  //   · 「配置岗位（工种）」：把各行「/」**前**的部分（车间）提取在一起、去重后用「,」连接；
  //     「/」**后**的部分（岗位）同样提取、去重后用「,」连接；两者之间再用「/」连接
  //     （例：`车间A/岗位1` + `车间B/岗位2,岗位3` → `车间A,车间B/岗位1,岗位2,岗位3`）；
  //   · 其余列（佩戴情况 / 更换周期 / 备注…）：合并相同值去重（逗号连接）。
  // 实现：保留组内第一行并改写其单元格文本，组内其余行删除后按原顺序重建表体
  //   （表3-6 的数据行没有纵向合并，删除安全；表头两行的 vMerge/gridSpan 原样保留）。
  const PPE_KEY_HEADS = ['分类', '类别', '生产厂家', '型号', '防护性能'];
  const PPE_JOB_HEAD = '配置岗位';

  function splitListItems(s) {
    return String(s === null || s === undefined ? '' : s)
      .split(/[,，、;；]/)
      .map(function (x) { return x.trim(); })
      .filter(function (x) { return x !== ''; });
  }
  function pushUniq(arr, v) {
    const s = String(v === null || v === undefined ? '' : v).trim();
    if (s && arr.indexOf(s) < 0) arr.push(s);
  }
  // 合并「配置岗位（工种）」：前段（车间）/ 后段（岗位）各自去重后用「,」连接，中间用「/」
  function mergePpeJobText(values) {
    const before = [];
    const after = [];
    for (const v of values) {
      const s = String(v === null || v === undefined ? '' : v).trim();
      if (!s || isNoHazardText(s)) continue;              // 空 / 「-」/「/」视为无内容
      let i = s.indexOf('/');
      const j = s.indexOf('／');
      if (i < 0 || (j >= 0 && j < i)) i = j;
      if (i < 0) {
        splitListItems(s).forEach(function (x) { pushUniq(before, x); });
        continue;
      }
      splitListItems(s.slice(0, i)).forEach(function (x) { pushUniq(before, x); });
      splitListItems(s.slice(i + 1)).forEach(function (x) { pushUniq(after, x); });
    }
    if (before.length && after.length) return before.join(',') + '/' + after.join(',');
    return before.length ? before.join(',') : after.join(',');
  }
  // 其余列：合并相同值去重（占位值「-」「/」不参与合并，整列都是占位值时原样保留）
  function mergePpePlainText(values) {
    const out = [];
    let firstRaw = '';
    for (const v of values) {
      const s = String(v === null || v === undefined ? '' : v).trim();
      if (!s) continue;
      if (!firstRaw) firstRaw = s;
      if (isNoHazardText(s)) continue;
      if (out.indexOf(s) < 0) out.push(s);
    }
    return out.length ? out.join(',') : firstRaw;
  }

  // 返回 { xml, changed }；changed = 被合并掉的行数
  function mergePpeRows(xml) {
    if (xml.indexOf(PPE_JOB_HEAD) < 0) return { xml: xml, changed: 0 };
    let changed = 0;
    const out = mapTables(xml, function (tbl) {
      // 只处理表3-6：表头必含「生产厂家」「型号」「防护性能」与「配置岗位」
      if (tbl.indexOf('生产厂家') < 0 || tbl.indexOf('防护性能') < 0) return tbl;
      const rows = tbl.match(/<w:tr(?:\s[^>]*)?>[\s\S]*?<\/w:tr>/g) || [];
      if (rows.length < 3) return tbl;
      const head0 = rowCells(rows[0]);
      const head1 = rowCells(rows[1]);
      // 数据行的列数（取各行最大值，容忍个别行少格）
      let nCol = Math.max(head0.length, head1.length);
      const dataStart = (cellText(head1[0] || '').indexOf('分类') >= 0
        || cellText(head1[1] || '').indexOf('类别') >= 0) ? 2 : 1;
      for (let ri = dataStart; ri < rows.length; ri++) {
        nCol = Math.max(nCol, rowCells(rows[ri]).length);
      }
      // 列名：子表头（分类/类别）优先，其余按表头的 gridSpan 占槽位补空
      const names = new Array(nCol).fill('');
      for (let i = 0; i < head1.length && i < nCol; i++) {
        const t = cellText(head1[i]);
        if (t && !names[i]) names[i] = t;
      }
      let p = 0;
      for (const c of head0) {
        const spanM = /<w:gridSpan w:val="(\d+)"/.exec(c);
        const span = spanM ? parseInt(spanM[1], 10) : 1;
        const t = cellText(c);
        for (let k = 0; k < span; k++) {
          if (p + k < nCol && !names[p + k] && t) names[p + k] = t;
        }
        p += span;
      }
      const findCol = function (kw) {
        const want = normKey(kw);
        for (let i = 0; i < names.length; i++) if (normKey(names[i]).indexOf(want) >= 0) return i;
        return -1;
      };
      const iJob = findCol(PPE_JOB_HEAD);
      const keyIdx = PPE_KEY_HEADS.map(findCol);
      if (iJob < 0 || keyIdx.some(function (i) { return i < 0; })) return tbl;

      // 按「五列关键列」归组（保持首次出现顺序）
      const groups = [];
      const map = new Map();
      for (let ri = dataStart; ri < rows.length; ri++) {
        const cs = rowCells(rows[ri]);
        if (cs.length !== nCol) continue;                 // 非标准行（含合并格）不参与
        const texts = cs.map(cellText);
        const key = keyIdx.map(function (i) { return normKey(texts[i]); }).join('\u0001');
        let g = map.get(key);
        if (!g) { g = { rows: [], texts: [] }; map.set(key, g); groups.push(g); }
        g.rows.push(ri);
        g.texts.push(texts);
      }

      const drop = new Set();
      const keep = new Map();                             // ri -> 改写后的行 XML
      for (const g of groups) {
        if (g.rows.length < 2) continue;
        const first = g.rows[0];
        const firstXml = rows[first];
        const cells = rowCells(firstXml);
        let newXml = firstXml;
        for (let ci = 0; ci < nCol && ci < cells.length; ci++) {
          if (keyIdx.indexOf(ci) >= 0) continue;           // 关键列本来就相同，不动
          const col = g.texts.map(function (t) { return t[ci]; });
          const val = (ci === iJob) ? mergePpeJobText(col) : mergePpePlainText(col);
          if (!val || val === cellText(cells[ci])) continue;
          newXml = newXml.replace(cells[ci], setCellText(cells[ci], xmlEsc(val)));
        }
        if (newXml !== firstXml) keep.set(first, newXml);
        for (let k = 1; k < g.rows.length; k++) drop.add(g.rows[k]);
        changed += g.rows.length - 1;
      }
      if (!drop.size) return tbl;
      // 按原顺序重建表体（表头/表格属性/tblGrid 原样保留）
      const firstRowXml = rows[0];
      const lastRowXml = rows[rows.length - 1];
      const headXml = tbl.slice(0, tbl.indexOf(firstRowXml));
      const tailXml = tbl.slice(tbl.lastIndexOf(lastRowXml) + lastRowXml.length);
      let body = '';
      for (let ri = 0; ri < rows.length; ri++) {
        if (drop.has(ri)) continue;
        body += keep.has(ri) ? keep.get(ri) : rows[ri];
      }
      return headXml + body + tailXml;
    });
    return { xml: out, changed: changed };
  }

  // ---------------- 报告后处理③：正文措辞替换 ----------------
  // 平台模板里的几处说法要改成规范表述（用户要求，2026-10-12）。
  // ⚠ 目标文字常被拆到多个 run（如「7.2物理有害因素检测结果分析与评价」= ['7.2','物理','有害因素…']），
  //   因此按「段落内跨 run」替换，并尽量保留原有格式（替换文字落在起始 run 上）。
  const TEXT_FIXES = [
    { from: '工作场所粉尘和化学有害因素职业接触限值', to: '工作场所空气中化学有害因素职业接触限值' },
    { from: '工作场所空气中粉尘及化学有害因素职业接触限值', to: '工作场所空气中化学有害因素职业接触限值' },
    { from: '物理有害因素检测结果分析与评价', to: '物理因素检测结果分析与评价' },
  ];

  // 在**单个段落** XML 内做跨 run 的文本替换；命中返回新 XML，未命中返回 null。
  // ★ 统一实现：真正的切片逻辑只有 `spliceParaRange` 一份，这里只是「按子串定位 → 调它」。
  function replaceAcrossRuns(pXml, from, to) {
    if (!from) return null;
    const full = paraTextConcat(pXml);
    const idx = full.indexOf(from);
    if (idx < 0) return null;
    // keepTag=true：保留原 <w:t> 的属性（与原实现逐字节一致）
    return spliceParaRange(pXml, idx, idx + from.length, to, true);
  }

  // 对 document.xml 逐段落做全部文字替换，返回 { xml, changed }
  // ★ 性能：① 每段只解析一次文本（原来 3 条 TEXT_FIXES 各解析一次）；
  //        ② 只扫「含粉 / 含害」的段落（TEXT_FIXES 每条都至少含其中一个字），不再全篇扫。
  //   ⚠ 以后往 TEXT_FIXES 加条目时，必须保证新条目的 from 里含 TEXT_FIX_NEEDLES 中的某个字，
  //     否则该条永远匹配不到（下面有启动期自检，会打印警告）。
  // 定位字（只做「哪些段落可能命中」的廉价预筛，真正的匹配仍是整串比对）：
  //   2026-10-21 性能：原来用 ['粉','害']，这两个字在报告里极常见（上千段都要做一次 paraTextConcat），
  //   实测 68ms；改用 ['限','价']（三条替换串都含其一：两条含「接触限值」、一条含「分析与评价」）
  //   后候选段落骤减到个位数。启动期自检会校验「每条 from 至少含一个定位字」。
  const TEXT_FIX_NEEDLES = ['限', '价'];
  TEXT_FIXES.forEach(function (f) {
    const ok = TEXT_FIX_NEEDLES.some(function (n) { return f.from.indexOf(n) >= 0; });
    if (!ok) console.warn('[检评报告] TEXT_FIXES 条目缺少定位字（会被漏掉）：' + f.from);
  });
  function fixReportText(xml) {
    let changed = 0;
    const out = mapParasContaining(xml, TEXT_FIX_NEEDLES, function (pXml) {
      if (pXml.indexOf('<w:t') < 0) return pXml;
      let cur = pXml;
      let full = null;
      for (const f of TEXT_FIXES) {
        if (full === null) full = paraTextConcat(cur);
        const i = full.indexOf(f.from);
        if (i < 0) continue;
        cur = spliceParaRange(cur, i, i + f.from.length, f.to, true);   // keepTag：保留原 <w:t> 属性
        full = null;                                                    // 文本已变 → 下次重算
        changed++;
      }
      return cur;
    });
    return { xml: out, changed: changed };
  }

  // ---------------- 报告后处理④：表1-1 的「注3」按备注列实际标记裁剪 ----------------
  // 表1-1（化学有害因素接触限值表）的「备注」列里没出现某个标记时，把注3 里对应的说明删掉，
  // 并把该说明**前面**的「，」/「；」改成「。」（因为后面那句已经被删了）。
  const NOTE3_MARKS = [
    { mark: 'G1', text: '“G1”表示确认人类致癌物。' },
    { mark: 'G2A', text: '“G2A”表示可能人类致癌物，' },
    { mark: 'G2B', text: '“G2B”表示可疑人类致癌物；' },
    { mark: '皮', text: '“皮”表示可因皮肤、黏膜和眼睛直接接触蒸气、液体和固体，通过完整皮肤吸收引起全身效应；' },
    { mark: '敏', text: '“敏”是指已被人或动物资料证实该物质可能有致敏作用。' },
  ];

  // 收集表1-1「备注」列里出现过的标记；找不到该表返回 null
  function collectHazardMarks(xml) {
    let found = null;
    // ★ 性能（2026-10-18）：原来用 mapTables(xml, cb) —— 那是 xml.replace 全量重建一个 6MB 字符串，
    //   而这里**根本不需要重建**（只读不改）。改为 tablesOf 逐表只读遍历；
    //   再加单字预筛：目标表头必含「临界不良健康效应」和「备注」→ 原始 XML 必含「临」「备」。
    for (const tbl of tablesOf(xml)) {
      if (found) break;
      if (tbl.indexOf('临') < 0 || tbl.indexOf('备') < 0) continue;
      const rows = tbl.match(/<w:tr(?:\s[^>]*)?>[\s\S]*?<\/w:tr>/g) || [];
      if (rows.length < 2) continue;
      const headCells = rows[0].match(/<w:tc(?:\s[^>]*)?>[\s\S]*?<\/w:tc>/g) || [];
      const heads = headCells.map(cellText);
      const has = function (kw) { return heads.some(function (h) { return h.indexOf(kw) >= 0; }); };
      const hIdx = heads.findIndex(function (h) { return h.indexOf('备注') >= 0; });
      // 化学有害因素接触限值表：表头同时含「中文名」「临界不良健康效应」「备注」
      if (hIdx < 0 || !has('中文名') || !has('临界不良健康效应')) continue;
      found = {};
      for (let ri = 1; ri < rows.length; ri++) {
        const cells = rows[ri].match(/<w:tc(?:\s[^>]*)?>[\s\S]*?<\/w:tc>/g) || [];
        // 表头可能因合并单元格比数据行少若干格，按差值平移
        const ci = hIdx + (cells.length - headCells.length);
        if (ci < 0 || ci >= cells.length) continue;
        const v = cellText(cells[ci]);
        NOTE3_MARKS.forEach(function (it) { if (v.indexOf(it.mark) >= 0) found[it.mark] = 1; });
      }
    }
    return found;
  }

  // 按实际标记裁剪注3 文本；无改动返回 null
  function trimNote3(noteText, marks) {
    const spans = [];
    NOTE3_MARKS.forEach(function (it) {
      if (marks[it.mark]) return;              // 有该标记 → 保留
      const i = noteText.indexOf(it.text);
      if (i < 0) return;                       // 报告里本来就没有这句
      spans.push({ s: i, e: i + it.text.length });
    });
    if (!spans.length) return null;
    const del = {};
    const starts = {};
    spans.forEach(function (sp) {
      starts[sp.s] = 1;
      for (let k = sp.s; k < sp.e; k++) del[k] = 1;
    });
    let out = '';
    for (let i = 0; i < noteText.length; i++) {
      if (del[i]) continue;
      let ch = noteText[i];
      // 紧邻被删片段之前的「，」「；」→ 改成「。」（「：」等前缀标点不动）
      if (starts[i + 1] && (ch === '，' || ch === '；')) ch = '。';
      out += ch;
    }
    return out;
  }

  // 注文是否还剩内容（去掉「注N：」前缀后是否为空）
  function noteHasContent(t) {
    const i = Math.max(t.indexOf('：'), t.indexOf(':'));
    const body = i >= 0 ? t.slice(i + 1) : t;
    return body.replace(/\s+/g, '') !== '';
  }

  // 对 document.xml 处理注3，返回 { xml, changed }
  // 若注3 被整段删除，则把它**后面**的注号依次前移（注4 → 注3 …），前面的不动。
  //
  // ★ 性能：分两段式处理，避免「注」字把大量「备注」单元格也拉进来：
  //   ① 先只扫含「癌」的段落（注3 那句含「人类致癌物」）——通常 0~1 个段落；
  //   ② 只有真的整段删掉了注3，才再扫一遍含「注」的段落做注号前移。
  function fixHazardNote(xml) {
    const marks = collectHazardMarks(xml);
    if (!marks) return { xml: xml, changed: 0 };
    let changed = 0;
    let deleted = false;   // 是否已整段删除注3
    let removedNo = 0;     // 被删的注号

    const out1 = mapParasContaining(xml, ['癌'], function (pXml) {
      const full = cellText(pXml);
      if (full.indexOf('人类致癌物') < 0) return pXml;       // 只处理注3 那句
      const next = trimNote3(full, marks);
      if (next === null || next === full) return pXml;
      changed++;
      // 说明全被删掉、只剩「注3：」时，整段删除（避免报告里留下空注）
      if (!noteHasContent(next)) {
        const m = /^注(\d+)\s*[：:]/.exec(full);
        removedNo = m ? parseInt(m[1], 10) : 0;
        deleted = true;
        return '';
      }
      return setCellText(pXml, next);          // 文本写回首个 <w:t>，其余清空（保留段落格式）
    });
    if (!deleted || !removedNo) return { xml: out1, changed: changed };

    const out2 = mapParasContaining(out1, ['注'], function (pXml) {
      const full = cellText(pXml);
      const m = /^(注\d+)(\s*[：:])/.exec(full);
      if (!m) return pXml;
      const n = parseInt(m[1].slice(1), 10);
      if (!(n > removedNo)) return pXml;                  // 前面的注号不动
      const np = replaceAcrossRuns(pXml, m[0], '注' + (n - 1) + m[2]);
      if (np === null) return pXml;
      changed++;
      return np;
    });
    return { xml: out2, changed: changed };
  }

  // ---------------- 报告后处理⑤：表4-1 下方的「注：」段落 + 表3-1 人数联动 ----------------
  // 需求（用户 2026-10-16，已二次细化）：
  //   ① 表4-1（岗位（工种）设置、劳动者作业情况及职业病危害因素识别情况表）下方**相邻**的
  //      含「注：」的段落，整段替换为固定文本（两处人数为变量）：
  //        注：用人单位职工总数<在册职工人数>，接触职业病危害人数<接触人数>。打磨岗位实行两班制，
  //        上班时间为白班8:00～20:00、晚班20:00～8:00（次日），其余各岗位均实行白班制，
  //        上班时间为8:00～11:30、12:30～17:00。M车间A岗位、N车间B岗位接触的噪声经过验证性检测，
  //        当天工况噪声检测结果小于80dB(A)，根据《职业卫生名词术语》（GBZ/T224-2010）定义属于
  //        非噪声作业岗位（仅代表当天工况），故噪声不作为上述岗位的主要职业病危害因素。
  //   ② 「接触人数」=「表8-1 检测结果与评价结论」作业人数列**按单元+岗位去重**后的合计；
  //      同时把该合计回填到表3-1「接触职业病危害人数」右边单元格（保留「人」后缀）。
  //   ③ 「在册职工人数」取表3-1「在册职工人数」右边单元格的值（如「410人」）。
  //   ④（2026-10-16 追加）注段落拆成**两段**：「M车间A岗位」前另起一段；
  //      「M车间A岗位、N车间B岗位」替换为「噪声数据结果」页签下「非噪声车间岗位」筛选结果
  //      （格式：车间岗位、岗位；车间岗位）；
  //      若该筛选结果为空 → **整段（第二段）删除**（不再输出）。
  // 结果：两段（第二段可能不存在）：
  //   A: 注：用人单位职工总数{STAFF}，接触职业病危害人数{CONTACT}。打磨岗位实行两班制…12:30～17:00。
  //   B: {NON_NOISE_TEXT}接触的噪声经过验证性检测…故噪声不作为上述岗位的主要职业病危害因素。

  const NOTE_PARA_A =
    '注：用人单位职工总数{STAFF}，接触职业病危害人数{CONTACT}。打磨岗位实行两班制，'
    + '上班时间为白班8:00～20:00、晚班20:00～8:00（次日），其余各岗位均实行白班制，'
    + '上班时间为8:00～11:30、12:30～17:00。';
  const NOTE_PARA_B_TPL =
    '{JOBS}接触的噪声经过验证性检测，当天工况噪声检测结果小于80dB(A)，'
    + '根据《职业卫生名词术语》（GBZ/T224-2010）定义属于非噪声作业岗位（仅代表当天工况），'
    + '故噪声不作为上述岗位的主要职业病危害因素。';
  // 兼容旧名（单段模板，保留供外部/测试参考，不再直接使用）
  const NOTE_TPL = NOTE_PARA_A + NOTE_PARA_B_TPL;

  // 取一行里的单元格 XML 数组（保持书写顺序）
  function rowCells(trXml) {
    return trXml.match(/<w:tc(?:\s[^>]*)?>[\s\S]*?<\/w:tc>/g) || [];
  }
  function tableRows(tblXml) {
    return tblXml.match(/<w:tr(?:\s[^>]*)?>[\s\S]*?<\/w:tr>/g) || [];
  }
  // XML 文本转义（写入 <w:t> 前必须转义 & < >）
  function xmlEsc(s) {
    return String(s === null || s === undefined ? '' : s)
      .replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
  }
  // 把值补成「N人」（若原值已含「人」则不再叠加；取不到数字则原样返回）
  function withPersonUnit(v) {
    const s = String(v === null || v === undefined ? '' : v).trim();
    if (!s) return s;
    if (/人\s*$/.test(s)) return s;
    return s + '人';
  }
  function pickNumber(s) {
    const m = String(s === null || s === undefined ? '' : s).match(/\d+(?:\.\d+)?/);
    return m ? m[0] : '';
  }

  // 在 document.xml 里找**表头含指定列名**的表，返回其行/表头单元格等信息（找不到返回 null）
  //
  // ★ 性能（2026-10-18）：新增可选第 3 参 needChars —— 一组**单字**，必须**全部**原样出现在
  //   该表的原始 XML 里才继续解析表头。用于在「扫表头」之前廉价跳过无关表（单字不会被 run 拆开，
  //   判定安全）。不传则保持原行为（全表扫描）。
  function findTableByHead(xml, predicate, needChars) {
    const tables = tablesOf(xml);
    for (const tbl of tables) {
      if (needChars) {
        let miss = false;
        for (const ch of needChars) { if (tbl.indexOf(ch) < 0) { miss = true; break; } }
        if (miss) continue;
      }
      const rows = tableRows(tbl);
      if (rows.length < 2) continue;
      const head = rowCells(rows[0]);
      const heads = head.map(cellText);
      if (predicate(heads, tbl, rows)) return { tbl: tbl, rows: rows, head: head, heads: heads };
    }
    return null;
  }

  // 找表头（可能两行合并）里某列的下标；先精确、再包含
  function headIndexOf(heads, name) {
    const want = normKey(name);
    for (let i = 0; i < heads.length; i++) if (normKey(heads[i]) === want) return i;
    for (let i = 0; i < heads.length; i++) if (normKey(heads[i]).indexOf(want) >= 0) return i;
    return -1;
  }

  // 表3-1 用人单位基本情况表：该表**不是**「表头 + 数据行」结构，而是「标签 | 值 | 标签 | 值」
  // 的键值对表格。因此不按表头定位，而是**逐行扫所有单元格**找标签，取同行右邻单元格作为值。
  // 返回 { staffText, contactText, staffLabelCell, contactLabelCell, tbl, rows }（找不到返回 null）
  //
  // ★ 性能（2026-10-18）：原实现会遍历**全文每一张表的每一个单元格**（真实报告 22 张表 ≈ 6400 个
  //   单元格 → 单次 130ms+，是整条后处理管线最贵的一步）。现在先做一次**单字预筛**：
  //   「在册职工人数」必含「册」、「接触职业病危害人数」必含「触」，且两标签**必须同表**才可能命中，
  //   因此不含这两个字的表可以直接跳过。单字预筛是安全的——汉字即使被拆到多个 <w:t> 里，
  //   该字本身仍原样出现在原始 XML 中（同 mapParasContaining 的 needles 约定）。
  //   实测 22 张表 → 1 张，步耗时 139ms → 4ms。
  function readBasicInfoTable(xml) {
    const tables = tablesOf(xml);
    const K_STAFF = normKey('在册职工人数');
    const K_CONTACT = normKey('接触职业病危害人数');
    for (const tbl of tables) {
      if (tbl.indexOf('册') < 0 || tbl.indexOf('触') < 0) continue;   // 预筛：两标签不可能同表
      const rows = tableRows(tbl);
      if (rows.length < 2) continue;
      let staffText = null, contactText = null, staffCi = -1, contactCi = -1, staffRow = -1, contactRow = -1;
      for (let ri = 0; ri < rows.length; ri++) {
        const texts = rowCells(rows[ri]).map(cellText);
        for (let ci = 0; ci < texts.length; ci++) {
          const k = normKey(texts[ci]);
          if (staffText === null && k === K_STAFF) {
            staffText = texts[ci + 1] || ''; staffCi = ci; staffRow = ri;
          } else if (contactText === null && k === K_CONTACT) {
            contactText = texts[ci + 1] || ''; contactCi = ci; contactRow = ri;
          }
        }
        if (staffText !== null && contactText !== null) break;
      }
      if (staffText !== null && contactText !== null) {
        return {
          tbl: tbl, rows: rows,
          staffCi: staffCi, contactCi: contactCi,
          staffRow: staffRow, contactRow: contactRow,
          staffText: staffText, contactText: contactText,
        };
      }
    }
    return null;
  }

  // ---------------- 接触职业病危害人数：从表4-1「人数 / 总数」列统计（2026-10-19 口径变更） ----------------
  // 用户要求：接触职业病危害人数 = **表4-1「岗位（工种）设置、劳动者作业情况及职业病危害因素识别情况」**
  //   「人数」列下「总数」子列的数字合计；某个「人数」单元格所对应的行（同一岗位的多个点位行、
  //   含纵向合并的续行）里「职业病危害因素」列**全部为空或「-」**时，该人数不纳入统计。
  // ⚠ 关键（1121 样例实测）：同一岗位的多个点位行会把**同一个人数**重复写在每行
  //   （包装工 20 / 20 / 20、压机工 30×4…），所以必须**按「车间+岗位」去重后只计一次**，
  //   否则会把人数成倍放大（1121 逐格相加得 231，按岗位去重才是 77，与报告表3-1 的 77 人一致）。
  // ⚠ 必须在 ②/②c 删完噪声之后执行：被删空的单元格已填「-」，自然被排除在外。
  // ⚠ 表4-1 找不到（或没有人数列）时返回 null，调用方退回旧的表8-1 去重合计（保守兜底）。
  const NO_HAZARD_TEXTS = ['', '-', '—', '–', '/', '／', '无'];
  function isNoHazardText(t) {
    const k = normKey(t);
    for (const p of NO_HAZARD_TEXTS) if (k === normKey(p)) return true;
    return false;
  }

  // 返回 { total, groups, counted, skipped } 或 null
  //   · groups ：表4-1 里不同「车间+岗位」的个数（= 人数单元格去重后的个数）
  //   · counted：实际纳入统计的岗位数
  //   · skipped：因危害因素列全为空/「-」被整组排除的岗位数
  function sumExposedFromT41(xml) {
    const t = findTableByHead(xml, function (heads) {
      const has = function (n) {
        const want = normKey(n);
        return heads.some(function (h) { return normKey(h).indexOf(want) >= 0; });
      };
      return has('职业病危害因素') && has('人数') && has('岗位');
    }, ['岗', '数', '危']);
    if (!t) return null;
    const hNum = headIndexOf(t.heads, '人数');
    const hHaz = headIndexOf(t.heads, '职业病危害因素');
    if (hNum < 0 || hHaz < 0) return null;
    // 「人数」在数据行里被拆成「总数 / 数/班」两格：先按表头下方那行子表头定位「总数」列；
    // 找不到子表头时退回表头「人数」列号（多出来的那格在人数列之后，故列号不变）。
    let cTotal = hNum;
    if (t.rows.length > 1) {
      const sub = rowCells(t.rows[1]).map(cellText);
      if (!normKey(sub[0]) && !normKey(sub[1])) {          // 只有子表头行才前两格为空
        for (let i = 0; i < sub.length; i++) {
          if (normKey(sub[i]) === normKey('总数')) { cTotal = i; break; }
        }
      }
    }
    // 按「车间+岗位」归组：同一岗位的多行（不同点位 / 纵向合并续行）只保留一个人数，
    // 但把**每一行**的危害因素都收进该组，供「整组都为 -/空」判定。
    const map = new Map();
    const order = [];
    for (let ri = 1; ri < t.rows.length; ri++) {
      const cs = rowCells(t.rows[ri]);
      if (cs.length < 3) continue;
      const unit = cellText(cs[0]);
      const job = cellText(cs[1]);
      if (!unit && !job) continue;                        // 子表头 / 空行
      const kUnit = normKey(unit);
      const kJob = normKey(job);
      if (kUnit.indexOf('合计') >= 0 || kJob.indexOf('合计') >= 0) continue;
      const hi = hHaz + (cs.length - t.head.length);      // 合并单元格导致的列号平移（同②）
      const hazard = (hi >= 0 && hi < cs.length) ? cellText(cs[hi]) : '';
      const numText = (cTotal >= 0 && cTotal < cs.length) ? cellText(cs[cTotal]) : '';
      const num = parseFloat(pickNumber(numText));
      const key = kUnit + '\u0001' + kJob;
      let g = map.get(key);
      if (numText !== '' && isFinite(num)) {
        if (!g) { g = { num: num, hazards: [] }; map.set(key, g); order.push(key); }
        g.hazards.push(hazard);
      } else if (g) {
        g.hazards.push(hazard);                           // 人数格为空（纵向合并/续行）→ 沿用该岗位的人数
      }
    }
    if (!order.length) return null;
    let total = 0, counted = 0, skipped = 0;
    for (const key of order) {
      const g = map.get(key);
      let hasHaz = false;
      for (const h of g.hazards) { if (!isNoHazardText(h)) { hasHaz = true; break; } }
      if (!hasHaz) { skipped++; continue; }               // 整组危害因素都是「-」/空 → 不纳入统计
      total += g.num;
      counted++;
    }
    return { total: total, groups: order.length, counted: counted, skipped: skipped };
  }

  // 表8-1：作业人数列按「单元/工作场所 + 岗位/工种」去重后求和；返回 { total, groups } 或 null
  // ★ 性能：预筛单字「数」（作业人数）与「论」（评价结论），跳过无关表。
  function sumExposedWorkers(xml) {
    const t = findTableByHead(xml, function (heads) {
      return heads.some(function (h) { return normKey(h).indexOf('作业人数') >= 0; })
        && heads.some(function (h) { return normKey(h).indexOf('评价结论') >= 0; });
    }, ['数', '论']);
    if (!t) return null;
    const cUnit = headIndexOf(t.heads, '单元/工作场所');
    const cJob = headIndexOf(t.heads, '岗位/工种');
    const cNum = headIndexOf(t.heads, '作业人数');
    if (cUnit < 0 || cJob < 0 || cNum < 0) return null;
    // 表头可能因合并单元格比数据行少若干格，按差值平移
    const shift = function (cells, ci) { return ci + (cells.length - t.head.length); };
    const seen = Object.create(null);
    for (let ri = 1; ri < t.rows.length; ri++) {
      const cs = rowCells(t.rows[ri]);
      if (cs.length < 3) continue;
      const ui = shift(cs, cUnit), ji = shift(cs, cJob), ni = shift(cs, cNum);
      if (ui < 0 || ji < 0 || ni < 0 || ui >= cs.length || ji >= cs.length || ni >= cs.length) continue;
      const unit = cellText(cs[ui]);
      const job = cellText(cs[ji]);
      if (!unit && !job) continue;                       // 表头续行
      if (normKey(unit).indexOf('合计') >= 0 || normKey(job).indexOf('合计') >= 0) continue;
      const num = parseFloat(pickNumber(cellText(cs[ni])));
      if (!isFinite(num)) continue;
      const key = normKey(unit) + '\u0001' + normKey(job);
      if (!(key in seen)) seen[key] = num;               // 首次出现记录（去重：同岗位多检测项目只算一次）
    }
    const groups = Object.keys(seen);
    const total = groups.reduce(function (n, k) { return n + seen[k]; }, 0);
    return { total: total, groups: groups.length };
  }

  // 找「表4-1 岗位（工种）设置…」下方**相邻**的含「注：」的段落，用段落数组 texts 整体替换。
  // texts.length === 1 → 原地替换该段落；texts.length === 2 → 用第 1 段替换原段落，克隆其
  // <w:pPr> 生成第 2 段紧随其后；texts.length === 0 → 删除该段落（**仅在 opts.allowDelete=true 时**）。
  // 返回 { xml, changed }。identify：表头同时含「岗位（工种）」与「职业病危害因素」。
  //
  // ⚠ 关键：只看**顶层** <w:p>（不在 <w:tbl> 里的），否则会误匹配到表格单元格里的「注：」段落，
  //    导致替换/删除到错误位置（曾出现「表4-1 下方注段落被整段删掉」的实测现象）。
  // 做法：把 xml 按「表格」切成块，只保留每块里**表格之外**的文本，再在其中找「注：」段落。
  //
  // ⚠ 安全兜底：默认**绝不删除**段落（arr 为空时原样返回）。删除目标注段落本身属于误操作，
  //    我们只允许「替换成内容更少的段落」，永远不允许把这段整体抹掉。
  function fixNoteParagraph(xml, texts, opts) {
    const allowDelete = !!(opts && opts.allowDelete === true);
    const arr = Array.isArray(texts) ? texts.filter(function (t) { return t !== null && t !== undefined && t !== ''; })
      : [texts];
    // 空数组 = 删除整段；默认禁止（防误删），只有显式 allowDelete 才放行
    if (arr.length === 0 && !allowDelete) return { xml: xml, changed: 0 };
    const tables = tablesOf(xml);
    let target = null;
    for (const tbl of tables) {
      // 性能预筛：目标表头必含「职业病危害因素」→ 原始 XML 必含「职」字
      if (tbl.indexOf('职') < 0) continue;
      const rows = tableRows(tbl);
      if (!rows.length) continue;
      const heads = rowCells(rows[0]).map(cellText).map(normKey);
      const has = function (kw) { return heads.some(function (h) { return h.indexOf(kw) >= 0; }); };
      if (has('岗位（工种）') && has('职业病危害因素')) { target = tbl; break; }
    }
    if (!target) return { xml: xml, changed: 0 };
    // 表结束位置之后的文本；逐段扫描，且**跳过表格内部**的 <w:p>（表内段落不参与）
    const idx = xml.indexOf(target) + target.length;
    // 取表之后到「下一个 <w:tbl> 之前」的**非表格区**；若其间没有目标注段落，再继续往后找
    function findNotePara(startPos) {
      let pos = startPos;
      while (pos < xml.length) {
        const nextTbl = xml.indexOf('<w:tbl', pos);
        const segEnd = nextTbl < 0 ? xml.length : nextTbl;
        const seg = xml.slice(pos, segEnd);
        const re = new RegExp(PARA_SRC, 'g');
        let m;
        while ((m = re.exec(seg))) {
          const t = cellText(m[0]);
          if (t && t.indexOf('注：') >= 0) return { absIndex: pos + m.index, para: m[0] };
          // 段落后遇到数字标题（如「5检测项目确定」）→ 该表后无注段落，停止
          if (t && /^\s*\d/.test(t)) return null;
        }
        if (nextTbl < 0) return null;
        // 跳过这张表，继续往后（注段落可能在别处，但仅限相邻几张）
        const closeTbl = xml.indexOf('</w:tbl>', nextTbl);
        if (closeTbl < 0) return null;
        pos = closeTbl + '</w:tbl>'.length;
      }
      return null;
    }
    const found = findNotePara(idx);
    if (!found) return { xml: xml, changed: 0 };
    const { absIndex, para: oldP } = found;
    // 安全约束：arr 为空时已在函数入口拦截（除非 allowDelete），此处 arr 必非空 → 一定会写出 A 段
    let repl;
    if (arr.length === 0) {
      repl = '';                                        // 仅在 allowDelete=true 时到达
    } else {
      // 克隆原段落的 <w:pPr>（保留字体/缩进/行距等格式），第 2 段紧随其后。
      // ⚠ 第 2 段是**新建**的 <w:p>，必须同时带上原段落首个 <w:r> 的 <w:rPr>，
      //   否则新 run 无字符格式，会回落到样式默认值（字体/字号/东亚字体 hint 与第 1 段不一致）。
      const pPrM = /<w:pPr(?:\s[^>]*)?>[\s\S]*?<\/w:pPr>/.exec(oldP);
      const pPr = pPrM ? pPrM[0] : '';
      // 取原段落第一个 <w:r> 自身的 <w:rPr>（先剔除 <w:pPr>，避免命中其内嵌的段落标记 rPr）
      const pBody = pPr ? oldP.replace(pPr, '') : oldP;
      const runM = /<w:r(?:\s[^>]*)?>[\s\S]*?<\/w:r>/.exec(pBody);
      let rPr = '';
      if (runM) {
        const rPrM = /<w:rPr(?:\s[^>]*)?>[\s\S]*?<\/w:rPr>/.exec(runM[0]);
        if (rPrM) rPr = rPrM[0];
      }
      repl = setCellText(oldP, xmlEsc(arr[0]));
      for (let i = 1; i < arr.length; i++) {
        // 用同一 pPr，但去掉 paraId（避免 Word 里两段共用同一 paraId）
        const pPrClean = pPr.replace(/\sw14:paraId="[^"]*"/g, '');
        repl += '<w:p>' + pPrClean
          + '<w:r>' + rPr + '<w:t xml:space="preserve">' + xmlEsc(arr[i]) + '</w:t></w:r></w:p>';
      }
    }
    // 兜底自检：替换结果里若连 A 段首句都不见了，说明定位/替换异常 → 放弃本次改动（保原样）
    if (arr.length > 0 && repl.indexOf(xmlEsc(arr[0].slice(0, 6))) < 0) {
      console.warn('[检评报告] 注段落替换自检未通过，放弃本步（原段落保持不动）');
      return { xml: xml, changed: 0 };
    }
    const out = xml.slice(0, absIndex) + repl + xml.slice(absIndex + oldP.length);
    return { xml: out, changed: 1 };
  }

  // 把表3-1「接触职业病危害人数」右邻单元格改成 newText（保留格式，只改首段文本）
  function setBasicInfoContact(xml, basic, newText) {
    if (!basic || basic.contactRow < 0) return { xml: xml, changed: 0 };
    const r = basic.rows[basic.contactRow];
    const cs = rowCells(r);
    const label = cs[basic.contactCi];
    const target = cs[basic.contactCi + 1];
    if (label === undefined || target === undefined) return { xml: xml, changed: 0 };
    if (normKey(cellText(label)).indexOf('接触职业病危害人数') < 0) return { xml: xml, changed: 0 };
    const oldVal = cellText(target);
    if (oldVal === newText) return { xml: xml, changed: 0 };
    const newCell = setCellText(target, xmlEsc(newText));
    return { xml: xml.replace(target, newCell), changed: 1 };
  }

  // 整合：注段落替换（含拆两段/整段删除）+ 接触职业病危害人数联动（不依赖网络）
  // nonNoiseText：「非噪声车间岗位」筛选结果文本（车间岗位、岗位；车间岗位）；
  //   空串 → 第二段（M车间A岗位那段）整段删除。
  function fixNoteAndContactCount(xml, nonNoiseText) {
    const basic = readBasicInfoTable(xml);
    // ★ 2026-10-19 口径变更：接触职业病危害人数优先取**表4-1「人数/总数」列合计**
    //   （危害因素列全为「-」/空的行不纳入）；表4-1 找不到时退回表8-1 去重合计（兜底）。
    const t41 = sumExposedFromT41(xml);
    const sum = t41 || sumExposedWorkers(xml);
    if (!basic && !sum) return { xml: xml, changed: 0 };
    let changed = 0;
    let out = xml;
    // ① 表3-1「接触职业病危害人数」右邻 = 表8-1 去重合计
    if (sum && basic && basic.contactText !== '') {
      const contactVal = withPersonUnit(String(sum.total));
      const r = setBasicInfoContact(out, basic, contactVal);
      if (r.changed) { out = r.xml; changed++; }
    }
    // ② 注段落替换（第 1 段两处人数用实时值；第 2 段用非噪声岗位文本，为空则**只删第 2 段**）
    //    ★ A 段始终保留：texts 至少为 [paraA]，永远不传空数组（否则会整段删除）。
    const staffVal = basic && basic.staffText ? withPersonUnit(basic.staffText) : '';
    const contactVal2 = sum ? withPersonUnit(String(sum.total)) : '';
    const paraA = NOTE_PARA_A
      .replace('{STAFF}', staffVal || '0人')
      .replace('{CONTACT}', contactVal2 || '0人');
    const jobs = String(nonNoiseText === null || nonNoiseText === undefined ? '' : nonNoiseText).trim();
    const texts = jobs ? [paraA, NOTE_PARA_B_TPL.replace('{JOBS}', jobs)] : [paraA];
    const rn = fixNoteParagraph(out, texts);
    if (rn.changed) { out = rn.xml; changed++; }
    return {
      xml: out, changed: changed, staffVal: staffVal, contactVal: contactVal2,
      total: sum ? sum.total : null, jobs: jobs, noteTexts: texts,
      // 口径来源：t41 = 表4-1 人数/总数合计（新口径）；t81 = 表8-1 去重合计（兜底）
      totalFrom: t41 ? 't41' : (sum ? 't81' : 'none'),
      t41: t41,
    };
  }

  // ---------------- 报告后处理⑥：按「噪声 / 物化性质 / 存在高毒物品」条件删除条目 ----------------
  // 需求（用户 2026-10-17）：
  //   A. 「噪声数据结果」里**噪声车间岗位为空**（全项目无噪声岗位）时：
  //        · 删 1.2.3 评价依据 下的 《听力防护装备的选择、使用和维护》(GB 23466-2025)；
  //        · 删 9.2 职业卫生管理 下的「产生噪声的工作场所设置“噪声有害”“戴护耳器”等警示标识; 」
  //   B. 取「检测项目确认表」的「职业病危害因素」列每个单元格 → 按危害因素库「系统名称」列匹配 →
  //      读该行「物化性质」列；全部对应内容**都不含**关键词时按下表删除（含前面的序号）：
  //        · 无毒物（物化性质不含「毒物」）→ 删毒物组（有毒物品条例 / 高毒物品目录 / 防毒面具 / 混合性粉(烟)尘警示句）
  //        · 无粉尘（物化性质不含「粉尘」）→ 删粉尘组（尘肺病防治条例 / 防颗粒物呼吸器 / 产生粉尘警示句）
  //        · 粉尘、毒物都没有      → 在以上基础上再删 《呼吸防护装备的选择、使用和维护》(GB 18664-2025)；
  //   C. 同上取「存在高毒物品」列，全部都不是「是」→ 删 《高毒物品目录》（卫法监发（2003）142号）；
  //   D. 删除后，被删条目所属**章节**内的「（1）（2）…」序号**重新按顺序排列**。
  //
  // ⚠ 判定口径（用户确认）：
  //   · 关键词判定用**子串包含**（物化性质文本含「毒物」即视为有毒物、含「粉尘」即视为有粉尘）；
  //   · **缺数据不删**：某个危害因素在库中找不到（或物化性质为空）→ 视为「证据不足」，整组**不删**（保守，防误删）；
  //   · 条目定位用**全局文本匹配**（不限定章节），序号重排则**限定在所属章节内**。
  //
  // 设计：把「法规/警示句」按段落（<w:p>）整段删除——报告里每条独占一段，
  //   段落文本（去首尾空白）以目标串开头/包含目标串即删整段。
  //   序号重排：先记录每个删除段落所属章节（最近的上一个「标题段落」），删除后在每个受影响章节内
  //   把剩余的「（N）」按出现顺序重编为「（1）（2）…」。
  // ★ 条件删除目标（2026-10-21 按用户口径全面改写：粉尘/毒物/腐蚀/高温/紫外/噪声 逐条独立判定）
  //   每条目标的判定条件见 fixConditionalDeletions 里的编号注释。
  const PJBG_DEL_TARGETS = {
    // 1.1 法律、法规、规章
    dustLaw: '《中华人民共和国尘肺病防治条例》',                        // ① 无「粉尘」（「粉尘、毒物」不算）
    toxLaw: '《使用有毒物品作业场所劳动保护条例》',                     // ② 无「毒物」（「粉尘、毒物」不算）
    highToxList: '《高毒物品目录》',                                    // ③ 无高毒物品（判读方式同旧版）
    // 1.2.3 评价依据
    resp18664: '《呼吸防护装备的选择、使用和维护》(GB 18664-2025)',     // ④ 既无「毒物」也无「粉尘」
    hearing23466: '《听力防护装备的选择、使用和维护》(GB 23466-2025)',  // ⑤ 表4-1 危害因素列（处理后）无「噪声」
    mask2890: '《呼吸防护 自吸过滤式防毒面具》(GB 2890-2022)',         // ⑥ 无「毒物」（「粉尘、毒物」不算）
    resp2626: '《呼吸防护 自吸过滤式防颗粒物呼吸器》(GB 2626-2019)',   // ⑦ 无「粉尘」或「粉尘、毒物」
  };

  // ★ 分句裁剪关键词（9.2 第四段是一条长句，各危害各占一个分句，只能切分句不能整条删）
  const PJBG_TRIM_KWS = {
    dust: ['产生粉尘的工作场所设置', '对皮肤有刺激性或经皮肤吸收的粉尘工作场所'],  // ⑧ 无「粉尘」或「粉尘、毒物」
    mixed: ['产生含有有毒物质的混合性粉(烟)尘的工作场所应设置'],                  // ⑨ 无「粉尘、毒物」
    tox: ['有毒物品工作场所设置'],                                                // ⑩ 无「毒物」（「粉尘、毒物」不算）
    corrode: ['能引起职业性灼伤或腐蚀的化学品工作场所'],                          // ⑪ 无「腐蚀」
    noise: ['产生噪声的工作场所设置'],                                            // ⑫ 表4-1 危害因素列（处理后）无「噪声」
    heat: ['高温工作场所设置'],                                                   // ⑬ 无「高温」
    uv: ['产生紫外辐射的工作场所设置'],                                           // ⑭ 无「紫外辐射」
  };

  // 表4-1「岗位（工种）设置、劳动者作业情况及职业病危害因素识别情况」的
  // 「职业病危害因素」列里是否还出现「噪声」。
  //   ★ 必须在本步**之前**的 ②（非噪声岗位删「噪声」）/②c（低噪点位删「噪声」）跑完之后判断，
  //     即「经过处理后再看还剩不剩噪声」——patchPjbgDocx 的管线顺序已保证（⑥ 在 ② 之后）。
  //   返回 true / false；**找不到该表（或没有危害因素列）返回 null** → 调用方退回旧口径
  //   （噪声车间岗位是否为空，`noiseWorkshopEmpty`）。
  function t41HasNoise(xml) {
    // 表4-1 识别：表头同时含「职业病危害因素」「岗位」「人数」（同 sumExposedFromT41，
    //   避免误抓「岗位/工种/采样/测量点/对象 + 职业病危害因素 + 检测项目」那张检测表）
    const t = findTableByHead(xml, function (heads) {
      const has = function (n) {
        const want = normKey(n);
        return heads.some(function (h) { return normKey(h).indexOf(want) >= 0; });
      };
      return has('职业病危害因素') && has('岗位') && has('人数');
    }, ['岗', '数', '危']);
    if (!t) return null;
    const ci = headIndexOf(t.heads, '职业病危害因素');
    if (ci < 0) return null;
    for (let ri = 1; ri < t.rows.length; ri++) {
      const cs = rowCells(t.rows[ri]);
      if (cs.length < 3) continue;                        // 子表头行（总数 / 数/班）
      const idx = ci + (cs.length - t.head.length);       // 合并单元格导致的列号平移
      const txt = cellText(cs[idx >= 0 && idx < cs.length ? idx : ci]);
      if (hazNorm(txt).indexOf('噪声') >= 0) return true;
    }
    return false;
  }

  // 取段落可见文本（跨 run 拼接，去空白），用于关键词定位
  function paraText(pXml) { return cellText(pXml); }

  // 从 document.xml 里取「检测项目确认表」中指定列的每个单元格文本（按表头的列名定位）。
  // 返回 { found: 是否找到该表, texts: [单元格文本...] }；找不到表返回 { found:false, texts:[] }。
  //
  // ⚠⚠ 报告里有多张表都带「职业病危害因素」列（如「岗位/危害因素一览表」也有一列同名），
  //   早先「表头含该列名即命中」会**抓错表**（抓到一览表），那一列常出现 `-` 等占位值，
  //   导致后续按因素名查库失败 → 「缺数据不删」→ 整组永远不删（实测 1095/1009 均如此）。
  //   现在改为**两级优先**：
  //     ① 表头同时含「职业病危害因素」和「检测项目」→ 这才是真正的检测项目确认表（首选）；
  //     ② 没有①时才退回「只要含职业病危害因素」，并且优先**行数多**的那张。
  function collectConfirmColTexts(xml, colNames) {
    const out = { found: false, texts: [] };
    const tables = tablesOf(xml);
    let best = null;   // { score, tbl, ci, headCells }
    for (const tbl of tables) {
      // 性能预筛：必须有「职业病危害因素」列 → 该表原始 XML 必含「职」字（单字不会被 run 拆开）
      if (tbl.indexOf('职') < 0) continue;
      const rows = tableRows(tbl);
      if (rows.length < 2) continue;
      const headCells = rowCells(rows[0]).map(cellText);
      const hasHaz = headCells.some(function (h) { return h.indexOf('职业病危害因素') >= 0; });
      if (!hasHaz) continue;                       // 必须有「职业病危害因素」列
      const hasItem = headCells.some(function (h) { return h.indexOf('检测项目') >= 0; });
      let ci = -1;
      for (let i = 0; i < headCells.length; i++) {
        for (const nm of colNames) {
          if (headCells[i].indexOf(nm) >= 0) { ci = i; break; }
        }
        if (ci >= 0) break;
      }
      if (ci < 0) continue;
      // 首选（含「检测项目」列）权重远高于备选；同级再按行数多者优先
      const score = (hasItem ? 1e6 : 0) + rows.length;
      if (!best || score > best.score) best = { score: score, tbl: tbl, ci: ci, headCells: headCells, rows: rows };
    }
    if (!best) return out;
    out.found = true;
    for (let ri = 1; ri < best.rows.length; ri++) {
      const cells = rowCells(best.rows[ri]);
      if (cells.length < 2) continue;
      const shifted = best.ci + (cells.length - best.headCells.length);   // 合并单元格导致的列号平移
      const idx = (shifted >= 0 && shifted < cells.length) ? shifted : best.ci;
      const t = cellText(cells[idx]);
      if (t) out.texts.push(t);
    }
    return out;
  }

  // 占位值（`-`/`—`/`/`/`无` 等）表示「本行无危害因素」，**跳过**而不是当成查不到的因素，
  //   否则一张表里只要有一个 `-` 就会让整组判定失效（实测 1095 就是这样卡住的）。
  const HAZ_PLACEHOLDER = /^(?:[-—–－~\/／\\]+|无|空|N\/A|NA)$/i;
  function isHazPlaceholder(s) {
    const t = String(s === null || s === undefined ? '' : s).replace(/\s+/g, '');
    if (!t) return true;
    return HAZ_PLACEHOLDER.test(t);
  }
  // 归一化：去空白 + NFKC（全角→半角，解决报告写 `<10%` 而库写 `＜10%` 对不上的问题）
  function hazNorm(s) {
    let t = String(s === null || s === undefined ? '' : s).replace(/\s+/g, '');
    try { t = t.normalize('NFKC'); } catch (e) {}
    return t.replace(/\s+/g, '');
  }
  // 把危害因素库建成「归一化系统名称 → 行」的索引。**同一轮里只建一次**（见 fixConditionalDeletions），
  //   593 条 × 3 次字段查询 → 1 次，省掉两轮 NFKC 归一化。
  function buildHazardIndex(hf) {
    const idx = {};
    if (!Array.isArray(hf)) return idx;
    for (const h of hf) { const k = hazNorm(h && h.name); if (k) idx[k] = h; }
    return idx;
  }
  // 确认表单元格去重（同一因素常在多行重复出现，例如 1095 的 19 格只有 7 种）：
  //   按「归一化后的整格文本」去重，顺序保持首次出现。
  function uniqHazTexts(texts) {
    const seen = new Set();
    const out = [];
    for (const raw of (texts || [])) {
      const k = hazNorm(raw);
      if (!k || seen.has(k)) continue;
      seen.add(k); out.push(raw);
    }
    return out;
  }

  // 依据确认表的危害因素名，**一次性**匹配出危害因素库对应的行（多字段共用同一批结果）。
  // 匹配口径：危害因素库「系统名称」列（字段 name）。
  // ⚠ 占位值（`-`/`—`/`/`/`无` 等）表示「本行无危害因素」，**跳过**而不是当成查不到的因素，
  //   否则一张表里只要有一个 `-` 就会让整组判定失效（实测 1095 就是这样卡住的）。
  // 返回 { rows: [库行...], missing: [没查到的因素] }；missing 为空才算「可判定」。
  // prepared：可选的 `buildHazardIndex(hf)` 结果，传入则复用（避免重复建索引）。
  function matchHazardRows(hf, texts, prepared) {
    if (!Array.isArray(hf) || !hf.length) return { rows: [], missing: [] };
    const norm = hazNorm;
    const idx = prepared || buildHazardIndex(hf);
    // 三级匹配：① 原名 ② 去掉尾部「(总尘)/(呼尘)」 ③ 拆逗号后逐段
    //   ② 是因为库里「滑石粉尘(游离SiO2含量＜10%)」用 dust 列区分总尘/呼尘，名字里不带后缀，
    //      而报告确认表写的是「滑石粉尘(游离SiO2含量<10%)(呼尘)」。
    const findHaz = function (key) {
      const k = norm(key);
      if (k && idx[k]) return idx[k];
      const k2 = norm(String(key).replace(/[（(]\s*(?:总尘|呼尘)\s*[)）]\s*$/, ''));
      if (k2 && idx[k2]) return idx[k2];
      return null;
    };
    const rows = [];
    const missing = [];
    for (const raw of uniqHazTexts(texts)) {
      const whole = String(raw === null || raw === undefined ? '' : raw).trim();
      if (whole && !isHazPlaceholder(whole)) {
        const h0 = findHaz(whole);          // ★ 先按**整格**匹配，避免把「1,2-二氯乙烷」按逗号拆坏
        if (h0) { rows.push(h0); continue; }
      }
      // 单元格含多个因素（逗号/顿号分隔）时逐个匹配
      const parts = String(raw).split(/[,，、;；\n]/).map(function (x) { return x.trim(); }).filter(Boolean);
      for (const p of parts) {
        if (isHazPlaceholder(p)) continue;              // 占位值 → 跳过（不算缺数据）
        const h = findHaz(p);
        if (!h) { missing.push(p); continue; }          // 记下缺哪些，最后统一判 ok
        rows.push(h);
      }
    }
    return { rows: rows, missing: missing };
  }

  // 依据确认表的危害因素名，查危害因素库对应行的某字段。
  // 返回 { ok: 是否可判定, values: [字段值...], missing: [没查到的因素] }；ok=false 表示「缺数据不能删」。
  function lookupHazardField(hf, texts, field, prepared) {
    if (!Array.isArray(hf) || !hf.length) return { ok: false, values: [], missing: [] };
    const r = matchHazardRows(hf, texts, prepared);
    const values = r.rows.map(function (h) { return String(h[field] === null || h[field] === undefined ? '' : h[field]); });
    return { ok: r.missing.length === 0, values: values, missing: r.missing };
  }

  // 判断一组文本里是否**全部都不含**关键词（空数组视为「无」→ true，配合 ok 使用）
  function noneContains(values, keywords) {
    if (!values.length) return true;
    for (const v of values) {
      for (const kw of keywords) {
        if (v.indexOf(kw) >= 0) return false;   // 有任一含关键词 → 不算「都没有」
      }
    }
    return true;
  }

  // 判断一串文本是否是「章节标题」段落（用于确定所属章节，做序号重排边界）
  // 规则：形如「1.2.3评价依据」「9.2 职业卫生管理」——以「数字(.数字)*」开头，
  //   后接**不以《/（/〔 开头**的短标题（≤30 字），且**不含法规书名号**（避免把「1.《xx》」这种条目当标题）。
  function isHeadingPara(t) {
    const s = String(t || '').replace(/\s+/g, '');
    if (!s) return false;
    const m = /^(\d+(?:\.\d+)*)(.+)$/.exec(s);
    if (!m) return false;
    const title = m[2];
    if (title.length > 30) return false;
    if (/^[《（(〔【]/.test(title)) return false;   // 后面紧跟书名号/括号 → 是条目，不是标题
    if (/[《》]/.test(title)) return false;         // 含书名号 → 是法规条目
    return true;
  }

  // 从段落文本里提取「（N）」序号（全角括号）；找不到返回 null
  function paraIndexNo(t) {
    const m = /^\s*[（(]\s*(\d+)\s*[)）]/.exec(String(t || ''));
    return m ? { no: parseInt(m[1], 10), len: m[0].length } : null;
  }

  // ---------------- 段落「按行」编辑工具（真实报告里一个 <w:p> 常含多行） ----------------
  // ★★ 关键事实（2026-10-17 实测 1095/1009）：报告里「1.1 法律、法规、规章」「1.2.3 评价依据」
  //   「9.2 建议」这些列表**不是一段一条**，而是**整节挤在一个 <w:p> 里、用 <w:br/> 软换行分行**。
  //   所以「命中即删整段」会把整节条目全删掉（灾难性过度删除），必须先按 <w:br/> 切行再逐行处理。
  //
  // 把段落切成「行」：返回 { full, breaks, lines:[{start,end,text}] }，偏移量对应**所有 <w:t> 拼接后的纯文本**。
  function paraLineMap(pXml) {
    const tokRe = /<w:t(?:\s[^>]*)?>([\s\S]*?)<\/w:t>|<w:br(?:\s[^>]*)?\/?>/g;
    let m;
    let full = '';
    const breaks = [];
    while ((m = tokRe.exec(pXml))) {
      if (m[0].indexOf('<w:br') === 0) breaks.push(full.length);
      else full += (m[1] || '');
    }
    const lines = [];
    let s = 0;
    for (const b of breaks) { lines.push({ start: s, end: b, text: full.slice(s, b) }); s = b; }
    lines.push({ start: s, end: full.length, text: full.slice(s) });
    return { full: full, breaks: breaks, lines: lines };
  }

  // 在段落内把纯文本区间 [start,end) 替换为 to（跨 run，未涉及的 run 原样保留）。
  // keepTag=true → 沿用原 <w:t> 的起始标签（与旧 replaceAcrossRuns 逐字节一致）；
  // 否则统一写成 <w:t xml:space="preserve">（⑥⑦ 用，保证首尾空格不被 Word 吞掉）。
  // ★★ 2026-10-21 修 bug：区间**跨越多个 run** 时，to 只能插入一次（落在第一个相交 run 上），
  //    其余相交 run 只保留区间之后的部分。原实现每个相交 run 都 `slice(0,a) + to + slice(b)`，
  //    于是「物理有害因素检测结果分析与评价」这种被 Word 拆成两个 run 的标题，
  //    替换后变成「物理因素检测结果分析与评价」重复两遍（实测 测试.docx 段落 #437：
  //    run0='7.1' run1=to run2=to）。单 run 内替换不受影响。
  function spliceParaRange(pXml, start, end, to, keepTag) {
    const re = /<w:t(?:\s[^>]*)?>([\s\S]*?)<\/w:t>/g;
    const nodes = [];
    let full = '';
    let m;
    while ((m = re.exec(pXml))) { const t = m[1] || ''; nodes.push({ s: full.length, e: full.length + t.length }); full += t; }
    let k = 0;
    let inserted = false;   // to 是否已插入（跨 run 时只插一次）
    return pXml.replace(re, function (raw, t) {
      const nd = nodes[k++];
      if (!nd || nd.e <= start || nd.s >= end) return raw;
      const a = Math.max(nd.s, start) - nd.s;
      const b = Math.min(nd.e, end) - nd.s;
      const ins = inserted ? '' : to;
      inserted = true;
      const nt = t.slice(0, a) + ins + t.slice(b);
      if (keepTag) {
        const open = /^<w:t(?:\s[^>]*)?>/.exec(raw);
        return (open ? open[0] : '<w:t>') + nt + '</w:t>';
      }
      return '<w:t xml:space="preserve">' + nt + '</w:t>';
    });
  }

  // 删掉段落里「纯文本偏移量 = offset」处的那个 <w:br/>（用于删行时一并去掉换行）
  function removeParaBreakAt(pXml, offset) {
    const tokRe = /<w:t(?:\s[^>]*)?>([\s\S]*?)<\/w:t>|<w:br(?:\s[^>]*)?\/?>/g;
    let m;
    let full = 0;
    const marks = [];
    while ((m = tokRe.exec(pXml))) {
      if (m[0].indexOf('<w:br') === 0) marks.push({ index: m.index, len: m[0].length, at: full });
      else full += (m[1] || '').length;
    }
    const hit = marks.filter(function (x) { return x.at === offset; })[0];
    if (!hit) return pXml;
    return pXml.slice(0, hit.index) + pXml.slice(hit.index + hit.len);
  }

  // 段内「（N）/（N）」按出现顺序重编为 1..k（只改行首序号，其余文本不动）
  function renumberParaLines(pXml) {
    const map = paraLineMap(pXml);
    const acts = [];
    let n = 0;
    for (const ln of map.lines) {
      const mm = /^\s*[（(]\s*(\d+)\s*[)）]/.exec(ln.text);
      if (!mm) continue;
      n++;
      if (parseInt(mm[1], 10) === n) continue;
      const prefix = mm[0];
      acts.push({ start: ln.start, end: ln.start + prefix.length, to: prefix.replace(/\d+/, String(n)) });
    }
    let out = pXml;
    acts.sort(function (a, b) { return b.start - a.start; });
    for (const a of acts) out = spliceParaRange(out, a.start, a.end, a.to);
    return out;
  }

  // 分句裁剪：把「含某组关键词」的分句从行里切掉（用于 9.2 那条「一句话覆盖粉尘/毒物/噪声/高温/紫外」的警示标识）
  //   · 按「；/;」切大句、再按「，/,」切分句；
  //   · 命中关键词的分句：从**最近的「。」之后**切到分句末尾（保留「…设置相应的警示标识和指令标识。」这类引导句）；
  //     找不到「。」→ 整个分句丢弃。
  //   ★ 2026-10-21：若命中的分句是「…工作场所」这种**主语分句**、而紧随其后的分句以「设置」开头，
  //     则那个谓语分句也一并丢弃。原因：⑪ 腐蚀那句原文是「…化学品工作场所**，**设置“当心腐蚀”…」，
  //     中间有逗号 → 关键词只能截到「工作场所」，分句裁剪会把主语删掉、留下「设置“当心腐蚀”…等警示标识」
  //     这样一句没有主语的残句。其他六条关键词都含「设置」，命中分句本身即完整，不会触发该分支。
  function trimClausesByKeywords(line, kws) {
    const src = String(line);
    const endedWithPeriod = /[。]\s*$/.test(src);
    const parts = src.split(/([；;])/);
    let res = '';
    for (let i = 0; i < parts.length; i += 2) {
      // 分隔符之后常带一个空格（如「…等警示标识; 有毒物品…」），拼接时去掉，避免出现「。 有毒物品」
      const seg = i === 0 ? parts[i] : parts[i].replace(/^\s+/, '');
      const sep = parts[i + 1] || '';
      if (!seg) { res += sep; continue; }
      const subs = seg.split(/([，,])/);
      let segOut = '';
      let swallowNext = false;   // 上一条命中分句是「…工作场所」主语 → 下一条「设置…」谓语一并丢
      for (let j = 0; j < subs.length; j += 2) {
        const sub = subs[j];
        const subSep = subs[j + 1] || '';
        if (!sub) continue;
        if (swallowNext) { swallowNext = false; continue; }
        const hit = kws.filter(function (k) { return sub.indexOf(k) >= 0; })[0];
        if (!hit) { segOut += sub + subSep; continue; }
        const at = sub.indexOf(hit);
        const dot = sub.lastIndexOf('。', at);
        if (dot >= 0) {
          const head = sub.slice(0, dot + 1);
          if (head.trim()) segOut += head;     // 保留引导句，不保留分隔符（避免留下孤立逗号）
        }
        // 无「。」→ 整分句丢弃；有「。」→ 只丢「。」之后的部分。
        // ★ 两种分支都要检查**紧跟的下一分句是否以「设置」开头**：
        //   若上一分句被删掉的正是「能引起职业性灼伤或腐蚀的化学品工作场所」这类**主语**，
        //   紧随的「设置“当心腐蚀”…」就是它的谓语，必须一并丢弃，否则留下无主残句。
        //   典型触发链：mixed 先删掉混合性粉(烟)尘分句时连分隔符一起吃掉，使
        //   「…指令标识。能引起…工作场所」粘成同一分句 → corrode 走的是「保留句号前引导句」分支
        //   （dot >= 0），若只在这个分支外检查就会漏掉 → 实测 0986 报告就是这么残留的。
        const nxt = subs[j + 2];
        if (nxt !== undefined && /^\s*设置/.test(nxt)) swallowNext = true;
      }
      segOut = segOut.replace(/[，,]\s*$/, '');
      if (segOut.trim()) {
        // 若这一句已经以句号收尾（删掉的正是句号后的分句），就不要再补分号，避免出现「。;」
        res += /[。]\s*$/.test(segOut) ? segOut : (segOut + sep);
      }
    }
    res = res.replace(/[；;，,]\s*$/, '');                 // 去掉结尾残留的分隔符
    res = res.replace(/。\s*[；;]\s*/g, '。');             // 「。;」→「。」
    if (endedWithPeriod && res && !/[。]\s*$/.test(res)) res += '。';
    return res;
  }

  // 删除命中的条目并做序号重排。
  // targets: 需要删除的文本数组。
  //   · 段落**只有一行**（无 <w:br/>）→ 沿用「整段删」语义；
  //   · 段落**含多行**（<w:br/> 分行）→ 只删命中的那一行（保留同段其它条目），并在段内重排序号。
  // opts.trimClauses: [{ kws:[...] }] —— 对「整行不是纯条目、只是一句话里的一个分句」的情况做分句裁剪。
  // 返回 { xml, changed, sections: 受影响的章节数 }
  function deletePjbgParas(xml, targets, opts) {
    const o = opts || {};
    const list = (targets || []).filter(Boolean);
    const trims = (o.trimClauses || []).filter(function (r) { return r && r.kws && r.kws.length; });
    if (!list.length && !trims.length) return { xml: xml, changed: 0, sections: 0 };
    // 逐段扫描，标记删除 + 记录章节（顺带记下段落在全文中的位置，末尾只替换改过的段落）
    const segs = [];   // { pos, orig, xml, text, del, heading, hasBr }
    eachParagraph(xml, function (pXml, pos) {
      const t = paraText(pXml);
      // ★ 性能：绝大多数段落没有 <w:br/>。先做一次子串预判，避免对 6000+ 个段落各建两次行映射。
      segs.push({ pos: pos, orig: pXml, xml: pXml, text: t, del: false,
        heading: isHeadingPara(t), hasBr: pXml.indexOf('<w:br') >= 0 });
    });
    let changed = 0;
    // ---- 先做「行级」处理（多行段落） ----
    for (const seg of segs) {
      if (seg.heading || !seg.hasBr) continue;      // 无 <w:br/> → 单行段，留给下面的整段删逻辑
      const map = paraLineMap(seg.xml);
      if (map.lines.length <= 1) continue;
      const acts = [];
      let lineDel = 0;
      for (const ln of map.lines) {
        const txt = ln.text;
        if (!txt) continue;
        const body = txt.replace(/^\s*[（(]\s*\d+\s*[)）]/, '');   // 去掉行首序号后再判断
        let hitTarget = null;
        for (const tg of list) { if (txt.indexOf(tg) >= 0) { hitTarget = tg; break; } }
        if (hitTarget) {
          // 「整行就是这个条目」→ 删整行；否则（条目只是长句里的一小段）→ 走分句裁剪
          const wholeLine = body.indexOf(hitTarget) === 0 || body.length <= hitTarget.length + 24;
          if (wholeLine) {
            acts.push({ start: ln.start, end: ln.end, to: '', brAt: ln.start });
            lineDel++;
            continue;
          }
        }
        let newTxt = txt;
        for (const rule of trims) {
          if (rule.kws.some(function (k) { return newTxt.indexOf(k) >= 0; })) {
            newTxt = trimClausesByKeywords(newTxt, rule.kws);
          }
        }
        if (newTxt !== txt) acts.push({ start: ln.start, end: ln.end, to: newTxt, brAt: -1 });
      }
      if (!acts.length) continue;
      // 从后往前应用，保证前面的偏移量仍然有效
      acts.sort(function (a, b) { return b.start - a.start; });
      let out = seg.xml;
      for (const a of acts) {
        out = spliceParaRange(out, a.start, a.end, a.to);
        if (a.brAt >= 0) out = removeParaBreakAt(out, a.brAt);   // 删行后文本左移，换行现在落在 start 处
      }
      if (lineDel > 0) out = renumberParaLines(out);
      seg.xml = out; seg.text = paraText(out);
      changed += acts.length;
    }
    // ---- 再做「整段删」（单行段落） ----
    for (const seg of segs) {
      if (seg.heading || seg.del) continue;
      if (seg.hasBr) continue;   // 多行段已在上面按行处理，绝不整段删
      for (const tg of list) { if (seg.text.indexOf(tg) >= 0) { seg.del = true; changed++; break; } }
    }
    // 计算每个非标题段所属章节索引（最近的上一个标题段）
    let sect = -1;
    const sectOf = [];
    for (let i = 0; i < segs.length; i++) {
      if (segs[i].heading) sect = i;
      sectOf[i] = sect;
    }
    if (!changed) return { xml: xml, changed: 0, sections: 0 };
    // 标记受影响章节（有删除的章节）
    const dirty = new Set();
    for (let i = 0; i < segs.length; i++) if (segs[i].del && sectOf[i] >= 0) dirty.add(sectOf[i]);
    // 章节内序号重排：对每个受影响章节，按顺序把未删段的「（N）」重编为 1..k
    for (const sh of dirty) {
      let n = 0;
      for (let i = sh + 1; i < segs.length; i++) {
        if (segs[i].heading) break;                 // 到达下一章节
        if (segs[i].del) continue;
        const cur = segs[i].text;
        const idx = paraIndexNo(cur);
        if (!idx) continue;
        n++;
        if (idx.no === n) continue;                 // 序号已正确，不动
        // 把首部「（N）」替换为「（n）」——跨 run 替换，保留格式
        const from = /^\s*[（(]\s*\d+\s*[)）]/.exec(cur)[0];
        const to = from.replace(/\d+/, String(n));
        const np = replaceAcrossRuns(segs[i].xml, from, to);
        if (np !== null) { segs[i].xml = np; segs[i].text = paraText(np); }
      }
    }
    // 重组 XML：**只就地替换改过的段落**，表格 / sectPr / 其它非段落内容一律原样保留。
    // ⚠⚠ 早先的实现是「把所有段落拼起来当输出」（out += seg.xml），会把**段落之外的全部内容**
    //   整块丢掉——实测 1095 报告 21 张表、约 1.14MB 内容会凭空消失（文档体积从 7.7MB 掉到 6.6MB）。
    //   因为此前 changed 恒为 0（判定永远不成立）从未走到这一步，所以这个坑一直没暴露。
    // ★ 性能：从后往前 splice 改过的段落即可，不再对整篇 7.7MB 做一次 replace。
    let out = xml;
    for (let i = segs.length - 1; i >= 0; i--) {
      const s = segs[i];
      const next = s.del ? '' : s.xml;
      if (next === s.orig) continue;                      // 未改动 → 跳过
      out = out.slice(0, s.pos) + next + out.slice(s.pos + s.orig.length);
    }
    return { xml: out, changed: changed, sections: dirty.size };
  }

  // 主入口⑥：按条件计算删除目标，再执行「删条目 + 重排」
  // opt = { noiseWorkshopEmpty: bool, hazardFactors: [] }
  // 主入口⑥：按条件计算删除目标，再执行「删条目 + 重排」
  // opt = { noiseWorkshopEmpty: bool, hazardFactors: [] }
  // ★ 2026-10-21 口径（用户逐条指定，编号与需求一一对应）：
  //   ① 无「粉尘」（「粉尘、毒物」不算）      → 删《中华人民共和国尘肺病防治条例》
  //   ② 无「毒物」（★ 2026-10-21 用户改为**宽口径**：「粉尘、毒物」**算**含毒物）
  //                                          → 删《使用有毒物品作业场所劳动保护条例》
  //   ③ 无高毒物品（判读同旧版）             → 删《高毒物品目录》
  //   ④ 既无「毒物」也无「粉尘」             → 删《呼吸防护装备的选择、使用和维护》(GB 18664-2025)
  //   ⑤ 表4-1 危害因素列（**处理后**）无「噪声」 → 删《听力防护装备的选择、使用和维护》(GB 23466-2025)
  //   ⑥ 无「毒物」（「粉尘、毒物」不算）      → 删《呼吸防护 自吸过滤式防毒面具》(GB 2890-2022)
  //   ⑦ 无「粉尘」或「粉尘、毒物」           → 删《呼吸防护 自吸过滤式防颗粒物呼吸器》(GB 2626-2019)
  //   ⑧ 同上                                 → 删 9.2「产生粉尘的工作场所设置…」分句
  //   ⑨ 无「粉尘、毒物」                     → 删「产生含有有毒物质的混合性粉(烟)尘…」分句
  //   ⑩ 无「毒物」（「粉尘、毒物」不算）      → 删「有毒物品工作场所设置…」分句
  //   ⑪ 无「腐蚀」                           → 删「能引起职业性灼伤或腐蚀的化学品工作场所…」分句
  //   ⑫ 表4-1 危害因素列（**处理后**）无「噪声」 → 删「产生噪声的工作场所设置…」分句
  //   ⑬ 无「高温」                           → 删「高温工作场所设置…」分句
  //   ⑭ 无「紫外辐射」                       → 删「产生紫外辐射的工作场所设置…」分句
  //   「有/无」按**检测项目确认表匹配到的因素在危害因素库里的「物化性质」**判定；
  //   确认表里有因素在库中找不到 → ①-④⑥-⑪⑬⑭ 一律不删（保守）；⑤⑫ 只看表4-1，不受影响。
  function fixConditionalDeletions(xml, opt) {
    const o = opt || {};
    const targets = [];
    const trims = [];
    const reasons = [];
    const pushTrim = function (kws, why) {
      for (const k of kws) targets.push(k);        // 兜底：该句单独成段时整段删
      trims.push({ kws: kws });                    // 主路径：长句里按分句裁掉
      reasons.push(why);
    };

    // ---- ⑤⑫ 噪声：看**表4-1「职业病危害因素」列处理之后还剩不剩「噪声」** ----
    //   ②/②c 已把非噪声岗位、低噪点位的「噪声」删掉，这里看到的就是「处理后的结果」；
    //   找不到表4-1 时退回旧口径（噪声车间岗位是否为空）。
    const t41Noise = t41HasNoise(xml);             // true / false / null
    const hasNoise = t41Noise === null ? (o.noiseWorkshopEmpty === true ? false : true) : t41Noise;
    if (!hasNoise) {
      targets.push(PJBG_DEL_TARGETS.hearing23466);
      pushTrim(PJBG_TRIM_KWS.noise, '⑤⑫表4-1 危害因素列（处理后）无噪声→删 GB23466 + 噪声警示句');
    } else {
      reasons.push(t41Noise === null ? '⑤⑫表4-1 未找到→噪声按旧口径保留' : '⑤⑫表4-1 危害因素列仍有噪声→保留噪声类条目');
    }

    // ---- ①-④⑥-⑪⑬⑭ 依赖危害因素库 ----
    const hf = o.hazardFactors;
    const cf = collectConfirmColTexts(xml, ['职业病危害因素']);
    if (cf.found && Array.isArray(hf) && hf.length) {
      // ★ 性能：确认表单元格先去重，库索引只建一次，只匹配一次（各字段都从同一批行取）
      const cells = uniqHazTexts(cf.texts);
      const hIdx = buildHazardIndex(hf);
      const mr = matchHazardRows(hf, cells, hIdx);
      if (mr.missing.length === 0) {
        const phy = mr.rows.map(function (h) { return hazNorm(h && h.physchem); });
        const names = mr.rows.map(function (h) { return hazNorm(h && h.name); });
        const MIX = hazNorm('粉尘、毒物');
        const anyOf = function (arr, kw) { return arr.some(function (v) { return v.indexOf(kw) >= 0; }); };
        const hasMixed = phy.some(function (v) { return v === MIX; });
        const hasDustWide = anyOf(phy, '粉尘');                                         // 含「粉尘、毒物」
        const hasDustStrict = phy.some(function (v) { return v.indexOf('粉尘') >= 0 && v !== MIX; });
        const hasToxWide = anyOf(phy, '毒物');
        const hasToxStrict = phy.some(function (v) { return v.indexOf('毒物') >= 0 && v !== MIX; });
        const hasCorrode = anyOf(phy, '腐蚀');
        const hasHeat = names.some(function (v) { return v === hazNorm('高温'); });
        const hasUV = names.some(function (v) { return v === hazNorm('紫外辐射'); });
        const hasHighTox = mr.rows.some(function (h) { return hazNorm(h && h.highTox).indexOf('是') >= 0; });

        if (!hasDustStrict) {
          targets.push(PJBG_DEL_TARGETS.dustLaw);
          reasons.push('①无「粉尘」（粉尘、毒物不算）→删尘肺病条例');
        }
        // ② 这条法规按**宽口径**：只要字段里出现「毒物」二字就保留
        //    （含「粉尘、毒物」—— 2026-10-21 用户明确：混合性质也算涉及有毒物品）
        if (!hasToxWide) {
          targets.push(PJBG_DEL_TARGETS.toxLaw);
          reasons.push('②无「毒物」（含「粉尘、毒物」也算有毒物）→删使用有毒物品条例');
        }
        // ⑥⑩ 仍按**严格口径**：「粉尘、毒物」不算有毒物
        if (!hasToxStrict) {
          targets.push(PJBG_DEL_TARGETS.mask2890);
          pushTrim(PJBG_TRIM_KWS.tox, '⑥⑩无「毒物」（粉尘、毒物不算）→删 GB2890/有毒物品警示句');
        }
        if (!hasHighTox) {
          targets.push(PJBG_DEL_TARGETS.highToxList);
          reasons.push('③无高毒物品→删高毒目录');
        }
        if (!hasDustWide && !hasToxWide) {
          targets.push(PJBG_DEL_TARGETS.resp18664);
          reasons.push('④无毒物且无粉尘→删 GB18664');
        }
        if (!hasDustWide) {
          targets.push(PJBG_DEL_TARGETS.resp2626);
          pushTrim(PJBG_TRIM_KWS.dust, '⑦⑧无「粉尘」或「粉尘、毒物」→删 GB2626 + 粉尘警示句');
        }
        if (!hasMixed) pushTrim(PJBG_TRIM_KWS.mixed, '⑨无「粉尘、毒物」→删混合性粉(烟)尘警示句');
        if (!hasCorrode) pushTrim(PJBG_TRIM_KWS.corrode, '⑪无「腐蚀」→删腐蚀类警示句');
        if (!hasHeat) pushTrim(PJBG_TRIM_KWS.heat, '⑬无「高温」→删高温警示句');
        if (!hasUV) pushTrim(PJBG_TRIM_KWS.uv, '⑭无「紫外辐射」→删紫外辐射警示句');
      } else {
        reasons.push('确认表有因素在库中找不到（' + mr.missing.slice(0, 4).join('、') + '）→粉尘/毒物/高毒/腐蚀/高温/紫外 均不删');
      }
    } else if (!cf.found) {
      reasons.push('未找到「检测项目确认表」→粉尘/毒物/高毒/腐蚀/高温/紫外 均不删');
    } else {
      reasons.push('危害因素库为空→粉尘/毒物/高毒/腐蚀/高温/紫外 均不删');
    }

    if (!targets.length && !trims.length) return { xml: xml, changed: 0, reasons: reasons };
    const r = deletePjbgParas(xml, targets, { trimClauses: trims });
    return { xml: r.xml, changed: r.changed, sections: r.sections, reasons: reasons };
  }

  // ---------------- ⑦ 「两班制」作息说明句标红 ----------------
  // 需求（用户 2026-10-17）：报告导出时把
  //   「打磨岗位实行两班制，上班时间为白班8:00～20:00、晚班20:00～8:00（次日），
  //     其余各岗位均实行白班制，上班时间为8:00～11:30、12:30～17:00。」
  // 这句**标红**（字体颜色 FF0000），提醒人工核对班制与作息时间。
  // 定位方式：在段落文本里找锚点「实行两班制」→ 向左回到上一个「。」之后、向右到下一个「。」（含），
  //   只把这一句上色，同段其它文字（如「注：职工总数…接触人数…。」）保持原样。
  // ⚠ 锚点用「实行两班制」而不是「打磨岗位…」：岗位名/时间可能随项目变化，锚点越短越稳；
  //   实测两份真实报告里「两班制」都只出现 1 次，不会误伤其它章节。
  const SHIFT_ANCHOR = '实行两班制';
  const SHIFT_COLOR = 'FF0000';

  // 在 <w:rPr> 里设置字体颜色。按 OOXML 子元素顺序插入（w:color 必须在 spacing/kern/sz/lang 之前），
  // 否则 Word 可能提示文件格式错误。
  function setRPrColor(rPr, hex) {
    let s = String(rPr || '');
    if (!s) s = '<w:rPr></w:rPr>';
    s = s.replace(/<w:color\b[^>]*\/>/g, '');          // 去掉已有颜色，避免重复
    const color = '<w:color w:val="' + hex + '"/>';
    const later = /<w:(?:spacing|w|kern|position|sz|szCs|highlight|u|effect|bdr|shd|fitText|vertAlign|rtl|cs|em|lang|eastAsianLayout|specVanish|oMath)\b/;
    const m = later.exec(s);
    if (m) return s.slice(0, m.index) + color + s.slice(m.index);
    return s.replace(/<\/w:rPr>\s*$/, color + '</w:rPr>');
  }

  // 把段落内纯文本区间 [start,end) 上色：按 run 切分（前段/着色段/后段），其余 run 原样保留。
  // 返回新段落 XML；没命中返回 null。
  function colorParaRange(pXml, start, end, hex) {
    const reRun = /<w:r(?:\s[^>]*)?>[\s\S]*?<\/w:r>/g;
    let off = 0;
    let hit = false;
    const out = pXml.replace(reRun, function (runXml) {
      // 该 run 的全部 <w:t> 文本（用于推进偏移；图片/制表符 run 没有 <w:t>）
      const all = runXml.match(/<w:t(?:\s[^>]*)?>([\s\S]*?)<\/w:t>/g) || [];
      const texts = all.map(function (x) { return (/<w:t(?:\s[^>]*)?>([\s\S]*?)<\/w:t>/.exec(x) || [])[1] || ''; });
      const txt = texts.join('');
      const rs = off, re2 = off + txt.length;
      off = re2;
      if (re2 <= start || rs >= end) return runXml;      // 不在范围内
      // 只处理「单个 <w:t>、无软换行」的简单 run（真实报告就是这种）；
      // 复杂的 run（多 <w:t> / 含 <w:br/>）原样放过，宁可不上色也不破坏结构。
      if (texts.length !== 1 || runXml.indexOf('<w:br') >= 0) return runXml;
      const a = Math.max(rs, start) - rs;
      const b = Math.min(re2, end) - rs;
      const rPrM = /<w:rPr>[\s\S]*?<\/w:rPr>/.exec(runXml);
      const rPr = rPrM ? rPrM[0] : '<w:rPr></w:rPr>';
      const rPrRed = setRPrColor(rPr, hex);
      const mk = function (text, rp) {
        return '<w:r>' + rp + '<w:t xml:space="preserve">' + text + '</w:t></w:r>';
      };
      const parts = [];
      if (txt.slice(0, a)) parts.push(mk(txt.slice(0, a), rPr));
      if (txt.slice(a, b)) parts.push(mk(txt.slice(a, b), rPrRed));
      if (txt.slice(b)) parts.push(mk(txt.slice(b), rPr));
      hit = true;
      return parts.join('');
    });
    return hit ? out : null;
  }

  // 主入口⑦：把「…实行两班制…」那句标红。返回 { xml, changed }
  function markShiftNoteRed(xml) {
    if (xml.indexOf(SHIFT_ANCHOR) < 0) return { xml: xml, changed: 0 };
    let changed = 0;
    // ★ 性能：只扫「含两」的段落（锚点「实行两班制」必含「两」，跨 run 拆分也成立），
    //   避免全篇 6000+ 段落扫描。
    const out = mapParasContaining(xml, ['两'], function (pXml) {
      if (pXml.indexOf('两') < 0) return pXml;
      const full = paraTextConcat(pXml);            // 原始拼接文本（不 trim，保证偏移与 <w:t> 对齐）
      const ai = full.indexOf(SHIFT_ANCHOR);
      if (ai < 0) return pXml;
      const prev = full.lastIndexOf('。', ai);
      const start = prev >= 0 ? prev + 1 : 0;       // 句首：上一个句号之后
      const nextDot = full.indexOf('。', ai);
      const end = nextDot >= 0 ? nextDot + 1 : full.length;   // 句尾：下一个句号（含）
      const np = colorParaRange(pXml, start, end, SHIFT_COLOR);
      if (np && np !== pXml) { changed++; return np; }
      return pXml;
    });
    return { xml: out, changed: changed };
  }

  // ---------------- ⑨ 表3-3「原辅材料情况」下方插入「注1~注4」（整段标红） ----------------
  // 需求（用户 2026-10-21）：在表3-3 表格下面增加一段，说明「漆料主要成分」的取值口径，**文字标红**。
  //   定位：先找含「表3-3」且含「原辅材料」的标题段落 → 取其后第一张 <w:tbl> → 在 </w:tbl> 之后插入新段落；
  //         找不到标题时按表头（物料名称 + 主要成分 + 年用量）兜底找表。
  //   格式：克隆「标题段落之前那一段正文」的 <w:pPr> 与首个 <w:r> 的 <w:rPr>（保持仿宋/字号/缩进一致），
  //         并强制把字符颜色改成 FF0000；4 条注用 <w:br/> 分行（= 一段 4 行）。
  //   幂等：文档里已出现「注1：本次检测对用人单位漆料」→ 直接跳过（重复导出不会叠加）。
  const PJBG_MATERIAL_NOTES = [
    '注1：本次检测对用人单位漆料进行有机组分定性分析，体现的“主要成分”指：定性结果中纳入《工作场所有害因素职业接触限值 第1部分：化学有害因素》（GBZ 2.1）和《职业病危害因素分类目录》（国卫疾控发〔2015〕92号）的化学组分。',
    '注2：本次检测对用人单位漆料进行有机组分定性分析，定性结果中的有机组分均未纳入《工作场所有害因素职业接触限值 第1部分：化学有害因素》（GBZ 2.1）和《职业病危害因素分类目录》（国卫疾控发〔2015〕92号），故物料的主要成分按“/”体现。',
    '注3：用人单位2025年度的《工作场所职业病危害因素检测报告》[报告编号：BTC（ZJ）250100]显示，漆料已进行有机组分定性分析，本次检测期间使用的漆料未发生变更，故漆料主要成分按照该报告的定性结果体现。',
    '注4：漆料的主要成分根据用人单位提供的漆料SDS/物料成分说明等资料体现。',
  ];
  const PJBG_NOTE_RED = 'FF0000';

  // 给 <w:rPr> 强制写入红色（保留其它字符格式）
  function forceRedRPr(rPr) {
    let out = rPr || '';
    if (!out) return '<w:rPr><w:color w:val="' + PJBG_NOTE_RED + '"/></w:rPr>';
    out = out.replace(/<w:color\b[^>]*\/>/g, '');          // 去掉已有颜色
    // 去掉加粗（表格表头单元格是加粗的，万一兜底克隆到表头，注段落也不该变粗）
    out = out.replace(/<w:b\s*\/>/g, '<w:b w:val="0"/>').replace(/<w:bCs\s*\/>/g, '<w:bCs w:val="0"/>');
    if (out.indexOf('</w:rPr>') >= 0) return out.replace('</w:rPr>', '<w:color w:val="' + PJBG_NOTE_RED + '"/></w:rPr>');
    return '<w:rPr><w:color w:val="' + PJBG_NOTE_RED + '"/></w:rPr>';
  }
  // 强制「首行不缩进 + 行距固定 14 磅」（14pt = 280 twips → w:line="280" w:lineRule="exact"）
  //   ⚠ OOXML 的 pPr 子元素有固定顺序（spacing 在 ind 之前，二者都在 jc/rPr 之前），
  //     所以按顺序插到第一个「后继元素」前面，而不是直接追加到 </w:pPr>。
  function forceNoteSpacing(pPr) {
    let out = String(pPr || '');
    if (!out) out = '<w:pPr></w:pPr>';
    out = out.replace(/<w:pPr\s*\/>/, '<w:pPr></w:pPr>');
    out = out.replace(/<w:spacing\b[^>]*\/>/g, '').replace(/<w:ind\b[^>]*\/>/g, '');
    const spacing = '<w:spacing w:line="280" w:lineRule="exact"/>';
    const ind = '<w:ind w:firstLine="0" w:firstLineChars="0"/>';
    const beforeSpacing = /<w:(?:ind|contextualSpacing|mirrorIndents|jc|textDirection|textAlignment|outlineLvl|rPr|sectPr)\b/;
    let m = beforeSpacing.exec(out);
    out = m ? out.slice(0, m.index) + spacing + out.slice(m.index) : out.replace('</w:pPr>', spacing + '</w:pPr>');
    const beforeInd = /<w:(?:contextualSpacing|mirrorIndents|jc|textDirection|textAlignment|outlineLvl|rPr|sectPr)\b/;
    m = beforeInd.exec(out);
    out = m ? out.slice(0, m.index) + ind + out.slice(m.index) : out.replace('</w:pPr>', ind + '</w:pPr>');
    return out;
  }

  // 取段落里第一个 run 的 <w:rPr>（先剔除 <w:pPr>，避免命中段落标记自带的 rPr）
  function firstRunRPr(pXml) {
    const pPrM = /<w:pPr(?:\s[^>]*)?>[\s\S]*?<\/w:pPr>/.exec(pXml);
    const pPr = pPrM ? pPrM[0] : '';
    const body = pPr ? pXml.replace(pPr, '') : pXml;
    const runM = /<w:r(?:\s[^>]*)?>[\s\S]*?<\/w:r>/.exec(body);
    if (!runM) return '';
    const rM = /<w:rPr(?:\s[^>]*)?>[\s\S]*?<\/w:rPr>/.exec(runM[0]);
    return rM ? rM[0] : '';
  }

  function addMaterialNote(xml) {
    // 幂等：已插入过就不再插
    if (xml.indexOf('注1：本次检测对用人单位漆料') >= 0) return { xml: xml, changed: 0, skipped: true };
    // 标题段落：优先「以 表3-3 开头」（真正的表题，如「表3-3 原辅材料情况」）；
    // 退而求其次才用「正文里提到 表3-3 的句子」（如「用人单位原辅材料情况详见表3-3。」）——
    // 否则格式模板会取到章节小标题（加粗/黑体），段落缩进也不对。
    // ★ 性能（2026-10-21）：原来 eachParagraph 全篇扫 4767 段（~8ms），改为只按「表3-3」定位
    //   （全篇出现次数个位数），命中后再向两侧扩到段落边界。
    let capStart = -1, capEnd = -1;
    let looseStart = -1, looseEnd = -1;
    let i = xml.indexOf('表3-3');
    while (i >= 0) {
      const s0 = paraStartBefore(xml, i);
      const e0 = xml.indexOf('</w:p>', i);
      if (s0 >= 0 && e0 >= 0) {
        const end0 = e0 + '</w:p>'.length;
        const t = cellText(xml.slice(s0, end0)).replace(/\s+/g, '');
        if (t && t.indexOf('表3-3') === 0) { capStart = s0; capEnd = end0; break; }
        if (looseStart < 0 && t && t.indexOf('原辅材料') >= 0) { looseStart = s0; looseEnd = end0; }
      }
      i = xml.indexOf('表3-3', i + 4);
    }
    if (capStart < 0 && looseStart >= 0) { capStart = looseStart; capEnd = looseEnd; }
    // 表题前一段（做段落格式模板）
    let prevPara = '';
    if (capStart > 0) {
      const pe = xml.lastIndexOf('</w:p>', capStart);
      if (pe > 0) {
        const ps = paraStartBefore(xml, pe);
        if (ps >= 0 && ps < pe) prevPara = xml.slice(ps, pe + '</w:p>'.length);
      }
    }
    let tblStart = -1, tblEnd = -1;
    if (capStart >= 0) {
      tblStart = xml.indexOf('<w:tbl', capEnd);
      if (tblStart >= 0) {
        const close = xml.indexOf('</w:tbl>', tblStart);
        if (close >= 0) tblEnd = close + '</w:tbl>'.length;
      }
    }
    if (tblEnd < 0) {
      // 兜底：按表头找「原辅材料情况」表（物料名称 / 主要成分 / 年用量）
      const tblRe = new RegExp(TBL_SRC, 'g');
      let tm;
      while ((tm = tblRe.exec(xml))) {
        const tbl = tm[0];
        if (tbl.indexOf('成') < 0 || tbl.indexOf('料') < 0) continue;
        const rows = tableRows(tbl);
        if (!rows.length) continue;
        const heads = rowCells(rows[0]).map(cellText).map(normKey);
        const has = function (kw) { return heads.some(function (h) { return h.indexOf(normKey(kw)) >= 0; }); };
        if (has('主要成分') && (has('物料名称') || has('年用量'))) {
          tblStart = tm.index; tblEnd = tm.index + tbl.length;
          break;
        }
      }
    }
    if (tblEnd < 0) return { xml: xml, changed: 0 };

    // 格式模板：
    //   · 段落格式（缩进/对齐/行距）沿用「标题前一段正文」——表格单元格是居中对齐，整段注居中也难看；
    //   · 字符格式（字体/字号）取**表3-3 数据单元格**里第一个 run —— 用户要求「和前面表格内字体格式一致」
    //     （实测表内 仿宋_GB2312 / sz 21，而正文是 sz 28，差两个字号）。跳过表头行（表头加粗）。
    const src = prevPara || (capStart >= 0 ? xml.slice(capStart, capEnd) : '');
    const pPrM = /<w:pPr(?:\s[^>]*)?>[\s\S]*?<\/w:pPr>/.exec(src || '');
    // 去掉段落标记自带的 <w:rPr>（那是「标题/正文」的字体，比如黑体 sz28 加粗），
    // 否则整段的段落标记仍是旧字号，与表格字体不一致
    const pPrRaw = pPrM ? pPrM[0].replace(/<w:rPr(?:\s[^>]*)?>[\s\S]*?<\/w:rPr>/g, '') : '';
    // 段落格式：**首行不缩进 + 行距固定 14 磅**（用户 2026-10-21）
    const pPr = forceNoteSpacing(pPrRaw);
    let rPrSrc = '';
    const tblXml = (tblStart >= 0 && tblEnd > tblStart) ? xml.slice(tblStart, tblEnd) : '';
    if (tblXml) {
      const trs = tableRows(tblXml);
      for (let ti = 1; ti < trs.length && !rPrSrc; ti++) {
        const tcs = rowCells(trs[ti]);
        for (const tc of tcs) {
          const pm = new RegExp(PARA_SRC).exec(tc);
          if (!pm) continue;
          const r = firstRunRPr(pm[0]);
          if (r) { rPrSrc = r; break; }
        }
      }
      if (!rPrSrc) {                                    // 兜底：整表第一个段落（可能是表头）
        const pm = new RegExp(PARA_SRC).exec(tblXml);
        if (pm) rPrSrc = firstRunRPr(pm[0]);
      }
    }
    const rPr = forceRedRPr(rPrSrc || firstRunRPr(src || ''));
    const runs = PJBG_MATERIAL_NOTES.map(function (line, i) {
      const br = i === 0 ? '' : '<w:r>' + rPr + '<w:br/></w:r>';
      return br + '<w:r>' + rPr + '<w:t xml:space="preserve">' + xmlEsc(line) + '</w:t></w:r>';
    }).join('');
    const para = '<w:p>' + pPr + runs + '</w:p>';
    return { xml: xml.slice(0, tblEnd) + para + xml.slice(tblEnd), changed: 1 };
  }

  // ---------------- ⑩ 表8-2「超标原因分析」按检测项目自动填「超标原因分析」列 ----------------
  // 需求（用户 2026-10-21）：报告 8.2「不符合岗位/工种超标原因分析」章节若有表8-2，
  //   逐行看「检测项目」列，按因素类型把对应模板写进「超标原因分析」列（多行，分号/句号分行）。
  //   · 噪声（检测项目含「噪声」）→ 4 行
  //   · 高温（检测项目含「高温」）→ 6 行
  //   · 检测项目按危害因素库「系统名称」匹配后，物化性质 =「粉尘」或「粉尘、毒物」→ 粉尘 6 行
  //   · 物化性质除「粉尘、毒物」外含「毒物」→ 毒物 5 行
  //   只填**空 / 占位（-、—、/、无）**的单元格，已有内容不动（避免覆盖人工填写的分析）。
  const PJBG_OVER_LIMIT_TEXTS = {
    noise: [
      'XX设备运行产生较大的机械性噪声（空气动力性噪声、电磁性噪声）；',
      '设备敞开式作业，缺少隔音防护；',
      '多台设备同时运转存在噪声叠加；',
      '气动装置高压气体排气口未安装消音装置，或消音装置消音效果不佳。',
    ],
    heat: [
      'XX设备运行存在生产性热源；',
      '设备敞开式作业或未设置隔热措施，热辐射易扩散到设备外部；',
      '环境温度较高；',
      '未设置防暑降温设施；',
      '工业风扇防暑降温效果欠佳；',
      '车间整体通风效果欠佳，存在热量积聚。',
    ],
    dust: [
      '打磨作业、搅拌作业为敞开式作业，易产生扬尘；',
      '投料、出料作业落差较大，易产生扬尘；',
      '高压吹扫物料表面易产生扬尘；',
      '抛丸机密封不严，作业过程粉尘易扩散到外部环境；',
      '打磨区除尘设施（罩口设置不合理、距离远、未设围挡）除尘效果欠佳，未能及时捕集逸散的粉尘；',
      '现场设备表面、地面等存在明显积尘，易形成二次扬尘。',
    ],
    tox: [
      '漆料中的甲苯含量较高，易挥发逸散至空气中；',
      '现场水帘机的吸风排毒效果欠佳，未能有效捕集逸散的废气；',
      '现场设备、管道存在跑冒滴漏现象，有废气的逸散；',
      '车间整体通风效果欠佳，导致废气蓄积；',
      '打磨区未与喷漆区进行物理隔离，导致喷漆区废气扩散到周边区域。',
    ],
  };

  // 把一个单元格的所有段落替换成「一段 N 行」（沿用首段的 pPr / 首个 run 的 rPr）
  function setCellLines(tcXml, lines) {
    const paraRe = new RegExp(PARA_SRC, 'g');
    const paras = tcXml.match(paraRe) || [];
    if (!paras.length || !lines || !lines.length) return tcXml;
    const tpl = paras[0];
    const pPrM = /<w:pPr(?:\s[^>]*)?>[\s\S]*?<\/w:pPr>/.exec(tpl);
    const pPr = pPrM ? pPrM[0] : '';
    const rPr = firstRunRPr(tpl);
    const runs = lines.map(function (ln, i) {
      const br = i === 0 ? '' : '<w:r>' + rPr + '<w:br/></w:r>';
      return br + '<w:r>' + rPr + '<w:t xml:space="preserve">' + xmlEsc(ln) + '</w:t></w:r>';
    }).join('');
    const np = '<w:p>' + pPr + runs + '</w:p>';
    let first = true;
    return tcXml.replace(paraRe, function () {
      if (first) { first = false; return np; }
      return '';
    });
  }

  // 检测项目文本 → 需要写入的「超标原因分析」行（按 噪声 / 高温 / 粉尘 / 毒物 顺序，去重）
  function overLimitLines(itemText, hIdx) {
    const out = [];
    const push = function (kind) {
      for (const ln of PJBG_OVER_LIMIT_TEXTS[kind]) if (out.indexOf(ln) < 0) out.push(ln);
    };
    const whole = hazNorm(itemText);
    if (whole.indexOf('噪声') >= 0) push('noise');
    if (whole.indexOf('高温') >= 0) push('heat');
    // 化学因素：按库「系统名称」逐个匹配（三级：原名 → 去「(总尘)/(呼尘)」→ 原样）
    const parts = String(itemText === null || itemText === undefined ? '' : itemText)
      .split(/[,，、;；\n]/).map(function (x) { return x.trim(); }).filter(Boolean);
    const MIX = hazNorm('粉尘、毒物');
    let dustDone = false, toxDone = false;
    for (const p of parts) {
      const k = hazNorm(p);
      if (!k || k.indexOf('噪声') >= 0 || k.indexOf('高温') >= 0) continue;
      let h = hIdx[k];
      if (!h) {
        const k2 = hazNorm(String(p).replace(/[（(]\s*(?:总尘|呼尘)\s*[)）]\s*$/, ''));
        if (k2) h = hIdx[k2];
      }
      if (!h) continue;
      const ph = hazNorm(h.physchem);
      if (ph === MIX || ph === hazNorm('粉尘')) { if (!dustDone) { push('dust'); dustDone = true; } }
      else if (ph.indexOf('毒物') >= 0) { if (!toxDone) { push('tox'); toxDone = true; } }
    }
    return out;
  }

  // 在 document.xml 里找「表8-2 超标原因分析表」（表头含「检测项目」+「超标原因分析」），逐行填空。
  // 返回 { xml, changed, filled }。
  function fillOverLimitCause(xml, hazardFactors) {
    if (xml.indexOf('超标原因分析') < 0) return { xml: xml, changed: 0, filled: 0 };
    const hIdx = buildHazardIndex(hazardFactors || []);
    const re = new RegExp(TBL_SRC, 'g');
    const jobs = [];
    let m;
    while ((m = re.exec(xml))) {
      const tbl = m[0];
      if (tbl.indexOf('超') < 0 || tbl.indexOf('项') < 0) continue;   // 廉价预筛
      const rows = tableRows(tbl);
      if (rows.length < 2) continue;
      const heads = rowCells(rows[0]).map(cellText);
      const cItem = headIndexOf(heads, '检测项目');
      const cWhy = headIndexOf(heads, '超标原因分析');
      if (cItem < 0 || cWhy < 0) continue;
      jobs.push({ start: m.index, len: tbl.length, tbl: tbl, rows: rows, heads: heads, cItem: cItem, cWhy: cWhy });
    }
    if (!jobs.length) return { xml: xml, changed: 0, filled: 0 };
    let out = xml;
    let filled = 0;
    // 从后往前替换，避免前面的长度变化让后面的偏移失效
    for (let ji = jobs.length - 1; ji >= 0; ji--) {
      const j = jobs[ji];
      let newTbl = j.tbl;
      for (let ri = 1; ri < j.rows.length; ri++) {
        const row = j.rows[ri];
        const cells = rowCells(row);
        if (cells.length < 2) continue;
        const shift = cells.length - j.heads.length;
        const ii = j.cItem + shift, wi = j.cWhy + shift;
        const iIdx = (ii >= 0 && ii < cells.length) ? ii : j.cItem;
        const wIdx = (wi >= 0 && wi < cells.length) ? wi : j.cWhy;
        if (wIdx >= cells.length) continue;
        const itemTxt = cellText(cells[iIdx]);
        const curWhy = cellText(cells[wIdx]);
        if (!isHazPlaceholder(curWhy)) continue;          // 已有内容 → 不覆盖
        const lines = overLimitLines(itemTxt, hIdx);
        if (!lines.length) continue;
        newTbl = newTbl.replace(row, row.replace(cells[wIdx], setCellLines(cells[wIdx], lines)));
        filled++;
      }
      if (newTbl !== j.tbl) out = out.slice(0, j.start) + newTbl + out.slice(j.start + j.len);
    }
    return { xml: out, changed: filled, filled: filled };
  }

  // ---------------- 快速重打包（B1）：只重压 document.xml，其余条目原始压缩字节直接拷贝 ----------------
  // 背景（2026-10-21 实测，0959：docx 1.88MB / document.xml 6.14MB / 31 个条目）：
  //   JSZip 的 generateAsync 会把**所有条目**重新 DEFLATE（图片 1.7MB 也重压一遍），耗时 584ms；
  //   而实际只有 document.xml 变了。这里直接按 ZIP 结构重建：
  //     · 其余条目：拷贝原始压缩数据 + 原始 method/crc/大小/时间戳/外部属性 → 零重压；
  //     · document.xml：原生 `CompressionStream('deflate-raw')` 压缩（实测 6.4MB → 145ms）；
  //     · 去掉 data descriptor（bit 3），本地头写真实大小。
  //   任何异常（ZIP64 / 结构不认识 / 没有原生 CompressionStream）→ 返回 null，调用方回退 JSZip。
  let PJBG_CRC_TABLE = null;
  function pjbgCrc32(u8) {
    if (!PJBG_CRC_TABLE) {
      const t = new Uint32Array(256);
      for (let n = 0; n < 256; n++) {
        let c = n;
        for (let k = 0; k < 8; k++) c = (c & 1) ? (0xEDB88320 ^ (c >>> 1)) : (c >>> 1);
        t[n] = c >>> 0;
      }
      PJBG_CRC_TABLE = t;
    }
    let crc = 0xFFFFFFFF;
    for (let i = 0; i < u8.length; i++) crc = PJBG_CRC_TABLE[(crc ^ u8[i]) & 0xFF] ^ (crc >>> 8);
    return (crc ^ 0xFFFFFFFF) >>> 0;
  }

  // 原生 deflate-raw；不支持则返回 null（走 JSZip 回退）
  async function pjbgDeflateRaw(u8) {
    if (typeof CompressionStream !== 'function') return null;
    try {
      const cs = new CompressionStream('deflate-raw');
      const w = cs.writable.getWriter();
      w.write(u8); w.close();
      const rd = cs.readable.getReader();
      const chunks = [];
      let total = 0;
      for (;;) {
        const r = await rd.read();
        if (r.done) break;
        chunks.push(r.value); total += r.value.length;
      }
      const out = new Uint8Array(total);
      let o = 0;
      for (const c of chunks) { out.set(c, o); o += c.length; }
      return out;
    } catch (e) { return null; }
  }

  // 解析 ZIP 的中央目录（只读，不碰数据）→ [{name, flags, method, time, date, crc, csize, usize, extAttr, dataOff}]
  function pjbgParseZip(buf) {
    const dv = new DataView(buf);
    const len = buf.byteLength;
    if (len < 22) return null;
    let eocd = -1;
    const floor = Math.max(0, len - 65558);
    for (let i = len - 22; i >= floor; i--) {
      if (dv.getUint32(i, true) === 0x06054b50) { eocd = i; break; }
    }
    if (eocd < 0) return null;
    const count = dv.getUint16(eocd + 10, true);
    let p = dv.getUint32(eocd + 16, true);
    const dec = new TextDecoder('utf-8');
    const entries = [];
    for (let i = 0; i < count; i++) {
      if (p + 46 > len || dv.getUint32(p, true) !== 0x02014b50) return null;
      const flags = dv.getUint16(p + 8, true);
      const method = dv.getUint16(p + 10, true);
      const time = dv.getUint16(p + 12, true);
      const date = dv.getUint16(p + 14, true);
      const crc = dv.getUint32(p + 16, true);
      const csize = dv.getUint32(p + 20, true);
      const usize = dv.getUint32(p + 24, true);
      const nlen = dv.getUint16(p + 28, true);
      const elen = dv.getUint16(p + 30, true);
      const clen = dv.getUint16(p + 32, true);
      const extAttr = dv.getUint32(p + 38, true);
      const off = dv.getUint32(p + 42, true);
      if (csize === 0xFFFFFFFF || usize === 0xFFFFFFFF || off === 0xFFFFFFFF) return null;  // ZIP64 → 不处理
      if (off + 30 > len || dv.getUint32(off, true) !== 0x04034b50) return null;
      const lnlen = dv.getUint16(off + 26, true);
      const lelen = dv.getUint16(off + 28, true);
      const dataOff = off + 30 + lnlen + lelen;
      if (dataOff + csize > len) return null;
      entries.push({
        name: dec.decode(new Uint8Array(buf, p + 46, nlen)),
        flags: flags, method: method, time: time, date: date,
        crc: crc, csize: csize, usize: usize, extAttr: extAttr, dataOff: dataOff,
      });
      p += 46 + nlen + elen + clen;
    }
    return entries;
  }

  // 按条目重新拼 ZIP（本地头 + 数据 + 中央目录 + EOCD）
  function pjbgBuildZip(items) {
    const enc = new TextEncoder();
    const parts = [];
    const cds = [];
    let offset = 0;
    for (const it of items) {
      const nameU8 = enc.encode(it.name);
      const lh = new Uint8Array(30 + nameU8.length);
      const dv = new DataView(lh.buffer);
      dv.setUint32(0, 0x04034b50, true);
      dv.setUint16(4, 20, true);                       // version needed
      dv.setUint16(6, it.flags & ~0x08, true);         // 清 data descriptor
      dv.setUint16(8, it.method, true);
      dv.setUint16(10, it.time, true);
      dv.setUint16(12, it.date, true);
      dv.setUint32(14, it.crc, true);
      dv.setUint32(18, it.csize, true);
      dv.setUint32(22, it.usize, true);
      dv.setUint16(26, nameU8.length, true);
      dv.setUint16(28, 0, true);                       // 不带 extra
      lh.set(nameU8, 30);
      parts.push(lh, it.data);
      const cd = new Uint8Array(46 + nameU8.length);
      const cdv = new DataView(cd.buffer);
      cdv.setUint32(0, 0x02014b50, true);
      cdv.setUint16(4, 20, true);
      cdv.setUint16(6, 20, true);
      cdv.setUint16(8, it.flags & ~0x08, true);
      cdv.setUint16(10, it.method, true);
      cdv.setUint16(12, it.time, true);
      cdv.setUint16(14, it.date, true);
      cdv.setUint32(16, it.crc, true);
      cdv.setUint32(20, it.csize, true);
      cdv.setUint32(24, it.usize, true);
      cdv.setUint16(28, nameU8.length, true);
      cdv.setUint16(30, 0, true);
      cdv.setUint16(32, 0, true);
      cdv.setUint16(34, 0, true);
      cdv.setUint16(36, 0, true);
      cdv.setUint32(38, it.extAttr, true);
      cdv.setUint32(42, offset, true);
      cd.set(nameU8, 46);
      cds.push(cd);
      offset += lh.length + it.data.length;
    }
    let cdSize = 0;
    for (const c of cds) cdSize += c.length;
    const eocd = new Uint8Array(22);
    const edv = new DataView(eocd.buffer);
    edv.setUint32(0, 0x06054b50, true);
    edv.setUint16(8, items.length, true);
    edv.setUint16(10, items.length, true);
    edv.setUint32(12, cdSize, true);
    edv.setUint32(16, offset, true);
    const total = offset + cdSize + 22;
    const out = new Uint8Array(total);
    let o = 0;
    for (const p of parts) { out.set(p, o); o += p.length; }
    for (const c of cds) { out.set(c, o); o += c.length; }
    out.set(eocd, o);
    return out.buffer;
  }

  // ---------------- 报告后处理⑪：1.1 / 1.2 章节的列表项改成 Word 自动编号 ----------------
  // 用户要求（2026-10-21）：报告 1.1（法律、法规、规章）与 1.2（技术规范与标准）里的列表项
  //   「(1)(2)(3)…」现在是**文字写死**的，删/增条目后不会自动调整。改成 Word 原生自动编号：
  //   · 原来「一个段落 + <w:br/> 软换行装 N 条」→ 每条拆成**独立段落**；
  //   · 段落格式（字体/字号/行距/首行缩进/对齐）逐项沿用原段落 → 视觉不变；
  //   · 编号格式「(N)」、编号后**不加空格/制表符**（suff=nothing），编号从 1 开始；
  //   · 每个原段落（= 一个列表）分配独立 numId（各自从 1 开始），编号定义写入 word/numbering.xml
  //     （平台模板原本没有该部件 → 需同时补 [Content_Types].xml override 与 document.xml.rels 关系）。
  // ⚠ 放在 ⑥ 之后执行：⑥ 先按条目文本删条目并重排文字序号，这里再把剩余条目的文字序号换成自动编号。
  const AUTONUM_ABSTRACT_ID = 90;
  const AUTONUM_FIRST_NUM_ID = 90;
  // pPr 子元素的 schema 顺序（numPr 必须插在 widowControl 之后、suppressLineNumbers 之前）
  const PPR_ORDER = ['pStyle', 'keepNext', 'keepLines', 'pageBreakBefore', 'framePr', 'widowControl',
    'numPr', 'suppressLineNumbers', 'pBdr', 'shd', 'tabs', 'suppressAutoHyphens', 'kinsoku', 'wordWrap',
    'overflowPunct', 'topLinePunct', 'autoSpaceDE', 'autoSpaceDN', 'bidi', 'adjustRightInd', 'snapToGrid',
    'spacing', 'ind', 'contextualSpacing', 'mirrorIndents', 'suppressOverlap', 'jc', 'textDirection',
    'textAlignment', 'textboxTightWrap', 'outlineLvl', 'divId', 'cnfStyle', 'rPr', 'sectPr', 'pPrChange'];

  // 把 <w:numPr> 按 schema 顺序插进 pPr（先删掉原有 numPr，如模板里的 numId=0 空编号）
  function insertPPrNumPr(pPr, numPrXml) {
    let s = String(pPr || '').replace(/<w:numPr>[\s\S]*?<\/w:numPr>/, '');
    if (s.indexOf('<w:pPr') < 0) s = '<w:pPr>' + s + '</w:pPr>';
    const idxNum = PPR_ORDER.indexOf('numPr');
    const re = /<w:([A-Za-z]+)(?=[\s/>])[^>]*?>/g;
    let m;
    let pos = -1;
    while ((m = re.exec(s))) {
      if (m.index === 0) continue;                       // 跳过 <w:pPr ...> 本身
      const oi = PPR_ORDER.indexOf(m[1]);
      if (oi > idxNum) { pos = m.index; break; }
    }
    if (pos < 0) return s.replace(/<\/w:pPr>\s*$/, numPrXml + '</w:pPr>');
    return s.slice(0, pos) + numPrXml + s.slice(pos);
  }

  // 把一个「多行列表段落」拆成 N 个自动编号段落；无法安全拆分时返回 null
  // 返回 { xml, items }
  function buildNumberedParas(pXml, numId) {
    const pPrM = /<w:pPr(?:\s[^>]*)?>[\s\S]*?<\/w:pPr>/.exec(pXml);
    if (!pPrM) return null;
    const bodyEnd = pXml.lastIndexOf('</w:p>');
    if (bodyEnd < 0) return null;
    const body = pXml.slice(pPrM.index + pPrM[0].length, bodyEnd);
    // ⚠ 保守：<w:br/> 若写在 <w:t> 内部（非标准写法），这里不处理
    if (/<w:t(?:\s[^>]*)?>[^<]*<w:br/.test(body)) return null;
    const firstRun = /<w:r(?:\s[^>]*)?>[\s\S]*?<\/w:r>/.exec(body);
    if (!firstRun) return null;
    const rPrM = /<w:rPr(?:\s[^>]*)?>[\s\S]*?<\/w:rPr>/.exec(firstRun[0]);
    const rPr = rPrM ? rPrM[0] : '';
    // 按 <w:br/> 切行（<w:t> 内容原样保留，<w:tab/> 等非文本节点自然丢弃）
    const lines = [];
    let cur = [];
    const TOK = /<w:t(?:\s[^>]*)?>([\s\S]*?)<\/w:t>|<w:br(?:\s[^>]*)?\/?>/g;
    let m;
    while ((m = TOK.exec(body))) {
      if (m[1] !== undefined) cur.push(m[1]);
      else { lines.push(cur.join('')); cur = []; }
    }
    lines.push(cur.join(''));
    // 只保留「(N)」开头的行，并去掉行首的编号
    const items = [];
    for (let raw of lines) {
      const s = String(raw).replace(/^[\s\u3000]+/, '');
      const mm = /^[（(]\s*(\d+)\s*[)）]/.exec(s);
      if (!mm) continue;
      const text = s.slice(mm[0].length).replace(/^[\s\u3000]+/, '');
      if (text) items.push(text);
    }
    if (!items.length) return null;
    const numPr = '<w:numPr><w:ilvl w:val="0"/><w:numId w:val="' + numId + '"/></w:numPr>';
    const newPPr = insertPPrNumPr(pPrM[0], numPr);
    let out = '';
    for (const it of items) {
      out += '<w:p>' + newPPr + '<w:r>' + rPr + '<w:t xml:space="preserve">' + it + '</w:t></w:r></w:p>';
    }
    return { xml: out, items: items.length };
  }

  // 把 document.xml 里 1.1 / 1.2 章节的列表项改成自动编号
  // 返回 { xml, changed(改了几个列表), items(总条目数), numIds: [用到的 numId] }
  function autoNumberItemLists(xml) {
    const out = { xml: xml, changed: 0, items: 0, numIds: [] };
    if (xml.indexOf('1.1') < 0 && xml.indexOf('1.2') < 0) return out;
    // 表格区间（表内段落不参与）
    const tblSpans = [];
    {
      const re = new RegExp(TBL_SRC, 'g');
      let tm;
      while ((tm = re.exec(xml))) tblSpans.push([tm.index, tm.index + tm[0].length]);
    }
    const inTable = function (pos) {
      for (const s of tblSpans) if (pos >= s[0] && pos < s[1]) return true;
      return false;
    };
    // 扫顶层段落：进入 1.1 / 1.2 后开始收集「(N)」开头的列表段落，遇到 1.3 结束
    const targets = [];
    let inScope = false;
    const re = new RegExp(PARA_SRC, 'g');
    let m;
    while ((m = re.exec(xml))) {
      const p = m[0];
      if (p.indexOf('<w:t') < 0) continue;
      if (inTable(m.index)) continue;
      const text = cellText(p);
      if (!text) continue;
      const first = String(text).split('\n')[0].trim();
      if (/^1\.1(?![.\d])/.test(first) || /^1\.2(?![.\d])/.test(first)) { inScope = true; continue; }
      if (/^1\.3(?![.\d])/.test(first)) { inScope = false; continue; }
      if (!inScope) continue;
      if (/^[（(]\s*\d+\s*[)）]/.test(first)) targets.push({ xml: p, index: m.index });
    }
    if (!targets.length) return out;
    let cur = xml;
    const numIds = [];
    let changed = 0;
    let items = 0;
    // 从后往前替换（前面的偏移不失效）；numId 仍按出现顺序分配
    const jobs = targets.map(function (t, i) { return { xml: t.xml, index: t.index, numId: AUTONUM_FIRST_NUM_ID + i }; });
    for (let i = jobs.length - 1; i >= 0; i--) {
      const built = buildNumberedParas(jobs[i].xml, jobs[i].numId);
      if (!built) continue;
      cur = cur.slice(0, jobs[i].index) + built.xml + cur.slice(jobs[i].index + jobs[i].xml.length);
      numIds.unshift(jobs[i].numId);
      changed++;
      items += built.items;
    }
    if (!changed) return out;
    return { xml: cur, changed: changed, items: items, numIds: numIds };
  }

  // 生成 word/numbering.xml（只含本步用到的 numId）
  function buildNumberingXml(numIds) {
    const lvl = '<w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="decimal"/>'
      + '<w:suff w:val="nothing"/><w:lvlText w:val="(%1)"/><w:lvlJc w:val="left"/>'
      + '<w:pPr><w:ind w:left="0" w:firstLine="420" w:firstLineChars="150"/></w:pPr>'
      + '<w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman" w:cs="Times New Roman"'
      + ' w:eastAsia="仿宋_GB2312"/><w:sz w:val="28"/><w:szCs w:val="28"/></w:rPr></w:lvl>';
    let s = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
      + '<w:numbering xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
      + '<w:abstractNum w:abstractNumId="' + AUTONUM_ABSTRACT_ID + '">'
      + '<w:multiLevelType w:val="singleLevel"/>' + lvl + '</w:abstractNum>';
    (numIds || []).forEach(function (id) {
      s += '<w:num w:numId="' + id + '"><w:abstractNumId w:val="' + AUTONUM_ABSTRACT_ID + '"/></w:num>';
    });
    return s + '</w:numbering>';
  }

  // 准备三个部件（numbering.xml / [Content_Types].xml / document.xml.rels）的文本；
  // 模板已有 numbering.xml 时返回 null（不冒险合并，调用方回退：不做自动编号）
  async function buildNumberingParts(zip, numIds) {
    if (zip.file('word/numbering.xml')) return null;
    const ctFile = zip.file('[Content_Types].xml');
    const relsFile = zip.file('word/_rels/document.xml.rels');
    if (!ctFile || !relsFile) return null;
    let ct = await ctFile.async('string');
    let rels = await relsFile.async('string');
    const CT_TYPE = 'application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml';
    if (ct.indexOf('/word/numbering.xml') < 0) {
      ct = ct.replace('</Types>', '<Override PartName="/word/numbering.xml" ContentType="' + CT_TYPE + '"/></Types>');
    }
    if (rels.indexOf('relationships/numbering') < 0) {
      let id = 'rIdNum1';
      let n = 1;
      while (rels.indexOf('Id="' + id + '"') >= 0) { n++; id = 'rIdNum' + n; }
      rels = rels.replace('</Relationships>', '<Relationship Id="' + id + '"'
        + ' Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering"'
        + ' Target="numbering.xml"/></Relationships>');
    }
    return [
      { name: 'word/numbering.xml', text: buildNumberingXml(numIds) },
      { name: '[Content_Types].xml', text: ct },
      { name: 'word/_rels/document.xml.rels', text: rels },
    ];
  }

  // 用新 document.xml 重打包；不可用时返回 null
  async function repackDocxFast(buf, newXml, extras) {
    try {
      // buf 可能是 ArrayBuffer（浏览器 fetch）或 Uint8Array/Buffer（Node 测试）→ 统一成 ArrayBuffer
      //   ⚠ 不能用 `instanceof ArrayBuffer`：测试在 vm 沙箱里跑，跨 realm 的 instanceof 恒为 false
      const isAB = function (x) { return !!x && Object.prototype.toString.call(x) === '[object ArrayBuffer]'; };
      const ab = isAB(buf) ? buf
        : (buf && isAB(buf.buffer) ? buf.buffer.slice(buf.byteOffset, buf.byteOffset + buf.byteLength) : null);
      if (!ab) return null;
      const entries = pjbgParseZip(ab);
      if (!entries || !entries.length) return null;
      if (!entries.some(function (e) { return e.name === 'word/document.xml'; })) return null;
      const xmlBytes = new TextEncoder().encode(newXml);
      const comp = await pjbgDeflateRaw(xmlBytes);
      if (!comp) return null;
      const docEntry = entries.find(function (e) { return e.name === 'word/document.xml'; });
      const items = entries.map(function (e) {
        if (e.name === 'word/document.xml') {
          return { name: e.name, flags: e.flags, method: 8, time: e.time, date: e.date,
            crc: pjbgCrc32(xmlBytes), csize: comp.length, usize: xmlBytes.length, extAttr: e.extAttr, data: comp };
        }
        return { name: e.name, flags: e.flags, method: e.method, time: e.time, date: e.date,
          crc: e.crc, csize: e.csize, usize: e.usize, extAttr: e.extAttr,
          data: new Uint8Array(ab, e.dataOff, e.csize) };
      });
      // 附加/替换部件（如 ⑪ 自动编号需要的 word/numbering.xml、[Content_Types].xml、document.xml.rels）
      for (const ex of (extras || [])) {
        if (!ex || !ex.name) continue;
        const bytes = new TextEncoder().encode(ex.text);
        const c = await pjbgDeflateRaw(bytes);
        if (!c) return null;
        const item = { name: ex.name, flags: 0x0800, method: 8,
          time: docEntry ? docEntry.time : 0, date: docEntry ? docEntry.date : 0,
          crc: pjbgCrc32(bytes), csize: c.length, usize: bytes.length, extAttr: 0, data: c };
        const idx = items.findIndex(function (it) { return it.name === ex.name; });
        if (idx >= 0) items[idx] = item; else items.push(item);
      }
      return pjbgBuildZip(items);
    } catch (e) {
      return null;
    }
  }

  // 对下载到的检评报告 docx 做全部后处理（读一次 document.xml，改完统一重打包）；
  // 任何一步失败都不影响出报告（回退到上一步结果 / 原始文件）。
  async function patchPjbgDocx(buf) {    try {
      const JSZip = await ensureJSZip();
      const zip = await JSZip.loadAsync(buf);
      const f = zip.file('word/document.xml');
      if (!f) return buf;
      let xml = await readZipEntryText(f);
      let changed = 0;
      // 危害因素库快照：⑥（条件删除）与⑩（表8-2 超标原因）共用，故提到函数作用域
      let hazardFactors = [];
      // ① 「存在高毒物品」列 0/1 → 否/是；占位横线「-」→ 否
      const r1 = fixHighToxColumn(xml);
      xml = r1.xml; changed += r1.changed;
      if (r1.changed) console.info('[检评报告] 「存在高毒物品」列 0/1/「-」→ 否/是，共修正 ' + r1.changed + ' 处');
      // ② 非噪声岗位（LEX 最大值 < 80 dB）删掉「噪声」危害；同时取「非噪声车间岗位」文本供⑤用
      //
      // ★ 性能（2026-10-18）：②⑤⑥ 都要读噪声表。原先是三个 *FromBuf 函数各自调用一次，
      //   每次都 JSZip.loadAsync 解压整份 docx + DOMParser 解析全文（同一报告被完整解析 3 次）。
      //   现在**只算一次** noiseInfo：优先复用本函数已解压出来的 xml（连解压都省掉），
      //   拿不到该接口时退回按 buf 解析（带记忆）。
      const u = root.SamplingUpload;
      let nonNoiseText = '';
      let noiseInfo = null;
      // ① 最高优先：用户已打开过「噪声数据结果」页签 → 该 docx 早已解析过，直接取缓存（零成本）
      if (u && typeof u.getCachedNoiseInfo === 'function') {
        try { noiseInfo = u.getCachedNoiseInfo(buf); } catch (e) { noiseInfo = null; }
      }
      // ② 次优：复用本函数**已经解压出来**的 xml 解析一次（省掉再次解压）
      if (!noiseInfo && u && typeof u.noiseInfoFromXml === 'function') {
        try { noiseInfo = u.noiseInfoFromXml(xml); } catch (e) { noiseInfo = null; }
      }
      // ③ 兜底：按 buf 解析（内部带记忆，同一 buf 只解压一次）
      if (!noiseInfo && u && typeof u.noiseInfoFromBuf === 'function') {
        try { noiseInfo = await u.noiseInfoFromBuf(buf); } catch (e) { noiseInfo = null; }
      }
      if (noiseInfo) {
        const r2 = fixNonNoiseHazards(xml, noiseInfo.ok ? noiseInfo.nonNoisePairs : null);
        xml = r2.xml; changed += r2.changed;
        if (r2.changed) console.info('[检评报告] 非噪声岗位删除「噪声」危害，共修正 ' + r2.changed + ' 处');
        nonNoiseText = noiseInfo.nonNoiseText || '';
        // ②c 噪声岗位（LEX 最大值 ≥ 80 dB）内「3 个测量值全部 < 80 dB」的点位，单独删掉「噪声」危害
        //   （前提：同岗位至少有一个点位 3 个测量值全部 ≥ 80 dB，由 upload.js 侧算好）
        const r2c = fixQuietPointHazards(xml, noiseInfo.ok ? noiseInfo.quietPointPairs : null);
        xml = r2c.xml; changed += r2c.changed;
        if (r2c.changed) console.info('[检评报告] 低噪点位删除「噪声」危害，共修正 ' + r2c.changed + ' 处');
      } else if (u && typeof u.noiseNonNoiseFromBuf === 'function') {
        // 兜底：老接口（页面缓存了旧版 upload.js 时）
        const pairs = await u.noiseNonNoiseFromBuf(buf);
        const r2 = fixNonNoiseHazards(xml, pairs);
        xml = r2.xml; changed += r2.changed;
        if (r2.changed) console.info('[检评报告] 非噪声岗位删除「噪声」危害，共修正 ' + r2.changed + ' 处');
        if (typeof u.noiseNonNoiseTextFromBuf === 'function') {
          try { nonNoiseText = await u.noiseNonNoiseTextFromBuf(buf); } catch (e) { nonNoiseText = ''; }
        }
      }
      // ②b 表4-1「职业病危害因素」列：氮氧化物(一氧化氮和二氧化氮):氮氧化物(以二氧化氮计) → 氮氧化物(一氧化氮和二氧化氮)
      const r2b = fixNitrogenOxide(xml);
      xml = r2b.xml; changed += r2b.changed;
      if (r2b.changed) console.info('[检评报告] 表4-1 氮氧化物措辞精简，共修正 ' + r2b.changed + ' 处');
      // ③ 正文措辞替换（1.4 章节的接触限值说法、7.2 标题）
      const r3 = fixReportText(xml);
      xml = r3.xml; changed += r3.changed;
      if (r3.changed) console.info('[检评报告] 正文措辞替换，共修正 ' + r3.changed + ' 处');
      // ④ 表1-1 注3：按「备注」列实际标记裁剪说明
      const r4 = fixHazardNote(xml);
      xml = r4.xml; changed += r4.changed;
      if (r4.changed) console.info('[检评报告] 表1-1 注3 按备注列标记裁剪，共修正 ' + r4.changed + ' 处');
      // ⑤ 表4-1 下方「注：」段落替换（拆两段/可整段删除）+ 表3-1「接触职业病危害人数」联动表8-1 去重合计
      const r5 = fixNoteAndContactCount(xml, nonNoiseText);
      xml = r5.xml; changed += r5.changed;
      if (r5.changed) {
        console.info('[检评报告] 注段落/接触人数联动，共修正 ' + r5.changed + ' 处'
          + '（在册职工人数=' + r5.staffVal + '，接触职业病危害人数=' + r5.contactVal
          + (r5.total != null
            ? '，来源=' + (r5.totalFrom === 't41' ? '表4-1 人数/总数合计' : '表8-1 作业人数去重合计')
              + '=' + r5.total
              + (r5.t41 ? '（岗位 ' + r5.t41.groups + ' 个：纳入 ' + r5.t41.counted
                + '，危害因素列全为「-」/空而排除 ' + r5.t41.skipped + '）' : '')
            : '')
          + '，非噪声岗位文本=' + (r5.jobs ? '有(' + r5.jobs.length + '字)' : '空→仅保留注段A') + '）');
        console.info('[检评报告] 注段落文本：' + (r5.noteTexts || []).length + ' 段 → '
          + (r5.noteTexts || []).map(function (t) { return '「' + String(t).slice(0, 20) + '…」'; }).join(' + '));
      }
      // ⑥ 条件删除条目（噪声为空 / 无粉尘·毒物 / 无高毒物品 → 删对应评价依据、警示标识条目）+ 章节内序号重排
      {
        // 复用 ② 已算出的噪声信息（noiseWorkshopGroups 为空 = 全项目无噪声岗位）
        // ⚠ 语义与旧实现保持一致：解析失败（ok=false）时分组也是 []，等同「噪声车间岗位为空」→ 仍删；
        //   只有**根本没有 SamplingUpload**（Node 测试沙箱）时才不删（noiseWorkshopEmpty 保持 false）。
        let noiseWorkshopEmpty = false;
        if (noiseInfo) {
          noiseWorkshopEmpty = (noiseInfo.noiseWorkshopGroups || []).length === 0;
        } else if (u && typeof u.noiseWorkshopGroupsFromBuf === 'function') {
          try {
            const ws = await u.noiseWorkshopGroupsFromBuf(buf);
            noiseWorkshopEmpty = Array.isArray(ws) && ws.length === 0;
          } catch (e) { noiseWorkshopEmpty = false; }
        }
        const app = root.SamplingApp;
        if (app && typeof app.librarySnapshot === 'function') {
          try {
            const lib = app.librarySnapshot();
            if (lib && Array.isArray(lib.hazardFactors)) hazardFactors = lib.hazardFactors;
          } catch (e) { hazardFactors = []; }
        }
        const r6 = fixConditionalDeletions(xml, { noiseWorkshopEmpty: noiseWorkshopEmpty, hazardFactors: hazardFactors });
        xml = r6.xml; changed += r6.changed;
        if (r6.changed) {
          console.info('[检评报告] 条件删除条目，共删除 ' + r6.changed + ' 段'
            + (r6.sections ? '（涉及 ' + r6.sections + ' 个章节，序号已重排）' : '')
            + '：' + (r6.reasons || []).join('；'));
        }
      }
      // ⑦ 「…实行两班制…」作息说明句标红（提醒人工核对班制/作息时间）
      const r7 = markShiftNoteRed(xml);
      xml = r7.xml; changed += r7.changed;
      if (r7.changed) console.info('[检评报告] 「两班制」作息说明句已标红，共 ' + r7.changed + ' 处');
      // ⑧ 表3-6「个人使用的职业病防护用品配置及使用情况」：同「分类/类别/厂家/型号/参数」的多行合并
      const r8 = mergePpeRows(xml);
      xml = r8.xml; changed += r8.changed;
      if (r8.changed) console.info('[检评报告] 表3-6 防护用品合并同规格行，共合并 ' + r8.changed + ' 行');
      // ⑨ 表3-3「原辅材料情况」下方插入「注1~注4」说明（整段标红）
      const r9 = addMaterialNote(xml);
      xml = r9.xml; changed += r9.changed;
      if (r9.changed) console.info('[检评报告] 表3-3 下方已插入「注1~注4」（标红）');
      // ⑩ 表8-2「超标原因分析」按检测项目自动填「超标原因分析」列（只填空/占位）
      const r10 = fillOverLimitCause(xml, hazardFactors);
      xml = r10.xml; changed += r10.changed;
      if (r10.changed) console.info('[检评报告] 表8-2 超标原因分析已填写 ' + r10.filled + ' 行');
      // ⑪ 1.1 / 1.2 章节的「(1)(2)(3)…」列表项改成 Word 自动编号（删/增条目后自动重排）
      let numParts = null;
      const r11 = autoNumberItemLists(xml);
      if (r11.changed) {
        numParts = await buildNumberingParts(zip, r11.numIds);
        if (numParts) {
          xml = r11.xml; changed += r11.changed;
          console.info('[检评报告] 1.1/1.2 列表项改为自动编号，共 ' + r11.changed + ' 个列表 / ' + r11.items + ' 项');
        } else {
          console.warn('[检评报告] 自动编号部件不可用（模板已有 numbering.xml 或缺 rels），跳过本步');
        }
      }
      if (!changed) return buf;
      // 快速重打包（2026-10-21）：只重压 document.xml，其余条目原始压缩字节直接拷贝。
      //   实测 0959：JSZip generateAsync 584ms → 本路径约 150~200ms。
      //   不支持（老内核无 CompressionStream / ZIP64 / 结构异常）时回退 JSZip，行为与以前一致。
      const fast = await repackDocxFast(buf, xml, numParts);
      if (fast) return fast;
      zip.file('word/document.xml', xml);
      if (numParts) numParts.forEach(function (e) { zip.file(e.name, e.text); });
      return await zip.generateAsync({ type: 'arraybuffer', compression: 'DEFLATE' });
    } catch (e) {
      console.warn('[检评报告] 后处理失败，改用原始文件：', e);
      return buf;
    }
  }

  // 噪声页签下方「生成检评报告」按钮：按「数据上传」里选中的项目生成并下载 Word。
  // 若「噪声数据结果」页签已经生成并下载过该项目的报告，直接复用缓存（省掉平台 ~20s 生成）。
  // 全程不弹窗：进度 / 结果 / 错误都写在按钮旁的 #pjbg-status（错误红色加粗）。
  async function generatePjbgReport() {
    const status = $('pjbg-status');
    const setStatus = (msg, isErr) => {
      if (!status) return;
      status.textContent = msg || '';
      status.classList.toggle('err', !!isErr);
    };
    const api = uploadApi();
    const ctx = api && typeof api.getUpstreamContext === 'function' ? api.getUpstreamContext() : null;
    if (!ctx || !ctx.loggedIn || !ctx.project) {
      setStatus('尚未选择项目：请先到「数据上传」查询并选中项目', true);
      return;
    }
    const btn = $('btn-pjbg');
    const old = btn ? btn.textContent : '';
    try {
      const cached = typeof api.getPjbgCache === 'function' ? api.getPjbgCache(ctx.project.id) : null;
      let g;
      if (cached && cached.buf) {
        // 复用「噪声数据结果」已下载的报告：只需做后处理（高毒物品 / 非噪声岗位），秒级完成
        if (btn) { btn.disabled = true; btn.textContent = '处理中…'; }
        setStatus('正在复用「噪声数据结果」已生成的报告…');
        g = {
          buf: await patchPjbgDocx(cached.buf),
          fileName: cached.fileName || '',
          modName: cached.modName || '',
        };
      } else {
        if (btn) { btn.disabled = true; btn.textContent = '生成中…'; }
        setStatus('正在生成报告（平台约需 20 秒）…');
        g = await genPjbgDocx(ctx);
      }
      // 文件名沿用平台生成的原始文件名（与「直接在网站上导出」一致）
      const base = String(ctx.project.code || 'report').replace(/[\\/:*?"<>|]/g, '_');
      const name = g.fileName || (base + '-检评报告.docx');
      saveBlobAs(new Blob([g.buf], {
        type: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
      }), name);
      // 生成成功后不再显示「已复用报告缓存，耗时 Xs：<文件名>」这类长文案（用户要求去掉）；
      // 同时清掉「正在生成报告…」进度提示，避免残留。文件已自动下载，无需文字确认。
      setStatus('');
    } catch (e) {
      console.error(e);
      setStatus('生成检评报告失败：' + (e && e.message ? e.message : e), true);
    } finally {
      if (btn) { btn.disabled = false; btn.textContent = old || '生成检评报告'; }
    }
  }

  // 并行下载清单里的图片；失败的单张跳过
  async function fetchImageItems(c, list) {
    const out = await Promise.all(list.map(async (it) => {
      try {
        const enc = String(it.path).split('/').map(encodeURIComponent).join('/');
        const r = await fetch(c.apiBase + enc, { headers: c.token ? { Authorization: c.token } : {} });
        if (!r.ok) return null;
        return { name: it.name, mime: r.headers.get('content-type') || '', data: new Uint8Array(await r.arrayBuffer()) };
      } catch (e) { return null; }
    }));
    return out.filter(Boolean);
  }

  async function exportFlowAndLayout(project) {
    const api = uploadApi();
    const ctx = api && typeof api.getUpstreamContext === 'function' ? api.getUpstreamContext(project) : null;
    if (!ctx || !ctx.loggedIn) { alert('登录状态已失效，请重新登录后再试。'); return; }
    const btn = $('blf-export');
    const old = btn ? btn.textContent : '';
    if (btn) { btn.disabled = true; btn.textContent = '获取图片清单…'; }
    try {
      const q = '?belongProject=' + encodeURIComponent(ctx.project.id) + '&pageNumber=1&pageSize=-1';
      // 并行取两组清单
      const [ppJ, plJ] = await Promise.all([
        upRequest(ctx, 'GET', '/api/productionProcess' + q).then((r) => r.json()).catch(() => null),
        upRequest(ctx, 'GET', '/api/samplingPlan/projectLayout' + q).then((r) => r.json()).catch(() => null),
      ]);
      const flowList = pickRows(ppJ)
        .map((r) => ({ name: String(txt(r.processName)).trim(), path: String(txt(r.processPicturePath)).trim() }))
        .filter((x) => x.name && x.path);
      const layoutList = pickRows(plJ)
        .slice()
        .sort((a, b) => (a.id || 0) - (b.id || 0))   // 按 id 升序还原「图片1/2/3」的自然顺序
        .map((r) => ({ name: String(txt(r.layoutPictureName)).trim(), path: String(txt(r.layoutPicturePath)).trim() }))
        .filter((x) => x.path);
      if (!flowList.length && !layoutList.length) {
        alert('未找到工艺流程图或采样/测量布点图。\n\n请确认该项目已在平台上传：\n· 「生产工艺调查」的工艺图；\n· 「检测布点图」的布点图。');
        return;
      }
      if (btn) btn.textContent = '下载图片中…';
      const [flowImgs, layoutImgs] = await Promise.all([
        fetchImageItems(ctx, flowList),
        fetchImageItems(ctx, layoutList),
      ]);
      const extOf = (im) => (/jpe?g/i.test(im.mime) || /\.jpe?g$/i.test(im.name) ? 'jpg' : 'png');
      const safe = (s) => String(s || '').replace(/[\\/:*?"<>|]/g, '_').replace(/\s+/g, ' ').trim();
      if (btn) btn.textContent = '打包中…';
      const JSZip = await ensureJSZip();
      const pack = new JSZip();
      flowImgs.forEach((im, i) => pack.file('工艺流程图/' + (safe(im.name) || ('工艺流程图-' + (i + 1))) + '.' + extOf(im), im.data));
      layoutImgs.forEach((im, i) => pack.file('采样测量布点图/采样测量布点图-' + (i + 1) + '.' + extOf(im), im.data));
      const blob = await pack.generateAsync({ type: 'blob', compression: 'STORE' });  // 图片已压缩，STORE 更快
      const zipName = String((ctx.project && ctx.project.code) || 'report').replace(/[\\/:*?"<>|]/g, '_') + '-工艺点位图.zip';
      saveBlobAs(blob, zipName);
      alert('已保存压缩包：' + zipName + '\n' +
        '· 工艺流程图 ' + flowImgs.length + ' 张\n' +
        '· 采样/测量布点图 ' + layoutImgs.length + ' 张\n\n' +
        '共 ' + (flowImgs.length + layoutImgs.length) + ' 张图片，解压后为「工艺流程图」「采样测量布点图」两个文件夹。');
    } catch (e) {
      console.error(e);
      alert('保存图片失败：' + (e && e.message ? e.message : e));
    } finally {
      if (btn) { btn.disabled = false; btn.textContent = old || '工艺、点位图保存'; }
    }
  }
  // ---------------- 组装：导入模板（20 页） ----------------

  // 组装「导入模板」工作表。opts.worker / opts.survey / opts.dicts 控制包含哪些页：
  //   · 采样方案导入 → { worker: true,  survey: false, dicts: false }  ⇒ 只含主表格页，导入时不会碰调查表
  //   · 调查表导入 → { worker: false, survey: true,  dicts: false }  ⇒ 只含 6 张调查表，导入时不会碰主表格
  //   不传 opts 时全部包含（保持旧行为）。
  //   注：13 张字典页只作参考，importWorkbook 并不读取，故两个导入都不带（少写/少解析，更快）。
  function buildTemplateSheets(up, opts) {
    const wantWorker = !opts || opts.worker !== false;
    const wantSurvey = !opts || opts.survey !== false;
    const wantDicts = !opts || opts.dicts !== false;
    const D = data();
    const widthsTpl = D.widthsTemplate || {};
    const localSv = surveyData();
    const useUp = !!(up && (up.worker39 || up.survey));

    const sheets = [];
    // 1) 劳动定员和职业病危害因素接触情况调查（39 列，取自上游）—— 主表格（测点布局）的数据来源
    if (wantWorker) {
      const workerBody = (up && up.worker39) ? up.worker39.map((r) => WORKER_HEADERS.map((_, i) => txt(r[i]))) : [];
      sheets.push({
        name: '劳动定员和职业病危害因素接触情况调查',
        rows: [WORKER_HEADERS].concat(workerBody),
        widths: widthArray(widthsTpl['劳动定员和职业病危害因素接触情况调查'], WORKER_HEADERS.length),
        freeze: 1,
        dataValidations: [
          { sqref: 'M2:M1048576', formula1: '"是,否"' },
          { sqref: 'N2:N1048576', formula1: '"是,否"' },
          { sqref: 'AL2:AL1048576', formula1: '"是,否"' },
        ],
      });
    }
    // 2) 现场调查 6 张表
    if (wantSurvey) {
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
    }
    // 3) 13 张字典页
    if (wantDicts) {
      for (const d of D.dicts || []) {
        const rows = (d.rows || []).map((r) => r.map((v) => txt(v)));
        sheets.push({ name: d.name, rows: rows.length ? rows : [['']], widths: widthArray(widthsTpl[d.name], (rows[0] || [1]).length) });
      }
    }
    return sheets;
  }

  // ---------------- 项目搜索（移植自「数据上传」搜索模块） ----------------

  let blfProject = null;   // 本弹窗独立选中的项目
  let blfSeq = 0;          // 搜索序号（丢弃过期响应）
  let blfAbort = null;     // 进行中的搜索请求
  let blfDebounce = null;  // 输入防抖

  function uploadApi() { return root.SamplingUpload || null; }

  // 弹窗底部三个「必须先选中项目」的按钮（采样方案导入 / 调查表导入 / 工艺、点位图保存）
  const BLF_ACTION_IDS = ['blf-do', 'blf-survey', 'blf-export'];
  function setBlfButtons(on) {
    BLF_ACTION_IDS.forEach((id) => { const b = $(id); if (b) b.disabled = !on; });
  }

  function setBlfSearching(on) {
    const el = $('blf-searching');
    if (el) el.classList.toggle('hidden', !on);
  }

  function renderBlfSuggest(list) {
    const box = $('blf-suggest');
    if (!box) return;
    box.innerHTML = '';
    if (!list || !list.length) {
      const empty = document.createElement('div');
      empty.className = 'suggest-empty';
      empty.textContent = '未找到匹配的项目';
      box.appendChild(empty);
      box.classList.remove('hidden');
      return;
    }
    list.forEach((p) => {
      const item = document.createElement('div');
      item.className = 'suggest-item';
      const codeSpan = document.createElement('span');
      codeSpan.className = 'suggest-code';
      codeSpan.textContent = p.code || '';
      const nameSpan = document.createElement('span');
      nameSpan.className = 'suggest-name';
      nameSpan.textContent = p.belongInspectName || '';
      item.appendChild(codeSpan);
      item.appendChild(nameSpan);
      item.addEventListener('click', () => selectBlfProject(p));
      box.appendChild(item);
    });
    box.classList.remove('hidden');
  }

  function selectBlfProject(p) {
    blfProject = p;
    const u = $('blf-unitName'); if (u) u.value = '';
    const c = $('blf-projectCode'); if (c) c.value = '';
    const box = $('blf-suggest'); if (box) box.classList.add('hidden');
    const info = $('blf-projectInfo');
    if (info) info.textContent = '已选择项目：' + (p.code || '') + '（' + (p.belongInspectName || '') + '）';
    setBlfButtons(true);
  }

  function doBlfSearch() {
    const u = $('blf-unitName'), c = $('blf-projectCode'), y = $('blf-yearSelect');
    const box = $('blf-suggest'), info = $('blf-projectInfo');
    if (!u || !c || !box) return;
    const unitName = u.value.trim();
    const code = c.value.trim();
    const year = y ? y.value : '';
    if (!unitName && !code && !year) { box.classList.add('hidden'); return; }
    const api = uploadApi();
    if (!api || typeof api.searchProjectsFor !== 'function') return;

    blfProject = null;
    if (info) info.textContent = '';
    setBlfButtons(false);

    const seq = ++blfSeq;
    if (blfAbort) blfAbort.abort();
    blfAbort = new AbortController();
    setBlfSearching(true);
    // 过滤与截断由 SamplingUpload.searchProjectsFor 完成（与「数据上传」共用 searchCache / yearCache）
    api.searchProjectsFor({ code: code, unitName: unitName, year: year }, blfAbort.signal)
      .then((list) => {
        if (seq !== blfSeq) return;
        renderBlfSuggest(list || []);
      })
      .catch((e) => {
        if (seq !== blfSeq) return;
        box.classList.add('hidden');
        console.warn('[网站数据导入] 搜索失败：', e);
      })
      .finally(() => { if (seq === blfSeq) setBlfSearching(false); });
  }

  // ---------------- 入口 ----------------

  // 打开「网站数据导入」弹窗（先搜索并选中项目，才能导入）
  function openBlfModal() {
    const api = uploadApi();
    const sess = api && typeof api.getUpstreamContext === 'function' ? api.getUpstreamContext() : null;
    if (!sess || !sess.loggedIn) {
      alert('尚未登录平台。\n\n请先点击顶部工具栏的「数据上传」登录后，再进行网站数据导入。');
      return;
    }
    blfProject = null;
    const u = $('blf-unitName'); if (u) u.value = '';
    const c = $('blf-projectCode'); if (c) c.value = '';
    const box = $('blf-suggest'); if (box) { box.classList.add('hidden'); box.innerHTML = ''; }
    const info = $('blf-projectInfo'); if (info) info.textContent = '';
    setBlfButtons(false);
    const modal = $('blf-modal'); if (modal) modal.classList.remove('hidden');
    if (u) u.focus();
  }

  function closeBlfModal() {
    const modal = $('blf-modal'); if (modal) modal.classList.add('hidden');
    if (blfAbort) { blfAbort.abort(); blfAbort = null; }
    setBlfSearching(false);
  }

  // kind = 'plan'   采样方案导入：只取「劳动定员…调查」→ 只写主表格（测点布局）
  // kind = 'survey' 调查表导入：只取 6 张现场调查表（含工艺图）→ 只写调查表页签
  // 两者各自生成**只含自己那几张工作表**的导入模板，因此互不干扰。
  async function runImport(project, kind) {
    const isSurvey = kind === 'survey';
    const btn = $(isSurvey ? 'blf-survey' : 'blf-do');
    const old = btn ? btn.textContent : '';
    if (btn) btn.textContent = '导入中…';
    setBlfButtons(false);   // 导入期间禁用底部按钮，避免两个导入并发
    try {
      const api = uploadApi();
      const ctx = api.getUpstreamContext(project);
      if (!ctx || !ctx.loggedIn) { alert('登录状态已失效，请重新登录后再试。'); return; }

      const up = await fetchUpstream(ctx, { worker: !isSurvey, survey: isSurvey });
      if (up.errors && up.errors.length) console.warn('[网站数据导入] 部分取数失败：', up.errors);

      const app = root.SamplingApp;
      if (!app || typeof app.importWorkbook !== 'function') {
        alert('无法导入：缺少导入能力（SamplingApp.importWorkbook）。');
        return;
      }

      // 工艺图属于「生产工艺调查」这张调查表的内容 → 只在调查表导入时处理
      let picNote = '';
      if (isSurvey) {
        if (btn) btn.textContent = '导入工艺图…';
        try {
          const pr = await importProcessPictures(ctx, up.processPictures);
          if (pr.total) {
            const parts = ['工艺图 ' + pr.ok + ' 张'];
            if (pr.skip) parts.push('已存在跳过 ' + pr.skip);
            if (pr.fail) parts.push('失败 ' + pr.fail);
            picNote = '工艺图片已导入「生产工艺调查」：' + parts.join('，') + '。';
          }
        } catch (e) { console.warn('[网站数据导入] 工艺图导入失败：', e); }
      }

      // 生成导入模板 → 直接走网页的「导入 Excel」逻辑（不下载文件）
      if (btn) btn.textContent = '导入中…';
      const X = await ensureXlsx();
      const bytes = await X.writeWorkbook(buildTemplateSheets(up, {
        worker: !isSurvey, survey: isSurvey, dicts: false,
      }));
      await app.importWorkbook(bytes, picNote);
      // 导入成功后不自动关闭弹窗：保留已选项目，方便接着做另一种导入（采样方案 / 调查表）
    } catch (e) {
      console.error(e);
      alert((isSurvey ? '调查表导入失败：' : '采样方案导入失败：') + (e && e.message ? e.message : e));
    } finally {
      if (btn) btn.textContent = old || '导入';
      setBlfButtons(!!blfProject);
    }
  }

  // 注：SamplingBlf 的完整导出在文件末尾统一给出（含生产接口 + 测试钩子），此处不再重复赋值。

  // 绑定按钮（主工具栏 + 全屏工具栏）与弹窗内交互
  function bind() {
    ['btn-blf', 'fs-blf'].forEach((id) => {
      const b = $(id);
      if (b && !b.__blfBound) {
        b.__blfBound = true;
        b.addEventListener('click', openBlfModal);
      }
    });
    const closeBtn = $('blf-close'); if (closeBtn) closeBtn.addEventListener('click', closeBlfModal);
    const cancelBtn = $('blf-cancel'); if (cancelBtn) cancelBtn.addEventListener('click', closeBlfModal);
    // 页签栏「生成检评报告」（有 class=tab 但无 data-tab，不参与页签切换）
    const pjbgBtn = $('btn-pjbg');
    if (pjbgBtn && !pjbgBtn.__pjbgBound) {
      pjbgBtn.__pjbgBound = true;
      pjbgBtn.addEventListener('click', generatePjbgReport);
    }
    const doBtn = $('blf-do');
    if (doBtn) doBtn.addEventListener('click', () => {
      if (!blfProject) { alert('请先在上方搜索并选中项目，再导入采样方案。'); return; }
      runImport(blfProject, 'plan');
    });
    const svBtn = $('blf-survey');
    if (svBtn) svBtn.addEventListener('click', () => {
      if (!blfProject) { alert('请先在上方搜索并选中项目，再导入调查表。'); return; }
      runImport(blfProject, 'survey');
    });
    const expBtn = $('blf-export');
    if (expBtn) expBtn.addEventListener('click', () => {
      if (!blfProject) { alert('请先在上方搜索并选中项目，再导出工艺、点位图。'); return; }
      exportFlowAndLayout(blfProject);
    });
    const modal = $('blf-modal');
    if (modal) modal.addEventListener('click', (e) => { if (e.target.id === 'blf-modal') closeBlfModal(); });
    const u = $('blf-unitName'), c = $('blf-projectCode');
    const onInput = () => { clearTimeout(blfDebounce); blfDebounce = setTimeout(doBlfSearch, 350); };
    if (u) u.addEventListener('input', onInput);
    if (c) c.addEventListener('input', onInput);
    const y = $('blf-yearSelect');
    if (y) y.addEventListener('change', doBlfSearch);
  }
  if (document.readyState === 'loading') document.addEventListener('DOMContentLoaded', bind);
  else bind();

  // 测试/复用钩子：后处理相关纯函数（供 Node 测试直接调用，不依赖 DOM）
  root.SamplingBlf = {
    // 生产接口（页面调用）
    open: openBlfModal,
    buildTemplateSheets: buildTemplateSheets,
    fetchUpstream: fetchUpstream,
    workerTo39: workerTo39,
    exportFlowAndLayout: exportFlowAndLayout,
    generatePjbgReport: generatePjbgReport,
    genPjbgDocx: genPjbgDocx,
    fixHighToxColumn: fixHighToxColumn,
    fixNonNoiseHazards: fixNonNoiseHazards,
    fixQuietPointHazards: fixQuietPointHazards,
    fixNitrogenOxide: fixNitrogenOxide,
    autoNumberItemLists: autoNumberItemLists,
    buildNumberedParas: buildNumberedParas,
    buildNumberingXml: buildNumberingXml,
    insertPPrNumPr: insertPPrNumPr,
    mergePpeRows: mergePpeRows,
    mergePpeJobText: mergePpeJobText,
    // ⑨ 表3-3 下方插入「注1~注4」（整段标红）；供测试/复用
    addMaterialNote: addMaterialNote,
    PJBG_MATERIAL_NOTES: PJBG_MATERIAL_NOTES,
    // ⑩ 表8-2 超标原因分析：按检测项目填「超标原因分析」列
    fillOverLimitCause: fillOverLimitCause,
    PJBG_OVER_LIMIT_TEXTS: PJBG_OVER_LIMIT_TEXTS,
    fixReportText: fixReportText,
    fixHazardNote: fixHazardNote,
    collectHazardMarks: collectHazardMarks,
    trimNote3: trimNote3,
    patchPjbgDocx: patchPjbgDocx,
    repackDocxFast: repackDocxFast,   // 供测试/基准：只重压 document.xml 的快速重打包
    // 后处理⑥（条件删除条目）
    fixConditionalDeletions: fixConditionalDeletions,
    collectConfirmColTexts: collectConfirmColTexts,
    lookupHazardField: lookupHazardField,
    deletePjbgParas: deletePjbgParas,
    // 后处理⑦（「两班制」作息说明句标红）
    markShiftNoteRed: markShiftNoteRed,
    colorParaRange: colorParaRange,
    setRPrColor: setRPrColor,
    // 测试用纯函数 / 常量
    fixNoteAndContactCount: fixNoteAndContactCount,
    sumExposedWorkers: sumExposedWorkers,
    sumExposedFromT41: sumExposedFromT41,
    isNoHazardText: isNoHazardText,
    readBasicInfoTable: readBasicInfoTable,
    fixNoteParagraph: fixNoteParagraph,
    NOX_FROM_HALF: NOX_FROM_HALF,
    NOX_TO: NOX_TO,
    NOX_FROM_FULL: NOX_FROM_FULL,
    NOX_TO_FULL: NOX_TO_FULL,
    NOTE_TPL: NOTE_TPL,
    NOTE_PARA_A: NOTE_PARA_A,
    NOTE_PARA_B_TPL: NOTE_PARA_B_TPL,
  };
})(typeof self !== 'undefined' ? self : this);
