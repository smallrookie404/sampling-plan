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

  // 上游字段在新旧版本中会改名（如 belongUnitName ↔ belongUnit、facilityOperations ↔ useNum）：
  // 按候选顺序取第一个非空值，做到新旧版兼容
  function pick(r, keys) {
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

  // ---------------- 组装：导入模板（20 页） ----------------

  function buildTemplateSheets(up) {
    const D = data();
    const widthsTpl = D.widthsTemplate || {};
    const localSv = surveyData();
    const useUp = !!(up && (up.worker39 || up.survey));

    const sheets = [];
    // 1) 劳动定员和职业病危害因素接触情况调查（39 列，取自上游）
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

  // ---------------- 项目搜索（移植自「数据上传」搜索模块） ----------------

  let blfProject = null;   // 本弹窗独立选中的项目
  let blfSeq = 0;          // 搜索序号（丢弃过期响应）
  let blfAbort = null;     // 进行中的搜索请求
  let blfDebounce = null;  // 输入防抖

  function uploadApi() { return root.SamplingUpload || null; }

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
    const btn = $('blf-do'); if (btn) btn.disabled = false;
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
    const doBtn = $('blf-do'); if (doBtn) doBtn.disabled = true;

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
    const doBtn = $('blf-do'); if (doBtn) doBtn.disabled = true;
    const modal = $('blf-modal'); if (modal) modal.classList.remove('hidden');
    if (u) u.focus();
  }

  function closeBlfModal() {
    const modal = $('blf-modal'); if (modal) modal.classList.add('hidden');
    if (blfAbort) { blfAbort.abort(); blfAbort = null; }
    setBlfSearching(false);
  }

  async function runImport(project) {
    const btn = $('blf-do');
    const old = btn ? btn.textContent : '';
    if (btn) { btn.disabled = true; btn.textContent = '导入中…'; }
    try {
      const api = uploadApi();
      const ctx = api.getUpstreamContext(project);
      if (!ctx || !ctx.loggedIn) { alert('登录状态已失效，请重新登录后再试。'); return; }

      const up = await fetchUpstream(ctx);
      if (up.errors && up.errors.length) console.warn('[网站数据导入] 部分取数失败：', up.errors);

      const app = root.SamplingApp;
      if (!app || typeof app.importWorkbook !== 'function') {
        alert('无法导入：缺少导入能力（SamplingApp.importWorkbook）。');
        return;
      }

      // 生成导入模板 → 直接走网页的「导入 Excel」逻辑（不下载文件）
      const X = await ensureXlsx();
      const bytes = await X.writeWorkbook(buildTemplateSheets(up));
      await app.importWorkbook(bytes);
      closeBlfModal();
    } catch (e) {
      console.error(e);
      alert('网站数据导入失败：' + (e && e.message ? e.message : e));
    } finally {
      if (btn) { btn.disabled = false; btn.textContent = old || '导入'; }
    }
  }

  root.SamplingBlf = { open: openBlfModal, buildTemplateSheets, fetchUpstream, workerTo39 };

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
    const doBtn = $('blf-do');
    if (doBtn) doBtn.addEventListener('click', () => { if (blfProject) runImport(blfProject); });
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
})(typeof self !== 'undefined' ? self : this);
