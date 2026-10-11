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
    const out = { worker: null, worker39: null, survey: {}, processPictures: [], errors: [] };
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

  // ---------------- 工艺、点位图导出（生成检评报告 Word → 提取内嵌图片） ----------------
  // 两个导出目标：
  //   ① 工艺流程图：报告「3.2.1 生产工艺」小节内的图片（用图注「图3-x …」命名）
  //   ② 采样/测量布点图：报告「附录2 采样/测量布点图」小节内的图片

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

  function saveBlobAs(blob, name) {
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = name;
    document.body.appendChild(a);
    a.click();
    setTimeout(() => { URL.revokeObjectURL(url); a.remove(); }, 5000);
  }

  // 解析「编号标题」（3.2.1生产工艺 / 3.2 生产工艺 / 9建议 / 10附件）
  // 或「附录标题」（附录1 xxx / 附录2采样/测量布点图）。
  // 注意：Word 里编号常被拆成多个 run（document.xml 里搜不到「3.2.1」），
  // 必须用段落 textContent 拼回整段后再匹配。
  function parseHeading(raw) {
    const s = String(raw || '').trim();
    if (!s || s.length > 50) return null;
    // 附录标题
    let m = s.match(/^附\s*录\s*([0-9０-９一二三四五六七八九十]+)\s*([\s\S]*)$/);
    if (m) {
      const rest0 = m[2].trim();
      if (rest0.length <= 30) return { kind: 'appendix', num: m[1], rest: rest0, depth: 1, text: s };
    }
    // 编号标题
    m = s.match(/^(\d+(?:[.．]\d+)*)([、.．\s]*)([\s\S]*)$/);
    if (!m) return null;
    const num = m[1].replace(/．/g, '.');
    const sep = m[2] || '';
    const rest = m[3].trim();
    if (!rest || rest.length > 24) return null;
    const hierarchical = num.indexOf('.') >= 0;                 // 3.2 / 3.2.1
    // 无分隔符时只接受「1~2 位数字 + 非数字开头」的短标题（9建议 / 10附件），
    // 避免把「2018年12月…」这类正文误判为标题
    const shortNoSep = !hierarchical && !sep.length && num.length <= 2 && !/^\d/.test(rest);
    if (!hierarchical && !sep.length && !shortNoSep) return null;
    return { kind: 'num', num: num, rest: rest, depth: num.split('.').length, text: s };
  }

  // 解析 docx 段落 → [{ text, images:[media 路径,…] }]
  async function readDocxParagraphs(zip) {
    const docFile = zip.file('word/document.xml');
    const relFile = zip.file('word/_rels/document.xml.rels');
    if (!docFile || !relFile) return null;
    const docXml = await docFile.async('string');
    const relXml = await relFile.async('string');
    const P = new DOMParser();
    const W_NS = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
    const A_NS = 'http://schemas.openxmlformats.org/drawingml/2006/main';
    const R_NS = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
    const relMap = {};
    const relDom = P.parseFromString(relXml, 'application/xml');
    const rels = relDom.getElementsByTagName('Relationship');
    for (let i = 0; i < rels.length; i++) {
      const id = rels[i].getAttribute('Id');
      const tgt = rels[i].getAttribute('Target');
      if (id && tgt) relMap[id] = 'word/' + tgt.replace(/^\.\//, '').replace(/^\//, '');
    }
    const dom = P.parseFromString(docXml, 'application/xml');
    const body = dom.getElementsByTagNameNS(W_NS, 'body')[0];
    if (!body) return null;
    const ps = body.getElementsByTagNameNS(W_NS, 'p');
    const out = [];
    for (let i = 0; i < ps.length; i++) {
      const p = ps[i];
      const images = [];
      const blips = p.getElementsByTagNameNS(A_NS, 'blip');
      for (let k = 0; k < blips.length; k++) {
        const id = blips[k].getAttributeNS(R_NS, 'embed') || blips[k].getAttribute('r:embed');
        if (id && relMap[id]) images.push(relMap[id]);
      }
      out.push({ text: String(p.textContent || '').trim(), images: images });
    }
    return out;
  }

  // 按标题挑选一个小节（匹配多个时取「rest 更短 / 层级更深」者），返回该小节内的图片与标题文本
  function pickSection(paras, match) {
    let target = null;
    for (let i = 0; i < paras.length; i++) {
      const h = parseHeading(paras[i].text);
      if (!h || !match(h)) continue;
      if (!target) { target = { i: i, h: h }; continue; }
      const t = target.h;
      if (h.rest.length < t.rest.length || (h.rest.length === t.rest.length && h.depth > t.depth)) target = { i: i, h: h };
    }
    if (!target) return { items: [], headText: '' };
    // 终点：下一个同级或更浅的标题
    let end = paras.length;
    for (let j = target.i + 1; j < paras.length; j++) {
      const h = parseHeading(paras[j].text);
      if (h && h.depth <= target.h.depth) { end = j; break; }
    }
    const items = [];
    for (let j = target.i + 1; j < end; j++) {
      const imgs = paras[j].images;
      if (!imgs || !imgs.length) continue;
      let caption = '';
      for (let k = j + 1; k < Math.min(j + 4, end); k++) {   // 图注：紧随其后、以「图」开头的短段
        const ct = paras[k].text;
        if (!ct) continue;
        if (/^图/.test(ct) && ct.length <= 60) caption = ct;
        break;
      }
      for (const p of imgs) items.push({ path: p, caption: caption });
    }
    return { items: items, headText: target.h.text };
  }

  async function readZipImages(zip, items) {
    const out = [];
    for (const it of items) {
      const f = zip.file(it.path);
      if (!f) continue;
      out.push({
        name: it.path.split('/').pop(),
        caption: it.caption,
        mime: /\.png$/i.test(it.path) ? 'image/png' : (/\.jpe?g$/i.test(it.path) ? 'image/jpeg' : 'application/octet-stream'),
        data: await f.async('uint8array'),
      });
    }
    return out;
  }

  // 生成检评报告 Word：模板列表 api/pjbgModule → /api/jpbgReportData/jpbgYL → 同路径 .docx
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
    return { buf: await resp.arrayBuffer(), modName: mod.name || '' };
  }

  async function exportFlowAndLayout(project) {
    const api = uploadApi();
    const ctx = api && typeof api.getUpstreamContext === 'function' ? api.getUpstreamContext(project) : null;
    if (!ctx || !ctx.loggedIn) { alert('登录状态已失效，请重新登录后再试。'); return; }
    const btn = $('blf-export');
    const old = btn ? btn.textContent : '';
    if (btn) { btn.disabled = true; btn.textContent = '生成报告中…'; }
    try {
      const g = await genPjbgDocx(ctx);
      if (btn) btn.textContent = '提取图片中…';
      const JSZip = await ensureJSZip();
      const zip = await JSZip.loadAsync(g.buf);
      const paras = await readDocxParagraphs(zip);
      if (!paras) throw new Error('报告结构无法解析');

      const flow = pickSection(paras, (h) => h.kind === 'num' && /^生产工艺/.test(h.rest));
      const layout = pickSection(paras, (h) => h.kind === 'appendix' && (/布点/.test(h.rest) || /^(2|二)$/.test(String(h.num).trim())));
      const flowImgs = await readZipImages(zip, flow.items);
      const layoutImgs = await readZipImages(zip, layout.items);
      if (!flowImgs.length && !layoutImgs.length) {
        alert('未在报告里找到工艺流程图或布点图。\n\n请确认报告包含「3.2.1 生产工艺」与「附录2 采样/测量布点图」。');
        return;
      }
      const base = String((ctx.project && ctx.project.code) || 'report').replace(/[\\/:*?"<>|]/g, '_');
      const safe = (s) => String(s || '').replace(/[\\/:*?"<>|]/g, '_').replace(/\s+/g, ' ').trim();
      const files = [];
      flowImgs.forEach((im, i) => {
        const ext = /jpe?g$/i.test(im.name) ? 'jpg' : 'png';
        const nm = im.caption ? safe(im.caption) : (base + '-工艺流程图' + (flowImgs.length > 1 ? ('-' + (i + 1)) : ''));
        files.push({ im: im, nm: nm + '.' + ext });
      });
      layoutImgs.forEach((im, i) => {
        const ext = /jpe?g$/i.test(im.name) ? 'jpg' : 'png';
        files.push({ im: im, nm: base + '-采样测量布点图' + (layoutImgs.length > 1 ? ('-' + (i + 1)) : '') + '.' + ext });
      });
      for (const f of files) {
        saveBlobAs(new Blob([f.im.data], { type: f.im.mime }), f.nm);
        await new Promise((r) => setTimeout(r, 400));
      }
      alert('已从「' + g.modName + '」报告导出 ' + files.length + ' 张图片：\n' +
        '· 工艺流程图 ' + flowImgs.length + ' 张' + (flow.headText ? '（' + flow.headText + '）' : '') + '\n' +
        '· 采样/测量布点图 ' + layoutImgs.length + ' 张' + (layout.headText ? '（' + layout.headText + '）' : '') +
        '\n\n文件已保存到浏览器下载目录。');
    } catch (e) {
      console.error(e);
      alert('导出图片失败：' + (e && e.message ? e.message : e));
    } finally {
      if (btn) { btn.disabled = false; btn.textContent = old || '工艺、点位图导出'; }
    }
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

      // 平台上的工艺图 → 「生产工艺调查」对应工艺名称的工艺图槽位（先写入，随导入一起渲染出来）
      let picNote = '';
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

      // 生成导入模板 → 直接走网页的「导入 Excel」逻辑（不下载文件）
      if (btn) btn.textContent = '导入中…';
      const X = await ensureXlsx();
      const bytes = await X.writeWorkbook(buildTemplateSheets(up));
      await app.importWorkbook(bytes, picNote);
      closeBlfModal();
    } catch (e) {
      console.error(e);
      alert('网站数据导入失败：' + (e && e.message ? e.message : e));
    } finally {
      if (btn) { btn.disabled = false; btn.textContent = old || '导入'; }
    }
  }

  root.SamplingBlf = { open: openBlfModal, buildTemplateSheets, fetchUpstream, workerTo39, exportFlowAndLayout };
  // 供测试/复用：docx 解析与「按标题取小节图片」
  root.__samplingBlfDocx = { parseHeading: parseHeading, pickSection: pickSection, readDocxParagraphs: readDocxParagraphs, readZipImages: readZipImages };

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
})(typeof self !== 'undefined' ? self : this);
