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
    // 登录失效（401）：交给统一处理（提示 + 引导重新登录）
    if (r.status === 401) {
      const u = root.SamplingUpload;
      if (u && typeof u.notifyAuthExpired === 'function') u.notifyAuthExpired();
    }
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

  // 取一个 <w:tc>…</w:tc> 里的可见文本。
  // <w:br/> 的两种写法（与 <w:t> 并列 / 写在 <w:t> 内部）都还原成换行。
  function cellText(tcXml) {
    let s = String(tcXml === null || tcXml === undefined ? '' : tcXml);
    s = s.replace(/<w:br(?:\s[^>]*)?\/?>/g, '\u0001').replace(/<w:tab\s*\/?>/g, '\t');
    // 只保留 <w:t> 边界，其余标签（含属性）丢弃
    s = s.replace(/<(?!\/?w:t(?:\s|>))[^>]*>/g, '');
    s = s.replace(/<\/?w:t(?:\s[^>]*)?>/g, '');
    return s
      .replace(/\u0001/g, '\n')
      .replace(/&lt;/g, '<').replace(/&gt;/g, '>').replace(/&quot;/g, '"')
      .replace(/&apos;/g, "'").replace(/&amp;/g, '&')
      .trim();
  }

  // 把单元格文本改成 text：第一个 <w:t> 承载内容，其余清空（保留段落结构）
  function setCellText(tcXml, text) {
    let first = true;
    return tcXml.replace(/<w:t(?:\s[^>]*)?>[\s\S]*?<\/w:t>/g, function () {
      if (first) { first = false; return '<w:t xml:space="preserve">' + text + '</w:t>'; }
      return '<w:t xml:space="preserve"></w:t>';
    });
  }

  // 在 document.xml 里找表头含「高毒物品」的表，把该列值为 0/1 的单元格换成 否/是。
  // 表头可能被拆成多个 run / 带软换行（如「存在高\n毒物品」），故用**包含**匹配而非全等。
  // 返回 { xml, changed }。
  function fixHighToxColumn(xml) {
    if (xml.indexOf(HIGHTOX_HEAD) < 0) return { xml: xml, changed: 0 };
    let changed = 0;
    const out = xml.replace(/<w:tbl(?:\s[^>]*)?>[\s\S]*?<\/w:tbl>/g, function (tbl) {
      const rows = tbl.match(/<w:tr(?:\s[^>]*)?>[\s\S]*?<\/w:tr>/g) || [];
      if (rows.length < 2) return tbl;
      const headCells = rows[0].match(/<w:tc(?:\s[^>]*)?>[\s\S]*?<\/w:tc>/g) || [];
      let col = -1;
      for (let i = 0; i < headCells.length; i++) {
        if (cellText(headCells[i]).indexOf(HIGHTOX_HEAD) >= 0) { col = i; break; }
      }
      if (col < 0) return tbl;   // 不是这张表
      let newTbl = tbl;
      for (let ri = 1; ri < rows.length; ri++) {
        const cells = rows[ri].match(/<w:tc(?:\s[^>]*)?>[\s\S]*?<\/w:tc>/g) || [];
        if (col >= cells.length) continue;
        const t = cellText(cells[col]);
        if (t !== '0' && t !== '1') continue;
        const newRow = rows[ri].replace(cells[col], setCellText(cells[col], t === '1' ? '是' : '否'));
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

  function normKey(s) { return String(s === null || s === undefined ? '' : s).replace(/\s+/g, ''); }

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
    const out = xml.replace(/<w:tbl(?:\s[^>]*)?>[\s\S]*?<\/w:tbl>/g, function (tbl) {
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

  // ---------------- 报告后处理③：正文措辞替换 ----------------
  // 平台模板里的几处说法要改成规范表述（用户要求，2026-10-12）。
  // ⚠ 目标文字常被拆到多个 run（如「7.2物理有害因素检测结果分析与评价」= ['7.2','物理','有害因素…']），
  //   因此按「段落内跨 run」替换，并尽量保留原有格式（替换文字落在起始 run 上）。
  const TEXT_FIXES = [
    { from: '工作场所粉尘和化学有害因素职业接触限值', to: '工作场所空气中化学有害因素职业接触限值' },
    { from: '工作场所空气中粉尘及化学有害因素职业接触限值', to: '工作场所空气中化学有害因素职业接触限值' },
    { from: '物理有害因素检测结果分析与评价', to: '物理因素检测结果分析与评价' },
  ];

  // 在**单个段落** XML 内做跨 run 的文本替换；命中返回新 XML，未命中返回 null
  function replaceAcrossRuns(pXml, from, to) {
    const re = /<w:t(?:\s[^>]*)?>[\s\S]*?<\/w:t>/g;
    const nodes = [];
    let m;
    while ((m = re.exec(pXml))) {
      const inner = /<w:t(?:\s[^>]*)?>([\s\S]*?)<\/w:t>/.exec(m[0]);
      nodes.push(inner ? inner[1] : '');
    }
    if (!nodes.length) return null;
    let full = '';
    const offs = nodes.map(function (t) { const s = full.length; full += t; return s; });
    const idx = full.indexOf(from);
    if (idx < 0) return null;
    const end = idx + from.length;
    let hit = false;
    const news = nodes.map(function (t, i) {
      const s = offs[i], e = s + t.length;
      if (e <= idx || s >= end) return null;   // 不涉及
      hit = true;
      if (s <= idx && e >= end) return t.slice(0, idx - s) + to + t.slice(end - s);  // 整段落在该 run 内
      if (s <= idx && idx < e) return t.slice(0, idx - s) + to;                      // 起始 run
      if (s < end && end <= e) return t.slice(end - s);                              // 结束 run
      return '';                                                                     // 完全被覆盖
    });
    if (!hit) return null;
    let k = 0;
    return pXml.replace(re, function (raw) {
      const nv = news[k++];
      if (nv === null) return raw;
      return raw.replace(/(<w:t(?:\s[^>]*)?>)[\s\S]*?(<\/w:t>)/, function (_, a, b) { return a + nv + b; });
    });
  }

  // 对 document.xml 逐段落做全部文字替换，返回 { xml, changed }
  function fixReportText(xml) {
    let changed = 0;
    const out = xml.replace(/<w:p(?:\s[^>]*)?>[\s\S]*?<\/w:p>/g, function (pXml) {
      let cur = pXml;
      for (const f of TEXT_FIXES) {
        const np = replaceAcrossRuns(cur, f.from, f.to);
        if (np !== null) { cur = np; changed++; }
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
    xml.replace(/<w:tbl(?:\s[^>]*)?>[\s\S]*?<\/w:tbl>/g, function (tbl) {
      if (found) return tbl;
      const rows = tbl.match(/<w:tr(?:\s[^>]*)?>[\s\S]*?<\/w:tr>/g) || [];
      if (rows.length < 2) return tbl;
      const headCells = rows[0].match(/<w:tc(?:\s[^>]*)?>[\s\S]*?<\/w:tc>/g) || [];
      const heads = headCells.map(cellText);
      const has = function (kw) { return heads.some(function (h) { return h.indexOf(kw) >= 0; }); };
      const hIdx = heads.findIndex(function (h) { return h.indexOf('备注') >= 0; });
      // 化学有害因素接触限值表：表头同时含「中文名」「临界不良健康效应」「备注」
      if (hIdx < 0 || !has('中文名') || !has('临界不良健康效应')) return tbl;
      found = {};
      for (let ri = 1; ri < rows.length; ri++) {
        const cells = rows[ri].match(/<w:tc(?:\s[^>]*)?>[\s\S]*?<\/w:tc>/g) || [];
        // 表头可能因合并单元格比数据行少若干格，按差值平移
        const ci = hIdx + (cells.length - headCells.length);
        if (ci < 0 || ci >= cells.length) continue;
        const v = cellText(cells[ci]);
        NOTE3_MARKS.forEach(function (it) { if (v.indexOf(it.mark) >= 0) found[it.mark] = 1; });
      }
      return tbl;
    });
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
  function fixHazardNote(xml) {
    const marks = collectHazardMarks(xml);
    if (!marks) return { xml: xml, changed: 0 };
    let changed = 0;
    let deleted = false;   // 是否已整段删除注3
    let removedNo = 0;     // 被删的注号
    const out = xml.replace(/<w:p(?:\s[^>]*)?>[\s\S]*?<\/w:p>/g, function (pXml) {
      const full = cellText(pXml);
      if (deleted) {
        if (!removedNo) return pXml;
        const m = /^(注\d+)(\s*[：:])/.exec(full);
        if (!m) return pXml;
        const n = parseInt(m[1].slice(1), 10);
        if (!(n > removedNo)) return pXml;                  // 前面的注号不动
        const np = replaceAcrossRuns(pXml, m[0], '注' + (n - 1) + m[2]);
        if (np === null) return pXml;
        changed++;
        return np;
      }
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
    return { xml: out, changed: changed };
  }

  // 对下载到的检评报告 docx 做全部后处理（读一次 document.xml，改完统一重打包）；
  // 任何一步失败都不影响出报告（回退到上一步结果 / 原始文件）。
  async function patchPjbgDocx(buf) {
    try {
      const JSZip = await ensureJSZip();
      const zip = await JSZip.loadAsync(buf);
      const f = zip.file('word/document.xml');
      if (!f) return buf;
      let xml = await f.async('string');
      let changed = 0;
      // ① 「存在高毒物品」列 0/1 → 否/是
      const r1 = fixHighToxColumn(xml);
      xml = r1.xml; changed += r1.changed;
      if (r1.changed) console.info('[检评报告] 「存在高毒物品」列 0/1 → 否/是，共修正 ' + r1.changed + ' 处');
      // ② 非噪声岗位（LEX 最大值 < 80 dB）删掉「噪声」危害
      const u = root.SamplingUpload;
      if (u && typeof u.noiseNonNoiseFromBuf === 'function') {
        const pairs = await u.noiseNonNoiseFromBuf(buf);
        const r2 = fixNonNoiseHazards(xml, pairs);
        xml = r2.xml; changed += r2.changed;
        if (r2.changed) console.info('[检评报告] 非噪声岗位删除「噪声」危害，共修正 ' + r2.changed + ' 处');
      }
      // ③ 正文措辞替换（1.4 章节的接触限值说法、7.2 标题）
      const r3 = fixReportText(xml);
      xml = r3.xml; changed += r3.changed;
      if (r3.changed) console.info('[检评报告] 正文措辞替换，共修正 ' + r3.changed + ' 处');
      // ④ 表1-1 注3：按「备注」列实际标记裁剪说明
      const r4 = fixHazardNote(xml);
      xml = r4.xml; changed += r4.changed;
      if (r4.changed) console.info('[检评报告] 表1-1 注3 按备注列标记裁剪，共修正 ' + r4.changed + ' 处');
      if (!changed) return buf;
      zip.file('word/document.xml', xml);
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

  root.SamplingBlf = { open: openBlfModal, buildTemplateSheets, fetchUpstream, workerTo39, exportFlowAndLayout, generatePjbgReport, genPjbgDocx, fixHighToxColumn, fixNonNoiseHazards, fixReportText, fixHazardNote, collectHazardMarks, trimNote3, patchPjbgDocx };

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
})(typeof self !== 'undefined' ? self : this);
