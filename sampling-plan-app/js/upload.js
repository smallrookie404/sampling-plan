/* =====================================================================
 * 现场调查 Excel 上传（整合进采样计划软件）
 * - 登录界面（账号/密码/验证码）作为整个软件的总登录门禁；
 * - 「数据上传」按钮切换到本视图；
 * - 上传文件自动使用当前表格导出的 Excel（无需手动选择）。
 * 原项目：zhiyeweishen/site-survey-upload-web
 * ===================================================================== */
(function () {
  'use strict';

  // 共享工具（转义 / 脚本懒加载 / 列表结构兼容）：见 js/util.js
  const Util = window.SamplingUtil;

  // ------------------- 配置区 -------------------
  // HTTPS 页面（Cloudflare Pages 部署，如 sampling-plan.pages.dev）走同源代理 /api/platform，
  // 避免浏览器混合内容拦截 http:// 平台接口；本地 http/file 打开时直连平台。
  const API_BASE = (function () {
    try {
      if (typeof location !== 'undefined' && location.protocol === 'https:' && /pages\.dev$/i.test(location.hostname)) {
        return '/api/platform';
      }
    } catch (e) {}
    return 'http://223.93.144.122:27800';
  })();
  const RSA_PUB_B64 = 'MFwwDQYJKoZIhvcNAQEBBQADSwAwSAJBANL378k3RiZHWx5AfJqdH9xRNBmD9wGD2iRe41HdTNF8RUhNnHit5NpMNtGL0NPTSSpPjjI1kJfVorRvaQerUgkCAwEAAQ==';
  const SESSION_KEY = 'xcdc_session_v1';       // 登录会话（sessionStorage，关闭浏览器失效）
  const REM_KEY = 'xcdc_upload_remember';
  const USR_KEY = 'xcdc_upload_username';
  const PWD_KEY = 'xcdc_upload_password';
  // 登录失效（401）时的统一处理：由 initApp 注入（提示 + 引导重新登录）
  let authExpiredHandler = null;
  // ----------------------------------------------

  // ---------------- RSA (PKCS#1 v1.5) ----------------
  function b64ToBytes(b64) {
    const bin = atob(b64);
    const bytes = new Uint8Array(bin.length);
    for (let i = 0; i < bin.length; i++) bytes[i] = bin.charCodeAt(i);
    return bytes;
  }

  function bytesToB64(bytes) {
    let bin = '';
    for (let i = 0; i < bytes.length; i++) bin += String.fromCharCode(bytes[i]);
    return btoa(bin);
  }

  function bytesToBigInt(bytes) {
    let hex = '0x';
    for (let i = 0; i < bytes.length; i++) {
      hex += bytes[i].toString(16).padStart(2, '0');
    }
    return BigInt(hex);
  }

  function bigIntToBytes(x, len) {
    const bytes = new Uint8Array(len);
    for (let i = len - 1; i >= 0 && x > 0n; i--) {
      bytes[i] = Number(x & 0xffn);
      x >>= 8n;
    }
    return bytes;
  }

  function modPow(base, exp, mod) {
    let r = 1n;
    base %= mod;
    while (exp > 0n) {
      if (exp & 1n) r = (r * base) % mod;
      base = (base * base) % mod;
      exp >>= 1n;
    }
    return r;
  }

  function rsaEncrypt(plainText) {
    const der = b64ToBytes(RSA_PUB_B64);
    const modulus = bytesToBigInt(der.slice(25, 89));   // 64 字节模数
    const exponent = bytesToBigInt(der.slice(91, 94));  // 010001
    const k = 64; // 密钥长度（字节）
    const data = new TextEncoder().encode(plainText);
    const padLen = k - data.length - 3;
    if (padLen < 8) throw new Error('密码过长，无法加密');
    const em = new Uint8Array(k);
    em[0] = 0x00;
    em[1] = 0x02;
    const rand = crypto.getRandomValues(new Uint8Array(padLen));
    for (let i = 0; i < padLen; i++) em[2 + i] = rand[i] === 0 ? 1 : rand[i]; // 非零填充
    em[2 + padLen] = 0x00;
    em.set(data, 3 + padLen);
    const m = bytesToBigInt(em);
    const c = modPow(m, exponent, modulus);
    return bytesToB64(bigIntToBytes(c, k));
  }

  // ---------------- HTTP 封装 ----------------
  async function apiRequest(method, path, opts) {
    opts = opts || {};
    const headers = {};
    if (opts.token) headers['Authorization'] = opts.token;
    if (opts.orgId) headers['organizationId'] = String(opts.orgId);
    if (opts.belongProject) headers['belongProject'] = String(opts.belongProject);
    const init = { method: method, headers: headers };
    const ac = new AbortController();
    // timeout 为 0 表示不限时（如大文件上传，平台处理时间不可预估）
    const timer = opts.timeout === 0 ? null : setTimeout(function () { ac.abort(); }, opts.timeout || 30000);
    if (opts.signal) {
      opts.signal.addEventListener('abort', function () { ac.abort(); });
    }
    init.signal = ac.signal;
    if (opts.json !== undefined) {
      headers['Content-Type'] = 'application/json';
      init.body = JSON.stringify(opts.json);
    } else if (opts.form) {
      init.body = opts.form;
    }
    try {
      const resp = await fetch(API_BASE + path, init);
      const text = await resp.text();
      let data = null;
      try { data = JSON.parse(text); } catch (e) { data = text; }
      // 登录失效：平台对过期会话返回 401（"找不到当前登录的信息"）→ 提示并引导重新登录
      // （opts.skipAuthExpired：账号池令牌探针等内部调用使用，避免误伤主会话）
      if (resp.status === 401 && opts.token && !opts.skipAuthExpired && typeof authExpiredHandler === 'function') authExpiredHandler();
      return { status: resp.status, data: data, text: text };
    } finally {
      clearTimeout(timer);
    }
  }

  // ---------------- 平台接口 ----------------
  async function fetchCaptcha() {
    const r = await apiRequest('GET', '/auth/code');
    if (r.status !== 200 || !r.data || !r.data.img || !r.data.uuid) {
      throw new Error('获取验证码失败(HTTP ' + r.status + ')：' + r.text);
    }
    return r.data; // { uuid, img }
  }

  async function login(username, password, code, uuid) {
    return apiRequest('POST', '/auth/login', {
      json: { username: username, password: rsaEncrypt(password), code: code, uuid: uuid }
    });
  }

  // 登录后获取权威用户信息（与平台网站一致），用于解析真实机构
  async function fetchAuthInfo(token) {
    const r = await apiRequest('GET', '/auth/info', { token: token });
    if (r.status === 200 && r.data && r.data.user) return r.data.user;
    return null;
  }

  // 机构 ID 解析：优先用户顶层的 organizationId（管理员切换机构后的当前机构），
  // 其次 user.organization.organizationId
  function resolveOrgId(user) {
    if (!user) return null;
    const top = user.organizationId;
    const nested = user.organization && user.organization.organizationId;
    return top || nested || null;
  }


  const searchCache = new Map(); // 搜索结果缓存：key -> { time, data }
  const yearCache = new Map();   // 年份范围缓存：key -> { time, data }
  const searchInFlight = new Map(); // 进行中的搜索请求：key -> Promise
  const yearInFlight = new Map();   // 进行中的年份范围请求：key -> Promise
  const CACHE_TTL = 5 * 60 * 1000; // 5 分钟
  const CACHE_TTL_YEAR = 30 * 60 * 1000; // 年份项目列表缓存 30 分钟
  const CACHE_MAX = 30;
  const YEAR_SS_PREFIX = 'xcdc_year_projects_v1:';

  function yearSsKey(yearPrefix, unitName) {
    return YEAR_SS_PREFIX + (yearPrefix || '') + '|' + (unitName || '');
  }

  // 平台项目搜索：用「初版原始记录管理」同款接口 /api/reportData/findList，
  // 全年份（2026 新项目与 2026 之前老项目）全覆盖，需携带 organizationId 头
  async function fetchProjects(token, orgId, query, pageSize, signal, timeout) {
    const q = 'pageNumber=1&pageSize=' + pageSize + '&' + query;
    const r = await apiRequest('GET', '/api/reportData/findList?' + q, { token: token, orgId: orgId, signal: signal, timeout: timeout });
    if (r.status !== 200) throw new Error('搜索项目失败(HTTP ' + r.status + ')');
    return (r.data && r.data.body && r.data.body.records) || [];
  }

  async function searchProjects(token, orgId, keyword, signal) {
    const params = [];
    if (keyword && keyword.code && String(keyword.code).trim()) {
      params.push('code=' + encodeURIComponent(String(keyword.code).trim()));
    }
    if (keyword && keyword.unitName && String(keyword.unitName).trim()) {
      params.push('belongInspectName=' + encodeURIComponent(String(keyword.unitName).trim()));
    }
    // year：服务端按年度归属过滤（与 code / belongInspectName 为 AND 关系）
    if (keyword && keyword.year && String(keyword.year).trim()) {
      params.push('year=' + encodeURIComponent(String(keyword.year).trim()));
    }
    if (params.length === 0) return [];
    const cacheKey = params.join('&');
    const cached = searchCache.get(cacheKey);
    if (cached && Date.now() - cached.time < CACHE_TTL) {
      return cached.data;
    }
    if (searchInFlight.has(cacheKey)) {
      return searchInFlight.get(cacheKey);
    }
    const p = (async function () {
      // 平台接口经 Cloudflare 代理转发较慢，搜索超时放宽到 60 秒；
      // findList 接口（初版原始记录管理同款）全年份全覆盖，单接口即可
      const records = await fetchProjects(token, orgId, cacheKey, 50, signal, 60000);
      searchCache.set(cacheKey, { time: Date.now(), data: records });
      if (searchCache.size > CACHE_MAX) {
        const oldest = searchCache.keys().next().value;
        searchCache.delete(oldest);
      }
      return records;
    })();
    searchInFlight.set(cacheKey, p);
    try {
      return await p;
    } finally {
      searchInFlight.delete(cacheKey);
    }
  }

  async function fetchYearProjects(token, orgId, yearPrefix, unitName, signal) {
    const key = yearPrefix + '|' + (unitName || '');
    const cached = yearCache.get(key);
    if (cached && Date.now() - cached.time < CACHE_TTL_YEAR) {
      return cached.data;
    }
    if (yearInFlight.has(key)) {
      return yearInFlight.get(key);
    }
    // sessionStorage 兜底：页面刷新后免重新拉取整年项目
    try {
      if (typeof sessionStorage !== 'undefined') {
        const ssRaw = sessionStorage.getItem(yearSsKey(yearPrefix, unitName));
        if (ssRaw) {
          const ss = JSON.parse(ssRaw);
          if (ss && ss.time && Array.isArray(ss.data) && Date.now() - ss.time < CACHE_TTL_YEAR) {
            yearCache.set(key, ss);
            return ss.data;
          }
        }
      }
    } catch (e) {}
    const p = (async function () {
      const params = [];
      // 年份浏览改用 year 参数（findList 按年度归属过滤，比 code 前缀更完整）
      const ym = /^BTC(\d{2})$/.exec(yearPrefix);
      if (ym) params.push('year=20' + ym[1]);
      else if (yearPrefix) params.push('code=' + encodeURIComponent(yearPrefix));
      if (unitName) params.push('belongInspectName=' + encodeURIComponent(String(unitName).trim()));
      // 整年项目批量拉取可能更慢，超时放宽到 90 秒
      const records = await fetchProjects(token, orgId, params.join('&'), 2000, signal, 90000);
      const entry = { time: Date.now(), data: records };
      yearCache.set(key, entry);
      try {
        if (typeof sessionStorage !== 'undefined') {
          sessionStorage.setItem(yearSsKey(yearPrefix, unitName), JSON.stringify(entry));
        }
      } catch (e) {}
      if (yearCache.size > 20) {
        const oldest = yearCache.keys().next().value;
        yearCache.delete(oldest);
      }
      return records;
    })();
    yearInFlight.set(key, p);
    try {
      return await p;
    } finally {
      yearInFlight.delete(key);
    }
  }

  async function uploadExcel(token, orgId, projectId, file) {
    const fd = new FormData();
    fd.append('file', file);
    fd.append('projectId', String(projectId));
    return apiRequest('POST', '/api/evaluationUnit/import331?id=' + projectId, {
      token: token,
      orgId: orgId,
      belongProject: projectId,
      form: fd,
      timeout: 0 // 上传不限时：平台导入处理时间可能较长，超时中断会造成「实际成功却报异常」
    });
  }

  // 工艺图片补传：调查表导入成功后，按「工艺名称」匹配平台的生产工艺记录，
  // 逐条调 /api/productionProcess/upload?id=<记录id>（multipart 字段名 imgs）上传工艺图。
  // 返回一段可拼接的中文说明（无图片时返回空串）。
  async function uploadProcessImages(token, orgId, projectId) {
    const all = (window.SurveySheets && window.SurveySheets.getProcessImages) ? (window.SurveySheets.getProcessImages() || {}) : {};
    const names = Object.keys(all).filter(function (k) { return k && all[k]; });
    if (!names.length) return '';
    let recs = [];
    try {
      const r = await apiRequest(
        'GET',
        '/api/productionProcess?belongProject=' + encodeURIComponent(projectId) + '&pageNumber=1&pageSize=-1',
        { token: token, orgId: orgId, belongProject: projectId, timeout: 60000 }
      );
      const body = r.data && r.data.body;
      recs = (body && (body.records || body.content)) || (Array.isArray(body) ? body : []);
    } catch (e) {
      return '工艺图片未上传（查询工艺记录失败：' + e.message + '）';
    }
    let ok = 0;
    const failed = [];
    for (const nm of names) {
      const hit = recs.find(function (x) { return String((x && x.processName) || '').trim() === nm; });
      if (!hit || hit.id === undefined || hit.id === null) { failed.push(nm + '（未匹配到工艺记录）'); continue; }
      try {
        const blob = await (await fetch(all[nm])).blob();
        const fd = new FormData();
        fd.append('imgs', new File([blob], nm + '.jpg', { type: 'image/jpeg' }));
        const ur = await apiRequest('POST', '/api/productionProcess/upload?id=' + encodeURIComponent(hit.id), {
          token: token, orgId: orgId, form: fd, timeout: 0
        });
        if (ur.status === 200 && ur.data && ur.data.code === '200') ok++;
        else failed.push(nm + '（' + ((ur.data && (ur.data.message || ur.data.msg)) || ('HTTP ' + ur.status)) + '）');
      } catch (e) {
        failed.push(nm + '（' + e.message + '）');
      }
    }
    if (!failed.length) return '工艺图片已上传 ' + ok + ' 张';
    return '工艺图片：成功 ' + ok + ' 张，失败 ' + failed.length + ' 张（' + failed.slice(0, 5).join('；') + (failed.length > 5 ? ' 等' : '') + '）';
  }

  // ---------------- 企业信息（受检单位信息）：在职人数 / 接害人数 ----------------
  // 平台「现场调查 - 企业信息」页签字段：enterpriseNum=在职人数、hazardNum=接害人数。
  // 读：GET  /api/inspectedUnitInfo?id=<belongInspect>（reportData 的 belongInspect，不是项目 id）
  // 写：PUT  /api/inspectedUnitInfo/edit（整对象回写，数组字段按平台前端口径拼回字符串）
  // 2026-10-18 真机验证（BTC26-SZJC0959，id=6205/belongInspect=1450）：
  //   GET id=1450 → enterpriseNum=410/hazardNum=53；PUT 原值回写 → code:"200" 修改成功。
  var ENTINFO_JOIN = { area: '/', belongIndustry: '/', registrant: '/', jsfwField: ',' };

  function normalizeEntInfoRecord(rec) {
    var t = JSON.parse(JSON.stringify(rec));
    Object.keys(ENTINFO_JOIN).forEach(function (k) {
      if (Array.isArray(t[k])) t[k] = t[k].join(ENTINFO_JOIN[k]);
    });
    var ih = t.inspectedHelp;
    if (ih) {
      if (Array.isArray(ih.newPneumDisease)) ih.newPneumDisease = ih.newPneumDisease.join(',');
      if (Array.isArray(ih.newHxDisease)) ih.newHxDisease = ih.newHxDisease.join(',');
      if (Array.isArray(ih.meetStandardDetails)) ih.meetStandardDetails = ih.meetStandardDetails.join('/');
    }
    return t;
  }

  async function fetchEnterpriseInfo(token, orgId, belongInspect) {
    if (belongInspect === undefined || belongInspect === null || belongInspect === '') return null;
    var r = await apiRequest('GET', '/api/inspectedUnitInfo?id=' + encodeURIComponent(belongInspect), {
      token: token, orgId: orgId, timeout: 30000
    });
    if (r.status !== 200) throw new Error('获取企业信息失败(HTTP ' + r.status + ')');
    var recs = (r.data && r.data.body && r.data.body.records) || [];
    return recs.length ? recs[0] : null;
  }

  async function editEnterpriseInfo(token, orgId, record) {
    return apiRequest('PUT', '/api/inspectedUnitInfo/edit', {
      token: token, orgId: orgId, json: normalizeEntInfoRecord(record), timeout: 30000
    });
  }

  // 输入框值 → 补丁（纯函数，便于测试）：都为空 → {empty:true}；格式非法 → {error}；
  // 否则 {patch:{enterpriseNum?, hazardNum?}}（只含有值的一侧，数字串转 Number）
  function buildEntInfoPatch(values) {
    var onJob = String((values && values.onJobNum) == null ? '' : values.onJobNum).trim();
    var hazard = String((values && values.hazardNum) == null ? '' : values.hazardNum).trim();
    if (!onJob && !hazard) return { empty: true, patch: {} };
    var pairs = [['在职人数', onJob], ['接害人数', hazard]];
    for (var i = 0; i < pairs.length; i++) {
      if (pairs[i][1] && !/^\d+$/.test(pairs[i][1])) return { error: pairs[i][0] + '须为非负整数' };
    }
    var patch = {};
    if (onJob) patch.enterpriseNum = Number(onJob);
    if (hazard) patch.hazardNum = Number(hazard);
    return { patch: patch };
  }

  // 上传分工（用户约定）：主表格（测点布局调查）上传只带在职人数，调查表（6 合一）上传只带接害人数
  function entPayloadForUpload(kind, values) {
    var v = values || {};
    if (kind === 'survey') return { onJobNum: '', hazardNum: v.hazardNum };
    return { onJobNum: v.onJobNum, hazardNum: '' };
  }

  // ---------------- 企业基本信息自动补全（调查表上传时可选勾选，2026-10-20 用户需求） ----------------
  // 规则（只补空，绝不覆盖已有值）：
  //   ① 职卫管理部门 occupationalOrg 空 → 「办公室」
  //   ② 职卫负责人 occupationalPerson 空 → 抄「联系人」connectPerson；
  //      职卫负责人电话 occupationalPersonTel 空 → 抄「联系电话」connectTel
  //   ③ 危害风险 hazardRisk：一般("1") → 职卫兼职人数 occupationalPartNum 空则填 1；
  //                          严重("0") → 职卫专职人数 occupationalFullNum 空则填 1
  // ⚠ 平台计数列的默认值是 0（不是 null），「0 人」与「没填」在平台里长得一样；
  //   本模块默认把 0 也算「空」（opts.zeroAsEmpty=false 可关掉，改成只认 null/''）。
  var ENT_BASIC_DEFAULT_ORG = '办公室';

  function entValueBlank(v, zeroAsEmpty) {
    if (v === undefined || v === null) return true;
    if (typeof v === 'string' && v.trim() === '') return true;
    if (zeroAsEmpty !== false && (v === 0 || (typeof v === 'string' && v.trim() === '0'))) return true;
    return false;
  }

  // 纯函数（便于测试）：rec = 平台企业信息记录 → { patch:{字段:值}, notes:[说明…] }
  //   notes 只用于日志/弹窗文案，patch 为空表示无需补全。
  function buildEntBasicPatch(rec, opts) {
    var r = rec || {};
    var zero = !opts || opts.zeroAsEmpty !== false;
    var blank = function (v) { return entValueBlank(v, zero); };
    var patch = {};
    var notes = [];

    if (blank(r.occupationalOrg)) {
      patch.occupationalOrg = ENT_BASIC_DEFAULT_ORG;
      notes.push('职卫管理部门（空 → ' + ENT_BASIC_DEFAULT_ORG + '）');
    }
    if (blank(r.occupationalPerson)) {
      if (!blank(r.connectPerson)) {
        patch.occupationalPerson = String(r.connectPerson).trim();
        notes.push('职卫负责人（空 → 联系人 ' + patch.occupationalPerson + '）');
      } else {
        notes.push('职卫负责人（空，但联系人也是空，无法补全）');
      }
    }
    if (blank(r.occupationalPersonTel)) {
      if (!blank(r.connectTel)) {
        patch.occupationalPersonTel = String(r.connectTel).trim();
        notes.push('职卫负责人电话（空 → 联系电话 ' + patch.occupationalPersonTel + '）');
      } else {
        notes.push('职卫负责人电话（空，但联系电话也是空，无法补全）');
      }
    }

    var risk = String(r.hazardRisk == null ? '' : r.hazardRisk).trim();
    if (risk === '1') {
      if (blank(r.occupationalPartNum)) {
        patch.occupationalPartNum = 1;
        notes.push('职卫兼职人数（危害风险一般 → 空补 1）');
      }
    } else if (risk === '0') {
      if (blank(r.occupationalFullNum)) {
        patch.occupationalFullNum = 1;
        notes.push('职卫专职人数（危害风险严重 → 空补 1）');
      }
    } else {
      notes.push('危害风险未选，跳过兼职/专职人数补全');
    }
    return { patch: patch, notes: notes };
  }

  // 上传流程调用：把输入框里的在职人数/接害人数写进平台企业信息。
  // extraBuilder（可选）：function(rec) → { patch, notes }，用于「企业基本信息自动补全」，
  //   与输入框补丁**合并成一次 PUT**（避免连写两次）；即使输入框都为空，只要给了 extraBuilder 也会读一次平台。
  // 两个都为空且无 extraBuilder → 跳过（不调任何接口）；失败不阻断上传主体，只返回说明文本供日志/弹窗展示。
  async function applyEntInfoOnUpload(token, orgId, project, values, onSaved, extraBuilder) {
    var built = buildEntInfoPatch(values);
    if (built.error) return '企业信息未写入：' + built.error;
    var hasInput = !built.empty;
    var hasExtra = typeof extraBuilder === 'function';
    if (!hasInput && !hasExtra) return '';
    var belongInspect = project && project.belongInspect;
    if (belongInspect === undefined || belongInspect === null || belongInspect === '') {
      return '企业信息未写入：项目未带受检单位标识（belongInspect）';
    }
    var rec = null;
    try {
      rec = await fetchEnterpriseInfo(token, orgId, belongInspect);
    } catch (e) {
      return '企业信息未写入：' + e.message;
    }
    if (!rec || rec.id === undefined || rec.id === null) {
      return '企业信息未写入：平台无该受检单位记录（id=' + belongInspect + '）';
    }
    // 输入框补丁优先（用户显式填的值），企业基本信息补全只补还没被占用的字段
    var patch = Object.assign({}, built.patch);
    var notes = [];
    if (hasExtra) {
      var extra = null;
      try {
        extra = extraBuilder(rec) || {};
      } catch (e) {
        extra = { patch: {}, notes: ['企业基本信息补全计算异常：' + e.message] };
      }
      var ep = extra.patch || {};
      Object.keys(ep).forEach(function (k) {
        if (patch[k] === undefined) patch[k] = ep[k];
      });
      notes = extra.notes || [];
    }
    if (!Object.keys(patch).length) {
      // 无需写入：只把「为什么没补」的说明带回去
      return notes.length ? '企业基本信息未改动：' + notes.join('；') : '';
    }
    try {
      var r = await editEnterpriseInfo(token, orgId, Object.assign({}, rec, patch));
      if (r.status === 200 && r.data && r.data.code === '200') {
        if (typeof onSaved === 'function') { try { onSaved(patch); } catch (e) {} }
        var msg = '企业信息已更新（在职 ' +
          (patch.enterpriseNum !== undefined ? patch.enterpriseNum + ' 人' : '未改') +
          '，接害 ' + (patch.hazardNum !== undefined ? patch.hazardNum + ' 人' : '未改') + '）';
        if (notes.length) msg += '；企业基本信息补全：' + notes.join('；');
        return msg;
      }
      return '企业信息未写入：' + ((r.data && (r.data.message || r.data.msg)) || ('HTTP ' + r.status));
    } catch (e) {
      return '企业信息未写入：' + e.message;
    }
  }

  // ---------------- 项目团队相关接口 ----------------
  async function fetchUsers(token, orgId) {
    const r = await apiRequest('GET', '/api/users?pageSize=-1&organizationId=' + orgId, { token: token });
    if (r.status !== 200) throw new Error('获取人员列表失败(HTTP ' + r.status + ')');
    return (r.data && r.data.body && r.data.body.content) || [];
  }

  async function fetchProjectMembers(token, projectId) {
    const r = await apiRequest('GET', '/api/allocation/projectHuman?projectId=' + projectId + '&pageNumber=1&pageSize=200', { token: token });
    if (r.status !== 200) throw new Error('获取项目组成员失败(HTTP ' + r.status + ')');
    return (r.data && r.data.body && r.data.body.records) || [];
  }

  async function addProjectMembers(token, projectId, ids) {
    return apiRequest('POST', '/api/allocation/projectHuman/add', {
      token: token,
      json: {
        humanId: ids.map(function (id) { return Number(id); }),
        projectId: projectId,
        humanType: '0'
      }
    });
  }

  async function deleteProjectMembers(token, humanIds) {
    return apiRequest('DELETE', '/api/allocation/projectHuman/delete', {
      token: token,
      json: humanIds.map(function (id) { return Number(id); })
    });
  }

  async function fetchXcdcUser(token, projectId) {
    const r = await apiRequest('GET', '/api/projectXcdcUser?belongProject=' + projectId + '&pageSize=-1', { token: token });
    if (r.status !== 200) throw new Error('获取现场调查人员失败(HTTP ' + r.status + ')');
    const recs = (r.data && r.data.body && r.data.body.records) || [];
    return recs.length ? recs[0] : null;
  }

  async function saveXcdcUser(token, data) {
    return apiRequest('PUT', '/api/projectXcdcUser/edit', { token: token, json: data });
  }

  // 导入成功后同步：项目组成员（差量）+ 调查人/复核人/调查日期
  async function syncProjectInfo(token, projectId, opts) {
    const logs = [];
    const want = (opts.memberIds || []).map(String);
    const existing = await fetchProjectMembers(token, projectId);
    const existingIds = existing.map(function (m) { return String(m.humanId); });
    const toAdd = want.filter(function (id) { return existingIds.indexOf(id) < 0; });
    // 项目负责人保护：平台项目表 responser 字段即负责人用户 id，负责人不能被移除
    // （另有成员记录带非 '0' humanType 的负责人标记，也一并保护，双保险）
    const leaderIds = [];
    if (opts.project && opts.project.responser != null && opts.project.responser !== '') {
      leaderIds.push(String(opts.project.responser));
    }
    const toDel = existing.filter(function (m) {
      if (want.indexOf(String(m.humanId)) >= 0) return false;
      if (leaderIds.indexOf(String(m.humanId)) >= 0) return false;
      if (m.humanType != null && String(m.humanType) !== '0') return false; // 非 '0' 类型=负责人/管理员，不移除
      return true;
    });
    // 增删成员与调查人同步三者互不依赖：并行执行（原串行 5 次往返，现约 2 次）
    const jobs = [];
    if (toAdd.length > 0) {
      const names = toAdd.map(function (id) {
        const u = (opts.users || []).filter(function (x) { return String(x.id) === id; })[0];
        return u ? u.userName : id;
      });
      jobs.push(addProjectMembers(token, projectId, toAdd).then(function (r) {
        if (r.status === 200 && r.data && r.data.code === '200') {
          logs.push('添加项目成员成功：' + names.join('、'));
        } else {
          logs.push('添加项目成员失败（' + names.join('、') + '）：' + ((r.data && (r.data.message || r.data.msg)) || ('HTTP ' + r.status)));
        }
      }));
    }
    if (toDel.length > 0) {
      // 逐个删除：平台会拒绝删除项目负责人（整批提交时一人被拒全批失败）。
      // 单个失败只记日志提示（负责人本就不该被删，平台拒绝=符合预期），不进失败摘要弹窗
      for (const m of toDel) {
        const name = m.humanIdName || m.humanId;
        // eslint-disable-next-line no-await-in-loop
        const r = await deleteProjectMembers(token, [m.humanId]);
        if (r.status === 200 && r.data && r.data.code === '200') {
          logs.push('移除项目成员成功：' + name);
        } else {
          // 平台拒删（负责人等受保护成员）：静默保留，仅记提示日志
          logs.push('成员「' + name + '」保留（平台不允许移除，如负责人）');
        }
      }
    }
    if (opts.investigatorId || opts.reviewerId || opts.investigateDate) {
      jobs.push((async function () {
        const rec = await fetchXcdcUser(token, projectId);
        const payload = {
          id: rec ? rec.id : null,
          belongProject: projectId,
          investigatePerson: opts.investigatorId ? String(opts.investigatorId) : (rec ? rec.investigatePerson : null),
          accompanyPerson: opts.reviewerId ? String(opts.reviewerId) : (rec ? rec.accompanyPerson : null),
          investigateTime: opts.investigateDate ? opts.investigateDate : (rec ? rec.investigateTime : null)
        };
        const r = await saveXcdcUser(token, payload);
        if (r.status === 200 && r.data && r.data.code === '200') {
          logs.push('调查人/复核人/调查日期已同步');
        } else {
          // 逐项列出本次要同步的内容与平台返回，便于定位哪一项失败
          const parts = [];
          if (opts.investigatorId) parts.push('调查人');
          if (opts.reviewerId) parts.push('复核人');
          if (opts.investigateDate) parts.push('调查日期=' + opts.investigateDate);
          logs.push('同步失败（' + parts.join('、') + '）：' + ((r.data && (r.data.message || r.data.msg)) || ('HTTP ' + r.status)));
        }
      })());
    }
    await Promise.all(jobs);
    return logs;
  }

  // ---------------- 页面逻辑 ----------------
  function initApp() {
    const $ = function (id) { return document.getElementById(id); };

    const loginOverlay = $('xcdc-login');
    const uploadOverlay = $('xcdc-upload');
    const loginPanel = $('loginPanel');
    const captchaImg = $('captchaImg');
    const loginMsg = $('loginMsg');

    let session = loadSession();
    let uuid = null;
    // 登录方式：默认「云端代登录」——浏览器只提交账号密码，验证码由 CF 账号桥取图 + OCR + 提交，
    // 界面不显示验证码。captchaMode = true 仅在云端登录失败、用户点「改用验证码登录」后启用（兜底）。
    let captchaMode = false;
    // 验证码自动填写（仅 captchaMode 下用）：captchaAutoFilled = 输入框里的码是自动填入的（用户没改过）；
    //   captchaAutoFails = 连续失败次数；captchaAutoOff = 本次页面会话内已停用自动填写。
    let captchaAutoFilled = false;
    let captchaAutoFails = 0;
    let captchaAutoOff = false;
    const CAPTCHA_AUTO_MAX_FAILS = 2;
    let token = session ? session.token : null;
    let orgId = session ? session.orgId : null;
    let userInfo = session ? session.userInfo : null;
    let selectedProject = null;
    let selectedFile = null;
    let userList = [];
    let confirmedMemberIds = [];
    let teamDirty = false;
    let teamSaved = null; // 本账号已确认保存的团队设置

    function teamStorageKey() {
      const code = userInfo && (userInfo.userCode || userInfo.id) ? String(userInfo.userCode || userInfo.id) : 'anonymous';
      return 'xcdc_team_v1_' + code;
    }

    function teamAccountKey() {
      return userInfo && (userInfo.userCode || userInfo.id) ? String(userInfo.userCode || userInfo.id) : 'anonymous';
    }

    function greetingText() {
      const who = userInfo ? (userInfo.userName + '（' + userInfo.userCode + '）') : '';
      return '已登录：' + who;
    }

    // 是否运行在 Cloudflare Pages（HTTPS 部署）：团队配置云端同步仅在 pages.dev 上可用
    function isCloudDeploy() {
      try {
        return typeof location !== 'undefined' && location.protocol === 'https:' && /pages\.dev$/i.test(location.hostname);
      } catch (e) {
        return false;
      }
    }

    function normalizeTeamConfig(obj) {
      if (!obj || typeof obj !== 'object' || Array.isArray(obj)) return null;
      const has = Array.isArray(obj.memberIds) || obj.investigatorId || obj.reviewerId || obj.investigateDate;
      return has ? obj : null;
    }

    // 同源云端接口（/api/team，Cloudflare Pages Function 读写 KV）
    async function cloudTeamRequest(method, query, body) {
      const resp = await fetch('/api/team' + (query ? '?' + query : ''), {
        method: method,
        headers: body ? { 'Content-Type': 'application/json' } : {},
        body: body ? JSON.stringify(body) : undefined,
      });
      let data = null;
      try { data = await resp.json(); } catch (e) {}
      return { status: resp.status, data: data };
    }

    // 加载本账号团队配置：云端优先（换电脑可读取），本地 localStorage 兜底
    async function loadTeamSettings() {
      if (isCloudDeploy()) {
        try {
          const r = await cloudTeamRequest('GET', 'account=' + encodeURIComponent(teamAccountKey()));
          if (r.status === 200) {
            const cfg = normalizeTeamConfig(r.data);
            if (cfg) return cfg;
          }
        } catch (e) {}
      }
      try {
        const s = localStorage.getItem(teamStorageKey());
        return s ? normalizeTeamConfig(JSON.parse(s)) : null;
      } catch (e) {
        return null;
      }
    }

    function log(msg) {
      const box = $('log');
      if (!box) return;
      const line = document.createElement('div');
      line.className = 'log-line';
      line.textContent = '[' + new Date().toLocaleTimeString('zh-CN', { hour12: false }) + '] ' + msg;
      box.appendChild(line);
      box.scrollTop = box.scrollHeight;
    }

    function loadSession() {
      try {
        const s = sessionStorage.getItem(SESSION_KEY);
        return s ? JSON.parse(s) : null;
      } catch (e) {
        return null;
      }
    }

    function saveSession() {
      try {
        sessionStorage.setItem(SESSION_KEY, JSON.stringify({ token: token, orgId: orgId, userInfo: userInfo }));
      } catch (e) {}
    }

    function clearSession() {
      try { sessionStorage.removeItem(SESSION_KEY); } catch (e) {}
    }

    // 进入软件（关闭登录遮罩）
    function enterApp() {
      if (loginOverlay) loginOverlay.classList.add('hidden');
      authExpiredHandled = false; // 重新登录成功，允许再次提示
    }

    // ---------- 右上角用户菜单：按钮显示登录者姓名，点击展开菜单可选退出登录 ----------
    function syncUserMenu() {
      const wrap = $('user-menu-wrap');
      const btn = $('btn-logout');
      const nameEl = $('user-menu-name');
      if (!wrap || !btn) return;
      if (userInfo && (userInfo.userName || userInfo.userCode)) {
        wrap.style.display = '';
        btn.textContent = userInfo.userName || userInfo.userCode;
        if (nameEl) nameEl.textContent = userInfo.userName + '（' + userInfo.userCode + '）';
      } else {
        wrap.style.display = 'none';
        closeUserMenu();
      }
    }
    function closeUserMenu() {
      const menu = $('user-menu');
      if (menu) menu.style.display = 'none';
    }
    (function initUserMenu() {
      const btn = $('btn-logout');
      const menu = $('user-menu');
      const confirmBtn = $('btn-logout-confirm');
      if (btn && menu) {
        btn.addEventListener('click', function (e) {
          e.stopPropagation();
          menu.style.display = menu.style.display === 'none' ? '' : 'none';
        });
      }
      if (confirmBtn) {
        confirmBtn.addEventListener('click', function () {
          closeUserMenu();
          logout();
        });
      }
      // 点击菜单外任意位置关闭菜单
      document.addEventListener('pointerdown', function (e) {
        const wrap = $('user-menu-wrap');
        if (wrap && e.target.closest && !e.target.closest('#user-menu-wrap')) closeUserMenu();
      });
    })();

    // 显示登录遮罩
    function showLogin() {
      if (loginOverlay) loginOverlay.classList.remove('hidden');
      if (uploadOverlay) uploadOverlay.classList.add('hidden');
      // 密码框复位为密文（上次可能点过「眼睛」切成明文）
      try { resetPwdEye(); } catch (e) {}
      // 默认云端代登录，界面不显示验证码 → 不需要预先取图；兜底模式才取
      if (captchaMode) refreshCaptcha();
    }

    // 上传视图：打开/关闭
    async function showUpload() {
      if (!token) { showLogin(); return; }
      if (uploadOverlay) uploadOverlay.classList.remove('hidden');
      if ($('greeting')) $('greeting').textContent = greetingText();
      if (userList.length === 0 && token && orgId) await loadUserList();
      // 打开上传视图时从云端刷新，换电脑也能读到最新团队配置
      teamSaved = await loadTeamSettings();
      if (teamSaved) applyTeamSettings(teamSaved);
      generateExportFile();
      generateSurveyFile(); // 进入上传页面自动重新生成两个上传文件（主表 + 调查表 6 合一）
    }

    function hideUpload() {
      if (uploadOverlay) uploadOverlay.classList.add('hidden');
    }

    // GitHub 配置按钮仅 btc-wf 账号可见（登录/登出/恢复会话时联动显隐）
    const GH_ALLOWED_USER = 'btc-wf';
    function syncGhButton() {
      const code = userInfo ? String(userInfo.userCode || userInfo.id || '') : '';
      const btnGh = document.getElementById('btn-gh');
      const btnFsGh = document.getElementById('fs-gh');
      if (btnGh) btnGh.style.display = code === GH_ALLOWED_USER ? '' : 'none';
      if (btnFsGh) btnFsGh.style.display = code === GH_ALLOWED_USER ? '' : 'none';
    }

    function logout() {
      clearSession();
      session = null;
      token = null;
      orgId = null;
      userInfo = null;
      syncGhButton();
      syncUserMenu();
      teamSaved = null;
      selectedProject = null;
      selectedFile = null;
      userList = [];
      confirmedMemberIds = [];
      teamDirty = false;
      showLogin();
    }

    // ---------- 登录失效（401）：提示并引导重新登录 ----------
    // 会话过期后，平台接口会返回 401（"找不到当前登录的信息"），表现为「搜索不出内容」。
    // 这里统一处理：清理失效会话 → 弹登录遮罩 → 提示用户重新登录（同一轮只提示一次）。
    let authExpiredHandled = false;
    function notifyAuthExpired() {
      if (authExpiredHandled) return;
      authExpiredHandled = true;
      const who = userInfo && (userInfo.userName || userInfo.userCode)
        ? ('（' + (userInfo.userName || userInfo.userCode) + '）')
        : '';
      try { logout(); } catch (e) { /* 清理失败也要弹登录 */ }
      // 让登录遮罩先渲染出来，再弹提示（避免提示被遮罩挡住的观感）
      setTimeout(function () {
        alert('登录已失效' + who + '，请重新登录。\n\n' +
          '· 长时间未操作，或账号在其它地方登录，平台会失效当前会话；\n' +
          '· 重新登录后再执行刚才的操作即可（刚才的搜索结果为空是因为登录失效）。');
      }, 50);
    }
    // 注入给顶层的 apiRequest：任何接口返回 401 都会走到这里
    authExpiredHandler = notifyAuthExpired;

    // 账号池验证码 OCR 端点：与 /token 同源同级（把结尾的 /token 换成 /code）。
    // bridgeBase() 在下面定义（函数声明会提升），返回 /api/bridge/token 或 http://127.0.0.1:8787/token。
    function bridgeSibling(suffix) {
      return bridgeBase().then(function (b) {
        const s = String(b).replace(/\/+$/, '');
        // 正常返回 …/token（同源或直连）；xcdc_bridge_url 覆盖时可能只给到根地址，两种都兼容
        return /\/token$/.test(s) ? s.replace(/\/token$/, '/' + suffix) : s + '/' + suffix;
      });
    }
    async function bridgeCodeUrl() { return bridgeSibling('code'); }

    // 云端代登录：把账号密码 POST 给账号桥，桥在服务端取图 + OCR + 调 /auth/login，
    // 只把 token（及用户信息）回给浏览器。验证码不出现在界面上、也不回传。
    // 成功返回 { ok, token, user, orgId, via, conf }；失败抛错（消息可直接展示）。
    async function loginViaBridge(user, pass) {
      const url = await bridgeSibling('login');
      const ctrl = typeof AbortController !== 'undefined' ? new AbortController() : null;
      const timer = ctrl ? setTimeout(function () { ctrl.abort(); }, 30000) : null;
      try {
        const r = await fetch(url, {
          method: 'POST',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify({ username: user, password: pass }),
          cache: 'no-store',
          signal: ctrl ? ctrl.signal : undefined,
        });
        let d = null;
        try { d = await r.json(); } catch (e) { d = null; }
        if (!d) throw new Error('账号桥返回异常（HTTP ' + r.status + '）');
        if (!d.ok || !d.token) throw new Error(d.message || ('HTTP ' + r.status));
        return d;
      } catch (e) {
        if (e && e.name === 'AbortError') throw new Error('云端登录超时（30 秒），请重试');
        throw e;
      } finally {
        if (timer) clearTimeout(timer);
      }
    }

    // 云端登录失败后露出「改用验证码登录」入口（正常情况界面不显示验证码）
    function showCaptchaFallback() {
      const link = $('btnCaptchaFallback');
      if (link) link.hidden = false;
    }
    function enableCaptchaMode() {
      captchaMode = true;
      const field = $('captchaField');
      if (field) field.hidden = false;
      const link = $('btnCaptchaFallback');
      if (link) link.hidden = true;
      loginMsg.textContent = '已切换为验证码登录';
      refreshCaptcha();
    }

    // 走账号池 OCR 取一张验证码：返回 { ok, uuid, code, conf, img } 或 null（桥不可用/识别失败）。
    // 只取「验证码答案」，账号密码仍由浏览器带着用户手输值去登录。
    async function fetchCaptchaViaBridge() {
      let url;
      try { url = await bridgeCodeUrl(); } catch (e) { return null; }
      const ctrl = typeof AbortController !== 'undefined' ? new AbortController() : null;
      const timer = ctrl ? setTimeout(function () { ctrl.abort(); }, 8000) : null;
      try {
        const r = await fetch(url, { cache: 'no-store', signal: ctrl ? ctrl.signal : undefined });
        if (!r.ok) return null;
        const d = await r.json();
        if (!d || !d.ok || !d.uuid || !d.code) return null;
        return d;
      } catch (e) {
        return null;   // 桥未启动 / 跨域被拒 / 超时 → 静默回退手输
      } finally {
        if (timer) clearTimeout(timer);
      }
    }

    // 刷新验证码：优先账号池 OCR 自动填写；桥不可用或连续失败则回退「浏览器取图 + 手输」。
    // opts.forceManual = true 时跳过 OCR，直接给一张干净的手输图。
    // 返回 true 表示已拿到一张验证码（uuid 有效）。
    async function refreshCaptcha(opts) {
      opts = opts || {};
      const useAuto = !opts.forceManual && !captchaAutoOff;

      // ① 账号池 OCR：拿到 uuid + 识别结果 + 图片，自动填入输入框
      if (useAuto) {
        // 识别需要 1~3 秒，给个提示，避免用户以为验证码加载失败
        if (loginMsg) loginMsg.textContent = '正在自动识别验证码…';
        const auto = await fetchCaptchaViaBridge();
        if (auto) {
          uuid = auto.uuid;
          if (auto.img) captchaImg.src = auto.img;
          $('captcha').value = auto.code;
          captchaAutoFilled = true;
          captchaAutoFails = 0;
          loginMsg.textContent = '';
          return true;
        }
        // OCR 不可用/失败：累计失败，到上限就本次会话内停用，避免每次开页都白等
        captchaAutoFails++;
        if (captchaAutoFails >= CAPTCHA_AUTO_MAX_FAILS) captchaAutoOff = true;
      }

      // ② 回退：浏览器直取图，用户手输
      captchaAutoFilled = false;
      try {
        const d = await fetchCaptcha();
        uuid = d.uuid;
        captchaImg.src = d.img;
        $('captcha').value = '';
        loginMsg.textContent = '';
        return true;
      } catch (e) {
        let msg = e.message;
        // GitHub Pages 等 HTTPS 静态托管无法直连平台 HTTP 接口，给出明确指引
        try {
          if (location.protocol === 'https:' && !/pages\.dev$/i.test(location.hostname)) {
            msg = 'HTTPS 页面无法直连平台 HTTP 接口（被浏览器拦截）。请使用 Cloudflare Pages 部署：https://sampling-plan.pages.dev';
          }
        } catch (e2) {}
        loginMsg.textContent = msg;
        return false;
      }
    }

    // 手动刷新 = 明确要换一张（仍优先走 OCR 自动填；OCR 已停用时自然回退手输）
    $('btnRefresh').addEventListener('click', function () { refreshCaptcha(); });
    captchaImg.addEventListener('click', function () { refreshCaptcha(); });
    // 用户自己改了验证码 → 视为手输，登录失败时不再自动换图重试
    (function () {
      const el = $('captcha');
      if (el) el.addEventListener('input', function () { captchaAutoFilled = false; });
    })();
    // 云端登录失败后的兜底入口：切到「手输验证码」模式
    (function () {
      const link = $('btnCaptchaFallback');
      if (link) link.addEventListener('click', function (e) { e.preventDefault(); enableCaptchaMode(); });
    })();

    // 密码框右侧「眼睛」：点击在明文/密文之间切换（图标与提示同步）
    let resetPwdEye = function () {};
    (function () {
      const btn = $('btnPwdEye');
      const pwd = $('password');
      if (!btn || !pwd) return;
      const sync = function (shown) {
        btn.dataset.shown = shown ? 'true' : 'false';
        const label = shown ? '隐藏密码' : '显示密码';
        btn.title = label;
        btn.setAttribute('aria-label', label);
        btn.setAttribute('aria-pressed', shown ? 'true' : 'false');
      };
      sync(false);
      // 回到登录界面时复位为密文，避免上次切成明文后密码裸露在登录页
      resetPwdEye = function () { pwd.type = 'password'; sync(false); };
      btn.addEventListener('click', function () {
        const shown = pwd.type === 'text';
        pwd.type = shown ? 'password' : 'text';
        sync(!shown);
        // 切换后保持焦点与光标在末尾，方便继续输入
        try {
          pwd.focus({ preventScroll: true });
          const n = pwd.value.length;
          pwd.setSelectionRange(n, n);
        } catch (e) { /* 老内核忽略 */ }
      });
    })();

    // 打开页面时自动填充已保存的账号密码
    try {
      if (localStorage.getItem(REM_KEY) === '1') {
        $('rememberMe').checked = true;
        $('username').value = localStorage.getItem(USR_KEY) || '';
        $('password').value = localStorage.getItem(PWD_KEY) || '';
      }
    } catch (e) { }

    // 账号/密码/验证码框内按回车直接登入（等价点击「登 录」）
    ['username', 'password', 'captcha'].forEach(function (id) {
      const el = document.getElementById(id);
      if (el) el.addEventListener('keydown', function (e) {
        if (e.key === 'Enter') { e.preventDefault(); $('btnLogin').click(); }
      });
    });

    $('btnLogin').addEventListener('click', async function () {
      const user = $('username').value.trim();
      const pass = $('password').value;
      if (!user || !pass) { loginMsg.textContent = '请填写账号和密码'; return; }
      // 兜底模式才需要验证码；默认云端代登录，界面没有验证码输入框
      const code = captchaMode ? $('captcha').value.trim() : '';
      if (captchaMode && !code) { loginMsg.textContent = '请填写验证码'; return; }
      const btn = $('btnLogin');
      btn.disabled = true;
      btn.textContent = '登录中...';
      let nextToken = null;
      let nextUser = null;
      let nextOrgId = null;
      try {
        if (!captchaMode) {
          // ① 云端代登录：账号密码发到账号桥，桥取图 + OCR + 调 /auth/login，只回 token
          loginMsg.textContent = '正在云端识别验证码并登录…';
          const br = await loginViaBridge(user, pass);
          nextToken = br.token;
          nextUser = br.user;
          nextOrgId = br.orgId;
        } else {
          // ② 兜底：浏览器自己取图 + 手输验证码
          let r = await login(user, pass, code, uuid);
          // 验证码是自动填入且用户没改过 → 失败极可能是 OCR 认错了：换一张自动再试一次。
          // 手输的验证码不自动重试（避免把「密码错」掩盖成「验证码错」而反复提交）。
          if ((r.status !== 200 || !r.data || !r.data.token) && captchaAutoFilled) {
            const fresh = await refreshCaptcha();
            const code2 = $('captcha').value.trim();
            if (fresh && captchaAutoFilled && code2) {
              loginMsg.textContent = '验证码自动识别可能有误，已换一张重试…';
              r = await login(user, pass, code2, uuid);
            } else {
              captchaAutoFails++;
              if (captchaAutoFails >= CAPTCHA_AUTO_MAX_FAILS) captchaAutoOff = true;
            }
          }
          if (r.status !== 200 || !r.data || !r.data.token) {
            const msg = (r.data && (r.data.message || r.data.msg)) || ('HTTP ' + r.status);
            throw new Error(msg);
          }
          nextToken = r.data.token;
          nextUser = r.data.user;
        }
        token = nextToken;
        // 平台登录响应的 user 可能是 {user:{…}}，也可能直接是用户对象，两种都兼容
        userInfo = (nextUser && nextUser.user && typeof nextUser.user === 'object') ? nextUser.user : (nextUser || {});
        // 登录后调用 auth/info 获取权威用户信息（与平台网站一致），解析真实机构
        try {
          const full = await fetchAuthInfo(token);
          if (full) {
            userInfo = Object.assign({}, userInfo, full);
            if (full.userName) userInfo.userName = full.userName;
            if (full.userCode) userInfo.userCode = full.userCode;
          }
        } catch (e) {}
        orgId = nextOrgId || resolveOrgId(userInfo) || (userInfo.organization && userInfo.organization.organizationId) || null;
        session = { token: token, orgId: orgId, userInfo: userInfo };
        saveSession();
        teamSaved = await loadTeamSettings(); // 按登录账号恢复已确认的团队设置（云端优先）
        teamDirty = !!teamSaved;
        // 记住账号密码
        try {
          if ($('rememberMe').checked) {
            localStorage.setItem(REM_KEY, '1');
            localStorage.setItem(USR_KEY, user);
            localStorage.setItem(PWD_KEY, pass);
          } else {
            localStorage.removeItem(REM_KEY);
            localStorage.removeItem(USR_KEY);
            localStorage.removeItem(PWD_KEY);
          }
        } catch (e) { }
        loginMsg.textContent = '';          // 清掉「正在云端识别…」，下次登出回来不残留
        $('greeting').textContent = greetingText();
        syncGhButton();
        syncUserMenu();
        enterApp();
        log('登录成功：' + userInfo.userName + '（' + userInfo.userCode + '），上传主体为当前账号');
        // 登录账号是池成员时：本会话令牌直接入池共享（回退/审核用该账号不再经桥重复登录）
        if (userInfo && POOL_ACCOUNTS.indexOf(userInfo.userCode) !== -1) {
          poolTokens[userInfo.userCode] = { token: token, orgId: orgId };
          log('「' + userInfo.userCode + '」为本会话登录账号，令牌已入池共享');
        }
        prefetchYear();
        loadUserList();
        warmAccountPool(); // 后台补齐池内其余账号（已有缓存的跳过，不重复登录）
      } catch (e) {
        loginMsg.textContent = '登录失败：' + e.message;
        if (!captchaMode) {
          // 云端代登录失败：露出兜底入口（正常界面不显示验证码）
          showCaptchaFallback();
        } else {
          // 这次提交的是自动填入的验证码 → 记为一次 OCR 失败；到上限就停用自动填写，
          // 并直接给一张干净的手输图（否则用户会被反复自动填错值卡住）。
          if (captchaAutoFilled) {
            captchaAutoFails++;
            if (captchaAutoFails >= CAPTCHA_AUTO_MAX_FAILS) captchaAutoOff = true;
          }
          refreshCaptcha({ forceManual: captchaAutoOff });
        }
      } finally {
        btn.disabled = false;
        btn.textContent = '登 录';
      }
    });

    // 自动生成当前表格的导出 Excel 作为上传文件
    async function generateExportFile() {
      const fileNameEl = $('fileName');
      try {
        if (!window.SamplingApp || !window.SamplingApp.exportWorkbookBytes) throw new Error('导出模块未就绪');
        const bytes = await window.SamplingApp.exportWorkbookBytes();
        const name = (window.SamplingApp.exportName ? window.SamplingApp.exportName() : '系统测点布局调查_自动计算区.xlsx');
        selectedFile = new File([bytes], name, { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
        if (fileNameEl) fileNameEl.textContent = name + '（' + Math.max(1, Math.round(bytes.length / 1024)) + ' KB）';
        log('已生成上传文件：' + name);
      } catch (e) {
        selectedFile = null;
        if (fileNameEl) fileNameEl.textContent = '';
        log('生成导出文件失败：' + e.message);
      }
    }

    // 受检单位 / 项目编号 自动搜索
    const unitInput = $('unitName');
    const codeInput = $('projectCode');
    const yearSelect = $('yearSelect');
    const suggestBox = $('suggest');
    const memberSearch = $('memberSearch');
    const memberList = $('memberList');
    const memberTags = $('memberTags');
    const btnMemberConfirm = $('btnMemberConfirm');
    const investigatorSelect = $('investigatorSelect');
    const reviewerSelect = $('reviewerSelect');
    const investigateDate = $('investigateDate');
    const entOnJobNum = $('entOnJobNum');
    const entHazardNum = $('entHazardNum');
    const chkEntBasic = $('chkEntBasicUpload');   // 可选元素：老页面缓存里没有时不影响启动
    if (!unitInput || !codeInput || !yearSelect || !suggestBox || !$('searching') || !$('projectInfo') ||
        !memberSearch || !memberList || !memberTags || !btnMemberConfirm || !investigatorSelect || !reviewerSelect || !investigateDate ||
        !entOnJobNum || !entHazardNum) {
      if (typeof console !== 'undefined') console.error('上传页面元素缺失，请强制刷新浏览器（Ctrl+F5）后重试');
      alert('上传页面加载不完整，请按 Ctrl+F5 强制刷新后重试');
      return;
    }
    let searchSeq = 0;
    let currentAbort = null;
    let debounceTimer = null;

    // 企业信息（在职人数/接害人数）：输入框 ↔ localStorage 缓存 ↔ 平台现值
    const ENTINFO_KEY = 'xcdc_entinfo_v1';
    let entInfoSeq = 0;

    // 「企业基本信息上传」勾选项（只在「上传调查表（6 合一）」时生效）：记住上次选择
    const ENTBASIC_KEY = 'xcdc_entbasic_v1';
    function entBasicChecked() {
      return !!(chkEntBasic && chkEntBasic.checked);
    }
    function initEntBasicToggle() {
      if (!chkEntBasic) return;
      try { chkEntBasic.checked = localStorage.getItem(ENTBASIC_KEY) === '1'; } catch (e) {}
      chkEntBasic.addEventListener('change', function () {
        try { localStorage.setItem(ENTBASIC_KEY, chkEntBasic.checked ? '1' : '0'); } catch (e) {}
      });
    }
    initEntBasicToggle();

    function entInfoCacheRead() {
      try { return JSON.parse(localStorage.getItem(ENTINFO_KEY)) || {}; } catch (e) { return {}; }
    }
    function entInfoCacheSave(projectId, values) {
      try {
        const all = entInfoCacheRead();
        all[String(projectId)] = {
          onJobNum: String((values && values.onJobNum) || ''),
          hazardNum: String((values && values.hazardNum) || '')
        };
        localStorage.setItem(ENTINFO_KEY, JSON.stringify(all));
      } catch (e) {}
    }
    function setEntInputs(onJob, hazard) {
      if (entOnJobNum) entOnJobNum.value = onJob == null ? '' : String(onJob);
      if (entHazardNum) entHazardNum.value = hazard == null ? '' : String(hazard);
    }
    function entInputValues() {
      return { onJobNum: entOnJobNum ? entOnJobNum.value : '', hazardNum: entHazardNum ? entHazardNum.value : '' };
    }
    // 上传成功后清空「本次实际写入」的那个输入框 + 它在本地的缓存值（2026-10-18 用户需求）。
    //   · 只清写入的那一侧：另一个文件还没上传，提前清掉会让人重新填一遍；
    //   · 由 applyEntInfoOnUpload 的 onSaved 回调触发，即**平台确认写入成功（code=200）**时才走到这里，
    //     所以清空不会丢数据（值已在平台上）；
    //   · 缓存同步置空，避免下次选中同一项目时又回填出旧值（若平台有值则按平台现值回填）。
    // 入参 patch 是真正写进平台的补丁：{ enterpriseNum }=只写了在职人数、{ hazardNum }=只写了接害人数。
    // 返回被清空的输入框名称（'在职人数' / '接害人数' / '在职人数、接害人数'），无则返回 ''。
    function clearEntInputsAfterUpload(patch, projectId) {
      const p = patch || {};
      const pid = String(projectId);
      const prev = entInfoCacheRead()[pid] || {};
      const next = { onJobNum: prev.onJobNum, hazardNum: prev.hazardNum };
      let cleared = '';
      if (p.enterpriseNum !== undefined) {
        if (entOnJobNum) entOnJobNum.value = '';
        next.onJobNum = '';
        cleared = '在职人数';
      }
      if (p.hazardNum !== undefined) {
        if (entHazardNum) entHazardNum.value = '';
        next.hazardNum = '';
        cleared = cleared ? cleared + '、接害人数' : '接害人数';
      }
      if (cleared) entInfoCacheSave(pid, next);
      return cleared;
    }
    // 上传前预校验：非法值直接拦截（等上传完再报就晚了）
    function validateEntInputs() {
      const v = entInputValues();
      const pairs = [['在职人数', v.onJobNum], ['接害人数', v.hazardNum]];
      for (let i = 0; i < pairs.length; i++) {
        const s = String(pairs[i][1] || '').trim();
        if (s && !/^\d+$/.test(s)) return pairs[i][0] + '须为非负整数（当前值：' + s + '）';
      }
      const a = parseInt(v.onJobNum, 10);
      const b = parseInt(v.hazardNum, 10);
      if (!isNaN(a) && !isNaN(b) && b > a) return '接害人数不能大于在职人数（' + b + ' > ' + a + '）';
      return '';
    }
    // 选中项目后预填：先按项目 id 取本地缓存，再拉平台当前值覆盖（过期响应丢弃）
    function prefillEntInfo(project) {
      const seq = ++entInfoSeq;
      if (!project) { setEntInputs('', ''); return; }
      const cached = entInfoCacheRead()[String(project.id)];
      setEntInputs(cached && cached.onJobNum, cached && cached.hazardNum);
      if (!token || project.belongInspect === undefined || project.belongInspect === null || project.belongInspect === '') return;
      fetchEnterpriseInfo(token, orgId, project.belongInspect).then(function (rec) {
        if (seq !== entInfoSeq) return; // 已切换项目，丢弃过期响应
        if (!rec) return;
        const onJob = rec.enterpriseNum == null ? '' : String(rec.enterpriseNum);
        const hazard = rec.hazardNum == null ? '' : String(rec.hazardNum);
        setEntInputs(onJob, hazard);
        entInfoCacheSave(project.id, { onJobNum: onJob, hazardNum: hazard });
      }).catch(function () { /* 拉取失败：保留缓存/已填值 */ });
    }

    function selectedMemberIds() {
      return confirmedMemberIds.slice();
    }

    function renderMemberList() {
      memberList.innerHTML = '';
      userList.forEach(function (u) {
        const label = document.createElement('label');
        label.className = 'member-item';
        const cb = document.createElement('input');
        cb.type = 'checkbox';
        cb.value = String(u.id);
        cb.dataset.name = u.userName || '';
        cb.dataset.code = u.userCode || '';
        const span = document.createElement('span');
        span.textContent = u.userName + '（' + (u.userCode || '') + '）';
        label.appendChild(cb);
        label.appendChild(span);
        label.addEventListener('click', function (e) {
          if (e.target !== cb) cb.checked = !cb.checked;
          updateMemberCount();
        });
        memberList.appendChild(label);
      });
      filterMemberList();
    }

    function filterMemberList() {
      const kw = memberSearch.value.trim().toLowerCase();
      Array.from(memberList.children).forEach(function (label) {
        const cb = label.querySelector('input');
        const text = ((cb.dataset.name || '') + (cb.dataset.code || '')).toLowerCase();
        label.classList.toggle('hidden', !!kw && text.indexOf(kw) < 0);
      });
    }

    function checkedMemberIds() {
      return Array.from(memberList.querySelectorAll('input:checked')).map(function (cb) { return cb.value; });
    }

    function updateMemberCount() {
      $('memberCount').textContent = '（已勾选 ' + checkedMemberIds().length + ' 人）';
    }

    function renderMemberTags() {
      memberTags.innerHTML = '';
      confirmedMemberIds.forEach(function (id) {
        const u = userList.filter(function (x) { return String(x.id) === id; })[0];
        if (!u) return;
        const tag = document.createElement('span');
        tag.className = 'member-tag';
        const text = document.createElement('span');
        text.textContent = u.userName + '（' + (u.userCode || '') + '）';
        const rm = document.createElement('button');
        rm.type = 'button';
        rm.className = 'member-tag-rm';
        rm.textContent = '×';
        rm.title = '移除该成员';
        rm.addEventListener('click', function (e) {
          e.stopPropagation();
          removeConfirmedMember(id);
        });
        tag.appendChild(text);
        tag.appendChild(rm);
        memberTags.appendChild(tag);
      });
    }

    // 点击成员标签上的 × 移除该成员
    function removeConfirmedMember(id) {
      const u = userList.filter(function (x) { return String(x.id) === String(id); })[0];
      confirmedMemberIds = confirmedMemberIds.filter(function (x) { return String(x) !== String(id); });
      Array.from(memberList.querySelectorAll('input')).forEach(function (cb) {
        if (cb.value === String(id)) cb.checked = false;
      });
      teamDirty = true;
      renderMemberTags();
      rebuildRoleOptions();
      updateMemberCount();
      saveTeamSettings().catch(function () {});
      log('已移除成员：' + (u ? u.userName + '（' + (u.userCode || '') + '）' : id));
    }

    async function confirmMembers(quiet, markDirty) {
      confirmedMemberIds = checkedMemberIds();
      if (markDirty) {
        teamDirty = true;
        memberSearch.value = '';
        filterMemberList();
        await saveTeamSettings(); // 确认后按当前登录账号保存（云端），不再清空
      }
      renderMemberTags();
      rebuildRoleOptions();
      updateMemberCount();
      if (!quiet) log('已确认项目成员 ' + confirmedMemberIds.length + ' 人');
    }

    // 按账号保存已确认的团队设置（成员/调查人/复核人/日期）：
    // 本地始终留一份兜底，部署在 Cloudflare Pages 时同步到云端（跨电脑读取）
    async function saveTeamSettings() {
      const data = {
        memberIds: confirmedMemberIds.slice(),
        investigatorId: investigatorSelect ? investigatorSelect.value : '',
        reviewerId: reviewerSelect ? reviewerSelect.value : '',
        investigateDate: investigateDate ? investigateDate.value : '',
        savedAt: Date.now()
      };
      teamSaved = data;
      try {
        localStorage.setItem(teamStorageKey(), JSON.stringify(data));
      } catch (e) {}
      if (isCloudDeploy()) {
        try {
          const r = await cloudTeamRequest('PUT', '', { account: teamAccountKey(), config: data });
          if (r.status === 200 && r.data && r.data.ok) {
            log('团队配置已同步到云端');
          } else {
            log('团队配置云端同步失败：' + ((r.data && r.data.message) || ('HTTP ' + r.status)));
          }
        } catch (e) {
          log('团队配置云端同步失败：' + e.message);
        }
      }
      return data;
    }

    function applyTeamSettings(data) {
      if (!data) return false;
      confirmedMemberIds = Array.isArray(data.memberIds) ? data.memberIds.slice() : [];
      Array.from(memberList.querySelectorAll('input')).forEach(function (cb) {
        cb.checked = confirmedMemberIds.indexOf(cb.value) >= 0;
      });
      renderMemberTags();
      rebuildRoleOptions();
      updateMemberCount();
      if (data.investigatorId && Array.from(investigatorSelect.options).some(function (o) { return o.value === data.investigatorId; })) {
        investigatorSelect.value = data.investigatorId;
      }
      if (data.reviewerId && Array.from(reviewerSelect.options).some(function (o) { return o.value === data.reviewerId; })) {
        reviewerSelect.value = data.reviewerId;
      }
      if (data.investigateDate) investigateDate.value = String(data.investigateDate).slice(0, 10);
      return true;
    }

    function fillRoleSelect(sel) {
      const cur = sel.value;
      sel.innerHTML = '';
      const empty = document.createElement('option');
      empty.value = '';
      empty.textContent = '（请选择）';
      sel.appendChild(empty);
      selectedMemberIds().forEach(function (id) {
        const u = userList.filter(function (x) { return String(x.id) === id; })[0];
        if (!u) return;
        const opt = document.createElement('option');
        opt.value = String(u.id);
        opt.textContent = u.userName + '（' + (u.userCode || '') + '）';
        sel.appendChild(opt);
      });
      sel.value = cur && Array.from(sel.options).some(function (o) { return o.value === cur; }) ? cur : '';
    }

    function rebuildRoleOptions() {
      fillRoleSelect(investigatorSelect);
      fillRoleSelect(reviewerSelect);
    }

    function fmtDate(d) {
      return d.getFullYear() + '-' + String(d.getMonth() + 1).padStart(2, '0') + '-' + String(d.getDate()).padStart(2, '0');
    }

    async function loadUserList() {
      try {
        userList = await fetchUsers(token, orgId);
        renderMemberList();
        teamSaved = await loadTeamSettings(); // 云端优先，登录/刷新后立即恢复本账号已确认的团队配置
        if (teamSaved) applyTeamSettings(teamSaved);
        log('已加载人员列表：' + userList.length + ' 人');
        if (selectedProject) loadProjectTeam(selectedProject.id);
      } catch (e) {
        log('加载人员列表失败：' + e.message);
      }
    }

    async function loadProjectTeam(projectId) {
      try {
        // 本账号已确认保存过团队设置：恢复使用，不覆盖、不清空
        if (teamSaved) {
          if (applyTeamSettings(teamSaved)) {
            log('已按账号恢复已确认的团队设置（切换项目不覆盖）');
            return;
          }
        }
        if (teamDirty) {
          log('已保留团队设置，切换项目不覆盖');
          return;
        }
        const members = await fetchProjectMembers(token, projectId);
        const ids = members.map(function (m) { return String(m.humanId); });
        Array.from(memberList.querySelectorAll('input')).forEach(function (cb) { cb.checked = ids.indexOf(cb.value) >= 0; });
        confirmMembers(true, false);
        log('已加载项目组成员：' + members.length + ' 人');
        const rec = await fetchXcdcUser(token, projectId);
        if (rec) {
          if (rec.investigatePerson) investigatorSelect.value = String(rec.investigatePerson);
          if (rec.accompanyPerson) reviewerSelect.value = String(rec.accompanyPerson);
          if (rec.investigateTime) {
            const t = rec.investigateTime;
            investigateDate.value = typeof t === 'number' ? fmtDate(new Date(t)) : String(t).slice(0, 10);
          }
          rebuildRoleOptions();
        }
      } catch (e) {
        log('加载项目团队信息失败：' + e.message);
      }
    }

    function setSearching(on) {
      const s = $('searching');
      if (s) s.classList.toggle('hidden', !on);
    }

    function renderSuggest(list) {
      suggestBox.innerHTML = '';
      if (!list || list.length === 0) {
        const empty = document.createElement('div');
        empty.className = 'suggest-empty';
        empty.textContent = '未找到匹配的项目';
        suggestBox.appendChild(empty);
        suggestBox.classList.remove('hidden');
        return;
      }
      list.forEach(function (p) {
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
        item.addEventListener('click', function () {
          selectedProject = p;
          unitInput.value = '';
          codeInput.value = '';
          suggestBox.classList.add('hidden');
          $('projectInfo').textContent = '已选择项目：' + p.code + '（' + (p.belongInspectName || '') + '）';
          $('projectInfo').className = 'info ok';
          log('已选择项目：' + p.code + ' ' + (p.belongInspectName || ''));
          loadProjectTeam(p.id);
          prefillEntInfo(p);   // 企业信息输入框预填（本地缓存 + 平台现值）
          refreshReportPanel();   // 「数据报告处理」页签同步显示当前项目
        });
        suggestBox.appendChild(item);
      });
      suggestBox.classList.remove('hidden');
    }

    // 项目搜索（「数据上传」搜索框与「网站数据导入」共用同一套逻辑与缓存）：
    // 服务端 code 为「包含匹配」（如 0959 可命中 BTC26-SZJC0959），belongInspectName 为模糊匹配，
    // 与 year 三者 AND；只在「仅选年份、未填编号与单位」时才走整年列表浏览。
    function runProjectSearch(keyword, signal) {
      const unitName = keyword && keyword.unitName ? String(keyword.unitName).trim() : '';
      const code = keyword && keyword.code ? String(keyword.code).trim() : '';
      const yearNum = keyword && keyword.year ? String(keyword.year).trim() : '';
      if (!unitName && !code && !yearNum) return Promise.resolve([]);
      if (code || unitName) {
        const kw = {};
        if (code) kw.code = code;
        if (unitName) kw.unitName = unitName;
        if (yearNum) kw.year = '20' + yearNum;
        return searchProjects(token, orgId, kw, signal).then(function (list) {
          return (list || []).slice(0, 15);
        });
      }
      return fetchYearProjects(token, orgId, 'BTC' + yearNum, '', signal).then(function (list) {
        return (list || []).slice(0, 15);
      });
    }

    function doSearch() {
      try {
        const unitName = unitInput.value.trim();
        const code = codeInput.value.trim();
        const year = yearSelect.value;
        if (!unitName && !code && !year) {
          suggestBox.classList.add('hidden');
          log('请先输入受检单位或项目编号，再点击搜索');
          return;
        }
        selectedProject = null;
        $('projectInfo').textContent = '';
        $('projectInfo').className = 'info';
        refreshReportPanel();
        const seq = ++searchSeq;
        if (currentAbort) currentAbort.abort();
        currentAbort = new AbortController();
        setSearching(true);
        const promise = runProjectSearch({ code: code, unitName: unitName, year: year }, currentAbort.signal);
        promise
          .then(function (list) {
            if (seq !== searchSeq) return;
            renderSuggest(list);
          })
          .catch(function (e) {
            if (seq !== searchSeq) return;
            suggestBox.classList.add('hidden');
            if (e && e.name === 'AbortError') {
              log('搜索超时，请稍后重试');
            } else {
              log('搜索失败：' + e.message);
            }
          })
          .finally(function () {
            if (seq === searchSeq) setSearching(false);
          });
      } catch (e) {
        log('搜索异常：' + e.message);
        setSearching(false);
      }
    }

    function prefetchYear() {
      const prefix = yearSelect.value ? 'BTC' + yearSelect.value : '';
      if (prefix && token) {
        fetchYearProjects(token, orgId, prefix, '', null).catch(function () { });
      }
    }

    unitInput.addEventListener('keydown', function (e) {
      if (e.key === 'Enter') { doSearch(); }
    });
    codeInput.addEventListener('keydown', function (e) {
      if (e.key === 'Enter') { doSearch(); }
    });
    unitInput.addEventListener('input', function () {
      clearTimeout(debounceTimer);
      debounceTimer = setTimeout(doSearch, 300);
    });
    codeInput.addEventListener('input', function () {
      clearTimeout(debounceTimer);
      debounceTimer = setTimeout(doSearch, 300);
    });
    yearSelect.addEventListener('change', function () {
      prefetchYear();
      // 受检单位/项目编号有内容时，切换年份重新触发搜索（按新年份过滤结果）
      if (unitInput.value.trim() || codeInput.value.trim()) {
        clearTimeout(debounceTimer);
        debounceTimer = setTimeout(doSearch, 300);
      }
    });
    memberSearch.addEventListener('input', filterMemberList);
    memberList.addEventListener('change', updateMemberCount);
    btnMemberConfirm.addEventListener('click', function () { confirmMembers(false, true).catch(function () {}); });
    investigatorSelect.addEventListener('change', function () { teamDirty = true; saveTeamSettings().catch(function () {}); });
    reviewerSelect.addEventListener('change', function () { teamDirty = true; saveTeamSettings().catch(function () {}); });
    investigateDate.addEventListener('change', function () { teamDirty = true; saveTeamSettings().catch(function () {}); });
    document.addEventListener('click', function (e) {
      if (e.target !== unitInput && e.target !== codeInput && !suggestBox.contains(e.target)) {
        suggestBox.classList.add('hidden');
      }
    });

    // 上传文件自动使用当前表格导出文件；点击可重新生成
    const dropZone = $('dropZone');
    dropZone.addEventListener('click', function () { generateExportFile(); });
    dropZone.addEventListener('dragover', function (e) {
      e.preventDefault();
      dropZone.classList.add('dragover');
    });
    dropZone.addEventListener('dragleave', function () { dropZone.classList.remove('dragover'); });
    dropZone.addEventListener('drop', function (e) {
      e.preventDefault();
      dropZone.classList.remove('dragover');
      generateExportFile();
    });

    // 自动生成调查表（6 合一）导出 Excel 作为上传文件
    let selectedSurveyFile = null;
    async function generateSurveyFile() {
      const fileNameEl = $('fileNameSurvey');
      try {
        if (!window.SurveySheets || !window.SurveySheets.exportAllWorkbook) throw new Error('调查表导出模块未就绪');
        const bytes = await window.SurveySheets.exportAllWorkbook();
        const name = '现场调查表汇总_' + new Date().toISOString().slice(0, 10) + '.xlsx';
        selectedSurveyFile = new File([bytes], name, { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
        if (fileNameEl) fileNameEl.textContent = name + '（' + Math.max(1, Math.round(bytes.length / 1024)) + ' KB）';
        log('已生成上传文件：' + name);
      } catch (e) {
        selectedSurveyFile = null;
        if (fileNameEl) fileNameEl.textContent = '';
        log('生成调查表文件失败：' + e.message);
      }
    }

    // 上传流程共用：生成文件 → 确认 → 上传 → 同步团队信息 → 提示保存
    async function runUploadFlow(opts) {
      if (!selectedProject) { log('请先查询项目（输入项目编号后点搜索并选择）'); return; }
      if (!token) { showLogin(); return; }
      const isSurvey = opts.kind === 'survey';
      const btn = $(isSurvey ? 'btnUploadSurvey' : 'btnUpload');
      // 上传前重新生成一次，确保文件为最新内容
      const file = isSurvey ? (await generateSurveyFile(), selectedSurveyFile) : (await generateExportFile(), selectedFile);
      if (!file) { log('上传文件生成失败，请先处理表格内容后重试'); return; }
      // 企业信息输入框预校验：非法值直接拦截（等上传完再报就晚了）
      const entErr = validateEntInputs();
      if (entErr) { alert(entErr); return; }
      if (!isSurvey) {
        const errs = window.SamplingApp && window.SamplingApp.countErrors ? window.SamplingApp.countErrors() : 0;
        if (errs > 0 && !confirm('当前表格有 ' + errs + ' 处校验错误，仍要上传吗？')) return;
      }
      const projectId = selectedProject.id;
      const ok = confirm('确认将文件上传到项目 ' + selectedProject.code + '（ID=' + projectId + '）？\n\n' + file.name);
      if (!ok) return;
      btn.disabled = true;
      btn.textContent = '上传中...';
      try {
        log('开始上传（主体：' + userInfo.userCode + ' ' + userInfo.userName + '）：' + file.name);
        const r = await uploadExcel(token, orgId, projectId, file);
        const d = r.data;
        if (r.status === 200 && d && d.code === '200') {
          const msgs = [];
          if (Array.isArray(d.body)) {
            d.body.forEach(function (m) { if (m && m.msg) msgs.push(String(m.msg)); });
          }
          let syncMsg = '';
          try {
            const syncLogs = await syncProjectInfo(token, projectId, {
              memberIds: selectedMemberIds(),
              users: userList,
              project: selectedProject, // 带 responser（负责人），同步时不移除负责人
              investigatorId: investigatorSelect.value,
              reviewerId: reviewerSelect.value,
              investigateDate: investigateDate.value
            });
            syncLogs.forEach(function (l) { log(l); });
            // 「保留（平台不允许移除）」是预期行为（负责人等受保护成员），不算失败
            const failed = syncLogs.filter(function (l) {
              return l.indexOf('失败') >= 0 && l.indexOf('保留（平台不允许移除') < 0;
            });
            // 弹窗直接列出失败项明细（冒号前短句），不再只给一句「有失败项」
            syncMsg = failed.length > 0 ? '，同步失败项：' + failed.map(function (l) { return l.split('：')[0]; }).join('；') : '，团队信息已同步';
          } catch (e) {
            log('同步团队信息失败：' + e.message);
            syncMsg = '，团队信息同步失败：' + e.message;
          }
          // 企业信息：主表格上传带在职人数、调查表上传带接害人数（失败不影响上传主体）
          // 勾选「企业基本信息上传」时，调查表上传额外补全职卫管理部门/负责人/兼职专职人数（见 buildEntBasicPatch）
          let entMsg = '';
          let entCleared = '';   // 本次上传后清空的输入框名（只用于提示文案）
          const entBasicOn = isSurvey && entBasicChecked();
          try {
            entMsg = await applyEntInfoOnUpload(
              token, orgId, selectedProject,
              entPayloadForUpload(isSurvey ? 'survey' : 'main', entInputValues()),
              function (patch) {
                // 平台已确认写入成功 → 清空本次写入的那个输入框及其缓存值（详见 clearEntInputsAfterUpload）
                entCleared = clearEntInputsAfterUpload(patch, projectId);
              },
              entBasicOn ? function (rec) { return buildEntBasicPatch(rec); } : null
            );
            if (entMsg) {
              if (entCleared) entMsg += '；已清空「' + entCleared + '」输入框';
              log(entMsg);
            }
          } catch (e) {
            entMsg = '企业信息写入异常：' + e.message;
            log(entMsg);
          }
          // 上传成功后询问是否保存数据，默认名：年份 + 公司名称（从项目编号取年份前缀）
          try {
            const ym = /^BTC(\d{2})/.exec(selectedProject.code || '');
            const defaultName = (ym ? ym[1] + '年' : '') + (selectedProject.belongInspectName || '');
            if (window.SamplingApp && window.SamplingApp.promptSave) {
              setTimeout(function () {
                window.SamplingApp.promptSave(defaultName).catch(function () {});
              }, 300);
            }
          } catch (e) {}
          // 调查表上传：导入成功后补传工艺图片（按「工艺名称」匹配平台的生产工艺记录）
          let imgMsg = '';
          if (isSurvey) {
            try {
              imgMsg = await uploadProcessImages(token, orgId, projectId);
              if (imgMsg) log(imgMsg);
            } catch (e) {
              imgMsg = '工艺图片上传异常：' + e.message;
              log(imgMsg);
            }
          }
          if (msgs.length > 0) {
            log('导入完成，但有以下提示：' + msgs.join('；'));
            alert('导入完成，提示：\n' + msgs.join('\n') + (entMsg ? '\n' + entMsg : '') + (imgMsg ? '\n' + imgMsg : ''));
          } else {
            log('导入成功！' + syncMsg + (entMsg ? '，' + entMsg : '') + (imgMsg ? '，' + imgMsg : ''));
            alert('导入成功！' + syncMsg + (entMsg ? '，' + entMsg : '') + (imgMsg ? '，' + imgMsg : ''));
          }
        } else {
          const msg = (d && (d.message || d.msg)) || ('HTTP ' + r.status);
          log('导入失败：' + msg);
          alert('导入失败：' + msg);
        }
      } catch (e) {
        // 传输中断（如网络断开）时平台可能已接收成功，提示核实而非断言失败
        const aborted = e && (e.name === 'AbortError' || /abort/i.test(e.message || ''));
        log(aborted ? '上传连接中断（结果未知，请到平台核实）' : '上传异常：' + e.message);
        alert(aborted ? '上传连接中断，结果未知：\n请到平台该项目下核实是否已导入成功，避免重复上传。' : '上传异常：' + e.message);
        if (!aborted) log('上传异常：' + e.message);
      } finally {
        btn.disabled = false;
        btn.textContent = isSurvey ? '上传调查表（6 合一）' : '上传测点布局调查';
      }
    }

    $('btnUpload').addEventListener('click', function () { runUploadFlow({ kind: 'main' }); });
    $('btnUploadSurvey').addEventListener('click', function () { runUploadFlow({ kind: 'survey' }); });

    // 调查表上传区：点击/拖放重新生成
    const dropZoneSurvey = $('dropZoneSurvey');
    dropZoneSurvey.addEventListener('click', function () { generateSurveyFile(); });
    dropZoneSurvey.addEventListener('dragover', function (e) {
      e.preventDefault();
      dropZoneSurvey.classList.add('dragover');
    });
    dropZoneSurvey.addEventListener('dragleave', function () { dropZoneSurvey.classList.remove('dragover'); });
    dropZoneSurvey.addEventListener('drop', function (e) {
      e.preventDefault();
      dropZoneSurvey.classList.remove('dragover');
      generateSurveyFile();
    });

    // 视图切换：返回采样计划 / 退出登录
    $('xcdcBack').addEventListener('click', hideUpload);

    // ---------- 噪声数据结果页签（按「数据上传」选中的项目生成检测报告 → 解析「工作场所噪声评判结果」表） ----------
    let jszipPromise = null;
    function ensureJSZip() {
      if (window.JSZip) return Promise.resolve(window.JSZip);
      if (!jszipPromise) {
        const load = (window.SamplingApp && window.SamplingApp.loadScript) || Util.loadScript;
        jszipPromise = load('js/jszip.min.js')
          .then(function () {
            if (!window.JSZip) throw new Error('jszip 加载失败');
            return window.JSZip;
          });
      }
      return jszipPromise;
    }

    const W_NS = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
    function noiseEsc(s) {
      return Util.escHtml(s);
    }
    // 取文本：把 <w:br/> 还原成换行。两种写法都要支持：
    //   ① 非标准——<w:br/> 写在 <w:t> **内部**（WPS/Aspose 的部分模板，如检测报告）
    //   ② 标准  ——<w:br/> 与 <w:t> **并列**在 <w:r> 里（如检评报告）
    // 故改为按文档顺序遍历，遇到 t 取文本、遇到 br 换行、遇到 tab 制表；
    // 单元格内多个段落由 xmlText 用空格连接。
    function runText(scope) {
      let out = '';
      (function walk(node) {
        const kids = node.childNodes;
        for (let n = 0; n < kids.length; n++) {
          const c = kids[n];
          if (c.nodeType === 3) { out += c.nodeValue || ''; continue; }
          if (c.nodeType !== 1) continue;
          const ln = c.localName;
          if (ln === 't') {
            const tk = c.childNodes;
            for (let k = 0; k < tk.length; k++) {
              const d = tk[k];
              if (d.nodeType === 3) out += d.nodeValue || '';
              else if (d.nodeType === 1 && d.localName === 'br') out += '\n';
              else if (d.nodeType === 1 && d.localName === 'tab') out += '\t';
            }
          } else if (ln === 'br') {
            out += '\n';
          } else if (ln === 'tab') {
            out += '\t';
          } else {
            walk(c);
          }
        }
      })(scope);
      return out.replace(/[ \t]+/g, ' ').replace(/\n{2,}/g, '\n').replace(/^\s+|\s+$/g, '');
    }
    function xmlText(el) {
      const ps = el.getElementsByTagNameNS(W_NS, 'p');
      if (ps.length) {
        const parts = [];
        for (let i = 0; i < ps.length; i++) {
          const t = runText(ps[i]);
          if (t) parts.push(t);
        }
        return parts.join(' ');
      }
      return runText(el);
    }

    // 兼容平台列表接口的多种返回结构（统一走共享工具 js/util.js）
    function pickListRows(j) {
      return Util.pickRows(j);
    }

    // 通用：从 **document.xml 字符串**里取「某个标题段落下面的第一张表」，返回 { rows: string[][] }
    //
    // ★ 性能（2026-10-18）：把「解压 + 取 document.xml」与「解析表格」拆开。
    //   原来每次取表都要 JSZip.loadAsync + async('string') 重新解压整份 docx；
    //   现在上层可以复用**已经解出来**的 xml（检评报告后处理本来就已解压过一次），
    //   省掉重复解压（1MB+ 的 document.xml 解压+转字符串是百毫秒级开销）。
    function parseTableAfterHeadingFromXml(xml, keyword, label) {
      const doc = new DOMParser().parseFromString(xml, 'application/xml');
      const body = doc.getElementsByTagNameNS(W_NS, 'body')[0];
      if (!body) throw new Error('docx 结构异常：缺少 body');
      let prev = '';
      let tbl = null;
      for (let i = 0; i < body.childNodes.length; i++) {
        const node = body.childNodes[i];
        if (node.nodeType !== 1) continue;
        if (node.localName === 'p') {
          const t = xmlText(node);
          if (t) prev = t;
        } else if (node.localName === 'tbl' && prev.indexOf(keyword) >= 0) {
          tbl = node;
          break;
        }
      }
      if (!tbl) throw new Error('未在报告中找到「' + (label || keyword) + '」表');
      const rows = [];
      const trs = tbl.getElementsByTagNameNS(W_NS, 'tr');
      for (let i = 0; i < trs.length; i++) {
        const cells = [];
        const tcs = trs[i].getElementsByTagNameNS(W_NS, 'tc');
        for (let j = 0; j < tcs.length; j++) cells.push(xmlText(tcs[j]));
        rows.push(cells);
      }
      return { rows: rows };
    }

    // 从 docx 里取「某个标题段落下面的第一张表」，返回 { rows: string[][] }
    async function parseTableAfterHeadingFromDocx(buf, keyword, label) {
      const xml = await readDocxDocumentXml(buf);
      return parseTableAfterHeadingFromXml(xml, keyword, label);
    }

    // 解压 docx 并取出 word/document.xml 字符串（唯一入口，便于上层复用）
    // 取文本优先用原生 TextDecoder（比 JSZip 自带 async('string') 的 JS 版 UTF-8 解码器快）；
    // Node 的 vm 沙箱可能没有 TextDecoder，做存在性判断后回退。
    async function readDocxDocumentXml(buf) {
      const JSZip = await ensureJSZip();
      const zip = await JSZip.loadAsync(buf);
      const f = zip.file('word/document.xml');
      if (!f) throw new Error('docx 结构异常：缺少 word/document.xml');
      if (typeof TextDecoder !== 'undefined') {
        try {
          const u8 = await f.async('uint8array');
          return new TextDecoder('utf-8').decode(u8);
        } catch (e) { /* 回退到 JSZip 自带解码 */ }
      }
      return await f.async('string');
    }

    // 从**检评报告**里取「表7-3 噪声检测结果分析与评价」表，返回 { rows: string[][] }
    // （该表列：测量编号 / 单元·工作场所 / 岗位·工种·测量点·对象 / 接触时间 / 噪声性质 /
    //   测量范围 / 测量结果 / LEX,8h / LEX,W / 接触限值 / 单项结论）
    async function parsePjbgNoiseTableFromDocx(buf) {
      return parseTableAfterHeadingFromDocx(buf, '噪声检测结果分析与评价', '噪声检测结果分析与评价');
    }

    // 展示用表格：① 删除「测量起止时间」列；② 在「岗位/工种/测量点/对象」右侧新增一列，
    // 把该列每个值按首个「/」拆分——「/」前留在原列，「/」后放进新列（多行值逐行处理）
    function splitBySlash(v) {
      const lines = String(v === null || v === undefined ? '' : v).split('\n');
      const left = [];
      const right = [];
      for (const ln of lines) {
        const i = ln.indexOf('/');
        if (i < 0) { left.push(ln); right.push(''); }
        else { left.push(ln.slice(0, i).trim()); right.push(ln.slice(i + 1).trim()); }
      }
      return { left: left.join('\n'), right: right.join('\n') };
    }

    function buildNoiseDisplay(rows) {
      if (!rows || !rows.length) return rows;
      const head = rows[0];
      const norm = (s) => String(s === null || s === undefined ? '' : s).replace(/\s+/g, '');
      const dropIdx = head.findIndex((h) => norm(h).indexOf('测量起止时间') >= 0);
      let jobIdx = head.findIndex((h) => norm(h).indexOf('岗位/工种/测量点/对象') >= 0);
      if (jobIdx < 0) jobIdx = head.findIndex((h) => norm(h).indexOf('岗位/工种') >= 0);
      const outHead = [];
      const keep = [];
      for (let i = 0; i < head.length; i++) {
        if (i === dropIdx) continue;
        outHead.push(head[i]);
        keep.push(i);
        if (i === jobIdx) outHead.push('测量点/对象');
      }
      const out = [outHead];
      for (let r = 1; r < rows.length; r++) {
        const src = rows[r] || [];
        const line = [];
        for (const i of keep) {
          const v = src[i] || '';
          if (i === jobIdx) {
            const sp = splitBySlash(v);
            line.push(sp.left, sp.right);
          } else {
            line.push(v);
          }
        }
        out.push(line);
      }
      return out;
    }

    // 渲染并记录当前「展示用行」（供「非噪声岗位同步」复用）
    function renderNoiseTable(box, data) {
      const rows = buildNoiseDisplay((data && data.rows) || []);
      noiseCache = { projectId: (noiseCache && noiseCache.projectId) || null, data: data, display: rows };
      renderNoiseSummary(rows);
      if (!rows.length) { box.innerHTML = ''; return rows; }
      const head = rows[0];
      let html = '<table class="noise-table"><thead><tr>';
      for (const h of head) html += '<th>' + noiseEsc(h) + '</th>';
      html += '</tr></thead><tbody>';
      for (let r = 1; r < rows.length; r++) {
        html += '<tr>';
        for (let c = 0; c < head.length; c++) html += '<td>' + noiseEsc(rows[r][c] || '') + '</td>';
        html += '</tr>';
      }
      html += '</tbody></table>';
      box.innerHTML = html;
      return rows;
    }

    let noiseCache = null; // { projectId, data }
    // 噪声页签加载时下载到的检评报告原始 docx：供「生成检评报告」直接复用，
    // 省掉平台再生成一次（约 20s）。{ projectId, buf, fileName, modName }
    let pjbgBufCache = null;
    let noiseLoading = false;

    // ---------- 噪声筛选汇总（非噪声车间岗位 / 噪声超标车间岗位） ----------
    // 判定口径与参考表《非噪声岗位筛选26.10.10.xlsx》右侧数组公式一致：
    //   ① 分组键 = 单元/工作场所 + 岗位（岗位取「/」前的部分；同一岗位的多个测量点合并为一组）；
    //   ② 组内取值 = 该组内所有含「LEX」列的所有数值的最大值（非数字按 0 处理）；
    //   ③ 非噪声车间岗位：max < 80；噪声超标车间岗位：max ≥ 85；
    //   ④ 输出按车间合并：同一车间只出现一次，岗位归到该车间下——
    //      「车间岗位、岗位；车间岗位、岗位」（车间之间用「；」，同车间岗位之间用「、」）。
    const NOISE_NON_LIMIT = 80;
    const NOISE_OVER_LIMIT = 85;
    const NOISE_SUMMARY_DEFS = [
      { key: 'nonNoise', title: '非噪声车间岗位', rule: 'LEX 最大值 < 80 dB', cls: 'ok' },
      { key: 'noiseWorkshop', title: '噪声车间岗位', rule: 'LEX 最大值 ≥ 80 dB', cls: 'noise' },
      { key: 'overWorkshop', title: '噪声超标车间岗位', rule: 'LEX 最大值 ≥ 85 dB', cls: 'over' },
    ];
    let noiseSummaryData = null;

    function beforeSlash(v) {
      const s = String(v === null || v === undefined ? '' : v);
      const i = s.indexOf('/');
      return (i < 0 ? s : s.slice(0, i)).trim();
    }
    function noiseNorm(s) { return String(s === null || s === undefined ? '' : s).replace(/\s+/g, ''); }

    // 由「展示用噪声行」计算筛选结果，返回
    //   { nonNoise, noiseWorkshop, overWorkshop, groupCount, quietPoints }
    // 前三个列表均为「按车间合并」后的数组：[{ unit, jobs:[岗位, ...] }]（顺序=首次出现）
    // quietPoints：「噪声岗位（LEX 最大值 ≥ 80 dB）内 3 个测量值全部 < 80 dB 的点位」，
    //   形如 [{ unit, job, point }]，供报告后处理②c 按点位删「噪声」危害（见 noiseInfoFromRows）。
    //   ⚠ 仅在「多个点位挤在同一单元格（多行）」的形态下判定；单点位单行不参与。
    function buildNoiseSummary(display) {
      const out = { nonNoise: [], noiseWorkshop: [], overWorkshop: [], groupCount: 0, quietPoints: [] };
      if (!display || display.length < 2) return out;
      const head = display[0];
      const findCol = function (name) {
        const want = noiseNorm(name);
        // 优先精确匹配：「测量点/对象」是「岗位/工种/测量点/对象」的子串，含匹配会取错列
        for (let i = 0; i < head.length; i++) if (noiseNorm(head[i]) === want) return i;
        for (let i = 0; i < head.length; i++) if (noiseNorm(head[i]).indexOf(want) >= 0) return i;
        return -1;
      };
      const cUnit = findCol('单元/工作场所');
      const cJob = findCol('岗位/工种/测量点/对象');
      if (cUnit < 0 || cJob < 0) return out;
      const cPoint = findCol('测量点/对象');     // buildNoiseDisplay 从「岗位/工种/测量点/对象」拆出来的点位列
      const cResult = findCol('测量结果');       // 「测量结果[dB(A)]」，每个点位 3 个测量值（多行）
      const lexCols = [];
      for (let i = 0; i < head.length; i++) if (noiseNorm(head[i]).indexOf('LEX') >= 0) lexCols.push(i);

      // 展开为「车间+岗位」分组：先对空单元格做向下填充（与参考表 SCAN 一致，兼容纵向合并的报表），
      // 再把单元格内的多行值逐行对齐展开（与「非噪声岗位同步」同一方式）
      const groups = new Map();
      let prevUnit = '';
      let prevJob = '';
      for (let r = 1; r < display.length; r++) {
        const src = display[r] || [];
        const unitRaw = String(src[cUnit] === null || src[cUnit] === undefined ? '' : src[cUnit]);
        const jobRaw = String(src[cJob] === null || src[cJob] === undefined ? '' : src[cJob]);
        if (unitRaw.trim() !== '') prevUnit = unitRaw;
        if (jobRaw.trim() !== '') prevJob = jobRaw;
        const units = splitLines(prevUnit);
        const jobs = splitLines(prevJob);
        // ★ 行内多点位（2026-10-19）：报告有时把同一岗位的多个点位挤在**同一个单元格**里（单元格多行），
        //   「测量结果[dB(A)」单元格同样多行，且每个点位固定 3 个测量值、按点位顺序排列
        //   （第 1 个点位 = 前 3 行，第 2 个点位 = 第 4~6 行，依次类推）。
        //   ⚠ 只在「**多行挤在一格**」（点位数 n > 1 且点数 = 结果行数 / 3）时才做点位级判断；
        //     单点位单行的报告形态（n = 1）**不参与**（用户 2026-10-19 明确要求），整行跳过。
        const points = cPoint >= 0 ? splitLines(src[cPoint]) : [];
        const results = cResult >= 0 ? splitLines(src[cResult]) : [];
        const n = Math.max(units.length, jobs.length, points.length);
        const perPoint = (n > 1 && points.length === n && results.length === n * 3);
        for (let i = 0; i < n; i++) {
          const unit = pickLine(units, i);
          const job = beforeSlash(pickLine(jobs, i));
          if (!unit && !job) continue;
          const key = noiseNorm(unit) + '\u0001' + noiseNorm(job);
          let g = groups.get(key);
          if (!g) { g = { unit: unit, job: job, max: 0, pts: [] }; groups.set(key, g); }
          for (let k = 0; k < lexCols.length; k++) {
            const num = toNum(pickLine(splitLines(src[lexCols[k]]), i));
            if (num !== null && num > g.max) g.max = num;
          }
          if (perPoint) {
            const point = beforeSlash(pickLine(points, i));
            if (point) {
              const vals = [];
              for (let k = 0; k < 3; k++) vals.push(toNum(results[i * 3 + k]));
              g.pts.push({ point: point, vals: vals });
            }
          }
        }
      }
      out.groupCount = groups.size;
      // 按车间合并：同一车间只保留一项，岗位依次归入（车间/岗位均保持首次出现顺序）
      const addGroup = function (list, index, g) {
        let ws = index.get(g.unit);
        if (!ws) { ws = { unit: g.unit, jobs: [] }; index.set(g.unit, ws); list.push(ws); }
        if (g.job && ws.jobs.indexOf(g.job) < 0) ws.jobs.push(g.job);
      };
      const nonIdx = new Map();
      const noiseIdx = new Map();
      const overIdx = new Map();
      groups.forEach(function (g) {
        if (g.max < NOISE_NON_LIMIT) addGroup(out.nonNoise, nonIdx, g);
        else addGroup(out.noiseWorkshop, noiseIdx, g);   // 噪声车间岗位：LEX 最大值 ≥ 80 dB
        if (g.max >= NOISE_OVER_LIMIT) addGroup(out.overWorkshop, overIdx, g);
        // ★ 低噪点位（2026-10-19）：岗位整体属噪声岗位（LEX 最大值 ≥ 80 dB）时，逐点位看
        //   「测量结果[dB(A)」的 3 个测量值：全部 < 80 dB → 该点位实际不接触噪声，
        //   报告表4-1 里该点位那一行也应删掉「噪声」危害。
        //   前提：同岗位**至少有一个点位 3 个测量值全部 ≥ 80 dB**，否则不删——
        //   否则整岗位的噪声会被删光（那种情况本就该走②的整岗位规则）。
        if (g.max >= NOISE_NON_LIMIT && g.pts.length) {
          let hasNoisePoint = false;
          const low = [];
          for (const p of g.pts) {
            if (p.vals.length !== 3 || p.vals.some(function (v) { return v === null; })) continue;
            if (p.vals.every(function (v) { return v >= NOISE_NON_LIMIT; })) hasNoisePoint = true;
            else if (p.vals.every(function (v) { return v < NOISE_NON_LIMIT; })) low.push(p);
          }
          if (hasNoisePoint) {
            low.forEach(function (p) {
              out.quietPoints.push({ unit: g.unit, job: g.job, point: p.point });
            });
          }
        }
      });
      return out;
    }

    // ---------- 噪声信息：一次解析，多方复用（★ 2026-10-18 性能重构） ----------
    // 背景：检评报告后处理 ②⑤⑥ 三步分别需要「非噪声岗位 pairs」「非噪声岗位文本」
    //   「噪声车间岗位分组」。原先各自独立调用一个 *FromBuf 函数，每次都要
    //   JSZip.loadAsync 解压整份 docx + DOMParser 解析全文 + 遍历噪声表 ——
    //   同一份报告被完整解析 3 次（浏览器实测 300ms+）。
    // 现在统一由 noiseInfoFromXml / noiseInfoFromBuf 算一次，各调用方共享结果。
    //
    // 信息对象结构：{ rows, nonNoiseGroups, nonNoiseText, nonNoisePairs, noiseWorkshopGroups,
    //                 quietPointPairs, ok }
    //   · nonNoiseGroups / noiseWorkshopGroups：按车间合并后的 [{ unit, jobs:[...] }]
    //   · nonNoisePairs：展开成 [{ unit, job }]（② 删除噪声危害用）
    //   · nonNoiseText：'车间岗位、岗位；车间岗位'（⑤ 注段落拼接用），空列表为 ''
    //   · quietPointPairs：噪声岗位内「3 个测量值全部 < 80 dB」的点位 [{ unit, job, point }]
    //     （②c 按点位删除噪声危害用；point 取噪声表「测量点/对象」列，即「岗位/点位」里「/」后的部分）
    //   · ok：是否成功解析（false 时各字段为空，调用方按「缺数据」保守处理）

    // 由「已解析出的噪声表行」构造噪声信息（noiseInfoFromXml / noiseInfoFromBuf /
    // loadNoiseResult 共用同一份组装逻辑，避免三处各写一遍）
    function noiseInfoFromRows(rows) {
      const data = { rows: rows || [] };
      const s = buildNoiseSummary(buildNoiseDisplay(data.rows));
      const nonNoiseGroups = (s && s.nonNoise) || [];
      const noiseWorkshopGroups = (s && s.noiseWorkshop) || [];
      const pairs = [];
      for (const g of nonNoiseGroups) for (const j of g.jobs) pairs.push({ unit: g.unit, job: j });
      const text = nonNoiseGroups.length
        ? nonNoiseGroups.map(function (g) { return g.unit + g.jobs.join('、'); }).join('；')
        : '';
      return {
        rows: data.rows,
        nonNoiseGroups: nonNoiseGroups,
        nonNoiseText: text,
        nonNoisePairs: pairs,
        noiseWorkshopGroups: noiseWorkshopGroups,
        quietPointPairs: (s && s.quietPoints) || [],
        ok: true,
      };
    }

    function noiseInfoFromXml(xml) {
      try {
        const data = parseTableAfterHeadingFromXml(xml, '噪声检测结果分析与评价', '噪声检测结果分析与评价');
        return noiseInfoFromRows(data.rows);
      } catch (e) {
        return { rows: [], nonNoiseGroups: [], nonNoiseText: '', nonNoisePairs: [], noiseWorkshopGroups: [], quietPointPairs: [], ok: false };
      }
    }

    // 按 docx buffer 取噪声信息；同一个 buf 只解析一次（后处理链里 ②⑤⑥ 共用同一份 buffer）。
    // 用单个槽位以 buf 对象身份为键；同一时刻只会处理一份报告，避免内存堆积。
    let noiseInfoCache = { buf: null, info: null };
    async function noiseInfoFromBuf(buf) {
      if (noiseInfoCache.buf === buf && noiseInfoCache.info) return noiseInfoCache.info;
      let info;
      try {
        const xml = await readDocxDocumentXml(buf);
        info = noiseInfoFromXml(xml);
      } catch (e) {
        info = { rows: [], nonNoiseGroups: [], nonNoiseText: '', nonNoisePairs: [], noiseWorkshopGroups: [], quietPointPairs: [], ok: false };
      }
      noiseInfoCache = { buf: buf, info: info };
      return info;
    }

    // 只读缓存：同一 buf 已解析过就返回，否则返回 null（**不做任何 IO**）。
    // 供「生成检评报告」优先复用「噪声数据结果」页签已解析的结果（用户先看过噪声页签时零额外解析）。
    function getCachedNoiseInfo(buf) {
      return (noiseInfoCache.buf === buf && noiseInfoCache.info) ? noiseInfoCache.info : null;
    }

    // 从**检评报告** docx 直接算出「非噪声车间岗位」清单（供「生成检评报告」删除噪声危害用，
    // 不依赖是否打开过噪声页签）。返回 [{ unit, job }]；解析失败返回 null。
    async function noiseNonNoiseFromBuf(buf) {
      const info = await noiseInfoFromBuf(buf);
      return info.ok ? info.nonNoisePairs : null;
    }

    // 同上，但返回「按车间合并」的分组（供注段落拼接文本用）：[{ unit, jobs:[...] }]；
    // 解析失败或无内容时返回 []（不返回 null，便于直接判空）。
    async function noiseNonNoiseGroupsFromBuf(buf) {
      const info = await noiseInfoFromBuf(buf);
      return info.nonNoiseGroups || [];
    }

    // 「非噪声车间岗位」文本：车间岗位、岗位；车间岗位、岗位（与页签「筛选结果」显示一致）；
    // 空列表返回 ''（供「生成检评报告」判断是否整段删除）。
    async function noiseNonNoiseTextFromBuf(buf) {
      const info = await noiseInfoFromBuf(buf);
      return info.nonNoiseText || '';
    }

    // 「噪声车间岗位」（LEX 最大值 ≥ 80 dB）分组：[{ unit, jobs:[...] }]；解析失败或无内容返回 []。
    // 供「生成检评报告」后处理⑥判断「噪声车间岗位为空 → 删噪声类评价依据/警示标识条目」。
    async function noiseWorkshopGroupsFromBuf(buf) {
      const info = await noiseInfoFromBuf(buf);
      return info.noiseWorkshopGroups || [];
    }

    // 把「按车间合并」的结果拼成文本：车间岗位、岗位；车间岗位、岗位
    function noiseSummaryText(list) {      return (list || []).map(function (g) { return g.unit + g.jobs.join('、'); }).join('；');
    }
    // 岗位总数（用于计数徽标）
    function noiseSummaryJobCount(list) {
      return (list || []).reduce(function (n, g) { return n + g.jobs.length; }, 0);
    }

    function copyNoiseText(text, btn) {
      const t = String(text === null || text === undefined ? '' : text);
      const ok = function () {
        if (!btn) return;
        const old = btn.textContent;
        btn.textContent = '已复制';
        setTimeout(function () { btn.textContent = old; }, 1200);
      };
      const fallback = function () {
        try {
          const ta = document.createElement('textarea');
          ta.value = t;
          ta.style.position = 'fixed';
          ta.style.opacity = '0';
          document.body.appendChild(ta);
          ta.select();
          document.execCommand('copy');
          document.body.removeChild(ta);
          ok();
        } catch (err) { alert('复制失败，请手动选择复制。'); }
      };
      if (navigator.clipboard && navigator.clipboard.writeText) navigator.clipboard.writeText(t).then(ok, fallback);
      else fallback();
    }

    function renderNoiseSummary(display) {
      const box = document.getElementById('noise-summary');
      if (!box) return;
      const s = buildNoiseSummary(display);
      noiseSummaryData = s;
      const hasAny = s.nonNoise.length || s.noiseWorkshop.length || s.overWorkshop.length;
      if (!hasAny) { box.innerHTML = ''; box.classList.add('hidden'); return; }
      let html = '<div class="noise-summary-head"><strong>筛选结果</strong>'
        + '<span class="noise-summary-hint">同一「车间+岗位」合并全部测量点，取所有 LEX 列的最大值（非数字按 0）；结果按车间合并</span>'
        + '<button class="btn small ghost" data-copy-all="1">复制全部</button></div>';
      html += '<div class="noise-summary-grid">';
      for (const d of NOISE_SUMMARY_DEFS) {
        const list = s[d.key] || [];
        const n = noiseSummaryJobCount(list);
        html += '<div class="noise-summary-item ' + d.cls + '">'
          + '<div class="nsi-head"><span class="nsi-title">' + d.title + '</span>'
          + '<span class="nsi-rule">' + d.rule + '</span>'
          + '<span class="nsi-count" title="共 ' + list.length + ' 个车间 / ' + n + ' 个岗位">' + n + '</span>'
          + '<button class="btn small ghost nsi-copy" data-copy-key="' + d.key + '">复制</button></div>'
          + '<div class="nsi-body' + (list.length ? '' : ' empty') + '">' + noiseEsc(list.length ? noiseSummaryText(list) : '（无）') + '</div>'
          + '</div>';
      }
      html += '</div>';
      box.innerHTML = html;
      box.classList.remove('hidden');
    }

    const noiseSummaryEl = document.getElementById('noise-summary');
    if (noiseSummaryEl) {
      noiseSummaryEl.addEventListener('click', function (e) {
        const btn = e.target && e.target.closest ? e.target.closest('button') : null;
        if (!btn) return;
        let text = '';
        if (btn.hasAttribute('data-copy-all')) {
          text = NOISE_SUMMARY_DEFS.map(function (d) {
            const list = (noiseSummaryData && noiseSummaryData[d.key]) || [];
            return d.title + '（' + d.rule + '）：' + (list.length ? noiseSummaryText(list) : '无');
          }).join('\n');
        } else {
          const key = btn.getAttribute('data-copy-key');
          text = noiseSummaryText((noiseSummaryData && noiseSummaryData[key]) || []);
        }
        if (!text) return;
        copyNoiseText(text, btn);
      });
    }

    // 页签按钮加载态：加载中显示「数据加载中」+ 转圈，结束恢复文案
    const noiseTabBtnEl = document.querySelector('.tab[data-tab="noise"]');
    function setNoiseTabLoading(on) {
      if (!noiseTabBtnEl) return;
      noiseTabBtnEl.innerHTML = on ? '数据加载中<span class="spinner"></span>' : '噪声数据结果';
    }

    // 状态行里的项目标识：项目代码 + 企业名称（有则显示，形如「项目 BTC26-SZJC0971（浙江省黄岩服装机械厂）」）
    function noiseProjectLabel() {
      const code = (selectedProject && selectedProject.code) || '';
      const ent = (selectedProject && selectedProject.belongInspectName) || '';
      return '项目 ' + code + (ent ? '（' + ent + '）' : '');
    }

    // 加载并展示噪声数据结果：项目取「数据上传」中选用的项目
    async function loadNoiseResult(force) {
      const box = document.getElementById('noise-table-box');
      const empty = document.getElementById('noise-empty');
      const status = document.getElementById('noise-status');
      if (!box) return;
      if (!token) { showLogin(); return; }
      if (!selectedProject || !selectedProject.id) {
        box.innerHTML = '';
        renderNoiseSummary(null);
        if (empty) {
          empty.textContent = '尚未在「数据上传」中选择项目。请先到「数据上传」查询并选中项目，再查看噪声数据结果。';
          empty.classList.remove('hidden');
        }
        if (status) status.textContent = '';
        alert('尚未在「数据上传」中选择项目。\n请先到「数据上传」查询并选中项目，再查看噪声数据结果。');
        return;
      }
      const pid = selectedProject.id;
      if (noiseLoading) return;
      if (!force && noiseCache && noiseCache.projectId === pid) {
        if (empty) empty.classList.add('hidden');
        renderNoiseTable(box, noiseCache.data);
        if (status) status.textContent = noiseProjectLabel() + ' · 共 ' + Math.max(0, noiseCache.data.rows.length - 1) + ' 行';
        enterNoisePanel();   // 数据已就绪 → 允许进入
        return;
      }
      noiseLoading = true;
      setNoiseTabLoading(true);
      if (empty) empty.classList.add('hidden');
      box.innerHTML = '';
      if (status) status.textContent = '加载中…（正在生成并解析检评报告）';
      try {
        // 数据源：**检评报告**（不是检测报告）——模板列表 → jpbgYL 生成 → 同路径 .docx 下载
        // → 取「噪声检测结果分析与评价」标题下的表，与报告里的表保持同一格式
        const tr = await apiRequest(
          'GET',
          '/api/pjbgModule?pageNumber=1&pageSize=-1&organizationId=' + encodeURIComponent(orgId || ''),
          { token: token, orgId: orgId, timeout: 60000 }
        );
        const recs = pickListRows(tr.data);
        if (!recs.length) throw new Error('未获取到检评报告模板，请先在平台创建模板');
        let mod = null;
        for (const x of recs) { if (/检评/.test(String(x.name || ''))) { mod = x; break; } }
        if (!mod) mod = recs[0];
        const rr = await apiRequest(
          'GET',
          '/api/jpbgReportData/jpbgYL?projectId=' + encodeURIComponent(pid) +
            '&moduleId=' + encodeURIComponent(mod.id) + '&organizationId=' + encodeURIComponent(orgId || ''),
          { token: token, orgId: orgId, timeout: 300000 }
        );
        const body = rr.data && rr.data.body;
        if (rr.status !== 200 || typeof body !== 'string' || !body) {
          throw new Error((rr.data && (rr.data.message || rr.data.msg)) || ('HTTP ' + rr.status));
        }
        const enc = body.replace(/\.pdf$/i, '.docx').split('/').map(encodeURIComponent).join('/');
        const resp = await fetch(API_BASE + enc, { headers: token ? { Authorization: token } : {} });
        if (!resp.ok) throw new Error('下载检评报告失败(HTTP ' + resp.status + ')');
        const raw = await resp.arrayBuffer();
        const data = await parsePjbgNoiseTableFromDocx(raw);
        // 缓存原始 docx（未做后处理），供「生成检评报告」复用，避免平台再生成一次（~20s）
        pjbgBufCache = {
          projectId: pid,
          buf: raw,
          fileName: String(body).split('/').pop().replace(/\.pdf$/i, '.docx'),
          modName: mod.name || '',
        };
        // ★ 顺手把噪声信息写进缓存：之后「生成检评报告」复用这份 docx 时可直接取用，
        //   不必再 DOMParser 解析一次全文（用户先看过噪声页签是最常见的路径）。
        //   try/catch 兜底：缓存只是加速手段，任何异常都不能影响噪声页签本身的加载。
        try { noiseInfoCache = { buf: raw, info: noiseInfoFromRows(data.rows) }; } catch (e) { noiseInfoCache = { buf: null, info: null }; }
        noiseCache = { projectId: pid, data: data };
        renderNoiseTable(box, data);
        if (status) status.textContent = noiseProjectLabel() + ' · 共 ' + Math.max(0, data.rows.length - 1) + ' 行';
        enterNoisePanel();   // 数据就绪 → 首次进入该页签
      } catch (e) {
        box.innerHTML = '';
        renderNoiseSummary(null);
        if (empty) {
          empty.textContent = '加载失败：' + (e && e.message ? e.message : e);
          empty.classList.remove('hidden');
        }
        if (status) status.textContent = '';
        // 失败时**不进入**该页签，直接弹提示（否则用户停在当前页看不到面板里的错误）
        alert('噪声数据加载失败：' + (e && e.message ? e.message : e));
      } finally {
        noiseLoading = false;
        setNoiseTabLoading(false);
      }
    }

    const noiseTabBtn = document.querySelector('.tab[data-tab="noise"]');
    // 进入「噪声数据结果」页签的策略：
    //   · 未选中项目 → 不允许进入（只提示）
    //   · 已选项目但数据未就绪 → 先在当前页签加载，**加载成功后**才切进去（首次进入）
    //   · 数据已就绪（或本次刷新已完成）→ 正常切换
    // 做法：捕获阶段拦截 app.js 的页签切换；就绪后程序化 click 一次放行。
    let noiseEnterAllowed = false;
    function enterNoisePanel() {
      if (!noiseTabBtn) return;
      noiseEnterAllowed = true;
      try { noiseTabBtn.click(); } finally { noiseEnterAllowed = false; }
    }
    if (noiseTabBtn) {
      noiseTabBtn.addEventListener('click', function (e) {
        if (noiseEnterAllowed) return;                 // 已放行，交给 app.js 正常切换
        e.preventDefault();
        e.stopImmediatePropagation();                  // 拦下 app.js 的页签切换
        loadNoiseResult(false);                        // 未就绪：先加载，成功后再进入
      }, true);
    }
    const noiseRefreshBtn = document.getElementById('noise-refresh');
    if (noiseRefreshBtn) noiseRefreshBtn.addEventListener('click', function () { loadNoiseResult(true); });

    // ---------- 数据报告处理（回退 / 审核） ----------
    // 对应平台「流程审批」里两个界面的「重置审批状态」，走的是**同一个接口**，只是 projectProcess 不同：
    //   检测结果报告单 = 'jcbg'    检测底稿 = 'jcdg'
    //   POST /api/projectXccyLastAffix/resetProjectApprovalProcess
    //   body { belongProject, projectProcess }
    // 审核（通过）走 POST /api/projectJcbgLastAffixNew/audit（单个对象 body，见下）
    //   body [{ belongProject, projectName, suggestions, type, isPass }]
    // 平台原生「重置审批状态」是**三次连续确认**，这里照搬。
    // ⚠️ 回退是**两步**：先重置「检测结果报告单」，再重置「检测底稿」（顺序与人工操作一致）。
    // 回退 = 两步顺序重置（与人工操作一致）：
    //   ① 检测结果报告单（jcbg）→ ② 检测底稿（jcdg）
    // 两者走同一接口，仅 projectProcess 不同。**B 方案**：两步分别以 btc-gxz / btc-wcy 身份执行。
    // 浏览器不持有这些账号令牌，统一从本地「账号桥」(captcha-ocr/bridge.py) 取（桥调账号池，缓存命中不消耗验证码）。
    const RESET_STEPS = [
      { process: 'jcbg', label: '检测结果报告单', account: 'btc-gxz' },
      { process: 'jcdg', label: '检测底稿', account: 'btc-wcy' },
    ];

    // 「数据报告审核」= 一条**跨账号的审批链**（与人工在平台上的操作顺序一致）：
    //   ① btc-jyg：检测底稿页「提交审复核」        POST /api/projectJcdgLastAffixNew/submitJcdgReview {belongProject}
    //   ② btc-wcy：检测底稿 理化复核通过            POST /api/projectJcdgLastAffixNew/audit  {…,type:当前节点,isPass:'1'}   （单对象！）
    //   ③ btc-zhj：检测结果报告单 理化审核通过       POST /api/projectJcbgLastAffixNew/audit  {…,type:当前节点,isPass:'1'}   （单对象！）
    //   ④ btc-gxz：检测结果报告单 理化签发通过       POST /api/projectJcbgLastAffixNew/audit  {…,type:当前节点,isPass:'1'}   （单对象！）
    //
    // ⚠⚠ 实测结论（2026-10-15 真机验证，见 交接文档-现状总览.md §2.1）：
    //   1) 通过类接口是 **`audit`（单个对象 body）**，**不是 `auditAll`**（auditAll 传数组 → HTTP 400）。
    //   2) `type` **必须取「该流程当前节点号」**，来源是 **`getAffix` 返回的 `type`**，不是列表行的 `status`：
    //        · jcdg 的列表 `projectJcdgLastAffixNew` 里 status 恒为 "0"（**不可用**）；
    //        · jcbg 的列表 status 恰好 == getAffix 的 type，但为统一稳妥，两者都用 getAffix 的 type。
    //   3) **jcbg 的 `type` 必须是数字**（如 `2`）；传字符串 `"2"` → 400「Required request body is missing」。
    //      jcdg 对 type 字符串宽容，但统一转 int 最稳。
    //   4) 节点号含义（来自前端 statusFn）：0=理化复核 1=复核 2=理化审核 3=审核 4=职卫复核 5=职卫审核
    //      6=理化签发 7=职卫签发 8=签发 9=第三方审核 10=签章 11=批准。
    //      流程节点序列见 `/api/ohitBusinessProcess`（type=8 检测底稿 = "0"；type=3 报告单 = "2,6"）。
    //   5) 每步通过后节点会推进，所以**下一步前重新调 getAffix** 取新节点号。
    //   6) ★★ 第①步「提交审复核」是**一次提交两条流程**：检测底稿(jcdg) **和** 检测报告单(jcbg) 一起提交
    //      （平台返回 body.msg 形如「检测底稿提交成功！;检测报告提交成功」，前端只弹一次通知）。
    //      ⚠ 若提交后只等 jcdg 就绪就往下走，第③步对 jcbg 审批时会报「该流程已无待审的检测记录」
    //      （jcbg 还没提交完/未流转到位）。**必须等两条 getAffix 都出现待审记录**再继续 → pollBothNodesReady()。
    const AUDIT_STEPS = [
      { kind: 'submitJcdg', account: 'btc-jyg', label: '提交审复核（检测底稿+检测报告单）' },
      { kind: 'auditJcdg', account: 'btc-wcy', label: '检测底稿 理化复核通过' },
      { kind: 'auditJcbg', account: 'btc-zhj', label: '检测结果报告单 理化审核通过' },
      { kind: 'auditJcbg', account: 'btc-gxz', label: '检测结果报告单 理化签发通过' },
    ];

    // ★★ 断点续审：**只审预设链路内的节点**，且**跳过已经走完的步骤**。
    //   平台里两个流程的节点序列（GET /api/ohitBusinessProcess）：
    //     · 检测底稿   type=8 → "0"      （唯一节点：理化复核）
    //     · 检测报告单 type=3 → "2,6"    （理化审核 → 理化签发）
    //   把「节点号」映射回本工具的步骤索引（AUDIT_STEPS 的下标）：
    //     0 = 提交审复核（该步无节点号，靠「两条流程是否已有待审记录」判断）
    //     1 = jcdg 节点 0（理化复核）
    //     2 = jcbg 节点 2（理化审核）
    //     3 = jcbg 节点 6（理化签发）
    //   作用：如果项目**已经在链路中途**（别人/上次已把前面几步走完，或本工具上次只跑到一半），
    //   点「数据报告审核」时不再从头执行，也不再因「该流程已无待审的检测记录」直接失败。
    //   ⚠ 若某流程的当前节点号**不在**上面的映射里（如职卫复核 4 / 第三方审核 9 / 签章 10），
    //     说明项目走的是本工具未覆盖的分支 → **停下并提示**，绝不乱动平台状态。
    const JCDG_NODE_STEP = { '0': 1 };
    const JCBG_NODE_STEP = { '2': 2, '6': 3 };
    const STEP_LABELS = AUDIT_STEPS.map(function (s) { return s.label; });

    // 依据两个流程的**当下节点快照**，算出「还需执行哪些步骤」（返回 AUDIT_STEPS 下标数组）。
    //   nodes = { jcdg: {type,isOver}|null, jcbg: {type,isOver}|null }
    //     null = 该流程无待审记录（未提交 / 已完结被清）
    //   返回 { indices:[…], unknowns:[…], completed:bool }：
    //     · indices  —— 按升序排列、尚未完成的步骤下标；
    //     · unknowns —— 落在预设链路之外的节点描述（非空时应中止，不执行 indices）；
    //     · completed —— 两条流程都**已完结**（isOver=1）⇒ 整条链已走完，无需再审。
    //   判定规则：
    //     · 两条流程**都无**待审记录且**都未完结** → 项目未提交 → 从 ① 提交开始（indices=[0,1,2,3]）。
    //     · **`isOver='1'` 的节点 = 该流程已完结**（实测 1009：jcdg `type=0,isOver=1` 表示理化复核已审完，
    //       列表已进 `isDel=1`）→ **不产生任何步骤**。⚠ 若忽略 isOver 会重复审已完成的节点，
    //       平台报「审批流程中的人员不可重复！」（同一个人不能连审两个节点）。
    //     · jcdg 已完结而 jcbg 仍在 `2`/`6` → 只审 jcbg 剩余节点（不重复提交）。
    //     · 两条**都已完结** → `completed:true`、`indices=[]`（调用方提示「已完结」而非重跑）。
    function planAuditSteps(nodes) {
      nodes = nodes || {};
      function over(n) { return !!n && String(n.isOver) === '1'; }
      // 该流程是否存在**待审节点**（有记录且未完结）
      function live(n) { return n && !over(n) ? n : null; }
      const jn = live(nodes.jcdg), bn = live(nodes.jcbg);
      const unknowns = [];
      const jcdgStep = jn ? JCDG_NODE_STEP[String(jn.type)] : undefined;
      const jcbgStep = bn ? JCBG_NODE_STEP[String(bn.type)] : undefined;
      // 两条都无待审节点：
      if (!jn && !bn) {
        // ⚠ 判「全部完结」必须**报告单有明确完结证据**（与 nextAuditStep 同源规则）：
        //   · 两条都 isOver=1 → 完结；· 报告单 isOver=1、底稿无记录 → 完结；
        //   · **底稿 isOver=1、报告单无记录 → 不判完结**，而是标记「需补提交报告单」
        //     （实测 BTC26-SZJC0722：底稿审完后报告单根本没有记录，被误报「审批链全部完结」）。
        const jo = over(nodes.jcdg), bo = over(nodes.jcbg);
        if (jo && bo) return { indices: [], unknowns: [], completed: true };
        if (bo) return { indices: [], unknowns: [], completed: true }; // 报告单已完结
        if (jo) return { indices: [], unknowns: [], completed: false, needSubmitJcbg: true };
        return { indices: AUDIT_STEPS.map(function (_, i) { return i; }), unknowns: [], completed: false };
      }
      // 链路外的节点 → 记入 unknowns（调用方据此中止）
      if (jn && jcdgStep === undefined) {
        unknowns.push('检测底稿 停在「' + nodeName(jn.type) + '」节点，不在本工具可审范围内');
      }
      if (bn && jcbgStep === undefined) {
        unknowns.push('检测结果报告单 停在「' + nodeName(bn.type) + '」节点，不在本工具可审范围内');
      }
      if (unknowns.length) return { indices: [], unknowns: unknowns, completed: false };
      // 走到这里说明至少有一个流程仍**有待审节点**（未完结）
      // ⇒ 项目已经提交过 → 不再重复执行 ① 提交步骤。
      const idx = [];
      if (jcdgStep !== undefined) idx.push(jcdgStep);
      if (jcbgStep !== undefined) idx.push(jcbgStep);
      // 去重 + 升序（jcdg=1 恒早于 jcbg=2/3，但保持通用性）
      const uniq = [];
      idx.forEach(function (i) { if (uniq.indexOf(i) === -1) uniq.push(i); });
      uniq.sort(function (a, b) { return a - b; });
      return { indices: uniq, unknowns: [], completed: false };
    }

    // ★★ 动态续审：给定**当下节点快照**，返回「下一个该执行的步骤下标」（无则 null）。
    //   —— 与 planAuditSteps（静态一次性规划）不同：审核是**边走边推进**的，
    //      每完成一步平台节点就会前移（如 jcbg 2→6→完结），因此必须**每步后重新求下一步**，
    //      直到走到链路终点（jcbg 签发完结）为止，而**不是**执行完初始规划就收尾。
    //   返回 { index, unknowns, done, why }：
    //     · index —— 下一步的 AUDIT_STEPS 下标；null 表示已无可执行步骤；
    //     · unknowns —— 链外节点（非空应中止）；
    //     · done —— true 表示整链已完结/无可执行步骤（正常收尾）。
    //   规则（与 planAuditSteps 同源，单步版）：
    //     · 两条都无待审节点：有任一条已完结 ⇒ done；否则未提交 ⇒ 下一步 = ① 提交(0)。
    //     · jcdg 未完结且节点在映射内 ⇒ 先做 jcdg 对应步（1）。
    //     · jcbg 未完结且节点在映射内 ⇒ 做 jcbg 对应步（2 或 3）。
    //     · 任一流程有效节点落在映射外 ⇒ unknowns（中止）。
    //     · 优先序：**先 jcdg(1) 再 jcbg(2→3)**（与平台链路一致：底稿复核在报告单审核之前）。
    function nextAuditStep(nodes) {
      nodes = nodes || {};
      function over(n) { return !!n && String(n.isOver) === '1'; }
      function live(n) { return n && !over(n) ? n : null; }
      const jn = live(nodes.jcdg), bn = live(nodes.jcbg);
      if (!jn && !bn) {
        // ⚠ 判「全部完结」必须**报告单有明确完结证据**：
        //   · 两条都 isOver=1 → 完结；
        //   · 报告单 isOver=1、底稿无记录 → 完结（底稿必然早已审完，其记录被清属正常）；
        //   · **底稿 isOver=1、报告单无记录 → 不判完结**，标记 needSubmitJcbg →
        //     由调用方「补提交检测报告单」后继续跑（实测 BTC26-SZJC0722 即此情形）。
        const jo = over(nodes.jcdg), bo = over(nodes.jcbg);
        if (jo && bo) return { index: null, unknowns: [], done: true, why: '已全部完结' };
        if (bo) return { index: null, unknowns: [], done: true, why: '已全部完结' }; // 报告单已完结
        if (jo) return { index: null, unknowns: [], done: false, needSubmitJcbg: true };
        return { index: 0, unknowns: [], done: false, why: '未提交' }; // 从未提交 → 从头
      }
      const unknowns = [];
      const jcdgStep = jn ? JCDG_NODE_STEP[String(jn.type)] : undefined;
      const jcbgStep = bn ? JCBG_NODE_STEP[String(bn.type)] : undefined;
      if (jn && jcdgStep === undefined) unknowns.push('检测底稿 停在「' + nodeName(jn.type) + '」节点，不在本工具可审范围内');
      if (bn && jcbgStep === undefined) unknowns.push('检测结果报告单 停在「' + nodeName(bn.type) + '」节点，不在本工具可审范围内');
      if (unknowns.length) return { index: null, unknowns: unknowns, done: false };
      // 先 jcdg，再 jcbg（jcdg 的 1 恒早于 jcbg 的 2/3）
      if (jcdgStep !== undefined) return { index: jcdgStep, unknowns: [], done: false };
      if (jcbgStep !== undefined) return { index: jcbgStep, unknowns: [], done: false };
      return { index: null, unknowns: [], done: true, why: '无可执行步骤' };
    }

    // 节点号 → 中文名（来自前端 statusFn，见 AUDIT_STEPS 注释）
    const NODE_NAMES = { '0': '理化复核', '1': '复核', '2': '理化审核', '3': '审核', '4': '职卫复核',
      '5': '职卫审核', '6': '理化签发', '7': '职卫签发', '8': '签发', '9': '第三方审核',
      '10': '签章', '11': '批准' };
    function nodeName(t) {
      if (t === null || t === undefined || t === '') return '未开始';
      return NODE_NAMES[String(t)] || ('节点' + t);
    }

    // 审批日期 yyyy-MM-dd（本地时区）
    function auditDateStr() {
      const d = new Date();
      const p = function (n) { return String(n).padStart(2, '0'); };
      return d.getFullYear() + '-' + p(d.getMonth() + 1) + '-' + p(d.getDate());
    }

    // 审批相关接口：list = 提交审复核后轮询用；affix = 读**当前节点号**；audit = 单对象通过
    // ⚠ kind 既可能是 'jcdg'/'jcbg'（简写，如 pollBothNodesReady 里），也可能是 'auditJcdg'/'auditJcbg'
    //   （AUDIT_STEPS 的 kind）。两种都要能正确映射，否则 'jcdg' 会误落到 jcbg 分支。
    function auditPaths(kind) {
      const isJcdg = kind === 'auditJcdg' || kind === 'jcdg';
      return isJcdg
        ? { list: '/api/projectJcdgLastAffixNew',
            affix: '/api/projectJcdgLastAffixNew/getAffix',
            audit: '/api/projectJcdgLastAffixNew/audit' }
        : { list: '/api/projectJcbgLastAffixNew',
            affix: '/api/projectJcbgLastAffix/getAffix',
            audit: '/api/projectJcbgLastAffixNew/audit' };
    }

    // 读某流程「当前节点号」：GET getAffix {belongProject} → body[0].type（+ isOver）
    // 返回 {type, isOver, id}；body 为空时返回 null（说明该流程已完结/无待审记录）。
    async function readAffixNode(bearer, org, affixPath, projectId) {
      const r = await poolApiRequest('GET',
        affixPath + '?belongProject=' + encodeURIComponent(String(projectId)),
        { token: bearer, orgId: org, timeout: 60000 });
      if (r.status !== 200) throw new Error('读取当前审批节点失败(HTTP ' + r.status + ')');
      const rows = (r.data && r.data.body) || [];
      if (!rows.length) return null;
      return { type: rows[0].type, isOver: rows[0].isOver, id: rows[0].id };
    }

    // ★ 读节点（三态：有记录 / 无记录 / **读取失败**）。
    //   「读取失败」绝不能当成「无待审记录」——否则会被下一步判定推成「审批链已全部完结」
    //   （实测 BTC26-SZJC0722：报告单无记录时，底稿一审完就误报全部完结）。
    //   失败重试一次（平台偶发抖动），仍失败返回 { failed:true }，由调用方中止并提示。
    async function readNodeState(bearer, org, affixPath, projectId) {
      let lastMsg = '';
      for (let i = 0; i < 2; i++) {
        try { return { node: await readAffixNode(bearer, org, affixPath, projectId) }; }
        catch (e) {
          lastMsg = (e && e.message) || String(e);
          if (i === 0) await new Promise(function (res) { setTimeout(res, 300); });
        }
      }
      return { failed: true, message: lastMsg };
    }

    // 人类可读描述两个流程的当下节点（用于确认框/提示），如「检测底稿：理化复核 ／ 检测结果报告单：理化审核」
    function describeNodes(jcdgNode, jcbgNode) {
      function one(name, n) {
        return name + '：' + (n ? (nodeName(n.type) + (String(n.isOver) === '1' ? '（已完结）' : '')) : '无待审记录');
      }
      return one('检测底稿', jcdgNode) + ' ／ ' + one('检测结果报告单', jcbgNode);
    }

    // 轮询间隔：**快起步 + 退避**。平台状态机通常在数百毫秒内就推进了，
    // 原来固定 2s 会让每次「等流转」白等近 2s（一条审核链累计可达 8~10s）。
    // 这里首轮 250ms 起、逐轮退避到 1.5s 封顶；opts.intervalMs 显式指定时用固定间隔（自动化测试用）。
    const POLL_RAMP = [250, 400, 600, 800, 1000, 1500];
    function pollDelay(round, opts) {
      if (opts && opts.intervalMs) return opts.intervalMs;
      return POLL_RAMP[Math.min(round, POLL_RAMP.length - 1)];
    }

    // ★ 提交审复核后，**检测底稿(jcdg) 与 检测报告单(jcbg) 是「一起提交」的**
    //   （前端点一次「提交审复核」，平台返回 body.msg 形如
    //    「检测底稿提交成功！;检测报告提交成功」——两条一起）。
    //   因此下一步(②)之前，必须**两条流程都已生成待审记录**才能继续，
    //   否则等到后面 ③ 用 jcbg 审批时会报「该流程已无待审的检测记录」。
    //   本函数轮询等待：jcdg 与 jcbg 的 getAffix **都**出现 type（即都存在待审记录）。
    //   返回 { ready, jcdg, jcbg, waited }；超时未齐则 ready=false（附带各自最新节点，便于提示）。
    // ★ 两条流程**并行**读（原来串行 → 每轮多一个 RTT）
    async function pollBothNodesReady(bearer, org, projectId, opts) {
      opts = opts || {};
      const maxWait = opts.maxWaitMs || 60000;
      const jcdgAffix = auditPaths('jcdg').affix;
      const jcbgAffix = auditPaths('jcbg').affix;
      const start = Date.now();
      let jn = null, bn = null, round = 0;
      while (true) {
        // 任一读失败不致命（当作还没就绪）
        const pair = await Promise.all([
          readAffixNode(bearer, org, jcdgAffix, projectId).catch(function () { return null; }),
          readAffixNode(bearer, org, jcbgAffix, projectId).catch(function () { return null; }),
        ]);
        jn = pair[0]; bn = pair[1];
        if (jn && bn) {
          return { ready: true, jcdg: jn, jcbg: bn, waited: Date.now() - start };
        }
        if (Date.now() - start >= maxWait) {
          return { ready: false, jcdg: jn, jcbg: bn, waited: Date.now() - start };
        }
        await new Promise(function (res) { setTimeout(res, pollDelay(round++, opts)); });
      }
    }

    // 轮询等待「当前节点号」变化：通过后反复读 getAffix，直到 type 变了 / body 消失 / isOver 变 1（最多等 30s）
    async function pollNodeChange(bearer, org, affixPath, projectId, oldType, opts) {
      opts = opts || {};
      const maxWait = opts.maxWaitMs || 30000;
      const start = Date.now();
      let round = 0;
      while (Date.now() - start < maxWait) {
        const node = await readAffixNode(bearer, org, affixPath, projectId);
        if (!node) {
          // 该流程已无待审记录 → 视为已流转（结束/进入下一阶段的新记录）
          return { changed: true, gone: true, waited: Date.now() - start };
        }
        if (String(node.isOver) === '1') {
          return { changed: true, gone: false, over: true, waited: Date.now() - start };
        }
        if (String(node.type) !== String(oldType)) {
          return { changed: true, gone: false, newType: node.type, waited: Date.now() - start };
        }
        await new Promise(function (res) { setTimeout(res, pollDelay(round++, opts)); });
      }
      return { changed: false, waited: Date.now() - start };
    }

    // 轮询等待某流程**出现待审记录**（补提交检测报告单后使用）。
    // 间隔与其它轮询一致（pollDelay：快起步 + 退避）。
    async function pollNodeReady(bearer, org, affixPath, projectId, opts) {
      opts = opts || {};
      const maxWait = opts.maxWaitMs || 60000;
      const start = Date.now();
      let round = 0;
      while (true) {
        try { if (await readAffixNode(bearer, org, affixPath, projectId)) return true; } catch (e) { /* 读失败当作还没就绪 */ }
        if (Date.now() - start >= maxWait) return false;
        await new Promise(function (res) { setTimeout(res, pollDelay(round++, opts)); });
      }
    }

    const ACCOUNT_BRIDGE_URL = 'http://127.0.0.1:8787';

    // 账号桥地址（只解析一次）：
    //  1) localStorage.xcdc_bridge_url 显式覆盖优先（末尾的 /token 由下面统一拼）；
    //  2) 本地服务模式（/api/bridge/ping 有响应）→ 用同源 /api/bridge/token，
    //     本地服务(server.mjs)会在首次使用时**自动拉起** captcha-ocr/bridge.py，
    //     所以点「回退」不需要任何人工启动步骤；
    //  3) 其它情况（直接双击 index.html / 静态部署）→ 直连 127.0.0.1:8787。
    // 同时记录 cloud：Cloudflare Pages Function 的 ping 会带 cloud:true（登录方式据此区分）。
    let bridgeInfoPromise = null;
    function bridgeInfo() {
      if (bridgeInfoPromise) return bridgeInfoPromise;
      bridgeInfoPromise = (async function () {
        try {
          const ov = (typeof localStorage !== 'undefined') && localStorage.getItem('xcdc_bridge_url');
          if (ov) return { url: String(ov).replace(/\/+$/, ''), same: false, cloud: false };
        } catch (e) {}
        try {
          const r = await fetch('/api/bridge/ping', { cache: 'no-store' });
          if (r.ok) {
            let d = null;
            try { d = await r.json(); } catch (e) {}
            return { url: '/api/bridge/token', same: true, cloud: !!(d && d.cloud) };
          }
        } catch (e) {}
        return { url: ACCOUNT_BRIDGE_URL + '/token', same: false, cloud: false };
      })();
      return bridgeInfoPromise;
    }
    function bridgeBase() {
      return bridgeInfo().then(function (i) { return i.url; });
    }

    // 登录方式按运行环境区分（只判定一次）：
    //   · Cloudflare 部署（桥 ping 返回 cloud:true）→ 云端代登录，界面**不显示验证码**；
    //   · 本地打开（file:// / 本地服务）→ **沿用原来的步骤**：浏览器取图 + 手输验证码，
    //     不调云端 /login，也不用 /code 自动填。
    // ⚠ 必须在 bridgeInfo / ACCOUNT_BRIDGE_URL 定义之后调用（否则 let/const 处于 TDZ，
    //   异常会被 try/catch 吞掉，导致永远判成本地）。
    async function initLoginMode() {
      let info = { cloud: false };
      try { info = await bridgeInfo(); } catch (e) {}
      if (info.cloud) return;              // 云端：保持默认（无验证码 UI）
      captchaMode = true;
      captchaAutoOff = true;               // 原步骤：纯手输，不做 OCR 自动填
      const field = $('captchaField');
      if (field) field.hidden = false;
      const link = $('btnCaptchaFallback');
      if (link) link.hidden = true;        // 本地本来就是手输，不需要兜底入口
      if (!uuid) refreshCaptcha();         // showLogin 可能已先取过图，避免重复请求
    }

    // ---------- 账号池共享令牌缓存 ----------
    // 池内账号共用一份令牌缓存：桥预热、主界面登录（登录账号是池成员时）、
    // 回退/审核调用都读写同一份，同一账号不重复登录；
    // 调用时鉴权失效 → 只作废该账号缓存 → 桥重登刷新 → 重试一次。
    const POOL_ACCOUNTS = ['btc-gxz', 'btc-wcy', 'btc-jyg', 'btc-zhj'];
    const poolTokens = {};   // account -> { token, orgId }
    const poolFetching = {}; // account -> Promise（并发去重：冷启动/预热只触发一次桥登录）

    async function fetchBridgeToken(account, opts) {
      opts = opts || {};
      // 缓存命中直接用（可能刚被主界面登录或预热填好）
      if (!opts.force && poolTokens[account]) return poolTokens[account];
      if (!opts.force && poolFetching[account]) return poolFetching[account];
      const p = (async function () {
        const base = await bridgeBase();
        // 同源路由（本地服务 / Cloudflare）带上本会话令牌：
        // 云端版靠它校验调用方是已登录平台的用户（本地版忽略该头）
        const headers = {};
        if (base.charAt(0) === '/') {
          if (token) headers.Authorization = token;
          if (orgId) headers.organizationId = String(orgId);
        }
        let resp;
        try {
          resp = await fetch(base + '/' + encodeURIComponent(account), { cache: 'no-store', headers: headers });
        } catch (e) {
          // fetch 只会抛 TypeError: Failed to fetch，这里换成可操作的提示
          throw new Error('连不上账号桥 ' + base
            + '（若用本地服务打开，请确认 Python 3 可用；也可双击 captcha-ocr/start-bridge.bat 手动启动）');
        }
        if (!resp.ok) {
          let msg = '账号桥返回 HTTP ' + resp.status;
          try { const d = await resp.json(); if (d && d.message) msg = d.message; } catch (e) {}
          throw new Error(msg);
        }
        const d = await resp.json();
        if (!d || !d.token) throw new Error('账号桥未返回 token：' + account);
        const entry = { token: d.token, orgId: d.orgId };
        poolTokens[account] = entry;
        return entry;
      })();
      if (!opts.force) {
        poolFetching[account] = p;
        const clear = function () { delete poolFetching[account]; };
        p.then(clear, clear);
      }
      return p;
    }

    // 平台响应/异常是否像「鉴权失败」（令牌过期、未认证）
    function poolAuthFailed(x) {
      if (!x) return false;
      if (typeof x.status === 'number') {
        if (x.status === 401 || x.status === 403) return true;
        const d = x.data;
        const code = d && d.code != null ? String(d.code) : '';
        if (code === '401' || code === '403') return true;
        const txt = (d && (d.message || d.msg)) || '';
        return /未登录|登录失效|登录已过期|令牌|token|未认证|认证失败/i.test(txt);
      }
      return /未登录|登录失效|登录已过期|令牌|token|未认证|认证失败/i.test(String(x.message || ''));
    }

    // ---------- 池账号「失效 → 自动重登 → 原地重试」----------
    // ★ 不变量：同一账号始终只有**一份令牌对象**——调用方持有的 entry 与 poolTokens[account]
    //   是同一个对象，重登只就地改写它的 token/orgId，持有者无需重新取令牌即可用上新令牌。
    const POOL_RELOGIN_BRIDGE_TRIES = 2; // 桥重登自身也可能因验证码识别失败而失败 → 重试
    const POOL_RELOGIN_RETRIES = 1;      // 单次调用最多重试 1 次（共 2 次尝试）

    // 该账号作废缓存 → 经账号桥**强制重登** → 就地刷新同一个 entry
    async function reloginPoolAccount(account) {
      const prev = poolTokens[account];
      let lastErr = null;
      for (let i = 1; i <= POOL_RELOGIN_BRIDGE_TRIES; i++) {
        try {
          const ne = await fetchBridgeToken(account, { force: true }); // force：绕过缓存真正重登
          const target = prev || ne;
          target.token = ne.token;
          target.orgId = ne.orgId;
          poolTokens[account] = target; // 统一回同一对象，引用保持稳定
          if (i > 1) log('「' + account + '」重新登录成功（第 ' + i + ' 次尝试）');
          return target;
        } catch (e) {
          lastErr = e;
          log('「' + account + '」重新登录失败（第 ' + i + '/' + POOL_RELOGIN_BRIDGE_TRIES + ' 次）：' + e.message);
          if (i < POOL_RELOGIN_BRIDGE_TRIES) await new Promise(function (res) { setTimeout(res, 800 * i); });
        }
      }
      if (prev) poolTokens[account] = prev; // 全部失败：恢复原引用，不留下半残状态
      throw lastErr || new Error('账号「' + account + '」重新登录失败');
    }

    // 按令牌反查所属池账号（失效后据此决定重登哪个账号）
    function poolAccountOfToken(tk) {
      if (!tk) return null;
      for (const acc in poolTokens) {
        const e = poolTokens[acc];
        if (e && e.token === tk) return acc;
      }
      return null;
    }

    // ★★ 池账号专用请求入口（任务「中途不终止」的关键）★★
    // 1) 一律 skipAuthExpired：池令牌失效**与用户主会话无关**，绝不能触发
    //    authExpiredHandler() → notifyAuthExpired() → logout() + 弹「登录已失效」，
    //    那会把主会话和整条回退/审核任务一起打掉。
    // 2) 调用**返回或抛出**鉴权失败 → 该账号经桥自动重登 → 用新令牌**原地重试**
    //    （最多 POOL_RELOGIN_RETRIES 次）：令牌在「探针之后、真正调用之前」过期也不会中断任务。
    // 3) 非鉴权错误（业务报错 / 网络抖动 / 超时）原样返回或抛出，不做无意义重登。
    // ⚠ 判定与重登都只作用于**该账号**，其他步骤、其他账号完全不受影响。
    async function poolApiRequest(method, path, opts) {
      const o = Object.assign({}, opts || {});
      o.skipAuthExpired = true;
      let r = null;
      for (let attempt = 0; ; attempt++) {
        try {
          r = await apiRequest(method, path, o);
        } catch (e) {
          const acc = poolAuthFailed(e) ? poolAccountOfToken(o.token) : null;
          if (!acc || attempt >= POOL_RELOGIN_RETRIES) throw e;
          log('「' + acc + '」调用异常疑似令牌失效（' + e.message + '），账号桥重新登录后重试…');
          await reloginPoolAccount(acc);
          o.token = poolTokens[acc].token;
          o.orgId = poolTokens[acc].orgId;
          continue;
        }
        if (!poolAuthFailed(r)) return r; // 正常业务响应（含业务失败）→ 交上层判断
        const acc = poolAccountOfToken(o.token);
        if (!acc || attempt >= POOL_RELOGIN_RETRIES) return r; // 反查不到 / 重登后仍失效 → 交上层报错
        log('「' + acc + '」令牌已失效（HTTP ' + r.status + '），账号桥重新登录后重试…');
        await reloginPoolAccount(acc);
        o.token = poolTokens[acc].token;
        o.orgId = poolTokens[acc].orgId;
      }
    }

    // 调用前校验池令牌：用轻量已鉴权 GET（auth/info）探测；失效（401/403/登录
    // 失效文案）→ 该账号经桥重登 → 就地刷新 entry（共享缓存同步更新）。
    // 网络抖动等非鉴权错误不触发重登，按原令牌继续。
    async function ensurePoolTokenFresh(account, entry) {
      if (!entry || !entry.token) return entry;
      let probe = null;
      try {
        probe = await apiRequest('GET', '/auth/info', { token: entry.token, orgId: entry.orgId, timeout: 15000, skipAuthExpired: true });
      } catch (e) {
        if (!poolAuthFailed(e)) return entry; // 非鉴权错误：按原令牌继续
        probe = { status: 0, message: e.message };
      }
      // 顺手记录该账号的平台用户 id（用于「同一人不可连审两节点」的前置校验）
      if (probe && probe.status === 200) {
        const u = (probe.data && probe.data.user) || {};
        if (u.id !== undefined && u.id !== null) entry.userId = String(u.id);
        if (u.userName) entry.userName = u.userName;
      }
      if (!poolAuthFailed(probe)) return entry;
      log('「' + account + '」令牌已失效，账号桥重新登录…');
      return await reloginPoolAccount(account);
    }

    // 账号池后台预热：进入软件后静默补齐池内账号令牌。
    // 只补缓存里还没有的账号——主界面登录是池成员时已直接入池（不重复登录），
    // 其余走桥（桥自身也有缓存，命中不耗验证码）。预热后「数据报告回退/审核」
    // 用池内账号即取即用。任何失败（无桥/离线/验证码识别失败）都静默跳过，
    // 不影响主账号的任何功能——调用失效时会再按需重登。
    function warmAccountPool() {
      (async function () {
        for (const acc of POOL_ACCOUNTS) {
          try {
            if (poolTokens[acc]) continue; // 已有共享令牌（登录入池/上轮预热），不重复登录
            await fetchBridgeToken(acc);
            log('账号池后台登录就绪：' + acc);
          } catch (e) {
            log('账号池后台登录跳过（' + acc + '）：' + e.message);
          }
        }
      })();
    }

    // 测试钩子（同 window.__AUDIT_POLL_OPTS__ 模式）：供自动化测试检查/驱动账号池
    if (typeof window !== 'undefined') {
      window.__POOL_TEST__ = {
        poolTokens: poolTokens,
        fetchBridgeToken: fetchBridgeToken,
        ensurePoolTokenFresh: ensurePoolTokenFresh,
      };
    }

    function reportEls() {
      return {
        status: document.getElementById('report-status'),
        empty: document.getElementById('report-empty'),
        info: document.getElementById('report-info'),
        state: document.getElementById('report-state'),
        badge: document.getElementById('report-state-badge'),
        detail: document.getElementById('report-state-detail'),
      };
    }

    // ★★ 「项目状态」文案映射——**照搬平台前端 statusFn(status, isDel)**（权威）。
    //   来源：chunks2/chunk-882c4b1e.js（检测结果报告单列表页模板）
    //     "statusFn",(function(e,t){ return "1"===t ? "已完结" : "0"===e?"待理化复核": … "9"===e?"待第三方审核":"10"===e?"待签章":void 0 })
    //     → 即 statusFn(row.status, row.isDel)；表格「项目状态」列就是它的返回值。
    //   ⚠ 注意两套 statusFn 含义不同，别混：
    //     · 本函数（列表页）：入参是**列表行的 status + isDel**，输出「待理化复核/…/已完结」；
    //     · nodeName()（审批节点）：入参是 **getAffix 的 type**，输出「理化复核/…/批准」。
    //   判定优先级：isDel==='1' 直接「已完结」；否则按 status 查表；查不到时回落到 status 原文。
    const PROJECT_STATUS_TEXT = {
      '0': '待理化复核', '1': '待复核', '2': '待理化审核', '3': '待审核', '4': '待职卫复核',
      '5': '待职卫审核', '6': '待理化签发', '7': '待职卫签发', '8': '待签发', '9': '待第三方审核',
      '10': '待签章',
    };
    function projectStatusText(status, isDel) {
      if (String(isDel) === '1') return '已完结';
      if (status === null || status === undefined || status === '') return '未开始';
      return PROJECT_STATUS_TEXT[String(status)] || ('节点' + status);
    }

    // 查「选中项目」在**检测结果报告单**（数据报告）里的项目状态。
    //   数据源 = 平台「流程审批 → 检测结果报告单」的列表接口
    //     GET /api/projectJcbgLastAffix?pageNumber=1&pageSize=50&isDel=<1|0>&code=<项目编号>
    //   —— 这正是平台该列表页的数据源，行里的 status/isDel 就是 statusFn 的入参。
    //   ⚠「已完结」的项目查 isDel=1，未完结的查 isDel=0（与本文件 findFinishedProject 同一约定）。
    //   返回 { status, isDel, row } 或 null（两个都查不到 = 该编号在报告流程里无记录）。
    async function fetchReportStatus(project) {
      const kw = String((project && project.code) || '').trim();
      if (!kw) return null;
      const q = '?pageNumber=1&pageSize=50&code=' + encodeURIComponent(kw) + '&isDel=';
      // ① 先查未完结（isDel=0），② 再查已完结（isDel=1）
      for (const isDel of ['0', '1']) {
        const r = await apiRequest('GET',
          '/api/projectJcbgLastAffix' + q + isDel,
          { token: token, orgId: orgId, timeout: 60000 });
        if (r.status !== 200) throw new Error('查询项目状态失败(HTTP ' + r.status + ')');
        const rows = (r.data && r.data.body && r.data.body.records) || [];
        const up = kw.toUpperCase();
        for (let i = 0; i < rows.length; i++) {
          if (String(rows[i].code || '').trim().toUpperCase() === up) {
            return { status: rows[i].status, isDel: rows[i].isDel, row: rows[i] };
          }
        }
      }
      return null;
    }

    // 把状态渲染进「数据报告状态」徽标。phase ∈ busy|done|pending|none|err
    function renderReportState(phase, text, detail) {
      const el = reportEls();
      if (!el.state) return;
      el.state.classList.remove('hidden');
      if (el.badge) {
        el.badge.className = 'report-state-badge ' + (phase || '');
        el.badge.textContent = text;
      }
      if (el.detail) el.detail.textContent = detail || '';
    }
    function hideReportState() {
      const el = reportEls();
      if (el.state) el.state.classList.add('hidden');
    }

    // 拉取并渲染「选中项目」的数据报告状态（异步；失败只写徽标，不打断其它逻辑）
    let reportStateSeq = 0;
    // 最近一次成功取到的「数据报告状态」令牌（status|isDel）——用于面板缓存键判定。
    //   · null/undefined = 尚无结果；'none' = 平台无该编号的报告流程记录；'err' = 查询失败。
    let reportStateToken = null;
    async function loadReportStatus() {
      if (!selectedProject || !selectedProject.id) { hideReportState(); return; }
      const seq = ++reportStateSeq;
      renderReportState('busy', '查询中…', '');
      try {
        const st = await fetchReportStatus(selectedProject);
        if (seq !== reportStateSeq) return; // 期间已切换项目 → 丢弃过期结果
        if (!st) {
          reportStateToken = 'none';
          renderReportState('none', '无报告流程记录',
            '该项目在「检测结果报告单」流程中尚无记录（可能未提交审复核）');
          maybeCacheReportPanel();
          return;
        }
        const txt = projectStatusText(st.status, st.isDel);
        const done = String(st.isDel) === '1';
        reportStateToken = String(st.status) + '|' + String(st.isDel);
        // 明细不再暴露 status/isDel 等技术字段，只保留对用户有意义的信息
        renderReportState(done ? 'done' : 'pending', txt, '');
        maybeCacheReportPanel();
      } catch (e) {
        if (seq !== reportStateSeq) return;
        reportStateToken = 'err';
        renderReportState('err', '状态查询失败', (e && e.message) ? e.message : String(e));
      }
    }

    // 状态与报告 PDF 都就绪后，把面板写入缓存（供页签来回切换时秒开）。
    //   用一个「待缓存」标记协调两条异步链（loadReportStatus / autoShowCurrentReport 谁后到谁触发）。
    let reportPdfCacheInfo = null;   // { fname, isOver, node } —— 最近一次自动浏览成功的报告信息
    let reportPdfPending = false;    // 自动浏览是否仍在进行
    function maybeCacheReportPanel() {
      if (!selectedProject || !selectedProject.id) return;
      if (reportStateToken === null || reportStateToken === 'err') return; // 状态未就绪/失败 → 不缓存
      if (reportPdfPending) return;    // PDF 仍在加载 → 等它结束再统一缓存（避免缓存到半成品）
      const el = reportEls();
      const badge = document.getElementById('report-state-badge');
      // 状态徽标仍停在「查询中…」→ 还没出结果，不缓存
      if (el.state && !el.state.classList.contains('hidden') && badge &&
          /查询中/.test(badge.textContent || '')) return;
      cacheReportPanel(reportStateToken, reportPdfCacheInfo);
    }

    function reportProjectLabel() {
      const p = selectedProject;
      if (!p) return '';
      return '项目 ' + (p.code || '') + (p.belongInspectName ? '（' + p.belongInspectName + '）' : '');
    }

    // ★★ 「数据报告处理」面板缓存（2026-10-16）
    //   目的：在**项目未切换**且**数据报告状态未变化**时，来回切换页签无需重复联网，
    //         直接复原上一次已经渲染好的面板（状态徽标 + 报告 PDF 内嵌框）→ 秒开。
    //   · 缓存键 = projectId + '|' + 数据报告状态（status/isDel，未取到则用 'none'）。
    //   · 项目切换 / 状态变化 → 键变化 → 自动失效并重新拉取。
    //   · 状态在**审核/回退**后必然变化，故结算后主动清缓存（invalidateReportCache）。
    //   · 缓存里保存 PDF 的 blob 对象 URL：切回时直接复用 iframe，不再重新下载整份 PDF（最大耗时点）。
    //   ⚠ blob URL 只能被一个 iframe 复用；页面内只保留一个 #report-pdf-frame-wrap，故安全。
    let reportPanelCache = null; // { key, badgePhase, badgeText, detail, statusText, pdf: { objUrl, fname, isOver, node } | null }

    function reportCacheKey() {
      if (!selectedProject || !selectedProject.id) return '';
      return String(selectedProject.id) + '|' + (reportPanelCache && reportPanelCache.stateToken ? reportPanelCache.stateToken : '');
    }
    // 清空「数据报告处理」的面板缓存；若连带释放 PDF blob URL，避免内存泄漏。
    function invalidateReportCache() {
      if (reportPanelCache && reportPanelCache.pdf && reportPanelCache.pdf.objUrl) {
        try { URL.revokeObjectURL(reportPanelCache.pdf.objUrl); } catch (e) {}
      }
      reportPanelCache = null;
    }

    function refreshReportPanel() {
      const el = reportEls();
      if (!el.empty) return;
      if (!selectedProject || !selectedProject.id) {
        // 切换项目 / 清空选择时，隐藏并释放上一次的报告 PDF 预览 + 缓存
        try { if (typeof hideReportPdf === 'function') hideReportPdf(); } catch (e) {}
        invalidateReportCache();
        el.empty.textContent = '尚未在「数据上传」中选择项目。请先到「数据上传」查询并选中项目，再进行数据报告处理。';
        el.empty.classList.remove('hidden');
        if (el.info) el.info.classList.add('hidden');
        if (el.status) el.status.textContent = '';
        hideReportState();
        return;
      }
      // ★ 命中缓存（同一项目 + 同一数据报告状态）→ 直接复原，无需联网
      if (reportPanelCache && reportPanelCache.key === reportCacheKey()) {
        restoreReportPanelFromCache(el);
        return;
      }
      // 未命中 → 隐藏并释放上一次的报告 PDF 预览（也释放其内存）
      try { if (typeof hideReportPdf === 'function') hideReportPdf(); } catch (e) {}
      el.empty.classList.add('hidden');
      if (el.info) {
        // 已按要求去除面板内的文字说明，仅保留工具栏与状态栏
        el.info.innerHTML = '';
        el.info.classList.add('hidden');
      }
      if (el.status) el.status.textContent = '';
      // ★ 新项目/新状态 → 先重置状态令牌与 PDF 信息，避免用上一次项目的残留值错写缓存键
      reportStateToken = null;
      reportPdfCacheInfo = null;
      // 选中项目 → 异步查询并显示「数据报告项目状态」
      loadReportStatus();
      // ★ 默认自动弹出「当前状态下的检测结果报告单」浏览
      //   —— 项目在「检测结果报告单」流程里的审批记录附件（getAffix.jcbgAffix）就是该状态下的数据报告，
      //      与人工在平台「流程审批 → 检测结果报告单 → 附件」看到的完全一致（可能未完结，如「已提交审复核」）。
      autoShowCurrentReport();
    }

    // 从缓存复原面板：状态徽标 + 状态栏文案 + 内嵌 PDF（复用已下载的 blob URL，不重新联网）。
    function restoreReportPanelFromCache(el) {
      const c = reportPanelCache;
      el.empty.classList.add('hidden');
      if (el.info) { el.info.innerHTML = ''; el.info.classList.add('hidden'); }
      if (c.badgePhase) renderReportState(c.badgePhase, c.badgeText, c.detail);
      else hideReportState();
      if (el.status) el.status.textContent = c.statusText || '';
      // 复原内嵌 PDF：blob URL 仍在缓存中 → 直接重建 iframe（不重新下载）。
      //   ⚠ 所有权始终归缓存，这里只是把「当前显示指针」指向它；缓存条目保留该 URL 不被清空。
      const box = document.getElementById('report-pdf-box');
      const wrap = document.getElementById('report-pdf-frame-wrap');
      const pdf = c.pdf;
      if (pdf && pdf.objUrl && wrap && box) {
        reportPdfObjUrl = pdf.objUrl;   // 显示指针指向缓存持有的 blob（缓存条目仍保留 objUrl）
        wrap.innerHTML = '';
        const fr = document.createElement('iframe');
        fr.src = reportPdfObjUrl;
        fr.title = '当下流程中的数据报告';
        wrap.appendChild(fr);
        box.classList.remove('hidden');
        try { document.body.classList.add('pdf-open'); } catch (e) {}
      } else {
        if (box) box.classList.add('hidden');
        if (wrap) wrap.innerHTML = '';
      }
    }

    // 把当前面板状态写入缓存（在状态与 PDF 都就绪后调用）。pdfInfo 为 { fname, isOver, node } 或 null。
    function cacheReportPanel(stateToken, pdfInfo) {
      if (!selectedProject || !selectedProject.id) return;
      const el = reportEls();
      const badge = document.getElementById('report-state-badge');
      const detail = document.getElementById('report-state-detail');
      const visible = el.state && !el.state.classList.contains('hidden');
      let phase = '', text = '';
      if (badge) {
        const cls = String(badge.className || '').replace('report-state-badge', '').trim();
        phase = cls.split(/\s+/)[0] || '';
        text = badge.textContent || '';
      }
      // PDF 只要当前已内嵌（reportPdfObjUrl 非空）就缓存——不依赖 pdfInfo 是否由自动浏览链路写入
      const hasPdf = !!reportPdfObjUrl;
      reportPanelCache = {
        key: String(selectedProject.id) + '|' + (stateToken || 'none'),
        stateToken: stateToken || 'none',
        badgePhase: visible ? phase : '',
        badgeText: text,
        detail: detail ? detail.textContent : '',
        statusText: el.status ? el.status.textContent : '',
        // ★ 缓存**当前**的 reportPdfObjUrl —— 切回时复用，避免重新下载整份 PDF（最大耗时点）
        pdf: hasPdf
          ? { objUrl: reportPdfObjUrl, fname: (pdfInfo && pdfInfo.fname) || '', isOver: !!(pdfInfo && pdfInfo.isOver), node: pdfInfo ? pdfInfo.node : undefined }
          : null,
      };
    }

    // 进入「数据报告处理」页签 / 切换项目时，默认自动加载并内嵌显示**当下流程中**的检测结果报告单。
    //   与「审核成功后」「回退成功后」共用同一个 `showFinalReportPdf({auto:true})`——只是自动浏览不弹窗、不滚动、失败静默。
    //   ⚠ 切换项目会通过 reportPdfSeq 丢弃过期响应。
    let reportPdfSeq = 0;
    async function autoShowCurrentReport() {
      if (!selectedProject || !selectedProject.id) return;
      const seq = ++reportPdfSeq;
      reportPdfPending = true;
      reportPdfCacheInfo = null;
      try {
        const r = await showFinalReportPdf({ auto: true });
        if (seq !== reportPdfSeq) return; // 期间已切换项目 → 该 PDF 已由新项目覆盖/释放
        // 自动浏览成功 → 记录信息，等待与状态查询一起写入缓存
        reportPdfCacheInfo = { fname: (r && r.fname) || '', isOver: !!(r && r.isOver) };
        return r;
      } catch (e) {
        if (seq !== reportPdfSeq) return;
        // 静默失败：状态栏给出可操作提示（如「尚无报告文件」）
        const el = reportEls();
        if (el.status) el.status.textContent = '未能加载当下流程中的数据报告：' + ((e && e.message) ? e.message : String(e));
        try { if (typeof hideReportPdf === 'function') hideReportPdf(); } catch (e2) {}
      } finally {
        if (seq === reportPdfSeq) {
          reportPdfPending = false;
          maybeCacheReportPanel();   // ★ 两条链都结束 → 尝试写缓存
        }
      }
    }

    function esc(s) {
      return Util.escHtml(s);
    }

    // 平台原生是三次连续确认；照搬，避免误触
    function tripleConfirm(what) {
      const tail = '\n\n' + reportProjectLabel();
      const qs = [
        '确认' + what + '所选中项目的审批状态？',
        '慎重考虑，确认' + what + '所选中项目的审批状态？',
        '再次强调，确定' + what + '所选中项目的审批状态？',
      ];
      for (let i = 0; i < qs.length; i++) {
        if (!window.confirm(qs[i] + tail)) return false;
      }
      return true;
    }

    // 回退 = 两步顺序重置（与人工操作一致）：
    //   ① 检测结果报告单（jcbg）→ ② 检测底稿（jcdg）
    // 两者走同一接口，仅 projectProcess 不同。当前用 App 当前会话（Plan A）执行；
    // 若日后要求以 btc-gxz / btc-wcy 各自的账号身份执行，则需接入账号池（见交接文档）。
    // ★ 回退只对**已完结**项目有效。平台「流程审批 → 检测结果报告单」的列表接口
    //     GET /api/projectJcbgLastAffix?isDel=1
    //   才是回退对象的权威来源（平台里「已完结」的定义就是 isDel==='1'）。
    //   ⚠ /api/reportData/findList 会忽略 isDel（实测 1/0/不带返回完全相同），
    //     用它搜出来的项目直接拿去回退会「找不到」，所以回退前必须先在该接口里定位。
    async function findFinishedProject(code) {
      const kw = String(code || '').trim();
      if (!kw) return null;
      const r = await apiRequest('GET',
        '/api/projectJcbgLastAffix?pageNumber=1&pageSize=50&isDel=1&code=' + encodeURIComponent(kw),
        { token: token, orgId: orgId, timeout: 60000 });
      if (r.status !== 200) throw new Error('查询「已完结」项目失败(HTTP ' + r.status + ')');
      const rows = (r.data && r.data.body && r.data.body.records) || [];
      const up = kw.toUpperCase();
      for (let i = 0; i < rows.length; i++) {
        if (String(rows[i].code || '').trim().toUpperCase() === up) return rows[i];
      }
      return null;
    }

    async function resetReportApproval() {
      const el = reportEls();
      if (!token) { showLogin(); return; }
      if (!selectedProject || !selectedProject.id) {
        alert('尚未在「数据上传」中选择项目。\n请先到「数据上传」查询并选中项目，再进行数据报告处理。');
        return;
      }
      const btn = document.getElementById('report-reset');
      const old = btn ? btn.textContent : '';
      // ① 先在「已完结」项目中定位（回退只能作用于已完结项目）
      let target = null;
      if (btn) { btn.disabled = true; btn.textContent = '定位中…'; }
      if (el.status) el.status.textContent = '正在「已完结」项目中定位 ' + (selectedProject.code || '') + '…';
      try {
        target = await findFinishedProject(selectedProject.code);
      } catch (e) {
        if (btn) { btn.disabled = false; btn.textContent = old || '数据报告回退'; }
        if (el.status) el.status.textContent = '❌ 定位失败：' + e.message;
        alert('数据报告回退失败：' + e.message);
        return;
      }
      if (btn) { btn.disabled = false; btn.textContent = old || '数据报告回退'; }
      if (!target) {
        if (el.status) el.status.textContent = '❌ 「已完结」项目中未找到 ' + (selectedProject.code || '');
        alert('在「已完结」项目中未找到该项目：' + (selectedProject.code || '')
          + '\n\n回退（重置审批状态）只对**已完结**项目有效，请确认该项目是否已审批完结。');
        return;
      }
      if (el.status) el.status.textContent = '✅ 已在「已完结」项目中定位到 ' + (target.code || '') + '，等待确认…';
      // ② 三重确认（照搬平台）
      if (!tripleConfirm('重置')) return;
      if (btn) { btn.disabled = true; }
      try {
        let done = 0;
        let changedAny = false;
        for (let i = 0; i < RESET_STEPS.length; i++) {
          const step = RESET_STEPS[i];
          if (btn) btn.textContent = '回退中…（' + (i + 1) + '/' + RESET_STEPS.length + ' ' + step.account + '·' + step.label + '）';
          if (el.status) el.status.textContent = '正在以 ' + step.account + ' 身份重置「' + step.label + '」（' + (i + 1) + '/' + RESET_STEPS.length + '）…';
          const tk = await fetchBridgeToken(step.account);
          await ensurePoolTokenFresh(step.account, tk); // 调用前校验：失效即经桥重登刷新
          const r = await poolApiRequest('POST', '/api/projectXccyLastAffix/resetProjectApprovalProcess', {
            token: tk.token, orgId: tk.orgId, timeout: 60000,
            json: { belongProject: target.id, projectProcess: step.process },
          });
          // ⚠ 平台返回形如 {"code":"200","message":"重置成功！","body":true|false}：
          //   body 只表示「本次是否真的发生了状态变更」——已处于初始状态时返回 false，
          //   但**仍是成功**（幂等，实测重复调用/已重置过的项目都返回 body:false）。
          //   所以成败只认 code==='200'，不能拿 body 当成功标志（否则会误报「重置失败：重置成功！」）。
          if (!(r.status === 200 && r.data && String(r.data.code) === '200')) {
            throw new Error('「' + step.label + '」以 ' + step.account + ' 身份重置失败：' + ((r.data && (r.data.message || r.data.msg)) || ('HTTP ' + r.status + ' ' + String(r.text || '').slice(0, 120))));
          }
          const changed = !!(r.data && r.data.body);
          if (changed) changedAny = true;
          done++;
          if (el.status) el.status.textContent = '✅ 已以 ' + step.account + ' 身份重置「' + step.label + '」'
            + (changed ? '' : '（无需变更，原本已是初始状态）')
            + '（' + done + '/' + RESET_STEPS.length + '）';
        }
        if (el.status) el.status.textContent = '✅ 两步回退全部完成 · ' + reportProjectLabel()
          + (changedAny ? '' : '（两步都无需变更，原本已是初始状态）');
        alert('数据报告回退成功（已依次重置：' + RESET_STEPS.map(function (s) { return s.label; }).join(' → ') + '）。'
          + (changedAny ? '' : '\n注：两步都无需变更，该项目原本已是初始状态。') + '\n' + reportProjectLabel());
        // ★ 回退成功后：与「审核成功后」「进页签自动浏览」共用同一套「当下流程中的数据报告」逻辑
        invalidateReportCache();   // 审批状态已变 → 作废面板缓存，强制重新拉取
        if (el.status) el.status.textContent = '✅ 两步回退全部完成，正在加载当前流程中的数据报告…';
        try {
          await showFinalReportPdf({});
        } catch (e) {
          if (el.status) el.status.textContent = '✅ 两步回退全部完成，但数据报告加载失败：' + (e && e.message ? e.message : e);
        }
        await loadReportStatus(); // 回退改变了状态 → 同步刷新「数据报告状态」
      } catch (e) {
        if (el.status) el.status.textContent = '❌ 回退失败：' + e.message;
        alert('数据报告回退失败：' + e.message);
      } finally {
        if (btn) { btn.disabled = false; btn.textContent = old || '数据报告回退'; }
      }
    }

    // 审核 = 跨账号审批链（见 AUDIT_STEPS 注释）。每步都用**各自账号**的令牌（经账号桥获取）。
    // ★ 支持**断点续审 + 一路审到底**：
    //   · 执行前先读两个流程的当下节点快照，规划出「未完成的步骤」（跳过已完成项）；
    //   · 但**审核是动态推进的**——每完成一步节点会前移（jcbg 2→6→完结），
    //     因此循环内**每步后重新读取节点、重新求下一步**，一直执行到链路终点（签发完结），
    //     而**不是**执行完初始规划就收尾。见 nextAuditStep 注释。
    async function auditReport() {
      const el = reportEls();
      if (!token) { showLogin(); return; }
      if (!selectedProject || !selectedProject.id) {
        alert('尚未在「数据上传」中选择项目。\n请先到「数据上传」查询并选中项目，再进行数据报告处理。');
        return;
      }
      const btn = document.getElementById('report-audit');
      const old = btn ? btn.textContent : '';
      if (btn) { btn.disabled = true; }
      const msgs = [];
      // 测试可通过 window.__AUDIT_POLL_OPTS__ 覆盖轮询参数（避免 mock 固定响应导致超时）
      const pollOpts = (typeof window !== 'undefined' && window.__AUDIT_POLL_OPTS__) || {};
      try {
        // ── 第一步：读「两个流程的当下节点快照」→ 规划本次需要执行的步骤（断点续审的核心）──
        if (el.status) el.status.textContent = '正在读取当前审批节点…';
        // ★ 两条流程**并行**读（原来串行，白白多一个 RTT）；读取失败必须中止，不能当成「无记录」
        const firstPair = await Promise.all([
          readNodeState(token, orgId, auditPaths('jcdg').affix, selectedProject.id),
          readNodeState(token, orgId, auditPaths('jcbg').affix, selectedProject.id),
        ]);
        if (firstPair[0].failed || firstPair[1].failed) {
          const which = [firstPair[0].failed ? '检测底稿' : '', firstPair[1].failed ? '检测结果报告单' : '']
            .filter(Boolean).join(' / ');
          const why = (firstPair[0].failed && firstPair[0].message)
            || (firstPair[1].failed && firstPair[1].message) || '未知错误';
          if (el.status) el.status.textContent = '❌ 读取审批节点失败：' + why;
          alert('读取「' + which + '」当前审批节点失败：' + why
            + '\n\n为避免误判平台状态，本次未执行任何审批操作。请稍后重试。');
          return;
        }
        const jcdgNode = firstPair[0].node;
        const jcbgNode = firstPair[1].node;
        const plan = planAuditSteps({ jcdg: jcdgNode, jcbg: jcbgNode });
        if (plan.unknowns.length) {
          const msg = '该项目当前审批节点不在本工具可审范围内：\n\n· ' + plan.unknowns.join('\n· ')
            + '\n\n为避免污染平台状态，已中止（未执行任何审批操作）。';
          if (el.status) el.status.textContent = '⏹ ' + plan.unknowns.join('；');
          alert(msg);
          return;
        }
        if (!plan.indices.length && !plan.needSubmitJcbg) {
          const msg = plan.completed
            ? '该项目的「检测底稿 + 检测报告单」审批链**已全部完结**，无需再审。\n\n' + reportProjectLabel()
            : '该项目在当前节点已无可执行的审核步骤。\n\n' + reportProjectLabel();
          if (el.status) el.status.textContent = plan.completed
            ? '⏹ 审批链已全部完结' : '⏹ 已无可执行的审核步骤';
          alert(msg);
          return;
        }
        // 断点续审提示（初始规划，仅用于确认框说明；实际执行是动态续跑到底）
        // ⚠ 需补提交报告单时，链路上那 4 步都还没做，不能算「已跳过」。
        const skipped = plan.needSubmitJcbg ? []
          : AUDIT_STEPS.map(function (s, i) { return i; })
            .filter(function (i) { return plan.indices.indexOf(i) === -1; });
        const heading = plan.needSubmitJcbg
          ? '（检测底稿已完结，但检测报告单没有待审记录 → 先补提交报告单，再一路审到签发完结）'
          : (skipped.length
            ? '（断点续审：从当前节点续跑，跳过 ' + skipped.length + ' 步已完成项）'
            : '');
        const repairLine = plan.needSubmitJcbg
          ? '0. ' + AUDIT_STEPS[0].account + ' → 补提交检测报告单（当前无待审记录）\n'
          : '';
        const chain = plan.indices.map(function (i, n) {
          return (n + 1) + '. ' + AUDIT_STEPS[i].account + ' → ' + AUDIT_STEPS[i].label;
        }).join('\n');
        const skipNote = skipped.length
          ? '\n\n当前节点：' + describeNodes(jcdgNode, jcbgNode)
            + '\n\n已跳过（视为已完成/无需重复）：\n'
            + skipped.map(function (i) { return '· ' + STEP_LABELS[i]; }).join('\n')
          : '';
        if (!window.confirm('确认对以下项目执行「数据报告审核」？\n\n' + reportProjectLabel()
          + (heading ? '\n' + heading : '')
          + skipNote
          + '\n\n本次将从当前节点**一路审到链路终点（签发完结）**。'
          + (plan.needSubmitJcbg
            ? '先补提交检测报告单，再动态推进：\n' + repairLine + chain
            : '预计执行 ' + plan.indices.length + ' 步起（各步用各自账号，动态推进）：\n' + chain))) return;

        let done = 0;
        let abortMsg = ''; // 非空 = 中途判定无法继续（区别于「正常跑到终点」）
        const ranAccounts = []; // 本次实际执行过的账号（按顺序）
        // ★★ 动态循环：每步开始前重新读节点 → nextAuditStep 求「下一步」→ 执行 → 再循环。
        //   maxSteps 兜底（=AUDIT_STEPS 长度 + 余量），防止极端情况下死循环。
        const maxSteps = AUDIT_STEPS.length + 4;
        let guard = 0;
        let jcbgRepaired = false; // 「补提交检测报告单」只做一次，避免死循环
        while (true) {
          if (++guard > maxSteps) { msgs.push('⚠ 已达最大步数上限(' + maxSteps + ')，中止以防死循环'); break; }
          // 重新读当下节点（上一步通过后节点会前移；这也是「续跑到底」的关键）
          let curJcdg, curJcbg;
          // ★ 两条流程**并行**读（原来串行）；这两条快照同时供下面「本步节点」直接复用。
          //   ⚠ 读取失败必须中止：若当成「无待审记录」，会直接推出「审批链已全部完结」。
          const pair = await Promise.all([
            readNodeState(token, orgId, auditPaths('jcdg').affix, selectedProject.id),
            readNodeState(token, orgId, auditPaths('jcbg').affix, selectedProject.id),
          ]);
          if (pair[0].failed || pair[1].failed) {
            const which = [pair[0].failed ? '检测底稿' : '', pair[1].failed ? '检测结果报告单' : '']
              .filter(Boolean).join(' / ');
            const why = (pair[0].failed && pair[0].message)
              || (pair[1].failed && pair[1].message) || '未知错误';
            throw new Error('读取「' + which + '」当前审批节点失败（' + why
              + '）；已中止，未继续审批，以免误判平台状态。请稍后重试。');
          }
          curJcdg = pair[0].node; curJcbg = pair[1].node;
          const nx = nextAuditStep({ jcdg: curJcdg, jcbg: curJcbg });
          if (nx.unknowns.length) {
            abortMsg = nx.unknowns.join('；');
            msgs.push('⏹ 中止：' + abortMsg);
            if (el.status) el.status.textContent = '⏹ ' + abortMsg;
            break;
          }
          if (nx.needSubmitJcbg) {
            // ★ 底稿已完结、报告单无待审记录（多半从未提交）→ **补提交报告单**后继续跑完
            if (jcbgRepaired) {
              abortMsg = '已补提交过检测报告单，但仍无待审记录，无法继续';
              msgs.push('⏹ 中止：' + abortMsg);
              if (el.status) el.status.textContent = '⏹ ' + abortMsg;
              break;
            }
            jcbgRepaired = true;
            const sub = AUDIT_STEPS[0]; // 提交审复核用 btc-jyg（与第①步同账号）
            if (btn) btn.textContent = '补提交报告单… ' + sub.account;
            if (el.status) el.status.textContent = '检测报告单无待审记录，正在以 ' + sub.account + ' 补提交…';
            const tks = await fetchBridgeToken(sub.account);
            await ensurePoolTokenFresh(sub.account, tks);
            const rr = await poolApiRequest('POST', '/api/reportFile/submitJcbgReview', {
              token: tks.token, orgId: tks.orgId, timeout: 60000,
              json: { belongProject: selectedProject.id },
            });
            if (!(rr.status === 200 && rr.data && String(rr.data.code) === '200')) {
              throw new Error('补提交「检测报告单」（' + sub.account + '）失败：'
                + ((rr.data && (rr.data.message || rr.data.msg)) || ('HTTP ' + rr.status + ' ' + String(rr.text || '').slice(0, 120))));
            }
            const rb = rr.data && rr.data.body;
            msgs.push('补提交检测报告单：' + ((rb && (rb.msg || rb.message)) || '已提交'));
            if (el.status) el.status.textContent = '已补提交检测报告单，等待待审记录出现…';
            const ready = await pollNodeReady(tks.token, tks.orgId, auditPaths('jcbg').affix, selectedProject.id, pollOpts);
            if (!ready) {
              throw new Error('补提交后 ' + Math.round((pollOpts.maxWaitMs || 60000) / 1000)
                + 's 内检测报告单仍无待审记录，已中止');
            }
            done++;
            ranAccounts.push(sub.account);
            continue;
          }
          if (nx.index === null) {
            // 已无可执行步骤：正常收尾
            msgs.push('链路到达终点（' + (nx.why || '无可执行步骤') + '）');
            break;
          }
          const i = nx.index;
          const step = AUDIT_STEPS[i];
          const tag = '（第 ' + (done + 1) + ' 步）';
          if (btn) btn.textContent = '审核中…' + tag + ' ' + step.account;
          if (el.status) el.status.textContent = '正在以 ' + step.account + ' 执行「' + step.label + '」' + tag + '…';
          const tk = await fetchBridgeToken(step.account);
          await ensurePoolTokenFresh(step.account, tk); // 调用前校验：失效即经桥重登刷新

          if (step.kind === 'submitJcdg') {
            // ★ 平台里「检测底稿」与「检测报告单」是**两个独立页面、两个独立的提交接口**：
            //   检测底稿  → /api/projectJcdgLastAffixNew/submitJcdgReview
            //   检测报告单 → /api/reportFile/submitJcbgReview   ← 必须**另外**提交！
            //   （实测 BTC26-SZJC1009：只调 submitJcdgReview 时 jcdg 有记录、jcbg 完全无记录，
            //     导致后面对 jcbg 的审批报「该流程已无待审的检测记录」。）
            //   因此本步先用当前账号（btc-jyg）提交检测底稿，再提交检测报告单，然后等两条流程都就绪。
            if (el.status) el.status.textContent = '正在提交「检测底稿」…' + tag + '（' + step.account + '）';
            const r = await poolApiRequest('POST', '/api/projectJcdgLastAffixNew/submitJcdgReview', {
              token: tk.token, orgId: tk.orgId, timeout: 60000,
              json: { belongProject: selectedProject.id },
            });
            if (!(r.status === 200 && r.data && String(r.data.code) === '200')) {
              throw new Error('「检测底稿提交审复核」（' + step.account + '）失败：'
                + ((r.data && (r.data.message || r.data.msg)) || ('HTTP ' + r.status + ' ' + String(r.text || '').slice(0, 120))));
            }
            const b = r.data && r.data.body;
            const submitMsg = b && (b.msg || b.message);
            if (submitMsg) msgs.push('检测底稿提交审复核：' + (b.prjCode || '') + ':' + submitMsg);

            // ★ 再提交「检测报告单」（独立接口）——这是原实现遗漏的关键一步
            if (el.status) el.status.textContent = '正在提交「检测报告单」…' + tag + '（' + step.account + '）';
            const r2 = await poolApiRequest('POST', '/api/reportFile/submitJcbgReview', {
              token: tk.token, orgId: tk.orgId, timeout: 60000,
              json: { belongProject: selectedProject.id },
            });
            if (!(r2.status === 200 && r2.data && String(r2.data.code) === '200')) {
              throw new Error('「检测报告单提交审复核」（' + step.account + '）失败：'
                + ((r2.data && (r2.data.message || r2.data.msg)) || ('HTTP ' + r2.status + ' ' + String(r2.text || '').slice(0, 120))));
            }
            const b2 = r2.data && r2.data.body;
            const submitMsg2 = b2 && (b2.msg || b2.message);
            if (submitMsg2) msgs.push('检测报告单提交审复核：' + (b2.prjCode || '') + ':' + submitMsg2);

            // ★ 等待 jcdg 与 jcbg **两条**待审记录都就绪（最多 60s，每 2s 一次）
            if (el.status) el.status.textContent = '等待「检测底稿 + 检测报告单」都就绪…' + tag + '（' + step.account + '）';
            const poll = await pollBothNodesReady(tk.token, tk.orgId, selectedProject.id, pollOpts);
            if (!poll.ready) {
              // ⚠ 两条没齐就**中止**，绝不能再往下走（否则后面 jcbg 审批必然报错，且会污染平台状态）
              const miss = [];
              if (!poll.jcdg) miss.push('检测底稿');
              if (!poll.jcbg) miss.push('检测报告单');
              throw new Error('等待 ' + (poll.waited / 1000).toFixed(1) + 's 后，'
                + miss.join(' / ') + ' 仍无待审记录；已中止（不继续后续审批，避免污染平台状态）');
            }
            msgs.push('提交审复核：✅ 检测底稿（节点 ' + nodeName(poll.jcdg.type)
              + '）+ 检测报告单（节点 ' + nodeName(poll.jcbg.type) + '）均已就绪（等待 '
              + (poll.waited / 1000).toFixed(1) + 's）');
          } else {
            const ap = auditPaths(step.kind);
            // ★ 直接复用本轮循环刚读到的节点快照（nextAuditStep 正是据此判定「该做这一步」的），
            //   不再重复 getAffix —— 前置节点读取是审核链里最耗时的一环。
            //   只有快照里该流程确实没有节点时，才回落到短轮询兜底（平台流转延迟/读取抖动）。
            let node = step.kind === 'auditJcdg' ? curJcdg : curJcbg;
            if (!node) {
              const waitMs = pollOpts.nodeWaitMs || 20000;
              const t0 = Date.now();
              let round = 0;
              while (!node && Date.now() - t0 < waitMs) {
                await new Promise(function (res) { setTimeout(res, pollDelay(round++, pollOpts)); });
                try { node = await readAffixNode(tk.token, tk.orgId, ap.affix, selectedProject.id); } catch (e) { node = null; }
              }
            }
            if (!node) {
              // nextAuditStep 刚判定该流程有待审节点，这里却读不到 → 状态异常（流转延迟/已被他人处理）。
              // 不静默跳过：报错中止并给出可读原因（下次点击会重新读节点、自然续跑）。
              const hint = step.kind === 'auditJcbg'
                ? '——检测报告单可能尚未提交成功，或已被他人在平台上处理'
                : '——上一步可能尚未生效，或该项目本就不在该节点';
              throw new Error('「' + step.label + '」（' + step.account + '）：该流程已无待审的检测记录' + hint);
            }
            const oldType = node.type;
            // ★ 前置校验：平台不允许**同一人连续审批两个相邻节点**（报「审批流程中的人员不可重复！」）。
            //   节点的 createBy/updateBy = 上一步经手人；若本次审批账号与上一步是同一人 → 提前中止，不发请求。
            const prevBy = node.updateBy !== undefined && node.updateBy !== null ? node.updateBy : node.createBy;
            const acctId = tk.userId !== undefined && tk.userId !== null ? String(tk.userId) : null;
            if (acctId && prevBy !== undefined && prevBy !== null && String(prevBy) === acctId) {
              throw new Error('「' + step.label + '」（' + step.account + '）无法执行：平台不允许同一人连续审批两个相邻节点。'
                + '\n上一步（' + nodeName(oldType) + '）已由该账号（用户ID ' + acctId + '）经手，'
                + '本步需换其他人审批。\n当前节点：' + describeNodes(curJcdg, curJcbg));
            }
            const r = await poolApiRequest('POST', ap.audit, {
              token: tk.token, orgId: tk.orgId, timeout: 60000,
              // ⚠ 单个对象（不是数组）；jcbg 的 type 必须为数字
              json: {
                belongProject: selectedProject.id,
                projectName: selectedProject.code || '',
                suggestions: '通过',
                type: /^\d+$/.test(String(oldType)) ? Number(oldType) : oldType,
                isPass: '1',
                auditDate: auditDateStr(),
              },
            });
            if (!(r.status === 200 && r.data && String(r.data.code) === '200')) {
              const raw = (r.data && (r.data.message || r.data.msg)) || ('HTTP ' + r.status + ' ' + String(r.text || '').slice(0, 120));
              // ★ 翻译平台「审批流程中的人员不可重复！」：同一人不能连审两个相邻节点。
              //   本工具链路的账号本应各不相同（jyg→wcy→zhj→gxz），正常不该出现；
              //   若出现，多半是该节点已被 profile 里某账号审过、或账号池令牌错配 → 给出可操作提示。
              if (/人员不可重复|不可重复/.test(String(raw))) {
                throw new Error('「' + step.label + '」（' + step.account + '）失败：' + raw
                  + '\n\n原因：平台不允许同一人连续审批两个相邻节点。'
                  + '该节点可能已由 ' + step.account + '（或同一人）审过，请刷新后用「数据报告审核」重新读取当前节点。'
                  + '当前节点：' + describeNodes(curJcdg, curJcbg));
              }
              throw new Error('「' + step.label + '」（' + step.account + '）失败：' + raw);
            }
            msgs.push(step.label + '：' + ((r.data && r.data.message) || '审批成功') + '（节点 ' + nodeName(oldType) + '）');
            // 轮询等待节点推进
            if (el.status) el.status.textContent = '等待状态流转中…' + tag + '（' + step.account + ' · ' + step.label + '）';
            const poll = await pollNodeChange(tk.token, tk.orgId, ap.affix, selectedProject.id, oldType, pollOpts);
            if (!poll.changed) {
              msgs.push(step.label + '：状态流转等待超时（' + (poll.waited / 1000).toFixed(1) + 's），继续下一步');
            } else {
              msgs.push(step.label + '：状态流转完成（等待 ' + (poll.waited / 1000).toFixed(1) + 's'
                + (poll.gone ? '，该流程已完结' : (poll.newType !== undefined ? '，推进到 ' + nodeName(poll.newType) : '')) + '）');
            }
          }
          done++;
          ranAccounts.push(step.account);
          if (el.status) el.status.textContent = '✅ 已执行 ' + done + ' 步（' + step.account + ' · ' + step.label + '）';
        }
        const tail = msgs.length ? ' · ' + msgs.join('；') : '';
        const acctChain = ranAccounts.join(' → ');
        // ★ 中途因「无法判定后续节点」而停止 → **不能**说「审核完成」
        if (abortMsg) {
          if (el.status) el.status.textContent = '⏹ 已中止（本次 ' + done + ' 步）：' + abortMsg;
          alert('数据报告审核已中止。\n' + reportProjectLabel()
            + '\n本次执行 ' + done + ' 步' + (acctChain ? '：' + acctChain : '') + '。'
            + '\n\n' + abortMsg);
          invalidateReportCache();
          try { await loadReportStatus(); } catch (e) {}
          return;
        }
        if (el.status) el.status.textContent = '✅ 审核链全部完成（本次 ' + done + ' 步'
          + (skipped.length ? '，跳过已完成 ' + skipped.length + ' 步' : '') + '）· ' + reportProjectLabel() + tail;
        // ★ 弹窗只保留**极简摘要**（用户反馈：弹窗文字太多）——项目 + 步数 + 账号链，明细走状态栏。
        alert('数据报告审核完成。\n' + reportProjectLabel()
          + '\n本次执行 ' + done + ' 步' + (acctChain ? '：' + acctChain : '') + '。');
        // ★ 审核成功后：与「回退成功后」「进页签自动浏览」共用同一套「当下流程中的数据报告」逻辑
        invalidateReportCache();   // 审批状态已变 → 作废面板缓存，强制重新拉取
        if (el.status) el.status.textContent = '✅ 审核链全部完成，正在加载当前流程中的数据报告…';
        try {
          await showFinalReportPdf({});
        } catch (e) {
          if (el.status) el.status.textContent = '✅ 审核链全部完成，但数据报告加载失败：' + (e && e.message ? e.message : e);
        }
        await loadReportStatus(); // 审核完结改变了状态 → 同步刷新「数据报告状态」
      } catch (e) {
        if (el.status) el.status.textContent = '❌ 审核失败（已完成 ' + msgs.length + ' 步）：' + e.message;
        alert('数据报告审核失败：' + e.message);
      } finally {
        if (btn) { btn.disabled = false; btn.textContent = old || '数据报告审核'; }
      }
    }

    // ★ 解析「当下流程中的数据报告」PDF 来源路径（**只拿路径、不下载**）。返回 { path, isOver, node, src } 或抛错。
    //
    //   设计要点：**不区分「回退后 / 审核后 / 自动浏览」**，一律按项目**当下的审批状态**取报告 —— 依次尝试、取到就停：
    //   ① 审批链记录（**当下流程中的那份**，首选）：
    //      GET /api/projectJcbgLastAffix/getAffix?belongProject= → body[0].jcbgAffix
    //      ★ 这就是平台「流程审批 → 检测结果报告单 → 附件」里那份数据报告：
    //        未完结时形如 …_检测报告_<用户id>.pdf（如项目 1009 已提交审复核 → isOver=0、type="2"）；
    //        审批链走到「签发」完结后，路径变为 …/review/xxx_检测报告_签发.pdf。
    //   ② 检测结果报告单本体（**回退后**审批链被清空时的兜底）：
    //      GET /api/jcbgReport/jcbgYl?projectId=&organizationId= → body（相对路径，直接拼 API_BASE）
    //      ★ 平台「检测结果报告单」页（路由 /recordPageNew/pageReport 的 jcbg 标签）用的就是这个接口。
    //      ⚠ 回退会**清空 jcbg 审批链记录**（getAffix 变空），但报告文件本体仍在 → 只能从 ② 取。
    //   用当前会话令牌直连平台取路径；展示方式：**当前页面下方内嵌 iframe**（不新开标签页）。
    let reportPdfObjUrl = null;
    // opts.keepBlob=true → 只隐藏 iframe、**不释放** blob URL，供面板缓存复用（页签来回切换秒开）。
    //   默认（无参）= 完全释放（切项目 / 改状态 / 主动关闭时用）。
    function hideReportPdf(opts) {
      const keepBlob = !!(opts && opts.keepBlob);
      const box = document.getElementById('report-pdf-box');
      const wrap = document.getElementById('report-pdf-frame-wrap');
      if (wrap) wrap.innerHTML = '';
      if (box) box.classList.add('hidden');
      try { document.body.classList.remove('pdf-open'); } catch (e) {}
      if (reportPdfObjUrl) {
        // 缓存仍持有该 blob（cachedObjUrl 指向它）→ 不释放，仅断开当前引用；复原时再接管
        const owned = reportPanelCache && reportPanelCache.pdf && reportPanelCache.pdf.objUrl === reportPdfObjUrl;
        if (!keepBlob && !owned) { try { URL.revokeObjectURL(reportPdfObjUrl); } catch (e) {} }
        reportPdfObjUrl = null;
      }
    }
    // 解析报告来源路径（不下载，只拿路径）。返回 { path, isOver, node, src } 或抛错。
    async function resolveFinalReportPath(opts) {
      opts = opts || {};
      const pid = String(selectedProject.id);
      const el = reportEls();
      // ① 审批链记录（当下流程中的那份：未完结=<用户id> 版；完结=review/签发 版）
      try {
        const r = await apiRequest('GET',
          auditPaths('jcbg').affix + '?belongProject=' + encodeURIComponent(pid),
          { token: token, orgId: orgId, timeout: 60000 });
        if (r.status === 200) {
          const rows = (r.data && r.data.body) || [];
          if (rows.length && rows[0] && rows[0].jcbgAffix) {
            return { path: String(rows[0].jcbgAffix), isOver: String(rows[0].isOver) === '1',
              node: rows[0].type, src: 'affix' };
          }
        }
      } catch (e) { /* 落到 ② */ }
      // ② 检测结果报告单本体（回退后审批链已清空 → 走这里）。
      //   ⚠ 该接口是「生成并返回路径」类，**可能非常慢**（实测个别项目 >8min 仍不返回）→ 给状态提示。
      if (el.status) el.status.textContent = '正在向平台请求数据报告（审批链已重置，改从报告单本体取，可能较慢）…';
      const r2 = await apiRequest('GET',
        '/api/jcbgReport/jcbgYl?projectId=' + encodeURIComponent(pid) + '&organizationId=' + encodeURIComponent(String(orgId)),
        { token: token, orgId: orgId, timeout: 300000 });
      if (r2.status !== 200) throw new Error('读取检测结果报告单失败(HTTP ' + r2.status + ')');
      const p = r2.data && (r2.data.body || r2.data.data);
      if (!p || typeof p !== 'string') throw new Error('该项目尚无「检测结果报告单」文件（可能未提交/未生成）');
      return { path: p, isOver: false, node: null, src: 'jcbgYl' };
    }
    // 取「检测结果报告单」PDF 并内嵌显示（与「数据报告审核」同一套界面/逻辑）。
    //   opts.quiet         ：失败/未完结时只写状态栏、不弹 alert（供审核/回退成功后、以及进页签自动浏览时调用）
    //   opts.context       ：状态栏文案里的场景名（默认「审批完结」），如「回退后」「当前状态」
    //   opts.silentPending ：未完结时**不弹**「尚未审批完结」提示（进页签自动浏览当前状态版时用）
    //   返回 { ok, fname, isOver } 或抛错。
    // ★★★ 统一的「当下流程中的数据报告」显示入口（2026-10-15 整合）
    //
    //   三个触发场景**共用这一套逻辑**，不再各写一份：
    //     ① 进「数据报告处理」页签 / 切换项目 → 自动浏览（auto=true）
    //     ② 「数据报告审核」4 步链成功后     → 自动显示（auto=false，滚动定位）
    //     ③ 「数据报告回退」两步成功后       → 自动显示（auto=false，滚动定位）
    //
    //   取源由 `resolveFinalReportPath()` 按**当下状态**决定（依次尝试、取到就停）：
    //     ① 审批记录附件 `getAffix.jcbgAffix`（**当下流程中的那份**：未完结 = …_检测报告_<用户id>.pdf；
    //        签发完结 = …/review/…_检测报告_签发.pdf）—— 与平台「流程审批 → 检测结果报告单 → 附件」一致；
    //     ② 回退后审批链被清空 → 落回「检测结果报告单本体」`jcbgReport/jcbgYl`。
    //
    //   opts.auto  ：true=进页签自动浏览（不弹窗、不滚动、失败静默）；false=用户主动操作后（滚动定位）
    //   返回 { ok, fname, isOver } 或抛错。
    async function showFinalReportPdf(opts) {
      opts = opts || {};
      const isAuto = !!opts.auto;
      // 主动操作（非自动浏览）→ 递增序号，取消可能仍在挂起的「自动浏览」请求，避免互相覆盖
      if (!isAuto) { try { reportPdfSeq++; } catch (e) {} }
      const el = reportEls();
      const resolved = await resolveFinalReportPath(opts);
      const isOver = resolved.isOver;
      const type = resolved.node;
      // jcbgAffix 形如 /src/ohit/jcbgReport/.../xxx_检测报告_签发.pdf —— 逐段编码以兼容中文路径
      const url = API_BASE + String(resolved.path).split('/').map(encodeURIComponent).join('/');
      const fname = String(resolved.path).split('/').pop() || '检测结果报告单.pdf';
      // 未完结：仅**用户主动点击**（理论上不会触发，因两个按钮都会改变状态）时才提示；
      // 自动浏览与流程执行后一律静默——显示的本来就是「当下流程中」的版本。
      if (!isOver && !isAuto && opts.warnPending) {
        alert('该项目「检测结果报告单」尚未审批完结' + (type ? '（当前节点：' + nodeName(type) + '）' : '') + '。\n'
          + '仍将显示当下流程中的报告文件，如需「审批完结」版请先跑完审核链（或等待其余节点审批）。\n\n' + fname);
      }
      // 平台下载接口用 Authorization 头鉴权，iframe 直开带不了头 → 先 fetch 成 blob，再内嵌 blob URL
      if (el.status) el.status.textContent = '正在加载报告 PDF…' + fname;
      const resp = await fetch(url, { headers: token ? { Authorization: token } : {} });
      if (!resp.ok) throw new Error('下载报告失败(HTTP ' + resp.status + ')');
      const blob = await resp.blob();
      // 释放上一次的 blob URL（⚠ 若该 URL 仍被面板缓存持有，则不释放——所有权归缓存）
      if (reportPdfObjUrl) {
        const owned = reportPanelCache && reportPanelCache.pdf && reportPanelCache.pdf.objUrl === reportPdfObjUrl;
        if (!owned) { try { URL.revokeObjectURL(reportPdfObjUrl); } catch (e) {} }
      }
      reportPdfObjUrl = URL.createObjectURL(blob);
      // 内嵌显示在「数据报告处理」页签下方
      const box = document.getElementById('report-pdf-box');
      const wrap = document.getElementById('report-pdf-frame-wrap');
      if (wrap) {
        wrap.innerHTML = '';
        const fr = document.createElement('iframe');
        fr.src = reportPdfObjUrl;
        fr.title = '当下流程中的数据报告';
        wrap.appendChild(fr);
      }
      if (box) { box.classList.remove('hidden'); }
      // 让整个页面可纵向滚动，方便把 PDF 区域继续往上滚查看
      try { document.body.classList.add('pdf-open'); } catch (e) {}
      // ⚠ 只有**用户主动操作**（审核/回退成功后）才滚动定位；
      //    进页签/切项目的自动浏览**不滚动**，避免主页面滚动条突然下滑。
      if (!isAuto) {
        try {
          const tab = document.getElementById('tab-report');
          if (tab) tab.scrollIntoView({ behavior: 'smooth', block: 'start' });
        } catch (e) {}
      }
      if (el.status) {
        // 统一文案：都描述为「当下流程中的数据报告」；未完结时附当前节点说明
        const tail = (!isOver && type) ? '（当下流程：节点 ' + nodeName(type) + '，未完结）' : '';
        el.status.textContent = '✅ 已显示该项目当下流程中的数据报告：' + fname + tail;
      }
      return { ok: true, fname: fname, isOver: isOver };
    }

    // 进入「数据报告处理」页签的策略（与「噪声数据结果」页签一致）：
    //   · 未选中项目 → **不允许进入**（拦下 app.js 的页签切换，只提示用户先去「数据上传」选项目）
    //   · 已选项目   → 正常进入；面板内容由 refreshReportPanel() 负责（命中缓存则秒开）
    // 做法：捕获阶段拦截 app.js 的 click；未选项目才拦下；已选项目放行后于微任务里刷新面板
    //        （下一 tick 再刷新，确保 app.js 已把 #tab-report 切为 active，避免在隐藏面板上操作）。
    const reportTabBtn = document.querySelector('.tab[data-tab="report"]');
    if (reportTabBtn) {
      reportTabBtn.addEventListener('click', function (e) {
        if (!selectedProject || !selectedProject.id) {
          e.preventDefault();
          e.stopImmediatePropagation();   // 拦下 app.js 的页签切换 → 不进入该页签
          alert('尚未在「数据上传」中选择项目。\n请先到「数据上传」查询并选中项目，再进行数据报告处理。');
          return;
        }
        // 已选项目 → 放行，交给 app.js 切换；随后刷新面板（缓存命中则直接复原）
        setTimeout(function () { refreshReportPanel(); }, 0);
      }, true);
    }
    const reportResetBtn = document.getElementById('report-reset');
    if (reportResetBtn) reportResetBtn.addEventListener('click', resetReportApproval);
    const reportAuditBtn = document.getElementById('report-audit');
    if (reportAuditBtn) reportAuditBtn.addEventListener('click', auditReport);
    refreshReportPanel();

    // ---------- 非噪声岗位同步 ----------
    function splitLines(v) {
      const arr = String(v === null || v === undefined ? '' : v).split('\n').map(function (s) { return s.trim(); });
      while (arr.length > 1 && arr[arr.length - 1] === '') arr.pop();
      return arr.length ? arr : [''];
    }
    function pickLine(arr, i) { return arr.length === 1 ? arr[0] : (arr[i] === undefined ? '' : arr[i]); }
    function toNum(v) {
      const n = parseFloat(String(v === null || v === undefined ? '' : v).replace(/[^\d.\-]/g, ''));
      return isNaN(n) ? null : n;
    }

    // 取噪声表中 单元/岗位/点位 与 LEX 值，交给主程序按自动计算区 W/X/AL 匹配并置「是否噪声作业岗位」为「否」
    function syncNonNoiseJobs() {
      const disp = noiseCache && noiseCache.display;
      if (!disp || disp.length < 2) {
        alert('请先加载噪声数据结果（需在「数据上传」中选择项目），再执行同步。');
        return;
      }
      const head = disp[0];
      const findCol = function (name) {
        const want = noiseNorm(name);
        // 优先精确匹配：「测量点/对象」是「岗位/工种/测量点/对象」的子串，含匹配会取错列
        for (let i = 0; i < head.length; i++) if (noiseNorm(head[i]) === want) return i;
        for (let i = 0; i < head.length; i++) if (noiseNorm(head[i]).indexOf(want) >= 0) return i;
        return -1;
      };
      const cUnit = findCol('单元/工作场所');
      const cJob = findCol('岗位/工种/测量点/对象');
      const cSite = findCol('测量点/对象');
      const lexCols = [];
      for (let i = 0; i < head.length; i++) if (noiseNorm(head[i]).indexOf('LEX') >= 0) lexCols.push(i);
      if (cUnit < 0 || cJob < 0 || cSite < 0) { alert('噪声表缺少「单元/工作场所」「岗位/工种/测量点/对象」「测量点/对象」列，无法同步。'); return; }

      // 口径与参考表《非噪声岗位筛选》一致：先对空单元格向下填充，再按「车间+岗位」分组，
      // 组内取所有 LEX 列的最大值（非数字按 0）；每个点位都带上其所属岗位的组最大值。
      const list = [];
      const groupMax = new Map();
      let prevUnit = '';
      let prevJob = '';
      for (let r = 1; r < disp.length; r++) {
        const src = disp[r] || [];
        const unitRaw = String(src[cUnit] === null || src[cUnit] === undefined ? '' : src[cUnit]);
        const jobRaw = String(src[cJob] === null || src[cJob] === undefined ? '' : src[cJob]);
        if (unitRaw.trim() !== '') prevUnit = unitRaw;
        if (jobRaw.trim() !== '') prevJob = jobRaw;
        const units = splitLines(prevUnit);
        const jobs = splitLines(prevJob);
        const sites = splitLines(src[cSite]);
        const n = Math.max(units.length, jobs.length, sites.length);
        for (let i = 0; i < n; i++) {
          const unit = pickLine(units, i);
          const job = beforeSlash(pickLine(jobs, i));
          if (!unit && !job) continue;
          let mx = null;
          for (let k = 0; k < lexCols.length; k++) {
            const num = toNum(pickLine(splitLines(src[lexCols[k]]), i));
            if (num !== null && (mx === null || num > mx)) mx = num;
          }
          const key = noiseNorm(unit) + '\u0001' + noiseNorm(job);
          const cur = groupMax.get(key);
          groupMax.set(key, (cur === undefined || cur === null) ? mx : (mx === null ? cur : Math.max(cur, mx)));
          list.push({ unit: unit, job: job, site: pickLine(sites, i), key: key });
        }
      }
      // 把「同岗位最大值」写回每个点位（组内无任何数值时按 0，与参考表一致）
      for (const it of list) {
        const g = groupMax.get(it.key);
        it.lex = (g === null || g === undefined) ? 0 : g;
      }

      const app = window.SamplingApp;
      if (!app || typeof app.syncNonNoiseJobs !== 'function') { alert('主程序未提供同步能力，请刷新页面后重试。'); return; }
      const res = app.syncNonNoiseJobs(list) || { total: 0, matched: 0, changed: 0, unmatched: [] };
      const un = res.unmatched || [];
      let msg = '非噪声岗位同步完成：\n'
        + '· 噪声数据结果表共 ' + (res.total || 0) + ' 个岗位\n'
        + '· 与主表格匹配到 ' + (res.matched || 0) + ' 个\n'
        + '· 其中 ' + (res.changed || 0) + ' 个「是否噪声作业岗位」已置为「否」';
      if (un.length) {
        const MAX = 20;
        msg += '\n\n未匹配到的岗位（' + un.length + ' 个）：\n' + un.slice(0, MAX).map(function (it) {
          const parts = [];
          if (it.unit) parts.push(it.unit);
          if (it.job) parts.push(it.job);
          if (it.site) parts.push(it.site);
          return '· ' + (parts.join(' / ') || '(空)');
        }).join('\n');
        if (un.length > MAX) msg += '\n· …（其余 ' + (un.length - MAX) + ' 个未列出）';
      } else {
        msg += '\n\n未匹配到的岗位：无';
      }
      alert(msg);
    }
    const noiseSyncBtn = document.getElementById('noise-sync');
    if (noiseSyncBtn) noiseSyncBtn.addEventListener('click', syncNonNoiseJobs);

    // 暴露给采样计划主程序：数据上传按钮调用
    window.SamplingUpload = {
      show: showUpload,
      hide: hideUpload,
      logout: logout,
      // 登录失效（401）时统一提示并引导重新登录（供其它模块复用）
      notifyAuthExpired: notifyAuthExpired,
      // 加载「噪声数据结果」页签（供其它模块触发）
      loadNoise: loadNoiseResult,
      // 数据报告处理（检测结果报告单：回退 / 审核，供其它模块触发/复用）
      refreshReportPanel: refreshReportPanel,
      resetReportApproval: resetReportApproval,
      auditReport: auditReport,
      showFinalReportPdf: showFinalReportPdf,
      hideReportPdf: hideReportPdf,
      // 面板缓存（供测试/复用）：命中缓存则页签来回切换不联网
      invalidateReportCache: invalidateReportCache,
      getReportPanelCache: function () { return reportPanelCache; },
      // 「数据报告状态」查询/映射（供测试/复用）
      loadReportStatus: loadReportStatus,
      fetchReportStatus: fetchReportStatus,
      projectStatusText: projectStatusText,
      // 断点续审规划：按两个流程的当下节点快照算出「还需执行哪些步骤」（供测试/复用）
      planAuditSteps: planAuditSteps,
      // 动态续审：按当下节点快照求「下一个该执行的步骤」（一路审到底的核心，供测试/复用）
      nextAuditStep: nextAuditStep,
      nodeName: nodeName,
      // 解析**检评报告** docx 中「噪声检测结果分析与评价」标题下的表（供测试/复用）
      parseNoiseTable: parsePjbgNoiseTableFromDocx,
      // 渲染噪声表（供测试/复用）
      renderNoiseTable: renderNoiseTable,
      // 噪声筛选汇总：由展示用行计算四类结果（供测试/复用）
      buildNoiseSummary: buildNoiseSummary,
      renderNoiseSummary: renderNoiseSummary,
      // 非噪声岗位同步（供测试/复用）
      syncNonNoise: syncNonNoiseJobs,
      // 从检评报告 docx 算出「非噪声车间岗位」清单 [{unit, job}]（供「生成检评报告」删除噪声危害）
      noiseNonNoiseFromBuf: noiseNonNoiseFromBuf,
      // 从检评报告 docx 算出「非噪声车间岗位」分组/文本（供注段落拼接；空返回 ''）
      noiseNonNoiseGroupsFromBuf: noiseNonNoiseGroupsFromBuf,
      noiseNonNoiseTextFromBuf: noiseNonNoiseTextFromBuf,
      // 从检评报告 docx 算出「噪声车间岗位」（LEX 最大值 ≥ 80）分组（供后处理⑥判断是否为空）
      noiseWorkshopGroupsFromBuf: noiseWorkshopGroupsFromBuf,
      // ★ 一次性取全部噪声信息（②⑤⑥ 共用，避免同一份 docx 被解析多次）：
      //   { nonNoiseGroups, nonNoiseText, nonNoisePairs, noiseWorkshopGroups, quietPointPairs, ok }
      //   noiseInfoFromXml(xml) —— 复用上层已解压的 document.xml（最省，推荐）
      //   noiseInfoFromBuf(buf) —— 自行解压（同一 buf 只解析一次，带记忆）
      noiseInfoFromXml: noiseInfoFromXml,
      noiseInfoFromBuf: noiseInfoFromBuf,
      getCachedNoiseInfo: getCachedNoiseInfo,
      noiseInfoFromRows: noiseInfoFromRows,
      // 解压 docx 取 word/document.xml 字符串（供上层复用，避免重复解压）
      readDocxDocumentXml: readDocxDocumentXml,
      // 复用噪声页签已生成并下载的检评报告（供「生成检评报告」缩短时间）：
      // 命中同一项目时返回 { buf, fileName, modName }，否则 null
      getPjbgCache: function (projectId) {
        if (!pjbgBufCache) return null;
        if (projectId !== undefined && projectId !== null
          && String(pjbgBufCache.projectId) !== String(projectId)) return null;
        return pjbgBufCache;
      },
      // 工艺图片补传（供测试/复用）
      uploadProcessImages: uploadProcessImages,
      // 企业信息（在职人数/接害人数）读写（供测试/复用）
      fetchEnterpriseInfo: fetchEnterpriseInfo,
      editEnterpriseInfo: editEnterpriseInfo,
      buildEntInfoPatch: buildEntInfoPatch,
      entPayloadForUpload: entPayloadForUpload,
      applyEntInfoOnUpload: applyEntInfoOnUpload,
      // 企业基本信息自动补全（勾选「企业基本信息上传」时随调查表上传写入）：
      //   buildEntBasicPatch(rec) → { patch, notes }；entBasicChecked() = 当前勾选状态
      buildEntBasicPatch: buildEntBasicPatch,
      entValueBlank: entValueBlank,
      entBasicChecked: entBasicChecked,
      // 上传成功后清空「本次写入」的输入框 + 其缓存值（供测试/复用）；
      // 入参 patch = { enterpriseNum } / { hazardNum }，返回被清空的名称（'' 表示没清）
      clearEntInputsAfterUpload: clearEntInputsAfterUpload,
      // 已选择项目时的默认保存名（年份+单位名称，与上传成功后弹出的保存框同公式）；未选项目返回 null
      defaultSaveName: function () {
        if (!selectedProject) return null;
        const ym = /^BTC(\d{2})/.exec(selectedProject.code || '');
        return (ym ? ym[1] + '年' : '') + (selectedProject.belongInspectName || '');
      },
      // 供「网站数据导入」到上游取数：登录态 + 项目（不传项目时用「数据上传」当前选中的项目）
      getUpstreamContext: function (project) {
        const p = project || selectedProject;
        return {
          apiBase: API_BASE,
          token: token,
          orgId: orgId,
          loggedIn: !!token,
          project: p
            ? { id: p.id, code: p.code, name: p.belongInspectName || '' }
            : null,
        };
      },
      // 供「网站数据导入」独立搜索项目：与「数据上传」搜索完全同一套逻辑与缓存
      // （服务端 code 为包含匹配，见 runProjectSearch）
      searchProjectsFor: runProjectSearch
    };

    // 启动：先定登录方式（云端代登录 / 本地手输验证码），再决定是否取图
    (async function () {
      await initLoginMode();
      if (session && token) {
        teamSaved = await loadTeamSettings(); // 云端优先恢复本账号团队配置
        teamDirty = !!teamSaved;
        // 会话恢复时刷新一次权威用户信息，确保机构与平台网站一致
        try {
          const full = await fetchAuthInfo(token);
          if (full) {
            userInfo = Object.assign({}, userInfo || {}, full);
            const freshOrg = resolveOrgId(userInfo);
            if (freshOrg) {
              orgId = freshOrg;
              saveSession();
            }
          }
        } catch (e) {}
        syncGhButton();
        syncUserMenu();
        // 会话恢复期间若因登录失效被登出（token 已被清空），保持登录遮罩，不要再进入软件
        if (token) { enterApp(); warmAccountPool(); }
        else showLogin();
      } else {
        showLogin();
      }
    })();
  }

  if (typeof document !== 'undefined') {
    if (document.readyState === 'loading') {
      document.addEventListener('DOMContentLoaded', initApp);
    } else {
      initApp();
    }
  }

  // Node 环境导出（便于测试/后续集成）
  if (typeof module !== 'undefined' && module.exports) {
    module.exports = {
      API_BASE: API_BASE,
      rsaEncrypt: rsaEncrypt,
      fetchCaptcha: fetchCaptcha,
      login: login,
      searchProjects: searchProjects,
      fetchYearProjects: fetchYearProjects,
      uploadExcel: uploadExcel,
      fetchEnterpriseInfo: fetchEnterpriseInfo,
      editEnterpriseInfo: editEnterpriseInfo,
      buildEntInfoPatch: buildEntInfoPatch,
      entPayloadForUpload: entPayloadForUpload,
      applyEntInfoOnUpload: applyEntInfoOnUpload,
      buildEntBasicPatch: buildEntBasicPatch,
      entValueBlank: entValueBlank,
      fetchUsers: fetchUsers,
      fetchProjectMembers: fetchProjectMembers,
      addProjectMembers: addProjectMembers,
      deleteProjectMembers: deleteProjectMembers,
      fetchXcdcUser: fetchXcdcUser,
      saveXcdcUser: saveXcdcUser,
      syncProjectInfo: syncProjectInfo
    };
  }
})();
