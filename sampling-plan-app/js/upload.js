/* =====================================================================
 * 现场调查 Excel 上传（整合进采样计划软件）
 * - 登录界面（账号/密码/验证码）作为整个软件的总登录门禁；
 * - 「数据上传」按钮切换到本视图；
 * - 上传文件自动使用当前表格导出的 Excel（无需手动选择）。
 * 原项目：zhiyeweishen/site-survey-upload-web
 * ===================================================================== */
(function () {
  'use strict';

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
    const timer = setTimeout(function () { ac.abort(); }, opts.timeout || 30000);
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
      form: fd
    });
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
    const toDel = existing.filter(function (m) { return want.indexOf(String(m.humanId)) < 0; });
    if (toAdd.length > 0) {
      const names = toAdd.map(function (id) {
        const u = (opts.users || []).filter(function (x) { return String(x.id) === id; })[0];
        return u ? u.userName : id;
      });
      const r = await addProjectMembers(token, projectId, toAdd);
      if (r.status === 200 && r.data && r.data.code === '200') {
        logs.push('添加项目成员成功：' + names.join('、'));
      } else {
        logs.push('添加项目成员失败：' + ((r.data && r.data.message) || ('HTTP ' + r.status)));
      }
    }
    if (toDel.length > 0) {
      const r = await deleteProjectMembers(token, toDel.map(function (m) { return m.humanId; }));
      if (r.status === 200 && r.data && r.data.code === '200') {
        logs.push('移除项目成员成功：' + toDel.map(function (m) { return m.humanIdName || m.humanId; }).join('、'));
      } else {
        logs.push('移除项目成员失败：' + ((r.data && r.data.message) || ('HTTP ' + r.status)));
      }
    }
    if (opts.investigatorId || opts.reviewerId || opts.investigateDate) {
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
        logs.push('调查人/复核人同步失败：' + ((r.data && r.data.message) || ('HTTP ' + r.status)));
      }
    }
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
    }

    // 显示登录遮罩
    function showLogin() {
      if (loginOverlay) loginOverlay.classList.remove('hidden');
      if (uploadOverlay) uploadOverlay.classList.add('hidden');
      refreshCaptcha();
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
    }

    function hideUpload() {
      if (uploadOverlay) uploadOverlay.classList.add('hidden');
    }

    function logout() {
      clearSession();
      session = null;
      token = null;
      orgId = null;
      userInfo = null;
      teamSaved = null;
      selectedProject = null;
      selectedFile = null;
      userList = [];
      confirmedMemberIds = [];
      teamDirty = false;
      showLogin();
    }

    async function refreshCaptcha() {
      try {
        const d = await fetchCaptcha();
        uuid = d.uuid;
        captchaImg.src = d.img;
        $('captcha').value = '';
        loginMsg.textContent = '';
      } catch (e) {
        let msg = e.message;
        // GitHub Pages 等 HTTPS 静态托管无法直连平台 HTTP 接口，给出明确指引
        try {
          if (location.protocol === 'https:' && !/pages\.dev$/i.test(location.hostname)) {
            msg = 'HTTPS 页面无法直连平台 HTTP 接口（被浏览器拦截）。请使用 Cloudflare Pages 部署：https://sampling-plan.pages.dev';
          }
        } catch (e2) {}
        loginMsg.textContent = msg;
      }
    }

    $('btnRefresh').addEventListener('click', refreshCaptcha);
    captchaImg.addEventListener('click', refreshCaptcha);

    // 打开页面时自动填充已保存的账号密码
    try {
      if (localStorage.getItem(REM_KEY) === '1') {
        $('rememberMe').checked = true;
        $('username').value = localStorage.getItem(USR_KEY) || '';
        $('password').value = localStorage.getItem(PWD_KEY) || '';
      }
    } catch (e) { }

    $('btnLogin').addEventListener('click', async function () {
      const user = $('username').value.trim();
      const pass = $('password').value;
      const code = $('captcha').value.trim();
      if (!user || !pass || !code) { loginMsg.textContent = '请填写账号、密码和验证码'; return; }
      const btn = $('btnLogin');
      btn.disabled = true;
      btn.textContent = '登录中...';
      try {
        const r = await login(user, pass, code, uuid);
        if (r.status !== 200 || !r.data || !r.data.token) {
          const msg = (r.data && (r.data.message || r.data.msg)) || ('HTTP ' + r.status);
          throw new Error(msg);
        }
        token = r.data.token;
        userInfo = r.data.user.user || {};
        // 登录后调用 auth/info 获取权威用户信息（与平台网站一致），解析真实机构
        try {
          const full = await fetchAuthInfo(token);
          if (full) {
            userInfo = Object.assign({}, userInfo, full);
            if (full.userName) userInfo.userName = full.userName;
            if (full.userCode) userInfo.userCode = full.userCode;
          }
        } catch (e) {}
        orgId = resolveOrgId(userInfo) || (userInfo.organization && userInfo.organization.organizationId) || null;
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
        $('greeting').textContent = greetingText();
        enterApp();
        log('登录成功：' + userInfo.userName + '（' + userInfo.userCode + '），上传主体为当前账号');
        prefetchYear();
        loadUserList();
      } catch (e) {
        loginMsg.textContent = '登录失败：' + e.message;
        refreshCaptcha();
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
    if (!unitInput || !codeInput || !yearSelect || !suggestBox || !$('searching') || !$('projectInfo') ||
        !memberSearch || !memberList || !memberTags || !btnMemberConfirm || !investigatorSelect || !reviewerSelect || !investigateDate) {
      if (typeof console !== 'undefined') console.error('上传页面元素缺失，请强制刷新浏览器（Ctrl+F5）后重试');
      alert('上传页面加载不完整，请按 Ctrl+F5 强制刷新后重试');
      return;
    }
    let searchSeq = 0;
    let currentAbort = null;
    let debounceTimer = null;

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
        });
        suggestBox.appendChild(item);
      });
      suggestBox.classList.remove('hidden');
    }

    function doSearch() {
      try {
        const unitName = unitInput.value.trim();
        const code = codeInput.value.trim();
        const yearPrefix = yearSelect.value ? 'BTC' + yearSelect.value : '';
        if (!unitName && !code && !yearPrefix) {
          suggestBox.classList.add('hidden');
          log('请先输入受检单位或项目编号，再点击搜索');
          return;
        }
        selectedProject = null;
        $('projectInfo').textContent = '';
        $('projectInfo').className = 'info';
        const seq = ++searchSeq;
        if (currentAbort) currentAbort.abort();
        currentAbort = new AbortController();
        setSearching(true);
        const typedHasYearPrefix = /^BTC\d{2}/i.test(code);
        let promise;
        if (unitName) {
          // 受检单位走服务端模糊匹配（小请求），避免每敲一个字都拉取整年项目
          promise = searchProjects(token, orgId, { code: code || yearPrefix, unitName: unitName }, currentAbort.signal)
            .then(function (list) {
              let filtered = list;
              // 未带年份前缀的编号片段：在服务端结果内再做本地包含匹配
              if (code && !typedHasYearPrefix) {
                filtered = list.filter(function (p) { return p.code && String(p.code).indexOf(code) >= 0; });
              }
              return filtered.slice(0, 15);
            });
        } else if (yearPrefix && !typedHasYearPrefix) {
          // 仅按年份/编号片段浏览：拉取当年范围（30 分钟缓存 + sessionStorage），本地按编号过滤
          promise = fetchYearProjects(token, orgId, yearPrefix, '', currentAbort.signal)
            .then(function (list) {
              let filtered = list;
              if (code) {
                filtered = list.filter(function (p) {
                  return p.code && String(p.code).indexOf(code) >= 0;
                });
              }
              return filtered.slice(0, 15);
            });
        } else if (code || unitName) {
          promise = searchProjects(token, orgId, { code: code, unitName: unitName }, currentAbort.signal);
        } else {
          promise = Promise.resolve([]);
        }
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

    $('btnUpload').addEventListener('click', async function () {
      if (!selectedProject) { log('请先查询项目（输入项目编号后点搜索并选择）'); return; }
      if (!token) { showLogin(); return; }
      const errs = window.SamplingApp && window.SamplingApp.countErrors ? window.SamplingApp.countErrors() : 0;
      if (errs > 0 && !confirm('当前表格有 ' + errs + ' 处校验错误，仍要上传吗？')) return;
      // 上传前重新生成一次，确保文件为最新表格内容
      await generateExportFile();
      if (!selectedFile) { log('上传文件生成失败，请先处理表格内容后重试'); return; }
      const projectId = selectedProject.id;
      const ok = confirm('确认将文件上传到项目 ' + selectedProject.code + '（ID=' + projectId + '）？\n\n' + selectedFile.name);
      if (!ok) return;
      const btn = $('btnUpload');
      btn.disabled = true;
      btn.textContent = '上传中...';
      try {
        log('开始上传（主体：' + userInfo.userCode + ' ' + userInfo.userName + '）：' + selectedFile.name);
        const r = await uploadExcel(token, orgId, projectId, selectedFile);
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
              investigatorId: investigatorSelect.value,
              reviewerId: reviewerSelect.value,
              investigateDate: investigateDate.value
            });
            syncLogs.forEach(function (l) { log(l); });
            const failed = syncLogs.filter(function (l) { return l.indexOf('失败') >= 0; });
            syncMsg = failed.length > 0 ? '，团队信息同步有失败项，详见日志' : '，团队信息已同步';
          } catch (e) {
            log('同步团队信息失败：' + e.message);
            syncMsg = '，团队信息同步失败：' + e.message;
          }
          if (msgs.length > 0) {
            log('导入完成，但有以下提示：' + msgs.join('；'));
            alert('导入完成，提示：\n' + msgs.join('\n'));
          } else {
            log('导入成功！' + syncMsg);
            alert('导入成功！' + syncMsg);
          }
        } else {
          const msg = (d && (d.message || d.msg)) || ('HTTP ' + r.status);
          log('导入失败：' + msg);
          alert('导入失败：' + msg);
        }
      } catch (e) {
        log('上传异常：' + e.message);
        alert('上传异常：' + e.message);
      } finally {
        btn.disabled = false;
        btn.textContent = '上传到现场调查';
      }
    });

    // 视图切换：返回采样计划 / 退出登录
    $('xcdcBack').addEventListener('click', hideUpload);

    // 暴露给采样计划主程序：数据上传按钮调用
    window.SamplingUpload = {
      show: showUpload,
      hide: hideUpload,
      logout: logout
    };

    // 启动：有会话直接进入（登录遮罩保持隐藏），否则显示登录
    (async function () {
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
        enterApp();
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
