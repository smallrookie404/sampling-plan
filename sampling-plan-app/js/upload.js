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
        syncGhButton();
        syncUserMenu();
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
        const load = window.SamplingApp && window.SamplingApp.loadScript;
        jszipPromise = (load ? load('js/jszip.min.js') : Promise.reject(new Error('缺少脚本加载能力')))
          .then(function () {
            if (!window.JSZip) throw new Error('jszip 加载失败');
            return window.JSZip;
          });
      }
      return jszipPromise;
    }

    const W_NS = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
    function noiseEsc(s) {
      return String(s === null || s === undefined ? '' : s)
        .replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
    }
    // 取文本：报告里一个单元格内的多行值写在同一个 <w:t> 里、用 <w:br/> 分隔
    // （WPS/Aspose 生成的非标准写法），须把 <w:br/> 还原成换行；单元格内多个段落用空格连接
    function runText(scope) {
      const ts = scope.getElementsByTagNameNS(W_NS, 't');
      const parts = [];
      for (let i = 0; i < ts.length; i++) {
        let s = '';
        const kids = ts[i].childNodes;
        for (let n = 0; n < kids.length; n++) {
          const c = kids[n];
          if (c.nodeType === 3) s += c.nodeValue || '';
          else if (c.nodeType === 1 && c.localName === 'br') s += '\n';
          else if (c.nodeType === 1 && c.localName === 'tab') s += '\t';
        }
        parts.push(s);
      }
      return parts.join('').replace(/[ \t]+/g, ' ').replace(/\n{2,}/g, '\n').replace(/^\s+|\s+$/g, '');
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

    // 从 docx 二进制提取「工作场所噪声评判结果」表，返回 { rows: string[][] }
    async function parseNoiseTableFromDocx(buf) {
      const JSZip = await ensureJSZip();
      const zip = await JSZip.loadAsync(buf);
      const f = zip.file('word/document.xml');
      if (!f) throw new Error('docx 结构异常：缺少 word/document.xml');
      const xml = await f.async('string');
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
        } else if (node.localName === 'tbl' && prev.indexOf('工作场所噪声评判结果') >= 0) {
          tbl = node;
          break;
        }
      }
      if (!tbl) throw new Error('未在报告中找到「工作场所噪声评判结果」表');
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
      { key: 'overWorkshop', title: '噪声超标车间岗位', rule: 'LEX 最大值 ≥ 85 dB', cls: 'over' },
    ];
    let noiseSummaryData = null;

    function beforeSlash(v) {
      const s = String(v === null || v === undefined ? '' : v);
      const i = s.indexOf('/');
      return (i < 0 ? s : s.slice(0, i)).trim();
    }
    function noiseNorm(s) { return String(s === null || s === undefined ? '' : s).replace(/\s+/g, ''); }

    // 由「展示用噪声行」计算筛选结果，返回 { nonNoise, overWorkshop, groupCount }
    // nonNoise / overWorkshop 为「按车间合并」后的数组：[{ unit, jobs:[岗位, ...] }]（顺序=首次出现）
    function buildNoiseSummary(display) {
      const out = { nonNoise: [], overWorkshop: [], groupCount: 0 };
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
        const n = Math.max(units.length, jobs.length);
        for (let i = 0; i < n; i++) {
          const unit = pickLine(units, i);
          const job = beforeSlash(pickLine(jobs, i));
          if (!unit && !job) continue;
          const key = noiseNorm(unit) + '\u0001' + noiseNorm(job);
          let g = groups.get(key);
          if (!g) { g = { unit: unit, job: job, max: 0 }; groups.set(key, g); }
          for (let k = 0; k < lexCols.length; k++) {
            const num = toNum(pickLine(splitLines(src[lexCols[k]]), i));
            if (num !== null && num > g.max) g.max = num;
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
      const overIdx = new Map();
      groups.forEach(function (g) {
        if (g.max < NOISE_NON_LIMIT) addGroup(out.nonNoise, nonIdx, g);
        if (g.max >= NOISE_OVER_LIMIT) addGroup(out.overWorkshop, overIdx, g);
      });
      return out;
    }

    // 把「按车间合并」的结果拼成文本：车间岗位、岗位；车间岗位、岗位
    function noiseSummaryText(list) {
      return (list || []).map(function (g) { return g.unit + g.jobs.join('、'); }).join('；');
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
      const hasAny = s.nonNoise.length || s.overWorkshop.length;
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
        if (status) status.textContent = '项目 ' + (selectedProject.code || '') + ' · 共 ' + Math.max(0, noiseCache.data.rows.length - 1) + ' 行';
        return;
      }
      noiseLoading = true;
      setNoiseTabLoading(true);
      if (empty) empty.classList.add('hidden');
      box.innerHTML = '';
      if (status) status.textContent = '加载中…（正在生成并解析检测报告）';
      try {
        const rr = await apiRequest(
          'GET',
          '/api/jcbgReport/jcbgYl?projectId=' + encodeURIComponent(pid) + '&organizationId=' + encodeURIComponent(orgId || ''),
          { token: token, orgId: orgId, timeout: 120000 }
        );
        const body = rr.data && rr.data.body;
        if (rr.status !== 200 || typeof body !== 'string' || !body) {
          throw new Error((rr.data && (rr.data.message || rr.data.msg)) || ('HTTP ' + rr.status));
        }
        const enc = body.replace(/\.pdf$/i, '.docx').split('/').map(encodeURIComponent).join('/');
        const resp = await fetch(API_BASE + enc, { headers: token ? { Authorization: token } : {} });
        if (!resp.ok) throw new Error('下载检测报告失败(HTTP ' + resp.status + ')');
        const data = await parseNoiseTableFromDocx(await resp.arrayBuffer());
        noiseCache = { projectId: pid, data: data };
        renderNoiseTable(box, data);
        if (status) status.textContent = '项目 ' + (selectedProject.code || '') + ' · 共 ' + Math.max(0, data.rows.length - 1) + ' 行';
      } catch (e) {
        box.innerHTML = '';
        renderNoiseSummary(null);
        if (empty) {
          empty.textContent = '加载失败：' + (e && e.message ? e.message : e);
          empty.classList.remove('hidden');
        }
        if (status) status.textContent = '';
      } finally {
        noiseLoading = false;
        setNoiseTabLoading(false);
      }
    }

    const noiseTabBtn = document.querySelector('.tab[data-tab="noise"]');
    if (noiseTabBtn) noiseTabBtn.addEventListener('click', function () { loadNoiseResult(false); });
    const noiseRefreshBtn = document.getElementById('noise-refresh');
    if (noiseRefreshBtn) noiseRefreshBtn.addEventListener('click', function () { loadNoiseResult(true); });

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
      const res = app.syncNonNoiseJobs(list) || { matched: 0, changed: 0 };
      alert('非噪声岗位同步完成：\n· 与主表格匹配到 ' + res.matched + ' 个岗位\n· 其中 ' + res.changed + ' 个「是否噪声作业岗位」已置为「否」');
    }
    const noiseSyncBtn = document.getElementById('noise-sync');
    if (noiseSyncBtn) noiseSyncBtn.addEventListener('click', syncNonNoiseJobs);

    // 暴露给采样计划主程序：数据上传按钮调用
    window.SamplingUpload = {
      show: showUpload,
      hide: hideUpload,
      logout: logout,
      // 加载「噪声数据结果」页签（供其它模块触发）
      loadNoise: loadNoiseResult,
      // 解析报告 docx 中的「工作场所噪声评判结果」表（供测试/复用）
      parseNoiseTable: parseNoiseTableFromDocx,
      // 渲染噪声表（供测试/复用）
      renderNoiseTable: renderNoiseTable,
      // 噪声筛选汇总：由展示用行计算四类结果（供测试/复用）
      buildNoiseSummary: buildNoiseSummary,
      renderNoiseSummary: renderNoiseSummary,
      // 非噪声岗位同步（供测试/复用）
      syncNonNoise: syncNonNoiseJobs,
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
        syncGhButton();
        syncUserMenu();
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
