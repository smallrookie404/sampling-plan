/* =====================================================================
 * SamplingUtil —— 全站共享的通用工具函数（无构建、经典脚本、IIFE）
 *
 * 背景：app.js / blf.js / upload.js / survey.js 各自重复实现了同一批
 * 小工具（脚本懒加载、HTML/XML 转义、列表结构兼容、空白归一化…）。
 * 这里统一一份，业务脚本只引用 window.SamplingUtil，避免「一处改了、
 * 另一处没改」的维护陷阱。本模块**不依赖任何其它模块**，且必须在业务
 * 脚本之前加载（见 index.html 的 <script> 顺序）。
 * ===================================================================== */
(function (root) {
  'use strict';

  // ---------- 脚本懒加载 ----------
  // 经典脚本按 src 去重：同一 src 只插一次 <script>（多次调用共享同一个 Promise）。
  const scriptPromises = Object.create(null);

  function loadScript(src) {
    if (scriptPromises[src]) return scriptPromises[src];
    scriptPromises[src] = new Promise(function (resolve, reject) {
      const s = document.createElement('script');
      s.src = src;
      s.onload = function () { resolve(); };
      s.onerror = function () { reject(new Error('脚本加载失败：' + src)); };
      document.head.appendChild(s);
    });
    return scriptPromises[src];
  }

  // 顺序加载一组脚本
  function loadScripts(list) {
    return (list || []).reduce(function (p, src) {
      return p.then(function () { return loadScript(src); });
    }, Promise.resolve());
  }

  // ---------- 转义 ----------
  const AMP = /&/g, LT = /</g, GT = />/g, QUOT = /"/g, APOS = /'/g;

  // HTML 文本转义（& < > 及可选引号）
  function escHtml(s) {
    return String(s === null || s === undefined ? '' : s)
      .replace(AMP, '&amp;').replace(LT, '&lt;').replace(GT, '&gt;')
      .replace(QUOT, '&quot;').replace(APOS, '&#39;');
  }

  // HTML 属性转义（只处理引号，用于已确定无 <> 的属性值）
  function escAttr(s) {
    return String(s === null || s === undefined ? '' : s)
      .replace(AMP, '&amp;').replace(QUOT, '&quot;');
  }

  // XML 内容转义（xlsx / docx 生成用，含单引号）
  function escXml(s) {
    return String(s === null || s === undefined ? '' : s)
      .replace(AMP, '&amp;').replace(LT, '&lt;').replace(GT, '&gt;')
      .replace(QUOT, '&quot;').replace(APOS, '&apos;');
  }

  // ---------- 空白归一化 ----------
  // 去掉全部空白（含全角空格）：用于表头/名称比对
  function stripWs(s) {
    return String(s === null || s === undefined ? '' : s).replace(/[\s\u3000]+/g, '');
  }

  // 去掉首尾空白并把连续空白压成一个空格
  function squashWs(s) {
    return String(s === null || s === undefined ? '' : s).replace(/\s+/g, ' ').trim();
  }

  // ---------- 平台接口返回结构兼容 ----------
  // 平台列表接口的返回形态多样：可能是数组本身，也可能是
  // { body | data | rows | records: [...] }，或再嵌一层。按优先级取第一个数组。
  function pickRows(j) {
    if (Array.isArray(j)) return j;
    const cands = [j && j.body, j && j.data, j && j.rows, j && j.records, j];
    for (let i = 0; i < cands.length; i++) {
      const c = cands[i];
      if (Array.isArray(c)) return c;
      if (c && typeof c === 'object') {
        const keys = ['records', 'rows', 'list', 'data'];
        for (let k = 0; k < keys.length; k++) {
          if (Array.isArray(c[keys[k]])) return c[keys[k]];
        }
      }
    }
    return [];
  }

  // 行对象按候选字段名取第一个「非空」值（新旧版字段改名兼容）
  function pickField(r, keys) {
    if (!r) return '';
    for (let i = 0; i < keys.length; i++) {
      const v = r[keys[i]];
      if (v !== undefined && v !== null && v !== '') return v;
    }
    return '';
  }

  root.SamplingUtil = {
    loadScript: loadScript,
    loadScripts: loadScripts,
    escHtml: escHtml,
    escAttr: escAttr,
    escXml: escXml,
    stripWs: stripWs,
    squashWs: squashWs,
    pickRows: pickRows,
    pickField: pickField,
  };
})(typeof self !== 'undefined' ? self : this);
