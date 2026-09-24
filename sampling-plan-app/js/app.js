(function () {
  "use strict";

  const L = window.SamplingLogic;
  const DEFAULT_DATA = window.SamplingData;

  // xlsx 读写模块（jszip + xlsxio）改为按需加载：首次导入/导出时才加载，加快页面启动
  let X = null;
  let xlsxPromise = null;
  function loadScript(src) {
    return new Promise((resolve, reject) => {
      const s = document.createElement("script");
      s.src = src;
      s.onload = resolve;
      s.onerror = () => reject(new Error("脚本加载失败：" + src));
      document.head.appendChild(s);
    });
  }
  function ensureXlsx() {
    if (X) return Promise.resolve(X);
    if (window.SamplingXlsx) { X = window.SamplingXlsx; return Promise.resolve(X); }
    if (!xlsxPromise) {
      xlsxPromise = (async () => {
        await loadScript("js/jszip.min.js");
        await loadScript("js/xlsxio.js");
        X = window.SamplingXlsx;
        if (!X) throw new Error("xlsx 模块加载失败");
        return X;
      })();
    }
    return xlsxPromise;
  }

  // ---------- 列定义 ----------
  const INPUT_COLS = L.INPUT_COLS; // A..V（V 为录入区备注列）
  const COMPUTED_COLS = L.COMPUTED_COLS; // W..BI
  const MANUAL_COLS = L.MANUAL_COLS;
  const OVERRIDE_COLS = L.OVERRIDE_COLS;
  const TEXT_OVERRIDE_COLS = L.TEXT_OVERRIDE_COLS || ["BI"];
  const SELECT_COLS = new Set(["Y", "Z", "AO", "AR", "U", "AI", "AJ", "BH"]); // 下拉单元格列
  // 录入区下拉联想选项：作业方式 / 采样方式（datalist，可手动录入）
  const ZUOYE_FS = ["手工作业", "半手工作业", "全自动作业"];
  const CAIYANG_FS = ["定点", "个体"];
  const NUM_COLS = new Set(["E", "F", "G", "H", "I", "AA", "AB", "AG", "AH", "AY", "AZ", "BA", "BB"]);
  // 显示顺序：A..M → V（备注，位于“是否采样”后面）→ N..U → W..BI
  const ALL_COLS = [
    ...INPUT_COLS.slice(0, 13),
    "V",
    ...INPUT_COLS.slice(13, 21),
    ...COMPUTED_COLS,
  ]; // 61 列
  const MAIN_HEADERS = DEFAULT_DATA.mainHeaders; // A..U + "" 分隔列 + W..BI
  const HEADERS = [
    ...MAIN_HEADERS.slice(0, 13),
    "备注",
    ...MAIN_HEADERS.slice(13, 21),
    ...MAIN_HEADERS.slice(22),
  ]; // 61 项，与 ALL_COLS 一一对应
  const ROW_H = 30;
  const MIN_ROW_H = 24;
  const MAX_ROW_H = 300;
  const MIN_COL_W = 30;
  const MAX_COL_W = 600;

  const colIdx = (letter) => ALL_COLS.indexOf(letter);
  // 各列宽度（Excel 字符单位近似）
  const COL_W = ALL_COLS.map((c) => {
    const w = {
      A: 12, B: 12, C: 10, D: 18, E: 9, F: 9, G: 9, H: 7, I: 8, J: 16,
      K: 14, L: 12, M: 8, N: 8, O: 9, P: 8, Q: 8, R: 11, S: 13, T: 11,
      U: 10, V: 14, W: 12, X: 12, Y: 9, Z: 16, AA: 10, AB: 9, AC: 12,
      AD: 11, AE: 13, AF: 11, AG: 12, AH: 12, AI: 12, AJ: 12, AK: 11,
      AL: 12, AM: 20, AN: 20, AO: 12, AP: 11, AQ: 11, AR: 9, AS: 11,
      AT: 9, AU: 9, AV: 8, AW: 10, AX: 10, AY: 11, AZ: 10, BA: 9,
      BB: 9, BC: 9, BD: 12, BE: 14, BF: 8, BG: 8, BH: 8, BI: 18,
    }[c];
    return w || 10;
  });
  const COMPUTED_HEADERS = COMPUTED_COLS.map((c) => HEADERS[colIdx(c)]);
  const COMPUTED_WIDTHS = COMPUTED_COLS.map((c) => COL_W[colIdx(c)]);

  // ---------- 列宽自适应（按浏览器实际文字宽度测量） ----------
  const measureCtx = (() => {
    const c = document.createElement("canvas");
    return c.getContext("2d");
  })();
  const CELL_FONT = "'Microsoft YaHei','PingFang SC','Segoe UI','Arial',sans-serif";
  const measureCache = new Map(); // 列宽测量缓存：同一文本只测量一次

  function measureTextWidth(text, sizePx, bold) {
    const key = (bold ? "b" : "n") + "|" + sizePx + "|" + String(text);
    const hit = measureCache.get(key);
    if (hit !== undefined) return hit;
    measureCtx.font = `${bold ? "bold " : ""}${sizePx}px ${CELL_FONT}`;
    const w = measureCtx.measureText(String(text)).width;
    if (measureCache.size > 30000) measureCache.clear(); // 防缓存无限增长
    measureCache.set(key, w);
    return w;
  }

  function columnWidths(cols, getHeader, getValue, rowsList, dataSize, headerSize, min, max, pad) {
    return cols.map((_, i) => {
      let w = measureTextWidth(getHeader(i), headerSize, true) + pad + 2;
      for (const r of rowsList) {
        const v = getValue(r, i);
        if (v === "" || v === null || v === undefined) continue;
        const t = String(v);
        if (t.length > 60) continue; // 超长文本由省略号处理，避免列宽失控
        const cw = measureTextWidth(t, dataSize, false) + pad;
        if (cw > w) w = cw;
      }
      return Math.max(min, Math.min(max, Math.ceil(w)));
    });
  }

  // 按内容计算主表各列所需宽度（“自动调整”按钮使用）
  function computeGridWidths() {
    const w = columnWidths(
      ALL_COLS,
      (i) => HEADERS[i],
      (r, i) => {
        const c = ALL_COLS[i];
        if (INPUT_COLS.includes(c)) return r.input[c];
        if (MANUAL_COLS.includes(c)) return r.manual[c];
        return r.values[c];
      },
      rows,
      12,
      11.5,
      56,
      380,
      14
    );
    // 下拉单元格额外预留右侧箭头空间，文字完整显示不被截断
    ALL_COLS.forEach((c, i) => { if (SELECT_COLS.has(c)) w[i] += 24; });
    return w;
  }

  // 应用当前列宽（手动拖动或自动调整后的缓存值；首次启动按内容计算一次）
  function applyGridWidths() {
    const w = colWidths || (colWidths = computeGridWidths());
    $("grid-cols").innerHTML =
      "<col style='width:42px'>" + w.map((x) => `<col style="width:${x}px">`).join("");
    // 前四列（车间/岗位/工种/点位/接害因素）与行号列锁定：设置各锁定列的 left 偏移
    const rownoW = 42;
    gridWrap.style.setProperty("--sticky-a", rownoW + "px");
    gridWrap.style.setProperty("--sticky-b", rownoW + w[0] + "px");
    gridWrap.style.setProperty("--sticky-c", rownoW + w[0] + w[1] + "px");
    gridWrap.style.setProperty("--sticky-d", rownoW + w[0] + w[1] + w[2] + "px");
  }

  function applyHazardWidths() {
    const w = columnWidths(
      HAZARD_KEYS,
      (i) => HAZARD_HEADERS[i],
      (h, i) => h[HAZARD_KEYS[i]],
      hazardFactors,
      12,
      12,
      60,
      360,
      14
    );
    // 下拉单元格（粉尘性质、是/否 等）额外预留箭头空间
    w.forEach((x, i) => { if (i === 2 || (i >= 3 && i <= 8) || i === 10) w[i] += 24; });
    $("hazard-cols").innerHTML =
      "<col style='width:42px'>" + w.map((x) => `<col style="width:${x}px">`).join("");
  }

  // ---------- 状态 ----------
  let rows = [];
  let rowHeights = []; // 每行高度（px），与 rows 一一对应；默认 ROW_H，可手动拖动或“自适应行距”调整
  let colWidths = null; // 列宽（px），与 ALL_COLS 一一对应；手动拖动或“自动调整”时更新
  let hazardFactors = [];
  let detectionItems = [];
  let selectedRow = -1;
  let selectedHazard = -1;
  let activeTab = "main";

  const $ = (id) => document.getElementById(id);
  const gridWrap = $("grid-wrap");
  const gridBody = $("grid-body");
  const gridHead = $("grid-head");

  // ---------- 行模型 ----------
  function blankRow() {
    const input = {};
    for (const c of INPUT_COLS) input[c] = "";
    const manual = {};
    for (const c of MANUAL_COLS) manual[c] = "";
    return { input, manual, overridden: {}, values: {}, errors: {} };
  }

  function emptyGrid(n = 50) {
    rows = [];
    for (let i = 0; i < n; i++) rows.push(blankRow());
    rowHeights = rows.map(() => ROW_H);
    rowOffsets = null;
  }

  // ---------- 表头 ----------
  function buildHead() {
    const groupTh = (label, colSpan) =>
      `<th class="group" colspan="${colSpan}">${label}</th>`;
    const fieldTh = (label, cls, col) => `<th class="field${cls ? " " + cls : ""}"${col ? ` data-c="${col}"` : ""}>${label || "&nbsp;"}</th>`;
    let groupHtml = `<th class="corner sticky-corner" rowspan="2" style="width:40px">行</th>`;
    groupHtml += groupTh("录 入 区", INPUT_COLS.length);
    groupHtml += groupTh("自动计算区（与原表 W~BI 列一致）", COMPUTED_COLS.length);
    let fieldHtml = "";
    for (const c of ALL_COLS) {
      const cls = c === "A" ? "sticky-col-a" : c === "B" ? "sticky-col-b" : c === "C" ? "sticky-col-c" : c === "D" ? "sticky-col-d" : "";
      fieldHtml += fieldTh(HEADERS[colIdx(c)], cls, c);
    }
    gridHead.innerHTML =
      `<tr class="group-row">${groupHtml}</tr>` +
      `<tr class="field-row">${fieldHtml}</tr>`;
  }

  // ---------- 单元格渲染 ----------
  function cellClass(col) {
    if (OVERRIDE_COLS.includes(col)) return "override";
    if (TEXT_OVERRIDE_COLS.includes(col)) return "override";
    if (MANUAL_COLS.includes(col)) return "manual";
    if (COMPUTED_COLS.includes(col)) return "computed";
    return "";
  }

  function cellHtml(row, idx, col) {
    let cls = cellClass(col);
    if (col === "A") cls += " sticky-col-a";
    else if (col === "B") cls += " sticky-col-b";
    else if (col === "C") cls += " sticky-col-c";
    else if (col === "D") cls += " sticky-col-d";
    const num = NUM_COLS.has(col) ? " cell-num" : "";
    let inner;
    if (OVERRIDE_COLS.includes(col)) {
      inner = selectHtml(col, row.values[col], row.overridden[col] === true);
    } else if (col === "U") {
      inner = selectHtml(col, row.input[col], false, true);
    } else if (col === "R" || col === "S") {
      // 岗位工作班制 / 岗位班制数：下拉联想 + 可手动录入（支持复制粘贴）
      const dl = col === "R" ? ' list="banzhi-r-dl"' : ' list="banzhi-s-dl"';
      inner = `<input data-c="${col}"${dl} value="${escAttr(fmt(row.input[col]))}">`;
    } else if (MANUAL_COLS.includes(col)) {
      if (col === "AI" || col === "AJ" || col === "BH") {
        inner = selectHtml(col, row.manual[col], false, true);
      } else {
        inner = `<input data-c="${col}" value="${escAttr(fmt(row.manual[col]))}">`;
      }
    } else if (TEXT_OVERRIDE_COLS.includes(col)) {
      inner = `<input data-c="${col}" value="${escAttr(fmt(row.values[col]))}">`;
    } else if (COMPUTED_COLS.includes(col)) {
      // 只读计算列用纯文本而非 input：行数多时常驻输入框节点拖慢渲染（卡顿主因）
      inner = `<div class="computed-text" data-c="${col}">${escHtml(fmt(row.values[col]))}</div>`;
    } else {
      // D 接害因素 / L 作业方式 / N 采样方式 / O 岗位性质：下拉联想 + 可手动录入
      const dlMap = { D: "hazard-dl", L: "zuoye-dl", N: "caiyang-dl", O: "gangwei-dl" };
      const dl = dlMap[col] ? ` list="${dlMap[col]}"` : "";
      inner = `<input data-c="${col}"${dl} value="${escAttr(fmt(row.input[col]))}">`;
    }
    return `<td class="${cls}${num}" data-r="${idx}" data-c="${col}">${inner}</td>`;
  }

  function selectHtml(col, value, isOverride, editable) {
    const options = {
      R: ["", ...L.BANZHI_R],
      S: ["", ...L.BANZHI_S],
      U: ["", ...L.TILI_LD],
      Y: ["", ...L.GANGWEI_XZ],
      Z: ["", ...L.JIECHU_LX],
      AO: ["", ...L.SHI_FOU],
      AR: ["", ...L.JIANCE_FS],
      AI: ["", ...L.SHI_FOU],
      AJ: ["", ...L.SHI_FOU],
      BH: ["", ...L.SHI_FOU],
    }[col] || [];
    const cur = value;
    let opts = "";
    if (isOverride) opts += `<option value="">（自动计算）</option>`;
    else opts += `<option value=""></option>`;
    for (const o of options) {
      if (o === "") continue;
      opts += `<option value="${escAttr(o)}"${cur === o ? " selected" : ""}>${escHtml(o)}</option>`;
    }
    return `<select data-c="${col}" ${editable ? "" : ""}>${opts}</select>`;
  }

  function fmt(v) {
    return v === null || v === undefined ? "" : String(v);
  }

  function escHtml(s) {
    return String(s).replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;");
  }
  function escAttr(s) {
    return escHtml(s).replace(/"/g, "&quot;");
  }

  // ---------- 虚拟滚动渲染 ----------
  let renderedRows = [];
  let scrollRenderQueued = false;
  let rowOffsets = null; // 行高前缀和缓存（行高变更时置空重建）
  let lastRenderedStart = -1;

  function rowHeightAt(i) {
    const h = rowHeights[i];
    return typeof h === "number" && h > 0 ? h : ROW_H;
  }

  function buildRowOffsets() {
    const arr = new Array(rows.length + 1);
    let acc = 0;
    arr[0] = 0;
    for (let i = 0; i < rows.length; i++) {
      acc += rowHeightAt(i);
      arr[i + 1] = acc;
    }
    rowOffsets = arr;
  }

  // 第 i 行顶边的累计 Y 偏移（0 起点），基于缓存前缀和
  function rowOffsetAt(i) {
    if (!rowOffsets) buildRowOffsets();
    return rowOffsets[i] ?? 0;
  }

  function totalGridHeight() {
    return rowOffsetAt(rows.length);
  }

  // 由滚动位置找到对应的行号（行高不一致时按累计高度定位）
  function rowIndexAt(y) {
    if (!rowOffsets) buildRowOffsets();
    if (rows.length === 0) return 0;
    let lo = 0;
    let hi = rows.length - 1;
    while (lo < hi) {
      const mid = (lo + hi + 1) >> 1;
      if (rowOffsets[mid] <= y) lo = mid;
      else hi = mid - 1;
    }
    return lo;
  }

  // 渲染缓冲：上下各多渲染的行数。增量渲染接管滚动后无需过大缓冲
  const RENDER_OVERSCAN = 24;

  // 单行 HTML（renderWindow 与增量渲染共用）
  function rowHtml(i) {
    const r = rows[i];
    const h = rowHeightAt(i);
    let cells = `<td class="rowno sticky-corner${i === selectedRow ? " selected" : ""}" data-r="${i}">${i + 1}</td>`;
    for (const c of ALL_COLS) cells += cellHtml(r, i, c);
    return `<tr data-r="${i}" style="height:${h}px"${i === selectedRow ? ' class="selected"' : ""}>${cells}</tr>`;
  }

  function renderWindow() {
    const total = rows.length;
    const st = gridWrap.scrollTop;
    const ch = gridWrap.clientHeight;
    const start = Math.max(0, rowIndexAt(st) - RENDER_OVERSCAN);
    const visible = Math.ceil(ch / ROW_H) + RENDER_OVERSCAN * 2;
    let end = start;
    let count = 0;
    while (end < total && count < visible) {
      end++;
      count++;
    }
    let html = "";
    renderedRows = [];
    lastRenderedStart = start;
    // 顶部占位行：保证总高度恒定，虚拟滚动才能稳定滚动到底（始终保留，供增量渲染调整高度）
    const topSpacer = rowOffsetAt(start);
    html += `<tr class="row-spacer" style="height:${topSpacer}px${topSpacer > 0 ? "" : ";display:none"}"><td colspan="${ALL_COLS.length + 1}"></td></tr>`;
    for (let i = start; i < end; i++) {
      renderedRows.push(i);
      html += rowHtml(i);
    }
    const spacer = Math.max(0, totalGridHeight() - rowOffsetAt(end));
    // 底部占位行始终保留（高度可为 0），供增量渲染调整
    html += `<tr class="row-spacer-b" style="height:${spacer}px${spacer > 0 ? "" : ";display:none"}"><td colspan="${ALL_COLS.length + 1}"></td></tr>`;
    gridBody.innerHTML = html;
    refreshStatus();
    updateSelectionClasses();
  }

  // 增量渲染：快速滚动时不整体重建表格，只裁掉离开窗口的行、追加进入窗口的行。
  // 整体 innerHTML 重建在行数多时耗时明显，跟不上滚动就会露出空白；增量操作只动少数行，
  // 且旧行在被裁掉前一直显示，滚动过程中不再出现空白。
  function ensureWindow() {
    const total = rows.length;
    if (!total) return;
    const st = gridWrap.scrollTop;
    const ch = gridWrap.clientHeight;
    const visStart = rowIndexAt(st);
    const visEnd = Math.min(total, rowIndexAt(st + ch) + 1);
    const curStart = renderedRows.length ? renderedRows[0] : -1;
    const curEnd = renderedRows.length ? renderedRows[renderedRows.length - 1] + 1 : -1;
    if (curStart < 0) { renderWindow(); return; }
    // 迟滞区间：可视区（含半缓冲余量）仍落在已渲染窗口内时不动 DOM，滚动不触发每帧重排
    const half = Math.floor(RENDER_OVERSCAN / 2);
    if (visStart - half >= curStart && visEnd + half <= curEnd) return;
    const wantStart = Math.max(0, visStart - RENDER_OVERSCAN);
    const wantEnd = Math.min(total, visEnd + RENDER_OVERSCAN);
    if (curStart === wantStart && curEnd === wantEnd) return;
    if (wantStart >= curEnd || wantEnd <= curStart) { renderWindow(); return; } // 窗口不相交，整体重建
    const topSpacer = gridBody.firstElementChild;
    const bottomSpacer = gridBody.lastElementChild;
    if (!topSpacer || !bottomSpacer || !topSpacer.classList.contains("row-spacer") || !bottomSpacer.classList.contains("row-spacer-b")) { renderWindow(); return; }
    const trimTop = () => {
      // 顶部数据行紧跟 topSpacer，按序删除即可，避免逐行 querySelector
      let node = topSpacer.nextElementSibling;
      for (let i = curStart; i < wantStart && node && node !== bottomSpacer; i++) {
        const next = node.nextElementSibling;
        node.remove();
        node = next;
      }
    };
    const trimBottom = () => {
      // 底部数据行紧邻 bottomSpacer，从后往前删
      let node = bottomSpacer.previousElementSibling;
      for (let i = curEnd - 1; i >= wantEnd && node && node !== topSpacer; i--) {
        const prev = node.previousElementSibling;
        node.remove();
        node = prev;
      }
    };
    if (wantStart > curStart) {
      trimTop();
      if (wantEnd > curEnd) {
        let html = "";
        for (let i = curEnd; i < wantEnd; i++) html += rowHtml(i);
        bottomSpacer.insertAdjacentHTML("beforebegin", html);
      } else {
        trimBottom();
      }
    } else {
      let html = "";
      for (let i = wantStart; i < curStart; i++) html += rowHtml(i);
      topSpacer.insertAdjacentHTML("afterend", html);
      if (wantEnd < curEnd) trimBottom();
      else {
        let html2 = "";
        for (let i = curEnd; i < wantEnd; i++) html2 += rowHtml(i);
        bottomSpacer.insertAdjacentHTML("beforebegin", html2);
      }
    }
    topSpacer.style.height = rowOffsetAt(wantStart) + "px";
    topSpacer.style.display = wantStart > 0 ? "" : "none";
    const bottomH = Math.max(0, totalGridHeight() - rowOffsetAt(wantEnd));
    bottomSpacer.style.height = bottomH + "px";
    bottomSpacer.style.display = bottomH > 0 ? "" : "none";
    renderedRows = [];
    for (let i = wantStart; i < wantEnd; i++) renderedRows.push(i);
    lastRenderedStart = wantStart;
    updateSelectionClasses();
  }

  gridWrap.addEventListener("scroll", () => {
    if (scrollRenderQueued) return;
    scrollRenderQueued = true;
    requestAnimationFrame(() => {
      scrollRenderQueued = false;
      // 增量渲染：仅在窗口变化时追加/裁剪行，快速滚动不整体重建、不出现空白
      ensureWindow();
    });
  });

  // ---------- 输入联动 ----------
  let recomputeTimer = null;
  // 输入过程中合并重算请求，避免每次按键都全量计算（离散操作仍走同步 recomputeAndRefresh）
  function scheduleRecompute() {
    if (recomputeTimer) clearTimeout(recomputeTimer);
    recomputeTimer = setTimeout(() => {
      recomputeTimer = null;
      recomputeAndRefresh();
    }, 120);
  }

  gridBody.addEventListener("input", (e) => {
    const el = e.target;
    if (el.tagName !== "INPUT") return;
    const tr = el.closest("tr");
    const r = Number(tr.dataset.r);
    const c = el.dataset.c;
    const row = rows[r];
    if (INPUT_COLS.includes(c)) {
      row.input[c] = el.value;
      // 接害因素重新录入时，备注列按数据库重新自动生成
      if (c === "D") {
        delete row.overridden.BI;
        row.values.BI = "";
      }
    } else if (TEXT_OVERRIDE_COLS.includes(c)) {
      row.values[c] = el.value;
      row.overridden[c] = true;
    }
    scheduleRecompute();
  });

  gridBody.addEventListener("change", (e) => {
    const el = e.target;
    if (el.tagName !== "SELECT") return;
    const tr = el.closest("tr");
    const r = Number(tr.dataset.r);
    const c = el.dataset.c;
    const row = rows[r];
    if (OVERRIDE_COLS.includes(c)) {
      row.values[c] = el.value;
      if (el.value === "") delete row.overridden[c];
      else row.overridden[c] = true;
    } else if (MANUAL_COLS.includes(c)) {
      row.manual[c] = el.value;
    } else {
      row.input[c] = el.value;
    }
    scheduleRecompute();
  });

  gridBody.addEventListener("click", (e) => {
    const td = e.target.closest("td.rowno");
    if (!td) return;
    const rRect = td.getBoundingClientRect();
    if (e.clientY >= rRect.bottom - 6) return; // 下边缘为调整行高拖动区，不触发选中整行
    const r = Number(td.closest("tr").dataset.r);
    selectedRow = selectedRow === r ? -1 : r;
    renderWindow();
  });

  function recomputeAndRefresh() {
    if (recomputeTimer) { clearTimeout(recomputeTimer); recomputeTimer = null; }
    L.computeRows(rows, { hazardFactors, detectionItems });
    updateVisibleCells();
    refreshStatus();
  }

  function updateVisibleCells() {
    for (const tr of gridBody.querySelectorAll("tr[data-r]")) {
      const r = Number(tr.dataset.r);
      const row = rows[r];
      if (!row) continue; // 行数变化后 DOM 可能残留旧行
      for (const cell of tr.querySelectorAll("td[data-c]")) {
        const c = cell.dataset.c;
        const isErr = !!row.errors[c];
        cell.classList.toggle("error", isErr);
        cell.title = isErr ? row.errors[c] : "";
        const el = cell.querySelector("input,select");
        if (COMPUTED_COLS.includes(c)) {
          // 只读计算列为纯文本单元格（无 input），直接同步文本内容
          const txt = cell.querySelector("div.computed-text");
          if (txt) {
            const v = escHtml(fmt(row.values[c]));
            if (txt.innerHTML !== v) txt.innerHTML = v;
          }
          continue;
        }
        if (!el) continue;
        let val;
        if (OVERRIDE_COLS.includes(c)) val = fmt(row.values[c]);
        else if (MANUAL_COLS.includes(c)) val = fmt(row.manual[c]);
        else val = fmt(row.input[c]);
        if (el.tagName === "SELECT") {
          if (el.value !== val) el.value = val;
        } else if (el.value !== val) {
          el.value = val;
        }
      }
    }
  }

  function refreshStatus() {
    const { total, byCol } = L.countErrors(rows);
    const status = $("grid-status");
    status.textContent = `共 ${rows.length} 行 · 错误 ${total} 处`;
    status.className = "status" + (total > 0 ? " err" : "");
    if (total > 0) status.title = Object.entries(byCol).map(([c, n]) => `${c}列 ${n}处`).join("，");
    else status.title = "";
  }

  // 错误定位：点击右上角错误计数，按从上到下顺序依次滚动并闪烁到错误单元格；循环点击循环定位
  let errorLocateIdx = -1;
  $("grid-status").addEventListener("click", () => {
    const { total } = L.countErrors(rows);
    if (total === 0) return;
    // 按行、按列序收集错误位置（从上到下、从左到右）
    const errs = [];
    for (let r = 0; r < rows.length; r++) {
      const row = rows[r];
      if (!row || !row.errors) continue;
      for (let ci = 0; ci < ALL_COLS.length; ci++) {
        const c = ALL_COLS[ci];
        if (row.errors[c]) errs.push({ r, c });
      }
    }
    if (!errs.length) return;
    errorLocateIdx = (errorLocateIdx + 1) % errs.length;
    const target = errs[errorLocateIdx];
    const cIdx = ALL_COLS.indexOf(target.c);
    // 虚拟滚动：revealCell 会滚动到目标行并按需重渲染，随后单元格已在 DOM 中
    revealCell(target.r, cIdx);
    requestAnimationFrame(() => {
      const td = findTd(target.r, cIdx);
      if (!td) return;
      td.classList.remove("error-flash");
      void td.offsetWidth; // 强制重启动画
      td.classList.add("error-flash");
      setTimeout(() => td.classList.remove("error-flash"), 1400);
    });
  });

  // ---------- 通用输入弹窗 ----------
  function askInput({ title, hint, value, type }) {
    return new Promise((resolve) => {
      const modal = $("prompt-modal");
      const input = $("prompt-input");
      const okBtn = $("prompt-ok");
      const cancelBtn = $("prompt-cancel");
      $("prompt-title").textContent = title || "输入";
      $("prompt-hint").textContent = hint || "";
      $("prompt-msg").textContent = "";
      input.type = type || "text";
      input.value = value || "";
      modal.classList.remove("hidden");
      input.focus();
      input.select();
      const cleanup = () => {
        modal.classList.add("hidden");
        okBtn.onclick = null;
        cancelBtn.onclick = null;
        input.onkeydown = null;
      };
      okBtn.onclick = () => { const v = input.value.trim(); cleanup(); resolve(v); };
      cancelBtn.onclick = () => { cleanup(); resolve(null); };
      input.onkeydown = (e) => {
        if (e.key === "Enter") { const v = input.value.trim(); cleanup(); resolve(v); }
        if (e.key === "Escape") { cleanup(); resolve(null); }
      };
    });
  }

  async function askRowCount(label, def) {
    const v = await askInput({ title: label, hint: "请输入正整数行数（1 ~ 1000）", value: String(def), type: "number" });
    if (v === null) return null;
    const n = parseInt(v, 10);
    if (!Number.isInteger(n) || n < 1) { alert("请输入大于 0 的整数行数。"); return null; }
    return Math.min(n, 1000);
  }

  function cloneRow(src) {
    const c = blankRow();
    for (const k of INPUT_COLS) c.input[k] = src.input[k] ?? "";
    for (const k of MANUAL_COLS) c.manual[k] = src.manual[k] ?? "";
    c.overridden = { ...(src.overridden || {}) };
    for (const k of OVERRIDE_COLS) c.values[k] = src.values[k] ?? "";
    for (const k of TEXT_OVERRIDE_COLS) c.values[k] = src.values[k] ?? "";
    return c;
  }

  // ---------- 行操作（支持自定义行数） ----------
  $("btn-add").addEventListener("click", async () => {
    const n = await askRowCount("新增行数", 1);
    if (n === null) return;
    if (rows.length + n > 5000) { alert("总行数不能超过 5000。"); return; }
    for (let i = 0; i < n; i++) rows.push(blankRow());
    rowHeights.push(...Array(n).fill(ROW_H));
    rowOffsets = null;
    recomputeAndRefresh();
    gridWrap.scrollTop = gridWrap.scrollHeight;
    renderWindow();
  });
  $("btn-insert").addEventListener("click", async () => {
    const n = await askRowCount("插入行数", 1);
    if (n === null) return;
    if (rows.length + n > 5000) { alert("总行数不能超过 5000。"); return; }
    // 优先用行号选中的行；没有时回退到当前选中单元格所在行
    const anchor = selectedRow >= 0 ? selectedRow : cur ? cur.r : -1;
    const at = anchor >= 0 ? anchor + 1 : rows.length;
    const ins = [];
    for (let i = 0; i < n; i++) ins.push(blankRow());
    rows.splice(at, 0, ...ins);
    rowHeights.splice(at, 0, ...Array(n).fill(ROW_H));
    rowOffsets = null;
    if (cur && cur.r >= at) cur.r += n;
    selectedRow = at;
    recomputeAndRefresh();
    renderWindow();
  });
  $("btn-copy").addEventListener("click", async () => {
    // 存在多格选区时，以选区首行为基准、默认复制选区覆盖的行数
    const rect = selRect();
    const selRows = rect ? rect.r2 - rect.r1 + 1 : 1;
    const n = await askRowCount("复制行数", selRows);
    if (n === null) return;
    // 优先用行号选中的行；没有时回退到当前选中单元格所在行，有选区时用选区首行
    const at = rect ? rect.r1 : selectedRow >= 0 ? selectedRow : cur ? cur.r : -1;
    if (at < 0) return;
    if (rows.length + n > 5000) { alert("总行数不能超过 5000。"); return; }
    const src = rows[at];
    const copies = [];
    for (let i = 0; i < n; i++) copies.push(cloneRow(src));
    rows.splice(at + 1, 0, ...copies);
    rowHeights.splice(at + 1, 0, ...Array(n).fill(ROW_H));
    rowOffsets = null;
    if (cur && cur.r > at) cur.r += n;
    selectedRow = at + n;
    recomputeAndRefresh();
    renderWindow();
  });
  // 自动调整：点击一次按内容统一重算列宽与行高（列宽按内容自适应，行高恢复为内容所需高度）
  $("btn-fit-rows").addEventListener("click", () => {
    colWidths = computeGridWidths();
    applyGridWidths();
    for (let i = 0; i < rows.length; i++) rowHeights[i] = ROW_H;
    rowOffsets = null;
    recomputeAndRefresh();
    renderWindow();
  });
  $("btn-del").addEventListener("click", async () => {
    // 存在多格选区时，以选区首行起删、默认删除选区覆盖的行数
    const rect = selRect();
    const selRows = rect ? rect.r2 - rect.r1 + 1 : 1;
    // 优先用行号选中的行；没有时回退到当前选中单元格所在行，有选区时用选区首行
    const delAt = rect ? rect.r1 : selectedRow >= 0 ? selectedRow : cur ? cur.r : -1;
    if (delAt < 0) { alert("请先选中要删除的行（点击行号或点击任意单元格）。"); return; }
    const n = await askRowCount("删除行数", selRows);
    if (n === null) return;
    const cnt = Math.min(n, rows.length - delAt);
    if (cnt < 1) return;
    if (!confirm(`确定删除从第 ${delAt + 1} 行起的 ${cnt} 行？`)) return;
    rows.splice(delAt, cnt);
    rowHeights.splice(delAt, cnt);
    rowOffsets = null;
    selectedRow = -1;
    clampCur();
    recomputeAndRefresh();
    renderWindow();
  });
  $("btn-reset").addEventListener("click", () => {
    if (!confirm("清空测点表格（保留危害因素库与检测项目）？")) return;
    emptyGrid();
    selectedRow = -1;
    selStart = selEnd = selAnchor = null;
    cur = null;
    recomputeAndRefresh();
    renderWindow();
  });

  // ---------- 录入区：Excel 式编辑（单元格光标 / 键盘导航 / 区域选择 / 复制粘贴） ----------
  let selStart = null;
  let selEnd = null;
  let selAnchor = null;
  let selDragging = false;
  let selFrame = false;
  let cur = null;       // 当前单元格 { r, c }，c 为 ALL_COLS 索引
  let editing = false;  // 是否处于单元格编辑模式（F2 / 双击 / 直接输入）
  let editOriginal = null; // 进入编辑时的原始值，Esc 还原用
  let lastMouse = null; // 最近一次鼠标位置，双击时用于放置光标

  function editableCellAt(r, cIdx) {
    if (cIdx < 0 || cIdx >= ALL_COLS.length) return false;
    const col = ALL_COLS[cIdx];
    return INPUT_COLS.includes(col) || MANUAL_COLS.includes(col) || TEXT_OVERRIDE_COLS.includes(col);
  }

  function findTd(r, cIdx) {
    return gridBody.querySelector(`tr[data-r="${r}"] > td[data-c="${ALL_COLS[cIdx]}"]`);
  }

  function getCellModelValue(r, cIdx) {
    if (!rows[r]) return "";
    const col = ALL_COLS[cIdx];
    if (INPUT_COLS.includes(col)) return rows[r].input[col] ?? "";
    if (MANUAL_COLS.includes(col)) return rows[r].manual[col] ?? "";
    return rows[r].values[col] ?? "";
  }

  function selRect() {
    if (!selStart || !selEnd) return null;
    return {
      r1: Math.min(selStart.r, selEnd.r),
      r2: Math.max(selStart.r, selEnd.r),
      c1: Math.min(selStart.c, selEnd.c),
      c2: Math.max(selStart.c, selEnd.c),
    };
  }

  // 开始单元格选区时，取消整行选中（避免整行浅蓝高亮残留，看起来像同行全部被选中）
  function clearRowSelection() {
    if (selectedRow < 0) return;
    const tr = gridBody.querySelector(`tr[data-r="${selectedRow}"]`);
    if (tr) {
      tr.classList.remove("selected");
      const rowno = tr.querySelector("td.rowno");
      if (rowno) rowno.classList.remove("selected");
    }
    selectedRow = -1;
  }

  function updateSelectionClasses() {
    const rect = selRect();
    // 单击也会产生 1×1 选区；仅当选区多于一个单元格时才视为“拖选”，让行列高亮让位
    const multi = rect && (rect.r1 !== rect.r2 || rect.c1 !== rect.c2);
    for (const tr of gridBody.querySelectorAll("tr[data-r]")) {
      const r = Number(tr.dataset.r);
      const isCurRow = !!(cur && r === cur.r); // 当前单元格所在行高亮
      for (const td of tr.querySelectorAll("td[data-c]")) {
        const c = ALL_COLS.indexOf(td.dataset.c);
        const inSel = rect && r >= rect.r1 && r <= rect.r2 && c >= rect.c1 && c <= rect.c2;
        td.classList.toggle("sel", inSel);
        td.classList.toggle("cur", !!(cur && r === cur.r && c === cur.c));
        // 选中单元格的整行、整列高亮（拖选多格时以选区代替，避免大片高亮）
        td.classList.toggle("hl-row", !multi && isCurRow && c !== cur.c);
        td.classList.toggle("hl-col", !multi && !!cur && c === cur.c && r !== cur.r);
      }
    }
  }

  // 把当前单元格滚动到可见区域（虚拟滚动下按需重渲染）
  function revealCell(r, cIdx) {
    const wrap = gridWrap;
    const headH = gridHead.offsetHeight || 60; // sticky 表头高度
    const rowTop = headH + rowOffsetAt(r);
    const ch = wrap.clientHeight;
    if (rowTop < wrap.scrollTop + headH) wrap.scrollTop = Math.max(0, rowTop - headH);
    else if (rowTop + rowHeightAt(r) > wrap.scrollTop + ch) wrap.scrollTop = rowTop + rowHeightAt(r) - ch;
    if (!renderedRows.includes(r)) renderWindow();
    const td = findTd(r, cIdx);
    if (!td) return;
    const wrapRect = wrap.getBoundingClientRect();
    const tdRect = td.getBoundingClientRect();
    const left = tdRect.left - wrapRect.left + wrap.scrollLeft;
    const right = tdRect.right - wrapRect.left + wrap.scrollLeft;
    if (left < wrap.scrollLeft) wrap.scrollLeft = left;
    else if (right > wrap.scrollLeft + wrap.clientWidth) wrap.scrollLeft = right - wrap.clientWidth;
    if (!renderedRows.includes(r)) renderWindow();
    updateSelectionClasses();
  }

  function focusCurrentCell() {
    if (!cur) return;
    const td = findTd(cur.r, cur.c);
    if (!td) return;
    const el = td.querySelector("input,select");
    if (!el) return;
    el.focus({ preventScroll: true });
    if (el.tagName === "INPUT") {
      try { el.setSelectionRange(el.value.length, el.value.length); } catch {}
    }
  }

  // 移动/设置当前单元格；extend=true 表示 Shift+方向键扩展选区
  function setCur(r, cIdx, opts = {}) {
    if (rows.length === 0) return;
    const { extend = false, focus = true } = opts;
    clearRowSelection();
    r = Math.max(0, Math.min(r, rows.length - 1));
    cIdx = Math.max(0, Math.min(cIdx, ALL_COLS.length - 1));
    const prev = cur;
    cur = { r, c: cIdx };
    editing = false;
    editOriginal = null;
    if (extend && (selAnchor || prev)) {
      if (!selAnchor && prev) selAnchor = prev;
      selStart = selAnchor;
      selEnd = { r, c: cIdx };
    } else {
      selAnchor = { r, c: cIdx };
      selStart = { r, c: cIdx };
      selEnd = { r, c: cIdx };
    }
    revealCell(r, cIdx);
    if (focus) focusCurrentCell();
  }

  function moveCur(dr, dc, opts = {}) {
    if (!cur || rows.length === 0) return;
    let r = cur.r + dr;
    let c = cur.c + dc;
    if (opts.wrap) {
      while (c >= ALL_COLS.length) { c -= ALL_COLS.length; r += 1; }
      while (c < 0) { c += ALL_COLS.length; r -= 1; }
      if (r < 0) { r = 0; c = 0; }
    }
    if (opts.grow && r > rows.length - 1) {
      if (rows.length >= 5000) r = rows.length - 1;
      else while (rows.length < 5000 && r > rows.length - 1) { rows.push(blankRow()); rowHeights.push(ROW_H); }
    }
    rowOffsets = null;
    setCur(r, c, opts);
  }

  function commitCurrent() {
    if (!cur) return;
    const td = findTd(cur.r, cur.c);
    if (!td) return;
    const el = td.querySelector("input,select");
    if (!el) return;
    const col = ALL_COLS[cur.c];
    const row = rows[cur.r];
    if (el.tagName === "SELECT") {
      if (OVERRIDE_COLS.includes(col)) {
        row.values[col] = el.value;
        if (el.value === "") delete row.overridden[col];
        else row.overridden[col] = true;
      } else if (MANUAL_COLS.includes(col)) {
        row.manual[col] = el.value;
      } else {
        row.input[col] = el.value;
      }
    } else if (TEXT_OVERRIDE_COLS.includes(col)) {
      row.values[col] = el.value;
      row.overridden[col] = true;
    } else if (INPUT_COLS.includes(col)) {
      row.input[col] = el.value;
    } else if (MANUAL_COLS.includes(col)) {
      row.manual[col] = el.value;
    }
    editing = false;
    recomputeAndRefresh();
    editOriginal = null;
  }

  function revertCurrent() {
    if (!cur) return;
    const restore = editOriginal === null ? getCellModelValue(cur.r, cur.c) : editOriginal;
    const col = ALL_COLS[cur.c];
    const row = rows[cur.r];
    if (INPUT_COLS.includes(col)) row.input[col] = restore;
    else if (MANUAL_COLS.includes(col)) row.manual[col] = restore;
    else if (TEXT_OVERRIDE_COLS.includes(col)) row.values[col] = restore;
    const td = findTd(cur.r, cur.c);
    if (td) {
      const el = td.querySelector("input");
      if (el) el.value = restore;
    }
    editing = false;
    editOriginal = null;
    updateVisibleCells();
  }

  // 清空选区内的可编辑内容（Delete/Backspace）
  function clearRange(rect) {
    if (!rect) return false;
    let cleared = false;
    for (let rr = rect.r1; rr <= Math.min(rect.r2, rows.length - 1); rr++) {
      const row = rows[rr];
      for (let cc = rect.c1; cc <= rect.c2; cc++) {
        const col = ALL_COLS[cc];
        if (INPUT_COLS.includes(col)) {
          if (row.input[col] !== "") { row.input[col] = ""; cleared = true; }
        } else if (MANUAL_COLS.includes(col)) {
          if (row.manual[col] !== "") { row.manual[col] = ""; cleared = true; }
        } else if (OVERRIDE_COLS.includes(col)) {
          if (row.values[col] !== "" || row.overridden[col]) {
            row.values[col] = "";
            delete row.overridden[col];
            cleared = true;
          }
        } else if (TEXT_OVERRIDE_COLS.includes(col)) {
          if (row.values[col] !== "" || row.overridden[col]) {
            row.values[col] = "";
            row.overridden[col] = true;
            cleared = true;
          }
        }
      }
    }
    return cleared;
  }

  function clampCur() {
    if (!cur) return;
    if (rows.length === 0) { cur = null; return; }
    cur.r = Math.max(0, Math.min(cur.r, rows.length - 1));
    cur.c = Math.max(0, Math.min(cur.c, ALL_COLS.length - 1));
  }

  function lastDataRow() {
    for (let r = rows.length - 1; r >= 0; r--) {
      const row = rows[r];
      const has =
        INPUT_COLS.some((col) => (row.input[col] ?? "") !== "") ||
        MANUAL_COLS.some((col) => (row.manual[col] ?? "") !== "");
      if (has) return r;
    }
    return 0;
  }

  function lastDataCol(rowIdx) {
    const r = rowIdx === undefined ? (cur ? cur.r : 0) : rowIdx;
    const row = rows[r];
    if (!row) return INPUT_COLS.length - 1;
    for (let c = ALL_COLS.length - 1; c >= 0; c--) {
      const col = ALL_COLS[c];
      if (INPUT_COLS.includes(col) && (row.input[col] ?? "") !== "") return c;
      if (MANUAL_COLS.includes(col) && (row.manual[col] ?? "") !== "") return c;
      if ((row.values[col] ?? "") !== "") return c;
    }
    return INPUT_COLS.length - 1;
  }

  // Ctrl+A 全选整张表（当前单元格保留为锚点）
  function selectAllCells() {
    if (rows.length === 0) return;
    const anchor = cur || { r: 0, c: 0 };
    selAnchor = anchor;
    selStart = { r: 0, c: 0 };
    selEnd = { r: rows.length - 1, c: ALL_COLS.length - 1 };
    updateSelectionClasses();
  }

  // ---------- 行高拖动调整：按住行号列下边缘上下拖动（类似 Excel） ----------
  let resizingRow = null; // { r, startY, startH }

  gridBody.addEventListener("mousedown", (e) => {
    const rno = e.target.closest("td.rowno");
    if (!rno || e.target.closest("input,select")) return;
    const rRect = rno.getBoundingClientRect();
    if (e.clientY >= rRect.bottom - 6) {
      const r = Number(rno.closest("tr[data-r]").dataset.r);
      if (r >= 0 && r < rows.length) {
        resizingRow = { r, startY: e.clientY, startH: rowHeightAt(r) };
        e.preventDefault();
      }
    }
  });

  document.addEventListener("mousemove", (e) => {
    if (!resizingRow) return;
    const h = Math.max(MIN_ROW_H, Math.min(MAX_ROW_H, Math.round(resizingRow.startH + (e.clientY - resizingRow.startY))));
    if (h !== rowHeights[resizingRow.r]) {
      rowHeights[resizingRow.r] = h;
      rowOffsets = null;
      renderWindow();
    }
  });

  document.addEventListener("mouseup", () => {
    resizingRow = null;
  });

  // ---------- 列宽拖动调整：按住表头列右边缘左右拖动（类似 Excel） ----------
  let resizingCol = null; // { idx, startX, startW }

  gridHead.addEventListener("mousedown", (e) => {
    if (e.button !== 0) return;
    for (const th of gridHead.querySelectorAll("th.field")) {
      const idx = th.dataset.c !== undefined ? colIdx(th.dataset.c) : -1;
      if (idx < 0) continue;
      const r = th.getBoundingClientRect();
      if (Math.abs(e.clientX - r.right) <= 6) {
        resizingCol = { idx, startX: e.clientX, startW: (colWidths || computeGridWidths())[idx] };
        e.preventDefault();
        return;
      }
    }
  });

  document.addEventListener("mousemove", (e) => {
    if (!resizingCol) return;
    const w = Math.max(MIN_COL_W, Math.min(MAX_COL_W, Math.round(resizingCol.startW + (e.clientX - resizingCol.startX))));
    if (w !== colWidths[resizingCol.idx]) {
      colWidths[resizingCol.idx] = w;
      applyGridWidths();
    }
  });

  document.addEventListener("mouseup", () => {
    resizingCol = null;
  });

  gridBody.addEventListener("mousedown", (e) => {
    const td = e.target.closest("td[data-c]");
    if (!td) return;
    const r = Number(td.closest("tr").dataset.r);
    const c = ALL_COLS.indexOf(td.dataset.c);
    const cell = { r, c };
    lastMouse = { x: e.clientX, y: e.clientY };
    clearRowSelection();
    editing = false;
    editOriginal = null;
    if (e.shiftKey) {
      // Shift+点击扩展选区：当前单元格保持在锚点（Excel 行为）
      if (!cur) selAnchor = cell;
      else if (!selAnchor) selAnchor = cur;
      selStart = selAnchor;
      selEnd = cell;
    } else {
      selAnchor = cell;
      selStart = cell;
      selEnd = cell;
      cur = { r, c };
    }
    selDragging = true;
    updateSelectionClasses();
  });

  gridBody.addEventListener("mouseover", (e) => {
    if (!selDragging) return;
    const td = e.target.closest("td[data-c]");
    if (!td) return;
    selEnd = { r: Number(td.closest("tr").dataset.r), c: ALL_COLS.indexOf(td.dataset.c) };
    if (!selFrame) {
      selFrame = true;
      requestAnimationFrame(() => {
        selFrame = false;
        updateSelectionClasses();
      });
    }
  });

  document.addEventListener("mouseup", () => {
    selDragging = false;
  });

  gridBody.addEventListener("dblclick", (e) => {
    const td = e.target.closest("td[data-c]");
    if (!td) return;
    const r = Number(td.closest("tr[data-r]").dataset.r);
    const c = ALL_COLS.indexOf(td.dataset.c);
    cur = { r, c };
    selAnchor = selStart = selEnd = { r, c };
    if (editableCellAt(r, c)) {
      editOriginal = getCellModelValue(r, c);
      editing = true;
      const el = td.querySelector("input");
      if (el) {
        el.focus({ preventScroll: true });
        let pos = el.value.length;
        if (lastMouse && document.caretRangeFromPoint) {
          const range = document.caretRangeFromPoint(lastMouse.x, lastMouse.y);
          if (range && range.startContainer === el) pos = range.startOffset;
        }
        try { el.setSelectionRange(pos, pos); } catch {}
      }
    }
    updateSelectionClasses();
  });

  gridBody.addEventListener("focusout", () => {
    editing = false;
  });

  gridBody.addEventListener("keydown", (e) => {
    const td = e.target && e.target.closest ? e.target.closest("td[data-c]") : null;
    if (!td) return; // 行号列、工具栏等不参与
    const tr = td.closest("tr[data-r]");
    if (!tr) return;
    const r = Number(tr.dataset.r);
    const c = ALL_COLS.indexOf(td.dataset.c);
    const key = e.key;
    const isInput = e.target.tagName === "INPUT";
    const isSelect = e.target.tagName === "SELECT";
    const mod = e.ctrlKey || e.metaKey;
    const shift = e.shiftKey;
    if (e.isComposing || key === "Process") return; // 中文输入法组合中不拦截

    if (mod) {
      if (editing) return; // 编辑中保留浏览器原生快捷键
      if (key === "Home") { e.preventDefault(); setCur(0, 0); return; }
      if (key === "End") {
        e.preventDefault();
        const lastR = lastDataRow();
        setCur(lastR, lastDataCol(lastR));
        return;
      }
      if (key.toLowerCase() === "a") {
        e.preventDefault();
        selectAllCells();
        return;
      }
      if (["c", "v", "x"].includes(key.toLowerCase())) return; // 原生复制/粘贴/剪切
      return;
    }

    // F2：进入编辑模式
    if (key === "F2") {
      e.preventDefault();
      if (isInput && editableCellAt(r, c)) {
        editOriginal = getCellModelValue(r, c);
        editing = true;
        try { e.target.setSelectionRange(e.target.value.length, e.target.value.length); } catch {}
      }
      return;
    }

    // 编辑模式下：Enter/Tab 提交并移动，Escape 还原，其余按键交给输入框
    if (editing) {
      if (key === "Enter") {
        e.preventDefault();
        commitCurrent();
        moveCur(shift ? -1 : 1, 0, { grow: true, select: false });
        return;
      }
      if (key === "Tab") {
        e.preventDefault();
        commitCurrent();
        moveCur(0, shift ? -1 : 1, { grow: true, select: false, wrap: true });
        return;
      }
      if (key === "Escape") {
        e.preventDefault();
        revertCurrent();
        return;
      }
      return; // 方向键 / Home / End / 退格等在编辑时走原生行为
    }

    // 非编辑模式：单元格光标移动
    if (key === "Enter") {
      e.preventDefault();
      moveCur(shift ? -1 : 1, 0, { grow: true });
      return;
    }
    if (key === "Tab") {
      e.preventDefault();
      moveCur(0, shift ? -1 : 1, { grow: true, wrap: true });
      return;
    }
    if (key === "Escape") {
      selStart = selEnd = selAnchor = null;
      updateSelectionClasses();
      return;
    }
    if (key === "ArrowDown") {
      if (isSelect && e.altKey) return; // Alt+↓ 打开下拉框
      e.preventDefault();
      moveCur(1, 0, { extend: shift });
      return;
    }
    if (key === "ArrowUp") {
      e.preventDefault();
      moveCur(-1, 0, { extend: shift });
      return;
    }
    if (key === "ArrowRight") {
      e.preventDefault();
      moveCur(0, 1, { extend: shift });
      return;
    }
    if (key === "ArrowLeft") {
      e.preventDefault();
      moveCur(0, -1, { extend: shift });
      return;
    }
    if (key === "Home") {
      e.preventDefault();
      setCur(r, 0);
      return;
    }
    if (key === "End") {
      e.preventDefault();
      setCur(r, lastDataCol());
      return;
    }
    if (key === "PageDown") {
      e.preventDefault();
      const rh = Math.max(1, rowHeightAt(cur ? cur.r : 0));
      const page = Math.max(1, Math.floor(gridWrap.clientHeight / rh) - 1);
      moveCur(page, 0, { extend: shift });
      return;
    }
    if (key === "PageUp") {
      e.preventDefault();
      const rh = Math.max(1, rowHeightAt(cur ? cur.r : 0));
      const page = Math.max(1, Math.floor(gridWrap.clientHeight / rh) - 1);
      moveCur(-page, 0, { extend: shift });
      return;
    }
    if (key === "Backspace") {
      // 文本单元格：进入编辑状态并逐字删除（不清空整格），与录入区编辑一致
      const el = e.target;
      if (el && el.tagName === "INPUT" && editableCellAt(r, c)) {
        e.preventDefault();
        editing = true;
        if (editOriginal === null) editOriginal = getCellModelValue(r, c);
        const start = el.selectionStart ?? el.value.length;
        const end = el.selectionEnd ?? el.value.length;
        let caret;
        if (start !== end) {
          el.value = el.value.slice(0, start) + el.value.slice(end);
          caret = start;
        } else {
          el.value = el.value.slice(0, Math.max(0, start - 1)) + el.value.slice(start);
          caret = Math.max(0, start - 1);
        }
        el.dispatchEvent(new Event("input", { bubbles: true }));
        try { el.setSelectionRange(caret, caret); } catch {}
        return;
      }
      if (clearRange(selRect())) {
        e.preventDefault();
        editing = false;
        editOriginal = null;
        recomputeAndRefresh();
      }
      return;
    }
    if (key === "Delete") {
      if (clearRange(selRect())) {
        e.preventDefault();
        editing = false;
        editOriginal = null;
        recomputeAndRefresh();
      }
      return;
    }
    // 下拉框输入首字符快速选中（Excel 行为）
    if (isSelect && key.length === 1 && !shift) {
      const opts = Array.from(e.target.options).filter((o) => o.value && o.value !== "");
      const hit = opts.find((o) => o.text.trim().toLowerCase().startsWith(key.toLowerCase()));
      if (hit) {
        e.preventDefault();
        e.target.value = hit.value;
        e.target.dispatchEvent(new Event("change", { bubbles: true }));
      }
      return;
    }
    // 直接输入字符：在光标处插入（有选中区域则替换选中内容），不清空原有内容
    if (isInput && editableCellAt(r, c) && key.length === 1) {
      e.preventDefault();
      const el = e.target;
      editOriginal = getCellModelValue(r, c);
      const start = el.selectionStart ?? el.value.length;
      const end = el.selectionEnd ?? el.value.length;
      el.value = el.value.slice(0, start) + key + el.value.slice(end);
      editing = true;
      el.dispatchEvent(new Event("input", { bubbles: true }));
      try { el.setSelectionRange(start + key.length, start + key.length); } catch {}
    }
  });

  gridBody.addEventListener("copy", (e) => {
    const rect = selRect();
    if (!rect) return;
    const lines = [];
    for (let r = rect.r1; r <= rect.r2; r++) {
      const cells = [];
      for (let c = rect.c1; c <= rect.c2; c++) {
        const col = ALL_COLS[c];
        let v = "";
        if (r < rows.length) {
          if (INPUT_COLS.includes(col)) v = rows[r].input[col];
          else if (MANUAL_COLS.includes(col)) v = rows[r].manual[col];
          else v = rows[r].values[col];
        }
        cells.push(v === null || v === undefined ? "" : String(v));
      }
      lines.push(cells.join("\t"));
    }
    if (e.clipboardData) {
      e.clipboardData.setData("text/plain", lines.join("\n"));
      e.preventDefault();
    }
  });

  gridBody.addEventListener("paste", (e) => {
    const text = e.clipboardData ? e.clipboardData.getData("text/plain") : "";
    if (!text) return;
    const lines = String(text).replace(/\r\n/g, "\n").replace(/\r/g, "\n").split("\n");
    // 去掉末尾空行（复制内容末尾常带换行）
    while (lines.length && lines[lines.length - 1] === "") lines.pop();
    const srcRows = lines.map((l) => l.split("\t"));
    const srcH = Math.max(srcRows.length, 1);
    const srcW = srcRows.reduce((m, r) => Math.max(m, r.length), 1);
    // 未指定选中区域（无选区或仅当前单元格）时，从当前单元格开始，按剪贴板全部内容大小直接粘贴
    let rect = selRect();
    if (!rect || (rect.r1 === rect.r2 && rect.c1 === rect.c2)) {
      const r0 = cur ? cur.r : 0;
      const c0 = cur ? cur.c : 0;
      rect = { r1: r0, c1: c0, r2: r0 + srcH - 1, c2: c0 + srcW - 1 };
    }
    let changed = false;
    // 将复制内容按行列规律重复填充到整个选中区域（复制单值则整片填入）
    for (let r = rect.r1; r <= rect.r2; r++) {
      for (let c = rect.c1; c <= rect.c2; c++) {
        if (c >= ALL_COLS.length) break;
        const col = ALL_COLS[c];
        if (!INPUT_COLS.includes(col) && !TEXT_OVERRIDE_COLS.includes(col)) continue;
        const val = (srcRows[(r - rect.r1) % srcH] || [])[(c - rect.c1) % srcW] ?? "";
        while (rows.length <= r) { rows.push(blankRow()); rowHeights.push(ROW_H); }
        rowOffsets = null;
        if (INPUT_COLS.includes(col)) {
          rows[r].input[col] = val;
          // 接害因素粘贴变化时，备注列按数据库重新自动生成
          if (col === "D") {
            delete rows[r].overridden.BI;
            rows[r].values.BI = "";
          }
        } else {
          rows[r].values[col] = val;
          rows[r].overridden[col] = true;
        }
        changed = true;
      }
    }
    if (changed) {
      e.preventDefault();
      cur = { r: rect.r1, c: rect.c1 };
      selAnchor = { r: rect.r1, c: rect.c1 };
      recomputeAndRefresh();
      renderWindow();
      // 粘贴后恢复焦点到区域左上角，保证可立即用 Delete/方向键继续操作
      revealCell(rect.r1, rect.c1);
      updateSelectionClasses();
      focusCurrentCell();
    }
  });

  // ---------- 一键清空录入区 ----------
  $("btn-clear-input").addEventListener("click", () => {
    if (!confirm("确定清空录入区全部内容（A~U 录入列与手工填写列）？行结构保留。")) return;
    for (const r of rows) {
      for (const c of INPUT_COLS) r.input[c] = "";
      for (const c of MANUAL_COLS) r.manual[c] = "";
      r.overridden = {};
      for (const c of OVERRIDE_COLS) r.values[c] = "";
      for (const c of TEXT_OVERRIDE_COLS) { r.values[c] = ""; delete r.overridden[c]; }
    }
    selStart = selEnd = selAnchor = null;
    cur = null;
    recomputeAndRefresh();
    renderWindow();
  });

  // ---------- 全屏显示：主表铺满浏览器窗口，操作按钮浮动在顶部 ----------
  const FS_BTN = $("btn-fs");
  const FS_BAR = $("fs-bar");

  function isFsActive() {
    return !!(document.fullscreenElement || document.webkitFullscreenElement);
  }

  function enterFullscreen() {
    const el = document.documentElement;
    const req = el.requestFullscreen || el.webkitRequestFullscreen;
    if (req) req.call(el);
    else {
      // 浏览器不支持 Fullscreen API 时降级为铺满视口
      document.body.classList.add("fullscreen-mode");
      FS_BAR.classList.remove("hidden");
      FS_BTN.textContent = "退出全屏";
    }
  }

  function exitFullscreen() {
    const ex = document.exitFullscreen || document.webkitExitFullscreen;
    if (ex && (document.fullscreenElement || document.webkitFullscreenElement)) ex.call(document);
    else {
      document.body.classList.remove("fullscreen-mode");
      FS_BAR.classList.add("hidden");
      FS_BTN.textContent = "⛶ 全屏";
    }
  }

  function syncFsUi() {
    const fs = isFsActive();
    document.body.classList.toggle("fullscreen-mode", fs);
    FS_BAR.classList.toggle("hidden", !fs);
    FS_BTN.textContent = fs ? "退出全屏" : "⛶ 全屏";
  }

  FS_BTN.addEventListener("click", () => {
    if (isFsActive()) exitFullscreen();
    else enterFullscreen();
  });
  document.addEventListener("fullscreenchange", syncFsUi);
  document.addEventListener("webkitfullscreenchange", syncFsUi);
  document.addEventListener("keydown", (e) => {
    if (e.key === "Escape" && isFsActive() && !document.querySelector(".modal:not(.hidden)")) {
      exitFullscreen();
    }
  });

  // 浮动操作条按钮转发到顶部同名按钮，复用原有逻辑
  const fsActions = {
    "fs-import": "btn-import",
    "fs-export": "btn-export",
    "fs-upload": "btn-upload",
    "fs-save": "btn-save",
    "fs-db": "btn-db",
    "fs-gh": "btn-gh",
    "fs-reset": "btn-reset",
  };
  for (const [fsId, origId] of Object.entries(fsActions)) {
    $(fsId).addEventListener("click", () => $(origId).click());
  }
  $("fs-exit").addEventListener("click", exitFullscreen);

  // ---------- 数据记录与参考库存储：本地服务 / GitHub API / 临时模式 ----------
  const RECORDS_KEY = "samplingPlanRecords_v1"; // 仅用于“临时模式”兜底
  const LIBRARY_KEY = "samplingPlanLibrary_v1";
  const GH_CONFIG_KEY = "samplingPlanGithubConfig_v1";
  const GH_PATH_DEFAULT = "sampling-plan-app/data/records.json";
  const GH_LIB_PATH_DEFAULT = "sampling-plan-app/data/library.json";
  let storageMode = "detecting";
  let ghSha = null;
  let libSha = null;
  let libSyncTimer = null;

  function loadGithubConfig() {
    try {
      return JSON.parse(localStorage.getItem(GH_CONFIG_KEY)) || null;
    } catch {
      return null;
    }
  }
  function saveGithubConfig(cfg) {
    try {
      localStorage.setItem(GH_CONFIG_KEY, JSON.stringify(cfg));
      return true;
    } catch {
      return false;
    }
  }
  function clearGithubConfig() {
    try {
      localStorage.removeItem(GH_CONFIG_KEY);
    } catch {}
  }

  function setStorageNotice(show, text) {
    const el = $("storage-notice");
    if (!el) return;
    if (text) el.innerHTML = text;
    el.classList.toggle("hidden", !show);
  }

  function setSyncStatus(state, msg) {
    const el = $("sync-status");
    if (!el) return;
    const cfg = loadGithubConfig();
    if (!cfg || !cfg.repo) {
      el.classList.add("hidden");
      return;
    }
    el.classList.remove("hidden");
    const t = new Date().toLocaleTimeString("zh-CN", { hour12: false });
    if (state === "ok") {
      el.className = "sync-status ok";
      el.innerHTML = `✓ 已从 GitHub 同步（${t}）：<b>${escHtml(cfg.repo)}</b>，保存的数据将直接写回 GitHub。`;
    } else if (state === "err") {
      el.className = "sync-status err";
      el.innerHTML = `✕ GitHub 同步失败（${t}）：${escHtml(msg || "未知错误")}，本次使用本地数据。`;
    } else {
      el.className = "sync-status syncing";
      el.innerHTML = `⟳ 正在从 GitHub 同步…（${escHtml(cfg.repo)}）`;
    }
  }

  function hasGithubConfig() {
    const cfg = loadGithubConfig();
    return !!(cfg && cfg.repo && cfg.token);
  }

  // 将远端索引合并到本地镜像（保留已缓存的完整记录，作为离线兜底）
  function mirrorToLocal(index) {
    try {
      const cached = new Map(readLocalFull().map((r) => [r.id, r]));
      const merged = index.map((r) => cached.get(r.id) || r);
      localStorage.setItem(RECORDS_KEY, JSON.stringify(merged));
    } catch {}
  }

  function noticeHtml() {
    if (storageMode === "server") return "";
    if (storageMode === "github") {
      return "当前为静态网页模式，数据通过 GitHub API 保存到仓库中的 <b>data/records.json</b>。";
    }
    if (storageMode === "static-unconfigured") {
      return '当前为静态网页模式且未配置 GitHub：请在「<b>GitHub 配置</b>」中填写仓库与 Token，保存的数据将直接写入 GitHub 仓库的 <b>sampling-plan-app/data/records.json</b>；未配置时数据仅暂存浏览器。';
    }
    return '当前为「直接打开页面」模式，数据仅暂存于浏览器。请通过「<b>启动采样计划软件.bat</b>」打开软件，数据将保存到项目文件夹 <b>data\\records.json</b>。';
  }

  async function detectStorageMode() {
    if (window.location.protocol === "http:" || window.location.protocol === "https:") {
      try {
        const res = await fetch("/api/health", { cache: "no-store" });
        if (res.ok) {
          storageMode = "server";
          return;
        }
      } catch {}
      const cfg = loadGithubConfig();
      storageMode = cfg && cfg.repo && cfg.token ? "github" : "static-unconfigured";
      return;
    }
    storageMode = "temp";
  }

  async function ensureMode() {
    if (storageMode === "detecting") await detectStorageMode();
  }

  // 新存储结构：远端只存「索引」（不含 rows），单条记录按 id 拉取/保存；
  // localStorage 镜像保存完整记录，作为离线兜底与内容预览缓存
  function readLocalFull() {
    try {
      const s = localStorage.getItem(RECORDS_KEY);
      const v = s ? JSON.parse(s) : [];
      return Array.isArray(v) ? v : [];
    } catch {
      return [];
    }
  }

  function writeLocalFull(list) {
    try { localStorage.setItem(RECORDS_KEY, JSON.stringify(list)); } catch {}
  }

  // 索引 + 本地缓存合并：有缓存的记录带 rows（供预览与直接调用），没有的仅元信息
  function attachCachedRows(index) {
    const map = new Map(readLocalFull().filter((r) => r && r.rows).map((r) => [r.id, r]));
    return index.map((r) => {
      const c = map.get(r.id);
      return c ? { ...r, rows: c.rows } : r;
    });
  }

  function upsertLocalFull(rec) {
    const list = readLocalFull().filter((r) => r.id !== rec.id);
    list.unshift(rec);
    writeLocalFull(list);
  }

  function removeFromLocalFull(id) {
    writeLocalFull(readLocalFull().filter((r) => r.id !== id));
  }

  async function loadRecords() {
    await ensureMode();
    // 配置了 GitHub：优先从 GitHub 读取索引，实现“打开即同步”
    if (hasGithubConfig()) {
      setSyncStatus("syncing");
      try {
        const index = await githubLoadIndex();
        mirrorToLocal(index);
        setSyncStatus("ok");
        return attachCachedRows(index);
      } catch (e) {
        setSyncStatus("err", e.message);
      }
    }
    if (storageMode === "server") {
      try {
        const res = await fetch("/api/records", { cache: "no-store" });
        if (!res.ok) throw new Error("HTTP " + res.status);
        const index = await res.json();
        return attachCachedRows(Array.isArray(index) ? index : []);
      } catch {
        setStorageNotice(true);
        return localLoad();
      }
    }
    if (storageMode === "github") {
      try {
        const index = await githubLoadIndex();
        return attachCachedRows(index);
      } catch (e) {
        alert("读取 GitHub 数据失败：" + e.message);
        return [];
      }
    }
    setStorageNotice(true, noticeHtml());
    return localLoad();
  }

  // 按 id 拉取单条完整记录（索引里没有 rows 缓存时调用）
  async function loadRecordById(id) {
    await ensureMode();
    if (hasGithubConfig() && storageMode !== "server") {
      try {
        const rec = await githubLoadRecord(id);
        if (rec) upsertLocalFull(rec);
        return rec;
      } catch (e) {
        throw new Error("读取记录失败：" + e.message);
      }
    }
    if (storageMode === "server") {
      const res = await fetch("/api/records?id=" + encodeURIComponent(id), { cache: "no-store" });
      if (!res.ok) throw new Error("HTTP " + res.status);
      const rec = await res.json();
      if (rec && rec.rows) upsertLocalFull(rec);
      return rec;
    }
    return readLocalFull().find((r) => r.id === id) || null;
  }

  // 保存单条记录（写单条存储 + 更新索引 + 本地镜像）
  async function persistRecord(rec) {
    await ensureMode();
    rec.updatedAt = rec.updatedAt || new Date().toISOString();
    // 配置了 GitHub：优先写回 GitHub，随后镜像到本地
    if (hasGithubConfig() && storageMode !== "server") {
      setSyncStatus("syncing");
      try {
        const ok = await githubPersistRecord(rec);
        upsertLocalFull(rec);
        setSyncStatus("ok");
        return ok;
      } catch (e) {
        setSyncStatus("err", e.message);
        alert("保存到 GitHub 失败：" + e.message + "，已改存本地。");
        upsertLocalFull(rec);
        return true;
      }
    }
    if (storageMode === "server") {
      try {
        const res = await fetch("/api/records?id=" + encodeURIComponent(rec.id), {
          method: "PUT",
          headers: { "Content-Type": "application/json" },
          body: JSON.stringify(rec),
        });
        if (!res.ok) throw new Error("HTTP " + res.status);
        upsertLocalFull(rec);
        return true;
      } catch (e) {
        setStorageNotice(true);
        alert("未能保存到项目文件夹，已暂存到浏览器：" + e.message);
        upsertLocalFull(rec);
        return true;
      }
    }
    if (storageMode === "github") {
      try {
        return await githubPersistRecord(rec);
      } catch (e) {
        alert("保存到 GitHub 失败：" + e.message);
        return false;
      }
    }
    setStorageNotice(true, noticeHtml());
    // 静态模式未配置 GitHub：静默存浏览器，不再弹提示
    upsertLocalFull(rec);
    return true;
  }

  // 删除单条记录（远端 + 索引 + 本地镜像）
  async function deleteRecordById(id) {
    await ensureMode();
    if (hasGithubConfig() && storageMode !== "server") {
      setSyncStatus("syncing");
      try {
        await githubDeleteRecord(id);
        setSyncStatus("ok");
      } catch (e) {
        setSyncStatus("err", e.message);
        alert("从 GitHub 删除失败：" + e.message);
        return false;
      }
    } else if (storageMode === "server") {
      try {
        const res = await fetch("/api/records?id=" + encodeURIComponent(id), { method: "DELETE" });
        if (!res.ok) throw new Error("HTTP " + res.status);
      } catch (e) {
        alert("删除失败：" + e.message);
        return false;
      }
    }
    removeFromLocalFull(id);
    return true;
  }

  function localLoad() {
    return readLocalFull();
  }

  // GitHub API 读写（静态网页部署模式）：
  // 索引存 cfg.path（records.json），单条记录存同目录 records/<id>.json
  function githubBaseDir(cfg) {
    const p = (cfg.path || GH_PATH_DEFAULT).split("/");
    p.pop(); // 去掉文件名，得到目录
    return p.join("/");
  }

  function githubApiUrl(cfg) {
    const p = (cfg.path || GH_PATH_DEFAULT).split("/").map(encodeURIComponent).join("/");
    const base = `https://api.github.com/repos/${cfg.repo}/contents/${p}`;
    return cfg.branch ? `${base}?ref=${encodeURIComponent(cfg.branch)}` : base;
  }

  function githubFileUrl(cfg, relPath) {
    const p = relPath.split("/").map(encodeURIComponent).join("/");
    const base = `https://api.github.com/repos/${cfg.repo}/contents/${p}`;
    return cfg.branch ? `${base}?ref=${encodeURIComponent(cfg.branch)}` : base;
  }

  function githubRequestUrl(cfg, method, url, body) {
    const ctrl = new AbortController();
    const timer = setTimeout(() => ctrl.abort(), 20000);
    return fetch(url, {
      method,
      headers: {
        Authorization: `Bearer ${cfg.token}`,
        Accept: "application/vnd.github+json",
        "X-GitHub-Api-Version": "2022-11-28",
        ...(body ? { "Content-Type": "application/json" } : {}),
      },
      body: body ? JSON.stringify(body) : undefined,
      signal: ctrl.signal,
    }).finally(() => clearTimeout(timer));
  }

  function githubRequest(cfg, method, body) {
    return githubRequestUrl(cfg, method, githubApiUrl(cfg), body);
  }

  async function githubFetchJson(cfg, relPath) {
    const res = await githubRequestUrl(cfg, "GET", githubFileUrl(cfg, relPath));
    if (res.status === 404) return null;
    if (!res.ok) throw new Error("HTTP " + res.status);
    const data = await res.json();
    return { sha: data.sha, json: JSON.parse(L.decodeUnicodeBase64(data.content)) };
  }

  async function githubPutJson(cfg, relPath, obj, message, sha) {
    const payload = {
      message: message || "更新数据记录（采样计划软件）",
      content: L.encodeUnicodeBase64(JSON.stringify(obj, null, 2)),
      ...(sha ? { sha } : {}),
      ...(cfg.branch ? { branch: cfg.branch } : {}),
    };
    const res = await githubRequestUrl(cfg, "PUT", githubFileUrl(cfg, relPath), payload);
    if (!res.ok) {
      const err = await res.json().catch(() => ({}));
      throw new Error("HTTP " + res.status + "：" + (err.message || "请检查 Token 的 Contents 读/写权限"));
    }
    return (await res.json()).content?.sha || null;
  }

  async function githubDeleteJson(cfg, relPath, sha) {
    const payload = {
      message: "删除数据记录（采样计划软件）",
      sha,
      ...(cfg.branch ? { branch: cfg.branch } : {}),
    };
    const res = await githubRequestUrl(cfg, "DELETE", githubFileUrl(cfg, relPath), payload);
    if (!res.ok && res.status !== 404) {
      const err = await res.json().catch(() => ({}));
      throw new Error("HTTP " + res.status + "：" + (err.message || "删除失败"));
    }
  }

  const recRelPath = (id) => githubBaseDir(loadGithubConfig()) + "/records/" + encodeURIComponent(id) + ".json";

  // 读取索引（旧格式兼容：若索引文件本身是完整记录数组，则整体迁移为分文件结构）
  async function githubLoadIndex() {
    const cfg = loadGithubConfig();
    if (!cfg || !cfg.repo || !cfg.token) throw new Error("未配置 GitHub 仓库或 Token");
    const data = await githubFetchJson(cfg, cfg.path || GH_PATH_DEFAULT);
    if (!data) {
      ghSha = null;
      return [];
    }
    ghSha = data.sha;
    const idx = data.json;
    if (!Array.isArray(idx)) return [];
    if (idx.some((r) => r && Array.isArray(r.rows))) {
      // 旧格式：整表数组 → 逐条写入 records/<id>.json，索引替换为元信息
      const newIndex = [];
      for (const rec of idx) {
        if (!rec || !rec.id) continue;
        await githubPutJson(cfg, recRelPath(rec.id), rec, "迁移数据记录（采样计划软件）", null);
        newIndex.push({ id: rec.id, name: rec.name || "", createdAt: rec.createdAt || "", updatedAt: rec.updatedAt || "" });
      }
      ghSha = await githubPutJson(cfg, cfg.path || GH_PATH_DEFAULT, newIndex, "迁移数据记录索引（采样计划软件）", ghSha);
      return newIndex;
    }
    return idx.filter((r) => r && r.id);
  }

  async function githubLoadRecord(id) {
    const cfg = loadGithubConfig();
    if (!cfg || !cfg.repo || !cfg.token) throw new Error("未配置 GitHub 仓库或 Token");
    const data = await githubFetchJson(cfg, recRelPath(id));
    return data ? data.json : null;
  }

  async function githubPersistRecord(rec) {
    const cfg = loadGithubConfig();
    if (!cfg || !cfg.repo || !cfg.token) throw new Error("未配置 GitHub 仓库或 Token");
    // 写单条文件（409 冲突重试一次）
    let sha = null;
    try {
      const cur = await githubFetchJson(cfg, recRelPath(rec.id));
      sha = cur ? cur.sha : null;
    } catch {}
    for (let attempt = 0; attempt < 2; attempt++) {
      try {
        await githubPutJson(cfg, recRelPath(rec.id), rec, "更新数据记录（采样计划软件）", sha);
        break;
      } catch (e) {
        if (attempt === 0 && String(e.message).includes("409")) {
          const cur = await githubFetchJson(cfg, recRelPath(rec.id)).catch(() => null);
          sha = cur ? cur.sha : null;
          continue;
        }
        throw e;
      }
    }
    // 更新索引
    const index = await githubLoadIndex().catch(() => []);
    const meta = { id: rec.id, name: rec.name || "", createdAt: rec.createdAt || "", updatedAt: rec.updatedAt || "" };
    const pos = index.findIndex((r) => r.id === rec.id);
    if (pos >= 0) index[pos] = meta;
    else index.unshift(meta);
    const idxSha = await githubFetchJson(cfg, cfg.path || GH_PATH_DEFAULT).then((d) => (d ? d.sha : null)).catch(() => null);
    ghSha = await githubPutJson(cfg, cfg.path || GH_PATH_DEFAULT, index, "更新数据记录索引（采样计划软件）", idxSha ?? ghSha);
    return true;
  }

  async function githubDeleteRecord(id) {
    const cfg = loadGithubConfig();
    if (!cfg || !cfg.repo || !cfg.token) throw new Error("未配置 GitHub 仓库或 Token");
    const cur = await githubFetchJson(cfg, recRelPath(id)).catch(() => null);
    if (cur) await githubDeleteJson(cfg, recRelPath(id), cur.sha);
    const index = (await githubLoadIndex().catch(() => [])).filter((r) => r.id !== id);
    const idxSha = await githubFetchJson(cfg, cfg.path || GH_PATH_DEFAULT).then((d) => (d ? d.sha : null)).catch(() => null);
    ghSha = await githubPutJson(cfg, cfg.path || GH_PATH_DEFAULT, index, "更新数据记录索引（采样计划软件）", idxSha ?? ghSha);
    return true;
  }

  // ---------- 危害因素库 / 检测项目参考库：加载、保存与自动同步 ----------
  function setLibStatus(state, msg) {
    for (const el of document.querySelectorAll(".lib-status")) {
      const t = new Date().toLocaleTimeString("zh-CN", { hour12: false });
      if (state === "ok") {
        el.className = "status lib-status lib-ok";
        el.textContent = `✓ 参考库已同步 GitHub（${t}）`;
      } else if (state === "local-ok") {
        el.className = "status lib-status lib-ok";
        el.textContent = `✓ 参考库已保存到本地（${t}）`;
      } else if (state === "static-ok") {
        el.className = "status lib-status lib-ok";
        el.textContent = "✓ 参考库：已加载站点 library.json";
      } else if (state === "err") {
        el.className = "status lib-status err";
        el.textContent = `✕ 参考库同步失败：${msg || "未知错误"}`;
      } else if (state === "pending") {
        el.className = "status lib-status";
        el.textContent = "⟳ 参考库待保存…";
      } else if (state === "github-loading") {
        el.className = "status lib-status";
        el.textContent = "⟳ 正在从 GitHub 加载参考库…";
      } else {
        el.className = "status lib-status";
        el.textContent = "参考库：内置数据";
      }
    }
  }

  function normalizeLibrary(obj) {
    if (!obj || typeof obj !== "object") return null;
    const hf = Array.isArray(obj.hazardFactors) ? obj.hazardFactors.filter((h) => h && typeof h === "object") : null;
    const di = Array.isArray(obj.detectionItems) ? obj.detectionItems.map(String) : null;
    if (!hf && !di) return null;
    const lib = {};
    if (hf) lib.hazardFactors = hf;
    if (di) lib.detectionItems = di;
    return lib;
  }

  function libraryObject() {
    return {
      hazardFactors: hazardFactors.map((h) => ({ ...h })),
      detectionItems: detectionItems.slice(),
    };
  }

  function localLibraryLoad() {
    try {
      const s = localStorage.getItem(LIBRARY_KEY);
      return s ? normalizeLibrary(JSON.parse(s)) : null;
    } catch {
      return null;
    }
  }

  function localLibraryPersist(lib) {
    try {
      localStorage.setItem(LIBRARY_KEY, JSON.stringify(lib));
      return true;
    } catch (e) {
      alert("参考库本地保存失败：" + e.message);
      return false;
    }
  }

  function mirrorLibraryToLocal(lib) {
    try {
      if (storageMode === "server") {
        fetch("/api/library", {
          method: "PUT",
          headers: { "Content-Type": "application/json" },
          body: JSON.stringify(lib),
        }).catch(() => {});
      } else {
        localStorage.setItem(LIBRARY_KEY, JSON.stringify(lib));
      }
    } catch {}
  }

  function githubLibConfig() {
    const cfg = loadGithubConfig();
    if (!cfg) return null;
    return { ...cfg, path: cfg.libPath || GH_LIB_PATH_DEFAULT };
  }

  async function githubLibLoad() {
    const cfg = githubLibConfig();
    if (!cfg || !cfg.repo || !cfg.token) throw new Error("未配置 GitHub 仓库或 Token");
    const res = await githubRequest(cfg, "GET");
    if (res.status === 404) { libSha = null; return null; }
    if (!res.ok) throw new Error("HTTP " + res.status + "（请检查仓库名与 Token 权限）");
    const data = await res.json();
    libSha = data.sha;
    return JSON.parse(L.decodeUnicodeBase64(data.content));
  }

  async function githubLibPersist(lib) {
    const cfg = githubLibConfig();
    if (!cfg || !cfg.repo || !cfg.token) throw new Error("未配置 GitHub 仓库或 Token");
    const content = L.encodeUnicodeBase64(JSON.stringify(lib, null, 2));
    for (let attempt = 0; attempt < 2; attempt++) {
      if (!libSha) {
        const res = await githubRequest(cfg, "GET");
        if (res.status === 404) libSha = null;
        else if (res.ok) libSha = (await res.json()).sha;
        else throw new Error("HTTP " + res.status);
      }
      const payload = {
        message: "更新参考库（危害因素/检测项目）",
        content,
        ...(libSha ? { sha: libSha } : {}),
        ...(cfg.branch ? { branch: cfg.branch } : {}),
      };
      const res = await githubRequest(cfg, "PUT", payload);
      if (res.status === 409 && attempt === 0) {
        libSha = null; // 文件被其他端修改，重取 sha 后重试一次
        continue;
      }
      if (!res.ok) {
        const err = await res.json().catch(() => ({}));
        throw new Error("HTTP " + res.status + "：" + (err.message || "请检查 Token 的 Contents 读/写权限"));
      }
      libSha = (await res.json()).content?.sha || libSha;
      return true;
    }
    throw new Error("更新冲突，请稍后重试");
  }

  async function loadLibrary() {
    await ensureMode();
    const cands = [];
    if (hasGithubConfig()) {
      setLibStatus("github-loading");
      try {
        const lib = await githubLibLoad();
        if (lib) cands.push({ src: "github", lib });
      } catch (e) {
        setLibStatus("err", e.message);
      }
    }
    if (storageMode === "server") {
      try {
        const lib = await cloudLibLoad();
        if (lib) cands.push({ src: "cloud", lib });
      } catch {
        // Cloudflare KV 不可用时忽略，继续其它来源
      }
    }
    if (cands.length) {
      cands.sort((a, b) => libTimestamp(b.lib) - libTimestamp(a.lib));
      const winner = cands[0];
      const lib = normalizeLibrary(winner.lib);
      if (lib) {
        const loser = cands[1];
        // GitHub 与 Cloudflare 双端都可用时：取较新一端，并把它同步到较旧一端（失败不阻塞）
        if (loser && libTimestamp(loser.lib) < libTimestamp(winner.lib)) {
          const syncTo =
            winner.src === "cloud"
              ? () => githubLibPersist(winner.lib)
              : () => cloudLibPersist(winner.lib);
          syncTo().catch(() => {});
        }
        mirrorLibraryToLocal(winner.lib);
        setLibStatus("ok");
        return lib;
      }
    }
    // 没有可用的远端参考库时，回退到站点自带的静态文件
    const staticLib = await fetchStaticLibrary();
    if (staticLib) {
      setLibStatus("static-ok");
      return staticLib;
    }
    return localLibraryLoad();
  }

  // 参考库版本时间戳：用于 GitHub 与 Cloudflare 双端“取最新”
  function libTimestamp(lib) {
    const t = lib && lib.updatedAt;
    const ts = t ? new Date(t).getTime() : NaN;
    return isNaN(ts) ? 0 : ts;
  }

  function withLibTimestamp(lib) {
    return { ...lib, updatedAt: new Date().toISOString() };
  }

  async function cloudLibLoad() {
    const res = await fetch("/api/library", { cache: "no-store" });
    if (!res.ok) throw new Error("HTTP " + res.status);
    const data = await res.json();
    return normalizeLibrary(data) ? data : null;
  }

  async function cloudLibPersist(lib) {
    const res = await fetch("/api/library", {
      method: "PUT",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify(lib),
    });
    if (!res.ok) throw new Error("HTTP " + res.status);
    return true;
  }

  // 读取站点自带的静态 library.json（GitHub Pages 部署时随仓库同步更新）
  async function fetchStaticLibrary() {
    try {
      const res = await fetch("data/library.json", { cache: "no-store" });
      if (!res.ok) return null;
      return normalizeLibrary(await res.json());
    } catch {
      return null;
    }
  }

  async function persistLibrary() {
    await ensureMode();
    const stamped = withLibTimestamp(libraryObject());
    const targets = [];
    if (hasGithubConfig()) targets.push({ name: "GitHub", fn: () => githubLibPersist(stamped) });
    if (storageMode === "server") targets.push({ name: "Cloudflare", fn: () => cloudLibPersist(stamped) });
    if (!targets.length) {
      const ok = localLibraryPersist(stamped);
      setLibStatus(ok ? "local-ok" : "err", ok ? "" : "本地保存失败");
      return ok;
    }
    const errors = [];
    let anyOk = false;
    for (const t of targets) {
      try {
        await t.fn();
        anyOk = true;
      } catch (e) {
        errors.push(t.name + "：" + e.message);
      }
    }
    if (errors.length) {
      setLibStatus("err", errors.join("；"));
      const localOk = localLibraryPersist(stamped);
      return anyOk || localOk;
    }
    setLibStatus("ok");
    return true;
  }

  // 参考库变更后防抖自动同步（新增/删除/修改后约 1.2 秒写回）
  function scheduleLibrarySync() {
    clearTimeout(libSyncTimer);
    setLibStatus("pending");
    libSyncTimer = setTimeout(() => {
      persistLibrary().catch(() => {});
    }, 1200);
  }

  function isBlankRow(r) {
    return (
      INPUT_COLS.every((c) => (r.input[c] ?? "") === "") &&
      MANUAL_COLS.every((c) => (r.manual[c] ?? "") === "")
    );
  }

  function trimBlankRows(list) {
    let end = list.length;
    while (end > 0 && isBlankRow(list[end - 1])) end--;
    return list.slice(0, end);
  }

  function defaultRecordName() {
    const first = rows.find((r) => !isBlankRow(r));
    const a = first && first.input.A ? first.input.A : "未命名车间";
    const d = first && first.input.D ? first.input.D : "";
    const now = new Date();
    const ymd = `${now.getFullYear()}-${String(now.getMonth() + 1).padStart(2, "0")}-${String(now.getDate()).padStart(2, "0")}`;
    return `${a}_${d}_${ymd}`;
  }

  function recordSearchText(rec) {
    const parts = [rec.name];
    if (Array.isArray(rec.rows)) {
      for (const r of rec.rows) parts.push([r.input.A, r.input.B, r.input.C, r.input.D, r.values.AN].filter((v) => v).join(" "));
    }
    return parts.join(" ").toLowerCase();
  }

  // 保存当前表格为一条数据记录（btn-save 与上传成功后联动共用）
  async function saveRecordFlow(defaultName) {
    const contentRows = trimBlankRows(rows);
    if (!contentRows.length) { alert("当前没有可保存的数据。"); return false; }
    const name = await askInput({
      title: "保存数据",
      hint: "为当前数据命名，之后可在「数据记录」中搜索并调用，减少重复输入",
      value: defaultName,
    });
    if (name === null) return false;
    if (name === "") { alert("名称不能为空。"); return false; }
    const list = await loadRecords();
    const now = new Date().toISOString();
    const snap = L.snapshotRows(contentRows);
    const existing = list.find((r) => r.name === name);
    let rec;
    if (existing) {
      if (!confirm(`已存在同名记录「${name}」，是否覆盖？`)) return false;
      rec = { ...existing, rows: snap, updatedAt: now };
    } else {
      rec = {
        id: Date.now() + "_" + Math.random().toString(36).slice(2, 8),
        name,
        createdAt: now,
        updatedAt: now,
        rows: snap,
      };
    }
    const ok = await persistRecord(rec);
    if (ok) alert(`已保存「${name}」（${contentRows.length} 行）。`);
    return ok;
  }

  $("btn-save").addEventListener("click", async () => {
    await saveRecordFlow(defaultRecordName());
  });

  async function renderDbList() {
    const q = $("db-search").value.trim().toLowerCase();
    const list = (await loadRecords()).sort((a, b) => (b.updatedAt || "").localeCompare(a.updatedAt || ""));
    const shown = q ? list.filter((r) => recordSearchText(r).includes(q)) : list;
    $("db-refresh").style.display = hasGithubConfig() ? "" : "none";
    $("db-count").textContent = `共 ${list.length} 条${q ? ` · 匹配 ${shown.length} 条` : ""}`;
    const box = $("db-list");
    if (!shown.length) {
      box.innerHTML = `<div class="db-empty">${q ? "无匹配结果" : "暂无记录，可先点击「保存数据」保存当前表格"}</div>`;
      return;
    }
    box.innerHTML = shown
      .map((r) => {
        const meta = new Date(r.updatedAt).toLocaleString("zh-CN", { hour12: false });
        const rowsArr = Array.isArray(r.rows) ? r.rows : [];
        const preview = rowsArr
          .slice(0, 3)
          .map((row) => `${row.input.A || "?"}｜${row.input.D || ""}`)
          .join("；");
        return (
          `<div class="db-item" data-id="${escAttr(r.id)}">` +
          `<div class="db-item-main">` +
          `<div class="db-item-name">${escHtml(r.name)}</div>` +
          `<div class="db-item-meta">${rowsArr.length ? rowsArr.length + " 行 · " : ""}保存于 ${escHtml(meta)}</div>` +
          `<div class="db-item-preview">${escHtml(preview)}</div>` +
          `</div>` +
          `<div class="db-item-actions">` +
          `<button class="btn small primary" data-act="load">调用</button>` +
          `<button class="btn small danger" data-act="del">删除</button>` +
          `</div></div>`
        );
      })
      .join("");
  }

  $("btn-db").addEventListener("click", async () => {
    await renderDbList();
    $("db-modal").classList.remove("hidden");
  });
  $("db-close").addEventListener("click", () => $("db-modal").classList.add("hidden"));
  $("db-modal").addEventListener("click", (e) => {
    if (e.target.id === "db-modal") $("db-modal").classList.add("hidden");
  });
  $("db-search").addEventListener("input", () => renderDbList());
  $("db-refresh").addEventListener("click", () => renderDbList());

  $("db-list").addEventListener("click", async (e) => {
    const btn = e.target.closest("button[data-act]");
    if (!btn) return;
    const item = btn.closest(".db-item");
    const id = item.dataset.id;
    const meta = (await loadRecords()).find((r) => r.id === id);
    if (!meta) return;
    if (btn.dataset.act === "load") {
      // 索引项可能不含 rows（本地无缓存），按 id 拉取完整记录
      const rec = Array.isArray(meta.rows) && meta.rows.length ? meta : await loadRecordById(id);
      if (!rec || !Array.isArray(rec.rows)) { alert("读取记录内容失败，请检查网络后重试。"); return; }
      const hasData = rows.some((r) => !isBlankRow(r));
      if (hasData && !confirm(`将用「${rec.name}」（${rec.rows.length} 行）替换当前表格，是否继续？`)) return;
      rows = L.restoreRows(rec.rows);
      rowHeights = rows.map(() => ROW_H);
      rowOffsets = null;
      selectedRow = -1;
      recomputeAndRefresh();
      renderWindow();
      $("db-modal").classList.add("hidden");
      alert(`已调用「${rec.name}」（${rec.rows.length} 行）。`);
    } else if (btn.dataset.act === "del") {
      if (!confirm(`确定删除记录「${meta.name}」？此操作不可恢复。`)) return;
      if (await deleteRecordById(id)) {
        await renderDbList();
        alert("已删除。");
      }
    }
  });

  // 备份导出 / 导入
  $("db-export").addEventListener("click", async () => {
    const list = await loadRecords();
    if (!list.length) { alert("数据库为空，无需备份。"); return; }
    await ensureXlsx();
    const blob = new Blob(
      [JSON.stringify({ app: "采样计划软件", version: 1, exportedAt: new Date().toISOString(), records: list }, null, 2)],
      { type: "application/json" }
    );
    X.downloadBlob(blob, `采样计划数据库备份_${new Date().toISOString().slice(0, 10)}.json`);
  });
  $("db-import").addEventListener("click", () => $("db-file-input").click());
  $("db-file-input").addEventListener("change", async (e) => {
    const file = e.target.files[0];
    if (!file) return;
    try {
      const parsed = JSON.parse(await file.text());
      const list = Array.isArray(parsed) ? parsed : parsed.records;
      if (!Array.isArray(list)) throw new Error("备份文件格式不正确");
      const clean = list.filter((r) => r && typeof r.name === "string" && Array.isArray(r.rows));
      if (!clean.length) throw new Error("备份文件中没有有效记录");
      if (!confirm(`将导入 ${clean.length} 条记录（与现有记录按名称合并，同名覆盖），是否继续？`)) {
        e.target.value = "";
        return;
      }
      // 按名称合并：同名取导入版本，逐条保存（单条存储）
      const existing = await loadRecords();
      const nameMap = new Map(existing.map((r) => [r.name, r]));
      let saved = 0;
      for (const rec of clean) {
        const dup = nameMap.get(rec.name);
        const toSave = dup ? { ...dup, rows: rec.rows, updatedAt: new Date().toISOString() } : rec;
        if (await persistRecord(toSave)) saved++;
      }
      if (saved) {
        await renderDbList();
        alert(`已导入 ${saved} 条记录。`);
      }
    } catch (err) {
      alert("导入失败：" + err.message);
    } finally {
      e.target.value = "";
    }
  });

  // ---------- GitHub 配置（静态网页部署模式） ----------
  $("btn-gh").addEventListener("click", () => {
    const cfg = loadGithubConfig() || {};
    $("gh-repo").value = cfg.repo || "";
    $("gh-branch").value = cfg.branch || "main";
    $("gh-path").value = cfg.path || GH_PATH_DEFAULT;
    $("gh-token").value = cfg.token || "";
    $("gh-show-token").checked = false;
    $("gh-token").type = "password";
    const hint = $("gh-saved-hint");
    if (cfg.token) {
      hint.className = "gh-saved ok";
      hint.textContent = "本机已保存 Token（末尾 " + cfg.token.slice(-6) + "），勾选「显示 Token」可查看完整值。";
    } else {
      hint.className = "gh-saved warn";
      hint.textContent = "本机浏览器中未找到已保存的 Token：请在下框粘贴 Token 后点「保存配置」；仅点「测试连接」不会保存。";
    }
    $("gh-msg").textContent = "";
    $("gh-modal").classList.remove("hidden");
  });
  $("gh-show-token").addEventListener("change", () => {
    $("gh-token").type = $("gh-show-token").checked ? "text" : "password";
  });
  $("gh-close").addEventListener("click", () => $("gh-modal").classList.add("hidden"));
  $("gh-modal").addEventListener("click", (e) => {
    if (e.target.id === "gh-modal") $("gh-modal").classList.add("hidden");
  });
  $("gh-test").addEventListener("click", async () => {
    const repo = $("gh-repo").value.trim();
    const token = $("gh-token").value.trim();
    if (!repo || !repo.includes("/") || !token) {
      $("gh-msg").textContent = "请先填写仓库（用户名/仓库名）与 Token";
      return;
    }
    $("gh-msg").textContent = "正在测试连接…";
    const cfg = {
      repo,
      branch: $("gh-branch").value.trim() || "main",
      path: $("gh-path").value.trim() || GH_PATH_DEFAULT,
      token,
    };
    try {
      const res = await githubRequest(cfg, "GET");
      if (res.status === 404) {
        $("gh-msg").textContent = "连接成功：数据文件尚不存在，首次保存时自动创建。仍需点「保存配置」才会保存并启用。";
      } else if (res.ok) {
        $("gh-msg").textContent = "连接成功：可读写 data/records.json。仍需点「保存配置」才会保存并启用。";
      } else {
        const err = await res.json().catch(() => ({}));
        $("gh-msg").textContent = "连接失败：" + res.status + " " + (err.message || "");
      }
    } catch (e) {
      $("gh-msg").textContent = "连接失败：" + e.message;
    }
  });
  $("gh-save").addEventListener("click", async () => {
    const repo = $("gh-repo").value.trim();
    const token = $("gh-token").value.trim();
    if (!repo || !repo.includes("/")) {
      $("gh-msg").textContent = "仓库格式应为：用户名/仓库名";
      return;
    }
    if (!token) {
      $("gh-msg").textContent = "请填写 Token";
      return;
    }
    const cfg = {
      repo,
      branch: $("gh-branch").value.trim() || "main",
      path: $("gh-path").value.trim() || GH_PATH_DEFAULT,
      token,
    };
    if (saveGithubConfig(cfg)) {
      if (storageMode !== "server") storageMode = "github";
      ghSha = null;
      setStorageNotice(false);
      $("gh-modal").classList.add("hidden");
      alert("GitHub 配置已保存，正在从 GitHub 同步数据…");
      loadRecords().catch(() => {}); // 打开即同步
    } else {
      $("gh-msg").textContent = "保存失败（浏览器存储不可用）";
    }
  });
  $("gh-clear").addEventListener("click", () => {
    if (!confirm("确定清除 GitHub 配置？")) return;
    clearGithubConfig();
    ghSha = null;
    setSyncStatus(null);
    if (storageMode === "github") storageMode = "static-unconfigured";
    setStorageNotice(storageMode !== "server", noticeHtml());
    $("gh-msg").textContent = "已清除配置。";
  });

  // ---------- 页签 ----------
  document.querySelectorAll(".tab").forEach((t) => {
    t.addEventListener("click", () => {
      activeTab = t.dataset.tab;
      document.querySelectorAll(".tab").forEach((x) => x.classList.toggle("active", x === t));
      document.querySelectorAll(".tab-panel").forEach((p) => p.classList.toggle("active", p.id === "tab-" + activeTab));
      if (activeTab === "hazard") renderHazard();
      if (activeTab === "survey" && window.SurveySheets) window.SurveySheets.refresh();
      if (activeTab === "items") renderItems();
      if (activeTab === "main") renderWindow();
    });
  });

  // ---------- 危害因素库 ----------
  const HAZARD_HEADERS = ["序号", "识别", "系统名称", "粉尘性质", "定性分析", "委外检测", "计算TWA", "计算STEL", "计算CPE", "计算MAC", "结果保留位数", "存在高毒物品", "不检测原因说明"];
  const HAZARD_KEYS = ["rec", "name", "dust", "qual", "outsource", "twa", "stel", "cpe", "mac", "digits", "highTox", "noTestReason"];
  const YESNO_COLS = new Set([3, 4, 5, 6, 7, 8, 10]);
  const HAZARD_ROW_H = 30;
  let lastHazardScrollTop = 0;

  function blankHazard() {
    const h = {};
    for (const k of HAZARD_KEYS) h[k] = "";
    return h;
  }

  function renderHazardHead() {
    // 筛选模式下每列表头显示下拉箭头（Excel 式）；有激活筛选的列高亮，当前打开弹层的列为蓝色选中态。
    // 注意：HAZARD_HEADERS 首列是「序号」，不在 HAZARD_KEYS 中，键索引 = 表头索引 - 1
    $("hazard-head").innerHTML =
      `<tr>${HAZARD_HEADERS.map((h, ci) => {
        if (ci === 0) return `<th>${h}</th>`;
        const ki = ci - 1;
        return `<th>${h}${hazardFilterMode ? `<button class="th-filter${hazardFilter[ki] ? " active" : ""}${hazardFilterCi === ki ? " open" : ""}" data-ki="${ki}" title="筛选本列">▼</button>` : ""}</th>`;
      }).join("")}</tr>`;
  }

  // ---------- 危害因素库 Excel 式列筛选 ----------
  // hazardFilter: { 列索引: Set(勾选的显示值) }；无键表示该列未筛选
  let hazardFilter = {};
  let hazardFilterMode = false; // 是否开启筛选模式（表头显示每列箭头）
  let hazardFilterCi = -1; // 当前打开弹层的列

  function closeHazardFilterPop() {
    const pop = $("hazard-filter-pop");
    if (pop) pop.remove();
    if (hazardFilterCi >= 0) {
      hazardFilterCi = -1;
      renderHazardHead(); // 移除列箭头的蓝色选中态
    }
    document.removeEventListener("click", hazardFilterOutside, true);
  }

  function hazardFilterOutside(e) {
    const pop = $("hazard-filter-pop");
    if (pop && !pop.contains(e.target) && !e.target.closest(".th-filter")) closeHazardFilterPop();
  }

  function openHazardFilter(ki, anchor) {
    closeHazardFilterPop();
    hazardFilterCi = ki;
    // 计数基准：搜索 + 其他列筛选后的行（不含本列），与列表实际显示一致
    const q = $("hazard-search").value.trim();
    const others = Object.entries(hazardFilter).filter(([k]) => Number(k) !== ki);
    const base = hazardFactors.filter((h) =>
      (!q || h.rec.includes(q) || h.name.includes(q)) &&
      others.every(([k, set]) => set.has(fmt(h[HAZARD_KEYS[k]])))
    );
    const counts = new Map();
    for (const h of base) {
      const v = fmt(h[HAZARD_KEYS[ki]]);
      counts.set(v, (counts.get(v) || 0) + 1);
    }
    const values = [...counts.keys()].sort((a, b) => a.localeCompare(b, "zh-CN"));
    const checked = hazardFilter[ki] || new Set(values);
    const headerName = HAZARD_HEADERS[HAZARD_KEYS.indexOf(HAZARD_KEYS[ki]) + 1] || HAZARD_KEYS[ki];
    const pop = document.createElement("div");
    pop.id = "hazard-filter-pop";
    pop.innerHTML =
      `<div class="hfp-title">筛选（${escHtml(headerName)}）</div>` +
      `<div class="hfp-ops">` +
      `<button type="button" data-op="all">全选</button>` +
      `<button type="button" data-op="none">反选</button>` +
      `<button type="button" data-op="clear">清除本列</button>` +
      `</div>` +
      `<div class="hfp-list">` +
      values.map((v) => {
        const has = checked.has(v);
        return `<label class="hfp-item"><input type="checkbox" value="${escAttr(v)}"${has ? " checked" : ""}/>` +
          `<span class="hfp-val">${v === "" ? "（空）" : escHtml(v)}</span>` +
          `<span class="hfp-cnt">${counts.get(v)}</span></label>`;
      }).join("") +
      `</div>` +
      `<div class="hfp-foot"><button type="button" class="btn small primary" data-op="ok">确定</button>` +
      `<button type="button" class="btn small ghost" data-op="cancel">取消</button></div>`;
    document.body.appendChild(pop);
    const r = anchor.getBoundingClientRect();
    pop.style.left = Math.min(r.left, window.innerWidth - 260) + "px";
    pop.style.top = Math.min(r.bottom + 4, window.innerHeight - 320) + "px";
    // 勾选即时生效（Excel 是点确定生效，这里即时刷新更直观；取消则还原）
    const snapshot = new Set(checked);
    // 从弹层勾选框读取选中值集合（存字符串值，不能存元素）
    const readChecked = () => new Set([...pop.querySelectorAll(".hfp-item input:checked")].map((el) => el.value));
    pop.addEventListener("change", () => {
      const cur = readChecked();
      if (cur.size === values.length) delete hazardFilter[ki];
      else hazardFilter[ki] = cur;
      renderHazard();
      renderHazardHead();
      reopenAt(anchor);
    });
    pop.addEventListener("click", (e) => {
      const btn = e.target.closest("button[data-op]");
      if (!btn) return;
      const op = btn.dataset.op;
      if (op === "all" || op === "none") {
        const target = op === "all";
        for (const el of pop.querySelectorAll(".hfp-item input")) el.checked = target;
      } else if (op === "clear") {
        delete hazardFilter[ki];
        closeHazardFilterPop();
        renderHazard();
        renderHazardHead();
        return;
      } else if (op === "cancel") {
        // 取消：清除本列筛选，恢复显示全部内容（不保留打开弹层时的状态）
        delete hazardFilter[ki];
        closeHazardFilterPop();
        renderHazard();
        renderHazardHead();
        return;
      } else if (op === "ok") {
        const cur = readChecked();
        if (cur.size === values.length) delete hazardFilter[ki];
        else hazardFilter[ki] = cur;
        closeHazardFilterPop();
        renderHazard();
        renderHazardHead();
        return;
      }
      // 全选/反选后即时刷新
      const cur = readChecked();
      if (cur.size === values.length) delete hazardFilter[ki];
      else hazardFilter[ki] = cur;
      renderHazard();
      renderHazardHead();
      reopenAt(anchor);
    });
    function reopenAt(a) {
      // 重渲染表头后按钮节点被替换，重新定位弹层
      const btn2 = document.querySelector(`.th-filter[data-ki="${ki}"]`);
      if (btn2 && hazardFilterCi === ki) {
        const r2 = btn2.getBoundingClientRect();
        pop.style.left = Math.min(r2.left, window.innerWidth - 260) + "px";
        pop.style.top = Math.min(r2.bottom + 4, window.innerHeight - 320) + "px";
      }
    }
    setTimeout(() => document.addEventListener("click", hazardFilterOutside, true), 0);
  }

  $("hazard-head").addEventListener("click", (e) => {
    const btn = e.target.closest(".th-filter");
    if (!btn) return;
    e.stopPropagation();
    const ki = Number(btn.dataset.ki);
    if (hazardFilterCi === ki) closeHazardFilterPop();
    else openHazardFilter(ki, btn);
  });

  // 工具栏「筛选」按钮：切换筛选模式（表头显示/收起每列箭头）；
  // 关闭时清空全部列筛选，恢复显示原始内容
  $("hazard-filter").addEventListener("click", () => {
    hazardFilterMode = !hazardFilterMode;
    closeHazardFilterPop();
    if (!hazardFilterMode) hazardFilter = {};
    renderHazardHead();
    renderHazard();
  });

  function renderHazard() {
    const q = $("hazard-search").value.trim();
    const list = hazardFactors
      .map((h, idx) => ({ h, idx }))
      .filter((x) => !q || x.h.rec.includes(q) || x.h.name.includes(q))
      // 列筛选：每列勾选值之间是「或」，列与列之间是「与」（Excel 语义）
      .filter((x) => Object.entries(hazardFilter).every(([ci, set]) => set.has(fmt(x.h[HAZARD_KEYS[ci]]))));
    const body = $("hazard-body");
    const wrap = document.querySelector(".hazard-wrap");
    // 窗口化渲染：只生成可见行，大幅降低点击/删除/搜索时的 DOM 开销
    const st = wrap ? wrap.scrollTop : 0;
    const ch = wrap ? wrap.clientHeight : 600;
    const start = Math.max(0, Math.floor(st / HAZARD_ROW_H) - 10);
    const visible = Math.ceil(ch / HAZARD_ROW_H) + 22;
    const end = Math.min(list.length, start + visible);
    let html = "";
    // 顶部占位行：保证表格总高度恒定（list.length × 行高），滚动位置才能稳定推进
    const topSpacer = start * HAZARD_ROW_H;
    if (topSpacer > 0) html += `<tr class="hazard-spacer" style="height:${topSpacer}px"><td colspan="${HAZARD_KEYS.length + 1}"></td></tr>`;
    for (let i = start; i < end; i++) {
      const { h, idx: absIdx } = list[i];
      const sel = absIdx === selectedHazard;
      html += `<tr data-h="${absIdx}"${sel ? ' class="selected"' : ""}>` +
        `<td class="rowno${sel ? " selected" : ""}" data-h="${absIdx}">${absIdx + 1}</td>` +
        HAZARD_KEYS.map((k, ci) => {
          const val = fmt(h[k]);
          let inner;
          if (ci === 2) inner = `<select data-k="${k}">${optionList(["", "总尘", "呼尘"], val)}</select>`;
          else if (YESNO_COLS.has(ci)) inner = `<select data-k="${k}">${optionList(["", "是", "否"], val)}</select>`;
          else if (ci === 9) inner = `<input data-k="${k}" type="number" value="${escAttr(val)}">`;
          else inner = `<input data-k="${k}" value="${escAttr(val)}">`;
          return `<td class="${ci === 9 ? "cell-num" : ""}">${inner}</td>`;
        }).join("") + `</tr>`;
    }
    const spacer = Math.max(0, list.length - end) * HAZARD_ROW_H;
    if (spacer > 0) html += `<tr class="hazard-spacer" style="height:${spacer}px"><td colspan="${HAZARD_KEYS.length + 1}"></td></tr>`;
    body.innerHTML = html;
    $("hazard-status").textContent = `共 ${hazardFactors.length} 条 · 显示 ${list.length} 条`;
    applyHazardWidths();
  }

  const hazardWrap = document.querySelector(".hazard-wrap");
  if (hazardWrap) {
    hazardWrap.addEventListener("scroll", () => {
      if (Math.abs(hazardWrap.scrollTop - lastHazardScrollTop) > HAZARD_ROW_H) {
        lastHazardScrollTop = hazardWrap.scrollTop;
        renderHazard();
      }
    });
  }

  function optionList(options, cur) {
    return options.map((o) => `<option value="${escAttr(o)}"${cur === o ? " selected" : ""}>${o === "" ? "（空）" : escHtml(o)}</option>`).join("");
  }

  $("hazard-body").addEventListener("input", (e) => {
    const el = e.target;
    if (el.tagName !== "INPUT") return;
    const idx = Number(el.closest("tr").dataset.h);
    hazardFactors[idx][el.dataset.k] = el.value;
    onHazardChange();
  });
  $("hazard-body").addEventListener("change", (e) => {
    const el = e.target;
    if (el.tagName !== "SELECT") return;
    const idx = Number(el.closest("tr").dataset.h);
    hazardFactors[idx][el.dataset.k] = el.value;
    onHazardChange();
  });
  $("hazard-body").addEventListener("click", (e) => {
    const tr = e.target.closest("tr[data-h]");
    if (!tr) return;
    const idx = Number(tr.dataset.h);
    selectedHazard = idx;
    // 只切换高亮类，不重建表格，点击无卡顿
    for (const t of $("hazard-body").querySelectorAll("tr[data-h]")) {
      const sel = Number(t.dataset.h) === idx;
      t.classList.toggle("selected", sel);
      const rno = t.querySelector("td.rowno");
      if (rno) rno.classList.toggle("selected", sel);
    }
  });
  $("hazard-search").addEventListener("input", () => {
    if (hazardWrap) hazardWrap.scrollTop = 0;
    renderHazard();
  });
  $("hazard-add").addEventListener("click", () => {
    hazardFactors.push(blankHazard());
    selectedHazard = hazardFactors.length - 1;
    if (hazardWrap) hazardWrap.scrollTop = hazardWrap.scrollHeight;
    renderHazard();
    onHazardChange();
  });
  $("hazard-del").addEventListener("click", () => {
    if (selectedHazard < 0) { alert("请先点击选中要删除的因素行。"); return; }
    if (!confirm(`确定删除「${hazardFactors[selectedHazard].rec || hazardFactors[selectedHazard].name || selectedHazard + 1}」？`)) return;
    hazardFactors.splice(selectedHazard, 1);
    selectedHazard = -1;
    renderHazard();
    onHazardChange();
  });
  $("hazard-restore").addEventListener("click", () => {
    if (!confirm("恢复为软件内置的默认危害因素库？当前库将被替换。")) return;
    hazardFactors = DEFAULT_DATA.hazardFactors.map((h) => ({ ...h }));
    selectedHazard = -1;
    renderHazard();
    onHazardChange();
  });

  function onHazardChange() {
    rebuildDatalist();
    $("hazard-count").textContent = hazardFactors.length;
    recomputeAndRefresh();
    scheduleLibrarySync();
  }

  // ---------- 检测项目 ----------
  let selectedItem = -1;

  function renderItems() {
    const q = $("items-search").value.trim();
    const dups = new Set(L.findDuplicates(detectionItems));
    const list = detectionItems
      .map((it, idx) => ({ it, idx }))
      .filter((x) => !q || x.it.includes(q));
    $("items-list").innerHTML = list
      .map((x) => `<div class="item${x.idx === selectedItem ? " selected" : ""}${dups.has(x.it) ? " dup" : ""}" data-idx="${x.idx}">${escHtml(x.it)}${dups.has(x.it) ? "（重复）" : ""}</div>`)
      .join("");
    $("items-status").textContent = `共 ${detectionItems.length} 项 · 显示 ${list.length} 项${dups.size ? " · 重复 " + dups.size + " 项" : ""}`;
  }
  $("items-search").addEventListener("input", renderItems);
  $("items-list").addEventListener("click", (e) => {
    const div = e.target.closest(".item");
    if (!div) return;
    selectedItem = Number(div.dataset.idx);
    renderItems();
  });
  $("items-add").addEventListener("click", async () => {
    const v = await askInput({ title: "新增检测项目", hint: "输入标准化检测项目名称" });
    if (v === null) return;
    const name = v.trim();
    if (!name) { alert("名称不能为空。"); return; }
    detectionItems.push(name);
    selectedItem = detectionItems.length - 1;
    renderItems();
    $("items-count").textContent = detectionItems.length;
    scheduleLibrarySync();
  });
  $("items-del").addEventListener("click", () => {
    if (selectedItem < 0) { alert("请先点击选择要删除的项目。"); return; }
    const name = detectionItems[selectedItem];
    if (!confirm(`确定删除检测项目「${name}」？`)) return;
    detectionItems.splice(selectedItem, 1);
    selectedItem = -1;
    renderItems();
    $("items-count").textContent = detectionItems.length;
    scheduleLibrarySync();
  });

  // ---------- 联想提示 ----------
  function rebuildDatalist() {
    const seen = new Set();
    const opts = [];
    for (const h of hazardFactors) {
      for (const v of [h.rec, h.name]) {
        if (v && !seen.has(v)) {
          seen.add(v);
          opts.push(`<option value="${escAttr(v)}">`);
        }
      }
    }
    const dl = document.createElement("datalist");
    dl.id = "hazard-dl";
    dl.innerHTML = opts.join("");
    const old = document.getElementById("hazard-dl");
    if (old) old.remove();
    document.body.appendChild(dl);
    // 岗位工作班制 / 岗位班制数：下拉联想（可手动录入）
    const buildBanzhi = (id, options) => {
      let bdl = document.getElementById(id);
      if (!bdl) {
        bdl = document.createElement("datalist");
        bdl.id = id;
        document.body.appendChild(bdl);
      }
      bdl.innerHTML = (options || []).map((o) => `<option value="${escAttr(o)}">`).join("");
    };
    buildBanzhi("banzhi-r-dl", L.BANZHI_R);
    buildBanzhi("banzhi-s-dl", L.BANZHI_S);
    // 作业方式 / 采样方式 / 岗位性质：录入区下拉联想（可手动录入）
    buildBanzhi("zuoye-dl", ZUOYE_FS);
    buildBanzhi("caiyang-dl", CAIYANG_FS);
    buildBanzhi("gangwei-dl", L.GANGWEI_XZ);
  }

  // ---------- 导出 ----------
  // 导出行范围：到“检测项目”列（AN）最后一个非空单元格为止；
  // 若整表 AN 均为空（如接害因素未识别），则回退到最后一条有录入内容的行，避免导出空表
  function lastExportRow() {
    let lastAn = -1;
    let lastContent = -1;
    for (let i = 0; i < rows.length; i++) {
      const r = rows[i];
      if ((r.values.AN ?? "") !== "") lastAn = i;
      if (!isBlankRow(r)) lastContent = i;
    }
    return lastAn >= 0 ? lastAn : lastContent;
  }

  async function exportWorkbook() {
    await ensureXlsx();
    // 仅导出自动计算区（W~BI 共 39 列）
    const mainRows = [COMPUTED_HEADERS];
    const end = lastExportRow() + 1;
    for (let i = 0; i < end; i++) {
      const r = rows[i];
      const line = [];
      for (const c of COMPUTED_COLS) {
        if (MANUAL_COLS.includes(c)) line.push(r.manual[c]);
        else line.push(NUM_COLS.has(c) ? L.toNum(r.values[c]) : r.values[c]);
      }
      mainRows.push(line);
    }
    return X.writeWorkbook({ mainRows, mainWidths: COMPUTED_WIDTHS });
  }

  // 供上传模块（js/upload.js）复用的导出能力
  window.SamplingApp = {
    exportWorkbookBytes: exportWorkbook,
    exportName: () => "系统测点布局调查_自动计算区.xlsx",
    countErrors: () => L.countErrors(rows).total,
    // 上传成功后联动保存：传入默认名（如「25年XX公司」），弹出保存命名框走统一保存链路
    promptSave: (defaultName) => saveRecordFlow(defaultName || defaultRecordName()),
    // 行数输入弹窗（调查表等模块复用，保持交互一致）
    askRowCount,
    // 主表格已填车间名称（去重、保序），供调查表「单元/工作场所」下拉引用
    workshopNames: () => {
      const out = [];
      for (const row of rows) {
        const v = (row.input["A"] ?? "").trim();
        if (v && !out.includes(v)) out.push(v);
      }
      return out;
    },
  };

  $("btn-export").addEventListener("click", async () => {
    const { total } = L.countErrors(rows);
    if (total > 0 && !confirm(`当前有 ${total} 处校验错误，仍要导出吗？`)) return;
    const bytes = await exportWorkbook();
    const blob = new Blob([bytes], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" });
    X.downloadBlob(blob, "系统测点布局调查_自动计算区.xlsx");
  });

  // 数据上传：切换到上传视图（登录/选项目/上传当前表格导出的 Excel）
  $("btn-upload").addEventListener("click", () => {
    if (window.SamplingUpload && window.SamplingUpload.show) window.SamplingUpload.show();
    else alert("上传模块未加载，请刷新页面后重试。");
  });

  // 退出登录（主页面顶栏）
  $("btn-logout").addEventListener("click", () => {
    if (window.SamplingUpload && window.SamplingUpload.logout) window.SamplingUpload.logout();
    else alert("登录模块未加载，请刷新页面后重试。");
  });

  // ---------- 导入（仅主表测点布局，不导入危害因素库/检测项目） ----------

  $("btn-import").addEventListener("click", () => $("file-input").click());
  $("file-input").addEventListener("change", async (e) => {
    const file = e.target.files[0];
    if (!file) return;
    try {
      await ensureXlsx();
      const sheets = await X.readWorkbook(await file.arrayBuffer());
      const byName = {};
      for (const s of sheets) byName[s.name] = X.sheetToArray(s);

      // 测点布局：优先按常用表名匹配，找不到时按表头智能识别（含“*检测项目”列的工作表即主表）
      let mainSheetName = byName["测点布局情况调查"] ? "测点布局情况调查"
        : byName["劳动定员和职业病危害因素接触情况调查"] ? "劳动定员和职业病危害因素接触情况调查" : null;
      if (!mainSheetName) {
        for (const name of Object.keys(byName)) {
          const hdr = (byName[name][0] || []).map((h) => String(h ?? "").trim());
          if (hdr.includes("*检测项目")) { mainSheetName = name; break; }
        }
      }
      const mainArr = mainSheetName ? byName[mainSheetName] : null;
      if (mainArr) {
        const arr = mainArr;
        const hdr = arr[0] || [];
        const idx = {};
        hdr.forEach((h, i) => { if (h !== null && h !== undefined && h !== "") idx[String(h).trim()] = i; });
        // 录入区独有的表头（自动计算区不含这些名称），用于区分“完整结构”与“仅自动计算区”文件
        const UNIQUE_INPUT_HEADERS = ["车间", "接害因素", "接触时间h/d", "人数"];
        const hasInput = UNIQUE_INPUT_HEADERS.some((h) => h in idx);
        // 仅自动计算区文件：把计算列反推回录入列（与计算引擎推导方向相反），
        // 导入后录入区可继续编辑，计算区随之联动重算
        const COMPUTED_TO_INPUT = {
          W: "A", X: "B", AL: "C", AN: "D", AZ: "E", AH: "F", AG: "G", AA: "H", BA: "I",
          AM: "J", BE: "K", AC: "L", AO: "M", AR: "N", Y: "O", BC: "P", AQ: "Q",
          AD: "R", AE: "S", AF: "T", AW: "U",
        };
        const imported = [];
        const importedVals = [];
        for (const r of arr.slice(1)) {
          if (!r.some((v) => v !== "" && v !== null && v !== undefined)) continue;
          const row = blankRow();
          const vals = {};
          if (hasInput) {
            for (const c of INPUT_COLS) {
              const h = HEADERS[colIdx(c)];
              if (h in idx) row.input[c] = fmt(r[idx[h]]);
            }
          } else {
            for (const [comp, inp] of Object.entries(COMPUTED_TO_INPUT)) {
              const h = HEADERS[colIdx(comp)];
              const v = h in idx ? r[idx[h]] : "";
              if (v !== "" && v !== null && v !== undefined) row.input[inp] = fmt(v);
            }
            // AE 的自动默认值“长白班”不是录入区 S 列的合法下拉值，跳过反推（引擎会自动重算出来）
            if (row.input.S === "长白班") row.input.S = "";
            // 接害因素：优先用库内“识别名”，其次用系统名对应的识别名，保证导入后能反查出检测项目
            const anText = row.input.D;
            if (anText) {
              const byName = hazardFactors.find((x) => x.name === anText);
              const byRec = hazardFactors.find((x) => x.rec === anText);
              row.input.D = byName && !byRec ? byName.rec : anText;
            }
          }
          for (const c of MANUAL_COLS) {
            const h = HEADERS[colIdx(c)];
            if (h in idx) row.manual[c] = fmt(r[idx[h]]);
          }
          if (hasInput) {
            for (const c of COMPUTED_COLS) {
              const h = HEADERS[colIdx(c)];
              if (h in idx) vals[c] = fmt(r[idx[h]]);
            }
          }
          imported.push(row);
          importedVals.push(vals);
        }
        if (imported.length) {
          rows = imported;
          rowHeights = rows.map(() => ROW_H);
          rowOffsets = null;
          L.computeRows(rows, { hazardFactors, detectionItems });
          if (hasInput) {
            // 完整结构文件：保留文件里的覆盖值
            rows.forEach((row, i) => {
              for (const c of [...OVERRIDE_COLS, ...TEXT_OVERRIDE_COLS]) {
                const importedV = importedVals[i][c];
                if (importedV !== undefined && importedV !== "" && importedV !== row.values[c]) {
                  row.values[c] = importedV;
                  row.overridden[c] = true;
                } else {
                  delete row.overridden[c];
                }
              }
            });
            L.computeRows(rows, { hazardFactors, detectionItems });
          } else {
            // 仅自动计算区文件：备注列（BI）允许手动编辑，导入时保留手动值
            rows.forEach((row, i) => {
              for (const c of TEXT_OVERRIDE_COLS) {
                const importedV = importedVals[i][c];
                if (importedV !== undefined && importedV !== "" && importedV !== row.values[c]) {
                  row.values[c] = importedV;
                  row.overridden[c] = true;
                }
              }
            });
            L.computeRows(rows, { hazardFactors, detectionItems });
          }
          selectedRow = -1;
        }
      }
      rebuildDatalist(); // 参考库不随导入变化，仅重建录入区的联想列表
      renderItems();
      renderWindow();
      activeTab = "main";
      document.querySelectorAll(".tab").forEach((t) => t.classList.toggle("active", t.dataset.tab === "main"));
      document.querySelectorAll(".tab-panel").forEach((p) => p.classList.toggle("active", p.id === "tab-main"));
      alert(`导入成功：测点 ${rows.length} 行，已同步到录入区（危害因素库与检测项目保持现有）。`);
    } catch (err) {
      console.error(err);
      alert("导入失败：" + err.message);
    } finally {
      e.target.value = "";
    }
  });

  // ---------- 帮助 ----------
  $("btn-help").addEventListener("click", () => $("help-modal").classList.remove("hidden"));
  $("help-close").addEventListener("click", () => $("help-modal").classList.add("hidden"));
  $("help-modal").addEventListener("click", (e) => {
    if (e.target.id === "help-modal") $("help-modal").classList.add("hidden");
  });

  // ---------- 启动 ----------
  async function init() {
    hazardFactors = DEFAULT_DATA.hazardFactors.map((h) => ({ ...h }));
    detectionItems = DEFAULT_DATA.detectionItems.slice();
    emptyGrid();
    buildHead();
    renderHazardHead();
    rebuildDatalist();
    await detectStorageMode();
    // GitHub 已配置并同步时不显示提示条，仅未配置/临时模式时提示
    setStorageNotice(storageMode !== "server" && storageMode !== "github", noticeHtml());
    try {
      const lib = await loadLibrary();
      if (lib) {
        if (lib.hazardFactors) hazardFactors = lib.hazardFactors.map((h) => ({ ...h }));
        if (lib.detectionItems) detectionItems = lib.detectionItems.map(String);
      }
    } catch {}
    rebuildDatalist();
    if (hasGithubConfig()) {
      loadRecords().catch(() => {}); // 打开时自动从 GitHub 同步
    }
    L.computeRows(rows, { hazardFactors, detectionItems });
    $("hazard-count").textContent = hazardFactors.length;
    $("items-count").textContent = detectionItems.length;
    renderHazard();
    renderItems();
    applyGridWidths();
    applyHazardWidths();
    renderWindow();
  }
  init();
})();
