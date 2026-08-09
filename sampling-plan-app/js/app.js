(function () {
  "use strict";

  const L = window.SamplingLogic;
  const X = window.SamplingXlsx;
  const DEFAULT_DATA = window.SamplingData;

  // ---------- 列定义 ----------
  const INPUT_COLS = L.INPUT_COLS; // A..U
  const COMPUTED_COLS = L.COMPUTED_COLS; // W..BI
  const MANUAL_COLS = L.MANUAL_COLS;
  const OVERRIDE_COLS = L.OVERRIDE_COLS;
  const NUM_COLS = new Set(["E", "F", "G", "H", "I", "AA", "AB", "AG", "AH", "AY", "AZ", "BA", "BB"]);
  const ALL_COLS = [...INPUT_COLS, "V", ...COMPUTED_COLS]; // 61 列（A..BI）
  const HEADERS = DEFAULT_DATA.mainHeaders; // 61 项
  const ROW_H = 30;

  const colIdx = (letter) => {
    let c = 0;
    for (const ch of letter) c = c * 26 + (ch.charCodeAt(0) - 64);
    return c - 1;
  };
  // 各列宽度（Excel 字符单位近似）
  const COL_W = ALL_COLS.map((c) => {
    const w = {
      A: 12, B: 12, C: 10, D: 18, E: 9, F: 9, G: 9, H: 7, I: 8, J: 16,
      K: 14, L: 12, M: 8, N: 8, O: 9, P: 8, Q: 8, R: 11, S: 13, T: 11,
      U: 10, V: 4, W: 12, X: 12, Y: 9, Z: 16, AA: 10, AB: 9, AC: 12,
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

  function measureTextWidth(text, sizePx, bold) {
    measureCtx.font = `${bold ? "bold " : ""}${sizePx}px ${CELL_FONT}`;
    return measureCtx.measureText(String(text)).width;
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

  function applyGridWidths() {
    const w = columnWidths(
      ALL_COLS,
      (i) => HEADERS[i],
      (r, i) => {
        const c = ALL_COLS[i];
        if (c === "V") return "";
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
    w[21] = 18; // V 列是分隔空列，保持窄
    $("grid-cols").innerHTML =
      "<col style='width:42px'>" + w.map((x) => `<col style="width:${x}px">`).join("");
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
    $("hazard-cols").innerHTML =
      "<col style='width:42px'>" + w.map((x) => `<col style="width:${x}px">`).join("");
  }

  let widthTimer = null;
  function scheduleWidths() {
    clearTimeout(widthTimer);
    widthTimer = setTimeout(() => {
      applyGridWidths();
      applyHazardWidths();
    }, 350);
  }

  // ---------- 状态 ----------
  let rows = [];
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

  function emptyGrid(n = 20) {
    rows = [];
    for (let i = 0; i < n; i++) rows.push(blankRow());
  }

  // 示例数据（与原表第 2~9 行一致）
  function loadSample() {
    const samples = [
      { A: "3号车间", B: "操作工", C: "投料", D: "二氧化钛粉尘(总尘)", E: "4", F: "8", G: "6", H: "4", I: "1", P: "是" },
      { A: "3号车间", B: "操作工", C: "投料", D: "硫酸钡", P: "是" },
      { A: "3号车间", B: "操作工", C: "投料", D: "丙烯酸", P: "是" },
      { D: "苯乙烯", P: "是" },
      { A: "3号车间", B: "操作工", C: "投料", D: "噪声", P: "是" },
      { A: "3号车间", B: "操作工", C: "放料", D: "丙烯酸", E: "1", P: "是" },
      { D: "苯乙烯", P: "是" },
      { A: "3号车间", B: "操作工", C: "放料", D: "噪声", P: "是" },
    ];
    rows = samples.map((s) => {
      const r = blankRow();
      for (const k of Object.keys(s)) r.input[k] = s[k];
      return r;
    });
  }

  // ---------- 表头 ----------
  function buildHead() {
    const groupTh = (label, colSpan) =>
      `<th class="group" colspan="${colSpan}">${label}</th>`;
    const fieldTh = (label) => `<th class="field">${label || "&nbsp;"}</th>`;
    let groupHtml = `<th class="corner" rowspan="2" style="width:40px">行</th>`;
    groupHtml += groupTh("录 入 区", INPUT_COLS.length);
    groupHtml += groupTh("&nbsp;", 1);
    groupHtml += groupTh("自动计算区（与原表 W~BI 列一致）", COMPUTED_COLS.length);
    let fieldHtml = "";
    for (const c of ALL_COLS) fieldHtml += fieldTh(HEADERS[colIdx(c)]);
    gridHead.innerHTML =
      `<tr class="group-row">${groupHtml}</tr>` +
      `<tr class="field-row">${fieldHtml}</tr>`;
  }

  // ---------- 单元格渲染 ----------
  function cellClass(col) {
    if (OVERRIDE_COLS.includes(col)) return "override";
    if (MANUAL_COLS.includes(col)) return "manual";
    if (COMPUTED_COLS.includes(col)) return "computed";
    return "";
  }

  function cellHtml(row, idx, col) {
    const cls = cellClass(col);
    const num = NUM_COLS.has(col) ? " cell-num" : "";
    let inner;
    if (col === "V") {
      inner = `<input readonly data-c="${col}" value="">`;
    } else if (OVERRIDE_COLS.includes(col)) {
      inner = selectHtml(col, row.values[col], row.overridden[col] === true);
    } else if (col === "R" || col === "S" || col === "U") {
      inner = selectHtml(col, row.input[col], false, true);
    } else if (MANUAL_COLS.includes(col)) {
      if (col === "AI" || col === "AJ" || col === "BH") {
        inner = selectHtml(col, row.manual[col], false, true);
      } else {
        inner = `<input data-c="${col}" value="${escAttr(fmt(row.manual[col]))}">`;
      }
    } else if (COMPUTED_COLS.includes(col)) {
      inner = `<input readonly data-c="${col}" value="${escAttr(fmt(row.values[col]))}">`;
    } else {
      const dl = col === "D" ? ' list="hazard-dl"' : "";
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
  let lastScrollTop = 0;
  let renderedRows = [];

  function renderWindow() {
    const total = rows.length;
    const st = gridWrap.scrollTop;
    const ch = gridWrap.clientHeight;
    const start = Math.max(0, Math.floor(st / ROW_H) - 8);
    const visible = Math.ceil(ch / ROW_H) + 16;
    const end = Math.min(total, start + visible);
    let html = "";
    renderedRows = [];
    for (let i = start; i < end; i++) {
      renderedRows.push(i);
      const r = rows[i];
      let cells = `<td class="rowno${i === selectedRow ? " selected" : ""}" data-r="${i}">${i + 1}</td>`;
      for (const c of ALL_COLS) cells += cellHtml(r, i, c);
      html += `<tr data-r="${i}"${i === selectedRow ? ' class="selected"' : ""}>${cells}</tr>`;
    }
    const spacer = Math.max(0, total - end) * ROW_H;
    if (spacer > 0) html += `<tr style="height:${spacer}px"><td></td></tr>`;
    gridBody.innerHTML = html;
    refreshStatus();
    updateSelectionClasses();
  }

  gridWrap.addEventListener("scroll", () => {
    if (Math.abs(gridWrap.scrollTop - lastScrollTop) > ROW_H) {
      lastScrollTop = gridWrap.scrollTop;
      renderWindow();
    }
  });

  // ---------- 输入联动 ----------
  gridBody.addEventListener("input", (e) => {
    const el = e.target;
    if (el.tagName !== "INPUT") return;
    const tr = el.closest("tr");
    const r = Number(tr.dataset.r);
    const c = el.dataset.c;
    const row = rows[r];
    if (INPUT_COLS.includes(c)) row.input[c] = el.value;
    recomputeAndRefresh();
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
    recomputeAndRefresh();
  });

  gridBody.addEventListener("click", (e) => {
    const td = e.target.closest("td.rowno");
    if (!td) return;
    const r = Number(td.closest("tr").dataset.r);
    selectedRow = selectedRow === r ? -1 : r;
    renderWindow();
  });

  function recomputeAndRefresh() {
    L.computeRows(rows, { hazardFactors, detectionItems });
    updateVisibleCells();
    refreshStatus();
    scheduleWidths();
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
        if (!el) continue;
        let val;
        if (OVERRIDE_COLS.includes(c)) val = fmt(row.values[c]);
        else if (MANUAL_COLS.includes(c)) val = fmt(row.manual[c]);
        else if (COMPUTED_COLS.includes(c)) val = fmt(row.values[c]);
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
    return c;
  }

  // ---------- 行操作（支持自定义行数） ----------
  $("btn-add").addEventListener("click", async () => {
    const n = await askRowCount("新增行数", 1);
    if (n === null) return;
    if (rows.length + n > 5000) { alert("总行数不能超过 5000。"); return; }
    for (let i = 0; i < n; i++) rows.push(blankRow());
    recomputeAndRefresh();
    gridWrap.scrollTop = gridWrap.scrollHeight;
    renderWindow();
  });
  $("btn-insert").addEventListener("click", async () => {
    const n = await askRowCount("插入行数", 1);
    if (n === null) return;
    if (rows.length + n > 5000) { alert("总行数不能超过 5000。"); return; }
    const at = selectedRow >= 0 ? selectedRow + 1 : rows.length;
    const ins = [];
    for (let i = 0; i < n; i++) ins.push(blankRow());
    rows.splice(at, 0, ...ins);
    if (cur && cur.r >= at) cur.r += n;
    selectedRow = at;
    recomputeAndRefresh();
    renderWindow();
  });
  $("btn-copy").addEventListener("click", async () => {
    const n = await askRowCount("复制行数", 1);
    if (n === null) return;
    const at = selectedRow >= 0 ? selectedRow : rows.length - 1;
    if (at < 0) return;
    if (rows.length + n > 5000) { alert("总行数不能超过 5000。"); return; }
    const src = rows[at];
    const copies = [];
    for (let i = 0; i < n; i++) copies.push(cloneRow(src));
    rows.splice(at + 1, 0, ...copies);
    if (cur && cur.r > at) cur.r += n;
    selectedRow = at + n;
    recomputeAndRefresh();
    renderWindow();
  });
  $("btn-del").addEventListener("click", async () => {
    if (selectedRow < 0) { alert("请先点击行号选中要删除的行。"); return; }
    const n = await askRowCount("删除行数", 1);
    if (n === null) return;
    const cnt = Math.min(n, rows.length - selectedRow);
    if (cnt < 1) return;
    if (!confirm(`确定删除从第 ${selectedRow + 1} 行起的 ${cnt} 行？`)) return;
    rows.splice(selectedRow, cnt);
    selectedRow = -1;
    clampCur();
    recomputeAndRefresh();
    renderWindow();
  });
  $("btn-reset").addEventListener("click", () => {
    if (!confirm("清空测点表格（保留危害因素库与检测项目）？")) return;
    emptyGrid(20);
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
    return INPUT_COLS.includes(col) || MANUAL_COLS.includes(col);
  }

  function findTd(r, cIdx) {
    return gridBody.querySelector(`tr[data-r="${r}"] > td[data-c="${ALL_COLS[cIdx]}"]`);
  }

  function getCellModelValue(r, cIdx) {
    if (!rows[r]) return "";
    const col = ALL_COLS[cIdx];
    if (col === "V") return "";
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

  function updateSelectionClasses() {
    const rect = selRect();
    for (const tr of gridBody.querySelectorAll("tr[data-r]")) {
      const r = Number(tr.dataset.r);
      for (const td of tr.querySelectorAll("td[data-c]")) {
        const c = ALL_COLS.indexOf(td.dataset.c);
        const inSel = rect && r >= rect.r1 && r <= rect.r2 && c >= rect.c1 && c <= rect.c2;
        td.classList.toggle("sel", inSel);
        td.classList.toggle("cur", !!(cur && r === cur.r && c === cur.c));
      }
    }
  }

  // 把当前单元格滚动到可见区域（虚拟滚动下按需重渲染）
  function revealCell(r, cIdx) {
    const wrap = gridWrap;
    const headH = gridHead.offsetHeight || 60; // sticky 表头高度
    const rowTop = headH + r * ROW_H;
    const ch = wrap.clientHeight;
    if (rowTop < wrap.scrollTop + headH) wrap.scrollTop = Math.max(0, rowTop - headH);
    else if (rowTop + ROW_H > wrap.scrollTop + ch) wrap.scrollTop = rowTop + ROW_H - ch;
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
      else while (rows.length < 5000 && r > rows.length - 1) rows.push(blankRow());
    }
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
    const td = findTd(cur.r, cur.c);
    if (td) {
      const el = td.querySelector("input");
      if (el) el.value = restore;
    }
    editing = false;
    editOriginal = null;
    updateVisibleCells();
  }

  function clearCell(r, cIdx) {
    const col = ALL_COLS[cIdx];
    const row = rows[r];
    if (INPUT_COLS.includes(col)) row.input[col] = "";
    else if (MANUAL_COLS.includes(col)) row.manual[col] = "";
    else if (OVERRIDE_COLS.includes(col)) {
      row.values[col] = "";
      delete row.overridden[col];
    } else return;
    editing = false;
    editOriginal = null;
    recomputeAndRefresh();
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
      if (col === "V") continue;
      if (INPUT_COLS.includes(col) && (row.input[col] ?? "") !== "") return c;
      if (MANUAL_COLS.includes(col) && (row.manual[col] ?? "") !== "") return c;
      if ((row.values[col] ?? "") !== "") return c;
    }
    return INPUT_COLS.length - 1;
  }

  gridBody.addEventListener("mousedown", (e) => {
    const td = e.target.closest("td[data-c]");
    if (!td) return;
    const r = Number(td.closest("tr").dataset.r);
    const c = ALL_COLS.indexOf(td.dataset.c);
    const cell = { r, c };
    lastMouse = { x: e.clientX, y: e.clientY };
    editing = false;
    editOriginal = null;
    if (e.shiftKey) {
      if (!cur) selAnchor = cell;
      else if (!selAnchor) selAnchor = cur;
      selStart = selAnchor;
      selEnd = cell;
    } else {
      selAnchor = cell;
      selStart = cell;
      selEnd = cell;
    }
    cur = { r, c };
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
      if (["c", "v", "x", "a"].includes(key.toLowerCase())) return; // 原生复制/粘贴/剪切/全选
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
      const page = Math.max(1, Math.floor(gridWrap.clientHeight / ROW_H) - 1);
      moveCur(page, 0, { extend: shift });
      return;
    }
    if (key === "PageUp") {
      e.preventDefault();
      const page = Math.max(1, Math.floor(gridWrap.clientHeight / ROW_H) - 1);
      moveCur(-page, 0, { extend: shift });
      return;
    }
    if ((key === "Delete" || key === "Backspace") && editableCellAt(r, c)) {
      e.preventDefault();
      clearCell(r, c);
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
    // 直接输入字符：替换当前单元格内容
    if (isInput && editableCellAt(r, c) && key.length === 1) {
      e.preventDefault();
      const el = e.target;
      editOriginal = getCellModelValue(r, c);
      el.value = key;
      editing = true;
      el.dispatchEvent(new Event("input", { bubbles: true }));
      try { el.setSelectionRange(el.value.length, el.value.length); } catch {}
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
          if (col === "V") v = "";
          else if (INPUT_COLS.includes(col)) v = rows[r].input[col];
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
    const rect = selRect();
    if (!rect) return;
    const text = e.clipboardData ? e.clipboardData.getData("text/plain") : "";
    if (!text) return;
    const lines = String(text).replace(/\r\n/g, "\n").replace(/\r/g, "\n").split("\n");
    let changed = false;
    for (let i = 0; i < lines.length; i++) {
      const cells = lines[i].split("\t");
      for (let j = 0; j < cells.length; j++) {
        const r = rect.r1 + i;
        const c = rect.c1 + j;
        if (c >= ALL_COLS.length) break;
        const col = ALL_COLS[c];
        if (!INPUT_COLS.includes(col)) continue; // 仅写入录入区
        while (rows.length <= r) rows.push(blankRow());
        rows[r].input[col] = cells[j];
        changed = true;
      }
    }
    if (changed) {
      e.preventDefault();
      cur = { r: rect.r1, c: rect.c1 };
      selAnchor = { r: rect.r1, c: rect.c1 };
      recomputeAndRefresh();
      renderWindow();
      updateSelectionClasses();
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
    }
    selStart = selEnd = selAnchor = null;
    cur = null;
    recomputeAndRefresh();
    renderWindow();
  });

  // ---------- 数据记录存储：本地服务 / GitHub API / 临时模式 ----------
  const RECORDS_KEY = "samplingPlanRecords_v1"; // 仅用于“临时模式”兜底
  const GH_CONFIG_KEY = "samplingPlanGithubConfig_v1";
  const GH_PATH_DEFAULT = "sampling-plan-app/data/records.json";
  let storageMode = "detecting";
  let ghSha = null;

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

  // 将最新数据镜像到本地（本地文件或浏览器），作为离线兜底
  function mirrorToLocal(list) {
    try {
      if (storageMode === "server") {
        fetch("/api/records", {
          method: "PUT",
          headers: { "Content-Type": "application/json" },
          body: JSON.stringify(list),
        }).catch(() => {});
      } else {
        localStorage.setItem(RECORDS_KEY, JSON.stringify(list));
      }
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

  async function loadRecords() {
    await ensureMode();
    // 配置了 GitHub：优先从 GitHub 读取，实现“打开即同步”
    if (hasGithubConfig()) {
      setSyncStatus("syncing");
      try {
        const list = await githubLoad();
        mirrorToLocal(list);
        setSyncStatus("ok");
        return list;
      } catch (e) {
        setSyncStatus("err", e.message);
      }
    }
    if (storageMode === "server") {
      try {
        const res = await fetch("/api/records", { cache: "no-store" });
        if (!res.ok) throw new Error("HTTP " + res.status);
        const list = await res.json();
        return Array.isArray(list) ? list : [];
      } catch {
        setStorageNotice(true);
        return localLoad();
      }
    }
    if (storageMode === "github") {
      try {
        return await githubLoad();
      } catch (e) {
        alert("读取 GitHub 数据失败：" + e.message);
        return [];
      }
    }
    setStorageNotice(true, noticeHtml());
    return localLoad();
  }

  async function persistRecords(list) {
    await ensureMode();
    // 配置了 GitHub：优先写回 GitHub，随后镜像到本地
    if (hasGithubConfig()) {
      setSyncStatus("syncing");
      try {
        const ok = await githubPersist(list);
        mirrorToLocal(list);
        setSyncStatus("ok");
        return ok;
      } catch (e) {
        setSyncStatus("err", e.message);
        alert("保存到 GitHub 失败：" + e.message + "，已改存本地。");
      }
    }
    if (storageMode === "server") {
      try {
        const res = await fetch("/api/records", {
          method: "PUT",
          headers: { "Content-Type": "application/json" },
          body: JSON.stringify(list),
        });
        if (!res.ok) throw new Error("HTTP " + res.status);
        return true;
      } catch (e) {
        setStorageNotice(true);
        alert("未能保存到项目文件夹，已暂存到浏览器：" + e.message);
        return localPersist(list);
      }
    }
    if (storageMode === "github") {
      try {
        return await githubPersist(list);
      } catch (e) {
        alert("保存到 GitHub 失败：" + e.message);
        return false;
      }
    }
    setStorageNotice(true, noticeHtml());
    alert('当前为静态网页模式且未配置 GitHub，数据仅暂存到浏览器。请在「GitHub 配置」中填写仓库与 Token 后保存。');
    return localPersist(list);
  }

  function localLoad() {
    try {
      const s = localStorage.getItem(RECORDS_KEY);
      return s ? JSON.parse(s) : [];
    } catch {
      return [];
    }
  }

  function localPersist(list) {
    try {
      localStorage.setItem(RECORDS_KEY, JSON.stringify(list));
      return true;
    } catch (e) {
      alert("保存失败：" + e.message);
      return false;
    }
  }

  // GitHub API 读写（静态网页部署模式）
  function githubApiUrl(cfg) {
    const p = (cfg.path || GH_PATH_DEFAULT).split("/").map(encodeURIComponent).join("/");
    const base = `https://api.github.com/repos/${cfg.repo}/contents/${p}`;
    return cfg.branch ? `${base}?ref=${encodeURIComponent(cfg.branch)}` : base;
  }

  function githubRequest(cfg, method, body) {
    const ctrl = new AbortController();
    const timer = setTimeout(() => ctrl.abort(), 20000);
    return fetch(githubApiUrl(cfg), {
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

  async function githubLoad() {
    const cfg = loadGithubConfig();
    if (!cfg || !cfg.repo || !cfg.token) throw new Error("未配置 GitHub 仓库或 Token");
    const res = await githubRequest(cfg, "GET");
    if (res.status === 404) {
      ghSha = null;
      return [];
    }
    if (!res.ok) throw new Error("HTTP " + res.status + "（请检查仓库名与 Token 权限）");
    const data = await res.json();
    ghSha = data.sha;
    const list = JSON.parse(L.decodeUnicodeBase64(data.content));
    return Array.isArray(list) ? list : [];
  }

  async function githubPersist(list) {
    const cfg = loadGithubConfig();
    if (!cfg || !cfg.repo || !cfg.token) throw new Error("未配置 GitHub 仓库或 Token");
    const content = L.encodeUnicodeBase64(JSON.stringify(list, null, 2));
    for (let attempt = 0; attempt < 2; attempt++) {
      if (!ghSha) {
        const res = await githubRequest(cfg, "GET");
        if (res.status === 404) ghSha = null;
        else if (res.ok) ghSha = (await res.json()).sha;
        else throw new Error("HTTP " + res.status);
      }
      const payload = {
        message: "更新数据记录（采样计划软件）",
        content,
        ...(ghSha ? { sha: ghSha } : {}),
        ...(cfg.branch ? { branch: cfg.branch } : {}),
      };
      const res = await githubRequest(cfg, "PUT", payload);
      if (res.status === 409 && attempt === 0) {
        ghSha = null; // 文件被其他端修改，重取 sha 后重试一次
        continue;
      }
      if (!res.ok) {
        const err = await res.json().catch(() => ({}));
        throw new Error("HTTP " + res.status + "：" + (err.message || "请检查 Token 的 Contents 读/写权限"));
      }
      ghSha = (await res.json()).content?.sha || ghSha;
      return true;
    }
    throw new Error("更新冲突，请稍后重试");
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
    return [
      rec.name,
      ...rec.rows.map((r) => [r.input.A, r.input.B, r.input.C, r.input.D, r.values.AN].filter((v) => v).join(" ")),
    ].join(" ").toLowerCase();
  }

  $("btn-save").addEventListener("click", async () => {
    const contentRows = trimBlankRows(rows);
    if (!contentRows.length) { alert("当前没有可保存的数据。"); return; }
    const name = await askInput({
      title: "保存数据",
      hint: "为当前数据命名，之后可在「数据记录」中搜索并调用，减少重复输入",
      value: defaultRecordName(),
    });
    if (name === null) return;
    if (name === "") { alert("名称不能为空。"); return; }
    const list = await loadRecords();
    const now = new Date().toISOString();
    const snap = L.snapshotRows(contentRows);
    const existing = list.find((r) => r.name === name);
    if (existing) {
      if (!confirm(`已存在同名记录「${name}」，是否覆盖？`)) return;
      existing.rows = snap;
      existing.updatedAt = now;
    } else {
      list.unshift({
        id: Date.now() + "_" + Math.random().toString(36).slice(2, 8),
        name,
        createdAt: now,
        updatedAt: now,
        rows: snap,
      });
    }
    if (await persistRecords(list)) alert(`已保存「${name}」（${contentRows.length} 行）。`);
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
        const preview = r.rows
          .slice(0, 3)
          .map((row) => `${row.input.A || "?"}｜${row.input.D || ""}`)
          .join("；");
        return (
          `<div class="db-item" data-id="${escAttr(r.id)}">` +
          `<div class="db-item-main">` +
          `<div class="db-item-name">${escHtml(r.name)}</div>` +
          `<div class="db-item-meta">${r.rows.length} 行 · 保存于 ${escHtml(meta)}</div>` +
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
    const rec = (await loadRecords()).find((r) => r.id === id);
    if (!rec) return;
    if (btn.dataset.act === "load") {
      const hasData = rows.some((r) => !isBlankRow(r));
      if (hasData && !confirm(`将用「${rec.name}」（${rec.rows.length} 行）替换当前表格，是否继续？`)) return;
      rows = L.restoreRows(rec.rows);
      selectedRow = -1;
      recomputeAndRefresh();
      renderWindow();
      $("db-modal").classList.add("hidden");
      alert(`已调用「${rec.name}」（${rec.rows.length} 行）。`);
    } else if (btn.dataset.act === "del") {
      if (!confirm(`确定删除记录「${rec.name}」？此操作不可恢复。`)) return;
      const list = (await loadRecords()).filter((r) => r.id !== id);
      if (await persistRecords(list)) {
        await renderDbList();
        alert("已删除。");
      }
    }
  });

  // 备份导出 / 导入
  $("db-export").addEventListener("click", async () => {
    const list = await loadRecords();
    if (!list.length) { alert("数据库为空，无需备份。"); return; }
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
      const map = new Map((await loadRecords()).map((r) => [r.name, r]));
      for (const r of clean) map.set(r.name, r);
      if (await persistRecords([...map.values()])) {
        await renderDbList();
        alert(`已导入 ${clean.length} 条记录。`);
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
      if (activeTab === "items") renderItems();
      if (activeTab === "main") renderWindow();
    });
  });

  // ---------- 危害因素库 ----------
  const HAZARD_HEADERS = ["识别", "系统名称", "粉尘性质", "定性分析", "委外检测", "计算TWA", "计算STEL", "计算CPE", "计算MAC", "结果保留位数", "存在高毒物品", "不检测原因说明"];
  const HAZARD_KEYS = ["rec", "name", "dust", "qual", "outsource", "twa", "stel", "cpe", "mac", "digits", "highTox", "noTestReason"];
  const YESNO_COLS = new Set([3, 4, 5, 6, 7, 8, 10]);

  function blankHazard() {
    const h = {};
    for (const k of HAZARD_KEYS) h[k] = "";
    return h;
  }

  function renderHazardHead() {
    $("hazard-head").innerHTML =
      `<tr>${HAZARD_HEADERS.map((h) => `<th>${h}</th>`).join("")}</tr>`;
  }

  function renderHazard() {
    const q = $("hazard-search").value.trim();
    const list = hazardFactors.filter((h) => !q || h.rec.includes(q) || h.name.includes(q));
    const body = $("hazard-body");
    let html = "";
    list.forEach((h, idx) => {
      const absIdx = hazardFactors.indexOf(h);
      html += `<tr data-h="${absIdx}">` +
        `<td class="rowno${absIdx === selectedHazard ? " selected" : ""}" data-h="${absIdx}">${absIdx + 1}</td>` +
        HAZARD_KEYS.map((k, ci) => {
          const val = fmt(h[k]);
          let inner;
          if (ci === 2) inner = `<select data-k="${k}">${optionList(["", "总尘", "呼尘"], val)}</select>`;
          else if (YESNO_COLS.has(ci)) inner = `<select data-k="${k}">${optionList(["", "是", "否"], val)}</select>`;
          else if (ci === 9) inner = `<input data-k="${k}" type="number" value="${escAttr(val)}">`;
          else inner = `<input data-k="${k}" value="${escAttr(val)}">`;
          return `<td class="${ci === 9 ? "cell-num" : ""}">${inner}</td>`;
        }).join("") + `</tr>`;
    });
    body.innerHTML = html;
    $("hazard-status").textContent = `共 ${hazardFactors.length} 条 · 显示 ${list.length} 条`;
    scheduleWidths();
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
    const td = e.target.closest("td.rowno");
    if (!td) return;
    selectedHazard = Number(td.dataset.h);
    renderHazard();
  });
  $("hazard-search").addEventListener("input", renderHazard);
  $("hazard-add").addEventListener("click", () => {
    hazardFactors.push(blankHazard());
    selectedHazard = hazardFactors.length - 1;
    renderHazard();
    onHazardChange();
  });
  $("hazard-del").addEventListener("click", () => {
    if (selectedHazard < 0) { alert("请先点击行号选中要删除的因素。"); return; }
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
  }

  // ---------- 检测项目 ----------
  function renderItems() {
    const q = $("items-search").value.trim();
    const dups = new Set(L.findDuplicates(detectionItems));
    const list = detectionItems.filter((it) => !q || it.includes(q));
    $("items-list").innerHTML = list
      .map((it) => `<div class="item${dups.has(it) ? " dup" : ""}">${escHtml(it)}${dups.has(it) ? "（重复）" : ""}</div>`)
      .join("");
    $("items-status").textContent = `共 ${detectionItems.length} 项 · 显示 ${list.length} 项${dups.size ? " · 重复 " + dups.size + " 项" : ""}`;
  }
  $("items-search").addEventListener("input", renderItems);

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
  }

  // ---------- 导出 ----------
  function exportWorkbook() {
    // 仅导出自动计算区（W~BI 共 39 列）
    const mainRows = [COMPUTED_HEADERS];
    for (const r of rows) {
      const line = [];
      for (const c of COMPUTED_COLS) {
        if (MANUAL_COLS.includes(c)) line.push(r.manual[c]);
        else line.push(NUM_COLS.has(c) ? L.toNum(r.values[c]) : r.values[c]);
      }
      mainRows.push(line);
    }
    return X.writeWorkbook({ mainRows, mainWidths: COMPUTED_WIDTHS });
  }

  $("btn-export").addEventListener("click", async () => {
    const { total } = L.countErrors(rows);
    if (total > 0 && !confirm(`当前有 ${total} 处校验错误，仍要导出吗？`)) return;
    const bytes = await exportWorkbook();
    const blob = new Blob([bytes], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" });
    X.downloadBlob(blob, "系统测点布局调查_自动计算区.xlsx");
  });

  $("btn-csv").addEventListener("click", () => {
    const sep = ",";
    const q = (v) => {
      const s = fmt(v);
      return /[",\n\r]/.test(s) ? `"${s.replace(/"/g, '""')}"` : s;
    };
    const lines = [COMPUTED_HEADERS.map(q).join(sep)];
    for (const r of rows) {
      const line = [];
      for (const c of COMPUTED_COLS) {
        if (MANUAL_COLS.includes(c)) line.push(r.manual[c]);
        else line.push(r.values[c]);
      }
      lines.push(line.map(q).join(sep));
    }
    const blob = new Blob(["\ufeff" + lines.join("\r\n")], { type: "text/csv;charset=utf-8" });
    X.downloadBlob(blob, "系统测点布局调查_自动计算区.csv");
  });

  // ---------- 导入 ----------
  const HEADER_ALIAS = {
    识别: "rec", 系统名称: "name", 职业卫生检测管理系统: "name", 粉尘性质: "dust",
    定性分析: "qual", 委外检测: "outsource", 计算TWA: "twa", 计算STEL: "stel",
    计算CPE: "cpe", 计算MAC: "mac", 结果保留位数: "digits", 存在高毒物品: "highTox",
    不检测原因说明: "noTestReason",
  };

  $("btn-import").addEventListener("click", () => $("file-input").click());
  $("file-input").addEventListener("change", async (e) => {
    const file = e.target.files[0];
    if (!file) return;
    try {
      const sheets = await X.readWorkbook(await file.arrayBuffer());
      const byName = {};
      for (const s of sheets) byName[s.name] = X.sheetToArray(s);

      // 危害因素库
      if (byName["危害因素"]) {
        const arr = byName["危害因素"];
        const hdr = arr[0] || [];
        const map = hdr.map((h) => HEADER_ALIAS[String(h).trim()] || null);
        const lib = arr.slice(1).filter((r) => r.some((v) => v !== "" && v !== null && v !== undefined));
        const parsed = lib.map((r) => {
          const h = blankHazard();
          map.forEach((k, i) => { if (k) h[k] = r[i] ?? ""; });
          return h;
        });
        if (parsed.length) hazardFactors = parsed;
      }

      // 检测项目
      if (byName["检测项目"]) {
        const arr = byName["检测项目"];
        const items = arr.slice(1).map((r) => r[0]).filter((v) => v !== "" && v !== null && v !== undefined);
        if (items.length) detectionItems = items.map(fmt);
      }

      // 测点布局
      if (byName["测点布局情况调查"]) {
        const arr = byName["测点布局情况调查"];
        const hdr = arr[0] || [];
        const idx = {};
        hdr.forEach((h, i) => { if (h !== null && h !== undefined && h !== "") idx[String(h).trim()] = i; });
        // 录入区独有的表头（自动计算区不含这些名称），用于区分“完整结构”与“仅自动计算区”文件
        const UNIQUE_INPUT_HEADERS = ["车间", "接害因素", "接触时间h/d", "人数"];
        const hasInput = UNIQUE_INPUT_HEADERS.some((h) => h in idx);
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
          }
          for (const c of MANUAL_COLS) {
            const h = HEADERS[colIdx(c)];
            if (h in idx) row.manual[c] = fmt(r[idx[h]]);
          }
          for (const c of COMPUTED_COLS) {
            const h = HEADERS[colIdx(c)];
            if (h in idx) vals[c] = fmt(r[idx[h]]);
          }
          imported.push(row);
          importedVals.push(vals);
        }
        if (imported.length) {
          rows = imported;
          L.computeRows(rows, { hazardFactors, detectionItems });
          rows.forEach((row, i) => {
            const cols = hasInput ? OVERRIDE_COLS : COMPUTED_COLS;
            for (const c of cols) {
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
          selectedRow = -1;
        }
      }
      onHazardChange();
      rebuildDatalist();
      renderItems();
      renderWindow();
      activeTab = "main";
      document.querySelectorAll(".tab").forEach((t) => t.classList.toggle("active", t.dataset.tab === "main"));
      document.querySelectorAll(".tab-panel").forEach((p) => p.classList.toggle("active", p.id === "tab-main"));
      alert(`导入成功：测点 ${rows.length} 行。${byName["危害因素"] ? `危害因素 ${hazardFactors.length} 条，` : "危害因素库保持现有，"}${byName["检测项目"] ? `检测项目 ${detectionItems.length} 项。` : "检测项目保持现有。"}`);
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
    loadSample();
    buildHead();
    renderHazardHead();
    rebuildDatalist();
    await detectStorageMode();
    setStorageNotice(storageMode !== "server", noticeHtml());
    if (hasGithubConfig()) {
      loadRecords().catch(() => {}); // 打开时自动从 GitHub 同步
    }
    L.computeRows(rows, { hazardFactors, detectionItems });
    $("hazard-count").textContent = hazardFactors.length;
    $("items-count").textContent = detectionItems.length;
    renderItems();
    applyGridWidths();
    applyHazardWidths();
    renderWindow();
  }
  init();
})();
