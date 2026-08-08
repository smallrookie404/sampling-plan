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
    const fieldTh = (label, c) =>
      `<th class="field" style="width:${COL_W[colIdx(c)]}ch">${label || "&nbsp;"}</th>`;
    let groupHtml = `<th class="corner" rowspan="2" style="width:40px">行</th>`;
    groupHtml += groupTh("录 入 区", INPUT_COLS.length);
    groupHtml += groupTh("&nbsp;", 1);
    groupHtml += groupTh("自动计算区（与原表 W~BI 列一致）", COMPUTED_COLS.length);
    let fieldHtml = "";
    for (const c of ALL_COLS) fieldHtml += fieldTh(HEADERS[colIdx(c)], c);
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
    recomputeAndRefresh();
    renderWindow();
  });
  $("btn-reset").addEventListener("click", () => {
    if (!confirm("清空测点表格（保留危害因素库与检测项目）？")) return;
    emptyGrid(20);
    selectedRow = -1;
    recomputeAndRefresh();
    renderWindow();
  });

  // ---------- 数据记录（保存到项目文件夹 data/records.json） ----------
  const RECORDS_KEY = "samplingPlanRecords_v1"; // 仅用于“临时模式”兜底
  const serverMode = window.location.protocol === "http:" || window.location.protocol === "https:";

  function setStorageNotice(show) {
    const el = $("storage-notice");
    if (el) el.classList.toggle("hidden", !show);
  }

  async function loadRecords() {
    if (serverMode) {
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
    setStorageNotice(true);
    return localLoad();
  }

  async function persistRecords(list) {
    if (serverMode) {
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
    setStorageNotice(true);
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
      `<tr>${HAZARD_HEADERS.map((h, i) => `<th style="width:${i === 0 || i === 1 ? 22 : 12}ch">${h}</th>`).join("")}</tr>`;
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
  function init() {
    hazardFactors = DEFAULT_DATA.hazardFactors.map((h) => ({ ...h }));
    detectionItems = DEFAULT_DATA.detectionItems.slice();
    loadSample();
    buildHead();
    renderHazardHead();
    rebuildDatalist();
    if (!serverMode) setStorageNotice(true);
    L.computeRows(rows, { hazardFactors, detectionItems });
    $("hazard-count").textContent = hazardFactors.length;
    $("items-count").textContent = detectionItems.length;
    renderItems();
    renderWindow();
  }
  init();
})();
