  // ---------- 调查表（现场调查导入模板中的 6 个工作表，按原格式） ----------
  // 自包含模块：通过 window.SurveySheets.refresh() 在页签切换时渲染
  const $S = (id) => document.getElementById(id);
  const esc = (v) => (v === null || v === undefined ? "" : String(v));
  const escHtml = (v) => esc(v).replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;");
  const escAttr = escHtml;
  const SURVEY_SHEETS = [
    { key: "process", name: "生产工艺调查", headers: ["*工艺名称", "*单元/工作场所", "*工艺描述"] },
    {
      key: "equipment", name: "设备设施调查",
      headers: ["*设备名称", "*设备型号及规格", "*设备总数", "*设备运行数", "单元/工作场所", "操作岗位(工种)", "工作地点", "*设施布局", "备注"],
    },
    {
      key: "material", name: "原辅物料调查",
      headers: ["*物料名称", "单元/工作场所", "使用岗位(工种)", "使用地点", "*物理状态", "储存量", "*年用量", "规格", "*主要成分", "*储存方式", "*加药/投料方式", "*运输方式", "*装卸方式", "*装卸周期", "备注"],
    },
    {
      key: "product", name: "主要产品调查",
      headers: ["*类型", "*产品名称", "单元/工作场所", "影响岗位(工种)", "*产量/产值", "*物理状态", "*主要成分", "*包装方式", "性状", "储存方式"],
    },
    {
      key: "protect", name: "职业防护调查",
      headers: ["*防护类型", "*防护设施名称", "单元/工作场所", "设置岗位(工种)", "设置地点", "*设施总数", "*设施运行数", "*设施运行情况", "维修情况", "技术参数", "备注"],
    },
    {
      key: "ppe", name: "个体防护调查",
      headers: ["*防护用品分类", "*防护用品类别", "单元/工作场所", "配置岗位(工种)", "*型号或规格", "*生产厂家", "*佩戴情况", "*更换周期", "*防护性能参数", "备注"],
    },
  ];
  const SURVEY_KEY = "samplingPlanSurvey_v2"; // v2：统一默认 10 行（旧 v1 缓存弃用）
  let surveyData = null; // { key: [[cell,...],...] }
  let surveyCur = "process"; // 当前显示的子表
  // 选区状态（与主表格一致：cur 当前单元格 + 拖选矩形）
  let sCur = null; // { r, c }
  let sSelStart = null;
  let sSelEnd = null;
  let sSelAnchor = null; // 选区锚点（Shift+方向扩展用，与主表格 selAnchor 同义）
  let sDragging = false;

  function sSelRect() {
    if (!sSelStart || !sSelEnd) return null;
    return {
      r1: Math.min(sSelStart.r, sSelEnd.r),
      r2: Math.max(sSelStart.r, sSelEnd.r),
      c1: Math.min(sSelStart.c, sSelEnd.c),
      c2: Math.max(sSelStart.c, sSelEnd.c),
    };
  }

  // 选中样式同步（与主表格 updateSelectionClasses 同语义：sel 选区 / cur 当前格 / hl-row hl-col 行列高亮）
  function updateSurveySelection() {
    const rect = sSelRect();
    const multi = rect && (rect.r1 !== rect.r2 || rect.c1 !== rect.c2);
    for (const tr of $S("survey-body").querySelectorAll("tr[data-r]")) {
      const r = Number(tr.dataset.r);
      for (const td of tr.querySelectorAll("td[data-c]")) {
        const c = Number(td.dataset.c);
        const inSel = rect && r >= rect.r1 && r <= rect.r2 && c >= rect.c1 && c <= rect.c2;
        td.classList.toggle("sel", inSel);
        td.classList.toggle("cur", !!(sCur && r === sCur.r && c === sCur.c));
        td.classList.toggle("hl-row", !multi && !!sCur && r === sCur.r && c !== sCur.c);
        td.classList.toggle("hl-col", !multi && !!sCur && c === sCur.c && r !== sCur.r);
      }
    }
  }

  function surveyLoad() {
    if (surveyData) return surveyData;
    try {
      const s = localStorage.getItem(SURVEY_KEY);
      surveyData = s ? JSON.parse(s) : {};
    } catch {
      surveyData = {};
    }
    for (const sh of SURVEY_SHEETS) {
      // 默认 10 行空白行（仅当该表尚无任何数据时填充）
      if (!Array.isArray(surveyData[sh.key]) || surveyData[sh.key].length === 0) {
        surveyData[sh.key] = Array.from({ length: 10 }, () => sh.headers.map(() => ""));
      }
    }
    return surveyData;
  }

  function surveySave() {
    try { localStorage.setItem(SURVEY_KEY, JSON.stringify(surveyData)); } catch {}
  }

  function surveyRows() {
    return surveyLoad()[surveyCur];
  }

  function surveyBlankRow() {
    const sh = SURVEY_SHEETS.find((s) => s.key === surveyCur);
    return sh.headers.map(() => "");
  }

  function buildSurveyTabs() {
    $S("survey-tabs").innerHTML = SURVEY_SHEETS.map((s) =>
      `<button class="stab${s.key === surveyCur ? " active" : ""}" data-k="${s.key}">${escHtml(s.name)}</button>`
    ).join("");
  }

  function renderSurvey() {
    sCloseWsPanel(); // tbody 重建会销毁编辑框，下拉面板一并清理
    const sh = SURVEY_SHEETS.find((s) => s.key === surveyCur);
    const rows = surveyRows();
    // 表头（模板原格式：必填列带 *，样式上以浅红底提示）
    $S("survey-head").innerHTML =
      `<tr>${sh.headers.map((h, i) =>
        `<th class="${h.startsWith("*") ? "req" : ""}" data-i="${i}">${escHtml(h)}</th>`
      ).join("")}</tr>`;
    // colgroup：首列稍窄，其余均分
    $S("survey-cols").innerHTML = sh.headers.map((h, i) => `<col style="width:${i === 0 ? 150 : 140}px">`).join("");
    // 全量渲染（调查表行数有限，无需虚拟化）；不显示序号列
    $S("survey-body").innerHTML = rows.length
      ? rows.map((r, ri) =>
          `<tr data-r="${ri}">` +
          sh.headers.map((h, ci) =>
            `<td data-c="${ci}"><div class="ctext">${escHtml(r[ci])}</div></td>`
          ).join("") +
          `</tr>`
        ).join("")
      : `<tr class="empty-row"><td colspan="${sh.headers.length}" style="text-align:center;color:#94a3b8;padding:16px">暂无数据，点击「+ 新增行」开始填写</td></tr>`;
    $S("survey-status").textContent = `${sh.name} · 共 ${rows.length} 行`;
    updateSurveySelection();
  }

  // 切换子表
  $S("survey-tabs").addEventListener("click", (e) => {
    const btn = e.target.closest(".stab");
    if (!btn) return;
    surveyCur = btn.dataset.k;
    sCur = sSelAnchor = sSelStart = sSelEnd = null;
    buildSurveyTabs();
    renderSurvey();
  });

  // ---------- Excel 式编辑（与主表格一致：单击选中 / 双击或直接输入编辑 / 键盘导航 / 复制粘贴） ----------
  let sEditing = false;
  let sEditOriginal = null;

  // 车间下拉面板（body 级 fixed 定位，不被表格裁剪）
  let sWsPanel = null;
  function sCloseWsPanel() {
    if (sWsPanel) { sWsPanel.remove(); sWsPanel = null; }
  }

  function sFindTd(r, c) {
    return $S("survey-body").querySelector(`tr[data-r="${r}"] > td[data-c="${c}"]`);
  }

  function sCellBeginEdit(commitCurrent) {
    if (!sCur || sEditing) return;
    const td = sFindTd(sCur.r, sCur.c);
    const rows = surveyRows();
    if (!td || !rows[sCur.r]) return;
    sEditing = true;
    sEditOriginal = rows[sCur.r][sCur.c] ?? "";
    // 「单元/工作场所」列（各调查表均含，列位置不同）：下拉选项 = 主表格已填车间名称（与主表格下拉联想同机制，可输可选）
    const sh = SURVEY_SHEETS.find((s) => s.key === surveyCur);
    const wsOpts = (sh && sh.headers[sCur.c] && sh.headers[sCur.c].includes("单元/工作场所") && window.SamplingApp && window.SamplingApp.workshopNames)
      ? window.SamplingApp.workshopNames()
      : null;
    if (wsOpts) {
      sCloseWsPanel();
      td.innerHTML = `<input data-r="${sCur.r}" data-c="${sCur.c}" value="${escAttr(sEditOriginal)}">`;
      const el = td.querySelector("input");
      el.focus();
      if (commitCurrent) el.value = commitCurrent + sEditOriginal;
      el.setSelectionRange(el.value.length, el.value.length);
      // 自定义下拉面板：始终显示全部选项（datalist 会按已有内容过滤，单元格有值时只剩匹配项）
      if (wsOpts.length) {
        const panel = document.createElement("div");
        panel.className = "survey-ws-panel";
        panel.innerHTML = wsOpts.map((o) => `<div class="survey-ws-opt" data-v="${escAttr(o)}">${escHtml(o)}</div>`).join("");
        document.body.appendChild(panel);
        const r = el.getBoundingClientRect();
        panel.style.left = r.left + "px";
        panel.style.top = r.bottom + 2 + "px";
        panel.style.minWidth = Math.max(r.width, 120) + "px";
        panel.addEventListener("mousedown", (ev) => {
          const opt = ev.target.closest(".survey-ws-opt");
          if (!opt) return;
          ev.preventDefault(); // 阻止 input 失焦
          el.value = opt.dataset.v;
          el.focus();
          el.setSelectionRange(el.value.length, el.value.length);
          sCloseWsPanel();
        });
        sWsPanel = panel;
      }
      return;
    }
    // 一律用 textarea：无论内容是否含换行，编辑态排版与显示态完全一致（单段长文本同样自动折行，不变成一行）
    // 编辑不改变行高：textarea 高度 = td 内容区高度（clientHeight 减上下内边距），替换前量好
    const cs = getComputedStyle(td);
    const h = td.clientHeight - (parseFloat(cs.paddingTop) || 0) - (parseFloat(cs.paddingBottom) || 0);
    td.innerHTML = `<textarea data-r="${sCur.r}" data-c="${sCur.c}">${escHtml(sEditOriginal)}</textarea>`;
    const ta = td.querySelector("textarea");
    ta.style.height = h + "px";
    // 字体度量差异可能导致内容溢出出现滚动条：以内容实际高度为准（不出现滚动条）
    requestAnimationFrame(() => {
      if (sEditing && ta.isConnected && ta.scrollHeight > ta.clientHeight) {
        // border-box 下边框占高，补偿后内容区才能完整容纳
        const bh = ta.offsetHeight - ta.clientHeight;
        ta.style.height = (ta.scrollHeight + bh) + "px";
      }
    });
    ta.style.overflowY = "hidden";
    ta.focus();
    if (commitCurrent) ta.value = commitCurrent + sEditOriginal;
    ta.setSelectionRange(ta.value.length, ta.value.length);
  }

  function sCellCommit() {
    if (!sEditing || !sCur) return;
    sCloseWsPanel();
    const td = sFindTd(sCur.r, sCur.c);
    const input = td && td.querySelector("input,textarea");
    const rows = surveyRows();
    if (input && rows[sCur.r]) {
      rows[sCur.r][sCur.c] = input.value;
      surveySave();
    }
    sEditing = false;
    sEditOriginal = null;
    if (td) td.innerHTML = `<div class="ctext">${escHtml(rows[sCur.r] ? (rows[sCur.r][sCur.c] ?? "") : "")}</div>`;
  }

  function sCellCancel() {
    if (!sEditing || !sCur) return;
    sCloseWsPanel();
    const td = sFindTd(sCur.r, sCur.c);
    sEditing = false;
    sEditOriginal = null;
    if (td) td.innerHTML = `<div class="ctext">${escHtml((surveyRows()[sCur.r] || [])[sCur.c] ?? "")}</div>`;
  }

  function sMove(dr, dc, opts = {}) {
    if (!sCur) return;
    const rows = surveyRows();
    const sh = SURVEY_SHEETS.find((s) => s.key === surveyCur);
    const maxC = sh.headers.length - 1;
    let r = sCur.r + dr;
    let c = sCur.c + dc;
    if (opts.wrap) {
      while (c > maxC) { c -= maxC + 1; r += 1; }
      while (c < 0) { c += maxC + 1; r -= 1; }
      if (r < 0) { r = 0; c = 0; }
    }
    if (opts.grow && r > rows.length - 1) {
      if (rows.length >= 5000) r = rows.length - 1;
      else {
        while (rows.length < 5000 && r > rows.length - 1) rows.push(surveyBlankRow());
        surveySave();
        renderSurvey(); // 新行需要渲染（内部会同步选区样式）
      }
    }
    r = Math.max(0, Math.min(r, rows.length - 1));
    c = Math.max(0, Math.min(c, maxC));
    const prev = sCur;
    sCur = { r, c };
    if (opts.extend && (sSelAnchor || prev)) {
      if (!sSelAnchor && prev) sSelAnchor = prev;
      sSelStart = sSelAnchor;
      sSelEnd = { r, c };
    } else {
      sSelAnchor = { r, c };
      sSelStart = { r, c };
      sSelEnd = { r, c };
    }
    updateSurveySelection();
    const td = sFindTd(r, c);
    if (td && td.scrollIntoView) td.scrollIntoView({ block: "nearest", inline: "nearest" });
  }

  function sMoveTo(r, c) {
    const rows = surveyRows();
    const sh = SURVEY_SHEETS.find((s) => s.key === surveyCur);
    r = Math.max(0, Math.min(r, rows.length - 1));
    c = Math.max(0, Math.min(c, sh.headers.length - 1));
    sCur = { r, c };
    sSelAnchor = sSelStart = sSelEnd = { r, c };
    updateSurveySelection();
    const td = sFindTd(r, c);
    if (td && td.scrollIntoView) td.scrollIntoView({ block: "nearest", inline: "nearest" });
  }

  // 单击选中；松开且未拖动时立即进入编辑（与主表格一致），光标落在点击位置
  let sLastMouse = null;
  let sDidDrag = false;

  // 估算点击位置在 input/textarea 值中的字符偏移（canvas 测宽）
  function sCaretOffset(el, x, y) {
    try {
      const r = el.getBoundingClientRect();
      const cs = getComputedStyle(el);
      const canvas = sCaretOffset._c || (sCaretOffset._c = document.createElement("canvas"));
      const ctx = canvas.getContext("2d");
      ctx.font = `${cs.fontWeight} ${cs.fontSize} ${cs.fontFamily}`;
      const px = x - r.left - (parseFloat(cs.borderLeftWidth) || 0) - (parseFloat(cs.paddingLeft) || 0) + el.scrollLeft;
      const colOf = (line) => {
        let acc = 0, w = 0;
        for (let i = 0; i < line.length; i++) {
          const cw = ctx.measureText(line[i]).width;
          if (acc + cw / 2 > px) return i;
          acc += cw; w = i + 1;
        }
        return line.length;
      };
      if (el.tagName === "TEXTAREA") {
        const lh = parseFloat(cs.lineHeight) || parseFloat(cs.fontSize) * 1.2 || 16;
        const py = y - r.top - (parseFloat(cs.borderTopWidth) || 0) - (parseFloat(cs.paddingTop) || 0) + el.scrollTop;
        const lines = el.value.split("\n");
        let li = Math.floor(py / lh);
        li = Math.max(0, Math.min(lines.length - 1, li));
        let off = 0;
        for (let k = 0; k < li; k++) off += lines[k].length + 1;
        return off + colOf(lines[li]);
      }
      return colOf(el.value);
    } catch {
      return el.value.length;
    }
  }

  $S("survey-body").addEventListener("mousedown", (e) => {
    if (e.button !== 0) return;
    const td = e.target.closest("td[data-c]");
    if (!td) return;
    const tr = td.closest("tr[data-r]");
    if (!tr) return;
    const r = Number(tr.dataset.r);
    const c = Number(td.dataset.c);
    if (sEditing) {
      // 编辑中点击当前格：保持编辑态，光标由浏览器原生定位（与主表格一致，不闪、不重建）
      if (sCur && r === sCur.r && c === sCur.c) return;
      sCellCommit(); // 点其他格：先提交再切换
    }
    sCur = { r, c };
    sSelAnchor = sSelStart = sSelEnd = { r, c };
    sDragging = true;
    sDidDrag = false;
    sLastMouse = { x: e.clientX, y: e.clientY };
    e.preventDefault();
    updateSurveySelection();
  });

  $S("survey-body").addEventListener("dblclick", (e) => {
    // 单击已进入编辑，双击无需重复处理
    e.preventDefault();
  });

  document.addEventListener("mousemove", (e) => {
    if (!sDragging) return;
    if (Math.abs(e.clientX - sLastMouse.x) > 3 || Math.abs(e.clientY - sLastMouse.y) > 3) sDidDrag = true;
    const td = e.target.closest && e.target.closest("td[data-c]");
    if (!td) return;
    const tr = td.closest("tr[data-r]");
    if (!tr) return;
    sSelEnd = { r: Number(tr.dataset.r), c: Number(td.dataset.c) };
    updateSurveySelection();
  });

  document.addEventListener("mouseup", () => {
    if (sDragging && !sDidDrag && sCur) {
      // 单击（未拖动）：立即进入编辑，光标落在点击位置
      const td = sFindTd(sCur.r, sCur.c);
      sCellBeginEdit();
      const el = td && td.querySelector("input,textarea");
      if (el && sLastMouse) {
        const pos = sCaretOffset(el, sLastMouse.x, sLastMouse.y);
        try { el.setSelectionRange(pos, pos); } catch {}
      }
    }
    sDragging = false;
  });

  // 键盘逻辑与主表格完全一致：IME 守卫 / Ctrl 组合键 / F2 / 编辑态 Enter·Tab·Esc / 导航（Shift 扩展、Tab 回绕、Enter 自动加行）/ Home·End / PageUp·Down / Delete 清空 / copy 事件复制
  // 挂 document：单元格选中后焦点不在表格内也能响应
  document.addEventListener("keydown", (e) => {
    if (!sCur) return;
    const ae = document.activeElement;
    // 焦点在其他输入框/文本域（如保存对话框）时不劫持按键
    if (ae && (ae.tagName === "INPUT" || ae.tagName === "TEXTAREA" || ae.tagName === "SELECT") && !$S("survey-body").contains(ae)) return;
    const mod = e.ctrlKey || e.metaKey;
    if (e.isComposing || e.key === "Process") return; // 中文输入法组合中不拦截
    // 编辑中的输入框获得焦点时进入编辑态分支
    const editInput = ae && (ae.tagName === "INPUT" || ae.tagName === "TEXTAREA") && $S("survey-body").contains(ae) ? ae : null;

    if (mod) {
      if (editInput) {
        // 编辑中保留浏览器原生快捷键；仅 Ctrl+Enter 提交并下移（textarea 换行与提交并存）
        if (e.key === "Enter") {
          e.preventDefault();
          sCellCommit();
          sMove(e.shiftKey ? -1 : 1, 0, { grow: true });
        }
        return;
      }
      if (e.key === "Home") { e.preventDefault(); sMoveTo(0, 0); return; }
      if (e.key === "End") {
        e.preventDefault();
        const rows = surveyRows();
        const sh = SURVEY_SHEETS.find((s) => s.key === surveyCur);
        sMoveTo(rows.length - 1, sh.headers.length - 1);
        return;
      }
      if (e.key.toLowerCase() === "a") {
        e.preventDefault();
        const rows = surveyRows();
        const sh = SURVEY_SHEETS.find((s) => s.key === surveyCur);
        sSelAnchor = { r: 0, c: 0 };
        sSelStart = { r: 0, c: 0 };
        sSelEnd = { r: rows.length - 1, c: sh.headers.length - 1 };
        updateSurveySelection();
        return;
      }
      if (["c", "v", "x"].includes(e.key.toLowerCase())) return; // 原生复制/粘贴/剪切（copy/paste 事件处理）
      return;
    }

    // F2：进入编辑（编辑中则光标移到末尾）
    if (e.key === "F2") {
      e.preventDefault();
      if (editInput) {
        try { editInput.setSelectionRange(editInput.value.length, editInput.value.length); } catch {}
      } else {
        sCellBeginEdit();
      }
      return;
    }

    // 编辑模式：Enter/Tab 提交并移动，Escape 还原，其余按键（方向键/Home/End/退格）走原生行为
    if (editInput) {
      if (e.key === "Enter") {
        // textarea 中 Enter 保留原生换行（Ctrl+Enter 提交下移）
        if (editInput.tagName === "TEXTAREA" && !mod) return;
        e.preventDefault();
        sCellCommit();
        sMove(e.shiftKey ? -1 : 1, 0, { grow: true });
      } else if (e.key === "Tab") {
        e.preventDefault();
        sCellCommit();
        sMove(0, e.shiftKey ? -1 : 1, { grow: true, wrap: true });
      } else if (e.key === "Escape") {
        e.preventDefault();
        sCellCancel();
      }
      return;
    }

    // 非编辑模式：单元格光标移动
    if (e.key === "Enter") { e.preventDefault(); sMove(e.shiftKey ? -1 : 1, 0, { grow: true }); return; }
    if (e.key === "Tab") { e.preventDefault(); sMove(0, e.shiftKey ? -1 : 1, { grow: true, wrap: true }); return; }
    if (e.key === "Escape") {
      sSelAnchor = sSelStart = sSelEnd = null;
      updateSurveySelection();
      return;
    }
    if (e.key === "ArrowDown") { e.preventDefault(); sMove(1, 0, { extend: e.shiftKey }); return; }
    if (e.key === "ArrowUp") { e.preventDefault(); sMove(-1, 0, { extend: e.shiftKey }); return; }
    if (e.key === "ArrowRight") { e.preventDefault(); sMove(0, 1, { extend: e.shiftKey }); return; }
    if (e.key === "ArrowLeft") { e.preventDefault(); sMove(0, -1, { extend: e.shiftKey }); return; }
    if (e.key === "Home") { e.preventDefault(); sMoveTo(sCur.r, 0); return; }
    if (e.key === "End") {
      e.preventDefault();
      const sh = SURVEY_SHEETS.find((s) => s.key === surveyCur);
      sMoveTo(sCur.r, sh.headers.length - 1);
      return;
    }
    if (e.key === "PageDown" || e.key === "PageUp") {
      e.preventDefault();
      const wrap = $S("survey-wrap");
      const td = sFindTd(sCur.r, sCur.c);
      const rh = Math.max(1, td ? td.getBoundingClientRect().height : 28);
      const page = Math.max(1, Math.floor(wrap.clientHeight / rh) - 1);
      sMove(e.key === "PageDown" ? page : -page, 0, { extend: e.shiftKey });
      return;
    }
    // Delete/Backspace：清空选区（或当前格）内容
    if (e.key === "Delete" || e.key === "Backspace") {
      e.preventDefault();
      const rect = sSelRect();
      const rows = surveyRows();
      const r1 = rect ? rect.r1 : sCur.r;
      const r2 = rect ? rect.r2 : sCur.r;
      const c1 = rect ? rect.c1 : sCur.c;
      const c2 = rect ? rect.c2 : sCur.c;
      for (let r = r1; r <= r2; r++) {
        if (!rows[r]) continue;
        for (let c = c1; c <= c2; c++) rows[r][c] = "";
      }
      surveySave();
      renderSurvey();
      return;
    }
    // 可打印字符直接进入编辑并追加（编辑框内继续输入走原生）
    if (e.key.length === 1 && !e.altKey) {
      e.preventDefault();
      sCellBeginEdit(e.key);
    }
  });

  // 复制：与主表格一致走 copy 事件（Ctrl+C 原生触发），选区内容以 TSV 写入剪贴板
  document.addEventListener("copy", (e) => {
    if (!sCur) return;
    const ae = document.activeElement;
    if (ae && (ae.tagName === "INPUT" || ae.tagName === "TEXTAREA") && $S("survey-body").contains(ae)) return; // 编辑中走原生复制
    const rect = sSelRect();
    if (!rect) return;
    const rows = surveyRows();
    const lines = [];
    for (let r = rect.r1; r <= rect.r2; r++) {
      const line = [];
      for (let c = rect.c1; c <= rect.c2; c++) line.push(rows[r] ? (rows[r][c] ?? "") : "");
      lines.push(line.join("\t"));
    }
    if (e.clipboardData) {
      e.clipboardData.setData("text/plain", lines.join("\n"));
      e.preventDefault();
    }
  });

  // 解析 Excel 粘贴文本：Tab 分列、换行分行；含换行/引号/Tab 的单元格会被 Excel 包在双引号内（"" 为字面引号）
  function parseTsvGrid(text) {
    const grid = [];
    let row = [];
    let cell = "";
    let inQuotes = false;
    for (let i = 0; i < text.length; i++) {
      const ch = text[i];
      if (inQuotes) {
        if (ch === '"') {
          if (text[i + 1] === '"') { cell += '"'; i++; } // "" → 字面引号
          else inQuotes = false;
        } else cell += ch;
      } else if (ch === '"') {
        inQuotes = true;
      } else if (ch === "\t") {
        row.push(cell); cell = "";
      } else if (ch === "\n") {
        row.push(cell); grid.push(row); row = []; cell = "";
      } else if (ch === "\r") {
        // 跳过（\r\n 或独立 \r 都由 \n 分支处理；行尾 \r 在 \n 前被忽略）
      } else {
        cell += ch;
      }
    }
    // 末尾单元格（文本不以换行结尾时）
    if (cell !== "" || row.length) { row.push(cell); grid.push(row); }
    return grid;
  }

  // 粘贴：Excel 复制的内容以 text/plain TSV 到达 paste 事件（比 clipboard.readText 兼容性更好、无需权限）
  // 挂 document：单元格选中后焦点不在表格内，paste 事件派发到 body，挂在 survey-body 上收不到
  // 粘贴基准 = 选区左上角；多行多列依次写入，行数不够自动扩行
  document.addEventListener("paste", (e) => {
    if (!sCur) return;
    // 焦点在其他输入框/文本域时不劫持（如保存对话框）
    const ae = document.activeElement;
    if (ae && (ae.tagName === "INPUT" || ae.tagName === "TEXTAREA" || ae.tagName === "SELECT") && !$S("survey-body").contains(ae)) return;
    const text = e.clipboardData && e.clipboardData.getData("text/plain");
    if (!text) return;
    e.preventDefault();
    const rows = surveyRows();
    const grid = parseTsvGrid(text);
    // Excel 复制区域末尾常带一个空行，去掉（中间空行保留）
    if (grid.length > 1 && grid[grid.length - 1].length === 1 && grid[grid.length - 1][0] === "") grid.pop();
    if (!grid.length) return;
    const r0 = sCur.r;
    const c0 = sCur.c;
    const need = r0 + grid.length;
    while (rows.length < need) rows.push(surveyBlankRow());
    for (let i = 0; i < grid.length; i++) {
      const cols = grid[i]; // parseTsvGrid 已按列拆好
      for (let j = 0; j < cols.length; j++) {
        if (c0 + j < rows[r0 + i].length) rows[r0 + i][c0 + j] = cols[j];
      }
    }
    surveySave();
    sEditing = false; // renderSurvey 重建 tbody 会销毁编辑框，编辑态标志必须同步复位
    sEditOriginal = null;
    renderSurvey();
  });

  // ---------- 行操作（与主表格一致：行数弹窗、选区默认行数、基准行取选区首行） ----------
  function askN(label, def) {
    if (window.SamplingApp && window.SamplingApp.askRowCount) return window.SamplingApp.askRowCount(label, def);
    // 兜底（SamplingApp 未就绪时）
    const raw = prompt(label + "：", String(def));
    if (raw === null) return null;
    const n = parseInt(raw, 10);
    if (!Number.isInteger(n) || n < 1) { alert("请输入大于 0 的整数行数。"); return null; }
    return Math.min(n, 1000);
  }

  function anchorRow() {
    // 与主表格一致：有选区用选区首行；否则用当前单元格所在行
    const rect = sSelRect();
    if (rect) return rect.r1;
    if (sCur) return sCur.r;
    return -1;
  }

  $S("survey-add").addEventListener("click", async () => {
    const n = await askN("新增行数", 1);
    if (n === null) return;
    const rows = surveyRows();
    if (rows.length + n > 5000) { alert("总行数不能超过 5000。"); return; }
    rows.push(...Array.from({ length: n }, () => surveyBlankRow()));
    surveySave();
    renderSurvey();
    const wrap = $S("survey-wrap");
    if (wrap) wrap.scrollTop = wrap.scrollHeight;
  });

  $S("survey-insert").addEventListener("click", async () => {
    const n = await askN("插入行数", 1);
    if (n === null) return;
    const rows = surveyRows();
    if (rows.length + n > 5000) { alert("总行数不能超过 5000。"); return; }
    const at = anchorRow() >= 0 ? anchorRow() + 1 : rows.length;
    rows.splice(at, 0, ...Array.from({ length: n }, () => surveyBlankRow()));
    if (sCur && sCur.r >= at) sCur = { ...sCur, r: sCur.r + n };
    surveySave();
    renderSurvey();
  });

  $S("survey-copy").addEventListener("click", async () => {
    const rect = sSelRect();
    const selRows = rect ? rect.r2 - rect.r1 + 1 : 1;
    const n = await askN("复制行数", selRows);
    if (n === null) return;
    const rows = surveyRows();
    const at = anchorRow();
    if (at < 0) return;
    if (rows.length + n > 5000) { alert("总行数不能超过 5000。"); return; }
    const src = rows[at];
    rows.splice(at + 1, 0, ...Array.from({ length: n }, () => src.slice()));
    if (sCur && sCur.r > at) sCur = { ...sCur, r: sCur.r + n };
    surveySave();
    renderSurvey();
  });

  // 清空录入区：6 个子表全部重置为默认 10 行空表
  $S("survey-clear").addEventListener("click", () => {
    if (!confirm("确定清空调查表全部 6 个表格的录入内容？此操作不可恢复。")) return;
    sCloseWsPanel();
    sEditing = false;
    sEditOriginal = null;
    sCur = sSelAnchor = sSelStart = sSelEnd = null;
    SURVEY_SHEETS.forEach((sh) => {
      surveyData[sh.key] = Array.from({ length: 10 }, () => sh.headers.map(() => ""));
    });
    surveySave();
    renderSurvey();
  });

  $S("survey-del").addEventListener("click", async () => {
    const rect = sSelRect();
    const selRows = rect ? rect.r2 - rect.r1 + 1 : 1;
    const delAt = anchorRow();
    if (delAt < 0) { alert("请先选中要删除的行（点击任意单元格）。"); return; }
    const n = await askN("删除行数", selRows);
    if (n === null) return;
    const cnt = Math.min(n, surveyRows().length - delAt);
    if (cnt < 1) return;
    if (!confirm(`确定删除从第 ${delAt + 1} 行起的 ${cnt} 行？`)) return;
    surveyRows().splice(delAt, cnt);
    sCur = null;
    sSelAnchor = sSelStart = sSelEnd = null;
    surveySave();
    renderSurvey();
  });

  // xlsx 导出辅助（供主界面导出菜单生成「调查表 6 子表合一份」xlsx）
  function sExportAoaSheet(name, headers, rows) {
    // xlsxio 通用形式：rows 为纯 aoa（首行表头）
    return { name, rows: [headers, ...rows.map((r) => r.map((v) => (v == null ? "" : String(v))))] };
  }
  $S("survey-export").addEventListener("click", async (e) => {
    if (!window.SamplingApp) { alert("导出组件未就绪，请稍后重试。"); return; }
    e.stopPropagation();
    if (sExportMenu) { sCloseExportMenu(); return; }
    const menu = document.createElement("div");
    menu.className = "survey-export-menu";
    menu.innerHTML =
      `<div class="survey-export-item" data-k="main">测点布局调查（主表格）</div>` +
      `<div class="survey-export-item" data-k="survey">调查表（6 个子表合一份）</div>` +
      `<div class="survey-export-item" data-k="cur">仅当前子表（${escHtml(SURVEY_SHEETS.find((s) => s.key === surveyCur).name)}）</div>`;
    const r = e.currentTarget.getBoundingClientRect();
    menu.style.left = r.left + "px";
    menu.style.top = r.bottom + 2 + "px";
    document.body.appendChild(menu);
    sExportMenu = menu;
  });
  // 加载 xlsxio（经典脚本，挂 window.SamplingXlsx；与主应用 ensureXlsx 同机制，动态 import 不适用）
  let sXlsxPromise = null;
  function sEnsureXlsx() {
    if (window.SamplingXlsx) return Promise.resolve(window.SamplingXlsx);
    if (!sXlsxPromise) {
      sXlsxPromise = (async () => {
        const load = (src) => new Promise((res, rej) => {
          const s = document.createElement("script");
          s.src = src; s.onload = res; s.onerror = () => rej(new Error("加载失败: " + src));
          document.head.appendChild(s);
        });
        if (!window.JSZip) await load("js/jszip.min.js");
        await load("js/xlsxio.js");
        if (!window.SamplingXlsx) throw new Error("xlsx 模块加载失败");
        return window.SamplingXlsx;
      })();
    }
    return sXlsxPromise;
  }

  // 裁剪掉尾部全空行（保存数据时减小体积）
  function sTrimRows(rows) {
    const arr = (rows || []).map((r) => r.map((v) => (v == null ? "" : String(v))));
    while (arr.length && arr[arr.length - 1].every((v) => v === "")) arr.pop();
    return arr;
  }
  function sAllEmpty(data) {
    const sh0 = SURVEY_SHEETS[0];
    return !data || SURVEY_SHEETS.every((sh) => !(data[sh.key] || []).some((r) => (r || []).some((v) => v !== "")));
  }

  // 对外接口：页签切换时刷新渲染；供主界面导出菜单生成「调查表 6 子表合一份」xlsx；供保存/调用数据读写 6 表内容
  window.SurveySheets = {
    refresh: () => { buildSurveyTabs(); renderSurvey(); },
    exportAllWorkbook: async () => {
      const X = await sEnsureXlsx();
      const sheets = SURVEY_SHEETS.map((sh) => sExportAoaSheet(sh.name, sh.headers, surveyData[sh.key] || []));
      return X.writeWorkbook(sheets);
    },
    // 保存数据用：6 表内容（已裁剪空行）；全部为空时返回 null
    getData: () => {
      const rows = surveyRows(); // 确保内存数据已初始化
      void rows;
      const out = {};
      SURVEY_SHEETS.forEach((sh) => { out[sh.key] = sTrimRows(surveyData[sh.key]); });
      return sAllEmpty(out) ? null : out;
    },
    // 调用数据用：写入 6 表内容（缺省子表补默认 10 行空表）并重渲染
    setData: (data) => {
      surveyData = data && typeof data === "object" ? data : {};
      SURVEY_SHEETS.forEach((sh) => {
        if (!Array.isArray(surveyData[sh.key])) surveyData[sh.key] = Array.from({ length: 10 }, () => sh.headers.map(() => ""));
      });
      surveySave();
      renderSurvey();
    },
  };

  // 加载完成立即渲染一次（面板隐藏不影响 DOM 渲染），确保进入页签即见全部子表按钮
  buildSurveyTabs();
  renderSurvey();
