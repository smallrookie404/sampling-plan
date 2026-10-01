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

  // 固定下拉列选项（按表头名匹配；「单元/工作场所」列动态取主表车间名，不在此列）
  const SURVEY_COL_OPTIONS = {
    "*设施布局": ["机群式布局", "U型生产线布局", "直线型生产线布局", "Y型生产线布局", "其他"],
    "*物理状态": ["气态", "液态", "粉末", "颗粒状", "片状", "条状", "块状", "固态"],
    "*储存方式": ["袋装", "桶装", "罐装", "散装"],
    "*加药/投料方式": ["人工", "自动"],
    "*运输方式": ["槽罐车", "叉车", "货车", "皮带"],
    "*装卸方式": ["接驳", "人工装卸"],
    "*类型": ["产品", "中间产品", "副产品", "联产品"],
    "*包装方式": ["袋装", "桶装", "罐装", "散装"],
    "*防护类型": ["防毒", "防尘", "防噪", "减振", "防暑降温", "防低温", "防非电离辐射", "防电离辐射"],
    "*防护用品分类": ["眼面防护", "听力防护", "呼吸防护", "防护服装", "手部防护"],
  };
  // 个体防护表：「防护用品类别」下拉选项按同行「防护用品分类」联动
  const PPE_CATEGORY_MAP = {
    "听力防护": ["防噪耳塞", "防噪耳罩"],
    "呼吸防护": ["防尘口罩", "防毒面具", "防尘毒半面具", "一次性口罩"],
    "眼面防护": ["电焊面罩", "护目镜"],
  };
  // 个体防护表：按「防护用品类别」联动的列选项（无匹配时该列退回普通文本编辑）
  const PPE_FIELD_OPTIONS = {
    "*型号或规格": {
      "防噪耳塞": ["1110", "1100", "1270", "1250"],
      "防尘口罩": ["KN90", "KN95", "KP90", "KP95", "KP100"],
      "防毒面具": ["配A1滤毒盒", "配E1滤毒盒", "配A1E1滤毒盒", "配3#滤毒盒", "配7#滤毒盒"],
    },
    "*防护性能参数": {
      "防噪耳塞": ["NRR=29dB，标称降噪值11dB", "NRR=24dB，标称降噪值8.5dB", "NRR=33dB，标称降噪值13dB"],
      "防毒面具": [
        "APF=10，防护有机气体或蒸气",
        "APF=10，防护酸性气体或蒸气",
        "APF=10，防护有机及酸性气体或蒸气",
      ],
      "防尘口罩": [
        "KN90，APF=10，对非油性颗粒物的过滤效率≥90%",
        "KN95，APF=10，对非油性颗粒物的过滤效率≥95%",
        "KP90，APF=10，对油性及非油性颗粒物的过滤效率≥90%",
        "KP95，APF=10，对油性及非油性颗粒物的过滤效率≥95%",
        "KP100，APF=10，对油性及非油性颗粒物的过滤效率≥99.97%",
      ],
    },
  };
  let surveyData = null; // { key: [[cell,...],...] }
  let surveyCur = "process"; // 当前显示的子表
  // 选区状态（与主表格一致：cur 当前单元格 + 拖选矩形）
  let sCur = null; // { r, c }
  let sSelStart = null;
  let sSelEnd = null;
  let sSelAnchor = null; // 选区锚点（Shift+方向扩展用，与主表格 selAnchor 同义）
  let sDragging = false;
  let sDragCell = null; // 拖动期间鼠标所在的单元格（判断是否真的跨格）
  let sDragMovedCell = false; // 本次按下后是否拖到过其他单元格（区别于原地手抖）

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

  // 本地缓存落盘：打字路径每个按键都会调用，改为 300ms 防抖合并写入，
  // 避免 JSON.stringify 全部 6 表 + localStorage 磁盘 IO 阻塞输入；
  // 离开页面/切后台时强制落盘，保证数据不丢（内存数据始终实时，落盘仅影响崩溃恢复）
  let surveySaveTimer = null;
  function surveySave() {
    if (surveySaveTimer) clearTimeout(surveySaveTimer);
    surveySaveTimer = setTimeout(surveySaveFlush, 300);
  }
  function surveySaveFlush() {
    if (surveySaveTimer) { clearTimeout(surveySaveTimer); surveySaveTimer = null; }
    try { localStorage.setItem(SURVEY_KEY, JSON.stringify(surveyData)); } catch {}
  }
  document.addEventListener("beforeunload", surveySaveFlush);
  document.addEventListener("visibilitychange", () => {
    if (document.visibilityState === "hidden") surveySaveFlush();
  });

  // ---------- Ctrl+Z 撤销（快照式：批量变更前存全量 surveyData，逐级回退；与主表格一致） ----------
  const sUndoStack = [];
  const S_UNDO_MAX = 50;
  function sPushUndo() {
    sUndoStack.push({ type: "all", data: JSON.stringify(surveyData) });
    if (sUndoStack.length > S_UNDO_MAX) sUndoStack.shift();
  }
  // 行级快照：单表单行（首次打字进入编辑用，序列化单行比全量 6 表快百倍）
  function sPushUndoRow(key, r) {
    const rows = surveyData[key];
    if (!rows || !rows[r]) return;
    sUndoStack.push({ type: "row", key, r, data: JSON.stringify(rows[r]) });
    if (sUndoStack.length > S_UNDO_MAX) sUndoStack.shift();
  }
  function sUndoLast() {
    if (!sUndoStack.length) return;
    const snap = sUndoStack.pop();
    if (snap.type === "row" && surveyData[snap.key] && surveyData[snap.key][snap.r]) {
      surveyData[snap.key][snap.r] = JSON.parse(snap.data);
    } else {
      surveyData = JSON.parse(snap.data);
    }
    surveySaveFlush();
    sEditing = false;
    sEditOriginal = null;
    sRemoveArm();
    sCur = null;
    sSelAnchor = sSelStart = sSelEnd = null;
    renderSurvey();
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
    sCellCommit(); // 重建 tbody 前先提交未完成的编辑，防止 sEditing 残留导致首次点击无法进入编辑
    sCloseWsPanel(); // tbody 重建会销毁编辑框，下拉面板一并清理
    sEditing = false; // tbody 重建销毁编辑框，编辑态同步复位（布防框在末尾重建）
    sEditOriginal = null;
    const sh = SURVEY_SHEETS.find((s) => s.key === surveyCur);
    const rows = surveyRows();
    // 表头（模板原格式：必填列带 *，样式上以浅红底提示）；首列（设备名称/物料名称等）横向锁定
    $S("survey-head").innerHTML =
      `<tr>${sh.headers.map((h, i) =>
        `<th class="${h.startsWith("*") ? "req" : ""}${i === 0 ? " s-sticky-name" : ""}" data-i="${i}">${escHtml(h)}</th>`
      ).join("")}</tr>`;
    // colgroup：首列稍窄，其余均分；列数少的表（原辅物料/主要产品/职业防护/个体防护）所有列平均分铺满整表宽度
    if (["material", "product", "protect", "ppe"].includes(sh.key)) {
      const rest = sh.headers.length;
      $S("survey-cols").innerHTML = sh.headers.map(() => `<col style="width:calc(100% / ${rest})">`).join("");
    } else {
      $S("survey-cols").innerHTML = sh.headers.map((h, i) => `<col style="width:${i === 0 ? 150 : 140}px">`).join("");
    }
    // 全量渲染（调查表行数有限，无需虚拟化）；不显示序号列
    $S("survey-body").innerHTML = rows.length
      ? rows.map((r, ri) =>
          `<tr data-r="${ri}">` +
          sh.headers.map((h, ci) =>
            `<td data-c="${ci}"${ci === 0 ? ' class="s-sticky-name"' : ""}><div class="ctext">${escHtml(r[ci])}</div></td>`
          ).join("") +
          `</tr>`
        ).join("")
      : `<tr class="empty-row"><td colspan="${sh.headers.length}" style="text-align:center;color:#94a3b8;padding:16px">暂无数据，点击「+ 新增行」开始填写</td></tr>`;
    $S("survey-status").textContent = `${sh.name} · 共 ${rows.length} 行`;
    updateSurveySelection();
    // 重建 tbody 后为当前格重新布防隐形编辑器（焦点常在，打字/输入法随时可写）
    if (sCur) sArmCell(sCur.r, sCur.c);
  }

  // 切换子表
  $S("survey-tabs").addEventListener("click", (e) => {
    const btn = e.target.closest(".stab");
    if (!btn) return;
    sCellCommit(); // 先提交当前编辑（此时 sCur 尚未清空，编辑内容不丢失）
    sRemoveArm();
    surveyCur = btn.dataset.k;
    sCur = sSelAnchor = sSelStart = sSelEnd = null;
    buildSurveyTabs();
    renderSurvey();
  });

  // ---------- Excel 式编辑（与主表格一致：单击选中即布防编辑框 / 双击定位光标 / 键盘导航 / 复制粘贴） ----------
  let sEditing = false;
  let sEditTyping = false; // 编辑进入方式：true=直接键入（输入模式：方向键提交并移动），false=双击/F2/退格（编辑模式：方向键格内移光标）
  let sEditOriginal = null;

  // 车间下拉面板（body 级 fixed 定位，不被表格裁剪）
  let sWsPanel = null;
  function sCloseWsPanel() {
    if (sWsPanel) { sWsPanel.remove(); sWsPanel = null; }
  }

  function sFindTd(r, c) {
    return $S("survey-body").querySelector(`tr[data-r="${r}"] > td[data-c="${c}"]`);
  }

  function sCellValue(r, c) {
    return (surveyRows()[r] || [])[c] ?? "";
  }

  // ---------- 隐形布防编辑器（Excel 式输入法支持） ----------
  // 单击选中单元格时，把一个透明 textarea 固定定位到该格上并聚焦：输入法上下文
  // 始终挂在可编辑元素上，任何时刻打字/组字都从首字母进入输入法。首次输入
  // （beforeinput/input）或 IME 组合开始（compositionstart）时显形为真正的编辑框。
  let sArm = null; // { r, c, ta, composing, typed }
  function sRemoveArm() {
    if (sArm) { sArm.ta.remove(); sArm = null; }
  }

  function sArmCell(r, c) {
    sRemoveArm();
    if (!sCur || sCur.r !== r || sCur.c !== c) return;
    const td = sFindTd(r, c);
    if (!td) return;
    const rect = td.getBoundingClientRect();
    const wrapRect = $S("survey-wrap").getBoundingClientRect();
    const ta = document.createElement("textarea");
    ta.className = "survey-arm";
    ta.value = "";
    ta.style.left = rect.left + "px";
    ta.style.top = rect.top + "px";
    ta.style.width = rect.width + "px";
    ta.style.height = rect.height + "px";
    document.body.appendChild(ta);
    ta.focus();
    // IME 组合期间绝不移除布防框（移除会打断组合、吞掉首字母）：
    // compositionstart 仅标记，compositionend（上屏/取消）后才显形并带入全部组合文本
    ta.addEventListener("compositionstart", () => {
      if (!sArm) return;
      sArm.composing = true;
    });
    ta.addEventListener("compositionend", (e) => {
      if (!sArm) return;
      sArm.composing = false;
      sArmReveal(); // 组合结束显形：ta.value 已含组合文本（或空=取消），交给编辑框
      void e;
    });
    ta.addEventListener("input", () => {
      if (!sArm) return;
      // 非组合的真实输入（英文直录/数字等）：显形进入编辑
      if (!sArm.composing) sArmReveal();
    });
    sArm = { r, c, ta, composing: false, typed: false };
  }

  // 布防框显形：就地转为正式编辑框（进入编辑态，初值 = 布防框已输入文本 + 单元格原值）
  function sArmReveal() {
    if (!sArm) return;
    const { r, c, ta } = sArm;
    const pending = ta.value; // 组合文本/首键已写入布防框，必须带入编辑框（首字母不丢）
    sRemoveArm();
    if (!sCur || sCur.r !== r || sCur.c !== c) return;
    const td = sFindTd(r, c);
    if (!td) return;
    sEditing = true;
    sEditTyping = true; // 直接键入/输入法显形：输入模式，方向键提交并移动
    sEditOriginal = sCellValue(r, c);
    sBuildEditor(td, pending);
  }

  // 创建编辑控件并同步聚焦（armed）：编辑框与显示态度量一致，选中/编辑共用。
  // 编辑框存在且聚焦后，字符与中文输入法组合全程走浏览器原生路径，
  // 不再有「首个按键落在 body 上被吞、第二个字母才进输入法」的问题。
  function sBuildEditor(td, typed) {
    const rows = surveyRows();
    if (!td || !rows[sCur.r]) return;
    // Excel 替换语义：布防态直接打字覆盖原内容，编辑框以键入文本开头（typed 未传=双击/F2，编辑原值）
    const initVal = typed !== undefined ? typed : sCellValue(sCur.r, sCur.c);
    // 下拉联想列：「单元/工作场所」动态取主表格已填车间名称；设备设施表「操作岗位(工种)」= 同行车间在主表中的岗位/工种；其余固定选项列取 SURVEY_COL_OPTIONS（可输可选）
    const sh = SURVEY_SHEETS.find((s) => s.key === surveyCur);
    const headerName = sh && sh.headers[sCur.c];
    let wsOpts = null;
    if (headerName && headerName.includes("单元/工作场所") && window.SamplingApp && window.SamplingApp.workshopNames) {
      wsOpts = window.SamplingApp.workshopNames();
    } else if (/(操作|使用|影响|设置|配置)岗位\(工种\)/.test(headerName) && window.SamplingApp && window.SamplingApp.workshopPosts) {
      // 岗位(工种)列（设备设施/原辅物料/主要产品/职业防护/个体防护）：选项 = 同行「单元/工作场所」车间在主表中的岗位/工种；多车间（「、」分隔）时合并去重
      const wsCol = sh.headers.indexOf("单元/工作场所");
      const postsMap = window.SamplingApp.workshopPosts();
      const wsNames = (wsCol >= 0 ? String(surveyRows()[sCur.r]?.[wsCol] ?? "") : "").split("、").map((s) => s.trim()).filter(Boolean);
      const merged = [];
      for (const name of wsNames) {
        for (const p of postsMap[name] || []) {
          if (!merged.includes(p)) merged.push(p);
        }
      }
      wsOpts = merged;
    } else if (headerName === "*防护用品类别") {
      // 个体防护表「防护用品类别」：选项按同行「防护用品分类」联动；分类无对应（未填/防护服装/手部防护）时不显示下拉
      const clsCol = sh.headers.indexOf("*防护用品分类");
      const cls = clsCol >= 0 ? String(surveyRows()[sCur.r]?.[clsCol] ?? "").trim() : "";
      wsOpts = PPE_CATEGORY_MAP[cls] || null;
    } else if (headerName && PPE_FIELD_OPTIONS[headerName]) {
      // 个体防护表「型号或规格」「防护性能参数」：选项按同行「防护用品类别」联动
      const catCol = sh.headers.indexOf("*防护用品类别");
      const cat = catCol >= 0 ? String(surveyRows()[sCur.r]?.[catCol] ?? "").trim() : "";
      wsOpts = PPE_FIELD_OPTIONS[headerName][cat] || null;
    } else if (headerName && SURVEY_COL_OPTIONS[headerName]) {
      wsOpts = SURVEY_COL_OPTIONS[headerName];
    }
    if (wsOpts) {
      sCloseWsPanel();
      // 一律用 textarea（与下方普通编辑路径一致）：多行内容编辑时排版保持不变，不变成一行
      const cs = getComputedStyle(td);
      const h = td.clientHeight - (parseFloat(cs.paddingTop) || 0) - (parseFloat(cs.paddingBottom) || 0);
      td.innerHTML = `<textarea data-r="${sCur.r}" data-c="${sCur.c}">${escHtml(initVal)}</textarea>`;
      const el = td.querySelector("textarea");
      el.style.height = h + "px";
      el.style.overflowY = "hidden";
      el.setSelectionRange(el.value.length, el.value.length);
      el.focus();
      // 布防即实时同步模型：焦点常驻编辑框，数据不能等到提交才取。
      // 首次输入（含 IME 组合）同时进入编辑态，Esc 才能还原到布防时的原值
      const armOriginal = initVal;
      const markEditing = () => { if (!sEditing) { sEditing = true; sEditTyping = true; sEditOriginal = armOriginal; sPushUndoRow(surveyCur, sCur.r); } };
      el.addEventListener("input", () => {
        markEditing();
        if (sCur && Number(el.dataset.r) === sCur.r && Number(el.dataset.c) === sCur.c) {
          surveyRows()[sCur.r][sCur.c] = el.value;
          surveySave();
        }
      });
      // 多选仅限「单元/工作场所」与岗位(工种)列；其余固定选项列单选（点选即替换）
      const multi = !!(
        (headerName && headerName.includes("单元/工作场所")) ||
        (headerName && /(操作|使用|影响|设置|配置)岗位\(工种\)/.test(headerName))
      );
      // 自定义下拉面板：始终显示全部选项（datalist 会按已有内容过滤，单元格有值时只剩匹配项）
      if (wsOpts.length) {
        const panel = document.createElement("div");
        panel.className = "survey-ws-panel";
        panel.innerHTML = wsOpts.map((o) => `<div class="survey-ws-opt${multi && el.value.split("、").includes(o) ? " selected" : ""}" data-v="${escAttr(o)}">${escHtml(o)}</div>`).join("");
        document.body.appendChild(panel);
        const r = el.getBoundingClientRect();
        panel.style.left = r.left + "px";
        panel.style.minWidth = Math.max(r.width, 120) + "px";
        // 先渲染量出面板实际高度，再决定向下弹出（紧贴单元格底边）或向上弹出（紧贴单元格顶边）
        const ph = panel.offsetHeight || 234;
        const below = window.innerHeight - r.bottom;
        if (below >= ph + 4) {
          panel.style.top = r.bottom + 2 + "px";
        } else {
          // 视口底部空间不足：向上弹出，面板底边紧贴单元格顶边（不留缝隙）
          panel.style.top = Math.max(4, r.top - ph - 1) + "px";
        }
        panel.addEventListener("mousedown", (ev) => {
          const opt = ev.target.closest(".survey-ws-opt");
          if (!opt) return;
          ev.preventDefault(); // 阻止 input 失焦
          if (multi) {
            // 多选：点选追加（以「、」分隔），重复点选则移除（支持取消）
            const SEP = "、";
            const curParts = el.value.split(SEP).map((s) => s.trim()).filter(Boolean);
            const v = opt.dataset.v;
            const idx = curParts.indexOf(v);
            if (idx >= 0) curParts.splice(idx, 1);
            else curParts.push(v);
            el.value = curParts.join(SEP);
            opt.classList.toggle("selected", idx < 0);
          } else {
            // 单选：点选即替换并关闭面板
            el.value = opt.dataset.v;
            sCloseWsPanel();
          }
          el.focus();
          el.setSelectionRange(el.value.length, el.value.length);
        });
        sWsPanel = panel;
      }
      return;
    }
    // 一律用 textarea：无论内容是否含换行，编辑态排版与显示态完全一致（单段长文本同样自动折行，不变成一行）
    // 编辑不改变行高：textarea 高度 = td 内容区高度（clientHeight 减上下内边距），替换前量好
    const cs = getComputedStyle(td);
    const h = td.clientHeight - (parseFloat(cs.paddingTop) || 0) - (parseFloat(cs.paddingBottom) || 0);
    td.innerHTML = `<textarea data-r="${sCur.r}" data-c="${sCur.c}">${escHtml(initVal)}</textarea>`;
    const ta = td.querySelector("textarea");
    ta.style.height = h + "px";
    ta.style.overflowY = "hidden";
    ta.setSelectionRange(ta.value.length, ta.value.length);
    // 字体度量差异可能导致内容溢出出现滚动条：以内容实际高度为准（不出现滚动条）
    if (ta.scrollHeight > ta.clientHeight) {
      // border-box 下边框占高，补偿后内容区才能完整容纳
      const bh = ta.offsetHeight - ta.clientHeight;
      ta.style.height = (ta.scrollHeight + bh) + "px";
    }
    ta.focus();
    // 布防即实时同步模型：焦点常驻编辑框，数据不能等到提交才取。
    // 首次输入（含 IME 组合）同时进入编辑态，Esc 才能还原到布防时的原值
    const armOriginal = initVal;
    const markEditing = () => { if (!sEditing) { sEditing = true; sEditTyping = true; sEditOriginal = armOriginal; sPushUndoRow(surveyCur, sCur.r); } };
    ta.addEventListener("input", () => {
      markEditing();
      if (sCur && Number(ta.dataset.r) === sCur.r && Number(ta.dataset.c) === sCur.c) {
        surveyRows()[sCur.r][sCur.c] = ta.value;
        surveySave();
      }
    });
  }

  function sCellCommit() {
    if (!sEditing) return;
    sCloseWsPanel();
    const cur = sCur;
    const td = cur ? sFindTd(cur.r, cur.c) : null;
    const input = td && td.querySelector("input,textarea");
    const rows = surveyRows();
    if (input && cur && rows[cur.r]) {
      rows[cur.r][cur.c] = input.value;
      surveySave();
    }
    // 状态无条件复位：即使 sCur 已被清空（如切子表）或编辑框已被重建销毁，也不残留编辑态
    sEditing = false;
    sEditOriginal = null;
    if (td && cur) td.innerHTML = `<div class="ctext">${escHtml(rows[cur.r] ? (rows[cur.r][cur.c] ?? "") : "")}</div>`;
  }

  function sCellCancel() {
    if (!sEditing || !sCur) return;
    sCloseWsPanel();
    const td = sFindTd(sCur.r, sCur.c);
    // 布防期间 input 事件已实时写入模型，取消需还原为进入编辑时的原值
    const rows = surveyRows();
    if (rows[sCur.r]) rows[sCur.r][sCur.c] = sEditOriginal ?? "";
    surveySave();
    sEditing = false;
    sEditOriginal = null;
    if (td) td.innerHTML = `<div class="ctext">${escHtml(sCellValue(sCur.r, sCur.c))}</div>`;
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
    const movedToNewCell = !prev || prev.r !== r || prev.c !== c;
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
    // 键盘导航落点重新布防隐形编辑器（Excel 式：单击/导航仅选中，随时可打字）
    if (movedToNewCell) {
      sEditing = false;
      sEditOriginal = null;
      sRemoveArm();
      sArmCell(r, c);
    }
  }

  function sMoveTo(r, c) {
    const rows = surveyRows();
    const sh = SURVEY_SHEETS.find((s) => s.key === surveyCur);
    r = Math.max(0, Math.min(r, rows.length - 1));
    c = Math.max(0, Math.min(c, sh.headers.length - 1));
    const movedToNewCell = !sCur || sCur.r !== r || sCur.c !== c;
    sCur = { r, c };
    sSelAnchor = sSelStart = sSelEnd = { r, c };
    updateSurveySelection();
    const td = sFindTd(r, c);
    if (td && td.scrollIntoView) td.scrollIntoView({ block: "nearest", inline: "nearest" });
    if (movedToNewCell) {
      sEditing = false;
      sEditOriginal = null;
      sRemoveArm();
      sArmCell(r, c);
    }
  }

  // 单击仅选中；双击进入编辑（与主表格一致），光标落在点击位置
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
    if (sEditing && sCur && r === sCur.r && c === sCur.c) {
      // 编辑中按下当前格：允许格内拖动选文本（走原生选择）；拖出才提交并进入拖选
      sDragging = true;
      sDidDrag = false;
      sDragCell = { r, c };
      sDragMovedCell = false;
      sLastMouse = { x: e.clientX, y: e.clientY };
      return;
    }
    if (sEditing) {
      sCellCommit(); // 点其他格：先提交再切换
    }
    sRemoveArm();
    sCur = { r, c };
    sSelAnchor = sSelStart = sSelEnd = { r, c };
    sDragging = true;
    sDidDrag = false;
    sDragCell = { r, c };
    sDragMovedCell = false;
    sLastMouse = { x: e.clientX, y: e.clientY };
    // 阻止 mousedown 默认聚焦（点击目标 div 无需聚焦，焦点交给隐形布防框）
    e.preventDefault();
    sArmCell(r, c);
    updateSurveySelection();
  });

  $S("survey-body").addEventListener("dblclick", (e) => {
    // 双击进入编辑（Excel 语义）：显形编辑框并定位光标到点击位置
    e.preventDefault();
    const td = e.target.closest("td[data-c]");
    if (!td) return;
    const tr = td.closest("tr[data-r]");
    if (!tr) return;
    const r = Number(tr.dataset.r), c = Number(td.dataset.c);
    if (!sCur || sCur.r !== r || sCur.c !== c) {
      // 双击目标与当前格不同：先提交旧格，再选中新格
      if (sEditing) sCellCommit();
      sRemoveArm();
      sCur = { r, c };
      sSelAnchor = sSelStart = sSelEnd = { r, c };
      updateSurveySelection();
    }
    if (sEditing && sCur.r === r && sCur.c === c) {
      // 已在编辑：仅定位光标
      const el0 = td.querySelector("input,textarea");
      if (el0 && sLastMouse) {
        const pos = sCaretOffset(el0, sLastMouse.x, sLastMouse.y);
        try { el0.setSelectionRange(pos, pos); } catch {}
      }
      return;
    }
    sRemoveArm();
    sEditing = true;
    sEditTyping = false; // 双击进入：编辑模式，方向键在格内移动光标
    sEditOriginal = sCellValue(r, c);
    sBuildEditor(td);
    const el = td.querySelector("input,textarea");
    if (el && sLastMouse) {
      const pos = sCaretOffset(el, sLastMouse.x, sLastMouse.y);
      try { el.setSelectionRange(pos, pos); } catch {}
    }
  });

  document.addEventListener("mousemove", (e) => {
    if (!sDragging) return;
    // 位移未超过阈值（原地手抖/边缘微抖）：不改变任何状态，松开仍按单击进入编辑
    if (Math.abs(e.clientX - sLastMouse.x) <= 3 && Math.abs(e.clientY - sLastMouse.y) <= 3) return;
    const td = e.target.closest && e.target.closest("td[data-c]");
    if (!td) return;
    const tr = td.closest("tr[data-r]");
    if (!tr) return;
    const r = Number(tr.dataset.r), c = Number(td.dataset.c);
    // 编辑态下鼠标仍在原格内：走浏览器原生文本选择，不提交、不进入拖选（与主表格一致）
    if (sEditing && sCur && r === sCur.r && c === sCur.c) return;
    // 跨到其他单元格：进入拖选（首次跨出编辑格先提交编辑，隐形布防框随后重建到拖选起点）
    if (!sDragCell || sDragCell.r !== r || sDragCell.c !== c) {
      if (!sDragMovedCell && sEditing) {
        sCellCommit();
        sRemoveArm();
        sCur = { r, c };
        sSelAnchor = { r, c };
        sArmCell(r, c);
      }
      sDragCell = { r, c };
      sDragMovedCell = true;
      sDidDrag = true;
    }
    sSelEnd = { r, c };
    updateSurveySelection();
  });

  document.addEventListener("mouseup", () => {
    // 单击/拖选结束：拖选跨格时编辑框在 mousemove 提交后已被目标格布防重建
    sDragging = false;
  });

  // 键盘逻辑与主表格完全一致：IME 守卫 / Ctrl 组合键 / F2 / 编辑态 Enter·Tab·Esc / 导航（Shift 扩展、Tab 回绕、Enter 自动加行）/ Home·End / PageUp·Down / Delete 清空 / copy 事件复制
  // 挂 document：单元格选中后焦点不在表格内也能响应
  document.addEventListener("keydown", (e) => {
    if (!sCur) return;
    const ae = document.activeElement;
    // 焦点在其他输入框/文本域（如保存对话框）时不劫持按键；隐形布防框除外（它代表当前格）
    const inBody = ae && $S("survey-body").contains(ae);
    const onArm = sArm && ae === sArm.ta;
    if (ae && (ae.tagName === "INPUT" || ae.tagName === "TEXTAREA" || ae.tagName === "SELECT") && !inBody && !onArm) return;
    const mod = e.ctrlKey || e.metaKey;
    if (e.isComposing || e.key === "Process") return; // 中文输入法组合中不拦截
    // 编辑中的输入框获得焦点时进入编辑态分支（布防框不算：它处于「选中未编辑」态）
    const editInput = inBody && (ae.tagName === "INPUT" || ae.tagName === "TEXTAREA") ? ae : null;

    if (mod) {
      if (editInput) {
        // 编辑中保留浏览器原生快捷键；Ctrl+Enter 在单元格内换行
        if (e.key === "Enter") return;
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
      if (e.key.toLowerCase() === "z") {
        e.preventDefault();
        sUndoLast(); // Ctrl+Z 撤销最近一次变更
        return;
      }
      return;
    }

    // F2：进入编辑（编辑中则光标移到末尾）
    if (e.key === "F2") {
      e.preventDefault();
      if (editInput) {
        try { editInput.setSelectionRange(editInput.value.length, editInput.value.length); } catch {}
      } else {
        sRemoveArm();
        sEditing = true;
        sEditTyping = false; // F2 进入：编辑模式，方向键在格内移动光标
        sEditOriginal = sCellValue(sCur.r, sCur.c);
        const td = sFindTd(sCur.r, sCur.c);
        if (td) sBuildEditor(td);
      }
      return;
    }

    // 编辑模式：Enter 提交并下移（不在单元格内换行，换行内容经粘贴保留 / Alt+Enter 手动换行），Tab 提交并右移，Escape 还原
    if (editInput) {
      if (e.key === "Enter") {
        if (e.altKey) {
          // Alt+Enter：手动插入换行符（浏览器对 Alt+Enter 无原生插入行为）
          e.preventDefault();
          editInput.setRangeText("\n", editInput.selectionStart, editInput.selectionEnd, "end");
          return;
        }
        e.preventDefault();
        sCellCommit();
        sMove(1, 0, { grow: true });
      } else if (e.key === "Tab") {
        e.preventDefault();
        sCellCommit();
        sMove(0, e.shiftKey ? -1 : 1, { grow: true, wrap: true });
      } else if (e.key === "Escape") {
        e.preventDefault();
        sCellCancel();
        // 取消后回到布防态（隐形框重新聚焦），继续打字/导航不中断
        sRemoveArm();
        sArmCell(sCur.r, sCur.c);
      } else if (sEditTyping && !e.altKey && (e.key === "ArrowDown" || e.key === "ArrowUp" || e.key === "ArrowLeft" || e.key === "ArrowRight")) {
        // 直接键入进入的编辑（WPS 输入模式）：方向键提交当前格并移动选中单元格
        e.preventDefault();
        sCellCommit();
        sMove(e.key === "ArrowDown" ? 1 : e.key === "ArrowUp" ? -1 : 0, e.key === "ArrowRight" ? 1 : e.key === "ArrowLeft" ? -1 : 0, { grow: true });
        return;
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
    // Delete/Backspace：清空选区（或当前格）内容；焦点在布防框内且单格选区时走原生逐字删除
    if (e.key === "Delete" || e.key === "Backspace") {
      const td = sFindTd(sCur.r, sCur.c);
      const el = td && td.querySelector("input,textarea");
      if (el && document.activeElement === el) {
        // 布防框聚焦中：Backspace/Delete 走浏览器原生编辑（不清空整格）
        return;
      }
      e.preventDefault();
      const rect = sSelRect();
      const rows = surveyRows();
      const r1 = rect ? rect.r1 : sCur.r;
      const r2 = rect ? rect.r2 : sCur.r;
      const c1 = rect ? rect.c1 : sCur.c;
      const c2 = rect ? rect.c2 : sCur.c;
      // 清空前存快照（清空前状态），Ctrl+Z 可复原（与主表格一致）
      sPushUndo();
      for (let r = r1; r <= r2; r++) {
        if (!rows[r]) continue;
        for (let c = c1; c <= c2; c++) rows[r][c] = "";
      }
      surveySave();
      renderSurvey();
      return;
    }
    // 可打印字符：焦点应已在隐形布防框（单击/导航时布防），首个键经原生写入并触发显形；
    // 焦点意外丢失时重新布防兜底（不影响本键之后的输入）
    if (e.key.length === 1 && !e.altKey) {
      if (!sArm || document.activeElement !== sArm.ta) {
        if (!editInput) sArmCell(sCur.r, sCur.c);
      }
    }
  });

  // 复制：与主表格一致走 copy 事件（Ctrl+C 原生触发），选区内容以 TSV 写入剪贴板。
  // 仅调查表页签生效：主表的 copy 处理器同样挂 document，需避免互相覆盖剪贴板内容
  // 点击调查表区域之外（含表格行下方的容器空白）退出编辑（与主表格/Excel 一致），清理布防框
  document.addEventListener("pointerdown", (e) => {
    if (!$S("tab-survey").classList.contains("active")) return;
    const t = e.target;
    if (t.closest && (t.closest("#survey-grid td") || t.closest("#survey-grid th") || t.closest(".grid-toolbar") || t.closest(".survey-ws-panel"))) return;
    if (!sEditing && !sArm) return;
    if (sEditing) sCellCommit();
    sRemoveArm();
  });

  document.addEventListener("copy", (e) => {
    if (!$S("tab-survey").classList.contains("active")) return;
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
  // 粘贴基准 = 选区左上角；与主表格一致：有选区时按选区大小将剪贴板内容按行列规律重复填充，无选区（或单格选区）时按剪贴板内容大小直接粘贴
  document.addEventListener("paste", (e) => {
    if (!$S("tab-survey").classList.contains("active")) return;
    if (!sCur) return;
    // 焦点在其他输入框/文本域时不劫持（如保存对话框）；隐形布防框代表当前格，不在此列
    const ae = document.activeElement;
    const onArm = sArm && ae === sArm.ta;
    if (ae && (ae.tagName === "INPUT" || ae.tagName === "TEXTAREA" || ae.tagName === "SELECT") && !$S("survey-body").contains(ae) && !onArm) return;
    const text = e.clipboardData && e.clipboardData.getData("text/plain");
    if (!text) return;
    e.preventDefault();
    const rows = surveyRows();
    const editEl = ae && (ae.tagName === "INPUT" || ae.tagName === "TEXTAREA") && $S("survey-body").contains(ae) ? ae : null;
    // 编辑中：无 Tab 的多行纯文本（如段落描述）视为单元格内容，在光标处插入换行，不拆成上下两格
    if (editEl && !text.includes("\t")) {
      const merged = text.replace(/\r\n/g, "\n").replace(/\r/g, "\n").replace(/\n+$/, "");
      const st = editEl.selectionStart ?? editEl.value.length;
      const en = editEl.selectionEnd ?? st;
      editEl.setRangeText(merged, st, en, "end");
      // 布防框粘贴：手动同步模型（setRangeText 不触发 input 事件）
      if (sCur && Number(editEl.dataset.r) === sCur.r && Number(editEl.dataset.c) === sCur.c) {
        surveyRows()[sCur.r][sCur.c] = editEl.value;
        surveySave();
      }
      return;
    }
    // 非编辑态（或含 Tab）：按 TSV 网格解析，无 Tab 多行文本按行拆分为多行（原行为）
    const grid = parseTsvGrid(text);
    // Excel 复制区域末尾常带一个空行，去掉（中间空行保留）
    if (grid.length > 1 && grid[grid.length - 1].length === 1 && grid[grid.length - 1][0] === "") grid.pop();
    if (!grid.length) return;
    const srcH = grid.length;
    const srcW = grid.reduce((m, r) => Math.max(m, r.length), 1);
    let rect = sSelRect();
    if (!rect || (rect.r1 === rect.r2 && rect.c1 === rect.c2)) {
      // 无选区或仅单格：从当前单元格起按剪贴板全部内容大小粘贴
      rect = { r1: sCur.r, c1: sCur.c, r2: sCur.r + srcH - 1, c2: sCur.c + srcW - 1 };
    }
    const need = rect.r2 + 1;
    while (rows.length < need) rows.push(surveyBlankRow());
    sPushUndo(); // 粘贴覆盖前存快照，可 Ctrl+Z 撤销
    for (let r = rect.r1; r <= rect.r2; r++) {
      for (let c = rect.c1; c <= rect.c2; c++) {
        if (c >= rows[r].length) break;
        // 按剪贴板内容行列规律重复填充（单值则整片填入）
        rows[r][c] = (grid[(r - rect.r1) % srcH] || [])[(c - rect.c1) % srcW] ?? "";
      }
    }
    surveySave();
    sEditing = false; // renderSurvey 重建 tbody 会销毁编辑框，编辑态标志必须同步复位
    sEditOriginal = null;
    sRemoveArm();
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
    sPushUndo();
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
    sPushUndo();
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
    sPushUndo();
    rows.splice(at + 1, 0, ...Array.from({ length: n }, () => src.slice()));
    if (sCur && sCur.r > at) sCur = { ...sCur, r: sCur.r + n };
    surveySave();
    renderSurvey();
  });

  // 清空录入区：6 个子表全部重置为默认 10 行空表
  $S("survey-clear").addEventListener("click", () => {
    if (!confirm("确定清空调查表全部 6 个表格的录入内容？此操作不可恢复。")) return;
    sPushUndo(); // 清空前存快照，可 Ctrl+Z 撤销
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
    sPushUndo();
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
      // 首表：劳动定员和职业病危害因素接触情况调查（首行 = 主表格导出表头，其余内容为空）
      const heads = (window.SamplingApp && window.SamplingApp.computedHeaders ? window.SamplingApp.computedHeaders() : []).map((v) => (v == null ? "" : String(v)));
      const ldSheet = { name: "劳动定员和职业病危害因素接触情况调查", rows: [heads] };
      const sheets = [ldSheet, ...SURVEY_SHEETS.map((sh) => sExportAoaSheet(sh.name, sh.headers, surveyData[sh.key] || []))];
      return X.writeWorkbook(sheets);
    },
    // 导入：从已解析的工作表数组（{ 表名: aoa，首行表头 }）识别 6 个子表并写入；返回匹配到的子表数
    importFromArray: (byName) => {
      surveyRows(); // 确保 surveyData 已初始化
      let matched = 0;
      for (const sh of SURVEY_SHEETS) {
        const arr = byName[sh.name];
        if (!Array.isArray(arr) || !arr.length) continue;
        // 按表头名映射列（兼容列序变化），缺失列留空
        const idx = {};
        (arr[0] || []).forEach((h, i) => { const k = String(h ?? "").trim(); if (k && idx[k] === undefined) idx[k] = i; });
        const rows = [];
        for (const r of arr.slice(1)) {
          if (!r || !r.some((v) => v !== "" && v !== null && v !== undefined)) continue;
          rows.push(sh.headers.map((h) => (idx[h] !== undefined ? String(r[idx[h]] ?? "") : "")));
        }
        while (rows.length < 10) rows.push(sh.headers.map(() => ""));
        surveyData[sh.key] = rows;
        matched++;
      }
      if (matched) {
        surveySave();
        sCloseWsPanel();
        sEditing = false;
        sEditOriginal = null;
        sCur = sSelAnchor = sSelStart = sSelEnd = null;
        buildSurveyTabs();
        renderSurvey();
      }
      return matched;
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
