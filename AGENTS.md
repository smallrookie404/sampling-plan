# 采样计划软件（sampling-plan）维护指南

## 项目简介

职业卫生检测「系统测点布局调查 · 采样计划软件」：把原 Excel 公式表升级为可独立运行的网页工具。核心功能（计算、编辑、校验、导出）全部在浏览器端完成；并整合了「现场调查 Excel 上传」工具（原 zhiyeweishen 项目），平台登录界面同时作为整个软件的总登录门禁。

技术形态：**无构建、无框架、无打包**的原生 JS + IIFE 模块，后端只负责「存数据」与「代理平台接口」，且有本地服务 / GitHub Contents API / Cloudflare KV 三种可选实现。

## 目录结构

```
sampling-plan/
├── README.md                      # 使用与部署文档
├── AGENTS.md                       # 本文件
├── wrangler.jsonc                  # Cloudflare Pages 部署配置（KV 绑定）
├── 部署到Cloudflare.bat            # Cloudflare 一键部署脚本
├── 同步数据库到GitHub.bat          # 提交 records.json / library.json 并推送
├── .github/workflows/
│   ├── deploy-pages.yml            # push main → GitHub Pages
│   └── deploy-cloudflare.yml       # push main → Cloudflare Pages（含 Functions）
└── sampling-plan-app/
    ├── index.html                  # 唯一页面：主工具 + 调查表页签 + 登录遮罩 + 上传视图
    ├── server.mjs                  # 本地服务（Node 标准库，端口 8017，/api/records、/api/library）
    ├── css/styles.css              # 主界面样式（CSS 变量 --primary 等，.xcdc 作用域给登录/上传视图）
    ├── .assetsignore               # Cloudflare 部署时排除 tests/ scripts/ data/ server.mjs *.bat README.md
    ├── data/
    │   ├── records.json            # 记录「索引」[{id,name,createdAt,updatedAt}]（旧版整表格式会自动迁移）
    │   ├── records/<id>.json       # 单条完整记录（含 rows / survey），本地服务按需创建
    │   └── library.json            # 危害因素库 + 检测项目（应用自动写回，勿手工改格式）
    ├── js/
    │   ├── data.js                 # 内置危害因素库/检测项目（由 scripts/extract-data.mjs 生成，勿手改）
    │   ├── logic.js                # 计算引擎（浏览器/Node 通用，纯函数）
    │   ├── xlsxio.js + jszip.min.js# xlsx 读写（懒加载，首次导入/导出时加载）
    │   ├── app.js                  # 主工具界面逻辑；暴露 window.SamplingApp（导出/错误数/车间岗位映射）
    │   ├── survey.js               # 调查表模块（6 个现场调查子表，Excel 式编辑）；暴露 window.SurveySheets
    │   └── upload.js               # 统一登录 + 现场调查上传（平台接口、团队配置、云端同步）
    ├── functions/api/              # Cloudflare Pages Functions（编译为 Worker）
    │   ├── records.js / library.js # KV 数据读写（binding: SAMPLING_RECORDS）
    │   ├── team.js                 # 团队配置云端存取（按账号分 key team:<账号>）
    │   └── platform/[[path]].js    # 平台接口同源代理（cloudflare:sockets 原始 TCP 转发）
    ├── scripts/extract-data.mjs    # 从原 Excel 生成 js/data.js
    └── tests/                      # 自动化测试（Node + Playwright）；另有 check_shots.py、validate_export.py
```

## 运行时组成与模块契约

- **加载顺序**（`index.html` 末尾，不可随意调整）：`data.js → logic.js → survey.js → app.js → upload.js`；`xlsxio.js` 与 `jszip.min.js` 由 app.js/survey.js 的 `ensureXlsx()` 在首次导入/导出时懒加载。app.js/survey.js 有大量顶层 `$()` 元素绑定与 `init()`，因此脚本必须放在 `</body>` 前。
- **模块间只通过全局对象通信**：

| 全局对象 | 定义处 | 主要成员 |
|---|---|---|
| `SamplingData` | js/data.js | 内置 `{ hazardFactors, detectionItems }` |
| `SamplingLogic` | js/logic.js | `computeRows` `validateRow` `countErrors` `snapshotRows` `restoreRows` `findDuplicates` |
| `SamplingXlsx` | js/xlsxio.js | `readWorkbook` `sheetToArray` `writeWorkbook`（支持 `[{name,rows,widths}]` 多表）`downloadBlob` `colToIndex` `indexToCol` |
| `SamplingApp` | app.js `window.SamplingApp` | `exportWorkbookBytes` `exportName` `countErrors` `promptSave` `askRowCount` `workshopNames` `computedHeaders` `workshopPosts` |
| `SurveySheets` | survey.js `window.SurveySheets` | `refresh` `exportAllWorkbook` `importFromArray` `getData` `setData` |
| `SamplingUpload` | upload.js | `show` `bootstrap`（登录态判断 + 上传视图） |

- **反向依赖（改动签名必须同时改两处）**：
  - survey.js 依赖 `SamplingApp`：`workshopNames()`（主表 A 列车间去重保序）、`workshopPosts()`（自动计算区 W→X 映射，含下填）、`computedHeaders()`、`askRowCount()`；
  - upload.js 依赖 `SamplingApp.exportWorkbookBytes/exportName/countErrors/promptSave` 与 `SurveySheets.exportAllWorkbook`；
  - app.js 依赖 `SurveySheets.getData/setData/refresh/importFromArray/exportAllWorkbook` 与 `SamplingUpload.show`。

## 数据模型与三端存储契约

**主表行对象**（logic.js `restoreRows`/`snapshotRows`）：`{ input:{A..V}, manual:{...}, overridden:{...}, values:{W..BI}, errors:{} }`。

**列结构**：共 61 列 = 录入区 22 列（`A–U` + `V` 备注）+ 自动计算区 39 列（`W–BI`）。渲染顺序、表头、列宽在 app.js 的 `ALL_COLS` / `HEADERS` / `COL_W`；计算侧的列常量在 logic.js：`INPUT_COLS` `COMPUTED_COLS` `MANUAL_COLS`（自动区中仍手工填写的列）`OVERRIDE_COLS`（有下拉、可覆盖自动值：Y/Z/AO/AR）`TEXT_OVERRIDE_COLS`（BI）。**任何增删列都要同时改这两处**。

**记录格式**：`{ id, name, createdAt, updatedAt, rows, survey? }`；远端只存**索引**（不含 rows），单条按 id 拉取，localStorage 里保存完整镜像作为离线兜底与预览缓存。

| 存储端 | 索引位置 | 单条记录位置 | 写入入口 |
|---|---|---|---|
| 本地服务（:8017） | `data/records.json` | `data/records/<id>.json` | `server.mjs` `/api/records` |
| GitHub | `cfg.path`（默认 `sampling-plan-app/data/records.json`） | 同目录 `records/<id>.json` | app.js `github*` 系列（Contents API） |
| Cloudflare KV | key `records-index` | key `record:<id>` | `functions/api/records.js` |

- 环境探测：app.js `detectStorageMode()` 用 `/api/health` 判定 `server`，否则看 GitHub 配置 → `github` / `static-unconfigured`；`file://` 打开为 `temp`（仅浏览器暂存）。
- **「取较新一端并互相同步」只实现在参考库（library）**（app.js `loadLibrary`，按 `updatedAt` 比较后回写较旧端）。**记录（records）侧不做时间比较**，读取优先级固定为 GitHub > 本地服务 > localStorage 镜像。
- 兼容迁移：`server.mjs` 与 `functions/api/records.js` 都带 `migrateLegacy()`，会把旧版「整表记录数组」自动拆成「索引 + 单条」（本地会建出 `data/records/`）；仓库中的 `records.json` 仍是旧格式属于正常现象。

**浏览器存储键**：

| 键 | 位置 | 内容 |
|---|---|---|
| `samplingPlanRecords_v1` | localStorage | 完整记录镜像（离线兜底） |
| `samplingPlanLibrary_v1` | localStorage | 参考库镜像 |
| `samplingPlanGithubConfig_v1` | localStorage | 仓库/分支/路径/**Token 明文** |
| `samplingPlanSurvey_v2` | localStorage | 调查表 6 子表数据（改动结构须升版本号） |
| `xcdc_session_v1` | sessionStorage | 平台登录会话 `{token, orgId, userInfo}` |
| `xcdc_team_v1_<账号>` | localStorage | 团队配置兜底（云端 `/api/team` 优先） |
| `xcdc_year_projects_v1:*` | sessionStorage | 按年份的项目列表缓存（30 分钟） |
| `xcdc_upload_remember / username / password` | localStorage | 记住账号密码（**明文，仅本机**） |

## 运行与测试

- 本地启动：双击 `sampling-plan-app\启动采样计划软件.bat`（自动起服务并打开浏览器），或直接双击 `index.html`（临时模式）。
- 自动化测试（需 Node 与 jszip/playwright 依赖）：
  ```bash
  node tests/engine.test.mjs     # 计算引擎 vs 原 Excel 缓存
  node tests/xlsxio.test.mjs     # xlsx 导出/导入往返
  node tests/server.test.mjs     # 本地服务接口
  node tests/ui.smoke.mjs        # 浏览器冒烟（依赖 Chrome/Edge）
  node tests/survey.smoke.mjs    # 调查表冒烟（Playwright）
  node tests/survey.edit.mjs     # 调查表编辑交互（Playwright）
  node tests/cloudflare.test.mjs # Cloudflare Functions 接口
  ```
- 测试注意：应用会自动把参考库写回 `data/library.json`；跑会改动参考库的测试前先备份，测试后 `git checkout -- sampling-plan-app/data/library.json` 恢复。
- 环境注意：本机系统 PATH 中可能没有 `node`（`启动采样计划软件.bat` 会按顺序回退到内置运行时）。可先取内置 Node 绝对路径再执行，例如 `C:\Users\<用户>\.dsh\dsh-runtimes\dsh-primary-runtime\dependencies\node\bin\node.exe`；`sampling-plan-app/node_modules` 已含 jszip 与 playwright-core。
- 快速自检：对改动文件跑 `node --check`（对全部 js/mjs 跑一遍也能一次性发现语法问题）。

## 部署

| 目标 | 触发方式 | 能力 |
|---|---|---|
| GitHub Pages（smallrookie404.github.io/sampling-plan） | push main → deploy-pages.yml | 静态工具可用；**无法调用平台 HTTP 接口**（HTTPS 混合内容被浏览器拦截） |
| Cloudflare Pages（sampling-plan.pages.dev） | push main → deploy-cloudflare.yml，或 部署到Cloudflare.bat | 静态工具 + Functions + KV；登录/验证码/上传**必须用此版本** |

Cloudflare 工作流需要 GitHub 仓库 Secrets：`CLOUDFLARE_API_TOKEN`、`CLOUDFLARE_ACCOUNT_ID`。KV 命名空间绑定 `SAMPLING_RECORDS`（wrangler.jsonc）。

## 关键机制与维护要点

1. **数据同步**：记录与参考库按上表三端存储契约写入。**只有参考库**做「取较新一端并互相同步」；记录侧读取优先级固定为 GitHub > 本地服务 > localStorage 镜像。
2. **统一登录**（upload.js）：平台登录（账号/密码 RSA 加密/验证码）成功后调 `auth/info` 获取权威用户信息并解析机构（顶层 `organizationId` 优先，其次 `organization.organizationId`），机构仅用于接口数据范围、**不在界面显示**。会话存 sessionStorage；登录遮罩默认显示、有会话时内联脚本立即隐藏，避免主界面闪现。
3. **团队配置云端保存**：成员/调查人/复核人/日期确认后，本地 localStorage 兜底 + Cloudflare KV（`/api/team`，按账号 `team:<userCode>`）同步，换电脑登录同一账号可读取；确认后切换项目不清空，手动修改才更新。
4. **平台代理**：`functions/api/platform/[[path]].js` 用 `cloudflare:sockets` 原始 TCP 连接平台 `http://223.93.144.122:27800`，透传请求头/响应头/Cookie 并解压 gzip，绕过 Cloudflare 直连 IP 限制；`upload.js` 在 pages.dev 上自动走 `/api/platform`，本地 http/file 直连平台。代理只透传固定头白名单（authorization / organizationId / belongProject / content-type / user-agent / accept / cookie），新增自定义头必须同步白名单。
5. **导出上传联动**：「数据上传」按钮生成当前表格的导出 Excel（`window.SamplingApp.exportWorkbookBytes()`）作为上传文件，无需手动选文件；导出范围到「检测项目」列最后一个非空单元格。
6. **界面约定**：主界面样式变量在 `css/styles.css`（`--primary: #1f4e79` 等）；登录/上传视图样式统一用 `.xcdc` 作用域并复用同一套变量。新增页面元素需与 `upload.js` 中的 id 引用保持一致（上传页会一次性校验全部关键 id，缺任一个即 alert 并中止）。
7. **调查表模块**（survey.js）：6 个现场调查子表（生产工艺/设备设施/原辅物料/主要产品/职业防护/个体防护），页签位于「测点布局调查」旁。Excel 式交互与主表格对齐：单击选中、双击编辑、拖选矩形、Enter 提交并下移、Alt+Enter 或编辑态粘贴多行为单元格内换行（非编辑态粘贴按行拆分填入多行）；下拉多选仅「车间」与「岗位(工种)」列，其余固定选项列为单选，「单元/工作场所」动态取主表车间名，岗位下拉数据源为主表自动计算区 W→X（含下填）。数据存 localStorage 键 `samplingPlanSurvey_v2`（统一默认 10 行），随记录保存/还原；无 Tab 多行粘贴兼容。导入 Excel 支持调查表；导出 6 合一首表为「劳动定员和职业病危害因素接触情况调查」（首行＝主表自动计算区表头，无数据行），上传页进入时自动生成两个文件。主表无数据时仍可保存仅含调查表的记录（两者均空才拦截）。
8. **主表格编辑模型**：单击仅选中并「布防」隐形 textarea，输入/输入法经 `armReveal` 显形，双击进入编辑；Enter/Tab 提交并移动、Esc 还原、Delete 清空选区。粘贴统一走 `parseTsvGrid`：编辑态且无 Tab 时视为单元格内换行，否则按选区左上角平铺。`W–BI` 自动列可被覆盖的机制由 `overridden` 标记保留手动值，导入「仅自动计算区」文件时靠它反推录入区。

## 已知注意事项

- **Token/凭据安全**：GitHub fine-grained PAT 与平台密码仅存本机浏览器（localStorage/sessionStorage），**严禁写入仓库或内置到代码**；公开仓库勿存放企业敏感调查数据。
- 平台接口为真实生产系统，上传/成员同步等写操作会影响真实数据，改动接口契约前先核对（请求体数组/对象格式易踩坑：成员新增传对象、成员删除传裸数字数组、上传用 FormData 且 query/body/请求头三处联动）。
- `data/library.json`、`data/records.json` 由应用自动写回，仓库中保持提交以利备份；`.gitignore` 已排除原版 Excel 模板（含调查数据）。
- `data/records.json` 目前仍是旧版「整表记录」格式（元素含 `rows`），首次访问 `/api/records` 时会自动迁移为索引 + `data/records/<id>.json`；迁移是幂等的，不要手工去改。
- 打开页面、修改参考库都会触发 `scheduleLibrarySync` 防抖写回 `data/library.json`（可能同时写 GitHub / KV），调试或跑测试前后注意备份与恢复。
- 测试契约：`tests/ui.smoke.mjs` 依赖 `td[data-r]/td[data-c]`、`td.rowno`、`div.computed-text`、`#grid-status` 文案与 GitHub 配置键名；`tests/xlsxio.test.mjs` 依赖「只导出 W~BI 自动计算区」的行为。改 DOM 结构或导出范围前先看这些断言。
- 纯函数引擎（logic.js）是唯一可被 Node 直接测试的部分，新增计算请优先放这里，不要在 app.js 里再写一份公式。
- 源文件统一 UTF-8；页面迭代频繁，部署后浏览器需 Ctrl+F5 强刷。

## 改动定位速查（按需求）

> 定位以函数名/变量名为主：行号、行数会随改动漂移，不要依赖行号，也不要在文档里累积行号快照。

| 需求 | 主要改动点 |
|---|---|
| 新增/修改一列自动计算 | logic.js 列常量 + `computeRows` 组装 → app.js `ALL_COLS`/`HEADERS`/`COL_W` → 视情况加 `OVERRIDE_COLS`/`MANUAL_COLS` |
| 改公式/校验规则 | logic.js `computeRows`（组级计算、下填、库查找）、`validateRow`、`countErrors` |
| 改主表格交互（选中/编辑/移动/粘贴/清空） | app.js `armCell`/`armReveal`/`commitCurrent`/`moveCur`/`clearRange`、`parseTsvGrid`+粘贴/复制、键盘路由与鼠标处理 |
| 改虚拟滚动/列宽/表头样式 | app.js `buildHead`/`cellHtml`/`rowHtml`/`renderWindow`（spacer 必须恒为首末子节点）+ `computeGridWidths` |
| 加/改调查表字段 | survey.js `SURVEY_SHEETS`（列名即键）+ `SURVEY_COL_OPTIONS`（下拉）；联动列用 `headers.indexOf` 定位，**改列名会使联动与旧 Excel 导入静默失效** |
| 改调查表导出/导入 | survey.js `exportAllWorkbook` / `importFromArray`（按工作表名+表头名匹配，兼容列序变化） |
| 改记录保存/读取/同步 | app.js `detectStorageMode`/`loadRecords`/`loadRecordById`/`persistRecord`/`deleteRecordById` + 三端实现（server.mjs、functions/api/records.js、app.js `github*`） |
| 改参考库同步与时间戳 | app.js `loadLibrary`/`persistLibrary`/`scheduleLibrarySync`/`libTimestamp` |
| 改登录/验证码/会话 | upload.js `API_BASE`/RSA 实现/`resolveOrgId`/`saveSession` + index.html 登录遮罩与内联隐藏脚本 |
| 改上传流程或平台接口 | upload.js `apiRequest`/`runUploadFlow`/`syncProjectInfo`；代理侧 `functions/api/platform/[[path]].js`（目标 IP、头白名单） |
| 改团队配置 | upload.js `loadTeamSettings`/`saveTeamSettings`/`loadProjectTeam` + `functions/api/team.js`（KV key、64KB 上限） |
| 改样式/主题 | 主界面 `css/styles.css`（CSS 变量）；登录/上传视图是 index.html 内联 `.xcdc`（老内核兼容写法勿改成 `inset` 简写等新语法） |

## 常见任务速查

- 改了主表/上传功能 → `node --check sampling-plan-app/js/app.js sampling-plan-app/js/upload.js` + 跑对应测试。
- 改了计算引擎 → `node --check sampling-plan-app/js/logic.js` + `node tests/engine.test.mjs`（对照原 Excel 缓存）。
- 改了调查表 → `node --check sampling-plan-app/js/survey.js` + 跑 `tests/survey.smoke.mjs` / `tests/survey.edit.mjs`。
- 改了导入导出 → `node tests/xlsxio.test.mjs`（会产出 `tests/out.xlsx`，已被 .gitignore 排除）。
- 改了本地服务 → `node tests/server.test.mjs`。
- 改了 Cloudflare Functions → 本地用 `wrangler pages functions build` 验证语法，或直接依赖工作流部署后测 `https://sampling-plan.pages.dev/api/health`，再跑 `node tests/cloudflare.test.mjs`。
- 需要平台白名单 → 把 Cloudflare 出口 IP（https://www.cloudflare.com/ips/）提供给平台管理员。
