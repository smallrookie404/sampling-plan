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
    ├── .assetsignore               # Cloudflare 部署时排除 tests/ scripts/ *.log server.mjs *.bat README.md data/
    ├── data/
    │   ├── records.json            # 记录「索引」[{id,name,createdAt,updatedAt}]（旧版整表格式会自动迁移）
    │   ├── records/<id>.json       # 单条完整记录（含 rows / survey），本地服务按需创建
    │   └── library.json            # 危害因素库（应用自动写回，勿手工改格式）
    ├── js/
    │   ├── util.js                 # ★ 全站共享工具（转义 / 脚本懒加载 / 列表结构兼容）；须最先加载
    │   ├── data.js                 # 全量数据 {mainHeaders,hazardFactors}，仅 Node 测试 require（勿手改）
    │   ├── data-core.js            # 主表列结构（仅 mainHeaders），页面启动同步加载（勿手改）
    │   ├── data-library.js         # 内置危害因素库，页面端懒加载（勿手改）
    │   ├── logic.js                # 计算引擎（浏览器/Node 通用，纯函数）
    │   ├── xlsxio.js + jszip.min.js# xlsx 读写（懒加载，首次导入/导出时加载）
    │   ├── app.js                  # 主工具界面逻辑；暴露 window.SamplingApp（导出/错误数/车间岗位映射）
    │   ├── survey.js               # 调查表模块（6 个现场调查子表，Excel 式编辑）；暴露 window.SurveySheets
    │   ├── blf.js                  # 「网站数据导入」「检评报告生成」等；暴露 window.SamplingBlf
    │   └── upload.js               # 统一登录 + 现场调查上传（平台接口、团队配置、云端同步）
    ├── functions/api/              # Cloudflare Pages Functions（编译为 Worker）
    │   ├── records.js / library.js # KV 数据读写（binding: SAMPLING_RECORDS）
    │   ├── team.js                 # 团队配置云端存取（按账号分 key team:<账号>）
    │   ├── bridge/[[path]].js      # 云端账号桥（OCR + RSA + 登录 + 令牌缓存；`/login` 供登录页云端代登录、`/code` 备用只 OCR）
    │   └── platform/[[path]].js    # 平台接口同源代理（cloudflare:sockets 原始 TCP 转发）
    ├── scripts/extract-data.mjs    # 从原 Excel 生成 js/data.js + data-core.js + data-library.js
    └── tests/                      # 自动化测试（Node + Playwright）；另有 _browser.mjs、check_shots.py、validate_export.py
```

## 运行时组成与模块契约

- **加载顺序**（`index.html` 末尾，不可随意调整）：`util.js → data-core.js → logic.js → survey.js → app.js → upload.js → blf-data.js → blf.js`；`data-library.js`（内置库，~125KB）、`xlsxio.js` 与 `jszip.min.js` 均为**懒加载**——内置库由 app.js `ensureLibraryDefaults()` 在「无云端/本地参考库」时经 `SamplingUtil.loadScript()` 拉取，xlsx 由各模块的 `ensureXlsx()` 在首次导入/导出时拉取。app.js/survey.js 有大量顶层 `$()` 元素绑定与 `init()`，因此脚本必须放在 `</body>` 前。
  - **`util.js` 必须最先加载**：其余业务脚本在顶层就引用 `window.SamplingUtil`（app/survey/upload 为硬依赖；blf.js 在 Node 单测里单独加载，故对 `SamplingUtil` 缺失做了内联兜底）。
- **模块间只通过全局对象通信**：

| 全局对象 | 定义处 | 主要成员 |
|---|---|---|
| `SamplingUtil` | js/util.js | `loadScript` `loadScripts` `escHtml` `escAttr` `escXml` `stripWs` `squashWs` `pickRows` `pickField` |
| `SamplingData` | js/data-core.js（页面启动）/ js/data.js（Node 测试） | data-core.js 仅 `{ mainHeaders }`；data.js 为 `{ mainHeaders, hazardFactors }` |
| `SamplingLibrary` | js/data-library.js | 内置 `{ hazardFactors }`（页面端懒加载） |
| `SamplingLogic` | js/logic.js | `computeRows` `validateRow` `countErrors` `snapshotRows` `restoreRows` |
| `SamplingXlsx` | js/xlsxio.js | `readWorkbook` `sheetToArray` `writeWorkbook`（支持 `[{name,rows,widths}]` 多表）`downloadBlob` `colToIndex` `indexToCol` |
| `SamplingApp` | app.js `window.SamplingApp` | `exportWorkbookBytes` `exportName` `countErrors` `promptSave` `askRowCount` `workshopNames` `computedHeaders` `workshopPosts` `loadScript` `importWorkbook` `syncNonNoiseJobs` |
| `SurveySheets` | survey.js `window.SurveySheets` | `refresh` `exportAllWorkbook` `importFromArray` `getData` `setData` |
| `SamplingBlf` | blf.js | `patchPjbgDocx` `genPjbgDocx` 等（网站数据导入 / 检评报告后处理） |
| `SamplingUpload` | upload.js | `show` `bootstrap`（登录态判断 + 上传视图）；数据报告回退/审核、账号池、报告 PDF 显示 |

- **反向依赖（改动签名必须同时改两处）**：
  - survey.js 依赖 `SamplingApp`：`workshopNames()`（主表 A 列车间去重保序）、`workshopPosts()`（自动计算区 W→X 映射，含下填）、`computedHeaders()`、`askRowCount()`；
  - upload.js 依赖 `SamplingApp.exportWorkbookBytes/exportName/countErrors/promptSave` 与 `SurveySheets.exportAllWorkbook`；
  - app.js 依赖 `SurveySheets.getData/setData/refresh/importFromArray/exportAllWorkbook` 与 `SamplingUpload.show`；
  - blf.js 依赖 `SamplingUpload`（搜索项目 / 401 处理）与 `SamplingApp.importWorkbook`。

## 数据模型与三端存储契约

**主表行对象**（logic.js `restoreRows`/`snapshotRows`）：`{ input:{A..V}, manual:{...}, overridden:{...}, values:{W..BI}, errors:{} }`。

**列结构**：共 61 列 = 录入区 22 列（`A–U` + `V` 备注）+ 自动计算区 39 列（`W–BI`）。渲染顺序、表头、列宽在 app.js 的 `ALL_COLS` / `HEADERS` / `COL_W`；计算侧的列常量在 logic.js：`INPUT_COLS` `COMPUTED_COLS` `MANUAL_COLS`（自动区中仍手工填写的列）`OVERRIDE_COLS`（有下拉、可覆盖自动值：Y/Z/AO/AR）`TEXT_OVERRIDE_COLS`（BI）。**任何增删列都要同时改这两处**。

**录入区→计算区的「组级填充」**（logic.js `computeRows`）：以下录入列有内容时，会按组把值扩散到计算区同组各行；组内取**第一个非空**值（J 还要求该行 AM 未被手工覆盖）；手工覆盖（`overridden`）始终优先。
- `J` 工作内容 → `AM`，分组键 = **车间(W)|岗位(X)|点位(AL)**
- `O` 岗位性质 → `Y`（`Ybase`），分组键 = **车间(W)|岗位(X)**（不含点位）；组内 O 全空 → 默认「固定」
- `K` 危害因素其他来源 → `BE`/`BD`、`L` 作业方式 → `AC`，分组键 = 车间(W)|岗位(X)|点位(AL)

**记录格式**：`{ id, name, createdAt, updatedAt, rows, survey? }`；远端只存**索引**（不含 rows），单条按 id 拉取，localStorage 里保存完整镜像作为离线兜底与预览缓存。

| 存储端 | 索引位置 | 单条记录位置 | 写入入口 |
|---|---|---|---|
| 本地服务（:8017） | `data/records.json` | `data/records/<id>.json` | `server.mjs` `/api/records` |
| GitHub | `cfg.path`（默认 `sampling-plan-app/data/records.json`） | 同目录 `records/<id>.json` | app.js `github*` 系列（Contents API） |
| Cloudflare KV | key `records-index` | key `record:<id>` | `functions/api/records.js` |

- 环境探测：app.js `detectStorageMode()` 用 `/api/health` 判定 `server`，否则看 GitHub 配置 → `github` / `static-unconfigured`；`file://` 打开为 `temp`（仅浏览器暂存）。
- **⚠ 三端各自独立存储，互不同步（2026-10-13 起）**：每类数据**只写一个「主端」**，读取时按 `updatedAt` 取最新，**不再把较新一端回写到较旧一端**。
  - 主端判定（记录与参考库同一套规则）：静态部署 + 已配 GitHub → **GitHub**；否则当前服务端（本地版 = 项目文件夹；Cloudflare Pages = **KV**）；再否则只写浏览器 localStorage。
  - 参考库：`persistLibrary` 只写主端（原来会**同时写 GitHub + KV**）；`loadLibrary` 只读不写（原来会回写较旧端）。**打开页面不再写参考库**，杜绝每次开页写一次 KV。
  - 记录侧读取优先级仍为 GitHub > 本地服务 > localStorage 镜像；写入与读取保持同一端。
  - 目的：Cloudflare KV 免费额度**写只有 1000/天**，杜绝无谓写。剩余 KV 写只来自：Pages 版保存/删除记录、Pages 版改动参考库、`/api/team` 团队配置（低频）。
- 兼容迁移：`server.mjs` 与 `functions/api/records.js` 都带 `migrateLegacy()`，会把旧版「整表记录数组」自动拆成「索引 + 单条」（本地会建出 `data/records/`）；仓库中的 `records.json` 仍是旧格式属于正常现象。

**浏览器存储键**：

| 键 | 位置 | 内容 |
|---|---|---|
| `samplingPlanRecords_v1` | localStorage | 完整记录镜像（离线兜底） |
| `samplingPlanLibrary_v1` | localStorage | 参考库镜像 |
| `samplingPlanLibrary_v1` | localStorage | 参考库镜像（`<id>_<租户标识>` 命名空间） |
| `samplingPlanLibraryDefaults_v1` | localStorage | 内置参考库默认值缓存（懒加载 data-library.js 解析后写入，见 app.js `ensureLibraryDefaults`） |
| `samplingPlanGithubConfig_v1` | localStorage | 仓库/分支/路径/**Token 明文** |
| `samplingPlanSurvey_v2` | localStorage | 调查表 6 子表数据（改动结构须升版本号） |
| `xcdc_session_v1` | sessionStorage | 平台登录会话 `{token, orgId, userInfo}` |
| `xcdc_team_v1_<账号>` | localStorage | 团队配置兜底（云端 `/api/team` 优先） |
| `xcdc_year_projects_v1:*` | sessionStorage | 按年份的项目列表缓存（30 分钟） |
| `xcdc_upload_remember / username / password` | localStorage | 记住账号密码（**明文，仅本机**） |

## 运行与测试

- 本地启动：双击**仓库根目录的 `一键启动.bat`**（推荐，一键）或 `sampling-plan-app\启动采样计划软件.bat`（同一实现），自动起服务并打开浏览器；或直接双击 `index.html`（临时模式）。
  - 实现集中在 `sampling-plan-app\start-app.bat`，两个中文名脚本只是薄封装（**改启动逻辑只改 start-app.bat**）。
  - 幂等：服务已在跑时只重开浏览器，不会起第二个实例（先探 `/api/health` 再决定是否 spawn）。
  - ⚠ 三个 bat 必须保持 **纯 ASCII 内容 + CRLF 换行**（中文只出现在文件名里）；`goto :label` 在 LF-only 文件里可能解析异常。
  - ⚠ 代理免疫：若系统导出 `HTTP_PROXY` 与 `http_proxy` 两个大小写不同的同名键，`Start-Process` 会抛「字典中重复的关键字」，且 `Invoke-WebRequest` 会把 `127.0.0.1` 也送进代理（报「拒绝访问」）。所以启动器先在自身 PowerShell 子进程里清掉 4 个代理变量，健康检查改用 `HttpWebRequest` + `Proxy = $null`。
  - ⚠ 不要用 `timeout /t` 做等待（stdin 被重定向时会失败），等待在 PowerShell 内用 `Start-Sleep`；整段「检查→启动→轮询」只起 **1 个** PowerShell 子进程（每次 spawn 约 4 秒，重复 spawn 会明显拖慢启动）。
- 自动化测试（需 Node 与 jszip/playwright 依赖；浏览器路径由 `tests/_browser.mjs` 自动探测 Chrome/Edge，缺失则跳过）：
  ```bash
  node tests/engine.test.mjs     # 计算引擎 vs 原 Excel 缓存（需原表 xlsx，可用 SAMPLING_SRC_XLSX 指定；缺失自动跳过）
  node tests/xlsxio.test.mjs     # xlsx 导出/导入往返（需 @oai/artifact-tool；缺失自动跳过）
  node tests/logic.ygroup.test.mjs # 岗位性质(Y)按车间+岗位组填充（纯 Node）
  node tests/server.test.mjs     # 本地服务接口
  node tests/ui.smoke.mjs        # 浏览器冒烟（自动探测 Chrome/Edge）
  node tests/survey.smoke.mjs    # 调查表冒烟（Playwright）
  node tests/survey.edit.mjs     # 调查表编辑交互（Playwright）
  node tests/cloudflare.test.mjs # Cloudflare Functions 接口（索引 + 分条 KV）
  node tests/captcha.test.mjs    # 验证码 OCR（JS 移植）准确率 + RSA 自检
  node tests/bridge-cloudflare.test.mjs # 云端账号桥（假 KV + 假平台，真跑 OCR 与登录；含 `/login` 代登录、`/code` 契约）
  node tests/login.ui.mjs        # 登录界面默认不显示验证码 + 「改用验证码登录」兜底可展开（Playwright）
  node tests/report.smoke.mjs    # 「数据报告处理」页签（回退/审核/断点续审/报告显示/面板缓存）
  node tests/entinfo.smoke.mjs   # 在职人数/接害人数：输入框预填、写入分工、上传成功后清空输入框与缓存（Playwright）
  node tests/hazard.edit.smoke.mjs # 危害因素库单元格编辑（Excel 语义：布防/键入覆盖/双击/F2/导航/清空/选区/物化性质列；Playwright）
  node tests/pjbg.note.test.mjs  # 检评报告后处理②b/⑤（需样例 docx，见下）
  node tests/pjbg.note1009.test.mjs # 后处理⑤回归 + ②b（需 1009 报告样例；缺失自动跳过）
  node tests/pjbg.conddel.test.mjs # 后处理⑥ 条件删除条目 + 章节内序号重排（纯 XML 夹具，无需真实报告）
  node tests/pjbg.shiftred.test.mjs # 后处理⑦ 「两班制」作息说明句标红（纯 XML 夹具）
  node tests/pjbg.noisereuse.test.mjs # 噪声表「只解析一次」的复用/缓存/退化逻辑（纯夹具，无需真实报告）
  node tests/pjbg.noiseapi.smoke.mjs # 真实 DOMParser 下 noiseInfoFromXml 的分组/拆分/文本（Playwright）
  ```
- 测试注意：应用会自动把参考库写回 `data/library.json`；跑会改动参考库的测试前先备份，测试后 `git checkout -- sampling-plan-app/data/library.json` 恢复。
- `tests/pjbg.note.test.mjs` / `pjbg.note1009.test.mjs` 需要真实检评报告 docx 作样例（默认读 `%TEMP%/sampling-plan-login/`，可用 `PJBG_SAMPLE` / `PJBG_1009` 或 argv 指定）。纯 Node（`vm` 沙箱加载 `js/blf.js` + `JSZip`），不依赖浏览器；**blf.js 为此保留了 `SamplingUtil` 缺失时的内联兜底**。
- ⚠ 已知**预存在失败**（与近期改动无关，勿误判）：`tests/ui.smoke.mjs` 第 4 步「覆盖联动」（AR 覆盖为「个体」时期望 AS=长时间/AK=采样对象）——当前逻辑产出短时间/采样点，断言陈旧。
- 环境注意：本机系统 PATH 中可能没有 `node`（`启动采样计划软件.bat` 会按顺序回退到内置运行时）。可先取内置 Node 绝对路径再执行；`sampling-plan-app/node_modules` 已含 jszip 与 playwright-core。
- 快速自检：对改动文件跑 `node --check`（对全部 js/mjs 跑一遍也能一次性发现语法问题）。

## 部署

| 目标 | 触发方式 | 能力 |
|---|---|---|
| GitHub Pages（smallrookie404.github.io/sampling-plan） | push main → deploy-pages.yml | 静态工具可用；**无法调用平台 HTTP 接口**（HTTPS 混合内容被浏览器拦截） |
| Cloudflare Pages（sampling-plan.pages.dev） | push main → deploy-cloudflare.yml，或 部署到Cloudflare.bat | 静态工具 + Functions + KV；登录/验证码/上传**必须用此版本** |

Cloudflare 工作流需要 GitHub 仓库 Secrets：`CLOUDFLARE_API_TOKEN`、`CLOUDFLARE_ACCOUNT_ID`。KV 命名空间绑定 `SAMPLING_RECORDS`（wrangler.jsonc）。

## 关键机制与维护要点

1. **数据同步**：记录与参考库按上表三端存储契约写入，但**三端各自独立、互不同步**（每类数据只写「主端」，读取按 `updatedAt` 取最新，不回写另一端）；记录侧读取优先级为 GitHub > 本地服务 > localStorage 镜像。**改同步逻辑前先看「三端存储契约」一节**——重点别引入「开页即写」或「双写」（会烧 Cloudflare KV 的 1000 写/天额度）。
2. **统一登录**（upload.js）：平台登录（账号/密码 RSA 加密/验证码）成功后调 `auth/info` 获取权威用户信息并解析机构（顶层 `organizationId` 优先，其次 `organization.organizationId`），机构仅用于接口数据范围、**不在界面显示**。会话存 sessionStorage；登录遮罩默认显示、有会话时内联脚本立即隐藏，避免主界面闪现。
3. **团队配置云端保存**：成员/调查人/复核人/日期确认后，本地 localStorage 兜底 + Cloudflare KV（`/api/team`，按账号 `team:<userCode>`）同步，换电脑登录同一账号可读取；确认后切换项目不清空，手动修改才更新。
4. **平台代理**：`functions/api/platform/[[path]].js` 用 `cloudflare:sockets` 原始 TCP 连接平台 `http://223.93.144.122:27800`，透传请求头/响应头/Cookie 并解压 gzip，绕过 Cloudflare 直连 IP 限制；`upload.js` 在 pages.dev 上自动走 `/api/platform`，本地 http/file 直连平台。代理只透传固定头白名单（authorization / organizationId / belongProject / content-type / user-agent / accept / cookie），新增自定义头必须同步白名单。
5. **导出上传联动**：「数据上传」按钮生成当前表格的导出 Excel（`window.SamplingApp.exportWorkbookBytes()`）作为上传文件，无需手动选文件；导出范围到「检测项目」列最后一个非空单元格。
   - **企业信息（在职人数 / 接害人数）**：工具栏两个输入框 `#entOnJobNum` / `#entHazardNum`，随上传写入平台「现场调查 → 企业信息」（`enterpriseNum` / `hazardNum`，PUT 整对象回写）。
     · **上传分工**：主表格（测点布局调查）只带在职人数、调查表（6 合一）只带接害人数（`entPayloadForUpload`），另一侧平台原值不动。
     · **写入成功后自动清空**（2026-10-18 用户需求）：`applyEntInfoOnUpload` 的 `onSaved` 回调 → `clearEntInputsAfterUpload(patch, projectId)`，**只清本次真正写入的那一侧**（另一侧保留，因为另一个文件还没上传），并同步清空该侧在 `xcdc_entinfo_v1` 里的缓存值。失败/未写入一律不清（值还在输入框里，方便直接重传）。
6. **界面约定**：主界面样式变量在 `css/styles.css`（`--primary: #1f4e79` 等）；登录/上传视图样式统一用 `.xcdc` 作用域并复用同一套变量。新增页面元素需与 `upload.js` 中的 id 引用保持一致（上传页会一次性校验全部关键 id，缺任一个即 alert 并中止）。
7. **调查表模块**（survey.js）：6 个现场调查子表（生产工艺/设备设施/原辅物料/主要产品/职业防护/个体防护），页签位于「测点布局调查」旁。Excel 式交互与主表格对齐：单击选中、双击编辑、拖选矩形、Enter 提交并下移、Alt+Enter 或编辑态粘贴多行为单元格内换行（非编辑态粘贴按行拆分填入多行）；下拉多选仅「车间」与「岗位(工种)」列，其余固定选项列为单选，「单元/工作场所」动态取主表车间名，岗位下拉数据源为主表自动计算区 W→X（含下填）。数据存 localStorage 键 `samplingPlanSurvey_v2`（统一默认 10 行），随记录保存/还原；无 Tab 多行粘贴兼容。导入 Excel 支持调查表；导出 6 合一首表为「劳动定员和职业病危害因素接触情况调查」（首行＝主表自动计算区表头，无数据行），上传页进入时自动生成两个文件。主表无数据时仍可保存仅含调查表的记录（两者均空才拦截）。
8. **主表格编辑模型**：单击仅选中并「布防」隐形 textarea，输入/输入法经 `armReveal` 显形，双击进入编辑；Enter/Tab 提交并移动、Esc 还原、Delete 清空选区。粘贴统一走 `parseTsvGrid`：编辑态且无 Tab 时视为单元格内换行，否则按选区左上角平铺。`W–BI` 自动列可被覆盖的机制由 `overridden` 标记保留手动值，导入「仅自动计算区」文件时靠它反推录入区。
9. **危害因素库编辑模型**（app.js，`haz*` 前缀）：与主表录入区**完全同构**的 Excel 语义编辑器 —— 单击仅选中（布防隐形 `textarea.grid-arm`）、直接键入按替换语义覆盖、双击/F2 进入编辑、方向键/Enter/Tab/Home/End/PageUp/PageDown 导航、Delete/Backspace 清空选区、拖选矩形高亮（`sel`/`cur`/`hl-row`/`hl-col`）、Ctrl+A 全选、Ctrl+C 复制为 TSV、粘贴按选区左上角平铺、Escape 取消。
   - **★ 已取消全部下拉列**：危害因素库不再渲染任何 `<select>`，全列均为文本 `input`（`hazIsSelectCol()` 恒返回 `false`，保留函数以便未来扩展）。
   - 网格为**窗口化渲染**（只渲染可视行，行高恒为 `HAZARD_ROW_H=30`），因此 `hzCur/hzSel*` 一律用 hazardFactors 的**绝对下标**；跨行移动需经 `hazardDisplayList()`（显示顺序）换算，`hazTd()` 只在当前已渲染的行里取 DOM。
   - 布防框与滚动协同：`scroll` 事件在布防态下调 `hazSyncArmPosition()` 重定位（不做窗口化重绘），移出可视区才移除；`hazRevealCell()` 负责滚动到目标格并触发重绘。
   - 清空/还原等改动模型的操作必须**同步已渲染的 input DOM**（`hazClearRange`/`hazRevertCurrent`），否则窗口化渲染会残留旧值。
   - 列结构：`HAZARD_HEADERS`（14 项，首项「序号」不在 keys 中）+ `HAZARD_KEYS`（13 项，末项为 2026-10 新增的 `physchem`「物化性质」）；**键索引 = 表头索引 - 1**。新增列只需在两数组**同位置**追加，`blankHazard`/`renderHazard`/`hazGetValue`/`hazSetValue`/窗口化 `colspan` 均由 `HAZARD_KEYS` 派生，会自动生效。
   - 专用回归测试：`tests/hazard.edit.smoke.mjs`（24 项）。**注意**：`tests/ui.smoke.mjs` 在「覆盖联动」既有失败处提前退出，其 4a 危害因素库用例实际不可达，故危害库编辑一律以该专项测试为准。新增的末列可能超出视口宽度，测试点击前需先 `hazard-wrap.scrollLeft = scrollWidth`。
   - ⚠ **命名坑**：函数一律 `haz*` 前缀（`hazSetCur`/`hazRemoveArm`…），状态变量一律 `hz*` 前缀（`hzCur`/`hzArm`…）。两套前缀极易写错（曾出现 `hzSetCur`/`hzRemoveArm`/`hazArm` 混用导致运行时报 “is not defined”，且 `node --check` **无法**发现）。改动后请用脚本核对「`haz*` 调用均有 `function haz*` 定义」「`hz*` 标识符均有 `let/const` 声明」。
10. **检评报告后处理管线**（blf.js `patchPjbgDocx`）：下载到 docx 后**只读一次** `word/document.xml`，顺序跑 6 步、最后统一 `DEFLATE` 重打包；**任一步失败回退上一步结果/原始文件，不影响出报告**。步骤：① 高毒物品 0/1→否/是（`fixHighToxColumn`）② 非噪声岗位删「噪声」危害（`fixNonNoiseHazards`，口径 LEX 最大值 < 80）②b 氮氧化物措辞精简（`fixNitrogenOxide`）③ 正文措辞替换（`fixReportText` + `TEXT_FIXES`）④ 表1-1 注3 按备注列裁剪（`fixHazardNote`）⑤ 表4-1 注段落替换 + 表3-1 接触人数联动（`fixNoteAndContactCount`）⑥ **条件删除条目 + 章节内序号重排**（`fixConditionalDeletions`）⑦ **「两班制」作息说明句标红**（`markShiftNoteRed`）。
   - ⑥ 依据（用户 2026-10-17，2026-10-17 晚修正）：**噪声车间岗位为空**（`SamplingUpload.noiseWorkshopGroupsFromBuf` 返回 `[]`）→ 删 1.2.3 评价依据下 GB 23466-2025 条目、9.2 下「产生噪声的工作场所设置…」；**取「检测项目确认表」的「职业病危害因素」列**（`collectConfirmColTexts`）→ 按危害因素库**「系统名称」列**（字段 `name`）匹配（`lookupHazardField`）→ 判三组：
     · **粉尘**：`dust`「粉尘性质」列（总尘/呼尘）非空 **或** `physchem` 含「粉尘」→ 有粉尘；两者皆无 → 删粉尘组。
     · **毒物**：`physchem` 含「毒物」→ 有毒物；`physchem` 整列为空 → **未知，不删**（保守）；确认无毒物 → 删毒物组。
     · **高毒**：`highTox` 全部不是「是」→ 删《高毒物品目录》。
     · 粉尘毒物**俱无**（且毒物可判定）→ 再删 GB 18664-2025。
     · **9.2 第四段另外三句**（按「确认表里有没有这个因素」**独立**判定，与粉尘/毒物组无关）：
       - 无 **`physchem` 完全等于「粉尘、毒物」** 的因素 → 删「产生含有有毒物质的混合性粉(烟)尘的工作场所应设置…」（只有「粉尘」或只有「毒物」都不算）；
       - 无 **`name` 为「高温」** 的因素 → 删「高温工作场所设置…」；
       - 无 **`name` 为「紫外辐射」** 的因素 → 删「产生紫外辐射的工作场所设置…」。
       ⚠ 用**因素名**判高温/紫外辐射（不能用 physchem：云端把物理因素的 physchem 全填成了「紫外辐射」）。
   - ⚠⚠ ⑥ **判定口径（含 2026-10-17 晚修正的三处根因）**：
     1. **确认表要选对**：报告里「岗位/危害因素一览表」也有「职业病危害因素」列，早先「表头含该列名即命中」会抓错表（那列常是 `-` 占位）→ 改为**必须同时含「检测项目」列**优先，同级按行数多者优先。
     2. **占位值跳过**：`-`/`—`/`/`/`无` 表示「本行无因素」，`isHazPlaceholder` 直接跳过，**不再当成「查不到」**（否则一个 `-` 就让整组判定失效）。
     3. **名字匹配三级**：① 整格（避免把 `1,2-二氯乙烷` 按逗号拆坏）② NFKC 全半角归一（报告 `<10%` vs 库 `＜10%`）③ 去掉尾部 `(总尘)/(呼尘)`（库里用 `dust` 列区分，名字不带后缀）④ 最后才按 `,，、;；` 拆分。
     4. **条目要按行删，不能整段删**：真实报告里「1.1 法律法规」「1.2.3 评价依据」「9.2 建议」是**一个 `<w:p>` 内用 `<w:br/>` 软换行装十几行**（实测 1095/1009 都是），「命中即删整段」会把整节条目全删。现在 `deletePjbgParas` 先 `paraLineMap` 切行：整行就是该条目→删该行；只是长句里的一个分句→`trimClausesByKeywords` 切分句；并做**段内序号重排**。
     5. **重组必须就地替换段落**：`xml.replace(re, cb)` 而不是「把段落拼起来」——否则表格/`sectPr` 等**段落之外的内容会整块丢失**（1095 会掉 21 张表、约 1.14MB）。
     6. 关键词用**子串包含**；**缺数据不删**（库中找不到即视为证据不足，整组不动）；序号「（N）」重排**限定在所属章节内**（`isHeadingPara` 判标题边界——形如 `1.2.3评价依据`，**排除** `1.《xx》` 这类法规条目）。
     7. **性能**（2026-10-17 晚，1095 实测）：确认表单元格先 `uniqHazTexts` **去重**（140 格→15 种，匹配量降 89%）；库索引 `buildHazardIndex` **一轮只建一次**给 dust/physchem/highTox 三次查询复用（原来 3 次）；段落扫描加 `hasBr` 廉价预判——只有含 `<w:br/>` 的段落才建行映射（1095 共 6035 段，仅 78 段含 `<w:br/>`，即 1.29%），省掉约 12000 次全量行映射。三项合计约省 24ms/次（⑥ 总耗时约 138ms，其余是 7.7MB 文档的段落正则扫描，属固有开销）。
   - ⑦ **「两班制」作息说明句标红**（用户 2026-10-17）：把「…实行两班制，上班时间为白班…、晚班…，其余各岗位均实行白班制，上班时间为…」整句改成**红色字体**（`<w:color w:val="FF0000"/>`），提醒人工核对班制/作息。
     · 定位：在段落文本里找锚点 **`实行两班制`** → 向左回到上一个「。」之后、向右到下一个「。」（含），只给这一句上色；同段的「注：职工总数…接触人数…。」保持原色。
     · 锚点用短句而不是「打磨岗位…」：岗位名与时间随项目变化；实测两份真实报告里「两班制」都只出现 1 次。
     · `colorParaRange()` 按字符偏移切 run（前段/着色段/后段）；`setRPrColor()` **按 OOXML 子元素顺序**插入 `w:color`（必须在 `spacing/kern/sz/szCs/lang` 之前），否则 Word 可能报格式错误。
     · 保守回退：run 含多个 `<w:t>` 或 `<w:br/>` 时**原样放过**（宁可不标红也不破坏结构）。
     · 必须在 ⑤ 之后执行（⑤ 会重写「注：」段落）。
   - ⑦ 专项测试：`tests/pjbg.shiftred.test.mjs`（26 项：真实单 run 切分、前缀不上色、color 插入位置、无锚点不动、已有颜色替换、句尾无句号、跨 run 上色、含 br 放过）。
   - ⑥ 纯 XML 夹具测试：`tests/pjbg.conddel.test.mjs`（**72 项**，含「多行段落行级删除」「长句分句裁剪」两条真实形态用例，无需真实报告）。改后处理链路回归：`report.smoke`、`pjbg.note.test`（52）、`pjbg.note1009.test`（18）、`pjbg.conddel.test`（72）。
   - ⑧ **整条链路性能优化（2026-10-18）**——改动集中在「重复劳动」上，输出结果**逐字节不变**：
     1. **噪声表只解析一次**：②⑤⑥ 原来各自调一个 `noise*FromBuf(buf)`，每次都 `JSZip.loadAsync` 解压整份 docx + `DOMParser` 解析全文，**同一报告被完整解析 3 次**。现在 upload.js 提供 `noiseInfoFromXml(xml)` / `noiseInfoFromBuf(buf)`（带记忆）/ `getCachedNoiseInfo(buf)`，`patchPjbgDocx` 只算一次并复用给三步；且**优先复用 `patchPjbgDocx` 已解压出的 xml**，连解压都省掉。若用户先打开过「噪声数据结果」页签，`loadNoiseResult` 会顺手把结果写进缓存 → 导出时**零额外解析**。
     2. **扫表前先做单字预筛**：`readBasicInfoTable`（在册职工人数/接触职业病危害人数）、`findTableByHead`（新增可选第 3 参 `needChars`）、`collectConfirmColTexts`、`fixNoteParagraph`、`collectHazardMarks` 原先会遍历**全文每张表的每个单元格**，现在先用单字（`册`/`触`/`数`/`论`/`职`/`临`/`备`）跳过无关表。单字预筛是安全的——汉字即使被拆到多个 `<w:t>`，该字仍原样出现在原始 XML 里（与 `mapParasContaining` 的 needles 约定同源）。
     3. **`collectHazardMarks` 不再 `mapTables`**：原来用 `xml.replace` 全量重建一个 6MB 字符串却**丢弃返回值**（只读不改），改为 `tablesOf` 只读遍历。
     4. **`readZipEntryText` 用原生 `TextDecoder`** 替代 JSZip 自带的 JS 版 UTF-8 解码器（`typeof TextDecoder` 判空回退，Node 的 vm 沙箱里可能没有）。
     5. **不做的事**：`generateAsync` 的 DEFLATE 级别**保持默认 6**。实测 level 4 只省 ~46ms（9%）但体积 +2%，性价比不足。JSZip 对未改动条目会复用原压缩字节（图片不会被重压），所以重打包耗时几乎全在 `document.xml`（解压后 6MB+）的 deflate 上，属必要成本。
     · 实测（真实报告 `pjbg.docx` 1.88MB / document.xml 6.05MB）：6 步后处理 **352ms → 237ms（-33%）**；浏览器端另省掉 2 次整份解析（每次解压 ~180ms + DOMParser 数百 ms）。
     · 专项测试：`tests/pjbg.noisereuse.test.mjs`（**22 项**：只解析一次、缓存命中零解析、老接口退化、解析失败语义不变）+ `tests/pjbg.noiseapi.smoke.mjs`（**11 项**，Playwright：真实 DOMParser 下的分组/「/」拆分/文本格式/缓存只读/无表返回 ok=false）。
     · 基准脚本：`tests/_bench_pjbg.mjs`（端到端）。

## 已知注意事项

- **Token/凭据安全**：GitHub fine-grained PAT 与平台密码仅存本机浏览器（localStorage/sessionStorage），**严禁写入仓库或内置到代码**；公开仓库勿存放企业敏感调查数据。
- 平台接口为真实生产系统，上传/成员同步等写操作会影响真实数据，改动接口契约前先核对（请求体数组/对象格式易踩坑：成员新增传对象、成员删除传裸数字数组、上传用 FormData 且 query/body/请求头三处联动）。
- `data/library.json`、`data/records.json` 由应用自动写回，仓库中保持提交以利备份；`.gitignore` 已排除原版 Excel 模板（含调查数据）。
- `data/records.json` 目前仍是旧版「整表记录」格式（元素含 `rows`），首次访问 `/api/records` 时会自动迁移为索引 + `data/records/<id>.json`；迁移是幂等的，不要手工去改。
- **只有修改参考库**才会触发 `scheduleLibrarySync` 防抖写回（约 1.2s，只写一个主端）；**打开页面不再写参考库**。调试或跑测试前后仍建议备份与恢复 `data/library.json`。
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
| 改危害因素库交互（编辑/导航/选区/行操作） | app.js `haz*` 函数族（`hazSetCur`/`hazMoveCur`/`hazArmCell`/`hazArmReveal`/`hazClearRange`/`hazApplySelection`/`hazTd`/`hazardDisplayList`）+ `hz*` 状态 + `renderHazard`（列结构 `HAZARD_HEADERS`/`HAZARD_KEYS`）；行操作用 `hazardAnchor`/`afterHazardRowsChanged` |
| 改虚拟滚动/列宽/表头样式 | app.js `buildHead`/`cellHtml`/`rowHtml`/`renderWindow`（spacer 必须恒为首末子节点）+ `computeGridWidths` |
| 加/改调查表字段 | survey.js `SURVEY_SHEETS`（列名即键）+ `SURVEY_COL_OPTIONS`（下拉）；联动列用 `headers.indexOf` 定位，**改列名会使联动与旧 Excel 导入静默失效** |
| 改调查表导出/导入 | survey.js `exportAllWorkbook` / `importFromArray`（按工作表名+表头名匹配，兼容列序变化） |
| 改记录保存/读取/同步 | app.js `detectStorageMode`/`loadRecords`/`loadRecordById`/`persistRecord`/`deleteRecordById` + 三端实现（server.mjs、functions/api/records.js、app.js `github*`） |
| 改参考库同步 | app.js `loadLibrary`/`persistLibrary`/`scheduleLibrarySync` |
| 改检评报告后处理（0/1、措辞、注段落、条件删条目） | blf.js `patchPjbgDocx` 内 6 步：`fixHighToxColumn`/`fixNonNoiseHazards`/`fixNitrogenOxide`/`fixReportText`(+`TEXT_FIXES`)/`fixHazardNote`/`fixNoteAndContactCount`/`fixConditionalDeletions`(+`PJBG_DEL_TARGETS`/`deletePjbgParas`/`collectConfirmColTexts`/`lookupHazardField`)；噪声数据侧 upload.js `buildNoiseSummary`/`noiseInfoFromXml`/`noiseInfoFromBuf`/`getCachedNoiseInfo` |
| 改检评报告处理性能 | blf.js `readZipEntryText`/`readBasicInfoTable`/`findTableByHead`(needChars)/`collectHazardMarks`/`collectConfirmColTexts`/`fixNoteParagraph`；upload.js `noiseInfoFromXml`/`noiseInfoFromBuf`/`readDocxDocumentXml`。⚠ 新增「扫全表」逻辑时先加单字预筛 |
| 改通用工具（转义/懒加载/列表兼容） | js/util.js（`window.SamplingUtil`）；业务脚本只引用它，不要在 app/upload/survey/blf 里各写一份 |
| 改登录/验证码/会话 | upload.js `API_BASE`/RSA 实现/`resolveOrgId`/`saveSession` + index.html 登录遮罩与内联隐藏脚本 |
| 改上传流程或平台接口 | upload.js `apiRequest`/`runUploadFlow`/`syncProjectInfo`；代理侧 `functions/api/platform/[[path]].js`（目标 IP、头白名单） |
| 改在职人数/接害人数（企业信息） | upload.js `entInputValues`/`validateEntInputs`/`entPayloadForUpload`/`buildEntInfoPatch`/`applyEntInfoOnUpload`/`clearEntInputsAfterUpload`/`prefillEntInfo` + index.html `#entOnJobNum`/`#entHazardNum`；测试 `tests/entinfo.smoke.mjs` |
| 改团队配置 | upload.js `loadTeamSettings`/`saveTeamSettings`/`loadProjectTeam` + `functions/api/team.js`（KV key、64KB 上限） |
| 改样式/主题 | 主界面 `css/styles.css`（CSS 变量）；登录/上传视图是 index.html 内联 `.xcdc`（老内核兼容写法勿改成 `inset` 简写等新语法） |

## 常见任务速查

- 改了主表/上传功能 → `node --check sampling-plan-app/js/app.js sampling-plan-app/js/upload.js` + 跑对应测试。
- 改了危害因素库编辑 → `node --check sampling-plan-app/js/app.js` + `node tests/hazard.edit.smoke.mjs`（24 项）。
- 改了检评报告后处理 → `node --check sampling-plan-app/js/blf.js` + `node tests/pjbg.conddel.test.mjs`（72 项，纯夹具）+ `node tests/pjbg.noisereuse.test.mjs`（22 项，噪声复用）+（有样例时）`pjbg.note.test` / `pjbg.note1009.test` + `report.smoke`（41 项）。
- 改了计算引擎 → `node --check sampling-plan-app/js/logic.js` + `node tests/engine.test.mjs`（对照原 Excel 缓存）。
- 改了调查表 → `node --check sampling-plan-app/js/survey.js` + 跑 `tests/survey.smoke.mjs` / `tests/survey.edit.mjs`。
- 改了导入导出 → `node tests/xlsxio.test.mjs`（会产出 `tests/out.xlsx`，已被 .gitignore 排除）。
- 改了本地服务 → `node tests/server.test.mjs`。
- 改了 Cloudflare Functions → 本地用 `wrangler pages functions build` 验证语法，或直接依赖工作流部署后测 `https://sampling-plan.pages.dev/api/health`，再跑 `node tests/cloudflare.test.mjs`。
- 需要平台白名单 → 把 Cloudflare 出口 IP（https://www.cloudflare.com/ips/）提供给平台管理员。
