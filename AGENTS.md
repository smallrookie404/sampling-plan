# 采样计划软件（sampling-plan）维护指南

## 项目简介

职业卫生检测「系统测点布局调查 · 采样计划软件」：把原 Excel 公式表升级为可独立运行的网页工具。核心功能（计算、编辑、校验、导出）全部在浏览器端完成；并整合了「现场调查 Excel 上传」工具（原 zhiyeweishen 项目），平台登录界面同时作为整个软件的总登录门禁。

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
    ├── data/
    │   ├── records.json            # 数据记录（随仓库同步）
    │   └── library.json            # 危害因素库 + 检测项目（应用自动写回，勿手工改格式）
    ├── js/
    │   ├── data.js                 # 内置危害因素库/检测项目（由 scripts/extract-data.mjs 生成，勿手改）
    │   ├── logic.js                # 计算引擎（浏览器/Node 通用，纯函数）
    │   ├── xlsxio.js + jszip.min.js# xlsx 读写（懒加载，首次导入/导出时加载）
    │   ├── app.js                  # 主工具界面逻辑；暴露 window.SamplingApp（导出/错误数）
    │   ├── survey.js               # 调查表模块（6 个现场调查子表，Excel 式编辑）；暴露 window.SurveySheets
    │   └── upload.js               # 统一登录 + 现场调查上传（平台接口、团队配置、云端同步）
    ├── functions/api/              # Cloudflare Pages Functions（编译为 Worker）
    │   ├── records.js / library.js # KV 数据读写（binding: SAMPLING_RECORDS）
    │   ├── team.js                 # 团队配置云端存取（按账号分 key team:<账号>）
    │   └── platform/[[path]].js    # 平台接口同源代理（cloudflare:sockets 原始 TCP 转发）
    └── tests/                      # 自动化测试（Node + Playwright）
```

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

## 部署

| 目标 | 触发方式 | 能力 |
|---|---|---|
| GitHub Pages（smallrookie404.github.io/sampling-plan） | push main → deploy-pages.yml | 静态工具可用；**无法调用平台 HTTP 接口**（HTTPS 混合内容被浏览器拦截） |
| Cloudflare Pages（sampling-plan.pages.dev） | push main → deploy-cloudflare.yml，或 部署到Cloudflare.bat | 静态工具 + Functions + KV；登录/验证码/上传**必须用此版本** |

Cloudflare 工作流需要 GitHub 仓库 Secrets：`CLOUDFLARE_API_TOKEN`、`CLOUDFLARE_ACCOUNT_ID`。KV 命名空间绑定 `SAMPLING_RECORDS`（wrangler.jsonc）。

## 关键机制与维护要点

1. **数据同步**：records.json / library.json 可在「GitHub 配置」（fine-grained PAT，仅存浏览器）下经 GitHub Contents API 写回仓库；Cloudflare 部署下经 `/api/records`、`/api/library` 写 Cloudflare KV；本地服务写项目文件夹。三方读取时取较新一端并互相同步。
2. **统一登录**（upload.js）：平台登录（账号/密码 RSA 加密/验证码）成功后调 `auth/info` 获取权威用户信息并解析机构（顶层 `organizationId` 优先，其次 `organization.organizationId`），机构仅用于接口数据范围、**不在界面显示**。会话存 sessionStorage；登录遮罩默认显示、有会话时内联脚本立即隐藏，避免主界面闪现。
3. **团队配置云端保存**：成员/调查人/复核人/日期确认后，本地 localStorage 兜底 + Cloudflare KV（`/api/team`，按账号 `team:<userCode>`）同步，换电脑登录同一账号可读取；确认后切换项目不清空，手动修改才更新。
4. **平台代理**：`functions/api/platform/[[path]].js` 用 `cloudflare:sockets` 原始 TCP 连接平台 `http://223.93.144.122:27800`，透传请求头/响应头/Cookie 并解压 gzip，绕过 Cloudflare 直连 IP 限制；`upload.js` 在 pages.dev 上自动走 `/api/platform`，本地 http/file 直连平台。
5. **导出上传联动**：「数据上传」按钮生成当前表格的导出 Excel（`window.SamplingApp.exportWorkbookBytes()`）作为上传文件，无需手动选文件；导出范围到「检测项目」列最后一个非空单元格。
6. **界面约定**：主界面样式变量在 `css/styles.css`（`--primary: #1f4e79` 等）；登录/上传视图样式统一用 `.xcdc` 作用域并复用同一套变量。新增页面元素需与 `upload.js` 中的 id 引用保持一致。
7. **调查表模块**（survey.js）：6 个现场调查子表（生产工艺/设备设施/原辅物料/主要产品/职业防护/个体防护），页签位于「测点布局调查」旁。Excel 式交互与主表格对齐：单击选中、双击编辑、拖选矩形、Enter 提交并下移、Alt+Enter 或编辑态粘贴多行为单元格内换行（非编辑态粘贴按行拆分填入多行）；下拉多选仅「车间」与「岗位(工种)」列，其余固定选项列为单选，「单元/工作场所」动态取主表车间名，岗位下拉数据源为主表自动计算区 W→X（含下填）。数据存 localStorage 键 `samplingPlanSurvey_v2`（统一默认 10 行），随记录保存/还原；无 Tab 多行粘贴兼容。导入 Excel 支持调查表；导出 6 合一首表为「劳动定员接触情况调查」，上传页进入时自动生成两个文件。主表无数据时仍可保存仅含调查表的记录（两者均空才拦截）。

## 已知注意事项

- **Token/凭据安全**：GitHub fine-grained PAT 与平台密码仅存本机浏览器（localStorage/sessionStorage），**严禁写入仓库或内置到代码**；公开仓库勿存放企业敏感调查数据。
- 平台接口为真实生产系统，上传/成员同步等写操作会影响真实数据，改动接口契约前先核对（请求体数组/对象格式易踩坑）。
- `data/library.json`、`data/records.json` 由应用自动写回，仓库中保持提交以利备份；`.gitignore` 已排除原版 Excel 模板（含调查数据）。
- 源文件统一 UTF-8；页面迭代频繁，部署后浏览器需 Ctrl+F5 强刷。

## 常见任务速查

- 改了主表/上传功能 → `node --check sampling-plan-app/js/app.js sampling-plan-app/js/upload.js` + 跑对应测试。
- 改了调查表 → `node --check sampling-plan-app/js/survey.js` + 跑 `tests/survey.smoke.mjs` / `tests/survey.edit.mjs`。
- 改了 Cloudflare Functions → 本地用 `wrangler pages functions build` 验证语法，或直接依赖工作流部署后测 `https://sampling-plan.pages.dev/api/health`。
- 需要平台白名单 → 把 Cloudflare 出口 IP（https://www.cloudflare.com/ips/）提供给平台管理员。
