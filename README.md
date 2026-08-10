# 系统测点布局调查 · 采样计划软件

职业卫生检测“采样点布局调查”自动化工具，把原来的 Excel 公式表升级为可独立运行的本地软件。

## 仓库结构

```text
├─ sampling-plan-app/        软件（网页界面 + 计算引擎 + 本地服务）
│  ├─ index.html             软件入口
│  ├─ server.mjs             本地服务（静态页面 + 数据文件读写）
│  ├─ data/records.json      数据记录（数据库文件，随仓库同步到 GitHub）
│  └─ tests/                 自动化测试
├─ 同步数据库到GitHub.bat    把 data/records.json 提交并推送到 GitHub
└─ .gitignore
```

## 在本机运行

进入 `sampling-plan-app`，双击 `启动采样计划软件.bat`，软件会自动启动本地服务并打开浏览器。数据保存到 `sampling-plan-app\data\records.json`。

## 部署到 GitHub

本仓库已初始化为 git 仓库并完成首次提交，数据库文件 `sampling-plan-app/data/records.json` 已在版本管理中。

首次推送（需要你的 GitHub 账号，执行一次即可）：

```bash
git remote add origin https://github.com/<你的用户名>/<仓库名>.git
git branch -M main
git push -u origin main
```

之后每次在软件里保存了数据，双击 **同步数据库到GitHub.bat**，即可把 `data/records.json` 提交并推送到 GitHub，实现数据库云端备份。

> 提示：
> - 原版 Excel 模板默认未入库（见 `.gitignore`），如需要一起提交请删除 `.gitignore` 中对应行；
> - 仓库为私有或公开由你的 GitHub 设置决定，公开仓库请不要包含敏感调查数据；
> - 首次推送时 git 会弹出 GitHub 登录授权窗口（Git Credential Manager），按提示完成即可。

## 部署到 GitHub Pages（静态网页）

软件核心功能（计算、编辑、校验、导出）全部在浏览器端，可以发布为静态网页；数据保存通过 **GitHub API** 直接读写仓库中的 `sampling-plan-app/data/records.json`，数据库仍然在 GitHub 上。

1. 把仓库推送到 GitHub（main 分支已包含 `.github/workflows/deploy-pages.yml` 自动部署工作流）；
2. 在仓库 Settings → Pages 中把 Source 设为 **GitHub Actions**，等待部署完成；
3. 打开 `https://<你的用户名>.github.io/<仓库名>/`；
4. 点软件右上角「**GitHub 配置**」，填写仓库（用户名/仓库名）、分支、以及一个 **Fine-grained PAT**（仅授予该仓库 Contents 读/写权限）；
5. 之后「保存数据」会直接更新 GitHub 仓库里的 `records.json`，与本机版的文件夹存储行为一致。

> 安全提示：
> - Token 只保存在你本机的浏览器中，**切勿提交到仓库或告诉他人**；
> - 公开仓库请谨慎存放企业调查数据；
> - 静态版没有本地服务，无法写本机文件夹，数据以 GitHub 仓库中的 records.json 为准。

## 数据记录与 GitHub 同步（本地版 / 网页版通用）

在软件右上角「GitHub 配置」中填写仓库与 Token 后：

- **打开软件时自动从 GitHub 同步**最新记录（打开即同步，顶部有同步状态提示）；
- 「保存数据」直接写回 GitHub 仓库的 `sampling-plan-app/data/records.json`，同时镜像到本地作为离线兜底；
- 「数据记录」弹窗中的「从 GitHub 刷新」可手动拉取最新数据；
- Token 仅保存在本机浏览器中，请勿提交到仓库或泄露。

## 部署到 Cloudflare（dash.cloudflare.com）

软件可以部署到 Cloudflare Pages，数据记录存入 **Cloudflare KV**（不需要依赖 GitHub）。

仓库已包含：

- `sampling-plan-app/functions/`：Cloudflare Pages Functions（`/api/health`、`/api/records`，读写 KV）；
- `wrangler.jsonc`：Pages 部署配置（含 KV 绑定占位）；
- `部署到Cloudflare.bat`：一键登录、创建 KV、部署脚本。

部署步骤（首次需要你的 Cloudflare 账号登录一次）：

1. 双击 **部署到Cloudflare.bat**，按提示在浏览器中完成 Cloudflare 登录（或先设置环境变量 `CLOUDFLARE_API_TOKEN` / `CLOUDFLARE_ACCOUNT_ID`，可跳过登录）；
2. 若 `wrangler.jsonc` 中的 KV id 仍是占位符，脚本会创建 KV 命名空间 `SAMPLING_RECORDS` 并提示你填入；
3. 脚本把 `sampling-plan-app/functions/` 编译成 Worker（`_worker.js`），并把静态页面与 `data/` 组装到 `.wrangler\pages-dist`；
4. 脚本执行 `npx wrangler pages deploy .wrangler\pages-dist --project-name sampling-plan`，`wrangler.jsonc` 中的 KV 绑定（`SAMPLING_RECORDS`）会自动生效；
5. 完成后访问 `https://sampling-plan.pages.dev`，软件会自动识别「服务模式」：数据记录保存到 Cloudflare KV，打开页面时自动同步读取。

> 注意：`wrangler pages deploy` 的旧式参数（`--branch`、`--commit-dirty`）会关闭 Functions 编译，本脚本已避免使用；若手动部署请保持 `wrangler pages deploy .wrangler\pages-dist --project-name sampling-plan` 的写法。

> 提示：Cloudflare KV 为最终一致，保存后建议刷新「数据记录」确认；如需强一致数据库可改用 D1（SQLite），需要时我可以再适配。
