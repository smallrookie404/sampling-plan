# 任务交接：数据报告审核（4步跨账号审批链）

> **状态：✅ 已解决并真机验证（2026-10-15 00:1x）**
> 之前的「第 ② 步阻塞」根因已查明，是实现用错了接口与节点判定，而非平台无此接口。

## 任务目标
实现「数据报告审核」自动化（与人工在平台上操作顺序一致）：
1. **① btc-jyg**：检测底稿「提交审复核」
2. **② btc-wcy**：检测底稿 **理化复核** 通过
3. **③ btc-zhj**：检测结果报告单 **理化审核** 通过
4. **④ btc-gxz**：检测结果报告单 **理化签发** 通过

## 根因与正确做法（★ 核心结论）

之前 400 的**三个错误**（现已全部修正）：

| # | 错误做法 | 正确做法 |
|---|---|---|
| 1 | 通过类调 `POST /api/projectJcdgLastAffixNew/auditAll`（**数组**） | 调 **`audit`（单个对象 body）** |
| 2 | `type` 取「列表行的 `status`」 | `type` 取 **`getAffix` 返回的 `type`（当前节点号）** |
| 3 | jcbg 的 `type` 传字符串 | jcbg 的 `type` **必须是数字**（传 `"2"` → 400「Required request body is missing」） |

### 接口清单（实测）
| 用途 | 接口 | Body |
|---|---|---|
| 提交审复核 | `POST /api/projectJcdgLastAffixNew/submitJcdgReview` | `{belongProject}` |
| jcdg 读当前节点 | `GET /api/projectJcdgLastAffixNew/getAffix?belongProject=<id>` | → `body[0].type` |
| jcdg 通过 | `POST /api/projectJcdgLastAffixNew/audit` | `{belongProject, projectName, suggestions:"通过", type:<节点号 int>, isPass:"1", auditDate}` |
| jcbg 读当前节点 | `GET /api/projectJcbgLastAffix/getAffix?belongProject=<id>` | → `body[0].type` |
| jcbg 通过 | `POST /api/projectJcbgLastAffixNew/audit` | 同上（**type 必须 int**） |

> ⚠ 注意两个「读当前节点」的路径**不同名**：jcdg 用 `projectJcdgLastAffixNew/getAffix`，
> jcbg 用 `projectJcbgLastAffix/getAffix`（**没有 New**）。

### 为什么列表不可用
- jcdg 的列表 `projectJcdgLastAffixNew` 里**所有行 `status` 恒为 "0"**（实测 131 条全 0）→ 不能用来判断节点。
- jcbg 的列表 `status` 恰好 == getAffix 的 type，但为统一稳妥，**两者都改用 getAffix 的 type**。

## 节点号含义（前端 statusFn，权威）
```
0=理化复核  1=复核  2=理化审核  3=审核  4=职卫复核  5=职卫审核
6=理化签发  7=职卫签发  8=签发  9=第三方审核  10=签章  11=批准
```
流程节点序列见 `GET /api/ohitBusinessProcess?type=<业务类型>`（返回 `record.processOrder`）：
- **type=8（检测底稿 jcdg）** → `"0"`（**只有 1 个节点**：理化复核）
- **type=3（检测结果报告单 jcbg）** → `"2,6"`（理化审核 → 理化签发，**2 个节点**）
- type=26（检评报告 jpbg）→ `"0"`

> 这解释了为什么链是 4 步：jcdg 1 个节点 + jcbg 2 个节点 + 1 次提交 = 4。

## 前端逆向来源（`%TEMP%/sampling-plan-login/chunks2/`）
- 路由 `/projectAuditReviewNewPage/page` → 模块 `f091`（`chunk-161ad49e.js`）。
- 模块 `66cb` 导出：`onReviewJcdg = POST /api/projectJcdgLastAffixNew/audit`（**单对象**）、
  `onReviewAllNew = .../auditAll`。→ 通过用的是 **audit**。
- 模块 `1ed4` 导出：`onReviewNew = POST /api/projectJcbgLastAffixNew/audit`（**单对象**）。
- 详情页 `handleAuditReviewPass`：jcdg 分支设置 `form.type = processStatus`（当前节点），
  再 `submitCU()` → `reportOperation()` → `onReviewJcdg(this.form)`。

## 真机验证记录（BTC26-SZJC1009，id=6255）
| 步 | 账号 | 操作 | 结果 |
|---|---|---|---|
| ① | btc-jyg | submitJcdgReview | ✅ 已提交 |
| ② | btc-wcy | jcdg audit，type=`0`(理化复核) | ✅ `{"code":"200","message":"审批成功！"}`；历史 id=22717 type=0 isPass=1 createBy=62 |
| ③ | btc-zhj | jcbg audit，type=`2`(理化审核) | ✅ 成功；节点推进到 `6`(理化签发) |
| ④ | btc-gxz | jcbg audit，type=`6`(理化签发) | ✅ 成功；`isOver=1`，项目从 isDel=0 列表消失、进入 **isDel=1（已完结）** |

## 代码位置
| 文件 | 说明 |
|---|---|
| `sampling-plan-app/js/upload.js` | `AUDIT_STEPS` / `auditPaths()` / `readAffixNode()` / `pollNodeChange()` / `nodeName()` / `auditDateStr()` / `auditReport()` |
| `sampling-plan-app/tests/report.smoke.mjs` | 冒烟测试（mock getAffix 状态机），**15/15 通过** |
| `sampling-plan-app/functions/api/bridge/[[path]].js` | 云端账号桥 |
| `captcha-ocr/plat.py` / `pool.py` / `bridge.py` | 平台客户端 / 账号池 / 本地账号桥 |

## 测试账号
- btc-jyg（金云龚）/ btc-wcy（王慈媛）/ btc-zhj（周海江）/ btc-gxz（高循洲），orgId=1，口令本地存 `accounts.json`
- 云端 `BRIDGE_ACCOUNTS` 已含 4 账号，部署于 `sampling-plan.pages.dev`

## 遗留事项
- `tests/cloudflare.test.mjs` 有一处**预存在**失败（"初始 GET 不符"），与本次改动无关（本次未触碰 KV 记录模块）。
- 如后续流程配置（`ohitBusinessProcess`）变更节点序列，`NODE_NAMES` 映射表需同步。
