# 验证码 OCR + 账号池原型（captcha-ocr）

> 独立原型，**不属于采样计划应用本体**，不参与构建与部署。
> 建立时间：2026-10-14 ｜ 状态：**端到端跑通**

## 一、结论

### OCR 识别（60 张人工标注，留一法 LOO）

| 指标 | 值 |
|---|---|
| 字符分割成功率 | **60/60 = 100%**（每图稳定切出 5 个字符） |
| 单字符准确率 | **92.2%**（第1位 93.3% / 运算符 96.7% / 第2位 86.7%） |
| 整式全对率 | **76.7%** |
| 未见样本抽查 | 5 张中 4 张正确 |

### ★ 置信度门控（关键增益）

单字符"最优 vs 次优"的距离比作为置信度，**低于阈值就丢弃换图**（不消耗登录尝试）：

| 阈值 | 接受率 | 接受后正确率 |
|---|---|---|
| 不设 | 100% | 76.7% |
| 0.20 | 73.3% | 90.9% |
| **0.30** | 65.0% | **97.4%** |

默认阈值取 **0.25**（接受率约 70%，正确率约 95%）。

### 账号池（端到端实测通过）

```
$ py pool.py login btc-wf
  [btc-wf] 登录成功（第1张，7×5=35，conf 0.74）
  令牌校验 /auth/info → HTTP 200 王飞（BTC-WF）
```

- ✅ 第 1 张验证码即识别正确并登录成功
- ✅ 令牌落盘，**重复 acquire 消耗 0 张验证码**（实测 9 → 9）
- ✅ 验证码用量按天计数（`pool.py budget`）

## 二、验证码结构（实测）

- 固定 **111×36** PNG，base64 内嵌在 `/auth/code` 的 JSON（`{uuid, img}`）
- 内容是算术题，格式恒为 **`数字 运算符 数字 = ?`**，共 **5 个字符**；已见 `+` `-` `×`
- **每个字符是一种独立主色**、背景纯白、**无干扰线** → 分割几乎免费
- ⚠️ 但**每个字符形变随机且夸张**（手绘扭曲字体）——识别难度的主要来源

## 三、★ 平台接口实测要点（踩过的坑）

| 项 | 实测结论 |
|---|---|
| `POST /auth/login` 请求体 | `{username, password(RSA PKCS#1 v1.5), code, uuid}`（前端还带 `systemType`） |
| 登录**成功**响应 | `{"user":{...},"token":"Bearer eyJ..."}` —— **没有 `code` 字段**；判定请用 `token` 是否存在 |
| **token 自带 `Bearer ` 前缀** | 返回什么样就原样发送，**不要再自己拼 Bearer** |
| 验证码错误 | HTTP **400**，body 是 JSON（失败响应不在 2xx，必须捕获 `HTTPError` 才能读到 body） |
| `GET /auth/info` | 需 `Authorization: <token>` + `organizationId`，返回 `user.userName` 等 |
| RSA 公钥 | `MFwwDQYJKoZIhvcNAQEBBQADSwAwSAJBANL378k3RiZHWx5AfJqdH9xRNBmD9wGD2iRe41HdTNF8RUhNnHit5NpMNtGL0NPTSSpPjjI1kJfVorRvaQerUgkCAwEAAQ==` |

## 四、账号池

### 当前池内账号（均已登入并校验通过）

| 账号 | 姓名 | 状态 |
|---|---|---|
| `btc-wf` | 王飞（BTC-WF） | 令牌有效 |
| `btc-gxz` | 高循洲（BTC-GXZ） | 令牌有效 |
| `btc-zhj` | 周海江（BTC-ZHJ） | 令牌有效 |
| `btc-wcy` | 王慈媛（BTC-WCY） | 令牌有效 |

> 口令**不写入仓库**：真实凭据只存在于本地 `accounts.json`（**已 gitignore**），仓库仅保留 `accounts.example.json` 占位。

### 命令

```bash
py pool.py status                    # 列出账号与登入状态
py pool.py check                     # 探测各令牌是否仍有效（走 /auth/info，不消耗验证码）
py pool.py keepalive [间隔秒] [轮数]  # 巡检保活：有效的保留、失效的重登（不主动登出）
py pool.py warm                      # 预热：把没登入的账号都登入一次（消耗验证码）
py pool.py add <账号> <密码> [orgId] [备注]   # 新增账号（登录界面手动输入的账号走这里）
py pool.py adopt <账号> <token> [orgId]      # 接管外部（界面登录）拿到的令牌，避免重复登录
py pool.py login <账号>               # 强制登录（调试用）
py pool.py budget                     # 今日验证码用量
```

### 如何做到「有效期内存活、不退出、减少登录次数」

| 你的要求 | 实现 |
|---|---|
| **账号有效期内不退出** | ① 池内**从不调用 `/auth/logout`**；② 令牌落盘 `tokens.json`，进程重启不丢；③ `keepalive` 只在探测到失效时才重登，有效的一律原样保留 |
| **后台留存** | 令牌持久化 + `keepalive` 可循环运行（`py pool.py keepalive 600 144` = 每 10 分钟巡检一次、共 144 轮≈24h） |
| **减少登录次数** | **惰性登录**：`acquire()` 命中缓存直接返回（实测**消耗 0 张验证码**）；只有令牌真失效才重登 |
| **登录界面手动输入的账号** | `add` 入池（长期保留）；或 `adopt` 直接接管界面已拿到的令牌，**连一次登录都不用** |

### 成本估算

平台 token 有效期约 **半天**。所以稳态下每个账号每天约需 **2 次重登**（= 2 张验证码）。
4 个账号 ≈ **8 张/天**；若池子扩到 20 个账号 ≈ 40 张/天。若验证码有硬性日限，请据此控制池子规模，
或降低 `keepalive` 频率（失效了才重登，不巡检也不会更费）。

### 核心 API（给程序调用）

```python
import pool
t = pool.acquire('BTC-WF')      # 有缓存直接用；过期才登录（惰性）
pool.invalidate('BTC-WF')       # 令牌失效（如接口返回 401）后标记，下次用会自动重登
pool.keepalive()                # 巡检保活
pool.budget()                   # 今日验证码用量
```

---

## 五、★ 实际取数流程（已接通）

`plat.py` 是平台接口客户端，**令牌全部来自账号池**，并在 **401 时自动失效→重登→重试一次**。
任何脚本/服务都通过它取数，**无需手动登录**。

### 用法

```bash
py fetch.py <项目编号> [--account btc-wf] [--out 输出目录] [--no-images]

# 例：把 BTC26-SZJC0959 整包拉下来
py fetch.py BTC26-SZJC0959 --account btc-gxz
```

### 实测（项目 BTC26-SZJC0959 / id=6205 / 浙江永贵电器）

```
✅ 搜索项目            项目 id=6205  BTC26-SZJC0959  浙江永贵电器股份有限公司
✅ 劳动定员 export331 (xlsx)     ✅ 生产工艺   ✅ 设备设施   ✅ 原辅物料
✅ 主要产品             ✅ 职业防护   ✅ 个体防护
✅ 工艺图 7 张          ✅ 布点图 3 张         ✅ 噪声采样元数据
完成 → out/BTC26-SZJC0959    文件 9 个，错误 0 个
今日验证码用量：15 → 15   ← 令牌命中缓存，0 消耗
```

产物结构：

```
out/BTC26-SZJC0959/
├── BTC26-SZJC0959_劳动定员.xlsx
├── survey_productionProcess.json   survey_eqpLayoutInfo.json
├── survey_rawMaterialsInfo.json    survey_mainProduct.json
├── survey_occupational_0.json      survey_occupational_1.json
├── noise_sampleInfo.json
├── images/                         （7 张工艺图 + 3 张布点图）
└── manifest.json                   （项目信息 + 文件清单 + 错误清单）
```

### 401 自动恢复（实测）

把令牌改成无效值后调用：

```
[btc-wf] 令牌失效(401)，重新登入后重试…
[btc-wf] 登录成功（第1张，5+2=7，conf 0.78）
恢复成功 → 王飞 BTC-WF
今日验证码用量：15 → 16      ← 恰好 1 张，用于恢复
```

### 已实现的接口

| 方法 | 接口 |
|---|---|
| `find_projects/find_project` | `GET /api/reportData/findList` |
| `worker331` | `POST /api/newEvaluationUnit/export331`（xlsx 字节） |
| `production_process` / `eqp_layout` / `raw_materials` / `main_product` | `GET /api/<res>` |
| `occupational_product(pid, 0\|1)` | `GET /api/occupationalProduct`（按 `occupationalType` 过滤） |
| `project_layout` | `GET /api/samplingPlan/projectLayout` |
| `download(path)` | 图片（逐段编码 + Authorization） |
| `noise_sample_info` / `noise_details` | `GET /api/noiseSampleInfo` / `siteNoiseHazardSampleInfos` |
| `pjbg_modules` / `gen_pjbg` | 检评报告模板与生成 |

### 设计如何对上三条约束

| 约束 | 对应设计 |
|---|---|
| 会话不会互踢 | 可同时持有多个账号的有效令牌，各自独立 |
| 员工已明确授权 | 凭据存 `accounts.json`（已 gitignore） |
| **验证码限量 + 登入后停止请求** | 惰性登录 + 置信度门控 + 令牌落盘 + 串行登录 + 按天用量统计 |

## 五、文件

| 文件 | 作用 |
|---|---|
| `fetch.mjs` | 采集验证码（`node fetch.mjs [数量] [起始]`） |
| `analyze.py` | 结构诊断（颜色/字符数/尺寸） |
| `segment3.py` | 二值掩码 + 列间隙切分 + 强制归 5 段 |
| `classify.py` | **核心**：特征提取 + LOO 评估（内含 60 张人工标注 `LABELS`） |
| `conf_eval.py` | 置信度阈值收益分析 |
| `ocr.py` | OCR CLI（`py ocr.py <图片>` / `--fetch`），输出 JSON |
| `pool.py` | **账号池**（`login` / `status` / `check` / `keepalive` / `warm` / `add` / `adopt` / `budget`） |
| `plat.py` | **平台接口客户端**（令牌取自账号池，401 自动重登重试） |
| `fetch.py` | **实际取数流程**（按项目编号整包拉取） |
| `accounts.example.json` | 账号配置模板（复制为 `accounts.json`） |
| `segments_preview.png` / `verify5.png` | 分割效果 / 识别验证 |

**已 gitignore**：`accounts.json`、`tokens.json`、`captcha-budget.json`、`samples/`、`out/`、`__pycache__/`

> ⚠️ 模块名不要用 `platform.py` —— 会与 Python 标准库 `platform` 冲突，故命名为 `plat.py`。

## 六、环境注意

- Python：`py`（Pillow + numpy 已具备）
- ⚠️ **Node 无法创建子进程**（本机沙箱限制，`spawnSync` 任何命令都 EBUSY）⇒ 因此账号池用**纯 Python** 实现，不经子进程调用 OCR。

## 七、提升准确率的路径（按性价比）

1. **加数据**（最确定）：现仅 60 张标注。采集免费，瓶颈只在标注 ⇒ 做**半自动标注**（用现分类器预标注，人工只改错的），扩到 500~1000 张，预计整式全对率 90%+
2. **换小 CNN**：同数据下通常再涨 3~5 个百分点
3. **多尺度特征集成**：对同一字符取多个模糊核/尺寸的特征拼接
4. **失败重试**（零成本，已内置）

## 八、⚠️ 仍需注意

原型已按你确认的前提实现（会话不互踢 / 已授权 / 验证码限量）。仍建议：

- 凭据文件（`accounts.json`）**不要入库、不要进日志**；
- 账号**逐个登记授权**，保留授权记录；
- 若某账号被平台提示异常登录，及时 `invalidate` 并暂停该账号。
