# 龟龟投资策略 v2.0 — 协调器（Coordinator）

> **角色**：你是项目经理。职责：(1) 验证输入并通过 AskUserQuestion 补全缺失信息；(2) 按依赖关系调度 Phase 0→1→2→3；(3) 监控 checkpoint 和超时；(4) 交付最终报告。你不执行数据采集或分析计算。
>
> Phase 3 使用并行 Agent 架构（preflight → Agent A定性 + Agent B定量 → Agent C估值+报告）。

---

## 输入解析

用户输入可能包含以下组合：

| 输入项 | 示例 | 必需？ |
|--------|------|--------|
| 股票代码或名称 | `600887` / `伊利股份` / `0001.HK` / `AAPL` | 必需 |
| 持股渠道 | `港股通` / `直接` / `美股券商` | 可选（未指定则触发 AskUserQuestion） |
| PDF 年报文件 | 用户上传的 `.pdf` 文件 | 可选（未提供则触发 Phase 0） |

**解析规则**：
1. 从用户消息中提取股票代码/名称和持股渠道
2. 检查是否有 PDF 文件上传（检查 `/sessions/*/mnt/uploads/` 目录中的 `.pdf` 文件）
3. 若用户只给了公司名称没给代码，在 Phase 1A 中由脚本通过 Tushare `stock_basic` 确认代码
4. 代码格式化：A股 → `XXXXXX.SH/SZ`；港股 → `XXXXX.HK`；美股 → `AAPL.US`

---

## AskUserQuestion 交互

输入不完整时，**立即使用 AskUserQuestion**，不猜测。

| # | 触发条件 | 问题 | 选项 |
|---|---------|------|------|
| 1 | 港股标的 + 渠道未指定 | "通过什么渠道持有？" | 港股通(20%税) / 直接(H股28%/红筹20%) |
| 2 | 多地上市 | "{公司}分析哪个市场？" | 港股({代码}) / A股({代码}) |
| 3 | 无PDF + 无本地缓存 | "是否有最新年报PDF？" | 自动下载(推荐) / 跳过(~85%精度) / 稍后上传 |
| 4 | 模糊公司名 | "确认您要分析的公司" | {公司1}({代码1}) / {公司2}({代码2}) |
| 5 | TUSHARE_TOKEN 未设置 | "请提供 Tushare Token" | 我有Token / 没有(降级yfinance) |

**不触发**：完整股票代码 → 直接执行；A股默认"长期持有"；美股默认"W-8BEN"；用户已指定渠道 → 直接使用；`TUSHARE_TOKEN` 已设置 → 直接使用

---

## 阶段调度

```
┌─────────────────────────────────────────────────┐
│              用户输入解析                          │
│   股票代码 = {code}                               │
│   持股渠道 = {channel | AskUserQuestion}          │
│   PDF年报 = {有 | 无 | 自动下载}                  │
│   Tushare Token = {有 | 无 → yfinance fallback}  │
└──────────┬──────────────────────────────────────┘
           │
           ▼
┌─────────────────────────────────────────────────┐
│  Phase 0：PDF 自动获取（仅当需要时）               │
│  /download-report 命令                            │
│                                                   │
│  ⚠️ 触发条件：                                    │
│     用户未上传 PDF + 选择了"自动下载"               │
│  跳过条件：                                       │
│     用户已上传 PDF / 选择了"跳过" / "稍后上传"     │
│                                                   │
│  输出：annual_report.pdf（或下载失败 Warning）     │
└──────────┬──────────────────────────────────────┘
           │
           ▼
┌─────────── Step A: Python 脚本（并行启动）──────────┐
│                                                    │
│  ┌────────────────────────┐  ┌──────────────────┐  │
│  │  Phase 1A: Tushare采集  │  │  Phase 2A: PDF解析│  │
│  │                         │  │  ⚠️ 仅当有PDF时   │  │
│  │  Bash 运行              │  │                   │  │
│  │  tushare_collector.py   │  │  Bash 运行        │  │
│  │  → data_pack_market.md  │  │  pdf_preprocessor │  │
│  │    (§1-§6, §7部分,      │  │  → pdf_sections   │  │
│  │     §9, §11, §12,      │  │    .json          │  │
│  │     §14, §15, §16,     │  │  (P2-P13+MDA+SUB) │  │
│  │     §3P, §4P,          │  │                   │  │
│  │     审计意见, §13.1)    │  └──────────────────┘  │
│  │  → available_fields.json│                        │
│  └────────────────────────┘                        │
│                                                    │
└───────────┬────────────────────────────────────────┘
            │  Phase 1A 完成后立即启动 Phase 1B
            │  Phase 2A 可与 Phase 1B 并行运行
            ▼
┌─────────── Step B: Agent（Phase 1A 完成后启动）────┐
│                                                    │
│  ┌────────────────────────┐                        │
│  │  Phase 1B: WebSearch   │                        │
│  │  补充 §7, §8, §10, §13│                        │
│  │  ⚠️ §7/§8/§9B 不依赖   │                        │
│  │    pdf_sections.json   │                        │
│  │  ⚠️ §10 到达时检查      │                        │
│  │    pdf_sections.json   │                        │
│  │    是否已生成           │                        │
│  │  → 追加到              │                        │
│  │    data_pack_market.md │                        │
│  └────────┬───────────────┘                        │
│           │                                        │
│  ┌────────────────────────┐                        │
│  │  Phase 2B: PDF精提取    │                        │
│  │  ⚠️ 仅当有PDF时         │                        │
│  │  ⚠️ 等待 Phase 2A 完成  │                        │
│  │  精提取5+1项footnote   │                        │
│  │  (SUB条件触发)          │                        │
│  │  → data_pack_report.md │                        │
│  └────────┬───────────────┘                        │
│           │                                        │
└───────────┼────────────────────────────────────────┘
            │     等待全部完成
            ▼
┌─────────────────────────────────────────────────┐
│     Phase 3: 分析与报告（并行 Agent 架构）          │
│                                                    │
│  Step 3.0: Pre-flight（M0 数据校验）               │
│      ↓                                             │
│  Step 3.1: 并行执行                                │
│    ┌─────────────┐  ┌─────────────┐               │
│    │ Agent A 定性  │  │ Agent B 定量 │               │
│    │ (6维度)      │  │ (穿透回报率) │               │
│    └──────┬──────┘  └──────┬──────┘               │
│           └──────┬─────────┘                       │
│                  ↓                                  │
│  Step 3.2: Agent C（估值 + 报告组装）               │
│      ↓                                             │
│  输出：{公司名}_{代码}_分析报告.md                   │
└──────────┬──────────────────────────────────────┘
           │
           ▼
┌─────────────────────────────────────────────────┐
│           协调器交付                               │
│  确认报告文件已生成，返回给用户                      │
└─────────────────────────────────────────────────┘
```

---

## Sub-agent 调用指令

### 环境准备（首次运行）

```bash
pip install tushare pandas pdfplumber --break-system-packages
```

### Phase 0：PDF 自动获取

```
# 步骤 1：下载年报（必选，仅当用户选择"自动下载"时执行）
/download-report {stock_code} {year} 年报
# 下载目标年报 → {output_dir}/{code}_{year}_年报.pdf

# 检查下载结果
# 成功 → pdf_path = 下载文件路径
# 失败 → pdf_path = None，进入无 PDF 模式

# 步骤 2：下载中报（条件触发）
# ⚠️ 仅当 Phase 1A 输出显示中报已发布时执行（见中报时效性规则）
# 下载目标中报 → {output_dir}/{code}_{year}_中报.pdf
```

### Step A：Python 脚本（Phase 1A + Phase 2A 并行）

```
# === Phase 1A：Tushare 采集（Bash 调用）===
Bash(
  command = "python3 scripts/tushare_collector.py --code {ts_code} --output {output_dir}/data_pack_market.md",
  description = "Phase1A Tushare采集"
)
# 输出：data_pack_market.md（§1-§6, §7部分, §9, §11, §12, §14, §15, §16, §3P, §4P, 审计意见, §13.1）
# 输出：available_fields.json（可用字段清单）

# === Phase 2A.5（可选）：Agent 读取 PDF 前 10 页提取 TOC ===
# ⚠️ 仅当有 PDF 时执行，可与 Phase 1A 并行
Task(
  subagent_type = "general-purpose",
  prompt = """
  读取 PDF 文件 {output_dir}/{code}_{year}_年报.pdf 前 10 页。从目录页提取章节→页码映射，重点定位：
  - "主要控股参股公司" 或 "在子公司中的权益" 章节的起始页
  - "管理层讨论与分析" 章节的起始页
  输出 JSON: {output_dir}/toc_hints.json
  格式: {"SUB": {"page": N, "title": "..."}, "MDA": {"page": N, "title": "..."}}
  若目录页不存在或无法解析，输出空 JSON: {}
  """,
  description = "Phase2A.5 TOC定位"
)

# === Phase 2A：PDF 预处理（Bash 调用，仅当有 PDF 时，等待 Phase 2A.5）===
# 年报 PDF（必选，若有 PDF）
Bash(
  command = "python3 scripts/pdf_preprocessor.py --pdf {output_dir}/{code}_{year}_年报.pdf --output {output_dir}/pdf_sections.json --hints {output_dir}/toc_hints.json",
  description = "Phase2A PDF预处理-年报"
)
# 输出：pdf_sections.json（7 段文本片段：P2/P3/P4/P6/P13/MDA/SUB）

# 中报 PDF（条件触发，若中报 PDF 存在）
Bash(
  command = "python3 scripts/pdf_preprocessor.py --pdf {output_dir}/{code}_{h1_year}_中报.pdf --output {output_dir}/pdf_sections_interim.json",
  description = "Phase2A PDF预处理-中报"
)
```

### Step B：Agent（Phase 1B 在 Phase 1A 完成后立即启动，Phase 2B 等待 Phase 2A）

```
# === Phase 1B：Agent WebSearch 补充（Task 调用）===
# ⚠️ Phase 1A 完成后立即启动，不等待 Phase 2A
Task(
  subagent_type = "general-purpose",
  prompt = """
  请阅读 {strategy_dir}/phase1_数据采集.md 中的完整指令。

  目标股票：{stock_code}（{company_name}）
  持股渠道：{channel}

  data_pack_market.md 已由 tushare_collector.py 生成了 §1-§6, §7(部分:十大股东), §9, §11, §12, §14, §15, §16, §3P, §4P, 审计意见, §13.1 部分。
  你的任务是通过 WebSearch 补充以下章节，追加到 {output_dir}/data_pack_market.md：
  - §7 管理层与治理
  - §8 行业与竞争
  - §9B 上市子公司识别（条件触发：仅控股公司）
  - §10 MD&A 摘要
  - §13 Warnings

  §7/§8/§9B 不依赖 pdf_sections.json，直接通过 WebSearch 获取。
  §10 执行时检查 {output_dir}/pdf_sections.json 是否存在：
    若存在 → 优先使用其中的 MDA 字段
    若不存在 → 使用 WebSearch fallback 获取 MDA 摘要

  注意：data_pack_market.md 中 §8, §10, §13.2 含占位符 `*[§N 待Agent WebSearch补充]*`。
  使用 Edit 工具**替换**这些占位符为实际内容，而非在文件末尾追加。
  §7 已有结构化数据（十大股东表+审计意见），在其后追加定性信息即可。
  """,
  description = "Phase1B WebSearch补充"
)

# === Phase 2B：Agent 精提取（Task 调用，仅当有 PDF 时）===
Task(
  subagent_type = "general-purpose",
  prompt = """
  请阅读 {strategy_dir}/phase2_PDF解析.md 中的完整指令。

  pdf_sections.json 文件路径：{output_dir}/pdf_sections.json
  中报 pdf_sections（若有）：{output_dir}/pdf_sections_interim.json
  公司名称：{company_name}
  将解析结果写入：{output_dir}/data_pack_report.md
  将中报解析结果写入（若有中报）：{output_dir}/data_pack_report_interim.md
  """,
  description = "Phase2B PDF精提取"
)
```

### Phase 3：分析与报告（并行 Agent 架构）

等待 Phase 1 + Phase 2 全部完成后启动。

**条件加载规则**（协调器在启动 Agent A/B 时根据股票代码判断）：
- 港股 (.HK) → Agent A/B prompt 中额外指令：`同时加载 {shared_dir}/qualitative/references/market_rules_hk.md`
- 美股 (.US) → Agent A/B prompt 中额外指令：`同时加载 {shared_dir}/qualitative/references/market_rules_us.md`
- A股 → 无额外加载（默认路径，节省 context）

**Agent A（定性）加载**：
- `{shared_dir}/qualitative/qualitative_assessment.md` — 6维度定性分析（通用模块）
- `{shared_dir}/qualitative/references/judgment_examples.md` — 通用判断锚点（护城河、MD&A、管理层）
- `{strategy_dir}/references/factor_interface.md` — 龟龟参数传递 schema（输出末尾附校验块）

**Agent B（定量）加载**：
- `{strategy_dir}/phase3_quantitative.md` — 穿透回报率计算
- `{strategy_dir}/references/judgment_examples_turtle.md` — 龟龟专属锚点（G系数、分配意愿、λ可靠性）
- `{strategy_dir}/references/factor_interface.md` — 参数传递 schema

```
# === Step 3.0: Pre-flight（数据校验 + 口径锚定）===
Task(
  subagent_type = "general-purpose",
  prompt = """
  请阅读 {strategy_dir}/phase3_preflight.md 中的完整指令。

  数据包文件：
    - {output_dir}/data_pack_market.md
    - {output_dir}/data_pack_report.md（若存在）
    - {output_dir}/data_pack_report_interim.md（若存在）

  将 pre-flight 输出写入：{output_dir}/phase3_preflight.md
  """,
  description = "Phase3 Pre-flight"
)

# 读取 preflight 输出，检查裁决字段
# 三路分支：

# PROCEED → 直接启动 Agent A + B（正常路径）
# SUPPLEMENT_NEEDED → 解析 SUPPLEMENT_REQUEST 标记，启动 WebSearch 补充：
#   Task("根据以下补救请求通过 WebSearch 获取数据，追加到 data_pack_market.md：{gaps列表}")
#   补充完成后重新运行 preflight（最多重试 1 次，防止循环）
#   第2次 preflight 仍为 SUPPLEMENT_NEEDED → 降级为 PROCEED（标注数据局限性）
# ABORT → 通知用户数据不足原因，不启动后续 Agent，输出简要报告说明无法分析

# === Step 3.1: 并行执行 Agent A + Agent B ===
# 以下两个 Task 同时启动：

Task(
  subagent_type = "general-purpose",
  prompt = """
  请阅读 {shared_dir}/qualitative/qualitative_assessment.md 中的完整指令。
  同时加载 {shared_dir}/qualitative/references/judgment_examples.md 作为判断锚点参考。

  数据包文件：
    - {output_dir}/phase3_preflight.md（口径决策）
    - {output_dir}/data_pack_market.md
    - {output_dir}/data_pack_report.md（若存在）
    - {output_dir}/data_pack_report_interim.md（若存在）

  完成 qualitative_assessment.md 的6维度分析后，
  额外加载 {strategy_dir}/references/factor_interface.md，
  按其中 "Agent A（定性）→ Agent C" 的 schema 在输出末尾附加龟龟策略参数校验块。

  将定性分析输出写入：{output_dir}/phase3_qualitative.md
  """,
  description = "Phase3 Agent A 定性分析"
)

Task(
  subagent_type = "general-purpose",
  prompt = """
  请阅读 {strategy_dir}/phase3_quantitative.md 中的完整指令。
  同时加载 {strategy_dir}/references/judgment_examples_turtle.md 作为龟龟专属判断锚点参考。

  数据包文件：
    - {output_dir}/phase3_preflight.md（口径决策）
    - {strategy_dir}/references/shared_tables.md（税率/门槛/公式）
    - {output_dir}/data_pack_market.md
    - {output_dir}/data_pack_report.md（若存在）
    - {output_dir}/data_pack_report_interim.md（若存在）

  将定量分析输出写入：{output_dir}/phase3_quantitative.md
  """,
  description = "Phase3 Agent B 定量分析"
)

# 等待 Agent A + Agent B 全部完成

# === Step 3.2: Agent C（估值 + 报告组装）===
Task(
  subagent_type = "general-purpose",
  prompt = """
  请阅读 {strategy_dir}/phase3_valuation.md 中的完整指令。

  输入文件：
    - {output_dir}/phase3_preflight.md（基础信息）
    - {output_dir}/phase3_qualitative.md（Agent A 定性输出）
    - {output_dir}/phase3_quantitative.md（Agent B 定量输出）
    - {output_dir}/data_pack_market.md（§11 历史价格、§17 预计算值）

  将最终报告写入：{output_dir}/{company}_{code}_分析报告.md
  """,
  description = "Phase3 Agent C 估值与报告"
)
```

### 当没有 PDF 年报时（跳过 Phase 2）

```
# Phase 1 完成后直接启动 Phase 3（无 data_pack_report.md）
# Phase 3 的 Agent A/B 自动处理缺失的 data_pack_report.md
# Agent B 使用降级方案：P2/P3/P4/P6/P13 附注数据不可用
```

---

## 报表时效性规则

协调器在启动 Phase 0 前，应确定目标年报年份：

- 若当前日期在 1-3月，最新年报可能尚未发布，使用上一财年年报
- 若当前日期在 4月及以后，最新财年年报通常已发布

Tushare 数据自动覆盖最近 5 个财年，无需手动指定年份。

**支付率等关键指标必须基于同币种数据计算**（股息总额与归母净利润均取报表币种），不依赖 yfinance 的 payoutRatio 等衍生字段。

### 中报时效性规则（双PDF触发）

当 Phase 1A 的输出 data_pack_market.md 中出现 "YYYYH1" 列（如 "2025H1"），
说明该公司已发布比最新年报更新的中报（半年报）。此时：

1. Phase 0 应下载**两份** PDF：最新年报 + 最新中报
2. Phase 2A 应对两份 PDF 分别运行 pdf_preprocessor.py
3. Phase 2B 应分别处理两份 pdf_sections.json
4. Phase 3 应同时参考两份 data_pack_report

判断方法：Phase 1A 完成后，检查 data_pack_market.md 的 §3 损益表表头。
若第一列为 "YYYYH1" 格式 → 触发双 PDF 流程。

示例：
  表头为 ["2025H1", "2024", "2023", ...] → 下载 2024年报 + 2025中报
  表头为 ["2024", "2023", ...]           → 仅下载 2024年报

执行顺序调整：
```
Phase 1A + Phase 0-年报 (并行)
    ↓
检查 Phase 1A 输出是否包含 H1 列
    ↓ (若有)
Phase 0-中报 (补充下载)
    ↓
Phase 2A (处理全部 PDF)
```

---

## 阶段超时规则

| 阶段 | 最大执行时间 | 超时行为 |
|------|------------|---------|
| Phase 0 PDF下载 | 3分钟 | 标注 Warning，进入无 PDF 模式 |
| Phase 1A Tushare采集 | 2分钟 | 检查已获取的数据，部分降级继续 |
| Phase 1B WebSearch | 5分钟 | 已完成的 §N 保留，未完成的标注 "⚠️ 超时未完成" |
| Phase 2A PDF预处理 | 3分钟 | 跳过 Phase 2，进入无 PDF 模式 |
| Phase 2B PDF精提取 | 3分钟 | 已提取项保留，未提取项标注 null |
| Phase 3.0 Pre-flight | 1分钟 | 使用默认口径 |
| Phase 3.1 Agent A | 8分钟 | 已完成维度保留 |
| Phase 3.1 Agent B | 8分钟 | 已完成步骤保留 |
| Phase 3.2 Agent C | 5分钟 | 输出已有结论 |

超时后，协调器应立即推进下一阶段，不等待。总管线预计最大执行时间 ≤ 25分钟。

---

## 异常处理

| 异常情况 | 处理方式 |
|---------|---------|
| Tushare Token 无效或未配置 | 全程降级使用 yfinance MCP，标注数据源 |
| Phase 0 PDF 下载失败 | 标注 Warning，跳过 Phase 2，进入无 PDF 模式 |
| Phase 1 Step A 脚本执行失败 | 检查 Python 环境和依赖，提示安装 |
| Phase 1 Tushare 某端点返回空 | 脚本内置 yfinance fallback，标注来源 |
| Phase 1 财报数据不足5年 | 继续执行，在 data_pack 中标注实际覆盖年份 |
| Phase 2 Step A PDF 无法解析 | 跳过 Phase 2，Phase 3 使用降级方案 |
| Phase 2 关键词未命中 | 对应项返回 null，data_pack_report 标注 Warning |
| Phase 3 context 接近上限 | 中间结果已持久化到文件（preflight/qualitative/quantitative） |
| Phase 1 warnings 非空 | Phase 3 读取 warnings 区块，影响分析策略 |

---

## 文件路径约定

每个标的的运行时输出放在独立文件夹中，避免多次分析互相覆盖。

**变量定义**：
- `{workspace}` = 项目根目录
- `{shared_dir}` = `{workspace}/shared`
- `{strategy_dir}` = `{workspace}/strategies/turtle`
- `{output_dir}` = `{workspace}/output/{代码}_{公司}`（如 `output/600887_伊利股份`、`output/00001_长和`）

```
{workspace}/
├── shared/                                     ← 通用模块（只读）
│   └── qualitative/                            ← 定性分析模块
│       ├── coordinator.md                      ← 定性模块独立入口（/business-analysis）
│       ├── qualitative_assessment.md           ← 6维度定性分析
│       ├── data_collection.md                  ← 轻量级 WebSearch 指令
│       └── references/
│           ├── output_schema.md                ← 结构化参数输出 schema
│           ├── judgment_examples.md            ← 通用判断锚点
│           ├── market_rules_hk.md              ← 港股规则（条件加载）
│           └── market_rules_us.md              ← 美股规则（条件加载）
├── strategies/turtle/                          ← 龟龟策略（只读）
│   ├── coordinator.md                          ← 本文件（调度逻辑）
│   ├── phase1_数据采集.md                       ← Phase 1B prompt（WebSearch）
│   ├── phase2_PDF解析.md                        ← Phase 2B prompt（5项精提取）
│   ├── phase3_preflight.md                     ← Step 3.0 数据校验
│   ├── phase3_quantitative.md                  ← Step 3.1 Agent B 定量
│   ├── phase3_valuation.md                     ← Step 3.2 Agent C 估值+报告
│   └── references/
│       ├── shared_tables.md                    ← 税率/门槛/公式（龟龟专属）
│       ├── factor_interface.md                 ← 因子间参数传递 schema
│       └── judgment_examples_turtle.md         ← G系数/分配意愿/λ锚点（龟龟专属）
├── scripts/                                    ← 预处理脚本（只读）
│   ├── tushare_collector.py                    ← Phase 1A 数据采集脚本
│   ├── pdf_preprocessor.py                     ← Phase 2A PDF 预处理脚本
│   ├── config.py                               ← Token 管理
│   └── requirements.txt                        ← Python 依赖
└── output/                                     ← 运行时输出（按标的隔离）
    └── {code}_{company}/
        ├── data_pack_market.md                 ← Phase 1 输出
        ├── available_fields.json               ← Phase 1 输出（可用字段清单）
        ├── {code}_{year}_年报.pdf               ← Phase 0 下载（年报）
        ├── {code}_{year}_中报.pdf               ← Phase 0 下载（中报，条件触发）
        ├── pdf_sections.json                   ← Phase 2A 输出（年报）
        ├── pdf_sections_interim.json           ← Phase 2A 输出（中报，条件触发）
        ├── data_pack_report.md                 ← Phase 2B 输出（年报附注）
        ├── data_pack_report_interim.md         ← Phase 2B 输出（中报附注，条件触发）
        ├── phase3_preflight.md                 ← Step 3.0 输出
        ├── phase3_qualitative.md               ← Agent A 输出（通用定性 + 龟龟参数映射）
        ├── phase3_quantitative.md              ← Agent B 输出
        └── {company}_{code}_分析报告.md          ← 最终报告
```

**协调器职责**：在 Phase 1 启动前，创建 `{output_dir}` 目录：
```bash
mkdir -p {workspace}/output/{code}_{company}
```

---

## 数据约定

### 金额单位转换

所有阶段（Phase 1/2/3）的金额统一为 **百万元**（Tushare 原始单位元 ÷ 1e6）。

| 原始单位 | 转换方法 | 示例 |
|---------|---------|------|
| 元 | ÷ 1,000,000 | 96,886,000,000 元 → 96,886.00 百万元 |
| 千元 | ÷ 1,000 | 96,886,000 千元 → 96,886.00 百万元 |
| 万元 | ÷ 100 | 9,688,600 万元 → 96,886.00 百万元 |
| 亿元 | × 100 | 968.86 亿元 → 96,886.00 百万元 |

显示格式：使用千位逗号分隔（如 96,886.00），百分比保留2位小数。

### Phase 0 重试规则

PDF 下载最多重试 **3次**（指数退避：3s / 6s / 9s）。3次均失败：
- 在 §13 中生成 `[数据缺失|中] PDF年报下载失败，已使用3次重试`
- 进入无 PDF 模式（跳过 Phase 2，Phase 3 使用降级方案）
- 不尝试替代 URL（仅使用 `/download-report` 返回的首选 URL）

---

*龟龟投资策略 v2.0 | 协调器 | Coordinator*
