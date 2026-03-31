# 定性分析模块 — 独立协调器

> **角色**：你是项目经理。职责：(1) 验证输入；(2) 调度数据采集；(3) 启动6维度定性分析；(4) 交付完整报告。
>
> 本协调器用于**独立运行**定性分析模块（`/business-analysis` 命令）。
> 当被龟龟、烟蒂等策略调用时，策略自身的协调器会直接加载 `qualitative_assessment.md`，不经过本文件。

---

## 输入解析

用户输入可能包含以下组合：

| 输入项 | 示例 | 必需？ |
|--------|------|--------|
| 股票代码或名称 | `600887` / `伊利股份` / `0001.HK` / `AAPL` | 必需 |

**解析规则**：
1. 从用户消息中提取股票代码/名称
2. 若用户只给了公司名称没给代码，在数据采集阶段由脚本通过 Tushare `stock_basic` 确认代码
3. 代码格式化：A股 → `XXXXXX.SH/SZ`；港股 → `XXXXX.HK`；美股 → `AAPL.US`

**AskUserQuestion 触发**：
- 多地上市 → "{公司}分析哪个市场？"
- 模糊公司名 → "确认您要分析的公司"
- TUSHARE_TOKEN 未设置 → "请提供 Tushare Token"（无 Token → 降级 yfinance）

---

## 执行流程

```
┌─────────────────────────────────────────────────┐
│              用户输入解析                          │
│   股票代码 = {code}                               │
│   Tushare Token = {有 | 无 → yfinance fallback}  │
└──────────┬──────────────────────────────────────┘
           │
           ▼
┌─────────────────────────────────────────────────┐
│  Step 1：Tushare 数据采集                         │
│                                                   │
│  Bash 运行：                                      │
│  python3 scripts/tushare_collector.py             │
│    --code {ts_code}                               │
│    --output {output_dir}/data_pack_market.md      │
│                                                   │
│  输出：data_pack_market.md                        │
│    (§1-§6, §7部分, §9, §11, §12, §14-§16,       │
│     §3P, §4P, 审计意见, §13.1)                   │
└──────────┬──────────────────────────────────────┘
           │
           ▼
┌─────────────────────────────────────────────────┐
│  Step 2：WebSearch 补充                           │
│                                                   │
│  Agent 读取 data_collection.md 指令               │
│  补充 §7(定性), §8(行业), §10(MD&A)              │
│  → 写入 data_pack_market.md                       │
└──────────┬──────────────────────────────────────┘
           │
           ▼
┌─────────────────────────────────────────────────┐
│  Step 3：6维度定性分析                             │
│                                                   │
│  Agent 读取 qualitative_assessment.md 指令        │
│  输入：data_pack_market.md                        │
│  → 输出：qualitative_report.md                    │
│    （完整报告 + 结构化参数表）                      │
└──────────┬──────────────────────────────────────┘
           │
           ▼
┌─────────────────────────────────────────────────┐
│           交付                                    │
│  确认报告文件已生成，返回给用户                      │
└─────────────────────────────────────────────────┘
```

---

## Sub-agent 调用指令

### 环境准备（首次运行）

```bash
pip install tushare pandas --break-system-packages
```

### Step 1：Tushare 数据采集

```
Bash(
  command = "python3 scripts/tushare_collector.py --code {ts_code} --output {output_dir}/data_pack_market.md",
  description = "Tushare数据采集"
)
```

### Step 2：WebSearch 数据补充

```
Task(
  subagent_type = "general-purpose",
  prompt = """
  请阅读 {shared_dir}/qualitative/data_collection.md 中的完整指令。

  目标股票：{stock_code}（{company_name}）
  data_pack_market.md 路径：{output_dir}/data_pack_market.md

  data_pack_market.md 已由 tushare_collector.py 生成了结构化数据。
  你的任务是通过 WebSearch 补充以下章节：
  - §7 管理层与治理（定性信息）
  - §8 行业与竞争
  - §10 MD&A 摘要

  注意：data_pack_market.md 中含占位符 `*[§N 待Agent WebSearch补充]*`。
  使用 Edit 工具**替换**这些占位符为实际内容。
  §7 已有结构化数据（十大股东表+审计意见），在其后追加定性信息即可。
  """,
  description = "WebSearch数据补充"
)
```

### Step 3：6维度定性分析

```
Task(
  subagent_type = "general-purpose",
  prompt = """
  请阅读 {shared_dir}/qualitative/qualitative_assessment.md 中的完整指令。

  数据包文件：
    - {output_dir}/data_pack_market.md

  注意：本次为独立运行模式，不存在 phase3_preflight.md。
  使用 GAAP 归母净利润作为默认利润口径。

  将定性分析输出写入：{output_dir}/qualitative_report.md
  """,
  description = "6维度定性分析"
)
```

### Step 4：生成 HTML 仪表盘报告

```
# 将 MD 报告转换为带样式的 HTML 仪表盘
Bash(
  command = "python3 scripts/report_to_html.py --input {output_dir}/qualitative_report.md --output {output_dir}/qualitative_report.html",
  description = "生成HTML仪表盘报告"
)
# 输出：qualitative_report.html（信息仪表盘风格，可直接浏览器打开或打印为PDF）
```

---

## 条件加载规则

协调器在启动 Step 3 Agent 时根据股票代码判断：
- 港股 (.HK) → 额外指令：`同时加载 {shared_dir}/qualitative/references/market_rules_hk.md`
- 美股 (.US) → 额外指令：`同时加载 {shared_dir}/qualitative/references/market_rules_us.md`
- A股 → 无额外加载

所有 Agent 加载：
- `{shared_dir}/qualitative/references/judgment_examples.md` — 判断锚点

---

## 文件路径约定

**变量定义**：
- `{workspace}` = 项目根目录
- `{shared_dir}` = `{workspace}/shared`
- `{output_dir}` = `{workspace}/output/{代码}_{公司}`（如 `output/600887_伊利股份`）

```
{workspace}/
├── shared/qualitative/                        ← 本模块（只读）
│   ├── coordinator.md                         ← 本文件
│   ├── qualitative_assessment.md              ← 6维度分析
│   ├── data_collection.md                     ← WebSearch 指令
│   ├── references/
│   │   ├── output_schema.md                   ← 参数 schema
│   │   ├── judgment_examples.md               ← 判断锚点
│   │   ├── framework_guide.md                 ← 框架说明（固定附录）
│   │   ├── market_rules_hk.md                 ← 港股规则（条件加载）
│   │   └── market_rules_us.md                 ← 美股规则（条件加载）
│   └── templates/
│       └── dashboard.html                     ← HTML 仪表盘模板
├── scripts/                                   ← 共享脚本
│   ├── tushare_collector.py                   ← 数据采集
│   ├── report_to_html.py                      ← MD→HTML 转换
│   └── config.py                              ← Token 管理
└── output/{code}_{company}/                   ← 运行时输出
    ├── data_pack_market.md                    ← Step 1+2 输出
    ├── qualitative_report.md                  ← Step 3 输出（MD报告）
    └── qualitative_report.html                ← Step 4 输出（HTML仪表盘）
```

**协调器职责**：在 Step 1 启动前，创建 `{output_dir}` 目录：
```bash
mkdir -p {workspace}/output/{code}_{company}
```

---

## 阶段超时规则

| 阶段 | 最大执行时间 | 超时行为 |
|------|------------|---------|
| Step 1 Tushare采集 | 2分钟 | 检查已获取的数据，部分降级继续 |
| Step 2 WebSearch | 5分钟 | 已完成的章节保留，未完成的标注 "⚠️ 超时未完成" |
| Step 3 定性分析 | 8分钟 | 已完成维度保留 |

---

## 异常处理

| 异常情况 | 处理方式 |
|---------|---------|
| Tushare Token 无效或未配置 | 全程降级使用 yfinance，标注数据源 |
| Step 1 脚本执行失败 | 检查 Python 环境和依赖，提示安装 |
| Step 1 Tushare 某端点返回空 | 脚本内置 yfinance fallback，标注来源 |
| Step 2 WebSearch 搜索无结果 | 对应项标注 "⚠️ 未搜索到"，Step 3 降级标注 |

---

## 数据约定

### 金额单位转换

所有金额统一为 **百万元**（Tushare 原始单位元 ÷ 1e6）。

| 原始单位 | 转换方法 | 示例 |
|---------|---------|------|
| 元 | ÷ 1,000,000 | 96,886,000,000 元 → 96,886.00 百万元 |
| 千元 | ÷ 1,000 | 96,886,000 千元 → 96,886.00 百万元 |
| 万元 | ÷ 100 | 9,688,600 万元 → 96,886.00 百万元 |
| 亿元 | × 100 | 968.86 亿元 → 96,886.00 百万元 |

显示格式：使用千位逗号分隔（如 96,886.00），百分比保留2位小数。

---

*通用定性分析模块 v1.0 | 独立协调器*
