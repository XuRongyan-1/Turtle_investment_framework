# Report Aggregator Skill

## Skill Info
- **Name**: report-aggregator
- **Description**: 扫描所有已生成的分析报告，汇总为Excel表格
- **Entry Point**: `scripts/export_excel.py`
- **Slash Command**: `/report-summary`

## Dependencies
- Python 3.8+
- pandas
- openpyxl

## Required Files
- `scripts/scan_reports.py` - 报告扫描模块
- `scripts/export_excel.py` - Excel导出模块

## Output
- `output/reports_summary_YYYYMMDD_HHMMSS.xlsx` - Excel汇总文件

## Usage
```
/report-summary
```

## Features
- 扫描output目录下所有分析报告
- 提取四因子评分（盈利能力、资本配置、安全边际、护城河）
- 汇总价格区间建议（增持/配置/持有/减仓）
- 计算历史分位和折价/溢价率
- 评估风险等级
- 生成带颜色格式的Excel表格

## Excel Columns
1. 股票代码
2. 公司名称
3. 报告日期
4. 因子1-盈利能力
5. 因子2-资本配置
6. 因子3-安全边际
7. 因子4-护城河
8. 综合评分
9. 当前价格
10. 币种
11. 历史分位(%)
12. 中位数价格
13. 相对中位数溢价率(%)
14. 相对中位数折价率(%)
15. 增持上限
16. 配置区间下限
17. 配置区间上限
18. 持有区间下限
19. 持有区间上限
20. 减仓下限
21. 风险等级
22. 源文件
