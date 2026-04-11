# Turtle Investment Framework — Skill Definition

## Skill Info
- **Name**: turtle-analysis
- **Description**: Run a full 龟龟投资策略 (Turtle Investment Framework) multi-phase fundamental analysis on Chinese A-share or HK stocks
- **Entry Point**: `prompts/coordinator.md`
- **Slash Command**: `/turtle-analysis <stock_code>`

## Dependencies
- **download-report**: Built-in `/download-annual-report` slash command for Phase 0 PDF acquisition
- **Python venv**: `.venv/` with `tushare`, `pandas`, `pdfplumber` (created by `bash init.sh`)
- **Tushare Pro API**: Requires `TUSHARE_TOKEN` environment variable

## Required Environment Variables
| Variable | Description | Required |
|----------|-------------|----------|
| `TUSHARE_TOKEN` | Tushare Pro API token | Yes |

## Pipeline Phases
1. **Phase 0**: PDF acquisition (`/download-annual-report` slash command)
2. **Phase 1A**: Tushare data collection (`scripts/tushare_collector.py`)
3. **Phase 1B**: Agent WebSearch for qualitative data
4. **Phase 2A**: PDF preprocessing (`scripts/pdf_preprocessor.py`)
5. **Phase 2B**: Agent PDF structured extraction
6. **Phase 3**: 4-factor analysis and report generation

## Output
- `output/{code}_{company}/` — all intermediate and final files
- Final report: `{company}_{code}_分析报告.md`
