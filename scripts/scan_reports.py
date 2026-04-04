"""
报告扫描器模块 - 用于扫描output目录中的所有分析报告并提取关键数据
"""

import os
import re
from pathlib import Path
from typing import Dict, List, Optional, NamedTuple
from dataclasses import dataclass


@dataclass
class ReportData:
    """报告数据结构"""
    stock_code: str
    company_name: str
    report_date: str

    # 四因子评分
    factor1_profitability: float  # 因子1: 盈利能力
    factor2_capital_alloc: float  # 因子2: 资本配置
    factor3_safety_margin: float  # 因子3: 安全边际
    factor4_moat: float           # 因子4: 护城河
    overall_score: float          # 综合评分

    # 价格数据
    current_price: float
    price_currency: str

    # 价格区间
    buy_threshold: float          # 积极增持价格上限
    accumulate_low: float         # 适度配置区间下限
    accumulate_high: float        # 适度配置区间上限
    hold_low: float               # 持有区间下限
    hold_high: float              # 持有区间上限
    sell_threshold: float         # 考虑减仓价格下限

    # 估值数据
    hist_percentile: float        # 历史分位
    median_price: float           # 中位数价格

    # 风险评级
    risk_level: str               # 风险等级 (低/中/高)

    # 溢价/折价率
    discount_premium_ratio: float  # 相对中位数折价/溢价率

    # 源文件路径
    source_file: str


class ReportScanner:
    """报告扫描器"""

    def __init__(self, output_dir: str = "output"):
        self.output_dir = Path(output_dir)
        self.reports: List[ReportData] = []

    def scan_all(self) -> List[ReportData]:
        """扫描所有报告"""
        self.reports = []

        if not self.output_dir.exists():
            print(f"输出目录不存在: {self.output_dir}")
            return self.reports

        # 遍历所有子目录
        for subdir in self.output_dir.iterdir():
            if subdir.is_dir() and not subdir.name.startswith('.'):
                report = self._parse_report_directory(subdir)
                if report:
                    self.reports.append(report)

        # 按股票代码排序
        self.reports.sort(key=lambda x: x.stock_code)
        return self.reports

    def _parse_report_directory(self, subdir: Path) -> Optional[ReportData]:
        """解析单个报告目录"""
        # 查找分析报告文件
        report_files = list(subdir.glob("*_分析报告.md"))
        if not report_files:
            return None

        report_file = report_files[0]
        try:
            content = report_file.read_text(encoding='utf-8')
            return self._extract_data(content, str(report_file))
        except Exception as e:
            print(f"解析文件失败 {report_file}: {e}")
            return None

    def _extract_data(self, content: str, source_file: str) -> Optional[ReportData]:
        """从报告内容提取数据"""
        try:
            # 提取基本信息 - 尝试多种格式
            stock_code = self._extract_pattern(content, r'\*\*股票代码\*\*\s*\|\s*([^|\n]+)')
            if not stock_code:
                stock_code = self._extract_pattern(content, r'\*\*股票代码\*\*[:：]\s*([^\n]+)')
            if not stock_code:
                stock_code = self._extract_pattern(content, r'股票代码[:：]\s*([^\n]+)')

            company_name = self._extract_pattern(content, r'\*\*公司全称\*\*\s*\|\s*([^|\n]+)')
            if not company_name:
                company_name = self._extract_pattern(content, r'\*\*公司名称\*\*[:：]\s*([^\n]+)')
            if not company_name:
                # 尝试从标题提取
                company_name = self._extract_pattern(content, r'#\s*([^（(]+)[（(]')
            if not company_name:
                # 尝试从文件名提取
                company_name = self._extract_pattern(content, r'([^\n]+)龟龟投资策略')

            report_date = self._extract_pattern(content, r'\*报告日期：([^*\n]+)\*')
            if not report_date:
                report_date = self._extract_pattern(content, r'\*\*分析日期\*\*[:：]\s*([^\n]+)')

            # 提取四因子评分 - 支持多种命名
            f1 = self._extract_score(content, '盈利能力', '因子1', '盈利')
            f2 = self._extract_score(content, '资本配置', '因子2', '利润归属', '资本')
            f3 = self._extract_score(content, '安全边际', '因子3', '现金创造', '估值定价', '安全')
            f4 = self._extract_score(content, '护城河', '因子4', '竞争')
            overall = self._extract_overall_score(content)

            # 如果找不到综合评分，计算四因子平均值
            if overall == 0.0 and any([f1, f2, f3, f4]):
                valid_scores = [s for s in [f1, f2, f3, f4] if s > 0]
                if valid_scores:
                    overall = round(sum(valid_scores) / len(valid_scores), 1)

            # 提取当前价格
            current_price, currency = self._extract_current_price(content)

            # 提取价格区间
            price_zones = self._extract_price_zones(content)

            # 提取历史分位和中位数
            hist_percentile = self._extract_hist_percentile(content)
            median_price = self._extract_median_price(content)

            # 计算折价/溢价率
            discount_premium = 0.0
            if current_price > 0 and median_price > 0:
                discount_premium = ((current_price - median_price) / median_price) * 100

            # 提取风险等级
            risk_level = self._extract_risk_level(content)

            return ReportData(
                stock_code=stock_code or 'N/A',
                company_name=company_name or 'N/A',
                report_date=report_date or 'N/A',
                factor1_profitability=f1,
                factor2_capital_alloc=f2,
                factor3_safety_margin=f3,
                factor4_moat=f4,
                overall_score=overall,
                current_price=current_price,
                price_currency=currency,
                buy_threshold=price_zones.get('buy', 0),
                accumulate_low=price_zones.get('accumulate_low', 0),
                accumulate_high=price_zones.get('accumulate_high', 0),
                hold_low=price_zones.get('hold_low', 0),
                hold_high=price_zones.get('hold_high', 0),
                sell_threshold=price_zones.get('sell', 0),
                hist_percentile=hist_percentile,
                median_price=median_price,
                risk_level=risk_level,
                discount_premium_ratio=discount_premium,
                source_file=source_file
            )
        except Exception as e:
            print(f"提取数据失败: {e}")
            return None

    def _extract_pattern(self, content: str, pattern: str) -> Optional[str]:
        """通用正则提取"""
        match = re.search(pattern, content)
        if match:
            return match.group(1).strip()
        return None

    def _extract_score(self, content: str, *keywords) -> float:
        """提取评分（支持多个关键词）"""
        for keyword in keywords:
            # 匹配格式: **评分：7.5/10** 或 评分：**7.5/10**
            # 也支持: | 因子 | 8.5/10 | 说明 |
            patterns = [
                rf'{keyword}.*[：:]\s*\*\*([\d.]+)/10\*\*',
                rf'{keyword}.*评分[：:]\s*\*\*?([\d.]+)\*\*?',
                rf'{keyword}.*\*\*([\d.]+)/10\*\*',
                rf'\|\s*{keyword}[^|]*\|\s*([\d.]+)/10\s*\|',
                rf'\|\s*{keyword}[^|]*[：:]\s*\|\s*([\d.]+)/10\s*\|',
                rf'因子[12]?[：:]?\s*[^|]*{keyword}[^|]*\|\s*([\d.]+)/10'
            ]
            for pattern in patterns:
                match = re.search(pattern, content, re.IGNORECASE)
                if match:
                    try:
                        return float(match.group(1))
                    except ValueError:
                        continue
        return 0.0

    def _extract_overall_score(self, content: str) -> float:
        """提取综合评分"""
        patterns = [
            r'综合评分[：:]\s*\*\*?([\d.]+)/10\*\*?',
            r'\*\*([\d.]+)/10\*\*.*综合',
            r'评分汇总.*\n.*\*\*([\d.]+)/10\*\*',
            r'\|\s*综合评分\s*\|\s*\*\*([\d.]+)/10\*\*',
            r'\|\s*\*\*综合评分\*\*\s*\|\s*([\d.]+)/10'
        ]
        for pattern in patterns:
            match = re.search(pattern, content)
            if match:
                try:
                    return float(match.group(1))
                except ValueError:
                    continue
        # 如果找不到综合评分，计算四因子平均值
        return 0.0

    def _extract_current_price(self, content: str) -> tuple:
        """提取当前价格和币种"""
        # 匹配当前股价 - 支持多种格式
        patterns = [
            r'\*\*当前股价[（(].*[)）]\*\*[：:]\s*([\d.]+)\s*([A-Za-z]*)',
            r'当前股价.*?([\d.]+)\s*([A-Za-z]*)',
            r'最新收盘价.*?([\d.]+)\s*([A-Za-z]*)',
            r'当前股价[:：]\s*([\d.]+)\s*([A-Za-z]*)',
            r'当前价格.*?(\d+\.?\d*)\s*元',
            r'\*\*当前价格\*\*[:：]\s*([\d.]+)'
        ]
        for pattern in patterns:
            match = re.search(pattern, content)
            if match:
                try:
                    price = float(match.group(1))
                    currency = match.group(2) if len(match.groups()) > 1 and match.group(2) else 'CNY'
                    if not currency:
                        currency = 'CNY'
                    return price, currency
                except ValueError:
                    continue
        return 0.0, 'CNY'

    def _extract_price_zones(self, content: str) -> Dict[str, float]:
        """提取操作建议价格区间"""
        zones = {}

        # 查找操作建议表格
        # 格式: | < 6.0港元 | **积极增持** |
        buy_pattern = r'\|\s*<\s*([\d.]+).*?\|\s*\*\*积极增持\*\*'
        match = re.search(buy_pattern, content)
        if match:
            zones['buy'] = float(match.group(1))

        # 适度配置区间
        acc_pattern = r'([\d.]+)\s*-\s*([\d.]+).*?\|\s*\*\*适度配置\*\*'
        match = re.search(acc_pattern, content)
        if match:
            zones['accumulate_low'] = float(match.group(1))
            zones['accumulate_high'] = float(match.group(2))

        # 持有区间
        hold_pattern = r'([\d.]+)\s*-\s*([\d.]+).*?\|\s*\*\*持有观望\*\*'
        match = re.search(hold_pattern, content)
        if match:
            zones['hold_low'] = float(match.group(1))
            zones['hold_high'] = float(match.group(2))

        # 减仓阈值
        sell_pattern = r'>\s*([\d.]+).*?\|\s*\*\*考虑减仓\*\*'
        match = re.search(sell_pattern, content)
        if match:
            zones['sell'] = float(match.group(1))

        return zones

    def _extract_hist_percentile(self, content: str) -> float:
        """提取历史分位"""
        patterns = [
            r'历史分位[：:]\s*\*\*?([\d.]+)%\*\*?',
            r'([\d.]+)%\s*历史分位',
            r'当前股价.*?([\d.]+)%\s*历史分位'
        ]
        for pattern in patterns:
            match = re.search(pattern, content)
            if match:
                try:
                    return float(match.group(1))
                except ValueError:
                    continue
        return 0.0

    def _extract_median_price(self, content: str) -> float:
        """提取中位数价格"""
        patterns = [
            r'50%分位.*?(?:中位数)?.*?([\d.]+)\s*[A-Za-z]*',
            r'中位数.*?([\d.]+)\s*[A-Za-z]*'
        ]
        for pattern in patterns:
            match = re.search(pattern, content)
            if match:
                try:
                    return float(match.group(1))
                except ValueError:
                    continue
        return 0.0

    def _extract_risk_level(self, content: str) -> str:
        """提取风险等级"""
        # 查找风险分析部分
        risk_section = re.search(r'## 风险分析([\s\S]*?)(?:##|$)', content)
        if risk_section:
            section = risk_section.group(1)
            # 统计各等级风险数量
            high_risks = len(re.findall(r'\*\*高\*\*', section))
            medium_risks = len(re.findall(r'\*\*中\*\*', section))
            low_risks = len(re.findall(r'\*\*低\*\*', section))

            if high_risks > 0:
                return '高'
            elif medium_risks > 0:
                return '中'
            else:
                return '低'
        return '中'  # 默认中等风险

    def get_summary(self) -> Dict:
        """获取汇总统计"""
        if not self.reports:
            return {}

        scores = [r.overall_score for r in self.reports]
        return {
            'total_count': len(self.reports),
            'avg_score': sum(scores) / len(scores) if scores else 0,
            'max_score': max(scores) if scores else 0,
            'min_score': min(scores) if scores else 0,
            'high_score_count': len([s for s in scores if s >= 8]),
            'low_score_count': len([s for s in scores if s < 6])
        }


if __name__ == "__main__":
    # 测试扫描功能
    scanner = ReportScanner()
    reports = scanner.scan_all()

    print(f"扫描到 {len(reports)} 份报告")
    print("\n" + "="*80)

    for r in reports:
        print(f"\n股票: {r.stock_code} - {r.company_name}")
        print(f"  综合评分: {r.overall_score}/10")
        print(f"  四因子: {r.factor1_profitability}, {r.factor2_capital_alloc}, {r.factor3_safety_margin}, {r.factor4_moat}")
        print(f"  当前价格: {r.current_price} {r.price_currency}")
        print(f"  历史分位: {r.hist_percentile}%")
        print(f"  折价/溢价率: {r.discount_premium_ratio:.1f}%")
        print(f"  风险等级: {r.risk_level}")

    summary = scanner.get_summary()
    print("\n" + "="*80)
    print("汇总统计:")
    print(f"  报告总数: {summary.get('total_count', 0)}")
    print(f"  平均评分: {summary.get('avg_score', 0):.2f}")
    print(f"  最高评分: {summary.get('max_score', 0):.2f}")
    print(f"  最低评分: {summary.get('min_score', 0):.2f}")
