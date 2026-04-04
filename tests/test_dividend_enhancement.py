"""Tests for dividend enhancement features.

Feature List Tests:
- #4: Verify report period detection from end_date
- #5: Verify dividend table output format with report period column
- #6: Verify payout ratio calculation uses annual sum
"""

import pandas as pd


def test_get_report_period():
    """Test #4: Verify report period detection from end_date."""
    from scripts.tushare_modules.financials import FinancialsMixin

    # Test various end_date formats
    test_cases = [
        ("20241231", "年报"),
        ("20240630", "中报"),
        ("20240331", "一季报"),
        ("20240930", "三季报"),
        ("20231231", "年报"),
        ("20230630", "中报"),
        ("invalid", "其他"),
        ("", "其他"),
        (None, "其他"),
    ]

    for end_date, expected in test_cases:
        result = FinancialsMixin._get_report_period(end_date)
        assert result == expected, f"Failed for end_date={end_date}: expected {expected}, got {result}"


def test_dividend_table_format():
    """Test #5: Verify dividend table includes report period column."""
    from scripts.tushare_modules.financials import FinancialsMixin
    from scripts.tushare_modules.infrastructure import InfrastructureMixin
    from unittest.mock import MagicMock, patch

    # Create a mock client
    class MockClient(FinancialsMixin, InfrastructureMixin):
        def __init__(self):
            self._store = {}

        def _is_hk(self, ts_code):
            return False

        def _is_us(self, ts_code):
            return False

        def _safe_call(self, api_name, **kwargs):
            # Mock dividend data with multiple dividends in same year
            data = {
                "ts_code": ["000001.SZ", "000001.SZ", "000001.SZ", "000001.SZ"],
                "end_date": ["20241231", "20240630", "20231231", "20221231"],
                "ann_date": ["20250301", "20240801", "20240301", "20230301"],
                "div_proc": ["实施", "实施", "实施", "实施"],
                "stk_div": [0.0, 0.0, 0.0, 0.0],
                "cash_div_tax": [1.5, 0.8, 1.2, 1.0],
                "record_date": ["20250501", "20240901", "20240501", "20230501"],
                "ex_date": ["20250502", "20240902", "20240502", "20230502"],
                "base_share": [100.0, 100.0, 100.0, 100.0],
            }
            return pd.DataFrame(data)

    client = MockClient()
    result = client.get_dividends("000001.SZ")

    # Verify output contains report period column
    assert "报告期" in result, "Table should contain '报告期' column"
    assert "年报" in result, "Table should contain '年报'"
    assert "中报" in result, "Table should contain '中报'"

    # Verify the table structure
    lines = result.split("\n")
    header_line = [l for l in lines if "年度" in l and "报告期" in l][0]
    assert "| 年度 | 报告期 |" in header_line, f"Header format incorrect: {header_line}"


def test_payout_ratio_annual_sum():
    """Test #6: Verify payout ratio calculation uses annual sum of dividends."""
    from scripts.tushare_modules.infrastructure import InfrastructureMixin
    from unittest.mock import MagicMock

    # Create a mock client
    class MockClient(InfrastructureMixin):
        def __init__(self):
            self._store = {}

        def _get_annual_df(self, name):
            # Mock income data
            if name == "income":
                data = {
                    "end_date": ["20241231", "20231231", "20221231"],
                    "n_income_attr_p": [1000.0, 900.0, 800.0],  # 归母净利润（百万元）
                }
                return pd.DataFrame(data)
            return pd.DataFrame()

        def _safe_float(self, val):
            if pd.isna(val) or val is None:
                return None
            return float(val)

    client = MockClient()

    # Mock dividend data: 2024 has two dividends (年报1.5元 + 中报0.8元)
    dividend_data = {
        "end_date": ["20241231", "20240630", "20231231"],
        "cash_div_tax": [1.5, 0.8, 1.2],  # 每股现金分红
        "base_share": [100.0, 100.0, 100.0],  # 总股本（万股）
    }
    client._store["dividends"] = pd.DataFrame(dividend_data)

    result = client._get_payout_by_year()

    # Expected calculation:
    # 2024: (1.5 + 0.8) * 100万股 * 10000 / 1000百万净利润 * 100 = 230%
    # 2023: 1.2 * 100 * 10000 / 900 * 100 = 133.33%

    assert "2024" in result, "Should have payout ratio for 2024"
    assert "2023" in result, "Should have payout ratio for 2023"

    # 2024 payout should be based on sum of both dividends (2.3 * 1000000 / 1000 * 100 = 230)
    expected_2024 = (1.5 + 0.8) * 100 * 10000 / 1000 * 100
    assert abs(result["2024"] - expected_2024) < 0.1, f"2024 payout ratio incorrect: expected {expected_2024}, got {result['2024']}"

    # 2023 payout should be based on single dividend
    expected_2023 = 1.2 * 100 * 10000 / 900 * 100
    assert abs(result["2023"] - expected_2023) < 0.1, f"2023 payout ratio incorrect: expected {expected_2023}, got {result['2023']}"


if __name__ == "__main__":
    print("Running test #4: report period detection...")
    test_get_report_period()
    print("PASSED: Report period detection")

    print("\nRunning test #5: dividend table format...")
    test_dividend_table_format()
    print("PASSED: Dividend table format")

    print("\nRunning test #6: payout ratio annual sum...")
    test_payout_ratio_annual_sum()
    print("PASSED: Payout ratio annual sum")

    print("\n=== All tests passed! ===")
