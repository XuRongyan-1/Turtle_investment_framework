"""Tests for Turtle Screener (龟龟选股器).

Tests cover:
- ScreenerConfig defaults, overrides, validation
- Tier 1 bulk data, filtering, ranking
- Tier 2 hard vetoes, financial quality, Factor 2/4, floor price
- Composite scoring and full pipeline
"""

import json
import math
import os
from unittest.mock import MagicMock, patch

import pandas as pd
import pytest

from screener_config import ScreenerConfig
from screener_core import TushareScreener

MOCK_DIR = os.path.join(os.path.dirname(__file__), "fixtures", "mock_tushare_responses")


def _load_bulk_mock(filename: str) -> pd.DataFrame:
    """Load a bulk mock fixture as DataFrame."""
    with open(os.path.join(MOCK_DIR, filename)) as f:
        data = json.load(f)
    return pd.DataFrame(data)


# ============================================================
# Feature #96: ScreenerConfig
# ============================================================


class TestScreenerConfig:
    """Tests for ScreenerConfig dataclass."""

    def test_default_values(self):
        cfg = ScreenerConfig()
        assert cfg.min_listing_years == 3
        assert cfg.min_market_cap_yi == 5.0
        assert cfg.min_turnover_pct == 0.1
        assert cfg.max_pb == 10.0
        assert cfg.max_pe == 50.0
        assert cfg.obs_channel_limit == 50
        assert cfg.tier2_main_limit == 150
        assert cfg.min_roe == 8.0
        assert cfg.min_gross_margin == 15.0
        assert cfg.max_debt_ratio == 70.0
        assert cfg.max_pledge_pct == 70.0

    def test_scoring_weights_sum_to_one(self):
        cfg = ScreenerConfig()
        total = sum(cfg.scoring_weights.values())
        assert abs(total - 1.0) < 0.001

    def test_tier2_max_stocks(self):
        cfg = ScreenerConfig()
        assert cfg.tier2_max_stocks == 200  # 150 + 50

    def test_override_from_dict(self):
        overrides = {"min_roe": 15.0, "max_pe": 30.0, "tier2_main_limit": 100}
        cfg = ScreenerConfig.from_dict(overrides)
        assert cfg.min_roe == 15.0
        assert cfg.max_pe == 30.0
        assert cfg.tier2_main_limit == 100
        # Defaults still intact
        assert cfg.min_listing_years == 3

    def test_from_dict_ignores_unknown_keys(self):
        overrides = {"min_roe": 10.0, "unknown_key": "hello"}
        cfg = ScreenerConfig.from_dict(overrides)
        assert cfg.min_roe == 10.0

    def test_to_dict(self):
        cfg = ScreenerConfig(min_roe=12.0)
        d = cfg.to_dict()
        assert d["min_roe"] == 12.0
        assert "min_listing_years" in d

    def test_validate_ok(self):
        cfg = ScreenerConfig()
        errors = cfg.validate()
        assert errors == []

    def test_validate_bad_weights(self):
        cfg = ScreenerConfig(weight_roe=0.5, weight_fcf_yield=0.5,
                             weight_penetration_r=0.5, weight_ev_ebitda=0.0,
                             weight_floor_premium=0.0)
        errors = cfg.validate()
        assert any("weights" in e.lower() for e in errors)

    def test_validate_negative_listing_years(self):
        cfg = ScreenerConfig(min_listing_years=-1)
        errors = cfg.validate()
        assert any("listing" in e.lower() for e in errors)

    def test_scoring_weights_keys(self):
        cfg = ScreenerConfig()
        keys = set(cfg.scoring_weights.keys())
        assert keys == {"roe", "fcf_yield", "penetration_r", "ev_ebitda", "floor_premium"}

    def test_dv_pe_pb_weights(self):
        cfg = ScreenerConfig()
        assert cfg.dv_weight == 0.4
        assert cfg.pe_weight == 0.3
        assert cfg.pb_weight == 0.3

    def test_cache_defaults(self):
        cfg = ScreenerConfig()
        assert cfg.cache_stock_basic_ttl_days == 7
        assert cfg.cache_daily_basic_ttl_days == 0
        assert cfg.cache_dir == "output/.screener_cache"


class TestBulkMockFixtures:
    """Verify mock fixture files exist and have correct shape."""

    def test_stock_basic_bulk_exists(self):
        df = _load_bulk_mock("stock_basic_bulk.json")
        assert len(df) == 10
        assert "ts_code" in df.columns
        assert "name" in df.columns
        assert "list_date" in df.columns

    def test_daily_basic_bulk_exists(self):
        df = _load_bulk_mock("daily_basic_bulk.json")
        assert len(df) == 10
        assert "pe_ttm" in df.columns
        assert "pb" in df.columns
        assert "total_mv" in df.columns
        assert "dv_ttm" in df.columns
        assert "turnover_rate" in df.columns

    def test_stock_basic_has_st_stock(self):
        df = _load_bulk_mock("stock_basic_bulk.json")
        st_stocks = df[df["name"].str.contains(r"\*ST|ST", na=False)]
        assert len(st_stocks) >= 1

    def test_stock_basic_has_new_ipo(self):
        df = _load_bulk_mock("stock_basic_bulk.json")
        # 301234 listed 2025-01-01, less than 3 years from 2026
        new = df[df["list_date"] == "20250101"]
        assert len(new) == 1

    def test_daily_basic_has_negative_pe(self):
        df = _load_bulk_mock("daily_basic_bulk.json")
        neg_pe = df[df["pe_ttm"] < 0]
        assert len(neg_pe) >= 1

    def test_daily_basic_has_zero_dividend(self):
        df = _load_bulk_mock("daily_basic_bulk.json")
        zero_div = df[df["dv_ttm"] == 0]
        assert len(zero_div) >= 1

    def test_daily_basic_has_negative_pb(self):
        df = _load_bulk_mock("daily_basic_bulk.json")
        neg_pb = df[df["pb"] < 0]
        assert len(neg_pb) >= 1


# ============================================================
# Feature #97: TushareScreener + Cache
# ============================================================


def _make_screener(tmp_path=None, config=None):
    """Create a TushareScreener with mocked tushare."""
    cfg = config or ScreenerConfig()
    if tmp_path:
        cfg = ScreenerConfig(**{**cfg.to_dict(), "cache_dir": str(tmp_path / "cache")})
    with patch("screener_core.get_token", return_value="test_token"):
        screener = TushareScreener(token="test_token", config=cfg)
    return screener


def _merged_mock_df() -> pd.DataFrame:
    """Load and merge stock_basic + daily_basic mock data."""
    sb = _load_bulk_mock("stock_basic_bulk.json")
    db = _load_bulk_mock("daily_basic_bulk.json")
    return sb.merge(db, on="ts_code", how="inner")


class TestScreenerCache:
    """Tests for ScreenerCache."""

    def test_put_and_get(self, tmp_path):
        from screener_core import ScreenerCache
        cache = ScreenerCache(str(tmp_path / "cache"))
        df = pd.DataFrame({"a": [1, 2, 3]})
        cache.put("test_key", df)
        result = cache.get("test_key", ttl_seconds=3600)
        assert result is not None
        assert len(result) == 3

    def test_get_expired(self, tmp_path):
        from screener_core import ScreenerCache
        cache = ScreenerCache(str(tmp_path / "cache"))
        df = pd.DataFrame({"a": [1]})
        cache.put("test_key", df)
        # Force expiry by using TTL=0
        result = cache.get("test_key", ttl_seconds=0)
        assert result is None

    def test_get_missing(self, tmp_path):
        from screener_core import ScreenerCache
        cache = ScreenerCache(str(tmp_path / "cache"))
        result = cache.get("nonexistent", ttl_seconds=3600)
        assert result is None

    def test_invalidate(self, tmp_path):
        from screener_core import ScreenerCache
        cache = ScreenerCache(str(tmp_path / "cache"))
        df = pd.DataFrame({"a": [1]})
        cache.put("test_key", df)
        cache.invalidate("test_key")
        result = cache.get("test_key", ttl_seconds=3600)
        assert result is None

    def test_clear(self, tmp_path):
        from screener_core import ScreenerCache
        cache = ScreenerCache(str(tmp_path / "cache"))
        cache.put("k1", pd.DataFrame({"a": [1]}))
        cache.put("k2", pd.DataFrame({"b": [2]}))
        cache.clear()
        assert cache.get("k1", 3600) is None
        assert cache.get("k2", 3600) is None


class TestTier1BulkData:
    """Tests for _tier1_bulk_data and _get_latest_trade_date."""

    def test_get_latest_trade_date(self, tmp_path):
        screener = _make_screener(tmp_path)
        cal_df = pd.DataFrame({
            "cal_date": ["20260306", "20260307", "20260308"],
            "is_open": [1, 0, 0],
        })
        screener._safe_call = MagicMock(return_value=cal_df)
        result = screener._get_latest_trade_date()
        assert result == "20260306"

    def test_tier1_bulk_data_merge(self, tmp_path):
        screener = _make_screener(tmp_path)
        sb = _load_bulk_mock("stock_basic_bulk.json")
        db = _load_bulk_mock("daily_basic_bulk.json")
        call_count = 0

        def _mock_call(api_name, **kwargs):
            nonlocal call_count
            call_count += 1
            if api_name == "trade_cal":
                return pd.DataFrame({"cal_date": ["20260306"], "is_open": [1]})
            elif api_name == "stock_basic":
                return sb
            elif api_name == "daily_basic":
                return db
            return pd.DataFrame()

        screener._safe_call = _mock_call
        result = screener._tier1_bulk_data(force_refresh=True)
        assert len(result) == 10
        assert "name" in result.columns
        assert "pe_ttm" in result.columns


# ============================================================
# Feature #98: Tier 1 Filter
# ============================================================


class TestTier1Filter:
    """Tests for _tier1_filter: each filter individually."""

    def _get_screener(self, tmp_path):
        return _make_screener(tmp_path)

    def test_removes_st_stocks(self, tmp_path):
        screener = self._get_screener(tmp_path)
        df = _merged_mock_df()
        result = screener._tier1_filter(df)
        names = result["name"].tolist()
        assert not any("ST" in n for n in names)

    def test_removes_new_ipo(self, tmp_path):
        screener = self._get_screener(tmp_path)
        df = _merged_mock_df()
        result = screener._tier1_filter(df)
        # 301234 listed 2025-01-01 (< 3 years from 2026-03-08)
        assert "301234.SZ" not in result["ts_code"].values

    def test_removes_low_market_cap(self, tmp_path):
        """total_mv < 50000 万 (5亿) should be removed."""
        screener = self._get_screener(tmp_path)
        df = pd.DataFrame({
            "ts_code": ["A.SH"], "name": ["小盘股"],
            "industry": ["测试"], "list_date": ["20100101"],
            "close": [5.0], "pe_ttm": [10.0], "pb": [1.5],
            "total_mv": [30000], "circ_mv": [20000],
            "dv_ttm": [2.0], "turnover_rate": [0.5],
        })
        result = screener._tier1_filter(df)
        assert len(result) == 0

    def test_removes_low_turnover(self, tmp_path):
        screener = self._get_screener(tmp_path)
        df = pd.DataFrame({
            "ts_code": ["A.SH"], "name": ["僵尸股"],
            "industry": ["测试"], "list_date": ["20100101"],
            "close": [5.0], "pe_ttm": [10.0], "pb": [1.5],
            "total_mv": [500000], "circ_mv": [400000],
            "dv_ttm": [2.0], "turnover_rate": [0.05],
        })
        result = screener._tier1_filter(df)
        assert len(result) == 0

    def test_removes_negative_pb(self, tmp_path):
        screener = self._get_screener(tmp_path)
        df = _merged_mock_df()
        result = screener._tier1_filter(df)
        # 000666 has PB = -0.30
        assert "000666.SZ" not in result["ts_code"].values

    def test_removes_high_pb(self, tmp_path):
        screener = self._get_screener(tmp_path)
        df = pd.DataFrame({
            "ts_code": ["A.SH"], "name": ["高PB"],
            "industry": ["测试"], "list_date": ["20100101"],
            "close": [100.0], "pe_ttm": [20.0], "pb": [15.0],
            "total_mv": [500000], "circ_mv": [400000],
            "dv_ttm": [1.0], "turnover_rate": [0.5],
        })
        result = screener._tier1_filter(df)
        assert len(result) == 0

    def test_removes_zero_dividend(self, tmp_path):
        screener = self._get_screener(tmp_path)
        df = _merged_mock_df()
        result = screener._tier1_filter(df)
        # 301234 and 000666 have dv_ttm=0
        for code in ["301234.SZ", "000666.SZ"]:
            assert code not in result["ts_code"].values

    def test_dual_channel_pe(self, tmp_path):
        screener = self._get_screener(tmp_path)
        df = _merged_mock_df()
        result = screener._tier1_filter(df)
        if not result.empty:
            main = result[result["channel"] == "main"]
            obs = result[result["channel"] == "observation"]
            # All main channel PE should be > 0 and <= 50
            if not main.empty:
                assert (main["pe_ttm"] > 0).all()
                assert (main["pe_ttm"] <= 50).all()
            # Observation channel PE should be < 0
            if not obs.empty:
                assert (obs["pe_ttm"] < 0).all()

    def test_high_pe_excluded_from_main(self, tmp_path):
        """PE > 50 should be excluded from main channel."""
        screener = self._get_screener(tmp_path)
        df = _merged_mock_df()
        result = screener._tier1_filter(df)
        # 688981 has pe_ttm=55 — should not be in result (also has low dv_ttm=0.20)
        main = result[result["channel"] == "main"]
        if not main.empty:
            assert (main["pe_ttm"] <= 50).all()

    def test_observation_channel_limit(self, tmp_path):
        cfg = ScreenerConfig(obs_channel_limit=1)
        screener = _make_screener(tmp_path, config=cfg)
        # Create data with multiple negative PE stocks
        df = pd.DataFrame({
            "ts_code": ["A.SH", "B.SH", "C.SH"],
            "name": ["负PE1", "负PE2", "负PE3"],
            "industry": ["测试", "测试", "测试"],
            "list_date": ["20100101", "20100101", "20100101"],
            "close": [10.0, 20.0, 30.0],
            "pe_ttm": [-5.0, -10.0, -15.0],
            "pb": [1.0, 2.0, 3.0],
            "total_mv": [500000, 800000, 300000],
            "circ_mv": [400000, 700000, 250000],
            "dv_ttm": [1.0, 2.0, 0.5],
            "turnover_rate": [0.5, 0.6, 0.3],
        })
        result = screener._tier1_filter(df)
        obs = result[result["channel"] == "observation"]
        assert len(obs) <= 1

    def test_filter_preserves_columns(self, tmp_path):
        screener = self._get_screener(tmp_path)
        df = _merged_mock_df()
        result = screener._tier1_filter(df)
        if not result.empty:
            assert "channel" in result.columns
            assert "ts_code" in result.columns
            assert "pe_ttm" in result.columns

    def test_empty_input(self, tmp_path):
        screener = self._get_screener(tmp_path)
        result = screener._tier1_filter(pd.DataFrame())
        assert result.empty

    def test_full_mock_filter_results(self, tmp_path):
        """With mock data, verify expected stocks pass/fail."""
        screener = self._get_screener(tmp_path)
        df = _merged_mock_df()
        result = screener._tier1_filter(df)
        passed_codes = set(result["ts_code"].values)
        # Should pass: 600887 (normal), 601398 (bank, low PE), 000858, 000001
        # Should fail: 000666 (ST, neg PB, zero div), 301234 (new IPO, zero div)
        # 688981: PE=55 (>50), dv_ttm=0.20 → passes other filters but PE too high
        # 300750: dv_ttm=0.80 → passes, PE=35 → main channel OK
        # 600519: PE=25, PB=8.5 → PB ≤ 10, passes
        # 600100: PE=-12 → observation channel
        assert "000666.SZ" not in passed_codes
        assert "301234.SZ" not in passed_codes


# ============================================================
# Feature #99: Tier 1 Rank & Cut
# ============================================================


class TestTier1RankAndCut:
    """Tests for _tier1_rank_and_cut."""

    def test_sort_order(self, tmp_path):
        screener = _make_screener(tmp_path)
        df = pd.DataFrame({
            "ts_code": ["A.SH", "B.SH", "C.SH"],
            "channel": ["main", "main", "main"],
            "pe_ttm": [10.0, 20.0, 30.0],
            "pb": [2.0, 3.0, 4.0],
            "dv_ttm": [5.0, 3.0, 1.0],
            "total_mv": [100000, 200000, 300000],
        })
        result = screener._tier1_rank_and_cut(df)
        # A.SH should rank highest (lowest PE, PB + highest div)
        assert result.iloc[0]["ts_code"] == "A.SH"
        assert "tier1_score" in result.columns

    def test_channel_merge(self, tmp_path):
        screener = _make_screener(tmp_path)
        df = pd.DataFrame({
            "ts_code": ["A.SH", "B.SH", "C.SH"],
            "channel": ["main", "main", "observation"],
            "pe_ttm": [10.0, 20.0, -5.0],
            "pb": [2.0, 3.0, 1.5],
            "dv_ttm": [5.0, 3.0, 2.0],
            "total_mv": [100000, 200000, 300000],
        })
        result = screener._tier1_rank_and_cut(df)
        assert len(result) == 3
        assert "C.SH" in result["ts_code"].values

    def test_cutoff_at_limit(self, tmp_path):
        cfg = ScreenerConfig(tier2_main_limit=2)
        screener = _make_screener(tmp_path, config=cfg)
        df = pd.DataFrame({
            "ts_code": [f"S{i}.SH" for i in range(5)],
            "channel": ["main"] * 5,
            "pe_ttm": [10.0, 15.0, 20.0, 25.0, 30.0],
            "pb": [2.0] * 5,
            "dv_ttm": [5.0, 4.0, 3.0, 2.0, 1.0],
            "total_mv": [100000] * 5,
        })
        result = screener._tier1_rank_and_cut(df)
        main = result[result["channel"] == "main"]
        assert len(main) <= 2

    def test_observation_gets_zero_score(self, tmp_path):
        screener = _make_screener(tmp_path)
        df = pd.DataFrame({
            "ts_code": ["A.SH", "B.SH"],
            "channel": ["main", "observation"],
            "pe_ttm": [10.0, -5.0],
            "pb": [2.0, 1.5],
            "dv_ttm": [5.0, 2.0],
            "total_mv": [100000, 300000],
        })
        result = screener._tier1_rank_and_cut(df)
        obs = result[result["channel"] == "observation"]
        assert (obs["tier1_score"] == 0.0).all()

    def test_empty_input(self, tmp_path):
        screener = _make_screener(tmp_path)
        result = screener._tier1_rank_and_cut(pd.DataFrame())
        assert result.empty


# ============================================================
# Feature #100: Tier 2 Hard Vetoes
# ============================================================


class TestTier2HardVetoes:
    """Tests for _check_hard_vetoes."""

    def test_high_pledge_ratio_vetoed(self, tmp_path):
        screener = _make_screener(tmp_path)
        pledge_df = pd.DataFrame({
            "ts_code": ["A.SH"], "end_date": ["20251231"],
            "pledge_count": [5], "pledge_ratio": [85.0],
        })
        audit_df = pd.DataFrame({
            "ts_code": ["A.SH"], "end_date": ["20251231"],
            "audit_result": ["标准无保留意见"],
        })

        def _mock_call(api_name, **kwargs):
            if api_name == "pledge_stat":
                return pledge_df
            if api_name == "fina_audit":
                return audit_df
            return pd.DataFrame()

        screener._safe_call = _mock_call
        passed, reason = screener._check_hard_vetoes("A.SH")
        assert not passed
        assert "pledge" in reason.lower()

    def test_non_standard_audit_vetoed(self, tmp_path):
        screener = _make_screener(tmp_path)
        pledge_df = pd.DataFrame({
            "ts_code": ["A.SH"], "end_date": ["20251231"],
            "pledge_count": [1], "pledge_ratio": [10.0],
        })
        audit_df = pd.DataFrame({
            "ts_code": ["A.SH"], "end_date": ["20251231"],
            "audit_result": ["保留意见"],
        })

        def _mock_call(api_name, **kwargs):
            if api_name == "pledge_stat":
                return pledge_df
            if api_name == "fina_audit":
                return audit_df
            return pd.DataFrame()

        screener._safe_call = _mock_call
        passed, reason = screener._check_hard_vetoes("A.SH")
        assert not passed
        assert "audit" in reason.lower()

    def test_clean_stock_passes(self, tmp_path):
        screener = _make_screener(tmp_path)
        pledge_df = pd.DataFrame({
            "ts_code": ["A.SH"], "end_date": ["20251231"],
            "pledge_count": [1], "pledge_ratio": [10.0],
        })
        audit_df = pd.DataFrame({
            "ts_code": ["A.SH"], "end_date": ["20251231"],
            "audit_result": ["标准无保留意见"],
        })

        def _mock_call(api_name, **kwargs):
            if api_name == "pledge_stat":
                return pledge_df
            if api_name == "fina_audit":
                return audit_df
            return pd.DataFrame()

        screener._safe_call = _mock_call
        passed, reason = screener._check_hard_vetoes("A.SH")
        assert passed
        assert reason == ""

    def test_missing_data_passes(self, tmp_path):
        """Missing pledge/audit data should not veto."""
        screener = _make_screener(tmp_path)
        screener._safe_call = MagicMock(return_value=pd.DataFrame())
        passed, reason = screener._check_hard_vetoes("A.SH")
        assert passed


# ============================================================
# Feature #101: Tier 2 Financial Quality
# ============================================================


class TestTier2FinancialQuality:
    """Tests for _check_financial_quality."""

    def _make_fina_df(self, roe=15.0, gm=30.0, debt=50.0, profit_dedt=1e8):
        return pd.DataFrame({
            "ts_code": ["A.SH"], "end_date": ["20251231"],
            "roe_waa": [roe], "grossprofit_margin": [gm],
            "debt_to_assets": [debt], "profit_dedt": [profit_dedt],
        })

    def test_good_stock_passes(self, tmp_path):
        screener = _make_screener(tmp_path)
        screener._safe_call = MagicMock(return_value=self._make_fina_df())
        passed, metrics = screener._check_financial_quality("A.SH")
        assert passed
        assert metrics["roe_waa"] == 15.0

    def test_low_roe_fails(self, tmp_path):
        screener = _make_screener(tmp_path)
        screener._safe_call = MagicMock(return_value=self._make_fina_df(roe=5.0))
        passed, metrics = screener._check_financial_quality("A.SH")
        assert not passed

    def test_low_gross_margin_fails(self, tmp_path):
        screener = _make_screener(tmp_path)
        screener._safe_call = MagicMock(return_value=self._make_fina_df(gm=10.0))
        passed, metrics = screener._check_financial_quality("A.SH")
        assert not passed

    def test_high_debt_fails(self, tmp_path):
        screener = _make_screener(tmp_path)
        screener._safe_call = MagicMock(return_value=self._make_fina_df(debt=80.0))
        passed, metrics = screener._check_financial_quality("A.SH")
        assert not passed

    def test_observation_channel_positive_dedt_passes(self, tmp_path):
        screener = _make_screener(tmp_path)
        screener._safe_call = MagicMock(return_value=self._make_fina_df(profit_dedt=5e7))
        passed, _ = screener._check_financial_quality("A.SH", channel="observation")
        assert passed

    def test_observation_channel_negative_dedt_fails(self, tmp_path):
        screener = _make_screener(tmp_path)
        screener._safe_call = MagicMock(return_value=self._make_fina_df(profit_dedt=-1e7))
        passed, _ = screener._check_financial_quality("A.SH", channel="observation")
        assert not passed

    def test_empty_data_fails(self, tmp_path):
        screener = _make_screener(tmp_path)
        screener._safe_call = MagicMock(return_value=pd.DataFrame())
        passed, _ = screener._check_financial_quality("A.SH")
        assert not passed


# ============================================================
# Feature #102: Factor 2 Penetration Return
# ============================================================


class TestFactor2Metrics:
    """Tests for _extract_factor2_metrics."""

    def test_basic_computation(self, tmp_path):
        screener = _make_screener(tmp_path)
        screener._rf_cache = 2.5  # Rf = 2.5%

        income_df = pd.DataFrame({
            "ts_code": ["A.SH"] * 3,
            "end_date": ["20251231", "20241231", "20231231"],
            "n_income_attr_p": [1e9, 9e8, 8e8],  # yuan
        })
        div_df = pd.DataFrame({
            "ts_code": ["A.SH"] * 3,
            "end_date": ["20251231", "20241231", "20231231"],
            "cash_div_tax": [0.5, 0.45, 0.40],  # per share
            "base_share": [1e8, 1e8, 1e8],  # 万股... actually shares
        })

        call_count = {"n": 0}

        def _mock_call(api_name, **kwargs):
            call_count["n"] += 1
            if api_name == "income":
                return income_df
            if api_name == "dividend":
                return div_df
            return pd.DataFrame()

        screener._safe_call = _mock_call
        # total_mv_wan = 50000 万元 → market cap = 5亿元 = 500百万元
        result = screener._extract_factor2_metrics("A.SH", total_mv_wan=50000)
        assert result["Rf"] == 2.5
        assert result["II"] == max(3.5, 2.5 + 2.0)  # 4.5
        assert result["M"] is not None
        assert result["R"] is not None

    def test_rf_cached_globally(self, tmp_path):
        screener = _make_screener(tmp_path)
        rf_df = pd.DataFrame({
            "trade_date": ["20260306"], "yield": [2.8],
        })

        def _mock_call(api_name, **kwargs):
            if api_name == "yc_cb":
                return rf_df
            return pd.DataFrame()

        screener._safe_call = _mock_call
        result = screener._extract_factor2_metrics("A.SH", total_mv_wan=50000)
        assert screener._rf_cache == 2.8

    def test_missing_dividend_no_crash(self, tmp_path):
        screener = _make_screener(tmp_path)
        screener._rf_cache = 2.5
        income_df = pd.DataFrame({
            "ts_code": ["A.SH"], "end_date": ["20251231"],
            "n_income_attr_p": [1e9],
        })

        def _mock_call(api_name, **kwargs):
            if api_name == "income":
                return income_df
            return pd.DataFrame()

        screener._safe_call = _mock_call
        result = screener._extract_factor2_metrics("A.SH", total_mv_wan=50000)
        assert result["M"] is None
        assert result["R"] is None


# ============================================================
# Feature #103: Factor 4 Valuation Metrics
# ============================================================


class TestFactor4Metrics:
    """Tests for _extract_factor4_metrics."""

    def _make_mock_data(self):
        income_df = pd.DataFrame({
            "ts_code": ["A.SH"], "end_date": ["20251231"],
            "operate_profit": [5e9], "finance_exp": [2e8],
            "n_income_attr_p": [4e9],
        })
        bs_df = pd.DataFrame({
            "ts_code": ["A.SH"], "end_date": ["20251231"],
            "money_cap": [3e9], "trad_asset": [1e9],
            "st_borr": [5e8], "lt_borr": [2e9],
            "bond_payable": [0], "non_cur_liab_due_1y": [3e8],
            "goodwill": [1e8], "total_assets": [3e10],
            "total_hldr_eqy_exc_min_int": [1.5e10],
        })
        cf_df = pd.DataFrame({
            "ts_code": ["A.SH"] * 5,
            "end_date": ["20251231", "20241231", "20231231", "20221231", "20211231"],
            "n_cashflow_act": [6e9, 5.5e9, 5e9, 4.5e9, 4e9],
            "c_pay_acq_const_fiolta": [2e9, 1.8e9, 1.6e9, 1.5e9, 1.4e9],
            "depr_fa_coga_dpba": [8e8, 7e8, 6e8, 5e8, 4e8],
            "amort_intang_assets": [1e8, 1e8, 1e8, 1e8, 1e8],
            "lt_amort_deferred_exp": [5e7, 5e7, 5e7, 5e7, 5e7],
        })
        return income_df, bs_df, cf_df

    def test_ev_ebitda_computation(self, tmp_path):
        screener = _make_screener(tmp_path)
        inc, bs, cf = self._make_mock_data()

        def _mock_call(api_name, **kwargs):
            if api_name == "income":
                return inc
            if api_name == "balancesheet":
                return bs
            if api_name == "cashflow":
                return cf
            return pd.DataFrame()

        screener._safe_call = _mock_call
        # total_mv_wan = 2e7 万元 → mkt_cap = 2e11 yuan = 200000 百万元
        result = screener._extract_factor4_metrics("A.SH", close=100.0,
                                                    total_mv_wan=2e7)
        assert "ev_ebitda" in result
        assert result["ev_ebitda"] > 0

    def test_fcf_yield_positive(self, tmp_path):
        screener = _make_screener(tmp_path)
        inc, bs, cf = self._make_mock_data()

        def _mock_call(api_name, **kwargs):
            if api_name == "income":
                return inc
            if api_name == "balancesheet":
                return bs
            if api_name == "cashflow":
                return cf
            return pd.DataFrame()

        screener._safe_call = _mock_call
        result = screener._extract_factor4_metrics("A.SH", close=100.0,
                                                    total_mv_wan=2e7)
        assert result.get("fcf_yield") is not None
        assert result["fcf_yield"] > 0

    def test_fcf_consistency(self, tmp_path):
        screener = _make_screener(tmp_path)
        inc, bs, cf = self._make_mock_data()

        def _mock_call(api_name, **kwargs):
            if api_name == "income":
                return inc
            if api_name == "balancesheet":
                return bs
            if api_name == "cashflow":
                return cf
            return pd.DataFrame()

        screener._safe_call = _mock_call
        result = screener._extract_factor4_metrics("A.SH", close=100.0,
                                                    total_mv_wan=2e7)
        assert result.get("fcf_consistency") == 1.0  # all 5 years positive

    def test_goodwill_ratio(self, tmp_path):
        screener = _make_screener(tmp_path)
        inc, bs, cf = self._make_mock_data()

        def _mock_call(api_name, **kwargs):
            if api_name == "income":
                return inc
            if api_name == "balancesheet":
                return bs
            if api_name == "cashflow":
                return cf
            return pd.DataFrame()

        screener._safe_call = _mock_call
        result = screener._extract_factor4_metrics("A.SH", close=100.0,
                                                    total_mv_wan=2e7)
        # goodwill=1e8 / total_assets=3e10 = 0.33%
        assert result.get("goodwill_ratio") is not None
        assert result["goodwill_ratio"] < 1.0

    def test_empty_data_no_crash(self, tmp_path):
        screener = _make_screener(tmp_path)
        screener._safe_call = MagicMock(return_value=pd.DataFrame())
        result = screener._extract_factor4_metrics("A.SH", close=100.0,
                                                    total_mv_wan=2e7)
        assert isinstance(result, dict)


# ============================================================
# Feature #104: Floor Price
# ============================================================


class TestFloorPrice:
    """Tests for _extract_floor_price."""

    def test_basic_computation(self, tmp_path):
        screener = _make_screener(tmp_path)
        screener._rf_cache = 2.5

        bs_df = pd.DataFrame({
            "ts_code": ["A.SH"], "end_date": ["20251231"],
            "money_cap": [5e9], "trad_asset": [1e9],
            "st_borr": [5e8], "lt_borr": [1e9],
            "bond_payable": [0], "non_cur_liab_due_1y": [2e8],
            "total_hldr_eqy_exc_min_int": [8e9],
        })
        cf_df = pd.DataFrame({
            "ts_code": ["A.SH"] * 3,
            "end_date": ["20251231", "20241231", "20231231"],
            "n_cashflow_act": [3e9, 2.5e9, 2e9],
            "c_pay_acq_const_fiolta": [1e9, 9e8, 8e8],
        })
        weekly_df = pd.DataFrame({
            "ts_code": ["A.SH"] * 3,
            "trade_date": ["20260101", "20250601", "20200101"],
            "close": [50.0, 45.0, 30.0],
        })
        div_df = pd.DataFrame({
            "ts_code": ["A.SH"] * 3,
            "end_date": ["20251231", "20241231", "20231231"],
            "cash_div_tax": [1.0, 0.9, 0.8],
        })

        def _mock_call(api_name, **kwargs):
            if api_name == "balancesheet":
                return bs_df
            if api_name == "cashflow":
                return cf_df
            if api_name == "weekly":
                return weekly_df
            if api_name == "dividend":
                return div_df
            return pd.DataFrame()

        screener._safe_call = _mock_call
        # close=50, total_mv_wan=100000 → total_shares = 100000*10000/50 = 2e7
        result = screener._extract_floor_price("A.SH", close=50.0,
                                                total_mv_wan=100000)
        assert "baselines" in result
        assert len(result["baselines"]) >= 3  # at least NLA, BVPS, 10yr_low
        assert result["composite_baseline"] is not None
        assert result["premium"] is not None

    def test_10yr_low_included(self, tmp_path):
        screener = _make_screener(tmp_path)
        screener._rf_cache = None
        weekly_df = pd.DataFrame({
            "ts_code": ["A.SH"] * 2,
            "trade_date": ["20260101", "20200101"],
            "close": [50.0, 20.0],
        })

        def _mock_call(api_name, **kwargs):
            if api_name == "weekly":
                return weekly_df
            return pd.DataFrame()

        screener._safe_call = _mock_call
        result = screener._extract_floor_price("A.SH", close=50.0,
                                                total_mv_wan=100000)
        method_names = [n for n, _ in result.get("baselines", [])]
        assert "10yr_low" in method_names

    def test_missing_data_partial_result(self, tmp_path):
        screener = _make_screener(tmp_path)
        screener._rf_cache = None
        screener._safe_call = MagicMock(return_value=pd.DataFrame())
        result = screener._extract_floor_price("A.SH", close=50.0,
                                                total_mv_wan=100000)
        assert isinstance(result, dict)


# ============================================================
# Feature #105: Composite Scoring
# ============================================================


class TestCompositeScoring:
    """Tests for _compute_rankings."""

    def test_percentile_ranking(self, tmp_path):
        screener = _make_screener(tmp_path)
        df = pd.DataFrame({
            "roe_waa": [20.0, 15.0, 10.0, 5.0],
            "fcf_yield": [8.0, 6.0, 4.0, 2.0],
            "R": [5.0, 4.0, 3.0, 2.0],
            "ev_ebitda": [6.0, 8.0, 10.0, 12.0],
            "floor_premium": [10.0, 20.0, 30.0, 50.0],
        })
        result = screener._compute_rankings(df)
        assert "composite_score" in result.columns
        # First row should have highest score (best on all dimensions)
        assert result.iloc[0]["composite_score"] >= result.iloc[-1]["composite_score"]

    def test_sort_order(self, tmp_path):
        screener = _make_screener(tmp_path)
        df = pd.DataFrame({
            "roe_waa": [5.0, 20.0, 10.0],
            "fcf_yield": [2.0, 8.0, 4.0],
            "R": [2.0, 5.0, 3.0],
            "ev_ebitda": [12.0, 6.0, 9.0],
            "floor_premium": [50.0, 10.0, 30.0],
        })
        result = screener._compute_rankings(df)
        scores = result["composite_score"].tolist()
        assert scores == sorted(scores, reverse=True)

    def test_handles_nan(self, tmp_path):
        screener = _make_screener(tmp_path)
        df = pd.DataFrame({
            "roe_waa": [20.0, None, 10.0],
            "fcf_yield": [8.0, 6.0, None],
            "R": [5.0, None, 3.0],
            "ev_ebitda": [6.0, 8.0, None],
            "floor_premium": [10.0, None, 30.0],
        })
        result = screener._compute_rankings(df)
        assert len(result) == 3
        assert not result["composite_score"].isna().all()

    def test_empty_input(self, tmp_path):
        screener = _make_screener(tmp_path)
        result = screener._compute_rankings(pd.DataFrame())
        assert result.empty

    def test_weight_application(self, tmp_path):
        """Verify weights actually affect the scoring."""
        cfg = ScreenerConfig(weight_roe=1.0, weight_fcf_yield=0.0,
                             weight_penetration_r=0.0, weight_ev_ebitda=0.0,
                             weight_floor_premium=0.0)
        screener = _make_screener(tmp_path, config=cfg)
        df = pd.DataFrame({
            "roe_waa": [5.0, 20.0, 10.0],
            "fcf_yield": [100.0, 1.0, 50.0],  # should be ignored
            "R": [100.0, 1.0, 50.0],
            "ev_ebitda": [1.0, 100.0, 50.0],
            "floor_premium": [1.0, 100.0, 50.0],
        })
        result = screener._compute_rankings(df)
        # With 100% weight on ROE, highest ROE should win
        assert result.iloc[0]["roe_waa"] == 20.0


# ============================================================
# Feature #106: Pipeline & Export
# ============================================================


class TestPipelineAndExport:
    """Tests for run() and export methods."""

    def test_tier1_only_mode(self, tmp_path):
        screener = _make_screener(tmp_path)
        sb = _load_bulk_mock("stock_basic_bulk.json")
        db = _load_bulk_mock("daily_basic_bulk.json")

        def _mock_call(api_name, **kwargs):
            if api_name == "trade_cal":
                return pd.DataFrame({"cal_date": ["20260306"], "is_open": [1]})
            if api_name == "stock_basic":
                return sb
            if api_name == "daily_basic":
                return db
            return pd.DataFrame()

        screener._safe_call = _mock_call
        result = screener.run(tier1_only=True)
        assert not result.empty
        assert "channel" in result.columns

    def test_csv_export(self, tmp_path):
        screener = _make_screener(tmp_path)
        df = pd.DataFrame({
            "ts_code": ["A.SH"], "name": ["测试"],
            "composite_score": [0.85],
        })
        csv_path = str(tmp_path / "test.csv")
        screener.export_csv(df, csv_path)
        assert os.path.exists(csv_path)
        loaded = pd.read_csv(csv_path)
        assert len(loaded) == 1

    def test_html_export(self, tmp_path):
        screener = _make_screener(tmp_path)
        df = pd.DataFrame({
            "ts_code": ["A.SH"], "name": ["测试"],
            "composite_score": [0.85],
        })
        html_path = str(tmp_path / "test.html")
        screener.export_html(df, html_path)
        assert os.path.exists(html_path)
        with open(html_path) as f:
            content = f.read()
        assert "龟龟选股器" in content

    def test_full_pipeline_mock(self, tmp_path):
        """End-to-end mock pipeline with tier2_limit=1."""
        screener = _make_screener(tmp_path)
        sb = _load_bulk_mock("stock_basic_bulk.json")
        db = _load_bulk_mock("daily_basic_bulk.json")

        # Mock all API calls
        pledge_df = pd.DataFrame({
            "ts_code": ["600887.SH"], "end_date": ["20251231"],
            "pledge_count": [1], "pledge_ratio": [5.0],
        })
        audit_df = pd.DataFrame({
            "ts_code": ["600887.SH"], "end_date": ["20251231"],
            "audit_result": ["标准无保留意见"],
        })
        fina_df = pd.DataFrame({
            "ts_code": ["600887.SH"], "end_date": ["20251231"],
            "roe_waa": [20.0], "grossprofit_margin": [35.0],
            "debt_to_assets": [45.0], "profit_dedt": [5e9],
        })
        income_df = pd.DataFrame({
            "ts_code": ["600887.SH"] * 3,
            "end_date": ["20251231", "20241231", "20231231"],
            "operate_profit": [5e9, 4.5e9, 4e9],
            "finance_exp": [2e8, 1.5e8, 1e8],
            "n_income_attr_p": [4e9, 3.5e9, 3e9],
        })
        bs_df = pd.DataFrame({
            "ts_code": ["600887.SH"], "end_date": ["20251231"],
            "money_cap": [3e9], "trad_asset": [1e9],
            "st_borr": [5e8], "lt_borr": [2e9],
            "bond_payable": [0], "non_cur_liab_due_1y": [3e8],
            "goodwill": [1e8], "total_assets": [3e10],
            "total_hldr_eqy_exc_min_int": [1.5e10],
        })
        cf_df = pd.DataFrame({
            "ts_code": ["600887.SH"] * 3,
            "end_date": ["20251231", "20241231", "20231231"],
            "n_cashflow_act": [6e9, 5.5e9, 5e9],
            "c_pay_acq_const_fiolta": [2e9, 1.8e9, 1.6e9],
            "depr_fa_coga_dpba": [8e8, 7e8, 6e8],
            "amort_intang_assets": [1e8, 1e8, 1e8],
            "lt_amort_deferred_exp": [5e7, 5e7, 5e7],
        })
        weekly_df = pd.DataFrame({
            "ts_code": ["600887.SH"] * 3,
            "trade_date": ["20260101", "20250601", "20200101"],
            "close": [28.0, 25.0, 15.0],
        })
        div_df = pd.DataFrame({
            "ts_code": ["600887.SH"] * 3,
            "end_date": ["20251231", "20241231", "20231231"],
            "cash_div_tax": [1.0, 0.9, 0.8],
            "base_share": [6.4e9, 6.4e9, 6.4e9],
        })
        rf_df = pd.DataFrame({
            "trade_date": ["20260306"], "yield": [2.5],
        })

        def _mock_call(api_name, **kwargs):
            if api_name == "trade_cal":
                return pd.DataFrame({"cal_date": ["20260306"], "is_open": [1]})
            if api_name == "stock_basic":
                return sb
            if api_name == "daily_basic":
                return db
            if api_name == "pledge_stat":
                return pledge_df
            if api_name == "fina_audit":
                return audit_df
            if api_name == "fina_indicator":
                return fina_df
            if api_name == "income":
                return income_df
            if api_name == "balancesheet":
                return bs_df
            if api_name == "cashflow":
                return cf_df
            if api_name == "weekly":
                return weekly_df
            if api_name == "dividend":
                return div_df
            if api_name == "yc_cb":
                return rf_df
            return pd.DataFrame()

        screener._safe_call = _mock_call
        result = screener.run(tier2_limit=1)
        # Should have at least 1 result (or 0 if all vetoed)
        assert isinstance(result, pd.DataFrame)
