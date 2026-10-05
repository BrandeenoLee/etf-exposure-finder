"""
End-to-end test on synthetic N-PORT data. Run with:  python -m pytest
"""

import sys
from pathlib import Path

import pandas as pd

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
sys.path.insert(0, str(Path(__file__).resolve().parent))

import etf_exposure  # noqa: E402
from make_sample_data import GME_CUSIP, build_all  # noqa: E402


def run(tmp_path, *extra):
    build_all(tmp_path)
    out = tmp_path / "out"
    code = etf_exposure.main(["--tickers", "GME", "--offline", "--data-dir", str(tmp_path),
                              "--out", str(out), "--start", "2020q4", "--end", "2021q1", *extra])
    assert code == 0
    book = next(out.glob("*.xlsx"))
    return pd.read_excel(book, sheet_name=None)


def test_finds_holders_and_cusip(tmp_path):
    sheets = run(tmp_path)
    assert sheets["CUSIPs"]["CUSIPS_USED"].iloc[0] == GME_CUSIP
    snap = sheets["Latest_Snapshot"]
    # Vanguard's holding has no ticker tag, only a CUSIP; it must still be found
    assert "Vanguard Extended Market Index Fund" in set(snap["FUND_NAME"])
    # The unrelated Macy's holding must not appear
    assert not sheets["All_Holdings"]["FUND_NAME"].isna().all()


def test_amendment_replaces_original(tmp_path):
    snap = run(tmp_path)["Latest_Snapshot"]
    xrt = snap[snap["FUND_NAME"] == "SPDR S&P Retail ETF"].iloc[0]
    assert round(xrt["WEIGHT_PCT"], 2) == 10.30      # amended value, not 9.80
    assert xrt["SHARES_HELD"] == 120000


def test_weights_over_time(tmp_path):
    grid = run(tmp_path)["Weights_GME"].set_index("FUND")
    row = [i for i in grid.index if i.startswith("SPDR S&P Retail ETF")][0]
    assert grid.loc[row, "2020-12"] == 1.5
    assert grid.loc[row, "2021-03"] == 10.3


def test_derivatives_separated(tmp_path):
    d = run(tmp_path)["Derivatives"]
    assert set(d["REF_TYPE"]) == {"single-name", "basket/index component"}
    # Swaps must not be counted as share holdings
    snap = run(tmp_path)["Latest_Snapshot"]
    assert "T-Rex 2X Long GME Daily Target ETF" not in set(snap["FUND_NAME"])


def test_etf_only_filter(tmp_path):
    snap = run(tmp_path, "--etf-only")["Latest_Snapshot"]
    assert set(snap["FUND_NAME"]) == {"SPDR S&P Retail ETF"}


def test_peak_weights_and_report_months(tmp_path):
    peaks = run(tmp_path)["Peak_Weights"]
    xrt = peaks[peaks["FUND_NAME"] == "SPDR S&P Retail ETF"].iloc[0]
    assert round(xrt["PEAK_WEIGHT_PCT"], 2) == 10.30
    assert str(xrt["PEAK_DATE"]).startswith("2021-03-31")
    assert xrt["FIRST_WEIGHT_PCT"] == 1.5
    assert xrt["REPORT_MONTHS"] == "2020-12, 2021-03"


def test_size_qualified_funds_rank_first(tmp_path):
    # With a $1B threshold only Vanguard ($90B) qualifies, so it must rank first
    snap = run(tmp_path, "--min-aum", "1e9")["Latest_Snapshot"]
    assert snap.iloc[0]["FUND_NAME"] == "Vanguard Extended Market Index Fund"
    assert bool(snap.iloc[0]["SIZE_OK"]) is True


def test_ticker_overrides(tmp_path):
    ov = tmp_path / "ov.csv"
    ov.write_text("MATCH,TICKER\ngrowth opportunities fund,GROWX\nS000002,VXF\n")
    snap = run(tmp_path, "--overrides", str(ov))["Latest_Snapshot"].set_index("FUND_NAME")
    assert snap.loc["Growth Opportunities Fund", "FUND_TICKERS"] == "GROWX"
    assert snap.loc["Vanguard Extended Market Index Fund", "FUND_TICKERS"] == "VXF"
