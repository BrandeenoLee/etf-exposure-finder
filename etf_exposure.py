#!/usr/bin/env python3
"""
etf_exposure.py - Find every fund (and flag likely ETFs) holding a given stock,
using the SEC's free Form N-PORT bulk data sets.

Works for any US-listed stock: GME, AMC, or anything else.

What it does
------------
1. Downloads the quarterly N-PORT data set zips from sec.gov (and caches them).
2. Finds every fund holding the stock(s) you ask for, either directly
   (shares) or through derivatives (options/swaps that reference the stock).
3. Joins in fund name, net assets, report date, and the fund's own ticker(s).
4. Writes an Excel workbook + CSVs ranking funds by weight and size for each
   reporting date, so you can see which funds could act as a shorting vehicle
   and how that changed over time.

Examples
--------
  # GME, every quarter the SEC has published (Oct 2019 onward)
  python etf_exposure.py --tickers GME --email you@example.com

  # GME and AMC, just 2020 Q4 through 2021 Q3
  python etf_exposure.py --tickers GME AMC --start 2020q4 --end 2021q3 --email you@example.com

  # Only the latest published quarter, ETFs only
  python etf_exposure.py --tickers GME --latest --etf-only --email you@example.com

Notes
-----
* The SEC requires a User-Agent with contact info for automated downloads,
  which is why --email is required. It is sent only to sec.gov.
* Each quarterly zip is ~350-700 MB. They are cached in --data-dir so later
  runs (including runs for other tickers) don't re-download.
  Use --delete-zips to remove each zip after it is processed.
* N-PORT data is published with a lag (the newest data is usually 3-6 months
  old). For today's holdings, check the issuer's daily holdings file.
"""

import argparse
import csv
import io
import json
import math
import re
import sys
import time
import zipfile
from datetime import date
from pathlib import Path

try:
    import pandas as pd
    import requests
except ImportError:
    sys.exit("Missing packages. Run:  pip install pandas requests openpyxl")

NPORT_URL = "https://www.sec.gov/files/dera/data/form-n-port-data-sets/{q}_nport.zip"
FUND_TICKER_URL = "https://www.sec.gov/files/company_tickers_mf.json"
FIRST_QUARTER = (2019, 4)
CHUNK = 500_000

# Words in a fund's series name that suggest it is an exchange-traded fund.
# This is a heuristic: some ETFs (notably Vanguard's, which are share classes
# of index mutual funds) won't match. Use --all-funds output to check.
ETF_NAME_PATTERN = re.compile(
    r"\bETF\b|\bETFs\b|EXCHANGE[- ]TRADED|\bSPDR\b|\bISHARES\b|POWERSHARES|"
    r"\bPROSHARES\b|\bDIREXION\b|\bYIELDMAX\b|\bT-REX\b|\bROUNDHILL\b|\bDEFIANCE\b",
    re.IGNORECASE,
)


# ----------------------------------------------------------------------------
# Helpers
# ----------------------------------------------------------------------------

def log(msg):
    print(f"[{time.strftime('%H:%M:%S')}] {msg}", flush=True)


def parse_quarter(text):
    m = re.fullmatch(r"(\d{4})\s*[qQ]([1-4])", text.strip())
    if not m:
        raise argparse.ArgumentTypeError(f"Quarter must look like 2021q1, got '{text}'")
    return int(m.group(1)), int(m.group(2))


def quarter_range(start, end):
    y, q = start
    out = []
    while (y, q) <= end:
        out.append(f"{y}q{q}")
        q += 1
        if q == 5:
            y, q = y + 1, 1
    return out


def current_quarter():
    t = date.today()
    return t.year, (t.month - 1) // 3 + 1


def to_num(series):
    return pd.to_numeric(series, errors="coerce")


def session_for(email):
    s = requests.Session()
    s.headers.update({
        "User-Agent": f"etf-exposure-research-script {email}",
        "Accept-Encoding": "gzip, deflate",
    })
    return s


# ----------------------------------------------------------------------------
# Downloading
# ----------------------------------------------------------------------------

def download_quarter(sess, q, data_dir):
    """Download one quarterly zip if not cached. Returns path or None if unavailable."""
    path = data_dir / f"{q}_nport.zip"
    if path.exists() and zipfile.is_zipfile(path):
        log(f"{q}: using cached {path.name}")
        return path
    url = NPORT_URL.format(q=q)
    log(f"{q}: downloading {url}")
    tmp = path.with_suffix(".part")
    for attempt in range(3):
        try:
            with sess.get(url, stream=True, timeout=120) as r:
                if r.status_code == 404:
                    log(f"{q}: not published (404), skipping")
                    return None
                r.raise_for_status()
                total = int(r.headers.get("content-length", 0))
                done = 0
                last = time.time()
                with open(tmp, "wb") as f:
                    for block in r.iter_content(chunk_size=1 << 20):
                        f.write(block)
                        done += len(block)
                        if time.time() - last > 10:
                            pct = f"{done / total:.0%}" if total else f"{done >> 20} MB"
                            log(f"{q}: {pct}")
                            last = time.time()
            tmp.rename(path)
            if not zipfile.is_zipfile(path):
                path.unlink()
                raise IOError("downloaded file is not a valid zip")
            return path
        except Exception as e:  # noqa: BLE001
            log(f"{q}: download attempt {attempt + 1} failed: {e}")
            time.sleep(5 * (attempt + 1))
    log(f"{q}: giving up")
    return None


def load_fund_tickers(sess, cache_dir, offline=False):
    """SERIES_ID -> 'XRT' (or 'VTI, VTSAX, ...'). Best effort."""
    cache = cache_dir / "company_tickers_mf.json"
    data = None
    if not offline:
        try:
            r = sess.get(FUND_TICKER_URL, timeout=60)
            r.raise_for_status()
            data = r.json()
            cache.write_text(json.dumps(data))
        except Exception as e:  # noqa: BLE001
            log(f"Could not fetch fund ticker list ({e}); using cache if present")
    if data is None and cache.exists():
        data = json.loads(cache.read_text())
    if not data:
        return {}
    fields = data.get("fields", [])
    try:
        si, sy = fields.index("seriesId"), fields.index("symbol")
    except ValueError:
        return {}
    out = {}
    for row in data.get("data", []):
        sid, sym = row[si], row[sy]
        if sid and sym:
            out.setdefault(sid, set()).add(str(sym).upper())
    return {k: ", ".join(sorted(v)) for k, v in out.items()}


# ----------------------------------------------------------------------------
# Reading the zip
# ----------------------------------------------------------------------------

def member(zf, table):
    """Find a table's file inside the zip regardless of folder/extension."""
    for name in zf.namelist():
        base = name.rsplit("/", 1)[-1].upper()
        if base.split(".")[0] == table:
            return name
    return None


def read_table(zf, table, usecols, chunks=False):
    name = member(zf, table)
    if name is None:
        return iter(()) if chunks else pd.DataFrame(columns=usecols)
    raw = zf.open(name)
    text = io.TextIOWrapper(raw, encoding="utf-8", errors="replace", newline="")
    # Read header to keep only the columns that exist in this vintage
    header = text.readline().rstrip("\r\n").split("\t")
    cols = [c for c in usecols if c in header]
    kwargs = dict(sep="\t", dtype=str, quoting=csv.QUOTE_NONE, names=header,
                  usecols=cols, header=None, on_bad_lines="skip", keep_default_na=False)
    if chunks:
        return pd.read_csv(text, chunksize=CHUNK, **kwargs)
    return pd.read_csv(text, **kwargs)


def scan_quarter(zip_path, targets, include_derivs=True):
    """
    targets: dict ticker -> set of CUSIPs (may be empty, filled by discovery).
    Returns (direct_df, deriv_df, discovered: dict ticker -> Counter of CUSIPs)
    """
    tick_set = set(targets)
    cusip_to_tick = {c: t for t, cs in targets.items() for c in cs}

    with zipfile.ZipFile(zip_path) as zf:
        # 1) Holdings whose ticker identifier matches
        id_to_tick = {}
        for ch in read_table(zf, "IDENTIFIERS", ["HOLDING_ID", "IDENTIFIER_TICKER"], chunks=True):
            t = ch["IDENTIFIER_TICKER"].str.strip().str.upper()
            # Some filers write "GME US" or "GME UN"; take the first token
            t = t.str.split().str[0]
            m = t.isin(tick_set)
            id_to_tick.update(zip(ch["HOLDING_ID"][m], t[m]))

        # 2) Derivatives that reference the stock
        deriv_ref = {}  # holding_id -> (ticker, ref_type, component_value, component_notional)
        if include_derivs:
            def match(ch):
                """Vectorized: ticker from TICKER column, else from CUSIP."""
                t = ch["TICKER"].str.strip().str.upper().str.split().str[0]
                c = ch["CUSIP"].str.strip().str.upper()
                tk = t.where(t.isin(tick_set), c.map(cusip_to_tick))
                return tk

            for ch in read_table(zf, "DESC_REF_OTHER", ["HOLDING_ID", "CUSIP", "TICKER"], chunks=True):
                tk = match(ch)
                hit = tk.notna()
                for hid, t in zip(ch["HOLDING_ID"][hit], tk[hit]):
                    deriv_ref[hid] = (t, "single-name", None, None)
            for ch in read_table(zf, "DESC_REF_INDEX_COMPONENT",
                                 ["HOLDING_ID", "CUSIP", "TICKER", "NOTIONAL_AMOUNT", "VALUE"], chunks=True):
                tk = match(ch)
                hit = tk.notna()
                for hid, t, v, n in zip(ch["HOLDING_ID"][hit], tk[hit], ch["VALUE"][hit],
                                        ch["NOTIONAL_AMOUNT"][hit]):
                    if hid not in deriv_ref:
                        deriv_ref[hid] = (t, "basket/index component", v, n)

        # 3) Holdings table: keep matches by id or CUSIP
        hcols = ["ACCESSION_NUMBER", "HOLDING_ID", "ISSUER_NAME", "ISSUER_TITLE", "ISSUER_CUSIP",
                 "BALANCE", "UNIT", "CURRENCY_VALUE", "PERCENTAGE", "PAYOFF_PROFILE",
                 "ASSET_CAT", "DERIVATIVE_CAT"]
        keep = []
        want_ids = set(id_to_tick) | set(deriv_ref)
        want_cusips = set(cusip_to_tick)
        for ch in read_table(zf, "FUND_REPORTED_HOLDING", hcols, chunks=True):
            cus = ch["ISSUER_CUSIP"].str.strip().str.upper()
            m = ch["HOLDING_ID"].isin(want_ids) | cus.isin(want_cusips)
            if m.any():
                keep.append(ch[m])
        hold = pd.concat(keep, ignore_index=True) if keep else pd.DataFrame(columns=hcols)

        if hold.empty:
            return pd.DataFrame(), pd.DataFrame(), {}

        # Assign ticker to each row
        hold["ISSUER_CUSIP"] = hold["ISSUER_CUSIP"].str.strip().str.upper()
        hold["TICKER"] = hold["HOLDING_ID"].map(id_to_tick)
        hold["TICKER"] = hold["TICKER"].fillna(hold["ISSUER_CUSIP"].map(cusip_to_tick))
        hold["TICKER"] = hold["TICKER"].fillna(hold["HOLDING_ID"].map(lambda h: deriv_ref.get(h, (None,))[0]))
        hold["IS_DERIV"] = hold["HOLDING_ID"].isin(deriv_ref) & (hold["DERIVATIVE_CAT"].str.strip() != "")

        # CUSIP discovery: which CUSIPs do ticker-tagged direct holdings use?
        discovered = {}
        direct_tagged = hold[hold["HOLDING_ID"].isin(id_to_tick) & ~hold["IS_DERIV"]]
        for tk, grp in direct_tagged.groupby("TICKER"):
            discovered[tk] = grp["ISSUER_CUSIP"][grp["ISSUER_CUSIP"].str.len() == 9].value_counts()

        # Fund-level tables, only for accessions we need
        accs = set(hold["ACCESSION_NUMBER"])
        sub = read_table(zf, "SUBMISSION", ["ACCESSION_NUMBER", "FILING_DATE", "SUB_TYPE",
                                            "REPORT_DATE", "REPORT_ENDING_PERIOD"])
        reg = read_table(zf, "REGISTRANT", ["ACCESSION_NUMBER", "CIK", "REGISTRANT_NAME"])
        fri_cols = ["ACCESSION_NUMBER", "SERIES_NAME", "SERIES_ID", "NET_ASSETS",
                    "SALES_FLOW_MON1", "SALES_FLOW_MON2", "SALES_FLOW_MON3",
                    "REDEMPTION_FLOW_MON1", "REDEMPTION_FLOW_MON2", "REDEMPTION_FLOW_MON3"]
        fri = read_table(zf, "FUND_REPORTED_INFO", fri_cols)
        info = (sub[sub["ACCESSION_NUMBER"].isin(accs)]
                .merge(reg[reg["ACCESSION_NUMBER"].isin(accs)], on="ACCESSION_NUMBER", how="left")
                .merge(fri[fri["ACCESSION_NUMBER"].isin(accs)], on="ACCESSION_NUMBER", how="left"))

    hold = hold.merge(info, on="ACCESSION_NUMBER", how="left")

    # Split direct vs derivative
    direct = hold[~hold["IS_DERIV"]].copy()
    deriv = hold[hold["IS_DERIV"]].copy()
    if not deriv.empty:
        deriv["REF_TYPE"] = deriv["HOLDING_ID"].map(lambda h: deriv_ref[h][1])
        deriv["TICKER"] = deriv["HOLDING_ID"].map(lambda h: deriv_ref[h][0])
        deriv["COMPONENT_VALUE"] = deriv["HOLDING_ID"].map(lambda h: deriv_ref[h][2])
        deriv["COMPONENT_NOTIONAL"] = deriv["HOLDING_ID"].map(lambda h: deriv_ref[h][3])
    return direct, deriv, discovered


# ----------------------------------------------------------------------------
# Post-processing
# ----------------------------------------------------------------------------

def dedupe_latest(df):
    """Keep only the latest filing (amendments win) per fund per report date."""
    if df.empty:
        return df
    df = df.copy()
    df["FILING_DATE_D"] = pd.to_datetime(df["FILING_DATE"], errors="coerce")
    key = ["SERIES_ID", "REPORT_DATE"]
    latest = (df.sort_values("FILING_DATE_D")
                .groupby(key, dropna=False)["ACCESSION_NUMBER"].last().reset_index())
    out = df.merge(latest, on=key + ["ACCESSION_NUMBER"], how="inner")
    return out.drop(columns=["FILING_DATE_D"])


def tidy(direct, fund_tickers, min_aum):
    if direct.empty:
        return direct
    d = direct.copy()
    d["REPORT_DATE"] = pd.to_datetime(d["REPORT_DATE"], errors="coerce").dt.date
    for c in ["BALANCE", "CURRENCY_VALUE", "PERCENTAGE", "NET_ASSETS"] + \
             [c for c in d.columns if "_FLOW_MON" in c]:
        if c in d:
            d[c] = to_num(d[c])
    # A fund can report the same stock on several lines (e.g. lots, or a short line)
    d["SIGNED_VALUE"] = d["CURRENCY_VALUE"].where(d["PAYOFF_PROFILE"].str.lower() != "short",
                                                  -d["CURRENCY_VALUE"])
    d["SIGNED_SHARES"] = d["BALANCE"].where(d["PAYOFF_PROFILE"].str.lower() != "short", -d["BALANCE"])
    d["SIGNED_PCT"] = d["PERCENTAGE"].where(d["PAYOFF_PROFILE"].str.lower() != "short", -d["PERCENTAGE"])
    flows_in = d[[c for c in d.columns if c.startswith("SALES_FLOW_MON")]].sum(axis=1, min_count=1)
    flows_out = d[[c for c in d.columns if c.startswith("REDEMPTION_FLOW_MON")]].sum(axis=1, min_count=1)
    d["NET_FLOW_3MO"] = flows_in - flows_out

    g = (d.groupby(["TICKER", "REPORT_DATE", "SERIES_ID"], dropna=False)
           .agg(FUND_NAME=("SERIES_NAME", "first"),
                REGISTRANT=("REGISTRANT_NAME", "first"),
                CIK=("CIK", "first"),
                SHARES_HELD=("SIGNED_SHARES", "sum"),
                VALUE_USD=("SIGNED_VALUE", "sum"),
                WEIGHT_PCT=("SIGNED_PCT", "sum"),
                FUND_NET_ASSETS=("NET_ASSETS", "first"),
                NET_FLOW_3MO=("NET_FLOW_3MO", "first"),
                HAS_SHORT_LINE=("PAYOFF_PROFILE", lambda s: (s.str.lower() == "short").any()),
                FILING_DATE=("FILING_DATE", "first"),
                ACCESSION_NUMBER=("ACCESSION_NUMBER", "first"))
           .reset_index())
    g["FUND_TICKERS"] = g["SERIES_ID"].map(fund_tickers).fillna("")
    g["LIKELY_ETF"] = g["FUND_NAME"].fillna("").str.contains(ETF_NAME_PATTERN) | \
                      g["REGISTRANT"].fillna("").str.contains(ETF_NAME_PATTERN)
    g["NET_FLOW_PCT_OF_ASSETS"] = 100 * g["NET_FLOW_3MO"] / g["FUND_NET_ASSETS"]

    # "Vehicle candidate": meaningful weight AND big enough fund to short in size.
    # Score = weight x log10(net assets): favours funds that are both concentrated and large.
    g["VEHICLE_SCORE"] = g["WEIGHT_PCT"].clip(lower=0) * \
        g["FUND_NET_ASSETS"].clip(lower=1).map(lambda x: math.log10(x) if x and x > 0 else 0)
    g["SIZE_OK"] = g["FUND_NET_ASSETS"] >= min_aum
    return g.sort_values(["TICKER", "REPORT_DATE", "WEIGHT_PCT"], ascending=[True, True, False])


def period_rollup(g):
    """One row per ticker per report month: how many funds, total shares, top fund."""
    if g.empty:
        return g
    g = g.copy()
    g["MONTH"] = pd.to_datetime(g["REPORT_DATE"]).dt.to_period("M").astype(str)
    rows = []
    for (tk, m), grp in g.groupby(["TICKER", "MONTH"]):
        etf = grp[grp["LIKELY_ETF"]]
        top = grp.sort_values("WEIGHT_PCT", ascending=False).iloc[0]
        rows.append({
            "TICKER": tk, "MONTH": m,
            "FUNDS_REPORTING": grp["SERIES_ID"].nunique(),
            "LIKELY_ETFS_REPORTING": etf["SERIES_ID"].nunique(),
            "TOTAL_SHARES_ALL_FUNDS": grp["SHARES_HELD"].sum(),
            "TOTAL_SHARES_LIKELY_ETFS": etf["SHARES_HELD"].sum(),
            "TOTAL_VALUE_USD_ALL_FUNDS": grp["VALUE_USD"].sum(),
            "FUNDS_WEIGHT_OVER_2PCT": int((grp["WEIGHT_PCT"] >= 2).sum()),
            "TOP_FUND_BY_WEIGHT": top["FUND_NAME"],
            "TOP_FUND_TICKERS": top["FUND_TICKERS"],
            "TOP_WEIGHT_PCT": top["WEIGHT_PCT"],
        })
    return pd.DataFrame(rows)


def latest_snapshot(g, top_n):
    """For each ticker, each fund's most recent report, ranked."""
    if g.empty:
        return g
    last = g.sort_values("REPORT_DATE").groupby(["TICKER", "SERIES_ID"]).tail(1)
    # Drop funds whose latest report is stale (>9 months older than the newest)
    newest = pd.to_datetime(last.groupby("TICKER")["REPORT_DATE"].transform("max"))
    last = last[pd.to_datetime(last["REPORT_DATE"]) >= newest - pd.DateOffset(months=9)]
    last = last.sort_values(["TICKER", "VEHICLE_SCORE"], ascending=[True, False])
    return last.groupby("TICKER").head(top_n)


def weight_pivot(g, top_funds=60):
    """Fund x report-month table of weights, for funds that ever ranked high."""
    if g.empty:
        return {}
    out = {}
    for tk, grp in g.groupby("TICKER"):
        grp = grp.copy()
        grp["MONTH"] = pd.to_datetime(grp["REPORT_DATE"]).dt.to_period("M").astype(str)
        best = grp.groupby("SERIES_ID")["WEIGHT_PCT"].max().nlargest(top_funds).index
        sub = grp[grp["SERIES_ID"].isin(best)]
        label = sub["FUND_NAME"].fillna("") + " [" + sub["FUND_TICKERS"].fillna("") + "]"
        sub = sub.assign(FUND=label)
        p = sub.pivot_table(index="FUND", columns="MONTH", values="WEIGHT_PCT", aggfunc="max")
        p["MAX_WEIGHT"] = p.max(axis=1)
        out[tk] = p.sort_values("MAX_WEIGHT", ascending=False).round(3)
    return out


def tidy_derivs(deriv, fund_tickers):
    if deriv.empty:
        return deriv
    d = deriv.copy()
    d["REPORT_DATE"] = pd.to_datetime(d["REPORT_DATE"], errors="coerce").dt.date
    for c in ["BALANCE", "CURRENCY_VALUE", "PERCENTAGE", "NET_ASSETS",
              "COMPONENT_VALUE", "COMPONENT_NOTIONAL"]:
        if c in d:
            d[c] = to_num(d[c])
    d["FUND_TICKERS"] = d["SERIES_ID"].map(fund_tickers).fillna("")
    keep = ["TICKER", "REPORT_DATE", "SERIES_NAME", "FUND_TICKERS", "REGISTRANT_NAME",
            "DERIVATIVE_CAT", "REF_TYPE", "ISSUER_NAME", "ISSUER_TITLE", "PAYOFF_PROFILE",
            "BALANCE", "UNIT", "CURRENCY_VALUE", "PERCENTAGE", "COMPONENT_VALUE",
            "COMPONENT_NOTIONAL", "NET_ASSETS", "SERIES_ID", "ACCESSION_NUMBER"]
    return d[[c for c in keep if c in d.columns]].sort_values(["TICKER", "REPORT_DATE"])


# ----------------------------------------------------------------------------
# Output
# ----------------------------------------------------------------------------

README_TEXT = """\
HOW TO READ THIS WORKBOOK

Source: SEC Form N-PORT data sets (registered funds' reported portfolio holdings).
Each row is what a fund reported holding on its REPORT_DATE, not today.

Sheets
  Latest_Snapshot  Each fund's most recent report, ranked by VEHICLE_SCORE.
  Rollup           One row per stock per month: fund count, total shares held by funds,
                   top fund by weight.
  Weights_<TICKER> Fund x month grid of the stock's weight in each fund (%). Funds that
                   ever ranked in the top 60 by weight.
  All_Holdings     Every fund/report-date row (direct share holdings).
  Derivatives      Options/swaps/baskets that reference the stock (swap-based leveraged
                   funds, option-income funds, index swaps that include the stock).
  CUSIPs           CUSIP(s) used to match each ticker. Verify these.

Key columns
  WEIGHT_PCT       Stock's % of the fund's net assets (fund-reported). Short lines count negative.
  FUND_NET_ASSETS  Fund size in USD. Small funds are poor vehicles for large shorts.
  VEHICLE_SCORE    WEIGHT_PCT x log10(net assets). Higher = more concentrated AND bigger.
                   This is a rough screen, not a measure of actual shorting.
  SIZE_OK          Net assets at or above the --min-aum threshold.
  LIKELY_ETF       Name-based guess. Vanguard ETFs (share classes of index funds) and some
                   others will show False; check FUND_TICKERS.
  NET_FLOW_3MO     Shares sold minus redeemed (USD) over the 3 months in the filing. For ETFs
                   this reflects creations minus redemptions.

Caveats
  * Data lags: the newest quarter is typically 3-6 months behind today.
  * Holding a stock does not mean a fund is being used to short it. To test that, compare
    a candidate ETF's short interest, FTDs, and shares outstanding against the stock's.
  * Liquidity (trading volume, borrow availability) of each ETF is not in N-PORT.
"""


def write_outputs(out_dir, tickers, g, rollup, snap, pivots, derivs, cusips, etf_only):
    out_dir.mkdir(parents=True, exist_ok=True)
    stamp = date.today().isoformat()
    base = f"{'_'.join(tickers)}_fund_exposure_{stamp}"
    xlsx = out_dir / f"{base}.xlsx"

    if etf_only and not g.empty:
        g = g[g["LIKELY_ETF"]]
        snap = snap[snap["LIKELY_ETF"]] if not snap.empty else snap

    g.to_csv(out_dir / f"{base}_all_holdings.csv", index=False)
    if not derivs.empty:
        derivs.to_csv(out_dir / f"{base}_derivatives.csv", index=False)

    snap_cols = ["TICKER", "REPORT_DATE", "FUND_NAME", "FUND_TICKERS", "LIKELY_ETF", "WEIGHT_PCT",
                 "FUND_NET_ASSETS", "SIZE_OK", "VEHICLE_SCORE", "SHARES_HELD", "VALUE_USD",
                 "NET_FLOW_3MO", "NET_FLOW_PCT_OF_ASSETS", "HAS_SHORT_LINE", "REGISTRANT",
                 "SERIES_ID"]
    with pd.ExcelWriter(xlsx, engine="openpyxl") as xw:
        pd.DataFrame({"README": README_TEXT.splitlines()}).to_excel(xw, sheet_name="README", index=False)
        if not snap.empty:
            snap[[c for c in snap_cols if c in snap]].to_excel(xw, sheet_name="Latest_Snapshot", index=False)
        if not rollup.empty:
            rollup.to_excel(xw, sheet_name="Rollup", index=False)
        for tk, p in pivots.items():
            p.to_excel(xw, sheet_name=f"Weights_{tk}"[:31])
        if not g.empty:
            g.to_excel(xw, sheet_name="All_Holdings", index=False)
        if not derivs.empty:
            derivs.to_excel(xw, sheet_name="Derivatives", index=False)
        pd.DataFrame([{"TICKER": t, "CUSIPS_USED": ", ".join(sorted(c))} for t, c in cusips.items()]) \
            .to_excel(xw, sheet_name="CUSIPs", index=False)
        # Light formatting: widths and freeze header
        for ws in xw.book.worksheets:
            ws.freeze_panes = "B2" if ws.title.startswith("Weights_") else "A2"
            for col in ws.columns:
                width = max((len(str(c.value)) for c in col[:200] if c.value is not None), default=8)
                ws.column_dimensions[col[0].column_letter].width = min(max(width + 2, 8), 60)
    return xlsx


# ----------------------------------------------------------------------------
# Main
# ----------------------------------------------------------------------------

def main(argv=None):
    p = argparse.ArgumentParser(description="Find funds/ETFs holding a stock using SEC N-PORT data.")
    p.add_argument("--tickers", nargs="+", required=True, help="Stock ticker(s), e.g. GME AMC")
    p.add_argument("--cusip", nargs="*", default=[],
                   help="Optional TICKER=CUSIP pairs to force matching, e.g. GME=36467W109")
    p.add_argument("--email", required=False, help="Your email for the SEC User-Agent header (required to download)")
    p.add_argument("--start", type=parse_quarter, default=FIRST_QUARTER, help="First data-set quarter, e.g. 2020q4")
    p.add_argument("--end", type=parse_quarter, default=None, help="Last data-set quarter (default: newest)")
    p.add_argument("--latest", action="store_true", help="Only process the newest available quarter")
    p.add_argument("--data-dir", type=Path, default=Path("nport_data"), help="Where zips are cached")
    p.add_argument("--out", type=Path, default=Path("output"), help="Output folder")
    p.add_argument("--etf-only", action="store_true", help="Keep only funds that look like ETFs in outputs")
    p.add_argument("--min-aum", type=float, default=100e6, help="Net assets threshold for SIZE_OK (default $100M)")
    p.add_argument("--top", type=int, default=50, help="Funds per ticker in Latest_Snapshot")
    p.add_argument("--no-derivatives", action="store_true", help="Skip derivative (options/swaps) matching")
    p.add_argument("--delete-zips", action="store_true", help="Delete each zip after processing to save disk")
    p.add_argument("--offline", action="store_true", help="Use only zips already in --data-dir")
    a = p.parse_args(argv)

    tickers = [t.upper() for t in a.tickers]
    if not a.offline and not a.email:
        p.error("--email is required to download from sec.gov (or use --offline with cached zips)")
    a.data_dir.mkdir(parents=True, exist_ok=True)
    sess = session_for(a.email or "unknown@example.com")

    # Quarters to process
    end = a.end or current_quarter()
    quarters = quarter_range(a.start, end)
    if a.offline:
        quarters = [q for q in quarters if (a.data_dir / f"{q}_nport.zip").exists()]

    # CUSIP cache so a ticker's CUSIP is discovered once
    cusip_cache_path = a.data_dir / "cusip_cache.json"
    cusip_cache = json.loads(cusip_cache_path.read_text()) if cusip_cache_path.exists() else {}
    targets = {t: set(cusip_cache.get(t, [])) for t in tickers}
    for pair in a.cusip:
        t, _, c = pair.partition("=")
        if t.upper() in targets and len(c.strip()) == 9:
            targets[t.upper()] = {c.strip().upper()}

    fund_tickers = load_fund_tickers(sess, a.data_dir, offline=a.offline)
    log(f"Loaded tickers for {len(fund_tickers):,} fund series")

    # Newest-first so --latest and CUSIP discovery use the most recent data
    order = list(reversed(quarters))
    directs, derivs = [], []
    processed = 0
    for q in order:
        zp = (a.data_dir / f"{q}_nport.zip") if a.offline else download_quarter(sess, q, a.data_dir)
        if zp is None or not zp.exists():
            continue
        log(f"{q}: scanning for {', '.join(tickers)}")
        try:
            direct, deriv, disc = scan_quarter(zp, targets, include_derivs=not a.no_derivatives)
            # If any ticker had no CUSIP yet, adopt the most common one and rescan once
            newly = []
            for t, counts in disc.items():
                if not targets[t] and len(counts):
                    targets[t] = {counts.index[0]}
                    newly.append(t)
                    log(f"{t}: matched to CUSIP {counts.index[0]} "
                        f"(seen {counts.iloc[0]}x; other candidates: {list(counts.index[1:4])})")
            if newly:
                cusip_cache.update({t: sorted(c) for t, c in targets.items() if c})
                cusip_cache_path.write_text(json.dumps(cusip_cache, indent=2))
                log(f"{q}: rescanning with CUSIP(s) to catch untagged holdings")
                direct, deriv, _ = scan_quarter(zp, targets, include_derivs=not a.no_derivatives)
        except zipfile.BadZipFile:
            log(f"{q}: corrupt zip, deleting; rerun to re-download")
            zp.unlink(missing_ok=True)
            continue
        log(f"{q}: {len(direct):,} direct holding rows, {len(deriv):,} derivative rows")
        if not direct.empty:
            directs.append(direct)
        if not deriv.empty:
            derivs.append(deriv)
        processed += 1
        if a.delete_zips:
            zp.unlink(missing_ok=True)
        if a.latest:
            break

    if processed == 0:
        log("No quarters processed. Check your date range or connection.")
        return 1

    for t in tickers:
        if not targets[t]:
            log(f"WARNING: no holdings found for {t}. Check the ticker, or pass --cusip {t}=XXXXXXXXX")

    direct_all = dedupe_latest(pd.concat(directs, ignore_index=True)) if directs else pd.DataFrame()
    deriv_all = dedupe_latest(pd.concat(derivs, ignore_index=True)) if derivs else pd.DataFrame()

    g = tidy(direct_all, fund_tickers, a.min_aum)
    rollup = period_rollup(g if not a.etf_only else g[g["LIKELY_ETF"]]) if not g.empty else pd.DataFrame()
    snap = latest_snapshot(g, a.top)
    pivots = weight_pivot(g if not a.etf_only else g[g["LIKELY_ETF"]]) if not g.empty else {}
    dv = tidy_derivs(deriv_all, fund_tickers)

    xlsx = write_outputs(a.out, tickers, g, rollup, snap, pivots, dv, targets, a.etf_only)
    log(f"Done. Workbook: {xlsx.resolve()}")

    # Console preview
    if not snap.empty:
        show = snap if not a.etf_only else snap[snap["LIKELY_ETF"]]
        for tk, grp in show.groupby("TICKER"):
            print(f"\nTop vehicle candidates for {tk} (latest reports):")
            cols = ["REPORT_DATE", "FUND_TICKERS", "FUND_NAME", "WEIGHT_PCT", "FUND_NET_ASSETS"]
            prev = grp[cols].head(15).copy()
            prev["FUND_NET_ASSETS"] = (prev["FUND_NET_ASSETS"] / 1e6).round(1).astype(str) + "M"
            prev["WEIGHT_PCT"] = prev["WEIGHT_PCT"].round(2)
            prev["FUND_NAME"] = prev["FUND_NAME"].str.slice(0, 45)
            print(prev.to_string(index=False))
    return 0


if __name__ == "__main__":
    sys.exit(main())
