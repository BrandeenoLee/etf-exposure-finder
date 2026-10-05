"""
Builds small fake N-PORT quarterly zips that mirror the SEC's real file layout
(tab-delimited tables inside a folder in a zip), so the pipeline can be tested
without downloading ~450 MB per quarter.

Scenarios covered:
  * a holding tagged with the ticker (normal case)
  * a holding with no ticker, only the CUSIP (must still be found)
  * a ticker written as "GME US" (filer formatting quirk)
  * an amended filing that should replace the original
  * a single-stock total return swap referencing the stock
  * a basket swap that includes the stock as one component
  * an unrelated holding that must be ignored
"""

import zipfile
from pathlib import Path

GME_CUSIP = "36467W109"


def _tsv(cols, rows):
    return "\t".join(cols) + "\n" + "".join("\t".join(map(str, r)) + "\n" for r in rows)


def _holding(acc, hid, name, cusip, bal, val, pct, payoff="Long", cat="EC", dcat=""):
    return [acc, hid, name, "", "Common", cusip, bal, "NS", "", "USD", val, "", pct,
            payoff, cat, "", "CORP", "", "US", "N", "1", dcat]


def build_quarter(folder, q, report_date, filing_date, xrt_weight, amend=False):
    folder = Path(folder)
    accs = ["A1", "A2", "A3", "A4"] + (["A1B"] if amend else [])

    subs = [[a, filing_date, "", "NPORT-P", "", report_date, "N"] for a in accs[:4]]
    if amend:
        subs.append(["A1B", "2099-01-01", "", "NPORT-P/A", "", report_date, "N"])

    reg_names = {"A1": ("111", "SPDR Series Trust"), "A2": ("222", "Vanguard Index Funds"),
                 "A3": ("333", "Tidal Trust"), "A4": ("444", "Some Mutual Fund Trust"),
                 "A1B": ("111", "SPDR Series Trust")}
    regs = [[a, *reg_names[a], "", ""] for a in accs]

    fund = {"A1": ("SPDR S&P Retail ETF", "S000001", 500e6),
            "A2": ("Vanguard Extended Market Index Fund", "S000002", 90e9),
            "A3": ("T-Rex 2X Long GME Daily Target ETF", "S000003", 20e6),
            "A4": ("Growth Opportunities Fund", "S000004", 800e6),
            "A1B": ("SPDR S&P Retail ETF", "S000001", 500e6)}
    fris = [[a, fund[a][0], fund[a][1], "", fund[a][2], 10, 10, 10, 200, 5, 5] for a in accs]

    holds = [
        _holding("A1", "1", "GameStop Corp", GME_CUSIP, 100000, 2500000, xrt_weight),
        _holding("A1", "2", "Macy's", "55616P104", 100000, 2000000, 1.4),
        _holding("A2", "3", "GAMESTOP CORP-A", GME_CUSIP, 5000000, 125000000, 0.14),
        _holding("A3", "4", "Total return swap GME", "N/A", 0, 15000000, 75, "Long", "DE", "SWP"),
        _holding("A4", "5", "GameStop Corp", GME_CUSIP, 300000, 7500000, 0.9),
        _holding("A4", "6", "Custom basket swap", "N/A", 0, 100000, 0.01, "Long", "DE", "SWP"),
    ]
    ids = [["1", "1", "US36467W1099", "GME", "", ""], ["2", "2", "", "M", "", ""],
           ["5", "3", "", "GME US", "", ""]]
    if amend:
        holds.append(_holding("A1B", "7", "GameStop Corp", GME_CUSIP, 120000, 3000000, xrt_weight + 0.5))
        ids.append(["7", "4", "", "GME", "", ""])

    tables = {
        "SUBMISSION": (["ACCESSION_NUMBER", "FILING_DATE", "FILE_NUM", "SUB_TYPE",
                        "REPORT_ENDING_PERIOD", "REPORT_DATE", "IS_LAST_FILING"], subs),
        "REGISTRANT": (["ACCESSION_NUMBER", "CIK", "REGISTRANT_NAME", "FILE_NUM", "LEI"], regs),
        "FUND_REPORTED_INFO": (["ACCESSION_NUMBER", "SERIES_NAME", "SERIES_ID", "SERIES_LEI", "NET_ASSETS",
                                "SALES_FLOW_MON1", "SALES_FLOW_MON2", "SALES_FLOW_MON3",
                                "REDEMPTION_FLOW_MON1", "REDEMPTION_FLOW_MON2", "REDEMPTION_FLOW_MON3"], fris),
        "FUND_REPORTED_HOLDING": (["ACCESSION_NUMBER", "HOLDING_ID", "ISSUER_NAME", "ISSUER_LEI", "ISSUER_TITLE",
                                   "ISSUER_CUSIP", "BALANCE", "UNIT", "OTHER_UNIT_DESC", "CURRENCY_CODE",
                                   "CURRENCY_VALUE", "EXCHANGE_RATE", "PERCENTAGE", "PAYOFF_PROFILE",
                                   "ASSET_CAT", "OTHER_ASSET", "ISSUER_TYPE", "OTHER_ISSUER",
                                   "INVESTMENT_COUNTRY", "IS_RESTRICTED_SECURITY", "FAIR_VALUE_LEVEL",
                                   "DERIVATIVE_CAT"], holds),
        "IDENTIFIERS": (["HOLDING_ID", "IDENTIFIERS_ID", "IDENTIFIER_ISIN", "IDENTIFIER_TICKER",
                         "OTHER_IDENTIFIER", "OTHER_IDENTIFIER_DESC"], ids),
        "DESC_REF_OTHER": (["HOLDING_ID", "DESC_REF_OTHER_ID", "ISSUER_NAME", "ISSUE_TITLE", "CUSIP",
                            "ISIN", "TICKER", "OTHER_IDENTIFIER", "OTHER_DESC"],
                           [["4", "1", "GameStop Corp", "Common", GME_CUSIP, "", "GME", "", ""]]),
        "DESC_REF_INDEX_COMPONENT": (["HOLDING_ID", "DESC_REF_INDEX_COMPONENT_ID", "NAME", "CUSIP", "ISIN",
                                      "TICKER", "OTHER_IDENTIFIER", "OTHER_DESC", "NOTIONAL_AMOUNT",
                                      "CURRENCY_CODE", "VALUE", "ISSUER_CURRENCY_CODE"],
                                     [["6", "1", "GameStop", GME_CUSIP, "", "", "", "", "50000",
                                       "USD", "51000", "USD"]]),
    }
    path = folder / f"{q}_nport.zip"
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as z:
        for name, (cols, rows) in tables.items():
            z.writestr(f"{q}_nport/{name}.tsv", _tsv(cols, rows))
    return path


def build_all(folder):
    build_quarter(folder, "2020q4", "2020-12-31", "2021-02-20", 1.5)
    build_quarter(folder, "2021q1", "2021-03-31", "2021-05-20", 9.8, amend=True)


if __name__ == "__main__":
    build_all(".")
    print("Wrote 2020q4_nport.zip and 2021q1_nport.zip")
