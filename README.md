# ETF Exposure Finder

![tests](../../actions/workflows/tests.yml/badge.svg)

A Python tool that finds every SEC-registered fund holding a given stock, tracks how that stock's weight in each fund changes over time, and flags which ETFs could serve as an indirect shorting vehicle.

Built with pandas on the SEC's free Form N-PORT bulk data (every registered fund's portfolio holdings, Oct 2019 to present). Works for any US-listed ticker.

## Why I built this

During the January 2021 GameStop squeeze, GME's weight in the SPDR S&P Retail ETF (XRT) jumped from about 1.5% to nearly 20%, and the SEC's staff report on the episode noted that shorting XRT could have served as an indirect way to short GME. A popular theory holds that short sellers use ETFs this way to keep exposure out of a stock's reported short interest.

I wanted to test that theory with data instead of opinion. The first step is knowing which ETFs could plausibly be used this way at any point in time. A fund only works as a vehicle if the stock is a meaningful share of it *and* the fund is big enough to short in size. This tool answers that question for any stock, across nearly seven years of filings.

## What it does

1. **Downloads and caches** the SEC's quarterly N-PORT data sets (~400 MB each, streamed and read directly from the zip, never fully extracted).
2. **Finds every holding** of the target stock across every fund that files N-PORT, matching by ticker and by CUSIP. It auto-detects the CUSIP, so holdings that filers tagged inconsistently (or not at all) are still caught.
3. **Separates derivative exposure** (swaps, options, and basket swaps that reference the stock) from direct share holdings. This matters because leveraged single-stock ETFs hold swaps, not shares.
4. **Cleans the data**: keeps only the latest version of each filing when funds file amendments, nets out short positions, and joins in fund names, net assets, flows, and each fund's own ticker.
5. **Scores and ranks** each fund as a potential vehicle (weight × log of fund size) and outputs an Excel workbook.

## Example output

Using the 2021 squeeze window, the `Weights_GME` sheet shows each fund's GME weight by month, which makes surges like XRT's easy to spot:

| Fund | 2020-12 | 2021-03 | ... |
|---|---|---|---|
| SPDR S&P Retail ETF [XRT] | 1.50 | 10.30 | |
| ... | | | |

*(Illustrative values from the test data. Run it yourself for real figures.)*

The workbook also includes a ranked **Latest_Snapshot**, a monthly **Rollup** (fund count, total shares held by funds), the full **All_Holdings** table, and a **Derivatives** sheet.

## Quick start

Requires Python 3.9+.

```bash
git clone https://github.com/<BrandeenoLee>/etf-exposure-finder.git
cd etf-exposure-finder
pip install -r requirements.txt

# Newest quarter only (one ~450 MB download)
python etf_exposure.py --tickers GME --latest --email you@example.com

# The 2021 squeeze window, two stocks at once
python etf_exposure.py --tickers GME AMC --start 2020q4 --end 2021q3 --email you@example.com
```

The SEC requires a contact email in the User-Agent for automated downloads; it is sent only to sec.gov.

### Options

| Option | What it does |
|---|---|
| `--tickers GME AMC` | One or more stocks, scanned in a single pass |
| `--start 2020q4 --end 2021q3` | Data-set quarters to process |
| `--latest` | Only the newest published quarter |
| `--etf-only` | Keep only funds that look like ETFs |
| `--min-aum 100e6` | Fund-size threshold for the `SIZE_OK` flag |
| `--cusip GME=36467W109` | Force a CUSIP instead of auto-detecting |
| `--no-derivatives` | Skip swaps/options |
| `--delete-zips` | Delete each zip after use to save disk |
| `--offline` | Use only already-downloaded data |

## Tests

```bash
python -m pytest
```

The tests build small synthetic N-PORT files that mirror the SEC's real format and cover the tricky cases: holdings with no ticker tag, filer formatting quirks (`"GME US"`), amended filings, swaps vs. shares, and the ETF filter. They run automatically on every push via GitHub Actions.

## Design notes

- **Memory:** the holdings table is several gigabytes uncompressed. The tool streams it in 500k-row chunks and keeps only matching rows, so a laptop handles it fine.
- **CUSIP discovery:** filers don't always report tickers. The tool learns the stock's CUSIP from tagged holdings, caches it, and rescans to catch untagged ones.
- **Amendments:** when a fund refiles, only the latest filing per fund per report date is kept.

## Limitations

- N-PORT data lags reality by roughly 3–6 months.
- The ETF flag is a name-based heuristic; Vanguard ETFs (share classes of index mutual funds) aren't flagged.
- Holding a stock doesn't mean a fund is being used to short it. This identifies candidates; testing the theory requires each ETF's short interest, failures-to-deliver, and creation/redemption activity, which is the planned next stage.

## Data source

[SEC Form N-PORT Data Sets](https://www.sec.gov/data-research/sec-markets-data/form-n-port-data-sets) (public domain).

This project is for research and education, not investment advice.

## License

MIT
