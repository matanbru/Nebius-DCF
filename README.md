# Nebius DCF – live market inputs

A discounted cash flow valuation of Nebius Group (NASDAQ: NBIS), built in Excel from the
company's 20-F and 6-K filings, with a Python script that refreshes the market inputs
from live data before each valuation run.

## What the script does

`update_market_inputs.py` pulls:

| Input | Source | Workbook cell |
|---|---|---|
| NBIS share price | [Twelve Data](https://twelvedata.com) `/quote` | `Market_Inputs!B5` |
| US 10-year Treasury yield (risk-free rate) | [FRED](https://fred.stlouisfed.org/series/DGS10) series DGS10 | `Market_Inputs!B18` |
| Levered beta | Twelve Data `/statistics` (paid plans only) | `Market_Inputs!B20` |

and writes the as-of date next to each value. The DCF, WACC, sensitivity tables and
Bear / Base / Bull summary then recalculate in Excel.

It edits the workbook through Excel itself (via [xlwings](https://www.xlwings.org)) rather
than rewriting the file, so formulas, dropdowns and data tables stay intact. If the
workbook is already open, the open copy is updated. A timestamped backup is saved next
to the workbook before every run.

## Setup

Requires Windows or macOS with Excel, and Python 3.10+.

```
pip install -r requirements.txt
copy .env.example .env
```

Fill in `.env` with your Twelve Data and FRED API keys and the path to the workbook.
`.env` is git-ignored, so the keys never leave your machine.

## Run

```
py update_market_inputs.py
```

Example output:

```
Share price:    $169.99  (2026-10-20)
Beta:           skipped - Twelve Data /statistics returned HTTP 403
Risk-free rate: 5.09%  (10y Treasury, 2026-10-16)
```

## Inputs still set by hand

- **Equity risk premium** – from Damodaran's website, updated a few times a year.
- **Beta** – if the API plan doesn't include it; a peer beta (e.g. CoreWeave) is a
  reasonable alternative given Nebius's short trading history as an AI cloud company.
- **Pre-tax cost of debt** – a judgement based on the secured loan (SOFR + 2.50%) and the
  convertible notes.
