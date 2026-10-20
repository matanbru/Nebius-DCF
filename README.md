# Nebius DCF – live market inputs

A discounted cash flow valuation of Nebius Group (NASDAQ: NBIS), built in Excel from the
company's 20-F and 6-K filings, with Python scripts that refresh the market inputs and
estimate a regression beta from live data before each valuation run.

The workbook itself isn't in this repository; the scripts find it through a setting in `.env`.

## What you can run

| Script | What it does |
|---|---|
| `update_market_inputs.py` | **Run this one.** Refreshes everything: share price, risk-free rate and the regression beta. |
| `calculate_beta.py` | Refreshes only the regression beta. `update_market_inputs.py` already runs it. |
| `common.py` | Shared code (settings, API calls, opening and saving the workbook). Not run on its own. |

Both scripts update and save the workbook, and open it in Excel if it was not already
open. If the workbook is already open, the open copy is updated – unless it has unsaved
changes, in which case the script stops and asks you to save or discard them first.

## What the scripts update

| Input | Source | Workbook cell |
|---|---|---|
| NBIS share price | [Twelve Data](https://twelvedata.com) `/quote` | `Market_Inputs!B5` |
| US 10-year Treasury yield (risk-free rate) | [FRED](https://fred.stlouisfed.org/series/DGS10) series DGS10 | `Market_Inputs!B18` |
| Regression beta (adjusted) | Twelve Data weekly prices, see below | WACC tab, "NBIS regression beta (weekly, Nov 2024 onwards)" |
| Beta workings | Twelve Data weekly prices | `Beta_Data` tab |

The as-of date is written next to each value. The DCF, WACC, sensitivity tables and
Bear / Base / Bull summary then recalculate in Excel. `Market_Inputs!B20` (the beta the
model uses) is a formula linked to the beta dropdown on the WACC tab, so the scripts never
write to it – choose "Regression" in that dropdown to use the regression beta.

The scripts edit the workbook through Excel itself (via [xlwings](https://www.xlwings.org))
rather than rewriting the file, so formulas, dropdowns and charts stay intact. A backup
with today's date is saved in a `backups` folder next to the workbook before every run.

## How the regression beta is calculated

- Weekly closing prices for NBIS and the S&P 500 from Twelve Data `/time_series`. The
  script tries the index (SPX) first and falls back to the SPY ETF if the API plan doesn't
  include the index; the benchmark used is recorded in the workbook.
- Period: from 1 November 2024 to the latest complete week. NBIS was suspended from
  trading from 2022 until October 2024, and before that it was Yandex – a different
  business – so earlier prices aren't used.
- Raw beta = slope of NBIS weekly returns against S&P 500 weekly returns, with R-squared,
  standard error and a 95% confidence range.
- Adjusted beta = 0.67 × raw beta + 0.33 (the Bloomberg adjustment, which pulls the
  estimate towards the market average of 1).
- The `Beta_Data` tab recalculates the same numbers with Excel formulas (`SLOPE`, `RSQ`,
  `COUNT`, `LINEST`), and the script checks that Excel and Python agree.

## Setup

Requires Windows or macOS with Excel, and Python 3.10+.

```
pip install -r requirements.txt
copy .env.example .env
```

Fill in `.env` with your Twelve Data and FRED API keys and the path to the workbook
(`WORKBOOK_PATH`). The workbook must already exist at that path. `.env` is git-ignored,
so the keys are kept out of Git.

## Run

```
py update_market_inputs.py
```

Example output:

```
Share price:    $170.76  (2026-10-20)
Risk-free rate: 5.09%  (10y Treasury, 2026-10-16)

Regression beta (NBIS vs SPY ETF (S&P 500 proxy), weekly)
  Period:          4 Nov 2024 – 12 Oct 2026  (101 weeks of returns)
  Raw beta:        3.01
  95% range:       1.63 – 4.38  (standard error 0.69)
  R-squared:       0.16
  Adjusted beta:   2.34  -> WACC!B34
  Excel check:     SLOPE() on Beta_Data = 3.0060, matches Python
```

## Inputs still set by hand

- **Equity risk premium** – from Damodaran's website, updated a few times a year.
- **Beta choice** – the dropdown on the WACC tab decides whether the model uses the
  regression beta or a bottom-up (industry) beta.
- **Pre-tax cost of debt** – a judgement based on the secured loan (SOFR + 2.50%) and the
  convertible notes.
