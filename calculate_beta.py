"""
Regression beta for Nebius (NBIS) against the S&P 500.

Downloads weekly closing prices from Twelve Data, regresses NBIS weekly returns on
S&P 500 weekly returns, and writes the results into the DCF workbook:
  - WACC tab: the adjusted beta goes in the "NBIS regression beta (weekly, Nov 2024
    onwards)" cell (found by its label), with the source and date next to it.
  - Beta_Data tab: the weekly prices and returns, with live Excel formulas
    (SLOPE, RSQ, COUNT, LINEST) that recalculate the beta, so the calculation is visible.

Only data from 1 Nov 2024 is used. NBIS was suspended from trading from 2022 until
October 2024, and before that the company was Yandex - a different business.

Run on its own:  py calculate_beta.py
update_market_inputs.py also runs this, so one script refreshes everything.
"""

from datetime import date, timedelta

import pandas as pd
from scipy import stats

from common import (WORKBOOK, backup_workbook, fmt_date, open_book, require_settings,
                    save_book, twelve_data)

TICKER = "NBIS"
BENCHMARKS = [("SPX", "S&P 500 index (SPX)"), ("SPY", "SPY ETF (S&P 500 proxy)")]
START = date(2024, 11, 1)

WACC_SHEET = "WACC"
BETA_LABEL = "NBIS regression beta (weekly, Nov 2024 onwards)"
DATA_SHEET = "Beta_Data"

BLUE = (0, 0, 255)
BLACK = (0, 0, 0)


# ---------------------------------------------------------------- data and maths

def weekly_closes(symbol):
    """Weekly closes from START up to the latest complete week, indexed by week start."""
    data = twelve_data("time_series", symbol=symbol, interval="1week",
                       start_date=START.isoformat(), order="ASC", outputsize=5000)
    closes = pd.Series(
        {pd.Timestamp(v["datetime"]): float(v["close"]) for v in data["values"]}, name=symbol)
    # Twelve Data dates each weekly bar by its Monday. A week is complete once its
    # Friday close is in, so drop the current, still-moving week.
    complete = closes.index + timedelta(days=5) <= pd.Timestamp(date.today())
    return closes[complete & (closes.index >= pd.Timestamp(START))]


def benchmark_closes():
    """Try the S&P 500 index first; fall back to SPY if the plan doesn't include it."""
    errors = []
    for symbol, description in BENCHMARKS:
        try:
            return weekly_closes(symbol), symbol, description
        except RuntimeError as e:
            errors.append(f"{symbol}: {e}")
    raise RuntimeError("no S&P 500 data available - " + "; ".join(errors))


def run_regression():
    """Download prices and calculate the beta. Doesn't touch the workbook."""
    nbis = weekly_closes(TICKER)
    bench, bench_symbol, bench_desc = benchmark_closes()

    # Line up by date and keep only weeks present in both series.
    prices = pd.concat([nbis, bench], axis=1, join="inner").dropna()
    prices.columns = ["nbis", "bench"]
    returns = prices.pct_change().dropna()

    x, y = returns["bench"].to_numpy(), returns["nbis"].to_numpy()
    fit = stats.linregress(x, y)
    n = len(x)
    t_crit = stats.t.ppf(0.975, n - 2)
    raw = fit.slope

    return dict(
        prices=prices,
        benchmark=bench_symbol,
        benchmark_desc=bench_desc,
        raw=raw,
        r2=fit.rvalue ** 2,
        se=fit.stderr,
        ci=(raw - t_crit * fit.stderr, raw + t_crit * fit.stderr),
        weeks=n,
        adjusted=0.67 * raw + 0.33,
        first=prices.index[0],
        last=prices.index[-1],
    )


# ---------------------------------------------------------------- workbook

def find_label(sheet, label):
    """Find the cell containing exactly this label text."""
    found = sheet.api.UsedRange.Find(What=label, LookAt=1)  # 1 = xlWhole
    if found is None:
        raise RuntimeError(f"can't find '{label}' on the {sheet.name} tab")
    return sheet.range(found.Address)


def write_beta_data(book, res):
    """Create or refresh the Beta_Data tab: prices, returns and the Excel regression."""
    if DATA_SHEET in [s.name for s in book.sheets]:
        ws = book.sheets[DATA_SHEET]
        ws.clear()  # keep the sheet itself, so any links to it keep working
    else:
        ws = book.sheets.add(DATA_SHEET, after=book.sheets[WACC_SHEET])

    prices = res["prices"]
    n_rows = len(prices)
    first, last = 5, 5 + n_rows - 1  # data rows

    ws["A1"].value = "NBIS regression beta – weekly returns vs S&P 500"
    ws["A2"].value = ("Blue = data downloaded from Twelve Data by calculate_beta.py; "
                      "black = Excel formulas. Weeks run Monday to Friday; Date = week start.")
    ws["A4"].value = ["Date", "NBIS close ($)", f"S&P close – {res['benchmark']} ($)",
                      "NBIS weekly return", "S&P weekly return"]

    ws[f"A{first}"].value = [[d.to_pydatetime(), a, b] for d, a, b in
                             zip(prices.index, prices["nbis"], prices["bench"])]
    # Returns as formulas: this week's close / last week's close - 1.
    ret_formulas = [[f"=B{r}/B{r-1}-1", f"=C{r}/C{r-1}-1"] for r in range(first + 1, last + 1)]
    ws[f"D{first + 1}"].value = ret_formulas

    y = f"D{first + 1}:D{last}"
    x = f"E{first + 1}:E{last}"
    s = last + 2  # summary block starts two rows below the data
    summary = [
        ("Regression results (Excel formulas)", None),
        ("Benchmark", res["benchmark_desc"]),
        ("Date range (week starting)",
         f"{fmt_date(res['first'])} – {fmt_date(res['last'])}"),
        ("Weeks of returns used", f"=COUNT({y})"),
        ("Raw beta = SLOPE(NBIS returns, S&P returns)", f"=SLOPE({y},{x})"),
        ("R-squared", f"=RSQ({y},{x})"),
        ("Standard error of beta", f"=INDEX(LINEST({y},{x},TRUE,TRUE),2,1)"),
        ("95% range – low", f"=B{s+4}-T.INV.2T(0.05,B{s+3}-2)*B{s+6}"),
        ("95% range – high", f"=B{s+4}+T.INV.2T(0.05,B{s+3}-2)*B{s+6}"),
        ("Adjusted beta = 0.67 × raw + 0.33 (Bloomberg)", f"=0.67*B{s+4}+0.33"),
    ]
    ws[f"A{s}"].value = [list(row) for row in summary]

    # Formatting: Arial throughout, blue data, black formulas.
    ws.used_range.font.name = "Arial"
    ws.used_range.font.color = BLACK
    ws["A1"].font.bold = True
    ws["A4:E4"].font.bold = True
    ws[f"A{s}"].font.bold = True
    ws[f"A{first}:C{last}"].font.color = BLUE
    ws[f"B{s+1}:B{s+2}"].font.color = BLUE  # benchmark and dates are typed-in text
    ws[f"A{first}:A{last}"].number_format = "d mmm yyyy"
    ws[f"B{first}:C{last}"].number_format = "#,##0.00"
    ws[f"D{first}:E{last}"].number_format = "0.0%"
    ws[f"B{s+3}"].number_format = "0"
    ws[f"B{s+4}:B{s+9}"].number_format = "0.00"
    ws[f"B{s+5}"].number_format = "0.0%"
    ws.range("A:A").column_width = 44
    ws.range("B:E").column_width = 20
    ws["A2"].wrap_text = False

    return ws, s


def write_beta(book, res):
    """Write the regression results into the workbook and check Excel agrees with Python."""
    ws, s = write_beta_data(book, res)

    excel_raw = ws[f"B{s+4}"].value
    excel_weeks = ws[f"B{s+3}"].value
    matches = abs(excel_raw - res["raw"]) < 1e-9 and int(excel_weeks) == res["weeks"]

    wacc = book.sheets[WACC_SHEET]
    cell = find_label(wacc, BETA_LABEL).offset(0, 1)  # value column, as for the other betas
    cell.value = round(res["adjusted"], 4)
    cell.offset(0, 1).value = (
        f"Twelve Data weekly closes, NBIS vs {res['benchmark_desc']}, "
        f"{fmt_date(res['first'])} – {fmt_date(res['last'])} ({res['weeks']} weeks). "
        f"Raw beta {res['raw']:.2f}, R² {res['r2']:.2f}; adjusted = 0.67 × raw + 0.33. "
        f"Updated {fmt_date(date.today())}. Workings on {DATA_SHEET} tab."
    )

    return dict(cell=cell.address.replace("$", ""), excel_raw=excel_raw, matches=matches,
                bottom_up_tax=wacc["B46"].value, bottom_up_no_tax=wacc["B47"].value,
                choice=wacc["B49"].value)


def print_summary(res, check):
    lo, hi = res["ci"]
    print(f"Regression beta (NBIS vs {res['benchmark_desc']}, weekly)")
    print(f"  Period:          {fmt_date(res['first'])} – {fmt_date(res['last'])}"
          f"  ({res['weeks']} weeks of returns)")
    print(f"  Raw beta:        {res['raw']:.2f}")
    print(f"  95% range:       {lo:.2f} – {hi:.2f}  (standard error {res['se']:.2f})")
    print(f"  R-squared:       {res['r2']:.2f}")
    print(f"  Adjusted beta:   {res['adjusted']:.2f}  -> WACC!{check['cell']}")
    if check["matches"]:
        print(f"  Excel check:     SLOPE() on {DATA_SHEET} = {check['excel_raw']:.4f}, matches Python")
    else:
        print(f"  Excel check:     MISMATCH - Excel SLOPE() = {check['excel_raw']:.4f}, "
              f"Python = {res['raw']:.4f}. Look at the {DATA_SHEET} tab before using this beta.")
    print(f"  Bottom-up betas: {check['bottom_up_tax']:.2f} (with tax), "
          f"{check['bottom_up_no_tax']:.2f} (0% tax). Beta choice on WACC tab: {check['choice']}")


def main():
    require_settings()
    res = run_regression()  # download and calculate before touching the workbook
    book, opened_here = open_book()
    backup = backup_workbook()
    try:
        check = write_beta(book, res)
        save_book(book, opened_here)
    except Exception:
        if opened_here:
            book.app.quit()  # close the copy we opened without saving the half-done changes
        raise
    print_summary(res, check)
    print(f"Backup saved to {backup}")


if __name__ == "__main__":
    main()
