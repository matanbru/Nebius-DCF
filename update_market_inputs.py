"""
Update the live market inputs in the Nebius DCF workbook - one script refreshes everything.

  - NBIS share price      (Twelve Data /quote)  -> Market_Inputs!B5
  - US 10-year Treasury   (FRED series DGS10)   -> Market_Inputs!B18  (if FRED_API_KEY is set)
  - Regression beta       (calculate_beta.py)   -> WACC tab beta cell and the Beta_Data tab
and writes today's date in column C next to B5 and B18.

Market_Inputs!B20 (beta) is a formula linked to the beta dropdown on the WACC tab, so
this script never writes to it.

The workbook is edited through Excel itself (xlwings), so formulas, dropdowns and charts
are left untouched. If the workbook is already open in Excel, the script updates the open
copy - unless it has unsaved changes, in which case it stops. Otherwise it opens the
workbook visibly and leaves it open after saving. A backup copy with today's date is saved
in a backups folder next to the workbook before anything changes.

API keys and the workbook location are read from a .env file in this folder (see
.env.example) - never type them into this file.

Run:  py update_market_inputs.py
"""

from datetime import date, datetime

import requests

import calculate_beta
from common import (FRED_KEY, PLACEHOLDER, WORKBOOK, backup_workbook, fmt_date, open_book,
                    require_settings, save_book, twelve_data)

SHEET = "Market_Inputs"
PRICE_CELL = "B5"
RISK_FREE_CELL = "B18"


def get_price():
    quote = twelve_data("quote", symbol=calculate_beta.TICKER)
    # During market hours "close" is the latest price; after hours it is the day's close.
    return float(quote["close"]), quote.get("datetime", "")


def get_risk_free():
    r = requests.get(
        "https://api.stlouisfed.org/fred/series/observations",
        params=dict(series_id="DGS10", api_key=FRED_KEY, file_type="json",
                    sort_order="desc", limit=10),
        timeout=20,
    )
    if not r.ok:
        raise RuntimeError(f"FRED returned HTTP {r.status_code}")
    # FRED marks market holidays with "."; take the most recent real observation.
    for obs in r.json()["observations"]:
        if obs["value"] != ".":
            return float(obs["value"]) / 100, obs["date"]
    raise RuntimeError("no recent DGS10 observation")


def main():
    require_settings()

    # Download everything first, so a failed download doesn't leave a half-updated workbook.
    updates = {}  # cell -> (value, as-of text)

    price, price_time = get_price()
    updates[PRICE_CELL] = (price, fmt_date(date.today()))
    print(f"Share price:    ${price:,.2f}  ({price_time})")

    if FRED_KEY and FRED_KEY != PLACEHOLDER:
        rf, rf_date = get_risk_free()
        updates[RISK_FREE_CELL] = (rf, fmt_date(datetime.strptime(rf_date, "%Y-%m-%d")))
        print(f"Risk-free rate: {rf:.2%}  (10y Treasury, {rf_date})")
    else:
        print("Risk-free rate: skipped - add FRED_API_KEY to .env to update it")

    try:
        beta = calculate_beta.run_regression()
    except Exception as e:
        beta = None
        print(f"Beta:           skipped - {e}")

    book, opened_here = open_book()
    backup = backup_workbook()
    try:
        ws = book.sheets[SHEET]
        for cell, (value, as_of) in updates.items():
            ws[cell].value = value
            # Column C = "As of". The apostrophe stores it as text, as the other rows are;
            # otherwise Excel turns it into a date serial and the Summary heading shows "46315".
            ws[cell].offset(0, 1).value = "'" + as_of
        check = calculate_beta.write_beta(book, beta) if beta else None
        save_book(book, opened_here)
    except Exception:
        if opened_here:
            book.app.quit()  # close the copy we opened without saving the half-done changes
        raise

    print(f"\nUpdated {WORKBOOK.name}. Backup saved to backups\\{backup.name}\n")
    if beta:
        calculate_beta.print_summary(beta, check)


if __name__ == "__main__":
    main()
