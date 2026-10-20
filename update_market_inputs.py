"""
Update the live market inputs in the Nebius DCF workbook.

Pulls:
  - NBIS share price      (Twelve Data /quote)       -> Market_Inputs!B5
  - US 10-year Treasury   (FRED series DGS10)        -> Market_Inputs!B18
  - NBIS beta             (Twelve Data /statistics)  -> Market_Inputs!B20
and writes today's date in column C next to each value.

The workbook is edited through Excel itself (xlwings), so formulas, the scenario
dropdown and the sensitivity tables are left untouched. If the workbook is already
open in Excel, the script updates the open copy; otherwise it opens, saves and
closes it. A backup copy is saved in a backups folder next to the workbook before
anything changes.

API keys and the workbook location are read from a .env file in this folder (see
.env.example) - never type them into this file.

Run:  py update_market_inputs.py
"""

import os
import shutil
import sys
from datetime import date, datetime
from pathlib import Path

import requests
import xlwings as xw
from dotenv import load_dotenv

FOLDER = Path(__file__).resolve().parent
load_dotenv(FOLDER / ".env")

# Full path to the workbook, or just its file name if it sits in this folder.
WORKBOOK = FOLDER / os.getenv("WORKBOOK_PATH", "Nebius DCF document.xlsx").strip().strip('"')
SHEET = "Market_Inputs"
TICKER = "NBIS"

PRICE_CELL = "B5"
RISK_FREE_CELL = "B18"
BETA_CELL = "B20"

TWELVE_DATA_KEY = os.getenv("TWELVE_DATA_API_KEY", "").strip()
FRED_KEY = os.getenv("FRED_API_KEY", "").strip()


def twelve_data(endpoint, **params):
    """Call a Twelve Data endpoint. It reports errors in the JSON body, not the HTTP status."""
    params.update(symbol=TICKER, apikey=TWELVE_DATA_KEY)
    r = requests.get(f"https://api.twelvedata.com/{endpoint}", params=params, timeout=20)
    if not r.ok:
        # Don't use raise_for_status(): its message includes the URL, and the URL includes the key.
        raise RuntimeError(f"Twelve Data /{endpoint} returned HTTP {r.status_code}")
    data = r.json()
    if data.get("status") == "error":
        raise RuntimeError(data.get("message", "unknown Twelve Data error"))
    return data


def get_price():
    quote = twelve_data("quote")
    # During market hours "close" is the latest price; after hours it is the day's close.
    return float(quote["close"]), quote.get("datetime", "")


def get_beta():
    stats = twelve_data("statistics")
    beta = stats["statistics"]["stock_price_summary"].get("beta")
    if beta in (None, ""):
        raise RuntimeError("Twelve Data returned no beta for NBIS")
    return float(beta)


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


def fmt_date(d):
    return d.strftime("%d %b %Y").lstrip("0")


def open_book():
    """Use the workbook if it's already open in Excel, otherwise open it in a hidden Excel."""
    for app in xw.apps:
        for book in app.books:
            if Path(book.fullname).name == WORKBOOK.name:
                return book, False
    app = xw.App(visible=False, add_book=False)
    return app.books.open(str(WORKBOOK)), True


def main():
    if not TWELVE_DATA_KEY or TWELVE_DATA_KEY == "paste_your_key_here":
        sys.exit("Add your Twelve Data key to the .env file first (TWELVE_DATA_API_KEY=...).")
    if not WORKBOOK.exists():
        sys.exit(f"Can't find the workbook at {WORKBOOK} - check WORKBOOK_PATH in .env")

    updates = {}  # cell -> (value, as-of text)
    today = fmt_date(date.today())

    price, price_time = get_price()
    updates[PRICE_CELL] = (price, today)
    print(f"Share price:    ${price:,.2f}  ({price_time})")

    try:
        beta = get_beta()
        updates[BETA_CELL] = (beta, today)
        print(f"Beta:           {beta:.2f}")
    except Exception as e:
        print(f"Beta:           skipped - {e}")
        print("                (/statistics may need a paid Twelve Data plan; enter beta by hand)")

    if FRED_KEY and FRED_KEY != "paste_your_key_here":
        rf, rf_date = get_risk_free()
        updates[RISK_FREE_CELL] = (rf, fmt_date(datetime.strptime(rf_date, "%Y-%m-%d")))
        print(f"Risk-free rate: {rf:.2%}  (10y Treasury, {rf_date})")
    else:
        print("Risk-free rate: skipped - add FRED_API_KEY to .env to update it")

    backups = WORKBOOK.parent / "backups"
    backups.mkdir(exist_ok=True)
    backup = backups / f"{WORKBOOK.stem} {datetime.now():%Y-%m-%d %H%M}{WORKBOOK.suffix}"
    shutil.copy2(WORKBOOK, backup)

    book, opened_here = open_book()
    try:
        ws = book.sheets[SHEET]
        for cell, (value, as_of) in updates.items():
            ws[cell].value = value
            # Column C = "As of". The apostrophe stores it as text, as the other rows are;
            # otherwise Excel turns it into a date serial and the Summary heading shows "46315".
            ws[cell].offset(0, 1).value = "'" + as_of
        book.save()
    finally:
        if opened_here:
            book.app.quit()

    print(f"\nUpdated {WORKBOOK.name} ({SHEET}). Backup saved to backups\\{backup.name}")
    if not opened_here:
        print("The workbook was open in Excel, so the open copy was updated and saved.")


if __name__ == "__main__":
    main()
