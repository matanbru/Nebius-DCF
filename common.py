"""
Shared helpers for the Nebius DCF scripts: settings from .env, the Twelve Data API,
and opening, backing up and saving the workbook through Excel (xlwings).

The workbook is never saved with openpyxl, which would delete its charts.
"""

import os
import shutil
import sys
from datetime import date
from pathlib import Path

import requests
import xlwings as xw
from dotenv import load_dotenv

FOLDER = Path(__file__).resolve().parent
load_dotenv(FOLDER / ".env")

TWELVE_DATA_KEY = os.getenv("TWELVE_DATA_API_KEY", "").strip()
FRED_KEY = os.getenv("FRED_API_KEY", "").strip()
PLACEHOLDER = "paste_your_key_here"

# Full path to the workbook, or just its file name if it sits in this folder.
WORKBOOK = FOLDER / os.getenv("WORKBOOK_PATH", "Nebius DCF document.xlsx").strip().strip('"')


def require_settings():
    if not TWELVE_DATA_KEY or TWELVE_DATA_KEY == PLACEHOLDER:
        sys.exit("Add your Twelve Data key to the .env file first (TWELVE_DATA_API_KEY=...).")
    if not WORKBOOK.exists():
        sys.exit(f"Can't find the workbook at {WORKBOOK} - check WORKBOOK_PATH in .env")


def twelve_data(endpoint, **params):
    """Call a Twelve Data endpoint. It reports errors in the JSON body, not the HTTP status."""
    params["apikey"] = TWELVE_DATA_KEY
    r = requests.get(f"https://api.twelvedata.com/{endpoint}", params=params, timeout=30)
    if not r.ok:
        # Don't use raise_for_status(): its message includes the URL, and the URL includes the key.
        raise RuntimeError(f"Twelve Data /{endpoint} returned HTTP {r.status_code}")
    data = r.json()
    if data.get("status") == "error":
        raise RuntimeError(data.get("message", "unknown Twelve Data error"))
    return data


def fmt_date(d):
    return d.strftime("%d %b %Y").lstrip("0")


def open_book():
    """Use the workbook if it's already open in Excel, otherwise open it visibly.

    Returns (book, opened_here). Stops if the open copy has unsaved changes, because
    saving it would also save whatever you were in the middle of.
    """
    for app in xw.apps:
        for book in app.books:
            if Path(book.fullname).name == WORKBOOK.name:
                if not book.api.Saved:
                    sys.exit(f"{WORKBOOK.name} is open in Excel with unsaved changes. "
                             "Save or discard them in Excel, then run this again.")
                return book, False
    app = xw.App(visible=True, add_book=False)
    return app.books.open(str(WORKBOOK)), True


def backup_workbook():
    """Copy the workbook to a backups folder next to it, with today's date in the name."""
    backups = WORKBOOK.parent / "backups"
    backups.mkdir(exist_ok=True)
    stamp = date.today().isoformat()
    backup = backups / f"{WORKBOOK.stem} {stamp}{WORKBOOK.suffix}"
    n = 2
    while backup.exists():  # keep every backup if you run it more than once a day
        backup = backups / f"{WORKBOOK.stem} {stamp} ({n}){WORKBOOK.suffix}"
        n += 1
    shutil.copy2(WORKBOOK, backup)
    return backup


def save_book(book, opened_here):
    """Save through Excel and leave the workbook open so you can look at the result."""
    book.save()
    try:
        book.activate()  # bring Excel to the front; nice to have, so don't fail if Windows refuses
    except Exception:
        pass
    if opened_here:
        print("The workbook is now open in Excel.")
    else:
        print("The workbook was open in Excel, so the open copy was updated and saved.")
