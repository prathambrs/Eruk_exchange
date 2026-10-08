"""
ERUK (exchangerates.org.uk) daily GBP rates, converted to USD base -> Excel in the company template format

Workbook (eruk_exchange.xlsx), updated in place on every run:
All sheets use BASE (USD): 1 USD = x currency, worked out daily as (GBP/CCY) / (GBP/USD).
    Average Currency_PL          subs currencies only, monthly averages,
                                 one table per year (a new table appears each January)
    Closing Rate_CASH Position   month-end close, 18 template currencies, USD first column = 1
    Daily Rate GESCO             latest rate, Banque de France currency list
    ERUK History                 every daily rate for the same list, Date | Currency | ISO_Code | Rate_USD

Edits made in Excel are kept: highlights, colours, comments, column widths and extra sheets stay,
and the script only writes rate values and new rows/tables (typed-over rates go back to ERUK's
value on the next run). Charts, images and pivot tables are not
supported by openpyxl, so keep those in a separate workbook. Don't insert or delete rows or columns
inside the generated tables, and close the file in Excel before running.

History is also kept in eruk_history.csv next to this script. After the first run the script only
loads the 180-day page per currency and fills in older gaps from the yearly pages when needed.

Setup (once):
    pip install playwright pandas openpyxl
    playwright install chromium

Run:
    python eruk.py
"""
import calendar
import os
import re
import sys
import time
from datetime import date, datetime, timedelta
from pathlib import Path

import pandas as pd
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, PatternFill
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.table import Table, TableStyleInfo
from playwright.sync_api import sync_playwright

# ---------------- settings ----------------
BASE_URL = os.environ.get("ERUK_BASE", "https://www.exchangerates.org.uk")
TODAY = date.fromisoformat(os.environ["ERUK_TODAY"]) if os.environ.get("ERUK_TODAY") else date.today()
START_YEAR = 2026                        # first year in the workbook; set earlier to backfill
ONLY_COMPLETE_MONTHS = True              # running month stays out of averages and closes
HEADLESS = os.environ.get("ERUK_HEADLESS", "0") == "1"   # visible browser by default
HERE = Path(__file__).resolve().parent
OUTPUT = HERE / "eruk_exchange.xlsx"
HISTORY_CSV = HERE / "eruk_history.csv"
DEBUG_DIR = HERE / "debug"
PROFILE_DIR = HERE / "browser_profile"   # keeps Cloudflare cookies between runs
WAIT_FOR_YOU = 120                       # seconds to tick "Verify you are human" when challenged
WAIT_NORMAL = 15                         # seconds to wait for the table on a normal page
PAUSE_BETWEEN_PAGES = 3                  # raise this if Cloudflare keeps re-challenging
MAX_GAP_DAYS = 7                         # a bigger hole in a year's history triggers the yearly page
BASE = "USD"                             # base currency for every sheet (1 USD = x); any listed code works
# ------------------------------------------

# Subsidiaries and their currency, same as the template
SUBS = [
    ("ASIA SHIPPING", "USD"), ("LONDON", "GBP"), ("FUTURES", "GBP"), ("BRSI", "EUR"),
    ("COPENHAGEN", "DKK"), ("AXS MARINE", "EUR"), ("PARIS", "EUR"), ("SEABORNE", "CHF"),
    ("GENEVA", "CHF"), ("ATHENS", "EUR"), ("DUBAI", "USD"), ("HAMBURG", "EUR"),
    ("BRS SHIPPING", "USD"), ("SINGAPORE", "USD"), ("BAXI", "SGD"), ("HONG KONG", "USD"),
    ("BAILI BEIJING", "RMB"), ("BAILI SHANGHAI", "RMB"), ("SINO", "USD"), ("SOFIA", "BGN"),
    ("USA", "USD"), ("AMAZONAS", "COP"), ("Indian Rupee", "INR"), ("UAE Dirham", "AED"),
    ("Indonesian Rupee", "IDR"),
]
TO_ISO = {"RMB": "CNY"}                  # subs list says RMB, ERUK and the ISO code say CNY
iso = lambda code: TO_ISO.get(code, code)   # RMB and CNY are the same currency, both pull ERUK's CNY page

# 1) Average Currency_PL: only the subs currencies, in subs order, no repeats
# labels as shown: template order first, then INR/AED/IDR, then the "Add" list in its order
# (RMB and CNY are the same rate but both kept as asked; TND and BGN already in the template rows)
AVG_TEMPLATE = ["EUR", "USD", "GBP", "CHF", "SGD", "RMB", "NOK", "COP", "BGN", "TND", "DKK"]
AVG_ADD = ["TND", "CNY", "MYR", "VND", "MKD", "XOF", "HKD", "BGN"]
AVG_ROWS = list(dict.fromkeys(AVG_TEMPLATE + ["INR", "AED", "IDR"] + AVG_ADD))

# 2) Closing Rate_CASH Position: the template's 18 currencies with the base moved to the first column (= 1)
TEMPLATE_CLOSING = ["EUR", "USD", "CNY", "GBP", "CHF", "SGD", "AED", "HKD", "BGN", "NOK",
                    "CAD", "IDR", "COP", "VND", "XOF", "DKK", "MKD", "TND"]
CLOSING = [BASE] + [c for c in TEMPLATE_CLOSING if c != BASE]

# 3) + 4) Daily Rate GESCO and ERUK History: the Banque de France list sorted by ISO code (as in the
#         file), then the "Add" list in the order it was given. GESCO rows are written in this order.
DAILY = [
    ("AUD", "Australian dollar"), ("BRL", "Brasilian Reais"), ("CAD", "Canadian dollar"),
    ("CHF", "Swiss franc"), ("CNY", "Chinese yuan renminbi"), ("CZK", "Czeck koruna"),
    ("DKK", "Danish krone"), ("GBP", "Pound sterling"), ("HKD", "Hong Kong dollar"),
    ("HUF", "Hungarian forint"), ("IDR", "Indonesian rupiah"), ("ILS", "New sheqalim"),
    ("INR", "Indian Rupee"), ("JPY", "Japanese yen"), ("KRW", "Korean won"),
    ("MXN", "Mexican Peso"), ("MYR", "Malaysian ringgit"), ("NOK", "Norwegian krone"),
    ("NZD", "New Zealand dollar"), ("PHP", "Philippine peso"), ("PLN", "Polish zloty"),
    ("RON", "New romanian leu"), ("SEK", "Swedish krona"), ("SGD", "Singaporean dollar"),
    ("THB", "Thai baht"), ("TRY", "New turkish lira"), ("USD", "US dollar"),
    ("ZAR", "South African rand"),
    # "Add" list, in its order (RMB is the same rate as CNY, kept as asked; BGN listed once)
    ("COP", "Colombian peso"), ("BGN", "Bulgarian lev"), ("EUR", "Euro"), ("RMB", "Chinese yuan renminbi"),
    ("AED", "UAE dirham"), ("TND", "Tunisian dinar"), ("VND", "Vietnamese dong"),
    ("MKD", "Macedonian denar"), ("XOF", "West African CFA franc"),
]
NAMES = dict(DAILY)

# everything any sheet needs (scraped once)
CURRENCIES = list(dict.fromkeys(CLOSING + [iso(c) for c in AVG_ROWS] + [iso(c) for c, _ in DAILY]))
PEGGED_TO_EUR = {}                       # no fixed-rate calculations: every currency is pulled from ERUK

MONTHS = "January|February|March|April|May|June|July|August|September|October|November|December"
MONTH_NAMES = MONTHS.split("|")
WEEKDAY_START = re.compile(r"^\s*(Monday|Tuesday|Wednesday|Thursday|Friday|Saturday|Sunday)\b", re.I)
SLASH_DATE = re.compile(r"\b(\d{1,2})/(\d{1,2})/(\d{4})\b")                           # 02/10/2026
TEXT_DATE = re.compile(rf"\b(\d{{1,2}})(?:st|nd|rd|th)?\s+({MONTHS})\s+(\d{{4}})\b")  # 2 October 2026


# ---------------- scraping ----------------
def parse_date(text):
    try:
        m = TEXT_DATE.search(text)
        if m:
            d, mo, y = m.groups()
            return datetime.strptime(f"{d} {mo} {y}", "%d %B %Y").date()
        m = SLASH_DATE.search(text)
        if m:
            d, mo, y = map(int, m.groups())
            return date(y, mo, d)
    except ValueError:
        pass
    return None


def rate_regex(code):
    # "1 GBP = 1.3456 USD", "£1 GBP = $1.3456", "1 GBP = 21,345.67 IDR", "£1 GBP = Rp 21,345"
    return re.compile(rf"1\s*GBP\s*=\s*[^\d\n]{{0,8}}?([\d,]+(?:\.\d+)?)\s*(?:{code})?", re.I)


def save_debug(page, name):
    DEBUG_DIR.mkdir(exist_ok=True)
    stem = DEBUG_DIR / name
    page.screenshot(path=f"{stem}.png", full_page=True)
    Path(f"{stem}.html").write_text(page.content(), encoding="utf-8")
    print(f"    saved {stem.name}.png/.html in the debug folder")


def is_challenge(page):
    html = page.content().lower()
    return any(k in html for k in ("just a moment", "verify you are human", "challenge-platform",
                                   "security verification"))


def scrape_page(page, url, code):
    """Returns {date: rate} for one ERUK page."""
    print(f"  {url}")
    try:
        resp = page.goto(url, wait_until="domcontentloaded", timeout=60000)
    except Exception as e:
        print(f"    could not load ({e.__class__.__name__})")
        return {}
    if resp is not None and resp.status == 404:
        print("    page does not exist (404)")
        return {}

    challenged = is_challenge(page)
    if challenged:
        print("    Cloudflare check showing - tick 'Verify you are human' in the browser window")
    try:
        page.wait_for_function("() => /GBP\\s*=/.test(document.body.innerText)",
                               timeout=(WAIT_FOR_YOU if challenged else WAIT_NORMAL) * 1000)
    except Exception:
        pass

    rx = rate_regex(code)
    found = []
    for text in page.locator("tr").all_inner_texts():
        m, d = rx.search(text), parse_date(text)
        if m and d and d <= TODAY:       # ERUK pads the yearly page with future Saturdays
            try:
                rate = float(m.group(1).replace(",", ""))
            except ValueError:
                continue
            if rate > 0:
                # rows starting with a weekday are the real daily rows; summary rows lose ties
                found.append((0 if WEEKDAY_START.match(text) else 1, d, round(rate, 6)))

    out = {}
    for _, d, rate in sorted(found, key=lambda x: x[0], reverse=True):
        out[d] = rate                    # priority-0 rows written last, so they win
    if out:
        print(f"    {len(out)} days")
    else:
        print("    no rates found on this page")
        save_debug(page, url.rsplit("/", 1)[-1].replace(".html", ""))
    return out


def year_has_gaps(dates, year):
    start, end = date(year, 1, 1), min(date(year, 12, 31), TODAY - timedelta(days=5))
    if end < start:
        return False
    ds = sorted(d for d in dates if start <= d <= end)
    if not ds:
        return True
    points = [start - timedelta(days=1)] + ds + [end + timedelta(days=1)]
    return max((b - a).days for a, b in zip(points, points[1:])) > MAX_GAP_DAYS


def load_history():
    store = {c: {} for c in CURRENCIES}
    if HISTORY_CSV.exists():
        df = pd.read_csv(HISTORY_CSV, parse_dates=["Date"])
        for row in df.itertuples(index=False):
            if row.Currency in store and row.Date.date() <= TODAY:
                store[row.Currency][row.Date.date()] = float(row.Rate)
        print(f"Loaded {len(df)} stored rates from {HISTORY_CSV.name}")
    return store


def save_history(store):
    rows = [(d, c, r) for c, series in store.items() for d, r in series.items()]
    df = pd.DataFrame(rows, columns=["Date", "Currency", "Rate"]).sort_values(["Date", "Currency"])
    df.to_csv(HISTORY_CSV, index=False, date_format="%Y-%m-%d")


def scrape_all(store):
    to_scrape = [c for c in CURRENCIES if c != "GBP" and c not in PEGGED_TO_EUR]
    with sync_playwright() as p:
        browser = p.chromium.launch_persistent_context(
            PROFILE_DIR,
            headless=HEADLESS,
            args=["--disable-blink-features=AutomationControlled", "--use-fake-ui-for-media-stream"],
            ignore_default_args=["--enable-automation"],
        )
        page = browser.pages[0] if browser.pages else browser.new_page()
        page.add_init_script("Object.defineProperty(navigator, 'webdriver', {get: () => undefined})")
        try:
            for ccy in to_scrape:
                print(f"\nGBP/{ccy}")
                store[ccy].update(scrape_page(page, f"{BASE_URL}/GBP-{ccy}-exchange-rate-history.html", ccy))
                time.sleep(PAUSE_BETWEEN_PAGES)
                for year in range(START_YEAR, TODAY.year + 1):
                    if year_has_gaps(store[ccy].keys(), year):
                        url = f"{BASE_URL}/GBP-{ccy}-spot-exchange-rates-history-{year}.html"
                        store[ccy].update(scrape_page(page, url, ccy))
                        time.sleep(PAUSE_BETWEEN_PAGES)
        finally:
            browser.close()
            save_history(store)          # keep whatever was pulled, even if the run stops halfway


# ---------------- tables ----------------
def build_daily(store):
    frames = [pd.Series(s, name=c, dtype=float) for c, s in store.items() if s]
    if not frames:
        return None
    wide = pd.concat(frames, axis=1)
    wide.index = pd.to_datetime(wide.index)
    wide = wide[(wide.index >= pd.Timestamp(START_YEAR, 1, 1)) & (wide.index.dayofweek != 5)]
    for c in CURRENCIES:
        if c not in wide:
            wide[c] = float("nan")
    wide.sort_index(inplace=True)
    wide["GBP"] = 1.0
    for ccy, per_eur in PEGGED_TO_EUR.items():
        wide[ccy] = wide["EUR"] * per_eur
    return wide[CURRENCIES].round(6)


def monthly_tables(wide):
    month = wide.index.to_period("M")
    avg = wide.groupby(month).mean().round(6)
    close = wide.groupby(month).agg(lambda s: s.dropna().iloc[-1] if s.notna().any() else float("nan")).round(6)
    if ONLY_COMPLETE_MONTHS:
        current = pd.Period(TODAY, "M")
        avg, close = avg[avg.index < current], close[close.index < current]
    return avg, close


# ---------------- Excel ----------------
# The workbook is updated in place: values are written into the existing file and styling is only
# applied to cells the script creates, so highlights, comments, colours, widths and extra sheets
# added by the team are kept. Don't insert or delete rows/columns inside the generated tables.
FONT = Font(name="Arial", size=10)
BOLD = Font(name="Arial", size=10, bold=True)
NOTE = Font(name="Arial", size=9, italic=True)
HEAD_FILL = PatternFill("solid", start_color="D9E1F2")
RATE_FMT = "#,##0.000000"
SHEETS = ["Average Currency_PL", "Closing Rate_CASH Position", "Daily Rate GESCO", "ERUK History"]
SOURCE_CELL = (24, 5)                       # E24, just above the first average table


def put(ws, row, col, value, fmt=None, font=None, fill=None):
    """Write a value; style the cell only if it was empty, so existing formatting is never touched."""
    c = ws.cell(row, col)
    new = c.value is None
    c.value = value
    if new and value is not None:
        c.font = font or FONT
        if fmt:
            c.number_format = fmt
        if fill:
            c.fill = fill
    return c


def as_date(v):
    if isinstance(v, datetime):
        return v.date()
    if isinstance(v, date):
        return v
    if isinstance(v, str):
        for fmt in ("%d/%m/%Y", "%d/%m/%y", "%Y-%m-%d", "%d.%m.%Y"):
            try:
                return datetime.strptime(v.strip(), fmt).date()
            except ValueError:
                pass
    return None


LOCKED_MSG = (f"{OUTPUT.name} is locked. Close it in Excel (check other Excel windows too), "
              "wait for OneDrive to finish syncing, then run again.")


def check_not_locked():
    """Fail fast, before scraping, if Excel or OneDrive has the workbook locked."""
    if OUTPUT.exists():
        try:
            with open(OUTPUT, "r+b"):
                pass
        except PermissionError:
            sys.exit(LOCKED_MSG)


def open_workbook():
    if OUTPUT.exists():
        try:
            wb = load_workbook(OUTPUT)
        except PermissionError:
            sys.exit(LOCKED_MSG + " (Today's rates are already saved in the CSV, so the rerun is quick.)")
        except Exception as e:
            sys.exit(f"Could not open {OUTPUT.name} ({e}). Rename it and run again to start a fresh one.")
        if "Overrides" in wb.sheetnames:            # tab from an earlier version, no longer used
            del wb["Overrides"]
        created = {name: name not in wb.sheetnames for name in SHEETS}
        for name in SHEETS:
            if created[name]:
                wb.create_sheet(name)
        return wb, created
    wb = Workbook()
    wb.active.title = SHEETS[0]
    for name in SHEETS[1:]:
        wb.create_sheet(name)
    return wb, {name: True for name in SHEETS}


def write_average_sheet(wb, created, avg, run_note):
    ws = wb["Average Currency_PL"]
    put(ws, 1, 3, f"List of rates to convert towards {BASE} (1 {BASE} = x currency)", font=BOLD)
    put(ws, 2, 2, "Subs", font=BOLD, fill=HEAD_FILL)
    put(ws, 2, 3, "Currency", font=BOLD, fill=HEAD_FILL)
    for i, (sub, ccy) in enumerate(SUBS, start=3):
        put(ws, i, 2, sub)
        put(ws, i, 3, ccy)

    # source line lives at E24; clear it from anywhere an older version of the script put it
    for row in ws.iter_rows(min_col=5, max_col=5):
        c = row[0]
        if isinstance(c.value, str) and c.value.startswith("Source: exchangerates") and c.row != SOURCE_CELL[0]:
            c.value, c.font = None, FONT
    put(ws, *SOURCE_CELL, run_note, font=NOTE)

    top = 26                                         # same anchor as the template
    for year in range(START_YEAR, TODAY.year + 1):   # one table per year, a new one each January
        head = dict(font=BOLD, fill=HEAD_FILL)
        put(ws, top, 5, BASE, **head)
        put(ws, top + 1, 5, "Avrg rate", **head)
        put(ws, top + 2, 5, None)
        for m in range(12):
            put(ws, top, 6 + m, "Monthly", **head)
            put(ws, top + 1, 6 + m, f"{m + 1:02d}.{year}", **head)
            put(ws, top + 2, 6 + m, MONTH_NAMES[m], **head)
        put(ws, top, 18, BASE, **head)
        put(ws, top + 1, 18, year, **head)
        put(ws, top + 2, 18, "Average", **head)
        if ws.cell(top + 2, 5).fill.fgColor.rgb in (None, "00000000"):
            ws.cell(top + 2, 5).fill = HEAD_FILL

        year_avg = avg[avg.index.year == year]
        for i, label in enumerate(AVG_ROWS):
            r, ccy = top + 3 + i, iso(label)
            put(ws, r, 5, label, font=BOLD)
            for m in range(1, 13):
                per = pd.Period(year=year, month=m, freq="M")
                v = year_avg.at[per, ccy] if per in year_avg.index else float("nan")
                put(ws, r, 5 + m, float(v) if pd.notna(v) else None, fmt=RATE_FMT)
            put(ws, r, 18, f"=IFERROR(AVERAGEA(F{r}:Q{r}),)", fmt=RATE_FMT)
        top += 3 + len(AVG_ROWS) + 2

    if created["Average Currency_PL"]:
        ws.column_dimensions["B"].width = 18
        ws.column_dimensions["E"].width = 11
        for col in range(6, 19):
            ws.column_dimensions[get_column_letter(col)].width = 17   # fits IDR/COP like 21,377.207908


def write_closing_sheet(wb, created, close):
    ws = wb["Closing Rate_CASH Position"]
    # B..S = 18 currencies with the base first (= 1), T = zero/empty (index 0), as in the template
    layout = [(c, i + 1) for i, c in enumerate(CLOSING)] + [("zero/empty", 0)]

    put(ws, 1, 1, BASE_URL).hyperlink = BASE_URL
    put(ws, 2, 1, "=COUNT(A3:A10000)+2", font=BOLD, fill=PatternFill("solid", start_color="FFFF00"))
    for j, (head, idx) in enumerate(layout, start=2):
        put(ws, 1, j, head, font=BOLD, fill=HEAD_FILL)
        put(ws, 2, j, idx, font=BOLD)

    rows = {}                                         # month-end date -> existing row
    for r in range(3, ws.max_row + 1):
        d = as_date(ws.cell(r, 1).value)
        if d:
            rows[d] = r
    next_row = max(rows.values(), default=2) + 1

    for per, vals in close.iterrows():
        month_end = date(per.year, per.month, calendar.monthrange(per.year, per.month)[1])
        r = rows.get(month_end)
        if r is None:
            r, next_row = next_row, next_row + 1
        put(ws, r, 1, month_end, fmt="dd/mm/yy;@")
        for j, (head, _) in enumerate(layout, start=2):
            if head == "zero/empty":
                put(ws, r, j, 0, fmt="0.0000")
            else:
                put(ws, r, j, float(vals[head]) if pd.notna(vals[head]) else None, fmt=RATE_FMT)

    if created["Closing Rate_CASH Position"]:
        ws.column_dimensions["A"].width = 12
        for j in range(2, len(layout) + 2):
            ws.column_dimensions[get_column_letter(j)].width = 16
        ws.freeze_panes = "B3"


def ensure_table(ws, name, ref):
    # rebuilt every run so the column names always match the header cells (Excel errors otherwise)
    style = TableStyleInfo(name="TableStyleLight1", showRowStripes=True)
    if name in ws.tables:
        style = ws.tables[name].tableStyleInfo or style
        del ws.tables[name]
    t = Table(displayName=name, ref=ref)
    t.tableStyleInfo = style
    ws.add_table(t)


def list_headers(ws, row):
    for j, h in enumerate(["Date", "Currency", "ISO_Code", f"Rate_{BASE}"], start=1):
        put(ws, row, j, h)


def write_latest_sheet(wb, created, wide):
    ws = wb["Daily Rate GESCO"]
    list_headers(ws, 1)
    last = 1 + len(DAILY)
    for r, (label, name) in enumerate(DAILY, start=2):   # always in DAILY order
        s = wide[iso(label)].dropna()
        put(ws, r, 1, s.index[-1].date() if not s.empty else None, fmt="dd/mm/yyyy")
        put(ws, r, 2, name)
        put(ws, r, 3, label)
        put(ws, r, 4, float(s.iloc[-1]) if not s.empty else None, fmt=RATE_FMT)
    r = last + 1                                          # clear leftovers from older layouts
    while any(ws.cell(r, j).value is not None for j in range(1, 5)):
        for j in range(1, 5):
            ws.cell(r, j).value = None
        r += 1
    ensure_table(ws, "Daily_Rate_GESCO", f"A1:D{last}")
    if created["Daily Rate GESCO"]:
        for col, w in zip("ABCD", (12, 26, 10, 16)):
            ws.column_dimensions[col].width = w


def write_history_sheet(wb, created, wide):
    ws = wb["ERUK History"]
    put(ws, 1, 1, f"Daily rate {BASE} to XXX _ Source Exchange Rates UK",
        font=Font(name="Arial", size=10, bold=True, italic=True))
    put(ws, 2, 1, BASE_URL, font=Font(name="Arial", size=10, underline="single", color="0563C1")).hyperlink = BASE_URL
    put(ws, 4, 1, "History Database", font=Font(name="Arial", size=10, bold=True, italic=True, underline="single"))
    list_headers(ws, 6)

    existing, last = {}, 6                            # (date, code) -> row
    r = 7
    while ws.cell(r, 1).value is not None:
        existing[(as_date(ws.cell(r, 1).value), ws.cell(r, 3).value)] = r
        last, r = r, r + 1

    labels = [c for c, _ in DAILY]
    for d, vals in wide.iterrows():
        for label in labels:
            v = vals[iso(label)]
            if pd.isna(v):
                continue
            row = existing.get((d.date(), label))
            if row is None:                           # new day: append at the bottom
                last += 1
                row = last
                put(ws, row, 1, d.date(), fmt="dd/mm/yyyy")
                put(ws, row, 2, NAMES[label])
                put(ws, row, 3, label)
            put(ws, row, 4, float(v), fmt=RATE_FMT)
    if last > 6:
        ensure_table(ws, "ERUK_History", f"A6:D{last}")
    if created["ERUK History"]:
        for col, w in zip("ABCD", (12, 26, 10, 16)):
            ws.column_dimensions[col].width = w
        ws.freeze_panes = "A7"


def main():
    check_not_locked()
    store = load_history()
    scrape_all(store)
    wide = build_daily(store)
    if wide is None or wide.drop(columns="GBP").isna().all().all():
        sys.exit("Nothing scraped. Check the files in the debug folder.")

    base_wide = wide if BASE == "GBP" else wide.div(wide[BASE], axis=0).round(6)   # 1 BASE = x, day by day
    avg, close = monthly_tables(base_wide)
    run_note = f"Source: exchangerates.org.uk, pulled {datetime.now():%d/%m/%Y %H:%M}."

    wb, created = open_workbook()
    write_average_sheet(wb, created, avg, run_note)
    write_closing_sheet(wb, created, close)
    write_latest_sheet(wb, created, base_wide)
    write_history_sheet(wb, created, base_wide)
    try:
        wb.save(OUTPUT)
    except PermissionError:
        sys.exit(LOCKED_MSG + " (Today's rates are already saved in the CSV, so the rerun is quick.)")

    print(f"\nDone. {len(wide)} days, {len(close)} closed months -> {OUTPUT.name}")
    missing = [c for c in CURRENCIES if wide[c].isna().all()]
    if missing:
        print("No data for: " + ", ".join(missing) + " (ERUK may not list them, see the debug folder)")
    latest = wide.drop(columns="GBP").dropna(how="all").index.max().date()
    if (TODAY - latest).days > 7:
        print(f"WARNING: newest rate is from {latest:%d/%m/%Y}. Pages probably failed to load this run.")
    for c in CURRENCIES:
        s = wide[c].dropna()
        if not s.empty and s.index.min().date() > date(START_YEAR, 1, 7):
            print(f"Note: {c} only has rates from {s.index.min():%d/%m/%Y}, earlier months are blank.")


if __name__ == "__main__":
    main()
