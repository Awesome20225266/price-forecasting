#!/usr/bin/env python3
"""
HPX India market-data auto-downloader
=====================================

Downloads these reports from https://www.hpxindia.com (Market Data menu) and
keeps them up to date in Excel. Every time you run it, it only fetches what is
new (plus the last few days again, in case HPX revised them) and merges it into
your existing files, with duplicates removed.

  Excel file                        Sheets
  --------------------------------  ---------------------------------------------
  Contingency_TAM.xlsx              Trade Details, Trade Summary
  TAM_Monthly.xlsx                  Trade Details
  GTAM_Monthly.xlsx                 Trade Details
  HighPrice_Contingency_TAM.xlsx    Trade Details, Trade Summary
  AS_DAM.xlsx                       UP, UP Daily Summary, DOWN, DOWN Daily Summary
  Green_Contingency_TAM.xlsx        Trade Details, Trade Summary

Setup (once):
    pip install requests pandas openpyxl

Run:
    python hpx_downloader.py                    # normal update
    python hpx_downloader.py --from 2025-04-01  # re-download everything from a date
    python hpx_downloader.py --only as_dam      # just one report (see --list)
    python hpx_downloader.py --dry-run          # show what would be fetched

Files are written to an "HPX_Data" folder next to this script. The "_store"
sub-folder holds the master CSV copies and the download progress, so don't
delete it unless you want a fresh start. Close the Excel files before running;
if one is open, the script tells you, and the data is still saved in _store.
"""

import argparse
import json
import os
import re
import sys
import time
from datetime import date, datetime, timedelta
from pathlib import Path

try:
    import pandas as pd
    import requests
    from requests.adapters import HTTPAdapter
    from urllib3.util.retry import Retry
except ImportError:
    sys.exit("Missing packages. Run:  pip install requests pandas openpyxl")

# =============================== SETTINGS ===================================
START_DATE = date(2023, 4, 1)  # first run downloads history from this date
OVERLAP_DAYS = 3               # re-fetch the last N days every run (catches revisions)
CHUNK_DAYS = 30                # days per API request
OUTPUT_DIR = Path(__file__).resolve().parent / "HPX_Data"
# ============================================================================

API = "https://api.hpxindia.com/api/"
HEADERS = {
    "User-Agent": ("Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
                   "(KHTML, like Gecko) Chrome/128.0 Safari/537.36"),
    "Accept": "application/json, text/plain, */*",
    "Origin": "https://www.hpxindia.com",
    "Referer": "https://www.hpxindia.com/",
}
EXCEL_MAX_ROWS = 1_000_000
TEXT_COLS = {"contracttype", "instrument name", "timeblock", "delivery month", "summary"}
SORT_COLS = ["contracttype", "instrument name", "timeblock", "summary"]


def ymd(d):
    return d.strftime("%Y-%m-%d")


def ymd8(d):
    return d.strftime("%Y%m%d")


def trade_details(s, e):
    return {"contracttype": "All", "flag": "web", "reporttype": "Trade Date Wise",
            "tradeperiod": "Select Range", "fromdate": ymd(s), "Todate": ymd(e)}


def trade_summary(s, e):
    return {"contracttype": "All", "flag": "web", "tradeperiod": "Select Range",
            "startdate": ymd(s), "enddate": ymd(e)}


def monthly(s, e):
    return {"TradePeriod": "Select Range", "deliveryperiod": "DAYWISE",
            "fromdate": ymd8(s), "Todate": ymd8(e)}


def green_details(s, e):
    return {"reporttype": "Trade Date Wise", "flag": "web", "contracttype": "All",
            "solarType": "All", "tradeperiod": "Select Range",
            "startdate": ymd(s), "enddate": ymd(e)}


def asdam(trans_type):
    def params(s, e):
        return {"interval": "15-Minute Block", "deliveryperiod": "Select Range",
                "flag": "web", "fromweek": "", "toweek": "", "frommonth": "",
                "tomonth": "", "years": "", "fromdate": ymd(s) + " 00:00",
                "todate": ymd(e) + " 23:59", "TransType": trans_type}
    return params


# Each report = one API call per date chunk. "outputs" maps the API's table
# name -> (Excel file, sheet name, store name).
REPORTS = [
    dict(key="contingency_tam_details", endpoint="getTradeDetails/w", params=trade_details,
         date_col="tradedate",
         outputs={"table": ("Contingency_TAM.xlsx", "Trade Details", "contingency_tam_details")}),
    dict(key="contingency_tam_summary", endpoint="getTradeSummaryData2/w", params=trade_summary,
         date_col="tradedate",
         outputs={"table": ("Contingency_TAM.xlsx", "Trade Summary", "contingency_tam_summary")}),
    dict(key="tam_monthly", endpoint="TAMonthlyData1/w", params=monthly,
         date_col="tradedate",
         outputs={"table": ("TAM_Monthly.xlsx", "Trade Details", "tam_monthly")}),
    dict(key="gtam_monthly", endpoint="GTAMonthlyData/w", params=monthly,
         date_col="tradedate",
         outputs={"table": ("GTAM_Monthly.xlsx", "Trade Details", "gtam_monthly")}),
    dict(key="hp_tam_details", endpoint="getTradeDetails1/w", params=trade_details,
         date_col="tradedate",
         outputs={"table": ("HighPrice_Contingency_TAM.xlsx", "Trade Details", "hp_tam_details")}),
    dict(key="hp_tam_summary", endpoint="getTradeSummaryData212/w", params=trade_summary,
         date_col="tradedate",
         outputs={"table": ("HighPrice_Contingency_TAM.xlsx", "Trade Summary", "hp_tam_summary")}),
    dict(key="as_dam_up", endpoint="getDataMarkSnapASDAM/w", params=asdam("UP"),
         date_col="deliverydate", end_offset=1,  # DAM delivers next day
         outputs={"table": ("AS_DAM.xlsx", "UP", "as_dam_up"),
                  "table1": ("AS_DAM.xlsx", "UP Daily Summary", "as_dam_up_daily")}),
    dict(key="as_dam_down", endpoint="getDataMarkSnapASDAM/w", params=asdam("DOWN"),
         date_col="deliverydate", end_offset=1,
         outputs={"table": ("AS_DAM.xlsx", "DOWN", "as_dam_down"),
                  "table1": ("AS_DAM.xlsx", "DOWN Daily Summary", "as_dam_down_daily")}),
    dict(key="green_tam_details", endpoint="getGTamDetails/w", params=green_details,
         date_col="tradedate",
         outputs={"table": ("Green_Contingency_TAM.xlsx", "Trade Details", "green_tam_details")}),
    dict(key="green_tam_summary", endpoint="getTradeSummaryData211/w", params=trade_summary,
         date_col="tradedate",
         outputs={"table": ("Green_Contingency_TAM.xlsx", "Trade Summary", "green_tam_summary")}),
]


# ------------------------------------------------------------------ helpers --
def make_session():
    s = requests.Session()
    s.headers.update(HEADERS)
    retry = Retry(total=4, backoff_factor=3, status_forcelist=(429, 500, 502, 503, 504),
                  allowed_methods=("GET",))
    s.mount("https://", HTTPAdapter(max_retries=retry))
    return s


def date_chunks(start, end, size):
    cur = start
    while cur <= end:
        stop = min(cur + timedelta(days=size - 1), end)
        yield cur, stop
        cur = stop + timedelta(days=1)


def fetch(session, endpoint, params):
    r = session.get(API + endpoint, params=params, timeout=180)
    r.raise_for_status()
    data = r.json()
    return data if isinstance(data, dict) else {}


def clean(df, date_col):
    """Parse the dd/mm/yyyy date column and turn numeric-looking text into numbers."""
    if df.empty:
        return df
    df = df.copy()
    if date_col in df.columns:
        df[date_col] = pd.to_datetime(df[date_col], format="%d/%m/%Y", errors="coerce")
    for c in df.columns:
        if (c == date_col or c in TEXT_COLS or pd.api.types.is_numeric_dtype(df[c])
                or pd.api.types.is_datetime64_any_dtype(df[c])):
            continue
        conv = pd.to_numeric(df[c], errors="coerce")
        if conv.notna().sum() == df[c].notna().sum():
            df[c] = conv
    return df


def store_path(name):
    return OUTPUT_DIR / "_store" / f"{name}.csv"


def load_store(name, date_col):
    p = store_path(name)
    if not p.exists():
        return pd.DataFrame()
    df = pd.read_csv(p)
    if date_col in df.columns:
        df[date_col] = pd.to_datetime(df[date_col], errors="coerce")
    return df


def save_store(name, df):
    p = store_path(name)
    p.parent.mkdir(parents=True, exist_ok=True)
    tmp = p.with_suffix(".tmp")
    df.to_csv(tmp, index=False, date_format="%Y-%m-%d")
    os.replace(tmp, p)


def merge(existing, new, date_col, win_start, win_end):
    """Replace the rows of `existing` inside [win_start, win_end] with `new`.
    An empty response never wipes rows we already have."""
    if new.empty:
        return existing
    if not existing.empty and date_col in existing.columns:
        d = existing[date_col]
        inside = (d >= pd.Timestamp(win_start)) & (d <= pd.Timestamp(win_end))
        existing = existing[~inside]
    frames = [f for f in (existing, new) if not f.empty]
    if not frames:
        return pd.DataFrame()
    out = pd.concat(frames, ignore_index=True).drop_duplicates()
    sort = [c for c in [date_col] + SORT_COLS if c in out.columns]
    return out.sort_values(sort, kind="stable").reset_index(drop=True)


def pretty(col):
    names = {"tradedate": "Trade Date", "deliverydate": "Delivery Date",
             "contracttype": "Contract Type", "timeblock": "Time Block"}
    if col in names:
        return names[col]
    s = re.sub(r"(\S)\(", r"\1 (", col).title()
    for a, b in [("(In Mus)", "(MUs)"), ("(In Mwh)", "(MWh)"), ("(In Rs/Mwh)", "(Rs/MWh)"),
                 ("Rs/Mwh", "Rs/MWh"), ("(Mwh)", "(MWh)"), ("Mcp", "MCP"),
                 ("Tras ", "TRAS "), ("No. Of", "No. of")]:
        s = s.replace(a, b)
    return s


def write_excel(filename, sheets):
    """sheets: list of (sheet_name, DataFrame, date_col). Returns True on success."""
    path = OUTPUT_DIR / filename
    tmp = path.with_name("~tmp_" + filename)
    with pd.ExcelWriter(tmp, engine="openpyxl", date_format="DD-MM-YYYY",
                        datetime_format="DD-MM-YYYY") as xw:
        for sheet, df, date_col in sheets:
            if df.empty:
                df = pd.DataFrame({"Info": ["No data yet for the downloaded period"]})
            parts = [df.iloc[i:i + EXCEL_MAX_ROWS] for i in range(0, len(df), EXCEL_MAX_ROWS)]
            for n, part in enumerate(parts, 1):
                name = sheet if n == 1 else f"{sheet} {n}"
                part.rename(columns=pretty).to_excel(xw, sheet_name=name[:31], index=False)
                ws = xw.sheets[name[:31]]
                ws.freeze_panes = "A2"
                ws.auto_filter.ref = ws.dimensions
                for i, col in enumerate(part.columns, 1):
                    letter = ws.cell(row=1, column=i).column_letter
                    if col == date_col:
                        for (cell,) in ws.iter_rows(min_row=2, min_col=i, max_col=i):
                            cell.number_format = "DD-MM-YYYY"
                        ws.column_dimensions[letter].width = 14
                        continue
                    sample = part[col].head(500).astype(str).str.len().max() if len(part) else 0
                    width = min(max(len(pretty(col)), int(sample or 0), 10) + 2, 45)
                    ws.column_dimensions[letter].width = width
    try:
        os.replace(tmp, path)
        return True
    except PermissionError:
        tmp.unlink(missing_ok=True)
        print(f"  ! Could not update {filename} - is it open in Excel? Close it and run again "
              f"(the data is already saved in _store).")
        return False


# --------------------------------------------------------------------- main --
def main():
    ap = argparse.ArgumentParser(description="Download / update HPX India market data.")
    ap.add_argument("--from", dest="from_date", help="re-download everything from YYYY-MM-DD")
    ap.add_argument("--only", help="comma-separated report keys to run (see --list)")
    ap.add_argument("--list", action="store_true", help="list report keys and exit")
    ap.add_argument("--dry-run", action="store_true", help="show requests without downloading")
    args = ap.parse_args()

    if args.list:
        for r in REPORTS:
            print(f"  {r['key']:26s} -> {', '.join(f'{o[0]} [{o[1]}]' for o in r['outputs'].values())}")
        return

    reports = REPORTS
    if args.only:
        wanted = {k.strip() for k in args.only.split(",")}
        reports = [r for r in REPORTS if r["key"] in wanted or any(r["key"].startswith(w) for w in wanted)]
        if not reports:
            sys.exit(f"No report matches --only {args.only}. Use --list to see the keys.")

    forced = datetime.strptime(args.from_date, "%Y-%m-%d").date() if args.from_date else None
    state_file = OUTPUT_DIR / "_store" / "state.json"
    state = json.loads(state_file.read_text()) if state_file.exists() else {}
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    session = make_session()
    today = date.today()
    touched_files = set()
    failures = []

    print(f"HPX downloader - {datetime.now():%d-%m-%Y %H:%M}  ->  {OUTPUT_DIR}\n")
    for rep in reports:
        key, dcol = rep["key"], rep["date_col"]
        end = today + timedelta(days=rep.get("end_offset", 0))
        if forced:
            start = forced
        elif key in state:
            start = datetime.strptime(state[key], "%Y-%m-%d").date() - timedelta(days=OVERLAP_DAYS)
        else:
            start = START_DATE
        start = min(start, end)
        print(f"- {key}: {start:%d-%m-%Y} to {end:%d-%m-%Y}")

        if args.dry_run:
            for cs, ce in date_chunks(start, end, CHUNK_DAYS):
                req = requests.Request("GET", API + rep["endpoint"], params=rep["params"](cs, ce)).prepare()
                print("    ", req.url)
            continue

        stores = {t: load_store(o[2], dcol) for t, o in rep["outputs"].items()}
        added = {t: 0 for t in stores}
        for cs, ce in date_chunks(start, end, CHUNK_DAYS):
            print(f"    {cs:%d-%m-%Y}..{ce:%d-%m-%Y}", flush=True)
            try:
                data = fetch(session, rep["endpoint"], rep["params"](cs, ce))
            except Exception as exc:  # network / server error: keep what we have
                print(f"    ! {cs:%d-%m-%Y}..{ce:%d-%m-%Y} failed: {exc}")
                failures.append(key)
                break
            for table, (fname, sheet, sname) in rep["outputs"].items():
                new = clean(pd.DataFrame(data.get(table) or []), dcol)
                stores[table] = merge(stores[table], new, dcol, cs, ce)
                added[table] += len(new.drop_duplicates()) if not new.empty else 0
            state[key] = ymd(ce)
            time.sleep(0.5)

        for table, (fname, sheet, sname) in rep["outputs"].items():
            save_store(sname, stores[table])
            touched_files.add(fname)
            print(f"    {sheet:20s} fetched {added[table]:>7,} rows | total {len(stores[table]):>8,}")
        state_file.parent.mkdir(parents=True, exist_ok=True)
        state_file.write_text(json.dumps(state, indent=2))

    if args.dry_run:
        return

    # Rebuild each Excel file that had any report updated.
    print()
    for fname in sorted(touched_files):
        sheets = []
        for rep in REPORTS:
            for table, (f, sheet, sname) in rep["outputs"].items():
                if f == fname:
                    sheets.append((sheet, load_store(sname, rep["date_col"]), rep["date_col"]))
        if write_excel(fname, sheets):
            print(f"  saved {fname}")

    print("\nDone." if not failures else
          f"\nFinished with errors in: {', '.join(sorted(set(failures)))}. "
          "Run again later - it will resume where it stopped.")


if __name__ == "__main__":
    main()
