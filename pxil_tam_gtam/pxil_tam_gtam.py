#!/usr/bin/env python3
"""
PXIL TAM + GTAM data collector  ->  one master Excel workbook
=============================================================

Pulls every Term Ahead Market (TAM) and Green Term Ahead Market (GTAM)
section from https://powerexindia.in/Pages/Market.html, keeps a local
SQLite cache, back-fills any missing dates, and rebuilds a single Excel
workbook with one sheet per section.

Sections collected
------------------
TAM   : Market Snapshot, DAC, Intraday, Any Day, Daily, Weekly, Monthly
        (+ the site's Max/Min/Avg/Total summary for every product, per day)
GTAM  : Market Snapshot, DAC, Intraday, Any Day
        (+ per-day summary split by Solar / Non Solar)

How "missing dates" work
------------------------
Every (section, day) that was downloaded successfully is recorded in the
cache. Each run works out which days between START and END are not
recorded yet and downloads only those. Days that failed (network error,
site down, PC off) are simply not recorded, so the next run picks them up.
The most recent REFRESH_DAYS are always re-downloaded, because the
exchange keeps adding trades for today / tomorrow.

Usage
-----
    python pxil_tam_gtam.py                      # back-fill + update (normal daily run)
    python pxil_tam_gtam.py --start 2024-04-01   # extend the back-fill further back
    python pxil_tam_gtam.py --no-summaries       # skip the per-day summary sheets
    python pxil_tam_gtam.py --only GTAM          # just the GTAM sections
    python pxil_tam_gtam.py --export-only        # rebuild the Excel from the cache
    python pxil_tam_gtam.py --redownload 2026-09-01 2026-09-30   # force a range again

Requires: Python 3.9+, requests, pandas, xlsxwriter  (pip install -r requirements.txt)
"""

from __future__ import annotations

import argparse
import datetime as dt
import logging
import os
import sqlite3
import sys
import threading
import time
from concurrent.futures import ThreadPoolExecutor, as_completed
from logging.handlers import RotatingFileHandler
from pathlib import Path

import pandas as pd
import requests

# --------------------------------------------------------------------------
# Settings (all can be overridden from the command line)
# --------------------------------------------------------------------------
BASE_URL = "https://powerexindia.in"
DEFAULT_START = "2023-04-01"          # back-fill horizon (start of FY 2023-24)
DAYS_AHEAD = 1                        # DAC is traded for next-day delivery -> include tomorrow
REFRESH_DAYS = 3                      # always re-download the latest N days
CHUNK_DAYS = 31                       # days per request for the detail tables
WORKERS = 4                           # parallel requests (keep this small, be polite)
EXCEL_NAME = "PXIL_TAM_GTAM.xlsx"
DB_NAME = "pxil_tam_gtam_cache.sqlite"
LOG_NAME = "pxil_tam_gtam.log"

IST = dt.timezone(dt.timedelta(hours=5, minutes=30))
EXCEL_MAX_ROWS = 1_048_575            # data rows per sheet (1 header row)

# --------------------------------------------------------------------------
# Section catalogue  (endpoint + POST params + column labels as on the site)
# --------------------------------------------------------------------------
QTY = {"total_buy": "Total Buy (MWh)", "total_sell": "Total Sell (MWh)",
       "traded_qty": "Traded Quantity (MWh)", "traded_price": "Traded Price (Rs/MWh)"}

SUMMARY_COLS = {
    "max_qty": "Max Traded Qty (MWh)", "min_qty": "Min Traded Qty (MWh)",
    "avg_qty": "Avg Traded Qty (MWh)", "total_qty": "Total Traded Qty (MWh)",
    "max_price": "Max Price (Rs/MWh)", "min_price": "Min Price (Rs/MWh)",
    "avg_price": "Avg Price (Rs/MWh)",
}


def _detail(market, name, endpoint, params, date_col, columns, title):
    return dict(market=market, name=name, endpoint=endpoint, params=params,
                kind="detail", date_col=date_col, columns=columns, title=title)


def _summary(market, name, endpoint, params, title):
    cols = {"_date": "Delivery Date"}
    if market == "GTAM":
        cols["product_name"] = "Product"
    cols.update(SUMMARY_COLS)
    return dict(market=market, name=name, endpoint=endpoint, params=params,
                kind="summary", date_col="_date", columns=cols, title=title)


GT = {"dacmvpType": "ALL"}   # GTAM: ALL = Solar + Non Solar

SECTIONS = [
    # ---------------------------- TAM ----------------------------------
    _detail("TAM", "TAM_MarketSnapshot", "get_MS_mvp_data",
            {"interval": "Daily", "contract_type": "ALL"}, "date",
            {"contract_type": "Contract Type", "date": "Delivery Date",
             "total_volume": "Total Traded Volume (MWh)",
             "price": "Average Traded Price (Rs./MWh)"},
            "TAM - Market Snapshot / Market Volume Profile (daily, all contracts)"),
    _detail("TAM", "TAM_DAC", "get_dac_bil_data", {}, "delivery_date",
            {"delivery_date": "Delivery Date", "product_name": "Product Name",
             "hours": "Hours", **QTY},
            "TAM - DAC - Market Volume Profile Report"),
    _summary("TAM", "TAM_DAC_Summary", "get_dac_bil_summary", {},
             "TAM - DAC - daily summary"),
    _detail("TAM", "TAM_Intraday", "get_intra_bil_data", {}, "delivery_date",
            {"delivery_date": "Delivery Date", "product_name": "Product Name",
             "hours": "Hours", **QTY},
            "TAM - Intraday - Market Volume Profile Report"),
    _summary("TAM", "TAM_Intraday_Summary", "get_intra_bil_summary", {},
             "TAM - Intraday - daily summary"),
    _detail("TAM", "TAM_AnyDay", "get_anyday_bil_data", {}, "from_delivery_date",
            {"from_delivery_date": "From Delivery Date", "to_delivery_date": "To Delivery Date",
             "time_block": "Time Block", **QTY},
            "TAM - Any Day - Market Volume Profile Report"),
    _summary("TAM", "TAM_AnyDay_Summary", "get_anyday_bil_summary", {},
             "TAM - Any Day - daily summary"),
    _detail("TAM", "TAM_Daily", "get_daily_bil_data", {}, "delivery_date",
            {"delivery_date": "Delivery Date", "time_block": "Time Block",
             "traded_qty": QTY["traded_qty"], "traded_price": QTY["traded_price"]},
            "TAM - Daily - Market Volume Profile Report"),
    _summary("TAM", "TAM_Daily_Summary", "get_daily_bil_summary", {},
             "TAM - Daily - daily summary"),
    _detail("TAM", "TAM_Weekly", "get_tweekly_bil_data", {}, "delivery_date",
            {"delivery_date": "Delivery Date", "time_block": "Time Block",
             "traded_qty": QTY["traded_qty"], "traded_price": QTY["traded_price"]},
            "TAM - Weekly - Market Volume Profile Report"),
    _summary("TAM", "TAM_Weekly_Summary", "get_tweekly_bil_summary", {},
             "TAM - Weekly - daily summary"),
    _detail("TAM", "TAM_Monthly", "get_monthly_bilateral_data", {}, "delivery_date",
            {"delivery_date": "Delivery Date", "time_block": "Time Block",
             "traded_qty": QTY["traded_qty"], "traded_price": QTY["traded_price"]},
            "TAM - Monthly - Market Volume Profile Report"),
    _summary("TAM", "TAM_Monthly_Summary", "get_monthly_bil_summary", {},
             "TAM - Monthly - daily summary"),
    # ---------------------------- GTAM ---------------------------------
    _detail("GTAM", "GTAM_MarketSnapshot", "get_gMS_mvp_data",
            {"interval": "Daily", "contract": "ALL"}, "date",
            {"contract_type": "Contract", "date": "Delivery Date",
             "green_volume": "Solar Volume (MWh)",
             "nongreen_volume": "Non Solar Volume (MWh)",
             "total_volume": "Total Traded Volume (MWh)"},
            "GTAM - Market Snapshot / Market Volume Profile (daily, all contracts)"),
    _detail("GTAM", "GTAM_DAC", "get_gdac_bil_data", GT, "delivery_from_date",
            {"delivery_from_date": "Delivery From Date", "delivery_to_date": "Delivery To Date",
             "time_block": "Time Block", "product_code": "Product", **QTY},
            "GTAM - Day Ahead Contingency - Market Volume Profile Report (Solar + Non Solar)"),
    _summary("GTAM", "GTAM_DAC_Summary", "get_gdac_bil_summary", GT,
             "GTAM - Day Ahead Contingency - daily summary by product"),
    _detail("GTAM", "GTAM_Intraday", "get_gintra_bil_data", GT, "delivery_from_date",
            {"delivery_from_date": "Delivery From Date", "delivery_to_date": "Delivery To Date",
             "time_block": "Time Block", "product_code": "Product", **QTY},
            "GTAM - Intra Day - Market Volume Profile Report (Solar + Non Solar)"),
    _summary("GTAM", "GTAM_Intraday_Summary", "get_gintra_bil_summary", GT,
             "GTAM - Intra Day - daily summary by product"),
    _detail("GTAM", "GTAM_AnyDay", "get_ganyday_bil_data", GT, "delivery_from_date",
            {"delivery_from_date": "Delivery From Date", "delivery_to_date": "Delivery To Date",
             "time_block": "Time Block", "product_code": "Product", **QTY},
            "GTAM - Any Day - Market Volume Profile Report (Solar + Non Solar)"),
    _summary("GTAM", "GTAM_AnyDay_Summary", "get_ganyday_bil_summary", GT,
             "GTAM - Any Day - daily summary by product"),
]

log = logging.getLogger("pxil")


# --------------------------------------------------------------------------
# Small helpers
# --------------------------------------------------------------------------
def parse_day(value) -> str | None:
    """'29-Sep-2026' / '29-September-2026' / '2026-09-29' -> '2026-09-29'."""
    if value is None:
        return None
    s = str(value).strip()
    for fmt in ("%d-%b-%Y", "%d-%B-%Y", "%Y-%m-%d", "%d/%m/%Y", "%d-%m-%Y"):
        try:
            return dt.datetime.strptime(s, fmt).date().isoformat()
        except ValueError:
            pass
    return None


def daterange(start: dt.date, end: dt.date):
    d = start
    while d <= end:
        yield d
        d += dt.timedelta(days=1)


def chunk_days(days: list[dt.date], max_len: int) -> list[tuple[dt.date, dt.date]]:
    """Group sorted days into contiguous (from, to) ranges of at most max_len days."""
    ranges: list[tuple[dt.date, dt.date]] = []
    for d in sorted(days):
        if ranges:
            a, b = ranges[-1]
            if d == b + dt.timedelta(days=1) and (d - a).days < max_len:
                ranges[-1] = (a, d)
                continue
        ranges.append((d, d))
    return ranges


def energy_type(code) -> str:
    c = str(code or "").upper()
    if c.endswith("_NSL"):
        return "Non Solar"
    if c.endswith("_SL"):
        return "Solar"
    return ""


def qident(name: str) -> str:
    return '"' + name.replace('"', '""') + '"'


# --------------------------------------------------------------------------
# HTTP
# --------------------------------------------------------------------------
class Client:
    def __init__(self, base_url: str, timeout: int = 90, retries: int = 4):
        self.base_url = base_url.rstrip("/")
        self.timeout = timeout
        self.retries = retries
        self._local = threading.local()

    def _session(self) -> requests.Session:
        s = getattr(self._local, "s", None)
        if s is None:
            s = requests.Session()
            s.headers.update({
                "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
                              "AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124 Safari/537.36",
                "X-Requested-With": "XMLHttpRequest",
                "Accept": "application/json, text/javascript, */*; q=0.01",
                "Referer": f"{self.base_url}/Pages/Market.html",
                "Origin": self.base_url,
            })
            self._local.s = s
        return s

    def post(self, endpoint: str, data: dict):
        url = f"{self.base_url}/{endpoint.strip('/')}/"
        last = None
        for attempt in range(1, self.retries + 1):
            try:
                r = self._session().post(url, data=data, timeout=self.timeout)
                if r.status_code >= 500 or r.status_code == 429:
                    raise requests.HTTPError(f"HTTP {r.status_code}")
                r.raise_for_status()
                return r.json()
            except (requests.RequestException, ValueError) as e:
                last = e
                if attempt < self.retries:
                    time.sleep(2 ** attempt)
        raise RuntimeError(f"{endpoint} {data.get('from_date')}..{data.get('to_date')}: {last}")


# --------------------------------------------------------------------------
# Local cache (SQLite)
# --------------------------------------------------------------------------
class Store:
    def __init__(self, path: Path):
        self.conn = sqlite3.connect(str(path))
        self.conn.execute(
            "CREATE TABLE IF NOT EXISTS fetch_log ("
            " section TEXT, day TEXT, rows INTEGER, fetched_at TEXT,"
            " PRIMARY KEY (section, day))")
        for sec in SECTIONS:
            self._ensure_table(sec)
        self.conn.commit()

    @staticmethod
    def raw_cols(sec) -> list[str]:
        cols = [c for c in sec["columns"] if c != "_date"]
        return cols

    def _ensure_table(self, sec):
        name = sec["name"]
        cols = self.raw_cols(sec)
        self.conn.execute(
            f"CREATE TABLE IF NOT EXISTS {qident(name)} ("
            + ", ".join([f"{qident(c)}" for c in cols] + ['"_date" TEXT', '"_fetched_at" TEXT'])
            + ")")
        existing = {r[1] for r in self.conn.execute(f"PRAGMA table_info({qident(name)})")}
        for c in cols:                              # future-proof: new columns on the site
            if c not in existing:
                self.conn.execute(f"ALTER TABLE {qident(name)} ADD COLUMN {qident(c)}")
        self.conn.execute(
            f"CREATE INDEX IF NOT EXISTS {qident('ix_' + name)} ON {qident(name)} (\"_date\")")

    def done_days(self, section: str) -> set[str]:
        return {r[0] for r in self.conn.execute(
            "SELECT day FROM fetch_log WHERE section=?", (section,))}

    def save(self, sec, day_from: dt.date, day_to: dt.date, rows: list[dict]):
        """Replace everything for [day_from, day_to] with `rows`, then mark the days done."""
        name, cols = sec["name"], self.raw_cols(sec)
        a, b = day_from.isoformat(), day_to.isoformat()
        now = dt.datetime.now(IST).strftime("%Y-%m-%d %H:%M:%S")
        col_sql = ", ".join(qident(c) for c in cols + ["_date", "_fetched_at"])
        ph = ", ".join("?" for _ in range(len(cols) + 2))
        in_range, out_range = [], []
        for r in rows:
            day = r.get("_date") or parse_day(r.get(sec["date_col"]))
            vals = tuple(r.get(c) for c in cols) + (day, now)
            (in_range if day and a <= day <= b else out_range).append(vals)

        per_day: dict[str, int] = {}
        for v in in_range:
            per_day[v[-2]] = per_day.get(v[-2], 0) + 1

        with self.conn:  # one transaction
            self.conn.execute(
                f"DELETE FROM {qident(name)} WHERE \"_date\" BETWEEN ? AND ?", (a, b))
            self.conn.executemany(f"INSERT INTO {qident(name)} ({col_sql}) VALUES ({ph})", in_range)
            if out_range:  # rows dated outside the request window: add only if not present
                cond = " AND ".join(f"{qident(c)} IS ?" for c in cols + ["_date"])
                for v in out_range:
                    if not self.conn.execute(
                            f"SELECT 1 FROM {qident(name)} WHERE {cond} LIMIT 1", v[:-1]).fetchone():
                        self.conn.execute(
                            f"INSERT INTO {qident(name)} ({col_sql}) VALUES ({ph})", v)
            self.conn.executemany(
                "INSERT OR REPLACE INTO fetch_log (section, day, rows, fetched_at) VALUES (?,?,?,?)",
                [(name, d.isoformat(), per_day.get(d.isoformat(), 0), now)
                 for d in daterange(day_from, day_to)])

    def frame(self, sec) -> pd.DataFrame:
        cols = self.raw_cols(sec)
        sel = ", ".join(qident(c) for c in cols + ["_date"])
        return pd.read_sql_query(f"SELECT {sel} FROM {qident(sec['name'])}", self.conn)

    def coverage(self, section: str):
        return self.conn.execute(
            "SELECT MIN(day), MAX(day), COUNT(*), MAX(fetched_at) FROM fetch_log WHERE section=?",
            (section,)).fetchone()


# --------------------------------------------------------------------------
# Fetch jobs
# --------------------------------------------------------------------------
def fetch_detail(client: Client, sec, a: dt.date, b: dt.date) -> list[dict]:
    payload = {"from_date": a.isoformat(), "to_date": b.isoformat(), **sec["params"]}
    j = client.post(sec["endpoint"], payload)
    rows = j.get("data", []) if isinstance(j, dict) else j
    if not isinstance(rows, list):
        raise RuntimeError(f"{sec['name']}: unexpected response {str(j)[:200]}")
    return [r for r in rows if isinstance(r, dict)]


def fetch_summary(client: Client, sec, day: dt.date) -> list[dict]:
    payload = {"from_date": day.isoformat(), "to_date": day.isoformat(), **sec["params"]}
    j = client.post(sec["endpoint"], payload)
    if not isinstance(j, dict):
        raise RuntimeError(f"{sec['name']}: unexpected response {str(j)[:200]}")
    keys = list(SUMMARY_COLS)
    products = j.get("product_name")
    out = []
    if isinstance(products, list):             # GTAM: one value per product
        for i, p in enumerate(products):
            row = {"product_name": p}
            for k in keys:
                v = j.get(k)
                row[k] = v[i] if isinstance(v, list) and i < len(v) else v
            out.append(row)
    else:                                      # TAM: single row
        row = {k: j.get(k) for k in keys}
        if "product_name" in j:
            row["product_name"] = products
        out.append(row)
    # drop "no trades" rows (the site returns '-' everywhere)
    out = [r for r in out if any(str(r.get(k, "-")).strip() not in ("-", "", "None") for k in keys)]
    for r in out:
        r["_date"] = day.isoformat()
    return out


def plan_jobs(store: Store, sections, start: dt.date, end: dt.date, refresh_from: dt.date,
              force: tuple[dt.date, dt.date] | None):
    jobs = []
    for sec in sections:
        done = store.done_days(sec["name"])
        todo = []
        for d in daterange(start, end):
            iso = d.isoformat()
            if iso not in done or d >= refresh_from or (force and force[0] <= d <= force[1]):
                todo.append(d)
        if sec["kind"] == "detail":
            for a, b in chunk_days(todo, CHUNK_DAYS):
                jobs.append((sec, a, b))
        else:
            for d in todo:
                jobs.append((sec, d, d))
    return jobs


def run_fetch(store: Store, client: Client, jobs, workers: int) -> int:
    if not jobs:
        log.info("Nothing to download - everything is up to date.")
        return 0
    log.info("Downloading %d request(s) ...", len(jobs))
    failures = 0
    done = 0
    t0 = time.time()
    breaker = {"streak": 0, "open": False}      # stop hammering if the site is down
    lock = threading.Lock()

    def call(fn, *a):
        if breaker["open"]:
            raise RuntimeError("skipped - site not responding")
        try:
            res = fn(*a)
            with lock:
                breaker["streak"] = 0
            return res
        except Exception:
            with lock:
                breaker["streak"] += 1
                if breaker["streak"] >= 8 and not breaker["open"]:
                    breaker["open"] = True
                    log.error("PXIL is not responding - stopping this run early. "
                              "Missing dates will be downloaded next run.")
            raise

    def work(job):
        """Returns ([(from, to, rows), ...], [(day, error), ...])."""
        sec, a, b = job
        if sec["kind"] == "summary":
            return [(a, b, call(fetch_summary, client, sec, a))], []
        try:
            return [(a, b, call(fetch_detail, client, sec, a, b))], []
        except Exception as e:
            if a == b or breaker["open"]:
                raise
            # a bad day inside a big range should not block the whole range -> go day by day
            log.warning("%s %s..%s failed (%s) - retrying day by day", sec["name"], a, b, e)
            segs, errs, streak = [], [], 0
            for d in daterange(a, b):
                try:
                    segs.append((d, d, call(fetch_detail, client, sec, d, d)))
                    streak = 0
                except Exception as e2:  # noqa: BLE001
                    errs.append((d, e2))
                    streak += 1
                    if streak >= 3 or breaker["open"]:
                        errs.append((None, "gave up on the rest of this range"))
                        break
            return segs, errs

    with ThreadPoolExecutor(max_workers=max(1, workers)) as pool:
        futures = {pool.submit(work, j): j for j in jobs}
        for fut in as_completed(futures):
            sec, a, b = futures[fut]
            done += 1
            try:
                segs, errs = fut.result()
            except Exception as e:  # noqa: BLE001 - keep going, gap is retried next run
                segs, errs = [], [((a, b), e)]
            for sa, sb, rows in segs:
                store.save(sec, sa, sb, rows)   # DB writes stay on the main thread
            nrows = sum(len(s[2]) for s in segs)
            if segs and (sec["kind"] == "detail" or done % 100 == 0 or done == len(jobs)):
                log.info("[%d/%d] %-22s %s..%s  %6d rows", done, len(jobs), sec["name"], a, b, nrows)
            failures += 1 if errs else 0
            for day, err in errs:
                if not breaker["open"] or "skipped" not in str(err):
                    log.warning("[%d/%d] FAILED %s %s: %s (will retry next run)",
                                done, len(jobs), sec["name"],
                                f"{day[0]}..{day[1]}" if isinstance(day, tuple) else (day or ""), err)
    log.info("Download finished in %.0fs, %d failure(s).", time.time() - t0, failures)
    return failures


# --------------------------------------------------------------------------
# Excel export
# --------------------------------------------------------------------------
def tidy(sec, df: pd.DataFrame) -> pd.DataFrame:
    cols = sec["columns"]
    df = df.copy()
    if sec["kind"] == "summary":
        df = df[["_date"] + [c for c in cols if c != "_date"]]
    else:
        df = df[list(cols) + ["_date"]]
    # dates
    for c in df.columns:
        if "date" in c:
            df[c] = pd.to_datetime(df[c].map(parse_day), errors="coerce")
    # numbers (the site sometimes sends "360.00" strings or '-')
    for c in df.columns:
        if any(k in c for k in ("qty", "buy", "sell", "price", "volume")):
            df[c] = pd.to_numeric(df[c], errors="coerce")
    sort_cols = ["_date"] + [c for c in ("contract_type", "product_name", "product_code",
                                         "hours", "time_block") if c in df.columns]
    df = df.sort_values(sort_cols, kind="stable")
    if "product_code" in df.columns:
        df.insert(df.columns.get_loc("product_code") + 1, "_energy",
                  df["product_code"].map(energy_type))
    if sec["market"] == "GTAM" and "product_name" in df.columns:
        df.insert(df.columns.get_loc("product_name") + 1, "_energy",
                  df["product_name"].map(energy_type))
    if sec["kind"] == "detail":
        df = df.drop(columns=["_date"])
    labels = dict(cols)
    labels["_energy"] = "Energy Type"
    return df.rename(columns=labels).reset_index(drop=True)


def export_excel(store: Store, out_path: Path, start: dt.date, end: dt.date):
    try:
        import xlsxwriter  # noqa: F401
        engine = "xlsxwriter"
    except ImportError:
        engine = "openpyxl"
    tmp = out_path.with_name(out_path.stem + ".tmp.xlsx")
    overview = []
    log.info("Writing Excel (%s) ...", engine)

    with pd.ExcelWriter(tmp, engine=engine,
                        **({"datetime_format": "dd-mmm-yyyy", "date_format": "dd-mmm-yyyy"}
                           if engine == "xlsxwriter" else {})) as xw:
        # placeholder so Overview is the first sheet
        pd.DataFrame().to_excel(xw, sheet_name="Overview", index=False)
        for sec in SECTIONS:
            df = tidy(sec, store.frame(sec))
            first, last, days_logged, last_fetch = store.coverage(sec["name"])
            done = store.done_days(sec["name"])
            missing = sum(1 for d in daterange(start, end) if d.isoformat() not in done)
            parts = max(1, -(-len(df) // EXCEL_MAX_ROWS))
            sheets = []
            for p in range(parts):
                sname = sec["name"] if p == 0 else f"{sec['name'][:28]}_{p + 1}"
                part = df.iloc[p * EXCEL_MAX_ROWS:(p + 1) * EXCEL_MAX_ROWS]
                part.to_excel(xw, sheet_name=sname, index=False)
                sheets.append(sname)
                if engine == "xlsxwriter":
                    ws = xw.sheets[sname]
                    ws.freeze_panes(1, 0)
                    ws.autofilter(0, 0, max(len(part), 1), max(len(part.columns) - 1, 0))
                    for i, c in enumerate(part.columns):
                        ws.set_column(i, i, max(12, min(34, len(str(c)) + 2)))
            date_col = next(c for c in df.columns if "Date" in c)
            overview.append({
                "Market": sec["market"], "Sheet": ", ".join(sheets), "Section": sec["title"],
                "Rows": len(df),
                "First Date": df[date_col].min() if len(df) else None,
                "Last Date": df[date_col].max() if len(df) else None,
                "Days Checked": days_logged or 0,
                f"Missing Days ({start:%d-%b-%Y} to {end:%d-%b-%Y})": missing,
                "Last Downloaded (IST)": last_fetch,
                "Source Endpoint": f"{BASE_URL}/{sec['endpoint']}/",
            })

        ov = pd.DataFrame(overview)
        ov.to_excel(xw, sheet_name="Overview", index=False, startrow=3)
        if engine == "xlsxwriter":
            wb, ws = xw.book, xw.sheets["Overview"]
            ws.write(0, 0, "PXIL - TAM & GTAM market data", wb.add_format({"bold": True, "font_size": 14}))
            ws.write(1, 0, f"Source: {BASE_URL}/Pages/Market.html   |   Updated: "
                           f"{dt.datetime.now(IST):%d-%b-%Y %H:%M} IST   |   "
                           "Rows with no trades are not listed. Summary sheets = site's "
                           "Max/Min/Avg/Total per delivery day.")
            ws.freeze_panes(4, 0)
            widths = [8, 26, 62, 10, 13, 13, 13, 30, 22, 52]
            for i, w in enumerate(widths):
                ws.set_column(i, i, w)

    try:
        os.replace(tmp, out_path)
        return out_path
    except PermissionError:
        alt = out_path.with_name(f"{out_path.stem}_{dt.datetime.now():%Y%m%d_%H%M}.xlsx")
        os.replace(tmp, alt)
        log.warning("%s is open in Excel - saved as %s instead. Close the file and "
                    "re-run to update the master copy.", out_path.name, alt.name)
        return alt


# --------------------------------------------------------------------------
# Main
# --------------------------------------------------------------------------
def main(argv=None):
    ap = argparse.ArgumentParser(description="Download PXIL TAM + GTAM data into one Excel workbook.")
    ap.add_argument("--start", default=DEFAULT_START, help="back-fill from this date (YYYY-MM-DD)")
    ap.add_argument("--end", default=None, help="last delivery date (default: today + 1, IST)")
    ap.add_argument("--refresh-days", type=int, default=REFRESH_DAYS,
                    help="always re-download the latest N days (default 3)")
    ap.add_argument("--only", choices=["TAM", "GTAM"], help="limit to one market")
    ap.add_argument("--sections", help="comma list of sheet names, e.g. GTAM_DAC,TAM_DAC")
    ap.add_argument("--no-summaries", action="store_true", help="skip the *_Summary sheets")
    ap.add_argument("--redownload", nargs=2, metavar=("FROM", "TO"),
                    help="force re-download of a date range")
    ap.add_argument("--export-only", action="store_true", help="only rebuild the Excel from cache")
    ap.add_argument("--out", default=None, help="output folder (default: this script's folder)")
    ap.add_argument("--workers", type=int, default=WORKERS)
    ap.add_argument("--base-url", default=BASE_URL, help=argparse.SUPPRESS)
    args = ap.parse_args(argv)

    out_dir = Path(args.out) if args.out else Path(__file__).resolve().parent
    out_dir.mkdir(parents=True, exist_ok=True)
    logging.basicConfig(
        level=logging.INFO, format="%(asctime)s  %(levelname)-7s %(message)s", datefmt="%Y-%m-%d %H:%M:%S",
        handlers=[logging.StreamHandler(sys.stdout),
                  RotatingFileHandler(out_dir / LOG_NAME, maxBytes=2_000_000, backupCount=3,
                                      encoding="utf-8")])

    today = dt.datetime.now(IST).date()
    start = dt.date.fromisoformat(args.start)
    end = dt.date.fromisoformat(args.end) if args.end else today + dt.timedelta(days=DAYS_AHEAD)
    if start > end:
        ap.error("--start is after --end")
    refresh_from = max(start, end - dt.timedelta(days=max(args.refresh_days, 0) - 1)) \
        if args.refresh_days > 0 else end + dt.timedelta(days=1)
    force = None
    if args.redownload:
        force = (dt.date.fromisoformat(args.redownload[0]), dt.date.fromisoformat(args.redownload[1]))
        start = min(start, force[0])

    sections = SECTIONS
    if args.only:
        sections = [s for s in sections if s["market"] == args.only]
    if args.sections:
        wanted = {x.strip().lower() for x in args.sections.split(",")}
        sections = [s for s in sections if s["name"].lower() in wanted]
    if args.no_summaries:
        sections = [s for s in sections if s["kind"] != "summary"]

    store = Store(out_dir / DB_NAME)
    log.info("PXIL TAM/GTAM  |  range %s -> %s  |  refresh from %s  |  %d section(s)",
             start, end, refresh_from, len(sections))

    failures = 0
    if not args.export_only:
        jobs = plan_jobs(store, sections, start, end, refresh_from, force)
        failures = run_fetch(store, Client(args.base_url), jobs, args.workers)

    path = export_excel(store, out_dir / EXCEL_NAME, start, end)
    log.info("Excel ready: %s", path)
    if failures:
        log.warning("%d request(s) failed; those dates stay 'missing' and will be retried next run.",
                    failures)
    return 1 if failures else 0


if __name__ == "__main__":
    sys.exit(main())
