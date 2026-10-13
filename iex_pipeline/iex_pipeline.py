#!/usr/bin/env python3
"""
IEX (Indian Energy Exchange) market-data pipeline
=================================================

Keeps a local SQLite database up to date with the data behind iexindia.com's market-data pages.

Datasets (use the name with --markets, or a group name: snapshot | intraday | auctions | all):

  15-min market snapshot (REST JSON  /api/v1/{mkt}/market-snapshot)
    dam, rtm, gdam, hpdam      one row per delivery date x 15-min block

  Intra-Day / DAC / Term-Ahead - instrument-level trade details (Next.js page data)
    tam      /market-data/intra-day-ahead-contingency/trade-details
    gtam     /market-data/green-intra-day/g-market-data            (solarType=ALL)
    hptam    /market-data/high-price-intra-day/trade-details
             one row per trade date x contract type x instrument

  Intra-Day / DAC / Term-Ahead - daily totals incl. total buy/sell bid volume (REST JSON)
    tam_summary, gtam_summary  /api/v1/{tam|gtam}/market-snapshot?range=DAILY

  Auctions (list pages only; result PDFs are not downloaded)
    tam_anyday   /tam-anyday   TAM Any-Day(s) Single-Side Reverse Auction list
    tam_monthly  /tam-monthly  TAM Monthly auctions (volumes/prices in the list)

Usage:
    python iex_pipeline.py update                  # first run = full backfill, later runs = incremental
    python iex_pipeline.py update --markets intraday auctions
    python iex_pipeline.py backfill --markets tam --start 2020-04-01 --end 2020-12-31
    python iex_pipeline.py check
    python iex_pipeline.py export --out csv
"""

from __future__ import annotations

import argparse
import csv
import json
import logging
import os
import re
import sqlite3
import sys
import time
from datetime import date, datetime, timedelta, timezone
from pathlib import Path
from urllib.parse import urlencode, urlsplit

import requests
from requests.adapters import HTTPAdapter
from urllib3.util.retry import Retry

# --------------------------------------------------------------------------- config

SITE = os.environ.get("IEX_SITE", "https://www.iexindia.com").rstrip("/")
API = SITE + "/api/v1"
HERE = Path(__file__).resolve().parent
DB_PATH = Path(os.environ.get("IEX_DB_PATH", HERE / "data" / "iex_market_data.db"))
LOG_PATH = Path(os.environ.get("IEX_LOG_PATH", HERE / "logs" / "iex_pipeline.log"))

DEFAULT_START = date(2020, 4, 1)
REQUEST_PAUSE_SEC = float(os.environ.get("IEX_PAUSE", "1.0"))    # be polite to the portal

LOOKBACK = {"blocks": 5, "trades": 7, "summary": 7}  # days re-pulled on every incremental run

KEYS = {
    "blocks":  ["delivery_date", "block"],
    "trades":  ["trade_date", "contract_type", "instrument_name"],
    "summary": ["trade_date", "contract_type"],
}

# `launch` = first date worth requesting (earlier dates come back empty).
DATASETS: dict[str, dict] = {
    "dam":   dict(kind="blocks", api="dam",   page="day-ahead-market",            launch=date(2020, 4, 1),
                  label="Day Ahead Market"),
    "rtm":   dict(kind="blocks", api="rtm",   page="real-time-market",            launch=date(2020, 6, 1),
                  label="Real Time Market"),
    "gdam":  dict(kind="blocks", api="gdam",  page="green-day-ahead-market",      launch=date(2021, 10, 27),
                  label="Green Day Ahead Market"),
    "hpdam": dict(kind="blocks", api="hpdam", page="high-price-day-ahead-market", launch=date(2023, 3, 11),
                  label="High Price Day Ahead Market"),

    "tam":   dict(kind="trades", page="/market-data/intra-day-ahead-contingency/trade-details", extra={},
                  launch=date(2020, 4, 1), label="Intra Day, DAC, Term Ahead Market - trade details"),
    "gtam":  dict(kind="trades", page="/market-data/green-intra-day/g-market-data", extra={"solarType": "ALL"},
                  launch=date(2020, 8, 21), label="Green Intra Day, DAC, Term Ahead Market - trade details"),
    "hptam": dict(kind="trades", page="/market-data/high-price-intra-day/trade-details", extra={},
                  launch=date(2023, 1, 1), label="High Price Intra Day, DAC, Term Ahead Market - trade details"),

    "tam_summary":  dict(kind="summary", api="tam", extra={}, launch=date(2020, 4, 1),
                         label="Intra Day, DAC, TAM - daily totals per contract type"),
    "gtam_summary": dict(kind="summary", api="gtam", extra={"solarType": "ALL"}, launch=date(2020, 8, 21),
                         label="Green Intra Day, DAC, TAM - daily totals per contract type"),

    "tam_anyday":  dict(kind="docs", page="/tam-anyday",
                        keys=["auction_date", "auction_entity_name", "pdf_url"],
                        label="TAM - Any Day(s) Single Side Reverse Auction"),
    "tam_monthly": dict(kind="docs", page="/tam-monthly",
                        keys=["auction_date", "instrument_name", "delivery_month", "pdf_url"],
                        label="TAM - Monthly"),
}
GROUPS = {
    "snapshot": ["dam", "rtm", "gdam", "hpdam"],
    "intraday": ["tam", "gtam", "hptam", "tam_summary", "gtam_summary"],
    "auctions": ["tam_anyday", "tam_monthly"],
}
GROUPS["all"] = [m for g in ("snapshot", "intraday", "auctions") for m in GROUPS[g]]

TEXT_COLS = {"period", "congestion", "contract_type", "instrument_name", "auction_entity_name",
             "delivery_month", "pdf_url", "instr_product", "instr_resolution", "instr_suffix"}

HEADERS = {
    "User-Agent": ("Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
                   "(KHTML, like Gecko) Chrome/128.0 Safari/537.36"),
    "Accept": "application/json, text/plain, */*",
    "Accept-Language": "en-US,en;q=0.9",
}

log = logging.getLogger("iex")
ERRORS = 0            # failed chunks in this run (retried automatically on the next `update`)
USE_BROWSER = os.environ.get("IEX_USE_BROWSER", "0") == "1"


# --------------------------------------------------------------------------- helpers

def setup_logging(verbose: bool = False) -> None:
    LOG_PATH.parent.mkdir(parents=True, exist_ok=True)
    fmt = logging.Formatter("%(asctime)s %(levelname)-7s %(message)s", "%Y-%m-%d %H:%M:%S")
    log.setLevel(logging.DEBUG if verbose else logging.INFO)
    for h in (logging.StreamHandler(sys.stdout), logging.FileHandler(LOG_PATH, encoding="utf-8")):
        h.setFormatter(fmt)
        log.addHandler(h)


def ist_today() -> date:
    # IST = UTC+5:30, no DST. Avoids depending on the machine's timezone.
    return (datetime.now(timezone.utc) + timedelta(hours=5, minutes=30)).date()


def parse_iso(s: str) -> date:
    return datetime.strptime(s, "%Y-%m-%d").date()


def to_api(d: date) -> str:
    return d.strftime("%d-%m-%Y")


def dmy_to_iso(s: str) -> str:
    return datetime.strptime(s.strip(), "%d-%m-%Y").date().isoformat()


def snake(name: str) -> str:
    name = re.sub(r"([a-z0-9])([A-Z])", r"\1_\2", str(name))          # camelCase -> camel_Case
    return re.sub(r"[^a-z0-9]+", "_", name.lower()).strip("_")


def to_num(v):
    if v is None or (isinstance(v, str) and v.strip() in ("", "-", "NA", "N/A")):
        return None
    try:
        return float(str(v).replace(",", ""))
    except ValueError:
        return str(v)


def month_chunks(start: date, end: date):
    cur = start
    while cur <= end:
        nxt = (cur.replace(day=1) + timedelta(days=32)).replace(day=1)
        chunk_end = min(end, nxt - timedelta(days=1))
        yield cur, chunk_end
        cur = chunk_end + timedelta(days=1)


def expand_markets(names: list[str]) -> list[str]:
    out: list[str] = []
    for n in names:
        for m in GROUPS.get(n, [n]):
            if m not in DATASETS:
                raise SystemExit(f"unknown market {m!r}; choose from {sorted(DATASETS) + sorted(GROUPS)}")
            if m not in out:
                out.append(m)
    return out


# --------------------------------------------------------------------------- HTTP

class BlockedError(RuntimeError):
    pass


def make_session():
    if USE_BROWSER:
        log.info("using headless browser fetch (--browser / IEX_USE_BROWSER=1)")
        return BrowserSession()
    s = requests.Session()
    s.headers.update(HEADERS)
    retry = Retry(total=5, backoff_factor=2, status_forcelist=(429, 500, 502, 503, 504),
                  allowed_methods=("GET",), respect_retry_after_header=True)
    s.mount("https://", HTTPAdapter(max_retries=retry))
    s.mount("http://", HTTPAdapter(max_retries=retry))
    return s


class _Resp:
    def __init__(self, status: int, ctype: str, text: str = "", content: bytes = b""):
        self.status_code, self.text, self.content = status, text, content
        self.headers = {"Content-Type": ctype}

    def raise_for_status(self):
        if self.status_code >= 400:
            raise requests.HTTPError(f"HTTP {self.status_code}: {self.text[:200]}")

    def json(self):
        return json.loads(self.text)


class BrowserSession:
    """Fallback if the portal ever blocks plain Python requests: opens the portal in a headless
    Chromium (Playwright) and runs each call as a same-origin fetch() inside that page.
    Install once:  pip install playwright  &&  python -m playwright install chromium"""

    def __init__(self):
        from playwright.sync_api import sync_playwright
        self._pw = sync_playwright().start()
        self._browser = self._pw.chromium.launch(headless=True)
        self.page = self._browser.new_page(user_agent=HEADERS["User-Agent"])
        landing = SITE + ("/market-data/day-ahead-market/market-snapshot" if "iexindia.com" in SITE else "/")
        self.page.goto(landing, wait_until="domcontentloaded", timeout=120_000)

    def get(self, url, params=None, timeout=120, headers=None):
        if urlsplit(url).netloc != urlsplit(SITE).netloc:          # other host: no CORS limits
            r = self.page.context.request.get(url, params=params, timeout=timeout * 1000)
            return _Resp(r.status, r.headers.get("content-type", ""), "", r.body())
        full = url + ("?" + urlencode(params) if params else "")
        res = self.page.evaluate(
            """async ([u, h]) => { const r = await fetch(u, {headers: h});
                                   return [r.status, r.headers.get('content-type') || '', await r.text()]; }""",
            [full, {"Accept": "application/json", **(headers or {})}])
        return _Resp(res[0], res[1], res[2], res[2].encode("utf-8"))

    def close(self):
        self._browser.close()
        self._pw.stop()


def get_json(session, url: str, params: dict, what: str, referer: str | None = None) -> dict:
    r = session.get(url, params=params, timeout=120, headers={"Referer": referer or SITE + "/"})
    ctype = r.headers.get("Content-Type", "")
    if "json" not in ctype:
        raise BlockedError(f"{what}: got HTTP {r.status_code} with {ctype or 'no content-type'} instead of JSON "
                           f"(bot protection or site change?). First bytes: {r.text[:200]!r}")
    r.raise_for_status()
    return r.json()


def get_rsc(session, path: str, params: dict, what: str) -> str:
    """The market-data pages are Next.js server components. Requesting the page with the header
    `RSC: 1` returns the page's data payload (text/x-component) - the same thing the site's own
    front-end downloads when you click 'Update Report'."""
    r = session.get(SITE + path, params=params, timeout=180, headers={"RSC": "1", "Accept": "*/*",
                                                                     "Referer": SITE + path})
    ctype = r.headers.get("Content-Type", "")
    if "x-component" not in ctype:
        raise BlockedError(f"{what}: got HTTP {r.status_code} with {ctype or 'no content-type'} instead of "
                           f"text/x-component (bot protection or site change?). First bytes: {r.text[:200]!r}")
    r.raise_for_status()
    return r.text


# --------------------------------------------------------------------------- RSC payload parsing

_DEC = json.JSONDecoder()
_REF = re.compile(r"^\$([0-9a-f]+)$")


def _rsc_lines(text: str) -> dict[str, str]:
    lines = {}
    for ln in text.split("\n"):
        m = re.match(r"^([0-9a-f]+):(.*)$", ln)
        if m:
            lines[m.group(1)] = m.group(2)
    return lines


def _resolve(obj, lines, depth=0):
    """React Flight may outline repeated objects as "$<hex id>" references to another line."""
    if depth > 20:
        return obj
    if isinstance(obj, str):
        if obj.startswith("$$"):
            return obj[1:]
        m = _REF.match(obj)
        if m and m.group(1) in lines:
            try:
                return _resolve(json.loads(lines[m.group(1)]), lines, depth + 1)
            except ValueError:
                return obj
        return obj
    if isinstance(obj, list):
        return [_resolve(x, lines, depth + 1) for x in obj]
    if isinstance(obj, dict):
        return {k: _resolve(v, lines, depth + 1) for k, v in obj.items()}
    return obj


def rsc_array(text: str, key: str) -> list | None:
    """First JSON array stored under "key": in the payload (None if the key is absent)."""
    lines = None
    for m in re.finditer(r'"%s":\[' % re.escape(key), text):
        arr, _ = _DEC.raw_decode(text, m.end() - 1)
        if isinstance(arr, list):
            lines = lines or _rsc_lines(text)
            return _resolve(arr, lines)
    return None


# --------------------------------------------------------------------------- normalisers

PERIOD_RE = re.compile(r"^\s*(\d{1,2}):(\d{2})\s*-\s*(\d{1,2}):(\d{2})\s*$")
INSTR_RE = re.compile(r"^(DAC|ITD)-([BH])(\d{2})-([A-Z0-9]+)$")


def norm_block(market: str, rec: dict) -> dict:
    row: dict = {"delivery_date": dmy_to_iso(rec["date"])}
    m = PERIOD_RE.match(str(rec.get("period", "")))
    if not m:
        raise ValueError(f"{market}: unexpected period format {rec.get('period')!r}")
    sh, sm, eh, em = map(int, m.groups())
    row["block"] = (sh * 60 + sm) // 15 + 1                    # 1..96
    row["hour"] = sh + 1                                       # 1..24
    row["period"] = f"{sh:02d}:{sm:02d}-{eh:02d}:{em:02d}"     # same format for all markets
    for k, v in rec.items():
        if k in ("date", "period"):
            continue
        col = snake(k)
        row[col] = (None if v is None else str(v)) if col in TEXT_COLS else to_num(v)
    return row


def norm_trade(rec: dict) -> dict:
    row = {"trade_date": dmy_to_iso(rec["tradeDate"]),
           "contract_type": str(rec.get("contractType") or "").strip(),
           "instrument_name": str(rec.get("instrumentName") or "").strip()}
    for k, v in rec.items():
        if k in ("tradeDate", "contractType", "instrumentName"):
            continue
        row[snake(k)] = to_num(v)
    # DAC-B01-ER / ITD-H18-WR -> product, B(15-min block)|H(hour), block no., region / RE-type code
    m = INSTR_RE.match(row["instrument_name"])
    row["instr_product"], row["instr_resolution"], row["instr_block"], row["instr_suffix"] = (
        (m.group(1), "15MIN" if m.group(2) == "B" else "HOUR", int(m.group(3)), m.group(4)) if m
        else (None, None, None, None))
    return row


def norm_summary(rec: dict) -> dict:
    row = {"trade_date": dmy_to_iso(rec["tradeDate"]), "contract_type": str(rec.get("contractType") or "").strip()}
    for k, v in rec.items():
        if k not in ("tradeDate", "contractType"):
            row[snake(k)] = to_num(v)
    return row


# --------------------------------------------------------------------------- fetchers

def fetch_blocks(session, name: str, a: date, b: date) -> list[dict]:
    ds = DATASETS[name]
    payload = get_json(session, f"{API}/{ds['api']}/market-snapshot",
                       {"interval": "ONE_FOURTH_HOUR", "fromDate": to_api(a), "toDate": to_api(b)},
                       f"{name} {a}..{b}", referer=f"{SITE}/market-data/{ds['page']}/market-snapshot")
    data = payload.get("data") or []
    if not isinstance(data, list):
        raise RuntimeError(f"{name}: unexpected payload shape: {str(payload)[:300]}")
    return [norm_block(name, r) for r in data]


def fetch_trades(session, name: str, a: date, b: date) -> list[dict]:
    ds = DATASETS[name]
    params = {"reportType": "TRADE_DATE_WISE", "dp": "SELECT_RANGE", "fromDate": to_api(a),
              "toDate": to_api(b), "contractType": "ALL", **ds["extra"]}
    text = get_rsc(session, ds["page"], params, f"{name} {a}..{b}")
    data = rsc_array(text, "tableData")
    if data is None:
        raise RuntimeError(f"{name}: no tableData in page payload (page layout changed?)")
    return [norm_trade(r) for r in data if isinstance(r, dict)]


def fetch_summary(session, name: str, a: date, b: date) -> list[dict]:
    ds = DATASETS[name]
    params = {"reportType": "TRADE_DATE_WISE", "contractType": "ALL", "year": str(a.year), "range": "DAILY",
              "fromDate": to_api(a), "toDate": to_api(b), **ds["extra"]}
    payload = get_json(session, f"{API}/{ds['api']}/market-snapshot", params, f"{name} {a}..{b}")
    data = payload.get("data") or []
    if not isinstance(data, list):
        raise RuntimeError(f"{name}: unexpected payload shape: {str(payload)[:300]}")
    return [norm_summary(r) for r in data]


def fetch_docs(session, name: str, page_size: int = 5000) -> list[dict]:
    ds = DATASETS[name]
    out, page = [], 0
    while True:
        text = get_rsc(session, ds["page"], {"page": page, "limit": page_size}, f"{name} page {page}")
        heads, rows = rsc_array(text, "tableHeads"), rsc_array(text, "tableRows")
        if heads is None or rows is None:
            raise RuntimeError(f"{name}: no tableHeads/tableRows in page payload (page layout changed?)")
        cols = [snake(h) for h in heads]
        for r in rows:
            rec = {}
            for col, cell in zip(cols, r):
                if isinstance(cell, dict) and "link" in cell:
                    rec["pdf_url"] = cell["link"]
                elif col == "auction_date":
                    rec[col] = dmy_to_iso(str(cell))
                elif col in TEXT_COLS:
                    rec[col] = None if cell is None else str(cell).strip()
                else:
                    rec[col] = to_num(cell)
            for k in ds["keys"]:
                rec.setdefault(k, "")
            out.append(rec)
        if len(rows) < page_size:
            return out
        page += 1
        time.sleep(REQUEST_PAUSE_SEC)


FETCHERS = {"blocks": fetch_blocks, "trades": fetch_trades, "summary": fetch_summary}


# --------------------------------------------------------------------------- storage

def connect() -> sqlite3.Connection:
    DB_PATH.parent.mkdir(parents=True, exist_ok=True)
    con = sqlite3.connect(DB_PATH)
    con.execute("PRAGMA journal_mode=WAL")
    con.execute("""CREATE TABLE IF NOT EXISTS load_log(
        run_at TEXT, market TEXT, from_date TEXT, to_date TEXT, rows INTEGER, status TEXT, message TEXT)""")
    return con


def table_keys(name: str) -> list[str]:
    ds = DATASETS[name]
    return ds.get("keys") or KEYS[ds["kind"]]


def table_exists(con, name: str) -> bool:
    return con.execute("SELECT 1 FROM sqlite_master WHERE type='table' AND name=?", (name,)).fetchone() is not None


def ensure_table(con: sqlite3.Connection, name: str, rows: list[dict]) -> None:
    keys = table_keys(name)
    keydefs = ", ".join(f'"{k}" {"INTEGER" if k == "block" else "TEXT"} NOT NULL' for k in keys)
    extra = ', hour INTEGER, period TEXT' if DATASETS[name]["kind"] == "blocks" else ""
    con.execute(f'CREATE TABLE IF NOT EXISTS "{name}" ({keydefs}{extra}, loaded_at TEXT, '
                f'PRIMARY KEY ({", ".join(keys)}))')
    existing = {r[1] for r in con.execute(f'PRAGMA table_info("{name}")')}
    seen: dict = {}
    for row in rows:
        for col, val in row.items():
            if col not in existing and (col not in seen or seen[col] is None):
                seen[col] = val
    for col, val in seen.items():
        sqltype = ("TEXT" if col in TEXT_COLS or isinstance(val, str)
                   else "INTEGER" if col in ("instr_block",) else "REAL")
        con.execute(f'ALTER TABLE "{name}" ADD COLUMN "{col}" {sqltype}')
        log.debug("  %s: added column %s (%s)", name, col, sqltype)


def upsert(con: sqlite3.Connection, name: str, rows: list[dict]) -> int:
    if not rows:
        return 0
    ensure_table(con, name, rows)
    keys = table_keys(name)
    stamp = datetime.now().isoformat(timespec="seconds")
    cols = sorted({c for r in rows for c in r} | {"loaded_at"})
    collist = ",".join(f'"{c}"' for c in cols)
    updates = ",".join(f'"{c}"=excluded."{c}"' for c in cols if c not in keys)
    sql = (f'INSERT INTO "{name}" ({collist}) VALUES ({",".join("?" for _ in cols)}) '
           f'ON CONFLICT({", ".join(keys)}) DO UPDATE SET {updates}')
    con.executemany(sql, [[{**r, "loaded_at": stamp}.get(c) for c in cols] for r in rows])
    con.commit()
    return len(rows)


def log_run(con, market, start, end, rows, status, msg=""):
    con.execute("INSERT INTO load_log VALUES (?,?,?,?,?,?,?)",
                (datetime.now().isoformat(timespec="seconds"), market,
                 start.isoformat() if start else None, end.isoformat() if end else None,
                 rows, status, msg[:500]))
    con.commit()


def last_complete_block_day(con, name: str) -> date | None:
    if not table_exists(con, name):
        return None
    row = con.execute(f'SELECT MAX(delivery_date) FROM (SELECT delivery_date FROM "{name}" '
                      f'GROUP BY delivery_date HAVING COUNT(*) >= 96)').fetchone()
    return parse_iso(row[0]) if row and row[0] else None


def last_ok_date(con, name: str) -> date | None:
    row = con.execute("SELECT MAX(to_date) FROM load_log WHERE market=? AND status='ok'", (name,)).fetchone()
    return parse_iso(row[0]) if row and row[0] else None


def failed_chunks(con, name: str) -> list[tuple[date, date]]:
    """Chunks that failed and were not re-loaded successfully afterwards."""
    rows = con.execute("""
        SELECT DISTINCT f.from_date, f.to_date FROM load_log f
        WHERE f.market=? AND f.status IN ('error','blocked') AND f.from_date IS NOT NULL
          AND NOT EXISTS (SELECT 1 FROM load_log o WHERE o.market=f.market AND o.status='ok'
                          AND o.from_date<=f.from_date AND o.to_date>=f.to_date AND o.run_at>=f.run_at)
        ORDER BY 1""", (name,)).fetchall()
    return [(parse_iso(a), parse_iso(b)) for a, b in rows]


def block_gaps(con, name: str, before: date) -> tuple[list[date], list[tuple[date, int]]]:
    """Days (launch .. before-1) with no rows, and days with fewer than 96 blocks."""
    if not table_exists(con, name):
        return [], []
    have = dict(con.execute(f'SELECT delivery_date, COUNT(*) FROM "{name}" GROUP BY delivery_date'))
    if not have:
        return [], []
    missing, partial = [], []
    d = max(DATASETS[name]["launch"], parse_iso(min(have)))
    while d < before:
        n = have.get(d.isoformat(), 0)
        if n == 0:
            missing.append(d)
        elif n < 96:
            partial.append((d, n))
        d += timedelta(days=1)
    return missing, partial


def refresh_views(con) -> None:
    """Convenience views for analysis (rebuilt every run from whichever tables exist)."""
    parts = [f"""SELECT '{t}' AS segment, trade_date, contract_type, COUNT(*) AS instruments,
                 SUM(number_of_trades) AS trades, ROUND(SUM(total_traded_volume), 2) AS traded_mwh,
                 ROUND(SUM(total_traded_volume * weighted_average_price) / NULLIF(SUM(total_traded_volume), 0), 2)
                   AS vwap_rs_mwh, MAX(highest_price) AS max_price_rs_mwh
               FROM "{t}" GROUP BY trade_date, contract_type"""
             for t in ("tam", "gtam", "hptam") if table_exists(con, t)]
    con.execute("DROP VIEW IF EXISTS v_intraday_daily")
    if parts:
        con.execute("CREATE VIEW v_intraday_daily AS " + " UNION ALL ".join(parts))
    blk = [f"""SELECT '{t}' AS segment, trade_date, instr_product, instr_resolution, instr_block,
                 ROUND(SUM(total_traded_volume), 2) AS traded_mwh,
                 ROUND(SUM(total_traded_volume * weighted_average_price) / NULLIF(SUM(total_traded_volume), 0), 2)
                   AS vwap_rs_mwh
               FROM "{t}" WHERE instr_product IS NOT NULL
               GROUP BY trade_date, instr_product, instr_resolution, instr_block"""
           for t in ("tam", "gtam", "hptam") if table_exists(con, t)]
    con.execute("DROP VIEW IF EXISTS v_dac_itd_blocks")
    if blk:
        con.execute("CREATE VIEW v_dac_itd_blocks AS " + " UNION ALL ".join(blk))
    con.commit()


# --------------------------------------------------------------------------- jobs

def load_range(session, con, name: str, start: date, end: date) -> int:
    global ERRORS
    ds = DATASETS[name]
    start = max(start, ds["launch"])
    if start > end:
        log.info("%s: nothing to load (data starts %s)", name, ds["launch"])
        return 0
    fetch, key0, total = FETCHERS[ds["kind"]], table_keys(name)[0], 0
    for a, b in month_chunks(start, end):
        try:
            rows = fetch(session, name, a, b)
            n = upsert(con, name, rows)
            total += n
            log.info("%-12s %s..%s  %6d rows  %3d days", name, a, b, n, len({r[key0] for r in rows}))
            log_run(con, name, a, b, n, "ok")
        except BlockedError:
            log_run(con, name, a, b, 0, "blocked")
            raise
        except Exception as e:  # keep going with the next chunk, record the failure
            ERRORS += 1
            log.error("%-12s %s..%s  FAILED: %s", name, a, b, e)
            log_run(con, name, a, b, 0, "error", str(e))
        time.sleep(REQUEST_PAUSE_SEC)
    return total


def load_docs(session, con, name: str) -> None:
    global ERRORS
    try:
        rows = fetch_docs(session, name)
        n = upsert(con, name, rows)
        dates = sorted(r["auction_date"] for r in rows if r.get("auction_date"))
        log.info("%-12s %6d rows  %s .. %s", name, n, dates[0] if dates else "-", dates[-1] if dates else "-")
        log_run(con, name, None, None, n, "ok")
    except BlockedError:
        log_run(con, name, None, None, 0, "blocked")
        raise
    except Exception as e:
        ERRORS += 1
        log.error("%-12s FAILED: %s", name, e)
        log_run(con, name, None, None, 0, "error", str(e))


def cmd_backfill(args):
    start = parse_iso(args.start) if args.start else DEFAULT_START
    session, con = make_session(), connect()
    for m in expand_markets(args.markets):
        ds = DATASETS[m]
        if ds["kind"] == "docs":
            load_docs(session, con, m)
            continue
        end = parse_iso(args.end) if args.end else ist_today() + timedelta(days=1 if ds["kind"] == "blocks" else 0)
        log.info("=== backfill %s (%s) %s -> %s", m, ds["label"], start, end)
        load_range(session, con, m, start, end)
    refresh_views(con)


def cmd_update(args):
    """Incremental. Block markets: from (last complete day - 5) up to tomorrow (DAM/GDAM/HP-DAM results
    for D+1 are out ~13:00 IST). Trade/summary markets: last loaded date - 7 days up to today.
    Empty dataset -> full backfill from --start. Failed chunks from earlier runs are retried."""
    session, con = make_session(), connect()
    today = ist_today()
    default_start = parse_iso(args.start) if args.start else DEFAULT_START
    for m in expand_markets(args.markets):
        ds, kind = DATASETS[m], DATASETS[m]["kind"]
        if kind == "docs":
            log.info("=== %s (%s)", m, ds["label"])
            load_docs(session, con, m)
            continue

        lookback = args.lookback if args.lookback is not None else LOOKBACK[kind]
        if kind == "blocks":
            last, end = last_complete_block_day(con, m), today + timedelta(days=1)
        else:
            last, end = last_ok_date(con, m), today
        start = default_start if last is None else min(last, end) - timedelta(days=lookback)
        log.info("=== %s %s: %s -> %s", "backfill" if last is None else "incremental", m, start, end)
        load_range(session, con, m, start, end)

        # self-heal older problems (e.g. a chunk that failed on a previous run)
        if kind == "blocks":
            missing, partial = block_gaps(con, m, before=start)
            bad = sorted(set(missing) | {d for d, _ in partial})
            months = sorted({d.replace(day=1) for d in bad})
            if bad:
                log.info("  %s: %d incomplete older days in %d month(s) -> re-fetching", m, len(bad), len(months))
            for mo in months[:args.max_gap_months]:
                days_in = [d for d in bad if d.replace(day=1) == mo]
                load_range(session, con, m, min(days_in), max(days_in))
        else:
            redo = [c for c in failed_chunks(con, m) if c[0] < start]
            if redo:
                log.info("  %s: %d previously failed chunk(s) -> re-fetching", m, len(redo))
            for a, b in redo[:args.max_gap_months]:
                load_range(session, con, m, a, b)
    refresh_views(con)


def cmd_check(args):
    con = connect()
    today = ist_today()
    for m in expand_markets(args.markets):
        kind = DATASETS[m]["kind"]
        if not table_exists(con, m):
            print(f"{m:12s} no table yet")
            continue
        k0 = table_keys(m)[0]
        nrows, ndays, first, last = con.execute(
            f'SELECT COUNT(*), COUNT(DISTINCT {k0}), MIN({k0}), MAX({k0}) FROM "{m}"').fetchone()
        line = f"{m:12s} {nrows:8d} rows  {ndays:5d} dates  {first} .. {last}"
        if kind == "blocks":
            missing, partial = block_gaps(con, m, before=today)   # today/tomorrow may be incomplete
            print(f"{line}  missing days: {len(missing)}  partial days: {len(partial)}")
            if missing:
                s = [d.isoformat() for d in missing]
                print(f"             missing: {', '.join(s[:15])}{' ...' if len(s) > 15 else ''}")
            if partial:
                s = [f"{d.isoformat()}({n})" for d, n in partial]
                print(f"             partial: {', '.join(s[:15])}{' ...' if len(s) > 15 else ''}")
        elif kind == "docs":
            print(line)
        else:
            bad = failed_chunks(con, m)
            print(f"{line}  failed chunks pending retry: {len(bad)}"
                  + (f"  ({', '.join(f'{a}..{b}' for a, b in bad[:6])})" if bad else ""))


def cmd_export(args):
    con = connect()
    out = Path(args.out)
    for m in expand_markets(args.markets):
        if not table_exists(con, m):
            continue
        keys = table_keys(m)
        cur = con.execute(f'SELECT * FROM "{m}" ORDER BY {", ".join(keys)}')
        cols = [c[0] for c in cur.description]
        k0 = cols.index(keys[0])
        writers, files = {}, []
        for row in cur:
            year = str(row[k0])[:4]
            if year not in writers:
                path = out / m / f"{m}_{year}.csv"
                path.parent.mkdir(parents=True, exist_ok=True)
                f = open(path, "w", newline="", encoding="utf-8")
                files.append(f)
                writers[year] = csv.writer(f)
                writers[year].writerow(cols)
            writers[year].writerow(row)
        for f in files:
            f.close()
        log.info("exported %s -> %s (%d files)", m, out / m, len(files))


# --------------------------------------------------------------------------- CLI

def main(argv=None):
    p = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    p.add_argument("-v", "--verbose", action="store_true")
    p.add_argument("--browser", action="store_true",
                   help="fetch through headless Chromium (Playwright) instead of plain HTTP")
    sub = p.add_subparsers(dest="cmd", required=True)
    choices = list(DATASETS) + list(GROUPS)

    def add_markets(sp):
        sp.add_argument("--markets", nargs="+", choices=choices, default=["all"],
                        help="dataset names or groups: snapshot, intraday, auctions, all (default)")

    sp = sub.add_parser("update", help="incremental update (full backfill on first run)")
    add_markets(sp)
    sp.add_argument("--start", help="first date for an empty dataset, YYYY-MM-DD (default 2020-04-01)")
    sp.add_argument("--lookback", type=int, default=None, help="days to re-pull (default 5 blocks / 7 trades)")
    sp.add_argument("--max-gap-months", type=int, default=12,
                    help="max number of older months with gaps to re-fetch per run")
    sp.set_defaults(func=cmd_update)

    sp = sub.add_parser("backfill", help="(re)load an explicit date range")
    add_markets(sp)
    sp.add_argument("--start", help="YYYY-MM-DD (default 2020-04-01)")
    sp.add_argument("--end", help="YYYY-MM-DD (default tomorrow IST for block markets, today otherwise)")
    sp.set_defaults(func=cmd_backfill)

    sp = sub.add_parser("check", help="completeness report")
    add_markets(sp)
    sp.set_defaults(func=cmd_check)

    sp = sub.add_parser("export", help="export tables to CSV (one file per dataset per year)")
    add_markets(sp)
    sp.add_argument("--out", default=str(HERE / "data" / "csv"))
    sp.set_defaults(func=cmd_export)

    args = p.parse_args(argv)
    setup_logging(args.verbose)
    global USE_BROWSER
    USE_BROWSER = USE_BROWSER or args.browser
    try:
        args.func(args)
    except BlockedError as e:
        log.error("%s", e)
        log.error("The portal refused a script request. See README 'If requests get blocked'.")
        return 2
    if ERRORS:
        log.warning("%d request(s) failed this run; the next `update` will retry them.", ERRORS)
        return 1
    return 0


if __name__ == "__main__":
    sys.exit(main())
