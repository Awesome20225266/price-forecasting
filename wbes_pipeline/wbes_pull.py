"""
WBES (Grid-India) data puller: PX + Net Schedule. Browser only for the OTP login.
=================================================================================

Two jobs, one login
-------------------
* px : Power Exchange dashboard. All PX (IEX, PXI, HPX) x all markets, latest
       revision for yesterday / today / tomorrow, re-checked every POLL_MINUTES.
* ns : Net Schedule report. All regions x all utilities, LATEST revision of every
       date. Each run makes sure the database is complete up to yesterday (D-1):
         - dates never downloaded (back to NS_START_DATE) are filled in, newest first;
         - the last NS_RECHECK_DAYS days are re-checked (1 cheap call per date), because
           the portal keeps revising past days for several days (seen: 26-Sep got
           revision 152 on 30-Sep). A date is re-downloaded only if its revision went up;
         - today is skipped (still changing all day), unless NS_INCLUDE_TODAY = True;
         - if the run is stopped halfway, the next run continues where it left off.
       Portal data starts around mid-2024. One date = ~775 API calls (~5 min).

Run
---
    python wbes_pull.py                         # both jobs, keeps running (PX every 15 min, NS every 3 h)
    python wbes_pull.py --once                  # both jobs, one pass (what the daily task runs)
    python wbes_pull.py --job ns --once         # Net Schedule only: bring DB up to D-1, then stop
    python wbes_pull.py --job ns --once --from 01-09-2026 --to 30-09-2026   # NS for a date range
    python wbes_pull.py --job px                # PX only
    run_pipeline.bat                            # same as --once; Task Scheduler runs this daily at 18:00

Why no inactivity logout
------------------------
* The portal's inactivity logout (15 min idle + 2 min warning) is JavaScript that
  runs in the browser tab. This script does not keep a portal tab open, so that
  timer never runs. There is nothing to "keep alive" except the login token.
* The login token (JWT) lasts about 10 minutes. The portal renews it with
      POST https://newwbes-api.grid-india.in/auth/api/Auth/refresh-token
      body {"Token": <RefreshToken>, "UtilityId": <utility id>}
  and gets back a new Token and a new RefreshToken (the refresh token rotates).
  This script does the same thing every ~8 minutes, so the session keeps going
  without OTP until the server itself ends it (refresh-token max age, portal
  maintenance, password change, or a login somewhere else).
* When that happens the script opens a Chrome window and waits for YOU to log in
  (username, OTP, captcha). It can alert you on Teams/Slack. It reads the new
  tokens, closes the window and carries on. It logs how long each session lasted,
  so you learn how often OTP is actually needed.

What gets saved on every run
----------------------------
* Excel: wbes_px_data/<date>/<PX>_<MARKET>_<date>_revNNN.xlsx  (Info, Region, Link, Entity sheets)
* Database (SQLite file wbes_px.db by default, or SQL Server / PostgreSQL / MySQL via DB_URL):
      px_load    one row per date+PX+market: which revision is stored, when it was pulled
      px_region  date, px, market, revision, time_block, time_desc, region, injection, drawal
      px_link    date, px, market, revision, time_block, time_desc, link_name, net_amount
      px_entity  date, px, market, revision, time_block, time_desc, entity, unit, px_code, region, injection, drawal
      ns_schedule     one row per date + region + utility + schedule type (ISGS, URS, OA_GNA,
                      OA_TGNA, OA_PX, OA_REMC, AS, SCED, SCUC, Total) with 96 block columns b01..b96
      ns_util_status  which utilities are done for a date (for resuming)
      ns_day          one row per completed date: revision stored, number of utilities
* Net Schedule Excel: wbes_ns_data/<yyyy-mm>/NetSchedule_<date>_revNNN.xlsx, one sheet per region
  Only the latest revision is kept. When a higher revision appears, the old rows for
  that date+PX+market are replaced in one transaction.
* Excel and DB are checked separately, so if one is missing (e.g. you add the DB
  later) it's filled on the next run without re-downloading duplicates.

Which browser opens?
--------------------
Set BROWSER below ("chrome", "msedge" or "chromium"). The script prints exactly which
browser and version it opened. It's always a separate window with a blank profile
(no bookmarks, not your normal logged-in Chrome). Chrome and Edge show a bar saying
"controlled by automated test software", which is how you can tell the script's window.

Setup (one time)
----------------
    pip install playwright pandas openpyxl requests sqlalchemy
    playwright install chromium      # only needed if BROWSER = "chromium" or Chrome/Edge are not installed

Important
---------
* wbes_session.json holds your live login tokens. Keep it private, never email or
  commit it. Delete it to force a fresh login.
* Don't use the same WBES user in a normal browser while this runs. That tab would
  refresh with the same rotating refresh token (one of the two gets kicked out).
  When its idle timer fires it also calls revoke-token.
* Run it on a machine that doesn't go to sleep. If the token lapses for longer
  than ~10 minutes, the refresh may be refused and you'll need to log in again.
"""

from __future__ import annotations

import argparse
import base64
import datetime as dt
import json
import os
import re
import sys
import time
from pathlib import Path

HERE = Path(__file__).resolve().parent

import pandas as pd
import requests
from sqlalchemy import (Column, Date, DateTime, Float, Index, Integer, MetaData, PrimaryKeyConstraint,
                        String, Table, create_engine, select)

# ----------------------------------------------------------------------------- config
PORTAL_URL = "https://newwbes.grid-india.in/site/open-access/px"
PORTAL_ORIGIN = "https://newwbes.grid-india.in"
API_BASE = os.environ.get("WBES_API_BASE", "https://newwbes-api.grid-india.in")

PX_NAMES: list[str] | None = None  # None = ALL exchanges the portal lists (IEX, PXI, HPX); or e.g. ["IEX"]
MARKETS: list[str] | None = None   # None = all markets the portal lists (DAM, RTM, GDAM, HPDAM)

# Which browser opens for the login (only for the OTP step; it closes afterwards):
#   "chrome"   = your installed Google Chrome   (separate window + blank profile, NOT your normal Chrome)
#   "msedge"   = your installed Microsoft Edge  (same, separate window + blank profile)
#   "chromium" = Playwright's own built-in Chromium (needs: playwright install chromium)
# If the chosen one is not installed, the script falls back to "chromium" and says so.
BROWSER = "chrome"

# Where data is saved on every run (both on by default)
SAVE_EXCEL = True
SAVE_DB = True
# SQLite = a single file, no setup. For a server database change the URL, e.g.
#   SQL Server : "mssql+pyodbc://USER:PASS@SERVER/DBNAME?driver=ODBC+Driver+17+for+SQL+Server"  (pip install pyodbc)
#   PostgreSQL : "postgresql+psycopg2://USER:PASS@HOST:5432/DBNAME"                              (pip install psycopg2-binary)
#   MySQL      : "mysql+pymysql://USER:PASS@HOST:3306/DBNAME"                                     (pip install pymysql)
DB_URL = os.environ.get(
    "WBES_DB_URL",
    "sqlite:///" + (HERE / "data" / "wbes.db").resolve().as_posix(),
)

# Which jobs run by default ("px", "ns"); override with --job
JOBS = ["px", "ns"]

# Net Schedule settings
NS_START_DATE = "01-06-2024"       # earliest date to keep complete in the DB (portal data starts ~mid-2024)
NS_RECHECK_DAYS = 7                # re-check this many recent days for late revisions
NS_INCLUDE_TODAY = False           # today keeps changing all day; leave False to stop at D-1
NS_REGIONS = [1, 2, 3, 4, 5, 6]    # ER, WR, NR, SR, NER, NLDC
NS_PAUSE_SEC = 0.2                 # pause between calls (be gentle with the portal)
NS_EVERY_MINUTES = 180             # in continuous mode, run the NS sync every 3 hours (PX still every POLL_MINUTES)
NS_SKIP_ZERO_ROWS = True           # don't store schedule-type rows that are zero in all 96 blocks
                                   # (each utility's "Total" row is always kept, so every utility appears)
NS_SAVE_EXCEL = True
NS_OUT_DIR = HERE / "wbes_ns_data"
DAY_OFFSETS = [-1, 0, 1]           # yesterday (final RTM), today, tomorrow (day-ahead)
POLL_MINUTES = 15
REFRESH_MARGIN_SEC = 150           # renew the token when it has less than this left
KEEP_ONLY_LATEST = True
SAVE_RAW_JSON = True

UTILITY_ID: str | None = None      # auto-read from the portal at login; or hard-code,
                                   # e.g. "32BBDF0D01463ADCE065000000000001" (GSPL_BHDL2)
NOTIFY_WEBHOOK = ""                # optional Slack / Teams incoming-webhook URL for "login needed" alerts
LOGIN_WAIT_MINUTES = 60            # how long the login window waits for you

OUT_DIR = HERE / "wbes_px_data"
SESSION_FILE = HERE / "wbes_session.json"

REGIONS = {1: "ER", 2: "WR", 3: "NR", 4: "SR", 5: "NER", 6: "NLDC"}


class LoginNeeded(Exception):
    pass


def log(msg: str) -> None:
    print(f"[{dt.datetime.now():%Y-%m-%d %H:%M:%S}] {msg}", flush=True)


def notify(msg: str) -> None:
    log(msg)
    if NOTIFY_WEBHOOK:
        try:
            requests.post(NOTIFY_WEBHOOK, json={"text": msg}, timeout=10)
        except Exception as e:
            log(f"(notify failed: {e!r})")


# ----------------------------------------------------------------------------- session
def jwt_exp(token: str) -> float:
    """Expiry (epoch seconds) from a JWT; 0 if unreadable."""
    try:
        part = token.split(".")[1]
        part += "=" * (-len(part) % 4)
        return float(json.loads(base64.urlsafe_b64decode(part))["exp"])
    except Exception:
        return 0.0


class WbesSession:
    def __init__(self, token: str, refresh: str, utility_id: str, login_at: float | None = None):
        self.token, self.refresh, self.utility_id = token, refresh, utility_id
        self.login_at = login_at or time.time()
        self.http = requests.Session()
        self.http.headers.update({"Accept": "application/json", "Content-Type": "application/json",
                                  "Origin": PORTAL_ORIGIN, "Referer": PORTAL_ORIGIN + "/"})

    # -- persistence
    def save(self) -> None:
        SESSION_FILE.write_text(json.dumps({"token": self.token, "refresh": self.refresh,
                                            "utility_id": self.utility_id, "login_at": self.login_at}))
        try:
            os.chmod(SESSION_FILE, 0o600)
        except Exception:
            pass

    @classmethod
    def load(cls) -> "WbesSession | None":
        try:
            d = json.loads(SESSION_FILE.read_text())
            return cls(d["token"], d["refresh"], d["utility_id"], d.get("login_at"))
        except Exception:
            return None

    # -- token handling
    def refresh_now(self) -> bool:
        """True = renewed. False = server refused (login needed). Network errors are retried, then raised."""
        for attempt in range(1, 6):
            try:
                r = self.http.post(f"{API_BASE}/auth/api/Auth/refresh-token",
                                   headers={"Authorization": f"Bearer {self.token}"},
                                   json={"Token": self.refresh, "UtilityId": self.utility_id}, timeout=30)
                break
            except requests.RequestException as e:
                log(f"refresh: network error ({attempt}/5): {e!r}")
                if attempt == 5:
                    raise
                time.sleep(15)
        if r.status_code != 200:
            log(f"refresh refused: HTTP {r.status_code}")
            return False
        try:
            d = r.json()
            if d.get("IsExpired") or not d.get("Token") or not d.get("RefreshToken"):
                log("refresh refused: session expired on server")
                return False
            self.token, self.refresh = d["Token"], d["RefreshToken"]
        except ValueError:
            log("refresh: unexpected response")
            return False
        self.save()
        return True

    def ensure_fresh(self) -> None:
        exp = jwt_exp(self.token)
        if exp and exp - time.time() > REFRESH_MARGIN_SEC:
            return
        if not self.refresh_now():
            raise LoginNeeded("token refresh failed")

    # -- API
    def call(self, path: str, method: str = "GET", body: dict | None = None):
        self.ensure_fresh()
        for attempt in (1, 2):
            r = self.http.request(method, API_BASE + path, json=body, timeout=60,
                                  headers={"Authorization": f"Bearer {self.token}"})
            if r.status_code in (401, 403):
                if attempt == 1 and self.refresh_now():
                    continue
                raise LoginNeeded(f"HTTP {r.status_code} on {path}")
            r.raise_for_status()
            return r.json()


# ----------------------------------------------------------------------------- manual login (browser)
JS_READ_SESSION = """
() => {
  try {
    const u = JSON.parse(localStorage.getItem('currentUser') || 'null').user;
    const sel = [...document.querySelectorAll('select')].find(s => /^[0-9A-F]{32}$/i.test(s.value));
    return {token: u.Token || null, refresh: u.RefreshToken || null, utility: sel ? sel.value : null};
  } catch (e) { return null; }
}
"""


def browser_login() -> WbesSession:
    """Open a fresh Chrome window, wait for the human to log in, capture tokens, close."""
    from playwright.sync_api import sync_playwright

    names = {"chrome": "Google Chrome", "msedge": "Microsoft Edge", "chromium": "Playwright Chromium"}
    with sync_playwright() as p:
        used = BROWSER
        try:
            browser = p.chromium.launch(headless=False, **({} if BROWSER == "chromium" else {"channel": BROWSER}))
        except Exception as e:
            log(f"{names.get(BROWSER, BROWSER)} could not be started ({str(e).splitlines()[0]}); using Playwright Chromium")
            used = "chromium"
            browser = p.chromium.launch(headless=False)
        log(f"Opened {names.get(used, used)} {browser.version} (separate window, blank profile)")
        notify(f"WBES login needed: please log in (OTP) in the {names.get(used, used)} window the script opened.")
        ctx = browser.new_context(viewport={"width": 1280, "height": 850})   # fresh: no stale tokens
        page = ctx.new_page()
        deadline = time.time() + LOGIN_WAIT_MINUTES * 60
        while time.time() < deadline:
            try:
                page.goto(PORTAL_URL, wait_until="domcontentloaded", timeout=20000)
                break
            except Exception as e:
                log(f"Portal page not open yet ({str(e).splitlines()[0]}). Retrying in 15 s — log in once the page appears.")
                page.wait_for_timeout(15000)
        else:
            ctx.close()
            browser.close()
            raise SystemExit("Portal did not load before the login window timed out.")

        got, got_at, nudged = None, 0.0, False
        while time.time() < deadline:
            try:
                d = page.evaluate(JS_READ_SESSION)
            except Exception:          # page navigating
                d = None
            if d and d.get("token") and d.get("refresh"):
                got = d
                got_at = got_at or time.time()
                if d.get("utility") or UTILITY_ID or time.time() - got_at > 30:
                    break
                if not nudged and not page.url.rstrip("/").endswith("/px"):
                    nudged = True      # open the PX page so the header utility dropdown renders
                    try:
                        page.goto(PORTAL_URL, wait_until="domcontentloaded")
                    except Exception:
                        pass
            page.wait_for_timeout(3000)
        ctx.close()
        browser.close()

    if not got:
        raise SystemExit("Login was not completed in time.")
    utility = UTILITY_ID or got.get("utility")
    if not utility:
        raise SystemExit("Logged in, but could not read the Utility ID. Set UTILITY_ID at the top of the script.")
    s = WbesSession(got["token"], got["refresh"], utility)
    s.save()
    log("Login captured; browser closed. Continuing without a browser.")
    return s


def get_session() -> WbesSession:
    s = WbesSession.load()
    ok = False
    if s:
        try:
            ok = s.refresh_now()
        except requests.RequestException:
            ok = False
    if s and ok:
        log(f"Reusing saved session (logged in {(time.time() - s.login_at) / 3600:.1f} h ago).")
        return s
    return browser_login()


# ----------------------------------------------------------------------------- parsing (same as v1)
def block_desc(n: int) -> str:
    s, e = (n - 1) * 15, n * 15
    return f"{s // 60:02d}:{s % 60:02d}-{e // 60:02d}:{e % 60:02d}"


def _blocks_frame(n: int) -> pd.DataFrame:
    return pd.DataFrame({"Time Block": range(1, n + 1), "Time Desc": [block_desc(i) for i in range(1, n + 1)]})


def region_frame(rec: dict) -> pd.DataFrame:
    rows = rec.get("AllRegionPowerExJsonData", {}).get("RegionWiseJsonData", []) or []
    df = _blocks_frame(max((len(r.get("Injection") or []) for r in rows), default=96))
    for r in rows:
        name = REGIONS.get(r.get("Region"), f"Region{r.get('Region')}")
        df[f"{name} Injection"] = pd.Series(r.get("Injection") or [])
        df[f"{name} Drawal"] = pd.Series(r.get("Drawal") or [])
    return df


def link_frame(rec: dict) -> pd.DataFrame:
    rows = rec.get("AllLinkPowerExJsonData", {}).get("LinkWiseJsonData", []) or []
    df = _blocks_frame(max((len(r.get("NetAmount") or []) for r in rows), default=96))
    for r in rows:
        df[r.get("LinkName") or f"Link{r.get('LinkNumber')}"] = pd.Series(r.get("NetAmount") or [])
    return df


def entity_frame(rec: dict) -> pd.DataFrame:
    out = []
    for e in rec.get("AllEntityPowerExJsonData", {}).get("EntityWiseJsonData", []) or []:
        inj, drw = e.get("Injection") or [], e.get("Drawal") or []
        for i in range(max(len(inj), len(drw))):
            out.append({"Time Block": i + 1, "Time Desc": block_desc(i + 1),
                        "Entity": e.get("EntityAcronym"), "Unit": e.get("EntityUnitAcronym"),
                        "PX Code": e.get("EntityPxCode"),
                        "Region": REGIONS.get(e.get("EntityRegionId"), e.get("EntityRegionId")),
                        "Injection": inj[i] if i < len(inj) else None,
                        "Drawal": drw[i] if i < len(drw) else None})
    return pd.DataFrame(out)


def info_frame(rec: dict, px: str, market: str) -> pd.DataFrame:
    return pd.DataFrame([("Date", rec.get("EffectiveForDate")), ("PX", px), ("Market", market),
                         ("Revision", rec.get("RevisionNo")), ("PowerExId", rec.get("PowerExId")),
                         ("Created On", rec.get("CreatedOn")),
                         ("Pulled At", dt.datetime.now().isoformat(timespec="seconds"))],
                        columns=["Field", "Value"])


def save_record(rec: dict, px: str, market: str, out_dir: Path) -> Path | None:
    date_iso = dt.datetime.strptime(rec["EffectiveForDate"], "%d-%m-%Y").strftime("%Y-%m-%d")
    rev = int(rec["RevisionNo"])
    folder = out_dir / date_iso
    folder.mkdir(parents=True, exist_ok=True)
    stem = f"{px}_{market}_{date_iso}_rev{rev:03d}"
    xlsx = folder / f"{stem}.xlsx"
    if xlsx.exists():
        return None
    tmp = xlsx.with_suffix(".tmp.xlsx")
    with pd.ExcelWriter(tmp, engine="openpyxl") as xw:
        info_frame(rec, px, market).to_excel(xw, sheet_name="Info", index=False)
        region_frame(rec).to_excel(xw, sheet_name="Region", index=False)
        link_frame(rec).to_excel(xw, sheet_name="Link", index=False)
        entity_frame(rec).to_excel(xw, sheet_name="Entity", index=False)
    tmp.replace(xlsx)
    if SAVE_RAW_JSON:
        (folder / f"{stem}.json").write_text(json.dumps(rec), encoding="utf-8")
    if KEEP_ONLY_LATEST:
        pat = re.compile(rf"^{re.escape(px)}_{re.escape(market)}_{date_iso}_rev(\d+)\.(xlsx|json)$")
        for f in folder.iterdir():
            m = pat.match(f.name)
            if m and int(m.group(1)) < rev:
                f.unlink()
    return xlsx


# ----------------------------------------------------------------------------- database
_md = MetaData()
_KEY = lambda: [Column("trade_date", Date, nullable=False), Column("px", String(10), nullable=False),
                Column("market", String(10), nullable=False), Column("revision", Integer, nullable=False)]
_BLOCK = lambda: [Column("time_block", Integer, nullable=False), Column("time_desc", String(11))]

px_load = Table("px_load", _md, *_KEY(),
                Column("power_ex_id", String(64)), Column("created_on", String(32)), Column("pulled_at", DateTime),
                PrimaryKeyConstraint("trade_date", "px", "market"))
px_region = Table("px_region", _md, *_KEY(), *_BLOCK(), Column("region", String(8)),
                  Column("injection", Float), Column("drawal", Float),
                  Index("ix_px_region_key", "trade_date", "px", "market"))
px_link = Table("px_link", _md, *_KEY(), *_BLOCK(), Column("link_name", String(50)), Column("net_amount", Float),
                Index("ix_px_link_key", "trade_date", "px", "market"))
px_entity = Table("px_entity", _md, *_KEY(), *_BLOCK(), Column("entity", String(100)), Column("unit", String(100)),
                  Column("px_code", String(50)), Column("region", String(8)),
                  Column("injection", Float), Column("drawal", Float),
                  Index("ix_px_entity_key", "trade_date", "px", "market"))

_engine = None


def db():
    global _engine
    if _engine is None:
        if DB_URL.startswith("sqlite"):
            (HERE / "data").mkdir(parents=True, exist_ok=True)
        kw = {"fast_executemany": True} if DB_URL.startswith("mssql+pyodbc") else {}
        _engine = create_engine(DB_URL, **kw)
        _md.create_all(_engine)          # creates the tables the first time only
    return _engine


def _f(v):
    try:
        return None if v is None else float(v)
    except (TypeError, ValueError):
        return None


def db_revision(trade_date: dt.date, px: str, market: str) -> int | None:
    with db().connect() as c:
        return c.execute(select(px_load.c.revision).where(
            px_load.c.trade_date == trade_date, px_load.c.px == px, px_load.c.market == market)).scalar()


def save_db(rec: dict, px: str, market: str) -> bool:
    """Store the record; replaces any older revision for the same date+PX+market. False if already stored."""
    d = dt.datetime.strptime(rec["EffectiveForDate"], "%d-%m-%Y").date()
    rev = int(rec["RevisionNo"])
    have = db_revision(d, px, market)
    if have is not None and have >= rev:
        return False
    key = {"trade_date": d, "px": px, "market": market, "revision": rev}

    regions = []
    for r in rec.get("AllRegionPowerExJsonData", {}).get("RegionWiseJsonData", []) or []:
        inj, drw = r.get("Injection") or [], r.get("Drawal") or []
        name = REGIONS.get(r.get("Region"), str(r.get("Region")))
        for i in range(max(len(inj), len(drw))):
            regions.append({**key, "time_block": i + 1, "time_desc": block_desc(i + 1), "region": name,
                            "injection": _f(inj[i] if i < len(inj) else None),
                            "drawal": _f(drw[i] if i < len(drw) else None)})
    links = []
    for r in rec.get("AllLinkPowerExJsonData", {}).get("LinkWiseJsonData", []) or []:
        for i, v in enumerate(r.get("NetAmount") or []):
            links.append({**key, "time_block": i + 1, "time_desc": block_desc(i + 1),
                          "link_name": r.get("LinkName") or f"Link{r.get('LinkNumber')}", "net_amount": _f(v)})
    ents = []
    for e in rec.get("AllEntityPowerExJsonData", {}).get("EntityWiseJsonData", []) or []:
        inj, drw = e.get("Injection") or [], e.get("Drawal") or []
        reg = REGIONS.get(e.get("EntityRegionId"), str(e.get("EntityRegionId")))
        for i in range(max(len(inj), len(drw))):
            ents.append({**key, "time_block": i + 1, "time_desc": block_desc(i + 1),
                         "entity": e.get("EntityAcronym"), "unit": e.get("EntityUnitAcronym"),
                         "px_code": e.get("EntityPxCode"), "region": reg,
                         "injection": _f(inj[i] if i < len(inj) else None),
                         "drawal": _f(drw[i] if i < len(drw) else None)})

    with db().begin() as c:              # one transaction: old revision out, new revision in
        for t in (px_region, px_link, px_entity, px_load):
            c.execute(t.delete().where(t.c.trade_date == d, t.c.px == px, t.c.market == market))
        for t, rows in ((px_region, regions), (px_link, links), (px_entity, ents)):
            for i in range(0, len(rows), 5000):
                c.execute(t.insert(), rows[i:i + 5000])
        c.execute(px_load.insert(), {**key, "power_ex_id": rec.get("PowerExId"),
                                     "created_on": str(rec.get("CreatedOn")), "pulled_at": dt.datetime.now()})
    return True


# ----------------------------------------------------------------------------- pull
def pull_dates(s: WbesSession, dates: list[dt.date]) -> None:
    for d in dates:
        ds = d.strftime("%d-%m-%Y")
        pxs = {p["PowerExchangeTypeName"]: p["PowerExchangeTypeId"]
               for p in s.call(f"/powerexchange/api/PowerExchange/GetPxTypeList/date/{ds}")}
        txs = {t["PowerExTransactionTypeName"]: t["PowerExTransactionTypeId"]
               for t in s.call(f"/powerexchange/api/PowerExchange/GetPxTransTypeList/date/{ds}")}
        for px in (PX_NAMES or list(pxs)):
            if px not in pxs:
                log(f"{ds} {px}: not listed, skipping")
                continue
            for mkt, mkt_id in txs.items():
                if MARKETS and mkt not in MARKETS:
                    continue
                data = s.call("/powerexchange/api/PowerExchange/GetPowerExchangeByDetails", "POST",
                              {"InputDate": ds, "PowerExchangeTypeId": pxs[px],
                               "PowerExTransactionTypeId": mkt_id, "RevisionNo": -1})
                recs = data.get("ResponseBody") or []
                if not recs:
                    log(f"{ds} {px} {mkt}: not published yet")
                    continue
                rec = max(recs, key=lambda r: int(r.get("RevisionNo", -1)))
                done = []
                if SAVE_EXCEL and save_record(rec, px, mkt, OUT_DIR):
                    done.append("Excel")
                if SAVE_DB and save_db(rec, px, mkt):
                    done.append("DB")
                log(f"{ds} {px} {mkt}: rev {rec['RevisionNo']} " +
                    (f"saved to {' + '.join(done)}" if done else "already saved"))
                time.sleep(0.5)


def sleep_keeping_token(s: WbesSession, seconds: int) -> None:
    end = time.time() + seconds
    while time.time() < end:
        time.sleep(min(30, max(1, end - time.time())))
        s.ensure_fresh()                     # raises LoginNeeded if the server ends the session


# ============================================================================= NET SCHEDULE
NS_API = "/schedule/api/ScheduleReport/"
BLOCK_COLS = [f"b{i:02d}" for i in range(1, 97)]

ns_schedule = Table(
    "ns_schedule", _md,
    Column("trade_date", Date, nullable=False), Column("region", String(8), nullable=False),
    Column("utility_id", String(40), nullable=False), Column("utility", String(100)),
    Column("utility_type_id", Integer), Column("schedule_type_id", Integer, nullable=False),
    Column("schedule_type", String(30)), Column("direction_flag", Integer),
    Column("revision", Integer, nullable=False), Column("created_on", String(32)),
    *[Column(c, Float) for c in BLOCK_COLS],
    Index("ix_ns_sched_date_util", "trade_date", "region", "utility_id"))
ns_util_status = Table(
    "ns_util_status", _md,
    Column("trade_date", Date, nullable=False), Column("region", String(8), nullable=False),
    Column("utility_id", String(40), nullable=False), Column("utility", String(100)),
    Column("revision", Integer), Column("n_rows", Integer), Column("pulled_at", DateTime),
    PrimaryKeyConstraint("trade_date", "region", "utility_id"))
ns_day = Table(
    "ns_day", _md,
    Column("trade_date", Date, nullable=False), Column("revision", Integer),
    Column("n_utilities", Integer), Column("completed_at", DateTime),
    PrimaryKeyConstraint("trade_date"))


class NsError(Exception):
    pass


def ns_latest_rev(s: WbesSession, d: dt.date) -> int:
    j = s.call(f"{NS_API}GetLatestFullSchdRevNo/regionId/0/date/{d:%d-%m-%Y}")
    try:
        return int(j.get("ResponseBody"))
    except (TypeError, ValueError):
        return -1


def ns_utilities(s: WbesSession, region_id: int, d: dt.date) -> list[dict]:
    j = s.call(f"{NS_API}GetSchdFilteredUtilUnitList", "POST",
               {"UtilRegionId": str(region_id), "InputDate": f"{d:%d-%m-%Y}"})
    return ((j or {}).get("ResponseBody") or {}).get("UtilityList") or []


def ns_fetch(s: WbesSession, region_id: int, d: dt.date, utility_id: str, rev: int) -> list[dict]:
    j = s.call(f"{NS_API}GetFullScheduleTypeWiseSummaryDetails", "POST",
               {"UtilRegionId": region_id, "InputDate": f"{d:%d-%m-%Y}", "EnergyScheduleTypeId": 0,
                "NetInjDrwFlag": 0, "UtilIdList": [utility_id], "IsFlowAreaSelected": 0, "RevisionNo": rev})
    code = ((j or {}).get("ResponseStatus") or {}).get("Code")
    if code and code != "WBES-200":
        raise NsError(f"{code}: {(j.get('ResponseStatus') or {}).get('Message')}")
    return (j or {}).get("ResponseBody") or []


def ns_store(d: dt.date, region: str, rev: int, batch: list[tuple[dict, list[dict]]]) -> None:
    now = dt.datetime.now()
    with db().begin() as c:
        for u, items in batch:
            uid = u["UtilityId"]
            where = lambda t: t.delete().where(t.c.trade_date == d, t.c.region == region, t.c.utility_id == uid)
            c.execute(where(ns_schedule))
            c.execute(where(ns_util_status))
            rows = []
            for it in items:
                amt = list(it.get("FinalAmount") or [])[:96]
                amt += [None] * (96 - len(amt))
                if (NS_SKIP_ZERO_ROWS and it.get("EnergyScheduleTypeName") != "Total"
                        and all(not _f(v) for v in amt)):
                    continue
                rows.append({"trade_date": d, "region": region, "utility_id": uid, "utility": u.get("Acronym"),
                             "utility_type_id": u.get("UtilityTypeId"),
                             "schedule_type_id": it.get("EnergyScheduleTypeId"),
                             "schedule_type": it.get("EnergyScheduleTypeName"),
                             "direction_flag": it.get("DirectionFlag"), "revision": rev,
                             "created_on": str(it.get("CreatedOn")),
                             **{BLOCK_COLS[i]: _f(amt[i]) for i in range(96)}})
            if rows:
                c.execute(ns_schedule.insert(), rows)
            c.execute(ns_util_status.insert(), {"trade_date": d, "region": region, "utility_id": uid,
                                                "utility": u.get("Acronym"), "revision": rev,
                                                "n_rows": len(rows), "pulled_at": now})


def ns_export_excel(d: dt.date, rev: int) -> Path:
    with db().connect() as c:
        df = pd.read_sql(select(ns_schedule).where(ns_schedule.c.trade_date == d)
                         .order_by(ns_schedule.c.region, ns_schedule.c.utility, ns_schedule.c.schedule_type_id), c)
    folder = NS_OUT_DIR / f"{d:%Y-%m}"
    folder.mkdir(parents=True, exist_ok=True)
    xlsx = folder / f"NetSchedule_{d:%Y-%m-%d}_rev{rev:03d}.xlsx"
    tmp = xlsx.with_suffix(".tmp.xlsx")
    df = df.rename(columns={c: block_desc(i + 1) for i, c in enumerate(BLOCK_COLS)})
    keep = ["utility", "utility_type_id", "schedule_type"] + [block_desc(i) for i in range(1, 97)]
    with pd.ExcelWriter(tmp, engine="openpyxl") as xw:
        pd.DataFrame([("Date", f"{d:%d-%m-%Y}"), ("Revision", rev), ("Utilities", df["utility_id"].nunique()),
                      ("Exported At", dt.datetime.now().isoformat(timespec="seconds"))],
                     columns=["Field", "Value"]).to_excel(xw, sheet_name="Info", index=False)
        for region in [REGIONS[r] for r in NS_REGIONS if REGIONS.get(r) in set(df["region"])]:
            df.loc[df["region"] == region, keep].to_excel(xw, sheet_name=region, index=False)
    last_err = None
    for _ in range(8):
        try:
            if tmp.exists():
                tmp.replace(xlsx)
            if xlsx.exists():
                break
        except OSError as e:   # Excel or antivirus briefly locks or moves the new file
            last_err = e
            if xlsx.exists() and not tmp.exists():
                break
            time.sleep(1.5)
    else:
        log(f"{d:%d-%m-%Y} NS Excel save was interrupted; data is in the database ({last_err})")
        return tmp if tmp.exists() else xlsx
    for f in folder.glob(f"NetSchedule_{d:%Y-%m-%d}_rev*.xlsx"):
        if f != xlsx and ".tmp" not in f.name:
            try:
                f.unlink()
            except PermissionError:
                log(f"{d:%d-%m-%Y} NS could not remove older file {f.name} (it is open)")
    return xlsx


def ns_pull_date(s: WbesSession, d: dt.date, rev: int) -> bool:
    """Download every utility of every region for date d at revision rev. True if complete."""
    with db().connect() as c:
        done = {(r.region, r.utility_id) for r in c.execute(
            select(ns_util_status.c.region, ns_util_status.c.utility_id).where(
                ns_util_status.c.trade_date == d, ns_util_status.c.revision == rev))}
    t0, total, fetched, failed = time.time(), 0, 0, 0
    for region_id in NS_REGIONS:
        region = REGIONS.get(region_id, str(region_id))
        utils = ns_utilities(s, region_id, d)
        batch: list[tuple[dict, list[dict]]] = []
        for u in utils:
            total += 1
            if (region, u["UtilityId"]) in done:
                continue
            try:
                items = ns_fetch(s, region_id, d, u["UtilityId"], rev)
            except NsError as e:
                failed += 1
                log(f"{d:%d-%m-%Y} NS {region} {u.get('Acronym')}: {e} (will retry next run)")
                continue
            batch.append((u, items))
            fetched += 1
            if len(batch) >= 25:
                ns_store(d, region, rev, batch)
                batch = []
            time.sleep(NS_PAUSE_SEC)
        if batch:
            ns_store(d, region, rev, batch)
        if utils:
            log(f"{d:%d-%m-%Y} NS {region}: {len(utils)} utilities (rev {rev})")

    if failed:
        log(f"{d:%d-%m-%Y} NS incomplete: {failed} utilities failed; the next run will retry them")
        return False
    with db().begin() as c:          # drop rows left over from an older revision, mark the day complete
        c.execute(ns_schedule.delete().where(ns_schedule.c.trade_date == d, ns_schedule.c.revision != rev))
        c.execute(ns_util_status.delete().where(ns_util_status.c.trade_date == d, ns_util_status.c.revision != rev))
        c.execute(ns_day.delete().where(ns_day.c.trade_date == d))
        c.execute(ns_day.insert(), {"trade_date": d, "revision": rev, "n_utilities": total,
                                    "completed_at": dt.datetime.now()})
    msg = f"{d:%d-%m-%Y} NS complete: rev {rev}, {total} utilities ({fetched} downloaded) in {time.time() - t0:.0f}s"
    if NS_SAVE_EXCEL:
        msg += f" -> {ns_export_excel(d, rev)}"
    log(msg)
    return True


def ns_sync(s: WbesSession, start: dt.date | None = None, end: dt.date | None = None) -> None:
    """Make the DB complete (latest revision) for every date from start to end, newest first."""
    today = dt.date.today()
    start = start or dt.datetime.strptime(NS_START_DATE, "%d-%m-%Y").date()
    end = end or (today if NS_INCLUDE_TODAY else today - dt.timedelta(days=1))
    with db().connect() as c:
        stored = {r.trade_date: r.revision for r in c.execute(select(ns_day.c.trade_date, ns_day.c.revision))}
    days = [end - dt.timedelta(days=i) for i in range((end - start).days + 1)]
    todo = [d for d in days if d not in stored or (today - d).days <= NS_RECHECK_DAYS]
    log(f"NS: {len(days)} dates in range, {len(days) - len(todo)} already complete, checking {len(todo)}")
    for d in todo:
        rev = ns_latest_rev(s, d)
        if rev < 0:
            log(f"{d:%d-%m-%Y} NS: no schedule on portal")
            continue
        if d in stored and stored[d] >= rev:
            continue                              # recent date, no new revision
        if d in stored:
            log(f"{d:%d-%m-%Y} NS: revision {stored[d]} -> {rev}, re-downloading")
        ns_pull_date(s, d, rev)
    log("NS: database is up to date" + ("" if NS_INCLUDE_TODAY else " (up to yesterday)"))


def main() -> None:
    ap = argparse.ArgumentParser(description="Pull WBES PX + Net Schedule data, no browser after login")
    ap.add_argument("--job", choices=["px", "ns", "all"], default=None, help="which job(s) to run")
    ap.add_argument("--once", action="store_true")
    ap.add_argument("--from", dest="date_from")
    ap.add_argument("--to", dest="date_to")
    args = ap.parse_args()

    def dates_to_pull() -> list[dt.date]:
        if args.date_from:
            a = dt.datetime.strptime(args.date_from, "%d-%m-%Y").date()
            b = dt.datetime.strptime(args.date_to or args.date_from, "%d-%m-%Y").date()
            return [a + dt.timedelta(days=i) for i in range((b - a).days + 1)]
        return [dt.date.today() + dt.timedelta(days=o) for o in DAY_OFFSETS]

    jobs = JOBS if args.job is None else (["px", "ns"] if args.job == "all" else [args.job])
    if "ns" in jobs and not SAVE_DB:
        raise SystemExit("The Net Schedule job needs the database (SAVE_DB = True).")
    ns_range = (None, None)
    if args.date_from:
        a = dt.datetime.strptime(args.date_from, "%d-%m-%Y").date()
        ns_range = (a, dt.datetime.strptime(args.date_to, "%d-%m-%Y").date() if args.date_to else a)

    log(f"Jobs: {', '.join(jobs)}")
    if SAVE_DB:                              # fail fast on a wrong DB_URL, before asking for OTP
        log(f"Database: {db().url.render_as_string(hide_password=True)}")
    if SAVE_EXCEL:
        log(f"Excel folders: {OUT_DIR.resolve()}" + (f" and {NS_OUT_DIR.resolve()}" if "ns" in jobs else ""))
    s = get_session()
    last_ns = 0.0
    while True:
        try:
            if "px" in jobs:
                pull_dates(s, dates_to_pull())
            if "ns" in jobs and (args.once or args.date_from or time.time() - last_ns >= NS_EVERY_MINUTES * 60):
                ns_sync(s, *ns_range)
                last_ns = time.time()
            if args.once or args.date_from:
                break
            log(f"Sleeping {POLL_MINUTES} min (token renewed automatically)...")
            sleep_keeping_token(s, POLL_MINUTES * 60)
        except LoginNeeded as e:
            log(f"Session ended after {(time.time() - s.login_at) / 3600:.1f} h ({e}).")
            s = browser_login()
        except requests.RequestException as e:
            log(f"Network error: {e!r}; retrying in 60 s")
            time.sleep(60)
        except KeyboardInterrupt:
            break


if __name__ == "__main__":
    main()
