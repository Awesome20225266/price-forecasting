"""Download every WBES PX and Net Schedule day the portal still has, then stop.

The daily task (run_pipeline.bat) only covers yesterday/today/tomorrow and recent
net-schedule revisions. This run fills everything older than that. It waits if
the portal is down, then keeps going until the range is stored.
"""
from __future__ import annotations

import datetime as dt
import subprocess
import sys
import time

import requests

import wbes_pull as w

TASK = "WBES Market Data Pipeline"
FLOOR = dt.date(2023, 1, 1)


def portal_up() -> bool:
    try:
        requests.get("https://newwbes.grid-india.in/", timeout=25)
        return True
    except Exception:
        return False


def has_px(s: w.WbesSession, d: dt.date) -> bool:
    ds = f"{d:%d-%m-%Y}"
    pxs = s.call(f"/powerexchange/api/PowerExchange/GetPxTypeList/date/{ds}")
    n = len(pxs) if isinstance(pxs, list) else 0
    w.log(f"probe PX {ds}: {n} exchange(s)")
    return n > 0


def has_ns(s: w.WbesSession, d: dt.date) -> bool:
    rev = w.ns_latest_rev(s, d)
    w.log(f"probe NS {d:%d-%m-%Y}: rev {rev}")
    return rev >= 0


def earliest(s: w.WbesSession, pred) -> dt.date:
    """Earliest date that has data, assuming it stays available once it starts."""
    hi = dt.date.today()
    if not pred(s, hi):
        hi = hi - dt.timedelta(days=1)
        if not pred(s, hi):
            fallback = dt.datetime.strptime(w.NS_START_DATE, "%d-%m-%Y").date()
            w.log(f"Could not confirm a recent date; starting at {fallback:%d-%m-%Y}")
            return fallback
    known = hi
    probe = hi
    while True:
        probe = max(FLOOR, probe - dt.timedelta(days=30))
        if pred(s, probe):
            known = probe
            if probe == FLOOR:
                return FLOOR
            continue
        left, right = probe, known
        break
    while (right - left).days > 1:
        mid = left + dt.timedelta(days=(right - left).days // 2)
        if pred(s, mid):
            right = mid
        else:
            left = mid
    return right


def ns_missing(start: dt.date, end: dt.date) -> list[dt.date]:
    from sqlalchemy import select

    with w.db().connect() as c:
        stored = {r.trade_date for r in c.execute(select(w.ns_day.c.trade_date))}
    return [start + dt.timedelta(days=i) for i in range((end - start).days + 1) if (start + dt.timedelta(days=i)) not in stored]


def enable_task() -> None:
    subprocess.run(["schtasks", "/Change", "/TN", TASK, "/ENABLE"], check=False)


def main() -> None:
    w.log("Historical WBES download: PX + Net Schedule, every date the portal still has.")
    s = None
    px_dates: list[dt.date] | None = None
    ns_start = ns_end = None
    phase = "discover"
    while True:
        try:
            if s is None:
                if portal_up():
                    s = w.get_session()
                else:
                    w.log("Portal is not answering yet. Opening Chrome — log in with OTP once the page loads.")
                    s = w.browser_login()
            if phase == "discover":
                px_start = earliest(s, has_px)
                ns_start = earliest(s, has_ns)
                px_end = dt.date.today() + dt.timedelta(days=1)   # day-ahead is already published
                ns_end = dt.date.today() - dt.timedelta(days=1)   # today is still being revised
                px_dates = [px_start + dt.timedelta(days=i) for i in range((px_end - px_start).days + 1)]
                w.log(
                    f"PX {px_start:%d-%m-%Y} .. {px_end:%d-%m-%Y} ({len(px_dates)} days); "
                    f"Net Schedule {ns_start:%d-%m-%Y} .. {ns_end:%d-%m-%Y}"
                )
                phase = "px"
            if phase == "px":
                w.pull_dates(s, px_dates)
                w.log("PX historical pass finished.")
                phase = "ns"
            if phase == "ns":
                for attempt in range(1, 4):
                    w.ns_sync(s, ns_start, ns_end)
                    still = [d for d in ns_missing(ns_start, ns_end) if has_ns(s, d)]
                    w.log(f"NS pass {attempt}: {len(still)} date(s) still on the portal but not stored")
                    if not still:
                        break
                phase = "done"
            w.log("HISTORICAL COMPLETE. Later runs only need new and recently revised days.")
            enable_task()
            w.log("Daily task re-enabled (18:00, yesterday/today/tomorrow + recent net schedule).")
            return
        except w.LoginNeeded as e:
            w.log(f"Session ended ({e}). A Chrome window will open for OTP.")
            s = None
            time.sleep(15)
        except SystemExit as e:
            w.log(f"Login window closed without a session ({e}). Retrying in 2 min.")
            s = None
            time.sleep(120)
        except requests.RequestException as e:
            w.log(f"Network error during {phase}: {e!r}. Retrying in 60 s.")
            time.sleep(60)
        except Exception as e:
            w.log(f"Unexpected error during {phase}: {e!r}. Retrying in 2 min.")
            time.sleep(120)


if __name__ == "__main__":
    main()
