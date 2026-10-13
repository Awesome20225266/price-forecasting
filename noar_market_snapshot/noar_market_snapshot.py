"""
NOAR (Grid-India) - TRAS Market Snapshot bulk export
====================================================

Drives the public page https://noar.in/static-analytics/market-snapshot in a real,
visible Chrome window the same way a person would (Date -> Tras Type -> Process Type
-> Submit) and saves the JSON the page receives. It loops every date x
Tras Type (Up/Down) x Process Type (DAM/RTM).

Output
  noar_market_snapshot.csv   all rows, appended as it goes
  noar_progress.txt          finished (date|tras|process) keys, so a re-run resumes
  noar_market_snapshot.xlsx  clean, de-duplicated Excel (built at the end)

Setup (once)
  pip install playwright pandas openpyxl
  playwright install chromium

Run
  python noar_market_snapshot.py                          # full history: 01/04/2023 -> yesterday
  python noar_market_snapshot.py --start 01/01/2026 --end 31/01/2026
  python noar_market_snapshot.py --daily                  # yesterday only (for a daily schedule)
  python noar_market_snapshot.py --excel-only             # just rebuild the Excel from the CSV

Notes
  * The site protects its API with Google reCAPTCHA v3. This script does NOT try to get
    around it: it just uses the page normally, at a slow pace. If the site starts
    rejecting requests, the script stops; re-run later and it resumes where it left off.
  * Keep the browser window open/visible while it runs (headless mode is more likely
    to be rejected).
"""

import argparse
import csv
import json
import os
import random
import sys
import time
from datetime import datetime, timedelta, timezone

from playwright.sync_api import sync_playwright, TimeoutError as PWTimeout

PAGE_URL = "https://noar.in/static-analytics/market-snapshot"
API_PART = "GetTRASMarketSnapshotReport"
FIRST_DATE = "01/04/2023"

TRAS_TYPES = [(1, "Up"), (2, "Down")]
PROCESS_TYPES = [(1, "DAM"), (3, "RTM")]  # portal uses value 3 for RTM
TRAS_LABEL = {1: "Up", 2: "Down"}
PROCESS_LABEL = {1: "DAM", 2: "RTM", 3: "RTM"}  # portal's own table shows RTM as "NA" (bug)

DATE_INPUT = 'input[formcontrolname="Date"]'
TRAS_SELECT = 'select[formcontrolname="TrasType"]'
PROC_SELECT = 'select[formcontrolname="ProcessType"]'

COLUMNS = [
    "Date", "Time_Block_Id", "Time_Block_Description", "Tras_Type", "Process_Type",
    "Requirement_MW", "Sell_Bid_MW", "Cleared_Volume_MW",
    "Max_Clearing_Price_INR_MWh", "Min_Clearing_Price_INR_MWh", "Pulled_At",
]

IST = timezone(timedelta(hours=5, minutes=30))


class FetchError(Exception):
    pass


def parse_ddmmyyyy(s):
    return datetime.strptime(s, "%d/%m/%Y").date()


def date_range(start, end):
    d = start
    while d <= end:
        yield d
        d += timedelta(days=1)


def load_progress(path):
    if not os.path.exists(path):
        return set()
    with open(path, encoding="utf-8") as f:
        return {line.strip() for line in f if line.strip()}


def count_data_rows(path):
    if not os.path.exists(path):
        return 0
    with open(path, encoding="utf-8") as f:
        n = sum(1 for _ in f)
    return max(0, n - 1)


def format_bar(done, total, width=28):
    frac = (done / total) if total else 1
    filled = int(round(width * frac))
    filled = max(0, min(width, filled))
    return "#" * filled + "-" * (width - filled)


def write_status(path, payload):
    """Update the progress file. A locked file must not stop the download."""
    data = json.dumps(payload)
    tmp = path + ".tmp"
    for attempt in range(6):
        try:
            with open(tmp, "w", encoding="utf-8") as f:
                f.write(data)
            os.replace(tmp, path)
            return
        except PermissionError:
            time.sleep(0.15 * (attempt + 1))
    try:
        with open(path, "w", encoding="utf-8") as f:
            f.write(data)
    except PermissionError:
        pass


def open_page(pw, headless):
    browser = pw.chromium.launch(headless=headless)
    context = browser.new_context(accept_downloads=False, viewport={"width": 1366, "height": 900})
    page = context.new_page()
    page.goto(PAGE_URL, wait_until="networkidle", timeout=90_000)
    page.wait_for_selector(DATE_INPUT, timeout=60_000)
    return browser, page


def fetch_one(page, date_str, tras_id, proc_id):
    """Set the three filters, click Submit, return the rows the page receives."""
    page.fill(DATE_INPUT, date_str)
    page.select_option(TRAS_SELECT, str(tras_id))
    page.select_option(PROC_SELECT, str(proc_id))

    with page.expect_response(
        lambda r: API_PART in r.url and r.request.method == "POST", timeout=60_000
    ) as info:
        page.get_by_role("button", name="Submit").click()
    resp = info.value

    if resp.status != 200:
        raise FetchError(f"HTTP {resp.status}")

    sent = resp.request.post_data_json or {}
    if sent.get("Date") != date_str or sent.get("TrasTypeID") != tras_id or sent.get("ProcessTypeID") != proc_id:
        raise FetchError(f"page sent different filters than expected: "
                         f"{ {k: sent.get(k) for k in ('Date', 'TrasTypeID', 'ProcessTypeID')} }")

    data = resp.json()
    status = data.get("ResponseStatus") or {}
    if status.get("Code") != "NOAR-200":
        raise FetchError(f"{status.get('Code')}: {status.get('Message')}")
    return data.get("ResponseBody") or []


def to_csv_rows(rows, pulled_at):
    out = []
    for r in rows:
        out.append({
            "Date": r.get("dateId"),
            "Time_Block_Id": r.get("TimeBlockId"),
            "Time_Block_Description": r.get("TimeBlockDescription"),
            "Tras_Type": TRAS_LABEL.get(r.get("TrasType"), r.get("TrasType")),
            "Process_Type": PROCESS_LABEL.get(r.get("ProcessType"), r.get("ProcessType")),
            "Requirement_MW": r.get("TrasRequirement"),
            "Sell_Bid_MW": r.get("SellBid"),
            "Cleared_Volume_MW": r.get("ClearedMW"),
            "Max_Clearing_Price_INR_MWh": r.get("MCP"),
            "Min_Clearing_Price_INR_MWh": r.get("MinMCP"),
            "Pulled_At": pulled_at,
        })
    return out


def build_excel(csv_path, xlsx_path):
    import pandas as pd

    if not os.path.exists(csv_path):
        print("No CSV yet - nothing to export.")
        return
    df = pd.read_csv(csv_path)
    if df.empty:
        print("CSV is empty - nothing to export.")
        return
    df["_d"] = pd.to_datetime(df["Date"], format="%d/%m/%Y")
    # keep the latest pull if a date/block was fetched more than once
    df = (df.sort_values("Pulled_At")
            .drop_duplicates(["Date", "Time_Block_Id", "Tras_Type", "Process_Type"], keep="last")
            .sort_values(["_d", "Process_Type", "Tras_Type", "Time_Block_Id"])
            .drop(columns="_d"))
    df.to_excel(xlsx_path, index=False, sheet_name="MarketSnapshot")
    print(f"Excel written: {xlsx_path} ({len(df):,} rows)")


def main():
    yesterday = (datetime.now(IST) - timedelta(days=1)).strftime("%d/%m/%Y")

    ap = argparse.ArgumentParser(description="Bulk export NOAR TRAS Market Snapshot")
    ap.add_argument("--start", default=FIRST_DATE, help="DD/MM/YYYY (default %(default)s)")
    ap.add_argument("--end", default=yesterday, help="DD/MM/YYYY (default: yesterday)")
    ap.add_argument("--daily", action="store_true", help="fetch yesterday only (re-fetches it)")
    ap.add_argument("--force", action="store_true", help="re-fetch combinations already done")
    ap.add_argument("--delay", type=float, default=3.0, help="seconds between requests (default 3)")
    ap.add_argument("--out", default="noar_market_snapshot.csv")
    ap.add_argument("--progress", default="noar_progress.txt")
    ap.add_argument("--xlsx", default="noar_market_snapshot.xlsx")
    ap.add_argument("--status", default="noar_status.json", help="live progress file for the progress page")
    ap.add_argument("--headless", action="store_true", help="hide the browser (may get rejected)")
    ap.add_argument("--excel-only", action="store_true", help="only rebuild the Excel from the CSV")
    args = ap.parse_args()

    if args.excel_only:
        build_excel(args.out, args.xlsx)
        return

    if args.daily:
        args.start = args.end = yesterday
        args.force = True

    start, end = parse_ddmmyyyy(args.start), parse_ddmmyyyy(args.end)
    done = set() if args.force else load_progress(args.progress)

    all_jobs = [(d.strftime("%d/%m/%Y"), t, p)
                for d in date_range(start, end) for t in TRAS_TYPES for p in PROCESS_TYPES]
    jobs = [j for j in all_jobs if f"{j[0]}|{j[1][0]}|{j[2][0]}" not in done]
    total = len(all_jobs)
    completed = total - len(jobs)
    csv_rows = count_data_rows(args.out)
    run_started = time.time()
    finished_this_run = 0

    def publish(state, current="", last_rows=0, message=""):
        elapsed = time.time() - run_started
        eta = None
        if finished_this_run > 0 and completed < total:
            eta = int((total - completed) * (elapsed / finished_this_run))
        write_status(args.status, {
            "state": state,
            "done": completed,
            "total": total,
            "percent": round(100 * completed / total, 1) if total else 100.0,
            "current": current,
            "last_rows": last_rows,
            "csv_rows": csv_rows,
            "eta_seconds": eta,
            "message": message,
            "range": f"{args.start} -> {args.end}",
            "updated_at": datetime.now(IST).isoformat(timespec="seconds"),
        })

    print(f"{len(jobs):,} requests left ({args.start} -> {args.end}); "
          f"{completed:,} of {total:,} already saved; "
          f"~{len(jobs) * (args.delay + 2) / 3600:.1f} h at this pace.")
    publish("running" if jobs else "complete", message="Opening the NOAR page" if jobs else "Nothing left to fetch")
    if not jobs:
        build_excel(args.out, args.xlsx)
        return

    new_file = not os.path.exists(args.out)
    csv_f = open(args.out, "a", newline="", encoding="utf-8")
    writer = csv.DictWriter(csv_f, fieldnames=COLUMNS)
    if new_file:
        writer.writeheader()
    prog_f = open(args.progress, "a", encoding="utf-8")

    fails_in_a_row, max_fails = 0, 5
    stopped_early = False
    with sync_playwright() as pw:
        browser, page = open_page(pw, args.headless)
        i = 0
        while i < len(jobs):
            date_str, (tras_id, tras_name), (proc_id, proc_name) = jobs[i]
            key = f"{date_str}|{tras_id}|{proc_id}"
            current = f"{date_str}  {tras_name}  {proc_name}"
            try:
                rows = fetch_one(page, date_str, tras_id, proc_id)
                pulled_at = datetime.now(IST).isoformat(timespec="seconds")
                writer.writerows(to_csv_rows(rows, pulled_at))
                csv_f.flush()
                prog_f.write(key + "\n")
                prog_f.flush()
                completed += 1
                finished_this_run += 1
                csv_rows += len(rows)
                bar = format_bar(completed, total)
                pct = 100 * completed / total if total else 100
                print(f"[{bar}] {pct:5.1f}%  {completed}/{total}  {current}: {len(rows)} rows", flush=True)
                publish("running", current, len(rows))
                fails_in_a_row = 0
                i += 1
                time.sleep(args.delay + random.uniform(0, 1.5))
            except (FetchError, PWTimeout) as e:
                fails_in_a_row += 1
                print(f"[FAILED {fails_in_a_row}/{max_fails}] {current}: {e}", flush=True)
                publish("running", current, message=f"Retry {fails_in_a_row} of {max_fails}: {e}")
                if fails_in_a_row >= max_fails:
                    print(f"\n{max_fails} failures in a row - the site is probably rejecting "
                          "automated requests right now. Stopping. Re-run later to resume.")
                    publish("stopped", current, message="The site rejected 5 requests in a row. Re-run later to resume.")
                    stopped_early = True
                    break
                wait = 30 * fails_in_a_row
                print(f"  waiting {wait}s, reloading page, then retrying ...", flush=True)
                time.sleep(wait)
                try:
                    page.reload(wait_until="networkidle", timeout=90_000)
                    page.wait_for_selector(DATE_INPUT, timeout=60_000)
                except PWTimeout:
                    pass
        browser.close()

    csv_f.close()
    prog_f.close()
    if not stopped_early:
        publish("complete", message="Building the Excel file")
    build_excel(args.out, args.xlsx)
    if not stopped_early:
        publish("complete", message="Excel file is ready")


if __name__ == "__main__":
    try:
        main()
    except KeyboardInterrupt:
        print("\nStopped by user - progress saved; re-run to resume.")
        sys.exit(1)
