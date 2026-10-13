# PXIL TAM + GTAM → Excel (auto-updating)

Downloads every TAM and GTAM section from powerexindia.in into one master workbook, **PXIL_TAM_GTAM.xlsx**, and fills in any missing dates on each run.

## What you get

| Sheet | Site section |
|---|---|
| Overview | Row counts, date coverage, **missing-day count** per sheet, last download time |
| TAM_MarketSnapshot | TAM › Market Snapshot › Market Volume Profile (daily, all contract types) |
| TAM_DAC / _Summary | TAM › DAC › Market Volume Profile Report (+ Max/Min/Avg/Total per day) |
| TAM_Intraday / _Summary | TAM › Intraday |
| TAM_AnyDay / _Summary | TAM › Anyday |
| TAM_Daily / _Summary | TAM › Daily |
| TAM_Weekly / _Summary | TAM › Weekly |
| TAM_Monthly / _Summary | TAM › Monthly |
| GTAM_MarketSnapshot | GTAM › Market Snapshot (Solar / Non Solar / Total volume) |
| GTAM_DAC / _Summary | GTAM › Day Ahead Contingency (Solar + Non Solar, with an *Energy Type* column) |
| GTAM_Intraday / _Summary | GTAM › Intra Day Product |
| GTAM_AnyDay / _Summary | GTAM › Any Day Product |

Dates are real Excel dates, and numbers are real numbers, so filters, pivots and formulas work straight away.

## One-time setup (Windows)

1. Install Python 3.9 or later from python.org (tick **"Add Python to PATH"**).
2. Put this folder somewhere permanent, e.g. `C:\PXIL`.
3. Open Command Prompt in that folder and run:
   ```
   pip install -r requirements.txt
   ```
4. Double-click **run_pxil.bat**. The first run back-fills from **1 Apr 2025**, which takes about 10–20 minutes. Later runs take seconds.

## Schedule it daily

In Command Prompt (change the path to match your folder):

```
schtasks /Create /SC DAILY /ST 09:15 /TN "PXIL TAM GTAM update" /TR "C:\PXIL\run_pxil.bat"
```

If the PC was off or the site was down, those dates stay "missing", and the next run downloads them automatically. The latest 3 days (yesterday, today and tomorrow's DAC) are always re-downloaded, because the exchange keeps adding trades for them.

On Mac or Linux, use cron: `15 9 * * * cd /path/to/folder && python3 pxil_tam_gtam.py`

## Options

```
python pxil_tam_gtam.py --start 2024-04-01            # back-fill further back
python pxil_tam_gtam.py --only GTAM                   # just GTAM (or TAM)
python pxil_tam_gtam.py --sections GTAM_DAC,TAM_DAC   # specific sheets
python pxil_tam_gtam.py --no-summaries                # skip *_Summary sheets (fewer requests)
python pxil_tam_gtam.py --redownload 2026-09-01 2026-09-30
python pxil_tam_gtam.py --export-only                 # rebuild Excel from the cache
python pxil_tam_gtam.py --out "D:\Reports\PXIL"       # write outputs elsewhere
```

## Files it creates

- `PXIL_TAM_GTAM.xlsx`: the master workbook, rebuilt on every run.
- `pxil_tam_gtam_cache.sqlite`: the local store behind the workbook. It records which days have already been downloaded. **Keep this file**; if you delete it, the next run starts the back-fill again.
- `pxil_tam_gtam.log`: the run log.

If the workbook is open in Excel during a run, the update is saved as `PXIL_TAM_GTAM_<timestamp>.xlsx` instead. Close Excel and run again to update the master copy.

## Notes

- Data comes from the same JSON endpoints the PXIL web page uses (`/get_gdac_bil_data/`, `/get_MS_mvp_data/`, and so on). No login is needed.
- Days with no trades have no rows. That's common for Any Day, for example. The Overview still counts those days as checked, not missing.
- If PXIL changes its website, the endpoint names at the top of the script (the `SECTIONS` list) are the only place that needs editing.
