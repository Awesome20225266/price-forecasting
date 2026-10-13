HPX India Market Data Downloader
================================

What it does
  Downloads these HPX reports into Excel and keeps them up to date.
  Each run fetches only new data (plus the last 3 days again, in case
  HPX revised them) and removes duplicates.

    Contingency_TAM.xlsx            Trade Details, Trade Summary
    TAM_Monthly.xlsx                Trade Details
    GTAM_Monthly.xlsx               Trade Details
    HighPrice_Contingency_TAM.xlsx  Trade Details, Trade Summary
    AS_DAM.xlsx                     UP, UP Daily Summary, DOWN, DOWN Daily Summary
    Green_Contingency_TAM.xlsx      Trade Details, Trade Summary

Setup (one time)
  1. Install Python 3 from https://www.python.org/downloads/
     (on Windows, tick "Add Python to PATH" during install).
  2. Extract this zip to any folder.

How to run
  Windows : double-click run_hpx_update.bat
  Mac/Linux: python3 hpx_downloader.py
  (The first run installs the needed packages and downloads history
   from 1 Apr 2026, which takes a few minutes.)

  Your Excel files appear in the "HPX_Data" folder next to the script.

Options (run from a command prompt in this folder)
  python hpx_downloader.py --from 2025-04-01   re-download everything from a date
  python hpx_downloader.py --only as_dam       update just one report
  python hpx_downloader.py --list              show report names
  python hpx_downloader.py --dry-run           show what would be fetched

Notes
  - Close the Excel files before running. If one is open, the script
    tells you, and your data is still saved.
  - Do not delete HPX_Data\_store - it holds the master data and the
    download progress.
  - To change the first-run start date, edit START_DATE at the top of
    hpx_downloader.py.
