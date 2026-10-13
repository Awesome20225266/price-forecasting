# Price Forecasting: Indian power-market data pipelines

This repo holds five small Python pipelines. Each one downloads market data from an Indian power exchange or grid portal and keeps a local copy up to date every day. That data is the input for price forecasting.

| Folder | Source | What it pulls | Saved as | Login? | Daily task (IST) |
|---|---|---|---|---|---|
| [`pxil_tam_gtam`](pxil_tam_gtam/README.md) | PXIL (powerexindia.in) | TAM and GTAM: market snapshot, DAC, Intraday, Any-Day, Daily, Weekly, Monthly | `PXIL_TAM_GTAM.xlsx` + SQLite cache | No | 09:15 |
| [`noar_market_snapshot`](noar_market_snapshot/noar_market_snapshot.py) | NOAR (noar.in) | TRAS market snapshot: Up/Down x DAM/RTM, per 15-min block | `noar_market_snapshot.xlsx` / `.csv` | No (opens a visible Chrome window) | 10:00, only if you register it |
| [`HPX_Downloader`](HPX_Downloader/README.txt) | HPX (hpxindia.com) | Contingency TAM, TAM/GTAM Monthly, High-Price TAM, AS DAM, Green TAM | Excel files in `HPX_Data\` | No | 10:15 |
| [`iex_pipeline`](iex_pipeline/README.md) | IEX (iexindia.com) | DAM, RTM, GDAM, HP-DAM, Intra-Day/DAC/TAM (conventional, green, high-price), TAM auction lists | `data\iex_market_data.db` (SQLite) + CSVs | No | 14:15 |
| [`wbes_pipeline`](wbes_pipeline/wbes_pull.py) | WBES (Grid-India) | PX dashboard (IEX/PXI/HPX, all markets) and Net Schedule (all regions, all utilities) | `data\wbes.db` (SQLite) + Excel | **Yes: your own WBES user + OTP** | 18:00 |

Each folder works on its own. Set up only the ones you need.

> **What is NOT in this repo, on purpose:** downloaded data, databases, logs, `.venv` folders and WBES login tokens. They are listed in `.gitignore`. Each person's setup downloads fresh data from the portals. Never add data files or `wbes_session.json` to Git.

---

## 1. What you need first (one time per PC)

1. **Windows 10 or 11.** The setup and scheduling scripts are `.bat` and PowerShell files.
2. **Python 3.9 or newer** from <https://www.python.org/downloads/>. During install, tick **"Add Python to PATH"**.
   - If typing `python` opens the Microsoft Store, turn off the alias: *Settings > Apps > Advanced app settings > App execution aliases* and switch off `python.exe` and `python3.exe`.
   - The original machine runs these pipelines on Python 3.14.
3. **Git** from <https://git-scm.com/download/win>, with the default options.
4. **Google Chrome**, only for `wbes_pipeline` (OTP login window) and `noar_market_snapshot` (drives the page).
5. **Disk space.** Rough sizes after a full history download on the original machine: IEX about 350 MB, PXIL about 40 MB, HPX about 50 MB, NOAR about 40 MB, **WBES about 5 GB**.

## 2. Get the code

Open **Command Prompt** in the folder where you keep projects (for example `C:\Projects`) and run:

```bat
git clone https://github.com/Awesome20225266/price-forecasting.git
cd price-forecasting
```

This creates a `price-forecasting` folder with one sub-folder per pipeline.

## 3. Set up a pipeline

Pick a pipeline folder and double-click its **`setup_windows.bat`**. Every setup creates a private Python environment (`.venv`) inside that folder and installs that pipeline's packages from `requirements.txt`. Your other Python projects are not touched.

What each setup does after that:

| Folder | `setup_windows.bat` then... | First run takes |
|---|---|---|
| `pxil_tam_gtam` | back-fills from 2023-04-01, then registers the daily task | tens of minutes |
| `HPX_Downloader` | downloads history from 2023-04-01, then registers the daily task | a few minutes |
| `iex_pipeline` | full back-fill from 2020-04-01, completeness check, CSV export, then registers the daily task | long (hours possible): leave the window open |
| `noar_market_snapshot` | installs Playwright's Chromium. **No download and no task yet.** Run `run_noar.bat` for full history (from 01/04/2023) or `run_noar.bat --daily` for yesterday only | full history is slow by design |
| `wbes_pipeline` | registers the daily task. **No download yet.** Run `run_pipeline.bat` once **while you are at the PC**, so you can complete the OTP login in the Chrome window it opens | depends on the date range |

Where a pipeline's own README gives a different start date, the script's setting is the one that runs (`START_DATE`, `DEFAULT_START`, `FIRST_DATE` or `NS_START_DATE` at the top of each script).

### Daily schedule

The setup registers a Windows **Task Scheduler** job, except for NOAR. If the PC was off at that time, the job runs as soon as it is back on. Each run fetches only new days plus a few recent days again, in case the portal revised them. Running it twice does not create duplicates.

- Change the time: in the pipeline folder, run `powershell -ExecutionPolicy Bypass -File setup_task.ps1 -At 15:00`
- Schedule NOAR: in its folder, run `powershell -ExecutionPolicy Bypass -File setup_task.ps1` (needs you logged in, because a browser window opens)
- Run a job now: `Start-ScheduledTask -TaskName "IEX Market Data Pipeline"` (in PowerShell)
- Remove a job: `Unregister-ScheduledTask -TaskName "IEX Market Data Pipeline" -Confirm:$false`

Task names: `PXIL TAM GTAM update`, `NOAR market snapshot update`, `HPX Market Data update`, `IEX Market Data Pipeline`, `WBES Market Data Pipeline`.

### Agree as a team who runs what

If everyone runs every setup, every PC downloads the same data every day and the portals see many repeated requests. **Agree on one machine per pipeline** that keeps the master data. Everyone else can set up a pipeline just to develop or test.

## 4. WBES: read this before you run it

- You need **your own WBES user**. The first run opens a separate Chrome window. Log in there with username, OTP and captcha. The script reads the login tokens, closes the window and carries on without a browser.
- The tokens are saved to `wbes_pipeline\wbes_session.json`. **That file is your live login. Never commit, email or share it.** Git already ignores it. Delete it to force a fresh login.
- **One WBES user = one machine.** Don't run the script on two PCs with the same user, and don't use that user in a normal browser while it runs. The portal rotates the refresh token, so one of the two gets logged out.
- Run it on a PC that does not go to sleep. If the token lapses for more than about 10 minutes, you will need to log in with OTP again.
- To download older history the daily job does not cover, run `.venv\Scripts\python.exe run_historical.py` in the `wbes_pipeline` folder.
- To store in SQL Server, PostgreSQL or MySQL instead of SQLite, set the environment variable `WBES_DB_URL`. See the comments near `DB_URL` in `wbes_pull.py`. Never write a database password into the script.

## 5. Get the latest code safely

When someone has pushed changes, update your copy from the `price-forecasting` folder:

```bat
git pull --ff-only
```

- `--ff-only` only updates when it can do so cleanly. If you edited a file that also changed on GitHub, it stops and changes nothing, so your work is never overwritten. Run `git status` to see your edits, then ask before going further.
- Your data, logs, `.venv` and WBES tokens are ignored by Git, so a pull never touches them.
- If a `requirements.txt` changed, re-run that folder's `setup_windows.bat`, or run `.venv\Scripts\python.exe -m pip install -r requirements.txt` in that folder.

## 6. Making changes

1. Make a branch: `git switch -c my-change`
2. Edit and test by running the pipeline by hand (for example `run_pipeline.bat`).
3. Before you commit, run `git status` and check that **no data files, logs or `wbes_session.json`** are listed.
4. Commit, push the branch, and open a pull request on GitHub for review.

## 7. Troubleshooting

| Symptom | What to do |
|---|---|
| `Python 3.9+ is required` | Install Python and tick "Add Python to PATH", or switch off the Store alias (step 1). |
| Excel file "in use" | Close the Excel file and run again. HPX and PXIL keep the data in their own store, so nothing is lost. |
| IEX log shows `HTTP 403 ... instead of JSON` | IEX is blocking script requests. Install the browser fallback (see `iex_pipeline/README.md`). `run_pipeline.bat` retries with it automatically. |
| NOAR stops with rejected requests | The site uses reCAPTCHA. The script does not bypass it. Run again later; it resumes where it stopped. |
| WBES keeps asking for OTP | Someone else may be using the same WBES user, or the PC slept. See section 4. |
| Where are the logs? | `logs\` in HPX, IEX and WBES. `pxil_tam_gtam.log` in PXIL. NOAR prints to its window. |

## Notes

- All dates are **IST**. Each pipeline's README or script header says whether a date is a delivery date or a trade date.
- These scripts read the same endpoints the portals' own web pages use. Those endpoints are not documented and can change without notice. The scripts fail loudly rather than store bad data. Check each portal's terms on automated access before relying on this commercially. IEX also offers an official market-data API (see `iex_pipeline/README.md`).
- More detail for each pipeline is in its own README or at the top of its main script.
