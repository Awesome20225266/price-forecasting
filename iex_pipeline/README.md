# IEX market data pipeline

This pipeline downloads IEX market data into a local SQLite database and keeps it up to date. It covers:

- **DAM, RTM, GDAM and HP-DAM** (15-minute blocks)
- **Intra-Day / DAC / Term-Ahead**, in conventional, Green and High-Price versions (every traded contract)
- **TAM Any-Day Single-Side Reverse Auctions** and **TAM Monthly** auctions (list data only)

Auction result PDFs are not downloaded. Everything is kept back to 1 April 2020, or to the market's launch if that was later.

## Datasets

| `--markets` name | What it is | Source on the portal | Row = | History |
|---|---|---|---|---|
| `dam` `rtm` `gdam` `hpdam` | Market snapshot: bids, MCV, MCP, congestion | `/api/v1/{mkt}/market-snapshot` (JSON) | delivery date × 15-min block | DAM 2020-04-01 · RTM 2020-06-01 · GDAM 2021-10-27 · HP-DAM 2023-03-11 |
| `tam` | Intra-Day, DAC, Daily and Weekly TAM: trade details | `/market-data/intra-day-ahead-contingency/trade-details` | trade date × contract type × instrument | 2020-04-01 |
| `gtam` | Green Intra-Day / DAC / TAM: trade details (solar and non-solar) | `/market-data/green-intra-day/g-market-data` | same | 2020-08-21 |
| `hptam` | High-Price Intra-Day / DAC / TAM: trade details | `/market-data/high-price-intra-day/trade-details` | same | first trade 2024-04-26 (trades are rare) |
| `tam_summary` `gtam_summary` | Daily totals per contract type, including **total buy/sell bid volume** | `/api/v1/{tam,gtam}/market-snapshot?range=DAILY` (JSON) | trade date × contract type | 2020 |
| `tam_anyday` | TAM Any Day(s) Single-Side Reverse Auction list | `/tam-anyday` | one auction | 2022 |
| `tam_monthly` | TAM Monthly auctions: instrument, bids, volume (MUs), price, delivery month | `/tam-monthly` | auction × instrument | 2022-08 |

Groups: `snapshot` (the 4 block markets), `intraday` (tam, gtam, hptam and both summaries), `auctions` (tam_anyday, tam_monthly), and `all` (the default).

### How each one is pulled

- **Block markets and summaries** come from the JSON API behind the site.
- **Trade details and auction lists** are pulled the same way the site's own **Update Report** button does it. The script asks for the page with the header `RSC: 1`, which makes the server return the page's data payload instead of the HTML. It then reads the `tableData` or `tableRows` array out of that payload. You don't need to log in or send cookies. One month of TAM trades is about 15,000 rows (about 5 MB) and comes back in about 2 s.

Block-market fields:

| Market | Fields |
|---|---|
| DAM, RTM, HP-DAM | purchase_bid, sell_bid, mcv, final_scheduled_volume, mcp, congestion (HP-DAM mcp is empty on days with no trade) |
| GDAM | purchase_bid, total/hydro/wind/other_re/dre × sell_bid, unconstrained MCV, cleared volume, unconstrained_m_c_p, congestion |

## Understanding the Intra-Day / DAC / TAM data

**Contract types.**
- `INTRADAY` is delivered the same day.
- `DAC` (Day-Ahead Contingency) is traded after DAM for next-day delivery.
- `DAILY` and `WEEKLY` are term-ahead contracts.
- In Green, the contract type can also name the RE source, for example `WIND` or `OTHER RE`.

**Instrument names** are parsed into extra columns where the pattern is clear:

| Instrument | instr_product | instr_resolution | instr_block | instr_suffix |
|---|---|---|---|---|
| `DAC-B01-ER` | DAC | 15MIN | 1 | ER |
| `ITD-H18-WR` | ITD | HOUR | 18 | WR |

The suffix is the region code (NR / ER / WR / SR / AR) in conventional TAM. In Green it's the RE source code (WD = wind, OT = other RE, DR, …). Daily and weekly names such as `DB5-FRI-ER` and `6T18FR2WR` are kept as they are.

**Before 8 May 2025, bid counts and bid volumes are missing.** IEX publishes `-` for them before that date, which is stored as NULL.

**The two views agree.** Summing `total_traded_volume` in the trade details for a day and contract type gives the same number as `total_cleared_volume` in the summary. For 1 Sep 2026 DAC, both are 18,048.89 MWh.

In `gtam_summary`, the field IEX calls `averageClearingVolume` is actually the average **price**. It's stored as `average_clearing_volume` so the source name is kept.

**Ready-made views:**
- `v_intraday_daily`: per segment, trade date and contract type. Gives instruments, trades, traded MWh, VWAP and max price.
- `v_dac_itd_blocks`: DAC and ITD volume and VWAP per delivery block, with all regions summed.

```sql
-- Monthly DAC volume & VWAP, conventional vs green
SELECT segment, substr(trade_date,1,7) AS month, contract_type,
       ROUND(SUM(traded_mwh)) AS mwh, ROUND(SUM(traded_mwh*vwap_rs_mwh)/SUM(traded_mwh),2) AS vwap
FROM v_intraday_daily WHERE contract_type='DAC' GROUP BY 1,2,3 ORDER BY 2,1;

-- Which 15-min DAC blocks clear highest (last 90 days)
SELECT instr_block, ROUND(AVG(vwap_rs_mwh)) AS avg_vwap, ROUND(SUM(traded_mwh)) AS mwh
FROM v_dac_itd_blocks WHERE segment='tam' AND instr_product='DAC' AND instr_resolution='15MIN'
  AND trade_date >= date('now','-90 day') GROUP BY 1 ORDER BY avg_vwap DESC LIMIT 10;
```

**TAM Any-Day auctions:** the list page only gives the date, title, and a link. Result PDFs are not downloaded.

**TAM Monthly:** the volumes (MUs) and discovered price come straight from the list.

## Setup on Windows (one time)

1. Install Python 3.9+ from python.org and tick "Add to PATH".
2. Copy this folder somewhere permanent, for example `C:\iex_pipeline`.
3. Double-click **`setup_windows.bat`**. It creates a `.venv`, installs `requests`, and runs the full backfill of every dataset. It then runs a completeness check and exports the CSVs. Finally, it registers the Task Scheduler job **"IEX Market Data Pipeline"** to run daily at 14:15.
   - The first run is longer than a snapshot-only backfill because it also loads Intra-Day / DAC / TAM history.

To pick a different time, run: `powershell -ExecutionPolicy Bypass -File setup_task.ps1 -At 15:00`

## What each daily run does (`run_pipeline.bat`, about 1–2 minutes)

| Dataset | Window pulled |
|---|---|
| Block markets | last complete day − 5 days → **tomorrow** (DAM, GDAM and HP-DAM results for D+1 are out by about 13:00) |
| Trades and summaries | last loaded day − 7 days → today |
| Auction lists | the full list is re-read |

Each run also:
- **Self-heals.** Missing or partial block days, and any request that failed on an earlier run, are fetched again.
- **Upserts** on each dataset's natural key, so repeated runs never create duplicates.
- **Exports CSVs** to `data\csv\<dataset>\<dataset>_<year>.csv`.

If the PC was off at the scheduled time, the task runs as soon as it's back on.

## Output

- `data\iex_market_data.db`: SQLite tables for each dataset, plus `load_log`. Block tables are keyed by delivery date and block (1–96). Trade tables are keyed by trade date, contract type, and instrument.
- `data\csv\...`: the same data as CSV.
- `logs\iex_pipeline.log`: the run log.

Example query (daily average DAM MCP):
```sql
SELECT delivery_date, ROUND(AVG(mcp),2) AS avg_mcp FROM dam GROUP BY delivery_date ORDER BY delivery_date;
```

## Commands

```
python iex_pipeline.py update                                   # what the scheduler runs (all datasets)
python iex_pipeline.py update --markets intraday auctions       # just the newer markets
python iex_pipeline.py backfill --markets tam --start 2020-04-01 --end 2020-12-31
python iex_pipeline.py check                                    # rows, date range, gaps, failed chunks
python iex_pipeline.py export --out D:\exports
```

Exit codes: 0 = ok, 1 = some requests failed (the next run retries them), 2 = the portal refused script requests.

## If requests get blocked

If IEX adds bot protection and the log shows "got HTTP 403 … instead of JSON / text/x-component", install the browser fallback once:

```
.venv\Scripts\pip install playwright
.venv\Scripts\python -m playwright install chromium
```

`run_pipeline.bat` then retries on its own with `--browser`. That mode makes the same calls from inside a hidden Chromium page.

## Notes

- All dates are IST. Block markets are keyed by **delivery** date, and TAM, GTAM and HP-TAM by **trade** date.
- These are the site's internal endpoints and aren't documented, so they can change without notice. The script fails loudly rather than storing bad data. IEX also offers an **official market-data API**: the portal says to email api_access@iexindia.com. That's the better option if this feeds anything commercial or critical. Please also check IEX's website terms on automated access.
- The pipeline waits 1 s between data calls.
