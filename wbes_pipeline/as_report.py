"""
AS (Ancillary Services) schedule report from the WBES Net Schedule Excel files.
===============================================================================

What it answers
---------------
* Day-wise: how much AS each plant was scheduled each day (Up and Down, MWh) and in which time slots,
  next to the plant's whole schedule that day by type (ISGS, URS, OA_GNA, ..., Total).
* Cumulative: the total AS of each plant over the whole date range.
* Time of day: what share of the AS energy falls in each hour and time band, for any date range.

Source
------
wbes_ns_data/<yyyy-mm>/NetSchedule_<date>_revNNN.xlsx, written by wbes_pull.py (one sheet per region,
one row per utility + schedule type, 96 block columns). Only rows with schedule_type "AS" are used.

How the numbers are read (checked against the data; the workbook's Read Me sheet repeats this)
-------------------------------------------------------------------------------------------
* Block values are MW (average over the 15-min block). Energy (MWh) = MW x 0.25 h. 1 MU = 1,000 MWh.
* WBES sign convention: drawal positive, injection negative (a generator's ISGS row is negative).
  So for a plant:  AS value < 0 -> extra injection = AS Up;   AS value > 0 -> backing down = AS Down.
* VAE_ER, VAE_WR, ... (utility_type_id 17) are the Virtual Ancillary Entities, the regional counter-party
  of every AS schedule: in each block, VAE = minus the sum of all plants of that region. They are left
  out of plant figures and used only as a checksum.
* Technology comes from each plant's own "Total" schedule row and its name (see technology()):
  solar = no output at night, wind = output day and night, BESS = charges and discharges,
  hybrid parts = WBES codes ending _HS / _HW, conventional = no renewable (OA_REMC) schedule.
  plant_master.csv (beside this script, editable in Excel) overrides technology and adds nameplate MW.

Run (in the wbes_pipeline folder)
---------------------------------
    .venv\\Scripts\\python.exe as_report.py                                 # every file in wbes_ns_data
    .venv\\Scripts\\python.exe as_report.py --from 2026-06-01 --to 2026-09-30

Output in reports\\ : AS_Schedule_Analysis_<from>_to_<to>.xlsx (Excel computes the formulas when it opens)
and AS_Schedule_Dashboard_<from>_to_<to>.html (interactive charts, works offline).
"""

from __future__ import annotations

import argparse
import datetime as dt
import json
import pickle
import re
from concurrent.futures import ProcessPoolExecutor
from pathlib import Path

import numpy as np
import pandas as pd

from wbes_pull import block_desc, log

HERE = Path(__file__).resolve().parent
NS_DIR = HERE / "wbes_ns_data"
OUT_DIR = HERE / "reports"
HTML_TEMPLATE = HERE / "as_dashboard_template.html"
PLANT_MASTER = HERE / "plant_master.csv"          # optional overrides: technology, nameplate MW, source

BLOCK_HOURS = 0.25                 # one block = 15 min
VAE_TYPE = 17
TYPE_NAMES = {4: "ISGS / central generator", 16: "IPP / RE / storage", 5: "State control area",
              VAE_TYPE: "VAE (regional AS pool)"}
REGION_ORDER = ["ER", "WR", "NR", "SR", "NER"]
# Time bands by the hour the block starts in. Bands must cover all 24 hours exactly once.
BANDS = [("Night 23:00-06:00", [23, 0, 1, 2, 3, 4, 5]),
         ("Morning 06:00-09:00", [6, 7, 8]),
         ("Solar 09:00-17:00", list(range(9, 17))),
         ("Evening 17:00-23:00", list(range(17, 23)))]
VAE_TOLERANCE_MW = 0.5             # plants + VAE should cancel to ~0 in every block
# Technology rules (see technology()). Night = 00:00-05:30 and 19:00-24:00, when solar cannot generate.
NIGHT_BLOCKS = (np.arange(96) < 22) | (np.arange(96) >= 76)
SOLAR_MAX_NIGHT_SHARE = 0.02       # solar: under 2 % of its energy at night (seen: 0-0.8 %)
WIND_MIN_NIGHT_SHARE = 0.35        # wind / round-the-clock: over 35 % at night (seen: 39-65 %)
CONVENTIONAL = "Conventional (thermal / hydro)"
# Schedule types in WBES order; any other type found is added after these.
SCHEDULE_ORDER = ["ISGS", "URS", "OA_GNA", "OA_TGNA", "OA_PX", "OA_REMC", "AS", "SCED", "SCUC", "Total"]
TECH_ORDER = ["Solar", "Wind", "Hybrid - solar part", "Hybrid - wind part", "BESS (standalone)", CONVENTIONAL,
              "State control area", "Unclear"]
# Thermal plants that also carry an OA_REMC row, so the schedule shape alone would call them wind.
KNOWN_CONVENTIONAL = {"APL_Raipur_TPP", "APL_Raigarh_TPP", "JPNIGRIE_JNSTPP", "DGEN", "MAHAN_ENERGEN",
                      "MEENAKSHI", "JHABUA_IPP"}
WORKERS = 8

FILE_RE = re.compile(r"NetSchedule_(\d{4}-\d{2}-\d{2})_rev(\d+)\.xlsx$")
BLOCK_LABELS = [block_desc(n) for n in range(1, 97)]
HOUR_OF_BLOCK = np.arange(96) // 4
assert sorted(h for _, hrs in BANDS for h in hrs) == list(range(24))


# ----------------------------------------------------------------------------- read
def read_file(path: Path) -> tuple[dict, pd.DataFrame, pd.DataFrame]:
    """One Net Schedule file: AS + Total rows (96 blocks), daily MWh of every schedule row, file info."""
    m = FILE_RE.match(path.name)
    d, rev = dt.date.fromisoformat(m.group(1)), int(m.group(2))
    with pd.ExcelFile(path) as xl:
        info = xl.parse("Info").set_index("Field")["Value"]
        if (pd.to_datetime(info.get("Date"), dayfirst=True).date() != d
                or int(info.get("Revision")) != rev):
            raise ValueError(f"{path.name}: Info sheet says {info.get('Date')} rev {info.get('Revision')}")
        parts, mix = [], []
        for region in xl.sheet_names[1:]:
            df = xl.parse(region)
            if list(df.columns[3:]) != BLOCK_LABELS:
                raise ValueError(f"{path.name} {region}: unexpected block columns")
            remc = set(df.loc[df["schedule_type"] == "OA_REMC", "utility"])
            keep = df[df["schedule_type"].isin(["AS", "Total"])]
            parts.append(keep.assign(region=region, has_remc=keep["utility"].isin(remc)))
            mw = df[BLOCK_LABELS].to_numpy(dtype=float)
            mix.append(pd.DataFrame({"region": region, "plant": df["utility"].to_numpy(),
                                     "schedule_type": df["schedule_type"].to_numpy(),
                                     "mwh": -np.nansum(mw, axis=1) * BLOCK_HOURS}))   # WBES: injection < 0
    rows = pd.concat(parts, ignore_index=True)
    mix = pd.concat(mix, ignore_index=True)
    mix.insert(0, "date", pd.Timestamp(d))
    rows.insert(0, "date", pd.Timestamp(d))
    meta = {"date": pd.Timestamp(d), "file": path.name, "revision": rev,
            "exported_at": str(info.get("Exported At")), "regions": ",".join(xl.sheet_names[1:])}
    return meta, rows, mix


def load(ns_dir: Path, start: dt.date | None, end: dt.date | None):
    """Read every file in range (highest revision per date). Returns files, AS/Total rows, mix, problems."""
    found: dict[dt.date, list[tuple[int, Path]]] = {}
    for f in ns_dir.glob("*/NetSchedule_*_rev*.xlsx"):
        m = FILE_RE.match(f.name)
        if not m:
            continue
        d = dt.date.fromisoformat(m.group(1))
        if (start is None or d >= start) and (end is None or d <= end):
            found.setdefault(d, []).append((int(m.group(2)), f))
    if not found:
        raise SystemExit(f"No NetSchedule files in {ns_dir} for that range")
    problems = []
    for d, revs in sorted(found.items()):
        if len(revs) > 1:
            problems.append(f"{d}: {len(revs)} files, used the highest revision {max(revs)[0]}")
    paths = [max(revs)[1] for _, revs in sorted(found.items())]
    first, last = start or min(found), end or max(found)
    missing = [first + dt.timedelta(days=i) for i in range((last - first).days + 1)
               if first + dt.timedelta(days=i) not in found]
    problems += [f"{d}: no file (date missing)" for d in missing]

    log(f"Reading {len(paths)} files from {ns_dir} ...")
    with ProcessPoolExecutor(WORKERS) as pool:
        results = list(pool.map(read_file, paths))
    files = pd.DataFrame([m for m, _, _ in results])
    rows = pd.concat([r for _, r, _ in results], ignore_index=True)
    mix = pd.concat([x for _, _, x in results], ignore_index=True)
    rows = rows.rename(columns={"utility": "plant", "utility_type_id": "type_id"})
    rows[BLOCK_LABELS] = rows[BLOCK_LABELS].astype(float).fillna(0.0)
    return files, rows, mix, problems


# ----------------------------------------------------------------------------- aggregate
def runs_text(mask: np.ndarray) -> str:
    """Contiguous True blocks as '18:00-22:30, 23:15-24:00'."""
    out, i = [], 0
    while i < 96:
        if mask[i]:
            j = i
            while j + 1 < 96 and mask[j + 1]:
                j += 1
            out.append(f"{block_desc(i + 1)[:5]}-{block_desc(j + 1)[6:]}")
            i = j + 1
        else:
            i += 1
    return ", ".join(out)


def plant_profile(total: pd.DataFrame) -> pd.DataFrame:
    """Shape of each plant's own Total schedule over the period: the evidence for its technology."""
    mw = total[BLOCK_LABELS].to_numpy()
    inj, drw = np.clip(-mw, 0, None), np.clip(mw, 0, None)        # injection is negative in WBES
    f = pd.DataFrame({"plant": total["plant"].to_numpy(), "inj": inj.sum(axis=1),
                      "inj_night": inj[:, NIGHT_BLOCKS].sum(axis=1), "drw": drw.sum(axis=1),
                      "max_inj": inj.max(axis=1), "max_drw": drw.max(axis=1), "has_remc": total["has_remc"].to_numpy()})
    g = f.groupby("plant").agg(inj=("inj", "sum"), inj_night=("inj_night", "sum"), drw=("drw", "sum"),
                               max_sched_mw=("max_inj", "max"), max_charge_mw=("max_drw", "max"),
                               has_remc=("has_remc", "any"))
    g["night_share"] = (g["inj_night"] / g["inj"]).where(g["inj"] > 0)
    g["charge_ratio"] = (g["drw"] / g["inj"]).where(g["inj"] > 0)
    return g[["night_share", "charge_ratio", "max_sched_mw", "max_charge_mw", "has_remc"]]


def twin_code(plant: str, names: set[str]) -> str | None:
    """The other leg of a hybrid: same WBES code with its S token swapped for W (or W for S), if it exists."""
    parts = plant.split("_")
    upper = {n.upper(): n for n in names}
    for i, tok in enumerate(parts):
        if tok.upper() in ("S", "W"):
            other = "_".join(parts[:i] + ["W" if tok.upper() == "S" else "S"] + parts[i + 1:])
            if other.upper() in upper:
                return upper[other.upper()]
    return None


def technology(plant: str, type_id: int, p, twin: str | None = None) -> tuple[str, str]:
    """(technology, evidence) for one plant; p is its plant_profile() row, twin its twin_code()."""
    u, ns, cr = plant.upper(), p.night_share, p.charge_ratio
    shape = f"{ns:.1%} of its output at night" if ns == ns else "no injection scheduled in the period"
    if type_id == 5:
        return "State control area", "WBES utility type 5 (state)"
    if "BESS" in u:
        return "BESS (standalone)", f"name has BESS; it draws {cr:.0%} of what it injects (charging)"
    if u.endswith("_HS"):
        return "Hybrid - solar part", f"code ends _HS; {shape}"
    if u.endswith("_HW"):
        return "Hybrid - wind part", f"code ends _HW; {shape}"
    if twin and p.has_remc and ns == ns:
        if ns <= SOLAR_MAX_NIGHT_SHARE:
            return "Hybrid - solar part", f"solar shape ({shape}); same code has a wind leg: {twin}"
        if ns >= WIND_MIN_NIGHT_SHARE:
            return "Hybrid - wind part", f"wind shape ({shape}); same code has a solar leg: {twin}"
    if plant in KNOWN_CONVENTIONAL:
        return CONVENTIONAL, "known thermal plant (has an OA_REMC row, but is not renewable)"
    if not p.has_remc:
        return CONVENTIONAL, "no renewable (OA_REMC) schedule"
    if ns <= SOLAR_MAX_NIGHT_SHARE:
        return "Solar", shape
    if ns >= WIND_MIN_NIGHT_SHARE:
        return "Wind", shape
    return "Unclear", shape


MASTER_COLS = ["plant", "technology", "nameplate_mw", "ac_or_dc", "bess_mwh", "project_name", "developer",
               "source_url", "confidence", "note"]


def load_master() -> pd.DataFrame:
    """plant_master.csv: one row per plant; a filled 'technology' overrides the automatic one."""
    if not PLANT_MASTER.exists():
        return pd.DataFrame(columns=MASTER_COLS).set_index("plant")
    m = pd.read_csv(PLANT_MASTER, dtype=str, encoding="utf-8-sig").fillna("")
    missing = [c for c in MASTER_COLS if c not in m.columns]
    if missing:
        raise SystemExit(f"{PLANT_MASTER.name} is missing columns {missing}")
    bad = sorted(set(m["technology"]) - set(TECH_ORDER) - {""})
    if bad:
        raise SystemExit(f"{PLANT_MASTER.name}: unknown technology {bad}; use one of {TECH_ORDER}")
    for c in ("nameplate_mw", "bess_mwh"):
        m[c] = pd.to_numeric(m[c], errors="coerce")
    return m[MASTER_COLS].drop_duplicates("plant", keep="last").set_index("plant")


def aggregate(files: pd.DataFrame, rows: pd.DataFrame, mix: pd.DataFrame) -> dict:
    total = rows[rows["schedule_type"] == "Total"]
    rows = rows[rows["schedule_type"] == "AS"]
    vae = rows[rows["type_id"] == VAE_TYPE]
    pl = rows[rows["type_id"] != VAE_TYPE].reset_index(drop=True)

    mw = pl[BLOCK_LABELS].to_numpy()
    up_mw, dn_mw = np.clip(-mw, 0, None), np.clip(mw, 0, None)
    up_h = np.stack([up_mw[:, HOUR_OF_BLOCK == h].sum(axis=1) for h in range(24)], axis=1) * BLOCK_HOURS
    dn_h = np.stack([dn_mw[:, HOUR_OF_BLOCK == h].sum(axis=1) for h in range(24)], axis=1) * BLOCK_HOURS

    day = pd.DataFrame({
        "date": pl["date"], "month": pl["date"].dt.strftime("%Y-%m"), "region": pl["region"],
        "plant": pl["plant"], "type_id": pl["type_id"], "type": pl["type_id"].map(TYPE_NAMES).fillna("Other"),
        "up_mwh": up_h.sum(axis=1), "down_mwh": dn_h.sum(axis=1),
        "up_blocks": (up_mw > 0).sum(axis=1), "down_blocks": (dn_mw > 0).sum(axis=1),
        "peak_up_mw": up_mw.max(axis=1), "peak_down_mw": dn_mw.max(axis=1),
        "up_slots": [runs_text(r > 0) for r in up_mw], "down_slots": [runs_text(r > 0) for r in dn_mw],
    })
    for name, hrs in BANDS:
        day[f"up|{name}"] = up_h[:, hrs].sum(axis=1)
    for name, hrs in BANDS:
        day[f"down|{name}"] = dn_h[:, hrs].sum(axis=1)
    hours = pd.DataFrame(np.hstack([up_h, dn_h]),
                         columns=[f"up_h{h:02d}" for h in range(24)] + [f"down_h{h:02d}" for h in range(24)])
    day = pd.concat([day, hours], axis=1)
    reg_rank = {r: i for i, r in enumerate(REGION_ORDER)}
    day = day.sort_values(["date", "region", "plant"], key=lambda s: s.map(reg_rank) if s.name == "region" else s,
                          kind="stable").reset_index(drop=True)

    # Checksum: in every date + region + block, plants + VAE must cancel.
    p_sum = pl.groupby(["date", "region"])[BLOCK_LABELS].sum()
    v_sum = vae.groupby(["date", "region"])[BLOCK_LABELS].sum().reindex(p_sum.index, fill_value=0.0)
    check = pd.DataFrame({"plants": pl.groupby(["date", "region"]).size(),
                          "vae_row": vae.groupby(["date", "region"]).size().reindex(p_sum.index, fill_value=0),
                          "max_diff_mw": (p_sum + v_sum).abs().max(axis=1)}).reset_index()
    check["status"] = np.where(check["max_diff_mw"] <= VAE_TOLERANCE_MW, "OK", "CHECK")

    plants = (day.groupby(["region", "plant", "type", "type_id"], as_index=False)
              .agg(up=("up_mwh", "sum"), down=("down_mwh", "sum")))
    plants["gross"] = plants["up"] + plants["down"]
    plants = plants.sort_values(["gross", "plant"], ascending=[False, True]).reset_index(drop=True)

    # Technology and capacity: schedule shape + name, then plant_master.csv overrides
    prof = plant_profile(total[total["plant"].isin(set(plants["plant"]))])
    plants = plants.join(prof, on="plant")
    all_names = set(mix["plant"])
    auto = [technology(p.plant, p.type_id, p, twin_code(p.plant, all_names)) for p in plants.itertuples()]
    plants["technology"] = [a[0] for a in auto]
    plants["tech_evidence"] = [a[1] for a in auto]
    master = load_master()
    plants = plants.join(master.drop(columns="technology"), on="plant")
    set_tech = plants["plant"].map(master["technology"]).fillna("")
    override = set_tech.ne("") & set_tech.ne(plants["technology"])
    plants.loc[override, "tech_evidence"] = ("set in plant_master.csv (automatic: " + plants.loc[override, "technology"]
                                             + "; " + plants.loc[override, "tech_evidence"] + ")")
    plants.loc[set_tech.ne(""), "technology"] = set_tech[set_tech.ne("")]
    day["technology"] = day["plant"].map(plants.set_index("plant")["technology"])

    # Whole schedule of the AS plants, every day they appear: MWh per schedule type, + = into the grid
    m = mix[mix["plant"].isin(set(plants["plant"]))]
    sched = m.pivot_table(index=["date", "region", "plant"], columns="schedule_type", values="mwh",
                          aggfunc="sum", fill_value=0.0)
    stypes = [x for x in SCHEDULE_ORDER if x in sched.columns] + sorted(set(sched.columns) - set(SCHEDULE_ORDER))
    sched = sched[stypes].reset_index()
    sched.insert(1, "month", sched["date"].dt.strftime("%Y-%m"))
    sched.insert(4, "technology", sched["plant"].map(plants.set_index("plant")["technology"]))
    parts = [x for x in stypes if x != "Total"]
    sched["check"] = sched[parts].sum(axis=1) - sched["Total"]      # types must add up to Total
    sched = sched.sort_values(["date", "region", "plant"], kind="stable").reset_index(drop=True)
    tot = sched.set_index(["date", "plant"])["Total"]
    day["total_mwh"] = [tot.get((d, p), np.nan) for d, p in zip(day["date"], day["plant"])]
    net = day["up_mwh"] - day["down_mwh"]
    day["as_share"] = (net / day["total_mwh"]).where(day["total_mwh"].abs() > 1e-9)

    per_file = day.groupby("date").agg(plants=("plant", "size"), up=("up_mwh", "sum"), down=("down_mwh", "sum"))
    files = files.merge(per_file, left_on="date", right_index=True, how="left")
    return {"day": day, "plants": plants, "check": check, "files": files, "sched": sched, "stypes": stypes,
            "dates": sorted(day["date"].unique()), "generated": dt.datetime.now()}


# ----------------------------------------------------------------------------- HTML
def html_data(t: dict) -> dict:
    """Compact data for the dashboard: one record per plant-day, hours stored as sparse pairs."""
    day, plants = t["day"], t["plants"]
    dates = [pd.Timestamp(d).strftime("%Y-%m-%d") for d in t["dates"]]
    d_idx = {d: i for i, d in enumerate(dates)}
    p_idx = {p: i for i, p in enumerate(plants["plant"])}
    types = list(dict.fromkeys(plants["type"]))
    techs = [x for x in TECH_ORDER if x in set(plants["technology"])]
    num = lambda v: None if v != v else round(float(v), 1)
    url = lambda v: v if isinstance(v, str) and v.startswith("http") else None

    def sparse(vals) -> list:
        return [x for h, v in enumerate(vals) if v > 0 for x in (h, round(float(v), 4))]   # exact: MW (2 dp) x 0.25

    up_cols, dn_cols = [f"up_h{h:02d}" for h in range(24)], [f"down_h{h:02d}" for h in range(24)]
    rec = [[d_idx[r.date.strftime("%Y-%m-%d")], p_idx[r.plant], sparse(u), sparse(dn),
            r.up_slots, r.down_slots, round(float(r.peak_up_mw), 2), round(float(r.peak_down_mw), 2)]
           for r, u, dn in zip(day.itertuples(), day[up_cols].to_numpy(), day[dn_cols].to_numpy())]
    return {"meta": {"from": dates[0], "to": dates[-1], "days": len(dates), "files": len(t["files"]),
                     "generated": t["generated"].strftime("%d-%b-%Y %H:%M"), "source": "wbes_ns_data",
                     "bands": [[n, hrs] for n, hrs in BANDS]},
            "dates": dates, "regions": REGION_ORDER, "types": types, "techs": techs, "stypes": t["stypes"],
            "mix": [[d_idx[r[0].strftime("%Y-%m-%d")], p_idx[r[1]]] + [round(float(v), 4) for v in r[2:]]
                    for r in t["sched"][["date", "plant"] + t["stypes"]].itertuples(index=False)],
            "plants": [[p.plant, p.region, types.index(p.type), techs.index(p.technology), num(p.max_sched_mw),
                        num(p.nameplate_mw), url(p.source_url)] for p in plants.itertuples()],
            "rec": rec}


def write_html(t: dict, path: Path) -> Path:
    data = json.dumps(html_data(t), separators=(",", ":"))
    page = HTML_TEMPLATE.read_text(encoding="utf-8").replace("__AS_DATA__", data)
    path.write_text(page, encoding="utf-8")
    return path


# ----------------------------------------------------------------------------- main
def main() -> None:
    ap = argparse.ArgumentParser(description="AS schedule report from WBES Net Schedule files")
    ap.add_argument("--from", dest="date_from", type=dt.date.fromisoformat, help="first date, YYYY-MM-DD")
    ap.add_argument("--to", dest="date_to", type=dt.date.fromisoformat, help="last date, YYYY-MM-DD")
    ap.add_argument("--src", type=Path, default=NS_DIR, help="folder with the NetSchedule files")
    ap.add_argument("--cache", type=Path, help="save / reuse the parsed data in this pickle (faster reruns)")
    ap.add_argument("--no-excel", action="store_true")
    ap.add_argument("--no-html", action="store_true")
    args = ap.parse_args()

    if args.cache and args.cache.exists():
        files, rows, mix, problems = pickle.loads(args.cache.read_bytes())
        log(f"Using cached data from {args.cache}")
    else:
        files, rows, mix, problems = load(args.src, args.date_from, args.date_to)
        if args.cache:
            args.cache.write_bytes(pickle.dumps((files, rows, mix, problems)))
    t = aggregate(files, rows, mix)
    t["problems"] = problems
    for p in problems:
        log(f"NOTE {p}")
    bad = t["check"][t["check"]["status"] != "OK"]
    off = t["sched"]["check"].abs() > 0.5
    log(f"Schedule types {t['stypes']}: {int(off.sum())} plant-days where the types do not add up to Total")
    log("Technology: " + ", ".join(f"{k} {v}" for k, v in t["plants"]["technology"].value_counts().items()))
    log(f"{len(t['dates'])} dates, {t['day']['plant'].nunique()} plants, {len(t['day'])} plant-days; "
        f"VAE checksum: {len(t['check']) - len(bad)} of {len(t['check'])} date-regions OK")

    OUT_DIR.mkdir(exist_ok=True)
    tag = f"{pd.Timestamp(t['dates'][0]):%Y-%m-%d}_to_{pd.Timestamp(t['dates'][-1]):%Y-%m-%d}"
    if not args.no_excel:
        from as_report_excel import write_excel
        log(f"Excel -> {write_excel(t, OUT_DIR / f'AS_Schedule_Analysis_{tag}.xlsx', BANDS, REGION_ORDER, TECH_ORDER)}")
    if not args.no_html:
        log(f"HTML  -> {write_html(t, OUT_DIR / f'AS_Schedule_Dashboard_{tag}.html')}")


if __name__ == "__main__":
    main()
