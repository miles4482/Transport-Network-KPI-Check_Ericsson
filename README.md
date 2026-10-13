# Transport-Network-KPI-Check_Ericsson

FEGE site-choking (Transmission Link Health Review) report for the Ericsson zone.

All rules are saved in `RULES.md`. They were copied from agent `bc-be4dd871-6780-44f4-807d-7266aad9b277` (NCCD FEGE health report v6, branch `cursor/fege-choke-rules-b277` on `Transport-Network-KPI-Check_NCCD`).

## Files required to prepare a similar report

### Rule files (copied from the NCCD agent, in this repo)

| File | Purpose |
|---|---|
| `RULES.md` | The full v84 rule set: detection, listing, severity, charts, GeoPlot, UrgentTxBWInc, OSS cell-to-site rollup, URGENTBWINC_BusyHour, Site List 4H/1H block, Missing FEG Port, Immediate actions, Comparison, workbook layout |
| `.cursor/rules/fege-choke-rules.mdc` | Always-on Cursor rule that points the agent to `RULES.md` and states how to read uploaded files |
| `rules/V14_LOGIC.md` | Frozen V14 detection baseline (crowded level / hard ceiling shapes) that every later version builds on |

### Scripts (copied from the NCCD agent, in this repo)

| File | Purpose |
|---|---|
| `scripts/fege_choke_check.py` | Reads the FEG and OSS files, runs detection, builds every sheet, writes `FEGE_Choked_Flat_Sites_<zone>_vN.xlsx` |
| `scripts/geoplot_satellite.py` | Renders the national, thana, and cluster satellite maps (JPEG for the GeoPlot sheet, Leaflet HTML) |
| `scripts/fege_choke_check_ericsson.py` | Ericsson runner (v1, v2). Maps the Ericsson export columns onto the rule roles, reuses the detection from `fege_choke_check.py`, and writes Dashboard, 2. HourlyChartOfIssueSites and 3. Hourly KPI only |
| `scripts/fege_peak_touch_ericsson.py` | Ericsson runner (v3). Peak-touch method on `Max DL Capacity Usage (Mbps)`: hours within 5% below the window Peak, graded Severe ≥70% / High ≥50% / Medium ≥40% / Low ≥20%. Same three sheets |

### Ericsson column mapping

| Rule role | Ericsson column in `FEG_KPI_Eric.xlsb` |
|---|---|
| eNodeB Name | `ERBS` |
| RxMaxSpeed (choke trace, Mbit/s) | v2: `Total Capcity,Mbps` only (`--trace total`). v1: `Max DL Capacity Usage (Mbps)` (`--trace dl`) |
| Tx Total BW (reference only, median) | v2: none. v1: `Max UL Capacity Usage(Mbps)` |
| Date / Time | `Date` (Excel serial) / `Hour` (0–23) |

v2 lists only sites with hours on cap ≥ 50% (`--min-hours-pct 50`).

Run (v2 method): `python3 scripts/fege_choke_check_ericsson.py FEG_KPI_Eric.xlsb --trace total --min-hours-pct 50 -o FEGE_Choked_Flat_Sites_Ericsson_vN.xlsx`

Run (v3 Peak-touch method): `python3 scripts/fege_peak_touch_ericsson.py FEG_KPI_Eric.xlsb -o FEGE_Choked_Flat_Sites_Ericsson_vN.xlsx`

Python packages used: `pandas`, `numpy`, `xlsxwriter`, `pillow`. The `unrar` command is needed to open `.rar` uploads.

### Input files

| Input | Name hint | In this repo |
|---|---|---|
| FEG port KPI (detection source) | `FEG` or `FEGE` in the name | `FEG_KPI_Eric.xlsb` (3–5 Oct 2026, uploaded) |
| OSS cell-level hourly KPI (one date = V1/V2 × part1/part2) | `OSS` in the name | not uploaded (NCCD example: `OSS KPI Hourly Cel Level_08Oct_V1.part1.rar`) |
| Physical site database (lat/long, district, region, thana, Tech, Site Type) | physical site / site database | not uploaded (NCCD example: `Physical_Site_Database_11Oct26.xlsx`) |
| Top thana list for GeoPlot | thana | not uploaded (`Top Thana.xlsx`) |
| Previous-window FEG file (Comparison sheet only, optional) | `FEG` + earlier date | not uploaded |

### Outputs produced

| File | Contents |
|---|---|
| `FEGE_Choked_Flat_Sites_Ericsson_vN.xlsx` | Dashboard, 1. Site List, 2. HourlyChartOfIssueSites, 3. Hourly KPI, 4. Method, 5. GeoPlot, UrgentTxBWInc, URGENTBWINC_BusyHour, 6. Missing FEG Port, 7. Immediate actions, Comparison (when a previous window exists) |
| `FEGE_Choked_Flat_Sites_Ericsson_vN_GeoPlot.html` | Interactive Leaflet map on Google satellite tiles |

Working files (`/tmp/.../OSS KPI Hourly site  level_capacity.csv`, combined cell CSVs, extracted RARs, `fege_report_cache.pkl`) are not committed.
