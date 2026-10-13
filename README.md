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

Python packages used: `pandas`, `numpy`, `xlsxwriter`, `pillow`. The `unrar` command is needed to open `.rar` uploads.

### Input files the user must upload (not in this repo yet)

| Input | Name hint | Example from the NCCD run |
|---|---|---|
| FEG port KPI (detection source, `RxMaxSpeed`) | `FEG` or `FEGE` in the name | `FEG_KPI_NCCD_11 OCT.part1.rar`, `FEG_KPI_NCCD_11 OCT.part2.rar` |
| OSS cell-level hourly KPI (one date = V1/V2 × part1/part2) | `OSS` in the name | `OSS KPI Hourly Cel Level_08Oct_V1.part1.rar` … `OSS KPI Hourly Cel Level_10Oct_V2.part2.rar` |
| Physical site database (lat/long, district, region, thana, Tech, Site Type) | physical site / site database | `Physical_Site_Database_11Oct26.xlsx` |
| Top thana list for GeoPlot | thana | `Top Thana.xlsx` |
| Previous-window FEG file (Comparison sheet only, optional) | `FEG` + earlier date | `FEG_KPI_DHAKA_5Oct.rar` |

### Outputs produced

| File | Contents |
|---|---|
| `FEGE_Choked_Flat_Sites_Ericsson_vN.xlsx` | Dashboard, 1. Site List, 2. HourlyChartOfIssueSites, 3. Hourly KPI, 4. Method, 5. GeoPlot, UrgentTxBWInc, URGENTBWINC_BusyHour, 6. Missing FEG Port, 7. Immediate actions, Comparison (when a previous window exists) |
| `FEGE_Choked_Flat_Sites_Ericsson_vN_GeoPlot.html` | Interactive Leaflet map on Google satellite tiles |

Working files (`/tmp/.../OSS KPI Hourly site  level_capacity.csv`, combined cell CSVs, extracted RARs, `fege_report_cache.pkl`) are not committed.
