# Transport-Network-KPI-Check_Ericsson

FEGE site-choking (Transmission Link Health Review) report for the Ericsson zone.

All rules are saved in `RULES.md`. The operative set is the section **Current Ericsson report — v31**, for the Excel workbook and the chart page. The same rules are written on the workbook sheet `Method`. Later sections are the history copied from agent `bc-be4dd871-6780-44f4-807d-7266aad9b277` (NCCD FEGE health report v6, branch `cursor/fege-choke-rules-b277` on `Transport-Network-KPI-Check_NCCD`). Where they disagree, v30 wins.

## Files required to prepare a similar report

### Rule files (copied from the NCCD agent, in this repo)

| File | Purpose |
|---|---|
| `RULES.md` | Current v30 rules for the Excel workbook and the chart page, then the older v84 history |
| `.cursor/rules/fege-choke-rules.mdc` | Always-on Cursor rule that points the agent to `RULES.md` and states how to read uploaded files |
| `rules/V14_LOGIC.md` | Frozen V14 detection baseline (crowded level / hard ceiling shapes) that every later version builds on |

### Scripts (copied from the NCCD agent, in this repo)

| File | Purpose |
|---|---|
| `scripts/fege_choke_check.py` | Reads the FEG and OSS files, runs detection, builds every sheet, writes `FEGE_Choked_Flat_Sites_<zone>_vN.xlsx` |
| `scripts/geoplot_satellite.py` | Renders the national, thana, and cluster satellite maps (JPEG for the GeoPlot sheet, Leaflet HTML) |
| `scripts/fege_choke_check_ericsson.py` | Ericsson runner (v1, v2). Maps the Ericsson export columns onto the rule roles, reuses the detection from `fege_choke_check.py`, and writes Dashboard, 2. HourlyChartOfIssueSites and 3. Hourly KPI only |
| `scripts/fege_peak_touch_ericsson.py` | Ericsson runner (v3 onward). Peak-touch method on `Max DL Capacity Usage (Mbps)`, license section on Dashboard, Site List Tech and Site Type, and 5. GeoPlot |

### Ericsson column mapping

| Rule role | Ericsson column in `FEG_KPI_Eric.xlsb` |
|---|---|
| eNodeB Name | `ERBS` |
| RxMaxSpeed (choke trace, Mbit/s) | v2: `Total Capcity,Mbps` only (`--trace total`). v1: `Max DL Capacity Usage (Mbps)` (`--trace dl`) |
| Tx Total BW (reference only, median) | v2: none. v1: `Max UL Capacity Usage(Mbps)` |
| Date / Time | `Date` (Excel serial) / `Hour` (0–23) |

v2 lists only sites with hours on cap ≥ 50% (`--min-hours-pct 50`).

Run (v2 method): `python3 scripts/fege_choke_check_ericsson.py FEG_KPI_Eric.xlsb --trace total --min-hours-pct 50 -o FEGE_Choked_Flat_Sites_Ericsson_vN.xlsx`

Run (v22: Tx Flat/Chocked first, then Peak Touch and License usage): `python3 scripts/fege_peak_touch_ericsson.py FEG_KPI_Eric.xlsb -o FEGE_Choked_Flat_Sites_Ericsson_vN.xlsx`

v6 lists a license site when `Max DL Capacity Usage / dlBbCapacityNet` is at least 85% for at least 40% of its hours.

Python packages used: `pandas`, `numpy`, `xlsxwriter`, `pillow`. The `unrar` command is needed to open `.rar` uploads.

### Input files

| Input | Name hint | In this repo |
|---|---|---|
| FEG port KPI (detection source) | `FEG` or `FEGE` in the name | `FEG_KPI_Eric.xlsb` (11–13 Oct 2026, uploaded) |
| OSS cell-level hourly KPI (one date = V1/V2 × part1/part2) | `OSS` in the name | `OSS KPI Hourly Cel Level_11Oct`, `_12Oct`, and `_13Oct` (V1 and V2, part1 and part2) |
| Physical site database (lat/long, district, region, thana, Tech, Site Type) | physical site / site database | `Physical_Site_Database_14Oct26.xlsx` |
| DL capacity license (`NodeId`, `dlBbCapacityNet`) | license | `DL_License_File.xlsb` (uploaded; older `DL_License_File.xlsx` kept) |
| Top thana list for GeoPlot | thana | `Top Thana.xlsx` |
| Previous-window FEG file (Comparison sheet only, optional) | `FEG` + earlier date | not uploaded |

### Outputs produced

| File | Contents |
|---|---|
| `FEGE_Choked_Flat_Sites_Ericsson_vN.xlsx` | Dashboard, 1. Site List_ERBS, 1. Site List_SiteLevel, 2. HourlyChartOfIssueSites, 3. Hourly KPI, 5. GeoPlot, URGENTBWINC_BusyHour, 7. Immediate actions. Comparison waits for a previous-window FEG file. Sheet 6. Missing FEG Port is not written from v29 |
| `FEGE_Choked_Flat_Sites_Ericsson_v23.xlsx` | v22 ERBS dashboard and `1. Site List` kept. Site Level Dashboard added underneath. New sheet `1. Site List_SiteLevel`. Site charts tagged on `2. HourlyChartOfIssueSites`. GeoPlot includes a physical site when any ERBS is an issue. The chart-page HTML stays at 22.6 |
| `FEGE_Choked_Flat_Sites_Ericsson_v24.xlsx` | v23 plus TP Range cell counts on `1. Site List` and `1. Site List_SiteLevel`, after each 4h and 1h OSS block. The chart-page HTML stays at 22.6 |
| `FEGE_Choked_Flat_Sites_Ericsson_v25.xlsx` | v24 plus the TP Range cell-count table on `2. HourlyChartOfIssueSites`, to the right of the OSS KPI table. 1 BusyHour pools 3 hours and 4 BusyHour pools 12 hours. The chart-page HTML stays at 22.6 |
| `FEGE_Choked_Flat_Sites_Ericsson_v26.xlsx` | v25 plus per-ERBS TP Range counts, including L900. `1. Site List` gains an ERBS Band column. OSS 4H/1H still drops L900. The chart-page HTML stays at 22.6 |
| `FEGE_Choked_Flat_Sites_Ericsson_v27.xlsx` | v26 with the ERBS sheet renamed `1. Site List_ERBS`. ERBS Band drops the word cells. `1. Site List_SiteLevel` gains a Band column. The chart-page HTML stays at 22.6 |
| `FEGE_Choked_Flat_Sites_Ericsson_v28.xlsx` | v27 with L900 removed from the TP Range. An ERBS that has only L09 cells has a blank TP Range. ERBS Band and the site Band are unchanged |
| `FEGE_Choked_Flat_Sites_Ericsson_v29.xlsx` | v28 with sheet `6. Missing FEG Port` removed. `7. Immediate actions` is rebuilt from `1. Site List_SiteLevel` and the BH Flag 1H blocks only. A site is counted on the first matching priority row and not again |
| `FEGE_Choked_Flat_Sites_Ericsson_v30.xlsx` | v29 plus 10 non-issue sites at the bottom of `1. Site List_ERBS` and `1. Site List_SiteLevel`, with comparison charts tagged Non Issue. Dashboard, GeoPlot, and Immediate actions are unchanged |
| `FEGE_Choked_Flat_Sites_Ericsson_v31.xlsx` | v30 plus sheet `Method`. That sheet holds the Excel rules and the chart-page rules |
| `FEGE_Choked_Flat_Sites_Ericsson_vN_GeoPlot.html` | Interactive Leaflet map on Google satellite tiles |
| `FEGE_Choked_Flat_Sites_Ericsson_v22.6_Charts.html` | Earlier chart page for Excel v22. One ERBS list only |
| `FEGE_Choked_Flat_Sites_Ericsson_v28_Charts.html` | Chart page built for Excel v28. Two tabs: **ERBS** and **SiteLevel** |
| `FEGE_Choked_Flat_Sites_Ericsson_v29_Charts.html` | The v28 chart page, renamed to the v29 Excel version. The charts were not rebuilt |
| `FEGE_Choked_Flat_Sites_Ericsson_v30_Charts.html` | The same chart page, renamed to the v30 Excel version. The charts were not rebuilt |
| `FEGE_Choked_Flat_Sites_Ericsson_v31_Charts.html` | The same chart page, renamed to the v31 Excel version. The charts were not rebuilt |
| `FEGE_Choked_Flat_Sites_Ericsson_v30_Charts.zip` | The same page, zipped so it can be downloaded. Unzip it and open the `.html` file. GitHub shows a raw `.html` as plain text, so that link does not run the page |
| `FEGE_Choked_Flat_Sites_Ericsson_v31_Charts.zip` | The v31-named chart page, zipped. Unzip it and open the `.html` file |

Chart page: `python3 scripts/fege_site_charts_html.py FEGE_Choked_Flat_Sites_Ericsson_v28.xlsx -o FEGE_Choked_Flat_Sites_Ericsson_v28_Charts.html`

The workbook is about 70 MB. GitHub’s file page will not download it (“we can’t show files that are this big”). Use the `raw.githubusercontent.com` link for that file.

Working files (`/tmp/.../OSS KPI Hourly site  level_capacity.csv`, combined cell CSVs, extracted RARs, `fege_report_cache.pkl`) are not committed.
