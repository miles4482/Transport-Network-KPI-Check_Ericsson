# FEGE site choking rules (v84) — Ericsson

This file is the saved copy of every rule used for the FEGE choke report.

Copied from agent `bc-be4dd871-6780-44f4-807d-7266aad9b277` (“NCCD FEGE health report”), branch `cursor/fege-choke-rules-b277` on `Transport-Network-KPI-Check_NCCD`. Last published NCCD file: `FEGE_Choked_Flat_Sites_NCCD_v6.xlsx`. That agent in turn copied the v84 rules from agent `bc-a7709455-0c4d-48fe-8e9d-0c3e0f5a86b3` (“Fege site choking detection_Huawei”), branch `cursor/fege-choked-sites-86b3` on `Transport-Network-KPI-Check`.

This repo is the **Ericsson** report. Use these same rules. Do not build a report until the Ericsson files are in the repo. When they arrive, score the dates in those files and keep every threshold below. The Dhaka reference window was **5–7 Oct 2026**; the NCCD reference window was **8–10 Oct 2026**.

The counter names below (`VS.FEGE.RxMaxSpeed`, `VS.FEGE.TxTotalBW`, the OSS `_N` / `_D` columns, `MID(Cell Name, 8, 3)`) are the Huawei export names. When the Ericsson export uses different column names, map them to the same roles. Do not change the thresholds or the shapes.

## Ericsson settings (user instructions, 13 Oct 2026)

Source file: `FEG_KPI_Eric.xlsb`. The 13 Oct upload had columns `Date`, `Hour`, `ERBS`, `Max DL Capacity Usage (Mbps)`, `Max UL Capacity Usage(Mbps)`, `Total Capcity,Mbps`, window 3–5 Oct 2026. The 14 Oct upload (v14) has `Date`, `Hour`, `ERBS`, `Max DL Capacity Usage (Mbps)`, `Max UL Capacity Usage (Mbps)`, window **11–13 Oct 2026**, every hour. `Total Capcity,Mbps` is not in that file. The Peak-touch method does not use it.

| Version | RxMaxSpeed role | Reference column | Listing floor | Sheets |
|---|---|---|---|---|
| v1 | `Max DL Capacity Usage (Mbps)` | `Max UL Capacity Usage(Mbps)` median | none (all four severities) | Dashboard, 2. HourlyChartOfIssueSites, 3. Hourly KPI |
| v2 | `Total Capcity,Mbps` only | none shown (Max Total Capacity shown instead) | hours on cap ≥ **50%** (Severe only) | same three sheets |
| v3 | `Max DL Capacity Usage (Mbps)` only | none | Peak-touch method (below) | same three sheets |

v2: read **only** `Total Capcity,Mbps` as the choke trace and show no other counter. List a site only when its hours on cap are at least 50% of the window. Detection shapes, tolerances, and the day / last-day rules are unchanged; the floor is applied after detection. Runner: `scripts/fege_choke_check_ericsson.py --trace total --min-hours-pct 50`.

### v3 Peak-touch method (replaces the v84 shape detection for Ericsson)

Runner: `scripts/fege_peak_touch_ericsson.py`. Counter: only `Max DL Capacity Usage (Mbps)`.

- **Peak** = the highest hourly value of the window (the red Peak line on each chart).
- **Below5%Peak** = Peak × 95%. An hour **touches** when Max DL is at or above Below5%Peak (within 5% below the Peak, the orange line). Below5%Peak is one number, 5% under that site's own highest hour. It is not an average of the hours. v3–v9 called this line touch levelbelow5%.
- **Touch ratio** = touch hours ÷ total hours with data for that site.
- Peak ratio = touch hours ÷ hours with data (an hour touches when Max DL is at or above Below5%Peak). Severity (v17): **Severe ≥ 40%**, **High 30% to under 40%**, **Medium 20% to under 30%**, **Low 10% to under 20%**. Under 10% is not listed. v16 used Severe ≥ 50%, High 40% to under 50%, Medium 30% to under 40%, Low 20% to under 30%. v10–v15 used Severe ≥ 60%. v3–v9 used Severe ≥ 70%.
- **Tx Flat/Chocked** (v18): a listed site whose Gap = (License − Peak) / License is at least 10%. Severity is Severe ≥ 40%, High 30% to under 40%, Medium 20% to under 30%, Low 15% to under 20%, Decision 10% to under 15%. The comment is Tx Issues Suspected. v17 used one label, Tx Issues Suspected, for a gap of at least 15%.
- **License usage** (v18): hours at ≥ 85% of license. Severe ≥ 50% of hours, High 40% to under 50%, Medium 30% to under 40%, Low 20% to under 30%. Under 20% is not listed. v6–v17 listed a site at ≥ 40% of hours and grouped ≥ 70% / 50% to under 70% / 40% to under 50%.
- Dashboard category table has Peak Touch, License usage, and Tx Flat/Chocked. Sheet 1. Site List and 2. HourlyChartOfIssueSites hold each ERBS once. 5. GeoPlot uses Severe / High / Moderate / Low / Decision. Medium is drawn as Moderate.
- **BH Flag 1H** (v19, Site List only): after Severity, two columns. DL Throughput 1h (Mbps) is Yes when either DataBusyHour (1 BusyHour) or UserBusyHour (1 BusyHour) DL Throughput is ≤ 7 Mbps. Data Volume 1h (GB) is Yes when either block's Data Volume is ≥ 50 GB. Both values missing is blank. This flag does not add or remove sites. Site Type sits immediately after Tech.
- Guards kept from section 4: at least 24 hourly samples and Peak ≥ 20 Mbit/s.
- Busy-hour touch (08:00–22:00) and Where touched (Busy + off-peak / Busy hours / Off-peak) are shown as information; they do not change the grade.
- Sort: severity, then touch ratio descending, then site name.

### v4 and v5 Dashboard_License

v4–v13 source: `DL_License_File.xlsx`, sheet `dlbbcapacity_dhk`. v14 source: `DL_License_File.xlsb`, sheet `Sheet1`. One row per node: `NodeId` (matches ERBS), `ENodeBFunctionId`, `dlBbCapacityNet` (the DL capacity license installed on the node, Mbps, one value for the window). v14 also has `ulBbCapacityNet`; the report does not use it.

| Version | Hour counts when usage / license | Site listed when | Sheet order |
|---|---|---|---|
| v4 | ≥ 90% | those hours ≥ 40% | Dashboard_License last |
| v5 | ≥ 80% | those hours ≥ 30% | Dashboard, then Dashboard_License, then the hourly charts |
| v6 | ≥ **85%** | those hours ≥ **40%** | License section is on Dashboard. Charts cover both lists |

v6 puts the license list on **Dashboard**, under the Peak-touch section. There is no separate Dashboard_License sheet.

v12 groups that license list into one category, the share of hours at ≥ 85% of `dlBbCapacityNet`: **≥ 70%**, **50% to under 70%**, **40% to under 50%**. Under 40% is not listed. This category is separate from the Peak-touch severity.

**2. HourlyChartOfIssueSites** has one chart for every site on either list. A site on both lists is drawn once. Peak-touch sites keep their order; license-only sites follow. Every issue chart has the same five lines: Max DL Capacity Usage, Peak, Below5%Peak, the license, and 85% of the license. There is no Avg peak line. Sheet **1. Site List** lists the same sites. Both tables link the site name to that chart. `3. Hourly KPI` holds the same sites.

### v9 Site List identity and GeoPlot

Physical database: `Physical_Site_Database_14Oct26.xlsx`. Top thanas: `Top Thana.xlsx`.

The physical site code is the first 7 characters of the ERBS. That code matches the Site Code column.

**1. Site List** copies Tech (4G or 4G+5G) and Site Type (GF, RTT, GFRT, and the other values in the file) from that row. A code with no row shows Tech **4G** and Site Type **Not found**. Column order (v19): No., ERBS, Tech, Site Type, Category, Severity, then the BH Flag 1H group, then Peak and the license columns, then the OSS 4H/1H blocks.

### v11 OSS cell rollup

OSS files: `OSS KPI Hourly Cel Level_11Oct`, `_12Oct`, and `_13Oct`, each V1.part1+part2 and V2.part1+part2. V1 is hours 0–11 and V2 is hours 12–23. Concatenate them. Do not commit the combined cell CSV.

Ericsson columns map to the same roles:

| Role | Ericsson column |
|---|---|
| Cell Name | `EUtranCell Id` |
| Site | first 7 characters of `EUtranCell Id` |
| Data Volume,GB | `Total Traffic (GB)` |
| DL User Throughput N/D | `Average User Throughput in DownLink (Mbps)_Num` / `_Denom` |
| Mean User | `Average Number of Users in a Cell` |
| Max User | `Maximum Number of Users in a Cell` |
| DL PRB Utilization N/D | `DL PRB Utilization (%)_Num` / `_Denom` |

Layer token is still Excel `MID(EUtranCell Id, 8, 3)`. Drop L900 when that token starts with `L09`. Keep every other cell and sum to one row per site-hour. Throughput = Num/Denom. PRB % = (Num/Denom) × 100.

v11–v13: the FEG window was 3–5 Oct and the OSS window was 11–13 Oct, so Hours matched stayed 0. v14 scores the 11–13 Oct FEG file, the same days as the OSS files, so Hours matched counts a busy hour that is also a Below5%Peak touch hour on that date. v15 writes that same window on every sheet: Dashboard, Site List, each issue chart, Hourly KPI, GeoPlot, BusyHour, Missing FEG Port, and Immediate actions. The 4H/1H KPI values are filled in both cases. Sheet `6. Missing FEG Port` compares OSS site codes with the first 7 characters of the ERBS. Sheet `7. Immediate actions` is Severe or High only, using the OSS 4H/1H volume and throughput. `URGENTBWINC_BusyHour` scans every capacity OSS site. On that sheet, skip the BH Flag block and the three Filtering blocks (1H only, 4H only, 7H/4H/1H). Keep Site, and the Data/User 7H, 4H, and 1H blocks. The chart sheet has the same 4H/1H table to the right of each chart.

**5. GeoPlot** plots one dot per physical site. A site is an issue when any of its ERBS is on the Dashboard. Severity is the worst Peak-touch grade at that site. Medium is drawn as Moderate, the name on the map legend. A site listed only by the license rule is drawn as Low. Maps: national, the four top thanas, then clusters until every geo-located issue site is on a map. Each map has a severity companion. Sheet order: Dashboard, 1. Site List, 2. HourlyChartOfIssueSites, 3. Hourly KPI, 5. GeoPlot.

A license of 0 or below, or an ERBS with no license row, is not scored. License sort: share of hours descending, then site name.

Dhaka sample sites **DHAPT35** and **DHAPT48** designed the shape. They are not a name filter. Do not fail an Ericsson run because those two names are absent.

Spelling **Chocked** on the BusyHour flag columns is intentional. Keep it.

---

## How to read uploaded files

Read each uploaded file from its name. Do not wait for an exact Dhaka filename, and do not ignore a file because the date, version, or part number is new. The name is the hint. Open the file and use the columns to confirm.

| Hint in the file name | What the file is |
|---|---|
| `OSS` | OSS KPI. Cell-level hourly radio KPI for the BusyHour sheet and the Site List 4H/1H block |
| `FEG` or `FEGE` | FEG port KPI. `RxMaxSpeed` is the choke trace. This is the detection source |
| Physical site, or a site database | Latitude, longitude, district, region, thana, Tech, Site Type |
| Thana | Top-thana list for GeoPlot |

Dates, `V1` / `V2`, and `part1` / `part2` in the name say how to combine that file with the others. They do not change what the file is.

`OSS` and `FEG` use the same combine order. For each date, merge that version’s part1 + part2, then merge the versions (V1 + V2) into the final file for that date. Keep OSS finals and FEG finals separate. If a later upload uses a different version or part count, follow the tokens in the names the same way.

Ericsson files in the repo: `FEG_KPI_Eric.xlsb` (11–13 Oct 2026), `DL_License_File.xlsb` (latest license; the older `DL_License_File.xlsx` remains), `Physical_Site_Database_14Oct26.xlsx`, `Top Thana.xlsx`, and the 11–13 Oct OSS KPI RARs. Skip the Comparison sheet until a previous-window FEG file is provided.

---

## 1. What a listed site is

A site is an issue only when **RxMaxSpeed** is stuck on a wall:

- After the night dip the line locks to one level, or
- The trace still moves during the day, but every rise stops at the same top and never goes above it.

A busy site whose peaks keep changing (tens to hundreds of Mbit/s apart) is not listed. Rejected shape examples: DHGULN2, DHDHN40, DHBDD32, DHDHN09.

This is a transmission link health review for improvement planning, not a blame list.

Report title: **Transmission Link Health Review**.  
GeoPlot title: **Transmission Link Health Monitoring**.

---

## 2. Counters

| Name | Formula | Role |
|---|---|---|
| RxMaxSpeed (Mbit/s) | `VS.FEGE.RxMaxSpeed(bit/s) / 1000 / 1000` | The only counter that decides choke |
| Tx Total BW (Mbit/s) | `VS.FEGE.TxTotalBW(kbit/s) / 1000` | Reference column only. Median across the window. Not used to list, grade, or set severity |

Do not revive the old “At / Above / Below Tx BW” 85–120% grade.

Skip a site when median Tx Total BW is ≤ 0.

Duplicate site + timestamp rows keep the last value. Dates such as `5-Oct` and hours such as `0:00` are parsed. Year is **2026**.

---

## 3. Window

- Score only the analysis window (Dhaka last run: **5, 6 and 7 Oct 2026**). Earlier days in the same file are ignored.
- Score **every hour**, busy **08:00–22:00** inclusive and off-peak / night.
- A site that is only flat at night can still be listed.
- Snapshot charts show the **same three days**.
- Night valley used for Night Rx: **02:00–06:00** inclusive. Night Rx is the median RxMaxSpeed in that valley. It is not a listing input. It was removed from Site List (v48). A site that stays on the cap through the night is still listed.

Listing thresholds stay fixed. Do not loosen or tighten them because the window is three days.

---

## 4. Detection (first shape that passes wins)

Skip a site with fewer than **24** hourly samples, or a max RxMaxSpeed below **20 Mbit/s**.

### Shape 1 — Crowded level

Tried first. DHAPT35 / DHAPT48 shape.

- Search the upper half of the trace (quantile **0.55** to max).
- Grid step **0.5 Mbit/s**. Pick the level that captures the most samples.
- Tolerance: **±4 Mbit/s**, or **±1.5% of p95** if that is wider.
- Recenter the level on the **median** of the in-band samples, then rebuild the band.
- The **99th percentile** may sit at most **40 Mbit/s** or **12% of the level** above the center. If it climbs past that, this shape fails (busy cluster with spikes, not a choke).

### Shape 2 — Hard ceiling

Tried only if shape 1 does not pass tightness and day rules.

- Same search, only over the **top 10%** of samples (quantile 0.90 to max).
- At most **1** hour may sit more than **8 Mbit/s** or **3% of the level** above the top.
- Catches traces that move during the day but clip at one value (DHSVRS5, GPSDR01, DHKKT87, DHKKTE2).

### Tightness (both shapes)

In-band standard deviation ≤ **2.5 Mbit/s** or **1% of the level**. Center ≥ **20 Mbit/s**.

---

## 5. Who is listed

A **day** is on-cap when at least **3 hours** that day (busy or off-peak, not necessarily consecutive) sit on the stuck level.

List the site if **either**:

1. At least **3 of the scored days** are on-cap, or
2. The **last day** of the window has at least **2 hours** on the cap.

A site can be listed for the last-day flag alone.

`Listed because`:

- `≥3 days + last day`
- `≥3 days`
- `Last day`

Last-day column text: `7-Oct-26 · 17h (busy 13, off-peak 4)` (date, hours, busy / off-peak split).

Longest consecutive run on the cap is extra information only. A gap over 90 minutes breaks the run. A day still counts when 3 hours sit on the level even if they are not adjacent.

### Severity

Hours on the cap as a share of **all hours in the window**. Not a comparison with Tx BW.

| Severity | Hours on cap |
|---|---|
| Severe | ≥ 50% |
| High | 25% to < 50% |
| Moderate | 10% to < 25% |
| Low | < 10% |

Low is still choked (usually a hard ceiling or a last-day flag).

### Where capped

| Cap when | Meaning |
|---|---|
| Busy + off-peak | Hours on cap in both 08:00–22:00 and the rest of the day |
| Busy hours | Only busy hours |
| Off-peak | Only off-peak / night |

### Sort

Site List: Severity (Severe first), then listed-because, then hours-on-cap % descending, then site name.

---

## 6. Charts

Sheet `2. HourlyChartOfIssueSites`. One chart per issue site for the three scored days.

- Keep the RxMaxSpeed series.
- **72-hour Peak** line: max RxMaxSpeed in the chart window.
- **Avg peak** line: mean of every chart-window hour that sits on the cap, reaches that day’s highest value, or sits at/above the stuck wall.
- Do not draw the old stuck-level dashed line (replaced at v40 / v47).
- Y-axis adapts to the trace so a ~33 Mbit/s site is not drawn on 0–100.
- Date is shown once per day (at 00:00). Hours 00:00–23:00. Chart title is the site name.
- Site name on Site List jumps to its chart.
- Hidden `_ChartData` holds the chart source.
- The chart header keeps only Tx Total BW, Stuck RxMaxSpeed, Hours on cap, and Days on cap.
- To the right of each hourly chart, with a gap after the chart, a table has columns Busy Hour, KPI, and Value. It has four blocks, same numbers as 1. Site List: Data 4H (orange), User 4H (gold), Data 1H (purple), User 1H (navy). Each block lists DL Throughput, DLPRBUtilization,%, Data Volume, Max User, and Mean User (the average of Mean User over that block’s busy hours). The Busy Hour cell is the window plus the hour type: `4H (Data Busy hour)`, `4H (User Busy Hour)`, `1H (Data Busy hour)`, `1H (User Busy Hour)`. Do not put a title above that table.
- Dashboard row 3 states how many OSS KPI sites are missing from the FEG file. Those sites are not choke-scored because the FEG port is not configured. The full list is sheet `6. Missing FEG Port`. The comparison is FEG eNodeB names versus the capacity OSS KPI, not the physical site database.
- Sheet `7. Immediate actions` lists Severe or High sites only. OR across Data 4H, User 4H, Data 1H, and User 1H. One tag per site, strongest match: Data Volume > 100 GB, then DL throughput < 5 Mbps, then 5 up to 8 Mbps, then 8 to 10 Mbps.
- Every eNodeB in the FEG file is choke-checked. A missing physical-database row does not drop the site. Issue sites with no database row are named at the top of GeoPlot and are not plotted.

### Double check

Site List **Double check = Yes** when `(Peak − Avg peak) / Peak ≥ 30%` (DHGULAP type). Detection listing does not change. UrgentTxBWInc leaves these sites off.

---

## 7. Technology and site type

From the physical site database the user provides (Dhaka used `Physical_Site_Database_24Sep26.xlsx`, NCCD used `Physical_Site_Database_11Oct26.xlsx`):

- **Tech** on Dashboard, Site List, and UrgentTxBWInc comes from the Tech column. Unmapped = **4G**.
- **Site Type** (GF, RTT, RTP, IBS, and so on) on Site List and URGENTBWINC_BusyHour comes from the Site Type column. Unmapped = **Not found**.

Do not mix physical radio Tech with a separate planning tech list. The live v84 rule uses the physical database Tech column.

GeoPlot coverage colours:

| Legend | Colour | Meaning |
|---|---|---|
| Coverage 4G | Cyan `#00E5FF` | Non-issue 4G |
| Coverage 5G | Yellow `#FFD600` | Non-issue 5G |
| Issue 4G | Red `#FF1744` | Issue, tech 4G |
| Issue 4G+5G | Orange `#FF6D00` | Issue, tech 4G+5G |

Severity map colours (v46): bright red Severe, deep red High, muted amber Moderate, calm green Low.

---

## 8. GeoPlot

Sheet `5. GeoPlot` is maps only. Coordinates live on hidden `_GeoData`. Google satellite background. No connecting lines. Freeze panes off.

Order:

1. National map (full physical-site layer, compact same-size markers, no site-name labels) plus a national issue-severity map beside it.
2. Top thana maps from `Top Thana.xlsx`, matched to the physical **Thana** column. Short title `Thana (4G: n 4G+5G: m)`. Each thana map has an issue-severity map beside it.
3. Dynamic issue clusters after the thana maps, urban-first, until every geo-located issue site is on a map. Each cluster map has an issue-severity map beside it.

Cluster rules:

- Cover every geo-located issue. Isolated sites get their own small map.
- Urban first: Dhaka Metro, Gazipur, Dhaka North / West / South, Narayanganj, Narsingdi / Manikganj / Munshiganj, Tangail, then anything else. For this repo, keep this urban-first idea using the districts that are actually in the Ericsson file.
- Inside each band, larger packs first.
- Cluster rank uses the **majority** of sites in the cluster.
- Pack size: urban (rank ≤ 2) target **18**, max **22**. Other areas target **16**, max **20**. Nearby leftovers may merge up to 22.
- Label: `Cluster NN — {majority district} / {majority region}`.
- Neighbouring issue sites in a padded view are drawn as coverage, not as extra issue dots. The title count is only that map’s own issues.
- Site names are labelled on thana and cluster maps, not on the national map.

Also write an interactive HTML (Leaflet + Google satellite) and embed JPEG copies on the GeoPlot sheet.

---

## 9. UrgentTxBWInc

Urgent bandwidth-increase list only. Not every issue site.

**Listed** when all of these are true:

- Severity is Severe or High (both cap shapes).
- Hours on cap (all 24h) ≥ **25%** (Severe is already ≥ 50%).
- Busy hours 08:00–22:00 sit on the cap at least **50%** of those hours.
- Double check is not Yes.

**Left off:** Moderate, Low, off-peak-only, busy-hour cap below 50%, and Double check = Yes.

Order: Severe first, then hours-on-cap % descending, then busy-hour cap % descending, then site name. Site name jumps to the hourly chart.

Thana summary: Name / Count pairs sized from the thana count. More thanas add pairs to the right. Highest counts fill column 1 first, then column 2, and so on.

---

## 10. OSS cell → site (capacity only)

Do **not** read `OSS KPI Hourly site  level.rar`.

### OSS file layout (one date = four RARs)

Any file whose name contains `OSS` is an OSS KPI file. The NCCD 8 Oct upload is the pattern, not the only names to accept. For each date there are four files. Example for 8 Oct:

- `OSS KPI Hourly Cel Level_08Oct_V1.part1.rar`
- `OSS KPI Hourly Cel Level_08Oct_V1.part2.rar`
- `OSS KPI Hourly Cel Level_08Oct_V2.part1.rar`
- `OSS KPI Hourly Cel Level_08Oct_V2.part2.rar`

Merge in this order, then do the same for every other date:

1. Merge V1 part1 + V1 part2 to get the V1 file.
2. Merge V2 part1 + V2 part2 to get the V2 file.
3. Merge V1 + V2 to get the final file for that date.
4. Repeat for the other dates, then combine those daily finals before the cell-to-site rollup.

### FEG file layout (same merge)

Any file whose name contains `FEG` or `FEGE` is the FEG port KPI. `RxMaxSpeed` from that file decides the choke list. It is not an OSS file.

If a date is split the same way as OSS, merge it the same way:

1. Merge V1 part1 + V1 part2 to get the V1 file.
2. Merge V2 part1 + V2 part2 to get the V2 file.
3. Merge V1 + V2 to get the final FEG file for that date.
4. Repeat for the other dates, then combine those daily finals before detection.

Do not mix a FEG final into the OSS cell rollup. Skip Comparison until the previous FEG file arrives.

### Cell rollup

Build site-hour OSS from those final daily cell files:

1. Layer token = Excel `MID(Cell Name, 8, 3)` (1-based start 8, length 3).
2. **Drop L900** when the token starts with `L09`.
3. **Keep** L2100 (`L21`), L2600 (`L26`), and everything else as L1800 capacity.
4. Roll remaining cells to one row per site-hour (`Date`, `Time`, `eNodeB Name`).

Aggregation:

- Drop cell-only columns: Cell TDD Indication, Cell Name, LocalCell Id, eNodeB Function Name, Integrity.
- **Sum:** Data Volume,GB; DL User Throughput_N; DL User Throughput_D; Mean User; Max User; DL PRB Utilization_N; DL PRB Utilization_D.
- **Average:** DL User Throughput,Mbps; DL PRB Utilization.
- Keep original KPI names. Do not rename them to “Sum of …”.
- Working file: `OSS KPI Hourly site  level_capacity.csv`. Do not commit the combined cell CSV.

BusyHour throughput and PRB are always recomputed from N/D, not from the file’s Mbps or % columns:

- DL User Throughput,Mbps = `DL User Throughput_N / DL User Throughput_D`
- DLPRBUtilization,% = `(DL PRB Utilization_N / DL PRB Utilization_D) × 100`

---

## 11. URGENTBWINC_BusyHour

Independent of UrgentTxBWInc. Scan **every** capacity-only OSS site. Do not limit this sheet to the UrgentTxBWInc list.

For each day in the analysis window, any hour of the day:

| Block | How the hours are chosen |
|---|---|
| DataBusyHour 7 / 4 / 1 | Sort Data Volume,GB largest to lowest. Take the top 7, top 4, or the single highest hour |
| UserBusyHour 7 / 4 / 1 | Sort Max User the same way |

A site is on this sheet when both DataBusyHour and UserBusyHour can be filled for any day in the window, or for the last day, on the 7, 4, or 1 set.

### Hours matched

A matched hour is a busy hour that is also FEGE on-cap. Each block has its own Hours matched, Hours matched %, Last day match, and Listed because.

| Block | Denominator | Last day |
|---|---|---|
| 7 BusyHour | 7 × number of scored days | choked hours / 7 |
| 4 BusyHour | 4 × number of scored days | choked hours / 4 |
| 1 BusyHour | 1 × number of scored days | choked hours / 1 |

`Listed because` for a block: `Any day + last day`, `Last day`, `Any day`, or `—`.

Each block also shows, over **that block’s own busy hours only**:

- DL Throughput (Mbps) from N/D
- DLPRBUtilization,% from N/D × 100
- Average Data Volume (GB)
- Average Max User
- Average Mean User (the mean of Mean User over those same busy hours)

Data Volume, Max User, Mean User, and DLPRBUtilization,% are **not** inputs to BH Flag.

### Yes flags (column order)

Identity: No., eNodeB Name, Tech, Severity, Site Type.

Then **BH Flag**:

| Flag | Yes when |
|---|---|
| 4 BH Flag | (Chocked %_Data OR Chocked %_User) **and** (Data Busy Hour_TP OR User Busy Hour_TP) in **Filtering(4H only)** |
| 1 BH Flag | Same pair in **Filtering(1H only)** |
| 4/1BH Flag | 4 BH Flag OR 1 BH Flag |

Chocked % is Yes when that section’s Hours matched % ≥ **30%**.  
TP is Yes when that section’s DL Throughput ≤ **7 Mbps**.

Then three filter groups, four columns each:

1. **Filtering(1H only):** Chocked %_Data, Chocked %_User, Data Busy Hour_TP, User Busy Hour_TP — 1h section only.
2. **Filtering(4H only):** same four names — 4h section only.
3. **Filtering(7H/4H/1H):** Chocked %_Data, Chocked %_User, Data Busy Hour, User Busy Hour. Yes when **any** of 7h OR 4h OR 1h meets the 30% or ≤ 7 Mbps test.

Filter the four columns of a group together with OR.

Then the six blocks, left to right:

1. DataBusyHour 7H (one column per day + Hours matched, %, Last day match, Listed because, DL Throughput 7h, DLPRBUtilization,%, Data Volume 7h, Max User 7h, Mean User 7h)
2. UserBusyHour 7H (same)
3. DataBusyHour 4H (same, with 4h names, including Mean User 4h)
4. UserBusyHour 4H (same)
5. DataBusyHour 1H (same, with 1h names, including Mean User 1h)
6. UserBusyHour 1H (same)

Header colours stay distinct per section (7H blue/teal, 4H orange/gold, 1H purple/navy, BH Flag amber, filter groups their own colours).

Sort: 4/1BH Flag, then 4 BH Flag, then 1 BH Flag, then the 1H choke/TP flags, then 4H, then 7H, then matched % descending, then severity, then site name.

---

## 12. Site List OSS block

On `1. Site List`, immediately after Site Type, copy four blocks from URGENTBWINC_BusyHour:

1. DataBusyHour (4 BusyHour): DL Throughput 4h, DLPRBUtilization,%, Data Volume 4h, Max User 4h, Mean User 4h
2. UserBusyHour (4 BusyHour): same
3. DataBusyHour (1 BusyHour): DL Throughput 1h, DLPRBUtilization,%, Data Volume 1h, Max User 1h, Mean User 1h
4. UserBusyHour (1 BusyHour): same

Blank means the issue site is not in the OSS BusyHour scan.

Site List columns before that block:

No., eNodeB Name, Tech, Severity, Cap shape, Tx Total BW (Mbit/s), Stuck RxMaxSpeed (Mbit/s), Max RxMaxSpeed (Mbit/s), Hours on cap (all 24h), Where capped, Days on cap (≥3h), Last day cap, Listed because, Longest flat run (h), Double check, Site Type.

Night Rx and sample-snap text are not on Site List.

---

## 13. Comparison

**Skipped until further notice.** The previous FEG file is not uploaded. Do not build the Comparison sheet, and do not invent a previous window.

When that file is provided, the sheet rules are:

Sheet **Comparison**. All sites. RxMaxSpeed only (no Tx Total BW columns).

- Busy-hour RxMaxSpeed **p90** for hours 08:00–22:00.
- Previous window vs current window. Dhaka last run: 2–4 Oct (old part1/part2 file) vs 5–7 Oct (current RAR).
- For this repo, build this sheet only when a previous-window file is actually provided. Do not invent the previous window.
- **BW Increased = Yes** when the new p90 is at least **40 Mbit/s** or **12%** above the old p90 (`need = max(40, 0.12 × old)`; Yes when `new − old ≥ need`). DHKKTE2 / GPSDR01 step-up shape.
- Sort BW Increased first, then delta descending, then site name.
- Columns include Tech, Site Type, Severity, old RxMaxSpeed, new RxMaxSpeed, delta, delta %, BW Increased.

---

## 14. Workbook

| Sheet | Contents |
|---|---|
| Dashboard | Opening view. Technology tiles. Cap window observation. Severity tiles. No Tx BW on the dashboard lists. No scored-on footnote (that lives on Method). |
| 1. Site List | Full issue list. Site name links to the chart. 4H/1H OSS block after Site Type. |
| 2. HourlyChartOfIssueSites | One chart per issue site |
| _ChartData | Hidden chart source |
| 3. Hourly KPI | Hourly numbers behind the charts |
| 4. Method | Detection, listing, severity, Double check, UrgentTxBWInc, BusyHour, Comparison |
| 5. GeoPlot | Maps only |
| _GeoData | Hidden coordinates |
| UrgentTxBWInc | Urgent upgrade list + thana summary |
| URGENTBWINC_BusyHour | Capacity-only OSS busy hours, flags, 7/4/1 blocks, including Mean User |
| 6. Missing FEG Port | OSS KPI sites with no FEG row. Action is to configure the FEG port. Compare FEG with OSS, not the physical site database |
| 7. Immediate actions | Severe or High only. OR across Data 4H, User 4H, Data 1H, and User 1H. One tag per site |
| Comparison | RxMaxSpeed ceiling step-up, when a previous window exists |

Report file name must include the version: `FEGE_Choked_Flat_Sites_Ericsson_vN.xlsx`.

### Missing FEG port

A site that is in the capacity OSS KPI and absent from the FEG file is not choke-scored. The FEG port is not configured, so FEG KPI is not collected. Dashboard row 3 states the count. Sheet `6. Missing FEG Port` lists every such site with action **Configure FEG port**. Do not use the physical site database for this compare.

### Immediate actions

Sheet `7. Immediate actions`. Severe or High only. OR across the four windows Data 4H, User 4H, Data 1H, and User 1H: if any one window matches, the site is listed.

Each site is listed once, under the strongest customer-experience tag:

1. High Priority: any of the four Data Volume values **> 100 GB**
2. High Priority: any of the four DL throughput values **< 5 Mbps**
3. Mid Priority: any of the four DL throughput values from **5 Mbps up to but not including 8 Mbps**
4. Low Priority: any of the four DL throughput values from **8 Mbps through 10 Mbps**

A louder window wins over a milder one. Volume above 100 GB wins over throughput. Throughput above 10 Mbps with volume at or under 100 GB is not listed.

Excel freeze panes are off. Formula and scoring notes stay on Method, not on Dashboard or Site List.

Default output name follows the presentation version. Changing maps or Dashboard does not change the detection thresholds in sections 4 and 5.

---

## 15. What not to do

- Do not change the detection thresholds in sections 4 and 5.
- Do not use Tx Total BW to list, grade, or set severity.
- Do not use the old site-level OSS RAR. Capacity-only cell rollup only.
- Do not limit URGENTBWINC_BusyHour to the UrgentTxBWInc list.
- Do not add Data Volume, Max User, Mean User, or DLPRBUtilization,% into BH Flag.
- Do not drop geo-located issue sites off the maps.
- Do not commit combined KPI CSVs or RAR working files.
- Do not add the Comparison sheet until the previous FEG file is provided.
