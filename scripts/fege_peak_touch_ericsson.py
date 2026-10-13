#!/usr/bin/env python3
"""Ericsson FEGE report, Peak-touch method (v3 onward).

Counter: only `Max DL Capacity Usage (Mbps)` from `FEG_KPI_Eric.xlsb`.

For every ERBS over the scored window (the 3 days in the file, every hour):

    Peak        = highest hourly value in the window (the red Peak line)
    Touch level = Peak x (1 - TOUCH_BAND)   (TOUCH_BAND = 5%)
    Touch hour  = an hour whose value is at or above the Touch level,
                  i.e. within 5% below the Peak
    Touch ratio = touch hours / total hours with data

Grade by touch ratio:

    >= 70%  Severe
    >= 50%  High
    >= 40%  Medium
    >= 20%  Low
    <  20%  not listed

Guards kept from RULES.md: a site needs at least 24 hourly samples and a
Peak of at least 20 Mbit/s to be scored.

Sheets written: Dashboard, 2. HourlyChartOfIssueSites (with hidden
_ChartData), 3. Hourly KPI.
"""

from __future__ import annotations

import argparse
import sys
import time as pytime
from collections import defaultdict
from datetime import time
from pathlib import Path

import numpy as np
import pandas as pd

sys.path.insert(0, str(Path(__file__).resolve().parent))
import fege_choke_check as fcc  # noqa: E402
import fege_choke_check_ericsson as eric  # noqa: E402

OUTPUT_DEFAULT = "FEGE_Choked_Flat_Sites_Ericsson_v3.xlsx"
TRACE_COL = eric.ERIC_DL_COL
TRACE_LABEL = "Max DL Capacity Usage"
TOUCH_BAND = 0.05
MIN_SAMPLES = 24
MIN_PEAK_MBPS = fcc.MIN_CENTER_MBPS

SEV_SEVERE = "Severe"
SEV_HIGH = "High"
SEV_MEDIUM = "Medium"
SEV_LOW = "Low"
SEVERITY_BANDS = (
    (SEV_SEVERE, 70.0),
    (SEV_HIGH, 50.0),
    (SEV_MEDIUM, 40.0),
    (SEV_LOW, 20.0),
)
SEVERITY_ORDER = {name: i for i, (name, _pct) in enumerate(SEVERITY_BANDS)}
SEVERITY_STYLE = {
    SEV_SEVERE: "severe",
    SEV_HIGH: "high",
    SEV_MEDIUM: "moderate",
    SEV_LOW: "low",
}
SEVERITY_NOTES = (
    (SEV_SEVERE, "Severe: at least 70% of the hours touch the Peak (within 5% below it)."),
    (SEV_HIGH, "High: 50% to under 70% of the hours touch the Peak."),
    (SEV_MEDIUM, "Medium: 40% to under 50% of the hours touch the Peak."),
    (SEV_LOW, "Low: 20% to under 40% of the hours touch the Peak. Under 20% is not listed."),
)
CHART_DATA_PER_ROW = eric.CHART_DATA_PER_ROW
CHART_DATA_BLOCK_ROWS = eric.CHART_DATA_BLOCK_ROWS


def severity(touch_pct: float) -> str | None:
    for name, floor in SEVERITY_BANDS:
        if touch_pct >= floor:
            return name
    return None


def analyse_touch(raw: pd.DataFrame) -> tuple[pd.DataFrame, list[dict]]:
    """Peak-touch scoring on the trace column. Returns (hourly frame, records)."""
    work = pd.DataFrame(
        {
            "Date": fcc.parse_kpi_date(raw["Date"]),
            "Hour": raw["Hour"].astype(int),
            "eNodeB Name": raw[eric.ERIC_SITE_COL].astype(str).str.strip(),
            "Rx": pd.to_numeric(raw[TRACE_COL], errors="coerce"),
        }
    )
    work["ts"] = work["Date"] + pd.to_timedelta(work["Hour"], unit="h")
    work = work.sort_values(["eNodeB Name", "ts"]).drop_duplicates(
        ["eNodeB Name", "ts"], keep="last"
    )
    end = work["Date"].max().normalize()
    start = max(work["Date"].min().normalize(), end - pd.Timedelta(days=fcc.SNAP_DAYS - 1))
    work = work[(work["Date"] >= start) & (work["Date"] <= end)].reset_index(drop=True)
    last_day = end.date()
    last_day_label = end.strftime("%-d-%b-%y")

    touch = np.zeros(len(work), dtype=bool)
    peak_col = np.full(len(work), np.nan)
    level_col = np.full(len(work), np.nan)
    records: list[dict] = []

    for site, idx in work.groupby("eNodeB Name", sort=False).groups.items():
        s = work.loc[idx]
        rx = s["Rx"].to_numpy(dtype=float)
        valid = np.isfinite(rx)
        if valid.sum() < MIN_SAMPLES:
            continue
        peak = float(np.nanmax(rx))
        if peak < MIN_PEAK_MBPS:
            continue
        level = peak * (1.0 - TOUCH_BAND)
        band = valid & (rx >= level)
        hours_n = int(valid.sum())
        hours_on = int(band.sum())
        touch_pct = 100.0 * hours_on / hours_n
        sev = severity(touch_pct)
        hours = s["Hour"].to_numpy()
        busy = (hours >= fcc.BUSY_START) & (hours <= fcc.BUSY_END)
        pos = work.index.get_indexer(s.index)
        peak_col[pos] = peak
        level_col[pos] = level
        touch[pos] = band
        if sev is None:
            continue
        days = s["Date"].dt.date.to_numpy()
        day_on: dict = defaultdict(int)
        for day, flag in zip(days, band):
            if flag:
                day_on[day] += 1
        busy_on = int(band[busy].sum())
        off_on = int(band[~busy].sum())
        busy_n = int((busy & valid).sum())
        records.append(
            {
                "site": site,
                "severity": sev,
                "peak": peak,
                "level": level,
                "hours_on": hours_on,
                "hours_n": hours_n,
                "hours_pct": touch_pct,
                "busy_on": busy_on,
                "busy_n": busy_n,
                "busy_pct": 100.0 * busy_on / busy_n if busy_n else 0.0,
                "off_on": off_on,
                "cap_when": fcc.cap_window(busy_on, off_on),
                "days_total": int(s["Date"].dt.normalize().nunique()),
                "days_touch": sum(1 for v in day_on.values() if v >= fcc.STUCK_DAY_HOURS),
                "last_on": int(day_on.get(last_day, 0)),
                "last_day_label": last_day_label,
                "center": level,
            }
        )

    work["OnCap"] = touch
    work["Peak"] = peak_col
    work["TouchLevel"] = level_col
    records.sort(key=lambda r: (SEVERITY_ORDER[r["severity"]], -r["hours_pct"], r["site"]))
    return work, records


def _page(ws, title: str, *, fit_width: bool = True) -> None:
    fcc._page(ws, title, fit_width=fit_width)
    ws.set_header(f"&L&8&K1F4E79{title}&R&8{fcc.REPORT_TITLE}  |  {eric.ZONE}")
    ws.set_footer(f"&L&8Touch = {TRACE_LABEL} within 5% below the Peak&R&8Page &P of &N")


def write_dashboard(book, styles, source_name, period_txt, n_sites, n_scored, records) -> None:
    ws = book.add_worksheet("Dashboard")
    _page(ws, fcc.REPORT_TITLE)
    ws.set_tab_color("#C65911")
    for i, w in enumerate([16, 16, 16, 18, 18, 14, 14, 16, 18, 16]):
        ws.set_column(i, i, w)
    tf = eric._tile_formats(book)

    counts = defaultdict(int)
    for rec in records:
        counts[rec["severity"]] += 1

    ws.set_row(0, 28)
    ws.merge_range("A1:J1", fcc.REPORT_TITLE, styles["title"])
    ws.set_row(1, 18)
    ws.merge_range(
        "A2:J2",
        f"{eric.ZONE} transmission    ·    {period_txt}    ·    hourly    ·    {source_name}",
        styles["subtitle"],
    )
    ws.set_row(2, 64)
    ws.merge_range(
        "A3:J3",
        f"Method: only {TRACE_COL} is read. Peak = the highest hourly value of the window "
        f"(red Peak line). An hour touches the Peak when its value is within "
        f"{TOUCH_BAND * 100:.0f}% below the Peak (at or above {100 - TOUCH_BAND * 100:.0f}% of "
        "Peak, the orange Touch level line). Touch ratio = touch hours / total hours with data. "
        "Severe ≥ 70%, High ≥ 50%, Medium ≥ 40%, Low ≥ 20%; under 20% is not listed. "
        f"A site is scored when it has at least {MIN_SAMPLES} hourly samples and a Peak of at "
        f"least {MIN_PEAK_MBPS:.0f} Mbit/s. "
        "This version carries Dashboard, 2. HourlyChartOfIssueSites and 3. Hourly KPI only.",
        styles["note"],
    )

    row = eric._write_tiles(
        ws,
        3,
        [
            (0, "Sites checked", n_sites, False),
            (2, "Listed (touch ≥ 20%)", len(records), True),
            (4, "Severe (≥ 70%)", counts[SEV_SEVERE], True),
            (
                6,
                "High / Medium / Low",
                f"{counts[SEV_HIGH]}  /  {counts[SEV_MEDIUM]}  /  {counts[SEV_LOW]}",
                False,
            ),
        ],
        tf,
    )
    row += 1

    sevs = [name for name, _pct in SEVERITY_BANDS]
    matrix = {sev: {when: 0 for when in fcc.CAP_WHEN_COLS} for sev in sevs}
    for rec in records:
        if rec["cap_when"] in matrix[rec["severity"]]:
            matrix[rec["severity"]][rec["cap_when"]] += 1
    ws.set_row(row, 22)
    ws.merge_range(row, 0, row, 9, "Touch Window Observation", tf["group"])
    row += 1
    ws.set_row(row, 28)
    ws.write(row, 0, "Severity", styles["header"])
    for offset, text in enumerate(fcc.CAP_WHEN_COLS):
        ws.write(row, 1 + offset, text, styles["header"])
    ws.write(row, 4, "Total", styles["header"])
    row += 1
    for sev in sevs:
        ws.set_row(row, 18)
        ws.write(row, 0, sev, styles[SEVERITY_STYLE[sev]])
        for offset, when in enumerate(fcc.CAP_WHEN_COLS):
            value = matrix[sev][when]
            if value:
                ws.write_number(row, 1 + offset, value, styles["int"])
            else:
                ws.write_blank(row, 1 + offset, None, styles["center"])
        ws.write_number(row, 4, counts[sev], styles["int"])
        row += 1
    ws.set_row(row, 18)
    ws.write(row, 0, "Grand Total", tf["total_label"])
    for offset, when in enumerate(fcc.CAP_WHEN_COLS):
        ws.write_number(row, 1 + offset, sum(matrix[sev][when] for sev in sevs), tf["total"])
    ws.write_number(row, 4, len(records), tf["total"])
    row += 2

    headers = [
        "No.",
        "ERBS",
        "Severity",
        "Peak (Mbit/s)",
        "Touch level (Mbit/s)",
        "Touch hours",
        "Total hours",
        "Touch ratio (%)",
        "Busy-hour touch (%)",
        "Where touched",
    ]
    indexed = list(enumerate(records))
    for name, note in SEVERITY_NOTES:
        group = [(i, rec) for i, rec in indexed if rec["severity"] == name]
        if not group:
            continue
        ws.set_row(row, 22)
        ws.merge_range(
            row,
            0,
            row,
            9,
            f"{name}    ·    {len(group)} site{'s' if len(group) != 1 else ''}",
            styles["section"],
        )
        row += 1
        ws.set_row(row, 24)
        ws.merge_range(row, 0, row, 9, note, styles["body"])
        row += 1
        ws.set_row(row, 30)
        for col, text in enumerate(headers):
            ws.write(row, col, text, styles["header"])
        row += 1
        for n_in_group, (i, rec) in enumerate(group):
            zebra = n_in_group % 2 == 1
            ws.set_row(row, 18)
            nfmt = styles["num_z"] if zebra else styles["num"]
            ifmt = styles["int_z"] if zebra else styles["int"]
            cfmt = styles["center_z"] if zebra else styles["center"]
            pfmt = styles["pct_z"] if zebra else styles["pct"]
            link = styles["link_z"] if zebra else styles["link"]
            ws.write_number(row, 0, i + 1, ifmt)
            ws.write_url(
                row,
                1,
                f"internal:'{fcc.SNAP_SHEET}'!A{i * fcc.BLOCK_ROWS + 1}",
                link,
                string=rec["site"],
            )
            ws.write_string(row, 2, rec["severity"], styles[SEVERITY_STYLE[rec["severity"]]])
            ws.write_number(row, 3, rec["peak"], nfmt)
            ws.write_number(row, 4, rec["level"], nfmt)
            ws.write_number(row, 5, rec["hours_on"], ifmt)
            ws.write_number(row, 6, rec["hours_n"], ifmt)
            ws.write_number(row, 7, rec["hours_pct"], pfmt)
            ws.write_number(row, 8, rec["busy_pct"], pfmt)
            ws.write_string(row, 9, rec["cap_when"], cfmt)
            row += 1
        row += 1
    ws.set_zoom(110)


def write_snapshots(book, styles, work, records, chart_start, chart_end) -> None:
    ws = book.add_worksheet(fcc.SNAP_SHEET)
    _page(ws, "Hourly chart of issue sites", fit_width=False)
    ws.set_zoom(100)
    ws.set_print_scale(100)
    ws.set_tab_color(fcc.BLUE)
    data = book.add_worksheet("_ChartData")
    data.hide()
    for col, width in enumerate([18, 16, 18, 16, 18, 16, 18, 16]):
        ws.set_column(col, col, width)
    block = fcc.BLOCK_ROWS
    cat_fmt = book.add_format({"font_name": "Calibri", "font_size": 8, "align": "center"})
    chart_days = list(pd.date_range(chart_start, chart_end, freq="D"))
    by_site = {site: g for site, g in work.groupby("eNodeB Name", sort=False)}

    print(f"Writing {len(records)} snapshot charts...")
    for i, rec in enumerate(records):
        if i and i % 200 == 0:
            print(f"  snapshots {i}/{len(records)}", flush=True)
        top = i * block
        site = rec["site"]
        s = by_site[site]
        s = s[(s["Date"] >= chart_start) & (s["Date"] <= chart_end)].sort_values("ts")
        c0 = (i % CHART_DATA_PER_ROW) * 5
        r0 = (i // CHART_DATA_PER_ROW) * CHART_DATA_BLOCK_ROWS
        for offset, name in enumerate(("Date", "Time", "Rx", "Peak", "TouchLevel")):
            data.write(r0, c0 + offset, name)
        rx_by_key = {
            (d.normalize(), int(h)): float(v) for d, h, v in zip(s["Date"], s["Hour"], s["Rx"])
        }
        peak = float(rec["peak"])
        level = float(rec["level"])
        dates: list[str] = []
        times: list[str] = []
        plotted: list[float] = [peak, level]
        r = r0
        for day in chart_days:
            date_label = f"{int(day.day)}/{fcc.MONTHS[int(day.month) - 1]}"
            for hour in range(24):
                r += 1
                dates.append(date_label if hour == 0 else "")
                times.append(f"{hour:02d}:00")
                if hour == 0:
                    data.write_string(r, c0, date_label, cat_fmt)
                else:
                    data.write_blank(r, c0, None)
                data.write_string(r, c0 + 1, f"{hour:02d}:00", cat_fmt)
                value = rx_by_key.get((day.normalize(), hour))
                if value is None or not np.isfinite(value):
                    data.write_blank(r, c0 + 2, None)
                else:
                    data.write_number(r, c0 + 2, value)
                    plotted.append(value)
                data.write_number(r, c0 + 3, peak)
                data.write_number(r, c0 + 4, level)
        n = r
        first = r0 + 1

        ws.set_row(top, 26)
        ws.merge_range(
            top, 0, top, 7, f"{i + 1:02d}    {site}    ·    {rec['severity']}", styles["banner"]
        )
        labels = [
            (0, "Peak (Mbit/s)", f"{peak:.2f}"),
            (2, f"Touch level ({100 - TOUCH_BAND * 100:.0f}% of Peak, Mbit/s)", f"{level:.2f}"),
            (
                4,
                "Touch hours (all 24h)",
                f"{rec['hours_on']} / {rec['hours_n']}  ({rec['hours_pct']:.1f}%)",
            ),
            (
                6,
                "Busy-hour touch (08–22)",
                f"{rec['busy_on']} / {rec['busy_n']}  ({rec['busy_pct']:.1f}%)",
            ),
        ]
        ws.set_row(top + 1, 16)
        ws.set_row(top + 2, 20)
        for col, label, _value in labels:
            ws.merge_range(top + 1, col, top + 1, col + 1, label, styles["kpi_label"])
        for col, _label, value in labels:
            ws.merge_range(top + 2, col, top + 2, col + 1, value, styles["kpi_value"])
        ws.set_row(top + 3, 8)

        chart = book.add_chart({"type": "line"})
        chart.show_blanks_as("gap")
        chart.set_title(
            {"name": site, "name_font": {"name": "Calibri", "size": 16, "color": "#595959"}}
        )
        series = (
            (f"{TRACE_LABEL} (Mbps)", 2, {"color": "#5B9BD5", "width": 2.25}),
            (f"Peak ({peak:.2f} Mbit/s)", 3, {"color": "#C00000", "width": 1.5, "dash_type": "dash"}),
            (
                f"Touch level ({level:.2f} Mbit/s)",
                4,
                {"color": "#ED7D31", "width": 1.5, "dash_type": "dash"},
            ),
        )
        for name, offset, line in series:
            chart.add_series(
                {
                    "name": name,
                    "categories": ["_ChartData", first, c0, n, c0 + 1],
                    "categories_data": [dates, times],
                    "values": ["_ChartData", first, c0 + offset, n, c0 + offset],
                    "line": line,
                    "marker": {"type": "none"},
                }
            )
        chart.set_x_axis(
            {
                "num_font": {"name": "Calibri", "size": 8, "color": "#595959", "rotation": -90},
                "label_position": "nextTo",
                "label_align": "center",
                "position_axis": "on_tick",
                "major_tick_mark": "none",
                "minor_tick_mark": "none",
                "major_gridlines": {"visible": False},
                "minor_gridlines": {"visible": False},
                "line": {"color": "#B0B0B0"},
            }
        )
        y_max, y_unit = fcc._chart_y_scale(plotted)
        chart.set_y_axis(
            {
                "num_font": {"name": "Calibri", "size": 9, "color": "#595959"},
                "min": 0,
                "max": y_max,
                "major_unit": y_unit,
                "major_gridlines": {"visible": True, "line": {"color": "#D9D9D9"}},
                "line": {"none": True},
            }
        )
        chart.set_legend({"position": "top", "font": {"name": "Calibri", "size": 9}})
        chart.set_chartarea({"border": {"color": "#D0D0D0"}, "fill": {"color": "white"}})
        chart.set_plotarea(
            {
                "border": {"none": True},
                "fill": {"color": "white"},
                "layout": {"x": 0.06, "y": 0.16, "width": 0.91, "height": 0.58},
            }
        )
        chart.set_size({"width": 1200, "height": 410})
        ws.insert_chart(top + 4, 0, chart, {"x_offset": 6, "y_offset": 6, "object_position": 2})
        for rr in range(top + 4, top + block):
            ws.set_row(rr, 16)

    if records:
        ws.set_h_pagebreaks([i * block for i in range(1, len(records))])


def write_hourly(book, styles, work, records) -> None:
    ws = book.add_worksheet("3. Hourly KPI")
    _page(ws, "Hourly FEGE KPI — listed sites")
    ws.set_tab_color("#5B9BD5")
    headers = [
        "Date",
        "Time",
        eric.ERIC_SITE_COL,
        TRACE_COL,
        "Touches Peak",
        "Peak (Mbit/s)",
        f"Touch level ({100 - TOUCH_BAND * 100:.0f}% of Peak)",
    ]
    for i, w in enumerate([14, 10, 16, 28, 14, 16, 24]):
        ws.set_column(i, i, w)
    ws.set_row(0, 30)
    for col, text in enumerate(headers):
        ws.write(0, col, text, styles["header"])

    sites = [r["site"] for r in records]
    hourly = work.loc[work["eNodeB Name"].isin(sites)].sort_values(["eNodeB Name", "ts"])
    parity = {site: i % 2 for i, site in enumerate(sites)}
    rows = zip(
        hourly["eNodeB Name"].tolist(),
        hourly["Date"].tolist(),
        hourly["Hour"].tolist(),
        hourly["Rx"].tolist(),
        hourly["OnCap"].tolist(),
        hourly["Peak"].tolist(),
        hourly["TouchLevel"].tolist(),
    )
    excel_row = 0
    for site, date, hour, rx, on, peak, level in rows:
        excel_row += 1
        z = "_z" if parity[site] else ""
        ws.write_datetime(excel_row, 0, date.to_pydatetime(), styles["date" + z])
        ws.write_datetime(excel_row, 1, time(int(hour), 0), styles["time" + z])
        ws.write_string(excel_row, 2, site, styles["text" + z])
        if rx is None or not np.isfinite(rx):
            ws.write_blank(excel_row, 3, None, styles["num" + z])
        else:
            ws.write_number(excel_row, 3, float(rx), styles["num" + z])
        ws.write_string(excel_row, 4, "Yes" if bool(on) else "No", styles["center" + z])
        ws.write_number(excel_row, 5, float(peak), styles["num" + z])
        ws.write_number(excel_row, 6, float(level), styles["num" + z])
    ws.autofilter(0, 0, max(excel_row, 1), len(headers) - 1)
    ws.repeat_rows(0, 0)
    ws.set_zoom(110)


def write_workbook(path: Path, source_name: str, work: pd.DataFrame, records: list[dict]) -> None:
    import xlsxwriter

    book = xlsxwriter.Workbook(str(path))
    styles = fcc._styles(book)
    period_start = work["Date"].min()
    period_end = work["Date"].max()
    n_sites = int(work["eNodeB Name"].nunique())
    n_scored = int(work.loc[work["Peak"].notna(), "eNodeB Name"].nunique())
    period_txt = f"{period_start.strftime('%d %b %Y')} – {period_end.strftime('%d %b %Y')}"
    chart_end = period_end.normalize()
    chart_start = chart_end - pd.Timedelta(days=fcc.SNAP_DAYS - 1)
    write_dashboard(book, styles, source_name, period_txt, n_sites, n_scored, records)
    write_snapshots(book, styles, work, records, chart_start, chart_end)
    print("Writing hourly KPI...", flush=True)
    write_hourly(book, styles, work, records)
    book.close()


def main() -> None:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("source", nargs="?", default=eric.SOURCE_DEFAULT)
    parser.add_argument("-o", "--output", default=OUTPUT_DEFAULT)
    args = parser.parse_args()
    t0 = pytime.perf_counter()

    source = Path(args.source)
    raw = eric.load_ericsson(source)
    required = {"Date", "Hour", eric.ERIC_SITE_COL, TRACE_COL}
    missing = required - set(raw.columns)
    if missing:
        raise SystemExit(f"Missing columns: {', '.join(sorted(missing))}")
    work, records = analyse_touch(raw)
    print(
        f"Window: {work['Date'].min().date()} – {work['Date'].max().date()}  ·  "
        f"trace {TRACE_COL}  ·  touch band {TOUCH_BAND * 100:.0f}% below Peak"
    )
    print(f"Analysis {pytime.perf_counter() - t0:.1f}s")
    output = Path(args.output)
    write_workbook(output, source.name, work, records)

    by_sev = defaultdict(int)
    by_when = defaultdict(int)
    for rec in records:
        by_sev[rec["severity"]] += 1
        by_when[rec["cap_when"]] += 1
    print(f"Sites checked: {work['eNodeB Name'].nunique()}")
    print(f"Sites scored (≥{MIN_SAMPLES} samples, Peak ≥ {MIN_PEAK_MBPS:.0f}): "
          f"{work.loc[work['Peak'].notna(), 'eNodeB Name'].nunique()}")
    print(f"Listed (touch ≥ 20%): {len(records)}")
    for name, floor in SEVERITY_BANDS:
        print(f"  {name} (≥{floor:.0f}%): {by_sev[name]}")
    print("Where touched:")
    for name in fcc.CAP_WHEN_COLS:
        print(f"  {name}: {by_when[name]}")
    print(f"Wrote {output} in {pytime.perf_counter() - t0:.1f}s")


if __name__ == "__main__":
    main()
