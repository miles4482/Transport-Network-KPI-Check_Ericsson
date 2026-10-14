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

OUTPUT_DEFAULT = "FEGE_Choked_Flat_Sites_Ericsson_v7.xlsx"
TOUCH_LEVEL_NAME = "touch levelbelow5%"
LICENSE_FILE = "DL_License_File.xlsx"
LICENSE_SHEET = "Dashboard_License"
LICENSE_NODE_COL = "NodeId"
LICENSE_CAP_COL = "dlBbCapacityNet"
# An hour counts when Max DL Capacity Usage / dlBbCapacityNet is at or above
# this share. A site is listed when such hours are at least LICENSE_HOURS_PCT
# of its hours.
LICENSE_USE = 0.85
LICENSE_HOURS_PCT = 40.0
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


def issue_charts(touch_records: list[dict], license_records: list[dict]) -> list[dict]:
    """One chart per site listed by the Peak-touch rule or the license rule.

    Peak-touch sites keep their order. License-only sites follow, highest
    share of hours first. A site on both lists is drawn once.
    """
    license_by = {rec["site"]: rec for rec in license_records}
    charts: list[dict] = []
    seen: set[str] = set()
    for rec in touch_records:
        charts.append({"site": rec["site"], "touch": rec, "license": license_by.get(rec["site"])})
        seen.add(rec["site"])
    for rec in license_records:
        if rec["site"] not in seen:
            charts.append({"site": rec["site"], "touch": None, "license": rec})
    return charts


def write_dashboard(
    book,
    styles,
    source_name,
    period_txt,
    n_sites,
    n_scored,
    records,
    *,
    link_charts: bool = True,
    chart_index: dict[str, int] | None = None,
    license_summary: dict | None = None,
) -> None:
    ws = book.add_worksheet("Dashboard")
    _page(ws, fcc.REPORT_TITLE)
    ws.set_tab_color("#C65911")
    for i, w in enumerate([8, 16, 18, 16, 16, 16, 20, 18, 18, 18]):
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
        f"Peak, the orange {TOUCH_LEVEL_NAME} line). Touch ratio = touch hours / total hours with data. "
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
        f"{TOUCH_LEVEL_NAME} (Mbit/s)",
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
            chart_i = (chart_index or {}).get(rec["site"], i)
            ws.write_number(row, 0, chart_i + 1, ifmt)
            if link_charts:
                ws.write_url(
                    row,
                    1,
                    f"internal:'{fcc.SNAP_SHEET}'!A{chart_i * fcc.BLOCK_ROWS + 1}",
                    link,
                    string=rec["site"],
                )
            else:
                ws.write_string(row, 1, rec["site"], styles["text_z"] if zebra else styles["text"])
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
    if license_summary is not None:
        _write_license_on_dashboard(book, ws, styles, row + 1, license_summary, chart_index or {})
    ws.set_footer("&L&8Peak touch, or Max DL Capacity Usage / dlBbCapacityNet&R&8Page &P of &N")
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
            (2, f"{TOUCH_LEVEL_NAME} (Mbit/s)", f"{level:.2f}"),
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
                f"{TOUCH_LEVEL_NAME} ({level:.2f} Mbit/s)",
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
        f"{TOUCH_LEVEL_NAME} (Mbit/s)",
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


def analyse_license(raw: pd.DataFrame, lic: pd.DataFrame) -> dict:
    """Sites whose Max DL usage sits at or above the license threshold.

    dlBbCapacityNet is the capacity license on the node (NodeId = ERBS), one
    value for the whole window. An hour counts when
    Max DL Capacity Usage (Mbps) / dlBbCapacityNet >= LICENSE_USE. A site is
    listed when those hours are at least LICENSE_HOURS_PCT of its hours.
    """
    cap = pd.to_numeric(lic[LICENSE_CAP_COL], errors="coerce")
    by_node = dict(zip(lic[LICENSE_NODE_COL].astype(str).str.strip(), cap))
    df = pd.DataFrame(
        {
            "site": raw[eric.ERIC_SITE_COL].astype(str).str.strip(),
            "date": fcc.parse_kpi_date(raw["Date"]),
            "hour": raw["Hour"].astype(int),
            "rx": pd.to_numeric(raw[TRACE_COL], errors="coerce"),
        }
    )
    df = df.dropna(subset=["rx"]).sort_values(["site", "date", "hour"])
    df = df.drop_duplicates(["site", "date", "hour"], keep="last")
    df["lic"] = df["site"].map(by_node)
    scored = df["lic"].notna() & (df["lic"] > 0)
    df["hit"] = False
    df.loc[scored, "hit"] = df.loc[scored, "rx"] >= LICENSE_USE * df.loc[scored, "lic"]
    df["busy"] = (df["hour"] >= fcc.BUSY_START) & (df["hour"] <= fcc.BUSY_END)

    records: list[dict] = []
    for site, g in df[scored].groupby("site", sort=False):
        hours_n = int(len(g))
        hours_on = int(g["hit"].sum())
        share = 100.0 * hours_on / hours_n
        if share < LICENSE_HOURS_PCT:
            continue
        license_mbps = float(g["lic"].iloc[0])
        peak = float(g["rx"].max())
        busy = g["busy"].to_numpy()
        hit = g["hit"].to_numpy()
        records.append(
            {
                "site": site,
                "license": license_mbps,
                "hours_on": hours_on,
                "hours_n": hours_n,
                "share": share,
                "peak": peak,
                "peak_ratio": 100.0 * peak / license_mbps,
                "median_ratio": 100.0 * float(np.median(g["rx"] / license_mbps)),
                "cap_when": fcc.cap_window(int(hit[busy].sum()), int(hit[~busy].sum())),
            }
        )
    records.sort(key=lambda r: (-r["share"], r["site"]))
    listed = [rec["site"] for rec in records]
    hourly = df.loc[df["site"].isin(listed)].copy()
    hourly["ratio"] = hourly["rx"] / hourly["lic"]
    return {
        "records": records,
        "hourly": hourly,
        "feg_sites": int(df["site"].nunique()),
        "license_nodes": int(lic[LICENSE_NODE_COL].nunique()),
        "matched": int(df.loc[df["lic"].notna(), "site"].nunique()),
        "no_license": int(df.loc[df["lic"].isna(), "site"].nunique()),
        "bad_license": int(df.loc[df["lic"].notna() & (df["lic"] <= 0), "site"].nunique()),
        "scored": int(df.loc[scored, "site"].nunique()),
    }


def _write_license_on_dashboard(book, ws, styles, row: int, summary: dict, chart_index: dict[str, int]) -> None:
    """License list on the Dashboard sheet. Site names jump to the shared charts."""
    records = summary["records"]
    use_pct = f"{LICENSE_USE * 100:.0f}"
    group = book.add_format(
        {
            "font_name": "Calibri",
            "font_size": 14,
            "bold": True,
            "font_color": fcc.NAVY,
            "valign": "vcenter",
        }
    )
    ws.set_row(row, 22)
    ws.merge_range(row, 0, row, 9, "License usage", group)
    row += 1
    ws.set_row(row, 48)
    ws.merge_range(
        row,
        0,
        row,
        9,
        f"dlBbCapacityNet is the DL capacity license installed on the node "
        f"({LICENSE_NODE_COL} = ERBS). An hour counts when {TRACE_COL} / dlBbCapacityNet "
        f"≥ {use_pct}%. A site is listed when those hours are at least "
        f"{LICENSE_HOURS_PCT:.0f}% of its hours. "
        f"{summary['no_license']:,} ERBS have no row in {LICENSE_FILE} and "
        f"{summary['bad_license']:,} have a license of 0 or below, so they are not scored.",
        styles["note"],
    )
    row += 1
    eric._write_tiles(
        ws,
        row,
        [
            (0, "ERBS in the FEG file", summary["feg_sites"], False),
            (2, "Matched to a license", summary["matched"], False),
            (4, "License scored (> 0)", summary["scored"], False),
            (
                6,
                f"Listed (≥{LICENSE_HOURS_PCT:.0f}% of hours at ≥{use_pct}%)",
                len(records),
                True,
            ),
        ],
        eric._tile_formats(book),
    )
    row += 4
    headers = [
        "No.",
        "ERBS",
        "License (Mbps)",
        f"Hours ≥{use_pct}%",
        "Total hours",
        "Share of hours (%)",
        "Peak Max DL (Mbps)",
        "Peak / license (%)",
        "Median usage (%)",
        f"Where ≥{use_pct}%",
    ]
    ws.set_row(row, 22)
    ws.merge_range(
        row,
        0,
        row,
        9,
        f"Sites at ≥{use_pct}% of license for ≥{LICENSE_HOURS_PCT:.0f}% of hours"
        f"    ·    {len(records)}",
        styles["section"],
    )
    row += 1
    ws.set_row(row, 32)
    for col, text in enumerate(headers):
        ws.write(row, col, text, styles["header"])
    row += 1
    for n, rec in enumerate(records):
        zebra = n % 2 == 1
        chart_i = chart_index[rec["site"]]
        ws.set_row(row, 18)
        nfmt = styles["num_z"] if zebra else styles["num"]
        ifmt = styles["int_z"] if zebra else styles["int"]
        cfmt = styles["center_z"] if zebra else styles["center"]
        pfmt = styles["pct_z"] if zebra else styles["pct"]
        link = styles["link_z"] if zebra else styles["link"]
        ws.write_number(row, 0, chart_i + 1, ifmt)
        ws.write_url(
            row,
            1,
            f"internal:'{fcc.SNAP_SHEET}'!A{chart_i * fcc.BLOCK_ROWS + 1}",
            link,
            string=rec["site"],
        )
        ws.write_number(row, 2, rec["license"], nfmt)
        ws.write_number(row, 3, rec["hours_on"], ifmt)
        ws.write_number(row, 4, rec["hours_n"], ifmt)
        ws.write_number(row, 5, rec["share"], pfmt)
        ws.write_number(row, 6, rec["peak"], nfmt)
        ws.write_number(row, 7, rec["peak_ratio"], pfmt)
        ws.write_number(row, 8, rec["median_ratio"], pfmt)
        ws.write_string(row, 9, rec["cap_when"], cfmt)
        row += 1


def write_license_dashboard(book, styles, source_name, period_txt, summary: dict) -> None:
    """Dashboard_License: sites using the installed DL license at the set threshold."""
    ws = book.add_worksheet(LICENSE_SHEET)
    _page(ws, "License usage")
    ws.set_footer("&L&8License usage = Max DL Capacity Usage / dlBbCapacityNet&R&8Page &P of &N")
    ws.set_tab_color("#375623")
    for i, w in enumerate([8, 16, 18, 16, 14, 16, 18, 16, 16, 18]):
        ws.set_column(i, i, w)
    tf = eric._tile_formats(book)
    records = summary["records"]

    ws.set_row(0, 28)
    ws.merge_range("A1:J1", "License Usage Review", styles["title"])
    ws.set_row(1, 18)
    ws.merge_range(
        "A2:J2",
        f"{eric.ZONE}    ·    {period_txt}    ·    {source_name}    ·    {LICENSE_FILE}",
        styles["subtitle"],
    )
    ws.set_row(2, 64)
    ws.merge_range(
        "A3:J3",
        f"dlBbCapacityNet is the DL capacity license installed on the node "
        f"({LICENSE_NODE_COL} = ERBS), one value for the window. "
        f"An hour counts when {TRACE_COL} / dlBbCapacityNet "
        f"≥ {LICENSE_USE * 100:.0f}%. A site is listed when those hours are at least "
        f"{LICENSE_HOURS_PCT:.0f}% of its hours. "
        f"{summary['no_license']:,} ERBS have no row in {LICENSE_FILE} and "
        f"{summary['bad_license']:,} have a license of 0 or below, so they are not scored.",
        styles["note"],
    )
    eric._write_tiles(
        ws,
        3,
        [
            (0, "ERBS in the FEG file", summary["feg_sites"], False),
            (2, "Matched to a license", summary["matched"], False),
            (4, "License scored (> 0)", summary["scored"], False),
            (
                6,
                f"Listed (≥{LICENSE_HOURS_PCT:.0f}% of hours at ≥{LICENSE_USE * 100:.0f}%)",
                len(records),
                True,
            ),
        ],
        tf,
    )

    use_pct = f"{LICENSE_USE * 100:.0f}"
    headers = [
        "No.",
        "ERBS",
        "License (Mbps)",
        f"Hours ≥{use_pct}%",
        "Total hours",
        "Share of hours (%)",
        "Peak Max DL (Mbps)",
        "Peak / license (%)",
        "Median usage (%)",
        f"Where ≥{use_pct}%",
    ]
    row = 7
    ws.set_row(row, 22)
    ws.merge_range(
        row, 0, row, 9,
        f"Sites at ≥{LICENSE_USE * 100:.0f}% of license for ≥{LICENSE_HOURS_PCT:.0f}% of hours"
        f"    ·    {len(records)}",
        styles["section"],
    )
    row += 1
    ws.set_row(row, 32)
    for col, text in enumerate(headers):
        ws.write(row, col, text, styles["header"])
    row += 1
    for i, rec in enumerate(records):
        zebra = i % 2 == 1
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
        ws.write_number(row, 2, rec["license"], nfmt)
        ws.write_number(row, 3, rec["hours_on"], ifmt)
        ws.write_number(row, 4, rec["hours_n"], ifmt)
        ws.write_number(row, 5, rec["share"], pfmt)
        ws.write_number(row, 6, rec["peak"], nfmt)
        ws.write_number(row, 7, rec["peak_ratio"], pfmt)
        ws.write_number(row, 8, rec["median_ratio"], pfmt)
        ws.write_string(row, 9, rec["cap_when"], cfmt)
        row += 1
    ws.autofilter(8, 0, max(row - 1, 8), len(headers) - 1)
    ws.set_zoom(110)


def write_license_snapshots(book, styles, summary: dict, chart_start, chart_end) -> None:
    """One chart per Dashboard_License site: usage, the license, and the 80% line."""
    records = summary["records"]
    hourly = summary["hourly"]
    ws = book.add_worksheet(fcc.SNAP_SHEET)
    _page(ws, "Hourly chart of license issue sites", fit_width=False)
    ws.set_footer("&L&8License usage = Max DL Capacity Usage / dlBbCapacityNet&R&8Page &P of &N")
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
    by_site = {site: g for site, g in hourly.groupby("site", sort=False)}
    use_pct = LICENSE_USE * 100

    print(f"Writing {len(records)} license charts...")
    for i, rec in enumerate(records):
        if i and i % 100 == 0:
            print(f"  charts {i}/{len(records)}", flush=True)
        top = i * block
        site = rec["site"]
        s = by_site[site]
        s = s[(s["date"] >= chart_start) & (s["date"] <= chart_end)].sort_values(["date", "hour"])
        c0 = (i % CHART_DATA_PER_ROW) * 5
        r0 = (i // CHART_DATA_PER_ROW) * CHART_DATA_BLOCK_ROWS
        for offset, name in enumerate(("Date", "Time", "Rx", "License", "Threshold")):
            data.write(r0, c0 + offset, name)
        rx_by_key = {
            (d.normalize(), int(h)): float(v) for d, h, v in zip(s["date"], s["hour"], s["rx"])
        }
        license_mbps = float(rec["license"])
        threshold = license_mbps * LICENSE_USE
        dates: list[str] = []
        times: list[str] = []
        plotted: list[float] = [license_mbps, threshold, float(rec["peak"])]
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
                data.write_number(r, c0 + 3, license_mbps)
                data.write_number(r, c0 + 4, threshold)
        n = r
        first = r0 + 1

        ws.set_row(top, 26)
        ws.merge_range(
            top,
            0,
            top,
            7,
            f"{i + 1:02d}    {site}    ·    {rec['share']:.1f}% of hours at ≥{use_pct:.0f}% of license",
            styles["banner"],
        )
        labels = [
            (0, "License dlBbCapacityNet (Mbps)", f"{license_mbps:.2f}"),
            (2, f"{use_pct:.0f}% of license (Mbps)", f"{threshold:.2f}"),
            (4, f"Hours ≥{use_pct:.0f}%", f"{rec['hours_on']} / {rec['hours_n']}  ({rec['share']:.1f}%)"),
            (6, "Peak Max DL (Mbps)", f"{rec['peak']:.2f}  ({rec['peak_ratio']:.1f}%)"),
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
            (
                f"License ({license_mbps:.2f} Mbps)",
                3,
                {"color": "#C00000", "width": 1.5, "dash_type": "dash"},
            ),
            (
                f"{use_pct:.0f}% of license ({threshold:.2f} Mbps)",
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


def write_license_hourly(book, styles, summary: dict) -> None:
    """Hourly rows for the Dashboard_License sites, in the same order as the charts."""
    ws = book.add_worksheet("3. Hourly KPI")
    _page(ws, "Hourly license usage — listed sites")
    ws.set_footer("&L&8License usage = Max DL Capacity Usage / dlBbCapacityNet&R&8Page &P of &N")
    ws.set_tab_color("#5B9BD5")
    use_pct = f"{LICENSE_USE * 100:.0f}"
    headers = [
        "Date",
        "Time",
        eric.ERIC_SITE_COL,
        TRACE_COL,
        "dlBbCapacityNet",
        "Usage / license (%)",
        f"≥{use_pct}% of license",
    ]
    for i, w in enumerate([14, 10, 16, 28, 18, 20, 18]):
        ws.set_column(i, i, w)
    ws.set_row(0, 30)
    for col, text in enumerate(headers):
        ws.write(0, col, text, styles["header"])

    order = {rec["site"]: i for i, rec in enumerate(summary["records"])}
    hourly = summary["hourly"].copy()
    hourly["_order"] = hourly["site"].map(order)
    hourly = hourly.sort_values(["_order", "date", "hour"])
    parity = {rec["site"]: i % 2 for i, rec in enumerate(summary["records"])}
    excel_row = 0
    for site, date, hour, rx, lic, ratio, hit in zip(
        hourly["site"],
        hourly["date"],
        hourly["hour"],
        hourly["rx"],
        hourly["lic"],
        hourly["ratio"],
        hourly["hit"],
    ):
        excel_row += 1
        z = "_z" if parity[site] else ""
        ws.write_datetime(excel_row, 0, date.to_pydatetime(), styles["date" + z])
        ws.write_datetime(excel_row, 1, time(int(hour), 0), styles["time" + z])
        ws.write_string(excel_row, 2, site, styles["text" + z])
        ws.write_number(excel_row, 3, float(rx), styles["num" + z])
        ws.write_number(excel_row, 4, float(lic), styles["num" + z])
        ws.write_number(excel_row, 5, 100.0 * float(ratio), styles["pct" + z])
        ws.write_string(excel_row, 6, "Yes" if bool(hit) else "No", styles["center" + z])
    ws.autofilter(0, 0, max(excel_row, 1), len(headers) - 1)
    ws.repeat_rows(0, 0)
    ws.set_zoom(110)


def write_issue_snapshots(book, styles, work, charts: list[dict], chart_start, chart_end) -> None:
    """One chart per issue site from either list. Lines match why the site is listed."""
    ws = book.add_worksheet(fcc.SNAP_SHEET)
    _page(ws, "Hourly chart of issue sites", fit_width=False)
    ws.set_footer("&L&8Peak touch, or Max DL Capacity Usage / dlBbCapacityNet&R&8Page &P of &N")
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
    use_pct = LICENSE_USE * 100
    cols = 7

    print(f"Writing {len(charts)} issue charts...")
    for i, item in enumerate(charts):
        if i and i % 100 == 0:
            print(f"  charts {i}/{len(charts)}", flush=True)
        top = i * block
        site = item["site"]
        touch = item["touch"]
        lic = item["license"]
        s = by_site.get(site)
        if s is None:
            s = work.iloc[0:0]
        else:
            s = s[(s["Date"] >= chart_start) & (s["Date"] <= chart_end)].sort_values("ts")
        c0 = (i % CHART_DATA_PER_ROW) * cols
        r0 = (i // CHART_DATA_PER_ROW) * CHART_DATA_BLOCK_ROWS
        for offset, name in enumerate(("Date", "Time", "Rx", "Peak", "Touch", "License", "Threshold")):
            data.write(r0, c0 + offset, name)
        rx_by_key = {
            (d.normalize(), int(h)): float(v) for d, h, v in zip(s["Date"], s["Hour"], s["Rx"])
        }
        peak = float(touch["peak"]) if touch else float("nan")
        touch_level = float(touch["level"]) if touch else float("nan")
        license_mbps = float(lic["license"]) if lic else float("nan")
        threshold = license_mbps * LICENSE_USE if np.isfinite(license_mbps) else float("nan")
        dates: list[str] = []
        times: list[str] = []
        plotted = [v for v in (peak, touch_level, license_mbps, threshold) if np.isfinite(v)]
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
                if np.isfinite(peak):
                    data.write_number(r, c0 + 3, peak)
                if np.isfinite(touch_level):
                    data.write_number(r, c0 + 4, touch_level)
                if np.isfinite(license_mbps):
                    data.write_number(r, c0 + 5, license_mbps)
                if np.isfinite(threshold):
                    data.write_number(r, c0 + 6, threshold)
        n = r
        first = r0 + 1

        tags = []
        if touch:
            tags.append(f"Peak touch {touch['severity']}")
        if lic:
            tags.append(f"License {lic['share']:.1f}% of hours ≥{use_pct:.0f}%")
        ws.set_row(top, 26)
        ws.merge_range(top, 0, top, 7, f"{i + 1:02d}    {site}    ·    {'    ·    '.join(tags)}", styles["banner"])
        if touch and lic:
            labels = [
                (0, "Peak (Mbit/s)", f"{peak:.2f}"),
                (2, "Touch ratio", f"{touch['hours_on']} / {touch['hours_n']}  ({touch['hours_pct']:.1f}%)"),
                (4, "License (Mbps)", f"{license_mbps:.2f}"),
                (6, f"Hours ≥{use_pct:.0f}% of license", f"{lic['hours_on']} / {lic['hours_n']}  ({lic['share']:.1f}%)"),
            ]
        elif touch:
            labels = [
                (0, "Peak (Mbit/s)", f"{peak:.2f}"),
                (2, f"{TOUCH_LEVEL_NAME} (Mbit/s)", f"{touch_level:.2f}"),
                (4, "Touch hours (all 24h)", f"{touch['hours_on']} / {touch['hours_n']}  ({touch['hours_pct']:.1f}%)"),
                (6, "Busy-hour touch (08–22)", f"{touch['busy_on']} / {touch['busy_n']}  ({touch['busy_pct']:.1f}%)"),
            ]
        else:
            labels = [
                (0, "License dlBbCapacityNet (Mbps)", f"{license_mbps:.2f}"),
                (2, f"{use_pct:.0f}% of license (Mbps)", f"{threshold:.2f}"),
                (4, f"Hours ≥{use_pct:.0f}%", f"{lic['hours_on']} / {lic['hours_n']}  ({lic['share']:.1f}%)"),
                (6, "Peak Max DL (Mbps)", f"{lic['peak']:.2f}  ({lic['peak_ratio']:.1f}%)"),
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
        chart.set_title({"name": site, "name_font": {"name": "Calibri", "size": 16, "color": "#595959"}})
        series = [(f"{TRACE_LABEL} (Mbps)", 2, {"color": "#5B9BD5", "width": 2.25})]
        if touch:
            series.append((f"Peak ({peak:.2f} Mbit/s)", 3, {"color": "#C00000", "width": 1.5, "dash_type": "dash"}))
            series.append((f"{TOUCH_LEVEL_NAME} ({touch_level:.2f} Mbit/s)", 4, {"color": "#ED7D31", "width": 1.5, "dash_type": "dash"}))
        if lic:
            series.append((f"License ({license_mbps:.2f} Mbps)", 5, {"color": "#548235", "width": 1.5, "dash_type": "dash"}))
            series.append((f"{use_pct:.0f}% of license ({threshold:.2f} Mbps)", 6, {"color": "#833C0C", "width": 1.5, "dash_type": "dash"}))
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
        y_max, y_unit = fcc._chart_y_scale(plotted or [50.0])
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
    if charts:
        ws.set_h_pagebreaks([i * block for i in range(1, len(charts))])


def write_issue_hourly(book, styles, work, charts: list[dict]) -> None:
    """Hourly rows for every charted site, in chart order."""
    ws = book.add_worksheet("3. Hourly KPI")
    _page(ws, "Hourly KPI — issue sites")
    ws.set_footer("&L&8Peak touch, or Max DL Capacity Usage / dlBbCapacityNet&R&8Page &P of &N")
    ws.set_tab_color("#5B9BD5")
    use_pct = f"{LICENSE_USE * 100:.0f}"
    headers = [
        "Date",
        "Time",
        eric.ERIC_SITE_COL,
        TRACE_COL,
        "Touches Peak",
        "Peak (Mbit/s)",
        f"{TOUCH_LEVEL_NAME} (Mbit/s)",
        "dlBbCapacityNet",
        "Usage / license (%)",
        f"≥{use_pct}% of license",
    ]
    for i, w in enumerate([14, 10, 16, 28, 14, 16, 18, 18, 20, 18]):
        ws.set_column(i, i, w)
    ws.set_row(0, 30)
    for col, text in enumerate(headers):
        ws.write(0, col, text, styles["header"])

    order = {item["site"]: i for i, item in enumerate(charts)}
    license_of = {item["site"]: item["license"]["license"] for item in charts if item["license"]}
    hourly = work.loc[work["eNodeB Name"].isin(order)].copy()
    hourly["_order"] = hourly["eNodeB Name"].map(order)
    hourly = hourly.sort_values(["_order", "ts"])
    parity = {item["site"]: i % 2 for i, item in enumerate(charts)}
    excel_row = 0
    for site, date, hour, rx, on_cap, peak, level in zip(
        hourly["eNodeB Name"],
        hourly["Date"],
        hourly["Hour"],
        hourly["Rx"],
        hourly["OnCap"],
        hourly["Peak"],
        hourly["TouchLevel"],
    ):
        excel_row += 1
        z = "_z" if parity[site] else ""
        ws.write_datetime(excel_row, 0, date.to_pydatetime(), styles["date" + z])
        ws.write_datetime(excel_row, 1, time(int(hour), 0), styles["time" + z])
        ws.write_string(excel_row, 2, site, styles["text" + z])
        if rx is None or not np.isfinite(rx):
            ws.write_blank(excel_row, 3, None, styles["num" + z])
        else:
            ws.write_number(excel_row, 3, float(rx), styles["num" + z])
        if peak is None or not np.isfinite(peak):
            ws.write_blank(excel_row, 4, None, styles["center" + z])
            ws.write_blank(excel_row, 5, None, styles["num" + z])
            ws.write_blank(excel_row, 6, None, styles["num" + z])
        else:
            ws.write_string(excel_row, 4, "Yes" if bool(on_cap) else "No", styles["center" + z])
            ws.write_number(excel_row, 5, float(peak), styles["num" + z])
            ws.write_number(excel_row, 6, float(level), styles["num" + z])
        lic = license_of.get(site)
        if lic is None or not np.isfinite(rx):
            ws.write_blank(excel_row, 7, None, styles["num" + z])
            ws.write_blank(excel_row, 8, None, styles["pct" + z])
            ws.write_blank(excel_row, 9, None, styles["center" + z])
        else:
            ws.write_number(excel_row, 7, float(lic), styles["num" + z])
            ws.write_number(excel_row, 8, 100.0 * float(rx) / float(lic), styles["pct" + z])
            ws.write_string(
                excel_row, 9, "Yes" if float(rx) >= LICENSE_USE * float(lic) else "No", styles["center" + z]
            )
    ws.autofilter(0, 0, max(excel_row, 1), len(headers) - 1)
    ws.repeat_rows(0, 0)
    ws.set_zoom(110)


def write_workbook(
    path: Path,
    source_name: str,
    work: pd.DataFrame,
    records: list[dict],
    license_summary: dict | None = None,
) -> None:
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
    license_records = license_summary["records"] if license_summary else []
    charts = issue_charts(records, license_records)
    chart_index = {item["site"]: i for i, item in enumerate(charts)}
    # License results sit on the Dashboard, under the Peak-touch section.
    # One chart per site from either list, and both tables link to it.
    write_dashboard(
        book, styles, source_name, period_txt, n_sites, n_scored, records,
        link_charts=True,
        chart_index=chart_index,
        license_summary=license_summary,
    )
    if license_summary is not None:
        write_issue_snapshots(book, styles, work, charts, chart_start, chart_end)
        print("Writing hourly KPI...", flush=True)
        write_issue_hourly(book, styles, work, charts)
    else:
        write_snapshots(book, styles, work, records, chart_start, chart_end)
        print("Writing hourly KPI...", flush=True)
        write_hourly(book, styles, work, records)
    book.close()


def main() -> None:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("source", nargs="?", default=eric.SOURCE_DEFAULT)
    parser.add_argument("-o", "--output", default=OUTPUT_DEFAULT)
    parser.add_argument("--license", default=LICENSE_FILE, help="DL capacity license file")
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
    license_summary = None
    license_path = Path(args.license)
    if license_path.exists():
        lic = pd.read_excel(license_path)
        missing_cols = {LICENSE_NODE_COL, LICENSE_CAP_COL} - set(lic.columns)
        if missing_cols:
            raise SystemExit(f"{license_path.name} is missing columns: {', '.join(sorted(missing_cols))}")
        license_summary = analyse_license(raw, lic)
        print(
            f"License: {license_summary['matched']:,} ERBS matched, "
            f"{len(license_summary['records']):,} listed "
            f"(≥{LICENSE_USE * 100:.0f}% of license for ≥{LICENSE_HOURS_PCT:.0f}% of hours)"
        )
    else:
        print(f"License file not found ({license_path}); license section not written")
    output = Path(args.output)
    write_workbook(output, source.name, work, records, license_summary)

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
