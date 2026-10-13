#!/usr/bin/env python3
"""Ericsson FEGE choke report: Dashboard, hourly charts, and hourly KPI only.

The Ericsson transport export (`FEG_KPI_Eric.xlsb`) has these columns:

    Date | Hour | ERBS | Max DL Capacity Usage (Mbps)
         | Max UL Capacity Usage(Mbps) | Total Capcity,Mbps

They map onto the Huawei roles in RULES.md like this:

    RxMaxSpeed (choke trace)   <- Max DL Capacity Usage (Mbps)
    Tx Total BW (reference)    <- Max UL Capacity Usage (Mbps)
    eNodeB Name                <- ERBS
    Date / Time                <- Date (Excel serial) / Hour (0-23)

Detection, listing, severity, chart lines, and the Dashboard layout are the
v84 rules imported from fege_choke_check.py. Only the three requested sheets
are written: Dashboard, 2. HourlyChartOfIssueSites (with its hidden
_ChartData), and 3. Hourly KPI. OSS, physical site database, and thana files
are not uploaded for Ericsson, so Tech, Site Type, GeoPlot, BusyHour, Site
List and the OSS side tables are not built.
"""

from __future__ import annotations

import argparse
import pickle
import sys
import time as pytime
from collections import defaultdict
from datetime import time
from pathlib import Path

import numpy as np
import pandas as pd

sys.path.insert(0, str(Path(__file__).resolve().parent))
import fege_choke_check as fcc  # noqa: E402

ZONE = "Ericsson"
SOURCE_DEFAULT = "FEG_KPI_Eric.xlsb"
OUTPUT_DEFAULT = "FEGE_Choked_Flat_Sites_Ericsson_v1.xlsx"
PARSE_CACHE = Path("/tmp/eric_feg_parsed.pkl")

ERIC_SITE_COL = "ERBS"
ERIC_DL_COL = "Max DL Capacity Usage (Mbps)"
ERIC_UL_COL = "Max UL Capacity Usage(Mbps)"
ERIC_TOTAL_COL = "Total Capcity,Mbps"
# Derived columns in the units analyse() expects (bit/s and kbit/s), so the
# Huawei detection code runs unchanged on the Ericsson Mbps values.
RX_BITS_COL = "Max DL Capacity Usage (bit/s)"
TX_KBITS_COL = "Max UL Capacity Usage (kbit/s)"
RX_LABEL = "Max DL Capacity Usage"
TX_LABEL = "Max UL Capacity Usage"
# _ChartData holds 5 columns per chart. Excel stops at 16,384 columns, so the
# sites are laid out in row blocks of CHART_DATA_PER_ROW sites each.
CHART_DATA_PER_ROW = 2000
CHART_DATA_BLOCK_ROWS = 1 + 24 * fcc.SNAP_DAYS + 1


def load_ericsson(path: Path) -> pd.DataFrame:
    """Read the Ericsson export once; cache the parsed frame next to its mtime."""
    stamp = (str(path.resolve()), path.stat().st_mtime_ns)
    if PARSE_CACHE.exists():
        try:
            cached = pickle.loads(PARSE_CACHE.read_bytes())
            if cached.get("stamp") == stamp:
                print(f"Loaded parsed source from {PARSE_CACHE}")
                return cached["frame"]
        except Exception:
            pass
    t0 = pytime.perf_counter()
    raw = fcc.load_kpi(path)
    print(f"Read {path.name}: {len(raw):,} rows in {pytime.perf_counter() - t0:.1f}s")
    PARSE_CACHE.write_bytes(pickle.dumps({"stamp": stamp, "frame": raw}))
    return raw


def to_rule_frame(raw: pd.DataFrame) -> pd.DataFrame:
    """Rename and rescale the Ericsson columns into the frame analyse() reads."""
    required = {"Date", "Hour", ERIC_SITE_COL, ERIC_DL_COL, ERIC_UL_COL}
    missing = required - set(raw.columns)
    if missing:
        raise SystemExit(f"Missing columns: {', '.join(sorted(missing))}")
    df = pd.DataFrame(
        {
            "Date": raw["Date"],
            "Time": raw["Hour"].astype(int).astype(str) + ":00",
            "eNodeB Name": raw[ERIC_SITE_COL].astype(str).str.strip(),
            RX_BITS_COL: pd.to_numeric(raw[ERIC_DL_COL], errors="coerce") * 1_000_000.0,
            TX_KBITS_COL: pd.to_numeric(raw[ERIC_UL_COL], errors="coerce") * 1_000.0,
            ERIC_DL_COL: pd.to_numeric(raw[ERIC_DL_COL], errors="coerce"),
            ERIC_UL_COL: pd.to_numeric(raw[ERIC_UL_COL], errors="coerce"),
        }
    )
    if ERIC_TOTAL_COL in raw.columns:
        df[ERIC_TOTAL_COL] = pd.to_numeric(raw[ERIC_TOTAL_COL], errors="coerce")
    else:
        df[ERIC_TOTAL_COL] = df[ERIC_DL_COL] + df[ERIC_UL_COL]
    return df


def set_window_from_file(df: pd.DataFrame) -> tuple[pd.Timestamp, pd.Timestamp]:
    """Score the dates inside the file: the last SNAP_DAYS calendar days present."""
    dates = fcc.parse_kpi_date(df["Date"]).dt.normalize()
    end = dates.max()
    start = max(dates.min(), end - pd.Timedelta(days=fcc.SNAP_DAYS - 1))
    fcc.ANALYSIS_START = start
    fcc.ANALYSIS_END = end
    fcc.ANALYSIS_DAYS = int((end - start).days) + 1
    return start, end


def _page(ws, title: str, *, fit_width: bool = True) -> None:
    fcc._page(ws, title, fit_width=fit_width)
    ws.set_header(f"&L&8&K1F4E79{title}&R&8{fcc.REPORT_TITLE}  |  {ZONE}")
    ws.set_footer(f"&L&8Choked = {RX_LABEL} stuck flat&R&8Page &P of &N")


def _tile_formats(book):
    def fmt(size, color, bg):
        return book.add_format(
            {
                "font_name": "Calibri",
                "font_size": size,
                "bold": True,
                "font_color": color,
                "align": "center",
                "valign": "vcenter",
                "bg_color": bg,
            }
        )

    return {
        "big": fmt(22, fcc.NAVY, fcc.PALE),
        "big_alert": fmt(22, fcc.ORANGE_FONT, fcc.ORANGE),
        "tile": fmt(9, fcc.GREY, fcc.PALE),
        "tile_alert": fmt(9, fcc.ORANGE_FONT, fcc.ORANGE),
        "group": book.add_format(
            {
                "font_name": "Calibri",
                "font_size": 14,
                "bold": True,
                "font_color": fcc.NAVY,
                "valign": "vcenter",
            }
        ),
        "total": book.add_format(
            {
                "font_name": "Calibri",
                "font_size": 10,
                "bold": True,
                "align": "center",
                "valign": "vcenter",
                "bg_color": fcc.PALE,
                "border": 1,
                "border_color": fcc.LINE,
            }
        ),
        "total_label": book.add_format(
            {
                "font_name": "Calibri",
                "font_size": 10,
                "bold": True,
                "align": "left",
                "valign": "vcenter",
                "bg_color": fcc.PALE,
                "border": 1,
                "border_color": fcc.LINE,
            }
        ),
    }


def _write_tiles(ws, row: int, tiles, tf) -> int:
    ws.set_row(row, 18)
    ws.set_row(row + 1, 32)
    for col, label, value, alert in tiles:
        ws.merge_range(row, col, row, col + 1, label, tf["tile_alert"] if alert else tf["tile"])
        ws.merge_range(
            row + 1, col, row + 1, col + 1, value, tf["big_alert"] if alert else tf["big"]
        )
    return row + 2


def write_dashboard(book, styles, source_name, period_txt, n_sites, records) -> None:
    ws = book.add_worksheet("Dashboard")
    _page(ws, fcc.REPORT_TITLE)
    ws.set_tab_color("#C65911")
    for i, w in enumerate([16, 16, 16, 20, 16, 20, 36, 20, 16, 16]):
        ws.set_column(i, i, w)
    tf = _tile_formats(book)

    counts = defaultdict(int)
    by_shape = defaultdict(int)
    by_listed = defaultdict(int)
    last_n = 0
    for rec in records:
        counts[rec["severity"]] += 1
        by_shape[rec["cap_type"]] += 1
        by_listed[rec["listed_by"]] += 1
        last_n += int(rec["last_flag"])

    ws.set_row(0, 28)
    ws.merge_range("A1:J1", fcc.REPORT_TITLE, styles["title"])
    ws.set_row(1, 18)
    ws.merge_range(
        "A2:J2",
        f"{ZONE} transmission    ·    {period_txt}    ·    hourly    ·    {source_name}",
        styles["subtitle"],
    )
    ws.set_row(2, 48)
    ws.merge_range(
        "A3:J3",
        f"Ericsson counter mapping: RxMaxSpeed = {RX_LABEL} (Mbps), the only counter that "
        f"decides choke. Tx Total BW reference = {TX_LABEL} (Mbps), median over the window, "
        "not used to list or grade. Site = ERBS. "
        "This version carries Dashboard, 2. HourlyChartOfIssueSites and 3. Hourly KPI only. "
        "The physical site database, OSS KPI and top-thana files are not uploaded, so Tech, "
        "Site Type, GeoPlot, BusyHour and the OSS 4H/1H blocks are not built.",
        styles["note"],
    )

    row = _write_tiles(
        ws,
        3,
        [
            (0, "Sites checked", n_sites, False),
            (2, "Issue sites", len(records), True),
            (4, "Severe choke", counts[fcc.SEV_SEVERE], True),
            (
                6,
                "High / Moderate / Low",
                f"{counts[fcc.SEV_HIGH]}  /  {counts[fcc.SEV_MODERATE]}  /  {counts[fcc.SEV_LOW]}",
                False,
            ),
        ],
        tf,
    )

    ws.set_row(row, 22)
    ws.merge_range(row, 0, row, 9, "Cap shape and listing", tf["group"])
    row += 1
    row = _write_tiles(
        ws,
        row,
        [
            (0, "Crowded level", by_shape[fcc.CAP_CROWDED], False),
            (2, "Hard ceiling", by_shape[fcc.CAP_CEILING], False),
            (4, "≥3 days on cap", by_listed["≥3 days + last day"] + by_listed["≥3 days"], False),
            (
                6,
                f"Last-day cap flag ({records[0]['last_day_label'] if records else 'n/a'})",
                last_n,
                True,
            ),
        ],
        tf,
    )
    row += 1

    matrix = fcc.cap_when_matrix(records, {})["total"]
    sevs = (fcc.SEV_SEVERE, fcc.SEV_HIGH, fcc.SEV_MODERATE, fcc.SEV_LOW)
    ws.set_row(row, 22)
    ws.merge_range(row, 0, row, 9, "Cap Window Observation", tf["group"])
    row += 1
    ws.set_row(row, 28)
    ws.write(row, 0, "Severity", styles["header"])
    for offset, text in enumerate(fcc.CAP_WHEN_COLS):
        ws.write(row, 1 + offset, text, styles["header"])
    ws.write(row, 4, "Total", styles["header"])
    row += 1
    for sev in sevs:
        ws.set_row(row, 18)
        ws.write(row, 0, sev, fcc._severity_format(styles, sev, False))
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
        "Cap shape",
        f"Stuck {RX_LABEL} (Mbit/s)",
        "Hours on cap (%)",
        "Days on cap (≥3h)",
        "Last day cap",
        "Listed because",
    ]
    indexed = list(enumerate(records))
    for name, note in fcc.SEVERITY_NOTES:
        group = [(i, rec) for i, rec in indexed if rec["severity"] == name]
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
        ws.set_row(row, 28)
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
            cfmt = styles["center_z"] if zebra else styles["center"]
            link = styles["link_z"] if zebra else styles["link"]
            excel_anchor = i * fcc.BLOCK_ROWS + 1
            ws.write_number(row, 0, i + 1, styles["int_z"] if zebra else styles["int"])
            ws.write_url(
                row,
                1,
                f"internal:'{fcc.SNAP_SHEET}'!A{excel_anchor}",
                link,
                string=rec["site"],
            )
            ws.write_string(row, 2, rec["cap_type"], cfmt)
            ws.write_number(row, 3, rec["center"], nfmt)
            ws.write_number(row, 4, rec["hours_pct"], styles["pct_z"] if zebra else styles["pct"])
            ws.write_string(row, 5, f"{rec['days_on_cap']}/{rec['days_total']}", cfmt)
            ws.write_string(row, 6, fcc.last_day_text(rec), cfmt)
            ws.write_string(row, 7, rec["listed_by"], cfmt)
            row += 1
        row += 1
    ws.set_zoom(110)


def write_snapshots(book, styles, work, records, chart_start, chart_end) -> None:
    """One chart per issue site, same lines and axis rules as the NCCD report."""
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
        s = by_site.get(site)
        s = s[(s["Date"] >= chart_start) & (s["Date"] <= chart_end)].sort_values("ts")

        c0 = (i % CHART_DATA_PER_ROW) * 5
        r0 = (i // CHART_DATA_PER_ROW) * CHART_DATA_BLOCK_ROWS
        for offset, name in enumerate(("Date", "Time", "Rx", "Peak", "AvgPeak")):
            data.write(r0, c0 + offset, name)
        rx_by_key = {
            (d.normalize(), int(h)): float(v)
            for d, h, v in zip(s["Date"], s["Hour"], s["Rx"])
        }
        peak_72 = rec.get("peak_72")
        avg_peak = rec.get("avg_peak")
        if peak_72 is None or avg_peak is None:
            on_cap_by_key = {
                (d.normalize(), int(h)): bool(v)
                for d, h, v in zip(s["Date"], s["Hour"], s["OnCap"])
            }
            chart_rx = [v for v in rx_by_key.values() if np.isfinite(v)]
            peak_72 = float(max(chart_rx)) if chart_rx else float("nan")
            avg_peak = fcc._avg_wall_peak(
                rx_by_key, on_cap_by_key, float(rec["center"]), chart_start, chart_end
            )
        dates: list[str] = []
        times: list[str] = []
        plotted: list[float] = [v for v in (peak_72, avg_peak) if np.isfinite(v)]
        r = r0
        for day in chart_days:
            date_label = f"{int(day.day)}/{fcc.MONTHS[int(day.month) - 1]}"
            for hour in range(24):
                r += 1
                hour_label = f"{hour:02d}:00"
                dates.append(date_label if hour == 0 else "")
                times.append(hour_label)
                if hour == 0:
                    data.write_string(r, c0, date_label, cat_fmt)
                else:
                    data.write_blank(r, c0, None)
                data.write_string(r, c0 + 1, hour_label, cat_fmt)
                value = rx_by_key.get((day.normalize(), hour))
                if value is None or not np.isfinite(value):
                    data.write_blank(r, c0 + 2, None)
                else:
                    data.write_number(r, c0 + 2, value)
                    plotted.append(value)
                if np.isfinite(peak_72):
                    data.write_number(r, c0 + 3, peak_72)
                if np.isfinite(avg_peak):
                    data.write_number(r, c0 + 4, avg_peak)
        n = r
        first = r0 + 1

        ws.set_row(top, 26)
        ws.merge_range(
            top, 0, top, 7, f"{i + 1:02d}    {site}    ·    {rec['severity']}", styles["banner"]
        )
        labels = [
            (0, f"{TX_LABEL}, median (Mbit/s)", f"{rec['bw']:.2f}"),
            (2, f"Stuck {RX_LABEL} (Mbit/s)", f"{rec['center']:.2f}"),
            (
                4,
                "Hours on cap (all 24h)",
                f"{rec['hours_on']} / {rec['hours_n']}  ({rec['hours_pct']:.1f}%)",
            ),
            (6, "Days on cap (≥3h)", f"{rec['days_on_cap']} / {rec['days_total']}"),
        ]
        ws.set_row(top + 1, 16)
        ws.set_row(top + 2, 20)
        for col, label, _value in labels:
            ws.merge_range(top + 1, col, top + 1, col + 1, label, styles["kpi_label"])
        for col, _label, value in labels:
            ws.merge_range(top + 2, col, top + 2, col + 1, value, styles["kpi_value"])
        ws.set_row(top + 3, 8)

        if not plotted or n < first:
            ws.merge_range(
                top + 5, 0, top + 5, 7, "No hourly samples in the latest 3 days.", styles["note"]
            )
            for rr in range(top + 4, top + block):
                ws.set_row(rr, 15)
            continue

        chart = book.add_chart({"type": "line"})
        chart.show_blanks_as("gap")
        chart.set_title(
            {"name": site, "name_font": {"name": "Calibri", "size": 16, "color": "#595959"}}
        )
        chart.add_series(
            {
                "name": f"{RX_LABEL} (Mbps)",
                "categories": ["_ChartData", first, c0, n, c0 + 1],
                "categories_data": [dates, times],
                "values": ["_ChartData", first, c0 + 2, n, c0 + 2],
                "line": {"color": "#5B9BD5", "width": 2.25},
                "marker": {"type": "none"},
            }
        )
        if np.isfinite(peak_72):
            chart.add_series(
                {
                    "name": f"Peak ({peak_72:.2f} Mbit/s)",
                    "categories": ["_ChartData", first, c0, n, c0 + 1],
                    "categories_data": [dates, times],
                    "values": ["_ChartData", first, c0 + 3, n, c0 + 3],
                    "line": {"color": "#C00000", "width": 1.5, "dash_type": "dash"},
                    "marker": {"type": "none"},
                }
            )
        if np.isfinite(avg_peak):
            chart.add_series(
                {
                    "name": f"Avg peak ({avg_peak:.2f} Mbit/s)",
                    "categories": ["_ChartData", first, c0, n, c0 + 1],
                    "categories_data": [dates, times],
                    "values": ["_ChartData", first, c0 + 4, n, c0 + 4],
                    "line": {"color": "#ED7D31", "width": 1.5, "dash_type": "dash"},
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
    _page(ws, "Hourly FEGE KPI — choked sites")
    ws.set_tab_color("#5B9BD5")
    headers = [
        "Date",
        "Time",
        ERIC_SITE_COL,
        ERIC_DL_COL,
        "Max UL Capacity Usage (Mbps)",
        "Total Capacity Usage (Mbps)",
        "On cap",
        "Stuck level (Mbit/s)",
    ]
    for i, w in enumerate([14, 10, 16, 26, 26, 26, 12, 20]):
        ws.set_column(i, i, w)
    ws.set_row(0, 30)
    for col, text in enumerate(headers):
        ws.write(0, col, text, styles["header"])

    sites = [r["site"] for r in records]
    hourly = work.loc[work["eNodeB Name"].isin(sites)].sort_values(["eNodeB Name", "ts"])
    site_parity = {site: i % 2 for i, site in enumerate(sites)}
    cols = (
        hourly["eNodeB Name"].tolist(),
        hourly["Date"].tolist(),
        hourly["Hour"].tolist(),
        hourly[ERIC_DL_COL].tolist(),
        hourly[ERIC_UL_COL].tolist(),
        hourly[ERIC_TOTAL_COL].tolist(),
        hourly["OnCap"].tolist(),
        hourly["StuckLevel"].tolist(),
    )
    excel_row = 0
    for site, date, hour, dl, ul, total, on_cap, level in zip(*cols):
        excel_row += 1
        z = "_z" if site_parity[site] else ""
        ws.write_datetime(excel_row, 0, date.to_pydatetime(), styles["date" + z])
        ws.write_datetime(excel_row, 1, time(int(hour), 0), styles["time" + z])
        ws.write_string(excel_row, 2, site, styles["text" + z])
        for col, value in ((3, dl), (4, ul), (5, total)):
            if value is None or not np.isfinite(value):
                ws.write_blank(excel_row, col, None, styles["num" + z])
            else:
                ws.write_number(excel_row, col, float(value), styles["num" + z])
        ws.write_string(excel_row, 6, "Yes" if bool(on_cap) else "No", styles["center" + z])
        ws.write_number(excel_row, 7, float(level), styles["num" + z])

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
    period_txt = f"{period_start.strftime('%d %b %Y')} – {period_end.strftime('%d %b %Y')}"
    chart_end = period_end.normalize()
    chart_start = chart_end - pd.Timedelta(days=fcc.SNAP_DAYS - 1)
    if records and "peak_72" not in records[0]:
        fcc._attach_chart_peaks(work, records, chart_start, chart_end)
    write_dashboard(book, styles, source_name, period_txt, n_sites, records)
    write_snapshots(book, styles, work, records, chart_start, chart_end)
    print("Writing hourly KPI...", flush=True)
    write_hourly(book, styles, work, records)
    book.close()


def main() -> None:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("source", nargs="?", default=SOURCE_DEFAULT)
    parser.add_argument("-o", "--output", default=OUTPUT_DEFAULT)
    args = parser.parse_args()
    t0 = pytime.perf_counter()

    source = Path(args.source)
    raw = load_ericsson(source)
    df = to_rule_frame(raw)
    start, end = set_window_from_file(df)
    fcc.RX_COL = RX_BITS_COL
    fcc.TXBW_COL = TX_KBITS_COL
    print(f"Window: {start.date()} – {end.date()} ({fcc.ANALYSIS_DAYS} days, all hours)")

    work, records = fcc.analyse(df)
    print(f"Analysis {pytime.perf_counter() - t0:.1f}s")
    output = Path(args.output)
    write_workbook(output, source.name, work, records)

    by_sev = defaultdict(int)
    by_shape = defaultdict(int)
    by_listed = defaultdict(int)
    by_when = defaultdict(int)
    last_n = 0
    for rec in records:
        by_sev[rec["severity"]] += 1
        by_shape[rec["cap_type"]] += 1
        by_listed[rec["listed_by"]] += 1
        by_when[rec["cap_when"]] += 1
        last_n += int(rec["last_flag"])
    print(f"Sites checked: {work['eNodeB Name'].nunique()}")
    print(f"Issue sites: {len(records)}")
    print("Severity (hours on cap):")
    for name in (fcc.SEV_SEVERE, fcc.SEV_HIGH, fcc.SEV_MODERATE, fcc.SEV_LOW):
        print(f"  {name}: {by_sev[name]}")
    print("Listed because:")
    for name in ("≥3 days + last day", "≥3 days", "Last day"):
        print(f"  {name}: {by_listed[name]}")
    print(f"Last-day cap flag: {last_n}")
    print("Where capped:")
    for name in fcc.CAP_WHEN_COLS:
        print(f"  {name}: {by_when[name]}")
    print("Cap shape:")
    for name in (fcc.CAP_CROWDED, fcc.CAP_CEILING):
        print(f"  {name}: {by_shape[name]}")
    print(f"Wrote {output} in {pytime.perf_counter() - t0:.1f}s")


if __name__ == "__main__":
    main()
