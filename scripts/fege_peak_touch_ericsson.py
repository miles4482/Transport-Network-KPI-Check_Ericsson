#!/usr/bin/env python3
"""Ericsson FEGE report, Peak-touch method (v3 onward).

Counter: only `Max DL Capacity Usage (Mbps)` from `FEG_KPI_Eric.xlsb`.

For every ERBS over the scored window (the 3 days in the file, every hour):

    Peak         = highest hourly value in the window (the red Peak line)
    Below5%Peak  = Peak x (1 - TOUCH_BAND)   (TOUCH_BAND = 5%)
    Touch hour   = an hour whose value is at or above Below5%Peak,
                   i.e. within 5% below the Peak
    Touch ratio  = touch hours / total hours with data

Grade by peak ratio (touch hours / hours with data):

    >= 40%  Severe
    >= 30%  High      (30% to under 40%)
    >= 20%  Medium    (20% to under 30%)
    >= 10%  Low       (10% to under 20%)
    <  10%  not listed

Tx Issues Suspected: a listed site whose license is at least 15% above its Peak.
Gap = (License - Peak) / License.

Guards kept from RULES.md: a site needs at least 24 hourly samples and a
Peak of at least 20 Mbit/s to be scored.

Sheets written: Dashboard, 1. Site List, 2. HourlyChartOfIssueSites
(with hidden _ChartData), 3. Hourly KPI, 5. GeoPlot.
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
import fege_choke_check_ericsson as eric  # noqa: E402

OUTPUT_DEFAULT = "FEGE_Choked_Flat_Sites_Ericsson_v22.xlsx"
SITE_LIST_SHEET = "1. Site List"
GEO_FILE = "Physical_Site_Database_14Oct26.xlsx"
TOUCH_LEVEL_NAME = "Below5%Peak"
LICENSE_FILE = "DL_License_File.xlsb"
REPORT_PERIOD = ""
LICENSE_SHEET = "Dashboard_License"
LICENSE_NODE_COL = "NodeId"
LICENSE_CAP_COL = "dlBbCapacityNet"
# An hour counts when Max DL Capacity Usage / dlBbCapacityNet is at or above
# this share. A site is listed when such hours are at least LICENSE_HOURS_PCT
# of its hours.
LICENSE_USE = 0.85
LICENSE_HOURS_PCT = 20.0
# Share of hours at ≥ LICENSE_USE. Under LICENSE_HOURS_PCT is not listed.
LICENSE_CATEGORIES = (
    ("≥ 50%", 50.0, "severe", "At least 50% of the hours at ≥85% of license"),
    ("40% to under 50%", 40.0, "high", "From 40% up to 50% of the hours at ≥85% of license"),
    ("30% to under 40%", 30.0, "moderate", "From 30% up to 40% of the hours at ≥85% of license"),
    ("20% to under 30%", 20.0, "low", "From 20% up to 30% of the hours at ≥85% of license"),
)
TRACE_COL = eric.ERIC_DL_COL
TRACE_LABEL = "Max DL Capacity Usage"
TOUCH_BAND = 0.05
MIN_SAMPLES = 24
MIN_PEAK_MBPS = fcc.MIN_CENTER_MBPS

SEV_SEVERE = "Severe"
SEV_HIGH = "High"
SEV_MEDIUM = "Medium"
SEV_LOW = "Low"
SEV_DECISION = "Decision"
CAT_PEAK = "Peak Touch"
CAT_LICENSE = "License usage"
CAT_TX = "Tx Flat/Chocked"
# Gap = (License - Peak) / License. A listed site at or above 10% is Tx Flat/Chocked.
GAP_MIN = 0.10
TX_COMMENT = "Tx Issues Suspected"
TX_BANDS = (
    (SEV_SEVERE, 40.0, "Gap ≥ 40%"),
    (SEV_HIGH, 30.0, "Gap 30% to under 40%"),
    (SEV_MEDIUM, 20.0, "Gap 20% to under 30%"),
    (SEV_LOW, 15.0, "Gap 15% to under 20%"),
    (SEV_DECISION, 10.0, "Gap 10% to under 15%"),
)
SEVERITY_BANDS = (
    (SEV_SEVERE, 40.0),
    (SEV_HIGH, 30.0),
    (SEV_MEDIUM, 20.0),
    (SEV_LOW, 10.0),
)
SEVERITY_ORDER = {name: i for i, (name, _pct) in enumerate(SEVERITY_BANDS)}
SEVERITY_ORDER[SEV_DECISION] = len(SEVERITY_ORDER)
SEVERITY_STYLE = {
    SEV_SEVERE: "severe",
    SEV_HIGH: "high",
    SEV_MEDIUM: "moderate",
    SEV_LOW: "low",
    SEV_DECISION: "decision",
}
LICENSE_TYPE = {"severe": SEV_SEVERE, "high": SEV_HIGH, "moderate": SEV_MEDIUM, "low": SEV_LOW}
PEAK_RULES = {
    SEV_SEVERE: "Peak ratio ≥ 40%",
    SEV_HIGH: "Peak ratio 30% to under 40%",
    SEV_MEDIUM: "Peak ratio 20% to under 30%",
    SEV_LOW: "Peak ratio 10% to under 20%",
}
LICENSE_RULES = {
    SEV_SEVERE: "≥ 50% of hours at ≥85% of license",
    SEV_HIGH: "40% to under 50% of hours at ≥85% of license",
    SEV_MEDIUM: "30% to under 40% of hours at ≥85% of license",
    SEV_LOW: "20% to under 30% of hours at ≥85% of license",
}
CHART_DATA_PER_ROW = eric.CHART_DATA_PER_ROW
CHART_DATA_BLOCK_ROWS = eric.CHART_DATA_BLOCK_ROWS


def license_category(share: float) -> str:
    """Band for a listed license site. Share is hours at ≥85% of the license."""
    for name, floor, _style, _rule in LICENSE_CATEGORIES:
        if share >= floor:
            return name
    return LICENSE_CATEGORIES[-1][0]


def severity(touch_pct: float) -> str | None:
    for name, floor in SEVERITY_BANDS:
        if touch_pct >= floor:
            return name
    return None


def tx_band(gap: float) -> str | None:
    """Gap is (License − Peak) / License. None when the gap is under 10%."""
    if not np.isfinite(gap):
        return None
    pct = gap * 100.0
    for name, floor, _rule in TX_BANDS:
        if pct >= floor:
            return name
    return None


def apply_tx(records: list[dict], stats_by_site: dict) -> None:
    """Move a listed site into Tx Flat/Chocked when the license is ≥10% above its Peak."""
    for rec in records:
        stat = stats_by_site.get(rec["site"])
        lic = float(stat["license"]) if stat else float("nan")
        peak = float(rec["peak"])
        gap = (lic - peak) / lic if np.isfinite(lic) and lic > 0 and np.isfinite(peak) else float("nan")
        rec["license_mbps"] = lic
        rec["gap"] = gap
        rec["touch_severity"] = rec["severity"]
        band = tx_band(gap)
        rec["tx_severity"] = band
        rec["double_check"] = band is not None
        if band:
            rec["category"] = CAT_TX
            rec["severity"] = band
        else:
            rec["category"] = CAT_PEAK
    records.sort(
        key=lambda r: (
            0 if r["category"] == CAT_TX else 1,
            SEVERITY_ORDER[r["severity"]],
            -r["hours_pct"],
            r["site"],
        )
    )


def license_type(share: float) -> str:
    for _name, _floor, style, _rule in LICENSE_CATEGORIES:
        if share >= _floor:
            return LICENSE_TYPE[style]
    return LICENSE_TYPE[LICENSE_CATEGORIES[-1][2]]


def stamp_charts(charts: list[dict], records: list[dict]) -> None:
    """One category and display severity per chart. License can sit on top of the other two."""
    by_site = {rec["site"]: rec for rec in records}
    for item in charts:
        rec = by_site.get(item["site"])
        lic = item.get("license")
        item["license_listed"] = lic is not None
        if rec is not None:
            item["category"] = rec["category"]
            item["display_severity"] = rec["severity"]
        else:
            item["category"] = CAT_LICENSE
            item["display_severity"] = license_type(float(lic["share"])) if lic else SEV_LOW
    cat_rank = {CAT_TX: 0, CAT_PEAK: 1, CAT_LICENSE: 2}
    charts.sort(
        key=lambda item: (
            cat_rank[item["category"]],
            SEVERITY_ORDER.get(item["display_severity"], 9),
            -(float(item["touch"]["hours_pct"]) if item.get("touch") else 0.0),
            -(float(item["license"]["share"]) if item.get("license") else 0.0),
            item["site"],
        )
    )


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


_FCC_PAGE = fcc._page


def _page(ws, title: str, *, fit_width: bool = True) -> None:
    _FCC_PAGE(ws, title, fit_width=fit_width)
    window = f"  |  {REPORT_PERIOD}" if REPORT_PERIOD else ""
    ws.set_header(f"&L&8&K1F4E79{title}&R&8{fcc.REPORT_TITLE}  |  {eric.ZONE}{window}")
    ws.set_footer(f"&L&8Touch = {TRACE_LABEL} at or above Below5%Peak (Peak × 95%)&R&8Page &P of &N")


def site_code(erbs: str) -> str:
    """Physical site code is the first 7 characters of the ERBS."""
    return str(erbs).strip()[:7]


# GeoPlot legend uses Moderate. The Peak-touch grade for that band is Medium.
GEO_SEV = {
    SEV_SEVERE: "Severe",
    SEV_HIGH: "High",
    SEV_MEDIUM: "Moderate",
    SEV_LOW: "Low",
    SEV_DECISION: "Decision",
}
GEO_SEV_RANK = {"Severe": 0, "High": 1, "Moderate": 2, "Low": 3, "Decision": 4}


def load_physical() -> tuple[pd.DataFrame, dict[str, str], dict[str, str]]:
    """Site code, Tech, and Site Type from the physical database."""
    path = fcc._find_input(GEO_FILE)
    if path is None:
        raise SystemExit(f"Physical site database not found: {GEO_FILE}")
    geo = fcc.load_site_geo(path)
    tech = fcc.tech_by_site_from_geo(geo)
    types = fcc.site_type_by_site_from_geo(geo)
    print(f"Physical database {path.name}: {len(geo):,} sites", flush=True)
    return geo, tech, types


def attach_site_identity(
    charts: list[dict],
    tech_by_code: dict[str, str],
    type_by_code: dict[str, str],
) -> None:
    """Stamp Tech and Site Type onto each issue ERBS from its site code."""
    missing = 0
    for item in charts:
        code = site_code(item["site"])
        item["site_code"] = code
        item["tech"] = tech_by_code.get(code, fcc.TECH_4G)
        item["site_type"] = type_by_code.get(code) or "Not found"
        if code not in type_by_code:
            missing += 1
    print(f"Site list: {len(charts)} ERBS, {missing} with no physical-database row", flush=True)


def geo_issue_records(charts: list[dict]) -> tuple[list[dict], int]:
    """One GeoPlot point per physical site that has a dashboard issue ERBS.

    Severity is the worst Peak-touch grade among those ERBS. Medium is drawn
    as Moderate so it uses the existing map legend. A site listed only by the
    license rule (no Peak-touch grade) is drawn as Low.
    """
    by_code: dict[str, dict] = {}
    for item in charts:
        code = item["site_code"]
        sev = GEO_SEV.get(item.get("display_severity")) or "Low"
        cur = by_code.get(code)
        if cur is None:
            by_code[code] = {
                "site": code,
                "severity": sev,
                "license_only": item.get("category") == CAT_LICENSE,
            }
            continue
        if item.get("category") != CAT_LICENSE:
            cur["license_only"] = False
        if GEO_SEV_RANK[sev] < GEO_SEV_RANK[cur["severity"]]:
            cur["severity"] = sev
    records = [{"site": row["site"], "severity": row["severity"]} for row in by_code.values()]
    records.sort(key=lambda row: (GEO_SEV_RANK[row["severity"]], row["site"]))
    license_only = sum(1 for row in by_code.values() if row["license_only"])
    return records, license_only


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


def _blank_or_num(ws, row: int, col: int, value: int, num_fmt, blank_fmt) -> None:
    if value:
        ws.write_number(row, col, value, num_fmt)
    else:
        ws.write_blank(row, col, None, blank_fmt)


def _write_category_matrix(book, ws, styles, row: int, records: list[dict], license_records: list[dict], tf) -> int:
    """Dashboard category table: Tx Flat/Chocked, then Peak Touch, then License usage."""
    comment_fmt = book.add_format(
        {
            "font_name": "Calibri",
            "font_size": 11,
            "bold": True,
            "font_color": "#C00000",
            "align": "center",
            "valign": "vcenter",
            "border": 1,
            "border_color": "#D9D9D9",
            "text_wrap": True,
        }
    )
    rule_fmt = book.add_format(
        {
            "font_name": "Calibri",
            "font_size": 10,
            "align": "left",
            "valign": "vcenter",
            "border": 1,
            "border_color": "#D9D9D9",
            "text_wrap": True,
        }
    )
    cat_fmt = book.add_format(
        {
            "font_name": "Calibri",
            "font_size": 11,
            "bold": True,
            "font_color": fcc.NAVY,
            "align": "center",
            "valign": "vcenter",
            "bg_color": "#DEEBF7",
            "border": 1,
            "border_color": "#D9D9D9",
            "text_wrap": True,
        }
    )
    blocks = []
    tx_rows = []
    for name, _floor, rule in TX_BANDS:
        chosen = [rec for rec in records if rec["category"] == CAT_TX and rec["severity"] == name]
        tx_rows.append((name, chosen, TX_COMMENT, rule, rec_when(chosen)))
    blocks.append((CAT_TX, tx_rows))
    peak_rows = []
    for name, _floor in SEVERITY_BANDS:
        chosen = [rec for rec in records if rec["category"] == CAT_PEAK and rec["severity"] == name]
        peak_rows.append((name, chosen, "", PEAK_RULES[name], rec_when(chosen)))
    blocks.append((CAT_PEAK, peak_rows))
    lic_rows = []
    for sev in (SEV_SEVERE, SEV_HIGH, SEV_MEDIUM, SEV_LOW):
        chosen = [rec for rec in license_records if license_type(rec["share"]) == sev]
        lic_rows.append((sev, chosen, "", LICENSE_RULES[sev], rec_when(chosen)))
    blocks.append((CAT_LICENSE, lic_rows))

    ws.set_row(row, 22)
    ws.merge_range(row, 0, row, 7, "Category summary", tf["group"])
    row += 1
    ws.set_row(row, 30)
    for col, text in enumerate(
        ("Category", "Severity", *fcc.CAP_WHEN_COLS, "Total", "Comments", "Rule Details")
    ):
        ws.write(row, col, text, styles["header"])
    row += 1
    grand = {when: 0 for when in fcc.CAP_WHEN_COLS}
    for label, block_rows in blocks:
        start = row
        for sev, _chosen, comment, rule, when_counts in block_rows:
            ws.set_row(row, 22)
            ws.write(row, 1, sev, styles[SEVERITY_STYLE[sev]])
            total = 0
            for offset, when in enumerate(fcc.CAP_WHEN_COLS):
                value = when_counts.get(when, 0)
                total += value
                grand[when] += value
                _blank_or_num(ws, row, 2 + offset, value, styles["int"], styles["center"])
            ws.write_number(row, 5, total, styles["int"])
            if comment:
                ws.write(row, 6, comment, comment_fmt)
            else:
                ws.write_blank(row, 6, None, styles["center"])
            ws.write(row, 7, rule, rule_fmt)
            row += 1
        if row - 1 == start:
            ws.write(start, 0, label, cat_fmt)
        else:
            ws.merge_range(start, 0, row - 1, 0, label, cat_fmt)
    ws.set_row(row, 18)
    ws.merge_range(row, 0, row, 1, "Grand Total", tf["total_label"])
    grand_n = sum(grand.values())
    for offset, when in enumerate(fcc.CAP_WHEN_COLS):
        ws.write_number(row, 2 + offset, grand[when], tf["total"])
    ws.write_number(row, 5, grand_n, tf["total"])
    ws.write_blank(row, 6, None, tf["total"])
    ws.write(row, 7, "Distinct findings added across the three categories. A site in two categories is counted in each.", rule_fmt)
    row += 1
    ws.set_row(row, 32)
    ws.merge_range(
        row, 0, row, 7,
        "Grand Total adds the rows above, so an ERBS in License usage and also in Peak Touch or "
        "Tx Flat/Chocked is counted twice here. Sheet 1. Site List and the charts keep that ERBS once.",
        styles["note"],
    )
    return row + 1


def rec_when(chosen: list[dict]) -> dict[str, int]:
    counts = {when: 0 for when in fcc.CAP_WHEN_COLS}
    for rec in chosen:
        when = rec.get("cap_when")
        if when in counts:
            counts[when] += 1
    return counts


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
    oss_note: str = "",
) -> None:
    ws = book.add_worksheet("Dashboard")
    _page(ws, fcc.REPORT_TITLE)
    ws.set_tab_color("#C65911")
    for i, w in enumerate([22, 16, 22, 16, 18, 22, 14, 16, 18, 18, 16, 12, 22]):
        ws.set_column(i, i, w)
    tf = eric._tile_formats(book)

    counts = defaultdict(int)
    for rec in records:
        counts[rec["severity"]] += 1

    ws.set_row(0, 28)
    ws.merge_range("A1:L1", fcc.REPORT_TITLE, styles["title"])
    ws.set_row(1, 18)
    ws.merge_range(
        "A2:L2",
        f"{eric.ZONE} transmission    ·    {period_txt}    ·    hourly    ·    {source_name}",
        styles["subtitle"],
    )
    ws.set_row(2, 150 if oss_note else 100)
    method = (
        f"Method: only {TRACE_COL} is read. Peak = the highest hourly value in the window "
        f"(red line). {TOUCH_LEVEL_NAME} = Peak × {100 - TOUCH_BAND * 100:.0f}% (orange line). "
        f"It is {TOUCH_BAND * 100:.0f}% under that site's own Peak. It is not an average of the hours. "
        f"An hour touches when its value is at or above {TOUCH_LEVEL_NAME}. "
        "Three categories, with Tx Flat/Chocked first. "
        "Tx Flat/Chocked: a listed site whose Gap = (License − Peak) / License is at least 10%. "
        "Severe ≥ 40%, High 30% to under 40%, Medium 20% to under 30%, Low 15% to under 20%, "
        "Decision 10% to under 15%. Comment is Tx Issues Suspected. "
        "Peak Touch: peak ratio Severe ≥ 40%, High 30% to under 40%, "
        "Medium 20% to under 30%, Low 10% to under 20%. Under 10% is not listed. "
        "License usage: share of hours at ≥85% of license, Severe ≥ 50%, High 40% to under 50%, "
        "Medium 30% to under 40%, Low 20% to under 30%. "
        f"A site is scored when it has at least {MIN_SAMPLES} hourly samples and a Peak of at "
        f"least {MIN_PEAK_MBPS:.0f} Mbit/s."
    )
    if oss_note:
        method = f"{method} {oss_note}"
    ws.merge_range("A3:L3", method, styles["note"])

    license_records = (license_summary or {}).get("records") or []
    n_peak = sum(1 for rec in records if rec["category"] == CAT_PEAK)
    n_tx = sum(1 for rec in records if rec["category"] == CAT_TX)
    row = eric._write_tiles(
        ws,
        3,
        [
            (0, "Sites checked", n_sites, False),
            (2, CAT_PEAK, n_peak, False),
            (4, CAT_TX, n_tx, True),
            (6, CAT_LICENSE, len(license_records), True),
        ],
        tf,
    )
    row += 1
    row = _write_category_matrix(book, ws, styles, row, records, license_records, tf)
    row += 2

    headers = [
        "No.",
        "ERBS",
        "Severity",
        "Peak (Mbit/s)",
        f"{TOUCH_LEVEL_NAME} (Mbit/s)",
        "Touch hours",
        "Total hours",
        "Peak ratio (%)",
        "Busy-hour touch (%)",
        "Where touched",
        "License (Mbps)",
        "Gap (%)",
    ]
    indexed = list(enumerate(records))
    sections: list[tuple[str, str, list]] = []
    for name, _floor, rule in TX_BANDS:
        group = [
            (i, rec) for i, rec in indexed
            if rec["category"] == CAT_TX and rec["severity"] == name
        ]
        sections.append((f"{CAT_TX} · {name}", f"{rule}. {TX_COMMENT}.", group))
    for name, _floor in SEVERITY_BANDS:
        group = [
            (i, rec) for i, rec in indexed
            if rec["category"] == CAT_PEAK and rec["severity"] == name
        ]
        sections.append((f"{CAT_PEAK} · {name}", PEAK_RULES[name], group))
    for name, note, group in sections:
        if not group:
            continue
        ws.set_row(row, 22)
        ws.merge_range(
            row,
            0,
            row,
            11,
            f"{name}    ·    {len(group)} site{'s' if len(group) != 1 else ''}",
            styles["section"],
        )
        row += 1
        ws.set_row(row, 24)
        ws.merge_range(row, 0, row, 11, note, styles["body"])
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
            gap = rec.get("gap", float("nan"))
            lic = rec.get("license_mbps", float("nan"))
            if np.isfinite(lic):
                ws.write_number(row, 10, float(lic), nfmt)
            else:
                ws.write_blank(row, 10, None, nfmt)
            if np.isfinite(gap):
                ws.write_number(row, 11, float(gap) * 100.0, pfmt)
            else:
                ws.write_blank(row, 11, None, nfmt)
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
            top, 0, top, 7,
            f"{i + 1:02d}    {site}    ·    {REPORT_PERIOD}    ·    {rec['severity']}",
            styles["banner"],
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

    all_stats = []
    for site, g in df[scored].groupby("site", sort=False):
        hours_n = int(len(g))
        hours_on = int(g["hit"].sum())
        license_mbps = float(g["lic"].iloc[0])
        peak = float(g["rx"].max())
        busy = g["busy"].to_numpy()
        hit = g["hit"].to_numpy()
        all_stats.append(
            {
                "site": site,
                "license": license_mbps,
                "hours_on": hours_on,
                "hours_n": hours_n,
                "share": 100.0 * hours_on / hours_n,
                "peak": peak,
                "peak_ratio": 100.0 * peak / license_mbps,
                "median_ratio": 100.0 * float(np.median(g["rx"] / license_mbps)),
                "cap_when": fcc.cap_window(int(hit[busy].sum()), int(hit[~busy].sum())),
            }
        )
    records = [rec for rec in all_stats if rec["share"] >= LICENSE_HOURS_PCT]
    records.sort(key=lambda r: (-r["share"], r["site"]))
    listed = [rec["site"] for rec in records]
    hourly = df.loc[df["site"].isin(listed)].copy()
    hourly["ratio"] = hourly["rx"] / hourly["lic"]
    return {
        "records": records,
        "hourly": hourly,
        "stats_by_site": {rec["site"]: rec for rec in all_stats},
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
    ws.set_row(row, 64)
    ws.merge_range(
        row,
        0,
        row,
        9,
        f"dlBbCapacityNet is the DL capacity license installed on the node "
        f"({LICENSE_NODE_COL} = ERBS). An hour counts when {TRACE_COL} / dlBbCapacityNet "
        f"≥ {use_pct}%. A site is listed when those hours are at least "
        f"{LICENSE_HOURS_PCT:.0f}% of its hours. "
        f"License usage type is that share of hours: Severe ≥ 50%, High 40% to under 50%, "
        f"Medium 30% to under 40%, Low 20% to under 30%. "
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
    counts = {name: 0 for name, _floor, _style, _rule in LICENSE_CATEGORIES}
    for rec in records:
        counts[license_category(rec["share"])] += 1
    ws.set_row(row, 22)
    ws.merge_range(row, 0, row, 9, "License usage category", styles["section"])
    row += 1
    ws.set_row(row, 28)
    ws.write(row, 0, "Category", styles["header"])
    ws.merge_range(row, 1, row, 3, f"Share of hours at ≥{use_pct}% of license", styles["header"])
    ws.write(row, 4, "Sites", styles["header"])
    ws.write(row, 5, "Type", styles["header"])
    row += 1
    for name, _floor, style, rule in LICENSE_CATEGORIES:
        ws.set_row(row, 18)
        ws.write(row, 0, name, styles[style])
        ws.merge_range(row, 1, row, 3, rule, styles["text"])
        ws.write_number(row, 4, counts[name], styles["int"])
        ws.write(row, 5, LICENSE_TYPE[style], styles[style])
        row += 1
    ws.set_row(row, 18)
    ws.write(row, 0, "Total listed", styles["header"])
    ws.merge_range(
        row, 1, row, 3,
        f"At least {LICENSE_HOURS_PCT:.0f}% of hours at ≥{use_pct}% of license",
        styles["header"],
    )
    ws.write_number(row, 4, len(records), styles["int"])
    row += 2

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
    for name, _floor, _style, _rule in LICENSE_CATEGORIES:
        group_rows = [rec for rec in records if license_category(rec["share"]) == name]
        if not group_rows:
            continue
        ws.set_row(row, 22)
        ws.merge_range(
            row, 0, row, 9,
            f"{name} of hours at ≥{use_pct}% of license    ·    {len(group_rows)} sites",
            styles["section"],
        )
        row += 1
        ws.set_row(row, 32)
        for col, text in enumerate(headers):
            ws.write(row, col, text, styles["header"])
        row += 1
        for n, rec in enumerate(group_rows):
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
            f"{i + 1:02d}    {site}    ·    {REPORT_PERIOD}    ·    {rec['share']:.1f}% of hours at ≥{use_pct:.0f}% of license",
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


def trace_lines(s: pd.DataFrame) -> dict:
    """Peak and Below5%Peak from the site's own hourly Max DL trace."""
    if s is None or s.empty:
        return {
            "peak": float("nan"),
            "level": float("nan"),
            "hours_on": 0,
            "hours_n": 0,
            "hours_pct": float("nan"),
            "cap_when": "—",
        }
    rx = s["Rx"].to_numpy(dtype=float)
    hours = s["Hour"].to_numpy(dtype=int)
    valid = np.isfinite(rx)
    peak = float(np.nanmax(rx)) if valid.any() else float("nan")
    level = peak * (1.0 - TOUCH_BAND) if np.isfinite(peak) else float("nan")
    band = valid & (rx >= level) if np.isfinite(level) else np.zeros(len(rx), dtype=bool)
    busy = (hours >= fcc.BUSY_START) & (hours <= fcc.BUSY_END)
    hours_n = int(valid.sum())
    hours_on = int(band.sum())
    return {
        "peak": peak,
        "level": level,
        "hours_on": hours_on,
        "hours_n": hours_n,
        "hours_pct": 100.0 * hours_on / hours_n if hours_n else float("nan"),
        "cap_when": fcc.cap_window(int((band & busy).sum()), int((band & ~busy).sum())),
    }


def attach_chart_lines(work: pd.DataFrame, charts: list[dict], stats_by_site: dict) -> None:
    """Give every issue site the five-line inputs, including sites on only one list."""
    by_site = {site: g for site, g in work.groupby("eNodeB Name", sort=False)}
    missing = []
    for item in charts:
        item["lines"] = trace_lines(by_site.get(item["site"]))
        stat = stats_by_site.get(item["site"])
        item["lic_stat"] = stat
        item["license_mbps"] = float(stat["license"]) if stat else float("nan")
        if not np.isfinite(item["license_mbps"]):
            missing.append(item["site"])
    if missing:
        print(f"  {len(missing)} issue sites have no positive license, so those two lines are blank")


OSS_PART_GLOB = "OSS KPI Hourly Cel Level_*.part1.rar"
OSS_CACHE = Path("/tmp/eric_oss_site_hour.pkl")
ERIC_OSS_RENAME = {
    "EUtranCell Id": fcc.OSS_CELL_COL,
    "Total Traffic (GB)": fcc.OSS_VOL_COL,
    "Average User Throughput in DownLink (Mbps)_Num": fcc.OSS_TP_N_COL,
    "Average User Throughput in DownLink (Mbps)_Denom": fcc.OSS_TP_D_COL,
    "Average Number of Users in a Cell": fcc.OSS_MEAN_USER_COL,
    "Maximum Number of Users in a Cell": fcc.OSS_USER_COL,
    "DL PRB Utilization (%)_Num": fcc.OSS_PRB_N_COL,
    "DL PRB Utilization (%)_Denom": fcc.OSS_PRB_D_COL,
}
OSS_VALUE_KEYS = (
    "tp_data_4h",
    "prb_data_4h",
    "vol_data_4h",
    "users_data_4h",
    "mean_data_4h",
    "tp_user_4h",
    "prb_user_4h",
    "vol_user_4h",
    "users_user_4h",
    "mean_user_4h",
    "tp_data_1h",
    "prb_data_1h",
    "vol_data_1h",
    "users_data_1h",
    "mean_data_1h",
    "tp_user_1h",
    "prb_user_1h",
    "vol_user_1h",
    "users_user_1h",
    "mean_user_1h",
)


def load_ericsson_oss() -> pd.DataFrame:
    """Cell KPI → one row per physical site-hour. L900 cells are dropped.

    V1 and V2 are the two halves of the day (hours 0–11 and 12–23). The site
    code is the first 7 characters of EUtranCell Id, the same key as the ERBS.
    """
    parts = sorted(Path("/workspace").glob(OSS_PART_GLOB))
    if not parts:
        raise SystemExit(f"No OSS parts matching {OSS_PART_GLOB}")
    stamp = tuple((path.name, path.stat().st_mtime_ns) for path in parts)
    if OSS_CACHE.exists():
        cached = pickle.loads(OSS_CACHE.read_bytes())
        if cached.get("stamp") == stamp:
            frame = cached["frame"]
            print(f"OSS cache: {len(frame):,} site-hours, {frame['site'].nunique():,} sites", flush=True)
            return frame
    frames = []
    for part1 in parts:
        dest = Path("/tmp/oss_cell_extract") / part1.name.replace(".part1.rar", "")
        csv_path = fcc._extract_split_rar(part1, dest)
        print(f"Loaded {csv_path.name}", flush=True)
        frames.append(pd.read_csv(csv_path))
    raw = pd.concat(frames, ignore_index=True)
    raw = raw.rename(columns=ERIC_OSS_RENAME)
    raw = raw.dropna(subset=["Date", "Hour", fcc.OSS_CELL_COL])
    cell = raw[fcc.OSS_CELL_COL].astype(str).str.strip()
    raw[fcc.OSS_CELL_COL] = cell
    raw = raw.loc[~cell.str.lower().isin(["", "nan", "none"])].copy()
    text = raw["Date"].astype(str).str.strip()
    text = text.mask(text.str.fullmatch(r"\d{1,2}-[A-Za-z]{3}"), text + "-2026")
    raw["Date"] = pd.to_datetime(text, format="%d-%b-%Y", errors="coerce").dt.normalize()
    raw["Hour"] = pd.to_numeric(raw["Hour"], errors="coerce")
    raw = raw.dropna(subset=["Date", "Hour"])
    raw["Hour"] = raw["Hour"].astype(int)
    raw["eNodeB Name"] = raw[fcc.OSS_CELL_COL].str[:7]
    raw["Time"] = raw["Hour"].map(lambda hour: f"{int(hour):02d}:00")
    n_cell = len(raw)
    kept = int(fcc._capacity_mask(raw[fcc.OSS_CELL_COL]).sum())
    print(f"L900 coverage dropped: {n_cell - kept:,} of {n_cell:,} cells", flush=True)
    site = fcc.aggregate_capacity_cells_to_site(raw)
    frame = pd.DataFrame(
        {
            "Date": pd.to_datetime(site["Date"]).dt.normalize(),
            "Hour": site["Time"].astype(str).str.slice(0, 2).astype(int),
            "site": site["eNodeB Name"].astype(str).str.strip(),
            "vol": pd.to_numeric(site[fcc.OSS_VOL_COL], errors="coerce"),
            "users": pd.to_numeric(site[fcc.OSS_USER_COL], errors="coerce"),
            "mean_user": pd.to_numeric(site[fcc.OSS_MEAN_USER_COL], errors="coerce"),
            "tp_n": pd.to_numeric(site[fcc.OSS_TP_N_COL], errors="coerce"),
            "tp_d": pd.to_numeric(site[fcc.OSS_TP_D_COL], errors="coerce"),
            "prb_n": pd.to_numeric(site[fcc.OSS_PRB_N_COL], errors="coerce"),
            "prb_d": pd.to_numeric(site[fcc.OSS_PRB_D_COL], errors="coerce"),
        }
    )
    frame = frame.drop_duplicates(["site", "Date", "Hour"], keep="last")
    OSS_CACHE.write_bytes(pickle.dumps({"stamp": stamp, "frame": frame}))
    days = ", ".join(day.strftime("%-d %b") for day in sorted(frame["Date"].unique()))
    print(
        f"OSS site-hours: {len(frame):,}  ·  sites {frame['site'].nunique():,}  ·  {days}",
        flush=True,
    )
    return frame


def site_oncap_frame(work: pd.DataFrame) -> pd.DataFrame:
    """A physical site-hour is on-cap when any of its ERBS touches Below5%Peak."""
    tmp = pd.DataFrame(
        {
            "eNodeB Name": work["eNodeB Name"].map(site_code),
            "Date": pd.to_datetime(work["Date"]).dt.normalize(),
            "Hour": work["Hour"].astype(int),
            "OnCap": work["OnCap"].astype(bool),
        }
    )
    return tmp.groupby(["eNodeB Name", "Date", "Hour"], as_index=False)["OnCap"].any()


def oss_issue_records(charts: list[dict]) -> list[dict]:
    """One Peak-touch grade per physical site, worst ERBS wins."""
    rank = {SEV_SEVERE: 0, SEV_HIGH: 1, SEV_MEDIUM: 2, SEV_LOW: 3}
    best: dict[str, dict] = {}
    for item in charts:
        code = item["site_code"]
        touch = item["touch"]
        sev = touch["severity"] if touch else None
        grade = (touch.get("touch_severity") or sev) if touch else None
        cand = rank.get(grade, 9) if grade else 9
        cur = best.get(code)
        if cur is not None and cand >= cur["_rank"]:
            continue
        best[code] = {
            "site": code,
            "severity": sev or "—",
            "hours_pct": float(touch["hours_pct"]) if touch else 0.0,
            "tech": item.get("tech") or fcc.TECH_4G,
            "site_type": item.get("site_type") or "Not found",
            "_rank": cand,
        }
    return [{key: value for key, value in row.items() if key != "_rank"} for row in best.values()]


def prepare_oss(
    work: pd.DataFrame,
    charts: list[dict],
    tech_by_code: dict[str, str],
    type_by_code: dict[str, str],
) -> tuple[list[dict], list[str], str]:
    """Score 7/4/1 busy hours. Returns busy rows, missing-FEG site codes, and a note."""
    oss = load_ericsson_oss()
    start = pd.Timestamp(oss["Date"].min())
    end = pd.Timestamp(oss["Date"].max())
    feg_start = pd.Timestamp(pd.to_datetime(work["Date"]).min()).normalize()
    feg_end = pd.Timestamp(pd.to_datetime(work["Date"]).max()).normalize()
    fcc.ANALYSIS_START = start
    fcc.ANALYSIS_END = end
    fcc.OSS_FILE = f"OSS cell KPI {start.strftime('%-d %b')}–{end.strftime('%-d %b %Y')}, L900 dropped"
    busy = fcc.analyse_urgent_busy_hours(
        oss, oss_issue_records(charts), site_oncap_frame(work), tech_by_code
    )
    for row in busy:
        code = row["site"]
        row["site_type"] = type_by_code.get(code) or row.get("site_type") or "Not found"
        row["tech"] = tech_by_code.get(code, row.get("tech") or fcc.TECH_4G)
    feg_codes = set(work["eNodeB Name"].map(site_code))
    missing = sorted(set(oss["site"]) - feg_codes)
    overlap = feg_start <= end.normalize() and start.normalize() <= feg_end
    if overlap:
        calendar = (
            "The two calendars overlap, so Hours matched counts a busy hour that is also "
            "a Below5%Peak touch hour on the same date."
        )
    else:
        calendar = (
            "Those calendars do not overlap, so Hours matched is 0 and the choke flags stay blank."
        )
    note = (
        f"OSS KPI is {start.strftime('%-d %b')}–{end.strftime('%-d %b %Y')} "
        f"({oss['site'].nunique():,} sites after L900 is dropped). "
        f"FEG touch hours are {feg_start.strftime('%-d %b')}–{feg_end.strftime('%-d %b %Y')}. "
        f"{calendar} "
        "DL Throughput, DLPRBUtilization,%, Data Volume, Max User, and Mean User are "
        "taken from the OSS busy hours. "
        f"{len(missing):,} OSS sites have no ERBS in the FEG file (sheet 6. Missing FEG Port)."
    )
    print(
        f"BusyHour sites {len(busy):,}  ·  OSS sites missing from FEG {len(missing):,}",
        flush=True,
    )
    return busy, missing, note


def _bh_yes(values: tuple[float, float], test) -> str | None:
    """Yes when either 1h value matches. No when values exist and neither matches."""
    finite = [value for value in values if np.isfinite(value)]
    if not finite:
        return None
    return "Yes" if any(test(value) for value in finite) else "No"


def _finite(value: object) -> float:
    try:
        number = float(value)
    except (TypeError, ValueError):
        return float("nan")
    return number if np.isfinite(number) else float("nan")


OSS_BLOCKS = (
    ("DataBusyHour (4 BusyHour)", "4h", "#C65911", (
        "DL Throughput 4h (Mbps)",
        "DLPRBUtilization,%",
        "Data Volume 4h (GB)",
        "Max User 4h",
        "Mean User 4h",
    )),
    ("UserBusyHour (4 BusyHour)", "4h", "#B7950B", (
        "DL Throughput 4h (Mbps)",
        "DLPRBUtilization,%",
        "Data Volume 4h (GB)",
        "Max User 4h",
        "Mean User 4h",
    )),
    ("DataBusyHour (1 BusyHour)", "1h", "#6C3483", (
        "DL Throughput 1h (Mbps)",
        "DLPRBUtilization,%",
        "Data Volume 1h (GB)",
        "Max User 1h",
        "Mean User 1h",
    )),
    ("UserBusyHour (1 BusyHour)", "1h", "#1A5276", (
        "DL Throughput 1h (Mbps)",
        "DLPRBUtilization,%",
        "Data Volume 1h (GB)",
        "Max User 1h",
        "Mean User 1h",
    )),
)


def _listed_because(item: dict) -> str:
    parts = []
    if item.get("category") == CAT_TX:
        parts.append(CAT_TX)
    elif item.get("category") == CAT_PEAK:
        parts.append(CAT_PEAK)
    elif item.get("category") == CAT_LICENSE:
        parts.append(CAT_LICENSE)
    if item.get("license_listed") and item.get("category") != CAT_LICENSE:
        parts.append(CAT_LICENSE)
    return " + ".join(parts) or CAT_LICENSE


def write_site_list(
    book,
    styles,
    period_txt,
    charts: list[dict],
    busy_by_code: dict[str, dict],
) -> None:
    """Issue sites from the Dashboard. Tech, Site Type, and OSS 4H/1H are filled."""
    ws = book.add_worksheet(SITE_LIST_SHEET)
    _page(ws, "Site list", fit_width=False)
    ws.set_footer("&L&8OSS 4H/1H from the cell KPI, rolled to the first 7 characters of the ERBS&R&8Page &P of &N")
    ws.set_tab_color("#375623")
    use_pct = f"{LICENSE_USE * 100:.0f}"
    left = ["No.", "ERBS", "Tech", "Site Type", "Category", "Severity"]
    bh_flag = [
        "DL Throughput<7Mbps",
        "Data Vol>50GB",
        "DL Throughput<5Mbps",
        "Data Vol>100GB",
    ]
    right = [
        "Peak (Mbit/s)",
        f"{TOUCH_LEVEL_NAME} (Mbit/s)",
        "Touch hours",
        "Total hours",
        "Peak ratio (%)",
        "Where touched",
        "License (Mbps)",
        "Gap (%)",
        f"Hours ≥{use_pct}%",
        "Share of hours (%)",
        f"Where ≥{use_pct}%",
    ]
    widths = [8, 16, 12, 14, 28, 16, 22, 16, 22, 16, 16, 24, 14, 14, 16, 18, 16, 12, 14, 18, 18]
    for i, w in enumerate(widths):
        ws.set_column(i, i, w)
    bh_start = len(left)
    right_start = bh_start + len(bh_flag)
    oss_start = right_start + len(right)
    col = oss_start
    for _title, _kind, _color, names in OSS_BLOCKS:
        for _name in names:
            ws.set_column(col, col, 18)
            col += 1
    ws.set_row(0, 22)
    ws.merge_range(0, 0, 0, col - 1, "Site list — Dashboard issue sites", styles["title"])
    ws.set_row(1, 48)
    ws.merge_range(
        1, 0, 1, col - 1,
        f"{period_txt}. Tech and Site Type are copied from {GEO_FILE}. "
        "OSS 4H/1H numbers are the physical site (first 7 characters of the ERBS): "
        "L900 cells dropped, then Data Volume and Max User pick the busy hours. "
        "DL Throughput and DLPRBUtilization,% are recomputed from Num/Denom. "
        "BH Flag 1H is Yes when either 1 BusyHour block matches: "
        "DL Throughput<7Mbps at or below 7 Mbps, Data Vol>50GB at least 50 GB, "
        "DL Throughput<5Mbps at or below 5 Mbps, Data Vol>100GB at least 100 GB. "
        "A blank cell means that 1h value is missing in both blocks. "
        "There is no Avg peak line. Below5%Peak is Peak × 95%. "
        "Tx Flat/Chocked: listed, and (License − Peak) / License is at least 10%. Comment: Tx Issues Suspected.",
        styles["note"],
    )
    header_row = 3
    ws.set_row(header_row, 18)
    ws.set_row(header_row + 1, 36)
    for c, text in enumerate(left):
        ws.merge_range(header_row, c, header_row + 1, c, text, styles["header"])
    for c, text in enumerate(right):
        ws.merge_range(header_row, right_start + c, header_row + 1, right_start + c, text, styles["header"])
    bh_group = book.add_format(
        {
            "font_name": "Calibri",
            "bold": True,
            "font_size": 10,
            "font_color": "white",
            "align": "center",
            "valign": "vcenter",
            "bg_color": "#1F4E79",
            "text_wrap": True,
            "border": 1,
            "border_color": "white",
        }
    )
    bh_sub = book.add_format(
        {
            "font_name": "Calibri",
            "bold": True,
            "font_size": 9,
            "font_color": "white",
            "align": "center",
            "valign": "vcenter",
            "bg_color": "#1F4E79",
            "text_wrap": True,
            "border": 1,
            "border_color": "white",
        }
    )
    ws.merge_range(
        header_row, bh_start, header_row, bh_start + len(bh_flag) - 1, "BH Flag 1H", bh_group
    )
    for offset, name in enumerate(bh_flag):
        ws.write(header_row + 1, bh_start + offset, name, bh_sub)
    yes_fmt = book.add_format(
        {
            "font_name": "Calibri",
            "bold": True,
            "font_size": 11,
            "font_color": "#006100",
            "align": "center",
            "valign": "vcenter",
            "bg_color": "#C6EFCE",
            "border": 1,
            "border_color": "#D9D9D9",
        }
    )
    col = oss_start
    for title, _kind, color, names in OSS_BLOCKS:
        group = book.add_format(
            {
                "font_name": "Calibri",
                "bold": True,
                "font_size": 10,
                "font_color": "white",
                "align": "center",
                "valign": "vcenter",
                "bg_color": color,
                "text_wrap": True,
                "border": 1,
                "border_color": "white",
            }
        )
        sub = book.add_format(
            {
                "font_name": "Calibri",
                "bold": True,
                "font_size": 9,
                "font_color": "white",
                "align": "center",
                "valign": "vcenter",
                "bg_color": color,
                "text_wrap": True,
                "border": 1,
                "border_color": "white",
            }
        )
        if len(names) == 1:
            ws.write(header_row, col, title, group)
        else:
            ws.merge_range(header_row, col, header_row, col + len(names) - 1, title, group)
        for offset, name in enumerate(names):
            ws.write(header_row + 1, col + offset, name, sub)
        col += len(names)

    for i, item in enumerate(charts):
        row = header_row + 2 + i
        zebra = i % 2 == 1
        ws.set_row(row, 18)
        nfmt = styles["num_z"] if zebra else styles["num"]
        ifmt = styles["int_z"] if zebra else styles["int"]
        cfmt = styles["center_z"] if zebra else styles["center"]
        pfmt = styles["pct_z"] if zebra else styles["pct"]
        tfmt = styles["text_z"] if zebra else styles["text"]
        link = styles["link_z"] if zebra else styles["link"]
        lines = item["lines"]
        touch = item["touch"]
        stat = item.get("lic_stat")
        severity = item.get("display_severity") or (touch["severity"] if touch else "—")
        where_touch = touch["cap_when"] if touch else lines["cap_when"]
        ws.write_number(row, 0, i + 1, ifmt)
        ws.write_url(
            row, 1, f"internal:'{fcc.SNAP_SHEET}'!A{i * fcc.BLOCK_ROWS + 1}", link, string=item["site"]
        )
        ws.write_string(row, 2, item.get("tech") or fcc.TECH_4G, cfmt)
        ws.write_string(row, 3, item.get("site_type") or "Not found", cfmt)
        ws.write_string(row, 4, _listed_because(item), tfmt)
        sev_style = SEVERITY_STYLE.get(severity)
        ws.write_string(row, 5, severity, styles[sev_style] if sev_style else cfmt)
        busy = busy_by_code.get(item.get("site_code", ""), {})
        tp_1h = (_finite(busy.get("tp_data_1h")), _finite(busy.get("tp_user_1h")))
        vol_1h = (_finite(busy.get("vol_data_1h")), _finite(busy.get("vol_user_1h")))
        flags = (
            _bh_yes(tp_1h, lambda value: value <= 7.0),
            _bh_yes(vol_1h, lambda value: value >= 50.0),
            _bh_yes(tp_1h, lambda value: value <= 5.0),
            _bh_yes(vol_1h, lambda value: value >= 100.0),
        )
        for offset, flag in enumerate(flags):
            if flag == "Yes":
                ws.write_string(row, bh_start + offset, "Yes", yes_fmt)
            elif flag == "No":
                ws.write_string(row, bh_start + offset, "No", cfmt)
            else:
                ws.write_blank(row, bh_start + offset, None, cfmt)
        ws.write_number(row, right_start, lines["peak"], nfmt)
        ws.write_number(row, right_start + 1, lines["level"], nfmt)
        ws.write_number(row, right_start + 2, lines["hours_on"], ifmt)
        ws.write_number(row, right_start + 3, lines["hours_n"], ifmt)
        ws.write_number(row, right_start + 4, lines["hours_pct"], pfmt)
        ws.write_string(row, right_start + 5, where_touch, cfmt)
        peak_mbps = float(lines["peak"])
        if stat:
            lic_mbps = float(stat["license"])
            ws.write_number(row, right_start + 6, lic_mbps, nfmt)
            gap = (
                (lic_mbps - peak_mbps) / lic_mbps
                if lic_mbps > 0 and np.isfinite(peak_mbps)
                else float("nan")
            )
            if np.isfinite(gap):
                ws.write_number(row, right_start + 7, gap * 100.0, pfmt)
            else:
                ws.write_blank(row, right_start + 7, None, nfmt)
            ws.write_number(row, right_start + 8, stat["hours_on"], ifmt)
            ws.write_number(row, right_start + 9, stat["share"], pfmt)
            ws.write_string(row, right_start + 10, stat["cap_when"], cfmt)
        else:
            for c in range(right_start + 6, right_start + 11):
                ws.write_blank(row, c, None, cfmt)
        for offset, key in enumerate(OSS_VALUE_KEYS):
            value = _finite(busy.get(key))
            if np.isfinite(value):
                ws.write_number(row, oss_start + offset, value, nfmt)
            else:
                ws.write_blank(row, oss_start + offset, None, nfmt)
    last = header_row + 1 + max(len(charts), 1)
    ws.autofilter(header_row + 1, 0, last, col - 1)
    ws.repeat_rows(header_row, header_row + 1)
    ws.set_zoom(100)


def _oss_side_formats(book) -> dict:
    def _fmt(bg: str, *, num: bool = False, hour: bool = False):
        spec = {
            "font_name": "Calibri",
            "bold": True,
            "font_size": 9 if not hour else 10,
            "font_color": "white",
            "align": "center" if hour or num else "left",
            "valign": "vcenter",
            "bg_color": bg,
            "border": 1,
            "border_color": "white",
            "text_wrap": bool(hour),
        }
        if num:
            spec["num_format"] = "0.00"
        return book.add_format(spec)

    return {
        "head": book.add_format(
            {
                "font_name": "Calibri",
                "bold": True,
                "font_size": 10,
                "font_color": "white",
                "align": "center",
                "valign": "vcenter",
                "bg_color": fcc.NAVY,
                "border": 1,
                "border_color": "white",
            }
        ),
        "data4": _fmt("#C65911"),
        "data4_num": _fmt("#C65911", num=True),
        "data4_hour": _fmt("#C65911", hour=True),
        "user4": _fmt("#B7950B"),
        "user4_num": _fmt("#B7950B", num=True),
        "user4_hour": _fmt("#B7950B", hour=True),
        "data1": _fmt("#6C3483"),
        "data1_num": _fmt("#6C3483", num=True),
        "data1_hour": _fmt("#6C3483", hour=True),
        "user1": _fmt("#1A5276"),
        "user1_num": _fmt("#1A5276", num=True),
        "user1_hour": _fmt("#1A5276", hour=True),
    }


def write_issue_snapshots(
    book,
    styles,
    work,
    charts: list[dict],
    chart_start,
    chart_end,
    busy_by_code: dict[str, dict] | None = None,
) -> None:
    """One chart per issue site. Every chart has the same five lines."""
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
    for col in range(8, 12):
        ws.set_column(col, col, 12)
    ws.set_column(12, 12, 20)
    ws.set_column(13, 13, 28)
    ws.set_column(14, 14, 12)
    side_formats = _oss_side_formats(book)
    busy_by_code = busy_by_code or {}
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
        lines = item["lines"]
        peak = float(lines["peak"])
        touch_level = float(lines["level"])
        license_mbps = float(item.get("license_mbps", float("nan")))
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
        if item.get("category") == CAT_TX and touch:
            tags.append(
                f"{CAT_TX} {item['display_severity']} · gap {float(touch['gap']) * 100:.1f}% · {TX_COMMENT}"
            )
        elif item.get("category") == CAT_PEAK and touch:
            tags.append(f"{CAT_PEAK} {item['display_severity']}")
        elif item.get("category") == CAT_LICENSE:
            tags.append(f"{CAT_LICENSE} {item.get('display_severity', '')}")
        if lic and item.get("category") != CAT_LICENSE:
            tags.append(f"License {lic['share']:.1f}% of hours ≥{use_pct:.0f}%")
        ws.set_row(top, 26)
        ws.merge_range(
            top, 0, top, 7,
            f"{i + 1:02d}    {site}    ·    {REPORT_PERIOD}    ·    {'    ·    '.join(tags)}",
            styles["banner"],
        )
        labels = [
            (0, "Peak (Mbit/s)", f"{peak:.2f}" if np.isfinite(peak) else "—"),
            (2, f"{TOUCH_LEVEL_NAME} (Mbit/s)", f"{touch_level:.2f}" if np.isfinite(touch_level) else "—"),
            (4, "License (Mbps)", f"{license_mbps:.2f}" if np.isfinite(license_mbps) else "No license"),
            (6, f"{use_pct:.0f}% of license (Mbps)", f"{threshold:.2f}" if np.isfinite(threshold) else "—"),
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
        series = [
            (f"{TRACE_LABEL} (Mbps)", 2, {"color": "#5B9BD5", "width": 2.25}, np.isfinite(peak) or True),
            (f"Peak ({peak:.2f} Mbit/s)", 3, {"color": "#C00000", "width": 1.5, "dash_type": "dash"}, np.isfinite(peak)),
            (f"{TOUCH_LEVEL_NAME} ({touch_level:.2f} Mbit/s)", 4, {"color": "#ED7D31", "width": 1.5, "dash_type": "dash"}, np.isfinite(touch_level)),
            (f"License ({license_mbps:.2f} Mbps)", 5, {"color": "#548235", "width": 1.5, "dash_type": "dash"}, np.isfinite(license_mbps)),
            (f"{use_pct:.0f}% of license ({threshold:.2f} Mbps)", 6, {"color": "#833C0C", "width": 1.5, "dash_type": "dash"}, np.isfinite(threshold)),
        ]
        series = [(name, offset, line) for name, offset, line, ok in series if ok]
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
        fcc._chart_side_table(ws, top, busy_by_code.get(item.get("site_code", ""), {}), side_formats)
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
    ws.set_row(0, 28)
    ws.merge_range(0, 0, 0, len(headers) - 1, "Hourly KPI — issue sites", styles["title"])
    ws.set_row(1, 18)
    ws.merge_range(
        1, 0, 1, len(headers) - 1,
        f"{REPORT_PERIOD}    ·    hourly    ·    every issue site",
        styles["subtitle"],
    )
    header_row = 2
    ws.set_row(header_row, 30)
    for col, text in enumerate(headers):
        ws.write(header_row, col, text, styles["header"])

    order = {item["site"]: i for i, item in enumerate(charts)}
    license_of = {
        item["site"]: item["license_mbps"]
        for item in charts
        if np.isfinite(item.get("license_mbps", float("nan")))
    }
    hourly = work.loc[work["eNodeB Name"].isin(order)].copy()
    hourly["_order"] = hourly["eNodeB Name"].map(order)
    hourly = hourly.sort_values(["_order", "ts"])
    parity = {item["site"]: i % 2 for i, item in enumerate(charts)}
    excel_row = header_row
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
    ws.autofilter(header_row, 0, max(excel_row, header_row + 1), len(headers) - 1)
    ws.repeat_rows(header_row, header_row)
    ws.freeze_panes(header_row + 1, 0)
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
    fcc._page = _page
    styles = fcc._styles(book)
    period_start = work["Date"].min()
    period_end = work["Date"].max()
    n_sites = int(work["eNodeB Name"].nunique())
    n_scored = int(work.loc[work["Peak"].notna(), "eNodeB Name"].nunique())
    period_txt = f"{period_start.strftime('%d %b %Y')} – {period_end.strftime('%d %b %Y')}"
    global REPORT_PERIOD
    REPORT_PERIOD = period_txt
    fcc.REPORT_PERIOD = period_txt
    chart_end = period_end.normalize()
    chart_start = chart_end - pd.Timedelta(days=fcc.SNAP_DAYS - 1)
    license_records = license_summary["records"] if license_summary else []
    charts = issue_charts(records, license_records)
    # License results sit on the Dashboard, under the Peak-touch section.
    # One chart per site from either list, and both tables link to it.
    stats = license_summary.get("stats_by_site", {}) if license_summary else {}
    attach_chart_lines(work, charts, stats)
    apply_tx(records, stats)
    stamp_charts(charts, records)
    chart_index = {item["site"]: i for i, item in enumerate(charts)}
    n_tx = sum(1 for rec in records if rec["category"] == CAT_TX)
    print(f"{CAT_TX} (gap ≥ {GAP_MIN * 100:.0f}%): {n_tx}  ·  charts {len(charts)}", flush=True)
    geo, tech_by_code, type_by_code = load_physical()
    attach_site_identity(charts, tech_by_code, type_by_code)
    busy_rows, missing_feg, oss_note = prepare_oss(work, charts, tech_by_code, type_by_code)
    busy_by_code = {row["site"]: row for row in busy_rows}
    write_dashboard(
        book, styles, source_name, period_txt, n_sites, n_scored, records,
        link_charts=True,
        chart_index=chart_index,
        license_summary=license_summary,
        oss_note=oss_note,
    )
    write_site_list(book, styles, period_txt, charts, busy_by_code)
    if license_summary is not None:
        write_issue_snapshots(
            book, styles, work, charts, chart_start, chart_end, busy_by_code
        )
        print("Writing hourly KPI...", flush=True)
        write_issue_hourly(book, styles, work, charts)
    else:
        write_snapshots(book, styles, work, records, chart_start, chart_end)
        print("Writing hourly KPI...", flush=True)
        write_hourly(book, styles, work, records)
    geo_records, license_only_sites = geo_issue_records(charts)
    print(
        f"GeoPlot: {len(geo_records)} physical sites from {len(charts)} ERBS "
        f"({license_only_sites} license-only, drawn as Low)",
        flush=True,
    )
    fcc.GEO_FILE = GEO_FILE
    fcc._write_geoplot(book, styles, work, geo_records, geo, path, tech_by_code)
    geo_ws = book.get_worksheet_by_name(fcc.GEO_SHEET)
    geo_ws.set_row(4, 32)
    geo_ws.merge_range(
        4, 0, 4, 17,
        "Each dot is one physical site (the first 7 characters of the ERBS). "
        "Severity is the worst Peak-touch grade at that site. Medium is drawn as Moderate. "
        "Map severity is the site's category grade. Medium is drawn as Moderate. Decision is the 10% to under 15% gap. "
        f"A site listed only by the license rule is drawn as Low ({license_only_sites} sites).",
        styles["note"],
    )
    link_records = [{"site": item["site_code"]} for item in charts]
    print(f"Writing BusyHour ({len(busy_rows):,} sites)...", flush=True)
    fcc._write_busy_hour_plan(book, styles, period_txt, link_records, busy_rows)
    fcc._write_missing_feg_port(book, styles, missing_feg)
    immediate_records = []
    immediate_busy = []
    for item in charts:
        touch = item["touch"]
        rec = {
            "site": item["site"],
            "severity": (touch.get("touch_severity") or touch["severity"]) if touch else "—",
            "hours_pct": float(touch["hours_pct"]) if touch else 0.0,
            "tech": item.get("tech") or fcc.TECH_4G,
            "site_type": item.get("site_type") or "Not found",
        }
        immediate_records.append(rec)
        src = busy_by_code.get(item["site_code"])
        if src:
            copy = dict(src)
            copy.update(rec)
            immediate_busy.append(copy)
    fcc._write_immediate_actions(book, styles, immediate_records, immediate_busy)
    book.close()
    for jpeg in path.parent.glob(path.stem + "_GeoPlot_*.jpg"):
        jpeg.unlink()


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
        engine = "pyxlsb" if license_path.suffix.lower() == ".xlsb" else None
        lic = pd.read_excel(license_path, engine=engine) if engine else pd.read_excel(license_path)
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
    print(f"Listed (peak ratio ≥ 10%): {len(records)}")
    for cat in (CAT_PEAK, CAT_TX):
        print(f"  {cat}:")
        names = [name for name, _floor in SEVERITY_BANDS] if cat == CAT_PEAK else [name for name, _f, _r in TX_BANDS]
        for name in names:
            print(f"    {name}: {sum(1 for rec in records if rec['category'] == cat and rec['severity'] == name)}")
    print("Where touched:")
    for name in fcc.CAP_WHEN_COLS:
        print(f"  {name}: {by_when[name]}")
    print(f"Wrote {output} in {pytime.perf_counter() - t0:.1f}s")


if __name__ == "__main__":
    main()
