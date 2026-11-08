#!/usr/bin/env python3
"""Add the vendor EthernetPort counter to the current Ericsson workbook.

The 30 Oct–1 Nov radio sheets stay as they are. This reads `EthernetPort File_*.rar`
(5–7 Nov 2026) and writes them in the same shape as the Site Level dashboard:

* a Front Node Site section at the bottom of Dashboard
* sheet FrontNode_Site
* one chart per listed site, appended on 2. HourlyChartOfIssueSites
* the chart page FrontNode tab
* GeoPlot HTML and hidden _GeoData, with the same worst-grade rule

Site Name = the first 7 characters of NE Name. Each site-hour is the MAX of
In Throughput (Mbps) across every row for that site and hour. Front Node is
the NE whose own maximum is the highest. The scored trace is that hourly maximum.
"""

from __future__ import annotations

import re
import subprocess
import sys
import zipfile
from pathlib import Path
from xml.etree import ElementTree as ET

import numpy as np
import pandas as pd

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(Path(__file__).resolve().parent))

import fege_choke_check as fcc  # noqa: E402
import fege_peak_touch_ericsson as pt  # noqa: E402
import fege_site_charts_html as charts_html  # noqa: E402
from geoplot_satellite import (  # noqa: E402
    classify_rows,
    cluster_bounds,
    cluster_issue_sites,
    cluster_label,
    render_html,
    thana_focus_specs,
)

BASE_VERSION = "v39"
REPORT_VERSION = "v41"
BASE_BOOK = ROOT / f"FEGE_Choked_Flat_Sites_Ericsson_{BASE_VERSION}.xlsx"
OUT_BOOK = ROOT / f"FEGE_Choked_Flat_Sites_Ericsson_{REPORT_VERSION}.xlsx"
CHARTS_HTML = ROOT / f"FEGE_Choked_Flat_Sites_Ericsson_{REPORT_VERSION}_Charts.html"
CHARTS_ZIP = ROOT / f"FEGE_Choked_Flat_Sites_Ericsson_{REPORT_VERSION}_Charts.zip"
GEOPLOT_HTML = ROOT / f"FEGE_Choked_Flat_Sites_Ericsson_{REPORT_VERSION}_GeoPlot.html"
EXTRACT_DIR = Path("/tmp/ethernet_frontnode")
NS = {"m": "http://schemas.openxmlformats.org/spreadsheetml/2006/main"}
GEO_RANK = {"Severe": 0, "High": 1, "Moderate": 2, "Low": 3, "Decision": 4}
CHART_COLS = 7
PERIOD = "05 Nov 2026 – 07 Nov 2026"
SNAP_SHEET = "2. HourlyChartOfIssueSites"
S_CHART_BANNER, S_CHART_LABEL, S_CHART_VALUE, S_CHART_TAG, S_CHART_AXIS = "49", "50", "51", "52", "70"

S_TITLE, S_SUB, S_NOTE, S_BANNER = "1", "2", "3", "4"
S_TILE_L, S_TILE_A, S_TILE_LV, S_TILE_AV = "5", "6", "7", "8"
S_GROUP, S_HEADER, S_CAT = "9", "10", "11"
S_INT, S_INT_Z = "13", "28"
S_COMMENT, S_RULE, S_TOTAL = "14", "15", "21"
S_SECTION, S_BODY = "23", "24"
S_LINK, S_LINK_Z = "25", "29"
S_NUM, S_NUM_Z = "26", "30"
S_PCT, S_PCT_Z = "27", "31"
S_CENTER, S_CENTER_Z = "17", "32"
S_YES = "46"
S_SEV = {"Severe": "12", "High": "16", "Medium": "18", "Low": "19", "Decision": "20"}
S_METHOD_A, S_METHOD_B = "76", "77"


def col_name(index: int) -> str:
    name = ""
    n = index + 1
    while n:
        n, rem = divmod(n - 1, 26)
        name = chr(65 + rem) + name
    return name


def xml_text(value: object) -> str:
    return (
        str(value)
        .replace("&", "&amp;")
        .replace("<", "&lt;")
        .replace(">", "&gt;")
        .replace('"', "&quot;")
    )


def c_str(ref: str, style: str, text: object) -> str:
    return (
        f'<c r="{ref}" s="{style}" t="inlineStr"><is><t xml:space="preserve">'
        f"{xml_text(text)}</t></is></c>"
    )


def c_num(ref: str, style: str, value: float | int | None) -> str:
    if value is None:
        return f'<c r="{ref}" s="{style}"/>'
    number = float(value)
    if not np.isfinite(number):
        return f'<c r="{ref}" s="{style}"/>'
    return f'<c r="{ref}" s="{style}"><v>{number:.10g}</v></c>'


def c_blank(ref: str, style: str) -> str:
    return f'<c r="{ref}" s="{style}"/>'


def row_xml(r: int, height: float | None, cells: list[str], last_col: int) -> str:
    ht = f' ht="{height}" customHeight="1"' if height else ""
    return f'<row r="{r}" spans="1:{last_col + 1}"{ht}>{"".join(cells)}</row>'


def span(r: int, height: float, style: str, text: str, last_col: int) -> tuple[str, str]:
    cells = [c_str(f"A{r}", style, text)]
    cells.extend(c_blank(f"{col_name(col)}{r}", style) for col in range(1, last_col + 1))
    return row_xml(r, height, cells, last_col), f"A{r}:{col_name(last_col)}{r}"


def finite(value: object) -> float | None:
    try:
        number = float(value)  # type: ignore[arg-type]
    except (TypeError, ValueError):
        return None
    if not np.isfinite(number):
        return None
    return number


def extract_csvs() -> list[Path]:
    EXTRACT_DIR.mkdir(parents=True, exist_ok=True)
    archives = sorted(ROOT.glob("EthernetPort File_*.rar"))
    if not archives:
        raise SystemExit("No EthernetPort File_*.rar files in the repo")
    paths = []
    for archive in archives:
        dest = EXTRACT_DIR / (archive.stem + ".csv")
        if not dest.exists() or dest.stat().st_size == 0:
            subprocess.check_call(["unrar", "x", "-o+", str(archive), str(EXTRACT_DIR)])
            extracted = EXTRACT_DIR / archive.with_suffix(".csv").name
            if extracted != dest and extracted.exists():
                extracted.replace(dest)
        if not dest.exists():
            found = list(EXTRACT_DIR.glob(archive.stem + ".csv"))
            if not found:
                raise SystemExit(f"Could not extract {archive.name}")
            dest = found[0]
        paths.append(dest)
    return paths


def load_ethernet() -> tuple[pd.DataFrame, dict[str, tuple[str, float]], int]:
    """Hourly site maximum, and the front node of each site.

    Site Name is the first 7 characters of NE Name. The site-hour value is the
    maximum In Throughput (Mbps) among every row for that site and hour.
    """
    frames = []
    for path in extract_csvs():
        frames.append(
            pd.read_csv(
                path,
                usecols=["Date", "Hour", "NE Name", "In Throughput (Mbps)"],
            )
        )
    raw = pd.concat(frames, ignore_index=True)
    raw["NE Name"] = raw["NE Name"].astype(str).str.strip()
    raw["Site Name"] = raw["NE Name"].str[:7]
    raw["Hour"] = pd.to_numeric(raw["Hour"], errors="coerce")
    raw["Rx"] = pd.to_numeric(raw["In Throughput (Mbps)"], errors="coerce")
    raw["Date"] = fcc.parse_kpi_date(raw["Date"])
    raw = raw.dropna(subset=["Date", "Hour", "Rx"])
    raw = raw.loc[~raw["Site Name"].str.lower().isin(["", "nan", "none"])]
    raw["Hour"] = raw["Hour"].astype(int)
    sample = raw.loc[
        (raw["Site Name"] == "BBAKD01")
        & (raw["Date"].dt.month == 11)
        & (raw["Date"].dt.day == 5)
        & (raw["Hour"] == 0),
        "Rx",
    ]
    if sample.empty or abs(float(sample.max()) - 979.9746635) > 1e-6:
        raise SystemExit(f"BBAKD01 5 Nov hour 0 max is {sample.max() if len(sample) else 'missing'}")
    print(f"  BBAKD01 5 Nov 00:00 max = {float(sample.max()):.7f} from {len(sample)} rows", flush=True)
    ne_peak = raw.groupby(["Site Name", "NE Name"], as_index=False)["Rx"].max()
    front: dict[str, tuple[str, float]] = {}
    for site, group in ne_peak.groupby("Site Name", sort=False):
        group = group.sort_values(["Rx", "NE Name"], ascending=[False, True])
        top = group.iloc[0]
        front[str(site)] = (str(top["NE Name"]), float(top["Rx"]))
    site_hour = raw.groupby(["Date", "Hour", "Site Name"], as_index=False)["Rx"].max()
    site_hour = site_hour.rename(columns={"Site Name": pt.eric.ERIC_SITE_COL, "Rx": pt.TRACE_COL})
    site_hour["Date"] = pd.to_datetime(site_hour["Date"]).dt.strftime("%-d-%b")
    return site_hour, front, int(raw["Site Name"].nunique())


def score(site_hour: pd.DataFrame) -> tuple[pd.DataFrame, list[dict], dict, list[dict]]:
    work, records = pt.analyse_touch(site_hour)
    lic = pd.read_excel(ROOT / pt.LICENSE_FILE, engine="pyxlsb")
    site_lic = pt.aggregate_license_to_site(lic)
    summary = pt.analyse_license(site_hour, site_lic)
    charts = pt._prepare_charts(work, records, summary, pt.LEVEL_SITE)
    return work, records, summary, charts


def load_bh_flags() -> dict[str, list[str | None]]:
    """BH Flag 1H from the OSS 1 BusyHour blocks already on URGENTBWINC_BusyHour.

    Yes when either Data 1H or User 1H matches. No when both values exist and
    neither matches. Blank when both 1H values are missing.
    """
    from openpyxl import load_workbook

    book = load_workbook(BASE_BOOK, read_only=True, data_only=True)
    ws = book["URGENTBWINC_BusyHour"]
    flags: dict[str, list[str | None]] = {}
    for row in ws.iter_rows(min_row=6, values_only=True):
        site = str(row[1] or "").strip()
        if not site:
            continue
        tp = (pt._finite(row[60]), pt._finite(row[72]))
        vol = (pt._finite(row[62]), pt._finite(row[74]))
        flags[site] = [
            pt._bh_yes(tp, lambda value: value <= 7.0),
            pt._bh_yes(vol, lambda value: value >= 50.0),
            pt._bh_yes(tp, lambda value: value <= 5.0),
            pt._bh_yes(vol, lambda value: value >= 100.0),
        ]
    book.close()
    return flags


def load_bands() -> dict[str, str]:
    from openpyxl import load_workbook

    book = load_workbook(BASE_BOOK, read_only=True, data_only=True)
    ws = book["1. Site List_SiteLevel"]
    bands: dict[str, str] = {}
    name_i = band_i = None
    for row in ws.iter_rows(values_only=True):
        if not row:
            continue
        if row[0] == "No.":
            labels = [str(cell).strip() if cell is not None else "" for cell in row]
            name_i = labels.index("Site Name")
            band_i = labels.index("Band")
            continue
        if name_i is None or not isinstance(row[0], (int, float)) or not row[name_i]:
            continue
        band = str(row[band_i] or "").strip()
        bands[str(row[name_i]).strip()] = band or "—"
    book.close()
    return bands


def shared_strings(blob: bytes) -> list[str]:
    root = ET.fromstring(blob)
    tag = "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}t"
    return ["".join(node.text or "" for node in si.iter(tag)) for si in root]


def read_geodata(blob: bytes, strings: list[str]) -> list[dict]:
    root = ET.fromstring(blob)
    rows = []
    header = None
    for row in root.find("m:sheetData", NS):
        values = {}
        for cell in row:
            ref = cell.attrib.get("r", "")
            col = "".join(ch for ch in ref if ch.isalpha())
            node = cell.find("m:v", NS)
            if node is None or node.text is None:
                values[col] = ""
                continue
            if cell.attrib.get("t") == "s":
                values[col] = strings[int(node.text)]
            else:
                values[col] = node.text
        if header is None:
            header = values
            continue
        rows.append(
            {
                "site": values.get("A", "").strip(),
                "lat": float(values["B"]) if values.get("B") else None,
                "lon": float(values["C"]) if values.get("C") else None,
                "status": values.get("D", ""),
                "tech": values.get("E", ""),
                "issue_tech": values.get("F", ""),
                "district": values.get("G", ""),
                "region": values.get("H", ""),
                "thana": values.get("I", ""),
                "severity": values.get("J", ""),
            }
        )
    return [row for row in rows if row["site"]]


def worse(left: str, right: str) -> str:
    if left not in GEO_RANK:
        return right
    if right not in GEO_RANK:
        return left
    return left if GEO_RANK[left] <= GEO_RANK[right] else right


def sheet_document(rows: list[str], merges: list[str], links: list[str], *, tab: str, dimension: str, cols: str, header: str, footer: str, freeze: int | None) -> str:
    merge_xml = ""
    if merges:
        merge_xml = f'<mergeCells count="{len(merges)}">' + "".join(
            f'<mergeCell ref="{ref}"/>' for ref in merges
        ) + "</mergeCells>"
    link_xml = ""
    if links:
        link_xml = "<hyperlinks>" + "".join(links) + "</hyperlinks>"
    view = '<sheetView showGridLines="0" workbookViewId="0">'
    if freeze:
        view += (
            f'<pane ySplit="{freeze}" topLeftCell="A{freeze + 1}" activePane="bottomLeft" state="frozen"/>'
            f'<selection pane="bottomLeft" activeCell="A{freeze + 1}" sqref="A{freeze + 1}"/>'
        )
    view += "</sheetView>"
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
        'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
        f'<sheetPr><tabColor rgb="{tab}"/><pageSetUpPr fitToPage="1"/></sheetPr>'
        f'<dimension ref="{dimension}"/>'
        f"<sheetViews>{view}</sheetViews>"
        '<sheetFormatPr defaultRowHeight="15"/>'
        f"{cols}<sheetData>{''.join(rows)}</sheetData>{merge_xml}{link_xml}"
        '<pageMargins left="0.45" right="0.45" top="0.6" bottom="0.5" header="0.3" footer="0.3"/>'
        '<pageSetup paperSize="9" fitToHeight="0" orientation="landscape"/>'
        f"<headerFooter><oddHeader>{header}</oddHeader><oddFooter>{footer}</oddFooter></headerFooter>"
        "</worksheet>"
    )


def rels_xml(pairs: list[tuple[str, str]]) -> str:
    body = "".join(
        f'<Relationship Id="{rid}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" '
        f'Target="{xml_text(target)}" TargetMode="External"/>'
        for rid, target in pairs
    )
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        f"{body}</Relationships>"
    )


def _header_row(r: int, labels: list[str], height: float = 30) -> str:
    return row_xml(r, height, [c_str(f"{col_name(i)}{r}", S_HEADER, text) for i, text in enumerate(labels)], len(labels) - 1)


def dashboard_block(start: int, records, summary, charts, n_sites: int, chart_rows: dict[str, int]):
    chart_index = {item["site"]: i for i, item in enumerate(charts)}
    rows: list[str] = []
    merges: list[str] = []
    links: list[str] = []
    r = start

    def add_span(height, style, text, last_col):
        nonlocal r
        xml, ref = span(r, height, style, text, last_col)
        rows.append(xml)
        merges.append(ref)
        r += 1
        return r - 1

    def add(xml):
        nonlocal r
        rows.append(xml)
        r += 1

    add_span(28, S_BANNER, "Front Node Site Dashboard", 11)
    add_span(
        120,
        S_NOTE,
        "Vendor counter In Throughput (Mbps), from EthernetPort File_, window 5 Nov 2026 – 7 Nov 2026. "
        "The radio sheets above stay on 30 Oct 2026 – 1 Nov 2026. "
        "Site Name is the first 7 characters of NE Name. "
        "Each hour, Tx usage is the maximum In Throughput (Mbps) among every row for that site and hour, "
        "every NE and every EthernetPort. "
        "BBAKD01 at 00:00 on 5 Nov is 979.9746635, the highest of its 14 rows. "
        "License is the SUM of dlBbCapacityNet for every NodeId with that Site Name. UL is not used. "
        "Front Node is the NE Name whose own maximum In Throughput is the highest. "
        "The scored trace is that hourly maximum, with the same floors as Site Level. "
        "Peak is the highest hourly maximum. Below5%Peak is Peak × 95%. "
        "An hour touches when the hourly maximum is at or above Below5%Peak. "
        "Tx Flat/Chocked, Peak Touch, and License usage use the same floors as the Site Level Dashboard. "
        f"A site is scored when it has at least {pt.MIN_SAMPLES} hourly samples and a Peak of at least {pt.MIN_PEAK_MBPS:.0f} Mbit/s. "
        "The site name opens that chart on 2. HourlyChartOfIssueSites. "
        "GeoPlot HTML and _GeoData include these sites.",
        11,
    )
    license_records = summary["records"]
    n_peak = sum(1 for rec in records if rec["category"] == pt.CAT_PEAK)
    n_tx = sum(1 for rec in records if rec["category"] == pt.CAT_TX)
    label_cells = []
    value_cells = []
    for col, label, value, alert in (
        (0, "Sites checked", n_sites, False),
        (2, pt.CAT_PEAK, n_peak, False),
        (4, pt.CAT_TX, n_tx, True),
        (6, pt.CAT_LICENSE, len(license_records), True),
    ):
        label_style = S_TILE_A if alert else S_TILE_L
        value_style = S_TILE_AV if alert else S_TILE_LV
        label_cells += [c_str(f"{col_name(col)}{r}", label_style, label), c_blank(f"{col_name(col + 1)}{r}", label_style)]
        value_cells += [c_num(f"{col_name(col)}{r + 1}", value_style, value), c_blank(f"{col_name(col + 1)}{r + 1}", value_style)]
        merges.append(f"{col_name(col)}{r}:{col_name(col + 1)}{r}")
        merges.append(f"{col_name(col)}{r + 1}:{col_name(col + 1)}{r + 1}")
    add(row_xml(r, 18, label_cells, 7))
    add(row_xml(r, 32, value_cells, 7))
    r += 1

    add_span(22, S_GROUP, "Category summary", 7)
    add(_header_row(r, ["Category", "Severity", *fcc.CAP_WHEN_COLS, "Total", "Comments", "Rule Details"]))
    blocks = []
    tx_rows = []
    for name, _floor, rule in pt.TX_BANDS:
        chosen = [rec for rec in records if rec["category"] == pt.CAT_TX and rec["severity"] == name]
        tx_rows.append((name, pt.TX_COMMENT, rule, pt.rec_when(chosen)))
    blocks.append((pt.CAT_TX, tx_rows))
    peak_rows = []
    for name, _floor in pt.SEVERITY_BANDS:
        chosen = [rec for rec in records if rec["category"] == pt.CAT_PEAK and rec["severity"] == name]
        peak_rows.append((name, "", pt.PEAK_RULES[name], pt.rec_when(chosen)))
    blocks.append((pt.CAT_PEAK, peak_rows))
    lic_rows = []
    for sev in (pt.SEV_SEVERE, pt.SEV_HIGH, pt.SEV_MEDIUM, pt.SEV_LOW):
        chosen = [rec for rec in license_records if pt.license_type(rec["share"]) == sev]
        lic_rows.append((sev, "", pt.LICENSE_RULES[sev], pt.rec_when(chosen)))
    blocks.append((pt.CAT_LICENSE, lic_rows))
    grand = {when: 0 for when in fcc.CAP_WHEN_COLS}
    for label, block_rows in blocks:
        start_row = r
        for sev, comment, rule, when_counts in block_rows:
            cells = [c_blank(f"A{r}", S_CAT), c_str(f"B{r}", S_SEV[sev], sev)]
            total = 0
            for offset, when in enumerate(fcc.CAP_WHEN_COLS):
                value = when_counts.get(when, 0)
                total += value
                grand[when] += value
                cells.append(c_num(f"{col_name(2 + offset)}{r}", S_INT, value or None))
            cells.append(c_num(f"F{r}", S_INT, total))
            cells.append(c_str(f"G{r}", S_COMMENT, comment) if comment else c_blank(f"G{r}", S_CENTER))
            cells.append(c_str(f"H{r}", S_RULE, rule))
            add(row_xml(r, 22, cells, 7))
        end_row = r - 1
        if end_row == start_row:
            rows[-1] = rows[-1].replace(f'<c r="A{start_row}" s="{S_CAT}"/>', c_str(f"A{start_row}", S_CAT, label), 1)
        else:
            rows[len(rows) - (end_row - start_row + 1)] = rows[len(rows) - (end_row - start_row + 1)].replace(
                f'<c r="A{start_row}" s="{S_CAT}"/>', c_str(f"A{start_row}", S_CAT, label), 1
            )
            merges.append(f"A{start_row}:A{end_row}")
    total_cells = [
        c_str(f"A{r}", S_TOTAL, "Grand Total"),
        c_blank(f"B{r}", S_TOTAL),
    ]
    merges.append(f"A{r}:B{r}")
    grand_n = sum(grand.values())
    for offset, when in enumerate(fcc.CAP_WHEN_COLS):
        total_cells.append(c_num(f"{col_name(2 + offset)}{r}", S_TOTAL, grand[when]))
    total_cells += [
        c_num(f"F{r}", S_TOTAL, grand_n),
        c_blank(f"G{r}", S_TOTAL),
        c_str(f"H{r}", S_RULE, "Distinct findings added across the three categories. A site in two categories is counted in each."),
    ]
    add(row_xml(r, 18, total_cells, 7))
    add_span(
        32,
        S_NOTE,
        "Grand Total adds the rows above, so a site in License usage and also in Peak Touch or "
        "Tx Flat/Chocked is counted twice here. Sheet FrontNode_Site and the charts keep that site once.",
        7,
    )
    r += 1

    sections = []
    for name, _floor, rule in pt.TX_BANDS:
        group = [rec for rec in records if rec["category"] == pt.CAT_TX and rec["severity"] == name]
        sections.append((f"{pt.CAT_TX} · {name}", f"{rule}. {pt.TX_COMMENT}.", group))
    for name, _floor in pt.SEVERITY_BANDS:
        group = [rec for rec in records if rec["category"] == pt.CAT_PEAK and rec["severity"] == name]
        sections.append((f"{pt.CAT_PEAK} · {name}", pt.PEAK_RULES[name], group))
    detail_headers = [
        "No.", "Site Name", "Severity", "Peak (Mbit/s)", "Below5%Peak (Mbit/s)",
        "Touch hours", "Total hours", "Peak ratio (%)", "Busy-hour touch (%)",
        "Where touched", "License (Mbps)", "Gap (%)",
    ]
    for title, note, group in sections:
        if not group:
            continue
        noun = "site" if len(group) == 1 else "sites"
        add_span(22, S_SECTION, f"{title}    ·    {len(group)} {noun}", 11)
        add_span(24, S_BODY, note, 11)
        add(_header_row(r, detail_headers))
        for n, rec in enumerate(group):
            zebra = n % 2 == 1
            nfmt = S_NUM_Z if zebra else S_NUM
            ifmt = S_INT_Z if zebra else S_INT
            cfmt = S_CENTER_Z if zebra else S_CENTER
            pfmt = S_PCT_Z if zebra else S_PCT
            link = S_LINK_Z if zebra else S_LINK
            chart_i = chart_index[rec["site"]]
            anchor = chart_rows[rec["site"]]
            cells = [
                c_num(f"A{r}", ifmt, chart_i + 1),
                c_str(f"B{r}", link, rec["site"]),
                c_str(f"C{r}", S_SEV[rec["severity"]], rec["severity"]),
                c_num(f"D{r}", nfmt, rec["peak"]),
                c_num(f"E{r}", nfmt, rec["level"]),
                c_num(f"F{r}", ifmt, rec["hours_on"]),
                c_num(f"G{r}", ifmt, rec["hours_n"]),
                c_num(f"H{r}", pfmt, rec["hours_pct"]),
                c_num(f"I{r}", pfmt, rec["busy_pct"]),
                c_str(f"J{r}", cfmt, rec["cap_when"]),
                c_num(f"K{r}", nfmt, finite(rec.get("license_mbps"))),
                c_num(f"L{r}", pfmt, finite(rec.get("gap")) * 100.0 if finite(rec.get("gap")) is not None else None),
            ]
            add(row_xml(r, 18, cells, 11))
            links.append(
                f'<hyperlink ref="B{r - 1}" location="\'{SNAP_SHEET}\'!A{anchor}" display="{xml_text(rec["site"])}"/>'
            )
        r += 1

    r = _license_block(r, rows, merges, links, summary, chart_index, chart_rows)
    return "".join(rows), merges, links, r - 1


def _license_block(r, rows, merges, links, summary, chart_index, chart_rows):
    records = summary["records"]
    use = f"{pt.LICENSE_USE * 100:.0f}"

    def add_span(height, style, text, last_col):
        nonlocal r
        xml, ref = span(r, height, style, text, last_col)
        rows.append(xml)
        merges.append(ref)
        r += 1

    def add(xml):
        nonlocal r
        rows.append(xml)
        r += 1

    add_span(22, S_SECTION, "License usage", 9)
    add_span(
        64,
        S_NOTE,
        "dlBbCapacityNet is summed for every NodeId that shares the Site Name. UL is not used. "
        f"An hour counts when the hourly maximum of In Throughput (Mbps) / that combined license is at least {use}%. "
        f"A site is listed when those hours are at least {pt.LICENSE_HOURS_PCT:.0f}% of its hours. "
        "License usage type is that share of hours: Severe at least 50%, High 40% to under 50%, "
        "Medium 30% to under 40%, Low 20% to under 30%. "
        f"{summary['no_license']:,} sites have no positive combined license and are not scored.",
        9,
    )
    label_cells = []
    value_cells = []
    tiles = (
        (0, "Sites in the EthernetPort file", summary["feg_sites"], False),
        (2, "Matched to a license", summary["matched"], False),
        (4, "License scored (> 0)", summary["scored"], False),
        (6, f"Listed (≥{pt.LICENSE_HOURS_PCT:.0f}% of hours at ≥{use}%)", len(records), True),
    )
    for col, label, value, alert in tiles:
        label_style = S_TILE_A if alert else S_TILE_L
        value_style = S_TILE_AV if alert else S_TILE_LV
        label_cells += [c_str(f"{col_name(col)}{r}", label_style, label), c_blank(f"{col_name(col + 1)}{r}", label_style)]
        value_cells += [c_num(f"{col_name(col)}{r + 1}", value_style, value), c_blank(f"{col_name(col + 1)}{r + 1}", value_style)]
        merges.append(f"{col_name(col)}{r}:{col_name(col + 1)}{r}")
        merges.append(f"{col_name(col)}{r + 1}:{col_name(col + 1)}{r + 1}")
    add(row_xml(r, 18, label_cells, 7))
    add(row_xml(r, 32, value_cells, 7))
    r += 1
    counts = {name: 0 for name, _floor, _style, _rule in pt.LICENSE_CATEGORIES}
    for rec in records:
        counts[pt.license_category(rec["share"])] += 1
    add_span(22, S_SECTION, "License usage category", 5)
    add(_header_row(r, ["Category", f"Share of hours at ≥{use}% of license", "", "", "Sites", "Type"]))
    merges.append(f"B{r - 1}:D{r - 1}")
    for name, _floor, style, rule in pt.LICENSE_CATEGORIES:
        sev = pt.LICENSE_TYPE[style]
        cells = [
            c_str(f"A{r}", S_SEV[sev], name),
            c_str(f"B{r}", S_RULE, rule),
            c_blank(f"C{r}", S_RULE),
            c_blank(f"D{r}", S_RULE),
            c_num(f"E{r}", S_INT, counts[name]),
            c_str(f"F{r}", S_SEV[sev], sev),
        ]
        merges.append(f"B{r}:D{r}")
        add(row_xml(r, 18, cells, 5))
    cells = [
        c_str(f"A{r}", S_HEADER, "Total listed"),
        c_str(f"B{r}", S_HEADER, f"At least {pt.LICENSE_HOURS_PCT:.0f}% of hours at ≥{use}% of license"),
        c_blank(f"C{r}", S_HEADER),
        c_blank(f"D{r}", S_HEADER),
        c_num(f"E{r}", S_INT, len(records)),
    ]
    merges.append(f"B{r}:D{r}")
    add(row_xml(r, 18, cells, 4))
    r += 1
    headers = [
        "No.", "Site Name", "License (Mbps)", f"Hours ≥{use}%", "Total hours",
        "Share of hours (%)", "Peak In Throughput (Mbps)", "Peak / license (%)",
        "Median usage (%)", f"Where ≥{use}%",
    ]
    for name, _floor, _style, _rule in pt.LICENSE_CATEGORIES:
        group = [rec for rec in records if pt.license_category(rec["share"]) == name]
        if not group:
            continue
        noun = "site" if len(group) == 1 else "sites"
        add_span(22, S_SECTION, f"{name} of hours at ≥{use}% of license    ·    {len(group)} {noun}", 9)
        add(_header_row(r, headers, 32))
        for n, rec in enumerate(group):
            zebra = n % 2 == 1
            nfmt = S_NUM_Z if zebra else S_NUM
            ifmt = S_INT_Z if zebra else S_INT
            cfmt = S_CENTER_Z if zebra else S_CENTER
            pfmt = S_PCT_Z if zebra else S_PCT
            link = S_LINK_Z if zebra else S_LINK
            chart_i = chart_index[rec["site"]]
            anchor = chart_rows[rec["site"]]
            cells = [
                c_num(f"A{r}", ifmt, chart_i + 1),
                c_str(f"B{r}", link, rec["site"]),
                c_num(f"C{r}", nfmt, rec["license"]),
                c_num(f"D{r}", ifmt, rec["hours_on"]),
                c_num(f"E{r}", ifmt, rec["hours_n"]),
                c_num(f"F{r}", pfmt, rec["share"]),
                c_num(f"G{r}", nfmt, rec["peak"]),
                c_num(f"H{r}", pfmt, rec["peak_ratio"]),
                c_num(f"I{r}", pfmt, rec["median_ratio"]),
                c_str(f"J{r}", cfmt, rec["cap_when"]),
            ]
            add(row_xml(r, 18, cells, 9))
            links.append(
                f'<hyperlink ref="B{r - 1}" location="\'{SNAP_SHEET}\'!A{anchor}" display="{xml_text(rec["site"])}"/>'
            )
        r += 1
    return r


def gap_pct(item: dict) -> float | None:
    stat = item.get("lic_stat")
    peak = finite((item.get("lines") or {}).get("peak"))
    if not stat or peak is None:
        return None
    lic = finite(stat.get("license"))
    if lic is None or lic <= 0:
        return None
    return (lic - peak) / lic * 100.0


def site_sheet(charts, front, bands, tech, types, chart_rows: dict[str, int], bh_flags: dict[str, list[str | None]]):
    rows: list[str] = []
    merges: list[str] = []
    links: list[str] = []
    rels: list[tuple[str, str]] = []
    r = 1
    intro = (
        "5 Nov 2026 – 7 Nov 2026. These rows are the sites on the Front Node Site Dashboard. "
        "Site Name is the first 7 characters of NE Name. "
        "Each hour is the maximum In Throughput (Mbps) among every row for that site and hour. "
        "License is the SUM of dlBbCapacityNet. The floors match Site Level. "
        "BH Flag 1H uses the OSS 1 BusyHour blocks for 30 Oct–1 Nov, the same test as Site Level. "
        "Yes when either Data 1H or User 1H matches. "
        "DL Throughput<7Mbps is at or below 7 Mbps. Data Vol>50GB is at least 50 GB. "
        "DL Throughput<5Mbps is at or below 5 Mbps. Data Vol>100GB is at least 100 GB. "
        "No when both 1H values exist and neither matches. Blank when both are missing. "
        "Band is copied from 1. Site List_SiteLevel when that site is listed there. "
        "The site name opens 2. HourlyChartOfIssueSites. Map opens the GeoPlot HTML on that site."
    )
    xml, ref = span(r, 28, S_TITLE, "Front Node Site — sites on the Front Node Site Dashboard", 24)
    rows.append(xml)
    merges.append(ref)
    r += 1
    xml, ref = span(r, 18, S_SUB, "EthernetPort File_    ·    5 Nov 2026 – 7 Nov 2026    ·    In Throughput (Mbps)", 24)
    rows.append(xml)
    merges.append(ref)
    r += 1
    xml, ref = span(r, 78, S_NOTE, intro, 24)
    rows.append(xml)
    merges.append(ref)
    r += 1
    headers = [
        "No.", "Site Name", "Tech", "Site Type", "Band", "Category", "Severity",
        "DL Throughput<7Mbps", "Data Vol>50GB", "DL Throughput<5Mbps", "Data Vol>100GB",
        "Peak (Mbit/s)", "Below5%Peak (Mbit/s)", "Touch hours", "Total hours", "Peak ratio (%)",
        "Where touched", "Gap (%)", "Front Node", "Front Node peak (Mbit/s)", "Map",
    ]
    rows.append(_header_row(r, headers))
    header_row = r
    r += 1
    first_data = r
    for i, item in enumerate(charts):
        zebra = i % 2 == 1
        nfmt = S_NUM_Z if zebra else S_NUM
        ifmt = S_INT_Z if zebra else S_INT
        cfmt = S_CENTER_Z if zebra else S_CENTER
        pfmt = S_PCT_Z if zebra else S_PCT
        link = S_LINK_Z if zebra else S_LINK
        lines = item["lines"]
        touch = item.get("touch")
        severity = item.get("display_severity") or "—"
        where = touch["cap_when"] if touch else lines["cap_when"]
        code = pt.site_code(item["site"])
        anchor = chart_rows[item["site"]]
        front_ne, front_peak = front.get(item["site"], ("", float("nan")))
        cells = [
            c_num(f"A{r}", ifmt, i + 1),
            c_str(f"B{r}", link, item["site"]),
            c_str(f"C{r}", cfmt, tech.get(code, fcc.TECH_4G)),
            c_str(f"D{r}", cfmt, types.get(code) or "Not found"),
            c_str(f"E{r}", cfmt, bands.get(item["site"]) or "—"),
            c_str(f"F{r}", cfmt, item.get("category") or ""),
            c_str(f"G{r}", S_SEV.get(severity, cfmt), severity),
        ]
        for col, flag in enumerate(bh_flags.get(item["site"], [None, None, None, None]), start=7):
            if flag == "Yes":
                cells.append(c_str(f"{col_name(col)}{r}", S_YES, "Yes"))
            elif flag == "No":
                cells.append(c_str(f"{col_name(col)}{r}", cfmt, "No"))
            else:
                cells.append(c_blank(f"{col_name(col)}{r}", cfmt))
        cells += [
            c_num(f"L{r}", nfmt, lines["peak"]),
            c_num(f"M{r}", nfmt, lines["level"]),
            c_num(f"N{r}", ifmt, lines["hours_on"]),
            c_num(f"O{r}", ifmt, lines["hours_n"]),
            c_num(f"P{r}", pfmt, lines["hours_pct"]),
            c_str(f"Q{r}", cfmt, where),
            c_num(f"R{r}", pfmt, gap_pct(item)),
            c_str(f"S{r}", cfmt, front_ne),
            c_num(f"T{r}", nfmt, front_peak),
            c_str(f"U{r}", link, "Map"),
        ]
        rows.append(row_xml(r, 18, cells, 20))
        links.append(
            f'<hyperlink ref="B{r}" location="\'{SNAP_SHEET}\'!A{anchor}" display="{xml_text(item["site"])}"/>'
        )
        rid = f"rId{len(rels) + 1}"
        rels.append((rid, f"{GEOPLOT_HTML.name}#{item['site']}"))
        links.append(f'<hyperlink ref="U{r}" r:id="{rid}" display="Map"/>')
        r += 1
    last = r - 1
    cols = "".join(
        f'<col min="{i}" max="{i}" width="{width}" customWidth="1"/>'
        for i, width in enumerate(
            [8, 16, 12, 14, 22, 18, 14, 20, 16, 20, 16, 16, 20, 14, 14, 16, 18, 12, 18, 24, 12],
            start=1,
        )
    )
    xml = sheet_document(
        rows, merges, links,
        tab="FFC65911",
        dimension=f"A1:U{last}",
        cols=f"<cols>{cols}</cols>",
        header="&amp;L&amp;8&amp;K1F4E79Front Node Site&amp;R&amp;8Transmission Link Health Review  |  Ericsson  |  5 Nov 2026 – 7 Nov 2026",
        footer="&amp;L&amp;8Each hour is the maximum In Throughput (Mbps)&amp;R&amp;8Page &amp;P of &amp;N",
        freeze=header_row,
    )
    xml = xml.replace(
        "</sheetData>",
        f'</sheetData><autoFilter ref="A{header_row}:U{last}"/>',
        1,
    )
    return xml, rels_xml(rels), first_data, last


def _banner(item: dict, number: int) -> str:
    tags = []
    touch = item.get("touch")
    if item.get("category") == pt.CAT_TX and touch:
        tags.append(
            f"{pt.CAT_TX} {item['display_severity']} · gap {float(touch['gap']) * 100:.1f}% · {pt.TX_COMMENT}"
        )
    elif item.get("category") == pt.CAT_PEAK and touch:
        tags.append(f"{pt.CAT_PEAK} {item['display_severity']}")
    elif item.get("category") == pt.CAT_LICENSE:
        tags.append(f"{pt.CAT_LICENSE} {item.get('display_severity', '')}")
    tail = "    ·    ".join(tags)
    return f"{number:02d}    {item['site']}    ·    {PERIOD}    ·    {tail}"


def _num_cell(ref: str, value: float) -> str:
    return f'<c r="{ref}"><v>{value:.10g}</v></c>'


def append_hourly_charts(sheet4: bytes, sheet5: bytes, drawing: bytes, drawing_rels: bytes, content_types: bytes, charts: list[dict], work: pd.DataFrame):
    """Append one Excel line chart per Front Node site onto the existing hourly-chart sheet."""
    import io
    import xlsxwriter

    match = re.search(br'<dimension ref="A1:[A-Z]+(\d+)"', sheet4)
    if not match:
        raise SystemExit("Hourly chart sheet has no dimension")
    last_row = int(match.group(1))
    existing = last_row // fcc.BLOCK_ROWS
    if existing * fcc.BLOCK_ROWS != last_row:
        raise SystemExit(f"Hourly chart sheet ends at row {last_row}, not a multiple of {fcc.BLOCK_ROWS}")
    days = list(pd.date_range(work["Date"].min().normalize(), work["Date"].max().normalize(), freq="D"))
    if len(days) != fcc.SNAP_DAYS:
        raise SystemExit(f"Expected {fcc.SNAP_DAYS} chart days, found {len(days)}")
    if (existing + len(charts)) * CHART_COLS > 16384:
        raise SystemExit("Front Node chart data would pass Excel's column limit")
    by_site = {site: group for site, group in work.groupby("eNodeB Name", sort=False)}

    buf = io.BytesIO()
    book = xlsxwriter.Workbook(buf, {"in_memory": True, "strings_to_urls": False})
    data = book.add_worksheet("_ChartData")
    dummy = book.add_worksheet("dummy")
    cat_fmt = book.add_format({"font_name": "Calibri", "font_size": 8, "align": "center"})
    data_cells: dict[int, list[str]] = {row: [] for row in range(1, 74)}
    block_rows: list[str] = []
    merges: list[str] = []
    anchors: list[str] = []
    rels: list[str] = []

    print(f"Drawing {len(charts)} Front Node charts after chart {existing}...", flush=True)
    for i, item in enumerate(charts):
        g = existing + i
        c0 = g * CHART_COLS
        excel_top = g * fcc.BLOCK_ROWS + 1
        lines = item["lines"]
        peak = float(lines["peak"])
        level = float(lines["level"])
        headers = ("Date", "Time", "Rx", "Peak", "Touch")
        for offset, name in enumerate(headers):
            ref = f"{col_name(c0 + offset)}1"
            data.write(0, c0 + offset, name)
            data_cells[1].append(c_str(ref, "0", name).replace(' s="0"', "", 1))
        site_rows = by_site.get(item["site"])
        rx_by_key = {}
        if site_rows is not None:
            rx_by_key = {
                (pd.Timestamp(day).normalize(), int(hour)): float(value)
                for day, hour, value in zip(site_rows["Date"], site_rows["Hour"], site_rows["Rx"])
                if np.isfinite(value)
            }
        dates: list[str] = []
        times: list[str] = []
        plotted = [value for value in (peak, level) if np.isfinite(value)]
        r = 0
        for day in days:
            date_label = f"{int(day.day)}/{fcc.MONTHS[int(day.month) - 1]}"
            for hour in range(24):
                r += 1
                excel_r = r + 1
                dates.append(date_label if hour == 0 else "")
                times.append(f"{hour:02d}:00")
                if hour == 0:
                    data.write_string(r, c0, date_label, cat_fmt)
                    data_cells[excel_r].append(c_str(f"{col_name(c0)}{excel_r}", S_CHART_AXIS, date_label))
                else:
                    data.write_blank(r, c0, None)
                data.write_string(r, c0 + 1, f"{hour:02d}:00", cat_fmt)
                data_cells[excel_r].append(c_str(f"{col_name(c0 + 1)}{excel_r}", S_CHART_AXIS, f"{hour:02d}:00"))
                value = rx_by_key.get((day.normalize(), hour))
                if value is not None and np.isfinite(value):
                    data.write_number(r, c0 + 2, value)
                    data_cells[excel_r].append(_num_cell(f"{col_name(c0 + 2)}{excel_r}", value))
                    plotted.append(value)
                if np.isfinite(peak):
                    data.write_number(r, c0 + 3, peak)
                    data_cells[excel_r].append(_num_cell(f"{col_name(c0 + 3)}{excel_r}", peak))
                if np.isfinite(level):
                    data.write_number(r, c0 + 4, level)
                    data_cells[excel_r].append(_num_cell(f"{col_name(c0 + 4)}{excel_r}", level))
        chart = book.add_chart({"type": "line"})
        chart.show_blanks_as("gap")
        chart.set_title({"name": item["site"], "name_font": {"name": "Calibri", "size": 16, "color": "#595959"}})
        series = [
            ("In Throughput (Mbps)", 2, {"color": "#5B9BD5", "width": 2.25}, True),
            (f"Peak ({peak:.2f} Mbit/s)", 3, {"color": "#C00000", "width": 1.5, "dash_type": "dash"}, np.isfinite(peak)),
            (f"Below5%Peak ({level:.2f} Mbit/s)", 4, {"color": "#ED7D31", "width": 1.5, "dash_type": "dash"}, np.isfinite(level)),
        ]
        for name, offset, line, ok in series:
            if not ok:
                continue
            chart.add_series(
                {
                    "name": name,
                    "categories": ["_ChartData", 1, c0, r, c0 + 1],
                    "categories_data": [dates, times],
                    "values": ["_ChartData", 1, c0 + offset, r, c0 + offset],
                    "line": line,
                    "marker": {"type": "none"},
                }
            )
        y_max, y_unit = fcc._chart_y_scale(plotted or [50.0])
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
        dummy.insert_chart(0, 0, chart)

        labels = [
            (0, "Peak (Mbit/s)", f"{peak:.2f}" if np.isfinite(peak) else "—"),
            (2, "Below5%Peak (Mbit/s)", f"{level:.2f}" if np.isfinite(level) else "—"),
        ]
        banner_cells = [c_str(f"A{excel_top}", S_CHART_BANNER, _banner(item, i + 1))]
        banner_cells += [c_blank(f"{col_name(col)}{excel_top}", S_CHART_BANNER) for col in range(1, 8)]
        label_cells = []
        value_cells = []
        for col, label, value in labels:
            label_cells += [
                c_str(f"{col_name(col)}{excel_top + 1}", S_CHART_LABEL, label),
                c_blank(f"{col_name(col + 1)}{excel_top + 1}", S_CHART_LABEL),
            ]
            value_cells += [
                c_str(f"{col_name(col)}{excel_top + 2}", S_CHART_VALUE, value),
                c_blank(f"{col_name(col + 1)}{excel_top + 2}", S_CHART_VALUE),
            ]
            merges.append(f"{col_name(col)}{excel_top + 1}:{col_name(col + 1)}{excel_top + 1}")
            merges.append(f"{col_name(col)}{excel_top + 2}:{col_name(col + 1)}{excel_top + 2}")
        tag_cells = [c_str(f"A{excel_top + 3}", S_CHART_TAG, "Front Node")]
        tag_cells += [c_blank(f"{col_name(col)}{excel_top + 3}", S_CHART_TAG) for col in range(1, 3)]
        block_rows.append(row_xml(excel_top, 26.25, banner_cells, 7).replace('spans="1:8"', 'spans="1:23"', 1))
        block_rows.append(row_xml(excel_top + 1, 15.75, label_cells, 7).replace('spans="1:8"', 'spans="1:23"', 1))
        block_rows.append(row_xml(excel_top + 2, 20.25, value_cells, 7).replace('spans="1:8"', 'spans="1:23"', 1))
        block_rows.append(row_xml(excel_top + 3, 18, tag_cells, 2).replace('spans="1:3"', 'spans="1:23"', 1))
        for offset in range(4, fcc.BLOCK_ROWS):
            block_rows.append(f'<row r="{excel_top + offset}" spans="1:23" ht="15.75" customHeight="1"/>')
        merges.append(f"A{excel_top}:H{excel_top}")
        merges.append(f"A{excel_top + 3}:C{excel_top + 3}")
        chart_no = g + 1
        from_row = g * fcc.BLOCK_ROWS + 4
        anchors.append(
            "<xdr:twoCellAnchor editAs=\"oneCell\">"
            f"<xdr:from><xdr:col>0</xdr:col><xdr:colOff>57150</xdr:colOff><xdr:row>{from_row}</xdr:row><xdr:rowOff>57150</xdr:rowOff></xdr:from>"
            f"<xdr:to><xdr:col>10</xdr:col><xdr:colOff>342900</xdr:colOff><xdr:row>{from_row + 19}</xdr:row><xdr:rowOff>161925</xdr:rowOff></xdr:to>"
            "<xdr:graphicFrame macro=\"\"><xdr:nvGraphicFramePr>"
            f"<xdr:cNvPr id=\"{chart_no + 1}\" name=\"Chart {chart_no}\"/><xdr:cNvGraphicFramePr/>"
            "</xdr:nvGraphicFramePr><xdr:xfrm><a:off x=\"0\" y=\"0\"/><a:ext cx=\"0\" cy=\"0\"/></xdr:xfrm>"
            "<a:graphic><a:graphicData uri=\"http://schemas.openxmlformats.org/drawingml/2006/chart\">"
            "<c:chart xmlns:c=\"http://schemas.openxmlformats.org/drawingml/2006/chart\" "
            "xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\" "
            f"r:id=\"rId{chart_no}\"/>"
            "</a:graphicData></a:graphic></xdr:graphicFrame><xdr:clientData/></xdr:twoCellAnchor>"
        )
        rels.append(
            f'<Relationship Id="rId{chart_no}" '
            'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart" '
            f'Target="../charts/chart{chart_no}.xml"/>'
        )
    book.close()
    buf.seek(0)
    with zipfile.ZipFile(buf) as temp:
        chart_parts = {
            f"xl/charts/chart{existing + i + 1}.xml": temp.read(f"xl/charts/chart{i + 1}.xml")
            for i in range(len(charts))
        }
    first_formula = re.search(br"<c:f>_ChartData!\$[A-Z]+\$2:\$[A-Z]+\$73</c:f>", next(iter(chart_parts.values())))
    if not first_formula:
        raise SystemExit("Front Node chart is missing its _ChartData formula")

    new_last = (existing + len(charts)) * fcc.BLOCK_ROWS
    sheet4 = re.sub(br'<dimension ref="A1:W\d+"', f'<dimension ref="A1:W{new_last}"'.encode(), sheet4, count=1)
    sheet4 = insert_before(sheet4, b"</sheetData>", "".join(block_rows).encode())
    merge_count = re.search(br'<mergeCells count="(\d+)">', sheet4)
    sheet4 = sheet4[: merge_count.start()] + f'<mergeCells count="{int(merge_count.group(1)) + len(merges)}">'.encode() + sheet4[merge_count.end() :]
    sheet4 = sheet4.replace(
        b"</mergeCells>",
        "".join(f'<mergeCell ref="{ref}"/>' for ref in merges).encode() + b"</mergeCells>",
        1,
    )
    last_col = (existing + len(charts)) * CHART_COLS
    sheet5_text = sheet5.decode()
    sheet5_text = re.sub(r'<dimension ref="[^"]+"', f'<dimension ref="A1:{col_name(last_col - 1)}73"', sheet5_text, count=1)
    sheet5_text = sheet5_text.replace('spans="1:11193"', f'spans="1:{last_col}"')
    pieces = []
    cursor = 0
    for row, cells in data_cells.items():
        token = f'<row r="{row}"'
        at = sheet5_text.find(token, cursor)
        end = sheet5_text.find("</row>", at)
        if at < 0 or end < 0:
            raise SystemExit(f"_ChartData is missing row {row}")
        pieces.append(sheet5_text[cursor:end])
        pieces.append("".join(cells))
        cursor = end
    pieces.append(sheet5_text[cursor:])
    sheet5 = "".join(pieces).encode()
    drawing = insert_before(drawing, b"</xdr:wsDr>", "".join(anchors).encode())
    drawing_rels = insert_before(drawing_rels, b"</Relationships>", "".join(rels).encode())
    overrides = "".join(
        f'<Override PartName="/xl/charts/chart{existing + i + 1}.xml" '
        'ContentType="application/vnd.openxmlformats-officedocument.drawingml.chart+xml"/>'
        for i in range(len(charts))
    )
    content_types = insert_before(content_types, b"</Types>", overrides.encode())
    chart_rows = {item["site"]: (existing + i) * fcc.BLOCK_ROWS + 1 for i, item in enumerate(charts)}
    return sheet4, sheet5, drawing, drawing_rels, content_types, chart_parts, chart_rows


def method_rows(text_rows: list[tuple[str, str]]) -> tuple[str, list[str]]:
    rows = []
    merges = []
    r = 59
    xml, ref = span(r, 22, S_SECTION, "Excel report — Front Node Site", 1)
    rows.append(xml)
    merges.append(ref)
    r += 1
    for label, body in text_rows:
        cells = [c_str(f"A{r}", S_METHOD_A, label), c_str(f"B{r}", S_METHOD_B, body)]
        rows.append(row_xml(r, 48, cells, 1))
        r += 1
    return "".join(rows), merges


def geodata_sheet(rows: list[dict]) -> str:
    headers = ["Site", "Latitude", "Longitude", "Status", "Tech", "Issue Tech", "District", "Region", "Thana", "Severity"]
    body = [_header_row(1, headers, 18)]
    for i, row in enumerate(rows, start=2):
        cells = [
            c_str(f"A{i}", "0", row["site"]),
            c_num(f"B{i}", "0", row["lat"]),
            c_num(f"C{i}", "0", row["lon"]),
            c_str(f"D{i}", "0", row["status"]),
            c_str(f"E{i}", "0", row["tech"]),
            c_str(f"F{i}", "0", row["issue_tech"]),
            c_str(f"G{i}", "0", row["district"]),
            c_str(f"H{i}", "0", row["region"]),
            c_str(f"I{i}", "0", row["thana"]),
            c_str(f"J{i}", "0", row["severity"]),
        ]
        body.append(row_xml(i, None, cells, 9))
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
        f'<dimension ref="A1:J{len(rows) + 1}"/>'
        '<sheetViews><sheetView workbookViewId="0"/></sheetViews>'
        '<sheetFormatPr defaultRowHeight="15"/>'
        f"<sheetData>{''.join(body)}</sheetData></worksheet>"
    )


def series_map(work: pd.DataFrame) -> tuple[list[str], dict[str, list]]:
    start = work["Date"].min().normalize()
    end = work["Date"].max().normalize()
    days = list(pd.date_range(start, end, freq="D"))
    keys = [(day, hour) for day in days for hour in range(24)]
    labels = [f"{day.day}/{day.strftime('%b')} {hour:02d}:00" for day, hour in keys]
    index = {key: i for i, key in enumerate(keys)}
    series: dict[str, list] = {}
    for site, group in work.groupby("eNodeB Name", sort=False):
        values = [None] * len(keys)
        for day, hour, value in zip(group["Date"], group["Hour"], group["Rx"]):
            slot = index.get((pd.Timestamp(day).normalize(), int(hour)))
            if slot is None or not np.isfinite(value):
                continue
            values[slot] = round(float(value), 2)
        series[str(site)] = values
    return labels, series


def listed_payload(item, front, bands, tech, types, rx, number):
    lines = item["lines"]
    stat = item.get("lic_stat")
    touch = item.get("touch")
    lic = finite(stat["license"]) if stat else None
    front_ne, front_peak = front.get(item["site"], ("", None))
    return {
        "n": number,
        "site": item["site"],
        "tech": tech.get(pt.site_code(item["site"]), fcc.TECH_4G),
        "siteType": types.get(pt.site_code(item["site"])) or "Not found",
        "band": bands.get(item["site"]) or "—",
        "category": item.get("category") or "",
        "severity": item.get("display_severity") or "",
        "flags": ["", "", "", ""],
        "peak": round(float(lines["peak"]), 2),
        "level": round(float(lines["level"]), 2),
        "touchHours": lines["hours_on"],
        "totalHours": lines["hours_n"],
        "ratio": round(float(lines["hours_pct"]), 1) if np.isfinite(lines["hours_pct"]) else None,
        "where": (touch or {}).get("cap_when") or lines["cap_when"],
        "license": round(lic, 2) if lic is not None else None,
        "threshold": round(lic * pt.LICENSE_USE, 2) if lic is not None else None,
        "gap": round(gap_pct(item), 1) if gap_pct(item) is not None else None,
        "listed": True,
        "frontNode": front_ne,
        "frontPeak": round(float(front_peak), 2) if front_peak is not None and np.isfinite(front_peak) else None,
        "rx": rx,
    }


def unlisted_payload(site, rx, front, bands, tech, types, stats):
    vals = [value for value in rx if value is not None]
    peak = max(vals) if vals else None
    level = round(peak * (1 - pt.TOUCH_BAND), 2) if peak is not None else None
    ratio = None
    if peak is not None and vals:
        ratio = round(100.0 * sum(1 for value in vals if value >= peak * (1 - pt.TOUCH_BAND)) / len(vals), 1)
    stat = stats.get(site)
    lic = finite(stat["license"]) if stat else None
    gap = None
    if lic and peak is not None and lic > 0:
        gap = round((lic - peak) / lic * 100.0, 1)
    front_ne, front_peak = front.get(site, ("", None))
    return {
        "site": site,
        "tech": tech.get(site, fcc.TECH_4G),
        "siteType": types.get(site) or "Not found",
        "band": bands.get(site) or "—",
        "category": "",
        "severity": "",
        "flags": ["", "", "", ""],
        "peak": round(peak, 2) if peak is not None else None,
        "level": level,
        "ratio": ratio,
        "where": "",
        "license": round(lic, 2) if lic is not None else None,
        "threshold": round(lic * pt.LICENSE_USE, 2) if lic is not None else None,
        "gap": gap,
        "listed": False,
        "frontNode": front_ne,
        "frontPeak": round(float(front_peak), 2) if isinstance(front_peak, float) and np.isfinite(front_peak) else None,
        "rx": rx,
    }


def front_tab(charts, work, front, bands, tech, types, summary) -> dict:
    labels, series = series_map(work)
    sites = []
    seen = set()
    for i, item in enumerate(charts):
        sites.append(listed_payload(item, front, bands, tech, types, series.get(item["site"], [None] * len(labels)), i + 1))
        seen.add(item["site"])
    for site in sorted(series):
        if site in seen:
            continue
        sites.append(unlisted_payload(site, series[site], front, bands, tech, types, summary["stats_by_site"]))
    return {
        "label": "FrontNode",
        "bandLabel": "Band",
        "listTitle": "FrontNode_Site",
        "search": "Search site, front node, severity",
        "hint": "Each chart is one site. In Throughput is the maximum of every row for that site and hour, 5–7 Nov 2026. The lines are In Throughput, Peak, and Below5%Peak.",
        "noun": "issue sites",
        "period": "05 Nov 2026 – 07 Nov 2026",
        "rxName": "In Throughput (Mbps)",
        "showOss": False,
        "filters": ["All", "Tx Flat/Chocked", "Peak Touch", "License usage", "All Sites"],
        "labels": labels,
        "sites": sites,
    }


def merge_geo(rows: list[dict], charts, tech_by_code: dict[str, str]) -> tuple[list[dict], dict]:
    by_site = {row["site"]: row for row in rows}
    added = worsened = 0
    missing = []
    for item in charts:
        code = pt.site_code(item["site"])
        grade = pt.GEO_SEV.get(item.get("display_severity")) or "Low"
        current = by_site.get(code)
        if current is None or current.get("lat") is None:
            missing.append(code)
            continue
        previous = current.get("severity") or ""
        merged = worse(previous, grade)
        if previous not in GEO_RANK:
            added += 1
            current["status"] = "Issue"
            current["issue_tech"] = tech_by_code.get(code, fcc.TECH_4G)
        elif GEO_RANK[merged] < GEO_RANK.get(previous, 9):
            worsened += 1
        current["severity"] = merged
    stats = {
        "added": added,
        "worsened": worsened,
        "missing": sorted(set(missing)),
        "issues": sum(1 for row in rows if row.get("severity") in GEO_RANK),
    }
    return rows, stats


def write_geo_html(rows: list[dict], tech_by_code: dict[str, str]) -> None:
    issue_set = {row["site"] for row in rows if row.get("severity") in GEO_RANK}
    geo_rows = [
        {
            "site": row["site"],
            "lat": row["lat"],
            "lon": row["lon"],
            "district": row.get("district") or "",
            "region": row.get("region") or "",
            "tech": row.get("tech") or "",
            "thana": row.get("thana") or "",
            "severity": row.get("severity") if row.get("severity") in GEO_RANK else "",
        }
        for row in rows
        if row.get("lat") is not None and row.get("lon") is not None
    ]
    classified = classify_rows(geo_rows, issue_set, tech_by_code)
    thanas = fcc.load_top_thanas(ROOT / fcc.TOP_THANA_FILE)
    specs = thana_focus_specs(classified, thanas, tech_by_code)
    issue_rows = [row for row in classified if str(row.get("kind", "")).startswith("Issue")]
    issue_rows.sort(key=lambda row: (GEO_RANK.get(row.get("severity"), 9), row["site"]))
    for index, cluster in enumerate(cluster_issue_sites(issue_rows), start=1):
        lat_min, lat_max, lon_min, lon_max = cluster_bounds(cluster)
        names = {row["site"] for row in cluster}
        zoom_rows = []
        for row in classified:
            if not (lat_min <= row["lat"] <= lat_max and lon_min <= row["lon"] <= lon_max):
                continue
            if str(row.get("kind", "")).startswith("Issue") and row["site"] not in names:
                zoom_rows.append({**row, "kind": row["tech"] if row["tech"] in ("4G", "5G") else "4G"})
            else:
                zoom_rows.append(row)
        specs.append(
            {
                "label": cluster_label(cluster, index),
                "issue_count": len(cluster),
                "issue_sites": [row["site"] for row in cluster],
                "rows": zoom_rows,
                "lat_min": lat_min,
                "lat_max": lat_max,
                "lon_min": lon_min,
                "lon_max": lon_max,
            }
        )
    render_html(classified, specs, GEOPLOT_HTML, fcc.GEO_TITLE)
    text = GEOPLOT_HTML.read_text(encoding="utf-8")
    text = text.replace(
        "Severity map: bright red Severe / deep red High / muted amber Moderate / calm green Low / blue Decision</p>",
        "Severity map: bright red Severe / deep red High / muted amber Moderate / calm green Low / blue Decision. "
        "Issue dots include the 30 Oct–1 Nov radio lists and the Front Node sites for 5–7 Nov 2026. "
        "A FrontNode_Site Map link opens this page on that site.</p>",
        1,
    )
    GEOPLOT_HTML.write_text(text, encoding="utf-8")


def patch_zip(parts: dict[str, bytes]) -> None:
    tmp = OUT_BOOK.with_suffix(".xlsx.tmp")
    with zipfile.ZipFile(BASE_BOOK) as source, zipfile.ZipFile(tmp, "w", compression=zipfile.ZIP_DEFLATED) as dest:
        seen = set()
        for info in source.infolist():
            data = parts[info.filename] if info.filename in parts else source.read(info.filename)
            dest.writestr(info.filename, data)
            seen.add(info.filename)
        for name, data in parts.items():
            if name not in seen:
                dest.writestr(name, data)
    tmp.replace(OUT_BOOK)


def insert_before(blob: bytes, marker: bytes, extra: bytes) -> bytes:
    at = blob.find(marker)
    if at < 0:
        raise SystemExit(f"Marker not found: {marker[:80]!r}")
    return blob[:at] + extra + blob[at:]


def main() -> None:
    print("Loading EthernetPort files...", flush=True)
    site_hour, front, n_tn_sites = load_ethernet()
    print(f"Sites: {n_tn_sites:,}  ·  site-hours: {len(site_hour):,}", flush=True)
    work, records, summary, charts = score(site_hour)
    peak_max = float(np.nanmax(work["Rx"])) if len(work) else 0
    print(
        f"Scored {work['eNodeB Name'].nunique():,}  ·  peak-ratio listed {len(records):,}  ·  "
        f"charts {len(charts):,}  ·  license listed {len(summary['records']):,}  ·  max peak {peak_max:.1f}",
        flush=True,
    )
    for code in ("CGPTA90", "CGLHG01"):
        stat = summary["stats_by_site"].get(code)
        if stat:
            print(f"  {code} license {stat['license']:.0f}  peak {stat['peak']:.1f}  share {stat['share']:.1f}%", flush=True)
    bands = load_bands()
    bh_flags = load_bh_flags()
    matched = sum(1 for item in charts if item["site"] in bh_flags)
    print(f"BH Flag 1H rows {len(bh_flags):,}  ·  matched to FrontNode_Site {matched:,} of {len(charts):,}", flush=True)
    _geo, tech, types = pt.load_physical()
    print(f"Reading the {BASE_VERSION} workbook shell...", flush=True)
    with zipfile.ZipFile(BASE_BOOK) as book:
        dashboard = book.read("xl/worksheets/sheet1.xml")
        method = book.read("xl/worksheets/sheet7.xml")
        geo_sheet = book.read("xl/worksheets/sheet8.xml")
        geo_rels = book.read("xl/worksheets/_rels/sheet8.xml.rels")
        geodata = book.read("xl/worksheets/sheet9.xml")
        strings_blob = book.read("xl/sharedStrings.xml")
        workbook = book.read("xl/workbook.xml")
        workbook_rels = book.read("xl/_rels/workbook.xml.rels")
        content_types = book.read("[Content_Types].xml")
        sheet4 = book.read("xl/worksheets/sheet4.xml")
        sheet5 = book.read("xl/worksheets/sheet5.xml")
        drawing = book.read("xl/drawings/drawing1.xml")
        drawing_rels = book.read("xl/drawings/_rels/drawing1.xml.rels")
    sheet4, sheet5, drawing, drawing_rels, content_types, chart_parts, chart_rows = append_hourly_charts(
        sheet4, sheet5, drawing, drawing_rels, content_types, charts, work
    )
    strings = shared_strings(strings_blob)
    geo_rows = read_geodata(geodata, strings)
    geo_rows, geo_stats = merge_geo(geo_rows, charts, tech)
    print(
        f"GeoPlot issues {geo_stats['issues']:,}  ·  newly added {geo_stats['added']:,}  ·  "
        f"worse grade {geo_stats['worsened']:,}  ·  no coordinates {len(geo_stats['missing']):,}",
        flush=True,
    )

    start = max(int(match) for match in re.findall(br'<row r="(\d+)"', dashboard)) + 2
    block, merges, links, last_row = dashboard_block(
        start, records, summary, charts, int(work["eNodeB Name"].nunique()), chart_rows
    )
    dashboard = insert_before(dashboard, b"</sheetData>", block.encode())
    dashboard = re.sub(br'<dimension ref="[^"]+"', f'<dimension ref="A1:L{last_row}"/>'.encode(), dashboard, count=1)
    merge_xml = "".join(f'<mergeCell ref="{ref}"/>' for ref in merges).encode()
    count_match = re.search(br'<mergeCells count="(\d+)">', dashboard)
    dashboard = dashboard[: count_match.start()] + f'<mergeCells count="{int(count_match.group(1)) + len(merges)}">'.encode() + dashboard[count_match.end() :]
    dashboard = dashboard.replace(b"</mergeCells>", merge_xml + b"</mergeCells>", 1)
    dashboard = dashboard.replace(b"</hyperlinks>", "".join(links).encode() + b"</hyperlinks>", 1)

    n_tx = sum(1 for rec in records if rec["category"] == pt.CAT_TX)
    n_peak = sum(1 for rec in records if rec["category"] == pt.CAT_PEAK)
    method_body, method_merges = method_rows(
        [
            (
                "Files",
                "EthernetPort File_1_05Nov.rar, EthernetPort File_1_06Nov.rar, EthernetPort File_1_07Nov.rar. "
                "Counter In Throughput (Mbps). The FEG, OSS, and license sheets are not rescored.",
            ),
            (
                "Window",
                "5 Nov 2026 – 7 Nov 2026, 72 hours. This is a different window from the radio report (30 Oct 2026 – 1 Nov 2026).",
            ),
            (
                "Site hour",
                "Site Name is the first 7 characters of NE Name. "
                "Tx usage for a site and hour is the maximum In Throughput (Mbps) among every row of that site and hour, "
                "every NE and every EthernetPort. "
                "BBAKD01 at 00:00 on 5 Nov is 979.9746635, the highest of its 14 rows.",
            ),
            (
                "Front Node",
                "The NE Name at the site whose own maximum In Throughput is the highest. Shown on FrontNode_Site. "
                "The grade uses the hourly site maximum, not that NE alone.",
            ),
            (
                "Same floors as Site Level",
                "At least 24 samples and Peak at least 20 Mbit/s. "
                "Tx Flat/Chocked when Gap = (License − Peak) / License is at least 10%, with the same gap bands. "
                "Peak Touch: Severe at least 40%, High 30% to under 40%, Medium 20% to under 30%, Low 10% to under 20%. "
                "License usage: hours at or above 85% of the summed dlBbCapacityNet, listed from 20% of hours, "
                "Severe at least 50%, High 40% to under 50%, Medium 30% to under 40%, Low 20% to under 30%.",
            ),
            (
                "This window",
                f"Sites checked {work['eNodeB Name'].nunique():,}. "
                f"Tx Flat/Chocked {n_tx:,}. Peak Touch {n_peak:,}. License usage {len(summary['records']):,}. "
                f"FrontNode_Site rows {len(charts):,}. "
                f"GeoPlot newly marked from this list {geo_stats['added']:,}. "
                f"Grade worsened {geo_stats['worsened']:,}.",
            ),
            (
                "Links",
                "Dashboard and FrontNode_Site names open that site's chart on 2. HourlyChartOfIssueSites. "
                "The Map column opens the GeoPlot HTML on that site. "
                "BH Flag 1H is filled from the OSS 1 BusyHour blocks on URGENTBWINC_BusyHour, "
                "30 Oct–1 Nov. Yes when either Data 1H or User 1H matches: "
                "DL Throughput at or below 7 Mbps, Data Volume at least 50 GB, "
                "DL Throughput at or below 5 Mbps, Data Volume at least 100 GB. "
                "No when both values exist and neither matches. Blank when both are missing.",
            ),
        ]
    )
    method = insert_before(method, b"</sheetData>", method_body.encode())
    method = method.replace(b'<dimension ref="A1:B58"/>', f'<dimension ref="A1:B{58 + 1 + 7}"/>'.encode(), 1)
    method_count = re.search(br'<mergeCells count="(\d+)">', method)
    method = method[: method_count.start()] + f'<mergeCells count="{int(method_count.group(1)) + len(method_merges)}">'.encode() + method[method_count.end() :]
    method = method.replace(
        b"</mergeCells>",
        "".join(f'<mergeCell ref="{ref}"/>' for ref in method_merges).encode() + b"</mergeCells>",
        1,
    )

    note = (
        "Each dot is one physical site (the first 7 characters of the ERBS). "
        "The full site is plotted when any of its ERBS is an issue, when the Site Level Dashboard lists that site, "
        "or when the Front Node Site Dashboard lists that site. Severity is the worst grade of those findings. "
        "Medium is drawn as Moderate. Decision is the 10% to under 15% gap. "
        "Front Node uses In Throughput (Mbps) for 5–7 Nov 2026. "
        "The pictures on this sheet remain the 30 Oct–1 Nov radio maps. "
        "The GeoPlot HTML and _GeoData include the Front Node sites."
    )
    geo_sheet = geo_sheet.replace(
        b'<c r="A5" s="3" t="s"><v>11987</v></c>',
        c_str("A5", S_NOTE, note).encode(),
        1,
    )
    geo_row = row_xml(
        6, 22,
        [c_str("A6", S_LINK, f"Open {GEOPLOT_HTML.name} (includes Front Node sites)")]
        + [c_blank(f"{col_name(col)}6", S_LINK) for col in range(1, 18)],
        17,
    )
    geo_sheet = insert_before(geo_sheet, b'<row r="7"', geo_row.encode())
    geo_count = re.search(br'<mergeCells count="(\d+)">', geo_sheet)
    geo_sheet = geo_sheet[: geo_count.start()] + f'<mergeCells count="{int(geo_count.group(1)) + 1}">'.encode() + geo_sheet[geo_count.end() :]
    geo_sheet = geo_sheet.replace(b"</mergeCells>", b'<mergeCell ref="A6:R6"/></mergeCells>', 1)
    geo_sheet = insert_before(
        geo_sheet,
        b"<pageMargins",
        f'<hyperlinks><hyperlink ref="A6" r:id="rId2" display="{xml_text(GEOPLOT_HTML.name)}"/></hyperlinks>'.encode(),
    )
    geo_rels = geo_rels.replace(
        b"</Relationships>",
        f'<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" '
        f'Target="{GEOPLOT_HTML.name}" TargetMode="External"/></Relationships>'.encode(),
        1,
    )

    strings_blob = strings_blob.replace(
        b"Site Level Dashboard is first. ERBS Dashboard follows. No. starts at 1 on the Site Level charts and continues on the ERBS charts.",
        b"Site Level Dashboard is first. ERBS Dashboard follows. Front Node Site Dashboard is last. "
        b"No. starts at 1 on the Site Level charts and continues on the ERBS charts. "
        b"Front Node numbers start again at 1. Those charts follow the radio charts on 2. HourlyChartOfIssueSites.",
    )
    strings_blob = strings_blob.replace(
        f"FEGE_Choked_Flat_Sites_Ericsson_{BASE_VERSION}".encode(),
        f"FEGE_Choked_Flat_Sites_Ericsson_{REPORT_VERSION}".encode(),
    )
    strings_blob = strings_blob.replace(f"Excel {BASE_VERSION}".encode(), f"Excel {REPORT_VERSION}".encode())
    strings_blob = strings_blob.replace(f"{BASE_VERSION}    ".encode(), f"{REPORT_VERSION}    ".encode())

    site_xml, site_rels, _first, _last = site_sheet(charts, front, bands, tech, types, chart_rows, bh_flags)
    workbook = workbook.replace(
        b'<sheet name="1. Site List_ERBS" sheetId="3" r:id="rId3"/>',
        b'<sheet name="1. Site List_ERBS" sheetId="3" r:id="rId3"/>'
        b'<sheet name="FrontNode_Site" sheetId="14" r:id="rId17"/>',
        1,
    )
    workbook_rels = workbook_rels.replace(
        b"</Relationships>",
        b'<Relationship Id="rId17" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet14.xml"/>'
        b"</Relationships>",
        1,
    )
    content_types = content_types.replace(
        b"</Types>",
        b'<Override PartName="/xl/worksheets/sheet14.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>'
        b"</Types>",
        1,
    )
    print(f"Writing {REPORT_VERSION} workbook...", flush=True)
    patch_zip(
        {
            "xl/worksheets/sheet1.xml": dashboard,
            "xl/worksheets/sheet4.xml": sheet4,
            "xl/worksheets/sheet5.xml": sheet5,
            "xl/worksheets/sheet7.xml": method,
            "xl/worksheets/sheet8.xml": geo_sheet,
            "xl/worksheets/sheet9.xml": geodata_sheet(geo_rows).encode(),
            "xl/worksheets/sheet14.xml": site_xml.encode(),
            "xl/worksheets/_rels/sheet8.xml.rels": geo_rels,
            "xl/worksheets/_rels/sheet14.xml.rels": site_rels.encode(),
            "xl/drawings/drawing1.xml": drawing,
            "xl/drawings/_rels/drawing1.xml.rels": drawing_rels,
            "xl/sharedStrings.xml": strings_blob,
            "xl/workbook.xml": workbook,
            "xl/_rels/workbook.xml.rels": workbook_rels,
            "[Content_Types].xml": content_types,
            **chart_parts,
        }
    )
    with zipfile.ZipFile(OUT_BOOK) as book:
        workbook_xml = book.read("xl/workbook.xml")
        if b"FrontNode_Charts" in workbook_xml or b"FrontNode_Site" not in workbook_xml:
            raise SystemExit("Workbook sheet list is missing FrontNode_Site or still has FrontNode_Charts")
        sheet4_out = book.read("xl/worksheets/sheet4.xml")
        first_anchor = chart_rows[charts[0]["site"]]
        if f"'{SNAP_SHEET}'!A{first_anchor}".encode() not in book.read("xl/worksheets/sheet1.xml"):
            raise SystemExit("Dashboard does not open the hourly chart sheet")
        if f"'{SNAP_SHEET}'!A{first_anchor}".encode() not in book.read("xl/worksheets/sheet14.xml"):
            raise SystemExit("FrontNode_Site does not open the hourly chart sheet")
        if b">Front Node<" not in sheet4_out:
            raise SystemExit("Hourly chart sheet has no Front Node block")
        front_sheet = book.read("xl/worksheets/sheet14.xml")
        if b">Yes<" not in front_sheet:
            raise SystemExit("BH Flag 1H is still blank on FrontNode_Site")
        if b"License (Mbps)" in front_sheet or "Hours ≥85%".encode() in front_sheet:
            raise SystemExit("FrontNode_Site still has the license columns")
        front_chart = book.read(f"xl/charts/chart{1599 + 1}.xml")
        if b"85% of license" in front_chart or b"License (" in front_chart:
            raise SystemExit("Front Node chart still draws the license lines")
        chart_name = f"xl/charts/chart{1599 + len(charts)}.xml"
        if chart_name not in book.namelist():
            raise SystemExit(f"Missing {chart_name}")
        for name in (
            "xl/worksheets/sheet1.xml",
            "xl/worksheets/sheet14.xml",
            "xl/worksheets/sheet9.xml",
        ):
            ET.fromstring(book.read(name))
    print(f"Workbook {OUT_BOOK.name} {OUT_BOOK.stat().st_size / 1e6:.1f} MB", flush=True)

    print("Building the chart page...", flush=True)
    charts_html.EXCEL_VERSION = REPORT_VERSION
    charts_html.HTML_VERSION = REPORT_VERSION.lstrip("v")
    view = charts_html.load_chart_view(OUT_BOOK)
    view["htmlVersion"] = REPORT_VERSION.lstrip("v")
    view["oss"] = charts_html.load_oss_1h(OUT_BOOK)
    charts_html.add_all_sites(view, ROOT / "FEG_KPI_Eric.xlsb", ROOT / pt.LICENSE_FILE, ROOT / pt.GEO_FILE)
    charts_html.attach_hup(view, OUT_BOOK)
    view["tabs"]["FrontNode"] = front_tab(charts, work, front, bands, tech, types, summary)
    charts_html.write_html(view, CHARTS_HTML)
    with zipfile.ZipFile(CHARTS_ZIP, "w", compression=zipfile.ZIP_DEFLATED) as archive:
        archive.write(CHARTS_HTML, CHARTS_HTML.name)
    print(
        f"Charts {CHARTS_HTML.name} {CHARTS_HTML.stat().st_size / 1e6:.1f} MB  ·  "
        f"FrontNode listed {sum(1 for site in view['tabs']['FrontNode']['sites'] if site['listed']):,}",
        flush=True,
    )
    print("Building GeoPlot HTML...", flush=True)
    write_geo_html(geo_rows, tech)
    print(f"GeoPlot {GEOPLOT_HTML.name} {GEOPLOT_HTML.stat().st_size / 1e6:.1f} MB", flush=True)
    counts = {
        "sites_checked": int(work["eNodeB Name"].nunique()),
        "tx": n_tx,
        "peak": n_peak,
        "license": len(summary["records"]),
        "rows": len(charts),
        "geo_added": geo_stats["added"],
        "geo_worsened": geo_stats["worsened"],
        "geo_missing": len(geo_stats["missing"]),
        "max_peak": round(peak_max, 2),
    }
    print("SUMMARY " + repr(counts), flush=True)


if __name__ == "__main__":
    main()
