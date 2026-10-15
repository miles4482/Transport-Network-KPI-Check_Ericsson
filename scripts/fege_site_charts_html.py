#!/usr/bin/env python3
"""HTML chart browser for sheet 1. Site List / 2. HourlyChartOfIssueSites.

Reads a published workbook and writes one HTML file. The Excel report is not modified.
"""

from __future__ import annotations

import argparse
import json
from datetime import datetime, time, timedelta
from pathlib import Path

from openpyxl import load_workbook

EXCEL_EPOCH = datetime(1899, 12, 30)
FLAG_NAMES = (
    "DL Throughput<7Mbps",
    "Data Vol>50GB",
    "DL Throughput<5Mbps",
    "Data Vol>100GB",
)


def _num(value) -> float | None:
    if isinstance(value, (int, float)) and value == value:
        return float(value)
    return None


def _text(value) -> str:
    if value is None:
        return ""
    return str(value).strip()


def _as_date(value) -> datetime | None:
    if isinstance(value, datetime):
        return value.replace(hour=0, minute=0, second=0, microsecond=0)
    if isinstance(value, (int, float)) and value > 1000:
        return EXCEL_EPOCH + timedelta(days=float(value))
    return None


def _as_hour(value) -> int | None:
    if isinstance(value, datetime):
        return int(value.hour)
    if isinstance(value, time):
        return int(value.hour)
    if isinstance(value, (int, float)):
        number = float(value)
        if 0 <= number <= 23 and number == int(number):
            return int(number)
        if 0 <= number < 1:
            return int(round(number * 24)) % 24
    text = _text(value)
    if len(text) >= 2 and text[:2].isdigit():
        return int(text[:2])
    return None


def _round(value: float | None, digits: int = 2):
    if value is None:
        return None
    return round(value, digits)


def load_chart_view(path: Path) -> dict:
    book = load_workbook(path, read_only=True, data_only=True)
    sites_ws = book["1. Site List"]
    period = ""
    sites = []
    for row in sites_ws.iter_rows(min_row=1, max_col=18, values_only=True):
        if row[0] == "No.":
            continue
        if isinstance(row[0], str) and not period and "Oct" in row[0]:
            period = row[0].split(".", 1)[0].strip()
        if not isinstance(row[0], (int, float)):
            continue
        license_mbps = _num(row[16])
        sites.append(
            {
                "n": int(row[0]),
                "site": _text(row[1]),
                "tech": _text(row[2]),
                "siteType": _text(row[3]),
                "category": _text(row[4]),
                "severity": _text(row[5]),
                "flags": [_text(row[6]), _text(row[7]), _text(row[8]), _text(row[9])],
                "peak": _round(_num(row[10])),
                "level": _round(_num(row[11])),
                "touchHours": _num(row[12]),
                "totalHours": _num(row[13]),
                "ratio": _round(_num(row[14]), 1),
                "where": _text(row[15]),
                "license": _round(license_mbps),
                "threshold": _round(license_mbps * 0.85) if license_mbps else None,
                "gap": _round(_num(row[17]), 1),
            }
        )
    by_site = {item["site"]: item for item in sites}
    hourly = book["3. Hourly KPI"]
    hours: set[tuple[datetime, int]] = set()
    series: dict[str, dict[tuple[datetime, int], float | None]] = {item["site"]: {} for item in sites}
    for row in hourly.iter_rows(min_row=1, max_col=8, values_only=True):
        site = _text(row[2])
        if site not in by_site:
            continue
        day = _as_date(row[0])
        hour = _as_hour(row[1])
        if day is None or hour is None:
            continue
        key = (day, hour)
        hours.add(key)
        series[site][key] = _round(_num(row[3]))
    book.close()
    axis = sorted(hours)
    labels = [f"{day.day}/{day.strftime('%b')} {hour:02d}:00" for day, hour in axis]
    for item in sites:
        values = series[item["site"]]
        item["rx"] = [values.get(key) for key in axis]
    return {"period": period, "labels": labels, "flagNames": list(FLAG_NAMES), "sites": sites}


HTML = r"""<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>Hourly charts — Site List</title>
<style>
  :root {
    --navy: #1F4E79;
    --ink: #1a1a1a;
    --line: #d9d9d9;
    --paper: #f4f6f8;
    --severe: #c00000;
    --high: #c65911;
    --medium: #bf8f00;
    --low: #548235;
    --decision: #1a5276;
  }
  * { box-sizing: border-box; }
  html, body { margin: 0; height: 100%; }
  body {
    font-family: Calibri, "Segoe UI", sans-serif;
    color: var(--ink);
    background: var(--paper);
    display: grid;
    grid-template-columns: 340px 1fr;
    height: 100vh;
  }
  aside {
    background: #fff;
    border-right: 1px solid var(--line);
    display: flex;
    flex-direction: column;
    min-height: 0;
  }
  header, .chart-head { padding: 14px 16px 10px; }
  h1 { margin: 0; font-size: 18px; color: var(--navy); }
  .sub { margin: 4px 0 0; color: #595959; font-size: 12px; }
  .tools { padding: 0 12px 10px; display: flex; flex-direction: column; gap: 8px; }
  input[type="search"] {
    width: 100%;
    border: 1px solid #b0b0b0;
    border-radius: 4px;
    padding: 8px 10px;
    font: inherit;
    font-size: 14px;
  }
  .filters { display: flex; flex-wrap: wrap; gap: 6px; }
  .filters button, .nav button {
    border: 1px solid #b0b0b0;
    background: #fff;
    border-radius: 14px;
    padding: 4px 10px;
    font: inherit;
    font-size: 12px;
    cursor: pointer;
  }
  .filters button.on { background: var(--navy); color: #fff; border-color: var(--navy); }
  #list { overflow: auto; flex: 1; }
  .site {
    width: 100%;
    text-align: left;
    border: 0;
    border-bottom: 1px solid #eee;
    background: #fff;
    padding: 8px 12px;
    cursor: pointer;
    font: inherit;
  }
  .site.on { background: #deebf7; }
  .site .name { font-weight: 700; color: var(--navy); }
  .site .meta { color: #595959; font-size: 12px; margin-top: 2px; }
  .sev { font-weight: 700; }
  .sev.Severe { color: var(--severe); }
  .sev.High { color: var(--high); }
  .sev.Medium { color: var(--medium); }
  .sev.Low { color: var(--low); }
  .sev.Decision { color: var(--decision); }
  main { min-width: 0; min-height: 0; display: flex; flex-direction: column; }
  .chart-head { display: flex; justify-content: space-between; gap: 12px; align-items: flex-start; }
  .kpis { display: flex; flex-wrap: wrap; gap: 8px; padding: 0 16px 8px; }
  .kpi {
    background: #fff;
    border: 1px solid var(--line);
    border-radius: 4px;
    padding: 6px 10px;
    min-width: 120px;
  }
  .kpi b { display: block; font-size: 16px; }
  .kpi span { color: #595959; font-size: 11px; }
  .flags { display: flex; flex-wrap: wrap; gap: 6px; padding: 0 16px 8px; }
  .flag { font-size: 12px; border-radius: 4px; padding: 3px 8px; background: #fff; border: 1px solid var(--line); }
  .flag.yes { background: #c6efce; color: #006100; font-weight: 700; border-color: #a9d08e; }
  #stage { flex: 1; min-height: 280px; margin: 0 16px 16px; background: #fff; border: 1px solid #d0d0d0; position: relative; }
  canvas { width: 100%; height: 100%; display: block; }
  .legend { display: flex; flex-wrap: wrap; gap: 10px; padding: 8px 16px 0; font-size: 12px; }
  .legend label { display: flex; align-items: center; gap: 6px; cursor: pointer; }
  .swatch { width: 22px; height: 0; border-top-width: 3px; border-top-style: solid; }
  .swatch.dash { border-top-style: dashed; }
  #tip {
    position: absolute; display: none; pointer-events: none;
    background: #fff; border: 1px solid #b0b0b0; border-radius: 4px;
    padding: 6px 8px; font-size: 12px; box-shadow: 0 2px 8px rgba(0,0,0,.12);
  }
  .nav { display: flex; gap: 6px; }
  @media (max-width: 860px) {
    body { grid-template-columns: 1fr; grid-template-rows: 42vh 1fr; }
    aside { border-right: 0; border-bottom: 1px solid var(--line); }
  }
</style>
</head>
<body>
<aside>
  <header>
    <h1>1. Site List</h1>
    <p class="sub" id="period"></p>
  </header>
  <div class="tools">
    <input id="q" type="search" placeholder="Search ERBS, tech, severity" autocomplete="off">
    <div class="filters" id="filters"></div>
  </div>
  <div id="list"></div>
</aside>
<main>
  <div class="chart-head">
    <div>
      <h1 id="title">Hourly chart</h1>
      <p class="sub" id="caption"></p>
    </div>
    <div class="nav">
      <button id="prev" type="button">Previous</button>
      <button id="next" type="button">Next</button>
    </div>
  </div>
  <div class="kpis" id="kpis"></div>
  <div class="flags" id="flags"></div>
  <div class="legend" id="legend"></div>
  <div id="stage">
    <canvas id="chart"></canvas>
    <div id="tip"></div>
  </div>
</main>
<script>
const VIEW = __DATA__;
const SERIES = [
  { key: "rx", name: "Max DL Capacity Usage (Mbps)", color: "#5B9BD5", dash: false },
  { key: "peak", name: "Peak", color: "#C00000", dash: true },
  { key: "level", name: "Below5%Peak", color: "#ED7D31", dash: true },
  { key: "license", name: "License", color: "#548235", dash: true },
  { key: "threshold", name: "85% of license", color: "#833C0C", dash: true }
];
const on = Object.fromEntries(SERIES.map(s => [s.key, true]));
let filter = "All";
let query = "";
let current = 0;
const filters = ["All", "Tx Flat/Chocked", "Peak Touch", "License usage"];

function primary(site) {
  return (site.category || "").split(" + ")[0];
}
function visible() {
  const q = query.trim().toLowerCase();
  return VIEW.sites.filter(site => {
    if (filter !== "All" && primary(site) !== filter) return false;
    if (!q) return true;
    return [site.site, site.tech, site.siteType, site.category, site.severity, String(site.n)]
      .join(" ").toLowerCase().includes(q);
  });
}
function yScale(values) {
  const finite = values.filter(v => v != null && isFinite(v));
  const peak = finite.length ? Math.max(...finite) : 50;
  const target = Math.max(peak * 1.15, peak + 4);
  const brackets = [[25,5],[50,10],[75,25],[100,25],[125,25],[150,25],[200,50],[250,50],[300,50],[350,50],[400,50],[500,100],[600,100],[800,100],[1000,200],[1500,250],[2000,250]];
  for (const pair of brackets) if (target <= pair[0]) return pair;
  const yMax = Math.max(200, Math.ceil(target / 200) * 200);
  return [yMax, 200];
}
function lineValues(site, key) {
  if (key === "rx") return site.rx;
  const value = site[key];
  return site.rx.map(() => value);
}
function draw() {
  const site = VIEW.sites[current];
  const canvas = document.getElementById("chart");
  const stage = document.getElementById("stage");
  const dpr = window.devicePixelRatio || 1;
  const width = stage.clientWidth;
  const height = stage.clientHeight;
  canvas.width = Math.max(1, Math.floor(width * dpr));
  canvas.height = Math.max(1, Math.floor(height * dpr));
  canvas.style.width = width + "px";
  canvas.style.height = height + "px";
  const ctx = canvas.getContext("2d");
  ctx.setTransform(dpr, 0, 0, dpr, 0, 0);
  ctx.clearRect(0, 0, width, height);
  const pad = { l: 58, r: 16, t: 16, b: 54 };
  const plotW = width - pad.l - pad.r;
  const plotH = height - pad.t - pad.b;
  const packed = [];
  SERIES.forEach(spec => { if (on[spec.key]) packed.push(...lineValues(site, spec.key)); });
  const [yMax, yUnit] = yScale(packed);
  const n = VIEW.labels.length;
  const xAt = i => pad.l + (n <= 1 ? plotW / 2 : (i / (n - 1)) * plotW);
  const yAt = v => pad.t + plotH - (v / yMax) * plotH;
  ctx.strokeStyle = "#d9d9d9";
  ctx.fillStyle = "#595959";
  ctx.font = "12px Calibri, sans-serif";
  ctx.textAlign = "right";
  ctx.textBaseline = "middle";
  for (let y = 0; y <= yMax + 0.1; y += yUnit) {
    const py = yAt(y);
    ctx.beginPath();
    ctx.moveTo(pad.l, py);
    ctx.lineTo(width - pad.r, py);
    ctx.stroke();
    ctx.fillText(String(Math.round(y)), pad.l - 8, py);
  }
  ctx.textAlign = "center";
  ctx.textBaseline = "top";
  VIEW.labels.forEach((label, i) => {
    if (!label.endsWith("00:00") && !label.endsWith("12:00")) return;
    const x = xAt(i);
    ctx.save();
    ctx.translate(x, pad.t + plotH + 8);
    ctx.rotate(-Math.PI / 3);
    ctx.fillText(label, 0, 0);
    ctx.restore();
  });
  ctx.strokeStyle = "#b0b0b0";
  ctx.beginPath();
  ctx.moveTo(pad.l, pad.t);
  ctx.lineTo(pad.l, pad.t + plotH);
  ctx.lineTo(width - pad.r, pad.t + plotH);
  ctx.stroke();
  SERIES.forEach(spec => {
    if (!on[spec.key]) return;
    const values = lineValues(site, spec.key);
    ctx.beginPath();
    ctx.strokeStyle = spec.color;
    ctx.lineWidth = spec.key === "rx" ? 2.25 : 1.5;
    ctx.setLineDash(spec.dash ? [6, 4] : []);
    let started = false;
    values.forEach((value, i) => {
      if (value == null || !isFinite(value)) { started = false; return; }
      const x = xAt(i), y = yAt(value);
      if (!started) { ctx.moveTo(x, y); started = true; }
      else ctx.lineTo(x, y);
    });
    ctx.stroke();
  });
  ctx.setLineDash([]);
  canvas._geom = { pad, plotW, plotH, n, xAt, yMax };
}
function showTip(event) {
  const canvas = document.getElementById("chart");
  const geom = canvas._geom;
  const tip = document.getElementById("tip");
  if (!geom) return;
  const rect = canvas.getBoundingClientRect();
  const x = event.clientX - rect.left;
  const y = event.clientY - rect.top;
  if (x < geom.pad.l || x > geom.pad.l + geom.plotW || y < geom.pad.t || y > geom.pad.t + geom.plotH) {
    tip.style.display = "none";
    return;
  }
  const i = Math.max(0, Math.min(geom.n - 1, Math.round((x - geom.pad.l) / geom.plotW * (geom.n - 1))));
  const site = VIEW.sites[current];
  const rows = SERIES.filter(spec => on[spec.key]).map(spec => {
    const value = lineValues(site, spec.key)[i];
    const shown = value == null ? "—" : value.toFixed(2);
    return `<div style="color:${spec.color}">${spec.name}: ${shown}</div>`;
  }).join("");
  tip.innerHTML = `<b>${VIEW.labels[i]}</b>${rows}`;
  tip.style.display = "block";
  const left = Math.min(x + 12, rect.width - 220);
  tip.style.left = Math.max(8, left) + "px";
  tip.style.top = Math.max(8, y + 12) + "px";
}
function kpi(label, value, unit) {
  const shown = value == null || value === "" ? "—" : value;
  return `<div class="kpi"><b>${shown}${unit && shown !== "—" ? unit : ""}</b><span>${label}</span></div>`;
}
function renderChart() {
  const site = VIEW.sites[current];
  document.getElementById("title").textContent = `${String(site.n).padStart(2, "0")}    ${site.site}`;
  document.getElementById("caption").textContent = `${VIEW.period}    ·    ${site.category}    ·    ${site.severity}`;
  document.getElementById("kpis").innerHTML = [
    kpi("Peak (Mbit/s)", site.peak, ""),
    kpi("Below5%Peak (Mbit/s)", site.level, ""),
    kpi("License (Mbps)", site.license, ""),
    kpi("85% of license", site.threshold, ""),
    kpi("Gap (%)", site.gap, ""),
    kpi("Peak ratio (%)", site.ratio, ""),
    kpi("Where touched", site.where, ""),
    kpi("Tech", site.tech, ""),
    kpi("Site Type", site.siteType, "")
  ].join("");
  document.getElementById("flags").innerHTML = VIEW.flagNames.map((name, i) => {
    const value = site.flags[i] || "";
    return `<span class="flag ${value === "Yes" ? "yes" : ""}">${name}: ${value || "—"}</span>`;
  }).join("");
  const rows = visible();
  document.querySelectorAll(".site").forEach(button => {
    button.classList.toggle("on", Number(button.dataset.index) === current);
  });
  const selected = document.querySelector(`.site[data-index="${current}"]`);
  if (selected) selected.scrollIntoView({ block: "nearest" });
  try { history.replaceState(null, "", "#" + encodeURIComponent(site.site)); } catch (err) {}
  draw();
}
function renderList() {
  const rows = visible();
  const list = document.getElementById("list");
  list.innerHTML = rows.map(site => `
    <button class="site ${site.n - 1 === current ? "on" : ""}" data-index="${site.n - 1}" type="button">
      <div><span class="name">${site.n}. ${site.site}</span></div>
      <div class="meta">${primary(site)} · <span class="sev ${site.severity}">${site.severity}</span> · ${site.tech} · ${site.siteType}</div>
    </button>`).join("") || `<p class="sub" style="padding:12px">No sites match.</p>`;
  document.getElementById("filters").innerHTML = filters.map(name => {
    const count = name === "All" ? VIEW.sites.length : VIEW.sites.filter(site => primary(site) === name).length;
    return `<button type="button" data-filter="${name}" class="${filter === name ? "on" : ""}">${name} ${count}</button>`;
  }).join("");
}
function selectIndex(index) {
  current = Math.max(0, Math.min(VIEW.sites.length - 1, index));
  renderList();
  renderChart();
}
function step(delta) {
  const rows = visible();
  if (!rows.length) return;
  let pos = rows.findIndex(site => site.n - 1 === current);
  if (pos < 0) pos = 0;
  pos = Math.max(0, Math.min(rows.length - 1, pos + delta));
  selectIndex(rows[pos].n - 1);
}
document.getElementById("list").addEventListener("click", event => {
  const button = event.target.closest(".site");
  if (!button) return;
  selectIndex(Number(button.dataset.index));
});
document.getElementById("filters").addEventListener("click", event => {
  const button = event.target.closest("button");
  if (!button) return;
  filter = button.dataset.filter;
  const rows = visible();
  if (rows.length && !rows.some(site => site.n - 1 === current)) current = rows[0].n - 1;
  renderList();
  renderChart();
});
document.getElementById("q").addEventListener("input", event => {
  query = event.target.value;
  const rows = visible();
  if (rows.length && !rows.some(site => site.n - 1 === current)) current = rows[0].n - 1;
  renderList();
  renderChart();
});
document.getElementById("prev").addEventListener("click", () => step(-1));
document.getElementById("next").addEventListener("click", () => step(1));
document.getElementById("chart").addEventListener("mousemove", showTip);
document.getElementById("chart").addEventListener("mouseleave", () => {
  document.getElementById("tip").style.display = "none";
});
window.addEventListener("resize", draw);
document.getElementById("legend").innerHTML = SERIES.map(spec => `
  <label><input type="checkbox" data-key="${spec.key}" checked>
  <i class="swatch ${spec.dash ? "dash" : ""}" style="border-top-color:${spec.color}"></i>${spec.name}</label>`).join("");
document.getElementById("legend").addEventListener("change", event => {
  const key = event.target.dataset.key;
  if (!key) return;
  on[key] = event.target.checked;
  draw();
});
document.getElementById("period").textContent = `${VIEW.period} · ${VIEW.sites.length} sites · same order as 1. Site List`;
const hash = decodeURIComponent(location.hash.replace(/^#/, ""));
const hashed = VIEW.sites.findIndex(site => site.site === hash || String(site.n) === hash);
current = hashed >= 0 ? hashed : 0;
renderList();
renderChart();
</script>
</body>
</html>
"""


def write_html(view: dict, dest: Path) -> None:
    payload = json.dumps(view, separators=(",", ":")).replace("<", "\\u003c")
    dest.write_text(HTML.replace("__DATA__", payload), encoding="utf-8")


def main() -> None:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("workbook")
    parser.add_argument("-o", "--output", required=True)
    args = parser.parse_args()
    view = load_chart_view(Path(args.workbook))
    dest = Path(args.output)
    write_html(view, dest)
    hours = len(view["labels"])
    print(f"Wrote {dest} · {len(view['sites'])} sites · {hours} hours · {dest.stat().st_size:,} bytes")


if __name__ == "__main__":
    main()
