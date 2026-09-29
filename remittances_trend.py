#!/usr/bin/env python3
"""
Kenya diaspora remittances: fetch Central Bank of Kenya (CBK) data and plot trends.

Sources (both on https://www.centralbank.go.ke/diaspora-remittances/):
  * the monthly "Remittances by Source ('000 USD Equivalent)" Excel file linked
    on the page (country/region detail; currently starts Feb 2019);
  * the HTML table on the same page (Year, Month, North America, Europe,
    Rest of World, Total; starts Jan 2004), used for the long-run total.

Outputs are written to  data/  next to this script. Run:  python remittances_trend.py
"""
from __future__ import annotations

import datetime as dt
import html as htmllib
import re
import sys
import time
from pathlib import Path
from urllib.parse import urljoin

import requests
import pandas as pd
import matplotlib

matplotlib.use("Agg")  # no GUI needed; works on Windows/macOS/headless
import matplotlib.pyplot as plt
import matplotlib.dates as mdates

# --------------------------------------------------------------------------- #
# Configuration
# --------------------------------------------------------------------------- #
PAGE_URL = "https://www.centralbank.go.ke/diaspora-remittances/"
# Used only if no Excel link can be found on the page (update if it goes stale).
FALLBACK_EXCEL_URL = "https://www.centralbank.go.ke/wp-content/uploads/2026/09/August2026.xlsx"

BASE_DIR = Path(__file__).resolve().parent
DATA_DIR = BASE_DIR / "data"
RAW_DIR = DATA_DIR / "raw"
CHART_DIR = DATA_DIR / "charts"

TAX_DATE = pd.Timestamp("2026-01-01")
TAX_LABEL = "1 Jan 2026: US 1% tax on cash-funded remittances"
SOURCE_NOTE = "Source: Central Bank of Kenya. Chart: Elly Okinyo."
CHART1_START = pd.Timestamp("2015-01-01")

TIMEOUT = 45  # seconds
RETRIES = 3
PLAIN_UA = "kenya-remittances-trend/1.0 (+python-requests)"
BROWSER_UA = ("Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
              "(KHTML, like Gecko) Chrome/128.0 Safari/537.36")

MONTHS = {m: i for i, m in enumerate(
    ["jan", "feb", "mar", "apr", "may", "jun", "jul", "aug", "sep", "oct", "nov", "dec"], 1)}

COLORS = {"total": "#1f3b73", "roll": "#c0392b", "usa": "#1f77b4",
          "canada": "#d62728", "europe": "#2ca02c", "na": "#7f7f7f"}


# --------------------------------------------------------------------------- #
# Download helpers
# --------------------------------------------------------------------------- #
def http_get(url: str) -> requests.Response:
    """GET with timeout + retries; switch to a browser User-Agent on 403/406."""
    ua = PLAIN_UA
    last_err: Exception | None = None
    for attempt in range(1, RETRIES + 1):
        try:
            r = requests.get(url, timeout=TIMEOUT, headers={
                "User-Agent": ua,
                "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8",
                "Accept-Language": "en-GB,en;q=0.9"})
            if r.status_code in (403, 406, 429) and ua != BROWSER_UA:
                ua = BROWSER_UA
                last_err = requests.HTTPError(f"HTTP {r.status_code}")
                continue  # retry immediately with browser UA
            r.raise_for_status()
            return r
        except requests.RequestException as e:
            last_err = e
            ua = BROWSER_UA
            if attempt < RETRIES:
                time.sleep(2 * attempt)
    raise RuntimeError(f"Could not download {url}: {last_err}")


def find_excel_link(page_html: str) -> str | None:
    """Pick the remittances Excel link from the CBK page (prefer 'by source')."""
    links = re.findall(r'<a[^>]+href="([^"]+\.xlsx?)"[^>]*>(.*?)</a>', page_html, re.I | re.S)
    if not links:
        return None
    def score(item):
        href, text = item
        t = (re.sub(r"<[^>]+>", " ", text) + " " + href).lower()
        return (("remittance" in t) * 2 + ("source" in t), href)
    href = max(links, key=score)[0]
    return urljoin(PAGE_URL, htmllib.unescape(href))


def parse_page_table(page_html: str) -> pd.DataFrame | None:
    """Parse the HTML table (Year, Month, North America, Europe, RoW, Total)."""
    for table in re.findall(r"<table.*?</table>", page_html, re.S | re.I):
        rows = []
        for tr in re.findall(r"<tr.*?</tr>", table, re.S | re.I):
            cells = [htmllib.unescape(re.sub(r"<[^>]+>", " ", c)).strip()
                     for c in re.findall(r"<t[dh][^>]*>(.*?)</t[dh]>", tr, re.S | re.I)]
            rows.append([re.sub(r"\s+", " ", c) for c in cells])
        header = next((r for r in rows if r and r[0].lower() == "year"), None)
        if not header or len(header) < 3 or "month" not in header[1].lower():
            continue
        unit = 1000.0 if "'000" in " ".join(header) or "000" in " ".join(header) else 1.0
        recs = []
        for r in rows:
            if len(r) != len(header) or not r[0].isdigit():
                continue
            try:
                date = pd.Timestamp(int(r[0]), int(r[1]), 1)
            except ValueError:
                continue
            rec = {"date": date}
            for name, val in zip(header[2:], r[2:]):
                name = re.sub(r"\s*\(.*?\)", "", name).strip()
                try:
                    rec[name] = float(val.replace(",", ""))
                except ValueError:
                    rec[name] = float("nan")
            recs.append(rec)
        if recs:
            df = pd.DataFrame(recs).drop_duplicates("date").set_index("date").sort_index()
            # Early years show 0.00 for regions (not broken down) -> treat as missing.
            for c in df.columns:
                if c.lower() != "total remittances" and not c.lower().startswith("total"):
                    df.loc[df[c] == 0, c] = float("nan")
            return df * unit / 1e6  # -> USD millions
    return None


# --------------------------------------------------------------------------- #
# Excel parsing (robust to wide or long layouts)
# --------------------------------------------------------------------------- #
def month_of(v) -> int | None:
    if isinstance(v, (dt.date, dt.datetime, pd.Timestamp)):
        return v.month
    if isinstance(v, str):
        s = v.strip().lower().rstrip(".")
        if s[:3] in MONTHS and (len(s) <= 3 or s.isalpha()):
            return MONTHS[s[:3]]
    return None


def year_of(v) -> int | None:
    if isinstance(v, (dt.date, dt.datetime, pd.Timestamp)):
        return v.year
    try:
        f = float(str(v).strip().replace("*", ""))
        if f.is_integer() and 1990 <= f <= 2100:
            return int(f)
    except ValueError:
        pass
    return None


def detect_unit_divisor(grid: pd.DataFrame) -> tuple[float, str]:
    """Return divisor to convert sheet values to USD millions, and a description."""
    text = " ".join(str(v) for v in grid.head(12).values.ravel() if isinstance(v, str)).lower()
    text = text.replace("\u2018", "'").replace("\u2019", "'")
    if re.search(r"million|usd\s*m\b|us\$\s*m\b", text):
        return 1.0, "USD millions (per sheet title)"
    if re.search(r"'000|000'|thousand|000s", text):
        return 1000.0, "USD thousands (per sheet title)"
    return 1000.0, "USD thousands (ASSUMED - no unit found in sheet title)"


def to_num(v):
    if v is None or (isinstance(v, float) and pd.isna(v)):
        return float("nan")
    if isinstance(v, (int, float)):
        return float(v)
    try:
        return float(str(v).replace(",", "").strip())
    except ValueError:
        return float("nan")


def parse_wide(grid: pd.DataFrame) -> tuple[pd.DataFrame, dict] | None:
    """Regions in rows, months across columns (current CBK layout)."""
    best = None
    for i in range(min(40, len(grid))):
        n = sum(month_of(v) is not None for v in grid.iloc[i])
        if n >= 6 and (best is None or n > best[1]):
            best = (i, n)
    if best is None:
        return None
    hr = best[0]
    months = grid.iloc[hr]
    # Year row: nearest row above with year-like cells (forward-filled across).
    year_row = None
    for j in range(hr - 1, max(-1, hr - 4), -1):
        if sum(year_of(v) is not None for v in grid.iloc[j]) >= 1:
            year_row = grid.iloc[j]
            break
    col_dates = {}
    cur_year, prev_m = None, None
    for c in grid.columns:
        m = month_of(months[c])
        if year_row is not None and year_of(year_row[c]) is not None:
            cur_year = year_of(year_row[c])
        elif isinstance(months[c], (dt.date, dt.datetime, pd.Timestamp)):
            cur_year = months[c].year
        if m is None:
            continue
        if year_row is None and prev_m is not None and m < prev_m and cur_year:
            cur_year += 1  # no year row: roll over at Dec -> Jan
        prev_m = m
        if cur_year:
            col_dates[c] = pd.Timestamp(cur_year, m, 1)
    if not col_dates:
        return None
    # Label column: first column left of the data with text in rows below header.
    first_data = min(col_dates)
    label_col = None
    for c in [c for c in grid.columns if c < first_data][::-1]:
        if grid.loc[hr + 1:, c].map(lambda v: isinstance(v, str)).sum() >= 3:
            label_col = c
            break
    if label_col is None:
        return None
    records, group = [], None
    for i in range(hr + 1, len(grid)):
        raw = grid.at[i, label_col]
        if not isinstance(raw, str) or not raw.strip():
            continue
        label = raw.strip()
        vals = {d: to_num(grid.at[i, c]) for c, d in col_dates.items()}
        has_data = any(pd.notna(v) for v in vals.values())
        if label.lower().startswith("source"):
            break
        if not has_data:
            group = label  # section heading, e.g. "America", "Europe"
            continue
        indent = len(raw) - len(raw.lstrip())
        if label.lower().startswith("total") or label.lower() == "grand total" or indent == 0 and not group:
            series = label
        elif indent == 0:
            series = label
        else:
            series = f"{group}: {label}" if group else label
        for d, v in vals.items():
            records.append((d, series, group or "", v))
    df = pd.DataFrame(records, columns=["date", "series", "group", "value"])
    info = {"layout": f"wide: header row {hr + 1} (month names), year row above, "
                      f"labels in column {label_col + 1}, {len(col_dates)} month columns"}
    return df, info


def parse_long(grid: pd.DataFrame) -> tuple[pd.DataFrame, dict] | None:
    """Months in rows (Year / Month columns or a date column), regions across."""
    for hr in range(min(40, len(grid))):
        row = [str(v).strip().lower() if isinstance(v, str) else "" for v in grid.iloc[hr]]
        if "year" in row and "month" in row:
            yc, mc = row.index("year"), row.index("month")
            break
    else:
        return None
    series_cols = {c: str(grid.iat[hr, c]).strip() for c in range(grid.shape[1])
                   if c not in (yc, mc) and isinstance(grid.iat[hr, c], str) and grid.iat[hr, c].strip()}
    records, cur_year = [], None
    for i in range(hr + 1, len(grid)):
        y = year_of(grid.iat[i, yc])
        cur_year = y or cur_year
        mv = grid.iat[i, mc]
        m = month_of(mv) or (int(to_num(mv)) if pd.notna(to_num(mv)) and 1 <= to_num(mv) <= 12 else None)
        if not (cur_year and m):
            continue
        for c, name in series_cols.items():
            records.append((pd.Timestamp(cur_year, m, 1), re.sub(r"\s*\(.*?\)", "", name), "", to_num(grid.iat[i, c])))
    if not records:
        return None
    return (pd.DataFrame(records, columns=["date", "series", "group", "value"]),
            {"layout": f"long: header row {hr + 1} with Year/Month columns"})


def parse_excel(path: Path) -> tuple[pd.DataFrame, dict]:
    try:
        sheets = pd.read_excel(path, sheet_name=None, header=None)
    except ImportError as e:
        sys.exit(f"Reading {path.name} needs an extra package ({e}). "
                 f"For old .xls files run:  pip install xlrd")
    best = None
    for name, grid in sheets.items():
        grid = grid.reset_index(drop=True)
        grid.columns = range(grid.shape[1])
        for parser in (parse_wide, parse_long):
            res = parser(grid)
            if res and (best is None or len(res[0]) > len(best[0])):
                div, unit = detect_unit_divisor(grid)
                best = (res[0], {**res[1], "sheet": name, "unit": unit, "divisor": div})
    if best is None:
        sys.exit(f"Could not recognise the layout of {path}. Please check the file.")
    df, info = best
    df["usd_millions"] = df.pop("value") / info["divisor"]
    df = df.dropna(subset=["usd_millions"])
    return df, info


# --------------------------------------------------------------------------- #
# Series selection helpers
# --------------------------------------------------------------------------- #
def pick(wide: pd.DataFrame, *patterns: str) -> str | None:
    for p in patterns:
        for c in wide.columns:
            if re.fullmatch(p, c.strip(), re.I):
                return c
    return None


# --------------------------------------------------------------------------- #
# Charting
# --------------------------------------------------------------------------- #
def style():
    plt.rcParams.update({
        "figure.dpi": 100, "savefig.dpi": 300, "font.size": 10.5,
        "font.family": "DejaVu Sans", "axes.titlesize": 13, "axes.titleweight": "bold",
        "axes.labelsize": 10.5, "axes.spines.top": False, "axes.spines.right": False,
        "axes.grid": True, "grid.color": "#dddddd", "grid.linewidth": 0.6,
        "legend.frameon": False, "legend.fontsize": 9.5})


def finish(fig, axes, path: Path):
    for ax in axes:
        ax.xaxis.set_major_locator(mdates.YearLocator())
        ax.xaxis.set_major_formatter(mdates.DateFormatter("%Y"))
    text = SOURCE_NOTE  # footer is the source line only
    # Place the note below the bottom axis (clear of the tick labels); the
    # tight bounding box on save expands the canvas to include it.
    axes[-1].annotate(text, xy=(0, 0), xycoords="axes fraction", xytext=(-40, -30),
                      textcoords="offset points", ha="left", va="top",
                      fontsize=8, color="#555555")
    fig.savefig(path, bbox_inches="tight")
    plt.close(fig)


def tax_line(ax, last_date, label=True):
    if last_date <= TAX_DATE:
        return
    ax.axvline(TAX_DATE, color="#333333", ls="--", lw=0.9, zorder=1)
    if label:
        ax.annotate(TAX_LABEL, xy=(TAX_DATE, 1), xycoords=("data", "axes fraction"),
                    xytext=(-4, -4), textcoords="offset points", ha="right", va="top",
                    fontsize=8, color="#444444",
                    bbox=dict(boxstyle="round,pad=0.2", fc="white", ec="none", alpha=0.85))


def chart_total(total: pd.Series, path: Path):
    s = total.sort_index()
    roll = s.rolling(12, min_periods=12).sum()
    s, roll = s[s.index >= CHART1_START], roll[roll.index >= CHART1_START]
    fig, (ax1, ax2) = plt.subplots(2, 1, figsize=(10, 7.5), sharex=True,
                                   gridspec_kw={"height_ratios": [1, 1], "hspace": 0.18})
    ax1.plot(s.index, s.values, color=COLORS["total"], lw=1.6, label="Monthly total")
    ax1.set_ylabel("USD millions")
    ax1.set_title("Kenya: diaspora remittance inflows, monthly", loc="left")
    ax2.plot(roll.index, roll.values, color=COLORS["roll"], lw=1.8, label="12-month rolling sum")
    ax2.set_ylabel("USD millions")
    ax2.set_title("12-month rolling sum", loc="left", fontsize=11.5)
    for ax in (ax1, ax2):
        ax.set_ylim(0, ax.get_ylim()[1] * 1.12)  # headroom for the tax-line label
        ax.legend(loc="lower right")
    tax_line(ax1, s.index.max())
    tax_line(ax2, s.index.max(), label=False)
    finish(fig, (ax1, ax2), path)


def chart_regions(wide: pd.DataFrame, cols: dict, path: Path):
    fig, ax = plt.subplots(figsize=(10, 5.8))
    last = wide.index.max()
    for label, (col, color) in cols.items():
        if col is None:
            continue
        s = wide[col].dropna()
        ax.plot(s.index, s.values, color=color, lw=0.9, alpha=0.45)
        ax.plot(s.index, s.rolling(12, min_periods=12).mean(), color=color, lw=2.2,
                label=f"{label}, 12-month average")
    ax.set_ylabel("USD millions")
    ax.set_ylim(0, ax.get_ylim()[1] * 1.12)
    ax.set_title("Kenya: remittance inflows from the USA, Canada and Europe", loc="left")
    # Explain the thin lines in the legend (the footer carries the source only).
    ax.plot([], [], color="#888888", lw=0.9, alpha=0.6, label="Thin lines: monthly values")
    ax.legend(loc="upper left")
    tax_line(ax, last)
    finish(fig, [ax], path)


def chart_share(us_share: pd.Series, path: Path):
    fig, ax = plt.subplots(figsize=(10, 5.8))
    ax.plot(us_share.index, us_share.values, color=COLORS["usa"], lw=2.0,
            label="USA share of total diaspora remittances")
    ax.set_ylabel("USA share of total (%)")
    top = min(100, (int(us_share.max() * 1.25) // 10 + 1) * 10)  # headroom for the tax-line label
    ax.set_ylim(0, top)
    ax.set_title("Kenya: USA share of total diaspora remittances", loc="left")
    ax.legend(loc="lower left")
    tax_line(ax, us_share.index.max())
    finish(fig, [ax], path)


# --------------------------------------------------------------------------- #
# Main
# --------------------------------------------------------------------------- #
def manual_instructions(reason: str):
    print(f"""
ERROR: {reason}

What to do:
  1. Open {PAGE_URL} in your browser.
  2. Click the link 'Remittances by Source ('000 USD Equivalent)' and save the
     Excel file into this folder:
       {RAW_DIR}
  3. Run this script again. It will use the newest .xlsx/.xls file in that folder.
""")
    sys.exit(1)


def get_data():
    RAW_DIR.mkdir(parents=True, exist_ok=True)
    page_html, page_table, excel_path, excel_url = None, None, None, None
    try:
        page_html = http_get(PAGE_URL).text
        (RAW_DIR / "cbk_diaspora_page.html").write_text(page_html, encoding="utf-8")
    except RuntimeError as e:
        print(f"Warning: could not load the CBK page ({e}).")
        cached = RAW_DIR / "cbk_diaspora_page.html"
        if cached.exists():
            print("  Using the previously saved copy of the page.")
            page_html = cached.read_text(encoding="utf-8", errors="replace")
    if page_html:
        page_table = parse_page_table(page_html)
        excel_url = find_excel_link(page_html)
    candidates = [u for u in (excel_url, FALLBACK_EXCEL_URL) if u]
    for url in dict.fromkeys(candidates):
        try:
            content = http_get(url).content
            if not content[:4] in (b"PK\x03\x04", b"\xd0\xcf\x11\xe0"):
                raise RuntimeError("response is not an Excel file")
            excel_path = RAW_DIR / Path(url.split("?")[0]).name
            excel_path.write_bytes(content)
            excel_url = url
            break
        except RuntimeError as e:
            print(f"Warning: Excel download failed from {url}: {e}")
    if excel_path is None:
        manual = sorted([p for p in RAW_DIR.glob("*.xls*") if not p.name.startswith("~$")],
                        key=lambda p: p.stat().st_mtime, reverse=True)
        if manual:
            excel_path, excel_url = manual[0], f"manually saved file {manual[0].name}"
            print(f"Using manually saved file: {excel_path}")
        elif page_table is None:
            manual_instructions("Could not download the CBK remittances data.")
        else:
            print("Warning: no Excel file; country detail (USA, Canada) will be missing.")
    return page_table, excel_path, excel_url


def main():
    DATA_DIR.mkdir(exist_ok=True)
    CHART_DIR.mkdir(parents=True, exist_ok=True)
    style()
    page_table, excel_path, excel_url = get_data()

    tidy_parts, info = [], {}
    if excel_path is not None:
        xdf, info = parse_excel(excel_path)
        xdf["source_table"] = "CBK Excel: Remittances by Source"
        tidy_parts.append(xdf)
    if page_table is not None:
        pt = page_table.reset_index().melt("date", var_name="series", value_name="usd_millions")
        pt["group"] = ""
        pt["source_table"] = "CBK web page table"
        pt["series"] = "Page table: " + pt["series"]
        tidy_parts.append(pt.dropna(subset=["usd_millions"]))
    tidy = pd.concat(tidy_parts, ignore_index=True)[["date", "series", "group", "usd_millions", "source_table"]]
    tidy = tidy.sort_values(["source_table", "series", "date"])
    tidy.to_csv(DATA_DIR / "kenya_remittances_tidy.csv", index=False, date_format="%Y-%m-%d",
                float_format="%.4f")
    wide = tidy.pivot_table(index="date", columns="series", values="usd_millions", aggfunc="first").sort_index()
    wide.to_csv(DATA_DIR / "kenya_remittances_wide.csv", date_format="%Y-%m-%d", float_format="%.4f")

    # Pick series
    x_total = pick(wide, r"grand total", r"total remittances", r"total")
    x_usa = pick(wide, r".*:\s*u\.?\s*s\.?\s*a\.?", r".*united states.*", r".*:\s*usa")
    x_can = pick(wide, r".*:\s*canada")
    x_eur = pick(wide, r"total europe", r"europe")
    x_na = pick(wide, r"total (north )?america")
    p_total = pick(wide, r"page table: total remittances", r"page table: total.*")
    p_na = pick(wide, r"page table: north america")
    p_eur = pick(wide, r"page table: europe")
    main_total = x_total or p_total
    if main_total is None:
        sys.exit("No total series found in the data.")
    last = wide[main_total].dropna().index.max()

    # Chart 1: long-run total from page table if it covers 2015, else Excel.
    if p_total and wide[p_total].dropna().index.min() <= CHART1_START:
        total_long = wide[p_total].dropna()
        if x_total:  # prefer Excel values where available (identical in practice)
            total_long = wide[x_total].dropna().combine_first(total_long)
    else:
        total_long = wide[main_total].dropna()
    f1 = CHART_DIR / "1_monthly_total_with_12m_sum.png"
    chart_total(total_long, f1)

    # Chart 2: USA vs Canada vs Europe
    notes2 = ["Thin lines: monthly values. Thick lines: 12-month rolling averages."]
    usa_col, usa_label = (x_usa, "USA") if x_usa else (x_na or p_na, "North America (USA + Canada + other)")
    eur_col = x_eur or p_eur
    if x_can is None:
        notes2.append("CBK data used here has no separate Canada series; Canada is not shown.")
    else:
        can = wide[x_can]
        start_all = wide[usa_col].dropna().index.min()
        if can.dropna().index.min() > start_all or can.loc[can.first_valid_index():].isna().any():
            notes2.append("Canada is not reported separately for some months; gaps are left blank.")
        notes2.append(f"Country detail from the CBK Excel file, available from {can.dropna().index.min():%b %Y}.")
    if eur_col == x_eur and x_eur:
        notes2.append("Europe = CBK 'Total Europe' (all European countries, incl. UK and Switzerland).")
    reg_wide = wide[[c for c in (usa_col, x_can, eur_col) if c]].dropna(how="all")
    f2 = CHART_DIR / "2_usa_canada_europe.png"
    chart_regions(reg_wide, {usa_label: (usa_col, COLORS["usa"]),
                             "Canada": (x_can, COLORS["canada"]),
                             "Europe": (eur_col, COLORS["europe"])}, f2)
    for n in notes2:  # chart notes go to the console; the chart footer is the source line only
        print(f"Chart 2 note: {n}")

    # Chart 3: USA share of the total (USA row / grand total, both from the Excel file)
    f3 = CHART_DIR / "3_usa_share_of_total.png"
    if x_usa and x_total:
        us_share = (wide[x_usa] / wide[x_total] * 100).dropna()
        chart_share(us_share, f3)
    else:
        print("Warning: no USA series in the Excel file; chart 3 (USA share) not drawn.")

    # ---------------- Summary ----------------
    tot = wide[main_total].dropna()
    ttm = tot.iloc[-12:].sum() if len(tot) >= 12 else float("nan")
    ttm_prev = tot.iloc[-24:-12].sum() if len(tot) >= 24 else float("nan")
    print("\n=== Kenya diaspora remittances (CBK) ===")
    print(f"Excel source : {excel_url}")
    if info:
        print(f"Sheet layout : sheet '{info['sheet']}', {info['layout']}; units: {info['unit']}")
    print(f"Latest month : {last:%B %Y}  total = USD {tot.iloc[-1]:,.1f}m")
    prev_y = tot.get(last - pd.DateOffset(years=1))
    if prev_y:
        print(f"               y/y change vs {last - pd.DateOffset(years=1):%b %Y}: {(tot.iloc[-1] / prev_y - 1) * 100:+.1f}%")
    print(f"Trailing 12m : USD {ttm:,.1f}m (USD {ttm / 1000:.3f}bn)"
          + (f", {(ttm / ttm_prev - 1) * 100:+.1f}% vs prior 12m" if ttm_prev == ttm_prev else ""))
    ytd_lbl = f"Jan-{last:%b} {last.year}"
    for label, col in ((usa_label, usa_col), ("Total", main_total)):
        s = wide[col].dropna()
        cur = s[(s.index.year == last.year) & (s.index <= last)]
        prv = s[(s.index.year == last.year - 1) & (s.index.month <= last.month)]
        if len(cur) == last.month and len(prv) == last.month:
            print(f"{label:<13}: {ytd_lbl} USD {cur.sum():,.1f}m vs USD {prv.sum():,.1f}m "
                  f"a year earlier ({(cur.sum() / prv.sum() - 1) * 100:+.1f}% y/y)")
    if x_total and p_total:
        diff = (wide[x_total] - wide[p_total]).abs().max()
        print(f"Check        : Excel grand total vs page-table total, max abs difference USD {diff:.2f}m")
    print(f"\nOutputs in {DATA_DIR}:")
    for p in [DATA_DIR / "kenya_remittances_tidy.csv", DATA_DIR / "kenya_remittances_wide.csv", f1, f2, f3]:
        if p.exists():
            print("  ", p.relative_to(BASE_DIR))


if __name__ == "__main__":
    main()
