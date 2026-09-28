#!/usr/bin/env python3
"""
MTN F5 Weekly Security Metrics Report Generator
=================================================

Reads the weekly F5 WAF/LTM source workbook (MTN_F5_ATTACKS_SOURCE.xlsx) and
produces:
  1. A filled-in PowerPoint deck (from MTN_Security_Metrics_Report_SOURCE.pptx template)
  2. (optional, only with --out-xlsx) a small KPI summary Excel workbook -
     no longer produced by run_weekly_reports.py (user: not needed, 28-Sep-2026)

Usage:
    python generate_weekly_report.py \
        --source MTN_F5_ATTACKS_SOURCE.xlsx \
        --template MTN_Security_Metrics_Report_SOURCE.pptx \
        --out-pptx OUT_REPORT.pptx \
        [--out-xlsx OUT_SUMMARY.xlsx] \
        [--week-end 2026-07-19]   # optional, defaults to "most recent complete week"

Design notes (see conversation for full rationale):
  - Sheet names are irregular ("13th - 19th July 2026 for IKY",
    "29th - 5th July 2026 for OJT " <- note trailing space). We NEVER match
    sheet names by exact string; we parse the date range + site out of each
    sheet name with regex, trim whitespace, and group sheets into "weeks".
  - Table positions inside a sheet are NOT fixed row numbers - the same
    table can start on a different row on different sheets. We locate every
    table by searching for its title text, then read relative to that.
  - The line chart and the FP chart both look back across the last 4
    weeks already present in the workbook (no external history file needed,
    since Treten appends 3 new sheets to this same workbook every week).
"""

import argparse
import copy
import math
import re
import sys
from datetime import date, datetime, timedelta
from dataclasses import dataclass, field

import openpyxl
from pptx import Presentation
from pptx.chart.data import CategoryChartData
from pptx.dml.color import RGBColor
from pptx.opc.constants import RELATIONSHIP_TYPE as RT
from pptx.util import Emu, Inches, Pt
from pptx.oxml.ns import qn

SITES = ["IKY", "VGC", "OJT"]

FONT_NAME = "MTN Brighter Sans"

MONTHS = {
    "january": 1, "february": 2, "march": 3, "april": 4, "may": 5, "june": 6,
    "july": 7, "august": 8, "september": 9, "october": 10, "november": 11,
    "december": 12,
}


# --------------------------------------------------------------------------
# Sheet-name parsing
# --------------------------------------------------------------------------

@dataclass
class SheetRef:
    sheet_name: str
    site: str
    start_date: date        # Monday of the week this sheet belongs to
    end_date: date          # Sunday of that week
    raw_start: date = None  # the dates exactly as written in the sheet name
    raw_end: date = None
    notes: list = field(default_factory=list)   # anything we had to correct/assume


def _resolve_month(month_text: str):
    """
    Matches a month name from a sheet title against MONTHS, tolerating a
    typo'd or truncated spelling (e.g. "Septemb" instead of "September") -
    seen on the 31st Aug - 6th Sep 2026 sheets. Matches as long as the
    given text is a prefix of exactly one full month name and at least 3
    characters long, so it never guesses between two ambiguous months.
    """
    month_text = month_text.lower()
    if month_text in MONTHS:
        return MONTHS[month_text]
    if len(month_text) < 3:
        return None
    matches = [full for full in MONTHS if full.startswith(month_text)]
    if len(matches) == 1:
        return MONTHS[matches[0]]
    return None


_ORD = r"(\d{1,2})(?:st|nd|rd|th)?"
_MON = r"([A-Za-z]{3,9})\.?"
_YEAR = r"(\d{4})"
# The date part of a sheet name, once the site code and the word "for" are
# taken out. Each pattern -> which groups are (d1, m1, d2, m2, year).
_DATE_PATTERNS = [
    # "7th - 13th Sep 2026", "28th - 4th Oct 2026", "7-13 Sep 2026"
    (re.compile(rf"^{_ORD}\s*-\s*{_ORD}\s+{_MON}\s*,?\s*{_YEAR}$", re.I), ("d1", "d2", "m2", "y")),
    # "28th Sep - 4th Oct 2026", "27th June - 2nd August 2026"
    (re.compile(rf"^{_ORD}\s+{_MON}\s*-\s*{_ORD}\s+{_MON}\s*,?\s*{_YEAR}$", re.I), ("d1", "m1", "d2", "m2", "y")),
    # "Sep 7 - 13 2026", "Sep 28 - Oct 4, 2026"
    (re.compile(rf"^{_MON}\s+{_ORD}\s*-\s*(?:{_MON}\s+)?{_ORD}\s*,?\s*{_YEAR}$", re.I), ("m1", "d1", "m2", "d2", "y")),
]
_SITE_RE = re.compile(r"\b(" + "|".join(SITES) + r")\b", re.I)


def monday_of(d: date) -> date:
    return d - timedelta(days=d.weekday())


def _week_for_range(raw_start: date, raw_end: date) -> date:
    """
    Monday of the week a sheet belongs to, given the dates written in its
    name. Reports run Monday-Sunday, so a correct name is Mon -> Sun. When
    it isn't, trust whichever end of the range still looks right:
      - ends on a Sunday   -> the week ending that Sunday
                              ("19th - 26th July" -> 20-26 Jul,
                               "27th June - 2nd August" -> 27 Jul - 2 Aug)
      - starts on a Monday -> the week starting that Monday
                              ("7th - 14th Sep" -> 7-13 Sep)
      - neither            -> the week the middle of the range falls in if
                              it's roughly a week long ("6th - 12th" ->
                              7-13), otherwise the week of the end date.
    """
    if raw_end.weekday() == 6:
        return raw_end - timedelta(days=6)
    if raw_start.weekday() == 0:
        return raw_start
    span = (raw_end - raw_start).days + 1
    if 5 <= span <= 9:
        return monday_of(raw_start + (raw_end - raw_start) / 2)
    return monday_of(raw_end)


def parse_sheet_name_verbose(name: str):
    """
    Reads a weekly sheet name and returns (SheetRef, None), or (None, reason)
    when it can't be read. Deliberately forgiving (confirmed with the user
    28-Sep-2026 - a sheet name typo shouldn't silently drop a week):
      - site code anywhere, any case, "for" optional:  "IKY 7th - 13th Sep 2026"
      - "st/nd/rd/th" optional, "-", "–", "—" or "to" between the dates
      - month on the end date only ("28th - 4th Oct 2026") OR on both
        ("28th Sep - 4th Oct 2026"), full / short / truncated ("Septemb")
      - month first also works ("Sep 7 - 13 2026")
    Every sheet is then placed in a Monday-Sunday week (see
    _week_for_range), so a slip like "19th - 26th July" (8 days) or "7th -
    14th" still lands in the right week and lines up with the other two
    sites - with a note saying what was assumed.
    Note Excel caps sheet names at 31 characters, which is how "September"
    got cut to "Septemb" - suggest_sheet_names() keeps new names under it.
    """
    text = (name or "").strip()
    sites = {m.upper() for m in _SITE_RE.findall(text)}
    if not sites:
        return None, "no site code (IKY / VGC / OJT) in the name"
    if len(sites) > 1:
        return None, f"more than one site code in the name ({', '.join(sorted(sites))})"
    site = sites.pop()

    rest = _SITE_RE.sub(" ", text)
    rest = re.sub(r"[\u2013\u2014]", "-", rest)             # en / em dash
    rest = re.sub(r"\bto\b", "-", rest, flags=re.I)
    rest = re.sub(r"\bfor\b", " ", rest, flags=re.I)
    rest = re.sub(r"\s+", " ", rest).strip(" -_,")

    parts = None
    for pattern, fields in _DATE_PATTERNS:
        m = pattern.match(rest)
        if m:
            parts = dict(zip(fields, m.groups()))
            break
    if parts is None:
        return None, f"couldn't read dates from '{rest}' (expected e.g. '7th - 13th Sep 2026')"

    end_month = _resolve_month(parts["m2"]) if parts.get("m2") else None
    start_month = _resolve_month(parts["m1"]) if parts.get("m1") else None
    if parts.get("m2") and end_month is None:
        return None, f"'{parts['m2']}' isn't a month name"
    if parts.get("m1") and start_month is None:
        return None, f"'{parts['m1']}' isn't a month name"
    end_month = end_month or start_month
    year = int(parts["y"])
    d1, d2 = int(parts["d1"]), int(parts["d2"])

    if start_month is None:
        # month written only once: "29th - 5th July" => start is in June
        start_month = end_month if d1 <= d2 else (end_month - 1 or 12)
    start_year = year - 1 if start_month > end_month else year
    try:
        raw_start = date(start_year, start_month, d1)
        raw_end = date(year, end_month, d2)
    except ValueError as e:
        return None, f"not a real date ({e})"
    if raw_end < raw_start:
        return None, f"end date {raw_end:%d %b %Y} is before start date {raw_start:%d %b %Y}"

    week_start = _week_for_range(raw_start, raw_end)
    week_end = week_start + timedelta(days=6)
    notes = []
    if (raw_start, raw_end) != (week_start, week_end):
        span = (raw_end - raw_start).days + 1
        notes.append(f"dates in the name ({raw_start:%d %b} - {raw_end:%d %b %Y}, {span} days) aren't a "
                     f"Monday-Sunday week - treated as {week_start:%d %b} - {week_end:%d %b %Y}")
    return SheetRef(sheet_name=name, site=site, start_date=week_start, end_date=week_end,
                    raw_start=raw_start, raw_end=raw_end, notes=notes), None


def parse_sheet_name(name: str):
    """SheetRef for a weekly sheet name, or None if it can't be read (see parse_sheet_name_verbose)."""
    return parse_sheet_name_verbose(name)[0]


def group_sheets_by_week(wb):
    """
    Group all readable sheets into weeks keyed by (Monday, Sunday). If a
    week has two sheets for the same site, the right-most one (added
    last) wins and a note is left on it.
    """
    weeks = {}
    for name in wb.sheetnames:
        ref = parse_sheet_name(name)
        if ref is None:
            continue
        key = (ref.start_date, ref.end_date)
        previous = weeks.setdefault(key, {}).get(ref.site)
        if previous is not None:
            ref.notes.append(f"two {ref.site} sheets for this week ('{previous.sheet_name.strip()}' and "
                             f"'{ref.sheet_name.strip()}') - using '{ref.sheet_name.strip()}', the one further right")
        weeks[key][ref.site] = ref
    return weeks


def _looks_like_week_sheet(name):
    """True for an unreadable name that was probably MEANT to be a weekly sheet
    (has a site code or a year in it) - so e.g. a stray 'Sheet1' isn't nagged about."""
    return bool(_SITE_RE.search(name) or re.search(r"\b20\d\d\b", name))


def latest_real_week(weeks, today=None):
    """The newest week that has actually started - a sheet dated in the
    future (e.g. a 2027 typo) must not decide what "next week" is."""
    today = today or date.today()
    started = [k for k in weeks if k[0] <= today]
    return max(started or weeks, key=lambda k: k[1])


def check_sheet_names(sheetnames, focus_weeks=None, today=None):
    """
    The "sheet check" shown before generating (in the web app, and as
    [WARN] lines in validate_outputs.py). Returns a list of plain-English
    warning strings: unreadable sheet names, weeks missing a site,
    corrected dates, duplicate sheets, weeks dated in the future, and gaps
    between weeks. An empty list means every sheet name was fine.

    focus_weeks: optional list of (Monday, Sunday) keys - per-week problems
    (missing site, corrected dates, duplicates, gaps) are then only
    reported for those weeks, so an old typo from months ago doesn't show
    up on every single run. Unreadable names and future-dated weeks are
    always reported, since either can mean the CURRENT week went missing.
    """
    today = today or date.today()

    class _Names:  # group_sheets_by_week only needs .sheetnames
        pass
    holder = _Names()
    holder.sheetnames = list(sheetnames)

    warnings = []
    for name in sheetnames:
        ref, reason = parse_sheet_name_verbose(name)
        if ref is None and _looks_like_week_sheet(name):
            warnings.append(f"Sheet '{name.strip()}' was ignored - {reason}")

    weeks = group_sheets_by_week(holder)
    keys = sorted(weeks)
    focus = set(keys if focus_weeks is None else focus_weeks)
    for key in keys:
        start, end = key
        label = f"{start:%d %b} - {end:%d %b %Y}"
        if start > today:
            warnings.append(f"Week {label} is in the future - check the month/year in its sheet names")
        if key not in focus:
            continue
        missing = [s for s in SITES if s not in weeks[key]]
        if missing:
            warnings.append(f"Week {label} has no sheet for {', '.join(missing)}")
        for site in SITES:
            ref = weeks[key].get(site)
            if ref is not None:
                for note in ref.notes:
                    warnings.append(f"{site} sheet '{ref.sheet_name.strip()}': {note}")
    started = [k for k in keys if k[0] <= today]   # future typos are reported above
    for a, b in zip(started, started[1:]):
        if a not in focus and b not in focus:
            continue
        gap_weeks = (b[0] - a[0]).days // 7 - 1
        if gap_weeks > 0:
            warnings.append(f"No sheets for {gap_weeks} week(s) between {a[1]:%d %b} and {b[0]:%d %b %Y}")
    return warnings


def _ordinal(d: int) -> str:
    if 11 <= d % 100 <= 13:
        return f"{d}th"
    return f"{d}{ {1: 'st', 2: 'nd', 3: 'rd'}.get(d % 10, 'th') }"


def suggest_sheet_names(after_week_end: date):
    """
    Exact sheet names for the week after `after_week_end` (a Sunday), in
    the same style people already use, always within Excel's 31-character
    sheet-name limit (so nobody is forced to cut "September" short again):
        "21st - 27th Sep 2026 for IKY"        (same month)
        "28th Sep - 4th Oct 2026 for IKY"     (two months - 31 chars max)
    """
    start = monday_of(after_week_end) + timedelta(days=7)
    end = start + timedelta(days=6)
    if start.month == end.month:
        dates = f"{_ordinal(start.day)} - {_ordinal(end.day)} {end:%b} {end.year}"
    else:
        dates = f"{_ordinal(start.day)} {start:%b} - {_ordinal(end.day)} {end:%b} {end.year}"
    names = [f"{dates} for {site}" for site in SITES]
    assert all(len(n) <= 31 for n in names), names
    return (start, end), names


def pick_target_week(weeks, week_end=None):
    """
    Pick the week to report on.
    - If week_end given, pick the week whose end_date matches (or is closest
      without going over).
    - Else pick the most recent complete week (end_date <= today), falling
      back to the latest available week if none qualifies.
    """
    keys = sorted(weeks.keys(), key=lambda k: k[1])  # sort by end_date
    if not keys:
        raise ValueError("No parsable weekly sheets found in workbook.")

    if week_end is not None:
        candidates = [k for k in keys if k[1] <= week_end]
        return candidates[-1] if candidates else keys[-1]

    today = date.today()
    candidates = [k for k in keys if k[1] <= today]
    return candidates[-1] if candidates else keys[-1]


def last_n_weeks(weeks, target_key, n=4):
    keys = sorted(weeks.keys(), key=lambda k: k[1])
    idx = keys.index(target_key)
    start = max(0, idx - n + 1)
    return keys[start: idx + 1]


def load_source_workbook(path, week_end=None, n_weeks=4):
    """
    Loads ONLY what the reports need from the source workbook: the target
    week plus the n_weeks-1 weeks before it, values only.

    Why: every week's sheets are pasted straight from the F5 web GUI, which
    drags along formatting on ~180k empty cells and ~7,000 hyperlinks per
    sheet, and the workbook keeps every week since June. Loading all of it
    the normal way took ~2.3 GB of RAM and ~75 s (26 MB file, 39 sheets,
    Sep 2026). Only the last 4 weeks' VALUES are ever used, so this streams
    just those sheets (openpyxl read-only mode) into a small in-memory
    workbook - no formatting, no links - and nobody has to delete or
    archive old sheets by hand. Confirmed with the user 28-Sep-2026.

    Returns (wb, weeks, target_key) - `weeks` only contains the loaded weeks,
    and every sheet in `wb` keeps its original name, so all the existing
    readers (wb[sheet_name], ws.cell(...), iter_rows) work unchanged.
    """
    src = openpyxl.load_workbook(path, read_only=True, data_only=True)
    try:
        all_weeks = group_sheets_by_week(src)      # sheet NAMES only - no cell data read
        target_key = pick_target_week(all_weeks, week_end)
        keep_keys = last_n_weeks(all_weeks, target_key, n=n_weeks)

        wb = openpyxl.Workbook()
        wb.remove(wb.active)
        for key in keep_keys:
            for ref in all_weeks[key].values():
                src_ws = src[ref.sheet_name]
                # don't trust the sheet's stored size - read every row there is
                src_ws.reset_dimensions()
                ws = wb.create_sheet(ref.sheet_name)
                for r_idx, row in enumerate(src_ws.iter_rows(values_only=True), start=1):
                    for c_idx, value in enumerate(row, start=1):
                        if value is not None:
                            ws.cell(row=r_idx, column=c_idx, value=value)
        # every sheet name in the ORIGINAL file, for the sheet check and the
        # "next week's names" hint (the slim workbook only has the kept weeks)
        wb.source_sheetnames = list(src.sheetnames)
        wb.latest_week = latest_real_week(all_weeks)
    finally:
        src.close()

    weeks = {k: all_weeks[k] for k in keep_keys}
    return wb, weeks, target_key


# --------------------------------------------------------------------------
# In-sheet table readers (locate by title text, not fixed row numbers)
# --------------------------------------------------------------------------

def find_row_containing(ws, text, max_row=None, col_range=range(1, 10)):
    """Return the first row number where any cell in col_range contains `text`."""
    max_row = max_row or ws.max_row
    text_lower = text.lower()
    for r in range(1, max_row + 1):
        for c in col_range:
            v = ws.cell(row=r, column=c).value
            if isinstance(v, str) and text_lower in v.lower():
                return r
    return None


WEEKDAYS = ("Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday", "Sunday")


def find_summary_row(ws, title_text, max_row=None, col_range=range(1, 10)):
    """
    Robust title-row finder for the weekly/summary tables (e.g. 'Weekly
    Request Throughput', 'Weekly WAF Attack Breakdown'). Some source
    sheets inconsistently drop the word 'Weekly' from these titles (e.g.
    just 'Request Throughput ' or 'WAF Attack Breakdown'), which would
    otherwise cause an exact-match search to find nothing and silently
    treat that site's numbers as 0.

    Tries an exact match on `title_text` first. If that fails and
    `title_text` starts with "Weekly ", retries with that prefix
    stripped - but explicitly skips any row whose text starts with a
    day-of-week name, since that would be a *daily* table using the same
    core wording (e.g. "Monday Request Throughput"), not the weekly one.
    """
    row = find_row_containing(ws, title_text, max_row=max_row, col_range=col_range)
    if row is not None:
        return row

    if title_text.lower().startswith("weekly "):
        core = title_text[len("weekly "):]
        core_lower = core.lower()
        scan_max = max_row or ws.max_row
        for r in range(1, scan_max + 1):
            for c in col_range:
                v = ws.cell(row=r, column=c).value
                if isinstance(v, str) and core_lower in v.lower():
                    stripped = v.strip().lower()
                    if any(stripped.startswith(day.lower()) for day in WEEKDAYS):
                        continue
                    return r
    return None


def read_kv_table(ws, title_text, max_scan=30, max_col=15):
    """
    Generic reader for the small 'label -> count' tables (Weekly Request
    Throughput, Sunday Pool Member Status, WAF False Positives, etc.)

    These tables are NOT all in the same columns - e.g. 'WAF False
    Positives' and 'Weekly Request Throughput' put their label/value pair
    in columns C/D, while other tables use B/C. Rather than hardcode a
    column, we locate `title_text`, then scan forward for a header row
    containing a cell whose text is exactly "Count" (every one of these
    tables uses that header) and take (that column - 1) as the label
    column and that column as the value column.
    """
    title_row = find_summary_row(ws, title_text)
    if title_row is None:
        return {}

    header_row, val_col = None, None
    for r in range(title_row, title_row + max_scan):
        for c in range(2, max_col):
            v = ws.cell(row=r, column=c).value
            if isinstance(v, str) and v.strip().lower() == "count":
                header_row, val_col = r, c
                break
        if header_row is not None:
            break
    if header_row is None:
        return {}
    key_col = val_col - 1

    data = {}
    r = header_row + 1
    while r < header_row + max_scan:
        k = ws.cell(row=r, column=key_col).value
        v = ws.cell(row=r, column=val_col).value
        if k is None:
            break
        data[str(k).strip()] = v
        r += 1
    return data


def read_attack_breakdown(ws, title_text="Weekly WAF Attack Breakdown", max_scan=60):
    """
    Reads the Attack Type -> Violation Count table under a given title.
    Returns dict {attack_type: violation_count} (all rows, not just top 10 -
    caller decides how many to keep).
    """
    title_row = find_summary_row(ws, title_text)
    if title_row is None:
        return {}
    header_row = find_row_containing(ws, "Attack Type", max_row=title_row + 10)
    if header_row is None or header_row < title_row:
        return {}
    data = {}
    r = header_row + 1
    while r < header_row + max_scan:
        attack_type = ws.cell(row=r, column=2).value
        violation_count = ws.cell(row=r, column=3).value
        if attack_type is None:
            break
        if str(attack_type).strip().upper() != "TOTAL":
            data[str(attack_type).strip()] = violation_count or 0
        r += 1
    return data


# --------------------------------------------------------------------------
# Weekly KPI / chart-data aggregation
# --------------------------------------------------------------------------

@dataclass
class WeekMetrics:
    label: str                      # e.g. "13th - 19th July 2026"
    start: date
    end: date
    requests_inspected: int = 0
    waf_blocks: int = 0
    pool_available: int = 0
    pool_unavailable: int = 0
    pool_offline: int = 0
    pool_unknown: int = 0
    pool_total: int = 0
    attack_counts: dict = field(default_factory=dict)   # merged across sites
    fp_by_site: dict = field(default_factory=dict)       # {site: count}


def week_label(start: date, end: date) -> str:
    def ord_suffix(d):
        if 11 <= d % 100 <= 13:
            return "th"
        return {1: "st", 2: "nd", 3: "rd"}.get(d % 10, "th")
    return f"{start.strftime('%B')} {start.day}{ord_suffix(start.day)} - {end.strftime('%B')} {end.day}{ord_suffix(end.day)}, {end.year}"


def compute_week_metrics(wb, week_key, site_refs, sites=None):
    """
    site_refs: dict {site: SheetRef} for this particular week.
    sites: which sites to include - defaults to all of SITES (merged/
    combined). Pass e.g. sites=["IKY"] to compute for a single site only.
    """
    sites = sites if sites is not None else SITES
    start, end = week_key
    wm = WeekMetrics(label=week_label(start, end), start=start, end=end)

    for site in sites:
        ref = site_refs.get(site)
        if ref is None:
            continue
        ws = wb[ref.sheet_name]

        # --- Weekly Request Throughput: Requests Inspected / WAF Blocks
        throughput = read_kv_table(ws, "Weekly Request Throughput")
        wm.requests_inspected += int(throughput.get("Requests Inspected", 0) or 0)
        wm.waf_blocks += int(throughput.get("WAF Blocks", 0) or 0)

        # --- Sunday Pool Member Status: Available / Offline / Unknown / TOTAL
        # NOTE: the raw table actually has a 4th category, "Unavailable",
        # which is distinct from "Offline". Since only 3 categories
        # (Available/Offline/Unknown) were specified for the pie chart,
        # "Unavailable" is folded into "Offline" here so the 3 slices still
        # sum to the true TOTAL. Change this line if "Unavailable" should
        # be handled differently.
        pool = read_kv_table(ws, "Sunday Pool Member Status")
        wm.pool_available += int(pool.get("Available", 0) or 0)
        wm.pool_unavailable += int(pool.get("Unavailable", 0) or 0)
        wm.pool_offline += int(pool.get("Offline", 0) or 0)
        wm.pool_unknown += int(pool.get("Unknown", 0) or 0)
        wm.pool_total += int(pool.get("TOTAL", 0) or 0)

        # --- Weekly WAF Attack Breakdown: merge attack-type counts
        attacks = read_attack_breakdown(ws, "Weekly WAF Attack Breakdown")
        for atype, count in attacks.items():
            wm.attack_counts[atype] = wm.attack_counts.get(atype, 0) + int(count or 0)

        # --- WAF False Positives: Source -> Count (Source = the site itself)
        fp = read_kv_table(ws, "WAF False Positives")
        # The "Source" rows name a site; TOTAL row is dropped. We attribute
        # this sheet's FP total to its own site per the confirmed mapping.
        fp_total = 0
        for k, v in fp.items():
            if k.strip().upper() != "TOTAL":
                fp_total += int(v or 0)
        wm.fp_by_site[site] = fp_total

    return wm


# --------------------------------------------------------------------------
# Node changes (day-to-day pool member status changes within the week)
# --------------------------------------------------------------------------

DAYS = ["MONDAY", "TUESDAY", "WEDNESDAY", "THURSDAY", "FRIDAY", "SATURDAY", "SUNDAY"]
NODE_STATUSES = ("AVAILABLE", "OFFLINE", "UNKNOWN", "UNAVAILABLE")
_IP_RE = re.compile(r"^\d{1,3}(\.\d{1,3}){3}(%\d+)?$")


def read_daily_node_states(ws):
    """
    Reads the daily pool-member lists pasted into each site sheet. Every
    day has one block per status, headed e.g. "FRIDAY OFFLINE NODE COUNT",
    and each row under it is one virtual server: name, (optional
    description), IP, port, ... The rows inside a block are not always in
    the same column (they shift by one now and then), so each row is read
    by finding its IP cell and taking the first text cell before it as the
    name, rather than trusting fixed columns.

    Returns {day: {(name, ip, port): {status, ...}}}.
    """
    rows = [list(r) for r in ws.iter_rows(values_only=True)]
    header_idx = None
    for idx, r in enumerate(rows[:15]):
        if any(isinstance(v, str) and "NODE COUNT" in v.upper() for v in r):
            header_idx = idx
            break
    if header_idx is None:
        return {}

    hdr = rows[header_idx]
    heads = []
    for col, v in enumerate(hdr):
        if isinstance(v, str) and "NODE COUNT" in v.upper():
            words = v.upper().split()
            day = next((w for w in words if w in DAYS), None)
            status = next((w for w in words if w in NODE_STATUSES), None)
            if day and status:
                heads.append((col, day, status))

    states = {}
    for k, (col, day, status) in enumerate(heads):
        end = heads[k + 1][0] if k + 1 < len(heads) else len(hdr)
        for r in rows[header_idx + 1:]:
            seg = r[col:end]
            ip_idx = next((j for j, v in enumerate(seg)
                           if isinstance(v, str) and _IP_RE.match(v.strip())), None)
            if ip_idx is None:
                continue
            names = [v for v in seg[:ip_idx] if isinstance(v, str) and v.strip()]
            if not names:
                continue
            port = seg[ip_idx + 1] if ip_idx + 1 < len(seg) else None
            key = (names[0].strip(), seg[ip_idx].strip(), "" if port is None else str(port).strip())
            states.setdefault(day, {}).setdefault(key, set()).add(status)
    return states


def compute_node_changes(states):
    """
    Confirmed with the user 28-Sep-2026: report a node only when it is in
    the list on BOTH days being compared and its status changed. Nodes that
    are merely added to / removed from the list are NOT reported - some
    days' snapshots include a whole extra partition (e.g. IKY's Sunday list
    jumping by ~196 members), which is a change in what was captured, not
    an outage.

    A node listed under two statuses on the same day (a duplicated row in
    the snapshot) can't be read reliably, so it's skipped and returned
    separately as `ambiguous` for the console log.

    Returns (changes, ambiguous) where changes is a list of
    (day, name, ip, old_status, new_status).
    """
    changes, ambiguous = [], []
    prev_day = None
    for day in DAYS:
        if day not in states:
            continue
        if prev_day is not None:
            prev, cur = states[prev_day], states[day]
            for key in sorted(set(prev) & set(cur)):
                before, after = prev[key], cur[key]
                if before == after:
                    continue
                if len(before) > 1 or len(after) > 1:
                    ambiguous.append((day, key))
                    continue
                name, ip, _port = key
                changes.append((day, name, ip, next(iter(before)), next(iter(after))))
        prev_day = day
    return changes, ambiguous


def format_node_change(name, ip, old, new):
    text = f"{name} ({ip}) — {old.title()} → {new.title()}"
    if new == "AVAILABLE":
        text += " (recovered)"
    return text


# --------------------------------------------------------------------------
# SSL / device certificates (from Certificate.xlsx)
# --------------------------------------------------------------------------

CERT_SHEET = "LTM"
# Which device in Certificate.xlsx each site's certificate slides come from
# (matched by the start of the "Node Name" cell). Same devices as the
# user's hand-built 7-13 Sep 2026 deck.
CERT_NODES = {
    "IKY": "NG-IKY-LTM-ASM-F5BIGIP-01",
    "VGC": "NG-VGC-LTM-ASM-F5BIGIP-01",
    "OJT": "NG-OJT-LTM-ASM-F5BIGIP-01",
}
CERT_NAME_FILL = "F8CBAD"   # the peach fill on the certificate-name/contents cells
_CERT_DATE_RE = re.compile(r"([A-Za-z]{3,9})\.?\s+(\d{1,2})(?:st|nd|rd|th)?\s*,?\s*(\d{4})(?!\d)")


@dataclass
class CertRow:
    name: str
    contents: str
    expiry_text: str
    expiry: object          # date, or None if the text couldn't be read
    fill: object            # "RRGGBB" copied from the workbook cell, or None


def parse_cert_expiry(value):
    """
    Expiry cells are free text in several formats: "Dec 13, 2026",
    "Apr 22 2026", "Aug 1st,2036", "Nov 8,2028", and bundles with a range
    ("Oct 19, 2025 - Oct 6, 2046" - the expiry is the END of the range).
    Some are cut off in the workbook itself ("Jan 15, 203") - those return
    None rather than a guess.
    """
    if isinstance(value, datetime):
        return value.date()
    if isinstance(value, date):
        return value
    if not isinstance(value, str):
        return None
    part = value.split(" - ")[-1]
    m = _CERT_DATE_RE.search(part)
    if not m:
        return None
    month = _resolve_month(m.group(1))
    if month is None:
        return None
    try:
        return date(int(m.group(3)), month, int(m.group(2)))
    except ValueError:
        return None


def _cell_fill_rgb(cell):
    fill = cell.fill
    if fill is None or fill.fill_type != "solid":
        return None
    color = fill.fgColor
    if color is None or color.type != "rgb" or not isinstance(color.rgb, str):
        return None
    return color.rgb[-6:].upper()


def read_certificates(path, year):
    """
    Reads Certificate.xlsx (LTM sheet) and returns, per site:
      {"node", "checked", "engineer", "ssl": [CertRow], "device": [CertRow],
       "unreadable": [CertRow]}
    Confirmed with the user 28-Sep-2026:
      - only certificates that expired, or will expire, in `year` (the
        report week's year) are shown;
      - expiry colours are copied from the workbook as the engineer filled
        them in, not recalculated;
      - the device-certificate slide shows only the first device cert row
        (server.crt), again only if it expires in `year`.
    Columns are located by header text, not fixed letters.
    """
    wb = openpyxl.load_workbook(path, data_only=True)
    ws = wb[CERT_SHEET] if CERT_SHEET in wb.sheetnames else next(
        (s for s in wb.worksheets if CERT_SHEET.lower() in s.title.lower()), None)
    if ws is None:
        raise ValueError(f"No '{CERT_SHEET}' sheet found in {path}")

    hdr_row, cols = None, {}
    for r in range(1, 6):
        for c in range(1, ws.max_column + 1):
            v = ws.cell(r, c).value
            if not isinstance(v, str):
                continue
            u = v.upper()
            if "NODE NAME" in u:
                cols["node"] = c
                hdr_row = r
            elif u.startswith("DATE"):
                cols["date"] = c
            elif "ENGINEER" in u:
                cols["engineer"] = c
            elif "SSL" in u and "ssl" not in cols:
                cols["ssl"] = c
            elif "DEVICE CERTIFICATE" in u and "device" not in cols:
                cols["device"] = c
        if hdr_row:
            break
    missing = [k for k in ("node", "ssl", "device") if k not in cols]
    if hdr_row is None or missing:
        raise ValueError(f"Couldn't find the {missing or 'header'} column(s) in {path}")

    # Each device is a block of rows: the node name/date/engineer sit in the
    # (merged) first row, certificates run down the rows below it.
    blocks, current = [], None
    for r in range(hdr_row + 1, ws.max_row + 1):
        node = ws.cell(r, cols["node"]).value
        if isinstance(node, str) and node.strip():
            checked = ws.cell(r, cols["date"]).value if "date" in cols else None
            current = {
                "node": node.strip(),
                "checked": checked.date() if isinstance(checked, datetime) else checked,
                "engineer": ws.cell(r, cols["engineer"]).value if "engineer" in cols else None,
                "ssl": [], "device": [],
            }
            blocks.append(current)
        if current is None:
            continue
        for kind in ("ssl", "device"):
            c = cols[kind]
            name = ws.cell(r, c).value
            if name is None or not str(name).strip():
                continue
            exp_cell = ws.cell(r, c + 2)
            current[kind].append(CertRow(
                name=str(name).strip(),
                contents=str(ws.cell(r, c + 1).value or "").strip(),
                expiry_text=str(exp_cell.value or "").strip(),
                expiry=parse_cert_expiry(exp_cell.value),
                fill=_cell_fill_rgb(exp_cell),
            ))

    result = {}
    for site, prefix in CERT_NODES.items():
        block = next((b for b in blocks if b["node"].upper().startswith(prefix.upper())), None)
        if block is None:
            continue
        unreadable = [c for c in block["ssl"] + block["device"][:1] if c.expiry is None]
        result[site] = {
            "node": block["node"],
            "checked": block["checked"],
            "engineer": block["engineer"],
            "ssl": [c for c in block["ssl"] if c.expiry and c.expiry.year == year],
            "device": [c for c in block["device"][:1] if c.expiry and c.expiry.year == year],
            "unreadable": unreadable,
        }
    return result


# --------------------------------------------------------------------------
# PPTX filling
# --------------------------------------------------------------------------

def hide_all_data_labels(chart):
    """
    Force-hide every data label on a chart at every XML level.

    Some charts (this doughnut in particular) have label-visibility flags
    baked in at more than one level - a per-series <c:dLbls> AND
    individual per-point <c:dLbl> overrides for specific slices - and
    python-pptx's high-level `data_labels` property only reaches one of
    those levels. A per-point override (e.g. showPercent=1 on just one
    slice) wins over a group-level setting, so setting the group property
    alone can silently fail to hide a specific slice's label. This walks
    the raw XML and blanks every show* flag everywhere in the chart, and
    deletes per-point <c:dLbl> override elements entirely, so nothing is
    left that could re-enable a label.
    """
    chartSpace = chart._chartSpace
    show_tags = [
        qn("c:showLegendKey"), qn("c:showVal"), qn("c:showCatName"),
        qn("c:showSerName"), qn("c:showPercent"), qn("c:showBubbleSize"),
    ]
    for dLbls in chartSpace.findall(".//" + qn("c:dLbls")):
        for child in list(dLbls):
            if child.tag == qn("c:dLbl"):
                dLbls.remove(child)
            elif child.tag in show_tags:
                child.set("val", "0")


def apply_font(text_frame, size=None, bold=None):
    """Set MTN Brighter Sans (and optionally size/bold) on every run in a text frame."""
    for para in text_frame.paragraphs:
        for run in para.runs:
            run.font.name = FONT_NAME
            if size is not None:
                run.font.size = size
            if bold is not None:
                run.font.bold = bold


def set_run_text(shape, new_text):
    """Replace text while preserving the first run's formatting, then force the font."""
    tf = shape.text_frame
    first_para = tf.paragraphs[0]
    if first_para.runs:
        first_para.runs[0].text = new_text
        for extra in first_para.runs[1:]:
            extra.text = ""
    else:
        first_para.text = new_text
    apply_font(tf)


def fill_footer(slide, current: WeekMetrics):
    for shape in slide.shapes:
        if shape.has_text_frame and "CONFIDENTIAL" in shape.text_frame.text:
            tf = shape.text_frame
            full = tf.paragraphs[0].runs
            if full:
                joined = "".join(r.text for r in full)
                new_text = re.sub(
                    r"\[.*?\]",
                    f"[{current.start.strftime('%d %b').upper()} - {current.end.strftime('%d %b %Y').upper()}]",
                    joined,
                )
                full[0].text = new_text
                for extra in full[1:]:
                    extra.text = ""
                apply_font(tf)


def fill_dashboard_slide(slide, current: WeekMetrics, history: list, is_combined: bool):
    """
    Fills the 4 KPI callouts + 4 charts on one dashboard slide.
    `current` / `history` are already scoped to the right site (or to all
    sites combined, for the final TOTAL dashboard).
    is_combined controls the WAF False Positive chart:
      - combined slide: 3 bars, one per site (as before)
      - per-site slide: 4 bars, that site's FP count for each of the last
        4 weeks (a trend, not a site comparison)
    """
    pool_pct = (current.pool_available / current.pool_total * 100) if current.pool_total else 0

    kpi_targets = {
        "Requests Inspected": f"{current.requests_inspected:,}",
        "WAF Blocks": f"{current.waf_blocks:,}",
        "Pool Member Avail.": f"{pool_pct:.0f}%",
        "F5 HA Uptime": "100%",
    }
    shapes = list(slide.shapes)

    # The WAF False Positive chart is dropped from every dashboard slide -
    # it isn't tracked reliably enough to report - and the Pool Member
    # Status doughnut next to it is stretched to take over the freed
    # space instead of leaving a gap.
    fp_shape, dough_shape = None, None
    for shape in shapes:
        if shape.has_chart and shape.chart.has_title:
            title = shape.chart.chart_title.text_frame.text
            if "False Positive" in title:
                fp_shape = shape
            elif "Pool Member Status" in title:
                dough_shape = shape
    if fp_shape is not None and dough_shape is not None:
        dough_shape.width = (fp_shape.left + fp_shape.width) - dough_shape.left
        fp_shape._element.getparent().remove(fp_shape._element)
        shapes = list(slide.shapes)

    for i, shape in enumerate(shapes):
        if shape.has_text_frame:
            label = shape.text_frame.text.strip()
            if label in kpi_targets:
                # the VALUE textbox is the one immediately preceding the label textbox
                value_shape = shapes[i - 1]
                set_run_text(value_shape, kpi_targets[label])

        if shape.has_chart:
            chart = shape.chart
            title = chart.chart_title.text_frame.text if chart.has_title else ""

            if "Attack Category" in title:
                # Top 10 by count, then displayed smallest -> largest
                top10 = sorted(current.attack_counts.items(), key=lambda kv: kv[1], reverse=True)[:10]
                top10_ascending = list(reversed(top10))
                if not top10_ascending:
                    # e.g. that site's sheet is missing for this week - keep
                    # the chart (empty) instead of crashing the whole report;
                    # validate_outputs.py FAILs loudly about the missing site
                    top10_ascending = [("No data for this week", 0)]
                cd = CategoryChartData()
                cd.categories = [k for k, _ in top10_ascending]
                cd.add_series("Blocks", [v for _, v in top10_ascending])
                chart.replace_data(cd)

            elif "Request Rate vs Block Rate" in title:
                cd = CategoryChartData()
                cd.categories = [f"Week {h.start.isocalendar()[1]}" for h in history]
                cd.add_series("Requests/min (k)", [h.requests_inspected for h in history])
                cd.add_series("Blocks/min", [h.waf_blocks for h in history])
                chart.replace_data(cd)

            elif "Pool Member Status" in title:
                total = current.pool_total or 1  # avoid div/0
                pct_available = current.pool_available / total * 100
                pct_unavailable = current.pool_unavailable / total * 100
                pct_offline = current.pool_offline / total * 100
                pct_unknown = current.pool_unknown / total * 100

                # Bake each percentage into the legend text itself (via the
                # category name) rather than relying on on-slice data
                # labels: this is a doughnut chart, and thin slices (e.g.
                # "Unknown" at ~1%) have almost no room at their center,
                # which causes the renderer to wrap long label text one
                # letter per line. Legend text is plain text and can't
                # garble the same way, so every percentage stays legible
                # regardless of how small its slice is.
                cd = CategoryChartData()
                cd.categories = [
                    f"Available ({pct_available:.0f}%)",
                    f"Unavailable ({pct_unavailable:.0f}%)",
                    f"Offline ({pct_offline:.0f}%)",
                    f"Unknown ({pct_unknown:.0f}%)",
                ]
                cd.add_series("Pool Health", [
                    pct_available, pct_unavailable, pct_offline, pct_unknown,
                ])
                chart.replace_data(cd)
                # Turn off on-slice labels entirely - the legend now carries
                # all the percentage information. Done at the raw-XML level
                # since this chart has per-point label overrides that a
                # plot-level python-pptx setting alone doesn't reach.
                hide_all_data_labels(chart)



def _footer_top(prs, slide):
    return min(
        (s.top for s in slide.shapes if s.has_text_frame and "CONFIDENTIAL" in s.text_frame.text),
        default=int(prs.slide_height * 0.9),
    )


def _add_label_value_para(tf, label, value, first=False):
    """One paragraph with a bold 'Label:' run followed by a plain value run."""
    para = tf.paragraphs[0] if first else tf.add_paragraph()
    r1 = para.add_run()
    r1.text = label
    r1.font.bold = True
    r2 = para.add_run()
    r2.text = value
    return para


def fill_exec_summary_slide(prs, slide, current: WeekMetrics, site_label=None,
                            site_breakdown=None, node_note=None):
    """
    site_breakdown: {site: WeekMetrics} - only for the combined (TOTAL)
        slide; adds an IKY / VGC / OJT column per site to the pool table
        next to the combined Count and Percentage (user's 7-13 Sep 2026 edit).
    node_note: text shown after a bold "Node changes:" label on per-site
        slides (e.g. "None" or "See next slide"); None = no line at all.
    """
    pool_pct = (current.pool_available / current.pool_total * 100) if current.pool_total else 0

    footer_top = _footer_top(prs, slide)
    content_top = int(prs.slide_height * 0.22)
    left = Emu(int(prs.slide_width * 0.08))
    width = Emu(int(prs.slide_width * 0.84))

    # --- Summary text: labels bold, values plain (user's 7-13 Sep edit) ---
    n_lines = 4 + (1 if node_note else 0)
    text_height = int(Inches(0.1) + n_lines * Inches(0.32))
    box = slide.shapes.add_textbox(Emu(left), Emu(content_top), Emu(width), Emu(text_height))
    tf = box.text_frame
    tf.word_wrap = True
    tf.paragraphs[0].add_run().text = "We had 100% Up time on the active and standby nodes."
    _add_label_value_para(tf, "Requests Inspected:", f" {current.requests_inspected:,}")
    _add_label_value_para(tf, "WAF Blocks:", f" {current.waf_blocks:,}")
    _add_label_value_para(tf, "Pool Availability:", f" {pool_pct:.0f}%")
    if node_note:
        _add_label_value_para(tf, "Node changes:", f" {node_note}")
    apply_font(tf)

    # --- Pool Member Status table (below the text), with a Percentage
    # column alongside the raw counts ---
    table_top = content_top + text_height + int(prs.slide_height * 0.03)
    rows = [
        ("Available", "pool_available"),
        ("Unavailable", "pool_unavailable"),
        ("Offline", "pool_offline"),
        ("Unknown", "pool_unknown"),
    ]
    n_rows = len(rows) + 2  # header + 4 status rows + TOTAL row
    total = current.pool_total or 1  # avoid div/0

    if site_breakdown:
        site_cols = [s for s in SITES if s in site_breakdown]
        headers = ["Pool Member Status"] + site_cols + ["Count", "Percentage"]
        table_width = int(prs.slide_width * 0.70)
        ratios = [0.30] + [0.125] * len(site_cols) + [0.14, 0.185]
    else:
        site_cols = []
        headers = ["Pool Member Status", "Count", "Percentage"]
        table_width = int(prs.slide_width * 0.55)
        ratios = [0.45, 0.25, 0.30]
    n_cols = len(headers)
    table_height = min(int(prs.slide_height * 0.32), footer_top - table_top - int(Inches(0.1)))

    graphic_frame = slide.shapes.add_table(n_rows, n_cols, Emu(left), Emu(table_top),
                                           Emu(table_width), Emu(table_height))
    table = graphic_frame.table
    ratio_sum = sum(ratios)
    for c, ratio in enumerate(ratios):
        table.columns[c].width = Emu(int(table_width * ratio / ratio_sum))
    for c, h in enumerate(headers):
        table.cell(0, c).text = h
    for r, (label, attr) in enumerate(rows, start=1):
        value = getattr(current, attr)
        cells = [label] + [f"{getattr(site_breakdown[s], attr):,}" for s in site_cols] \
            + [f"{value:,}", f"{value / total * 100:.2f}%"]
        for c, text in enumerate(cells):
            table.cell(r, c).text = text
    totals = ["TOTAL"] + [f"{site_breakdown[s].pool_total:,}" for s in site_cols] \
        + [f"{current.pool_total:,}", "100%"]
    for c, text in enumerate(totals):
        table.cell(n_rows - 1, c).text = text
    for r in range(n_rows):
        for c in range(n_cols):
            apply_font(table.cell(r, c).text_frame, size=Pt(12), bold=(r == 0 or r == n_rows - 1))


def order_slides(prs, ordered_slides):
    """
    Physically puts the slides of `prs` into the order given by
    `ordered_slides` (a list of Slide objects). python-pptx has no
    built-in "move slide" API - the slide order lives in the <p:sldIdLst>
    element as a sequence of <p:sldId> references, so this reorders that
    element directly rather than moving slide content around.
    """
    sldIdLst = prs.slides._sldIdLst
    by_part = {prs.part.related_part(el.rId): el for el in sldIdLst}
    new_order = [by_part[s.part] for s in ordered_slides]
    for el in list(sldIdLst):
        sldIdLst.remove(el)
    for el in new_order:
        sldIdLst.append(el)


def clone_header_slide(prs, src_slide, title, subtitle):
    """
    Adds a new slide carrying only the header bar / logo / title /
    subtitle / footer of `src_slide` (an untouched template slide), with
    the title and subtitle text replaced. Used for the node-change and
    certificate slides, which have no slide of their own in the template.
    """
    new = prs.slides.add_slide(src_slide.slide_layout)
    for shp in list(new.shapes):
        shp._element.getparent().remove(shp._element)
    for shape in src_slide.shapes:
        if shape.has_chart or shape.has_table:
            continue
        el = copy.deepcopy(shape._element)
        # pictures (the logo) point at an image part via a relationship id
        # that only exists on the source slide - re-link it on the new one
        for blip in el.iter(qn("a:blip")):
            old_rid = blip.get(qn("r:embed"))
            if old_rid:
                image_part = src_slide.part.related_part(old_rid)
                blip.set(qn("r:embed"), new.part.relate_to(image_part, RT.IMAGE))
        new.shapes._spTree.append(el)
    for shape in new.shapes:
        if not shape.has_text_frame:
            continue
        text = shape.text_frame.text
        if text.startswith("F5 BIG-IP"):
            set_run_text(shape, title)
        elif text.startswith("Attack category"):
            set_run_text(shape, subtitle)
    return new


def _content_box(prs, slide):
    """(left, top, width, height) of the usable area between header and footer."""
    left = int(prs.slide_width * 0.08)
    top = int(prs.slide_height * 0.22)
    return left, top, int(prs.slide_width * 0.84), _footer_top(prs, slide) - top - int(Inches(0.05))


NODE_LINES_PER_SLIDE = 13


def build_node_change_slides(prs, template_exec_slide, site, changes):
    """
    One or more slides listing that site's node changes, grouped by day,
    in the same layout as the user's hand-built IKY slide (bold "Node
    changes" heading, day name, then one line per change).
    """
    lines = [("Node changes", True)]
    last_day = None
    for day, name, ip, old, new in changes:
        if day != last_day:
            if last_day is not None:
                lines.append(("", False))
            lines.append((day.title(), True))
            last_day = day
        lines.append((format_node_change(name, ip, old, new), False))

    pages = [lines[i:i + NODE_LINES_PER_SLIDE] for i in range(0, len(lines), NODE_LINES_PER_SLIDE)]
    slides = []
    for p, page in enumerate(pages, start=1):
        suffix = f" ({p}/{len(pages)})" if len(pages) > 1 else ""
        slide = clone_header_slide(
            prs, template_exec_slide,
            f"F5 BIG-IP — {site} NODE CHANGES{suffix}",
            "Pool member status changes during the week (each day compared with the day before)",
        )
        left, top, width, height = _content_box(prs, slide)
        box = slide.shapes.add_textbox(Emu(left), Emu(top), Emu(width), Emu(height))
        tf = box.text_frame
        tf.word_wrap = True
        for i, (text, bold) in enumerate(page):
            para = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
            run = para.add_run()
            run.text = text
            run.font.bold = bold
        apply_font(tf, size=Pt(16))
        slides.append(slide)
    return slides


CERT_ROWS_PER_SLIDE = 20


def _set_cell(cell, text, size, fill=None, bold=True):
    cell.text = text
    cell.margin_top = cell.margin_bottom = Emu(int(Inches(0.02)))
    cell.margin_left = cell.margin_right = Emu(int(Inches(0.06)))
    if fill:
        cell.fill.solid()
        cell.fill.fore_color.rgb = RGBColor.from_string(fill)
    for para in cell.text_frame.paragraphs:
        for run in para.runs:
            run.font.name = FONT_NAME
            run.font.size = size
            run.font.bold = bold
            if fill:
                run.font.color.rgb = RGBColor(0, 0, 0)


def build_cert_slides(prs, template_exec_slide, site, kind, certs, info, year):
    """
    Certificate table slide(s) for one site. kind = "ssl" or "device".
    Columns match the user's slides (name / contents / expiration), with
    the name + contents cells in the same peach fill and the expiration
    cell in whatever colour the engineer gave it in Certificate.xlsx. Long
    lists are split evenly over as many slides as needed so the text stays
    readable (the hand-pasted 40-45 row tables were ~5pt).
    """
    if not certs:
        return []
    heading = "SSL CERTIFICATE" if kind == "ssl" else "DEVICE CERTIFICATE"
    checked = info.get("checked")
    checked_txt = checked.strftime("%d %b %Y") if isinstance(checked, date) else str(checked or "?")
    what = "Traffic certificates" if kind == "ssl" else "Device certificate"
    subtitle = f"{what} expired or expiring in {year} · {info['node']} · checked {checked_txt}"

    n_pages = math.ceil(len(certs) / CERT_ROWS_PER_SLIDE)
    per_page = math.ceil(len(certs) / n_pages)
    slides = []
    for p in range(n_pages):
        page = certs[p * per_page:(p + 1) * per_page]
        suffix = f" ({p + 1}/{n_pages})" if n_pages > 1 else ""
        slide = clone_header_slide(prs, template_exec_slide,
                                   f"F5 BIG-IP — {site} {heading}{suffix}", subtitle)
        left, top, width, height = _content_box(prs, slide)
        left = int(Inches(0.45))
        width = int(prs.slide_width - 2 * Inches(0.45))
        n_rows = len(page) + 1
        row_h = min(int(Inches(0.32)), int(height / n_rows))
        size = Pt(11) if n_rows <= 12 else Pt(10) if n_rows <= 16 else Pt(9)
        gf = slide.shapes.add_table(n_rows, 3, Emu(left), Emu(top), Emu(width), Emu(row_h * n_rows))
        table = gf.table
        for c, ratio in enumerate([0.40, 0.42, 0.18]):
            table.columns[c].width = Emu(int(width * ratio))
        for r in range(n_rows):
            table.rows[r].height = Emu(row_h)
        for c, h in enumerate(["Certificate", "Contents", "Expiration"]):
            _set_cell(table.cell(0, c), h, size)
        for r, cert in enumerate(page, start=1):
            _set_cell(table.cell(r, 0), cert.name, size, fill=CERT_NAME_FILL)
            _set_cell(table.cell(r, 1), cert.contents, size, fill=CERT_NAME_FILL)
            _set_cell(table.cell(r, 2), cert.expiry_text, size, fill=cert.fill or "FFFFFF")
        slides.append(slide)
    return slides


def fill_pptx(template_path, out_path, combined_current, combined_history, per_site,
              node_changes=None, certs=None, cert_year=None):
    """
    per_site: dict {site: (current: WeekMetrics, history: list)} for
    IKY/VGC/OJT, each already scoped to that site only.
    combined_current/combined_history: merged across all 3 sites.
    node_changes: {site: [(day, name, ip, old, new), ...]} or None to skip
        the node-change lines/slides entirely.
    certs: output of read_certificates() or None to skip certificate slides.
    """
    prs = Presentation(template_path)
    slides = list(prs.slides)
    # untouched copy of a per-site exec slide's header, for cloning new slides
    # from - taken before anything is filled in
    template_exec = slides[2]
    extra = {site: [] for site in SITES}
    for site in SITES:
        if node_changes is not None and node_changes.get(site):
            extra[site] += build_node_change_slides(prs, template_exec, site, node_changes[site])
        if certs and site in certs:
            extra[site] += build_cert_slides(prs, template_exec, site, "ssl",
                                             certs[site]["ssl"], certs[site], cert_year)
            extra[site] += build_cert_slides(prs, template_exec, site, "device",
                                             certs[site]["device"], certs[site], cert_year)

    week_str = f"{combined_current.start.strftime('%B')} {combined_current.start.day}th - " \
               f"{combined_current.end.strftime('%B')} {combined_current.end.day}th, {combined_current.end.year}"

    # ---- Slide 1: cover week label (no-op if the template doesn't have
    # this text box - e.g. a section-divider-only cover) ----
    for shape in slides[0].shapes:
        if shape.has_text_frame and shape.text_frame.text.startswith(("July", "June",
                                                                       "August", "January",
                                                                       "February", "March",
                                                                       "April", "May",
                                                                       "September", "October",
                                                                       "November", "December")):
            tf = shape.text_frame
            if len(tf.paragraphs) >= 1 and tf.paragraphs[0].runs:
                tf.paragraphs[0].runs[0].text = week_str
                apply_font(tf)
            break

    # ---- Slides 2-7: 3 per-site (dashboard, exec summary) pairs ----
    slide_positions = [(1, 2, "IKY"), (3, 4, "VGC"), (5, 6, "OJT")]
    for dash_idx, exec_idx, site in slide_positions:
        cur, hist = per_site[site]
        fill_dashboard_slide(slides[dash_idx], cur, hist, is_combined=False)
        fill_footer(slides[dash_idx], cur)
        node_note = None
        if node_changes is not None:
            node_note = "See next slide" if node_changes.get(site) else "None this week"
        fill_exec_summary_slide(prs, slides[exec_idx], cur, site_label=site, node_note=node_note)
        fill_footer(slides[exec_idx], cur)
        for s in extra[site]:
            fill_footer(s, cur)

    # ---- Slides 8-9: combined dashboard + exec summary ----
    fill_dashboard_slide(slides[7], combined_current, combined_history, is_combined=True)
    fill_footer(slides[7], combined_current)
    fill_exec_summary_slide(prs, slides[8], combined_current, site_label=None,
                            site_breakdown={s: per_site[s][0] for s in SITES})
    fill_footer(slides[8], combined_current)

    # ---- Final order: cover, then the combined (TOTAL) dashboard + exec
    # summary, then per site: dashboard, exec summary, node changes, SSL
    # certificates, device certificate. Each slide's title text ("F5 BIG-IP
    # - IKY ...") is baked into that specific slide rather than generated,
    # so this moves the actual slides (not just which data gets filled
    # where) - otherwise a slide keeps its old title with new data on it.
    ordered = [slides[0], slides[7], slides[8]]
    for dash_idx, exec_idx, site in slide_positions:
        ordered += [slides[dash_idx], slides[exec_idx]] + extra[site]
    order_slides(prs, ordered)

    prs.save(out_path)


def load_node_changes_and_certs(wb, weeks, target_key, certs_path, log=print):
    """
    Reads the week's node changes (from the source workbook) and the
    certificate lists (from Certificate.xlsx), printing a short log plus
    [WARN] lines for anything that needs a human look. Returns
    (node_changes, certs) ready for fill_pptx(); certs is None if the
    certificate file isn't there (the deck is still built, just without
    certificate slides).
    """
    import os
    start, end = target_key
    node_changes = {}
    for site in SITES:
        ref = weeks[target_key].get(site)
        if ref is None:
            node_changes[site] = []
            continue
        changes, ambiguous = compute_node_changes(read_daily_node_states(wb[ref.sheet_name]))
        node_changes[site] = changes
        log(f"  {site}: {len(changes)} node change(s)")
        for day, name, ip, old, new in changes:
            log(f"      {day.title()}: {format_node_change(name, ip, old, new)}")
        for day, (name, ip, _port) in ambiguous:
            log(f"  [WARN] {site} {day.title()}: {name} ({ip}) is listed under more than one "
                  f"status in the same day's snapshot - left off the slide, check it by hand")

    certs = None
    if certs_path and os.path.exists(certs_path):
        certs = read_certificates(certs_path, end.year)
        for site in SITES:
            info = certs.get(site)
            if info is None:
                log(f"  [WARN] {site}: device {CERT_NODES[site]} not found in {os.path.basename(certs_path)}"
                      f" - no certificate slides for {site}")
                continue
            log(f"  {site}: {len(info['ssl'])} SSL + {len(info['device'])} device certificate(s) "
                  f"expiring in {end.year} ({info['node']})")
            for c in info["unreadable"]:
                log(f"  [WARN] {site}: couldn't read the expiry date '{c.expiry_text}' for {c.name}"
                      f" - left off the slide, check Certificate.xlsx")
            checked = info.get("checked")
            if isinstance(checked, date) and checked < start - timedelta(days=7):
                log(f"  [WARN] {site}: certificate check is dated {checked:%d %b %Y}, more than a week"
                      f" before this report's week - is Certificate.xlsx up to date?")
    else:
        log(f"  [WARN] {certs_path or 'Certificate.xlsx'} not found - certificate slides skipped")
    return node_changes, certs


# --------------------------------------------------------------------------
# Excel summary output
# --------------------------------------------------------------------------

def write_summary_xlsx(out_path, current: WeekMetrics, history: list):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "KPI Summary"

    ws.append(["Week", current.label])
    ws.append(["Requests Inspected", current.requests_inspected])
    ws.append(["WAF Blocks", current.waf_blocks])
    ws.append(["Pool Available", current.pool_available])
    ws.append(["Pool Total", current.pool_total])
    pool_pct = (current.pool_available / current.pool_total * 100) if current.pool_total else 0
    ws.append(["Pool Availability %", round(pool_pct, 2)])
    ws.append([])

    ws.append(["Attack Type", "Violation Count (merged, all sites)"])
    for atype, count in sorted(current.attack_counts.items(), key=lambda kv: kv[1], reverse=True):
        ws.append([atype, count])

    ws2 = wb.create_sheet("FP by Site")
    ws2.append(["Site", "FP Count"])
    for site, count in current.fp_by_site.items():
        ws2.append([site, count])

    ws3 = wb.create_sheet("4-Week Trend")
    ws3.append(["Week", "Requests Inspected", "WAF Blocks", "Pool Availability %"])
    for h in history:
        pct = (h.pool_available / h.pool_total * 100) if h.pool_total else 0
        ws3.append([h.label, h.requests_inspected, h.waf_blocks, round(pct, 2)])

    wb.save(out_path)


# --------------------------------------------------------------------------
# Main
# --------------------------------------------------------------------------

def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--source", required=True)
    ap.add_argument("--template", required=True)
    ap.add_argument("--out-pptx", required=True)
    ap.add_argument("--out-xlsx", default=None, help="optional KPI summary workbook")
    ap.add_argument("--week-end", default=None, help="YYYY-MM-DD, optional")
    ap.add_argument("--certs", default="Certificate.xlsx",
                    help="Certificate workbook for the SSL/device certificate slides (skipped if missing)")
    args = ap.parse_args()

    week_end = date.fromisoformat(args.week_end) if args.week_end else None

    wb, weeks, target_key = load_source_workbook(args.source, week_end, n_weeks=4)
    week_keys = last_n_weeks(weeks, target_key, n=4)

    # Combined (all 3 sites merged) - used for the final TOTAL dashboard/summary
    combined_history = [compute_week_metrics(wb, k, weeks[k]) for k in week_keys]
    combined_current = combined_history[-1]

    # Per-site (IKY/VGC/OJT individually) - used for the 3 per-site dashboards
    per_site = {}
    for site in SITES:
        hist = [compute_week_metrics(wb, k, weeks[k], sites=[site]) for k in week_keys]
        per_site[site] = (hist[-1], hist)

    print(f"Target week: {combined_current.label}")
    print(f"History weeks used for trend: {[h.label for h in combined_history]}")
    print(f"Requests Inspected (combined): {combined_current.requests_inspected:,}")
    print(f"WAF Blocks (combined): {combined_current.waf_blocks:,}")
    print(f"Pool Available/Total (combined): {combined_current.pool_available}/{combined_current.pool_total}")
    for site in SITES:
        cur, _ = per_site[site]
        print(f"  {site}: Requests Inspected={cur.requests_inspected:,}  WAF Blocks={cur.waf_blocks:,}")

    node_changes, certs = load_node_changes_and_certs(wb, weeks, target_key, args.certs)
    fill_pptx(args.template, args.out_pptx, combined_current, combined_history, per_site,
              node_changes=node_changes, certs=certs, cert_year=target_key[1].year)
    if args.out_xlsx:
        write_summary_xlsx(args.out_xlsx, combined_current, combined_history)
    print(f"\nWrote {args.out_pptx}")
    if args.out_xlsx:
        print(f"Wrote {args.out_xlsx}")


if __name__ == "__main__":
    main()
