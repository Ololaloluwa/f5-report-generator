#!/usr/bin/env python3
"""
MTN F5 WAF Weekly Report - Excel Automation
=============================================

Fills MTN_F5_WAF_Weekly_Report_Temp.xlsx from the F5 source workbook
(MTN_F5_ATTACKS_SOURCE.xlsx). Produces:
  - Sheet 1 "Weekly Report Data": updates the Report Range date only.
  - Sheets 2-4 "Detailed Vuln Matrix <SITE>": per-site attack-type rows
    enriched with Violation Name(s)/Title/Description from a fixed
    reference lookup table, Severity hardcoded to "High", Associated IP
    pulled straight from the source's Associated IPs column.
  - Sheets 5-7 "Summary Dashboard <SITE>": 4 side-by-side weekly blocks
    (current + up to 3 prior weeks), each sorted by Violation Count
    descending with a TOTAL row, a standalone "X more/less compared to
    last week" label per block (not tied to any attack-type row), a
    Severity Count table (Critical always 0, High = number of attack
    types that week), and the site's 3 charts updated in place (their
    cell references are edited, not rebuilt, so template styling/colors
    are preserved).

Usage:
    python generate_weekly_excel.py \
        --source MTN_F5_ATTACKS_SOURCE.xlsx \
        --template MTN_F5_WAF_Weekly_Report_Temp.xlsx \
        --out OUT_WEEKLY_EXCEL.xlsx \
        [--week-end 2026-07-19]

Known assumptions (flagged explicitly - change here if wrong):
  - "Violation Type" column is hardcoded to "Input Violation" for every
    row (100% of the example data used this single value).
  - If an attack type isn't found in the lookup table (even after
    stripping a parenthetical abbreviation), Violation Name/Title/
    Description are written as "N/A".
  - Associated IP is copied as-is from the source's "Associated IPs"
    column for that attack type/site/week; if empty, "N/A" is written.
"""

import argparse
import copy
import re
from datetime import date

import openpyxl
from openpyxl.styles import Font, PatternFill, Border, Alignment
from openpyxl.utils import get_column_letter

from generate_weekly_report import (
    SITES,
    group_sheets_by_week,
    pick_target_week,
    last_n_weeks,
    load_source_workbook,
    week_label,
    find_row_containing,
    find_summary_row,
)

SHEET_MAP = {
    "IKY": ("Detailed Vuln Matrix IKY", "Summary Dashboard IKY"),
    "VGC": ("Detailed Vuln Matrix VGC", "Summary Dashboard VGC"),
    "OJT": ("Detailed Vuln Matrix OJT", "Summary Dashboard OJT"),
}

BLOCK_WIDTH = 8   # 5 data columns + 3 spacer columns per weekly block
MAX_HISTORY_WEEKS = 4
TOP_N_INCIDENT_CHART = 14  # matches the original template's row range

VIOLATION_TYPE_CONST = "Input Violation"


# --------------------------------------------------------------------------
# Attack Type -> Violation Name(s) / Title / Description lookup table
# (as provided by the user)
# --------------------------------------------------------------------------

ATTACK_LOOKUP_RAW = [
    ("Server-Side Code Injection (SSCI)", ["VIOL_SERVER_SIDE_CODE_INJECTION"],
     "Server-Side Code Injection Attempt",
     "Request contained patterns indicating execution of server-side code (e.g., command injection / code execution attempt)."),
    ("Forceful Browsing", ["VIOL_URL"],
     "Illegal URL Access Attempt",
     "Request attempted to access a URL or resource that is not defined or permitted by the security policy. This indicates a possible forced browsing attempt where an attacker is trying to discover or access hidden, restricted, or unauthorized application paths."),
    ("Information Leakage", ["VIOL_INFORMATION_LEAKAGE"],
     "Information Leakage Detected",
     "The application response exposed sensitive information such as server details, system paths, software versions, error messages, configuration data, or debugging information that could assist an attacker in reconnaissance or further exploitation attempts."),
    ("HTTP Parser Attack", ["VIOL_COOKIE_RFC_NON_COMPLIANT"],
     "Cookie not RFC compliant",
     "Cookie header does not conform to RFC standards."),
    ("Detection Evasion", ["VIOL_EVASION"],
     "Evasion Technique Detected",
     "Request used encoding, obfuscation, malformed data, or other techniques designed to bypass security inspection mechanisms."),
    ("Path Traversal", ["VIOL_URL"],
     "Illegal URL Access Attempt",
     "Request attempted to access a URL or resource that is not defined or permitted by the security policy. This indicates a possible forced browsing attempt where an attacker is trying to discover or access hidden, restricted, or unauthorized application paths."),
    ("Vulnerability Scan", ["VIOL_INFORMATION_LEAKAGE"],
     "Information Leakage Detected",
     "The application response exposed sensitive information such as server details, system paths, software versions, error messages, configuration data, or debugging information that could assist an attacker in reconnaissance or further exploitation attempts."),
    ("Predictable Resource Location", ["VIOL_URL"],
     "Illegal URL Access Attempt",
     "Request attempted to access a predictable or hidden resource location (e.g., administrative pages, backup files, or restricted directories) that is not allowed by the security policy."),
    ("Buffer Overflow", ["VIOL_URL"],
     "Illegal URL Access Attempt",
     "Request attempted abnormal or unauthorized URL access patterns that may indicate probing for vulnerable resources or buffer exhaustion conditions."),
    ("Injection Attempt", ["VIOL_ATTACK_SIGNATURE_DETECTED"],
     "Attack Signature Detected",
     "Request contained patterns matching known attack signatures associated with injection attempts, such as SQL injection, command injection, or malicious payload execution."),
    ("Cross Site Scripting (XSS)", ["VIOL_SERVER_SIDE_CODE_INJECTION"],
     "Server-Side Code Injection Attempt",
     "Request contained patterns indicating execution of server-side code (e.g., command injection / code execution attempt)."),
    ("Command Execution", ["VIOL_PARAMETER_VALUE"],
     "Illegal Parameter Value",
     "Request contained modified or manipulated parameter values that violate expected application behavior or security policy."),
    ("Other Application Attacks", ["VIOL_SERVER_SIDE_CODE_INJECTION"],
     "Server-Side Request Forgery Attempt",
     "Request attempted to manipulate server-side functionality to make unauthorized requests to internal or external resources."),
    ("Denial of Service", ["VIOL_URL"],
     "Illegal URL Access Attempt",
     "Request attempted to access a URL not defined in the security policy, indicating possible forced browsing or resource exhaustion attempt."),
    ("Abuse of Functionality", ["VIOL_ABUSE_OF_FUNCTIONALITY"],
     "Application Functionality Abuse Attempt",
     "Request indicated misuse of legitimate application functionality in a manner that could bypass intended business logic or application controls."),
    ("SQL Injection", ["VIOL_SQL_INJECTION"],
     "SQL Injection Attempt",
     "Request contained SQL syntax or payloads indicative of an attempt to manipulate backend database queries and gain unauthorized access to data."),
    ("Remote File Include", ["VIOL_REMOTE_FILE_INCLUDE"],
     "Remote File Inclusion Attempt",
     "Request attempted to load or execute a remote file or resource, potentially leading to remote code execution or unauthorized content inclusion."),
    ("Authentication/Authorization Attacks",
     ["VIOL_AUTHORIZATION_ATTACK", "VIOL_AUTHENTICATION_ATTACK"],
     "Authentication or Authorization Attack",
     "Request attempted to bypass, manipulate, or abuse authentication or authorization mechanisms to gain unauthorized access to protected resources."),
    ("XML Parser Attack", ["VIOL_XML_PARSER_ATTACK"],
     "XML Parser Attack Detected",
     "Request contained malicious XML content intended to exploit XML parser vulnerabilities, including XML External Entity (XXE) or parser manipulation attacks."),
    ("Trojan/Backdoor/Spyware", ["VIOL_TROJAN_BACKDOOR_SPYWARE"],
     "Malware Distribution Attempt",
     "Request contained indicators associated with trojans, backdoors, spyware, or other malicious software intended for delivery or execution."),
    ("Malicious File Upload", ["VIOL_MALICIOUS_FILE_UPLOAD"],
     "Malicious File Upload Attempt",
     "Request attempted to upload a file containing malicious content, executable code, or file types that could compromise the application or server."),
    ("Non-Browser Client", ["VIOL_NON_BROWSER_CLIENT"],
     "Non-Browser Client Access",
     "Request originated from a non-browser client or automated tool that may indicate scripted activity or automated interaction with the application."),
    ("Session Hijacking", ["VIOL_SESSION_HIJACKING"],
     "Session Hijacking Attempt",
     "Request exhibited characteristics consistent with an attempt to steal, reuse, or manipulate session identifiers to impersonate a legitimate user."),
    ("LDAP Injection", ["VIOL_LDAP_INJECTION"],
     "LDAP Injection Attempt",
     "Request contained LDAP query manipulation patterns that could alter directory service queries or bypass authentication controls."),
    ("JSON Parser Attack", ["VIOL_JSON_PARSER_ATTACK"],
     "JSON Parser Attack Detected",
     "Request contained malformed or malicious JSON intended to exploit parser vulnerabilities or disrupt application processing."),
    ("HTTP Response Splitting", ["VIOL_HTTP_RESPONSE_SPLITTING"],
     "HTTP Response Splitting Attempt",
     "Request attempted to inject carriage return and line feed (CRLF) characters to manipulate HTTP response headers or facilitate web cache poisoning."),
    ("Parameter Tampering", ["VIOL_PARAMETER_TAMPERING"],
     "Parameter Tampering Attempt",
     "Request contained modified or unexpected parameter values indicating an attempt to manipulate application logic or bypass input validation."),
    ("Server-Side Request Forgery (SSRF)", ["VIOL_SERVER_SIDE_REQUEST_FORGERY"],
     "Server-Side Request Forgery Attempt",
     "Request attempted to induce the server to make unauthorized requests to internal or external resources on behalf of the attacker."),
    ("Other Application Activity", ["VIOL_OTHER_APPLICATION_ACTIVITY"],
     "Suspicious Application Activity",
     "Request exhibited unusual application behavior or patterns that did not match a specific attack category but warranted further inspection."),
    ("XPath Injection", ["VIOL_XPATH_INJECTION"],
     "XPath Injection Attempt",
     "Request contained XPath query manipulation patterns intended to alter XML query execution, bypass authentication, or retrieve unauthorized data from XML-based data stores."),
    ("HTTP Request Smuggling Attack", ["VIOL_HTTP_REQUEST_SMUGGLING"],
     "HTTP Request Smuggling Attempt",
     "Request contained malformed or ambiguous HTTP headers designed to exploit inconsistencies between front-end and back-end servers, potentially enabling request smuggling or unauthorized request processing."),
    ("Cross-Site Request Forgery (CSRF)", ["VIOL_CROSS_SITE_REQUEST_FORGERY"],
     "Cross-Site Request Forgery (CSRF) Attempt",
     "Request exhibited characteristics of a Cross-Site Request Forgery attack, attempting to trick an authenticated user into performing unintended actions without their consent."),
]

# Known spelling/typo variants seen in real source data -> canonical form
ALIASES = {
    "command exection": "command execution",   # typo seen in source
    "sql-injection": "sql injection",
    "non-browser client": "non browser client",
}


def normalize(s):
    s = s.lower().strip()
    s = s.replace("-", " ").replace("/", " ")
    s = re.sub(r"\s+", " ", s)
    return s.strip()


def base_name(s):
    """Strip a trailing parenthetical abbreviation, e.g. '(SSCI)'."""
    return normalize(re.sub(r"\(.*?\)", "", s))


LOOKUP_BY_FULL = {}
LOOKUP_BY_BASE = {}
for attack_type, viol_names, title, desc in ATTACK_LOOKUP_RAW:
    entry = {"violation_names": viol_names, "title": title, "description": desc}
    full_key = normalize(attack_type)
    base_key = base_name(attack_type)
    LOOKUP_BY_FULL.setdefault(full_key, entry)
    LOOKUP_BY_BASE.setdefault(base_key, entry)


def lookup_attack(attack_type):
    n = normalize(attack_type)
    n = ALIASES.get(n, n)
    if n in LOOKUP_BY_FULL:
        return LOOKUP_BY_FULL[n]
    b = base_name(attack_type)
    b = ALIASES.get(b, b)
    if b in LOOKUP_BY_BASE:
        return LOOKUP_BY_BASE[b]
    return None


# --------------------------------------------------------------------------
# Source reading
# --------------------------------------------------------------------------

def _safe_metric(value, fallback):
    """
    The source workbook has occasionally shown broken formulas in the
    Blocked/Remediated columns of the Weekly WAF Attack Breakdown table
    (a literal '#REF!' cached value - seen so far on IKY's 17th-23rd Aug
    2026 sheet, affecting every row of that week). Writing '#REF!' itself
    into the output would silently corrupt the generated workbook with no
    error, exactly the kind of drift this pipeline is supposed to guard
    against. Every clean row we've seen has Blocked == Violation Count
    and Remediated == Violation Count (Allowed/Logged is always 0), so on
    any Excel-error value (or other non-numeric junk) we fall back to
    that known-good relationship instead of propagating the error.
    """
    if isinstance(value, str) and value.strip().startswith("#"):
        return fallback
    return value or 0


def read_full_attack_breakdown(ws, title_text="Weekly WAF Attack Breakdown", max_scan=200):
    """
    Reads the full Weekly WAF Attack Breakdown table (Attack Type,
    Violation Count, Blocked, Allowed/Logged, Remediated, Associated IPs)
    for one site's sheet, preserving source row order. Skips the TOTAL row.
    """
    title_row = find_summary_row(ws, title_text)
    if title_row is None:
        return []
    header_row = find_row_containing(ws, "Attack Type", max_row=title_row + 10)
    if header_row is None:
        return []

    rows = []
    r = header_row + 1
    scanned = 0
    while scanned < max_scan:
        attack_type = ws.cell(row=r, column=2).value
        if attack_type is None:
            break
        if str(attack_type).strip().upper() != "TOTAL":
            violation_count = ws.cell(row=r, column=3).value or 0
            rows.append({
                "attack_type": str(attack_type).strip(),
                "violation_count": violation_count,
                "blocked": _safe_metric(ws.cell(row=r, column=4).value, violation_count),
                "allowed": _safe_metric(ws.cell(row=r, column=5).value, 0),
                "remediated": _safe_metric(ws.cell(row=r, column=6).value, violation_count),
                "ips": ws.cell(row=r, column=7).value,
            })
        r += 1
        scanned += 1
    return rows


def short_week_label(start: date, end: date) -> str:
    def ord_suffix(d):
        if 11 <= d % 100 <= 13:
            return "th"
        return {1: "st", 2: "nd", 3: "rd"}.get(d % 10, "th")
    return f"{start.strftime('%B')} {start.day}{ord_suffix(start.day)} - {end.strftime('%B')} {end.day}{ord_suffix(end.day)} "


# --------------------------------------------------------------------------
# Excel writers
# --------------------------------------------------------------------------

def clear_range(ws, min_row, max_row, min_col, max_col):
    for r in range(min_row, max_row + 1):
        for c in range(min_col, max_col + 1):
            ws.cell(row=r, column=c).value = None


def capture_style(cell):
    """Snapshot a cell's visual style so it can be re-applied elsewhere.
    Must be called BEFORE the cell's own value/style is overwritten."""
    return {
        "font": copy.copy(cell.font),
        "fill": copy.copy(cell.fill),
        "border": copy.copy(cell.border),
        "alignment": copy.copy(cell.alignment),
        "number_format": cell.number_format,
    }


def apply_style(cell, style, bold=None):
    font = style["font"]
    if bold is not None:
        font = copy.copy(font)
        font.bold = bold
    cell.font = font
    cell.fill = copy.copy(style["fill"])
    cell.border = copy.copy(style["border"])
    cell.alignment = copy.copy(style["alignment"])
    cell.number_format = style["number_format"]


def reset_style(ws, min_row, max_row, min_col, max_col):
    """Blank out any leftover formatting (fill/border/font) in a range -
    used so a week with fewer rows/blocks than a previous week doesn't
    leave stray styled-but-empty cells behind, and so a range the
    template never originally styled doesn't stay visually broken."""
    blank_font, blank_fill = Font(), PatternFill()
    blank_border, blank_align = Border(), Alignment()
    for r in range(min_row, max_row + 1):
        for c in range(min_col, max_col + 1):
            cell = ws.cell(row=r, column=c)
            cell.font = blank_font
            cell.fill = blank_fill
            cell.border = blank_border
            cell.alignment = blank_align
            cell.number_format = "General"


def write_sheet1(ws, full_label):
    for row in ws.iter_rows():
        for cell in row:
            if isinstance(cell.value, str) and cell.value.strip() == "Report Range":
                ws.cell(row=cell.row, column=cell.column + 1).value = full_label
                return


def write_detailed_vuln_matrix(ws, rows):
    old_max_row = ws.max_row
    n_cols = 9

    # Capture the existing data-row style (row 2) BEFORE we touch anything,
    # so we can re-apply consistent formatting to however many rows we
    # actually write this week - fixes rows looking unstyled when this
    # week has more rows than the template did, or stray old styling
    # lingering when this week has fewer.
    style_row = 2 if old_max_row >= 2 else 1
    col_styles = [capture_style(ws.cell(row=style_row, column=c)) for c in range(1, n_cols + 1)]

    new_last_row = 1 + len(rows) + 1  # header + data rows + TOTAL row
    reset_style(ws, 2, max(old_max_row, new_last_row), 1, n_cols)
    clear_range(ws, 2, max(old_max_row, new_last_row), 1, n_cols)

    total_count = 0
    for i, row in enumerate(rows, start=1):
        r = i + 1
        entry = lookup_attack(row["attack_type"])
        violation_names = ", ".join(entry["violation_names"]) if entry else "N/A"
        title = entry["title"] if entry else "N/A"
        description = entry["description"] if entry else "N/A"
        ips = row["ips"] if row["ips"] else "N/A"
        count = row["violation_count"]
        total_count += count

        values = [i, row["attack_type"], violation_names, title, "High",
                  description, ips, VIOLATION_TYPE_CONST, count]
        for c, val in enumerate(values, start=1):
            cell = ws.cell(row=r, column=c)
            cell.value = val
            apply_style(cell, col_styles[c - 1])

    # TOTAL row at the end. The "TOTAL" label goes under column G
    # (Associated IP), not column B (Attack Type) - user preference,
    # confirmed 31-Aug-2026.
    total_row = len(rows) + 2
    ws.cell(row=total_row, column=7).value = "TOTAL"
    apply_style(ws.cell(row=total_row, column=7), col_styles[6], bold=True)
    ws.cell(row=total_row, column=9).value = total_count
    apply_style(ws.cell(row=total_row, column=9), col_styles[8], bold=True)
    for c in [1, 2, 3, 4, 5, 6, 8]:
        apply_style(ws.cell(row=total_row, column=c), col_styles[c - 1], bold=True)


def _set_chart_cat(series, formula):
    if series.cat is not None:
        if series.cat.numRef is not None:
            series.cat.numRef.f = formula
        elif series.cat.strRef is not None:
            series.cat.strRef.f = formula


def _set_chart_val(series, formula):
    if series.val is not None and series.val.numRef is not None:
        series.val.numRef.f = formula


def _chart_title_text(chart):
    try:
        return chart.title.tx.rich.p[0].r[0].t
    except Exception:
        return ""


def write_summary_dashboard(ws, weeks_data):
    """weeks_data: list of (label, rows) ordered newest -> oldest, up to 4."""
    old_max_row = ws.max_row
    old_max_col = ws.max_column

    # Capture block-0's original header/data row styles BEFORE touching
    # anything. These get re-applied to every block we write - including
    # the 4th block, which never existed in the original 3-block template
    # and would otherwise render with zero formatting.
    header_styles = [capture_style(ws.cell(row=1, column=c)) for c in range(1, BLOCK_WIDTH + 1)]
    data_row = 2 if old_max_row >= 2 else 1
    data_styles = [capture_style(ws.cell(row=data_row, column=c)) for c in range(1, 6)]

    # Unmerge everything first - old merges belonged to the old block
    # layout and would make some cells read-only ("MergedCell") otherwise.
    for merged_range in list(ws.merged_cells.ranges):
        ws.unmerge_cells(str(merged_range))

    max_col_needed = 1 + len(weeks_data) * BLOCK_WIDTH + 10  # blocks + severity table + margin
    max_row_needed = 100
    reset_style(ws, 1, max(old_max_row, max_row_needed), 1, max(old_max_col, max_col_needed))
    clear_range(ws, 1, max(old_max_row, max_row_needed), 1, max(old_max_col, max_col_needed))

    headers = ["Attack Type", "Violation Count", "Blocked", "Allowed / Logged", "Remediated"]
    block_meta = []

    for i, (label, rows) in enumerate(weeks_data):
        start_col = 1 + i * BLOCK_WIDTH
        for j, h in enumerate(headers):
            cell = ws.cell(row=1, column=start_col + j)
            cell.value = h
            apply_style(cell, header_styles[j])
        label_cell = ws.cell(row=1, column=start_col + 5)
        label_cell.value = label
        apply_style(label_cell, header_styles[5])

        sorted_rows = sorted(rows, key=lambda x: x["violation_count"], reverse=True)
        for k, row in enumerate(sorted_rows):
            r = 2 + k
            values = [row["attack_type"], row["violation_count"], row["blocked"],
                      row["allowed"], row["remediated"]]
            for j, val in enumerate(values):
                cell = ws.cell(row=r, column=start_col + j)
                cell.value = val
                apply_style(cell, data_styles[j])

        total_row = 2 + len(sorted_rows)
        total_count = sum(row["violation_count"] for row in sorted_rows)
        total_values = ["TOTAL ", total_count, f"{total_count}(100%)", 0,
                         f"{total_count}(Remediated at WAF Layer)"]
        for j, val in enumerate(total_values):
            cell = ws.cell(row=total_row, column=start_col + j)
            cell.value = val
            apply_style(cell, data_styles[j], bold=True)

        block_meta.append({
            "start_col": start_col,
            "n_rows": len(sorted_rows),
            "total_row": total_row,
            "total_count": total_count,
        })

    # Standalone comparison labels - own cell, not tied to any attack-type
    # row, merged across all 3 spacer columns so the text has room to
    # display instead of being squeezed into a single narrow cell.
    for i in range(len(block_meta) - 1):
        cur, older = block_meta[i], block_meta[i + 1]
        diff = cur["total_count"] - older["total_count"]
        word = "more" if diff >= 0 else "less"
        text = f"{abs(diff):,} {word} compared to last week"
        # Fixed row right under the block's week-label header, rather than
        # total_row + 2: total_row varies with how many attack-type rows
        # that particular week had, which put this caption at a different,
        # jagged row per block/site instead of lining up under the label.
        comp_row = 2
        comp_col = cur["start_col"] + 5
        ws.merge_cells(start_row=comp_row, start_column=comp_col,
                        end_row=comp_row, end_column=comp_col + 2)
        cell = ws.cell(row=comp_row, column=comp_col)
        cell.value = text
        apply_style(cell, header_styles[5])
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

    # Severity Count table (Critical always 0 since all rows are High)
    n_blocks = len(block_meta)
    sev_col = 1 + n_blocks * BLOCK_WIDTH + 1
    high_count = block_meta[0]["n_rows"] if block_meta else 0
    sev_header = [("Severity", "Count")]
    for j, (a, b) in enumerate(sev_header):
        ws.cell(row=1, column=sev_col).value = a
        ws.cell(row=1, column=sev_col + 1).value = b
        apply_style(ws.cell(row=1, column=sev_col), header_styles[0])
        apply_style(ws.cell(row=1, column=sev_col + 1), header_styles[1])
    for r, (label, val) in enumerate([("Critical", 0), ("High", high_count)], start=2):
        ws.cell(row=r, column=sev_col).value = label
        ws.cell(row=r, column=sev_col + 1).value = val
        apply_style(ws.cell(row=r, column=sev_col), data_styles[0])
        apply_style(ws.cell(row=r, column=sev_col + 1), data_styles[1])

    # --- Update the 3 existing charts in place (preserves styling) ---
    if block_meta:
        col0 = get_column_letter(block_meta[0]["start_col"])
        col1 = get_column_letter(block_meta[0]["start_col"] + 1)
        n0 = block_meta[0]["n_rows"]
        top_n = min(TOP_N_INCIDENT_CHART, n0)
        sev_col_letter = get_column_letter(sev_col)
        sev_val_letter = get_column_letter(sev_col + 1)
        sheet_name = ws.title

        for chart in ws._charts:
            title = _chart_title_text(chart)
            if "Incident Counts" in title:
                for s in chart.series:
                    _set_chart_cat(s, f"'{sheet_name}'!${col0}$2:${col0}${1+top_n}")
                    _set_chart_val(s, f"'{sheet_name}'!${col1}$2:${col1}${1+top_n}")
            elif "High Severity Attack Distribution" in title:
                for s in chart.series:
                    _set_chart_cat(s, f"'{sheet_name}'!${col0}$2:${col0}${1+n0}")
                    _set_chart_val(s, f"'{sheet_name}'!${col1}$2:${col1}${1+n0}")
            elif "High vs Critical" in title:
                for s in chart.series:
                    _set_chart_cat(s, f"'{sheet_name}'!${sev_col_letter}$2:${sev_col_letter}$3")
                    _set_chart_val(s, f"'{sheet_name}'!${sev_val_letter}$2:${sev_val_letter}$3")


# --------------------------------------------------------------------------
# Main
# --------------------------------------------------------------------------

def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--source", required=True)
    ap.add_argument("--template", required=True)
    ap.add_argument("--out", required=True)
    ap.add_argument("--week-end", default=None, help="YYYY-MM-DD, optional")
    args = ap.parse_args()

    week_end = date.fromisoformat(args.week_end) if args.week_end else None

    src_wb, weeks, target_key = load_source_workbook(args.source, week_end, n_weeks=MAX_HISTORY_WEEKS)
    week_keys_asc = last_n_weeks(weeks, target_key, n=MAX_HISTORY_WEEKS)   # oldest -> newest
    week_keys_desc = list(reversed(week_keys_asc))                        # newest -> oldest

    out_wb = openpyxl.load_workbook(args.template)

    full_label = week_label(*target_key)
    write_sheet1(out_wb["Weekly Report Data"], full_label)
    print(f"Target week: {full_label}")

    for site, (matrix_sheet, dash_sheet) in SHEET_MAP.items():
        cur_ref = weeks[target_key].get(site)
        cur_rows = read_full_attack_breakdown(src_wb[cur_ref.sheet_name]) if cur_ref else []
        write_detailed_vuln_matrix(out_wb[matrix_sheet], cur_rows)
        print(f"  {site}: {len(cur_rows)} attack-type rows in Detailed Vuln Matrix")

        weeks_data = []
        for k in week_keys_desc:
            ref = weeks[k].get(site)
            rows = read_full_attack_breakdown(src_wb[ref.sheet_name]) if ref else []
            weeks_data.append((short_week_label(*k), rows))
        write_summary_dashboard(out_wb[dash_sheet], weeks_data)
        print(f"  {site}: Summary Dashboard built with {len(weeks_data)} weekly block(s)")

    out_wb.save(args.out)
    print(f"\nWrote {args.out}")


if __name__ == "__main__":
    main()
