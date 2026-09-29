#!/usr/bin/env python3
"""
MTN F5 WAF MONTHLY Report - Excel
===================================

Fills the same Excel template as the weekly report
(MTN_F5_WAF_Weekly_Report_Temp.xlsx), with months instead of weeks:
  - Sheet 1 is renamed "Monthly Report Data" and its Report Range becomes
    e.g. "September 1st - September 30th, 2026".
  - "Detailed Vuln Matrix <SITE>": every attack type for the whole month
    (biggest first), with the Associated IPs of every week/day combined.
  - "Summary Dashboard <SITE>": up to 4 month blocks side by side (this
    month first), each with "X more/less compared to last month", and the
    3 charts pointed at this month's block.

The month figures come from generate_monthly_report.compute_site_month(),
so the Excel totals are always the same numbers as the PowerPoint's WAF
Blocks. The writers themselves are the weekly ones (imported, unchanged).
"""

import openpyxl

import generate_weekly_excel as wx
from generate_monthly_report import SITES, month_days, month_label

MAX_IP_TEXT = 32000        # an Excel cell holds at most 32,767 characters


def _ord(d):
    return "th" if 11 <= d % 100 <= 13 else {1: "st", 2: "nd", 3: "rd"}.get(d % 10, "th")


def month_range_label(ym):
    first, last = month_days(ym)[0], month_days(ym)[-1]
    return f"{first:%B} {first.day}{_ord(first.day)} - {last:%B} {last.day}{_ord(last.day)}, {last.year}"


def _ip_text(ips):
    text = ", ".join(ips)
    if len(text) <= MAX_IP_TEXT:
        return text or None
    cut = text[:MAX_IP_TEXT].rsplit(", ", 1)[0]
    shown = cut.count(", ") + 1
    return f"{cut} ... (+{len(ips) - shown:,} more)"


def excel_rows(mm):
    """MonthMetrics.attack_rows -> the row dicts the weekly writers take, biggest first."""
    rows = []
    for r in mm.attack_rows.values():
        rows.append({
            "attack_type": r["attack_type"],
            "violation_count": r["violation_count"],
            "blocked": r["blocked"],
            "allowed": r["allowed"],
            "remediated": r["remediated"],
            "ips": _ip_text(r["ips"]),
        })
    return sorted(rows, key=lambda x: x["violation_count"], reverse=True)


def write_monthly_excel(template_path, out_path, ym, site_trends):
    """
    site_trends: {site: [MonthMetrics]} oldest -> newest (last = the month
    being reported), the same list the PowerPoint uses.
    """
    wb = openpyxl.load_workbook(template_path)
    first = wb.worksheets[0]
    first.title = "Monthly Report Data"
    wx.write_sheet1(first, month_range_label(ym))

    for site in SITES:
        matrix_sheet, dash_sheet = wx.SHEET_MAP[site]
        trend = site_trends[site]
        wx.write_detailed_vuln_matrix(wb[matrix_sheet], excel_rows(trend[-1]))
        blocks = [(month_label(m.ym), excel_rows(m)) for m in reversed(trend)]   # newest first
        ws = wb[dash_sheet]
        wx.write_summary_dashboard(ws, blocks)
        for cell in ws[2]:
            if isinstance(cell.value, str) and cell.value.endswith("compared to last week"):
                cell.value = cell.value.replace("last week", "last month")
    wb.save(out_path)
