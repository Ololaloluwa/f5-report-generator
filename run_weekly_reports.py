#!/usr/bin/env python3
"""
Weekly MTN F5 Reports - run everything
=========================================

One command that runs BOTH generators:
  1. generate_weekly_report.py  -> PowerPoint deck
  2. generate_weekly_excel.py   -> Detailed WAF Excel workbook

Requires generate_weekly_report.py and generate_weekly_excel.py to be in
the same folder as this file (this script imports them directly rather
than shelling out, so you get one clean error message instead of two
separate tracebacks if something goes wrong).

Usage:
    python run_weekly_reports.py \
        --source MTN_F5_ATTACKS_SOURCE.xlsx \
        --pptx-template MTN_Security_Metrics_Report_SOURCE.pptx \
        --xlsx-template MTN_F5_WAF_Weekly_Report_Temp.xlsx \
        --out-dir ./weekly_output \
        [--week-end 2026-07-19] \
        [--certs Certificate.xlsx]   # default; certificate slides are skipped if it's missing

Output files are named automatically using the target week's dates, e.g.:
    weekly_output/MTN_Security_Metrics_Report_13-19Jul2026.pptx
    weekly_output/MTN_F5_WAF_Weekly_Report_13-19Jul2026.xlsx

The actual work lives in generate_reports(), which the command line (and
the .bat file) and the web app both call - so there is only ONE code path
producing the reports, whichever way they're run.
"""

import argparse
import os
from dataclasses import dataclass, field
from datetime import date

import openpyxl

import generate_weekly_report as pptx_gen
import generate_weekly_excel as xlsx_gen


@dataclass
class ReportResult:
    week: tuple              # (start_date, end_date) of the week reported on
    week_label: str          # e.g. "September 7th - September 13th, 2026"
    pptx_path: str
    xlsx_path: str
    log: list = field(default_factory=list)       # every line printed during the run
    warnings: list = field(default_factory=list)  # just the [WARN] lines


def generate_reports(source, pptx_template, xlsx_template, out_dir,
                     week_end=None, certs="Certificate.xlsx", log=print):
    """
    Builds the PowerPoint deck and the detailed WAF Excel workbook for one
    week and returns a ReportResult. `week_end` (a date) picks the week the
    same way --week-end does; None = most recent complete week. `log` gets
    every progress line (print by default; the web app passes its own).
    """
    lines = []

    def out(msg=""):
        lines.append(msg)
        log(msg)

    os.makedirs(out_dir, exist_ok=True)

    # Figure out the target week once, up front, so both outputs are
    # guaranteed to use the identical week even if this runs right at a
    # week boundary.
    # Only the target week + the 3 before it are loaded, values only - see
    # load_source_workbook() for why (the full workbook needs ~2.3 GB).
    wb, weeks, target_key = pptx_gen.load_source_workbook(source, week_end, n_weeks=4)
    start, end = target_key
    tag = f"{start.day}-{end.day}{end.strftime('%b%Y')}"

    pptx_out = os.path.join(out_dir, f"MTN_Security_Metrics_Report_{tag}.pptx")
    waf_xlsx_out = os.path.join(out_dir, f"MTN_F5_WAF_Weekly_Report_{tag}.xlsx")

    out("=" * 60)
    out("STEP 1/2: Generating PowerPoint deck")
    out("=" * 60)
    week_keys_asc = pptx_gen.last_n_weeks(weeks, target_key, n=4)
    history = [pptx_gen.compute_week_metrics(wb, k, weeks[k]) for k in week_keys_asc]
    current = history[-1]
    out(f"Target week: {current.label}")
    out(f"Requests Inspected (combined): {current.requests_inspected:,}")
    out(f"WAF Blocks (combined): {current.waf_blocks:,}")

    per_site = {}
    for site in pptx_gen.SITES:
        site_hist = [pptx_gen.compute_week_metrics(wb, k, weeks[k], sites=[site]) for k in week_keys_asc]
        per_site[site] = (site_hist[-1], site_hist)
        out(f"  {site}: Requests Inspected={site_hist[-1].requests_inspected:,}  WAF Blocks={site_hist[-1].waf_blocks:,}")

    out("Node changes / certificates:")
    node_changes, cert_data = pptx_gen.load_node_changes_and_certs(wb, weeks, target_key, certs, log=out)
    pptx_gen.fill_pptx(pptx_template, pptx_out, current, history, per_site,
                       node_changes=node_changes, certs=cert_data, cert_year=end.year)
    out(f"Wrote {pptx_out}")

    out()
    out("=" * 60)
    out("STEP 2/2: Generating detailed WAF Excel workbook")
    out("=" * 60)
    out_wb = openpyxl.load_workbook(xlsx_template)
    full_label = pptx_gen.week_label(*target_key)
    xlsx_gen.write_sheet1(out_wb["Weekly Report Data"], full_label)
    week_keys_desc = list(reversed(week_keys_asc))
    for site, (matrix_sheet, dash_sheet) in xlsx_gen.SHEET_MAP.items():
        cur_ref = weeks[target_key].get(site)
        cur_rows = xlsx_gen.read_full_attack_breakdown(wb[cur_ref.sheet_name]) if cur_ref else []
        xlsx_gen.write_detailed_vuln_matrix(out_wb[matrix_sheet], cur_rows)
        out(f"  {site}: {len(cur_rows)} attack-type rows in Detailed Vuln Matrix")

        weeks_data = []
        for k in week_keys_desc:
            ref = weeks[k].get(site)
            rows = xlsx_gen.read_full_attack_breakdown(wb[ref.sheet_name]) if ref else []
            weeks_data.append((xlsx_gen.short_week_label(*k), rows))
        xlsx_gen.write_summary_dashboard(out_wb[dash_sheet], weeks_data)
        out(f"  {site}: Summary Dashboard built with {len(weeks_data)} weekly block(s)")
    out_wb.save(waf_xlsx_out)
    out(f"Wrote {waf_xlsx_out}")

    out()
    out("=" * 60)
    out(f"DONE - 2 files written to {out_dir}")
    out("=" * 60)
    out(f"  {os.path.basename(pptx_out)}")
    out(f"  {os.path.basename(waf_xlsx_out)}")

    return ReportResult(
        week=target_key, week_label=full_label, pptx_path=pptx_out, xlsx_path=waf_xlsx_out,
        log=lines, warnings=[l.strip() for l in lines if "[WARN]" in l],
    )


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--source", required=True, help="MTN_F5_ATTACKS_SOURCE.xlsx")
    ap.add_argument("--pptx-template", required=True, help="PowerPoint template (MTN_Security_Metrics_Report_SOURCE.pptx)")
    ap.add_argument("--xlsx-template", required=True, help="Excel template (MTN_F5_WAF_Weekly_Report_Temp.xlsx)")
    ap.add_argument("--out-dir", default=".", help="Folder to write outputs into (created if missing)")
    ap.add_argument("--week-end", default=None, help="YYYY-MM-DD, optional - defaults to most recent complete week")
    ap.add_argument("--certs", default="Certificate.xlsx",
                    help="Certificate workbook for the SSL/device certificate slides (skipped with a warning if missing)")
    args = ap.parse_args()

    week_end = date.fromisoformat(args.week_end) if args.week_end else None
    generate_reports(args.source, args.pptx_template, args.xlsx_template, args.out_dir,
                     week_end=week_end, certs=args.certs)


if __name__ == "__main__":
    main()
