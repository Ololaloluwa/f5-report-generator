#!/usr/bin/env python3
"""
Validation checks for the MTN F5 weekly report pipeline.

Run this AFTER generate_weekly_report.py / generate_weekly_excel.py /
run_weekly_reports.py, BEFORE treating the run as successful. It re-reads
the same source workbook, re-derives what the target week SHOULD be, and
sanity-checks the numbers - catching the class of bug we've hit repeatedly:
the scripts running with no error, but silently producing wrong output
because a table/column/sheet-name didn't match what the parser expected.

This does NOT fix anything. It only reports. A human (or the agent,
after explaining the finding and getting confirmation) decides the fix.

Usage:
    python validate_outputs.py --source MTN_F5_ATTACKS_SOURCE.xlsx [--week-end YYYY-MM-DD]

Exit code 0 = all checks passed. Exit code 1 = at least one check failed
(see printed [FAIL] lines for details).
"""

import argparse
import sys
from dataclasses import dataclass, field
from datetime import date

import openpyxl

from generate_weekly_report import (
    SITES,
    group_sheets_by_week,
    pick_target_week,
    last_n_weeks,
    load_source_workbook,
    read_attack_breakdown,
    read_kv_table,
    compute_week_metrics,
    check_sheet_names,
    suggest_sheet_names,
)

MAX_REASONABLE_WEEK_SPAN_DAYS = 9   # a real Mon-Sun week is 7; allow slack for off-by-one typos
MIN_REASONABLE_WEEK_SPAN_DAYS = 5
KPI_JUMP_WARN_RATIO = 0.5            # warn if this week's KPI is <50% or >200% of last week's


@dataclass
class ValidationResult:
    ok: bool                 # False if any [FAIL]
    week: tuple              # (start_date, end_date) that was checked
    lines: list = field(default_factory=list)   # every [PASS]/[WARN]/[FAIL] line, in order


def run_checks(source, week_end=None, log=print):
    """
    Runs every check and returns a ValidationResult. The command line (and
    the .bat file) print the lines; the web app shows them on the page.
    `week_end` (a date) picks the week the same way --week-end does.
    """
    lines = []

    def out(msg=""):
        lines.append(msg)
        log(msg)

    def check(label, condition, detail=""):
        status = "PASS" if condition else "FAIL"
        out(f"[{status}] {label}" + (f" - {detail}" if detail and not condition else ""))
        return condition

    # target week + the week before it only, values only (see
    # load_source_workbook in generate_weekly_report.py)
    wb, weeks, target_key = load_source_workbook(source, week_end, n_weeks=2)
    start, end = target_key
    span_days = (end - start).days + 1

    all_ok = True
    out(f"Target week resolved to: {start} - {end} ({span_days} days)")
    out()

    # --- Check 1: the dates WRITTEN in this week's sheet names span ~7
    # days. Sheets are placed in their Monday-Sunday week even when the
    # dates are slightly off (a [WARN] from the sheet check below says so);
    # a span way off 7 days means the name needs a human look.
    bad_spans = {}
    for site, ref in weeks.get(target_key, {}).items():
        written = (ref.raw_end - ref.raw_start).days + 1
        if not MIN_REASONABLE_WEEK_SPAN_DAYS <= written <= MAX_REASONABLE_WEEK_SPAN_DAYS:
            bad_spans[site] = written
    ok = check(
        "Target week span is ~7 days",
        not bad_spans,
        "dates written in the sheet name span " + ", ".join(f"{s}: {d} days" for s, d in bad_spans.items())
        + " - likely a sheet-naming typo (e.g. wrong month/day); check which week it was meant to be",
    )
    all_ok &= ok

    # --- Check 2: all 3 sites present for target week ---
    site_refs = weeks.get(target_key, {})
    missing_sites = [s for s in SITES if s not in site_refs]
    ok = check(
        "All 3 sites (IKY/VGC/OJT) found for target week",
        not missing_sites,
        f"missing: {missing_sites} - check sheet naming for that site this week "
        f"(missing 'for', wrong site code, trailing space, etc.)",
    )
    all_ok &= ok

    # --- Check 3: each present site has >0 attack-type rows and >0 throughput ---
    for site in SITES:
        ref = site_refs.get(site)
        if ref is None:
            continue
        ws = wb[ref.sheet_name]
        attacks = read_attack_breakdown(ws)
        throughput = read_kv_table(ws, "Weekly Request Throughput")
        requests = throughput.get("Requests Inspected", 0) or 0

        ok = check(
            f"{site}: Weekly WAF Attack Breakdown has attack-type rows",
            len(attacks) > 0,
            "0 rows found - table title or column layout likely doesn't match what the "
            "parser expects this week (e.g. shifted columns, renamed title)",
        )
        all_ok &= ok

        ok = check(
            f"{site}: Requests Inspected > 0",
            requests > 0,
            "0 or missing - 'Weekly Request Throughput' table probably wasn't found "
            "(missing 'Weekly' prefix, different title, etc.)",
        )
        all_ok &= ok

    # --- Check 4: week-over-week KPI sanity (warn only, not a hard fail) ---
    try:
        week_keys = last_n_weeks(weeks, target_key, n=2)
        if len(week_keys) == 2:
            prev_metrics = compute_week_metrics(wb, week_keys[0], weeks[week_keys[0]])
            cur_metrics = compute_week_metrics(wb, week_keys[1], weeks[week_keys[1]])
            for label, prev_val, cur_val in [
                ("Requests Inspected", prev_metrics.requests_inspected, cur_metrics.requests_inspected),
                ("WAF Blocks", prev_metrics.waf_blocks, cur_metrics.waf_blocks),
            ]:
                if prev_val > 0:
                    ratio = cur_val / prev_val
                    within_range = (1 - KPI_JUMP_WARN_RATIO) <= ratio <= (1 + KPI_JUMP_WARN_RATIO)
                    status = "PASS" if within_range else "WARN"
                    out(f"[{status}] {label} vs last week: {prev_val:,} -> {cur_val:,} "
                          f"({ratio:.1%} of last week)" + ("" if within_range else
                          " - large swing, worth a manual glance (could be legitimate)"))
    except Exception as e:
        out(f"[WARN] Could not run week-over-week comparison: {e}")

    # --- Check 5: sheet names (warn only) - unreadable names, a site
    # missing, dates that had to be corrected, duplicates, future weeks ---
    for warning in check_sheet_names(wb.source_sheetnames, focus_weeks=list(weeks)):
        out(f"[WARN] {warning}")

    out()
    if all_ok:
        out("All hard checks passed.")
    else:
        out("One or more checks FAILED - do not treat this run as reliable. "
            "Investigate the flagged table(s) in the raw source file before trusting the output.")
    # --- What next week's sheets should be called (always under Excel's
    # 31-character sheet-name limit) ---
    _, next_names = suggest_sheet_names(wb.latest_week[1])
    out()
    out("Next week's sheets should be named:")
    for name in next_names:
        out(f"  {name}")
    return ValidationResult(ok=all_ok, week=target_key, lines=lines)


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--source", required=True)
    ap.add_argument("--week-end", default=None)
    args = ap.parse_args()

    week_end = date.fromisoformat(args.week_end) if args.week_end else None
    result = run_checks(args.source, week_end)
    sys.exit(0 if result.ok else 1)


if __name__ == "__main__":
    main()
