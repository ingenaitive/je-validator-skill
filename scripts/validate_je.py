#!/usr/bin/env python3
"""Validate a QuickBooks-style journal-entry Excel export against a chart of
accounts (bundled manufacturing COA by default, or a user-supplied one),
assuming double-entry bookkeeping. Produces a validation report and, for
entries that pass, an import-ready CSV for QBO's native JE importer.

Usage:
    python validate_je.py <je_file.xlsx> [--coa coa_file.xlsx]
                           [--report validation_report.csv]
                           [--out je_import_ready.csv]

With no arguments, validates the bundled sample files (references/) so a
new user can see real output with zero setup.
"""
from __future__ import annotations

import argparse
import csv
import os
import re
import sys

import pandas as pd

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from column_detect import detect_je_columns, detect_coa_columns  # noqa: E402
from coa_match import ChartOfAccounts, load_coa_rows_from_dicts  # noqa: E402

SKILL_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
DEFAULT_COA = os.path.join(SKILL_DIR, "references", "manufacturing_coa.xlsx")
DEFAULT_JE = os.path.join(SKILL_DIR, "references", "sample_journal_entries.xlsx")

JUNK_LABELS = re.compile(r"^\s*(total|subtotal|grand total|totals?)\b", re.IGNORECASE)


def parse_amount(value):
    if value is None:
        return None
    s = str(value).strip()
    if s == "" or s.lower() == "nan":
        return None
    neg = s.startswith("(") and s.endswith(")")
    s = s.strip("()").replace("$", "").replace(",", "").strip()
    if s == "":
        return None
    try:
        amt = float(s)
    except ValueError:
        return None
    if neg:
        amt = -amt
    return round(amt, 2)


def row_is_blank(row: dict, cols: list[str]) -> bool:
    for c in cols:
        v = row.get(c)
        if v is not None and str(v).strip() not in ("", "nan"):
            return False
    return True


def load_sheet(path: str) -> tuple[list[str], list[dict]]:
    df = pd.read_excel(path, dtype=str)
    headers = [str(h) for h in df.columns]
    records = df.to_dict(orient="records")
    return headers, records


def detect_je_columns_with_fallback(headers, records):
    detection = detect_je_columns(headers)
    if not detection.missing and not detection.ambiguous:
        return detection
    return detection  # caller decides how to react


def resolve_amounts(mapping, record) -> tuple[float | None, float | None]:
    if "debit" in mapping and "credit" in mapping:
        return parse_amount(record.get(mapping["debit"])), parse_amount(record.get(mapping["credit"]))
    if "amount" in mapping and mapping.get("amount"):
        amt = parse_amount(record.get(mapping["amount"]))
        if amt is None:
            return None, None
        return (amt, None) if amt > 0 else (None, -amt) if amt < 0 else (None, None)
    return None, None


def strip_trailing_junk(records: list[dict], account_col: str):
    """Returns (data_records, junk_notes). Drops a trailing run of blank rows
    and/or a trailing Total/Subtotal row; does NOT touch blank/junk rows that
    occur before real data (those are left for the caller's phantom-row logic
    inside entry grouping, since a blank row there may be a real sheet
    noise or a genuine phantom line within an entry)."""
    all_cols = list(records[0].keys()) if records else []
    notes = []
    end = len(records)
    blank_dropped = 0

    while end > 0:
        row = records[end - 1]
        acct_val = str(row.get(account_col, "")).strip()
        if row_is_blank(row, all_cols):
            blank_dropped += 1
            end -= 1
            continue
        if JUNK_LABELS.match(acct_val):
            notes.append(f"Row {end}: trailing '{acct_val}' row ignored (sheet-level total, not a journal line).")
            end -= 1
            continue
        break

    if blank_dropped:
        notes.insert(0, f"{blank_dropped} trailing blank row(s) at end of sheet ignored.")

    return records[:end], notes


def group_entries(records: list[dict], mapping: dict):
    """Returns (entries: dict[str, list[dict]], notes: list[str]).
    Groups by entry_id if present (forward-filling blanks as continuation
    lines, flagged via notes), else by contiguous blocks split on blank rows.
    """
    notes = []
    account_col = mapping["account"]
    all_cols = list(records[0].keys()) if records else []

    if "entry_id" in mapping:
        id_col = mapping["entry_id"]
        entries: dict[str, list[dict]] = {}
        order: list[str] = []
        current_id = None
        for idx, row in enumerate(records):
            raw_id = str(row.get(id_col, "")).strip()
            if raw_id and raw_id.lower() != "nan":
                current_id = raw_id
            else:
                if row_is_blank(row, all_cols):
                    continue  # sheet-level separator blank row, not data
                if current_id is None:
                    current_id = f"(unlabeled-row-{idx})"
                    notes.append(
                        f"Row {idx}: no entry id and none seen yet -- treated as its own entry "
                        f"'{current_id}'."
                    )
                else:
                    notes.append(
                        f"Row {idx}: blank entry id on a continuation line -- assumed to belong "
                        f"to entry '{current_id}' (grouping key should repeat on every line)."
                    )
            entries.setdefault(current_id, [])
            if current_id not in order:
                order.append(current_id)
            entries[current_id].append(row)
        return {k: entries[k] for k in order}, notes

    # Fallback: group by contiguous blocks separated by blank rows.
    entries = {}
    block = []
    block_num = 0
    for row in records:
        if row_is_blank(row, all_cols):
            if block:
                block_num += 1
                entries[f"block-{block_num}"] = block
                block = []
            continue
        block.append(row)
    if block:
        block_num += 1
        entries[f"block-{block_num}"] = block
    notes.append(
        "No journal/entry-id column found -- grouped rows into entries by "
        "blank-row-separated blocks."
    )
    return entries, notes


def validate_entry(entry_id, rows, mapping, coa: ChartOfAccounts):
    errors = []
    total_debit = 0.0
    total_credit = 0.0
    account_col = mapping["account"]
    all_cols = list(rows[0].keys()) if rows else []
    resolved_lines = []

    for i, row in enumerate(rows):
        acct_raw = str(row.get(account_col, "")).strip()
        debit, credit = resolve_amounts(mapping, row)
        no_account = not acct_raw or acct_raw.lower() == "nan"
        no_amount = (debit is None or debit == 0.0) and (credit is None or credit == 0.0)

        if no_account and no_amount:
            errors.append(f"Line {i}: phantom row (no account, no amount).")
            continue
        if not no_account and no_amount:
            errors.append(f"Line {i}: phantom row (account '{acct_raw}' has zero/empty debit and credit).")
            continue
        if no_account:
            errors.append(f"Line {i}: amount present but no account name provided.")
            continue

        if debit and credit:
            errors.append(f"Line {i}: both Debit and Credit populated for '{acct_raw}' -- exactly one is allowed.")
            continue
        if debit is not None and debit < 0:
            errors.append(f"Line {i}: negative debit ({debit}) for '{acct_raw}'.")
            continue
        if credit is not None and credit < 0:
            errors.append(f"Line {i}: negative credit ({credit}) for '{acct_raw}'.")
            continue

        match = coa.match(acct_raw)
        resolved_fqn = None
        if match["status"] == "exact":
            resolved_fqn = match["fqn"]
        elif match["status"] == "ambiguous_leaf":
            errors.append(
                f"Line {i}: account '{acct_raw}' is not a fully-qualified name. "
                f"Matches {len(match['candidates'])} account(s) by leaf name only "
                f"({', '.join(match['candidates'])}) -- use the full Parent:Child path."
            )
        elif match["status"] == "fuzzy_suggestion":
            suggestions = ", ".join(f"{c} ({s:.0%})" for c, s in match["candidates"])
            errors.append(
                f"Line {i}: account '{acct_raw}' not found in chart of accounts. "
                f"Possible match (not auto-applied): {suggestions}."
            )
        else:
            errors.append(f"Line {i}: account '{acct_raw}' not found in chart of accounts (phantom account).")

        total_debit += debit or 0.0
        total_credit += credit or 0.0
        resolved_lines.append({
            "account": resolved_fqn or acct_raw,
            "debit": debit or 0.0,
            "credit": credit or 0.0,
            "description": row.get(mapping.get("description"), "") if mapping.get("description") else "",
            "date": row.get(mapping.get("date"), "") if mapping.get("date") else "",
        })

    if round(total_debit, 2) != round(total_credit, 2):
        diff = round(total_debit - total_credit, 2)
        errors.append(
            f"Entry does not balance: total debits {total_debit:.2f} != "
            f"total credits {total_credit:.2f} (diff {diff:+.2f})."
        )

    return errors, resolved_lines


def load_coa(path: str) -> ChartOfAccounts:
    headers, records = load_sheet(path)
    detection = detect_coa_columns(headers)
    if detection.missing or detection.ambiguous:
        raise SystemExit(
            f"Could not confidently map chart-of-accounts columns in {path}.\n"
            f"Missing: {detection.missing}\nAmbiguous: {detection.ambiguous}\n"
            f"Headers found: {headers}"
        )
    rows = load_coa_rows_from_dicts(records, detection.mapping)
    return ChartOfAccounts(rows)


def write_report(report_path, results):
    with open(report_path, "w", newline="") as f:
        writer = csv.writer(f)
        writer.writerow(["SourceFile", "EntryId", "Status", "ErrorCount", "Errors"])
        for r in results:
            writer.writerow([r["source"], r["entry_id"], r["status"], r["error_count"], r["errors"]])


def write_import_ready(out_path, results):
    with open(out_path, "w", newline="") as f:
        writer = csv.writer(f)
        writer.writerow(["Journal No.", "Journal Date", "Account Name", "Description", "Debits", "Credits"])
        for r in results:
            if r["status"] != "OK":
                continue
            for line in r["lines"]:
                writer.writerow([
                    r["entry_id"], line.get("date", ""), line["account"], line.get("description", ""),
                    f"{line['debit']:.2f}" if line["debit"] else "",
                    f"{line['credit']:.2f}" if line["credit"] else "",
                ])


def validate_file(je_path: str, coa: ChartOfAccounts):
    headers, records = load_sheet(je_path)
    if not records:
        return [], [f"{je_path}: no data rows found."]

    detection = detect_je_columns(headers)
    if detection.missing or detection.ambiguous:
        msg = (
            f"Could not confidently map columns in {je_path}.\n"
            f"Missing: {detection.missing}\nAmbiguous: {detection.ambiguous}\n"
            f"Headers found: {headers}\n"
            f"Fix the column headers or rename them to something like "
            f"Account / Debit / Credit / Journal No. / Date."
        )
        return [], [msg]

    mapping = detection.mapping
    sheet_notes = list(detection.notes)

    data_records, junk_notes = strip_trailing_junk(records, mapping["account"])
    sheet_notes += junk_notes

    entries, grouping_notes = group_entries(data_records, mapping)
    sheet_notes += grouping_notes

    results = []
    for entry_id, rows in entries.items():
        errors, lines = validate_entry(entry_id, rows, mapping, coa)
        results.append({
            "source": os.path.basename(je_path),
            "entry_id": entry_id,
            "status": "OK" if not errors else "INVALID",
            "error_count": len(errors),
            "errors": " | ".join(errors),
            "lines": lines,
        })
    return results, sheet_notes


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("je_file", nargs="?", default=None, help="Journal entry .xlsx file")
    parser.add_argument("--coa", default=None, help="Chart of accounts .xlsx (defaults to bundled manufacturing COA)")
    parser.add_argument("--report", default="validation_report.csv")
    parser.add_argument("--out", default="je_import_ready.csv")
    args = parser.parse_args()

    je_path = args.je_file or DEFAULT_JE
    coa_path = args.coa or DEFAULT_COA
    using_demo = args.je_file is None

    if using_demo:
        print(f"No file given -- running the bundled demo:\n  JE file: {je_path}\n  COA:     {coa_path}\n")

    coa = load_coa(coa_path)
    results, sheet_notes = validate_file(je_path, coa)

    if not results:
        for line in sheet_notes:
            print(line)
        sys.exit(1)

    for note in sheet_notes:
        print(f"NOTE: {note}")

    passed = failed = 0
    issues = 0
    for r in results:
        if r["status"] == "OK":
            passed += 1
            print(f"[OK]      {r['entry_id']}")
        else:
            failed += 1
            issues += r["error_count"]
            print(f"[INVALID] {r['entry_id']}")
            for e in r["errors"].split(" | "):
                print(f"    - {e}")

    write_report(args.report, results)
    write_import_ready(args.out, results)

    print(f"\n{len(results)} entries checked, {passed} passed, {failed} failed, {issues} total issues found.")
    print(f"Report written to {args.report}")
    print(f"Import-ready lines (passing entries only) written to {args.out}")

    sys.exit(1 if failed else 0)


if __name__ == "__main__":
    main()
