---
name: je-validator
description: Use this skill to validate a QuickBooks-style journal-entry Excel export before importing it into QuickBooks Online. It maps each transaction's account name against a chart of accounts (a bundled manufacturing-company COA by default, or your own COA file), checks double-entry debits/credits balance to the cent per entry, flags blank/zero/trailing junk rows, and writes an import-ready CSV for QBO's native JE importer. Trigger on requests like "validate this journal entry file," "check this JE export against our chart of accounts," "is this ready to import into QuickBooks," or when the user hands you a journal-entry .xlsx file.
---

# JE Validator

Validates a double-entry journal-entry Excel export against a chart of accounts, before it goes anywhere near QuickBooks Online's native CSV/Excel importer (Settings → Import Data → Journal Entries).

## When to use this

- The user hands you (or names) a journal-entry `.xlsx` file and asks you to check, validate, or confirm it's import-ready.
- The user wants to try the tool out with no files of their own yet — run it with no arguments to validate the bundled demo.

## Quick start

```bash
# Try it with the bundled sample (manufacturing COA + sample JE file, no setup needed)
python scripts/validate_je.py

# Validate a real file against the bundled manufacturing COA
python scripts/validate_je.py path/to/journal_entries.xlsx

# Validate against the user's own chart of accounts instead
python scripts/validate_je.py path/to/journal_entries.xlsx --coa path/to/their_coa.xlsx

# Choose output paths
python scripts/validate_je.py path/to/journal_entries.xlsx --report report.csv --out import_ready.csv
```

All script paths are relative to this skill's directory.

## What it checks

1. **Account mapping** — every account name is matched against the chart of accounts in tiers: exact `FullyQualifiedName` match → ambiguous leaf-name match (flagged, lists every full-path candidate) → fuzzy "possible match" suggestion (never auto-applied) → phantom account (no match at all, flagged).
2. **Double-entry balance** — total debits must equal total credits per journal entry, to the cent, after rounding.
3. **Per-line sanity** — exactly one of Debit/Credit per line and it's positive; blank rows and zero-amount rows are flagged as phantom rows, never silently dropped.
4. **Sheet-level noise** — a trailing "Total"/"Subtotal" row or trailing blank rows at the end of the sheet are detected and excluded from entries (reported as a `NOTE`, not a failure, since they're sheet formatting, not transaction data).
5. **Grouping-key continuity** — if a journal/entry-ID column exists, a blank ID on what looks like a continuation line is forward-filled to the previous entry and flagged via a `NOTE`, rather than silently becoming its own (wrong) one-line entry. If no ID column exists at all, rows are grouped into entries by blank-row-separated blocks instead.

Columns are auto-detected from the file's own headers (Account/GL Account/Account Name, Debit/Credit or a single signed Amount column, Journal No./Entry ID/Transaction ID, Date, Description/Memo) — there's no fixed template to match. If required columns can't be confidently identified, the script stops and reports exactly what it found and what's missing/ambiguous, rather than guessing.

## Output

- **Validation report** (`validation_report.csv` by default): one row per journal entry — `SourceFile, EntryId, Status, ErrorCount, Errors`.
- **Import-ready CSV** (`je_import_ready.csv` by default): only entries that passed validation, rewritten in QBO's exact native-importer header format (`Journal No., Journal Date, Account Name, Description, Debits, Credits`), with account names normalized to their full `FullyQualifiedName`.

## Files

- `scripts/validate_je.py` — main entry point (run this).
- `scripts/column_detect.py` — header-mapping heuristics.
- `scripts/coa_match.py` — chart-of-accounts loading and account-name matching tiers.
- `scripts/build_fixtures.py` — regenerates the bundled COA/demo/test fixtures; not needed at run time.
- `references/manufacturing_coa.xlsx` — bundled default chart of accounts (manufacturing company).
- `references/sample_journal_entries.xlsx` — demo JE file (2 clean entries, 2 with intentional defects) used when no file is given.
- `test_files/` — regression fixtures, each isolating exactly one defect: `je_good.xlsx`, `je_bad_unbalanced.xlsx`, `je_bad_rounding.xlsx`, `je_bad_phantom_blank_row.xlsx`, `je_bad_phantom_zero_row.xlsx`, `je_bad_ambiguous_leaf.xlsx`, `je_bad_trailing_total_row.xlsx`, `je_bad_broken_grouping.xlsx`. Re-run all of these after any change to validation logic.
