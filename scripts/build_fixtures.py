#!/usr/bin/env python3
"""One-off generator for the bundled manufacturing COA, demo JE file, and the
test_files/ regression fixtures. Not needed at skill run time -- rerun this
only if you want to regenerate the .xlsx fixtures from scratch."""
import os

import pandas as pd

SKILL_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
REF_DIR = os.path.join(SKILL_DIR, "references")
TEST_DIR = os.path.join(SKILL_DIR, "test_files")

MANUFACTURING_COA = [
    # Name, FullyQualifiedName, AccountType
    ("Checking", "Checking", "Bank"),
    ("Accounts Receivable (A/R)", "Accounts Receivable (A/R)", "Accounts Receivable"),
    ("Raw Materials Inventory", "Inventory Asset:Raw Materials Inventory", "Other Current Asset"),
    ("Work in Process Inventory", "Inventory Asset:Work in Process Inventory", "Other Current Asset"),
    ("Finished Goods Inventory", "Inventory Asset:Finished Goods Inventory", "Other Current Asset"),
    ("Factory Equipment", "Fixed Assets:Factory Equipment", "Fixed Asset"),
    ("Accumulated Depreciation - Factory Equipment", "Fixed Assets:Accumulated Depreciation - Factory Equipment", "Fixed Asset"),
    ("Accounts Payable (A/P)", "Accounts Payable (A/P)", "Accounts Payable"),
    ("Accrued Payroll", "Accrued Payroll", "Other Current Liability"),
    ("Common Stock", "Common Stock", "Equity"),
    ("Retained Earnings", "Retained Earnings", "Equity"),
    ("Product Sales", "Product Sales", "Income"),
    ("Direct Materials", "Cost of Goods Sold:Direct Materials", "Cost of Goods Sold"),
    ("Direct Labor", "Cost of Goods Sold:Direct Labor", "Cost of Goods Sold"),
    ("Manufacturing Overhead", "Cost of Goods Sold:Manufacturing Overhead", "Cost of Goods Sold"),
    ("Factory Supplies", "Manufacturing Overhead:Factory Supplies", "Expense"),
    ("Factory Utilities", "Manufacturing Overhead:Factory Utilities", "Expense"),
    ("Equipment Maintenance", "Manufacturing Overhead:Equipment Maintenance", "Expense"),
    ("Depreciation Expense", "Depreciation Expense", "Expense"),
    ("Office Supplies", "Office Supplies", "Expense"),
]


def write_coa():
    df = pd.DataFrame(MANUFACTURING_COA, columns=["Name", "FullyQualifiedName", "AccountType"])
    os.makedirs(REF_DIR, exist_ok=True)
    df.to_excel(os.path.join(REF_DIR, "manufacturing_coa.xlsx"), index=False)


def df_from_rows(rows, columns=("Journal No.", "Journal Date", "Account Name", "Description", "Debits", "Credits")):
    return pd.DataFrame(rows, columns=list(columns))


def write_xlsx(rows, path, columns=("Journal No.", "Journal Date", "Account Name", "Description", "Debits", "Credits")):
    df_from_rows(rows, columns).to_excel(path, index=False)


def write_sample_demo():
    rows = [
        # JE1: clean -- raw materials issued to WIP
        ("JE1", "2026-11-01", "Inventory Asset:Work in Process Inventory", "Issue raw materials to WIP - Batch 101", 4200.00, None),
        ("JE1", "2026-11-01", "Inventory Asset:Raw Materials Inventory", "Issue raw materials to WIP - Batch 101", None, 4200.00),
        # JE2: clean -- finished goods completed
        ("JE2", "2026-11-03", "Inventory Asset:Finished Goods Inventory", "Batch 101 complete, move WIP to FG", 6100.00, None),
        ("JE2", "2026-11-03", "Inventory Asset:Work in Process Inventory", "Batch 101 complete, move WIP to FG", None, 6100.00),
        # JE3: unbalanced (defect, for the demo)
        ("JE3", "2026-11-05", "Cost of Goods Sold:Direct Labor", "Payroll allocation to production", 3000.00, None),
        ("JE3", "2026-11-05", "Accrued Payroll", "Payroll allocation to production", None, 2800.00),
        # JE4: ambiguous leaf name (defect, for the demo)
        ("JE4", "2026-11-06", "Factory Supplies", "Shop towels and lubricant", 180.00, None),
        ("JE4", "2026-11-06", "Checking", "Shop towels and lubricant", None, 180.00),
    ]
    write_xlsx(rows, os.path.join(REF_DIR, "sample_journal_entries.xlsx"))


def write_test_files():
    os.makedirs(TEST_DIR, exist_ok=True)

    write_xlsx([
        ("JE1001", "2026-11-01", "Inventory Asset:Work in Process Inventory", "Issue raw materials - Batch 1", 450.00, None),
        ("JE1001", "2026-11-01", "Inventory Asset:Raw Materials Inventory", "Issue raw materials - Batch 1", None, 450.00),
    ], os.path.join(TEST_DIR, "je_good.xlsx"))

    write_xlsx([
        ("JE1002", "2026-11-02", "Cost of Goods Sold:Direct Labor", "Payroll allocation", 500.00, None),
        ("JE1002", "2026-11-02", "Accrued Payroll", "Payroll allocation", None, 450.00),
    ], os.path.join(TEST_DIR, "je_bad_unbalanced.xlsx"))

    write_xlsx([
        ("JE1003", "2026-11-04", "Inventory Asset:Raw Materials Inventory", "Split 3-way materials cost", 33.33, None),
        ("JE1003", "2026-11-04", "Inventory Asset:Raw Materials Inventory", "Split 3-way materials cost", 33.33, None),
        ("JE1003", "2026-11-04", "Inventory Asset:Raw Materials Inventory", "Split 3-way materials cost", 33.33, None),
        ("JE1003", "2026-11-04", "Checking", "Split 3-way materials cost", None, 100.00),
    ], os.path.join(TEST_DIR, "je_bad_rounding.xlsx"))

    write_xlsx([
        ("JE1004", "2026-11-05", "Inventory Asset:Work in Process Inventory", "WIP transfer", 450.00, None),
        ("JE1004", "2026-11-05", "Checking", "WIP transfer", None, 450.00),
        ("JE1004", "2026-11-05", "", "", None, None),
    ], os.path.join(TEST_DIR, "je_bad_phantom_blank_row.xlsx"))

    write_xlsx([
        ("JE1005", "2026-11-06", "Inventory Asset:Work in Process Inventory", "WIP transfer", 450.00, None),
        ("JE1005", "2026-11-06", "Checking", "WIP transfer", None, 450.00),
        ("JE1005", "2026-11-06", "Miscellaneous Holding Account", "", 0.00, 0.00),
    ], os.path.join(TEST_DIR, "je_bad_phantom_zero_row.xlsx"))

    write_xlsx([
        ("JE1006", "2026-11-07", "Factory Supplies", "Shop supplies", 180.00, None),
        ("JE1006", "2026-11-07", "Checking", "Shop supplies", None, 180.00),
    ], os.path.join(TEST_DIR, "je_bad_ambiguous_leaf.xlsx"))

    write_xlsx([
        ("JE1007", "2026-11-08", "Inventory Asset:Work in Process Inventory", "WIP transfer", 450.00, None),
        ("JE1007", "2026-11-08", "Checking", "WIP transfer", None, 450.00),
        (None, "", "Total", "", 450.00, 450.00),
    ], os.path.join(TEST_DIR, "je_bad_trailing_total_row.xlsx"))

    write_xlsx([
        ("JE1008", "2026-11-09", "Inventory Asset:Raw Materials Inventory", "Batch 9 materials, 3 lines", 150.00, None),
        (None, "2026-11-09", "Inventory Asset:Raw Materials Inventory", "Batch 9 materials, 3 lines", 150.00, None),
        (None, "2026-11-09", "Checking", "Batch 9 materials, 3 lines", None, 300.00),
    ], os.path.join(TEST_DIR, "je_bad_broken_grouping.xlsx"))


if __name__ == "__main__":
    write_coa()
    write_sample_demo()
    write_test_files()
    print("Fixtures written.")
