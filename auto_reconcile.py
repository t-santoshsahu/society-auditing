#!/usr/bin/env python3
"""
Automated month-wise reconciliation orchestrator.
===================================================

Given a single month directory (structured like ``test_data/2026-2027/Aug_26``,
containing an ``input/`` folder with the Bank Statement, NBH Bank Book and
NBH GST report), this script:

  1. Detects the month/year from the directory name.
  2. Auto-classifies the three required input files by inspecting their
     contents (no reliance on file naming conventions).
  3. Locates the Vendor Bills workbook (shared across months).
  4. Runs the same matching pipeline as ``reconcile.py`` (expenses, income,
     FD interest, cash withdrawals, other income sources).
    5. Rebuilds ``report_Aug_2026.xlsx`` from the current input files on
      every run, carrying forward only old values that the fresh run leaves
      blank.
  6. Looks for expenses that remain unmatched in the current month across
     every other month directory under the same root (cheques/split
     payments frequently land in a different month) and annotates any hit
     found, including partial matches for bills paid via multiple cheques.

Usage:
    python3 auto_reconcile.py test_data/2026-2027/Aug_26
    python3 auto_reconcile.py test_data/2026-2027/Aug_26 --vendor-bills test_data/2026-2027/VendorBills.xlsx
"""

from __future__ import annotations

import argparse
import re
import os
import sys
from copy import copy
from datetime import date, datetime
from pathlib import Path
from typing import Any, Optional

from openpyxl import load_workbook
from openpyxl.worksheet.datavalidation import DataValidation

import reconcile as rc

MONTH_NAMES = {
    "jan": 1, "january": 1, "feb": 2, "february": 2, "mar": 3, "march": 3,
    "apr": 4, "april": 4, "may": 5, "jun": 6, "june": 6, "jul": 7, "july": 7,
    "aug": 8, "august": 8, "sep": 9, "sept": 9, "september": 9, "oct": 10,
    "october": 10, "nov": 11, "november": 11, "dec": 12, "december": 12,
}

CHEQUE_NUMBER_PATTERN = re.compile(r'\b\d{4,}\b')
QUARTERLY_REFERENCE_PATTERN = re.compile(
    r'^(?:\d{4}(?:-\d{2})?|\d{2}-\d{2})/Q\d+/\d+$',
    re.IGNORECASE,
)


# =============================================================
# MONTH DIRECTORY PARSING
# =============================================================

def parse_month_dir_name(name: str) -> tuple[int, int]:
    match = re.match(r"([A-Za-z]+)[_\-\s]?(\d{2,4})", name)
    if not match:
        raise ValueError(f"Could not determine month/year from directory name: '{name}'")
    month_text, year_text = match.groups()
    month = MONTH_NAMES.get(month_text.lower())
    if not month:
        raise ValueError(f"Unrecognized month name '{month_text}' in directory: '{name}'")
    year = int(year_text)
    if year < 100:
        year += 2000
    return year, month


# =============================================================
# INPUT FILE CLASSIFICATION
# =============================================================

def _has_bank_header(ws) -> bool:
    for r in range(1, min(ws.max_row, 15) + 1):
        headers = {
            rc.clean_string(ws.cell(r, c).value).upper()
            for c in range(1, ws.max_column + 1)
            if ws.cell(r, c).value not in (None, "")
        }
        if "TRANSACTION DATE" in headers and "NARRATION" in headers:
            return True
    return False


def _top_text(ws, rows: int = 5, cols: int = 8) -> str:
    parts = []
    for r in range(1, min(ws.max_row, rows) + 1):
        for c in range(1, min(ws.max_column, cols) + 1):
            value = ws.cell(r, c).value
            if value:
                parts.append(str(value))
    return " ".join(parts).upper()


def classify_input_file(path: Path) -> str:
    """Return one of 'bank', 'income', 'other_income', or 'unknown'."""
    try:
        wb = rc.load_input_workbook(path, data_only=True)
        ws = wb.worksheets[0]
    except Exception:
        return "unknown"

    if _has_bank_header(ws):
        return "bank"

    top_text = _top_text(ws)
    if "BANK BOOK" in top_text:
        return "income"

    for r in range(1, min(ws.max_row, 5) + 1):
        row_headers = {
            rc.clean_string(ws.cell(r, c).value).upper() for c in range(1, ws.max_column + 1)
        }
        if "BILL NUMBER" in row_headers and "LEDGER NAME" in row_headers:
            return "other_income"

    if "INVOICE" in top_text and "GST" in path.name.upper():
        return "other_income"

    return "unknown"


def discover_input_files(input_dir: Path) -> dict[str, Path]:
    if not input_dir.is_dir():
        raise FileNotFoundError(f"Input directory not found: {input_dir}")

    candidates: dict[str, list[Path]] = {"bank": [], "income": [], "other_income": []}
    for f in sorted(input_dir.iterdir()):
        if not f.is_file() or f.suffix.lower() not in {".xlsx", ".xls"} or f.name.startswith("~$"):
            continue
        kind = classify_input_file(f)
        if kind in candidates:
            candidates[kind].append(f)

    labels = {
        "bank": "Bank Statement",
        "income": "NBH Bank Book (Income)",
        "other_income": "NBH GST Report (Other Income)",
    }
    resolved: dict[str, Path] = {}
    for kind, label in labels.items():
        files = candidates[kind]
        if not files:
            raise FileNotFoundError(f"Could not find the {label} file in {input_dir}")
        if len(files) > 1:
            names = ", ".join(f.name for f in files)
            raise ValueError(f"Multiple candidate files found for {label} in {input_dir}: {names}")
        resolved[kind] = files[0]
    return resolved


# =============================================================
# VENDOR BILLS WORKBOOK DISCOVERY
# =============================================================

def find_vendor_bills_file(month_dir: Path, data_root: Path, explicit: Optional[Path]) -> Path:
    if explicit:
        if not explicit.exists():
            raise FileNotFoundError(f"Vendor bills file not found: {explicit}")
        return explicit

    # Search in priority order and stop at the first directory with any match,
    # so an unrelated VendorBills.xlsx higher up the tree doesn't cause a false
    # "ambiguous" error.
    search_dirs = [month_dir, month_dir / "input", data_root, data_root.parent]
    seen: set[Path] = set()
    for d in search_dirs:
        if not d.is_dir():
            continue
        resolved_d = d.resolve()
        if resolved_d in seen:
            continue
        seen.add(resolved_d)

        matches = sorted({
            f.resolve() for f in d.glob("*[Vv]endor*[Bb]ill*.xls*")
            if f.is_file() and not f.name.startswith("~$")
        })
        if not matches:
            continue
        if len(matches) > 1:
            names = ", ".join(str(f) for f in matches)
            raise ValueError(f"Multiple candidate Vendor Bills workbooks found in {d}: {names}")
        return matches[0]

    raise FileNotFoundError(
        "Could not locate a Vendor Bills workbook (expected a file matching "
        f"'*Vendor*Bill*.xlsx') under {month_dir} or {data_root}. "
        "Pass one explicitly with --vendor-bills."
    )


# =============================================================
# SHEET MERGE HELPERS (PRESERVE MANUAL EDITS ACROSS RE-RUNS)
# =============================================================

def _row_signature(ws, row: int, cols: list[int]) -> tuple:
    signature = []
    for c in cols:
        value = ws.cell(row, c).value
        if isinstance(value, (datetime, date)):
            value = value.isoformat() if hasattr(value, "isoformat") else str(value)
        elif isinstance(value, (int, float)):
            value = round(float(value), 2)
        else:
            value = rc.clean_string(value).upper()
        signature.append(value)
    return tuple(signature)


def _carry_forward(old_value: Any, new_value: Any) -> Any:
    """Keep a manually-resolved old value when the fresh run left the cell unresolved."""
    blank_markers = ("", None)
    if new_value not in blank_markers:
        return new_value
    if old_value not in blank_markers:
        return old_value
    return new_value


def _merge_rows(old_ws, old_start_row: int, old_key_cols: list[int],
                 new_ws, new_start_row: int, new_key_cols: list[int],
                 status_cols: list[int]) -> int:
    old_index: dict[tuple, int] = {}
    for r in range(old_start_row + 1, old_ws.max_row + 1):
        key = _row_signature(old_ws, r, old_key_cols)
        if any(part not in ("", None) for part in key):
            old_index.setdefault(key, r)

    merged = 0
    for r in range(new_start_row + 1, new_ws.max_row + 1):
        key = _row_signature(new_ws, r, new_key_cols)
        old_row = old_index.get(key)
        if old_row is None:
            continue
        for c in status_cols:
            old_value = old_ws.cell(old_row, c).value
            new_value = new_ws.cell(r, c).value
            carried = _carry_forward(old_value, new_value)
            if carried != new_value:
                new_ws.cell(r, c).value = carried
                merged += 1
    return merged


def _header_map(ws, header_row: int = 1) -> dict[str, int]:
    return {
        rc.clean_string(ws.cell(header_row, col).value).upper(): col
        for col in range(1, ws.max_column + 1)
        if rc.clean_string(ws.cell(header_row, col).value)
    }


def _row_key(ws, row: int, columns: list[int]) -> tuple:
    return _row_signature(ws, row, columns)


def _index_rows(ws, start_row: int, columns: list[int]) -> dict[tuple, int]:
    index: dict[tuple, int] = {}
    for row in range(start_row, ws.max_row + 1):
        key = _row_key(ws, row, columns)
        if any(value not in ("", None) for value in key):
            index.setdefault(key, row)
    return index


def _copy_columns(old_ws, old_row: int, new_ws, new_row: int, columns: list[int]) -> int:
    copied = 0
    for column in columns:
        old_value = old_ws.cell(old_row, column).value
        if old_value in (None, ""):
            continue
        if new_ws.cell(new_row, column).value != old_value:
            new_ws.cell(new_row, column).value = old_value
            copied += 1
    return copied


def _clone_sheet(dst_wb, src_ws, sheet_name: str):
    del dst_wb[sheet_name]
    new_ws = dst_wb.create_sheet(sheet_name)
    for row in src_ws.iter_rows():
        for cell in row:
            rc.copy_cell(cell, new_ws.cell(cell.row, cell.column))
    for col_letter, dim in src_ws.column_dimensions.items():
        new_ws.column_dimensions[col_letter].width = dim.width
    for merged_range in src_ws.merged_cells.ranges:
        new_ws.merge_cells(str(merged_range))
    for dv in src_ws.data_validations.dataValidation:
        new_dv = DataValidation(
            type=dv.type, formula1=dv.formula1, formula2=dv.formula2, allow_blank=dv.allow_blank
        )
        new_dv.sqref = dv.sqref
        new_ws.add_data_validation(new_dv)
    new_ws.freeze_panes = src_ws.freeze_panes
    return new_ws


def refresh_other_income_categories(other_income_ws) -> bool:
    """Refresh category validation and GST summary in an existing draft."""
    summary_start = next(
        (
            row for row in range(1, other_income_ws.max_row + 1)
            if rc.clean_string(other_income_ws.cell(row, 1).value).upper() == "INCOME TYPE"
            and rc.clean_string(other_income_ws.cell(row, 2).value).upper() == "TOTAL AMOUNT"
        ),
        None,
    )
    total_row = next(
        (
            row for row in range((summary_start or 1) + 1, other_income_ws.max_row + 1)
            if rc.clean_string(other_income_ws.cell(row, 1).value).upper() == "TOTAL GST"
        ),
        None,
    )
    if summary_start is None or total_row is None:
        return False

    manual_entry_end = summary_start - 1
    other_income_ws.delete_rows(summary_start, total_row - summary_start + 1)
    rc.add_other_income_gst_summary(other_income_ws, manual_entry_end)

    other_income_ws.data_validations.dataValidation = []
    income_type_validation = DataValidation(
        type="list",
        formula1=f'"{",".join(rc.OTHER_INCOME_TYPES)}"',
        allow_blank=True,
    )
    income_type_validation.add(f"L2:L{manual_entry_end}")
    other_income_ws.add_data_validation(income_type_validation)
    return True


def migrate_other_income_schema(other_income_ws) -> bool:
    """Convert the legacy 20-column GST sheet to the current 16-column layout."""
    headers = [
        rc.clean_string(other_income_ws.cell(1, col).value).upper()
        for col in range(1, (other_income_ws.max_column or 0) + 1)
    ]
    legacy_columns = ["GST NO", "PLACE OF SUPPLY", "TYPE", "UTILITY"]
    is_legacy = len(headers) >= 20 and headers[3:7] == legacy_columns
    current_columns = [
        "BILL NUMBER", "LEDGER NAME", "DATE", "VOUCHER TYPE", "TOTAL AMOUNT",
        "NON TAXABLE AMOUNT", "TAXABLE AMOUNT", "PERCENTAGE", "CGST", "SGST",
        "IGST", "INCOME TYPE", "BANK DEPOSIT AMOUNT", "BANK REFERENCE ID",
        "BANK TRANSACTION DATE", "MATCH STATUS",
    ]
    if not is_legacy and headers[:16] != current_columns:
        return False

    if is_legacy:
        other_income_ws.delete_cols(4, 4)
    current_headers = [
        "Bill Number", "Ledger Name", "Date", "Voucher Type", "Total Amount",
        "Non Taxable Amount", "Taxable Amount", "Percentage", "CGST", "SGST",
        "IGST", "Income Type", "Bank Deposit Amount", "Bank Reference ID",
        "Bank Transaction Date", "Match Status",
    ]
    for col, header in enumerate(current_headers, 1):
        other_income_ws.cell(1, col).value = header

    summary_start = next(
        (
            row for row in range(2, (other_income_ws.max_row or 1) + 1)
            if rc.clean_string(other_income_ws.cell(row, 1).value).upper() == "INCOME TYPE"
            and rc.clean_string(other_income_ws.cell(row, 2).value).upper() == "TOTAL AMOUNT"
        ),
        None,
    )
    data_end = (summary_start - 1) if summary_start else (other_income_ws.max_row or 1)
    for row in range(2, data_end + 1):
        has_data = any(
            other_income_ws.cell(row, col).value not in (None, "")
            for col in range(1, 17)
        )
        if not has_data:
            continue
        other_income_ws.cell(row, 7).value = f'=ROUND((E{row}-F{row})/(1+2*H{row}),2)'
        other_income_ws.cell(row, 9).value = f'=ROUND(G{row}*H{row},2)'
        other_income_ws.cell(row, 10).value = f'=ROUND(G{row}*H{row},2)'
        other_income_ws.cell(row, 11).value = 0.0

    if summary_start:
        total_row = next(
            (
                row for row in range(summary_start + 1, (other_income_ws.max_row or 1) + 1)
                if rc.clean_string(other_income_ws.cell(row, 1).value).upper() == "TOTAL GST"
            ),
            None,
        )
        if total_row:
            other_income_ws.delete_rows(summary_start, total_row - summary_start + 1)
    if summary_start:
        rc.add_other_income_gst_summary(other_income_ws, data_end)

    other_income_ws.data_validations.dataValidation = []
    income_type_validation = DataValidation(
        type="list",
        formula1=f'"{",".join(rc.OTHER_INCOME_TYPES)}"',
        allow_blank=True,
    )
    income_type_validation.add(f"L2:L{data_end}")
    other_income_ws.add_data_validation(income_type_validation)
    rc.style_header(other_income_ws, 1)
    rc.autofit(other_income_ws)
    return True


def _build_bank_meta(bank_ws) -> dict:
    header_row, headers = rc.locate_bank_header(bank_ws)
    return {
        "header_row": header_row,
        "date_col": headers.get("TRANSACTION DATE", 1),
        "narration_col": headers.get("NARRATION", 2),
        "ref_col": headers.get("CHEQUE NUMBER/SETTLEMENT ID", 3),
        "debit_col": headers.get("DEBIT", 4),
        "credit_col": headers.get("CREDIT", 5),
    }


def _color_rows_by_status(ws, status_col: int, rows: list[int], columns: range) -> int:
    colored = 0
    for r in rows:
        status = rc.clean_string(ws.cell(r, status_col).value).upper()
        if status.startswith("MATCHED") or status == "FD TRANSFER":
            for c in columns:
                ws.cell(r, c).fill = copy(rc.GREEN_FILL)
            colored += 1
        elif status.startswith("PARTIALLY MATCHED"):
            for c in columns:
                ws.cell(r, c).fill = copy(rc.YELLOW_FILL)
        elif status.startswith(("UNMATCHED", "REVIEW", "PENDING")) or status == "NOT EXPENSE":
            for c in columns:
                ws.cell(r, c).fill = copy(rc.RED_FILL)
    return colored


def expense_payment_amount(expense_ws, row: int) -> float:
    expense_group = rc.clean_string(expense_ws.cell(row, 19).value).upper()
    vendor_name = rc.clean_string(expense_ws.cell(row, rc.VENDOR_COL_NAME).value).upper()
    if expense_group == "BANK CHARGES" or vendor_name == "CHEQUE BOOK ISSUE CHARGE":
        bank_paid = rc.numeric_amount(expense_ws.cell(row, rc.EXPENSE_COL_BANK_PAID).value)
        if bank_paid > 0:
            return bank_paid
    amount_paid_value = expense_ws.cell(row, rc.VENDOR_COL_AMOUNT_PAID).value
    amount_paid = rc.numeric_amount(amount_paid_value)
    if amount_paid > 0 and not (isinstance(amount_paid_value, str) and amount_paid_value.startswith("=")):
        return amount_paid
    amount = rc.numeric_amount(expense_ws.cell(row, rc.VENDOR_COL_AMOUNT).value)
    sgst = rc.numeric_amount(expense_ws.cell(row, rc.VENDOR_COL_SGST).value)
    cgst = rc.numeric_amount(expense_ws.cell(row, rc.VENDOR_COL_CGST).value)
    tds_rate = rc.numeric_amount(expense_ws.cell(row, rc.VENDOR_COL_TDS_RATE).value)
    return round(amount + sgst + cgst - round(amount * tds_rate / 100, 0), 2)


def is_expense_detail_row(expense_ws, row: int) -> bool:
    vendor = expense_ws.cell(row, rc.VENDOR_COL_NAME).value
    return not (isinstance(vendor, str) and vendor.startswith("=")) and rc.is_real_expense_row(expense_ws, row)


def reference_parts(value: object) -> list[str]:
    return [part.strip().upper() for part in rc.clean_string(value).replace(";", ",").split(",") if part.strip()]


def is_legacy_balance_reference(value: object) -> bool:
    return bool(re.fullmatch(r"[\d,]+(?:\.\d+)?\s*[DC]", rc.clean_string(value).upper()))


def compact_fd_transactions(ws) -> None:
    totals_gap = 10
    summary_row = next(
        (
            row for row in range(2, ws.max_row + 1)
            if rc.clean_string(ws.cell(row, 5).value).upper() == "FD TRANSACTION TOTALS"
        ),
        None,
    )
    if summary_row is None:
        return

    empty_rows = [
        row for row in range(2, summary_row)
        if not any(ws.cell(row, column).value not in (None, "") for column in range(1, 9))
    ]
    for row in reversed(empty_rows):
        ws.delete_rows(row, 1)

    data_rows = [
        row for row in range(2, ws.max_row + 1)
        if rc.clean_string(ws.cell(row, 3).value)
    ]
    summary_row = next(
        row for row in range(2, ws.max_row + 1)
        if rc.clean_string(ws.cell(row, 5).value).upper() == "FD TRANSACTION TOTALS"
    )
    last_data_row = max(data_rows, default=1)
    desired_summary_row = last_data_row + totals_gap
    if summary_row < desired_summary_row:
        ws.insert_rows(summary_row, amount=desired_summary_row - summary_row)
        summary_row = desired_summary_row
    elif summary_row > desired_summary_row:
        ws.delete_rows(desired_summary_row, amount=summary_row - desired_summary_row)
        summary_row = desired_summary_row
    ws.cell(summary_row + 1, 6).value = f"=SUM(F2:F{last_data_row})"
    ws.cell(summary_row + 2, 7).value = f"=SUM(G2:G{last_data_row})"


def link_expense_references_with_bank(expense_ws, bank_ws, bank_meta: dict) -> int:
    """Match expenses by user reference, then unique amount within Rs 1."""
    if expense_ws.max_column < rc.EXPENSE_COL_BANK_PAID:
        expense_ws.cell(1, rc.EXPENSE_COL_BANK_PAID).value = "Bank Amount Paid"

    rows_by_ref: dict[str, list[int]] = {}
    for row in range(2, expense_ws.max_row + 1):
        if not is_expense_detail_row(expense_ws, row):
            continue
        references = reference_parts(expense_ws.cell(row, 16).value)
        for reference in references:
            if rc.clean_string(expense_ws.cell(row, rc.VENDOR_COL_NAME).value):
                rows_by_ref.setdefault(reference, []).append(row)

    bank_rows_by_ref: dict[str, list[int]] = {}
    header_row = bank_meta["header_row"]
    last_row = rc.find_last_bank_data_row(bank_ws, header_row)
    for row in range(header_row + 1, last_row + 1):
        reference = rc.clean_string(bank_ws.cell(row, 6).value).upper()
        if reference:
            bank_rows_by_ref.setdefault(reference, []).append(row)

    linked = 0
    processed_expense_rows: set[int] = set()
    for reference, reference_expense_rows in rows_by_ref.items():
        expense_rows = [row for row in reference_expense_rows if row not in processed_expense_rows]
        if not expense_rows:
            continue
        reference_values = reference_parts(expense_ws.cell(expense_rows[0], 16).value)
        bank_rows = list(dict.fromkeys(row for value in reference_values for row in bank_rows_by_ref.get(value, [])))
        if len(bank_rows) != len(reference_values):
            continue
        bank_amount = sum(rc.numeric_amount(bank_ws.cell(row, bank_meta["debit_col"]).value) for row in bank_rows)
        expected_amount = sum(expense_payment_amount(expense_ws, row) for row in expense_rows)
        difference = round(bank_amount - expected_amount, 2)
        if abs(difference) <= 1:
            status = "MATCHED"
            fill = rc.GREEN_FILL
        else:
            status = f"PARTIALLY MATCHED (Difference: {difference:,.2f})"
            fill = rc.YELLOW_FILL
        for row in expense_rows:
            expense_ws.cell(row, 18).value = status
            expense_ws.cell(row, rc.EXPENSE_COL_BANK_PAID).value = bank_amount
            for col in range(1, 20):
                expense_ws.cell(row, col).fill = copy(fill)
        for bank_row in bank_rows:
            bank_ws.cell(bank_row, 8).value = f"Expense reference {rc.clean_string(expense_ws.cell(expense_rows[0], 16).value)} | {status}"
            bank_ws.cell(bank_row, 9).value = status
            for col in range(1, 10):
                bank_ws.cell(bank_row, col).fill = copy(fill)
        processed_expense_rows.update(expense_rows)
        linked += 1

    # Retry unresolved red/yellow rows by amount on every reconciliation.
    used_rows = {row for rows in rows_by_ref.values() for row in rows}
    for bank_row in range(header_row + 1, last_row + 1):
        bank_status = rc.clean_string(bank_ws.cell(bank_row, 9).value).upper()
        if bank_status.startswith("MATCHED"):
            continue
        bank_amount = rc.numeric_amount(bank_ws.cell(bank_row, bank_meta["debit_col"]).value)
        if bank_amount <= 0:
            continue
        candidates = []
        for row in range(2, expense_ws.max_row + 1):
            if row in used_rows or not rc.clean_string(expense_ws.cell(row, rc.VENDOR_COL_NAME).value):
                continue
            if rc.clean_string(expense_ws.cell(row, 19).value).upper() == "TAX AND DUTIES":
                continue
            if abs(expense_payment_amount(expense_ws, row) - bank_amount) <= 1:
                candidates.append(row)
        if len(candidates) != 1:
            continue
        top_row = candidates[0]
        expected_amount = expense_payment_amount(expense_ws, top_row)
        difference = round(bank_amount - expected_amount, 2)
        if abs(difference) <= 1:
            status = "MATCHED"
            fill = rc.GREEN_FILL
        else:
            status = f"PARTIALLY MATCHED (Difference: {difference:,.2f})"
            fill = rc.YELLOW_FILL
        expense_ws.cell(top_row, 16).value = rc.clean_string(bank_ws.cell(bank_row, 6).value)
        expense_ws.cell(top_row, 17).value = bank_ws.cell(bank_row, bank_meta["date_col"]).value
        expense_ws.cell(top_row, 18).value = status
        expense_ws.cell(top_row, rc.EXPENSE_COL_BANK_PAID).value = bank_amount
        for col in range(1, rc.EXPENSE_COL_BANK_PAID + 1):
            expense_ws.cell(top_row, col).fill = copy(fill)
        bank_ws.cell(bank_row, 7).value = expense_ws.cell(top_row, rc.VENDOR_COL_BILL_NO).value
        bank_ws.cell(bank_row, 8).value = f"{expense_ws.cell(top_row, rc.VENDOR_COL_NAME).value} | {status}"
        bank_ws.cell(bank_row, 9).value = status
        for col in range(1, 10):
            bank_ws.cell(bank_row, col).fill = copy(fill)
        used_rows.add(top_row)
        linked += 1
    return linked


def link_other_income_with_bank(other_income_ws, bank_ws, bank_meta: dict) -> int:
    """Propagate Income from Other Sources matches into the Bank_Statement sheet.

    ``process_other_income_sources`` never wrote back to Bank_Statement, so a
    deposit resolved here (by reference, or manually categorised by the user)
    never showed as MATCHED/green on the bank sheet. This closes that loop in
    both directions: bank reference -> other income row, and vice versa.
    """
    header_row = bank_meta["header_row"]
    last_row = rc.find_last_bank_data_row(bank_ws, header_row)

    bank_rows_by_ref: dict[str, list[int]] = {}
    for r in range(header_row + 1, last_row + 1):
        ref = rc.clean_string(bank_ws.cell(r, 6).value).upper()
        if ref:
            bank_rows_by_ref.setdefault(ref, []).append(r)

    linked = 0
    for r in range(2, other_income_ws.max_row + 1):
        bill_no = rc.clean_string(other_income_ws.cell(r, 1).value)
        ledger = rc.clean_string(other_income_ws.cell(r, 2).value)
        income_type = rc.clean_string(other_income_ws.cell(r, 12).value)
        bank_amount = rc.numeric_amount(other_income_ws.cell(r, 13).value)
        bank_ref = rc.clean_string(other_income_ws.cell(r, 14).value).upper()
        status = rc.clean_string(other_income_ws.cell(r, 16).value).upper()

        if not bank_ref and bank_amount <= 0:
            continue  # blank manual-entry slot or summary row, not real data

        # Auto-created deposit rows are only candidates until the user supplies
        # both identifying details required for an other-income match.
        if not ledger or not income_type:
            if status == "MATCHED":
                other_income_ws.cell(r, 16).value = "PENDING"
            if bank_ref:
                stale_bank_rows = bank_rows_by_ref.get(bank_ref, [])
                if len(stale_bank_rows) == 1:
                    bank_row = stale_bank_rows[0]
                    if rc.clean_string(bank_ws.cell(bank_row, 8).value).startswith("Other Income:"):
                        bank_ws.cell(bank_row, 7).value = None
                        bank_ws.cell(bank_row, 8).value = None
            continue

        candidates = bank_rows_by_ref.get(bank_ref, []) if bank_ref else []
        if not candidates and bank_amount > 0:
            candidates = [
                rr for rr in range(header_row + 1, last_row + 1)
                if rc.money_equal(rc.numeric_amount(bank_ws.cell(rr, bank_meta["credit_col"]).value), bank_amount)
            ]

        if len(candidates) != 1:
            continue
        bank_row = candidates[0]

        current_bank_status = rc.clean_string(bank_ws.cell(bank_row, 9).value).upper()
        if current_bank_status not in {"MATCHED", "FD TRANSFER"}:
            description = ledger or bill_no or "Income from Other Sources"
            bank_ws.cell(bank_row, 7).value = bill_no or bank_ws.cell(bank_row, 7).value
            bank_ws.cell(bank_row, 8).value = f"Other Income: {description}"
            bank_ws.cell(bank_row, 9).value = "MATCHED"
            for c in range(1, 10):
                bank_ws.cell(bank_row, c).fill = copy(rc.GREEN_FILL)
            linked += 1

        if status in {"", "PENDING", "UNMATCHED"}:
            other_income_ws.cell(r, 16).value = "MATCHED"
            if not bank_ref:
                other_income_ws.cell(r, 14).value = rc.clean_string(bank_ws.cell(bank_row, 6).value)
            if not other_income_ws.cell(r, 15).value:
                other_income_ws.cell(r, 15).value = bank_ws.cell(bank_row, bank_meta["date_col"]).value

    return linked


def link_recorded_income_with_bank(income_ws, bank_ws, bank_meta: dict) -> int:
    """Resolve NBH income rows using settlement/reference IDs and amounts."""
    headers = {
        rc.clean_string(income_ws.cell(1, col).value).upper(): col
        for col in range(1, income_ws.max_column + 1)
    }
    status_col = headers.get("BANK MATCH STATUS")
    reference_cols = [
        headers[name]
        for name in ("SETTLEMENT ID", "TXN ID / CHEQUE NO.", "BANK REFERENCE ID", "REFERENCE NUMBER")
        if headers.get(name)
    ]
    amount_cols = [headers[name] for name in ("CREDIT", "DEBIT") if headers.get(name)]
    flat_col = headers.get("FLAT NO.")
    reference_number_col = headers.get("REFERENCE NUMBER")
    date_col = headers.get("BANK TRANSACTION DATE")
    if not reference_cols or not status_col:
        return 0

    header_row = bank_meta["header_row"]
    last_row = rc.find_last_bank_data_row(bank_ws, header_row)
    bank_rows_by_ref: dict[str, list[int]] = {}
    for row in range(header_row + 1, last_row + 1):
        reference = rc.clean_string(bank_ws.cell(row, 6).value).upper()
        if reference:
            bank_rows_by_ref.setdefault(reference, []).append(row)

    rows_by_bank_row: dict[int, list[int]] = {}
    for row in range(2, income_ws.max_row + 1):
        references = []
        for reference_col in reference_cols:
            references.extend(reference_parts(income_ws.cell(row, reference_col).value))
        references = list(dict.fromkeys(references))
        if not references:
            continue
        candidates = []
        for reference in references:
            reference_candidates = bank_rows_by_ref.get(reference, [])
            if len(reference_candidates) == 1:
                candidates = reference_candidates
                break
        if len(candidates) != 1:
            continue
        rows_by_bank_row.setdefault(candidates[0], []).append(row)

    linked = 0
    for bank_row, income_rows in rows_by_bank_row.items():
        income_amount = sum(
            next(
                (rc.numeric_amount(income_ws.cell(row, amount_col).value) for amount_col in amount_cols
                 if rc.numeric_amount(income_ws.cell(row, amount_col).value) > 0),
                0.0,
            )
            for row in income_rows
        )
        bank_amounts = (
            rc.numeric_amount(bank_ws.cell(bank_row, bank_meta["credit_col"]).value),
            rc.numeric_amount(bank_ws.cell(bank_row, bank_meta["debit_col"]).value),
        )
        if income_amount > 0 and not any(rc.money_equal(income_amount, amount) for amount in bank_amounts):
            continue

        bank_reference = rc.clean_string(bank_ws.cell(bank_row, 6).value)
        descriptions = []
        for row in income_rows:
            income_ws.cell(row, status_col).value = "MATCHED"
            if headers.get("BANK REFERENCE ID"):
                income_ws.cell(row, headers["BANK REFERENCE ID"]).value = bank_reference
            if date_col and not income_ws.cell(row, date_col).value:
                income_ws.cell(row, date_col).value = bank_ws.cell(bank_row, bank_meta["date_col"]).value
            for col in range(1, income_ws.max_column + 1):
                income_ws.cell(row, col).fill = copy(rc.GREEN_FILL)
            description = []
            if flat_col and rc.clean_string(income_ws.cell(row, flat_col).value):
                description.append(rc.clean_string(income_ws.cell(row, flat_col).value))
            if reference_number_col and rc.clean_string(income_ws.cell(row, reference_number_col).value):
                description.append(f"Ref: {rc.clean_string(income_ws.cell(row, reference_number_col).value)}")
            if description:
                text = " | ".join(description)
                if text not in descriptions:
                    descriptions.append(text)
        description = " | ".join(descriptions)
        if description:
            bank_ws.cell(bank_row, 8).value = description
        bank_ws.cell(bank_row, 9).value = "MATCHED"
        for col in range(1, 10):
            bank_ws.cell(bank_row, col).fill = copy(rc.GREEN_FILL)
        linked += 1

    return linked


def link_cash_withdrawals_with_bank(wb, bank_ws, bank_meta: dict) -> int:
    """Mark populated cash withdrawal registers on their bank rows."""
    header_row = bank_meta["header_row"]
    last_row = rc.find_last_bank_data_row(bank_ws, header_row)
    bank_rows_by_ref = {
        rc.clean_string(bank_ws.cell(r, 6).value).upper(): r
        for r in range(header_row + 1, last_row + 1)
        if rc.clean_string(bank_ws.cell(r, 6).value)
    }
    linked = 0

    for sheet_name in wb.sheetnames:
        if not sheet_name.startswith("Cash Withdrawal "):
            continue
        cash_ws = wb[sheet_name]
        reference = rc.clean_string(cash_ws["E1"].value).upper()
        withdrawal_amount = rc.numeric_amount(cash_ws["H1"].value)
        if not reference or withdrawal_amount <= 0:
            continue

        payment_header_row = next(
            (row for row in range(1, cash_ws.max_row + 1)
             if any(rc.clean_string(cash_ws.cell(row, col).value).upper() == "PAYMENT AMOUNT"
                    for col in range(1, cash_ws.max_column + 1))),
            None,
        )
        total_paid_row = next(
            (row for row in range(1, cash_ws.max_row + 1)
             if any(rc.clean_string(cash_ws.cell(row, col).value).upper() == "TOTAL PAID"
                    for col in range(1, cash_ws.max_column + 1))),
            None,
        )
        if not payment_header_row or not total_paid_row:
            continue

        payment_col = next(
            col for col in range(1, cash_ws.max_column + 1)
            if rc.clean_string(cash_ws.cell(payment_header_row, col).value).upper() == "PAYMENT AMOUNT"
        )
        payment_entries = sum(
            1 for row in range(payment_header_row + 1, total_paid_row)
            if rc.numeric_amount(cash_ws.cell(row, payment_col).value) > 0
        )
        if not payment_entries:
            continue

        bank_row = bank_rows_by_ref.get(reference)
        if bank_row is None:
            continue
        bank_ws.cell(bank_row, 8).value = f"Cash withdrawal | {sheet_name}"
        bank_ws.cell(bank_row, 9).value = "MATCHED"
        for col in range(1, 10):
            bank_ws.cell(bank_row, col).fill = copy(rc.GREEN_FILL)
        linked += 1

    return linked


def reapply_match_highlighting(bank_ws, header_row: int) -> int:
    """Re-color rows green whenever the Match Status resolves to a matched state.

    This runs after merging, so rows whose match status/description came from a
    previously-resolved draft (rather than this run's own auto-matching) still
    get the green highlight applied consistently.
    """
    last_row = rc.find_last_bank_data_row(bank_ws, header_row)
    colored = 0
    for r in range(header_row + 1, last_row + 1):
        status = rc.clean_string(bank_ws.cell(r, 9).value).upper()
        if status.startswith("MATCHED") or status == "FD TRANSFER":
            for col in range(1, 10):
                bank_ws.cell(r, col).fill = copy(rc.GREEN_FILL)
            colored += 1
        elif status.startswith("PARTIALLY MATCHED"):
            for col in range(1, 10):
                bank_ws.cell(r, col).fill = copy(rc.YELLOW_FILL)
        elif status.startswith(("UNMATCHED", "REVIEW", "PENDING")) or status == "NOT EXPENSE":
            for col in range(1, 10):
                bank_ws.cell(r, col).fill = copy(rc.RED_FILL)
    return colored


def merge_draft(old_wb, new_wb) -> None:
    """Carry only stable identifiers from the previous draft into a fresh one."""

    old_fd_references = set()
    if "FD Transactions" in old_wb.sheetnames:
        old_fd_ws = old_wb["FD Transactions"]
        old_fd_references = {
            reference.upper()
            for row in range(2, old_fd_ws.max_row + 1)
            for reference in reference_parts(old_fd_ws.cell(row, 3).value)
            if reference
        }

    if "Bank_Statement" in old_wb.sheetnames and "Bank_Statement" in new_wb.sheetnames:
        old_ws, new_ws = old_wb["Bank_Statement"], new_wb["Bank_Statement"]
        old_header, old_headers = rc.locate_bank_header(old_ws)
        new_header, new_headers = rc.locate_bank_header(new_ws)
        old_keys = [old_headers.get("TRANSACTION DATE", 1), old_headers.get("NARRATION", 2),
                    old_headers.get("DEBIT", 4), old_headers.get("CREDIT", 5)]
        new_keys = [new_headers.get("TRANSACTION DATE", 1), new_headers.get("NARRATION", 2),
                    new_headers.get("DEBIT", 4), new_headers.get("CREDIT", 5)]
        old_index = _index_rows(old_ws, old_header + 1, old_keys)
        merged = 0
        for row in range(new_header + 1, new_ws.max_row + 1):
            old_row = old_index.get(_row_key(new_ws, row, new_keys))
            if old_row is not None:
                merged += _copy_columns(old_ws, old_row, new_ws, row, [6])
        colored = reapply_match_highlighting(new_ws, new_header)
        print(f"Bank_Statement           : preserved {merged} generated reference ID(s), highlighted {colored} matched row(s)")

    if "Expense" in old_wb.sheetnames and "Expense" in new_wb.sheetnames:
        old_ws, new_ws = old_wb["Expense"], new_wb["Expense"]
        key_cols = [rc.VENDOR_COL_BILL_NO, rc.VENDOR_COL_NAME, rc.VENDOR_COL_AMOUNT]
        old_index = _index_rows(old_ws, 2, key_cols)
        old_reference_index: dict[str, int] = {}
        for old_row in range(2, old_ws.max_row + 1):
            for reference in reference_parts(old_ws.cell(old_row, 16).value):
                old_reference_index.setdefault(reference, old_row)
        merged = 0
        for row in range(2, new_ws.max_row + 1):
            old_row = None
            for reference in reference_parts(new_ws.cell(row, 16).value):
                old_row = old_reference_index.get(reference)
                if old_row is not None:
                    break
            if old_row is None:
                old_row = old_index.get(_row_key(new_ws, row, key_cols))
            if old_row is not None:
                merged += _copy_columns(old_ws, old_row, new_ws, row, [5, 16])
        removed = 0
        for row in range(new_ws.max_row, 1, -1):
            payment_references = {
                reference.upper()
                for reference in reference_parts(new_ws.cell(row, 16).value)
            }
            if payment_references & old_fd_references:
                new_ws.delete_rows(row, 1)
                removed += 1
        print(f"Expense                  : preserved {merged} Payment Reference ID(s), removed {removed} FD-moved row(s)")

    if "Income_recorded_in_nbh" in old_wb.sheetnames and "Income_recorded_in_nbh" in new_wb.sheetnames:
        old_ws, new_ws = old_wb["Income_recorded_in_nbh"], new_wb["Income_recorded_in_nbh"]
        old_headers = _header_map(old_ws)
        new_headers = _header_map(new_ws)
        key_names = ("DATE", "FLAT NO.", "PARTICULARS", "TXN ID / CHEQUE NO.", "DEBIT", "CREDIT")
        if all(name in old_headers and name in new_headers for name in key_names):
            old_keys = [old_headers[name] for name in key_names]
            new_keys = [new_headers[name] for name in key_names]
            old_index = _index_rows(old_ws, 2, old_keys)
            id_names = ("SETTLEMENT ID", "TXN ID / CHEQUE NO.", "BANK REFERENCE ID")
            merged = 0
            for row in range(2, new_ws.max_row + 1):
                old_row = old_index.get(_row_key(new_ws, row, new_keys))
                if old_row is None:
                    continue
                columns = [new_headers[name] for name in id_names if name in new_headers and name in old_headers]
                if "BANK REFERENCE ID" in old_headers and is_legacy_balance_reference(
                    old_ws.cell(old_row, old_headers["BANK REFERENCE ID"]).value
                ):
                    columns = [column for column in columns if column != new_headers.get("BANK REFERENCE ID")]
                merged += _copy_columns(old_ws, old_row, new_ws, row, columns)
            print(f"Income_recorded_in_nbh   : preserved {merged} identifier(s) by stable row fields")

    if "Income from Other Sources" in old_wb.sheetnames and "Income from Other Sources" in new_wb.sheetnames:
        old_ws, new_ws = old_wb["Income from Other Sources"], new_wb["Income from Other Sources"]
        old_index = _index_rows(old_ws, 2, [1])
        generated_bill_numbers = {
            rc.clean_string(old_ws.cell(row, 17).value).upper()
            for row in range(2, old_ws.max_row + 1)
            if rc.clean_string(old_ws.cell(row, 17).value)
            and rc.clean_string(old_ws.cell(row, 14).value)
        }
        removed = 0
        for row in range(new_ws.max_row, 1, -1):
            bill_number = rc.clean_string(new_ws.cell(row, 1).value).upper()
            if bill_number and bill_number in generated_bill_numbers:
                new_ws.delete_rows(row, 1)
                removed += 1
        merged = 0
        for row in range(2, new_ws.max_row + 1):
            old_row = old_index.get(_row_key(new_ws, row, [1]))
            if old_row is not None:
                merged += _copy_columns(old_ws, old_row, new_ws, row, [2, 12, 14, 17])
        print(f"Income from Other Sources: preserved {merged} values/identifier(s), removed {removed} merged duplicate row(s)")

    if "FD Transactions" in old_wb.sheetnames and "FD Transactions" in new_wb.sheetnames:
        old_ws, new_ws = old_wb["FD Transactions"], new_wb["FD Transactions"]
        old_index = _index_rows(old_ws, 2, [3])
        fresh_refs = {
            rc.clean_string(new_ws.cell(row, 3).value).upper()
            for row in range(2, new_ws.max_row + 1)
            if rc.clean_string(new_ws.cell(row, 3).value)
        }
        merged = 0
        for row in range(2, new_ws.max_row + 1):
            old_row = old_index.get(_row_key(new_ws, row, [3]))
            if old_row is None:
                continue
            merged += _copy_columns(old_ws, old_row, new_ws, row, list(range(1, 9)))
            new_ws.cell(row, 5).value = new_ws.cell(row, 3).value
        for row in range(2, new_ws.max_row + 1):
            if rc.clean_string(new_ws.cell(row, 3).value):
                new_ws.cell(row, 5).value = new_ws.cell(row, 3).value

        unmatched_fd_rows = []
        for old_row in range(2, old_ws.max_row + 1):
            fd_reference = rc.clean_string(old_ws.cell(old_row, 3).value)
            if fd_reference and fd_reference.upper() not in fresh_refs:
                unmatched_fd_rows.append((old_row, fd_reference))

        appended_refs: set[str] = set()
        if unmatched_fd_rows:
            summary_row = next(
                (
                    row for row in range(2, new_ws.max_row + 1)
                    if rc.clean_string(new_ws.cell(row, 5).value).upper() == "FD TRANSACTION TOTALS"
                ),
                new_ws.max_row + 1,
            )
            new_ws.insert_rows(summary_row, amount=len(unmatched_fd_rows))
            for offset, (old_row, fd_reference) in enumerate(unmatched_fd_rows):
                target_row = summary_row + offset
                for column in range(1, old_ws.max_column + 1):
                    rc.copy_cell(old_ws.cell(old_row, column), new_ws.cell(target_row, column))
                new_ws.cell(target_row, 5).value = fd_reference
                fresh_refs.add(fd_reference.upper())
                appended_refs.add(fd_reference.upper())
                merged += 1

        if appended_refs and "Bank_Statement" in new_wb.sheetnames:
            bank_ws = new_wb["Bank_Statement"]
            bank_header, _ = rc.locate_bank_header(bank_ws)
            bank_last = rc.find_last_bank_data_row(bank_ws, bank_header)
            for bank_row in range(bank_header + 1, bank_last + 1):
                bank_reference = rc.clean_string(bank_ws.cell(bank_row, 6).value).upper()
                if bank_reference not in appended_refs:
                    continue
                bank_ws.cell(bank_row, 8).value = f"FD Transfer: {bank_reference}"
                bank_ws.cell(bank_row, 9).value = "FD TRANSFER"
                for column in range(1, 10):
                    bank_ws.cell(bank_row, column).fill = copy(rc.GREEN_FILL)

        if appended_refs and "Income from Other Sources" in new_wb.sheetnames:
            other_ws = new_wb["Income from Other Sources"]
            for row in range(other_ws.max_row, 1, -1):
                bank_reference = rc.clean_string(other_ws.cell(row, 14).value).upper()
                if bank_reference in appended_refs:
                    other_ws.delete_rows(row, 1)
        compact_fd_transactions(new_ws)
        print(f"FD Transactions          : preserved {merged} existing FD row value(s)")

    # Cash withdrawal registers are fully manual entry areas - keep them intact.
    cloned = 0
    for sheet_name in old_wb.sheetnames:
        if sheet_name.startswith("Cash Withdrawal") and sheet_name in new_wb.sheetnames:
            _clone_sheet(new_wb, old_wb[sheet_name], sheet_name)
            cloned += 1
    if cloned:
        print(f"Cash Withdrawal sheets   : preserved {cloned} manually-filled register(s)")


# =============================================================
# CROSS-MONTH / SPLIT-PAYMENT ENRICHMENT
# =============================================================

def discover_sibling_month_dirs(root: Path, exclude: Path) -> list[Path]:
    dirs = []
    exclude_resolved = exclude.resolve()
    if not root.is_dir():
        return dirs
    for d in sorted(root.iterdir()):
        if not d.is_dir() or d.resolve() == exclude_resolved:
            continue
        try:
            parse_month_dir_name(d.name)
        except ValueError:
            continue
        if (d / "input").is_dir():
            dirs.append(d)
    return dirs


def build_cross_month_bank_index(month_dirs: list[Path], drafts_only: bool = False) -> list[dict]:
    """Flatten sibling bank statements from current input workbooks."""
    index: list[dict] = []
    for month_dir in month_dirs:
        try:
            files = discover_input_files(month_dir / "input")
            bank_wb = rc.load_input_workbook(files["bank"])
            bank_ws = bank_wb.worksheets[0]
        except Exception:
            continue

        header_row, headers = rc.locate_bank_header(bank_ws)
        date_col = headers.get("TRANSACTION DATE", 1)
        narration_col = headers.get("NARRATION", 2)
        ref_col = headers.get("CHEQUE NUMBER/SETTLEMENT ID", 3)
        debit_col = headers.get("DEBIT", 4)
        credit_col = headers.get("CREDIT", 5)
        last_row = rc.find_last_bank_data_row(bank_ws, header_row)

        for r in range(header_row + 1, last_row + 1):
            narration = bank_ws.cell(r, narration_col).value
            orig_ref = bank_ws.cell(r, ref_col).value
            debit = rc.numeric_amount(bank_ws.cell(r, debit_col).value)
            credit = rc.numeric_amount(bank_ws.cell(r, credit_col).value)
            if debit <= 0 and credit <= 0:
                continue
            index.append({
                "month": month_dir.name,
                "row": r,
                "date": rc.parse_date(bank_ws.cell(r, date_col).value),
                "narration": rc.clean_string(narration).upper(),
                "ref": rc.generate_reference_id(
                    narration,
                    orig_ref,
                    bank_ws.cell(r, date_col).value,
                    debit,
                    credit,
                ).upper(),
                "debit": debit,
                "credit": credit,
            })
    return index


def _extract_cheque_numbers(text: str) -> list[str]:
    if not text or not re.search(r'CHQ|CHEQUE', text.upper()):
        return []
    return CHEQUE_NUMBER_PATTERN.findall(text.upper())


def annotate_cross_month_draft(
    root: Path, month_name: str, generated_ref: str, note: str,
    status: str = "MATCHED (Cross-Month Reference)",
) -> bool:
    """Leave a matching reference note on another month's draft, if it already exists.

    Only annotates a pre-existing draft (never creates/regenerates one), and
    only touches the Match Status/fill when that row isn't already MATCHED.
    """
    other_draft = resolve_draft_path(root / month_name)
    if not other_draft.exists():
        return False
    try:
        other_wb = load_workbook(other_draft)
        if "Bank_Statement" not in other_wb.sheetnames:
            return False
        other_ws = other_wb["Bank_Statement"]
        header_row, _ = rc.locate_bank_header(other_ws)
        last_row = rc.find_last_bank_data_row(other_ws, header_row)

        updated = False
        for r in range(header_row + 1, last_row + 1):
            if rc.clean_string(other_ws.cell(r, 6).value).upper() != generated_ref.upper():
                continue
            existing_desc = rc.clean_string(other_ws.cell(r, 8).value)
            if note not in existing_desc:
                other_ws.cell(r, 8).value = f"{existing_desc} | {note}".strip(" |") if existing_desc else note
            existing_status = rc.clean_string(other_ws.cell(r, 9).value).upper()
            if existing_status != "FD TRANSFER":
                other_ws.cell(r, 9).value = status
                for c in range(1, 10):
                    other_ws.cell(r, c).fill = copy(rc.GREEN_FILL)
            updated = True
            break

        if updated:
            other_wb.save(other_draft)
        return updated
    except Exception:
        return False


def enrich_unmatched_expenses(
    wb, bank_index: list[dict], current_month_name: str, root: Optional[Path] = None
) -> dict[str, int]:
    """Search other months' bank statements for expenses left unmatched this month."""
    if "Expense" not in wb.sheetnames or not bank_index:
        return {"cross_month": 0, "split": 0}

    expense_ws = wb["Expense"]
    current_bank_refs = set()
    current_bank_transactions: dict[str, dict] = {}
    if "Bank_Statement" in wb.sheetnames:
        current_bank_ws = wb["Bank_Statement"]
        current_header_row, current_headers = rc.locate_bank_header(current_bank_ws)
        current_date_col = current_headers.get("TRANSACTION DATE", 1)
        current_narration_col = current_headers.get("NARRATION", 2)
        current_debit_col = current_headers.get("DEBIT", 4)
        current_credit_col = current_headers.get("CREDIT", 5)
        current_last_row = rc.find_last_bank_data_row(current_bank_ws, current_header_row)
        for row in range(current_header_row + 1, current_last_row + 1):
            reference = rc.clean_string(current_bank_ws.cell(row, 6).value).upper()
            if not reference:
                continue
            current_bank_refs.add(reference)
            current_bank_transactions[reference] = {
                "month": current_month_name,
                "row": row,
                "date": rc.parse_date(current_bank_ws.cell(row, current_date_col).value),
                "narration": rc.clean_string(current_bank_ws.cell(row, current_narration_col).value).upper(),
                "ref": reference,
                "debit": rc.numeric_amount(current_bank_ws.cell(row, current_debit_col).value),
                "credit": rc.numeric_amount(current_bank_ws.cell(row, current_credit_col).value),
            }
    cross_month_count = 0
    split_count = 0

    for r in range(2, expense_ws.max_row + 1):
        vendor = rc.clean_string(expense_ws.cell(r, rc.VENDOR_COL_NAME).value)
        if not vendor:
            break
        status = rc.clean_string(expense_ws.cell(r, 18).value).upper()
        bill_no = rc.clean_string(expense_ws.cell(r, rc.VENDOR_COL_BILL_NO).value)
        expense_references = reference_parts(expense_ws.cell(r, 16).value)
        comments = rc.clean_string(expense_ws.cell(r, rc.VENDOR_COL_COMMENTS).value)
        amount = expense_payment_amount(expense_ws, r)
        if amount <= 0:
            continue

        # Explicit expense references take priority, including multi-payment rows.
        # Do not infer a cross-month match from amount or narration alone.
        current_matches = []
        missing_references = []
        for reference in expense_references:
            if reference in current_bank_transactions:
                current_matches.append(current_bank_transactions[reference])
                continue
            missing_references.append(reference)
        if expense_references and current_matches and not missing_references:
            continue

        # First account for references in this month's bank sheet. Missing
        # references are then searched in sibling months and added to the same
        # expense total, so split payments show the local and cross-month parts.
        cross_month_matches = []
        for reference in missing_references:
            reference_hits = [
                tx for tx in bank_index
                if tx["ref"] == reference or reference in tx["narration"]
            ]
            if len(reference_hits) == 1:
                cross_month_matches.append(reference_hits[0])

        reference_matches = current_matches + cross_month_matches
        if expense_references and reference_matches:
            found_total = sum(tx["debit"] for tx in reference_matches)
            difference = round(found_total - amount, 2)
            fully_matched = (
                len(reference_matches) == len(expense_references)
                and abs(difference) <= 1
            )
            status = "MATCHED" if fully_matched else "PARTIALLY MATCHED"
            expense_ws.cell(r, rc.EXPENSE_COL_BANK_PAID).value = found_total
            expense_ws.cell(r, 18).value = (
                f"{status} (Difference: {difference:,.2f} | "
                f"Current: {len(current_matches)} | Cross-month: {len(cross_month_matches)})"
            )
            for c in range(1, 19):
                expense_ws.cell(r, c).fill = copy(
                    rc.GREEN_FILL if fully_matched else rc.YELLOW_FILL
                )
            for tx in reference_matches:
                vendor = rc.clean_string(expense_ws.cell(r, rc.VENDOR_COL_NAME).value)
                note = (
                    f"Cross-month match | Detail sheet: Expense | Detail month: {current_month_name} | "
                    f"Bank sheet: {tx['month']} | Bank entry: {tx['row']} | "
                    f"Vendor: {vendor} | Bill: {bill_no or 'n/a'}"
                )
                if tx["month"] == current_month_name and "Bank_Statement" in wb.sheetnames:
                    current_bank_ws.cell(tx["row"], 8).value = note
                    current_bank_ws.cell(tx["row"], 9).value = status
                    for column in range(1, 10):
                        current_bank_ws.cell(tx["row"], column).fill = copy(
                            rc.GREEN_FILL if fully_matched else rc.YELLOW_FILL
                        )
                elif root is not None:
                    annotate_cross_month_draft(
                        root,
                        tx["month"],
                        tx["ref"],
                        note + f" | Result: {status} | Difference: {difference:,.2f}",
                        f"{status} (Detail: Expense/{current_month_name} | Bank: {tx['month']}/{tx['row']})",
                    )
            if fully_matched:
                cross_month_count += 1
            continue

        cheque_numbers = _extract_cheque_numbers(comments) or _extract_cheque_numbers(bill_no)

        if len(cheque_numbers) >= 2:
            found_parts = []
            found_total = 0.0
            for chq in cheque_numbers:
                hit = next((tx for tx in bank_index if tx["ref"] == chq or chq in tx["narration"]), None)
                if hit:
                    hit_amount = hit["debit"] if hit["debit"] > 0 else hit["credit"]
                    found_total += hit_amount
                    found_parts.append(f"{hit['month']}:Rs{hit_amount:,.0f}")

            if found_parts:
                if rc.money_equal(found_total, amount) and len(found_parts) == len(cheque_numbers):
                    expense_ws.cell(r, 18).value = "MATCHED (SPLIT PAYMENT - FULLY FOUND)"
                    expense_ws.cell(r, rc.EXPENSE_COL_BANK_PAID).value = found_total
                else:
                    expense_ws.cell(r, 18).value = (
                        f"REVIEW (SPLIT PAYMENT - {len(found_parts)}/{len(cheque_numbers)} cheques found)"
                    )
                expense_ws.cell(r, 16).value = "; ".join(found_parts)
                if root is not None:
                    vendor = rc.clean_string(expense_ws.cell(r, rc.VENDOR_COL_NAME).value)
                    split_note = (
                        f"Cross-month match | Detail sheet: Expense | Detail month: {current_month_name} | "
                        f"Vendor: {vendor} | Bill: {bill_no or 'n/a'}"
                    )
                    for cheque_number in cheque_numbers:
                        hit = next(
                            (tx for tx in bank_index
                             if tx["ref"] == cheque_number or cheque_number in tx["narration"]),
                            None,
                        )
                        if hit:
                            note = f"{split_note} | Bank sheet: {hit['month']} | Bank entry: {hit['row']}"
                            annotate_cross_month_draft(
                                root,
                                hit["month"],
                                hit["ref"],
                                note,
                                f"MATCHED (Detail: Expense/{current_month_name} | Bank: {hit['month']}/{hit['row']})"
                                if len(found_parts) == len(cheque_numbers)
                                else f"REVIEW (Detail: Expense/{current_month_name} | Bank: {hit['month']}/{hit['row']})",
                            )
                split_count += 1
            continue

        if not expense_references:
            continue
        reference_matches = [
            tx for tx in bank_index
            if any(reference == tx["ref"] or reference in tx["narration"] for reference in expense_references)
            and rc.money_equal(tx["debit"], amount)
        ]
        chosen = reference_matches

        if len(chosen) == 1:
            tx = chosen[0]
            expense_ws.cell(r, 16).value = tx["ref"]
            expense_ws.cell(r, 17).value = tx["date"]
            expense_ws.cell(r, rc.EXPENSE_COL_BANK_PAID).value = tx["debit"]
            expense_ws.cell(r, 18).value = (
                f"MATCHED (Cross-month | Bank sheet: {tx['month']} | Entry: {tx['row']})"
            )
            for c in range(1, 19):
                expense_ws.cell(r, c).fill = copy(rc.GREEN_FILL)
            cross_month_count += 1
            if root is not None:
                vendor = rc.clean_string(expense_ws.cell(r, rc.VENDOR_COL_NAME).value)
                note = (
                    f"Cross-month match | Detail sheet: Expense | Detail month: {current_month_name} | "
                    f"Vendor: {vendor} | Bill: {bill_no or 'n/a'} | "
                    f"Bank sheet: {tx['month']} | Bank entry: {tx['row']}"
                )
                annotate_cross_month_draft(
                    root,
                    tx["month"],
                    tx["ref"],
                    note,
                    f"MATCHED (Detail: Expense/{current_month_name} | Bank: {tx['month']}/{tx['row']})",
                )
        elif len(chosen) > 1:
            months = ", ".join(sorted({tx["month"] for tx in chosen}))
            expense_ws.cell(r, 18).value = f"REVIEW (possible match found in: {months})"

    return {"cross_month": cross_month_count, "split": split_count}


def enrich_cross_month_income(
    wb, bank_index: list[dict], current_month_name: str, root: Optional[Path] = None
) -> int:
    """Match NBH and other-income rows across drafts by reference ID only."""
    if not bank_index:
        return 0

    matched = 0
    income_specs = (
        ("Income_recorded_in_nbh", ("BANK REFERENCE ID", "SETTLEMENT ID", "TXN ID / CHEQUE NO."), ("CREDIT", "DEBIT"), "BANK MATCH STATUS", "BANK TRANSACTION DATE"),
        ("Income from Other Sources", ("BANK REFERENCE ID",), ("BANK DEPOSIT AMOUNT",), "MATCH STATUS", "BANK TRANSACTION DATE"),
    )
    for sheet_name, reference_headers, amount_headers, status_header, date_header in income_specs:
        if sheet_name not in wb.sheetnames:
            continue
        ws = wb[sheet_name]
        headers = {
            rc.clean_string(ws.cell(1, col).value).upper(): col
            for col in range(1, ws.max_column + 1)
        }
        reference_cols = [headers[name] for name in reference_headers if headers.get(name)]
        amount_cols = [headers[name] for name in amount_headers if headers.get(name)]
        status_col = headers.get(status_header)
        date_col = headers.get(date_header)
        if not reference_cols or not status_col:
            continue
        for row in range(2, ws.max_row + 1):
            status = rc.clean_string(ws.cell(row, status_col).value).upper()
            if status.startswith("MATCHED"):
                continue
            references = []
            for reference_col in reference_cols:
                references.extend(reference_parts(ws.cell(row, reference_col).value))
            references = list(dict.fromkeys(references))
            if not references:
                continue
            amount = next(
                (rc.numeric_amount(ws.cell(row, amount_col).value) for amount_col in amount_cols
                 if rc.numeric_amount(ws.cell(row, amount_col).value) > 0),
                0.0,
            )
            candidates = [
                tx for tx in bank_index
                if any(reference == tx["ref"] or reference in tx["narration"] for reference in references)
                and (amount <= 0 or rc.money_equal(tx["credit"], amount))
            ]
            if len(candidates) != 1:
                continue

            tx = candidates[0]
            note = (
                f"Cross-month match | Detail sheet: {sheet_name} | Detail month: {current_month_name} | "
                f"Bank sheet: {tx['month']} | Bank entry: {tx['row']} | Reference: {references[0]}"
            )
            ws.cell(row, status_col).value = (
                f"MATCHED (Cross-month | Bank sheet: {tx['month']} | Entry: {tx['row']})"
            )
            bank_reference_col = headers.get("BANK REFERENCE ID")
            if bank_reference_col and not ws.cell(row, bank_reference_col).value:
                ws.cell(row, bank_reference_col).value = tx["ref"]
            if not ws.cell(row, date_col).value:
                ws.cell(row, date_col).value = tx["date"]
            for col in range(1, ws.max_column + 1):
                ws.cell(row, col).fill = copy(rc.GREEN_FILL)
            if root is not None:
                annotate_cross_month_draft(
                    root,
                    tx["month"],
                    tx["ref"],
                    note,
                    f"MATCHED (Detail: {sheet_name}/{current_month_name} | Bank: {tx['month']}/{tx['row']})",
                )
            matched += 1
    return matched


def reapply_sibling_cross_month_links(
    wb, sibling_dirs: list[Path], current_month_name: str
) -> int:
    """Restore reverse pointers when a current bank row belongs to a sibling draft."""
    if "Bank_Statement" not in wb.sheetnames:
        return 0

    bank_ws = wb["Bank_Statement"]
    bank_header, _ = rc.locate_bank_header(bank_ws)
    bank_last = rc.find_last_bank_data_row(bank_ws, bank_header)
    bank_rows = {
        rc.clean_string(bank_ws.cell(row, 6).value).upper(): row
        for row in range(bank_header + 1, bank_last + 1)
        if rc.clean_string(bank_ws.cell(row, 6).value)
    }
    updated = 0

    for month_dir in sibling_dirs:
        draft_path = resolve_draft_path(month_dir)
        if not draft_path.exists():
            continue
        try:
            sibling_wb = load_workbook(draft_path, data_only=False)
        except Exception:
            continue
        try:
            specs = (
                ("Expense", 16, 18, "Expense"),
                ("Income_recorded_in_nbh", None, None, "Income_recorded_in_nbh"),
                ("Income from Other Sources", None, None, "Income from Other Sources"),
            )
            for sheet_name, fixed_ref_col, fixed_status_col, detail_name in specs:
                if sheet_name not in sibling_wb.sheetnames:
                    continue
                ws = sibling_wb[sheet_name]
                headers = _header_map(ws)
                ref_cols = [fixed_ref_col] if fixed_ref_col else [
                    headers[name] for name in (
                        "BANK REFERENCE ID", "SETTLEMENT ID", "TXN ID / CHEQUE NO."
                    ) if headers.get(name)
                ]
                status_col = fixed_status_col or headers.get(
                    "BANK MATCH STATUS" if sheet_name == "Income_recorded_in_nbh" else "MATCH STATUS"
                )
                if not ref_cols or not status_col:
                    continue
                for detail_row in range(2, ws.max_row + 1):
                    status = rc.clean_string(ws.cell(detail_row, status_col).value).upper()
                    if not status.startswith("MATCHED") or "CROSS" not in status:
                        continue
                    references = []
                    for ref_col in ref_cols:
                        references.extend(reference_parts(ws.cell(detail_row, ref_col).value))
                    for reference in dict.fromkeys(references):
                        bank_row = bank_rows.get(reference)
                        if bank_row is None:
                            continue
                        note = (
                            f"Cross-month match | Detail sheet: {detail_name} | "
                            f"Detail month: {month_dir.name} | Bank sheet: {current_month_name} | "
                            f"Bank entry: {bank_row} | Reference: {reference}"
                        )
                        existing = rc.clean_string(bank_ws.cell(bank_row, 8).value)
                        if note not in existing:
                            bank_ws.cell(bank_row, 8).value = f"{existing} | {note}".strip(" |") if existing else note
                        if rc.clean_string(bank_ws.cell(bank_row, 9).value).upper() != "FD TRANSFER":
                            bank_ws.cell(bank_row, 9).value = (
                                f"MATCHED (Detail: {detail_name}/{month_dir.name} | "
                                f"Bank: {current_month_name}/{bank_row})"
                            )
                            for column in range(1, 10):
                                bank_ws.cell(bank_row, column).fill = copy(rc.GREEN_FILL)
                        updated += 1
        finally:
            sibling_wb.close()
    return updated


def remove_quarterly_other_income_duplicates(
    wb, sibling_dirs: list[Path]
) -> int:
    """Remove GST-free Other Sources rows represented by quarterly NBH income."""
    if "Income from Other Sources" not in wb.sheetnames:
        return 0

    quarterly_bank_references: set[str] = set()
    workbooks = [(wb, None)]
    for month_dir in sibling_dirs:
        draft_path = resolve_draft_path(month_dir)
        if draft_path.exists():
            try:
                workbooks.append((load_workbook(draft_path, data_only=False), month_dir))
            except Exception:
                continue

    for source_wb, _ in workbooks:
        if "Income_recorded_in_nbh" not in source_wb.sheetnames:
            continue
        ws = source_wb["Income_recorded_in_nbh"]
        headers = _header_map(ws)
        bank_ref_col = headers.get("BANK REFERENCE ID")
        if not bank_ref_col:
            continue
        for row in range(2, ws.max_row + 1):
            bank_reference = rc.clean_string(ws.cell(row, bank_ref_col).value).upper()
            has_quarterly_reference = any(
                QUARTERLY_REFERENCE_PATTERN.fullmatch(rc.clean_string(ws.cell(row, column).value))
                for column in range(1, ws.max_column + 1)
            )
            if bank_reference and has_quarterly_reference:
                quarterly_bank_references.add(bank_reference)

    other_ws = wb["Income from Other Sources"]
    removed = 0
    for row in range(other_ws.max_row, 1, -1):
        bank_reference = rc.clean_string(other_ws.cell(row, 14).value).upper()
        if bank_reference not in quarterly_bank_references:
            continue
        has_gst_component = any(
            value not in (None, "")
            and not (isinstance(value, str) and value.startswith("="))
            and rc.numeric_amount(value) > 0
            for column in (9, 10, 11)
            for value in (other_ws.cell(row, column).value,)
        )
        if has_gst_component:
            continue
        other_ws.delete_rows(row, 1)
        removed += 1

    for source_wb, _ in workbooks[1:]:
        source_wb.close()
    return removed


# =============================================================
# MAIN ORCHESTRATION
# =============================================================

def resolve_draft_path(month_dir: Path) -> Path:
    year, month = parse_month_dir_name(month_dir.name)
    month_label = datetime(year, month, 1).strftime("%B_%Y")
    return month_dir / f"report_{month_label}.xlsx"


def finalize_workbook_layout(wb) -> None:
    for ws in wb.worksheets:
        rc.autofit(ws)

    sheet_order = [
        "Bank_Statement",
        "FD Transactions",
        "Income from Other Sources",
        "Expense",
        "Income_recorded_in_nbh",
    ]
    ordered_names = [name for name in sheet_order if name in wb.sheetnames]
    ordered_names.extend(
        name for name in wb.sheetnames
        if name not in ordered_names and name.startswith("Cash Withdrawal")
    )
    ordered_names.extend(name for name in wb.sheetnames if name not in ordered_names)
    wb._sheets = [wb[name] for name in ordered_names]


def normalize_tax_and_duties_rows(expense_ws) -> None:
    """Keep generated tax rows limited to their identifying fields."""
    retained_columns = {2, 5, 12, 16, 17, 18, 19}
    for row in range(2, expense_ws.max_row + 1):
        if rc.clean_string(expense_ws.cell(row, 19).value).upper() != "TAX AND DUTIES":
            continue
        for column in range(1, 20):
            if column not in retained_columns:
                expense_ws.cell(row, column).value = None


def add_missing_tax_expenses(expense_ws, bank_ws, bank_meta: dict) -> int:
    """Add statutory tax payments to an existing draft's Expense sheet."""
    summary_row = next(
        (row for row in range(1, expense_ws.max_row + 1)
         if rc.clean_string(expense_ws.cell(row, 1).value).upper() == "BILLING CATEGORY"),
        expense_ws.max_row + 1,
    )
    existing_refs = {
        rc.clean_string(expense_ws.cell(row, 16).value).upper()
        for row in range(2, summary_row)
        if rc.clean_string(expense_ws.cell(row, 16).value)
    }
    additions = []
    header_row = bank_meta["header_row"]
    last_row = rc.find_last_bank_data_row(bank_ws, header_row)
    for bank_row in range(header_row + 1, last_row + 1):
        narration = rc.clean_string(bank_ws.cell(bank_row, bank_meta["narration_col"]).value)
        original_reference = rc.clean_string(bank_ws.cell(bank_row, bank_meta["ref_col"]).value)
        generated_reference = rc.clean_string(bank_ws.cell(bank_row, 6).value)
        reference_text = f"{narration} {original_reference} {generated_reference}".upper()
        if "CBT TIN" in reference_text or "CBDT TIN" in reference_text:
            tax_name, expense_group = "TDS Paid", "Tax and Duties"
        elif "GST" in reference_text:
            tax_name, expense_group = "GST Paid", "Tax and Duties"
        else:
            continue
        debit = rc.numeric_amount(bank_ws.cell(bank_row, bank_meta["debit_col"]).value)
        if debit <= 0 or generated_reference.upper() in existing_refs:
            continue
        additions.append((bank_row, tax_name, expense_group, debit, generated_reference))

    if not additions:
        return 0
    expense_ws.insert_rows(summary_row, amount=len(additions))
    for offset, (bank_row, tax_name, expense_group, debit, reference) in enumerate(additions):
        row = summary_row + offset
        expense_ws.cell(row, 1).value = row - 1
        expense_ws.cell(row, 2).value = bank_ws.cell(bank_row, bank_meta["date_col"]).value
        expense_ws.cell(row, 5).value = tax_name
        expense_ws.cell(row, 12).value = debit
        expense_ws.cell(row, 16).value = reference
        expense_ws.cell(row, 17).value = bank_ws.cell(bank_row, bank_meta["date_col"]).value
        expense_ws.cell(row, 18).value = "MATCHED"
        expense_ws.cell(row, 19).value = expense_group
        expense_ws.cell(row, rc.EXPENSE_COL_BANK_PAID).value = debit
        for col in range(1, rc.EXPENSE_COL_BANK_PAID + 1):
            expense_ws.cell(row, col).fill = copy(rc.GREEN_FILL)
        bank_ws.cell(bank_row, 8).value = f"{tax_name}: {reference}"
        bank_ws.cell(bank_row, 9).value = "MATCHED"
        for col in range(1, 10):
            bank_ws.cell(bank_row, col).fill = copy(rc.GREEN_FILL)
    return len(additions)


def update_expense_bank_summary(expense_ws) -> None:
    """Add bank-paid totals to an existing Expense summary section."""
    summary_row = next(
        (row for row in range(1, expense_ws.max_row + 1)
         if rc.clean_string(expense_ws.cell(row, 1).value).upper() == "BILLING CATEGORY"),
        None,
    )
    if not summary_row:
        return
    data_end = summary_row - 11
    expense_ws.cell(summary_row, 9).value = "Bank Amount Paid"
    clubhouse_row = summary_row + 1
    maintenance_row = summary_row + 2
    expense_ws.cell(clubhouse_row, 9).value = (
        f'=SUMIF($E$2:$E${data_end},"*CLUBHOUSE*",$T$2:$T${data_end})'
    )
    expense_ws.cell(maintenance_row, 9).value = (
        f'=SUM($T$2:$T${data_end})-I{clubhouse_row}'
    )
    for row in (clubhouse_row, maintenance_row):
        expense_ws.cell(row, 9).number_format = rc.AMOUNT_FORMAT
    total_row = summary_row + 3
    expense_ws.cell(total_row, 9).value = f'=SUM(I{summary_row + 1}:I{summary_row + 2})'
    expense_ws.cell(total_row, 9).number_format = rc.AMOUNT_FORMAT


def ensure_reconciliation_totals(wb, bank_ws, bank_meta: dict) -> None:
    """Restore summary formulas after draft merge/current-month processing."""
    has_bank_totals = any(
        rc.clean_string(bank_ws.cell(row, 3).value).upper() == "BANK STATEMENT TOTALS"
        for row in range(1, bank_ws.max_row + 1)
    )
    if not has_bank_totals:
        rc.add_bank_totals_summary(bank_ws, bank_meta["header_row"])

    if "FD Transactions" not in wb.sheetnames:
        return
    fd_ws = wb["FD Transactions"]
    has_fd_totals = any(
        rc.clean_string(fd_ws.cell(row, 5).value).upper() == "FD TRANSACTION TOTALS"
        for row in range(1, fd_ws.max_row + 1)
    )
    if has_fd_totals:
        return

    data_rows = [
        row for row in range(2, fd_ws.max_row + 1)
        if rc.clean_string(fd_ws.cell(row, 3).value)
    ]
    if not data_rows:
        return
    last_data_row = max(data_rows)
    summary_row = last_data_row + 10
    fd_ws.cell(summary_row, 5).value = "FD Transaction Totals"
    fd_ws.cell(summary_row + 1, 5).value = "Total Debit"
    fd_ws.cell(summary_row + 1, 6).value = f"=SUM(F2:F{last_data_row})"
    fd_ws.cell(summary_row + 2, 5).value = "Total Credit"
    fd_ws.cell(summary_row + 2, 7).value = f"=SUM(G2:G{last_data_row})"
    for row, column in ((summary_row + 1, 6), (summary_row + 2, 7)):
        fd_ws.cell(row, column).number_format = rc.AMOUNT_FORMAT
        fd_ws.cell(row, column).font = rc.Font(bold=True)
        fd_ws.cell(row, column).border = rc.DOUBLE_BOTTOM


def cleanup_draft_consistency(wb, bank_ws, bank_meta: dict) -> int:
    """Flag stale resolved rows whose references or amounts no longer agree."""
    header_row = bank_meta["header_row"]
    last_row = rc.find_last_bank_data_row(bank_ws, header_row)
    expense_references: set[str] = set()
    if "Expense" in wb.sheetnames:
        expense_ws = wb["Expense"]
        for row in range(2, expense_ws.max_row + 1):
            if not is_expense_detail_row(expense_ws, row):
                continue
            reference = rc.clean_string(expense_ws.cell(row, 16).value).upper()
            if not reference:
                continue
            expense_references.update(reference_parts(reference))

        for row in range(2, expense_ws.max_row + 1):
            if not is_expense_detail_row(expense_ws, row):
                continue
            if rc.clean_string(expense_ws.cell(row, 16).value):
                continue
            expense_ws.cell(row, 17).value = None
            expense_ws.cell(row, 18).value = "PENDING"
            expense_ws.cell(row, 19).value = None
            expense_ws.cell(row, rc.EXPENSE_COL_BANK_PAID).value = None
            for col in range(1, rc.EXPENSE_COL_BANK_PAID + 1):
                expense_ws.cell(row, col).fill = copy(rc.RED_FILL)

    bank_by_ref: dict[str, list[tuple[int, float, str]]] = {}
    for row in range(header_row + 1, last_row + 1):
        reference = rc.clean_string(bank_ws.cell(row, 6).value).upper()
        if reference:
            debit = rc.numeric_amount(bank_ws.cell(row, bank_meta["debit_col"]).value)
            credit = rc.numeric_amount(bank_ws.cell(row, bank_meta["credit_col"]).value)
            bank_by_ref.setdefault(reference, []).append((row, debit or credit, "debit" if debit > 0 else "credit"))

    for row in range(header_row + 1, last_row + 1):
        fill = bank_ws.cell(row, 1).fill
        is_yellow = fill.fill_type == "solid" and fill.fgColor.type == "rgb" and fill.fgColor.rgb in {
            rc.YELLOW_FILL.fgColor.rgb,
        }
        reference = rc.clean_string(bank_ws.cell(row, 6).value).upper()
        if is_yellow and reference not in expense_references:
            bank_ws.cell(row, 8).value = None
            bank_ws.cell(row, 9).value = "PENDING"
            for col in range(1, 10):
                bank_ws.cell(row, col).fill = copy(rc.RED_FILL)

    resolved_refs: dict[str, list[tuple[float, str]]] = {}
    flagged = 0

    def check_detail(ws, ref_cols: list[int], amount_cols: list[int], status_col: int, columns: range, direction: str) -> None:
        nonlocal flagged
        for row in range(2, ws.max_row + 1):
            status = rc.clean_string(ws.cell(row, status_col).value).upper()
            if not status.startswith("MATCHED"):
                continue
            if "FOUND IN" in status or "CROSS-MONTH" in status:
                continue
            references = list(dict.fromkeys(
                reference
                for ref_col in ref_cols
                for reference in reference_parts(ws.cell(row, ref_col).value)
            ))
            amount = next(
                (rc.numeric_amount(ws.cell(row, amount_col).value) for amount_col in amount_cols
                 if rc.numeric_amount(ws.cell(row, amount_col).value) > 0),
                0.0,
            )
            matches = [
                item
                for reference in references
                for item in bank_by_ref.get(reference, [])
                if direction == "credit_or_debit" or item[2] == direction
            ]
            if direction == "credit_or_debit" and amount > 0:
                matches = [item for item in matches if rc.money_equal(amount, item[1], tolerance=1.0)]
            matches = list({(item[0], item[1], item[2]): item for item in matches}.values())
            matched_amount = sum(item[1] for item in matches)
            valid = bool(
                references
                and matches
                and (len(matches) == len(references) or direction == "credit_or_debit")
                and (amount <= 0 or rc.money_equal(amount, matched_amount, tolerance=1.0))
            )
            if valid:
                for item in matches:
                    resolved_refs.setdefault(references[matches.index(item)], []).append((item[1], direction))
                continue
            ws.cell(row, status_col).value = "REVIEW (Source Mismatch)"
            for col in columns:
                ws.cell(row, col).fill = copy(rc.RED_FILL)
            flagged += 1

    if "Expense" in wb.sheetnames:
        check_detail(wb["Expense"], [16], [rc.EXPENSE_COL_BANK_PAID], 18, range(1, 21), "debit")
    if "Income_recorded_in_nbh" in wb.sheetnames:
        income_ws = wb["Income_recorded_in_nbh"]
        income_headers = {
            rc.clean_string(income_ws.cell(1, col).value).upper(): col
            for col in range(1, income_ws.max_column + 1)
        }
        check_detail(
            income_ws,
            [income_headers[name] for name in ("BANK REFERENCE ID", "SETTLEMENT ID", "TXN ID / CHEQUE NO.") if income_headers.get(name)],
            [income_headers[name] for name in ("CREDIT", "DEBIT") if income_headers.get(name)],
            income_headers.get("BANK MATCH STATUS", income_ws.max_column),
            range(1, income_ws.max_column + 1),
            "credit_or_debit",
        )
    if "Income from Other Sources" in wb.sheetnames:
        check_detail(wb["Income from Other Sources"], [14], [13], 16, range(1, 17), "credit")

    for sheet_name in ("FD Transactions",):
        if sheet_name not in wb.sheetnames:
            continue
        ws = wb[sheet_name]
        for row in range(2, ws.max_row + 1):
            reference = rc.clean_string(ws.cell(row, 3).value).upper()
            if reference:
                resolved_refs.setdefault(reference, []).append((rc.numeric_amount(ws.cell(row, 6).value) or rc.numeric_amount(ws.cell(row, 7).value), "debit" if rc.numeric_amount(ws.cell(row, 6).value) > 0 else "credit"))

    for sheet_name in wb.sheetnames:
        if not sheet_name.startswith("Cash Withdrawal "):
            continue
        reference = rc.clean_string(wb[sheet_name]["E1"].value).upper()
        if reference:
            resolved_refs.setdefault(reference, []).append((rc.numeric_amount(wb[sheet_name]["H1"].value), "debit"))

    for row in range(header_row + 1, last_row + 1):
        status = rc.clean_string(bank_ws.cell(row, 9).value).upper()
        if "INT TRF FRM" in rc.clean_string(bank_ws.cell(row, bank_meta["narration_col"]).value).upper():
            if rc.clean_string(bank_ws.cell(row, 8).value).upper().startswith("FD INTEREST:"):
                bank_ws.cell(row, 9).value = "MATCHED"
                for col in range(1, 10):
                    bank_ws.cell(row, col).fill = copy(rc.GREEN_FILL)
                continue
        if not status.startswith("MATCHED"):
            continue
        if "CROSS-MONTH" in status or "FOUND IN" in status:
            continue
        if "INT TRF FRM" in rc.clean_string(bank_ws.cell(row, bank_meta["narration_col"]).value).upper():
            if rc.clean_string(bank_ws.cell(row, 8).value).upper().startswith("FD INTEREST:"):
                continue
        reference = rc.clean_string(bank_ws.cell(row, 6).value).upper()
        if reference in expense_references:
            continue
        bank_amount = rc.numeric_amount(bank_ws.cell(row, bank_meta["debit_col"]).value) or rc.numeric_amount(bank_ws.cell(row, bank_meta["credit_col"]).value)
        direction = "debit" if rc.numeric_amount(bank_ws.cell(row, bank_meta["debit_col"]).value) > 0 else "credit"
        candidates = [item for item in resolved_refs.get(reference, []) if item[1] == direction]
        if not candidates or not any(amount <= 0 or rc.money_equal(amount, bank_amount, tolerance=1.0) for amount, _ in candidates):
            bank_ws.cell(row, 9).value = "REVIEW (Source Mismatch)"
            for col in range(1, 10):
                bank_ws.cell(row, col).fill = copy(rc.RED_FILL)
            flagged += 1
    return flagged


def run_for_month(month_dir: Path, vendor_bills: Optional[Path], data_root: Optional[Path]) -> None:
    month_dir = month_dir.resolve()
    if not month_dir.is_dir():
        raise FileNotFoundError(f"Month directory not found: {month_dir}")

    year, month = parse_month_dir_name(month_dir.name)
    month_label = datetime(year, month, 1).strftime("%B %Y")

    print("=" * 60)
    print("AUTOMATED MONTH-WISE RECONCILIATION")
    print("=" * 60)
    print(f"Month directory       : {month_dir}")
    print(f"Detected period       : {month_label}")

    root = data_root.resolve() if data_root else month_dir.parent
    draft_path = resolve_draft_path(month_dir)
    existing_draft_path = draft_path if draft_path.exists() else None
    print(f"Draft output          : {draft_path}")
    print(f"Existing draft found  : {'yes - merging manual values' if existing_draft_path else 'no - creating new draft'}")
    print()

    fresh_path = draft_path.with_name(draft_path.stem + ".fresh.xlsx")
    files = discover_input_files(month_dir / "input")
    print(f"Bank statement        : {files['bank'].name}")
    print(f"NBH bank book (income): {files['income'].name}")
    print(f"NBH GST report (other): {files['other_income'].name}")
    vendor_path = find_vendor_bills_file(month_dir, root, vendor_bills)
    print(f"Vendor bills workbook : {vendor_path}")
    rc.run_reconciliation(
        bank_path=files["bank"],
        vendor_path=vendor_path,
        income_path=files["income"],
        other_income_path=files["other_income"],
        output_path=fresh_path,
    )
    fresh_wb = load_workbook(fresh_path)

    if existing_draft_path:
        print()
        print("-" * 60)
        print("Merging values from existing draft")
        print("-" * 60)
        old_wb = load_workbook(existing_draft_path)
        merge_draft(old_wb, fresh_wb)
        old_wb.close()

    bank_ws = fresh_wb["Bank_Statement"]
    bank_meta = _build_bank_meta(bank_ws)

    consistency_flags = cleanup_draft_consistency(fresh_wb, bank_ws, bank_meta)
    print(f"Draft consistency mismatches flagged: {consistency_flags}")

    tax_expenses_added = 0
    if "Expense" in fresh_wb.sheetnames:
        tax_expenses_added = add_missing_tax_expenses(fresh_wb["Expense"], bank_ws, bank_meta)
    if tax_expenses_added:
        print(f"Statutory tax expenses added: {tax_expenses_added}")

    print()
    print("-" * 60)
    print("Linking Income from Other Sources <-> Bank Statement")
    print("-" * 60)
    other_income_links = 0
    if "Income from Other Sources" in fresh_wb.sheetnames:
        migrated = migrate_other_income_schema(fresh_wb["Income from Other Sources"])
        if migrated:
            print("Income from Other Sources: migrated legacy columns to current format")
        refreshed = refresh_other_income_categories(fresh_wb["Income from Other Sources"])
        if refreshed:
            print("Income from Other Sources: refreshed category options and GST summary")
        other_income_links = link_other_income_with_bank(fresh_wb["Income from Other Sources"], bank_ws, bank_meta)
    print(f"Other Income entries linked to Bank Statement: {other_income_links}")

    expense_reference_links = 0
    if "Expense" in fresh_wb.sheetnames:
        expense_reference_links = link_expense_references_with_bank(
            fresh_wb["Expense"], bank_ws, bank_meta
        )
    print(f"Expense references linked to Bank Statement: {expense_reference_links}")

    recorded_income_links = 0
    if "Income_recorded_in_nbh" in fresh_wb.sheetnames:
        recorded_income_links = link_recorded_income_with_bank(
            fresh_wb["Income_recorded_in_nbh"], bank_ws, bank_meta
        )
    print(f"NBH income entries linked by Bank Reference ID: {recorded_income_links}")

    cash_withdrawal_links = link_cash_withdrawals_with_bank(fresh_wb, bank_ws, bank_meta)
    print(f"Cash withdrawal registers linked to Bank Statement: {cash_withdrawal_links}")

    colored_bank = reapply_match_highlighting(bank_ws, bank_meta["header_row"])

    expense_colored = 0
    if "Expense" in fresh_wb.sheetnames:
        expense_ws = fresh_wb["Expense"]
        normalize_tax_and_duties_rows(expense_ws)
        update_expense_bank_summary(expense_ws)
        expense_rows = [
            r for r in range(2, expense_ws.max_row + 1)
            if rc.clean_string(expense_ws.cell(r, rc.VENDOR_COL_NAME).value)
        ]
        expense_colored = _color_rows_by_status(expense_ws, 18, expense_rows, range(1, 21))

    income_colored = 0
    if "Income_recorded_in_nbh" in fresh_wb.sheetnames:
        income_ws = fresh_wb["Income_recorded_in_nbh"]
        income_rows = list(range(2, income_ws.max_row + 1))
        income_colored = _color_rows_by_status(income_ws, income_ws.max_column, income_rows, range(1, income_ws.max_column + 1))

    other_income_colored = 0
    if "Income from Other Sources" in fresh_wb.sheetnames:
        oi_ws = fresh_wb["Income from Other Sources"]
        oi_rows = [
            r for r in range(2, oi_ws.max_row + 1)
            if rc.numeric_amount(oi_ws.cell(r, 13).value) > 0 or rc.clean_string(oi_ws.cell(r, 14).value)
        ]
        other_income_colored = _color_rows_by_status(oi_ws, 16, oi_rows, range(1, 17))

    print(f"Green-highlighted        : Bank {colored_bank}, Expense {expense_colored}, "
          f"Income {income_colored}, Other Income {other_income_colored}")

    sibling_dirs = discover_sibling_month_dirs(root, exclude=month_dir)
    if sibling_dirs:
        print()
        print("-" * 60)
        print(f"Cross-month search across {len(sibling_dirs)} other month folder(s)")
        print("-" * 60)
        bank_index = build_cross_month_bank_index(sibling_dirs)
        enrichment = enrich_unmatched_expenses(fresh_wb, bank_index, month_dir.name, root=root)
        print(f"Cross-month expense matches found : {enrichment['cross_month']}")
        print(f"Split-cheque payments resolved     : {enrichment['split']}")
        cross_income = enrich_cross_month_income(fresh_wb, bank_index, month_dir.name, root=root)
        print(f"Cross-month income matches found   : {cross_income}")
        reverse_links = reapply_sibling_cross_month_links(
            fresh_wb, sibling_dirs, month_dir.name
        )
        print(f"Reverse cross-month bank links restored: {reverse_links}")

    finalize_workbook_layout(fresh_wb)
    fresh_wb.active = fresh_wb.sheetnames.index("Bank_Statement")
    fresh_wb.save(fresh_path)
    fresh_wb.close()
    os.replace(fresh_path, draft_path)

    print()
    print("=" * 60)
    print(f"Draft saved: {draft_path}")
    print("=" * 60)


def discover_month_dirs(root: Path) -> list[Path]:
    """Return immediate month directories in chronological/name order."""
    if not root.is_dir():
        raise FileNotFoundError(f"Data root not found: {root}")
    month_dirs = []
    for child in sorted(root.iterdir()):
        if not child.is_dir():
            continue
        try:
            parse_month_dir_name(child.name)
        except ValueError:
            continue
        if (child / "input").is_dir() or resolve_draft_path(child).exists():
            month_dirs.append(child)
    if not month_dirs:
        raise FileNotFoundError(f"No month directories found under {root}")
    return sorted(month_dirs, key=lambda path: parse_month_dir_name(path.name))


def run_for_root(data_root: Path, vendor_bills: Optional[Path]) -> None:
    """Run reconciliation in five repository-wide phases."""
    month_dirs = discover_month_dirs(data_root.resolve())
    print(f"Found {len(month_dirs)} month director{'y' if len(month_dirs) == 1 else 'ies'} under {data_root.resolve()}")

    # Phase 1 and 2: every month gets a fresh workbook, then every existing
    # draft is merged before any current-month or cross-month matching starts.
    for month_dir in month_dirs:
        _prepare_root_month(month_dir, vendor_bills, data_root)

    # Phase 3: reconcile each month only against its own bank sheet.
    for month_dir in month_dirs:
        _run_current_month_pass(month_dir)

    # Phase 4 and 5: reload each completed current-month draft, apply
    # cross-month links, and save it. Sibling annotations are already on disk
    # when the next month's draft is loaded, so both sides retain the link.
    for month_dir in month_dirs:
        _run_cross_month_pass(month_dir, month_dirs, data_root)


def _prepare_root_month(month_dir: Path, vendor_bills: Optional[Path], data_root: Path) -> None:
    month_dir = month_dir.resolve()
    root = data_root.resolve()
    draft_path = resolve_draft_path(month_dir)
    existing_draft_path = draft_path if draft_path.exists() else None
    fresh_path = draft_path.with_name(draft_path.stem + ".fresh.xlsx")
    files = discover_input_files(month_dir / "input")
    vendor_path = find_vendor_bills_file(month_dir, root, vendor_bills)
    rc.run_reconciliation(
        bank_path=files["bank"], vendor_path=vendor_path,
        income_path=files["income"], other_income_path=files["other_income"],
        output_path=fresh_path,
    )
    fresh_wb = load_workbook(fresh_path)
    if existing_draft_path:
        old_wb = load_workbook(existing_draft_path)
        merge_draft(old_wb, fresh_wb)
        old_wb.close()
    fresh_wb.save(draft_path)
    fresh_wb.close()
    if fresh_path.exists():
        fresh_path.unlink()


def _run_current_month_pass(month_dir: Path) -> None:
    draft_path = resolve_draft_path(month_dir)
    wb = load_workbook(draft_path)
    bank_ws = wb["Bank_Statement"]
    bank_meta = _build_bank_meta(bank_ws)
    cleanup_draft_consistency(wb, bank_ws, bank_meta)
    if "Expense" in wb.sheetnames:
        add_missing_tax_expenses(wb["Expense"], bank_ws, bank_meta)
    if "Income from Other Sources" in wb.sheetnames:
        migrate_other_income_schema(wb["Income from Other Sources"])
        refresh_other_income_categories(wb["Income from Other Sources"])
        link_other_income_with_bank(wb["Income from Other Sources"], bank_ws, bank_meta)
    if "Expense" in wb.sheetnames:
        link_expense_references_with_bank(wb["Expense"], bank_ws, bank_meta)
    if "Income_recorded_in_nbh" in wb.sheetnames:
        link_recorded_income_with_bank(wb["Income_recorded_in_nbh"], bank_ws, bank_meta)
    link_cash_withdrawals_with_bank(wb, bank_ws, bank_meta)
    reapply_match_highlighting(bank_ws, bank_meta["header_row"])
    if "Expense" in wb.sheetnames:
        normalize_tax_and_duties_rows(wb["Expense"])
        update_expense_bank_summary(wb["Expense"])
        rows = [r for r in range(2, wb["Expense"].max_row + 1)
                if rc.clean_string(wb["Expense"].cell(r, rc.VENDOR_COL_NAME).value)]
        _color_rows_by_status(wb["Expense"], 18, rows, range(1, 21))
    if "Income_recorded_in_nbh" in wb.sheetnames:
        ws = wb["Income_recorded_in_nbh"]
        _color_rows_by_status(ws, ws.max_column, range(2, ws.max_row + 1), range(1, ws.max_column + 1))
    if "Income from Other Sources" in wb.sheetnames:
        ws = wb["Income from Other Sources"]
        rows = [r for r in range(2, ws.max_row + 1)
                if rc.numeric_amount(ws.cell(r, 13).value) > 0 or rc.clean_string(ws.cell(r, 14).value)]
        _color_rows_by_status(ws, 16, rows, range(1, 17))
    ensure_reconciliation_totals(wb, bank_ws, bank_meta)
    finalize_workbook_layout(wb)
    wb.active = wb.sheetnames.index("Bank_Statement")
    wb.save(draft_path)
    wb.close()


def _run_cross_month_pass(month_dir: Path, month_dirs: list[Path], data_root: Path) -> None:
    draft_path = resolve_draft_path(month_dir)
    wb = load_workbook(draft_path)
    siblings = [path for path in month_dirs if path.resolve() != month_dir.resolve()]
    bank_index = build_cross_month_bank_index(siblings)
    enrichment = enrich_unmatched_expenses(wb, bank_index, month_dir.name, root=data_root)
    enrich_cross_month_income(wb, bank_index, month_dir.name, root=data_root)
    reapply_sibling_cross_month_links(wb, siblings, month_dir.name)
    quarterly_duplicates = remove_quarterly_other_income_duplicates(wb, siblings)
    if quarterly_duplicates:
        print(f"{month_dir.name}: removed {quarterly_duplicates} quarterly Other Sources duplicate(s)")
    finalize_workbook_layout(wb)
    wb.active = wb.sheetnames.index("Bank_Statement")
    wb.save(draft_path)
    wb.close()
    print(f"{month_dir.name}: cross-month expenses {enrichment['cross_month']}, split {enrichment['split']}")


def main():
    parser = argparse.ArgumentParser(
        description="Automated month-wise bank reconciliation orchestrator",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""Examples:
    python3 auto_reconcile.py test_data/2026-2027
    python3 auto_reconcile.py test_data/2026-2027 --vendor-bills test_data/2026-2027/VendorBills.xlsx
    python3 auto_reconcile.py test_data/2026-2027/Aug_26 --vendor-bills test_data/2026-2027/VendorBills.xlsx
        """,
    )
    parser.add_argument("data_path", nargs="?", help="Data root or a single month directory")
    parser.add_argument("--vendor-bills", default=None, help="Explicit path to the Vendor Bills workbook")
    parser.add_argument(
        "--data-root", default=None,
        help="Root directory containing all month folders (defaults to the month directory's parent)"
    )
    args = parser.parse_args()

    if not args.data_path and not args.data_root:
        parser.error("provide a data root or month directory")
    data_path = Path(args.data_path) if args.data_path else Path(args.data_root)
    vendor_bills = Path(args.vendor_bills) if args.vendor_bills else None
    data_root = Path(args.data_root) if args.data_root else None

    try:
        if data_path.is_dir() and (data_path / "input").is_dir():
            run_for_month(data_path, vendor_bills, data_root)
        else:
            run_for_root(data_path, vendor_bills)
    except (FileNotFoundError, ValueError, RuntimeError) as error:
        print(f"\nERROR: {error}")
        sys.exit(1)


if __name__ == "__main__":
    main()
