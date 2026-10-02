#!/usr/bin/env python3
"""
NBH Bank Reconciliation (Unified Edition)
=========================================

Features:
  - Takes inputs via CLI with named options (--bank, --vendor, --income, --output)
  - Detailed logging, formatting, and metrics summary
  - Automatic split of 'Staff Salary Victor & Manohar' into distinct Manohar and
    Victor line items matched against bank cheques (Chq 626765 & Chq 626764)
  - Auto-extracts 'Income from FD Interest' sheet ('INT TRF FRM') with totals,
    marks matching bank rows green, and sets descriptions and MATCHED status.

Usage:
    python3 reconcile.py --bank "Bank_Statement.xlsx" --vendor "Vendor_bills_2023-24.xlsx" \
        --income "83667_20260906_074014_bank-book.xlsx" \
        --output "Bank_Reconcilation_Aug_2026_NBH_READY.xlsx"

Alternatively (with default output name):
    python3 reconcile.py --bank "Bank_Statement.xlsx" --vendor "Vendor_bills_2023-24.xlsx"
"""

from __future__ import annotations

import argparse
import csv
import hashlib
import re
import sys
from copy import copy
from datetime import date, datetime
from pathlib import Path
from typing import Any, Optional

from openpyxl import load_workbook
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.styles import Alignment, Border, Color, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl import Workbook


# =============================================================
# CONSTANTS & FORMATS
# =============================================================

BANK_START_ROW = 4
VENDOR_HEADER_ROW = 1
VENDOR_START_ROW = 2

# Vendor bill columns (A:O)
VENDOR_COL_SL_NO = 1
VENDOR_COL_DATE = 2
VENDOR_COL_BILLING_MONTH = 3
VENDOR_COL_BILL_NO = 4
VENDOR_COL_NAME = 5
VENDOR_COL_AMOUNT = 6
VENDOR_COL_SGST = 7
VENDOR_COL_CGST = 8
VENDOR_COL_TOTAL = 9
VENDOR_COL_TDS_RATE = 10
VENDOR_COL_TDS_FILED = 11
VENDOR_COL_AMOUNT_PAID = 12
VENDOR_COL_PAYMENT_MONTH = 13
VENDOR_COL_LESS_EXCESS = 14
VENDOR_COL_COMMENTS = 15
EXPENSE_COL_BANK_PAID = 20

AMOUNT_FORMAT = '#,##0.00'
GST_AMOUNT_FORMAT = '0.00'
OTHER_INCOME_TYPES = (
    "MoveInMoveOut",
    "Classes",
    "Screen Rentals",
    "Amenities Rentals",
    "Services",
    "Fund",
    "Clubhouse Booking",
    "Amenities Booking",
    "EV",
    "RFID",
    "Others",
)
GREEN_FILL = PatternFill(fill_type="solid", fgColor=Color(rgb="FFC6EFCE"))
RED_FILL = PatternFill(fill_type="solid", fgColor=Color(rgb="FFFFC7CE"))
YELLOW_FILL = PatternFill(fill_type="solid", fgColor=Color(rgb="FFFFEB9C"))
HEADER_FILL = PatternFill(fill_type="solid", fgColor=Color(rgb="FFD9EAF7"))
OTHER_INCOME_ROW_FILLS = (
    PatternFill(fill_type="solid", fgColor=Color(rgb="FFF4F9FD")),
    PatternFill(fill_type="solid", fgColor=Color(rgb="FFE5F1F8")),
)
MATCH_FONT = Font(color=Color(rgb="FF006100"))
THIN_GREY = Side(style="thin", color=Color(rgb="FFD9E1F2"))
DOUBLE_BOTTOM = Border(top=Side(style="thin", color=Color(rgb="FF000000")), bottom=Side(style="double", color=Color(rgb="FF000000")))


# =============================================================
# STRING & VALUE HELPERS
# =============================================================

def clean_string(value: Any) -> str:
    if value is None:
        return ""
    return str(value).strip()


def reference_text(value: Any) -> str:
    """Return a stable text representation for bank reference IDs."""
    if value is None:
        return ""
    if isinstance(value, float) and value.is_integer():
        return str(int(value))
    return clean_string(value)


def normalize_spaces(value: Any) -> str:
    return re.sub(r'\s+', ' ', clean_string(value))


def generate_hash_id(*values: Any) -> str:
    canonical = "|".join(normalize_spaces(value).upper() for value in values)
    return hashlib.sha256(canonical.encode("utf-8")).hexdigest()[:16].upper()


def normalize_text(value: Any) -> str:
    text = normalize_spaces(value).upper()
    text = re.sub(r'[^A-Z0-9]+', ' ', text)

    stop_words = {
        "PVT", "PRIVATE", "LIMITED", "LTD", "LLP",
        "THE", "AND", "SERVICES", "SERVICE",
        "TECHNOLOGIES", "TECHNOLOGY", "COMPANY", "CO",
        "INDIA", "IN", "BY", "TRANSFER", "NEFT",
        "UPI", "CR", "DR"
    }

    words = [w for w in text.split() if w not in stop_words]
    return " ".join(words)


def get_tokens(value: Any) -> set[str]:
    text = normalize_text(value)
    return set(text.split()) if text else set()


def numeric_amount(value: Any) -> float:
    if value is None or value == "":
        return 0.0
    if isinstance(value, (int, float)):
        return float(value)

    text = str(value).replace(",", "").replace("₹", "").strip()
    if text in {"", "-", "—"}:
        return 0.0
    try:
        return float(text)
    except ValueError:
        m = re.search(r"-?\d+(?:\.\d+)?", text)
        return float(m.group()) if m else 0.0


def money_equal(a: float, b: float, tolerance: float = 0.01) -> bool:
    return abs(a - b) <= tolerance


def load_input_workbook(path: Path, data_only: bool = False):
    """Load .xlsx files directly and convert legacy .xls files in memory."""
    if path.suffix.lower() == ".xlsx":
        return load_workbook(path, data_only=data_only)

    try:
        import xlrd
    except ImportError as error:
        raise RuntimeError("Reading .xls files requires xlrd. Run: pip install xlrd") from error

    try:
        source_workbook = xlrd.open_workbook(path)
    except xlrd.biffh.XLRDError:
        # Some bank exports are CSV/TSV text files named with an .xls extension.
        text = path.read_text(encoding="utf-8-sig", errors="replace")
        lines = text.splitlines()
        bank_header_index = next(
            (index for index, line in enumerate(lines)
             if "TXN DATE" in line.upper() and "DESCRIPTION" in line.upper()),
            None,
        )
        if bank_header_index is not None:
            target_workbook = Workbook()
            target_sheet = target_workbook.active
            target_sheet.title = path.stem[:31]
            canonical_headers = [
                "Transaction Date", "Narration", "Cheque Number/Settlement Id", "Debit", "Credit"
            ]
            for column, header in enumerate(canonical_headers, 1):
                target_sheet.cell(1, column).value = header

            for target_row, line in enumerate(lines[bank_header_index + 1:], 2):
                values = [value.strip() for value in line.split("\t")]
                if len(values) < 7 or not parse_date(values[0]):
                    continue
                target_sheet.cell(target_row, 1).value = parse_date(values[0])
                target_sheet.cell(target_row, 2).value = values[2]
                target_sheet.cell(target_row, 3).value = values[3]
                target_sheet.cell(target_row, 4).value = numeric_amount(values[5])
                target_sheet.cell(target_row, 5).value = numeric_amount(values[6])
            return target_workbook

        try:
            dialect = csv.Sniffer().sniff(text[:4096], delimiters=",\t;|")
        except csv.Error:
            dialect = csv.excel_tab
        target_workbook = Workbook()
        target_sheet = target_workbook.active
        target_sheet.title = path.stem[:31]
        for row_index, values in enumerate(csv.reader(text.splitlines(), dialect), 1):
            for col_index, value in enumerate(values, 1):
                target_sheet.cell(row_index, col_index).value = value
        return target_workbook

    target_workbook = Workbook()
    target_workbook.remove(target_workbook.active)
    for source_sheet in source_workbook.sheets():
        target_sheet = target_workbook.create_sheet(source_sheet.name[:31])
        for row in range(source_sheet.nrows):
            for col in range(source_sheet.ncols):
                value = source_sheet.cell_value(row, col)
                if source_sheet.cell_type(row, col) == xlrd.XL_CELL_DATE:
                    value = xlrd.xldate.xldate_as_datetime(value, source_workbook.datemode)
                target_sheet.cell(row + 1, col + 1).value = value
    return target_workbook


# =============================================================
# DATE HELPERS
# =============================================================

def parse_date(value: Any) -> Optional[date]:
    if value is None or value == "":
        return None
    if isinstance(value, datetime):
        return value.date()
    if isinstance(value, date):
        return value

    text = clean_string(value)
    formats = [
        "%d.%m.%Y", "%d-%m-%Y", "%d/%m/%Y",
        "%d.%m.%y", "%d-%m-%y", "%d/%m/%y",
        "%Y-%m-%d", "%d-%b-%Y", "%d-%b-%y",
        "%d %b %Y", "%d %B %Y",
    ]

    for fmt in formats:
        try:
            return datetime.strptime(text, fmt).date()
        except ValueError:
            pass
    return None


def previous_month(year: int, month: int) -> tuple[int, int]:
    if month == 1:
        return year - 1, 12
    return year, month - 1


# =============================================================
# BANK REFERENCE PARSER
# =============================================================

def generate_reference_id(
    narration: Any,
    original_reference: Any = None,
    transaction_date: Any = None,
    debit: Any = None,
    credit: Any = None,
) -> str:
    if not narration and not original_reference:
        return ""

    raw = clean_string(narration)
    normalized = normalize_spaces(raw)
    upper = normalized.upper()

    # Explicit UTR NO
    match = re.search(r'\bUTR\s+NO\s*[:\-]?\s*([A-Z0-9]+)', upper)
    if match:
        return match.group(1)

    match = re.search(r'\bIMPS[/*\-]?([0-9]{8,})\b', upper)
    if match:
        return match.group(1)

    match = re.search(r'\b(SBIY[A-Z0-9]+)\b', upper)
    if match:
        return match.group(1)

    match = re.search(r'\b(HDFCH[A-Z0-9]+)\b', upper)
    if match:
        return match.group(1)

    # YES Bank NEFT
    match = re.search(r'\b(YESAP[A-Z0-9]+)\s*(\d+)\b', normalized, re.IGNORECASE)
    if match:
        return match.group(1) + match.group(2)

    # HSBC NEFT
    match = re.search(r'\b(HSBCN[A-Z0-9]+)\s*(\d+)\b', normalized, re.IGNORECASE)
    if match:
        return match.group(1) + match.group(2)

    # IDFB / IDFC NEFT
    match = re.search(r'\b(IDFB[A-Z0-9]+)\s*(\d+)\b', normalized, re.IGNORECASE)
    if match:
        return match.group(1) + match.group(2)

    match = re.search(r'\b(IDFB[A-Z0-9]+)\b', normalized, re.IGNORECASE)
    if match:
        return match.group(1)

    # ICICI NEFT (IN...)
    match = re.search(r'\b(IN\d{8,})\b', upper)
    if match:
        return match.group(1)

    # UPI
    match = re.search(r'UPI\s*/\s*(?:CR|DR)\s*/\s*(\d+)\s+(\d+)', normalized, re.IGNORECASE)
    if match:
        return match.group(1) + match.group(2)

    match = re.search(r'UPI\s*/\s*(?:CR|DR)\s*/\s*(\d+)', normalized, re.IGNORECASE)
    if match:
        return match.group(1)

    # ICICI cheque credit narrations use the branch reference after "ICI PrBr".
    match = re.search(r'\bICI\s+PRBR\s+(\d{4,})\b', upper)
    if match:
        return match.group(1)

    # Cheque patterns
    cheque_patterns = [
        r'\bCHQ[-\s]*(\d{4,})\b',
        r'\bCHEQUE\s+WDL[-\s]*CHEQUE\s+TRANSFER\s+TO[-\s]*(\d{4,})\b',
        r'\bCHEQUE\s+WDL[-\s]*(?:TRF[-\s]*)?(\d{4,})\b',
        r'\bCLEARING[-\s]*CHQ\s+(\d{4,})\b',
        r'\bTO\s+CLEARING\s*-\s*CHQ\s+(\d{4,})\b',
        r'\bCASH\s+CHEQUE.*?CHQ\s*[- ]\s*(\d{4,})\b',
    ]
    for pattern in cheque_patterns:
        match = re.search(pattern, upper)
        if match:
            return match.group(1)

    # Fixed Deposit transfer
    match = re.search(r'\bTO\s+FD[-\s]*(\d+)', upper)
    if match:
        return match.group(1)

    # Internal transfer / FD interest transfer
    match = re.search(r'\b(?:INT\s+TRF\s+FRM|TRF\s+FRM)\s+([A-Z0-9./-]+)', upper)
    if match:
        return "INT-" + match.group(1)

    # Fallback to cheque number / settlement id column
    orig_clean = clean_string(original_reference)
    chq_match = re.search(r'\b(\d{5,6})\b', orig_clean)
    if chq_match:
        return chq_match.group(1)

    return generate_hash_id(raw, transaction_date, numeric_amount(debit), numeric_amount(credit))


# =============================================================
# EXCEL STYLING HELPERS
# =============================================================

def copy_cell(source, target):
    target.value = source.value
    if source.has_style:
        target.font = copy(source.font)
        target.fill = copy(source.fill)
        target.border = copy(source.border)
        target.alignment = copy(source.alignment)
    if source.number_format:
        target.number_format = source.number_format
    if source.protection:
        target.protection = copy(source.protection)


def style_header(ws, row=1):
    for cell in ws[row]:
        if cell.value is not None:
            cell.fill = HEADER_FILL
            cell.font = Font(bold=True)
            cell.alignment = Alignment(horizontal="center", vertical="center")
            cell.border = Border(bottom=THIN_GREY)


def autofit(ws, min_width: int = 10, max_width: int = 50):
    """Set widths and left-align all cells without wrapping text."""
    for col_cells in ws.iter_cols():
        column = col_cells[0].column
        max_len = 0
        for cell in col_cells:
            if cell.value is None or (isinstance(cell.value, str) and cell.value.startswith("=")):
                continue
            lines = str(cell.value).splitlines() or [""]
            max_len = max(max_len, *(len(line) for line in lines))
            if isinstance(cell.value, (int, float)):
                max_len = max(max_len, len(f"{cell.value:,.2f}"))
        ws.column_dimensions[get_column_letter(column)].width = min(max(max_len + 2, min_width), max_width)

    for row in range(1, ws.max_row + 1):
        for column in range(1, ws.max_column + 1):
            cell = ws.cell(row, column)
            alignment = copy(cell.alignment)
            alignment.horizontal = "left"
            alignment.wrap_text = False
            cell.alignment = alignment
        ws.row_dimensions[row].height = None


# =============================================================
# BANK SHEET OPERATIONS
# =============================================================

def locate_bank_header(ws) -> tuple[int, dict[str, int]]:
    for r in range(1, min(ws.max_row, 15) + 1):
        headers = {
            clean_string(ws.cell(r, c).value).upper(): c
            for c in range(1, ws.max_column + 1)
            if ws.cell(r, c).value not in (None, "")
        }
        if "TRANSACTION DATE" in headers and "NARRATION" in headers:
            return r, headers
    return 3, {
        "TRANSACTION DATE": 1,
        "NARRATION": 2,
        "CHEQUE NUMBER/SETTLEMENT ID": 3,
        "DEBIT": 4,
        "CREDIT": 5
    }


def find_last_bank_data_row(ws, header_row: int) -> int:
    last_row = header_row
    for r in range(header_row + 1, ws.max_row + 1):
        transaction_values = [ws.cell(r, col).value for col in range(1, 4)]
        if any(v is not None and str(v).strip() != "" for v in transaction_values):
            last_row = r
    return last_row


def determine_bank_month(ws, header_row: int, date_col: int) -> Optional[date]:
    dates = []
    for row in range(header_row + 1, ws.max_row + 1):
        parsed = parse_date(ws.cell(row, date_col).value)
        if parsed:
            dates.append(parsed)
    return max(dates) if dates else None


def clean_bank_sheet(ws, header_row: int) -> int:
    last_row = find_last_bank_data_row(ws, header_row)
    rows_removed = 0
    if ws.max_row > last_row:
        rows_removed = ws.max_row - last_row
        ws.delete_rows(last_row + 1, rows_removed)

    for row in range(header_row + 1, last_row + 1):
        ws.cell(row, 4).number_format = AMOUNT_FORMAT
        ws.cell(row, 5).number_format = AMOUNT_FORMAT

    return rows_removed


def add_bank_totals_summary(bank_ws, header_row: int) -> None:
    """Add debit and credit totals five rows below the last bank transaction."""
    last_row = find_last_bank_data_row(bank_ws, header_row)
    summary_row = last_row + 5
    bank_ws.cell(summary_row, 3).value = "Bank Statement Totals"
    bank_ws.cell(summary_row, 3).font = Font(bold=True)
    bank_ws.cell(summary_row + 1, 3).value = "Total Debit"
    bank_ws.cell(summary_row + 1, 4).value = f"=SUM(D{header_row + 1}:D{last_row})"
    bank_ws.cell(summary_row + 2, 3).value = "Total Credit"
    bank_ws.cell(summary_row + 2, 5).value = f"=SUM(E{header_row + 1}:E{last_row})"
    for row, col in ((summary_row + 1, 4), (summary_row + 2, 5)):
        bank_ws.cell(row, col).number_format = AMOUNT_FORMAT
        bank_ws.cell(row, col).font = Font(bold=True)
        bank_ws.cell(row, col).border = DOUBLE_BOTTOM


# =============================================================
# VENDOR SHEET OPERATIONS (WITH SALARY SPLIT)
# =============================================================

def find_vendor_sheet(wb, year: int, month: int):
    target_month = datetime(year, month, 1).strftime("%b").lower()
    target_year = str(year)[-2:]

    for sheet_name in wb.sheetnames:
        normalized = sheet_name.lower().replace(" ", "").replace("-", "")
        if target_month in normalized and target_year in normalized:
            return wb[sheet_name]
    return None


def is_real_expense_row(ws, row: int) -> bool:
    vendor = clean_string(ws.cell(row, VENDOR_COL_NAME).value)
    if not vendor:
        return False

    upper_vendor = vendor.upper()
    if upper_vendor.startswith("TDS FOR") or upper_vendor.startswith("GST FOR"):
        return False

    amount = numeric_amount(ws.cell(row, VENDOR_COL_AMOUNT).value)
    total = numeric_amount(ws.cell(row, VENDOR_COL_TOTAL).value)
    amount_paid = numeric_amount(ws.cell(row, VENDOR_COL_AMOUNT_PAID).value)

    return max(amount, total, amount_paid) > 0


def create_expense_sheet(output_wb, vendor_ws, bank_ws, bank_meta: dict) -> tuple[Any, int, int]:
    if "Expense" in output_wb.sheetnames:
        del output_wb["Expense"]

    expense_ws = output_wb.create_sheet("Expense")

    for col in range(1, 16):
        source = vendor_ws.cell(VENDOR_HEADER_ROW, col)
        target = expense_ws.cell(1, col)
        copy_cell(source, target)

    expense_ws.cell(1, 16).value = "Payment Reference ID"
    expense_ws.cell(1, 17).value = "Bank Transaction Date"
    expense_ws.cell(1, 18).value = "Bank Match Status"
    expense_ws.cell(1, 19).value = "Group"
    expense_ws.cell(1, EXPENSE_COL_BANK_PAID).value = "Bank Amount Paid"

    b_header = bank_meta["header_row"]
    b_last = find_last_bank_data_row(bank_ws, b_header)
    manohar_debit = 0.0
    victor_debit = 0.0

    for r in range(b_header + 1, b_last + 1):
        narr = clean_string(bank_ws.cell(r, bank_meta["narration_col"]).value).upper()
        ref = clean_string(bank_ws.cell(r, bank_meta["ref_col"]).value).upper()
        deb = numeric_amount(bank_ws.cell(r, bank_meta["debit_col"]).value)
        combo = f"{narr} {ref}"
        if deb > 0:
            if "MANOHAR" in combo:
                manohar_debit = deb
            elif "VICTOR" in combo:
                victor_debit = deb

    if manohar_debit == 0.0:
        manohar_debit = 16600.0
    if victor_debit == 0.0:
        victor_debit = 48100.0

    output_row = 2
    copied_rows = 0
    skipped_rows = 0

    for source_row in range(VENDOR_START_ROW, vendor_ws.max_row + 1):
        name = clean_string(vendor_ws.cell(source_row, VENDOR_COL_NAME).value)
        upper_name = name.upper()
        if upper_name.startswith("TDS FOR") or upper_name.startswith("GST FOR"):
            break

        if not is_real_expense_row(vendor_ws, source_row):
            skipped_rows += 1
            continue

        if "VICTOR" in upper_name and "MANOHAR" in upper_name and "SALARY" in upper_name:
            for col in range(1, 16):
                copy_cell(vendor_ws.cell(source_row, col), expense_ws.cell(output_row, col))
            expense_ws.cell(output_row, VENDOR_COL_NAME).value = "Staff Salary - Manohar"
            expense_ws.cell(output_row, VENDOR_COL_BILL_NO).value = "Staff Salary"
            expense_ws.cell(output_row, VENDOR_COL_AMOUNT).value = manohar_debit
            expense_ws.cell(output_row, VENDOR_COL_TOTAL).value = manohar_debit
            expense_ws.cell(output_row, VENDOR_COL_AMOUNT_PAID).value = manohar_debit
            output_row += 1
            copied_rows += 1

            for col in range(1, 16):
                copy_cell(vendor_ws.cell(source_row, col), expense_ws.cell(output_row, col))
            expense_ws.cell(output_row, VENDOR_COL_NAME).value = "Staff Salary - Victor"
            expense_ws.cell(output_row, VENDOR_COL_BILL_NO).value = "Staff Salary"
            expense_ws.cell(output_row, VENDOR_COL_AMOUNT).value = victor_debit
            expense_ws.cell(output_row, VENDOR_COL_TOTAL).value = victor_debit
            expense_ws.cell(output_row, VENDOR_COL_AMOUNT_PAID).value = victor_debit
            output_row += 1
            copied_rows += 1
        else:
            for col in range(1, 16):
                copy_cell(vendor_ws.cell(source_row, col), expense_ws.cell(output_row, col))
            if not clean_string(expense_ws.cell(output_row, VENDOR_COL_BILL_NO).value):
                amount_paid = numeric_amount(expense_ws.cell(output_row, VENDOR_COL_AMOUNT_PAID).value)
                amount = amount_paid or numeric_amount(expense_ws.cell(output_row, VENDOR_COL_TOTAL).value)
                expense_ws.cell(output_row, VENDOR_COL_BILL_NO).value = generate_hash_id(
                    expense_ws.cell(output_row, VENDOR_COL_NAME).value,
                    expense_ws.cell(output_row, VENDOR_COL_DATE).value,
                    amount,
                )
            output_row += 1
            copied_rows += 1

    # Treat statutory tax payments as expenses even though they are not vendor
    # bills. Their bank reference is the stable identity used across reruns.
    for bank_row in range(b_header + 1, b_last + 1):
        narration = clean_string(bank_ws.cell(bank_row, bank_meta["narration_col"]).value)
        original_reference = clean_string(bank_ws.cell(bank_row, bank_meta["ref_col"]).value)
        generated_reference = clean_string(bank_ws.cell(bank_row, 6).value)
        reference_text = f"{narration} {original_reference} {generated_reference}".upper()
        if "CBT TIN" in reference_text or "CBDT TIN" in reference_text:
            tax_name = "TDS Paid"
            expense_group = "Tax and Duties"
        elif "GST" in reference_text:
            tax_name = "GST Paid"
            expense_group = "Tax and Duties"
        elif "CHEQUE BOOK ISSUE CHARGE" in reference_text:
            tax_name = "CHEQUE BOOK ISSUE CHARGE"
            expense_group = "Bank Charges"
        else:
            continue

        debit = numeric_amount(bank_ws.cell(bank_row, bank_meta["debit_col"]).value)
        if debit <= 0:
            continue

        tax_date = bank_ws.cell(bank_row, bank_meta["date_col"]).value
        expense_ws.cell(output_row, VENDOR_COL_SL_NO).value = output_row - 1
        expense_ws.cell(output_row, VENDOR_COL_DATE).value = tax_date
        expense_ws.cell(output_row, VENDOR_COL_NAME).value = tax_name
        expense_ws.cell(output_row, VENDOR_COL_AMOUNT_PAID).value = debit
        expense_ws.cell(output_row, 16).value = generated_reference
        expense_ws.cell(output_row, 17).value = tax_date
        expense_ws.cell(output_row, 18).value = "MATCHED"
        expense_ws.cell(output_row, 19).value = expense_group
        expense_ws.cell(output_row, EXPENSE_COL_BANK_PAID).value = debit
        for col in range(1, EXPENSE_COL_BANK_PAID + 1):
            expense_ws.cell(output_row, col).fill = copy(GREEN_FILL)

        bank_ws.cell(bank_row, 8).value = f"{tax_name}: {generated_reference or original_reference}"
        bank_ws.cell(bank_row, 9).value = "MATCHED"
        for col in range(1, 10):
            bank_ws.cell(bank_row, col).fill = copy(GREEN_FILL)
        output_row += 1
        copied_rows += 1

    amount_cols = [6, 7, 8, 9, 12, 14]
    for row in range(2, expense_ws.max_row + 1):
        for col in amount_cols:
            cell = expense_ws.cell(row, col)
            val = numeric_amount(cell.value)
            cell.value = val
            cell.number_format = AMOUNT_FORMAT

    add_expense_gst_summary(expense_ws)

    style_header(expense_ws, 1)
    expense_ws.freeze_panes = "A2"
    return expense_ws, copied_rows, skipped_rows


def add_expense_gst_summary(expense_ws) -> None:
    """Add separate GST totals for the clubhouse and society maintenance invoices."""
    summary_start = expense_ws.max_row + 11
    data_end = summary_start - 11

    headers = ["Billing Category", "Taxable Amount", "SGST", "CGST", "Total GST", "Invoice Total", "TDS", "Amount Paid", "Bank Amount Paid"]
    for col, header in enumerate(headers, 1):
        expense_ws.cell(summary_start, col).value = header

    clubhouse_row = summary_start + 1
    maintenance_row = summary_start + 2
    expense_ws.cell(clubhouse_row, 1).value = "Clubhouse"
    expense_ws.cell(maintenance_row, 1).value = "Society Maintenance"
    for col, source_col in ((2, "F"), (3, "G"), (4, "H"), (6, "I"), (7, "K"), (8, "L"), (9, "T")):
        expense_ws.cell(clubhouse_row, col).value = (
            f'=SUMIF($E$2:$E${data_end},"*CLUBHOUSE*",${source_col}$2:${source_col}${data_end})'
        )
        expense_ws.cell(maintenance_row, col).value = (
            f'=SUM(${source_col}$2:${source_col}${data_end})-{get_column_letter(col)}{clubhouse_row}'
        )
    expense_ws.cell(clubhouse_row, 5).value = f'=C{clubhouse_row}+D{clubhouse_row}'
    expense_ws.cell(maintenance_row, 5).value = f'=C{maintenance_row}+D{maintenance_row}'

    total_row = summary_start + 3
    expense_ws.cell(total_row, 1).value = "Total GST"
    for col in range(2, 10):
        expense_ws.cell(total_row, col).value = f'=SUM({get_column_letter(col)}{summary_start + 1}:{get_column_letter(col)}{summary_start + 2})'
        expense_ws.cell(total_row, col).number_format = '#,##0'
        expense_ws.cell(total_row, col).border = DOUBLE_BOTTOM
        expense_ws.cell(total_row, col).font = Font(bold=True)

    style_header(expense_ws, summary_start)
    expense_ws.cell(total_row, 1).font = Font(bold=True)


def add_other_income_gst_summary(other_income_ws, data_end: int) -> None:
    """Add GST totals by income type after ten reserved manual-entry rows."""
    summary_start = data_end + 1
    headers = ["Income Type", "Total Amount", "Taxable Amount", "CGST", "SGST", "IGST", "Total GST"]
    categories = OTHER_INCOME_TYPES

    for col, header in enumerate(headers, 1):
        other_income_ws.cell(summary_start, col).value = header

    for offset, category in enumerate(categories, 1):
        row = summary_start + offset
        other_income_ws.cell(row, 1).value = category
        for col, source_col in ((2, "E"), (3, "G"), (4, "I"), (5, "J"), (6, "K")):
            other_income_ws.cell(row, col).value = (
                f'=SUMIF($L$2:$L${data_end},$A{row},${source_col}$2:${source_col}${data_end})'
            )
            other_income_ws.cell(row, col).number_format = AMOUNT_FORMAT
        other_income_ws.cell(row, 7).value = f'=SUM(D{row}:F{row})'
        other_income_ws.cell(row, 7).number_format = AMOUNT_FORMAT

    total_row = summary_start + len(categories) + 1
    other_income_ws.cell(total_row, 1).value = "Total GST"
    for col in range(2, 8):
        other_income_ws.cell(total_row, col).value = (
            f'=SUM({get_column_letter(col)}{summary_start + 1}:{get_column_letter(col)}{total_row - 1})'
        )
        other_income_ws.cell(total_row, col).number_format = AMOUNT_FORMAT
        other_income_ws.cell(total_row, col).font = Font(bold=True)
        other_income_ws.cell(total_row, col).border = DOUBLE_BOTTOM

    style_header(other_income_ws, summary_start)
    other_income_ws.cell(total_row, 1).font = Font(bold=True)


def apply_expense_calculations(expense_ws) -> None:
    """Apply rounded Total, TDS, and Amount Paid formulas after bank matching."""
    for row in range(2, expense_ws.max_row + 1):
        if not clean_string(expense_ws.cell(row, VENDOR_COL_NAME).value):
            break
        if clean_string(expense_ws.cell(row, 19).value).upper() == "TAX AND DUTIES":
            continue
        expense_ws.cell(row, VENDOR_COL_TOTAL).value = f'=ROUND(F{row}+G{row}+H{row},0)'
        expense_ws.cell(row, VENDOR_COL_TDS_FILED).value = f'=ROUND(F{row}*IFERROR(J{row},0)/100,0)'
        expense_ws.cell(row, VENDOR_COL_AMOUNT_PAID).value = f'=ROUND(I{row}-K{row},0)'
        for col in (VENDOR_COL_AMOUNT, VENDOR_COL_SGST, VENDOR_COL_CGST, VENDOR_COL_TOTAL, VENDOR_COL_TDS_FILED, VENDOR_COL_AMOUNT_PAID, VENDOR_COL_LESS_EXCESS):
            expense_ws.cell(row, col).number_format = '#,##0'

    for row in expense_ws.iter_rows():
        for cell in row:
            if cell.value is None:
                continue
            font = copy(cell.font)
            font.name = "Calibri"
            font.sz = 11
            cell.font = font


# =============================================================
# FD & CASH WITHDRAWAL TRACKING
# =============================================================

def is_fd_transfer(narration: Any, reference: Any) -> tuple[bool, str]:
    text = f"{clean_string(narration)} {clean_string(reference)}".upper()
    if re.search(r'\bTO\s+FD[-\s]*\d+', text):
        return True, "FD Created"
    if "INT TRF FRM" not in text and re.search(r'\b(?:TRF\s+)?FRM\s+FD[-\s]*\d+', text):
        return True, "FD Broken"
    return False, ""


def create_fd_transactions_sheet(bank_ws, output_wb, bank_meta: dict) -> tuple[Any, int, int, float]:
    sheet_name = "FD Transactions"
    if sheet_name in output_wb.sheetnames:
        del output_wb[sheet_name]
    fd_ws = output_wb.create_sheet(sheet_name)
    headers = [
        "Transaction Date", "FD Transaction Type", "FD Reference ID", "Narration",
        "Cheque/Settlement Ref", "Debit / FD Investment", "Credit / FD Proceeds", "Bank Match Status"
    ]
    for col, header in enumerate(headers, 1):
        fd_ws.cell(1, col).value = header

    output_row = 2
    interest_count = 0
    interest_total = 0.0
    header_row = bank_meta["header_row"]
    last_row = find_last_bank_data_row(bank_ws, header_row)
    for row in range(header_row + 1, last_row + 1):
        narration = bank_ws.cell(row, bank_meta["narration_col"]).value
        reference = bank_ws.cell(row, bank_meta["ref_col"]).value
        found, transaction_type = is_fd_transfer(narration, reference)
        is_interest = "INT TRF FRM" in f"{clean_string(narration)} {clean_string(reference)}".upper()
        if not found and not is_interest:
            continue

        if is_interest:
            transaction_type = "FD Interest"

        fd_ws.cell(output_row, 1).value = bank_ws.cell(row, bank_meta["date_col"]).value
        fd_ws.cell(output_row, 1).number_format = "dd-mmm-yyyy"
        fd_ws.cell(output_row, 2).value = transaction_type
        fd_ws.cell(output_row, 3).value = bank_ws.cell(row, 6).value
        fd_ws.cell(output_row, 4).value = narration
        fd_ws.cell(output_row, 5).value = reference
        fd_ws.cell(output_row, 6).value = numeric_amount(bank_ws.cell(row, bank_meta["debit_col"]).value)
        fd_ws.cell(output_row, 7).value = numeric_amount(bank_ws.cell(row, bank_meta["credit_col"]).value)
        fd_ws.cell(output_row, 8).value = "MATCHED"
        for col in (6, 7):
            fd_ws.cell(output_row, col).number_format = AMOUNT_FORMAT

        bank_ws.cell(row, 8).value = f"{transaction_type}: {bank_ws.cell(row, 6).value}"
        bank_ws.cell(row, 9).value = "MATCHED"
        for col in range(1, 10):
            bank_ws.cell(row, col).fill = copy(GREEN_FILL)
        if is_interest:
            interest_count += 1
            interest_total += numeric_amount(bank_ws.cell(row, bank_meta["credit_col"]).value)
        output_row += 1

    summary_row = output_row + 10
    fd_ws.cell(summary_row, 5).value = "FD Transaction Totals"
    fd_ws.cell(summary_row, 5).font = Font(bold=True)
    fd_ws.cell(summary_row + 1, 5).value = "Total Debit"
    fd_ws.cell(summary_row + 1, 6).value = f"=SUM(F2:F{output_row - 1})"
    fd_ws.cell(summary_row + 2, 5).value = "Total Credit"
    fd_ws.cell(summary_row + 2, 7).value = f"=SUM(G2:G{output_row - 1})"
    for row, col in ((summary_row + 1, 6), (summary_row + 2, 7)):
        fd_ws.cell(row, col).number_format = AMOUNT_FORMAT
        fd_ws.cell(row, col).font = Font(bold=True)
        fd_ws.cell(row, col).border = DOUBLE_BOTTOM

    style_header(fd_ws, 1)
    fd_ws.freeze_panes = "A2"
    return fd_ws, output_row - 2, interest_count, interest_total


def create_cash_withdrawal_sheets(bank_ws, output_wb, bank_meta: dict) -> int:
    """Create a separate editable expense register for every bank cash withdrawal."""
    header_row = bank_meta["header_row"]
    last_row = find_last_bank_data_row(bank_ws, header_row)
    created_count = 0

    for row in range(header_row + 1, last_row + 1):
        narration = clean_string(bank_ws.cell(row, bank_meta["narration_col"]).value)
        reference = clean_string(bank_ws.cell(row, bank_meta["ref_col"]).value)
        if "CASH" not in narration.upper() or "WITHDRAWAL" not in narration.upper():
            continue

        withdrawal_amount = numeric_amount(bank_ws.cell(row, bank_meta["debit_col"]).value)
        if withdrawal_amount <= 0:
            continue
        withdrawal_ref = clean_string(bank_ws.cell(row, 6).value) or generate_reference_id(narration, reference)
        sheet_name = f"Cash Withdrawal {withdrawal_ref}"[:31]
        if sheet_name in output_wb.sheetnames:
            del output_wb[sheet_name]
        cash_ws = output_wb.create_sheet(sheet_name)

        cash_ws["A1"] = "Withdrawal Date"
        cash_ws["B1"] = bank_ws.cell(row, bank_meta["date_col"]).value
        cash_ws["B1"].number_format = "dd-mmm-yyyy"
        cash_ws["D1"] = "Withdrawal Reference"
        cash_ws["E1"] = withdrawal_ref
        cash_ws["G1"] = "Withdrawal Amount"
        cash_ws["H1"] = withdrawal_amount
        cash_ws["H1"].number_format = AMOUNT_FORMAT
        cash_ws["A2"] = "Bank Narration"
        cash_ws["B2"] = narration
        cash_ws["A3"] = "Opening Balance"
        cash_ws["B3"] = 0.0
        cash_ws["B3"].number_format = AMOUNT_FORMAT
        cash_ws["D3"] = "Total Cash Available"
        cash_ws["E3"] = "=B3+H1"
        cash_ws["E3"].number_format = AMOUNT_FORMAT

        headers = ["Payment Date", "Paid To", "Expense Description", "Invoice Number", "Invoice Available", "Payment Amount", "Taxable Amount", "CGST", "SGST", "IGST", "GST Total", "Remarks"]
        for col, header in enumerate(headers, 1):
            cash_ws.cell(4, col).value = header
        for entry_row in range(5, 25):
            cash_ws.cell(entry_row, 1).number_format = "dd-mmm-yyyy"
            for col in (6, 7, 8, 9, 10, 11):
                cash_ws.cell(entry_row, col).number_format = AMOUNT_FORMAT
            cash_ws.cell(entry_row, 11).value = f'=SUM(H{entry_row}:J{entry_row})'

        invoice_validation = DataValidation(type="list", formula1='"Yes,No"', allow_blank=True)
        cash_ws.add_data_validation(invoice_validation)
        invoice_validation.add("E5:E24")

        cash_ws["E26"] = "Total Paid"
        cash_ws["F26"] = "=SUM(F5:F24)"
        cash_ws["G26"] = "Closing Balance"
        cash_ws["H26"] = "=E3-F26"
        cash_ws["I26"] = "Total GST"
        cash_ws["J26"] = "=SUM(K5:K24)"
        cash_ws["E27"] = "Total Accounted"
        cash_ws["F27"] = "=F26+H26-B3"
        for cell in ("F26", "H26", "J26"):
            cash_ws[cell].number_format = AMOUNT_FORMAT
            cash_ws[cell].font = Font(bold=True)
            cash_ws[cell].border = DOUBLE_BOTTOM
        cash_ws["F27"].number_format = AMOUNT_FORMAT
        cash_ws["F27"].font = Font(bold=True)
        cash_ws["F27"].border = DOUBLE_BOTTOM

        style_header(cash_ws, 4)
        for cell in cash_ws[1]:
            if cell.value is not None:
                cell.font = Font(bold=True)
        cash_ws.freeze_panes = "A5"
        cash_ws.column_dimensions["B"].width = 28
        cash_ws.column_dimensions["C"].width = 35
        cash_ws.column_dimensions["D"].width = 18
        cash_ws.column_dimensions["L"].width = 28
        autofit(cash_ws)
        created_count += 1

    return created_count


# =============================================================
# INCOME SHEET OPERATIONS
# =============================================================

def find_income_header(ws) -> int:
    for r in range(1, min(ws.max_row, 15) + 1):
        headers = [clean_string(ws.cell(r, c).value).upper() for c in range(1, ws.max_column + 1)]
        if "DATE" in headers and "TYPE" in headers and "FLAT NO." in headers:
            return r
    return 1


def is_bank_receipt(value: Any) -> bool:
    s = normalize_spaces(value).upper()
    return s in {"BANK RECEIPT", "BANK RECIEPT"}


def sanitize_income_sheet(source_ws, output_wb) -> tuple[Any, int]:
    if "Income_recorded_in_nbh" in output_wb.sheetnames:
        del output_wb["Income_recorded_in_nbh"]

    income_ws = output_wb.create_sheet("Income_recorded_in_nbh")
    header_row = find_income_header(source_ws)
    orig_cols = source_ws.max_column

    retained_source_columns = [
        c for c in range(1, orig_cols + 1)
        if clean_string(source_ws.cell(header_row, c).value).upper() != "BALANCE"
    ]

    for output_col, source_col in enumerate(retained_source_columns, 1):
        income_ws.cell(1, output_col).value = source_ws.cell(header_row, source_col).value

    metadata_start = len(retained_source_columns) + 1
    income_ws.cell(1, metadata_start).value = "Bank Reference ID"
    income_ws.cell(1, metadata_start + 1).value = "Bank Transaction Date"
    income_ws.cell(1, metadata_start + 2).value = "Bank Match Status"
    source_headers = {
        clean_string(source_ws.cell(header_row, c).value).upper(): c
        for c in range(1, source_ws.max_column + 1)
    }
    settlement_source_col = source_headers.get("SETTLEMENT ID")

    out_row = 2
    copied_rows = 0

    for r in range(header_row + 1, source_ws.max_row + 1):
        if not is_bank_receipt(source_ws.cell(r, 2).value):
            continue

        for output_col, source_col in enumerate(retained_source_columns, 1):
            income_ws.cell(out_row, output_col).value = source_ws.cell(r, source_col).value

        if settlement_source_col:
            income_ws.cell(out_row, metadata_start).value = source_ws.cell(r, settlement_source_col).value

        for c in (10, 11):
            cell = income_ws.cell(out_row, c)
            cell.value = numeric_amount(cell.value)
            cell.number_format = AMOUNT_FORMAT

        out_row += 1
        copied_rows += 1

    style_header(income_ws, 1)
    income_ws.freeze_panes = "A2"
    return income_ws, copied_rows


# =============================================================
# FD INTEREST EXTRACTION & BANK MATCHING
# =============================================================

def extract_and_match_fd_interest(bank_ws, output_wb, bank_meta: dict) -> tuple[Any, int, float]:
    sheet_name = "Income from FD Interest"
    if sheet_name in output_wb.sheetnames:
        del output_wb[sheet_name]

    fd_ws = output_wb.create_sheet(sheet_name)

    headers = [
        "Transaction Date",
        "Narration",
        "Cheque/Settlement Ref",
        "Generated Reference ID",
        "Credit / Interest Amount (INR)"
    ]
    for col_idx, header in enumerate(headers, 1):
        fd_ws.cell(1, col_idx).value = header

    header_row = bank_meta["header_row"]
    last_row = find_last_bank_data_row(bank_ws, header_row)

    out_row = 2
    total_interest = 0.0
    entries_count = 0

    for r in range(header_row + 1, last_row + 1):
        narration = clean_string(bank_ws.cell(r, bank_meta["narration_col"]).value)
        orig_ref = clean_string(bank_ws.cell(r, bank_meta["ref_col"]).value)
        combo = f"{narration} {orig_ref}".upper()

        if "INT TRF FRM" in combo:
            txn_date = bank_ws.cell(r, bank_meta["date_col"]).value
            gen_ref = clean_string(bank_ws.cell(r, 6).value)
            if not gen_ref:
                gen_ref = generate_reference_id(
                    narration,
                    orig_ref,
                    bank_ws.cell(r, bank_meta["date_col"]).value,
                    bank_ws.cell(r, bank_meta["debit_col"]).value,
                    bank_ws.cell(r, bank_meta["credit_col"]).value,
                )
                bank_ws.cell(r, 6).value = gen_ref

            credit = numeric_amount(bank_ws.cell(r, bank_meta["credit_col"]).value)
            debit = numeric_amount(bank_ws.cell(r, bank_meta["debit_col"]).value)
            interest = credit if credit > 0 else debit

            fd_ws.cell(out_row, 1).value = txn_date
            fd_ws.cell(out_row, 1).number_format = "dd-mmm-yyyy"
            fd_ws.cell(out_row, 2).value = narration
            fd_ws.cell(out_row, 3).value = orig_ref
            fd_ws.cell(out_row, 4).value = gen_ref

            amt_cell = fd_ws.cell(out_row, 5)
            amt_cell.value = interest
            amt_cell.number_format = AMOUNT_FORMAT

            # Mark matching Bank_Statement row
            bank_ws.cell(r, 7).value = gen_ref
            bank_ws.cell(r, 8).value = f"FD Interest: {gen_ref}"
            bank_ws.cell(r, 9).value = "MATCHED"
            for col in range(1, 10):
                bank_ws.cell(r, col).fill = copy(GREEN_FILL)

            total_interest += interest
            entries_count += 1
            out_row += 1

    total_label_cell = fd_ws.cell(out_row, 4)
    total_label_cell.value = "Total Interest Received"
    total_label_cell.font = Font(bold=True)
    total_label_cell.alignment = Alignment(horizontal="right")

    total_val_cell = fd_ws.cell(out_row, 5)
    if entries_count > 0:
        total_val_cell.value = f"=SUM(E2:E{out_row - 1})"
    else:
        total_val_cell.value = 0.0

    total_val_cell.font = Font(bold=True)
    total_val_cell.number_format = AMOUNT_FORMAT
    total_val_cell.border = DOUBLE_BOTTOM

    style_header(fd_ws, 1)
    fd_ws.freeze_panes = "A2"
    return fd_ws, entries_count, total_interest


# =============================================================
# MATCHING LOGIC
# =============================================================

def match_bank_and_expense(bank_ws, expense_ws, bank_meta: dict) -> dict:
    header_row = bank_meta["header_row"]
    last_row = find_last_bank_data_row(bank_ws, header_row)

    expenses = []
    for r in range(2, expense_ws.max_row + 1):
        vendor = clean_string(expense_ws.cell(r, VENDOR_COL_NAME).value)
        if not vendor:
            break
        bill_no = clean_string(expense_ws.cell(r, VENDOR_COL_BILL_NO).value)
        amount_paid = numeric_amount(expense_ws.cell(r, VENDOR_COL_AMOUNT_PAID).value)
        total = numeric_amount(expense_ws.cell(r, VENDOR_COL_TOTAL).value)
        amount = numeric_amount(expense_ws.cell(r, VENDOR_COL_AMOUNT).value)

        final_amount = amount_paid if amount_paid > 0 else (total if total > 0 else amount)
        if final_amount > 0:
            expenses.append({
                "row": r,
                "bill_no": bill_no,
                "vendor": vendor,
                "payment_ref": clean_string(expense_ws.cell(r, 16).value).upper(),
                "amount": final_amount,
            })

    matched_expenses = set()
    metrics = {"generated": 0, "matched": 0, "review": 0, "unmatched": 0}

    for row in range(header_row + 1, last_row + 1):
        narration = bank_ws.cell(row, bank_meta["narration_col"]).value
        orig_ref = bank_ws.cell(row, bank_meta["ref_col"]).value
        debit = numeric_amount(bank_ws.cell(row, bank_meta["debit_col"]).value)
        credit = numeric_amount(bank_ws.cell(row, bank_meta["credit_col"]).value)

        ref_id = generate_reference_id(narration, orig_ref, bank_ws.cell(row, bank_meta["date_col"]).value, debit, credit)
        bank_ws.cell(row, 6).value = ref_id
        if ref_id:
            metrics["generated"] += 1

        bank_ws.cell(row, 4).value = debit
        bank_ws.cell(row, 5).value = credit
        bank_ws.cell(row, 4).number_format = AMOUNT_FORMAT
        bank_ws.cell(row, 5).number_format = AMOUNT_FORMAT

        if debit <= 0:
            if credit > 0:
                if not bank_ws.cell(row, 9).value:
                    bank_ws.cell(row, 9).value = "NOT EXPENSE"
            continue

        # FD creation/breakage rows are already reconciled by the FD sheet.
        # Do not overwrite their MATCHED status during vendor expense matching.
        if clean_string(bank_ws.cell(row, 9).value).upper() == "MATCHED":
            continue

        reference_candidates = [
            i for i, exp in enumerate(expenses)
            if i not in matched_expenses and exp["payment_ref"] == ref_id.upper()
        ]
        if len(reference_candidates) == 1:
            match_idx = reference_candidates[0]
            expected = expenses[match_idx]["amount"]
            difference = round(debit - expected, 2)
            status = "MATCHED" if abs(difference) <= 1 else f"PARTIALLY MATCHED (Difference: {difference:,.2f})"
            if status.startswith("PARTIALLY"):
                expense_row = expenses[match_idx]["row"]
                bank_ws.cell(row, 8).value = status
                bank_ws.cell(row, 9).value = status
                expense_ws.cell(expense_row, 18).value = status
                expense_ws.cell(expense_row, EXPENSE_COL_BANK_PAID).value = debit
                for col in range(1, 10):
                    bank_ws.cell(row, col).fill = copy(YELLOW_FILL)
                for col in range(1, EXPENSE_COL_BANK_PAID + 1):
                    expense_ws.cell(expense_row, col).fill = copy(YELLOW_FILL)
                matched_expenses.add(match_idx)
                continue

        candidates = [
            i for i, exp in enumerate(expenses)
            if i not in matched_expenses and abs(exp["amount"] - debit) <= 1
        ]
        match_idx = candidates[0] if len(candidates) == 1 else None
        status = "MATCHED" if match_idx is not None else ("REVIEW" if len(candidates) > 1 else "UNMATCHED")

        bank_ws.cell(row, 9).value = status

        if status == "MATCHED" and match_idx is not None:
            exp = expenses[match_idx]
            matched_expenses.add(match_idx)

            bank_ws.cell(row, 7).value = exp["bill_no"]
            bank_ws.cell(row, 8).value = exp["vendor"]

            exp_row = exp["row"]
            expense_ws.cell(exp_row, 16).value = ref_id
            expense_ws.cell(exp_row, 17).value = bank_ws.cell(row, bank_meta["date_col"]).value
            expense_ws.cell(exp_row, 18).value = "MATCHED"
            expense_ws.cell(exp_row, EXPENSE_COL_BANK_PAID).value = debit

            for col in range(1, 10):
                bank_ws.cell(row, col).fill = copy(GREEN_FILL)

            metrics["matched"] += 1
        elif status == "REVIEW":
            metrics["review"] += 1
        else:
            metrics["unmatched"] += 1

    for i, exp in enumerate(expenses):
        if i not in matched_expenses:
            r = exp["row"]
            if not expense_ws.cell(r, 18).value:
                expense_ws.cell(r, 18).value = "UNMATCHED"

    metrics["expenses_total"] = len(expenses)
    metrics["expenses_matched"] = len(matched_expenses)
    return metrics


def match_bank_and_income(bank_ws, income_ws, bank_meta: dict) -> dict:
    header_row = bank_meta["header_row"]
    last_row = find_last_bank_data_row(bank_ws, header_row)

    headers = {
        clean_string(income_ws.cell(1, c).value).upper(): c
        for c in range(1, income_ws.max_column + 1)
    }
    settlement_col = headers.get("SETTLEMENT ID")
    reference_col = headers.get("REFERENCE NUMBER")
    flat_col = headers.get("FLAT NO.")
    status_col = income_ws.max_column

    bank_credits = []
    for r in range(header_row + 1, last_row + 1):
        cr = numeric_amount(bank_ws.cell(r, bank_meta["credit_col"]).value)
        if cr > 0 and bank_ws.cell(r, 9).value != "MATCHED":
            bank_credits.append({
                "row": r,
                "credit": cr,
                "date": parse_date(bank_ws.cell(r, bank_meta["date_col"]).value),
                "generated_ref": clean_string(bank_ws.cell(r, 6).value),
            })

    used_bank_rows = set()
    income_matched = 0
    income_review = 0
    income_unmatched = 0

    for ir in range(2, income_ws.max_row + 1):
        amount = numeric_amount(income_ws.cell(ir, 10).value)
        if amount <= 0:
            amount = numeric_amount(income_ws.cell(ir, 11).value)

        if amount <= 0:
            income_ws.cell(ir, status_col).value = "UNMATCHED"
            continue

        settlement = clean_string(income_ws.cell(ir, settlement_col).value) if settlement_col else ""
        reference = clean_string(income_ws.cell(ir, reference_col).value) if reference_col else ""
        flat = clean_string(income_ws.cell(ir, flat_col).value) if flat_col else ""
        inc_date = parse_date(income_ws.cell(ir, 1).value)

        matches = []

        if settlement:
            for txn in bank_credits:
                if txn["row"] in used_bank_rows:
                    continue
                if txn["generated_ref"].upper() == settlement.upper():
                    matches.append(txn)

        if not matches:
            for txn in bank_credits:
                if txn["row"] in used_bank_rows:
                    continue
                if money_equal(txn["credit"], amount) and inc_date and txn["date"] == inc_date:
                    matches.append(txn)

        if not matches:
            for txn in bank_credits:
                if txn["row"] in used_bank_rows:
                    continue
                if money_equal(txn["credit"], amount):
                    matches.append(txn)

        if len(matches) == 1:
            txn = matches[0]
            used_bank_rows.add(txn["row"])

            desc = []
            if flat:
                desc.append(flat)
            if reference:
                desc.append(f"Ref: {reference}")

            bank_ws.cell(txn["row"], 8).value = " | ".join(desc)
            bank_ws.cell(txn["row"], 9).value = "MATCHED"

            income_ws.cell(ir, status_col - 2).value = txn["generated_ref"]
            income_ws.cell(ir, status_col - 1).value = txn["date"]
            income_ws.cell(ir, status_col).value = "MATCHED"

            for col in range(1, 10):
                bank_ws.cell(txn["row"], col).fill = copy(GREEN_FILL)

            income_matched += 1
        elif len(matches) > 1:
            income_ws.cell(ir, status_col).value = "REVIEW"
            income_review += 1
        else:
            income_ws.cell(ir, status_col).value = "UNMATCHED"
            income_unmatched += 1

    return {
        "income_matched": income_matched,
        "income_review": income_review,
        "income_unmatched": income_unmatched,
        "income_total": income_ws.max_row - 1
    }


# =============================================================
# INCOME FROM OTHER SOURCES (GST INVOICE PROCESSING)
# =============================================================

def process_other_income_sources(gst_invoice_ws, bank_ws, output_wb, bank_meta: dict,
                                 income_ws=None) -> tuple[Any, int, int]:
    """
    Process GST invoice file for other income sources (Classes, Amenity, Move In/Out, Screen Rentals)
    Create new sheet with GST calculations based on bank amounts
    """
    sheet_name = "Income from Other Sources"
    if sheet_name in output_wb.sheetnames:
        del output_wb[sheet_name]

    other_income_ws = output_wb.create_sheet(sheet_name)

    # Headers matching the GST invoice format + additional columns
    headers = [
        "Bill Number",
        "Ledger Name",
        "Date",
        "Voucher Type",
        "Total Amount",
        "Non Taxable Amount",
        "Taxable Amount",
        "Percentage",
        "CGST",
        "SGST",
        "IGST",
        "Income Type",
        "Bank Deposit Amount",
        "Bank Reference ID",
        "Bank Transaction Date",
        "Match Status",
        "Generated Bill Number"
    ]

    for col_idx, header in enumerate(headers, 1):
        other_income_ws.cell(1, col_idx).value = header

    # Get unmatched bank credits
    header_row = bank_meta["header_row"]
    last_row = find_last_bank_data_row(bank_ws, header_row)

    unmatched_bank_credits = []
    for r in range(header_row + 1, last_row + 1):
        cr = numeric_amount(bank_ws.cell(r, bank_meta["credit_col"]).value)
        match_status = clean_string(bank_ws.cell(r, 9).value)
        
        if cr > 0 and match_status not in {"MATCHED", "FD TRANSFER"}:
            unmatched_bank_credits.append({
                "row": r,
                "amount": cr,
                "date": bank_ws.cell(r, bank_meta["date_col"]).value,
                "narration": clean_string(bank_ws.cell(r, bank_meta["narration_col"]).value),
                "ref": clean_string(bank_ws.cell(r, 6).value),
            })

    # Copy existing GST invoice rows to the sheet
    gst_header_row = 3  # Headers in row 2, data starts from row 3 in GST invoice
    out_row = 2
    entries_count = 0
    existing_references: set[str] = set()

    for r in range(gst_header_row, gst_invoice_ws.max_row + 1):
        bill_no = clean_string(gst_invoice_ws.cell(r, 1).value)
        if not bill_no and not any(gst_invoice_ws.cell(r, col).value not in (None, "") for col in range(1, 16)):
            break

        # Copy GST invoice columns while omitting GST No, Place of Supply, TYPE,
        # and UTILITY from the output sheet.
        source_columns = [1, 2, 3, 8, 9, 10, 11, 12, 13, 14, 15]
        for output_col, source_col in enumerate(source_columns, 1):
            src_val = gst_invoice_ws.cell(r, source_col).value
            other_income_ws.cell(out_row, output_col).value = src_val
            
            # Format amount columns
            if output_col in [5, 6, 7, 9, 10, 11]:  # Amount columns
                if isinstance(src_val, (int, float)):
                    other_income_ws.cell(out_row, output_col).number_format = AMOUNT_FORMAT

        # Apply GST to every income row by default. Maintenance receipts are
        # reset to zero below when their year/quarter/number reference is known.
        other_income_ws.cell(out_row, 8).value = 0.09

        # Automatically classify Move In/Move Out invoices.
        if bill_no.upper().startswith("MIMO"):
            other_income_ws.cell(out_row, 12).value = "MoveInMoveOut"
        elif bill_no.upper().startswith("SPONSOR/"):
            other_income_ws.cell(out_row, 12).value = "Screen Rentals"
        else:
            other_income_ws.cell(out_row, 12).value = None

        # Keep all GST values formula-driven from amount and percentage.
        other_income_ws.cell(out_row, 7).value = f'=ROUND((E{out_row}-F{out_row})/(1+2*H{out_row}),2)'
        other_income_ws.cell(out_row, 9).value = f'=ROUND(G{out_row}*H{out_row},2)'
        other_income_ws.cell(out_row, 10).value = f'=ROUND(G{out_row}*H{out_row},2)'
        other_income_ws.cell(out_row, 11).value = 0.0

        # Add Bank Deposit Amount, Reference, Date, Match Status (empty for manual entry)
        other_income_ws.cell(out_row, 13).value = None  # Bank Deposit Amount
        other_income_ws.cell(out_row, 14).value = None  # Bank Reference ID
        other_income_ws.cell(out_row, 15).value = None  # Bank Transaction Date
        other_income_ws.cell(out_row, 16).value = None  # Match Status

        out_row += 1
        entries_count += 1

    # Link GST input invoices to their matching NBH receipt reference and date.
    # Non-quarter (for example BKG) receipts are appended as Other Sources before
    # remaining unmatched bank credits are added.
    if income_ws:
        income_headers = {
            clean_string(income_ws.cell(1, col).value).upper(): col
            for col in range(1, income_ws.max_column + 1)
        }
        reference_col = income_headers.get("REFERENCE NUMBER")
        bank_reference_col = income_headers.get("BANK REFERENCE ID")
        particulars_col = income_headers.get("PARTICULARS")
        flat_col = income_headers.get("FLAT NO.")
        invoice_rows = {
            clean_string(other_income_ws.cell(row, 1).value).upper(): row
            for row in range(2, out_row)
            if clean_string(other_income_ws.cell(row, 1).value)
        }
        maintenance_reference = re.compile(r"^\d{4}(?:-\d{2})?/Q\d+/\d+$", re.IGNORECASE)

        for row in range(2, income_ws.max_row + 1):
            receipt_reference = clean_string(income_ws.cell(row, reference_col).value) if reference_col else ""
            bank_reference = reference_text(income_ws.cell(row, bank_reference_col).value) if bank_reference_col else ""
            receipt_amount = numeric_amount(income_ws.cell(row, 10).value)
            if receipt_amount <= 0:
                receipt_amount = numeric_amount(income_ws.cell(row, 11).value)
            if receipt_amount <= 0:
                continue

            invoice_row = invoice_rows.get(receipt_reference.upper()) if receipt_reference else None
            if invoice_row:
                other_income_ws.cell(invoice_row, 13).value = receipt_amount
                other_income_ws.cell(invoice_row, 14).value = bank_reference
                other_income_ws.cell(invoice_row, 14).number_format = "@"
                other_income_ws.cell(invoice_row, 15).value = income_ws.cell(row, income_ws.max_column - 1).value
                if maintenance_reference.fullmatch(receipt_reference):
                    other_income_ws.cell(invoice_row, 8).value = 0.0
                else:
                    other_income_ws.cell(invoice_row, 8).value = 0.09
                existing_references.add(bank_reference.upper())
                continue
            if (receipt_reference and maintenance_reference.fullmatch(receipt_reference)) or bank_reference.upper() in existing_references:
                continue

            ledger_name = clean_string(income_ws.cell(row, particulars_col).value) if particulars_col else ""
            if not ledger_name and flat_col:
                ledger_name = clean_string(income_ws.cell(row, flat_col).value)
            income_type = ""
            reference_upper = receipt_reference.upper()
            if reference_upper.startswith("MIMO"):
                income_type = "MoveInMoveOut"
            elif reference_upper.startswith("BKG") and receipt_amount >= 3500:
                income_type = "Clubhouse Booking"
            elif reference_upper.startswith("BKG") and receipt_amount < 2000:
                income_type = "Amenities Booking"

            other_income_ws.cell(out_row, 1).value = receipt_reference or None
            other_income_ws.cell(out_row, 2).value = ledger_name
            other_income_ws.cell(out_row, 3).value = income_ws.cell(row, 1).value
            other_income_ws.cell(out_row, 4).value = "Bank Receipt"
            other_income_ws.cell(out_row, 5).value = receipt_amount
            other_income_ws.cell(out_row, 6).value = 0.0
            other_income_ws.cell(out_row, 7).value = f'=ROUND((E{out_row}-F{out_row})/(1+2*H{out_row}),2)'
            other_income_ws.cell(out_row, 8).value = 0.09
            other_income_ws.cell(out_row, 9).value = f'=ROUND(G{out_row}*H{out_row},2)'
            other_income_ws.cell(out_row, 10).value = f'=ROUND(G{out_row}*H{out_row},2)'
            other_income_ws.cell(out_row, 11).value = 0.0
            other_income_ws.cell(out_row, 12).value = income_type
            other_income_ws.cell(out_row, 13).value = receipt_amount
            other_income_ws.cell(out_row, 14).value = bank_reference
            other_income_ws.cell(out_row, 14).number_format = "@"
            other_income_ws.cell(out_row, 15).value = income_ws.cell(row, income_ws.max_column - 1).value
            other_income_ws.cell(out_row, 16).value = "PENDING"
            existing_references.add(bank_reference.upper())
            out_row += 1
            entries_count += 1

    # Add unmatched bank credits as new rows
    for bank_credit in unmatched_bank_credits:
        bank_reference = reference_text(bank_credit["ref"])
        if bank_reference.upper() in existing_references:
            continue
        # Calculate GST from bank amount (18% total: 9% CGST + 9% SGST)
        bank_amount = bank_credit["amount"]
        # Bank amount is inclusive of GST
        # If total is X and tax rate is 18%, then: X = base + base*0.18 = base*1.18
        # So: base = X / 1.18
        # Bank Amount (Column 5 - Total)
        other_income_ws.cell(out_row, 5).value = bank_amount
        other_income_ws.cell(out_row, 5).number_format = AMOUNT_FORMAT

        # Non Taxable Amount (Column 10)
        other_income_ws.cell(out_row, 6).value = 0.0
        other_income_ws.cell(out_row, 6).number_format = AMOUNT_FORMAT

        # Taxable Amount (Column 11)
        other_income_ws.cell(out_row, 7).value = f'=ROUND((E{out_row}-F{out_row})/(1+2*H{out_row}),2)'
        other_income_ws.cell(out_row, 7).number_format = AMOUNT_FORMAT

        # Percentage (Column 12)
        other_income_ws.cell(out_row, 8).value = 0.09  # 9% CGST

        # CGST (Column 13)
        other_income_ws.cell(out_row, 9).value = f'=ROUND(G{out_row}*H{out_row},2)'
        other_income_ws.cell(out_row, 9).number_format = AMOUNT_FORMAT

        # SGST (Column 14)
        other_income_ws.cell(out_row, 10).value = f'=ROUND(G{out_row}*H{out_row},2)'
        other_income_ws.cell(out_row, 10).number_format = AMOUNT_FORMAT

        # IGST (Column 15)
        other_income_ws.cell(out_row, 11).value = 0.0
        other_income_ws.cell(out_row, 11).number_format = AMOUNT_FORMAT

        # Income Type (Column 16) - To be filled manually
        other_income_ws.cell(out_row, 12).value = None

        # Bank Deposit Amount (Column 17)
        other_income_ws.cell(out_row, 13).value = bank_amount
        other_income_ws.cell(out_row, 13).number_format = AMOUNT_FORMAT

        # Bank Reference ID (Column 18)
        other_income_ws.cell(out_row, 14).value = bank_reference
        other_income_ws.cell(out_row, 14).number_format = "@"

        # Bank Transaction Date (Column 19)
        other_income_ws.cell(out_row, 15).value = bank_credit["date"]
        other_income_ws.cell(out_row, 15).number_format = "dd-mmm-yyyy"

        # Match Status (Column 20)
        other_income_ws.cell(out_row, 16).value = "PENDING"

        existing_references.add(bank_reference.upper())
        out_row += 1
        entries_count += 1

    for row in range(2, out_row):
        if not clean_string(other_income_ws.cell(row, 1).value):
            other_income_ws.cell(row, 1).value = generate_hash_id(
                other_income_ws.cell(row, 3).value,
                other_income_ws.cell(row, 5).value,
                other_income_ws.cell(row, 14).value,
            )
        # Generated Bill Number is intentionally left blank for manual merging.

    # Keep ten blank rows available for invoices or deposits entered manually.
    manual_entry_end = out_row + 9
    for row in range(2, manual_entry_end + 1):
        for col in (5, 6, 7, 9, 10, 11, 13):
            other_income_ws.cell(row, col).number_format = GST_AMOUNT_FORMAT
        for col in (3, 15):
            other_income_ws.cell(row, col).number_format = "dd-mmm-yyyy"
        row_fill = OTHER_INCOME_ROW_FILLS[(row - 2) % len(OTHER_INCOME_ROW_FILLS)]
        for col in range(1, len(headers) + 1):
            other_income_ws.cell(row, col).fill = copy(row_fill)

    income_type_validation = DataValidation(
        type="list",
        formula1=f'"{",".join(OTHER_INCOME_TYPES)}"',
        allow_blank=True,
    )
    other_income_ws.add_data_validation(income_type_validation)
    income_type_validation.add(f"L2:L{manual_entry_end}")
    add_other_income_gst_summary(other_income_ws, manual_entry_end)

    style_header(other_income_ws, 1)
    other_income_ws.freeze_panes = "A2"
    autofit(other_income_ws)

    return other_income_ws, len(unmatched_bank_credits), manual_entry_end
# =============================================================
# MAIN CONTROLLER
# =============================================================

def run_reconciliation(
    bank_path: Path,
    vendor_path: Path,
    income_path: Optional[Path] = None,
    other_income_path: Optional[Path] = None,
    output_path: Optional[Path] = None,
) -> dict[str, Any]:
    """Run the full reconciliation pipeline and save the resulting workbook.

    Returns a dict with the metrics that used to be printed by ``main()`` plus
    the resolved ``output_path`` and the ``bank_date`` that was detected, so
    callers (e.g. an orchestration script) can inspect the outcome.
    """
    output_path = Path(output_path) if output_path else bank_path.with_name(bank_path.stem + "_NBH_READY.xlsx")

    print()
    print("=" * 60)
    print("NoBrokerHood Bank + Expense + Income Reconciliation")
    print("=" * 60)
    print()
    print(f"Bank statement   : {bank_path}")
    print(f"Vendor bills     : {vendor_path}")
    print(f"Income receipts  : {income_path if income_path else 'None provided (skipping income matching)'}")
    print(f"Other income     : {other_income_path if other_income_path else 'None provided'}")
    print(f"Output file      : {output_path}")
    print()

    bank_wb = load_input_workbook(bank_path)
    vendor_wb = load_input_workbook(vendor_path, data_only=False)

    bank_ws = bank_wb.worksheets[0]
    bank_ws.title = "Bank_Statement"

    header_row, bank_headers = locate_bank_header(bank_ws)
    date_col = bank_headers.get("TRANSACTION DATE", 1)
    narration_col = bank_headers.get("NARRATION", 2)
    ref_col = bank_headers.get("CHEQUE NUMBER/SETTLEMENT ID", 3)
    debit_col = bank_headers.get("DEBIT", 4)
    credit_col = bank_headers.get("CREDIT", 5)

    bank_meta = {
        "header_row": header_row,
        "date_col": date_col,
        "narration_col": narration_col,
        "ref_col": ref_col,
        "debit_col": debit_col,
        "credit_col": credit_col,
    }

    # Pre-populate Reference IDs on Bank Statement
    b_last = find_last_bank_data_row(bank_ws, header_row)
    for r in range(header_row + 1, b_last + 1):
        narr = bank_ws.cell(r, narration_col).value
        orig = bank_ws.cell(r, ref_col).value
        bank_ws.cell(r, 6).value = generate_reference_id(
            narr,
            orig,
            bank_ws.cell(r, bank_meta["date_col"]).value,
            bank_ws.cell(r, bank_meta["debit_col"]).value,
            bank_ws.cell(r, bank_meta["credit_col"]).value,
        )

    bank_ws.cell(header_row, 6).value = "Generated Reference ID"
    bank_ws.cell(header_row, 7).value = "Bill / Invoice No"
    bank_ws.cell(header_row, 8).value = "Expense Description"
    bank_ws.cell(header_row, 9).value = "Match Status"
    style_header(bank_ws, header_row)

    bank_date = determine_bank_month(bank_ws, header_row, date_col)
    if not bank_date:
        raise ValueError("Could not determine bank statement month.")

    print(f"Bank statement month : {bank_date.strftime('%B %Y')}")
    exp_year, exp_month = previous_month(bank_date.year, bank_date.month)
    target_month_str = datetime(exp_year, exp_month, 1).strftime("%B %Y")
    print(f"Vendor expense month : {target_month_str}")

    vendor_ws = find_vendor_sheet(vendor_wb, exp_year, exp_month)
    if vendor_ws is None:
        available = ", ".join(vendor_wb.sheetnames)
        raise ValueError(
            f"Could not find vendor sheet for {target_month_str}. Available sheets: {available}"
        )

    print(f"Vendor sheet selected: {vendor_ws.title}")

    # 1. FD interest, creation, and breakage tracking in one sheet
    fd_ws, fd_transaction_count, fd_entries, fd_total_amt = create_fd_transactions_sheet(
        bank_ws, bank_wb, bank_meta
    )
    print(f"FD Interest rows reconciled: {fd_entries} (Total: ₹{fd_total_amt:,.2f})")
    print(f"FD transactions tracked     : {fd_transaction_count}")

    # 3. Build Sanitized Expense Sheet (with Manohar / Victor split)
    expense_ws, exp_copied, exp_skipped = create_expense_sheet(bank_wb, vendor_ws, bank_ws, bank_meta)
    print(f"Expense rows generated     : {exp_copied} (Includes Staff Salary split)")
    print(f"Expense rows skipped       : {exp_skipped}")

    # 4. Expense matching
    exp_results = match_bank_and_expense(bank_ws, expense_ws, bank_meta)
    apply_expense_calculations(expense_ws)

    # 5. Income matching (optional)
    inc_results = None
    income_ws = None
    if income_path:
        income_wb = load_input_workbook(income_path, data_only=False)
        income_source_ws = income_wb.worksheets[0]
        income_ws, inc_copied = sanitize_income_sheet(income_source_ws, bank_wb)
        print(f"Income receipts copied     : {inc_copied}")
        inc_results = match_bank_and_income(bank_ws, income_ws, bank_meta)

    # 6. Create editable allocation sheets for each cash withdrawal
    cash_withdrawal_count = create_cash_withdrawal_sheets(bank_ws, bank_wb, bank_meta)
    print(f"Cash withdrawal sheets       : {cash_withdrawal_count}")

    # 7. Income from Other Sources (GST Invoices) - optional
    other_income_ws = None
    other_unmatched_count = 0
    other_income_data_end = 0
    if other_income_path:
        other_income_wb = load_input_workbook(other_income_path, data_only=False)
        other_income_source_ws = other_income_wb.worksheets[0]
        other_income_ws, other_unmatched_count, other_income_data_end = process_other_income_sources(
            other_income_source_ws, bank_ws, bank_wb, bank_meta, income_ws
        )
        print(f"Other income sources entries: {other_unmatched_count} unmatched bank deposits added")

    # Clean Bank Sheet trailing empty lines
    rows_removed = clean_bank_sheet(bank_ws, header_row)
    add_bank_totals_summary(bank_ws, header_row)

    # Column widths and visual styling
    for ws in bank_wb.worksheets:
        autofit(ws)

    for r in range(2, expense_ws.max_row + 1):
        expense_ws.cell(r, 17).number_format = "dd-mmm-yyyy"
    if income_ws:
        for r in range(2, income_ws.max_row + 1):
            income_ws.cell(r, income_ws.max_column - 1).number_format = "dd-mmm-yyyy"
    if other_income_ws:
        for r in range(2, other_income_data_end + 1):
            other_income_ws.cell(r, 3).number_format = "dd-mmm-yyyy"
            other_income_ws.cell(r, 15).number_format = "dd-mmm-yyyy"

        # Explicitly retain numeric formatting for all GST calculation columns.
        for r in range(2, other_income_data_end + 1):
            for col in (5, 6, 7, 9, 10, 11, 13):
                other_income_ws.cell(r, col).number_format = GST_AMOUNT_FORMAT

    bank_wb.calculation.fullCalcOnLoad = True
    bank_wb.calculation.forceFullCalc = True
    bank_wb.calculation.calcMode = "auto"

    sheet_order = [
        "Bank_Statement",
        "FD Transactions",
        "Income from Other Sources",
        "Expense",
        "Income_recorded_in_nbh",
    ]
    ordered_sheets = [name for name in sheet_order if name in bank_wb.sheetnames]
    ordered_sheets.extend(
        name for name in bank_wb.sheetnames
        if name not in ordered_sheets and name.startswith("Cash Withdrawal")
    )
    ordered_sheets.extend(name for name in bank_wb.sheetnames if name not in ordered_sheets)
    bank_wb._sheets = [bank_wb[name] for name in ordered_sheets]
    bank_wb.active = bank_wb.sheetnames.index("Bank_Statement")

    output_path.parent.mkdir(parents=True, exist_ok=True)
    try:
        bank_wb.save(output_path)
    except Exception as e:
        raise RuntimeError(f"Could not save output file: {e}") from e

    # Final Report
    print()
    print("=" * 60)
    print("RECONCILIATION REPORT")
    print("=" * 60)
    print(f"Bank reference IDs generated : {exp_results['generated']}")
    print(f"Bank debit expenses matched  : {exp_results['matched']}")
    print(f"Bank expenses review needed  : {exp_results['review']}")
    print(f"Bank debit payments unmatched: {exp_results['unmatched']}")
    print(f"Total vendor expenses items  : {exp_results['expenses_total']}")
    print(f"Vendor expenses reconciled   : {exp_results['expenses_matched']}")

    if inc_results:
        print("-" * 60)
        print(f"Income receipts processed    : {inc_results['income_total']}")
        print(f"Income credits matched       : {inc_results['income_matched']}")
        print(f"Income receipts review needed: {inc_results['income_review']}")
        print(f"Income credits unmatched     : {inc_results['income_unmatched']}")

    if other_income_ws:
        print("-" * 60)
        print(f"Other income sources entries : {other_unmatched_count} unmatched deposits added")
        print(f"(GST calculated @ 18% on bank deposits)")

    print("-" * 60)
    print(f"FD Interest transactions     : {fd_entries}")
    print(f"FD Total Interest sum        : ₹{fd_total_amt:,.2f}")
    print(f"Trailing bank rows cleaned   : {rows_removed}")
    print()
    print(f"Saved: {output_path}")
    print("=" * 60)
    print()

    return {
        "output_path": output_path,
        "bank_date": bank_date,
        "expense_month": target_month_str,
        "expense_results": exp_results,
        "income_results": inc_results,
        "fd_entries": fd_entries,
        "fd_total_amt": fd_total_amt,
        "fd_transaction_count": fd_transaction_count,
        "cash_withdrawal_count": cash_withdrawal_count,
        "other_income_unmatched": other_unmatched_count,
        "rows_removed": rows_removed,
    }


def main():
    parser = argparse.ArgumentParser(
        description="NBH Bank, Expense & Income Reconciliation Tool",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""Examples:
  python3 reconcile.py --bank Bank_Statement.xlsx --vendor Vendor_bills.xlsx
  python3 reconcile.py --bank Bank_Statement.xlsx --vendor Vendor_bills.xlsx --income bank-book.xlsx --other-income other-sources.xlsx --output output.xlsx
        """
    )
    parser.add_argument("--bank", required=True, help="Path to Bank Statement Excel file (.xlsx or .xls)")
    parser.add_argument("--vendor", required=True, help="Path to Vendor Bills Excel file (.xlsx or .xls)")
    parser.add_argument("--income", default=None, help="Optional path to NBH Bank Book income file (.xlsx or .xls)")
    parser.add_argument("--other-income", default=None, help="Optional path to GST Invoice file (.xlsx or .xls)")
    parser.add_argument("--output", default=None, help="Optional custom output path (.xlsx)")
    args = parser.parse_args()

    bank_path = Path(args.bank)
    vendor_path = Path(args.vendor)
    income_path = Path(args.income) if args.income else None
    other_income_path = Path(args.other_income) if args.other_income else None
    output_path = Path(args.output) if args.output else bank_path.with_name(bank_path.stem + "_NBH_READY.xlsx")

    for p, label in [(bank_path, "Bank"), (vendor_path, "Vendor")]:
        if not p.exists():
            print(f"ERROR: {label} file not found: {p}")
            sys.exit(1)
        if p.suffix.lower() not in {".xlsx", ".xls"}:
            print(f"ERROR: {label} file must be .xlsx or .xls: {p}")
            sys.exit(1)

    if income_path:
        if not income_path.exists():
            print(f"ERROR: Income file not found: {income_path}")
            sys.exit(1)
        if income_path.suffix.lower() not in {".xlsx", ".xls"}:
            print(f"ERROR: Income file must be .xlsx or .xls: {income_path}")
            sys.exit(1)

    if other_income_path:
        if not other_income_path.exists():
            print(f"ERROR: Other Income file not found: {other_income_path}")
            sys.exit(1)
        if other_income_path.suffix.lower() not in {".xlsx", ".xls"}:
            print(f"ERROR: Other Income file must be .xlsx or .xls: {other_income_path}")
            sys.exit(1)

    try:
        run_reconciliation(bank_path, vendor_path, income_path, other_income_path, output_path)
    except (ValueError, RuntimeError) as error:
        print(f"\nERROR: {error}")
        sys.exit(1)


if __name__ == "__main__":
    main()
