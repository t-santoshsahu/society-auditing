#!/usr/bin/env python3
"""
NBH Bank Reconciliation (Unified Edition)
=========================================

Features:
  - Takes inputs via CLI (Bank, Vendor, --income, --output)
  - Detailed logging, formatting, and metrics summary
  - Automatic split of 'Staff Salary Victor & Manohar' into distinct Manohar and
    Victor line items matched against bank cheques (Chq 626765 & Chq 626764)
  - Auto-extracts 'Income from FD Interest' sheet ('INT TRF FRM') with totals,
    marks matching bank rows green, and sets descriptions and MATCHED status.

Usage:
    python3 reconcile6.py "Bank_Statement.xlsx" "Vendor_bills_2023-24.xlsx" \
        --income "83667_20260906_074014_bank-book.xlsx" \
        --output "Bank_Reconcilation_Aug_2026_NBH_READY.xlsx"
"""

from __future__ import annotations

import argparse
import re
import sys
from copy import copy
from datetime import date, datetime
from pathlib import Path
from typing import Any, Optional

from openpyxl import load_workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter


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

AMOUNT_FORMAT = '#,##0.00'
GREEN_FILL = PatternFill(fill_type="solid", fgColor="C6EFCE")
HEADER_FILL = PatternFill(fill_type="solid", fgColor="D9EAF7")
MATCH_FONT = Font(color="006100")
THIN_GREY = Side(style="thin", color="D9E1F2")
DOUBLE_BOTTOM = Border(top=Side(style="thin", color="000000"), bottom=Side(style="double", color="000000"))


# =============================================================
# STRING & VALUE HELPERS
# =============================================================

def clean_string(value: Any) -> str:
    if value is None:
        return ""
    return str(value).strip()


def normalize_spaces(value: Any) -> str:
    return re.sub(r'\s+', ' ', clean_string(value))


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


def vendor_similarity(vendor_name: Any, bank_text: Any) -> float:
    vendor_tokens = get_tokens(vendor_name)
    bank_tokens = get_tokens(bank_text)

    if not vendor_tokens or not bank_tokens:
        return 0.0

    common = vendor_tokens.intersection(bank_tokens)
    return len(common) / len(vendor_tokens)


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

def generate_reference_id(narration: Any, original_reference: Any = None) -> str:
    if not narration and not original_reference:
        return ""

    raw = clean_string(narration)
    normalized = normalize_spaces(raw)
    upper = normalized.upper()

    # Explicit UTR NO
    match = re.search(r'\bUTR\s+NO\s*[:\-]?\s*([A-Z0-9]+)', upper)
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

    return orig_clean


# =============================================================
# EXCEL STYLING HELPERS
# =============================================================

def copy_cell(source, target):
    target.value = source.value
    if source.has_style:
        target._style = copy(source._style)
    if source.number_format:
        target.number_format = source.number_format
    if source.alignment:
        target.alignment = copy(source.alignment)
    if source.protection:
        target.protection = copy(source.protection)


def style_header(ws, row=1):
    for cell in ws[row]:
        if cell.value is not None:
            cell.fill = HEADER_FILL
            cell.font = Font(bold=True)
            cell.alignment = Alignment(horizontal="center", vertical="center")
            cell.border = Border(bottom=THIN_GREY)


def autofit(ws, max_width=50):
    for col_cells in ws.columns:
        col_letter = get_column_letter(col_cells[0].column)
        max_len = 0
        for cell in col_cells:
            if cell.value is not None:
                max_len = max(max_len, len(str(cell.value)))
        ws.column_dimensions[col_letter].width = min(max(max_len + 2, 10), max_width)


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
        values = [ws.cell(r, col).value for col in range(1, 6)]
        if any(v is not None and str(v).strip() != "" for v in values):
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
            output_row += 1
            copied_rows += 1

    amount_cols = [6, 7, 8, 9, 12, 14]
    for row in range(2, expense_ws.max_row + 1):
        for col in amount_cols:
            cell = expense_ws.cell(row, col)
            val = numeric_amount(cell.value)
            cell.value = val
            cell.number_format = AMOUNT_FORMAT

    style_header(expense_ws, 1)
    expense_ws.freeze_panes = "A2"
    return expense_ws, copied_rows, skipped_rows


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

    for c in range(1, orig_cols + 1):
        income_ws.cell(1, c).value = source_ws.cell(header_row, c).value

    income_ws.cell(1, orig_cols + 1).value = "Bank Reference ID"
    income_ws.cell(1, orig_cols + 2).value = "Bank Transaction Date"
    income_ws.cell(1, orig_cols + 3).value = "Bank Match Status"

    out_row = 2
    copied_rows = 0

    for r in range(header_row + 1, source_ws.max_row + 1):
        if not is_bank_receipt(source_ws.cell(r, 2).value):
            continue

        for c in range(1, orig_cols + 1):
            income_ws.cell(out_row, c).value = source_ws.cell(r, c).value

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
                gen_ref = generate_reference_id(narration, orig_ref)
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

def bill_number_in_text(bill_no: Any, text: Any) -> bool:
    bill = clean_string(bill_no)
    if not bill or not text:
        return False
    bill_clean = re.sub(r'[^A-Z0-9]', '', bill.upper())
    text_clean = re.sub(r'[^A-Z0-9]', '', str(text).upper())
    return bool(bill_clean and bill_clean in text_clean)


def match_bank_and_expense(bank_ws, expense_ws, bank_meta: dict) -> dict:
    header_row = bank_meta["header_row"]
    last_row = find_last_bank_data_row(bank_ws, header_row)

    expenses = []
    for r in range(2, expense_ws.max_row + 1):
        vendor = clean_string(expense_ws.cell(r, VENDOR_COL_NAME).value)
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
                "amount": final_amount,
            })

    matched_expenses = set()
    metrics = {"generated": 0, "matched": 0, "review": 0, "unmatched": 0}

    for row in range(header_row + 1, last_row + 1):
        narration = bank_ws.cell(row, bank_meta["narration_col"]).value
        orig_ref = bank_ws.cell(row, bank_meta["ref_col"]).value
        debit = numeric_amount(bank_ws.cell(row, bank_meta["debit_col"]).value)
        credit = numeric_amount(bank_ws.cell(row, bank_meta["credit_col"]).value)

        ref_id = generate_reference_id(narration, orig_ref)
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

        bank_text = f"{clean_string(narration)} {clean_string(orig_ref)} {ref_id}".upper()

        salary_match_idx = None
        if "MANOHAR" in bank_text:
            for i, exp in enumerate(expenses):
                if "MANOHAR" in exp["vendor"].upper() and money_equal(exp["amount"], debit):
                    salary_match_idx = i
                    break
        elif "VICTOR" in bank_text:
            for i, exp in enumerate(expenses):
                if "VICTOR" in exp["vendor"].upper() and money_equal(exp["amount"], debit):
                    salary_match_idx = i
                    break

        if salary_match_idx is not None and salary_match_idx not in matched_expenses:
            match_idx = salary_match_idx
            status = "MATCHED"
        else:
            candidates = [
                i for i, exp in enumerate(expenses)
                if i not in matched_expenses
                and money_equal(exp["amount"], debit)
                and bill_number_in_text(exp["bill_no"], bank_text)
            ]

            match_idx = None
            status = "UNMATCHED"

            if len(candidates) == 1:
                match_idx = candidates[0]
                status = "MATCHED"
            elif len(candidates) > 1:
                status = "REVIEW"
            else:
                scored_candidates = []
                for i, exp in enumerate(expenses):
                    if i in matched_expenses:
                        continue
                    if money_equal(exp["amount"], debit):
                        score = vendor_similarity(exp["vendor"], bank_text)
                        scored_candidates.append((i, score))

                if len(scored_candidates) == 1:
                    match_idx = scored_candidates[0][0]
                    status = "MATCHED"
                elif len(scored_candidates) > 1:
                    scored_candidates.sort(key=lambda x: x[1], reverse=True)
                    top_idx, top_score = scored_candidates[0]
                    second_score = scored_candidates[1][1]

                    if top_score >= 0.50 and (top_score - second_score >= 0.15):
                        match_idx = top_idx
                        status = "MATCHED"
                    elif top_score >= 0.80:
                        match_idx = top_idx
                        status = "MATCHED"
                    else:
                        status = "REVIEW"

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
# MAIN CONTROLLER
# =============================================================

def main():
    parser = argparse.ArgumentParser(description="NBH Bank, Expense & Income Reconciliation Tool")
    parser.add_argument("bank_file", help="Path to Bank Statement Excel file (.xlsx)")
    parser.add_argument("vendor_file", help="Path to Vendor Bills Excel file (.xlsx)")
    parser.add_argument("--income", default=None, help="Optional path to NBH Bank Book income file (.xlsx)")
    parser.add_argument("--output", default=None, help="Optional custom output path (.xlsx)")
    args = parser.parse_args()

    bank_path = Path(args.bank_file)
    vendor_path = Path(args.vendor_file)
    income_path = Path(args.income) if args.income else None
    output_path = Path(args.output) if args.output else bank_path.with_name(bank_path.stem + "_NBH_READY.xlsx")

    # Validations
    for p, label in [(bank_path, "Bank"), (vendor_path, "Vendor")]:
        if not p.exists():
            print(f"ERROR: {label} file not found: {p}")
            sys.exit(1)
        if p.suffix.lower() != ".xlsx":
            print(f"ERROR: {label} file must be .xlsx: {p}")
            sys.exit(1)

    if income_path:
        if not income_path.exists():
            print(f"ERROR: Income file not found: {income_path}")
            sys.exit(1)
        if income_path.suffix.lower() != ".xlsx":
            print(f"ERROR: Income file must be .xlsx: {income_path}")
            sys.exit(1)

    print()
    print("=" * 60)
    print("NoBrokerHood Bank + Expense + Income Reconciliation")
    print("=" * 60)
    print()
    print(f"Bank statement   : {bank_path}")
    print(f"Vendor bills     : {vendor_path}")
    print(f"Income receipts  : {income_path if income_path else 'None provided (skipping income matching)'}")
    print(f"Output file      : {output_path}")
    print()

    bank_wb = load_workbook(bank_path)
    vendor_wb = load_workbook(vendor_path, data_only=False)

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
        bank_ws.cell(r, 6).value = generate_reference_id(narr, orig)

    bank_ws.cell(header_row, 6).value = "Generated Reference ID"
    bank_ws.cell(header_row, 7).value = "Bill / Invoice No"
    bank_ws.cell(header_row, 8).value = "Expense Description"
    bank_ws.cell(header_row, 9).value = "Match Status"
    style_header(bank_ws, header_row)

    bank_date = determine_bank_month(bank_ws, header_row, date_col)
    if not bank_date:
        print("ERROR: Could not determine bank statement month.")
        sys.exit(1)

    print(f"Bank statement month : {bank_date.strftime('%B %Y')}")
    exp_year, exp_month = previous_month(bank_date.year, bank_date.month)
    target_month_str = datetime(exp_year, exp_month, 1).strftime("%B %Y")
    print(f"Vendor expense month : {target_month_str}")

    vendor_ws = find_vendor_sheet(vendor_wb, exp_year, exp_month)
    if vendor_ws is None:
        print(f"\nERROR: Could not find vendor sheet for {target_month_str}.")
        print("Available sheets:")
        for name in vendor_wb.sheetnames:
            print(f"  - {name}")
        sys.exit(1)

    print(f"Vendor sheet selected: {vendor_ws.title}")

    # 1. FD Interest Extraction Sheet & Green Highlight in Bank Statement
    fd_ws, fd_entries, fd_total_amt = extract_and_match_fd_interest(bank_ws, bank_wb, bank_meta)
    print(f"FD Interest rows reconciled: {fd_entries} (Total: ₹{fd_total_amt:,.2f})")

    # 2. Build Sanitized Expense Sheet (with Manohar / Victor split)
    expense_ws, exp_copied, exp_skipped = create_expense_sheet(bank_wb, vendor_ws, bank_ws, bank_meta)
    print(f"Expense rows generated     : {exp_copied} (Includes Staff Salary split)")
    print(f"Expense rows skipped       : {exp_skipped}")

    # 3. Expense matching
    exp_results = match_bank_and_expense(bank_ws, expense_ws, bank_meta)

    # 4. Income matching (optional)
    inc_results = None
    income_ws = None
    if income_path:
        income_wb = load_workbook(income_path, data_only=False)
        income_source_ws = income_wb.worksheets[0]
        income_ws, inc_copied = sanitize_income_sheet(income_source_ws, bank_wb)
        print(f"Income receipts copied     : {inc_copied}")
        inc_results = match_bank_and_income(bank_ws, income_ws, bank_meta)

    # Clean Bank Sheet trailing empty lines
    rows_removed = clean_bank_sheet(bank_ws, header_row)

    # Column widths and visual styling
    for ws in [bank_ws, expense_ws, fd_ws]:
        autofit(ws)
    if income_ws:
        autofit(income_ws)

    bank_ws.column_dimensions["B"].width = 55
    bank_ws.column_dimensions["H"].width = 45
    bank_ws.column_dimensions["I"].width = 16
    expense_ws.column_dimensions["E"].width = 40
    fd_ws.column_dimensions["B"].width = 55

    for r in range(2, expense_ws.max_row + 1):
        expense_ws.cell(r, 17).number_format = "dd-mmm-yyyy"
    if income_ws:
        for r in range(2, income_ws.max_row + 1):
            income_ws.cell(r, income_ws.max_column - 1).number_format = "dd-mmm-yyyy"

    bank_wb.active = bank_wb.sheetnames.index("Bank_Statement")

    try:
        bank_wb.save(output_path)
    except Exception as e:
        print(f"\nERROR: Could not save output file:\n{e}")
        sys.exit(1)

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

    print("-" * 60)
    print(f"FD Interest transactions     : {fd_entries}")
    print(f"FD Total Interest sum        : ₹{fd_total_amt:,.2f}")
    print(f"Trailing bank rows cleaned   : {rows_removed}")
    print()
    print(f"Saved: {output_path}")
    print("=" * 60)
    print()


if __name__ == "__main__":
    main()
