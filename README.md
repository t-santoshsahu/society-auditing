# Society Auditing - Bank Reconciliation Suite

A comprehensive Python-based bank reconciliation tool for NoBrokerHood (NBH) society management. Automatically reconciles bank statements, vendor bills, and income receipts with intelligent matching and detailed reporting.

## Overview

This project provides a unified reconciliation system that:
- Matches bank transactions against vendor bills and expenses
- Reconciles income receipts with bank credit entries
- Automatically extracts and matches FD (Fixed Deposit) interest transactions
- Splits combined salary entries (e.g., "Staff Salary Victor & Manohar") into distinct line items
- Generates detailed Excel reports with color-coded matches and match status indicators
- Uses advanced text matching algorithms for robust vendor name matching

## Features

### Core Capabilities
- **Multi-file processing**: Accepts bank statements, vendor bills, and income receipts as Excel files
- **Intelligent matching**: Matches transactions using amount, bill numbers, dates, and vendor names
- **Automatic splitting**: Splits combined salary entries based on bank debit amounts
- **FD Interest extraction**: Auto-extracts "Income from FD Interest" sheet with totals
- **Reference ID generation**: Extracts and standardizes transaction references (UTR, Cheque, UPI, NEFT, etc.)
- **Status indicators**: Marks matches as MATCHED, REVIEW, or UNMATCHED
- **Detailed metrics**: Generates comprehensive reconciliation reports with match statistics

### Transaction Matching
- **Expense matching**: Matches vendor bills against bank debit transactions
- **Income matching**: Matches income receipts against bank credit transactions  
- **Multiple algorithms**: Uses amount matching, bill number matching, date matching, and vendor name similarity scoring
- **Salary splitting**: Automatically identifies and splits "Staff Salary Victor & Manohar" entries

### Output Sheets
- **Bank_Statement**: Enhanced with Generated Reference IDs, Bill Numbers, Descriptions, and Match Status
- **Expense**: Sanitized vendor expenses with payment references, bank dates, and match status
- **Income_recorded_in_nbh**: Processed income receipts with reconciliation details
- **Income from FD Interest**: Extracted FD interest transactions with totals and match status

## Usage

```bash
python3 reconcile.py "Bank_Statement.xlsx" "Vendor_bills_2023-24.xlsx" \
    --income "bank-book.xlsx" \
    --output "Bank_Reconcilation_READY.xlsx"
```

### Arguments
- **bank_file** (required): Path to Bank Statement Excel file
- **vendor_file** (required): Path to Vendor Bills Excel file  
- **--income** (optional): Path to NBH Bank Book income file
- **--output** (optional): Custom output path (default: bank_file_NBH_READY.xlsx)

## File Requirements

### Bank Statement Format
- Column headers: TRANSACTION DATE, NARRATION, CHEQUE NUMBER/SETTLEMENT ID, DEBIT, CREDIT
- Data starting from row 4 (or specified header row)
- Amounts in numeric format

### Vendor Bills Format
- Headers in row 1 with columns: SL.NO, Date, Billing Month, Bill No, Name, Amount, SGST, CGST, Total, TDS Rate, TDS Filed, Amount Paid, Payment Month, Less/Excess, Comments
- Data starting from row 2
- Must contain a sheet for the previous month from bank statement

### Income File Format
- Headers include: DATE, TYPE, FLAT NO., SETTLEMENT ID, REFERENCE NUMBER
- Type must be "BANK RECEIPT" or "BANK RECIEPT"
- Columns 10-11 contain amounts

## Output Report

The reconciliation generates a detailed report including:
- Bank reference IDs generated
- Expenses matched against bank debits
- Income matched against bank credits
- Items requiring manual review
- Unmatched items
- FD Interest total extracted

## Matching Algorithms

The tool uses a sophisticated multi-stage matching approach:
1. **Salary matching**: Special handling for identified salary payments
2. **Bill number matching**: Matches bill numbers found in bank narrations
3. **Amount + Date matching**: Precise match on amount and transaction date
4. **Similarity scoring**: Token-based vendor name matching with configurable thresholds
5. **Multiple candidate review**: Flags ambiguous matches requiring manual review

See [AGENTS.md](AGENTS.md) for detailed technical documentation on matching strategies.
