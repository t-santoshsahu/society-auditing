# Quick Reference - Test Commands

## Setup (One-Time)

Create and activate a Python virtual environment with required dependencies:

```bash
cd /Users/santosh/test/society-auditing
python3 -m venv venv
source venv/bin/activate
pip install openpyxl
```

## Run Test with All Files

**Using the shell script (recommended):**
```bash
bash run_test.sh
```

**Or directly:**
```bash
source venv/bin/activate
python3 reconcile.py \
    --bank test_data/Bank_Reconcilation_Aug_2026.xlsx \
    --vendor test_data/VendorBills.xlsx \
    --income test_data/83667_20260906_074014_bank-book.xlsx \
    --other-income test_data/NBH_GST_Invoice.xlsx \
    --output test_data/Reconciliation_Result.xlsx
```

### Test Output Validation

The generated Excel file opens without repair errors or warnings. The file contains the core reconciliation sheets plus one sheet per cash withdrawal:
- **Bank_Statement**: 107 rows with matching data
- **Income from FD Interest**: 12 rows with FD interest transactions
- **Expense**: 21 rows with matched vendor expenses
- **Income_recorded_in_nbh**: 50 rows with income reconciliation
- **Income from Other Sources**: GST invoices + unmatched bank deposits with 18% GST calculated
- **FD Transactions**: FD creation and FD breakage transfers, with the bank reference and debit/credit amount
- **Cash Withdrawal <reference>**: One editable payment allocation register for each cash withdrawal, including invoice, taxable amount, CGST, SGST, IGST, total GST, paid amount, and remaining cash

## Files Used in Test

| File | Purpose | Location |
|------|---------|----------|
| Bank_Reconcilation_Aug_2026.xlsx | Bank statement for August 2026 | test_data/ |
| VendorBills.xlsx | Vendor bills for July 2026 (previous month) | test_data/ |
| 83667_20260906_074014_bank-book.xlsx | Income receipts from bank book | test_data/ |
| NBH_GST_Invoice.xlsx | GST invoices for other income sources (Classes, Amenity, etc.) | test_data/ |
| Reconciliation_Result.xlsx | Output file (generated) | test_data/ |

## Expected Output Sheets

The reconciliation will generate the following sheets in the output file:

1. **Bank_Statement** - Enhanced with:
   - Generated Reference IDs
   - Bill/Invoice Numbers
   - Expense Descriptions
   - Match Status (MATCHED, REVIEW, or UNMATCHED)
   - Green highlighting for matched rows

2. **Expense** - Sanitized vendor expenses with:
   - All vendor bill details
   - Payment Reference IDs from bank matches
   - Bank transaction dates
   - Match status indicators
   - Special handling for "Staff Salary Victor & Manohar" split
   - GST totals at the end, split between Clubhouse invoices and Society Maintenance invoices

3. **Income_recorded_in_nbh** - Income reconciliation with:
   - Flat-wise collection details
   - Settlement IDs and reference numbers
   - Bank Reference IDs
   - Bank transaction dates
   - Match status indicators

4. **Income from FD Interest** - Fixed Deposit interest extraction:
   - Transaction dates
   - Narrations
   - Generated reference IDs
   - Interest amounts
   - Total interest calculation (formula-based)

5. **FD Transactions** - Fixed Deposit lifecycle tracking:
   - FD Created transactions such as `CHEQUE WDL-TRF TO FD-898569`
   - FD Broken transactions when proceeds return to the bank account
   - Bank reference, narration, debit investment, credit proceeds, and tracking status

6. **Income from Other Sources** - Classes, Amenity Booking, Move-In/Out, Screen Rentals:
   - GST invoice entries (from NBH_GST_Invoice.xlsx)
   - Unmatched bank deposits automatically added with PENDING status
   - **Columns include:**
     - Bill Number, Ledger Name, Date, GST No, Place of Supply
     - TYPE, UTILITY, Voucher Type
     - Total Amount, Non Taxable Amount, Taxable Amount
     - Percentage (GST %), CGST, SGST, IGST
   - **Income Type** (MoveInMoveOut / Classes / Sponsorship / Amenities Rentals / Services / Fund / Clubhouse Booking / Amenities Booking / EV / RFID / Others) - Manual entry
     - **Bank Deposit Amount** (populated from unmatched bank transactions)
     - **Bank Reference ID** (populated from bank statement)
     - **Bank Transaction Date** (populated from bank statement)
     - **Match Status** (PENDING for unmatched deposits)
   - **GST Calculation:** 18% GST (9% CGST + 9% SGST) calculated from bank deposit amounts
     - Formula: If bank amount is inclusive of GST, Taxable = Bank Amount / 1.18
     - CGST = (Bank Amount - Taxable) / 2
     - SGST = (Bank Amount - Taxable) / 2

7. **Cash Withdrawal <reference>** - One sheet for every cash withdrawal:
    - Withdrawal date, reference, amount, and original bank narration
    - Twenty payment rows for recording payee, description, invoice number, and invoice availability
    - Payment, taxable amount, CGST, SGST, IGST, and total GST columns
    - Automatic Total Paid, Balance Cash, and Total GST formulas

## Installation & Dependencies

The project uses **openpyxl** for Excel file manipulation. It's already included in the `venv` virtual environment setup above.

**Dependencies:**
- `openpyxl` >= 3.0 (for Excel file handling)
- Python 3.9+

**Note:** Due to Python version management (Homebrew), a virtual environment is recommended to avoid conflicts with system packages.
