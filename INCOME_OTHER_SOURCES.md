# Income from Other Sources Sheet - Feature Documentation

## Overview

A new **"Income from Other Sources"** sheet has been added to the reconciliation tool to handle income from:
- Classes
- Amenity Booking
- Move In/Out
- Screen Rentals

This sheet processes GST invoices and automatically reconciles unmatched bank deposits with automatic GST calculations.

## Features

### 1. **GST Invoice Integration**
- Accepts GST invoice data from NBH (compatible format with existing GST reporting)
- Preserves all original invoice columns:
  - Bill Number
  - Ledger Name
  - Date
  - GST No
  - Place of Supply
  - TYPE (B2C, B2B, etc.)
  - UTILITY (YES/NO)
  - Voucher Type
  - Total Amount, Non Taxable Amount, Taxable Amount
  - GST % and calculated CGST/SGST/IGST

### 2. **Automatic Bank Deposit Matching**
- Scans bank statement for unmatched credit transactions
- Automatically identifies deposits not yet matched to:
  - Flat collections (Income_recorded_in_nbh)
  - Bank receipts
  - Expenses
  - FD Interest
- Adds unmatched deposits as new entries with status "PENDING"

### 3. **Automatic GST Calculation**
- Calculates 18% GST from bank deposit amounts (9% CGST + 9% SGST)
- Assumes bank amount is **inclusive of GST** (standard for India)
- Formula used:
  ```
  Taxable Amount = Bank Amount / 1.18
  GST Amount = Bank Amount - Taxable Amount
  CGST = GST Amount / 2
  SGST = GST Amount / 2
  ```

### 4. **Manual Income Type Entry**
- Provides "Income Type" column for later classification
- Allowed values: MoveInMoveOut, Classes, Sponsorship, Amenities Rentals, Services, Fund, Clubhouse Booking, Amenities Booking, EV, RFID, Others
- Supports free-form text entry for flexibility

### 5. **Bank Reconciliation Data**
Each unmatched deposit row includes:
- **Bank Deposit Amount**: Actual amount deposited
- **Bank Reference ID**: Generated reference from bank transaction
- **Bank Transaction Date**: Date from bank statement
- **Match Status**: "PENDING" for manual review

## Sheet Structure (20 Columns)

| Col | Name | Source | Notes |
|-----|------|--------|-------|
| 1-8 | Invoice Details | GST File | Bill No, Ledger Name, Date, GST No, Supply, TYPE, UTILITY, Voucher Type |
| 9 | Total Amount | GST File / Bank | Amount (inclusive of GST) |
| 10 | Non Taxable | GST File / Calc | Non-taxable portion |
| 11 | Taxable Amount | GST File / Calc | Taxable portion (Amount/1.18) |
| 12 | Percentage | GST File / Fixed | 9% for GST entries |
| 13 | CGST | GST File / Calc | Central GST (9%) |
| 14 | SGST | GST File / Calc | State GST (9%) |
| 15 | IGST | GST File | Inter-state GST (0% for B2C) |
| 16 | **Income Type** | **Manual** | **Classes / Amenity Booking / Move In/Out / Screen Rentals** |
| 17 | **Bank Deposit Amount** | Bank | Populated for unmatched deposits |
| 18 | **Bank Reference ID** | Bank | Generated reference from bank |
| 19 | **Bank Transaction Date** | Bank | Date from bank statement |
| 20 | **Match Status** | System | "PENDING" for unmatched deposits |

## Usage

### Run with Other Income File
```bash
python3 reconcile.py \
    --bank "Bank_Statement.xlsx" \
    --vendor "Vendor_bills.xlsx" \
    --income "bank-book.xlsx" \
    --other-income "NBH_GST_Invoice.xlsx" \
    --output "Reconciliation_Result.xlsx"
```

### Or use the test script
```bash
bash run_test.sh
```

## Test Results

From the test run with sample data:
- **GST Invoice entries**: 5 (from NBH_GST_Invoice.xlsx)
- **Unmatched bank deposits**: 19 (automatically added with GST calculated)
- **Total data rows**: 24

Example calculations for unmatched deposits:
- Bank Deposit: ₹900
  - Taxable: ₹762.71 (900/1.18)
  - GST (18%): ₹137.29
  - CGST (9%): ₹68.64
  - SGST (9%): ₹68.64

## Workflow

### Step 1: Prepare Data
- Gather GST invoices in NBH format (matching column structure)
- Ensure bank statement is processed first

### Step 2: Run Reconciliation
- Execute script with --other-income parameter
- Script automatically:
  - Copies existing GST invoices
  - Identifies unmatched bank deposits
  - Calculates GST
  - Creates "Income from Other Sources" sheet

### Step 3: Manual Review & Categorization
- Open output Excel file
- Review "Income from Other Sources" sheet
- Fill in "Income Type" column (Classes/Amenity/Move In/Out/Screen Rentals)
- Update Match Status when deposits are validated
- Add additional details as needed

### Step 4: Further Processing
- Export data for:
  - GST filing (use CGST/SGST columns)
  - Ledger posting
  - Income reporting
  - Collection tracking

## Benefits

✅ **Automated**: Unmatched deposits automatically added with calculations  
✅ **Accurate**: 18% GST calculated from bank amounts  
✅ **Organized**: Consolidated view of all "other income" in one place  
✅ **Flexible**: Income Type customizable for various revenue sources  
✅ **Auditable**: Bank references and dates preserved for verification  
✅ **Reconcilable**: PENDING status makes it clear which deposits need review  

## Future Enhancements

- [ ] Auto-categorization of income types based on narration patterns
- [ ] Batch status updates (PENDING → VERIFIED)
- [ ] Income trend analysis
- [ ] GST report generation
- [ ] Multi-period reconciliation support
- [ ] Integration with accounting software
