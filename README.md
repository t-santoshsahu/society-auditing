# Automated NBH Society Audit Reconciliation

This project generates month-wise Excel reconciliation reports from a bank statement, NBH bank-book export, NBH GST income report, and vendor invoice sheet.

The main workflow is `auto_reconcile.py`. It creates or updates one report per month and preserves valid manual corrections when the month is processed again.

## Setup

Requirements:

- Python 3.9 or newer
- `openpyxl`
- `xlrd` for legacy `.xls` bank statements

```bash
python3 -m venv venv
source venv/bin/activate
pip install openpyxl xlrd
```

## Folder Layout

The financial year is the data root. Each month folder must contain an `input` folder:

```text
test_data/
└── 2026-2027/
    ├── VendorBills.xlsx
    ├── Jun_26/
    │   ├── input/
    │   │   ├── bank-statement.xls
    │   │   ├── nbh-bank-book.xlsx
    │   │   └── nbh-gst-income.xlsx
    │   └── report_June_2026.xlsx
    └── Jul_26/
        └── input/
```

The month directory name must start with a month and contain a two- or four-digit year, for example `Jun_26`, `June-2026`, or `August 2026`.

## Required Inputs

Each month requires these four inputs:

1. **Bank statement for the month**
   - Excel `.xlsx` or legacy `.xls` format.
   - Must contain columns equivalent to `TRANSACTION DATE`, `NARRATION`, `CHEQUE NUMBER/SETTLEMENT ID`, `DEBIT`, and `CREDIT`.
   - The script identifies it by its headers, not by its filename.

2. **NBH bank book export**
   - Must contain a bank-book heading and columns including `DATE`, `TYPE`, `FLAT NO.`, `PARTICULARS`, `TXN ID / CHEQUE NO.`, `REFERENCE NUMBER`, `SETTLEMENT ID`, `DEBIT`, and `CREDIT`.
   - Rows with type `BANK RECEIPT` or `BANK RECIEPT` are copied into `Income_recorded_in_nbh`.

3. **NBH GST income report**
   - Must contain GST invoice data, including `BILL NUMBER` and `LEDGER NAME`, or be identifiable as an invoice/GST report.
   - It supplies recorded income such as classes, amenities, move-in/move-out, sponsorship, services, and other sources.

4. **Vendor invoice sheet**
   - Usually a shared workbook named `VendorBills.xlsx` in the financial-year root.
   - It must contain a sheet for the month before the bank statement month.
   - Expected fields include vendor name, bill number, billing month, amount, GST, amount paid, payment month, and comments.

There must be exactly one bank statement, one NBH bank book, and one GST income report in each month’s `input` directory. Extra Excel files may cause classification errors. Use `--vendor-bills` when the vendor workbook is stored elsewhere or has a different name.

## Running the Process

Run one month:

```bash
source venv/bin/activate
python3 auto_reconcile.py test_data/2026-2027/Jun_26
```

Run all month folders under a financial-year root:

```bash
python3 auto_reconcile.py test_data/2026-2027
```

Specify the vendor workbook explicitly:

```bash
python3 auto_reconcile.py \
  test_data/2026-2027/Jun_26 \
  --vendor-bills test_data/2026-2027/VendorBills.xlsx
```

When processing a single month, the script searches sibling month folders for cross-month payments. Use `--data-root` when the month’s parent is not the financial-year root:

```bash
python3 auto_reconcile.py path/to/Jun_26 \
  --data-root path/to/2026-2027
```

The output is written directly inside the month folder as `report_<Month>_<Year>.xlsx`. A rerun creates a fresh result, then carries forward valid identifiers, statuses, descriptions, and manually completed cash-withdrawal entries.

## How Matching Works

### Bank Statement

The bank statement is the transaction source. The script generates a normalized reference ID from each narration and the original reference field. Recognized formats include UTR, NEFT, UPI, cheque, fixed-deposit, and internal-transfer references.

The bank row is colored green when it is matched to an expense, NBH income, other income, FD transaction, or populated cash-withdrawal register. It is colored yellow for a partial match and red for `REVIEW`, `UNMATCHED`, or `PENDING` states.

### Expense Sheet

Vendor expenses are matched to bank debit transactions in this order:

1. **Explicit payment reference**: a bill number, cheque/reference ID, or comma-separated references in the expense payment-reference field are matched to generated bank reference IDs. Multiple references may represent split payments.
2. **Amount**: the sum of matching bank debits is compared with the expected vendor payment. A difference of up to `1.00` is accepted as `MATCHED`.
3. **Unique amount fallback**: an unresolved expense is matched when exactly one unused bank debit has the same expected amount within `1.00`.
4. **Cross-month search**: unresolved expenses are searched in sibling month bank statements by reference, vendor/bill details, and split-cheque numbers. A full cross-month match is marked `MATCHED`; a partial split payment is marked `PARTIALLY MATCHED` or `REVIEW`.

Special salary rows for Manohar and Victor are handled separately when their names and amounts identify the bank payment.

#### If an Expense row is not matched

- Check that the vendor invoice is in the correct previous-month sheet.
- Enter the bank-generated reference ID or cheque references in the expense payment-reference field. For split payments, enter all references separated by commas.
- Correct the vendor name, bill number, amount, GST, or payment amount if the invoice data is wrong.
- Rerun the same month. The script will retry the reference and amount match and will search other months.
- Investigate `REVIEW` when more than one transaction could match. Choose the correct reference manually; do not force a match by changing an unrelated amount.

### `Income_recorded_in_nbh`

NBH bank-book receipt rows are matched to bank credits using these fields, in priority order:

1. `SETTLEMENT ID` and the bank statement’s generated reference ID.
2. `TXN ID / CHEQUE NO.`, `BANK REFERENCE ID`, or `REFERENCE NUMBER` when available.
3. Credit/debit amount and transaction date.
4. Amount alone when it identifies exactly one unused bank credit.

The report initially copies `SETTLEMENT ID` into `BANK REFERENCE ID`. If a user corrects that reference, the corrected value is preserved on the next run and is used for matching. A unique match marks the income row and bank row `MATCHED` and colors them green. Ambiguous or missing matches remain `REVIEW` or `UNMATCHED`.

#### If an NBH income row is not matched

- Compare `SETTLEMENT ID` in the NBH bank book with the generated reference ID shown in `Bank_Statement`.
- If the bank statement’s generated ID is wrong, correct the reference on the bank row or enter the correct bank reference in `BANK REFERENCE ID` in `Income_recorded_in_nbh`.
- Confirm that the credit amount and date are correct.
- If the same amount appears more than once, provide the exact settlement/reference ID; amount alone is intentionally treated as ambiguous.
- Rerun the month. A corrected reference is retained and retried. Cross-month matching is also attempted when the receipt belongs to another month’s bank statement.

### `Income from Other Sources`

Existing GST income rows are combined with unmatched bank deposits. Automatically added deposit rows start as `PENDING` until the user supplies identifying details.

To match an other-income row:

1. Enter or confirm the income ledger/description.
2. Select an `Income Type`.
3. Provide the bank reference when known, or use the bank deposit amount when it identifies exactly one credit.
4. The script links the row to the bank statement, fills the bank date/reference when blank, marks it `MATCHED`, and colors the row green.

#### If an other-income row is not matched

- Fill in both the ledger/description and income type. Rows missing either value remain `PENDING`.
- Correct the bank reference or bank deposit amount.
- Use the exact generated bank reference when multiple credits have the same amount.
- Rerun the month and verify the bank row description and status.

### `FD Transactions`

Narrations such as `INT TRF FRM ...` or `TRF FRM ...` are treated as FD interest transfers. The script extracts the FD reference, records the interest amount, and marks the corresponding bank row as matched.

If an FD transaction is missing or incorrect, check the narration and original reference in the bank statement, correct those source values, and rerun the month.

### Cash Withdrawal Sheets

One editable `Cash Withdrawal <reference>` sheet is created for each bank cash withdrawal. Enter invoice details, GST, and allocation amounts in the register. The bank withdrawal is linked only when the register has payment entries and the values reconcile.

If a withdrawal is not matched, complete the register, check the withdrawal reference in cell `E1`, and verify the allocated amounts and closing balance. Rerun the month; manually filled registers are preserved.

## Status and Color Guide

| Status | Meaning | User action |
| --- | --- | --- |
| `MATCHED` | One reliable transaction match was found. | No action unless the match is wrong. |
| `MATCHED (Cross-month...)` | The match was found in another month. | Verify the referenced month and row. |
| `PARTIALLY MATCHED` | Some amount or split payment was found, but the total differs. | Check references and payment amounts. |
| `REVIEW` | More than one possible match or a source mismatch exists. | Enter the exact reference and rerun. |
| `UNMATCHED` | No reliable match was found. | Correct source data or add the missing reference. |
| `PENDING` | Other-income details are incomplete. | Fill the ledger and income type. |
| `NOT EXPENSE` | The bank transaction is not a vendor expense. | Review it under income or other-income sheets. |

Green means matched, yellow means partial, and red means review, unmatched, or pending. Manual edits to identifiers, descriptions, statuses, and cash registers are carried forward on the next run, so corrections should be made in the generated month report and then verified after rerunning.

## Validation Before Closing a Month

1. Confirm every required input file was classified correctly.
2. Review red and yellow rows in `Bank_Statement`, `Expense`, `Income_recorded_in_nbh`, and `Income from Other Sources`.
3. Resolve every `REVIEW`, `UNMATCHED`, and `PENDING` row or document why it remains unresolved.
4. Check that manually entered bank references exactly match the bank statement generated reference IDs.
5. Rerun the month after corrections and confirm that the corrections remain and the affected rows are green.
