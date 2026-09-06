# Matching Agents & Algorithms

This document describes the intelligent matching strategies and agents used in the bank reconciliation engine.

## Table of Contents
1. [Text Normalization](#text-normalization)
2. [Reference ID Extraction](#reference-id-extraction)
3. [Expense Matching Agent](#expense-matching-agent)
4. [Income Matching Agent](#income-matching-agent)
5. [FD Interest Extraction Agent](#fd-interest-extraction-agent)

---

## Text Normalization

### Core Text Processing Functions

#### `normalize_spaces(value)`
- Collapses multiple consecutive whitespace characters into single spaces
- Removes leading/trailing whitespace
- Purpose: Standardize string format for comparison

#### `normalize_text(value)`
- Converts text to uppercase
- Replaces all non-alphanumeric characters with spaces
- **Removes stop words** to reduce noise in matching:
  - Company suffixes: PVT, PRIVATE, LIMITED, LTD, LLP, COMPANY, CO
  - Common words: THE, AND, SERVICES, SERVICE, TECHNOLOGIES, TECHNOLOGY, INDIA, IN, BY
  - Banking terms: TRANSFER, NEFT, UPI, CR, DR

#### `get_tokens(value)`
- Tokenizes normalized text into individual words
- Returns as set for efficient intersection operations
- Example: "ABC Pvt Ltd" → {"ABC"}

### Similarity Scoring

#### `vendor_similarity(vendor_name, bank_text)`
Returns a float score (0.0 to 1.0) based on token overlap:
```
similarity = common_tokens / total_vendor_tokens
```

- **Score 0.0**: No matching tokens
- **Score 0.5+**: Moderate match (50%+ of vendor tokens found)
- **Score 0.8+**: Strong match (80%+ of vendor tokens found)

---

## Reference ID Extraction

The `generate_reference_id()` function automatically extracts standardized reference identifiers from transaction narrations, handling multiple bank and transfer types:

### 1. **UTR Numbers** (Universal Transaction Reference)
- Pattern: `UTR NO: XXXXX` or `UTR-XXXXX`
- Extracts: Bank NEFT reference numbers
- Priority: **HIGHEST** - checked first

### 2. **YES Bank NEFT**
- Pattern: `YESAPXXXXXXXX NNNN`
- Extracts: YES Bank NEFT identifier + sequence number
- Combination preserves unique identifier

### 3. **HSBC NEFT**
- Pattern: `HSBCNXXXXXXXX NNNN`
- Extracts: HSBC NEFT identifier + sequence number

### 4. **IDFC/IDFB NEFT**
- Pattern: `IDFBXXXXXXXX NNNN` or `IDFBXXXXXXXX`
- Extracts: IDFC NEFT reference (standalone or with sequence)

### 5. **ICICI NEFT**
- Pattern: `IN NNNNNNNN+` (IN followed by 8+ digits)
- Extracts: ICICI NEFT identifier

### 6. **UPI Transfers**
- Pattern: `UPI/CR/NNNN NNNN` or `UPI/DR/NNNN`
- Extracts: UPI transaction reference numbers

### 7. **Cheque Patterns** (Multiple formats)
- `CHQ-NNNN` or `CHQ NNNN`
- `CHEQUE WDL-CHEQUE TRANSFER TO-NNNN`
- `CHEQUE WDL-TRF-NNNN`
- `CLEARING-CHQ NNNN`
- `TO CLEARING - CHQ NNNN`
- `CASH CHEQUE...CHQ-NNNN`
- Extracts: Last 4+ digits of cheque number

### 8. **Fixed Deposit Transfer**
- Pattern: `TO FD-NNNN`
- Extracts: FD account identifier

### 9. **Internal/FD Interest Transfer**
- Pattern: `INT TRF FRM XXXXX` or `TRF FRM XXXXX`
- Extracts: Prepended with "INT-" prefix
- Used for income matching

### 10. **Fallback**
- Searches for 5-6 digit numbers in original reference column
- Returns first match or original reference string as-is

---

## Expense Matching Agent

Matches vendor bills against bank debit transactions using multi-stage intelligent algorithm.

### Input Data
- **Expense list**: Vendor bills with amounts, bill numbers, vendor names
- **Bank debits**: Transaction records with narrations, references, amounts

### Matching Process

#### Stage 1: Data Preparation
```python
expenses = [
    {
        "row": r,
        "bill_no": "vendor_bill_id",
        "vendor": "Vendor Name",
        "amount": 50000.00
    },
    ...
]
```

#### Stage 2: Salary Matching (Special Case)
**Purpose**: Handle "Staff Salary Victor & Manohar" splitting

For identified salary entries:
- **Manohar salary**: ₹16,600 (or value from bank record)
- **Victor salary**: ₹48,100 (or value from bank record)

Algorithm:
1. Check if bank narration contains "MANOHAR"
   - Search expense list for vendor containing "MANOHAR"
   - Match on amount equality
   - Mark as MATCHED with highest priority

2. Check if bank narration contains "VICTOR"
   - Search expense list for vendor containing "VICTOR"
   - Match on amount equality
   - Mark as MATCHED with highest priority

### Stage 3: Bill Number Matching

#### `bill_number_in_text(bill_no, bank_text)`
- Removes all non-alphanumeric characters from both strings
- Checks if cleaned bill number is substring of cleaned bank text
- **Precision**: Eliminates false matches from punctuation differences
- **Used when**: Vendor bill number appears in bank narration

**Match criteria**:
- Amount equals exactly
- Bill number found in bank narration/reference
- **Result**: MATCHED (high confidence)

### Stage 4: Amount + Date Matching

For remaining unmatched candidates:
- Amount matches (within ₹0.01 tolerance)
- Transaction date matches exactly
- **Result**: MATCHED if exactly one candidate, REVIEW if multiple

### Stage 5: Similarity-Based Matching

For vendors not matched by bill number:

#### Single Candidate by Amount
```python
candidates = expenses where amount matches AND not already matched
if len(candidates) == 1:
    status = "MATCHED"
elif len(candidates) > 1:
    status = "REVIEW"
```

#### Multiple Candidates - Scoring Algorithm
```python
scored_candidates = [
    (expense_index, vendor_similarity_score),
    ...
]
sorted by similarity score (descending)
```

**Score thresholds**:
1. **Top score ≥ 0.80 AND no score conflicts**: MATCHED
2. **Top score ≥ 0.50 AND score gap ≥ 0.15**:
   - If `top_score - second_score >= 0.15`: MATCHED
   - Ensures clear winner among candidates
3. **Otherwise**: REVIEW (manual intervention needed)

### Output Status Values

| Status | Meaning | Action Required |
|--------|---------|-----------------|
| MATCHED | Unique match found with high confidence | ✅ None - auto-accepted |
| REVIEW | Multiple candidates or weak match | ⚠️ Manual verification needed |
| UNMATCHED | No candidates found | ❌ Investigate discrepancy |
| NOT EXPENSE | Bank credit entry (income) | N/A |

### Matching Metrics

```python
metrics = {
    "generated": count of reference IDs extracted,
    "matched": count of expenses matched,
    "review": count of ambiguous matches,
    "unmatched": count of unmatched expenses,
    "expenses_total": total expense count,
    "expenses_matched": final matched count
}
```

---

## Income Matching Agent

Matches income receipts against bank credit transactions.

### Input Data
- **Income list**: Flat-wise collections with settlement IDs, reference numbers, flat numbers, amounts
- **Bank credits**: Unmatched credit transactions from bank statement

### Matching Process

#### Stage 1: Prepare Bank Credit List
```python
bank_credits = [
    {
        "row": r,
        "credit": amount,
        "date": transaction_date,
        "generated_ref": extracted_reference_id
    },
    ...
]
```

Filter: Credit amount > 0 AND not already matched as expense

#### Stage 2: Three-Level Matching Strategy

**Level 1: Settlement ID Matching** (Highest Priority)
```python
if income.settlement_id:
    for bank_credit in bank_credits:
        if bank_credit.generated_ref == income.settlement_id:
            match_found = True
            break
```
- Exact reference matching
- Most reliable method

**Level 2: Amount + Date Matching**
```python
if not found AND income.date:
    for bank_credit in bank_credits:
        if money_equal(bank_credit.credit, amount) 
           AND bank_credit.date == income.date:
            match_found = True
            break
```
- Precise amount and date match
- Secondary reliable method

**Level 3: Amount Matching Only**
```python
if not found:
    for bank_credit in bank_credits:
        if money_equal(bank_credit.credit, amount):
            candidates.append(bank_credit)
```
- Amount-based matching without date
- May result in multiple candidates

#### Stage 3: Decision Logic

```python
if len(matches) == 1:
    status = "MATCHED"
    mark bank row as matched with green fill
    populate bank description: "Flat # | Ref: #"
    
elif len(matches) > 1:
    status = "REVIEW"  # Multiple candidates
    
else:
    status = "UNMATCHED"  # No candidates
```

### Output Columns Added
- **Bank Match Status** (Column 9): MATCHED / REVIEW / UNMATCHED
- **Income Bank Reference ID**: Generated reference from bank
- **Income Bank Transaction Date**: Date from bank
- **Income Match Status**: MATCHED / REVIEW / UNMATCHED

### Income Metrics

```python
{
    "income_total": total income items,
    "income_matched": matched count,
    "income_review": ambiguous matches,
    "income_unmatched": no matches found
}
```

---

## FD Interest Extraction Agent

Automatically extracts and reconciles Fixed Deposit interest transactions from bank statements.

### Detection
**Narration pattern**: `INT TRF FRM XXXXX` or `TRF FRM XXXXX`
- Indicates interest transfer from FD account
- Extracted via reference ID generator with "INT-" prefix

### Processing Steps

#### Step 1: Locate FD Interest Rows
```python
for each bank_row:
    if "INT TRF FRM" in narration:
        extract as FD interest entry
```

#### Step 2: Extract Details
- Transaction date
- Narration text
- Original reference (cheque/settlement ID)
- Generated reference ID
- Credit amount (interest received)

#### Step 3: Create Income Sheet
**Sheet name**: "Income from FD Interest"

**Columns**:
1. Transaction Date
2. Narration
3. Cheque/Settlement Ref
4. Generated Reference ID
5. Credit / Interest Amount (INR)

#### Step 4: Highlight in Bank Statement
For each matched FD interest row in bank statement:
- **Generated Reference ID** (Col 6): Populated
- **Description** (Col 8): `FD Interest: [Reference]`
- **Match Status** (Col 9): `MATCHED`
- **Visual**: Filled with green background

#### Step 5: Calculate Totals
```python
total_interest = SUM(E2:E[last_row])
```
- Formula-based calculation
- Bold formatting with double border
- Labeled "Total Interest Received"

### Output Data
```python
{
    "entries_count": number of FD interest transactions,
    "total_interest": sum of all interest amounts
}
```

### Example

| Date | Narration | Ref | Generated ID | Interest |
|------|-----------|-----|--------------|----------|
| 2026-08-15 | INT TRF FRM 1001 | 1001 | INT-1001 | 1,234.50 |
| 2026-08-20 | INT TRF FRM 1002 | 1002 | INT-1002 | 2,345.75 |
| | | | **Total Interest Received** | **3,580.25** |

---

## Helper Functions

### Amount Processing

#### `numeric_amount(value)`
- Converts any value to float
- Handles: None, empty strings, formatted numbers with commas
- Strips currency symbols (₹)
- Regex extraction for partial numbers
- Returns: 0.0 if conversion fails

#### `money_equal(a, b, tolerance=0.01)`
- Compares amounts with tolerance
- Default: ₹0.01 tolerance
- Handles floating-point precision issues

### Date Processing

#### `parse_date(value)`
- Flexible date parsing
- Supported formats:
  - DD.MM.YYYY, DD-MM-YYYY, DD/MM/YYYY
  - DD.MM.YY, DD-MM-YY, DD/MM/YY
  - YYYY-MM-DD
  - DD-Mon-YYYY (e.g., 15-Aug-2026)
  - DD Mon YYYY / DD Month YYYY
- Returns: `date` object or None

#### `previous_month(year, month)`
- Calculates previous month for vendor expense matching
- Handles year boundaries (Jan → Dec previous year)

---

## Color Coding & Formatting

### Matched Row Styling
```python
GREEN_FILL = PatternFill(fill_type="solid", fgColor="C6EFCE")
MATCH_FONT = Font(color="006100")
```
- All matched rows: Green background
- Entire row columns 1-9 highlighted
- Applied to: Bank Statement, Expense, Income sheets

### Header Styling
```python
HEADER_FILL = PatternFill(fill_type="solid", fgColor="D9EAF7")
Font(bold=True)
Alignment(horizontal="center", vertical="center")
Border(bottom=THIN_GREY)
```
- Light blue background
- Bold text, centered
- Subtle bottom border

### Amount Formatting
```python
AMOUNT_FORMAT = '#,##0.00'
```
- Thousands separator (comma)
- Two decimal places
- Currency display ready

---

## Configuration Parameters

### Matching Thresholds
| Parameter | Value | Description |
|-----------|-------|-------------|
| Amount tolerance | ₹0.01 | Maximum diff for amount equality |
| Similarity threshold (strong) | 0.80 | Score for confident match |
| Similarity threshold (moderate) | 0.50 | Minimum for candidate consideration |
| Score gap (decision) | 0.15 | Minimum lead for top candidate |

### Salary Defaults
- Manohar: ₹16,600 (fallback if not found in bank)
- Victor: ₹48,100 (fallback if not found in bank)

### File Processing
- Bank header row: Auto-detected (rows 1-14)
- Vendor header row: 1
- Vendor data start: Row 2
- Month lookback: Previous month from bank statement

---

## Performance Notes

- **Text normalization**: O(n) where n = text length
- **Matching**: O(m×b) where m = expenses, b = bank debits
- **Similarity scoring**: O(m×b×t) where t = avg tokens per entry
- **Overall**: Suitable for monthly reconciliations (100s-1000s of entries)

---

## Future Enhancements

- [ ] Machine learning-based vendor matching
- [ ] Batch date fuzzy matching
- [ ] Multi-currency support
- [ ] Partial payment matching
- [ ] Split payment detection
- [ ] Automated REVIEW→MATCHED promotion with confidence scoring
