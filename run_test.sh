#!/bin/bash
# Test script for bank reconciliation

# Activate virtual environment if it exists
if [ -d "venv" ]; then
    source venv/bin/activate
fi

echo "Running Bank Reconciliation Test..."
echo "===================================="
echo ""

python3 reconcile.py \
    --bank test_data/Bank_Reconcilation_Aug_2026.xlsx \
    --vendor test_data/VendorBills.xlsx \
    --income test_data/83667_20260906_074014_bank-book.xlsx \
    --other-income test_data/NBH_GST_Invoice.xlsx \
    --output test_data/Reconciliation_Result.xlsx

RESULT=$?

echo ""
if [ $RESULT -eq 0 ]; then
    echo "✅ Test complete! Output saved to: test_data/Reconciliation_Result.xlsx"
else
    echo "❌ Test failed with exit code: $RESULT"
fi

exit $RESULT
