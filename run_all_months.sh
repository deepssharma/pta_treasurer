#!/bin/bash

# PTA Treasurer Report Generator - Batch Runner
# Runs the notebook for each month automatically

NOTEBOOK="PTA_Treasurer_Report_v4.ipynb"
LOG_DIR="logs"
mkdir -p $LOG_DIR

# Define months to process
# Format: "Month Year"
MONTHS=(
    "September 2025"
    "October 2025"
    "November 2025"
    "December 2025"
    "January 2026"
    "February 2026"
    "March 2026"
    "April 2026"
    "May 2026"
)

for MONTH_YEAR in "${MONTHS[@]}"; do
    MONTH=$(echo $MONTH_YEAR | cut -d' ' -f1)
    YEAR=$(echo $MONTH_YEAR | cut -d' ' -f2)
    
    echo "================================================"
    echo "Processing: $MONTH $YEAR"
    echo "================================================"
    
    # Check if input folder exists
    INPUT_DIR="input/${MONTH}_${YEAR}"
    if [ ! -d "$INPUT_DIR" ]; then
        echo "WARNING: $INPUT_DIR not found - skipping"
        continue
    fi
    
    # Check if QB file exists
    QB_FILE=$(ls $INPUT_DIR/quickbooks_*.csv 2>/dev/null | head -1)
    if [ -z "$QB_FILE" ]; then
        echo "WARNING: No QuickBooks file in $INPUT_DIR - skipping"
        continue
    fi
    
    # Set environment variables for the notebook
    export BATCH_MODE=1
    export INPUT_MONTH=$MONTH
    export FISCAL_YEAR=$YEAR
    
    # Run notebook with nbconvert
    jupyter nbconvert \
        --to notebook \
        --execute \
        --ExecutePreprocessor.timeout=300 \
        --output "output/executed_${MONTH}_${YEAR}.ipynb" \
        $NOTEBOOK \
        2>&1 | tee "$LOG_DIR/${MONTH}_${YEAR}.log"
    
    if [ $? -eq 0 ]; then
        echo "✅ $MONTH $YEAR - Done"
    else
        echo "❌ $MONTH $YEAR - Failed (check $LOG_DIR/${MONTH}_${YEAR}.log)"
    fi
    
    echo ""
done

echo "================================================"
echo "Batch processing complete"
echo "================================================"
