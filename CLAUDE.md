# QBO Invoice Matcher

Matches invoices from Gmail and local folders to QuickBooks Online transactions, then attaches them via browser automation.

## Workflow

1. **Prep**: Run email checker first to file stray invoices into Gmail label
2. **Export**: User downloads QBO transaction list as CSV (Reports > Transaction List > Export to Excel)
3. **Match**: Tool matches transactions → invoices (Gmail + local folder)
4. **Review**: User reviews match report, resolves uncertainties
5. **Attach**: Tool uploads confirmed matches to QBO via browser automation

## Config

- **Gmail Label**: accounting - Ltd Expenses - 24/25 Receipts
- **Local Receipts Folder**: /Users/matt/Desktop/Business Receipts copy
- **Tax Year**: August 2024 - July 2025
- **QBO URL**: https://qbo.intuit.co.uk

## Quick Start

Run `/transaction-audit` in a Claude session from this directory to start the full pipeline.

## Input Files

- **QBO Export**: `transactions.csv` — export from QBO (Reports > Transaction List by Date > Export to Excel)
  - Expected columns: Date, Transaction Type, Num, Name/Vendor, Memo/Description, Amount

## How to Run

```bash
# Step 1: Parse transactions + search for invoices
python3 qbo_matcher.py \
  --transactions path/to/transactions.csv \
  --gmail-label "accounting - Ltd Expenses - 24/25 Receipts" \
  --local-folder path/to/invoices/ \
  --output match_report.xlsx

# Step 2: Attach to QBO via browser (after review)
python3 qbo_matcher.py \
  --attach match_report.xlsx

# Dry-run attach (finds transactions but doesn't upload)
python3 qbo_matcher.py \
  --attach match_report.xlsx --dry-run
```

## Matching Logic

For each transaction, search for invoices by:
1. **Amount match** (strongest signal): Exact amount or within 1p tolerance
2. **Vendor/sender match**: Transaction vendor name ↔ email sender or PDF text
3. **Date proximity**: Invoice date within 14 days of transaction date

### Confidence Scoring
- **High** (80%+): Exact amount + vendor match + date within 7 days
- **Medium** (50-79%): Amount match + partial vendor or date match
- **Low** (<50%): Amount only or fuzzy matches

### Search Priority
1. Gmail label (attachments downloaded and cached)
2. Local folder (recursive PDF/image text extraction)
3. No match → flagged for manual action

## Output

Excel workbook:
- **Matched**: Transaction | Invoice Source | Confidence | Amount | Date | Vendor
- **Uncertain**: Transaction | Possible Matches | Confidence
- **Unmatched**: Transaction | Suggested Action (check platform X, contact vendor Y)
- **Summary**: Stats, match rates, action items
