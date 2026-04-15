# Transaction Audit

Run the full QBO invoice matching pipeline: parse transactions, search Gmail and local receipts folder, produce a match report, and optionally attach invoices to QBO.

## Steps

### 1. Find the transaction export

Look for the most recent QBO transaction export in the current directory:
- `Seven Hills Search Limited_Transaction List by Date*.xlsx` (preferred)
- Any `.xlsx` or `.csv` file that looks like a QBO export

If no export is found, tell the user to export from QBO: **Reports > Transaction List by Date > set date range > Export to Excel**.

### 2. Check dependencies

Run `pip3 install -r requirements.txt` if any imports fail. Ensure `playwright install chromium` has been run.

### 3. Run the matching pipeline

```bash
python3 qbo_matcher.py \
  -t "<transaction_file>" \
  --local-folder "/Users/matt/Desktop/Business Receipts copy" \
  -o match_report.xlsx
```

This will:
- Parse all transactions from the export
- Search Gmail (label: `accounting - Ltd Expenses - 24/25 Receipts`) for matching invoices
- Search `/Users/matt/Desktop/Business Receipts copy` for local receipt PDFs/images
- Score each match by amount, vendor, and date proximity
- Output `match_report.xlsx`

### 4. Report progress

While the pipeline runs, report progress to the user every 30-60 seconds:
- Number of transactions parsed
- Gmail emails processed / total
- Local files processed / total
- Running match counts: High confidence / Needs review / Unmatched

### 5. Summarise results

When complete, read the output Excel and report:
- **Matched (high confidence)**: count and percentage — these are ready to attach
- **Needs review**: count — list the uncertain ones with their best match and confidence score
- **Unmatched**: count — list the vendor names so the user knows which receipts to find
- **Total value** of matched vs unmatched transactions

### 6. Ask about next steps

Ask the user:
1. **Review uncertain matches** — go through the "Needs Review" items together
2. **Attach to QBO** — run `python3 qbo_matcher.py --attach match_report.xlsx --dry-run` first, then the real attach
3. **Export only** — just keep the Excel report for manual processing

## Config

- **Gmail Label**: `accounting - Ltd Expenses - 24/25 Receipts`
- **Local Receipts**: `/Users/matt/Desktop/Business Receipts copy`
- **QBO Region**: UK (`https://qbo.intuit.co.uk`)
- **Tax Year**: August 2024 - July 2025
