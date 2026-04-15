#!/usr/bin/env python3
"""
QBO Invoice Matcher

Matches invoices from Gmail and local folders to QuickBooks Online transactions.
Outputs an Excel match report for review, then optionally attaches via browser.

Usage:
    # Match transactions to invoices
    python qbo_matcher.py --transactions txns.csv --output report.xlsx

    # Match with local folder too
    python qbo_matcher.py --transactions txns.csv --local-folder ~/invoices --output report.xlsx

    # Attach confirmed matches to QBO (after review)
    python qbo_matcher.py --attach report.xlsx
"""

import argparse
import csv
import io
import os
import re
import sys
import tempfile
from datetime import datetime, timedelta
from pathlib import Path
from typing import Optional

TOOL_DIR = Path(__file__).parent.resolve()


# ---------------------------------------------------------------------------
# Transaction parsing
# ---------------------------------------------------------------------------

def parse_transactions(filepath: str) -> list[dict]:
    """
    Parse QBO transaction export CSV/Excel into standardised dicts.

    Returns list of:
        {date, vendor, amount, description, type, num, raw_row}
    """
    path = Path(filepath)

    if path.suffix in ('.xlsx', '.xls'):
        return _parse_excel_transactions(path)
    else:
        return _parse_csv_transactions(path)


def _parse_csv_transactions(path: Path) -> list[dict]:
    """Parse CSV transaction export from QBO."""
    transactions = []

    with open(path, 'r', encoding='utf-8-sig') as f:
        # Try to detect delimiter
        sample = f.read(2048)
        f.seek(0)

        dialect = csv.Sniffer().sniff(sample)
        reader = csv.DictReader(f, dialect=dialect)

        # Normalise column names (QBO exports vary)
        fieldnames = reader.fieldnames
        col_map = _map_columns(fieldnames)

        for row in reader:
            txn = _extract_transaction(row, col_map)
            if txn:
                transactions.append(txn)

    return transactions


def _parse_excel_transactions(path: Path) -> list[dict]:
    """Parse Excel transaction export from QBO."""
    try:
        import openpyxl
    except ImportError:
        print("ERROR: openpyxl required for Excel files. Install: pip install openpyxl")
        sys.exit(1)

    wb = openpyxl.load_workbook(path, read_only=True)
    ws = wb.active
    transactions = []

    rows = list(ws.iter_rows(values_only=True))
    if not rows:
        return []

    # Find header row (skip any QBO metadata rows)
    header_idx = 0
    for i, row in enumerate(rows):
        row_strs = [str(c).lower() if c else '' for c in row]
        if any('date' in s for s in row_strs):
            header_idx = i
            break

    headers = [str(c).strip() if c else f'col_{i}' for i, c in enumerate(rows[header_idx])]
    col_map = _map_columns(headers)

    for row in rows[header_idx + 1:]:
        row_dict = {headers[i]: row[i] for i in range(min(len(headers), len(row)))}
        txn = _extract_transaction(row_dict, col_map)
        if txn:
            transactions.append(txn)

    wb.close()
    return transactions


def _map_columns(headers: list[str]) -> dict:
    """Map QBO column names to standard names. QBO exports vary."""
    col_map = {}
    normalised = {h.lower().strip(): h for h in headers if h}

    # Date
    for key in ['date', 'txn date', 'transaction date']:
        if key in normalised:
            col_map['date'] = normalised[key]
            break

    # Vendor/Name
    for key in ['name', 'vendor', 'payee', 'name/vendor', 'customer/vendor', 'from/to']:
        if key in normalised:
            col_map['vendor'] = normalised[key]
            break

    # Amount (single column)
    for key in ['amount', 'total', 'net amount']:
        if key in normalised:
            col_map['amount'] = normalised[key]
            break

    # Split amount columns (bank feed format: Spent / Received)
    for key in ['spent', 'debit', 'money out']:
        if key in normalised:
            col_map['spent'] = normalised[key]
            break
    for key in ['received', 'credit', 'money in']:
        if key in normalised:
            col_map['received'] = normalised[key]
            break

    # Description
    for key in ['memo/description', 'memo', 'description', 'notes',
                'bank description', 'transaction posted']:
        if key in normalised:
            col_map['description'] = normalised[key]
            break

    # Bank description (used as vendor fallback for bank feed exports)
    for key in ['bank description']:
        if key in normalised:
            col_map['bank_description'] = normalised[key]
            break

    # Transaction type
    for key in ['transaction type', 'type', 'txn type']:
        if key in normalised:
            col_map['type'] = normalised[key]
            break

    # Reference number
    for key in ['num', 'ref no.', 'reference', 'doc number', 'ref no']:
        if key in normalised:
            col_map['num'] = normalised[key]
            break

    return col_map


def _extract_transaction(row: dict, col_map: dict) -> Optional[dict]:
    """Extract a standardised transaction from a row."""
    date_str = row.get(col_map.get('date', ''), '')
    if not date_str:
        return None

    # Parse date (QBO uses various formats)
    txn_date = _parse_date(str(date_str))
    if not txn_date:
        return None

    # Parse amount — handle single column or split Spent/Received columns
    amount = None
    if 'amount' in col_map:
        amount_str = str(row.get(col_map['amount'], '0'))
        amount = _parse_amount(amount_str)
    elif 'spent' in col_map or 'received' in col_map:
        spent_str = str(row.get(col_map.get('spent', ''), '') or '')
        received_str = str(row.get(col_map.get('received', ''), '') or '')
        spent = _parse_amount(spent_str) if spent_str.strip() else None
        received = _parse_amount(received_str) if received_str.strip() else None
        if spent:
            amount = -abs(spent)  # Outgoing = negative
        elif received:
            amount = abs(received)  # Incoming = positive

    if amount is None or amount == 0:
        return None

    # Vendor: prefer From/To, fall back to Bank description
    vendor = str(row.get(col_map.get('vendor', ''), '')).strip()
    if not vendor and 'bank_description' in col_map:
        vendor = str(row.get(col_map['bank_description'], '')).strip()

    description = str(row.get(col_map.get('description', ''), '')).strip()
    txn_type = str(row.get(col_map.get('type', ''), '')).strip()
    num = str(row.get(col_map.get('num', ''), '')).strip()

    return {
        'date': txn_date,
        'vendor': vendor,
        'amount': abs(amount),
        'amount_raw': amount,
        'description': description,
        'type': txn_type,
        'num': num,
        'raw_row': dict(row),
    }


def _parse_date(s: str) -> Optional[datetime]:
    """Try multiple date formats."""
    s = s.strip()
    for fmt in ['%d/%m/%Y', '%m/%d/%Y', '%Y-%m-%d', '%d-%m-%Y',
                '%d %b %Y', '%d %B %Y', '%m/%d/%y', '%d/%m/%y']:
        try:
            return datetime.strptime(s, fmt)
        except ValueError:
            continue

    # Handle datetime objects (from Excel)
    if isinstance(s, datetime):
        return s

    return None


def _parse_amount(s: str) -> Optional[float]:
    """Parse amount string, handling currency symbols and negatives."""
    if not s:
        return None
    # Remove currency symbols and whitespace
    cleaned = re.sub(r'[£$€,\s]', '', s)
    # Handle parentheses for negatives: (123.45) → -123.45
    if cleaned.startswith('(') and cleaned.endswith(')'):
        cleaned = '-' + cleaned[1:-1]
    try:
        return float(cleaned)
    except ValueError:
        return None


# ---------------------------------------------------------------------------
# Gmail invoice search
# ---------------------------------------------------------------------------

def search_gmail_invoices(label_name: str, transactions: list[dict],
                          cache_dir: Path) -> dict[int, list[dict]]:
    """
    Search Gmail label for invoices matching transactions.

    Returns: {txn_index: [list of candidate matches]}
    """
    from gmail_oauth import get_gmail_credentials
    from googleapiclient.discovery import build
    import base64

    creds = get_gmail_credentials()
    service = build('gmail', 'v1', credentials=creds)

    # Find the label ID
    labels = service.users().labels().list(userId='me').execute()
    label_id = None
    for label in labels.get('labels', []):
        if label['name'].lower() == label_name.lower():
            label_id = label['id']
            break

    if not label_id:
        # Try nested label path
        for label in labels.get('labels', []):
            if label['name'].lower().replace('/', ' - ').strip() == label_name.lower():
                label_id = label['id']
                break

    if not label_id:
        print(f"WARNING: Gmail label '{label_name}' not found")
        print("Available labels containing 'accounting' or 'expense' or 'receipt':")
        for label in labels.get('labels', []):
            name_lower = label['name'].lower()
            if any(k in name_lower for k in ['accounting', 'expense', 'receipt', 'invoice']):
                print(f"  - {label['name']}")
        return {}

    print(f"Found Gmail label: {label_name} (ID: {label_id})")

    # Fetch all messages in the label
    all_messages = []
    page_token = None

    while True:
        kwargs = {
            'userId': 'me',
            'labelIds': [label_id],
            'maxResults': 100,
        }
        if page_token:
            kwargs['pageToken'] = page_token

        results = service.users().messages().list(**kwargs).execute()
        messages = results.get('messages', [])
        all_messages.extend(messages)

        page_token = results.get('nextPageToken')
        if not page_token:
            break

    print(f"Found {len(all_messages)} emails in label")

    # Parse each email for invoice data
    invoices = []
    for i, msg_ref in enumerate(all_messages):
        if i % 20 == 0:
            print(f"  Processing email {i+1}/{len(all_messages)}...")

        msg = service.users().messages().get(
            userId='me', id=msg_ref['id'], format='full'
        ).execute()

        invoice = _parse_email_invoice(msg, service, cache_dir)
        if invoice:
            invoices.append(invoice)

    print(f"Extracted {len(invoices)} invoices from emails")

    # Match invoices to transactions
    matches = {}
    for txn_idx, txn in enumerate(transactions):
        candidates = []
        for inv in invoices:
            score = _score_match(txn, inv)
            if score > 0:
                candidates.append({**inv, 'score': score})

        if candidates:
            candidates.sort(key=lambda x: x['score'], reverse=True)
            matches[txn_idx] = candidates[:3]  # Top 3

    return matches


def _parse_email_invoice(msg: dict, service, cache_dir: Path) -> Optional[dict]:
    """Extract invoice data from a Gmail message."""
    headers = {h['name'].lower(): h['value']
               for h in msg.get('payload', {}).get('headers', [])}

    sender = headers.get('from', '')
    subject = headers.get('subject', '')
    date_str = headers.get('date', '')

    # Parse email date
    email_date = _parse_email_date(date_str)

    # Extract amounts from subject + body
    body = _get_email_body(msg)
    amounts = _extract_amounts(subject + ' ' + body)

    # Extract sender name/company
    sender_name = _extract_sender_name(sender)

    # Check for attachments
    attachments = _get_attachments_info(msg)
    attachment_path = None

    # Download first PDF/image attachment if present
    if attachments:
        for att in attachments:
            if att['filename'].lower().endswith(('.pdf', '.png', '.jpg', '.jpeg')):
                att_path = cache_dir / att['filename']
                if not att_path.exists():
                    _download_attachment(service, msg['id'], att['id'], att_path)
                attachment_path = str(att_path)

                # Extract amounts from PDF too
                if att_path.suffix.lower() == '.pdf':
                    pdf_text = _extract_pdf_text(att_path)
                    if pdf_text:
                        pdf_amounts = _extract_amounts(pdf_text)
                        amounts.extend(pdf_amounts)
                break

    if not amounts and not attachment_path:
        return None

    return {
        'source': 'gmail',
        'message_id': msg['id'],
        'sender': sender,
        'sender_name': sender_name,
        'subject': subject,
        'date': email_date,
        'amounts': list(set(amounts)),
        'attachment_path': attachment_path,
        'attachment_filename': attachments[0]['filename'] if attachments else None,
    }


def _parse_email_date(date_str: str) -> Optional[datetime]:
    """Parse email Date header."""
    if not date_str:
        return None
    # Remove timezone name in parens
    date_str = re.sub(r'\s*\([^)]+\)\s*$', '', date_str).strip()
    for fmt in ['%a, %d %b %Y %H:%M:%S %z', '%d %b %Y %H:%M:%S %z',
                '%a, %d %b %Y %H:%M:%S', '%d %b %Y %H:%M:%S']:
        try:
            return datetime.strptime(date_str, fmt).replace(tzinfo=None)
        except ValueError:
            continue
    return None


def _get_email_body(msg: dict) -> str:
    """Extract plain text body from email."""
    import base64

    payload = msg.get('payload', {})

    # Simple body
    if payload.get('body', {}).get('data'):
        return base64.urlsafe_b64decode(payload['body']['data']).decode('utf-8', errors='replace')

    # Multipart
    for part in payload.get('parts', []):
        if part.get('mimeType') == 'text/plain' and part.get('body', {}).get('data'):
            return base64.urlsafe_b64decode(part['body']['data']).decode('utf-8', errors='replace')
        # Nested multipart
        for subpart in part.get('parts', []):
            if subpart.get('mimeType') == 'text/plain' and subpart.get('body', {}).get('data'):
                return base64.urlsafe_b64decode(subpart['body']['data']).decode('utf-8', errors='replace')

    return ''


def _extract_amounts(text: str) -> list[float]:
    """Extract monetary amounts from text."""
    amounts = []
    # Match patterns like £123.45, $1,234.56, 123.45, etc.
    patterns = [
        r'[£$€]\s*(\d{1,3}(?:,\d{3})*(?:\.\d{2})?)',  # £123.45
        r'(\d{1,3}(?:,\d{3})*\.\d{2})\s*(?:GBP|USD|EUR)',  # 123.45 GBP
        r'(?:total|amount|due|subtotal|net|gross|balance)[:\s]*[£$€]?\s*(\d{1,3}(?:,\d{3})*\.\d{2})',  # total: £123.45
    ]

    for pattern in patterns:
        for match in re.finditer(pattern, text, re.IGNORECASE):
            amount_str = match.group(1).replace(',', '')
            try:
                amount = float(amount_str)
                if 0.01 < amount < 1_000_000:  # Sanity check
                    amounts.append(amount)
            except ValueError:
                pass

    return amounts


def _extract_sender_name(sender: str) -> str:
    """Extract company/person name from email sender."""
    # "Company Name <email@example.com>" → "Company Name"
    match = re.match(r'"?([^"<]+)"?\s*<', sender)
    if match:
        return match.group(1).strip()
    # Just email: extract domain
    match = re.search(r'@([^.]+)', sender)
    if match:
        return match.group(1).capitalize()
    return sender


def _get_attachments_info(msg: dict) -> list[dict]:
    """Get attachment metadata from email."""
    attachments = []
    payload = msg.get('payload', {})

    for part in payload.get('parts', []):
        filename = part.get('filename', '')
        if filename and part.get('body', {}).get('attachmentId'):
            attachments.append({
                'filename': filename,
                'id': part['body']['attachmentId'],
                'size': part['body'].get('size', 0),
            })
        # Check nested parts
        for subpart in part.get('parts', []):
            filename = subpart.get('filename', '')
            if filename and subpart.get('body', {}).get('attachmentId'):
                attachments.append({
                    'filename': filename,
                    'id': subpart['body']['attachmentId'],
                    'size': subpart['body'].get('size', 0),
                })

    return attachments


def _download_attachment(service, message_id: str, attachment_id: str, save_path: Path):
    """Download a Gmail attachment to disk."""
    import base64
    att = service.users().messages().attachments().get(
        userId='me', messageId=message_id, id=attachment_id
    ).execute()

    data = base64.urlsafe_b64decode(att['data'])
    save_path.parent.mkdir(parents=True, exist_ok=True)
    with open(save_path, 'wb') as f:
        f.write(data)


# ---------------------------------------------------------------------------
# Local file search
# ---------------------------------------------------------------------------

def search_local_invoices(folder: str, transactions: list[dict]) -> dict[int, list[dict]]:
    """
    Search local folder for invoice files matching transactions.

    Returns: {txn_index: [list of candidate matches]}
    """
    folder_path = Path(folder)
    if not folder_path.exists():
        print(f"WARNING: Local folder not found: {folder}")
        return {}

    # Collect all PDF/image files
    invoice_files = []
    for ext in ['*.pdf', '*.PDF', '*.png', '*.PNG', '*.jpg', '*.JPG', '*.jpeg', '*.JPEG']:
        invoice_files.extend(folder_path.rglob(ext))

    print(f"Found {len(invoice_files)} invoice files in {folder}")

    # Parse each file
    invoices = []
    for i, filepath in enumerate(invoice_files):
        if i % 10 == 0 and i > 0:
            print(f"  Processing file {i+1}/{len(invoice_files)}...")

        invoice = _parse_local_invoice(filepath)
        if invoice:
            invoices.append(invoice)

    print(f"Extracted data from {len(invoices)} local files")

    # Match to transactions
    matches = {}
    for txn_idx, txn in enumerate(transactions):
        candidates = []
        for inv in invoices:
            score = _score_match(txn, inv)
            if score > 0:
                candidates.append({**inv, 'score': score})

        if candidates:
            candidates.sort(key=lambda x: x['score'], reverse=True)
            matches[txn_idx] = candidates[:3]

    return matches


def _parse_local_invoice(filepath: Path) -> Optional[dict]:
    """Extract invoice data from a local file."""
    text = ''

    if filepath.suffix.lower() == '.pdf':
        text = _extract_pdf_text(filepath)
    elif filepath.suffix.lower() in ('.png', '.jpg', '.jpeg'):
        # Could add OCR here in future
        pass

    if not text:
        # Fall back to filename analysis
        amounts = _extract_amounts(filepath.stem.replace('-', ' ').replace('_', ' '))
        if not amounts:
            return None
        return {
            'source': 'local',
            'filepath': str(filepath),
            'filename': filepath.name,
            'sender_name': _guess_vendor_from_filename(filepath.name),
            'date': _guess_date_from_file(filepath),
            'amounts': amounts,
            'attachment_path': str(filepath),
        }

    amounts = _extract_amounts(text)
    if not amounts:
        return None

    # Try to extract vendor from text
    vendor = _extract_vendor_from_text(text)
    # Try to extract date from text
    invoice_date = _extract_date_from_text(text)

    return {
        'source': 'local',
        'filepath': str(filepath),
        'filename': filepath.name,
        'sender_name': vendor or _guess_vendor_from_filename(filepath.name),
        'date': invoice_date or _guess_date_from_file(filepath),
        'amounts': amounts,
        'attachment_path': str(filepath),
    }


def _extract_pdf_text(filepath: Path) -> str:
    """Extract text from PDF."""
    try:
        import pdfplumber
        with pdfplumber.open(filepath) as pdf:
            text = ''
            for page in pdf.pages[:3]:  # First 3 pages max
                page_text = page.extract_text()
                if page_text:
                    text += page_text + '\n'
            return text
    except ImportError:
        pass

    try:
        from PyPDF2 import PdfReader
        reader = PdfReader(filepath)
        text = ''
        for page in reader.pages[:3]:
            page_text = page.extract_text()
            if page_text:
                text += page_text + '\n'
        return text
    except ImportError:
        pass

    print("WARNING: No PDF library available. Install: pip install pdfplumber")
    return ''


def _extract_vendor_from_text(text: str) -> Optional[str]:
    """Try to extract vendor/company name from invoice text."""
    # Look for common patterns
    patterns = [
        r'(?:from|invoice from|billed by|company)[:\s]+([A-Z][A-Za-z\s&]+(?:Ltd|Limited|Inc|LLC|LLP)?)',
        r'^([A-Z][A-Z\s&]{2,}(?:LTD|LIMITED|INC|LLC)?)\s*$',  # All-caps company at start
    ]
    for pattern in patterns:
        match = re.search(pattern, text[:500], re.MULTILINE | re.IGNORECASE)
        if match:
            return match.group(1).strip()
    return None


def _extract_date_from_text(text: str) -> Optional[datetime]:
    """Extract date from invoice text."""
    patterns = [
        r'(?:date|invoice date|issued)[:\s]*(\d{1,2}[/\-]\d{1,2}[/\-]\d{2,4})',
        r'(\d{1,2}\s+(?:Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)\w*\s+\d{4})',
    ]
    for pattern in patterns:
        match = re.search(pattern, text[:1000], re.IGNORECASE)
        if match:
            return _parse_date(match.group(1))
    return None


def _guess_vendor_from_filename(filename: str) -> str:
    """Guess vendor from filename."""
    # Remove extension, dates, numbers
    name = Path(filename).stem
    name = re.sub(r'\d{4}[-_]\d{2}[-_]\d{2}', '', name)
    name = re.sub(r'[-_]', ' ', name)
    name = re.sub(r'\d+', '', name).strip()
    return name if name else 'Unknown'


def _guess_date_from_file(filepath: Path) -> Optional[datetime]:
    """Guess date from file modification time or filename."""
    # Try filename first
    match = re.search(r'(\d{4})[-_](\d{2})[-_](\d{2})', filepath.name)
    if match:
        try:
            return datetime(int(match.group(1)), int(match.group(2)), int(match.group(3)))
        except ValueError:
            pass

    # Fall back to file modification time
    mtime = os.path.getmtime(filepath)
    return datetime.fromtimestamp(mtime)


# ---------------------------------------------------------------------------
# Matching engine
# ---------------------------------------------------------------------------

def _score_match(transaction: dict, invoice: dict) -> float:
    """
    Score how well an invoice matches a transaction.
    Returns 0-100 confidence score.
    """
    score = 0.0
    txn_amount = transaction['amount']

    # Amount match (strongest signal) — up to 50 points
    amount_matched = False
    for inv_amount in invoice.get('amounts', []):
        diff = abs(txn_amount - inv_amount)
        if diff < 0.02:  # Exact match (within 1p)
            score += 50
            amount_matched = True
            break
        elif diff < 1.0:  # Within £1
            score += 40
            amount_matched = True
            break
        elif diff / max(txn_amount, 0.01) < 0.05:  # Within 5%
            score += 25
            amount_matched = True
            break

    if not amount_matched:
        return 0  # No amount match = no match

    # Vendor/sender match — up to 30 points
    txn_vendor = transaction.get('vendor', '').lower()
    inv_sender = invoice.get('sender_name', '').lower()

    if txn_vendor and inv_sender:
        # Check if vendor name appears in sender or vice versa
        txn_words = set(re.findall(r'\w+', txn_vendor))
        inv_words = set(re.findall(r'\w+', inv_sender))

        # Remove common words
        stop_words = {'ltd', 'limited', 'inc', 'llc', 'the', 'and', 'of', 'uk', 'com'}
        txn_words -= stop_words
        inv_words -= stop_words

        if txn_words and inv_words:
            overlap = txn_words & inv_words
            if overlap:
                coverage = len(overlap) / min(len(txn_words), len(inv_words))
                score += 30 * coverage
            elif txn_vendor in inv_sender or inv_sender in txn_vendor:
                score += 20

    # Date proximity — up to 20 points
    txn_date = transaction.get('date')
    inv_date = invoice.get('date')

    if txn_date and inv_date:
        days_diff = abs((txn_date - inv_date).days)
        if days_diff <= 3:
            score += 20
        elif days_diff <= 7:
            score += 15
        elif days_diff <= 14:
            score += 10
        elif days_diff <= 30:
            score += 5

    return min(score, 100)


# ---------------------------------------------------------------------------
# Output
# ---------------------------------------------------------------------------

def write_match_report(transactions: list[dict],
                       gmail_matches: dict[int, list[dict]],
                       local_matches: dict[int, list[dict]],
                       output_path: str):
    """Write Excel match report."""
    try:
        import openpyxl
        from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
    except ImportError:
        print("ERROR: openpyxl required. Install: pip install openpyxl")
        sys.exit(1)

    wb = openpyxl.Workbook()

    # Merge matches from both sources
    all_matches = {}
    for txn_idx in range(len(transactions)):
        candidates = []
        if txn_idx in gmail_matches:
            candidates.extend(gmail_matches[txn_idx])
        if txn_idx in local_matches:
            candidates.extend(local_matches[txn_idx])
        if candidates:
            # Deduplicate and sort by score
            candidates.sort(key=lambda x: x['score'], reverse=True)
            all_matches[txn_idx] = candidates

    # Categorise
    matched = []      # High confidence (80+)
    uncertain = []    # Medium confidence (40-79)
    unmatched = []    # No match or low confidence (<40)

    for txn_idx, txn in enumerate(transactions):
        if txn_idx in all_matches:
            best = all_matches[txn_idx][0]
            if best['score'] >= 80:
                matched.append((txn_idx, txn, all_matches[txn_idx]))
            elif best['score'] >= 40:
                uncertain.append((txn_idx, txn, all_matches[txn_idx]))
            else:
                unmatched.append((txn_idx, txn))
        else:
            unmatched.append((txn_idx, txn))

    # Styles
    header_font = Font(bold=True, color='FFFFFF')
    header_fill = PatternFill(start_color='333333', end_color='333333', fill_type='solid')
    green_fill = PatternFill(start_color='C6EFCE', end_color='C6EFCE', fill_type='solid')
    yellow_fill = PatternFill(start_color='FFEB9C', end_color='FFEB9C', fill_type='solid')
    red_fill = PatternFill(start_color='FFC7CE', end_color='FFC7CE', fill_type='solid')

    def write_header(ws, headers):
        for col, h in enumerate(headers, 1):
            cell = ws.cell(row=1, column=col, value=h)
            cell.font = header_font
            cell.fill = header_fill
            cell.alignment = Alignment(horizontal='center')
        ws.freeze_panes = 'A2'

    # --- Summary tab ---
    ws_summary = wb.active
    ws_summary.title = 'Summary'
    write_header(ws_summary, ['Metric', 'Count', '%'])

    total = len(transactions)
    rows = [
        ('Total Transactions', total, ''),
        ('Matched (High Confidence)', len(matched), f'{len(matched)/total*100:.0f}%' if total else ''),
        ('Uncertain (Needs Review)', len(uncertain), f'{len(uncertain)/total*100:.0f}%' if total else ''),
        ('Unmatched (No Invoice)', len(unmatched), f'{len(unmatched)/total*100:.0f}%' if total else ''),
    ]
    for i, (metric, count, pct) in enumerate(rows, 2):
        ws_summary.cell(row=i, column=1, value=metric)
        ws_summary.cell(row=i, column=2, value=count)
        ws_summary.cell(row=i, column=3, value=pct)

        if i == 3:
            ws_summary.cell(row=i, column=1).fill = green_fill
        elif i == 4:
            ws_summary.cell(row=i, column=1).fill = yellow_fill
        elif i == 5:
            ws_summary.cell(row=i, column=1).fill = red_fill

    ws_summary.column_dimensions['A'].width = 30
    ws_summary.column_dimensions['B'].width = 10
    ws_summary.column_dimensions['C'].width = 10

    # --- Matched tab ---
    ws_matched = wb.create_sheet('Matched')
    headers = ['Txn Date', 'Vendor', 'Amount', 'Description', 'Txn Type', 'Txn Num',
               'Invoice Source', 'Invoice File/Email', 'Attachment Path', 'Confidence',
               'Invoice Amount', 'Invoice Date', 'Sender/Filename', 'Attach Status']
    write_header(ws_matched, headers)

    for row_idx, (txn_idx, txn, candidates) in enumerate(matched, 2):
        best = candidates[0]
        ws_matched.cell(row=row_idx, column=1, value=txn['date'].strftime('%d/%m/%Y'))
        ws_matched.cell(row=row_idx, column=2, value=txn['vendor'])
        ws_matched.cell(row=row_idx, column=3, value=txn['amount'])
        ws_matched.cell(row=row_idx, column=4, value=txn['description'])
        ws_matched.cell(row=row_idx, column=5, value=txn.get('type', ''))
        ws_matched.cell(row=row_idx, column=6, value=txn.get('num', ''))
        ws_matched.cell(row=row_idx, column=7, value=best.get('source', ''))
        ws_matched.cell(row=row_idx, column=8, value=best.get('subject', ''))
        ws_matched.cell(row=row_idx, column=9, value=best.get('attachment_path', ''))
        ws_matched.cell(row=row_idx, column=10, value=f"{best['score']:.0f}%")
        ws_matched.cell(row=row_idx, column=11, value=best['amounts'][0] if best.get('amounts') else '')
        ws_matched.cell(row=row_idx, column=12, value=best['date'].strftime('%d/%m/%Y') if best.get('date') else '')
        ws_matched.cell(row=row_idx, column=13, value=best.get('sender_name', ''))
        ws_matched.cell(row=row_idx, column=14, value='')  # Attach Status — filled by qbo_attach

        for col in range(1, 15):
            ws_matched.cell(row=row_idx, column=col).fill = green_fill

    _auto_width(ws_matched)

    # --- Uncertain tab ---
    ws_uncertain = wb.create_sheet('Needs Review')
    headers = ['Txn Date', 'Vendor', 'Amount', 'Description',
               'Best Match Source', 'Best Match File/Email', 'Confidence',
               'Match Amount', 'Match Date', 'Other Candidates']
    write_header(ws_uncertain, headers)

    for row_idx, (txn_idx, txn, candidates) in enumerate(uncertain, 2):
        best = candidates[0]
        ws_uncertain.cell(row=row_idx, column=1, value=txn['date'].strftime('%d/%m/%Y'))
        ws_uncertain.cell(row=row_idx, column=2, value=txn['vendor'])
        ws_uncertain.cell(row=row_idx, column=3, value=txn['amount'])
        ws_uncertain.cell(row=row_idx, column=4, value=txn['description'])
        ws_uncertain.cell(row=row_idx, column=5, value=best.get('source', ''))
        ws_uncertain.cell(row=row_idx, column=6, value=best.get('attachment_path') or best.get('subject', ''))
        ws_uncertain.cell(row=row_idx, column=7, value=f"{best['score']:.0f}%")
        ws_uncertain.cell(row=row_idx, column=8, value=best['amounts'][0] if best.get('amounts') else '')
        ws_uncertain.cell(row=row_idx, column=9, value=best['date'].strftime('%d/%m/%Y') if best.get('date') else '')

        # Other candidates
        other = '; '.join([
            f"{c.get('sender_name', '')} ({c['score']:.0f}%)"
            for c in candidates[1:3]
        ])
        ws_uncertain.cell(row=row_idx, column=10, value=other)

        for col in range(1, 11):
            ws_uncertain.cell(row=row_idx, column=col).fill = yellow_fill

    _auto_width(ws_uncertain)

    # --- Unmatched tab ---
    ws_unmatched = wb.create_sheet('Unmatched')
    headers = ['Txn Date', 'Vendor', 'Amount', 'Description', 'Type', 'Suggested Action']
    write_header(ws_unmatched, headers)

    for row_idx, (txn_idx, txn) in enumerate(unmatched, 2):
        ws_unmatched.cell(row=row_idx, column=1, value=txn['date'].strftime('%d/%m/%Y'))
        ws_unmatched.cell(row=row_idx, column=2, value=txn['vendor'])
        ws_unmatched.cell(row=row_idx, column=3, value=txn['amount'])
        ws_unmatched.cell(row=row_idx, column=4, value=txn['description'])
        ws_unmatched.cell(row=row_idx, column=5, value=txn['type'])

        # Suggest where to find the invoice
        vendor = txn['vendor']
        if vendor:
            action = f"Check {vendor}'s portal or email for invoice"
        else:
            action = "Identify vendor and locate invoice"
        ws_unmatched.cell(row=row_idx, column=6, value=action)

        for col in range(1, 7):
            ws_unmatched.cell(row=row_idx, column=col).fill = red_fill

    _auto_width(ws_unmatched)

    # Save
    wb.save(output_path)
    print(f"\nMatch report saved to: {output_path}")
    print(f"  Matched:   {len(matched)} ({len(matched)/total*100:.0f}%)" if total else "")
    print(f"  Uncertain: {len(uncertain)} ({len(uncertain)/total*100:.0f}%)" if total else "")
    print(f"  Unmatched: {len(unmatched)} ({len(unmatched)/total*100:.0f}%)" if total else "")

    return {
        'matched': len(matched),
        'uncertain': len(uncertain),
        'unmatched': len(unmatched),
        'total': total,
    }


def _auto_width(ws):
    """Auto-adjust column widths."""
    for col in ws.columns:
        max_len = 0
        col_letter = col[0].column_letter
        for cell in col:
            try:
                if cell.value:
                    max_len = max(max_len, len(str(cell.value)))
            except Exception:
                pass
        ws.column_dimensions[col_letter].width = min(max_len + 2, 50)


# ---------------------------------------------------------------------------
# Main CLI
# ---------------------------------------------------------------------------

def main():
    parser = argparse.ArgumentParser(description='QBO Invoice Matcher')

    parser.add_argument('--transactions', '-t',
                        help='QBO transaction export CSV/Excel')
    parser.add_argument('--gmail-label', '-g',
                        default='accounting - Ltd Expenses - 24/25 Receipts',
                        help='Gmail label to search for invoices')
    parser.add_argument('--local-folder', '-l',
                        help='Local folder to search for invoice PDFs')
    parser.add_argument('--output', '-o', default='match_report.xlsx',
                        help='Output Excel file path')
    parser.add_argument('--no-gmail', action='store_true',
                        help='Skip Gmail search')
    parser.add_argument('--no-local', action='store_true',
                        help='Skip local folder search')
    parser.add_argument('--attach', metavar='REPORT',
                        help='Attach matched invoices to QBO (pass match report xlsx)')
    parser.add_argument('--dry-run', action='store_true',
                        help='With --attach: find transactions in QBO but do not upload')
    parser.add_argument('--qbo-url', default='https://qbo.intuit.co.uk',
                        help='QBO base URL (default: UK)')

    args = parser.parse_args()

    if args.attach:
        import asyncio
        from qbo_attach import attach_all_from_report
        asyncio.run(attach_all_from_report(
            args.attach,
            qbo_base_url=args.qbo_url,
            dry_run=args.dry_run,
        ))
        sys.exit(0)

    if not args.transactions:
        print("ERROR: --transactions required (QBO export CSV/Excel)")
        sys.exit(1)

    # Parse transactions
    print(f"Parsing transactions from: {args.transactions}")
    transactions = parse_transactions(args.transactions)
    print(f"Found {len(transactions)} transactions")

    if not transactions:
        print("No transactions found. Check your file format.")
        sys.exit(1)

    # Create cache dir for downloaded attachments
    cache_dir = TOOL_DIR / '.cache' / 'attachments'
    cache_dir.mkdir(parents=True, exist_ok=True)

    # Search Gmail
    gmail_matches = {}
    if not args.no_gmail:
        print(f"\nSearching Gmail label: {args.gmail_label}")
        gmail_matches = search_gmail_invoices(args.gmail_label, transactions, cache_dir)
        print(f"Gmail matches found for {len(gmail_matches)} transactions")

    # Search local folder
    local_matches = {}
    if not args.no_local and args.local_folder:
        print(f"\nSearching local folder: {args.local_folder}")
        local_matches = search_local_invoices(args.local_folder, transactions)
        print(f"Local matches found for {len(local_matches)} transactions")

    # Write report
    print(f"\nWriting match report...")
    write_match_report(transactions, gmail_matches, local_matches, args.output)


if __name__ == '__main__':
    main()
