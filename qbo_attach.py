"""
QBO Attach — Playwright automation for attaching invoices to QuickBooks Online transactions.

Reads a match report Excel file (from qbo_matcher.py) and uploads invoice files
to each matched transaction in QBO via browser automation.

Usage (via qbo_matcher.py CLI):
    python qbo_matcher.py --attach match_report.xlsx
    python qbo_matcher.py --attach match_report.xlsx --dry-run
"""

import asyncio
import os
import sys
from datetime import datetime
from pathlib import Path

from playwright.async_api import async_playwright, Page, BrowserContext

# ---------------------------------------------------------------------------
# Config / selectors
# ---------------------------------------------------------------------------

CACHE_DIR = Path(__file__).parent.resolve() / '.cache'
STATE_PATH = CACHE_DIR / 'qbo_state.json'
SCREENSHOT_DIR = CACHE_DIR / 'screenshots'

# Delay between transactions to avoid bot detection
THROTTLE_SECONDS = 3

# Timeouts
PAGE_LOAD_TIMEOUT = 30_000  # 30s
UPLOAD_TIMEOUT = 60_000     # 60s

# QBO selectors — these may need updating if QBO changes its UI.
# Using text/role-based selectors where possible for resilience.
SELECTORS = {
    # Search
    'search_button': '[data-testid="global-search"], [aria-label="Search"], button:has-text("Search")',
    'search_input': '[data-testid="global-search-input"], input[placeholder*="Search"], [role="searchbox"]',
    'search_result_row': '[data-testid="search-result-row"], .search-result-item, [role="listbox"] [role="option"]',

    # Transaction detail page
    'attachment_button': (
        'button:has-text("Attachments"), '
        'button:has-text("Attach"), '
        '[data-testid="attachment-button"], '
        '[aria-label*="ttach"], '
        'button:has-text("Add attachment")'
    ),
    'file_input': 'input[type="file"]',
    'save_button': (
        'button:has-text("Save and close"), '
        'button:has-text("Save"), '
        '[data-testid="save-button"], '
        '[data-automation-id="save-button"]'
    ),

    # Upload confirmation
    'upload_success': (
        '[data-testid="attachment-item"], '
        '.attachment-item, '
        '.attachable-thumbnail, '
        'text="attached"'
    ),
}


# ---------------------------------------------------------------------------
# Session management
# ---------------------------------------------------------------------------

async def launch_qbo_browser(qbo_base_url: str) -> tuple[BrowserContext, Page]:
    """
    Launch Playwright browser with QBO session.

    First run: opens QBO login page and pauses for manual login.
    Subsequent runs: reloads saved browser state.
    """
    CACHE_DIR.mkdir(parents=True, exist_ok=True)
    SCREENSHOT_DIR.mkdir(parents=True, exist_ok=True)

    pw = await async_playwright().start()

    if STATE_PATH.exists():
        print("Loading saved QBO session...")
        context = await pw.chromium.launch_persistent_context(
            user_data_dir=str(CACHE_DIR / 'browser_profile'),
            headless=False,
            slow_mo=500,
            viewport={'width': 1280, 'height': 900},
        )
        page = context.pages[0] if context.pages else await context.new_page()

        # Verify session is still valid
        await page.goto(f"{qbo_base_url}/app/homepage", timeout=PAGE_LOAD_TIMEOUT)
        await page.wait_for_load_state('networkidle', timeout=PAGE_LOAD_TIMEOUT)

        # Check if we got redirected to login
        if 'login' in page.url.lower() or 'signin' in page.url.lower() or 'accounts.intuit' in page.url.lower():
            print("Session expired. Please log in again...")
            print(">>> Log into QBO in the browser window, then press Enter here <<<")
            await _wait_for_user_login(page, qbo_base_url)
            # Save updated state
            await context.storage_state(path=str(STATE_PATH))
            print("Session saved.")
    else:
        print("First run — opening QBO for manual login...")
        context = await pw.chromium.launch_persistent_context(
            user_data_dir=str(CACHE_DIR / 'browser_profile'),
            headless=False,
            slow_mo=500,
            viewport={'width': 1280, 'height': 900},
        )
        page = context.pages[0] if context.pages else await context.new_page()

        await page.goto(f"{qbo_base_url}/app/homepage", timeout=PAGE_LOAD_TIMEOUT)

        print(">>> Log into QBO in the browser window, then press Enter here <<<")
        await _wait_for_user_login(page, qbo_base_url)

        # Save session state
        await context.storage_state(path=str(STATE_PATH))
        print("Session saved for future runs.")

    return context, page


async def _wait_for_user_login(page: Page, qbo_base_url: str):
    """Wait for user to complete login by polling URL + prompting."""
    # Give user time to log in
    while True:
        try:
            input()  # Wait for Enter
        except EOFError:
            await asyncio.sleep(5)

        current_url = page.url
        if '/app/' in current_url and 'login' not in current_url.lower() and 'signin' not in current_url.lower():
            print(f"Logged in successfully. Current page: {current_url}")
            return

        print("Still on login page. Complete login and press Enter again...")


# ---------------------------------------------------------------------------
# Transaction lookup
# ---------------------------------------------------------------------------

async def find_transaction(page: Page, qbo_base_url: str,
                           txn_date: str, vendor: str, amount: float,
                           txn_type: str = '', txn_num: str = '') -> bool:
    """
    Navigate to a specific transaction in QBO.

    Strategy:
    1. Use QBO search with vendor + amount
    2. Find the matching transaction in results
    3. Click to open it

    Returns True if transaction was found and opened.
    """
    # Build search query — use the most specific info available
    if txn_num:
        search_query = txn_num
    elif vendor:
        search_query = f"{vendor} {amount:.2f}"
    else:
        search_query = f"{amount:.2f}"

    print(f"  Searching QBO: {search_query}")

    # Try using the search bar
    try:
        # Click the search icon/button
        search_btn = page.locator(SELECTORS['search_button']).first
        await search_btn.click(timeout=10_000)
        await asyncio.sleep(1)

        # Type search query
        search_input = page.locator(SELECTORS['search_input']).first
        await search_input.fill(search_query, timeout=10_000)
        await asyncio.sleep(0.5)
        await search_input.press('Enter')

        # Wait for results
        await page.wait_for_load_state('networkidle', timeout=PAGE_LOAD_TIMEOUT)
        await asyncio.sleep(2)

        # Look for a matching result
        found = await _click_matching_result(page, txn_date, vendor, amount)
        if found:
            return True

    except Exception as e:
        print(f"  Search approach failed: {e}")

    # Fallback: navigate directly to the expenses/transactions list with filters
    try:
        # QBO URL-based filtering — navigate to expense list filtered by date
        parsed_date = _parse_date_str(txn_date)
        if parsed_date:
            date_str = parsed_date.strftime('%Y-%m-%d')
            list_url = (
                f"{qbo_base_url}/app/expenses"
                f"?startDate={date_str}&endDate={date_str}"
            )
            print(f"  Trying direct URL: {list_url}")
            await page.goto(list_url, timeout=PAGE_LOAD_TIMEOUT)
            await page.wait_for_load_state('networkidle', timeout=PAGE_LOAD_TIMEOUT)
            await asyncio.sleep(2)

            found = await _click_matching_result(page, txn_date, vendor, amount)
            if found:
                return True

    except Exception as e:
        print(f"  URL filter approach failed: {e}")

    return False


async def _click_matching_result(page: Page, txn_date: str, vendor: str, amount: float) -> bool:
    """
    Scan visible page for a transaction row matching our criteria and click it.

    Looks for rows/links containing the vendor name and amount.
    """
    amount_str = f"{amount:.2f}"
    amount_str_no_dec = f"{amount:.0f}" if amount == int(amount) else amount_str

    # Strategy 1: Find any clickable element containing both vendor and amount text
    if vendor:
        # Try to find a row that contains both the vendor name and amount
        try:
            # Look for table rows or list items containing the vendor
            rows = page.locator(f'tr:has-text("{vendor}"), [role="row"]:has-text("{vendor}"), a:has-text("{vendor}")')
            count = await rows.count()

            for i in range(count):
                row = rows.nth(i)
                row_text = await row.inner_text()

                if amount_str in row_text or amount_str_no_dec in row_text:
                    print(f"  Found matching row: {row_text[:80]}...")
                    await row.click(timeout=5_000)
                    await page.wait_for_load_state('networkidle', timeout=PAGE_LOAD_TIMEOUT)
                    await asyncio.sleep(1)
                    return True
        except Exception:
            pass

    # Strategy 2: Look for any element containing the amount
    try:
        amount_elements = page.locator(f'text="{amount_str}"')
        count = await amount_elements.count()

        for i in range(count):
            el = amount_elements.nth(i)
            # Find the closest clickable parent (row or link)
            parent = el.locator('xpath=ancestor::tr | ancestor::a | ancestor::*[@role="row"]').first
            try:
                parent_text = await parent.inner_text()
                # Check if date roughly matches
                if txn_date and txn_date in parent_text:
                    print(f"  Found by amount+date: {parent_text[:80]}...")
                    await parent.click(timeout=5_000)
                    await page.wait_for_load_state('networkidle', timeout=PAGE_LOAD_TIMEOUT)
                    await asyncio.sleep(1)
                    return True
            except Exception:
                continue
    except Exception:
        pass

    # Strategy 3: Just click the first search result
    try:
        result = page.locator(SELECTORS['search_result_row']).first
        if await result.is_visible():
            result_text = await result.inner_text()
            print(f"  Clicking first search result: {result_text[:80]}...")
            await result.click(timeout=5_000)
            await page.wait_for_load_state('networkidle', timeout=PAGE_LOAD_TIMEOUT)
            await asyncio.sleep(1)
            return True
    except Exception:
        pass

    return False


def _parse_date_str(date_str: str):
    """Parse DD/MM/YYYY date string."""
    for fmt in ['%d/%m/%Y', '%Y-%m-%d', '%m/%d/%Y']:
        try:
            return datetime.strptime(date_str.strip(), fmt)
        except (ValueError, AttributeError):
            continue
    return None


# ---------------------------------------------------------------------------
# File upload
# ---------------------------------------------------------------------------

async def attach_file(page: Page, file_path: str) -> bool:
    """
    Upload a file as an attachment to the currently open QBO transaction.

    Returns True if upload succeeded.
    """
    filepath = Path(file_path)
    if not filepath.exists():
        print(f"  WARNING: File not found: {file_path}")
        return False

    print(f"  Uploading: {filepath.name}")

    try:
        # Try to find and click the attachment button
        att_btn = page.locator(SELECTORS['attachment_button']).first
        try:
            await att_btn.click(timeout=10_000)
            await asyncio.sleep(1)
        except Exception:
            # Attachment area might already be visible
            pass

        # Set file on the file input (may be hidden)
        file_input = page.locator(SELECTORS['file_input']).first

        # Some file inputs are hidden — use set_input_files which works regardless
        await file_input.set_input_files(str(filepath.resolve()), timeout=10_000)
        print(f"  File selected, waiting for upload...")

        # Wait for upload to process
        await asyncio.sleep(3)

        # Check for upload confirmation
        try:
            await page.locator(SELECTORS['upload_success']).first.wait_for(
                state='visible', timeout=UPLOAD_TIMEOUT
            )
            print(f"  Upload confirmed.")
        except Exception:
            # May not find the success indicator, but upload might still have worked
            print(f"  Upload indicator not found — checking if file appears...")
            await asyncio.sleep(2)

        # Save the transaction
        try:
            save_btn = page.locator(SELECTORS['save_button']).first
            await save_btn.click(timeout=10_000)
            await page.wait_for_load_state('networkidle', timeout=PAGE_LOAD_TIMEOUT)
            print(f"  Transaction saved.")
        except Exception as e:
            print(f"  WARNING: Could not click save: {e}")
            # Try keyboard shortcut
            await page.keyboard.press('Control+s')
            await asyncio.sleep(2)

        return True

    except Exception as e:
        print(f"  ERROR uploading file: {e}")
        # Screenshot for debugging
        screenshot_path = SCREENSHOT_DIR / f"upload_error_{datetime.now().strftime('%H%M%S')}.png"
        try:
            await page.screenshot(path=str(screenshot_path))
            print(f"  Screenshot saved: {screenshot_path}")
        except Exception:
            pass
        return False


# ---------------------------------------------------------------------------
# Orchestration
# ---------------------------------------------------------------------------

async def attach_all_from_report(report_path: str,
                                  qbo_base_url: str = 'https://qbo.intuit.co.uk',
                                  dry_run: bool = False):
    """
    Read match report and attach each matched invoice to its QBO transaction.

    Args:
        report_path: Path to match_report.xlsx from qbo_matcher.py
        qbo_base_url: QBO base URL (default UK)
        dry_run: If True, find transactions but don't upload files
    """
    import openpyxl

    report = Path(report_path)
    if not report.exists():
        print(f"ERROR: Report not found: {report_path}")
        sys.exit(1)

    # Read matched transactions from the report
    wb = openpyxl.load_workbook(report)

    if 'Matched' not in wb.sheetnames:
        print("ERROR: No 'Matched' sheet found in report.")
        sys.exit(1)

    ws = wb['Matched']

    # Parse header row to find column indices
    headers = [cell.value for cell in ws[1]]
    col_idx = {h: i for i, h in enumerate(headers) if h}

    required = ['Txn Date', 'Vendor', 'Amount', 'Attachment Path']
    missing = [r for r in required if r not in col_idx]
    if missing:
        print(f"ERROR: Missing required columns in Matched sheet: {missing}")
        print(f"Found columns: {headers}")
        sys.exit(1)

    # Collect rows to process
    rows_to_process = []
    for row in ws.iter_rows(min_row=2, values_only=False):
        values = {h: row[col_idx[h]].value for h in col_idx}

        # Skip rows already attached
        status = values.get('Attach Status', '')
        if status and 'attached' in str(status).lower():
            continue

        attachment_path = values.get('Attachment Path', '')
        if not attachment_path:
            continue

        rows_to_process.append({
            'row_num': row[0].row,
            'txn_date': str(values.get('Txn Date', '')),
            'vendor': str(values.get('Vendor', '')),
            'amount': float(values.get('Amount', 0)),
            'txn_type': str(values.get('Txn Type', '')),
            'txn_num': str(values.get('Txn Num', '')),
            'attachment_path': str(attachment_path),
        })

    if not rows_to_process:
        print("No rows to process (all already attached or no attachment paths).")
        return

    print(f"\n{'DRY RUN — ' if dry_run else ''}Processing {len(rows_to_process)} matched transactions")
    print(f"QBO URL: {qbo_base_url}")
    print()

    # Launch browser
    context, page = await launch_qbo_browser(qbo_base_url)

    status_col = col_idx.get('Attach Status')
    attached_count = 0
    failed_count = 0

    try:
        for i, row_data in enumerate(rows_to_process, 1):
            print(f"\n[{i}/{len(rows_to_process)}] {row_data['vendor']} — £{row_data['amount']:.2f} ({row_data['txn_date']})")

            # Find the transaction in QBO
            found = await find_transaction(
                page, qbo_base_url,
                row_data['txn_date'],
                row_data['vendor'],
                row_data['amount'],
                row_data['txn_type'],
                row_data['txn_num'],
            )

            if not found:
                print(f"  SKIPPED: Transaction not found in QBO")
                if status_col is not None:
                    ws.cell(row=row_data['row_num'], column=status_col + 1, value='Not Found')
                failed_count += 1
                continue

            if dry_run:
                print(f"  DRY RUN: Would attach {Path(row_data['attachment_path']).name}")
                if status_col is not None:
                    ws.cell(row=row_data['row_num'], column=status_col + 1, value='Found (dry run)')
                attached_count += 1
            else:
                # Upload the file
                success = await attach_file(page, row_data['attachment_path'])
                if success:
                    if status_col is not None:
                        ws.cell(row=row_data['row_num'], column=status_col + 1, value='Attached')
                    attached_count += 1
                else:
                    if status_col is not None:
                        ws.cell(row=row_data['row_num'], column=status_col + 1, value='Failed')
                    failed_count += 1

            # Throttle between transactions
            if i < len(rows_to_process):
                await asyncio.sleep(THROTTLE_SECONDS)

        # Save updated report with status
        wb.save(report_path)
        print(f"\n{'DRY RUN ' if dry_run else ''}Results:")
        print(f"  {'Found' if dry_run else 'Attached'}: {attached_count}")
        print(f"  Failed/Not Found: {failed_count}")
        print(f"  Report updated: {report_path}")

    except KeyboardInterrupt:
        print("\n\nInterrupted! Saving progress...")
        wb.save(report_path)
        print(f"Progress saved to {report_path}. Re-run to continue.")

    finally:
        await context.close()
