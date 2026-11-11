#!/usr/bin/env python3
"""
PayPal to Google Sheets Automation Script

This script automates transferring PayPal transaction data from downloaded CSV files
into a Google Sheet for expense tracking. It handles:
- Pairing payment rows with deposit rows via Reference Txn ID
- Parsing "Song - Artist" format from payment notes
- Splitting payments among multiple artists
- Invoice payments (using Item Title instead of Note)
- Currency conversions (always uses USD amount from General Card Deposit)
- Creating/updating monthly sheets automatically
"""

import os
import re
import sys
import time
from datetime import datetime
from collections import defaultdict
from typing import Dict, List, Optional, Tuple

import pandas as pd
import gspread
from gspread_formatting import (
    CellFormat,
    TextFormat,
    Color,
    NumberFormat,
    format_cell_range,
    DataValidationRule,
    BooleanCondition,
    set_data_validation_for_cell_range
)
from google.oauth2.service_account import Credentials

import config


# Google Sheets API scopes
SCOPES = [
    'https://www.googleapis.com/auth/spreadsheets',
    'https://www.googleapis.com/auth/drive'
]

# Sheet header row (row 1)
HEADERS = [
    'Date', 'Payment Method', 'Charge Type', 'Split Charge',
    'Artist Name', 'Song Name', 'Expense Description',
    'Initial Amount', 'Additional Fees', 'Final Amount'
]

# Dropdown options for data validation
PAYMENT_METHODS = ['Robinhood', 'AMEX', 'Apple', 'Chase', 'Wells Fargo']
CHARGE_TYPES = ['Cost', 'Reimbursement']
SPLIT_CHARGE_OPTIONS = ['Yes', 'No']

# Minimalist styling colors (RGB values 0-1)
# Header: Light blue #D9E7F8
HEADER_BG_COLOR = Color(0.85, 0.91, 0.97)
# Column B (Payment Method): Light cream #FFF8DC
PAYMENT_METHOD_BG_COLOR = Color(1.0, 0.97, 0.86)
# Column C (Charge Type): Light green #E8F5E9
CHARGE_TYPE_BG_COLOR = Color(0.91, 0.96, 0.91)
# Column D (Split Charge): Light blue #E3F2FD
SPLIT_CHARGE_BG_COLOR = Color(0.89, 0.95, 0.99)

# Regex pattern to match "Song - Artist" format in notes
# Matches text in quotes with a dash separator, handling variable whitespace
# Pattern: "([^"]+)\s*-\s*([^"]+)" captures Song (before dash) and Artist (after dash)
SONG_ARTIST_PATTERN = re.compile(r'"([^"]+?)\s*-\s*([^"]+?)"')


class PayPalProcessor:
    """Processes PayPal CSV files and uploads to Google Sheets."""

    def __init__(self):
        self.client = None
        self.spreadsheet = None
        self.warnings = []
        self.errors = []

    def authenticate(self) -> bool:
        """
        Authenticate with Google Sheets API using service account credentials.
        Returns True if successful, False otherwise.
        """
        try:
            credentials = Credentials.from_service_account_file(
                config.CREDENTIALS_FILE,
                scopes=SCOPES
            )
            self.client = gspread.authorize(credentials)
            self.spreadsheet = self.client.open_by_key(config.GOOGLE_SHEET_ID)
            print(f"Successfully authenticated with Google Sheets")
            return True
        except FileNotFoundError:
            print(f"Error: Credentials file '{config.CREDENTIALS_FILE}' not found")
            return False
        except gspread.exceptions.SpreadsheetNotFound:
            print(f"Error: Google Sheet not accessible. Make sure credentials.json is valid and Sheet ID is correct")
            return False
        except Exception as e:
            print(f"Error authenticating: {e}")
            return False

    def apply_sheet_styling(self, worksheet: gspread.Worksheet, last_data_row: int = 1000) -> None:
        """
        Apply minimalist professional styling to the worksheet.

        Styling applied:
        - Header row (row 1): Light blue background, bold text
        - Column B (Payment Method): Light cream background
        - Column C (Charge Type): Light green background
        - Column D (Split Charge): Light blue background
        - Columns H, I, J (amounts): Right-aligned, currency format

        Args:
            worksheet: The worksheet to style
            last_data_row: Last row to apply column styling to (default 1000)
        """
        try:
            # Header row: Light blue background with bold text
            header_format = CellFormat(
                backgroundColor=HEADER_BG_COLOR,
                textFormat=TextFormat(bold=True)
            )
            format_cell_range(worksheet, 'A1:J1', header_format)

            # Column B (Payment Method): Light cream background
            payment_method_format = CellFormat(backgroundColor=PAYMENT_METHOD_BG_COLOR)
            format_cell_range(worksheet, f'B2:B{last_data_row}', payment_method_format)

            # Column C (Charge Type): Light green background
            charge_type_format = CellFormat(backgroundColor=CHARGE_TYPE_BG_COLOR)
            format_cell_range(worksheet, f'C2:C{last_data_row}', charge_type_format)

            # Column D (Split Charge): Light blue background
            split_charge_format = CellFormat(backgroundColor=SPLIT_CHARGE_BG_COLOR)
            format_cell_range(worksheet, f'D2:D{last_data_row}', split_charge_format)

            # Columns H, I, J (amounts): Right-aligned with currency format
            amount_format = CellFormat(
                horizontalAlignment='RIGHT',
                numberFormat=NumberFormat(type='CURRENCY', pattern='$#,##0.00')
            )
            format_cell_range(worksheet, f'H2:J{last_data_row}', amount_format)

            print("Applied minimalist styling to sheet")

        except Exception as e:
            print(f"Warning: Could not apply sheet styling: {e}")

    def setup_dropdown_validation(self, worksheet: gspread.Worksheet, start_row: int = 2, end_row: int = 1000) -> None:
        """
        Set up data validation (dropdowns) for columns B, C, and D.

        Column B (Payment Method): Robinhood, AMEX, Apple, Chase, Wells Fargo
        Column C (Charge Type): Cost, Reimbursement
        Column D (Split Charge): Yes, No

        Args:
            worksheet: The worksheet to apply validation to
            start_row: First row to apply validation (default 2, after header)
            end_row: Last row to apply validation (default 1000)
        """
        try:
            # Column B - Payment Method
            payment_rule = DataValidationRule(
                BooleanCondition('ONE_OF_LIST', PAYMENT_METHODS),
                showCustomUi=True,
                strict=True
            )
            set_data_validation_for_cell_range(
                worksheet,
                f'B{start_row}:B{end_row}',
                payment_rule
            )

            # Column C - Charge Type
            charge_type_rule = DataValidationRule(
                BooleanCondition('ONE_OF_LIST', CHARGE_TYPES),
                showCustomUi=True,
                strict=True
            )
            set_data_validation_for_cell_range(
                worksheet,
                f'C{start_row}:C{end_row}',
                charge_type_rule
            )

            # Column D - Split Charge
            split_charge_rule = DataValidationRule(
                BooleanCondition('ONE_OF_LIST', SPLIT_CHARGE_OPTIONS),
                showCustomUi=True,
                strict=True
            )
            set_data_validation_for_cell_range(
                worksheet,
                f'D{start_row}:D{end_row}',
                split_charge_rule
            )

            print(f"Applied dropdown validation to columns B, C, D (rows {start_row}-{end_row})")
        except Exception as e:
            print(f"Warning: Could not set up dropdown validation: {e}")

    def remove_instruction_row_if_exists(self, worksheet: gspread.Worksheet) -> bool:
        """
        Check if row 2 contains instruction text and remove it if found.
        Returns True if instruction row was found and removed.
        """
        try:
            # Get row 2 values
            row2 = worksheet.row_values(2)
            if not row2:
                return False

            # Check if this looks like the old instruction row
            # The instruction row typically has 'MM/DD/YYYY' in cell A2
            if row2 and row2[0] == 'MM/DD/YYYY':
                worksheet.delete_rows(2)
                print("Removed legacy instruction row from sheet")
                return True
            return False
        except Exception as e:
            print(f"Warning: Could not check/remove instruction row: {e}")
            return False

    def calculate_text_width(self, text: str, is_bold: bool = False) -> int:
        """
        Calculate approximate pixel width needed for text.

        Uses character-based estimation:
        - Average character width: ~7 pixels for normal text, ~8 for bold
        - Adds padding for cell borders

        Args:
            text: The text to measure
            is_bold: Whether the text is bold (slightly wider)

        Returns:
            Estimated pixel width needed
        """
        if not text:
            return 0
        text_str = str(text)
        char_width = 8 if is_bold else 7
        # Add extra padding (20px) for cell borders and breathing room
        return (len(text_str) * char_width) + 20

    def auto_resize_columns(self, worksheet: gspread.Worksheet, min_width: int = 80, max_width: int = 400) -> None:
        """
        Auto-resize columns A-J to fit content, ensuring headers and data are fully visible.

        Width calculation:
        - Calculates width needed for header text (bold)
        - Calculates width needed for longest data in each column
        - Sets column width to max(header_width, data_width, min_width)
        - Caps at max_width to prevent excessively wide columns

        Args:
            worksheet: The worksheet to resize columns for
            min_width: Minimum column width in pixels (default 80)
            max_width: Maximum column width in pixels (default 400)
        """
        try:
            sheet_id = worksheet.id

            # Get all values to calculate widths (header + data)
            all_values = worksheet.get_all_values()
            if not all_values:
                return

            # Calculate required width for each column
            column_widths = []
            num_columns = min(10, len(all_values[0]) if all_values else 0)  # Columns A-J

            for col_idx in range(num_columns):
                max_content_width = 0

                for row_idx, row in enumerate(all_values):
                    if col_idx < len(row):
                        cell_value = row[col_idx]
                        # Row 1 (index 0) is the header - it's bold
                        is_header = (row_idx == 0)
                        text_width = self.calculate_text_width(cell_value, is_bold=is_header)
                        max_content_width = max(max_content_width, text_width)

                # Apply min/max constraints
                final_width = max(min_width, min(max_content_width, max_width))
                column_widths.append(final_width)

            # Build batch update requests to set column widths
            requests = []
            for col_idx, width in enumerate(column_widths):
                requests.append({
                    'updateDimensionProperties': {
                        'range': {
                            'sheetId': sheet_id,
                            'dimension': 'COLUMNS',
                            'startIndex': col_idx,
                            'endIndex': col_idx + 1
                        },
                        'properties': {'pixelSize': width},
                        'fields': 'pixelSize'
                    }
                })

            # Apply all column width updates
            if requests:
                self.spreadsheet.batch_update({'requests': requests})

            print("Auto-resized columns A-J to fit content")

        except Exception as e:
            print(f"Warning: Could not auto-resize columns: {e}")

    def parse_sheet_name_date(self, sheet_name: str) -> Optional[datetime]:
        """
        Parse a sheet name in "Month Year" format (e.g., "January 2026") to a datetime.

        Args:
            sheet_name: The name of the sheet to parse

        Returns:
            datetime object if parsing succeeds, None otherwise
        """
        try:
            # Try to parse "Month Year" format (e.g., "December 2025")
            return datetime.strptime(sheet_name, '%B %Y')
        except ValueError:
            return None

    def reorder_sheets_chronologically(self) -> None:
        """
        Reorder all monthly sheets so newest months appear on the left (first position)
        and oldest months on the right. Non-monthly sheets are kept at the end.

        Sheet order example: [January 2026] [December 2025] [November 2025] ...
        """
        try:
            all_worksheets = self.spreadsheet.worksheets()

            # Separate monthly sheets from non-monthly sheets
            monthly_sheets = []
            non_monthly_sheets = []

            for ws in all_worksheets:
                sheet_date = self.parse_sheet_name_date(ws.title)
                if sheet_date:
                    monthly_sheets.append((ws, sheet_date))
                else:
                    non_monthly_sheets.append(ws)

            # Sort monthly sheets by date (newest first)
            monthly_sheets.sort(key=lambda x: x[1], reverse=True)

            # Build the new order: monthly sheets (newest first), then non-monthly sheets
            new_order = [ws for ws, _ in monthly_sheets] + non_monthly_sheets

            # Only reorder if we have sheets and the order would change
            if new_order and new_order != all_worksheets:
                self.spreadsheet.reorder_worksheets(new_order)
                print("Reordered sheets chronologically (newest first)")

        except Exception as e:
            print(f"Warning: Could not reorder sheets: {e}")

    def read_csv(self, filepath: str) -> Optional[pd.DataFrame]:
        """
        Read a PayPal CSV file with proper encoding handling.
        Tries UTF-8 first, then falls back to latin-1.
        """
        for encoding in ['utf-8', 'latin-1']:
            try:
                df = pd.read_csv(filepath, encoding=encoding)
                return df
            except UnicodeDecodeError:
                continue
            except Exception as e:
                print(f"Error reading CSV: {e}")
                return None
        print(f"Error: Could not read CSV file with any supported encoding")
        return None

    def parse_amount(self, amount_str: str) -> float:
        """
        Parse amount string to float, handling commas and negative values.
        Examples: "-4,120.00" -> 4120.00 (absolute value)
        """
        if pd.isna(amount_str):
            return 0.0
        # Remove commas and convert to float
        amount_str = str(amount_str).replace(',', '')
        try:
            return abs(float(amount_str))
        except ValueError:
            return 0.0

    def parse_note(self, note: str) -> List[Dict[str, str]]:
        """
        Parse the payment note to extract Song/Artist pairs.

        The expected note format follows this pattern:
        - Single: "Song Name - Artist Name" description
        - Multiple: "Song1 - Artist1" deliverable1 "Song2 - Artist2" deliverable2 @shared_info

        Returns a list of dicts with 'song', 'artist' keys.
        Empty list if no valid patterns found.
        """
        if not note or pd.isna(note):
            return []

        # Find all "Song - Artist" patterns
        matches = SONG_ARTIST_PATTERN.findall(note)

        results = []
        for match in matches:
            # match is a tuple: (song_part, artist_part)
            song = match[0].strip()
            artist = match[1].strip()
            results.append({'song': song, 'artist': artist})

        return results

    def extract_description(self, note: str, song_artist_pairs: List[Dict]) -> str:
        """
        Extract the description from the note by removing all "Song - Artist" patterns.
        The remaining text (handles, deliverables, etc.) becomes the description.
        """
        if not note or pd.isna(note):
            return ''

        # Remove all quoted "Song - Artist" patterns from the note
        description = SONG_ARTIST_PATTERN.sub('', note)

        # Clean up extra whitespace
        description = ' '.join(description.split()).strip()

        return description

    def pair_transactions(self, df: pd.DataFrame) -> List[Dict]:
        """
        Pair payment rows (Mobile Payment/BillPay/PreApproved) with their corresponding
        General Card Deposit rows using the Reference Txn ID field.

        For outgoing payments:
        - Mobile Payment / BillPay transaction / PreApproved Payment: Contains Note, Fee, recipient info
        - General Card Deposit: Contains actual USD amount charged

        The Reference Txn ID in the deposit row points to the payment row's Transaction ID.
        """
        # Filter deposits to only Completed (these show actual charges)
        # For payments, include both Completed and Pending since we pair by deposit
        deposits = df[(df['Type'] == 'General Card Deposit') & (df['Status'] == 'Completed')].copy()

        # Separate payment rows - include PreApproved Payment Bill User Payment for TikTok ads etc.
        payment_types = ['Mobile Payment', 'BillPay transaction', 'PreApproved Payment Bill User Payment']
        payments = df[df['Type'].isin(payment_types)].copy()

        # SAFETY CHECK: Warn about any unknown transaction types with negative amounts (outgoing money)
        # This catches new PayPal transaction types we haven't accounted for
        known_types = payment_types + ['General Card Deposit', 'General Currency Conversion']
        for _, row in df.iterrows():
            txn_type = row.get('Type', '')
            gross = str(row.get('Gross', '0')).replace(',', '')
            try:
                gross_val = float(gross)
            except ValueError:
                continue
            # If it's a negative amount (outgoing money) and not a known type, warn
            if gross_val < 0 and txn_type not in known_types:
                self.warnings.append(
                    f"UNKNOWN TRANSACTION TYPE with outgoing money: '{txn_type}' on "
                    f"{row.get('Date', 'unknown')} for {row.get('Gross', 'unknown')} "
                    f"[ID: {row.get('Transaction ID', 'unknown')}] - REVIEW MANUALLY"
                )

        # Build a lookup dict: Reference Txn ID -> deposit row
        # The deposit's "Reference Txn ID" points to the payment's "Transaction ID"
        deposit_lookup = {}
        for _, row in deposits.iterrows():
            ref_id = row.get('Reference Txn ID', '')
            if ref_id and not pd.isna(ref_id):
                deposit_lookup[ref_id] = row

        paired_transactions = []

        for _, payment in payments.iterrows():
            txn_id = payment.get('Transaction ID', '')

            # Skip if this is an incoming payment (positive Gross)
            gross = str(payment.get('Gross', '0')).replace(',', '')
            try:
                if float(gross) > 0:
                    continue
            except ValueError:
                continue

            # Look for matching deposit using the payment's Transaction ID
            deposit = deposit_lookup.get(txn_id)

            if deposit is None:
                self.errors.append(
                    f"No matching General Card Deposit found for transaction "
                    f"{payment.get('Date', 'unknown')} [ID: {txn_id}]"
                )
                continue

            # Extract amounts
            # Final Amount = USD amount from General Card Deposit (absolute value)
            final_amount = self.parse_amount(deposit.get('Gross', 0))

            # Additional Fees = Fee from the payment row (absolute value)
            additional_fees = self.parse_amount(payment.get('Fee', 0))

            # Initial Amount = Final Amount - Additional Fees
            initial_amount = final_amount - additional_fees

            # Determine if this is an invoice payment
            invoice_number = payment.get('Invoice Number', '')
            is_invoice = bool(invoice_number and not pd.isna(invoice_number))

            # Get the note and item title
            note = payment.get('Note', '')
            item_title = deposit.get('Item Title', '') or payment.get('Item Title', '')

            paired_transactions.append({
                'date': payment.get('Date', ''),
                'txn_id': txn_id,
                'note': note,
                'item_title': item_title,
                'is_invoice': is_invoice,
                'final_amount': final_amount,
                'additional_fees': additional_fees,
                'initial_amount': initial_amount,
                'type': payment.get('Type', '')
            })

        # SAFETY CHECK: Find orphan deposits (charges that weren't matched to any payment)
        # This catches cases where money was charged but we don't know what for
        processed_txn_ids = {txn['txn_id'] for txn in paired_transactions}
        for ref_id, deposit in deposit_lookup.items():
            if ref_id not in processed_txn_ids:
                # This deposit wasn't matched - money was charged but we missed the payment
                deposit_amount = deposit.get('Gross', 'unknown')
                deposit_date = deposit.get('Date', 'unknown')
                item_title = deposit.get('Item Title', '')
                self.errors.append(
                    f"UNMATCHED DEPOSIT (money charged but no matching payment found): "
                    f"{deposit_date} for ${deposit_amount} - Item: '{item_title}' "
                    f"[Ref ID: {ref_id}] - REVIEW MANUALLY AND ADD TO SHEET"
                )

        return paired_transactions

    def process_transaction(self, txn: Dict) -> List[List]:
        """
        Process a single paired transaction and return rows for the Google Sheet.

        Handles:
        - Invoice payments: Uses Item Title as description, leaves Artist/Song blank
        - Single payments: One row with parsed Song/Artist
        - Split payments: Multiple rows, dividing amounts equally among artists

        Returns a list of rows, where each row is a list of cell values.
        """
        rows = []

        # Format date as MM/DD/YYYY
        date_str = txn['date']
        try:
            # PayPal dates are already in MM/DD/YYYY format
            date_formatted = date_str
        except:
            date_formatted = date_str

        # Invoice handling: Use Item Title as description
        if txn['is_invoice']:
            # For invoices, use Item Title as Expense Description
            description = txn['item_title'] if txn['item_title'] else ''

            row = [
                date_formatted,          # Date
                '',                       # Payment Method (blank)
                'Cost',                   # Charge Type
                'No',                     # Split Charge
                '',                       # Artist Name (blank for invoice)
                '',                       # Song Name (blank for invoice)
                description,              # Expense Description
                f"${txn['initial_amount']:.2f}",  # Initial Amount
                f"${txn['additional_fees']:.2f}", # Additional Fees
                'FORMULA_PLACEHOLDER'     # Final Amount (will be replaced with formula)
            ]
            rows.append(row)
            return rows

        # Parse the note for Song/Artist pairs
        note = txn['note']
        song_artist_pairs = self.parse_note(note)
        description = self.extract_description(note, song_artist_pairs)

        # Handle case where note doesn't match "Song - Artist" format
        # A properly formatted note MUST have: quotes containing text with a dash
        # Example: "Midnight Drive - Nova Lane" @sample_creator 5 posts
        # If no valid "Song - Artist" pattern found, treat entire note as description
        if not song_artist_pairs:
            # Malformed note - put entire note text in Expense Description
            # Leave Artist Name and Song Name blank for manual review
            # Strip surrounding quotes if present for cleaner description
            note_text = note if note else ''
            if note_text.startswith('"') and note_text.endswith('"'):
                note_text = note_text[1:-1].strip()

            # Log warning for malformed notes (helps identify transactions needing review)
            if note:
                self.warnings.append(
                    f"Note doesn't match 'Song - Artist' format for transaction "
                    f"{txn['date']} [ID: {txn['txn_id']}]: {note}"
                )

            row = [
                date_formatted,
                '',                       # Payment Method (blank)
                'Cost',                   # Charge Type
                'No',                     # Split Charge
                '',                       # Artist Name (blank - malformed note)
                '',                       # Song Name (blank - malformed note)
                note_text,                # Expense Description (entire note)
                f"${txn['initial_amount']:.2f}",
                f"${txn['additional_fees']:.2f}",
                'FORMULA_PLACEHOLDER'
            ]
            rows.append(row)
            return rows

        # Determine if this is a split payment
        num_artists = len(song_artist_pairs)
        is_split = num_artists > 1

        # Calculate amounts per artist for split payments
        final_per_artist = txn['final_amount'] / num_artists
        fees_per_artist = txn['additional_fees'] / num_artists
        initial_per_artist = final_per_artist - fees_per_artist

        # Create a row for each artist
        for pair in song_artist_pairs:
            row = [
                date_formatted,
                '',                           # Payment Method (blank)
                'Cost',                       # Charge Type
                'Yes' if is_split else 'No',  # Split Charge
                pair['artist'],               # Artist Name
                pair['song'],                 # Song Name
                description,                  # Expense Description
                f"${initial_per_artist:.2f}", # Initial Amount (split if multiple)
                f"${fees_per_artist:.2f}",    # Additional Fees (split if multiple)
                'FORMULA_PLACEHOLDER'         # Final Amount (will be replaced with formula)
            ]
            rows.append(row)

        return rows

    def get_month_year(self, date_str: str) -> Tuple[str, int]:
        """
        Extract month name and year from a date string.
        Returns tuple of (month_name, year) e.g., ("December", 2025)
        """
        try:
            # Parse MM/DD/YYYY format
            date_obj = datetime.strptime(date_str, '%m/%d/%Y')
            month_name = date_obj.strftime('%B')
            year = date_obj.year
            return (month_name, year)
        except ValueError:
            return (None, None)

    def group_by_month(self, transactions: List[Dict]) -> Dict[str, List[Dict]]:
        """
        Group transactions by their month/year.
        Returns a dict with keys like "December 2025" and values as lists of transactions.
        """
        groups = defaultdict(list)

        for txn in transactions:
            month_name, year = self.get_month_year(txn['date'])
            if month_name and year:
                key = f"{month_name} {year}"
                groups[key].append(txn)

        return dict(groups)

    def get_or_create_sheet(self, sheet_name: str) -> Optional[gspread.Worksheet]:
        """
        Get an existing sheet or create a new one with headers.
        For new sheets: adds bold headers (row 1) and dropdown validation.
        Returns the worksheet object or None on failure.
        """
        try:
            # Try to get existing sheet
            worksheet = self.spreadsheet.worksheet(sheet_name)
            print(f"Found existing sheet '{sheet_name}'")
            return worksheet
        except gspread.exceptions.WorksheetNotFound:
            if not config.AUTO_CREATE_MONTHLY_SHEETS:
                print(f"Sheet '{sheet_name}' not found and AUTO_CREATE is disabled")
                return None

            print(f"Sheet '{sheet_name}' not found - creating new sheet...")

            # Create new sheet with headers
            for attempt in range(config.MAX_RETRIES):
                try:
                    worksheet = self.spreadsheet.add_worksheet(
                        title=sheet_name,
                        rows=1000,
                        cols=10
                    )

                    # Add header row (row 1 only, no instruction row)
                    worksheet.update(range_name='A1:J1', values=[HEADERS])

                    # Set up dropdown validation for columns B, C, D
                    self.setup_dropdown_validation(worksheet)

                    # Apply minimalist styling (header, dropdown columns, amount columns)
                    self.apply_sheet_styling(worksheet)

                    # Auto-resize columns to fit content
                    self.auto_resize_columns(worksheet)

                    # Reorder sheets so newest months appear first
                    self.reorder_sheets_chronologically()

                    print(f"Created sheet '{sheet_name}' with headers, dropdowns, and styling")
                    return worksheet

                except gspread.exceptions.APIError as e:
                    delay = config.RETRY_DELAY_BASE * (2 ** attempt)
                    print(f"API error, retrying in {delay}s... ({e})")
                    time.sleep(delay)

            print(f"Failed to create sheet after {config.MAX_RETRIES} attempts")
            return None

    def append_rows_to_sheet(self, worksheet: gspread.Worksheet, rows: List[List]) -> bool:
        """
        Append rows to the worksheet, starting after the last row with data.
        Converts FORMULA_PLACEHOLDER to actual formulas.

        For existing sheets:
        - Removes legacy instruction row (row 2) if present
        - Sets up dropdown validation for columns B, C, D
        - Applies minimalist styling (header colors, dropdown column colors, amount formatting)
        - Auto-resizes columns A-J to fit content
        - Reorders sheets chronologically (newest first)

        Returns True on success, False on failure.
        """
        if not rows:
            return True

        for attempt in range(config.MAX_RETRIES):
            try:
                # Check for and remove legacy instruction row if it exists
                self.remove_instruction_row_if_exists(worksheet)

                # Apply dropdown validation (for existing sheets)
                self.setup_dropdown_validation(worksheet)

                # Find the last row with data
                all_values = worksheet.get_all_values()
                next_row = len(all_values) + 1

                # Convert formula placeholders to actual formulas
                # Final Amount formula: =H[ROW]+I[ROW]
                for i, row in enumerate(rows):
                    row_num = next_row + i
                    if row[-1] == 'FORMULA_PLACEHOLDER':
                        row[-1] = f'=H{row_num}+I{row_num}'

                # Calculate the range for the update
                start_cell = f'A{next_row}'
                end_cell = f'J{next_row + len(rows) - 1}'
                range_notation = f'{start_cell}:{end_cell}'

                # Update with USER_ENTERED to interpret formulas
                worksheet.update(range_name=range_notation, values=rows, value_input_option='USER_ENTERED')

                # Calculate the new last row after appending
                new_last_row = next_row + len(rows) - 1

                # Apply minimalist styling (covers header and all data rows)
                self.apply_sheet_styling(worksheet, last_data_row=new_last_row)

                # Auto-resize columns to fit the new content
                self.auto_resize_columns(worksheet)

                # Reorder sheets so newest months appear first
                self.reorder_sheets_chronologically()

                return True

            except gspread.exceptions.APIError as e:
                delay = config.RETRY_DELAY_BASE * (2 ** attempt)
                print(f"API error while appending, retrying in {delay}s... ({e})")
                time.sleep(delay)

        print(f"Failed to append rows after {config.MAX_RETRIES} attempts")
        return False

    def process_csv_file(self, filepath: str) -> Dict[str, int]:
        """
        Process a single CSV file and upload to Google Sheets.
        Returns a dict with processing statistics.
        """
        print(f"\nProcessing CSV: {os.path.basename(filepath)}")

        # Read the CSV
        df = self.read_csv(filepath)
        if df is None:
            return {'error': 'Could not read CSV file'}

        # Pair transactions
        paired_transactions = self.pair_transactions(df)
        print(f"Found {len(paired_transactions)} transactions")

        # RECONCILIATION: Calculate and display total amounts for verification
        total_processed = sum(txn['final_amount'] for txn in paired_transactions)

        # Calculate total deposits in CSV for comparison
        deposits_in_csv = df[(df['Type'] == 'General Card Deposit') & (df['Status'] == 'Completed')]
        total_deposits = sum(self.parse_amount(row['Gross']) for _, row in deposits_in_csv.iterrows())

        print(f"\n--- RECONCILIATION CHECK ---")
        print(f"Total amount processed: ${total_processed:,.2f}")
        print(f"Total deposits in CSV:  ${total_deposits:,.2f}")
        if abs(total_processed - total_deposits) > 0.01:
            diff = total_deposits - total_processed
            print(f"⚠️  DISCREPANCY: ${diff:,.2f} in deposits not matched to payments!")
            print(f"    Check errors below for unmatched deposits.")
        else:
            print(f"✓ All deposits accounted for")
        print(f"----------------------------\n")

        if not paired_transactions:
            return {'transactions': 0, 'rows_added': 0}

        # Group by month
        monthly_groups = self.group_by_month(paired_transactions)
        print(f"Detected months: {', '.join(monthly_groups.keys())}")

        stats = {
            'transactions': len(paired_transactions),
            'regular_payments': 0,
            'invoice_payments': 0,
            'split_payments': 0,
            'total_rows': 0,
            'sheets_updated': []
        }

        # Process each month
        for month_name, transactions in monthly_groups.items():
            print(f"\nProcessing {month_name}:")

            # Get or create the monthly sheet
            worksheet = self.get_or_create_sheet(month_name)
            if worksheet is None:
                continue

            # Process all transactions for this month
            all_rows = []
            month_split_count = 0
            month_invoice_count = 0

            for txn in transactions:
                rows = self.process_transaction(txn)
                all_rows.extend(rows)

                if txn['is_invoice']:
                    month_invoice_count += 1
                elif len(rows) > 1:
                    month_split_count += 1

            month_regular = len(transactions) - month_invoice_count - month_split_count

            print(f"- {month_regular} regular payments")
            print(f"- {month_invoice_count} invoice payments")
            print(f"- {month_split_count} split payments ({len(all_rows)} total rows)")

            # Append rows to the sheet
            if self.append_rows_to_sheet(worksheet, all_rows):
                print(f"Successfully added {len(all_rows)} rows to sheet '{month_name}'")
                stats['regular_payments'] += month_regular
                stats['invoice_payments'] += month_invoice_count
                stats['split_payments'] += month_split_count
                stats['total_rows'] += len(all_rows)
                stats['sheets_updated'].append(month_name)
            else:
                print(f"Failed to add rows to sheet '{month_name}'")

        return stats

    def process_all_csvs(self, folder: str = None, specific_file: str = None):
        """
        Process all CSV files in the folder or a specific file.
        """
        if specific_file:
            files = [specific_file]
        else:
            folder = folder or config.CSV_FOLDER
            if not os.path.exists(folder):
                print(f"Error: CSV folder '{folder}' not found")
                return

            files = [
                os.path.join(folder, f)
                for f in os.listdir(folder)
                if f.lower().endswith('.csv')
            ]

        if not files:
            print("No CSV files found to process")
            return

        print(f"Found {len(files)} CSV file(s) to process")

        total_stats = {
            'total_rows': 0,
            'sheets_updated': set()
        }

        for filepath in files:
            stats = self.process_csv_file(filepath)
            if 'error' not in stats:
                total_stats['total_rows'] += stats.get('total_rows', 0)
                total_stats['sheets_updated'].update(stats.get('sheets_updated', []))

        # Print summary
        print(f"\n{'='*50}")
        print("Summary:")
        print(f"Total rows added: {total_stats['total_rows']}")
        print(f"Sheets updated: {', '.join(sorted(total_stats['sheets_updated']))}")

        # Print warnings (non-critical issues)
        if self.warnings:
            print(f"\n{'='*50}")
            print(f"⚠️  WARNINGS ({len(self.warnings)}):")
            print("These may need manual review:")
            for warning in self.warnings:
                print(f"  - {warning}")

        # Print errors (critical issues - money may be missing from sheet!)
        if self.errors:
            print(f"\n{'='*50}")
            print(f"🚨 ERRORS ({len(self.errors)}) - REVIEW IMMEDIATELY:")
            print("These transactions may be MISSING from your spreadsheet!")
            for error in self.errors:
                print(f"  - {error}")
            print(f"\n⚠️  ACTION REQUIRED: Review errors above and manually add missing transactions!")

        if not self.warnings and not self.errors:
            print(f"\n✓ No warnings or errors - all transactions processed successfully!")

        print(f"\n{'='*50}")
        print("Processing complete!")


def main():
    """Main entry point for the script."""
    processor = PayPalProcessor()

    # Authenticate with Google Sheets
    if not processor.authenticate():
        sys.exit(1)

    # Check for command line argument (specific file)
    if len(sys.argv) > 1:
        specific_file = sys.argv[1]
        if not os.path.exists(specific_file):
            print(f"Error: File '{specific_file}' not found")
            sys.exit(1)
        processor.process_all_csvs(specific_file=specific_file)
    else:
        processor.process_all_csvs()


if __name__ == '__main__':
    main()
