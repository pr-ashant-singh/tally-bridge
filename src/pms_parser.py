"""
PMS (Portfolio Management Service) parser for SageOne monthly statements.

Parses password-protected SageOne PMS PDFs which contain:
  - Portfolio Performance (summary returns)
  - Portfolio Appraisal (current holdings)
  - Transaction Statement (buys/sells for the period)
  - Statement of Capital Gain/Loss (realized gains with purchase info)
  - Bank Book (cash movements, expenses like custody/fund accounting fees)

The PDF may contain multiple portfolios (e.g., Smallcap + Core).
Each portfolio has its own set of sections.

Output is mapped to the existing ParsedData structure for Tally generation.
"""

import re
from typing import Optional
from datetime import datetime

import pdfplumber

from src.parser import ParsedData, Trade, Charge


def is_pms_statement(filepath: str, password: Optional[str] = None) -> bool:
    """
    Check if a PDF is a SageOne PMS statement by looking at the first page.

    Args:
        filepath: Path to the PDF file.
        password: Optional password for encrypted PDFs.

    Returns:
        True if this appears to be a SageOne PMS statement.
    """
    try:
        pdf = pdfplumber.open(filepath, password=password)
        text = pdf.pages[0].extract_text() or ""
        pdf.close()
        # SageOne PMS statements have a Table of Contents with these sections
        return (
            "Portfolio Performance" in text
            and "Statement of Capital Gain" in text
        ) or (
            "SAGEONE" in text.upper()
            and "Portfolio" in text
        )
    except Exception:
        return False


def parse_pms(filepath: str, password: Optional[str] = None) -> ParsedData:
    """
    Parse a SageOne PMS monthly statement PDF.

    Args:
        filepath: Path to the PDF file.
        password: Optional password for encrypted PDFs.

    Returns:
        ParsedData object with extracted capital gains, transactions, and charges.

    Raises:
        ValueError: If the file format is not recognized as a PMS statement.
    """
    pdf = pdfplumber.open(filepath, password=password)
    data = ParsedData()

    # Extract client info from first content page
    _parse_client_info(pdf, data)

    # Find all sections across pages
    sections = _identify_sections(pdf)

    # Parse capital gains (most important for Tally)
    _parse_capital_gains(pdf, data, sections)

    # Parse transactions (for reference/narration)
    transactions = _parse_transactions(pdf, sections)

    # Parse bank book for charges/expenses
    _parse_bank_charges(pdf, data, sections)

    # Enrich trades with transaction-level brokerage/STT from transaction statement
    _enrich_trades_with_transaction_data(data, transactions)

    pdf.close()
    return data


def _parse_client_info(pdf, data: ParsedData):
    """Extract account info from the first content page (page 2, index 1)."""
    # Try pages 1-4 for account info
    for page_idx in range(min(5, len(pdf.pages))):
        text = pdf.pages[page_idx].extract_text() or ""
        lines = text.split("\n")

        for line in lines:
            line = line.strip()

            # Account number and client name on same line
            if not data.client_id and (line.startswith("Account :") or line.startswith("Account:")):
                parts = line.replace("Account :", "").replace("Account:", "").strip()
                # Format: "100389 TWOPIR CONSULTING PRIVATE LIMITED"
                match = re.match(r"(\d+)\s+(.*)", parts)
                if match:
                    data.client_id = match.group(1)
                    data.client_name = match.group(2)

            # Date range (From DD/MM/YYYY to DD/MM/YYYY)
            if not data.financial_year and line.startswith("From ") and " to " in line:
                date_range = line.replace("From ", "").strip()
                # Convert DD/MM/YYYY to YYYY-MM-DD format for compatibility with generator
                range_parts = date_range.split(" to ")
                if len(range_parts) == 2:
                    start = _convert_date(range_parts[0].strip())
                    end = _convert_date(range_parts[1].strip())
                    data.financial_year = f"{start} to {end}"
                else:
                    data.financial_year = date_range

        if data.client_id and data.financial_year:
            break


def _identify_sections(pdf) -> dict:
    """
    Identify page ranges for each section type.

    Returns dict mapping section name to list of page indices.
    """
    sections = {
        "portfolio_performance": [],
        "portfolio_appraisal": [],
        "transaction_statement": [],
        "capital_gain": [],
        "bank_book": [],
    }

    for i, page in enumerate(pdf.pages):
        text = (page.extract_text() or "")[:200]  # Just check first few lines

        if "Statement of Capital Gain" in text:
            sections["capital_gain"].append(i)
        elif "Transaction Statement" in text:
            sections["transaction_statement"].append(i)
        elif "Bank Book" in text:
            sections["bank_book"].append(i)
        elif "Portfolio Performance" in text:
            sections["portfolio_performance"].append(i)
        elif "Portfolio Appraisal" in text:
            sections["portfolio_appraisal"].append(i)

    return sections


def _parse_capital_gains(pdf, data: ParsedData, sections: dict):
    """
    Parse Statement of Capital Gain/Loss pages.

    Each row format (single string from table extraction):
    "SECURITY_NAME-ISIN SALE_DATE QTY SALE_RATE SALE_AMT PURCHASE_DATE PURCHASE_RATE PRICE_31JAN18 PURCHASE_AMT EFFECTIVE_COST DAYS_HELD ST_GAIN LT_GAIN EFFECTIVE_GAIN"
    """
    capital_gain_pages = sections.get("capital_gain", [])
    if not capital_gain_pages:
        return

    for page_idx in capital_gain_pages:
        page = pdf.pages[page_idx]
        tables = page.extract_tables()

        for table in tables:
            for row in table:
                if not row or not row[0]:
                    continue

                cell = row[0].strip()

                # Skip headers, section labels, and summary rows
                if _is_capital_gain_header(cell):
                    continue

                # Skip "Total" and summary lines
                if cell.startswith("Total") or cell.startswith("Capital Gain"):
                    continue
                if cell.startswith("ST Gain") or cell.startswith("LT Gain"):
                    # Parse summary
                    _parse_gain_summary_line(cell, data)
                    continue

                # Parse trade data row
                trade = _parse_capital_gain_row(cell)
                if trade:
                    data.trades.append(trade)


def _is_capital_gain_header(cell: str) -> bool:
    """Check if a cell contains header text rather than data."""
    headers = [
        "Sale Purchase",
        "Security Sale Date",
        "Listed Shares",
        "(S)",
        "(M)",
        "01/04-",
    ]
    return any(h in cell for h in headers)


def _parse_gain_summary_line(cell: str, data: ParsedData):
    """Parse summary lines like 'ST Gain/Loss 0.00 0.00 0.00 0.00 0.00 0.00'."""
    parts = cell.split()
    if not parts:
        return

    # Get the "Total" value (last number)
    try:
        total = _parse_number(parts[-1])
    except (ValueError, IndexError):
        return

    if "ST Gain" in cell or "ST Loss" in cell:
        data.short_term_profit += total
    elif "LT Gain" in cell or "LT Loss" in cell:
        data.long_term_profit += total


def _parse_capital_gain_row(cell: str) -> Optional[Trade]:
    """
    Parse a capital gain row like:
    "FIEM INDUSTRIES LTD-INE737H01014 16/12/25 340 2,333.132 793,264.88 17/06/22 1,020.8157 942.05 347,077.34 347,077.34 1278 0.00 446,187.54 446,187.54"

    Format:
    SECURITY_NAME-ISIN SALE_DATE QTY SALE_RATE SALE_AMOUNT PURCHASE_DATE PURCHASE_RATE PRICE_31JAN18 PURCHASE_AMT EFFECTIVE_COST DAYS_HELD ST_GAIN LT_GAIN EFFECTIVE_GAIN
    """
    # Find the ISIN pattern (INExxxxxxxxx or INFxxxxxxxxx)
    isin_match = re.search(r"(IN[A-Z]\w{9})", cell)
    if not isin_match:
        return None

    isin = isin_match.group(1)
    isin_end = isin_match.end()

    # Security name is everything before the ISIN (minus the dash)
    security_part = cell[:isin_match.start()].rstrip("-").strip()

    # Everything after ISIN is the numeric data
    remainder = cell[isin_end:].strip()
    parts = remainder.split()

    if len(parts) < 13:
        return None

    try:
        sale_date_str = parts[0]       # 16/12/25
        quantity = _parse_number(parts[1])
        sale_rate = _parse_number(parts[2])
        sale_amount = _parse_number(parts[3])
        purchase_date_str = parts[4]   # 17/06/22
        purchase_rate = _parse_number(parts[5])
        # parts[6] = price on 31-Jan-18 (FMV)
        fair_market_value = _parse_number(parts[6])
        purchase_amount = _parse_number(parts[7])
        effective_cost = _parse_number(parts[8])
        days_held = int(_parse_number(parts[9]))
        st_gain = _parse_number(parts[10])
        lt_gain = _parse_number(parts[11])
        effective_gain = _parse_number(parts[12]) if len(parts) > 12 else lt_gain

    except (ValueError, IndexError):
        return None

    # Determine trade type based on holding period and gain columns
    profit = st_gain if st_gain != 0 else lt_gain
    if days_held <= 365:
        trade_type = "Equity - Short Term"
    else:
        trade_type = "Equity - Long Term"

    # Convert dates from DD/MM/YY to YYYY-MM-DD
    entry_date = _convert_date(purchase_date_str)
    exit_date = _convert_date(sale_date_str)

    return Trade(
        symbol=security_part,
        isin=isin,
        entry_date=entry_date,
        exit_date=exit_date,
        quantity=quantity,
        buy_value=purchase_amount,
        sell_value=sale_amount,
        profit=profit,
        period_of_holding=days_held,
        fair_market_value=fair_market_value,
        taxable_profit=effective_gain,
        turnover=sale_amount,
        brokerage=0.0,  # Will be enriched from transaction statement
        exchange_charges=0.0,
        ipft=0.0,
        sebi_charges=0.0,
        cgst=0.0,
        sgst=0.0,
        igst=0.0,
        stamp_duty=0.0,
        stt=0.0,
        trade_type=trade_type,
    )


def _parse_transactions(pdf, sections: dict) -> list:
    """
    Parse the Transaction Statement for brokerage and STT data.

    Returns list of dicts with transaction-level details.
    """
    transactions = []
    txn_pages = sections.get("transaction_statement", [])

    for page_idx in txn_pages:
        page = pdf.pages[page_idx]
        tables = page.extract_tables()

        for table in tables:
            for row in table:
                if not row or not row[0]:
                    continue

                cell = row[0].strip()

                # Skip headers and section labels
                if _is_transaction_header(cell):
                    continue

                # Parse transaction row
                txn = _parse_transaction_row(cell)
                if txn:
                    transactions.append(txn)

    return transactions


def _is_transaction_header(cell: str) -> bool:
    """Check if this is a transaction statement header or section label."""
    headers = [
        "Transaction Description",
        "Current Period",
        "Shares - Listed",
        "TRANSACTION STATEMENT",
        "Sell",  # Summary row starting with just "Sell" followed by totals
        "Buy",   # Summary row
    ]
    # The summary table has "Sell X X X X" format (all numbers after)
    if cell.startswith("Sell ") or cell.startswith("Buy "):
        parts = cell.split()
        if len(parts) >= 3:
            # Check if it's a data row (has a date) vs summary
            try:
                # Transaction rows have: Buy/Sell DD/MM/YYYY ...
                if re.match(r"\d{2}/\d{2}/\d{4}", parts[1]):
                    return False  # It's a data row
            except (IndexError, ValueError):
                pass
        return True  # It's a header/summary
    return any(h in cell for h in headers[:4])


def _parse_transaction_row(cell: str) -> Optional[dict]:
    """
    Parse a transaction row like:
    "Buy 11/12/2025 12/12/2025 SGMART NSE 1,044.000 327.6341 0.2621 342.05 342,665.68"
    "Sell 16/12/2025 17/12/2025 FIEM INDUSTRIES LTD NSE 340.000 2,335.00 1.868 793.90 792,470.98"

    Format:
    TYPE TRAN_DATE SETTLE_DATE SECURITY EXCHANGE QUANTITY UNIT_PRICE BROKERAGE STT SETTLEMENT_AMOUNT
    """
    parts = cell.split()
    if len(parts) < 8:
        return None

    txn_type = parts[0]
    if txn_type not in ("Buy", "Sell"):
        return None

    # Validate date format for tran_date
    if not re.match(r"\d{2}/\d{2}/\d{4}", parts[1]):
        return None

    tran_date = parts[1]
    settle_date = parts[2]

    # Find the exchange (NSE/BSE) - it's before the numbers start
    exchange_idx = None
    for idx in range(3, len(parts)):
        if parts[idx] in ("NSE", "BSE"):
            exchange_idx = idx
            break

    if exchange_idx is None:
        return None

    security = " ".join(parts[3:exchange_idx])
    exchange = parts[exchange_idx]

    # After exchange: Quantity, UnitPrice, Brokerage, STT, SettlementAmount
    numeric_parts = parts[exchange_idx + 1:]
    if len(numeric_parts) < 5:
        return None

    try:
        quantity = _parse_number(numeric_parts[0])
        unit_price = _parse_number(numeric_parts[1])
        brokerage = _parse_number(numeric_parts[2])
        stt = _parse_number(numeric_parts[3])
        settlement_amount = _parse_number(numeric_parts[4])
    except (ValueError, IndexError):
        return None

    return {
        "type": txn_type,
        "tran_date": tran_date,
        "settle_date": settle_date,
        "security": security,
        "exchange": exchange,
        "quantity": quantity,
        "unit_price": unit_price,
        "brokerage": brokerage,
        "stt": stt,
        "settlement_amount": settlement_amount,
    }


def _parse_bank_charges(pdf, data: ParsedData, sections: dict):
    """
    Parse Bank Book pages for expense entries (Custody Charges, Fund Accounting Fees, etc.)
    """
    bank_pages = sections.get("bank_book", [])

    for page_idx in bank_pages:
        text = pdf.pages[page_idx].extract_text() or ""
        lines = text.split("\n")

        # Determine which portfolio this belongs to
        portfolio = ""
        for line in lines[:5]:
            if "SAGEONE" in line.upper():
                portfolio = line.strip()
                break

        for line in lines:
            line = line.strip()
            if not line:
                continue

            # Look for expense lines (they have "Expense" in them)
            if "Expense" in line:
                charge = _parse_bank_expense_line(line, portfolio)
                if charge:
                    data.charges.append(charge)


def _parse_bank_expense_line(line: str, portfolio: str) -> Optional[Charge]:
    """
    Parse a bank book expense line like:
    "26/12/2025 DIRECT Custody Charges - Cash Expense 0.00 0.00 1,750.55 0.00 3,876,148.93"

    Format:
    DATE TRAN_ACCOUNT DESCRIPTION Buy/Sell_Amount Income Expenses Dep/With Balance
    """
    # Find date at the start
    date_match = re.match(r"(\d{2}/\d{2}/\d{4})", line)
    if not date_match:
        return None

    posting_date = date_match.group(1)
    remainder = line[date_match.end():].strip()

    # Find "Expense" keyword and extract description before it
    expense_idx = remainder.find("Expense")
    if expense_idx < 0:
        return None

    # Description is between the account code and "Expense"
    desc_part = remainder[:expense_idx].strip()
    # Remove the account code prefix (e.g., "DIRECT")
    # The description typically has format: "ACCOUNT Description - Cash"
    desc_match = re.match(r"\w+\s+(.*?)\s*-\s*Cash", desc_part)
    if desc_match:
        particulars = desc_match.group(1).strip()
    else:
        particulars = desc_part

    # Numbers after "Expense" - format: Buy/Sell Income Expenses Dep/With Balance
    numbers_part = remainder[expense_idx + len("Expense"):].strip()
    numbers = numbers_part.split()

    if len(numbers) < 3:
        return None

    try:
        # Expenses column is the 3rd value (index 2)
        expense_amount = _parse_number(numbers[2])
    except (ValueError, IndexError):
        return None

    if expense_amount <= 0:
        return None

    # Convert date from DD/MM/YYYY to YYYY-MM-DD
    date_parts = posting_date.split("/")
    if len(date_parts) == 3:
        formatted_date = f"{date_parts[2]}-{date_parts[1]}-{date_parts[0]}"
    else:
        formatted_date = posting_date

    segment = portfolio if portfolio else "PMS"

    return Charge(
        particulars=particulars,
        posting_date=formatted_date,
        debit=expense_amount,
        credit=0.0,
        segment=segment,
    )


def _enrich_trades_with_transaction_data(data: ParsedData, transactions: list):
    """
    Match capital gain trades with transaction statement data to add
    brokerage and STT information.
    """
    # Build a lookup: (security_name_normalized, date, quantity) -> transaction
    txn_lookup = {}
    for txn in transactions:
        if txn["type"] == "Sell":
            key = (
                _normalize_security_name(txn["security"]),
                txn["tran_date"],
                txn["quantity"],
            )
            txn_lookup[key] = txn

    for trade in data.trades:
        # Convert exit_date back to DD/MM/YYYY for matching
        exit_date_ddmmyyyy = _convert_date_to_ddmmyyyy(trade.exit_date)
        key = (
            _normalize_security_name(trade.symbol),
            exit_date_ddmmyyyy,
            trade.quantity,
        )

        if key in txn_lookup:
            txn = txn_lookup[key]
            trade.brokerage = txn["brokerage"]
            trade.stt = txn["stt"]


def _normalize_security_name(name: str) -> str:
    """Normalize security name for matching."""
    return name.upper().strip()


def _convert_date(date_str: str) -> str:
    """Convert DD/MM/YY or DD/MM/YYYY to YYYY-MM-DD format."""
    if not date_str:
        return ""

    parts = date_str.split("/")
    if len(parts) != 3:
        return date_str

    day, month, year = parts
    if len(year) == 2:
        # Assume 20xx for years < 50, 19xx otherwise
        year_int = int(year)
        year = f"20{year}" if year_int < 50 else f"19{year}"

    return f"{year}-{month}-{day}"


def _convert_date_to_ddmmyyyy(date_str: str) -> str:
    """Convert YYYY-MM-DD to DD/MM/YYYY format."""
    if not date_str:
        return ""

    parts = date_str.split("-")
    if len(parts) != 3:
        return date_str

    year, month, day = parts
    return f"{day}/{month}/{year}"


def _parse_number(value: str) -> float:
    """Parse a number string, removing commas."""
    if not value:
        return 0.0
    return float(value.replace(",", ""))
