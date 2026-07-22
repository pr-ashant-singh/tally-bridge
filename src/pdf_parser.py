"""
PDF parser module for Zerodha Tax P&L PDF files.

The Zerodha PDF tax report splits trade data across 4 "bands" of pages:
  Band 1 (pages 1-N):   Symbol, ISIN, Entry Date, Exit Date, Quantity
  Band 2 (pages N+1-M): Buy Value, Sell Value, Profit, Period of Holding, FMV, Taxable Profit
  Band 3 (pages M+1-P): Turnover, Brokerage, Exchange Charges, IPFT, SEBI Charges, CGST
  Band 4 (pages P+1-Q): SGST, IGST, Stamp Duty, STT

After the trade bands, additional pages contain:
  - Equity and Non Equity summary
  - Other Debits and Credits (charges)
  - Equity Dividends
"""

from typing import Optional
import pdfplumber

from src.parser import ParsedData, Trade, Charge, Dividend


# Section headers that appear in the trade data
TRADE_SECTIONS = [
    "Equity - Intraday",
    "Equity - Short Term",
    "Equity - Long Term",
    "Equity - Buyback",
    "Non Equity",
    "Mutual Funds",
    "F&O",
    "Currency",
    "Commodity",
]


def parse_pdf(filepath: str, password: Optional[str] = None) -> ParsedData:
    """
    Parse a Zerodha Tax P&L PDF file.

    Args:
        filepath: Path to the .pdf file.
        password: Optional password for encrypted/protected PDFs.

    Returns:
        ParsedData object with all extracted information.

    Raises:
        ValueError: If the file format is not recognized.
        FileNotFoundError: If the file does not exist.
    """
    pdf = pdfplumber.open(filepath, password=password)
    data = ParsedData()

    # Extract client info from first page
    _parse_client_info(pdf, data)

    # Find band boundaries
    bands = _find_band_boundaries(pdf)

    # Parse trades from all 4 bands
    _parse_trades(pdf, data, bands)

    # Parse summary, charges, and dividends from remaining pages
    _parse_summary_and_extras(pdf, data, bands)

    pdf.close()
    return data


def _parse_client_info(pdf, data: ParsedData):
    """Extract client ID, name, PAN, and FY from the first page."""
    text = pdf.pages[0].extract_text() or ""
    for line in text.split("\n"):
        line = line.strip()
        if line.startswith("Client ID"):
            data.client_id = line.replace("Client ID", "").strip()
        elif line.startswith("Client Name"):
            data.client_name = line.replace("Client Name", "").strip()
        elif line.startswith("PAN"):
            data.pan = line.replace("PAN", "").strip()
        elif "Tradewise Exits from" in line:
            # Extract "2025-04-01 to 2026-03-31"
            parts = line.split("from")
            if len(parts) > 1:
                data.financial_year = parts[1].strip()


def _find_band_boundaries(pdf) -> dict:
    """
    Identify the page ranges for each data band.

    Returns dict with keys: 'band1', 'band2', 'band3', 'band4', 'extras_start'
    Each value is a tuple (start_page_idx, end_page_idx) exclusive end.
    """
    bands = {
        "band1": None,
        "band2": None,
        "band3": None,
        "band4": None,
        "extras_start": None,
    }

    # Strategy: Find the first page of each band by its header pattern
    band2_start = None
    band3_start = None
    band4_start = None
    extras_start = None

    for i, page in enumerate(pdf.pages):
        text = page.extract_text() or ""
        first_line = text.split("\n")[0].strip() if text.strip() else ""

        # Band 2 starts with "Buy Value Sell Value Profit..."
        if band2_start is None and "Buy Value" in first_line and "Sell Value" in first_line and "Profit" in first_line:
            band2_start = i

        # Band 3 starts with "Turnover Brokerage..." or "Sell Value Profit Turnover..."
        elif band2_start is not None and band3_start is None:
            if "Turnover" in first_line and "Brokerage" in first_line:
                band3_start = i
            elif "Sell Value" in first_line and "Turnover" in first_line:
                band3_start = i

        # Band 4 starts with "SGST IGST..." or "SEBI Charges CGST SGST..."
        elif band3_start is not None and band4_start is None:
            if "SGST" in first_line and ("IGST" in first_line or "Stamp" in first_line):
                band4_start = i
            elif "SEBI" in first_line and "CGST" in first_line and "SGST" in first_line:
                band4_start = i

        # Extras start at summary page (Client ID header after band 4)
        elif band4_start is not None and extras_start is None:
            if "Client ID" in first_line or "Taxpnl Statement" in first_line:
                extras_start = i

    # Band 1 is always from page 0 to band2_start
    bands["band1"] = (0, band2_start)
    bands["band2"] = (band2_start, band3_start)
    bands["band3"] = (band3_start, band4_start)
    bands["band4"] = (band4_start, extras_start or len(pdf.pages))
    bands["extras_start"] = extras_start

    return bands


def _parse_trades(pdf, data: ParsedData, bands: dict):
    """Parse trade data from all 4 bands and combine into Trade objects."""
    # Extract Band 1: Symbol, ISIN, Entry Date, Exit Date, Quantity + section info
    band1_data = _extract_band1(pdf.pages[bands["band1"][0]:bands["band1"][1]])

    # Extract numeric bands
    band2_rows = _extract_numeric_rows(pdf.pages[bands["band2"][0]:bands["band2"][1]])
    band3_rows = _extract_numeric_rows(pdf.pages[bands["band3"][0]:bands["band3"][1]])
    band4_rows = _extract_numeric_rows(pdf.pages[bands["band4"][0]:bands["band4"][1]])

    # Verify alignment
    n = len(band1_data)
    if len(band2_rows) != n or len(band3_rows) != n:
        # Try to reconcile - use minimum
        n = min(len(band1_data), len(band2_rows), len(band3_rows), len(band4_rows))

    # Combine into Trade objects
    for i in range(n):
        b1 = band1_data[i]
        b2 = band2_rows[i] if i < len(band2_rows) else []
        b3 = band3_rows[i] if i < len(band3_rows) else []
        b4 = band4_rows[i] if i < len(band4_rows) else []

        trade = _build_trade(b1, b2, b3, b4)
        if trade:
            data.trades.append(trade)


def _extract_band1(pages) -> list:
    """
    Extract identifier data from Band 1 pages.

    Returns list of dicts with keys: section, symbol, isin, entry_date, exit_date, quantity
    """
    rows = []
    current_section = None

    for page in pages:
        text = page.extract_text() or ""
        for line in text.split("\n"):
            line = line.strip()
            if not line:
                continue

            # Skip non-data lines
            if any(line.startswith(skip) for skip in [
                "View Zerodha", "Client ID", "Client Name", "PAN",
                "Tradewise Exits", "Symbol ISIN", "Symbol Entry",
            ]):
                continue

            # Check for section headers
            found_section = False
            for sec in TRADE_SECTIONS:
                if line == sec:
                    current_section = sec
                    found_section = True
                    break
            if found_section:
                continue

            # Parse data row
            parts = line.split()
            if not current_section:
                continue

            # A valid trade row should have: Symbol ISIN EntryDate ExitDate Quantity
            # But some symbols contain special chars, so be flexible
            # ISIN is always 12 chars starting with IN
            if len(parts) >= 5:
                # Find ISIN position (INExxxxx or INFxxxxx pattern)
                isin_idx = None
                for idx, part in enumerate(parts):
                    if len(part) == 12 and (part.startswith("IN")):
                        isin_idx = idx
                        break

                if isin_idx is not None:
                    symbol = " ".join(parts[:isin_idx])
                    isin = parts[isin_idx]
                    entry_date = parts[isin_idx + 1] if isin_idx + 1 < len(parts) else ""
                    exit_date = parts[isin_idx + 2] if isin_idx + 2 < len(parts) else ""
                    quantity = parts[isin_idx + 3] if isin_idx + 3 < len(parts) else "0"

                    rows.append({
                        "section": current_section,
                        "symbol": symbol,
                        "isin": isin,
                        "entry_date": entry_date,
                        "exit_date": exit_date,
                        "quantity": quantity,
                    })
                elif len(parts) >= 5:
                    # Fallback: assume format is Symbol EntryDate ExitDate Quantity ...
                    # (some Non Equity entries don't have ISIN)
                    rows.append({
                        "section": current_section,
                        "symbol": parts[0],
                        "isin": "",
                        "entry_date": parts[1] if len(parts) > 1 else "",
                        "exit_date": parts[2] if len(parts) > 2 else "",
                        "quantity": parts[3] if len(parts) > 3 else "0",
                    })

    return rows


def _extract_numeric_rows(pages) -> list:
    """Extract numeric data rows from a band, skipping header lines."""
    rows = []
    for page in pages:
        text = page.extract_text() or ""
        for line in text.split("\n"):
            line = line.strip()
            if not line:
                continue

            parts = line.split()
            if not parts:
                continue

            # Skip lines that start with alphabetic text (headers)
            first_char = parts[0][0] if parts[0] else ""
            if first_char.isalpha():
                continue

            # Must start with a number (possibly negative)
            try:
                float(parts[0])
                rows.append(parts)
            except ValueError:
                continue

    return rows


def _build_trade(b1: dict, b2: list, b3: list, b4: list) -> Optional[Trade]:
    """Combine data from all 4 bands into a single Trade object."""

    def safe_float(lst, idx, default=0.0):
        if idx >= len(lst):
            return default
        try:
            return float(lst[idx])
        except (ValueError, TypeError):
            return default

    def safe_int(lst, idx, default=0):
        if idx >= len(lst):
            return default
        try:
            return int(float(lst[idx]))
        except (ValueError, TypeError):
            return default

    try:
        quantity = float(b1["quantity"])
    except (ValueError, TypeError):
        quantity = 0.0

    return Trade(
        symbol=b1["symbol"],
        isin=b1["isin"],
        entry_date=b1["entry_date"],
        exit_date=b1["exit_date"],
        quantity=quantity,
        buy_value=safe_float(b2, 0),
        sell_value=safe_float(b2, 1),
        profit=safe_float(b2, 2),
        period_of_holding=safe_int(b2, 3),
        fair_market_value=safe_float(b2, 4),
        taxable_profit=safe_float(b2, 5),
        turnover=safe_float(b3, 0),
        brokerage=safe_float(b3, 1),
        exchange_charges=safe_float(b3, 2),
        ipft=safe_float(b3, 3),
        sebi_charges=safe_float(b3, 4),
        cgst=safe_float(b3, 5),
        sgst=safe_float(b4, 0),
        igst=safe_float(b4, 1),
        stamp_duty=safe_float(b4, 2),
        stt=safe_float(b4, 3),
        trade_type=b1["section"],
    )


def _parse_summary_and_extras(pdf, data: ParsedData, bands: dict):
    """Parse summary, charges, and dividends from pages after the trade bands."""
    extras_start = bands.get("extras_start")
    if extras_start is None:
        return

    # Concatenate all text from extras pages
    all_text = ""
    for page in pdf.pages[extras_start:]:
        text = page.extract_text() or ""
        all_text += text + "\n"

    lines = all_text.split("\n")

    # Parse summary
    _parse_summary_from_lines(lines, data)

    # Parse charges
    _parse_charges_from_lines(lines, data)

    # Parse dividends
    _parse_dividends_from_lines(lines, data)


def _parse_summary_from_lines(lines: list, data: ParsedData):
    """Parse the Realized Profit Breakdown from text.
    
    Only takes the first match for each value to avoid overwriting
    with values from later sections (e.g., Mutual Funds summary).
    """
    import re

    found = {"intraday": False, "short_term": False, "long_term": False, "non_equity": False}

    for line in lines:
        line = line.strip()
        # Handle concatenated format like "Intraday/Speculative profit-7.8"
        # and spaced format like "Short Term profit -30406.06"
        if not found["intraday"] and "Intraday" in line and "profit" in line.lower():
            val = _extract_trailing_number(line)
            if val is not None:
                data.intraday_profit = val
                found["intraday"] = True
        elif not found["short_term"] and "Short Term profit" in line and "Equity" not in line and "Debt" not in line:
            val = _extract_trailing_number(line)
            if val is not None:
                data.short_term_profit = val
                found["short_term"] = True
        elif not found["long_term"] and "Long Term profit" in line and "Equity" not in line and "Debt" not in line:
            val = _extract_trailing_number(line)
            if val is not None:
                data.long_term_profit = val
                found["long_term"] = True
        elif not found["non_equity"] and "Non Equity profit" in line:
            val = _extract_trailing_number(line)
            if val is not None:
                data.non_equity_profit = val
                found["non_equity"] = True

        # Stop early if all found
        if all(found.values()):
            break


def _extract_trailing_number(line: str) -> Optional[float]:
    """Extract a number from the end of a line (handles 'profit-7.8' and 'profit 6714.8')."""
    import re

    # Try splitting by space and getting last value
    parts = line.split()
    if parts:
        try:
            return float(parts[-1])
        except ValueError:
            pass

    # Try finding number pattern at end (handles concatenated like "profit-7.8")
    match = re.search(r"(-?\d+\.?\d*)$", line.replace(" ", ""))
    if match:
        return float(match.group(1))
    return None


def _parse_charges_from_lines(lines: list, data: ParsedData):
    """Parse Other Debits and Credits section."""
    in_charges = False
    current_segment = None
    segments = ["Equity", "Mutual Funds", "F&O", "Currency", "Commodity"]

    for line in lines:
        line = line.strip()

        if "Other Debits and Credits" in line:
            in_charges = True
            continue

        if not in_charges:
            continue

        # Stop at Equity Dividends or next major section
        if "Equity Dividends" in line or "Open Positions" in line or "Ledger Balances" in line:
            break

        # Check for segment header
        if line in segments:
            current_segment = line
            continue

        # Skip the header row
        if line.startswith("Particulars"):
            continue

        # Skip client info lines
        if line.startswith("Client") or line.startswith("PAN"):
            continue

        # Parse charge data row
        if current_segment and line:
            charge = _parse_charge_line(line, current_segment)
            if charge:
                data.charges.append(charge)


def _parse_charge_line(line: str, segment: str) -> Optional[Charge]:
    """Parse a single charge line like 'AMC for Demat Account... 2026-02-19 88.5 0'."""
    import re

    # Pattern: text description, then date (YYYY-MM-DD), then 1-2 numbers
    # The date pattern helps us split description from data
    date_pattern = r"(\d{4}-\d{2}-\d{2})"
    match = re.search(date_pattern, line)

    if not match:
        return None

    date_start = match.start()
    particulars = line[:date_start].strip()
    remainder = line[date_start:].strip()

    parts = remainder.split()
    if len(parts) < 3:
        return None

    posting_date = parts[0]
    try:
        debit = float(parts[1])
        credit = float(parts[2]) if len(parts) > 2 else 0.0
    except ValueError:
        return None

    return Charge(
        particulars=particulars,
        posting_date=posting_date,
        debit=debit,
        credit=credit,
        segment=segment,
    )


def _parse_dividends_from_lines(lines: list, data: ParsedData):
    """Parse Equity Dividends section.
    
    The PDF splits dividend data across pages:
    - Pages with Symbol, ISIN, Ex-date, Quantity, Dividend Per Share
    - Separate page(s) with just Net Dividend Amount values
    """
    in_dividends = False
    in_net_amounts = False
    dividend_rows = []
    net_amounts = []

    for line in lines:
        line = line.strip()

        if "Equity Dividends from" in line:
            in_dividends = True
            continue

        if not in_dividends:
            continue

        # Stop at next section
        if "Ledger Balances" in line or "Open Positions" in line:
            break

        # Skip headers and client info
        if line.startswith("Symbol ISIN") or line.startswith("Client") or line.startswith("PAN"):
            continue

        # Net Dividend Amount column header - switch to amount mode
        if "Net Dividend Amount" in line:
            in_net_amounts = True
            continue

        # Total line
        if "Total Dividend" in line:
            parts = line.split()
            if parts:
                try:
                    data.total_dividend = float(parts[-1])
                except ValueError:
                    pass
            continue

        # Skip info lines
        if "Dividends are credited" in line:
            continue

        if in_net_amounts:
            # These are just numbers (net amounts), one per line
            try:
                val = float(line)
                net_amounts.append(val)
            except ValueError:
                continue
        else:
            # Parse dividend identifier row
            parts = line.split()
            if len(parts) >= 5:
                # Find ISIN
                isin_idx = None
                for idx, part in enumerate(parts):
                    if len(part) == 12 and part.startswith("IN"):
                        isin_idx = idx
                        break

                if isin_idx is not None:
                    symbol = " ".join(parts[:isin_idx])
                    isin = parts[isin_idx]
                    ex_date = parts[isin_idx + 1] if isin_idx + 1 < len(parts) else ""
                    quantity = parts[isin_idx + 2] if isin_idx + 2 < len(parts) else "0"
                    dps = parts[isin_idx + 3] if isin_idx + 3 < len(parts) else "0"
                    dividend_rows.append({
                        "symbol": symbol,
                        "isin": isin,
                        "ex_date": ex_date,
                        "quantity": quantity,
                        "dps": dps,
                    })

    # Combine identifier rows with net amounts
    # The last net_amount value might be the total if it matches data.total_dividend
    if net_amounts and len(net_amounts) > len(dividend_rows):
        # Last value is likely the total
        if not data.total_dividend and net_amounts:
            data.total_dividend = net_amounts[-1]
        net_amounts = net_amounts[:len(dividend_rows)]

    for i, row in enumerate(dividend_rows):
        try:
            quantity = float(row["quantity"])
            dps = float(row["dps"])
            net_amount = net_amounts[i] if i < len(net_amounts) else quantity * dps

            data.dividends.append(Dividend(
                symbol=row["symbol"],
                isin=row["isin"],
                ex_date=row["ex_date"],
                quantity=quantity,
                dividend_per_share=dps,
                net_amount=net_amount,
            ))
        except (ValueError, TypeError):
            continue
