# TallyBridge

Convert Zerodha Tax P&L Excel reports into Tally-compatible import files — zero install, single `.exe`.

## What It Does

Takes your Zerodha Tax P&L report (`taxpnl-*.xlsx`) and generates 4 separate Excel files ready for Tally import:

| Output File | Contents |
|---|---|
| `Profit_Entries_FY20XX-XX.xlsx` | All trades with positive P&L (capital gains) |
| `Loss_Entries_FY20XX-XX.xlsx` | All trades with negative P&L (capital losses) |
| `Charges_FY20XX-XX.xlsx` | DP charges, AMC, smallcase fees, etc. |
| `Dividends_FY20XX-XX.xlsx` | Dividend income entries |

Each file uses Tally's import format: **Date | Voucher Type | Particulars | Ledger Name | Debit | Credit | Narration**

## Download & Usage

### For Windows Users (No Installation Required)

1. Download `TallyBridge.exe` from [Releases](../../releases)
2. Double-click to run
3. Select your Zerodha Tax P&L Excel file
4. Click **Generate Tally Files**
5. Import the output files into Tally

> **Note:** On first run, Windows may show a SmartScreen warning. Click "More info" → "Run anyway". This is normal for unsigned apps.

### Run from Source (Mac/Linux/Windows)

```bash
# Clone
git clone https://github.com/pr-ashant-singh/tally-bridge.git
cd tally-bridge

# Install dependencies
pip install -r requirements.txt

# Run
python main.py
```

## Supported Input Format

Zerodha Tax P&L report downloaded from:
**Console → Reports → Tax P&L → Download .xlsx**

The file typically has a name like: `taxpnl-CLIENTID-YYYY_YYYY-Q1-Q4.xlsx`

### Sheets Parsed

- **Tradewise Exits** — Individual trade-level data (Intraday, Short Term, Long Term, Non-Equity)
- **Equity and Non Equity** — P&L summary (used for cross-validation)
- **Other Debits and Credits** — DP charges, AMC fees, smallcase fees
- **Equity Dividends** — Dividend income

## Output Tally Ledger Mapping

| Trade Type | Profit Ledger | Loss Ledger |
|---|---|---|
| Equity - Intraday | Speculative Income | Speculative Loss |
| Equity - Short Term | Short Term Capital Gains | Short Term Capital Loss |
| Equity - Long Term | Long Term Capital Gains | Long Term Capital Loss |
| Non Equity | Non-Equity Capital Gains | Non-Equity Capital Loss |

| Charge Type | Ledger |
|---|---|
| DP Charges | DP Charges |
| AMC/Demat | Demat AMC Charges |
| Smallcase Fee | Smallcase Fees |
| Dividends | Dividend Income |

## Building the Windows Exe

The `.exe` is built automatically via GitHub Actions on every version tag. To build manually:

```bash
pip install -r requirements.txt
pyinstaller tallybridge.spec
# Output: dist/TallyBridge.exe
```

## Development

```
tally-bridge/
├── main.py              # Entry point
├── tallybridge.spec     # PyInstaller config
├── requirements.txt     # Python dependencies
├── src/
│   ├── __init__.py
│   ├── parser.py        # Zerodha Excel parser
│   ├── generator.py     # Tally Excel output generator
│   └── gui.py           # CustomTkinter GUI
└── .github/
    └── workflows/
        └── build.yml    # CI: builds Windows exe on tag push
```

## Creating a Release

```bash
git tag v1.0.0
git push origin v1.0.0
```

GitHub Actions will automatically:
1. Build the Windows `.exe`
2. Create a GitHub Release
3. Attach the `.exe` to the release

## Tech Stack

- **Python 3.11** — Core language
- **openpyxl** — Excel reading/writing
- **CustomTkinter** — Modern dark-mode GUI
- **PyInstaller** — Single-file Windows executable
- **GitHub Actions** — Automated builds

## License

MIT
