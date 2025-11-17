# Stock Analysis Google Sheets Generator

An automated stock analysis tool that calculates technical indicators (SMAs and RSI) and generates buy/sell signals, outputting results to Google Sheets. Available in both **Python** and **Google Apps Script** versions.

## Features

- 📊 **Technical Indicators**: Calculates 20-day, 50-day, and 200-day Simple Moving Averages (SMA)
- 📈 **RSI Analysis**: Computes 14-period Relative Strength Index
- 🎯 **Buy/Sell Signals**:
  - **Buy Opportunity**: RSI < 30 AND Current Price > 200-day SMA
  - **Sell Opportunity**: RSI > 70
- 📋 **Google Sheets Integration**: Automatically writes formatted results to Google Sheets
- 🔄 **Batch Processing**: Handles large stock lists with rate limiting protection
- ⏰ **Scheduled Updates**: Can be automated to run daily after market close
- 📧 **Email Notifications**: Optional email alerts when analysis completes

## Project Structure

```
stock-analysis-google-sheets/
├── python/
│   ├── StockAnalysisGoogleSheet.py    # Python version
│   ├── requirements.txt                # Python dependencies
│   └── README.md                       # Python setup instructions
├── google-apps-script/
│   ├── StockAnalysis_GoogleAppsScript.gs  # Apps Script version
│   └── SETUP_INSTRUCTIONS.md              # Apps Script setup guide
├── data/
│   └── stock_symbols_list.txt         # Sample stock/ETF list (100 symbols)
└── README.md                          # This file
```

## Quick Start

### Python Version

1. **Install dependencies**:
   ```bash
   pip install -r python/requirements.txt
   ```

2. **Configure Google Sheets credentials**:
   - Download your Google Service Account JSON key file
   - Place it in the same directory as the script or update the path in the script

3. **Update configuration** in `StockAnalysisGoogleSheet.py`:
   ```python
   GOOGLE_SHEET_JSON_KEYFILE = 'your-credentials.json'
   GOOGLE_SHEET_NAME = 'StockAnalysis'
   YOUR_EMAIL = 'your-email@gmail.com'
   ```

4. **Run the script**:
   ```bash
   python StockAnalysisGoogleSheet.py
   ```

### Google Apps Script Version

1. **Open Google Sheets** and create a new spreadsheet
2. **Add stock symbols** to a worksheet named "Symbol List" (column A)
3. **Open Apps Script**: Extensions → Apps Script
4. **Paste the code** from `StockAnalysis_GoogleAppsScript.gs`
5. **Run** from the menu: Stock Analysis → Update Stock Analysis

See `google-apps-script/SETUP_INSTRUCTIONS.md` for detailed setup.

## Technical Indicators

### Simple Moving Averages (SMA)
- **20-Day SMA**: Short-term trend indicator
- **50-Day SMA**: Medium-term trend indicator
- **200-Day SMA**: Long-term trend indicator (bull/bear market indicator)

### Relative Strength Index (RSI)
- **RSI < 30**: Oversold condition (potential buy signal)
- **RSI > 70**: Overbought condition (potential sell signal)
- **RSI 30-70**: Neutral range

### Trading Signals

**Buy Signal**: 
- RSI < 30 (oversold)
- AND Current Price > 200-day SMA (above long-term trend)

**Sell Signal**:
- RSI > 70 (overbought)

## Output Format

The script generates a Google Sheet with the following columns:

| Column | Description |
|--------|-------------|
| Symbol | Stock ticker symbol |
| Current Price | Latest closing price |
| 20-Day SMA | 20-day Simple Moving Average |
| 50-Day SMA | 50-day Simple Moving Average |
| 200-Day SMA | 200-day Simple Moving Average |
| RSI | Relative Strength Index (14-period) |
| Buy Opportunity | YES/NO based on RSI < 30 and price > 200 SMA |
| Sell Opportunity | YES/NO based on RSI > 70 |
| Last Updated | Timestamp of when data was updated |
| Error | Any error messages |

## Rate Limiting

Both versions include rate limiting protection:
- **Python**: Configurable delays between requests (default: 2 seconds)
- **Apps Script**: Automatic delays and batch processing to avoid timeouts

If you encounter rate limiting:
1. Increase delay between requests
2. Reduce the number of stocks analyzed
3. Wait 15-30 minutes before retrying
4. Use a VPN to change IP address

## Scheduling

### Python Version
Use Windows Task Scheduler or cron to run daily:
```bash
# Windows Task Scheduler
# Schedule: Daily at 4:30 PM ET
python C:\path\to\StockAnalysisGoogleSheet.py
```

### Google Apps Script Version
1. In Apps Script editor, click **Triggers** (clock icon)
2. Click **+ Add Trigger**
3. Configure:
   - Function: `updateStockAnalysis`
   - Event: Time-driven → Day timer
   - Time: 4:30 PM (after market close)
4. Save

## Requirements

### Python Version
- Python 3.7+
- pandas
- yfinance
- gspread
- oauth2client
- requests

### Google Apps Script Version
- Google account
- Google Sheets access
- No additional dependencies (uses built-in GOOGLEFINANCE function)

## Configuration

### Python Configuration
```python
DELAY_BETWEEN_REQUESTS = 2  # Seconds between requests
GOOGLE_SHEET_NAME = 'StockAnalysis'
WORKSHEET_NAME = 'Stock Metrics'
YOUR_EMAIL = 'your-email@gmail.com'
```

### Apps Script Configuration
```javascript
const MAX_STOCKS_PER_RUN = 20;  // Batch size
const DELAY_BETWEEN_STOCKS = 1000;  // Milliseconds
const AUTO_CONTINUE = true;  // Auto-continue processing
const SEND_EMAIL = true;  // Email results
```

## Features Comparison

| Feature | Python Version | Apps Script Version |
|---------|---------------|---------------------|
| Data Source | Yahoo Finance API (yfinance) | GOOGLEFINANCE function |
| Execution | Local machine | Google Cloud |
| Rate Limiting | Manual delays | Automatic handling |
| Batch Processing | Manual | Automatic with triggers |
| Email Alerts | Not included | Built-in |
| Setup Complexity | Medium | Low |
| Best For | Bulk processing, advanced analysis | Daily automation, cloud-based |

## Use Cases

- **Day Trading**: Identify oversold/overbought conditions
- **Swing Trading**: Find stocks near support/resistance with RSI confirmation
- **Portfolio Management**: Monitor multiple stocks for entry/exit signals
- **Research**: Analyze technical indicators across large stock lists

## Limitations

- Requires at least 200 trading days of historical data for accurate 200-day SMA
- Rate limiting may affect large batch processing
- GOOGLEFINANCE has limitations on number of calls per sheet
- Not financial advice - use for research purposes only

## Contributing

Feel free to submit issues, fork the repository, and create pull requests for any improvements.

## License

This project is provided as-is for educational and research purposes.

## Disclaimer

This tool is for educational and research purposes only. It does not constitute financial advice. Always do your own research and consult with a financial advisor before making investment decisions.

