# Google Apps Script Setup Instructions

## Quick Start Guide

### Step 1: Open Your Google Sheet
1. Open your Google Sheet (the one with "StockAnalysis" or create a new one)
2. Make sure you have a worksheet named **"Symbol List"** with stock symbols in column A

### Step 2: Open Apps Script Editor
1. Click **Extensions** → **Apps Script**
2. This opens the Apps Script editor in a new tab

### Step 3: Paste the Code
1. Delete any existing code in the editor
2. Copy the entire contents of `StockAnalysis_GoogleAppsScript.gs`
3. Paste it into the Apps Script editor
4. Click **Save** (💾 icon) or press `Ctrl+S`
5. Name your project (e.g., "Stock Analysis")

### Step 4: Set Up Your Symbol List
1. Go back to your Google Sheet
2. Make sure you have a worksheet named **"Symbol List"**
3. Put your stock symbols in column A (one per row), starting from row 1
   - Example:
     ```
     A1: AAPL
     A2: MSFT
     A3: GOOGL
     ```

### Step 5: Run the Script
**Option A: Using the Menu (Recommended)**
1. Go back to your Google Sheet
2. Refresh the page (F5)
3. You should see a new menu: **"Stock Analysis"**
4. Click **Stock Analysis** → **Update Stock Analysis (Auto-Continue)**
5. Authorize the script when prompted (first time only)

**Option B: From Apps Script Editor**
1. In the Apps Script editor, select the function `updateStockAnalysis` from the dropdown
2. Click the **Run** button (▶️)
3. Authorize the script when prompted (first time only)

### Step 6: Authorize the Script (First Time Only)
1. Click **Review Permissions**
2. Choose your Google account
3. Click **Advanced** → **Go to [Project Name] (unsafe)**
4. Click **Allow**
5. The script will now run

## What the Script Does

1. **Reads symbols** from the "Symbol List" worksheet
2. **Fetches historical data** using GOOGLEFINANCE for each symbol
3. **Calculates indicators:**
   - Current Price
   - 20-Day SMA
   - 50-Day SMA
   - 200-Day SMA
   - RSI (14-period)
4. **Determines opportunities:**
   - **Buy Opportunity**: RSI < 30 AND Current Price > 200-day SMA
   - **Sell Opportunity**: RSI > 70
5. **Writes results** to the "Stock Metrics" worksheet

## Output Sheet Structure

The script creates/updates a worksheet called **"Stock Metrics"** with these columns:

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

## Batch Processing & Auto-Continue

The script processes stocks in batches to avoid the 6-minute execution limit:
- **Batch Size**: 20 stocks per run (configurable)
- **Auto-Continue**: Automatically schedules the next batch if not finished
- **Resume Capability**: Automatically resumes from where it left off

When you run "Update Stock Analysis (Auto-Continue)", the script will:
1. Process up to 20 stocks
2. Automatically schedule the next batch to run in 1 minute
3. Continue until all stocks are processed
4. Clean up triggers when complete

## Scheduling Automatic Daily Updates

To run the script automatically after market close every trading day:

### Option 1: Time-Based Trigger (Recommended)
1. In Apps Script editor, click the **clock icon** (Triggers) on the left
2. Click **+ Add Trigger** (bottom right)
3. Configure:
   - **Function to run**: `updateStockAnalysis`
   - **Event source**: Time-driven
   - **Type of time based trigger**: Day timer
   - **Time of day**: 4:30 PM (or your preferred time after market close)
4. Click **Save**
5. The script will now run automatically every day at the specified time

### Option 2: Weekday-Only Trigger (Better)
1. In Apps Script editor, click **Triggers** → **+ Add Trigger**
2. Configure:
   - **Function to run**: `updateStockAnalysis`
   - **Event source**: Time-driven
   - **Type of time based trigger**: Week timer
   - **Day of week**: Monday through Friday
   - **Time of day**: 4:30 PM
3. Click **Save**

## Menu Options

The script adds a custom menu with these options:

- **Update Stock Analysis (Auto-Continue)**: Process all stocks with automatic batching
- **Refresh All Data**: Update existing rows with fresh data
- **Reset and Start Over**: Clear all data and start fresh
- **Stop Auto-Continue**: Stop any scheduled batch processing

## Troubleshooting

### "No symbols found" error
- Make sure you have a worksheet named exactly **"Symbol List"**
- Check that column A has symbols (not headers like "Symbol" or "Ticker")
- Symbols should start from row 1

### "Insufficient data" errors
- Some symbols may not have enough historical data
- The script needs at least 200 days of data
- Check the "Error" column in the results

### Script runs slowly
- GOOGLEFINANCE has rate limits
- The script includes delays between symbols
- For 100+ symbols, it may take 5-10 minutes (with auto-continue)

### GOOGLEFINANCE returns no data
- Some symbols may be delisted or invalid
- Check the symbol spelling
- Some international stocks may not be supported

### Permission errors
- Make sure you authorized the script
- Try running it again and authorize when prompted

### Script times out
- The script automatically handles this with batch processing
- It will continue processing in the next batch
- Check the execution log to see progress

## Configuration

You can modify these constants at the top of the script:

```javascript
const SYMBOL_SHEET_NAME = 'Symbol List';  // Name of sheet with symbols
const METRICS_SHEET_NAME = 'Stock Metrics';  // Name of output sheet
const PERIOD_DAYS = 365;  // Days of historical data to fetch
const MAX_STOCKS_PER_RUN = 20;  // Batch size (to avoid timeout)
const DELAY_BETWEEN_STOCKS = 1000;  // Milliseconds between stocks
const AUTO_CONTINUE = true;  // Auto-continue processing
const REFRESH_MODE = true;  // Update existing rows instead of appending
```

## Notes

- The script creates a temporary sheet called "_TEMP_DATA" during execution (it's cleaned up automatically)
- GOOGLEFINANCE may have delays, so the script includes wait times
- The script processes symbols sequentially to avoid rate limiting
- Results are written to "Stock Metrics" worksheet (created automatically if it doesn't exist)
- The script automatically handles the 6-minute execution limit with batch processing

## Support

If you encounter issues:
1. Check the **Execution log** in Apps Script (View → Logs)
2. Make sure all worksheet names match exactly
3. Verify your symbols are valid ticker symbols
4. Check that GOOGLEFINANCE is working: try `=GOOGLEFINANCE("AAPL","price")` in a cell

