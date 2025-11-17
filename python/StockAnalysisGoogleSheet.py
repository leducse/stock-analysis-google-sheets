"""
Stock Analysis Google Sheet Generator

This script fetches stock data, calculates technical indicators (SMAs and RSI),
and writes the results to a Google Sheet with buy/sell opportunity indicators.

Buying Opportunity: RSI < 30 AND Current Price > 200-day SMA
Selling Opportunity: RSI > 70

NOTE: Yahoo Finance (yfinance) may rate limit or block IP addresses that make too many requests.
If you encounter 429 errors or empty data, try:
1. Wait 15-30 minutes before running again
2. Use a VPN to change your IP address
3. Increase the DELAY_BETWEEN_REQUESTS value below
4. Reduce the number of stocks being analyzed
"""

import pandas as pd
import yfinance as yf
import gspread
import os
from oauth2client.service_account import ServiceAccountCredentials
import datetime
import requests
import time

# Configuration: Adjust these if you're getting rate limited (429 errors)
DELAY_BETWEEN_REQUESTS = 2  # Seconds to wait between stock data requests (increase if rate limited)
DELAY_ON_ERROR = 5  # Seconds to wait after an error
DELAY_ON_RATE_LIMIT = 10  # Seconds to wait when rate limited

# Google Sheets Configuration
# TODO: Update these with your own values
GOOGLE_SHEET_JSON_KEYFILE = 'your-credentials.json'  # Path to your Google Service Account JSON file
GOOGLE_SHEET_NAME = 'StockAnalysis'  # Name of the Google Sheet to create/update
WORKSHEET_NAME = 'Stock Metrics'  # Name of the worksheet within the sheet
YOUR_EMAIL = 'your-email@gmail.com'  # Your email address - the sheet will be shared with this email


class StockAnalyzer:
    """Class to analyze stocks and calculate technical indicators."""
    
    def __init__(self):
        pass
    
    def calculate_rsi(self, series, period=14):
        """Calculate Relative Strength Index (RSI)."""
        delta = series.diff()
        gain = (delta.where(delta > 0, 0)).rolling(window=period).mean()
        loss = (-delta.where(delta < 0, 0)).rolling(window=period).mean()
        
        rs = gain / loss
        rsi = 100 - (100 / (1 + rs))
        return rsi
    
    def calculate_sma(self, series, period):
        """Calculate Simple Moving Average (SMA)."""
        return series.rolling(window=period).mean()
    
    def analyze_stock(self, ticker_symbol, period_days=250):
        """
        Analyze a stock and return key metrics.
        
        Returns:
            dict with keys: symbol, current_price, sma_20, sma_50, sma_200, rsi, 
                           buy_opportunity, sell_opportunity, error
        """
        try:
            ticker = yf.Ticker(ticker_symbol)
            # Get enough data for 200-day SMA (need at least 200 trading days)
            data = ticker.history(period=f"{period_days}d", timeout=15)
            
            if data.empty or len(data) < 200:
                return {
                    'symbol': ticker_symbol,
                    'current_price': None,
                    'sma_20': None,
                    'sma_50': None,
                    'sma_200': None,
                    'rsi': None,
                    'buy_opportunity': None,
                    'sell_opportunity': None,
                    'error': 'Insufficient data'
                }
            
            close_prices = data['Close']
            current_price = close_prices.iloc[-1]
            
            # Calculate SMAs
            sma_20 = self.calculate_sma(close_prices, 20).iloc[-1]
            sma_50 = self.calculate_sma(close_prices, 50).iloc[-1]
            sma_200 = self.calculate_sma(close_prices, 200).iloc[-1]
            
            # Calculate RSI
            rsi_series = self.calculate_rsi(close_prices)
            rsi = rsi_series.iloc[-1]
            
            # Determine opportunities
            buy_opportunity = "YES" if (rsi < 30 and current_price > sma_200) else "NO"
            sell_opportunity = "YES" if rsi > 70 else "NO"
            
            return {
                'symbol': ticker_symbol,
                'current_price': round(current_price, 2),
                'sma_20': round(sma_20, 2) if pd.notna(sma_20) else None,
                'sma_50': round(sma_50, 2) if pd.notna(sma_50) else None,
                'sma_200': round(sma_200, 2) if pd.notna(sma_200) else None,
                'rsi': round(rsi, 2) if pd.notna(rsi) else None,
                'buy_opportunity': buy_opportunity,
                'sell_opportunity': sell_opportunity,
                'error': None
            }
            
        except Exception as e:
            return {
                'symbol': ticker_symbol,
                'current_price': None,
                'sma_20': None,
                'sma_50': None,
                'sma_200': None,
                'rsi': None,
                'buy_opportunity': None,
                'sell_opportunity': None,
                'error': str(e)[:100]  # Truncate error message
            }


def find_credentials_file(filename):
    """Search for credentials file in multiple common locations."""
    search_paths = [
        filename,  # Current directory
        os.path.join(os.path.dirname(os.path.abspath(__file__)), filename),  # Script directory
        os.path.join(os.path.expanduser('~'), filename),  # User home directory
    ]
    
    for path in search_paths:
        if os.path.exists(path):
            return path
    
    return None


def fetch_tickers_from_google_sheets(json_keyfile, sheet_name, worksheet_name='Symbol List'):
    """Fetch ticker symbols from Google Sheets."""
    try:
        # Find credentials file in multiple locations
        creds_path = find_credentials_file(json_keyfile)
        if not creds_path:
            print(f"Warning: Google Sheets credentials file not found. Searched in:")
            print(f"  - Current directory")
            print(f"  - Script directory")
            print(f"  - User home directory")
            return []
        
        json_keyfile = creds_path
        
        scope = ['https://spreadsheets.google.com/feeds', 'https://www.googleapis.com/auth/drive']
        creds = ServiceAccountCredentials.from_json_keyfile_name(json_keyfile, scope)
        client = gspread.authorize(creds)
        
        # Open the spreadsheet and get the specified worksheet
        spreadsheet = client.open(sheet_name)
        try:
            worksheet = spreadsheet.worksheet(worksheet_name)
        except gspread.exceptions.WorksheetNotFound:
            # Fallback to first worksheet if "Symbol List" doesn't exist
            print(f"Warning: Worksheet '{worksheet_name}' not found, using first worksheet")
            worksheet = spreadsheet.sheet1
        
        # Get tickers from first column
        tickers = worksheet.col_values(1)
        # Filter out empty strings, header rows, and clean up
        tickers = [t.strip().upper() for t in tickers if t and t.strip() and t.strip().upper() not in ['SYMBOL', 'TICKER', 'STOCK', 'ETF']]
        return tickers
    except Exception as e:
        print(f"Error fetching tickers from Google Sheets: {e}")
        return []


def fetch_sp500_stocks():
    """Fetch S&P 500 stock symbols from Wikipedia."""
    url = "https://en.wikipedia.org/wiki/List_of_S%26P_500_companies"
    headers = {'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36'}
    
    try:
        response = requests.get(url, headers=headers)
        response.raise_for_status()
        tables = pd.read_html(response.text)
        sp500_table = tables[0]
        sp500_symbols = sp500_table['Symbol'].tolist()
        return sp500_symbols
    except Exception as e:
        print(f"Error fetching S&P 500 list: {e}")
        print("Using a default list of popular stocks instead...")
        # Fallback to a small list of popular stocks if Wikipedia fails
        return ['AAPL', 'MSFT', 'GOOGL', 'AMZN', 'NVDA', 'META', 'TSLA', 'BRK.B', 'V', 'JNJ']


def get_or_create_google_sheet(json_keyfile, sheet_name, worksheet_name):
    """Get existing Google Sheet or create a new one."""
    try:
        # Find credentials file in multiple locations
        creds_path = find_credentials_file(json_keyfile)
        if not creds_path:
            raise FileNotFoundError(
                f"Google Sheets credentials file not found: {json_keyfile}\n"
                f"Searched in: current directory, script directory, user home"
            )
        
        json_keyfile = creds_path
        
        scope = ['https://spreadsheets.google.com/feeds', 'https://www.googleapis.com/auth/drive']
        creds = ServiceAccountCredentials.from_json_keyfile_name(json_keyfile, scope)
        client = gspread.authorize(creds)
        
        # Try to open existing sheet
        try:
            spreadsheet = client.open(sheet_name)
            print(f"Opened existing Google Sheet: {sheet_name}")
        except gspread.exceptions.SpreadsheetNotFound:
            # Create new sheet
            spreadsheet = client.create(sheet_name)
            print(f"Created new Google Sheet: {sheet_name}")
        
        # Get or create worksheet
        try:
            worksheet = spreadsheet.worksheet(worksheet_name)
            print(f"Using existing worksheet: {worksheet_name}")
        except gspread.exceptions.WorksheetNotFound:
            worksheet = spreadsheet.add_worksheet(title=worksheet_name, rows=1000, cols=10)
            print(f"Created new worksheet: {worksheet_name}")
        
        # Always share the sheet with the user's email (if provided)
        # This ensures you have access to the sheet
        try:
            spreadsheet.share(YOUR_EMAIL, perm_type='user', role='writer')
            print(f"Shared sheet with: {YOUR_EMAIL}")
        except Exception as e:
            print(f"Warning: Could not share sheet with {YOUR_EMAIL}: {e}")
            print("   You may need to manually share the sheet or check the service account permissions")
        
        return worksheet, client
        
    except Exception as e:
        print(f"Error accessing Google Sheets: {e}")
        raise


def write_to_google_sheet(worksheet, stock_data_list):
    """Write stock analysis data to Google Sheet."""
    # Define headers
    headers = [
        'Symbol',
        'Current Price',
        '20-Day SMA',
        '50-Day SMA',
        '200-Day SMA',
        'RSI',
        'Buy Opportunity',
        'Sell Opportunity',
        'Last Updated',
        'Error'
    ]
    
    # Prepare data rows
    rows = [headers]
    for stock in stock_data_list:
        row = [
            stock['symbol'],
            stock['current_price'] if stock['current_price'] is not None else '',
            stock['sma_20'] if stock['sma_20'] is not None else '',
            stock['sma_50'] if stock['sma_50'] is not None else '',
            stock['sma_200'] if stock['sma_200'] is not None else '',
            stock['rsi'] if stock['rsi'] is not None else '',
            stock['buy_opportunity'] if stock['buy_opportunity'] else '',
            stock['sell_opportunity'] if stock['sell_opportunity'] else '',
            datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S'),
            stock['error'] if stock['error'] else ''
        ]
        rows.append(row)
    
    # Clear existing data and write new data
    worksheet.clear()
    # Use the older update syntax for compatibility
    worksheet.update('A1', rows)
    
    # Format header row (bold)
    worksheet.format('A1:J1', {'textFormat': {'bold': True}})
    
    # Format buy/sell opportunity columns with colors
    num_rows = len(rows)
    if num_rows > 1:
        # Format Buy Opportunity column (green for YES, red for NO)
        buy_range = f'G2:G{num_rows}'
        worksheet.format(buy_range, {
            'backgroundColor': {'red': 0.85, 'green': 0.95, 'blue': 0.85}
        })
        
        # Format Sell Opportunity column (red for YES, green for NO)
        sell_range = f'H2:H{num_rows}'
        worksheet.format(sell_range, {
            'backgroundColor': {'red': 0.95, 'green': 0.85, 'blue': 0.85}
        })
    
    print(f"Successfully wrote {len(stock_data_list)} stocks to Google Sheet")


if __name__ == '__main__':
    print("=" * 60)
    print("Stock Analysis Google Sheet Generator")
    print("=" * 60)
    
    analyzer = StockAnalyzer()
    
    # Get list of stocks to analyze
    print("\n1. Fetching stock list...")
    try:
        # Try to get tickers from Google Sheets "Symbol List" worksheet
        additional_tickers = fetch_tickers_from_google_sheets(GOOGLE_SHEET_JSON_KEYFILE, GOOGLE_SHEET_NAME, 'Symbol List')
        if additional_tickers:
            print(f"   Found {len(additional_tickers)} tickers from Google Sheets 'Symbol List' worksheet")
            all_stocks = sorted(set(additional_tickers))
        else:
            # Fallback to S&P 500
            print("   No tickers found in 'Symbol List' worksheet. Fetching S&P 500 stocks...")
            sp500_stocks = fetch_sp500_stocks()
            all_stocks = sorted(set(sp500_stocks[:50]))  # Limit to first 50 for testing
            print(f"   Using {len(all_stocks)} stocks from S&P 500")
    except Exception as e:
        print(f"   Error fetching stock list: {e}")
        print("   Using default stock list...")
        all_stocks = ['AAPL', 'MSFT', 'GOOGL', 'AMZN', 'NVDA', 'META', 'TSLA']
    
    if not all_stocks:
        print("No stocks to analyze. Exiting.")
        exit(0)
    
    print(f"\n2. Analyzing {len(all_stocks)} stocks...")
    print("   (This may take a while due to rate limiting)")
    
    stock_data_list = []
    for i, ticker in enumerate(all_stocks):
        try:
            if i > 0:
                time.sleep(DELAY_BETWEEN_REQUESTS)
            print(f"   [{i+1}/{len(all_stocks)}] Analyzing {ticker}...", end=" ")
            
            stock_data = analyzer.analyze_stock(ticker)
            stock_data_list.append(stock_data)
            
            if stock_data['error']:
                print(f"Error: {stock_data['error'][:50]}")
                time.sleep(DELAY_ON_ERROR)
            else:
                buy = "BUY" if stock_data['buy_opportunity'] == "YES" else ""
                sell = "SELL" if stock_data['sell_opportunity'] == "YES" else ""
                signal = f"{buy} {sell}".strip()
                print(f"RSI: {stock_data['rsi']:.2f} {signal}" if stock_data['rsi'] else "Success")
                
        except Exception as e:
            error_msg = str(e)
            if "429" in error_msg or "rate limit" in error_msg.lower() or "too many requests" in error_msg.lower():
                print(f"Rate limited. Waiting {DELAY_ON_RATE_LIMIT}s...")
                time.sleep(DELAY_ON_RATE_LIMIT)
            else:
                print(f"Error: {error_msg[:50]}")
                time.sleep(DELAY_ON_ERROR)
            
            # Add error entry
            stock_data_list.append({
                'symbol': ticker,
                'current_price': None,
                'sma_20': None,
                'sma_50': None,
                'sma_200': None,
                'rsi': None,
                'buy_opportunity': None,
                'sell_opportunity': None,
                'error': error_msg[:100]
            })
    
    print(f"\n3. Writing data to Google Sheet...")
    sheet_url = None
    try:
        worksheet, client = get_or_create_google_sheet(
            GOOGLE_SHEET_JSON_KEYFILE,
            GOOGLE_SHEET_NAME,
            WORKSHEET_NAME
        )
        
        # Get the URL of the sheet before writing (in case write fails)
        spreadsheet = client.open(GOOGLE_SHEET_NAME)
        sheet_url = f"https://docs.google.com/spreadsheets/d/{spreadsheet.id}"
        print(f"   Google Sheet URL: {sheet_url}")
        
        write_to_google_sheet(worksheet, stock_data_list)
        print(f"\nSuccess! View your sheet at: {sheet_url}")
        
    except Exception as e:
        print(f"\nError writing to Google Sheet: {e}")
        if sheet_url:
            print(f"\nSheet was created/accessed but write failed. Try accessing it at:")
            print(f"   {sheet_url}")
        else:
            print(f"\nTo find your sheet manually:")
            print(f"   1. Go to https://drive.google.com")
            print(f"   2. Search for: '{GOOGLE_SHEET_NAME}'")
            print(f"   3. Or look in 'Shared with me' or 'My Drive'")
        print("\nStock data (first 5 rows):")
        df = pd.DataFrame(stock_data_list)
        print(df.head().to_string())
    
    print("\n" + "=" * 60)
    print("Analysis complete!")
    print("=" * 60)

