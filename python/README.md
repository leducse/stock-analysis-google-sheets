# Python Version - Stock Analysis Google Sheets Generator

## Setup

1. **Install dependencies**:
   ```bash
   pip install -r requirements.txt
   ```

2. **Set up Google Sheets API**:
   - Go to [Google Cloud Console](https://console.cloud.google.com/)
   - Create a new project or select existing
   - Enable Google Sheets API and Google Drive API
   - Create a Service Account
   - Download the JSON key file
   - Place it in the same directory as the script

3. **Configure the script**:
   - Update `GOOGLE_SHEET_JSON_KEYFILE` with your JSON key filename
   - Update `GOOGLE_SHEET_NAME` with your desired sheet name
   - Update `YOUR_EMAIL` with your email address

4. **Run the script**:
   ```bash
   python StockAnalysisGoogleSheet.py
   ```

## Features

- Fetches stock data from Yahoo Finance
- Calculates technical indicators (SMAs and RSI)
- Writes formatted results to Google Sheets
- Handles rate limiting with configurable delays
- Supports reading stock lists from Google Sheets

## Configuration Options

```python
DELAY_BETWEEN_REQUESTS = 2  # Seconds between requests
DELAY_ON_ERROR = 5  # Seconds to wait after errors
DELAY_ON_RATE_LIMIT = 10  # Seconds when rate limited
```

## Troubleshooting

- **Rate Limiting**: Increase `DELAY_BETWEEN_REQUESTS` if you get 429 errors
- **Credentials**: Make sure your JSON key file is in the correct location
- **Permissions**: Ensure the service account has access to Google Sheets and Drive

