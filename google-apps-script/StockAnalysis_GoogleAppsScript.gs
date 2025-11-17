/**
 * Stock Analysis Google Apps Script
 * 
 * This script analyzes stocks using GOOGLEFINANCE data, calculates technical indicators
 * (SMAs and RSI), and generates buy/sell signals.
 * 
 * Features:
 * - Batch processing to handle large stock lists (avoids 6-minute execution limit)
 * - Auto-continue functionality via time-based triggers
 * - Refresh mode to update existing data
 * - Custom menu for easy access
 * 
 * Buying Opportunity: RSI < 30 AND Current Price > 200-day SMA
 * Selling Opportunity: RSI > 70
 */

// Configuration
const SYMBOL_SHEET_NAME = 'Symbol List';
const METRICS_SHEET_NAME = 'Stock Metrics';
const PERIOD_DAYS = 365;  // Days of historical data to fetch (need 200+ for 200-day SMA)
const MAX_STOCKS_PER_RUN = 20;  // Process this many stocks per execution (to avoid timeout)
const DELAY_BETWEEN_STOCKS = 1000;  // Milliseconds to wait between stocks
const AUTO_CONTINUE = true;  // Automatically schedule next batch if not finished
const REFRESH_MODE = true;  // If true, updates existing rows instead of appending

/**
 * Main function - called from menu or trigger
 */
function updateStockAnalysis() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const startTime = new Date();
  
  try {
    // Get symbols
    const symbols = getSymbols(ss);
    if (symbols.length === 0) {
      Logger.log('No symbols found. Please add symbols to "' + SYMBOL_SHEET_NAME + '" worksheet.');
      SpreadsheetApp.getUi().alert('No symbols found. Please add symbols to "' + SYMBOL_SHEET_NAME + '" worksheet.');
      return;
    }
    
    Logger.log('Found ' + symbols.length + ' symbols to process');
    
    // Setup metrics sheet
    const metricsSheet = setupMetricsSheet(ss);
    
    // Check if we're in refresh mode (updating existing data)
    const isUpdateMode = REFRESH_MODE && getLastUpdateTime(metricsSheet) !== null;
    let resumeIndex = 0;
    
    if (!isUpdateMode) {
      // Get resume index (where we left off in previous run)
      resumeIndex = getResumeIndex(metricsSheet, symbols.length);
      Logger.log('Resuming from index: ' + resumeIndex);
    }
    
    // Process stocks in batches
    let processed = 0;
    let startIndex = isUpdateMode ? 0 : resumeIndex;
    
    for (let i = startIndex; i < symbols.length && processed < MAX_STOCKS_PER_RUN; i++) {
      const symbol = symbols[i];
      Logger.log('Processing [' + (i + 1) + '/' + symbols.length + ']: ' + symbol);
      
      try {
        const result = analyzeStock(symbol);
        
        if (isUpdateMode) {
          updateExistingRow(metricsSheet, result, i + 2);  // +2 for header and 1-based index
        } else {
          appendResults(metricsSheet, result);
        }
        
        processed++;
        
        // Small delay to avoid rate limiting
        if (i < symbols.length - 1) {
          Utilities.sleep(DELAY_BETWEEN_STOCKS);
        }
        
      } catch (error) {
        Logger.log('Error processing ' + symbol + ': ' + error.toString());
        const errorResult = {
          symbol: symbol,
          currentPrice: null,
          sma20: null,
          sma50: null,
          sma200: null,
          rsi: null,
          buyOpportunity: 'NO',
          sellOpportunity: 'NO',
          lastUpdated: new Date(),
          error: error.toString()
        };
        
        if (isUpdateMode) {
          updateExistingRow(metricsSheet, errorResult, i + 2);
        } else {
          appendResults(metricsSheet, errorResult);
        }
        processed++;
      }
      
      // Check execution time (Apps Script has 6-minute limit)
      const elapsed = (new Date() - startTime) / 1000;
      if (elapsed > 300) {  // 5 minutes - leave 1 minute buffer
        Logger.log('Approaching time limit. Stopping at index ' + i);
        break;
      }
    }
    
    const endIndex = startIndex + processed;
    Logger.log('Processed ' + processed + ' stocks (indices ' + startIndex + ' to ' + (endIndex - 1) + ')');
    
    // Schedule next batch if needed
    if (AUTO_CONTINUE && endIndex < symbols.length && !isUpdateMode) {
      scheduleNextBatch();
      Logger.log('Scheduled next batch to continue from index ' + endIndex);
    } else if (endIndex >= symbols.length) {
      // Finished processing all stocks
      cleanupTriggers();
      Logger.log('All stocks processed!');
      
      // Format the sheet
      formatMetricsSheet(metricsSheet);
      
      // Update last update time
      if (metricsSheet.getRange(1, 9).getValue() === 'Last Updated') {
        // Header row exists, update timestamp in a specific cell
        metricsSheet.getRange(2, 9).setValue(new Date());
      }
    }
    
  } catch (error) {
    Logger.log('Error in updateStockAnalysis: ' + error.toString());
    SpreadsheetApp.getUi().alert('Error: ' + error.toString());
  }
}

/**
 * Get stock symbols from the Symbol List worksheet
 */
function getSymbols(ss) {
  try {
    const symbolSheet = ss.getSheetByName(SYMBOL_SHEET_NAME);
    if (!symbolSheet) {
      Logger.log('Symbol sheet not found: ' + SYMBOL_SHEET_NAME);
      return [];
    }
    
    const data = symbolSheet.getRange(1, 1, symbolSheet.getLastRow(), 1).getValues();
    const symbols = [];
    
    for (let i = 0; i < data.length; i++) {
      const symbol = String(data[i][0]).trim().toUpperCase();
      if (symbol && symbol !== 'SYMBOL' && symbol !== 'TICKER' && symbol !== 'STOCK' && symbol !== 'ETF') {
        symbols.push(symbol);
      }
    }
    
    return symbols;
  } catch (error) {
    Logger.log('Error getting symbols: ' + error.toString());
    return [];
  }
}

/**
 * Setup or get the metrics sheet
 */
function setupMetricsSheet(ss) {
  let metricsSheet = ss.getSheetByName(METRICS_SHEET_NAME);
  
  if (!metricsSheet) {
    metricsSheet = ss.insertSheet(METRICS_SHEET_NAME);
    // Add headers
    const headers = [
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
    ];
    metricsSheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    formatMetricsSheet(metricsSheet);
  } else {
    // Check if headers exist
    if (metricsSheet.getRange(1, 1).getValue() !== 'Symbol') {
      const headers = [
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
      ];
      metricsSheet.insertRowBefore(1);
      metricsSheet.getRange(1, 1, 1, headers.length).setValues([headers]);
      formatMetricsSheet(metricsSheet);
    }
  }
  
  return metricsSheet;
}

/**
 * Analyze a single stock
 */
function analyzeStock(symbol) {
  try {
    // Get historical data
    const historicalData = getHistoricalData(symbol);
    
    if (!historicalData || historicalData.length < 200) {
      throw new Error('Insufficient data (need at least 200 days)');
    }
    
    // Extract close prices
    const closePrices = historicalData.map(row => row[1]);  // Column 1 is close price
    const currentPrice = closePrices[closePrices.length - 1];
    
    // Calculate SMAs
    const sma20 = calculateSMA(closePrices, 20);
    const sma50 = calculateSMA(closePrices, 50);
    const sma200 = calculateSMA(closePrices, 200);
    
    // Calculate RSI
    const rsi = calculateRSI(closePrices, 14);
    
    // Determine opportunities
    const buyOpportunity = (rsi < 30 && currentPrice > sma200) ? 'YES' : 'NO';
    const sellOpportunity = (rsi > 70) ? 'YES' : 'NO';
    
    return {
      symbol: symbol,
      currentPrice: currentPrice,
      sma20: sma20,
      sma50: sma50,
      sma200: sma200,
      rsi: rsi,
      buyOpportunity: buyOpportunity,
      sellOpportunity: sellOpportunity,
      lastUpdated: new Date(),
      error: null
    };
    
  } catch (error) {
    Logger.log('Error analyzing ' + symbol + ': ' + error.toString());
    return {
      symbol: symbol,
      currentPrice: null,
      sma20: null,
      sma50: null,
      sma200: null,
      rsi: null,
      buyOpportunity: 'NO',
      sellOpportunity: 'NO',
      lastUpdated: new Date(),
      error: error.toString()
    };
  }
}

/**
 * Get historical data using GOOGLEFINANCE
 */
function getHistoricalData(symbol) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let tempSheet = ss.getSheetByName('_TEMP_DATA');
  
  if (!tempSheet) {
    tempSheet = ss.insertSheet('_TEMP_DATA');
  } else {
    tempSheet.clear();
  }
  
  try {
    // Use GOOGLEFINANCE to get historical data
    const formula = '=GOOGLEFINANCE("' + symbol + '","close",TODAY()-' + PERIOD_DAYS + ',TODAY())';
    tempSheet.getRange(1, 1).setFormula(formula);
    
    // Wait for formula to calculate (GOOGLEFINANCE can be slow)
    let attempts = 0;
    let data = [];
    
    while (attempts < 30) {  // Wait up to 30 seconds
      Utilities.sleep(1000);
      const range = tempSheet.getDataRange();
      const values = range.getValues();
      
      if (values.length > 1) {
        // Filter out header rows and invalid data
        data = values.filter(row => {
          if (!row[0] || !row[1]) return false;
          const date = row[0];
          const price = row[1];
          
          // Check if date is valid
          if (Object.prototype.toString.call(date) !== '[object Date]' || isNaN(date.getTime())) {
            return false;
          }
          
          // Check if price is valid number
          if (typeof price !== 'number' || isNaN(price) || price <= 0) {
            return false;
          }
          
          return true;
        });
        
        if (data.length >= 200) {
          break;
        }
      }
      
      attempts++;
    }
    
    if (data.length < 200) {
      Logger.log('Warning: Only got ' + data.length + ' data points for ' + symbol);
    }
    
    return data;
    
  } catch (error) {
    Logger.log('Error getting historical data for ' + symbol + ': ' + error.toString());
    return null;
  }
}

/**
 * Calculate Simple Moving Average
 */
function calculateSMA(prices, period) {
  if (prices.length < period) {
    return null;
  }
  
  const slice = prices.slice(-period);
  const sum = slice.reduce((a, b) => a + b, 0);
  return sum / period;
}

/**
 * Calculate Relative Strength Index (RSI)
 */
function calculateRSI(prices, period) {
  if (prices.length < period + 1) {
    return null;
  }
  
  const deltas = [];
  for (let i = 1; i < prices.length; i++) {
    deltas.push(prices[i] - prices[i - 1]);
  }
  
  let gains = 0;
  let losses = 0;
  
  // Calculate initial average gain and loss
  for (let i = deltas.length - period; i < deltas.length; i++) {
    if (deltas[i] > 0) {
      gains += deltas[i];
    } else {
      losses += Math.abs(deltas[i]);
    }
  }
  
  const avgGain = gains / period;
  const avgLoss = losses / period;
  
  if (avgLoss === 0) {
    return 100;
  }
  
  const rs = avgGain / avgLoss;
  const rsi = 100 - (100 / (1 + rs));
  
  return rsi;
}

/**
 * Append results to the metrics sheet
 */
function appendResults(sheet, result) {
  const row = [
    result.symbol,
    result.currentPrice,
    result.sma20,
    result.sma50,
    result.sma200,
    result.rsi,
    result.buyOpportunity,
    result.sellOpportunity,
    result.lastUpdated,
    result.error
  ];
  
  sheet.appendRow(row);
}

/**
 * Update existing row in the metrics sheet
 */
function updateExistingRow(sheet, result, rowIndex) {
  const row = [
    result.symbol,
    result.currentPrice,
    result.sma20,
    result.sma50,
    result.sma200,
    result.rsi,
    result.buyOpportunity,
    result.sellOpportunity,
    result.lastUpdated,
    result.error
  ];
  
  sheet.getRange(rowIndex, 1, 1, row.length).setValues([row]);
}

/**
 * Get the index to resume from (for batch processing)
 */
function getResumeIndex(sheet, totalSymbols) {
  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) {
    return 0;  // No data yet, start from beginning
  }
  
  // Count non-header rows
  const dataRows = lastRow - 1;
  return Math.min(dataRows, totalSymbols);
}

/**
 * Get last update time from the sheet
 */
function getLastUpdateTime(sheet) {
  try {
    const lastRow = sheet.getLastRow();
    if (lastRow <= 1) {
      return null;
    }
    
    // Check the timestamp in row 2, column 9 (Last Updated column)
    const timestamp = sheet.getRange(2, 9).getValue();
    if (timestamp instanceof Date) {
      return timestamp;
    }
    return null;
  } catch (error) {
    return null;
  }
}

/**
 * Schedule the next batch to run
 */
function scheduleNextBatch() {
  // Delete existing triggers for this function
  cleanupTriggers();
  
  // Create a new trigger to run in 1 minute
  ScriptApp.newTrigger('updateStockAnalysis')
    .timeBased()
    .after(60000)  // 1 minute from now
    .create();
}

/**
 * Clean up existing triggers
 */
function cleanupTriggers() {
  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(trigger => {
    if (trigger.getHandlerFunction() === 'updateStockAnalysis') {
      ScriptApp.deleteTrigger(trigger);
    }
  });
}

/**
 * Format the metrics sheet
 */
function formatMetricsSheet(sheet) {
  const headerRange = sheet.getRange(1, 1, 1, 10);
  headerRange.setFontWeight('bold');
  headerRange.setBackground('#4285f4');
  headerRange.setFontColor('#ffffff');
  
  // Freeze header row
  sheet.setFrozenRows(1);
  
  // Auto-resize columns
  sheet.autoResizeColumns(1, 10);
}

/**
 * Create custom menu when sheet opens
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('Stock Analysis')
    .addItem('Update Stock Analysis (Auto-Continue)', 'updateStockAnalysis')
    .addItem('Refresh All Data', 'refreshAllData')
    .addItem('Reset and Start Over', 'resetAndStartOver')
    .addItem('Stop Auto-Continue', 'stopAutoContinue')
    .addToUi();
}

/**
 * Refresh all existing data
 */
function refreshAllData() {
  // Force update mode by checking last update time
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const metricsSheet = ss.getSheetByName(METRICS_SHEET_NAME);
  
  if (metricsSheet && metricsSheet.getLastRow() > 1) {
    // Update existing data
    updateStockAnalysis();
  } else {
    SpreadsheetApp.getUi().alert('No existing data to refresh. Use "Update Stock Analysis" instead.');
  }
}

/**
 * Reset and start over (clear existing data)
 */
function resetAndStartOver() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const metricsSheet = ss.getSheetByName(METRICS_SHEET_NAME);
  
  if (metricsSheet) {
    const response = SpreadsheetApp.getUi().alert(
      'Are you sure you want to clear all existing data?',
      SpreadsheetApp.getUi().ButtonSet.YES_NO
    );
    
    if (response === SpreadsheetApp.getUi().Button.YES) {
      metricsSheet.clear();
      setupMetricsSheet(ss);
      cleanupTriggers();
      updateStockAnalysis();
    }
  }
}

/**
 * Stop auto-continue processing
 */
function stopAutoContinue() {
  cleanupTriggers();
  SpreadsheetApp.getUi().alert('Auto-continue stopped. All triggers have been removed.');
}

