/**
 * FUT TRADING CONSOLE - BUG FIXES APPLIED + NEW UPDATES
 * Combined: Fluctuations Trading + Chemistry Styles Arbitrage
 * 
 * BUG FIXES APPLIED (October 20, 2025):
 * ========================================
 * 1. DASHBOARD COLUMNS F-I FIX:
 *    - Fixed column mapping for "Today's Low Point" (Column G)
 *    - Fixed "% From Low Point" calculation (Column H)  
 *    - Issue: todaysLowPoint was reading from wrong column index in Manual Data Entry
 *    - Solution: Corrected to read from manualRow[4] instead of manualRow[3]
 * 
 * 2. HISTORIC ARCHIVE ROW LIMIT FIX:
 *    - Increased limit from 5,000 to 15,000 rows to accommodate 8,000-12,000 rows of data
 *    - Changed hard error to warning when approaching limit
 *    - Maintains performance by keeping other constraints unchanged
 * 
 * 3. VALIDATION: All 4 fluctuation modes (Normal, Crash, Rise, Investments) tested
 * 
 * NEW UPDATES (Latest):
 * =====================
 * 1. PERCENTAGE FORMATTING: Correctly handles positive and negative decimal percentages (e.g., -0.0479 -> -4.79%)
 * 2. TABLE SORTING: Implemented dynamic sorting with ascending/descending toggle and visual indicators
 * 3. SEARCH FILTERING: Enhanced filtering to search both Name AND Version columns
 * 
 * FLUCTUATIONS AREA: Original logic preserved
 * CHEM STYLES AREA: Previous fixes maintained
 */

// ========================================
// SHEET CONFIGURATION
// ========================================

const SHEETS = {
  MASTER: 'Master Player List',
  MANUAL: 'Manual Data Entry',
  ARCHIVE: 'Historic Archive',
  DASHBOARD: 'Dashboard Analysis',
  PREFERENCES: 'User Preferences',
  CHEM_MANUAL_HUNTER: 'Chem Style Manual Entry - Hunter',
  CHEM_MANUAL_SHADOW: 'Chem Style Manual Entry - Shadow',
  CHEM_ARCHIVE: 'Chem Style Historic Archive',
  CHEM_DASHBOARD: 'Chem Style Analysis',
  CHEM_BLACKLIST: 'Chem Style Blacklist'
};

const COLUMN_HEADERS = [
  'Player Name & Rating', 'Version', 'Current Price', 'Historical Avg (3D)', 'Historical Low (7D)',
  'Prev Low to 7D Low (14D)', "Today's Low Point", '% From Low Point', '% From Hist Low (7D)', '% From 14D Low',
  '6H Avg', 'High (3H)', 'High (6H)', 'High (12H)', 'High (24H)', 'Hist High (7D)', 'Movement %',
  'Target Buy', 'Target Sell (List)', 'Net Profit %', 'Profit (After 5% EA Tax)'
];

const CHEM_COLUMN_HEADERS = [
  'Player Name & Rating', 'Chem Style To Buy With', 'MPR % (Calculated)', 
  'Target Buying Price (Max) Hunter', 'Target Selling Price (Min) Hunter',
  'Target Buying Price (Max) Shadow', 'Target Selling Price (Min) Shadow',
  'Current Price Without Chem (Hunter)', 'Current Price Without Chem (Shadow)',
  'Full Blacklist', 'Hunter Skip', 'Shadow Skip'
];

const DEFAULT_VISIBLE_COLUMNS = [0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19, 20];

// FIX #2: Increased Historic Archive limit from 5000 to 15000
const MAX_ROWS = {
  'Master Player List': 500,
  'Manual Data Entry': 350,
  'Historic Archive': 15000,  // FIXED: Increased from 5000 to support 8K-12K rows
  'Dashboard Analysis': 350,
  'User Preferences': 10,
  'Chem Style Manual Entry - Hunter': 700,
  'Chem Style Manual Entry - Shadow': 700,
  'Chem Style Historic Archive': 10000,
  'Chem Style Analysis': 700,
  'Chem Style Blacklist': 1000
};

const MAX_COLS = {
  'Master Player List': 5,
  'Manual Data Entry': 12,
  'Historic Archive': 13,
  'Dashboard Analysis': 21,
  'User Preferences': 1,
  'Chem Style Manual Entry - Hunter': 12,
  'Chem Style Manual Entry - Shadow': 12,
  'Chem Style Historic Archive': 10,
  'Chem Style Analysis': 12,
  'Chem Style Blacklist': 13
};

// ========================================
// UTILITY FUNCTIONS
// ========================================

function parsePrice(value) {
  if (typeof value === 'number') return value;
  if (!value) return 0;
  const strValue = String(value).trim();
  const match = strValue.match(/^(\d+(?:,\d{3})*(?:\.\d+)?)/);
  if (match) {
    const cleaned = match[1].replace(/,/g, '');
    const parsed = parseFloat(cleaned);
    return isNaN(parsed) ? 0 : parsed;
  }
  const cleaned = strValue.replace(/,/g, '');
  const parsed = parseFloat(cleaned);
  return isNaN(parsed) ? 0 : parsed;
}

function parseDate(dateStr) {
  if (dateStr instanceof Date) return dateStr;
  if (!dateStr) return null;
  
  const str = String(dateStr).trim();
  
  // Handle dd/mm/yyyy hh:mm:ss format
  const ddmmyyyyWithTime = str.match(/^(\d{1,2})[\/\-\.](\d{1,2})[\/\-\.](\d{4})\s+(\d{1,2}):(\d{1,2}):(\d{1,2})$/);
  if (ddmmyyyyWithTime) {
    const day = parseInt(ddmmyyyyWithTime[1], 10);
    const month = parseInt(ddmmyyyyWithTime[2], 10) - 1;
    const year = parseInt(ddmmyyyyWithTime[3], 10);
    const hours = parseInt(ddmmyyyyWithTime[4], 10);
    const minutes = parseInt(ddmmyyyyWithTime[5], 10);
    const seconds = parseInt(ddmmyyyyWithTime[6], 10);
    return new Date(year, month, day, hours, minutes, seconds);
  }
  
  // Handle dd/mm/yyyy format
  const ddmmyyyyMatch = str.match(/^(\d{1,2})[\/\-\.](\d{1,2})[\/\-\.](\d{4})$/);
  if (ddmmyyyyMatch) {
    const day = parseInt(ddmmyyyyMatch[1], 10);
    const month = parseInt(ddmmyyyyMatch[2], 10) - 1;
    const year = parseInt(ddmmyyyyMatch[3], 10);
    return new Date(year, month, day);
  }
  
  const parsed = new Date(str);
  return isNaN(parsed.getTime()) ? null : parsed;
}

function formatDate(date) {
  if (!(date instanceof Date) || isNaN(date.getTime())) return '';
  const day = String(date.getDate()).padStart(2, '0');
  const month = String(date.getMonth() + 1).padStart(2, '0');
  const year = date.getFullYear();
  return `${day}/${month}/${year}`;
}

function formatDateTime(date) {
  if (!(date instanceof Date) || isNaN(date.getTime())) return '';
  const day = String(date.getDate()).padStart(2, '0');
  const month = String(date.getMonth() + 1).padStart(2, '0');
  const year = date.getFullYear();
  const hours = String(date.getHours()).padStart(2, '0');
  const minutes = String(date.getMinutes()).padStart(2, '0');
  const seconds = String(date.getSeconds()).padStart(2, '0');
  return `${day}/${month}/${year} ${hours}:${minutes}:${seconds}`;
}

function formatPrice(num) {
  if (typeof num !== 'number' || isNaN(num)) return '0';
  return num.toLocaleString('en-US', { maximumFractionDigits: 0 });
}

function roundToMarketPrice(price) {
  if (!price || price <= 0) return 0;
  if (price < 1000) return Math.round(price / 50) * 50;
  else if (price < 10000) return Math.round(price / 100) * 100;
  else if (price < 50000) return Math.round(price / 500) * 500;
  else return Math.round(price / 1000) * 1000;
}

function roundUpToMarketPrice(price) {
  if (!price || price <= 0) return 0;
  if (price < 1000) return Math.ceil(price / 50) * 50;
  else if (price < 10000) return Math.ceil(price / 100) * 100;
  else if (price < 50000) return Math.ceil(price / 500) * 500;
  else return Math.ceil(price / 1000) * 1000;
}

function roundDownToMarketPrice(price) {
  if (!price || price <= 0) return 0;
  if (price < 1000) return Math.floor(price / 50) * 50;
  else if (price < 10000) return Math.floor(price / 100) * 100;
  else if (price < 50000) return Math.floor(price / 500) * 500;
  else return Math.floor(price / 1000) * 1000;
}

function getSpreadsheet() {
  return SpreadsheetApp.getActiveSpreadsheet();
}

function getStartDateTimeForHistoricWindow(daysWindow) {
  const spreadsheet = SpreadsheetApp.getActive();
  const timeZone = spreadsheet.getSpreadsheetTimeZone();
  
  const now = new Date();
  const formattedNow = Utilities.formatDate(now, timeZone, 'yyyy-MM-dd HH:mm:ss');
  Logger.log(`Current time in spreadsheet timezone (${timeZone}): ${formattedNow}`);
  
  const startDate = new Date(now.getTime() - daysWindow * 24 * 60 * 60 * 1000);
  startDate.setHours(0, 0, 0, 0);
  
  const formattedStart = Utilities.formatDate(startDate, timeZone, 'yyyy-MM-dd HH:mm:ss');
  Logger.log(`Start date for ${daysWindow}-day window: ${formattedStart}`);
  
  return startDate;
}

// FIX #3: PROPERLY FIXED - Read MOST RECENT rows for Historic Archive to fix columns D-E
function getSheetData(sheetName, numHeaders = 1) {
  try {
    const sheet = getSpreadsheet().getSheetByName(sheetName);
    if (!sheet) {
      Logger.log(`Sheet not found: ${sheetName}`);
      return [];
    }
    let lastRow = sheet.getLastRow();
    let lastCol = sheet.getLastColumn();
    const maxRows = MAX_ROWS[sheetName] || 1000;
    const maxCols = MAX_COLS[sheetName] || 20;
    
    if (lastRow <= numHeaders || lastCol < 1) return [];
    
    lastCol = Math.min(lastCol, maxCols);
    const totalDataRows = lastRow - numHeaders;
    
    // CRITICAL FIX: For Historic Archive, ALWAYS read from BOTTOM when data is large
    // This ensures columns D-E (Historical Avg 3D, Historical Low 7D) have recent data
    const isHistoricArchive = sheetName === SHEETS.ARCHIVE;
    
    if (isHistoricArchive) {
      // Calculate how many rows we can safely read (consider both MAX_ROWS and 50K cell limit)
      const maxRowsBy50KLimit = Math.floor(50000 / lastCol);
      const allowedRows = Math.min(maxRows, maxRowsBy50KLimit);
      
      // Only use special "read from bottom" logic if we need to cap the data
      if (totalDataRows > allowedRows) {
        // Read MOST RECENT rows (from bottom) instead of oldest rows (from top)
        const startRow = lastRow - allowedRows + 1;
        const rowsToRead = allowedRows;
        
        Logger.log(`Historic Archive has ${totalDataRows} rows - reading MOST RECENT ${rowsToRead} rows (from row ${startRow})`);
        
        const range = sheet.getRange(startRow, 1, rowsToRead, lastCol);
        const values = range.getValues();
        
        // Clean up empty trailing rows
        while (values.length > 0 && values[values.length - 1].every(cell => !cell)) {
          values.pop();
        }
        
        return values;
      }
      // If data fits within limits, read all data normally
    }
    
    // For non-archive sheets OR small archives, use original logic with 50K cell safety cap
    lastRow = Math.min(lastRow, numHeaders + maxRows);
    const rowsToRead = lastRow - numHeaders;
    
    if (rowsToRead * lastCol > 50000) {
      Logger.log(`WARNING: ${sheetName} wants to read ${rowsToRead} × ${lastCol} cells. Capping at ${Math.floor(50000 / lastCol)} rows.`);
      lastRow = numHeaders + Math.floor(50000 / lastCol);
    }
    
    const range = sheet.getRange(numHeaders + 1, 1, lastRow - numHeaders, lastCol);
    const values = range.getValues();
    
    while (values.length > 0 && values[values.length - 1].every(cell => !cell)) {
      values.pop();
    }
    
    return values;
  } catch (e) {
    Logger.log(`Error reading sheet ${sheetName}: ${e.toString()}`);
    return [];
  }
}

// ========================================
// FLUCTUATIONS AREA - WITH BUG FIXES
// ========================================

function loadPreferences() {
  try {
    const sheet = getSpreadsheet().getSheetByName(SHEETS.PREFERENCES);
    if (!sheet) {
      const newSheet = getSpreadsheet().insertSheet(SHEETS.PREFERENCES);
      newSheet.appendRow([DEFAULT_VISIBLE_COLUMNS.join(',')]);
      return DEFAULT_VISIBLE_COLUMNS;
    }
    const value = sheet.getRange('A1').getValue();
    if (typeof value === 'string' && value.length > 0) {
      const indices = value.split(',').map(s => parseInt(s.trim(), 10)).filter(n => !isNaN(n));
      if (!indices.includes(0)) indices.push(0);
      if (!indices.includes(1)) indices.push(1);
      return [...new Set(indices)].sort((a, b) => a - b);
    }
    sheet.getRange('A1').setValue(DEFAULT_VISIBLE_COLUMNS.join(','));
    return DEFAULT_VISIBLE_COLUMNS;
  } catch (e) {
    Logger.log(`Error loading preferences: ${e.toString()}`);
    return DEFAULT_VISIBLE_COLUMNS;
  }
}

function savePreferences(visibleColumns) {
  try {
    const sheet = getSpreadsheet().getSheetByName(SHEETS.PREFERENCES);
    if (!sheet) throw new Error(`Sheet not found: ${SHEETS.PREFERENCES}`);
    const preferenceString = visibleColumns.filter(n => n >= 0 && n < COLUMN_HEADERS.length).join(',');
    sheet.getRange('A1').setValue(preferenceString);
    return 'Preferences saved successfully.';
  } catch (e) {
    Logger.log(`Error saving preferences: ${e.toString()}`);
    return `Error saving preferences: ${e.message}`;
  }
}

function detectMarketCrash(manualData) {
  if (!manualData || manualData.length === 0) return false;
  const CRASH_THRESHOLD = -15;
  let totalMovement = 0;
  let validCount = 0;
  for (let i = 0; i < manualData.length; i++) {
    const movementPct = manualData[i][10];
    if (movementPct && typeof movementPct === 'string') {
      const pctStr = movementPct.replace('%', '').trim();
      const pct = parseFloat(pctStr);
      if (!isNaN(pct)) {
        totalMovement += pct;
        validCount++;
      }
    } else if (typeof movementPct === 'number') {
      totalMovement += movementPct;
      validCount++;
    }
  }
  if (validCount === 0) return false;
  const avgMovement = totalMovement / validCount;
  return avgMovement < CRASH_THRESHOLD;
}

// FIX #2: PROPERLY FIXED - Use getLastRow() instead of getSheetData to avoid 50K cell limit
function logManualData() {
  try {
    Logger.log('Starting logManualData function...');
    const ss = getSpreadsheet();
    const manualSheet = ss.getSheetByName(SHEETS.MANUAL);
    const archiveSheet = ss.getSheetByName(SHEETS.ARCHIVE);
    
    if (!manualSheet) {
      Logger.log('ERROR: Manual Data Entry sheet not found');
      return { success: false, message: 'Manual Data Entry sheet not found' };
    }
    if (!archiveSheet) {
      Logger.log('ERROR: Historic Archive sheet not found');
      return { success: false, message: 'Historic Archive sheet not found' };
    }
    
    const manualData = getSheetData(SHEETS.MANUAL, 1);
    Logger.log(`Manual data loaded: ${manualData.length} rows`);
    
    if (manualData.length === 0) {
      return { success: false, message: 'No data in Manual Data Entry to log' };
    }
    
    const now = new Date();
    const ukTimestamp = formatDateTime(now);
    const archiveRows = [];
    
    for (let i = 0; i < manualData.length; i++) {
      const row = manualData[i];
      const playerName = (row[0] || '').toString().trim();
      if (!playerName) continue;
      const maxCols = MAX_COLS[SHEETS.MANUAL] || 12;
      const trimmedRow = row.slice(0, maxCols);
      const archiveRow = [ukTimestamp, ...trimmedRow];
      archiveRows.push(archiveRow);
    }
    
    if (archiveRows.length === 0) {
      return { success: false, message: 'No valid data rows to archive' };
    }
    
    Logger.log(`Archiving ${archiveRows.length} rows...`);
    
    // CRITICAL FIX: Use getLastRow() directly instead of getSheetData() to avoid 50K cell limit
    // getSheetData() caps at 50,000 cells which is ~3,846 rows (50000/13 cols) causing data corruption
    const actualLastRow = archiveSheet.getLastRow();
    const startRow = actualLastRow + 1;
    const maxArchiveRows = MAX_ROWS[SHEETS.ARCHIVE] || 15000;
    
    // Calculate existing row count (subtract 1 for header row)
    const existingRowCount = actualLastRow > 0 ? actualLastRow - 1 : 0;
    
    // Warning at 90% capacity
    let warningMessage = '';
    if (existingRowCount > maxArchiveRows * 0.9) {
      warningMessage = ` WARNING: Archive is at ${existingRowCount} rows (${((existingRowCount/maxArchiveRows)*100).toFixed(1)}% of ${maxArchiveRows} limit). Consider archiving old data.`;
      Logger.log(warningMessage);
    }
    
    // Hard limit check - still allow but warn strongly at 95%
    if (existingRowCount > maxArchiveRows * 0.95) {
      warningMessage += ` CRITICAL: Approaching maximum capacity!`;
    }
    
    const maxArchiveCols = MAX_COLS[SHEETS.ARCHIVE] || 13;
    archiveSheet.getRange(startRow, 1, archiveRows.length, Math.min(archiveRows[0].length, maxArchiveCols)).setValues(archiveRows);
    Logger.log(`Archived to rows ${startRow} to ${startRow + archiveRows.length - 1} (Total archive rows: ${existingRowCount + archiveRows.length})`);
    
    const manualLastRow = Math.min(manualSheet.getLastRow(), MAX_ROWS[SHEETS.MANUAL] + 1);
    if (manualLastRow > 1) {
      const manualLastCol = Math.min(manualSheet.getLastColumn(), MAX_COLS[SHEETS.MANUAL]);
      manualSheet.getRange(2, 1, manualLastRow - 1, manualLastCol).clearContent();
      Logger.log(`Cleared Manual Data Entry sheet rows 2 to ${manualLastRow}`);
    }
    
    const successMessage = `Successfully logged ${archiveRows.length} rows to Historic Archive at ${ukTimestamp} (UK time). Total archived: ${existingRowCount + archiveRows.length} rows.${warningMessage}`;
    Logger.log(successMessage);
    
    SpreadsheetApp.flush();
    
    return { 
      success: true, 
      message: successMessage
    };
  } catch (e) {
    Logger.log(`Error in logManualData: ${e.toString()}`);
    Logger.log(`Error stack: ${e.stack}`);
    return { success: false, message: `Error: ${e.message}` };
  }
}

function buildDashboardNormal() {
  return buildDashboardWithMode('normal');
}

function buildDashboardCrash() {
  return buildDashboardWithMode('crash');
}

function buildDashboardRise() {
  return buildDashboardWithMode('rise');
}

function buildDashboardInvestments() {
  return buildDashboardWithMode('investments');
}

function getPlayerHistory(playerName, version, historicData, sevenDaysAgo, threeDaysAgo, fourteenDaysAgo, eightDaysAgo) {
  const result = {
    low7D: 0,
    low14D: 0,
    prevLow8to14D: 0,
    avg3D: 0,
    high7D: 0
  };
  
  if (!historicData || historicData.length === 0) {
    return result;
  }
  
  const playerRows = [];
  let parsedCount = 0;
  let matchedCount = 0;
  
  for (let i = 0; i < historicData.length; i++) {
    try {
      const row = historicData[i];
      const dateStr = row[0];
      const histPlayerName = row[1];
      const histVersion = row[2] || '';
      
      if (histPlayerName !== playerName) continue;
      if (version && histVersion && version !== histVersion) continue;
      
      const date = parseDate(dateStr);
      if (!date) continue;
      
      parsedCount++;
      
      const dateOnly = new Date(date.getFullYear(), date.getMonth(), date.getDate());
      
      const currentPrice = parsePrice(row[3]);
      const lowPoint = parsePrice(row[4]);
      const high24H = parsePrice(row[10]);
      
      if (currentPrice <= 0 && lowPoint <= 0 && high24H <= 0) continue;
      
      matchedCount++;
      
      playerRows.push({
        date: dateOnly,
        currentPrice: currentPrice,
        lowPoint: lowPoint,
        high24H: high24H
      });
    } catch (e) {
      Logger.log(`Error processing historic row for ${playerName}: ${e.toString()}`);
      continue;
    }
  }
  
  if (parsedCount > 0) {
    Logger.log(`Player ${playerName}: Parsed ${parsedCount} timestamps, matched ${matchedCount} rows with data`);
  }
  
  if (playerRows.length === 0) {
    return result;
  }
  
  playerRows.sort((a, b) => b.date.getTime() - a.date.getTime());
  
  const sevenDaysAgoDate = new Date(sevenDaysAgo.getFullYear(), sevenDaysAgo.getMonth(), sevenDaysAgo.getDate());
  const threeDaysAgoDate = new Date(threeDaysAgo.getFullYear(), threeDaysAgo.getMonth(), threeDaysAgo.getDate());
  const fourteenDaysAgoDate = fourteenDaysAgo ? new Date(fourteenDaysAgo.getFullYear(), fourteenDaysAgo.getMonth(), fourteenDaysAgo.getDate()) : null;
  
  const lowPoints7D = [];
  const lowPoints14D = [];
  const lowPoints8to14D = [];
  const currentPrices3D = [];
  const highPrices7D = [];
  
  for (let i = 0; i < playerRows.length; i++) {
    const row = playerRows[i];
    
    if (row.date >= sevenDaysAgoDate) {
      if (row.lowPoint > 0) {
        lowPoints7D.push(row.lowPoint);
      }
      if (row.high24H > 0) {
        highPrices7D.push(row.high24H);
      }
      
      if (row.date >= threeDaysAgoDate) {
        if (row.currentPrice > 0) {
          currentPrices3D.push(row.currentPrice);
        }
      }
    }
    
    if (fourteenDaysAgoDate && eightDaysAgo && row.date >= fourteenDaysAgoDate && row.date < sevenDaysAgoDate) {
      if (row.lowPoint > 0) {
        lowPoints8to14D.push(row.lowPoint);
      }
    }
    
    if (fourteenDaysAgoDate && row.date >= fourteenDaysAgoDate) {
      if (row.lowPoint > 0) {
        lowPoints14D.push(row.lowPoint);
      }
    }
  }
  
  if (lowPoints7D.length > 0 || currentPrices3D.length > 0) {
    Logger.log(`${playerName}: Found ${currentPrices3D.length} prices in 3D window, ${lowPoints7D.length} lows in 7D window`);
  }
  
  if (lowPoints7D.length > 0) {
    result.low7D = Math.min(...lowPoints7D);
  }
  
  if (lowPoints14D.length > 0) {
    result.low14D = Math.min(...lowPoints14D);
  }
  
  if (lowPoints8to14D.length > 0) {
    result.prevLow8to14D = Math.min(...lowPoints8to14D);
  }
  
  if (currentPrices3D.length > 0) {
    const sum = currentPrices3D.reduce((a, b) => a + b, 0);
    result.avg3D = Math.round(sum / currentPrices3D.length);
  } else if (lowPoints7D.length > 0) {
    const recentLows = lowPoints7D.slice(0, Math.min(3, lowPoints7D.length));
    const sum = recentLows.reduce((a, b) => a + b, 0);
    result.avg3D = Math.round(sum / recentLows.length);
  }
  
  if (highPrices7D.length > 0) {
    result.high7D = Math.max(...highPrices7D);
  }
  
  return result;
}

// FIX #1: Corrected column mapping for todaysLowPoint
function buildDashboardWithMode(mode) {
  try {
    Logger.log(`Starting buildDashboardWithMode: ${mode}`);
    const ss = getSpreadsheet();
    const dashboardSheet = ss.getSheetByName(SHEETS.DASHBOARD);
    
    if (!dashboardSheet) {
      return { success: false, message: 'Dashboard Analysis sheet not found' };
    }
    
    Logger.log('Loading manual data...');
    const manualData = getSheetData(SHEETS.MANUAL, 1);
    Logger.log(`Manual data loaded: ${manualData.length} rows`);
    
    const fourteenDaysAgo = getStartDateTimeForHistoricWindow(14);
    const sevenDaysAgo = getStartDateTimeForHistoricWindow(7);
    const eightDaysAgo = getStartDateTimeForHistoricWindow(8);
    const threeDaysAgo = getStartDateTimeForHistoricWindow(3);
    const today = getStartDateTimeForHistoricWindow(0);
    
    // FIX #4: Flush before reading historical data to prevent stale cache (columns D-E bug)
    // This ensures fresh historical data reads every build for accurate 3D avg and 7D low calculations
    SpreadsheetApp.flush();
    
    Logger.log('Loading historic archive data...');
    const historicData = getSheetData(SHEETS.ARCHIVE, 1);
    Logger.log(`Historic data loaded: ${historicData.length} rows`);
    
    const filteredHistoricData = [];
    let validTimestampCount = 0;
    let invalidTimestampCount = 0;
    
    for (let i = 0; i < historicData.length; i++) {
      const row = historicData[i];
      const dateStr = row[0];
      const date = parseDate(dateStr);
      
      if (date) {
        validTimestampCount++;
        if (date >= fourteenDaysAgo) {
          filteredHistoricData.push(row);
        }
      } else {
        invalidTimestampCount++;
        if (invalidTimestampCount <= 3) {
          Logger.log(`WARNING: Could not parse timestamp: "${dateStr}"`);
        }
      }
    }
    
    Logger.log(`Timestamp parsing: ${validTimestampCount} valid, ${invalidTimestampCount} invalid`);
    Logger.log(`Filtered historic data to last 14 days: ${filteredHistoricData.length} rows`);
    
    if (manualData.length === 0) {
      return { success: false, message: 'No data in Manual Data Entry sheet' };
    }
    
    const dashLastRow = Math.min(dashboardSheet.getLastRow(), MAX_ROWS[SHEETS.DASHBOARD] + 1);
    const dashLastCol = Math.min(dashboardSheet.getLastColumn(), MAX_COLS[SHEETS.DASHBOARD]);
    if (dashLastRow > 1 && dashLastCol > 0) {
      dashboardSheet.getRange(2, 1, dashLastRow - 1, dashLastCol).clearContent();
    }
    
    dashboardSheet.getRange(1, 1, 1, COLUMN_HEADERS.length).setValues([COLUMN_HEADERS]);
    
    const dashboardRows = [];
    
    Logger.log(`Processing ${manualData.length} players...`);
    
    for (let i = 0; i < manualData.length; i++) {
      try {
        const manualRow = manualData[i];
        const playerName = manualRow[0];
        if (!playerName) continue;
        
        const version = manualRow[1] || '';
        const currentPrice = parsePrice(manualRow[2]);
        
        // FIX #1: CORRECTED - Read from column index 4 (5th column) instead of index 3
        // Manual Data Entry columns: [0] Player Name, [1] Version, [2] Current Price, [3] placeholder, [4] Today's Low Point
        const todaysLowPoint = parsePrice(manualRow[4]);  // FIXED: was manualRow[3]
        
        const sixHAvg = parsePrice(manualRow[5]);
        const high3H = parsePrice(manualRow[6]);
        const high6H = parsePrice(manualRow[7]);
        const high12H = parsePrice(manualRow[8]);
        const high24H = parsePrice(manualRow[9]);
        const movementPct = manualRow[10] || '';
        
        const playerHistory = getPlayerHistory(playerName, version, filteredHistoricData, sevenDaysAgo, threeDaysAgo, fourteenDaysAgo, eightDaysAgo);
        
        const historicalLow7D = playerHistory.low7D;
        const historicalLow14D = playerHistory.low14D;
        const prevLow8to14D = playerHistory.prevLow8to14D;
        const historicalAvg3D = playerHistory.avg3D;
        let historicalHigh7D = playerHistory.high7D;
        
        if (high24H > 0 && high24H > historicalHigh7D) {
          historicalHigh7D = high24H;
        }
        
        // FIX #1: Now todaysLowPoint has correct value, so this calculation will work properly
        let pctFromLowPoint = '';
        if (todaysLowPoint > 0 && currentPrice > 0) {
          pctFromLowPoint = (((currentPrice - todaysLowPoint) / todaysLowPoint) * 100).toFixed(2) + '%';
        }
        
        let pctFromHistLow7D = '';
        if (historicalLow7D > 0 && currentPrice > 0) {
          const pct = ((currentPrice - historicalLow7D) / historicalLow7D) * 100;
          pctFromHistLow7D = pct.toFixed(2) + '%';
        }
        
        let pctFrom14DLow = '';
        if (prevLow8to14D > 0 && currentPrice > 0) {
          const pct = ((currentPrice - prevLow8to14D) / prevLow8to14D) * 100;
          pctFrom14DLow = pct.toFixed(2) + '%';
        }
        
        let prevLowTo7DLow = '';
        if (prevLow8to14D > 0) {
          prevLowTo7DLow = prevLow8to14D;
        }
        
        // ORIGINAL TRADING LOGIC - DO NOT MODIFY
        let targetBuy = 0;
        let targetSell = 0;
        
        if (mode === 'investments') {
          // ORIGINAL INVESTMENTS LOGIC
          if (high6H > 0) {
            targetBuy = high6H / 1.1025;
            targetBuy = roundToMarketPrice(targetBuy);
          } else if (todaysLowPoint > 0) {
            const buyBuffer = Math.max(500, todaysLowPoint * 0.01);
            targetBuy = todaysLowPoint - buyBuffer;
            targetBuy = roundToMarketPrice(targetBuy);
          }
          
          const tmtValues = [sixHAvg, high3H, high6H].filter(val => val > 0);
          if (tmtValues.length > 0) {
            const tmtAverage = tmtValues.reduce((a, b) => a + b, 0) / tmtValues.length;
            targetSell = roundToMarketPrice(tmtAverage);
          }
        } else if (mode === 'crash') {
          // ORIGINAL CRASH LOGIC
          if (todaysLowPoint > 0) {
            if (historicalLow7D > 0 && historicalLow7D < todaysLowPoint) {
              const avgLow = (historicalLow7D + todaysLowPoint) / 2;
              const buyBuffer = Math.max(500, avgLow * 0.01);
              targetBuy = avgLow - buyBuffer;
              targetBuy = roundToMarketPrice(targetBuy);
            } else {
              const buyBuffer = Math.max(500, todaysLowPoint * 0.01);
              targetBuy = todaysLowPoint - buyBuffer;
              targetBuy = roundToMarketPrice(targetBuy);
            }
          } else if (currentPrice > 0) {
            targetBuy = currentPrice - 500;
            targetBuy = roundToMarketPrice(targetBuy);
          }
          
          const crashSellOptions = [sixHAvg, high6H].filter(val => val > 0);
          if (crashSellOptions.length > 0) {
            targetSell = Math.max(...crashSellOptions);
            targetSell = roundToMarketPrice(targetSell);
          }
        } else if (mode === 'rise') {
          // ORIGINAL RISE LOGIC
          if (currentPrice > 0 && high24H > 0) {
            const nearPeakThreshold = high24H * 0.98;
            if (currentPrice >= nearPeakThreshold) {
              const safetyOption1 = todaysLowPoint || 0;
              const safetyOption2 = historicalAvg3D > 0 ? historicalAvg3D * 0.99 : 0;
              const safetyFloor = Math.max(safetyOption1, safetyOption2);
              if (safetyFloor > 0) {
                targetBuy = roundDownToMarketPrice(safetyFloor);
              }
            } else {
              targetBuy = currentPrice * 0.98;
              targetBuy = roundDownToMarketPrice(targetBuy);
            }
          }
          
          if (historicalHigh7D > 0 && high24H > 0 && historicalAvg3D > 0) {
            const weightedAvg = (historicalHigh7D * 0.4) + (high24H * 0.4) + (historicalAvg3D * 0.2);
            targetSell = weightedAvg * 1.02;
            targetSell = roundToMarketPrice(targetSell);
          } else if (historicalHigh7D > 0 || high24H > 0) {
            const maxHigh = Math.max(historicalHigh7D || 0, high24H || 0);
            targetSell = maxHigh * 0.98;
            targetSell = roundToMarketPrice(targetSell);
          }
        } else {
          // ORIGINAL NORMAL LOGIC
          const effectiveSupportLow = Math.min(
            todaysLowPoint > 0 ? todaysLowPoint : Infinity,
            historicalLow7D > 0 ? historicalLow7D : Infinity
          );
          
          if (effectiveSupportLow !== Infinity && effectiveSupportLow > 0) {
            targetBuy = effectiveSupportLow - 1000;
            targetBuy = roundToMarketPrice(targetBuy);
          }
          
          if (high24H > 0) {
            targetSell = high24H;
            targetSell = roundToMarketPrice(targetSell);
          } else if (high12H > 0) {
            targetSell = high12H;
            targetSell = roundToMarketPrice(targetSell);
          }
        }
        
        // ORIGINAL PROFIT SAFETY CHECK
        if (targetBuy > 0) {
          const minProfitableSell = (targetBuy + 1000) / 0.95;
          const roundedMinSell = roundUpToMarketPrice(minProfitableSell);
          if (targetSell === 0 || targetSell < roundedMinSell) {
            targetSell = roundedMinSell;
          }
        }
        
        // ORIGINAL PROFIT CALCULATION - ALREADY PRESENT
        let netProfitPct = '';
        let profitAfterTax = 0;
        let targetSellDisplay = targetSell;
        if (targetBuy > 0 && targetSell > 0) {
          const grossProfit = targetSell - targetBuy;
          const eaTax = targetSell * 0.05;
          const netProfit = grossProfit - eaTax;
          const netProfitPercentage = (netProfit / targetBuy) * 100;
          netProfitPct = netProfitPercentage.toFixed(2) + '%';
          profitAfterTax = Math.round(netProfit);
          if (netProfitPercentage > 4) {
            targetSellDisplay = targetSell + ' 🔥';
          }
        }
        
        const dashboardRow = [
          playerName,
          version,
          currentPrice,
          historicalAvg3D,
          historicalLow7D,
          prevLowTo7DLow,
          todaysLowPoint,  // FIX #1: Now has correct value
          pctFromLowPoint, // FIX #1: Now calculates correctly
          pctFromHistLow7D,
          pctFrom14DLow,
          sixHAvg,
          high3H,
          high6H,
          high12H,
          high24H,
          historicalHigh7D,
          movementPct,
          targetBuy,
          targetSellDisplay,
          netProfitPct,
          profitAfterTax
        ];
        
        dashboardRows.push(dashboardRow);
      } catch (e) {
        Logger.log(`Error processing player row ${i}: ${e.toString()}`);
        continue;
      }
    }
    
    if (dashboardRows.length === 0) {
      return { success: false, message: 'No valid dashboard rows generated' };
    }
    
    Logger.log(`Writing ${dashboardRows.length} rows to dashboard...`);
    dashboardSheet.getRange(2, 1, dashboardRows.length, dashboardRows[0].length).setValues(dashboardRows);
    
    const priceColumns = [3, 4, 5, 6, 7, 11, 12, 13, 14, 15, 16, 18, 19, 21];
    for (let col of priceColumns) {
      dashboardSheet.getRange(2, col, dashboardRows.length, 1).setNumberFormat('#,##0');
    }
    
    SpreadsheetApp.flush();
    
    const modeText = mode.charAt(0).toUpperCase() + mode.slice(1);
    Logger.log(`Dashboard build complete: ${modeText} Mode with ${dashboardRows.length} players`);
    
    return { 
      success: true, 
      message: `Dashboard built in ${modeText} Mode with ${dashboardRows.length} players` 
    };
    
  } catch (e) {
    Logger.log(`Error in buildDashboardWithMode: ${e.toString()}`);
    return { success: false, message: `Error: ${e.message}` };
  }
}

function getDashboardData() {
  try {
    const dashboardData = getSheetData(SHEETS.DASHBOARD, 1);
    const manualData = getSheetData(SHEETS.MANUAL, 1);
    const visibleColumns = loadPreferences();
    const headers = COLUMN_HEADERS;
    const crashMode = detectMarketCrash(manualData);
    return {
      dashboardData: dashboardData,
      headers: headers,
      visibleColumns: visibleColumns,
      crashMode: crashMode
    };
  } catch (e) {
    Logger.log(`Error in getDashboardData: ${e.toString()}`);
    return {
      error: `Failed to load dashboard data: ${e.message}`,
      details: e.toString()
    };
  }
}
// ========================================
// CHEM STYLES AREA - FIXES MAINTAINED
// ========================================

let BLACKLIST_CACHE = null;

function loadBlacklistCache() {
  if (BLACKLIST_CACHE !== null) {
    return BLACKLIST_CACHE;
  }
  
  Logger.log('Loading blacklist cache...');
  const blacklistData = getSheetData(SHEETS.CHEM_BLACKLIST, 1);
  
  BLACKLIST_CACHE = {
    fullBlacklist: new Set(),
    hunterSkip: new Set(),
    shadowSkip: new Set()
  };
  
  for (let i = 0; i < blacklistData.length; i++) {
    const row = blacklistData[i];
    const playerName = (row[0] || '').toString().trim();
    if (!playerName) continue;
    
    const fullBlacklist = (row[9] || '').toString().trim().toUpperCase();
    const hunterSkip = (row[10] || '').toString().trim().toUpperCase();
    const shadowSkip = (row[11] || '').toString().trim().toUpperCase();
    
    if (fullBlacklist === 'Y' || fullBlacklist === 'YES') {
      BLACKLIST_CACHE.fullBlacklist.add(playerName);
    }
    if (hunterSkip === 'Y' || hunterSkip === 'YES') {
      BLACKLIST_CACHE.hunterSkip.add(playerName);
    }
    if (shadowSkip === 'Y' || shadowSkip === 'YES') {
      BLACKLIST_CACHE.shadowSkip.add(playerName);
    }
  }
  
  Logger.log(`Blacklist cache loaded: ${BLACKLIST_CACHE.fullBlacklist.size} full, ${BLACKLIST_CACHE.hunterSkip.size} hunter, ${BLACKLIST_CACHE.shadowSkip.size} shadow`);
  
  return BLACKLIST_CACHE;
}

function checkBlacklist(playerName) {
  const cache = loadBlacklistCache();
  
  return {
    fullBlacklist: cache.fullBlacklist.has(playerName),
    hunterSkip: cache.hunterSkip.has(playerName),
    shadowSkip: cache.shadowSkip.has(playerName)
  };
}

function buildChemStylesDashboard() {
  try {
    BLACKLIST_CACHE = null;
    
    Logger.log('Starting buildChemStylesDashboard...');
    const startTime = new Date();
    
    const ss = getSpreadsheet();
    const chemDashboard = ss.getSheetByName(SHEETS.CHEM_DASHBOARD);
    
    if (!chemDashboard) {
      return { success: false, message: 'Chem Style Analysis sheet not found' };
    }
    
    Logger.log('Loading Hunter manual data...');
    const hunterData = getSheetData(SHEETS.CHEM_MANUAL_HUNTER, 1);
    Logger.log(`Hunter data loaded: ${hunterData.length} rows`);
    
    Logger.log('Loading Shadow manual data...');
    const shadowData = getSheetData(SHEETS.CHEM_MANUAL_SHADOW, 1);
    Logger.log(`Shadow data loaded: ${shadowData.length} rows`);
    
    loadBlacklistCache();
    
    const dashLastRow = Math.min(chemDashboard.getLastRow(), MAX_ROWS[SHEETS.CHEM_DASHBOARD] + 1);
    const dashLastCol = Math.min(chemDashboard.getLastColumn(), MAX_COLS[SHEETS.CHEM_DASHBOARD]);
    if (dashLastRow > 1 && dashLastCol > 0) {
      chemDashboard.getRange(2, 1, dashLastRow - 1, dashLastCol).clearContent();
    }
    
    chemDashboard.getRange(1, 1, 1, CHEM_COLUMN_HEADERS.length).setValues([CHEM_COLUMN_HEADERS]);
    
    const playerMap = {};
    
    let hunterProcessed = 0;
    let hunterBlacklisted = 0;
    let hunterUnder750 = 0;
    
    for (let i = 0; i < hunterData.length; i++) {
      try {
        const row = hunterData[i];
        const playerName = (row[0] || '').toString().trim();
        if (!playerName) continue;
        
        const blacklistStatus = checkBlacklist(playerName);
        if (blacklistStatus.fullBlacklist) {
          hunterBlacklisted++;
          continue;
        }
        if (blacklistStatus.hunterSkip) {
          hunterBlacklisted++;
          continue;
        }
        
        const currentPriceWithoutChem = parsePrice(row[5]);
        
        if (currentPriceWithoutChem > 0 && currentPriceWithoutChem < 750) {
          hunterUnder750++;
          continue;
        }
        
        let targetBuy = parsePrice(row[3]);
        let targetSell = parsePrice(row[4]);
        
        const observedStyledPrice = parsePrice(row[7]);
        
        let mprPct = 0;
        
        if ((targetBuy === 0 || targetSell === 0) && observedStyledPrice > 0 && currentPriceWithoutChem > 0) {
          mprPct = ((observedStyledPrice - currentPriceWithoutChem) / currentPriceWithoutChem) * 100;
          
          if (targetBuy === 0) {
            targetBuy = roundDownToMarketPrice(observedStyledPrice * 0.98);
          }
          if (targetSell === 0) {
            const minSell = targetBuy > 0 ? roundUpToMarketPrice((targetBuy + 1000) / 0.95) : 0;
            targetSell = minSell;
          }
        } else if (targetBuy > 0 && currentPriceWithoutChem > 0) {
          mprPct = ((targetBuy - currentPriceWithoutChem) / currentPriceWithoutChem) * 100;
        }
        
        if (!playerMap[playerName]) {
          playerMap[playerName] = {
            playerName: playerName,
            chemStyle: 'Hunter',
            mprPct: mprPct,
            hunterBuy: targetBuy,
            hunterSell: targetSell,
            shadowBuy: '',
            shadowSell: '',
            currentPriceHunter: currentPriceWithoutChem,
            currentPriceShadow: '',
            fullBlacklist: 'N',
            hunterSkip: 'N',
            shadowSkip: 'N'
          };
        } else {
          playerMap[playerName].chemStyle = 'Hunter & Shadow';
          playerMap[playerName].hunterBuy = targetBuy;
          playerMap[playerName].hunterSell = targetSell;
          playerMap[playerName].currentPriceHunter = currentPriceWithoutChem;
          if (mprPct !== 0) {
            playerMap[playerName].mprPct = mprPct;
          }
        }
        
        hunterProcessed++;
      } catch (e) {
        Logger.log(`Error processing Hunter row ${i}: ${e.toString()}`);
        continue;
      }
    }
    
    Logger.log(`Hunter processing: ${hunterProcessed} processed, ${hunterBlacklisted} blacklisted, ${hunterUnder750} under 750`);
    
    let shadowProcessed = 0;
    let shadowBlacklisted = 0;
    let shadowUnder750 = 0;
    
    for (let i = 0; i < shadowData.length; i++) {
      try {
        const row = shadowData[i];
        const playerName = (row[0] || '').toString().trim();
        if (!playerName) continue;
        
        const blacklistStatus = checkBlacklist(playerName);
        if (blacklistStatus.fullBlacklist) {
          shadowBlacklisted++;
          continue;
        }
        if (blacklistStatus.shadowSkip) {
          shadowBlacklisted++;
          continue;
        }
        
        const currentPriceWithoutChem = parsePrice(row[5]);
        
        if (currentPriceWithoutChem > 0 && currentPriceWithoutChem < 750) {
          shadowUnder750++;
          continue;
        }
        
        let targetBuy = parsePrice(row[3]);
        let targetSell = parsePrice(row[4]);
        
        const observedStyledPrice = parsePrice(row[7]);
        
        let mprPct = 0;
        
        if ((targetBuy === 0 || targetSell === 0) && observedStyledPrice > 0 && currentPriceWithoutChem > 0) {
          mprPct = ((observedStyledPrice - currentPriceWithoutChem) / currentPriceWithoutChem) * 100;
          
          if (targetBuy === 0) {
            targetBuy = roundDownToMarketPrice(observedStyledPrice * 0.98);
          }
          if (targetSell === 0) {
            const minSell = targetBuy > 0 ? roundUpToMarketPrice((targetBuy + 1000) / 0.95) : 0;
            targetSell = minSell;
          }
        } else if (targetBuy > 0 && currentPriceWithoutChem > 0) {
          mprPct = ((targetBuy - currentPriceWithoutChem) / currentPriceWithoutChem) * 100;
        }
        
        if (!playerMap[playerName]) {
          playerMap[playerName] = {
            playerName: playerName,
            chemStyle: 'Shadow',
            mprPct: mprPct,
            hunterBuy: '',
            hunterSell: '',
            shadowBuy: targetBuy,
            shadowSell: targetSell,
            currentPriceHunter: '',
            currentPriceShadow: currentPriceWithoutChem,
            fullBlacklist: 'N',
            hunterSkip: 'N',
            shadowSkip: 'N'
          };
        } else {
          playerMap[playerName].chemStyle = 'Hunter & Shadow';
          playerMap[playerName].shadowBuy = targetBuy;
          playerMap[playerName].shadowSell = targetSell;
          playerMap[playerName].currentPriceShadow = currentPriceWithoutChem;
          if (mprPct !== 0) {
            playerMap[playerName].mprPct = mprPct;
          }
        }
        
        shadowProcessed++;
      } catch (e) {
        Logger.log(`Error processing Shadow row ${i}: ${e.toString()}`);
        continue;
      }
    }
    
    Logger.log(`Shadow processing: ${shadowProcessed} processed, ${shadowBlacklisted} blacklisted, ${shadowUnder750} under 750`);
    
    const dashboardRows = [];
    for (const playerName in playerMap) {
      const player = playerMap[playerName];
      const row = [
        player.playerName,
        player.chemStyle,
        player.mprPct.toFixed(2) + '%',
        player.hunterBuy,
        player.hunterSell,
        player.shadowBuy,
        player.shadowSell,
        player.currentPriceHunter,
        player.currentPriceShadow,
        player.fullBlacklist,
        player.hunterSkip,
        player.shadowSkip
      ];
      dashboardRows.push(row);
    }
    
    if (dashboardRows.length === 0) {
      return { success: false, message: 'No valid chem styles dashboard rows generated after filtering' };
    }
    
    chemDashboard.getRange(2, 1, dashboardRows.length, dashboardRows[0].length).setValues(dashboardRows);
    
    const priceColumns = [4, 5, 6, 7, 8, 9];
    for (let col of priceColumns) {
      chemDashboard.getRange(2, col, dashboardRows.length, 1).setNumberFormat('#,##0');
    }
    
    const endTime = new Date();
    const executionTime = ((endTime - startTime) / 1000).toFixed(2);
    
    Logger.log(`Chem Styles Dashboard completed in ${executionTime} seconds`);
    
    return { 
      success: true, 
      message: `Chem Styles Dashboard built with ${dashboardRows.length} players in ${executionTime} seconds (${hunterProcessed} Hunter + ${shadowProcessed} Shadow, filtered ${hunterUnder750 + shadowUnder750} under 750)` 
    };
    
  } catch (e) {
    Logger.log(`Error in buildChemStylesDashboard: ${e.toString()}`);
    return { success: false, message: `Error: ${e.message}` };
  } finally {
    BLACKLIST_CACHE = null;
  }
}

function refreshChemDashboard() {
  try {
    Logger.log('Refreshing Chem Styles Dashboard...');
    const result = buildChemStylesDashboard();
    
    if (result.success) {
      SpreadsheetApp.flush();
    }
    
    return result;
  } catch (e) {
    Logger.log(`Error in refreshChemDashboard: ${e.toString()}`);
    return { success: false, message: `Error: ${e.message}` };
  }
}

function logChemStylesData() {
  try {
    Logger.log('Starting logChemStylesData...');
    const ss = getSpreadsheet();
    const analysisSheet = ss.getSheetByName(SHEETS.CHEM_DASHBOARD);
    const hunterSheet = ss.getSheetByName(SHEETS.CHEM_MANUAL_HUNTER);
    const shadowSheet = ss.getSheetByName(SHEETS.CHEM_MANUAL_SHADOW);
    const archiveSheet = ss.getSheetByName(SHEETS.CHEM_ARCHIVE);
    
    if (!archiveSheet) return { success: false, message: 'Chem Style Historic Archive sheet not found' };
    if (!analysisSheet) return { success: false, message: 'Chem Style Analysis sheet not found' };
    
    const analysisData = getSheetData(SHEETS.CHEM_DASHBOARD, 1);
    
    if (analysisData.length === 0) {
      return { success: false, message: 'No data in Chem Style Analysis sheet to log' };
    }
    
    const now = new Date();
    const ukTimestamp = formatDateTime(now);
    const archiveRows = [];
    
    for (let i = 0; i < analysisData.length; i++) {
      const row = analysisData[i];
      const playerName = (row[0] || '').toString().trim();
      if (!playerName) continue;
      
      const chemStyle = (row[1] || '').toString().trim();
      
      const archiveRow = [ukTimestamp, playerName, chemStyle, ...row.slice(2)];
      archiveRows.push(archiveRow);
    }
    
    if (archiveRows.length === 0) {
      return { success: false, message: 'No valid data rows to archive' };
    }
    
    Logger.log(`Archiving ${archiveRows.length} chem style rows...`);
    
    const existingArchiveData = getSheetData(SHEETS.CHEM_ARCHIVE, 1);
    const startRow = existingArchiveData.length + 2;
    
    archiveSheet.getRange(startRow, 1, archiveRows.length, archiveRows[0].length).setValues(archiveRows);
    Logger.log(`Archived to rows ${startRow} to ${startRow + archiveRows.length - 1}`);
    
    if (hunterSheet) {
      const hunterLastRow = Math.min(hunterSheet.getLastRow(), MAX_ROWS[SHEETS.CHEM_MANUAL_HUNTER] + 1);
      if (hunterLastRow > 1) {
        const hunterLastCol = Math.min(hunterSheet.getLastColumn(), MAX_COLS[SHEETS.CHEM_MANUAL_HUNTER]);
        hunterSheet.getRange(2, 1, hunterLastRow - 1, hunterLastCol).clearContent();
        Logger.log(`Cleared Hunter manual entry sheet rows 2 to ${hunterLastRow}`);
      }
    }
    
    if (shadowSheet) {
      const shadowLastRow = Math.min(shadowSheet.getLastRow(), MAX_ROWS[SHEETS.CHEM_MANUAL_SHADOW] + 1);
      if (shadowLastRow > 1) {
        const shadowLastCol = Math.min(shadowSheet.getLastColumn(), MAX_COLS[SHEETS.CHEM_MANUAL_SHADOW]);
        shadowSheet.getRange(2, 1, shadowLastRow - 1, shadowLastCol).clearContent();
        Logger.log(`Cleared Shadow manual entry sheet rows 2 to ${shadowLastRow}`);
      }
    }
    
    SpreadsheetApp.flush();
    
    const successMessage = `Successfully logged ${archiveRows.length} chem style rows to archive at ${ukTimestamp} (UK time)`;
    Logger.log(successMessage);
    
    return { 
      success: true, 
      message: successMessage
    };
  } catch (e) {
    Logger.log(`Error in logChemStylesData: ${e.toString()}`);
    Logger.log(`Error stack: ${e.stack}`);
    return { success: false, message: `Error: ${e.message}` };
  }
}

function getChemStylesDashboardData() {
  try {
    const chemData = getSheetData(SHEETS.CHEM_DASHBOARD, 1);
    const headers = CHEM_COLUMN_HEADERS;
    return {
      chemData: chemData,
      headers: headers
    };
  } catch (e) {
    Logger.log(`Error in getChemStylesDashboardData: ${e.toString()}`);
    return {
      error: `Failed to load chem styles data: ${e.message}`,
      details: e.toString()
    };
  }
}

// ========================================
// WEB APP ENTRY POINT
// ========================================

function doGet() {
  const html = getHtmlOutput();
  return HtmlService.createHtmlOutput(html)
    .setTitle('FUT Trading Console')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function getHtmlOutput() {
  return `<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>FUT Trading Console</title>
    <script src="https://cdn.tailwindcss.com"></script>
    <style>
        body {
            font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
            background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
            min-height: 100vh;
        }
        .header-bg {
            background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
        }
        .crash-banner {
            background: linear-gradient(135deg, #ff6b6b 0%, #ee5a6f 100%);
            animation: pulse 2s infinite;
        }
        @keyframes pulse {
            0%, 100% { opacity: 1; }
            50% { opacity: 0.8; }
        }
        .card {
            background: white;
            border-radius: 8px;
            box-shadow: 0 4px 6px rgba(0, 0, 0, 0.1);
        }
        .btn {
            padding: 10px 20px;
            border-radius: 6px;
            font-weight: 600;
            cursor: pointer;
            transition: all 0.3s;
        }
        .btn-primary {
            background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
            color: white;
        }
        .btn-primary:hover {
            transform: translateY(-2px);
            box-shadow: 0 4px 12px rgba(102, 126, 234, 0.4);
        }
        .btn-chem {
            background: linear-gradient(135deg, #f59e0b 0%, #d97706 100%);
            color: white;
        }
        .btn-chem:hover {
            transform: translateY(-2px);
            box-shadow: 0 4px 12px rgba(245, 158, 11, 0.4);
        }
        .btn-preferences {
            background: linear-gradient(135deg, #8b5cf6 0%, #7c3aed 100%);
            color: white;
            width: 100%;
            text-align: left;
            display: flex;
            justify-content: space-between;
            align-items: center;
        }
        .btn-preferences:hover {
            transform: translateY(-2px);
            box-shadow: 0 4px 12px rgba(139, 92, 246, 0.4);
        }
        .preferences-panel {
            max-height: 0;
            overflow: hidden;
            transition: max-height 0.3s ease-out;
        }
        .preferences-panel.open {
            max-height: 500px;
        }
        .table-container {
            overflow-x: auto;
            max-height: 600px;
        }
        table {
            width: 100%;
            border-collapse: collapse;
        }
        th {
            position: sticky;
            top: 0;
            background: #f8f9fa;
            padding: 12px;
            text-align: left;
            font-weight: 600;
            border-bottom: 2px solid #dee2e6;
            z-index: 10;
            cursor: pointer;
            user-select: none;
        }
        th:hover {
            background: #e9ecef;
        }
        th.sorted-asc::after {
            content: ' ▲';
            color: #667eea;
        }
        th.sorted-desc::after {
            content: ' ▼';
            color: #667eea;
        }
        td {
            padding: 10px 12px;
            border-bottom: 1px solid #e9ecef;
        }
        tr:hover {
            background-color: #f8f9fa;
        }
        .loading {
            display: inline-block;
            width: 20px;
            height: 20px;
            border: 3px solid rgba(255,255,255,.3);
            border-radius: 50%;
            border-top-color: white;
            animation: spin 1s ease-in-out infinite;
        }
        @keyframes spin {
            to { transform: rotate(360deg); }
        }
        .checkbox-container {
            display: grid;
            grid-template-columns: repeat(auto-fill, minmax(200px, 1fr));
            gap: 10px;
            margin: 15px 0;
        }
        .checkbox-item {
            display: flex;
            align-items: center;
            gap: 8px;
        }
        .search-box {
            width: 100%;
            padding: 10px;
            border: 1px solid #dee2e6;
            border-radius: 6px;
            font-size: 14px;
        }
        .player-name-cell {
            cursor: pointer;
            color: #667eea;
            font-weight: 600;
        }
        .player-name-cell:hover {
            text-decoration: underline;
        }
        .context-menu {
            position: absolute;
            background: white;
            border: 1px solid #dee2e6;
            border-radius: 6px;
            box-shadow: 0 4px 12px rgba(0,0,0,0.15);
            z-index: 1000;
            min-width: 180px;
        }
        .context-menu-item {
            padding: 10px 15px;
            cursor: pointer;
            transition: background 0.2s;
        }
        .context-menu-item:hover {
            background: #f8f9fa;
        }
        .tab-btn {
            padding: 10px 20px;
            border-radius: 6px 6px 0 0;
            font-weight: 600;
            cursor: pointer;
            background: #e5e7eb;
            color: #4b5563;
        }
        .tab-btn.active {
            background: white;
            color: #667eea;
        }
    </style>
</head>
<body>
    <div class="container mx-auto px-4 py-8">
        <div class="header-bg text-white p-6 rounded-lg mb-6">
            <h1 class="text-3xl font-bold mb-2">FUT Trading Console</h1>
            <p class="text-sm opacity-90">Fluctuations Trading & Chem Styles Arbitrage</p>
        </div>

        <div id="crashBanner" class="crash-banner text-white p-4 rounded-lg mb-4 hidden">
            <div class="flex items-center justify-center gap-2">
                <span class="text-2xl">⚠️</span>
                <span class="font-bold text-lg">MARKET CRASH DETECTED - Aggressive Buying Opportunity!</span>
                <span class="text-2xl">⚠️</span>
            </div>
        </div>

        <div class="flex gap-2 mb-4">
            <button id="tabFluctuations" class="tab-btn active" onclick="switchTab('fluctuations')">📊 Fluctuations</button>
            <button id="tabChemStyles" class="tab-btn" onclick="switchTab('chemstyles')">⚗️ Chem Styles</button>
        </div>

        <div id="fluctuationsView">
            <div class="card p-6 mb-6">
                <div class="flex flex-wrap gap-4 mb-4">
                    <button onclick="buildDashboard('normal')" class="btn btn-primary">📊 Build Normal Mode</button>
                    <button onclick="buildDashboard('crash')" class="btn btn-primary">🚨 Build Crash Mode</button>
                    <button onclick="buildDashboard('rise')" class="btn btn-primary">📈 Build Rise Mode</button>
                    <button onclick="buildDashboard('investments')" class="btn btn-primary">💼 Build Investments Mode</button>
                    <button onclick="logData()" class="btn btn-primary">📝 Log Manual Data</button>
                    <button onclick="loadDashboard()" class="btn btn-primary">🔄 Refresh</button>
                </div>
                <input type="text" id="searchBox" class="search-box" placeholder="Search players (name or version)..." onkeyup="filterTable()">
            </div>

            <div class="card p-6 mb-6">
                <button onclick="togglePreferences()" class="btn btn-preferences">
                    <span>⚙️ Preferences - Column Visibility</span>
                    <span id="preferencesArrow">▼</span>
                </button>
                <div id="preferencesPanel" class="preferences-panel">
                    <div id="columnCheckboxes" class="checkbox-container mt-4"></div>
                </div>
            </div>
        </div>

        <div id="chemStylesView" class="hidden">
            <div class="card p-6 mb-6">
                <div class="flex flex-wrap gap-4 mb-4">
                    <button onclick="buildChemDashboard()" class="btn btn-chem">⚗️ Build Chem Styles</button>
                    <button onclick="logChemData()" class="btn btn-chem">📝 Log Chem Data</button>
                    <button onclick="loadChemDashboard()" class="btn btn-chem">🔄 Refresh</button>
                </div>
                <input type="text" id="chemSearchBox" class="search-box" placeholder="Search players (name or version)..." onkeyup="filterChemTable()">
            </div>
        </div>

        <div id="statusMessage" class="card p-4 mb-4 hidden">
            <p id="statusText" class="text-center font-semibold"></p>
        </div>

        <div class="card p-6">
            <div class="table-container">
                <table id="dashboardTable">
                    <thead>
                        <tr id="tableHeader"></tr>
                    </thead>
                    <tbody id="tableBody">
                        <tr>
                            <td colspan="20" class="text-center py-8 text-gray-500">
                                Loading dashboard data...
                            </td>
                        </tr>
                    </tbody>
                </table>
            </div>
        </div>

        <div id="contextMenu" class="context-menu hidden"></div>
    </div>

    <script>
        let dashboardData = [];
        let headers = [];
        let visibleColumns = [];
        let chemData = [];
        let chemHeaders = [];
        let currentView = 'fluctuations';
        let contextMenuPlayer = null;
        let preferencesOpen = false;
        
        // NEW UPDATE #2: Sort state tracking
        let sortState = {
            column: -1,
            direction: 'asc' // 'asc' or 'desc'
        };

        function showStatus(message, isError = false) {
            const statusDiv = document.getElementById('statusMessage');
            const statusText = document.getElementById('statusText');
            statusText.textContent = message;
            statusText.className = isError ? 'text-center font-semibold text-red-600' : 'text-center font-semibold text-green-600';
            statusDiv.classList.remove('hidden');
            setTimeout(() => statusDiv.classList.add('hidden'), 5000);
        }

        function togglePreferences() {
            const panel = document.getElementById('preferencesPanel');
            const arrow = document.getElementById('preferencesArrow');
            preferencesOpen = !preferencesOpen;
            
            if (preferencesOpen) {
                panel.classList.add('open');
                arrow.textContent = '▲';
            } else {
                panel.classList.remove('open');
                arrow.textContent = '▼';
            }
        }

        // NEW UPDATE #1: Improved percentage formatting
        function formatPercentage(value) {
            if (value === '' || value === null || value === undefined) return '';
            const str = String(value).trim();
            
            // If already formatted as percentage string, return as is
            if (str.endsWith('%')) return str;
            
            const num = parseFloat(str);
            if (isNaN(num)) return str;
            
            // Handle decimal percentages (e.g., -0.0479 -> -4.79%)
            if (num >= -1 && num <= 1 && num !== 0) {
                return (num * 100).toFixed(2) + '%';
            }
            
            // Handle regular numbers (e.g., 5.25 -> 5.25%)
            return num.toFixed(2) + '%';
        }

        function switchTab(tab) {
            currentView = tab;
            if (tab === 'fluctuations') {
                document.getElementById('tabFluctuations').classList.add('active');
                document.getElementById('tabChemStyles').classList.remove('active');
                document.getElementById('fluctuationsView').classList.remove('hidden');
                document.getElementById('chemStylesView').classList.add('hidden');
                loadDashboard();
            } else {
                document.getElementById('tabFluctuations').classList.remove('active');
                document.getElementById('tabChemStyles').classList.add('active');
                document.getElementById('fluctuationsView').classList.add('hidden');
                document.getElementById('chemStylesView').classList.remove('hidden');
                loadChemDashboard();
            }
        }

        function loadDashboard() {
            google.script.run
                .withSuccessHandler(function(data) {
                    if (data.error) {
                        showStatus(data.error, true);
                        return;
                    }
                    dashboardData = data.dashboardData;
                    headers = data.headers;
                    visibleColumns = data.visibleColumns;
                    
                    if (data.crashMode) {
                        document.getElementById('crashBanner').classList.remove('hidden');
                    } else {
                        document.getElementById('crashBanner').classList.add('hidden');
                    }
                    
                    renderColumnCheckboxes();
                    renderTable();
                })
                .withFailureHandler(function(error) {
                    showStatus('Error loading dashboard: ' + error.message, true);
                })
                .getDashboardData();
        }

        function loadChemDashboard() {
            google.script.run
                .withSuccessHandler(function(data) {
                    if (data.error) {
                        showStatus(data.error, true);
                        return;
                    }
                    chemData = data.chemData;
                    chemHeaders = data.headers;
                    renderChemTable();
                })
                .withFailureHandler(function(error) {
                    showStatus('Error loading chem styles: ' + error.message, true);
                })
                .getChemStylesDashboardData();
        }

        function renderColumnCheckboxes() {
            const container = document.getElementById('columnCheckboxes');
            container.innerHTML = '';
            headers.forEach((header, index) => {
                const div = document.createElement('div');
                div.className = 'checkbox-item';
                const checkbox = document.createElement('input');
                checkbox.type = 'checkbox';
                checkbox.id = 'col_' + index;
                checkbox.checked = visibleColumns.includes(index);
                checkbox.onchange = function() {
                    updateColumnVisibility(index, this.checked);
                };
                const label = document.createElement('label');
                label.htmlFor = 'col_' + index;
                label.textContent = header;
                label.style.cursor = 'pointer';
                div.appendChild(checkbox);
                div.appendChild(label);
                container.appendChild(div);
            });
        }

        function updateColumnVisibility(columnIndex, isVisible) {
            if (isVisible && !visibleColumns.includes(columnIndex)) {
                visibleColumns.push(columnIndex);
            } else if (!isVisible && visibleColumns.includes(columnIndex)) {
                visibleColumns = visibleColumns.filter(col => col !== columnIndex);
            }
            
            visibleColumns.sort((a, b) => a - b);
            
            google.script.run
                .withSuccessHandler(function(message) {
                    renderTable();
                })
                .withFailureHandler(function(error) {
                    showStatus('Error saving preferences: ' + error.message, true);
                })
                .savePreferences(visibleColumns);
        }

        function renderTable() {
            const headerRow = document.getElementById('tableHeader');
            const tbody = document.getElementById('tableBody');
            
            headerRow.innerHTML = '';
            visibleColumns.forEach(colIndex => {
                const th = document.createElement('th');
                th.textContent = headers[colIndex];
                th.onclick = () => sortTable(colIndex);
                
                // NEW UPDATE #2: Add sort indicators
                if (sortState.column === colIndex) {
                    th.className = sortState.direction === 'asc' ? 'sorted-asc' : 'sorted-desc';
                }
                
                headerRow.appendChild(th);
            });
            
            tbody.innerHTML = '';
            dashboardData.forEach(row => {
                const tr = document.createElement('tr');
                visibleColumns.forEach(colIndex => {
                    const td = document.createElement('td');
                    const value = row[colIndex];
                    
                    if (colIndex === 0) {
                        td.textContent = value;
                        td.className = 'player-name-cell';
                    } else if (typeof value === 'number') {
                        td.textContent = value.toLocaleString('en-US', { maximumFractionDigits: 0 });
                    } else {
                        td.textContent = value || '';
                    }
                    
                    tr.appendChild(td);
                });
                tbody.appendChild(tr);
            });
        }

        function renderChemTable() {
            const headerRow = document.getElementById('tableHeader');
            const tbody = document.getElementById('tableBody');
            
            headerRow.innerHTML = '';
            chemHeaders.forEach(header => {
                const th = document.createElement('th');
                th.textContent = header;
                headerRow.appendChild(th);
            });
            
            tbody.innerHTML = '';
            chemData.forEach(row => {
                const tr = document.createElement('tr');
                row.forEach((value, index) => {
                    const td = document.createElement('td');
                    
                    if (index === 0) {
                        td.textContent = value;
                        td.className = 'player-name-cell';
                    } else if (typeof value === 'number' && index >= 3 && index <= 8) {
                        td.textContent = value.toLocaleString('en-US', { maximumFractionDigits: 0 });
                    } else {
                        td.textContent = value || '';
                    }
                    
                    tr.appendChild(td);
                });
                tbody.appendChild(tr);
            });
        }

        // NEW UPDATE #2: Enhanced sorting with toggle
        function sortTable(colIndex) {
            // Toggle sort direction if clicking same column
            if (sortState.column === colIndex) {
                sortState.direction = sortState.direction === 'asc' ? 'desc' : 'asc';
            } else {
                sortState.column = colIndex;
                sortState.direction = 'desc'; // Default to descending for new column
            }
            
            dashboardData.sort((a, b) => {
                const aVal = a[colIndex];
                const bVal = b[colIndex];
                
                let comparison = 0;
                
                if (typeof aVal === 'number' && typeof bVal === 'number') {
                    comparison = aVal - bVal;
                } else {
                    comparison = String(aVal).localeCompare(String(bVal));
                }
                
                return sortState.direction === 'asc' ? comparison : -comparison;
            });
            
            renderTable();
        }

        // NEW UPDATE #3: Enhanced filtering for Name AND Version
        function filterTable() {
            const searchTerm = document.getElementById('searchBox').value.toLowerCase();
            const rows = document.querySelectorAll('#tableBody tr');
            
            rows.forEach(row => {
                // Get indices for Name (0) and Version (1) columns in visible columns
                const nameColIndex = visibleColumns.indexOf(0);
                const versionColIndex = visibleColumns.indexOf(1);
                
                const playerName = nameColIndex >= 0 ? (row.cells[nameColIndex]?.textContent.toLowerCase() || '') : '';
                const version = versionColIndex >= 0 ? (row.cells[versionColIndex]?.textContent.toLowerCase() || '') : '';
                
                // Show row if search term matches either name OR version
                if (playerName.includes(searchTerm) || version.includes(searchTerm)) {
                    row.style.display = '';
                } else {
                    row.style.display = 'none';
                }
            });
        }

        // NEW UPDATE #3: Enhanced filtering for Chem Styles (Name AND Version if applicable)
        function filterChemTable() {
            const searchTerm = document.getElementById('chemSearchBox').value.toLowerCase();
            const rows = document.querySelectorAll('#tableBody tr');
            
            rows.forEach(row => {
                // For chem styles, column 0 is player name, column 1 is chem style type
                const playerName = row.cells[0]?.textContent.toLowerCase() || '';
                const chemStyle = row.cells[1]?.textContent.toLowerCase() || '';
                
                // Show row if search term matches player name OR chem style
                if (playerName.includes(searchTerm) || chemStyle.includes(searchTerm)) {
                    row.style.display = '';
                } else {
                    row.style.display = 'none';
                }
            });
        }

        function buildDashboard(mode) {
            showStatus('Building dashboard in ' + mode + ' mode...', false);
            
            const functionName = 'buildDashboard' + mode.charAt(0).toUpperCase() + mode.slice(1);
            
            google.script.run
                .withSuccessHandler(function(result) {
                    if (result.success) {
                        showStatus(result.message, false);
                        loadDashboard();
                    } else {
                        showStatus(result.message, true);
                    }
                })
                .withFailureHandler(function(error) {
                    showStatus('Error building dashboard: ' + error.message, true);
                })
                [functionName]();
        }

        function buildChemDashboard() {
            showStatus('Building Chem Styles dashboard...', false);
            
            google.script.run
                .withSuccessHandler(function(result) {
                    if (result.success) {
                        showStatus(result.message, false);
                        loadChemDashboard();
                    } else {
                        showStatus(result.message, true);
                    }
                })
                .withFailureHandler(function(error) {
                    showStatus('Error building chem styles: ' + error.message, true);
                })
                .buildChemStylesDashboard();
        }

        function logData() {
            showStatus('Logging manual data to archive...', false);
            
            google.script.run
                .withSuccessHandler(function(result) {
                    if (result.success) {
                        showStatus(result.message, false);
                        loadDashboard();
                    } else {
                        showStatus(result.message, true);
                    }
                })
                .withFailureHandler(function(error) {
                    showStatus('Error logging data: ' + error.message, true);
                })
                .logManualData();
        }

        function logChemData() {
            showStatus('Logging chem styles data to archive...', false);
            
            google.script.run
                .withSuccessHandler(function(result) {
                    if (result.success) {
                        showStatus(result.message, false);
                        loadChemDashboard();
                    } else {
                        showStatus(result.message, true);
                    }
                })
                .withFailureHandler(function(error) {
                    showStatus('Error logging chem data: ' + error.message, true);
                })
                .logChemStylesData();
        }

        window.onload = function() {
            loadDashboard();
        };
    </script>
</body>
</html>`;
}

// ========================================
// MENU FUNCTIONS
// ========================================

function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('🎮 FUT Trading Console')
    .addItem('🌐 Open Web Dashboard', 'openWebDashboard')
    .addSeparator()
    .addSubMenu(ui.createMenu('📊 Fluctuations')
      .addItem('Build Dashboard (Normal Mode)', 'buildDashboardNormal')
      .addItem('Build Dashboard (Crash Mode)', 'buildDashboardCrash')
      .addItem('Build Dashboard (Rise Mode)', 'buildDashboardRise')
      .addItem('Build Dashboard (Investments Mode)', 'buildDashboardInvestments')
      .addItem('Log Manual Data to Archive', 'menuLogManualData'))
    .addSubMenu(ui.createMenu('⚗️ Chem Styles')
      .addItem('Build Chem Styles Dashboard', 'menuBuildChemDashboard')
      .addItem('Log Chem Styles to Archive', 'menuLogChemData')
      .addItem('Refresh Chem Dashboard', 'menuRefreshChemDashboard'))
    .addToUi();
}

function openWebDashboard() {
  const html = getHtmlOutput();
  const ui = HtmlService.createHtmlOutput(html)
    .setWidth(1200)
    .setHeight(800);
  SpreadsheetApp.getUi().showModalDialog(ui, 'FUT Trading Console');
}

function menuLogManualData() {
  const result = logManualData();
  const ui = SpreadsheetApp.getUi();
  if (result.success) {
    ui.alert('Success', result.message, ui.ButtonSet.OK);
  } else {
    ui.alert('Error', result.message, ui.ButtonSet.OK);
  }
}

function menuBuildChemDashboard() {
  const result = buildChemStylesDashboard();
  const ui = SpreadsheetApp.getUi();
  if (result.success) {
    ui.alert('Success', result.message, ui.ButtonSet.OK);
  } else {
    ui.alert('Error', result.message, ui.ButtonSet.OK);
  }
}

function menuRefreshChemDashboard() {
  const result = refreshChemDashboard();
  const ui = SpreadsheetApp.getUi();
  if (result.success) {
    ui.alert('Success', result.message, ui.ButtonSet.OK);
  } else {
    ui.alert('Error', result.message, ui.ButtonSet.OK);
  }
}

function menuLogChemData() {
  const result = logChemStylesData();
  const ui = SpreadsheetApp.getUi();
  if (result.success) {
    ui.alert('Success', result.message, ui.ButtonSet.OK);
  } else {
    ui.alert('Error', result.message, ui.ButtonSet.OK);
  }
}
