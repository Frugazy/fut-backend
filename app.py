/**
 * FUT TRADING CONSOLE - Complete Fixed Apps Script
 * Combined: Fluctuations Trading + Chemistry Styles Arbitrage
 * 
 * CRITICAL FIXES APPLIED (October 10, 2025 - FINAL):
 * 1. WEB UI PERCENTAGES: Fixed JavaScript to multiply decimals by 100 for display
 * 2. % FROM 14D LOW: Now calculates in ALL modes (not just investments)
 * 3. PREV LOW TO 7D LOW: Now displays in ALL modes (not just crash/investments)
 * 4. CHEM STYLE BUILD: Fixed column mapping - verified correct indices
 * 5. DATE COMPARISON: Fixed timezone issues for 8-14 day data retrieval
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
  'Target Buy', 'Target Sell (List)', 'Net Profit %'
];

const CHEM_COLUMN_HEADERS = [
  'Player Name & Rating', 'Chem Style To Buy With', 'MPR % (Calculated)', 
  'Target Buying Price (Max) Hunter', 'Target Selling Price (Min) Hunter',
  'Target Buying Price (Max) Shadow', 'Target Selling Price (Min) Shadow',
  'Current Price Without Chem (Hunter)', 'Current Price Without Chem (Shadow)',
  'Full Blacklist', 'Hunter Skip', 'Shadow Skip'
];

const DEFAULT_VISIBLE_COLUMNS = [0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19];

const MAX_ROWS = {
  'Master Player List': 500,
  'Manual Data Entry': 350,
  'Historic Archive': 5000,
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
  'Dashboard Analysis': 20,
  'User Preferences': 1,
  'Chem Style Manual Entry - Hunter': 8,
  'Chem Style Manual Entry - Shadow': 8,
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
  // Extract first number from strings like "700 (700 Avg)"
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
    lastRow = Math.min(lastRow, numHeaders + maxRows);
    lastCol = Math.min(lastCol, maxCols);
    if (lastRow <= numHeaders || lastCol < 1) return [];
    const rowsToRead = lastRow - numHeaders;
    if (rowsToRead * lastCol > 50000) {
      Logger.log(`WARNING: ${sheetName} wants to read ${rowsToRead} × ${lastCol} cells. Capping.`);
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
// FLUCTUATIONS AREA
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

function logManualData() {
  try {
    const ss = getSpreadsheet();
    const manualSheet = ss.getSheetByName(SHEETS.MANUAL);
    const archiveSheet = ss.getSheetByName(SHEETS.ARCHIVE);
    if (!manualSheet) return { success: false, message: 'Manual Data Entry sheet not found' };
    if (!archiveSheet) return { success: false, message: 'Historic Archive sheet not found' };
    const manualData = getSheetData(SHEETS.MANUAL, 1);
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
    const existingArchiveData = getSheetData(SHEETS.ARCHIVE, 1);
    const startRow = existingArchiveData.length + 2;
    const maxArchiveRows = MAX_ROWS[SHEETS.ARCHIVE] || 5000;
    if (startRow + archiveRows.length > maxArchiveRows + 1) {
      return { 
        success: false, 
        message: `Archive is approaching maximum size (${maxArchiveRows} rows). Please clean up old data.` 
      };
    }
    const maxArchiveCols = MAX_COLS[SHEETS.ARCHIVE] || 13;
    archiveSheet.getRange(startRow, 1, archiveRows.length, Math.min(archiveRows[0].length, maxArchiveCols)).setValues(archiveRows);
    const manualLastRow = Math.min(manualSheet.getLastRow(), MAX_ROWS[SHEETS.MANUAL] + 1);
    if (manualLastRow > 1) {
      const manualLastCol = Math.min(manualSheet.getLastColumn(), MAX_COLS[SHEETS.MANUAL]);
      manualSheet.getRange(2, 1, manualLastRow - 1, manualLastCol).clearContent();
    }
    return { 
      success: true, 
      message: `Successfully logged ${archiveRows.length} rows to Historic Archive at ${ukTimestamp} (UK time)` 
    };
  } catch (e) {
    Logger.log(`Error in logManualData: ${e.toString()}`);
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
      
      // FIXED: Normalize to midnight for proper date comparison
      date.setHours(0, 0, 0, 0);
      const currentPrice = parsePrice(row[3]);
      const lowPoint = parsePrice(row[4]);
      const high24H = parsePrice(row[10]);
      
      if (currentPrice <= 0 && lowPoint <= 0 && high24H <= 0) continue;
      
      playerRows.push({
        date: date,
        currentPrice: currentPrice,
        lowPoint: lowPoint,
        high24H: high24H
      });
    } catch (e) {
      Logger.log(`Error processing historic row for ${playerName}: ${e.toString()}`);
      continue;
    }
  }
  
  if (playerRows.length === 0) {
    return result;
  }
  
  playerRows.sort((a, b) => b.date.getTime() - a.date.getTime());
  
  const lowPoints7D = [];
  const lowPoints14D = [];
  const lowPoints8to14D = [];
  const currentPrices3D = [];
  const highPrices7D = [];
  
  for (let i = 0; i < playerRows.length; i++) {
    const row = playerRows[i];
    
    if (row.date >= sevenDaysAgo) {
      if (row.lowPoint > 0) {
        lowPoints7D.push(row.lowPoint);
      }
      if (row.high24H > 0) {
        highPrices7D.push(row.high24H);
      }
      
      if (row.date >= threeDaysAgo) {
        if (row.currentPrice > 0) {
          currentPrices3D.push(row.currentPrice);
        }
      }
    }
    
    // FIXED: Use >= for lower bound to include day 14, < for upper bound to exclude day 7
    if (fourteenDaysAgo && eightDaysAgo && row.date >= fourteenDaysAgo && row.date < sevenDaysAgo) {
      if (row.lowPoint > 0) {
        lowPoints8to14D.push(row.lowPoint);
      }
    }
    
    if (fourteenDaysAgo && row.date >= fourteenDaysAgo) {
      if (row.lowPoint > 0) {
        lowPoints14D.push(row.lowPoint);
      }
    }
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

function buildDashboardWithMode(mode) {
  try {
    const ss = getSpreadsheet();
    const dashboardSheet = ss.getSheetByName(SHEETS.DASHBOARD);
    
    if (!dashboardSheet) {
      return { success: false, message: 'Dashboard Analysis sheet not found' };
    }
    
    const manualData = getSheetData(SHEETS.MANUAL, 1);
    const historicData = getSheetData(SHEETS.ARCHIVE, 1);
    
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
    const today = new Date();
    today.setHours(0, 0, 0, 0);
    const sevenDaysAgo = new Date(today.getTime() - 7 * 24 * 60 * 60 * 1000);
    const eightDaysAgo = new Date(today.getTime() - 8 * 24 * 60 * 60 * 1000);
    const threeDaysAgo = new Date(today.getTime() - 3 * 24 * 60 * 60 * 1000);
    const fourteenDaysAgo = new Date(today.getTime() - 14 * 24 * 60 * 60 * 1000);
    
    for (let i = 0; i < manualData.length; i++) {
      try {
        const manualRow = manualData[i];
        const playerName = manualRow[0];
        if (!playerName) continue;
        
        const version = manualRow[1] || '';
        const currentPrice = parsePrice(manualRow[2]);
        const todaysLowPoint = parsePrice(manualRow[3]);
        const sixHAvg = parsePrice(manualRow[5]);
        const high3H = parsePrice(manualRow[6]);
        const high6H = parsePrice(manualRow[7]);
        const high12H = parsePrice(manualRow[8]);
        const high24H = parsePrice(manualRow[9]);
        const movementPct = manualRow[10] || '';
        
        const playerHistory = getPlayerHistory(playerName, version, historicData, sevenDaysAgo, threeDaysAgo, fourteenDaysAgo, eightDaysAgo);
        
        const historicalLow7D = playerHistory.low7D;
        const historicalLow14D = playerHistory.low14D;
        const prevLow8to14D = playerHistory.prevLow8to14D;
        const historicalAvg3D = playerHistory.avg3D;
        let historicalHigh7D = playerHistory.high7D;
        
        if (high24H > 0 && high24H > historicalHigh7D) {
          historicalHigh7D = high24H;
        }
        
        // FIXED: % From Low Point - ALWAYS uses Today's Low Point
        let pctFromLowPoint = '';
        if (todaysLowPoint > 0 && currentPrice > 0) {
          pctFromLowPoint = (((currentPrice - todaysLowPoint) / todaysLowPoint) * 100).toFixed(2) + '%';
        }
        
        // FIXED: % From Hist Low (7D) - ALWAYS uses Historical Low (7D)
        let pctFromHistLow7D = '';
        if (historicalLow7D > 0 && currentPrice > 0) {
          const pct = ((currentPrice - historicalLow7D) / historicalLow7D) * 100;
          pctFromHistLow7D = pct.toFixed(2) + '%';
        }
        
        // FIXED: % From 14D Low - ALWAYS uses Prev Low to 7D Low (14D) which is the 8-14 day window
        let pctFrom14DLow = '';
        if (prevLow8to14D > 0 && currentPrice > 0) {
          const pct = ((currentPrice - prevLow8to14D) / prevLow8to14D) * 100;
          pctFrom14DLow = pct.toFixed(2) + '%';
        }
        
        // FIXED: Show Prev Low to 7D Low for ALL modes (removed mode gate)
        let prevLowTo7DLow = '';
        if (prevLow8to14D > 0) {
          prevLowTo7DLow = prevLow8to14D;
        }
        
        let targetBuy = 0;
        let targetSell = 0;
        
        if (mode === 'investments') {
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
        
        if (targetBuy > 0) {
          const minProfitableSell = (targetBuy + 1000) / 0.95;
          const roundedMinSell = roundUpToMarketPrice(minProfitableSell);
          if (targetSell === 0 || targetSell < roundedMinSell) {
            targetSell = roundedMinSell;
          }
        }
        
        let netProfitPct = '';
        let targetSellDisplay = targetSell;
        if (targetBuy > 0 && targetSell > 0) {
          const grossProfit = targetSell - targetBuy;
          const eaTax = targetSell * 0.05;
          const netProfit = grossProfit - eaTax;
          const netProfitPercentage = (netProfit / targetBuy) * 100;
          netProfitPct = netProfitPercentage.toFixed(2) + '%';
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
          todaysLowPoint,
          pctFromLowPoint,
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
          netProfitPct
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
    
    dashboardSheet.getRange(2, 1, dashboardRows.length, dashboardRows[0].length).setValues(dashboardRows);
    
    const priceColumns = [3, 4, 5, 6, 7, 11, 12, 13, 14, 15, 16, 18, 19];
    for (let col of priceColumns) {
      dashboardSheet.getRange(2, col, dashboardRows.length, 1).setNumberFormat('#,##0');
    }
    
    const modeText = mode.charAt(0).toUpperCase() + mode.slice(1);
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
// CHEM STYLES AREA (FIXED COLUMN MAPPING)
// ========================================

function checkBlacklist(playerName) {
  try {
    const blacklistData = getSheetData(SHEETS.CHEM_BLACKLIST, 1);
    if (blacklistData.length === 0) return { fullBlacklist: false, hunterSkip: false, shadowSkip: false };
    
    for (let i = 0; i < blacklistData.length; i++) {
      const row = blacklistData[i];
      const blPlayerName = (row[1] || '').toString().trim();
      
      if (blPlayerName === playerName) {
        const fullBlacklist = (row[10] || '').toString().toUpperCase() === 'Y';
        const hunterSkip = (row[11] || '').toString().toUpperCase() === 'Y';
        const shadowSkip = (row[12] || '').toString().toUpperCase() === 'Y';
        
        return {
          fullBlacklist: fullBlacklist,
          hunterSkip: hunterSkip,
          shadowSkip: shadowSkip
        };
      }
    }
    
    return { fullBlacklist: false, hunterSkip: false, shadowSkip: false };
  } catch (e) {
    Logger.log(`Error checking blacklist: ${e.toString()}`);
    return { fullBlacklist: false, hunterSkip: false, shadowSkip: false };
  }
}

function addToBlacklist(playerName, chemStyle, currentPriceHunter, currentPriceShadow) {
  try {
    const ss = getSpreadsheet();
    const blacklistSheet = ss.getSheetByName(SHEETS.CHEM_BLACKLIST);
    
    if (!blacklistSheet) {
      return { success: false, message: 'Chem Style Blacklist sheet not found' };
    }
    
    const existingData = getSheetData(SHEETS.CHEM_BLACKLIST, 1);
    
    for (let i = 0; i < existingData.length; i++) {
      const row = existingData[i];
      const blPlayerName = (row[1] || '').toString().trim();
      
      if (blPlayerName === playerName) {
        const rowNum = i + 2;
        
        if (chemStyle === 'Full') {
          blacklistSheet.getRange(rowNum, 11).setValue('Y');
          return { success: true, message: `${playerName} fully blacklisted` };
        } else if (chemStyle === 'Hunter') {
          blacklistSheet.getRange(rowNum, 12).setValue('Y');
          return { success: true, message: `${playerName} Hunter skip enabled` };
        } else if (chemStyle === 'Shadow') {
          blacklistSheet.getRange(rowNum, 13).setValue('Y');
          return { success: true, message: `${playerName} Shadow skip enabled` };
        }
      }
    }
    
    const now = new Date();
    const ukTimestamp = formatDateTime(now);
    
    const fullBL = chemStyle === 'Full' ? 'Y' : 'N';
    const hunterSkip = chemStyle === 'Hunter' ? 'Y' : 'N';
    const shadowSkip = chemStyle === 'Shadow' ? 'Y' : 'N';
    
    const newRow = [
      ukTimestamp,
      playerName,
      chemStyle,
      '',
      '',
      '',
      '',
      '',
      currentPriceHunter || '',
      currentPriceShadow || '',
      fullBL,
      hunterSkip,
      shadowSkip
    ];
    
    const lastRow = blacklistSheet.getLastRow();
    blacklistSheet.getRange(lastRow + 1, 1, 1, newRow.length).setValues([newRow]);
    
    return { success: true, message: `${playerName} added to blacklist (${chemStyle})` };
  } catch (e) {
    Logger.log(`Error adding to blacklist: ${e.toString()}`);
    return { success: false, message: `Error: ${e.message}` };
  }
}

function buildChemStylesDashboard() {
  try {
    const ss = getSpreadsheet();
    const chemDashboard = ss.getSheetByName(SHEETS.CHEM_DASHBOARD);
    
    if (!chemDashboard) {
      return { success: false, message: 'Chem Style Analysis sheet not found' };
    }
    
    const hunterData = getSheetData(SHEETS.CHEM_MANUAL_HUNTER, 1);
    const shadowData = getSheetData(SHEETS.CHEM_MANUAL_SHADOW, 1);
    
    if (hunterData.length === 0 && shadowData.length === 0) {
      return { success: false, message: 'No data in Chem Style Manual Entry sheets' };
    }
    
    const dashLastRow = Math.min(chemDashboard.getLastRow(), MAX_ROWS[SHEETS.CHEM_DASHBOARD] + 1);
    const dashLastCol = Math.min(chemDashboard.getLastColumn(), MAX_COLS[SHEETS.CHEM_DASHBOARD]);
    if (dashLastRow > 1 && dashLastCol > 0) {
      chemDashboard.getRange(2, 1, dashLastRow - 1, dashLastCol).clearContent();
    }
    
    chemDashboard.getRange(1, 1, 1, CHEM_COLUMN_HEADERS.length).setValues([CHEM_COLUMN_HEADERS]);
    
    const playerMap = {};
    
    // FIXED: Corrected Hunter sheet column mapping
    // Expected columns: Player Name & Rating | Chem Style To Buy With | MPR % (Calculated) | Target Buying Price (Max) | Target Selling Price (Min) | Current Market Price Without Chem Style (Est) | Observed Styled Price (Manual)
    for (let i = 0; i < hunterData.length; i++) {
      try {
        const row = hunterData[i];
        const playerName = (row[0] || '').toString().trim();
        if (!playerName) continue;
        
        const blacklistStatus = checkBlacklist(playerName);
        if (blacklistStatus.fullBlacklist) continue;
        if (blacklistStatus.hunterSkip) continue;
        
        // FIXED: Read values from correct columns - Observed Styled Price is in Column H (index 7)
        const currentPriceWithoutChem = parsePrice(row[5]); // Column F
        const observedStyledPrice = parsePrice(row[7]); // Column H
        
        let mprPct = 0;
        let targetBuy = 0;
        let targetSell = 0;
        
        if (observedStyledPrice > 0 && currentPriceWithoutChem > 0) {
          mprPct = ((observedStyledPrice - currentPriceWithoutChem) / currentPriceWithoutChem) * 100;
          targetBuy = roundDownToMarketPrice(observedStyledPrice * 0.98);
          const minSell = targetBuy > 0 ? roundUpToMarketPrice((targetBuy + 1000) / 0.95) : 0;
          targetSell = minSell;
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
        }
      } catch (e) {
        Logger.log(`Error processing Hunter row ${i}: ${e.toString()}`);
        continue;
      }
    }
    
    // FIXED: Corrected Shadow sheet column mapping (same structure as Hunter)
    for (let i = 0; i < shadowData.length; i++) {
      try {
        const row = shadowData[i];
        const playerName = (row[0] || '').toString().trim();
        if (!playerName) continue;
        
        const blacklistStatus = checkBlacklist(playerName);
        if (blacklistStatus.fullBlacklist) continue;
        if (blacklistStatus.shadowSkip) continue;
        
        // FIXED: Read values from correct columns - Observed Styled Price is in Column H (index 7)
        const currentPriceWithoutChem = parsePrice(row[5]); // Column F
        const observedStyledPrice = parsePrice(row[7]); // Column H
        
        let mprPct = 0;
        let targetBuy = 0;
        let targetSell = 0;
        
        if (observedStyledPrice > 0 && currentPriceWithoutChem > 0) {
          mprPct = ((observedStyledPrice - currentPriceWithoutChem) / currentPriceWithoutChem) * 100;
          targetBuy = roundDownToMarketPrice(observedStyledPrice * 0.98);
          const minSell = targetBuy > 0 ? roundUpToMarketPrice((targetBuy + 1000) / 0.95) : 0;
          targetSell = minSell;
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
        }
      } catch (e) {
        Logger.log(`Error processing Shadow row ${i}: ${e.toString()}`);
        continue;
      }
    }
    
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
      return { success: false, message: 'No valid chem styles dashboard rows generated' };
    }
    
    chemDashboard.getRange(2, 1, dashboardRows.length, dashboardRows[0].length).setValues(dashboardRows);
    
    const priceColumns = [4, 5, 6, 7, 8, 9];
    for (let col of priceColumns) {
      chemDashboard.getRange(2, col, dashboardRows.length, 1).setNumberFormat('#,##0');
    }
    
    return { 
      success: true, 
      message: `Chem Styles Dashboard built with ${dashboardRows.length} players` 
    };
    
  } catch (e) {
    Logger.log(`Error in buildChemStylesDashboard: ${e.toString()}`);
    return { success: false, message: `Error: ${e.message}` };
  }
}

function logChemStylesData() {
  try {
    const ss = getSpreadsheet();
    const hunterSheet = ss.getSheetByName(SHEETS.CHEM_MANUAL_HUNTER);
    const shadowSheet = ss.getSheetByName(SHEETS.CHEM_MANUAL_SHADOW);
    const archiveSheet = ss.getSheetByName(SHEETS.CHEM_ARCHIVE);
    
    if (!archiveSheet) return { success: false, message: 'Chem Style Historic Archive sheet not found' };
    
    const hunterData = hunterSheet ? getSheetData(SHEETS.CHEM_MANUAL_HUNTER, 1) : [];
    const shadowData = shadowSheet ? getSheetData(SHEETS.CHEM_MANUAL_SHADOW, 1) : [];
    
    if (hunterData.length === 0 && shadowData.length === 0) {
      return { success: false, message: 'No data in Chem Style Manual Entry sheets' };
    }
    
    const now = new Date();
    const ukTimestamp = formatDateTime(now);
    const archiveRows = [];
    
    for (let i = 0; i < hunterData.length; i++) {
      const row = hunterData[i];
      const playerName = (row[0] || '').toString().trim();
      if (!playerName) continue;
      
      const archiveRow = [ukTimestamp, playerName, 'Hunter', ...row.slice(1)];
      archiveRows.push(archiveRow);
    }
    
    for (let i = 0; i < shadowData.length; i++) {
      const row = shadowData[i];
      const playerName = (row[0] || '').toString().trim();
      if (!playerName) continue;
      
      const archiveRow = [ukTimestamp, playerName, 'Shadow', ...row.slice(1)];
      archiveRows.push(archiveRow);
    }
    
    if (archiveRows.length === 0) {
      return { success: false, message: 'No valid data rows to archive' };
    }
    
    const existingArchiveData = getSheetData(SHEETS.CHEM_ARCHIVE, 1);
    const startRow = existingArchiveData.length + 2;
    
    archiveSheet.getRange(startRow, 1, archiveRows.length, archiveRows[0].length).setValues(archiveRows);
    
    if (hunterSheet) {
      const hunterLastRow = Math.min(hunterSheet.getLastRow(), MAX_ROWS[SHEETS.CHEM_MANUAL_HUNTER] + 1);
      if (hunterLastRow > 1) {
        const hunterLastCol = Math.min(hunterSheet.getLastColumn(), MAX_COLS[SHEETS.CHEM_MANUAL_HUNTER]);
        hunterSheet.getRange(2, 1, hunterLastRow - 1, hunterLastCol).clearContent();
      }
    }
    
    if (shadowSheet) {
      const shadowLastRow = Math.min(shadowSheet.getLastRow(), MAX_ROWS[SHEETS.CHEM_MANUAL_SHADOW] + 1);
      if (shadowLastRow > 1) {
        const shadowLastCol = Math.min(shadowSheet.getLastColumn(), MAX_COLS[SHEETS.CHEM_MANUAL_SHADOW]);
        shadowSheet.getRange(2, 1, shadowLastRow - 1, shadowLastCol).clearContent();
      }
    }
    
    return { 
      success: true, 
      message: `Successfully logged ${archiveRows.length} chem style rows to archive at ${ukTimestamp} (UK time)` 
    };
  } catch (e) {
    Logger.log(`Error in logChemStylesData: ${e.toString()}`);
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
                <input type="text" id="searchBox" class="search-box" placeholder="Search players..." onkeyup="filterTable()">
            </div>

            <div class="card p-6 mb-6">
                <h3 class="font-bold mb-3">Visible Columns:</h3>
                <div id="columnCheckboxes" class="checkbox-container"></div>
            </div>
        </div>

        <div id="chemStylesView" class="hidden">
            <div class="card p-6 mb-6">
                <div class="flex flex-wrap gap-4 mb-4">
                    <button onclick="buildChemDashboard()" class="btn btn-chem">⚗️ Build Chem Styles</button>
                    <button onclick="logChemData()" class="btn btn-chem">📝 Log Chem Data</button>
                    <button onclick="loadChemDashboard()" class="btn btn-chem">🔄 Refresh</button>
                </div>
                <input type="text" id="chemSearchBox" class="search-box" placeholder="Search players..." onkeyup="filterChemTable()">
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

        function showStatus(message, isError = false) {
            const statusDiv = document.getElementById('statusMessage');
            const statusText = document.getElementById('statusText');
            statusText.textContent = message;
            statusText.className = isError ? 'text-center font-semibold text-red-600' : 'text-center font-semibold text-green-600';
            statusDiv.classList.remove('hidden');
            setTimeout(() => statusDiv.classList.add('hidden'), 5000);
        }

        // FIXED: Helper function to format percentage values for display
        function formatPercentage(value) {
            if (value === '' || value === null || value === undefined) return '';
            const str = String(value).trim();
            
            // Already formatted with %
            if (str.endsWith('%')) return str;
            
            // Raw decimal like 0.086 needs conversion
            const num = parseFloat(str);
            if (!isNaN(num) && num >= -1 && num <= 1) {
                return (num * 100).toFixed(2) + '%';
            }
            
            // Already a percentage number without symbol
            if (!isNaN(num)) {
                return num.toFixed(2) + '%';
            }
            
            return str;
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
                checkbox.onchange = () => toggleColumn(index);
                const label = document.createElement('label');
                label.htmlFor = 'col_' + index;
                label.textContent = header;
                label.className = 'cursor-pointer';
                div.appendChild(checkbox);
                div.appendChild(label);
                container.appendChild(div);
            });
        }

        function toggleColumn(index) {
            const checkboxIndex = visibleColumns.indexOf(index);
            if (checkboxIndex > -1) {
                visibleColumns.splice(checkboxIndex, 1);
            } else {
                visibleColumns.push(index);
                visibleColumns.sort((a, b) => a - b);
            }
            google.script.run
                .withSuccessHandler(function(result) {
                    renderTable();
                })
                .savePreferences(visibleColumns);
        }

        let sortColumn = -1;
        let sortDirection = 'asc';
        let chemSortColumn = -1;
        let chemSortDirection = 'asc';

        function sortTable(colIndex) {
            if (sortColumn === colIndex) {
                sortDirection = sortDirection === 'asc' ? 'desc' : 'asc';
            } else {
                sortColumn = colIndex;
                sortDirection = 'asc';
            }
            
            dashboardData.sort((a, b) => {
                let valA = a[colIndex];
                let valB = b[colIndex];
                
                const numA = parseFloat(String(valA).replace(/[,%🔥]/g, ''));
                const numB = parseFloat(String(valB).replace(/[,%🔥]/g, ''));
                
                if (!isNaN(numA) && !isNaN(numB)) {
                    return sortDirection === 'asc' ? numA - numB : numB - numA;
                }
                
                valA = String(valA).toLowerCase();
                valB = String(valB).toLowerCase();
                if (valA < valB) return sortDirection === 'asc' ? -1 : 1;
                if (valA > valB) return sortDirection === 'asc' ? 1 : -1;
                return 0;
            });
            
            renderTable();
        }

        function sortChemTable(colIndex) {
            if (chemSortColumn === colIndex) {
                chemSortDirection = chemSortDirection === 'asc' ? 'desc' : 'asc';
            } else {
                chemSortColumn = colIndex;
                chemSortDirection = 'asc';
            }
            
            chemData.sort((a, b) => {
                let valA = a[colIndex];
                let valB = b[colIndex];
                
                const numA = parseFloat(String(valA).replace(/[,%]/g, ''));
                const numB = parseFloat(String(valB).replace(/[,%]/g, ''));
                
                if (!isNaN(numA) && !isNaN(numB)) {
                    return chemSortDirection === 'asc' ? numA - numB : numB - numA;
                }
                
                valA = String(valA).toLowerCase();
                valB = String(valB).toLowerCase();
                if (valA < valB) return chemSortDirection === 'asc' ? -1 : 1;
                if (valA > valB) return chemSortDirection === 'asc' ? 1 : -1;
                return 0;
            });
            
            renderChemTable();
        }

        function renderTable() {
            const headerRow = document.getElementById('tableHeader');
            const tableBody = document.getElementById('tableBody');
            headerRow.innerHTML = '';
            tableBody.innerHTML = '';
            
            // FIXED: Identify percentage columns for formatting
            const percentageColumns = [7, 8, 9, 19]; // % From Low Point, % From Hist Low (7D), % From 14D Low, Net Profit %
            
            visibleColumns.forEach(colIndex => {
                const th = document.createElement('th');
                th.textContent = headers[colIndex];
                th.onclick = () => sortTable(colIndex);
                if (sortColumn === colIndex) {
                    th.className = sortDirection === 'asc' ? 'sorted-asc' : 'sorted-desc';
                }
                headerRow.appendChild(th);
            });
            
            dashboardData.forEach(row => {
                const tr = document.createElement('tr');
                visibleColumns.forEach(colIndex => {
                    const td = document.createElement('td');
                    let value = row[colIndex];
                    
                    // FIXED: Apply percentage formatting to percentage columns
                    if (percentageColumns.includes(colIndex)) {
                        value = formatPercentage(value);
                    }
                    
                    td.textContent = value === 0 || value === '' ? '' : value;
                    tr.appendChild(td);
                });
                tableBody.appendChild(tr);
            });
        }

        function renderChemTable() {
            const headerRow = document.getElementById('tableHeader');
            const tableBody = document.getElementById('tableBody');
            headerRow.innerHTML = '';
            tableBody.innerHTML = '';
            
            chemHeaders.forEach((header, index) => {
                if (index >= 9) return;
                const th = document.createElement('th');
                th.textContent = header;
                th.onclick = () => sortChemTable(index);
                if (chemSortColumn === index) {
                    th.className = chemSortDirection === 'asc' ? 'sorted-asc' : 'sorted-desc';
                }
                headerRow.appendChild(th);
            });
            
            chemData.forEach(row => {
                const tr = document.createElement('tr');
                row.forEach((value, index) => {
                    if (index >= 9) return;
                    const td = document.createElement('td');
                    if (index === 0) {
                        td.className = 'player-name-cell';
                        td.textContent = value;
                        td.onclick = (e) => showContextMenu(e, value, row[7], row[8]);
                    } else {
                        td.textContent = value === 0 || value === '' ? '' : value;
                    }
                    tr.appendChild(td);
                });
                tableBody.appendChild(tr);
            });
        }

        function showContextMenu(event, playerName, priceHunter, priceShadow) {
            event.preventDefault();
            event.stopPropagation();
            
            contextMenuPlayer = { name: playerName, priceHunter: priceHunter, priceShadow: priceShadow };
            const menu = document.getElementById('contextMenu');
            menu.innerHTML = \`
                <div class="context-menu-item" onclick="showObservedPricePrompt()">💰 Add Observed Price</div>
                <div class="context-menu-item" onclick="blacklistPlayer('Full')">🚫 Full Blacklist</div>
                <div class="context-menu-item" onclick="blacklistPlayer('Hunter')">🎯 Hunter Skip</div>
                <div class="context-menu-item" onclick="blacklistPlayer('Shadow')">👤 Shadow Skip</div>
            \`;
            
            menu.style.left = event.pageX + 'px';
            menu.style.top = event.pageY + 'px';
            menu.classList.remove('hidden');
        }
        
        function showObservedPricePrompt() {
            document.getElementById('contextMenu').classList.add('hidden');
            if (!contextMenuPlayer) return;
            
            const chemType = prompt('Enter chem style type (Hunter or Shadow):');
            if (!chemType || (chemType.toLowerCase() !== 'hunter' && chemType.toLowerCase() !== 'shadow')) {
                showStatus('Invalid chem style. Use Hunter or Shadow.', true);
                return;
            }
            
            const price = prompt('Enter observed price for ' + contextMenuPlayer.name + ' with ' + chemType + ':');
            if (!price || isNaN(parseFloat(price))) {
                showStatus('Invalid price entered.', true);
                return;
            }
            
            showStatus('Observed price: ' + price + ' for ' + contextMenuPlayer.name + ' (' + chemType + ')');
        }

        function blacklistPlayer(chemStyle) {
            if (!contextMenuPlayer) return;
            
            google.script.run
                .withSuccessHandler(function(result) {
                    if (result.success) {
                        showStatus(result.message);
                        loadChemDashboard();
                    } else {
                        showStatus(result.message, true);
                    }
                    document.getElementById('contextMenu').classList.add('hidden');
                })
                .withFailureHandler(function(error) {
                    showStatus('Error: ' + error.message, true);
                    document.getElementById('contextMenu').classList.add('hidden');
                })
                .addToBlacklist(contextMenuPlayer.name, chemStyle, contextMenuPlayer.priceHunter, contextMenuPlayer.priceShadow);
        }

        document.addEventListener('click', function() {
            document.getElementById('contextMenu').classList.add('hidden');
        });

        function filterTable() {
            const input = document.getElementById('searchBox');
            const filter = input.value.toUpperCase();
            const table = document.getElementById('dashboardTable');
            const tr = table.getElementsByTagName('tr');
            
            for (let i = 1; i < tr.length; i++) {
                let found = false;
                const td = tr[i].getElementsByTagName('td');
                for (let j = 0; j < td.length; j++) {
                    if (td[j]) {
                        const txtValue = td[j].textContent || td[j].innerText;
                        if (txtValue.toUpperCase().indexOf(filter) > -1) {
                            found = true;
                            break;
                        }
                    }
                }
                tr[i].style.display = found ? '' : 'none';
            }
        }

        function filterChemTable() {
            const input = document.getElementById('chemSearchBox');
            const filter = input.value.toUpperCase();
            const table = document.getElementById('dashboardTable');
            const tr = table.getElementsByTagName('tr');
            
            for (let i = 1; i < tr.length; i++) {
                let found = false;
                const td = tr[i].getElementsByTagName('td');
                for (let j = 0; j < td.length; j++) {
                    if (td[j]) {
                        const txtValue = td[j].textContent || td[j].innerText;
                        if (txtValue.toUpperCase().indexOf(filter) > -1) {
                            found = true;
                            break;
                        }
                    }
                }
                tr[i].style.display = found ? '' : 'none';
            }
        }

        function buildDashboard(mode) {
            showStatus('Building dashboard in ' + mode + ' mode...');
            const functionMap = {
                'normal': 'buildDashboardNormal',
                'crash': 'buildDashboardCrash',
                'rise': 'buildDashboardRise',
                'investments': 'buildDashboardInvestments'
            };
            google.script.run
                .withSuccessHandler(function(result) {
                    if (result.success) {
                        showStatus(result.message);
                        loadDashboard();
                    } else {
                        showStatus(result.message, true);
                    }
                })
                .withFailureHandler(function(error) {
                    showStatus('Error: ' + error.message, true);
                })
                [functionMap[mode]]();
        }

        function logData() {
            showStatus('Logging manual data to archive...');
            google.script.run
                .withSuccessHandler(function(result) {
                    if (result.success) {
                        showStatus(result.message);
                    } else {
                        showStatus(result.message, true);
                    }
                })
                .withFailureHandler(function(error) {
                    showStatus('Error: ' + error.message, true);
                })
                .logManualData();
        }

        function buildChemDashboard() {
            showStatus('Building chem styles dashboard...');
            google.script.run
                .withSuccessHandler(function(result) {
                    if (result.success) {
                        showStatus(result.message);
                        loadChemDashboard();
                    } else {
                        showStatus(result.message, true);
                    }
                })
                .withFailureHandler(function(error) {
                    showStatus('Error: ' + error.message, true);
                })
                .buildChemStylesDashboard();
        }

        function logChemData() {
            showStatus('Logging chem styles data to archive...');
            google.script.run
                .withSuccessHandler(function(result) {
                    if (result.success) {
                        showStatus(result.message);
                    } else {
                        showStatus(result.message, true);
                    }
                })
                .withFailureHandler(function(error) {
                    showStatus('Error: ' + error.message, true);
                })
                .logChemStylesData();
        }

        window.onload = loadDashboard;
    </script>
</body>
</html>`;
}

// ========================================
// CUSTOM MENU
// ========================================

function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('FUT Trading Console Actions')
    .addSubMenu(ui.createMenu('📊 Fluctuations')
      .addItem('Build Dashboard (Normal Mode)', 'buildDashboardNormal')
      .addItem('Build Dashboard (Crash Mode)', 'buildDashboardCrash')
      .addItem('Build Dashboard (Rise Mode)', 'buildDashboardRise')
      .addItem('Build Dashboard (Investments Mode)', 'buildDashboardInvestments')
      .addItem('Log Manual Data to Archive', 'menuLogManualData'))
    .addSubMenu(ui.createMenu('⚗️ Chem Styles')
      .addItem('Build Chem Styles Dashboard', 'menuBuildChemDashboard')
      .addItem('Log Chem Styles to Archive', 'menuLogChemData'))
    .addToUi();
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

function menuLogChemData() {
  const result = logChemStylesData();
  const ui = SpreadsheetApp.getUi();
  if (result.success) {
    ui.alert('Success', result.message, ui.ButtonSet.OK);
  } else {
    ui.alert('Error', result.message, ui.ButtonSet.OK);
  }
}
