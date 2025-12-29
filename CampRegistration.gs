/**
 * Camp Registration Tracking System - Apps Script
 * Processes roster data from Paste Report tab and updates Master Data
 * Handles multi-year data with intelligent deduplication and cancellation tracking
 *
 * FIXED VERSION - Properly handles data replacement and cancellation tracking
 */

// Configuration
const CONFIG = {
  PASTE_SHEET: "Paste Report",
  MASTER_SHEET: "Master Data",
  CANCELLATION_SHEET: "Cancellations",
  HEADER_ROW: 1,
  DATA_START_ROW: 2
};

// Required columns from CSV (these must be present)
const REQUIRED_COLUMNS = [
  'Member ID',
  'Program',
  'Site',
  'Instance Name',
  'Program Start',
  'First Name',
  'Last Name'
];

// Calculated columns that will be added
const CALCULATED_COLUMNS = [
  'Year',
  'WeekNumber',
  'SiteDisplay',
  'DateProcessed',
  'LastUpdated'
];

/**
 * Creates custom menu when spreadsheet opens
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('📊 Camp Registration')
    .addItem('▶️ Process Updates', 'processRosterUpdate')
    .addSeparator()
    .addItem('🔄 Clear Paste Report', 'clearPasteReport')
    .addItem('ℹ️ Show Instructions', 'showInstructions')
    .addToUi();
}

/**
 * Main function to process roster updates
 */
function processRosterUpdate() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();

  try {
    // Step 1: Get and validate paste report data
    const pasteSheet = ss.getSheetByName(CONFIG.PASTE_SHEET);
    if (!pasteSheet) {
      throw new Error("Paste Report sheet not found");
    }

    const pasteData = pasteSheet.getDataRange().getValues();
    if (pasteData.length <= 1) {
      ui.alert('⚠️ No Data', 'Please paste roster data before processing.', ui.ButtonSet.OK);
      return;
    }

    // Step 2: Validate required columns
    const pasteHeaders = pasteData[0];
    const missingColumns = REQUIRED_COLUMNS.filter(col => !pasteHeaders.includes(col));
    if (missingColumns.length > 0) {
      throw new Error(`Missing required columns: ${missingColumns.join(', ')}`);
    }

    // Step 3: Extract metadata from first data row
    const firstRow = pasteData[1];
    const metadata = extractMetadata(firstRow, pasteHeaders);

    // Step 4: Confirm with user
    const response = ui.alert(
      '🔍 Detected Data',
      `Program: ${metadata.program}\nYear: ${metadata.year}\nRows: ${pasteData.length - 1}\n\nProceed with update?`,
      ui.ButtonSet.YES_NO
    );

    if (response !== ui.Button.YES) {
      return;
    }

    // Step 5: Process and update
    const result = processAndUpdate(ss, pasteData, pasteHeaders, metadata);

    // Step 6: Show results
    const cancelMsg = result.cancellations > 0
      ? `\n⚠️ ${result.cancellations} cancellations detected (see Cancellations tab)`
      : '';

    ui.alert(
      '✅ Update Complete',
      `Program: ${metadata.program}\nYear: ${metadata.year}\n\n` +
      `✨ New registrations: ${result.newCount}\n` +
      `🔄 Updated existing: ${result.updatedCount}\n` +
      `📋 Total active: ${result.totalCount}${cancelMsg}`,
      ui.ButtonSet.OK
    );

  } catch (error) {
    ui.alert('❌ Error', error.toString(), ui.ButtonSet.OK);
    Logger.log('Error in processRosterUpdate: ' + error);
    Logger.log(error.stack);
  }
}

/**
 * Extract metadata (year, program) from data row
 */
function extractMetadata(row, headers) {
  const programCol = headers.indexOf('Program');
  const programStartCol = headers.indexOf('Program Start');

  const programValue = row[programCol];
  const programStartValue = row[programStartCol];

  return {
    year: extractYear(programStartValue),
    program: identifyProgram(programValue)
  };
}

/**
 * Extract year from Program Start date
 */
function extractYear(dateValue) {
  if (dateValue instanceof Date) {
    return dateValue.getFullYear();
  }
  const date = new Date(dateValue);
  if (!isNaN(date.getTime())) {
    return date.getFullYear();
  }
  return new Date().getFullYear();
}

/**
 * Identify program type from Program column value
 */
function identifyProgram(programString) {
  const programLower = String(programString).toLowerCase();

  if (programLower.includes('camp winnebago')) {
    return 'Camp Winnebago';
  } else if (programLower.includes('adventure camps (say/neb)')) {
    return 'Adventure Camps (SAY/NEB)';
  } else if (programLower.includes('adventure camp (good shepherd)')) {
    return 'Adventure Camp (Good Shepherd)';
  }

  return programString;
}

/**
 * Main processing function - handles all data transformation and updates
 */
function processAndUpdate(ss, pasteData, pasteHeaders, metadata) {
  const now = new Date();

  // Build column mapping for paste data
  const pasteMap = buildColumnMap(pasteHeaders);

  // Process new data into normalized format
  const newRecords = [];
  for (let i = 1; i < pasteData.length; i++) {
    const row = pasteData[i];

    // Skip empty rows
    if (!row[pasteMap['Member ID']]) continue;

    const record = buildRecord(row, pasteMap, metadata, now);
    newRecords.push(record);
  }

  // Get or create master sheet
  let masterSheet = ss.getSheetByName(CONFIG.MASTER_SHEET);
  const isNewSheet = !masterSheet;

  if (isNewSheet) {
    masterSheet = ss.insertSheet(CONFIG.MASTER_SHEET);
  }

  // Get existing master data
  let existingRecords = [];
  if (!isNewSheet && masterSheet.getLastRow() > 0) {
    const masterData = masterSheet.getDataRange().getValues();
    const masterHeaders = masterData[0];
    const masterMap = buildColumnMap(masterHeaders);

    // Parse existing records
    for (let i = 1; i < masterData.length; i++) {
      const record = parseRecord(masterData[i], masterMap);
      existingRecords.push(record);
    }
  }

  // Perform intelligent merge
  const mergeResult = mergeRecords(existingRecords, newRecords, metadata, now);

  // Log cancellations if any
  if (mergeResult.cancellations.length > 0) {
    logCancellations(ss, mergeResult.cancellations);
  }

  // Write merged data back to master sheet
  writeMasterData(masterSheet, mergeResult.records);

  return {
    newCount: mergeResult.newCount,
    updatedCount: mergeResult.updatedCount,
    totalCount: mergeResult.records.length,
    cancellations: mergeResult.cancellations.length
  };
}

/**
 * Build column name to index mapping
 */
function buildColumnMap(headers) {
  const map = {};
  headers.forEach((header, index) => {
    map[header] = index;
  });
  return map;
}

/**
 * Build a normalized record object from row data
 */
function buildRecord(row, columnMap, metadata, timestamp) {
  const record = {
    // Core fields
    memberId: String(row[columnMap['Member ID']] || '').trim(),
    program: metadata.program,
    site: row[columnMap['Site']] || '',
    instanceName: row[columnMap['Instance Name']] || '',
    year: metadata.year,

    // Calculated fields
    weekNumber: extractWeekNumber(row[columnMap['Instance Name']]),
    siteDisplay: cleanSiteName(row[columnMap['Site']]),

    // Timestamps
    dateProcessed: timestamp,
    lastUpdated: timestamp,

    // All original data (for preservation)
    originalData: {}
  };

  // Store all original columns
  for (const [colName, colIndex] of Object.entries(columnMap)) {
    record.originalData[colName] = row[colIndex];
  }

  return record;
}

/**
 * Parse existing record from master data row
 */
function parseRecord(row, columnMap) {
  const record = {
    memberId: String(row[columnMap['Member ID']] || '').trim(),
    program: row[columnMap['Program']] || '',
    site: row[columnMap['Site']] || '',
    instanceName: row[columnMap['Instance Name']] || '',
    year: parseInt(row[columnMap['Year']]) || 0,
    weekNumber: row[columnMap['WeekNumber']] || '',
    siteDisplay: row[columnMap['SiteDisplay']] || '',
    dateProcessed: row[columnMap['DateProcessed']] || new Date(),
    lastUpdated: row[columnMap['LastUpdated']] || new Date(),
    originalData: {}
  };

  // Store all original columns
  for (const [colName, colIndex] of Object.entries(columnMap)) {
    record.originalData[colName] = row[colIndex];
  }

  return record;
}

/**
 * Build composite key for record identification
 */
function buildKey(record) {
  return `${record.memberId}|${record.program}|${record.site}|${record.instanceName}|${record.year}`;
}

/**
 * Merge existing and new records with intelligent deduplication
 * Returns: { records: [], cancellations: [], newCount: 0, updatedCount: 0 }
 */
function mergeRecords(existingRecords, newRecords, metadata, now) {
  const result = {
    records: [],
    cancellations: [],
    newCount: 0,
    updatedCount: 0
  };

  // Build map of new records by composite key
  const newRecordMap = new Map();
  newRecords.forEach(record => {
    const key = buildKey(record);
    newRecordMap.set(key, record);
  });

  // Track which new records we've seen
  const processedNewKeys = new Set();

  // Process existing records
  existingRecords.forEach(existingRecord => {
    const shouldUpdate = (
      existingRecord.program === metadata.program &&
      existingRecord.year === metadata.year
    );

    if (!shouldUpdate) {
      // Keep records from other programs/years unchanged
      result.records.push(existingRecord);
      return;
    }

    // This record is from the program/year being updated
    const key = buildKey(existingRecord);

    if (newRecordMap.has(key)) {
      // Record still exists - update it
      const newRecord = newRecordMap.get(key);
      newRecord.dateProcessed = existingRecord.dateProcessed; // Preserve original date
      newRecord.lastUpdated = now; // Update timestamp
      result.records.push(newRecord);
      processedNewKeys.add(key);
      result.updatedCount++;
    } else {
      // Record no longer exists - it's a cancellation
      result.cancellations.push({
        memberId: existingRecord.memberId,
        firstName: existingRecord.originalData['First Name'] || '',
        lastName: existingRecord.originalData['Last Name'] || '',
        program: existingRecord.program,
        site: existingRecord.siteDisplay || existingRecord.site,
        weekNumber: existingRecord.weekNumber,
        year: existingRecord.year
      });
    }
  });

  // Add new records that didn't exist before
  newRecords.forEach(newRecord => {
    const key = buildKey(newRecord);
    if (!processedNewKeys.has(key)) {
      result.records.push(newRecord);
      result.newCount++;
    }
  });

  return result;
}

/**
 * Write merged records to master sheet
 */
function writeMasterData(sheet, records) {
  // Clear existing content
  sheet.clear();

  if (records.length === 0) {
    return;
  }

  // Get all unique column names from all records
  const allColumns = new Set();
  records.forEach(record => {
    Object.keys(record.originalData).forEach(col => allColumns.add(col));
  });

  // Build final header list: original columns + calculated columns
  const originalColumns = Array.from(allColumns);
  const headers = [...originalColumns, ...CALCULATED_COLUMNS];

  // Build rows
  const rows = [headers];
  records.forEach(record => {
    const row = [];

    // Add original data columns
    originalColumns.forEach(colName => {
      row.push(record.originalData[colName] || '');
    });

    // Add calculated columns
    row.push(record.year);
    row.push(record.weekNumber);
    row.push(record.siteDisplay);
    row.push(record.dateProcessed);
    row.push(record.lastUpdated);

    rows.push(row);
  });

  // Write to sheet
  sheet.getRange(1, 1, rows.length, headers.length).setValues(rows);

  // Format header
  sheet.getRange(1, 1, 1, headers.length)
    .setFontWeight('bold')
    .setBackground('#4285f4')
    .setFontColor('white');

  sheet.setFrozenRows(1);

  // Auto-resize columns
  for (let i = 1; i <= headers.length; i++) {
    sheet.autoResizeColumn(i);
  }
}

/**
 * Extract week number from Instance Name
 */
function extractWeekNumber(instanceName) {
  if (!instanceName) return '';

  const str = String(instanceName);

  // Check for "Pre-Camp"
  if (str.toLowerCase().includes('pre-camp')) {
    return 'Pre-Camp';
  }

  // Extract "Week #"
  const match = str.match(/Week\s*\d+/i);
  return match ? match[0] : str;
}

/**
 * Clean site name for display
 */
function cleanSiteName(siteName) {
  if (!siteName) return '';

  const str = String(siteName);

  const siteMap = {
    '(KC) Kinder Camp': 'KinderCamp',
    '(AC) Animal Camp': 'Animal Camp',
    '(SB) Standing Bear Camp': 'Standing Bear',
    '(WC) Wilderness Camp': 'Wilderness',
    '(SE) Soaring Eagle Camp': 'Soaring Eagle',
    '(ADVCN) Adventure Camp - Northeast Family YMCA': 'NEB',
    '(ADVCS) Adventure Camp - SwedishAmerican YMCA': 'SAY',
    '(GSY) Good Shepherd YMCA': 'Good Shep'
  };

  // Exact match
  if (siteMap[str]) {
    return siteMap[str];
  }

  // Partial match
  for (const [key, value] of Object.entries(siteMap)) {
    if (str.includes(key)) {
      return value;
    }
  }

  return str;
}

/**
 * Log cancellations to dedicated sheet
 */
function logCancellations(ss, cancellations) {
  let cancelSheet = ss.getSheetByName(CONFIG.CANCELLATION_SHEET);

  // Create sheet if it doesn't exist
  if (!cancelSheet) {
    cancelSheet = ss.insertSheet(CONFIG.CANCELLATION_SHEET);
    const headers = [
      'Date Detected',
      'Member ID',
      'First Name',
      'Last Name',
      'Program',
      'Site',
      'Week',
      'Year'
    ];

    cancelSheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    cancelSheet.getRange(1, 1, 1, headers.length)
      .setFontWeight('bold')
      .setBackground('#dc3545')
      .setFontColor('white');

    cancelSheet.setFrozenRows(1);
  }

  // Append cancellations
  const now = new Date();
  const rows = cancellations.map(c => [
    now,
    c.memberId,
    c.firstName,
    c.lastName,
    c.program,
    c.site,
    c.weekNumber,
    c.year
  ]);

  if (rows.length > 0) {
    const lastRow = cancelSheet.getLastRow();
    const startRow = lastRow + 1;
    cancelSheet.getRange(startRow, 1, rows.length, 8).setValues(rows);
  }
}

/**
 * Clear Paste Report sheet
 */
function clearPasteReport() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const pasteSheet = ss.getSheetByName(CONFIG.PASTE_SHEET);
  const ui = SpreadsheetApp.getUi();

  if (!pasteSheet) {
    ui.alert('❌ Error', 'Paste Report sheet not found.', ui.ButtonSet.OK);
    return;
  }

  const response = ui.alert(
    '🔄 Clear Paste Report',
    'This will clear all data from the Paste Report tab. Continue?',
    ui.ButtonSet.YES_NO
  );

  if (response === ui.Button.YES) {
    pasteSheet.clear();
    ui.alert('✅ Cleared', 'Paste Report has been cleared.', ui.ButtonSet.OK);
  }
}

/**
 * Show usage instructions
 */
function showInstructions() {
  const ui = SpreadsheetApp.getUi();
  const message = `
📋 HOW TO USE THIS TRACKER

1️⃣ PASTE ROSTER DATA
   • Download CSV from Daxko
   • Go to "Paste Report" tab
   • Copy ALL data from CSV (Ctrl+A in Excel/Sheets)
   • Paste into cell A1 (important!)

2️⃣ PROCESS UPDATES
   • Menu: Camp Registration > Process Updates
   • Confirm detected program and year
   • System automatically:
     ✓ Updates existing registrations
     ✓ Adds new registrations
     ✓ Detects cancellations

3️⃣ VIEW RESULTS
   • Master Data: All current registrations
   • Cancellations: Dropped registrations
   • All analysis tabs update automatically

⚠️ IMPORTANT
   • Always paste FULL roster (not partial)
   • Start at cell A1
   • Include header row
   • Works with all 3 programs

✨ FEATURES
   • Smart deduplication (no duplicates!)
   • Multi-year tracking
   • Automatic cancellation detection
   • Preserves historical data

❓ Questions? Contact your administrator
`;

  ui.alert('📊 Camp Registration Tracker', message, ui.ButtonSet.OK);
}
