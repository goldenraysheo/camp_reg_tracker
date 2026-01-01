/**
 * Camp Registration Tracking System - Apps Script
 * Simple, working version that actually does what it's supposed to do
 *
 * Paste roster → Detect program/year → Replace that program/year's data → Track cancellations
 */

// Sheet names
const PASTE_SHEET = "Paste Report";
const MASTER_SHEET = "Master Data";
const CANCEL_SHEET = "Cancellations";
const REG_TRENDS_SHEET = "Reg Trends";

// Column indices in CSV (0-based)
const COL = {
  MEMBER_ID: 1,      // Member ID
  FIRST_NAME: 2,     // First Name
  LAST_NAME: 3,      // Last Name
  PROGRAM: 25,       // Program
  PROGRAM_START: 27, // Program Start
  SITE: 24,          // Site
  INSTANCE: 30       // Instance Name
};

/**
 * Creates custom menu when spreadsheet opens
 */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('📊 Camp Registration')
    .addItem('▶️ Process Updates', 'processRosterUpdate')
    .addToUi();
}

/**
 * Main processing function
 */
function processRosterUpdate() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();

  try {
    // Step 1: Get paste data
    const pasteSheet = ss.getSheetByName(PASTE_SHEET);
    if (!pasteSheet) {
      throw new Error("Paste Report sheet not found");
    }

    const lastRow = pasteSheet.getLastRow();
    const lastCol = pasteSheet.getLastColumn();

    if (lastRow < 2) {
      ui.alert('⚠️ No Data', 'Please paste roster data into Paste Report tab starting at A1.', ui.ButtonSet.OK);
      return;
    }

    // Get all data including headers
    const pasteData = pasteSheet.getRange(1, 1, lastRow, lastCol).getValues();
    const headers = pasteData[0];

    // Validate we have required columns
    if (headers[COL.MEMBER_ID] !== 'Member ID' || headers[COL.PROGRAM] !== 'Program') {
      throw new Error('Invalid data format. Make sure you paste with headers starting at A1.');
    }

    // Step 2: Detect program and year from first data row
    const firstDataRow = pasteData[1];
    const program = normalizeProgram(firstDataRow[COL.PROGRAM]);
    const year = extractYear(firstDataRow[COL.PROGRAM_START]);
    const rowCount = pasteData.length - 1; // excluding header

    // Step 3: Confirm with user
    const response = ui.alert(
      '🔍 Ready to Process',
      `Program: ${program}\nYear: ${year}\nRows: ${rowCount}\n\nThis will REPLACE all existing ${program} ${year} data.\n\nContinue?`,
      ui.ButtonSet.YES_NO
    );

    if (response !== ui.Button.YES) {
      return;
    }

    // Show progress indicator
    ss.toast('Processing roster data...', '⏳ Working', -1);

    // Step 4: Process the data
    const processedData = processData(pasteData, headers, program, year);

    // Step 5: Merge with existing Master Data
    ss.toast('Updating Master Data...', '⏳ Working', -1);
    const result = mergeWithMaster(ss, processedData, program, year);

    // Step 6: Log cancellations if any
    if (result.cancellations.length > 0) {
      ss.toast('Logging cancellations...', '⏳ Working', -1);
      logCancellations(ss, result.cancellations);
    }

    // Step 6.5: Organize tabs and navigate to Overview Dashboard
    ss.toast('Finalizing...', '⏳ Working', -1);
    organizeTabs(ss);

    // Step 6.75: Record registration snapshot
    ss.toast('Recording snapshot...', '⏳ Working', -1);
    recordRegistrationSnapshot();

    // Step 7: Clear Paste Report
    pasteSheet.clear();

    // Step 8: Success message
    ss.toast('Complete! 🎉', '✅ Success', 3);

    const cancelMsg = result.cancellations.length > 0
      ? `\n⚠️ Cancellations: ${result.cancellations.length} (see Cancellations tab)`
      : '';

    ui.alert(
      '✅ Update Complete',
      `Program: ${program}\nYear: ${year}\n\n` +
      `New registrations: ${result.newCount}\n` +
      `Updated existing: ${result.updatedCount}\n` +
      `Total in Master Data: ${result.totalCount}${cancelMsg}`,
      ui.ButtonSet.OK
    );

  } catch (error) {
    ui.alert('❌ Error', error.toString(), ui.ButtonSet.OK);
    Logger.log('Error: ' + error.toString());
    Logger.log(error.stack);
  }
}

/**
 * Process paste data: add calculated columns to each row
 * Returns: { headers: [...], rows: [[...], [...], ...] }
 */
function processData(pasteData, csvHeaders, program, year) {
  const now = new Date();

  // Build new headers: CSV headers + calculated columns
  const newHeaders = [
    ...csvHeaders,
    'Year',
    'WeekNumber',
    'SiteDisplay',
    'DateProcessed',
    'LastUpdated'
  ];

  // Process each data row
  const processedRows = [];

  for (let i = 1; i < pasteData.length; i++) {
    const csvRow = pasteData[i];

    // Skip empty rows
    if (!csvRow[COL.MEMBER_ID]) continue;

    // Add calculated columns
    const newRow = [
      ...csvRow,
      year,                                    // Year
      extractWeekNumber(csvRow[COL.INSTANCE]), // WeekNumber
      cleanSiteName(csvRow[COL.SITE]),        // SiteDisplay
      now,                                     // DateProcessed
      now                                      // LastUpdated
    ];

    processedRows.push(newRow);
  }

  return {
    headers: newHeaders,
    rows: processedRows
  };
}

/**
 * Merge new data with existing Master Data
 * Returns: { masterRows: [[...]], cancellations: [...], newCount: 0, updatedCount: 0, totalCount: 0 }
 */
function mergeWithMaster(ss, processedData, program, year) {
  let masterSheet = ss.getSheetByName(MASTER_SHEET);

  // If Master Data doesn't exist, create it and write all data
  if (!masterSheet) {
    masterSheet = ss.insertSheet(MASTER_SHEET);
    const allRows = [processedData.headers, ...processedData.rows];
    writeMasterSheet(masterSheet, allRows);

    return {
      masterRows: allRows,
      cancellations: [],
      newCount: processedData.rows.length,
      updatedCount: 0,
      totalCount: processedData.rows.length
    };
  }

  // Master Data exists - need to merge
  const existingData = masterSheet.getDataRange().getValues();
  const existingHeaders = existingData[0];

  // Find column indices in existing Master Data
  const existingCols = {
    memberId: existingHeaders.indexOf('Member ID'),
    program: existingHeaders.indexOf('Program'),
    site: existingHeaders.indexOf('Site'),
    instance: existingHeaders.indexOf('Instance Name'),
    year: existingHeaders.indexOf('Year'),
    firstName: existingHeaders.indexOf('First Name'),
    lastName: existingHeaders.indexOf('Last Name'),
    dateProcessed: existingHeaders.indexOf('DateProcessed')
  };

  // Find column indices in new data
  const newCols = {
    memberId: processedData.headers.indexOf('Member ID'),
    program: processedData.headers.indexOf('Program'),
    site: processedData.headers.indexOf('Site'),
    instance: processedData.headers.indexOf('Instance Name'),
    year: processedData.headers.indexOf('Year'),
    dateProcessed: processedData.headers.indexOf('DateProcessed'),
    lastUpdated: processedData.headers.indexOf('LastUpdated')
  };

  // Build composite key for new data
  const newDataMap = new Map();
  processedData.rows.forEach(row => {
    const key = buildKey(
      row[newCols.memberId],
      row[newCols.program],
      row[newCols.site],
      row[newCols.instance],
      row[newCols.year]
    );
    newDataMap.set(key, row);
  });

  // Process existing data
  const rowsToKeep = [];
  const existingKeys = new Set();
  const cancellations = [];
  let updatedCount = 0;

  for (let i = 1; i < existingData.length; i++) {
    const row = existingData[i];
    const rowProgram = normalizeProgram(row[existingCols.program]);
    const rowYear = parseInt(row[existingCols.year]);

    // Keep rows from other programs/years unchanged
    if (rowProgram !== program || rowYear !== year) {
      rowsToKeep.push(row);
      continue;
    }

    // This row is from the program/year being updated
    const key = buildKey(
      row[existingCols.memberId],
      row[existingCols.program],
      row[existingCols.site],
      row[existingCols.instance],
      row[existingCols.year]
    );

    if (newDataMap.has(key)) {
      // This registration still exists - update it
      const newRow = newDataMap.get(key);
      // Preserve original DateProcessed
      newRow[newCols.dateProcessed] = row[existingCols.dateProcessed];
      // Update LastUpdated
      newRow[newCols.lastUpdated] = new Date();
      rowsToKeep.push(newRow);
      existingKeys.add(key);
      updatedCount++;
    } else {
      // This registration is gone - it's a cancellation
      cancellations.push({
        memberId: row[existingCols.memberId],
        firstName: row[existingCols.firstName] || '',
        lastName: row[existingCols.lastName] || '',
        program: rowProgram,
        site: row[existingCols.site],
        instance: row[existingCols.instance],
        year: rowYear
      });
    }
  }

  // Add brand new registrations
  let newCount = 0;
  processedData.rows.forEach(row => {
    const key = buildKey(
      row[newCols.memberId],
      row[newCols.program],
      row[newCols.site],
      row[newCols.instance],
      row[newCols.year]
    );
    if (!existingKeys.has(key)) {
      rowsToKeep.push(row);
      newCount++;
    }
  });

  // Prepare final data with headers
  const allRows = [processedData.headers, ...rowsToKeep];

  // Write to Master Data
  writeMasterSheet(masterSheet, allRows);

  return {
    masterRows: allRows,
    cancellations: cancellations,
    newCount: newCount,
    updatedCount: updatedCount,
    totalCount: rowsToKeep.length
  };
}

/**
 * Write data to Master Data sheet
 */
function writeMasterSheet(sheet, rows) {
  // Clear existing content
  sheet.clear();

  if (rows.length === 0) return;

  // Write all data
  sheet.getRange(1, 1, rows.length, rows[0].length).setValues(rows);

  // Format header row
  sheet.getRange(1, 1, 1, rows[0].length)
    .setFontWeight('bold')
    .setBackground('#4285f4')
    .setFontColor('white');

  // Freeze header row
  sheet.setFrozenRows(1);
}

/**
 * Build composite key for deduplication
 */
function buildKey(memberId, program, site, instance, year) {
  return `${memberId}|${normalizeProgram(program)}|${site}|${instance}|${year}`;
}

/**
 * Normalize program name for consistent comparison
 */
function normalizeProgram(programName) {
  if (!programName) return '';

  const str = String(programName).toLowerCase();

  if (str.includes('camp winnebago')) {
    return 'Camp Winnebago';
  } else if (str.includes('adventure camps (say/neb)')) {
    return 'Adventure Camps (SAY/NEB)';
  } else if (str.includes('adventure camp (good shepherd)')) {
    return 'Adventure Camp (Good Shepherd)';
  }

  // Return as-is if no match
  return programName;
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
 * Extract week number from Instance Name
 * Example: "Week 8: 7/28 - 8/1" → "Week 8"
 */
function extractWeekNumber(instanceName) {
  if (!instanceName) return '';

  const str = String(instanceName);

  // Check for Pre-Camp
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

  // Try exact match first
  if (siteMap[str]) {
    return siteMap[str];
  }

  // Try partial match
  for (const [key, value] of Object.entries(siteMap)) {
    if (str.includes(key)) {
      return value;
    }
  }

  return str;
}

/**
 * Log cancellations to Cancellations tab
 */
function logCancellations(ss, cancellations) {
  let cancelSheet = ss.getSheetByName(CANCEL_SHEET);

  // Create sheet if it doesn't exist
  if (!cancelSheet) {
    cancelSheet = ss.insertSheet(CANCEL_SHEET);
    const headers = [
      'Date Detected',
      'Member ID',
      'First Name',
      'Last Name',
      'Program',
      'Site',
      'Instance',
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
    c.instance,
    c.year
  ]);

  if (rows.length > 0) {
    const lastRow = cancelSheet.getLastRow();
    cancelSheet.getRange(lastRow + 1, 1, rows.length, 8).setValues(rows);
  }
}

/**
 * Clear Paste Report sheet
 */
function clearPasteReport() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();
  const pasteSheet = ss.getSheetByName(PASTE_SHEET);

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
 * Organize tabs in preferred order and navigate to Overview Dashboard
 * Order: Paste Report, Overview Dashboard, Cancellations, Master Data
 */
function organizeTabs(ss) {
  if (!ss) {
    ss = SpreadsheetApp.getActiveSpreadsheet();
  }

  const desiredOrder = [
    PASTE_SHEET,
    'Overview Dashboard',
    CANCEL_SHEET,
    MASTER_SHEET
  ];

  // Move sheets to desired positions (only if they exist)
  let position = 1; // Start at position 1 (leftmost)
  desiredOrder.forEach(sheetName => {
    const sheet = ss.getSheetByName(sheetName);
    if (sheet) {
      // Only move if sheet exists
      ss.setActiveSheet(sheet);
      ss.moveActiveSheet(position);
      position++; // Increment for next sheet
    }
  });

  // Navigate to Overview Dashboard
  const overviewSheet = ss.getSheetByName('Overview Dashboard');
  if (overviewSheet) {
    ss.setActiveSheet(overviewSheet);
  }
}

/**
 * Record a snapshot of current registration counts
 * Creates a time-series record of registrations by program
 * This captures growth AND cancellations over time
 */
function recordRegistrationSnapshot() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const masterSheet = ss.getSheetByName(MASTER_SHEET);

  if (!masterSheet) {
    // No Master Data yet, nothing to record
    return;
  }

  // Get or create Reg Trends sheet
  let regSheet = ss.getSheetByName(REG_TRENDS_SHEET);
  const isNewSheet = !regSheet;

  if (!regSheet) {
    regSheet = ss.insertSheet(REG_TRENDS_SHEET);
    regSheet.setFrozenRows(1);
    // Hide the sheet (users can unhide if needed)
    regSheet.hideSheet();
  }

  // Read Master Data
  const masterData = masterSheet.getDataRange().getValues();
  const masterHeaders = masterData[0];

  // Find column indices
  const programCol = masterHeaders.indexOf('Program');
  const siteDisplayCol = masterHeaders.indexOf('SiteDisplay');
  const weekNumberCol = masterHeaders.indexOf('WeekNumber');
  const regDateCol = masterHeaders.indexOf('Registration Date');

  if (programCol === -1 || siteDisplayCol === -1 || weekNumberCol === -1) {
    throw new Error('Required columns not found in Master Data');
  }

  // TABLE 1: Process Date Snapshot (columns A-E)
  // Only initialize headers if this is a new sheet
  if (isNewSheet) {
    const headers1 = ['Updated', 'Camp', 'SAY', 'NEB', 'GS'];
    regSheet.getRange(1, 1, 1, headers1.length).setValues([headers1]);
    regSheet.getRange(1, 1, 1, headers1.length)
      .setFontWeight('bold')
      .setBackground('#34a853')
      .setFontColor('white');
  }

  // Count registrations by program (excluding Pre-Camp)
  let campCount = 0;
  let sayCount = 0;
  let nebCount = 0;
  let gsCount = 0;

  for (let i = 1; i < masterData.length; i++) {
    const row = masterData[i];
    const program = normalizeProgram(row[programCol]);
    const site = row[siteDisplayCol];
    const weekNumber = row[weekNumberCol];

    // Skip Pre-Camp weeks
    if (weekNumber && String(weekNumber).toLowerCase().includes('pre-camp')) {
      continue;
    }

    if (program === 'Camp Winnebago') {
      campCount++;
    } else if (program === 'Adventure Camps (SAY/NEB)') {
      if (site === 'SAY') {
        sayCount++;
      } else if (site === 'NEB') {
        nebCount++;
      }
    } else if (program === 'Adventure Camp (Good Shepherd)') {
      gsCount++;
    }
  }

  // Get today's date (formatted as date only, no time)
  const today = new Date();
  today.setHours(0, 0, 0, 0);

  // Check if we already have an entry for today in Table 1
  // We need to check only column A (not the entire sheet)
  const columnAValues = regSheet.getRange(2, 1, Math.max(1, regSheet.getLastRow() - 1), 1).getValues();
  let todayRowIndex = -1;

  for (let i = 0; i < columnAValues.length; i++) {
    if (columnAValues[i][0]) {
      const rowDate = new Date(columnAValues[i][0]);
      rowDate.setHours(0, 0, 0, 0);

      if (rowDate.getTime() === today.getTime()) {
        todayRowIndex = i + 2; // Convert to 1-based row number (i is 0-based, +2 because we started from row 2)
        break;
      }
    }
  }

  const newRow = [today, campCount, sayCount, nebCount, gsCount];

  if (todayRowIndex > 0) {
    // Update existing row for today
    regSheet.getRange(todayRowIndex, 1, 1, newRow.length).setValues([newRow]);
  } else {
    // Append new row - find last row with data in column A
    let lastRowInTableOne = 1; // Start at header
    for (let i = 0; i < columnAValues.length; i++) {
      if (columnAValues[i][0]) {
        lastRowInTableOne = i + 2; // Convert to 1-based row number
      }
    }
    regSheet.getRange(lastRowInTableOne + 1, 1, 1, newRow.length).setValues([newRow]);
  }

  // TABLE 2: Registration Date Table (columns H-L, 2 columns away)
  // Initialize headers if new sheet
  if (isNewSheet) {
    const headers2 = ['Reg Count', 'Camp', 'SAY', 'NEB', 'GS'];
    regSheet.getRange(1, 8, 1, headers2.length).setValues([headers2]);
    regSheet.getRange(1, 8, 1, headers2.length)
      .setFontWeight('bold')
      .setBackground('#34a853')
      .setFontColor('white');
  }

  // Group registrations by registration date
  const regDateMap = new Map();

  for (let i = 1; i < masterData.length; i++) {
    const row = masterData[i];
    const program = normalizeProgram(row[programCol]);
    const site = row[siteDisplayCol];
    const weekNumber = row[weekNumberCol];
    const regDate = row[regDateCol];

    // Skip Pre-Camp weeks
    if (weekNumber && String(weekNumber).toLowerCase().includes('pre-camp')) {
      continue;
    }

    // Skip if no registration date
    if (!regDate) continue;

    // Normalize the date (remove time component)
    const dateObj = new Date(regDate);
    dateObj.setHours(0, 0, 0, 0);
    const dateKey = dateObj.getTime();

    if (!regDateMap.has(dateKey)) {
      regDateMap.set(dateKey, {
        camp: 0,
        say: 0,
        neb: 0,
        gs: 0
      });
    }

    const counts = regDateMap.get(dateKey);

    if (program === 'Camp Winnebago') {
      counts.camp++;
    } else if (program === 'Adventure Camps (SAY/NEB)') {
      if (site === 'SAY') {
        counts.say++;
      } else if (site === 'NEB') {
        counts.neb++;
      }
    } else if (program === 'Adventure Camp (Good Shepherd)') {
      counts.gs++;
    }
  }

  // Generate ALL dates from 12/1/2025 to 9/1/2026
  const startDate = new Date(2025, 11, 1); // Dec 1, 2025 (month is 0-indexed)
  const endDate = new Date(2026, 8, 1);    // Sep 1, 2026
  const allDates = [];

  for (let d = new Date(startDate); d <= endDate; d.setDate(d.getDate() + 1)) {
    const dateKey = new Date(d).getTime();
    const counts = regDateMap.get(dateKey) || { camp: 0, say: 0, neb: 0, gs: 0 };
    allDates.push([new Date(d), counts.camp, counts.say, counts.neb, counts.gs]);
  }

  // Clear old Table 2 data (keep header, clear everything below)
  const lastRow = regSheet.getLastRow();
  const maxCols = regSheet.getMaxColumns();
  if (lastRow > 1 && maxCols >= 12) {
    // Clear columns H-L from row 2 onwards
    regSheet.getRange(2, 8, lastRow - 1, 5).clearContent();
  }

  // Write new Table 2 data
  if (allDates.length > 0) {
    regSheet.getRange(2, 8, allDates.length, 5).setValues(allDates);
  }
}
