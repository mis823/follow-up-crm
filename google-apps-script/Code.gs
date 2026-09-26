/**
 * Follow-Up CRM — Google Apps Script Webhook Bridge
 * 
 * SETUP INSTRUCTIONS:
 * 1. Open your Google Spreadsheet (with 'Master' and 'Responses1' sheets).
 * 2. In top menu, click: Extensions > Apps Script.
 * 3. Delete any code there, paste this entire file, and click Save (disk icon).
 * 4. In top right, click: Deploy > New deployment.
 * 5. Click the gear icon (Select type) > Choose "Web app".
 * 6. Set Description: "Follow-Up CRM Webhook"
 * 7. Set "Execute as": "Me" (your Google account)
 * 8. Set "Who has access": "Anyone"  <-- IMPORTANT
 * 9. Click "Deploy" (Review permissions and grant access with your account).
 * 10. Copy the Web App URL (ends with /exec) and paste it into your .env file:
 *     APPS_SCRIPT_WEBHOOK_URL="https://script.google.com/macros/s/.../exec"
 */

function doGet(e) {
  try {
    const params = e && e.parameter ? e.parameter : {};
    const action = params.action; // DO NOT default to getMaster!

    // 1. Instant Ping / Health Check (< 0.5s)
    // When visiting the Web App URL in a browser or pinging status
    if (!action || action === 'ping' || action === 'status') {
      return jsonResponse({
        success: true,
        status: 'online',
        message: 'Follow-Up CRM Webhook Bridge is running smoothly.',
        time: Utilities.formatDate(new Date(), 'Asia/Kolkata', 'dd/MM/yyyy HH:mm:ss')
      });
    }

    const ss = SpreadsheetApp.getActiveSpreadsheet();

    // 2. Save Follow-Up via GET (fastest, atomic append ~2-3s)
    if (action === 'save') {
      const payload = params.data ? JSON.parse(params.data) : params;
      return handleSaveRecord(ss, payload);
    }

    // 3. Fetch Master Data for in-memory caching (with CacheService acceleration)
    if (action === 'getMaster') {
      return handleGetMaster(ss, params.forceRefresh === 'true');
    }

    // 4. Check & Pre-allocate rows endpoint
    if (action === 'checkAndAddRows') {
      return handleCheckAndAddRows(ss);
    }

    return jsonResponse({ success: false, error: 'Unknown action: ' + action });
  } catch (err) {
    return jsonResponse({ success: false, error: err.toString() });
  }
}

function doPost(e) {
  try {
    let payload = {};
    if (e && e.postData && e.postData.contents) {
      try {
        payload = JSON.parse(e.postData.contents);
      } catch (ex) {
        payload = e.parameter || {};
      }
    } else if (e && e.parameter) {
      payload = e.parameter;
    }

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    return handleSaveRecord(ss, payload);
  } catch (err) {
    return jsonResponse({ success: false, error: err.toString() });
  }
}

/**
 * Optimized Master Data fetcher with CacheService & blank row filtering
 */
function handleGetMaster(ss, forceRefresh) {
  const cache = CacheService.getScriptCache();
  if (!forceRefresh) {
    const cached = cache.get('master_data_json');
    if (cached) {
      return ContentService.createTextOutput(cached)
        .setMimeType(ContentService.MimeType.JSON);
    }
  }

  const sheet = ss.getSheetByName('Master') || ss.getSheets()[0];
  if (!sheet) {
    return jsonResponse({ success: false, error: 'Master sheet not found' });
  }

  const lastRow = sheet.getLastRow();
  if (lastRow < 6) {
    return jsonResponse({ success: true, data: [] });
  }

  // Read only the required 6 columns (starts from row 6: Header row)
  const numCols = Math.min(sheet.getLastColumn(), 6);
  const rawData = sheet.getRange(6, 1, lastRow - 5, numCols).getValues();

  // Filter out completely blank rows so the payload is tiny and fast
  const compactData = [];
  compactData.push(rawData[0]); // Header row

  for (let i = 1; i < rawData.length; i++) {
    const row = rawData[i];
    // Keep row only if mobile or name has content
    if ((row[1] && String(row[1]).trim()) || (row[2] && String(row[2]).trim())) {
      compactData.push(row);
    }
  }

  const responseJson = JSON.stringify({ success: true, data: compactData });

  // Store in Apps Script cache for 6 hours (21600 seconds) if under 100KB
  try {
    if (responseJson.length < 100000) {
      cache.put('master_data_json', responseJson, 21600);
    }
  } catch (e) {}

  return ContentService.createTextOutput(responseJson)
    .setMimeType(ContentService.MimeType.JSON);
}

function handleSaveRecord(ss, payload) {
  const sheet = ss.getSheetByName('Responses1');
  if (!sheet) {
    return jsonResponse({ success: false, error: 'Responses1 sheet not found' });
  }

  const {
    timestamp,
    submissionId,
    mobile,
    name,
    followup,
    rating,
    remark,
    anyIssues,
    reason
  } = payload;

  // 1. Find the REAL last filled row by inspecting Column B (Timestamp) & Column C (Mobile)
  const lastFilledRow = getLastFilledRow(sheet);
  const targetRow = lastFilledRow + 1;

  // 2. Duplicate protection: Check recent 50 submissions in Column K around lastFilledRow
  if (lastFilledRow >= 2 && submissionId) {
    const startRow = Math.max(2, lastFilledRow - 50);
    const numRows = lastFilledRow - startRow + 1;
    const recentSubmissions = sheet.getRange(startRow, 11, numRows, 1).getValues();
    for (let i = 0; i < recentSubmissions.length; i++) {
      if (String(recentSubmissions[i][0]).trim() === String(submissionId).trim()) {
        return jsonResponse({ success: true, message: 'Already saved', alreadySaved: true });
      }
    }
  }

  const timeFormatted = timestamp || Utilities.formatDate(new Date(), 'Asia/Kolkata', 'dd/MM/yyyy HH:mm:ss');

  // Format follow-up date as DD/MM/YYYY
  let formattedFollowup = followup || '';
  if (formattedFollowup && formattedFollowup.indexOf('-') > -1) {
    const dParts = formattedFollowup.split('-');
    if (dParts.length === 3) {
      formattedFollowup = dParts[2] + '/' + dParts[1] + '/' + dParts[0];
    }
  }

  // 3. Pre-allocate 500 rows if sheet is running low on space (<= 20 rows left from targetRow)
  const maxRows = sheet.getMaxRows();
  if (targetRow > maxRows) {
    sheet.insertRowsAfter(maxRows, Math.max(500, targetRow - maxRows + 100));
  } else if (maxRows - targetRow <= 20) {
    sheet.insertRowsAfter(maxRows, 500);
  }

  // 4. Ensure formulas exist in Column A and Column E for targetRow
  if (targetRow >= 3) {
    try {
      const prevRow = targetRow - 1;
      // Copy formula in Col A if targetRow cell doesn't already have one
      const targetACell = sheet.getRange(targetRow, 1);
      if (!targetACell.getFormula()) {
        const prevACell = sheet.getRange(prevRow, 1);
        if (prevACell.getFormula()) {
          prevACell.copyTo(targetACell, SpreadsheetApp.CopyPasteType.PASTE_FORMULA, false);
        }
      }

      // Copy formula in Col E if targetRow cell doesn't already have one
      const targetECell = sheet.getRange(targetRow, 5);
      if (!targetECell.getFormula()) {
        const prevECell = sheet.getRange(prevRow, 5);
        if (prevECell.getFormula()) {
          prevECell.copyTo(targetECell, SpreadsheetApp.CopyPasteType.PASTE_FORMULA, false);
        }
      }
    } catch (e) {}
  }

  // 5. Write data into the EXACT next filled row (targetRow)
  // Column B to D: [Timestamp, Mobile, Name]
  sheet.getRange(targetRow, 2, 1, 3).setValues([
    [timeFormatted, mobile || '', name || '']
  ]);

  // Column F to K: [Follow-Up Date, Rating, Remark, Any Issues, Reason, Submission ID]
  sheet.getRange(targetRow, 6, 1, 6).setValues([
    [
      formattedFollowup,
      rating || '',
      remark || '',
      anyIssues || 'No',
      anyIssues === 'Yes' ? (reason || '') : '',
      submissionId || ''
    ]
  ]);

  return jsonResponse({
    success: true,
    message: 'Follow-up saved successfully',
    row: targetRow
  });
}

/**
 * Finds the actual last filled data row in Responses1
 * Checks Column B (Timestamp) and Column C (Mobile) from bottom to top
 */
function getLastFilledRow(sheet) {
  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) return 1;

  // Read Column B & C values up to lastRow
  const colBC = sheet.getRange(1, 2, lastRow, 2).getValues();

  // Scan backwards from bottom to top to find the first non-empty cell in Col B or Col C
  for (let r = colBC.length - 1; r >= 1; r--) {
    const b = colBC[r][0];
    const c = colBC[r][1];
    if ((b !== '' && b !== null && b !== undefined) ||
        (c !== '' && c !== null && c !== undefined)) {
      return r + 1; // 1-indexed row number
    }
  }

  return 1; // If only header row has content
}

/**
 * Endpoint helper to check and add rows
 */
function handleCheckAndAddRows(ss) {
  const sheet = ss.getSheetByName('Responses1');
  if (!sheet) {
    return jsonResponse({ success: false, error: 'Responses1 sheet not found' });
  }

  const maxRows = sheet.getMaxRows();
  const lastFilledRow = getLastFilledRow(sheet);
  const emptyRowsLeft = maxRows - lastFilledRow;
  let added = false;

  if (emptyRowsLeft <= 20) {
    sheet.insertRowsAfter(maxRows, 500);
    added = true;
  }

  return jsonResponse({
    success: true,
    added500Rows: added,
    lastFilledRow: lastFilledRow,
    totalRows: sheet.getMaxRows(),
    emptyRowsLeft: sheet.getMaxRows() - lastFilledRow
  });
}

/**
 * Adds a custom menu in Google Sheets so you can check and add rows with 1 click
 */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('CRM Tools')
    .addItem('Check & Add 500 Rows to Responses1', 'checkAndAdd500Rows')
    .addToUi();
}

/**
 * Checks Responses1 and adds 500 rows if 20 or fewer rows remain after last filled row.
 * Runs instantly from the "CRM Tools" menu inside your Google Spreadsheet.
 */
function checkAndAdd500Rows() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Responses1');
  if (!sheet) {
    SpreadsheetApp.getUi().alert('Responses1 sheet not found!');
    return;
  }

  const maxRows = sheet.getMaxRows();
  const lastFilledRow = getLastFilledRow(sheet);
  const emptyRowsLeft = maxRows - lastFilledRow;

  if (emptyRowsLeft <= 20) {
    sheet.insertRowsAfter(maxRows, 500);
    SpreadsheetApp.getUi().alert(
      '✅ Added 500 new rows to Responses1!\n\n' +
      'Last filled row (Col B): ' + lastFilledRow + '\n' +
      'Previous total rows: ' + maxRows + '\n' +
      'New total rows: ' + sheet.getMaxRows() + '\n' +
      'Empty rows available: ' + (sheet.getMaxRows() - lastFilledRow)
    );
  } else {
    SpreadsheetApp.getUi().alert(
      'ℹ️ Sufficient rows available!\n\n' +
      'Last filled row (Col B): ' + lastFilledRow + '\n' +
      'Total rows in sheet: ' + maxRows + '\n' +
      'Empty rows remaining: ' + emptyRowsLeft + '\n\n' +
      'No extra rows needed right now (buffer threshold is 20 rows).'
    );
  }
}

function jsonResponse(data) {
  return ContentService.createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}
