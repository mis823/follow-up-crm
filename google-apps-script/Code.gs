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

  const lastRow = sheet.getLastRow();

  // Duplicate protection: Check recent 50 submissions in Column K (Column 11)
  if (lastRow >= 2 && submissionId) {
    const startRow = Math.max(2, lastRow - 50);
    const numRows = lastRow - startRow + 1;
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

  // If Column A or Column E have formulas, read them once from previous row
  let colAVal = '';
  let colEVal = '';
  if (lastRow >= 2) {
    try {
      const formulas = sheet.getRange(lastRow, 1, 1, 5).getFormulasR1C1()[0];
      if (formulas && formulas[0]) colAVal = formulas[0];
      if (formulas && formulas[4]) colEVal = formulas[4];
    } catch (e) {}
  }

  // Pre-allocate 500 rows if sheet is running low on space (<= 20 rows left)
  // This completely eliminates Google Sheets delay on auto-expanding rows
  const maxRows = sheet.getMaxRows();
  if (maxRows - lastRow <= 20) {
    sheet.insertRowsAfter(maxRows, 500);
  }

  // Single fast atomic appendRow
  sheet.appendRow([
    colAVal,            // Column A (formula)
    timeFormatted,      // Column B: Timestamp (DD/MM/YYYY HH:mm:ss)
    mobile || '',       // Column C: Mobile No.
    name || '',         // Column D: Customer Name
    colEVal,            // Column E (formula)
    formattedFollowup,  // Column F: Follow-Up Date (DD/MM/YYYY)
    rating || '',       // Column G: Rating
    remark || '',       // Column H: Remark
    anyIssues || 'No',                         // Column I: Any Issues
    anyIssues === 'Yes' ? (reason || '') : '', // Column J: Reason
    submissionId || ''                         // Column K: Submission ID
  ]);

  return jsonResponse({
    success: true,
    message: 'Follow-up saved successfully'
  });
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
  const lastRow = sheet.getLastRow();
  const emptyRowsLeft = maxRows - lastRow;
  let added = false;

  if (emptyRowsLeft <= 20) {
    sheet.insertRowsAfter(maxRows, 500);
    added = true;
  }

  return jsonResponse({
    success: true,
    added500Rows: added,
    totalRows: sheet.getMaxRows(),
    emptyRowsLeft: sheet.getMaxRows() - sheet.getLastRow()
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
 * Checks Column B of Responses1 and adds 500 rows if 20 or fewer rows remain.
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
  const lastRow = sheet.getLastRow();
  const emptyRowsLeft = maxRows - lastRow;

  if (emptyRowsLeft <= 20) {
    sheet.insertRowsAfter(maxRows, 500);
    SpreadsheetApp.getUi().alert(
      '✅ Added 500 new rows to Responses1!\n\n' +
      'Previous total rows: ' + maxRows + '\n' +
      'New total rows: ' + sheet.getMaxRows() + '\n' +
      'Empty rows available: ' + (sheet.getMaxRows() - lastRow)
    );
  } else {
    SpreadsheetApp.getUi().alert(
      'ℹ️ Sufficient rows available!\n\n' +
      'Total rows in sheet: ' + maxRows + '\n' +
      'Last row with data: ' + lastRow + '\n' +
      'Empty rows remaining: ' + emptyRowsLeft + '\n\n' +
      'No extra rows needed right now (buffer threshold is 20 rows).'
    );
  }
}

function jsonResponse(data) {
  return ContentService.createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}
