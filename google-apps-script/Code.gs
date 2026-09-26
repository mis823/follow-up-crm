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
    const action = params.action || 'getMaster';
    const ss = SpreadsheetApp.getActiveSpreadsheet();

    // 1. Fetch Master Data for in-memory caching
    if (action === 'getMaster') {
      const sheet = ss.getSheetByName('Master') || ss.getSheets()[0];
      if (!sheet) {
        return jsonResponse({ success: false, error: 'Master sheet not found' });
      }

      const lastRow = sheet.getLastRow();
      const lastCol = sheet.getLastColumn();
      if (lastRow < 6 || lastCol < 1) {
        return jsonResponse({ success: true, data: [] });
      }

      // Read from Row 6 (header) down to last row
      const data = sheet.getRange(6, 1, lastRow - 5, lastCol).getValues();
      return jsonResponse({ success: true, data: data });
    }

    // 2. Save Follow-Up via GET (handles HTTP redirects seamlessly)
    if (action === 'save') {
      const payload = params.data ? JSON.parse(params.data) : params;
      return handleSaveRecord(ss, payload);
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

  const newRow = lastRow + 1;
  const timeFormatted = timestamp || Utilities.formatDate(new Date(), 'Asia/Kolkata', 'dd/MM/yyyy HH:mm:ss');

  // Column B: Timestamp, Column C: Mobile No., Column D: Customer Name
  sheet.getRange(newRow, 2, 1, 3).setValues([[
    timeFormatted,
    mobile || '',
    name || ''
  ]]);

  // Column F: Follow-Up Date, Column G: Rating, Column H: Remark, Column I: Any Issues, Column J: Reason, Column K: Submission ID
  sheet.getRange(newRow, 6, 1, 6).setValues([[
    followup || '',
    rating || '',
    remark || '',
    anyIssues || 'No',
    anyIssues === 'Yes' ? (reason || '') : '',
    submissionId || ''
  ]]);

  // If Column A (1) and Column E (5) have row formulas, copy them down to newRow
  if (lastRow >= 2) {
    const colAFormula = sheet.getRange(lastRow, 1).getFormulaR1C1();
    if (colAFormula) {
      sheet.getRange(newRow, 1).setFormulaR1C1(colAFormula);
    }

    const colEFormula = sheet.getRange(lastRow, 5).getFormulaR1C1();
    if (colEFormula) {
      sheet.getRange(newRow, 5).setFormulaR1C1(colEFormula);
    }
  }

  return jsonResponse({
    success: true,
    message: 'Follow-up saved successfully'
  });
}

function jsonResponse(data) {
  return ContentService.createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}
