const { google } = require('googleapis');
const config = require('./config');

// In-memory Master data store
let masterCustomers = [];
let prefixIndex = new Map(); // Map: 5-digit prefix -> Array of customer objects
let lastCacheTime = 0;
let isLoadingMaster = false;

// Processed submission IDs for fast duplicate protection
const processedSubmissions = new Set();

// Google Sheets API client instance
let sheetsClient = null;

// Mock master data used when real credentials are not yet configured in .env
const MOCK_CUSTOMERS = [
  { mobile: '9876543210', name: 'Rajesh Kumar', rating: '4', remark: 'Interested in annual maintenance plan' },
  { mobile: '9876512345', name: 'Priya Sharma', rating: '5', remark: 'Highly satisfied with previous service' },
  { mobile: '9876598765', name: 'Amitabh Verma', rating: '3', remark: 'Requested follow-up regarding pricing discounts' },
  { mobile: '9876533445', name: 'Sunita Patel', rating: '4', remark: 'Needs quote for bulk purchase' },
  { mobile: '9876588776', name: 'Vikram Singh', rating: '2', remark: 'Reported delayed delivery last month' },
  { mobile: '9876500112', name: 'Ananya Gupta', rating: '5', remark: 'VIP customer, prefers weekend calls' },
  { mobile: '9876522334', name: 'Rohan Mehra', rating: '', remark: '' },
  { mobile: '9876555667', name: 'Kavita Joshi', rating: '4', remark: 'Inquired about renewal terms' },
  { mobile: '9876577889', name: 'Deepak Rao', rating: '3', remark: 'Follow up next quarter' },
  { mobile: '9876599001', name: 'Neha Choudhary', rating: '5', remark: 'Requested product catalog' },
  { mobile: '9123456789', name: 'Sanjay Kapoor', rating: '4', remark: 'Asked for demo session' },
  { mobile: '9123498765', name: 'Meera Nambiar', rating: '5', remark: 'Referred by another client' },
  { mobile: '9988776655', name: 'Rahul Deshmukh', rating: '1', remark: 'Past issue with billing statement' }
];

/**
 * Initialize Google Sheets API client or Webhook Bridge
 */
function initializeGoogleSheets() {
  if (config.isAppsScriptConfigured) {
    console.log(`✅ [GoogleSheets] Using Google Apps Script Webhook Bridge (100% Free Zero-Cloud Mode).`);
    console.log(`🔗 Endpoint: ${config.appsScriptUrl.substring(0, 45)}...`);
    return true;
  }

  if (config.isGoogleCloudConfigured) {
    try {
      const auth = new google.auth.JWT({
        email: config.clientEmail,
        key: config.privateKey,
        scopes: ['https://www.googleapis.com/auth/spreadsheets']
      });

      sheetsClient = google.sheets({ version: 'v4', auth });
      console.log('✅ [GoogleSheets] Google Cloud Service account authentication initialized.');
      return sheetsClient;
    } catch (err) {
      console.error('❌ [GoogleSheets] Failed to initialize Google Sheets client:', err.message);
      sheetsClient = null;
      return null;
    }
  }

  console.warn('\n⚠️  [GoogleSheets] Live credentials not set in .env. Running in DEMO/MOCK mode.');
  console.warn('👉 To connect your live sheet for FREE, add APPS_SCRIPT_WEBHOOK_URL to your .env file.\n');
  return null;
}

/**
 * Rebuild the in-memory prefix index for fast search
 */
function rebuildIndex(customers) {
  masterCustomers = customers;
  prefixIndex.clear();

  for (let i = 0; i < customers.length; i++) {
    const c = customers[i];
    if (!c.mobile || c.mobile.length < 5) continue;

    // Index the first 5 digits
    const prefix5 = c.mobile.substring(0, 5);
    if (!prefixIndex.has(prefix5)) {
      prefixIndex.set(prefix5, []);
    }
    prefixIndex.get(prefix5).push(c);
  }
}

let masterLoadingPromise = null;

/**
 * Load customer Master data from Google Sheets into memory/index
 * Supports starting from Row 7 as specified in business logic.
 */
async function loadMasterData(forceRefresh = false) {
  const now = Date.now();
  if (!forceRefresh && masterCustomers.length > 0 && now - lastCacheTime < config.cacheTtlMs) {
    return masterCustomers;
  }

  if (masterLoadingPromise) {
    return masterLoadingPromise;
  }

  masterLoadingPromise = (async () => {
    const startTime = Date.now();

    try {
      let rows = [];

      // Mode 1: Google Apps Script Webhook
      if (config.isAppsScriptConfigured) {
        console.log(`🔄 [Master Cache] Fetching Master sheet via Apps Script Webhook...`);
        const fetchUrl = `${config.appsScriptUrl}?action=getMaster&_t=${Date.now()}`;
        const controller = new AbortController();
        const timeout = setTimeout(() => controller.abort(), 60000);

        try {
          const response = await fetch(fetchUrl, { redirect: 'follow', signal: controller.signal });
          clearTimeout(timeout);
          if (!response.ok) {
            throw new Error(`Webhook returned status ${response.status}`);
          }
          const json = await response.json();
          if (!json.success) {
            throw new Error(json.error || 'Failed to fetch Master data from Webhook');
          }
          rows = json.data || [];
        } catch (fetchErr) {
          clearTimeout(timeout);
          throw fetchErr;
        }
      }
      // Mode 2: Google Cloud Service Account API
      else if (config.isGoogleCloudConfigured && sheetsClient) {
        console.log(`🔄 [Master Cache] Fetching Master sheet starting from row ${config.masterStartRow}...`);
        const headerRow = Math.max(1, config.masterStartRow - 1);
        const range = `${config.masterSheetName}!A${headerRow}:Z`;

        const response = await sheetsClient.spreadsheets.values.get({
          spreadsheetId: config.spreadsheetId,
          range
        });
        rows = response.data.values || [];
      }
      // Mode 3: Demo / Mock Mode
      else {
        rebuildIndex(MOCK_CUSTOMERS);
        lastCacheTime = Date.now();
        console.log(`ℹ️ [Master Cache] Loaded ${MOCK_CUSTOMERS.length} demo customer records into memory index.`);
        return masterCustomers;
      }

      if (rows.length === 0) {
        console.warn(`⚠️ [Master Cache] No data returned for Master sheet.`);
        rebuildIndex([]);
        lastCacheTime = Date.now();
        return [];
      }

    // Determine column indices by matching first occurrences
    let mobileIdx = -1;
    let nameIdx = -1;
    let ratingIdx = -1;
    let remarkIdx = -1;

    const headers = rows[0] || [];
    let dataStartOffset = 1; // Since row 0 in 'rows' is row 6 (the header)

    // Inspect headers
    headers.forEach((h, idx) => {
      if (!h) return;
      const lower = String(h).trim().toLowerCase();
      if (mobileIdx === -1 && (lower === 'mobile no.' || lower.includes('mobile') || lower.includes('phone') || lower.includes('contact'))) {
        mobileIdx = idx;
      } else if (nameIdx === -1 && (lower === 'name' || lower.includes('customer name') || lower.includes('client'))) {
        nameIdx = idx;
      } else if (ratingIdx === -1 && (lower === 'rating' || lower.includes('previous rating') || lower.includes('rate'))) {
        ratingIdx = idx;
      } else if (remarkIdx === -1 && (lower === 'remark' || lower.includes('previous remark') || lower.includes('feedback'))) {
        remarkIdx = idx;
      }
    });

    // Fallbacks if not detected
    if (mobileIdx === -1) mobileIdx = 1;
    if (nameIdx === -1) nameIdx = 2;
    if (ratingIdx === -1) ratingIdx = 4;
    if (remarkIdx === -1) remarkIdx = 5;

    console.log(`📋 [Column Mapping] Mobile: Col ${mobileIdx}, Name: Col ${nameIdx}, Rating: Col ${ratingIdx}, Remark: Col ${remarkIdx}`);

    const parsedCustomers = [];
    for (let r = dataStartOffset; r < rows.length; r++) {
      const row = rows[r];
      if (!row || row.length === 0) continue;

      const rawMobile = row[mobileIdx] ? String(row[mobileIdx]).replace(/\D/g, '').trim() : '';
      if (!rawMobile) continue;

      const name = row[nameIdx] ? String(row[nameIdx]).trim() : '';
      const rating = row[ratingIdx] ? String(row[ratingIdx]).trim() : '';
      const remark = row[remarkIdx] ? String(row[remarkIdx]).trim() : '';

      parsedCustomers.push({
        mobile: rawMobile,
        name,
        rating,
        remark
      });
    }

    rebuildIndex(parsedCustomers);
    lastCacheTime = Date.now();
    const duration = Date.now() - startTime;
    console.log(`✅ [Master Cache] Successfully indexed ${parsedCustomers.length} customers from Master sheet in ${duration}ms.`);
    return masterCustomers;
  } catch (err) {
    console.error(`❌ [Master Cache] Failed to load Master sheet:`, err.message);
    if (masterCustomers.length === 0) {
      // Fallback to mock data so system continues operating
      rebuildIndex(MOCK_CUSTOMERS);
    }
    } finally {
      masterLoadingPromise = null;
    }
  })();

  return masterLoadingPromise;
}

/**
 * Fast search customer by mobile prefix OR customer name
 * Searches in-memory index; returns up to 20 matching customers
 * Triggered at 3 or more characters
 */
async function searchCustomers(query) {
  if (!query) return [];
  const rawQuery = String(query).trim();
  if (rawQuery.length < 3) return [];

  // Ensure cache is loaded
  if (masterCustomers.length === 0 || Date.now() - lastCacheTime > config.cacheTtlMs) {
    await loadMasterData();
  }

  const queryLower = rawQuery.toLowerCase();
  const digitQuery = rawQuery.replace(/\D/g, '');

  const results = [];
  const maxResults = 20;

  // 1. If query contains 3+ digits, search by mobile
  if (digitQuery.length >= 3) {
    // Check prefix index if 5+ digits
    if (digitQuery.length >= 5) {
      const prefix5 = digitQuery.substring(0, 5);
      const candidates = prefixIndex.get(prefix5) || [];
      for (let i = 0; i < candidates.length; i++) {
        const cust = candidates[i];
        if (cust.mobile.includes(digitQuery)) {
          results.push(cust);
          if (results.length >= maxResults) return results;
        }
      }
    }

    // Search across all customers by mobile number
    for (let i = 0; i < masterCustomers.length; i++) {
      const cust = masterCustomers[i];
      if (cust.mobile && cust.mobile.includes(digitQuery)) {
        if (!results.includes(cust)) {
          results.push(cust);
          if (results.length >= maxResults) return results;
        }
      }
    }
  }

  // 2. Search by customer name
  for (let i = 0; i < masterCustomers.length; i++) {
    const cust = masterCustomers[i];
    if (cust.name && cust.name.toLowerCase().includes(queryLower)) {
      if (!results.includes(cust)) {
        results.push(cust);
        if (results.length >= maxResults) return results;
      }
    }
  }

  return results;
}

/**
 * Append follow-up record to Responses1 sheet
 */
async function appendFollowUp(payload) {
  const {
    mobile,
    name,
    followup,
    rating,
    remark,
    anyIssues,
    reason,
    submissionId
  } = payload;

  // Duplicate protection check
  if (submissionId && processedSubmissions.has(submissionId)) {
    return {
      success: true,
      message: 'Already saved',
      alreadySaved: true
    };
  }

  // Validate mobile
  const cleanMobile = String(mobile || '').replace(/\D/g, '').trim();
  if (!cleanMobile || cleanMobile.length !== 10) {
    const err = new Error('Invalid mobile number. Mobile number must be exactly 10 digits.');
    err.status = 400;
    throw err;
  }

  // Validate name
  const cleanName = String(name || '').trim();
  if (!cleanName) {
    const err = new Error('Customer Name is required.');
    err.status = 400;
    throw err;
  }

  // Validate follow-up date
  if (!followup) {
    const err = new Error('Follow-up date is required.');
    err.status = 400;
    throw err;
  }

  const today = new Date();
  today.setHours(0, 0, 0, 0);

  const maxDate = new Date(today);
  maxDate.setDate(maxDate.getDate() + 30);
  maxDate.setHours(23, 59, 59, 999);

  const selectedDate = new Date(followup + 'T00:00:00');
  if (isNaN(selectedDate.getTime())) {
    const err = new Error('Invalid date format.');
    err.status = 400;
    throw err;
  }

  if (selectedDate < today) {
    const err = new Error('Follow-up date cannot be in the past.');
    err.status = 400;
    throw err;
  }

  if (selectedDate > maxDate) {
    const err = new Error('Follow-up date cannot be more than 30 days ahead.');
    err.status = 400;
    throw err;
  }

  // Validate rating
  const numRating = parseInt(rating, 10);
  if (isNaN(numRating) || numRating < 1 || numRating > 5) {
    const err = new Error('Rating is required and must be between 1 and 5.');
    err.status = 400;
    throw err;
  }

  // Validate Any Issues & Reason
  const normalizedAnyIssues = (anyIssues === 'Yes' || anyIssues === true || anyIssues === 'true') ? 'Yes' : 'No';
  const cleanReason = String(reason || '').trim();
  if (normalizedAnyIssues === 'Yes' && !cleanReason) {
    const err = new Error('Reason is required when "Any Issues" is checked.');
    err.status = 400;
    throw err;
  }

  const cleanRemark = String(remark || '').trim();
  const timestamp = new Date().toLocaleString('en-IN', { timeZone: 'Asia/Kolkata' });

  // Responses1 Row structure:
  // [Col A: (formula), Col B: Timestamp, Col C: Mobile, Col D: Name, Col E: (formula), Col F: Follow-up, Col G: Rating, Col H: Remark, Col I: Any Issues, Col J: Reason, Col K: Submission ID]
  const rowValues = [
    '', // Col A (formula)
    timestamp, // Col B
    cleanMobile, // Col C
    cleanName, // Col D
    '', // Col E (formula)
    followup, // Col F
    numRating, // Col G
    cleanRemark, // Col H
    normalizedAnyIssues, // Col I
    normalizedAnyIssues === 'Yes' ? cleanReason : '', // Col J
    submissionId || '' // Col K
  ];

  if (config.isAppsScriptConfigured) {
    // Mode 1: Save via Apps Script Webhook Bridge
    const savePayload = {
      timestamp,
      submissionId: submissionId || '',
      mobile: cleanMobile,
      name: cleanName,
      followup,
      rating: numRating,
      remark: cleanRemark,
      anyIssues: normalizedAnyIssues,
      reason: normalizedAnyIssues === 'Yes' ? cleanReason : ''
    };

    const webhookUrl = `${config.appsScriptUrl}?action=save&data=${encodeURIComponent(JSON.stringify(savePayload))}`;
    const res = await fetch(webhookUrl, {
      method: 'GET',
      redirect: 'follow'
    });

    if (!res.ok) {
      throw new Error(`Webhook save request failed with status ${res.status}`);
    }

    const json = await res.json();
    if (!json.success && !json.alreadySaved) {
      throw new Error(json.error || 'Failed to save follow-up via Apps Script Webhook');
    }

    if (json.alreadySaved) {
      if (submissionId) processedSubmissions.add(submissionId);
      return { success: true, message: 'Already saved', alreadySaved: true };
    }
  } else if (config.isGoogleCloudConfigured && sheetsClient) {
    // Mode 2: Save via Google Cloud Sheets API
    const range = `${config.responsesSheetName}!A:I`;
    await sheetsClient.spreadsheets.values.append({
      spreadsheetId: config.spreadsheetId,
      range,
      valueInputOption: 'USER_ENTERED',
      insertDataOption: 'INSERT_ROWS',
      requestBody: {
        values: [rowValues]
      }
    });
  } else {
    // Mode 3: Demo mode simulated instant append
    console.log(`ℹ️ [Demo Mode Save] Simulated append to ${config.responsesSheetName}:`, rowValues);
  }

  // Record submission ID in memory to prevent double submit
  if (submissionId) {
    processedSubmissions.add(submissionId);
  }

  return {
    success: true,
    message: 'Follow-up saved successfully',
    data: {
      submissionId,
      mobile: cleanMobile,
      name: cleanName
    }
  };
}

module.exports = {
  initializeGoogleSheets,
  loadMasterData,
  searchCustomers,
  appendFollowUp
};
