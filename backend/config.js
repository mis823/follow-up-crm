const path = require('path');
const dotenv = require('dotenv');

// Load environment variables from .env in project root
dotenv.config({ path: path.join(__dirname, '..', '.env') });

const rawPrivateKey = process.env.GOOGLE_PRIVATE_KEY || '';
// Handle formatted newlines in private key string
const formattedPrivateKey = rawPrivateKey.includes('\\n')
  ? rawPrivateKey.replace(/\\n/g, '\n')
  : rawPrivateKey;

const spreadsheetId = (process.env.GOOGLE_SPREADSHEET_ID || '').trim();
const clientEmail = (process.env.GOOGLE_CLIENT_EMAIL || '').trim();
const appsScriptUrl = (process.env.APPS_SCRIPT_WEBHOOK_URL || '').trim();

const isGoogleCloudConfigured = Boolean(
  spreadsheetId &&
  !spreadsheetId.includes('YOUR_SPREADSHEET_ID') &&
  clientEmail &&
  !clientEmail.includes('YOUR_SERVICE_ACCOUNT_EMAIL') &&
  formattedPrivateKey &&
  !formattedPrivateKey.includes('YOUR_PRIVATE_KEY')
);

const isAppsScriptConfigured = Boolean(
  appsScriptUrl &&
  !appsScriptUrl.includes('YOUR_APPS_SCRIPT_URL') &&
  appsScriptUrl.startsWith('http')
);

const isConfigured = isGoogleCloudConfigured || isAppsScriptConfigured;

module.exports = {
  port: parseInt(process.env.PORT || '3000', 10),
  spreadsheetId,
  clientEmail,
  privateKey: formattedPrivateKey,
  appsScriptUrl,
  isAppsScriptConfigured,
  isGoogleCloudConfigured,
  masterSheetName: process.env.MASTER_SHEET_NAME || 'Master',
  responsesSheetName: process.env.RESPONSES_SHEET_NAME || 'Responses1',
  masterStartRow: parseInt(process.env.MASTER_START_ROW || '7', 10),
  cacheTtlMs: parseInt(process.env.MASTER_CACHE_TTL_MINUTES || '10', 10) * 60 * 1000,
  isConfigured
};
