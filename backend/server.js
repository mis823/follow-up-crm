const express = require('express');
const cors = require('cors');
const path = require('path');
const config = require('./config');
const {
  initializeGoogleSheets,
  loadMasterData,
  searchCustomers,
  appendFollowUp
} = require('./googleSheets');

const app = express();

// Middlewares
app.use(cors());
app.use(express.json());
app.use(express.urlencoded({ extended: true }));

// Serve frontend static files
const frontendDir = path.join(__dirname, '..', 'frontend');
app.use(express.static(frontendDir));

/**
 * Health check endpoint
 * GET /api/health
 */
app.get('/api/health', (req, res) => {
  res.json({
    success: true,
    message: 'Follow-Up CRM API is running',
    configured: config.isConfigured
  });
});

/**
 * Customer search endpoint
 * GET /api/search?mobile=98765
 */
app.get('/api/search', async (req, res, next) => {
  const startTime = Date.now();
  const mobile = req.query.mobile || '';

  try {
    const results = await searchCustomers(mobile);
    const duration = Date.now() - startTime;
    console.log(`SEARCH prefix=${mobile} results=${results.length} duration=${duration}ms`);
    return res.json(results);
  } catch (err) {
    const duration = Date.now() - startTime;
    console.error(`SEARCH_ERROR prefix=${mobile} duration=${duration}ms:`, err.message);
    next(err);
  }
});

/**
 * Follow-up save endpoint
 * POST /api/save
 */
app.post('/api/save', async (req, res, next) => {
  const startTime = Date.now();
  const mobile = req.body.mobile || 'unknown';

  try {
    const result = await appendFollowUp(req.body);
    const duration = Date.now() - startTime;
    console.log(`SAVE mobile=${mobile} success duration=${duration}ms`);
    return res.json(result);
  } catch (err) {
    const duration = Date.now() - startTime;
    console.error(`SAVE_ERROR mobile=${mobile} duration=${duration}ms:`, err.message);
    const status = err.status || 500;
    return res.status(status).json({
      success: false,
      message: err.message || 'An error occurred while saving the follow-up record.'
    });
  }
});

// Fallback route for SPA
app.get('*', (req, res) => {
  res.sendFile(path.join(frontendDir, 'index.html'));
});

// Global error handler
app.use((err, req, res, next) => {
  console.error('Unhandled Server Error:', err);
  const status = err.status || 500;
  res.status(status).json({
    success: false,
    message: err.message || 'Internal Server Error'
  });
});

// Start Express server
const server = app.listen(config.port, async () => {
  console.log(`====================================================`);
  console.log(`🚀 Follow-Up CRM Server running on http://localhost:${config.port}`);
  console.log(`📂 Frontend served from: ${frontendDir}`);
  console.log(`====================================================`);

  // Initialize Google Sheets and warm up cache
  initializeGoogleSheets();
  try {
    await loadMasterData();
  } catch (err) {
    console.error('Cache initialization error:', err.message);
  }
});

module.exports = { app, server };
