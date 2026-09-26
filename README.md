# Follow-Up CRM

A high-performance standalone customer follow-up web application integrated with Google Sheets API and Google Apps Script Webhook.

## Features
- **In-Memory Search Index**: Searches 20,000+ customer records in milliseconds (< 5ms) by mobile prefix.
- **Fast Append**: Submits follow-up records directly to the `Responses1` sheet in seconds.
- **Duplicate Submission Protection**: Guarantees no accidental double-saves using unique submission IDs.
- **Modern Bootstrap 5 UI**:
  - 10-digit mobile number validation and autocomplete suggestions.
  - Readonly customer name locking with auto-filled previous ratings and remarks.
  - Follow-up date validation (today to today + 30 days).
  - 1–5 Star rating component.
  - Dynamic "Any Issues?" toggle with mandatory reason description.
- **100% Free / Zero Billing**: Operates with Google Apps Script Webhook bridge or Google Cloud Service Accounts.

---

## Project Structure
```text
follow-up-crm/
│
├── frontend/
│   ├── index.html       # Customer follow-up form & UI
│   ├── style.css        # Clean responsive styles & star animations
│   └── app.js           # Debounce search, validation, and REST API calls
│
├── backend/
│   ├── server.js        # Express HTTP server & static frontend hosting
│   ├── googleSheets.js  # In-memory index & Google Sheets connector
│   └── config.js        # Centralized configuration loader
│
├── google-apps-script/
│   └── Code.gs          # Free Google Apps Script Webhook bridge code
│
├── package.json
└── .env.example
```

---

## Local Setup

1. **Install Dependencies**:
   ```bash
   npm install
   ```

2. **Configure Environment Variables**:
   Copy `.env.example` to `.env`:
   ```bash
   cp .env.example .env
   ```
   Add your Google Apps Script Webhook URL (from your Google Sheet):
   ```env
   PORT=3000
   APPS_SCRIPT_WEBHOOK_URL="https://script.google.com/macros/s/YOUR_ID/exec"
   ```

3. **Start the Application**:
   ```bash
   npm start
   ```
   Open [http://localhost:3000](http://localhost:3000).

---

## 24/7 Deployment on Render.com (100% Free)

1. Sign up on [Render.com](https://render.com/) (Free Tier).
2. Click **New +** > **Web Service** and select this repository.
3. Configure:
   - **Runtime**: `Node`
   - **Build Command**: `npm install`
   - **Start Command**: `npm start`
   - **Plan**: `Free`
4. In **Environment Variables**, add:
   - `APPS_SCRIPT_WEBHOOK_URL` = your Google Apps Script Webhook URL.
5. Click **Deploy**. Your app is live 24/7 on the internet.
