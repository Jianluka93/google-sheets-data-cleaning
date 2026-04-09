/**
 * Google Sheets Data Cleaning Automation
 * 
 * Reads raw data from:   Raw_Data
 * Writes clean data to:  Clean_Data
 * Logs discarded rows:   Cleaning_Log
 *
 * Features:
 *  - Auto-detects header row
 *  - Flexible column name aliases
 *  - Email format validation
 *  - Duplicate email detection
 *  - Sales threshold filtering (≥ 100)
 *  - Full audit trail in Cleaning_Log
 *  - Summary popup after each run
 *  - Custom menu in Google Sheets UI
 */

// ─────────────────────────────────────────────
// COLUMN ALIASES  (extend freely)
// ─────────────────────────────────────────────
const NAME_ALIASES  = ["name", "full name", "customer name", "client name", "nome"];
const EMAIL_ALIASES = ["email", "e-mail", "mail", "email address"];
const SALES_ALIASES = ["sales", "amount", "revenue", "value", "total sales", "vendite", "totale"];

// Minimum sales value to keep a row
const SALES_THRESHOLD = 100;

// ─────────────────────────────────────────────
// CUSTOM MENU
// ─────────────────────────────────────────────

/**
 * Creates the "Data Tools" menu when the spreadsheet is opened.
 */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("🧹 Data Tools")
    .addItem("Run Data Cleaning", "runCleaning")
    .addSeparator()
    .addItem("Clear Clean_Data & Log", "clearOutputSheets")
    .addToUi();
}

// ─────────────────────────────────────────────
// PUBLIC ENTRY POINTS
// ─────────────────────────────────────────────

/**
 * Entry point called by the menu / button / manual run.
 * Shows a confirmation dialog, runs the cleaning, then shows a summary.
 */
function runCleaning() {
  const ui = SpreadsheetApp.getUi();

  const confirm = ui.alert(
    "Data Cleaning",
    "This will overwrite Clean_Data and append new rows to Cleaning_Log.\n\nContinue?",
    ui.ButtonSet.YES_NO
  );

  if (confirm !== ui.Button.YES) {
    ui.alert("Operation cancelled.");
    return;
  }

  try {
    const result = cleanData_withLogging();

    ui.alert(
      "✅ Cleaning complete",
      `Total data rows processed : ${result.totalRows}\n` +
      `Valid rows written         : ${result.keptRows}\n` +
      `Rows discarded (logged)    : ${result.totalRows - result.keptRows}`,
      ui.ButtonSet.OK
    );
  } catch (err) {
    ui.alert("❌ Error", err.message, ui.ButtonSet.OK);
    Logger.log("ERROR: " + err.message);
  }
}

/**
 * Clears Clean_Data content and the Cleaning_Log (keeps the header row).
 */
function clearOutputSheets() {
  const ui   = SpreadsheetApp.getUi();
  const ss   = SpreadsheetApp.getActiveSpreadsheet();

  const confirm = ui.alert(
    "Clear output sheets",
    "This will delete all content in Clean_Data and Cleaning_Log.\n\nContinue?",
    ui.ButtonSet.YES_NO
  );

  if (confirm !== ui.Button.YES) return;

  const clean = ss.getSheetByName("Clean_Data");
  const log   = ss.getSheetByName("Cleaning_Log");

  if (clean) clean.clearContents();
  if (log)   log.clearContents();

  ui.alert("Done. Both sheets have been cleared.");
}

// ─────────────────────────────────────────────
// CORE CLEANING LOGIC
// ─────────────────────────────────────────────

/**
 * Main cleaning function.
 * @returns {{ totalRows: number, keptRows: number }}
 */
function cleanData_withLogging() {
  const ss     = SpreadsheetApp.getActiveSpreadsheet();
  const source = ss.getSheetByName("Raw_Data");

  if (!source) {
    throw new Error('Sheet "Raw_Data" not found. Please create it first.');
  }

  // Get or create output sheets
  const dest     = ss.getSheetByName("Clean_Data")   || ss.insertSheet("Clean_Data");
  const logSheet = ss.getSheetByName("Cleaning_Log") || ss.insertSheet("Cleaning_Log");

  // Clear previous clean output; log is append-only
  dest.clearContents();

  // Read all raw data at once (faster than row-by-row API calls)
  const data = source.getDataRange().getValues();

  if (data.length === 0) {
    throw new Error("Raw_Data sheet is empty.");
  }

  // ── Detect header row ──────────────────────
  const { headerRowIndex, nameIndex, emailIndex, salesIndex } = detectHeaders(data);

  // ── Ensure log header exists ───────────────
  ensureLogHeader(logSheet);

  // ── Process rows ──────────────────────────
  const totalRows = data.length - (headerRowIndex + 1);
  const output    = [["Name", "Email", "Sales"]]; // clean output header
  const logBuffer = [];                            // batch-write to log
  const seenEmails = new Set();
  let keptRows = 0;

  for (let i = headerRowIndex + 1; i < data.length; i++) {
    const row       = data[i];
    const nome      = String(row[nameIndex]  ?? "").trim();
    const email     = String(row[emailIndex] ?? "").trim();
    const salesRaw  = row[salesIndex];

    // ── Validation checks (in priority order) ──

    // 1. Missing name or email
    if (!nome || !email) {
      logBuffer.push(buildLogRow(i + 1, "Missing Name or Email", nome, email, salesRaw));
      continue;
    }

    // 2. Email format
    if (!isValidEmail(email)) {
      logBuffer.push(buildLogRow(i + 1, "Invalid email format", nome, email, salesRaw));
      continue;
    }

    // 3. Duplicate email (within this run)
    if (seenEmails.has(email.toLowerCase())) {
      logBuffer.push(buildLogRow(i + 1, "Duplicate email", nome, email, salesRaw));
      continue;
    }

    // 4. Parse and clean the sales value
    const num = parseSales(salesRaw);

    if (isNaN(num)) {
      logBuffer.push(buildLogRow(i + 1, "Invalid sales value", nome, email, salesRaw));
      continue;
    }

    // 5. Sales threshold
    if (num < SALES_THRESHOLD) {
      logBuffer.push(buildLogRow(i + 1, `Sales below threshold (${SALES_THRESHOLD})`, nome, email, salesRaw));
      continue;
    }

    // ── Row is valid ──
    seenEmails.add(email.toLowerCase());
    output.push([nome, email, num]);
    keptRows++;
  }

  // ── Write output in one batch ──────────────
  if (output.length > 1) {
    dest.getRange(1, 1, output.length, 3).setValues(output);
  } else {
    // Write at least the header so the sheet isn't blank
    dest.getRange(1, 1, 1, 3).setValues([output[0]]);
  }

  // ── Append log rows in one batch ──────────
  if (logBuffer.length > 0) {
    const lastLogRow = logSheet.getLastRow();
    logSheet
      .getRange(lastLogRow + 1, 1, logBuffer.length, logBuffer[0].length)
      .setValues(logBuffer);
  }

  // ── Console summary ───────────────────────
  Logger.log(`Header row detected at index : ${headerRowIndex} (row ${headerRowIndex + 1})`);
  Logger.log(`Total data rows              : ${totalRows}`);
  Logger.log(`Valid rows kept              : ${keptRows}`);
  Logger.log(`Rows discarded               : ${totalRows - keptRows}`);

  return { totalRows, keptRows };
}

// ─────────────────────────────────────────────
// HELPER FUNCTIONS
// ─────────────────────────────────────────────

/**
 * Scans rows until it finds one containing all three required column types.
 * @param {any[][]} data
 * @returns {{ headerRowIndex, nameIndex, emailIndex, salesIndex }}
 */
function detectHeaders(data) {
  for (let r = 0; r < data.length; r++) {
    const headers = data[r].map(h => String(h).trim().toLowerCase());

    const nameIndex  = findAlias(headers, NAME_ALIASES);
    const emailIndex = findAlias(headers, EMAIL_ALIASES);
    const salesIndex = findAlias(headers, SALES_ALIASES);

    if (nameIndex !== -1 && emailIndex !== -1 && salesIndex !== -1) {
      return { headerRowIndex: r, nameIndex, emailIndex, salesIndex };
    }
  }

  throw new Error(
    "Could not detect header row. Make sure columns exist for Name, Email, and Sales.\n" +
    "Accepted aliases:\n" +
    `  Name  → ${NAME_ALIASES.join(", ")}\n` +
    `  Email → ${EMAIL_ALIASES.join(", ")}\n` +
    `  Sales → ${SALES_ALIASES.join(", ")}`
  );
}

/**
 * Returns the first index in `headers` that matches any alias, or -1.
 */
function findAlias(headers, aliases) {
  for (let i = 0; i < headers.length; i++) {
    if (aliases.includes(headers[i])) return i;
  }
  return -1;
}

/**
 * Strips currency symbols / whitespace and parses to float.
 * Handles both "." and "," as decimal separators.
 * Examples: "150€" → 150 | "1,250.00" → 1250 | "€ 99,5" → 99.5
 */
function parseSales(raw) {
  const str = String(raw ?? "").trim();

  // Remove everything that is not a digit, dot, or comma
  let cleaned = str.replace(/[^\d.,]/g, "");

  // If both separators present (e.g. "1,250.00"), remove comma as thousands sep
  if (cleaned.includes(",") && cleaned.includes(".")) {
    cleaned = cleaned.replace(/,/g, "");
  } else {
    // Treat comma as decimal separator
    cleaned = cleaned.replace(",", ".");
  }

  return parseFloat(cleaned);
}

/**
 * Basic email format check using a standard regex.
 */
function isValidEmail(email) {
  const re = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;
  return re.test(email);
}

/**
 * Builds a single log row array.
 */
function buildLogRow(rowNumber, reason, nome, email, salesRaw) {
  return [new Date(), rowNumber, reason, nome, email, salesRaw];
}

/**
 * Writes the Cleaning_Log header if the sheet is completely empty.
 */
function ensureLogHeader(logSheet) {
  if (logSheet.getLastRow() === 0) {
    logSheet.appendRow(["Timestamp", "Row #", "Reason", "Name", "Email", "Sales (raw)"]);
  }
}
