/**
 * Google Sheets Data Cleaning (Portfolio Project)
 * Reads data from: Raw_Data
 * Writes cleaned data to: Clean_Data
 * Includes logging, validation, summary stats and UI interaction.
 */


/**
 * Main cleaning logic
 */
function cleanData_withLogging() {

  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const source = ss.getSheetByName("Raw_Data");
  const dest = ss.getSheetByName("Clean_Data") || ss.insertSheet("Clean_Data");

  dest.clearContents();

  const data = source.getDataRange().getValues();

  let totalRows = data.length - 1;
  let keptRows = 0;
  let skippedMissing = 0;
  let skippedInvalid = 0;
  let skippedBelowThreshold = 0;

  const output = [];

  // Copy header
  if (data.length > 0) {
    output.push(data[0]);
  }

  for (let i = 1; i < data.length; i++) {

    const nome = String(data[i]?.[0] ?? "").trim();
    const email = String(data[i]?.[1] ?? "").trim();
    const venditeRaw = data[i]?.[2];

    // Check required fields
    if (!nome || !email) {
      skippedMissing++;
      Logger.log(`Row ${i + 1} skipped: missing Name or Email`);
      continue;
    }

    // Clean numeric value
    const cleaned = String(venditeRaw ?? "")
      .trim()
      .replace(/[^\d.,-]/g, "");

    const num = Number(cleaned.replace(",", "."));

    if (isNaN(num)) {
      skippedInvalid++;
      Logger.log(`Row ${i + 1} skipped: invalid number`);
      continue;
    }

    // Threshold check
    if (num < 100) {
      skippedBelowThreshold++;
      Logger.log(`Row ${i + 1} skipped: below threshold`);
      continue;
    }

    output.push([nome, email, num]);
    keptRows++;
  }

  dest.getRange(1, 1, output.length, output[0].length).setValues(output);

  Logger.log("=== CLEANING SUMMARY ===");
  Logger.log(`Total rows: ${totalRows}`);
  Logger.log(`Valid rows: ${keptRows}`);
  Logger.log(`Missing fields: ${skippedMissing}`);
  Logger.log(`Invalid numbers: ${skippedInvalid}`);
  Logger.log(`Below threshold: ${skippedBelowThreshold}`);

  return {
    totalRows,
    keptRows,
    skippedMissing,
    skippedInvalid,
    skippedBelowThreshold
  };
}


/**
 * UI wrapper with confirmation popup
 */
function runCleaning() {

  const ui = SpreadsheetApp.getUi();

  const response = ui.alert(
    "Run Data Cleaning",
    "Do you want to clean the dataset now?",
    ui.ButtonSet.YES_NO
  );

  if (response !== ui.Button.YES) {
    ui.alert("Operation cancelled.");
    return;
  }

  const result = cleanData_withLogging();

  ui.alert(
    "Cleaning Completed",
    `Total rows: ${result.totalRows}
Valid rows: ${result.keptRows}
Missing fields: ${result.skippedMissing}
Invalid numbers: ${result.skippedInvalid}
Below threshold: ${result.skippedBelowThreshold}`,
    ui.ButtonSet.OK
  );
}


/**
 * Creates custom menu when spreadsheet opens
 */
function onOpen() {

  const ui = SpreadsheetApp.getUi();

  ui.createMenu("Data Tools")
    .addItem("Run Data Cleaning", "runCleaning")
    .addToUi();
}