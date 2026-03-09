# Google Sheets Data Cleaning (Google Apps Script)

A small portfolio project that demonstrates how to clean, validate, and export spreadsheet data using Google Apps Script.

## Overview
This script reads raw data from a sheet named **Raw_Data**, applies validation + sanitization rules, and outputs clean rows into **Clean_Data**.

It also includes:
- Detailed execution logs (skipped rows + reasons)
- A summary report returned by the cleaning function
- A confirmation popup + final summary popup (UI)
- A custom menu added on spreadsheet open (`onOpen()`)

## Input format (Raw_Data)
Expected columns:
1. **Name**
2. **Email**
3. **Sales**

Examples of messy values the script can handle:
- Missing fields (empty Name/Email)
- Numbers stored as text (e.g., `"150"`)
- Currency symbols (e.g., `"150€"`)
- Extra spaces (e.g., `" 300 "`)

## Cleaning rules
The script:
1. Skips rows where **Name** or **Email** is missing
2. Cleans Sales values by removing non-numeric characters (currency symbols etc.)
3. Converts Sales to a number
4. Keeps only rows with **Sales >= 100**
5. Writes valid rows into **Clean_Data**

## How to use
1. Create a Google Sheet with two tabs:
   - `Raw_Data`
   - `Clean_Data` (optional: script can create it)
2. Go to **Extensions → Apps Script**
3. Paste the code from `src/cleaning.gs`
4. Save and run `runCleaning()` once to grant permissions
5. Reload the spreadsheet to see the custom menu:
   - **Data Tools → Run Data Cleaning**

## Button (optional)
You can add a button inside the sheet:
1. **Insert → Drawing**
2. Add a shape labeled “Run Data Cleaning”
3. Click the drawing → three dots → **Assign script**
4. Enter: `runCleaning`

## Tech notes
- Uses 2D arrays (`getValues()`) for performance (batch read/write)
- Uses optional chaining and nullish coalescing for robustness (`?.` and `??`)
- Logs skip reasons via `Logger.log()`

## License
MIT (optional)
