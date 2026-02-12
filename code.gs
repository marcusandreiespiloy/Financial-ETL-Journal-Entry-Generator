/**
 * FINANCIAL AUTOMATION ENGINE: Double-Entry Journal Entry Generator
 * * ARCHITECTURE: ETL (Extract, Transform, Load)
 * PURPOSE: Transforms raw operational logs into balanced, audit-ready JEs.
 * IMPACT: Optimized for high-volume transactions with built-in data integrity checks.
 * * @author Marcus Andrei Espiloy
 * @version 1.1
 */

function generateJournalEntries() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const logSheet = ss.getSheetByName("Sheet3");
  const tempSheet = ss.getSheetByName("TEMPLATE");
  const ui = SpreadsheetApp.getUi();
  
  // --- DATA INTEGRITY CHECK ---
  if (!logSheet || !tempSheet) {
    ui.alert("CRITICAL ERROR: Source 'Sheet3' or destination 'TEMPLATE' not found.");
    return;
  }

  // --- USER INPUT & PREFIX HANDLING ---
  const jeInput = ui.prompt("Setup", "Enter Starting JE Number", ui.ButtonSet.OK).getResponseText();
  if (!jeInput) return; // Exit if cancelled
  
  let jePrefix = jeInput.replace(/[0-9]/g, '');
  let jeNumber = parseInt(jeInput.replace(/\D/g, '')) || 0;

  const lastRow = logSheet.getLastRow();
  if (lastRow < 4) return;
  
  // Extracting data range (Columns A to W)
  const data = logSheet.getRange(4, 1, lastRow - 3, 23).getValues();
  
  // --- STEP 1: DEEP CLEAR (PREVENT DATA OVERLAP) ---
  const maxRows = tempSheet.getMaxRows();
  if (maxRows >= 3) {
    tempSheet.getRange(3, 1, maxRows - 2, 14).clearContent();
  }
  
  let output = [];

  // --- STEP 2: TRANSFORMATION LOGIC ---
  for (let i = 0; i < data.length; i++) {
    let row = data[i];
    
    // Strict numeric rounding to prevent floating-point imbalance (1-cent errors)
    let vNet = Math.round((parseFloat(row[8]) || 0) * 100) / 100;
    let vVAT = Math.round((parseFloat(row[9]) || 0) * 100) / 100;
    let vEWT = Math.round((parseFloat(row[11]) || 0) * 100) / 100;

    if (vNet > 0) {
      let currentJE = jePrefix + jeNumber;
      let rawTitle = (row[6] || row[14] || "").toString().trim();
      let accTitle = rawTitle;
      let checkTitle = rawTitle.toUpperCase();
      let checkPayee = (row[3] || "").toString().trim().toUpperCase();
      let cvNum = row[2];
      let ceCode = row[4];
      let memo = row[7];
      
      // Date Logic: Mapping based on Account Type
      let transactionDate = (checkTitle.includes("LOANS PAYABLE") || checkTitle.includes("INTEREST EXPENSE")) ? row[0] : row[1];

      // Normalized Account Mapping
      if (checkTitle === "VEHICLE HIRE") accTitle = "Rental Expense";

      // --- FLEXIBLE CREDIT ACCOUNT ROUTING ---
      let creditAcc = "AP - NonTrade"; 
      
      if (checkTitle.includes("ADVANCES") && checkTitle.includes("EMPLOYEE")) {
        creditAcc = "Cash in Bank";
      } 
      else if (checkTitle.includes("PAYABLE") || checkTitle.includes("INTEREST EXPENSE")) {
        if (checkPayee.includes("PBCOM") || checkPayee.includes("PHILIPPINE BANK OF COMMUNICATIONS")) {
          creditAcc = "Cash in Bank (PBCom)";
        } else if (checkPayee.includes("BPI") || checkPayee.includes("BANK OF PHILIPPINE ISLAND")) {
          creditAcc = "Cash in Bank (BPI)";
        } else {
          creditAcc = "Cash in Bank (AUB)";
        }
      }

      // Ensure credit perfectly offsets net debits
      let totalCreditAmount = Math.round((vNet + vVAT - vEWT) * 100) / 100;

      // --- BUFFERING OUTPUT ARRAY (FOR BATCH WRITING) ---
      output.push([null, null, cvNum, null, null, ceCode, null, currentJE, transactionDate, null, accTitle, vNet, 0, memo]);
      
      if (vVAT !== 0) {
        output.push([null, null, cvNum, null, null, ceCode, null, currentJE, transactionDate, null, "VAT input", vVAT, 0, memo]);
      }
      if (vEWT !== 0) {
        output.push([null, null, cvNum, null, null, ceCode, null, currentJE, transactionDate, null, "EWT", 0, vEWT, memo]);
      }
      
      output.push([null, null, cvNum, null, null, ceCode, null, currentJE, transactionDate, null, creditAcc, 0, totalCreditAmount, memo]);
      
      jeNumber++;
    }
  }

  // --- STEP 3: FINAL AUDIT & LOAD ---
  if (output.length > 0) {
    // Architectural Integrity Check: Total Debits must equal Total Credits
    let totalDebits = output.reduce((sum, r) => sum + (r[11] || 0), 0);
    let totalCredits = output.reduce((sum, r) => sum + (r[12] || 0), 0);
    let imbalance = Math.abs(totalDebits - totalCredits);

    tempSheet.getRange(3, 1, output.length, 14).setValues(output);

    if (imbalance > 0.01) {
      ui.alert("WARNING: The generated entries are out of balance by " + imbalance.toFixed(2) + ". Please check source math.");
    } else {
      ui.alert("SUCCESS: " + output.length + " rows generated. Data is balanced and audit-ready.");
    }
  }
}
