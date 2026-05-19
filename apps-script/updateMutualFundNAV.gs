/**
 * Robust API Fetcher (Prevents timeouts)
 */
function fetchUrlWithRetry(url) {
  const maxRetries = 3;
  for (let i = 0; i < maxRetries; i++) {
    try {
      const response = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
      if (response.getResponseCode() === 200) {
        return JSON.parse(response.getContentText());
      }
      Utilities.sleep(2000); // Wait 2 seconds and retry
    } catch (e) {
      Utilities.sleep(2000);
    }
  }
  return null; // Return null if all 3 attempts fail
}

/**
 * Updates Mutual Fund NAVs, Injects Formulas, and Formats Dashboard
 */
function updateMutualFundNAV() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("SIP");
  
  if (!sheet) return;

  // 1. Fetch NAVs from API
  const iciciCode = sheet.getRange("O4").getValue();
  const paragCode = sheet.getRange("O5").getValue();
  
  const iciciUrl = "https://api.mfapi.in/mf/" + iciciCode;
  const paragUrl = "https://api.mfapi.in/mf/" + paragCode;

  const iciciData = fetchUrlWithRetry(iciciUrl);
  const paragData = fetchUrlWithRetry(paragUrl);

  // 2. Safety Check & Update
  if (iciciData && paragData && iciciData.data.length > 0 && paragData.data.length > 0) {
    const iciciNAV = parseFloat(iciciData.data[0].nav);
    const paragNAV = parseFloat(paragData.data[0].nav);

    // Update the NAV cells
    sheet.getRange("O7").setValue(iciciNAV);
    sheet.getRange("O9").setValue(paragNAV);

    // --- NEW: AUTOMATICALLY INJECT FORMULAS SO THEY NEVER BREAK ---
    // Calculate individual current values (Units * NAV)
    sheet.getRange("O8").setFormula("=ROUND(K4*O7)"); 
    sheet.getRange("O10").setFormula("=ROUND(K5*O9)"); 

    // Calculate Totals in the Summary block
    sheet.getRange("K8").setFormula("=O8+O10"); // Total Current Value
    sheet.getRange("K9").setFormula("=K8-K3");  // Total Gain / Loss
    sheet.getRange("K10").setFormula("=K9/K3"); // Return %

    // 3. Apply clean formatting to the main dashboard
    sheet.getRange("K8:K9").setNumberFormat("0"); 
    sheet.getRange("K10").setNumberFormat("0.00%"); 
    sheet.getRange("K4:K5").setNumberFormat("0.00");
    
  } else {
    Logger.log("Failed to fetch fresh NAVs. Leaving previous values intact.");
  }
}