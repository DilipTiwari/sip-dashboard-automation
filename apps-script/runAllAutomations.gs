function runAllAutomations() {

  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // ==============================
  // 1️⃣ RUN DAILY TRACKER FIRST
  // ==============================

  let marketStatus = "OPEN";

  if (typeof recordDailyTracker === "function") {
    marketStatus = recordDailyTracker();
  }

  // Stop automation if market closed
  if (marketStatus === "CLOSED") {

    console.log("Market closed. Tracker updated only.");

    return;

  }

  try {

    ss.toast("Processing Portfolio Automation...", "System", 2);

    // ==============================
    // 2️⃣ FETCH NIFTY DATA
    // ==============================

    if (typeof fetchNiftyFromGoogleFinance === "function") {
      fetchNiftyFromGoogleFinance();
    }

    // ==============================
    // 3️⃣ UPDATE MUTUAL FUND NAV
    // ==============================

    if (typeof updateMutualFundNAV === "function") {
      updateMutualFundNAV();
    }

    // ==============================
    // 4️⃣ UPDATE CASHFLOW FOR XIRR
    // ==============================

    if (typeof updateCashflowForXIRR === "function") {
      updateCashflowForXIRR();
    }

    // ==============================
    // 5️⃣ SEND PORTFOLIO EMAIL
    // ==============================

    if (typeof sendPortfolioAlert === "function") {
      sendPortfolioAlert();
    }

    // ==============================
    // 6️⃣ SEND TELEGRAM ALERT
    // ==============================

    if (typeof sendTelegramAlert === "function") {
      sendTelegramAlert();
    }

    console.log("Nightly automation completed successfully.");

  }

  catch(err){

    console.error("Automation Error: " + err.message);

  }

}