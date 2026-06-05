/************************************************************
 FINAL MERGED SIP AUTOMATION
 Keep only this one .gs file after testing.

 Required sheets kept:
 1. SIP
 2. SIP Daily Tracker
 3. Graph
 4. SIP Dashboard

 Unused sheets can be deleted by running:
 cleanupUnusedSipSheets()

 Install final triggers by running:
 installFinalSipAutomationTriggers()
************************************************************/


/************************************************************
 SIP FINAL AUTOMATION - DAILY TRACKER + NAV UPDATE
 This section updates only these required sheets:
 SIP
 SIP Daily Tracker
************************************************************/

var SIP_SHEET_NAME_FINAL = "SIP";
var DAILY_TRACKER_SHEET_NAME_FINAL = "SIP Daily Tracker";

var DAILY_TRIGGER_HOUR_FINAL = 6;
var DAILY_TRIGGER_MINUTE_FINAL = 30;

var PARAG_SCHEME_CODE_FINAL = "122639";
var ICICI_SCHEME_CODE_FINAL = "120620";


/******************** NIFTY VALUE ********************/

function updateNiftyValue() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SIP_SHEET_NAME_FINAL);

  if (!sheet) {
    Logger.log("Sheet not found: " + SIP_SHEET_NAME_FINAL);
    return;
  }

  sheet.getRange("O6").setFormula('=GOOGLEFINANCE("INDEXNSE:NIFTY_50","price")');
  Logger.log("NIFTY formula updated in SIP!O6.");
}


/******************** MUTUAL FUND NAV UPDATE ********************/

function fetchUrlWithRetry(url) {
  var maxRetries = 3;

  for (var i = 0; i < maxRetries; i++) {
    try {
      var response = UrlFetchApp.fetch(url, { muteHttpExceptions: true });

      if (response.getResponseCode() === 200) {
        return JSON.parse(response.getContentText());
      }

      Utilities.sleep(2000);
    } catch (e) {
      Utilities.sleep(2000);
    }
  }

  return null;
}


function updateMutualFundNAV() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SIP_SHEET_NAME_FINAL);

  if (!sheet) {
    Logger.log("Sheet not found: " + SIP_SHEET_NAME_FINAL);
    return;
  }

  var iciciCode = String(sheet.getRange("O4").getValue() || ICICI_SCHEME_CODE_FINAL).trim();
  var paragCode = String(sheet.getRange("O5").getValue() || PARAG_SCHEME_CODE_FINAL).trim();

  var iciciNAV = 0;
  var paragNAV = 0;

  // First try mfapi.in
  try {
    var iciciData = fetchUrlWithRetry("https://api.mfapi.in/mf/" + iciciCode);
    var paragData = fetchUrlWithRetry("https://api.mfapi.in/mf/" + paragCode);

    if (iciciData && iciciData.data && iciciData.data.length > 0) {
      iciciNAV = Number(iciciData.data[0].nav) || 0;
    }

    if (paragData && paragData.data && paragData.data.length > 0) {
      paragNAV = Number(paragData.data[0].nav) || 0;
    }
  } catch (e) {
    Logger.log("mfapi NAV fetch failed: " + e);
  }

  // AMFI fallback
  if (!iciciNAV || !paragNAV) {
    var amfi = FINAL_fetchLatestNavsFromAMFI_();

    if (!iciciNAV) iciciNAV = amfi.icici || 0;
    if (!paragNAV) paragNAV = amfi.parag || 0;
  }

  if (iciciNAV) {
    sheet.getRange("O7").setValue(iciciNAV);
  }

  if (paragNAV) {
    sheet.getRange("O9").setValue(paragNAV);
  }

  // Keep formulas alive
  sheet.getRange("O8").setFormula("=ROUND(K4*O7,2)");
  sheet.getRange("O10").setFormula("=ROUND(K5*O9,2)");
  sheet.getRange("K8").setFormula("=O8+O10");
  sheet.getRange("K9").setFormula("=K8-K3");
  sheet.getRange("K10").setFormula("=K9/K3");

  sheet.getRange("K8:K9").setNumberFormat("₹#,##0.00");
  sheet.getRange("K10").setNumberFormat("0.00%");
  sheet.getRange("K4:K5").setNumberFormat("0.000");
  sheet.getRange("O7:O10").setNumberFormat("0.0000");

  SpreadsheetApp.flush();

  Logger.log("NAV update completed. ICICI NAV: " + iciciNAV + ", Parag NAV: " + paragNAV);
}


/******************** DAILY TRACKER A:S ********************/

function FIX_recordDailyTracker_A_to_S() {
  updateNiftyValue();
  updateMutualFundNAV();

  SpreadsheetApp.flush();
  Utilities.sleep(2000);

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sipSheet = ss.getSheetByName(SIP_SHEET_NAME_FINAL);
  var tracker = FINAL_getOrCreateSheet_(ss, DAILY_TRACKER_SHEET_NAME_FINAL);

  if (!sipSheet) {
    throw new Error("Sheet not found: " + SIP_SHEET_NAME_FINAL);
  }

  FINAL_ensureDailyTrackerHeader_(tracker);
  FINAL_deleteBadDailyTrackerRows_(tracker);

  var today = FINAL_dateOnly_(new Date());

  var iciciUnits = FINAL_num_(sipSheet.getRange("K4").getValue());
  var paragUnits = FINAL_num_(sipSheet.getRange("K5").getValue());

  var iciciNAV = FINAL_num_(sipSheet.getRange("O7").getValue());
  var paragNAV = FINAL_num_(sipSheet.getRange("O9").getValue());

  var iciciInvested = FINAL_sumRange_(sipSheet, "B2:B1000");
  var paragInvested = FINAL_sumRange_(sipSheet, "C2:C1000");

  var latestFallback = FINAL_getLatestValidDailyRow_(tracker);

  if ((!iciciUnits || !paragUnits || !iciciInvested || !paragInvested) && latestFallback) {
    var fallbackRow = latestFallback.values;

    if (!paragUnits) paragUnits = FINAL_num_(fallbackRow[7]);
    if (!paragNAV) paragNAV = FINAL_num_(fallbackRow[8]);
    if (!paragInvested) paragInvested = FINAL_num_(fallbackRow[9]);

    if (!iciciUnits) iciciUnits = FINAL_num_(fallbackRow[13]);
    if (!iciciNAV) iciciNAV = FINAL_num_(fallbackRow[14]);
    if (!iciciInvested) iciciInvested = FINAL_num_(fallbackRow[15]);
  }

  if (!iciciNAV || !paragNAV) {
    var amfi = FINAL_fetchLatestNavsFromAMFI_();
    if (!iciciNAV) iciciNAV = amfi.icici || 0;
    if (!paragNAV) paragNAV = amfi.parag || 0;
  }

  if (paragUnits <= 0 || paragInvested <= 0 || iciciUnits <= 0 || iciciInvested <= 0) {
    throw new Error("Invalid SIP values. Please check SIP sheet K4/K5 and B:C investment data.");
  }

  var paragCurrent = paragUnits * paragNAV;
  var paragPL = paragCurrent - paragInvested;
  var paragReturn = paragInvested ? paragPL / paragInvested : 0;

  var iciciCurrent = iciciUnits * iciciNAV;
  var iciciPL = iciciCurrent - iciciInvested;
  var iciciReturn = iciciInvested ? iciciPL / iciciInvested : 0;

  var totalInvested = paragInvested + iciciInvested;
  var currentValue = paragCurrent + iciciCurrent;
  var totalPL = currentValue - totalInvested;
  var totalReturn = totalInvested ? totalPL / totalInvested : 0;

  var rowValues = [
    today,
    FINAL_dayName_(today),
    FINAL_weekNumber_(today),

    FINAL_round2_(totalInvested),
    FINAL_round2_(currentValue),
    FINAL_round2_(totalPL),
    totalReturn,

    FINAL_round3_(paragUnits),
    FINAL_round4_(paragNAV),
    FINAL_round2_(paragInvested),
    FINAL_round2_(paragCurrent),
    FINAL_round2_(paragPL),
    paragReturn,

    FINAL_round3_(iciciUnits),
    FINAL_round4_(iciciNAV),
    FINAL_round2_(iciciInvested),
    FINAL_round2_(iciciCurrent),
    FINAL_round2_(iciciPL),
    iciciReturn
  ];

  var targetRow = FINAL_findDailyRowByDate_(tracker, today);

  if (!targetRow) {
    var latest = FINAL_getLatestValidDailyRow_(tracker);
    targetRow = latest ? latest.rowNumber + 1 : 2;
  }

  tracker.getRange(targetRow, 1, 1, 19).setValues([rowValues]);
  FINAL_formatDailyTrackerRow_(tracker, targetRow);

  SpreadsheetApp.flush();

  Logger.log("SIP Daily Tracker A:S updated at row: " + targetRow);
  return "OPEN";
}


// Backward-compatible alias. Old code can call recordDailyTracker safely now.
function recordDailyTracker() {
  return FIX_recordDailyTracker_A_to_S();
}


/******************** DAILY TRACKER HELPERS ********************/

function FINAL_ensureDailyTrackerHeader_(sheet) {
  var headers = [
    "Date",
    "Day",
    "Week",
    "Total Invested",
    "Current Value",
    "Total P/L",
    "Total Return %",
    "Parag Units",
    "Parag NAV",
    "Parag Invested",
    "Parag Current",
    "Parag P/L",
    "Parag Return %",
    "ICICI Units",
    "ICICI NAV",
    "ICICI Invested",
    "ICICI Current",
    "ICICI P/L",
    "ICICI Return %"
  ];

  sheet.getRange(1, 1, 1, 19).setValues([headers]);
  sheet.getRange(1, 1, 1, 19)
    .setFontWeight("bold")
    .setBackground("#d9e8fb")
    .setHorizontalAlignment("center")
    .setBorder(true, true, true, true, true, true);
}


function FINAL_deleteBadDailyTrackerRows_(sheet) {
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return;

  var rows = sheet.getRange(2, 1, lastRow - 1, 19).getDisplayValues();

  for (var i = rows.length - 1; i >= 0; i--) {
    var rowNumber = i + 2;
    var row = rows[i];

    var hasData = row.some(function(cell) {
      return String(cell || "").trim() !== "";
    });

    if (hasData && !FINAL_isValidDailyDisplayRow_(row)) {
      sheet.deleteRow(rowNumber);
    }
  }
}


function FINAL_getLatestValidDailyRow_(sheet) {
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return null;

  var values = sheet.getRange(2, 1, lastRow - 1, 19).getValues();
  var displayValues = sheet.getRange(2, 1, lastRow - 1, 19).getDisplayValues();

  for (var i = values.length - 1; i >= 0; i--) {
    if (FINAL_isValidDailyDisplayRow_(displayValues[i])) {
      return {
        rowNumber: i + 2,
        values: values[i],
        displayValues: displayValues[i]
      };
    }
  }

  return null;
}


function FINAL_findDailyRowByDate_(sheet, targetDate) {
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return null;

  var rows = sheet.getRange(2, 1, lastRow - 1, 19).getDisplayValues();

  for (var i = rows.length - 1; i >= 0; i--) {
    if (!FINAL_isValidDailyDisplayRow_(rows[i])) continue;

    var rowDate = FINAL_parseDate_(rows[i][0]);

    if (rowDate && rowDate.getTime() === targetDate.getTime()) {
      return i + 2;
    }
  }

  return null;
}


function FINAL_isValidDailyDisplayRow_(row) {
  var date = FINAL_parseDate_(row[0]);
  var day = String(row[1] || "").trim();
  var week = FINAL_num_(row[2]);
  var totalInvested = FINAL_num_(row[3]);
  var currentValue = FINAL_num_(row[4]);

  var paragUnits = FINAL_num_(row[7]);
  var paragInvested = FINAL_num_(row[9]);
  var iciciUnits = FINAL_num_(row[13]);
  var iciciInvested = FINAL_num_(row[15]);

  var validDays = ["Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday", "Sunday"];

  return (
    date &&
    validDays.indexOf(day) !== -1 &&
    week > 0 &&
    totalInvested > 0 &&
    Math.abs(currentValue) > 1 &&
    paragUnits > 0 &&
    paragInvested > 0 &&
    iciciUnits > 0 &&
    iciciInvested > 0
  );
}


function FINAL_formatDailyTrackerRow_(sheet, rowNumber) {
  sheet.getRange(rowNumber, 1).setNumberFormat("dd-mmm-yyyy");

  sheet.getRange(rowNumber, 4, 1, 2).setNumberFormat("₹#,##0.00");
  sheet.getRange(rowNumber, 6).setNumberFormat("₹#,##0.00");
  sheet.getRange(rowNumber, 7).setNumberFormat("0.00%");

  sheet.getRange(rowNumber, 8).setNumberFormat("0.000");
  sheet.getRange(rowNumber, 9).setNumberFormat("0.0000");
  sheet.getRange(rowNumber, 10, 1, 3).setNumberFormat("₹#,##0.00");
  sheet.getRange(rowNumber, 13).setNumberFormat("0.00%");

  sheet.getRange(rowNumber, 14).setNumberFormat("0.000");
  sheet.getRange(rowNumber, 15).setNumberFormat("0.0000");
  sheet.getRange(rowNumber, 16, 1, 3).setNumberFormat("₹#,##0.00");
  sheet.getRange(rowNumber, 19).setNumberFormat("0.00%");

  sheet.getRange(rowNumber, 1, 1, 19)
    .setHorizontalAlignment("center")
    .setBorder(true, true, true, true, true, true);
}


function FINAL_fetchLatestNavsFromAMFI_() {
  var result = { parag: 0, icici: 0 };

  try {
    var response = UrlFetchApp.fetch("https://www.amfiindia.com/spages/NAVAll.txt", {
      muteHttpExceptions: true
    });

    if (response.getResponseCode() !== 200) {
      return result;
    }

    var text = response.getContentText();
    result.parag = FINAL_findAMFINavBySchemeCode_(text, PARAG_SCHEME_CODE_FINAL);
    result.icici = FINAL_findAMFINavBySchemeCode_(text, ICICI_SCHEME_CODE_FINAL);

    if (!result.parag) {
      result.parag = FINAL_findAMFINavByKeywords_(text, ["parag", "flexi cap", "direct", "growth"]);
    }

    if (!result.icici) {
      result.icici = FINAL_findAMFINavByKeywords_(text, ["icici prudential", "nifty 50", "direct", "growth"]);
    }
  } catch (e) {
    Logger.log("AMFI fetch failed: " + e);
  }

  return result;
}


function FINAL_findAMFINavBySchemeCode_(text, schemeCode) {
  var lines = text.split(/\r?\n/);

  for (var i = 0; i < lines.length; i++) {
    var parts = lines[i].split(";");

    if (parts.length >= 6 && String(parts[0]).trim() === String(schemeCode)) {
      return FINAL_num_(parts[4]);
    }
  }

  return 0;
}


function FINAL_findAMFINavByKeywords_(text, keywords) {
  var lines = text.split(/\r?\n/);

  for (var i = 0; i < lines.length; i++) {
    var parts = lines[i].split(";");

    if (parts.length < 6) continue;

    var name = String(parts[3] || "").toLowerCase();

    var match = keywords.every(function(k) {
      return name.indexOf(String(k).toLowerCase()) !== -1;
    });

    if (match) {
      return FINAL_num_(parts[4]);
    }
  }

  return 0;
}


function FINAL_sumRange_(sheet, rangeA1) {
  var values = sheet.getRange(rangeA1).getValues();
  var sum = 0;

  values.forEach(function(row) {
    row.forEach(function(v) {
      sum += FINAL_num_(v);
    });
  });

  return sum;
}


function FINAL_getOrCreateSheet_(ss, name) {
  return ss.getSheetByName(name) || ss.insertSheet(name);
}


function FINAL_num_(value) {
  if (typeof value === "number") return value;
  if (value === null || value === undefined || value === "") return 0;

  var cleaned = String(value)
    .replace(/₹/g, "")
    .replace(/,/g, "")
    .replace(/%/g, "")
    .replace(/−/g, "-")
    .replace(/\(/g, "-")
    .replace(/\)/g, "")
    .replace(/\s/g, "")
    .trim();

  var n = Number(cleaned);
  return isNaN(n) ? 0 : n;
}


function FINAL_dateOnly_(date) {
  return new Date(date.getFullYear(), date.getMonth(), date.getDate());
}


function FINAL_parseDate_(value) {
  if (value instanceof Date) {
    return FINAL_dateOnly_(value);
  }

  if (!value) return null;

  var text = String(value).trim();
  var normalDate = new Date(text);

  if (!isNaN(normalDate.getTime())) {
    return FINAL_dateOnly_(normalDate);
  }

  var parts = text.split("-");

  if (parts.length === 3) {
    var day = Number(parts[0]);
    var monthText = String(parts[1]).toLowerCase();
    var year = Number(parts[2]);

    if (year < 100) year = 2000 + year;

    var months = {
      jan: 0, january: 0,
      feb: 1, february: 1,
      mar: 2, march: 2,
      apr: 3, april: 3,
      may: 4,
      jun: 5, june: 5,
      jul: 6, july: 6,
      aug: 7, august: 7,
      sep: 8, sept: 8, september: 8,
      oct: 9, october: 9,
      nov: 10, november: 10,
      dec: 11, december: 11
    };

    var month = isNaN(Number(monthText)) ? months[monthText] : Number(monthText) - 1;

    if (!isNaN(day) && month !== undefined && !isNaN(year)) {
      return new Date(year, month, day);
    }
  }

  return null;
}


function FINAL_dayName_(date) {
  return Utilities.formatDate(date, Session.getScriptTimeZone(), "EEEE");
}


function FINAL_weekNumber_(date) {
  var d = new Date(Date.UTC(date.getFullYear(), date.getMonth(), date.getDate()));
  var dayNum = d.getUTCDay() || 7;

  d.setUTCDate(d.getUTCDate() + 4 - dayNum);

  var yearStart = new Date(Date.UTC(d.getUTCFullYear(), 0, 1));

  return Math.ceil((((d - yearStart) / 86400000) + 1) / 7);
}


function FINAL_round2_(value) {
  return Math.round((Number(value) || 0) * 100) / 100;
}


function FINAL_round3_(value) {
  return Math.round((Number(value) || 0) * 1000) / 1000;
}


function FINAL_round4_(value) {
  return Math.round((Number(value) || 0) * 10000) / 10000;
}


/************************************************************
 DAILY PORTFOLIO MAIL - FIXED VERSION
 Reads correct data from SIP Daily Tracker columns A:S only.
 Does NOT overwrite SIP Daily Tracker headers.
 Does NOT save old 9-column tracker data.
************************************************************/


/******************** CONFIG ********************/

var MAIL_TO = "diliptiwari1dkt@gmail.com";
var MAIL_BCC = "diliptiwari12dkt@gmail.com,mehtrekrishna1996@gmail.com,shivani93808@gmail.com";

var MAIL_TRACKER_SHEET_NAME = "SIP Daily Tracker";


/******************** MAIN DAILY AUTOMATION ********************/

function morningSipAutomation() {
  // First update the Daily Tracker full A:S row using your fixed tracker function.
  // This function already worked for you.
  if (typeof FIX_recordDailyTracker_A_to_S === "function") {
    FIX_recordDailyTracker_A_to_S();
    SpreadsheetApp.flush();
    Utilities.sleep(5000);
  }

  // Then send mail using the latest valid A:S row.
  sendPortfolioAlert();
}


/******************** MAIN MAIL FUNCTION ********************/

function sendPortfolioAlert() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(MAIL_TRACKER_SHEET_NAME);

  if (!sheet) {
    throw new Error("Sheet not found: " + MAIL_TRACKER_SHEET_NAME);
  }

  var latest = MAIL_getLatestValidDailyRow_(sheet);

  if (!latest) {
    throw new Error("No valid A:S row found in SIP Daily Tracker.");
  }

  var previous = MAIL_getPreviousValidDailyRow_(sheet, latest.rowNumber);

  var latestObj = MAIL_rowToObject_(latest.values, latest.displayValues);
  var previousObj = previous ? MAIL_rowToObject_(previous.values, previous.displayValues) : null;

  var oneDayTotalChange = previousObj ? latestObj.currentValue - previousObj.currentValue : 0;
  var oneDayTotalChangePct = previousObj && previousObj.currentValue
    ? oneDayTotalChange / previousObj.currentValue
    : 0;

  var paragDayChange = previousObj ? latestObj.paragCurrent - previousObj.paragCurrent : 0;
  var paragDayChangePct = previousObj && previousObj.paragCurrent
    ? paragDayChange / previousObj.paragCurrent
    : 0;

  var iciciDayChange = previousObj ? latestObj.iciciCurrent - previousObj.iciciCurrent : 0;
  var iciciDayChangePct = previousObj && previousObj.iciciCurrent
    ? iciciDayChange / previousObj.iciciCurrent
    : 0;

  var subject = "📊 Portfolio Update (" + MAIL_formatPercent_(latestObj.totalReturnPct) + ")";

  var status = latestObj.totalPL >= 0 ? "PROFIT" : "LOSS";

  var htmlBody = MAIL_buildPortfolioHtml_(
    latestObj,
    status,
    oneDayTotalChange,
    oneDayTotalChangePct,
    paragDayChange,
    paragDayChangePct,
    iciciDayChange,
    iciciDayChangePct
  );

  MailApp.sendEmail({
    to: MAIL_TO,
    bcc: MAIL_BCC,
    subject: subject,
    htmlBody: htmlBody
  });

  Logger.log("Daily portfolio mail sent successfully.");
  Logger.log("Source row used: " + latest.rowNumber);
  Logger.log("Mail subject: " + subject);
}


// Alias, in case any old script/trigger calls sendDailyPortfolioMail
function sendDailyPortfolioMail() {
  sendPortfolioAlert();
}


/******************** GET LATEST VALID A:S ROW ********************/

function MAIL_getLatestValidDailyRow_(sheet) {
  var lastRow = sheet.getLastRow();

  if (lastRow < 2) return null;

  var values = sheet.getRange(2, 1, lastRow - 1, 19).getValues();
  var displayValues = sheet.getRange(2, 1, lastRow - 1, 19).getDisplayValues();

  for (var i = values.length - 1; i >= 0; i--) {
    if (MAIL_isValidDailyDisplayRow_(displayValues[i])) {
      return {
        rowNumber: i + 2,
        values: values[i],
        displayValues: displayValues[i]
      };
    }
  }

  return null;
}


function MAIL_getPreviousValidDailyRow_(sheet, latestRowNumber) {
  if (latestRowNumber <= 2) return null;

  var values = sheet.getRange(2, 1, latestRowNumber - 2, 19).getValues();
  var displayValues = sheet.getRange(2, 1, latestRowNumber - 2, 19).getDisplayValues();

  for (var i = values.length - 1; i >= 0; i--) {
    if (MAIL_isValidDailyDisplayRow_(displayValues[i])) {
      return {
        rowNumber: i + 2,
        values: values[i],
        displayValues: displayValues[i]
      };
    }
  }

  return null;
}


function MAIL_isValidDailyDisplayRow_(row) {
  var date = MAIL_parseDate_(row[0]);
  var day = String(row[1] || "").trim();
  var week = MAIL_num_(row[2]);

  var totalInvested = MAIL_num_(row[3]);
  var currentValue = MAIL_num_(row[4]);

  var paragUnits = MAIL_num_(row[7]);
  var paragInvested = MAIL_num_(row[9]);

  var iciciUnits = MAIL_num_(row[13]);
  var iciciInvested = MAIL_num_(row[15]);

  var validDays = [
    "Monday",
    "Tuesday",
    "Wednesday",
    "Thursday",
    "Friday",
    "Saturday",
    "Sunday"
  ];

  return (
    date &&
    validDays.indexOf(day) !== -1 &&
    week > 0 &&
    totalInvested > 0 &&
    Math.abs(currentValue) > 1 &&
    paragUnits > 0 &&
    paragInvested > 0 &&
    iciciUnits > 0 &&
    iciciInvested > 0
  );
}


/******************** CONVERT A:S ROW TO OBJECT ********************/

function MAIL_rowToObject_(row, displayRow) {
  return {
    date: MAIL_parseDate_(row[0]) || MAIL_parseDate_(displayRow[0]),
    day: displayRow[1],
    week: MAIL_num_(row[2]),

    totalInvested: MAIL_num_(row[3]),
    currentValue: MAIL_num_(row[4]),
    totalPL: MAIL_num_(row[5]),
    totalReturnPct: MAIL_percentDecimal_(row[6], displayRow[6]),

    paragUnits: MAIL_num_(row[7]),
    paragNAV: MAIL_num_(row[8]),
    paragInvested: MAIL_num_(row[9]),
    paragCurrent: MAIL_num_(row[10]),
    paragPL: MAIL_num_(row[11]),
    paragReturnPct: MAIL_percentDecimal_(row[12], displayRow[12]),

    iciciUnits: MAIL_num_(row[13]),
    iciciNAV: MAIL_num_(row[14]),
    iciciInvested: MAIL_num_(row[15]),
    iciciCurrent: MAIL_num_(row[16]),
    iciciPL: MAIL_num_(row[17]),
    iciciReturnPct: MAIL_percentDecimal_(row[18], displayRow[18])
  };
}


/******************** EMAIL HTML ********************/

function MAIL_buildPortfolioHtml_(
  latestObj,
  status,
  oneDayTotalChange,
  oneDayTotalChangePct,
  paragDayChange,
  paragDayChangePct,
  iciciDayChange,
  iciciDayChangePct
) {
  var green = "#16a34a";
  var red = "#dc2626";
  var black = "#111827";
  var muted = "#6b7280";

  return `
  <html>
  <body style="font-family:Arial;background:#f5f6f8;padding:20px">

  <div style="max-width:760px;margin:auto;background:white;padding:22px;border-radius:14px;box-shadow:0 4px 12px rgba(0,0,0,0.08)">

    <h2 style="margin-bottom:10px">Investments (2)</h2>

    <p style="color:${muted};margin:0">Current value</p>
    <h1 style="margin:5px 0 15px 0">${MAIL_formatMoney_(latestObj.currentValue)}</h1>

    <table style="width:100%;margin-bottom:15px">
      <tr style="color:${muted};font-size:13px">
        <td>Invested value</td>
        <td>1D returns</td>
        <td>Total returns</td>
        <td>Total Return %</td>
      </tr>

      <tr style="font-weight:bold">
        <td>${MAIL_formatMoney_(latestObj.totalInvested)}</td>

        <td style="color:${MAIL_color_(oneDayTotalChange)}">
          ${MAIL_formatSignedMoney_(oneDayTotalChange)}
          <br><span style="font-size:12px">(${MAIL_formatSignedPercent_(oneDayTotalChangePct)})</span>
        </td>

        <td style="color:${MAIL_color_(latestObj.totalPL)}">
          ${MAIL_formatSignedMoney_(latestObj.totalPL)}
          <br><span style="font-size:12px">(${MAIL_formatSignedPercent_(latestObj.totalReturnPct)})</span>
        </td>

        <td style="color:${MAIL_color_(latestObj.totalReturnPct)}">
          ${MAIL_formatPercent_(latestObj.totalReturnPct)}
        </td>
      </tr>
    </table>

    <hr style="border:none;border-top:1px solid #e5e7eb;margin:15px 0">

    <table style="width:100%;border-collapse:collapse">

      <tr style="color:${muted};font-size:13px;text-align:left">
        <th style="padding-bottom:8px">Scheme name</th>
        <th style="padding-bottom:8px;text-align:right">Day change</th>
        <th style="padding-bottom:8px;text-align:right">Returns</th>
        <th style="padding-bottom:8px;text-align:right">Current (Invested)</th>
      </tr>

      <tr style="border-top:1px solid #f1f5f9">
        <td style="padding:12px 0">
          <div style="font-weight:600;color:${black}">
            Parag Parikh Flexi Cap Fund Direct Growth
          </div>
          <div style="font-size:12px;color:${muted}">
            Units: ${latestObj.paragUnits.toFixed(3)}
          </div>
          <div style="font-size:12px;color:${muted}">
            NAV: ${latestObj.paragNAV.toFixed(4)}
          </div>
        </td>

        <td style="color:${MAIL_color_(paragDayChange)};font-weight:bold;text-align:right">
          ${MAIL_formatSignedMoney_(paragDayChange)}
          <br><span style="font-size:12px">(${MAIL_formatSignedPercent_(paragDayChangePct)})</span>
        </td>

        <td style="color:${MAIL_color_(latestObj.paragPL)};font-weight:bold;text-align:right">
          ${MAIL_formatSignedMoney_(latestObj.paragPL)}
          <br><span style="font-size:12px">(${MAIL_formatSignedPercent_(latestObj.paragReturnPct)})</span>
        </td>

        <td style="text-align:right">
          <b>${MAIL_formatMoney_(latestObj.paragCurrent)}</b><br>
          <span style="font-size:12px;color:${muted}">
            ${MAIL_formatMoney_(latestObj.paragInvested)}
          </span>
        </td>
      </tr>

      <tr style="border-top:1px solid #f1f5f9">
        <td style="padding:12px 0">
          <div style="font-weight:600;color:${black}">
            ICICI Prudential Nifty 50 Index Direct Plan Growth
          </div>
          <div style="font-size:12px;color:${muted}">
            Units: ${latestObj.iciciUnits.toFixed(3)}
          </div>
          <div style="font-size:12px;color:${muted}">
            NAV: ${latestObj.iciciNAV.toFixed(4)}
          </div>
        </td>

        <td style="color:${MAIL_color_(iciciDayChange)};font-weight:bold;text-align:right">
          ${MAIL_formatSignedMoney_(iciciDayChange)}
          <br><span style="font-size:12px">(${MAIL_formatSignedPercent_(iciciDayChangePct)})</span>
        </td>

        <td style="color:${MAIL_color_(latestObj.iciciPL)};font-weight:bold;text-align:right">
          ${MAIL_formatSignedMoney_(latestObj.iciciPL)}
          <br><span style="font-size:12px">(${MAIL_formatSignedPercent_(latestObj.iciciReturnPct)})</span>
        </td>

        <td style="text-align:right">
          <b>${MAIL_formatMoney_(latestObj.iciciCurrent)}</b><br>
          <span style="font-size:12px;color:${muted}">
            ${MAIL_formatMoney_(latestObj.iciciInvested)}
          </span>
        </td>
      </tr>

    </table>

    <p style="margin-top:20px;color:${muted};font-size:13px">
      Status: <b>${status}</b><br>
      Last Updated: ${MAIL_formatDate_(latestObj.date)}
    </p>

  </div>

  </body>
  </html>
  `;
}


/******************** HELPERS ********************/

function MAIL_num_(value) {
  if (value === null || value === "") return 0;

  if (typeof value === "number") return value;

  var text = String(value)
    .replace(/₹/g, "")
    .replace(/,/g, "")
    .replace(/%/g, "")
    .replace(/\(/g, "-")
    .replace(/\)/g, "")
    .trim();

  var n = Number(text);

  return isNaN(n) ? 0 : n;
}


function MAIL_percentDecimal_(rawValue, displayValue) {
  var displayText = String(displayValue || "").trim();

  if (displayText.indexOf("%") !== -1) {
    return MAIL_num_(displayText) / 100;
  }

  var n = MAIL_num_(rawValue);

  if (Math.abs(n) > 1) {
    return n / 100;
  }

  return n;
}


function MAIL_parseDate_(value) {
  if (value instanceof Date) {
    return new Date(value.getFullYear(), value.getMonth(), value.getDate());
  }

  if (!value) return null;

  var text = String(value).trim();

  var parts = text.split("-");

  if (parts.length === 3) {
    var day = Number(parts[0]);
    var monthText = String(parts[1]).toLowerCase();
    var year = Number(parts[2]);

    if (year < 100) year = 2000 + year;

    var months = {
      jan: 0,
      january: 0,
      feb: 1,
      february: 1,
      mar: 2,
      march: 2,
      apr: 3,
      april: 3,
      may: 4,
      jun: 5,
      june: 5,
      jul: 6,
      july: 6,
      aug: 7,
      august: 7,
      sep: 8,
      sept: 8,
      september: 8,
      oct: 9,
      october: 9,
      nov: 10,
      november: 10,
      dec: 11,
      december: 11
    };

    var month;

    if (isNaN(Number(monthText))) {
      month = months[monthText];
    } else {
      month = Number(monthText) - 1;
    }

    if (!isNaN(day) && month !== undefined && !isNaN(year)) {
      return new Date(year, month, day);
    }
  }

  var d = new Date(text);

  if (!isNaN(d.getTime())) {
    return new Date(d.getFullYear(), d.getMonth(), d.getDate());
  }

  return null;
}


function MAIL_formatMoney_(amount) {
  var n = Number(amount) || 0;

  return "₹" + Math.abs(n).toLocaleString("en-IN", {
    minimumFractionDigits: 2,
    maximumFractionDigits: 2
  });
}


function MAIL_formatSignedMoney_(amount) {
  var n = Number(amount) || 0;
  var sign = n < 0 ? "-₹" : "₹";

  return sign + Math.abs(n).toLocaleString("en-IN", {
    minimumFractionDigits: 2,
    maximumFractionDigits: 2
  });
}


function MAIL_formatPercent_(decimalValue) {
  var n = Number(decimalValue) || 0;

  return (n * 100).toFixed(2) + "%";
}


function MAIL_formatSignedPercent_(decimalValue) {
  var n = Number(decimalValue) || 0;
  var sign = n > 0 ? "+" : n < 0 ? "-" : "";

  return sign + Math.abs(n * 100).toFixed(2) + "%";
}


function MAIL_color_(value) {
  return Number(value) >= 0 ? "#16a34a" : "#dc2626";
}


function MAIL_formatDate_(date) {
  return Utilities.formatDate(date, Session.getScriptTimeZone(), "dd-MMM-yyyy");
}

/******************** CONFIG ********************/
var SD_DASHBOARD_SHEET = "SIP Dashboard";
var SD_TRACKER_SHEET = "SIP Daily Tracker";

var SD_SUNDAY_EMAIL_HOUR = 9;

var SD_MAIL_TO = "diliptiwari1dkt@gmail.com";
var SD_MAIL_BCC = "diliptiwari12dkt@gmail.com,mehtrekrishna1996@gmail.com,shivani93808@gmail.com";

/******************** MONTHLY SIP AUTOMATION CONFIG ********************/
var SD_SIP_START_YEAR = 2025;
var SD_SIP_START_MONTH_INDEX = 9; // Oct = 9, because JavaScript month starts from 0
var SD_MONTHLY_SIP_AMOUNT = 5000;
var SD_SIP_TOLERANCE = 20;

/******************** MENU ********************/
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("SIP Dashboard")
    .addItem("Refresh Dashboard", "refreshDashboard")
    .addItem("Send Dashboard Mail Now", "sendWeeklyDashboardMail")
    .addItem("Install Sunday Dashboard Mail Trigger", "installSundayDashboardMailTrigger")
    .addToUi();
}

/******************** MAIN DASHBOARD ********************/
function setupUltraDashboard(silent) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var dashboard = sdGetOrCreateSheet_(ss, SD_DASHBOARD_SHEET);
  var tracker = ss.getSheetByName(SD_TRACKER_SHEET);

  if (!tracker) {
    throw new Error("Sheet not found: " + SD_TRACKER_SHEET);
  }

  var allRows = sdGetCleanTrackerRows_(tracker);

  if (!allRows.length) {
    throw new Error("No valid data found in '" + SD_TRACKER_SHEET + "'.");
  }

  var latest = allRows[allRows.length - 1];

  var totalInvested = latest[1];
  var currentValue = latest[2];
  var overallPnL = latest[3];
  var returnPct = latest[4];

  sdPrepareDashboardSheet_(dashboard);

  dashboard.getRange("B2:X4")
    .merge()
    .setValue("📊 PERSONAL SIP INVESTMENT DASHBOARD")
    .setBackground("#08152f")
    .setFontColor("white")
    .setFontWeight("bold")
    .setFontSize(13)
    .setHorizontalAlignment("center")
    .setVerticalAlignment("middle");

  dashboard.getRange("B5:X5")
    .merge()
    .setValue(
      "Updated : " +
      Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "dd-MMM-yyyy hh:mm a")
    )
    .setBackground("#1f2937")
    .setFontColor("white")
    .setFontSize(8)
    .setHorizontalAlignment("center");

  sdCreateCard_(dashboard, "B7:F10", "💰 Total Invested", totalInvested, "#2563eb", "currency");
  sdCreateCard_(dashboard, "H7:L10", "📈 Current Value", currentValue, "#7c3aed", "currency");
  sdCreateCard_(dashboard, "N7:R10", "📉 Overall P/L", overallPnL, overallPnL >= 0 ? "#059669" : "#dc2626", "currency");
  sdCreateCard_(dashboard, "T7:X10", "🎯 Return %", returnPct, "#059669", "percent");

  dashboard.getRange("B12:X14")
    .merge()
    .setValue(overallPnL >= 0 ? "🟢 Portfolio Performing Well" : "🔴 Portfolio Under Recovery")
    .setBackground("#f3f4f6")
    .setFontWeight("bold")
    .setFontSize(10)
    .setHorizontalAlignment("center")
    .setVerticalAlignment("middle")
    .setBorder(true, true, true, true, true, true);

  sdCreateSection_(dashboard, "B16:L17", "📊 All Time P/L");
  sdCreateSection_(dashboard, "N16:X17", "📈 All Time Investment Growth");

  sdCreateSection_(dashboard, "B32:L33", "📅 Monthly SIP Tracker");
  sdCreateSection_(dashboard, "N32:X33", "🥧 Portfolio Allocation");

  sdCreatePanel_(dashboard, "B18:L29");
  sdCreatePanel_(dashboard, "N18:X29");
  sdCreatePanel_(dashboard, "N34:X43");

  var lastSipRow = sdCreateSipTable_(dashboard);

  SpreadsheetApp.flush();

  var pnlBlob = sdCreatePnLChartBlob_(allRows, 540, 230);
  var growthBlob = sdCreateGrowthChartBlob_(allRows, 540, 230);
  var pieBlob = sdCreatePieChartBlob_(520, 200);

  dashboard.insertImage(pnlBlob, 2, 18, 6, 6);
  dashboard.insertImage(growthBlob, 14, 18, 6, 6);
  dashboard.insertImage(pieBlob, 14, 34, 6, 6);

  SpreadsheetApp.flush();

  // Remove extra blank rows after dashboard.
  // Currently keeps till row 44 unless monthly SIP table needs more rows in future.
  sdTrimDashboardRows_(dashboard, Math.max(lastSipRow, 44));

  if (!silent) {
    SpreadsheetApp.getUi().alert("✅ Dashboard refreshed successfully.");
  }
}

function refreshDashboard() {
  setupUltraDashboard(false);
}

/******************** SUNDAY DASHBOARD MAIL ********************/
function sendWeeklyDashboardMail() {
  setupUltraDashboard(true);

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var tracker = ss.getSheetByName(SD_TRACKER_SHEET);

  if (!tracker) {
    throw new Error("Sheet not found: " + SD_TRACKER_SHEET);
  }

  var allRows = sdGetCleanTrackerRows_(tracker);

  if (!allRows.length) {
    throw new Error("No valid tracker data found.");
  }

  var weeklyRows = sdGetPreviousMondayToSaturdayRows_(tracker, new Date());
  var latest = allRows[allRows.length - 1];

  var invested = Number(latest[1]) || 0;
  var current = Number(latest[2]) || 0;
  var pnl = Number(latest[3]) || 0;
  var returnPct = Number(latest[4]) || 0;

  var statusText = pnl >= 0 ? "🟢 Portfolio Performing Well" : "🔴 Portfolio Under Recovery";
  var pnlColor = pnl >= 0 ? "#059669" : "#dc2626";

  var updatedText = Utilities.formatDate(
    new Date(),
    Session.getScriptTimeZone(),
    "dd-MMM-yyyy hh:mm a"
  );

  var weekLabel = sdGetPreviousWeekLabel_(new Date());

  var pnlBlob = sdCreatePnLChartBlob_(allRows, 520, 220).setName("pnl_chart.png");
  var growthBlob = sdCreateGrowthChartBlob_(allRows, 520, 220).setName("growth_chart.png");
  var pieBlob = sdCreatePieChartBlob_(420, 180).setName("pie_chart.png");

  var subject = "📊 Weekly SIP Dashboard Update - " + weekLabel + " (" + sdFormatPercentText_(returnPct) + ")";

  var htmlBody = sdBuildDashboardEmailHtml_({
    updatedText: updatedText,
    invested: invested,
    current: current,
    pnl: pnl,
    returnPct: returnPct,
    pnlColor: pnlColor,
    statusText: statusText,
    weeklyRows: weeklyRows,
    weekLabel: weekLabel
  });

  MailApp.sendEmail({
    to: SD_MAIL_TO,
    bcc: SD_MAIL_BCC,
    subject: subject,
    htmlBody: htmlBody,
    inlineImages: {
      pnlChart: pnlBlob,
      growthChart: growthBlob,
      pieChart: pieBlob
    }
  });

  Logger.log("Weekly dashboard mail sent successfully.");
}

/******************** DASHBOARD EMAIL HTML ********************/
function sdBuildDashboardEmailHtml_(data) {
  return `
  <div style="margin:0;padding:20px;background:#dfe3ea;font-family:Arial,Helvetica,sans-serif;">
    <table role="presentation" cellpadding="0" cellspacing="0" style="width:100%;max-width:980px;margin:0 auto;border-collapse:collapse;background:#dfe3ea;">
      
      <tr>
        <td colspan="4" style="background:#08152f;color:#ffffff;text-align:center;font-size:18px;font-weight:bold;padding:22px 10px;">
          📊 PERSONAL SIP INVESTMENT DASHBOARD
        </td>
      </tr>

      <tr>
        <td colspan="4" style="background:#1f2937;color:#ffffff;text-align:center;font-size:11px;font-weight:bold;padding:6px 10px;">
          Updated : ${data.updatedText}
        </td>
      </tr>

      <tr><td colspan="4" style="height:18px;"></td></tr>

      <tr>
        <td style="width:25%;padding:8px;">
          <div style="background:#2563eb;color:#ffffff;border:1px solid #111827;text-align:center;font-weight:bold;padding:22px 8px;font-size:13px;">
            💰 Total Invested<br>
            <span style="font-size:15px;">${sdFormatCurrency_(data.invested)}</span>
          </div>
        </td>

        <td style="width:25%;padding:8px;">
          <div style="background:#7c3aed;color:#ffffff;border:1px solid #111827;text-align:center;font-weight:bold;padding:22px 8px;font-size:13px;">
            📈 Current Value<br>
            <span style="font-size:15px;">${sdFormatCurrency_(data.current)}</span>
          </div>
        </td>

        <td style="width:25%;padding:8px;">
          <div style="background:${data.pnlColor};color:#ffffff;border:1px solid #111827;text-align:center;font-weight:bold;padding:22px 8px;font-size:13px;">
            📉 Overall P/L<br>
            <span style="font-size:15px;">${sdFormatCurrency_(data.pnl)}</span>
          </div>
        </td>

        <td style="width:25%;padding:8px;">
          <div style="background:#059669;color:#ffffff;border:1px solid #111827;text-align:center;font-weight:bold;padding:22px 8px;font-size:13px;">
            🎯 Return %<br>
            <span style="font-size:15px;">${sdFormatPercentText_(data.returnPct)}</span>
          </div>
        </td>
      </tr>

      <tr><td colspan="4" style="height:10px;"></td></tr>

      <tr>
        <td colspan="4" style="padding:0 8px;">
          <div style="background:#f3f4f6;border:1px solid #111827;text-align:center;font-size:14px;font-weight:bold;padding:18px;">
            ${data.statusText}
          </div>
        </td>
      </tr>

      <tr><td colspan="4" style="height:18px;"></td></tr>

      <tr>
        <td colspan="4" style="padding:8px;">
          <div style="background:#08152f;color:#ffffff;font-size:13px;font-weight:bold;padding:9px;">
            📅 Weekly Data: Previous Monday to Saturday - ${data.weekLabel}
          </div>

          <div style="background:#ffffff;border:1px solid #111827;padding:10px;">
            ${sdWeeklyRowsHtml_(data.weeklyRows)}
          </div>
        </td>
      </tr>

      <tr><td colspan="4" style="height:18px;"></td></tr>

      <tr>
        <td colspan="2" style="padding:8px;vertical-align:top;">
          <div style="background:#08152f;color:#ffffff;font-size:13px;font-weight:bold;padding:9px;">
            📊 All Time P/L
          </div>
          <div style="background:#ffffff;border:1px solid #111827;text-align:center;padding:10px;">
            <img src="cid:pnlChart" style="width:100%;max-width:520px;height:auto;display:block;margin:0 auto;">
          </div>
        </td>

        <td colspan="2" style="padding:8px;vertical-align:top;">
          <div style="background:#08152f;color:#ffffff;font-size:13px;font-weight:bold;padding:9px;">
            📈 All Time Investment Growth
          </div>
          <div style="background:#ffffff;border:1px solid #111827;text-align:center;padding:10px;">
            <img src="cid:growthChart" style="width:100%;max-width:520px;height:auto;display:block;margin:0 auto;">
          </div>
        </td>
      </tr>

      <tr><td colspan="4" style="height:18px;"></td></tr>

      <tr>
        <td colspan="2" style="padding:8px;vertical-align:top;">
          <div style="background:#08152f;color:#ffffff;font-size:13px;font-weight:bold;padding:9px;">
            📅 Monthly SIP Tracker
          </div>

          <table cellpadding="4" cellspacing="0" style="width:100%;border-collapse:collapse;background:#ffffff;font-size:12px;text-align:center;">
            <tr style="background:#08152f;color:#ffffff;font-weight:bold;">
              <td style="border:1px solid #111827;">Month</td>
              <td style="border:1px solid #111827;">Investment</td>
              <td style="border:1px solid #111827;">Status</td>
              <td style="border:1px solid #111827;">Remarks</td>
            </tr>
            ${sdSipRowsHtml_()}
          </table>
        </td>

        <td colspan="2" style="padding:8px;vertical-align:top;">
          <div style="background:#08152f;color:#ffffff;font-size:13px;font-weight:bold;padding:9px;">
            🥧 Portfolio Allocation
          </div>
          <div style="background:#ffffff;border:1px solid #111827;text-align:center;padding:10px;">
            <img src="cid:pieChart" style="width:100%;max-width:420px;height:auto;display:block;margin:0 auto;">
          </div>
        </td>
      </tr>

    </table>
  </div>
  `;
}

/******************** WEEKLY TABLE HTML ********************/
function sdWeeklyRowsHtml_(rows) {
  if (!rows || !rows.length) {
    return `<div style="font-size:13px;color:#dc2626;font-weight:bold;">No data found for previous Monday to Saturday.</div>`;
  }

  var html = `
    <table cellpadding="6" cellspacing="0" style="width:100%;border-collapse:collapse;font-size:12px;text-align:center;">
      <tr style="background:#08152f;color:#ffffff;font-weight:bold;">
        <td style="border:1px solid #111827;">Date</td>
        <td style="border:1px solid #111827;">Total Invested</td>
        <td style="border:1px solid #111827;">Current Value</td>
        <td style="border:1px solid #111827;">Daily P/L</td>
        <td style="border:1px solid #111827;">Return %</td>
      </tr>
  `;

  rows.forEach(function(r) {
    var pnlColor = Number(r[3]) >= 0 ? "#059669" : "#dc2626";

    html += `
      <tr>
        <td style="border:1px solid #111827;">${sdFormatDateFull_(r[0])}</td>
        <td style="border:1px solid #111827;">${sdFormatCurrency_(r[1])}</td>
        <td style="border:1px solid #111827;">${sdFormatCurrency_(r[2])}</td>
        <td style="border:1px solid #111827;color:${pnlColor};font-weight:bold;">${sdFormatCurrency_(r[3])}</td>
        <td style="border:1px solid #111827;color:${pnlColor};font-weight:bold;">${sdFormatPercentText_(r[4])}</td>
      </tr>
    `;
  });

  html += `</table>`;
  return html;
}

/******************** MONTHLY SIP HTML - AUTOMATED ********************/
function sdSipRowsHtml_() {
  var rows = sdGetMonthlySipRows_();

  return rows.map(function(r) {
    var statusColor = "#dc2626";

    if (r[2] === "Completed") {
      statusColor = "#059669";
    } else if (r[2] === "Partial") {
      statusColor = "#d97706";
    }

    return `
      <tr>
        <td style="border:1px solid #111827;">${r[0]}</td>
        <td style="border:1px solid #111827;">${r[1]}</td>
        <td style="border:1px solid #111827;color:${statusColor};font-weight:bold;">${r[2]}</td>
        <td style="border:1px solid #111827;">${r[3]}</td>
      </tr>
    `;
  }).join("");
}

/******************** SUNDAY EMAIL TRIGGER ********************/
function installSundayDashboardMailTrigger() {
  var triggers = ScriptApp.getProjectTriggers();

  triggers.forEach(function(trigger) {
    if (trigger.getHandlerFunction() === "sendWeeklyDashboardMail") {
      ScriptApp.deleteTrigger(trigger);
    }
  });

  ScriptApp.newTrigger("sendWeeklyDashboardMail")
    .timeBased()
    .onWeekDay(ScriptApp.WeekDay.SUNDAY)
    .atHour(SD_SUNDAY_EMAIL_HOUR)
    .create();

  SpreadsheetApp.getUi().alert("✅ Sunday Dashboard mail trigger installed successfully.");
}

/******************** WEEKLY DATA SHEET REMOVED ********************/
/******************** PREVIOUS MONDAY TO SATURDAY FILTER ********************/
function sdGetPreviousMondayToSaturdayRows_(tracker, refDate) {
  var allRows = sdGetCleanTrackerRows_(tracker);

  var today = new Date(refDate);
  var day = today.getDay();

  var sunday = new Date(today);
  sunday.setDate(today.getDate() - day);
  sunday.setHours(0, 0, 0, 0);

  var monday = new Date(sunday);
  monday.setDate(sunday.getDate() - 6);
  monday.setHours(0, 0, 0, 0);

  var saturday = new Date(sunday);
  saturday.setDate(sunday.getDate() - 1);
  saturday.setHours(23, 59, 59, 999);

  return allRows.filter(function(r) {
    var d = sdParseTrackerDate_(r[0]);
    return d >= monday && d <= saturday;
  });
}

function sdGetPreviousWeekLabel_(refDate) {
  var today = new Date(refDate);
  var day = today.getDay();

  var sunday = new Date(today);
  sunday.setDate(today.getDate() - day);
  sunday.setHours(0, 0, 0, 0);

  var monday = new Date(sunday);
  monday.setDate(sunday.getDate() - 6);

  var saturday = new Date(sunday);
  saturday.setDate(sunday.getDate() - 1);

  return sdFormatDateFull_(monday) + " to " + sdFormatDateFull_(saturday);
}

/******************** SHEET PREP ********************/
function sdPrepareDashboardSheet_(sheet) {
  try {
    var charts = sheet.getCharts();
    charts.forEach(function(ch) {
      sheet.removeChart(ch);
    });
  } catch (e) {}

  try {
    var images = sheet.getImages();
    images.forEach(function(img) {
      img.remove();
    });
  } catch (e) {}

  var requiredColumns = 24; // X column
  var requiredRowsForBuild = 80;

  if (sheet.getMaxColumns() > requiredColumns) {
    sheet.deleteColumns(requiredColumns + 1, sheet.getMaxColumns() - requiredColumns);
  }

  if (sheet.getMaxColumns() < requiredColumns) {
    sheet.insertColumnsAfter(sheet.getMaxColumns(), requiredColumns - sheet.getMaxColumns());
  }

  if (sheet.getMaxRows() < requiredRowsForBuild) {
    sheet.insertRowsAfter(sheet.getMaxRows(), requiredRowsForBuild - sheet.getMaxRows());
  }

  sheet.getRange("A1:X80").breakApart();
  sheet.clear();
  sheet.setHiddenGridlines(true);

  for (var c = 1; c <= 24; c++) {
    sheet.setColumnWidth(c, 52);
  }

  for (var r = 1; r <= 80; r++) {
    sheet.setRowHeight(r, 22);
  }

  sheet.getRange("A1:X80").setBackground("#dfe3ea");
}

/******************** DATA ********************/
function sdGetCleanTrackerRows_(tracker) {
  var lastRow = tracker.getLastRow();

  if (lastRow < 2) {
    return [];
  }

  var rawValues = tracker.getRange(2, 1, lastRow - 1, 7).getValues();
  var displayValues = tracker.getRange(2, 1, lastRow - 1, 7).getDisplayValues();

  var rows = [];

  for (var i = 0; i < rawValues.length; i++) {
    var raw = rawValues[i];
    var display = displayValues[i];

    var dateText = display[0];
    var parsedDate = sdParseTrackerDate_(dateText);

    var totalInvested = Number(raw[1]) || Number(String(display[1]).replace(/[₹,%\s,]/g, "")) || 0;
    var currentValue = Number(raw[2]) || Number(String(display[2]).replace(/[₹,%\s,]/g, "")) || 0;
    var dailyPnL = Number(raw[3]) || Number(String(display[3]).replace(/[₹,%\s,]/g, "")) || 0;
    var dailyPercent = sdNormalizePercent_(raw[4] !== "" ? raw[4] : display[4]);

    if (
      parsedDate instanceof Date &&
      !isNaN(parsedDate.getTime()) &&
      display[0] !== "" &&
      display[1] !== "" &&
      display[2] !== "" &&
      display[3] !== ""
    ) {
      rows.push([
        parsedDate,
        totalInvested,
        currentValue,
        dailyPnL,
        dailyPercent
      ]);
    }
  }

  rows.sort(function(a, b) {
    return a[0].getTime() - b[0].getTime();
  });

  return rows;
}

/******************** CHART BLOBS ********************/
function sdCreatePnLChartBlob_(rows, width, height) {
  var dataTableBuilder = Charts.newDataTable()
    .addColumn(Charts.ColumnType.STRING, "Date")
    .addColumn(Charts.ColumnType.NUMBER, "Daily P/L");

  rows.forEach(function(r) {
    dataTableBuilder.addRow([sdFormatDateShort_(r[0]), Number(r[3]) || 0]);
  });

  var chart = Charts.newColumnChart()
    .setDataTable(dataTableBuilder.build())
    .setDimensions(width, height)
    .setOption("title", "All Time Profit / Loss")
    .setOption("backgroundColor", "#ffffff")
    .setOption("legend", { position: "none" })
    .setOption("colors", ["#0f3b5f"])
    .setOption("chartArea", { left: 60, top: 35, width: "78%", height: "60%" })
    .setOption("hAxis", {
      slantedText: true,
      slantedTextAngle: 90,
      textStyle: { fontSize: 6 }
    })
    .setOption("vAxis", {
      title: "Profit / Loss",
      textStyle: { fontSize: 9 }
    })
    .build();

  return chart.getBlob();
}

function sdCreateGrowthChartBlob_(rows, width, height) {
  var dataTableBuilder = Charts.newDataTable()
    .addColumn(Charts.ColumnType.STRING, "Date")
    .addColumn(Charts.ColumnType.NUMBER, "Total Invested")
    .addColumn(Charts.ColumnType.NUMBER, "Current Value");

  rows.forEach(function(r) {
    dataTableBuilder.addRow([
      sdFormatDateShort_(r[0]),
      Number(r[1]) || 0,
      Number(r[2]) || 0
    ]);
  });

  var chart = Charts.newLineChart()
    .setDataTable(dataTableBuilder.build())
    .setDimensions(width, height)
    .setOption("title", "All Time Investment vs Current Value")
    .setOption("backgroundColor", "#ffffff")
    .setOption("curveType", "function")
    .setOption("lineWidth", 3)
    .setOption("colors", ["#2563eb", "#10b981"])
    .setOption("legend", { position: "top", textStyle: { fontSize: 9 } })
    .setOption("chartArea", { left: 60, top: 35, width: "78%", height: "60%" })
    .setOption("hAxis", {
      slantedText: true,
      slantedTextAngle: 90,
      textStyle: { fontSize: 6 }
    })
    .setOption("vAxis", {
      textStyle: { fontSize: 9 }
    })
    .build();

  return chart.getBlob();
}

function sdCreatePieChartBlob_(width, height) {
  var dataTable = Charts.newDataTable()
    .addColumn(Charts.ColumnType.STRING, "Fund")
    .addColumn(Charts.ColumnType.NUMBER, "Amount")
    .addRow(["ICICI", 16000])
    .addRow(["Parag Parikh", 24000])
    .build();

  var chart = Charts.newPieChart()
    .setDataTable(dataTable)
    .setDimensions(width, height)
    .setOption("title", "Fund Allocation")
    .setOption("backgroundColor", "#ffffff")
    .setOption("pieHole", 0.55)
    .setOption("legend", { position: "right", textStyle: { fontSize: 9 } })
    .setOption("chartArea", { left: 10, top: 25, width: "88%", height: "75%" })
    .build();

  return chart.getBlob();
}

/******************** UI HELPERS ********************/
function sdCreateCard_(sheet, range, title, value, color, type) {
  var cell = sheet.getRange(range);
  cell.merge();

  var formatted = value;

  if (type === "currency") {
    formatted = sdFormatCurrency_(value);
  }

  if (type === "percent") {
    formatted = sdFormatPercentText_(value);
  }

  cell.setValue(title + "\n" + formatted)
    .setBackground(color)
    .setFontColor("white")
    .setFontWeight("bold")
    .setFontSize(8)
    .setHorizontalAlignment("center")
    .setVerticalAlignment("middle")
    .setBorder(true, true, true, true, true, true);
}

function sdCreateSection_(sheet, range, text) {
  sheet.getRange(range)
    .merge()
    .setValue(text)
    .setBackground("#08152f")
    .setFontColor("white")
    .setFontWeight("bold")
    .setFontSize(8)
    .setVerticalAlignment("middle");
}

function sdCreatePanel_(sheet, range) {
  sheet.getRange(range)
    .merge()
    .setBackground("#ffffff")
    .setBorder(true, true, true, true, true, true);
}

/******************** MONTHLY SIP TABLE ********************/
function sdCreateSipTable_(sheet) {
  var sipRows = sdGetMonthlySipRows_();

  var startRow = 34;
  var endClearRow = Math.max(80, startRow + sipRows.length + 5);

  if (sheet.getMaxRows() < endClearRow) {
    sheet.insertRowsAfter(sheet.getMaxRows(), endClearRow - sheet.getMaxRows());
  }

  sheet.getRange("B" + startRow + ":L" + endClearRow).breakApart();
  sheet.getRange("B" + startRow + ":L" + endClearRow).clearContent().clearFormat();

  var headerRow = startRow;
  sdSetMonthlySipMergedRow_(sheet, headerRow, ["Month", "Investment", "Status", "Remarks"], true);

  for (var i = 0; i < sipRows.length; i++) {
    sdSetMonthlySipMergedRow_(sheet, headerRow + i + 1, sipRows[i], false);
  }

  var lastTableRow = headerRow + sipRows.length;

  sheet.getRange("B" + headerRow + ":L" + lastTableRow)
    .setHorizontalAlignment("center")
    .setVerticalAlignment("middle")
    .setFontSize(8)
    .setBorder(true, true, true, true, true, true);

  for (var r = headerRow; r <= lastTableRow; r++) {
    sheet.setRowHeight(r, 24);
  }

  return lastTableRow;
}

function sdSetMonthlySipMergedRow_(sheet, row, values, isHeader) {
  var ranges = [
    "B" + row + ":C" + row,
    "D" + row + ":F" + row,
    "G" + row + ":I" + row,
    "J" + row + ":L" + row
  ];

  for (var i = 0; i < ranges.length; i++) {
    var cell = sheet.getRange(ranges[i]);
    cell.merge();
    cell.setValue(values[i]);
    cell.setBorder(true, true, true, true, true, true);

    if (isHeader) {
      cell.setBackground("#08152f")
        .setFontColor("white")
        .setFontWeight("bold");
    } else {
      cell.setBackground("#dfe3ea")
        .setFontColor("black")
        .setFontWeight("normal");
    }
  }

  if (!isHeader) {
    var status = values[2];
    var statusCell = sheet.getRange("G" + row + ":I" + row);

    if (status === "Completed") {
      statusCell.setFontColor("#059669").setFontWeight("bold");
    } else if (status === "Partial") {
      statusCell.setFontColor("#d97706").setFontWeight("bold");
    } else {
      statusCell.setFontColor("#dc2626").setFontWeight("bold");
    }
  }
}

/******************** MONTHLY SIP AUTOMATION LOGIC ********************/
function sdGetMonthlySipRows_() {
  var latestInvested = sdGetLatestTotalInvested_();

  var rows = [];

  var startMonth = new Date(SD_SIP_START_YEAR, SD_SIP_START_MONTH_INDEX, 1);
  var currentMonth = new Date();
  currentMonth = new Date(currentMonth.getFullYear(), currentMonth.getMonth(), 1);

  var monthDate = new Date(startMonth);
  var monthCount = 0;

  while (monthDate <= currentMonth) {
    monthCount++;

    var expectedTillThisMonth = monthCount * SD_MONTHLY_SIP_AMOUNT;
    var expectedTillPreviousMonth = (monthCount - 1) * SD_MONTHLY_SIP_AMOUNT;

    var status = "Pending";
    var remarks = "";

    if (latestInvested + SD_SIP_TOLERANCE >= expectedTillThisMonth) {
      status = "Completed";
      remarks = "";
    } else if (latestInvested > expectedTillPreviousMonth + SD_SIP_TOLERANCE) {
      status = "Partial";

      var receivedThisMonth = latestInvested - expectedTillPreviousMonth;

      remarks =
        "Received " +
        sdFormatCurrencyNoDecimal_(receivedThisMonth) +
        " / " +
        sdFormatCurrencyNoDecimal_(SD_MONTHLY_SIP_AMOUNT);
    } else {
      status = "Pending";
      remarks = "";
    }

    rows.push([
      sdFormatMonthLabel_(monthDate),
      sdFormatCurrencyNoDecimal_(SD_MONTHLY_SIP_AMOUNT),
      status,
      remarks
    ]);

    monthDate.setMonth(monthDate.getMonth() + 1);
  }

  return rows;
}

function sdGetLatestTotalInvested_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var tracker = ss.getSheetByName(SD_TRACKER_SHEET);

  if (!tracker) {
    return 0;
  }

  var allRows = sdGetCleanTrackerRows_(tracker);

  if (!allRows.length) {
    return 0;
  }

  return Number(allRows[allRows.length - 1][1]) || 0;
}

/******************** ROW CLEANUP ********************/
function sdTrimDashboardRows_(sheet, keepRows) {
  var currentRows = sheet.getMaxRows();

  if (currentRows > keepRows) {
    sheet.deleteRows(keepRows + 1, currentRows - keepRows);
  }

  if (sheet.getMaxRows() < keepRows) {
    sheet.insertRowsAfter(sheet.getMaxRows(), keepRows - sheet.getMaxRows());
  }
}

/******************** COMMON HELPERS ********************/
function sdGetOrCreateSheet_(ss, name) {
  return ss.getSheetByName(name) || ss.insertSheet(name);
}

function sdParseTrackerDate_(value) {
  if (value instanceof Date) {
    return new Date(value.getFullYear(), value.getMonth(), value.getDate());
  }

  if (!value) return null;

  var text = String(value).trim();

  var normalDate = new Date(text);
  if (!isNaN(normalDate.getTime())) {
    return new Date(normalDate.getFullYear(), normalDate.getMonth(), normalDate.getDate());
  }

  var parts = text.split("-");
  if (parts.length === 3) {
    var day = Number(parts[0]);
    var monthText = String(parts[1]).toLowerCase();
    var year = Number(parts[2]);

    if (year < 100) year = 2000 + year;

    var months = {
      jan: 0, january: 0,
      feb: 1, february: 1,
      mar: 2, march: 2,
      apr: 3, april: 3,
      may: 4,
      jun: 5, june: 5,
      jul: 6, july: 6,
      aug: 7, august: 7,
      sep: 8, sept: 8, september: 8,
      oct: 9, october: 9,
      nov: 10, november: 10,
      dec: 11, december: 11
    };

    var month = isNaN(Number(monthText)) ? months[monthText] : Number(monthText) - 1;

    if (!isNaN(day) && month !== undefined && !isNaN(year)) {
      return new Date(year, month, day);
    }
  }

  return null;
}

function sdNumber_(value) {
  if (typeof value === "number") return value;
  if (value === null || value === undefined || value === "") return 0;

  var cleaned = String(value)
    .replace(/₹/g, "")
    .replace(/,/g, "")
    .replace(/%/g, "")
    .replace(/−/g, "-")
    .replace(/\s/g, "")
    .trim();

  var n = Number(cleaned);
  return isNaN(n) ? 0 : n;
}

function sdNormalizePercent_(value) {
  if (typeof value === "number") {
    return Math.abs(value) > 1 ? value / 100 : value;
  }

  if (typeof value === "string") {
    var cleaned = value.replace("%", "").trim();
    var num = Number(cleaned);

    if (!isNaN(num)) {
      return num / 100;
    }
  }

  return 0;
}

function sdFormatCurrency_(value) {
  var num = Number(value) || 0;

  return num.toLocaleString("en-IN", {
    style: "currency",
    currency: "INR",
    maximumFractionDigits: 2
  });
}

function sdFormatCurrencyNoDecimal_(value) {
  var num = Number(value) || 0;

  return "₹" + num.toLocaleString("en-IN", {
    maximumFractionDigits: 0
  });
}

function sdFormatPercentText_(value) {
  var num = Number(value) || 0;
  return (num * 100).toFixed(2) + "%";
}

function sdFormatDateShort_(value) {
  var d = sdParseTrackerDate_(value);

  return Utilities.formatDate(
    d,
    Session.getScriptTimeZone(),
    "dd-MMM-yy"
  );
}

function sdFormatDateFull_(value) {
  var d = sdParseTrackerDate_(value);

  return Utilities.formatDate(
    d,
    Session.getScriptTimeZone(),
    "dd-MMM-yyyy"
  );
}

function sdFormatMonthLabel_(dateValue) {
  return Utilities.formatDate(
    dateValue,
    Session.getScriptTimeZone(),
    "MMM-yy"
  );
}


/************************************************************
 WEEKLY GRAPH MAIL - NO EXTRA SOURCE SHEET
 Uses only:
 SIP Daily Tracker
 Graph
************************************************************/

function buildGraphOnlyFinalV2() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var tracker = ss.getSheetByName("SIP Daily Tracker");

  if (!tracker) {
    throw new Error("SIP Daily Tracker sheet not found.");
  }

  var graphSheet = FINAL_getOrCreateSheet_(ss, "Graph");
  var charts = graphSheet.getCharts();

  charts.forEach(function(chart) {
    graphSheet.removeChart(chart);
  });

  graphSheet.clear();
  graphSheet.setHiddenGridlines(true);

  var lastRow = tracker.getLastRow();

  if (lastRow < 2) {
    throw new Error("SIP Daily Tracker has no data.");
  }

  for (var c = 1; c <= 22; c++) {
    graphSheet.setColumnWidth(c, 80);
  }

  for (var r = 1; r <= 60; r++) {
    graphSheet.setRowHeight(r, 24);
  }

  graphSheet.getRange("A1:V60").setBackground("#ffffff");

  var dateRange = tracker.getRange(1, 1, lastRow, 1);   // A Date
  var totalRange = tracker.getRange(1, 6, lastRow, 1);  // F Total P/L
  var paragRange = tracker.getRange(1, 12, lastRow, 1); // L Parag P/L
  var iciciRange = tracker.getRange(1, 18, lastRow, 1); // R ICICI P/L

  var totalAxis = GRAPH_axisRangeFromRange_(tracker.getRange(2, 6, lastRow - 1, 1).getValues());
  var paragAxis = GRAPH_axisRangeFromRange_(tracker.getRange(2, 12, lastRow - 1, 1).getValues());
  var iciciAxis = GRAPH_axisRangeFromRange_(tracker.getRange(2, 18, lastRow - 1, 1).getValues());

  var width = 1250;
  var height = 360;

  var totalChart = graphSheet.newChart()
    .setChartType(Charts.ChartType.COLUMN)
    .addRange(dateRange)
    .addRange(totalRange)
    .setNumHeaders(1)
    .setPosition(1, 1, 0, 0)
    .setOption("title", "Total P/L by Date")
    .setOption("legend", { position: "none" })
    .setOption("height", height)
    .setOption("width", width)
    .setOption("chartArea", { left: 90, top: 55, width: "82%", height: "62%" })
    .setOption("hAxis", { slantedText: true, slantedTextAngle: 60, textStyle: { fontSize: 7 } })
    .setOption("vAxis", {
      title: "P/L amount",
      baseline: 0,
      viewWindow: { min: totalAxis.min, max: totalAxis.max },
      textStyle: { fontSize: 9 }
    })
    .build();

  graphSheet.insertChart(totalChart);

  var paragChart = graphSheet.newChart()
    .setChartType(Charts.ChartType.COLUMN)
    .addRange(dateRange)
    .addRange(paragRange)
    .setNumHeaders(1)
    .setPosition(19, 1, 0, 0)
    .setOption("title", "Parag Parikh P/L by Date")
    .setOption("legend", { position: "none" })
    .setOption("height", height)
    .setOption("width", width)
    .setOption("chartArea", { left: 90, top: 55, width: "82%", height: "62%" })
    .setOption("hAxis", { slantedText: true, slantedTextAngle: 60, textStyle: { fontSize: 7 } })
    .setOption("vAxis", {
      title: "P/L amount",
      baseline: 0,
      viewWindow: { min: paragAxis.min, max: paragAxis.max },
      textStyle: { fontSize: 9 }
    })
    .build();

  graphSheet.insertChart(paragChart);

  var iciciChart = graphSheet.newChart()
    .setChartType(Charts.ChartType.COLUMN)
    .addRange(dateRange)
    .addRange(iciciRange)
    .setNumHeaders(1)
    .setPosition(37, 1, 0, 0)
    .setOption("title", "ICICI Nifty 50 P/L by Date")
    .setOption("legend", { position: "none" })
    .setOption("height", height)
    .setOption("width", width)
    .setOption("chartArea", { left: 90, top: 55, width: "82%", height: "62%" })
    .setOption("hAxis", { slantedText: true, slantedTextAngle: 60, textStyle: { fontSize: 7 } })
    .setOption("vAxis", {
      title: "P/L amount",
      baseline: 0,
      viewWindow: { min: iciciAxis.min, max: iciciAxis.max },
      textStyle: { fontSize: 9 }
    })
    .build();

  graphSheet.insertChart(iciciChart);

  SpreadsheetApp.flush();
}


function sendSundayGraphMailFinalV2() {
  try {
    FIX_recordDailyTracker_A_to_S();
  } catch (e) {
    Logger.log("Daily tracker update failed before graph mail: " + e);
  }

  buildGraphOnlyFinalV2();

  SpreadsheetApp.flush();
  Utilities.sleep(3000);

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var graphSheet = ss.getSheetByName("Graph");

  if (!graphSheet) {
    throw new Error("Graph sheet not found.");
  }

  var charts = graphSheet.getCharts();

  if (!charts || charts.length < 3) {
    throw new Error("Expected 3 charts in Graph sheet, found: " + charts.length);
  }

  var totalBlob = charts[0].getAs("image/png").setName("total_pnl_by_date.png");
  var paragBlob = charts[1].getAs("image/png").setName("parag_pnl_by_date.png");
  var iciciBlob = charts[2].getAs("image/png").setName("icici_pnl_by_date.png");

  var todayText = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "dd-MMM-yyyy");
  var subject = "📊 Weekly SIP P/L Graphs - " + todayText;

  var htmlBody = `
  <html>
    <body style="font-family:Arial;background:#f5f6f8;padding:20px;">
      <div style="max-width:1100px;margin:auto;background:white;padding:20px;border-radius:14px;box-shadow:0 4px 12px rgba(0,0,0,0.08);">
        <h2 style="margin:0 0 10px 0;color:#111827;">📊 Weekly SIP P/L Graphs</h2>
        <p style="margin:0 0 20px 0;color:#6b7280;font-size:13px;">Updated on ${todayText}</p>

        <div style="margin-bottom:28px;">
          <h3 style="margin:0 0 10px 0;color:#111827;font-size:15px;">1. Total P/L by Date</h3>
          <img src="cid:totalChart" style="width:100%;max-width:1000px;height:auto;border:1px solid #e5e7eb;border-radius:10px;">
        </div>

        <div style="margin-bottom:28px;">
          <h3 style="margin:0 0 10px 0;color:#111827;font-size:15px;">2. Parag Parikh P/L by Date</h3>
          <img src="cid:paragChart" style="width:100%;max-width:1000px;height:auto;border:1px solid #e5e7eb;border-radius:10px;">
        </div>

        <div>
          <h3 style="margin:0 0 10px 0;color:#111827;font-size:15px;">3. ICICI Nifty 50 P/L by Date</h3>
          <img src="cid:iciciChart" style="width:100%;max-width:1000px;height:auto;border:1px solid #e5e7eb;border-radius:10px;">
        </div>

        <p style="margin-top:20px;color:#6b7280;font-size:12px;">
          Automated from SIP Daily Tracker.
        </p>
      </div>
    </body>
  </html>
  `;

  MailApp.sendEmail({
    to: MAIL_TO,
    bcc: MAIL_BCC,
    subject: subject,
    htmlBody: htmlBody,
    inlineImages: {
      totalChart: totalBlob,
      paragChart: paragBlob,
      iciciChart: iciciBlob
    }
  });

  Logger.log("Sunday graph mail sent successfully.");
}


// Backward-compatible alias.
function updateGraphsAndSendEmail() {
  sendSundayGraphMailFinalV2();
}


function GRAPH_axisRangeFromRange_(values) {
  var nums = values.map(function(row) {
    return FINAL_num_(row[0]);
  }).filter(function(v) {
    return !isNaN(v);
  });

  if (!nums.length) {
    return { min: -100, max: 100 };
  }

  var min = Math.min.apply(null, nums);
  var max = Math.max.apply(null, nums);

  min = Math.min(min, 0);
  max = Math.max(max, 0);

  var range = max - min;

  if (range === 0) {
    range = Math.abs(max) || 100;
  }

  var padding = range * 0.15;

  return {
    min: min - padding,
    max: max + padding
  };
}



/************************************************************
 FINAL TRIGGER + CLEANUP FUNCTIONS
************************************************************/

function installFinalSipAutomationTriggers() {
  FINAL_deleteTriggersByNames_([
    "morningSipAutomation",
    "sendPortfolioAlert",
    "sendDailyPortfolioMail",
    "FIX_recordDailyTracker_A_to_S",
    "recordDailyTracker",
    "sendWeeklyDashboardMail",
    "sendSundayGraphMailFinalV2",
    "sendSundayGraphMailFinal",
    "sendSundayGraphMail",
    "sendWeeklySipTrackerGraphMail",
    "updateGraphsAndSendEmail",
    "runAllAutomations"
  ]);

  ScriptApp.newTrigger("morningSipAutomation")
    .timeBased()
    .everyDays(1)
    .atHour(6)
    .nearMinute(30)
    .create();

  ScriptApp.newTrigger("sendWeeklyDashboardMail")
    .timeBased()
    .onWeekDay(ScriptApp.WeekDay.SUNDAY)
    .atHour(9)
    .create();

  ScriptApp.newTrigger("sendSundayGraphMailFinalV2")
    .timeBased()
    .onWeekDay(ScriptApp.WeekDay.SUNDAY)
    .atHour(11)
    .create();

  Logger.log("Final triggers installed: morningSipAutomation, sendWeeklyDashboardMail, sendSundayGraphMailFinalV2.");
}


function FINAL_deleteTriggersByNames_(functionNames) {
  var triggers = ScriptApp.getProjectTriggers();

  triggers.forEach(function(trigger) {
    var fn = trigger.getHandlerFunction();

    if (functionNames.indexOf(fn) !== -1) {
      ScriptApp.deleteTrigger(trigger);
      Logger.log("Deleted trigger: " + fn);
    }
  });
}


function cleanupUnusedSipSheets() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  var keepSheets = [
    "SIP",
    "SIP Daily Tracker",
    "Graph",
    "SIP Dashboard"
  ];

  var deleteSheets = [
    "Dashboard",
    "Weekly SIP Data",
    "Monthly Ledger",
    "Graph Source Data",
    "Dashboard Chart Data"
  ];

  deleteSheets.forEach(function(name) {
    var sheet = ss.getSheetByName(name);

    if (sheet && keepSheets.indexOf(name) === -1 && ss.getSheets().length > 1) {
      ss.deleteSheet(sheet);
      Logger.log("Deleted unused sheet: " + name);
    }
  });
}


// Optional manual test that runs full daily flow now.
function testDailyFlowNow() {
  morningSipAutomation();
}


// Optional manual test that refreshes dashboard and graph sheets.
function testWeeklyVisualsNow() {
  setupUltraDashboard(true);
  buildGraphOnlyFinalV2();
}


// Backward-compatible alias for old menu/trigger.
function runAllAutomations() {
  morningSipAutomation();
}
/************************************************************
 PATCH FIX - SIP DASHBOARD + GRAPH SOURCE MAPPING
 Paste this at the very bottom of SIP_Final_Merged_Automation.gs

 Fixes:
 1. SIP Dashboard reading wrong columns.
 2. Graph charts going blank.
 3. No extra sheet needed for graph source.
************************************************************/


/******************** PATCH: CORRECT DASHBOARD DATA SOURCE A:S ********************/

function sdGetCleanTrackerRows_(tracker) {
  var lastRow = tracker.getLastRow();

  if (lastRow < 2) {
    return [];
  }

  // Read full A:S from SIP Daily Tracker
  var rawValues = tracker.getRange(2, 1, lastRow - 1, 19).getValues();
  var displayValues = tracker.getRange(2, 1, lastRow - 1, 19).getDisplayValues();

  var rows = [];

  for (var i = 0; i < rawValues.length; i++) {
    var raw = rawValues[i];
    var display = displayValues[i];

    var parsedDate = FINAL_PARSE_DATE_(raw[0]) || FINAL_PARSE_DATE_(display[0]);

    // Correct A:S mapping
    var totalInvested = FINAL_NUM_(raw[3] !== "" ? raw[3] : display[3]);      // D
    var currentValue  = FINAL_NUM_(raw[4] !== "" ? raw[4] : display[4]);      // E
    var totalPL       = FINAL_NUM_(raw[5] !== "" ? raw[5] : display[5]);      // F
    var returnPct     = FINAL_PERCENT_(raw[6], display[6]);                  // G

    var paragUnits    = FINAL_NUM_(raw[7] !== "" ? raw[7] : display[7]);      // H
    var paragInvested = FINAL_NUM_(raw[9] !== "" ? raw[9] : display[9]);      // J
    var paragCurrent  = FINAL_NUM_(raw[10] !== "" ? raw[10] : display[10]);   // K
    var paragPL       = FINAL_NUM_(raw[11] !== "" ? raw[11] : display[11]);   // L
    var paragReturn   = FINAL_PERCENT_(raw[12], display[12]);                // M

    var iciciUnits    = FINAL_NUM_(raw[13] !== "" ? raw[13] : display[13]);   // N
    var iciciInvested = FINAL_NUM_(raw[15] !== "" ? raw[15] : display[15]);   // P
    var iciciCurrent  = FINAL_NUM_(raw[16] !== "" ? raw[16] : display[16]);   // Q
    var iciciPL       = FINAL_NUM_(raw[17] !== "" ? raw[17] : display[17]);   // R
    var iciciReturn   = FINAL_PERCENT_(raw[18], display[18]);                // S

    if (
      parsedDate instanceof Date &&
      !isNaN(parsedDate.getTime()) &&
      totalInvested > 0 &&
      Math.abs(currentValue) > 1 &&
      paragUnits > 0 &&
      iciciUnits > 0
    ) {
      rows.push([
        parsedDate,       // 0 Date
        totalInvested,    // 1 Total Invested
        currentValue,     // 2 Current Value
        totalPL,          // 3 Total P/L
        returnPct,        // 4 Total Return %

        paragInvested,    // 5 Parag Invested
        paragCurrent,     // 6 Parag Current
        paragPL,          // 7 Parag P/L
        paragReturn,      // 8 Parag Return %

        iciciInvested,    // 9 ICICI Invested
        iciciCurrent,     // 10 ICICI Current
        iciciPL,          // 11 ICICI P/L
        iciciReturn       // 12 ICICI Return %
      ]);
    }
  }

  rows.sort(function(a, b) {
    return a[0].getTime() - b[0].getTime();
  });

  return rows;
}


/******************** PATCH: PORTFOLIO ALLOCATION PIE CHART ********************/

function sdCreatePieChartBlob_(width, height) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var tracker = ss.getSheetByName("SIP Daily Tracker");
  var rows = tracker ? sdGetCleanTrackerRows_(tracker) : [];

  var paragInvested = 0;
  var iciciInvested = 0;

  if (rows.length) {
    var latest = rows[rows.length - 1];

    paragInvested = Number(latest[5]) || 0; // Parag Invested
    iciciInvested = Number(latest[9]) || 0; // ICICI Invested
  }

  if (!paragInvested && !iciciInvested) {
    paragInvested = 24000;
    iciciInvested = 16000;
  }

  var dataTable = Charts.newDataTable()
    .addColumn(Charts.ColumnType.STRING, "Fund")
    .addColumn(Charts.ColumnType.NUMBER, "Amount")
    .addRow(["ICICI", iciciInvested])
    .addRow(["Parag Parikh", paragInvested])
    .build();

  var chart = Charts.newPieChart()
    .setDataTable(dataTable)
    .setDimensions(width, height)
    .setOption("title", "Fund Allocation")
    .setOption("backgroundColor", "#ffffff")
    .setOption("pieHole", 0.55)
    .setOption("legend", { position: "right", textStyle: { fontSize: 9 } })
    .setOption("chartArea", { left: 10, top: 25, width: "88%", height: "75%" })
    .build();

  return chart.getBlob();
}


/******************** PATCH: GRAPH SOURCE WITHOUT EXTRA SHEET ********************/

function buildGraphOnlyFinalV2() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  var tracker = ss.getSheetByName("SIP Daily Tracker");
  if (!tracker) {
    throw new Error("SIP Daily Tracker sheet not found.");
  }

  var graphSheet = FINAL_getOrCreateSheet_(ss, "Graph");

  // Clear only Graph sheet
  var oldCharts = graphSheet.getCharts();
  oldCharts.forEach(function(chart) {
    graphSheet.removeChart(chart);
  });

  graphSheet.clear();
  graphSheet.setHiddenGridlines(true);

  var lastRow = tracker.getLastRow();

  if (lastRow < 2) {
    throw new Error("SIP Daily Tracker has no data.");
  }

  // Make sure Graph has enough rows/columns for charts + hidden source data
  if (graphSheet.getMaxRows() < 60) {
    graphSheet.insertRowsAfter(graphSheet.getMaxRows(), 60 - graphSheet.getMaxRows());
  }

  if (graphSheet.getMaxColumns() < 33) {
    graphSheet.insertColumnsAfter(graphSheet.getMaxColumns(), 33 - graphSheet.getMaxColumns());
  }

  for (var c = 1; c <= 22; c++) {
    graphSheet.setColumnWidth(c, 80);
  }

  for (var r = 1; r <= 60; r++) {
    graphSheet.setRowHeight(r, 24);
  }

  graphSheet.getRange("A1:V60").setBackground("#ffffff");

  // Read A:S
  var values = tracker.getRange(2, 1, lastRow - 1, 19).getValues();
  var displayValues = tracker.getRange(2, 1, lastRow - 1, 19).getDisplayValues();

  var tz = Session.getScriptTimeZone();

  var totalRows = [["Date", "Total P/L"]];
  var paragRows = [["Date", "Parag P/L"]];
  var iciciRows = [["Date", "ICICI P/L"]];

  for (var i = 0; i < values.length; i++) {
    var row = values[i];
    var display = displayValues[i];

    var dateValue = FINAL_PARSE_DATE_(row[0]) || FINAL_PARSE_DATE_(display[0]);
    if (!dateValue) continue;

    var label = Utilities.formatDate(dateValue, tz, "dd-MMM");

    var totalPL = FINAL_NUM_(row[5] !== "" ? row[5] : display[5]);     // F
    var paragPL = FINAL_NUM_(row[11] !== "" ? row[11] : display[11]);  // L
    var iciciPL = FINAL_NUM_(row[17] !== "" ? row[17] : display[17]);  // R

    totalRows.push([label, totalPL]);
    paragRows.push([label, paragPL]);
    iciciRows.push([label, iciciPL]);
  }

  if (totalRows.length <= 1) {
    throw new Error("No valid rows found for Graph.");
  }

  /*
    Hidden helper data inside Graph sheet itself:
    Z:AA  = Total P/L
    AC:AD = Parag P/L
    AF:AG = ICICI P/L
  */
  var totalStartCol = 26; // Z
  var paragStartCol = 29; // AC
  var iciciStartCol = 32; // AF

  graphSheet.getRange(1, totalStartCol, totalRows.length, 2).setValues(totalRows);
  graphSheet.getRange(1, paragStartCol, paragRows.length, 2).setValues(paragRows);
  graphSheet.getRange(1, iciciStartCol, iciciRows.length, 2).setValues(iciciRows);

  graphSheet.getRange(1, totalStartCol, 1, 2).setFontWeight("bold").setBackground("#dbeafe");
  graphSheet.getRange(1, paragStartCol, 1, 2).setFontWeight("bold").setBackground("#dbeafe");
  graphSheet.getRange(1, iciciStartCol, 1, 2).setFontWeight("bold").setBackground("#dbeafe");

  graphSheet.getRange(1, totalStartCol + 1, totalRows.length, 1).setNumberFormat("₹#,##0.00");
  graphSheet.getRange(1, paragStartCol + 1, paragRows.length, 1).setNumberFormat("₹#,##0.00");
  graphSheet.getRange(1, iciciStartCol + 1, iciciRows.length, 1).setNumberFormat("₹#,##0.00");

  SpreadsheetApp.flush();

  var width = 1250;
  var height = 360;

  var totalRange = graphSheet.getRange(1, totalStartCol, totalRows.length, 2);
  var paragRange = graphSheet.getRange(1, paragStartCol, paragRows.length, 2);
  var iciciRange = graphSheet.getRange(1, iciciStartCol, iciciRows.length, 2);

  var totalAxis = GRAPH_AXIS_FROM_ROWS_(totalRows);
  var paragAxis = GRAPH_AXIS_FROM_ROWS_(paragRows);
  var iciciAxis = GRAPH_AXIS_FROM_ROWS_(iciciRows);

  var totalChart = graphSheet.newChart()
    .setChartType(Charts.ChartType.COLUMN)
    .addRange(totalRange)
    .setNumHeaders(1)
    .setPosition(1, 1, 0, 0)
    .setOption("title", "Total P/L by Date")
    .setOption("legend", { position: "none" })
    .setOption("useFirstColumnAsDomain", true)
    .setOption("height", height)
    .setOption("width", width)
    .setOption("chartArea", { left: 90, top: 55, width: "82%", height: "62%" })
    .setOption("hAxis", { slantedText: true, slantedTextAngle: 60, textStyle: { fontSize: 7 } })
    .setOption("vAxis", {
      title: "P/L amount",
      baseline: 0,
      viewWindow: { min: totalAxis.min, max: totalAxis.max },
      textStyle: { fontSize: 9 }
    })
    .build();

  graphSheet.insertChart(totalChart);

  var paragChart = graphSheet.newChart()
    .setChartType(Charts.ChartType.COLUMN)
    .addRange(paragRange)
    .setNumHeaders(1)
    .setPosition(19, 1, 0, 0)
    .setOption("title", "Parag Parikh P/L by Date")
    .setOption("legend", { position: "none" })
    .setOption("useFirstColumnAsDomain", true)
    .setOption("height", height)
    .setOption("width", width)
    .setOption("chartArea", { left: 90, top: 55, width: "82%", height: "62%" })
    .setOption("hAxis", { slantedText: true, slantedTextAngle: 60, textStyle: { fontSize: 7 } })
    .setOption("vAxis", {
      title: "P/L amount",
      baseline: 0,
      viewWindow: { min: paragAxis.min, max: paragAxis.max },
      textStyle: { fontSize: 9 }
    })
    .build();

  graphSheet.insertChart(paragChart);

  var iciciChart = graphSheet.newChart()
    .setChartType(Charts.ChartType.COLUMN)
    .addRange(iciciRange)
    .setNumHeaders(1)
    .setPosition(37, 1, 0, 0)
    .setOption("title", "ICICI Nifty 50 P/L by Date")
    .setOption("legend", { position: "none" })
    .setOption("useFirstColumnAsDomain", true)
    .setOption("height", height)
    .setOption("width", width)
    .setOption("chartArea", { left: 90, top: 55, width: "82%", height: "62%" })
    .setOption("hAxis", { slantedText: true, slantedTextAngle: 60, textStyle: { fontSize: 7 } })
    .setOption("vAxis", {
      title: "P/L amount",
      baseline: 0,
      viewWindow: { min: iciciAxis.min, max: iciciAxis.max },
      textStyle: { fontSize: 9 }
    })
    .build();

  graphSheet.insertChart(iciciChart);

  // Hide helper columns Z:AG so Graph sheet stays clean
  try {
    graphSheet.hideColumns(26, 8);
  } catch (e) {
    Logger.log("Could not hide graph helper columns: " + e.message);
  }

  SpreadsheetApp.flush();

  Logger.log("Graph rebuilt successfully using hidden helper data inside Graph sheet.");
}


/******************** PATCH TEST FUNCTION ********************/

function fixDashboardAndGraphSourcesNow() {
  setupUltraDashboard(true);
  buildGraphOnlyFinalV2();
  SpreadsheetApp.flush();
  Logger.log("SIP Dashboard and Graph source fix completed.");
}

// Override old test function also
function testWeeklyVisualsNow() {
  fixDashboardAndGraphSourcesNow();
}


/******************** PATCH HELPERS ********************/

function GRAPH_AXIS_FROM_ROWS_(rows) {
  var values = [];

  for (var i = 1; i < rows.length; i++) {
    values.push(Number(rows[i][1]) || 0);
  }

  if (!values.length) {
    return { min: -100, max: 100 };
  }

  var min = Math.min.apply(null, values);
  var max = Math.max.apply(null, values);

  if (min === max) {
    min = min - 100;
    max = max + 100;
  }

  var padding = Math.max(100, Math.abs(max - min) * 0.15);

  return {
    min: Math.floor(min - padding),
    max: Math.ceil(max + padding)
  };
}


function FINAL_NUM_(value) {
  if (value === null || value === "") return 0;

  if (typeof value === "number") return value;

  var text = String(value)
    .replace(/₹/g, "")
    .replace(/,/g, "")
    .replace(/%/g, "")
    .replace(/\(/g, "-")
    .replace(/\)/g, "")
    .trim();

  var n = Number(text);

  return isNaN(n) ? 0 : n;
}


function FINAL_PERCENT_(rawValue, displayValue) {
  var displayText = String(displayValue || "").trim();

  if (displayText.indexOf("%") !== -1) {
    return FINAL_NUM_(displayText) / 100;
  }

  var n = FINAL_NUM_(rawValue);

  if (Math.abs(n) > 1) {
    return n / 100;
  }

  return n;
}


function FINAL_PARSE_DATE_(value) {
  if (value instanceof Date) {
    return new Date(value.getFullYear(), value.getMonth(), value.getDate());
  }

  if (!value) return null;

  var text = String(value).trim();

  var d = new Date(text);
  if (!isNaN(d.getTime())) {
    return new Date(d.getFullYear(), d.getMonth(), d.getDate());
  }

  var parts = text.split("-");

  if (parts.length !== 3) return null;

  var day = Number(parts[0]);
  var monthText = String(parts[1]).toLowerCase();
  var year = Number(parts[2]);

  if (year < 100) year = 2000 + year;

  var months = {
    jan: 0, january: 0,
    feb: 1, february: 1,
    mar: 2, march: 2,
    apr: 3, april: 3,
    may: 4,
    jun: 5, june: 5,
    jul: 6, july: 6,
    aug: 7, august: 7,
    sep: 8, sept: 8, september: 8,
    oct: 9, october: 9,
    nov: 10, november: 10,
    dec: 11, december: 11
  };

  var month;

  if (isNaN(Number(monthText))) {
    month = months[monthText];
  } else {
    month = Number(monthText) - 1;
  }

  if (isNaN(day) || month === undefined || isNaN(year)) return null;

  return new Date(year, month, day);
}
/************************************************************
 SIMPLE GRAPH FIX
 Source directly from SIP Daily Tracker:
 A = Date
 F = Total P/L
 L = Parag P/L
 R = ICICI P/L

 Output:
 Graph sheet with 3 charts only.
 No helper sheet.
 No hidden source data.
************************************************************/


function buildGraphOnlyFinalV2() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  var tracker = ss.getSheetByName("SIP Daily Tracker");
  if (!tracker) {
    throw new Error("SIP Daily Tracker sheet not found.");
  }

  var graphSheet = ss.getSheetByName("Graph");
  if (!graphSheet) {
    graphSheet = ss.insertSheet("Graph");
  }

  // Remove old charts
  var oldCharts = graphSheet.getCharts();
  oldCharts.forEach(function(chart) {
    graphSheet.removeChart(chart);
  });

  // Clear only Graph sheet UI
  graphSheet.clear();
  graphSheet.setHiddenGridlines(true);

  var lastRow = tracker.getLastRow();

  if (lastRow < 2) {
    throw new Error("SIP Daily Tracker has no data.");
  }

  // Keep only rows having date in Column A and values in F/L/R
  var lastValidRow = getLastValidGraphRow_(tracker);

  if (lastValidRow < 2) {
    throw new Error("No valid rows found in SIP Daily Tracker for graph.");
  }

  // Direct ranges from SIP Daily Tracker
  var dateRange = tracker.getRange(1, 1, lastValidRow, 1);       // A Date
  var totalPLRange = tracker.getRange(1, 6, lastValidRow, 1);    // F Total P/L
  var paragPLRange = tracker.getRange(1, 12, lastValidRow, 1);   // L Parag P/L
  var iciciPLRange = tracker.getRange(1, 18, lastValidRow, 1);   // R ICICI P/L

  // Layout
  if (graphSheet.getMaxRows() < 60) {
    graphSheet.insertRowsAfter(graphSheet.getMaxRows(), 60 - graphSheet.getMaxRows());
  }

  if (graphSheet.getMaxColumns() < 22) {
    graphSheet.insertColumnsAfter(graphSheet.getMaxColumns(), 22 - graphSheet.getMaxColumns());
  }

  for (var c = 1; c <= 22; c++) {
    graphSheet.setColumnWidth(c, 80);
  }

  for (var r = 1; r <= 60; r++) {
    graphSheet.setRowHeight(r, 24);
  }

  graphSheet.getRange("A1:V60").setBackground("#ffffff");

  var chartWidth = 1250;
  var chartHeight = 360;

  // Chart 1 - Total P/L
  var totalChart = graphSheet.newChart()
    .setChartType(Charts.ChartType.COLUMN)
    .addRange(dateRange)
    .addRange(totalPLRange)
    .setNumHeaders(1)
    .setPosition(1, 1, 0, 0)
    .setOption("title", "Total P/L by Date")
    .setOption("legend", { position: "none" })
    .setOption("useFirstColumnAsDomain", true)
    .setOption("width", chartWidth)
    .setOption("height", chartHeight)
    .setOption("chartArea", {
      left: 90,
      top: 55,
      width: "82%",
      height: "62%"
    })
    .setOption("hAxis", {
      slantedText: true,
      slantedTextAngle: 60,
      textStyle: { fontSize: 7 }
    })
    .setOption("vAxis", {
      title: "Total P/L",
      baseline: 0,
      textStyle: { fontSize: 9 }
    })
    .build();

  graphSheet.insertChart(totalChart);

  // Chart 2 - Parag P/L
  var paragChart = graphSheet.newChart()
    .setChartType(Charts.ChartType.COLUMN)
    .addRange(dateRange)
    .addRange(paragPLRange)
    .setNumHeaders(1)
    .setPosition(19, 1, 0, 0)
    .setOption("title", "Parag Parikh P/L by Date")
    .setOption("legend", { position: "none" })
    .setOption("useFirstColumnAsDomain", true)
    .setOption("width", chartWidth)
    .setOption("height", chartHeight)
    .setOption("chartArea", {
      left: 90,
      top: 55,
      width: "82%",
      height: "62%"
    })
    .setOption("hAxis", {
      slantedText: true,
      slantedTextAngle: 60,
      textStyle: { fontSize: 7 }
    })
    .setOption("vAxis", {
      title: "Parag P/L",
      baseline: 0,
      textStyle: { fontSize: 9 }
    })
    .build();

  graphSheet.insertChart(paragChart);

  // Chart 3 - ICICI P/L
  var iciciChart = graphSheet.newChart()
    .setChartType(Charts.ChartType.COLUMN)
    .addRange(dateRange)
    .addRange(iciciPLRange)
    .setNumHeaders(1)
    .setPosition(37, 1, 0, 0)
    .setOption("title", "ICICI Nifty 50 P/L by Date")
    .setOption("legend", { position: "none" })
    .setOption("useFirstColumnAsDomain", true)
    .setOption("width", chartWidth)
    .setOption("height", chartHeight)
    .setOption("chartArea", {
      left: 90,
      top: 55,
      width: "82%",
      height: "62%"
    })
    .setOption("hAxis", {
      slantedText: true,
      slantedTextAngle: 60,
      textStyle: { fontSize: 7 }
    })
    .setOption("vAxis", {
      title: "ICICI P/L",
      baseline: 0,
      textStyle: { fontSize: 9 }
    })
    .build();

  graphSheet.insertChart(iciciChart);

  SpreadsheetApp.flush();

  Logger.log("Simple graph created successfully from SIP Daily Tracker direct columns.");
}


/************************************************************
 Find last valid row for graph
************************************************************/

function getLastValidGraphRow_(tracker) {
  var lastRow = tracker.getLastRow();

  if (lastRow < 2) return 1;

  var values = tracker.getRange(2, 1, lastRow - 1, 18).getValues();

  for (var i = values.length - 1; i >= 0; i--) {
    var row = values[i];

    var dateValue = row[0];   // A
    var totalPL = row[5];     // F
    var paragPL = row[11];    // L
    var iciciPL = row[17];    // R

    var validDate = dateValue instanceof Date && !isNaN(dateValue.getTime());

    if (
      validDate &&
      totalPL !== "" &&
      paragPL !== "" &&
      iciciPL !== ""
    ) {
      return i + 2;
    }
  }

  return 1;
}


/************************************************************
 Test function
************************************************************/

function fixGraphSimpleNow() {
  buildGraphOnlyFinalV2();
}
/************************************************************
 FINAL SIMPLE GRAPH FIX - SAME GRAPH SHEET SOURCE DATA

Source:
SIP Daily Tracker
A = Date
F = Total P/L
L = Parag P/L
R = ICICI P/L

Output:
Graph sheet with 3 working charts.

No extra sheet.
No hidden columns.
Source data is copied inside Graph sheet in far-right columns:
Z:AA  = Total P/L
AC:AD = Parag P/L
AF:AG = ICICI P/L
************************************************************/


function buildGraphOnlyFinalV2() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  var tracker = ss.getSheetByName("SIP Daily Tracker");
  if (!tracker) {
    throw new Error("SIP Daily Tracker sheet not found.");
  }

  var graphSheet = ss.getSheetByName("Graph");
  if (!graphSheet) {
    graphSheet = ss.insertSheet("Graph");
  }

  // Remove old charts
  var oldCharts = graphSheet.getCharts();
  oldCharts.forEach(function(chart) {
    graphSheet.removeChart(chart);
  });

  // Clear Graph sheet
  graphSheet.clear();
  graphSheet.setHiddenGridlines(true);

  var lastRow = tracker.getLastRow();

  if (lastRow < 2) {
    throw new Error("SIP Daily Tracker has no data.");
  }

  // Read A:R from SIP Daily Tracker
  var values = tracker.getRange(2, 1, lastRow - 1, 18).getValues();
  var displayValues = tracker.getRange(2, 1, lastRow - 1, 18).getDisplayValues();

  var totalRows = [["Date", "Total P/L"]];
  var paragRows = [["Date", "Parag P/L"]];
  var iciciRows = [["Date", "ICICI P/L"]];

  for (var i = 0; i < values.length; i++) {
    var row = values[i];
    var display = displayValues[i];

    var dateValue = graphFinalParseDate_(row[0]) || graphFinalParseDate_(display[0]);

    if (!dateValue) continue;

    var totalPL = graphFinalNum_(row[5] !== "" ? row[5] : display[5]);     // F
    var paragPL = graphFinalNum_(row[11] !== "" ? row[11] : display[11]);  // L
    var iciciPL = graphFinalNum_(row[17] !== "" ? row[17] : display[17]);  // R

    totalRows.push([dateValue, totalPL]);
    paragRows.push([dateValue, paragPL]);
    iciciRows.push([dateValue, iciciPL]);
  }

  if (totalRows.length <= 1) {
    throw new Error("No valid graph rows found from SIP Daily Tracker.");
  }

  // Make sure Graph has enough rows and columns
  if (graphSheet.getMaxRows() < Math.max(totalRows.length + 5, 60)) {
    graphSheet.insertRowsAfter(
      graphSheet.getMaxRows(),
      Math.max(totalRows.length + 5, 60) - graphSheet.getMaxRows()
    );
  }

  if (graphSheet.getMaxColumns() < 34) {
    graphSheet.insertColumnsAfter(graphSheet.getMaxColumns(), 34 - graphSheet.getMaxColumns());
  }

  // Layout
  for (var c = 1; c <= 34; c++) {
    graphSheet.setColumnWidth(c, 80);
  }

  for (var r = 1; r <= 60; r++) {
    graphSheet.setRowHeight(r, 24);
  }

  graphSheet.getRange("A1:AH60").setBackground("#ffffff");

  /*
    Put chart source data inside Graph sheet itself.
    Do NOT hide columns because Google charts sometimes go blank when source is hidden.
  */
  var totalStartCol = 26; // Z
  var paragStartCol = 29; // AC
  var iciciStartCol = 32; // AF

  graphSheet.getRange(1, totalStartCol, totalRows.length, 2).setValues(totalRows);
  graphSheet.getRange(1, paragStartCol, paragRows.length, 2).setValues(paragRows);
  graphSheet.getRange(1, iciciStartCol, iciciRows.length, 2).setValues(iciciRows);

  graphSheet.getRange(1, totalStartCol, 1, 2).setFontWeight("bold").setBackground("#dbeafe");
  graphSheet.getRange(1, paragStartCol, 1, 2).setFontWeight("bold").setBackground("#dbeafe");
  graphSheet.getRange(1, iciciStartCol, 1, 2).setFontWeight("bold").setBackground("#dbeafe");

  graphSheet.getRange(2, totalStartCol, totalRows.length - 1, 1).setNumberFormat("dd-mmm-yyyy");
  graphSheet.getRange(2, paragStartCol, paragRows.length - 1, 1).setNumberFormat("dd-mmm-yyyy");
  graphSheet.getRange(2, iciciStartCol, iciciRows.length - 1, 1).setNumberFormat("dd-mmm-yyyy");

  graphSheet.getRange(2, totalStartCol + 1, totalRows.length - 1, 1).setNumberFormat("₹#,##0.00");
  graphSheet.getRange(2, paragStartCol + 1, paragRows.length - 1, 1).setNumberFormat("₹#,##0.00");
  graphSheet.getRange(2, iciciStartCol + 1, iciciRows.length - 1, 1).setNumberFormat("₹#,##0.00");

  SpreadsheetApp.flush();
  Utilities.sleep(2000);

  var chartWidth = 1250;
  var chartHeight = 360;

  var totalRange = graphSheet.getRange(1, totalStartCol, totalRows.length, 2);
  var paragRange = graphSheet.getRange(1, paragStartCol, paragRows.length, 2);
  var iciciRange = graphSheet.getRange(1, iciciStartCol, iciciRows.length, 2);

  // Chart 1 - Total P/L
  var totalChart = graphSheet.newChart()
    .setChartType(Charts.ChartType.COLUMN)
    .addRange(totalRange)
    .setNumHeaders(1)
    .setPosition(1, 1, 0, 0)
    .setOption("title", "Total P/L by Date")
    .setOption("legend", { position: "none" })
    .setOption("useFirstColumnAsDomain", true)
    .setOption("width", chartWidth)
    .setOption("height", chartHeight)
    .setOption("chartArea", {
      left: 90,
      top: 55,
      width: "82%",
      height: "62%"
    })
    .setOption("hAxis", {
      slantedText: true,
      slantedTextAngle: 60,
      textStyle: { fontSize: 7 }
    })
    .setOption("vAxis", {
      title: "Total P/L",
      baseline: 0,
      textStyle: { fontSize: 9 }
    })
    .build();

  graphSheet.insertChart(totalChart);

  // Chart 2 - Parag P/L
  var paragChart = graphSheet.newChart()
    .setChartType(Charts.ChartType.COLUMN)
    .addRange(paragRange)
    .setNumHeaders(1)
    .setPosition(19, 1, 0, 0)
    .setOption("title", "Parag Parikh P/L by Date")
    .setOption("legend", { position: "none" })
    .setOption("useFirstColumnAsDomain", true)
    .setOption("width", chartWidth)
    .setOption("height", chartHeight)
    .setOption("chartArea", {
      left: 90,
      top: 55,
      width: "82%",
      height: "62%"
    })
    .setOption("hAxis", {
      slantedText: true,
      slantedTextAngle: 60,
      textStyle: { fontSize: 7 }
    })
    .setOption("vAxis", {
      title: "Parag P/L",
      baseline: 0,
      textStyle: { fontSize: 9 }
    })
    .build();

  graphSheet.insertChart(paragChart);

  // Chart 3 - ICICI P/L
  var iciciChart = graphSheet.newChart()
    .setChartType(Charts.ChartType.COLUMN)
    .addRange(iciciRange)
    .setNumHeaders(1)
    .setPosition(37, 1, 0, 0)
    .setOption("title", "ICICI Nifty 50 P/L by Date")
    .setOption("legend", { position: "none" })
    .setOption("useFirstColumnAsDomain", true)
    .setOption("width", chartWidth)
    .setOption("height", chartHeight)
    .setOption("chartArea", {
      left: 90,
      top: 55,
      width: "82%",
      height: "62%"
    })
    .setOption("hAxis", {
      slantedText: true,
      slantedTextAngle: 60,
      textStyle: { fontSize: 7 }
    })
    .setOption("vAxis", {
      title: "ICICI P/L",
      baseline: 0,
      textStyle: { fontSize: 9 }
    })
    .build();

  graphSheet.insertChart(iciciChart);

  SpreadsheetApp.flush();

  Logger.log("Final graph created successfully.");
  Logger.log("Rows used: " + (totalRows.length - 1));
  Logger.log("Source columns: A Date, F Total P/L, L Parag P/L, R ICICI P/L");
}


/************************************************************
 Manual test function
************************************************************/

function fixGraphSimpleNow() {
  buildGraphOnlyFinalV2();
}


/************************************************************
 Helpers
************************************************************/

function graphFinalNum_(value) {
  if (value === null || value === "") return 0;

  if (typeof value === "number") return value;

  var text = String(value)
    .replace(/₹/g, "")
    .replace(/,/g, "")
    .replace(/%/g, "")
    .replace(/\(/g, "-")
    .replace(/\)/g, "")
    .trim();

  var n = Number(text);

  return isNaN(n) ? 0 : n;
}


function graphFinalParseDate_(value) {
  if (value instanceof Date) {
    return new Date(value.getFullYear(), value.getMonth(), value.getDate());
  }

  if (!value) return null;

  var text = String(value).trim();

  var d = new Date(text);
  if (!isNaN(d.getTime())) {
    return new Date(d.getFullYear(), d.getMonth(), d.getDate());
  }

  var parts = text.split("-");

  if (parts.length !== 3) return null;

  var day = Number(parts[0]);
  var monthText = String(parts[1]).toLowerCase();
  var year = Number(parts[2]);

  if (year < 100) year = 2000 + year;

  var months = {
    jan: 0, january: 0,
    feb: 1, february: 1,
    mar: 2, march: 2,
    apr: 3, april: 3,
    may: 4,
    jun: 5, june: 5,
    jul: 6, july: 6,
    aug: 7, august: 7,
    sep: 8, sept: 8, september: 8,
    oct: 9, october: 9,
    nov: 10, november: 10,
    dec: 11, december: 11
  };

  var month;

  if (isNaN(Number(monthText))) {
    month = months[monthText];
  } else {
    month = Number(monthText) - 1;
  }

  if (isNaN(day) || month === undefined || isNaN(year)) return null;

  return new Date(year, month, day);
}
/************************************************************
 OLD WORKING GRAPH METHOD - FIXED FOR NEW A:S TRACKER

Source:
SIP Daily Tracker
A = Date
F = Total P/L
L = Parag P/L
R = ICICI P/L

Output:
Graph sheet with 3 charts only.
No extra helper sheet.
No hidden source columns.
************************************************************/


function FIX_buildGraph_OldWorkingMethod() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  var dataSheet = ss.getSheetByName("SIP Daily Tracker");

  if (!dataSheet) {
    throw new Error("Could not find sheet: SIP Daily Tracker");
  }

  var graphSheet = ss.getSheetByName("Graph");

  if (!graphSheet) {
    graphSheet = ss.insertSheet("Graph");
  }

  // Remove old charts
  var charts = graphSheet.getCharts();

  for (var i = 0; i < charts.length; i++) {
    graphSheet.removeChart(charts[i]);
  }

  // Clear graph sheet only
  graphSheet.clear();
  graphSheet.setHiddenGridlines(true);

  var lastRow = dataSheet.getLastRow();

  if (lastRow < 2) {
    throw new Error("SIP Daily Tracker has no data.");
  }

  // Use your old working direct range method
  var dateRange = dataSheet.getRange("A1:A" + lastRow);
  var totalPnlRange = dataSheet.getRange("F1:F" + lastRow);
  var paragPnlRange = dataSheet.getRange("L1:L" + lastRow);
  var iciciPnlRange = dataSheet.getRange("R1:R" + lastRow);

  var chartWidth = 1200;
  var chartHeight = 400;

  // Layout
  for (var c = 1; c <= 20; c++) {
    graphSheet.setColumnWidth(c, 80);
  }

  for (var r = 1; r <= 70; r++) {
    graphSheet.setRowHeight(r, 24);
  }

  graphSheet.getRange("A1:T70").setBackground("#ffffff");

  // Chart 1: Total P/L
  var totalChart = graphSheet.newChart()
    .setChartType(Charts.ChartType.COLUMN)
    .addRange(dateRange)
    .addRange(totalPnlRange)
    .setPosition(2, 2, 0, 0)
    .setOption("title", "Total P/L by Date")
    .setOption("hAxis.title", "Date")
    .setOption("vAxis.title", "P/L (₹)")
    .setOption("legend", { position: "none" })
    .setOption("width", chartWidth)
    .setOption("height", chartHeight)
    .build();

  graphSheet.insertChart(totalChart);

  // Chart 2: Parag P/L
  var paragChart = graphSheet.newChart()
    .setChartType(Charts.ChartType.COLUMN)
    .addRange(dateRange)
    .addRange(paragPnlRange)
    .setPosition(24, 2, 0, 0)
    .setOption("title", "Parag Parikh P/L by Date")
    .setOption("hAxis.title", "Date")
    .setOption("vAxis.title", "P/L (₹)")
    .setOption("legend", { position: "none" })
    .setOption("width", chartWidth)
    .setOption("height", chartHeight)
    .build();

  graphSheet.insertChart(paragChart);

  // Chart 3: ICICI P/L
  var iciciChart = graphSheet.newChart()
    .setChartType(Charts.ChartType.COLUMN)
    .addRange(dateRange)
    .addRange(iciciPnlRange)
    .setPosition(46, 2, 0, 0)
    .setOption("title", "ICICI Nifty 50 P/L by Date")
    .setOption("hAxis.title", "Date")
    .setOption("vAxis.title", "P/L (₹)")
    .setOption("legend", { position: "none" })
    .setOption("width", chartWidth)
    .setOption("height", chartHeight)
    .build();

  graphSheet.insertChart(iciciChart);

  SpreadsheetApp.flush();

  Logger.log("Graph rebuilt using old working method.");
  Logger.log("Source ranges used:");
  Logger.log("Dates: A1:A" + lastRow);
  Logger.log("Total P/L: F1:F" + lastRow);
  Logger.log("Parag P/L: L1:L" + lastRow);
  Logger.log("ICICI P/L: R1:R" + lastRow);
}


/************************************************************
 Override old graph test function
************************************************************/

function buildGraphOnlyFinalV2() {
  FIX_buildGraph_OldWorkingMethod();
}


function fixGraphSimpleNow() {
  FIX_buildGraph_OldWorkingMethod();
}


/************************************************************
 Sunday graph mail using old working chart method
************************************************************/

function sendSundayGraphMailFinalV2() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  // First rebuild graph using the old working method
  FIX_buildGraph_OldWorkingMethod();

  SpreadsheetApp.flush();
  Utilities.sleep(3000);

  var graphSheet = ss.getSheetByName("Graph");

  if (!graphSheet) {
    throw new Error("Graph sheet not found.");
  }

  var charts = graphSheet.getCharts();

  if (!charts || charts.length < 3) {
    throw new Error("Expected 3 charts in Graph sheet. Found: " + charts.length);
  }

  var totalChartImg = charts[0].getAs("image/png").setName("totalPnl.png");
  var paragChartImg = charts[1].getAs("image/png").setName("paragPnl.png");
  var iciciChartImg = charts[2].getAs("image/png").setName("iciciPnl.png");

  var todayText = Utilities.formatDate(
    new Date(),
    Session.getScriptTimeZone(),
    "dd-MMM-yyyy"
  );

  var subject = "📊 Weekly SIP P/L Graphs - " + todayText;

  var htmlBody =
    "<html><body style='font-family:Arial;background:#f5f6f8;padding:20px;'>" +
    "<div style='max-width:1100px;margin:auto;background:white;padding:20px;border-radius:14px;'>" +
    "<h2>📊 Weekly SIP P/L Graphs</h2>" +
    "<p style='color:#6b7280;'>Updated on " + todayText + "</p>" +

    "<h3>1. Total P/L by Date</h3>" +
    "<img src='cid:totalPnlImg' style='width:100%;max-width:1000px;height:auto;border:1px solid #e5e7eb;border-radius:10px;'><br><br>" +

    "<h3>2. Parag Parikh P/L by Date</h3>" +
    "<img src='cid:paragPnlImg' style='width:100%;max-width:1000px;height:auto;border:1px solid #e5e7eb;border-radius:10px;'><br><br>" +

    "<h3>3. ICICI Nifty 50 P/L by Date</h3>" +
    "<img src='cid:iciciPnlImg' style='width:100%;max-width:1000px;height:auto;border:1px solid #e5e7eb;border-radius:10px;'>" +

    "<p style='margin-top:20px;color:#6b7280;'>Best Regards,<br>Automated Portfolio Tracker</p>" +
    "</div></body></html>";

  MailApp.sendEmail({
    to: "diliptiwari1dkt@gmail.com",
    bcc: "diliptiwari12dkt@gmail.com,mehtrekrishna1996@gmail.com,shivani93808@gmail.com",
    subject: subject,
    htmlBody: htmlBody,
    inlineImages: {
      totalPnlImg: totalChartImg,
      paragPnlImg: paragChartImg,
      iciciPnlImg: iciciChartImg
    }
  });

  Logger.log("Sunday graph mail sent using old working graph method.");
}