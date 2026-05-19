/******************** CONFIG ********************/
var SD_DASHBOARD_SHEET = "SIP Dashboard";
var SD_TRACKER_SHEET = "SIP Daily Tracker";
var SD_WEEKLY_SHEET = "Weekly SIP Data";

var SD_SUNDAY_EMAIL_HOUR = 9;

var SD_MAIL_TO = "diliptiwari12dkt@gmail.com";
var SD_MAIL_BCC = "mehrtekrishna1996@gmail.com,shivani93808@gmail.com";

/******************** MENU ********************/
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("SIP Dashboard")
    .addItem("Refresh Dashboard", "refreshDashboard")
    .addItem("Send Dashboard Mail Now", "sendWeeklyDashboardMail")
    .addItem("Create Weekly Data Sheet", "createWeeklyDataSheet")
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
  sdCreateSection_(dashboard, "B32:F33", "📅 Monthly SIP Tracker");
  sdCreateSection_(dashboard, "N32:X33", "🥧 Portfolio Allocation");

  sdCreatePanel_(dashboard, "B18:L29");
  sdCreatePanel_(dashboard, "N18:X29");
  sdCreatePanel_(dashboard, "N34:U43");

  sdCreateSipTable_(dashboard);

  SpreadsheetApp.flush();

  var pnlBlob = sdCreatePnLChartBlob_(allRows, 540, 230);
  var growthBlob = sdCreateGrowthChartBlob_(allRows, 540, 230);
  var pieBlob = sdCreatePieChartBlob_(430, 180);

  dashboard.insertImage(pnlBlob, 2, 18, 6, 6);
  dashboard.insertImage(growthBlob, 14, 18, 6, 6);
  dashboard.insertImage(pieBlob, 14, 34, 6, 6);

  SpreadsheetApp.flush();

  if (!silent) {
    SpreadsheetApp.getUi().alert("✅ Dashboard refreshed successfully.");
  }
}

function refreshDashboard() {
  setupUltraDashboard(false);
}

/******************** SUNDAY MAIL - PREVIOUS MONDAY TO SATURDAY ********************/
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
  sdWriteWeeklyDataSheet_(weeklyRows);

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

/******************** EMAIL HTML ********************/
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

          <table cellpadding="4" cellspacing="0" style="border-collapse:collapse;background:#ffffff;font-size:12px;text-align:center;">
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

/******************** MONTHLY SIP HTML ********************/
function sdSipRowsHtml_() {
  var rows = [
    ["Oct-25", "₹5,000", "Completed", ""],
    ["Nov-25", "₹5,000", "Completed", ""],
    ["Dec-25", "₹5,000", "Completed", ""],
    ["Jan-26", "₹5,000", "Completed", ""],
    ["Feb-26", "₹5,000", "Completed", ""],
    ["Mar-26", "₹5,000", "Completed", ""],
    ["Apr-26", "₹5,000", "Completed", ""],
    ["May-26", "₹5,000", "Completed", ""],
    ["Jun-26", "₹5,000", "Pending", ""]
  ];

  return rows.map(function(r) {
    return `
      <tr>
        <td style="border:1px solid #111827;">${r[0]}</td>
        <td style="border:1px solid #111827;">${r[1]}</td>
        <td style="border:1px solid #111827;">${r[2]}</td>
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

/******************** WEEKLY DATA SHEET ********************/
function createWeeklyDataSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var tracker = ss.getSheetByName(SD_TRACKER_SHEET);

  if (!tracker) {
    throw new Error("Sheet not found: " + SD_TRACKER_SHEET);
  }

  var weeklyRows = sdGetPreviousMondayToSaturdayRows_(tracker, new Date());
  sdWriteWeeklyDataSheet_(weeklyRows);

  SpreadsheetApp.getUi().alert("✅ Weekly data sheet updated successfully.");
}

function sdWriteWeeklyDataSheet_(rows) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var weeklySheet = sdGetOrCreateSheet_(ss, SD_WEEKLY_SHEET);

  weeklySheet.clear();

  weeklySheet.getRange("A1:E1").setValues([
    ["Date", "Total Invested", "Current Value", "Daily P/L", "Return %"]
  ]);

  if (rows && rows.length) {
    weeklySheet.getRange(2, 1, rows.length, 5).setValues(rows);
  }

  weeklySheet.getRange("A:A").setNumberFormat("dd-mmm-yyyy");
  weeklySheet.getRange("B:D").setNumberFormat("₹#,##0.00");
  weeklySheet.getRange("E:E").setNumberFormat("0.00%");
  weeklySheet.autoResizeColumns(1, 5);
}

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
    var d = new Date(r[0]);
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

  sheet.getRange("A1:Z70").breakApart();
  sheet.clear();
  sheet.setHiddenGridlines(true);

  for (var c = 1; c <= 26; c++) {
    sheet.setColumnWidth(c, 52);
  }

  for (var r = 1; r <= 60; r++) {
    sheet.setRowHeight(r, 22);
  }

  sheet.getRange("A1:Z60").setBackground("#dfe3ea");
}

/******************** DATA ********************/
function sdGetCleanTrackerRows_(tracker) {
  var lastRow = tracker.getLastRow();

  if (lastRow < 2) {
    return [];
  }

  var values = tracker.getRange(2, 1, lastRow - 1, 7).getValues();

  return values
    .filter(function(row) {
      return row[0] !== "" && row[1] !== "" && row[2] !== "" && row[3] !== "";
    })
    .map(function(row) {
      return [
        row[0],
        Number(row[1]) || 0,
        Number(row[2]) || 0,
        Number(row[3]) || 0,
        sdNormalizePercent_(row[4])
      ];
    });
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
    .setOption("chartArea", { left: 10, top: 25, width: "85%", height: "75%" })
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

function sdCreateSipTable_(sheet) {
  var data = [
    ["Month", "Investment", "Status", "Remarks"],
    ["Oct-25", "₹5,000", "Completed", ""],
    ["Nov-25", "₹5,000", "Completed", ""],
    ["Dec-25", "₹5,000", "Completed", ""],
    ["Jan-26", "₹5,000", "Completed", ""],
    ["Feb-26", "₹5,000", "Completed", ""],
    ["Mar-26", "₹5,000", "Completed", ""],
    ["Apr-26", "₹5,000", "Completed", ""],
    ["May-26", "₹5,000", "Completed", ""],
    ["Jun-26", "₹5,000", "Pending", ""]
  ];

  sheet.getRange("B34:E43").setValues(data);

  sheet.getRange("B34:E34")
    .setBackground("#08152f")
    .setFontColor("white")
    .setFontWeight("bold");

  sheet.getRange("B34:E43")
    .setBorder(true, true, true, true, true, true)
    .setHorizontalAlignment("center")
    .setVerticalAlignment("middle")
    .setFontSize(8);
}

/******************** COMMON HELPERS ********************/
function sdGetOrCreateSheet_(ss, name) {
  return ss.getSheetByName(name) || ss.insertSheet(name);
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

function sdFormatPercentText_(value) {
  var num = Number(value) || 0;
  return (num * 100).toFixed(2) + "%";
}

function sdFormatDateShort_(value) {
  return Utilities.formatDate(
    new Date(value),
    Session.getScriptTimeZone(),
    "dd-MMM"
  );
}

function sdFormatDateFull_(value) {
  return Utilities.formatDate(
    new Date(value),
    Session.getScriptTimeZone(),
    "dd-MMM-yyyy"
  );
}