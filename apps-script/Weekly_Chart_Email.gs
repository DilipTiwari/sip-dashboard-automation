/******************** WEEKLY SIP DAILY TRACKER GRAPH EMAIL ********************/
/*
  This script is separate from SIP Dashboard.gs.

  Purpose:
  - Read data from "SIP Daily Tracker"
  - Create Daily Profit/Loss graph
  - Send only that graph in email body
  - Install Sunday trigger separately

  It does NOT touch your SIP Dashboard design/mail.
*/

/******************** CONFIG ********************/
var WC_TRACKER_SHEET = "SIP Daily Tracker";

var WC_MAIL_TO = "diliptiwari12dkt@gmail.com";
var WC_MAIL_BCC = "mehrtekrishna1996@gmail.com,shivani93808@gmail.com";

var WC_SUNDAY_EMAIL_HOUR = 9;

// Use this for graph type:
// "ALL_TIME" = full SIP Daily Tracker graph
// "PREVIOUS_WEEK" = previous Monday to Saturday only
var WC_GRAPH_MODE = "ALL_TIME";

/******************** MENU ********************/
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("Weekly SIP Graph")
    .addItem("Send SIP Tracker Graph Now", "sendWeeklySipTrackerGraphMail")
    .addItem("Install Sunday SIP Tracker Graph Trigger", "installSundaySipTrackerGraphTrigger")
    .addToUi();
}

/******************** MAIN MAIL FUNCTION ********************/
function sendWeeklySipTrackerGraphMail() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var tracker = ss.getSheetByName(WC_TRACKER_SHEET);

  if (!tracker) {
    throw new Error("Sheet not found: " + WC_TRACKER_SHEET);
  }

  var allRows = wcGetCleanTrackerRows_(tracker);

  if (!allRows.length) {
    throw new Error("No valid data found in " + WC_TRACKER_SHEET);
  }

  var graphRows;

  if (WC_GRAPH_MODE === "PREVIOUS_WEEK") {
    graphRows = wcGetPreviousMondayToSaturdayRows_(allRows, new Date());
  } else {
    graphRows = allRows;
  }

  if (!graphRows.length) {
    throw new Error("No graph data found for selected mode: " + WC_GRAPH_MODE);
  }

  var latest = allRows[allRows.length - 1];
  var latestReturn = Number(latest[4]) || 0;

  var weekLabel = wcGetPreviousWeekLabel_(new Date());

  var titleText =
    WC_GRAPH_MODE === "PREVIOUS_WEEK"
      ? "Previous Monday to Saturday P/L Trend - " + weekLabel
      : "All Time Daily Profit/Loss Trend";

  var chartBlob = wcCreateDailyPnLChartBlob_(graphRows, 900, 420, titleText)
    .setName("sip_daily_tracker_graph.png");

  var subject =
    WC_GRAPH_MODE === "PREVIOUS_WEEK"
      ? "📊 Weekly SIP Daily Tracker Graph - " + weekLabel
      : "📊 SIP Daily Tracker Graph Update (" + wcFormatPercentText_(latestReturn) + ")";

  var htmlBody = wcBuildGraphOnlyEmailHtml_(titleText);

  MailApp.sendEmail({
    to: WC_MAIL_TO,
    bcc: WC_MAIL_BCC,
    subject: subject,
    htmlBody: htmlBody,
    inlineImages: {
      trackerGraph: chartBlob
    }
  });

  Logger.log("SIP Daily Tracker graph mail sent successfully.");
}

/******************** EMAIL HTML ********************/
function wcBuildGraphOnlyEmailHtml_(titleText) {
  return `
  <div style="margin:0;padding:24px;background:#f3f4f6;font-family:Arial,Helvetica,sans-serif;">
    <table role="presentation" cellpadding="0" cellspacing="0" style="width:100%;max-width:980px;margin:0 auto;border-collapse:collapse;background:#ffffff;border-radius:10px;overflow:hidden;">
      
      <tr>
        <td style="background:#08152f;color:#ffffff;text-align:center;font-size:18px;font-weight:bold;padding:18px 10px;">
          📊 ${titleText}
        </td>
      </tr>

      <tr>
        <td style="background:#1f2937;color:#ffffff;text-align:center;font-size:11px;font-weight:bold;padding:6px 10px;">
          Sent on : ${Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "dd-MMM-yyyy hh:mm a")}
        </td>
      </tr>

      <tr>
        <td style="padding:18px;text-align:center;background:#ffffff;">
          <img src="cid:trackerGraph" style="width:100%;max-width:900px;height:auto;display:block;margin:0 auto;border:1px solid #d1d5db;">
        </td>
      </tr>

    </table>
  </div>
  `;
}

/******************** SUNDAY TRIGGER ********************/
function installSundaySipTrackerGraphTrigger() {
  var triggers = ScriptApp.getProjectTriggers();

  triggers.forEach(function(trigger) {
    if (trigger.getHandlerFunction() === "sendWeeklySipTrackerGraphMail") {
      ScriptApp.deleteTrigger(trigger);
    }
  });

  ScriptApp.newTrigger("sendWeeklySipTrackerGraphMail")
    .timeBased()
    .onWeekDay(ScriptApp.WeekDay.SUNDAY)
    .atHour(WC_SUNDAY_EMAIL_HOUR)
    .create();

  SpreadsheetApp.getUi().alert("✅ Sunday SIP Daily Tracker graph trigger installed successfully.");
}

/******************** DATA READ ********************/
function wcGetCleanTrackerRows_(tracker) {
  var lastRow = tracker.getLastRow();

  if (lastRow < 2) {
    return [];
  }

  /*
    SIP Daily Tracker structure:
    A = Date
    B = Total Invested
    C = Current Value
    D = Daily P/L
    E = Daily P/L %
    F = Day
    G = Week
  */
  var values = tracker.getRange(2, 1, lastRow - 1, 7).getValues();

  return values
    .filter(function(row) {
      return row[0] !== "" && row[3] !== "";
    })
    .map(function(row) {
      return [
        row[0],
        Number(row[1]) || 0,
        Number(row[2]) || 0,
        Number(row[3]) || 0,
        wcNormalizePercent_(row[4])
      ];
    });
}

/******************** CHART CREATION ********************/
function wcCreateDailyPnLChartBlob_(rows, width, height, titleText) {
  var dataTableBuilder = Charts.newDataTable()
    .addColumn(Charts.ColumnType.STRING, "Date")
    .addColumn(Charts.ColumnType.NUMBER, "Daily P/L");

  rows.forEach(function(row) {
    dataTableBuilder.addRow([
      wcFormatDateShort_(row[0]),
      Number(row[3]) || 0
    ]);
  });

  var xAxisFontSize = rows.length > 25 ? 6 : 9;
  var xAxisAngle = rows.length > 12 ? 90 : 45;

  var chart = Charts.newColumnChart()
    .setDataTable(dataTableBuilder.build())
    .setDimensions(width, height)
    .setOption("title", titleText)
    .setOption("backgroundColor", "#ffffff")
    .setOption("legend", { position: "none" })
    .setOption("colors", ["#0f3b5f"])
    .setOption("chartArea", {
      left: 75,
      top: 45,
      width: "84%",
      height: "62%"
    })
    .setOption("hAxis", {
      slantedText: true,
      slantedTextAngle: xAxisAngle,
      textStyle: {
        fontSize: xAxisFontSize
      }
    })
    .setOption("vAxis", {
      title: "Profit / Loss",
      textStyle: {
        fontSize: 10
      }
    })
    .build();

  return chart.getBlob();
}

/******************** PREVIOUS MONDAY TO SATURDAY FILTER ********************/
function wcGetPreviousMondayToSaturdayRows_(allRows, refDate) {
  var today = new Date(refDate);
  var day = today.getDay(); // Sunday = 0

  var sunday = new Date(today);
  sunday.setDate(today.getDate() - day);
  sunday.setHours(0, 0, 0, 0);

  var monday = new Date(sunday);
  monday.setDate(sunday.getDate() - 6);
  monday.setHours(0, 0, 0, 0);

  var saturday = new Date(sunday);
  saturday.setDate(sunday.getDate() - 1);
  saturday.setHours(23, 59, 59, 999);

  return allRows.filter(function(row) {
    var d = new Date(row[0]);
    return d >= monday && d <= saturday;
  });
}

function wcGetPreviousWeekLabel_(refDate) {
  var today = new Date(refDate);
  var day = today.getDay();

  var sunday = new Date(today);
  sunday.setDate(today.getDate() - day);
  sunday.setHours(0, 0, 0, 0);

  var monday = new Date(sunday);
  monday.setDate(sunday.getDate() - 6);

  var saturday = new Date(sunday);
  saturday.setDate(sunday.getDate() - 1);

  return wcFormatDateFull_(monday) + " to " + wcFormatDateFull_(saturday);
}

/******************** HELPERS ********************/
function wcNormalizePercent_(value) {
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

function wcFormatPercentText_(value) {
  var num = Number(value) || 0;
  return (num * 100).toFixed(2) + "%";
}

function wcFormatDateShort_(value) {
  return Utilities.formatDate(
    new Date(value),
    Session.getScriptTimeZone(),
    "dd-MM-yy"
  );
}

function wcFormatDateFull_(value) {
  return Utilities.formatDate(
    new Date(value),
    Session.getScriptTimeZone(),
    "dd-MMM-yyyy"
  );
}