function recordDailyTracker() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sipSheet = ss.getSheetByName("SIP");
  const tracker = ss.getSheetByName("SIP Daily Tracker");

  if (!sipSheet || !tracker) return;

  const tz = Session.getScriptTimeZone();
  let today = new Date();

  const lastRow = tracker.getLastRow();

  /* Prevent duplicate entry */
  if (lastRow > 1) {
    const lastDateRaw = tracker.getRange(lastRow, 1).getValue();

    if (lastDateRaw) {
      const lastDateObj = new Date(lastDateRaw);

      lastDateObj.setHours(0, 0, 0, 0);
      today.setHours(0, 0, 0, 0);

      if (lastDateObj.getTime() === today.getTime()) return;
    }
  }

  /* Fetch data */
  const invested = sipSheet.getRange("K3").getValue();
  const current = sipSheet.getRange("K8").getValue();
  const pl = sipSheet.getRange("K9").getValue();
  const plPct = sipSheet.getRange("K10").getValue();

  if (!invested || !current) return;

  /* Day + Week */
  const dayName = Utilities.formatDate(today, tz, "EEEE");
  const weekNum = getISOWeekNumber(today);

  /* Append row */
  tracker.appendRow([
    today,
    invested,
    current,
    pl,
    plPct,
    dayName,
    weekNum
  ]);

  const newRow = tracker.getLastRow();

  /* Formatting */
  tracker.getRange(newRow, 1).setNumberFormat("dd-MM-yy");
  tracker.getRange(newRow, 2, 1, 3).setNumberFormat("0");
  tracker.getRange(newRow, 5).setNumberFormat("0.00%");

  /* Update Dashboard */
  updateDashboard();
}


/* ISO Week */
function getISOWeekNumber(date) {
  const tempDate = new Date(date);
  tempDate.setHours(0, 0, 0, 0);

  tempDate.setDate(tempDate.getDate() + 3 - ((tempDate.getDay() + 6) % 7));

  const week1 = new Date(tempDate.getFullYear(), 0, 4);

  return (
    1 +
    Math.round(
      ((tempDate - week1) / 86400000 - 3 + ((week1.getDay() + 6) % 7)) / 7
    )
  );
}


/* DASHBOARD */
function updateDashboard() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const tracker = ss.getSheetByName("SIP Daily Tracker");
  let dashboard = ss.getSheetByName("Dashboard");

  if (!tracker) return;

  if (!dashboard) {
    dashboard = ss.insertSheet("Dashboard");
  }

  dashboard.clear();

  const data = tracker.getDataRange().getValues();
  const rows = data.slice(1);

  let totalInvested = 0;
  let latestValue = 0;

  let weekMap = {};
  let dayMap = {};

  /* ✅ FILTER ONLY SIP (≤ 50K) */
  rows.forEach(row => {
    const invested = row[1];

    if (invested > 50000) return; // ❌ remove chit

    const current = row[2];
    const pl = row[3];
    const day = row[5];
    const week = row[6];

    if (invested) totalInvested = invested;
    if (current) latestValue = current;

    if (!weekMap[week]) weekMap[week] = 0;
    weekMap[week] += pl;

    if (!dayMap[day]) dayMap[day] = { pl: 0, count: 0 };
    dayMap[day].pl += pl;
    dayMap[day].count++;
  });

  const totalPL = latestValue - totalInvested;

  let r = 1;

  dashboard.getRange(r++, 1).setValue("📊 SIP DASHBOARD");

  dashboard.getRange(r, 1).setValue("Total Invested");
  dashboard.getRange(r++, 2).setValue(totalInvested);

  dashboard.getRange(r, 1).setValue("Current Value");
  dashboard.getRange(r++, 2).setValue(latestValue);

  dashboard.getRange(r, 1).setValue("Total P/L");
  dashboard.getRange(r++, 2).setValue(totalPL);

  r += 2;

  /* Weekly */
  dashboard.getRange(r++, 1).setValue("📅 Weekly Performance");
  dashboard.getRange(r++, 1, 1, 2).setValues([["Week", "Total P/L"]]);

  Object.keys(weekMap).sort((a,b)=>a-b).forEach(week => {
    dashboard.getRange(r, 1).setValue(week);
    dashboard.getRange(r++, 2).setValue(weekMap[week]);
  });

  r += 2;

  /* Day Analysis */
  dashboard.getRange(r++, 1).setValue("📆 Day Analysis");
  dashboard.getRange(r++, 1, 1, 3).setValues([["Day", "Avg P/L", "Total P/L"]]);

  Object.keys(dayMap).forEach(day => {
    const avg = dayMap[day].pl / dayMap[day].count;

    dashboard.getRange(r, 1).setValue(day);
    dashboard.getRange(r, 2).setValue(avg);
    dashboard.getRange(r++, 3).setValue(dayMap[day].pl);
  });

  /* ✅ Format decimals */
  dashboard.getRange("B2:B100").setNumberFormat("0.00");
  dashboard.getRange("C2:C100").setNumberFormat("0.00");

  dashboard.autoResizeColumns(1, 5);

  /* ✅ FILTERED CHART */
  const lastRow = tracker.getLastRow();
  const raw = tracker.getRange("A2:G" + lastRow).getValues();

  const filtered = raw.filter(r => r[1] <= 50000);

  if (filtered.length > 0) {
    const temp = ss.insertSheet("TEMP_CHART");

    temp.getRange(1,1,filtered.length,3)
      .setValues(filtered.map(r => [r[0], r[2], r[1]]));

    const chart = dashboard.newChart()
      .setChartType(Charts.ChartType.LINE)
      .addRange(temp.getRange(1,1,filtered.length,3))
      .setPosition(2, 5, 0, 0)
      .build();

    dashboard.insertChart(chart);

    ss.deleteSheet(temp);
  }
}