function sendPortfolioAlert() {

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("SIP");
  const dailySheet = ss.getSheetByName("SIP Daily Tracker");

  if (!sheet || !dailySheet) return;

  // =============================
  // SUMMARY DATA
  // =============================

  const totalInvested = sheet.getRange("K3").getValue();
  const iciciUnits = sheet.getRange("K4").getValue();
  const paragUnits = sheet.getRange("K5").getValue();

  const totalCurrentValue = sheet.getRange("K8").getValue();
  const totalGainLoss = sheet.getRange("K9").getValue();
  const returnPercentage = sheet.getRange("K10").getValue();
  const status = sheet.getRange("K11").getValue();

  // =============================
  // FUND DATA
  // =============================

  const iciciCurrent = sheet.getRange("O8").getValue();
  const paragCurrent = sheet.getRange("O10").getValue();

  // =============================
  // INVESTED VALUE (FIXED)
  // =============================

  // ICICI = Column B (B2:B100)
  const iciciData = sheet.getRange("B2:B100").getValues();
  const iciciInvested = iciciData.flat().reduce((sum, val) => sum + (Number(val) || 0), 0);

  // PARAG = Column C (C2:C100)
  const paragData = sheet.getRange("C2:C100").getValues();
  const paragInvested = paragData.flat().reduce((sum, val) => sum + (Number(val) || 0), 0);

  // =============================
  // DAILY MOVEMENT
  // =============================

  const lastRow = dailySheet.getLastRow();
  let dailyMove = 0;

  if (lastRow > 1) {
    const yesterdayValue = dailySheet.getRange(lastRow, 3).getValue();
    dailyMove = totalCurrentValue - yesterdayValue;
  }

  const dailyPercent = totalCurrentValue !== 0
    ? (dailyMove / totalCurrentValue) * 100
    : 0;

  // =============================
  // COLORS
  // =============================

  const green = "#16a34a";
  const red = "#dc2626";

  const gainColor = totalGainLoss >= 0 ? green : red;
  const dayColor = dailyMove >= 0 ? green : red;

  // =============================
  // SUBJECT
  // =============================

  const subject =
    "📊 Portfolio Update (" +
    (returnPercentage * 100).toFixed(2) +
    "%)";

  // =============================
  // EMAIL BODY
  // =============================

  const body = `
  <html>
  <body style="font-family:Arial;background:#f5f6f8;padding:20px">

  <div style="max-width:700px;margin:auto;background:white;padding:20px;border-radius:14px;box-shadow:0 4px 12px rgba(0,0,0,0.08)">

    <h2 style="margin-bottom:10px">Investments (2)</h2>

    <p style="color:#6b7280;margin:0">Current value</p>
    <h1 style="margin:5px 0 15px 0">₹${totalCurrentValue.toLocaleString('en-IN')}</h1>

    <table style="width:100%;margin-bottom:15px">
      <tr style="color:#6b7280;font-size:13px">
        <td>Invested value</td>
        <td>1D returns</td>
        <td>Total returns</td>
        <td>XIRR</td>
      </tr>

      <tr style="font-weight:bold">
        <td>₹${totalInvested.toLocaleString('en-IN')}</td>

        <td style="color:${dayColor}">
          ${dailyMove >= 0 ? "+" : ""}₹${dailyMove.toLocaleString('en-IN')}
          <br><span style="font-size:12px">(${dailyPercent.toFixed(2)}%)</span>
        </td>

        <td style="color:${gainColor}">
          ${totalGainLoss >= 0 ? "+" : ""}₹${totalGainLoss.toLocaleString('en-IN')}
          <br><span style="font-size:12px">(${(returnPercentage * 100).toFixed(2)}%)</span>
        </td>

        <td>${(returnPercentage * 100).toFixed(2)}%</td>
      </tr>
    </table>

    <hr style="border:none;border-top:1px solid #e5e7eb;margin:15px 0">

    <table style="width:100%">

      <tr style="color:#6b7280;font-size:13px;text-align:left">
        <th>Scheme name</th>
        <th>Day change</th>
        <th>Returns</th>
        <th>Current (Invested)</th>
      </tr>

      <!-- PARAG -->
      <tr style="border-top:1px solid #f1f5f9">
        <td style="padding:10px 0">
          <div style="font-weight:600;color:#111827">
            Parag Parikh Flexi Cap Fund Direct Growth
          </div>
          <div style="font-size:12px;color:#6b7280">
            Units: ${paragUnits}
          </div>
        </td>

        <td style="color:${green}">--</td>

        <td style="color:${red}">
          ₹0<br>
          <span style="font-size:12px">(0%)</span>
        </td>

        <td>
          ₹${paragCurrent.toLocaleString('en-IN')}<br>
          <span style="font-size:12px;color:#6b7280">
            ₹${paragInvested.toLocaleString('en-IN')}
          </span>
        </td>
      </tr>

      <!-- ICICI -->
      <tr style="border-top:1px solid #f1f5f9">
        <td style="padding:10px 0">
          <div style="font-weight:600;color:#111827">
            ICICI Prudential Nifty 50 Index Direct Plan Growth
          </div>
          <div style="font-size:12px;color:#6b7280">
            Units: ${iciciUnits}
          </div>
        </td>

        <td style="color:${green}">--</td>

        <td style="color:${red}">
          ₹0<br>
          <span style="font-size:12px">(0%)</span>
        </td>

        <td>
          ₹${iciciCurrent.toLocaleString('en-IN')}<br>
          <span style="font-size:12px;color:#6b7280">
            ₹${iciciInvested.toLocaleString('en-IN')}
          </span>
        </td>
      </tr>

    </table>

    <p style="margin-top:20px;color:#6b7280;font-size:13px">
      Status: <b>${status}</b><br>
      Last Updated: ${new Date().toLocaleDateString('en-IN')}
    </p>

  </div>

  </body>
  </html>
  `;

  // =============================
  // SEND EMAIL
  // =============================

  MailApp.sendEmail({
    to: "diliptiwari1dkt@gmail.com",
    bcc: "diliptiwari12dkt@gmail.com,mehtrekrishna1996@gmail.com,shivani93808@gmail.com",
    subject: subject,
    htmlBody: body
  });

}