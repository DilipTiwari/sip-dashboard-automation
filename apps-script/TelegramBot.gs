function sendTelegramAlert() {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("SIP");
    
    // Fetch values and format them as Indian Currency
    const invested = sheet.getRange("G3").getValue().toLocaleString('en-IN');
    const currentValue = sheet.getRange("G8").getValue().toLocaleString('en-IN');
    const gainLoss = sheet.getRange("G9").getValue().toLocaleString('en-IN');
    const returnPct = sheet.getRange("G10").getValue();
    const status = sheet.getRange("G11").getValue();

    const BOT_TOKEN = "8519528895:AAFidrPAFEY86liJo-6nZ_70zov-q5x02LU";
    const CHAT_ID = "1847207454";

    // Improved Markdown message
    const message = 
      `📊 *Daily Portfolio Summary*\n\n` +
      `💰 *Invested:* ₹${invested}\n` +
      `📈 *Current Value:* ₹${currentValue}\n` +
      `📉 *Gain/Loss:* ₹${gainLoss}\n` +
      `🔄 *Return %:* ${returnPct}%\n` +
      `🚦 *Status:* ${status}`;

    const url = `https://api.telegram.org/bot${BOT_TOKEN}/sendMessage`;
    const payload = {
      chat_id: CHAT_ID,
      text: message,
      parse_mode: "Markdown"
    };

    const options = {
      method: "post",
      contentType: "application/json",
      payload: JSON.stringify(payload),
      muteHttpExceptions: true
    };

    const response = UrlFetchApp.fetch(url, options);
    console.log("Telegram Response: " + response.getContentText());

  } catch (e) {
    console.error("Telegram Alert Failed: " + e.toString());
  }
}