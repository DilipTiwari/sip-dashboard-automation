/**
 * 📱 WhatsApp Alert Script (Final Verified Version)
 * Includes Timestamp and Rate-Limit protection.
 */
function sendWhatsAppAlert() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("SIP");

  const invested = sheet.getRange("G3").getValue();
  const currentValue = sheet.getRange("G8").getValue();
  const gainLoss = sheet.getRange("G9").getValue();
  const status = sheet.getRange("G11").getValue();
  const lastStatus = sheet.getRange("H3").getValue(); 

  // 🚫 Spam Prevention: Only send if status changes
  if (status === lastStatus) {
    console.log("Status unchanged. No WhatsApp sent.");
    return;
  }

  // 🕒 Create Timestamp
  const timestamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "dd-MM-yy HH:mm:ss");

  const apiKey = "YG7ynLpVnBqx";
  const phoneNumbers = ["919700635190", "918309719690", "919381833391"];

  const message = 
    `*📊 Portfolio Alert*\n` +
    `*Time:* ${timestamp}\n` +
    `*Status:* ${status}\n` +
    `*Invested:* ₹${invested}\n` +
    `*Current:* ₹${currentValue}\n` +
    `*P/L:* ₹${gainLoss}`;

  phoneNumbers.forEach((phone, index) => {
    try {
      if (index > 0) {
        console.log("Waiting 5 seconds to avoid rate limit...");
        Utilities.sleep(5500); 
      }

      const url = `https://api.textmebot.com/send.php?phone=${phone}&apikey=${apiKey}&text=${encodeURIComponent(message)}`;
      const response = UrlFetchApp.fetch(url, { "muteHttpExceptions": true });
      
      if (response.getResponseCode() === 200) {
        console.log(`SUCCESS for ${phone}: ${response.getContentText()}`);
      }
    } catch (e) {
      console.error(`Error for ${phone}: ${e.message}`);
    }
  });

  sheet.getRange("H5").setValue(status);
}