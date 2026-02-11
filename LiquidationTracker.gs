/**
 * Automated Liquidation & Overdue Tracker v2.0
 * Features: Email Threading, Dynamic CSV Attachments, Multi-DB Support.
 */

// --- GLOBAL SETTINGS ---
const TEST_MODE = false; // SET TO FALSE FOR FRIDAY LAUNCH
const TEST_EMAIL = "marcus@asoup.ph"; 

function sendLiquidationSummaryFromPivot() {
  const ss = SpreadsheetApp.getActive();
  const pivotSheet = ss.getSheetByName("GMA Summary"); 
  const overdueSheet = ss.getSheetByName("Overdue Summary"); 
  const rawDataSheet = ss.getSheetByName("DATA"); 
  const databaseNames = ["Database", "Region Database"];

  if (!pivotSheet || !overdueSheet || !rawDataSheet) {
    Logger.log("Required sheets missing.");
    return;
  }

  const today = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "MMMM d, yyyy");
  const allRawData = rawDataSheet.getDataRange().getValues();
  const NAME_COL_INDEX = 2;   // Column C (Payee Name)
  const STATUS_COL_INDEX = 7; // Column H (Status)

  const pivotData = pivotSheet.getRange(2, 1, pivotSheet.getLastRow() - 1, 5).getValues(); 
  const overdueData = overdueSheet.getRange(2, 1, overdueSheet.getLastRow() - 1, 2).getValues(); 

  databaseNames.forEach(sheetName => {
    const dbSheet = ss.getSheetByName(sheetName);
    if (!dbSheet) return;

    const dbData = dbSheet.getRange(2, 1, dbSheet.getLastRow() - 1, 5).getValues(); 

    dbData.forEach((dbRow, index) => {
      const [title, nameOnly, email, ccList, existingThreadId] = dbRow;

      if (!email || !nameOnly) return; 

      const pivotRow = pivotData.find(r => r[0] === nameOnly);
      if (!pivotRow || pivotRow[4] <= 0) return; 

      const remaining = pivotRow[4];
      const overdueRow = overdueData.find(r => r[0] === nameOnly);
      const overdue = overdueRow ? overdueRow[1] : 0;
      const emailName = title ? `${title} ${nameOnly}` : nameOnly;

      // Filter raw data for CSV
      const personalData = allRawData.filter((row, idx) => {
        if (idx === 0) return true; 
        const isOwner = row[NAME_COL_INDEX] === nameOnly;
        const statusStr = String(row[STATUS_COL_INDEX]).toLowerCase();
        return isOwner && (statusStr.includes("pending") || statusStr.includes("overdue"));
      });
      
      let attachment = null;
      if (personalData.length > 1) {
        const csvContent = personalData.map(row => {
          return row.map(cell => `"${String(cell).replace(/"/g, '""')}"`).join(",");
        }).join("\n");
        attachment = Utilities.newBlob(csvContent, "text/csv", `Pending_Liquidations_${nameOnly}.csv`);
      }

      const subject = `2026 Weekly Liquidation Summary Update - ${nameOnly}`;
      const overdueStyle = overdue > 0 ? `color: red; font-weight: bold;` : `color: black;`;
      const bodyHtml = `
        <p>Hi ${emailName},</p>
        <p>Please see the attached breakdown of your <b>Pending</b> and <b>Overdue</b> liquidations as of ${today}.</p>
        <table style="border-collapse: collapse; width: 300px; margin: 15px 0;">
          <tr><td style="padding: 5px; border: 1px solid #ddd;">Total Remaining:</td><td style="padding: 5px; border: 1px solid #ddd;"><b>${fmt(remaining)}</b></td></tr>
          <tr style="${overdueStyle}"><td style="padding: 5px; border: 1px solid #ddd;">Total Overdue:</td><td style="padding: 5px; border: 1px solid #ddd;"><b>${fmt(overdue)}</b></td></tr>
        </table>
        <p>Kindly submit receipts for the items listed in the attached file once available.</p>
        <p>Thank you!<br><br>Regards,<br><b>Marcus Espiloy</b></p>
      `;

      try {
        const mailOptions = { htmlBody: bodyHtml, attachments: attachment ? [attachment] : [] };
        let threadFound = false;
        
        if (existingThreadId) {
          try {
            const thread = GmailApp.getThreadById(existingThreadId);
            if (thread) {
              thread.reply("", mailOptions); 
              threadFound = true;
              Logger.log(`SUCCESS: Replied to thread for ${nameOnly}`);
            }
          } catch (e) { Logger.log(`Thread ID expired for ${nameOnly}`); }
        }

        if (!threadFound) {
          const recipient = TEST_MODE ? TEST_EMAIL : email;
          if (ccList && !TEST_MODE) mailOptions.cc = ccList;
          
          GmailApp.sendEmail(recipient, subject, "", mailOptions);
          
          // Capture and save the new thread ID
          const latestThread = GmailApp.search(`to:${recipient} subject:"${subject}"`, 0, 1)[0];
          dbSheet.getRange(index + 2, 5).setValue(latestThread.getId());
          Logger.log(`SUCCESS: New thread started for ${nameOnly}`);
        }
      } catch (err) { Logger.log(`Error for ${nameOnly}: ${err.toString()}`); }
    });
  });
}

function fmt(num) {
  return Number(num).toLocaleString('en-US', { minimumFractionDigits: 2, maximumFractionDigits: 2 });
}
