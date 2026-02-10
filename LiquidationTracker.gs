/**
 * Automated Liquidation & Overdue Tracker v2.0
 * * FEATURES:
 * - Email Threading: Automatically replies to previous email threads using saved Thread IDs.
 * - Dynamic CSV Generation: Filters raw data to attach a personalized breakdown for each payee.
 * - Multi-Database Support: Iterates through multiple contact sheets (e.g., Regional vs HQ).
 * - Audit-Ready: Logs Gmail Thread IDs back to the sheet for communication tracking.
 * * @author Marcus Espiloy
 * @license MIT
 */

// --- CONFIGURATION ---
const SETTINGS = {
  TEST_MODE: true,               // Set to FALSE for live production
  TEST_EMAIL: "admin@example.com", // Replace with your test email for verification
  
  SHEETS: {
    PIVOT_MAIN: "GMA Summary",   // Pivot table with total balances
    OVERDUE: "Overdue Summary",  // Pivot table with overdue balances
    RAW_DATA: "DATA",            // Source sheet for the CSV attachment breakdown
    DATABASES: ["Database", "Region Database"] // Add all contact sheet names here
  },
  
  // Data Index Mapping (0-based: A=0, B=1, C=2...)
  COL_INDEX: {
    NAME_IN_RAW: 2,   // Column C in 'DATA' sheet (Payee Name)
    STATUS_IN_RAW: 7  // Column H in 'DATA' sheet (Status)
  },
  
  EMAIL_SUBJECT_PREFIX: "2026 Weekly Liquidation Update",
  EMAIL_SIGNATURE: "<b>Marcus Espiloy</b><br>Finance Coordinator"
};

function sendLiquidationSummaryFromPivot() {
  const ss = SpreadsheetApp.getActive();
  const pivotSheet = ss.getSheetByName(SETTINGS.SHEETS.PIVOT_MAIN);
  const overdueSheet = ss.getSheetByName(SETTINGS.SHEETS.OVERDUE);
  const rawDataSheet = ss.getSheetByName(SETTINGS.SHEETS.RAW_DATA);

  // Safety check to ensure all required components exist
  if (!pivotSheet || !overdueSheet || !rawDataSheet) {
    Logger.log("Error: One or more required sheets are missing.");
    return;
  }

  const today = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "MMMM d, yyyy");
  const allRawData = rawDataSheet.getDataRange().getValues();

  // Load Pivot Data (Expected: A=Payee, E=Remaining Balance)
  const pivotData = pivotSheet.getRange(2, 1, pivotSheet.getLastRow() - 1, 5).getValues();
  const overdueData = overdueSheet.getRange(2, 1, overdueSheet.getLastRow() - 1, 2).getValues();

  // Iterate through each database sheet defined in SETTINGS
  SETTINGS.SHEETS.DATABASES.forEach(sheetName => {
    const dbSheet = ss.getSheetByName(sheetName);
    if (!dbSheet) return;

    // Database structure: A=Title, B=Name, C=Email, D=CC, E=ThreadID
    const dbData = dbSheet.getRange(2, 1, dbSheet.getLastRow() - 1, 5).getValues();

    dbData.forEach((dbRow, index) => {
      const [title, nameOnly, email, ccList, existingThreadId] = dbRow;

      if (!email || !nameOnly) return;

      // Match payee with Pivot Table data
      const pivotRow = pivotData.find(r => r[0] === nameOnly);
      if (!pivotRow || pivotRow[4] <= 0) return; // Skip if no balance is due

      const remaining = pivotRow[4];
      const overdueRow = overdueData.find(r => r[0] === nameOnly);
      const overdue = overdueRow ? overdueRow[1] : 0;
      const emailName = title ? `${title} ${nameOnly}` : nameOnly;

      // --- FILTER & GENERATE PERSONALIZED CSV BREAKDOWN ---
      const personalData = allRawData.filter((row, idx) => {
        if (idx === 0) return true; // Keep headers
        const isOwner = row[SETTINGS.COL_INDEX.NAME_IN_RAW] === nameOnly;
        const statusStr = String(row[SETTINGS.COL_INDEX.STATUS_IN_RAW]).toLowerCase();
        return isOwner && (statusStr.includes("pending") || statusStr.includes("overdue"));
      });

      let attachment = null;
      if (personalData.length > 1) {
        // Build CSV content with quote-wrapping for data safety
        const csvContent = personalData.map(row => {
          return row.map(cell => `"${String(cell).replace(/"/g, '""')}"`).join(",");
        }).join("\n");
        
        attachment = Utilities.newBlob(csvContent, "text/csv", `Pending_Items_${nameOnly}.csv`);
      }

      // --- EMAIL CONTENT (HTML) ---
      const subject = `${SETTINGS.EMAIL_SUBJECT_PREFIX} - ${nameOnly}`;
      const overdueStyle = overdue > 0 ? "color: red; font-weight: bold;" : "color: black;";
      const bodyHtml = `
        <p>Hi ${emailName},</p>
        <p>Please find the breakdown of your <b>Pending</b> and <b>Overdue</b> liquidations as of ${today}.</p>
        <table style="border-collapse: collapse; width: 300px; margin: 15px 0;">
          <tr><td style="padding: 5px; border: 1px solid #ddd;">Total Remaining:</td><td style="padding: 5px; border: 1px solid #ddd;"><b>${fmt(remaining)}</b></td></tr>
          <tr style="${overdueStyle}"><td style="padding: 5px; border: 1px solid #ddd;">Total Overdue:</td><td style="padding: 5px; border: 1px solid #ddd;"><b>${fmt(overdue)}</b></td></tr>
        </table>
        <p>Kindly submit receipts for the specific items listed in the attached CSV file.</p>
        <p>Regards,<br>${SETTINGS.EMAIL_SIGNATURE}</p>
      `;

      // --- DISPATCH LOGIC (Threading Support) ---
      try {
        const mailOptions = { htmlBody: bodyHtml, attachments: attachment ? [attachment] : [] };
        let threadFound = false;

        // Try to reply to an existing thread to prevent inbox clutter
        if (existingThreadId && !SETTINGS.TEST_MODE) {
          try {
            const thread = GmailApp.getThreadById(existingThreadId);
            if (thread) {
              thread.replyAll("", mailOptions);
              threadFound = true;
            }
          } catch (e) { 
            Logger.log(`Thread ID expired for ${nameOnly}. Starting new conversation.`); 
          }
        }

        // If no previous thread, send a new email and capture the Thread ID
        if (!threadFound) {
          const recipient = SETTINGS.TEST_MODE ? SETTINGS.TEST_EMAIL : email;
          if (ccList && !SETTINGS.TEST_MODE) mailOptions.cc = ccList;
          
          const msg = GmailApp.sendEmail(recipient, subject, "", mailOptions);
          
          if (!SETTINGS.TEST_MODE) {
            // Log the Thread ID back to Column E in the Database sheet
            dbSheet.getRange(index + 2, 5).setValue(msg.getThread().getId());
          }
        }
      } catch (err) { 
        Logger.log(`Critical Error for ${nameOnly}: ${err.toString()}`); 
      }
    });
  });
}

/**
 * Utility: Formats numbers to 2 decimal places with commas
 */
function fmt(num) {
  return Number(num).toLocaleString('en-US', { 
    minimumFractionDigits: 2, 
    maximumFractionDigits: 2 
  });
}
