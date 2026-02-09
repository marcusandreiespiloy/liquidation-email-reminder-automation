/**
 * Automated Liquidation & Overdue Tracker
 * * This script synchronizes pivot table data with a contact database 
 * to send personalized financial summaries via Gmail.
 * * @author Marcus Espiloy
 * @license MIT
 */

// --- CONFIGURATION ---
const SETTINGS = {
  TEST_MODE: true,               // Set to FALSE when ready for live deployment
  TEST_EMAIL: "your-test@email.com", 
  
  SHEETS: {
    PIVOT_MAIN: "GMA Summary",   // Source of Disbursed/Liquidated data
    OVERDUE: "Overdue Summary",  // Source of Overdue amounts
    DATABASE: "Database"         // Source of Name, Email, and CC info
  },
  
  EMAIL_SUBJECT: "Weekly Liquidation Summary Update",
  CURRENCY_LOCALE: "en-US"
};

function sendLiquidationSummaryFromPivot() {
  const ss = SpreadsheetApp.getActive();

  // --- Initialize Sheets ---
  const pivotSheet = ss.getSheetByName(SETTINGS.SHEETS.PIVOT_MAIN);
  const overdueSheet = ss.getSheetByName(SETTINGS.SHEETS.OVERDUE);
  const dbSheet = ss.getSheetByName(SETTINGS.SHEETS.DATABASE);

  // --- Safety checks ---
  if (!pivotSheet || !overdueSheet || !dbSheet) { 
    Logger.log("Error: One or more required sheets are missing."); 
    return; 
  }

  const today = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "MMMM d, yyyy");

  // --- Read Main Pivot Data (A=Payee, B=Disbursed, C=Cash Returned, D=Liquidated, E=Remaining) ---
  const pivotData = pivotSheet.getRange(2, 1, pivotSheet.getLastRow() - 1, 5).getValues(); 

  // --- Read Overdue Pivot Data (A=Payee, B=Overdue Amount) ---
  const overdueData = overdueSheet.getRange(2, 1, overdueSheet.getLastRow() - 1, 2).getValues(); 

  // --- Read Database (A=Title, B=Name, C=Email, D=CC) ---
  const dbData = dbSheet.getRange(2, 1, dbSheet.getLastRow() - 1, 4).getValues(); 

  // --- Legal Signature Template ---
  const htmlSignature = `
    <hr>
    <p style="font-size: 10px; color: #666;">
      This e-mail message, including any attached file, is confidential and legally privileged... 
      [Data Privacy Act Compliance Text]
    </p>`;

  // --- Process Each User in Database ---
  dbData.forEach(dbRow => {
    const title = dbRow[0];      
    const nameOnly = dbRow[1];   
    const email = dbRow[2];
    const ccList = dbRow[3];

    if (!email) return; 

    // --- Find payee in Main Pivot ---
    const pivotRow = pivotData.find(r => r[0] === nameOnly);
    if (!pivotRow) {
      Logger.log("Payee not found in main pivot: " + nameOnly);
      return;
    }

    const [payee, disbursed, cashReturned, liquidated, remaining] = pivotRow;

    if (remaining <= 0) return; 

    // --- Find payee in Overdue Pivot ---
    const overdueRow = overdueData.find(r => r[0] === nameOnly);
    const overdue = overdueRow ? overdueRow[1] : 0;

    const emailName = title ? `${title} ${nameOnly}` : nameOnly;

    // --- Construct Email Body ---
    const bodyText = `Hi ${emailName},

As of ${today} for CVs released starting Oct 1 to Present:

Total Disbursed: ${fmt(disbursed)}
Total Liquidated: ${fmt(liquidated)}
Total Cash Return: ${fmt(cashReturned)}
Total Remaining Liquidation: ${fmt(remaining)}
Total Overdue: ${fmt(overdue)}

Kindly submit receipts once available.

Thank you!

Regards,
[Your Name/Department Name]
`;

    const bodyHtml = bodyText.replace(/\n/g, '<br>') + htmlSignature;

    // --- Dispatch Email ---
    if (SETTINGS.TEST_MODE) {
      MailApp.sendEmail({
        to: SETTINGS.TEST_EMAIL,
        subject: `[TEST] ${SETTINGS.EMAIL_SUBJECT}`,
        htmlBody: bodyHtml
      });
      Logger.log("Test email sent for: " + nameOnly);
      return;
    }

    const options = { htmlBody: bodyHtml };
    if (ccList && ccList.trim() !== "") {
      options.cc = ccList;
    }

    MailApp.sendEmail(email, SETTINGS.EMAIL_SUBJECT, bodyText, options);
    Logger.log("Email sent to: " + email);
  });
}

/**
 * Formats numbers to currency string
 */
function fmt(num) {
  return Number(num).toLocaleString(SETTINGS.CURRENCY_LOCALE, { 
    minimumFractionDigits: 2, 
    maximumFractionDigits: 2 
  });
}
