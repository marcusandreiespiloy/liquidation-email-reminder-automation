**Built for Finance Teams who value accuracy over manual follow-ups.**

I created this because chasing people for receipts is the least productive part of finance. I’ve found that automation is the only way to ensure nothing slips through the cracks. This script bridges the gap between your Pivot Table reports and your Contact Database to send automated, personalized reminders.

🛠 Why I built it this way:
Accuracy over guesswork: It pulls directly from the GMA Summary and Overdue Summary pivots. If the data is in the sheet, it’s in the email.

Respecting Inbox Space: I programmed it to skip anyone with a zero balance. We only send emails when there is an actual "Remaining" amount to discuss.

The "Safety First" Approach: I’ve included a TEST_MODE toggle. As someone who handles financial data, I know you never "go live" without a dry run.

Compliance-Ready: It automatically attaches the required Data Privacy Act disclaimer to keep everything above board.

📊 The Setup (The Technical Bit)
To keep the logic locked in, your Google Sheet should be structured like this:

GMA Summary: Pivot Table (Payee in Col A, Remaining Balance in Col E).

Overdue Summary: Pivot Table (Overdue Amount in Col B).

Database: Your "Source of Truth" for contacts. Ensure the Name column matches your Pivot names exactly.

🚀 Getting it Running
Paste the code into your Apps Script project.

Check the SETTINGS block—I’ve kept the sheet names modular so you can tweak them without breaking the core logic.

Run a test. Verify the numbers match your pivot.

Set it and forget it: I recommend a weekly Time-driven trigger. It’s one less thing for you to monitor.
