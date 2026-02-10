📨 The "No-Nonsense" Liquidation & Overdue Tracker (v2.0)
Google Apps Script | Developed by Marcus Espiloy

I built this because, in finance, chasing receipts is a low-value use of time that carries high-value risks. Having handled pipelines with 75,000+ transaction lines, I know that manual follow-ups are where errors happen.

This system is my solution for Alphabet Soup Inc. (and any data-heavy finance team). It bridges the gap between raw financial pivots and employee inboxes, ensuring that every "nudge" is backed by actual data evidence.

🛠 Why I built it this way (My "Logic First" Approach):
Threaded Conversations: I hate cluttered inboxes. v2.0 logs Gmail Thread IDs so that every weekly reminder stays in one continuous conversation. It’s cleaner for the employee and creates a better audit trail for me.

Evidence-Based Nudging: I don’t just tell people they owe money; I show them the math. The script generates a custom CSV attachment for every recipient, filtered from the master data.

Accuracy Over Guesswork: It pulls directly from GMA and Overdue Pivots. If it's in the sheet, it’s in the email. If the balance is zero, the script is smart enough to skip them.

Safety & Compliance: I've included a TEST_MODE toggle and a Data Privacy Act disclaimer. As someone who handles P1.8M+ in advances, I never "go live" without a dry run.

📊 The Setup (The Technical Bit)
To keep the logic locked in, your Google Sheet needs:

GMA Summary & Overdue Summary: Your "Source of Truth" pivot tables.

DATA: The raw transaction dump (used for the CSV breakdowns).

Database: Your contact list. I’ve added support for multiple databases (Regions/HQ) because finance tools should be built to scale.

🚀 Getting it Running
Paste the code into your Apps Script project.

Check the SETTINGS block at the top—I’ve kept the sheet names and variables modular so you can tweak them without breaking the core engine.

Set a Time-driven trigger for Monday mornings. It’s one less thing for you to monitor.
