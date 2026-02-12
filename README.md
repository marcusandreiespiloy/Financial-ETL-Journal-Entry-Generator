Financial ETL: Automated Double-Entry Journal Generator
A high-integrity automation engine for converting operational logs into balanced, audit-ready accounting entries.

🚀 Business Impact
In traditional accounting, converting raw operational reports into General Ledger entries is a high-risk, manual hurdle. This engine was designed to:

Reduce Turnaround Time by 99%: Processes thousands of lines in seconds.

Eliminate Manual Errors: Built-in validation ensures zero "out-of-balance" journal entries.

Standardize Workflows: Maps legacy data (CSV/Excel) directly to ERP-ready templates (Xero/QuickBooks).

🛠️ Technical Architecture
This repository features a Google Apps Script implementation of an ETL (Extract, Transform, Load) pipeline.

1. Transformation Logic
Currency-Safe Math: Implements precise numeric rounding to handle JavaScript’s floating-point limitations, preventing 1-cent imbalances.

Dynamic Credit Routing: Uses string-matching algorithms to intelligently determine credit accounts (e.g., distinguishing between different bank accounts vs. Accounts Payable) based on payee and transaction metadata.

2. Security & Integrity (Audit Clauses)
Deep-Clear Safeguards: Every execution flushes the destination template to prevent "ghost data" overlap from previous reporting periods.

Final Balancing Audit: A post-transformation check sums all debits and credits; if an imbalance exists, the system halts and flags the discrepancy before data is committed.

3. Optimized Performance
Array-Buffering: Instead of writing to the sheet row-by-row, the script batches all entries into an output array and executes a single .setValues() call. This maximizes execution speed for datasets exceeding 10,000 rows.

📂 Project Structure
JournalEntryGenerator.gs: The core engine containing the transformation and validation logic.

Sample_Log.csv: (Reference) Example of raw transactional data.

Template_Output.xlsx: (Reference) The standardized format for accounting software uploads.

🔧 Setup & Usage
Copy the code from JournalEntryGenerator.gs into your Google Apps Script editor.

Ensure your source data is in a sheet named Sheet3 and your template is named TEMPLATE.

Run generateJournalEntries().

Provide the starting JE sequence number when prompted.
