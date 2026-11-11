# Task Tracking & Escalation Automation

A self-contained Google Sheets + Google Apps Script prototype for recurring task tracking, interactive employee submissions, completion logging, correction history, manager notifications, and leadership escalation.

## What It Demonstrates

* **One-Click Workbook Provisioning:** Automatically creates all 7 required tabs, headers, data validations, checkboxes, and live `FILTER`/`VLOOKUP` formulas via `setupAutomation()`.
* **Custom Spreadsheet Menu (`Allo Ops`):** Provides direct UI controls to repair sheets, install triggers, run manual checks, reset demo data, and execute automated tests.
* **Interactive Employee Submissions (`Employee View`):** Employees enter a ticket ID or check the completion box directly in the sheet to log work and advance the recurring due date.
* **Ticket Proof & Deduplication:** Validates required ticket IDs for portal tasks, prevents duplicate cycle completions, and uses `LockService` to prevent concurrent edit collisions.
* **Correction Tracking:** Allows up to 2 proof corrections per cycle while preserving full audit trails in `Correction History`.
* **Automated Escalation Pipeline:** Emails the manager once a task passes the configured `DUE_HOUR` (default 18:00), supports dropdown acknowledgements in the `Manager` tab, and escalates unresolved alerts to leadership after `ESCALATION_HOURS` (default 4 hours).
* **Built-In Test Suite:** Runs 8 end-to-end scenarios (`runAllTests()`) and outputs `PASS`/`FAIL` diagnostics to the `Test Log` tab.

## Repository Structure

```text
task-tracking-escalation-automation/
├── Code.gs
├── README.md
└── .gitignore
```
## Technology

* **Google Sheets** — Workbook storage, dynamic views, data validation, and interactive checkboxes
* **Google Apps Script** — Workflow automation, installable `onEdit` triggers, and hourly time-driven jobs
* **Apps Script Services** — `SpreadsheetApp`, `LockService`, `MailApp`, `ScriptApp`, and `Utilities`

## Workbook Architecture

You do **not** need to build the sheets manually. Running **Allo Ops → Set up / repair sheets** automatically creates and formats all seven tabs:

| Sheet Name | Purpose & Key Columns |
| :--- | :--- |
| **`Tasks`** | Master schedule (`Task ID`, `Task Name`, `Assigned Employee`, `Frequency (Days)`, `Task Type`, `Requires Ticket ID`, `Last Completed`, `Next Due`, `Status`). |
| **`Employee View`** | Live employee workspace synced via `FILTER` from `Tasks`, featuring `Last Proof`, `Enter Ticket ID / Update` (Col G), and `Mark Complete` checkboxes (Col H). |
| **`Manager`** | Live manager dashboard synced via `FILTER` from `Tasks`, featuring a `Manager Acknowledgement` dropdown (Col F). |
| **`Completion Log`** | Immutable audit log of completed task cycles, timestamps, submitted proof, cycle due dates, and correction counts. |
| **`Correction History`** | Audit trail recording original vs. replacement proof whenever an employee updates a completed cycle. |
| **`Manager Alerts`** | Tracks overdue alert timestamps, status (`Sent`, `Acknowledged`, `Resolved`), escalation status, and escalation timestamps. |
| **`Test Log`** | Stores `PASS`/`FAIL` outputs from the 8 automated test scenarios. |
## Setup Instructions

1. Create a new Google Spreadsheet.
2. Open **Extensions → Apps Script**.
3. Paste the contents of `Code.gs` into the script editor.
4. Update the configuration constants at the top of `Code.gs`:
   * `MANAGER_EMAIL` — Recipient email for overdue alerts
   * `LEADERSHIP_EMAIL` — Recipient email for SLA escalations
   * `ESCALATION_HOURS` — Hours before an uncompleted alert escalates (default: `4`)
   * `DUE_HOUR` — Hour of the day (24h format) when a due task becomes overdue (default: `18`)
   * `SEND_EMAILS` — Set to `true` to send live emails or `false` for dry-run logging
5. Save the project and reload the Google Spreadsheet.
6. In the spreadsheet menu bar, click **Allo Ops → Set up / repair sheets** and authorize the requested Google permissions.
7. *(Optional)* Click **Allo Ops → Reset demo data** to populate 5 sample tasks, then click **Allo Ops → Run all tests** to verify the installation.

## Important Implementation Notes

* **Installable Triggers Required:** Because `handleEdit` sends toast notifications and modifies multiple tabs, triggers must be installed via **Allo Ops → Install hourly trigger** (which registers both the hourly `runAutomation` job and the installable `handleEdit` trigger).
* **Cycle Deduplication:** Alerts and completions are deduplicated per task cycle using a composite key (`Task ID | yyyy-MM-dd`). Completing a task automatically marks any open alert for that cycle as `Resolved` and clears the manager acknowledgement dropdown.
* **Prototype Scope:** Employee matching relies on sheet row context rather than authenticated Google user identity. Before deploying in production, add `Session.getActiveUser().getEmail()` verification, sheet protection rules, and external secret management.

## Safety and Privacy

Do not commit real employee names, personal email addresses, internal ticket IDs, or private spreadsheet URLs to version control. Always replace `MANAGER_EMAIL` with a placeholder before pushing to GitHub.
