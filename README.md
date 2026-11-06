# Task Tracking & Escalation Automation

A Google Sheets + Google Apps Script prototype for recurring task tracking, completion logging, correction history, manager notifications, and leadership escalation.

## What it demonstrates

- Records task completion against a task ID and assigned employee name.
- Validates required ticket proof for tasks marked as requiring a ticket.
- Uses a script lock to reduce collisions when multiple submissions update the spreadsheet.
- Creates a completion-log entry and calculates the next due date from the configured recurrence interval.
- Allows up to two corrections to the latest completion before the next due date, while preserving correction history.
- Sends manager alerts for tasks already marked `Overdue`.
- Supports manager acknowledgement from the `Manager Dashboard` sheet.
- Escalates unresolved manager alerts after a proposed four-hour prototype interval.

## Repository structure

```text
task-tracking-escalation-automation/
├── Code.gs
├── README.md
└── .gitignore
```

## Technology

- Google Sheets — workbook, task data, logs, dashboards, and alert status
- Google Apps Script — workflow automation and scheduled functions
- JavaScript — business logic
- Apps Script services — `SpreadsheetApp`, `LockService`, `MailApp`, and `PropertiesService`

## Workbook requirements

The script expects these sheet names:

- `Tasks`
- `Completion Log`
- `Correction History`
- `Manager Alerts`
- `Manager Dashboard`

The code reads/writes these minimum column positions. Keep the workbook's actual headers aligned with these positions before using the script.

### Tasks

| Column | Expected data |
|---|---|
| A | Task ID |
| B | Task name |
| C | Assigned employee |
| D | Recurrence interval in days |
| F | Ticket required (`Yes`/other) |
| G | Last completed date |
| H | Due date |
| I | Status (`Pending`/`Overdue`, etc.) |

Column E is read only indirectly through the row width and is not interpreted by the current completion logic.

### Completion Log

| Column | Expected data |
|---|---|
| A | Task ID |
| B | Task name |
| C | Employee |
| D | Completion timestamp |
| E | Proof / ticket reference |
| F | Acknowledgement field |
| G | Previous due date |
| H | Correction count |

### Correction History

| Column | Expected data |
|---|---|
| A | Task ID |
| B | Task name |
| C | Employee |
| D | Original completion timestamp |
| E | Previous proof |
| F | Correction timestamp |
| G | Replacement proof |

### Manager Alerts

The current code uses columns A, E, F, G, and H for task ID, alert timestamp, alert status, escalation status, and acknowledgement timestamp respectively. Confirm the remaining columns against your workbook.

### Manager Dashboard

The `onEdit(e)` handler expects task ID in column A and the manager acknowledgement dropdown in column F. It writes the acknowledgement status and timestamp back to `Manager Alerts`.

## Setup

1. Open the target Google Spreadsheet.
2. Select **Extensions → Apps Script**.
3. Replace or add the script source using `Code.gs`.
4. Save the Apps Script project.
5. In Apps Script, open **Project Settings → Script properties**.
6. Add a property:
   - Property: `MANAGER_EMAIL`
   - Value: the intended manager/leadership email address
7. Review the sheet names, column positions, sample data, and formulas.
8. Run `testConnection()` and authorize the requested Google permissions.
9. Configure time-driven triggers for:
   - `sendManagerOverdueAlerts`
   - `sendLeadershipEscalations`
10. Test on a duplicate workbook with dummy data before using any real operational data.

## Important implementation notes

- The employee check compares the submitted employee name with the assigned employee name in the sheet. It is a data validation check, **not identity authentication**.
- Proof is stored as submitted text. The script does not independently verify that a WhatsApp message was sent or that an external ticket exists.
- The four-hour escalation interval is a prototype rule in the code, not a confirmed organisational SLA.
- Manager acknowledgement is driven by a dropdown edit in the dashboard; it is not automatically inferred from email replies.
- `sendManagerOverdueAlerts()` only sends an alert when the task status is already `Overdue` and a matching row exists in `Manager Alerts`.
- The included validation functions use sample task IDs and employee names. Update those values to match a dedicated test workbook before running them.
- `testManagerEmail()` sends a real email to the configured `MANAGER_EMAIL`. Run it only when you intend to send a test message.
- Time-driven triggers and email sending require Apps Script authorization and may be subject to Google quotas.

## Safety and privacy

Do not commit real employee names, email addresses, ticket IDs, task records, spreadsheet links with private access, or operational data. Use a sanitised workbook with fictional sample records for demonstrations. The recipient email is read from Script Properties rather than stored in the repository.

## Current scope and limitations

This repository documents and shares the Apps Script prototype. It does not include the original Google Sheet's private data, formulas, dashboard formatting, or trigger configuration. Those must be recreated or shared separately using a sanitised demonstration workbook.

The prototype should not be described as production-ready without additional work such as stronger identity/access controls, role-based permissions, robust error/retry handling, idempotency safeguards, and end-to-end testing of every escalation path.
