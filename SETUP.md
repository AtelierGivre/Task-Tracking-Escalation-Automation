# Runbook — install, demo, submit

Everything below happens inside **your** Google account. The Apps Script project belongs to the
account that creates it, so it cannot be pushed from another machine.

## 1. Install (5 minutes)

1. Create a Google Sheet named **Allo Recurring Task Tracker**.
2. **File → Settings → Time zone → (GMT+05:30) India Standard Time**. The 18:00 due time is read
   in the script timezone.
3. **Extensions → Apps Script**. Delete whatever is in `Code.gs` and paste this folder's
   `Code.gs`. Save.
4. In the function dropdown pick **setupAutomation** → **Run**.
5. Approve the permissions (Sheets, Gmail send, triggers). It is your own script; the prompt
   appears because it sends mail and installs triggers.
6. Reload the spreadsheet. An **Allo Ops** menu appears.

`setupAutomation` builds/repairs all tabs, the live views, the status formulas, the Read Me tab,
the conditional colours, the hourly trigger and the on-edit trigger.

If you ever see `#REF!` in `Manager` or `Employee View`, run **Allo Ops → Set up / repair sheets**
again. The old filter area is cleared before the new one is written, so this is now safe to run
repeatedly.

## 2. Seed the demo (30 seconds)

**Allo Ops → Reset demo data**

Creates the four brief tasks plus one deliberately overdue case:

| Task | Owner | Frequency | Due |
|---|---|---|---|
| T001 Daily Operations Update | Employee A | 1 day | today |
| T002 Portal Task A | Employee B | 7 days | +7 |
| T003 Portal Task B | Employee C | 15 days | +15 |
| T004 Portal Task C | Employee D | 30 days | +30 |
| T005 Overdue Demo | Employee E | 1 day | −2 (overdue) |

## 3. Manual demo (the order that reads best)

1. **Pending Now** — only T001 (due today) and T005 (overdue) appear.
2. **Employee View**, T002 row, column G → type `TK-2001`, press Enter.
   Toast: *Completion recorded. Next due …*; `Tasks` moves 7 days out; `Completion Log` gains one
   row; column G clears itself. No other tab was touched → Tests 1 + 2.
3. Same tab, T001 row, tick column H. Completion logged with `WhatsApp update` → Test 3.
4. **Allo Ops → Run pending checks now**. `Manager Alerts` gets a T005 row (`Sent`) and the
   manager mailbox gets *Overdue Task Alert*. Same-day alert → Test 4.
5. **Manager** tab, T005 row, column F → `Acknowledged`. `Manager Alerts!F` becomes
   `Acknowledged`, `H` is timestamped → Test 5.
6. **Allo Ops → Run all tests**. It backdates one alert to prove the 4-hour SLA and writes the
   PASS/FAIL grid to **Test Log** → Tests 6, 7, 8.
7. Screenshot: Test Log, Manager Alerts, Completion Log, one `Tasks` row before/after, and both
   emails.

`Run all tests` sends a handful of real emails (alert, escalation). That is intentional evidence;
run it once, not in a loop. To rehearse without mail, set `SEND_EMAILS = false` at the top of
`Code.gs`, run, then set it back to `true` and run once for the final screenshots.

## 4. Sharing for submission

1. **Share → General access → Anyone with the link → Editor** (the brief explicitly asks for
   editor access so they can test the automation).
2. Also share the Apps Script project? Not needed: editor access to the Sheet exposes
   **Extensions → Apps Script**. Attach `Code.gs` to the email as a backup.
3. Open the link in an incognito window once to confirm it is not restricted.
4. Send to `hr.ops.allo@gmail.com` with the draft in `SUBMISSION_EMAIL.md`.

## 5. Troubleshooting

| Symptom | Fix |
|---|---|
| Menu missing after reload | Run `onOpen` once from the editor, or refresh the tab. |
| `#REF!` in Manager / Employee View | Allo Ops → Set up / repair sheets. |
| Checkbox does nothing | The on-edit trigger is missing: Allo Ops → Install hourly trigger. |
| No email | Check spam; run `testManagerEmail`; confirm `MANAGER_EMAIL` is correct. |
| Alert not created for T005 | The due time is 18:00; before that the task is not yet late. Run the test suite or backdate `Tasks!H` by one day. |
| Wrong dates by a few hours | Sheet timezone is not IST (File → Settings). |
| Duplicate alerts after editing data | Should not happen: key = Task ID + cycle due date. If it does, capture the two `Manager Alerts` rows and check the due dates differ. |
