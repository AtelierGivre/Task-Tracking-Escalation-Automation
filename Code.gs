const MANAGER_EMAIL = "manager@example.com"; // set the manager / HR inbox here
const LEADERSHIP_EMAIL = MANAGER_EMAIL;
const ESCALATION_HOURS = 4;
const DUE_HOUR = 18;
const MAX_CORRECTIONS = 2;
const SEND_EMAILS = true;

const TASKS = "Tasks";
const LOG = "Completion Log";
const HISTORY = "Correction History";
const ALERTS = "Manager Alerts";
const MANAGER = "Manager";
const VIEW = "Employee View";
const PENDING = "Pending Now";
const TESTLOG = "Test Log";
const README = "Read Me";

// Adds the Allo Ops menu to the spreadsheet.
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("Allo Ops")
    .addItem("Set up / repair sheets", "setupAutomation")
    .addItem("Install hourly trigger", "installTriggers")
    .addSeparator()
    .addItem("Run pending checks now", "runAutomation")
    .addItem("Run all tests", "runAllTests")
    .addItem("Reset demo data", "resetDemoData")
    .addToUi();
}

// One-time setup: repair the sheets and install the triggers.
function setupAutomation() {
  repairSheets();
  installTriggers();
  SpreadsheetApp.getActive().toast("Sheets repaired and hourly trigger installed.", "Allo Ops", 6);
}

// Rebuilds headers, status formulas, and the live views.
function repairSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const ensure = function (name, headers) {
    let sh = ss.getSheetByName(name);
    if (!sh) sh = ss.insertSheet(name);
    const width = Math.max(sh.getLastColumn(), headers.length);
    sh.getRange(1, 1, 1, width).clearContent();
    sh.getRange(1, 1, 1, headers.length).setValues([headers])
      .setFontWeight("bold").setBackground("#1f3864").setFontColor("#ffffff");
    sh.setFrozenRows(1);
    return sh;
  };

  const tasks = ensure(TASKS, ["Task ID", "Task Name", "Assigned Employee", "Frequency (Days)",
    "Task Type", "Requires Ticket ID", "Last Completed", "Next Due", "Status"]);
  const log = ensure(LOG, ["Task ID", "Task Name", "Completed By", "Completion Date & Time",
    "Proof / Ticket ID", "Manager Notified", "Cycle Due Date", "Correction Count"]);
  ensure(HISTORY, ["Task ID", "Task Name", "Employee Name", "Original Completion Date & Time",
    "Original Proof / Ticket ID", "Correction Date & Time", "Corrected Proof / Ticket ID",
    "Correction Count"]);
  const alerts = ensure(ALERTS, ["Task ID", "Task Name", "Assigned Employee", "Cycle Due Date",
    "Alert Date & Time", "Alert Status", "Escalation Status", "Acknowledged At", "Escalated At"]);
  const manager = ensure(MANAGER, ["Task ID", "Task Name", "Assigned Employee", "Next Due",
    "Status", "Manager Acknowledgement"]);
  const view = ensure(VIEW, ["Employee Name", "Task ID", "Task Name", "Due Date", "Status",
    "Last Proof", "Enter Ticket ID / Update", "Mark Complete"]);
  const pending = ensure(PENDING, ["Employee Name", "Task ID", "Task Name", "Due Date", "Status"]);
  ensure(TESTLOG, ["Test", "Scenario", "Expected", "Actual", "Result", "Run At"]);

  // Legacy files kept a FILTER in Alerts!A2; it would block the script's writes.
  if (alerts.getRange("A2").getFormula()) {
    alerts.getRange(2, 1, Math.max(alerts.getMaxRows() - 1, 1), 9).clearContent();
  }

  // Tasks: keep the configured due date, give a brand-new task its first cycle.
  const lastTask = Math.max(tasks.getLastRow(), 2);
  const ids = tasks.getRange(2, 1, lastTask - 1, 1).getValues();
  const dueValues = tasks.getRange(2, 8, lastTask - 1, 1).getValues();
  const freqValues = tasks.getRange(2, 4, lastTask - 1, 1).getValues();
  const today = startOfDay_(new Date());
  const dueOut = [];
  const statusFormulas = [];

  for (let i = 0; i < ids.length; i++) {
    if (!String(ids[i][0] || "").trim()) break;
    const r = i + 2;
    const due = dueValues[i][0];
    if (due instanceof Date) {
      dueOut.push([due]);
    } else {
      const freq = Number(freqValues[i][0]) || 1;
      dueOut.push([new Date(today.getTime() + freq * 86400000)]);
    }
    statusFormulas.push(['=IF($A' + r + '="","",IF($H' + r + '<TODAY(),"Overdue",' +
      'IF($H' + r + '=TODAY(),"Due today","Scheduled")))']);
  }

  if (dueOut.length) {
    tasks.getRange(2, 8, dueOut.length, 1).setValues(dueOut).setNumberFormat("yyyy-mm-dd");
    tasks.getRange(2, 9, statusFormulas.length, 1).setFormulas(statusFormulas);
  }
  tasks.getRange(2, 6, Math.max(lastTask - 1, 1), 1).setDataValidation(
    SpreadsheetApp.newDataValidation().requireValueInList(["Yes", "No"], true).build());

  // Manager: clear the old spill before writing FILTER, or it returns #REF!.
  manager.getRange(2, 1, Math.max(manager.getMaxRows() - 1, 1), 5).clearContent();
  manager.getRange("A2").setFormula(
    '=IFERROR(FILTER({Tasks!A2:A1000,Tasks!B2:B1000,Tasks!C2:C1000,Tasks!H2:H1000,Tasks!I2:I1000},' +
    'Tasks!A2:A1000<>""),"")');
  manager.getRange(2, 4, Math.max(manager.getMaxRows() - 1, 1), 1).setNumberFormat("yyyy-mm-dd");
  manager.getRange(2, 6, Math.max(manager.getMaxRows() - 1, 1), 1).setDataValidation(
    SpreadsheetApp.newDataValidation().requireValueInList(["Acknowledged"], true).build());

  // Employee View: clear the old spill, keep the G/H input columns.
  view.getRange(2, 1, Math.max(view.getMaxRows() - 1, 1), 6).clearContent();
  view.getRange("A2").setFormula(
    '=IFERROR(FILTER({Tasks!C2:C1000,Tasks!A2:A1000,Tasks!B2:B1000,Tasks!H2:H1000,Tasks!I2:I1000},' +
    'Tasks!A2:A1000<>""),"")');
  view.getRange("F2").setFormula(
    '=IF(B2:B1000="","",IFNA(VLOOKUP(B2:B1000,SORT(\'Completion Log\'!A2:E1000,4,FALSE),5,FALSE),""))');
  view.getRange(2, 4, Math.max(view.getMaxRows() - 1, 1), 1).setNumberFormat("yyyy-mm-dd");
  view.getRange(2, 8, 100, 1).insertCheckboxes();

  // Pending Now: only the rows that need action today.
  pending.getRange(2, 1, Math.max(pending.getMaxRows() - 1, 1), 5).clearContent();
  pending.getRange("A2").setFormula(
    '=IFERROR(QUERY(\'Employee View\'!A2:E1000,"select A,B,C,D,E where E=\'Overdue\' or ' +
    'E=\'Due today\' order by A, D",0),"")');

  styleSheets_(view, manager, alerts);
  writeReadMe_(ss);
  SpreadsheetApp.flush();
}

// Colours overdue / due-today rows so the demo reads at a glance.
// Rules must be built per sheet: one rule cannot hold ranges from two sheets.
function styleSheets_(view, manager, alerts) {
  const statusRules_ = function (sheet) {
    const last = Math.max(sheet.getMaxRows(), 2);
    const range = sheet.getRange(2, 5, last - 1, 1);
    return [
      SpreadsheetApp.newConditionalFormatRule()
        .whenTextEqualTo("Overdue").setBackground("#fce8e6").setFontColor("#a50e0e").setBold(true)
        .setRanges([range]).build(),
      SpreadsheetApp.newConditionalFormatRule()
        .whenTextEqualTo("Due today").setBackground("#fef7e0").setFontColor("#b06000").setBold(true)
        .setRanges([range]).build(),
    ];
  };

  view.clearConditionalFormatRules();
  view.setConditionalFormatRules(statusRules_(view));
  manager.clearConditionalFormatRules();
  manager.setConditionalFormatRules(statusRules_(manager));

  const lastAlertRow = Math.max(alerts.getMaxRows(), 2);
  alerts.clearConditionalFormatRules();
  alerts.setConditionalFormatRules([
    SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo("Escalated").setBackground("#fce8e6").setFontColor("#a50e0e").setBold(true)
      .setRanges([alerts.getRange(2, 7, lastAlertRow - 1, 1)]).build(),
  ]);
}

// Writes (or refreshes) the one-page operating guide, so an evaluator can use the sheet cold.
function writeReadMe_(ss) {
  const sh = ss.getSheetByName(README) || ss.insertSheet(README, 0);
  sh.clear();

  const rows = [
    ["Allo Innoware - Recurring Task Tracker", "Google Sheets + Apps Script prototype"],
    ["", ""],
    ["EMPLOYEE - one place only", "Open 'Employee View' and find your name."],
    ["Portal task", "Paste the Ticket ID in column G and press Enter. That single entry completes the cycle - no other sheet to update."],
    ["Daily task", "Post the update in the WhatsApp group, then tick column H (Mark Complete). The proof is logged automatically as 'WhatsApp update'."],
    ["What happens next", "A Completion Log row is written, the task leaves 'Pending Now', and the next due date is calculated from the frequency."],
    ["", ""],
    ["MANAGER", "'Manager' tab = live status of every task. 'Pending Now' = only today's open items, grouped by employee."],
    ["Acknowledge", "After the 18:00 alert email, choose 'Acknowledged' in column F. It is timestamped in 'Manager Alerts'."],
    ["Escalation", "If the task is still open 4 hours after the alert, leadership is emailed once and 'Manager Alerts'!G shows Escalated."],
    ["", ""],
    ["AUTOMATION", "Hourly trigger runAutomation - nobody has to open the sheet."],
    ["No duplicates", "One alert per task per cycle, one escalation per alert, one completion per cycle."],
    ["Setup", "Allo Ops > Set up / repair sheets, then Allo Ops > Reset demo data."],
    ["Evidence", "Allo Ops > Run all tests writes PASS/FAIL to the 'Test Log' tab."],
  ];

  sh.getRange(1, 1, rows.length, 2).setValues(rows);
  sh.getRange("A1:B1").setFontSize(14).setFontWeight("bold");
  sh.getRange("A3:A15").setFontWeight("bold");
  sh.getRange("B2:B15").setWrap(true);
  sh.setColumnWidth(1, 210);
  sh.setColumnWidth(2, 760);
  sh.setHiddenGridlines(true);
  if (sh.getMaxColumns() > 2) sh.hideColumns(3, sh.getMaxColumns() - 2);
}

// Installs the hourly run and the on-edit trigger.

function installTriggers() {
  ScriptApp.getProjectTriggers().forEach(function (t) {
    const handler = t.getHandlerFunction();
    if (handler === "runAutomation" || handler === "onEdit" || handler === "handleEdit" ||
      handler === "sendManagerOverdueAlerts" || handler === "sendLeadershipEscalations") {
      ScriptApp.deleteTrigger(t);
    }
  });
  ScriptApp.newTrigger("runAutomation").timeBased().everyHours(1).create();

  ScriptApp.newTrigger("handleEdit")
    .forSpreadsheet(SpreadsheetApp.getActive())
    .onEdit()
    .create();
}

// Hourly job: send overdue alerts, then escalations.
function runAutomation() {
  const alerts = sendManagerOverdueAlerts();
  const escalations = sendLeadershipEscalations();
  Logger.log("Automation run: " + alerts + " alerts, " + escalations + " escalations.");
  return { alerts: alerts, escalations: escalations };
}

// Records one completion or correction for a task cycle.
function recordCompletion(taskId, employeeName, proof) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);

  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const tasks = ss.getSheetByName(TASKS);
    const log = ss.getSheetByName(LOG);
    const history = ss.getSheetByName(HISTORY);

    if (!tasks || !log || !history) throw new Error("A required sheet is missing.");

    const taskRow = findTaskRow_(tasks, taskId);
    if (taskRow === -1) throw new Error("Task ID not found: " + taskId);

    const task = tasks.getRange(taskRow, 1, 1, 9).getValues()[0];
    const assigned = String(task[2] || "").trim();

    if (assigned !== String(employeeName || "").trim()) {
      throw new Error("Employee name does not match the assigned employee.");
    }

    const requiresTicket = String(task[5] || "").trim().toLowerCase() === "yes";
    const text = String(proof == null ? "" : proof).trim();
    if (requiresTicket && !text) throw new Error("A ticket ID is required for this task.");

    const cycleDue = task[7] instanceof Date ? task[7] : new Date(task[7]);
    if (isNaN(cycleDue.getTime())) throw new Error("The task due date is invalid.");

    const submitted = text || "WhatsApp update";
    const now = new Date();
    const key = cycleKey_(taskId, cycleDue);
    const existing = findLogRow_(log, key);

    if (existing) {
      const currentProof = String(existing.values[4] || "").trim();

      if (currentProof === submitted) {
        return "Already completed for this cycle. Nothing changed.";
      }

      const count = Number(existing.values[7]) || 0;
      if (count >= MAX_CORRECTIONS) {
        throw new Error("Maximum of " + MAX_CORRECTIONS + " corrections already used.");
      }

      history.appendRow([
        taskId, task[1], assigned, existing.values[3], currentProof, now, submitted, count + 1,
      ]);
      log.getRange(existing.row, 5).setValue(submitted);
      log.getRange(existing.row, 8).setValue(count + 1);
      SpreadsheetApp.flush();
      resolveAlerts_(ss, taskId, cycleDue);
      return "Correction saved (" + (count + 1) + " of " + MAX_CORRECTIONS + ").";
    }

    log.appendRow([taskId, task[1], assigned, now, submitted, "No", cycleDue, 0]);

    const nextDue = new Date(now.getTime() + (Number(task[3]) || 1) * 86400000);
    tasks.getRange(taskRow, 7).setValue(now);
    tasks.getRange(taskRow, 8).setValue(nextDue);

    SpreadsheetApp.flush();
    resolveAlerts_(ss, taskId, cycleDue);
    return "Completion recorded. Next due " + ymd_(nextDue) + ".";
  } finally {
    lock.releaseLock();
  }
}

// Returns the Tasks row for a task ID, or -1.
function findTaskRow_(sheet, taskId) {
  const last = sheet.getLastRow();
  if (last < 2) return -1;
  const ids = sheet.getRange(2, 1, last - 1, 1).getValues();
  for (let i = 0; i < ids.length; i++) {
    if (String(ids[i][0]).trim() === String(taskId).trim()) return i + 2;
  }
  return -1;
}

// Returns the log row for a task ID and cycle due date.
function findLogRow_(sheet, key) {
  const last = sheet.getLastRow();
  if (last < 2) return null;
  const rows = sheet.getRange(2, 1, last - 1, 8).getValues();
  for (let i = 0; i < rows.length; i++) {
    const id = String(rows[i][0] || "").trim();
    const due = rows[i][6];
    if (!id || !(due instanceof Date)) continue;
    if (cycleKey_(id, due) === key) return { row: i + 2, values: rows[i] };
  }
  return null;
}

// Marks this cycle's alert Resolved and clears the manager tick.
function resolveAlerts_(ss, taskId, due) {
  const alerts = ss.getSheetByName(ALERTS);
  if (!alerts || alerts.getLastRow() < 2) return;
  const rows = alerts.getRange(2, 1, alerts.getLastRow() - 1, 9).getValues();
  for (let i = 0; i < rows.length; i++) {
    const id = String(rows[i][0] || "").trim();
    const alertDue = rows[i][3];
    if (id === String(taskId).trim() && alertDue instanceof Date && ymd_(alertDue) === ymd_(due)) {
      alerts.getRange(i + 2, 6).setValue("Resolved");
    }
  }

  clearManagerTick_(ss, taskId);
}

// Clears the manager acknowledgement tick for a task.
function clearManagerTick_(ss, taskId) {
  const manager = ss.getSheetByName(MANAGER);
  const tasks = ss.getSheetByName(TASKS);
  if (!manager || !tasks) return;
  const taskRow = findTaskRow_(tasks, taskId);
  if (taskRow === -1 || taskRow > manager.getMaxRows()) return;
  manager.getRange(taskRow, 6).clearContent();
}

// Records a manager acknowledgement with a timestamp.
function acknowledgeTask_(taskId) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const tasks = ss.getSheetByName(TASKS);
  const alerts = ss.getSheetByName(ALERTS);
  if (!tasks || !alerts) return false;

  const taskRow = findTaskRow_(tasks, taskId);
  if (taskRow === -1) return false;

  const due = tasks.getRange(taskRow, 8).getValue();
  if (!(due instanceof Date)) return false;

  const hit = findAlertRow_(alerts, taskId, due);
  if (!hit) return false;

  alerts.getRange(hit.row, 6).setValue("Acknowledged");
  alerts.getRange(hit.row, 8).setValue(new Date());
  return true;
}

// Returns the alert row for a task ID and cycle due date.
function findAlertRow_(sheet, taskId, due) {
  const last = sheet.getLastRow();
  if (last < 2) return null;
  const rows = sheet.getRange(2, 1, last - 1, 9).getValues();
  for (let i = 0; i < rows.length; i++) {
    const id = String(rows[i][0] || "").trim();
    const alertDue = rows[i][3];
    if (!id || !(alertDue instanceof Date)) continue;
    if (id === String(taskId).trim() && ymd_(alertDue) === ymd_(due)) {
      return { row: i + 2, values: rows[i] };
    }
  }
  return null;
}

// Emails the manager once per overdue task cycle.
function sendManagerOverdueAlerts() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const tasks = ss.getSheetByName(TASKS);
  const alerts = ss.getSheetByName(ALERTS);
  if (!tasks || !alerts) throw new Error("Tasks or Manager Alerts sheet is missing.");

  const lastRow = tasks.getLastRow();
  if (lastRow < 2) return 0;

  const data = tasks.getRange(2, 1, lastRow - 1, 9).getValues();
  const now = new Date();
  let sent = 0;

  data.forEach(function (task) {
    const id = String(task[0] || "").trim();
    if (!id) return;

    const due = task[7] instanceof Date ? task[7] : null;
    if (!due) return;

    // Due time is 18:00 in the sheet's timezone, so the alert lands on the due date.
    const dueDay = ymd_(due);
    const today = ymd_(now);
    if (today < dueDay) return;
    if (today === dueDay && Number(Utilities.formatDate(now, sheetTz_(), "H")) < DUE_HOUR) return;

    if (findAlertRow_(alerts, id, due)) return;

    alerts.appendRow([id, task[1], task[2], due, now, "Sent", "", "", ""]);

    const body = [
      "Hello Manager,",
      "",
      "The following recurring task is overdue:",
      "",
      "Task ID:          " + id,
      "Task:             " + task[1],
      "Assigned to:      " + task[2],
      "Cycle due:        " + ymd_(due),
      "",
      "Ask the employee to complete it in the Employee View tab, then tick",
      "Manager Acknowledgement in the Manager tab.",
      "",
      "If it is still open after " + ESCALATION_HOURS + " hours it is escalated automatically.",
      "",
      "Allo Innoware Task Tracker",
    ].join("\n");

    sendMail_(MANAGER_EMAIL, "Overdue Task Alert: " + task[1], body);
    sent++;
  });

  SpreadsheetApp.flush();
  return sent;
}

// Emails leadership when an alert passes the SLA window.
function sendLeadershipEscalations() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const alerts = ss.getSheetByName(ALERTS);
  if (!alerts) throw new Error("Manager Alerts sheet is missing.");
  if (alerts.getLastRow() < 2) return 0;

  const rows = alerts.getRange(2, 1, alerts.getLastRow() - 1, 9).getValues();
  const now = new Date();
  let sent = 0;

  rows.forEach(function (row, i) {
    const id = String(row[0] || "").trim();
    if (!id) return;

    const status = String(row[5] || "");
    const escalated = String(row[6] || "");
    const alertAt = row[4] instanceof Date ? row[4] : null;

    if (!alertAt) return;
    if (!(row[3] instanceof Date)) return;
    if (status === "Resolved" || escalated === "Escalated") return;

    const hours = (now - alertAt) / 3600000;
    if (hours < ESCALATION_HOURS) return;

    const body = [
      "Hello,",
      "",
      "This recurring task is still unresolved:",
      "",
      "Task ID:      " + id,
      "Task:         " + row[1],
      "Assigned to:  " + row[2],
      "Cycle due:    " + ymd_(row[3]),
      "Manager alert sent: " + ymd_(alertAt),
      "Manager acknowledged: " + (status === "Acknowledged" ? "Yes" : "No"),
      "",
      "It has been open for " + Math.round(hours) + " hours. Please review with the",
      "employee and the manager.",
      "",
      "Allo Innoware Task Tracker",
    ].join("\n");

    sendMail_(LEADERSHIP_EMAIL, "Leadership Escalation: " + row[1], body);
    alerts.getRange(i + 2, 7).setValue("Escalated");
    alerts.getRange(i + 2, 9).setValue(now);
    sent++;
  });

  SpreadsheetApp.flush();
  return sent;
}

// Sends an email, or logs it when SEND_EMAILS is false.
function sendMail_(to, subject, body) {
  if (!SEND_EMAILS) {
    Logger.log("[dry run] " + subject);
    return false;
  }
  try {
    MailApp.sendEmail(to, subject, body);
    return true;
  } catch (err) {
    Logger.log("Email failed: " + err);
    return false;
  }
}

// Handles edits: completions from Employee View, acks from Manager.
function handleEdit(e) {
  if (!e || !e.range) return;

  const sheet = e.range.getSheet();
  const name = sheet.getName();
  const row = e.range.getRow();
  const col = e.range.getColumn();

  if (name === MANAGER) {
    if (row < 2 || col !== 6) return;
    if (String(e.value || "").trim() !== "Acknowledged") return;

    const taskId = String(sheet.getRange(row, 1).getValue() || "").trim();
    if (!taskId) return;

    if (acknowledgeTask_(taskId)) {
      SpreadsheetApp.getActive().toast("Acknowledgement recorded for " + taskId + ".", "Manager", 5);
    } else {
      sheet.getRange(row, 6).setValue("");
      SpreadsheetApp.getActive().toast(
        "No alert has been sent for this cycle yet, so there is nothing to acknowledge.", "Manager", 6);
    }
    return;
  }

  if (name === VIEW) {
    if (row < 2) return;
    if (col !== 7 && col !== 8) return;
    if (e.range.getNumRows() > 1 || e.range.getNumColumns() > 1) return;

    const taskId = String(sheet.getRange(row, 2).getValue() || "").trim();
    if (!taskId) return;

    const proof = String(sheet.getRange(row, 7).getValue() || "").trim();
    const ticked = col === 8 && e.value === "TRUE";

    // One action is enough: a ticket ID commits on Enter, or the checkbox commits.
    if (!ticked && !proof) return;

    const employee = String(sheet.getRange(row, 1).getValue() || "").trim();
    try {
      const message = recordCompletion(taskId, employee, proof);
      SpreadsheetApp.getActive().toast(message, "Completion", 5);
    } catch (err) {
      SpreadsheetApp.getActive().toast(err.message, "Completion", 8);
    } finally {
      sheet.getRange(row, 7).setValue("");
      sheet.getRange(row, 8).setValue(false);
    }
  }
}

// Logs the spreadsheet name, to prove access works.
function testConnection() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  Logger.log("Spreadsheet name: " + ss.getName());
}

// Sends one test email to MANAGER_EMAIL.
function testManagerEmail() {
  MailApp.sendEmail(MANAGER_EMAIL, "Allo Innoware - Test Alert",
    "This is a test email from the Task Tracker prototype. Email automation is working.");
  Logger.log("Test email sent to: " + MANAGER_EMAIL);
}

// Checks that a wrong employee name is rejected.
function testCompletionValidation() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const log = ss.getSheetByName(LOG);
  const before = log.getLastRow();

  try {
    recordCompletion("T005", "Wrong Employee", "TEST-123");
    throw new Error("Test failed: incorrect employee was accepted.");
  } catch (error) {
    if (error.message.indexOf("Employee name does not match") === -1) throw error;
  }

  if (log.getLastRow() !== before) throw new Error("Test failed: a completion was recorded.");
  Logger.log("PASS: incorrect employee rejected, nothing written.");
}

// Checks that a missing ticket ID is rejected.
function testTicketValidation() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const log = ss.getSheetByName(LOG);
  const before = log.getLastRow();

  try {
    recordCompletion("T005", "Employee E", "");
    throw new Error("Test failed: missing ticket ID was accepted.");
  } catch (error) {
    if (error.message.indexOf("A ticket ID is required") === -1) throw error;
  }

  if (log.getLastRow() !== before) throw new Error("Test failed: a completion was recorded.");
  Logger.log("PASS: missing ticket ID rejected, nothing written.");
}

// Runs the eight scenarios and writes PASS/FAIL to Test Log.
function runAllTests() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(TESTLOG);
  if (!sh) throw new Error("Run setupAutomation() first to create the sheets.");
  if (sh.getLastRow() > 1) sh.getRange(2, 1, sh.getLastRow() - 1, 6).clearContent();
  cleanupTestRows_();

  const results = [];
  const add = function (test, scenario, expected, actual) {
    const pass = String(expected) === String(actual);
    results.push([test, scenario, String(expected), String(actual), pass ? "PASS" : "FAIL", new Date()]);
  };

  const now = new Date();
  const today = startOfDay_(now);
  const day = function (n) { return new Date(today.getTime() + n * 86400000); };

  seedTestTask_("T-TEST1", "Recurring 7 day", "Test User", 7, true, day(-1));
  recordCompletion("T-TEST1", "Test User", "TK-1001");
  const t1 = readTask_("T-TEST1");
  add("Test 1", "Complete a 7 day task and roll to the next cycle",
    ymd_(new Date(now.getTime() + 7 * 86400000)), t1 && t1.due ? ymd_(t1.due) : "none");

  const logs1 = readLogFor_("T-TEST1");
  add("Test 2", "Single completion entry, task row updated once",
    "1|yes", logs1.length + "|" + (t1 && t1.lastCompleted ? "yes" : "no"));

  seedTestTask_("T-TEST2", "Daily WhatsApp update", "Test User", 1, false, day(0));
  const r3 = recordCompletion("T-TEST2", "Test User", "WhatsApp update 09:15");
  const logs2 = readLogFor_("T-TEST2");
  add("Test 3", "Daily task completes without a ticket ID",
    "ok|WhatsApp update 09:15", (r3.indexOf("Completion recorded") === 0 ? "ok" : "fail") +
    "|" + (logs2[0] ? logs2[0][4] : "none"));

  seedTestTask_("T-TEST3", "Overdue ticket task", "Test User", 7, true, day(-1));
  sendManagerOverdueAlerts();
  const a4 = readAlert_("T-TEST3", day(-1));
  add("Test 4", "Manager alert created when the task is not completed by due time",
    "Sent", a4 ? String(a4.values[5]) : "none");

  acknowledgeTask_("T-TEST3");
  const a5 = readAlert_("T-TEST3", day(-1));
  add("Test 5", "Manager acknowledgement recorded with a timestamp",
    "Acknowledged|yes", (a5 ? String(a5.values[5]) : "none") + "|" +
    (a5 && a5.values[7] ? "yes" : "no"));

  const alertRow = a5 ? a5.row : 0;
  if (alertRow) ss.getSheetByName(ALERTS).getRange(alertRow, 5)
    .setValue(new Date(now.getTime() - (ESCALATION_HOURS + 1) * 3600000));
  sendLeadershipEscalations();
  const a6 = readAlert_("T-TEST3", day(-1));
  add("Test 6", "Unresolved task escalates to leadership after the SLA window",
    "Escalated", a6 ? String(a6.values[6]) : "none");

  seedTestTask_("T-TEST4", "Fresh overdue task", "Test User", 7, true, day(-1));
  sendManagerOverdueAlerts();
  sendLeadershipEscalations();
  const a7 = readAlert_("T-TEST4", day(-1));
  add("Test 7", "No escalation before the SLA window elapses",
    "", a7 ? String(a7.values[6]) : "none");

  const beforeAlerts = alertCount_();
  sendManagerOverdueAlerts();
  sendManagerOverdueAlerts();

  const t1row = readTask_("T-TEST1").row;
  ss.getSheetByName(TASKS).getRange(t1row, 8).setValue(day(-1));
  SpreadsheetApp.flush();
  const beforeLogs = readLogFor_("T-TEST1").length;
  const dupMessage = recordCompletion("T-TEST1", "Test User", "TK-1001");
  add("Test 8", "Repeated runs create no duplicate alerts or completions",
    beforeAlerts + "|" + beforeLogs + "|duplicate",
    alertCount_() + "|" + readLogFor_("T-TEST1").length + "|" +
    (dupMessage.indexOf("Already completed") === 0 ? "duplicate" : "new"));

  sh.getRange(2, 1, results.length, 6).setValues(results);
  cleanupTestRows_();

  const failed = results.filter(function (r) { return r[4] === "FAIL"; }).length;
  SpreadsheetApp.getActive().toast(
    (results.length - failed) + " of " + results.length + " tests passed. See the Test Log tab.",
    "Allo Ops", 8);
  return results;
}

// Adds a throwaway T-TEST task for the test run.
function seedTestTask_(id, name, employee, freq, requiresTicket, firstDue) {
  const sh = SpreadsheetApp.getActive().getSheetByName(TASKS);
  const row = sh.getLastRow() + 1;
  sh.getRange(row, 1, 1, 9).setValues([[
    id, name, employee, freq,
    requiresTicket ? "Portal Task" : "WhatsApp Update",
    requiresTicket ? "Yes" : "No",
    "", firstDue, "",
  ]]);
  SpreadsheetApp.flush();
}

// Returns the current Tasks values for a task ID.
function readTask_(id) {
  const sh = SpreadsheetApp.getActive().getSheetByName(TASKS);
  const row = findTaskRow_(sh, id);
  if (row === -1) return null;
  const r = sh.getRange(row, 1, 1, 9).getValues()[0];
  return { row: row, id: r[0], name: r[1], employee: r[2], freq: r[3], lastCompleted: r[6], due: r[7] };
}

// Returns every Completion Log row for a task ID.
function readLogFor_(id) {
  const sh = SpreadsheetApp.getActive().getSheetByName(LOG);
  if (sh.getLastRow() < 2) return [];
  return sh.getRange(2, 1, sh.getLastRow() - 1, 8).getValues()
    .filter(function (r) { return String(r[0]).trim() === id; });
}

// Reads one alert row as an object with row and values.
function readAlert_(id, due) {
  const sh = SpreadsheetApp.getActive().getSheetByName(ALERTS);
  const hit = findAlertRow_(sh, id, due);
  if (!hit) return null;
  const row = hit.row;
  const values = sh.getRange(row, 1, 1, 9).getValues()[0];
  return { row: row, values: values };
}

// Returns the number of alert rows, excluding the header.
function alertCount_() {
  const sh = SpreadsheetApp.getActive().getSheetByName(ALERTS);
  return Math.max(sh.getLastRow() - 1, 0);
}

// Deletes leftover T-TEST rows from every sheet.
function cleanupTestRows_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const targets = [
    [TASKS, 1], [LOG, 1], [HISTORY, 1], [ALERTS, 1],
  ];
  targets.forEach(function (pair) {
    const sh = ss.getSheetByName(pair[0]);
    if (!sh || sh.getLastRow() < 2) return;
    const col = pair[1];
    const values = sh.getRange(2, col, sh.getLastRow() - 1, 1).getValues();
    for (let i = values.length - 1; i >= 0; i--) {
      if (String(values[i][0]).indexOf("T-TEST") === 0) sh.deleteRow(i + 2);
    }
  });
}

// Clears all data and seeds the demo tasks.
function resetDemoData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss.getSheetByName(TASKS)) {
    throw new Error("Run setupAutomation() first to create the sheets.");
  }

  [TASKS, LOG, HISTORY, ALERTS, TESTLOG].forEach(function (name) {
    const sh = ss.getSheetByName(name);
    if (sh && sh.getLastRow() > 1) {
      sh.getRange(2, 1, sh.getLastRow() - 1, sh.getMaxColumns()).clearContent();
    }
  });

  const managerSheet = ss.getSheetByName(MANAGER);
  if (managerSheet) {
    managerSheet.getRange(2, 6, Math.max(managerSheet.getMaxRows() - 1, 1), 1).clearContent();
  }
  const viewSheet = ss.getSheetByName(VIEW);
  if (viewSheet) viewSheet.getRange(2, 7, 99, 2).clearContent();

  const today = startOfDay_(new Date());
  const day = function (n) { return new Date(today.getTime() + n * 86400000); };

  const rows = [
    ["T001", "Daily Operations Update", "Employee A", 1, "WhatsApp Update", "No", day(0)],
    ["T002", "Portal Task A", "Employee B", 7, "Portal Task", "Yes", day(7)],
    ["T003", "Portal Task B", "Employee C", 15, "Portal Task", "Yes", day(15)],
    ["T004", "Portal Task C", "Employee D", 30, "Portal Task", "Yes", day(30)],
    ["T005", "Overdue Demo", "Employee E", 1, "Portal Task", "Yes", day(-2)],
  ];

  const sh = ss.getSheetByName(TASKS);
  sh.getRange(2, 1, rows.length, 6).setValues(rows.map(function (r) {
    return [r[0], r[1], r[2], r[3], r[4], r[5]];
  }));
  sh.getRange(2, 7, rows.length, 1).setValue("");
  sh.getRange(2, 8, rows.length, 1).setValues(rows.map(function (r) { return [r[6]]; }));
  SpreadsheetApp.flush();

  repairSheets();

  SpreadsheetApp.getActive().toast(
    "Demo data seeded. Run 'Run pending checks now' to create the overdue alert.", "Allo Ops", 6);
}

// Returns midnight at the start of the given day.
function startOfDay_(d) {
  const x = new Date(d.getTime());
  x.setHours(0, 0, 0, 0);
  return x;
}

// Formats a date as yyyy-MM-dd in the spreadsheet timezone.
function ymd_(d) {
  return Utilities.formatDate(d, sheetTz_(), "yyyy-MM-dd");
}

// Spreadsheet timezone, so dates agree with the sheet's TODAY().
function sheetTz_() {
  return SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetTimeZone();
}

// Builds the dedupe key: task ID plus cycle due date.
function cycleKey_(taskId, due) {
  return String(taskId).trim() + "|" + ymd_(due);
}
