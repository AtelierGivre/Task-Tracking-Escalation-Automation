const MANAGER_EMAIL = "manager@example.com"; // Replace with actual email or use PropertiesService
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
const TESTLOG = "Test Log";

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
    sh.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight("bold");
    sh.setFrozenRows(1);
    return sh;
  };

  const tasks = ensure(TASKS, [
    "Task ID", "Task Name", "Assigned Employee", "Frequency (Days)",
    "Task Type", "Requires Ticket ID", "Last Completed", "Next Due", "Status"
  ]);
  const log = ensure(LOG, [
    "Task ID", "Task Name", "Completed By", "Completion Date & Time",
    "Proof / Ticket ID", "Manager Notified", "Cycle Due Date", "Correction Count"
  ]);
  ensure(HISTORY, [
    "Task ID", "Task Name", "Employee Name", "Original Completion Date & Time",
    "Original Proof / Ticket ID", "Correction Date & Time", "Corrected Proof / Ticket ID", "Correction Count"
  ]);
  const alerts = ensure(ALERTS, [
    "Task ID", "Task Name", "Assigned Employee", "Cycle Due Date",
    "Alert Date & Time", "Alert Status", "Escalation Status", "Acknowledged At", "Escalated At"
  ]);
  const manager = ensure(MANAGER, [
    "Task ID", "Task Name", "Assigned Employee", "Next Due", "Status", "Manager Acknowledgement"
  ]);
  const view = ensure(VIEW, [
    "Employee Name", "Task ID", "Task Name", "Due Date", "Status", "Last Proof",
    "Enter Ticket ID / Update", "Mark Complete"
  ]);
  ensure(TESTLOG, ["Test", "Scenario", "Expected", "Actual", "Result", "Run At"]);

  const alertSpill = alerts.getRange("A2").getFormula();
  if (alertSpill) {
    alerts.getRange(2, 1, alerts.getMaxRows() - 1, alerts.getMaxColumns()).clearContent();
  }
  if (alerts.getLastRow() < 2) {
    manager.getRange(2, 6, Math.max(manager.getMaxRows() - 1, 1), 1).clearContent();
  }

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
    statusFormulas.push([
      '=IF($A' + r + '="","",IF($H' + r + '<TODAY(),"Overdue",' +
      'IF($H' + r + '=TODAY(),"Due Today","Pending")))'
    ]);
  }

  if (dueOut.length) {
    tasks.getRange(2, 8, dueOut.length, 1).setValues(dueOut);
    tasks.getRange(2, 9, statusFormulas.length, 1).setFormulas(statusFormulas);
    tasks.getRange(2, 8, dueOut.length, 1).setNumberFormat("yyyy-mm-dd");
  }

  tasks.getRange(2, 6, Math.max(lastTask - 1, 1), 1).setDataValidation(
    SpreadsheetApp.newDataValidation().requireValueInList(["Yes", "No"], true).build()
  );

  manager.getRange("A2").setFormula(
    '=IFERROR(FILTER({Tasks!A2:A1000,Tasks!B2:B1000,Tasks!C2:C1000,Tasks!H2:H1000,Tasks!I2:I1000},' +
    'Tasks!A2:A1000<>""),"")'
  );
  manager.getRange(2, 6, Math.max(manager.getMaxRows() - 1, 1), 1).setDataValidation(
    SpreadsheetApp.newDataValidation().requireValueInList(["Acknowledged"], true).build()
  );

  view.getRange("A2").setFormula(
    '=IFERROR(FILTER({Tasks!C2:C1000,Tasks!A2:A1000,Tasks!B2:B1000,Tasks!H2:H1000,Tasks!I2:I1000},' +
    'Tasks!A2:A1000<>""),"")'
  );
  view.getRange("F2").setFormula(
    '=IF(B2:B1000="","",IFNA(VLOOKUP(B2:B1000,SORT(\'Completion Log\'!A2:E1000,4,FALSE),5,FALSE),""))'
  );
  view.getRange(2, 8, 100, 1).insertCheckboxes();
  view.setFrozenRows(1);

  SpreadsheetApp.flush();
}

// Installs the hourly run and the on-edit trigger.
function installTriggers() {
  ScriptApp.getProjectTriggers().forEach(function (t) {
    const handler = t.getHandlerFunction();
    if (handler === "runAutomation" || handler === "onEdit" || handler === "handleEdit") {
      ScriptApp.deleteTrigger(t);
    }
  });
  ScriptApp.newTrigger("runAutomation").timeBased().everyHours(1).create();
  ScriptApp.newTrigger("handleEdit").forSpreadsheet(SpreadsheetApp.getActive()).onEdit().create();
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

function findTaskRow_(sheet, taskId) {
  const last = sheet.getLastRow();
  if (last < 2) return -1;
  const ids = sheet.getRange(2, 1, last - 1, 1).getValues();
  for (let i = 0; i < ids.length; i++) {
    if (String(ids[i][0]).trim() === String(taskId).trim()) return i + 2;
  }
  return -1;
}

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

function clearManagerTick_(ss, taskId) {
  const manager = ss.getSheetByName(MANAGER);
  const tasks = ss.getSheetByName(TASKS);
  if (!manager || !tasks) return;
  const taskRow = findTaskRow_(tasks, taskId);
  if (taskRow === -1 || taskRow > manager.getMaxRows()) return;
  manager.getRange(taskRow, 6).clearContent();
}

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

function sendManagerOverdueAlerts() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const tasks = ss.getSheetByName(TASKS);
  const alerts = ss.getSheetByName(ALERTS);
  if (!tasks || !alerts) throw new Error("Tasks or Manager Alerts sheet is missing.");
  const lastRow = tasks.getLastRow();
  if (lastRow < 2) return 0;
  const data = tasks.getRange(2, 1, lastRow - 1, 9).getValues();
  let sent = 0;

  data.forEach(function (task) {
    const id = String(task[0] || "").trim();
    if (!id) return;
    const due = task[7] instanceof Date ? task[7] : null;
    if (!due) return;
    const dueAt = new Date(due.getTime());
    dueAt.setHours(DUE_HOUR, 0, 0, 0);
    if (new Date() < dueAt) return;
    if (findAlertRow_(alerts, id, due)) return;

    const now = new Date();
    alerts.appendRow([id, task[1], task[2], due, now, "Sent", "", "", ""]);
    const body = [
      "Hello Manager,", "", "The following recurring task is overdue:", "",
      "Task ID:          " + id, "Task:             " + task[1],
      "Assigned to:      " + task[2], "Cycle due:        " + ymd_(due), "",
      "Ask the employee to complete it in the Employee View tab, then tick",
      "Manager Acknowledgement in the Manager tab.", "",
      "If it is still open after " + ESCALATION_HOURS + " hours it is escalated automatically.", "",
      "Allo Innoware Task Tracker",
    ].join("\n");

    sendMail_(MANAGER_EMAIL, "Overdue Task Alert: " + task[1], body);
    sent++;
  });

  SpreadsheetApp.flush();
  return sent;
}

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
      "Hello,", "", "This recurring task is still unresolved:", "",
      "Task ID:      " + id, "Task:         " + row[1],
      "Assigned to:  " + row[2], "Cycle due:    " + ymd_(row[3]),
      "Manager alert sent: " + ymd_(alertAt),
      "Manager acknowledged: " + (status === "Acknowledged" ? "Yes" : "No"), "",
      "It has been open for " + Math.round(hours) + " hours. Please review with the",
      "employee and the manager.", "", "Allo Innoware Task Tracker",
    ].join("\n");

    sendMail_(LEADERSHIP_EMAIL, "Leadership Escalation: " + row[1], body);
    alerts.getRange(i + 2, 7).setValue("Escalated");
    alerts.getRange(i + 2, 9).setValue(now);
    sent++;
  });

  SpreadsheetApp.flush();
  return sent;
}

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
        "No alert has been sent for this cycle yet, so there is nothing to acknowledge.",
        "Manager", 6
      );
    }
    return;
  }

  if (name === VIEW) {
    if (row < 2) return;
    if (col !== 7 && col !== 8) return;
    const taskId = String(sheet.getRange(row, 2).getValue() || "").trim();
    if (!taskId) return;
    const proof = String(sheet.getRange(row, 7).getValue() || "").trim();
    if (col === 7 && !proof) return;
    if (col === 8 && e.value !== "TRUE") return;
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

function testConnection() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  Logger.log("Spreadsheet name: " + ss.getName());
}

function testManagerEmail() {
  sendMail_(MANAGER_EMAIL, "Allo Innoware - Test Alert", "This is a test email from the Task Tracker prototype. Email automation is working.");
  Logger.log("Test email sent to: " + MANAGER_EMAIL);
}

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
  add("Test 1", "Complete a 7 day task and roll to the next cycle", ymd_(new Date(now.getTime() + 7 * 86400000)), t1 && t1.due ? ymd_(t1.due) : "none");

  const logs1 = readLogFor_("T-TEST1");
  add("Test 2", "Single completion entry, task row updated once", "1|yes", logs1.length + "|" + (t1 && t1.lastCompleted ? "yes" : "no"));

  seedTestTask_("T-TEST2", "Daily WhatsApp update", "Test User", 1, false, day(0));
  const r3 = recordCompletion("T-TEST2", "Test User", "WhatsApp update 09:15");
  const logs2 = readLogFor_("T-TEST2");
  add("Test 3", "Daily task completes without a ticket ID", "ok|WhatsApp update 09:15", (r3.indexOf("Completion recorded") === 0 ? "ok" : "fail") + "|" + (logs2[0] ? logs2[0][4] : "none"));

  seedTestTask_("T-TEST3", "Overdue ticket task", "Test User", 7, true, day(-1));
  sendManagerOverdueAlerts();
  const a4 = readAlert_("T-TEST3", day(-1));
  add("Test 4", "Manager alert created when the task is not completed by due time", "Sent", a4 ? String(a4.values[5]) : "none");

  acknowledgeTask_("T-TEST3");
  const a5 = readAlert_("T-TEST3", day(-1));
  add("Test 5", "Manager acknowledgement recorded with a timestamp", "Acknowledged|yes", (a5 ? String(a5.values[5]) : "none") + "|" + (a5 && a5.values[7] ? "yes" : "no"));

  const alertRow = a5 ? a5.row : 0;
  if (alertRow) ss.getSheetByName(ALERTS).getRange(alertRow, 5).setValue(new Date(now.getTime() - (ESCALATION_HOURS + 1) * 3600000));
  sendLeadershipEscalations();
  const a6 = readAlert_("T-TEST3", day(-1));
  add("Test 6", "Unresolved task escalates to leadership after the SLA window", "Escalated", a6 ? String(a6.values[6]) : "none");

  seedTestTask_("T-TEST4", "Fresh overdue task", "Test User", 7, true, day(-1));
  sendManagerOverdueAlerts();
  sendLeadershipEscalations();
  const a7 = readAlert_("T-TEST4", day(-1));
  add("Test 7", "No escalation before the SLA window elapses", "", a7 ? String(a7.values[6]) : "none");

  const beforeAlerts = alertCount_();
  sendManagerOverdueAlerts();
  sendManagerOverdueAlerts();
  const t1row = readTask_("T-TEST1").row;
  ss.getSheetByName(TASKS).getRange(t1row, 8).setValue(day(-1));
  SpreadsheetApp.flush();
  const beforeLogs = readLogFor_("T-TEST1").length;
  const dupMessage = recordCompletion("T-TEST1", "Test User", "TK-1001");
  add("Test 8", "Repeated runs create no duplicate alerts or completions", beforeAlerts + "|" + beforeLogs + "|duplicate", alertCount_() + "|" + readLogFor_("T-TEST1").length + "|" + (dupMessage.indexOf("Already completed") === 0 ? "duplicate" : "new"));

  sh.getRange(2, 1, results.length, 6).setValues(results);
  cleanupTestRows_();

  const failed = results.filter(function (r) { return r[4] === "FAIL"; }).length;
  SpreadsheetApp.getActive().toast((results.length - failed) + " of " + results.length + " tests passed. See the Test Log tab.", "Allo Ops", 8);
  return results;
}

function seedTestTask_(id, name, employee, freq, requiresTicket, firstDue) {
  const sh = SpreadsheetApp.getActive().getSheetByName(TASKS);
  const row = sh.getLastRow() + 1;
  sh.getRange(row, 1, 1, 9).setValues([[
    id, name, employee, freq, requiresTicket ? "Portal Task" : "WhatsApp Update",
    requiresTicket ? "Yes" : "No", "", firstDue, ""
  ]]);
  SpreadsheetApp.flush();
}

function readTask_(id) {
  const sh = SpreadsheetApp.getActive().getSheetByName(TASKS);
  const row = findTaskRow_(sh, id);
  if (row === -1) return null;
  const r = sh.getRange(row, 1, 1, 9).getValues()[0];
  return { row: row, id: r[0], name: r[1], employee: r[2], freq: r[3], lastCompleted: r[6], due: r[7] };
}

function readLogFor_(id) {
  const sh = SpreadsheetApp.getActive().getSheetByName(LOG);
  if (sh.getLastRow() < 2) return [];
  return sh.getRange(2, 1, sh.getLastRow() - 1, 8).getValues()
    .filter(function (r) { return String(r[0]).trim() === id; });
}

function readAlert_(id, due) {
  const sh = SpreadsheetApp.getActive().getSheetByName(ALERTS);
  const hit = findAlertRow_(sh, id, due);
  if (!hit) return null;
  const row = hit.row;
  const values = sh.getRange(row, 1, 1, 9).getValues()[0];
  return { row: row, values: values };
}

function alertCount_() {
  const sh = SpreadsheetApp.getActive().getSheetByName(ALERTS);
  return Math.max(sh.getLastRow() - 1, 0);
}

function cleanupTestRows_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const targets = [[TASKS, 1], [LOG, 1], [HISTORY, 1], [ALERTS, 1]];
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
  sh.getRange(2, 1, rows.length, 6).setValues(rows.map(function (r) { return [r[0], r[1], r[2], r[3], r[4], r[5]]; }));
  sh.getRange(2, 7, rows.length, 1).setValue("");
  sh.getRange(2, 8, rows.length, 1).setValues(rows.map(function (r) { return [r[6]]; }));

  SpreadsheetApp.flush();
  repairSheets();
  SpreadsheetApp.getActive().toast("Demo data seeded. Run 'Run pending checks now' to create the overdue alert.", "Allo Ops", 6);
}

function startOfDay_(d) {
  const x = new Date(d.getTime());
  x.setHours(0, 0, 0, 0);
  return x;
}

function ymd_(d) {
  return Utilities.formatDate(d, Session.getScriptTimeZone(), "yyyy-MM-dd");
}

function cycleKey_(taskId, due) {
  return String(taskId).trim() + "|" + ymd_(due);
}
