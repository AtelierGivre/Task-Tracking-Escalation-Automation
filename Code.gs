/**
 * Allo Innoware — Task Tracking & Escalation Automation
 * Google Apps Script prototype.
 *
 * Before use:
 * 1. Attach this script to the intended Google Spreadsheet.
 * 2. Confirm the sheet names and column positions documented in README.md.
 * 3. Set MANAGER_EMAIL in Script Properties.
 * 4. Review and authorize email permissions before enabling triggers.
 *
 * This is a prototype, not a production-hardened access-control system.
 */

/**
 * Configure the recipient in Apps Script:
 * Project Settings > Script properties
 * Add MANAGER_EMAIL with the intended manager/leadership email address.
 */
function getManagerEmail_() {
  const email = PropertiesService.getScriptProperties().getProperty("MANAGER_EMAIL");
  if (!email) {
    throw new Error('Missing MANAGER_EMAIL. Add it under Project Settings > Script properties.');
  }
  return email;
}
function testConnection() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  Logger.log("Spreadsheet name: " + ss.getName());
}


function recordCompletion(taskId, employeeName, proof) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);

  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const tasks = ss.getSheetByName("Tasks");
    const log = ss.getSheetByName("Completion Log");
    const history = ss.getSheetByName("Correction History");

    if (!tasks || !log || !history) {
      throw new Error("A required sheet is missing.");
    }


    const lastRow = tasks.getLastRow();
    const data = lastRow < 2
      ? []
      : tasks.getRange(2, 1, lastRow - 1, 9).getValues();

    const taskIndex = data.findIndex(row => row[0] === taskId);

    if (taskIndex === -1) {
      throw new Error("Task ID not found: " + taskId);
    }

    const task = data[taskIndex];

    if (task[2] !== employeeName) {
      throw new Error("Employee name does not match the assigned employee.");
    }

    if (task[5] === "Yes" && !String(proof || "").trim()) {
      throw new Error("A ticket ID is required for this task.");
    }

    const submittedProof = String(proof || "WhatsApp update").trim();
    const now = new Date();
    const taskRow = taskIndex + 2;

    const currentDue = task[7] instanceof Date
      ? task[7]
      : new Date(task[7]);

    if (isNaN(currentDue.getTime())) {
      throw new Error("The task due date is invalid.");
    }

    const lastCompleted = task[6];

    const logLastRow = log.getLastRow();
    const logData = logLastRow < 2
      ? []
      : log.getRange(2, 1, logLastRow - 1, 8).getValues();

    let existingIndex = -1;

    if (lastCompleted instanceof Date) {
      existingIndex = logData.findIndex(row =>
        row[0] === taskId &&
        row[2] === employeeName &&
        row[3] instanceof Date &&
        row[3].getTime() === lastCompleted.getTime()
      );
    }

    // CORRECTION: Update the latest completion before the next due date.
    if (existingIndex !== -1 && now < currentDue) {
      const logRow = existingIndex + 2;
      const original = logData[existingIndex];
      const correctionCount = Number(original[7]) || 0;

      if (correctionCount >= 2) {
        throw new Error("Maximum of two corrections already used.");
      }

      const historyLastRow = history.getLastRow();
      const historyData = historyLastRow < 2
        ? []
        : history.getRange(2, 1, historyLastRow - 1, 7).getValues();

      const previousCorrections = historyData.filter(row =>
        row[0] === taskId &&
        row[2] === employeeName &&
        row[3] instanceof Date &&
        row[3].getTime() === original[3].getTime()
      );

      const originalProof = previousCorrections.length > 0
        ? previousCorrections[0][4]
        : original[4];

      const newHistoryRow = history.getLastRow() + 1;

      history.getRange(newHistoryRow, 1, 1, 7).setValues([[
        taskId,
        task[1],
        employeeName,
        original[3],
        originalProof,
        now,
        submittedProof
      ]]);

      const savedHistoryProof =
        history.getRange(newHistoryRow, 7).getValue();

      if (String(savedHistoryProof) !== submittedProof) {
        throw new Error("Correction History proof verification failed.");
      }

      log.getRange(logRow, 5).setValue(submittedProof);
      log.getRange(logRow, 8).setValue(correctionCount + 1);

      SpreadsheetApp.flush();

      const savedProof = log.getRange(logRow, 5).getValue();
      const savedCount = log.getRange(logRow, 8).getValue();

      Logger.log("Target log row: " + logRow);
      Logger.log("Saved correction count: " + savedCount);

      if (String(savedProof) !== submittedProof) {
        throw new Error("Completion Log proof verification failed.");
      }

      return "Correction saved successfully.";
    }

    // NEW COMPLETION: Start a new recurring cycle.
    log.appendRow([
      taskId,
      task[1],
      employeeName,
      now,
      submittedProof,
      "No",
      currentDue,
      0
    ]);

    const nextDue = new Date(
      now.getTime() + Number(task[3]) * 24 * 60 * 60 * 1000
    );

    tasks.getRange(taskRow, 7).setValue(now);
    tasks.getRange(taskRow, 8).setValue(nextDue);
    tasks.getRange(taskRow, 9).setValue("Pending");

    SpreadsheetApp.flush();

    return "New completion recorded successfully.";

  } finally {
    lock.releaseLock();
  }
}


// VALIDATION TESTS

function testCompletionValidation() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const log = ss.getSheetByName("Completion Log");
  const beforeCount = log.getLastRow();

  try {
    recordCompletion("T005", "Wrong Employee", "TEST-123");
    throw new Error("Test failed: incorrect employee was accepted.");
  } catch (error) {
    if (!error.message.includes("Employee name does not match")) {
      throw error;
    }
  }

  if (log.getLastRow() !== beforeCount) {
    throw new Error("Test failed: a completion was recorded.");
  }

  Logger.log("PASS: Incorrect employee rejected; no completion recorded.");
}


function testTicketValidation() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const log = ss.getSheetByName("Completion Log");
  const beforeCount = log.getLastRow();

  try {
    recordCompletion("T005", "Employee E", "");
    throw new Error("Test failed: missing ticket ID was accepted.");
  } catch (error) {
    if (!error.message.includes("A ticket ID is required")) {
      throw error;
    }
  }

  if (log.getLastRow() !== beforeCount) {
    throw new Error("Test failed: a completion was recorded.");
  }

  Logger.log("PASS: Missing ticket ID rejected; no completion recorded.");
}


// Diagnostic functions that modified sample records were intentionally excluded from this public version.\n\nfunction sendManagerOverdueAlerts() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const tasks = ss.getSheetByName("Tasks");
  const alerts = ss.getSheetByName("Manager Alerts");

  if (!tasks || !alerts) {
    throw new Error("Tasks or Manager Alerts sheet is missing.");
  }

  const today = new Date();
  today.setHours(0, 0, 0, 0);

  const lastRow = tasks.getLastRow();
  if (lastRow < 2) return;

  const taskData = tasks.getRange(2, 1, lastRow - 1, 9).getValues();

  const alertLastRow = alerts.getLastRow();
  const alertData = alertLastRow < 2
    ? []
    : alerts.getRange(2, 1, alertLastRow - 1, 7).getValues();

  const alertRows = new Map();

  alertData.forEach((row, index) => {
    if (row[0]) alertRows.set(row[0], index + 2);
  });

  taskData.forEach(task => {
    const taskId = task[0];
    const taskName = task[1];
    const employee = task[2];
    const dueDate = task[7];
    const status = task[8];

    if (!taskId || !(dueDate instanceof Date)) return;
    if (status !== "Overdue" || dueDate >= today) return;

    const alertRow = alertRows.get(taskId);
    if (!alertRow) return;

    const alertStatus = alerts.getRange(alertRow, 6).getValue();

    if (alertStatus === "Sent" || alertStatus === "Acknowledged") {
      return;
    }

    const subject = "Overdue Task Alert: " + taskName;

    const body =
      "Hello Manager,\n\n" +
      "The following task is overdue:\n\n" +
      "Task ID: " + taskId + "\n" +
      "Task: " + taskName + "\n" +
      "Assigned Employee: " + employee + "\n" +
      "Due Date: " + dueDate.toDateString() + "\n\n" +
      "Please review and acknowledge this task.\n\n" +
      "Allo Innoware Task Tracker";

    MailApp.sendEmail(getManagerEmail_(), subject, body);

    alerts.getRange(alertRow, 5).setValue(new Date());
    alerts.getRange(alertRow, 6).setValue("Sent");
  });

  SpreadsheetApp.flush();
  Logger.log("Overdue alert check completed.");
}
function testManagerEmail() {
  MailApp.sendEmail(
    MANAGER_EMAIL,
    "Allo Innoware - Test Alert",
    "This is a test email from the Task Tracker prototype. Email automation is working."
  );

  Logger.log("Test email sent.");
}

function onEdit(e) {
  if (!e || !e.range) return;

  const sheet = e.range.getSheet();

  if (
    sheet.getName() !== "Manager Dashboard" ||
    e.range.getColumn() !== 6 ||
    e.range.getRow() < 2 ||
    e.value !== "Acknowledged"
  ) return;

  const ss = e.source;
  const taskId = sheet.getRange(e.range.getRow(), 1).getValue();
  const alerts = ss.getSheetByName("Manager Alerts");

  if (!taskId || !alerts) return;

  const lastRow = alerts.getLastRow();
  if (lastRow < 2) return;

  const taskIds = alerts.getRange(2, 1, lastRow - 1, 1).getValues();

  for (let i = 0; i < taskIds.length; i++) {
    if (String(taskIds[i][0]).trim() === String(taskId).trim()) {
      alerts.getRange(i + 2, 6).setValue("Acknowledged");
      alerts.getRange(i + 2, 8).setValue(new Date());
      break;
    }
  }
}

function sendLeadershipEscalations() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const tasks = ss.getSheetByName("Tasks");
  const alerts = ss.getSheetByName("Manager Alerts");

  if (!tasks || !alerts) {
    throw new Error("Required sheet is missing.");
  }

  const taskData = tasks.getDataRange().getValues();
  const alertData = alerts.getDataRange().getValues();

  const now = new Date();
  const escalationEmail = getManagerEmail_();

  for (let i = 1; i < taskData.length; i++) {
    const task = taskData[i];
    const taskId = task[0];
    const status = task[8];

    if (!taskId || status !== "Overdue") continue;

    const alertIndex = alertData.findIndex(
      (row, index) => index > 0 && row[0] === taskId
    );

    if (alertIndex === -1) continue;

    const alert = alertData[alertIndex];
    const alertStatus = alert[5];
    const escalationStatus = alert[6];

    if (
      alertStatus !== "Sent" ||
      alertStatus === "Acknowledged" ||
      escalationStatus === "Escalated"
    ) continue;

    const alertTime = alert[4];

    if (!(alertTime instanceof Date)) continue;

    const hoursElapsed = (now - alertTime) / (1000 * 60 * 60);

    // Proposed prototype SLA: 4 elapsed hours.
    if (hoursElapsed < 4) continue;

    const subject = "Leadership Escalation: " + task[1];

    const body =
      "Hello,\n\n" +
      "The following overdue task remains unresolved:\n\n" +
      "Task ID: " + taskId + "\n" +
      "Task: " + task[1] + "\n" +
      "Assigned Employee: " + task[2] + "\n\n" +
      "The manager alert was sent more than 4 hours ago.\n" +
      "Please review this pending task.\n\n" +
      "Allo Innoware Task Tracker";

    MailApp.sendEmail(escalationEmail, subject, body);

    alerts.getRange(alertIndex + 1, 7).setValue("Escalated");
  }

  SpreadsheetApp.flush();
  Logger.log("Leadership escalation check completed.");
}
