// ============================================================
//  UTILITIES
// ============================================================

// สร้าง ID อัตโนมัติ
function generateId(prefix) {
  const timestamp = new Date().getTime();
  return `${prefix}-${timestamp}`;
}

// บันทึก Log
function logNotification(ruleId, contractId, channel, status, errorMsg = "") {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const logSheet = ss.getSheetByName("notification_logs");

  logSheet.appendRow([
    generateId("LOG"),
    ruleId,
    contractId,
    channel,
    status,
    errorMsg,
    new Date(),
  ]);
}

// สร้าง notification_rules อัตโนมัติเมื่อ Sale บันทึกสัญญาใหม่
function createNotificationRules(contractId, endDate, alertDaysArray) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const rulesSheet = ss.getSheetByName("notification_rules");

  alertDaysArray.forEach(days => {
    const scheduledDate = new Date(endDate);
    scheduledDate.setDate(scheduledDate.getDate() - days);

    rulesSheet.appendRow([
      generateId("NR"),   // rule_id
      contractId,          // contract_id
      days,                // alert_days_before
      scheduledDate,       // scheduled_date
      true,                // notify_email
      true,                // notify_teams
      false,               // is_sent
      "",                  // sent_at
    ]);
  });
}

// ตั้ง Time-based Trigger (รันครั้งเดียวตอน setup)
function setupDailyTrigger() {
  // ลบ trigger เก่าก่อน
  ScriptApp.getProjectTriggers().forEach(t => ScriptApp.deleteTrigger(t));

  // สร้าง trigger ใหม่ รันทุกวัน เวลา 08:00
  ScriptApp.newTrigger("checkAndSendNotifications")
    .timeBased()
    .everyDays(1)
    .atHour(8)
    .create();

  Logger.log("✅ Daily trigger ตั้งค่าเรียบร้อย — รันทุกวัน 08:00");
}
