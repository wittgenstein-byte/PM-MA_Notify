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
// hour: ชั่วโมงที่ต้องการให้ส่งแจ้งเตือน (0-23) ค่าเริ่มต้น = 8 (08:00)
// ตัวอย่าง: setupDailyTrigger()     → รันทุกวัน 08:00
//          setupDailyTrigger(9)    → รันทุกวัน 09:00
//          setupDailyTrigger(17)   → รันทุกวัน 17:00
function setupDailyTrigger(hour) {
  if (hour === undefined || hour === null) hour = 8;
  hour = Math.max(0, Math.min(23, Math.floor(hour)));

  // ลบ trigger เก่าก่อน
  ScriptApp.getProjectTriggers().forEach(t => ScriptApp.deleteTrigger(t));

  // สร้าง trigger ใหม่
  ScriptApp.newTrigger("checkAndSendNotifications")
    .timeBased()
    .everyDays(1)
    .atHour(hour)
    .create();

  const timeStr = String(hour).padStart(2, '0') + ':00';
  Logger.log(`✅ Daily trigger ตั้งค่าเรียบร้อย — รันทุกวัน ${timeStr}`);
}
