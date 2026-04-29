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
// notifyTime: เวลาที่ต้องการแจ้งเตือน เช่น "08:00", "09:30" (ค่าเริ่มต้น "08:00")
function createNotificationRules(contractId, endDate, alertDaysArray, notifyTime) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const rulesSheet = ss.getSheetByName("notification_rules");

  // ตรวจสอบ notifyTime — ถ้าไม่มีให้ใช้ค่าเริ่มต้น 08:00
  const time = notifyTime && /^\d{2}:\d{2}$/.test(notifyTime) ? notifyTime : '08:00';

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
      time,                // notify_time (เวลาที่ต้องการแจ้งเตือน เช่น "08:00")
    ]);
  });
}

// ตั้ง Time-based Trigger — รันทุกชั่วโมง
// เพื่อให้สามารถส่งแจ้งเตือนตามเวลาที่แต่ละสัญญากำหนดได้
// ตัวอย่าง: setupHourlyTrigger()  → รันทุกชั่วโมง
function setupHourlyTrigger() {
  // ลบ trigger เก่าก่อน
  ScriptApp.getProjectTriggers().forEach(t => ScriptApp.deleteTrigger(t));

  // สร้าง trigger ใหม่ — รันทุก 1 ชั่วโมง
  ScriptApp.newTrigger("checkAndSendNotifications")
    .timeBased()
    .everyHours(1)
    .create();

  Logger.log('✅ Hourly trigger ตั้งค่าเรียบร้อย — รันทุก 1 ชั่วโมง เพื่อเช็คเวลาแจ้งเตือนแต่ละสัญญา');
}

// Backward-compatible: ยังคงใช้ได้ แต่จะเปลี่ยนเป็น hourly trigger อัตโนมัติ
function setupDailyTrigger(hour) {
  Logger.log('⚠️ setupDailyTrigger() เปลี่ยนเป็น setupHourlyTrigger() แล้ว — รันทุกชั่วโมงเพื่อรองรับเวลาแจ้งเตือนแต่ละสัญญา');
  setupHourlyTrigger();
}
