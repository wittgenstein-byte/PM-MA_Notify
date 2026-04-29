// ============================================================
//  MAIN SCHEDULER — รันทุกชั่วโมง (hourly trigger)
//  เช็คเวลาแจ้งเตือนที่แต่ละ rule กำหนด (notify_time)
// ============================================================

function checkAndSendNotifications() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const rulesSheet = ss.getSheetByName("notification_rules");
  const contractsSheet = ss.getSheetByName("warranty_contracts");

  const now = new Date();
  const today = new Date(now);
  today.setHours(0, 0, 0, 0);

  const currentHour = now.getHours();

  const rulesData = rulesSheet.getDataRange().getValues();
  const contractsData = contractsSheet.getDataRange().getValues();

  // Map contracts by contract_id
  const contractMap = {};
  for (let i = 1; i < contractsData.length; i++) {
    const row = contractsData[i];
    contractMap[row[0]] = {
      contract_id:     row[0],
      po_number:       row[1],
      project_name:    row[2],
      customer_name:   row[3],
      service_type:    row[4],
      start_date:      row[5],
      end_date:        row[6],
      recipients_sale: row[7],
      recipients_eng:  row[8],
      teams_webhook:   row[9],
      note:            row[10],
      status:          row[11],
      line_group_id:   row[12],
      notify_line:     row[13],
    };
  }

  // Cache access token for email to reduce redundant calls
  let cachedEmailToken = null;

  // Loop notification rules
  // Schema: rule_id(0) | contract_id(1) | alert_days_before(2) | scheduled_date(3) |
  //         notify_email(4) | notify_teams(5) | is_sent(6) | sent_at(7) | notify_time(8)
  for (let i = 1; i < rulesData.length; i++) {
    const rule = rulesData[i];
    const ruleId           = rule[0];
    const contractId       = rule[1];
    const alertDays        = rule[2];
    const scheduledDate    = new Date(rule[3]);
    const notifyEmail      = rule[4];
    const notifyTeams      = rule[5];
    const isSent           = rule[6];
    const notifyTime       = rule[8] || '08:00'; // ค่าเริ่มต้น 08:00

    scheduledDate.setHours(0, 0, 0, 0);

    // ข้ามถ้าส่งแล้ว
    if (isSent) continue;

    // ข้ามถ้ายังไม่ถึงวัน
    if (scheduledDate > today) continue;

    // ────────────────────────────────────────────────────────
    // เช็คเวลาแจ้งเตือน: ส่งเฉพาะเมื่อชั่วโมงปัจจุบัน ตรงกับ
    // ชั่วโมงที่กำหนด (notify_time)
    // ────────────────────────────────────────────────────────
    const scheduledHour = parseInt(notifyTime.split(':')[0], 10) || 8;

    // ถ้าวันนี้ = scheduled_date → ส่งเฉพาะเมื่อถึงชั่วโมงที่กำหนด
    // ถ้าวันนี้เลย scheduled_date ไปแล้ว (สัญญาเก่าที่ยังไม่ได้ส่ง) → ส่งทันที
    const isExactDay = scheduledDate.getTime() === today.getTime();
    if (isExactDay && currentHour < scheduledHour) continue;

    const contract = contractMap[contractId];
    if (!contract || contract.status !== "active") continue;

    // คำนวณวันที่เหลือ
    const endDate = new Date(contract.end_date);
    const daysLeft = Math.ceil((endDate - today) / (1000 * 60 * 60 * 24));

    let emailSuccess = true;
    let teamsSuccess = true;
    let lineSuccess = true;

    // ส่ง Email
    if (notifyEmail) {
      // Get token once per execution session
      if (!cachedEmailToken) {
        try { cachedEmailToken = getAccessToken(); } catch(e) { Logger.log("Email Auth Error: " + e.message); }
      }
      emailSuccess = sendOutlookEmail(contract, daysLeft, alertDays, cachedEmailToken);
      logNotification(ruleId, contractId, "email",
                      emailSuccess ? "success" : "failed");
    }

    // ส่ง Teams
    if (notifyTeams) {
      teamsSuccess = sendTeamsMessage(contract, daysLeft, alertDays);
      logNotification(ruleId, contractId, "teams",
                      teamsSuccess ? "success" : "failed");
    }

    // ส่ง LINE
    if (contract.notify_line && contract.line_group_id) {
      lineSuccess = sendLineMessage(contract, daysLeft, alertDays);
      logNotification(ruleId, contractId, "line",
                      lineSuccess ? "success" : "failed");
    }

    // Mark is_sent = TRUE ถ้าส่งสำเร็จทุกช่องทาง
    if (emailSuccess && teamsSuccess && lineSuccess) {
      rulesSheet.getRange(i + 1, 7).setValue(true);       // is_sent
      rulesSheet.getRange(i + 1, 8).setValue(new Date()); // sent_at
    }

    // Rate Limiting: พัก 1 วินาทีถ้ามีการส่งแจ้งเตือน
    if (notifyEmail || notifyTeams || (contract.notify_line && contract.line_group_id)) {
      Utilities.sleep(1000); 
    }
  }
}
