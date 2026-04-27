// ============================================================
//  MAIN SCHEDULER — รันทุกวัน 08:00
// ============================================================

function checkAndSendNotifications() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const rulesSheet = ss.getSheetByName("notification_rules");
  const contractsSheet = ss.getSheetByName("warranty_contracts");

  const today = new Date();
  today.setHours(0, 0, 0, 0);

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
  for (let i = 1; i < rulesData.length; i++) {
    const rule = rulesData[i];
    const ruleId           = rule[0];
    const contractId       = rule[1];
    const alertDays        = rule[2];
    const scheduledDate    = new Date(rule[3]);
    const notifyEmail      = rule[4];
    const notifyTeams      = rule[5];
    const isSent           = rule[6];

    scheduledDate.setHours(0, 0, 0, 0);

    // ข้ามถ้าส่งแล้ว หรือยังไม่ถึงวัน
    if (isSent || scheduledDate > today) continue;

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

    // Rate Limiting: พัก 1 วินาทีถ้ามีการส่งแจ้งเตือน เพื่อป้องกันโควต้า Google เต็ม (Bandwidth/Rate limit)
    if (notifyEmail || notifyTeams || (contract.notify_line && contract.line_group_id)) {
      Utilities.sleep(1000); 
    }
  }
}
