// ============================================================
//  API ENDPOINTS — รับข้อมูลจาก Dashboard (Vercel)
//  Deploy เป็น Web App: Execute as "Me", Access "Anyone"
// ============================================================

/**
 * doGet — Health check + ดึงข้อมูล contracts/logs
 * GET ?action=getContracts
 * GET ?action=getLogs
 */
function doGet(e) {
  const action = e && e.parameter && e.parameter.action;

  try {
    if (action === 'getContracts') {
      const data = getContractsData();
      return jsonResponse({ status: 'ok', data });
    }
    if (action === 'getLogs') {
      const data = getLogsData();
      return jsonResponse({ status: 'ok', data });
    }
    // Health check
    return jsonResponse({ status: 'ok', message: 'PM-MA Notify API is running' });
  } catch (err) {
    return jsonResponse({ status: 'error', message: err.message }, 500);
  }
}

/**
 * doPost — รับข้อมูลสัญญาใหม่จาก Dashboard แล้วบันทึกลง Sheets
 * Body (JSON): { action: 'addContract', data: { ...contract fields } }
 */
function doPost(e) {
  try {
    const payload = JSON.parse(e.postData.contents);
    const action = payload.action;

    if (action === 'addContract') {
      const result = addContractToSheet(payload.data);
      return jsonResponse({ status: 'ok', contract_id: result.contract_id, message: 'บันทึกสัญญาเรียบร้อย' });
    }

    return jsonResponse({ status: 'error', message: 'Unknown action: ' + action }, 400);
  } catch (err) {
    return jsonResponse({ status: 'error', message: err.message }, 500);
  }
}

// ── Helpers ─────────────────────────────────────────────────

function jsonResponse(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

/**
 * บันทึกสัญญาใหม่ลง warranty_contracts + สร้าง notification_rules
 */
function addContractToSheet(data) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const contractsSheet = ss.getSheetByName('warranty_contracts');

  // สร้าง contract_id
  const contractId = 'WC-' + new Date().getFullYear() + '-' + String(Date.now()).slice(-6);
  const now = new Date();

  // คอลัมน์ตาม schema: contract_id | po_number | project_name | customer_name |
  //   service_type | start_date | end_date | recipients_sale | recipients_eng |
  //   teams_webhook | note | status | line_group_id | notify_line |
  //   created_by | created_at | updated_at
  contractsSheet.appendRow([
    contractId,
    data.po_number       || '',
    data.project_name    || '',
    data.customer_name   || '',
    data.service_type    || '',
    data.start_date      || '',
    data.end_date        || '',
    data.recipients_sale || '',
    data.recipients_eng  || '',
    data.teams_webhook   || '',
    data.note            || '',
    'active',
    data.line_group_id   || '',
    data.line_group_id ? true : false,
    data.created_by      || 'dashboard',
    now,
    now,
  ]);

  // สร้าง notification_rules
  const alertDays = data.alert_days && data.alert_days.length > 0
    ? data.alert_days
    : [90, 60, 30, 7];

  createNotificationRules(contractId, data.end_date, alertDays);

  return { contract_id: contractId };
}

/**
 * ดึงข้อมูล contracts ทั้งหมดเพื่อแสดงบน dashboard
 */
function getContractsData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('warranty_contracts');
  const rows = sheet.getDataRange().getValues();

  const headers = rows[0];
  const contracts = [];
  for (let i = 1; i < rows.length; i++) {
    const row = rows[i];
    if (!row[0]) continue; // skip empty rows
    contracts.push({
      contract_id:     row[0],
      po_number:       row[1],
      project_name:    row[2],
      customer_name:   row[3],
      service_type:    row[4],
      start_date:      row[5] ? Utilities.formatDate(new Date(row[5]), 'Asia/Bangkok', 'yyyy-MM-dd') : '',
      end_date:        row[6] ? Utilities.formatDate(new Date(row[6]), 'Asia/Bangkok', 'yyyy-MM-dd') : '',
      recipients_sale: row[7],
      recipients_eng:  row[8],
      teams_webhook:   row[9],
      note:            row[10],
      status:          row[11],
      line_group_id:   row[12],
      notify_line:     row[13],
    });
  }
  return contracts;
}

/**
 * ดึง notification_logs ล่าสุด 50 รายการ
 */
function getLogsData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('notification_logs');
  const rows = sheet.getDataRange().getValues();

  const logs = [];
  for (let i = Math.max(1, rows.length - 50); i < rows.length; i++) {
    const row = rows[i];
    if (!row[0]) continue;
    logs.push({
      log_id:      row[0],
      rule_id:     row[1],
      contract_id: row[2],
      channel:     row[3],
      status:      row[4],
      error_msg:   row[5],
      sent_at:     row[6] ? Utilities.formatDate(new Date(row[6]), 'Asia/Bangkok', 'yyyy-MM-dd HH:mm') : '',
    });
  }
  return logs.reverse(); // ล่าสุดขึ้นก่อน
}
