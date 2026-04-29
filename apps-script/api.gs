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

    if (action === 'updateContract') {
      const result = updateContractInSheet(payload.contract_id, payload.data);
      return jsonResponse({ status: 'ok', contract_id: result.contract_id, message: 'อัพเดทสัญญาเรียบร้อย' });
    }

    if (action === 'deleteContract') {
      deleteContractFromSheet(payload.contract_id);
      return jsonResponse({ status: 'ok', message: 'ลบสัญญาเรียบร้อย' });
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

  // ส่ง notify_time จาก frontend (เช่น "08:00", "09:30")
  const notifyTime = data.notify_time || '08:00';
  createNotificationRules(contractId, data.end_date, alertDays, notifyTime);

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

// ============================================================
//  UPDATE CONTRACT — อัพเดทข้อมูลสัญญาที่มีอยู่
// ============================================================

function updateContractInSheet(contractId, data) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('warranty_contracts');
  const rows = sheet.getDataRange().getValues();

  // หาแถวที่ตรงกับ contract_id
  let rowIndex = -1;
  for (let i = 1; i < rows.length; i++) {
    if (rows[i][0] === contractId) {
      rowIndex = i + 1; // sheet row is 1-indexed + header
      break;
    }
  }

  if (rowIndex === -1) {
    throw new Error('ไม่พบสัญญา: ' + contractId);
  }

  const now = new Date();

  // อัพเดทแต่ละคอลัมน์ (เฉพาะที่ส่งมา)
  // Schema: contract_id(1) | po_number(2) | project_name(3) | customer_name(4) |
  //   service_type(5) | start_date(6) | end_date(7) | recipients_sale(8) | recipients_eng(9) |
  //   teams_webhook(10) | note(11) | status(12) | line_group_id(13) | notify_line(14) |
  //   created_by(15) | created_at(16) | updated_at(17)

  if (data.po_number !== undefined)       sheet.getRange(rowIndex, 2).setValue(data.po_number);
  if (data.project_name !== undefined)    sheet.getRange(rowIndex, 3).setValue(data.project_name);
  if (data.customer_name !== undefined)   sheet.getRange(rowIndex, 4).setValue(data.customer_name);
  if (data.service_type !== undefined)    sheet.getRange(rowIndex, 5).setValue(data.service_type);
  if (data.start_date !== undefined)      sheet.getRange(rowIndex, 6).setValue(data.start_date);
  if (data.end_date !== undefined)        sheet.getRange(rowIndex, 7).setValue(data.end_date);
  if (data.recipients_sale !== undefined) sheet.getRange(rowIndex, 8).setValue(data.recipients_sale);
  if (data.recipients_eng !== undefined)  sheet.getRange(rowIndex, 9).setValue(data.recipients_eng);
  if (data.teams_webhook !== undefined)   sheet.getRange(rowIndex, 10).setValue(data.teams_webhook);
  if (data.note !== undefined)            sheet.getRange(rowIndex, 11).setValue(data.note);
  if (data.status !== undefined)          sheet.getRange(rowIndex, 12).setValue(data.status);
  if (data.line_group_id !== undefined) {
    sheet.getRange(rowIndex, 13).setValue(data.line_group_id);
    sheet.getRange(rowIndex, 14).setValue(!!data.line_group_id);
  }

  // updated_at
  sheet.getRange(rowIndex, 17).setValue(now);

  // ถ้า regenerate_rules = true → ลบ rules เก่า สร้างใหม่พร้อม notify_time
  if (data.regenerate_rules) {
    deleteRulesForContract(contractId);
    const alertDays = data.alert_days && data.alert_days.length > 0
      ? data.alert_days
      : [90, 60, 30, 7];
    const endDateForRules = data.end_date || sheet.getRange(rowIndex, 7).getValue();
    const notifyTime = data.notify_time || '08:00';
    createNotificationRules(contractId, endDateForRules, alertDays, notifyTime);
  }

  return { contract_id: contractId };
}

// ============================================================
//  DELETE CONTRACT — ลบสัญญา + notification_rules ที่เกี่ยวข้อง
// ============================================================

function deleteContractFromSheet(contractId) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // 1) ลบ notification_rules ที่เกี่ยวข้อง
  deleteRulesForContract(contractId);

  // 2) ลบ contract row
  const contractsSheet = ss.getSheetByName('warranty_contracts');
  const contractRows = contractsSheet.getDataRange().getValues();

  for (let i = contractRows.length - 1; i >= 1; i--) {
    if (contractRows[i][0] === contractId) {
      contractsSheet.deleteRow(i + 1);
      break;
    }
  }
}

/**
 * ลบ notification_rules ทั้งหมดที่เป็นของ contractId
 */
function deleteRulesForContract(contractId) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const rulesSheet = ss.getSheetByName('notification_rules');
  const rulesRows = rulesSheet.getDataRange().getValues();

  // ลบจากล่างขึ้นบน เพื่อไม่ให้ row index เลื่อน
  for (let i = rulesRows.length - 1; i >= 1; i--) {
    if (rulesRows[i][1] === contractId) {
      rulesSheet.deleteRow(i + 1);
    }
  }
}

