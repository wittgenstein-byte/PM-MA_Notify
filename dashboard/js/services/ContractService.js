// ============================================================
//  ContractService — GRASP: Information Expert + Pure Fabrication
//  SRP: รับผิดชอบเฉพาะการสื่อสารกับ Google Apps Script API
//  DIP: รับ gasUrl ผ่าน constructor — ไม่พึ่ง global variable
// ============================================================

class ContractService {
  /**
   * @param {string} gasUrl - GAS Web App URL (injected, not hardcoded)
   */
  constructor(gasUrl) {
    this._url = gasUrl;
  }

  // ── READ ─────────────────────────────────────────────────

  /**
   * ดึงข้อมูลสัญญาทั้งหมดจาก Google Sheets
   * @returns {Promise<{status:string, data?:Array}>}
   */
  async fetchContracts() {
    if (!this._url) return { status: 'demo', data: [] };

    const res = await fetch(`${this._url}?action=getContracts`, {
      redirect: 'follow',
    });
    return res.json();
  }

  /**
   * ดึง notification logs จาก Google Sheets
   * @returns {Promise<{status:string, data?:Array}>}
   */
  async fetchLogs() {
    if (!this._url) return { status: 'demo', data: [] };

    const res = await fetch(`${this._url}?action=getLogs`, {
      redirect: 'follow',
    });
    return res.json();
  }

  // ── CREATE ───────────────────────────────────────────────

  /**
   * ส่ง contract ใหม่ไปยัง GAS
   * @param {Object} contractData
   * @returns {Promise<{status:'ok'|'error', contract_id?:string, message?:string}>}
   */
  async addContract(contractData) {
    if (!this._url) {
      return { status: 'demo', message: 'No GAS URL configured' };
    }

    const res = await fetch(this._url, {
      method: 'POST',
      headers: { 'Content-Type': 'text/plain;charset=utf-8' },
      body: JSON.stringify({ action: 'addContract', data: contractData }),
      redirect: 'follow',
    });

    return res.json();
  }

  // ── UPDATE ───────────────────────────────────────────────

  /**
   * อัพเดทข้อมูลสัญญาที่มีอยู่
   * @param {string} contractId
   * @param {Object} data - ข้อมูลที่เปลี่ยน
   * @returns {Promise<{status:string, message?:string}>}
   */
  async updateContract(contractId, data) {
    if (!this._url) {
      return { status: 'demo', message: 'No GAS URL configured' };
    }

    const res = await fetch(this._url, {
      method: 'POST',
      headers: { 'Content-Type': 'text/plain;charset=utf-8' },
      body: JSON.stringify({ action: 'updateContract', contract_id: contractId, data }),
      redirect: 'follow',
    });

    return res.json();
  }

  // ── DELETE ───────────────────────────────────────────────

  /**
   * ลบสัญญา + notification_rules ที่เกี่ยวข้อง
   * @param {string} contractId
   * @returns {Promise<{status:string, message?:string}>}
   */
  async deleteContract(contractId) {
    if (!this._url) {
      return { status: 'demo', message: 'No GAS URL configured' };
    }

    const res = await fetch(this._url, {
      method: 'POST',
      headers: { 'Content-Type': 'text/plain;charset=utf-8' },
      body: JSON.stringify({ action: 'deleteContract', contract_id: contractId }),
      redirect: 'follow',
    });

    return res.json();
  }

  // ── UTILITY ──────────────────────────────────────────────

  /**
   * ตรวจว่า service ทำงานอยู่ในโหมด demo (ยังไม่มี URL)
   * @returns {boolean}
   */
  isDemoMode() {
    return !this._url || this._url === 'YOUR_GAS_WEB_APP_URL_HERE';
  }
}
