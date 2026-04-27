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

  /**
   * ตรวจว่า service ทำงานอยู่ในโหมด demo (ยังไม่มี URL)
   * @returns {boolean}
   */
  isDemoMode() {
    return !this._url || this._url === 'YOUR_GAS_WEB_APP_URL_HERE';
  }
}
