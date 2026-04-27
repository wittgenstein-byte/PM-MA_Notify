// ============================================================
//  EmailHistoryService — GRASP: Pure Fabrication
//  SRP: รับผิดชอบเฉพาะ email history ใน localStorage
//  ISP: แยกออกจาก EmailTagInput — UI ไม่ต้องรู้ storage detail
// ============================================================

class EmailHistoryService {
  constructor() {
    this._storageKey = 'pm_ma_recent_emails';
  }

  /**
   * รวบรวม emails จาก contracts + localStorage
   * @param {Array} contracts - sampleContracts array
   * @returns {string[]} sorted unique email list
   */
  collectAll(contracts = []) {
    const emailSet = new Set();

    contracts.forEach(c => {
      if (c.recipients_sale) {
        c.recipients_sale.split(',').forEach(e => emailSet.add(e.trim().toLowerCase()));
      }
      if (c.recipients_eng) {
        c.recipients_eng.split(',').forEach(e => emailSet.add(e.trim().toLowerCase()));
      }
    });

    // Merge from localStorage
    try {
      const stored = JSON.parse(localStorage.getItem(this._storageKey) || '[]');
      stored.forEach(e => emailSet.add(e.toLowerCase()));
    } catch (_) { }

    return Array.from(emailSet).filter(e => e && e.includes('@')).sort();
  }

  /**
   * บันทึก email ที่ใช้ล่าสุดลง localStorage
   * @param {string} email
   */
  save(email) {
    try {
      const stored = JSON.parse(localStorage.getItem(this._storageKey) || '[]');
      const lower = email.toLowerCase();
      if (!stored.includes(lower)) {
        stored.unshift(lower);
        if (stored.length > 50) stored.pop();
        localStorage.setItem(this._storageKey, JSON.stringify(stored));
      }
    } catch (_) { }
  }
}
