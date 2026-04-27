// ============================================================
//  ToastUI — SRP: รับผิดชอบเฉพาะการแสดง toast notification
//  GRASP: High Cohesion — ทุก method เกี่ยวกับ toast เท่านั้น
// ============================================================

class ToastUI {
  constructor(containerId = 'toastContainer') {
    this._container = document.getElementById(containerId);
  }

  /**
   * แสดง toast message
   * @param {'success'|'error'|'warning'} type
   * @param {string} message
   * @param {number} durationMs
   */
  show(type, message, durationMs = 3200) {
    if (!this._container) return;

    const toast = document.createElement('div');
    toast.className = `toast ${type}`;
    const icon = type === 'success' ? '✅' : type === 'warning' ? '⚠️' : '❌';
    toast.innerHTML = `<span>${icon}</span><span>${message}</span>`;
    this._container.appendChild(toast);

    setTimeout(() => toast.remove(), durationMs);
  }

  success(message) { this.show('success', message); }
  error(message)   { this.show('error', message); }
  warning(message) { this.show('warning', message); }
}
