// ── Utility Functions ────────────────────────────────────────

const $ = id => document.getElementById(id);
  function calcDaysLeft(endDateStr) {
    const end = new Date(endDateStr);
    end.setHours(0, 0, 0, 0);
    return Math.ceil((end - today) / (1000 * 60 * 60 * 24));
  }

  function getStatusCategory(daysLeft, status) {
    if (status === 'expired' || status === 'cancel') return 'expired';
    if (daysLeft <= 0) return 'expired';
    if (daysLeft <= 30) return 'critical';
    if (daysLeft <= 90) return 'warning';
    return 'safe';
  }

  function formatDate(dateStr) {
    if (!dateStr) return '-';
    const d = new Date(dateStr);
    const dd = String(d.getDate()).padStart(2, '0');
    const mm = String(d.getMonth() + 1).padStart(2, '0');
    const yyyy = d.getFullYear();
    return `${dd}/${mm}/${yyyy}`;
  }

  function formatThaiDate(date) {
    const months = ['ม.ค.', 'ก.พ.', 'มี.ค.', 'เม.ย.', 'พ.ค.', 'มิ.ย.',
      'ก.ค.', 'ส.ค.', 'ก.ย.', 'ต.ค.', 'พ.ย.', 'ธ.ค.'];
    const d = new Date(date);
    return `${d.getDate()} ${months[d.getMonth()]} ${d.getFullYear() + 543}`;
  }

  function generateId(prefix) {
    return `${prefix}-${Date.now()}`;
  }

  