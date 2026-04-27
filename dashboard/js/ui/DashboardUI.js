// ============================================================
//  DashboardUI — SRP: รับผิดชอบเฉพาะ render หน้า Dashboard
//  DIP: รับ contracts, logs เป็น parameter — ไม่ใช้ global โดยตรง
//  GRASP: Low Coupling — ไม่รู้จัก ContractsUI หรือ modal
// ============================================================

class DashboardUI {
  /**
   * @param {Function} onRowClick - callback(contractId) เมื่อคลิก row
   * @param {Function} renderStatusBadge
   */
  constructor(onRowClick, renderStatusBadge) {
    this._onRowClick = onRowClick;
    this._renderStatusBadge = renderStatusBadge;
  }

  /**
   * Render dashboard page ทั้งหมด
   * @param {Array} contracts - raw contract array
   * @param {Array} logs - raw logs array
   */
  render(contracts, logs) {
    const enriched = this._enrich(contracts);
    this._renderStats(enriched, contracts);
    this._renderTable(enriched);
    this._renderFeed(logs, contracts);
  }

  // ── Private ───────────────────────────────────────────────

  _enrich(contracts) {
    return contracts.map(c => {
      const daysLeft = calcDaysLeft(c.end_date);
      const category = getStatusCategory(daysLeft, c.status);
      return { ...c, daysLeft, category };
    });
  }

  _renderStats(enriched, allContracts) {
    let critical = 0, warning = 0, safe = 0;

    enriched.forEach(c => {
      if (c.status !== 'active') return;
      if (c.category === 'critical') critical++;
      else if (c.category === 'warning') warning++;
      else if (c.category === 'safe') safe++;
    });

    const expiredCount = allContracts.filter(
      c => c.status === 'expired' || c.status === 'cancel'
    ).length;
    const total = enriched.filter(c => c.status === 'active').length + expiredCount;

    animateCounter('criticalCount', critical);
    animateCounter('warningCount', warning);
    animateCounter('safeCount', safe);
    animateCounter('totalCount', total);
  }

  _renderTable(enriched) {
    const sorted = enriched
      .filter(c => c.status === 'active')
      .sort((a, b) => a.daysLeft - b.daysLeft)
      .slice(0, 8);

    const tbody = document.getElementById('dashboardTableBody');
    tbody.innerHTML = sorted.map(c => `
      <tr data-id="${c.contract_id}">
        <td><code style="color:var(--primary-400)">${c.po_number}</code></td>
        <td>${c.project_name}</td>
        <td>${c.customer_name}</td>
        <td><span class="badge badge-${c.service_type === 'MA+PM' ? 'email' : 'teams'}">${c.service_type}</span></td>
        <td>${formatDate(c.end_date)}</td>
        <td><span class="days-remaining ${c.category}">${c.daysLeft} วัน</span></td>
        <td>${this._renderStatusBadge(c.category)}</td>
      </tr>
    `).join('');

    tbody.querySelectorAll('tr').forEach(row => {
      row.addEventListener('click', () => this._onRowClick(row.dataset.id));
    });
  }

  _renderFeed(logs, contracts) {
    const feed = document.getElementById('notificationFeed');
    const recentLogs = [...logs].slice(-6).reverse();

    feed.innerHTML = recentLogs.map(log => {
      const contract = contracts.find(c => c.contract_id === log.contract_id);
      const icons = { email: '📧', teams: '💬', line: '💚' };
      const names = { email: 'Outlook Email', teams: 'MS Teams', line: 'LINE' };
      return `
        <div class="notif-item">
          <div class="notif-icon ${log.channel}">${icons[log.channel] || '🔔'}</div>
          <div class="notif-content">
            <div class="notif-title">${names[log.channel] || log.channel} — ${contract ? contract.project_name : log.contract_id}</div>
            <div class="notif-desc">
              <span class="badge badge-${log.status === 'success' ? 'success' : 'failed'}">${log.status}</span>
              ${log.error_msg ? `<span style="color:var(--red-400);margin-left:8px">${log.error_msg}</span>` : ''}
            </div>
          </div>
          <span class="notif-time">${log.sent_at}</span>
        </div>
      `;
    }).join('');
  }
}

// ── Helper: animate counter (ใช้ร่วมกันใน DashboardUI) ───────
function animateCounter(elementId, target) {
  const el = document.getElementById(elementId);
  if (!el) return;
  let current = 0;
  const step = Math.max(1, Math.ceil(target / 20));
  const interval = setInterval(() => {
    current += step;
    if (current >= target) { current = target; clearInterval(interval); }
    el.textContent = current;
  }, 30);
}
