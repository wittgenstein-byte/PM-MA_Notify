// ============================================================
//  LogsUI — SRP: รับผิดชอบเฉพาะ render ตาราง notification logs
//  GRASP: High Cohesion — ทำแค่เรื่อง logs
// ============================================================

class LogsUI {
  render(logs) {
    const tbody = document.getElementById('logsTableBody');
    if (!tbody) return;

    const sortedLogs = [...logs].sort((a, b) => new Date(b.sent_at) - new Date(a.sent_at));
    tbody.innerHTML = sortedLogs.map(log => `
      <tr>
        <td><code style="font-size:0.75rem;color:var(--text-muted)">${log.log_id}</code></td>
        <td><code style="font-size:0.75rem">${log.rule_id}</code></td>
        <td><code style="font-size:0.75rem">${log.contract_id}</code></td>
        <td>
          <span class="badge badge-${log.channel}">
            ${log.channel === 'email' ? '📧 Email' : log.channel === 'line' ? '💚 LINE' : '💬 Teams'}
          </span>
        </td>
        <td><span class="badge badge-${log.status}">${log.status}</span></td>
        <td style="color:${log.error_msg ? 'var(--red-400)' : 'var(--text-muted)'}">
          ${log.error_msg || '-'}
        </td>
        <td>${log.sent_at}</td>
      </tr>
    `).join('');
  }
}
