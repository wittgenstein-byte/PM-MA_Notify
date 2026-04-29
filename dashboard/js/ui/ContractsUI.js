// ============================================================
//  ContractsUI — SRP: รับผิดชอบเฉพาะ render ตาราง contracts
//  OCP: เพิ่ม filter ใหม่ได้โดยไม่แก้ class หลัก
//  GRASP: Low Coupling — รับ callback แทนการเรียก module อื่นโดยตรง
// ============================================================

class ContractsUI {
  /**
   * @param {Function} onRowClick - callback(contractId)
   * @param {Function} renderStatusBadge
   * @param {Function} onEdit - callback(contractId)
   * @param {Function} onDelete - callback(contractId)
   */
  constructor(onRowClick, renderStatusBadge, onEdit, onDelete) {
    this._onRowClick = onRowClick;
    this._renderStatusBadge = renderStatusBadge;
    this._onEdit = onEdit || (() => {});
    this._onDelete = onDelete || (() => {});
    this._currentFilter = 'all';
    this._contracts = [];
    this._scheduleMap = {};
    this._bindFilterTabs();
    this._bindSearch();
  }

  /**
   * Render ตาราง contracts
   * @param {Array} contracts - raw contract array
   * @param {string} [filter='all']
   * @param {Object} [scheduleMap={}] - { contract_id: [{ date, time, days, is_sent }] }
   */
  render(contracts, filter = this._currentFilter, scheduleMap = this._scheduleMap) {
    this._contracts = contracts;
    this._currentFilter = filter;
    this._scheduleMap = scheduleMap;

    const enriched = contracts.map(c => ({
      ...c,
      daysLeft: calcDaysLeft(c.end_date),
      category: getStatusCategory(calcDaysLeft(c.end_date), c.status),
    }));

    let filtered = filter === 'all' ? enriched : enriched.filter(c => c.category === filter);

    // Apply search
    const searchTerm = (document.getElementById('searchInput')?.value || '').toLowerCase();
    if (searchTerm) {
      filtered = filtered.filter(c =>
        c.po_number.toLowerCase().includes(searchTerm) ||
        c.project_name.toLowerCase().includes(searchTerm) ||
        c.customer_name.toLowerCase().includes(searchTerm) ||
        c.contract_id.toLowerCase().includes(searchTerm)
      );
    }

    const tbody = document.getElementById('contractsTableBody');

    if (filtered.length === 0) {
      tbody.innerHTML = `
        <tr>
          <td colspan="11" style="text-align:center;padding:40px;color:var(--text-muted)">
            ไม่พบข้อมูลสัญญา
          </td>
        </tr>`;
      return;
    }

    tbody.innerHTML = filtered.sort((a, b) => a.daysLeft - b.daysLeft).map(c => `
      <tr data-id="${c.contract_id}">
        <td><code style="font-size:0.75rem;color:var(--text-muted)">${c.contract_id}</code></td>
        <td><code style="color:var(--primary-400)">${c.po_number}</code></td>
        <td style="max-width:200px;overflow:hidden;text-overflow:ellipsis">${c.project_name}</td>
        <td>${c.customer_name}</td>
        <td><span class="badge badge-${c.service_type === 'MA+PM' ? 'email' : 'teams'}">${c.service_type}</span></td>
        <td>${formatDate(c.start_date)}</td>
        <td>${formatDate(c.end_date)}</td>
        <td><span class="days-remaining ${c.category}">${c.daysLeft > 0 ? c.daysLeft + ' วัน' : 'หมดอายุ'}</span></td>
        <td>${this._renderScheduleCell(c.contract_id)}</td>
        <td>${this._renderStatusBadge(c.category)}</td>
        <td>
          <div class="action-buttons">
            <button class="btn-icon" title="ดูรายละเอียด" data-view="${c.contract_id}">
              <svg width="15" height="15" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2">
                <path d="M1 12s4-8 11-8 11 8 11 8-4 8-11 8-11-8-11-8z"/>
                <circle cx="12" cy="12" r="3"/>
              </svg>
            </button>
            <button class="btn-icon btn-edit" title="แก้ไข" data-edit="${c.contract_id}">
              <svg width="15" height="15" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2">
                <path d="M11 4H4a2 2 0 0 0-2 2v14a2 2 0 0 0 2 2h14a2 2 0 0 0 2-2v-7"/>
                <path d="M18.5 2.5a2.121 2.121 0 0 1 3 3L12 15l-4 1 1-4 9.5-9.5z"/>
              </svg>
            </button>
            <button class="btn-icon btn-delete" title="ลบ" data-delete="${c.contract_id}">
              <svg width="15" height="15" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2">
                <polyline points="3 6 5 6 21 6"/>
                <path d="M19 6v14a2 2 0 0 1-2 2H7a2 2 0 0 1-2-2V6m3 0V4a2 2 0 0 1 2-2h4a2 2 0 0 1 2 2v2"/>
              </svg>
            </button>
          </div>
        </td>
      </tr>
    `).join('');

    // Bind row click (view modal)
    tbody.querySelectorAll('tr').forEach(row => {
      row.addEventListener('click', e => {
        if (e.target.closest('.btn-icon')) return;
        this._onRowClick(row.dataset.id);
      });
    });

    // Bind view buttons
    tbody.querySelectorAll('[data-view]').forEach(btn => {
      btn.addEventListener('click', e => {
        e.stopPropagation();
        this._onRowClick(btn.dataset.view);
      });
    });

    // Bind edit buttons
    tbody.querySelectorAll('[data-edit]').forEach(btn => {
      btn.addEventListener('click', e => {
        e.stopPropagation();
        this._onEdit(btn.dataset.edit);
      });
    });

    // Bind delete buttons
    tbody.querySelectorAll('[data-delete]').forEach(btn => {
      btn.addEventListener('click', e => {
        e.stopPropagation();
        this._onDelete(btn.dataset.delete);
      });
    });
  }

  // ── Private ───────────────────────────────────────────────

  /**
   * Render notification schedule cell for a contract
   * Shows upcoming and past notifications with badges
   */
  _renderScheduleCell(contractId) {
    const schedules = this._scheduleMap[contractId];
    if (!schedules || schedules.length === 0) {
      return '<span style="color:var(--text-muted);font-size:0.75rem">—</span>';
    }

    const todayTime = today.getTime();

    return '<div class="schedule-badges">' + schedules.map(s => {
      const sDate = new Date(s.date);
      const isPast = sDate.getTime() < todayTime;
      const isToday = sDate.toISOString().split('T')[0] === today.toISOString().split('T')[0];

      let cls = 'schedule-badge';
      let icon = '';

      if (s.is_sent) {
        cls += ' sent';
        icon = '✅';
      } else if (isToday) {
        cls += ' today';
        icon = '🔔';
      } else if (isPast) {
        cls += ' overdue';
        icon = '⚠️';
      } else {
        cls += ' pending';
        icon = '⏳';
      }

      const dayLabel = s.days === 0 ? 'วันหมด' : `${s.days}d`;
      const dateLabel = formatDate(s.date);

      return `<span class="${cls}" title="${dateLabel} เวลา ${s.time} น. (ก่อนหมด ${s.days} วัน)">${icon} ${dayLabel} <small>${s.time}</small></span>`;
    }).join('') + '</div>';
  }

  _bindFilterTabs() {
    document.querySelectorAll('.filter-tab').forEach(tab => {
      tab.addEventListener('click', () => {
        document.querySelectorAll('.filter-tab').forEach(t => t.classList.remove('active'));
        tab.classList.add('active');
        this.render(this._contracts, tab.dataset.filter);
      });
    });
  }

  _bindSearch() {
    document.getElementById('searchInput')?.addEventListener('input', () => {
      this.render(this._contracts, this._currentFilter);
    });
  }
}
