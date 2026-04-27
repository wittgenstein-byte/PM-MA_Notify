// ============================================================
//  ContractsUI — SRP: รับผิดชอบเฉพาะ render ตาราง contracts
//  OCP: เพิ่ม filter ใหม่ได้โดยไม่แก้ class หลัก
//  GRASP: Low Coupling — รับ callback แทนการเรียก module อื่นโดยตรง
// ============================================================

class ContractsUI {
  /**
   * @param {Function} onRowClick - callback(contractId)
   * @param {Function} renderStatusBadge
   */
  constructor(onRowClick, renderStatusBadge) {
    this._onRowClick = onRowClick;
    this._renderStatusBadge = renderStatusBadge;
    this._currentFilter = 'all';
    this._contracts = [];
    this._bindFilterTabs();
    this._bindSearch();
  }

  /**
   * Render ตาราง contracts
   * @param {Array} contracts - raw contract array
   * @param {string} [filter='all']
   */
  render(contracts, filter = this._currentFilter) {
    this._contracts = contracts;
    this._currentFilter = filter;

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
          <td colspan="10" style="text-align:center;padding:40px;color:var(--text-muted)">
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
        <td>${this._renderStatusBadge(c.category)}</td>
        <td>
          <button class="btn-icon" title="ดูรายละเอียด" data-view="${c.contract_id}">
            <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2">
              <path d="M1 12s4-8 11-8 11 8 11 8-4 8-11 8-11-8-11-8z"/>
              <circle cx="12" cy="12" r="3"/>
            </svg>
          </button>
        </td>
      </tr>
    `).join('');

    tbody.querySelectorAll('tr').forEach(row => {
      row.addEventListener('click', e => {
        if (e.target.closest('.btn-icon')) return;
        this._onRowClick(row.dataset.id);
      });
    });

    tbody.querySelectorAll('.btn-icon').forEach(btn => {
      btn.addEventListener('click', () => this._onRowClick(btn.dataset.view));
    });
  }

  // ── Private ───────────────────────────────────────────────

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
