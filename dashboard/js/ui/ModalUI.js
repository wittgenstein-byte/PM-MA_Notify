// ============================================================
//  ModalUI — SRP: รับผิดชอบเฉพาะ contract detail modal
//  DIP: รับ renderStatusBadge เป็น dependency injection
//  GRASP: Information Expert — รู้ว่าต้อง render อะไรใน modal
// ============================================================

class ModalUI {
  /**
   * @param {Function} renderStatusBadge - injected renderer function
   * @param {Function} onEdit - callback(contractId)
   * @param {Function} onDelete - callback(contractId)
   */
  constructor(renderStatusBadge, onEdit, onDelete) {
    this._renderStatusBadge = renderStatusBadge;
    this._onEdit = onEdit || (() => {});
    this._onDelete = onDelete || (() => {});
    this._overlay = document.getElementById('modalOverlay');
    this._title   = document.getElementById('modalTitle');
    this._body    = document.getElementById('modalBody');
    this._currentContractId = null;
    this._bindCloseEvents();
  }

  /**
   * เปิด modal แสดงรายละเอียดสัญญา
   * @param {Object} contract - enriched contract object (includes daysLeft, category)
   */
  open(contract) {
    const { daysLeft, category } = contract;
    this._currentContractId = contract.contract_id;
    this._title.textContent = `📄 ${contract.po_number}`;
    this._body.innerHTML = this._buildHTML(contract, daysLeft, category);
    this._overlay.classList.add('active');

    // Bind action buttons inside modal
    this._body.querySelector('[data-modal-edit]')?.addEventListener('click', () => {
      this.close();
      this._onEdit(contract.contract_id);
    });
    this._body.querySelector('[data-modal-delete]')?.addEventListener('click', () => {
      this._onDelete(contract.contract_id);
    });
  }

  close() {
    this._overlay.classList.remove('active');
    this._currentContractId = null;
  }

  _bindCloseEvents() {
    document.getElementById('modalClose')?.addEventListener('click', () => this.close());
    this._overlay?.addEventListener('click', e => {
      if (e.target === this._overlay) this.close();
    });
    document.addEventListener('keydown', e => {
      if (e.key === 'Escape') this.close();
    });
  }

  _buildHTML(contract, daysLeft, category) {
    const row = (label, value) =>
      `<div class="detail-row">
        <span class="detail-label">${label}</span>
        <span class="detail-value">${value}</span>
      </div>`;

    return `
      ${row('Contract ID', contract.contract_id)}
      ${row('PO Number', `<span style="color:var(--primary-400)">${contract.po_number}</span>`)}
      ${row('📁 โครงการ', contract.project_name)}
      ${row('🏢 ลูกค้า', contract.customer_name)}
      ${row('🛡️ ประเภท', contract.service_type)}
      ${row('📅 เริ่มต้น', formatDate(contract.start_date))}
      ${row('📅 หมดอายุ', `<span style="color:var(--red-400);font-weight:700">${formatDate(contract.end_date)}</span>`)}
      ${row('⏳ เหลืออีก', `<span class="days-remaining ${category}" style="font-size:1.1rem">${daysLeft > 0 ? daysLeft + ' วัน' : 'หมดอายุแล้ว'}</span>`)}
      ${row('สถานะ', this._renderStatusBadge(category))}
      ${row('📧 Email Sale', contract.recipients_sale || '-')}
      ${row('📧 Email Engineer', contract.recipients_eng || '-')}
      ${row('💬 Teams Webhook', `<span style="font-size:0.75rem;max-width:300px;word-break:break-all">${contract.teams_webhook || '-'}</span>`)}
      ${row('💚 LINE Group', contract.line_group_id
        ? '<span class="badge badge-line">✅ เชื่อมต่อแล้ว</span>'
        : '<span style="color:var(--text-muted)">— ไม่ได้ตั้งค่า</span>')}
      ${contract.note ? row('📝 หมายเหตุ', contract.note) : ''}

      <div class="modal-actions">
        <button class="btn btn-primary btn-sm" data-modal-edit title="แก้ไขสัญญานี้">
          <svg width="15" height="15" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2">
            <path d="M11 4H4a2 2 0 0 0-2 2v14a2 2 0 0 0 2 2h14a2 2 0 0 0 2-2v-7"/>
            <path d="M18.5 2.5a2.121 2.121 0 0 1 3 3L12 15l-4 1 1-4 9.5-9.5z"/>
          </svg>
          แก้ไข
        </button>
        <button class="btn btn-danger btn-sm" data-modal-delete title="ลบสัญญานี้">
          <svg width="15" height="15" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2">
            <polyline points="3 6 5 6 21 6"/>
            <path d="M19 6v14a2 2 0 0 1-2 2H7a2 2 0 0 1-2-2V6m3 0V4a2 2 0 0 1 2-2h4a2 2 0 0 1 2 2v2"/>
          </svg>
          ลบ
        </button>
      </div>
    `;
  }
}
