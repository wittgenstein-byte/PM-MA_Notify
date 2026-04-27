// ============================================================
//  ModalUI — SRP: รับผิดชอบเฉพาะ contract detail modal
//  DIP: รับ renderStatusBadge เป็น dependency injection
//  GRASP: Information Expert — รู้ว่าต้อง render อะไรใน modal
// ============================================================

class ModalUI {
  /**
   * @param {Function} renderStatusBadge - injected renderer function
   */
  constructor(renderStatusBadge) {
    this._renderStatusBadge = renderStatusBadge;
    this._overlay = document.getElementById('modalOverlay');
    this._title   = document.getElementById('modalTitle');
    this._body    = document.getElementById('modalBody');
    this._bindCloseEvents();
  }

  /**
   * เปิด modal แสดงรายละเอียดสัญญา
   * @param {Object} contract - enriched contract object (includes daysLeft, category)
   */
  open(contract) {
    const { daysLeft, category } = contract;
    this._title.textContent = `📄 ${contract.po_number}`;
    this._body.innerHTML = this._buildHTML(contract, daysLeft, category);
    this._overlay.classList.add('active');
  }

  close() {
    this._overlay.classList.remove('active');
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
    `;
  }
}
