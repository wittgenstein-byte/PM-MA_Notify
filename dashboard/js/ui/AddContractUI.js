// ============================================================
//  AddContractUI — SRP: รับผิดชอบเฉพาะ Add/Edit Contract form
//  GRASP: Creator — สร้าง EmailTagInput instances (มี data ที่ต้องการ)
//  DIP: รับ ContractService, ToastUI, router เป็น dependencies
// ============================================================

class AddContractUI {
  /**
   * @param {ContractService} contractService
   * @param {EmailHistoryService} historyService
   * @param {ToastUI} toast
   * @param {Function} getContracts - () => contract[] (getter ปัจจุบัน)
   * @param {Function} onContractAdded - callback(newContract) หลัง save สำเร็จ
   * @param {Function} navigateTo - router.navigateTo
   */
  constructor(contractService, historyService, toast, getContracts, onContractAdded, navigateTo) {
    this._service       = contractService;
    this._history       = historyService;
    this._toast         = toast;
    this._getContracts  = getContracts;
    this._onAdded       = onContractAdded;
    this._navigate      = navigateTo;

    this._customAlertDays = [];
    this._form = document.getElementById('contractForm');

    // ── Edit mode state ──────────────────────────────────
    this._editMode = false;
    this._editContractId = null;
    this._onUpdated = null; // set via setOnUpdated()

    // GRASP: Creator — AddContractUI เป็นผู้สร้าง EmailTagInput
    this._saleEmailTags = new EmailTagInput(
      'saleTagContainer', 'saleTags', 'saleEmailInput', 'saleDropdown',
      this._history,
      () => this._history.collectAll(this._getContracts()),
      () => this._renderRecentEmails()
    );
    this._engEmailTags = new EmailTagInput(
      'engTagContainer', 'engTags', 'engEmailInput', 'engDropdown',
      this._history,
      () => this._history.collectAll(this._getContracts()),
      () => this._renderRecentEmails()
    );

    this._bindDurationPreview();
    this._bindCustomAlertDays();
    this._bindFormSubmit();
    this._bindActionButtons();
    this._renderRecentEmails();
  }

  // ── Public ────────────────────────────────────────────────

  refreshRecentEmails() { this._renderRecentEmails(); }

  /**
   * Set callback for when a contract is updated (edit mode)
   * @param {Function} callback - (contractId, updatedData) => void
   */
  setOnUpdated(callback) { this._onUpdated = callback; }

  /**
   * Enter edit mode: populate form with existing contract data
   * @param {Object} contract - existing contract to edit
   */
  enterEditMode(contract) {
    this._editMode = true;
    this._editContractId = contract.contract_id;

    // Update page heading
    const heading = document.getElementById('heading-add-contract');
    if (heading) heading.textContent = '✏️ แก้ไขสัญญา MA/PM';

    // Update submit button
    const submitBtn = this._form.querySelector('[type="submit"]');
    if (submitBtn) {
      submitBtn.innerHTML = `
        <svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2">
          <path d="M19 21H5a2 2 0 0 1-2-2V5a2 2 0 0 1 2-2h11l5 5v11a2 2 0 0 1-2 2z"/>
          <polyline points="17 21 17 13 7 13 7 21"/><polyline points="7 3 7 8 15 8"/>
        </svg>
        บันทึกการแก้ไข
      `;
    }

    // Show edit info banner
    this._showEditBanner(contract);

    // Populate form fields
    document.getElementById('poNumber').value = contract.po_number || '';
    document.getElementById('projectName').value = contract.project_name || '';
    document.getElementById('customerName').value = contract.customer_name || '';
    document.getElementById('serviceType').value = contract.service_type || '';
    document.getElementById('startDate').value = contract.start_date || '';
    document.getElementById('endDate').value = contract.end_date || '';
    document.getElementById('teamsWebhook').value = contract.teams_webhook || '';
    document.getElementById('lineGroupId').value = contract.line_group_id || '';
    document.getElementById('contractNote').value = contract.note || '';

    // Populate email tags
    this._saleEmailTags.clear();
    this._engEmailTags.clear();
    if (contract.recipients_sale) {
      contract.recipients_sale.split(',').filter(e => e.trim()).forEach(e => this._saleEmailTags.addEmail(e.trim()));
    }
    if (contract.recipients_eng) {
      contract.recipients_eng.split(',').filter(e => e.trim()).forEach(e => this._engEmailTags.addEmail(e.trim()));
    }

    // Trigger duration preview
    const startInput = document.getElementById('startDate');
    const endInput = document.getElementById('endDate');
    if (startInput && endInput) {
      startInput.dispatchEvent(new Event('change'));
    }

    this._renderRecentEmails();
  }

  /**
   * Exit edit mode and reset form to "add new" state
   */
  exitEditMode() {
    this._editMode = false;
    this._editContractId = null;

    // Restore heading
    const heading = document.getElementById('heading-add-contract');
    if (heading) heading.textContent = '🔔 เพิ่มสัญญา MA/PM ใหม่';

    // Restore submit button
    const submitBtn = this._form.querySelector('[type="submit"]');
    if (submitBtn) {
      submitBtn.innerHTML = `
        <svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2">
          <path d="M19 21H5a2 2 0 0 1-2-2V5a2 2 0 0 1 2-2h11l5 5v11a2 2 0 0 1-2 2z"/>
          <polyline points="17 21 17 13 7 13 7 21"/><polyline points="7 3 7 8 15 8"/>
        </svg>
        บันทึกสัญญา
      `;
    }

    // Remove edit banner
    const banner = document.getElementById('editBanner');
    if (banner) banner.remove();

    this._resetForm();
  }

  // ── Private: Edit Banner ─────────────────────────────────

  _showEditBanner(contract) {
    // Remove existing banner
    const existing = document.getElementById('editBanner');
    if (existing) existing.remove();

    const banner = document.createElement('div');
    banner.id = 'editBanner';
    banner.className = 'edit-banner';
    banner.innerHTML = `
      <div class="edit-banner-content">
        <svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2">
          <path d="M11 4H4a2 2 0 0 0-2 2v14a2 2 0 0 0 2 2h14a2 2 0 0 0 2-2v-7"/>
          <path d="M18.5 2.5a2.121 2.121 0 0 1 3 3L12 15l-4 1 1-4 9.5-9.5z"/>
        </svg>
        <span>กำลังแก้ไข: <strong>${contract.po_number}</strong> — ${contract.project_name}</span>
        <code style="font-size:0.7rem;opacity:0.7;margin-left:8px">${contract.contract_id}</code>
      </div>
      <button type="button" class="btn btn-sm btn-ghost edit-banner-cancel" id="btnCancelEdit">
        ✕ ยกเลิกการแก้ไข
      </button>
    `;

    // Insert before form
    this._form.parentNode.insertBefore(banner, this._form);

    document.getElementById('btnCancelEdit')?.addEventListener('click', () => {
      this.exitEditMode();
      this._navigate('contracts');
    });
  }

  // ── Private: Duration Preview ─────────────────────────────

  _bindDurationPreview() {
    const start = document.getElementById('startDate');
    const end = document.getElementById('endDate');
    const update = () => {
      if (start.value && end.value) {
        const diff = Math.ceil((new Date(end.value) - new Date(start.value)) / (1000 * 60 * 60 * 24));
        document.getElementById('durationText').textContent =
          diff === 0 ? 'ภายในวันเดียวกัน'
          : diff > 0 ? `${diff} วัน (${Math.round(diff / 30)} เดือน)`
          : 'วันที่ไม่ถูกต้อง';
      }
    };
    start.addEventListener('change', update);
    end.addEventListener('change', update);
  }

  // ── Private: Custom Alert Days ────────────────────────────

  _bindCustomAlertDays() {
    const addBtn = document.getElementById('addCustomAlert');
    const input  = document.getElementById('customAlertDays');

    addBtn?.addEventListener('click', () => this._addCustomAlertDay());
    input?.addEventListener('keydown', e => {
      if (e.key === 'Enter') { e.preventDefault(); this._addCustomAlertDay(); }
    });
  }

  _addCustomAlertDay() {
    const input = document.getElementById('customAlertDays');
    const val = parseInt(input.value);
    if (isNaN(val) || val < 0 || val > 365) {
      this._toast.error('❌ กรุณาระบุจำนวนวันระหว่าง 0–365');
      return;
    }
    const presetDays = [90, 60, 30, 7];
    if (presetDays.includes(val) || this._customAlertDays.includes(val)) {
      this._toast.error(`⚠️ ${val} วัน มีอยู่แล้ว`);
      return;
    }
    this._customAlertDays.push(val);
    this._customAlertDays.sort((a, b) => b - a);
    this._renderCustomAlertTags();
    input.value = '';
    this._toast.success(`✅ เพิ่มแจ้งเตือนก่อน ${val} วัน`);
  }

  _renderCustomAlertTags() {
    const container = document.getElementById('customAlertTags');
    container.innerHTML = this._customAlertDays.map((days, i) => `
      <span class="custom-alert-tag">
        ${days} วัน
        <button type="button" class="tag-remove" data-index="${i}" title="ลบ">×</button>
      </span>
    `).join('');

    container.querySelectorAll('.tag-remove').forEach(btn => {
      btn.addEventListener('click', () => {
        this._customAlertDays.splice(parseInt(btn.dataset.index), 1);
        this._renderCustomAlertTags();
      });
    });
  }

  // ── Private: Recent Emails ────────────────────────────────

  _renderRecentEmails() {
    const chipsContainer = document.getElementById('recentEmailsChips');
    const allEmails = this._history.collectAll(this._getContracts());
    const allUsed = new Set([
      ...this._saleEmailTags.getEmails(),
      ...this._engEmailTags.getEmails(),
    ]);

    if (allEmails.length === 0) {
      chipsContainer.innerHTML = '<span style="font-size:0.78rem;color:var(--text-muted)">ยังไม่มี email ที่เคยใช้</span>';
      return;
    }

    chipsContainer.innerHTML = allEmails.slice(0, 15).map(email => {
      const isUsed = allUsed.has(email);
      return `
        <button class="recent-chip ${isUsed ? 'used' : ''}" data-email="${email}" type="button" title="${isUsed ? 'เพิ่มแล้ว' : 'คลิกเพื่อเพิ่ม'}">
          <span class="chip-avatar">${getInitials(email)}</span>
          <span>${email}</span>
          ${isUsed ? '' : '<span class="chip-add">+</span>'}
        </button>
      `;
    }).join('');

    chipsContainer.querySelectorAll('.recent-chip:not(.used)').forEach(chip => {
      chip.addEventListener('click', () => {
        this._showAddToFieldPopup(chip.dataset.email, chip);
      });
    });
  }

  _showAddToFieldPopup(email, anchorEl) {
    document.querySelectorAll('.field-picker-popup').forEach(p => p.remove());

    const popup = document.createElement('div');
    popup.className = 'field-picker-popup';
    popup.innerHTML = `
      <div style="font-size:0.72rem;color:var(--text-muted);margin-bottom:6px;font-weight:600">เพิ่ม "${email}" ให้:</div>
      <button class="field-picker-btn" data-target="sale" type="button">📧 Email Sale</button>
      <button class="field-picker-btn" data-target="eng" type="button">🔧 Email Engineer</button>
      <button class="field-picker-btn" data-target="both" type="button">✅ ทั้ง Sale + Engineer</button>
    `;

    popup.style.cssText = `
      position:fixed;background:var(--bg-secondary);border:1px solid var(--border-color);
      border-radius:var(--radius-md);padding:10px;box-shadow:0 8px 32px rgba(0,0,0,0.5);
      z-index:100;min-width:200px;animation:dropIn 0.2s ease;
    `;

    const rect = anchorEl.getBoundingClientRect();
    popup.style.left = rect.left + 'px';
    popup.style.top  = (rect.bottom + 6) + 'px';
    document.body.appendChild(popup);

    popup.querySelectorAll('.field-picker-btn').forEach(btn => {
      btn.style.cssText = `
        display:block;width:100%;text-align:left;padding:8px 10px;background:transparent;
        border:1px solid var(--border-subtle);border-radius:6px;color:var(--text-primary);
        font-family:inherit;font-size:0.8rem;cursor:pointer;margin-bottom:4px;transition:all 0.15s ease;
      `;
      btn.addEventListener('mouseenter', () => { btn.style.background = 'rgba(59,130,246,0.1)'; btn.style.borderColor = 'var(--primary-500)'; });
      btn.addEventListener('mouseleave', () => { btn.style.background = 'transparent'; btn.style.borderColor = 'var(--border-subtle)'; });
      btn.addEventListener('click', () => {
        const target = btn.dataset.target;
        if (target === 'sale' || target === 'both') this._saleEmailTags.addEmail(email);
        if (target === 'eng'  || target === 'both') this._engEmailTags.addEmail(email);
        popup.remove();
        this._renderRecentEmails();
      });
    });

    setTimeout(() => {
      const closeHandler = e => {
        if (!popup.contains(e.target)) { popup.remove(); document.removeEventListener('click', closeHandler); }
      };
      document.addEventListener('click', closeHandler);
    }, 10);
  }

  // ── Private: Form Submit ──────────────────────────────────

  _bindFormSubmit() {
    this._form.addEventListener('submit', async e => {
      e.preventDefault();
      if (this._editMode) {
        await this._handleUpdate();
      } else {
        await this._handleSubmit();
      }
    });
  }

  // ── Handle UPDATE (edit mode) ─────────────────────────────

  async _handleUpdate() {
    const saleEmails = this._saleEmailTags.getEmails();
    const engEmails  = this._engEmailTags.getEmails();

    // Collect alert days
    const alertDays = [];
    ['alert90', 'alert60', 'alert30', 'alert7', 'alert0'].forEach(id => {
      const el = document.getElementById(id);
      if (el?.checked) alertDays.push(parseInt(el.value));
    });
    this._customAlertDays.forEach(d => { if (!alertDays.includes(d)) alertDays.push(d); });
    alertDays.sort((a, b) => b - a);

    const notifyTime = document.getElementById('notifyTime').value || '08:00';

    const updatedData = {
      po_number:       document.getElementById('poNumber').value.trim(),
      project_name:    document.getElementById('projectName').value.trim(),
      customer_name:   document.getElementById('customerName').value.trim(),
      service_type:    document.getElementById('serviceType').value,
      start_date:      document.getElementById('startDate').value,
      end_date:        document.getElementById('endDate').value,
      recipients_sale: saleEmails.join(','),
      recipients_eng:  engEmails.join(','),
      teams_webhook:   document.getElementById('teamsWebhook').value.trim(),
      note:            document.getElementById('contractNote').value.trim(),
      line_group_id:   document.getElementById('lineGroupId').value.trim(),
      alert_days:      alertDays,
      notify_time:     notifyTime,
      regenerate_rules: true, // ให้ backend สร้าง rules ใหม่
    };

    const submitBtn = this._form.querySelector('[type="submit"]');
    const originalHTML = submitBtn?.innerHTML;
    if (submitBtn) { submitBtn.disabled = true; submitBtn.textContent = '⏳ กำลังบันทึก...'; }

    try {
      if (this._service.isDemoMode()) {
        // Demo mode — update local only
        this._toast.warning('⚠️ อัพเดทเฉพาะ local session (demo mode)');
        if (this._onUpdated) {
          this._onUpdated(this._editContractId, {
            ...updatedData,
            contract_id: this._editContractId,
            status: 'active',
            notify_line: !!updatedData.line_group_id,
          });
        }
        this.exitEditMode();
        this._navigate('contracts');
        return;
      }

      const result = await this._service.updateContract(this._editContractId, updatedData);

      if (result.status === 'ok') {
        this._toast.success(`✅ อัพเดทสัญญา ${updatedData.po_number} เรียบร้อย`);
        if (this._onUpdated) {
          this._onUpdated(this._editContractId, {
            ...updatedData,
            contract_id: this._editContractId,
            status: 'active',
            notify_line: !!updatedData.line_group_id,
          });
        }
        this.exitEditMode();
        this._navigate('contracts');
      } else {
        this._toast.error(`❌ อัพเดทไม่สำเร็จ: ${result.message || 'ไม่ทราบสาเหตุ'}`);
      }
    } catch (err) {
      console.error('Update error:', err);
      this._toast.error(`❌ เชื่อมต่อไม่ได้: ${err.message}`);
    } finally {
      if (submitBtn) { submitBtn.disabled = false; submitBtn.innerHTML = originalHTML; }
    }
  }

  // ── Handle CREATE (normal mode) ───────────────────────────

  async _handleSubmit() {
    const saleEmails = this._saleEmailTags.getEmails();
    const engEmails  = this._engEmailTags.getEmails();

    // Collect alert days
    const alertDays = [];
    ['alert90', 'alert60', 'alert30', 'alert7', 'alert0'].forEach(id => {
      const el = document.getElementById(id);
      if (el?.checked) alertDays.push(parseInt(el.value));
    });
    this._customAlertDays.forEach(d => { if (!alertDays.includes(d)) alertDays.push(d); });
    alertDays.sort((a, b) => b - a);

    const notifyTime = document.getElementById('notifyTime').value || '08:00';

    const newContract = {
      po_number:       document.getElementById('poNumber').value.trim(),
      project_name:    document.getElementById('projectName').value.trim(),
      customer_name:   document.getElementById('customerName').value.trim(),
      service_type:    document.getElementById('serviceType').value,
      start_date:      document.getElementById('startDate').value,
      end_date:        document.getElementById('endDate').value,
      recipients_sale: saleEmails.join(','),
      recipients_eng:  engEmails.join(','),
      teams_webhook:   document.getElementById('teamsWebhook').value.trim(),
      note:            document.getElementById('contractNote').value.trim(),
      line_group_id:   document.getElementById('lineGroupId').value.trim(),
      alert_days:      alertDays,
      notify_time:     notifyTime,
      created_by:      'dashboard',
    };

    const submitBtn = this._form.querySelector('[type="submit"]');
    const originalText = submitBtn?.textContent;
    if (submitBtn) { submitBtn.disabled = true; submitBtn.textContent = '⏳ กำลังบันทึก...'; }

    try {
      // Demo mode — no GAS URL
      if (this._service.isDemoMode()) {
        this._toast.warning('⚠️ ยังไม่ได้ตั้งค่า GAS URL — บันทึกเฉพาะ local session (demo mode)');
        this._onAdded({ ...newContract, contract_id: generateId('WC'), status: 'active', notify_line: !!newContract.line_group_id });
        this._resetForm();
        this._navigate('contracts');
        return;
      }

      const result = await this._service.addContract(newContract);

      if (result.status === 'ok') {
        const channelSummary = [];
        if (saleEmails.length > 0 || engEmails.length > 0) channelSummary.push('📧 Email');
        if (newContract.teams_webhook) channelSummary.push('💬 Teams');
        if (newContract.line_group_id) channelSummary.push('💚 LINE');
        const channelText = channelSummary.length > 0 ? channelSummary.join(', ') : 'ไม่มีช่องทาง';

        this._toast.success(
          `✅ บันทึกสัญญา ${newContract.po_number} ลง Google Sheets เรียบร้อย — ` +
          `แจ้งเตือน ${alertDays.length > 0 ? alertDays.join(', ') + ' วัน' : 'ค่าเริ่มต้น'} ` +
          `⏰ ${notifyTime} น. (${channelText})`
        );

        this._onAdded({
          ...newContract,
          contract_id: result.contract_id || generateId('WC'),
          status: 'active',
          notify_line: !!newContract.line_group_id,
        });
        this._resetForm();
        this._navigate('contracts');
      } else {
        this._toast.error(`❌ เกิดข้อผิดพลาด: ${result.message || 'ไม่ทราบสาเหตุ'}`);
      }
    } catch (err) {
      console.error('GAS fetch error:', err);
      this._toast.error(`❌ เชื่อมต่อ Google Apps Script ไม่ได้: ${err.message}`);
    } finally {
      if (submitBtn) { submitBtn.disabled = false; submitBtn.textContent = originalText; }
    }
  }

  _resetForm() {
    this._form.reset();
    this._saleEmailTags.clear();
    this._engEmailTags.clear();
    this._customAlertDays.length = 0;
    this._renderCustomAlertTags();
    this._renderRecentEmails();
  }

  _bindActionButtons() {
    document.getElementById('btnCancel')?.addEventListener('click', () => {
      if (this._editMode) {
        this.exitEditMode();
      }
      this._navigate('dashboard');
    });
  }
}
