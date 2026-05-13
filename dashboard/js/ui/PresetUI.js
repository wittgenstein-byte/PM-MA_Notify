// ============================================================
//  PresetUI — จัดการเกี่ยวกับ Preset (Save, Load, Manage)
// ============================================================

class PresetUI {
  constructor(contractService, addContractUI, toast) {
    this._service = contractService;
    this._addContractUI = addContractUI;
    this._toast = toast;
    this._presets = [];

    // DOM Elements
    this._selector = document.getElementById('presetSelector');
    this._btnManage = document.getElementById('btnManagePresets');
    this._btnSave = document.getElementById('btnSavePreset');
    this._modal = document.getElementById('presetModalOverlay');
    this._modalClose = document.getElementById('presetModalClose');
    this._tableBody = document.getElementById('presetsTableBody');

    this._bindEvents();
    this.loadPresets();
  }

  async loadPresets() {
    try {
      const result = await this._service.fetchPresets();
      if (result.status === 'ok') {
        this._presets = result.data || [];
        this._renderSelector();
        this._renderTable();
      }
    } catch (err) {
      console.error('Failed to load presets:', err);
    }
  }

  _renderSelector() {
    if (!this._selector) return;
    this._selector.innerHTML = '<option value="">-- เลือก Preset --</option>';
    this._presets.forEach(p => {
      const opt = document.createElement('option');
      opt.value = p.preset_id;
      opt.textContent = `${p.preset_name} (${p.customer_name || 'ไม่ระบุลูกค้า'})`;
      this._selector.appendChild(opt);
    });
  }

  _renderTable() {
    if (!this._tableBody) return;
    if (this._presets.length === 0) {
      this._tableBody.innerHTML = '<tr><td colspan="4" style="text-align:center;color:var(--text-muted)">ไม่มี Preset</td></tr>';
      return;
    }

    this._tableBody.innerHTML = this._presets.map(p => `
      <tr>
        <td><strong>${p.preset_name}</strong></td>
        <td>${p.customer_name || '-'}</td>
        <td><span class="badge ${p.service_type === 'PM' ? 'badge-primary' : (p.service_type === 'MA' ? 'badge-secondary' : 'badge-warning')}">${p.service_type || '-'}</span></td>
        <td>
          <button class="btn btn-outline btn-sm btn-delete-preset" data-id="${p.preset_id}" title="ลบ Preset">
            <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><polyline points="3 6 5 6 21 6"/><path d="M19 6v14a2 2 0 0 1-2 2H7a2 2 0 0 1-2-2V6m3 0V4a2 2 0 0 1 2-2h4a2 2 0 0 1 2 2v2"/></svg>
          </button>
        </td>
      </tr>
    `).join('');

    this._tableBody.querySelectorAll('.btn-delete-preset').forEach(btn => {
      btn.addEventListener('click', async (e) => {
        const id = e.currentTarget.dataset.id;
        if (confirm('คุณแน่ใจหรือไม่ว่าต้องการลบ Preset นี้?')) {
          await this._deletePreset(id);
        }
      });
    });
  }

  _bindEvents() {
    // โหลด Preset ลงฟอร์ม
    this._selector?.addEventListener('change', (e) => {
      const presetId = e.target.value;
      if (!presetId) return;
      
      const preset = this._presets.find(p => p.preset_id === presetId);
      if (preset) {
        this._addContractUI.fillFormWithPreset(preset);
        this._toast.success(`✅ โหลด Preset: ${preset.preset_name}`);
      }
    });

    // เปิด Modal จัดการ Preset
    this._btnManage?.addEventListener('click', () => {
      this._modal.classList.add('active');
    });

    // ปิด Modal
    this._modalClose?.addEventListener('click', () => {
      this._modal.classList.remove('active');
    });
    this._modal?.addEventListener('click', (e) => {
      if (e.target === this._modal) this._modal.classList.remove('active');
    });

    // บันทึกฟอร์มปัจจุบันเป็น Preset
    this._btnSave?.addEventListener('click', async () => {
      const presetName = prompt('กรุณาตั้งชื่อให้ Preset นี้:');
      if (!presetName || !presetName.trim()) return;

      const formData = this._addContractUI.getFormDataForPreset();
      formData.preset_name = presetName.trim();

      const originalText = this._btnSave.innerHTML;
      this._btnSave.disabled = true;
      this._btnSave.textContent = 'กำลังบันทึก...';

      try {
        const result = await this._service.addPreset(formData);
        if (result.status === 'ok') {
          this._toast.success(`✅ บันทึก Preset "${presetName}" เรียบร้อย`);
          await this.loadPresets();
        } else {
          this._toast.error(`❌ บันทึกไม่สำเร็จ: ${result.message}`);
        }
      } catch (err) {
        this._toast.error('❌ เชื่อมต่อระบบไม่สำเร็จ');
      } finally {
        this._btnSave.disabled = false;
        this._btnSave.innerHTML = originalText;
      }
    });
  }

  async _deletePreset(id) {
    try {
      const result = await this._service.deletePreset(id);
      if (result.status === 'ok') {
        this._toast.success('🗑️ ลบ Preset เรียบร้อย');
        this._presets = this._presets.filter(p => p.preset_id !== id);
        this._renderSelector();
        this._renderTable();
      } else {
        this._toast.error(`❌ ลบไม่สำเร็จ: ${result.message}`);
      }
    } catch (err) {
      this._toast.error('❌ เชื่อมต่อระบบไม่สำเร็จ');
    }
  }
}
