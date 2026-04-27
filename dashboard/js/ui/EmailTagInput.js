// ============================================================
//  EmailTagInput — SRP: จัดการ tag input สำหรับ email เดียว
//  GRASP: Creator — สร้างโดย AddContractUI (ผู้มีข้อมูลที่จำเป็น)
//  ISP: ไม่รู้จัก ContractService หรือ history storage
// ============================================================

class EmailTagInput {
  /**
   * @param {string} containerId
   * @param {string} tagsId
   * @param {string} inputId
   * @param {string} dropdownId
   * @param {EmailHistoryService} historyService
   * @param {Function} getEmailSuggestions - () => string[]
   * @param {Function} onChanged - callback เมื่อ email list เปลี่ยน
   */
  constructor(containerId, tagsId, inputId, dropdownId, historyService, getEmailSuggestions, onChanged) {
    this.container = document.getElementById(containerId);
    this.tagsEl    = document.getElementById(tagsId);
    this.input     = document.getElementById(inputId);
    this.dropdown  = document.getElementById(dropdownId);
    this._history  = historyService;
    this._getSuggestions = getEmailSuggestions;
    this._onChanged = onChanged || (() => {});
    this.emails    = [];
    this.highlightIndex = -1;
    this._bindEvents();
  }

  // ── Public ────────────────────────────────────────────────

  addEmail(email) {
    const lower = email.toLowerCase().trim();
    if (!lower || this.emails.includes(lower)) return;
    this.emails.push(lower);
    this.input.value = '';
    this._renderTags();
    this._hideDropdown();
    this._history.save(lower);
    this._onChanged();
  }

  removeEmail(email) {
    this.emails = this.emails.filter(e => e !== email);
    this._renderTags();
    this._onChanged();
  }

  getEmails() { return [...this.emails]; }

  clear() {
    this.emails = [];
    this.input.value = '';
    this._renderTags();
  }

  // ── Private ───────────────────────────────────────────────

  _bindEvents() {
    this.container.addEventListener('click', () => this.input.focus());

    this.input.addEventListener('keydown', e => {
      if (e.key === 'Enter' || e.key === 'Tab' || e.key === ',') {
        e.preventDefault();
        if (this.highlightIndex >= 0) {
          const items = this.dropdown.querySelectorAll('.email-dropdown-item');
          if (items[this.highlightIndex]) {
            this.addEmail(items[this.highlightIndex].dataset.email);
            this._hideDropdown();
            return;
          }
        }
        this._commitInput();
      } else if (e.key === 'Backspace' && !this.input.value) {
        if (this.emails.length > 0) this.removeEmail(this.emails[this.emails.length - 1]);
      } else if (e.key === 'ArrowDown') {
        e.preventDefault(); this._navigateDropdown(1);
      } else if (e.key === 'ArrowUp') {
        e.preventDefault(); this._navigateDropdown(-1);
      } else if (e.key === 'Escape') {
        this._hideDropdown();
      }
    });

    this.input.addEventListener('input', () => this._showAutocomplete());
    this.input.addEventListener('focus', () => {
      if (this.input.value.trim()) this._showAutocomplete();
    });
    this.input.addEventListener('blur', () => {
      setTimeout(() => this._hideDropdown(), 200);
    });
    this.input.addEventListener('paste', e => {
      e.preventDefault();
      const text = (e.clipboardData || window.clipboardData).getData('text');
      text.split(/[,;\s\n]+/).filter(s => s.trim()).forEach(email => {
        const trimmed = email.trim();
        if (isValidEmail(trimmed)) this.addEmail(trimmed);
      });
    });
  }

  _commitInput() {
    const val = this.input.value.replace(/,/g, '').trim();
    if (val && isValidEmail(val)) {
      this.addEmail(val);
    } else if (val) {
      this.container.style.borderColor = 'var(--red-500)';
      setTimeout(() => { this.container.style.borderColor = ''; }, 800);
    }
  }

  _renderTags() {
    this.tagsEl.innerHTML = this.emails.map(email => `
      <span class="email-tag" data-email="${email}">
        <span class="tag-text">${email}</span>
        <button class="tag-remove" type="button" title="ลบ ${email}">×</button>
      </span>
    `).join('');

    this.tagsEl.querySelectorAll('.tag-remove').forEach(btn => {
      btn.addEventListener('click', e => {
        e.stopPropagation();
        this.removeEmail(btn.closest('.email-tag').dataset.email);
      });
    });

    this.input.placeholder = this.emails.length > 0 ? 'เพิ่มอีก...' : 'พิมพ์ email แล้วกด Enter';
  }

  _showAutocomplete() {
    const query = this.input.value.trim().toLowerCase();
    if (!query) { this._hideDropdown(); return; }

    const matches = this._getSuggestions()
      .filter(e => e.includes(query) && !this.emails.includes(e))
      .slice(0, 6);

    if (matches.length === 0) { this._hideDropdown(); return; }

    this.highlightIndex = -1;
    this.dropdown.innerHTML = matches.map(email => `
      <div class="email-dropdown-item" data-email="${email}">
        <div class="dropdown-icon">${getInitials(email)}</div>
        <span class="dropdown-email">${email}</span>
        <span class="dropdown-hint">เพิ่ม</span>
      </div>
    `).join('');

    this.dropdown.querySelectorAll('.email-dropdown-item').forEach(item => {
      item.addEventListener('mousedown', e => {
        e.preventDefault();
        this.addEmail(item.dataset.email);
      });
    });

    this.dropdown.classList.add('show');
  }

  _hideDropdown() {
    this.dropdown.classList.remove('show');
    this.highlightIndex = -1;
  }

  _navigateDropdown(direction) {
    const items = this.dropdown.querySelectorAll('.email-dropdown-item');
    if (items.length === 0) return;
    items.forEach(i => i.classList.remove('highlighted'));
    this.highlightIndex += direction;
    if (this.highlightIndex < 0) this.highlightIndex = items.length - 1;
    if (this.highlightIndex >= items.length) this.highlightIndex = 0;
    items[this.highlightIndex].classList.add('highlighted');
    items[this.highlightIndex].scrollIntoView({ block: 'nearest' });
  }
}

// ── Email helpers (shared, used by EmailTagInput) ───────────
function isValidEmail(email) {
  return /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email);
}

function getInitials(email) {
  const name = email.split('@')[0];
  return name.length <= 2 ? name.toUpperCase() : name.substring(0, 2).toUpperCase();
}
