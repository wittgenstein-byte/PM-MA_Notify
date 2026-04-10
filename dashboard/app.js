// ============================================================
//  PM-MA Warranty Notification System — Dashboard App
// ============================================================

(function () {
  'use strict';

  // ── Sample Data ──────────────────────────────────────────────

  const today = new Date();
  today.setHours(0, 0, 0, 0);

  const sampleContracts = [
    {
      contract_id: 'WC-2026-001',
      po_number: 'PO-2026-0356',
      project_name: 'ระบบ CCTV โรงพยาบาลขอนแก่น',
      customer_name: 'โรงพยาบาลขอนแก่น',
      service_type: 'MA+PM',
      start_date: '2025-01-01',
      end_date: '2026-04-25',
      recipients_sale: 'sale@company.com',
      recipients_eng: 'eng1@company.com,eng2@company.com',
      teams_webhook: 'https://outlook.office.com/webhook/sample1',
      note: 'ดูแลกล้อง CCTV 120 ตัว',
      status: 'active',
    },
    {
      contract_id: 'WC-2026-002',
      po_number: 'PO-2026-0289',
      project_name: 'ระบบ Network มหาวิทยาลัยขอนแก่น',
      customer_name: 'มหาวิทยาลัยขอนแก่น',
      service_type: 'MA',
      start_date: '2025-06-01',
      end_date: '2026-05-15',
      recipients_sale: 'sale2@company.com',
      recipients_eng: 'eng3@company.com',
      teams_webhook: 'https://outlook.office.com/webhook/sample2',
      note: '',
      status: 'active',
    },
    {
      contract_id: 'WC-2026-003',
      po_number: 'PO-2026-0100',
      project_name: 'ระบบ Firewall การไฟฟ้าส่วนภูมิภาค',
      customer_name: 'การไฟฟ้าส่วนภูมิภาค (กฟภ.)',
      service_type: 'MA+PM',
      start_date: '2026-01-01',
      end_date: '2026-12-31',
      recipients_sale: 'sale@company.com',
      recipients_eng: 'eng1@company.com',
      teams_webhook: 'https://outlook.office.com/webhook/sample3',
      note: '',
      status: 'active',
    },
    {
      contract_id: 'WC-2026-004',
      po_number: 'PO-2025-0899',
      project_name: 'ระบบ Server ธนาคารกรุงไทย สาขาขอนแก่น',
      customer_name: 'ธนาคารกรุงไทย',
      service_type: 'PM',
      start_date: '2025-01-01',
      end_date: '2026-04-12',
      recipients_sale: 'sale3@company.com',
      recipients_eng: 'eng2@company.com',
      teams_webhook: '',
      note: 'PM ทุก 3 เดือน',
      status: 'active',
    },
    {
      contract_id: 'WC-2026-005',
      po_number: 'PO-2026-0501',
      project_name: 'ระบบ Wi-Fi โรงเรียนสาธิต มข.',
      customer_name: 'โรงเรียนสาธิต มข.',
      service_type: 'MA',
      start_date: '2026-01-01',
      end_date: '2026-06-30',
      recipients_sale: 'sale4@company.com',
      recipients_eng: 'eng4@company.com',
      teams_webhook: 'https://outlook.office.com/webhook/sample5',
      note: 'ดูแล AP 50 ตัว',
      status: 'active',
    },
    {
      contract_id: 'WC-2026-006',
      po_number: 'PO-2026-0620',
      project_name: 'ระบบ UPS ศูนย์ข้อมูล กระทรวงสาธารณสุข',
      customer_name: 'กระทรวงสาธารณสุข',
      service_type: 'PM',
      start_date: '2026-03-01',
      end_date: '2027-02-28',
      recipients_sale: 'sale@company.com',
      recipients_eng: 'eng1@company.com',
      teams_webhook: 'https://outlook.office.com/webhook/sample6',
      note: '',
      status: 'active',
    },
    {
      contract_id: 'WC-2026-007',
      po_number: 'PO-2026-0710',
      project_name: 'ระบบ Access Control โรงงาน CPF ขอนแก่น',
      customer_name: 'CPF ขอนแก่น',
      service_type: 'MA+PM',
      start_date: '2026-02-01',
      end_date: '2027-01-31',
      recipients_sale: 'sale5@company.com',
      recipients_eng: 'eng5@company.com',
      teams_webhook: '',
      note: '',
      status: 'active',
    },
    {
      contract_id: 'WC-2025-008',
      po_number: 'PO-2025-0200',
      project_name: 'ระบบ PA โรงพยาบาลศรีนครินทร์',
      customer_name: 'โรงพยาบาลศรีนครินทร์',
      service_type: 'MA',
      start_date: '2024-06-01',
      end_date: '2025-05-31',
      recipients_sale: 'sale@company.com',
      recipients_eng: 'eng1@company.com',
      teams_webhook: '',
      note: '',
      status: 'expired',
    },
    {
      contract_id: 'WC-2026-009',
      po_number: 'PO-2026-0888',
      project_name: 'ระบบ Digital Signage ห้าง Central ขอนแก่น',
      customer_name: 'Central ขอนแก่น',
      service_type: 'MA',
      start_date: '2026-01-15',
      end_date: '2027-01-14',
      recipients_sale: 'sale6@company.com',
      recipients_eng: 'eng6@company.com,eng7@company.com',
      teams_webhook: 'https://outlook.office.com/webhook/sample9',
      note: '',
      status: 'active',
    },
    {
      contract_id: 'WC-2026-010',
      po_number: 'PO-2026-0999',
      project_name: 'ระบบ VoIP สำนักงานจังหวัดขอนแก่น',
      customer_name: 'สำนักงานจังหวัดขอนแก่น',
      service_type: 'MA+PM',
      start_date: '2026-04-01',
      end_date: '2027-03-31',
      recipients_sale: 'sale7@company.com',
      recipients_eng: 'eng8@company.com',
      teams_webhook: 'https://outlook.office.com/webhook/sample10',
      note: '',
      status: 'active',
    },
  ];

  const sampleLogs = [
    { log_id: 'LOG-001', rule_id: 'NR-001', contract_id: 'WC-2026-001', channel: 'email', status: 'success', error_msg: '', sent_at: '2026-01-25 08:00' },
    { log_id: 'LOG-002', rule_id: 'NR-001', contract_id: 'WC-2026-001', channel: 'teams', status: 'success', error_msg: '', sent_at: '2026-01-25 08:00' },
    { log_id: 'LOG-003', rule_id: 'NR-002', contract_id: 'WC-2026-002', channel: 'email', status: 'success', error_msg: '', sent_at: '2026-02-14 08:00' },
    { log_id: 'LOG-004', rule_id: 'NR-002', contract_id: 'WC-2026-002', channel: 'teams', status: 'failed', error_msg: 'Webhook URL invalid', sent_at: '2026-02-14 08:01' },
    { log_id: 'LOG-005', rule_id: 'NR-003', contract_id: 'WC-2026-004', channel: 'email', status: 'success', error_msg: '', sent_at: '2026-03-13 08:00' },
    { log_id: 'LOG-006', rule_id: 'NR-004', contract_id: 'WC-2026-001', channel: 'email', status: 'success', error_msg: '', sent_at: '2026-02-24 08:00' },
    { log_id: 'LOG-007', rule_id: 'NR-004', contract_id: 'WC-2026-001', channel: 'teams', status: 'success', error_msg: '', sent_at: '2026-02-24 08:00' },
    { log_id: 'LOG-008', rule_id: 'NR-005', contract_id: 'WC-2026-003', channel: 'email', status: 'success', error_msg: '', sent_at: '2026-10-02 08:00' },
  ];

  // ── Utility Functions ────────────────────────────────────────

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

  // ── DOM Cache ────────────────────────────────────────────────

  const $ = id => document.getElementById(id);
  const pages = document.querySelectorAll('.page');
  const navItems = document.querySelectorAll('.nav-item');

  // ── Navigation ───────────────────────────────────────────────

  function navigateTo(pageId) {
    pages.forEach(p => p.classList.remove('active'));
    navItems.forEach(n => n.classList.remove('active'));

    const targetPage = $(`page-${pageId}`);
    const targetNav = document.querySelector(`[data-page="${pageId}"]`);

    if (targetPage) targetPage.classList.add('active');
    if (targetNav) targetNav.classList.add('active');

    // Close mobile sidebar
    $('sidebar').classList.remove('open');
  }

  navItems.forEach(item => {
    item.addEventListener('click', e => {
      e.preventDefault();
      const page = item.dataset.page;
      navigateTo(page);
    });
  });

  // Mobile menu toggle
  const menuToggle = $('menuToggle');
  if (menuToggle) {
    menuToggle.addEventListener('click', () => {
      $('sidebar').classList.toggle('open');
    });
  }

  // ── Date Display ─────────────────────────────────────────────

  $('currentDate').textContent = formatThaiDate(today);

  // ── Dashboard ────────────────────────────────────────────────

  function renderDashboard() {
    let critical = 0, warning = 0, safe = 0, total = 0;

    const enriched = sampleContracts.map(c => {
      const daysLeft = calcDaysLeft(c.end_date);
      const category = getStatusCategory(daysLeft, c.status);
      return { ...c, daysLeft, category };
    });

    enriched.forEach(c => {
      if (c.status === 'active') {
        total++;
        if (c.category === 'critical') critical++;
        else if (c.category === 'warning') warning++;
        else if (c.category === 'safe') safe++;
      }
    });

    // Add expired to total for overall count
    const expiredCount = sampleContracts.filter(c => c.status === 'expired' || c.status === 'cancel').length;
    total += expiredCount;

    // Animate stat counters
    animateCounter('criticalCount', critical);
    animateCounter('warningCount', warning);
    animateCounter('safeCount', safe);
    animateCounter('totalCount', total);

    // Dashboard table — sorted by daysLeft ascending
    const sorted = enriched
      .filter(c => c.status === 'active')
      .sort((a, b) => a.daysLeft - b.daysLeft)
      .slice(0, 8);

    const tbody = $('dashboardTableBody');
    tbody.innerHTML = sorted.map(c => `
      <tr data-id="${c.contract_id}">
        <td><code style="color:var(--primary-400)">${c.po_number}</code></td>
        <td>${c.project_name}</td>
        <td>${c.customer_name}</td>
        <td><span class="badge badge-${c.service_type === 'MA+PM' ? 'email' : 'teams'}">${c.service_type}</span></td>
        <td>${formatDate(c.end_date)}</td>
        <td><span class="days-remaining ${c.category}">${c.daysLeft} วัน</span></td>
        <td>${renderStatusBadge(c.category)}</td>
      </tr>
    `).join('');

    // Click row => open modal
    tbody.querySelectorAll('tr').forEach(row => {
      row.addEventListener('click', () => openContractModal(row.dataset.id));
    });

    // Notification feed
    renderNotificationFeed();
  }

  function animateCounter(elementId, target) {
    const el = $(elementId);
    let current = 0;
    const step = Math.max(1, Math.ceil(target / 20));
    const interval = setInterval(() => {
      current += step;
      if (current >= target) {
        current = target;
        clearInterval(interval);
      }
      el.textContent = current;
    }, 30);
  }

  function renderStatusBadge(category) {
    const labels = {
      critical: '🔴 วิกฤต',
      warning: '🟡 ปานกลาง',
      safe: '🟢 ปกติ',
      expired: '⚫ หมดอายุ'
    };
    return `<span class="badge badge-${category}">${labels[category] || category}</span>`;
  }

  function renderNotificationFeed() {
    const feed = $('notificationFeed');
    const recentLogs = sampleLogs.slice(-6).reverse();

    feed.innerHTML = recentLogs.map(log => {
      const contract = sampleContracts.find(c => c.contract_id === log.contract_id);
      const isEmail = log.channel === 'email';
      return `
        <div class="notif-item">
          <div class="notif-icon ${log.channel}">
            ${isEmail ? '📧' : '💬'}
          </div>
          <div class="notif-content">
            <div class="notif-title">${isEmail ? 'Outlook Email' : 'MS Teams'} — ${contract ? contract.project_name : log.contract_id}</div>
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

  // ── Contracts List ───────────────────────────────────────────

  let currentFilter = 'all';

  function renderContractsList(filter = 'all') {
    currentFilter = filter;
    const enriched = sampleContracts.map(c => {
      const daysLeft = calcDaysLeft(c.end_date);
      const category = getStatusCategory(daysLeft, c.status);
      return { ...c, daysLeft, category };
    });

    let filtered = enriched;
    if (filter !== 'all') {
      filtered = enriched.filter(c => c.category === filter);
    }

    // Apply search
    const searchTerm = ($('searchInput')?.value || '').toLowerCase();
    if (searchTerm) {
      filtered = filtered.filter(c =>
        c.po_number.toLowerCase().includes(searchTerm) ||
        c.project_name.toLowerCase().includes(searchTerm) ||
        c.customer_name.toLowerCase().includes(searchTerm) ||
        c.contract_id.toLowerCase().includes(searchTerm)
      );
    }

    const tbody = $('contractsTableBody');
    if (filtered.length === 0) {
      tbody.innerHTML = `
        <tr>
          <td colspan="10" style="text-align:center;padding:40px;color:var(--text-muted)">
            ไม่พบข้อมูลสัญญา
          </td>
        </tr>
      `;
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
        <td>${renderStatusBadge(c.category)}</td>
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

    // Click row => open modal
    tbody.querySelectorAll('tr').forEach(row => {
      row.addEventListener('click', e => {
        if (e.target.closest('.btn-icon')) return;
        openContractModal(row.dataset.id);
      });
    });

    tbody.querySelectorAll('.btn-icon').forEach(btn => {
      btn.addEventListener('click', () => openContractModal(btn.dataset.view));
    });
  }

  // Filter tabs
  document.querySelectorAll('.filter-tab').forEach(tab => {
    tab.addEventListener('click', () => {
      document.querySelectorAll('.filter-tab').forEach(t => t.classList.remove('active'));
      tab.classList.add('active');
      renderContractsList(tab.dataset.filter);
    });
  });

  // Search
  const searchInput = $('searchInput');
  if (searchInput) {
    searchInput.addEventListener('input', () => renderContractsList(currentFilter));
  }

  // ── Logs List ────────────────────────────────────────────────

  function renderLogs() {
    const tbody = $('logsTableBody');
    tbody.innerHTML = sampleLogs.reverse().map(log => `
      <tr>
        <td><code style="font-size:0.75rem;color:var(--text-muted)">${log.log_id}</code></td>
        <td><code style="font-size:0.75rem">${log.rule_id}</code></td>
        <td><code style="font-size:0.75rem">${log.contract_id}</code></td>
        <td><span class="badge badge-${log.channel}">${log.channel === 'email' ? '📧 Email' : '💬 Teams'}</span></td>
        <td><span class="badge badge-${log.status}">${log.status}</span></td>
        <td style="color:${log.error_msg ? 'var(--red-400)' : 'var(--text-muted)'}">${log.error_msg || '-'}</td>
        <td>${log.sent_at}</td>
      </tr>
    `).join('');
  }

  // ── Contract Modal ───────────────────────────────────────────

  function openContractModal(contractId) {
    const contract = sampleContracts.find(c => c.contract_id === contractId);
    if (!contract) return;

    const daysLeft = calcDaysLeft(contract.end_date);
    const category = getStatusCategory(daysLeft, contract.status);

    $('modalTitle').textContent = `📄 ${contract.po_number}`;
    $('modalBody').innerHTML = `
      <div class="detail-row">
        <span class="detail-label">Contract ID</span>
        <span class="detail-value">${contract.contract_id}</span>
      </div>
      <div class="detail-row">
        <span class="detail-label">PO Number</span>
        <span class="detail-value" style="color:var(--primary-400)">${contract.po_number}</span>
      </div>
      <div class="detail-row">
        <span class="detail-label">📁 โครงการ</span>
        <span class="detail-value">${contract.project_name}</span>
      </div>
      <div class="detail-row">
        <span class="detail-label">🏢 ลูกค้า</span>
        <span class="detail-value">${contract.customer_name}</span>
      </div>
      <div class="detail-row">
        <span class="detail-label">🛡️ ประเภท</span>
        <span class="detail-value">${contract.service_type}</span>
      </div>
      <div class="detail-row">
        <span class="detail-label">📅 เริ่มต้น</span>
        <span class="detail-value">${formatDate(contract.start_date)}</span>
      </div>
      <div class="detail-row">
        <span class="detail-label">📅 หมดอายุ</span>
        <span class="detail-value" style="color:var(--red-400);font-weight:700">${formatDate(contract.end_date)}</span>
      </div>
      <div class="detail-row">
        <span class="detail-label">⏳ เหลืออีก</span>
        <span class="detail-value">
          <span class="days-remaining ${category}" style="font-size:1.1rem">${daysLeft > 0 ? daysLeft + ' วัน' : 'หมดอายุแล้ว'}</span>
        </span>
      </div>
      <div class="detail-row">
        <span class="detail-label">สถานะ</span>
        <span class="detail-value">${renderStatusBadge(category)}</span>
      </div>
      <div class="detail-row">
        <span class="detail-label">📧 Email Sale</span>
        <span class="detail-value">${contract.recipients_sale}</span>
      </div>
      <div class="detail-row">
        <span class="detail-label">📧 Email Engineer</span>
        <span class="detail-value">${contract.recipients_eng}</span>
      </div>
      <div class="detail-row">
        <span class="detail-label">💬 Teams Webhook</span>
        <span class="detail-value" style="font-size:0.75rem;max-width:300px;word-break:break-all">${contract.teams_webhook || '-'}</span>
      </div>
      ${contract.note ? `
      <div class="detail-row">
        <span class="detail-label">📝 หมายเหตุ</span>
        <span class="detail-value">${contract.note}</span>
      </div>` : ''}
    `;

    $('modalOverlay').classList.add('active');
  }

  function closeModal() {
    $('modalOverlay').classList.remove('active');
  }

  $('modalClose').addEventListener('click', closeModal);
  $('modalOverlay').addEventListener('click', e => {
    if (e.target === $('modalOverlay')) closeModal();
  });
  document.addEventListener('keydown', e => {
    if (e.key === 'Escape') closeModal();
  });

  // ── Add Contract Form ────────────────────────────────────────

  const form = $('contractForm');
  const startDateInput = $('startDate');
  const endDateInput = $('endDate');

  // Duration preview
  function updateDuration() {
    const start = startDateInput.value;
    const end = endDateInput.value;
    if (start && end) {
      const diff = Math.ceil((new Date(end) - new Date(start)) / (1000 * 60 * 60 * 24));
      $('durationText').textContent = diff > 0 ? `${diff} วัน (${Math.round(diff / 30)} เดือน)` : 'วันที่ไม่ถูกต้อง';
    }
  }

  startDateInput.addEventListener('change', updateDuration);
  endDateInput.addEventListener('change', updateDuration);

  form.addEventListener('submit', e => {
    e.preventDefault();

    const alertDays = [];
    if ($('alert90').checked) alertDays.push(90);
    if ($('alert60').checked) alertDays.push(60);
    if ($('alert30').checked) alertDays.push(30);
    if ($('alert7').checked) alertDays.push(7);

    const newContract = {
      contract_id: generateId('WC'),
      po_number: $('poNumber').value,
      project_name: $('projectName').value,
      customer_name: $('customerName').value,
      service_type: $('serviceType').value,
      start_date: $('startDate').value,
      end_date: $('endDate').value,
      recipients_sale: $('recipientsSale').value,
      recipients_eng: $('recipientsEng').value,
      teams_webhook: $('teamsWebhook').value,
      note: $('contractNote').value,
      status: 'active',
    };

    sampleContracts.push(newContract);
    showToast('success', `✅ บันทึกสัญญา ${newContract.po_number} เรียบร้อย`);

    form.reset();
    renderDashboard();
    renderContractsList(currentFilter);
    navigateTo('contracts');
  });

  // Cancel button => navigate to dashboard
  $('btnCancel').addEventListener('click', () => navigateTo('dashboard'));

  // View all button
  $('btnViewAll').addEventListener('click', () => navigateTo('contracts'));

  // Add new button on contracts page
  $('btnAddNew').addEventListener('click', () => navigateTo('add-contract'));

  // ── Toast ────────────────────────────────────────────────────

  function showToast(type, message) {
    const container = $('toastContainer');
    const toast = document.createElement('div');
    toast.className = `toast ${type}`;
    toast.innerHTML = `
      <span>${type === 'success' ? '✅' : '❌'}</span>
      <span>${message}</span>
    `;
    container.appendChild(toast);

    setTimeout(() => {
      toast.remove();
    }, 3200);
  }

  // ── Initialize ───────────────────────────────────────────────

  renderDashboard();
  renderContractsList();
  renderLogs();

})();
