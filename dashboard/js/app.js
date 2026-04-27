// ============================================================
//  app.js — Application Bootstrap (Coordinator)
//  GRASP: Controller — wires ทุก module เข้าด้วยกัน
//  ทำหน้าที่เดียวคือ: สร้าง instances + inject dependencies
// ============================================================

(function () {
  'use strict';

  // ── Shared state ──────────────────────────────────────────
  let contracts = [...sampleContracts]; // start with sample, replace when Sheets loads
  let logs = [...sampleLogs];

  // ── Shared helper (used across modules) ──────────────────
  function renderStatusBadge(category) {
    const labels = {
      critical: '🔴 วิกฤต',
      warning:  '🟡 ปานกลาง',
      safe:     '🟢 ปกติ',
      expired:  '⚫ หมดอายุ',
    };
    return `<span class="badge badge-${category}">${labels[category] || category}</span>`;
  }

  // ── Instantiate Services (DIP: inject URL from config) ───
  const gasUrl         = (window.__ENV__ && window.__ENV__.GAS_SCRIPT_URL) || '';
  const contractSvc    = new ContractService(gasUrl);
  const emailHistorySvc = new EmailHistoryService();

  // ── Instantiate UI modules ────────────────────────────────
  const router      = new Router();
  const toast       = new ToastUI('toastContainer');
  const modal       = new ModalUI(renderStatusBadge, onEditContract, onDeleteContract);
  const dashboardUI = new DashboardUI(openModal, renderStatusBadge);
  const contractsUI = new ContractsUI(openModal, renderStatusBadge, onEditContract, onDeleteContract);
  const logsUI      = new LogsUI();

  const addContractUI = new AddContractUI(
    contractSvc,
    emailHistorySvc,
    toast,
    () => contracts,           // getter — ไม่ copy ค่า
    onContractAdded,
    (page) => router.navigateTo(page)
  );

  // Wire edit-mode callback
  addContractUI.setOnUpdated((contractId, updatedData) => {
    const idx = contracts.findIndex(c => c.contract_id === contractId);
    if (idx !== -1) {
      contracts[idx] = { ...contracts[idx], ...updatedData };
    }
    refreshAllViews();
  });

  // ── Callbacks: CRUD Operations ────────────────────────────

  function onContractAdded(newContract) {
    contracts.push(newContract);
    refreshAllViews();
  }

  function onEditContract(contractId) {
    const contract = contracts.find(c => c.contract_id === contractId);
    if (!contract) return;
    addContractUI.enterEditMode(contract);
    router.navigateTo('add-contract');
  }

  async function onDeleteContract(contractId) {
    const contract = contracts.find(c => c.contract_id === contractId);
    if (!contract) return;

    const confirmed = confirm(
      `⚠️ ยืนยันการลบสัญญา?\n\n` +
      `PO: ${contract.po_number}\n` +
      `โครงการ: ${contract.project_name}\n\n` +
      `การลบจะลบ notification rules ที่เกี่ยวข้องทั้งหมดด้วย`
    );
    if (!confirmed) return;

    // Close modal if open
    modal.close();

    if (contractSvc.isDemoMode()) {
      // Demo mode — ลบ local เท่านั้น
      contracts = contracts.filter(c => c.contract_id !== contractId);
      toast.warning('⚠️ ลบเฉพาะ local session (demo mode)');
      refreshAllViews();
      return;
    }

    try {
      toast.info('🗑️ กำลังลบสัญญา...');
      const result = await contractSvc.deleteContract(contractId);

      if (result.status === 'ok') {
        contracts = contracts.filter(c => c.contract_id !== contractId);
        toast.success(`✅ ลบสัญญา ${contract.po_number} เรียบร้อย (ลบ notification rules ด้วย)`);
        refreshAllViews();
      } else {
        toast.error(`❌ ลบไม่สำเร็จ: ${result.message || 'ไม่ทราบสาเหตุ'}`);
      }
    } catch (err) {
      console.error('Delete error:', err);
      toast.error(`❌ เชื่อมต่อไม่ได้: ${err.message}`);
    }
  }

  // ── Helper: เปิด modal พร้อม enrich ────────────────────
  function openModal(contractId) {
    const contract = contracts.find(c => c.contract_id === contractId);
    if (!contract) return;
    const daysLeft = calcDaysLeft(contract.end_date);
    const category = getStatusCategory(daysLeft, contract.status);
    modal.open({ ...contract, daysLeft, category });
  }

  // ── Helper: refresh ทุก view ─────────────────────────────
  function refreshAllViews() {
    dashboardUI.render(contracts, logs);
    contractsUI.render(contracts);
    logsUI.render(logs);
    addContractUI.refreshRecentEmails();
    updateConnectionBadge();
  }

  // ── Connection Status Badge ─────────────────────────────
  function updateConnectionBadge() {
    const badge = document.getElementById('connectionBadge');
    if (!badge) return;

    if (contractSvc.isDemoMode()) {
      badge.className = 'connection-badge demo';
      badge.innerHTML = '<span class="conn-dot"></span> Demo Mode';
    } else {
      badge.className = 'connection-badge connected';
      badge.innerHTML = '<span class="conn-dot"></span> เชื่อมต่อ Sheets';
    }
  }

  // ── Load data from Sheets ────────────────────────────────
  async function loadFromSheets() {
    if (contractSvc.isDemoMode()) {
      updateConnectionBadge();
      return;
    }

    const syncBtn = document.getElementById('btnSync');
    if (syncBtn) {
      syncBtn.disabled = true;
      syncBtn.innerHTML = '<span class="sync-spinner"></span> กำลังโหลด...';
    }

    try {
      const [contractsResult, logsResult] = await Promise.all([
        contractSvc.fetchContracts(),
        contractSvc.fetchLogs(),
      ]);

      if (contractsResult.status === 'ok' && contractsResult.data) {
        contracts = contractsResult.data;
        toast.success(`✅ โหลดข้อมูลจาก Sheets สำเร็จ (${contracts.length} สัญญา)`);
      }

      if (logsResult.status === 'ok' && logsResult.data) {
        logs = logsResult.data;
      }

      refreshAllViews();
    } catch (err) {
      console.error('Fetch from Sheets error:', err);
      toast.error('⚠️ โหลดข้อมูลจาก Sheets ไม่ได้ — ใช้ sample data แทน');
    } finally {
      if (syncBtn) {
        syncBtn.disabled = false;
        syncBtn.innerHTML = `
          <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2">
            <polyline points="23 4 23 10 17 10"/><path d="M20.49 15a9 9 0 1 1-2.12-9.36L23 10"/>
          </svg>
          Sync
        `;
      }
    }
  }

  // ── Navigation shortcut buttons ─────────────────────────
  document.getElementById('btnViewAll')?.addEventListener('click', () => router.navigateTo('contracts'));
  document.getElementById('btnAddNew')?.addEventListener('click',  () => router.navigateTo('add-contract'));
  document.getElementById('btnSync')?.addEventListener('click', loadFromSheets);

  // ── Date display ─────────────────────────────────────────
  document.getElementById('currentDate').textContent = formatThaiDate(today);

  // ── Initialize ───────────────────────────────────────────
  refreshAllViews();
  loadFromSheets(); // Attempt to load real data

})();
