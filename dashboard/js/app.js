// ============================================================
//  app.js — Application Bootstrap (Coordinator)
//  GRASP: Controller — wires ทุก module เข้าด้วยกัน
//  ทำหน้าที่เดียวคือ: สร้าง instances + inject dependencies
// ============================================================

(function () {
  'use strict';

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
  const modal       = new ModalUI(renderStatusBadge);
  const dashboardUI = new DashboardUI(openModal, renderStatusBadge);
  const contractsUI = new ContractsUI(openModal, renderStatusBadge);
  const logsUI      = new LogsUI();

  const addContractUI = new AddContractUI(
    contractSvc,
    emailHistorySvc,
    toast,
    () => sampleContracts,           // getter — ไม่ copy ค่า
    onContractAdded,
    (page) => router.navigateTo(page)
  );

  // ── Callback: เมื่อเพิ่มสัญญาใหม่ ──────────────────────
  function onContractAdded(newContract) {
    sampleContracts.push(newContract);
    refreshAllViews();
  }

  // ── Helper: เปิด modal พร้อม enrich ────────────────────
  function openModal(contractId) {
    const contract = sampleContracts.find(c => c.contract_id === contractId);
    if (!contract) return;
    const daysLeft = calcDaysLeft(contract.end_date);
    const category = getStatusCategory(daysLeft, contract.status);
    modal.open({ ...contract, daysLeft, category });
  }

  // ── Helper: refresh ทุก view ─────────────────────────────
  function refreshAllViews() {
    dashboardUI.render(sampleContracts, sampleLogs);
    contractsUI.render(sampleContracts);
    logsUI.render(sampleLogs);
    addContractUI.refreshRecentEmails();
  }

  // ── Navigation shortcut buttons ─────────────────────────
  document.getElementById('btnViewAll')?.addEventListener('click', () => router.navigateTo('contracts'));
  document.getElementById('btnAddNew')?.addEventListener('click',  () => router.navigateTo('add-contract'));

  // ── Date display ─────────────────────────────────────────
  document.getElementById('currentDate').textContent = formatThaiDate(today);

  // ── Initialize ───────────────────────────────────────────
  refreshAllViews();

})();
