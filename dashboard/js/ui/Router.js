// ============================================================
//  Router — GRASP: Indirection + Controller
//  SRP: รับผิดชอบเฉพาะ SPA navigation
//  OCP: เพิ่ม page ใหม่ได้โดยไม่แก้ core logic
// ============================================================

class Router {
  constructor() {
    this._pages = document.querySelectorAll('.page');
    this._navItems = document.querySelectorAll('.nav-item');
    this._bindNavEvents();
    this._bindMobileMenu();
  }

  /**
   * นำทางไปยัง page ที่กำหนด
   * @param {string} pageId - e.g. 'dashboard', 'contracts'
   */
  navigateTo(pageId) {
    this._pages.forEach(p => p.classList.remove('active'));
    this._navItems.forEach(n => n.classList.remove('active'));

    const targetPage = document.getElementById(`page-${pageId}`);
    const targetNav = document.querySelector(`[data-page="${pageId}"]`);

    if (targetPage) targetPage.classList.add('active');
    if (targetNav) targetNav.classList.add('active');

    // Close mobile sidebar
    const sidebar = document.getElementById('sidebar');
    if (sidebar) sidebar.classList.remove('open');
  }

  _bindNavEvents() {
    this._navItems.forEach(item => {
      item.addEventListener('click', e => {
        e.preventDefault();
        this.navigateTo(item.dataset.page);
      });
    });
  }

  _bindMobileMenu() {
    const menuToggle = document.getElementById('menuToggle');
    const sidebar = document.getElementById('sidebar');
    if (menuToggle && sidebar) {
      menuToggle.addEventListener('click', () => sidebar.classList.toggle('open'));
    }
  }
}
