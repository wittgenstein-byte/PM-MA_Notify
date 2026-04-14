// ── Theme Toggle ───────────────────────────────────────────────

  function getSavedTheme() {
    return localStorage.getItem('pm_ma_theme') || 'dark';
  }

  function applyTheme(theme) {
    if (theme === 'light') {
      document.documentElement.setAttribute('data-theme', 'light');
    } else {
      document.documentElement.removeAttribute('data-theme');
    }

    const toggles = [$('themeToggleSidebar'), $('themeToggleMobile')].filter(Boolean);
    toggles.forEach(btn => {
      const sun = btn.querySelector('.sun-icon');
      const moon = btn.querySelector('.moon-icon');
      if (theme === 'light') {
        if (sun) sun.style.display = 'none';
        if (moon) moon.style.display = 'block';
      } else {
        if (sun) sun.style.display = 'block';
        if (moon) moon.style.display = 'none';
      }
    });
  }

  function toggleTheme() {
    const current = getSavedTheme();
    const newTheme = current === 'dark' ? 'light' : 'dark';
    localStorage.setItem('pm_ma_theme', newTheme);
    applyTheme(newTheme);
  }

  const toggleSidebar = $('themeToggleSidebar');
  const toggleMobile = $('themeToggleMobile');
  if (toggleSidebar) toggleSidebar.addEventListener('click', toggleTheme);
  if (toggleMobile) toggleMobile.addEventListener('click', toggleTheme);

  applyTheme(getSavedTheme());

  