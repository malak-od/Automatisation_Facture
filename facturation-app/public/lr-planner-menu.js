// Préférences locales uniquement ; aucune identité fictive ni nouvelle session.
(() => {
  const root = document.documentElement;
  const theme = document.getElementById('lr-theme');
  const reduce = document.getElementById('lr-reduce');
  const mobile = document.getElementById('lr-open');
  function save(key, value) { try { localStorage.setItem(key, value); } catch {} }
  function refresh() {
    const dark = root.dataset.theme === 'dark';
    theme.querySelector('.lbl').textContent = dark ? 'Mode clair' : 'Mode sombre';
    theme.title = dark ? 'Mode clair' : 'Mode sombre';
    reduce.title = root.dataset.sidebar === 'collapsed' ? 'Agrandir le menu' : 'Réduire';
  }
  theme.addEventListener('click', () => {
    root.dataset.theme = root.dataset.theme === 'dark' ? 'light' : 'dark';
    save('facturation-theme', root.dataset.theme); refresh();
  });
  reduce.addEventListener('click', () => {
    root.dataset.sidebar = root.dataset.sidebar === 'collapsed' ? 'expanded' : 'collapsed';
    save('facturation-sidebar', root.dataset.sidebar); refresh();
  });
  const aside = document.querySelector('.lr-sidebar');
  const smallScreen = window.matchMedia('(max-width: 900px)');
  function syncAccess() { aside.inert = smallScreen.matches && !root.classList.contains('lr-menu-open'); }
  function close() { root.classList.remove('lr-menu-open'); mobile.setAttribute('aria-expanded', 'false'); syncAccess(); }
  smallScreen.addEventListener('change', () => { close(); });
  syncAccess();
  mobile.addEventListener('click', () => {
    const open = root.classList.toggle('lr-menu-open'); mobile.setAttribute('aria-expanded', String(open));
    syncAccess();
    if (open) document.querySelector('.lr-sidebar button').focus();
  });
  document.getElementById('lr-backdrop').addEventListener('click', close);
  document.addEventListener('keydown', event => { if (event.key === 'Escape' && root.classList.contains('lr-menu-open')) { close(); mobile.focus(); } });
  document.getElementById('lr-home').addEventListener('click', () => { goHome(); close(); });
  refresh();
})();
