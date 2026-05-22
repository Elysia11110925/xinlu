/* ═══════════════ 导航高亮 ═══════════════ */
(function() {
  const sections = document.querySelectorAll('section[id]');
  const links = document.querySelectorAll('nav a');
  window.addEventListener('scroll', () => {
    let current = '';
    sections.forEach(s => {
      if (window.scrollY >= s.offsetTop - 120) current = s.id;
    });
    links.forEach(a => {
      a.classList.toggle('active', a.getAttribute('href') === '#' + current);
    });
  });
})();

/* ═══════════════ 回到顶部按钮 ═══════════════ */
(function() {
  const btn = document.querySelector('.back-to-top');
  if (!btn) return;

  window.addEventListener('scroll', () => {
    btn.classList.toggle('visible', window.scrollY > 500);
  });

  btn.addEventListener('click', () => {
    window.scrollTo({ top: 0, behavior: 'smooth' });
  });
})();

/* ═══════════════ 滚动入场动画 ═══════════════ */
(function() {
  const observer = new IntersectionObserver((entries) => {
    entries.forEach(entry => {
      if (entry.isIntersecting) {
        entry.target.classList.add('visible');
        observer.unobserve(entry.target);
      }
    });
  }, { threshold: 0.1, rootMargin: '0px 0px -40px 0px' });

  document.querySelectorAll('.animate-on-scroll').forEach(el => {
    observer.observe(el);
  });
})();

/* ═══════════════ 暗色模式切换 ═══════════════ */
(function() {
  const toggle = document.querySelector('.theme-toggle');
  if (!toggle) return;

  const html = document.documentElement;

  // 从 localStorage 恢复
  const saved = localStorage.getItem('theme');
  if (saved) {
    html.setAttribute('data-theme', saved);
  }

  toggle.addEventListener('click', () => {
    const current = html.getAttribute('data-theme');
    const next = current === 'dark' ? 'light' : 'dark';
    html.setAttribute('data-theme', next);
    localStorage.setItem('theme', next);
  });

  // 更新按钮图标
  function updateIcon() {
    const isDark = getComputedStyle(html).getPropertyValue('--bg').trim() === '#1a1a1a';
    toggle.textContent = isDark ? '☀' : '🌙';
  }
  updateIcon();

  // 监听变化
  new MutationObserver(updateIcon).observe(html, { attributes: true, attributeFilter: ['data-theme'] });
})();
