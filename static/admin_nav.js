document.addEventListener('DOMContentLoaded', function () {
  const nav = document.querySelector('.side-scroll');
  if (!nav) return;

  const labels = Array.from(nav.querySelectorAll('.section-label'));
  labels.forEach(function (label, index) {
    const button = document.createElement('button');
    button.type = 'button';
    button.className = 'nav-section-toggle';
    button.innerHTML = '<span>' + label.textContent.trim() + '</span><span class="nav-section-chevron">⌄</span>';
    label.replaceWith(button);

    const items = [];
    let node = button.nextElementSibling;
    while (node && !node.classList.contains('nav-section-toggle')) {
      items.push(node);
      node = node.nextElementSibling;
    }

    const hasActive = items.some(function (el) { return el.classList && el.classList.contains('active'); });
    const key = 'pr-admin-nav-' + index;
    const saved = localStorage.getItem(key);
    const open = saved === null ? hasActive || index === 0 : saved === '1';

    button.setAttribute('aria-expanded', open ? 'true' : 'false');
    button.addEventListener('click', function () {
      const next = button.getAttribute('aria-expanded') !== 'true';
      button.setAttribute('aria-expanded', next ? 'true' : 'false');
      localStorage.setItem(key, next ? '1' : '0');
    });
  });
});
