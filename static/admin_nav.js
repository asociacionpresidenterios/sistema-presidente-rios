document.addEventListener('DOMContentLoaded', function () {
  const nav = document.querySelector('.side-scroll');
  if (!nav) return;

  const style = document.createElement('style');
  style.textContent = '.nav-section-toggle{width:100%;display:flex;align-items:center;justify-content:space-between;border:0;background:transparent;color:#7890ae;font:800 9px/1.2 Inter,-apple-system,BlinkMacSystemFont,"Segoe UI",Roboto,Arial,sans-serif;letter-spacing:.12em;text-transform:uppercase;padding:12px 10px 6px;cursor:pointer;text-align:left}.nav-section-toggle:hover{color:#d9e5f3}.nav-section-chevron{font-size:13px;transition:transform .18s ease}.nav-section-toggle[aria-expanded="false"] .nav-section-chevron{transform:rotate(-90deg)}';
  document.head.appendChild(style);

  const labels = Array.from(nav.querySelectorAll('.section-label'));
  labels.forEach(function (label, index) {
    const button = document.createElement('button');
    button.type = 'button';
    button.className = 'nav-section-toggle';
    button.innerHTML = '<span>' + label.textContent.trim() + '</span><span class="nav-section-chevron">⌄</span>';
    label.replaceWith(button);

    const items = [];
    let node = button.nextElementSibling;
    while (node && !node.classList.contains('section-label') && !node.classList.contains('nav-section-toggle')) {
      items.push(node);
      node = node.nextElementSibling;
    }

    const hasActive = items.some(function (el) { return el.classList && el.classList.contains('active'); });
    const key = 'pr-admin-nav-' + index;
    const saved = localStorage.getItem(key);
    const open = saved === null ? hasActive || index === 0 : saved === '1';

    function apply(value) {
      button.setAttribute('aria-expanded', value ? 'true' : 'false');
      items.forEach(function (el) { el.style.display = value ? '' : 'none'; });
    }

    apply(open);
    button.addEventListener('click', function () {
      const next = button.getAttribute('aria-expanded') !== 'true';
      apply(next);
      localStorage.setItem(key, next ? '1' : '0');
    });
  });
});
