document.addEventListener('DOMContentLoaded', async function () {
  const page = document.querySelector('.page');
  if (!page) return;

  const titles = Array.from(page.querySelectorAll('.section-title h2'));
  const nextTitle = titles.find(function (el) { return el.textContent.trim() === 'Próximos partidos'; });
  if (!nextTitle) return;

  nextTitle.textContent = 'Próxima fecha';
  const titleWrap = nextTitle.closest('.section-title');
  const subtitle = titleWrap && titleWrap.querySelector('span');
  if (subtitle) subtitle.textContent = 'Partidos programados';

  const champLink = page.querySelector('.champ .btn-mini[href*="panel_campeonato"]');
  if (!champLink) return;

  const match = champLink.getAttribute('href').match(/(\d+)(?:\/)?$/);
  if (!match) return;
  const campeonatoId = match[1];

  try {
    const response = await fetch('/publico/campeonato/' + campeonatoId + '/tabla', { credentials: 'same-origin' });
    if (!response.ok) return;
    const html = await response.text();
    const doc = new DOMParser().parseFromString(html, 'text/html');
    const sourceTable = doc.querySelector('table');
    if (!sourceTable) return;

    const section = document.createElement('div');
    section.className = 'section-title pr-standings-title';
    section.innerHTML = '<h2>Tabla de posiciones</h2><span>Campeonato activo</span><a class="alert-link" href="/publico/campeonato/' + campeonatoId + '/tabla" target="_blank" rel="noopener">Ver tabla completa →</a>';

    const card = document.createElement('div');
    card.className = 'card pr-standings-card';
    card.innerHTML = '<div class="card-head"><strong>Posiciones actuales</strong><span class="pill active">Actualizada</span></div>';
    const wrap = document.createElement('div');
    wrap.className = 'table-wrap';
    const table = sourceTable.cloneNode(true);
    table.className = 'table';
    wrap.appendChild(table);
    card.appendChild(wrap);

    const css = document.createElement('style');
    css.textContent = '.pr-standings-card{margin-bottom:18px}.pr-standings-card .table{min-width:620px}.pr-standings-card .table th,.pr-standings-card .table td{text-align:center}.pr-standings-card .table th:first-child,.pr-standings-card .table td:first-child{text-align:left}.pr-standings-title{margin-top:22px}.pr-standings-title .alert-link{margin-left:auto}.pr-standings-card .table tbody tr:hover{background:#f7faff}';
    document.head.appendChild(css);

    titleWrap.parentNode.insertBefore(section, titleWrap);
    titleWrap.parentNode.insertBefore(card, titleWrap);
  } catch (error) {
    console.warn('No fue posible cargar la tabla de posiciones en el dashboard.', error);
  }
});
