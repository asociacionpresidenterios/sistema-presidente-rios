import re

from flask import request, session

from app import app, index as original_index


# Dedicated player-list endpoint. The existing index view already contains the
# correct query and template, so the sports/data logic remains untouched.
if "admin_jugadores" not in app.view_functions:
    def admin_jugadores():
        admin_id = session.pop("admin_id", None)
        try:
            return original_index()
        finally:
            if admin_id is not None:
                session["admin_id"] = admin_id

    app.add_url_rule(
        "/admin/jugadores",
        endpoint="admin_jugadores",
        view_func=admin_jugadores,
        methods=["GET"],
    )


# Existing templates still reference url_for('index'). During an admin session,
# make those links stay inside Gestión de jugadores instead of returning Home.
_original_update_template_context = app.update_template_context


def _update_template_context(ctx, context):
    _original_update_template_context(ctx, context)
    if session.get("admin_id") and "url_for" in context:
        original_url_for = context["url_for"]

        def navigation_url_for(endpoint, *args, **kwargs):
            if endpoint == "index":
                endpoint = "admin_jugadores"
            return original_url_for(endpoint, *args, **kwargs)

        context["url_for"] = navigation_url_for


app.update_template_context = _update_template_context


@app.after_request
def repair_public_menu(response):
    if request.path != "/publico":
        return response
    if not response.content_type or "text/html" not in response.content_type:
        return response

    html = response.get_data(as_text=True)
    if 'data-target="posiciones"' not in html:
        return response

    match = re.search(r'value="champ-(\d+)"', html)
    if not match:
        return response
    campeonato_id = match.group(1)

    menu_old = '''<button class="active" data-target="posiciones" onclick="mostrarSeccion('posiciones',this)">📊 Posiciones</button>
<button data-target="programacion" onclick="mostrarSeccion('programacion',this)">📅 Programación</button>
<button data-target="resultados" onclick="mostrarSeccion('resultados',this)">⚽ Resultados</button>
<button data-target="goleadores" onclick="mostrarSeccion('goleadores',this)">🥅 Goleadores</button>
<button data-target="estadisticas" onclick="mostrarSeccion('estadisticas',this)">📈 Estadísticas</button>
<button data-target="fairplay" onclick="mostrarSeccion('fairplay',this)">🟧 Fair Play</button>
<button data-target="enlaces" onclick="mostrarSeccion('enlaces',this)">🔗 Más información</button>'''

    routes = {
        "posiciones": f"/publico/campeonato/{campeonato_id}/tabla",
        "programacion": f"/publico/campeonato/{campeonato_id}/programacion",
        "resultados": f"/publico/campeonato/{campeonato_id}/resultados",
        "goleadores": f"/publico/campeonato/{campeonato_id}/goleadores",
        "estadisticas": f"/publico/campeonato/{campeonato_id}/estadisticas",
        "fairplay": f"/publico/campeonato/{campeonato_id}/fair-play",
        "enlaces": f"/publico/campeonato/{campeonato_id}",
    }
    labels = {
        "posiciones": "📊 Posiciones",
        "programacion": "📅 Programación",
        "resultados": "⚽ Resultados",
        "goleadores": "🥅 Goleadores",
        "estadisticas": "📈 Estadísticas",
        "fairplay": "🟧 Fair Play",
        "enlaces": "🔗 Más información",
    }

    menu_new = "\n".join(
        f'<a class="menu-public-link{(" active" if key == "posiciones" else "")}" '
        f'data-target="{key}" href="{routes[key]}">{labels[key]}</a>'
        for key in routes
    )

    if menu_old in html:
        html = html.replace(menu_old, menu_new, 1)

    style = '''<style id="button-navigation-fix">\n.menu-public-link{width:100%;display:block;box-sizing:border-box;border:0;border-bottom:1px solid #edf0f4;background:#fff;text-align:left;padding:13px 14px;font-weight:800;color:#172033;cursor:pointer;text-decoration:none}.menu-public-link.active{background:#eaf2ff;color:#1459bd;border-left:4px solid #1769e0}@media(max-width:800px){.menu-public-link{display:block;flex:0 0 auto;width:auto;border-bottom:0;border-right:1px solid #edf0f4;padding:12px 13px;font-size:12px}.menu-public-link.active{border-left:0;border-bottom:3px solid #1769e0}}\n</style>'''
    if "button-navigation-fix" not in html:
        html = html.replace("</head>", style + "</head>", 1)

    script = f'''<script id="button-navigation-links">\n(function(){{\n  const base={{posiciones:'/publico/campeonato/ID/tabla',programacion:'/publico/campeonato/ID/programacion',resultados:'/publico/campeonato/ID/resultados',goleadores:'/publico/campeonato/ID/goleadores',estadisticas:'/publico/campeonato/ID/estadisticas',fairplay:'/publico/campeonato/ID/fair-play',enlaces:'/publico/campeonato/ID'}};\n  function actualizar(){{\n    const select=document.getElementById('champSelect');\n    const raw=select && select.value ? select.value.replace('champ-','') : '{campeonato_id}';\n    document.querySelectorAll('#menu a.menu-public-link[data-target]').forEach(function(a){{\n      const template=base[a.dataset.target];\n      if(template)a.href=template.replace('ID',raw);\n    }});\n  }}\n  actualizar();\n  const select=document.getElementById('champSelect');\n  if(select)select.addEventListener('change',function(){{setTimeout(actualizar,0);}});\n}})();\n</script>'''
    if "button-navigation-links" not in html:
        html = html.replace("</body>", script + "</body>", 1)

    response.set_data(html)
    return response
