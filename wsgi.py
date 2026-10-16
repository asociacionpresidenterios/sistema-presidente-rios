from flask import redirect, request, session, url_for
from flask import render_template

from app import app, db, Jugador, index as original_index


# Public portal: the main domain opens directly on the official public face.
# The original Players view is kept available for the administrative module.
def public_root():
    return redirect("/publico")

if "index" in app.view_functions:
    app.view_functions["index"] = public_root


# Real administrative Players endpoint. It reuses the existing player view only
# for its data/template; the root URL is reserved for the public portal.
if "admin_jugadores" not in app.view_functions:
    def admin_jugadores():
        if not session.get("admin_id"):
            return redirect(url_for("login", next=request.path))

        admin_id = session.pop("admin_id", None)
        try:
            return original_index()
        finally:
            if admin_id is not None:
                session["admin_id"] = admin_id

    app.add_url_rule("/admin/jugadores", endpoint="admin_jugadores", view_func=admin_jugadores, methods=["GET"])


# Keep existing templates that still use url_for('index') inside the Players
# module when an administrator is logged in. This is a Jinja global override,
# not a Flask request-context override.
_original_url_for = app.jinja_env.globals.get("url_for")
if _original_url_for:
    def _navigation_url_for(endpoint, *args, **kwargs):
        if endpoint == "index" and session.get("admin_id"):
            endpoint = "admin_jugadores"
        return _original_url_for(endpoint, *args, **kwargs)
    app.jinja_env.globals["url_for"] = _navigation_url_for


@app.after_request
def add_public_back_navigation(response):
    path = request.path or ""
    if not path.startswith("/publico/"):
        return response
    if not response.content_type or "text/html" not in response.content_type:
        return response

    html = response.get_data(as_text=True)
    if "public-back-navigation" in html:
        return response

    style = '''<style id="public-back-navigation">\n.public-nav{max-width:1240px;margin:0 auto;padding:10px 18px;display:flex;gap:8px;flex-wrap:wrap}.public-nav a{display:inline-block;text-decoration:none;background:#fff;color:#173b70;border:1px solid #dfe5ec;border-radius:9px;padding:8px 11px;font:800 12px Arial,sans-serif}.public-nav a.primary{background:#1769e0;color:#fff;border-color:#1769e0}\n</style>'''
    nav = '''<nav class="public-nav" aria-label="Navegación del portal">\n<a class="primary" href="/publico">← Portal público</a>\n<a href="javascript:history.back()">← Volver</a>\n</nav>'''
    html = html.replace("</head>", style + "</head>", 1)
    if "<main" in html:
        html = html.replace("<main", nav + "<main", 1)
    else:
        html = html.replace("<body>", "<body>" + nav, 1)
    response.set_data(html)
    return response
