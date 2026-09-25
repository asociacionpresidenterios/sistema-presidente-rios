from pathlib import Path
from datetime import datetime
import shutil

ROOT = Path(__file__).resolve().parent
APP = ROOT / "app.py"
BASE = ROOT / "templates" / "base.html"
TEMPLATE_SRC = ROOT / "templates" / "admin_centro_jugadores.html"

MARKER = "# V6.9 — CENTRO MAESTRO DE JUGADORES"

ROUTE = '# ============================================================\n# V6.9 — CENTRO MAESTRO DE JUGADORES\n# ============================================================\n\n@app.route("/admin/jugadores/centro")\ndef admin_centro_jugadores():\n    """Vista maestra del registro único de jugadores.\n\n    Reutiliza el modelo Jugador y la información ya existente en\n    PartidoJugador. No crea una segunda tabla ni modifica PostgreSQL.\n    """\n    q = request.args.get("q", "").strip()\n    club_filtro = request.args.get("club", "").strip()\n    serie_filtro = request.args.get("serie", "").strip()\n    estado_filtro = request.args.get("estado", "").strip()\n\n    query = Jugador.query\n\n    if q:\n        patron = f"%{q}%"\n        query = query.filter(\n            db.or_(\n                Jugador.nombre_completo.ilike(patron),\n                Jugador.rut.ilike(patron),\n            )\n        )\n\n    if club_filtro:\n        query = query.filter(Jugador.club == club_filtro)\n\n    if serie_filtro:\n        query = query.filter(Jugador.serie == serie_filtro)\n\n    if estado_filtro:\n        query = query.filter(Jugador.estado == estado_filtro)\n\n    jugadores = (\n        query\n        .order_by(Jugador.nombre_completo.asc(), Jugador.id.asc())\n        .limit(250)\n        .all()\n    )\n\n    ids = [j.id for j in jugadores]\n    stats = {}\n\n    if ids:\n        filas = (\n            db.session.query(\n                PartidoJugador.jugador_id,\n                db.func.count(PartidoJugador.id).label("partidos"),\n                db.func.coalesce(db.func.sum(PartidoJugador.goles), 0).label("goles"),\n                db.func.coalesce(db.func.sum(PartidoJugador.amarillas), 0).label("amarillas"),\n                db.func.coalesce(db.func.sum(PartidoJugador.rojas), 0).label("rojas"),\n            )\n            .filter(PartidoJugador.jugador_id.in_(ids))\n            .group_by(PartidoJugador.jugador_id)\n            .all()\n        )\n\n        for fila in filas:\n            stats[fila.jugador_id] = {\n                "partidos": int(fila.partidos or 0),\n                "goles": int(fila.goles or 0),\n                "amarillas": int(fila.amarillas or 0),\n                "rojas": int(fila.rojas or 0),\n            }\n\n    resumen = []\n    for jugador in jugadores:\n        resumen.append({\n            "jugador": jugador,\n            **stats.get(jugador.id, {\n                "partidos": 0,\n                "goles": 0,\n                "amarillas": 0,\n                "rojas": 0,\n            }),\n        })\n\n    clubes = [\n        valor[0] for valor in db.session.query(Jugador.club)\n        .filter(Jugador.club.isnot(None), Jugador.club != "")\n        .distinct()\n        .order_by(Jugador.club.asc())\n        .all()\n    ]\n\n    series = [\n        valor[0] for valor in db.session.query(Jugador.serie)\n        .filter(Jugador.serie.isnot(None), Jugador.serie != "")\n        .distinct()\n        .order_by(Jugador.serie.asc())\n        .all()\n    ]\n\n    estados = sorted(ESTADOS_PERMITIDOS)\n\n    return render_template(\n        "admin_centro_jugadores.html",\n        resumen=resumen,\n        q=q,\n        club_filtro=club_filtro,\n        serie_filtro=serie_filtro,\n        estado_filtro=estado_filtro,\n        clubes=clubes,\n        series=series,\n        estados=estados,\n        total=Jugador.query.count(),\n        vigentes=Jugador.query.filter_by(estado="Vigente").count(),\n        clubes_total=Jugador.query.filter(\n            Jugador.club.isnot(None), Jugador.club != ""\n        ).with_entities(Jugador.club).distinct().count(),\n        participaciones=PartidoJugador.query.count(),\n    )\n\n'

NAV = '    <a class="nav-item {% if request.endpoint==\'admin_centro_jugadores\' %}active{% endif %}" href="{{ url_for(\'admin_centro_jugadores\') }}"><span class="nav-icon">🧑\u200d💼</span>Centro maestro de jugadores</a>\n'

def backup(path):
    stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    dest = path.with_name(path.name + f".V69_BACKUP_{stamp}")
    shutil.copy2(path, dest)
    return dest

if not APP.exists():
    raise SystemExit("No se encontró app.py. Ejecuta este instalador en la carpeta raíz del proyecto.")
if not BASE.exists():
    raise SystemExit("No se encontró templates/base.html.")
if not TEMPLATE_SRC.exists():
    raise SystemExit("No se encontró templates/admin_centro_jugadores.html.")

app = APP.read_text(encoding="utf-8")
base = BASE.read_text(encoding="utf-8")

if MARKER not in app:
    anchor = '# ============================================================\n# FICHA ADMINISTRATIVA DE CLUB'
    if anchor not in app:
        anchor = '@app.route("/admin/clubes")'
    pos = app.find(anchor)
    if pos < 0:
        raise SystemExit("No se encontró un punto seguro para insertar V6.9 en app.py.")
    app = app[:pos] + ROUTE + "\n" + app[pos:]
    backup(APP)
    APP.write_text(app, encoding="utf-8")
    print("OK: ruta V6.9 agregada a app.py")
else:
    print("OK: V6.9 ya existe en app.py; no se modificó.")

if "Centro maestro de jugadores" not in base:
    anchor = '    <a class="nav-item {% if request.endpoint in [\'admin_registro\',\'index\',\'nuevo_jugador\',\'editar_jugador\',\'importar_jugadores\',\'ficha_jugador\',\'historial_jugador\'] %}active{% endif %}" href="{{ url_for(\'admin_registro\') }}"><span class="nav-icon">👥</span>Registro de jugadores</a>\n'
    if anchor not in base:
        raise SystemExit("No se encontró el enlace de Registro de jugadores en base.html.")
    base = base.replace(anchor, anchor + NAV, 1)
    backup(BASE)
    BASE.write_text(base, encoding="utf-8")
    print("OK: acceso V6.9 agregado a base.html")
else:
    print("OK: acceso V6.9 ya existe en base.html; no se modificó.")

print("V6.9 instalada de forma incremental. Se crearon respaldos antes de modificar archivos.")
