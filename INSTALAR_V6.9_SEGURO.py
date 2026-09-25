from pathlib import Path
from datetime import datetime
import shutil, ast

ROOT = Path(__file__).resolve().parent
APP = ROOT / 'app.py'
BASE = ROOT / 'templates' / 'base.html'
SRC = ROOT / 'V6.9_FINAL_TEMPLATE' / 'admin_centro_jugadores.html'
DST = ROOT / 'templates' / 'admin_centro_jugadores.html'

ROUTE_MARKER = '@app.route("/admin/clubes")'
MENU_MARKER = "    <a class=\"nav-item {% if request.endpoint in ['admin_clubes','admin_planteles','admin_ficha_club','admin_centro_jugador','admin_nuevo_jugador_club','admin_editar_jugador_plantel','admin_cambiar_estado_plantel','admin_mover_jugador_plantel'] %}active{% endif %}\" href=\"{{ url_for('admin_clubes') }}\"><span class=\"nav-icon\">🏟️</span>Clubes y planteles</a>"
MENU_LINE = "    <a class=\"nav-item {% if request.endpoint=='admin_centro_jugadores' %}active{% endif %}\" href=\"{{ url_for('admin_centro_jugadores') }}\"><span class=\"nav-icon\">🧑‍💼</span>Centro maestro de jugadores</a>\n"

ROUTE_CODE = r'''
# ============================================================
# V6.9 — CENTRO MAESTRO DE JUGADORES
# ============================================================
@app.route("/admin/jugadores/centro")
@admin_required
def admin_centro_jugadores():
    q = request.args.get("q", "").strip()
    club_filtro = request.args.get("club", "").strip()
    serie_filtro = request.args.get("serie", "").strip()
    estado_filtro = request.args.get("estado", "").strip()

    query = Jugador.query
    if q:
        query = query.filter(db.or_(
            Jugador.nombre_completo.ilike(f"%{q}%"),
            Jugador.rut.ilike(f"%{q}%")
        ))
    if club_filtro:
        query = query.filter(Jugador.club == club_filtro)
    if serie_filtro:
        query = query.filter(Jugador.serie == serie_filtro)
    if estado_filtro:
        query = query.filter(Jugador.estado == estado_filtro)

    jugadores = query.order_by(Jugador.nombre_completo.asc()).limit(250).all()

    clubes = [x[0] for x in db.session.query(Jugador.club)
              .filter(Jugador.club.isnot(None), Jugador.club != "")
              .distinct().order_by(Jugador.club.asc()).all()]
    series = [x[0] for x in db.session.query(Jugador.serie)
              .filter(Jugador.serie.isnot(None), Jugador.serie != "")
              .distinct().order_by(Jugador.serie.asc()).all()]
    estados = [x[0] for x in db.session.query(Jugador.estado)
               .filter(Jugador.estado.isnot(None), Jugador.estado != "")
               .distinct().order_by(Jugador.estado.asc()).all()]

    resumen = []
    for jugador in jugadores:
        partidos = db.session.query(db.func.count(PartidoJugador.id)).filter(
            PartidoJugador.jugador_id == jugador.id
        ).scalar() or 0
        goles = db.session.query(db.func.coalesce(db.func.sum(PartidoJugador.goles), 0)).filter(
            PartidoJugador.jugador_id == jugador.id
        ).scalar() or 0
        amarillas = db.session.query(db.func.coalesce(db.func.sum(PartidoJugador.amarillas), 0)).filter(
            PartidoJugador.jugador_id == jugador.id
        ).scalar() or 0
        rojas = db.session.query(db.func.coalesce(db.func.sum(PartidoJugador.rojas), 0)).filter(
            PartidoJugador.jugador_id == jugador.id
        ).scalar() or 0
        resumen.append({
            "jugador": jugador,
            "partidos": int(partidos),
            "goles": int(goles),
            "amarillas": int(amarillas),
            "rojas": int(rojas),
        })

    total = Jugador.query.count()
    vigentes = Jugador.query.filter_by(estado="Vigente").count()
    clubes_total = db.session.query(Jugador.club).filter(
        Jugador.club.isnot(None), Jugador.club != ""
    ).distinct().count()
    participaciones = db.session.query(PartidoJugador).count()

    return render_template(
        "admin_centro_jugadores.html",
        resumen=resumen,
        total=total,
        vigentes=vigentes,
        clubes_total=clubes_total,
        participaciones=participaciones,
        clubes=clubes,
        series=series,
        estados=estados,
        q=q,
        club_filtro=club_filtro,
        serie_filtro=serie_filtro,
        estado_filtro=estado_filtro,
    )

'''

if not APP.exists() or not BASE.exists() or not SRC.exists():
    raise SystemExit('ERROR: coloca este instalador en la raiz del proyecto V6.8 y conserva V6.9_FINAL_TEMPLATE.')

app_text = APP.read_text(encoding='utf-8')
base_text = BASE.read_text(encoding='utf-8')

def backup(path):
    stamp = datetime.now().strftime('%Y%m%d_%H%M%S')
    shutil.copy2(path, path.with_name(path.name + '.V68_BACKUP_' + stamp))

if 'def admin_centro_jugadores(' not in app_text:
    if ROUTE_MARKER not in app_text:
        raise SystemExit('ABORTADO: marcador de app.py no encontrado. No se modifico nada.')
    new_app = app_text.replace(ROUTE_MARKER, ROUTE_CODE + '\n' + ROUTE_MARKER, 1)
    try:
        ast.parse(new_app)
    except SyntaxError as exc:
        raise SystemExit(f'ABORTADO: error de sintaxis: {exc}')
    backup(APP)
    APP.write_text(new_app, encoding='utf-8')
else:
    print('La ruta V6.9 ya existe; no se duplica.')

if 'admin_centro_jugadores' not in base_text:
    if MENU_MARKER not in base_text:
        raise SystemExit('ABORTADO: menu esperado no encontrado en base.html. No se modifico base.html.')
    backup(BASE)
    BASE.write_text(base_text.replace(MENU_MARKER, MENU_MARKER + '\n' + MENU_LINE, 1), encoding='utf-8')
else:
    print('El enlace V6.9 ya existe; no se duplica.')

shutil.copy2(SRC, DST)
print('V6.9 integrada correctamente.')
print('Ruta: /admin/jugadores/centro')
print('Prueba LOCAL antes de hacer commit o push.')
