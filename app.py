import os
from functools import wraps
from datetime import date, datetime, timedelta
from io import BytesIO

import qrcode

from flask import (
    Flask,
    render_template,
    request,
    redirect,
    url_for,
    flash,
    jsonify,
    Response,
    session
)

from flask_sqlalchemy import SQLAlchemy
from openpyxl import load_workbook

from docx import Document
from docx.enum.section import WD_SECTION
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_CELL_VERTICAL_ALIGNMENT
from docx.shared import Cm, Pt
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from werkzeug.security import generate_password_hash, check_password_hash


# ============================================================
# CONFIGURACIÓN
# ============================================================

app = Flask(__name__)

app.config["SECRET_KEY"] = os.environ.get(
    "SECRET_KEY",
    "cambia-esta-clave-en-produccion"
)


# ============================================================
# BASE DE DATOS
# ============================================================

db_url = os.environ.get("DATABASE_URL")

if db_url:

    db_url = db_url.replace(
        "postgres://",
        "postgresql+psycopg2://",
        1
    )

    db_url = db_url.replace(
        "postgresql://",
        "postgresql+psycopg2://",
        1
    )


app.config["SQLALCHEMY_DATABASE_URI"] = (
    db_url or "sqlite:///jugadores.db"
)

app.config["SQLALCHEMY_TRACK_MODIFICATIONS"] = False

db = SQLAlchemy(app)


# ============================================================
# ESTADOS PERMITIDOS
# ============================================================

ESTADOS_PERMITIDOS = {
    "Vigente",
    "Pendiente",
    "Suspendido",
    "Inhabilitado"
}


def normalizar_estado(estado):

    if not estado:
        return "Vigente"

    estado = str(estado).strip()

    if estado not in ESTADOS_PERMITIDOS:
        return "Vigente"

    return estado


# ============================================================
# V5.9 — USUARIOS ADMINISTRADORES
# ============================================================

class AdminUser(db.Model):

    id = db.Column(db.Integer, primary_key=True)
    username = db.Column(db.String(80), unique=True, nullable=False, index=True)
    nombre = db.Column(db.String(160), nullable=False, default="Administrador")
    password_hash = db.Column(db.String(255), nullable=False)
    activo = db.Column(db.Boolean, nullable=False, default=True)
    creado_en = db.Column(db.DateTime, nullable=False, default=datetime.utcnow)

    def set_password(self, password):
        self.password_hash = generate_password_hash(password)

    def check_password(self, password):
        return check_password_hash(self.password_hash, password)


# ============================================================
# MODELO CLUB
# ============================================================

class Club(db.Model):

    id = db.Column(
        db.Integer,
        primary_key=True
    )

    nombre = db.Column(
        db.String(120),
        unique=True,
        nullable=False
    )

    activo = db.Column(
        db.Boolean,
        nullable=False,
        default=True
    )

    campeonatos_participantes = db.relationship(
        "CampeonatoClub",
        back_populates="club",
        cascade="all, delete-orphan",
        lazy=True
    )


# ============================================================
# MODELO SERIE
# ============================================================

class Serie(db.Model):

    id = db.Column(
        db.Integer,
        primary_key=True
    )

    nombre = db.Column(
        db.String(80),
        unique=True,
        nullable=False
    )

    activo = db.Column(
        db.Boolean,
        nullable=False,
        default=True
    )


# ============================================================
# MODELO CAMPEONATO
# ============================================================

class Campeonato(db.Model):

    id = db.Column(
        db.Integer,
        primary_key=True
    )

    nombre = db.Column(
        db.String(160),
        nullable=False
    )

    temporada = db.Column(
        db.String(20),
        nullable=False
    )

    serie = db.Column(
        db.String(80),
        nullable=False
    )

    fecha_inicio = db.Column(
        db.Date,
        nullable=True
    )

    fecha_termino = db.Column(
        db.Date,
        nullable=True
    )

    estado = db.Column(
        db.String(30),
        nullable=False,
        default="Activo"
    )

    descripcion = db.Column(
        db.Text,
        nullable=True
    )

    clubes_participantes = db.relationship(
        "CampeonatoClub",
        back_populates="campeonato",
        cascade="all, delete-orphan",
        lazy=True
    )


# ============================================================
# MODELO CAMPEONATO - CLUB PARTICIPANTE
# ============================================================

class CampeonatoClub(db.Model):

    id = db.Column(
        db.Integer,
        primary_key=True
    )

    campeonato_id = db.Column(
        db.Integer,
        db.ForeignKey("campeonato.id"),
        nullable=False,
        index=True
    )

    club_id = db.Column(
        db.Integer,
        db.ForeignKey("club.id"),
        nullable=False,
        index=True
    )

    campeonato = db.relationship(
        "Campeonato",
        back_populates="clubes_participantes"
    )

    club = db.relationship(
        "Club",
        back_populates="campeonatos_participantes"
    )

    __table_args__ = (
        db.UniqueConstraint(
            "campeonato_id",
            "club_id",
            name="uq_campeonato_club"
        ),
    )


# ============================================================
# MODELO PARTIDO / FIXTURE
# ============================================================

class Partido(db.Model):

    id = db.Column(
        db.Integer,
        primary_key=True
    )

    campeonato_id = db.Column(
        db.Integer,
        db.ForeignKey("campeonato.id"),
        nullable=False,
        index=True
    )

    jornada = db.Column(
        db.Integer,
        nullable=False,
        index=True
    )

    fecha = db.Column(
        db.Date,
        nullable=True
    )

    hora = db.Column(
        db.String(20),
        nullable=True
    )

    cancha = db.Column(
        db.String(120),
        nullable=True
    )

    local_club_id = db.Column(
        db.Integer,
        db.ForeignKey("club.id"),
        nullable=False,
        index=True
    )

    visitante_club_id = db.Column(
        db.Integer,
        db.ForeignKey("club.id"),
        nullable=False,
        index=True
    )

    turno_club_id = db.Column(
        db.Integer,
        db.ForeignKey("club.id"),
        nullable=True,
        index=True
    )

    goles_local = db.Column(
        db.Integer,
        nullable=True
    )

    goles_visitante = db.Column(
        db.Integer,
        nullable=True
    )

    estado = db.Column(
        db.String(30),
        nullable=False,
        default="Programado"
    )

    campeonato = db.relationship(
        "Campeonato",
        backref=db.backref(
            "partidos",
            lazy=True,
            cascade="all, delete-orphan"
        )
    )

    local_club = db.relationship(
        "Club",
        foreign_keys=[local_club_id],
        backref=db.backref("partidos_local", lazy=True)
    )

    visitante_club = db.relationship(
        "Club",
        foreign_keys=[visitante_club_id],
        backref=db.backref("partidos_visitante", lazy=True)
    )

    turno_club = db.relationship(
        "Club",
        foreign_keys=[turno_club_id],
        backref=db.backref("partidos_turno", lazy=True)
    )


# ============================================================
# MODELO JUGADOR
# ============================================================

class Jugador(db.Model):

    id = db.Column(
        db.Integer,
        primary_key=True
    )

    rut = db.Column(
        db.String(20),
        unique=True,
        nullable=False,
        index=True
    )

    nombre_completo = db.Column(
        db.String(160),
        nullable=False
    )

    fecha_nacimiento = db.Column(
        db.Date,
        nullable=False
    )

    serie = db.Column(
        db.String(80),
        nullable=False
    )

    club = db.Column(
        db.String(120),
        nullable=False
    )

    foto = db.Column(
        db.LargeBinary,
        nullable=True
    )

    estado = db.Column(
        db.String(30),
        nullable=False,
        default="Vigente"
    )


# ============================================================
# MODELO REGISTRO DISCIPLINARIO
# ============================================================

class RegistroDisciplinario(db.Model):

    id = db.Column(
        db.Integer,
        primary_key=True
    )

    jugador_id = db.Column(
        db.Integer,
        db.ForeignKey("jugador.id"),
        nullable=False,
        index=True
    )

    fecha = db.Column(
        db.Date,
        nullable=False,
        default=date.today
    )

    tipo = db.Column(
        db.String(30),
        nullable=False
    )

    cantidad = db.Column(
        db.Integer,
        nullable=False,
        default=1
    )

    motivo = db.Column(
        db.String(255),
        nullable=True
    )

    campeonato = db.Column(
        db.String(120),
        nullable=True
    )

    observaciones = db.Column(
        db.Text,
        nullable=True
    )

    campeonato_id = db.Column(
        db.Integer,
        db.ForeignKey("campeonato.id"),
        nullable=True,
        index=True
    )

    jugador = db.relationship(
        "Jugador",
        backref=db.backref(
            "registros_disciplinarios",
            lazy=True,
            cascade="all, delete-orphan"
        )
    )


# ============================================================
# MODELO GOLES
# ============================================================

class Gol(db.Model):

    id = db.Column(
        db.Integer,
        primary_key=True
    )

    jugador_id = db.Column(
        db.Integer,
        db.ForeignKey("jugador.id"),
        nullable=False,
        index=True
    )

    fecha = db.Column(
        db.Date,
        nullable=False,
        default=date.today
    )

    cantidad = db.Column(
        db.Integer,
        nullable=False,
        default=1
    )

    campeonato = db.Column(
        db.String(120),
        nullable=True
    )

    observaciones = db.Column(
        db.Text,
        nullable=True
    )

    campeonato_id = db.Column(
        db.Integer,
        db.ForeignKey("campeonato.id"),
        nullable=True,
        index=True
    )

    jugador = db.relationship(
        "Jugador",
        backref=db.backref(
            "goles_registrados",
            lazy=True,
            cascade="all, delete-orphan"
        )
    )


# ============================================================
# V5.7 — ACTA DIGITAL / NÓMINA DE PARTIDO
# ============================================================

class ActaPartido(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    partido_id = db.Column(db.Integer, db.ForeignKey("partido.id"), nullable=False, unique=True, index=True)
    numero_acta = db.Column(db.String(40), nullable=True)
    arbitro = db.Column(db.String(160), nullable=True)
    observaciones = db.Column(db.Text, nullable=True)
    estado = db.Column(db.String(30), nullable=False, default="Borrador")
    partido = db.relationship("Partido", backref=db.backref("acta", uselist=False, cascade="all, delete-orphan"))

class PartidoJugador(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    partido_id = db.Column(db.Integer, db.ForeignKey("partido.id"), nullable=False, index=True)
    jugador_id = db.Column(db.Integer, db.ForeignKey("jugador.id"), nullable=False, index=True)
    equipo = db.Column(db.String(20), nullable=False)
    condicion = db.Column(db.String(20), nullable=False, default="Suplente")
    capitan = db.Column(db.Boolean, nullable=False, default=False)
    ingreso = db.Column(db.String(10), nullable=True)
    salida = db.Column(db.String(10), nullable=True)
    goles = db.Column(db.Integer, nullable=False, default=0)
    amarillas = db.Column(db.Integer, nullable=False, default=0)
    rojas = db.Column(db.Integer, nullable=False, default=0)
    observaciones = db.Column(db.String(255), nullable=True)
    partido = db.relationship("Partido", backref=db.backref("nomina", lazy=True, cascade="all, delete-orphan"))
    jugador = db.relationship("Jugador", backref=db.backref("participaciones_partidos", lazy=True))
    __table_args__ = (db.UniqueConstraint("partido_id", "jugador_id", name="uq_partido_jugador"),)


# ============================================================
# CREACIÓN / ACTUALIZACIÓN SEGURA DE BASE DE DATOS
# ============================================================

def preparar_base_datos():

    db.create_all()

    try:

        inspector = db.inspect(
            db.engine
        )

        columnas = [
            columna["name"]
            for columna in inspector.get_columns(
                "jugador"
            )
        ]

        # ----------------------------------------------------
        # AGREGAR FOTO SI NO EXISTE
        # ----------------------------------------------------

        if "foto" not in columnas:

            if db.engine.dialect.name == "postgresql":

                db.session.execute(
                    db.text(
                        "ALTER TABLE jugador "
                        "ADD COLUMN IF NOT EXISTS foto BYTEA"
                    )
                )

            elif db.engine.dialect.name == "sqlite":

                db.session.execute(
                    db.text(
                        "ALTER TABLE jugador "
                        "ADD COLUMN foto BLOB"
                    )
                )

            db.session.commit()

        # ----------------------------------------------------
        # AGREGAR ESTADO SI NO EXISTE
        # ----------------------------------------------------

        inspector = db.inspect(
            db.engine
        )

        columnas = [
            columna["name"]
            for columna in inspector.get_columns(
                "jugador"
            )
        ]

        if "estado" not in columnas:

            if db.engine.dialect.name == "postgresql":

                db.session.execute(
                    db.text(
                        "ALTER TABLE jugador "
                        "ADD COLUMN IF NOT EXISTS "
                        "estado VARCHAR(30) "
                        "DEFAULT 'Vigente'"
                    )
                )

            elif db.engine.dialect.name == "sqlite":

                db.session.execute(
                    db.text(
                        "ALTER TABLE jugador "
                        "ADD COLUMN estado VARCHAR(30) "
                        "DEFAULT 'Vigente'"
                    )
                )

            db.session.commit()

        # ----------------------------------------------------
        # ASEGURAR ESTADO EN REGISTROS ANTIGUOS
        # ----------------------------------------------------

        db.session.execute(
            db.text(
                "UPDATE jugador "
                "SET estado = 'Vigente' "
                "WHERE estado IS NULL "
                "OR estado = ''"
            )
        )

        db.session.commit()

        # ----------------------------------------------------
        # ASEGURAR COLUMNA OBSERVACIONES EN DISCIPLINA
        # ----------------------------------------------------

        inspector = db.inspect(
            db.engine
        )

        columnas_disciplina = [
            columna["name"]
            for columna in inspector.get_columns(
                "registro_disciplinario"
            )
        ]

        if "observaciones" not in columnas_disciplina:

            if db.engine.dialect.name == "postgresql":

                db.session.execute(
                    db.text(
                        "ALTER TABLE registro_disciplinario "
                        "ADD COLUMN IF NOT EXISTS "
                        "observaciones TEXT"
                    )
                )

            elif db.engine.dialect.name == "sqlite":

                db.session.execute(
                    db.text(
                        "ALTER TABLE registro_disciplinario "
                        "ADD COLUMN observaciones TEXT"
                    )
                )

            db.session.commit()

    except Exception as error:

        db.session.rollback()

        print(
            "Advertencia al preparar la base de datos:",
            error
        )


def preparar_vinculos_campeonato():
    """Migra de forma segura las columnas nuevas usadas por PASO 8.

    db.create_all() no agrega columnas a tablas que ya existen en Railway,
    por eso aquí se revisa cada tabla y se agrega solamente lo que falta.
    """
    try:
        inspector = db.inspect(db.engine)
        dialecto = db.engine.dialect.name

        columnas_nuevas = {
            "gol": {
                "campeonato": "VARCHAR(120)",
                "observaciones": "TEXT",
                "campeonato_id": "INTEGER",
            },
            "registro_disciplinario": {
                "campeonato_id": "INTEGER",
            },
            "partido": {
                "turno_club_id": "INTEGER",
            },
        }

        for tabla, definiciones in columnas_nuevas.items():
            columnas_actuales = {
                c["name"] for c in inspector.get_columns(tabla)
            }

            for nombre, tipo in definiciones.items():
                if nombre in columnas_actuales:
                    continue

                if dialecto == "postgresql":
                    sql = (
                        f"ALTER TABLE {tabla} "
                        f"ADD COLUMN IF NOT EXISTS {nombre} {tipo}"
                    )
                elif dialecto == "sqlite":
                    sql = (
                        f"ALTER TABLE {tabla} "
                        f"ADD COLUMN {nombre} {tipo}"
                    )
                else:
                    # Para otros motores dejamos que SQLAlchemy reporte el
                    # problema en vez de ejecutar SQL incompatible.
                    raise RuntimeError(
                        f"Motor de base de datos no soportado para migración: {dialecto}"
                    )

                print(f"Agregando columna {tabla}.{nombre}...")
                db.session.execute(db.text(sql))
                db.session.commit()

                # Refrescar el inspector después de cada ALTER TABLE.
                inspector = db.inspect(db.engine)

    except Exception as error:
        db.session.rollback()
        print(
            "ERROR preparando columnas de estadísticas del campeonato:",
            repr(error)
        )

def asegurar_admin_inicial():
    """Crea un administrador inicial solo si no existe ninguno."""
    try:
        if AdminUser.query.count() > 0:
            return
        username = (os.environ.get("ADMIN_USERNAME") or "admin").strip()
        password = os.environ.get("ADMIN_PASSWORD") or "PresidenteRios2026!"
        nombre = (os.environ.get("ADMIN_NAME") or "Administrador principal").strip()
        admin = AdminUser(username=username, nombre=nombre, activo=True)
        admin.set_password(password)
        db.session.add(admin)
        db.session.commit()
        print(f"Administrador inicial creado: {username}")
    except Exception as error:
        db.session.rollback()
        print("ERROR creando administrador inicial:", repr(error))


with app.app_context():
    preparar_base_datos()
    preparar_vinculos_campeonato()
    asegurar_admin_inicial()


# ============================================================
# SEGURIDAD V5.9 — PORTAL PÚBLICO / ADMINISTRACIÓN
# ============================================================

PUBLIC_ENDPOINTS = {
    "login",
    "logout",
    "publico",
    "publico_campeonato",
    "publico_tabla",
    "publico_goleadores",
    "health",
    "static",
}


@app.before_request
def exigir_login_administrativo():
    endpoint = request.endpoint
    if endpoint in PUBLIC_ENDPOINTS or (request.path or "").startswith("/static/"):
        return None
    if session.get("admin_id"):
        return None
    destino = request.full_path.rstrip("?")
    return redirect(url_for("login", next=destino))


def admin_required(view):
    @wraps(view)
    def wrapped(*args, **kwargs):
        if not session.get("admin_id"):
            return redirect(url_for("login", next=request.full_path))
        return view(*args, **kwargs)
    return wrapped


# ============================================================
# FUNCIONES AUXILIARES
# ============================================================

def normalizar_rut(rut):

    if rut is None:
        return ""

    rut = str(rut).strip().upper()

    rut = rut.replace(
        " ",
        ""
    )

    rut_limpio = (
        rut
        .replace(".", "")
        .replace("-", "")
    )

    if len(rut_limpio) < 2:
        return ""

    cuerpo = rut_limpio[:-1]

    dv = rut_limpio[-1]

    if not cuerpo.isdigit():
        return ""

    cuerpo_formateado = ""

    while len(cuerpo) > 3:

        cuerpo_formateado = (
            "."
            + cuerpo[-3:]
            + cuerpo_formateado
        )

        cuerpo = cuerpo[:-3]

    cuerpo_formateado = (
        cuerpo
        + cuerpo_formateado
    )

    return f"{cuerpo_formateado}-{dv}"


def validar_rut(rut):

    if not rut:
        return False

    rut_limpio = (
        rut
        .replace(".", "")
        .replace("-", "")
        .replace(" ", "")
        .upper()
    )

    if len(rut_limpio) < 2:
        return False

    cuerpo = rut_limpio[:-1]

    dv = rut_limpio[-1]

    if not cuerpo.isdigit():
        return False

    suma = 0

    multiplicador = 2

    for digito in reversed(cuerpo):

        suma += (
            int(digito)
            * multiplicador
        )

        multiplicador += 1

        if multiplicador > 7:
            multiplicador = 2

    resto = suma % 11

    resultado = 11 - resto

    if resultado == 11:

        dv_calculado = "0"

    elif resultado == 10:

        dv_calculado = "K"

    else:

        dv_calculado = str(resultado)

    return dv == dv_calculado


def convertir_fecha(valor):

    if valor is None:
        return None

    if isinstance(valor, datetime):

        return valor.date()

    if isinstance(valor, date):

        return valor

    if isinstance(valor, str):

        valor = valor.strip()

        formatos = [
            "%d/%m/%Y",
            "%d-%m-%Y",
            "%Y-%m-%d",
            "%d.%m.%Y"
        ]

        for formato in formatos:

            try:

                return datetime.strptime(
                    valor,
                    formato
                ).date()

            except ValueError:

                continue

    return None


def obtener_fotografia():

    archivo = request.files.get(
        "foto"
    )

    if not archivo:
        return None, None

    if not archivo.filename:
        return None, None

    contenido = archivo.read()

    if not contenido:

        return (
            None,
            "La fotografía seleccionada está vacía."
        )

    maximo = 5 * 1024 * 1024

    if len(contenido) > maximo:

        return (
            None,
            "La fotografía no puede superar los 5 MB."
        )

    tipo = (
        archivo.mimetype or ""
    ).lower()

    tipos_permitidos = {
        "image/jpeg",
        "image/png",
        "image/webp"
    }

    if tipo not in tipos_permitidos:

        return (
            None,
            "La fotografía debe ser JPG, PNG o WEBP."
        )

    return contenido, None


# ============================================================
# ESTADÍSTICAS DEPORTIVAS
# ============================================================

def obtener_amarillas(jugador_id):

    try:
        registros = (
            RegistroDisciplinario.query
            .filter_by(
                jugador_id=jugador_id,
                tipo="Amarilla"
            )
            .all()
        )

        return sum(
            registro.cantidad or 0
            for registro in registros
        )

    except Exception as error:

        db.session.rollback()
        print("Advertencia obteniendo amarillas:", error)
        return 0


def obtener_rojas(jugador_id):

    try:
        registros = (
            RegistroDisciplinario.query
            .filter_by(
                jugador_id=jugador_id,
                tipo="Roja"
            )
            .all()
        )

        return sum(
            registro.cantidad or 0
            for registro in registros
        )

    except Exception as error:

        db.session.rollback()
        print("Advertencia obteniendo rojas:", error)
        return 0


def obtener_goles(jugador_id):

    try:
        registros = (
            Gol.query
            .filter_by(
                jugador_id=jugador_id
            )
            .all()
        )

        return sum(
            registro.cantidad or 0
            for registro in registros
        )

    except Exception as error:

        db.session.rollback()
        print("Advertencia obteniendo goles:", error)
        return 0


def obtener_suspensiones(jugador_id):

    try:
        registros = (
            RegistroDisciplinario.query
            .filter_by(
                jugador_id=jugador_id,
                tipo="Suspension"
            )
            .all()
        )

        return sum(
            registro.cantidad or 0
            for registro in registros
        )

    except Exception as error:

        db.session.rollback()
        print("Advertencia obteniendo suspensiones:", error)
        return 0


def crear_suspension_por_acumulacion(
    jugador,
    campeonato=None
):

    amarillas = obtener_amarillas(
        jugador.id
    )

    suspensiones_correspondientes = (
        amarillas // 4
    )

    suspensiones_existentes = (
        obtener_suspensiones(
            jugador.id
        )
    )

    nuevas_suspensiones = (
        suspensiones_correspondientes
        - suspensiones_existentes
    )

    if nuevas_suspensiones <= 0:
        return 0

    jugador.estado = "Suspendido"

    for _ in range(
        nuevas_suspensiones
    ):

        suspension = RegistroDisciplinario(

            jugador_id=jugador.id,

            fecha=date.today(),

            tipo="Suspension",

            cantidad=1,

            motivo=(
                "Suspensión automática "
                "por acumulación de 4 "
                "tarjetas amarillas."
            ),

            campeonato=campeonato
        )

        db.session.add(
            suspension
        )

    return nuevas_suspensiones


# ============================================================
# DATOS PARA FORMULARIO DE JUGADORES
# ============================================================

def obtener_datos_formulario_jugador():

    clubes = Club.query.filter_by(
        activo=True
    ).order_by(
        Club.nombre
    ).all()

    series = Serie.query.filter_by(
        activo=True
    ).order_by(
        Serie.nombre
    ).all()

    return clubes, series


# ============================================================
# V5.9 — LOGIN DE ADMINISTRADORES
# ============================================================

@app.route("/login", methods=["GET", "POST"])
def login():
    if session.get("admin_id"):
        return redirect(url_for("dashboard"))

    if request.method == "POST":
        username = request.form.get("username", "").strip()
        password = request.form.get("password", "")
        admin = AdminUser.query.filter_by(username=username).first()

        if admin and admin.activo and admin.check_password(password):
            session.clear()
            session["admin_id"] = admin.id
            session["admin_username"] = admin.username
            session["admin_nombre"] = admin.nombre
            session.permanent = True
            destino = request.form.get("next", "").strip()
            if not destino.startswith("/") or destino.startswith("//"):
                destino = url_for("dashboard")
            return redirect(destino)

        flash("Usuario o contraseña incorrectos, o el administrador está inactivo.", "error")

    return render_template("login.html", next=request.args.get("next", ""))


@app.route("/logout")
def logout():
    session.clear()
    flash("Sesión cerrada correctamente.", "success")
    return redirect(url_for("login"))


@app.route("/admin/mi-cuenta", methods=["GET", "POST"])
def mi_cuenta_admin():
    admin = db.get_or_404(AdminUser, session["admin_id"])
    if request.method == "POST":
        actual = request.form.get("password_actual", "")
        nueva = request.form.get("password_nueva", "")
        confirmar = request.form.get("password_confirmar", "")
        if not admin.check_password(actual):
            flash("La contraseña actual no es correcta.", "error")
        elif len(nueva) < 8:
            flash("La nueva contraseña debe tener al menos 8 caracteres.", "error")
        elif nueva != confirmar:
            flash("Las contraseñas nuevas no coinciden.", "error")
        else:
            admin.set_password(nueva)
            db.session.commit()
            flash("Contraseña actualizada correctamente.", "success")
            return redirect(url_for("mi_cuenta_admin"))
    return render_template("admin_cuenta.html", admin=admin)


@app.route("/admin/usuarios", methods=["GET", "POST"])
def admin_usuarios():
    if request.method == "POST":
        username = request.form.get("username", "").strip()
        nombre = request.form.get("nombre", "").strip() or "Administrador"
        password = request.form.get("password", "")
        if len(username) < 3 or len(password) < 8:
            flash("El usuario debe tener al menos 3 caracteres y la contraseña 8.", "error")
        elif AdminUser.query.filter_by(username=username).first():
            flash("Ese usuario administrador ya existe.", "error")
        else:
            admin = AdminUser(username=username, nombre=nombre, activo=True)
            admin.set_password(password)
            db.session.add(admin)
            db.session.commit()
            flash("Administrador creado correctamente.", "success")
            return redirect(url_for("admin_usuarios"))
    usuarios = AdminUser.query.order_by(AdminUser.username).all()
    return render_template("admin_usuarios.html", usuarios=usuarios)


@app.route("/admin/usuarios/<int:admin_id>/estado", methods=["POST"])
def cambiar_estado_admin(admin_id):
    admin = db.get_or_404(AdminUser, admin_id)
    if admin.id == session.get("admin_id"):
        flash("No puedes desactivar tu propia cuenta.", "error")
    else:
        admin.activo = not admin.activo
        db.session.commit()
        flash("Estado del administrador actualizado.", "success")
    return redirect(url_for("admin_usuarios"))


# ============================================================
# INICIO / LISTADO
# ============================================================

@app.route("/")
def index():

    q = request.args.get(
        "q",
        ""
    ).strip()

    query = Jugador.query

    if q:

        query = query.filter(
            db.or_(
                Jugador.rut.ilike(
                    f"%{q}%"
                ),

                Jugador.nombre_completo.ilike(
                    f"%{q}%"
                ),

                Jugador.club.ilike(
                    f"%{q}%"
                )
            )
        )

    jugadores = query.order_by(
        Jugador.nombre_completo
    ).all()

    return render_template(
        "index.html",
        jugadores=jugadores,
        q=q
    )




# ============================================================
# FICHA INDIVIDUAL
# ============================================================

@app.route(
    "/jugadores/<int:jugador_id>"
)
def ficha_jugador(jugador_id):

    jugador = db.get_or_404(
        Jugador,
        jugador_id
    )

    goles = obtener_goles(
        jugador.id
    )

    amarillas = obtener_amarillas(
        jugador.id
    )

    rojas = obtener_rojas(
        jugador.id
    )

    suspensiones = obtener_suspensiones(
        jugador.id
    )

    try:
        historial = (
            RegistroDisciplinario.query
            .filter_by(
                jugador_id=jugador.id
            )
            .order_by(
                RegistroDisciplinario.fecha.desc(),
                RegistroDisciplinario.id.desc()
            )
            .all()
        )
    except Exception as error:
        db.session.rollback()
        print("Advertencia obteniendo historial disciplinario:", error)
        historial = []

    try:
        historial_goles = (
            Gol.query
            .filter_by(
                jugador_id=jugador.id
            )
            .order_by(
                Gol.fecha.desc(),
                Gol.id.desc()
            )
            .all()
        )
    except Exception as error:
        db.session.rollback()
        print("Advertencia obteniendo historial de goles:", error)
        historial_goles = []

    return render_template(
        "jugador_detalle.html",
        jugador=jugador,
        goles=goles,
        amarillas=amarillas,
        rojas=rojas,
        suspensiones=suspensiones,
        historial=historial,
        historial_goles=historial_goles
    )


# ============================================================
# V5.9 — HISTORIAL DEPORTIVO DEL JUGADOR
# ============================================================

@app.route("/jugadores/<int:jugador_id>/historial")
def historial_jugador(jugador_id):
    jugador = db.get_or_404(Jugador, jugador_id)
    participaciones = (
        PartidoJugador.query
        .join(Partido, PartidoJugador.partido_id == Partido.id)
        .join(Campeonato, Partido.campeonato_id == Campeonato.id)
        .filter(PartidoJugador.jugador_id == jugador.id)
        .order_by(Partido.fecha.desc().nullslast(), Partido.id.desc())
        .all()
    )

    total_partidos = len(participaciones)
    total_titular = sum(1 for p in participaciones if p.condicion == "Titular")
    total_suplente = sum(1 for p in participaciones if p.condicion == "Suplente")
    total_goles = sum(int(p.goles or 0) for p in participaciones)
    total_amarillas = sum(int(p.amarillas or 0) for p in participaciones)
    total_rojas = sum(int(p.rojas or 0) for p in participaciones)

    por_campeonato = {}
    for p in participaciones:
        camp = p.partido.campeonato
        fila = por_campeonato.setdefault(camp.id, {
            "campeonato": camp, "partidos": 0, "titular": 0, "suplente": 0,
            "goles": 0, "amarillas": 0, "rojas": 0
        })
        fila["partidos"] += 1
        fila["titular"] += 1 if p.condicion == "Titular" else 0
        fila["suplente"] += 1 if p.condicion == "Suplente" else 0
        fila["goles"] += int(p.goles or 0)
        fila["amarillas"] += int(p.amarillas or 0)
        fila["rojas"] += int(p.rojas or 0)

    return render_template(
        "jugador_historial.html", jugador=jugador, participaciones=participaciones,
        campeonatos_historial=list(por_campeonato.values()), total_partidos=total_partidos,
        total_titular=total_titular, total_suplente=total_suplente, total_goles=total_goles,
        total_amarillas=total_amarillas, total_rojas=total_rojas
    )


# ============================================================
# FOTOGRAFÍA
# ============================================================

@app.route(
    "/jugadores/<int:jugador_id>/foto"
)
def foto_jugador(jugador_id):

    jugador = db.get_or_404(
        Jugador,
        jugador_id
    )

    if not jugador.foto:

        return (
            "",
            404
        )

    return Response(
        jugador.foto,
        mimetype="image/jpeg"
    )


# ============================================================
# NUEVO JUGADOR
# ============================================================

@app.route(
    "/jugadores/nuevo",
    methods=["GET", "POST"]
)
def nuevo_jugador():

    clubes, series = obtener_datos_formulario_jugador()

    if request.method == "POST":

        rut = normalizar_rut(
            request.form.get(
                "rut",
                ""
            )
        )

        nombre = request.form.get(
            "nombre_completo",
            ""
        ).strip()

        fecha = request.form.get(
            "fecha_nacimiento",
            ""
        ).strip()

        serie = request.form.get(
            "serie",
            ""
        ).strip()

        club = request.form.get(
            "club",
            ""
        ).strip()

        estado = normalizar_estado(
            request.form.get(
                "estado",
                "Vigente"
            )
        )

        if not all([
            rut,
            nombre,
            fecha,
            serie,
            club
        ]):

            flash(
                "Completa todos los campos.",
                "error"
            )

            return render_template(
                "jugador_form.html",
                jugador=None,
                clubes=clubes,
                series=series
            )

        if not validar_rut(rut):

            flash(
                "El RUT ingresado no es válido.",
                "error"
            )

            return render_template(
                "jugador_form.html",
                jugador=None,
                clubes=clubes,
                series=series
            )

        try:

            fecha_obj = date.fromisoformat(
                fecha
            )

        except ValueError:

            flash(
                "Fecha no válida.",
                "error"
            )

            return render_template(
                "jugador_form.html",
                jugador=None,
                clubes=clubes,
                series=series
            )

        if Jugador.query.filter_by(
            rut=rut
        ).first():

            flash(
                "Ese RUT ya está registrado.",
                "error"
            )

            return render_template(
                "jugador_form.html",
                jugador=None,
                clubes=clubes,
                series=series
            )

        foto, error_foto = obtener_fotografia()

        if error_foto:

            flash(
                error_foto,
                "error"
            )

            return render_template(
                "jugador_form.html",
                jugador=None,
                clubes=clubes,
                series=series
            )

        jugador = Jugador(
            rut=rut,
            nombre_completo=nombre,
            fecha_nacimiento=fecha_obj,
            serie=serie,
            club=club,
            foto=foto,
            estado=estado
        )

        try:

            db.session.add(
                jugador
            )

            db.session.commit()

        except Exception as error:

            db.session.rollback()

            print(
                "Error registrando jugador:",
                error
            )

            flash(
                "No fue posible guardar el jugador.",
                "error"
            )

            return render_template(
                "jugador_form.html",
                jugador=None,
                clubes=clubes,
                series=series
            )

        flash(
            "Jugador registrado correctamente.",
            "success"
        )

        return redirect(
            url_for(
                "ficha_jugador",
                jugador_id=jugador.id
            )
        )

    return render_template(
        "jugador_form.html",
        jugador=None,
        clubes=clubes,
        series=series
    )


# ============================================================
# EDITAR JUGADOR
# ============================================================

@app.route(
    "/jugadores/<int:jugador_id>/editar",
    methods=["GET", "POST"]
)
def editar_jugador(jugador_id):

    jugador = db.get_or_404(
        Jugador,
        jugador_id
    )

    clubes, series = obtener_datos_formulario_jugador()

    if request.method == "POST":

        rut = normalizar_rut(
            request.form.get(
                "rut",
                ""
            )
        )

        nombre = request.form.get(
            "nombre_completo",
            ""
        ).strip()

        fecha = request.form.get(
            "fecha_nacimiento",
            ""
        ).strip()

        serie = request.form.get(
            "serie",
            ""
        ).strip()

        club = request.form.get(
            "club",
            ""
        ).strip()

        estado = normalizar_estado(
            request.form.get(
                "estado",
                jugador.estado or "Vigente"
            )
        )

        if not all([
            rut,
            nombre,
            fecha,
            serie,
            club
        ]):

            flash(
                "Completa todos los campos.",
                "error"
            )

            return render_template(
                "jugador_form.html",
                jugador=jugador,
                clubes=clubes,
                series=series
            )

        if not validar_rut(rut):

            flash(
                "El RUT ingresado no es válido.",
                "error"
            )

            return render_template(
                "jugador_form.html",
                jugador=jugador,
                clubes=clubes,
                series=series
            )

        otro_jugador = Jugador.query.filter(
            Jugador.rut == rut,
            Jugador.id != jugador.id
        ).first()

        if otro_jugador:

            flash(
                "Ese RUT ya pertenece a otro jugador.",
                "error"
            )

            return render_template(
                "jugador_form.html",
                jugador=jugador,
                clubes=clubes,
                series=series
            )

        try:

            fecha_obj = date.fromisoformat(
                fecha
            )

        except ValueError:

            flash(
                "Fecha no válida.",
                "error"
            )

            return render_template(
                "jugador_form.html",
                jugador=jugador,
                clubes=clubes,
                series=series
            )

        archivo_foto = request.files.get(
            "foto"
        )

        if (
            archivo_foto
            and archivo_foto.filename
        ):

            foto, error_foto = obtener_fotografia()

            if error_foto:

                flash(
                    error_foto,
                    "error"
                )

                return render_template(
                    "jugador_form.html",
                    jugador=jugador,
                    clubes=clubes,
                    series=series
                )

            jugador.foto = foto

        jugador.rut = rut

        jugador.nombre_completo = nombre

        jugador.fecha_nacimiento = fecha_obj

        jugador.serie = serie

        jugador.club = club

        jugador.estado = estado

        try:

            db.session.commit()

        except Exception as error:

            db.session.rollback()

            print(
                "Error actualizando jugador:",
                error
            )

            flash(
                "No fue posible actualizar el jugador.",
                "error"
            )

            return render_template(
                "jugador_form.html",
                jugador=jugador,
                clubes=clubes,
                series=series
            )

        flash(
            "Datos actualizados correctamente.",
            "success"
        )

        return redirect(
            url_for(
                "ficha_jugador",
                jugador_id=jugador.id
            )
        )

    return render_template(
        "jugador_form.html",
        jugador=jugador,
        clubes=clubes,
        series=series
    )


# ============================================================
# ELIMINAR JUGADOR
# ============================================================

@app.route(
    "/jugadores/<int:jugador_id>/eliminar",
    methods=["POST"]
)
def eliminar_jugador(jugador_id):

    jugador = db.get_or_404(
        Jugador,
        jugador_id
    )

    try:

        db.session.delete(
            jugador
        )

        db.session.commit()

    except Exception as error:

        db.session.rollback()

        print(
            "Error eliminando jugador:",
            error
        )

        flash(
            "No fue posible eliminar el jugador.",
            "error"
        )

        return redirect(
            url_for(
                "ficha_jugador",
                jugador_id=jugador_id
            )
        )

    flash(
        "Jugador eliminado.",
        "success"
    )

    return redirect(
        url_for("index")
    )


# ============================================================
# IMPORTAR EXCEL
# ============================================================

@app.route(
    "/jugadores/importar",
    methods=["GET", "POST"]
)
def importar_jugadores():

    if request.method == "GET":

        return render_template(
            "importar.html"
        )

    archivo = request.files.get(
        "archivo"
    )

    if not archivo or not archivo.filename:

        flash(
            "Selecciona un archivo Excel.",
            "error"
        )

        return redirect(
            url_for("importar_jugadores")
        )

    extension = (
        archivo.filename
        .rsplit(".", 1)[-1]
        .lower()
    )

    if extension != "xlsx":

        flash(
            "El archivo debe ser Excel (.xlsx).",
            "error"
        )

        return redirect(
            url_for("importar_jugadores")
        )

    try:

        workbook = load_workbook(
            archivo,
            data_only=True
        )

        hoja = workbook.active

        filas = list(
            hoja.iter_rows(
                values_only=True
            )
        )

        if not filas:

            flash(
                "El archivo está vacío.",
                "error"
            )

            return redirect(
                url_for("importar_jugadores")
            )

        encabezados = [
            str(x).strip().lower()
            if x is not None else ""
            for x in filas[0]
        ]

        columnas = {}

        equivalencias = {

            "rut": [
                "rut",
                "r.u.t.",
                "r.u.t",
                "documento"
            ],

            "nombre": [
                "nombre",
                "nombre completo",
                "nombre_completo"
            ],

            "fecha": [
                "fecha de nacimiento",
                "fecha nacimiento",
                "nacimiento",
                "fecha_nacimiento"
            ],

            "serie": [
                "serie",
                "categoría",
                "categoria"
            ],

            "club": [
                "club",
                "equipo"
            ]

        }

        for nombre_columna, posibles in (
            equivalencias.items()
        ):

            for posible in posibles:

                if posible in encabezados:

                    columnas[
                        nombre_columna
                    ] = encabezados.index(
                        posible
                    )

                    break

        faltantes = [
            campo
            for campo in equivalencias
            if campo not in columnas
        ]

        if faltantes:

            flash(
                "Faltan columnas obligatorias: "
                + ", ".join(faltantes),
                "error"
            )

            return redirect(
                url_for(
                    "importar_jugadores"
                )
            )

        registrados = 0

        duplicados = 0

        errores = 0

        detalle_errores = []

        ruts_archivo = set()

        for numero_fila, fila in enumerate(
            filas[1:],
            start=2
        ):

            try:

                rut_original = fila[
                    columnas["rut"]
                ]

                nombre = fila[
                    columnas["nombre"]
                ]

                fecha_valor = fila[
                    columnas["fecha"]
                ]

                serie = fila[
                    columnas["serie"]
                ]

                club = fila[
                    columnas["club"]
                ]

                rut = normalizar_rut(
                    rut_original
                )

                nombre = (
                    str(nombre).strip()
                    if nombre is not None
                    else ""
                )

                serie = (
                    str(serie).strip()
                    if serie is not None
                    else ""
                )

                club = (
                    str(club).strip()
                    if club is not None
                    else ""
                )

                if not all([
                    rut,
                    nombre,
                    fecha_valor,
                    serie,
                    club
                ]):

                    errores += 1

                    detalle_errores.append(
                        f"Fila {numero_fila}: "
                        "faltan datos obligatorios."
                    )

                    continue

                if not validar_rut(rut):

                    errores += 1

                    detalle_errores.append(
                        f"Fila {numero_fila}: "
                        f"RUT inválido ({rut})."
                    )

                    continue

                if rut in ruts_archivo:

                    duplicados += 1

                    continue

                ruts_archivo.add(rut)

                if Jugador.query.filter_by(
                    rut=rut
                ).first():

                    duplicados += 1

                    continue

                fecha_obj = convertir_fecha(
                    fecha_valor
                )

                if not fecha_obj:

                    errores += 1

                    detalle_errores.append(
                        f"Fila {numero_fila}: "
                        "fecha de nacimiento inválida."
                    )

                    continue

                jugador = Jugador(
                    rut=rut,
                    nombre_completo=nombre,
                    fecha_nacimiento=fecha_obj,
                    serie=serie,
                    club=club,
                    estado="Vigente"
                )

                db.session.add(
                    jugador
                )

                registrados += 1

            except Exception as error:

                errores += 1

                detalle_errores.append(
                    f"Fila {numero_fila}: "
                    f"error al procesar: {error}"
                )

        try:

            db.session.commit()

        except Exception:

            db.session.rollback()

            flash(
                "Ocurrió un error al guardar "
                "los jugadores.",
                "error"
            )

            return redirect(
                url_for(
                    "importar_jugadores"
                )
            )

        return render_template(
            "importar_resultado.html",
            registrados=registrados,
            duplicados=duplicados,
            errores=errores,
            detalle_errores=detalle_errores
        )

    except Exception as error:

        db.session.rollback()

        print(
            "Error importando Excel:",
            error
        )

        flash(
            "Ocurrió un error inesperado "
            "al importar el archivo.",
            "error"
        )

        return redirect(
            url_for(
                "importar_jugadores"
            )
        )


# ============================================================
# API DE JUGADORES
# ============================================================

@app.route(
    "/api/jugadores"
)
def api_jugadores():

    rut = request.args.get(
        "rut",
        ""
    ).strip().upper()

    if not rut:

        return jsonify([])

    jugadores = Jugador.query.filter(
        Jugador.rut.ilike(
            f"%{rut}%"
        )
    ).order_by(
        Jugador.nombre_completo
    ).all()

    return jsonify([

        {
            "id": jugador.id,

            "rut": jugador.rut,

            "nombre_completo":
                jugador.nombre_completo,

            "fecha_nacimiento":
                jugador.fecha_nacimiento.isoformat(),

            "serie":
                jugador.serie,

            "club":
                jugador.club,

            "estado":
                jugador.estado or "Vigente",

            "tiene_foto":
                bool(jugador.foto),

            "goles":
                obtener_goles(jugador.id),

            "amarillas":
                obtener_amarillas(jugador.id),

            "rojas":
                obtener_rojas(jugador.id),

            "suspensiones":
                obtener_suspensiones(jugador.id)
        }

        for jugador in jugadores

    ])


# ============================================================
# REGISTRAR GOL
# ============================================================

@app.route(
    "/jugadores/<int:jugador_id>/gol",
    methods=["POST"]
)
def registrar_gol(jugador_id):

    jugador = db.get_or_404(
        Jugador,
        jugador_id
    )

    try:

        cantidad = int(
            request.form.get(
                "cantidad",
                1
            )
        )

    except ValueError:

        cantidad = 1

    if cantidad < 1:
        cantidad = 1

    campeonato = request.form.get(
        "campeonato",
        ""
    ).strip()

    observaciones = request.form.get(
        "observaciones",
        ""
    ).strip()

    gol = Gol(
        jugador_id=jugador.id,
        fecha=date.today(),
        cantidad=cantidad,
        campeonato=campeonato,
        observaciones=observaciones
    )

    try:

        db.session.add(
            gol
        )

        db.session.commit()

    except Exception as error:

        db.session.rollback()

        print(
            "Error registrando gol:",
            error
        )

        flash(
            "No fue posible registrar el gol.",
            "error"
        )

        return redirect(
            url_for(
                "ficha_jugador",
                jugador_id=jugador.id
            )
        )

    flash(
        f"Se registraron {cantidad} gol(es).",
        "success"
    )

    return redirect(
        url_for(
            "ficha_jugador",
            jugador_id=jugador.id
        )
    )


# ============================================================
# REGISTRAR TARJETA AMARILLA
# ============================================================

@app.route(
    "/jugadores/<int:jugador_id>/amarilla",
    methods=["POST"]
)
def registrar_amarilla(jugador_id):

    jugador = db.get_or_404(
        Jugador,
        jugador_id
    )

    campeonato = request.form.get(
        "campeonato",
        ""
    ).strip()

    motivo = request.form.get(
        "motivo",
        ""
    ).strip()

    observaciones = request.form.get(
        "observaciones",
        ""
    ).strip()

    try:

        registro = RegistroDisciplinario(
            jugador_id=jugador.id,
            fecha=date.today(),
            tipo="Amarilla",
            cantidad=1,
            motivo=motivo,
            campeonato=campeonato,
            observaciones=observaciones
        )

        db.session.add(registro)

        db.session.flush()

        nuevas_suspensiones = crear_suspension_por_acumulacion(
            jugador,
            campeonato
        )

        db.session.commit()

    except Exception as error:

        db.session.rollback()

        print(
            "ERROR REGISTRANDO TARJETA AMARILLA:",
            repr(error)
        )

        flash(
            "No fue posible registrar la tarjeta amarilla.",
            "error"
        )

        return redirect(
            url_for(
                "ficha_jugador",
                jugador_id=jugador.id
            )
        )

    amarillas = obtener_amarillas(
        jugador.id
    )

    if nuevas_suspensiones:

        flash(
            f"Tarjeta amarilla registrada. "
            f"El jugador alcanzó {amarillas} amarillas "
            f"y se generó automáticamente una suspensión.",
            "warning"
        )

    else:

        flash(
            f"Tarjeta amarilla registrada correctamente. "
            f"Acumuladas: {amarillas}/4.",
            "success"
        )

    return redirect(
        url_for(
            "ficha_jugador",
            jugador_id=jugador.id
        )
    )


# ============================================================
# REGISTRAR TARJETA ROJA
# ============================================================

@app.route(
    "/jugadores/<int:jugador_id>/roja",
    methods=["POST"]
)
def registrar_roja(jugador_id):

    jugador = db.get_or_404(
        Jugador,
        jugador_id
    )

    campeonato = request.form.get(
        "campeonato",
        ""
    ).strip()

    motivo = request.form.get(
        "motivo",
        ""
    ).strip()

    observaciones = request.form.get(
        "observaciones",
        ""
    ).strip()

    try:

        registro = RegistroDisciplinario(
            jugador_id=jugador.id,
            fecha=date.today(),
            tipo="Roja",
            cantidad=1,
            motivo=motivo,
            campeonato=campeonato,
            observaciones=observaciones
        )

        db.session.add(registro)

        db.session.commit()

    except Exception as error:

        db.session.rollback()

        import traceback

        print("==============================================")
        print("ERROR REGISTRANDO TARJETA ROJA")
        print("==============================================")
        print(repr(error))
        traceback.print_exc()
        print("==============================================")

        flash(
            "No fue posible registrar la tarjeta roja.",
            "error"
        )

        return redirect(
            url_for(
                "ficha_jugador",
                jugador_id=jugador.id
            )
        )

    flash(
        "Tarjeta roja registrada correctamente.",
        "success"
    )

    return redirect(
        url_for(
            "ficha_jugador",
            jugador_id=jugador.id
        )
    )


# ============================================================
# REGISTRAR SUSPENSIÓN
# ============================================================

@app.route(
    "/jugadores/<int:jugador_id>/suspension",
    methods=["POST"]
)
def registrar_suspension(jugador_id):

    jugador = db.get_or_404(
        Jugador,
        jugador_id
    )

    try:

        cantidad = int(
            request.form.get(
                "cantidad",
                1
            )
        )

    except (ValueError, TypeError):

        cantidad = 1

    if cantidad < 1:
        cantidad = 1

    campeonato = request.form.get(
        "campeonato",
        ""
    ).strip()

    motivo = request.form.get(
        "motivo",
        ""
    ).strip()

    observaciones = request.form.get(
        "observaciones",
        ""
    ).strip()

    try:

        suspension = RegistroDisciplinario(
            jugador_id=jugador.id,
            fecha=date.today(),
            tipo="Suspension",
            cantidad=cantidad,
            motivo=motivo,
            campeonato=campeonato,
            observaciones=observaciones
        )

        db.session.add(suspension)

        jugador.estado = "Suspendido"

        db.session.commit()

    except Exception as error:

        db.session.rollback()

        print(
            "ERROR REGISTRANDO SUSPENSION:",
            repr(error)
        )

        flash(
            "No fue posible registrar la suspensión.",
            "error"
        )

        return redirect(
            url_for(
                "ficha_jugador",
                jugador_id=jugador.id
            )
        )

    flash(
        f"Suspensión registrada correctamente. "
        f"Cantidad: {cantidad}.",
        "success"
    )

    return redirect(
        url_for(
            "ficha_jugador",
            jugador_id=jugador.id
        )
    )

# ============================================================
# ELIMINAR REGISTRO DE GOL
# ============================================================

@app.route(
    "/jugadores/<int:jugador_id>/gol/<int:gol_id>/eliminar",
    methods=["POST"]
)
def eliminar_gol(jugador_id, gol_id):

    jugador = db.get_or_404(
        Jugador,
        jugador_id
    )

    gol = db.get_or_404(
        Gol,
        gol_id
    )

    if gol.jugador_id != jugador.id:

        flash(
            "El registro de gol no pertenece a este jugador.",
            "error"
        )

        return redirect(
            url_for(
                "ficha_jugador",
                jugador_id=jugador.id
            )
        )

    try:

        db.session.delete(gol)
        db.session.commit()

        flash(
            "Registro de gol eliminado correctamente.",
            "success"
        )

    except Exception as error:

        db.session.rollback()

        print("Error eliminando gol:", error)

        flash(
            "No fue posible eliminar el registro de gol.",
            "error"
        )

    return redirect(
        url_for(
            "ficha_jugador",
            jugador_id=jugador.id
        )
    )


# ============================================================
# ELIMINAR REGISTRO DISCIPLINARIO
# ============================================================

@app.route(
    "/jugadores/<int:jugador_id>/disciplina/<int:registro_id>/eliminar",
    methods=["POST"]
)
def eliminar_registro_disciplinario(jugador_id, registro_id):

    jugador = db.get_or_404(
        Jugador,
        jugador_id
    )

    registro = db.get_or_404(
        RegistroDisciplinario,
        registro_id
    )

    if registro.jugador_id != jugador.id:

        flash(
            "El registro disciplinario no pertenece a este jugador.",
            "error"
        )

        return redirect(
            url_for(
                "ficha_jugador",
                jugador_id=jugador.id
            )
        )

    try:

        db.session.delete(registro)
        db.session.commit()

        flash(
            "Registro disciplinario eliminado correctamente.",
            "success"
        )

    except Exception as error:

        db.session.rollback()

        print(
            "Error eliminando registro disciplinario:",
            error
        )

        flash(
            "No fue posible eliminar el registro disciplinario.",
            "error"
        )

    return redirect(
        url_for(
            "ficha_jugador",
            jugador_id=jugador.id
        )
    )


# ============================================================
# QR DEL JUGADOR
# ============================================================

@app.route(
    "/jugadores/<int:jugador_id>/qr"
)
def qr_jugador(jugador_id):

    jugador = db.get_or_404(
        Jugador,
        jugador_id
    )

    url_verificacion = url_for(
        "verificacion_jugador",
        jugador_id=jugador.id,
        _external=True
    )

    imagen_qr = qrcode.make(
        url_verificacion
    )

    memoria = BytesIO()

    imagen_qr.save(
        memoria,
        format="PNG"
    )

    memoria.seek(0)

    return Response(
        memoria.getvalue(),
        mimetype="image/png"
    )


# ============================================================
# VERIFICACIÓN PÚBLICA
# ============================================================

@app.route(
    "/jugadores/<int:jugador_id>/verificar"
)
def verificacion_jugador(jugador_id):

    jugador = db.get_or_404(
        Jugador,
        jugador_id
    )

    return render_template(
        "verificacion_jugador.html",
        jugador=jugador
    )


# ============================================================
# CREDENCIAL
# ============================================================

@app.route(
    "/jugadores/<int:jugador_id>/credencial"
)
def credencial_jugador(jugador_id):

    jugador = db.get_or_404(
        Jugador,
        jugador_id
    )

    return render_template(
        "jugador_credencial.html",
        jugador=jugador
    )


# ============================================================
# CREDENCIAL COMPLETA — FRENTE + REVERSO
# ============================================================

@app.route(
    "/jugadores/<int:jugador_id>/credencial-completa"
)
def credencial_completa(jugador_id):

    jugador = db.get_or_404(
        Jugador,
        jugador_id
    )

    return render_template(
        "credencial_completa.html",
        jugador=jugador
    )


# ============================================================
# REVERSO DE CREDENCIAL
# ============================================================

@app.route(
    "/jugadores/<int:jugador_id>/credencial/reverso"
)
def credencial_reverso(jugador_id):

    jugador = db.get_or_404(
        Jugador,
        jugador_id
    )

    return render_template(
        "credencial_reverso.html",
        jugador=jugador
    )


# ============================================================
# DASHBOARD
# ============================================================

@app.route("/dashboard")
def dashboard():

    """Panel principal.

    Las estadísticas deportivas son complementarias al registro de jugadores.
    Si una tabla estadística antigua no existe todavía en producción, el
    dashboard continúa funcionando mostrando esos indicadores en cero.
    """

    # ------------------------------------------------------------
    # INDICADORES GENERALES
    # ------------------------------------------------------------

    total_jugadores = Jugador.query.count()

    vigentes = Jugador.query.filter_by(
        estado="Vigente"
    ).count()

    pendientes = Jugador.query.filter_by(
        estado="Pendiente"
    ).count()

    suspendidos = Jugador.query.filter_by(
        estado="Suspendido"
    ).count()

    inhabilitados = Jugador.query.filter_by(
        estado="Inhabilitado"
    ).count()

    # ------------------------------------------------------------
    # CONSULTAS SEGURAS
    # ------------------------------------------------------------
    # Las tablas Gol y RegistroDisciplinario pueden no existir en una
    # base de datos antigua. Nunca deben impedir que cargue el Dashboard.

    def safe_all(query, default=None):
        try:
            return query.all()
        except Exception:
            db.session.rollback()
            return [] if default is None else default

    def safe_scalar(query, default=0):
        try:
            value = query.scalar()
            return value if value is not None else default
        except Exception:
            db.session.rollback()
            return default

    # ------------------------------------------------------------
    # JUGADORES POR CLUB
    # ------------------------------------------------------------

    jugadores_por_club = safe_all(
        db.session.query(
            Jugador.club,
            db.func.count(Jugador.id)
        )
        .filter(
            Jugador.club.isnot(None),
            Jugador.club != ""
        )
        .group_by(Jugador.club)
        .order_by(
            db.func.count(Jugador.id).desc()
        )
    )

    # ------------------------------------------------------------
    # JUGADORES POR SERIE
    # ------------------------------------------------------------

    jugadores_por_serie = safe_all(
        db.session.query(
            Jugador.serie,
            db.func.count(Jugador.id)
        )
        .filter(
            Jugador.serie.isnot(None),
            Jugador.serie != ""
        )
        .group_by(Jugador.serie)
        .order_by(
            db.func.count(Jugador.id).desc()
        )
    )

    # ------------------------------------------------------------
    # ÚLTIMOS JUGADORES
    # ------------------------------------------------------------

    ultimos_jugadores = (
        Jugador.query
        .order_by(Jugador.id.desc())
        .limit(5)
        .all()
    )

    # ------------------------------------------------------------
    # ESTADÍSTICAS DEPORTIVAS
    # ------------------------------------------------------------

    total_goles = safe_scalar(
        db.session.query(
            db.func.coalesce(
                db.func.sum(Gol.cantidad),
                0
            )
        )
    )

    total_amarillas = safe_scalar(
        db.session.query(
            db.func.coalesce(
                db.func.sum(RegistroDisciplinario.cantidad),
                0
            )
        )
        .filter(
            RegistroDisciplinario.tipo == "Amarilla"
        )
    )

    total_rojas = safe_scalar(
        db.session.query(
            db.func.coalesce(
                db.func.sum(RegistroDisciplinario.cantidad),
                0
            )
        )
        .filter(
            RegistroDisciplinario.tipo == "Roja"
        )
    )

    total_suspensiones = safe_scalar(
        db.session.query(
            db.func.coalesce(
                db.func.sum(RegistroDisciplinario.cantidad),
                0
            )
        )
        .filter(
            RegistroDisciplinario.tipo == "Suspension"
        )
    )

    # ------------------------------------------------------------
    # INDICADORES DE PARTICIPACIÓN DEPORTIVA
    # ------------------------------------------------------------

    jugadores_con_goles = safe_scalar(
        db.session.query(
            db.func.count(db.func.distinct(Gol.jugador_id))
        ),
        0
    )

    jugadores_con_amarillas = safe_scalar(
        db.session.query(
            db.func.count(db.func.distinct(RegistroDisciplinario.jugador_id))
        )
        .filter(
            RegistroDisciplinario.tipo == "Amarilla"
        ),
        0
    )

    jugadores_con_rojas = safe_scalar(
        db.session.query(
            db.func.count(db.func.distinct(RegistroDisciplinario.jugador_id))
        )
        .filter(
            RegistroDisciplinario.tipo == "Roja"
        ),
        0
    )

    jugadores_suspendidos_registro = safe_scalar(
        db.session.query(
            db.func.count(db.func.distinct(RegistroDisciplinario.jugador_id))
        )
        .filter(
            RegistroDisciplinario.tipo == "Suspension"
        ),
        0
    )

    # ------------------------------------------------------------
    # GOLEADORES
    # ------------------------------------------------------------

    goleadores = safe_all(
        db.session.query(
            Jugador.id,
            Jugador.nombre_completo,
            Jugador.club,
            Jugador.serie,
            db.func.sum(Gol.cantidad).label("total")
        )
        .join(
            Gol,
            Gol.jugador_id == Jugador.id
        )
        .group_by(
            Jugador.id,
            Jugador.nombre_completo,
            Jugador.club,
            Jugador.serie
        )
        .order_by(
            db.func.sum(Gol.cantidad).desc(),
            Jugador.nombre_completo.asc()
        )
        .limit(10)
    )

    # ------------------------------------------------------------
    # RANKING AMARILLAS
    # ------------------------------------------------------------

    ranking_amarillas = safe_all(
        db.session.query(
            Jugador.id,
            Jugador.nombre_completo,
            Jugador.club,
            db.func.sum(
                RegistroDisciplinario.cantidad
            ).label("total")
        )
        .join(
            RegistroDisciplinario,
            RegistroDisciplinario.jugador_id == Jugador.id
        )
        .filter(
            RegistroDisciplinario.tipo == "Amarilla"
        )
        .group_by(
            Jugador.id,
            Jugador.nombre_completo,
            Jugador.club
        )
        .order_by(
            db.func.sum(
                RegistroDisciplinario.cantidad
            ).desc(),
            Jugador.nombre_completo.asc()
        )
        .limit(10)
    )

    # ------------------------------------------------------------
    # RANKING ROJAS
    # ------------------------------------------------------------

    ranking_rojas = safe_all(
        db.session.query(
            Jugador.id,
            Jugador.nombre_completo,
            Jugador.club,
            db.func.sum(
                RegistroDisciplinario.cantidad
            ).label("total")
        )
        .join(
            RegistroDisciplinario,
            RegistroDisciplinario.jugador_id == Jugador.id
        )
        .filter(
            RegistroDisciplinario.tipo == "Roja"
        )
        .group_by(
            Jugador.id,
            Jugador.nombre_completo,
            Jugador.club
        )
        .order_by(
            db.func.sum(
                RegistroDisciplinario.cantidad
            ).desc(),
            Jugador.nombre_completo.asc()
        )
        .limit(10)
    )

    # ------------------------------------------------------------
    # RANKING SUSPENSIONES
    # ------------------------------------------------------------

    ranking_suspensiones = safe_all(
        db.session.query(
            Jugador.id,
            Jugador.nombre_completo,
            Jugador.club,
            db.func.sum(
                RegistroDisciplinario.cantidad
            ).label("total")
        )
        .join(
            RegistroDisciplinario,
            RegistroDisciplinario.jugador_id == Jugador.id
        )
        .filter(
            RegistroDisciplinario.tipo == "Suspension"
        )
        .group_by(
            Jugador.id,
            Jugador.nombre_completo,
            Jugador.club
        )
        .order_by(
            db.func.sum(
                RegistroDisciplinario.cantidad
            ).desc(),
            Jugador.nombre_completo.asc()
        )
        .limit(10)
    )

    return render_template(
        "dashboard.html",
        total_jugadores=total_jugadores,
        vigentes=vigentes,
        pendientes=pendientes,
        suspendidos=suspendidos,
        inhabilitados=inhabilitados,
        jugadores_por_club=jugadores_por_club,
        jugadores_por_serie=jugadores_por_serie,
        ultimos_jugadores=ultimos_jugadores,
        total_goles=total_goles,
        total_amarillas=total_amarillas,
        total_rojas=total_rojas,
        total_suspensiones=total_suspensiones,
        jugadores_con_goles=jugadores_con_goles,
        jugadores_con_amarillas=jugadores_con_amarillas,
        jugadores_con_rojas=jugadores_con_rojas,
        jugadores_suspendidos_registro=jugadores_suspendidos_registro,
        goleadores=goleadores,
        ranking_amarillas=ranking_amarillas,
        ranking_rojas=ranking_rojas,
        ranking_suspensiones=ranking_suspensiones
    )


# ============================================================
# IDENTIDAD VISUAL
# ============================================================

@app.route("/identidad")
def identidad_visual():
    """Punto de entrada para Identidad Visual.

    La navegación institucional puede apuntar a este endpoint aunque la
    configuración visual avanzada todavía no esté habilitada en esta versión.
    Redirigimos a Configuración para evitar errores de url_for en producción.
    """
    flash(
        "La identidad visual se administra desde Configuración en esta versión.",
        "info"
    )
    return redirect(url_for("configuracion"))


# ============================================================
# ADMINISTRACIÓN DE CLUBES Y SERIES
# ============================================================

@app.route(
    "/configuracion"
)
def configuracion():

    clubes = Club.query.order_by(
        Club.nombre
    ).all()

    series = Serie.query.order_by(
        Serie.nombre
    ).all()

    return render_template(
        "configuracion.html",
        clubes=clubes,
        series=series
    )


# ============================================================
# CREAR CLUB
# ============================================================

@app.route(
    "/clubes/nuevo",
    methods=["POST"]
)
def nuevo_club():

    nombre = request.form.get(
        "nombre",
        ""
    ).strip()

    if not nombre:

        flash(
            "Debes ingresar el nombre del club.",
            "error"
        )

        return redirect(
            url_for("configuracion")
        )

    club_existente = Club.query.filter(
        db.func.lower(Club.nombre) ==
        nombre.lower()
    ).first()

    if club_existente:

        flash(
            "Ese club ya está registrado.",
            "error"
        )

        return redirect(
            url_for("configuracion")
        )

    club = Club(
        nombre=nombre,
        activo=True
    )

    db.session.add(
        club
    )

    db.session.commit()

    flash(
        f"Club '{nombre}' agregado correctamente.",
        "success"
    )

    return redirect(
        url_for("configuracion")
    )


# ============================================================
# CREAR SERIE
# ============================================================

@app.route(
    "/series/nueva",
    methods=["POST"]
)
def nueva_serie():

    nombre = request.form.get(
        "nombre",
        ""
    ).strip()

    if not nombre:

        flash(
            "Debes ingresar el nombre de la serie.",
            "error"
        )

        return redirect(
            url_for("configuracion")
        )

    serie_existente = Serie.query.filter(
        db.func.lower(Serie.nombre) ==
        nombre.lower()
    ).first()

    if serie_existente:

        flash(
            "Esa serie ya está registrada.",
            "error"
        )

        return redirect(
            url_for("configuracion")
        )

    serie = Serie(
        nombre=nombre,
        activo=True
    )

    db.session.add(
        serie
    )

    db.session.commit()

    flash(
        f"Serie '{nombre}' agregada correctamente.",
        "success"
    )

    return redirect(
        url_for("configuracion")
    )


# ============================================================
# ACTIVAR / DESACTIVAR CLUB
# ============================================================

@app.route(
    "/clubes/<int:club_id>/estado",
    methods=["POST"]
)
def cambiar_estado_club(club_id):

    club = db.get_or_404(
        Club,
        club_id
    )

    club.activo = not club.activo

    db.session.commit()

    estado = (
        "activado"
        if club.activo
        else "desactivado"
    )

    flash(
        f"Club {estado} correctamente.",
        "success"
    )

    return redirect(
        url_for("configuracion")
    )


# ============================================================
# ACTIVAR / DESACTIVAR SERIE
# ============================================================

@app.route(
    "/series/<int:serie_id>/estado",
    methods=["POST"]
)
def cambiar_estado_serie(serie_id):

    serie = db.get_or_404(
        Serie,
        serie_id
    )

    serie.activo = not serie.activo

    db.session.commit()

    estado = (
        "activada"
        if serie.activo
        else "desactivada"
    )

    flash(
        f"Serie {estado} correctamente.",
        "success"
    )

    return redirect(
        url_for("configuracion")
    )


# ============================================================
# MÓDULO DE CAMPEONATOS — PASOS 1 A 4
# ============================================================

@app.route("/campeonatos")
def campeonatos():

    campeonatos_registrados = (
        Campeonato.query
        .order_by(
            Campeonato.temporada.desc(),
            Campeonato.id.desc()
        )
        .all()
    )

    return render_template(
        "campeonatos.html",
        campeonatos=campeonatos_registrados
    )


@app.route("/campeonatos/nuevo", methods=["GET", "POST"])
def nuevo_campeonato():

    series = (
        Serie.query
        .filter_by(activo=True)
        .order_by(Serie.nombre)
        .all()
    )

    if request.method == "POST":

        nombre = request.form.get("nombre", "").strip()
        temporada = request.form.get("temporada", "").strip()
        serie = request.form.get("serie", "").strip()
        estado = request.form.get("estado", "Activo").strip() or "Activo"
        descripcion = request.form.get("descripcion", "").strip()

        fecha_inicio = convertir_fecha(
            request.form.get("fecha_inicio", "").strip()
        )

        fecha_termino = convertir_fecha(
            request.form.get("fecha_termino", "").strip()
        )

        if not nombre or not temporada or not serie:
            flash(
                "Debes completar nombre, temporada y serie.",
                "error"
            )
            return render_template(
                "campeonato_form.html",
                series=series
            )

        if estado not in {"Activo", "Finalizado"}:
            estado = "Activo"

        if fecha_inicio and fecha_termino and fecha_termino < fecha_inicio:
            flash(
                "La fecha de término no puede ser anterior a la fecha de inicio.",
                "error"
            )
            return render_template(
                "campeonato_form.html",
                series=series
            )

        campeonato = Campeonato(
            nombre=nombre,
            temporada=temporada,
            serie=serie,
            fecha_inicio=fecha_inicio,
            fecha_termino=fecha_termino,
            estado=estado,
            descripcion=descripcion
        )

        try:
            db.session.add(campeonato)
            db.session.commit()
        except Exception as error:
            db.session.rollback()
            print("ERROR CREANDO CAMPEONATO:", repr(error))
            flash(
                "No fue posible crear el campeonato. Revise el registro del servidor.",
                "error"
            )
            return render_template(
                "campeonato_form.html",
                series=series
            )

        flash(
            f"Campeonato '{nombre}' creado correctamente.",
            "success"
        )

        return redirect(
            url_for(
                "detalle_campeonato",
                campeonato_id=campeonato.id
            )
        )

    return render_template(
        "campeonato_form.html",
        series=series
    )


@app.route("/campeonatos/<int:campeonato_id>/eliminar", methods=["POST"])
def eliminar_campeonato(campeonato_id):
    """Elimina un campeonato y todos sus datos exclusivos, sin tocar clubes ni jugadores."""
    campeonato = db.get_or_404(Campeonato, campeonato_id)
    nombre = campeonato.nombre

    try:
        # Primero eliminamos los datos que dependen directamente del campeonato.
        # Se usan filtros por campeonato_id para no borrar registros de otros campeonatos.
        partidos_eliminados = Partido.query.filter_by(
            campeonato_id=campeonato.id
        ).delete(synchronize_session=False)

        goles_eliminados = Gol.query.filter_by(
            campeonato_id=campeonato.id
        ).delete(synchronize_session=False)

        disciplina_eliminada = RegistroDisciplinario.query.filter_by(
            campeonato_id=campeonato.id
        ).delete(synchronize_session=False)

        clubes_eliminados = CampeonatoClub.query.filter_by(
            campeonato_id=campeonato.id
        ).delete(synchronize_session=False)

        db.session.delete(campeonato)
        db.session.commit()

        flash(
            f"Campeonato '{nombre}' eliminado correctamente. "
            f"Partidos: {partidos_eliminados}, goles: {goles_eliminados}, "
            f"registros disciplinarios: {disciplina_eliminada}, clubes inscritos: {clubes_eliminados}.",
            "success"
        )

    except Exception as error:
        db.session.rollback()
        print("ERROR ELIMINANDO CAMPEONATO:", repr(error))
        flash(
            "No fue posible eliminar el campeonato. No se modificaron los datos.",
            "error"
        )

    return redirect(url_for("campeonatos"))


@app.route("/campeonatos/<int:campeonato_id>/panel")
def panel_campeonato(campeonato_id):
    """Centro de control PRO de un campeonato."""
    campeonato = db.get_or_404(Campeonato, campeonato_id)

    clubes_participantes = (
        CampeonatoClub.query
        .filter_by(campeonato_id=campeonato.id)
        .join(Club, CampeonatoClub.club_id == Club.id)
        .order_by(Club.nombre)
        .all()
    )
    club_ids = {r.club_id for r in clubes_participantes}

    partidos = (
        Partido.query
        .filter_by(campeonato_id=campeonato.id)
        .order_by(Partido.jornada, Partido.id)
        .all()
    )
    partidos_finalizados = [p for p in partidos if p.estado == "Finalizado"]
    partidos_pendientes = [p for p in partidos if p.estado != "Finalizado"]

    # Tabla de posiciones
    tabla = {}
    for registro in clubes_participantes:
        club = registro.club
        tabla[club.id] = {"club": club, "pj": 0, "pg": 0, "pe": 0, "pp": 0,
                          "gf": 0, "gc": 0, "dg": 0, "pts": 0}

    for partido in partidos_finalizados:
        if partido.local_club_id not in tabla or partido.visitante_club_id not in tabla:
            continue
        gl = partido.goles_local or 0
        gv = partido.goles_visitante or 0
        local, visitante = tabla[partido.local_club_id], tabla[partido.visitante_club_id]
        local["pj"] += 1; visitante["pj"] += 1
        local["gf"] += gl; local["gc"] += gv
        visitante["gf"] += gv; visitante["gc"] += gl
        if gl > gv:
            local["pg"] += 1; visitante["pp"] += 1; local["pts"] += 3
        elif gl < gv:
            visitante["pg"] += 1; local["pp"] += 1; visitante["pts"] += 3
        else:
            local["pe"] += 1; visitante["pe"] += 1; local["pts"] += 1; visitante["pts"] += 1

    filas = list(tabla.values())
    for fila in filas:
        fila["dg"] = fila["gf"] - fila["gc"]
    filas.sort(key=lambda f: (-f["pts"], -f["dg"], -f["gf"], f["club"].nombre.lower()))
    for pos, fila in enumerate(filas, 1):
        fila["pos"] = pos

    # Próximo partido: prioriza los programados con fecha, luego jornada.
    proximos = [p for p in partidos if p.estado != "Finalizado"]
    proximos.sort(key=lambda p: (p.fecha is None, p.fecha or date.max, p.jornada, p.id))
    proximo_partido = proximos[0] if proximos else None

    # Club libre por jornada
    todos = {r.club_id: r.club for r in clubes_participantes}
    libres_por_jornada = {}
    jornadas = {}
    for p in partidos:
        jornadas.setdefault(p.jornada, []).append(p)
    for jornada, lista in jornadas.items():
        jugando = set()
        for p in lista:
            jugando.add(p.local_club_id); jugando.add(p.visitante_club_id)
        libres_por_jornada[jornada] = [todos[cid] for cid in todos if cid not in jugando]

    # Estadísticas del campeonato existentes en el sistema.
    goles = amarillas = rojas = suspensiones = []
    try:
        goles, amarillas, rojas, suspensiones = estadisticas_campeonato_data(campeonato)
    except Exception as error:
        db.session.rollback()
        print("ERROR PANEL ESTADISTICAS:", repr(error))

    total_goles = sum(int(x or 0) for _, x in goles)
    total_amarillas = sum(int(x or 0) for _, x in amarillas)
    total_rojas = sum(int(x or 0) for _, x in rojas)
    total_suspensiones = sum(int(x or 0) for _, x in suspensiones)

    jornadas_count = len(jornadas)
    porcentaje = round((len(partidos_finalizados) / len(partidos) * 100), 1) if partidos else 0

    return render_template(
        "campeonato_panel.html",
        campeonato=campeonato,
        clubes_participantes=clubes_participantes,
        partidos=partidos,
        partidos_finalizados=partidos_finalizados,
        partidos_pendientes=partidos_pendientes,
        filas=filas,
        proximo_partido=proximo_partido,
        libres_por_jornada=libres_por_jornada,
        jornadas_count=jornadas_count,
        porcentaje=porcentaje,
        goleadores=goles[:8],
        ranking_amarillas=amarillas[:8],
        ranking_rojas=rojas[:8],
        ranking_suspensiones=suspensiones[:8],
        total_goles=total_goles,
        total_amarillas=total_amarillas,
        total_rojas=total_rojas,
        total_suspensiones=total_suspensiones,
    )



@app.route("/campeonatos/<int:campeonato_id>/afiches")
def afiches_campeonato(campeonato_id):
    """Generador de afiches básico y compatible con la versión estable V5.8."""
    campeonato = db.get_or_404(Campeonato, campeonato_id)
    clubes_participantes = (
        CampeonatoClub.query
        .filter_by(campeonato_id=campeonato.id)
        .join(Club, CampeonatoClub.club_id == Club.id)
        .order_by(Club.nombre)
        .all()
    )
    partidos = (
        Partido.query
        .filter_by(campeonato_id=campeonato.id)
        .order_by(Partido.jornada, Partido.fecha, Partido.hora, Partido.id)
        .all()
    )
    jornadas = {}
    for partido in partidos:
        jornadas.setdefault(partido.jornada, []).append(partido)

    todos = {r.club_id: r.club for r in clubes_participantes}
    libres_por_jornada = {}
    for jornada, lista in jornadas.items():
        jugando = set()
        for partido in lista:
            jugando.add(partido.local_club_id)
            jugando.add(partido.visitante_club_id)
        libres_por_jornada[jornada] = [club for cid, club in todos.items() if cid not in jugando]

    jornada_seleccionada = request.args.get("jornada", type=int)
    if jornada_seleccionada not in jornadas and jornadas:
        jornada_seleccionada = min(jornadas)

    return render_template(
        "campeonato_afiches.html",
        campeonato=campeonato,
        clubes_participantes=clubes_participantes,
        jornadas=jornadas,
        libres_por_jornada=libres_por_jornada,
        jornada_seleccionada=jornada_seleccionada,
    )

@app.route("/campeonatos/<int:campeonato_id>")
def detalle_campeonato(campeonato_id):

    campeonato = db.get_or_404(
        Campeonato,
        campeonato_id
    )

    clubes = (
        Club.query
        .filter_by(activo=True)
        .order_by(Club.nombre)
        .all()
    )

    clubes_participantes = (
        CampeonatoClub.query
        .filter_by(campeonato_id=campeonato.id)
        .join(Club, CampeonatoClub.club_id == Club.id)
        .order_by(Club.nombre)
        .all()
    )

    clubes_participantes_ids = {
        registro.club_id
        for registro in clubes_participantes
    }

    return render_template(
        "campeonato_detalle.html",
        campeonato=campeonato,
        clubes=clubes,
        clubes_participantes=clubes_participantes,
        clubes_participantes_ids=clubes_participantes_ids
    )


@app.route(
    "/campeonatos/<int:campeonato_id>/clubes",
    methods=["POST"]
)
def guardar_clubes_campeonato(campeonato_id):

    campeonato = db.get_or_404(
        Campeonato,
        campeonato_id
    )

    valores = request.form.getlist("club_ids")
    club_ids = set()

    for valor in valores:
        try:
            club_id = int(valor)
            if club_id > 0:
                club_ids.add(club_id)
        except (TypeError, ValueError):
            continue

    try:
        clubes_validos = (
            Club.query
            .filter(
                Club.id.in_(club_ids)
            )
            .all()
            if club_ids else []
        )

        ids_validos = {club.id for club in clubes_validos}

        # Reemplazamos la inscripción del campeonato por la selección actual.
        CampeonatoClub.query.filter_by(
            campeonato_id=campeonato.id
        ).delete(
            synchronize_session=False
        )

        for club_id in sorted(ids_validos):
            db.session.add(
                CampeonatoClub(
                    campeonato_id=campeonato.id,
                    club_id=club_id
                )
            )

        db.session.commit()

    except Exception as error:
        db.session.rollback()
        print("ERROR GUARDANDO CLUBES DEL CAMPEONATO:", repr(error))
        flash(
            "No fue posible guardar los clubes participantes.",
            "error"
        )
        return redirect(
            url_for(
                "detalle_campeonato",
                campeonato_id=campeonato.id
            )
        )

    flash(
        f"Clubes participantes actualizados correctamente: {len(ids_validos)}.",
        "success"
    )

    return redirect(
        url_for(
            "detalle_campeonato",
            campeonato_id=campeonato.id
        )
    )


# ============================================================
# PASO 5 — FIXTURE DEL CAMPEONATO
# ============================================================

def siguiente_sabado(fecha_base):

    if not fecha_base:
        fecha_base = date.today()

    dias_hasta_sabado = (5 - fecha_base.weekday()) % 7

    return fecha_base + timedelta(days=dias_hasta_sabado)


def generar_calendario_todos_contra_todos(club_ids):
    """Genera una rueda todos-contra-todos usando el método de rotación."""

    equipos = list(club_ids)

    if len(equipos) < 2:
        return []

    # Con cantidad impar se agrega un descanso (None).
    if len(equipos) % 2:
        equipos.append(None)

    cantidad = len(equipos)
    rondas = cantidad - 1
    calendario = []

    for jornada in range(rondas):

        partidos_jornada = []

        for i in range(cantidad // 2):

            a = equipos[i]
            b = equipos[cantidad - 1 - i]

            if a is None or b is None:
                continue

            # Alternancia simple de localía para repartirla durante la rueda.
            if (jornada + i) % 2 == 0:
                local_id, visitante_id = a, b
            else:
                local_id, visitante_id = b, a

            partidos_jornada.append((local_id, visitante_id))

        calendario.append(partidos_jornada)

        # El primer equipo queda fijo y los demás rotan.
        equipos = [
            equipos[0],
            equipos[-1],
            *equipos[1:-1]
        ]

    return calendario


@app.route("/campeonatos/<int:campeonato_id>/fixture/exportar-word")
def exportar_fixture_word(campeonato_id):
    """Exporta todas las jornadas del fixture a un documento Word, una jornada por página."""
    campeonato = db.get_or_404(Campeonato, campeonato_id)

    partidos = (
        Partido.query
        .filter_by(campeonato_id=campeonato.id)
        .order_by(Partido.jornada, Partido.fecha, Partido.hora, Partido.id)
        .all()
    )

    if not partidos:
        flash("No hay partidos para exportar. Primero genera el fixture.", "error")
        return redirect(url_for("fixture_campeonato", campeonato_id=campeonato.id))

    jornadas = {}
    for partido in partidos:
        jornadas.setdefault(partido.jornada, []).append(partido)

    clubes_participantes = (
        CampeonatoClub.query
        .filter_by(campeonato_id=campeonato.id)
        .join(Club, CampeonatoClub.club_id == Club.id)
        .order_by(Club.nombre)
        .all()
    )
    todos_los_clubes = {registro.club_id: registro.club for registro in clubes_participantes}

    # Documento pensado como base para crear afiches: cada jornada ocupa una página.
    doc = Document()
    section = doc.sections[0]
    section.top_margin = Cm(1.4)
    section.bottom_margin = Cm(1.4)
    section.left_margin = Cm(1.5)
    section.right_margin = Cm(1.5)

    styles = doc.styles
    styles["Normal"].font.name = "Arial"
    styles["Normal"].font.size = Pt(11)

    def sombrear_celda(cell, fill):
        tcPr = cell._tc.get_or_add_tcPr()
        shd = tcPr.find(qn("w:shd"))
        if shd is None:
            shd = OxmlElement("w:shd")
            tcPr.append(shd)
        shd.set(qn("w:fill"), fill)

    def bordes_celda(cell, color="D1D5DB", size="8"):
        tc = cell._tc
        tcPr = tc.get_or_add_tcPr()
        borders = tcPr.first_child_found_in("w:tcBorders")
        if borders is None:
            borders = OxmlElement("w:tcBorders")
            tcPr.append(borders)
        for edge in ("top", "left", "bottom", "right", "insideH", "insideV"):
            tag = "w:" + edge
            element = borders.find(qn(tag))
            if element is None:
                element = OxmlElement(tag)
                borders.append(element)
            element.set(qn("w:val"), "single")
            element.set(qn("w:sz"), size)
            element.set(qn("w:color"), color)

    jornadas_items = list(jornadas.items())
    for indice, (jornada, lista) in enumerate(jornadas_items):
        if indice > 0:
            doc.add_page_break()

        titulo = doc.add_paragraph()
        titulo.alignment = WD_ALIGN_PARAGRAPH.CENTER
        run = titulo.add_run(f"JORNADA {jornada}")
        run.bold = True
        run.font.name = "Arial"
        run.font.size = Pt(28)

        subtitulo = doc.add_paragraph()
        subtitulo.alignment = WD_ALIGN_PARAGRAPH.CENTER
        run = subtitulo.add_run(campeonato.nombre)
        run.bold = True
        run.font.size = Pt(17)

        info = doc.add_paragraph()
        info.alignment = WD_ALIGN_PARAGRAPH.CENTER
        fechas = [p.fecha.strftime("%d/%m/%Y") for p in lista if p.fecha]
        horas = [p.hora for p in lista if p.hora]
        canchas = sorted({p.cancha for p in lista if p.cancha})
        fecha_txt = fechas[0] if fechas and len(set(fechas)) == 1 else "Fecha por definir"
        hora_txt = horas[0] if horas and len(set(horas)) == 1 else "Horarios según programación"
        cancha_txt = " · ".join(canchas) if canchas else "Cancha por definir"
        run = info.add_run(f"{campeonato.serie} · Temporada {campeonato.temporada}\n{fecha_txt} · {hora_txt} · {cancha_txt}")
        run.font.size = Pt(12)

        clubes_que_juegan = set()
        for p in lista:
            clubes_que_juegan.add(p.local_club_id)
            clubes_que_juegan.add(p.visitante_club_id)
        libres = [todos_los_clubes[cid] for cid in todos_los_clubes if cid not in clubes_que_juegan]

        libre_p = doc.add_paragraph()
        libre_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        libre_run = libre_p.add_run(
            "🆓 LIBRE: " + ", ".join(c.nombre for c in libres) if libres else "SIN CLUB LIBRE"
        )
        libre_run.bold = True
        libre_run.font.size = Pt(14)

        tabla = doc.add_table(rows=1, cols=5)
        tabla.autofit = False
        anchos = [Cm(1.5), Cm(5.1), Cm(1.7), Cm(5.1), Cm(4.0)]
        encabezados = ["N°", "LOCAL", "VS", "VISITANTE", "PROGRAMACIÓN"]
        for i, texto in enumerate(encabezados):
            cell = tabla.rows[0].cells[i]
            cell.width = anchos[i]
            cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
            sombrear_celda(cell, "1F2937")
            bordes_celda(cell, "FFFFFF", "10")
            para = cell.paragraphs[0]
            para.alignment = WD_ALIGN_PARAGRAPH.CENTER
            r = para.add_run(texto)
            r.bold = True
            r.font.size = Pt(10)

        for numero, partido in enumerate(lista, start=1):
            cells = tabla.add_row().cells
            for i, cell in enumerate(cells):
                cell.width = anchos[i]
                cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
                bordes_celda(cell)
                cell.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER

            valores = [
                str(numero),
                partido.local_club.nombre,
                "VS",
                partido.visitante_club.nombre,
                f"{partido.fecha.strftime('%d/%m/%Y') if partido.fecha else 'Sin fecha'}\n{partido.hora or 'Sin hora'}\n{partido.cancha or 'Sin cancha'}"
            ]
            for i, valor in enumerate(valores):
                r = cells[i].paragraphs[0].add_run(valor)
                r.bold = i in (1, 2, 3)
                r.font.size = Pt(11 if i in (1, 3) else 10)

            # Segunda línea opcional para el club de turno.
            if partido.turno_club:
                p = cells[4].add_paragraph()
                p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                r = p.add_run(f"Turno: {partido.turno_club.nombre}")
                r.bold = True
                r.font.size = Pt(9)

        nota = doc.add_paragraph()
        nota.alignment = WD_ALIGN_PARAGRAPH.CENTER
        r = nota.add_run("Documento base para diseño de afiche · Asociación Presidente Ríos")
        r.italic = True
        r.font.size = Pt(9)

    salida = BytesIO()
    doc.save(salida)
    salida.seek(0)

    nombre_seguro = "".join(c if c.isalnum() or c in " _-" else "_" for c in campeonato.nombre).strip() or "fixture"
    filename = f"Fixture_{nombre_seguro}.docx"

    return Response(
        salida.getvalue(),
        mimetype="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        headers={"Content-Disposition": f'attachment; filename="{filename}"'}
    )


@app.route("/campeonatos/<int:campeonato_id>/fixture")
def fixture_campeonato(campeonato_id):

    campeonato = db.get_or_404(
        Campeonato,
        campeonato_id
    )

    clubes_participantes = (
        CampeonatoClub.query
        .filter_by(campeonato_id=campeonato.id)
        .join(Club, CampeonatoClub.club_id == Club.id)
        .order_by(Club.nombre)
        .all()
    )

    partidos = (
        Partido.query
        .filter_by(campeonato_id=campeonato.id)
        .order_by(Partido.jornada, Partido.id)
        .all()
    )

    jornadas = {}
    for partido in partidos:
        jornadas.setdefault(partido.jornada, []).append(partido)

    # Calculamos automáticamente el/los clubes libres de cada jornada.
    # Para un todos-contra-todos con cantidad impar debe existir exactamente uno.
    todos_los_clubes = {registro.club_id: registro.club for registro in clubes_participantes}
    libres_por_jornada = {}
    for jornada, lista_partidos in jornadas.items():
        clubes_que_juegan = set()
        for partido in lista_partidos:
            clubes_que_juegan.add(partido.local_club_id)
            clubes_que_juegan.add(partido.visitante_club_id)
        libres_por_jornada[jornada] = [
            todos_los_clubes[club_id]
            for club_id in todos_los_clubes
            if club_id not in clubes_que_juegan
        ]

    return render_template(
        "campeonato_fixture.html",
        campeonato=campeonato,
        clubes_participantes=clubes_participantes,
        partidos=partidos,
        jornadas=jornadas,
        libres_por_jornada=libres_por_jornada
    )


@app.route(
    "/campeonatos/<int:campeonato_id>/fixture/generar",
    methods=["POST"]
)
def generar_fixture_campeonato(campeonato_id):

    campeonato = db.get_or_404(
        Campeonato,
        campeonato_id
    )

    clubes_participantes = (
        CampeonatoClub.query
        .filter_by(campeonato_id=campeonato.id)
        .order_by(CampeonatoClub.id)
        .all()
    )

    club_ids = [registro.club_id for registro in clubes_participantes]

    if len(club_ids) < 2:
        flash(
            "Debes tener al menos 2 clubes inscritos para generar el fixture.",
            "error"
        )
        return redirect(url_for("detalle_campeonato", campeonato_id=campeonato.id))

    hora = request.form.get("hora", "17:00").strip() or "17:00"
    cancha = request.form.get("cancha", "Por definir").strip() or "Por definir"

    try:
        # Al regenerar, se eliminan solamente los partidos de este campeonato.
        Partido.query.filter_by(
            campeonato_id=campeonato.id
        ).delete(synchronize_session=False)

        calendario = generar_calendario_todos_contra_todos(club_ids)
        primera_fecha = siguiente_sabado(campeonato.fecha_inicio)

        contador = 0

        for indice_jornada, partidos_jornada in enumerate(calendario, start=1):

            fecha_jornada = primera_fecha + timedelta(days=(indice_jornada - 1) * 7)

            for local_id, visitante_id in partidos_jornada:
                db.session.add(
                    Partido(
                        campeonato_id=campeonato.id,
                        jornada=indice_jornada,
                        fecha=fecha_jornada,
                        hora=hora,
                        cancha=cancha,
                        local_club_id=local_id,
                        visitante_club_id=visitante_id,
                        turno_club_id=None,
                        goles_local=None,
                        goles_visitante=None,
                        estado="Programado"
                    )
                )
                contador += 1

        db.session.commit()

        flash(
            f"Fixture generado correctamente: {len(calendario)} jornadas y {contador} partidos.",
            "success"
        )

    except Exception as error:
        db.session.rollback()
        print("ERROR GENERANDO FIXTURE:", repr(error))
        flash(
            "No fue posible generar el fixture. Revise el registro del servidor.",
            "error"
        )

    return redirect(
        url_for(
            "fixture_campeonato",
            campeonato_id=campeonato.id
        )
    )


@app.route(
    "/campeonatos/<int:campeonato_id>/fixture/jornada/<int:jornada>/configurar",
    methods=["POST"]
)
def configurar_jornada_fixture(campeonato_id, jornada):
    """Asigna fecha, hora y cancha a todos los partidos de una jornada."""
    campeonato = db.get_or_404(Campeonato, campeonato_id)

    try:
        partidos_jornada = (
            Partido.query
            .filter_by(campeonato_id=campeonato.id, jornada=jornada)
            .all()
        )

        if not partidos_jornada:
            raise ValueError("La jornada no tiene partidos.")

        fecha_texto = request.form.get("fecha", "").strip()
        if not fecha_texto:
            raise ValueError("Debes indicar una fecha.")

        fecha = datetime.strptime(fecha_texto, "%Y-%m-%d").date()
        hora = request.form.get("hora", "").strip() or None
        cancha = request.form.get("cancha", "").strip() or None

        for partido in partidos_jornada:
            partido.fecha = fecha
            partido.hora = hora
            partido.cancha = cancha

        db.session.commit()
        flash(
            f"Jornada {jornada} actualizada: {fecha.strftime('%d/%m/%Y')} · {hora or 'sin hora'} · {cancha or 'sin cancha'}.",
            "success"
        )

    except Exception as error:
        db.session.rollback()
        print("ERROR CONFIGURANDO JORNADA DEL FIXTURE:", repr(error))
        flash("No fue posible configurar la jornada. Revisa la fecha, horario y cancha.", "error")

    return redirect(url_for("fixture_campeonato", campeonato_id=campeonato.id))


@app.route(
    "/campeonatos/<int:campeonato_id>/fixture/eliminar",
    methods=["POST"]
)
def eliminar_fixture_campeonato(campeonato_id):

    campeonato = db.get_or_404(
        Campeonato,
        campeonato_id
    )

    try:
        eliminados = Partido.query.filter_by(
            campeonato_id=campeonato.id
        ).delete(synchronize_session=False)

        db.session.commit()

        flash(
            f"Fixture eliminado correctamente. Partidos eliminados: {eliminados}.",
            "success"
        )

    except Exception as error:
        db.session.rollback()
        print("ERROR ELIMINANDO FIXTURE:", repr(error))
        flash(
            "No fue posible eliminar el fixture.",
            "error"
        )

    return redirect(
        url_for(
            "fixture_campeonato",
            campeonato_id=campeonato.id
        )
    )



@app.route(
    "/campeonatos/<int:campeonato_id>/fixture/partido/<int:partido_id>/editar",
    methods=["POST"]
)
def editar_partido_fixture(campeonato_id, partido_id):
    """Edita fecha, hora, cancha y equipos de un partido sin regenerar el fixture."""
    campeonato = db.get_or_404(Campeonato, campeonato_id)
    partido = db.get_or_404(Partido, partido_id)

    if partido.campeonato_id != campeonato.id:
        flash("El partido no pertenece a este campeonato.", "error")
        return redirect(url_for("fixture_campeonato", campeonato_id=campeonato.id))

    try:
        fecha_texto = request.form.get("fecha", "").strip()
        if fecha_texto:
            partido.fecha = datetime.strptime(fecha_texto, "%Y-%m-%d").date()

        partido.hora = request.form.get("hora", "").strip() or None
        partido.cancha = request.form.get("cancha", "").strip() or None

        local_id = int(request.form.get("local_club_id", partido.local_club_id))
        visitante_id = int(request.form.get("visitante_club_id", partido.visitante_club_id))
        turno_raw = request.form.get("turno_club_id", "").strip()
        turno_id = int(turno_raw) if turno_raw else None

        participantes = {
            registro.club_id
            for registro in CampeonatoClub.query.filter_by(campeonato_id=campeonato.id).all()
        }

        if local_id == visitante_id:
            raise ValueError("El club local y visitante no pueden ser el mismo.")
        if local_id not in participantes or visitante_id not in participantes:
            raise ValueError("Los clubes seleccionados no pertenecen a este campeonato.")
        if turno_id is not None and turno_id not in participantes:
            raise ValueError("El club de turno seleccionado no pertenece a este campeonato.")

        # Evita que un club aparezca en dos partidos de la misma jornada.
        otros_partidos = (
            Partido.query
            .filter(
                Partido.campeonato_id == campeonato.id,
                Partido.jornada == partido.jornada,
                Partido.id != partido.id
            )
            .all()
        )
        ocupados = set()
        for otro in otros_partidos:
            ocupados.add(otro.local_club_id)
            ocupados.add(otro.visitante_club_id)

        if local_id in ocupados or visitante_id in ocupados:
            raise ValueError("Uno de los clubes seleccionados ya juega en otro partido de esta jornada.")

        partido.local_club_id = local_id
        partido.visitante_club_id = visitante_id
        partido.turno_club_id = turno_id

        db.session.commit()
        flash(f"Partido de la jornada {partido.jornada} actualizado correctamente.", "success")

    except Exception as error:
        db.session.rollback()
        print("ERROR EDITANDO PARTIDO DEL FIXTURE:", repr(error))
        flash("No fue posible modificar el partido. Revise los datos ingresados.", "error")

    return redirect(url_for("fixture_campeonato", campeonato_id=campeonato.id))


@app.route(
    "/campeonatos/<int:campeonato_id>/fixture/regenerar-nuevo",
    methods=["POST"]
)
def regenerar_fixture_nuevo(campeonato_id):
    """Regenera un fixture nuevo cambiando el orden de los clubes."""
    campeonato = db.get_or_404(Campeonato, campeonato_id)

    clubes_participantes = (
        CampeonatoClub.query
        .filter_by(campeonato_id=campeonato.id)
        .order_by(CampeonatoClub.id)
        .all()
    )
    club_ids = [registro.club_id for registro in clubes_participantes]

    if len(club_ids) < 2:
        flash("Debes tener al menos 2 clubes inscritos para generar el fixture.", "error")
        return redirect(url_for("fixture_campeonato", campeonato_id=campeonato.id))

    try:
        import random

        # Un orden nuevo produce una distribución distinta de enfrentamientos/localías.
        random.shuffle(club_ids)

        hora = request.form.get("hora", "17:00").strip() or "17:00"
        cancha = request.form.get("cancha", "Por definir").strip() or "Por definir"
        calendario = generar_calendario_todos_contra_todos(club_ids)
        primera_fecha = siguiente_sabado(campeonato.fecha_inicio)

        Partido.query.filter_by(campeonato_id=campeonato.id).delete(synchronize_session=False)

        contador = 0
        for indice_jornada, partidos_jornada in enumerate(calendario, start=1):
            fecha_jornada = primera_fecha + timedelta(days=(indice_jornada - 1) * 7)
            for local_id, visitante_id in partidos_jornada:
                db.session.add(
                    Partido(
                        campeonato_id=campeonato.id,
                        jornada=indice_jornada,
                        fecha=fecha_jornada,
                        hora=hora,
                        cancha=cancha,
                        local_club_id=local_id,
                        visitante_club_id=visitante_id,
                        goles_local=None,
                        goles_visitante=None,
                        estado="Programado"
                    )
                )
                contador += 1

        db.session.commit()
        flash(f"Nuevo fixture generado: {len(calendario)} jornadas y {contador} partidos.", "success")

    except Exception as error:
        db.session.rollback()
        print("ERROR REGENERANDO NUEVO FIXTURE:", repr(error))
        flash("No fue posible generar el nuevo fixture.", "error")

    return redirect(url_for("fixture_campeonato", campeonato_id=campeonato.id))


# ============================================================
# PASO 6 — REGISTRO DE RESULTADOS DEL CAMPEONATO
# ============================================================

@app.route("/campeonatos/<int:campeonato_id>/resultados")
def resultados_campeonato(campeonato_id):
    campeonato = db.get_or_404(Campeonato, campeonato_id)

    partidos = (
        Partido.query
        .filter_by(campeonato_id=campeonato.id)
        .order_by(Partido.jornada, Partido.id)
        .all()
    )

    jornadas = {}
    for partido in partidos:
        jornadas.setdefault(partido.jornada, []).append(partido)

    return render_template(
        "campeonato_resultados.html",
        campeonato=campeonato,
        partidos=partidos,
        jornadas=jornadas
    )


@app.route(
    "/campeonatos/<int:campeonato_id>/resultados/<int:partido_id>",
    methods=["POST"]
)
def registrar_resultado(campeonato_id, partido_id):
    campeonato = db.get_or_404(Campeonato, campeonato_id)
    partido = db.get_or_404(Partido, partido_id)

    if partido.campeonato_id != campeonato.id:
        flash("El partido no pertenece a este campeonato.", "error")
        return redirect(
            url_for("resultados_campeonato", campeonato_id=campeonato.id)
        )

    try:
        goles_local_raw = request.form.get("goles_local", "").strip()
        goles_visitante_raw = request.form.get("goles_visitante", "").strip()

        if goles_local_raw == "" or goles_visitante_raw == "":
            raise ValueError("Debes ingresar ambos marcadores.")

        goles_local = int(goles_local_raw)
        goles_visitante = int(goles_visitante_raw)

        if goles_local < 0 or goles_visitante < 0:
            raise ValueError("Los goles no pueden ser negativos.")

        partido.goles_local = goles_local
        partido.goles_visitante = goles_visitante
        partido.estado = "Finalizado"

        db.session.commit()

        flash(
            f"Resultado guardado: {partido.local_club.nombre} {goles_local} - {goles_visitante} {partido.visitante_club.nombre}.",
            "success"
        )

    except (TypeError, ValueError) as error:
        db.session.rollback()
        flash(f"No se pudo guardar el resultado: {error}", "error")
    except Exception as error:
        db.session.rollback()
        print("ERROR REGISTRANDO RESULTADO:", repr(error))
        flash("No fue posible guardar el resultado. Revise el registro del servidor.", "error")

    return redirect(
        url_for("resultados_campeonato", campeonato_id=campeonato.id)
    )


@app.route(
    "/campeonatos/<int:campeonato_id>/resultados/<int:partido_id>/programado",
    methods=["POST"]
)
def marcar_partido_programado(campeonato_id, partido_id):
    campeonato = db.get_or_404(Campeonato, campeonato_id)
    partido = db.get_or_404(Partido, partido_id)

    if partido.campeonato_id != campeonato.id:
        flash("El partido no pertenece a este campeonato.", "error")
        return redirect(
            url_for("resultados_campeonato", campeonato_id=campeonato.id)
        )

    try:
        partido.goles_local = None
        partido.goles_visitante = None
        partido.estado = "Programado"
        db.session.commit()
        flash("El partido volvió a estado Programado.", "success")
    except Exception as error:
        db.session.rollback()
        print("ERROR RESTABLECIENDO PARTIDO:", repr(error))
        flash("No fue posible restablecer el partido.", "error")

    return redirect(
        url_for("resultados_campeonato", campeonato_id=campeonato.id)
    )


# ============================================================
# PASO 7 — TABLA DE POSICIONES AUTOMÁTICA
# ============================================================

@app.route("/campeonatos/<int:campeonato_id>/tabla")
def tabla_campeonato(campeonato_id):
    campeonato = db.get_or_404(Campeonato, campeonato_id)

    clubes_participantes = (
        CampeonatoClub.query
        .filter_by(campeonato_id=campeonato.id)
        .join(Club, CampeonatoClub.club_id == Club.id)
        .order_by(Club.nombre)
        .all()
    )

    partidos_finalizados = (
        Partido.query
        .filter_by(campeonato_id=campeonato.id, estado="Finalizado")
        .all()
    )

    tabla = {}
    for registro in clubes_participantes:
        club = registro.club
        tabla[club.id] = {
            "club": club,
            "pj": 0,
            "pg": 0,
            "pe": 0,
            "pp": 0,
            "gf": 0,
            "gc": 0,
            "dg": 0,
            "pts": 0,
        }

    for partido in partidos_finalizados:
        # Solo contabilizamos partidos cuyos clubes siguen inscritos
        # en este campeonato.
        if partido.local_club_id not in tabla or partido.visitante_club_id not in tabla:
            continue

        gl = partido.goles_local if partido.goles_local is not None else 0
        gv = partido.goles_visitante if partido.goles_visitante is not None else 0

        local = tabla[partido.local_club_id]
        visitante = tabla[partido.visitante_club_id]

        local["pj"] += 1
        visitante["pj"] += 1
        local["gf"] += gl
        local["gc"] += gv
        visitante["gf"] += gv
        visitante["gc"] += gl

        if gl > gv:
            local["pg"] += 1
            visitante["pp"] += 1
            local["pts"] += 3
        elif gl < gv:
            visitante["pg"] += 1
            local["pp"] += 1
            visitante["pts"] += 3
        else:
            local["pe"] += 1
            visitante["pe"] += 1
            local["pts"] += 1
            visitante["pts"] += 1

    filas = list(tabla.values())
    for fila in filas:
        fila["dg"] = fila["gf"] - fila["gc"]

    # Orden oficial de clasificación: puntos, diferencia de goles,
    # goles a favor y nombre del club como desempate final.
    filas.sort(
        key=lambda fila: (
            -fila["pts"],
            -fila["dg"],
            -fila["gf"],
            fila["club"].nombre.lower(),
        )
    )

    for posicion, fila in enumerate(filas, start=1):
        fila["pos"] = posicion

    return render_template(
        "campeonato_tabla.html",
        campeonato=campeonato,
        filas=filas,
        partidos_finalizados=len(partidos_finalizados),
        partidos_totales=Partido.query.filter_by(campeonato_id=campeonato.id).count(),
    )

# ============================================================
# PASO 8 — ESTADÍSTICAS POR CAMPEONATO
# ============================================================

def jugadores_campeonato(campeonato):
    """Jugadores habilitados para las estadísticas del campeonato.

    Usa la misma normalización de club/serie que el acta digital, para que
    los goles y tarjetas registrados desde el acta aparezcan también en
    las estadísticas aunque existan diferencias de mayúsculas, tildes,
    guiones o nombres como "Senior" / "Serie Senior".
    """
    import unicodedata

    def normalizar(valor, quitar_serie=False):
        texto = unicodedata.normalize("NFKD", str(valor or ""))
        texto = "".join(c for c in texto if not unicodedata.combining(c))
        texto = texto.lower().strip()
        for caracter in "._-/":
            texto = texto.replace(caracter, " ")
        texto = " ".join(texto.split())
        if quitar_serie:
            texto = texto.replace("serie ", " ")
            texto = " ".join(texto.split())
        return texto

    registros = CampeonatoClub.query.filter_by(campeonato_id=campeonato.id).join(Club).all()
    if not registros:
        return []

    clubes = [registro.club for registro in registros]
    nombres_club = {normalizar(club.nombre) for club in clubes}
    serie_objetivo = normalizar(campeonato.serie, quitar_serie=True)

    candidatos = (Jugador.query
                  .filter(Jugador.club.isnot(None))
                  .order_by(Jugador.club, Jugador.nombre_completo)
                  .all())

    resultado = []
    for jugador in candidatos:
        club_norm = normalizar(jugador.club)
        if club_norm not in nombres_club:
            continue
        serie_norm = normalizar(jugador.serie, quitar_serie=True)
        if serie_norm == serie_objetivo:
            resultado.append(jugador)

    # Respaldo para bases históricas donde la serie fue cargada con un
    # nombre incompatible. Se limita a clubes participantes para no mezclar
    # jugadores de clubes ajenos al campeonato.
    if not resultado:
        resultado = [
            jugador for jugador in candidatos
            if normalizar(jugador.club) in nombres_club
        ]

    return resultado

def estadisticas_campeonato_data(campeonato):
    ids=[j.id for j in jugadores_campeonato(campeonato)]
    if not ids: return [],[],[],[]
    def rank(tipo=None):
        q=(db.session.query(Jugador, db.func.coalesce(db.func.sum(RegistroDisciplinario.cantidad),0).label('total'))
           .outerjoin(RegistroDisciplinario, db.and_(RegistroDisciplinario.jugador_id==Jugador.id,
               *( [RegistroDisciplinario.tipo==tipo] if tipo else [] ),
               db.or_(RegistroDisciplinario.campeonato_id==campeonato.id,
                      db.and_(RegistroDisciplinario.campeonato_id.is_(None), RegistroDisciplinario.campeonato==campeonato.nombre)) ))
           .filter(Jugador.id.in_(ids)).group_by(Jugador.id))
        return q.order_by(db.desc('total'),Jugador.nombre_completo).all()
    goles=(db.session.query(Jugador,db.func.coalesce(db.func.sum(Gol.cantidad),0).label('total'))
      .outerjoin(Gol,db.and_(Gol.jugador_id==Jugador.id,db.or_(Gol.campeonato_id==campeonato.id,db.and_(Gol.campeonato_id.is_(None),Gol.campeonato==campeonato.nombre))))
      .filter(Jugador.id.in_(ids)).group_by(Jugador.id).order_by(db.desc('total'),Jugador.nombre_completo).all())
    return goles,rank('Amarilla'),rank('Roja'),rank('Suspension')

@app.route('/campeonatos/<int:campeonato_id>/estadisticas')
def estadisticas_campeonato(campeonato_id):
    campeonato=db.get_or_404(Campeonato,campeonato_id)
    jugadores=jugadores_campeonato(campeonato)
    goles,amarillas,rojas,suspensiones=estadisticas_campeonato_data(campeonato)
    return render_template('campeonato_estadisticas.html',campeonato=campeonato,jugadores=jugadores,
        goleadores=goles,ranking_amarillas=amarillas,ranking_rojas=rojas,ranking_suspensiones=suspensiones,
        total_goles=sum(int(x or 0) for _,x in goles),total_amarillas=sum(int(x or 0) for _,x in amarillas),
        total_rojas=sum(int(x or 0) for _,x in rojas),total_suspensiones=sum(int(x or 0) for _,x in suspensiones))

def jugador_valido_campeonato(campeonato,jugador_id):
    return next((j for j in jugadores_campeonato(campeonato) if j.id==jugador_id),None)

@app.route('/campeonatos/<int:campeonato_id>/estadisticas/gol',methods=['POST'])
def registrar_gol_campeonato(campeonato_id):
    campeonato=db.get_or_404(Campeonato,campeonato_id)
    try: jugador_id=int(request.form.get('jugador_id','0')); cantidad=max(1,int(request.form.get('cantidad','1')))
    except (TypeError,ValueError): jugador_id,cantidad=0,1
    jugador=jugador_valido_campeonato(campeonato,jugador_id)
    if not jugador: flash('El jugador no pertenece a la serie o clubes de este campeonato.','error'); return redirect(url_for('estadisticas_campeonato',campeonato_id=campeonato.id))
    try:
        db.session.add(Gol(jugador_id=jugador.id,fecha=date.today(),cantidad=cantidad,campeonato=campeonato.nombre,campeonato_id=campeonato.id,observaciones=request.form.get('observaciones','').strip()))
        db.session.commit(); flash(f'Se registraron {cantidad} gol(es) para {jugador.nombre_completo}.','success')
    except Exception as error:
        db.session.rollback(); print('ERROR GOL CAMPEONATO:',repr(error)); flash('No fue posible registrar el gol.','error')
    return redirect(url_for('estadisticas_campeonato',campeonato_id=campeonato.id))

@app.route('/campeonatos/<int:campeonato_id>/estadisticas/disciplina',methods=['POST'])
def registrar_disciplina_campeonato(campeonato_id):
    campeonato=db.get_or_404(Campeonato,campeonato_id)
    try: jugador_id=int(request.form.get('jugador_id','0')); cantidad=max(1,int(request.form.get('cantidad','1')))
    except (TypeError,ValueError): jugador_id,cantidad=0,1
    tipo=request.form.get('tipo','').strip(); jugador=jugador_valido_campeonato(campeonato,jugador_id)
    if tipo not in {'Amarilla','Roja','Suspension'}: flash('Tipo disciplinario no válido.','error'); return redirect(url_for('estadisticas_campeonato',campeonato_id=campeonato.id))
    if not jugador: flash('El jugador no pertenece a la serie o clubes de este campeonato.','error'); return redirect(url_for('estadisticas_campeonato',campeonato_id=campeonato.id))
    try:
        db.session.add(RegistroDisciplinario(jugador_id=jugador.id,fecha=date.today(),tipo=tipo,cantidad=cantidad,campeonato=campeonato.nombre,campeonato_id=campeonato.id,motivo=request.form.get('motivo','').strip(),observaciones=request.form.get('observaciones','').strip()))
        if tipo=='Suspension': jugador.estado='Suspendido'
        db.session.commit(); flash(f'{tipo} registrada para {jugador.nombre_completo}.','success')
    except Exception as error:
        db.session.rollback(); print('ERROR DISCIPLINA CAMPEONATO:',repr(error)); flash('No fue posible registrar la disciplina.','error')
    return redirect(url_for('estadisticas_campeonato',campeonato_id=campeonato.id))

# ============================================================
# V5.7 — ACTA DIGITAL Y CONTROL DE NÓMINA
# ============================================================

def _normalizar_texto_acta(valor):
    """Normaliza nombres de club/serie para evitar problemas por mayúsculas,
    tildes, guiones o espacios diferentes entre el registro del jugador y el campeonato."""
    import unicodedata
    texto = unicodedata.normalize("NFKD", str(valor or ""))
    texto = "".join(c for c in texto if not unicodedata.combining(c))
    texto = texto.lower().strip()
    for caracter in "._-":
        texto = texto.replace(caracter, " ")
    return " ".join(texto.split())


def jugadores_disponibles_para_equipo(campeonato, club):
    """Obtiene los jugadores habilitados para el club del partido.

    Primero intenta coincidir club + serie de forma normalizada. Si la serie
    fue registrada con una variante de nombre (por ejemplo, 'Senior' frente
    a 'Serie Senior'), utiliza los jugadores del mismo club como respaldo.
    Esto evita que la nómina aparezca vacía por diferencias de escritura.
    """
    if not club:
        return []

    nombre_club = _normalizar_texto_acta(club.nombre)
    serie_campeonato = _normalizar_texto_acta(campeonato.serie)

    jugadores_club = (
        Jugador.query
        .filter(Jugador.club.isnot(None))
        .order_by(Jugador.nombre_completo)
        .all()
    )

    por_club = [
        jugador for jugador in jugadores_club
        if _normalizar_texto_acta(jugador.club) == nombre_club
    ]

    por_club_y_serie = [
        jugador for jugador in por_club
        if _normalizar_texto_acta(jugador.serie) == serie_campeonato
    ]

    # Caso habitual: club y serie coinciden exactamente.
    if por_club_y_serie:
        return por_club_y_serie

    # Respaldo: algunas cargas históricas usan nombres distintos para la
    # misma serie. Para el acta es preferible mostrar los jugadores del club
    # antes que dejar la nómina completamente vacía.
    return por_club

def sincronizar_estadisticas_desde_acta(campeonato, partido, nomina):
    """Reemplaza únicamente los registros generados por esta acta.

    Los registros ingresados manualmente desde el módulo de estadísticas no
    se tocan. Así se evita duplicar goles/tarjetas cada vez que se guarda
    el acta.
    """
    marcador = f"ACTA_PARTIDO:{partido.id}"

    Gol.query.filter(
        Gol.campeonato_id == campeonato.id,
        Gol.observaciones == marcador,
    ).delete(synchronize_session=False)

    RegistroDisciplinario.query.filter(
        RegistroDisciplinario.campeonato_id == campeonato.id,
        RegistroDisciplinario.observaciones == marcador,
    ).delete(synchronize_session=False)

    goles_local = 0
    goles_visitante = 0

    for registro in nomina:
        cantidad_goles = max(0, int(registro.goles or 0))
        amarillas = max(0, int(registro.amarillas or 0))
        rojas = max(0, int(registro.rojas or 0))

        if registro.equipo == "local":
            goles_local += cantidad_goles
        else:
            goles_visitante += cantidad_goles

        if cantidad_goles:
            db.session.add(Gol(
                jugador_id=registro.jugador_id,
                fecha=partido.fecha or date.today(),
                cantidad=cantidad_goles,
                campeonato=campeonato.nombre,
                campeonato_id=campeonato.id,
                observaciones=marcador,
            ))

        if amarillas:
            db.session.add(RegistroDisciplinario(
                jugador_id=registro.jugador_id,
                fecha=partido.fecha or date.today(),
                tipo="Amarilla",
                cantidad=amarillas,
                motivo=f"Acta partido #{partido.id}",
                campeonato=campeonato.nombre,
                campeonato_id=campeonato.id,
                observaciones=marcador,
            ))

        if rojas:
            db.session.add(RegistroDisciplinario(
                jugador_id=registro.jugador_id,
                fecha=partido.fecha or date.today(),
                tipo="Roja",
                cantidad=rojas,
                motivo=f"Acta partido #{partido.id}",
                campeonato=campeonato.nombre,
                campeonato_id=campeonato.id,
                observaciones=marcador,
            ))

    return goles_local, goles_visitante


@app.route("/campeonatos/<int:campeonato_id>/partido/<int:partido_id>/acta", methods=["GET", "POST"])
def acta_partido(campeonato_id, partido_id):
    campeonato = db.get_or_404(Campeonato, campeonato_id)
    partido = db.get_or_404(Partido, partido_id)

    if partido.campeonato_id != campeonato.id:
        flash("El partido no pertenece a este campeonato.", "error")
        return redirect(url_for("fixture_campeonato", campeonato_id=campeonato.id))

    acta = partido.acta

    if request.method == "POST":
        accion = request.form.get("accion", "guardar").strip().lower()

        # Una vez cerrada, el acta queda protegida contra modificaciones.
        if accion == "reabrir":
            if acta is None:
                flash("El acta todavía no existe.", "error")
            elif acta.estado != "Cerrada":
                flash("El acta ya está abierta.", "info")
            else:
                try:
                    acta.estado = "Borrador"
                    db.session.commit()
                    flash("Acta reabierta para edición.", "success")
                except Exception as error:
                    db.session.rollback()
                    print("ERROR REABRIENDO ACTA V5.8:", repr(error))
                    flash("No fue posible reabrir el acta.", "error")
            return redirect(url_for("acta_partido", campeonato_id=campeonato.id, partido_id=partido.id))

        if acta is not None and acta.estado == "Cerrada":
            flash("El acta está cerrada. Debes reabrirla antes de modificarla.", "error")
            return redirect(url_for("acta_partido", campeonato_id=campeonato.id, partido_id=partido.id))

        try:
            if acta is None:
                acta = ActaPartido(partido_id=partido.id)
                db.session.add(acta)
                db.session.flush()

            nuevo_estado = request.form.get("estado", "Borrador").strip()
            if nuevo_estado not in {"Borrador", "Cerrada"}:
                nuevo_estado = "Borrador"

            acta.numero_acta = request.form.get("numero_acta", "").strip() or None
            acta.arbitro = request.form.get("arbitro", "").strip() or None
            acta.observaciones = request.form.get("observaciones", "").strip() or None
            acta.estado = nuevo_estado

            PartidoJugador.query.filter_by(partido_id=partido.id).delete(synchronize_session=False)

            jugadores_ids = request.form.getlist("jugador_id")
            equipos = request.form.getlist("equipo")
            condiciones = request.form.getlist("condicion")
            capitanes = request.form.getlist("capitan")
            ingresos = request.form.getlist("ingreso")
            salidas = request.form.getlist("salida")
            goles = request.form.getlist("goles")
            amarillas = request.form.getlist("amarillas")
            rojas = request.form.getlist("rojas")
            observs = request.form.getlist("obs_jugador")

            permitidos = {
                "local": {j.id for j in jugadores_disponibles_para_equipo(campeonato, partido.local_club)},
                "visitante": {j.id for j in jugadores_disponibles_para_equipo(campeonato, partido.visitante_club)},
            }

            for i, raw_id in enumerate(jugadores_ids):
                if not raw_id.strip():
                    continue

                jid = int(raw_id)
                equipo = equipos[i] if i < len(equipos) else "local"
                if jid not in permitidos.get(equipo, set()):
                    raise ValueError("Jugador no perteneciente al club seleccionado.")

                def iv(lst):
                    try:
                        return max(0, int(lst[i])) if i < len(lst) and lst[i].strip() else 0
                    except (ValueError, TypeError):
                        return 0

                db.session.add(PartidoJugador(
                    partido_id=partido.id,
                    jugador_id=jid,
                    equipo=equipo,
                    condicion=condiciones[i] if i < len(condiciones) and condiciones[i] in {"Titular", "Suplente"} else "Suplente",
                    capitan=str(jid) in capitanes,
                    ingreso=(ingresos[i].strip() if i < len(ingresos) else None) or None,
                    salida=(salidas[i].strip() if i < len(salidas) else None) or None,
                    goles=iv(goles),
                    amarillas=iv(amarillas),
                    rojas=iv(rojas),
                    observaciones=(observs[i].strip() if i < len(observs) else None) or None,
                ))

            db.session.flush()
            nomina_actual = PartidoJugador.query.filter_by(partido_id=partido.id).all()
            goles_local, goles_visitante = sincronizar_estadisticas_desde_acta(
                campeonato, partido, nomina_actual
            )

            # El resultado oficial se toma de los goles registrados en el acta
            # solamente al momento de cerrarla.
            if acta.estado == "Cerrada":
                partido.goles_local = goles_local
                partido.goles_visitante = goles_visitante
                partido.estado = "Finalizado"

            db.session.commit()

            if acta.estado == "Cerrada":
                flash(
                    f"Acta cerrada. Resultado actualizado: {partido.local_club.nombre} "
                    f"{goles_local} - {goles_visitante} {partido.visitante_club.nombre}. "
                    "Estadísticas y disciplina sincronizadas.",
                    "success",
                )
            else:
                flash(
                    "Acta guardada como borrador. Goles y disciplina quedaron sincronizados.",
                    "success",
                )

        except Exception as error:
            db.session.rollback()
            print("ERROR GUARDANDO ACTA V5.8:", repr(error))
            flash("No fue posible guardar el acta. Revise los datos e inténtelo nuevamente.", "error")

        return redirect(url_for("acta_partido", campeonato_id=campeonato.id, partido_id=partido.id))

    return render_template(
        "partido_acta_nomina.html",
        campeonato=campeonato,
        partido=partido,
        acta=acta,
        nomina=partido.nomina,
        locales=jugadores_disponibles_para_equipo(campeonato, partido.local_club),
        visitantes=jugadores_disponibles_para_equipo(campeonato, partido.visitante_club),
    )

@app.route("/campeonatos/<int:campeonato_id>/partido/<int:partido_id>/acta/word")
def exportar_acta_word(campeonato_id, partido_id):
    campeonato=db.get_or_404(Campeonato,campeonato_id); partido=db.get_or_404(Partido,partido_id)
    if partido.campeonato_id != campeonato.id: return redirect(url_for("fixture_campeonato",campeonato_id=campeonato.id))
    acta=partido.acta; doc=Document(); sec=doc.sections[0]
    sec.top_margin=Cm(1.5); sec.bottom_margin=Cm(1.5); sec.left_margin=Cm(1.5); sec.right_margin=Cm(1.5)
    p=doc.add_paragraph(); p.alignment=WD_ALIGN_PARAGRAPH.CENTER; r=p.add_run("ASOCIACIÓN DE FÚTBOL PRESIDENTE RÍOS"); r.bold=True; r.font.size=Pt(16)
    p=doc.add_paragraph(); p.alignment=WD_ALIGN_PARAGRAPH.CENTER; r=p.add_run(f"ACTA DIGITAL · {campeonato.nombre}"); r.bold=True
    t=doc.add_table(rows=6,cols=2); t.style="Table Grid"
    datos=[("Acta",acta.numero_acta if acta and acta.numero_acta else "-"),("Jornada",str(partido.jornada)),("Fecha",partido.fecha.strftime("%d/%m/%Y") if partido.fecha else "-"),("Hora / Cancha",f"{partido.hora or '-'} · {partido.cancha or '-'}"),("Partido",f"{partido.local_club.nombre} vs {partido.visitante_club.nombre}"),("Árbitro",acta.arbitro if acta and acta.arbitro else "-")]
    for row,(a,b) in zip(t.rows,datos): row.cells[0].text=a; row.cells[1].text=b
    doc.add_paragraph("NÓMINA Y EVENTOS")
    nt=doc.add_table(rows=1,cols=8); nt.style="Table Grid"
    for c,h in zip(nt.rows[0].cells,["Jugador","Equipo","Condición","Capitán","Ingreso","Salida","Goles","Tarjetas"]): c.text=h
    for n in sorted(partido.nomina,key=lambda x:(x.equipo,x.jugador.nombre_completo)):
        vals=[n.jugador.nombre_completo,n.equipo.title(),n.condicion,"Sí" if n.capitan else "",n.ingreso or "",n.salida or "",str(n.goles),f"A:{n.amarillas} / R:{n.rojas}"]
        for c,v in zip(nt.add_row().cells,vals): c.text=v
    if acta and acta.observaciones: doc.add_paragraph("Observaciones: "+acta.observaciones)
    doc.add_paragraph("\n____________________________        ____________________________\nFirma Árbitro                                      Firma Club de Turno")
    out=BytesIO(); doc.save(out); out.seek(0); safe="".join(c if c.isalnum() or c in " _-" else "_" for c in partido.local_club.nombre+"_vs_"+partido.visitante_club.nombre)
    return Response(out.getvalue(),mimetype="application/vnd.openxmlformats-officedocument.wordprocessingml.document",headers={"Content-Disposition":f'attachment; filename="Acta_{safe}.docx"'})


# ============================================================
# V5.9 — PORTAL PÚBLICO
# ============================================================

def obtener_tabla_publica(campeonato):
    registros = (CampeonatoClub.query.filter_by(campeonato_id=campeonato.id)
                 .join(Club, CampeonatoClub.club_id == Club.id).order_by(Club.nombre).all())
    finalizados = Partido.query.filter_by(campeonato_id=campeonato.id, estado="Finalizado").all()
    tabla = {}
    for r in registros:
        tabla[r.club.id] = {"club": r.club, "pj": 0, "pg": 0, "pe": 0, "pp": 0, "gf": 0, "gc": 0, "dg": 0, "pts": 0}
    for p in finalizados:
        if p.local_club_id not in tabla or p.visitante_club_id not in tabla:
            continue
        gl, gv = int(p.goles_local or 0), int(p.goles_visitante or 0)
        l, v = tabla[p.local_club_id], tabla[p.visitante_club_id]
        l["pj"] += 1; v["pj"] += 1; l["gf"] += gl; l["gc"] += gv; v["gf"] += gv; v["gc"] += gl
        if gl > gv:
            l["pg"] += 1; v["pp"] += 1; l["pts"] += 3
        elif gv > gl:
            v["pg"] += 1; l["pp"] += 1; v["pts"] += 3
        else:
            l["pe"] += 1; v["pe"] += 1; l["pts"] += 1; v["pts"] += 1
    filas = list(tabla.values())
    for f in filas: f["dg"] = f["gf"] - f["gc"]
    filas.sort(key=lambda f: (-f["pts"], -f["dg"], -f["gf"], f["club"].nombre.lower()))
    for pos, f in enumerate(filas, 1): f["pos"] = pos
    return filas


def obtener_proxima_jornada(campeonato):
    partidos = (Partido.query.filter_by(campeonato_id=campeonato.id)
                .filter(Partido.estado != "Finalizado")
                .order_by(Partido.jornada, Partido.fecha, Partido.hora, Partido.id).all())
    if not partidos:
        return None, []
    jornada = partidos[0].jornada
    return jornada, [p for p in partidos if p.jornada == jornada]


def obtener_goleadores_publicos(campeonato, limite=50):
    ids = [j.id for j in jugadores_campeonato(campeonato)]
    if not ids:
        return []
    filas = (db.session.query(Jugador, db.func.coalesce(db.func.sum(Gol.cantidad), 0).label("total"))
             .outerjoin(Gol, db.and_(Gol.jugador_id == Jugador.id,
                 db.or_(Gol.campeonato_id == campeonato.id,
                       db.and_(Gol.campeonato_id.is_(None), Gol.campeonato == campeonato.nombre))))
             .filter(Jugador.id.in_(ids)).group_by(Jugador.id)
             .order_by(db.desc("total"), Jugador.nombre_completo).limit(limite).all())
    return filas


@app.route("/publico")
def publico():
    campeonatos = (Campeonato.query.filter(Campeonato.estado.ilike("Activo"))
                   .order_by(Campeonato.fecha_inicio.desc().nullslast(), Campeonato.id.desc()).all())
    paneles = []
    series = {}
    for campeonato in campeonatos:
        jornada, partidos = obtener_proxima_jornada(campeonato)
        panel = {"campeonato": campeonato, "jornada": jornada, "partidos": partidos,
                 "tabla": obtener_tabla_publica(campeonato),
                 "goleadores": obtener_goleadores_publicos(campeonato, 10)}
        paneles.append(panel)
        nombre_serie = (campeonato.serie or "Sin serie").strip() or "Sin serie"
        series.setdefault(nombre_serie, []).append(panel)
    series_publicas = sorted(series.items(), key=lambda item: item[0].lower())
    return render_template("publico.html", paneles=paneles, campeonatos=campeonatos,
                           series_publicas=series_publicas)


@app.route("/publico/campeonato/<int:campeonato_id>")
def publico_campeonato(campeonato_id):
    # Página de detalle pública. Se mantiene independiente del resto del
    # portal para evitar que un problema con una plantilla secundaria
    # provoque un 500 en Railway. Los datos son los mismos que usa el
    # portal público principal.
    campeonato = db.get_or_404(Campeonato, campeonato_id)
    jornada, partidos = obtener_proxima_jornada(campeonato)
    tabla = obtener_tabla_publica(campeonato)
    goleadores = obtener_goleadores_publicos(campeonato, 50)

    return render_template(
        "publico_campeonato.html",
        campeonato=campeonato,
        jornada=jornada,
        partidos=partidos,
        tabla=tabla,
        goleadores=goleadores,
    )


@app.route("/publico/campeonato/<int:campeonato_id>/tabla")
def publico_tabla(campeonato_id):
    campeonato = db.get_or_404(Campeonato, campeonato_id)
    return render_template("publico_tabla.html", campeonato=campeonato, tabla=obtener_tabla_publica(campeonato))


@app.route("/publico/campeonato/<int:campeonato_id>/goleadores")
def publico_goleadores(campeonato_id):
    campeonato = db.get_or_404(Campeonato, campeonato_id)
    return render_template("publico_goleadores.html", campeonato=campeonato, goleadores=obtener_goleadores_publicos(campeonato, 100))


# ============================================================
# HEALTH CHECK
# ============================================================

@app.route(
    "/health"
)
def health():

    return {
        "status": "ok"
    }


# ============================================================
# EJECUCIÓN LOCAL
# ============================================================

if __name__ == "__main__":

    app.run(
        host="0.0.0.0",
        port=int(
            os.environ.get(
                "PORT",
                5000
            )
        )
    )
