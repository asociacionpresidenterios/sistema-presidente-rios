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
    rol = db.Column(db.String(30), nullable=False, default="Administrador")
    creado_en = db.Column(db.DateTime, nullable=False, default=datetime.utcnow)

    def set_password(self, password):
        self.password_hash = generate_password_hash(password)

    def check_password(self, password):
        return check_password_hash(self.password_hash, password)


# ============================================================
# V9.0 — MODELO DE TESORERÍA
# ============================================================

class RendicionTesoreria(db.Model):
    """
    V10.4 — Cierre formal de rendiciones mensuales.
    Un período puede estar en Borrador, Revisada o Cerrada.
    Al cerrar se guardan los totales como fotografía/auditoría.
    """
    __tablename__ = "rendicion_tesoreria"

    id = db.Column(db.Integer, primary_key=True)
    periodo_mes = db.Column(db.Integer, nullable=False, index=True)
    periodo_anio = db.Column(db.Integer, nullable=False, index=True)
    numero_rendicion = db.Column(db.String(40), nullable=False, unique=True)
    estado = db.Column(db.String(20), nullable=False, default="Borrador", index=True)

    saldo_inicial = db.Column(db.Numeric(14, 0), nullable=True)
    ingresos = db.Column(db.Numeric(14, 0), nullable=True)
    egresos = db.Column(db.Numeric(14, 0), nullable=True)
    saldo_final = db.Column(db.Numeric(14, 0), nullable=True)
    movimientos = db.Column(db.Integer, nullable=True)
    cuentas_por_cobrar = db.Column(db.Numeric(14, 0), nullable=True)
    cuentas_por_pagar = db.Column(db.Numeric(14, 0), nullable=True)

    creado_por = db.Column(db.String(160), nullable=True)
    creado_en = db.Column(db.DateTime, nullable=False, default=datetime.utcnow)
    revisado_por = db.Column(db.String(160), nullable=True)
    revisado_en = db.Column(db.DateTime, nullable=True)
    cerrado_por = db.Column(db.String(160), nullable=True)
    cerrado_en = db.Column(db.DateTime, nullable=True)
    observaciones = db.Column(db.Text, nullable=True)

    __table_args__ = (
        db.UniqueConstraint(
            "periodo_mes",
            "periodo_anio",
            name="uq_rendicion_tesoreria_periodo"
        ),
    )


class MovimientoTesoreria(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    fecha = db.Column(db.Date, nullable=False, default=date.today, index=True)
    tipo = db.Column(db.String(20), nullable=False, index=True)  # Ingreso / Egreso
    concepto = db.Column(db.String(180), nullable=False)
    categoria = db.Column(db.String(80), nullable=True)
    monto = db.Column(db.Numeric(14, 0), nullable=False, default=0)
    medio_pago = db.Column(db.String(50), nullable=True)
    referencia = db.Column(db.String(120), nullable=True)
    observaciones = db.Column(db.Text, nullable=True)
    creado_por = db.Column(db.String(160), nullable=True)
    club_id = db.Column(db.Integer, db.ForeignKey("club.id"), nullable=True, index=True)
    campeonato_id = db.Column(db.Integer, db.ForeignKey("campeonato.id"), nullable=True, index=True)
    serie = db.Column(db.String(80), nullable=True)
    cuenta_id = db.Column(db.Integer, nullable=True, index=True)
    creado_en = db.Column(db.DateTime, nullable=False, default=datetime.utcnow)

    club = db.relationship("Club", foreign_keys=[club_id])
    campeonato = db.relationship("Campeonato", foreign_keys=[campeonato_id])

# ============================================================
# MODELO CLUB
# ============================================================

class CuentaTesoreria(db.Model):

    id = db.Column(db.Integer, primary_key=True)
    fecha = db.Column(db.Date, nullable=False, default=date.today, index=True)
    tipo = db.Column(db.String(20), nullable=False, index=True)  # Por cobrar / Por pagar
    club_id = db.Column(db.Integer, db.ForeignKey("club.id"), nullable=True, index=True)
    campeonato_id = db.Column(db.Integer, db.ForeignKey("campeonato.id"), nullable=True, index=True)
    serie = db.Column(db.String(80), nullable=True)
    concepto = db.Column(db.String(180), nullable=False)
    monto_total = db.Column(db.Numeric(14, 0), nullable=False, default=0)
    monto_pagado = db.Column(db.Numeric(14, 0), nullable=False, default=0)
    vencimiento = db.Column(db.Date, nullable=True)
    estado = db.Column(db.String(30), nullable=False, default="Pendiente")
    observaciones = db.Column(db.Text, nullable=True)
    creado_por = db.Column(db.String(160), nullable=True)
    partido_id = db.Column(db.Integer, db.ForeignKey("partido.id"), nullable=True, index=True)
    origen = db.Column(db.String(40), nullable=True, index=True)
    creado_en = db.Column(db.DateTime, nullable=False, default=datetime.utcnow)

    club = db.relationship("Club", foreign_keys=[club_id])
    campeonato = db.relationship("Campeonato", foreign_keys=[campeonato_id])
    partido = db.relationship("Partido", foreign_keys=[partido_id])

    @property
    def saldo(self):
        return max(
            int(self.monto_total or 0) - int(self.monto_pagado or 0),
            0
        )

    def actualizar_estado(self):
        total = int(self.monto_total or 0)
        pagado = int(self.monto_pagado or 0)

        if pagado >= total and total > 0:
            self.estado = "Pagado"
        elif pagado > 0:
            self.estado = "Abono"
        else:
            self.estado = "Pendiente"


def rango_periodo_tesoreria(mes, anio):
    inicio = date(anio, mes, 1)
    fin = date(anio + 1, 1, 1) if mes == 12 else date(anio, mes + 1, 1)
    return inicio, fin


def rendicion_tesoreria_periodo(fecha):
    if not fecha:
        return None
    return RendicionTesoreria.query.filter_by(
        periodo_mes=fecha.month,
        periodo_anio=fecha.year,
    ).first()


def tesoreria_periodo_cerrado(fecha):
    rendicion = rendicion_tesoreria_periodo(fecha)
    return bool(rendicion and rendicion.estado == "Cerrada")


def mensaje_periodo_cerrado(fecha):
    rendicion = rendicion_tesoreria_periodo(fecha)
    if not rendicion or rendicion.estado != "Cerrada":
        return None
    return (
        f"El período {fecha.month:02d}/{fecha.year} está cerrado "
        f"({rendicion.numero_rendicion}). No se permiten modificaciones "
        f"con fecha dentro de ese período."
    )


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
# ETAPA 10 — HISTORIAL DE INTEGRACIÓN DEL REGISTRO MAESTRO
# ============================================================

class JugadorMovimiento(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    jugador_id = db.Column(db.Integer, db.ForeignKey("jugador.id"), nullable=False, index=True)
    fecha_hora = db.Column(db.DateTime, nullable=False, default=datetime.utcnow, index=True)
    tipo = db.Column(db.String(40), nullable=False)
    club_anterior = db.Column(db.String(120), nullable=True)
    serie_anterior = db.Column(db.String(80), nullable=True)
    estado_anterior = db.Column(db.String(30), nullable=True)
    club_nuevo = db.Column(db.String(120), nullable=True)
    serie_nueva = db.Column(db.String(80), nullable=True)
    estado_nuevo = db.Column(db.String(30), nullable=True)
    motivo = db.Column(db.String(255), nullable=True)
    realizado_por = db.Column(db.String(80), nullable=True)
    jugador = db.relationship("Jugador", backref=db.backref("movimientos", lazy=True, cascade="all, delete-orphan"))


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
# V11.0 — COMITÉ DE DISCIPLINA / RESOLUCIONES OFICIALES
# ============================================================

class ResolucionDisciplina(db.Model):
    __tablename__ = "resolucion_disciplina"

    id = db.Column(db.Integer, primary_key=True)
    numero_resolucion = db.Column(db.String(40), unique=True, nullable=False, index=True)
    fecha = db.Column(db.Date, nullable=False, default=date.today, index=True)
    fecha_publicacion = db.Column(db.Date, nullable=True)

    campeonato_id = db.Column(db.Integer, db.ForeignKey("campeonato.id"), nullable=True, index=True)
    partido_id = db.Column(db.Integer, db.ForeignKey("partido.id"), nullable=True, index=True)
    jugador_id = db.Column(db.Integer, db.ForeignKey("jugador.id"), nullable=True, index=True)
    club_id = db.Column(db.Integer, db.ForeignKey("club.id"), nullable=True, index=True)
    serie = db.Column(db.String(80), nullable=True)

    tipo_decision = db.Column(db.String(60), nullable=False, default="Resolución")
    titulo = db.Column(db.String(220), nullable=False)
    antecedentes = db.Column(db.Text, nullable=True)
    resolucion = db.Column(db.Text, nullable=False)
    sancion = db.Column(db.String(255), nullable=True)
    cantidad_sancion = db.Column(db.Integer, nullable=True)
    unidad_sancion = db.Column(db.String(50), nullable=True)
    fecha_inicio = db.Column(db.Date, nullable=True)
    fecha_fin = db.Column(db.Date, nullable=True)
    observaciones = db.Column(db.Text, nullable=True)

    estado = db.Column(db.String(30), nullable=False, default="Borrador", index=True)
    publicado = db.Column(db.Boolean, nullable=False, default=False, index=True)

    creado_por = db.Column(db.String(160), nullable=True)
    creado_en = db.Column(db.DateTime, nullable=False, default=datetime.utcnow)
    actualizado_por = db.Column(db.String(160), nullable=True)
    actualizado_en = db.Column(db.DateTime, nullable=True)

    campeonato = db.relationship("Campeonato", foreign_keys=[campeonato_id])
    partido = db.relationship("Partido", foreign_keys=[partido_id])
    jugador = db.relationship("Jugador", foreign_keys=[jugador_id])
    club = db.relationship("Club", foreign_keys=[club_id])


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
        # AGREGAR ROL A ADMINISTRADORES SI NO EXISTE
        # ----------------------------------------------------
        inspector = db.inspect(db.engine)
        columnas_admin = [c["name"] for c in inspector.get_columns("admin_user")]
        if "rol" not in columnas_admin:
            if db.engine.dialect.name == "postgresql":
                db.session.execute(db.text("ALTER TABLE admin_user ADD COLUMN IF NOT EXISTS rol VARCHAR(30) DEFAULT 'Administrador'"))
            elif db.engine.dialect.name == "sqlite":
                db.session.execute(db.text("ALTER TABLE admin_user ADD COLUMN rol VARCHAR(30) DEFAULT 'Administrador'"))
            db.session.commit()
        db.session.execute(db.text("UPDATE admin_user SET rol = 'Administrador' WHERE rol IS NULL OR rol = ''"))
        db.session.commit()

        # ----------------------------------------------------
        # TESORERÍA — MIGRACIÓN SEGURA
        # Vincula movimientos con club, campeonato, serie y cuenta.
        # ----------------------------------------------------
        inspector = db.inspect(db.engine)
        columnas_tesoreria = {
            c["name"] for c in inspector.get_columns("movimiento_tesoreria")
        }
        columnas_nuevas_tesoreria = {
            "club_id": "INTEGER",
            "campeonato_id": "INTEGER",
            "serie": "VARCHAR(80)",
            "cuenta_id": "INTEGER",
        }

        for nombre, tipo_columna in columnas_nuevas_tesoreria.items():
            if nombre in columnas_tesoreria:
                continue

            if db.engine.dialect.name == "postgresql":
                sql = (
                    "ALTER TABLE movimiento_tesoreria "
                    f"ADD COLUMN IF NOT EXISTS {nombre} {tipo_columna}"
                )
            elif db.engine.dialect.name == "sqlite":
                sql = (
                    "ALTER TABLE movimiento_tesoreria "
                    f"ADD COLUMN {nombre} {tipo_columna}"
                )
            else:
                raise RuntimeError(
                    "Motor de base de datos no soportado para Tesorería."
                )

            db.session.execute(db.text(sql))
            db.session.commit()
            inspector = db.inspect(db.engine)
            columnas_tesoreria = {
                c["name"]
                for c in inspector.get_columns("movimiento_tesoreria")
            }

        # V10.2 — trazabilidad de cuentas generadas desde partidos.
        inspector = db.inspect(db.engine)
        columnas_cuentas = {
            c["name"] for c in inspector.get_columns("cuenta_tesoreria")
        }
        columnas_nuevas_cuentas = {
            "partido_id": "INTEGER",
            "origen": "VARCHAR(40)",
        }

        for nombre, tipo_columna in columnas_nuevas_cuentas.items():
            if nombre in columnas_cuentas:
                continue

            if db.engine.dialect.name == "postgresql":
                sql = (
                    "ALTER TABLE cuenta_tesoreria "
                    f"ADD COLUMN IF NOT EXISTS {nombre} {tipo_columna}"
                )
            elif db.engine.dialect.name == "sqlite":
                sql = (
                    "ALTER TABLE cuenta_tesoreria "
                    f"ADD COLUMN {nombre} {tipo_columna}"
                )
            else:
                raise RuntimeError(
                    "Motor de base de datos no soportado para Tesorería."
                )

            db.session.execute(db.text(sql))
            db.session.commit()
            inspector = db.inspect(db.engine)
            columnas_cuentas = {
                c["name"] for c in inspector.get_columns("cuenta_tesoreria")
            }

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
        admin = AdminUser(username=username, nombre=nombre, activo=True, rol="Administrador")
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
    "publico_disciplina",
    "publico_resolucion_disciplina",
    "health",
    "static",
}


@app.before_request
def exigir_login_administrativo():
    endpoint = request.endpoint
    if endpoint in PUBLIC_ENDPOINTS or (request.path or "").startswith("/static/"):
        return None
    if session.get("admin_id"):
        rol = session.get("admin_rol") or "Administrador"
        if rol != "Administrador":
            permitidos = ROLE_ENDPOINTS.get(rol, set())
            if endpoint not in permitidos and endpoint not in {"login", "logout"}:
                flash("Tu perfil no tiene permisos para acceder a esta sección.", "error")
                return redirect(url_for("dashboard"))
        return None
    destino = request.full_path.rstrip("?")
    return redirect(url_for("login", next=destino))


# ============================================================
# PERFILES DE ACCESO
# ============================================================

ROLES = {"Administrador": "Administrador", "Disciplina": "Disciplina", "Tesoreria": "Tesorería"}

ROLE_ENDPOINTS = {
    # Disciplina trabaja exclusivamente con las actas y sus registros.
    "Disciplina": {"dashboard","admin_panel_maestro","mi_cuenta_admin","admin_actas","acta_partido","admin_centro_actas","admin_disciplina","crear_resolucion_disciplina","ver_resolucion_disciplina","publicar_resolucion_disciplina","registrar_gol","registrar_amarilla","registrar_roja","registrar_suspension","eliminar_gol","eliminar_registro_disciplinario","logout"},
    "Tesoreria": {"dashboard","admin_panel_maestro","mi_cuenta_admin","admin_tesoreria","logout"}
}

def rol_permitido(*roles):
    def decorator(view):
        @wraps(view)
        def wrapped(*args, **kwargs):
            if not session.get("admin_id"):
                return redirect(url_for("login", next=request.full_path))
            rol = session.get("admin_rol") or "Administrador"
            if rol not in roles:
                flash("Tu perfil no tiene permisos para acceder a esta sección.", "error")
                return redirect(url_for("dashboard"))
            return view(*args, **kwargs)
        return wrapped
    return decorator

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
            session["admin_rol"] = admin.rol or "Administrador"
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
@rol_permitido("Administrador")
def admin_usuarios():
    if request.method == "POST":
        username = request.form.get("username", "").strip()
        nombre = request.form.get("nombre", "").strip() or "Administrador"
        password = request.form.get("password", "")
        rol = request.form.get("rol", "Administrador").strip()
        if rol not in ROLES:
            rol = "Administrador"
        if len(username) < 3 or len(password) < 8:
            flash("El usuario debe tener al menos 3 caracteres y la contraseña 8.", "error")
        elif AdminUser.query.filter_by(username=username).first():
            flash("Ese usuario administrador ya existe.", "error")
        else:
            admin = AdminUser(username=username, nombre=nombre, activo=True, rol=rol)
            admin.set_password(password)
            db.session.add(admin)
            db.session.commit()
            flash("Administrador creado correctamente.", "success")
            return redirect(url_for("admin_usuarios"))
    usuarios = AdminUser.query.order_by(AdminUser.username).all()
    return render_template("admin_usuarios.html", usuarios=usuarios, roles=ROLES)


@app.route("/admin/usuarios/<int:admin_id>/estado", methods=["POST"])
@rol_permitido("Administrador")
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
    # V9.0: si el administrador entra a la raíz del sistema,
    # llevarlo siempre al Inicio de gestión y no al antiguo registro
    # de jugadores. El acceso público sin sesión conserva su comportamiento.
    if session.get("admin_id"):
        return redirect(url_for("admin_panel_maestro"))

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
@rol_permitido("Administrador")
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
@rol_permitido("Administrador")
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
@rol_permitido("Administrador")
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
@rol_permitido("Administrador")
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
# V11.0 — CENTRO DE DISCIPLINA / COMITÉ
# ============================================================

@app.route("/admin/disciplina")
@rol_permitido("Administrador", "Disciplina")
def admin_disciplina():
    estado = (request.args.get("estado") or "").strip()
    serie = (request.args.get("serie") or "").strip()
    campeonato_id = request.args.get("campeonato_id", type=int)

    query = ResolucionDisciplina.query
    if estado:
        query = query.filter_by(estado=estado)
    if serie:
        query = query.filter(ResolucionDisciplina.serie == serie)
    if campeonato_id:
        query = query.filter(ResolucionDisciplina.campeonato_id == campeonato_id)

    resoluciones = query.order_by(
        ResolucionDisciplina.fecha.desc(),
        ResolucionDisciplina.id.desc()
    ).limit(200).all()

    campeonatos = Campeonato.query.order_by(Campeonato.temporada.desc(), Campeonato.id.desc()).all()
    series = Serie.query.filter_by(activo=True).order_by(Serie.nombre).all()

    total = ResolucionDisciplina.query.count()
    publicadas = ResolucionDisciplina.query.filter_by(publicado=True).count()
    pendientes = ResolucionDisciplina.query.filter(ResolucionDisciplina.estado != "Cerrada").count()
    sanciones = ResolucionDisciplina.query.filter(
        ResolucionDisciplina.sancion.isnot(None),
        ResolucionDisciplina.sancion != ""
    ).count()

    return render_template(
        "admin_disciplina.html",
        resoluciones=resoluciones,
        campeonatos=campeonatos,
        series=series,
        estado=estado,
        serie=serie,
        campeonato_id=campeonato_id,
        resumen={"total": total, "publicadas": publicadas, "pendientes": pendientes, "sanciones": sanciones},
    )


@app.route("/admin/disciplina/resolucion/nueva", methods=["GET", "POST"])
@rol_permitido("Administrador", "Disciplina")
def crear_resolucion_disciplina():
    if request.method == "GET":
        return render_template(
            "admin_disciplina_form.html",
            resolucion=None,
            campeonatos=Campeonato.query.order_by(Campeonato.temporada.desc(), Campeonato.id.desc()).all(),
            clubes=Club.query.filter_by(activo=True).order_by(Club.nombre).all(),
            jugadores=Jugador.query.order_by(Jugador.club, Jugador.nombre_completo).all(),
            partidos=Partido.query.order_by(Partido.fecha.desc().nullslast(), Partido.id.desc()).limit(300).all(),
        )

    try:
        fecha = datetime.strptime(request.form.get("fecha") or date.today().isoformat(), "%Y-%m-%d").date()
        fecha_inicio_raw = request.form.get("fecha_inicio") or ""
        fecha_fin_raw = request.form.get("fecha_fin") or ""
        fecha_inicio = datetime.strptime(fecha_inicio_raw, "%Y-%m-%d").date() if fecha_inicio_raw else None
        fecha_fin = datetime.strptime(fecha_fin_raw, "%Y-%m-%d").date() if fecha_fin_raw else None

        jugador_id = request.form.get("jugador_id", type=int) or None
        club_id = request.form.get("club_id", type=int) or None
        campeonato_id = request.form.get("campeonato_id", type=int) or None
        partido_id = request.form.get("partido_id", type=int) or None
        cantidad = request.form.get("cantidad_sancion", type=int) or None
        aplicar = request.form.get("aplicar_sancion") == "1"

        ultimo = ResolucionDisciplina.query.order_by(ResolucionDisciplina.id.desc()).first()
        numero = f"CD-{fecha.year}-{(ultimo.id + 1 if ultimo else 1):04d}"

        resolucion = ResolucionDisciplina(
            numero_resolucion=numero,
            fecha=fecha,
            campeonato_id=campeonato_id,
            partido_id=partido_id,
            jugador_id=jugador_id,
            club_id=club_id,
            serie=(request.form.get("serie") or "").strip() or None,
            tipo_decision=(request.form.get("tipo_decision") or "Resolución").strip(),
            titulo=(request.form.get("titulo") or "").strip(),
            antecedentes=(request.form.get("antecedentes") or "").strip(),
            resolucion=(request.form.get("resolucion") or "").strip(),
            sancion=(request.form.get("sancion") or "").strip() or None,
            cantidad_sancion=cantidad,
            unidad_sancion=(request.form.get("unidad_sancion") or "").strip() or None,
            fecha_inicio=fecha_inicio,
            fecha_fin=fecha_fin,
            observaciones=(request.form.get("observaciones") or "").strip(),
            estado="Cerrada" if request.form.get("guardar_cerrar") == "1" else "Borrador",
            publicado=False,
            creado_por=session.get("admin_nombre") or "Administrador",
        )
        db.session.add(resolucion)
        db.session.flush()

        # Una resolución puede generar automáticamente un registro disciplinario.
        if aplicar and jugador_id:
            tipo_sancion = (request.form.get("tipo_registro") or "Suspension").strip()
            registro = RegistroDisciplinario(
                jugador_id=jugador_id,
                fecha=fecha,
                tipo=tipo_sancion,
                cantidad=max(cantidad or 1, 1),
                motivo=resolucion.titulo,
                campeonato=(resolucion.campeonato.nombre if resolucion.campeonato else ""),
                observaciones=f"Resolución {resolucion.numero_resolucion}: {resolucion.resolucion}",
                campeonato_id=campeonato_id,
            )
            db.session.add(registro)
            if tipo_sancion == "Suspension":
                jugador = db.session.get(Jugador, jugador_id)
                if jugador:
                    jugador.estado = "Suspendido"
            elif tipo_sancion == "Inhabilitacion":
                jugador = db.session.get(Jugador, jugador_id)
                if jugador:
                    jugador.estado = "Inhabilitado"

        db.session.commit()
        flash(f"Resolución {resolucion.numero_resolucion} registrada correctamente.", "success")
        return redirect(url_for("ver_resolucion_disciplina", resolucion_id=resolucion.id))
    except Exception as error:
        db.session.rollback()
        app.logger.exception("Error creando resolución disciplinaria")
        flash(f"No fue posible guardar la resolución: {error}", "error")
        return redirect(url_for("crear_resolucion_disciplina"))


@app.route("/admin/disciplina/resolucion/<int:resolucion_id>")
@rol_permitido("Administrador", "Disciplina")
def ver_resolucion_disciplina(resolucion_id):
    resolucion = db.get_or_404(ResolucionDisciplina, resolucion_id)
    return render_template("admin_disciplina_detalle.html", resolucion=resolucion)


@app.route("/admin/disciplina/resolucion/<int:resolucion_id>/publicar", methods=["POST"])
@rol_permitido("Administrador", "Disciplina")
def publicar_resolucion_disciplina(resolucion_id):
    resolucion = db.get_or_404(ResolucionDisciplina, resolucion_id)
    if not (resolucion.titulo and resolucion.resolucion):
        flash("La resolución debe tener título y decisión antes de publicarse.", "error")
        return redirect(url_for("ver_resolucion_disciplina", resolucion_id=resolucion.id))
    resolucion.publicado = True
    resolucion.estado = "Cerrada"
    resolucion.fecha_publicacion = date.today()
    resolucion.actualizado_por = session.get("admin_nombre") or "Administrador"
    resolucion.actualizado_en = datetime.utcnow()
    db.session.commit()
    flash(f"La resolución {resolucion.numero_resolucion} fue publicada en el portal público.", "success")
    return redirect(url_for("ver_resolucion_disciplina", resolucion_id=resolucion.id))


@app.route("/admin/disciplina/resolucion/<int:resolucion_id>/anular", methods=["POST"])
@rol_permitido("Administrador")
def anular_resolucion_disciplina(resolucion_id):
    resolucion = db.get_or_404(ResolucionDisciplina, resolucion_id)
    resolucion.estado = "Anulada"
    resolucion.publicado = False
    resolucion.actualizado_por = session.get("admin_nombre") or "Administrador"
    resolucion.actualizado_en = datetime.utcnow()
    db.session.commit()
    flash(f"La resolución {resolucion.numero_resolucion} fue anulada y retirada del portal público.", "warning")
    return redirect(url_for("ver_resolucion_disciplina", resolucion_id=resolucion.id))


# ============================================================
# V11.0 — PORTAL PÚBLICO DE DISCIPLINA
# ============================================================

@app.route("/disciplina")
def publico_disciplina():
    resoluciones = ResolucionDisciplina.query.filter_by(
        publicado=True
    ).filter(
        ResolucionDisciplina.estado != "Anulada"
    ).order_by(
        ResolucionDisciplina.fecha_publicacion.desc(),
        ResolucionDisciplina.id.desc()
    ).limit(100).all()
    return render_template("publico_disciplina.html", resoluciones=resoluciones)


@app.route("/disciplina/resolucion/<int:resolucion_id>")
def publico_resolucion_disciplina(resolucion_id):
    resolucion = db.get_or_404(ResolucionDisciplina, resolucion_id)
    if not resolucion.publicado or resolucion.estado == "Anulada":
        return redirect(url_for("publico_disciplina"))
    return render_template("publico_resolucion_disciplina.html", resolucion=resolucion)


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
# REGISTRO INTERNO DE CLUBES, SERIES Y PLANTELES
# ============================================================

def registrar_movimiento_jugador(jugador, tipo, anterior=None, motivo=None):
    """Registra cambios del registro maestro sin duplicar al jugador."""
    anterior = anterior or {}
    movimiento = JugadorMovimiento(
        jugador_id=jugador.id,
        tipo=tipo,
        club_anterior=anterior.get("club"),
        serie_anterior=anterior.get("serie"),
        estado_anterior=anterior.get("estado"),
        club_nuevo=jugador.club,
        serie_nuevo=jugador.serie,
        estado_nuevo=jugador.estado,
        motivo=motivo,
        realizado_por=(session.get("admin_username") or session.get("admin_id") or "Administrador"),
    )
    db.session.add(movimiento)


@app.route("/admin/planteles")
def admin_planteles():
    """Registro interno de jugadores organizado por club y serie.

    No crea nuevas tablas: utiliza el registro Jugador existente, por lo que
    es compatible con la base de datos actual de PostgreSQL/Railway.
    """
    q = request.args.get("q", "").strip()
    club_filtro = request.args.get("club", "").strip()
    serie_filtro = request.args.get("serie", "").strip()
    estado_filtro = request.args.get("estado", "").strip()

    query = Jugador.query

    if q:
        like = f"%{q}%"
        query = query.filter(
            db.or_(
                Jugador.nombre_completo.ilike(like),
                Jugador.rut.ilike(like),
                Jugador.club.ilike(like),
                Jugador.serie.ilike(like),
            )
        )

    if club_filtro:
        query = query.filter(Jugador.club == club_filtro)

    if serie_filtro:
        query = query.filter(Jugador.serie == serie_filtro)

    if estado_filtro:
        query = query.filter(Jugador.estado == estado_filtro)

    jugadores = query.order_by(
        Jugador.club.asc(),
        Jugador.serie.asc(),
        Jugador.nombre_completo.asc(),
    ).all()

    # Agrupación: Club -> Serie -> Jugadores
    planteles = {}
    for jugador in jugadores:
        club = jugador.club or "Sin club"
        serie = jugador.serie or "Sin serie"
        planteles.setdefault(club, {}).setdefault(serie, []).append(jugador)

    clubes = [
        row[0]
        for row in db.session.query(Jugador.club)
        .filter(Jugador.club.isnot(None), Jugador.club != "")
        .distinct()
        .order_by(Jugador.club.asc())
        .all()
    ]

    series = [
        row[0]
        for row in db.session.query(Jugador.serie)
        .filter(Jugador.serie.isnot(None), Jugador.serie != "")
        .distinct()
        .order_by(Jugador.serie.asc())
        .all()
    ]

    estados = [
        row[0]
        for row in db.session.query(Jugador.estado)
        .filter(Jugador.estado.isnot(None), Jugador.estado != "")
        .distinct()
        .order_by(Jugador.estado.asc())
        .all()
    ]

    return render_template(
        "admin_planteles.html",
        planteles=planteles,
        clubes=clubes,
        series=series,
        estados=estados,
        q=q,
        club_filtro=club_filtro,
        serie_filtro=serie_filtro,
        estado_filtro=estado_filtro,
        total_jugadores=len(jugadores),
        total_clubes=len(planteles),
        total_series=len({serie for club in planteles.values() for serie in club}),
    )



# ============================================================
# V6.8 — CENTRO DE CONTROL DE CLUBES Y PLANTELES
# ============================================================

@app.route("/admin/clubes/centro")
def admin_centro_clubes():
    """Centro consolidado de clubes, planteles y participación deportiva."""
    q = request.args.get("q", "").strip()
    query = Club.query
    if q:
        query = query.filter(Club.nombre.ilike(f"%{q}%"))
    clubes = query.order_by(Club.nombre.asc()).all()

    jugadores = Jugador.query.all()
    por_club = {}
    for j in jugadores:
        por_club.setdefault((j.club or "").strip(), []).append(j)

    tarjetas = []
    for club in clubes:
        js = por_club.get(club.nombre.strip(), [])
        vigentes = sum(1 for j in js if (j.estado or "Vigente") == "Vigente")
        series = sorted({j.serie for j in js if j.serie})
        campeonatos = CampeonatoClub.query.filter_by(club_id=club.id).count()
        partidos = Partido.query.filter(
            db.or_(Partido.local_club_id == club.id, Partido.visitante_club_id == club.id)
        ).count()
        tarjetas.append({
            "club": club,
            "total": len(js),
            "vigentes": vigentes,
            "series": series,
            "campeonatos": campeonatos,
            "partidos": partidos,
        })

    return render_template(
        "admin_centro_clubes.html",
        tarjetas=tarjetas,
        q=q,
        total_clubes=len(clubes),
        total_jugadores=sum(t["total"] for t in tarjetas),
        total_vigentes=sum(t["vigentes"] for t in tarjetas),
        total_partidos=sum(t["partidos"] for t in tarjetas),
    )


@app.route("/admin/clubes/<int:club_id>/centro")
def admin_centro_club(club_id):
    """Centro operativo de un club usando el registro maestro existente."""
    club = db.get_or_404(Club, club_id)
    jugadores = Jugador.query.filter(Jugador.club == club.nombre).order_by(
        Jugador.serie.asc(), Jugador.nombre_completo.asc()
    ).all()

    planteles = {}
    for j in jugadores:
        planteles.setdefault(j.serie or "Sin serie", []).append(j)

    participaciones = CampeonatoClub.query.filter_by(club_id=club.id).all()
    campeonatos = [x.campeonato for x in participaciones if x.campeonato]

    partidos = Partido.query.filter(
        db.or_(Partido.local_club_id == club.id, Partido.visitante_club_id == club.id)
    ).order_by(Partido.fecha.desc().nullslast(), Partido.id.desc()).all()

    jugados = [p for p in partidos if p.goles_local is not None and p.goles_visitante is not None]
    victorias = empates = derrotas = goles_favor = goles_contra = 0
    for p in jugados:
        es_local = p.local_club_id == club.id
        gf = p.goles_local if es_local else p.goles_visitante
        gc = p.goles_visitante if es_local else p.goles_local
        goles_favor += int(gf or 0)
        goles_contra += int(gc or 0)
        if gf > gc: victorias += 1
        elif gf == gc: empates += 1
        else: derrotas += 1

    actas_pendientes = sum(
        1 for p in partidos if not p.acta or (p.acta.estado or "Borrador") != "Cerrada"
    )

    return render_template(
        "admin_centro_club.html",
        club=club,
        jugadores=jugadores,
        planteles=planteles,
        campeonatos=campeonatos,
        partidos=partidos[:15],
        total_jugadores=len(jugadores),
        total_vigentes=sum(1 for j in jugadores if (j.estado or "Vigente") == "Vigente"),
        total_series=len(planteles),
        total_partidos=len(partidos),
        jugados=len(jugados),
        victorias=victorias,
        empates=empates,
        derrotas=derrotas,
        goles_favor=goles_favor,
        goles_contra=goles_contra,
        actas_pendientes=actas_pendientes,
    )

@app.route("/admin/clubes/<int:club_id>/series")
@admin_required
def admin_centro_club_series(club_id):
    """V7.4.13 — Series del club conectadas al registro maestro."""
    club = db.get_or_404(Club, club_id)

    jugadores = Jugador.query.filter(
        Jugador.club == club.nombre
    ).order_by(
        Jugador.serie.asc(), Jugador.nombre_completo.asc()
    ).all()

    participaciones = CampeonatoClub.query.filter_by(
        club_id=club.id
    ).all()

    campeonatos = sorted(
        [x.campeonato for x in participaciones if x.campeonato],
        key=lambda x: (x.temporada or "", x.id),
        reverse=True
    )

    series_map = {}
    for jugador in jugadores:
        nombre_serie = (jugador.serie or "Sin serie").strip()
        fila = series_map.setdefault(nombre_serie, {
            "nombre": nombre_serie,
            "jugadores": [],
            "campeonatos": [],
        })
        fila["jugadores"].append(jugador)

    for campeonato in campeonatos:
        nombre_serie = (campeonato.serie or "Sin serie").strip()
        fila = series_map.setdefault(nombre_serie, {
            "nombre": nombre_serie,
            "jugadores": [],
            "campeonatos": [],
        })
        fila["campeonatos"].append(campeonato)

    filas = list(series_map.values())
    for fila in filas:
        ids_campeonatos = {c.id for c in fila["campeonatos"]}
        stats = {
            "partidos": 0,
            "goles": 0,
            "amarillas": 0,
            "rojas": 0,
        }

        if ids_campeonatos and fila["jugadores"]:
            jugador_ids = [j.id for j in fila["jugadores"]]
            registros = (
                db.session.query(PartidoJugador, Partido, ActaPartido)
                .join(Partido, Partido.id == PartidoJugador.partido_id)
                .join(ActaPartido, ActaPartido.partido_id == Partido.id)
                .filter(
                    PartidoJugador.jugador_id.in_(jugador_ids),
                    Partido.campeonato_id.in_(ids_campeonatos),
                    ActaPartido.estado == "Cerrada",
                )
                .all()
            )
            partidos_ids = set()
            for pj, partido, acta in registros:
                partidos_ids.add(partido.id)
                stats["goles"] += int(pj.goles or 0)
                stats["amarillas"] += int(pj.amarillas or 0)
                stats["rojas"] += int(pj.rojas or 0)
            stats["partidos"] = len(partidos_ids)

        fila["vigentes"] = sum(
            1 for j in fila["jugadores"]
            if (j.estado or "Vigente") == "Vigente"
        )
        fila["stats"] = stats

    filas.sort(key=lambda x: x["nombre"].lower())

    return render_template(
        "admin_centro_club_series.html",
        club=club,
        filas=filas,
        total_series=len(filas),
        total_jugadores=len(jugadores),
        total_vigentes=sum(f["vigentes"] for f in filas),
    )

# ============================================================
# V7.4.3 — ESTADÍSTICAS INTEGRADAS POR CLUB
# Usa exclusivamente PartidoJugador + ActaPartido cerrada.
# No crea registros ni tablas paralelas.
# ============================================================

@app.route("/admin/clubes/<int:club_id>/estadisticas")
def admin_centro_club_estadisticas(club_id):
    club = db.get_or_404(Club, club_id)

    campeonato_id = request.args.get("campeonato_id", type=int)

    participaciones = CampeonatoClub.query.filter_by(club_id=club.id).all()
    campeonatos = [x.campeonato for x in participaciones if x.campeonato]
    campeonatos = sorted(campeonatos, key=lambda x: (x.temporada or "", x.id), reverse=True)

    base = (
        db.session.query(PartidoJugador, Jugador, Partido, Campeonato, ActaPartido)
        .join(Jugador, Jugador.id == PartidoJugador.jugador_id)
        .join(Partido, Partido.id == PartidoJugador.partido_id)
        .join(Campeonato, Campeonato.id == Partido.campeonato_id)
        .join(ActaPartido, ActaPartido.partido_id == Partido.id)
        .filter(ActaPartido.estado == "Cerrada")
        .filter(Jugador.club == club.nombre)
    )

    if campeonato_id:
        base = base.filter(Campeonato.id == campeonato_id)

    registros = base.all()

    partidos_ids = set()
    jugadores_ids = set()
    total_goles = total_amarillas = total_rojas = 0
    por_jugador = {}

    for pj, jugador, partido, campeonato, acta in registros:
        partidos_ids.add(partido.id)
        jugadores_ids.add(jugador.id)

        goles = int(pj.goles or 0)
        amarillas = int(pj.amarillas or 0)
        rojas = int(pj.rojas or 0)
        total_goles += goles
        total_amarillas += amarillas
        total_rojas += rojas

        key = (jugador.id, campeonato.id)
        if key not in por_jugador:
            por_jugador[key] = {
                "jugador": jugador,
                "campeonato": campeonato,
                "partidos": set(),
                "goles": 0,
                "amarillas": 0,
                "rojas": 0,
            }

        fila = por_jugador[key]
        fila["partidos"].add(partido.id)
        fila["goles"] += goles
        fila["amarillas"] += amarillas
        fila["rojas"] += rojas

    jugadores_stats = []
    for fila in por_jugador.values():
        fila["partidos_jugados"] = len(fila["partidos"])
        jugadores_stats.append(fila)

    goleadores = sorted(
        jugadores_stats,
        key=lambda x: (-x["goles"], x["jugador"].nombre_completo.lower())
    )
    disciplina = sorted(
        jugadores_stats,
        key=lambda x: (-x["amarillas"], -x["rojas"], x["jugador"].nombre_completo.lower())
    )

    goleadores = [x for x in goleadores if x["goles"] > 0]
    disciplina = [x for x in disciplina if x["amarillas"] > 0 or x["rojas"] > 0]

    return render_template(
        "admin_centro_club_estadisticas.html",
        club=club,
        campeonatos=campeonatos,
        campeonato_id=campeonato_id,
        goleadores=goleadores,
        disciplina=disciplina,
        resumen={
            "partidos": len(partidos_ids),
            "jugadores": len(jugadores_ids),
            "goles": total_goles,
            "amarillas": total_amarillas,
            "rojas": total_rojas,
        },
    )


# ============================================================
# V6.3 — CENTRO DE REGISTRO Y DIRECTORIO DE CLUBES
# ============================================================

@app.route("/admin/registro")
def admin_registro():
    """Centro único para altas e importación de jugadores."""
    total_jugadores = Jugador.query.count()
    total_vigentes = Jugador.query.filter_by(estado="Vigente").count()
    total_clubes = Club.query.filter_by(activo=True).count()
    total_series = Serie.query.filter_by(activo=True).count()

    return render_template(
        "admin_registro.html",
        total_jugadores=total_jugadores,
        total_vigentes=total_vigentes,
        total_clubes=total_clubes,
        total_series=total_series,
    )



# ============================================================
# V6.9 — CENTRO MAESTRO DE JUGADORES
# Consulta unificada del registro Jugador existente.
# No crea una segunda base ni modifica el centro individual.
# ============================================================

@app.route("/admin/jugadores/centro")
def admin_centro_jugadores():
    q = request.args.get("q", "").strip()
    club_filtro = request.args.get("club", "").strip()
    serie_filtro = request.args.get("serie", "").strip()
    estado_filtro = request.args.get("estado", "").strip()

    query = Jugador.query

    if q:
        like = f"%{q}%"
        query = query.filter(
            db.or_(
                Jugador.nombre_completo.ilike(like),
                Jugador.rut.ilike(like),
            )
        )

    if club_filtro:
        query = query.filter(Jugador.club == club_filtro)

    if serie_filtro:
        query = query.filter(Jugador.serie == serie_filtro)

    if estado_filtro:
        query = query.filter(Jugador.estado == estado_filtro)

    jugadores = query.order_by(
        Jugador.nombre_completo.asc()
    ).limit(250).all()

    clubes = [
        row[0]
        for row in db.session.query(Jugador.club)
        .filter(Jugador.club.isnot(None), Jugador.club != "")
        .distinct()
        .order_by(Jugador.club.asc())
        .all()
    ]

    series = [
        row[0]
        for row in db.session.query(Jugador.serie)
        .filter(Jugador.serie.isnot(None), Jugador.serie != "")
        .distinct()
        .order_by(Jugador.serie.asc())
        .all()
    ]

    estados = [
        row[0]
        for row in db.session.query(Jugador.estado)
        .filter(Jugador.estado.isnot(None), Jugador.estado != "")
        .distinct()
        .order_by(Jugador.estado.asc())
        .all()
    ]

    resumen = []
    for jugador in jugadores:
        partidos = (
            db.session.query(db.func.count(PartidoJugador.id))
            .filter(PartidoJugador.jugador_id == jugador.id)
            .scalar()
        ) or 0

        goles = (
            db.session.query(db.func.coalesce(db.func.sum(PartidoJugador.goles), 0))
            .filter(PartidoJugador.jugador_id == jugador.id)
            .scalar()
        ) or 0

        amarillas = (
            db.session.query(db.func.coalesce(db.func.sum(PartidoJugador.amarillas), 0))
            .filter(PartidoJugador.jugador_id == jugador.id)
            .scalar()
        ) or 0

        rojas = (
            db.session.query(db.func.coalesce(db.func.sum(PartidoJugador.rojas), 0))
            .filter(PartidoJugador.jugador_id == jugador.id)
            .scalar()
        ) or 0

        resumen.append({
            "jugador": jugador,
            "partidos": int(partidos),
            "goles": int(goles),
            "amarillas": int(amarillas),
            "rojas": int(rojas),
        })

    total = Jugador.query.count()
    vigentes = Jugador.query.filter_by(estado="Vigente").count()
    clubes_total = (
        db.session.query(Jugador.club)
        .filter(Jugador.club.isnot(None), Jugador.club != "")
        .distinct()
        .count()
    )
    participaciones = PartidoJugador.query.count()

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

@app.route("/admin/clubes")
def admin_clubes():
    """Directorio interno de todos los clubes, agrupando sus planteles por serie."""
    q = request.args.get("q", "").strip()
    club_id = request.args.get("club_id", "", type=int)

    query = Club.query
    if q:
        query = query.filter(Club.nombre.ilike(f"%{q}%"))
    if club_id:
        query = query.filter(Club.id == club_id)

    clubes = query.order_by(Club.nombre.asc()).all()
    jugadores = Jugador.query.order_by(
        Jugador.club.asc(), Jugador.serie.asc(), Jugador.nombre_completo.asc()
    ).all()

    agrupados = {}
    for jugador in jugadores:
        agrupados.setdefault(jugador.club or "Sin club", {}).setdefault(
            jugador.serie or "Sin serie", []
        ).append(jugador)

    tarjetas = []
    for club in clubes:
        series_map = agrupados.get(club.nombre, {})
        total = sum(len(lista) for lista in series_map.values())
        tarjetas.append({"club": club, "series": series_map, "total": total})

    return render_template(
        "admin_clubes.html",
        tarjetas=tarjetas,
        clubes=Club.query.order_by(Club.nombre.asc()).all(),
        q=q,
        club_id=club_id,
        total_clubes=Club.query.count(),
        total_jugadores=Jugador.query.count(),
        total_vigentes=Jugador.query.filter_by(estado="Vigente").count(),
    )


# ============================================================
# FICHA ADMINISTRATIVA DE CLUB
# ============================================================

@app.route("/admin/planteles/club/<path:club_nombre>")
def admin_ficha_club(club_nombre):
    """Ficha administrativa completa de un club.

    Agrupa el registro Jugador por serie y permite consultar rápidamente
    planteles, estados y acciones de gestión sin modificar el modelo de datos.
    """
    club_nombre = club_nombre.strip()
    q = request.args.get("q", "").strip()
    serie_filtro = request.args.get("serie", "").strip()

    query = Jugador.query.filter(Jugador.club == club_nombre)

    if q:
        like = f"%{q}%"
        query = query.filter(
            db.or_(
                Jugador.nombre_completo.ilike(like),
                Jugador.rut.ilike(like),
            )
        )

    if serie_filtro:
        query = query.filter(Jugador.serie == serie_filtro)

    jugadores = query.order_by(
        Jugador.serie.asc(),
        Jugador.nombre_completo.asc(),
    ).all()

    series_club = [
        row[0]
        for row in db.session.query(Jugador.serie)
        .filter(Jugador.club == club_nombre)
        .filter(Jugador.serie.isnot(None), Jugador.serie != "")
        .distinct()
        .order_by(Jugador.serie.asc())
        .all()
    ]

    # Agrupación Serie -> jugadores
    planteles = {}
    for jugador in jugadores:
        serie = jugador.serie or "Sin serie"
        planteles.setdefault(serie, []).append(jugador)

    todos = Jugador.query.filter(Jugador.club == club_nombre).all()
    total = len(todos)
    vigentes = sum(1 for j in todos if (j.estado or "Vigente") == "Vigente")
    pendientes = sum(1 for j in todos if j.estado == "Pendiente")
    suspendidos = sum(1 for j in todos if j.estado in {"Suspendido", "Inhabilitado"})

    resumen_series = []
    for serie in series_club:
        serie_jugadores = [j for j in todos if j.serie == serie]
        resumen_series.append({
            "nombre": serie,
            "total": len(serie_jugadores),
            "vigentes": sum(1 for j in serie_jugadores if (j.estado or "Vigente") == "Vigente"),
            "pendientes": sum(1 for j in serie_jugadores if j.estado == "Pendiente"),
            "suspendidos": sum(1 for j in serie_jugadores if j.estado in {"Suspendido", "Inhabilitado"}),
        })

    return render_template(
        "admin_club_ficha.html",
        club_nombre=club_nombre,
        planteles=planteles,
        series_club=series_club,
        resumen_series=resumen_series,
        total=total,
        vigentes=vigentes,
        pendientes=pendientes,
        suspendidos=suspendidos,
        q=q,
        serie_filtro=serie_filtro,
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


@app.route("/campeonatos/<int:campeonato_id>/panel")
@admin_required
def panel_campeonato(campeonato_id):
    """Panel estable del campeonato. Solo lectura sobre registros existentes."""
    campeonato = db.get_or_404(Campeonato, campeonato_id)

    clubes_participantes = (
        CampeonatoClub.query
        .filter_by(campeonato_id=campeonato.id)
        .join(Club, CampeonatoClub.club_id == Club.id)
        .order_by(Club.nombre.asc())
        .all()
    )

    partidos = (
        Partido.query
        .filter_by(campeonato_id=campeonato.id)
        .order_by(
            Partido.jornada.asc(),
            Partido.fecha.asc().nullslast(),
            Partido.hora.asc().nullslast(),
            Partido.id.asc(),
        )
        .all()
    )

    partidos_finalizados = [
        p for p in partidos
        if p.estado == "Finalizado"
        and p.goles_local is not None
        and p.goles_visitante is not None
    ]
    partidos_pendientes = [p for p in partidos if p.estado != "Finalizado"]

    tabla = {}
    for registro in clubes_participantes:
        club = registro.club
        tabla[club.id] = {
            "club": club, "pj": 0, "pg": 0, "pe": 0, "pp": 0,
            "gf": 0, "gc": 0, "dg": 0, "pts": 0
        }

    for partido in partidos_finalizados:
        if partido.local_club_id not in tabla or partido.visitante_club_id not in tabla:
            continue
        gl = int(partido.goles_local or 0)
        gv = int(partido.goles_visitante or 0)
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
        elif gv > gl:
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
    filas.sort(key=lambda f: (-f["pts"], -f["dg"], -f["gf"], f["club"].nombre.lower()))
    for pos, fila in enumerate(filas, 1):
        fila["pos"] = pos

    proximos = [p for p in partidos if p.estado != "Finalizado"]
    proximos.sort(key=lambda p: (p.fecha is None, p.fecha or date.max, p.jornada, p.id))
    proximo_partido = proximos[0] if proximos else None

    # Estadísticas oficiales: solo participaciones de actas cerradas.
    registros = (
        db.session.query(PartidoJugador, Jugador, Partido)
        .join(Jugador, Jugador.id == PartidoJugador.jugador_id)
        .join(Partido, Partido.id == PartidoJugador.partido_id)
        .join(ActaPartido, ActaPartido.partido_id == Partido.id)
        .filter(
            Partido.campeonato_id == campeonato.id,
            ActaPartido.estado == "Cerrada",
        )
        .all()
    )

    por_jugador = {}
    por_club = {}
    for pj, jugador, partido in registros:
        j = por_jugador.setdefault(
            jugador.id,
            {"nombre": jugador.nombre_completo, "goles": 0, "amarillas": 0, "rojas": 0},
        )
        j["goles"] += int(pj.goles or 0)
        j["amarillas"] += int(pj.amarillas or 0)
        j["rojas"] += int(pj.rojas or 0)

        nombre_club = (jugador.club or "Sin club").strip() or "Sin club"
        cc = por_club.setdefault(
            nombre_club,
            {"nombre": nombre_club, "amarillas": 0, "rojas": 0},
        )
        cc["amarillas"] += int(pj.amarillas or 0)
        cc["rojas"] += int(pj.rojas or 0)

    goleadores = sorted(
        [(x["nombre"], x["goles"]) for x in por_jugador.values() if x["goles"] > 0],
        key=lambda x: (-x[1], x[0].lower()),
    )[:8]

    total_amarillas = sum(x["amarillas"] for x in por_jugador.values())
    total_rojas = sum(x["rojas"] for x in por_jugador.values())

    total_suspensiones = (
        db.session.query(db.func.coalesce(db.func.sum(RegistroDisciplinario.cantidad), 0))
        .filter(
            RegistroDisciplinario.campeonato_id == campeonato.id,
            RegistroDisciplinario.tipo == "Suspension",
        )
        .scalar()
        or 0
    )

    # Datos simples y serializables para las funciones interactivas del panel.
    campaign_matches = [
        {
            "id": p.id,
            "jornada": p.jornada or 0,
            "local_id": p.local_club_id,
            "local": p.local_club.nombre,
            "visitante_id": p.visitante_club_id,
            "visitante": p.visitante_club.nombre,
            "gl": p.goles_local,
            "gv": p.goles_visitante,
            "estado": p.estado,
        }
        for p in partidos
    ]

    fair_play = sorted(
        [
            {
                "nombre": x["nombre"],
                "amarillas": x["amarillas"],
                "rojas": x["rojas"],
                "puntos": x["amarillas"] + x["rojas"] * 3,
            }
            for x in por_club.values()
        ],
        key=lambda x: (x["puntos"], x["nombre"].lower()),
    )

    # Incluir también clubes sin tarjetas.
    existentes = {x["nombre"] for x in fair_play}
    for registro in clubes_participantes:
        if registro.club.nombre not in existentes:
            fair_play.append({
                "nombre": registro.club.nombre,
                "amarillas": 0,
                "rojas": 0,
                "puntos": 0,
            })
    fair_play.sort(key=lambda x: (x["puntos"], x["nombre"].lower()))

    porcentaje = round((len(partidos_finalizados) / len(partidos)) * 100) if partidos else 0
    campaign_clubs = [{"id": r.club.id, "nombre": r.club.nombre} for r in clubes_participantes]

    return render_template(
        "campeonato_panel.html",
        campeonato=campeonato,
        clubes_participantes=clubes_participantes,
        partidos=partidos,
        partidos_finalizados=partidos_finalizados,
        partidos_pendientes=partidos_pendientes,
        porcentaje=porcentaje,
        filas=filas,
        proximo_partido=proximo_partido,
        goleadores=goleadores,
        total_amarillas=int(total_amarillas),
        total_rojas=int(total_rojas),
        total_suspensiones=int(total_suspensiones),
        campaign_matches=campaign_matches,
        campaign_clubs=campaign_clubs,
        fair_play=fair_play,
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
    "/campeonatos/<int:campeonato_id>/club/<int:club_id>/retirar",
    methods=["POST"]
)
@rol_permitido("Administrador")
def retirar_club_campeonato(campeonato_id, club_id):
    """Da de baja un club y reorganiza automáticamente todo el fixture pendiente."""
    campeonato = db.get_or_404(Campeonato, campeonato_id)
    club = db.get_or_404(Club, club_id)

    participacion = CampeonatoClub.query.filter_by(
        campeonato_id=campeonato.id, club_id=club.id
    ).first()

    if not participacion:
        flash("El club no está inscrito en este campeonato.", "error")
        return redirect(url_for("detalle_campeonato", campeonato_id=campeonato.id))

    try:
        partidos = Partido.query.filter_by(campeonato_id=campeonato.id).order_by(
            Partido.jornada.asc(), Partido.id.asc()
        ).all()

        finalizados = [
            p for p in partidos
            if (p.estado or "").strip().lower() == "finalizado"
            and p.goles_local is not None and p.goles_visitante is not None
        ]

        for p in partidos:
            if p not in finalizados:
                db.session.delete(p)

        db.session.delete(participacion)
        db.session.flush()

        clubes_actuales = [
            r.club_id for r in CampeonatoClub.query.filter_by(
                campeonato_id=campeonato.id
            ).order_by(CampeonatoClub.id).all()
        ]

        cruces_jugados = {
            frozenset((p.local_club_id, p.visitante_club_id))
            for p in finalizados
            if p.local_club_id in clubes_actuales and p.visitante_club_id in clubes_actuales
        }

        calendario = generar_calendario_todos_contra_todos(clubes_actuales)
        calendario_pendiente = []
        for jornada in calendario:
            nuevos = [
                par for par in jornada
                if frozenset(par) not in cruces_jugados
            ]
            if nuevos:
                calendario_pendiente.append(nuevos)

        ultima_jornada = max((p.jornada for p in finalizados), default=0)
        ultima_fecha = max((p.fecha for p in finalizados if p.fecha), default=None)
        primera_fecha = (
            ultima_fecha + timedelta(days=7)
            if ultima_fecha else siguiente_sabado(campeonato.fecha_inicio)
        )

        creados = 0
        for offset, jornada_partidos in enumerate(calendario_pendiente):
            jornada_num = ultima_jornada + offset + 1
            fecha_jornada = primera_fecha + timedelta(days=offset * 7)
            for local_id, visitante_id in jornada_partidos:
                db.session.add(Partido(
                    campeonato_id=campeonato.id, jornada=jornada_num,
                    fecha=fecha_jornada, hora="17:00", cancha="Por definir",
                    local_club_id=local_id, visitante_club_id=visitante_id,
                    turno_club_id=None, goles_local=None,
                    goles_visitante=None, estado="Programado"
                ))
                creados += 1

        db.session.commit()
        flash(
            f"{club.nombre} fue dado de baja. Se conservaron {len(finalizados)} "
            f"partidos finalizados y se reorganizaron {creados} partidos pendientes.",
            "success"
        )
    except Exception as error:
        db.session.rollback()
        print("ERROR RETIRANDO CLUB DEL CAMPEONATO:", repr(error))
        flash("No fue posible dar de baja el club ni reorganizar el fixture.", "error")

    return redirect(url_for("detalle_campeonato", campeonato_id=campeonato.id))


@app.route(
    "/campeonatos/<int:campeonato_id>/clubes",
    methods=["POST"]
)
@rol_permitido("Administrador")
def guardar_clubes_campeonato(campeonato_id):

    campeonato = db.get_or_404(Campeonato, campeonato_id)

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
        actuales = {
            registro.club_id
            for registro in CampeonatoClub.query.filter_by(
                campeonato_id=campeonato.id
            ).all()
        }
        ids_validos = {
            club.id
            for club in Club.query.filter(Club.id.in_(club_ids)).all()
        } if club_ids else set()

        retirados = actuales - ids_validos

        # Si ya existe fixture, una baja debe hacerse mediante el botón
        # "Dar de baja", que también reorganiza automáticamente el calendario.
        if retirados and Partido.query.filter_by(campeonato_id=campeonato.id).count() > 0:
            flash(
                "Para retirar un club de un campeonato con fixture, usa el botón «Dar de baja». "
                "Así el sistema reorganiza automáticamente los partidos.",
                "error"
            )
            return redirect(url_for("detalle_campeonato", campeonato_id=campeonato.id))

        CampeonatoClub.query.filter_by(
            campeonato_id=campeonato.id
        ).delete(synchronize_session=False)

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
@rol_permitido("Administrador")
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
@rol_permitido("Administrador")
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
@rol_permitido("Administrador")
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
@rol_permitido("Administrador")
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
@rol_permitido("Administrador")
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
@rol_permitido("Administrador")
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
@rol_permitido("Administrador")
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
        and normalizar_estado(jugador.estado) == "Vigente"
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


@app.route("/admin/partidos")
@admin_required
def admin_partidos():
    """Centro interno de partidos: consulta y acceso rápido a fixture y actas."""
    campeonatos = Campeonato.query.order_by(Campeonato.temporada.desc(), Campeonato.id.desc()).all()
    campeonato_id = request.args.get("campeonato_id", type=int)
    estado = (request.args.get("estado") or "").strip()
    q = (request.args.get("q") or "").strip()

    partidos_q = Partido.query.join(Campeonato).order_by(
        Partido.fecha.desc().nullslast(), Partido.jornada.desc(), Partido.id.desc()
    )
    if campeonato_id:
        partidos_q = partidos_q.filter(Partido.campeonato_id == campeonato_id)
    if estado:
        partidos_q = partidos_q.filter(Partido.estado == estado)
    if q:
        patron = f"%{q}%"
        partidos_q = partidos_q.filter(
            db.or_(Club.nombre.ilike(patron),
                   db.exists().where(db.and_(Club.id == Partido.local_club_id, Club.nombre.ilike(patron))),
                   db.exists().where(db.and_(Club.id == Partido.visitante_club_id, Club.nombre.ilike(patron))))
        )
    partidos = partidos_q.limit(250).all()
    resumen = {
        "total": Partido.query.count(),
        "programados": Partido.query.filter(Partido.estado != "Finalizado").count(),
        "finalizados": Partido.query.filter_by(estado="Finalizado").count(),
        "actas_cerradas": ActaPartido.query.filter_by(estado="Cerrada").count(),
    }
    return render_template("admin_partidos.html", partidos=partidos, campeonatos=campeonatos,
                           campeonato_id=campeonato_id, estado=estado, q=q, resumen=resumen)


@app.route("/admin/actas/centro")
@admin_required
def admin_centro_actas():
    """Centro maestro de actas. Consulta unificada sobre las actas existentes."""
    campeonatos = Campeonato.query.order_by(
        Campeonato.temporada.desc(), Campeonato.id.desc()
    ).all()

    estado = (request.args.get("estado") or "").strip()
    campeonato_id = request.args.get("campeonato_id", type=int)
    q = (request.args.get("q") or "").strip()

    query = (
        ActaPartido.query
        .join(Partido)
        .join(Campeonato)
        .order_by(
            Partido.fecha.desc().nullslast(),
            ActaPartido.id.desc()
        )
    )

    if estado:
        query = query.filter(ActaPartido.estado == estado)

    if campeonato_id:
        query = query.filter(Partido.campeonato_id == campeonato_id)

    if q:
        patron = f"%{q}%"
        query = query.filter(ActaPartido.numero_acta.ilike(patron))

    actas = query.limit(250).all()

    total = ActaPartido.query.count()
    borrador = ActaPartido.query.filter_by(estado="Borrador").count()
    cerradas = ActaPartido.query.filter_by(estado="Cerrada").count()

    sin_numero = ActaPartido.query.filter(
        db.or_(ActaPartido.numero_acta.is_(None), ActaPartido.numero_acta == "")
    ).count()

    return render_template(
        "admin_centro_actas.html",
        actas=actas,
        campeonatos=campeonatos,
        campeonato_id=campeonato_id,
        estado=estado,
        q=q,
        resumen={
            "total": total,
            "borrador": borrador,
            "cerradas": cerradas,
            "sin_numero": sin_numero,
        },
    )


@app.route("/admin/actas")
@admin_required
def admin_actas():
    """Centro interno de actas con estado y acceso directo a cada acta."""
    campeonatos = Campeonato.query.order_by(Campeonato.temporada.desc(), Campeonato.id.desc()).all()
    estado = (request.args.get("estado") or "").strip()
    campeonato_id = request.args.get("campeonato_id", type=int)
    q = (request.args.get("q") or "").strip()

    query = (ActaPartido.query.join(Partido).join(Campeonato)
             .order_by(Partido.fecha.desc().nullslast(), ActaPartido.id.desc()))
    if estado:
        query = query.filter(ActaPartido.estado == estado)
    if campeonato_id:
        query = query.filter(Partido.campeonato_id == campeonato_id)
    if q:
        patron = f"%{q}%"
        query = query.filter(ActaPartido.numero_acta.ilike(patron))
    actas = query.limit(250).all()
    resumen = {
        "total": ActaPartido.query.count(),
        "borrador": ActaPartido.query.filter_by(estado="Borrador").count(),
        "cerradas": ActaPartido.query.filter_by(estado="Cerrada").count(),
    }
    return render_template("admin_actas.html", actas=actas, campeonatos=campeonatos,
                           campeonato_id=campeonato_id, estado=estado, q=q, resumen=resumen)


@app.route("/campeonatos/<int:campeonato_id>/partido/<int:partido_id>/acta", methods=["GET", "POST"])
@admin_required
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
                    marcador = f"ACTA_PARTIDO:{partido.id}"
                    Gol.query.filter_by(campeonato_id=campeonato.id, observaciones=marcador).delete(synchronize_session=False)
                    RegistroDisciplinario.query.filter_by(campeonato_id=campeonato.id, observaciones=marcador).delete(synchronize_session=False)
                    acta.estado = "Borrador"
                    partido.goles_local = None
                    partido.goles_visitante = None
                    partido.estado = "Programado"
                    db.session.commit()
                    flash("Acta reabierta. Se retiraron sus estadísticas derivadas hasta volver a cerrarla.", "success")
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

            # V8.1: el resultado oficial se registra dentro del acta.
            # Debe coincidir con la suma de goles ingresados por jugador para
            # mantener una única fuente de verdad y evitar inconsistencias.
            resultado_local_raw = request.form.get("resultado_local", "").strip()
            resultado_visitante_raw = request.form.get("resultado_visitante", "").strip()

            def entero_resultado(valor):
                try:
                    return max(0, int(valor))
                except (ValueError, TypeError):
                    return None

            resultado_local = entero_resultado(resultado_local_raw)
            resultado_visitante = entero_resultado(resultado_visitante_raw)

            if acta.estado == "Cerrada":
                if resultado_local is None or resultado_visitante is None:
                    raise ValueError("Debes ingresar el resultado de ambos equipos antes de cerrar el acta.")

                # El marcador ingresado en el acta es el resultado oficial.
                # Los goles por jugador son datos de goleadores y pueden
                # registrarse aunque no exista detalle individual.
                sincronizar_estadisticas_desde_acta(
                    campeonato, partido, nomina_actual
                )
                partido.goles_local = resultado_local
                partido.goles_visitante = resultado_visitante
                partido.estado = "Finalizado"
                goles_local = resultado_local
                goles_visitante = resultado_visitante
            else:
                # Si se vuelve a guardar como borrador, eliminar cualquier
                # estadística derivada previa de esta acta para mantener una
                # sola fuente de verdad: el acta cerrada.
                marcador = f"ACTA_PARTIDO:{partido.id}"
                Gol.query.filter_by(campeonato_id=campeonato.id, observaciones=marcador).delete(synchronize_session=False)
                RegistroDisciplinario.query.filter_by(campeonato_id=campeonato.id, observaciones=marcador).delete(synchronize_session=False)
                partido.goles_local = None
                partido.goles_visitante = None
                partido.estado = "Programado"
                goles_local = goles_visitante = 0

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
# V9.6 — INFORMES OFICIALES
# ============================================================

@app.route("/campeonatos/<int:campeonato_id>/informes")
@admin_required
def informes_campeonato(campeonato_id):
    campeonato = db.get_or_404(Campeonato, campeonato_id)
    total_partidos = Partido.query.filter_by(campeonato_id=campeonato.id).count()
    finalizados = Partido.query.filter_by(
        campeonato_id=campeonato.id,
        estado="Finalizado"
    ).count()
    return render_template(
        "informes_campeonato.html",
        campeonato=campeonato,
        total_partidos=total_partidos,
        finalizados=finalizados,
    )


@app.route("/campeonatos/<int:campeonato_id>/partido/<int:partido_id>/informe")
@admin_required
def informe_partido(campeonato_id, partido_id):
    campeonato = db.get_or_404(Campeonato, campeonato_id)
    partido = db.get_or_404(Partido, partido_id)
    if partido.campeonato_id != campeonato.id:
        flash("El partido no pertenece a este campeonato.", "error")
        return redirect(url_for("informes_campeonato", campeonato_id=campeonato.id))
    return render_template(
        "informe_partido.html",
        campeonato=campeonato,
        partido=partido,
        acta=partido.acta,
        nomina=sorted(partido.nomina, key=lambda x: (x.equipo, x.jugador.nombre_completo)),
    )


@app.route("/campeonatos/<int:campeonato_id>/informe")
@admin_required
def informe_campeonato(campeonato_id):
    campeonato = db.get_or_404(Campeonato, campeonato_id)
    filas = obtener_tabla_publica(campeonato)
    goleadores = obtener_goleadores_publicos(campeonato, 20)
    partidos = (
        Partido.query.filter_by(campeonato_id=campeonato.id)
        .order_by(Partido.jornada.asc(), Partido.fecha.asc().nullslast(), Partido.id.asc())
        .all()
    )
    return render_template(
        "informe_campeonato.html",
        campeonato=campeonato,
        filas=filas,
        goleadores=goleadores,
        partidos=partidos,
    )


@app.route("/campeonatos/<int:campeonato_id>/club/<int:club_id>/informe")
@admin_required
def informe_club(campeonato_id, club_id):
    campeonato = db.get_or_404(Campeonato, campeonato_id)
    club = db.get_or_404(Club, club_id)
    partidos = (
        Partido.query.filter(
            Partido.campeonato_id == campeonato.id,
            db.or_(Partido.local_club_id == club.id, Partido.visitante_club_id == club.id),
        )
        .order_by(Partido.jornada.asc(), Partido.fecha.asc().nullslast(), Partido.id.asc())
        .all()
    )
    stats = {"pj": 0, "pg": 0, "pe": 0, "pp": 0, "gf": 0, "gc": 0, "pts": 0}
    for p in partidos:
        if p.estado != "Finalizado" or p.goles_local is None or p.goles_visitante is None:
            continue
        stats["pj"] += 1
        local = p.local_club_id == club.id
        gf = p.goles_local if local else p.goles_visitante
        gc = p.goles_visitante if local else p.goles_local
        stats["gf"] += int(gf or 0)
        stats["gc"] += int(gc or 0)
        if gf > gc:
            stats["pg"] += 1; stats["pts"] += 3
        elif gf == gc:
            stats["pe"] += 1; stats["pts"] += 1
        else:
            stats["pp"] += 1
    stats["dg"] = stats["gf"] - stats["gc"]
    plantel = sorted(
        [j for j in Jugador.query.filter_by(club=club.nombre).all()
         if (j.serie or "").strip().lower() == (campeonato.serie or "").strip().lower()],
        key=lambda j: j.nombre_completo.lower()
    )
    return render_template(
        "informe_club.html",
        campeonato=campeonato,
        club=club,
        partidos=partidos,
        stats=stats,
        plantel=plantel,
    )


@app.route("/jugadores/<int:jugador_id>/informe")
@admin_required
def informe_jugador(jugador_id):
    jugador = db.get_or_404(Jugador, jugador_id)
    participaciones = (
        PartidoJugador.query.filter_by(jugador_id=jugador.id)
        .join(Partido)
        .order_by(Partido.fecha.desc().nullslast(), Partido.id.desc())
        .all()
    )
    goles = Gol.query.filter_by(jugador_id=jugador.id).order_by(Gol.fecha.desc(), Gol.id.desc()).all()
    disciplina = (
        RegistroDisciplinario.query.filter_by(jugador_id=jugador.id)
        .order_by(RegistroDisciplinario.fecha.desc(), RegistroDisciplinario.id.desc())
        .all()
    )
    return render_template(
        "informe_jugador.html",
        jugador=jugador,
        participaciones=participaciones,
        goles=goles,
        disciplina=disciplina,
    )

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
             .having(db.func.sum(Gol.cantidad) > 0)
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
        resultados_recientes = (Partido.query
                               .filter(Partido.campeonato_id == campeonato.id,
                                       Partido.estado == "Finalizado",
                                       Partido.goles_local.isnot(None),
                                       Partido.goles_visitante.isnot(None))
                               .order_by(Partido.jornada.desc(), Partido.id.desc())
                               .limit(5).all())
        panel = {"campeonato": campeonato, "jornada": jornada, "partidos": partidos,
                 "resultados": resultados_recientes,
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
    resumen = resumen_publico_campeonato(campeonato)

    # Compatibilidad con versiones anteriores de la plantilla pública que
    # esperaban la variable goleadores_por_club. Mantener ambas variables
    # evita que un despliegue antiguo de la plantilla provoque un 500.
    goleadores_por_club = {}
    for jugador, total in goleadores:
        club_nombre = (getattr(jugador, "club", None) or "Sin club").strip() or "Sin club"
        goleadores_por_club.setdefault(club_nombre, []).append((jugador, total))

    return render_template(
        "publico_campeonato.html",
        campeonato=campeonato,
        jornada=jornada,
        partidos=partidos,
        tabla=tabla,
        goleadores=goleadores,
        goleadores_por_club=goleadores_por_club,
        resumen=resumen,
    )




def normalizar_nombre_publico(valor):
    """Normaliza nombres para agrupar clubes/series aunque existan tildes o diferencias de formato."""
    import unicodedata
    texto = unicodedata.normalize("NFKD", str(valor or ""))
    texto = "".join(c for c in texto if not unicodedata.combining(c))
    texto = texto.lower().strip()
    texto = texto.replace("_", " ").replace("-", " ").replace("/", " ")
    return " ".join(texto.split())


def normalizar_serie_publica(valor):
    texto = normalizar_nombre_publico(valor)
    if texto.startswith("serie "):
        texto = texto[6:].strip()
    return texto


def construir_directorio_clubes():
    """Construye el directorio público Club -> Series -> Jugadores."""
    clubes = Club.query.filter_by(activo=True).order_by(Club.nombre.asc()).all()
    jugadores = (Jugador.query
                 .filter(Jugador.club.isnot(None), Jugador.club != "")
                 .order_by(Jugador.club.asc(), Jugador.serie.asc(), Jugador.nombre_completo.asc())
                 .all())

    clubes_por_nombre = {normalizar_nombre_publico(c.nombre): c for c in clubes}
    grupos = {c.id: {} for c in clubes}

    for jugador in jugadores:
        club = clubes_por_nombre.get(normalizar_nombre_publico(jugador.club))
        if not club:
            continue
        serie = (jugador.serie or "Sin serie").strip() or "Sin serie"
        clave = normalizar_serie_publica(serie)
        grupos[club.id].setdefault(clave, {"nombre": serie, "jugadores": []})["jugadores"].append(jugador)

    # Añade también las series de campeonatos activos, aunque todavía no tengan jugadores.
    participantes = (CampeonatoClub.query
                     .join(Campeonato, CampeonatoClub.campeonato_id == Campeonato.id)
                     .filter(Campeonato.estado.ilike("Activo"))
                     .all())
    for participante in participantes:
        club = participante.club
        if not club or not club.activo:
            continue
        serie = (participante.campeonato.serie or "Sin serie").strip() or "Sin serie"
        clave = normalizar_serie_publica(serie)
        grupos.setdefault(club.id, {}).setdefault(clave, {"nombre": serie, "jugadores": []})

    resultado = []
    for club in clubes:
        series = list(grupos.get(club.id, {}).values())
        series.sort(key=lambda x: normalizar_serie_publica(x["nombre"]))
        total = sum(len(s["jugadores"]) for s in series)
        resultado.append({
            "club": club,
            "series": series,
            "total_jugadores": total,
            "total_series": len(series),
        })
    return resultado


@app.route("/publico/campeones")
def publico_campeones():
    campeonatos = Campeonato.query.order_by(
        Campeonato.temporada.desc(),
        Campeonato.fecha_inicio.desc().nullslast(),
        Campeonato.id.desc()
    ).all()

    historial = []
    for campeonato in campeonatos:
        partidos = Partido.query.filter_by(campeonato_id=campeonato.id).all()
        finalizados = [
            p for p in partidos
            if (p.estado or "").strip().lower() == "finalizado"
            and p.goles_local is not None
            and p.goles_visitante is not None
        ]
        tabla = obtener_tabla_publica(campeonato)
        terminado = (
            (campeonato.estado or "").strip().lower() == "finalizado"
            or (partidos and len(finalizados) == len(partidos))
        )
        campeon = tabla[0]["club"] if terminado and tabla else None
        if campeon:
            historial.append({
                "campeonato": campeonato,
                "club": campeon,
                "temporada": campeonato.temporada,
                "serie": campeonato.serie,
            })

    return render_template("publico_campeones.html", historial=historial)

@app.route("/publico/clubes")
def publico_clubes():
    """Directorio público de clubes, series y planteles registrados."""
    directorio = construir_directorio_clubes()
    return render_template("publico_clubes.html", directorio=directorio)


@app.route("/publico/club/<int:club_id>")
def publico_club_ficha(club_id):
    """Ficha global de un club con todas sus series y jugadores."""
    club = db.get_or_404(Club, club_id)
    if not club.activo:
        from flask import abort
        abort(404)

    directorio = construir_directorio_clubes()
    ficha = next((item for item in directorio if item["club"].id == club.id), None)
    if not ficha:
        ficha = {"club": club, "series": [], "total_jugadores": 0, "total_series": 0}

    campeonatos = (CampeonatoClub.query
                   .join(Campeonato, CampeonatoClub.campeonato_id == Campeonato.id)
                   .filter(CampeonatoClub.club_id == club.id,
                           Campeonato.estado.ilike("Activo"))
                   .order_by(Campeonato.temporada.desc(), Campeonato.id.desc())
                   .all())

    return render_template(
        "publico_club_ficha.html",
        club=club,
        series=ficha["series"],
        total_jugadores=ficha["total_jugadores"],
        total_series=ficha["total_series"],
        campeonatos=campeonatos,
    )


@app.route("/publico/campeonato/<int:campeonato_id>/club/<int:club_id>")
def publico_club(campeonato_id, club_id):
    """Ficha pública de un club dentro de un campeonato."""
    campeonato = db.get_or_404(Campeonato, campeonato_id)
    club = db.get_or_404(Club, club_id)

    participante = CampeonatoClub.query.filter_by(campeonato_id=campeonato.id, club_id=club.id).first()
    if not participante:
        from flask import abort
        abort(404)

    partidos = (Partido.query.filter(
        Partido.campeonato_id == campeonato.id,
        db.or_(Partido.local_club_id == club.id, Partido.visitante_club_id == club.id)
    ).order_by(Partido.jornada, Partido.fecha, Partido.hora, Partido.id).all())

    finalizados = [p for p in partidos if (p.estado or '').lower() == 'finalizado' and p.goles_local is not None and p.goles_visitante is not None]
    proximos = [p for p in partidos if p not in finalizados and (p.estado or '').lower() != 'suspendido']

    pj = pg = pe = pp = gf = gc = pts = 0
    for p in finalizados:
        gl, gv = int(p.goles_local or 0), int(p.goles_visitante or 0)
        if p.local_club_id == club.id:
            gf += gl; gc += gv
            if gl > gv: pg += 1; pts += 3
            elif gl == gv: pe += 1; pts += 1
            else: pp += 1
        else:
            gf += gv; gc += gl
            if gv > gl: pg += 1; pts += 3
            elif gv == gl: pe += 1; pts += 1
            else: pp += 1
        pj += 1

    ultimos = list(reversed(finalizados[-5:]))
    proximos = proximos[:5]

    jugadores = (Jugador.query.filter_by(club=club.nombre, serie=campeonato.serie, estado='Vigente')
                  .order_by(Jugador.nombre_completo.asc()).all())
    goleadores = obtener_goleadores_publicos(campeonato, 500)
    goleadores_club = [(j, t) for j, t in goleadores if (j.club or '').strip().lower() == club.nombre.strip().lower()]
    goleadores_club.sort(key=lambda x: (-x[1], x[0].nombre_completo.lower()))

    amarillas = rojas = 0
    for j in jugadores:
        registros = RegistroDisciplinario.query.filter_by(jugador_id=j.id, campeonato_id=campeonato.id).all()
        amarillas += sum((r.cantidad or 1) for r in registros if (r.tipo or '').lower() in ('amarilla','amarillas','tarjeta amarilla'))
        rojas += sum((r.cantidad or 1) for r in registros if (r.tipo or '').lower() in ('roja','rojas','tarjeta roja'))

    return render_template('publico_club.html', campeonato=campeonato, club=club, partidos=partidos, finalizados=finalizados, ultimos=ultimos, proximos=proximos, jugadores=jugadores, goleadores_club=goleadores_club, stats={'pj':pj,'pg':pg,'pe':pe,'pp':pp,'gf':gf,'gc':gc,'dg':gf-gc,'pts':pts,'amarillas':amarillas,'rojas':rojas})


@app.route("/publico/campeonato/<int:campeonato_id>/jugador/<int:jugador_id>")
def publico_jugador(campeonato_id, jugador_id):
    """Ficha pública de un jugador dentro de un campeonato."""
    campeonato = db.get_or_404(Campeonato, campeonato_id)
    jugador = db.get_or_404(Jugador, jugador_id)

    # El jugador debe pertenecer al club/serie del campeonato para ser visible.
    club = Club.query.filter(db.func.lower(Club.nombre) == (jugador.club or "").strip().lower()).first()
    if not club or not CampeonatoClub.query.filter_by(campeonato_id=campeonato.id, club_id=club.id).first():
        from flask import abort
        abort(404)
    if (jugador.serie or "").strip().lower() != (campeonato.serie or "").strip().lower():
        from flask import abort
        abort(404)

    # V7.4.10: fuente oficial pública = participaciones contenidas en actas cerradas.
    # No mezcla estadísticas históricas/manuales con las estadísticas oficiales.
    participaciones = (PartidoJugador.query
        .join(Partido, PartidoJugador.partido_id == Partido.id)
        .join(ActaPartido, ActaPartido.partido_id == Partido.id)
        .filter(PartidoJugador.jugador_id == jugador.id,
                Partido.campeonato_id == campeonato.id,
                ActaPartido.estado == "Cerrada")
        .order_by(Partido.fecha.desc().nullslast(), Partido.id.desc())
        .all())

    total_partidos = len(participaciones)
    total_titular = sum(1 for x in participaciones if (x.condicion or '').lower() == 'titular')
    total_suplente = sum(1 for x in participaciones if (x.condicion or '').lower() == 'suplente')
    total_goles = sum(int(x.goles or 0) for x in participaciones)
    total_amarillas = sum(int(x.amarillas or 0) for x in participaciones)
    total_rojas = sum(int(x.rojas or 0) for x in participaciones)

    return render_template(
        'publico_jugador.html',
        campeonato=campeonato,
        jugador=jugador,
        club=club,
        participaciones=participaciones,
        total_partidos=total_partidos,
        total_titular=total_titular,
        total_suplente=total_suplente,
        total_goles=total_goles,
        total_amarillas=total_amarillas,
        total_rojas=total_rojas,
    )


@app.route("/publico/campeonato/<int:campeonato_id>/tabla")
def publico_tabla(campeonato_id):
    campeonato = db.get_or_404(Campeonato, campeonato_id)
    return render_template("publico_tabla.html", campeonato=campeonato, tabla=obtener_tabla_publica(campeonato))


@app.route("/publico/campeonato/<int:campeonato_id>/goleadores")
def publico_goleadores(campeonato_id):
    campeonato = db.get_or_404(Campeonato, campeonato_id)
    return render_template("publico_goleadores.html", campeonato=campeonato, goleadores=obtener_goleadores_publicos(campeonato, 100))


def resumen_publico_campeonato(campeonato):
    """Indicadores ligeros para el nuevo portal público V6."""
    partidos = Partido.query.filter_by(campeonato_id=campeonato.id).all()
    finalizados = [p for p in partidos if p.estado == "Finalizado"]
    goles = sum(int(p.goles_local or 0) + int(p.goles_visitante or 0) for p in finalizados)
    equipos = CampeonatoClub.query.filter_by(campeonato_id=campeonato.id).count()
    tabla = obtener_tabla_publica(campeonato)
    lider = tabla[0] if tabla else None
    return {
        "equipos": equipos,
        "partidos_totales": len(partidos),
        "partidos_jugados": len(finalizados),
        "partidos_pendientes": max(0, len(partidos) - len(finalizados)),
        "goles": goles,
        "lider": lider,
    }



def datos_estadisticas_publicas(campeonato):
    """Calcula estadísticas públicas usando la misma fuente de datos del campeonato.
    Se evita depender de funciones SQL específicas para que funcione igual en PostgreSQL y SQLite.
    """
    partidos = Partido.query.filter_by(campeonato_id=campeonato.id).all()
    finalizados = [
        p for p in partidos
        if (p.estado or "").strip().lower() == "finalizado"
        and p.goles_local is not None
        and p.goles_visitante is not None
    ]

    total_goles = sum(int(p.goles_local or 0) + int(p.goles_visitante or 0) for p in finalizados)
    victorias_local = sum(1 for p in finalizados if int(p.goles_local or 0) > int(p.goles_visitante or 0))
    victorias_visitante = sum(1 for p in finalizados if int(p.goles_visitante or 0) > int(p.goles_local or 0))
    empates = sum(1 for p in finalizados if int(p.goles_local or 0) == int(p.goles_visitante or 0))

    tabla = obtener_tabla_publica(campeonato)
    goleadores = obtener_goleadores_publicos(campeonato, 100)
    jugadores = jugadores_campeonato(campeonato)

    total_amarillas = 0
    total_rojas = 0
    total_suspensiones = 0

    for jugador in jugadores:
        registros = RegistroDisciplinario.query.filter_by(jugador_id=jugador.id).all()
        for registro in registros:
            pertenece = (
                registro.campeonato_id == campeonato.id
                or (
                    registro.campeonato_id is None
                    and (registro.campeonato or "").strip().lower() == (campeonato.nombre or "").strip().lower()
                )
            )
            if not pertenece:
                continue

            tipo = (registro.tipo or "").strip().lower()
            cantidad = int(registro.cantidad or 1)

            if "amarilla" in tipo:
                total_amarillas += cantidad
            elif "roja" in tipo:
                total_rojas += cantidad
            elif "suspension" in tipo or "suspensión" in tipo:
                total_suspensiones += cantidad

    promedio_goles = round(total_goles / len(finalizados), 2) if finalizados else 0
    equipos = len(tabla)
    mejor_ataque = max(tabla, key=lambda x: (x.get("gf", 0), x.get("pts", 0))) if tabla else None
    mejor_defensa = min(tabla, key=lambda x: (x.get("gc", 0), -x.get("pts", 0))) if tabla else None
    mayor_diferencia = max(tabla, key=lambda x: (x.get("dg", 0), x.get("pts", 0))) if tabla else None

    return {
        "partidos": len(partidos),
        "finalizados": len(finalizados),
        "pendientes": max(0, len(partidos) - len(finalizados)),
        "goles": total_goles,
        "promedio_goles": promedio_goles,
        "victorias_local": victorias_local,
        "victorias_visitante": victorias_visitante,
        "empates": empates,
        "equipos": equipos,
        "tabla": tabla,
        "goleadores": goleadores,
        "jugadores": len(jugadores),
        "amarillas": total_amarillas,
        "rojas": total_rojas,
        "suspensiones": total_suspensiones,
        "mejor_ataque": mejor_ataque,
        "mejor_defensa": mejor_defensa,
        "mayor_diferencia": mayor_diferencia,
    }

@app.route('/publico/campeonato/<int:campeonato_id>/estadisticas')
def publico_estadisticas(campeonato_id):
    campeonato = db.get_or_404(Campeonato, campeonato_id)
    datos = datos_estadisticas_publicas(campeonato)
    return render_template('publico_estadisticas.html', campeonato=campeonato, datos=datos)


@app.route('/publico/campeonato/<int:campeonato_id>/fair-play')
def publico_fair_play(campeonato_id):
    campeonato = db.get_or_404(Campeonato, campeonato_id)
    datos = datos_estadisticas_publicas(campeonato)

    equipos = []
    jugadores = jugadores_campeonato(campeonato)

    for fila in datos["tabla"]:
        club = fila.get("club") if isinstance(fila, dict) else getattr(fila, "club", None)
        if not club:
            continue
        amarillas = 0
        rojas = 0
        suspensiones = 0

        for jugador in jugadores:
            if (jugador.club or "").strip().lower() != (club.nombre or "").strip().lower():
                continue

            registros = RegistroDisciplinario.query.filter_by(jugador_id=jugador.id).all()
            for registro in registros:
                pertenece = (
                    registro.campeonato_id == campeonato.id
                    or (
                        registro.campeonato_id is None
                        and (registro.campeonato or "").strip().lower() == (campeonato.nombre or "").strip().lower()
                    )
                )
                if not pertenece:
                    continue

                tipo = (registro.tipo or "").strip().lower()
                cantidad = int(registro.cantidad or 1)

                if "amarilla" in tipo:
                    amarillas += cantidad
                elif "roja" in tipo:
                    rojas += cantidad
                elif "suspension" in tipo or "suspensión" in tipo:
                    suspensiones += cantidad

        puntos = amarillas + rojas * 3 + suspensiones * 2
        # Solo publicar clubes con actividad disciplinaria.
        if amarillas == 0 and rojas == 0 and suspensiones == 0:
            continue

        equipos.append({
            "club": club,
            "amarillas": amarillas,
            "rojas": rojas,
            "suspensiones": suspensiones,
            "puntos": puntos,
        })

    equipos.sort(key=lambda x: (x["puntos"], x["amarillas"], x["rojas"], x["club"].nombre.lower()))
    for posicion, equipo in enumerate(equipos, 1):
        equipo["pos"] = posicion

    return render_template(
        "publico_fair_play.html",
        campeonato=campeonato,
        datos=datos,
        equipos=equipos,
    )

@app.route("/publico/campeonato/<int:campeonato_id>/programacion")
def publico_programacion(campeonato_id):
    campeonato = db.get_or_404(Campeonato, campeonato_id)
    partidos = (Partido.query.filter_by(campeonato_id=campeonato.id)
                .order_by(Partido.jornada, Partido.fecha, Partido.hora, Partido.id).all())
    jornadas = {}
    for partido in partidos:
        jornadas.setdefault(partido.jornada, []).append(partido)
    return render_template("publico_programacion.html", campeonato=campeonato, jornadas=jornadas)


@app.route("/publico/campeonato/<int:campeonato_id>/resultados")
def publico_resultados(campeonato_id):
    campeonato = db.get_or_404(Campeonato, campeonato_id)
    partidos = (Partido.query.filter_by(campeonato_id=campeonato.id, estado="Finalizado")
                .order_by(Partido.jornada.desc(), Partido.fecha.desc(), Partido.hora.desc(), Partido.id.desc()).all())
    jornadas = {}
    for partido in partidos:
        jornadas.setdefault(partido.jornada, []).append(partido)
    return render_template("publico_resultados.html", campeonato=campeonato, jornadas=jornadas)


# ============================================================
# ETAPA 8 — GESTIÓN INTERNA DE PLANTELES
# ============================================================

@app.route("/admin/planteles/club/<path:club_nombre>/nuevo", methods=["GET", "POST"])
def admin_nuevo_jugador_club(club_nombre):
    club_nombre = club_nombre.strip()
    clubes, series = obtener_datos_formulario_jugador()
    if request.method == "POST":
        rut = normalizar_rut(request.form.get("rut", ""))
        nombre = request.form.get("nombre_completo", "").strip()
        fecha = request.form.get("fecha_nacimiento", "").strip()
        serie = request.form.get("serie", "").strip()
        estado = normalizar_estado(request.form.get("estado", "Vigente"))

        if not all([rut, nombre, fecha, serie]):
            flash("Completa todos los campos obligatorios.", "error")
            return render_template("admin_jugador_form.html", modo="nuevo", jugador=None,
                                   club_fijo=club_nombre, clubes=clubes, series=series,
                                   estados=sorted(ESTADOS_PERMITIDOS))
        if not validar_rut(rut):
            flash("El RUT ingresado no es válido.", "error")
            return render_template("admin_jugador_form.html", modo="nuevo", jugador=None,
                                   club_fijo=club_nombre, clubes=clubes, series=series,
                                   estados=sorted(ESTADOS_PERMITIDOS))
        try:
            fecha_obj = date.fromisoformat(fecha)
        except ValueError:
            flash("Fecha de nacimiento no válida.", "error")
            return render_template("admin_jugador_form.html", modo="nuevo", jugador=None,
                                   club_fijo=club_nombre, clubes=clubes, series=series,
                                   estados=sorted(ESTADOS_PERMITIDOS))
        if Jugador.query.filter_by(rut=rut).first():
            flash("Ese RUT ya está registrado. Puedes buscarlo en el plantel y editarlo.", "error")
            return render_template("admin_jugador_form.html", modo="nuevo", jugador=None,
                                   club_fijo=club_nombre, clubes=clubes, series=series,
                                   estados=sorted(ESTADOS_PERMITIDOS))
        jugador = Jugador(rut=rut, nombre_completo=nombre, fecha_nacimiento=fecha_obj,
                          serie=serie, club=club_nombre, estado=estado)
        db.session.add(jugador)
        try:
            db.session.flush()
            registrar_movimiento_jugador(jugador, "ALTA", motivo="Inscripción de jugador")
            db.session.commit()
        except Exception as error:
            db.session.rollback()
            print("Error registrando jugador desde plantel:", error)
            flash("No fue posible guardar el jugador.", "error")
            return render_template("admin_jugador_form.html", modo="nuevo", jugador=None,
                                   club_fijo=club_nombre, clubes=clubes, series=series,
                                   estados=sorted(ESTADOS_PERMITIDOS))
        flash(f"{nombre} fue inscrito correctamente en {club_nombre} · {serie}.", "success")
        return redirect(url_for("admin_ficha_club", club_nombre=club_nombre, serie=serie))

    return render_template("admin_jugador_form.html", modo="nuevo", jugador=None,
                           club_fijo=club_nombre, clubes=clubes, series=series,
                           estados=sorted(ESTADOS_PERMITIDOS))


@app.route("/admin/planteles/jugador/<int:jugador_id>/editar", methods=["GET", "POST"])
def admin_editar_jugador_plantel(jugador_id):
    jugador = db.get_or_404(Jugador, jugador_id)
    clubes, series = obtener_datos_formulario_jugador()
    club_anterior = jugador.club
    serie_anterior = jugador.serie
    estado_anterior = jugador.estado

    if request.method == "POST":
        rut = normalizar_rut(request.form.get("rut", ""))
        nombre = request.form.get("nombre_completo", "").strip()
        fecha = request.form.get("fecha_nacimiento", "").strip()
        serie = request.form.get("serie", "").strip()
        club = request.form.get("club", "").strip()
        estado = normalizar_estado(request.form.get("estado", "Vigente"))

        if not all([rut, nombre, fecha, serie, club]):
            flash("Completa todos los campos obligatorios.", "error")
            return render_template("admin_jugador_form.html", modo="editar", jugador=jugador,
                                   clubes=clubes, series=series, club_fijo=None,
                                   estados=sorted(ESTADOS_PERMITIDOS))
        if not validar_rut(rut):
            flash("El RUT ingresado no es válido.", "error")
            return render_template("admin_jugador_form.html", modo="editar", jugador=jugador,
                                   clubes=clubes, series=series, club_fijo=None,
                                   estados=sorted(ESTADOS_PERMITIDOS))
        try:
            fecha_obj = date.fromisoformat(fecha)
        except ValueError:
            flash("Fecha de nacimiento no válida.", "error")
            return render_template("admin_jugador_form.html", modo="editar", jugador=jugador,
                                   clubes=clubes, series=series, club_fijo=None,
                                   estados=sorted(ESTADOS_PERMITIDOS))

        duplicado = Jugador.query.filter(Jugador.rut == rut, Jugador.id != jugador.id).first()
        if duplicado:
            flash("El RUT ya pertenece a otro jugador.", "error")
            return render_template("admin_jugador_form.html", modo="editar", jugador=jugador,
                                   clubes=clubes, series=series, club_fijo=None,
                                   estados=sorted(ESTADOS_PERMITIDOS))

        jugador.rut = rut
        jugador.nombre_completo = nombre
        jugador.fecha_nacimiento = fecha_obj
        jugador.serie = serie
        jugador.club = club
        jugador.estado = estado
        try:
            anterior = {"club": club_anterior, "serie": serie_anterior, "estado": getattr(jugador, "_estado_anterior", None)}
            # El estado anterior se obtiene de la instancia antes de modificarla en versiones nuevas;
            # si no existe, dejamos el campo vacío y registramos igualmente club/serie.
            registrar_movimiento_jugador(
                jugador,
                "ACTUALIZACION",
                anterior={"club": club_anterior, "serie": serie_anterior, "estado": estado_anterior},
                motivo="Edición del registro maestro"
            )
            db.session.commit()
        except Exception as error:
            db.session.rollback()
            print("Error editando jugador desde plantel:", error)
            flash("No fue posible actualizar el jugador.", "error")
            return render_template("admin_jugador_form.html", modo="editar", jugador=jugador,
                                   clubes=clubes, series=series, club_fijo=None,
                                   estados=sorted(ESTADOS_PERMITIDOS))

        if club_anterior != club or serie_anterior != serie:
            flash(f"Jugador actualizado y asignado a {club} · {serie}.", "success")
        else:
            flash("Jugador actualizado correctamente.", "success")
        return redirect(url_for("admin_ficha_club", club_nombre=club, serie=serie))

    return render_template("admin_jugador_form.html", modo="editar", jugador=jugador,
                           clubes=clubes, series=series, club_fijo=None,
                           estados=sorted(ESTADOS_PERMITIDOS))


@app.route("/admin/planteles/jugador/<int:jugador_id>/estado", methods=["POST"])
def admin_cambiar_estado_plantel(jugador_id):
    jugador = db.get_or_404(Jugador, jugador_id)
    estado = normalizar_estado(request.form.get("estado", "Vigente"))
    anterior = {"club": jugador.club, "serie": jugador.serie, "estado": jugador.estado}
    jugador.estado = estado
    registrar_movimiento_jugador(jugador, "CAMBIO_ESTADO", anterior=anterior, motivo="Cambio de estado administrativo")
    db.session.commit()
    flash(f"Estado de {jugador.nombre_completo} actualizado a {estado}.", "success")
    return redirect(url_for("admin_ficha_club", club_nombre=jugador.club, serie=jugador.serie))


@app.route("/admin/planteles/jugador/<int:jugador_id>/mover", methods=["POST"])
def admin_mover_jugador_plantel(jugador_id):
    jugador = db.get_or_404(Jugador, jugador_id)
    club = request.form.get("club", "").strip()
    serie = request.form.get("serie", "").strip()
    if not club or not serie:
        flash("Debes seleccionar club y serie.", "error")
        return redirect(url_for("admin_ficha_club", club_nombre=jugador.club))
    anterior = {"club": jugador.club, "serie": jugador.serie, "estado": jugador.estado}
    anterior_texto = f"{jugador.club} · {jugador.serie}"
    jugador.club = club
    jugador.serie = serie
    registrar_movimiento_jugador(jugador, "CAMBIO_CLUB_SERIE", anterior=anterior, motivo="Movimiento administrativo de plantel")
    db.session.commit()
    flash(f"Jugador movido desde {anterior_texto} a {club} · {serie}.", "success")
    return redirect(url_for("admin_ficha_club", club_nombre=club, serie=serie))


# ============================================================
# ETAPA 9 — CENTRO INTEGRADO DEL JUGADOR
# Conecta ficha, historial, credenciales, estadísticas y actas
# usando el mismo registro Jugador.id.
# ============================================================

@app.route("/admin/jugador/<int:jugador_id>/centro")
@admin_required
def admin_centro_jugador(jugador_id):
    jugador = db.get_or_404(Jugador, jugador_id)

    # V7.4.12: fuente oficial = participaciones de actas cerradas.
    participaciones = (
        PartidoJugador.query
        .join(Partido, PartidoJugador.partido_id == Partido.id)
        .join(Campeonato, Partido.campeonato_id == Campeonato.id)
        .join(ActaPartido, ActaPartido.partido_id == Partido.id)
        .filter(
            PartidoJugador.jugador_id == jugador.id,
            ActaPartido.estado == "Cerrada"
        )
        .order_by(Partido.fecha.desc().nullslast(), Partido.id.desc())
        .all()
    )

    total_partidos = len(participaciones)
    total_titular = sum(1 for p in participaciones if str(p.condicion or "").strip().lower() == "titular")
    total_suplente = sum(1 for p in participaciones if str(p.condicion or "").strip().lower() == "suplente")
    total_goles_acta = sum(int(p.goles or 0) for p in participaciones)
    total_amarillas_acta = sum(int(p.amarillas or 0) for p in participaciones)
    total_rojas_acta = sum(int(p.rojas or 0) for p in participaciones)

    campeonatos = []
    vistos = set()
    for p in participaciones:
        c = p.partido.campeonato if p.partido else None
        if c and c.id not in vistos:
            vistos.add(c.id)
            campeonatos.append(c)

    try:
        movimientos = (
            JugadorMovimiento.query
            .filter_by(jugador_id=jugador.id)
            .order_by(JugadorMovimiento.fecha_hora.desc())
            .limit(8)
            .all()
        )
    except Exception as error:
        db.session.rollback()
        print("Advertencia cargando movimientos del jugador:", repr(error))
        movimientos = []

    return render_template(
        "admin_jugador_centro.html",
        jugador=jugador,
        goles=total_goles_acta,
        amarillas=total_amarillas_acta,
        rojas=total_rojas_acta,
        suspensiones=obtener_suspensiones(jugador.id),
        participaciones=participaciones,
        movimientos=movimientos,
        campeonatos=campeonatos,
        total_partidos=total_partidos,
        total_titular=total_titular,
        total_suplente=total_suplente,
        total_goles_acta=total_goles_acta,
        total_amarillas_acta=total_amarillas_acta,
        total_rojas_acta=total_rojas_acta,
    )

# ============================================================
# ETAPA 10 — CENTRO DE INTEGRACIÓN Y AUDITORÍA
# ============================================================

@app.route("/admin/tesoreria", methods=["GET", "POST"])
@rol_permitido("Administrador", "Tesoreria")
def admin_tesoreria():
    if request.method == "POST":
        accion = (request.form.get("accion") or "movimiento").strip()

        if accion == "cuenta":
            tipo = (request.form.get("tipo_cuenta") or "").strip()
            concepto = (request.form.get("concepto") or "").strip()
            monto_raw = (request.form.get("monto_total") or "").strip().replace(".", "").replace(",", "")
            fecha_raw = (request.form.get("fecha") or "").strip()
            vencimiento_raw = (request.form.get("vencimiento") or "").strip()
            club_id = request.form.get("club_id", type=int)
            campeonato_id = request.form.get("campeonato_id", type=int)
            serie = (request.form.get("serie") or "").strip()
            observaciones = (request.form.get("observaciones") or "").strip()

            if tipo not in {"Por cobrar", "Por pagar"}:
                flash("Selecciona si la cuenta es por cobrar o por pagar.", "error")
                return redirect(url_for("admin_tesoreria"))

            if not concepto:
                flash("Debes ingresar un concepto para la cuenta.", "error")
                return redirect(url_for("admin_tesoreria"))

            try:
                monto_total = int(monto_raw)
            except (TypeError, ValueError):
                monto_total = 0

            if monto_total <= 0:
                flash("El monto total debe ser mayor que $0.", "error")
                return redirect(url_for("admin_tesoreria"))

            try:
                fecha_cuenta = datetime.strptime(fecha_raw, "%Y-%m-%d").date() if fecha_raw else date.today()
            except ValueError:
                fecha_cuenta = date.today()

            try:
                vencimiento = datetime.strptime(vencimiento_raw, "%Y-%m-%d").date() if vencimiento_raw else None
            except ValueError:
                vencimiento = None

            mensaje_cierre = mensaje_periodo_cerrado(fecha_cuenta)
            if mensaje_cierre:
                flash(mensaje_cierre, "error")
                return redirect(url_for("admin_tesoreria"))

            cuenta = CuentaTesoreria(
                fecha=fecha_cuenta,
                tipo=tipo,
                club_id=club_id or None,
                campeonato_id=campeonato_id or None,
                serie=serie or None,
                concepto=concepto,
                monto_total=monto_total,
                monto_pagado=0,
                vencimiento=vencimiento,
                observaciones=observaciones or None,
                creado_por=session.get("admin_nombre") or session.get("admin_username"),
            )
            cuenta.actualizar_estado()
            db.session.add(cuenta)
            db.session.commit()
            flash("Cuenta registrada correctamente.", "success")
            return redirect(url_for("admin_tesoreria"))

        if accion == "pago":
            cuenta_id = request.form.get("cuenta_id", type=int)
            monto_raw = (request.form.get("monto_pago") or "").strip().replace(".", "").replace(",", "")
            fecha_raw = (request.form.get("fecha_pago") or "").strip()
            medio_pago = (request.form.get("medio_pago_pago") or "").strip()
            referencia = (request.form.get("referencia_pago") or "").strip()
            observaciones = (request.form.get("observaciones_pago") or "").strip()

            cuenta = db.session.get(CuentaTesoreria, cuenta_id) if cuenta_id else None
            if not cuenta:
                flash("La cuenta seleccionada no existe.", "error")
                return redirect(url_for("admin_tesoreria"))

            try:
                monto_pago = int(monto_raw)
            except (TypeError, ValueError):
                monto_pago = 0

            saldo_cuenta = cuenta.saldo
            if monto_pago <= 0:
                flash("El monto del abono debe ser mayor que $0.", "error")
                return redirect(url_for("admin_tesoreria"))
            if monto_pago > saldo_cuenta:
                flash("El abono no puede superar el saldo pendiente de la cuenta.", "error")
                return redirect(url_for("admin_tesoreria"))

            try:
                fecha_pago = datetime.strptime(fecha_raw, "%Y-%m-%d").date() if fecha_raw else date.today()
            except ValueError:
                fecha_pago = date.today()

            mensaje_cierre = mensaje_periodo_cerrado(fecha_pago)
            if mensaje_cierre:
                flash(mensaje_cierre, "error")
                return redirect(url_for("admin_tesoreria"))

            cuenta.monto_pagado = int(cuenta.monto_pagado or 0) + monto_pago
            cuenta.actualizar_estado()

            movimiento = MovimientoTesoreria(
                fecha=fecha_pago,
                tipo="Ingreso" if cuenta.tipo == "Por cobrar" else "Egreso",
                concepto=f"Abono: {cuenta.concepto}",
                categoria="Cobros" if cuenta.tipo == "Por cobrar" else "Pagos",
                monto=monto_pago,
                medio_pago=medio_pago or None,
                referencia=referencia or None,
                observaciones=observaciones or None,
                club_id=cuenta.club_id,
                campeonato_id=cuenta.campeonato_id,
                serie=cuenta.serie,
                cuenta_id=cuenta.id,
                creado_por=session.get("admin_nombre") or session.get("admin_username"),
            )
            db.session.add(movimiento)
            db.session.commit()
            flash("Abono registrado correctamente.", "success")
            return redirect(url_for("admin_tesoreria"))

        # Registro rápido de ingreso/egreso, manteniendo la funcionalidad anterior.
        tipo = request.form.get("tipo", "").strip()
        concepto = request.form.get("concepto", "").strip()
        categoria = request.form.get("categoria", "").strip()
        monto_raw = request.form.get("monto", "").strip().replace(".", "").replace(",", "")
        fecha_raw = request.form.get("fecha", "").strip()
        medio_pago = request.form.get("medio_pago", "").strip()
        referencia = request.form.get("referencia", "").strip()
        observaciones = request.form.get("observaciones", "").strip()
        club_id = request.form.get("club_id", type=int)
        campeonato_id = request.form.get("campeonato_id", type=int)
        serie = request.form.get("serie", "").strip()

        if tipo not in {"Ingreso", "Egreso"}:
            flash("Selecciona si el movimiento es un ingreso o un egreso.", "error")
            return redirect(url_for("admin_tesoreria"))

        if not concepto:
            flash("Debes ingresar un concepto.", "error")
            return redirect(url_for("admin_tesoreria"))

        try:
            monto = int(monto_raw)
        except (TypeError, ValueError):
            monto = 0

        if monto <= 0:
            flash("El monto debe ser mayor que $0.", "error")
            return redirect(url_for("admin_tesoreria"))

        try:
            fecha_movimiento = datetime.strptime(fecha_raw, "%Y-%m-%d").date() if fecha_raw else date.today()
        except ValueError:
            fecha_movimiento = date.today()

        mensaje_cierre = mensaje_periodo_cerrado(fecha_movimiento)
        if mensaje_cierre:
            flash(mensaje_cierre, "error")
            return redirect(url_for("admin_tesoreria"))

        movimiento = MovimientoTesoreria(
            fecha=fecha_movimiento,
            tipo=tipo,
            concepto=concepto,
            categoria=categoria or None,
            monto=monto,
            medio_pago=medio_pago or None,
            referencia=referencia or None,
            observaciones=observaciones or None,
            club_id=club_id or None,
            campeonato_id=campeonato_id or None,
            serie=serie or None,
            creado_por=session.get("admin_nombre") or session.get("admin_username"),
        )
        db.session.add(movimiento)
        db.session.commit()
        flash(f"{tipo} registrado correctamente.", "success")
        return redirect(url_for("admin_tesoreria"))

    hoy = date.today()
    try:
        mes = int(request.args.get("mes")) if request.args.get("mes") else hoy.month
    except ValueError:
        mes = hoy.month
    try:
        anio = int(request.args.get("anio")) if request.args.get("anio") else hoy.year
    except ValueError:
        anio = hoy.year

    mes = min(max(mes, 1), 12)

    inicio_mes = date(anio, mes, 1)
    fin_mes = date(anio + 1, 1, 1) if mes == 12 else date(anio, mes + 1, 1)

    movimientos = (
        MovimientoTesoreria.query
        .filter(
            MovimientoTesoreria.fecha >= inicio_mes,
            MovimientoTesoreria.fecha < fin_mes,
        )
        .order_by(MovimientoTesoreria.fecha.desc(), MovimientoTesoreria.id.desc())
        .limit(250)
        .all()
    )

    cuentas = (
        CuentaTesoreria.query
        .order_by(
            db.case(
                (CuentaTesoreria.estado == "Pendiente", 0),
                (CuentaTesoreria.estado == "Abono", 1),
                else_=2,
            ),
            CuentaTesoreria.vencimiento.asc().nullslast(),
            CuentaTesoreria.fecha.desc(),
            CuentaTesoreria.id.desc(),
        )
        .limit(150)
        .all()
    )

    ingresos = db.session.query(
        db.func.coalesce(db.func.sum(MovimientoTesoreria.monto), 0)
    ).filter(
        MovimientoTesoreria.tipo == "Ingreso",
        MovimientoTesoreria.fecha >= inicio_mes,
        MovimientoTesoreria.fecha < fin_mes,
    ).scalar() or 0

    egresos = db.session.query(
        db.func.coalesce(db.func.sum(MovimientoTesoreria.monto), 0)
    ).filter(
        MovimientoTesoreria.tipo == "Egreso",
        MovimientoTesoreria.fecha >= inicio_mes,
        MovimientoTesoreria.fecha < fin_mes,
    ).scalar() or 0

    por_cobrar = sum(x.saldo for x in cuentas if x.tipo == "Por cobrar" and x.saldo > 0)
    por_pagar = sum(x.saldo for x in cuentas if x.tipo == "Por pagar" and x.saldo > 0)
    saldo = int(ingresos) - int(egresos)

    clubes = Club.query.filter_by(activo=True).order_by(Club.nombre.asc()).all()
    campeonatos = Campeonato.query.order_by(Campeonato.temporada.desc(), Campeonato.id.desc()).all()
    series = Serie.query.filter_by(activo=True).order_by(Serie.nombre.asc()).all()

    return render_template(
        "admin_tesoreria.html",
        movimientos=movimientos,
        cuentas=cuentas,
        total_ingresos=int(ingresos),
        total_egresos=int(egresos),
        saldo=int(saldo),
        por_cobrar=int(por_cobrar),
        por_pagar=int(por_pagar),
        today=hoy,
        mes=mes,
        anio=anio,
        clubes=clubes,
        campeonatos=campeonatos,
        series=series,
    )

# ============================================================
# V10.2 — TESORERÍA CONECTADA AL FIXTURE
# Genera cuentas por cobrar/pagar a partir de una jornada real.
# Las cuentas son obligaciones: NO generan movimiento de caja
# hasta que se registre un abono/pago.
# ============================================================

@app.route("/admin/tesoreria/jornada", methods=["GET", "POST"])
@rol_permitido("Administrador", "Tesoreria")
def admin_tesoreria_jornada():
    campeonatos = Campeonato.query.order_by(
        Campeonato.temporada.desc(), Campeonato.id.desc()
    ).all()

    campeonato_id = request.form.get("campeonato_id", type=int) if request.method == "POST" else request.args.get("campeonato_id", type=int)
    fecha_raw = (request.form.get("fecha") if request.method == "POST" else request.args.get("fecha") or "").strip()
    cobro_club_raw = (request.form.get("cobro_club") if request.method == "POST" else request.args.get("cobro_club") or "0").strip()
    cancha_raw = (request.form.get("cancha") if request.method == "POST" else request.args.get("cancha") or "0").strip()
    arbitraje_raw = (request.form.get("arbitraje") if request.method == "POST" else request.args.get("arbitraje") or "0").strip()
    vencimiento_raw = (request.form.get("vencimiento") if request.method == "POST" else request.args.get("vencimiento") or "").strip()

    try:
        fecha = datetime.strptime(fecha_raw, "%Y-%m-%d").date() if fecha_raw else None
    except ValueError:
        fecha = None

    def monto_formulario(valor):
        try:
            return max(0, int(str(valor).replace(".", "").replace(",", "").strip() or 0))
        except (TypeError, ValueError):
            return 0

    cobro_club = monto_formulario(cobro_club_raw)
    monto_cancha = monto_formulario(cancha_raw)
    monto_arbitraje = monto_formulario(arbitraje_raw)

    try:
        vencimiento = datetime.strptime(vencimiento_raw, "%Y-%m-%d").date() if vencimiento_raw else None
    except ValueError:
        vencimiento = None

    campeonato = db.session.get(Campeonato, campeonato_id) if campeonato_id else None
    partidos = []
    if campeonato and fecha:
        partidos = (
            Partido.query
            .filter(
                Partido.campeonato_id == campeonato.id,
                Partido.fecha == fecha,
            )
            .order_by(Partido.hora.asc().nullslast(), Partido.id.asc())
            .all()
        )

    if request.method == "POST":
        accion = (request.form.get("accion") or "").strip()

        if accion == "generar":
            if fecha and tesoreria_periodo_cerrado(fecha):
                flash(mensaje_periodo_cerrado(fecha), "error")
                return redirect(url_for(
                    "admin_tesoreria_jornada",
                    campeonato_id=campeonato.id if campeonato else campeonato_id,
                    fecha=fecha.isoformat() if fecha else "",
                    cobro_club=cobro_club,
                    cancha=monto_cancha,
                    arbitraje=monto_arbitraje,
                    vencimiento=vencimiento.isoformat() if vencimiento else "",
                ))

            if not campeonato:
                flash("Selecciona un campeonato.", "error")
                return redirect(url_for("admin_tesoreria_jornada"))
            if not fecha:
                flash("Selecciona una fecha de jornada válida.", "error")
                return redirect(url_for("admin_tesoreria_jornada", campeonato_id=campeonato.id))
            if not partidos:
                flash("No existen partidos programados para ese campeonato y fecha.", "error")
                return redirect(url_for(
                    "admin_tesoreria_jornada",
                    campeonato_id=campeonato.id,
                    fecha=fecha.isoformat(),
                    cobro_club=cobro_club,
                    cancha=monto_cancha,
                    arbitraje=monto_arbitraje,
                ))
            if cobro_club <= 0 and monto_cancha <= 0 and monto_arbitraje <= 0:
                flash("Debes ingresar al menos un monto mayor que $0.", "error")
                return redirect(url_for(
                    "admin_tesoreria_jornada",
                    campeonato_id=campeonato.id,
                    fecha=fecha.isoformat(),
                ))

            creador = session.get("admin_nombre") or session.get("admin_username")
            creadas = 0
            existentes = 0

            try:
                for partido in partidos:
                    origen = f"PARTIDO:{partido.id}"
                    venc = vencimiento or partido.fecha or date.today()
                    jornada_txt = f"Fecha {partido.jornada}"

                    clubes_partido = [partido.local_club, partido.visitante_club]

                    if cobro_club > 0:
                        for club in clubes_partido:
                            if not club:
                                continue
                            concepto = (
                                f"Cobro jornada {jornada_txt} · "
                                f"{partido.local_club.nombre} vs {partido.visitante_club.nombre}"
                            )
                            existe = (
                                CuentaTesoreria.query
                                .filter(
                                    CuentaTesoreria.partido_id == partido.id,
                                    CuentaTesoreria.tipo == "Por cobrar",
                                    CuentaTesoreria.club_id == club.id,
                                    CuentaTesoreria.origen == origen,
                                )
                                .first()
                            )
                            if existe:
                                existentes += 1
                            else:
                                cuenta = CuentaTesoreria(
                                    fecha=partido.fecha or fecha,
                                    tipo="Por cobrar",
                                    club_id=club.id,
                                    campeonato_id=campeonato.id,
                                    serie=campeonato.serie,
                                    concepto=concepto,
                                    monto_total=cobro_club,
                                    monto_pagado=0,
                                    vencimiento=venc,
                                    observaciones=f"Generado desde fixture · cancha: {partido.cancha or 'Sin cancha'} · hora: {partido.hora or 'Sin hora'}",
                                    creado_por=creador,
                                )
                                cuenta.origen = origen
                                cuenta.actualizar_estado()
                                db.session.add(cuenta)
                                creadas += 1

                    if monto_cancha > 0:
                        concepto_cancha = (
                            f"Cancha {partido.cancha or 'sin recinto'} · "
                            f"{jornada_txt} · {partido.local_club.nombre} vs {partido.visitante_club.nombre}"
                        )
                        existe = (
                            CuentaTesoreria.query
                            .filter(
                                CuentaTesoreria.partido_id == partido.id,
                                CuentaTesoreria.tipo == "Por pagar",
                                CuentaTesoreria.origen == origen,
                                CuentaTesoreria.concepto == concepto_cancha,
                            )
                            .first()
                        )
                        if existe:
                            existentes += 1
                        else:
                            cuenta = CuentaTesoreria(
                                fecha=partido.fecha or fecha,
                                tipo="Por pagar",
                                campeonato_id=campeonato.id,
                                serie=campeonato.serie,
                                concepto=concepto_cancha,
                                monto_total=monto_cancha,
                                monto_pagado=0,
                                vencimiento=venc,
                                observaciones=f"Generado desde fixture · {partido.cancha or 'Sin cancha'}",
                                creado_por=creador,
                            )
                            cuenta.origen = origen
                            cuenta.actualizar_estado()
                            db.session.add(cuenta)
                            creadas += 1

                    if monto_arbitraje > 0:
                        arbitro = ""
                        if partido.acta and partido.acta.arbitro:
                            arbitro = partido.acta.arbitro.strip()
                        nombre_arbitro = arbitro or "Árbitro del partido"
                        concepto_arbitro = (
                            f"Arbitraje · {nombre_arbitro} · {jornada_txt} · "
                            f"{partido.local_club.nombre} vs {partido.visitante_club.nombre}"
                        )
                        existe = (
                            CuentaTesoreria.query
                            .filter(
                                CuentaTesoreria.partido_id == partido.id,
                                CuentaTesoreria.tipo == "Por pagar",
                                CuentaTesoreria.origen == origen,
                                CuentaTesoreria.concepto == concepto_arbitro,
                            )
                            .first()
                        )
                        if existe:
                            existentes += 1
                        else:
                            cuenta = CuentaTesoreria(
                                fecha=partido.fecha or fecha,
                                tipo="Por pagar",
                                campeonato_id=campeonato.id,
                                serie=campeonato.serie,
                                concepto=concepto_arbitro,
                                monto_total=monto_arbitraje,
                                monto_pagado=0,
                                vencimiento=venc,
                                observaciones="Generado desde fixture y acta de partido.",
                                creado_por=creador,
                            )
                            cuenta.origen = origen
                            cuenta.actualizar_estado()
                            db.session.add(cuenta)
                            creadas += 1

                db.session.commit()
                flash(
                    f"Jornada procesada: {creadas} cuentas nuevas y {existentes} ya existentes. "
                    "Las cuentas nuevas quedan pendientes hasta registrar sus abonos/pagos.",
                    "success",
                )
            except Exception as error:
                db.session.rollback()
                print("ERROR GENERANDO TESORERÍA DESDE FIXTURE:", repr(error))
                flash("No fue posible generar las cuentas de la jornada.", "error")

            return redirect(url_for(
                "admin_tesoreria_jornada",
                campeonato_id=campeonato.id,
                fecha=fecha.isoformat(),
                cobro_club=cobro_club,
                cancha=monto_cancha,
                arbitraje=monto_arbitraje,
                vencimiento=vencimiento.isoformat() if vencimiento else "",
            ))

    return render_template(
        "admin_tesoreria_jornada.html",
        campeonatos=campeonatos,
        campeonato=campeonato,
        campeonato_id=campeonato_id,
        fecha=fecha,
        fecha_raw=fecha.isoformat() if fecha else "",
        cobro_club=cobro_club,
        monto_cancha=monto_cancha,
        monto_arbitraje=monto_arbitraje,
        vencimiento=vencimiento,
        vencimiento_raw=vencimiento.isoformat() if vencimiento else "",
        partidos=partidos,
    )

# ============================================================
# ============================================================
# V10.4 — RENDICIÓN MENSUAL Y CIERRE FORMAL
# ============================================================

def calcular_rendicion_mensual(mes, anio):
    inicio_mes, fin_mes = rango_periodo_tesoreria(mes, anio)

    saldo_inicial = db.session.query(
        db.func.coalesce(
            db.func.sum(
                db.case(
                    (MovimientoTesoreria.tipo == "Ingreso", MovimientoTesoreria.monto),
                    else_=-MovimientoTesoreria.monto,
                )
            ), 0
        )
    ).filter(MovimientoTesoreria.fecha < inicio_mes).scalar() or 0

    movimientos = (
        MovimientoTesoreria.query
        .filter(
            MovimientoTesoreria.fecha >= inicio_mes,
            MovimientoTesoreria.fecha < fin_mes,
        )
        .order_by(MovimientoTesoreria.fecha.asc(), MovimientoTesoreria.id.asc())
        .all()
    )

    ingresos = sum(int(m.monto or 0) for m in movimientos if m.tipo == "Ingreso")
    egresos = sum(int(m.monto or 0) for m in movimientos if m.tipo == "Egreso")
    saldo_final = int(saldo_inicial) + ingresos - egresos

    cuentas_periodo = (
        CuentaTesoreria.query
        .filter(
            CuentaTesoreria.fecha >= inicio_mes,
            CuentaTesoreria.fecha < fin_mes,
        )
        .order_by(CuentaTesoreria.fecha.asc(), CuentaTesoreria.id.asc())
        .all()
    )
    cuentas_cobrar_periodo = [c for c in cuentas_periodo if c.tipo == "Por cobrar"]
    cuentas_pagar_periodo = [c for c in cuentas_periodo if c.tipo == "Por pagar"]

    categorias = {}
    for m in movimientos:
        nombre = (m.categoria or "Sin categoría").strip() or "Sin categoría"
        item = categorias.setdefault(nombre, {"ingresos": 0, "egresos": 0, "movimientos": 0})
        item["movimientos"] += 1
        if m.tipo == "Ingreso":
            item["ingresos"] += int(m.monto or 0)
        elif m.tipo == "Egreso":
            item["egresos"] += int(m.monto or 0)

    categorias = [
        {"nombre": nombre, **datos, "neto": datos["ingresos"] - datos["egresos"]}
        for nombre, datos in sorted(
            categorias.items(),
            key=lambda item: max(item[1]["ingresos"], item[1]["egresos"]),
            reverse=True,
        )
    ]

    cuentas_pendientes = (
        CuentaTesoreria.query
        .filter(CuentaTesoreria.estado.in_(["Pendiente", "Abono"]))
        .order_by(
            CuentaTesoreria.tipo.asc(),
            CuentaTesoreria.vencimiento.asc().nullslast(),
            CuentaTesoreria.fecha.asc(),
            CuentaTesoreria.id.asc(),
        )
        .limit(300)
        .all()
    )
    pendientes_cobrar = [c for c in cuentas_pendientes if c.tipo == "Por cobrar" and c.saldo > 0]
    pendientes_pagar = [c for c in cuentas_pendientes if c.tipo == "Por pagar" and c.saldo > 0]

    clubes = Club.query.order_by(Club.nombre.asc()).all()
    resumen_clubes = []
    for club in clubes:
        mov_club = [m for m in movimientos if m.club_id == club.id]
        ing = sum(int(m.monto or 0) for m in mov_club if m.tipo == "Ingreso")
        egr = sum(int(m.monto or 0) for m in mov_club if m.tipo == "Egreso")
        if ing or egr:
            resumen_clubes.append({
                "club": club,
                "ingresos": ing,
                "egresos": egr,
                "neto": ing - egr,
                "movimientos": len(mov_club),
            })

    return {
        "inicio_mes": inicio_mes,
        "fin_mes": fin_mes,
        "fecha_fin_periodo": fin_mes - timedelta(days=1),
        "movimientos": movimientos,
        "categorias": categorias,
        "cuentas_periodo": cuentas_periodo,
        "cuentas_cobrar_periodo": cuentas_cobrar_periodo,
        "cuentas_pagar_periodo": cuentas_pagar_periodo,
        "pendientes_cobrar": pendientes_cobrar,
        "pendientes_pagar": pendientes_pagar,
        "resumen_clubes": resumen_clubes,
        "saldo_inicial": int(saldo_inicial),
        "ingresos": int(ingresos),
        "egresos": int(egresos),
        "saldo_final": int(saldo_final),
        "cuentas_cobrar": int(sum(int(c.monto_total or 0) for c in cuentas_cobrar_periodo)),
        "cuentas_pagar": int(sum(int(c.monto_total or 0) for c in cuentas_pagar_periodo)),
    }


@app.route("/admin/tesoreria/rendicion")
@rol_permitido("Administrador", "Tesoreria")
def admin_tesoreria_rendicion():
    hoy = date.today()
    try:
        mes = int(request.args.get("mes")) if request.args.get("mes") else hoy.month
    except ValueError:
        mes = hoy.month
    try:
        anio = int(request.args.get("anio")) if request.args.get("anio") else hoy.year
    except ValueError:
        anio = hoy.year

    mes = min(max(mes, 1), 12)
    calculo = calcular_rendicion_mensual(mes, anio)
    rendicion = RendicionTesoreria.query.filter_by(
        periodo_mes=mes,
        periodo_anio=anio,
    ).first()

    nombres_meses = [
        "Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio",
        "Julio", "Agosto", "Septiembre", "Octubre", "Noviembre", "Diciembre"
    ]

    return render_template(
        "admin_tesoreria_rendicion.html",
        mes=mes,
        anio=anio,
        nombre_mes=nombres_meses[mes - 1],
        today=hoy,
        rendicion=rendicion,
        **calculo,
        resumen={
            "saldo_inicial": calculo["saldo_inicial"],
            "ingresos": calculo["ingresos"],
            "egresos": calculo["egresos"],
            "saldo_final": calculo["saldo_final"],
            "cuentas_cobrar": calculo["cuentas_cobrar"],
            "cuentas_pagar": calculo["cuentas_pagar"],
            "movimientos": len(calculo["movimientos"]),
        },
    )


@app.route("/admin/tesoreria/rendicion/<int:rendicion_id>/comprobante")
@rol_permitido("Administrador", "Tesoreria")
def admin_tesoreria_comprobante(rendicion_id):
    rendicion = RendicionTesoreria.query.get_or_404(rendicion_id)
    calculo = calcular_rendicion_mensual(rendicion.periodo_mes, rendicion.periodo_anio)

    # El comprobante oficial usa los valores congelados al cierre.
    if rendicion.estado == "Cerrada":
        resumen = {
            "saldo_inicial": int(rendicion.saldo_inicial or 0),
            "ingresos": int(rendicion.ingresos or 0),
            "egresos": int(rendicion.egresos or 0),
            "saldo_final": int(rendicion.saldo_final or 0),
            "movimientos": int(rendicion.movimientos or 0),
            "cuentas_cobrar": int(rendicion.cuentas_por_cobrar or 0),
            "cuentas_pagar": int(rendicion.cuentas_por_pagar or 0),
        }
    else:
        resumen = {
            "saldo_inicial": calculo["saldo_inicial"],
            "ingresos": calculo["ingresos"],
            "egresos": calculo["egresos"],
            "saldo_final": calculo["saldo_final"],
            "movimientos": len(calculo["movimientos"]),
            "cuentas_cobrar": calculo["cuentas_cobrar"],
            "cuentas_pagar": calculo["cuentas_pagar"],
        }

    return render_template(
        "admin_tesoreria_comprobante.html",
        rendicion=rendicion,
        resumen=resumen,
        calculo=calculo,
        nombre_mes=[
            "Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio",
            "Julio", "Agosto", "Septiembre", "Octubre", "Noviembre", "Diciembre"
        ][rendicion.periodo_mes - 1],
        fecha_emision=date.today(),
    )


@app.route("/admin/tesoreria/cierre", methods=["GET", "POST"])
@rol_permitido("Administrador", "Tesoreria")
def admin_tesoreria_cierre():
    hoy = date.today()
    try:
        mes = int(request.form.get("mes") if request.method == "POST" else request.args.get("mes")) if (request.form.get("mes") if request.method == "POST" else request.args.get("mes")) else hoy.month
    except ValueError:
        mes = hoy.month
    try:
        anio = int(request.form.get("anio") if request.method == "POST" else request.args.get("anio")) if (request.form.get("anio") if request.method == "POST" else request.args.get("anio")) else hoy.year
    except ValueError:
        anio = hoy.year

    mes = min(max(mes, 1), 12)
    rendicion = RendicionTesoreria.query.filter_by(
        periodo_mes=mes,
        periodo_anio=anio,
    ).first()

    if request.method == "POST":
        accion = (request.form.get("accion") or "").strip()
        usuario = session.get("admin_nombre") or session.get("admin_username") or "Sistema"
        calculo = calcular_rendicion_mensual(mes, anio)

        if accion == "crear":
            if rendicion:
                flash("La rendición de este período ya existe.", "error")
            else:
                rendicion = RendicionTesoreria(
                    periodo_mes=mes,
                    periodo_anio=anio,
                    numero_rendicion=f"RND-{anio}-{mes:02d}",
                    estado="Borrador",
                    creado_por=usuario,
                )
                db.session.add(rendicion)
                db.session.commit()
                flash(f"Rendición {rendicion.numero_rendicion} creada en estado Borrador.", "success")

        elif accion == "revisar":
            if not rendicion:
                flash("Primero debes crear la rendición.", "error")
            elif rendicion.estado != "Borrador":
                flash("Solo una rendición en Borrador puede pasar a Revisada.", "error")
            else:
                rendicion.estado = "Revisada"
                rendicion.revisado_por = usuario
                rendicion.revisado_en = datetime.utcnow()
                db.session.commit()
                flash("Rendición marcada como Revisada.", "success")

        elif accion == "cerrar":
            if session.get("admin_rol") != "Administrador":
                flash("Solo un Administrador puede cerrar una rendición.", "error")
            elif not rendicion:
                flash("Primero debes crear la rendición.", "error")
            elif rendicion.estado != "Revisada":
                flash("La rendición debe estar Revisada antes de cerrarse.", "error")
            else:
                rendicion.estado = "Cerrada"
                rendicion.saldo_inicial = calculo["saldo_inicial"]
                rendicion.ingresos = calculo["ingresos"]
                rendicion.egresos = calculo["egresos"]
                rendicion.saldo_final = calculo["saldo_final"]
                rendicion.movimientos = len(calculo["movimientos"])
                rendicion.cuentas_por_cobrar = calculo["cuentas_cobrar"]
                rendicion.cuentas_por_pagar = calculo["cuentas_pagar"]
                rendicion.cerrado_por = usuario
                rendicion.cerrado_en = datetime.utcnow()
                db.session.commit()
                flash(f"Rendición {rendicion.numero_rendicion} cerrada. El período quedó bloqueado.", "success")

        elif accion == "reabrir":
            if session.get("admin_rol") != "Administrador":
                flash("Solo un Administrador puede reabrir una rendición.", "error")
            elif not rendicion or rendicion.estado != "Cerrada":
                flash("Solo una rendición Cerrada puede reabrirse.", "error")
            else:
                rendicion.estado = "Revisada"
                rendicion.observaciones = (
                    (rendicion.observaciones or "") +
                    f" | Reabierta por {usuario} el {datetime.utcnow().strftime('%d/%m/%Y %H:%M')}"
                ).strip(" |")
                db.session.commit()
                flash("Rendición reabierta en estado Revisada. El período vuelve a estar editable.", "success")

        return redirect(url_for(
            "admin_tesoreria_cierre",
            mes=mes,
            anio=anio,
        ))

    calculo = calcular_rendicion_mensual(mes, anio)
    return render_template(
        "admin_tesoreria_cierre.html",
        mes=mes,
        anio=anio,
        nombre_mes=[
            "Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio",
            "Julio", "Agosto", "Septiembre", "Octubre", "Noviembre", "Diciembre"
        ][mes - 1],
        today=hoy,
        rendicion=rendicion,
        calculo=calculo,
    )

# ============================================================
# V10.1 — ESTADO FINANCIERO POR CLUB
# Vista individual del historial financiero de cada club.
# Utiliza exclusivamente cuentas y movimientos de Tesorería.
# ============================================================

@app.route("/admin/tesoreria/club/<int:club_id>")
@rol_permitido("Administrador", "Tesoreria")
def admin_tesoreria_club(club_id):
    club = db.get_or_404(Club, club_id)

    hoy = date.today()
    try:
        mes = int(request.args.get("mes")) if request.args.get("mes") else hoy.month
    except ValueError:
        mes = hoy.month
    try:
        anio = int(request.args.get("anio")) if request.args.get("anio") else hoy.year
    except ValueError:
        anio = hoy.year

    mes = min(max(mes, 1), 12)
    inicio_mes = date(anio, mes, 1)
    fin_mes = date(anio + 1, 1, 1) if mes == 12 else date(anio, mes + 1, 1)

    movimientos = (
        MovimientoTesoreria.query
        .filter(MovimientoTesoreria.club_id == club.id)
        .order_by(MovimientoTesoreria.fecha.desc(), MovimientoTesoreria.id.desc())
        .limit(300)
        .all()
    )

    movimientos_periodo = [
        m for m in movimientos
        if m.fecha and inicio_mes <= m.fecha < fin_mes
    ]

    cuentas = (
        CuentaTesoreria.query
        .filter(CuentaTesoreria.club_id == club.id)
        .order_by(
            db.case(
                (CuentaTesoreria.estado == "Pendiente", 0),
                (CuentaTesoreria.estado == "Abono", 1),
                else_=2,
            ),
            CuentaTesoreria.vencimiento.asc().nullslast(),
            CuentaTesoreria.fecha.desc(),
            CuentaTesoreria.id.desc(),
        )
        .all()
    )

    cuentas_cobrar = [c for c in cuentas if c.tipo == "Por cobrar"]
    cuentas_pagar = [c for c in cuentas if c.tipo == "Por pagar"]

    total_por_cobrar = sum(c.saldo for c in cuentas_cobrar if c.saldo > 0)
    total_por_pagar = sum(c.saldo for c in cuentas_pagar if c.saldo > 0)
    total_cobrado = sum(int(c.monto_pagado or 0) for c in cuentas_cobrar)
    total_pagado = sum(int(c.monto_pagado or 0) for c in cuentas_pagar)

    ingresos_periodo = sum(
        int(m.monto or 0) for m in movimientos_periodo if m.tipo == "Ingreso"
    )
    egresos_periodo = sum(
        int(m.monto or 0) for m in movimientos_periodo if m.tipo == "Egreso"
    )

    ingresos_historicos = sum(
        int(m.monto or 0) for m in movimientos if m.tipo == "Ingreso"
    )
    egresos_historicos = sum(
        int(m.monto or 0) for m in movimientos if m.tipo == "Egreso"
    )

    campeonatos = (
        Campeonato.query
        .join(CuentaTesoreria, CuentaTesoreria.campeonato_id == Campeonato.id)
        .filter(CuentaTesoreria.club_id == club.id)
        .distinct()
        .order_by(Campeonato.temporada.desc(), Campeonato.id.desc())
        .all()
    )

    categorias = {}
    for m in movimientos:
        categoria = (m.categoria or "Sin categoría").strip() or "Sin categoría"
        categorias.setdefault(categoria, {"ingresos": 0, "egresos": 0})
        if m.tipo == "Ingreso":
            categorias[categoria]["ingresos"] += int(m.monto or 0)
        elif m.tipo == "Egreso":
            categorias[categoria]["egresos"] += int(m.monto or 0)

    categorias = [
        {"nombre": nombre, **datos}
        for nombre, datos in sorted(
            categorias.items(),
            key=lambda item: item[0].lower()
        )
    ]

    return render_template(
        "admin_tesoreria_club.html",
        club=club,
        movimientos=movimientos,
        cuentas=cuentas,
        cuentas_cobrar=cuentas_cobrar,
        cuentas_pagar=cuentas_pagar,
        campeonatos=campeonatos,
        categorias=categorias,
        mes=mes,
        anio=anio,
        inicio_mes=inicio_mes,
        fin_mes=fin_mes,
        today=hoy,
        resumen={
            "por_cobrar": int(total_por_cobrar),
            "por_pagar": int(total_por_pagar),
            "cobrado": int(total_cobrado),
            "pagado": int(total_pagado),
            "ingresos_periodo": int(ingresos_periodo),
            "egresos_periodo": int(egresos_periodo),
            "saldo_periodo": int(ingresos_periodo - egresos_periodo),
            "ingresos_historicos": int(ingresos_historicos),
            "egresos_historicos": int(egresos_historicos),
            "saldo_historico": int(ingresos_historicos - egresos_historicos),
        },
    )


@app.route("/admin/integracion")
def admin_integracion():
    """Panel técnico para verificar que los módulos usen el mismo registro maestro."""
    jugadores = Jugador.query.count()
    clubes = Club.query.count()
    series = Serie.query.count()
    campeonatos = Campeonato.query.count()
    partidos = Partido.query.count()
    actas = ActaPartido.query.count()
    participaciones = PartidoJugador.query.count()
    goles = Gol.query.count()
    disciplina = RegistroDisciplinario.query.count()

    actas_cerradas = ActaPartido.query.filter_by(estado="Cerrada").count()
    participaciones_no_vigentes = (
        PartidoJugador.query.join(Jugador, PartidoJugador.jugador_id == Jugador.id)
        .filter(Jugador.estado != "Vigente").count()
    )

    inconsistencias_club = 0
    for p in PartidoJugador.query.join(Partido).join(Jugador).all():
        partido = p.partido
        jugador = p.jugador
        club_id = partido.local_club_id if p.equipo == "local" else partido.visitante_club_id
        club = db.session.get(Club, club_id)
        if club and _normalizar_texto_acta(jugador.club) != _normalizar_texto_acta(club.nombre):
            inconsistencias_club += 1

    resumen = {
        "jugadores": jugadores, "clubes": clubes[:10], "series": series,
        "campeonatos": campeonatos, "partidos": partidos, "actas": actas,
        "actas_cerradas": actas_cerradas, "participaciones": participaciones,
        "goles": goles, "disciplina": disciplina,
        "participaciones_no_vigentes": participaciones_no_vigentes,
        "inconsistencias_club": inconsistencias_club,
    }
    return render_template("admin_integracion.html", resumen=resumen)


# ============================================================
# V6.5 — CENTRO DE INTEGRACIÓN TOTAL
# ============================================================

@app.route("/admin/centro-integracion")
@admin_required
def admin_centro_integracion():
    """Centro transversal del sistema.

    No crea registros ni modifica la base de datos. Reúne las relaciones
    existentes para que el administrador pueda navegar desde campeonato
    hasta club, jugador, partido, acta y estadísticas.
    """
    campeonatos = (
        Campeonato.query
        .order_by(Campeonato.temporada.desc(), Campeonato.id.desc())
        .all()
    )

    partidos = (
        Partido.query
        .join(Campeonato)
        .order_by(Partido.fecha.desc().nullslast(), Partido.id.desc())
        .limit(12)
        .all()
    )

    actas_abiertas = (
        ActaPartido.query
        .filter(ActaPartido.estado != "Cerrada")
        .join(Partido)
        .order_by(Partido.fecha.desc().nullslast(), ActaPartido.id.desc())
        .limit(12)
        .all()
    )

    # Campeonatos con sus métricas principales.
    resumen_campeonatos = []
    for campeonato in campeonatos[:10]:
        total_partidos = Partido.query.filter_by(campeonato_id=campeonato.id).count()
        finalizados = Partido.query.filter_by(
            campeonato_id=campeonato.id,
            estado="Finalizado"
        ).count()
        actas = (
            ActaPartido.query
            .join(Partido)
            .filter(Partido.campeonato_id == campeonato.id)
            .count()
        )
        actas_cerradas = (
            ActaPartido.query
            .join(Partido)
            .filter(
                Partido.campeonato_id == campeonato.id,
                ActaPartido.estado == "Cerrada"
            )
            .count()
        )
        clubes = CampeonatoClub.query.filter_by(campeonato_id=campeonato.id).count()
        resumen_campeonatos.append({
            "campeonato": campeonato,
            "partidos": total_partidos,
            "finalizados": finalizados,
            "actas": actas,
            "actas_cerradas": actas_cerradas,
            "clubes": clubes,
        })

    # Controles de integridad de lectura: no alteran ningún dato.
    partidos_sin_acta = (
        Partido.query
        .outerjoin(ActaPartido, ActaPartido.partido_id == Partido.id)
        .filter(ActaPartido.id.is_(None))
        .count()
    )

    actas_cerradas_sin_nomina = (
        ActaPartido.query
        .join(Partido)
        .filter(ActaPartido.estado == "Cerrada")
        .filter(~db.exists().where(PartidoJugador.partido_id == Partido.id))
        .count()
    )

    participaciones_no_vigentes = (
        PartidoJugador.query
        .join(Jugador, PartidoJugador.jugador_id == Jugador.id)
        .filter(Jugador.estado != "Vigente")
        .count()
    )

    inconsistencias_club = 0
    for participacion in (
        PartidoJugador.query
        .join(Partido, PartidoJugador.partido_id == Partido.id)
        .join(Jugador, PartidoJugador.jugador_id == Jugador.id)
        .all()
    ):
        partido = participacion.partido
        jugador = participacion.jugador
        club_id = (
            partido.local_club_id
            if participacion.equipo == "local"
            else partido.visitante_club_id
        )
        club = db.session.get(Club, club_id)
        if club and _normalizar_texto_acta(jugador.club) != _normalizar_texto_acta(club.nombre):
            inconsistencias_club += 1

    alertas = [
        ("Partidos sin acta", partidos_sin_acta, "Revisar desde Partidos"),
        ("Actas cerradas sin nómina", actas_cerradas_sin_nomina, "Revisar acta y nómina"),
        ("Participaciones de jugadores no vigentes", participaciones_no_vigentes, "Revisar Registro"),
        ("Inconsistencias jugador/club", inconsistencias_club, "Revisar plantel"),
    ]

    resumen = {
        "jugadores": Jugador.query.count(),
        "jugadores_vigentes": Jugador.query.filter_by(estado="Vigente").count(),
        "clubes": Club.query.count(),
        "clubes_activos": Club.query.filter_by(activo=True).count(),
        "series": Serie.query.count(),
        "campeonatos": Campeonato.query.count(),
        "partidos": Partido.query.count(),
        "partidos_finalizados": Partido.query.filter_by(estado="Finalizado").count(),
        "actas": ActaPartido.query.count(),
        "actas_cerradas": ActaPartido.query.filter_by(estado="Cerrada").count(),
        "participaciones": PartidoJugador.query.count(),
        "goles": Gol.query.count(),
        "disciplina": RegistroDisciplinario.query.count(),
    }

    return render_template(
        "admin_centro_integracion.html",
        resumen=resumen,
        campeonatos=resumen_campeonatos,
        partidos=partidos,
        actas_abiertas=actas_abiertas,
        alertas=alertas,
    )


# ============================================================
# V6.6 — PANEL MAESTRO DE LA ASOCIACIÓN
# ============================================================

@app.route("/admin/operaciones")
@admin_required
def admin_operaciones():
    """Centro de operaciones diario: agenda, pendientes y accesos de trabajo.

    Es una vista operativa sobre los mismos registros existentes; no crea
    tablas ni duplica jugadores, partidos o actas.
    """
    hoy = date.today()
    limite = hoy + timedelta(days=7)

    partidos_hoy = (Partido.query
        .filter(Partido.fecha == hoy)
        .order_by(Partido.hora.asc().nullslast(), Partido.id.asc())
        .all())

    proximos = (Partido.query
        .filter(Partido.fecha > hoy, Partido.fecha <= limite)
        .order_by(Partido.fecha.asc(), Partido.hora.asc().nullslast(), Partido.id.asc())
        .all())

    pendientes_acta = (Partido.query
        .outerjoin(ActaPartido, ActaPartido.partido_id == Partido.id)
        .filter(Partido.fecha <= hoy, Partido.estado != "Finalizado", ActaPartido.id.is_(None))
        .order_by(Partido.fecha.asc().nullslast(), Partido.hora.asc().nullslast(), Partido.id.asc())
        .limit(20).all())

    actas_borrador = (ActaPartido.query
        .join(Partido)
        .filter(ActaPartido.estado != "Cerrada")
        .order_by(Partido.fecha.desc().nullslast(), ActaPartido.id.desc())
        .limit(20).all())

    resultados_pendientes = (Partido.query
        .filter(
            Partido.estado == "Finalizado",
            db.or_(Partido.goles_local.is_(None), Partido.goles_visitante.is_(None))
        )
        .order_by(Partido.fecha.desc().nullslast(), Partido.id.desc())
        .limit(20).all())

    def acta_estado(partido):
        return partido.acta.estado if partido.acta else "Sin acta"

    return render_template(
        "admin_operaciones.html",
        hoy=hoy,
        limite=limite,
        partidos_hoy=partidos_hoy,
        proximos=proximos,
        pendientes_acta=pendientes_acta,
        actas_borrador=actas_borrador,
        resultados_pendientes=resultados_pendientes,
        acta_estado=acta_estado,
    )


@app.route("/admin/panel-maestro")
@admin_required
def admin_panel_maestro():
    """Panel ejecutivo de lectura del sistema.

    Reúne indicadores de las entidades existentes sin crear registros
    paralelos ni modificar la estructura de PostgreSQL.
    """

    def safe_count(query):
        try:
            return query.count()
        except Exception:
            db.session.rollback()
            return 0

    def safe_scalar(query, default=0):
        try:
            value = query.scalar()
            return default if value is None else value
        except Exception:
            db.session.rollback()
            return default

    # -------------------------
    # INDICADORES PRINCIPALES
    # -------------------------
    total_jugadores = safe_count(Jugador.query)
    jugadores_vigentes = safe_count(Jugador.query.filter_by(estado="Vigente"))
    total_clubes = safe_count(Club.query)
    clubes_activos = safe_count(Club.query.filter_by(activo=True))
    total_campeonatos = safe_count(Campeonato.query)
    campeonatos_activos = safe_count(Campeonato.query.filter_by(estado="Activo"))
    total_partidos = safe_count(Partido.query)
    partidos_finalizados = safe_count(Partido.query.filter_by(estado="Finalizado"))
    partidos_programados = safe_count(Partido.query.filter_by(estado="Programado"))
    total_actas = safe_count(ActaPartido.query)
    actas_cerradas = safe_count(ActaPartido.query.filter_by(estado="Cerrada"))
    actas_pendientes = safe_count(ActaPartido.query.filter(ActaPartido.estado != "Cerrada"))

    total_goles = safe_scalar(
        db.session.query(db.func.coalesce(db.func.sum(Gol.cantidad), 0))
    )
    total_amarillas = safe_scalar(
        db.session.query(db.func.coalesce(db.func.sum(RegistroDisciplinario.cantidad), 0))
        .filter(RegistroDisciplinario.tipo == "Amarilla")
    )
    total_rojas = safe_scalar(
        db.session.query(db.func.coalesce(db.func.sum(RegistroDisciplinario.cantidad), 0))
        .filter(RegistroDisciplinario.tipo == "Roja")
    )

    # -------------------------
    # PENDIENTES / CONTROL
    # -------------------------
    partidos_sin_acta = safe_count(
        Partido.query
        .outerjoin(ActaPartido, ActaPartido.partido_id == Partido.id)
        .filter(ActaPartido.id.is_(None))
    )

    actas_cerradas_sin_nomina = safe_count(
        ActaPartido.query
        .join(Partido)
        .filter(ActaPartido.estado == "Cerrada")
        .filter(~db.exists().where(PartidoJugador.partido_id == Partido.id))
    )

    partidos_sin_resultado = safe_count(
        Partido.query.filter(
            Partido.estado == "Finalizado",
            db.or_(Partido.goles_local.is_(None), Partido.goles_visitante.is_(None))
        )
    )

    jugadores_sin_club = safe_count(
        Jugador.query.filter(
            db.or_(Jugador.club.is_(None), Jugador.club == "")
        )
    )

    jugadores_sin_serie = safe_count(
        Jugador.query.filter(
            db.or_(Jugador.serie.is_(None), Jugador.serie == "")
        )
    )

    alertas = [
        {"tipo": "danger", "icono": "📋", "titulo": "Partidos sin acta", "cantidad": partidos_sin_acta,
         "texto": "Partidos que todavía no tienen acta asociada.", "url": url_for("admin_partidos")},
        {"tipo": "warning", "icono": "📝", "titulo": "Actas pendientes", "cantidad": actas_pendientes,
         "texto": "Actas que aún no están cerradas.", "url": url_for("admin_actas")},
        {"tipo": "warning", "icono": "👥", "titulo": "Actas cerradas sin nómina", "cantidad": actas_cerradas_sin_nomina,
         "texto": "Actas cerradas que no tienen jugadores asociados.", "url": url_for("admin_actas")},
        {"tipo": "warning", "icono": "⚽", "titulo": "Finalizados sin resultado", "cantidad": partidos_sin_resultado,
         "texto": "Partidos marcados como finalizados sin marcador completo.", "url": url_for("admin_partidos")},
        {"tipo": "info", "icono": "👤", "titulo": "Jugadores sin club", "cantidad": jugadores_sin_club,
         "texto": "Registros maestros que todavía no tienen club informado.", "url": url_for("admin_registro")},
        {"tipo": "info", "icono": "🏷️", "titulo": "Jugadores sin serie", "cantidad": jugadores_sin_serie,
         "texto": "Registros maestros que todavía no tienen serie informada.", "url": url_for("admin_registro")},
    ]

    # Solo mostrar pendientes reales en la tarjeta de control.
    alertas = [a for a in alertas if a["cantidad"] > 0]

    # -------------------------
    # CAMPEONATOS
    # -------------------------
    campeonatos = Campeonato.query.order_by(
        Campeonato.temporada.desc(), Campeonato.id.desc()
    ).limit(10).all()

    resumen_campeonatos = []
    for campeonato in campeonatos:
        partidos = Partido.query.filter_by(campeonato_id=campeonato.id)
        resumen_campeonatos.append({
            "campeonato": campeonato,
            "clubes": safe_count(CampeonatoClub.query.filter_by(campeonato_id=campeonato.id)),
            "partidos": safe_count(partidos),
            "finalizados": safe_count(partidos.filter_by(estado="Finalizado")),
            "actas": safe_count(
                ActaPartido.query.join(Partido).filter(Partido.campeonato_id == campeonato.id)
            ),
            "actas_cerradas": safe_count(
                ActaPartido.query.join(Partido).filter(
                    Partido.campeonato_id == campeonato.id,
                    ActaPartido.estado == "Cerrada"
                )
            ),
        })

    # -------------------------
    # PRÓXIMOS PARTIDOS
    # -------------------------
    hoy = date.today()
    proximos_partidos = (
        Partido.query
        .filter(
            Partido.estado == "Programado",
            db.or_(Partido.fecha.is_(None), Partido.fecha >= hoy)
        )
        .order_by(Partido.fecha.asc().nullslast(), Partido.hora.asc().nullslast(), Partido.id.asc())
        .limit(8)
        .all()
    )

    # -------------------------
    # ÚLTIMOS PARTIDOS FINALIZADOS
    # -------------------------
    ultimos_resultados = (
        Partido.query
        .filter(Partido.estado == "Finalizado")
        .order_by(Partido.fecha.desc().nullslast(), Partido.id.desc())
        .limit(8)
        .all()
    )

    # -------------------------
    # ÚLTIMOS JUGADORES
    # -------------------------
    ultimos_jugadores = Jugador.query.order_by(Jugador.id.desc()).limit(6).all()

    resumen = {
        "total_jugadores": total_jugadores,
        "jugadores_vigentes": jugadores_vigentes,
        "total_clubes": total_clubes,
        "clubes_activos": clubes_activos,
        "total_campeonatos": total_campeonatos,
        "campeonatos_activos": campeonatos_activos,
        "total_partidos": total_partidos,
        "partidos_finalizados": partidos_finalizados,
        "partidos_programados": partidos_programados,
        "total_actas": total_actas,
        "actas_cerradas": actas_cerradas,
        "actas_pendientes": actas_pendientes,
        "total_goles": total_goles,
        "total_amarillas": total_amarillas,
        "total_rojas": total_rojas,
    }

    return render_template(
        "admin_panel_maestro.html",
        resumen=resumen,
        alertas=alertas,
        campeonatos=resumen_campeonatos,
        proximos_partidos=proximos_partidos,
        ultimos_resultados=ultimos_resultados,
        ultimos_jugadores=ultimos_jugadores,
    )



# ============================================================
# V7.1 — CENTRO MAESTRO DE CAMPEONATOS
# Consulta unificada de competencias existentes.
# No crea tablas ni modifica registros.
# ============================================================

@app.route("/campeonatos/<int:campeonato_id>/eliminar", methods=["POST"])
@rol_permitido("Administrador")
def eliminar_campeonato(campeonato_id):
    """Elimina un campeonato y sus datos dependientes, sin tocar clubes ni jugadores."""
    campeonato = db.get_or_404(Campeonato, campeonato_id)

    try:
        partidos = Partido.query.filter_by(campeonato_id=campeonato.id).all()
        partido_ids = [p.id for p in partidos]

        # Eliminar primero los registros que dependen de los partidos.
        if partido_ids:
            ActaPartido.query.filter(
                ActaPartido.partido_id.in_(partido_ids)
            ).delete(synchronize_session=False)

            PartidoJugador.query.filter(
                PartidoJugador.partido_id.in_(partido_ids)
            ).delete(synchronize_session=False)

        # Eliminar estadísticas y disciplina asociadas al campeonato.
        Gol.query.filter_by(
            campeonato_id=campeonato.id
        ).delete(synchronize_session=False)

        RegistroDisciplinario.query.filter_by(
            campeonato_id=campeonato.id
        ).delete(synchronize_session=False)

        # Eliminar fixture y clubes inscritos.
        Partido.query.filter_by(
            campeonato_id=campeonato.id
        ).delete(synchronize_session=False)

        CampeonatoClub.query.filter_by(
            campeonato_id=campeonato.id
        ).delete(synchronize_session=False)

        nombre = campeonato.nombre
        db.session.delete(campeonato)
        db.session.commit()

        flash(
            f"El campeonato «{nombre}» fue eliminado correctamente. "
            "Los clubes y jugadores se conservaron.",
            "success"
        )
    except Exception:
        db.session.rollback()
        app.logger.exception(
            "Error al eliminar campeonato %s",
            campeonato_id
        )
        flash(
            "No fue posible eliminar el campeonato. No se realizaron cambios.",
            "error"
        )

    return redirect(url_for("campeonatos"))


@app.route("/admin/campeonatos/centro")
def admin_centro_campeonatos():
    campeonatos = Campeonato.query.order_by(
        Campeonato.temporada.desc(),
        Campeonato.id.desc()
    ).all()

    tarjetas = []
    for campeonato in campeonatos:
        club_count = CampeonatoClub.query.filter_by(
            campeonato_id=campeonato.id
        ).count()

        partidos_query = Partido.query.filter_by(
            campeonato_id=campeonato.id
        )
        partidos_count = partidos_query.count()
        finalizados = partidos_query.filter_by(
            estado="Finalizado"
        ).count()

        actas_count = (
            ActaPartido.query
            .join(Partido, ActaPartido.partido_id == Partido.id)
            .filter(Partido.campeonato_id == campeonato.id)
            .count()
        )

        actas_cerradas = (
            ActaPartido.query
            .join(Partido, ActaPartido.partido_id == Partido.id)
            .filter(
                Partido.campeonato_id == campeonato.id,
                ActaPartido.estado == "Cerrada"
            )
            .count()
        )

        tarjetas.append({
            "campeonato": campeonato,
            "clubes": club_count,
            "partidos": partidos_count,
            "finalizados": finalizados,
            "pendientes": partidos_count - finalizados,
            "actas": actas_count,
            "actas_cerradas": actas_cerradas,
            "actas_pendientes": actas_count - actas_cerradas,
        })

    return render_template(
        "admin_centro_campeonatos.html",
        tarjetas=tarjetas,
        total_campeonatos=len(campeonatos),
        activos=sum(1 for x in campeonatos if x.estado == "Activo"),
        finalizados=sum(1 for x in campeonatos if x.estado == "Finalizado"),
        total_partidos=sum(x["partidos"] for x in tarjetas),
        total_pendientes=sum(x["pendientes"] for x in tarjetas),
    )



# ============================================================
# V7.4.8 — CENTRO MAESTRO DE SERIES
# ============================================================

@app.route("/admin/series/centro")
@admin_required
def admin_centro_series():
    series_db = Serie.query.order_by(Serie.nombre.asc()).all()
    campeonatos = Campeonato.query.order_by(
        Campeonato.temporada.desc(), Campeonato.id.desc()
    ).all()

    filas = []
    for serie_obj in series_db:
        nombre = (serie_obj.nombre or "").strip()
        campeonatos_serie = [
            c for c in campeonatos
            if (c.serie or "").strip().lower() == nombre.lower()
        ]

        club_ids = set()
        for campeonato in campeonatos_serie:
            participaciones = CampeonatoClub.query.filter_by(
                campeonato_id=campeonato.id
            ).all()
            club_ids.update(x.club_id for x in participaciones)

        clubes = Club.query.filter(Club.id.in_(club_ids)).order_by(
            Club.nombre.asc()
        ).all() if club_ids else []

        jugadores = Jugador.query.filter(
            Jugador.serie == nombre
        ).order_by(
            Jugador.club.asc(), Jugador.nombre_completo.asc()
        ).all()

        filas.append({
            "serie": serie_obj,
            "campeonatos": campeonatos_serie[:8],
            "clubes": clubes,
            "jugadores": jugadores,
            "total_clubes": len(clubes),
            "total_jugadores": len(jugadores),
            "vigentes": sum(
                1 for j in jugadores
                if (j.estado or "Vigente") == "Vigente"
            ),
        })

    return render_template(
        "admin_centro_series.html",
        filas=filas,
        total_series=len(filas),
        total_clubes=len({
            c.id for f in filas for c in f["clubes"]
        }),
        total_jugadores=sum(f["total_jugadores"] for f in filas),
    )


@app.route("/admin/series/<int:serie_id>/centro")
@admin_required
def admin_centro_serie(serie_id):
    """V7.4.14 — Centro maestro de una serie."""
    serie = db.get_or_404(Serie, serie_id)
    nombre_serie = (serie.nombre or "").strip()

    campeonatos = Campeonato.query.filter(
        db.func.lower(db.func.trim(Campeonato.serie)) == nombre_serie.lower()
    ).order_by(
        Campeonato.temporada.desc(), Campeonato.id.desc()
    ).all()

    club_ids = set()
    for campeonato in campeonatos:
        club_ids.update(
            x.club_id for x in CampeonatoClub.query.filter_by(
                campeonato_id=campeonato.id
            ).all()
        )

    clubes = Club.query.filter(
        Club.id.in_(club_ids)
    ).order_by(Club.nombre.asc()).all() if club_ids else []

    jugadores = Jugador.query.filter(
        Jugador.serie == nombre_serie
    ).order_by(
        Jugador.club.asc(), Jugador.nombre_completo.asc()
    ).all()

    registros = (
        db.session.query(PartidoJugador, Jugador, Partido, Campeonato, ActaPartido)
        .join(Jugador, Jugador.id == PartidoJugador.jugador_id)
        .join(Partido, Partido.id == PartidoJugador.partido_id)
        .join(Campeonato, Campeonato.id == Partido.campeonato_id)
        .join(ActaPartido, ActaPartido.partido_id == Partido.id)
        .filter(
            ActaPartido.estado == "Cerrada",
            db.func.lower(db.func.trim(Campeonato.serie)) == nombre_serie.lower()
        )
        .all()
    )

    partidos_ids = set()
    jugadores_stats_ids = set()
    total_goles = total_amarillas = total_rojas = 0
    por_club = {}

    for pj, jugador, partido, campeonato, acta in registros:
        partidos_ids.add(partido.id)
        jugadores_stats_ids.add(jugador.id)

        goles = int(pj.goles or 0)
        amarillas = int(pj.amarillas or 0)
        rojas = int(pj.rojas or 0)
        total_goles += goles
        total_amarillas += amarillas
        total_rojas += rojas

        nombre_club = (jugador.club or "Sin club").strip()
        fila = por_club.setdefault(nombre_club, {
            "nombre": nombre_club,
            "goles": 0,
            "amarillas": 0,
            "rojas": 0,
            "jugadores": set(),
        })
        fila["goles"] += goles
        fila["amarillas"] += amarillas
        fila["rojas"] += rojas
        fila["jugadores"].add(jugador.id)

    resumen_clubes = list(por_club.values())
    resumen_clubes.sort(key=lambda x: (-x["goles"], x["nombre"].lower()))
    for fila in resumen_clubes:
        fila["total_jugadores"] = len(fila["jugadores"])

    return render_template(
        "admin_centro_serie.html",
        serie=serie,
        campeonatos=campeonatos,
        clubes=clubes,
        jugadores=jugadores,
        resumen_clubes=resumen_clubes,
        resumen={
            "campeonatos": len(campeonatos),
            "clubes": len(clubes),
            "jugadores": len(jugadores),
            "vigentes": sum(1 for j in jugadores if (j.estado or "Vigente") == "Vigente"),
            "partidos": len(partidos_ids),
            "jugadores_con_estadisticas": len(jugadores_stats_ids),
            "goles": total_goles,
            "amarillas": total_amarillas,
            "rojas": total_rojas,
        },
    )


# ============================================================
# V7.4.4 — CENTRO MAESTRO DE CAMPEONATO
# Integra clubes, planteles, partidos, actas y estadísticas.
# Solo lectura; utiliza los registros maestros existentes.
# ============================================================

@app.route("/admin/campeonatos/<int:campeonato_id>/centro")
@admin_required
def admin_centro_campeonato(campeonato_id):
    campeonato = db.get_or_404(Campeonato, campeonato_id)

    participaciones = CampeonatoClub.query.filter_by(
        campeonato_id=campeonato.id
    ).all()
    clubes = sorted(
        [x.club for x in participaciones if x.club],
        key=lambda x: (x.nombre or "").lower()
    )

    # Planteles maestros: jugadores ya registrados para los clubes participantes.
    club_ids = {c.id for c in clubes}
    club_nombres = {c.nombre for c in clubes}
    serie_campeonato = (campeonato.serie or "").strip()
    jugadores_query = Jugador.query.filter(Jugador.club.in_(club_nombres)) if club_nombres else Jugador.query.filter(db.false())
    if serie_campeonato:
        jugadores_query = jugadores_query.filter(Jugador.serie == serie_campeonato)
    jugadores = jugadores_query.order_by(
        Jugador.club.asc(), Jugador.serie.asc(), Jugador.nombre_completo.asc()
    ).all()

    planteles = {}
    for jugador in jugadores:
        clave = (jugador.club or "Sin club", jugador.serie or "Sin serie")
        planteles.setdefault(clave, []).append(jugador)

    resumen_planteles = []
    for (nombre_club, nombre_serie), lista in planteles.items():
        resumen_planteles.append({
            "club": nombre_club,
            "serie": nombre_serie,
            "total": len(lista),
            "vigentes": sum(1 for j in lista if (j.estado or "Vigente") == "Vigente"),
            "jugadores": lista,
        })
    resumen_planteles.sort(key=lambda x: (x["club"].lower(), x["serie"].lower()))

    series = sorted({x["serie"] for x in resumen_planteles}, key=str.lower)

    partidos = Partido.query.filter_by(
        campeonato_id=campeonato.id
    ).order_by(
        Partido.jornada.asc(),
        Partido.fecha.asc().nullslast(),
        Partido.hora.asc().nullslast(),
        Partido.id.asc()
    ).all()

    # Estadísticas oficiales: solamente actas cerradas.
    registros = (
        db.session.query(PartidoJugador, Jugador, Partido, ActaPartido)
        .join(Jugador, Jugador.id == PartidoJugador.jugador_id)
        .join(Partido, Partido.id == PartidoJugador.partido_id)
        .join(ActaPartido, ActaPartido.partido_id == Partido.id)
        .filter(
            Partido.campeonato_id == campeonato.id,
            ActaPartido.estado == "Cerrada"
        )
        .all()
    )

    jugadores_ids = set()
    partidos_cerrados = set()
    total_goles = total_amarillas = total_rojas = 0
    por_club = {}
    por_jugador = {}

    for pj, jugador, partido, acta in registros:
        jugadores_ids.add(jugador.id)
        partidos_cerrados.add(partido.id)

        goles = int(pj.goles or 0)
        amarillas = int(pj.amarillas or 0)
        rojas = int(pj.rojas or 0)
        total_goles += goles
        total_amarillas += amarillas
        total_rojas += rojas

        nombre_club = (jugador.club or "Sin club").strip()
        fila_club = por_club.setdefault(nombre_club, {
            "nombre": nombre_club,
            "jugadores": set(),
            "goles": 0,
            "amarillas": 0,
            "rojas": 0,
        })
        fila_club["jugadores"].add(jugador.id)
        fila_club["goles"] += goles
        fila_club["amarillas"] += amarillas
        fila_club["rojas"] += rojas

        key = jugador.id
        fila_jugador = por_jugador.setdefault(key, {
            "jugador": jugador,
            "partidos": set(),
            "goles": 0,
            "amarillas": 0,
            "rojas": 0,
        })
        fila_jugador["partidos"].add(partido.id)
        fila_jugador["goles"] += goles
        fila_jugador["amarillas"] += amarillas
        fila_jugador["rojas"] += rojas

    resumen_clubes = []
    for fila in por_club.values():
        fila["total_jugadores"] = len(fila["jugadores"])
        resumen_clubes.append(fila)

    resumen_clubes.sort(key=lambda x: (-(x["goles"]), x["nombre"].lower()))

    goleadores = [
        x for x in por_jugador.values()
        if x["goles"] > 0
    ]
    goleadores.sort(
        key=lambda x: (-x["goles"], x["jugador"].nombre_completo.lower())
    )

    disciplina = [
        x for x in por_jugador.values()
        if x["amarillas"] > 0 or x["rojas"] > 0
    ]
    disciplina.sort(
        key=lambda x: (
            -x["amarillas"],
            -x["rojas"],
            x["jugador"].nombre_completo.lower()
        )
    )

    return render_template(
        "admin_centro_campeonato.html",
        campeonato=campeonato,
        clubes=clubes,
        partidos=partidos,
        resumen_clubes=resumen_clubes,
        resumen_planteles=resumen_planteles,
        series=series,
        serie_campeonato=serie_campeonato,
        goleadores=goleadores[:30],
        disciplina=disciplina[:30],
        resumen={
            "clubes": len(clubes),
            "partidos": len(partidos),
            "finalizados": sum(1 for p in partidos if p.estado == "Finalizado"),
            "pendientes": sum(1 for p in partidos if p.estado != "Finalizado"),
            "actas_cerradas": len(partidos_cerrados),
            "jugadores": len(jugadores_ids),
            "goles": total_goles,
            "amarillas": total_amarillas,
            "rojas": total_rojas,
        },
    )


# ============================================================
# V7.4.6 — CENTRO DE PLANTELES POR CAMPEONATO
# ============================================================

@app.route("/admin/campeonatos/<int:campeonato_id>/planteles")
@admin_required
def admin_planteles_campeonato(campeonato_id):
    campeonato = db.get_or_404(Campeonato, campeonato_id)

    participaciones = CampeonatoClub.query.filter_by(
        campeonato_id=campeonato.id
    ).all()

    clubes = sorted(
        [x.club for x in participaciones if x.club],
        key=lambda x: (x.nombre or "").lower()
    )

    serie = (campeonato.serie or "").strip()

    filas = []
    for club in clubes:
        jugadores = Jugador.query.filter(
            Jugador.club == club.nombre,
            Jugador.serie == serie
        ).order_by(
            Jugador.nombre_completo.asc()
        ).all()

        filas.append({
            "club": club,
            "serie": serie or "Sin serie",
            "jugadores": jugadores,
            "total": len(jugadores),
            "vigentes": sum(
                1 for j in jugadores
                if (j.estado or "Vigente") == "Vigente"
            ),
        })

    return render_template(
        "admin_planteles_campeonato.html",
        campeonato=campeonato,
        filas=filas,
        serie=serie,
    )


@app.route("/admin/campeonatos/<int:campeonato_id>/plantel/<int:club_id>")
@admin_required
def admin_plantel_campeonato_club(campeonato_id, club_id):
    campeonato = db.get_or_404(Campeonato, campeonato_id)
    club = db.get_or_404(Club, club_id)

    participa = CampeonatoClub.query.filter_by(
        campeonato_id=campeonato.id,
        club_id=club.id
    ).first()

    if not participa:
        flash("El club no participa en este campeonato.", "warning")
        return redirect(
            url_for(
                "admin_planteles_campeonato",
                campeonato_id=campeonato.id
            )
        )

    serie = (campeonato.serie or "").strip()

    jugadores = Jugador.query.filter(
        Jugador.club == club.nombre,
        Jugador.serie == serie
    ).order_by(
        Jugador.nombre_completo.asc()
    ).all()

    # Estadísticas individuales del plantel: solo actas cerradas.
    jugador_ids = {j.id for j in jugadores}
    stats = {
        j.id: {
            "partidos": 0,
            "titulares": 0,
            "goles": 0,
            "amarillas": 0,
            "rojas": 0,
            "actas": 0,
        }
        for j in jugadores
    }

    if jugador_ids:
        registros = (
            db.session.query(PartidoJugador, Partido, ActaPartido)
            .join(Partido, Partido.id == PartidoJugador.partido_id)
            .join(ActaPartido, ActaPartido.partido_id == Partido.id)
            .filter(
                Partido.campeonato_id == campeonato.id,
                ActaPartido.estado == "Cerrada",
                PartidoJugador.jugador_id.in_(jugador_ids),
            )
            .all()
        )

        for pj, partido, acta in registros:
            fila = stats.get(pj.jugador_id)
            if not fila:
                continue
            fila["partidos"] += 1
            fila["titulares"] += (
                1 if str(pj.condicion or "").strip().lower() == "titular" else 0
            )
            fila["goles"] += int(pj.goles or 0)
            fila["amarillas"] += int(pj.amarillas or 0)
            fila["rojas"] += int(pj.rojas or 0)
            fila["actas"] += 1

    return render_template(
        "admin_plantel_campeonato_club.html",
        campeonato=campeonato,
        club=club,
        serie=serie or "Sin serie",
        jugadores=jugadores,
        stats=stats,
    )


# ============================================================
# V7.4.11 — CENTRO MAESTRO DE PARTIDO
# Partido → Clubes → Plantel → Acta → Estadísticas.
# Solo lectura; utiliza los registros maestros existentes.
# ============================================================

@app.route("/admin/partidos/<int:partido_id>/centro")
@admin_required
def admin_centro_partido(partido_id):
    partido = db.get_or_404(Partido, partido_id)

    acta = partido.acta

    nomina = (
        PartidoJugador.query
        .join(Jugador, PartidoJugador.jugador_id == Jugador.id)
        .filter(PartidoJugador.partido_id == partido.id)
        .order_by(
            PartidoJugador.equipo.asc(),
            Jugador.nombre_completo.asc()
        )
        .all()
    )

    local = [x for x in nomina if (x.equipo or "").strip().lower() in ("local", "home")]
    visitante = [x for x in nomina if (x.equipo or "").strip().lower() in ("visitante", "visita", "away")]

    # Estadísticas del partido: las oficiales se muestran solo si el acta está cerrada.
    acta_cerrada = bool(acta and (acta.estado or "").strip().lower() == "cerrada")
    if not acta_cerrada:
        local = []
        visitante = []

    def resumen_nomina(lista):
        return {
            "jugadores": len(lista),
            "titulares": sum(
                1 for x in lista
                if str(x.condicion or "").strip().lower() == "titular"
            ),
            "goles": sum(int(x.goles or 0) for x in lista),
            "amarillas": sum(int(x.amarillas or 0) for x in lista),
            "rojas": sum(int(x.rojas or 0) for x in lista),
        }

    resumen_local = resumen_nomina(local)
    resumen_visitante = resumen_nomina(visitante)

    return render_template(
        "admin_centro_partido.html",
        partido=partido,
        acta=acta,
        acta_cerrada=acta_cerrada,
        local=local,
        visitante=visitante,
        resumen_local=resumen_local,
        resumen_visitante=resumen_visitante,
    )


# ============================================================
# V7.2 — CENTRO MAESTRO DE PARTIDOS
# Vista unificada de partidos existentes y sus actas.
# Solo consulta; no crea ni modifica registros.
# ============================================================

@app.route("/admin/partidos/centro")
@admin_required
def admin_centro_partidos():
    campeonatos = Campeonato.query.order_by(
        Campeonato.temporada.desc(), Campeonato.id.desc()
    ).all()

    campeonato_id = request.args.get("campeonato_id", type=int)
    estado = (request.args.get("estado") or "").strip()
    q = (request.args.get("q") or "").strip()

    query = (
        Partido.query
        .join(Campeonato)
        .order_by(
            Partido.fecha.desc().nullslast(),
            Partido.jornada.desc(),
            Partido.id.desc()
        )
    )

    if campeonato_id:
        query = query.filter(Partido.campeonato_id == campeonato_id)

    if estado:
        query = query.filter(Partido.estado == estado)

    if q:
        patron = f"%{q}%"
        query = query.filter(
            db.or_(
                Club.nombre.ilike(patron),
                db.exists().where(
                    db.and_(
                        Club.id == Partido.local_club_id,
                        Club.nombre.ilike(patron)
                    )
                ),
                db.exists().where(
                    db.and_(
                        Club.id == Partido.visitante_club_id,
                        Club.nombre.ilike(patron)
                    )
                )
            )
        )

    partidos = query.limit(250).all()

    filas = []
    for partido in partidos:
        acta = partido.acta
        filas.append({
            "partido": partido,
            "acta": acta,
            "acta_estado": acta.estado if acta else "Sin acta",
        })

    total = Partido.query.count()
    programados = Partido.query.filter(Partido.estado != "Finalizado").count()
    finalizados = Partido.query.filter_by(estado="Finalizado").count()
    con_acta = ActaPartido.query.count()
    actas_cerradas = ActaPartido.query.filter_by(estado="Cerrada").count()

    return render_template(
        "admin_centro_partidos.html",
        filas=filas,
        campeonatos=campeonatos,
        campeonato_id=campeonato_id,
        estado=estado,
        q=q,
        resumen={
            "total": total,
            "programados": programados,
            "finalizados": finalizados,
            "con_acta": con_acta,
            "actas_cerradas": actas_cerradas,
        },
    )


# ============================================================
# V7.4 — CENTRO MAESTRO DE ESTADÍSTICAS
# Estadísticas oficiales derivadas exclusivamente de actas cerradas.
# No crea tablas ni registros paralelos.
# ============================================================

@app.route("/admin/estadisticas/centro")
@admin_required
def admin_centro_estadisticas():
    campeonatos = Campeonato.query.order_by(
        Campeonato.temporada.desc(), Campeonato.id.desc()
    ).all()

    campeonato_id = request.args.get("campeonato_id", type=int)
    serie = (request.args.get("serie") or "").strip()
    club = (request.args.get("club") or "").strip()
    q = (request.args.get("q") or "").strip()

    series = sorted({(c.serie or "").strip() for c in campeonatos if (c.serie or "").strip()}, key=str.lower)
    clubes = sorted({(j.club or "").strip() for j in Jugador.query.all() if (j.club or "").strip()}, key=str.lower)

    base = (
        db.session.query(PartidoJugador, Jugador, Partido, Campeonato, ActaPartido)
        .join(Jugador, Jugador.id == PartidoJugador.jugador_id)
        .join(Partido, Partido.id == PartidoJugador.partido_id)
        .join(Campeonato, Campeonato.id == Partido.campeonato_id)
        .join(ActaPartido, ActaPartido.partido_id == Partido.id)
        .filter(ActaPartido.estado == "Cerrada")
    )

    if campeonato_id:
        base = base.filter(Campeonato.id == campeonato_id)
    if serie:
        base = base.filter(Campeonato.serie == serie)
    if club:
        base = base.filter(Jugador.club == club)
    if q:
        patron = f"%{q}%"
        base = base.filter(
            db.or_(Jugador.nombre_completo.ilike(patron), Jugador.rut.ilike(patron))
        )

    registros = base.all()

    # Agregación en Python para mantener el template simple y evitar
    # expresiones Jinja complejas que puedan romper la vista.
    por_jugador = {}
    partidos_ids = set()
    jugadores_ids = set()
    total_goles = 0
    total_amarillas = 0
    total_rojas = 0
    total_titulares = 0

    for pj, jugador, partido, campeonato, acta in registros:
        partidos_ids.add(partido.id)
        jugadores_ids.add(jugador.id)
        goles = int(pj.goles or 0)
        amarillas = int(pj.amarillas or 0)
        rojas = int(pj.rojas or 0)
        titular = 1 if str(pj.condicion or "").strip().lower() == "titular" else 0

        total_goles += goles
        total_amarillas += amarillas
        total_rojas += rojas
        total_titulares += titular

        key = (jugador.id, campeonato.id)
        if key not in por_jugador:
            por_jugador[key] = {
                "jugador": jugador,
                "campeonato": campeonato,
                "partidos": set(),
                "titulares": 0,
                "goles": 0,
                "amarillas": 0,
                "rojas": 0,
            }
        fila = por_jugador[key]
        fila["partidos"].add(partido.id)
        fila["titulares"] += titular
        fila["goles"] += goles
        fila["amarillas"] += amarillas
        fila["rojas"] += rojas

    jugadores_stats = []
    for fila in por_jugador.values():
        fila["partidos_jugados"] = len(fila["partidos"])
        jugadores_stats.append(fila)

    goleadores = sorted(
        jugadores_stats,
        key=lambda x: (-x["goles"], x["jugador"].nombre_completo.lower())
    )
    disciplina = sorted(
        jugadores_stats,
        key=lambda x: (-x["amarillas"], -x["rojas"], x["jugador"].nombre_completo.lower())
    )
    rojas_ranking = sorted(
        jugadores_stats,
        key=lambda x: (-x["rojas"], x["jugador"].nombre_completo.lower())
    )

    # Solo mostrar jugadores con producción estadística real.
    goleadores = [x for x in goleadores if x["goles"] > 0]
    disciplina = [x for x in disciplina if x["amarillas"] > 0 or x["rojas"] > 0]
    rojas_ranking = [x for x in rojas_ranking if x["rojas"] > 0]

    return render_template(
        "admin_centro_estadisticas.html",
        campeonatos=campeonatos,
        campeonato_id=campeonato_id,
        series=series,
        serie=serie,
        clubes=clubes,
        club=club,
        q=q,
        goleadores=goleadores,
        disciplina=disciplina,
        rojas=rojas_ranking,
        resumen={
            "partidos": len(partidos_ids),
            "jugadores": len(jugadores_ids),
            "goles": total_goles,
            "amarillas": total_amarillas,
            "rojas": total_rojas,
            "titulares": total_titulares,
            "registros": len(registros),
        },
    )


# ============================================================
# V9.4 — HISTORIA DE LA ASOCIACIÓN
# Historial derivado de campeonatos y actas cerradas.
# No crea registros paralelos.
# ============================================================

@app.route("/admin/historia")
@admin_required
def admin_historia():
    campeonatos = Campeonato.query.order_by(
        Campeonato.temporada.desc(),
        Campeonato.fecha_inicio.desc().nullslast(),
        Campeonato.id.desc()
    ).all()

    historial = []
    titulos = {}

    for campeonato in campeonatos:
        clubes_participantes = CampeonatoClub.query.filter_by(
            campeonato_id=campeonato.id
        ).all()

        partidos = Partido.query.filter_by(
            campeonato_id=campeonato.id
        ).order_by(
            Partido.jornada.asc(),
            Partido.id.asc()
        ).all()

        finalizados = [
            p for p in partidos
            if p.estado == "Finalizado"
            and p.goles_local is not None
            and p.goles_visitante is not None
        ]

        filas = {}
        for registro in clubes_participantes:
            if not registro.club:
                continue
            filas[registro.club.id] = {
                "club": registro.club,
                "pj": 0,
                "pg": 0,
                "pe": 0,
                "pp": 0,
                "gf": 0,
                "gc": 0,
                "dg": 0,
                "pts": 0,
            }

        for partido in finalizados:
            local = filas.get(partido.local_club_id)
            visitante = filas.get(partido.visitante_club_id)
            if not local or not visitante:
                continue

            gl = int(partido.goles_local or 0)
            gv = int(partido.goles_visitante or 0)
            local["pj"] += 1
            visitante["pj"] += 1
            local["gf"] += gl
            local["gc"] += gv
            visitante["gf"] += gv
            visitante["gc"] += gl

            if gl > gv:
                local["pg"] += 1
                local["pts"] += 3
                visitante["pp"] += 1
            elif gl < gv:
                visitante["pg"] += 1
                visitante["pts"] += 3
                local["pp"] += 1
            else:
                local["pe"] += 1
                visitante["pe"] += 1
                local["pts"] += 1
                visitante["pts"] += 1

        for fila in filas.values():
            fila["dg"] = fila["gf"] - fila["gc"]

        tabla = sorted(
            filas.values(),
            key=lambda x: (
                -x["pts"],
                -x["dg"],
                -x["gf"],
                x["club"].nombre.lower(),
            )
        )

        campeonato_completo = bool(partidos) and len(finalizados) == len(partidos)
        es_finalizado = (campeonato.estado or "").strip().lower() == "finalizado" or campeonato_completo
        campeon = tabla[0] if es_finalizado and tabla and tabla[0]["pj"] > 0 else None

        goleador = None
        registros_goles = (
            db.session.query(Jugador, db.func.coalesce(db.func.sum(Gol.cantidad), 0))
            .join(Gol, Gol.jugador_id == Jugador.id)
            .filter(Gol.campeonato_id == campeonato.id)
            .group_by(Jugador.id)
            .order_by(db.func.sum(Gol.cantidad).desc(), Jugador.nombre_completo.asc())
            .first()
        )
        if registros_goles:
            goleador = {
                "jugador": registros_goles[0],
                "goles": int(registros_goles[1] or 0),
            }

        if campeon:
            nombre_campeon = campeon["club"].nombre
            titulos[nombre_campeon] = titulos.get(nombre_campeon, 0) + 1

        historial.append({
            "campeonato": campeonato,
            "clubes": len(clubes_participantes),
            "partidos": len(partidos),
            "finalizados": len(finalizados),
            "completo": campeonato_completo,
            "es_finalizado": es_finalizado,
            "campeon": campeon,
            "tabla": tabla[:3],
            "goleador": goleador,
        })

    titulos_ranking = sorted(
        [{"club": nombre, "titulos": total} for nombre, total in titulos.items()],
        key=lambda x: (-x["titulos"], x["club"].lower())
    )

    goleadores_historicos = (
        db.session.query(
            Jugador,
            db.func.coalesce(db.func.sum(Gol.cantidad), 0).label("goles")
        )
        .join(Gol, Gol.jugador_id == Jugador.id)
        .group_by(Jugador.id)
        .order_by(db.func.sum(Gol.cantidad).desc(), Jugador.nombre_completo.asc())
        .limit(10)
        .all()
    )

    total_campeonatos = len(campeonatos)
    campeonatos_finalizados = sum(1 for x in historial if x["es_finalizado"])
    total_clubes = len({
        registro.club_id
        for campeonato in campeonatos
        for registro in CampeonatoClub.query.filter_by(campeonato_id=campeonato.id).all()
    })
    total_partidos = sum(x["partidos"] for x in historial)

    return render_template(
        "admin_historia.html",
        historial=historial,
        titulos_ranking=titulos_ranking,
        goleadores_historicos=goleadores_historicos,
        resumen={
            "campeonatos": total_campeonatos,
            "finalizados": campeonatos_finalizados,
            "clubes": total_clubes,
            "partidos": total_partidos,
        },
    )


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
