import os
from datetime import date, datetime
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
    Response
)

from flask_sqlalchemy import SQLAlchemy
from openpyxl import load_workbook


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
# MODELO CLUB
# ============================================================

class Club(db.Model):

    __tablename__ = "club"

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


# ============================================================
# MODELO SERIE
# ============================================================

class Serie(db.Model):

    __tablename__ = "serie"

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
# MODELO JUGADOR
# ============================================================

class Jugador(db.Model):

    __tablename__ = "jugador"

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
# CREACIÓN / ACTUALIZACIÓN SEGURA DE BASE DE DATOS
# ============================================================

def preparar_base_datos():

    try:

        # ----------------------------------------------------
        # CREAR TABLAS QUE NO EXISTAN
        # ----------------------------------------------------

        db.create_all()

        inspector = db.inspect(
            db.engine
        )

        # ----------------------------------------------------
        # VERIFICAR TABLA JUGADOR
        # ----------------------------------------------------

        tablas = inspector.get_table_names()

        if "jugador" not in tablas:

            print(
                "La tabla jugador no existe."
            )

            return

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

        print(
            "Base de datos preparada correctamente."
        )

    except Exception as error:

        db.session.rollback()

        print(
            "Advertencia al preparar la base de datos:",
            error
        )


# ============================================================
# PREPARAR BASE DE DATOS
# ============================================================

with app.app_context():

    preparar_base_datos()


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
                ),

                Jugador.serie.ilike(
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

    return render_template(
        "jugador_detalle.html",
        jugador=jugador
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

    # --------------------------------------------------------
    # CATÁLOGOS ACTIVOS
    # --------------------------------------------------------

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

        except Exception as error:

            db.session.rollback()

            print(
                "Error guardando importación:",
                error
            )

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
                bool(jugador.foto)
        }

        for jugador in jugadores

    ])


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

@app.route(
    "/dashboard"
)
def dashboard():

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

    total_clubes = Club.query.count()

    clubes_activos = Club.query.filter_by(
        activo=True
    ).count()

    total_series = Serie.query.count()

    series_activas = Serie.query.filter_by(
        activo=True
    ).count()

    return render_template(
        "dashboard.html",
        total_jugadores=total_jugadores,
        vigentes=vigentes,
        pendientes=pendientes,
        suspendidos=suspendidos,
        inhabilitados=inhabilitados,
        total_clubes=total_clubes,
        clubes_activos=clubes_activos,
        total_series=total_series,
        series_activas=series_activas
    )


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

    try:

        db.session.add(
            club
        )

        db.session.commit()

    except Exception as error:

        db.session.rollback()

        print(
            "Error creando club:",
            error
        )

        flash(
            "No fue posible crear el club.",
            "error"
        )

        return redirect(
            url_for("configuracion")
        )

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

    try:

        db.session.add(
            serie
        )

        db.session.commit()

    except Exception as error:

        db.session.rollback()

        print(
            "Error creando serie:",
            error
        )

        flash(
            "No fue posible crear la serie.",
            "error"
        )

        return redirect(
            url_for("configuracion")
        )

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

    try:

        club.activo = not club.activo

        db.session.commit()

    except Exception as error:

        db.session.rollback()

        print(
            "Error cambiando estado del club:",
            error
        )

        flash(
            "No fue posible cambiar el estado del club.",
            "error"
        )

        return redirect(
            url_for("configuracion")
        )

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

    try:

        serie.activo = not serie.activo

        db.session.commit()

    except Exception as error:

        db.session.rollback()

        print(
            "Error cambiando estado de la serie:",
            error
        )

        flash(
            "No fue posible cambiar el estado de la serie.",
            "error"
        )

        return redirect(
            url_for("configuracion")
        )

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
# API DE CLUBES
# ============================================================

@app.route(
    "/api/clubes"
)
def api_clubes():

    clubes = Club.query.filter_by(
        activo=True
    ).order_by(
        Club.nombre
    ).all()

    return jsonify([

        {
            "id": club.id,
            "nombre": club.nombre,
            "activo": club.activo
        }

        for club in clubes

    ])


# ============================================================
# API DE SERIES
# ============================================================

@app.route(
    "/api/series"
)
def api_series():

    series = Serie.query.filter_by(
        activo=True
    ).order_by(
        Serie.nombre
    ).all()

    return jsonify([

        {
            "id": serie.id,
            "nombre": serie.nombre,
            "activo": serie.activo
        }

        for serie in series

    ])


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
