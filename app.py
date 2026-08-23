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

    # --------------------------------------------------------
    # FOTOGRAFÍA
    # --------------------------------------------------------

    foto = db.Column(
        db.LargeBinary,
        nullable=True
    )

    foto_tipo = db.Column(
        db.String(50),
        nullable=True
    )

    # --------------------------------------------------------
    # ESTADO DEL JUGADOR
    # --------------------------------------------------------

    estado = db.Column(
        db.String(30),
        nullable=False,
        default="HABILITADO"
    )

    motivo_estado = db.Column(
        db.String(300),
        nullable=True
    )

    fecha_estado = db.Column(
        db.DateTime,
        nullable=True
    )


# ============================================================
# CREACIÓN / ACTUALIZACIÓN SEGURA DE BASE DE DATOS
# ============================================================

def preparar_base_datos():

    db.create_all()

    try:

        inspector = db.inspect(db.engine)

        columnas = [
            columna["name"]
            for columna in inspector.get_columns("jugador")
        ]

        # ----------------------------------------------------
        # Agregar columnas nuevas si no existen
        # ----------------------------------------------------

        columnas_nuevas = {
            "foto_tipo": "VARCHAR(50)",
            "estado": "VARCHAR(30)",
            "motivo_estado": "VARCHAR(300)",
            "fecha_estado": "TIMESTAMP"
        }

        for nombre_columna, tipo_columna in columnas_nuevas.items():

            if nombre_columna not in columnas:

                if db.engine.dialect.name == "postgresql":

                    db.session.execute(
                        db.text(
                            f"""
                            ALTER TABLE jugador
                            ADD COLUMN IF NOT EXISTS
                            {nombre_columna}
                            {tipo_columna}
                            """
                        )
                    )

                elif db.engine.dialect.name == "sqlite":

                    db.session.execute(
                        db.text(
                            f"""
                            ALTER TABLE jugador
                            ADD COLUMN
                            {nombre_columna}
                            {tipo_columna}
                            """
                        )
                    )

        db.session.commit()

        # ----------------------------------------------------
        # Actualizar jugadores antiguos
        # ----------------------------------------------------

        try:

            db.session.execute(
                db.text(
                    """
                    UPDATE jugador
                    SET estado = 'HABILITADO'
                    WHERE estado IS NULL
                    """
                )
            )

            db.session.commit()

        except Exception:

            db.session.rollback()

    except Exception as error:

        db.session.rollback()

        print(
            "Advertencia al preparar la base de datos:",
            error
        )


with app.app_context():

    preparar_base_datos()


# ============================================================
# FUNCIONES AUXILIARES
# ============================================================

def normalizar_rut(rut):

    if rut is None:
        return ""

    rut = str(rut).strip().upper()

    rut = rut.replace(" ", "")

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
            "." +
            cuerpo[-3:] +
            cuerpo_formateado
        )

        cuerpo = cuerpo[:-3]

    cuerpo_formateado = (
        cuerpo +
        cuerpo_formateado
    )

    return f"{cuerpo_formateado}-{dv}"


# ============================================================
# VALIDAR RUT
# ============================================================

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
            int(digito) *
            multiplicador
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


# ============================================================
# CONVERTIR FECHA
# ============================================================

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


# ============================================================
# OBTENER FOTOGRAFÍA
# ============================================================

def obtener_fotografia():

    archivo = request.files.get("foto")

    if not archivo:
        return None, None, None

    if not archivo.filename:
        return None, None, None

    contenido = archivo.read()

    if not contenido:

        return (
            None,
            None,
            "La fotografía seleccionada está vacía."
        )

    # Máximo 5 MB

    maximo = 5 * 1024 * 1024

    if len(contenido) > maximo:

        return (
            None,
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
            None,
            "La fotografía debe ser JPG, PNG o WEBP."
        )

    return contenido, tipo, None


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

    mimetype = jugador.foto_tipo or "image/jpeg"

    return Response(
        jugador.foto,
        mimetype=mimetype
    )


# ============================================================
# NUEVO JUGADOR
# ============================================================

@app.route(
    "/jugadores/nuevo",
    methods=["GET", "POST"]
)
def nuevo_jugador():

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
                jugador=None
            )

        if not validar_rut(rut):

            flash(
                "El RUT ingresado no es válido.",
                "error"
            )

            return render_template(
                "jugador_form.html",
                jugador=None
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
                jugador=None
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
                jugador=None
            )

        foto, foto_tipo, error_foto = obtener_fotografia()

        if error_foto:

            flash(
                error_foto,
                "error"
            )

            return render_template(
                "jugador_form.html",
                jugador=None
            )

        jugador = Jugador(
            rut=rut,
            nombre_completo=nombre,
            fecha_nacimiento=fecha_obj,
            serie=serie,
            club=club,
            foto=foto,
            foto_tipo=foto_tipo,
            estado="HABILITADO",
            motivo_estado="Jugador registrado",
            fecha_estado=datetime.utcnow()
        )

        db.session.add(jugador)

        db.session.commit()

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
        jugador=None
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
                jugador=jugador
            )

        if not validar_rut(rut):

            flash(
                "El RUT ingresado no es válido.",
                "error"
            )

            return render_template(
                "jugador_form.html",
                jugador=jugador
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
                jugador=jugador
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
                jugador=jugador
            )

        archivo_foto = request.files.get(
            "foto"
        )

        if (
            archivo_foto
            and archivo_foto.filename
        ):

            foto, foto_tipo, error_foto = obtener_fotografia()

            if error_foto:

                flash(
                    error_foto,
                    "error"
                )

                return render_template(
                    "jugador_form.html",
                    jugador=jugador
                )

            jugador.foto = foto

            jugador.foto_tipo = foto_tipo

        jugador.rut = rut

        jugador.nombre_completo = nombre

        jugador.fecha_nacimiento = fecha_obj

        jugador.serie = serie

        jugador.club = club

        db.session.commit()

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
        jugador=jugador
    )


# ============================================================
# CAMBIAR ESTADO DEL JUGADOR
# ============================================================

@app.route(
    "/jugadores/<int:jugador_id>/estado",
    methods=["POST"]
)
def cambiar_estado_jugador(jugador_id):

    jugador = db.get_or_404(
        Jugador,
        jugador_id
    )

    estado = request.form.get(
        "estado",
        ""
    ).strip().upper()

    motivo = request.form.get(
        "motivo_estado",
        ""
    ).strip()

    estados_validos = {
        "HABILITADO",
        "SUSPENDIDO",
        "PENDIENTE"
    }

    if estado not in estados_validos:

        flash(
            "Estado de jugador no válido.",
            "error"
        )

        return redirect(
            url_for(
                "ficha_jugador",
                jugador_id=jugador.id
            )
        )

    if estado == "SUSPENDIDO" and not motivo:

        flash(
            "Debes indicar el motivo de la suspensión.",
            "error"
        )

        return redirect(
            url_for(
                "ficha_jugador",
                jugador_id=jugador.id
            )
        )

    jugador.estado = estado

    jugador.motivo_estado = (
        motivo
        if motivo
        else "Sin observaciones"
    )

    jugador.fecha_estado = datetime.utcnow()

    db.session.commit()

    flash(
        f"Estado actualizado a {estado}.",
        "success"
    )

    return redirect(
        url_for(
            "ficha_jugador",
            jugador_id=jugador.id
        )
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

    db.session.delete(jugador)

    db.session.commit()

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
                +
                ", ".join(faltantes),
                "error"
            )

            return redirect(
                url_for("importar_jugadores")
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
                    estado="HABILITADO",
                    motivo_estado="Importado desde Excel",
                    fecha_estado=datetime.utcnow()
                )

                db.session.add(jugador)

                registrados += 1

            except Exception as error:

                errores += 1

                detalle_errores.append(
                    f"Fila {numero_fila}: "
                    f"error al procesar ({error})."
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
                url_for("importar_jugadores")
            )

        return render_template(
            "importar_resultado.html",
            registrados=registrados,
            duplicados=duplicados,
            errores=errores,
            detalle_errores=detalle_errores
        )

    except Exception:

        db.session.rollback()

        flash(
            "Ocurrió un error inesperado "
            "al importar el archivo.",
            "error"
        )

        return redirect(
            url_for("importar_jugadores")
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
                jugador.estado,

            "motivo_estado":
                jugador.motivo_estado,

            "fecha_estado":
                jugador.fecha_estado.isoformat()
                if jugador.fecha_estado
                else None,

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

    url_credencial = url_for(
        "credencial_jugador",
        jugador_id=jugador.id,
        _external=True
    )

    imagen_qr = qrcode.make(
        url_credencial
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
# CREDENCIAL DEL JUGADOR
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
