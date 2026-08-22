import os
from datetime import date, datetime
from flask import Flask, render_template, request, redirect, url_for, flash, jsonify
from flask_sqlalchemy import SQLAlchemy
from werkzeug.utils import secure_filename
from openpyxl import load_workbook

app = Flask(__name__)

app.config["SECRET_KEY"] = os.environ.get(
    "SECRET_KEY",
    "cambia-esta-clave-en-produccion"
)

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


with app.app_context():
    db.create_all()


# ============================================================
# FUNCIONES AUXILIARES
# ============================================================

def normalizar_rut(rut):
    """
    Convierte distintos formatos de RUT al formato:
    11.111.111-1
    """

    if rut is None:
        return ""

    rut = str(rut).strip().upper()

    # Eliminar espacios
    rut = rut.replace(" ", "")

    # Eliminar puntos y guion
    rut_limpio = rut.replace(".", "").replace("-", "")

    if len(rut_limpio) < 2:
        return ""

    cuerpo = rut_limpio[:-1]
    dv = rut_limpio[-1]

    if not cuerpo.isdigit():
        return ""

    # Formatear con puntos
    cuerpo_formateado = ""

    while len(cuerpo) > 3:
        cuerpo_formateado = "." + cuerpo[-3:] + cuerpo_formateado
        cuerpo = cuerpo[:-3]

    cuerpo_formateado = cuerpo + cuerpo_formateado

    return f"{cuerpo_formateado}-{dv}"


def validar_rut(rut):
    """
    Valida el dígito verificador de un RUT chileno.
    """

    if not rut:
        return False

    rut_limpio = (
        rut.replace(".", "")
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
        suma += int(digito) * multiplicador

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
    """
    Convierte fechas provenientes de Excel
    a datetime.date.
    """

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
# INICIO / LISTADO DE JUGADORES
# ============================================================

@app.route("/")
def index():

    q = request.args.get("q", "").strip()

    query = Jugador.query

    if q:
        query = query.filter(
            db.or_(
                Jugador.rut.ilike(f"%{q}%"),
                Jugador.nombre_completo.ilike(f"%{q}%"),
                Jugador.club.ilike(f"%{q}%")
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
# NUEVO JUGADOR
# ============================================================

@app.route(
    "/jugadores/nuevo",
    methods=["GET", "POST"]
)
def nuevo_jugador():

    if request.method == "POST":

        rut = normalizar_rut(
            request.form.get("rut", "")
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
            fecha_obj = date.fromisoformat(fecha)

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

        jugador = Jugador(
            rut=rut,
            nombre_completo=nombre,
            fecha_nacimiento=fecha_obj,
            serie=serie,
            club=club
        )

        db.session.add(jugador)
        db.session.commit()

        flash(
            "Jugador registrado correctamente.",
            "success"
        )

        return redirect(
            url_for("index")
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
            request.form.get("rut", "")
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
            fecha_obj = date.fromisoformat(fecha)

        except ValueError:

            flash(
                "Fecha no válida.",
                "error"
            )

            return render_template(
                "jugador_form.html",
                jugador=jugador
            )

        jugador.rut = rut
        jugador.nombre_completo = nombre
        jugador.fecha_nacimiento = fecha_obj
        jugador.serie = serie
        jugador.club = club

        db.session.commit()

        flash(
            "Datos actualizados.",
            "success"
        )

        return redirect(
            url_for("index")
        )

    return render_template(
        "jugador_form.html",
        jugador=jugador
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

@app.route("/jugadores/importar", methods=["GET", "POST"])
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

        # ----------------------------------------------------
        # Buscar columnas
        # ----------------------------------------------------

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

        for nombre_columna, posibles in equivalencias.items():

            for posible in posibles:

                if posible in encabezados:

                    columnas[nombre_columna] = encabezados.index(
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
                url_for("importar_jugadores")
            )

        # ----------------------------------------------------
        # Contadores
        # ----------------------------------------------------

        registrados = 0
        duplicados = 0
        errores = 0

        detalle_errores = []

        ruts_archivo = set()

        # ----------------------------------------------------
        # Procesar filas
        # ----------------------------------------------------

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

                # --------------------------------------------
                # Validaciones
                # --------------------------------------------

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
                    club=club
                )

                db.session.add(jugador)

                registrados += 1

            except Exception as error:

                errores += 1

                detalle_errores.append(
                    f"Fila {numero_fila}: "
                    f"error al procesar."
                )

        # ----------------------------------------------------
        # Guardar todo
        # ----------------------------------------------------

        try:

            db.session.commit()

        except Exception:

            db.session.rollback()

            flash(
                "Ocurrió un error al guardar los jugadores.",
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
    except Exception as error:
        db.session.rollback()

        flash(
            "Ocurrió un error inesperado al importar el archivo.",
            "error"
        )

        return redirect(
            url_for("importar_jugadores")
        )

# ============================================================
# API
# ============================================================

@app.route("/api/jugadores")
def api_jugadores():

    rut = request.args.get(
        "rut",
        ""
    ).strip().upper()

    if not rut:
        return jsonify([])

    jugadores = Jugador.query.filter(
        Jugador.rut.ilike(f"%{rut}%")
    ).order_by(
        Jugador.nombre_completo
    ).all()

    return jsonify([
        {
            "id": j.id,
            "rut": j.rut,
            "nombre_completo": j.nombre_completo,
            "fecha_nacimiento": j.fecha_nacimiento.isoformat(),
            "serie": j.serie,
            "club": j.club
        }
        for j in jugadores
    ])


@app.route("/health")
def health():
    return {"status": "ok"}


if __name__ == "__main__":
    app.run(
        host="0.0.0.0",
        port=int(os.environ.get("PORT", 5000))
    )


# ============================================================

