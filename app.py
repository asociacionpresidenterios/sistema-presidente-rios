import os
from datetime import date
from flask import Flask, render_template, request, redirect, url_for, flash, jsonify
from flask_sqlalchemy import SQLAlchemy

app = Flask(__name__)
app.config["SECRET_KEY"] = os.environ.get("SECRET_KEY", "cambia-esta-clave-en-produccion")
db_url = os.environ.get("DATABASE_URL")
if db_url:
    db_url = db_url.replace("postgres://", "postgresql+psycopg://", 1)
    db_url = db_url.replace("postgresql://", "postgresql+psycopg://", 1)
app.config["SQLALCHEMY_DATABASE_URI"] = db_url or "sqlite:///jugadores.db"
app.config["SQLALCHEMY_TRACK_MODIFICATIONS"] = False
db = SQLAlchemy(app)

class Jugador(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    rut = db.Column(db.String(20), unique=True, nullable=False, index=True)
    nombre_completo = db.Column(db.String(160), nullable=False)
    fecha_nacimiento = db.Column(db.Date, nullable=False)
    serie = db.Column(db.String(80), nullable=False)
    club = db.Column(db.String(120), nullable=False)

with app.app_context():
    db.create_all()

@app.route("/")
def index():
    q = request.args.get("q", "").strip()
    query = Jugador.query
    if q:
        query = query.filter(db.or_(Jugador.rut.ilike(f"%{q}%"),
                                    Jugador.nombre_completo.ilike(f"%{q}%"),
                                    Jugador.club.ilike(f"%{q}%")))
    jugadores = query.order_by(Jugador.nombre_completo).all()
    return render_template("index.html", jugadores=jugadores, q=q)

@app.route("/jugadores/nuevo", methods=["GET","POST"])
def nuevo_jugador():
    if request.method == "POST":
        rut = request.form["rut"].strip().upper()
        nombre = request.form["nombre_completo"].strip()
        fecha = request.form["fecha_nacimiento"].strip()
        serie = request.form["serie"].strip()
        club = request.form["club"].strip()
        if not all([rut,nombre,fecha,serie,club]):
            flash("Completa todos los campos.", "error")
            return render_template("jugador_form.html", jugador=None)
        try:
            fecha_obj = date.fromisoformat(fecha)
        except ValueError:
            flash("Fecha no válida.", "error")
            return render_template("jugador_form.html", jugador=None)
        if Jugador.query.filter_by(rut=rut).first():
            flash("Ese RUT ya está registrado.", "error")
            return render_template("jugador_form.html", jugador=None)
        db.session.add(Jugador(rut=rut,nombre_completo=nombre,
                               fecha_nacimiento=fecha_obj,serie=serie,club=club))
        db.session.commit()
        flash("Jugador registrado correctamente.", "success")
        return redirect(url_for("index"))
    return render_template("jugador_form.html", jugador=None)

@app.route("/jugadores/<int:jugador_id>/editar", methods=["GET","POST"])
def editar_jugador(jugador_id):
    jugador = db.get_or_404(Jugador, jugador_id)
    if request.method == "POST":
        jugador.rut = request.form["rut"].strip().upper()
        jugador.nombre_completo = request.form["nombre_completo"].strip()
        jugador.fecha_nacimiento = date.fromisoformat(request.form["fecha_nacimiento"])
        jugador.serie = request.form["serie"].strip()
        jugador.club = request.form["club"].strip()
        db.session.commit()
        flash("Datos actualizados.", "success")
        return redirect(url_for("index"))
    return render_template("jugador_form.html", jugador=jugador)

@app.route("/jugadores/<int:jugador_id>/eliminar", methods=["POST"])
def eliminar_jugador(jugador_id):
    jugador = db.get_or_404(Jugador, jugador_id)
    db.session.delete(jugador)
    db.session.commit()
    flash("Jugador eliminado.", "success")
    return redirect(url_for("index"))

@app.route("/api/jugadores")
def api_jugadores():
    rut = request.args.get("rut", "").strip().upper()
    return jsonify([{
        "id": j.id, "rut": j.rut, "nombre_completo": j.nombre_completo,
        "fecha_nacimiento": j.fecha_nacimiento.isoformat(),
        "serie": j.serie, "club": j.club
    } for j in Jugador.query.filter(Jugador.rut.ilike(f"%{rut}%")).all()]) if rut else jsonify([])

@app.route("/health")
def health():
    return {"status":"ok"}

if __name__ == "__main__":
    app.run(host="0.0.0.0", port=int(os.environ.get("PORT",5000)))
