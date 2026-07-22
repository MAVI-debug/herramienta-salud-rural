"""
app.py — Servidor Web TSR (Flask)
===================================
Panel de administración para desparasitación y fluorización.
"""

import os
import hashlib
import tempfile
import zipfile
from datetime import timedelta, date, datetime

from flask import (Flask, render_template, request, redirect,
                   url_for, session, flash, g, send_file, make_response)
from flask_caching import Cache
from werkzeug.security import generate_password_hash, check_password_hash

from db import get_db, fetchone, fetchall, execute, commit, rollback, close_db

import sisca_logic

# ---------------------------------------------------------------------------
# Configuración
# ---------------------------------------------------------------------------
app = Flask(__name__)
app.secret_key = os.urandom(32).hex()
app.permanent_session_lifetime = timedelta(hours=8)
app.config["MAX_CONTENT_LENGTH"] = 200 * 1024 * 1024  # 200 MB

# Cache en memoria para métricas pesadas
cache = Cache(app, config={"CACHE_TYPE": "SimpleCache", "CACHE_DEFAULT_TIMEOUT": 600})

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
PLANTILLA_PATH = os.path.join(BASE_DIR, "plantilla_sisca.xlsx")
PLANTILLA_SIGSA22_PATH = os.path.join(BASE_DIR, "SIGSA_22_de_odontologia.xlsx")
SALIDA_SISCA_DIR = os.path.join(BASE_DIR, "descargas_sisca")


# Teardown: cierra la conexión híbrida
app.teardown_appcontext(close_db)


def hash_contrasena(contrasena: str) -> str:
    """Genera hash con werkzeug (bcrypt-style); compatible con SHA256 legacy."""
    return generate_password_hash(contrasena)


def verificar_contrasena(almacenada: str, proporcionada: str) -> bool:
    """Verifica contra werkzeug; fallback a SHA256 para usuarios existentes."""
    if almacenada.startswith("pbkdf2:") or almacenada.startswith("scrypt:"):
        return check_password_hash(almacenada, proporcionada)
    return hashlib.sha256(proporcionada.encode("utf-8")).hexdigest() == almacenada


# ---------------------------------------------------------------------------
# Helpers de sesión
# ---------------------------------------------------------------------------
def login_requerido():
    if "usuario_id" not in session:
        flash("Inicia sesión para acceder.", "warning")
        return redirect(url_for("login"))
    return None


# ---------------------------------------------------------------------------
# Ruta temporal — regenerar usuario demo
# ---------------------------------------------------------------------------
@app.route("/generar-demo")
def generar_demo():
    contrasena_plano = "demo123"
    hash_nuevo = hash_contrasena(contrasena_plano)
    execute("DELETE FROM usuarios WHERE usuario = %s", ("tsr_demo",))
    execute("""
        INSERT INTO usuarios
            (usuario, contrasena_hash, nombre_responsable, cargo, area_salud, distrito_salud)
        VALUES (%s, %s, %s, %s, %s, %s)
    """, ("tsr_demo", hash_nuevo, "Andrés Romeo Mazariegos Vicente",
          "Técnico en Salud Rural (TSR)", "TOTONICAPÁN", "TOTONICAPÁN"))
    commit()
    print(f"DEBUG /generar-demo: hash insertado = {hash_nuevo}")
    return f"Usuario tsr_demo recreado. Hash: {hash_nuevo}", 200


# ---------------------------------------------------------------------------
# Rutas — Login
# ---------------------------------------------------------------------------
@app.route("/", methods=["GET"])
def index():
    if "usuario_id" in session:
        return redirect(url_for("dashboard"))
    return redirect(url_for("login"))


@app.route("/login", methods=["GET", "POST"])
def login():
    if request.method == "POST":
        usuario = request.form.get("usuario", "").strip()
        contrasena = request.form.get("contrasena", "")

        if not usuario or not contrasena:
            flash("Completa todos los campos.", "danger")
            return render_template("login.html")

        fila = fetchone(
            "SELECT id, usuario, nombre_responsable, cargo, area_salud, "
            "distrito_salud, contrasena_hash "
            "FROM usuarios WHERE usuario = %s",
            (usuario,)
        )
        print("DEBUG LOGIN fila:", fila)

        if fila and verificar_contrasena(fila["contrasena_hash"], contrasena):
            session.permanent = True
            session["usuario_id"] = fila["id"]
            session["usuario"] = fila["usuario"]
            session["nombre_responsable"] = fila["nombre_responsable"]
            session["cargo"] = fila["cargo"]
            session["area_salud"] = fila["area_salud"]
            session["distrito_salud"] = fila["distrito_salud"]
            flash(f"Bienvenido, {fila['nombre_responsable']}.", "success")
            return redirect(url_for("dashboard"))
        else:
            flash("Usuario o contraseña incorrectos.", "danger")

    return render_template("login.html")


@app.route("/logout")
def logout():
    session.clear()
    flash("Sesión cerrada.", "info")
    return redirect(url_for("login"))


# ---------------------------------------------------------------------------
# Rutas — Registro de TSR
# ---------------------------------------------------------------------------

@app.route("/registrar_tsr", methods=["GET", "POST"])
def registrar_tsr():
    if "usuario_id" not in session:
        return login_requerido()

    if request.method == "POST":
        nombre = request.form.get("nombre_completo", "").strip()
        usuario = request.form.get("usuario", "").strip()
        contrasena = request.form.get("contrasena", "")
        cargo = request.form.get("cargo", "Técnico en Salud Rural (TSR)").strip()
        area_salud = request.form.get("area_salud", "").strip()
        distrito_salud = request.form.get("distrito_salud", "").strip()

        if not all([nombre, usuario, contrasena, area_salud, distrito_salud]):
            flash("Completa todos los campos obligatorios.", "danger")
            return render_template("registrar.html")

        existe = fetchone(
            "SELECT id FROM usuarios WHERE usuario = %s", (usuario,)
        )
        if existe:
            flash(f"El usuario «{usuario}» ya existe.", "danger")
            return render_template("registrar.html")

        execute("""
            INSERT INTO usuarios
                (usuario, contrasena_hash, nombre_responsable, cargo, area_salud, distrito_salud)
            VALUES (%s, %s, %s, %s, %s, %s)
        """, (usuario, hash_contrasena(contrasena), nombre, cargo, area_salud, distrito_salud))
        commit()

        flash(f"¡TSR «{nombre}» registrado exitosamente!", "success")
        return redirect(url_for("registrar_tsr"))

    return render_template("registrar.html")


# ---------------------------------------------------------------------------
# Ruta única — Cargar datos legacy a usuarios existentes
# ---------------------------------------------------------------------------

@app.route("/cargar-datos-iniciales")
def cargar_datos_iniciales():
    """
    UNA SOLA VEZ: duplica las escuelas, estudiantes y jornadas existentes
    (con usuario_id = 1 o nulo) para TODOS los usuarios actuales.
    Los nuevos usuarios registrados después NO heredan estos datos.
    """
    usuarios = fetchall("SELECT id FROM usuarios")
    if not usuarios:
        return "No hay usuarios en el sistema.", 400

    resultados = []
    for u in usuarios:
        uid = u["id"]
        # Escuelas
        execute("""
            INSERT INTO escuelas (codigo_centro, usuario_id, nombre_centro, tipo_centro, servicio_salud)
            SELECT e.codigo_centro, %s, e.nombre_centro, e.tipo_centro, e.servicio_salud
            FROM escuelas e
            WHERE e.usuario_id IS NULL OR e.usuario_id = 1
            GROUP BY e.codigo_centro
            ON CONFLICT (codigo_centro, usuario_id) DO NOTHING
        """, (uid,))
        # Estudiantes
        execute("""
            INSERT INTO estudiantes (cui, usuario_id, nombre_completo, sexo, fecha_nacimiento, grado, seccion)
            SELECT e.cui, %s, e.nombre_completo, e.sexo, e.fecha_nacimiento, e.grado, e.seccion
            FROM estudiantes e
            WHERE e.usuario_id IS NULL OR e.usuario_id = 1
            GROUP BY e.cui
            ON CONFLICT (cui, usuario_id) DO NOTHING
        """, (uid,))
        # Registros de salud
        execute("""
            INSERT INTO registros_salud
                (cui_estudiante, codigo_centro, tipo_intervencion,
                 campana, fecha_aplicacion, fecha_corte, edad_calculo, usuario_id)
            SELECT r.cui_estudiante, r.codigo_centro, r.tipo_intervencion,
                   r.campana, r.fecha_aplicacion, r.fecha_corte, r.edad_calculo, %s
            FROM registros_salud r
            WHERE r.usuario_id IS NULL OR r.usuario_id = 1
        """, (uid,))
        resultados.append(f"Usuario {uid} procesado")

    commit()
    return "<br>".join(resultados), 200


# ---------------------------------------------------------------------------
# Rutas — Perfil de Usuario
# ---------------------------------------------------------------------------

@app.route("/perfil", methods=["GET", "POST"])
def perfil():
    if "usuario_id" not in session:
        return login_requerido()

    if request.method == "POST":
        form_type = request.form.get("form_type", "")

        if form_type == "actualizar_perfil":
            nombre = request.form.get("nombre_completo", "").strip()
            usuario = request.form.get("usuario", "").strip()
            cargo = request.form.get("cargo", "").strip()
            area_salud = request.form.get("area_salud", "").strip()
            distrito_salud = request.form.get("distrito_salud", "").strip()

            if not all([nombre, usuario]):
                flash("El nombre y el usuario son obligatorios.", "danger")
                return redirect(url_for("perfil"))

            duplicado = fetchone(
                "SELECT id FROM usuarios WHERE usuario = %s AND id != %s",
                (usuario, session["usuario_id"])
            )
            if duplicado:
                flash(f"El usuario «{usuario}» ya está en uso.", "danger")
                return redirect(url_for("perfil"))

            execute("""
                UPDATE usuarios
                SET nombre_responsable=%s, usuario=%s, cargo=%s,
                    area_salud=%s, distrito_salud=%s
                WHERE id=%s
            """, (nombre, usuario, cargo, area_salud, distrito_salud, session["usuario_id"]))
            commit()

            session["nombre_responsable"] = nombre
            session["usuario"] = usuario
            session["cargo"] = cargo
            session["area_salud"] = area_salud
            session["distrito_salud"] = distrito_salud

            flash("Perfil actualizado correctamente.", "success")
            return redirect(url_for("perfil"))

        elif form_type == "cambiar_contrasena":
            actual = request.form.get("contrasena_actual", "")
            nueva = request.form.get("nueva_contrasena", "")
            confirmar = request.form.get("confirmar_contrasena", "")

            if not all([actual, nueva, confirmar]):
                flash("Completa todos los campos de contraseña.", "danger")
                return redirect(url_for("perfil"))

            if nueva != confirmar:
                flash("Las contraseñas nuevas no coinciden.", "danger")
                return redirect(url_for("perfil"))

            if len(nueva) < 6:
                flash("La contraseña debe tener al menos 6 caracteres.", "danger")
                return redirect(url_for("perfil"))

            fila = fetchone(
                "SELECT contrasena_hash FROM usuarios WHERE id = %s",
                (session["usuario_id"],)
            )
            if not fila or not verificar_contrasena(fila["contrasena_hash"], actual):
                flash("La contraseña actual no es correcta.", "danger")
                return redirect(url_for("perfil"))

            execute(
                "UPDATE usuarios SET contrasena_hash = %s WHERE id = %s",
                (hash_contrasena(nueva), session["usuario_id"])
            )
            commit()
            flash("Contraseña cambiada correctamente.", "success")
            return redirect(url_for("perfil"))

    return render_template("perfil.html")


# ---------------------------------------------------------------------------
# Rutas — Dashboard
# ---------------------------------------------------------------------------
@app.route("/dashboard")
def dashboard():
    if "usuario_id" not in session:
        return login_requerido()

    uid = session["usuario_id"]

    total_escuelas = fetchone(
        "SELECT COUNT(*) AS c FROM escuelas WHERE usuario_id = %s", (uid,)
    )["c"]

    total_estudiantes = fetchone("""
        SELECT COUNT(*) AS c FROM estudiantes e
        WHERE e.usuario_id = %s AND EXISTS (
            SELECT 1 FROM registros_salud r
            JOIN escuelas esc ON r.codigo_centro = esc.codigo_centro
                AND r.usuario_id = esc.usuario_id
            WHERE r.cui_estudiante = e.cui
              AND r.usuario_id = e.usuario_id
        )
    """, (uid,))["c"]

    escuelas = fetchall(
        "SELECT codigo_centro, nombre_centro FROM escuelas WHERE usuario_id = %s ORDER BY nombre_centro",
        (uid,)
    )

    # Desglose por tipo de escuela
    tipo_counts = fetchall(
        "SELECT tipo_centro, COUNT(*) AS c FROM escuelas WHERE usuario_id = %s GROUP BY tipo_centro",
        (uid,)
    )
    desglose_tipos = {r["tipo_centro"]: r["c"] for r in tipo_counts}

    # Escuela con mayor número de alumnos (a través de registros_salud)
    mayor_poblacion = fetchone("""
        SELECT e.nombre_centro, COUNT(*) AS c
        FROM registros_salud r
        JOIN estudiantes s ON r.cui_estudiante = s.cui AND r.usuario_id = s.usuario_id
        JOIN escuelas e ON r.codigo_centro = e.codigo_centro AND r.usuario_id = e.usuario_id
        WHERE r.usuario_id = %s
        GROUP BY e.nombre_centro, e.codigo_centro, e.usuario_id
        ORDER BY c DESC
        LIMIT 1
    """, (uid,))
    nombre_mayor = mayor_poblacion["nombre_centro"] if mayor_poblacion else "—"

    # Helper: subconsulta EXISTS para estudiantes vinculados a escuelas activas
    student_exists_join = """
        AND EXISTS (
            SELECT 1 FROM registros_salud r
            JOIN escuelas esc ON r.codigo_centro = esc.codigo_centro
                AND r.usuario_id = esc.usuario_id
            WHERE r.cui_estudiante = e.cui
              AND r.usuario_id = e.usuario_id
        )
    """

    # Distribución por sexo (solo estudiantes con escuela activa)
    sexo_counts = fetchall(
        "SELECT e.sexo, COUNT(*) AS c FROM estudiantes e"
        " WHERE e.usuario_id = %s" + student_exists_join +
        " GROUP BY e.sexo",
        (uid,)
    )
    total_m = next((r["c"] for r in sexo_counts if r["sexo"] == "Masculino"), 0)
    total_f = next((r["c"] for r in sexo_counts if r["sexo"] == "Femenino"), 0)

    # Fecha de corte para cálculo de edad
    fecha_corte = _obtener_fecha_corte()
    anio_c = fecha_corte.year
    mes_c = fecha_corte.month
    dia_c = fecha_corte.day

    # Alumnos en rango 6-14 años (solo con escuela activa)
    en_rango = fetchone("""
        SELECT COUNT(*) AS c FROM estudiantes e
        WHERE e.usuario_id = %s
    """ + student_exists_join + """
          AND (CAST(%s AS INTEGER) - CAST(substr(e.fecha_nacimiento,7,4) AS INTEGER)
               - CASE
                   WHEN CAST(substr(e.fecha_nacimiento,4,2) AS INTEGER) > %s
                        OR (CAST(substr(e.fecha_nacimiento,4,2) AS INTEGER) = %s
                            AND CAST(substr(e.fecha_nacimiento,1,2) AS INTEGER) > %s)
                   THEN 1 ELSE 0
                 END
              ) BETWEEN 6 AND 14
    """, (uid, anio_c, mes_c, mes_c, dia_c))["c"]
    pct_rango = round(en_rango / total_estudiantes * 100, 1) if total_estudiantes else 0

    # Desparasitados por jornada (vinculados a escuelas activas)
    desp_1ra = fetchone("""
        SELECT COUNT(DISTINCT r.cui_estudiante) AS c
        FROM registros_salud r
        JOIN escuelas e ON r.codigo_centro = e.codigo_centro AND r.usuario_id = e.usuario_id
        WHERE r.usuario_id = %s
          AND r.tipo_intervencion = 'Desparasitación'
          AND r.campana = 'Primera'
    """, (uid,))["c"]

    desp_2da = fetchone("""
        SELECT COUNT(DISTINCT r.cui_estudiante) AS c
        FROM registros_salud r
        JOIN escuelas e ON r.codigo_centro = e.codigo_centro AND r.usuario_id = e.usuario_id
        WHERE r.usuario_id = %s
          AND r.tipo_intervencion = 'Desparasitación'
          AND r.campana = 'Segunda'
    """, (uid,))["c"]

    fecha_corte_str = fecha_corte.strftime("%d/%m/%Y")

    return render_template(
        "dashboard.html",
        total_escuelas=total_escuelas,
        total_estudiantes=total_estudiantes,
        escuelas=escuelas,
        desglose_tipos=desglose_tipos,
        nombre_mayor=nombre_mayor,
        total_m=total_m,
        total_f=total_f,
        en_rango=en_rango,
        pct_rango=pct_rango,
        fecha_corte_str=fecha_corte_str,
        desp_1ra=desp_1ra,
        desp_2da=desp_2da,
    )


# ---------------------------------------------------------------------------
# Rutas — Jornadas (páginas completas)
# ---------------------------------------------------------------------------

@app.route("/desparasitacion")
def jornada_desparasitacion():
    if "usuario_id" not in session:
        return login_requerido()
    uid = session["usuario_id"]
    escuelas = fetchall(
        "SELECT codigo_centro, nombre_centro, tipo_centro FROM escuelas WHERE usuario_id = %s ORDER BY nombre_centro",
        (uid,)
    )
    return render_template("jornada_desparasitacion.html", escuelas=escuelas)


@app.route("/fluorizacion")
def jornada_fluorizacion():
    if "usuario_id" not in session:
        return login_requerido()
    uid = session["usuario_id"]
    escuelas = fetchall(
        "SELECT codigo_centro, nombre_centro, tipo_centro FROM escuelas WHERE usuario_id = %s ORDER BY nombre_centro",
        (uid,)
    )
    return render_template("jornada_fluorizacion.html", escuelas=escuelas)


# ---------------------------------------------------------------------------
# Rutas — Escuelas
# ---------------------------------------------------------------------------
@app.route("/escuelas")
def listar_escuelas():
    if "usuario_id" not in session:
        return login_requerido()

    uid = session["usuario_id"]
    escuelas = fetchall(
        "SELECT codigo_centro, nombre_centro, tipo_centro, servicio_salud "
        "FROM escuelas WHERE usuario_id = %s ORDER BY nombre_centro",
        (uid,)
    )

    return render_template("escuelas.html", escuelas=escuelas)


# ---------------------------------------------------------------------------
# Rutas — Formulario y procesamiento SISCA
# ---------------------------------------------------------------------------

@app.route("/sisca", methods=["GET"])
def sisca_form():
    if "usuario_id" not in session:
        return login_requerido()

    codigo_centro = request.args.get("codigo_centro", "").strip()
    tipo = request.args.get("tipo", "desparasitacion")
    uid = session["usuario_id"]

    # ── Vista masiva: no se especificó escuela ──────────────────────────────
    if not codigo_centro:
        escuelas = fetchall(
            "SELECT codigo_centro, nombre_centro, tipo_centro, servicio_salud "
            "FROM escuelas WHERE usuario_id = %s ORDER BY nombre_centro",
            (uid,)
        )
        if not escuelas:
            flash("No hay escuelas registradas. Sube un PDF desde el Consolidado.", "warning")
            return redirect(url_for("dashboard"))

        today = date.today()
        fc = date.today()
        return render_template("sisca_masiva.html",
                               tipo=tipo,
                               escuelas=[dict(e) for e in escuelas],
                               now=f"{today.day:02d}/{today.month:02d}/{today.year}",
                               fecha_corte_iso=fc.isoformat())

    # ── Vista individual: una escuela específica ────────────────────────────
    escuela = fetchone(
        "SELECT codigo_centro, nombre_centro, tipo_centro, servicio_salud "
        "FROM escuelas WHERE codigo_centro = %s AND usuario_id = %s",
        (codigo_centro, uid)
    )
    if not escuela:
        flash("Escuela no encontrada.", "danger")
        return redirect(url_for("dashboard"))

    hoy = date.today()
    fc = date.today()
    return render_template("sisca_form.html",
                           tipo=tipo,
                           escuela=dict(escuela),
                           now=f"{hoy.day:02d}/{hoy.month:02d}/{hoy.year}",
                           fecha_corte_iso=fc.isoformat())


# ---------------------------------------------------------------------------
# Rutas — Formulario SIGSA-22 (Fluorización)
# ---------------------------------------------------------------------------

@app.route("/sisca_fluor", methods=["GET"])
def sisca_fluor_form():
    if "usuario_id" not in session:
        return login_requerido()

    codigo_centro = request.args.get("codigo_centro", "").strip()
    escuela = None
    uid = session["usuario_id"]
    if codigo_centro:
        escuela = fetchone(
            "SELECT codigo_centro, nombre_centro, servicio_salud "
            "FROM escuelas WHERE codigo_centro = %s AND usuario_id = %s",
            (codigo_centro, uid)
        )

    hoy = date.today()
    fc = date.today()
    return render_template("sisca_fluor.html",
                           escuela=dict(escuela) if escuela else None,
                           fecha_corte_iso=fc.isoformat())


@app.route("/generar_sigsa22", methods=["POST"])
def generar_sigsa22():
    if "usuario_id" not in session:
        return login_requerido()

    codigo_centro = request.form.get("codigo_centro", "").strip()
    area_salud = request.form.get("area_salud", "").strip()
    distrito_salud = request.form.get("distrito_salud", "").strip()
    municipio = request.form.get("municipio", "").strip()
    servicio_salud = request.form.get("servicio_salud", "").strip()
    responsable_informacion = request.form.get("responsable_informacion", "").strip()
    cargo = request.form.get("cargo", "").strip()
    mes_reporte = request.form.get("mes_reporte", "").strip()
    anio_reporte = request.form.get("anio_reporte", "").strip()
    fecha_corte_str = request.form.get("fecha_corte", "").strip()

    if not all([codigo_centro, area_salud, distrito_salud, municipio,
                servicio_salud, responsable_informacion, cargo,
                mes_reporte, anio_reporte]):
        flash("Completa todos los campos obligatorios.", "danger")
        return redirect(url_for("sisca_fluor_form", codigo_centro=codigo_centro))

    if not os.path.exists(PLANTILLA_SIGSA22_PATH):
        flash("No se encontró la plantilla SIGSA-22 en el servidor.", "danger")
        return redirect(url_for("dashboard"))

    # ── Parsear fecha de corte ──────────────────────────────────────────────
    if fecha_corte_str:
        try:
            fc = datetime.strptime(fecha_corte_str.split("T")[0], "%Y-%m-%d").date()
        except Exception:
            try:
                fc = datetime.strptime(fecha_corte_str, "%d/%m/%Y").date()
            except Exception:
                fc = date.today()
    else:
        fc = date.today()

    uid = session["usuario_id"]
    escuela = fetchone(
        "SELECT codigo_centro, nombre_centro FROM escuelas WHERE codigo_centro = %s AND usuario_id = %s",
        (codigo_centro, uid)
    )
    if not escuela:
        flash("Escuela no encontrada.", "danger")
        return redirect(url_for("dashboard"))
    nombre_escuela = escuela["nombre_centro"]

    # ── Consultar estudiantes ───────────────────────────────────────────────
    registros = fetchall("""
        SELECT e.cui, e.nombre_completo, e.sexo, e.fecha_nacimiento,
               e.grado, e.seccion
        FROM registros_salud r
        JOIN estudiantes e ON r.cui_estudiante = e.cui AND r.usuario_id = e.usuario_id
        WHERE r.codigo_centro = %s AND r.usuario_id = %s
        GROUP BY e.cui, e.usuario_id
    """, (codigo_centro, uid))

    if not registros:
        flash("No hay estudiantes registrados.", "warning")
        return redirect(url_for("dashboard"))

    # ── Filtrar 6-14 años ───────────────────────────────────────────────────
    aptos = []
    for r in registros:
        try:
            dia, mes, anio = str(r["fecha_nacimiento"]).split("/")
            dia_i, mes_i, anio_i = int(dia), int(mes), int(anio)
        except Exception:
            continue
        edad = sisca_logic.calcular_edad_a_fecha_corte(dia_i, mes_i, anio_i, fc)
        gen = "M" if r["sexo"].startswith("M") else "F"
        if edad is not None and 6 <= edad <= 14:
            aptos.append({
                "nombre": r["nombre_completo"],
                "cui": r["cui"],
                "genero": gen,
                "dia": f"{dia_i:02d}",
                "mes": f"{mes_i:02d}",
                "anio": str(anio_i),
            })

    if not aptos:
        flash("No hay alumnos en edad apta (6-14 años).", "warning")
        return redirect(url_for("dashboard"))

    # ── Generar SIGSA-22 ────────────────────────────────────────────────────
    os.makedirs(SALIDA_SISCA_DIR, exist_ok=True)
    import re as re_mod
    nombre_archivo = re_mod.sub(r'[\\/*?:"<>|]', "", codigo_centro)[:60]
    ruta_salida = os.path.join(SALIDA_SISCA_DIR, f"SIGSA22_{nombre_archivo}.xlsx")

    estudiantes = []
    for a in aptos:
        g = "F" if a.get("genero", "").upper().startswith("F") else "M"
        estudiantes.append({
            "cui": a.get("cui", ""),
            "nombre_completo": a.get("nombre", ""),
            "sexo": g,
            "pueblo": "",
            "comunidad_linguistica": "",
            "dia_nac": a.get("dia", ""),
            "mes_nac": a.get("mes", ""),
            "anio_nac": a.get("anio", ""),
        })

    datos = {
        "area_salud": area_salud,
        "distrito_salud": distrito_salud,
        "municipio": municipio,
        "servicio_salud": servicio_salud,
        "responsable_informacion": responsable_informacion,
        "cargo": cargo,
        "mes_reporte": mes_reporte,
        "anio_reporte": anio_reporte,
        "estudiantes": estudiantes,
    }

    try:
        sisca_logic.generar_reporte_sigsa22(
            PLANTILLA_SIGSA22_PATH, ruta_salida, datos,
        )
    except Exception as e:
        flash(f"Error al generar SIGSA-22: {e}", "danger")
        return redirect(url_for("dashboard"))

    if not os.path.exists(ruta_salida) or os.path.getsize(ruta_salida) == 0:
        flash("Error: el archivo SIGSA-22 no se guardó correctamente en el servidor.", "danger")
        return redirect(url_for("dashboard"))

    response = make_response(send_file(
        ruta_salida,
        as_attachment=True,
        download_name=f"SIGSA22_{codigo_centro}.xlsx",
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    ))
    response.headers["Content-Length"] = os.path.getsize(ruta_salida)
    response.headers["Cache-Control"] = "no-cache, no-store, must-revalidate"
    return response


def _generar_sisca_para_escuela(codigo_centro, uid, fecha_corte, responsable,
                                 cargo, area_salud, distrito_salud,
                                 servicio_salud, tipo_centro, jornada,
                                 fecha_reporte=None):
    """
    Genera una ficha SISCA para una escuela y retorna (ruta_salida, nombre_escuela, error_msg).
    error_msg es None si todo salio bien.
    """
    escuela = fetchone(
        "SELECT codigo_centro, nombre_centro, tipo_centro FROM escuelas WHERE codigo_centro = %s AND usuario_id = %s",
        (codigo_centro, uid)
    )
    if not escuela:
        return None, None, f"Escuela no encontrada: {codigo_centro}"

    nombre_escuela = escuela["nombre_centro"]
    tipo_centro = escuela.get("tipo_centro", "PUBLICO") or "PUBLICO"
    registros = fetchall("""
        SELECT e.cui, e.nombre_completo, e.sexo, e.fecha_nacimiento,
               e.grado, e.seccion
        FROM registros_salud r
        JOIN estudiantes e ON r.cui_estudiante = e.cui AND r.usuario_id = e.usuario_id
        WHERE r.codigo_centro = %s AND r.usuario_id = %s
        GROUP BY e.cui, e.usuario_id
        ORDER BY MIN(r.id)
    """, (codigo_centro, uid))

    if not registros:
        return None, None, f"Sin estudiantes registrados para '{nombre_escuela}'."

    aptos = []
    for r in registros:
        try:
            dia, mes, anio = str(r["fecha_nacimiento"]).split("/")
            dia_i, mes_i, anio_i = int(dia), int(mes), int(anio)
        except Exception:
            continue
        edad = sisca_logic.calcular_edad_a_fecha_corte(dia_i, mes_i, anio_i, fecha_corte)
        gen = "F" if r["sexo"].startswith("F") else "M"
        nombre = r["nombre_completo"]
        cui = r["cui"]
        if edad is not None and 6 <= edad <= 14:
            aptos.append({
                "nombre": nombre,
                "cui": cui,
                "genero": gen,
                "dia": dia_i,
                "mes": mes_i,
                "anio": anio_i,
            })

    if not aptos:
        return None, None, f"No hay alumnos en edad apta (6-14 anos) para '{nombre_escuela}'."

    anio_campana = fecha_corte.year
    os.makedirs(SALIDA_SISCA_DIR, exist_ok=True)
    import re
    nombre_archivo = re.sub(r'[\\/*?:"<>|]', "", codigo_centro or nombre_escuela)[:60]
    ruta_salida = os.path.join(SALIDA_SISCA_DIR, f"SISCA_{nombre_archivo}.xlsx")

    try:
        sisca_logic.generar_ficha_sisca_escuela(
            PLANTILLA_PATH, ruta_salida,
            nombre_escuela, codigo_centro, aptos, anio_campana,
            tipo_centro=tipo_centro,
            responsable=responsable, cargo=cargo,
            area=area_salud, distrito=distrito_salud,
            servicio=servicio_salud,
            fecha_reporte_str=fecha_reporte if fecha_reporte else None,
            jornada=jornada,
        )
    except Exception as e:
        return None, None, f"Error al generar ficha para '{nombre_escuela}': {e}"

    return ruta_salida, nombre_escuela, None


# ---------------------------------------------------------------------------
# Rutas — SISCA Individual
# ---------------------------------------------------------------------------

@app.route("/procesar_sisca", methods=["POST"])
def procesar_sisca():
    if "usuario_id" not in session:
        return login_requerido()

    # ── Leer campos del formulario ────────────────────────────────────────
    codigo_centro = request.form.get("codigo_centro", "").strip()
    responsable = request.form.get("responsable", "").strip()
    cargo = request.form.get("cargo", "").strip()
    area_salud = request.form.get("area_salud", "").strip()
    distrito_salud = request.form.get("distrito_salud", "").strip()
    servicio_salud = request.form.get("servicio_salud", "").strip()
    tipo_centro = request.form.get("tipo_centro", "PUBLICO").strip().upper()
    fecha_reporte = request.form.get("fecha_reporte", "").strip()
    fecha_corte_str = request.form.get("fecha_corte", "").strip()
    jornada = request.form.get("jornada", "Primera Jornada").strip()
    tipo_intervencion = request.form.get("tipo_intervencion", "desparasitacion").strip()

    if not codigo_centro:
        flash("No se selecciono ninguna escuela.", "danger")
        return redirect(url_for("dashboard"))

    if not all([responsable, cargo, area_salud, distrito_salud, servicio_salud]):
        flash("Completa todos los campos obligatorios del encabezado.", "danger")
        return redirect(url_for("sisca_form", tipo=tipo_intervencion, codigo_centro=codigo_centro))

    if not os.path.exists(PLANTILLA_PATH):
        flash("No se encontro la plantilla SISCA en el servidor.", "danger")
        return redirect(url_for("dashboard"))

    # ── Parsear fecha de corte ────────────────────────────────────────────
    if fecha_corte_str:
        try:
            partes = fecha_corte_str.split("-")
            fecha_corte = date(int(partes[0]), int(partes[1]), int(partes[2]))
        except Exception:
            flash("Fecha de corte invalida.", "danger")
            return redirect(url_for("dashboard"))
    else:
        fecha_corte = date.today()

    # ── Actualizar perfil del usuario ─────────────────────────────────────
    execute("""
        UPDATE usuarios
        SET nombre_responsable=%s, cargo=%s, area_salud=%s, distrito_salud=%s
        WHERE id=%s
    """, (responsable, cargo, area_salud, distrito_salud, session["usuario_id"]))
    commit()
    session.update(nombre_responsable=responsable, cargo=cargo,
                   area_salud=area_salud, distrito_salud=distrito_salud)

    uid = session["usuario_id"]
    ruta_salida, nombre_escuela, error = _generar_sisca_para_escuela(
        codigo_centro, uid, fecha_corte, responsable, cargo,
        area_salud, distrito_salud, servicio_salud, "PUBLICO", jornada,
        fecha_reporte=fecha_reporte if fecha_reporte else None,
    )

    if error:
        flash(error, "danger" if "Error" in error else "warning")
        return redirect(url_for("dashboard"))

    return send_file(
        ruta_salida,
        as_attachment=True,
        download_name=f"SISCA_{codigo_centro}_{tipo_intervencion}.xlsx",
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )


# ---------------------------------------------------------------------------
# Rutas — SISCA Masivo (ExportaciÃ³n MÃºltiple a ZIP)
# ---------------------------------------------------------------------------

@app.route("/procesar_sisca_masiva", methods=["POST"])
def procesar_sisca_masiva():
    if "usuario_id" not in session:
        return login_requerido()

    codigos = request.form.getlist("codigos_centro")
    if not codigos:
        flash("No se selecciono ninguna escuela.", "danger")
        return redirect(url_for("sisca_form"))

    responsable = request.form.get("responsable", "").strip()
    cargo = request.form.get("cargo", "").strip()
    area_salud = request.form.get("area_salud", "").strip()
    distrito_salud = request.form.get("distrito_salud", "").strip()
    servicio_salud = request.form.get("servicio_salud", "").strip()
    tipo_centro = request.form.get("tipo_centro", "PUBLICO").strip().upper()
    fecha_reporte = request.form.get("fecha_reporte", "").strip()
    fecha_corte_str = request.form.get("fecha_corte", "").strip()
    jornada = request.form.get("jornada", "Primera Jornada").strip()
    tipo_intervencion = request.form.get("tipo_intervencion", "desparasitacion").strip()

    if not all([responsable, cargo, area_salud, distrito_salud, servicio_salud]):
        flash("Completa todos los campos obligatorios del encabezado.", "danger")
        return redirect(url_for("sisca_form"))

    if not os.path.exists(PLANTILLA_PATH):
        flash("No se encontro la plantilla SISCA en el servidor.", "danger")
        return redirect(url_for("dashboard"))

    if fecha_corte_str:
        try:
            partes = fecha_corte_str.split("-")
            fecha_corte = date(int(partes[0]), int(partes[1]), int(partes[2]))
        except Exception:
            flash("Fecha de corte invalida.", "danger")
            return redirect(url_for("sisca_form"))
    else:
        fecha_corte = date.today()

    # ── Actualizar perfil del usuario ─────────────────────────────────────
    uid = session["usuario_id"]
    execute("""
        UPDATE usuarios
        SET nombre_responsable=%s, cargo=%s, area_salud=%s, distrito_salud=%s
        WHERE id=%s
    """, (responsable, cargo, area_salud, distrito_salud, uid))
    commit()
    session.update(nombre_responsable=responsable, cargo=cargo,
                   area_salud=area_salud, distrito_salud=distrito_salud)

    # ── Generar ficha por cada escuela seleccionada ────────────────────────
    rutas_generadas = []
    errores = []
    for codigo in codigos:
        ruta, nombre, err = _generar_sisca_para_escuela(
            codigo, uid, fecha_corte, responsable, cargo,
            area_salud, distrito_salud, servicio_salud, "PUBLICO", jornada,
            fecha_reporte=fecha_reporte if fecha_reporte else None,
        )
        if err:
            errores.append(err)
        else:
            rutas_generadas.append(ruta)

    if not rutas_generadas:
        for e in errores:
            flash(e, "danger")
        return redirect(url_for("sisca_form"))

    # ── Empaquetar en ZIP ─────────────────────────────────────────────────
    import io, zipfile
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as zf:
        for ruta in rutas_generadas:
            nombre_zip = os.path.basename(ruta)
            zf.write(ruta, arcname=nombre_zip)

    for ruta in rutas_generadas:
        try:
            os.remove(ruta)
        except Exception:
            pass

    buf.seek(0)

    for err in errores:
        flash(err, "warning")

    return send_file(
        buf,
        as_attachment=True,
        download_name=f"SISCA_masivo_{tipo_intervencion}.zip",
        mimetype="application/zip",
    )


# ---------------------------------------------------------------------------
# Helpers — Consolidado
# ---------------------------------------------------------------------------

def _obtener_fecha_corte(uid=None):
    if uid is None:
        uid = session.get("usuario_id", 0)
    fila = fetchone("SELECT fecha_corte FROM registros_salud WHERE usuario_id = %s LIMIT 1", (uid,))
    if fila:
        try:
            partes = fila["fecha_corte"].split("/")
            return date(int(partes[2]), int(partes[1]), int(partes[0]))
        except Exception:
            pass
    return date.today()


def _consolidado_data(fecha_corte, uid=None):
    """Retorna (matriz, escuelas_detalle, totales) para vista y exportación."""
    if uid is None:
        uid = session.get("usuario_id", 0)
    escuelas = fetchall(
        "SELECT codigo_centro, nombre_centro FROM escuelas WHERE usuario_id = %s ORDER BY nombre_centro",
        (uid,)
    )

    # Única consulta: evita N+1 por escuela
    todos_registros = fetchall("""
        SELECT r.codigo_centro,
               e.cui, e.nombre_completo, e.sexo, e.fecha_nacimiento,
               e.grado, e.seccion
        FROM registros_salud r
        JOIN estudiantes e ON r.cui_estudiante = e.cui AND r.usuario_id = e.usuario_id
        WHERE r.usuario_id = %s
        GROUP BY e.cui, e.usuario_id, r.codigo_centro
        ORDER BY r.codigo_centro
    """, (uid,))

    # Agrupar en Python por código de centro
    alumnos_por_centro = {}
    for row in todos_registros:
        cc = row["codigo_centro"]
        if cc not in alumnos_por_centro:
            alumnos_por_centro[cc] = []
        edad = sisca_logic._edad_desde_fecha_nac(row["fecha_nacimiento"], fecha_corte)
        rango = sisca_logic._clasificar_rango(edad)
        genero = "F" if row["sexo"].startswith("F") else "M"
        alumnos_por_centro[cc].append({
            "nombre": row["nombre_completo"],
            "cui": row["cui"],
            "genero": genero,
            "fecha_nac": row["fecha_nacimiento"],
            "edad": edad,
            "grado": row["grado"] or "",
            "seccion": row["seccion"] or "",
            "rango": rango,
        })

    matriz = []
    escuelas_detalle = {}
    rangos = ["5_y_menos", "6_a_9", "10_a_14", "15_a_19"]
    totales = {f"{r}_{g}": 0 for r in rangos for g in ("f", "m")}
    totales["tot_f"] = 0
    totales["tot_m"] = 0
    totales["f_6_14"] = 0
    totales["m_6_14"] = 0
    totales["sub_6_14"] = 0
    totales["gral"] = 0

    for e in escuelas:
        codigo = e["codigo_centro"]
        alumnos = alumnos_por_centro.get(codigo, [])

        conteo = {}
        for a in alumnos:
            key = f"{a['rango']}_{a['genero'].lower()}"
            conteo[key] = conteo.get(key, 0) + 1

        escuelas_detalle[codigo] = alumnos
        fila = {"nombre": e["nombre_centro"], "codigo": codigo}
        for r in rangos:
            fila[f"{r}_f"] = conteo.get(f"{r}_f", 0)
            fila[f"{r}_m"] = conteo.get(f"{r}_m", 0)
        matriz.append(fila)

        for r in rangos:
            for g in ("f", "m"):
                totales[f"{r}_{g}"] += conteo.get(f"{r}_{g}", 0)
        totales["tot_f"] += sum(conteo.get(f"{r}_f", 0) for r in rangos)
        totales["tot_m"] += sum(conteo.get(f"{r}_m", 0) for r in rangos)
    totales["f_6_14"] = totales["6_a_9_f"] + totales["10_a_14_f"]
    totales["m_6_14"] = totales["6_a_9_m"] + totales["10_a_14_m"]
    totales["sub_6_14"] = totales["6_a_9_f"] + totales["6_a_9_m"] + totales["10_a_14_f"] + totales["10_a_14_m"]
    totales["gral"] = totales["tot_f"] + totales["tot_m"]

    return matriz, escuelas_detalle, totales


# ---------------------------------------------------------------------------
# Rutas — Consolidado e Historial
# ---------------------------------------------------------------------------

@app.route("/consolidado")
def consolidado():
    if "usuario_id" not in session:
        return login_requerido()

    # Fecha de corte: query param > session > today
    q_fc = request.args.get("fecha_corte", "").strip()
    if q_fc:
        try:
            fecha_corte = datetime.strptime(q_fc, "%d/%m/%Y").date()
            session["fecha_corte"] = fecha_corte.strftime("%d/%m/%Y")
        except (ValueError, TypeError):
            fecha_corte = _obtener_fecha_corte()
    elif "fecha_corte" in session:
        try:
            fecha_corte = datetime.strptime(session["fecha_corte"], "%d/%m/%Y").date()
        except (ValueError, TypeError):
            fecha_corte = date.today()
    else:
        fecha_corte = date.today()

    matriz, escuelas_detalle, totales = _consolidado_data(fecha_corte)

    for m in matriz:
        raw = escuelas_detalle.get(m["codigo"], [])
        grupos = {}
        for a in raw:
            key = (a.get("grado", ""), a.get("seccion", ""))
            grupos.setdefault(key, []).append(a)
        for k in grupos:
            grupos[k].sort(key=lambda x: x.get("nombre", ""))
        m["alumnos"] = raw
        m["alumnos_agrupados"] = [
            {"grado": k[0], "seccion": k[1], "alumnos": v}
            for k, v in sorted(grupos.items(),
                               key=lambda item: (
                                   sisca_logic._orden_grado(item[0][0]),
                                   item[0][1]
                               ))
        ]

    return render_template("consolidado.html",
                           matriz=matriz,
                           totales=totales,
                           fecha_corte=fecha_corte,
                           fecha_corte_iso=fecha_corte.isoformat(),
                           fecha_corte_str=fecha_corte.strftime("%d/%m/%Y"))


@app.route("/exportar_escuela/<codigo_centro>")
def exportar_escuela(codigo_centro):
    if "usuario_id" not in session:
        return login_requerido()
    q_fc = request.args.get("fecha_corte", "").strip()
    if q_fc:
        try:
            fecha_corte = datetime.strptime(q_fc, "%d/%m/%Y").date()
        except (ValueError, TypeError):
            fecha_corte = _obtener_fecha_corte()
    elif "fecha_corte" in session:
        try:
            fecha_corte = datetime.strptime(session["fecha_corte"], "%d/%m/%Y").date()
        except (ValueError, TypeError):
            fecha_corte = date.today()
    else:
        fecha_corte = date.today()

    uid = session["usuario_id"]
    escuela = fetchone(
        "SELECT codigo_centro, nombre_centro FROM escuelas WHERE codigo_centro = %s AND usuario_id = %s",
        (codigo_centro, uid)
    )
    if not escuela:
        flash("Escuela no encontrada.", "danger")
        return redirect(url_for("consolidado"))

    registros = fetchall("""
        SELECT e.cui, e.nombre_completo, e.sexo, e.fecha_nacimiento,
               e.grado, e.seccion
        FROM registros_salud r
        JOIN estudiantes e ON r.cui_estudiante = e.cui AND r.usuario_id = e.usuario_id
        WHERE r.codigo_centro = %s AND r.usuario_id = %s
        GROUP BY e.cui, e.usuario_id
    """, (codigo_centro, uid))

    alumnos = []
    for r in registros:
        alumnos.append({
            "nombre": r["nombre_completo"],
            "cui": r["cui"],
            "genero": "F" if r["sexo"].startswith("F") else "M",
            "fecha_nac": r["fecha_nacimiento"],
            "grado": r["grado"] or "",
            "seccion": r["seccion"] or "",
        })

    os.makedirs(SALIDA_SISCA_DIR, exist_ok=True)
    ruta = os.path.join(SALIDA_SISCA_DIR, f"escuela_{codigo_centro}.xlsx")
    sisca_logic.generar_excel_escuela(
        ruta, escuela["nombre_centro"], codigo_centro, alumnos, fecha_corte
    )

    return send_file(
        ruta, as_attachment=True,
        download_name=f"{escuela['nombre_centro'][:40]}.xlsx",
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )


@app.route("/exportar_consolidado_total")
def exportar_consolidado_total():
    if "usuario_id" not in session:
        return login_requerido()
    q_fc = request.args.get("fecha_corte", "").strip()
    if q_fc:
        try:
            fecha_corte = datetime.strptime(q_fc, "%d/%m/%Y").date()
        except (ValueError, TypeError):
            fecha_corte = _obtener_fecha_corte()
    elif "fecha_corte" in session:
        try:
            fecha_corte = datetime.strptime(session["fecha_corte"], "%d/%m/%Y").date()
        except (ValueError, TypeError):
            fecha_corte = date.today()
    else:
        fecha_corte = date.today()
    matriz, escuelas_detalle, _ = _consolidado_data(fecha_corte)

    os.makedirs(SALIDA_SISCA_DIR, exist_ok=True)
    ruta = os.path.join(SALIDA_SISCA_DIR, "consolidado_total.xlsx")
    sisca_logic.generar_excel_consolidado(
        ruta, matriz, escuelas_detalle, fecha_corte
    )

    return send_file(
        ruta, as_attachment=True,
        download_name="consolidado_TSR.xlsx",
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )


# ---------------------------------------------------------------------------
# Rutas — Eliminar escuela
# ---------------------------------------------------------------------------

@app.route("/eliminar_escuela/<codigo_centro>", methods=["POST"])
def eliminar_escuela(codigo_centro):
    if "usuario_id" not in session:
        return login_requerido()
    uid = session["usuario_id"]
    execute("DELETE FROM escuelas WHERE codigo_centro = %s AND usuario_id = %s",
            (codigo_centro, uid))
    execute("DELETE FROM estudiantes WHERE usuario_id = %s AND cui IN ("
            "SELECT cui_estudiante FROM registros_salud WHERE codigo_centro = %s AND usuario_id = %s"
            ")", (uid, codigo_centro, uid))
    execute("DELETE FROM registros_salud WHERE codigo_centro = %s AND usuario_id = %s",
            (codigo_centro, uid))
    commit()
    flash(f"Escuela {codigo_centro} eliminada.", "info")
    return redirect(url_for("consolidado"))


# ---------------------------------------------------------------------------
# Rutas — Cargar PDF desde Consolidado (alimentación directa)
# ---------------------------------------------------------------------------

@app.route("/cargar_pdf_consolidado", methods=["POST"])
def cargar_pdf_consolidado():
    if "usuario_id" not in session:
        return login_requerido()

    archivos = request.files.getlist("archivos_pdf")
    if not archivos or all(a.filename == "" for a in archivos):
        flash("No se seleccionó ningún archivo PDF.", "danger")
        return redirect(url_for("consolidado"))

    fecha_corte_str = request.form.get("fecha_corte", "").strip()
    if fecha_corte_str:
        try:
            fc = datetime.strptime(fecha_corte_str, "%d/%m/%Y").date()
            fecha_corte_db = fc.strftime("%d/%m/%Y")
        except (ValueError, TypeError):
            fc = date.today()
            fecha_corte_db = fc.strftime("%d/%m/%Y")
    else:
        fc = date.today()
        fecha_corte_db = fc.strftime("%d/%m/%Y")

    total_escuelas = 0
    errores = []
    tipo_intervencion = "desparasitacion"

    for archivo in archivos:
        if archivo.filename == "":
            continue
        if not archivo.filename.lower().endswith(".pdf"):
            errores.append(f"«{archivo.filename}» no es PDF, se omitió.")
            continue

        tmp_path = None
        try:
            tmp = tempfile.NamedTemporaryFile(delete=False, suffix=".pdf")
            tmp_path = tmp.name
            archivo.save(tmp_path)
            tmp.close()

            alumnos = sisca_logic.extraer_alumnos_pdf(tmp_path)
            if not alumnos:
                errores.append(f"«{archivo.filename}»: sin alumnos, se omitio.")
                continue

            nombre, direccion, codigo = sisca_logic.extraer_metadatos_encabezado_pdf(tmp_path)
            nombre_escuela = sisca_logic.construir_nombre_escolar_completo(nombre, direccion)
            if not nombre_escuela:
                nombre_escuela = "ESCUELA SIN NOMBRE"
            if not codigo:
                codigo = f"SIN_CODIGO_{total_escuelas + 1}"

            uid = session["usuario_id"]
            execute("""
                INSERT INTO escuelas
                    (codigo_centro, usuario_id, nombre_centro, tipo_centro, servicio_salud)
                VALUES (%s, %s, %s, %s, %s)
                ON CONFLICT (codigo_centro, usuario_id) DO NOTHING
            """, (codigo, uid, nombre_escuela, "PUBLICO", ""))

            for a in alumnos:
                dia, mes, anio = sisca_logic._split_fecha(a["fecha_nac"])
                gen = "F" if a.get("genero", "").startswith("F") else "M"
                nombre_completo = sisca_logic._nombre_completo(a)
                cui = a.get("cui", "")
                fecha_nac = a.get("fecha_nac", "")
                if not cui:
                    raw = f"{nombre_completo}{fecha_nac}{a.get('grado', '')}{a.get('seccion', '')}"
                    cui = f"TMP-{hashlib.sha256(raw.encode()).hexdigest()[:12].upper()}"
                edad = sisca_logic.calcular_edad_a_fecha_corte(dia, mes, anio, fc)
                sexo_db = "Femenino" if gen == "F" else "Masculino"
                execute("""
                    INSERT INTO estudiantes
                        (cui, usuario_id, nombre_completo, sexo, fecha_nacimiento, grado, seccion)
                    VALUES (%s, %s, %s, %s, %s, %s, %s)
                    ON CONFLICT (cui, usuario_id) DO NOTHING
                """, (cui, uid, nombre_completo, sexo_db, fecha_nac,
                      a.get("grado", ""), a.get("seccion", "")))
                execute("""
                    UPDATE estudiantes SET grado = %s, seccion = %s WHERE cui = %s AND usuario_id = %s
                """, (a.get("grado", ""), a.get("seccion", ""), cui, uid))
                execute("""
                    INSERT INTO registros_salud
                        (cui_estudiante, codigo_centro, tipo_intervencion,
                         campana, fecha_aplicacion, fecha_corte, edad_calculo, usuario_id)
                    VALUES (%s, %s, %s, %s, %s, %s, %s, %s)
                """, (cui, codigo, "Desparasitacion", "Primera",
                      date.today().isoformat(), fecha_corte_db, edad, uid))

            try:
                commit()
                total_escuelas += 1
            except Exception as e_commit:
                try:
                    rollback()
                except Exception:
                    pass
                errores.append(f"«{archivo.filename}»: error al guardar en BD - {e_commit}")

        except Exception as e:
            try:
                rollback()
            except Exception:
                pass
            errores.append(f"«{archivo.filename}»: {type(e).__name__} - {e}")
        finally:
            if tmp_path:
                try:
                    os.unlink(tmp_path)
                except Exception:
                    pass

    if total_escuelas:
        flash(f"¡Éxito! Se procesaron correctamente {total_escuelas} escuela(s) en la base de datos.", "success")
    for err in errores:
        flash(err, "warning" if total_escuelas else "danger")

    return redirect(url_for("consolidado"))


# ---------------------------------------------------------------------------
# Ejecución
# ---------------------------------------------------------------------------
if __name__ == "__main__":
    print(f"Servidor TSR iniciado en http://127.0.0.1:5000")
    app.run(debug=True, host="127.0.0.1", port=5000)
