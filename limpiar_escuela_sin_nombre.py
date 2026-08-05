"""
limpiar_escuela_sin_nombre.py
=============================
Elimina por completo la "ESCUELA SIN NOMBRE" (o escuelas corruptas sin nombre)
del usuario administrador 'mavi' y limpia los datos huerfanos asociados para
que el consolidado vuelva a la normalidad.

Funciona con SQLite (local) y PostgreSQL (produccion, via DATABASE_URL).
Uso:
    python limpiar_escuela_sin_nombre.py
"""

import sys

from flask import Flask

import db

_CTX = Flask(__name__)


def _eliminar_huerfanos(uid):
    """Barrido general de huerfanos para el usuario dado."""
    total = 0

    cur = db.execute(
        "DELETE FROM registros_salud WHERE usuario_id = %s"
        " AND codigo_centro NOT IN (SELECT codigo_centro FROM escuelas WHERE usuario_id = %s)",
        (uid, uid),
    )
    total += cur.rowcount if cur.rowcount and cur.rowcount > 0 else 0

    cur = db.execute(
        "DELETE FROM registros_salud WHERE usuario_id = %s"
        " AND (cui_estudiante IS NULL OR cui_estudiante NOT IN"
        " (SELECT cui FROM estudiantes WHERE usuario_id = %s))",
        (uid, uid),
    )
    total += cur.rowcount if cur.rowcount and cur.rowcount > 0 else 0

    cur = db.execute(
        "DELETE FROM estudiantes WHERE usuario_id = %s"
        " AND cui NOT IN (SELECT cui_estudiante FROM registros_salud WHERE usuario_id = %s"
        " AND cui_estudiante IS NOT NULL)",
        (uid, uid),
    )
    total += cur.rowcount if cur.rowcount and cur.rowcount > 0 else 0

    return total


def main():
    with _CTX.app_context():
        _ejecutar()


def _ejecutar():
    mavi = db.fetchone("SELECT id, usuario FROM usuarios WHERE usuario = 'mavi'")
    if not mavi:
        print("AVISO: no existe el usuario 'mavi' en esta base de datos.")
        return

    uid = mavi["id"]
    print(f"Usuario admin: {mavi['usuario']} (id={uid})")

    objetivo = db.fetchall(
        "SELECT codigo_centro, nombre_centro FROM escuelas WHERE usuario_id = %s"
        " AND (TRIM(nombre_centro) = '' OR UPPER(nombre_centro) LIKE '%SIN NOMBRE%')",
        (uid,),
    )
    print(f"Escuelas sin nombre encontradas: {len(objetivo)}")
    for f in objetivo:
        print("   - codigo:", repr(f["codigo_centro"]), "| nombre:", repr(f["nombre_centro"]))

    if not objetivo:
        print("No hay escuelas sin nombre que limpiar.")
    else:
        for f in objetivo:
            cod = f["codigo_centro"]
            db.execute(
                "DELETE FROM registros_salud WHERE codigo_centro = %s AND usuario_id = %s",
                (cod, uid),
            )
            db.execute(
                "DELETE FROM estudiantes WHERE usuario_id = %s"
                " AND cui NOT IN (SELECT cui_estudiante FROM registros_salud"
                " WHERE usuario_id = %s AND cui_estudiante IS NOT NULL)",
                (uid, uid),
            )
            db.execute(
                "DELETE FROM escuelas WHERE codigo_centro = %s AND usuario_id = %s",
                (cod, uid),
            )
            print(f"   -> eliminada escuela '{cod}' (con registros y estudiantes huerfanos).")

    huerfanos = _eliminar_huerfanos(uid)
    if huerfanos:
        print(f"Barrido de huerfanos: {huerfanos} registro(s) eliminado(s).")

    db.commit()
    print("Limpieza completada y confirmada.")


if __name__ == "__main__":
    try:
        main()
    except Exception as e:
        try:
            with _CTX.app_context():
                db.rollback()
        except Exception:
            pass
        print("ERROR durante la limpieza:", e)
        sys.exit(1)
