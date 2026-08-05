"""
crear_admin.py — Garantiza la cuenta de administrador "mavi".
=============================================================
Uso (SQLite local o PostgreSQL vía DATABASE_URL):

    set MAVI_PASSWORD=TU_CLAVE          (PowerShell)
    python crear_admin.py

    MAVI_PASSWORD=TU_CLAVE python crear_admin.py   (Linux / Render shell)

Solo la cuenta "mavi" (es_admin=1) puede crear usuarios en la plataforma.
Opcionales: MAVI_NOMBRE, MAVI_CARGO, MAVI_AREA, MAVI_DISTRITO.
La contraseña NUNCA queda escrita en el repositorio.
"""

import os
import sys

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, BASE_DIR)

from db import execute, fetchone, commit  # noqa: E402


def main():
    from app import app, hash_contrasena  # noqa: E402

    usuario = "mavi"
    password = os.environ.get("MAVI_PASSWORD", "").strip()
    nombre = os.environ.get("MAVI_NOMBRE", "").strip() or None
    cargo = os.environ.get("MAVI_CARGO", "Técnico en Salud Rural (TSR)").strip()
    area = os.environ.get("MAVI_AREA", "").strip()
    distrito = os.environ.get("MAVI_DISTRITO", "").strip()

    if not password:
        print(">>> AVISO: no se definió MAVI_PASSWORD; se conservará la contraseña actual de mavi.")

    with app.app_context():
        fila = fetchone("SELECT id FROM usuarios WHERE usuario = %s", (usuario,))

        if fila:
            updates = ["es_admin = 1"]
            params = []
            if password:
                updates.append("contrasena_hash = %s")
                params.append(hash_contrasena(password))
            if nombre:
                updates.append("nombre_responsable = %s")
                params.append(nombre)
            if cargo:
                updates.append("cargo = %s")
                params.append(cargo)
            if area:
                updates.append("area_salud = %s")
                params.append(area)
            if distrito:
                updates.append("distrito_salud = %s")
                params.append(distrito)
            params.append(usuario)
            execute("UPDATE usuarios SET " + ", ".join(updates) + " WHERE usuario = %s", params)
            print(">>> mavi actualizado: es_admin=1" + (" y contraseña" if password else "") + ".")
        else:
            if not password:
                print("ERROR: el usuario 'mavi' no existe y no se proporcionó MAVI_PASSWORD.")
                sys.exit(1)
            nombre = nombre or "MAVI"
            execute("""
                INSERT INTO usuarios
                    (usuario, contrasena_hash, nombre_responsable, cargo, area_salud, distrito_salud, es_admin)
                VALUES (%s, %s, %s, %s, %s, %s, 1)
            """, (usuario, hash_contrasena(password), nombre, cargo, area, distrito))
            print(">>> mavi creado como administrador (es_admin=1).")

        commit()
        print(">>> Listo. Solo 'mavi' puede crear usuarios (botón +).")


if __name__ == "__main__":
    main()
