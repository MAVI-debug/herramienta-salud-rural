"""
crear_base_datos.py
===================
Script independiente para crear la infraestructura SQLite
destinada a la futura migración web de TSR.
"""

import sqlite3
from werkzeug.security import generate_password_hash


def inicializar_bd(ruta_db="salud_rural.db", crear_usuario_demo=False):
    conn = sqlite3.connect(ruta_db)
    cursor = conn.cursor()

    cursor.execute("PRAGMA journal_mode=WAL;")
    cursor.execute("PRAGMA foreign_keys=ON;")

    # 1. USUARIOS
    cursor.execute("""
        CREATE TABLE IF NOT EXISTS usuarios (
            id              INTEGER PRIMARY KEY AUTOINCREMENT,
            usuario         TEXT UNIQUE NOT NULL,
            contrasena_hash TEXT NOT NULL,
            nombre_responsable TEXT NOT NULL,
            cargo           TEXT NOT NULL,
            area_salud      TEXT NOT NULL,
            distrito_salud  TEXT NOT NULL
        )
    """)

    # 2. ESCUELAS / CENTROS EDUCATIVOS
    cursor.execute("""
        CREATE TABLE IF NOT EXISTS escuelas (
            codigo_centro TEXT PRIMARY KEY,
            nombre_centro TEXT NOT NULL,
            tipo_centro   TEXT CHECK(tipo_centro IN ('PÚBLICO', 'PRIVADO')) NOT NULL,
            servicio_salud TEXT NOT NULL
        )
    """)

    # 3. ESTUDIANTES
    cursor.execute("""
        CREATE TABLE IF NOT EXISTS estudiantes (
            cui              TEXT PRIMARY KEY,
            nombre_completo  TEXT NOT NULL,
            sexo             TEXT CHECK(sexo IN ('Femenino', 'Masculino')) NOT NULL,
            fecha_nacimiento TEXT NOT NULL,
            grado            TEXT DEFAULT '',
            seccion          TEXT DEFAULT ''
        )
    """)
    for col in ("grado", "seccion"):
        try:
            cursor.execute(f"ALTER TABLE estudiantes ADD COLUMN {col} TEXT DEFAULT ''")
        except sqlite3.OperationalError:
            pass

    # 4. REGISTROS DE SALUD
    cursor.execute("""
        CREATE TABLE IF NOT EXISTS registros_salud (
            id                INTEGER PRIMARY KEY AUTOINCREMENT,
            cui_estudiante    TEXT,
            codigo_centro     TEXT,
            tipo_intervencion TEXT CHECK(tipo_intervencion IN ('Desparasitación', 'Fluorización')) NOT NULL,
            campana           TEXT CHECK(campana IN ('Primera', 'Segunda')) NOT NULL,
            fecha_aplicacion  TEXT NOT NULL,
            fecha_corte       TEXT NOT NULL DEFAULT '31/03/2026',
            edad_calculo      INTEGER DEFAULT NULL,
            usuario_id        INTEGER,
            FOREIGN KEY (cui_estudiante) REFERENCES estudiantes(cui)
                ON DELETE SET NULL ON UPDATE CASCADE,
            FOREIGN KEY (codigo_centro) REFERENCES escuelas(codigo_centro)
                ON DELETE CASCADE ON UPDATE CASCADE,
            FOREIGN KEY (usuario_id) REFERENCES usuarios(id)
                ON DELETE SET NULL ON UPDATE CASCADE
        )
    """)
    # Migración silenciosa para bases existentes sin la columna
    try:
        cursor.execute("ALTER TABLE registros_salud ADD COLUMN fecha_corte TEXT NOT NULL DEFAULT '31/03/2026'")
    except sqlite3.OperationalError:
        pass  # ya existe

    # 5. ÍNDICES PARA CONSULTAS FRECUENTES
    cursor.execute("""
        CREATE INDEX IF NOT EXISTS idx_usuarios_usuario
            ON usuarios(usuario)
    """)
    cursor.execute("""
        CREATE INDEX IF NOT EXISTS idx_estudiantes_sexo
            ON estudiantes(sexo)
    """)
    cursor.execute("""
        CREATE INDEX IF NOT EXISTS idx_estudiantes_fecha_nacimiento
            ON estudiantes(fecha_nacimiento)
    """)
    cursor.execute("""
        CREATE INDEX IF NOT EXISTS idx_registros_cui
            ON registros_salud(cui_estudiante)
    """)
    cursor.execute("""
        CREATE INDEX IF NOT EXISTS idx_registros_centro
            ON registros_salud(codigo_centro)
    """)
    cursor.execute("""
        CREATE INDEX IF NOT EXISTS idx_registros_intervencion
            ON registros_salud(tipo_intervencion, campana)
    """)
    cursor.execute("""
        CREATE INDEX IF NOT EXISTS idx_registros_usuario
            ON registros_salud(usuario_id)
    """)

    if crear_usuario_demo:
        cursor.execute("""
            INSERT INTO usuarios
                (usuario, contrasena_hash, nombre_responsable, cargo, area_salud, distrito_salud)
            VALUES (%s, %s, %s, %s, %s, %s)
            ON CONFLICT (usuario) DO NOTHING
        """, (
            "tsr_demo",
            generate_password_hash("demo123"),
            "Andrés Romeo Mazariegos Vicente",
            "Técnico en Salud Rural (TSR)",
            "TOTONICAPÁN",
            "TOTONICAPÁN",
        ))
        print("  -> Usuario demo creado: tsr_demo / demo123")

    conn.commit()
    conn.close()
    print(f"OK: Base de datos '{ruta_db}' lista - {sqlite3.sqlite_version}")


if __name__ == "__main__":
    inicializar_bd(crear_usuario_demo=True)
