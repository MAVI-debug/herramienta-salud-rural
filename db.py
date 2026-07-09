"""
db.py — Abstracción de base de datos híbrida SQLite / PostgreSQL.
=================================================================
Usa DATABASE_URL (variable de entorno) para PostgreSQL en producción;
cae a SQLite local si la variable no existe.
"""

import os
from flask import g

DATABASE_URL = os.environ.get("DATABASE_URL")
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
DB_PATH = os.path.join(BASE_DIR, "salud_rural.db")


def get_db():
    if "db" not in g:
        if DATABASE_URL:
            g.db = _connect_pg()
        else:
            g.db = _connect_sqlite()
    return g.db


def _connect_sqlite():
    import sqlite3

    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    conn.execute("PRAGMA foreign_keys=ON;")
    _run_migrations_sqlite(conn)
    return conn


def _connect_pg():
    import psycopg2
    from psycopg2.extras import RealDictCursor

    conn = psycopg2.connect(DATABASE_URL, cursor_factory=RealDictCursor)
    _run_migrations_pg(conn)
    return conn


def _is_pg():
    return DATABASE_URL is not None


def execute(sql, params=None):
    """Ejecuta SQL adaptando placeholders %s → ? para SQLite."""
    db = get_db()
    if not _is_pg() and params is not None:
        sql = sql.replace("%s", "?")
    if _is_pg():
        from psycopg2.extras import RealDictCursor
        cur = db.cursor(cursor_factory=RealDictCursor)
        cur.execute(sql, params or ())
    else:
        cur = db.execute(sql, params or ())
    return cur


def fetchone(sql, params=None):
    cur = execute(sql, params)
    row = cur.fetchone()
    if _is_pg():
        cur.close()
    return row


def fetchall(sql, params=None):
    cur = execute(sql, params)
    return cur.fetchall()


def commit():
    db = get_db()
    if _is_pg():
        db.commit()
    else:
        db.commit()


# ---------------------------------------------------------------------------
# Migraciones silenciosas
# ---------------------------------------------------------------------------

def _run_migrations_sqlite(conn):
    import sqlite3

    for col_def in [
        "ALTER TABLE registros_salud ADD COLUMN fecha_corte TEXT NOT NULL DEFAULT '31/03/2026'",
        "ALTER TABLE estudiantes ADD COLUMN grado TEXT DEFAULT ''",
        "ALTER TABLE estudiantes ADD COLUMN seccion TEXT DEFAULT ''",
        "ALTER TABLE registros_salud ADD COLUMN edad_calculo INTEGER DEFAULT NULL",
    ]:
        try:
            conn.execute(col_def)
        except sqlite3.OperationalError:
            pass


def _run_migrations_pg(conn):
    cur = conn.cursor()
    for col_def in [
        "ALTER TABLE registros_salud ADD COLUMN IF NOT EXISTS fecha_corte TEXT NOT NULL DEFAULT '31/03/2026'",
        "ALTER TABLE estudiantes ADD COLUMN IF NOT EXISTS grado TEXT DEFAULT ''",
        "ALTER TABLE estudiantes ADD COLUMN IF NOT EXISTS seccion TEXT DEFAULT ''",
        "ALTER TABLE registros_salud ADD COLUMN IF NOT EXISTS edad_calculo INTEGER DEFAULT NULL",
    ]:
        try:
            cur.execute(col_def)
        except Exception:
            pass
    conn.commit()
    cur.close()


# ---------------------------------------------------------------------------
# Teardown (registrar en app.py con: app.teardown_appcontext(close_db))
# ---------------------------------------------------------------------------

def close_db(exception=None):
    db = g.pop("db", None)
    if db is not None:
        try:
            db.close()
        except Exception:
            pass


def rollback():
    """Rollback transaction (usado en cargas masivas)."""
    db = get_db()
    try:
        db.rollback()
    except Exception:
        pass
