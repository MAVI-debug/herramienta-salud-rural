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

    # ── Migración silenciosa a PK compuesta con usuario_id ──────────────────
    cur = conn.execute("PRAGMA table_info(escuelas)")
    cols = [r[1] for r in cur.fetchall()]
    if "usuario_id" not in cols:
        conn.execute("PRAGMA foreign_keys=OFF")
        try:
            conn.executescript("""
                CREATE TABLE escuelas_nueva (
                    codigo_centro TEXT NOT NULL,
                    usuario_id    INTEGER NOT NULL,
                    nombre_centro TEXT NOT NULL,
                    tipo_centro   TEXT CHECK(tipo_centro IN ('PÚBLICO','PRIVADO')) NOT NULL,
                    servicio_salud TEXT NOT NULL,
                    PRIMARY KEY (codigo_centro, usuario_id),
                    FOREIGN KEY (usuario_id) REFERENCES usuarios(id)
                );
                INSERT INTO escuelas_nueva
                    (codigo_centro, usuario_id, nombre_centro, tipo_centro, servicio_salud)
                SELECT e.codigo_centro, u.id,
                       e.nombre_centro, e.tipo_centro, e.servicio_salud
                FROM escuelas e, usuarios u;
                DROP TABLE escuelas;
                ALTER TABLE escuelas_nueva RENAME TO escuelas;

                CREATE TABLE estudiantes_nueva (
                    cui              TEXT NOT NULL,
                    usuario_id       INTEGER NOT NULL,
                    nombre_completo  TEXT NOT NULL,
                    sexo             TEXT CHECK(sexo IN ('Femenino','Masculino')) NOT NULL,
                    fecha_nacimiento TEXT NOT NULL,
                    grado            TEXT DEFAULT '',
                    seccion          TEXT DEFAULT '',
                    PRIMARY KEY (cui, usuario_id),
                    FOREIGN KEY (usuario_id) REFERENCES usuarios(id)
                );
                INSERT INTO estudiantes_nueva
                    (cui, usuario_id, nombre_completo, sexo, fecha_nacimiento, grado, seccion)
                SELECT e.cui, u.id,
                       e.nombre_completo, e.sexo, e.fecha_nacimiento,
                       COALESCE(e.grado, ''), COALESCE(e.seccion, '')
                FROM estudiantes e, usuarios u;
                DROP TABLE estudiantes;
                ALTER TABLE estudiantes_nueva RENAME TO estudiantes;

                CREATE TABLE registros_salud_nueva (
                    id                INTEGER PRIMARY KEY AUTOINCREMENT,
                    cui_estudiante    TEXT,
                    codigo_centro     TEXT,
                    tipo_intervencion TEXT CHECK(tipo_intervencion IN ('Desparasitación','Fluorización')) NOT NULL,
                    campana           TEXT CHECK(campana IN ('Primera','Segunda')) NOT NULL,
                    fecha_aplicacion  TEXT NOT NULL,
                    fecha_corte       TEXT NOT NULL DEFAULT '31/03/2026',
                    edad_calculo      INTEGER DEFAULT NULL,
                    usuario_id        INTEGER NOT NULL,
                    FOREIGN KEY (cui_estudiante, usuario_id) REFERENCES estudiantes(cui, usuario_id)
                        ON DELETE SET NULL ON UPDATE CASCADE,
                    FOREIGN KEY (codigo_centro, usuario_id) REFERENCES escuelas(codigo_centro, usuario_id)
                        ON DELETE CASCADE ON UPDATE CASCADE,
                    FOREIGN KEY (usuario_id) REFERENCES usuarios(id)
                        ON DELETE SET NULL ON UPDATE CASCADE
                );
                INSERT INTO registros_salud_nueva
                    (cui_estudiante, codigo_centro, tipo_intervencion,
                     campana, fecha_aplicacion, fecha_corte, edad_calculo, usuario_id)
                SELECT r.cui_estudiante, r.codigo_centro, r.tipo_intervencion,
                       r.campana, r.fecha_aplicacion,
                       COALESCE(r.fecha_corte, '31/03/2026'), r.edad_calculo, u.id
                FROM registros_salud r, usuarios u;
                DROP TABLE registros_salud;
                ALTER TABLE registros_salud_nueva RENAME TO registros_salud;

                CREATE INDEX IF NOT EXISTS idx_estudiantes_sexo ON estudiantes(sexo);
                CREATE INDEX IF NOT EXISTS idx_estudiantes_fecha_nacimiento ON estudiantes(fecha_nacimiento);
                CREATE INDEX IF NOT EXISTS idx_registros_cui ON registros_salud(cui_estudiante);
                CREATE INDEX IF NOT EXISTS idx_registros_centro ON registros_salud(codigo_centro);
                CREATE INDEX IF NOT EXISTS idx_registros_intervencion ON registros_salud(tipo_intervencion, campana);
                CREATE INDEX IF NOT EXISTS idx_registros_usuario ON registros_salud(usuario_id);
            """)
        except Exception:
            conn.rollback()
            raise
        finally:
            conn.execute("PRAGMA foreign_keys=ON")

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

    # Migracion: relajar CHECK constraints para aceptar variantes sin tilde
    try:
        conn.execute("PRAGMA table_info(escuelas)")
        cur_check = conn.execute(
            "SELECT sql FROM sqlite_master WHERE type='table' AND name='escuelas'"
        )
        ddl = cur_check.fetchone()[0]
        if "CHECK" in (ddl or ""):
            conn.execute("PRAGMA foreign_keys=OFF")
            try:
                conn.executescript("""
                    CREATE TABLE escuelas_new (
                        codigo_centro TEXT NOT NULL,
                        usuario_id    INTEGER NOT NULL,
                        nombre_centro TEXT NOT NULL,
                        tipo_centro   TEXT NOT NULL DEFAULT 'PUBLICO',
                        servicio_salud TEXT NOT NULL DEFAULT '',
                        PRIMARY KEY (codigo_centro, usuario_id),
                        FOREIGN KEY (usuario_id) REFERENCES usuarios(id)
                    );
                    INSERT INTO escuelas_new
                        SELECT codigo_centro, usuario_id, nombre_centro,
                               COALESCE(tipo_centro, 'PUBLICO'),
                               COALESCE(servicio_salud, '')
                        FROM escuelas;
                    DROP TABLE escuelas;
                    ALTER TABLE escuelas_new RENAME TO escuelas;

                    CREATE TABLE registros_salud_new (
                        id                INTEGER PRIMARY KEY AUTOINCREMENT,
                        cui_estudiante    TEXT,
                        codigo_centro     TEXT,
                        tipo_intervencion TEXT NOT NULL DEFAULT 'Desparasitacion',
                        campana           TEXT NOT NULL DEFAULT 'Primera',
                        fecha_aplicacion  TEXT NOT NULL,
                        fecha_corte       TEXT NOT NULL DEFAULT '31/03/2026',
                        edad_calculo      INTEGER DEFAULT NULL,
                        usuario_id        INTEGER NOT NULL,
                        FOREIGN KEY (usuario_id) REFERENCES usuarios(id)
                            ON DELETE SET NULL ON UPDATE CASCADE
                    );
                    INSERT INTO registros_salud_new
                        SELECT id, cui_estudiante, codigo_centro,
                               COALESCE(tipo_intervencion, 'Desparasitacion'),
                               COALESCE(campana, 'Primera'),
                               fecha_aplicacion, fecha_corte,
                               edad_calculo, usuario_id
                        FROM registros_salud;
                    DROP TABLE registros_salud;
                    ALTER TABLE registros_salud_new RENAME TO registros_salud;

                    CREATE INDEX IF NOT EXISTS idx_registros_cui ON registros_salud(cui_estudiante);
                    CREATE INDEX IF NOT EXISTS idx_registros_centro ON registros_salud(codigo_centro);
                    CREATE INDEX IF NOT EXISTS idx_registros_intervencion ON registros_salud(tipo_intervencion, campana);
                    CREATE INDEX IF NOT EXISTS idx_registros_usuario ON registros_salud(usuario_id);
                """)
            except Exception:
                conn.rollback()
            finally:
                conn.execute("PRAGMA foreign_keys=ON")
    except Exception:
        pass

    # ── Migracion: columna es_admin ────────────────────────────────────
    try:
        conn.execute("ALTER TABLE usuarios ADD COLUMN es_admin INTEGER NOT NULL DEFAULT 0")
    except sqlite3.OperationalError:
        pass

    # Marcar tsr_demo como administrador
    try:
        conn.execute("UPDATE usuarios SET es_admin = 1 WHERE usuario = 'tsr_demo'")
        conn.commit()
    except Exception:
        pass

    # ── Migracion: columna fecha_corte en usuarios ─────────────────────
    try:
        from datetime import date as _date
        _hoy = _date.today().strftime("%d/%m/%Y")
        conn.execute(f"ALTER TABLE usuarios ADD COLUMN fecha_corte TEXT NOT NULL DEFAULT '{_hoy}'")
    except sqlite3.OperationalError:
        pass

    # ── Migracion: tabla jornadas_realizadas ──────────────────────────
    try:
        conn.execute("""
            CREATE TABLE IF NOT EXISTS jornadas_realizadas (
                id                INTEGER PRIMARY KEY AUTOINCREMENT,
                usuario_id        INTEGER NOT NULL,
                tipo_jornada      TEXT NOT NULL,
                fecha_jornada     TEXT NOT NULL,
                anio              INTEGER NOT NULL,
                fecha_corte       TEXT NOT NULL,
                total_entran_rango INTEGER NOT NULL DEFAULT 0,
                total_no_entran   INTEGER NOT NULL DEFAULT 0,
                total_f           INTEGER NOT NULL DEFAULT 0,
                total_m           INTEGER NOT NULL DEFAULT 0,
                total_general     INTEGER NOT NULL DEFAULT 0,
                observaciones     TEXT DEFAULT '',
                FOREIGN KEY (usuario_id) REFERENCES usuarios(id)
            )
        """)
        conn.commit()
    except Exception:
        pass


def _run_migrations_pg(conn):
    cur = conn.cursor()

    # ── Migración a PK compuesta con usuario_id ────────────────────────────
    try:
        cur.execute("""
            SELECT column_name FROM information_schema.columns
            WHERE table_name='escuelas' AND column_name='usuario_id'
        """)
        if not cur.fetchone():
            cur.execute("ALTER TABLE escuelas ADD COLUMN usuario_id INTEGER NOT NULL DEFAULT 1")
            cur.execute("ALTER TABLE escuelas DROP CONSTRAINT escuelas_pkey CASCADE")
            cur.execute("ALTER TABLE escuelas ADD PRIMARY KEY (codigo_centro, usuario_id)")
            cur.execute("""
                INSERT INTO escuelas (codigo_centro, usuario_id, nombre_centro, tipo_centro, servicio_salud)
                SELECT e.codigo_centro, u.id, e.nombre_centro, e.tipo_centro, e.servicio_salud
                FROM escuelas e, usuarios u
                WHERE e.usuario_id = 1 AND u.id != 1
                ON CONFLICT (codigo_centro, usuario_id) DO NOTHING
            """)
            cur.execute("ALTER TABLE escuelas ALTER COLUMN usuario_id SET NOT NULL")

            cur.execute("ALTER TABLE estudiantes ADD COLUMN usuario_id INTEGER NOT NULL DEFAULT 1")
            cur.execute("ALTER TABLE estudiantes DROP CONSTRAINT estudiantes_pkey CASCADE")
            cur.execute("ALTER TABLE estudiantes ADD PRIMARY KEY (cui, usuario_id)")
            cur.execute("""
                INSERT INTO estudiantes (cui, usuario_id, nombre_completo, sexo, fecha_nacimiento, grado, seccion)
                SELECT e.cui, u.id, e.nombre_completo, e.sexo, e.fecha_nacimiento, e.grado, e.seccion
                FROM estudiantes e, usuarios u
                WHERE e.usuario_id = 1 AND u.id != 1
                ON CONFLICT (cui, usuario_id) DO NOTHING
            """)
            cur.execute("ALTER TABLE estudiantes ALTER COLUMN usuario_id SET NOT NULL")

            # registros_salud: update NULLs to default user, then duplicate
            cur.execute("""
                INSERT INTO registros_salud (cui_estudiante, codigo_centro, tipo_intervencion,
                    campana, fecha_aplicacion, fecha_corte, edad_calculo, usuario_id)
                SELECT r.cui_estudiante, r.codigo_centro, r.tipo_intervencion,
                    r.campana, r.fecha_aplicacion,
                    COALESCE(r.fecha_corte, '31/03/2026'), r.edad_calculo, u.id
                FROM registros_salud r, usuarios u
                WHERE (r.usuario_id IS NULL OR r.usuario_id = 1) AND u.id != 1
            """)
            cur.execute("""
                UPDATE registros_salud SET usuario_id = 1
                WHERE usuario_id IS NULL
            """)
    except Exception:
        conn.rollback()
        raise

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

    # Migracion: relajar CHECK constraints para aceptar variantes sin tilde
    try:
        cur.execute("""
            SELECT c.conname, n.nspname, r.relname
            FROM pg_constraint c
            JOIN pg_class r ON c.conrelid = r.oid
            JOIN pg_namespace n ON r.relnamespace = n.oid
            WHERE c.contype = 'c'
              AND r.relname IN ('escuelas','registros_salud')
        """)
        for row in cur.fetchall():
            schema = row['nspname']
            table = row['relname']
            cname = row['conname']
            try:
                cur.execute(f"ALTER TABLE {schema}.{table} DROP CONSTRAINT IF EXISTS {cname}")
            except Exception:
                pass
    except Exception:
        pass

    # ── Migracion: columna es_admin ────────────────────────────────────
    try:
        cur.execute("ALTER TABLE usuarios ADD COLUMN IF NOT EXISTS es_admin INTEGER NOT NULL DEFAULT 0")
        cur.execute("UPDATE usuarios SET es_admin = 1 WHERE usuario = 'tsr_demo'")
    except Exception:
        pass

    # ── Migracion: columna fecha_corte en usuarios ─────────────────────
    try:
        from datetime import date as _date
        _hoy = _date.today().strftime("%d/%m/%Y")
        cur.execute("ALTER TABLE usuarios ADD COLUMN IF NOT EXISTS fecha_corte TEXT NOT NULL DEFAULT %s", (_hoy,))
    except Exception:
        pass

    # ── Migracion: tabla jornadas_realizadas ──────────────────────────
    try:
        cur.execute("""
            CREATE TABLE IF NOT EXISTS jornadas_realizadas (
                id                SERIAL PRIMARY KEY,
                usuario_id        INTEGER NOT NULL,
                tipo_jornada      TEXT NOT NULL,
                fecha_jornada     TEXT NOT NULL,
                anio              INTEGER NOT NULL,
                fecha_corte       TEXT NOT NULL,
                total_entran_rango INTEGER NOT NULL DEFAULT 0,
                total_no_entran   INTEGER NOT NULL DEFAULT 0,
                total_f           INTEGER NOT NULL DEFAULT 0,
                total_m           INTEGER NOT NULL DEFAULT 0,
                total_general     INTEGER NOT NULL DEFAULT 0,
                observaciones     TEXT DEFAULT '',
                FOREIGN KEY (usuario_id) REFERENCES usuarios(id)
            )
        """)
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
