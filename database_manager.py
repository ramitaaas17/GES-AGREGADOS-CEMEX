"""Gestión de SQLite para la captura RutaV3_2026."""

import sqlite3
from datetime import date, datetime
from sqlite3 import Error

# Todos los campos manuales (sin fórmula) en el orden exacto de la GUI (MANUAL_FIELDS)
TABLA_COLS = [
    "org_vta", "cedis", "nombre_cedis", "sf", "nombre_sf",
    "planta_obra", "nombre_planta", "articulo", "desc_articulo", "pv",
    "modalidad", "um_vta",
    "mp_vta_m3", "flete_vta_m3",
    "mp_comp_m3", "ca_comp_m3", "flete_comp_m3", "integral_comp_m3",
    "tipo", "fecha_procesamiento", "tipo_cliente", "observacion",
    "fecha_solicitud", "incidencia", "motivo", "mes", "fecha_conf",
    "formato_destino",
    # Auditoría (Fase 2 - concurrencia): fecha/hora de la última escritura
    # sobre cada fila, usada por la Sincronización de Recuperación para
    # detectar qué registros se trabajaron hoy.
    "timestamp_modificacion",
]

# Tipos SQL por columna (REAL para numéricos, TEXT para el resto)
TIPOS_SQL = {
    "pv": "REAL", "mp_vta_m3": "REAL", "flete_vta_m3": "REAL",
    "mp_comp_m3": "REAL", "ca_comp_m3": "REAL",
    "flete_comp_m3": "REAL", "integral_comp_m3": "REAL",
}


def create_connection(db_file):
    conn = None
    try:
        conn = sqlite3.connect(db_file, check_same_thread=False)
        return conn
    except Error as e:
        print(f"Error conectando a BD: {e}")
    return conn


def _col_sql(col):
    tipo = TIPOS_SQL.get(col, "TEXT")
    return f'"{col}" {tipo}'


def create_table(conn):
    """Crea la tabla staging_rutas con los 28 campos manuales (sin fórmula)."""
    try:
        c = conn.cursor()
        cols_sql = ', '.join([_col_sql(c) for c in TABLA_COLS])
        c.execute(f"""
            CREATE TABLE IF NOT EXISTS staging_rutas (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                {cols_sql}
            );
        """)
    except Error as e:
        print(f"Error creando tabla: {e}")


def migrate_if_needed(conn):
    """Migra staging_rutas a la versión más reciente (28 columnas).

    Estrategia: añade con ALTER TABLE las columnas que falten, sin perder datos.
    """
    try:
        cur = conn.cursor()
        cur.execute("SELECT name FROM sqlite_master WHERE type='table' AND name='staging_rutas'")
        if not cur.fetchone():
            create_table(conn)
            return

        cols = [r[1] for r in cur.execute("PRAGMA table_info(staging_rutas)").fetchall()]

        # Añadir columnas faltantes (preserva datos existentes)
        faltantes = [c for c in TABLA_COLS if c not in cols]
        for col in faltantes:
            cur.execute(f'ALTER TABLE staging_rutas ADD COLUMN {_col_sql(col)} DEFAULT NULL')
            print(f"Migrado: columna '{col}' añadida.")

        # Si la tabla vieja tenía 'um_venta' (v0) en vez de 'um_vta', renombrar
        if "um_venta" in cols and "um_vta" not in cols:
            cur.execute("ALTER TABLE staging_rutas RENAME COLUMN um_venta TO um_vta")
            print("Migrado: columna 'um_venta' renombrada a 'um_vta'.")

        # Si la tabla vieja tenía 'costo_mp_venta'/'costo_flete_venta' (v0),
        # volcar a mp_vta_m3/flete_vta_m3 si estas están vacías
        if "costo_mp_venta" in cols and "mp_vta_m3" in cols:
            cur.execute("""
                UPDATE staging_rutas SET mp_vta_m3 = costo_mp_venta
                WHERE (mp_vta_m3 IS NULL OR mp_vta_m3 = '') AND costo_mp_venta IS NOT NULL
            """)
            cur.execute("""
                UPDATE staging_rutas SET flete_vta_m3 = costo_flete_venta
                WHERE (flete_vta_m3 IS NULL OR flete_vta_m3 = '') AND costo_flete_venta IS NOT NULL
            """)
            print("Migrado: costos v0 volcados a mp_vta_m3/flete_vta_m3.")

        conn.commit()
    except Error as e:
        print(f"Error en migración: {e}")


def insert_data(conn, data):
    """Inserta una fila con los campos manuales. Retorna lastrowid o None.

    `data` puede venir con la longitud anterior (sin 'timestamp_modificacion')
    o con la longitud completa de TABLA_COLS; en ambos casos, si el valor de
    'timestamp_modificacion' llega vacío, se completa automáticamente con la
    fecha/hora actual (auditoría de la Fase 2 de concurrencia).
    """
    data = list(data)
    if len(data) == len(TABLA_COLS) - 1:
        data.append(None)
    ts_idx = TABLA_COLS.index("timestamp_modificacion")
    if not data[ts_idx]:
        data[ts_idx] = datetime.now().strftime("%d/%m/%Y %H:%M:%S")
    data = tuple(data)

    placeholders = ",".join(["?"] * len(TABLA_COLS))
    sql = f""" INSERT INTO staging_rutas({','.join(f'"{c}"' for c in TABLA_COLS)})
               VALUES({placeholders}) """
    try:
        cur = conn.cursor()
        cur.execute(sql, data)
        conn.commit()
        return cur.lastrowid
    except Error as e:
        print(f"Error insertando: {e}")
    return None


def touch_timestamp(conn, row_id):
    """Actualiza 'timestamp_modificacion' de una fila con la fecha/hora actual.

    Helper reutilizable para que cualquier UPDATE de negocio (fuera de este
    módulo) marque la fila como modificada hoy, de cara a la Sincronización
    de Recuperación.
    """
    try:
        cur = conn.cursor()
        cur.execute(
            'UPDATE staging_rutas SET "timestamp_modificacion"=? WHERE id=?',
            (datetime.now().strftime("%d/%m/%Y %H:%M:%S"), row_id),
        )
        conn.commit()
    except Error as e:
        print(f"Error actualizando timestamp_modificacion: {e}")


def get_rows_modificadas_hoy(conn):
    """Retorna las filas de staging_rutas cuyo 'timestamp_modificacion' es de
    hoy (fecha local), como lista de dict.

    El timestamp se guarda como texto 'dd/mm/yyyy HH:MM:SS', por lo que el
    filtro se hace en Python (no con funciones de fecha SQL) para evitar
    problemas de formato.
    """
    try:
        cur = conn.cursor()
        cur.execute(
            'SELECT * FROM staging_rutas WHERE "timestamp_modificacion" IS NOT NULL '
            'AND "timestamp_modificacion" != \'\''
        )
        cols = [d[0] for d in cur.description]
        hoy = date.today()
        resultado = []
        for r in cur.fetchall():
            fila = dict(zip(cols, r))
            ts = fila.get("timestamp_modificacion", "")
            try:
                fecha_ts = datetime.strptime(ts, "%d/%m/%Y %H:%M:%S").date()
            except (ValueError, TypeError):
                continue
            if fecha_ts == hoy:
                resultado.append(fila)
        return resultado
    except Error as e:
        print(f"Error consultando filas modificadas hoy: {e}")
        return []
