"""Módulo backend para el Área de Validaciones (FPA).

Encapsula toda la lógica nueva del módulo FPA sin tocar la Mesa de Captura:
  - Pegado masivo de 24 columnas (identificación FPA).
  - Búsqueda de solicitudes en SQLite y en el Excel Maestro (xlwings).
  - Proyección a M3 (UM == TN -> multiplicar por PV, según confirmación).
  - Actualización de "Fecha de Confirmación" (fecha_conf) en SQLite y en Excel.
  - Captura de imágenes desde el portapapeles (Pillow ImageGrab).
  - Preparación del portapapeles mixto texto/imagen para Outlook (win32clipboard).
"""

import base64
import io
import os
from datetime import datetime

import pandas as pd
import xlwings as xw
from PIL import ImageGrab

import database_manager
import etl_processor

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
DB_FILE = os.path.join(BASE_DIR, "staging_rutas.db")
EXCEL_SHEET_NAME = "BASE 2026"
EXCEL_TABLE_NAME = "Solicitudes"


def _excel_maestro_path():
    """Devuelve la ruta configurada del Excel Maestro (compartida con etl_processor)."""
    return etl_processor.get_excel_maestro_path()

# ---------------------------------------------------------------------------
# Identificación FPA: las 24 columnas exactas del reporte FPA
# ---------------------------------------------------------------------------
FPA_COLUMNS = [
    "Solicitud", "Esquema", "Org.Ventas", "ShipFrom",
    "Descripción de Ship From", "Centro", "Descripción de Centro",
    "Descripción Cliente", "No. De Destino SAP", "Descripción Destino",
    "Material", "Descripción Material", "Condición precio material",
    "Precio Material", "UM", "Condición precio flete", "Precio Flete",
    "Modalidad", "Descripción Modalidad", "Valido desde:", "Valido hasta:",
    "MOP", "% MOP", "PV",
]

# Clave de identificación FPA -> campo de staging_rutas (para búsqueda SQLite)
FPA_ID_FIELDS = {
    "ShipFrom": "sf",
    "Centro": "cedis",
    "No. De Destino SAP": "planta_obra",
    "Material": "articulo",
    "Modalidad": "modalidad",
    "UM": "um_vta",
}

# Posición (1-based) de los campos clave dentro de la tabla 'Solicitudes'
# del Excel Maestro. Coinciden con EXCEL_DATA_COLUMNS de etl_processor.py.
EXCEL_ID_POS = {
    "sf": 5,          # E SF
    "cedis": 3,       # C CEDIS
    "planta_obra": 7, # G NO PLANTA / OBRA
    "articulo": 9,    # I ARTICULO
    "modalidad": 12,  # L MODALIDAD
}

# Columna AF "FECHA DE PROCESAMIENTO" en el Excel (posición 32).
EXCEL_FECHA_PROCESAMIENTO_POS = 32


def _s(value):
    """Convierte a string seguro (None -> '')."""
    if value is None:
        return ""
    return str(value).strip()


def _num(value, default=0.0):
    """Convierte a float de forma segura, tolerando formato de dinero.

    Acepta símbolos de moneda ($, MXN, USD), espacios y separadores de miles
    con coma (1,234.56) o punto (1.234,56). El último separador se interpreta
    como decimal y el resto como miles.
    """
    if value is None:
        return default
    s = str(value).strip()
    if s == "":
        return default

    # Quitar símbolos de moneda y espacios
    s = s.replace("$", "").replace("MXN", "").replace("USD", "").replace(" ", "")
    s = s.replace("mxn", "").replace("usd", "")
    if s == "":
        return default

    try:
        return float(s)
    except ValueError:
        pass

    # Si tiene coma y punto, el de más a la derecha es el decimal.
    if "," in s and "." in s:
        if s.rfind(",") > s.rfind("."):
            s = s.replace(".", "").replace(",", ".")
        else:
            s = s.replace(",", "")
    elif "," in s and "." not in s:
        # Solo coma: si va seguida de exactamente 3 dígitos, es miles.
        if len(s.rsplit(",", 1)[1]) == 3 and s.count(",") == 1:
            s = s.replace(",", "")
        else:
            s = s.replace(",", ".")

    try:
        return float(s)
    except ValueError:
        return default


def _norm_id(value):
    """Normaliza un identificador numérico/texto para comparación exacta.

    Excel devuelve códigos como 900006225.0, 1001.0, 1.0, mientras que el
    portapapeles trae '900006225', '1001', '01'. Convierte números enteros a
    su representación sin '.0' y quita espacios.
    """
    if value is None:
        return ""
    if isinstance(value, bool):
        return str(value)
    if isinstance(value, (int, float)):
        if float(value).is_integer():
            return str(int(value))
        return str(value)
    s = str(value).strip()
    if s == "":
        return ""
    # '900006225.0' -> '900006225'
    try:
        f = float(s)
        if f.is_integer():
            return str(int(f))
    except (ValueError, TypeError):
        pass
    return s


def _norm_modalidad(value):
    """Normaliza modalidad '01' -> '1', '1.0' -> '1', '04' -> '4'."""
    n = _norm_id(value)
    if n == "":
        return ""
    try:
        return str(int(float(n)))
    except (ValueError, TypeError):
        return n


def _norm_modalidad_db(value):
    """Normaliza modalidad a formato canónico de 2 dígitos: '01', '02', '04'.

    '1' -> '01', '1.0' -> '01', '4' -> '04', '04' -> '04'.
    """
    n = _norm_modalidad(value)
    if n == "":
        return ""
    try:
        return f"{int(float(n)):02d}"
    except (ValueError, TypeError):
        return n


def _norm_fecha(value):
    """Normaliza una fecha a dd/mm/yyyy (acepta datetime, date, float de Excel,
    y strings con separadores /, -, o . en dd/mm/yyyy o yyyy-mm-dd).

    Reemplaza puntos y guiones por barras antes de intentar parsear, y acepta
    años de 2 o 4 dígitos. Retorna '' si no se puede parsear.
    """
    if value is None or value == "":
        return ""
    if isinstance(value, (datetime,)):
        return value.strftime("%d/%m/%Y")
    if isinstance(value, str):
        s = value.strip()
        if not s:
            return ""
        # Estándarizar separadores: '.' y '-' -> '/'.
        for sep in (".", "-"):
            s = s.replace(sep, "/")
        s = s.replace(" ", "/")
        for fmt in ("%d/%m/%Y", "%Y/%m/%d", "%d/%m/%y"):
            try:
                return datetime.strptime(s, fmt).strftime("%d/%m/%Y")
            except ValueError:
                continue
        return s
    if isinstance(value, (int, float)):
        # Serial de fecha de Excel (p.ej. 45900.0 -> días desde 1899-12-30)
        try:
            from datetime import timedelta
            return (datetime(1899, 12, 30) + timedelta(days=float(value))).strftime("%d/%m/%Y")
        except Exception:
            return ""
    return _s(value)


# ---------------------------------------------------------------------------
# 1. Pegado masivo (24 columnas)
# ---------------------------------------------------------------------------
# Mapeo del formato real del cliente (hoja FPA con celdas combinadas y
# columnas ocultas). Al copiar una fila, Excel entrega 26 valores tabulados:
#   ind 0  = Solicitud (A:B combinada, valor en A)
#   ind 1  = '' (columna B, parte de la combinación A:B)
#   ind 2  = Esquema (C:E combinada, valor en C)
#   ind 3  = '' (columna D, parte de la combinación C:E)
#   ind 4  = Org.Ventas (F; la columna E está oculta y se omite)
#   ind 5  = ShipFrom (G)
#   ind 6  = Descripción de Ship From (H)
#   ind 7  = Centro (I)
#   ind 8  = Descripción de Centro (J)
#   ind 9  = Descripción Cliente (K)
#   ind 10 = No. De Destino SAP (L)
#   ind 11 = Descripción Destino (M)
#   ind 12 = Material (N)
#   ind 13 = Descripción Material (O)
#   ind 14 = Condición precio material (P)
#   ind 15 = Precio Material (Q)
#   ind 16 = UM (R)
#   ind 17 = Condición precio flete (S)
#   ind 18 = Precio Flete (T)
#   ind 19 = Modalidad (U)
#   ind 20 = Descripción Modalidad (V)
#   ind 21 = Valido desde: (W)
#   ind 22 = Valido hasta: (X)
#   ind 23 = MOP (AL; columnas Y:AK ocultas se omiten)
#   ind 24 = % MOP (AM)
#   ind 25 = PV (AN)
FPA_CLIPBOARD_26_MAP = {
    "Solicitud": 0,
    "Esquema": 2,
    "Org.Ventas": 4,
    "ShipFrom": 5,
    "Descripción de Ship From": 6,
    "Centro": 7,
    "Descripción de Centro": 8,
    "Descripción Cliente": 9,
    "No. De Destino SAP": 10,
    "Descripción Destino": 11,
    "Material": 12,
    "Descripción Material": 13,
    "Condición precio material": 14,
    "Precio Material": 15,
    "UM": 16,
    "Condición precio flete": 17,
    "Precio Flete": 18,
    "Modalidad": 19,
    "Descripción Modalidad": 20,
    "Valido desde:": 21,
    "Valido hasta:": 22,
    "MOP": 23,
    "% MOP": 24,
    "PV": 25,
}


def parse_clipboard_24(raw: str):
    """Parsea texto tabulado del portapapeles en filas FPA (24 columnas).

    Soporta dos formatos:
      - Formato plano de 24 columnas (mapeo 1:1 con FPA_COLUMNS).
      - Formato del cliente con celdas combinadas/ocultas (26 columnas),
        usando FPA_CLIPBOARD_26_MAP para alinear correctamente.

    Devuelve una lista de dict con llaves FPA_COLUMNS.
    """
    rows = []
    if not raw:
        return rows

    for line in raw.splitlines():
        if not line.strip():
            continue
        values = [v.strip() for v in line.split("\t")]

        # Si trae exactamente 26 columnas, es el formato combinado/oculto.
        if len(values) >= 26:
            fila = {}
            for fpa_key, idx in FPA_CLIPBOARD_26_MAP.items():
                fila[fpa_key] = values[idx] if idx < len(values) else ""
            # Por seguridad, completar cualquier llave FPA que no esté mapeada
            for fpa_key in FPA_COLUMNS:
                fila.setdefault(fpa_key, "")
            rows.append(fila)
        else:
            # Formato plano de 24 columnas (comportamiento original)
            padded = values[:24] + [""] * (24 - len(values))
            rows.append({FPA_COLUMNS[i]: padded[i] for i in range(24)})
    return rows


# Lista maestra de las 24 columnas FPA (orden canónico). Se usa para
# `reindex` en el pegado con encabezados: cualquier columna omitida por el
# usuario se rellena con vacío, garantizando que el orden no se rompa.
FPA_COLUMNS_MASTER = list(FPA_COLUMNS)


def _normalizar_header(name):
    """Normaliza un nombre de columna para compararlo de forma tolerante:
    minúsculas, sin acentos y sin espacios en blanco."""
    import unicodedata
    s = str(name).strip().lower()
    s = "".join(c for c in unicodedata.normalize("NFD", s) if unicodedata.category(c) != "Mn")
    return "".join(s.split())


def parse_clipboard_fpa(raw: str):
    """Parsea el portapapeles FPA de forma robusta (pegado con o sin encabezados).

    Estrategia híbrida:
      1. Intenta interpretar el texto como un CSV tabulado con encabezados
         (`pandas` + header=0). Si detecta encabezados que coinciden con las
         24 columnas maestras (ignorando mayúsculas/acentos/espacios), hace
         `df.reindex(columns=FPA_COLUMNS_MASTER)` para rellenar con vacío las
         columnas omitidas y respetar el orden canónico.
      2. Si falla el parseo o no se reconocen encabezados, cae al parser
         posicional existente (parse_clipboard_24), que soporta el formato de
         26 columnas con celdas combinadas/ocultas.

    Devuelve una lista de dict con llaves FPA_COLUMNS.
    """
    if not raw or not raw.strip():
        return []

    # --- Intento 1: formato con encabezados (pandas) ---
    try:
        df = pd.read_csv(io.StringIO(raw), sep="\t", header=0, dtype=str)
        if not df.empty and len(df.columns) > 0:
            header_norm = {_normalizar_header(c) for c in df.columns}
            master_norm = {_normalizar_header(c) for c in FPA_COLUMNS_MASTER}
            # Reconocer como "con encabezados" si al menos una columna coincide
            if header_norm & master_norm:
                # Mapa: header normalizado -> columna maestra canónica
                norm_to_canon = {_normalizar_header(c): c for c in FPA_COLUMNS_MASTER}
                rename = {c: norm_to_canon.get(_normalizar_header(c), c) for c in df.columns}
                df = df.rename(columns=rename)
                df = df.reindex(columns=FPA_COLUMNS_MASTER)
                df = df.fillna("")
                return df.to_dict("records")
    except Exception:
        # Cualquier error de pandas (columnas duplicadas, formato roto) cae al fallback.
        pass

    # --- Intento 2: fallback posicional (formato 26/24 columnas) ---
    return parse_clipboard_24(raw)


# ---------------------------------------------------------------------------
# 2. Búsqueda de solicitudes
# ---------------------------------------------------------------------------
def _where_criteria(row, fecha_procesamiento=None):
    """Construye (where_clause, params) a partir de una fila FPA usando los
    campos de identificación no vacíos, más la fecha de procesamiento opcional.

    La modalidad se compara de forma tolerante: acepta tanto '01' como '1'
    para que coincida con registros viejos y nuevos de la base de datos.
    """
    clauses = []
    params = []
    for fpa_key, db_key in FPA_ID_FIELDS.items():
        value = _s(row.get(fpa_key) or row.get(db_key) or (row.get("destino") if db_key == "planta_obra" else ""))
        if not value:
            continue
        if db_key == "modalidad":
            canon = _norm_modalidad_db(value)
            corto = _norm_modalidad(value)
            if canon != corto:
                clauses.append(f'("{db_key}" = ? OR "{db_key}" = ?)')
                params.extend([canon, corto])
            else:
                clauses.append(f'"{db_key}" = ?')
                params.append(canon)
        else:
            clauses.append(f'"{db_key}" = ?')
            params.append(value)
    fecha = _norm_fecha(fecha_procesamiento)
    if fecha:
        # La columna puede guardar dd/mm/yyyy, dd-mm-yyyy, dd.mm.yyyy o
        # yyyy/mm/dd (con o sin hora). Normalizamos los separadores en SQL con
        # REPLACE y comparamos contra dd/mm/yyyy y yyyy/mm/dd para cubrir ambos
        # órdenes, de modo que la búsqueda en SQLite sea tan tolerante como en
        # Excel.
        try:
            fecha_iso = datetime.strptime(fecha, "%d/%m/%Y").strftime("%Y/%m/%d")
        except ValueError:
            fecha_iso = fecha
        clauses.append(
            '(REPLACE(REPLACE("fecha_procesamiento", \'-\', \'/\'), \'.\', \'/\') = ? '
            'OR REPLACE(REPLACE("fecha_procesamiento", \'-\', \'/\'), \'.\', \'/\') = ?)'
        )
        params.extend([fecha, fecha_iso])
    return clauses, params


def buscar_en_sqlite(conn, row, fecha_procesamiento=None):
    """Busca solicitudes en SQLite que coincidan con los campos de ID de la fila
    y, opcionalmente, con la fecha de procesamiento.

    Retorna lista de dict con los campos manuales de staging_rutas + 'id'.
    """
    if conn is None:
        return []
    clauses, params = _where_criteria(row, fecha_procesamiento)
    if not clauses:
        return []

    where = " AND ".join(clauses)
    sql = f'SELECT * FROM staging_rutas WHERE {where} ORDER BY id DESC'
    cur = conn.cursor()
    cur.execute(sql, params)
    cols = [d[0] for d in cur.description]
    return [dict(zip(cols, r)) for r in cur.fetchall()]


def buscar_en_excel(row, fecha_procesamiento=None):
    """Busca solicitudes en el Excel Maestro (tabla 'Solicitudes') por clave
    compuesta: SF, CEDIS, NO PLANTA, ARTICULO, MODALIDAD y, opcionalmente,
    la Fecha de Procesamiento (columna AF).

    Retorna lista de dict con los campos FPA_COLUMNS reconstruidos desde las
    posiciones conocidas del Excel, más la fila absoluta donde se encontró.
    No abre un nuevo Excel: usa el libro ya abierto (live hooking) o lo abre
    visible. Deja el libro abierto.
    """
    app = None
    wb = None
    try:
        wb = xw.Book(_excel_maestro_path())
        app = wb.app
        sht = wb.sheets[EXCEL_SHEET_NAME]

        # Leer el rango de la tabla 'Solicitudes' completa (valores)
        lo = None
        for candidate in sht.api.ListObjects:
            if candidate.Name == EXCEL_TABLE_NAME:
                lo = candidate
                break
        if lo is None:
            return []

        # Quitar filtros activos antes de leer, igual que en etl_processor.
        try:
            sht.api.AutoFilterMode = False
        except Exception:
            pass
        try:
            lo.ShowAutoFilter = False
        except Exception:
            pass

        first_row = lo.Range.Row
        last_row = lo.Range.Row + lo.Range.Rows.Count - 1
        last_col = lo.Range.Column + lo.Range.Columns.Count - 1

        target = {
            "sf": _norm_id(row.get("ShipFrom") or row.get("sf")),
            "cedis": _norm_id(row.get("Centro") or row.get("cedis")),
            "planta_obra": _norm_id(row.get("No. De Destino SAP") or row.get("planta_obra") or row.get("destino")),
            "articulo": _norm_id(row.get("Material") or row.get("articulo")),
            "modalidad": _norm_modalidad(row.get("Modalidad") or row.get("modalidad")),
        }
        target_fecha = _norm_fecha(fecha_procesamiento)

        # Leer todo el bloque de datos de una sola vez (más rápido y robusto)
        data = sht.range((first_row + 1, 1), (last_row, last_col)).value
        if data is None:
            return []
        # Normalizar a lista de filas (xlwings devuelve lista o tupla según el tamaño)
        if last_row == first_row + 1:
            data = [data]

        results = []
        for offset, vals in enumerate(data):
            if vals is None:
                continue
            match = True
            for key, pos in EXCEL_ID_POS.items():
                if not target[key]:
                    continue
                cell_val = _norm_modalidad(vals[pos - 1]) if key == "modalidad" else _norm_id(vals[pos - 1])
                if cell_val != target[key]:
                    match = False
                    break
            if match and target_fecha:
                cell_fecha = _norm_fecha(vals[EXCEL_FECHA_PROCESAMIENTO_POS - 1] if len(vals) >= EXCEL_FECHA_PROCESAMIENTO_POS else None)
                if cell_fecha != target_fecha:
                    match = False
            if match:
                r = first_row + 1 + offset
                results.append({
                    "_excel_row": r,
                    "ShipFrom": _s(vals[EXCEL_ID_POS["sf"] - 1]),
                    "Centro": _s(vals[EXCEL_ID_POS["cedis"] - 1]),
                    "No. De Destino SAP": _s(vals[EXCEL_ID_POS["planta_obra"] - 1]),
                    "Material": _s(vals[EXCEL_ID_POS["articulo"] - 1]),
                    "Modalidad": _s(vals[EXCEL_ID_POS["modalidad"] - 1]),
                    "UM": _s(vals[12] if len(vals) >= 12 else ""),  # M UM VTA
                    "FechaProcesamiento": _s(vals[EXCEL_FECHA_PROCESAMIENTO_POS - 1] if len(vals) >= EXCEL_FECHA_PROCESAMIENTO_POS else ""),
                })
        return results
    except Exception as e:
        print(f"Error buscando en Excel: {e}")
        return []
    finally:
        if app is not None:
            try:
                app.screen_updating = True
                app.calculation = "automatic"
            except Exception:
                pass
        # No cerrar el libro; el usuario sigue trabajando


# ---------------------------------------------------------------------------
# 3. Proyección a M3 (solo lectura, no sobrescribe la data cruda)
# ---------------------------------------------------------------------------
def proyectar_m3(row):
    """Devuelve (precio_material_m3, precio_flete_m3) mostrando la conversión.

    Si UM == 'TN', los precios se multiplican por PV para mostrarlos en M3.
    Si UM == 'M3' (u otro), se muestran sin conversión. Nunca muta la fila
    original; solo produce los valores proyectados para mostrar.
    """
    um = _s(row.get("UM")).upper()
    precio_mat = _num(row.get("Precio Material"))
    precio_flete = _num(row.get("Precio Flete"))
    pv = _num(row.get("PV"))

    if um == "TN" and pv > 0:
        return precio_mat * pv, precio_flete * pv
    return precio_mat, precio_flete


# ---------------------------------------------------------------------------
# 3.1 Migración Excel -> SQLite (solo filas faltantes, sin duplicar)
# ---------------------------------------------------------------------------
# Mapa: posición de columna del Excel -> columna de staging_rutas. Usa los
# mismos números de EXCEL_DATA_COLUMNS de etl_processor.py.
EXCEL_TO_DB_COLUMNS = {
    2: "org_vta",         # B ORGVTA
    3: "cedis",           # C CEDIS
    4: "nombre_cedis",    # D NOMBRE CEDIS
    5: "sf",              # E SF
    6: "nombre_sf",       # F NOMBRE SF
    7: "planta_obra",     # G NO PLANTA / OBRA
    8: "nombre_planta",   # H NOMBRE PLANTA / OBRA
    9: "articulo",        # I ARTICULO
    10: "desc_articulo",   # J DESC / ARTÍCULO
    11: "pv",              # K PV
    12: "modalidad",       # L MODALIDAD
    13: "um_vta",          # M UM VTA
    17: "mp_vta_m3",       # Q MP VTA (M3)
    18: "flete_vta_m3",    # R FLETE VTA (M3)
    21: "mp_comp_m3",      # U MP COMP (M3)
    22: "ca_comp_m3",      # V CA COMP (M3)
    23: "flete_comp_m3",   # W FLETE COMP (M3)
    31: "tipo",            # AE TIPO
    32: "fecha_procesamiento",  # AF FECHA DE PROCESAMIENTO
    33: "tipo_cliente",    # AG TIPO CLIENTE
    34: "observacion",      # AH OBSERVACION
    37: "fecha_solicitud",  # AK FECHA SOLICITUD
    38: "incidencia",      # AL INCIDENCIA?
    39: "motivo",          # AM MOTIVO
    40: "mes",             # AN MES
    41: "fecha_conf",      # AO FECHA CONF
}


def migrar_excel_a_sqlite(conn):
    """Lee la tabla 'Solicitudes' del Excel Maestro y sincroniza con SQLite.

    Por cada fila del Excel:
      - Si la clave no existe, la INSERTA.
      - Si la clave ya existe, ACTUALIZA sus campos (upsert), de modo que
        cambios de fechas/precios en Excel se reflejen en SQLite.

    Clave de identidad: (cedis, sf, planta_obra, articulo, modalidad).
    Retorna (insertadas, actualizadas, error).
    """
    if conn is None:
        return 0, 0, "Sin conexión a SQLite."

    app = None
    wb = None
    try:
        wb = xw.Book(_excel_maestro_path())
        app = wb.app
        sht = wb.sheets[EXCEL_SHEET_NAME]

        lo = None
        for candidate in sht.api.ListObjects:
            if candidate.Name == EXCEL_TABLE_NAME:
                lo = candidate
                break
        if lo is None:
            return 0, 0, f"Tabla '{EXCEL_TABLE_NAME}' no encontrada en Excel."

        # Quitar filtros activos antes de leer la tabla completa.
        try:
            sht.api.AutoFilterMode = False
        except Exception:
            pass
        try:
            lo.ShowAutoFilter = False
        except Exception:
            pass

        first_row = lo.Range.Row
        last_row = lo.Range.Row + lo.Range.Rows.Count - 1
        last_col = lo.Range.Column + lo.Range.Columns.Count - 1

        data = sht.range((first_row + 1, 1), (last_row, last_col)).value
        if data is None:
            return 0, 0, "No hay datos en la tabla de Excel."
        if last_row == first_row + 1:
            data = [data]

        # Columnas actualizables (las que vienen del Excel), en el orden exacto
        # de TABLA_COLS para INSERT.
        cols = list(database_manager.TABLA_COLS)

        insertadas = 0
        actualizadas = 0
        cur = conn.cursor()

        for vals in data:
            if vals is None:
                continue
            # Construir registro mapeando posiciones de Excel a columnas DB
            registro = {}
            for pos, db_col in EXCEL_TO_DB_COLUMNS.items():
                if pos - 1 < len(vals):
                    registro[db_col] = _s(vals[pos - 1])

            # Normalizar fechas para almacenarlas como dd/mm/yyyy
            for fecha_col in ("fecha_procesamiento", "fecha_solicitud", "fecha_conf"):
                if fecha_col in registro:
                    registro[fecha_col] = _norm_fecha(registro[fecha_col])

            # Normalizar identificadores numéricos para no guardar '900006225.0'
            for id_col in ("cedis", "sf", "planta_obra", "articulo", "org_vta"):
                if id_col in registro:
                    registro[id_col] = _norm_id(registro[id_col])

            # Normalizar modalidad a formato canónico de 2 dígitos ('01', '04')
            if "modalidad" in registro:
                registro["modalidad"] = _norm_modalidad_db(registro["modalidad"])

            # Clave de identidad (sin fecha, para detectar actualizaciones)
            cedis = _norm_id(registro.get("cedis"))
            sf = _norm_id(registro.get("sf"))
            planta = _norm_id(registro.get("planta_obra"))
            articulo = _norm_id(registro.get("articulo"))
            modalidad = _norm_modalidad_db(registro.get("modalidad"))

            # Filas completamente vacías
            if not any([cedis, sf, planta, articulo]):
                continue

            # Buscar registro existente por clave
            cur.execute(
                """SELECT id FROM staging_rutas
                   WHERE cedis=? AND sf=? AND planta_obra=? AND articulo=?
                     AND modalidad=?
                   LIMIT 1""",
                (cedis, sf, planta, articulo, modalidad),
            )
            existente = cur.fetchone()

            if existente is None:
                # Insertar fila nueva
                valores = tuple(registro.get(c, "") for c in cols)
                try:
                    database_manager.insert_data(conn, valores)
                    insertadas += 1
                except Exception as e:
                    print(f"Error insertando fila de Excel en SQLite: {e}")
            else:
                # Actualizar campos (upsert) excluyendo la clave
                row_id = existente[0]
                registro["timestamp_modificacion"] = datetime.now().strftime("%d/%m/%Y %H:%M:%S")
                set_cols = [c for c in registro.keys() if c in cols]
                if set_cols:
                    assignments = ", ".join(f'"{c}"=?' for c in set_cols)
                    params = [registro[c] for c in set_cols] + [row_id]
                    try:
                        cur.execute(
                            f"UPDATE staging_rutas SET {assignments} WHERE id=?",
                            params,
                        )
                        actualizadas += 1
                    except Exception as e:
                        print(f"Error actualizando fila {row_id} en SQLite: {e}")

        conn.commit()
        return insertadas, actualizadas, None
    except Exception as e:
        print(f"Error migrando Excel a SQLite: {e}")
        return 0, 0, str(e)
    finally:
        if app is not None:
            try:
                app.screen_updating = True
                app.calculation = "automatic"
            except Exception:
                pass


# ---------------------------------------------------------------------------
# 4. Actualización de "Fecha de Confirmación" (fecha_conf)
# ---------------------------------------------------------------------------
# En SQLite la columna ya existe como `fecha_conf`. En el Excel Maestro la
# columna 'AO FECHA CONF' es la posición 41 (1-based), igual que en
# etl_processor.EXCEL_DATA_COLUMNS.
EXCEL_FECHA_CONF_POS = 41  # AO


def _as_row_list(rows):
    """Normaliza la entrada a una lista de diccionarios (una fila o varias)."""
    if rows is None:
        return []
    if isinstance(rows, dict):
        return [rows]
    return list(rows)


def actualizar_fecha_confirmacion_sqlite(conn, rows, fecha=None):
    """Actualiza la Fecha de Confirmación de todas las filas de la solicitud.

    Para cada fila, busca sus coincidencias en SQLite (por clave compuesta) y
    actualiza `fecha_conf`. Retorna la lista de IDs de staging_rutas
    actualizados (vacía si no hubo coincidencias o error).
    """
    if conn is None:
        return []
    if fecha is None:
        fecha = datetime.now().strftime("%d/%m/%Y")

    ids_actualizados = []
    for row in _as_row_list(rows):
        matches = buscar_en_sqlite(conn, row)
        if not matches:
            continue
        try:
            cur = conn.cursor()
            ts = datetime.now().strftime("%d/%m/%Y %H:%M:%S")
            for m in matches:
                cur.execute(
                    'UPDATE staging_rutas SET fecha_conf=?, "timestamp_modificacion"=? WHERE id=?',
                    (fecha, ts, m["id"]),
                )
                ids_actualizados.append(m["id"])
            conn.commit()
        except Exception as e:
            print(f"Error actualizando fecha_conf en SQLite: {e}")
    return ids_actualizados


def actualizar_fecha_confirmacion_excel(rows, fecha=None):
    """Actualiza la 'Fecha de Confirmación' (columna AO) en el Excel Maestro
    para todas las filas de la solicitud activa.

    Usa live hooking (xw.Book), una sola conexión para todas las filas y deja
    el libro abierto. Retorna una lista con los números de fila absolutos de
    Excel que fueron actualizados.
    """
    filas = _as_row_list(rows)
    if not filas:
        return []
    if fecha is None:
        fecha = datetime.now().strftime("%d/%m/%Y")

    # Conjunto de claves normalizadas a buscar (evita duplicados)
    claves = []
    for row in filas:
        clave = {
            "sf": _norm_id(row.get("ShipFrom") or row.get("sf")),
            "cedis": _norm_id(row.get("Centro") or row.get("cedis")),
            "planta_obra": _norm_id(row.get("No. De Destino SAP") or row.get("planta_obra") or row.get("destino")),
            "articulo": _norm_id(row.get("Material") or row.get("articulo")),
            "modalidad": _norm_modalidad(row.get("Modalidad") or row.get("modalidad")),
        }
        if clave not in claves:
            claves.append(clave)

    app = None
    wb = None
    filas_actualizadas = []
    try:
        wb = xw.Book(_excel_maestro_path())
        app = wb.app
        sht = wb.sheets[EXCEL_SHEET_NAME]

        lo = None
        for candidate in sht.api.ListObjects:
            if candidate.Name == EXCEL_TABLE_NAME:
                lo = candidate
                break
        if lo is None:
            print("Tabla 'Solicitudes' no encontrada en Excel.")
            return []

        # Quitar filtros activos antes de leer/escribir.
        try:
            sht.api.AutoFilterMode = False
        except Exception:
            pass
        try:
            lo.ShowAutoFilter = False
        except Exception:
            pass

        first_row = lo.Range.Row
        last_row = lo.Range.Row + lo.Range.Rows.Count - 1
        last_col = lo.Range.Column + lo.Range.Columns.Count - 1

        app.screen_updating = False
        app.calculation = "manual"

        # Leer el bloque de datos una sola vez
        data = sht.range((first_row + 1, 1), (last_row, last_col)).value
        if data is None:
            return []
        if last_row == first_row + 1:
            data = [data]

        for offset, vals in enumerate(data):
            if vals is None:
                continue
            for clave in claves:
                match = True
                for key, pos in EXCEL_ID_POS.items():
                    if not clave[key]:
                        continue
                    cell_val = _norm_modalidad(vals[pos - 1]) if key == "modalidad" else _norm_id(vals[pos - 1])
                    if cell_val != clave[key]:
                        match = False
                        break
                if match:
                    # Columna AO = posición 41
                    num_fila = first_row + 1 + offset
                    sht.range((num_fila, EXCEL_FECHA_CONF_POS)).value = fecha
                    filas_actualizadas.append(num_fila)
                    break

        wb.save()
        return filas_actualizadas
    except Exception as e:
        print(f"Error actualizando fecha_conf en Excel: {e}")
        return filas_actualizadas
    finally:
        if app is not None:
            try:
                app.screen_updating = True
                app.calculation = "automatic"
            except Exception:
                pass


def actualizar_fecha_conf_excel_por_fila(num_fila_excel, fecha):
    """Actualiza exclusivamente la fila `num_fila_excel` (número absoluto de
    fila de Excel) en la columna 'FECHA CONF' (AO, posición 41).

    Usado cuando hay duplicados y el usuario ya eligió qué fila afectar.
    Retorna el número de fila actualizado (int) o None en caso de error.
    """
    if fecha is None:
        fecha = datetime.now().strftime("%d/%m/%Y")

    app = None
    wb = None
    try:
        wb = xw.Book(_excel_maestro_path())
        app = wb.app
        sht = wb.sheets[EXCEL_SHEET_NAME]

        app.screen_updating = False
        app.calculation = "manual"

        # Escribir directamente en la posición 41 (AO) de la fila elegida
        sht.range((num_fila_excel, EXCEL_FECHA_CONF_POS)).value = fecha

        wb.save()
        return num_fila_excel
    except Exception as e:
        print(f"Error actualizando FECHA CONF por fila {num_fila_excel}: {e}")
        return None
    finally:
        if app is not None:
            try:
                app.screen_updating = True
                app.calculation = "automatic"
            except Exception:
                pass


# ---------------------------------------------------------------------------
# 5. Evidencias FPA (imágenes desde el portapapeles)
# ---------------------------------------------------------------------------
def capturar_imagen_portapapeles():
    """Captura la imagen actual del portapapeles con Pillow ImageGrab.

    Retorna la imagen PIL o None si no hay imagen disponible.
    """
    try:
        img = ImageGrab.grabclipboard()
        if img is None:
            return None
        if isinstance(img, list):
            # Portapapeles con archivos; devolver None (solo se soporta imagen directa)
            return None
        return img
    except Exception as e:
        print(f"Error capturando imagen del portapapeles: {e}")
        return None


def _pil_to_png_base64(img):
    """Serializa una imagen PIL a base64 PNG."""
    buffer = io.BytesIO()
    img.convert("RGB").save(buffer, format="PNG")
    return base64.b64encode(buffer.getvalue()).decode("ascii")


def preparar_portapapeles_outlook(texto, imagenes):
    """Prepara el portapapeles de Windows con contenido mixto (texto + imágenes)
    en formato HTML listo para pegar en Outlook.

    texto    : cadena, p. ej. 'Ok 27/08/2026'
    imagenes : lista de objetos PIL.Image

    Retorna True si el portapapeles se escribió correctamente.
    """
    try:
        import win32clipboard
    except ImportError:
        print("win32clipboard no disponible. Instala pywin32.")
        return False

    if not imagenes:
        return _set_text_clipboard(texto)

    body = "".join(
        f'<img src="data:image/png;base64,{_pil_to_png_base64(img)}"><br>'
        for img in imagenes
    )

    try:
        win32clipboard.OpenClipboard()
        win32clipboard.EmptyClipboard()
        # CF_HTML con marcadores de fragmento obligatorios de Outlook.
        # Los offsets se calculan dinámicamente para que sean exactos.
        header = (
            "Version:0.9\r\n"
            "StartHTML:0000000000\r\n"
            "EndHTML:0000000000\r\n"
            "StartFragment:0000000000\r\n"
            "EndFragment:0000000000\r\n"
        )
        fragment = f"{texto}<br>{body}"
        doc = "<html><body><!--StartFragment-->" + fragment + "<!--EndFragment--></body></html>"

        start_html = len(header.encode("utf-8"))
        start_fragment = start_html + len("<html><body><!--StartFragment-->".encode("utf-8"))
        end_fragment = start_fragment + len(fragment.encode("utf-8"))
        end_html = start_html + len(doc.encode("utf-8"))

        header = (
            "Version:0.9\r\n"
            f"StartHTML:{start_html:010d}\r\n"
            f"EndHTML:{end_html:010d}\r\n"
            f"StartFragment:{start_fragment:010d}\r\n"
            f"EndFragment:{end_fragment:010d}\r\n"
        )
        cf_html = header + doc

        CF_HTML = win32clipboard.RegisterClipboardFormat("HTML Format")
        win32clipboard.SetClipboardData(CF_HTML, cf_html.encode("utf-8"))
        win32clipboard.CloseClipboard()
        return True
    except Exception as e:
        print(f"Error escribiendo portapapeles HTML: {e}")
        try:
            win32clipboard.CloseClipboard()
        except Exception:
            pass
        return False


def _set_text_clipboard(texto):
    """Escribe texto plano en el portapapeles (fallback cuando no hay imágenes)."""
    try:
        import win32clipboard
        win32clipboard.OpenClipboard()
        win32clipboard.EmptyClipboard()
        win32clipboard.SetClipboardText(texto)
        win32clipboard.CloseClipboard()
        return True
    except Exception as e:
        print(f"Error escribiendo portapapeles de texto: {e}")
        return False
