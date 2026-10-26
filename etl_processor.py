"""ETL Processor for RutaV3_2026 - Lee datos crudos de SQLite, aplica las
matemáticas base con PV y genera los 4 CSVs independientes."""

import pandas as pd
import numpy as np
import xlwings as xw
import sqlite3
import os
import re
import json
from datetime import datetime, date, timedelta

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
DB_FILE = os.path.join(BASE_DIR, 'staging_rutas.db')
OUTPUT_CSV = os.path.join(BASE_DIR, 'input_powerquery.csv')

# Ruta de exportación para el generador de archivos del sistema
EXPORT_BASE_DIR = os.path.join(BASE_DIR, "Exportaciones_Rutas")

# Archivo de configuración local (persiste la ruta del Excel Maestro)
SETTINGS_FILE = os.path.join(BASE_DIR, "config_ruta.json")

# ---------------------------------------------------------------------------
# Excel Maestro corporativo
# ---------------------------------------------------------------------------
# Nombre por defecto (se puede cambiar desde la UI y persiste en config_ruta.json)
EXCEL_MAESTRO_DEFAULT = "Extracción de Datos Agregados _ Formatos Para Alta de Ruta (ultimo)_2026FACV.xlsm"
EXCEL_SHEET_NAME = "BASE 2026"
EXCEL_TABLE_NAME = "Solicitudes"


def get_excel_maestro_path():
    """Devuelve la ruta configurada del Excel Maestro.

    Orden de prioridad:
      1. Ruta guardada en config_ruta.json (si el archivo existe).
      2. Ruta por defecto dentro de BASE_DIR.
    """
    try:
        if os.path.exists(SETTINGS_FILE):
            with open(SETTINGS_FILE, "r", encoding="utf-8") as f:
                data = json.load(f)
            path = data.get("excel_maestro_path")
            if path and os.path.exists(path):
                return path
    except Exception as e:
        print(f"AVISO: No se pudo leer {SETTINGS_FILE}: {e}")

    default = os.path.join(BASE_DIR, EXCEL_MAESTRO_DEFAULT)
    return default


def set_excel_maestro_path(path):
    """Guarda la ruta del Excel Maestro en config_ruta.json y la retorna."""
    if path:
        try:
            with open(SETTINGS_FILE, "w", encoding="utf-8") as f:
                json.dump({"excel_maestro_path": path}, f, ensure_ascii=False, indent=2)
        except Exception as e:
            print(f"AVISO: No se pudo guardar {SETTINGS_FILE}: {e}")
    return path


# Variable de conveniencia (se resuelve dinámicamente al momento de usarla).
# Se mantiene por compatibilidad; el código usa get_excel_maestro_path().
def _resolver_excel_path():
    return get_excel_maestro_path()

# Columnas de la tabla que contienen fórmulas y NO deben tocarse
EXCEL_FORMULA_COLUMNS = {1, 16, 19, 24, 26, 28, 29, 30, 35, 36, 42, 43, 44}

# Mapeo por posición/letra de columna (1-based) -> clave de calculate_row
EXCEL_DATA_COLUMNS = {
    2: 'B ORGVTA',
    3: 'C CEDIS',
    4: 'D NOMBRE CEDIS',
    5: 'E SF',
    6: 'F NOMBRE SF',
    7: 'G NO PLANTA / OBRA',
    8: 'H NOMBRE PLANTA / OBRA',
    9: 'I ARTICULO',
    10: 'J DESC / ARTÍCULO',
    11: 'K PV',
    12: 'L MODALIDAD',
    13: 'M UM VTA',
    14: 'N $ MP VTA (TN)',
    15: 'O $ FLETE VTA (TN)',
    17: 'Q MP VTA (M3)',
    18: 'R FLETE VTA (M3)',
    20: 'T UM COMP [Solicitud]',
    21: 'U MP COMP (M3)',
    22: 'V CA COMP (M3)',
    23: 'W FLETE COMP (M3)',
    25: 'Y MP COMP (TN)',
    27: 'AA FLETE COMP (TN)',
    31: 'AE TIPO',
    32: 'AF FECHA DE PROCESAMIENTO',
    33: 'AG TIPO CLIENTE',
    34: 'AH OBSERVACION',
    37: 'AK FECHA SOLICITUD',
    38: 'AL INCIDENCIA?',
    39: 'AM MOTIVO',
    40: 'AN MES',
    41: 'AO FECHA CONF',
}

# Claves de columnas que deben escribirse como fechas de Excel
EXCEL_DATE_KEYS = {
    'AF FECHA DE PROCESAMIENTO',
    'AK FECHA SOLICITUD',
    'AO FECHA CONF',
}


class ExcelBloqueadoError(Exception):
    """Se lanza cuando Excel está ocupado o en modo edición y no se puede escribir."""
    pass


def _es_error_bloqueo_excel(e: Exception) -> bool:
    """Detecta errores COM/xlwings típicos de un Excel en Modo Edición o bloqueado."""
    name = type(e).__name__.lower()
    msg = str(e).lower()
    if 'com_error' in name or 'comerror' in name:
        return True
    hints = (
        'edit',
        'busy',
        'retry',
        'call was rejected',
        'server threw an exception',
        '0x800ac472',  # RPC_E_SERVERCALL_RETRYLATER
        '0x80010001',  # RPC_E_CALL_REJECTED
        'rpc_e_',
        'no se puede',
        'modo edición',
        'en modo',
        'ocupado',
    )
    return any(h in msg for h in hints)


# ---------------------------------------------------------------------------
# Helpers de fecha (emular NETWORKDAYS.INTL con fin de semana sáb/dom)
# ---------------------------------------------------------------------------
def _parse_date(value):
    """Convierte un valor a date de forma robusta.

    Acepta datetime, date, pandas.Timestamp, numpy.datetime64 y cadenas en
    múltiples formatos (con o sin hora, con guiones, puntos o barras). Si no
    puede parsear, devuelve None (nunca lanza excepción). Además detecta
    NaT/NaN/None en cualquier punto y los convierte a None.
    """
    import pandas as pd

    if value is None:
        return None

    # Detectar NaT/NaN de pandas/numpy desde el inicio.
    try:
        if pd.isna(value):
            return None
    except (ValueError, TypeError):
        pass

    if isinstance(value, datetime):
        return value.date()
    if isinstance(value, date):
        return value

    # pandas Timestamp / numpy datetime64
    if hasattr(value, "to_pydatetime"):
        try:
            d = value.to_pydatetime().date()
        except Exception:
            d = None
        if d is not None:
            try:
                if pd.isna(d):
                    return None
            except (ValueError, TypeError):
                pass
            return d

    # NaT tiene .date() pero devuelve NaT (no lanza), por eso se valida el resultado.
    if hasattr(value, "date") and not isinstance(value, date):
        try:
            d = value.date()
        except Exception:
            d = None
        if d is not None:
            try:
                if pd.isna(d):
                    return None
            except (ValueError, TypeError):
                pass
            return d

    s = str(value).strip()
    if not s or s.lower() in ("nat", "nan", "none"):
        return None

    for fmt in (
        "%Y-%m-%d %H:%M:%S",
        "%d.%m.%Y",
        "%Y-%m-%d",
        "%d/%m/%Y",
        "%d-%m-%Y",
        "%d.%m.%Y %H:%M:%S",
        "%d/%m/%y",
        "%d.%m.%y",
    ):
        try:
            return datetime.strptime(s, fmt).date()
        except ValueError:
            continue

    # Último recurso: pandas (si está disponible) para formatos ambiguos.
    # Se valida explícitamente que no sea NaT antes de retornar.
    try:
        parsed = pd.to_datetime(s, errors="coerce")
        if parsed is not None and not pd.isna(parsed):
            return parsed.date()
    except Exception:
        pass

    return None


def _format_dd_mm_yyyy(value):
    """Formatea un valor de fecha a dd.mm.yyyy, dejando vacío si no es parseable."""
    d = _parse_date(value)
    if d is None:
        return ''
    return d.strftime("%d.%m.%Y")


def _networkdays(start, end):
    """Cuenta días hábiles (lun-vie) entre start y end, inclusive de ambos extremos."""
    if start is None or end is None:
        return None
    if end < start:
        return 0
    days = 0
    current = start
    while current <= end:
        if current.weekday() < 5:
            days += 1
        current += timedelta(days=1)
    return days


def calcular_estado_sla(fecha_procesamiento, fecha_confirmacion, hoy=None):
    """Calcula días hábiles, días transcurridos y estado de una solicitud.

    Réplica de Excel:
      - Días hábiles = DIAS.LAB.INTL(fecha_procesamiento, fecha_confirmacion, 1) - 1
        Equivalente en numpy: np.busday_count(fecha_procesamiento, fecha_confirmacion)
      - Días transcurridos = fecha_confirmacion - fecha_procesamiento (días calendario)

    Reglas de estado:
      - Si existe fecha de confirmación: CONFIRMADO.
      - Si días hábiles <= 4: VIGENTE.
      - Si días hábiles >= 5: VENCIDO.

    Blindado contra NaT/None/NaN de pandas: se valida con pd.isna() antes de
    cualquier resta matemática para evitar TypeError con NaTType.

    Retorna (dias_habiles, dias_transcurridos, estado).
    """
    import pandas as pd

    if hoy is None:
        hoy = date.today()

    # Cancelar de inmediato si la fecha de procesamiento es NaT/None/NaN/vacía.
    if fecha_procesamiento is None or pd.isna(fecha_procesamiento):
        return None, None, "SIN FECHA PROCESA"

    proc = _parse_date(fecha_procesamiento)
    if proc is None:
        return None, None, "SIN FECHA PROCESAMIENTO"

    conf = _parse_date(fecha_confirmacion)

    def _busday_count(start, end):
        try:
            return int(np.busday_count(start, end))
        except Exception:
            return None

    # NaT se evalúa como True en `if conf:`, por eso se valida explícitamente
    # con pd.isna() antes de restar.
    if conf is not None and not pd.isna(conf):
        dias_habiles = _busday_count(proc, conf)
        dias_transcurridos = (conf - proc).days
        estado = "CONFIRMADO"
    else:
        dias_habiles = _busday_count(proc, hoy)
        dias_transcurridos = (hoy - proc).days
        if dias_habiles is None:
            estado = "ERROR"
        else:
            estado = "VIGENTE" if dias_habiles <= 4 else "VENCIDO"

    return dias_habiles, dias_transcurridos, estado


# Posiciones (1-based) de las columnas de la tabla 'Solicitudes' en Excel
EXCEL_POS = {
    "cedis": 3,                 # C CEDIS
    "nombre_cedis": 4,          # D NOMBRE CEDIS
    "sf": 5,                    # E SF
    "nombre_sf": 6,             # F NOMBRE SF
    "planta_obra": 7,           # G NO PLANTA / OBRA
    "nombre_planta": 8,         # H NOMBRE PLANTA / OBRA
    "articulo": 9,              # I ARTICULO
    "desc_articulo": 10,        # J DESC / ARTÍCULO
    "pv": 11,                   # K PV
    "modalidad": 12,            # L MODALIDAD
    "um_vta": 13,               # M UM VTA
    "fecha_procesamiento": 32,  # AF FECHA DE PROCESAMIENTO
    "tipo_cliente": 33,         # AG TIPO CLIENTE
    "observacion": 34,          # AH OBSERVACION
    "fecha_solicitud": 37,      # AK FECHA SOLICITUD
    "mes": 40,                  # AN MES
    "fecha_conf": 41,           # AO FECHA CONF
}

# Columnas discriminadoras usadas en la búsqueda por ramas, en orden de
# prioridad definido por el usuario.
EXCEL_DISCRIMINADORES = [
    "fecha_procesamiento",
    "tipo_cliente",
    "fecha_solicitud",
    "um_vta",
]


def _norm_cell(v):
    """Normaliza una celda para comparación exacta (texto/numérico)."""
    if v is None:
        return ""
    if isinstance(v, (int, float)) and float(v).is_integer():
        return str(int(v))
    return str(v).strip()


def _norm_cell_fecha(v):
    """Normaliza una celda de fecha (dd/mm/yyyy o serial de Excel)."""
    if v is None or v == "":
        return ""
    if isinstance(v, (int, float)):
        try:
            return (datetime(1899, 12, 30) + timedelta(days=float(v))).strftime("%d/%m/%Y")
        except Exception:
            return _norm_cell(v)
    return _norm_cell(v)


def _norm_modalidad(v):
    """Normaliza modalidad: '01' -> '1', '1.0' -> '1', '04' -> '4'."""
    n = _norm_cell(v)
    if n == "":
        return ""
    try:
        return str(int(float(n)))
    except (ValueError, TypeError):
        return n


def _leer_filas_tabla_solicitudes():
    """Lee todas las filas de datos de la tabla 'Solicitudes' del Excel Maestro.

    Retorna (app, wb, sht, first_row, data) o (None, ...) en caso de error.
    `data` es una lista de filas (cada fila es una lista/tupla de celdas).
    """
    try:
        wb = xw.Book(get_excel_maestro_path())
        app = wb.app
        sht = wb.sheets[EXCEL_SHEET_NAME]

        lo = None
        for candidate in sht.api.ListObjects:
            if candidate.Name == EXCEL_TABLE_NAME:
                lo = candidate
                break
        if lo is None:
            print("Tabla 'Solicitudes' no encontrada en Excel.")
            return None, None, None, None, None

        # Quitar filtros activos: con AutoFilter activo, el rango .value solo
        # devuelve las filas visibles y la búsqueda/escritura falla o se pierde.
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

        data = sht.range((first_row + 1, 1), (last_row, last_col)).value
        if data is None:
            return app, wb, sht, first_row, []
        if last_row == first_row + 1:
            data = [data]
        return app, wb, sht, first_row, data
    except Exception as e:
        print(f"Error leyendo tabla 'Solicitudes' del Excel: {e}")
        return None, None, None, None, None


def _fila_a_candidato(offset, vals, first_row):
    """Convierte una fila cruda del Excel en un dict de candidato."""
    def cell(pos):
        return vals[pos - 1] if (vals is not None and len(vals) >= pos) else None

    return {
        "_excel_row": first_row + 1 + offset,
        "cedis": _norm_cell(cell(EXCEL_POS["cedis"])),
        "nombre_cedis": _norm_cell(cell(EXCEL_POS["nombre_cedis"])),
        "sf": _norm_cell(cell(EXCEL_POS["sf"])),
        "nombre_sf": _norm_cell(cell(EXCEL_POS["nombre_sf"])),
        "planta_obra": _norm_cell(cell(EXCEL_POS["planta_obra"])),
        "nombre_planta": _norm_cell(cell(EXCEL_POS["nombre_planta"])),
        "articulo": _norm_cell(cell(EXCEL_POS["articulo"])),
        "desc_articulo": _norm_cell(cell(EXCEL_POS["desc_articulo"])),
        "modalidad": _norm_modalidad(cell(EXCEL_POS["modalidad"])),
        "um_vta": _norm_cell(cell(EXCEL_POS["um_vta"])),
        "fecha_procesamiento": _norm_cell_fecha(cell(EXCEL_POS["fecha_procesamiento"])),
        "tipo_cliente": _norm_cell(cell(EXCEL_POS["tipo_cliente"])),
        "fecha_solicitud": _norm_cell_fecha(cell(EXCEL_POS["fecha_solicitud"])),
    }


def _coincide_candidato(cand, criterios):
    """Verifica si un candidato cumple todos los criterios (clave + discriminadores).

    `criterios` es un dict con las llaves normalizadas que deben coincidir.
    Los campos clave son obligatorios; los discriminadores solo se comparan si
    están presentes en el dict.
    """
    claves = ("cedis", "sf", "planta_obra", "articulo", "modalidad")
    for k in claves:
        if k in criterios and criterios[k] != "":
            if cand.get(k) != criterios[k]:
                return False
    for k in EXCEL_DISCRIMINADORES:
        if k in criterios and criterios[k] != "":
            if cand.get(k) != criterios[k]:
                return False
    return True


def buscar_candidatos_excel_sla(cedis, sf, planta_obra, articulo, modalidad,
                                fecha_procesamiento=None, tipo_cliente=None,
                                fecha_solicitud=None, um_vta=None):
    """Busca en el Excel Maestro las filas que coinciden con la clave compuesta
    y los discriminadores opcionales.

    Retorna una lista de dict (candidatos) con `_excel_row` y los campos clave.
    No escribe nada; solo lee. Deja el libro abierto.
    """
    app, wb, sht, first_row, data = _leer_filas_tabla_solicitudes()
    if data is None:
        return []
    try:
        criterios = {
            "cedis": _norm_cell(cedis),
            "sf": _norm_cell(sf),
            "planta_obra": _norm_cell(planta_obra),
            "articulo": _norm_cell(articulo),
            "modalidad": _norm_modalidad(modalidad),
            "fecha_procesamiento": _norm_cell_fecha(fecha_procesamiento) if fecha_procesamiento else "",
            "tipo_cliente": _norm_cell(tipo_cliente) if tipo_cliente else "",
            "fecha_solicitud": _norm_cell_fecha(fecha_solicitud) if fecha_solicitud else "",
            "um_vta": _norm_cell(um_vta) if um_vta else "",
        }

        resultados = []
        for offset, vals in enumerate(data):
            if vals is None:
                continue
            cand = _fila_a_candidato(offset, vals, first_row)
            if _coincide_candidato(cand, criterios):
                resultados.append(cand)
        return resultados
    finally:
        if app is not None:
            try:
                app.screen_updating = True
                app.calculation = "automatic"
            except Exception:
                pass


def actualizar_fecha_conf_excel_por_fila(num_fila_excel, fecha_conf):
    """Escribe `fecha_conf` en la columna 'AO FECHA CONF' de una fila específica
    del Excel Maestro (número absoluto de fila).

    Retorna el número de fila actualizado o None en caso de error.
    """
    if fecha_conf is None:
        fecha_conf = datetime.now().strftime("%d/%m/%Y")

    app = None
    wb = None
    try:
        wb = xw.Book(get_excel_maestro_path())
        app = wb.app
        sht = wb.sheets[EXCEL_SHEET_NAME]

        app.screen_updating = False
        app.calculation = "manual"

        sht.range((num_fila_excel, EXCEL_POS["fecha_conf"])).value = fecha_conf

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


def actualizar_fecha_conf_excel_sla(cedis, sf, planta_obra, articulo, modalidad, fecha_conf, fecha_procesamiento=None):
    """Actualiza la celda 'FECHA CONF' (columna AO) en el Excel Maestro para la
    solicitud identificada por su clave compuesta (+ fecha de procesamiento).

    Mantenida por compatibilidad: busca candidatos y escribe en todos. Para la
    búsqueda por ramas con preguntas usa `buscar_candidatos_excel_sla` y
    `actualizar_fecha_conf_excel_por_fila`.

    Retorna el número de celdas actualizadas o None si hubo error.
    """
    candidatos = buscar_candidatos_excel_sla(
        cedis, sf, planta_obra, articulo, modalidad,
        fecha_procesamiento=fecha_procesamiento,
    )
    if not candidatos:
        return 0

    actualizadas = 0
    for cand in candidatos:
        r = actualizar_fecha_conf_excel_por_fila(cand["_excel_row"], fecha_conf)
        if r is not None:
            actualizadas += 1
    return actualizadas


def _inyectar_en_excel_maestro(data):
    """Inserta filas calculadas en la tabla 'Solicitudes' del Excel Maestro.

    Función única y robusta para inyección en vivo. Recibe un DataFrame de
    pandas o una lista de diccionarios. Se engancha al archivo Excel ya
    abierto por el usuario (o lo abre visible) mediante `xw.Book`. Prepara la
    matriz de 44 columnas en memoria y la pega de golpe. Modo turbo: desactiva
    cálculo y actualización de pantalla durante la escritura.

    No cierra ni quita Excel; deja el libro abierto para que el usuario siga
    trabajando.
    """
    # Normalizar entrada a DataFrame
    if isinstance(data, list):
        df = pd.DataFrame(data)
    elif isinstance(data, pd.DataFrame):
        df = data
    else:
        raise TypeError("El argumento debe ser un DataFrame de pandas o una lista de diccionarios.")

    if df.empty:
        print("AVISO: No hay filas para insertar en el Excel Maestro.")
        return

    app = None
    wb = None
    try:
        # Live hooking: conectar al libro abierto o abrirlo visible
        wb = xw.Book(get_excel_maestro_path())
        app = wb.app

        if wb.api.ReadOnly:
            raise PermissionError(
                "El Excel Maestro está abierto como solo lectura. "
                "Ciérralo y ábrelo con permisos de escritura e intenta de nuevo."
            )

        sht = wb.sheets[EXCEL_SHEET_NAME]

        # Encontrar primera fila vacía de forma segura usando la columna B
        last_used_row = sht.cells.last_cell.row
        ultima_fila = sht.range(f"B{last_used_row}").end('up').row
        fila_destino = ultima_fila + 1

        n = len(df)
        print(f"[EXCEL MAESTRO DEBUG] Preparando inserción de {n} filas en fila {fila_destino}.")

        # Preparar matriz de inserción en memoria: 44 columnas por fila
        matriz_insercion = []
        for _, fila in df.iterrows():
            fila_excel = [None] * 44
            for col_idx, key in EXCEL_DATA_COLUMNS.items():
                value = fila.get(key)
                if value is None or (isinstance(value, float) and pd.isna(value)):
                    value = ""

                # Convertir fechas a datetime de Excel para no romper fórmulas dependientes
                if key in EXCEL_DATE_KEYS and value:
                    parsed = _parse_date(value)
                    if parsed:
                        value = datetime.combine(parsed, datetime.min.time())
                    else:
                        value = str(value)

                fila_excel[col_idx - 1] = value
            matriz_insercion.append(fila_excel)

        # Modo turbo: suspender cálculo y actualización de pantalla
        app.screen_updating = False
        app.calculation = 'manual'

        # Escritura en bloque de toda la matriz en la primera fila vacía
        sht.range(f"A{fila_destino}").value = matriz_insercion

        wb.save()
        print(f"[EXCEL MAESTRO] Insertadas {len(df)} filas en {get_excel_maestro_path()}")
    except Exception as e:
        # Intentar restaurar Excel antes de re-lanzar el error para no dejarlo en modo manual
        if app is not None:
            try:
                app.screen_updating = True
                app.calculation = 'automatic'
            except Exception:
                pass
        if _es_error_bloqueo_excel(e):
            raise ExcelBloqueadoError(
                "Excel está ocupado o en Modo Edición. Presiona ESC en Excel e intenta de nuevo"
            ) from e
        raise
    finally:
        # Restaurar siempre screen_updating y calculation, sin cerrar ni salir de Excel
        if app is not None:
            try:
                app.screen_updating = True
                app.calculation = 'automatic'
            except Exception as e:
                print(f"AVISO: No se pudieron restaurar los ajustes de Excel: {e}")
        # NO llamamos wb.close() ni app.quit() para dejar el archivo abierto


def _num(value, default=0.0):
    """Convierte a float de forma segura, tolerando formato de dinero.

    Acepta símbolos de moneda ($, MXN, USD), espacios y separadores de miles
    con coma (1,234.56) o punto (1.234,56). En caso de duda, interpreta el
    último separador como decimal y el resto como miles.
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
        # Si ya es un número limpio (p.ej. '123.45' o '123')
        return float(s)
    except ValueError:
        pass

    # Si tiene tanto coma como punto, el que esté más a la derecha es el
    # separador decimal; el otro es separador de miles.
    if "," in s and "." in s:
        if s.rfind(",") > s.rfind("."):
            # coma es decimal, punto es miles -> 1.234,56
            s = s.replace(".", "").replace(",", ".")
        else:
            # punto es decimal, coma es miles -> 1,234.56
            s = s.replace(",", "")
    elif "," in s and "." not in s:
        # Solo coma: si tiene exactamente 3 dígitos tras la coma, es miles
        # (1,234); en caso contrario, es decimal (123,45).
        if len(s.rsplit(",", 1)[1]) == 3 and s.count(",") == 1:
            s = s.replace(",", "")
        else:
            s = s.replace(",", ".")

    try:
        return float(s)
    except ValueError:
        return default


# ---------------------------------------------------------------------------
# Carga de catálogos SAP (simulados en CSV)
# ---------------------------------------------------------------------------
def load_sap_data():
    """Carga PVTA_MAT, PVTA_FTE y CONT_COMPRA desde CSV. Retorna dict de DataFrames."""
    files = {
        'PVTA_MAT VK13': 'PVTA_MAT VK13.csv',
        'PVTA_FTE VK13': 'PVTA_FTE VK13.csv',
        'CONT_COMPRA TRAOPE': 'CONT_COMPRA TRAOPE.csv',
    }
    sap = {}
    for key, filename in files.items():
        path = os.path.join(BASE_DIR, filename)
        try:
            df = pd.read_csv(path, dtype=str)
            # Normalizar nombres de columnas a minúsculas para matching robusto
            df.columns = [str(c).strip().lower() for c in df.columns]
            sap[key] = df
            print(f"OK: {filename} cargado ({len(df)} filas)")
        except FileNotFoundError:
            print(f"AVISO: {filename} no encontrado. Continuando sin él.")
            sap[key] = pd.DataFrame()
        except Exception as e:
            print(f"AVISO: Error cargando {filename}: {e}")
            sap[key] = pd.DataFrame()
    return sap


def _col(df, *names, default=None):
    """Busca una columna por varios nombres posibles (minúsculas)."""
    for n in names:
        n = n.lower()
        if n in df.columns:
            return n
    return default


# ---------------------------------------------------------------------------
# Cálculo de una fila (Columnas A - AR) según el PDF
# ---------------------------------------------------------------------------
def calculate_row(manual: dict, sap: dict = None):
    """Calcula todas las columnas A-AR a partir de los inputs manuales.

    Parámetros
    ----------
    manual : dict con llaves:
        org_vta, cedis, nombre_cedis, sf, nombre_sf, planta_obra, nombre_planta,
        articulo, desc_articulo, pv, modalidad, um_vta, mp_vta_m3, flete_vta_m3,
        integral_vta_m3, mp_comp_m3, ca_comp_m3, flete_comp_m3, integral_comp_m3,
        tipo, fecha_procesamiento, tipo_cliente, observacion, fecha_solicitud,
        incidencia, motivo, mes, fecha_conf, formato_destino

    sap : dict opcional con DataFrames cargados por load_sap_data().

    Retorna dict con TODAS las columnas A-AR (44 columnas) ya calculadas.
    """
    # --- Normalización de inputs manuales ---
    def g(key, default=""):
        v = manual.get(key, default)
        return v if v is not None else default

    cedis = str(g('cedis')).strip()
    sf = str(g('sf')).strip()
    planta = str(g('planta_obra')).strip()
    articulo = str(g('articulo')).strip()
    modalidad = str(g('modalidad')).strip()
    um_vta = str(g('um_vta')).strip()
    um_comp_sol = um_vta  # UM COMP [Solicitud] = UM VTA (según el modelo)

    # Los inputs de captura SIEMPRE están en M3.
    mp_vta_raw = _num(g('mp_vta_m3'))
    flete_vta_raw = _num(g('flete_vta_m3'))
    mp_comp_raw = _num(g('mp_comp_m3'))
    ca_comp_raw = _num(g('ca_comp_m3'))
    flete_comp_raw = _num(g('flete_comp_m3'))
    pv = _num(g('pv'))  # Peso volumétrico (factor TN <-> M3)

    # --- Columna A: RUTA ---
    ruta = "-".join([cedis, sf, planta, articulo, modalidad[-1] if modalidad else ""])

    # ------------------------------------------------------------------
    # Conversión M3 -> TN: siempre dividiendo entre PV
    # ------------------------------------------------------------------
    mp_vta_tn = (mp_vta_raw / pv) if pv else 0.0
    flete_vta_tn = (flete_vta_raw / pv) if pv else 0.0
    mp_comp_tn = (mp_comp_raw / pv) if pv else 0.0
    ca_comp_tn = (ca_comp_raw / pv) if pv else 0.0
    flete_comp_tn = (flete_comp_raw / pv) if pv else 0.0

    # Las variables de M3 conservan su valor capturado intacto
    mp_vta_m3_final = mp_vta_raw
    flete_vta_m3_final = flete_vta_raw
    mp_comp_m3_final = mp_comp_raw
    ca_comp_m3_final = ca_comp_raw
    flete_comp_m3_final = flete_comp_raw

    integral_vta_tn = mp_vta_tn + flete_vta_tn
    integral_vta_m3 = mp_vta_m3_final + flete_vta_m3_final

    integral_comp_m3 = mp_comp_m3_final + ca_comp_m3_final + flete_comp_m3_final
    integral_comp_tn = flete_comp_tn + mp_comp_tn + ca_comp_tn

    # ===== Márgenes operativos (dependen de la UM solicitada) =====
    if um_comp_sol.upper() == 'M3':
        mop_integral = integral_vta_m3 - integral_comp_m3
        mop_pct = (integral_vta_m3 - integral_comp_m3) / integral_vta_m3 if integral_vta_m3 else 0.0
        mop_m3 = mp_vta_m3_final - mp_comp_m3_final
        mop_pct2 = (mp_vta_m3_final - mp_comp_m3_final) / mp_vta_m3_final if mp_vta_m3_final else 0.0
    elif um_comp_sol.upper() == 'TN':
        mop_integral = integral_vta_tn - integral_comp_tn
        mop_pct = (integral_vta_tn - integral_comp_tn) / integral_vta_tn if integral_vta_tn else 0.0
        mop_m3 = mp_vta_tn - mp_comp_tn
        mop_pct2 = (mp_vta_tn - mp_comp_tn) / mp_vta_tn if mp_vta_tn else 0.0
    else:
        mop_integral = 0.0
        mop_pct = 0.0
        mop_m3 = 0.0
        mop_pct2 = 0.0

    # --- Columnas AP, AQ, AR (Días y Estado) ---
    fecha_proc = _parse_date(g('fecha_procesamiento'))
    fecha_conf = _parse_date(g('fecha_conf'))

    if fecha_proc and fecha_conf:
        dias_habiles = _networkdays(fecha_proc, fecha_conf) - 1 if fecha_conf >= fecha_proc else 0
        dias_transcurridos = (fecha_conf - fecha_proc).days
    else:
        dias_habiles = 0
        dias_transcurridos = 0

    # AR: Estado de solicitud
    if fecha_conf:
        estado = "CONFIRMADO"
    elif dias_transcurridos in (0, 1, 2, 3):
        estado = "VIGENTE"
    elif dias_transcurridos == 4:
        estado = "POR VENCER"
    else:
        estado = "VENCIDO"

    # --- Lookups Snowflake (AS - AZ) si hay SAP cargado ---
    resultado = {
        'A RUTA': ruta,
        'B ORGVTA': g('org_vta'),
        'C CEDIS': cedis,
        'D NOMBRE CEDIS': g('nombre_cedis'),
        'E SF': sf,
        'F NOMBRE SF': g('nombre_sf'),
        'G NO PLANTA / OBRA': planta,
        'H NOMBRE PLANTA / OBRA': g('nombre_planta'),
        'I ARTICULO': articulo,
        'J DESC / ARTÍCULO': g('desc_articulo'),
        'K PV': pv,
        'L MODALIDAD': modalidad,
        'M UM VTA': um_vta,
        'N $ MP VTA (TN)': round(mp_vta_tn, 4),
        'O $ FLETE VTA (TN)': round(flete_vta_tn, 4),
        'P $ INTEGRA VTA (TN)': round(integral_vta_tn, 4),
        'Q MP VTA (M3)': round(mp_vta_m3_final, 4),
        'R FLETE VTA (M3)': round(flete_vta_m3_final, 4),
        'S INTEGRAL VTA (M3)': round(integral_vta_m3, 4),
        'T UM COMP [Solicitud]': um_comp_sol,
        'U MP COMP (M3)': round(mp_comp_m3_final, 4),
        'V CA COMP (M3)': round(ca_comp_m3_final, 4),
        'W FLETE COMP (M3)': round(flete_comp_m3_final, 4),
        'X INTEGRAL COMP (M3)': round(integral_comp_m3, 4),
        'Y MP COMP (TN)': round(mp_comp_tn, 4),
        'Z CA COMPR (TN)': round(ca_comp_tn, 4),
        'AA FLETE COMP (TN)': round(flete_comp_tn, 4),
        'AB INTEGRAL COMP (TN)': round(integral_comp_tn, 4),
        'AC $ MOP INTEGRAL': round(mop_integral, 4),
        'AD % MOP INTEGRAL': round(mop_pct, 4),
        'AE TIPO': g('tipo'),
        'AF FECHA DE PROCESAMIENTO': g('fecha_procesamiento'),
        'AG TIPO CLIENTE': g('tipo_cliente'),
        'AH OBSERVACION': g('observacion'),
        'AI $ MOP/M3': round(mop_m3, 4),
        'AJ MOP %': round(mop_pct2, 4),
        'AK FECHA SOLICITUD': g('fecha_solicitud'),
        'AL INCIDENCIA?': g('incidencia'),
        'AM MOTIVO': g('motivo'),
        'AN MES': g('mes'),
        'AO FECHA CONF': g('fecha_conf'),
        'AP DÍAS HÁBILES': dias_habiles,
        'AQ DÍAS TRANSCURRIDOS': dias_transcurridos,
        'AR ESTADO DE SOLICITUD': estado,
        'FORMATO DESTINO': str(g('formato_destino')).strip(),
    }

    # --- Secciones 3-5 (AS - BO) con lookups SAP ---
    if sap is not None:
        resultado.update(_calculate_snowflake_columns(resultado, sap))

    return resultado


def _calculate_snowflake_columns(fila: dict, sap: dict) -> dict:
    """Calcula columnas AS-BO (validación Snowflake/SAP)."""
    out = {}

    ruta = fila['A RUTA']
    modalidad = str(fila['L MODALIDAD'])
    um_vta = str(fila['M UM VTA'])
    um_comp = um_vta  # UM COMP

    # --- AS, AT: MP VTA SNOW / UM MAT desde PVTA_MAT VK13 ---
    df_mat = sap.get('PVTA_MAT VK13', pd.DataFrame())
    if not df_mat.empty:
        col_ruta = _col(df_mat, 'ruta', 'rut', 'ruta_completa')
        col_importe = _col(df_mat, 'importe', 'precio', 'precio_neto')
        col_unidad = _col(df_mat, 'unidad', 'um', 'um_precio')
        if col_ruta and col_importe:
            match = df_mat[df_mat[col_ruta] == ruta]
            if not match.empty:
                out['AS MP VTA SNOV'] = _num(match.iloc[0].get(col_importe))
                out['AT UM MAT'] = str(match.iloc[0].get(col_unidad, '')) if col_unidad else 0
            else:
                out['AS MP VTA SNOV'] = 0.0
                out['AT UM MAT'] = 0
        else:
            out['AS MP VTA SNOV'] = 0.0
            out['AT UM MAT'] = 0
    else:
        out['AS MP VTA SNOV'] = 0.0
        out['AT UM MAT'] = 0

    # AU var MAT
    um_mat = out['AT UM MAT']
    if um_mat == 0:
        out['AU var MAT'] = 0
    elif str(um_mat).upper() == 'M3':
        out['AU var MAT'] = out['AS MP VTA SNOV'] - fila['Q MP VTA (M3)']
    else:
        out['AU var MAT'] = out['AS MP VTA SNOV'] - fila['N $ MP VTA (TN)']

    # AV var UMM
    out['AV var UMM'] = (str(um_mat) == str(um_vta))

    # --- AW, AX, AY, AZ: Flete Snowflake ---
    df_fte = sap.get('PVTA_FTE VK13', pd.DataFrame())
    if not df_fte.empty:
        col_ruta_f = _col(df_fte, 'ruta', 'rut', 'ruta_completa')
        col_importe_f = _col(df_fte, 'importe', 'precio', 'precio_neto')
        col_unidad_f = _col(df_fte, 'unidad', 'um', 'um_precio')
        if col_ruta_f and col_importe_f:
            match = df_fte[df_fte[col_ruta_f] == ruta]
            if not match.empty:
                out['AW FLETE VTA SNOV'] = _num(match.iloc[0].get(col_importe_f))
                out['AX UM FTE'] = str(match.iloc[0].get(col_unidad_f, '')) if col_unidad_f else 'RECOGIDO'
            else:
                out['AW FLETE VTA SNOV'] = 0.0
                out['AX UM FTE'] = 'RECOGIDO'
        else:
            out['AW FLETE VTA SNOV'] = 0.0
            out['AX UM FTE'] = 'RECOGIDO'
    else:
        out['AW FLETE VTA SNOV'] = 0.0
        out['AX UM FTE'] = 'RECOGIDO'

    tipo = str(fila['AE TIPO']).upper()
    um_fte = out['AX UM FTE']
    if tipo == 'CAMBIO':
        out['AY var FTE'] = 0
    elif um_fte == 0:
        out['AY var FTE'] = 0
    elif str(um_fte).upper() == 'M3':
        out['AY var FTE'] = out['AW FLETE VTA SNOV'] - fila['R FLETE VTA (M3)']
    else:
        out['AY var FTE'] = out['AW FLETE VTA SNOV'] - fila['O $ FLETE VTA (TN)']

    # AZ vs UMF
    out['AZ vs UMF'] = 'VERDADERO' if str(um_fte) == str(um_vta) else 'RECOGIDO'

    # --- BA, BB, BC, BD, BE: Compra Snowflake (TRAOPE) ---
    df_comp = sap.get('CONT_COMPRA TRAOPE', pd.DataFrame())
    last_digit = modalidad[-1] if modalidad else ''
    if last_digit != '4':
        comp_key = f"{fila['C CEDIS']}-{fila['E SF']}-{fila['I ARTICULO']}"
    else:
        comp_key = ruta

    if not df_comp.empty:
        col_ruta_c = _col(df_comp, 'ruta', 'rut', 'ruta_completa')
        col_precio_c = _col(df_comp, 'precio neto pedido', 'precio_neto', 'precio', 'importe')
        col_um_c = _col(df_comp, 'um precio pedido', 'um', 'unidad', 'um_compra')
        if col_ruta_c and col_precio_c:
            match = df_comp[df_comp[col_ruta_c] == comp_key]
            if not match.empty:
                out['BA MP COMP SNOV'] = _num(match.iloc[0].get(col_precio_c))
                out['BC UM COMP'] = str(match.iloc[0].get(col_um_c, '')) if col_um_c else 0
            else:
                out['BA MP COMP SNOV'] = 0.0
                out['BC UM COMP'] = 0
        else:
            out['BA MP COMP SNOV'] = 0.0
            out['BC UM COMP'] = 0
    else:
        out['BA MP COMP SNOV'] = 0.0
        out['BC UM COMP'] = 0

    # BB Validar Costo Compra vs Solicitud
    um_comp = out['BC UM COMP']
    mod2 = modalidad[-2:] if len(modalidad) >= 2 else modalidad
    if mod2 == '01':
        if um_comp == 0:
            out['BB Validar Costo Compra vs Solicitud'] = 0
        elif str(um_comp).upper() == 'M3':
            delta = abs(out['BA MP COMP SNOV'] - fila['X INTEGRAL COMP (M3)'])
            out['BB Validar Costo Compra vs Solicitud'] = 'ENTREGADO' if delta <= 0.1 else 'VERIFICAR'
        else:
            delta = abs(out['BA MP COMP SNOV'] - fila['AB INTEGRAL COMP (TN)'])
            out['BB Validar Costo Compra vs Solicitud'] = 'ENTREGADO' if delta <= 0.1 else 'VERIFICAR'
    elif mod2 == '04':
        if um_comp == 0:
            out['BB Validar Costo Compra vs Solicitud'] = 0
        elif str(um_comp).upper() == 'M3':
            delta = abs(out['BA MP COMP SNOV'] - fila['X INTEGRAL COMP (M3)'])
            out['BB Validar Costo Compra vs Solicitud'] = 'RECOGIDO' if delta <= 0.1 else 'VERIFICAR'
        else:
            delta = abs(out['BA MP COMP SNOV'] - fila['AB INTEGRAL COMP (TN)'])
            out['BB Validar Costo Compra vs Solicitud'] = 'RECOGIDO' if delta <= 0.1 else 'VERIFICAR'
    else:
        out['BB Validar Costo Compra vs Solicitud'] = ''

    # BD Variación Compra
    if mod2 == '01':
        out['BD Variación Compra'] = 'ENTREGADO'
    elif mod2 == '04':
        if um_comp == 0:
            out['BD Variación Compra'] = 0
        elif str(um_comp).upper() == 'M3':
            out['BD Variación Compra'] = out['BA MP COMP SNOV'] - fila['X INTEGRAL COMP (M3)']
        else:
            out['BD Variación Compra'] = out['BA MP COMP SNOV'] - fila['AB INTEGRAL COMP (TN)']
    else:
        out['BD Variación Compra'] = 'RECOGIDO'

    # BE Valida UM
    out['BE Valida UM'] = (str(out['BC UM COMP']) == str(um_comp))

    return out


# ---------------------------------------------------------------------------
# Migración de macros VBA a pandas: reportes CTO / FPA
# ---------------------------------------------------------------------------
def _get_col(df, col_name, default=None):
    """Devuelve la Serie de una columna por nombre (con o sin prefijo de letra).

    Si hay columnas duplicadas (p.ej. tras normalizar prefijos), devuelve la
    primera como Serie para evitar que pandas regrese un DataFrame.
    """
    if col_name in df.columns:
        serie = df[col_name]
        if isinstance(serie, pd.DataFrame):  # duplicados -> tomar primera
            serie = serie.iloc[:, 0]
        return serie
    # Buscar por terminación exacta (ej. 'RUTA' encuentra 'A RUTA' pero no 'NOMBRE RUTA')
    for c in df.columns:
        if str(c).upper() == str(col_name).upper():
            serie = df[c]
            if isinstance(serie, pd.DataFrame):
                serie = serie.iloc[:, 0]
            return serie
    return pd.Series([default] * len(df), index=df.index)


def _normalizar_columnas(df: pd.DataFrame) -> pd.DataFrame:
    """Traduce los nombres con prefijo de letra Excel ('A RUTA' -> 'RUTA') usando
    un mapa explícito de las columnas generadas por calculate_row(). No toca
    nombres legítimos como 'FORMATO DESTINO' o '$ MP VTA (TN)'."""
    # Mapa: nombre con prefijo -> nombre canónico (sin prefijo)
    PREFIJO_MAP = {
        "A RUTA": "RUTA",
        "B ORGVTA": "ORGVTA",
        "C CEDIS": "CEDIS",
        "D NOMBRE CEDIS": "NOMBRE CEDIS",
        "E SF": "SF",
        "F NOMBRE SF": "NOMBRE SF",
        "G NO PLANTA / OBRA": "NO PLANTA / OBRA",
        "H NOMBRE PLANTA / OBRA": "NOMBRE PLANTA / OBRA",
        "I ARTICULO": "ARTICULO",
        "J DESC / ARTÍCULO": "DESC / ARTÍCULO",
        "K PV": "PV",
        "L MODALIDAD": "MODALIDAD",
        "M UM VTA": "UM VTA",
        "N $ MP VTA (TN)": "$ MP VTA (TN)",
        "O $ FLETE VTA (TN)": "$ FLETE VTA (TN)",
        "P $ INTEGRA VTA (TN)": "$ INTEGRA VTA (TN)",
        "Q MP VTA (M3)": "MP VTA (M3)",
        "R FLETE VTA (M3)": "FLETE VTA (M3)",
        "S INTEGRAL VTA (M3)": "INTEGRAL VTA (M3)",
        "T UM COMP [Solicitud]": "UM COMP [Solicitud]",
        "U MP COMP (M3)": "MP COMP (M3)",
        "V CA COMP (M3)": "CA COMP (M3)",
        "W FLETE COMP (M3)": "FLETE COMP (M3)",
        "X INTEGRAL COMP (M3)": "INTEGRAL COMP (M3)",
        "Y MP COMP (TN)": "MP COMP (TN)",
        "Z CA COMPR (TN)": "CA COMPR (TN)",
        "AA FLETE COMP (TN)": "FLETE COMP (TN)",
        "AB INTEGRAL COMP (TN)": "INTEGRAL COMP (TN)",
        "AC $ MOP INTEGRAL": "$ MOP INTEGRAL",
        "AD % MOP INTEGRAL": "% MOP INTEGRAL",
        "AE TIPO": "TIPO",
        "AF FECHA DE PROCESAMIENTO": "FECHA DE PROCESAMIENTO",
        "AG TIPO CLIENTE": "TIPO CLIENTE",
        "AH OBSERVACION": "OBSERVACION",
        "AI $ MOP/M3": "$ MOP/M3",
        "AJ MOP %": "MOP %",
        "AK FECHA SOLICITUD": "FECHA SOLICITUD",
        "AL INCIDENCIA?": "INCIDENCIA?",
        "AM MOTIVO": "MOTIVO",
        "AN MES": "MES",
        "AO FECHA CONF": "FECHA CONF",
        "AP DÍAS HÁBILES": "DÍAS HÁBILES",
        "AQ DÍAS TRANSCURRIDOS": "DÍAS TRANSCURRIDOS",
        "AR ESTADO DE SOLICITUD": "ESTADO DE SOLICITUD",
        "AS MP VTA SNOV": "MP VTA SNOV",
        "AT UM MAT": "UM MAT",
        "AU var MAT": "var MAT",
        "AV var UMM": "var UMM",
        "AW FLETE VTA SNOV": "FLETE VTA SNOV",
        "AX UM FTE": "UM FTE",
        "AY var FTE": "var FTE",
        "AZ vs UMF": "vs UMF",
        "BA MP COMP SNOV": "MP COMP SNOV",
        "BB Validar Costo Compra vs Solicitud": "Validar Costo Compra vs Solicitud",
        "BC UM COMP": "UM COMP",
        "BD Variación Compra": "Variación Compra",
        "BE Valida UM": "Valida UM",
        "FORMATO DESTINO": "FORMATO DESTINO",
    }
    df = df.copy()
    df.columns = [PREFIJO_MAP.get(str(c).strip(), str(c).strip()) for c in df.columns]
    return df


# ---------------------------------------------------------------------------
# Construcción de reportes CTO/COMPRA, CTO/VENTA y FPA
# ---------------------------------------------------------------------------

def _build_cto_compra(base: pd.DataFrame) -> pd.DataFrame:
    """Devuelve DataFrame CTO COMPRA con los 33 encabezados oficiales."""
    cols = [
        'Contrato', 'VENDOR', 'PUR_ORG', 'COMPANY', 'PUR_GROUP', 'CURRENCY',
        'FCH_INI', 'FCH_FIN', 'PLANT', 'STOR_LOC', 'MATERIAL', 'QTY', 'PO_UM',
        'PRICE', 'PRIC_UNIT', 'PRIC_UM', 'TAX_CODE', 'ITEM_CAT', 'ACC_ASSIG', '',
        'COST_CENT', 'EXP_COND', 'KUNNR', 'PRICE', 'PRIC_UNIT', 'PRIC_UM',
        '¿Se requiere gasto adicional?', 'Condicion de precio', 'Importe',
        'Unidad de medida', 'Proveedor', 'Comentario', 'Posicion'
    ]

    rows = []
    for _, row in base.iterrows():
        um_val = str(row.get('UM COMP [Solicitud]', '')).strip().upper()
        modalidad = str(row.get('MODALIDAD', '')).strip()

        # KUNNR: solo si MODALIDAD termina en "4"
        kunnr = str(row.get('NO PLANTA / OBRA', '')).strip() if modalidad.endswith('4') else ''

        def costo_m3_tn(col_m3, col_tn):
            if um_val == 'TN':
                return row.get(col_tn, 0.0)
            return row.get(col_m3, 0.0)

        price_mat = costo_m3_tn('MP COMP (M3)', 'MP COMP (TN)')
        price_fte = costo_m3_tn('FLETE COMP (M3)', 'FLETE COMP (TN)')
        integral = costo_m3_tn('INTEGRAL COMP (M3)', 'INTEGRAL COMP (TN)')
        um_display = row.get('UM COMP [Solicitud]', '')

        record = [
            '',                          # Contrato
            row.get('SF', ''),           # VENDOR
            'MXPO',                      # PUR_ORG
            row.get('ORGVTA', ''),       # COMPANY
            'MXL',                       # PUR_GROUP
            'MXN',                       # CURRENCY
            row.get('FECHA DE PROCESAMIENTO', ''),  # FCH_INI
            '31.12.2030',                # FCH_FIN
            row.get('CEDIS', ''),        # PLANT
            "'001",                      # STOR_LOC
            row.get('ARTICULO', ''),     # MATERIAL
            '999999999',                 # QTY
            um_display,                  # PO_UM
            price_mat,                   # PRICE (material)
            '1',                         # PRIC_UNIT
            um_display,                   # PRIC_UM
            'V4',                        # TAX_CODE
            '',                          # ITEM_CAT (vacío)
            'K',                         # ACC_ASSIG
            '',                          # Columna sin encabezado
            '',                          # COST_CENT
            modalidad,                   # EXP_COND
            kunnr,                       # KUNNR
            price_fte,                   # PRICE (flete)
            '1',                         # PRIC_UNIT
            um_display,                   # PRIC_UM
            'SI' if price_fte else 'NO', # ¿Se requiere gasto adicional?
            '',                          # Condicion de precio
            integral,                    # Importe
            um_display,                   # Unidad de medida
            row.get('SF', ''),           # Proveedor
            row.get('OBSERVACION', ''),  # Comentario
            '',                          # Posicion (vacío)
        ]
        rows.append(record)

    return pd.DataFrame(rows, columns=cols)


def _build_cto_venta(base: pd.DataFrame) -> pd.DataFrame:
    """Devuelve DataFrame CTO VENTA con doble fila por registro (25 cols)."""
    cols = [
        'Contrato SAP (COC)', 'Clase de contrato', 'Org de ventas',
        'Canal de disT.', 'Sector', 'Cliente Nombre', 'Sucursal/obra', 'Frente',
        'Fecha de inicio de vigencia', 'Fecha de fin de vigencia', 'Moneda',
        'Cond Exp', 'Checkbox', 'Exp', 'Términos de pago', 'Checkbox',
        'Producto', 'Cantidad', 'Clase de condición', 'Motivo de pedido',
        'Denominación SHIPFROM', 'Texto a nivel cabecera', 'Centro', 'Checkbox', ''
    ]

    # Índices de columnas clave para la lógica de doble fila
    IDX_FCH_INI = 8
    IDX_FCH_FIN = 9
    IDX_FIRST_CHECKBOX = 12
    IDX_EXP = 13
    IDX_TERMINOS = 14
    IDX_MOTIVO = 19
    IDX_DENOM_SHIP = 20
    IDX_PRODUCTO = 16
    IDX_CANTIDAD = 17
    IDX_CENTRO = 22
    IDX_UM_EXTRA = 24

    rows = []
    for _, row in base.iterrows():
        fch_ini = _format_dd_mm_yyyy(row.get('FECHA DE PROCESAMIENTO', ''))
        fch_fin = '31.12.2029'
        org_vta = row.get('ORGVTA', '')
        modalidad = str(row.get('MODALIDAD', '')).strip()
        sucursal = row.get('NO PLANTA / OBRA', '')
        producto = row.get('ARTICULO', '')
        centro = row.get('CEDIS', '')
        sf_num = row.get('SF', '')
        observacion = row.get('OBSERVACION', '')
        um_vta = row.get('UM VTA', '')

        # Fila A: datos hasta el primer Checkbox; Motivo de pedido = 130
        fila_a = [''] * len(cols)
        fila_a[0] = ''                                  # Contrato SAP (COC)
        fila_a[1] = 'ZCP'                               # Clase de contrato
        fila_a[2] = org_vta                            # Org de ventas
        fila_a[3] = '00'                               # Canal de disT.
        fila_a[4] = '04'                               # Sector
        fila_a[5] = ''                                 # Cliente Nombre
        fila_a[6] = sucursal                           # Sucursal/obra
        fila_a[7] = ''                                 # Frente
        fila_a[8] = fch_ini                            # Fecha inicio
        fila_a[9] = fch_fin                            # Fecha fin
        fila_a[10] = 'MXN'                             # Moneda
        fila_a[11] = modalidad                         # Cond Exp
        fila_a[12] = ''                                # Checkbox
        fila_a[19] = '130'                             # Motivo de pedido
        rows.append(fila_a)

        # Fila B: vacío hasta primer Checkbox; llena a la derecha
        fila_b = [''] * len(cols)
        fila_b[IDX_EXP] = '01'                         # Exp
        fila_b[IDX_TERMINOS] = 'ZCON'                  # Términos de pago
        fila_b[IDX_PRODUCTO] = producto                # Producto
        fila_b[IDX_CANTIDAD] = '999999999'             # Cantidad
        fila_b[IDX_DENOM_SHIP] = sf_num                # Denominación SHIPFROM = SF
        fila_b[IDX_CENTRO] = centro                    # Centro
        fila_b[IDX_UM_EXTRA] = um_vta                  # UM VTA en columna extra
        rows.append(fila_b)

    return pd.DataFrame(rows, columns=cols)


def _translate_modalidad(modalidad: str) -> str:
    """Traduce el código de modalidad a su descripción."""
    m = str(modalidad).strip()
    mapping = {
        '01': '01 Entregado',
        '02': '02 Recogido',
        '04': '04 Entrega directa',
    }
    return mapping.get(m, m)


def _build_fpa(subset: pd.DataFrame, condicion_material: str, condicion_flete: str) -> pd.DataFrame:
    """Devuelve DataFrame FPA con los 24 encabezados oficiales."""
    cols = [
        'Solicitud', 'Esquema', 'Org.Ventas', 'ShipFrom',
        'Descripción de Ship From', 'Centro', 'Descripción de Centro',
        'Descripción Cliente', 'No. De Destino SAP', 'Descripción Destino',
        'Material', 'Descripción Material', 'Condición precio material', 'Precio Material', 'UM',
        'Condición precio flete', 'Precio Flete', 'Modalidad', 'Descripción Modalidad',
        'Valido desde:', 'Valido hasta:', 'MOP', '% MOP', 'PV'
    ]

    um = _get_col(subset, 'UM VTA', '').astype(str).str.strip().str.upper()

    # Selección según UM capturada
    precio_mat = _get_col(subset, 'MP VTA (M3)', 0.0).where(um != 'TN', _get_col(subset, '$ MP VTA (TN)', 0.0))
    precio_fte = _get_col(subset, 'FLETE VTA (M3)', 0.0).where(um != 'TN', _get_col(subset, '$ FLETE VTA (TN)', 0.0))

    # Esquema: deriva del TIPO CLIENTE capturado en Mesa de Captura.
    # TERCEROS -> 'Terceros' | FILIAL -> 'Concretos'
    tipo_cliente = _get_col(subset, 'TIPO CLIENTE', '').astype(str).str.strip().str.upper()
    esquema = tipo_cliente.map({'TERCEROS': 'Terceros', 'FILIAL': 'Concretos'}).fillna('')

    out = pd.DataFrame(index=subset.index)
    out['Solicitud'] = _get_col(subset, 'TIPO', '')
    out['Esquema'] = esquema
    out['Org.Ventas'] = _get_col(subset, 'ORGVTA', '')
    out['ShipFrom'] = _get_col(subset, 'SF', '')
    out['Descripción de Ship From'] = _get_col(subset, 'NOMBRE SF', '')
    out['Centro'] = _get_col(subset, 'CEDIS', '')
    out['Descripción de Centro'] = _get_col(subset, 'NOMBRE CEDIS', '')
    out['Descripción Cliente'] = ''
    out['No. De Destino SAP'] = _get_col(subset, 'NO PLANTA / OBRA', '')
    out['Descripción Destino'] = _get_col(subset, 'NOMBRE PLANTA / OBRA', '')
    out['Material'] = _get_col(subset, 'ARTICULO', '')
    out['Descripción Material'] = _get_col(subset, 'DESC / ARTÍCULO', '')
    out['Condición precio material'] = condicion_material
    out['Precio Material'] = precio_mat
    out['UM'] = _get_col(subset, 'UM VTA', '')
    out['Condición precio flete'] = condicion_flete
    out['Precio Flete'] = precio_fte
    out['Modalidad'] = _get_col(subset, 'MODALIDAD', '')
    out['Descripción Modalidad'] = _get_col(subset, 'MODALIDAD', '').apply(_translate_modalidad)
    out['Valido desde:'] = _get_col(subset, 'FECHA DE PROCESAMIENTO', '')
    out['Valido hasta:'] = '31.12.9999'
    out['MOP'] = _get_col(subset, '$ MOP INTEGRAL', 0.0)
    out['% MOP'] = _get_col(subset, '% MOP INTEGRAL', 0.0)
    out['PV'] = _get_col(subset, 'PV', '')

    return out[cols]


def generar_reportes_fpa_cto(df_validado: pd.DataFrame, output_dir: str = None, ruta_prefix: str = None) -> dict:
    """Enruta las filas según FORMATO DESTINO y exporta con los encabezados
    oficiales de las plantillas de la empresa en un único archivo Excel
    multi-hoja.

    Parámetros
    ----------
    df_validado : DataFrame con las columnas A-BE calculadas por calculate_row().
    output_dir  : directorio de salida (default: carpeta del proyecto).
    ruta_prefix : RUTA de la primera fila del lote; si se provee, se anexa al
                  nombre del archivo consolidado.

    Retorna dict con {nombre_archivo: ruta_completa}.
    """
    if output_dir is None:
        output_dir = BASE_DIR

    if df_validado is None or df_validado.empty:
        print("AVISO: generar_reportes_fpa_cto recibió un DataFrame vacío.")
        return {}

    # Normalizar nombres de columnas (quitar prefijos de letra Excel)
    df = _normalizar_columnas(df_validado)

    reportes = {}

    formato = _get_col(df, "FORMATO DESTINO").astype(str).str.strip().str.upper()
    mask_vta = formato.str.contains("CTO VENTA", na=False)
    mask_compra = formato.str.contains("CTO COMPRA", na=False)
    mask_fpa = formato.str.contains("FPA", na=False)

    # =====================================================================
    # 1) CTO VENTA: doble fila por registro
    # =====================================================================
    try:
        if mask_vta.any():
            cto_venta = _build_cto_venta(df[mask_vta])
            reportes["CTO_VENTA"] = cto_venta
        else:
            print("AVISO: Sin filas con FORMATO DESTINO = CTO VENTA.")
    except Exception as e:
        print(f"ERROR generando CTO VENTA: {e}")

    # =====================================================================
    # 2) CTO COMPRA: 33 encabezados oficiales
    # =====================================================================
    try:
        if mask_compra.any():
            cto_compra = _build_cto_compra(df[mask_compra])
            reportes["CTO_COMPRA"] = cto_compra
        else:
            print("AVISO: Sin filas con FORMATO DESTINO = CTO COMPRA.")
    except Exception as e:
        print(f"ERROR generando CTO COMPRA: {e}")

    # =====================================================================
    # 3) FPA: 24 encabezados oficiales (FILIAL y TERCEROS)
    # =====================================================================
    try:
        if mask_fpa.any():
            fpa = df[mask_fpa].copy()
            org_vta = _get_col(fpa, "ORGVTA", "").astype(str).str.strip()

            # FPA FILIAL: ORGVTA == 7100
            mask_filial = org_vta == "7100"
            if mask_filial.any():
                filial = _build_fpa(fpa[mask_filial], 'ZMAH', 'ZMPH')
                reportes["FPA_FILIAL"] = filial
            else:
                print("AVISO: Sin filas FPA con ORGVTA=7100.")

            # FPA TERCEROS: ORGVTA in (7180, 7277)
            mask_terceros = org_vta.isin(["7180", "7277"])
            if mask_terceros.any():
                terceros = _build_fpa(fpa[mask_terceros], 'ZMA6', 'ZMP1')
                reportes["FPA_TERCEROS"] = terceros
            else:
                print("AVISO: Sin filas FPA con ORGVTA 7180/7277.")
        else:
            print("AVISO: Sin filas con FORMATO DESTINO = FPA.")
    except Exception as e:
        print(f"ERROR generando FPA: {e}")

    salidas = {}
    if reportes:
        try:
            nombre_archivo = f"Reportes_Consolidados_{ruta_prefix}.xlsx" if ruta_prefix else "Reportes_Consolidados.xlsx"
            path = os.path.join(output_dir, nombre_archivo)
            with pd.ExcelWriter(path, engine='openpyxl') as writer:
                for sheet_name, df_reporte in reportes.items():
                    df_reporte.to_excel(writer, sheet_name=sheet_name, index=False)
            salidas[nombre_archivo] = path
            print(f"[EXPORTADO] {os.path.abspath(path)}  (sheets: {list(reportes.keys())})")
        except Exception as e:
            print(f"ERROR generando Reportes Consolidados: {e}")

    return salidas


# ---------------------------------------------------------------------------
# Procesamiento de Lotes (Batching)
# ---------------------------------------------------------------------------
def procesar_lote(rows: list, base_output_dir: str = None, sap_data: dict = None) -> str:
    """Procesa un lote de filas manuales, genera los reportes en una carpeta
    dinámica y abre el Explorador de Archivos en Windows.

    Parámetros
    ----------
    rows : list[dict]
        Lista de diccionarios con los campos manuales de cada fila.
    base_output_dir : str
        Carpeta raíz donde se creará Exportaciones_Rutas (default: EXPORT_BASE_DIR).
    sap_data : dict
        Catálogos SAP cargados; si es None se cargan automáticamente.

    Retorna
    -------
    str : ruta de la carpeta donde se generaron los archivos.
    """
    if not rows:
        print("AVISO: procesar_lote recibió un lote vacío.")
        return None

    if sap_data is None:
        sap_data = load_sap_data()

    # Calcular todas las filas
    calculated = []
    for manual in rows:
        calculated.append(calculate_row(manual, sap_data))

    result = pd.DataFrame(calculated)

    # Crear carpeta dinámica: Solicitud_<timestamp> dentro de la ruta base
    if base_output_dir is None:
        base_output_dir = EXPORT_BASE_DIR
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    folder = os.path.join(base_output_dir, f"Solicitud_{timestamp}")
    os.makedirs(folder, exist_ok=True)

    # RUTA de la primera fila para el prefijo del nombre de archivo
    ruta_primera = ''
    if calculated:
        ruta_primera = str(calculated[0].get('A RUTA', '')).strip()

    # Insertar lote en el Excel Maestro corporativo
    try:
        _inyectar_en_excel_maestro(result)
        print(f"[EXCEL MAESTRO] Insertadas {len(result)} filas en la tabla Solicitudes.")
    except PermissionError:
        raise
    except Exception as e:
        print("=" * 70)
        print("ERROR CRÍTICO: Fallo al inyectar datos en el Excel Maestro.")
        print(f"Archivo destino: {get_excel_maestro_path()}")
        import traceback
        traceback.print_exc()
        print(f"Excepción: {type(e).__name__}: {e}")
        print("=" * 70)
        raise

    # Generar reportes FPA/CTO en la misma carpeta
    try:
        generar_reportes_fpa_cto(result, output_dir=folder, ruta_prefix=ruta_primera)
    except Exception as e:
        print(f"AVISO: No se pudieron generar los reportes FPA/CTO del lote: {e}")

    # Abrir el Explorador de Archivos (solo Windows)
    try:
        if os.name == 'nt':
            os.startfile(os.path.abspath(folder))
            print(f"[EXPLORADOR] Abriendo carpeta: {os.path.abspath(folder)}")
    except Exception as e:
        print(f"AVISO: No se pudo abrir el Explorador de Archivos: {e}")

    return folder


# ---------------------------------------------------------------------------
# Procesamiento batch desde SQLite y exportación CSV
# ---------------------------------------------------------------------------
def process_and_validate_data(sqlite_db_path=DB_FILE, sap_data=None):
    """Lee staging de SQLite, calcula columnas y exporta input_powerquery.csv
    más los 4 reportes independientes."""
    conn = None
    try:
        conn = sqlite3.connect(sqlite_db_path)
        df = pd.read_sql_query("SELECT * FROM staging_rutas", conn)
    except Exception as e:
        print(f"Error leyendo SQLite: {e}")
        return pd.DataFrame()
    finally:
        if conn:
            conn.close()

    if df.empty:
        print("No hay registros en staging_rutas.")
        return pd.DataFrame()

    if sap_data is None:
        sap_data = load_sap_data()

    # Convertir cada fila de SQLite a dict manual y calcular
    column_map = {
        'org_vta': 'org_vta', 'cedis': 'cedis', 'nombre_cedis': 'nombre_cedis',
        'sf': 'sf', 'nombre_sf': 'nombre_sf', 'planta_obra': 'planta_obra',
        'nombre_planta': 'nombre_planta', 'articulo': 'articulo',
        'desc_articulo': 'desc_articulo', 'pv': 'pv', 'modalidad': 'modalidad',
        'um_vta': 'um_vta', 'mp_vta_m3': 'mp_vta_m3', 'flete_vta_m3': 'flete_vta_m3',
        'mp_comp_m3': 'mp_comp_m3', 'ca_comp_m3': 'ca_comp_m3',
        'flete_comp_m3': 'flete_comp_m3', 'tipo': 'tipo',
        'fecha_procesamiento': 'fecha_procesamiento', 'tipo_cliente': 'tipo_cliente',
        'observacion': 'observacion', 'fecha_solicitud': 'fecha_solicitud',
        'incidencia': 'incidencia', 'motivo': 'motivo', 'mes': 'mes',
        'fecha_conf': 'fecha_conf', 'formato_destino': 'formato_destino',
    }
    # Solo columnas que existen en el DataFrame
    valid_cols = [k for k in column_map if k in df.columns]

    rows = []
    for _, row in df.iterrows():
        manual = {k: row[k] for k in valid_cols}
        rows.append(calculate_row(manual, sap_data))

    result = pd.DataFrame(rows)

    # Insertar lote en el Excel Maestro corporativo
    try:
        _inyectar_en_excel_maestro(result)
        print(f"[EXCEL MAESTRO] Insertadas {len(result)} filas en la tabla Solicitudes.")
    except PermissionError:
        raise
    except Exception as e:
        print("=" * 70)
        print("ERROR CRÍTICO: Fallo al inyectar datos en el Excel Maestro.")
        print(f"Archivo destino: {get_excel_maestro_path()}")
        import traceback
        traceback.print_exc()
        print(f"Excepción: {type(e).__name__}: {e}")
        print("=" * 70)
        raise  # Re-lanzar para que la GUI muestre la alerta

    # Reportes FPA / CTO (migración de macros VBA)
    try:
        generar_reportes_fpa_cto(result)
    except Exception as e:
        print(f"AVISO: No se pudieron generar los reportes FPA/CTO: {e}")

    return result


if __name__ == '__main__':
    process_and_validate_data()
