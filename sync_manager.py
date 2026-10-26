"""Sincronización manual por botones para trabajo sin servidor.

Cada usuario tiene su copia local (SQLite + Excel maestro). Una carpeta
"drive" (OneDrive/SharePoint o cualquier ruta configurable) actúa como punto
de intercambio:

  - Subir al drive  : copia la BD y el Excel locales hacia la carpeta drive.
  - Cargar datos    : copia la BD y el Excel del drive hacia la copia local,
                      sobrescribiendo la versión actual.

La carpeta drive se configura por equipo en sync_config.json (ruta absoluta),
porque OneDrive usa rutas distintas por usuario.
"""

import ctypes
import json
import os
import shutil
import subprocess
import sys
import time
from datetime import datetime

import database_manager
import etl_processor
import fpa_validations

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
CONFIG_FILE = os.path.join(BASE_DIR, "sync_config.json")
DB_FILE = os.path.join(BASE_DIR, "staging_rutas.db")

# Nombre del archivo de Excel maestro se resuelve dinámicamente desde
# etl_processor (única fuente de verdad), para que el selector de la UI
# y la sincronización apunten siempre al mismo archivo.

# Huella (mtime) del drive al momento de la última descarga por este equipo,
# por tipo de archivo ("db"/"excel"). Se usa como semáforo de colisiones:
# si al intentar subir el mtime actual del drive es mayor que esta huella,
# significa que otro usuario subió una versión más nueva mientras trabajábamos.
_last_known_cloud_mtime = {}


def hay_conexion_internet():
    """Detecta si hay conexión a internet usando la API de Windows (wininet).

    Retorna True por defecto si la llamada falla (p. ej. entorno no Windows),
    para no bloquear la sincronización en plataformas donde ctypes no aplica.
    """
    try:
        return ctypes.windll.wininet.InternetGetConnectedState(None, 0) == 1
    except Exception:
        return True


def get_excel_file():
    """Devuelve la ruta actual del Excel Maestro (configurable en la UI)."""
    return etl_processor.get_excel_maestro_path()


def get_sync_files():
    """Construye la lista de archivos a sincronizar en cada operación.

    Se evalúa en cada llamada para reflejar cambios de ruta del Excel Maestro
    hechos en tiempo de ejecución.
    """
    excel = get_excel_file()
    return [
        ("db", DB_FILE, "staging_rutas.db"),
        ("excel", excel, os.path.basename(excel)),
    ]


def _one_drive_default():
    """Intenta detectar la carpeta OneDrive/SharePoint local del usuario."""
    # Windows define variables de entorno con la ruta OneDrive.
    for var in ("OneDriveCommercial", "OneDriveConsumer", "OneDrive"):
        val = os.environ.get(var, "").strip()
        if val and os.path.isdir(val):
            return val

    candidates = []
    home = os.path.expanduser("~")
    try:
        for name in os.listdir(home):
            low = name.lower()
            if "onedrive" in low or "sharepoint" in low:
                candidates.append(os.path.join(home, name))
    except Exception:
        pass
    if candidates:
        return candidates[0]
    return ""


def load_config():
    """Carga sync_config.json o devuelve un dict con la ruta drive por defecto.

    Acepta dos formatos:
      - JSON: {"drive_dir": "C:\\ruta"}
      - Texto plano: la ruta en una sola línea (más seguro para rutas Windows).
    """
    if os.path.exists(CONFIG_FILE):
        raw = None
        try:
            with open(CONFIG_FILE, "r", encoding="utf-8-sig") as f:
                raw = f.read()
        except Exception:
            pass

        if raw:
            raw = raw.strip()
            # Primero intentar JSON
            try:
                cfg = json.loads(raw)
                if isinstance(cfg, dict):
                    return cfg
            except Exception:
                pass
            # Si no es JSON, asumir que la ruta está en texto plano
            if raw and not raw.startswith("{"):
                return {"drive_dir": raw}
    return {"drive_dir": _one_drive_default()}


def save_config(config):
    """Guarda el diccionario de configuración en sync_config.json."""
    try:
        with open(CONFIG_FILE, "w", encoding="utf-8") as f:
            json.dump(config, f, ensure_ascii=False, indent=2)
        return True
    except Exception as e:
        print(f"Error guardando sync_config.json: {e}")
        return False


def get_drive_dir():
    config = load_config()
    drive = config.get("drive_dir", "").strip()
    if not drive:
        # Fallback: subcarpeta local "Drive_Compartido" dentro de la app.
        drive = os.path.join(BASE_DIR, "Drive_Compartido")
    return drive


def _timestamped_backup(path):
    """Crea una copia de seguridad del archivo antes de sobrescribirlo.

    Los backups se guardan en una subcarpeta `backups` dentro de la carpeta
    de la aplicación (BASE_DIR/backups), para no ensuciar la carpeta de datos.
    """
    if not os.path.exists(path):
        return None
    backups_dir = os.path.join(BASE_DIR, "backups")
    try:
        os.makedirs(backups_dir, exist_ok=True)
    except Exception as e:
        print(f"AVISO: no se pudo crear carpeta de backups: {e}")
        backups_dir = BASE_DIR

    stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    backup = os.path.join(backups_dir, f"{os.path.basename(path)}.bak_{stamp}")
    try:
        shutil.copy2(path, backup)
        return backup
    except Exception as e:
        print(f"AVISO: no se pudo crear backup de {path}: {e}")
        return None


def _activity_file():
    """Ruta del archivo de actividad en la carpeta drive."""
    return os.path.join(get_drive_dir(), "sync_activity.json")


def get_last_activity():
    """Lee la última actividad (subida/descarga) desde la carpeta drive.

    Retorna un string legible o '' si no hay registro.
    """
    af = _activity_file()
    if not os.path.exists(af):
        return ""
    try:
        with open(af, "r", encoding="utf-8") as f:
            data = json.load(f)
        accion = data.get("accion", "?")
        fecha = data.get("fecha", "")
        user = data.get("usuario", "")
        return f"Última: {accion} | {fecha} | {user}"
    except Exception:
        return ""


def _write_activity(accion):
    """Registra la última actividad en la carpeta drive."""
    af = _activity_file()
    try:
        data = {
            "accion": accion,
            "fecha": datetime.now().strftime("%d/%m/%Y %H:%M:%S"),
            "usuario": os.environ.get("USERNAME", os.environ.get("USER", "desconocido")),
        }
        with open(af, "w", encoding="utf-8") as f:
            json.dump(data, f, ensure_ascii=False, indent=2)
    except Exception as e:
        print(f"AVISO: no se pudo escribir actividad: {e}")


def check_upload_conflict():
    """Verifica si la carpeta drive tiene versiones más nuevas que las que
    este equipo descargó por última vez (semáforo de colisiones).

    Compara el mtime actual de cada archivo en el drive contra la huella
    guardada en `_last_known_cloud_mtime` (poblada en `download_from_drive`).
    Si un archivo no tiene huella registrada (primera sesión, nunca se
    descargó), no bloquea.

    Retorna (ok, mensaje): ok=True si no hay conflicto (se puede subir),
    ok=False con un mensaje detallado si algún archivo cambió en el drive.
    """
    drive = get_drive_dir()
    conflictos = []
    for kind, local_path, name in get_sync_files():
        src = os.path.join(drive, name)
        if not os.path.exists(src):
            continue
        huella = _last_known_cloud_mtime.get(kind)
        if huella is None:
            continue
        try:
            actual = os.path.getmtime(src)
        except Exception:
            continue
        if actual > huella:
            conflictos.append(f"{kind}: {name}")

    if conflictos:
        detalle = "\n".join(f"  - {c}" for c in conflictos)
        mensaje = (
            "Otro usuario subió una versión más nueva al drive desde tu "
            "última descarga. Si subes ahora, borrarías su trabajo.\n\n"
            f"Archivo(s) en conflicto:\n{detalle}\n\n"
            "Usa el botón 'Sincronización de Recuperación' para traer los "
            "cambios de la nube sin perder tus datos locales de hoy."
        )
        return False, mensaje
    return True, ""


def upload_to_drive(force=False):
    """Copia la BD y el Excel locales hacia la carpeta drive.

    Antes de copiar (salvo `force=True`) verifica que nadie haya subido una
    versión más nueva desde la última descarga de este equipo
    (`check_upload_conflict`); si hay conflicto, no copia nada y retorna
    (False, mensaje) sugiriendo la Sincronización de Recuperación.

    Retorna (ok, mensaje) con las rutas exactas de origen y destino de cada
    archivo, y verifica que el destino haya quedado con tamaño > 0.
    """
    if not force:
        ok_check, msg_check = check_upload_conflict()
        if not ok_check:
            return False, msg_check

    drive = get_drive_dir()
    try:
        os.makedirs(drive, exist_ok=True)
    except Exception as e:
        return False, f"No se pudo crear la carpeta drive:\n{drive}\n{e}"

    lineas = [f"Carpeta destino: {drive}", ""]
    errores = []
    for kind, local_path, name in get_sync_files():
        if not os.path.exists(local_path):
            errores.append(f"{kind}: NO EXISTE {local_path}")
            continue
        dest = os.path.join(drive, name)
        try:
            shutil.copyfile(local_path, dest)
            size = os.path.getsize(dest) if os.path.exists(dest) else 0
            if size <= 0:
                errores.append(f"{kind}: se copió pero quedó vacío ({dest})")
            else:
                lineas.append(f"OK {name}  ({size} bytes)")
                lineas.append(f"   {local_path}")
                lineas.append(f"   -> {dest}")
                lineas.append("")
        except Exception as e:
            errores.append(f"{kind}: {type(e).__name__}: {e}")

    if errores:
        return False, "\n".join(lineas) + "\nErrores:\n" + "\n".join(errores)

    # Actualizar la huella con el mtime resultante para no autobloquear el
    # siguiente upload de este mismo equipo.
    for kind, local_path, name in get_sync_files():
        dest = os.path.join(drive, name)
        try:
            _last_known_cloud_mtime[kind] = os.path.getmtime(dest)
        except Exception:
            pass

    _write_activity("Subida al drive")
    return True, "\n".join(lineas).rstrip()


def download_from_drive():
    """Copia la BD y el Excel del drive hacia la copia local (sobrescribe).

    Antes de sobrescribir crea un backup local con fecha. Retorna
    (ok, mensaje).
    """
    drive = get_drive_dir()
    if not os.path.isdir(drive):
        return False, f"No existe la carpeta drive:\n{drive}"

    descargados = []
    errores = []
    for kind, local_path, name in get_sync_files():
        src = os.path.join(drive, name)
        if not os.path.exists(src):
            errores.append(f"{kind}: no existe {name} en el drive")
            continue
        try:
            _timestamped_backup(local_path)
            shutil.copy2(src, local_path)
            descargados.append(name)
            # Guardar la huella de tiempo del drive para el semáforo de colisiones
            try:
                _last_known_cloud_mtime[kind] = os.path.getmtime(src)
            except Exception:
                pass
        except Exception as e:
            errores.append(f"{kind}: {e}")

    if errores:
        return False, f"Carpeta origen:\n{drive}\n\nErrores al cargar:\n" + "\n".join(errores)

    _write_activity("Descarga desde drive")
    return True, f"Carpeta origen:\n{drive}\n\nCargados localmente:\n" + "\n".join(descargados)


# ---------------------------------------------------------------------------
# Fase 3: Sincronización de Recuperación
# ---------------------------------------------------------------------------
# Clave compuesta de identidad de una fila (igual que migrar_excel_a_sqlite).
_CLAVE_DB = ("cedis", "sf", "planta_obra", "articulo", "modalidad")

# Posición Excel (1-based) de cada columna de staging_rutas, invirtiendo el
# mapeo EXCEL_TO_DB_COLUMNS de fpa_validations (única fuente de verdad).
_DB_TO_EXCEL_POS = {v: k for k, v in fpa_validations.EXCEL_TO_DB_COLUMNS.items()}


def _copiar_binario(src, dst, intentos=5, espera=1.5):
    """Copia binaria (rb/wb) con reintentos para evitar [WinError 32].

    Si el archivo de origen o destino está bloqueado (p. ej. por OneDrive o un
    antivirus), espera `espera` segundos y reintenta hasta `intentos` veces.
    Lanza la última excepción si todos los intentos fallan.
    """
    ultimo_error = None
    for i in range(intentos):
        try:
            with open(src, "rb") as f_src, open(dst, "wb") as f_dst:
                f_dst.write(f_src.read())
            return
        except Exception as e:
            ultimo_error = e
            if i < intentos - 1:
                time.sleep(espera)
    raise ultimo_error


def sincronizar_recuperacion(conn, on_done=None):
    """Recupera cambios de la nube sin perder el trabajo local de hoy.

    Flujo (pensado para correr en un hilo de fondo, ya que abre Excel):
      1. Descarga el Excel del drive a un archivo temporal (no toca el Excel
         local que el usuario pueda tener abierto en pantalla).
      2. Toma de SQLite las filas trabajadas hoy (timestamp_modificacion).
      3. Con una instancia de Excel invisible y aislada (xw.App(visible=False)),
         actualiza en el temporal las filas que ya existen y agrega al final
         las que son nuevas.
      4. Reemplaza el Excel local real por el temporal ya actualizado
         (con backup previo) y refresca la huella de colisión.

    `on_done(ok, mensaje)` se invoca al terminar (o al fallar) si se provee;
    quien llame desde la GUI debe marshalarlo con `self.after(0, ...)`.
    Retorna igualmente (ok, mensaje) para permitir uso síncrono/pruebas.
    """
    drive = get_drive_dir()
    if not os.path.isdir(drive):
        msg = f"No existe la carpeta drive:\n{drive}"
        if on_done:
            on_done(False, msg)
        return False, msg

    excel_local = get_excel_file()
    nombre_excel = os.path.basename(excel_local)
    src_excel = os.path.join(drive, nombre_excel)
    if not os.path.exists(src_excel):
        msg = f"No existe {nombre_excel} en el drive:\n{drive}"
        if on_done:
            on_done(False, msg)
        return False, msg

    # Nombre temporal único (con marca de tiempo) para no chocar con un archivo
    # de recuperación previo que haya quedado bloqueado por un proceso Excel
    # invisible de un intento anterior.
    tmp_excel = os.path.join(
        BASE_DIR,
        f"_recuperacion_tmp_{datetime.now().strftime('%Y%m%d_%H%M%S_%f')}{os.path.splitext(nombre_excel)[1]}",
    )

    try:
        _copiar_binario(src_excel, tmp_excel)
    except Exception as e:
        msg = f"No se pudo descargar el Excel de recuperación: {e}"
        if on_done:
            on_done(False, msg)
        return False, msg

    # Ejecutar la manipulación de Excel en un subproceso aislado con timeout.
    # Así la GUI nunca se bloquea y, si Excel se cuelga, lo matamos por timeout
    # en vez de dejar la app colgada con un proceso EXCEL.EXE invisible.
    payload = {
        "db_path": DB_FILE,
        "tmp_excel": tmp_excel,
        "sheet": etl_processor.EXCEL_SHEET_NAME,
        "table": etl_processor.EXCEL_TABLE_NAME,
    }
    worker_path = os.path.join(BASE_DIR, "recovery_worker.py")
    insertadas = 0
    actualizadas = 0
    try:
        proc = subprocess.run(
            [sys.executable, worker_path],
            input=json.dumps(payload),
            capture_output=True,
            text=True,
            timeout=120,
        )
        stdout = (proc.stdout or "").strip()
        try:
            resultado = json.loads(stdout)
        except Exception:
            resultado = {"ok": False, "error": f"Respuesta inválida del worker: {stdout or proc.stderr}"}
        if proc.returncode != 0 or not resultado.get("ok"):
            error = resultado.get("error") or (proc.stderr or "").strip() or "Error desconocido en el worker."
            msg = f"Error durante la Sincronización de Recuperación: {error}"
            if on_done:
                on_done(False, msg)
            return False, msg
        insertadas = resultado.get("insertadas", 0)
        actualizadas = resultado.get("actualizadas", 0)
    except subprocess.TimeoutExpired:
        msg = "La Sincronización de Recuperación excedió el tiempo límite (120 s). Se abortó para no dejar la aplicación colgada."
        if on_done:
            on_done(False, msg)
        return False, msg
    except Exception as e:
        msg = f"Error al ejecutar la Sincronización de Recuperación: {e}"
        if on_done:
            on_done(False, msg)
        return False, msg

    # Reemplazar el Excel local real por el temporal ya actualizado, con backup.
    try:
        _timestamped_backup(excel_local)
        _copiar_binario(tmp_excel, excel_local)
    except Exception as e:
        msg = f"Excel recuperado pero no se pudo aplicar localmente: {e}"
        if on_done:
            on_done(False, msg)
        return False, msg
    finally:
        try:
            os.remove(tmp_excel)
        except Exception:
            pass

    # Refrescar la huella de colisión con el mtime del drive ya sincronizado.
    try:
        _last_known_cloud_mtime["excel"] = os.path.getmtime(src_excel)
    except Exception:
        pass

    _write_activity("Sincronización de Recuperación")
    msg = (
        f"Sincronización de Recuperación completada.\n"
        f"Filas insertadas: {insertadas}\nFilas actualizadas: {actualizadas}"
    )
    if on_done:
        on_done(True, msg)
    return True, msg
