"""
Construye CEMEX_MATRIZ_INTEGRADO.xlsm: el archivo de extracción de datos (con sus consultas
Power Query, tablas, dinámica y macros) + la hoja Panel_Control con los botones, en UN solo libro.

Por qué el panel se añade AL archivo de extracción y no al revés: las consultas Power Query
(PVTA_MAT, PVTA_FTE, CONT_COMPRA, CONT_VTA) viven en el libro de extracción y no sobreviven a
copiar sus hojas a otro libro.

Uso (requiere Excel de escritorio + pywin32; ejecutar desde la raíz del proyecto):
    python herramientas/construir_integrado.py                  # crea CEMEX_MATRIZ_INTEGRADO.xlsm
    python herramientas/construir_integrado.py --forzar         # reconstruye aunque ya exista (pierde el VBA importado)
    python herramientas/construir_integrado.py --extraccion "otra extraccion.xlsm"

Los archivos originales nunca se modifican: se trabaja sobre copias en una carpeta temporal y es
Excel quien guarda el resultado (Windows no deja que Python escriba en Documentos si está activo
el "Acceso controlado a carpetas").

Después de construirlo hay que importar el VBA una vez (ver README, o usar herramientas/instalar_vba_en_master.py):
Alt+F11 > Archivo > Importar > vba/ModSelectorCEDIS.bas, y arrastrar UserForm_SelectorCEDIS desde el panel antiguo.
"""
import argparse
import glob
import os
import shutil
import sys
import tempfile

import win32com.client as win32

import panel_master

AQUI = os.path.dirname(os.path.abspath(__file__))       # herramientas/
RAIZ = os.path.dirname(AQUI)                              # raíz del proyecto: ahí viven los .xlsm
NOMBRE_SALIDA = "CEMEX_MATRIZ_INTEGRADO.xlsm"
HOJA_PANEL = "Panel_Control"

def buscar_extraccion():
    candidatos = [
        f for f in glob.glob(os.path.join(RAIZ, "*.xlsm"))
        if "2026" in os.path.basename(f)
        and not os.path.basename(f).startswith("~$")
        and "PANEL" not in os.path.basename(f).upper()
        and "INTEGRADO" not in os.path.basename(f).upper()
    ]
    if not candidatos:
        sys.exit("No se encontró el archivo de extracción (*2026*.xlsm). Indíquelo con --extraccion.")
    return sorted(candidatos)[0]


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--extraccion", help="Archivo de extracción (.xlsm). Default: el *2026*.xlsm de esta carpeta")
    ap.add_argument("--panel", default=os.path.join(RAIZ, "PANEL_CONTROL_CEMEX.xlsm"), help="Panel de control origen")
    ap.add_argument("--salida", default=os.path.join(RAIZ, NOMBRE_SALIDA), help="Archivo integrado a crear")
    ap.add_argument("--forzar", action="store_true", help="Sobrescribir el integrado si ya existe")
    args = ap.parse_args()

    extraccion = os.path.abspath(args.extraccion) if args.extraccion else buscar_extraccion()
    panel = os.path.abspath(args.panel)
    salida = os.path.abspath(args.salida)
    for f in (extraccion, panel):
        if not os.path.isfile(f):
            sys.exit(f"No existe: {f}")
    if os.path.exists(salida) and not args.forzar:
        sys.exit(f"Ya existe {salida}\nSi lo reconstruye se pierde el VBA que haya importado. Use --forzar para continuar.")

    print("Extracción:", os.path.basename(extraccion))
    print("Panel     :", os.path.basename(panel))
    print("Salida    :", salida)

    tmp = tempfile.mkdtemp(prefix="cemex_integrado_")
    xl = None
    try:
        t_ext = os.path.join(tmp, "extraccion.xlsm")
        t_pan = os.path.join(tmp, "panel.xlsm")
        print("Copiando originales a carpeta temporal...", flush=True)
        shutil.copy(extraccion, t_ext)
        shutil.copy(panel, t_pan)

        xl = win32.DispatchEx("Excel.Application")   # instancia propia, no toca su Excel abierto
        xl.Visible = False
        xl.DisplayAlerts = False
        xl.EnableEvents = False                       # no disparar las macros de la extracción
        xl.AskToUpdateLinks = False
        xl.AutomationSecurity = 3                     # msoAutomationSecurityForceDisable

        print("Abriendo extracción (puede tardar)...", flush=True)
        wb = xl.Workbooks.Open(t_ext, 0, False)       # UpdateLinks=0
        wb_pan = xl.Workbooks.Open(t_pan, 0, True)

        if any(s.Name == HOJA_PANEL for s in wb.Worksheets):
            sys.exit(f"El libro de extracción ya trae una hoja llamada {HOJA_PANEL}.")

        print("Copiando la hoja del panel...", flush=True)
        wb_pan.Worksheets(HOJA_PANEL).Copy(wb.Worksheets(1))   # Before = primera hoja
        wb_pan.Close(False)
        ws = wb.Worksheets(HOJA_PANEL)
        ws.Activate()

        # ---- Diseño del panel (simple, cálido; botones con su macro y estado de los datos)
        panel_master.disenar_panel(xl, ws)

        print("Guardando integrado (puede tardar)...", flush=True)
        wb.SaveAs(salida, 52)          # xlOpenXMLWorkbookMacroEnabled (Excel sobrescribe si --forzar)
        wb.Close(False)
        print("Listo:", salida)
    finally:
        if xl is not None:
            try:
                xl.Quit()
            except Exception:
                pass
        shutil.rmtree(tmp, ignore_errors=True)


if __name__ == "__main__":
    main()
