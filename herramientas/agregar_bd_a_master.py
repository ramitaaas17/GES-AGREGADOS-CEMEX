"""
Agrega (o actualiza) la hoja BD_Completa en el libro maestro CEMEX_MATRIZ_INTEGRADO.xlsm y le aplica el
diseño del panel. La base se arma con la hoja BASE 2026 del propio maestro + las fuentes de apoyo
(carpeta fuentes_apoyo).

Uso (el maestro debe estar CERRADO en Excel; requiere Excel de escritorio + pywin32; ejecutar desde la raíz):
    python herramientas/agregar_bd_a_master.py                # agrega la hoja (si ya existe, no la toca)
    python herramientas/agregar_bd_a_master.py --forzar       # reconstruye la hoja (se pierden las correcciones hechas a mano)
    python herramientas/agregar_bd_a_master.py --solo-panel   # solo aplica el diseño del panel

Antes de modificar, deja una copia: CEMEX_MATRIZ_INTEGRADO_respaldo.xlsm. Si algo falla al verificar, se restaura.
"""
import argparse
import os
import re
import shutil
import subprocess
import sys
import tempfile
import traceback
import zipfile

import pandas as pd
import win32com.client as win32

AQUI = os.path.dirname(os.path.abspath(__file__))       # herramientas/
RAIZ = os.path.dirname(AQUI)                              # raíz del proyecto: ahí viven los .xlsm
sys.path.insert(0, RAIZ)                                  # para poder importar base_rutas, que vive en la raíz
import base_rutas
import panel_master

HOJA = 'BD_Completa'
NARANJA = panel_master.NARANJA


def hojas_del_libro(ruta):
    with zipfile.ZipFile(ruta) as z:
        return re.findall(r'<sheet [^>]*name="([^"]+)"', z.read('xl/workbook.xml').decode('utf8', 'ignore'))


def resumen_paquete(ruta):
    with zipfile.ZipFile(ruta) as z:
        n = set(z.namelist())
        return {
            'conexiones': len(re.findall(r'<connection ', z.read('xl/connections.xml').decode('utf8', 'ignore'))) if 'xl/connections.xml' in n else 0,
            'tablas': len([x for x in n if re.match(r'xl/tables/table\d+\.xml', x)]),
            'vba': 'xl/vbaProject.bin' in n,
            'dinamicas': len([x for x in n if x.startswith('xl/pivotTables/pivotTable')]),
        }


def copiar(origen, destino):
    """Copia con PowerShell: Windows no deja que python.exe escriba en Documentos (Acceso controlado a carpetas)."""
    env = dict(os.environ, SRC=origen, DST=destino)
    r = subprocess.run(['powershell', '-NoProfile', '-Command',
                        'Copy-Item -LiteralPath $env:SRC -Destination $env:DST -Force'], env=env, capture_output=True, text=True)
    if r.returncode != 0:
        sys.exit(f'No se pudo copiar a {destino}: {r.stderr.strip()}')


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument('--maestro', default=os.path.join(RAIZ, 'CEMEX_MATRIZ_INTEGRADO.xlsm'))
    ap.add_argument('--apoyo', default=os.path.join(RAIZ, 'fuentes_apoyo'))
    ap.add_argument('--forzar', action='store_true', help='Reconstruir BD_Completa aunque ya exista')
    ap.add_argument('--solo-panel', action='store_true', help='Solo aplicar el diseño del panel')
    args = ap.parse_args()

    maestro = os.path.abspath(args.maestro)
    if not os.path.isfile(maestro):
        sys.exit(f'No existe: {maestro}')
    if os.path.exists(os.path.join(os.path.dirname(maestro), '~$' + os.path.basename(maestro))):
        sys.exit('El libro maestro está abierto en Excel. Guárdelo, ciérrelo y vuelva a ejecutar.')

    hojas = hojas_del_libro(maestro)
    hay_bd = HOJA in hojas
    poner_bd = not args.solo_panel and (args.forzar or not hay_bd)
    if hay_bd and not args.forzar and not args.solo_panel:
        print(f'{HOJA} ya existe: no se modifica (use --forzar para reconstruirla). Solo se aplica el diseño del panel.')

    tmp = tempfile.mkdtemp(prefix='cemex_bd_')
    xl = None
    error = None
    guardando = False
    respaldo = os.path.join(os.path.dirname(maestro), os.path.splitext(os.path.basename(maestro))[0] + '_respaldo.xlsm')
    try:
        ruta_bd = None
        if poner_bd:
            print('Leyendo BASE 2026 y fuentes de apoyo...', flush=True)
            base = None
            if 'BASE 2026' in hojas:
                base = pd.read_excel(maestro, sheet_name='BASE 2026', header=3, dtype=str)
            bd = base_rutas.construir_bd(base, base_rutas.cargar_apoyo(args.apoyo))
            if bd is None:
                sys.exit('No hay datos para construir la base (falta BASE 2026 y fuentes de apoyo).')
            ruta_bd = base_rutas.guardar_bd_excel(bd, os.path.join(tmp, 'bd.xlsx'))
            print(f'Base: {len(bd)} registros. {base_rutas.resumen_texto(bd)}')

        print('Respaldando el maestro...', flush=True)
        copiar(maestro, respaldo)

        xl = win32.DispatchEx('Excel.Application')
        xl.Visible = False
        xl.DisplayAlerts = False
        xl.EnableEvents = False
        xl.AskToUpdateLinks = False
        xl.AutomationSecurity = 3          # no ejecutar macros del libro; se conservan al guardar

        print('Abriendo el maestro (puede tardar)...', flush=True)
        wb = xl.Workbooks.Open(maestro, 0, False)
        if poner_bd:
            wbd = xl.Workbooks.Open(ruta_bd, 0, True)
            if hay_bd:
                wb.Worksheets(HOJA).Delete()
            wbd.Worksheets(HOJA).Copy(wb.Worksheets(wb.Worksheets.Count))
            wbd.Close(False)
            wb.Worksheets(HOJA).Tab.Color = NARANJA
        panel = next((s for s in wb.Worksheets if s.Name == 'Panel_Control'), None)
        if panel is not None:
            panel_master.disenar_panel(xl, panel)
            panel.Activate()
        print('Guardando...', flush=True)
        guardando = True          # desde aquí el maestro puede quedar a medias
        wb.Save()
        wb.Close(False)
    except BaseException as e:
        error = e
        traceback.print_exc()
    finally:
        if xl is not None:
            try:
                xl.Quit()
            except Exception:
                pass
        shutil.rmtree(tmp, ignore_errors=True)

    if error is not None:         # Excel ya está cerrado: se puede restaurar sin conflicto de bloqueo
        if guardando and os.path.exists(respaldo):
            print('Se restaura el maestro desde el respaldo.')
            copiar(respaldo, maestro)
        sys.exit(f'No se pudo completar: {error}')

    # Verificación: no se debe perder nada del maestro
    antes, despues = resumen_paquete(respaldo), resumen_paquete(maestro)
    hojas_n = hojas_del_libro(maestro)
    problemas = [k for k in antes if antes[k] != despues[k]]
    if poner_bd and HOJA not in hojas_n:
        problemas.append('falta la hoja ' + HOJA)
    if problemas:
        copiar(respaldo, maestro)
        sys.exit(f'Verificación fallida ({problemas}); se restauró el maestro.')
    print('Listo. Hojas:', ', '.join(hojas_n))
    print('Verificado: conexiones, tablas, dinámica y macros intactas.')


if __name__ == '__main__':
    main()
