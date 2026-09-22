"""
Instala en CEMEX_MATRIZ_INTEGRADO.xlsm la macro (ModSelectorCEDIS.bas) y el selector de CEDIS estilizado
(UserForm_SelectorCEDIS), y prueba la macro de verdad antes de dejarla.

Necesita que Excel permita el acceso al modelo de objetos de VBA. El script lo activa SOLO mientras
trabaja y lo deja exactamente como estaba (aunque falle).

Uso (el maestro debe estar cerrado en Excel; ejecutar desde la raíz del proyecto):
    python herramientas/instalar_vba_en_master.py            # instala en el maestro (deja CEMEX_MATRIZ_INTEGRADO_respaldo.xlsm)
    python herramientas/instalar_vba_en_master.py --probar   # solo prueba sobre una copia; no toca el maestro
"""
import argparse
import os
import re
import shutil
import gc
import subprocess
import sys
import tempfile
import time
import traceback
import winreg
import zipfile

import win32com.client as win32

import vba_selector

AQUI = os.path.dirname(os.path.abspath(__file__))       # herramientas/
RAIZ = os.path.dirname(AQUI)                              # raíz del proyecto: ahí vive el maestro
RUTA_BAS = os.path.join(RAIZ, "vba", "ModSelectorCEDIS.bas")
CLAVE = r"Software\Microsoft\Office\16.0\Excel\Security"
MARCA_PRUEBAS = "' ==== PRUEBAS TEMPORALES (se eliminan solas) ===="

PRUEBAS_VBA = MARCA_PRUEBAS + r'''
Public Function Prueba_Todo() As String
    Dim fso As Object, r As String
    Set fso = CreateObject("Scripting.FileSystemObject")
    On Error GoTo Fallo
    r = "norm1=" & NormalizarCEDIS("d836, dw66;  a&b")
    r = r & "|norm2=" & NormalizarCEDIS("d836 todos")
    r = r & "|norm3=[" & NormalizarCEDIS("   ") & "]"
    r = r & "|datos=" & TieneHojasDeDatos()
    r = r & "|actualizando=" & ActualizandoDatos()
    r = r & "|carpeta=" & CarpetaLibro(fso)
    r = r & "|cmd=" & RutaCmd()
    r = r & "|procesos=" & PuedeLanzarProcesos(CreateObject("WScript.Shell"))
    r = r & "|copia=" & fso.FileExists(PrepararCopiaDatos(fso, CarpetaTrabajo(fso)))
    Prueba_Todo = r
    Exit Function
Fallo:
    Prueba_Todo = r & "|ERROR " & Err.Number & ": " & Err.Description
End Function

Public Function Prueba_Formulario() As String
    Dim f As Object, i As Long, conNombre As Long
    Set f = VBA.UserForms.Add("UserForm_SelectorCEDIS")
    For i = 0 To f.LstCEDIS.ListCount - 1
        If Len(CStr(f.LstCEDIS.List(i, 1))) > 0 Then conNombre = conNombre + 1
    Next i
    Prueba_Formulario = "items=" & f.LstCEDIS.ListCount & "|con nombre=" & conNombre & "|ejemplo=" & _
                        f.LstCEDIS.List(f.LstCEDIS.ListCount - 1, 0) & " / " & f.LstCEDIS.List(f.LstCEDIS.ListCount - 1, 1) & _
                        "|conteo=" & f.LblConteo.Caption
    Unload f
End Function
'''


class AccesoVBOM:
    """Activa 'Confiar en el acceso al modelo de objetos de proyectos de VBA' solo mientras se usa."""

    def __enter__(self):
        self.k = winreg.CreateKey(winreg.HKEY_CURRENT_USER, CLAVE)
        try:
            self.previo = winreg.QueryValueEx(self.k, "AccessVBOM")[0]
        except FileNotFoundError:
            self.previo = None
        winreg.SetValueEx(self.k, "AccessVBOM", 0, winreg.REG_DWORD, 1)
        return self

    def _restaurar(self):
        try:
            if self.previo is None:
                winreg.DeleteValue(self.k, "AccessVBOM")
            else:
                winreg.SetValueEx(self.k, "AccessVBOM", 0, winreg.REG_DWORD, self.previo)
        except FileNotFoundError:
            pass

    def __exit__(self, *exc):
        try:
            self._restaurar()
            time.sleep(1.5)
            self._restaurar()       # por si Excel lo reescribió mientras terminaba de cerrarse
        finally:
            self.k.Close()
        return False


def pids_excel():
    salida = subprocess.run(["tasklist", "/FI", "IMAGENAME eq EXCEL.EXE", "/FO", "CSV", "/NH"],
                            capture_output=True, text=True).stdout
    return {int(l.split('","')[1]) for l in salida.strip().splitlines() if "EXCEL" in l.upper() and '","' in l}


def cerrar_excel(xl, pids_previos):
    """Cierra la instancia y espera a que su PROCESO termine (Excel reescribe sus ajustes al salir).
    Quit() no basta: mientras Python conserve la referencia COM, el proceso EXCEL.EXE no termina de
    salir; hay que soltarla (del + gc.collect()) antes de esperar."""
    propios = pids_excel() - pids_previos
    try:
        xl.Quit()
    except Exception:
        pass
    del xl
    gc.collect()
    limite = time.time() + 60
    while propios & pids_excel() and time.time() < limite:
        time.sleep(0.3)


def copiar(origen, destino):
    """Copia con PowerShell (Windows no deja que python.exe escriba en Documentos)."""
    env = dict(os.environ, SRC=origen, DST=destino)
    r = subprocess.run(["powershell", "-NoProfile", "-Command",
                        "Copy-Item -LiteralPath $env:SRC -Destination $env:DST -Force"], env=env, capture_output=True, text=True)
    if r.returncode != 0:
        sys.exit(f"No se pudo copiar a {destino}: {r.stderr.strip()}")


def texto_modulo(comp):
    cm = comp.CodeModule
    return cm.Lines(1, cm.CountOfLines) if cm.CountOfLines else ""


def instalar(vbp):
    """Quita versiones anteriores de la macro, importa el .bas y crea el formulario."""
    for i in range(vbp.VBComponents.Count, 0, -1):
        comp = vbp.VBComponents.Item(i)
        if comp.Type == 1 and "MostrarSelectorCEDIS" in texto_modulo(comp):     # módulo estándar con la macro vieja
            vbp.VBComponents.Remove(comp)
    vbp.VBComponents.Import(RUTA_BAS)
    vba_selector.crear_formulario(vbp)


def probar(xl, wb):
    """Añade funciones de prueba, las ejecuta de verdad en Excel y las elimina. Devuelve los resultados.

    OJO: CodeModule.AddFromString NO inserta al final del módulo, sino justo después de la sección de
    declaraciones (CountOfDeclarationLines) — verificado empíricamente. Por eso aquí se usa InsertLines
    en una posición explícita (antes+1 = una línea después de la última) y se borra por CONTEO de líneas,
    nunca buscando el marcador con texto: buscar y borrar "desde el marcador hasta el final" borró una vez
    todo el código real que había quedado después del bloque de pruebas."""
    comp = next(vbp_c for vbp_c in (wb.VBProject.VBComponents.Item(i) for i in range(1, wb.VBProject.VBComponents.Count + 1))
                if vbp_c.Type == 1 and "MostrarSelectorCEDIS" in texto_modulo(vbp_c))
    cm = comp.CodeModule
    antes = cm.CountOfLines
    cm.InsertLines(antes + 1, PRUEBAS_VBA)
    despues = cm.CountOfLines
    resultados = {}
    try:
        for nombre in ("Prueba_Todo", "Prueba_Formulario"):
            try:
                resultados[nombre] = xl.Run(f"'{wb.Name}'!{nombre}")
            except Exception as e:
                resultados[nombre] = f"FALLO AL EJECUTAR: {e}"
    finally:
        cm.DeleteLines(antes + 1, despues - antes)
        if cm.CountOfLines != antes:
            raise RuntimeError(f"Limpieza de pruebas incompleta: quedaron {cm.CountOfLines} líneas, "
                               f"debían ser {antes}. No se guarda el libro.")
    return resultados


def modulos_vba(ruta):
    from oletools.olevba import VBA_Parser
    p = VBA_Parser(ruta)
    try:
        return sorted({os.path.splitext(m[2])[0] for m in p.extract_macros()})
    finally:
        p.close()


def resumen_paquete(ruta):
    with zipfile.ZipFile(ruta) as z:
        n = set(z.namelist())
        return {
            "conexiones": len(re.findall(r"<connection ", z.read("xl/connections.xml").decode("utf8", "ignore"))) if "xl/connections.xml" in n else 0,
            "tablas": len([x for x in n if re.match(r"xl/tables/table\d+\.xml", x)]),
            "dinamicas": len([x for x in n if x.startswith("xl/pivotTables/pivotTable")]),
            "hojas": re.findall(r'<sheet [^>]*name="([^"]+)"', z.read("xl/workbook.xml").decode("utf8", "ignore")),
        }


def abrir_excel():
    xl = win32.DispatchEx("Excel.Application")
    xl.Visible = False
    xl.DisplayAlerts = False
    xl.EnableEvents = False              # no dispara las macros de la extracción
    xl.AskToUpdateLinks = False
    xl.AutomationSecurity = 1            # las macros deben poder ejecutarse para probarlas
    return xl


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--maestro", default=os.path.join(RAIZ, "CEMEX_MATRIZ_INTEGRADO.xlsm"))
    ap.add_argument("--probar", action="store_true", help="Solo probar sobre una copia (no toca el maestro)")
    args = ap.parse_args()
    maestro = os.path.abspath(args.maestro)
    if not os.path.isfile(maestro):
        sys.exit(f"No existe: {maestro}")
    if not args.probar and os.path.exists(os.path.join(os.path.dirname(maestro), "~$" + os.path.basename(maestro))):
        sys.exit("El libro maestro está abierto en Excel. Guárdelo (Ctrl+G), ciérrelo y vuelva a ejecutar.")

    tmp = tempfile.mkdtemp(prefix="cemex_vba_")
    respaldo = os.path.splitext(maestro)[0] + "_respaldo.xlsm"
    destino = os.path.join(tmp, "maestro_prueba.xlsm") if args.probar else maestro
    if args.probar:
        shutil.copy(maestro, destino)
    else:
        print("Respaldando el maestro...", flush=True)
        copiar(maestro, respaldo)

    xl, error, guardando, pids_previos = None, None, False, set()
    with AccesoVBOM():      # se restaura al salir de este bloque, DESPUÉS de cerrar Excel (Excel reescribe el ajuste al cerrar)
        try:
            pids_previos = pids_excel()
            xl = abrir_excel()
            print("Abriendo el libro (puede tardar)...", flush=True)
            wb = xl.Workbooks.Open(destino, 0, False)
            print("Instalando macro y formulario...", flush=True)
            instalar(wb.VBProject)
            print("Probando la macro en Excel...", flush=True)
            for k, v in probar(xl, wb).items():
                print(f"  {k}: {v}")
            comp = next(c for c in (wb.VBProject.VBComponents.Item(i) for i in range(1, wb.VBProject.VBComponents.Count + 1))
                        if c.Type == 1 and "AbrirTerminalMatriz" in texto_modulo(c))
            if comp.CodeModule.CountOfLines < 900:      # ModSelectorCEDIS.bas tiene ~945 líneas: red de seguridad
                raise RuntimeError(f"El módulo quedó con solo {comp.CodeModule.CountOfLines} líneas tras las "
                                   "pruebas; no se guarda (revise probar()).")
            if not args.probar:
                print("Guardando...", flush=True)
                guardando = True
                wb.Save()
            wb.Close(False)
            del wb
        except BaseException as e:
            error = e
            traceback.print_exc()
        finally:
            if xl is not None:
                cerrar_excel(xl, pids_previos)
                xl = None
                gc.collect()

    if error is not None:
        if guardando and os.path.exists(respaldo):
            print("Se restaura el maestro desde el respaldo.")
            copiar(respaldo, maestro)
        shutil.rmtree(tmp, ignore_errors=True)
        sys.exit(f"No se pudo completar: {error}")

    if not args.probar:
        antes, despues = resumen_paquete(respaldo), resumen_paquete(maestro)
        mods = modulos_vba(maestro)
        problemas = [k for k in antes if antes[k] != despues[k]]
        for necesario in ("ModSelectorCEDIS", "UserForm_SelectorCEDIS"):
            if necesario not in mods:
                problemas.append("falta " + necesario)
        if problemas:
            copiar(respaldo, maestro)
            sys.exit(f"Verificación fallida ({problemas}); se restauró el maestro.")
        print("Listo. Módulos VBA:", ", ".join(mods))
    shutil.rmtree(tmp, ignore_errors=True)


if __name__ == "__main__":
    main()
