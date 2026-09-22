Attribute VB_Name = "ModSelectorCEDIS"
Option Explicit

' =============================================================================
' CEMEX AGREGADOS - ORQUESTADOR UNIFICADO DE MATRIZ DE VENTAS Y GOBERNANZA
' Modulo de Automatizacion para Seleccion Multi-CEDIS y Ejecucion en Un Clic
'
' Diseno "funciona en cualquier PC":
'  - Todo con late binding: sin referencias, Excel 32/64 bits, cualquier idioma.
'  - Python se detecta solo (py, python, rutas tipicas, Anaconda) y se valida que
'    tenga pandas/openpyxl; si faltan, ofrece instalarlas.
'  - Python escribe SOLO en %LOCALAPPDATA%\CEMEX_MatrizVentas (carpeta que Windows
'    nunca protege). Asi funciona aunque este activo el "Acceso controlado a
'    carpetas" de Defender, que bloquea a python.exe/cmd.exe en Documentos.
'    Excel copia luego el resultado a <carpeta del libro>\_salidas_integradas.
'  - Soporta libros abiertos desde OneDrive/SharePoint (URL) y rutas con acentos.
'  - Ejecuta en segundo plano sin congelar Excel y siempre informa el paso exacto
'    donde ocurre cualquier error.
'  - Si la organizacion bloquea que Excel ejecute programas (Defender ASR: Office
'    no puede crear procesos secundarios), guarda la seleccion y pide abrir
'    EJECUTAR_MATRIZ.bat con doble clic (sin saltarse la politica).
'  - Libro integrado: si este libro trae las hojas de datos (PVTA_MAT, PVTA_FTE...), espera
'    a que terminen las consultas de Power Query y procesa una copia de los datos actuales.
'  - Segundo boton (AbrirTerminalMatriz): abre una terminal y ejecuta el .bat.
' =============================================================================

Private Const NOMBRE_APP As String = "CEMEX Matriz de Ventas"
Private Const CLAVE_REG As String = "CEMEX_MatrizVentas"
Private Const SCRIPT_PY As String = "matriz_integrada.py"
Private Const SUBCARPETA_SALIDA As String = "_salidas_integradas"
Private Const ANIO_FUENTE As String = "2026"
Private Const MAX_MINUTOS As Long = 45
Private Const MAX_ESPERA_DATOS_MIN As Long = 20
Private Const COPIA_DATOS As String = "datos_fuente.xlsm"

Private mEnEjecucion As Boolean
Private mErrorPython As String

' -----------------------------------------------------------------------------
' PUNTO DE ENTRADA (boton de la portada)
' -----------------------------------------------------------------------------
Public Sub MostrarSelectorCEDIS()
    Dim frm As Object
    Dim entrada As String
    Dim lista As String

    On Error GoTo ErrSelector

    If Not EsWindows() Then
        MsgBox "Esta macro solo funciona en Excel para Windows.", vbExclamation, NOMBRE_APP
        Exit Sub
    End If

    ' Si el UserForm existe se usa; se instancia por nombre para que el modulo
    ' compile aunque el formulario no este en el libro.
    On Error Resume Next
    Set frm = VBA.UserForms.Add("UserForm_SelectorCEDIS")
    On Error GoTo ErrSelector

    If Not frm Is Nothing Then
        frm.Show
        Exit Sub
    End If

    ' Modo alternativo (sin UserForm)
    entrada = InputBox("Ingrese los codigos de CEDIS a procesar separados por espacio o coma:" & vbCrLf & _
                       "(Ejemplo: D836 D838 DW66)" & vbCrLf & vbCrLf & _
                       "Escriba TODOS para procesar todos los centros del pais.", _
                       "CEMEX - Selector de CEDIS", "D836 D838")
    lista = NormalizarCEDIS(entrada)
    If Len(lista) = 0 Then Exit Sub

    EjecutarMatrizPython lista
    Exit Sub

ErrSelector:
    MsgBox "Error al iniciar el selector (" & Err.Number & "): " & Err.Description, vbCritical, NOMBRE_APP
End Sub

' -----------------------------------------------------------------------------
' EJECUCION PRINCIPAL. listaCEDIS: codigos separados por espacio, o "TODOS".
' Firma compatible con UserForm_SelectorCEDIS.
' -----------------------------------------------------------------------------
Public Sub EjecutarMatrizPython(ByVal listaCEDIS As String)
    If mEnEjecucion Then
        MsgBox "Ya hay una generacion en curso. Espere a que termine.", vbExclamation, NOMBRE_APP
        Exit Sub
    End If
    mEnEjecucion = True
    On Error GoTo Fin
    EjecutarInterno listaCEDIS
Fin:
    mEnEjecucion = False
    Application.StatusBar = False
End Sub

Private Sub EjecutarInterno(ByVal listaCEDIS As String)
    Dim paso As String
    Dim fso As Object, wsh As Object, env As Object
    Dim baseDir As String, scriptPath As String, fuente As String
    Dim trabajoDir As String, salidaTrabajo As String, cacheDir As String
    Dim logFile As String, doneFile As String
    Dim pyCmd As String, comando As String, cedisArg As String
    Dim existentes As Object
    Dim inicio As Date
    Dim segundos As Long
    Dim estado As String
    Dim nuevo As String, finalPath As String
    Dim fallo As Boolean

    On Error GoTo ErrEjecutar

    paso = "validar entorno"
    If Not EsWindows() Then
        MsgBox "Esta macro solo funciona en Excel para Windows.", vbExclamation, NOMBRE_APP
        Exit Sub
    End If

    paso = "validar seleccion de CEDIS"
    cedisArg = NormalizarCEDIS(listaCEDIS)
    If Len(cedisArg) = 0 Then
        MsgBox "Seleccione al menos un CEDIS.", vbExclamation, NOMBRE_APP
        Exit Sub
    End If

    Set fso = CreateObject("Scripting.FileSystemObject")
    Set wsh = CreateObject("WScript.Shell")

    paso = "ubicar la carpeta del proyecto"
    baseDir = CarpetaLibro(fso)

    paso = "buscar " & SCRIPT_PY
    scriptPath = baseDir & "\" & SCRIPT_PY
    If Not fso.FileExists(scriptPath) Then
        MsgBox "No se encontro el script de Python:" & vbCrLf & scriptPath & vbCrLf & vbCrLf & _
               "Deje este libro en la misma carpeta que " & SCRIPT_PY & ".", vbCritical, NOMBRE_APP
        Exit Sub
    End If

    paso = "preparar carpeta de trabajo"
    trabajoDir = CarpetaTrabajo(fso)
    salidaTrabajo = trabajoDir & "\salidas"
    cacheDir = trabajoDir & "\cache"
    logFile = trabajoDir & "\ejecucion.log"
    doneFile = trabajoDir & "\ejecucion.fin"
    CrearCarpeta fso, salidaTrabajo
    CrearCarpeta fso, cacheDir

    If TieneHojasDeDatos() Then
        paso = "esperar a que terminen las consultas de datos"
        EsperarActualizacion
        paso = "copiar los datos actuales del libro"
        fuente = PrepararCopiaDatos(fso, trabajoDir)
    Else
        paso = "buscar el archivo de extraccion " & ANIO_FUENTE
        fuente = BuscarFuente(fso, baseDir)
        If Len(fuente) = 0 Then Exit Sub
    End If

    paso = "verificar permiso para ejecutar programas"
    If Not PuedeLanzarProcesos(wsh) Then
        paso = "preparar ejecucion manual"
        ModoManual fso, baseDir, trabajoDir, cedisArg
        Exit Sub
    End If

    paso = "detectar Python"
    pyCmd = ObtenerPython(fso, wsh, trabajoDir)
    If Len(pyCmd) = 0 Then Exit Sub

    paso = "preparar ejecucion"
    BorrarSiExiste fso, logFile
    BorrarSiExiste fso, doneFile
    Set existentes = ListarMatrices(fso, salidaTrabajo)

    ' /d sin AutoRun, /s + comillas externas: forma robusta de pasar rutas con
    ' espacios, acentos y parentesis. El marcador .fin permite saber el resultado
    ' sin bloquear Excel.
    comando = Q(RutaCmd()) & " /d /s /c " & Chr(34) & _
              "cd /d " & Q(trabajoDir) & " && " & _
              pyCmd & " " & Q(scriptPath) & _
              " --fuente " & Q(fuente)
    If UCase$(cedisArg) <> "TODOS" Then comando = comando & " --cedis " & cedisArg
    comando = comando & " --output " & Q(salidaTrabajo) & _
              " --cache-dir " & Q(cacheDir) & _
              " > " & Q(logFile) & " 2>&1" & _
              " && echo OK>" & Q(doneFile) & _
              " || echo FAIL>" & Q(doneFile) & Chr(34)

    paso = "lanzar Python"
    Set env = wsh.Environment("Process")
    env.Item("PYTHONUTF8") = "1"            ' acentos/simbolos correctos en el log
    env.Item("PYTHONIOENCODING") = "utf-8"
    env.Item("PYTHONDONTWRITEBYTECODE") = "1"
    inicio = Now
    wsh.Run comando, 0, False

    paso = "esperar a Python"
    Do
        If fso.FileExists(doneFile) Then Exit Do
        segundos = CLng(DateDiff("s", inicio, Now))
        If segundos > 30 And Not fso.FileExists(logFile) Then
            fallo = True                    ' cmd nunca arranco
            Exit Do
        End If
        If segundos > MAX_MINUTOS * 60 Then
            fallo = True
            Exit Do
        End If
        Application.StatusBar = "CEMEX: generando matriz con Python... " & _
                                Format$(segundos \ 60, "00") & ":" & Format$(segundos Mod 60, "00") & _
                                " (no cierre Excel)"
        DoEvents
        Application.Wait Now + TimeSerial(0, 0, 1)
    Loop
    Application.StatusBar = False

    paso = "leer resultado"
    estado = ""
    For segundos = 1 To 4                     ' el .fin puede existir un instante sin contenido
        If fso.FileExists(doneFile) Then estado = UCase$(Trim$(LeerTexto(doneFile)))
        If Len(estado) > 0 Or fallo Then Exit For
        Application.Wait Now + TimeSerial(0, 0, 1)
    Next segundos
    nuevo = ArchivoNuevo(fso, salidaTrabajo, existentes)

    If Not fallo And Left$(estado, 2) = "OK" And Len(nuevo) > 0 Then
        paso = "copiar resultado a la carpeta del proyecto"
        finalPath = CopiarResultado(fso, nuevo, baseDir)
        paso = "abrir resultado"
        OfrecerAbrir fso, wsh, finalPath
    Else
        paso = "mostrar diagnostico"
        MostrarFallo fso, wsh, logFile, estado, fallo
    End If
    Exit Sub

ErrEjecutar:
    Dim num As Long, desc As String
    num = Err.Number
    desc = Err.Description
    Application.StatusBar = False
    MsgBox "No se pudo completar el proceso." & vbCrLf & vbCrLf & _
           "Paso: " & paso & vbCrLf & _
           "Error " & num & ": " & desc, vbCritical, NOMBRE_APP
End Sub

' -----------------------------------------------------------------------------
' CEDIS: solo letras/numeros (evita inyectar caracteres en la linea de comandos)
' Devuelve "TODOS", "D836 D838 ..." o "" si no hay nada valido.
' -----------------------------------------------------------------------------
Private Function NormalizarCEDIS(ByVal texto As String) As String
    Dim partes() As String
    Dim i As Long, j As Long
    Dim tok As String, limpio As String, ch As String
    Dim res As String

    texto = Replace(Replace(Replace(Replace(texto, ",", " "), ";", " "), vbTab, " "), vbCrLf, " ")
    partes = Split(UCase$(Trim$(texto)), " ")

    For i = LBound(partes) To UBound(partes)
        tok = Trim$(partes(i))
        If Len(tok) > 0 Then
            limpio = ""
            For j = 1 To Len(tok)
                ch = Mid$(tok, j, 1)
                If (ch >= "A" And ch <= "Z") Or (ch >= "0" And ch <= "9") Then limpio = limpio & ch
            Next j
            If limpio = "TODOS" Then
                NormalizarCEDIS = "TODOS"
                Exit Function
            End If
            If Len(limpio) > 0 Then res = res & limpio & " "
        End If
    Next i
    NormalizarCEDIS = Trim$(res)
End Function

' -----------------------------------------------------------------------------
' RUTAS
' -----------------------------------------------------------------------------
Private Function CarpetaLibro(ByVal fso As Object) As String
    Dim p As String
    p = ThisWorkbook.Path
    If Len(p) = 0 Then
        Err.Raise vbObjectError + 1000, , "Guarde el libro en una carpeta antes de ejecutar la macro."
    End If

    If LCase$(Left$(p, 4)) = "http" Then
        p = UrlALocal(fso, p)
        If Len(p) = 0 Then
            Err.Raise vbObjectError + 1001, , "El libro esta abierto desde OneDrive/SharePoint (URL) y no se " & _
                "encontro su copia sincronizada." & vbCrLf & vbCrLf & _
                "Solucion: sincronice la carpeta con OneDrive o copiela a una ruta local " & _
                "(ej. C:\CEMEX), abra el libro desde ahi y vuelva a ejecutar."
        End If
    End If

    If Right$(p, 1) = "\" Then p = Left$(p, Len(p) - 1)
    CarpetaLibro = p
End Function

' Traduce https://.../personal/usuario/Documents/A/B a <OneDrive local>\A\B
Private Function UrlALocal(ByVal fso As Object, ByVal url As String) As String
    Dim raices(1 To 3) As String
    Dim partes() As String
    Dim r As Long, i As Long, k As Long
    Dim cand As String, cola As String

    On Error GoTo Fallo
    raices(1) = Environ$("OneDriveCommercial")
    raices(2) = Environ$("OneDrive")
    raices(3) = Environ$("OneDriveConsumer")

    url = DecodificarUrl(Mid$(url, InStr(9, url, "/") + 1))   ' quita "https://host/"
    partes = Split(Replace(url, "/", "\"), "\")

    For r = 1 To 3
        If Len(raices(r)) > 0 Then
            For i = 0 To UBound(partes)
                cola = ""
                For k = i To UBound(partes)
                    cola = cola & "\" & partes(k)
                Next k
                cand = raices(r) & cola
                If fso.FolderExists(cand) Then
                    UrlALocal = cand
                    Exit Function
                End If
            Next i
        End If
    Next r
    Exit Function
Fallo:
    UrlALocal = ""
End Function

Private Function DecodificarUrl(ByVal s As String) As String
    Dim b() As Byte
    Dim n As Long, i As Long
    Dim ch As String
    Dim stm As Object

    On Error GoTo Fallo
    ReDim b(0 To Len(s) * 3 + 3)
    i = 1
    Do While i <= Len(s)
        ch = Mid$(s, i, 1)
        If ch = "%" And i + 2 <= Len(s) Then
            b(n) = CByte("&H" & Mid$(s, i + 1, 2))
            i = i + 3
        Else
            b(n) = AscW(ch) And &HFF
            i = i + 1
        End If
        n = n + 1
    Loop
    If n = 0 Then Exit Function
    ReDim Preserve b(0 To n - 1)

    Set stm = CreateObject("ADODB.Stream")
    stm.Type = 1
    stm.Open
    stm.Write b
    stm.Position = 0
    stm.Type = 2
    stm.Charset = "utf-8"
    DecodificarUrl = stm.ReadText
    stm.Close
    Exit Function
Fallo:
    DecodificarUrl = s
End Function

' Carpeta donde Python SI puede escribir siempre (no protegida por Defender)
Private Function CarpetaTrabajo(ByVal fso As Object) As String
    Dim raiz As String
    raiz = Environ$("LOCALAPPDATA")
    If Len(raiz) = 0 Or Not fso.FolderExists(raiz) Then raiz = Environ$("TEMP")
    raiz = raiz & "\CEMEX_MatrizVentas"
    CrearCarpeta fso, raiz
    CarpetaTrabajo = raiz
End Function

Private Sub CrearCarpeta(ByVal fso As Object, ByVal ruta As String)
    If Not fso.FolderExists(ruta) Then fso.CreateFolder ruta
End Sub

Private Sub BorrarSiExiste(ByVal fso As Object, ByVal ruta As String)
    On Error Resume Next
    If fso.FileExists(ruta) Then fso.DeleteFile ruta, True
    On Error GoTo 0
End Sub

Private Function BuscarFuente(ByVal fso As Object, ByVal baseDir As String) As String
    Dim fl As Object
    Dim elegido As String
    Dim dlg As Object

    For Each fl In fso.GetFolder(baseDir).Files
        If Left$(fl.Name, 2) <> "~$" _
           And LCase$(fso.GetExtensionName(fl.Name)) = "xlsm" _
           And InStr(1, fl.Name, ANIO_FUENTE, vbTextCompare) > 0 _
           And StrComp(fl.Name, ThisWorkbook.Name, vbTextCompare) <> 0 Then
            elegido = fl.Path
            Exit For
        End If
    Next fl

    If Len(elegido) = 0 Then
        MsgBox "No se encontro en la carpeta el archivo de extraccion " & ANIO_FUENTE & _
               " (Extraccion de Datos Agregados ... " & ANIO_FUENTE & ".xlsm)." & vbCrLf & _
               "Seleccionelo manualmente en la siguiente ventana.", vbInformation, NOMBRE_APP
        Set dlg = Application.FileDialog(3)      ' msoFileDialogFilePicker
        dlg.AllowMultiSelect = False
        dlg.Title = "Seleccione el archivo de extraccion (.xlsm)"
        dlg.Filters.Clear
        dlg.Filters.Add "Excel", "*.xlsm;*.xlsx"
        If dlg.Show <> -1 Then Exit Function
        elegido = dlg.SelectedItems(1)
    End If
    BuscarFuente = elegido
End Function

' -----------------------------------------------------------------------------
' PYTHON: detecta un interprete con pandas + openpyxl
' Devuelve el comando ya listo para anteponer (ej.  "C:\...\python.exe"  o  py -3)
' -----------------------------------------------------------------------------
Private Function ObtenerPython(ByVal fso As Object, ByVal wsh As Object, ByVal trabajoDir As String) As String
    Dim candidatos As Collection
    Dim c As Variant
    Dim guardado As String
    Dim sinLibs As String

    Application.StatusBar = "CEMEX: buscando Python..."

    guardado = GetSetting(CLAVE_REG, "Config", "PythonCmd", "")
    If Len(guardado) > 0 Then
        If PythonEjecuta(wsh, guardado, "import pandas, openpyxl") Then
            ObtenerPython = guardado
            Application.StatusBar = False
            Exit Function
        End If
    End If

    Set candidatos = ListarPythons(fso)
    For Each c In candidatos
        If PythonEjecuta(wsh, CStr(c), "import pandas, openpyxl") Then
            SaveSetting CLAVE_REG, "Config", "PythonCmd", CStr(c)
            ObtenerPython = CStr(c)
            Application.StatusBar = False
            Exit Function
        End If
        If Len(sinLibs) = 0 Then
            If PythonEjecuta(wsh, CStr(c), "import sys") Then sinLibs = CStr(c)
        End If
    Next c
    Application.StatusBar = False

    If Len(sinLibs) > 0 Then
        If InstalarLibrerias(fso, wsh, sinLibs, trabajoDir) Then
            SaveSetting CLAVE_REG, "Config", "PythonCmd", sinLibs
            ObtenerPython = sinLibs
        End If
        Exit Function
    End If

    MsgBox "No se encontro Python en este equipo." & vbCrLf & vbCrLf & _
           "1) Instale Python 3 (Centro de Software de la empresa o https://www.python.org/downloads/)." & vbCrLf & _
           "    Marque la casilla 'Add python.exe to PATH'." & vbCrLf & _
           "2) Vuelva a hacer clic en el boton; las librerias se instalan solas." & _
           IIf(Len(mErrorPython) > 0, vbCrLf & vbCrLf & "Detalle tecnico: " & mErrorPython, ""), _
           vbExclamation, NOMBRE_APP
End Function

' Ejecuta "<py> -c codigo" oculto y espera; True si termina con codigo 0
Private Function PythonEjecuta(ByVal wsh As Object, ByVal pyCmd As String, ByVal codigo As String) As Boolean
    Dim rc As Long
    On Error GoTo Fallo
    rc = wsh.Run(Q(RutaCmd()) & " /d /s /c " & Chr(34) & pyCmd & " -c " & Chr(34) & codigo & Chr(34) & _
                 " >nul 2>&1" & Chr(34), 0, True)
    PythonEjecuta = (rc = 0)
    Exit Function
Fallo:
    mErrorPython = "Error " & Err.Number & ": " & Err.Description
    PythonEjecuta = False
End Function

Private Function ListarPythons(ByVal fso As Object) As Collection
    Dim col As New Collection
    Dim locApp As String, prog As String, perfil As String, pdata As String

    locApp = Environ$("LOCALAPPDATA")
    prog = Environ$("ProgramFiles")
    perfil = Environ$("USERPROFILE")
    pdata = Environ$("ProgramData")

    ' Lanzador oficial (lee el registro; cubre instalaciones nuevas y antiguas)
    col.Add "py -3"
    col.Add "python"
    col.Add "python3"

    AgregarDeCarpeta fso, col, locApp & "\Programs\Python"     ' Python3xx
    AgregarDeCarpeta fso, col, locApp & "\Python"              ' pythoncore-3.x-64 (Python install manager)
    AgregarDeCarpeta fso, col, Environ$("ProgramW6432")        ' C:\Program Files\Python3xx
    AgregarDeCarpeta fso, col, prog
    AgregarDeCarpeta fso, col, "C:\"                           ' C:\Python3xx
    AgregarUnico fso, col, perfil & "\anaconda3\python.exe"
    AgregarUnico fso, col, perfil & "\miniconda3\python.exe"
    AgregarUnico fso, col, pdata & "\anaconda3\python.exe"
    AgregarUnico fso, col, pdata & "\miniconda3\python.exe"
    AgregarUnico fso, col, locApp & "\anaconda3\python.exe"
    Set ListarPythons = col
End Function

' Anade <raiz>\<Python*|pythoncore*>\python.exe; las carpetas mas nuevas primero
Private Sub AgregarDeCarpeta(ByVal fso As Object, ByVal col As Collection, ByVal raiz As String)
    Dim fld As Object
    Dim sf As Object
    Dim tmp As New Collection
    Dim exe As Variant

    If Len(raiz) = 0 Then Exit Sub
    On Error GoTo Salir
    If Not fso.FolderExists(raiz) Then Exit Sub
    Set fld = fso.GetFolder(raiz)
    For Each sf In fld.SubFolders
        If LCase$(Left$(sf.Name, 6)) = "python" Then
            If fso.FileExists(sf.Path & "\python.exe") Then
                If tmp.Count = 0 Then
                    tmp.Add sf.Path & "\python.exe"
                Else
                    tmp.Add sf.Path & "\python.exe", , 1     ' alfabetico inverso = version mas alta primero
                End If
            End If
        End If
    Next sf
Salir:
    For Each exe In tmp
        col.Add Q(CStr(exe))
    Next exe
End Sub

Private Sub AgregarUnico(ByVal fso As Object, ByVal col As Collection, ByVal exe As String)
    If Len(exe) > 2 Then
        If fso.FileExists(exe) Then col.Add Q(exe)
    End If
End Sub

Private Function InstalarLibrerias(ByVal fso As Object, ByVal wsh As Object, ByVal pyCmd As String, _
                                   ByVal trabajoDir As String) As Boolean
    Dim resp As VbMsgBoxResult
    Dim pipLog As String
    Dim rc As Long

    resp = MsgBox("Se encontro Python, pero faltan las librerias necesarias (pandas, openpyxl, xlrd)." & vbCrLf & vbCrLf & _
                  "Requiere conexion a Internet y tarda 1-3 minutos." & vbCrLf & _
                  ChrW(191) & "Desea instalarlas ahora?", vbQuestion + vbYesNo, NOMBRE_APP)
    If resp <> vbYes Then Exit Function

    pipLog = trabajoDir & "\instalacion_librerias.log"
    BorrarSiExiste fso, pipLog
    Application.StatusBar = "CEMEX: instalando librerias de Python (1-3 min)..."
    On Error Resume Next
    wsh.Environment("Process").Item("PYTHONUTF8") = "1"
    rc = wsh.Run(Q(RutaCmd()) & " /d /s /c " & Chr(34) & _
                 "cd /d " & Q(trabajoDir) & " && " & pyCmd & _
                 " -m pip install --user --disable-pip-version-check pandas openpyxl xlrd" & _
                 " > " & Q(pipLog) & " 2>&1" & Chr(34), 0, True)
    On Error GoTo 0
    Application.StatusBar = False

    If PythonEjecuta(wsh, pyCmd, "import pandas, openpyxl") Then
        InstalarLibrerias = True
    Else
        MsgBox "No se pudieron instalar las librerias." & vbCrLf & vbCrLf & _
               "Ultimas lineas del registro:" & vbCrLf & ColaTexto(LeerTexto(pipLog), 700) & vbCrLf & vbCrLf & _
               "Registro completo: " & pipLog, vbCritical, NOMBRE_APP
    End If
End Function

' -----------------------------------------------------------------------------
' RESULTADOS
' -----------------------------------------------------------------------------
Private Function ListarMatrices(ByVal fso As Object, ByVal carpeta As String) As Object
    Dim d As Object
    Dim fl As Object
    Set d = CreateObject("Scripting.Dictionary")
    d.CompareMode = 1
    If fso.FolderExists(carpeta) Then
        For Each fl In fso.GetFolder(carpeta).Files
            d(fl.Name) = fl.DateLastModified
        Next fl
    End If
    Set ListarMatrices = d
End Function

' Archivo Matriz_Precios_Integral*.xlsx creado/modificado en esta ejecucion
Private Function ArchivoNuevo(ByVal fso As Object, ByVal carpeta As String, ByVal antes As Object) As String
    Dim fl As Object
    Dim mejor As Date
    Dim res As String

    If Not fso.FolderExists(carpeta) Then Exit Function
    For Each fl In fso.GetFolder(carpeta).Files
        If InStr(1, fl.Name, "Matriz_Precios_Integral", vbTextCompare) > 0 _
           And Left$(fl.Name, 2) <> "~$" _
           And LCase$(fso.GetExtensionName(fl.Name)) = "xlsx" Then
            If Not antes.Exists(fl.Name) Then
                If Len(res) = 0 Or fl.DateLastModified > mejor Then
                    res = fl.Path
                    mejor = fl.DateLastModified
                End If
            ElseIf fl.DateLastModified > antes(fl.Name) Then
                If Len(res) = 0 Or fl.DateLastModified > mejor Then
                    res = fl.Path
                    mejor = fl.DateLastModified
                End If
            End If
        End If
    Next fl
    ArchivoNuevo = res
End Function

' Copia a <proyecto>\_salidas_integradas (lo hace Excel, que Windows si deja escribir).
' Si no se puede, el resultado se deja en la carpeta de trabajo.
Private Function CopiarResultado(ByVal fso As Object, ByVal origen As String, ByVal baseDir As String) As String
    Dim destDir As String, destino As String

    On Error GoTo Fallo
    destDir = baseDir & "\" & SUBCARPETA_SALIDA
    CrearCarpeta fso, destDir
    destino = destDir & "\" & fso.GetFileName(origen)
    fso.CopyFile origen, destino, True
    CopiarResultado = destino
    Exit Function
Fallo:
    CopiarResultado = origen
End Function

Private Sub OfrecerAbrir(ByVal fso As Object, ByVal wsh As Object, ByVal archivo As String)
    Dim resp As VbMsgBoxResult

    resp = MsgBox(ChrW(161) & "Matriz y Dashboard generados exitosamente!" & vbCrLf & vbCrLf & _
                  "Archivo: " & fso.GetFileName(archivo) & vbCrLf & _
                  "Carpeta: " & fso.GetParentFolderName(archivo) & vbCrLf & vbCrLf & _
                  ChrW(191) & "Desea abrir el reporte ahora?", vbInformation + vbYesNo, NOMBRE_APP)
    If resp <> vbYes Then Exit Sub

    On Error Resume Next
    Workbooks.Open archivo
    If Err.Number <> 0 Then
        Err.Clear
        ThisWorkbook.FollowHyperlink archivo          ' abre con la app predeterminada
        If Err.Number <> 0 Then
            Err.Clear
            wsh.Run "explorer.exe /select," & Q(archivo), 1, False
        End If
    End If
    On Error GoTo 0
End Sub

Private Sub MostrarFallo(ByVal fso As Object, ByVal wsh As Object, ByVal logFile As String, _
                         ByVal estado As String, ByVal noArranco As Boolean)
    Dim detalle As String
    Dim resp As VbMsgBoxResult

    If fso.FileExists(logFile) Then
        detalle = ColaTexto(LeerTexto(logFile), 600)
    End If
    If Len(Trim$(detalle)) = 0 Then
        If noArranco Then
            detalle = "Python no llego a iniciar o supero los " & MAX_MINUTOS & " minutos de espera."
        Else
            detalle = "El proceso termino sin generar el archivo de la matriz (estado: " & estado & ")."
        End If
    End If

    resp = MsgBox("No se pudo generar el archivo de matriz." & vbCrLf & vbCrLf & _
                  "Detalle del proceso:" & vbCrLf & detalle & vbCrLf & vbCrLf & _
                  ChrW(191) & "Desea abrir el registro completo?", vbCritical + vbYesNo, NOMBRE_APP)
    If resp = vbYes And fso.FileExists(logFile) Then
        wsh.Run "notepad.exe " & Q(logFile), 1, False
    End If
End Sub

' -----------------------------------------------------------------------------
' LIBRO INTEGRADO: datos de extraccion + panel en el mismo archivo
' -----------------------------------------------------------------------------
' True si este libro trae las hojas de datos (es el libro integrado)
Private Function TieneHojasDeDatos() As Boolean
    Dim nombres As Variant
    Dim i As Long
    Dim ws As Object

    nombres = Array("PVTA_MAT VK13", "PVTA_FTE VK13", "CONT_COMPRA TRAOPE", "PESO_VOL", "CONT_VTA ZSDD4501")
    For i = LBound(nombres) To UBound(nombres)
        Set ws = Nothing
        On Error Resume Next
        Set ws = ThisWorkbook.Worksheets(CStr(nombres(i)))
        On Error GoTo 0
        If ws Is Nothing Then Exit Function
    Next i
    TieneHojasDeDatos = True
End Function

' True si alguna consulta de Power Query / conexion de datos se esta actualizando ahora mismo
Private Function ActualizandoDatos() As Boolean
    Dim cn As Object, ws As Object, lo As Object
    Dim activo As Boolean

    On Error Resume Next
    For Each cn In ThisWorkbook.Connections
        activo = False
        activo = cn.OLEDBConnection.Refreshing
        If activo Then
            ActualizandoDatos = True
            Exit Function
        End If
    Next cn
    For Each ws In ThisWorkbook.Worksheets
        For Each lo In ws.ListObjects
            activo = False
            activo = lo.QueryTable.Refreshing
            If activo Then
                ActualizandoDatos = True
                Exit Function
            End If
        Next lo
    Next ws
End Function

' Si los datos se estan actualizando, espera a que terminen (asi la matriz nunca usa datos a medias)
Private Sub EsperarActualizacion()
    Dim inicio As Date
    Dim segundos As Long

    If Not ActualizandoDatos() Then Exit Sub
    inicio = Now
    Do While ActualizandoDatos()
        segundos = CLng(DateDiff("s", inicio, Now))
        If segundos > MAX_ESPERA_DATOS_MIN * 60 Then
            Err.Raise vbObjectError + 1002, , "Las consultas de datos siguen actualizandose despues de " & _
                MAX_ESPERA_DATOS_MIN & " minutos. Espere a que terminen y vuelva a pulsar el boton."
        End If
        Application.StatusBar = "CEMEX: los datos se estan actualizando (Power Query)... " & _
                                Format$(segundos \ 60, "00") & ":" & Format$(segundos Mod 60, "00") & _
                                " - la matriz se generara al terminar"
        DoEvents
        Application.Wait Now + TimeSerial(0, 0, 1)
    Loop
    On Error Resume Next
    Application.CalculateUntilAsyncQueriesDone
    On Error GoTo 0
    Application.StatusBar = False
End Sub

' Guarda una copia de los datos tal como estan ahora (aunque el libro no este guardado) en la
' carpeta de trabajo. Python lee esa copia: nunca usa datos viejos ni el archivo bloqueado.
Private Function PrepararCopiaDatos(ByVal fso As Object, ByVal trabajoDir As String) As String
    Dim destino As String

    destino = trabajoDir & "\" & COPIA_DATOS
    BorrarSiExiste fso, destino          ' que nunca quede una copia vieja
    Application.StatusBar = "CEMEX: copiando los datos actuales del libro..."
    ThisWorkbook.SaveCopyAs destino
    Application.StatusBar = False
    If Not fso.FileExists(destino) Then
        Err.Raise vbObjectError + 1003, , "No se pudo crear la copia de los datos en: " & destino
    End If
    PrepararCopiaDatos = destino
End Function

' -----------------------------------------------------------------------------
' BOTON "ABRIR TERMINAL": abre una terminal y ejecuta EJECUTAR_MATRIZ.bat.
' Si la politica de la organizacion impide que Excel abra programas, copia la ruta del
' .bat al portapapeles para ejecutarla con Win + R (sin saltarse la politica).
' -----------------------------------------------------------------------------
Public Sub AbrirTerminalMatriz()
    Dim paso As String
    Dim fso As Object, wsh As Object, ts As Object
    Dim baseDir As String, lanzador As String, trabajoDir As String
    Dim copiado As Boolean
    Dim num As Long, desc As String

    On Error GoTo ErrTerminal

    paso = "validar entorno"
    If Not EsWindows() Then
        MsgBox "Esta macro solo funciona en Excel para Windows.", vbExclamation, NOMBRE_APP
        Exit Sub
    End If

    Set fso = CreateObject("Scripting.FileSystemObject")
    Set wsh = CreateObject("WScript.Shell")

    paso = "ubicar la carpeta del proyecto"
    baseDir = CarpetaLibro(fso)
    lanzador = baseDir & "\EJECUTAR_MATRIZ.bat"
    If Not fso.FileExists(lanzador) Then
        MsgBox "No se encontro el lanzador:" & vbCrLf & lanzador & vbCrLf & vbCrLf & _
               "Deje este libro en la misma carpeta que EJECUTAR_MATRIZ.bat.", vbCritical, NOMBRE_APP
        Exit Sub
    End If

    If TieneHojasDeDatos() Then
        paso = "preparar carpeta de trabajo"
        trabajoDir = CarpetaTrabajo(fso)
        paso = "esperar a que terminen las consultas de datos"
        EsperarActualizacion
        paso = "copiar los datos actuales del libro"
        PrepararCopiaDatos fso, trabajoDir
        ' Solicitud sin CEDIS: el .bat usa esta copia de datos y pregunta los CEDIS
        Set ts = fso.CreateTextFile(trabajoDir & "\solicitud.txt", True, False)
        ts.Close
    End If

    paso = "abrir la terminal"
    If PuedeLanzarProcesos(wsh) Then
        wsh.Run Q(RutaCmd()) & " /d /s /k " & Chr(34) & "cd /d " & Q(baseDir) & " && " & Q(lanzador) & Chr(34), 1, False
    Else
        copiado = CopiarAlPortapapeles(Q(lanzador))
        MsgBox "Excel no puede abrir la terminal en este equipo (politica de seguridad de su organizacion)." & _
               vbCrLf & vbCrLf & _
               IIf(copiado, "La ruta de EJECUTAR_MATRIZ.bat ya esta en el portapapeles:", "Ruta de EJECUTAR_MATRIZ.bat:") & _
               vbCrLf & lanzador & vbCrLf & vbCrLf & _
               "Para ejecutarlo:" & vbCrLf & _
               "  1) Presione Win + R" & vbCrLf & _
               "  2) Pegue con Ctrl + V" & vbCrLf & _
               "  3) Presione Enter" & vbCrLf & vbCrLf & _
               "Los datos actuales de este libro ya quedaron listos para el .bat.", vbInformation, NOMBRE_APP
    End If
    Exit Sub

ErrTerminal:
    num = Err.Number
    desc = Err.Description
    Application.StatusBar = False
    MsgBox "No se pudo abrir la terminal." & vbCrLf & vbCrLf & _
           "Paso: " & paso & vbCrLf & _
           "Error " & num & ": " & desc, vbCritical, NOMBRE_APP
End Sub

Private Function CopiarAlPortapapeles(ByVal texto As String) As Boolean
    Dim dobj As Object
    On Error GoTo Fallo
    Set dobj = CreateObject("new:{1C3B4210-F441-11CE-B9EA-00AA006B1A69}")   ' MSForms.DataObject
    dobj.SetText texto
    dobj.PutInClipboard
    CopiarAlPortapapeles = True
    Exit Function
Fallo:
    CopiarAlPortapapeles = False
End Function

' -----------------------------------------------------------------------------
' EXCEL BLOQUEADO PARA EJECUTAR PROGRAMAS (politica de la organizacion)
' -----------------------------------------------------------------------------
' True si Excel puede crear procesos. La regla ASR "Bloquear que las aplicaciones
' de Office creen procesos secundarios" hace fallar wsh.Run (error 70) y Shell (error 5).
Private Function PuedeLanzarProcesos(ByVal wsh As Object) As Boolean
    Dim rc As Long
    On Error GoTo Bloqueado
    rc = wsh.Run(Q(RutaCmd()) & " /d /c exit 0", 0, True)
    PuedeLanzarProcesos = (rc = 0)
    Exit Function
Bloqueado:
    mErrorPython = "Error " & Err.Number & ": " & Err.Description
    PuedeLanzarProcesos = False
End Function

' Guarda la seleccion; el usuario abre EJECUTAR_MATRIZ.bat con doble clic desde el
' Explorador (proceso hijo de Explorer, no de Office).
Private Sub ModoManual(ByVal fso As Object, ByVal baseDir As String, ByVal trabajoDir As String, _
                       ByVal cedisArg As String)
    Dim ts As Object
    Dim lanzador As String
    Dim resumen As String

    lanzador = baseDir & "\EJECUTAR_MATRIZ.bat"
    Set ts = fso.CreateTextFile(trabajoDir & "\solicitud.txt", True, False)
    ts.WriteLine cedisArg
    ts.Close

    If Not fso.FileExists(lanzador) Then
        MsgBox "Excel no puede ejecutar programas en este equipo y falta el lanzador:" & vbCrLf & lanzador, _
               vbCritical, NOMBRE_APP
        Exit Sub
    End If

    resumen = cedisArg
    If Len(resumen) > 120 Then resumen = Left$(resumen, 120) & "..."

    MsgBox "La politica de seguridad de su organizacion no permite que Excel ejecute programas " & _
           "(Windows Defender, regla ASR), por eso la matriz se genera en 2 pasos." & vbCrLf & vbCrLf & _
           "Su seleccion ya esta guardada: " & resumen & vbCrLf & vbCrLf & _
           "AHORA haga doble clic en este archivo desde el Explorador de Windows:" & vbCrLf & lanzador & vbCrLf & vbCrLf & _
           "El reporte se abrira solo al terminar.", vbInformation, NOMBRE_APP
End Sub

' -----------------------------------------------------------------------------
' UTILIDADES
' -----------------------------------------------------------------------------
Private Function Q(ByVal s As String) As String
    Q = Chr(34) & s & Chr(34)
End Function

Private Function RutaCmd() As String
    Dim ruta As String
    ruta = Environ$("ComSpec")
    If Len(ruta) = 0 Then ruta = "cmd.exe"
    RutaCmd = ruta
End Function

Private Function EsWindows() As Boolean
    EsWindows = (InStr(1, Application.OperatingSystem, "Windows", vbTextCompare) > 0)
End Function

Private Function ColaTexto(ByVal s As String, ByVal maxLen As Long) As String
    s = Trim$(s)
    If Len(s) > maxLen Then s = "..." & Right$(s, maxLen)
    ColaTexto = s
End Function

' Lee texto en UTF-8 (el log de Python lo es); si falla, lectura ANSI
Private Function LeerTexto(ByVal ruta As String) As String
    Dim stm As Object
    Dim fso As Object
    Dim ts As Object

    On Error GoTo Alterno
    Set stm = CreateObject("ADODB.Stream")
    stm.Type = 2
    stm.Charset = "utf-8"
    stm.Open
    stm.LoadFromFile ruta
    LeerTexto = stm.ReadText
    stm.Close
    Exit Function
Alterno:
    On Error Resume Next
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set ts = fso.OpenTextFile(ruta, 1, False)
    LeerTexto = ts.ReadAll
    ts.Close
End Function
