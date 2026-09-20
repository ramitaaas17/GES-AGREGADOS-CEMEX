Attribute VB_Name = "ModSelectorCEDIS"
Option Explicit

' =============================================================================
' CEMEX AGREGADOS - ORQUESTADOR UNIFICADO DE MATRIZ DE VENTAS Y GOBERNANZA
' Módulo de Automatización para Selección Multi-CEDIS y Ejecución en Un Clic
' =============================================================================

Public Sub MostrarSelectorCEDIS()
    ' Macro asignada al botón principal de la portada de Excel
    On Error GoTo ErrSelector
    
    Dim frm As Object
    ' Si existe el UserForm compilado en el proyecto, lo muestra
    On Error Resume Next
    UserForm_SelectorCEDIS.Show
    If Err.Number = 0 Then Exit Sub
    On Error GoTo ErrSelector
    
    ' Modo Alternativo (Sin UserForm): Cuadro de diálogo interactivo directo
    Dim cedisInput As String
    cedisInput = InputBox("Ingrese los códigos de CEDIS a procesar separados por espacio o coma:" & vbCrLf & _
                          "(Ejemplo: D836 D838 DW66)" & vbCrLf & vbCrLf & _
                          "Escriba 'TODOS' para procesar todos los centros del país.", _
                          "CEMEX - Selector de CEDIS", "D836 D838")
                          
    If Trim(cedisInput) = "" Then Exit Sub
    
    EjecutarMatrizPython cedisInput
    Exit Sub

ErrSelector:
    MsgBox "Error al iniciar el selector: " & Err.Description, vbCritical, "CEMEX Error"
End Sub

Public Sub EjecutarMatrizPython(ByVal listaCEDIS As String)
    Dim wsh As Object
    Dim fso As Object
    Dim basePath As String
    Dim outDir As String
    Dim pyScript As String
    Dim sourceFile As String
    Dim pyExe As String
    Dim logFile As String
    Dim cmd As String
    Dim resCode As Long
    Dim tInicio As Date
    
    tInicio = Now - TimeSerial(0, 2, 0) ' Margen de 2 minutos
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    basePath = ThisWorkbook.Path
    outDir = basePath & "\_salidas_integradas"
    pyScript = basePath & "\matriz_integrada.py"
    logFile = outDir & "\vba_execution.log"
    
    If Not fso.FolderExists(outDir) Then
        On Error Resume Next
        fso.CreateFolder outDir
        On Error GoTo 0
    End If
    
    ' Localizar el archivo de extracción 2026
    Dim fl As Object
    Dim folderObj As Object
    Set folderObj = fso.GetFolder(basePath)
    For Each fl In folderObj.Files
        If InStr(1, fl.Name, "2026", vbTextCompare) > 0 And InStr(1, fl.Name, ".xlsm", vbTextCompare) > 0 Then
            sourceFile = fl.Path
            Exit For
        End If
    Next fl
    
    If sourceFile = "" Then
        sourceFile = basePath & "\Extracción de Datos Agregados _ Formatos Para Alta de Ruta (ultimo)_2026.xlsm"
    End If
    
    If Not fso.FileExists(pyScript) Then
        MsgBox "No se encontró el script Python: " & pyScript, vbCritical, "Error de Archivo"
        Exit Sub
    End If
    
    ' Detectar ejecutable de Python
    pyExe = "C:\Users\coazo\AppData\Local\Programs\Python\Python311\python.exe"
    If Not fso.FileExists(pyExe) Then
        pyExe = "python"
    End If
    
    Dim paramCedis As String
    If UCase(Trim(listaCEDIS)) = "TODOS" Or Trim(listaCEDIS) = "" Then
        paramCedis = ""
    Else
        paramCedis = " --cedis " & Trim(listaCEDIS)
    End If
    
    cmd = "cmd.exe /c cd /d """ & basePath & """ && """ & pyExe & """ """ & pyScript & """ --fuente """ & sourceFile & """" & paramCedis & " --output """ & outDir & """ > """ & logFile & """ 2>&1"
    
    Application.StatusBar = "Procesando matriz de precios en Python... por favor espere."
    Set wsh = CreateObject("WScript.Shell")
    resCode = wsh.Run(cmd, 0, True)
    Application.StatusBar = False
    
    ' Buscar el archivo recién generado
    Dim latestFile As String
    Dim latestDate As Date
    latestDate = tInicio
    
    If fso.FolderExists(outDir) Then
        Set folderObj = fso.GetFolder(outDir)
        For Each fl In folderObj.Files
            If InStr(1, fl.Name, "Matriz_Precios_Integral", vbTextCompare) > 0 _
               And Left(fl.Name, 2) <> "~$" _
               And LCase(fso.GetExtensionName(fl.Name)) = "xlsx" _
               And fl.DateLastModified >= latestDate Then
                latestDate = fl.DateLastModified
                latestFile = fl.Path
            End If
        Next fl
    End If
    
    If latestFile <> "" Then
        Dim resp As VbMsgBoxResult
        resp = MsgBox("¡Matriz y Dashboard generados exitosamente!" & vbCrLf & vbCrLf & _
                      "Archivo: " & fso.GetFileName(latestFile) & vbCrLf & vbCrLf & _
                      "¿Desea abrir el reporte ahora?", vbInformation + vbYesNo, "CEMEX Matriz de Ventas")
        If resp = vbYes Then
            On Error Resume Next
            Dim wbNew As Workbook
            Set wbNew = Workbooks.Open(latestFile)
            If Err.Number <> 0 Then
                wsh.Run "explorer.exe """ & latestFile & """"
            End If
            On Error GoTo 0
        End If
    Else
        Dim logContent As String
        logContent = ""
        If fso.FileExists(logFile) Then
            On Error Resume Next
            Dim ts As Object
            Set ts = fso.OpenTextFile(logFile, 1)
            logContent = ts.ReadAll
            ts.Close
            On Error GoTo 0
        End If
        MsgBox "No se pudo generar el archivo de matriz." & vbCrLf & vbCrLf & _
               "Detalle del proceso:" & vbCrLf & _
               IIf(logContent <> "", Left(logContent, 600), "Código de salida: " & resCode), vbCritical, "CEMEX Error de Generación"
    End If
End Sub
