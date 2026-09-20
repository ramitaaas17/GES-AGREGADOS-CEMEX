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
    Dim pyScript As String
    Dim sourceFile As String
    Dim cmd As String
    Dim resCode As Long
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    basePath = ThisWorkbook.Path
    pyScript = basePath & "\matriz_integrada.py"
    
    ' Buscar automáticamente el archivo de extracción 2026 en la carpeta
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
        MsgBox "No se encontró el script de procesamiento: " & pyScript, vbCritical, "Error de Archivo"
        Exit Sub
    End If
    
    ' Preparar comando de ejecución
    Dim paramCedis As String
    If UCase(Trim(listaCEDIS)) = "TODOS" Or Trim(listaCEDIS) = "" Then
        paramCedis = ""
    Else
        paramCedis = " --cedis " & Trim(listaCEDIS)
    End If
    
    cmd = "cmd.exe /c python """ & pyScript & """ --fuente """ & sourceFile & """" & paramCedis
    
    Application.StatusBar = "Procesando matriz de precios en Python... por favor espere."
    Set wsh = CreateObject("WScript.Shell")
    
    ' Ejecución síncrona
    resCode = wsh.Run(cmd, 1, True)
    Application.StatusBar = False
    
    If resCode = 0 Or resCode = 1 Then
        ' Buscar el archivo más reciente generado en _salidas_integradas
        Dim outDir As String
        outDir = basePath & "\_salidas_integradas"
        Dim latestFile As String
        Dim latestDate As Date
        latestDate = #1/1/1990#
        
        If fso.FolderExists(outDir) Then
            Set folderObj = fso.GetFolder(outDir)
            For Each fl In folderObj.Files
                If InStr(1, fl.Name, "Matriz_Precios_Integral", vbTextCompare) > 0 _
                   And Left(fl.Name, 2) <> "~$" _
                   And LCase(fso.GetExtensionName(fl.Name)) = "xlsx" _
                   And fl.DateLastModified > latestDate Then
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
                Workbooks.Open latestFile
                If Err.Number <> 0 Then
                    MsgBox "No se pudo abrir automáticamente el archivo (" & Err.Description & "). Puede abrirlo directamente desde la carpeta _salidas_integradas.", vbInformation, "CEMEX"
                End If
                On Error GoTo 0
            End If
        Else
            MsgBox "Proceso completado. Revise la carpeta _salidas_integradas.", vbInformation, "CEMEX Listo"
        End If
    Else
        MsgBox "Hubo un problema al ejecutar Python. Verifique la consola.", vbExclamation, "CEMEX Advertencia"
    End If
End Sub
