import os
import glob
import openpyxl
import win32com.client

def get_all_cedis():
    files = glob.glob(r'C:\Users\coazo\ramitaaas17 GES-AGREGADOS-CEMEX main MatrizVentas\*2026*.xlsm')
    if not files:
        return ['D836', 'D838', 'DW66', 'D285', 'D847', 'D865', 'D881', 'DW74', 'DW75', 'DW88']
    
    wb = openpyxl.load_workbook(files[0], read_only=True, data_only=True)
    ws = wb['PVTA_MAT VK13']
    headers = next(ws.iter_rows(max_row=1, values_only=True))
    centro_idx = headers.index('Centro')
    
    cedis_set = set()
    for r in ws.iter_rows(min_row=2, values_only=True):
        val = r[centro_idx]
        if val:
            cedis_set.add(str(val).strip().upper())
    wb.close()
    return sorted(list(cedis_set))

def crear_panel_xlsm():
    cedis_list = get_all_cedis()
    cedis_str_array = '","'.join(cedis_list)
    
    out_xlsm = r'C:\Users\coazo\ramitaaas17 GES-AGREGADOS-CEMEX main MatrizVentas\PANEL_CONTROL_CEMEX.xlsm'
    
    excel = win32com.client.Dispatch('Excel.Application')
    excel.Visible = False
    excel.DisplayAlerts = False
    
    wb = excel.Workbooks.Add()
    ws = wb.Worksheets(1)
    ws.Name = "Panel_Control"
    try:
        excel.ActiveWindow.DisplayGridlines = False
    except:
        pass
    
    # -------------------------------------------------------------------------
    # 1. DISEÑO DE LA HOJA DE CONTROL
    # -------------------------------------------------------------------------
    # Anchos de columna
    ws.Columns("A").ColumnWidth = 4
    for c in ["B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L"]:
        ws.Columns(c).ColumnWidth = 12.5
        
    # Alturas
    ws.Rows(1).RowHeight = 8
    ws.Rows(2).RowHeight = 32
    ws.Rows(3).RowHeight = 20
    ws.Rows(4).RowHeight = 12
    
    # Hero Banner
    rng_hero = ws.Range("B2:L3")
    rng_hero.Interior.Color = 0x722D00 # RGB(0, 45, 114) Azul Marino CEMEX
    
    ws.Range("B2:L2").Merge()
    ws.Range("B3:L3").Merge()
    
    c_title = ws.Range("B2")
    c_title.Value = "CEMEX AGREGADOS  |  SISTEMA INTEGRAL DE MATRIZ DE PRECIOS Y GOBERNANZA"
    c_title.Font.Name = "Segoe UI"
    c_title.Font.Size = 13
    c_title.Font.Bold = True
    c_title.Font.Color = 0xFFFFFF
    c_title.VerticalAlignment = -4108 # xlCenter
    
    c_sub = ws.Range("B3")
    c_sub.Value = "  Generador de Matriz de Precios Venta vs Costo, Validación de Márgenes y Dashboard Ejecutivo"
    c_sub.Font.Name = "Segoe UI"
    c_sub.Font.Size = 9
    c_sub.Font.Color = 0xE0E0E0
    c_sub.VerticalAlignment = -4108
    
    # Tarjeta de Instrucciones
    rng_card = ws.Range("B5:L12")
    rng_card.Interior.Color = 0xFAF8F8 # Fondo tarjeta
    for b_edge in [7, 8, 9, 10]: # xlEdgeLeft, Top, Bottom, Right
        rng_card.Borders(b_edge).LineStyle = 1
        rng_card.Borders(b_edge).Color = 0xE2E8F0
        
    ws.Range("B5:L5").Merge()
    c_ctit = ws.Range("B5")
    c_ctit.Value = "  INSTRUCCIONES DE USO:"
    c_ctit.Font.Name = "Segoe UI"
    c_ctit.Font.Size = 10.5
    c_ctit.Font.Bold = True
    c_ctit.Font.Color = 0x722D00
    
    instrucciones = [
        "1. Haga clic en el botón verde inferior 'SELECCIONAR CEDIS Y GENERAR MATRIZ'.",
        "2. Se abrirá una ventana interactiva donde podrá marcar uno, varios o todos los 102 CEDIS.",
        "3. El sistema procesará en segundo plano los datos de Snowflake/SAP (Venta, Costo, Contratos y Densidad).",
        "4. Al finalizar, se abrirá automáticamente el reporte consolidado con el Dashboard Ejecutivo y las alertas."
    ]
    for i, inst in enumerate(instrucciones, start=7):
        ws.Range(f"B{i}:L{i}").Merge()
        c_i = ws.Range(f"B{i}")
        c_i.Value = f"  {inst}"
        c_i.Font.Name = "Segoe UI"
        c_i.Font.Size = 9
        c_i.Font.Color = 0x333333
        
    # Crear Botón con Forma 3D interactiva
    # Left, Top, Width, Height en pt
    # B14:L16
    left_pt = ws.Range("D14").Left
    top_pt = ws.Range("D14").Top
    width_pt = ws.Range("D14:I16").Width
    height_pt = ws.Range("D14:I16").Height
    
    btn_shape = ws.Shapes.AddShape(5, left_pt, top_pt, width_pt, height_pt) # 5 = msoShapeRoundedRectangle
    btn_shape.Name = "BtnLanzadorPrincipal"
    btn_shape.Fill.Solid()
    btn_shape.Fill.ForeColor.RGB = 0x28A745 # Verde Esmeralda (RGB 40, 167, 69)
    btn_shape.Line.Visible = False
    
    btn_shape.TextFrame.Characters().Text = "▶  SELECCIONAR CEDIS Y GENERAR MATRIZ"
    btn_shape.TextFrame.Characters().Font.Name = "Segoe UI"
    btn_shape.TextFrame.Characters().Font.Size = 12
    btn_shape.TextFrame.Characters().Font.Bold = True
    btn_shape.TextFrame.Characters().Font.Color = 0xFFFFFF
    btn_shape.TextFrame.HorizontalAlignment = -4108 # xlCenter
    btn_shape.TextFrame.VerticalAlignment = -4108   # xlCenter
    
    # Asignar la macro al botón
    btn_shape.OnAction = "MostrarSelectorCEDIS"
    
    # -------------------------------------------------------------------------
    # 2. CÓDIGO VBA: MÓDULO ESTÁNDAR
    # -------------------------------------------------------------------------
    vb_proj = wb.VBProject
    mod_std = vb_proj.VBComponents.Add(1) # 1 = vbext_ct_StdModule
    mod_std.Name = "ModControlador"
    
    code_mod = f'''Option Explicit

Public Sub MostrarSelectorCEDIS()
    On Error GoTo ErrHandler
    UserForm_SelectorCEDIS.Show
    Exit Sub
ErrHandler:
    MsgBox "Error al abrir el selector: " & Err.Description, vbCritical, "CEMEX Error"
End Sub

Public Sub EjecutarProcesoPython(ByVal paramCedis As String)
    Dim wsh As Object
    Dim fso As Object
    Dim basePath As String
    Dim pyScript As String
    Dim sourceFile As String
    Dim cmd As String
    Dim resCode As Long
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    basePath = ThisWorkbook.Path
    pyScript = basePath & "\\matriz_integrada.py"
    
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
        sourceFile = basePath & "\\Extracción de Datos Agregados _ Formatos Para Alta de Ruta (ultimo)_2026.xlsm"
    End If
    
    If Not fso.FileExists(pyScript) Then
        MsgBox "No se encontró el script Python: " & pyScript, vbCritical, "Error de Archivo"
        Exit Sub
    End If
    
    Dim argCedis As String
    If UCase(Trim(paramCedis)) = "TODOS" Or Trim(paramCedis) = "" Then
        argCedis = ""
    Else
        argCedis = " --cedis " & Trim(paramCedis)
    End If
    
    cmd = "cmd.exe /c python """ & pyScript & """ --fuente """ & sourceFile & """" & argCedis
    
    Application.StatusBar = "Procesando matriz en Python... por favor espere."
    Set wsh = CreateObject("WScript.Shell")
    resCode = wsh.Run(cmd, 1, True)
    Application.StatusBar = False
    
    If resCode = 0 Or resCode = 1 Then
        Dim outDir As String
        outDir = basePath & "\\_salidas_integradas"
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
            MsgBox "Proceso finalizado. Revise la carpeta _salidas_integradas.", vbInformation, "CEMEX Listo"
        End If
    Else
        MsgBox "Hubo un problema al ejecutar Python. Verifique la consola.", vbExclamation, "CEMEX Advertencia"
    End If
End Sub
'''
    mod_std.CodeModule.AddFromString(code_mod)
    
    # -------------------------------------------------------------------------
    # 3. CÓDIGO VBA: USERFORM CON CASILLAS MULTI-CEDIS
    # -------------------------------------------------------------------------
    uf = vb_proj.VBComponents.Add(3) # 3 = vbext_ct_MSForm
    uf.Name = "UserForm_SelectorCEDIS"
    try:
        uf.Properties.Item("Caption").Value = "CEMEX | Selector de Centros (CEDIS)"
        uf.Properties.Item("Width").Value = 360
        uf.Properties.Item("Height").Value = 430
    except:
        pass
    
    # Controles en el UserForm
    # Label Header
    lbl_h = uf.Designer.Controls.Add("Forms.Label.1", "LblHeader")
    lbl_h.Caption = "SELECCIÓN DE CENTROS (CEDIS)"
    lbl_h.Top = 10
    lbl_h.Left = 15
    lbl_h.Width = 320
    lbl_h.Height = 16
    lbl_h.Font.Bold = True
    lbl_h.Font.Size = 10
    lbl_h.ForeColor = 0x722D00
    
    # Label Sub
    lbl_s = uf.Designer.Controls.Add("Forms.Label.1", "LblSub")
    lbl_s.Caption = "Marque las casillas de los CEDIS a incluir en la Matriz:"
    lbl_s.Top = 28
    lbl_s.Left = 15
    lbl_s.Width = 320
    lbl_s.Height = 14
    lbl_s.Font.Size = 8.5
    
    # TextBox Filtro
    txt_f = uf.Designer.Controls.Add("Forms.TextBox.1", "TxtFiltro")
    txt_f.Top = 46
    txt_f.Left = 15
    txt_f.Width = 315
    txt_f.Height = 20
    txt_f.Font.Size = 9
    
    # ListBox con Checkboxes (MultiSelect = 1, ListStyle = 1)
    lst = uf.Designer.Controls.Add("Forms.ListBox.1", "LstCEDIS")
    lst.Top = 72
    lst.Left = 15
    lst.Width = 315
    lst.Height = 240
    lst.MultiSelect = 1 # fmMultiSelectMulti
    lst.ListStyle = 1   # fmListStyleOption (Checkboxes!)
    lst.Font.Size = 9.5
    
    # Boton Marcar Todos
    btn_all = uf.Designer.Controls.Add("Forms.CommandButton.1", "BtnTodos")
    btn_all.Caption = "☑ Marcar Todos"
    btn_all.Top = 320
    btn_all.Left = 15
    btn_all.Width = 95
    btn_all.Height = 24
    btn_all.Font.Size = 8.5
    
    # Boton Desmarcar
    btn_none = uf.Designer.Controls.Add("Forms.CommandButton.1", "BtnNinguno")
    btn_none.Caption = "☐ Desmarcar"
    btn_none.Top = 320
    btn_none.Left = 115
    btn_none.Width = 95
    btn_none.Height = 24
    btn_none.Font.Size = 8.5
    
    # Boton Demo D836/D838
    btn_demo = uf.Designer.Controls.Add("Forms.CommandButton.1", "BtnDemo")
    btn_demo.Caption = "🎯 Demo D836/D838"
    btn_demo.Top = 320
    btn_demo.Left = 215
    btn_demo.Width = 115
    btn_demo.Height = 24
    btn_demo.Font.Size = 8.5
    
    # Boton Principal Generar
    btn_gen = uf.Designer.Controls.Add("Forms.CommandButton.1", "BtnGenerar")
    btn_gen.Caption = "🚀 GENERAR MATRIZ Y DASHBOARD"
    btn_gen.Top = 352
    btn_gen.Left = 15
    btn_gen.Width = 315
    btn_gen.Height = 32
    btn_gen.Font.Bold = True
    btn_gen.Font.Size = 10
    btn_gen.BackColor = 0x28A745
    btn_gen.ForeColor = 0xFFFFFF
    
    # Código del UserForm
    code_uf = f'''Option Explicit

Private MasterList As Variant

Private Sub UserForm_Initialize()
    MasterList = Array("{cedis_str_array}")
    CargarLista ""
    
    ' Marcar por defecto D836 y D838 para la demo
    Dim i As Long
    For i = 0 To LstCEDIS.ListCount - 1
        If LstCEDIS.List(i) = "D836" Or LstCEDIS.List(i) = "D838" Then
            LstCEDIS.Selected(i) = True
        End If
    Next i
End Sub

Private Sub CargarLista(ByVal filtro As String)
    LstCEDIS.Clear
    Dim item As Variant
    For Each item In MasterList
        If filtro = "" Or InStr(1, CStr(item), filtro, vbTextCompare) > 0 Then
            LstCEDIS.AddItem CStr(item)
        End If
    Next item
End Sub

Private Sub TxtFiltro_Change()
    CargarLista Trim(TxtFiltro.Text)
End Sub

Private Sub BtnTodos_Click()
    Dim i As Long
    For i = 0 To LstCEDIS.ListCount - 1
        LstCEDIS.Selected(i) = True
    Next i
End Sub

Private Sub BtnNinguno_Click()
    Dim i As Long
    For i = 0 To LstCEDIS.ListCount - 1
        LstCEDIS.Selected(i) = False
    Next i
End Sub

Private Sub BtnDemo_Click()
    Dim i As Long
    For i = 0 To LstCEDIS.ListCount - 1
        If LstCEDIS.List(i) = "D836" Or LstCEDIS.List(i) = "D838" Then
            LstCEDIS.Selected(i) = True
        Else
            LstCEDIS.Selected(i) = False
        End If
    Next i
End Sub

Private Sub BtnGenerar_Click()
    Dim seleccionados As String
    Dim totalMarcados As Long
    Dim i As Long
    
    totalMarcados = 0
    seleccionados = ""
    
    For i = 0 To LstCEDIS.ListCount - 1
        If LstCEDIS.Selected(i) Then
            seleccionados = seleccionados & LstCEDIS.List(i) & " "
            totalMarcados = totalMarcados + 1
        End If
    Next i
    
    If totalMarcados = 0 Then
        MsgBox "Por favor seleccione al menos un CEDIS de la lista o haga clic en 'Marcar Todos'.", vbExclamation, "CEMEX Selección"
        Exit Sub
    End If
    
    If totalMarcados = UBound(MasterList) - LBound(MasterList) + 1 Then
        seleccionados = "TODOS"
    End If
    
    Me.Hide
    EjecutarProcesoPython Trim(seleccionados)
    Unload Me
End Sub
'''
    uf.CodeModule.AddFromString(code_uf)
    
    # -------------------------------------------------------------------------
    # 4. GUARDAR COMO ARCHIVO XLSM HABILITADO PARA MACROS
    # -------------------------------------------------------------------------
    # 52 = xlOpenXMLWorkbookMacroEnabled (.xlsm)
    wb.SaveAs(out_xlsm, 52)
    wb.Close(False)
    excel.Quit()
    print(f"¡Libro .xlsm con UserForm y Botón creado exitosamente en: {out_xlsm}!")

if __name__ == "__main__":
    crear_panel_xlsm()
