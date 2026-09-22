"""
Formulario UserForm_SelectorCEDIS: selector de CEDIS con lista de casillas (código + nombre), buscador,
contador y colores cálidos del panel. Se crea por COM (requiere acceso al modelo de objetos de VBA).
El catálogo se lee del propio libro (tabla PVTA_MAT y hoja BD_Completa), así no hay que reconstruirlo
cuando aparece un CEDIS nuevo.
"""
from panel_master import ARENA, BLANCO, BORDE, CAFE, CREMA, NARANJA, TERRACOTA, _rgb

NOMBRE_FORMULARIO = "UserForm_SelectorCEDIS"
ROJO = _rgb(185, 28, 28)

CODIGO_FORMULARIO = r'''Option Explicit

Private catalogo As Object          ' codigo -> nombre
Private codigos() As String         ' codigos ordenados
Private nCod As Long
Private sel As Object               ' codigos seleccionados (se conservan al filtrar)
Private cargando As Boolean

Private Sub UserForm_Initialize()
    Set sel = CreateObject("Scripting.Dictionary")
    CargarCatalogo
    Refrescar ""
    ActualizarConteo
End Sub

' Codigos: tabla PVTA_MAT. Nombres: el mas frecuente de la hoja BD_Completa.
Private Sub CargarCatalogo()
    Dim nombres As Object, v As Variant, i As Long, j As Long, c As String, tmp As String
    Set catalogo = CreateObject("Scripting.Dictionary")
    catalogo.CompareMode = 1
    Set nombres = NombresDeCedis()

    On Error Resume Next
    v = ThisWorkbook.Worksheets("PVTA_MAT VK13").ListObjects("PVTA_MAT").ListColumns("Centro").DataBodyRange.Value
    If Err.Number <> 0 Then v = Empty
    Err.Clear
    On Error GoTo 0
    If IsArray(v) Then
        For i = 1 To UBound(v, 1)
            c = Trim$(CStr(v(i, 1)))
            If Len(c) > 0 Then
                If Not catalogo.Exists(c) Then
                    If nombres.Exists(c) Then catalogo.Add c, nombres(c) Else catalogo.Add c, ""
                End If
            End If
        Next i
    End If

    nCod = catalogo.Count
    If nCod = 0 Then Exit Sub
    ReDim codigos(0 To nCod - 1)
    i = 0
    Dim k As Variant
    For Each k In catalogo.Keys
        codigos(i) = CStr(k)
        i = i + 1
    Next k
    For i = 1 To nCod - 1                 ' orden alfabetico (insercion)
        tmp = codigos(i)
        j = i - 1
        Do While j >= 0
            If StrComp(codigos(j), tmp, vbTextCompare) <= 0 Then Exit Do
            codigos(j + 1) = codigos(j)
            j = j - 1
        Loop
        codigos(j + 1) = tmp
    Next i
End Sub

Private Function NombresDeCedis() As Object
    Dim res As Object
    Set res = CreateObject("Scripting.Dictionary")
    res.CompareMode = 1
    Set NombresDeCedis = res
    NombresDesdeTraope res          ' SAP primero
    NombresDesdeBase res            ' la base solo llena los que faltan
End Function

' Descripcion del centro en CONT_COMPRA (contratos de compra)
Private Sub NombresDesdeTraope(ByVal res As Object)
    Dim lo As Object, hdr As Variant, v As Variant, i As Long, cC As Long, cN As Long
    Dim h As String, c As String, n As String
    On Error Resume Next
    Set lo = ThisWorkbook.Worksheets("CONT_COMPRA TRAOPE").ListObjects("CONT_COMPRA")
    If lo Is Nothing Then Exit Sub
    hdr = lo.HeaderRowRange.Value
    v = lo.DataBodyRange.Value
    If Err.Number <> 0 Or Not IsArray(v) Then Exit Sub
    On Error GoTo 0
    For i = 1 To UBound(hdr, 2)
        h = CStr(hdr(1, i))
        If h = "Centro" Then cC = i
        If Left$(h, 5) = "Descr" And Right$(h, 6) = "Centro" Then cN = i
    Next i
    If cC = 0 Or cN = 0 Then Exit Sub
    For i = 1 To UBound(v, 1)
        c = Trim$(CStr(v(i, cC)))
        n = Trim$(CStr(v(i, cN)))
        If Len(c) > 0 And Len(n) > 0 Then
            If Not res.Exists(c) Then res.Add c, n
        End If
    Next i
End Sub

' Nombre mas frecuente de cada centro en la hoja BD_Completa
Private Sub NombresDesdeBase(ByVal res As Object)
    Dim cuenta As Object, mejor As Object, elegido As Object, ws As Object, v As Variant
    Dim i As Long, cC As Long, cN As Long
    Dim c As String, n As String, clave As String, k As Variant
    Set cuenta = CreateObject("Scripting.Dictionary")
    Set mejor = CreateObject("Scripting.Dictionary")
    Set elegido = CreateObject("Scripting.Dictionary")

    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets("BD_Completa")
    If ws Is Nothing Then Exit Sub
    v = ws.UsedRange.Value
    If Err.Number <> 0 Or Not IsArray(v) Then Exit Sub
    On Error GoTo 0
    For i = 1 To UBound(v, 2)
        If CStr(v(1, i)) = "Centro" Then cC = i
        If CStr(v(1, i)) = "Nombre Cedis" Then cN = i
    Next i
    If cC = 0 Or cN = 0 Then Exit Sub
    For i = 2 To UBound(v, 1)
        c = Trim$(CStr(v(i, cC)))
        n = Trim$(CStr(v(i, cN)))
        If Len(c) > 0 And Len(n) > 0 Then
            clave = c & "|" & n
            cuenta(clave) = cuenta(clave) + 1
        End If
    Next i
    For Each k In cuenta.Keys
        c = Left$(CStr(k), InStr(CStr(k), "|") - 1)
        n = Mid$(CStr(k), InStr(CStr(k), "|") + 1)
        If Not res.Exists(c) Then
            If Not mejor.Exists(c) Then
                mejor(c) = cuenta(k)
                elegido(c) = n
            ElseIf cuenta(k) > mejor(c) Then
                mejor(c) = cuenta(k)
                elegido(c) = n
            End If
        End If
    Next k
    For Each k In elegido.Keys
        res.Add CStr(k), elegido(k)
    Next k
End Sub

Private Sub Refrescar(ByVal filtro As String)
    Dim i As Long, texto As String
    cargando = True
    LstCEDIS.Clear
    filtro = Trim$(filtro)
    For i = 0 To nCod - 1
        texto = codigos(i) & " " & catalogo(codigos(i))
        If Len(filtro) = 0 Or InStr(1, texto, filtro, vbTextCompare) > 0 Then
            LstCEDIS.AddItem codigos(i)
            LstCEDIS.List(LstCEDIS.ListCount - 1, 1) = catalogo(codigos(i))
            If sel.Exists(codigos(i)) Then LstCEDIS.Selected(LstCEDIS.ListCount - 1) = True
        End If
    Next i
    cargando = False
End Sub

' Copia al diccionario lo marcado en la lista visible (asi no se pierde al filtrar)
Private Sub Sincronizar()
    Dim i As Long, c As String
    For i = 0 To LstCEDIS.ListCount - 1
        c = CStr(LstCEDIS.List(i, 0))
        If LstCEDIS.Selected(i) Then
            sel(c) = True
        ElseIf sel.Exists(c) Then
            sel.Remove c
        End If
    Next i
    ActualizarConteo
End Sub

Private Sub ActualizarConteo()
    LblConteo.ForeColor = &H17335C
    LblConteo.Caption = sel.Count & IIf(sel.Count = 1, " seleccionado", " seleccionados")
End Sub

Private Sub TxtBuscar_Change()
    Sincronizar
    Refrescar TxtBuscar.Text
End Sub

Private Sub LstCEDIS_Click()
    If Not cargando Then Sincronizar
End Sub

Private Sub BtnTodos_Click()
    Dim i As Long
    For i = 0 To LstCEDIS.ListCount - 1
        LstCEDIS.Selected(i) = True
    Next i
    Sincronizar
End Sub

Private Sub BtnNinguno_Click()
    sel.RemoveAll
    Refrescar TxtBuscar.Text
    ActualizarConteo
End Sub

Private Sub BtnGenerar_Click()
    Dim k As Variant, lista As String
    Sincronizar
    If sel.Count = 0 Then
        LblConteo.ForeColor = &H1C1CB9
        LblConteo.Caption = "Selecciona al menos un CEDIS"
        Exit Sub
    End If
    If sel.Count = nCod Then
        lista = "TODOS"
    Else
        For Each k In sel.Keys
            lista = lista & CStr(k) & " "
        Next k
    End If
    Me.Hide
    EjecutarMatrizPython Trim$(lista)
    Unload Me
End Sub
'''


def crear_formulario(vbp):
    """Crea (o reemplaza) el formulario en el proyecto VBA vbp."""
    for i in range(vbp.VBComponents.Count, 0, -1):
        comp = vbp.VBComponents.Item(i)
        if comp.Name == NOMBRE_FORMULARIO:
            vbp.VBComponents.Remove(comp)

    comp = vbp.VBComponents.Add(3)              # vbext_ct_MSForm
    comp.Name = NOMBRE_FORMULARIO
    p = comp.Properties
    p.Item("Caption").Value = "CEMEX | Selector de CEDIS"
    p.Item("Width").Value = 408
    p.Item("Height").Value = 540
    p.Item("BackColor").Value = CREMA
    try:
        p.Item("StartUpPosition").Value = 1     # centrado sobre Excel
    except Exception:
        pass

    d = comp.Designer

    def ctl(tipo, nombre, izq, arr, ancho, alto):
        c = d.Controls.Add(tipo, nombre, True)
        c.Left, c.Top, c.Width, c.Height = izq, arr, ancho, alto
        return c

    def fuente(c, tam, negrita=False):
        c.Font.Name = "Segoe UI"
        c.Font.Size = tam
        c.Font.Bold = negrita

    # Franja superior con el título
    banner = ctl("Forms.Label.1", "LblBanner", 0, 0, 408, 52)
    banner.BackStyle = 1
    banner.BackColor = TERRACOTA
    banner.Caption = ""
    titulo = ctl("Forms.Label.1", "LblTitulo", 20, 14, 360, 26)
    titulo.BackStyle = 0
    titulo.ForeColor = BLANCO
    titulo.Caption = "SELECCIONA LOS CEDIS"
    fuente(titulo, 15, True)

    # Buscador
    lbl = ctl("Forms.Label.1", "LblBuscar", 20, 64, 200, 16)
    lbl.BackStyle = 0
    lbl.ForeColor = CAFE
    lbl.Caption = "Buscar"
    fuente(lbl, 9)
    txt = ctl("Forms.TextBox.1", "TxtBuscar", 20, 82, 368, 28)
    txt.BackColor = BLANCO
    txt.ForeColor = CAFE
    txt.BorderStyle = 1
    txt.BorderColor = BORDE
    txt.SpecialEffect = 0
    fuente(txt, 11)

    # Lista con casillas: código + nombre
    lst = ctl("Forms.ListBox.1", "LstCEDIS", 20, 120, 368, 300)
    lst.ColumnCount = 2
    lst.ColumnWidths = "58 pt;280 pt"
    lst.MultiSelect = 1                         # fmMultiSelectMulti
    lst.ListStyle = 1                           # con casillas
    lst.BackColor = BLANCO
    lst.ForeColor = CAFE
    lst.BorderStyle = 1
    lst.BorderColor = BORDE
    lst.SpecialEffect = 0
    fuente(lst, 10)

    # Contador y atajos
    cont = ctl("Forms.Label.1", "LblConteo", 20, 432, 190, 20)
    cont.BackStyle = 0
    cont.ForeColor = CAFE
    cont.Caption = "0 seleccionados"
    fuente(cont, 10, True)
    for nombre, texto, izq in (("BtnTodos", "Todos", 216), ("BtnNinguno", "Ninguno", 304)):
        b = ctl("Forms.CommandButton.1", nombre, izq, 428, 84, 26)
        b.Caption = texto
        b.BackColor = ARENA
        b.ForeColor = CAFE
        fuente(b, 10)

    # Botón principal
    g = ctl("Forms.CommandButton.1", "BtnGenerar", 20, 468, 368, 44)
    g.Caption = "GENERAR MATRIZ"
    g.BackColor = NARANJA
    g.ForeColor = BLANCO
    fuente(g, 13, True)

    comp.CodeModule.AddFromString(CODIGO_FORMULARIO)
    return comp
