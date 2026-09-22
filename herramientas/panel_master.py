"""
Diseño de la hoja Panel_Control del libro maestro: simple, cálido y sin texto de más.
Título, dos botones y una línea de estado de los datos. Se usa por COM (Excel de escritorio).
"""

def _rgb(r, g, b):
    return r + g * 256 + b * 65536      # Color de Excel COM (BGR)


CREMA = _rgb(255, 248, 238)
TERRACOTA = _rgb(178, 74, 36)
NARANJA = _rgb(232, 119, 46)
ARENA = _rgb(250, 222, 178)
CAFE = _rgb(92, 51, 23)
BORDE = _rgb(240, 205, 165)
BLANCO = _rgb(255, 255, 255)
VERDE = _rgb(56, 142, 60)
AMBAR = _rgb(217, 119, 6)
ROJO = _rgb(185, 28, 28)

TITULO = "CEMEX AGREGADOS  |  SISTEMA INTEGRAL DE MATRIZ DE PRECIOS Y GOBERNANZA"
SUBTITULO = "  Generador de Matriz de Precios Venta vs Costo, Validación de Márgenes y Dashboard Ejecutivo"
BOTON_PRINCIPAL = "▶  SELECCIONAR CEDIS Y GENERAR MATRIZ"
BOTON_TERMINAL = ">_  ABRIR TERMINAL Y EJECUTAR EL .BAT"


def _celda(ws, direccion, valor=None, size=10, bold=False, color=CAFE, fondo=None, align=-4108, fusionar=True):
    r = ws.Range(direccion)
    if fusionar and ":" in direccion:
        r.Merge()
    c = ws.Range(direccion.split(":")[0])
    if valor is not None:
        c.Value = valor
    r.Font.Name = "Segoe UI"
    r.Font.Size = size
    r.Font.Bold = bold
    r.Font.Color = color
    r.VerticalAlignment = -4108
    r.HorizontalAlignment = align
    if fondo is not None:
        r.Interior.Color = fondo
    return r


def _boton(ws, nombre, rango, texto, fondo, color_texto, size, macro, linea=None):
    r = ws.Range(rango)
    existente = next((s for s in ws.Shapes if s.Name == nombre), None)
    b = existente or ws.Shapes.AddShape(5, r.Left, r.Top, r.Width, r.Height)   # 5 = rectángulo redondeado
    b.Name = nombre
    b.Left, b.Top, b.Width, b.Height = r.Left, r.Top, r.Width, r.Height
    b.Fill.Solid()
    b.Fill.ForeColor.RGB = fondo
    if linea is None:
        b.Line.Visible = False
    else:
        b.Line.Visible = True
        b.Line.ForeColor.RGB = linea
        b.Line.Weight = 1.25
    try:
        b.Adjustments.Item(1).Value = 0.32      # esquinas más redondeadas
    except Exception:
        pass
    tf = b.TextFrame.Characters()
    tf.Text = texto
    tf.Font.Name = "Segoe UI"
    tf.Font.Size = size
    tf.Font.Bold = True
    tf.Font.Color = color_texto
    b.TextFrame.HorizontalAlignment = -4108
    b.TextFrame.VerticalAlignment = -4108
    b.OnAction = macro
    return b


def disenar_panel(xl, ws):
    """Rehace el diseño de Panel_Control conservando los botones (y sus macros)."""
    ws.Activate()
    ws.Cells.UnMerge()
    ws.Cells.FormatConditions.Delete()
    ws.Cells.Clear()
    conservar = {"BtnLanzadorPrincipal", "BtnTerminalBat"}
    for i in range(ws.Shapes.Count, 0, -1):
        if ws.Shapes.Item(i).Name not in conservar:
            ws.Shapes.Item(i).Delete()

    ws.Cells.Interior.Color = CREMA
    ws.Columns("A").ColumnWidth = 3.5
    for c in "BCDEFGHIJKL":
        ws.Columns(c).ColumnWidth = 12.43
    for fila, alto in {1: 10, 2: 36, 3: 22, 4: 18, 5: 22, 6: 22, 7: 22, 8: 14, 9: 22, 10: 22, 11: 22, 12: 28, 13: 10}.items():
        ws.Rows(fila).RowHeight = alto

    # Encabezado
    _celda(ws, "B2:L2", TITULO, size=15, bold=True, color=BLANCO, fondo=TERRACOTA, align=-4131)
    _celda(ws, "B3:L3", SUBTITULO, size=9.5, color=_rgb(255, 232, 205), fondo=TERRACOTA, align=-4131)
    ws.Range("B2").IndentLevel = 1

    # Botones
    _boton(ws, "BtnLanzadorPrincipal", "D5:J7", BOTON_PRINCIPAL, NARANJA, BLANCO, 13, "MostrarSelectorCEDIS")
    _boton(ws, "BtnTerminalBat", "D9:J10", BOTON_TERMINAL, ARENA, CAFE, 11, "AbrirTerminalMatriz", linea=BORDE)

    # Estado de los datos: una sola línea
    tarjeta = ws.Range("B12:L12")
    tarjeta.Interior.Color = BLANCO
    for borde in (7, 8, 9, 10):
        tarjeta.Borders(borde).LineStyle = 1
        tarjeta.Borders(borde).Color = BORDE
    _celda(ws, "B12:F12", "Datos actualizados al", size=10, color=CAFE, fondo=BLANCO, align=-4152)
    ws.Range("B12").IndentLevel = 1
    fecha = _celda(ws, "G12:H12", None, size=11, bold=True, color=CAFE, fondo=BLANCO)
    ws.Range("G12").Formula = '=IF(N12>0,DATE(INT(N12/10000),MOD(INT(N12/100),100),MOD(N12,100)),"Sin datos")'
    cod = xl.International            # tupla indexada desde 0: 18 = año, 19 = mes, 20 = día (según el idioma de Excel)
    y, m, d = cod[18], cod[19], cod[20]
    ws.Range("G12").NumberFormatLocal = f"{d}{d}/{m}{m}/{y}{y}{y}{y}"
    estado = _celda(ws, "I12:L12", None, size=10, bold=True, color=VERDE, fondo=BLANCO, align=-4131)
    ws.Range("I12").Formula = ('=IF(NOT(ISNUMBER(G12)),"⛔ Sin datos",IF(TODAY()-G12<=7,"✔ Al día",'
                               '"⚠ "&(TODAY()-G12)&" días sin cambios"))')
    ws.Range("N13").Formula = "=IF(NOT(ISNUMBER(G12)),2,IF(TODAY()-G12>7,1,0))"      # 0 al día, 1 aviso, 2 sin datos
    fc = ws.Range("I12").FormatConditions
    fc.Add(2, None, "=$N$13=1").Font.Color = AMBAR
    fc.Add(2, None, "=$N$13=2").Font.Color = ROJO

    # Auxiliar (oculto): fecha más reciente de modificación en precios y fletes (aaaammdd, texto).
    # AGGREGATE(14,6,...) ignora vacíos/errores y no requiere Ctrl+Mayús+Enter (Excel 2010+).
    ws.Range("N12").Formula = ("=IFERROR(MAX(AGGREGATE(14,6,--PVTA_MAT[Modif_Date],1),"
                               "AGGREGATE(14,6,--PVTA_FTE[Modif_Date],1)),0)")
    ws.Columns("N").Hidden = True

    ws.Tab.Color = NARANJA
    try:
        xl.ActiveWindow.DisplayGridlines = False
    except Exception:
        pass
    ws.Range("A1").Select()
