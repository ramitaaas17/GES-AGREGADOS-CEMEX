"""
=============================================================================
CEMEX AGREGADOS - GENERADOR DE PROYECTO POWER BI (.PBIP)
RÉPLICA EXACTA DE PLANTILLA SAAS DASHBOARD USER PANEL (Vector 3133054)
=============================================================================
Estructura Visual Fiel al Mockup:
1. Sidebar Izquierdo (Deep Navy #1B2E6F):
   - Buscador con contorno Cyan #00C2FF.
   - Slicer interactivo de CEDIS.
   - Menú de navegación vertical con punto activo Cyan (● Dashboard, Inbox, etc.).
   - Pie de barra institucional.
2. Barra Superior Blanca:
   - Título ejecutivo CEMEX + Perfil de Usuario con Avatar circular.
3. Card 1 (Superior / "Devices"):
   - Barra de título Deep Navy #1B2E6F.
   - 6 Circular Donut Rings SVG idénticos al mockup con arcos Cyan/Navy, porcentaje central,
     etiqueta de categoría y barra de acento inferior.
4. Card 2 (Media Izquierda / "Countries"):
   - Header Deep Navy + Red Nacional de Plantas y CEDIS.
5. Card 3 (Media Derecha / "Countries Graphic"):
   - Header Deep Navy + Gráfico de barras horizontales ordenadas de mayor a menor.
6. Card 4 (Inferior Izquierda / "Calendar"):
   - Header Deep Navy + Calendario y Segmentador de Vigencias.
7. Card 5 (Inferior Derecha / "Weekly Access"):
   - Header Deep Navy + Auditoría de Rutas con Microvisuales SVG.
=============================================================================
"""

import os
import json
import logging
import glob

logging.basicConfig(level=logging.INFO, format="%(asctime)s - %(levelname)s - %(message)s")

def crear_proyecto_pbip(output_dir, data_excel_path):
    """
    Construye el proyecto PBIP replicando fielmente la plantilla vector 3133054.
    """
    os.makedirs(output_dir, exist_ok=True)
    
    project_name = "Dashboard_CEMEX_Agregados"
    pbip_file = os.path.join(output_dir, f"{project_name}.pbip")
    sm_dir = os.path.join(output_dir, f"{project_name}.SemanticModel")
    rep_dir = os.path.join(output_dir, f"{project_name}.Report")
    
    shared_themes_dir = os.path.join(rep_dir, "StaticResources", "SharedResources", "BaseThemes")
    reg_themes_dir = os.path.join(rep_dir, "StaticResources", "RegisteredResources")
    
    os.makedirs(os.path.join(sm_dir, ".pbi"), exist_ok=True)
    os.makedirs(os.path.join(rep_dir, ".pbi"), exist_ok=True)
    os.makedirs(shared_themes_dir, exist_ok=True)
    os.makedirs(reg_themes_dir, exist_ok=True)
    
    data_path_escaped = data_excel_path.replace("\\", "\\\\")
    
    # =========================================================================
    # 1. ARCHIVO RAÍZ .pbip
    # =========================================================================
    pbip_content = {
        "version": "1.0",
        "artifacts": [{"report": {"path": f"{project_name}.Report"}}],
        "settings": {"enableAutoRecovery": True}
    }
    with open(pbip_file, "w", encoding="utf-8") as f:
        json.dump(pbip_content, f, indent=2, ensure_ascii=False)
        
    # =========================================================================
    # 2. DEFINITIONS (.pbir y .pbism)
    # =========================================================================
    pbir_content = {
        "version": "1.0",
        "datasetReference": {
            "byPath": {"path": f"../{project_name}.SemanticModel"},
            "byConnection": None
        }
    }
    with open(os.path.join(rep_dir, "definition.pbir"), "w", encoding="utf-8") as f:
        json.dump(pbir_content, f, indent=2)

    pbism_content = {"version": "1.0"}
    with open(os.path.join(sm_dir, "definition.pbism"), "w", encoding="utf-8") as f:
        json.dump(pbism_content, f, indent=2)

    local_settings = {"version": "1.0"}
    with open(os.path.join(sm_dir, ".pbi", "localSettings.json"), "w", encoding="utf-8") as f:
        json.dump(local_settings, f, indent=2)
    with open(os.path.join(rep_dir, ".pbi", "localSettings.json"), "w", encoding="utf-8") as f:
        json.dump(local_settings, f, indent=2)

    # =========================================================================
    # 3. TEMA SAAS USER PANEL (#1B2E6F, #00C2FF, #F0F3F8)
    # =========================================================================
    base_theme_content = {
        "name": "CY23SU04",
        "version": "5.46",
        "type": 2
    }
    with open(os.path.join(shared_themes_dir, "CY23SU04.json"), "w", encoding="utf-8") as f:
        json.dump(base_theme_content, f, indent=2)

    theme_registered_name = "CEMEX_SaaS_UserPanel.json"
    cemex_theme = {
        "name": "CEMEX_SaaS_UserPanel",
        "dataColors": [
            "#1B2E6F",  # Deep Navy Principal
            "#00C2FF",  # Cyan Vibrante / Sky Blue (Acento activo)
            "#00A859",  # Verde Vertua
            "#3B5CA8",  # Azul Secundario
            "#D32027",  # Rojo Alerta
            "#7895D1",  # Azul Claro Menú
            "#94A3B8",  # Gris Neutro
            "#1E293B"   # Texto Oscuro
        ],
        "background": "#F0F3F8",
        "foreground": "#1E293B",
        "tableAccent": "#1B2E6F",
        "good": "#00A859",
        "neutral": "#00C2FF",
        "bad": "#D32027",
        "visualStyles": {
            "*": {
                "*": {
                    "fontFamily": [{"fontFamily": "'Segoe UI', 'Lato', Arial, sans-serif"}],
                    "background": [{"color": {"solid": {"color": "#FFFFFF"}}, "transparency": 0}],
                    "border": [{"show": True, "color": {"solid": {"color": "#DDE4F0"}}, "radius": 4}],
                    "dropShadow": [{"show": False}],
                    "visualHeader": [{"background": {"solid": {"color": "#FFFFFF"}}, "border": {"solid": {"color": "#FFFFFF"}}}]
                }
            },
            "page": {
                "*": {
                    "background": [{"color": {"solid": {"color": "#F0F3F8"}}, "transparency": 0}]
                }
            }
        }
    }
    with open(os.path.join(reg_themes_dir, theme_registered_name), "w", encoding="utf-8") as f:
        json.dump(cemex_theme, f, indent=2, ensure_ascii=False)

    # =========================================================================
    # 4. SEMANTIC MODEL (model.bim) CON LOS 6 DONUT RINGS SVG EXACTOS
    # =========================================================================
    def build_svg_ring_measure(measure_ref, label_text, default_pct=0.5):
        return (
            f"VAR _pct = {measure_ref}\n"
            f"VAR _pctClamped = MIN(MAX(IF(ISBLANK(_pct), {default_pct}, _pct), 0), 1)\n"
            f"VAR _circumference = 175.93\n"
            f"VAR _dashoffset = _circumference * (1 - _pctClamped)\n"
            f"VAR _label = FORMAT(_pctClamped, \"0%\")\n"
            f"RETURN\n"
            f"\"data:image/svg+xml;utf8,<svg xmlns='http://www.w3.org/2000/svg' width='140' height='130' viewBox='0 0 140 130'>\" &\n"
            f"\"<circle cx='70' cy='52' r='28' fill='none' stroke='#1B2E6F' stroke-width='7' />\" &\n"
            f"\"<circle cx='70' cy='52' r='28' fill='none' stroke='#00C2FF' stroke-width='7' stroke-dasharray='175.93' stroke-dashoffset='\" & _dashoffset & \"' stroke-linecap='round' transform='rotate(-90 70 52)' />\" &\n"
            f"\"<text x='70' y='58' font-family='Segoe UI, Lato, Arial' font-size='15' font-weight='700' fill='#1B2E6F' text-anchor='middle'>\" & _label & \"</text>\" &\n"
            f"\"<text x='70' y='98' font-family='Segoe UI, Lato, Arial' font-size='9.5' font-weight='600' fill='#1B2E6F' text-anchor='middle'>{label_text}</text>\" &\n"
            f"\"<line x1='50' y1='106' x2='90' y2='106' stroke='#00C2FF' stroke-width='2' />\" &\n"
            f"\"</svg>\""
        )

    model_bim = {
        "name": "SemanticModel",
        "compatibilityLevel": 1550,
        "model": {
            "culture": "es-MX",
            "dataAccessOptions": {"legacyRedirects": True, "returnErrorValuesAsNull": True},
            "defaultPowerBIDataSourceVersion": "powerBI_V3",
            "sourceQueryCulture": "es-MX",
            "tables": [
                {
                    "name": "Data_Agregados",
                    "columns": [
                        {"name": "Concat1", "dataType": "string", "sourceColumn": "Concat1"},
                        {"name": "Concat2", "dataType": "string", "sourceColumn": "Concat2"},
                        {"name": "Sociedad", "dataType": "string", "sourceColumn": "Sociedad"},
                        {"name": "Ship From", "dataType": "string", "sourceColumn": "Ship From"},
                        {"name": "Nombre SF", "dataType": "string", "sourceColumn": "Nombre SF"},
                        {"name": "Centro", "dataType": "string", "sourceColumn": "Centro"},
                        {"name": "Desc. Centro", "dataType": "string", "sourceColumn": "Desc. Centro"},
                        {"name": "Destino", "dataType": "string", "sourceColumn": "Destino"},
                        {"name": "Nombre Destino", "dataType": "string", "sourceColumn": "Nombre Destino"},
                        {"name": "Material", "dataType": "string", "sourceColumn": "Material"},
                        {"name": "Denominación", "dataType": "string", "sourceColumn": "Denominación"},
                        {"name": "PV", "dataType": "double", "sourceColumn": "PV", "formatString": "#,##0.000"},
                        {"name": "Inicio Vigencia", "dataType": "string", "sourceColumn": "Inicio Vigencia"},
                        {"name": "Fin Vigencia", "dataType": "string", "sourceColumn": "Fin Vigencia"},
                        {"name": "Modalidad Venta", "dataType": "string", "sourceColumn": "Modalidad Venta"},
                        {"name": "Clase Cond. MP", "dataType": "string", "sourceColumn": "Clase Cond. MP"},
                        {"name": "Importe MP", "dataType": "double", "sourceColumn": "Importe MP", "formatString": "$#,##0.00"},
                        {"name": "UM Venta", "dataType": "string", "sourceColumn": "UM Venta"},
                        {"name": "Importe Flete", "dataType": "double", "sourceColumn": "Importe Flete", "formatString": "$#,##0.00"},
                        {"name": "No. Contrato Compra", "dataType": "string", "sourceColumn": "No. Contrato Compra"},
                        {"name": "Inicio Vigencia Compra", "dataType": "string", "sourceColumn": "Inicio Vigencia Compra"},
                        {"name": "Fin Vigencia Compra", "dataType": "string", "sourceColumn": "Fin Vigencia Compra"},
                        {"name": "Costo Total TRAOPE", "dataType": "double", "sourceColumn": "Costo Total TRAOPE", "formatString": "$#,##0.00"},
                        {"name": "Flete Compra", "dataType": "double", "sourceColumn": "Flete Compra", "formatString": "$#,##0.00"},
                        {"name": "MP Compra (Costo Material)", "dataType": "double", "sourceColumn": "MP Compra (Costo Material)", "formatString": "$#,##0.00"},
                        {"name": "Margen Material (MOP %)", "dataType": "double", "sourceColumn": "Margen Material (MOP %)", "formatString": "0.0%"},
                        {"name": "UM Costo", "dataType": "string", "sourceColumn": "UM Costo"},
                        {"name": "Tipo Operación", "dataType": "string", "sourceColumn": "Tipo Operación"},
                        {"name": "Nivel Autorización / Alerta", "dataType": "string", "sourceColumn": "Nivel Autorización / Alerta"},
                        {"name": "Validacion 1", "dataType": "double", "sourceColumn": "Validacion 1", "formatString": "$#,##0.00"},
                        {"name": "Validacion 2", "dataType": "double", "sourceColumn": "Validacion 2", "formatString": "$#,##0.00"},
                        {"name": "Semaforo", "dataType": "string", "sourceColumn": "Semaforo"},
                        {"name": "No. Contrato Venta", "dataType": "string", "sourceColumn": "No. Contrato Venta"},
                        {"name": "UM Contrato", "dataType": "string", "sourceColumn": "UM Contrato"},
                        {"name": "Precio Contrato", "dataType": "double", "sourceColumn": "Precio Contrato", "formatString": "$#,##0.00"}
                    ],
                    "measures": [
                        {"name": "Total Rutas", "expression": "COUNTROWS('Data_Agregados')", "formatString": "#,##0"},
                        {"name": "Rutas Terceros", "expression": "CALCULATE(COUNTROWS('Data_Agregados'), 'Data_Agregados'[Sociedad] <> \"7100\")", "formatString": "#,##0"},
                        {"name": "Rutas Filial", "expression": "CALCULATE(COUNTROWS('Data_Agregados'), 'Data_Agregados'[Sociedad] = \"7100\")", "formatString": "#,##0"},
                        {"name": "Rutas Trading 7180", "expression": "CALCULATE(COUNTROWS('Data_Agregados'), 'Data_Agregados'[Sociedad] = \"7180\")", "formatString": "#,##0"},
                        {"name": "Rutas Trading 7277", "expression": "CALCULATE(COUNTROWS('Data_Agregados'), 'Data_Agregados'[Sociedad] = \"7277\")", "formatString": "#,##0"},
                        {"name": "Rutas Champion", "expression": "CALCULATE(COUNTROWS('Data_Agregados'), SEARCH(\"Champion\", 'Data_Agregados'[Nivel Autorización / Alerta], 1, 0) > 0)", "formatString": "#,##0"},
                        {"name": "Rutas Regional", "expression": "CALCULATE(COUNTROWS('Data_Agregados'), SEARCH(\"Regional\", 'Data_Agregados'[Nivel Autorización / Alerta], 1, 0) > 0)", "formatString": "#,##0"},
                        {"name": "Alertas Nacionales", "expression": "CALCULATE(COUNTROWS('Data_Agregados'), SEARCH(\"Alerta\", 'Data_Agregados'[Nivel Autorización / Alerta], 1, 0) > 0 || SEARCH(\"Nacional\", 'Data_Agregados'[Nivel Autorización / Alerta], 1, 0) > 0)", "formatString": "#,##0"},
                        {"name": "MOP Promedio Terceros", "expression": "CALCULATE(AVERAGE('Data_Agregados'[Margen Material (MOP %)]), 'Data_Agregados'[Sociedad] <> \"7100\")", "formatString": "0.0%"},
                        {"name": "Pct Champion", "expression": "DIVIDE([Rutas Champion], [Rutas Terceros], 0.75)", "formatString": "0.0%"},
                        {"name": "Pct Regional", "expression": "DIVIDE([Rutas Regional], [Rutas Terceros], 0.50)", "formatString": "0.0%"},
                        {"name": "Pct Alertas", "expression": "DIVIDE([Alertas Nacionales], [Rutas Terceros], 0.15)", "formatString": "0.0%"},
                        {"name": "Pct Filial", "expression": "DIVIDE([Rutas Filial], [Total Rutas], 0.30)", "formatString": "0.0%"},
                        {"name": "Pct Trading 7180", "expression": "DIVIDE([Rutas Trading 7180], [Total Rutas], 0.25)", "formatString": "0.0%"},
                        {"name": "Pct Trading 7277", "expression": "DIVIDE([Rutas Trading 7277], [Total Rutas], 0.10)", "formatString": "0.0%"},
                        
                        # =====================================================
                        # 6 CIRCULAR DONUT RINGS SVG IDÉNTICOS AL MOCKUP
                        # =====================================================
                        {"name": "SVG Ring Champion", "expression": build_svg_ring_measure("[Pct Champion]", "Champion (>8%)", 0.75), "dataCategory": "ImageUrl"},
                        {"name": "SVG Ring Regional", "expression": build_svg_ring_measure("[Pct Regional]", "Regional (5-8%)", 0.50), "dataCategory": "ImageUrl"},
                        {"name": "SVG Ring Alerta", "expression": build_svg_ring_measure("[Pct Alertas]", "Alerta (<5% VP)", 0.15), "dataCategory": "ImageUrl"},
                        {"name": "SVG Ring Filial", "expression": build_svg_ring_measure("[Pct Filial]", "Filial (7100)", 0.30), "dataCategory": "ImageUrl"},
                        {"name": "SVG Ring Trading7180", "expression": build_svg_ring_measure("[Pct Trading 7180]", "Trading 7180", 0.25), "dataCategory": "ImageUrl"},
                        {"name": "SVG Ring Trading7277", "expression": build_svg_ring_measure("[Pct Trading 7277]", "Trading 7277", 0.10), "dataCategory": "ImageUrl"},

                        # Microvisuales para tablas
                        {
                            "name": "SVG MOP Bullet Bar",
                            "expression": (
                                "VAR _MOP = AVERAGE('Data_Agregados'[Margen Material (MOP %)])\n"
                                "VAR _Val = IF(ISBLANK(_MOP), 0, _MOP)\n"
                                "VAR _PctClamped = MIN(MAX(_Val, 0), 0.25)\n"
                                "VAR _BarWidth = _PctClamped * 400\n"
                                "VAR _Color = SWITCH(TRUE(), _Val < 0.05, \"#D32027\", _Val > 0.08, \"#00A859\", \"#1B2E6F\")\n"
                                "VAR _Target5 = 0.05 * 400\n"
                                "VAR _Target8 = 0.08 * 400\n"
                                "VAR _TextLabel = FORMAT(_Val, \"0.0%\")\n"
                                "RETURN\n"
                                "IF(ISBLANK(_MOP), BLANK(),\n"
                                "\"data:image/svg+xml;utf8,<svg xmlns='http://www.w3.org/2000/svg' width='130' height='22' viewBox='0 0 130 22'>\" &\n"
                                "\"<rect x='0' y='5' width='100' height='12' rx='2' fill='#E2E8F4'/>\" &\n"
                                "\"<rect x='0' y='5' width='\" & _BarWidth & \"' height='12' rx='2' fill='\" & _Color & \"'/>\" &\n"
                                "\"<line x1='\" & _Target5 & \"' y1='3' x2='\" & _Target5 & \"' y2='19' stroke='#D32027' stroke-width='1.5' stroke-dasharray='2,1'/>\" &\n"
                                "\"<line x1='\" & _Target8 & \"' y1='3' x2='\" & _Target8 & \"' y2='19' stroke='#00A859' stroke-width='1.5'/>\" &\n"
                                "\"<text x='105' y='15' font-family='Segoe UI, Arial' font-size='10' font-weight='600' fill='\" & _Color & \"'>\" & _TextLabel & \"</text>\" &\n"
                                "\"</svg>\")"
                            ),
                            "dataCategory": "ImageUrl"
                        },
                        {
                            "name": "SVG Status Pill",
                            "expression": (
                                "VAR _Nivel = SELECTEDVALUE('Data_Agregados'[Nivel Autorización / Alerta], MAX('Data_Agregados'[Nivel Autorización / Alerta]))\n"
                                "VAR _Color = SWITCH(TRUE(), SEARCH(\"Champion\", _Nivel, 1, 0) > 0, \"#00A859\", SEARCH(\"Regional\", _Nivel, 1, 0) > 0, \"#1B2E6F\", SEARCH(\"Alerta\", _Nivel, 1, 0) > 0 || SEARCH(\"Nacional\", _Nivel, 1, 0) > 0, \"#D32027\", \"#64748B\")\n"
                                "VAR _Bg = SWITCH(TRUE(), SEARCH(\"Champion\", _Nivel, 1, 0) > 0, \"#F0FDF4\", SEARCH(\"Regional\", _Nivel, 1, 0) > 0, \"#F0F7FF\", SEARCH(\"Alerta\", _Nivel, 1, 0) > 0 || SEARCH(\"Nacional\", _Nivel, 1, 0) > 0, \"#FEF2F2\", \"#F8FAFC\")\n"
                                "VAR _Short = SWITCH(TRUE(), SEARCH(\"Champion\", _Nivel, 1, 0) > 0, \"Champion (>8%)\", SEARCH(\"Regional\", _Nivel, 1, 0) > 0, \"Regional (5-8%)\", SEARCH(\"Alerta\", _Nivel, 1, 0) > 0 || SEARCH(\"Nacional\", _Nivel, 1, 0) > 0, \"Alerta (<5%)\", \"Filial (N/A)\")\n"
                                "RETURN\n"
                                "IF(ISBLANK(_Nivel), BLANK(),\n"
                                "\"data:image/svg+xml;utf8,<svg xmlns='http://www.w3.org/2000/svg' width='105' height='22' viewBox='0 0 105 22'>\" &\n"
                                "\"<rect x='1' y='1' width='103' height='20' rx='10' fill='\" & _Bg & \"' stroke='\" & _Color & \"' stroke-width='1'/>\" &\n"
                                "\"<text x='52' y='14' font-family='Segoe UI, Arial' font-size='9.5' font-weight='600' fill='\" & _Color & \"' text-anchor='middle'>\" & _Short & \"</text>\" &\n"
                                "\"</svg>\")"
                            ),
                            "dataCategory": "ImageUrl"
                        }
                    ],
                    "partitions": [
                        {
                            "name": "Data_Agregados",
                            "mode": "import",
                            "source": {
                                "type": "m",
                                "expression": [
                                    "let",
                                    f"    RutaArchivo = \"{data_path_escaped}\",",
                                    "    Origen = Excel.Workbook(File.Contents(RutaArchivo), null, true),",
                                    "    Data_Sheet = Origen{[Item=\"Data_Agregados\",Kind=\"Sheet\"]}[Data],",
                                    "    Encabezados = Table.PromoteHeaders(Data_Sheet, [PromoteAllScalars=true]),",
                                    "    TiposCambiados = Table.TransformColumnTypes(Encabezados,{",
                                    "        {\"Concat1\", type text}, {\"Concat2\", type text}, {\"Sociedad\", type text}, {\"Ship From\", type text}, {\"Nombre SF\", type text},",
                                    "        {\"Centro\", type text}, {\"Desc. Centro\", type text}, {\"Destino\", type text}, {\"Nombre Destino\", type text}, {\"Material\", type text}, {\"Denominación\", type text},",
                                    "        {\"PV\", type number}, {\"Inicio Vigencia\", type text}, {\"Fin Vigencia\", type text}, {\"Modalidad Venta\", type text}, {\"Clase Cond. MP\", type text},",
                                    "        {\"Importe MP\", Currency.Type}, {\"UM Venta\", type text}, {\"Importe Flete\", Currency.Type},",
                                    "        {\"No. Contrato Compra\", type text}, {\"Inicio Vigencia Compra\", type text}, {\"Fin Vigencia Compra\", type text},",
                                    "        {\"Costo Total TRAOPE\", Currency.Type}, {\"Flete Compra\", Currency.Type}, {\"MP Compra (Costo Material)\", Currency.Type},",
                                    "        {\"Margen Material (MOP %)\", Percentage.Type}, {\"UM Costo\", type text}, {\"Tipo Operación\", type text}, {\"Nivel Autorización / Alerta\", type text},",
                                    "        {\"Validacion 1\", Currency.Type}, {\"Validacion 2\", Currency.Type}, {\"Semaforo\", type text},",
                                    "        {\"No. Contrato Venta\", type text}, {\"UM Contrato\", type text}, {\"Precio Contrato\", Currency.Type}",
                                    "    }, \"es-MX\")",
                                    "in",
                                    "    TiposCambiados"
                                ]
                            }
                        }
                    ]
                }
            ]
        }
    }
    with open(os.path.join(sm_dir, "model.bim"), "w", encoding="utf-8") as f:
        json.dump(model_bim, f, indent=2, ensure_ascii=False)

    # =========================================================================
    # 5. REPORTE VISUAL (report.json) - RÉPLICA EXACTA DEL MOCKUP 3133054
    # =========================================================================
    def card_header_bar(name, x, y, width, height, z, title_text):
        return [
            {
                "x": x, "y": y, "z": z, "width": width, "height": height,
                "config": json.dumps({
                    "name": f"{name}_Bg",
                    "layouts": [{"id": 0, "position": {"x": x, "y": y, "z": z, "width": width, "height": height}}],
                    "singleVisual": {
                        "visualType": "shape",
                        "objects": {
                            "shape": [{"properties": {"tileShape": {"expr": {"Literal": {"Value": "'rectangle'"}}}}}],
                            "fill": [{"properties": {"fillColor": {"solid": {"color": {"expr": {"Literal": {"Value": "'#1B2E6F'"}}}}}}}],
                            "line": [{"properties": {"lineColor": {"solid": {"color": {"expr": {"Literal": {"Value": "'#1B2E6F'"}}}}}}}]
                        }
                    }
                })
            },
            {
                "x": x + 10, "y": y + 3, "z": z + 1, "width": width - 20, "height": height - 4,
                "config": json.dumps({
                    "name": f"{name}_Text",
                    "layouts": [{"id": 0, "position": {"x": x + 10, "y": y + 3, "z": z + 1, "width": width - 20, "height": height - 4}}],
                    "singleVisual": {
                        "visualType": "textbox",
                        "objects": {
                            "general": [{
                                "properties": {
                                    "paragraphs": [
                                        {"textRuns": [{"value": title_text, "textStyle": {"fontWeight": "bold", "fontSize": "9pt", "color": "#FFFFFF"}}]}
                                    ]
                                }
                            }],
                            "background": [{"properties": {"show": {"expr": {"Literal": {"Value": "false"}}}}}]
                        }
                    }
                })
            }
        ]

    def svg_ring_card(name, x, y, width, height, z, query_ref):
        return {
            "x": x, "y": y, "z": z, "width": width, "height": height,
            "config": json.dumps({
                "name": name,
                "layouts": [{"id": 0, "position": {"x": x, "y": y, "z": z, "width": width, "height": height}}],
                "singleVisual": {
                    "visualType": "card",
                    "projections": {"Values": [{"queryRef": query_ref}]},
                    "prototypeQuery": {
                        "Version": 2,
                        "From": [{"Name": "d", "Entity": "Data_Agregados", "Type": 0}],
                        "Select": [{"Measure": {"Expression": {"SourceRef": {"Source": "d"}}, "Property": query_ref.split(".")[1]}, "Name": query_ref}]
                    },
                    "objects": {
                        "labels": [{"properties": {"imageHeight": {"expr": {"Literal": {"Value": "115L"}}}, "imageWidth": {"expr": {"Literal": {"Value": "135L"}}}}}],
                        "categoryLabels": [{"properties": {"show": {"expr": {"Literal": {"Value": "false"}}}}}]
                    },
                    "vcObjects": {
                        "border": [{"properties": {"show": {"expr": {"Literal": {"Value": "false"}}}}}],
                        "background": [{"properties": {"show": {"expr": {"Literal": {"Value": "true"}}}, "color": {"solid": {"color": {"expr": {"Literal": {"Value": "'#FFFFFF'"}}}}}, "transparency": {"expr": {"Literal": {"Value": "0D"}}}}}],
                        "dropShadow": [{"properties": {"show": {"expr": {"Literal": {"Value": "false"}}}}}]
                    }
                }
            })
        }

    report_json = {
        "config": json.dumps({
            "version": "5.48",
            "themeCollection": {
                "baseTheme": {"name": "CY23SU04", "version": "5.46", "type": 2},
                "customTheme": {"name": theme_registered_name, "version": "5.46", "type": 1}
            },
            "activeSectionIndex": 0,
            "defaultDrillFilterOtherVisuals": True,
            "settings": {
                "useNewFilterPaneExperience": True,
                "allowChangeFilterTypes": True,
                "useStylableVisualContainerHeader": True,
                "useEnhancedTooltips": True,
                "exportDataMode": 1
            }
        }),
        "layoutOptimization": 0,
        "resourcePackages": [
            {
                "resourcePackage": {
                    "disabled": False,
                    "items": [{"name": "CY23SU04", "path": "BaseThemes/CY23SU04.json", "type": 202}],
                    "name": "SharedResources", "type": 2
                }
            },
            {
                "resourcePackage": {
                    "disabled": False,
                    "items": [{"name": theme_registered_name, "path": theme_registered_name, "type": 201}],
                    "name": "RegisteredResources", "type": 1
                }
            }
        ],
        "sections": [
            # =================================================================
            # PÁGINA 1: RÉPLICA EXACTA DEL MOCKUP SAAS USER PANEL
            # =================================================================
            {
                "name": "SaaSUserPanel",
                "displayName": "📊 Dashboard",
                "filters": "[]",
                "ordinal": 0,
                "width": 1280,
                "height": 720,
                "config": json.dumps({
                    "objects": {
                        "background": [{
                            "properties": {
                                "color": {"solid": {"color": {"expr": {"Literal": {"Value": "'#F0F3F8'"}}}}},
                                "transparency": {"expr": {"Literal": {"Value": "0D"}}}
                            }
                        }]
                    }
                }),
                "visualContainers": [
                    # ---------------------------------------------------------
                    # 1. SIDEBAR IZQUIERDO (DEEP NAVY #1B2E6F)
                    # ---------------------------------------------------------
                    {
                        "x": 0, "y": 0, "z": 0, "width": 216, "height": 720,
                        "config": json.dumps({
                            "name": "Sidebar_Background",
                            "layouts": [{"id": 0, "position": {"x": 0, "y": 0, "z": 0, "width": 216, "height": 720}}],
                            "singleVisual": {
                                "visualType": "shape",
                                "objects": {
                                    "shape": [{"properties": {"tileShape": {"expr": {"Literal": {"Value": "'rectangle'"}}}}}],
                                    "fill": [{"properties": {"fillColor": {"solid": {"color": {"expr": {"Literal": {"Value": "'#1B2E6F'"}}}}}}}],
                                    "line": [{"properties": {"lineColor": {"solid": {"color": {"expr": {"Literal": {"Value": "'#1B2E6F'"}}}}}}}]
                                }
                            }
                        })
                    },
                    # 1.1 SEARCH BOX EN EL SIDEBAR (CON CONTORNO CYAN #00C2FF)
                    {
                        "x": 16, "y": 18, "z": 1, "width": 184, "height": 34,
                        "config": json.dumps({
                            "name": "Sidebar_Search_Box",
                            "layouts": [{"id": 0, "position": {"x": 16, "y": 18, "z": 1, "width": 184, "height": 34}}],
                            "singleVisual": {
                                "visualType": "shape",
                                "objects": {
                                    "shape": [{"properties": {"tileShape": {"expr": {"Literal": {"Value": "'rectangle'"}}}}}],
                                    "fill": [{"properties": {"fillColor": {"solid": {"color": {"expr": {"Literal": {"Value": "'#162559'"}}}}}}}],
                                    "line": [{"properties": {"lineColor": {"solid": {"color": {"expr": {"Literal": {"Value": "'#00C2FF'"}}}}}, "weight": {"expr": {"Literal": {"Value": "1.5D"}}}}}],
                                    "roundCorners": [{"properties": {"radius": {"expr": {"Literal": {"Value": "3D"}}}}}]
                                }
                            }
                        })
                    },
                    {
                        "x": 24, "y": 24, "z": 2, "width": 168, "height": 22,
                        "config": json.dumps({
                            "name": "Sidebar_Search_Placeholder",
                            "layouts": [{"id": 0, "position": {"x": 24, "y": 24, "z": 2, "width": 168, "height": 22}}],
                            "singleVisual": {
                                "visualType": "textbox",
                                "objects": {
                                    "general": [{
                                        "properties": {
                                            "paragraphs": [
                                                {"textRuns": [{"value": "🔍  Search...", "textStyle": {"fontWeight": "normal", "fontSize": "9pt", "color": "#7895D1"}}]}
                                            ]
                                        }
                                    }],
                                    "background": [{"properties": {"show": {"expr": {"Literal": {"Value": "false"}}}}}]
                                }
                            }
                        })
                    },
                    # 1.2 SLICER DE CENTROS / CEDIS (INTEGRADO EN EL SIDEBAR)
                    {
                        "x": 16, "y": 62, "z": 3, "width": 184, "height": 160,
                        "config": json.dumps({
                            "name": "Sidebar_Slicer_Centro",
                            "layouts": [{"id": 0, "position": {"x": 16, "y": 62, "z": 3, "width": 184, "height": 160}}],
                            "singleVisual": {
                                "visualType": "slicer",
                                "projections": {"Values": [{"queryRef": "Data_Agregados.Centro"}]},
                                "prototypeQuery": {
                                    "Version": 2,
                                    "From": [{"Name": "d", "Entity": "Data_Agregados", "Type": 0}],
                                    "Select": [{"Column": {"Expression": {"SourceRef": {"Source": "d"}}, "Property": "Centro"}, "Name": "Data_Agregados.Centro"}]
                                },
                                "objects": {
                                    "header": [{"properties": {"show": {"expr": {"Literal": {"Value": "false"}}}}}],
                                    "selection": [{"properties": {"selectAllCheckboxEnabled": {"expr": {"Literal": {"Value": "true"}}}}}]
                                },
                                "vcObjects": {
                                    "background": [{"properties": {"show": {"expr": {"Literal": {"Value": "true"}}}, "color": {"solid": {"color": {"expr": {"Literal": {"Value": "'#162559'"}}}}}, "transparency": {"expr": {"Literal": {"Value": "0D"}}}}}],
                                    "border": [{"properties": {"show": {"expr": {"Literal": {"Value": "true"}}}, "color": {"solid": {"color": {"expr": {"Literal": {"Value": "'#2A4287'"}}}}}, "radius": {"expr": {"Literal": {"Value": "3D"}}}}}],
                                    "title": [{
                                        "properties": {
                                            "text": {"expr": {"Literal": {"Value": "'CEDIS / PLANTAS'"}}},
                                            "show": {"expr": {"Literal": {"Value": "true"}}},
                                            "fontColor": {"solid": {"color": {"expr": {"Literal": {"Value": "'#00C2FF'"}}}}},
                                            "fontSize": {"expr": {"Literal": {"Value": "8D"}}},
                                            "fontFamily": {"expr": {"Literal": {"Value": "'''Segoe UI Bold'', wf_segoe-ui_bold, Arial, sans-serif'"}}}
                                        }
                                    }]
                                }
                            }
                        })
                    },
                    # 1.3 MENÚ VERTICAL DE NAVEGACIÓN (EXACTO AL MOCKUP CON ICONOS Y PUNTO ACTIVO)
                    {
                        "x": 16, "y": 236, "z": 4, "width": 184, "height": 420,
                        "config": json.dumps({
                            "name": "Sidebar_Nav_Menu",
                            "layouts": [{"id": 0, "position": {"x": 16, "y": 236, "z": 4, "width": 184, "height": 420}}],
                            "singleVisual": {
                                "visualType": "textbox",
                                "objects": {
                                    "general": [{
                                        "properties": {
                                            "paragraphs": [
                                                {"textRuns": [{"value": "●  ⊞  Dashboard\n\n", "textStyle": {"fontWeight": "bold", "fontSize": "10pt", "color": "#00C2FF"}}]},
                                                {"textRuns": [{"value": "    ✉  Inbox\n\n", "textStyle": {"fontWeight": "normal", "fontSize": "9.5pt", "color": "#7895D1"}}]},
                                                {"textRuns": [{"value": "    ✏  Posts / Materiales\n\n", "textStyle": {"fontWeight": "normal", "fontSize": "9.5pt", "color": "#7895D1"}}]},
                                                {"textRuns": [{"value": "    📄  Pages / Centros\n\n", "textStyle": {"fontWeight": "normal", "fontSize": "9.5pt", "color": "#7895D1"}}]},
                                                {"textRuns": [{"value": "    ⊞  Widgets / Gobernanza\n\n", "textStyle": {"fontWeight": "normal", "fontSize": "9.5pt", "color": "#7895D1"}}]},
                                                {"textRuns": [{"value": "    ⚙  Settings\n", "textStyle": {"fontWeight": "normal", "fontSize": "9.5pt", "color": "#7895D1"}}]}
                                            ]
                                        }
                                    }],
                                    "background": [{"properties": {"show": {"expr": {"Literal": {"Value": "false"}}}}}]
                                }
                            }
                        })
                    },
                    # 1.4 FOOTER DEL SIDEBAR
                    {
                        "x": 16, "y": 665, "z": 5, "width": 184, "height": 45,
                        "config": json.dumps({
                            "name": "Sidebar_Footer_Brand",
                            "layouts": [{"id": 0, "position": {"x": 16, "y": 665, "z": 5, "width": 184, "height": 45}}],
                            "singleVisual": {
                                "visualType": "textbox",
                                "objects": {
                                    "general": [{
                                        "properties": {
                                            "paragraphs": [
                                                {"textRuns": [{"value": "CEMEX AGREGADOS\n", "textStyle": {"fontWeight": "bold", "fontSize": "7.5pt", "color": "#00C2FF"}}]},
                                                {"textRuns": [{"value": "457 Rutas Vigentes | v2.6", "textStyle": {"fontWeight": "normal", "fontSize": "7pt", "color": "#7895D1"}}]}
                                            ]
                                        }
                                    }],
                                    "background": [{"properties": {"show": {"expr": {"Literal": {"Value": "false"}}}}}]
                                }
                            }
                        })
                    },
                    # ---------------------------------------------------------
                    # 2. TOP HEADER BAR (BLANCO CON USER PROFILE / AVATAR)
                    # ---------------------------------------------------------
                    {
                        "x": 216, "y": 0, "z": 0, "width": 1064, "height": 48,
                        "config": json.dumps({
                            "name": "Top_Header_Background",
                            "layouts": [{"id": 0, "position": {"x": 216, "y": 0, "z": 0, "width": 1064, "height": 48}}],
                            "singleVisual": {
                                "visualType": "shape",
                                "objects": {
                                    "shape": [{"properties": {"tileShape": {"expr": {"Literal": {"Value": "'rectangle'"}}}}}],
                                    "fill": [{"properties": {"fillColor": {"solid": {"color": {"expr": {"Literal": {"Value": "'#FFFFFF'"}}}}}}}],
                                    "line": [{"properties": {"lineColor": {"solid": {"color": {"expr": {"Literal": {"Value": "'#DDE4F0'"}}}}}}}]
                                }
                            }
                        })
                    },
                    {
                        "x": 234, "y": 12, "z": 1, "width": 600, "height": 26,
                        "config": json.dumps({
                            "name": "Top_Header_App_Title",
                            "layouts": [{"id": 0, "position": {"x": 234, "y": 12, "z": 1, "width": 600, "height": 26}}],
                            "singleVisual": {
                                "visualType": "textbox",
                                "objects": {
                                    "general": [{
                                        "properties": {
                                            "paragraphs": [
                                                {"textRuns": [{"value": "CEMEX AGREGADOS  |  ", "textStyle": {"fontWeight": "bold", "fontSize": "11pt", "color": "#1B2E6F"}}]},
                                                {"textRuns": [{"value": "Panel Ejecutivo de Control Comercial", "textStyle": {"fontWeight": "normal", "fontSize": "11pt", "color": "#64748B"}}]}
                                            ]
                                        }
                                    }],
                                    "background": [{"properties": {"show": {"expr": {"Literal": {"Value": "false"}}}}}]
                                }
                            }
                        })
                    },
                    # USER PROFILE / AVATAR A LA DERECHA (TAL CUAL EL MOCKUP)
                    {
                        "x": 1060, "y": 8, "z": 1, "width": 200, "height": 34,
                        "config": json.dumps({
                            "name": "Top_Header_User_Profile",
                            "layouts": [{"id": 0, "position": {"x": 1060, "y": 8, "z": 1, "width": 200, "height": 34}}],
                            "singleVisual": {
                                "visualType": "textbox",
                                "objects": {
                                    "general": [{
                                        "properties": {
                                            "paragraphs": [
                                                {
                                                    "textRuns": [
                                                        {"value": "User Name  ", "textStyle": {"fontWeight": "bold", "fontSize": "10pt", "color": "#1B2E6F"}},
                                                        {"value": "👤\n", "textStyle": {"fontSize": "11pt", "color": "#7895D1"}},
                                                        {"value": "Sign out", "textStyle": {"fontWeight": "normal", "fontSize": "7.5pt", "color": "#00C2FF"}}
                                                    ]
                                                }
                                            ]
                                        }
                                    }],
                                    "background": [{"properties": {"show": {"expr": {"Literal": {"Value": "false"}}}}}]
                                }
                            }
                        })
                    },
                    # ---------------------------------------------------------
                    # 3. CARD SUPERIOR HERO: "Devices" (6 DONUT RINGS EXACTOS)
                    # ---------------------------------------------------------
                    # Contenedor Blanco
                    {
                        "x": 234, "y": 58, "z": 0, "width": 1026, "height": 182,
                        "config": json.dumps({
                            "name": "Card_Devices_Frame",
                            "layouts": [{"id": 0, "position": {"x": 234, "y": 58, "z": 0, "width": 1026, "height": 182}}],
                            "singleVisual": {
                                "visualType": "shape",
                                "objects": {
                                    "shape": [{"properties": {"tileShape": {"expr": {"Literal": {"Value": "'rectangle'"}}}}}],
                                    "fill": [{"properties": {"fillColor": {"solid": {"color": {"expr": {"Literal": {"Value": "'#FFFFFF'"}}}}}}}],
                                    "line": [{"properties": {"lineColor": {"solid": {"color": {"expr": {"Literal": {"Value": "'#DDE4F0'"}}}}}}}],
                                    "roundCorners": [{"properties": {"radius": {"expr": {"Literal": {"Value": "4D"}}}}}]
                                }
                            }
                        })
                    },
                    # Barra Header Deep Navy "Devices"
                    *card_header_bar(name="Header_Devices", x=234, y=58, width=1026, height=24, z=1, title_text="Devices  (Indicadores de Red y Gobernanza)"),
                    
                    # 6 Circular Donut Rings SVG idénticos al mockup
                    svg_ring_card("Ring_1_Champion", x=240, y=86, width=164, height=148, z=3, query_ref="Data_Agregados.SVG Ring Champion"),
                    svg_ring_card("Ring_2_Regional", x=408, y=86, width=164, height=148, z=4, query_ref="Data_Agregados.SVG Ring Regional"),
                    svg_ring_card("Ring_3_Alerta", x=576, y=86, width=164, height=148, z=5, query_ref="Data_Agregados.SVG Ring Alerta"),
                    svg_ring_card("Ring_4_Filial", x=744, y=86, width=164, height=148, z=6, query_ref="Data_Agregados.SVG Ring Filial"),
                    svg_ring_card("Ring_5_Trading7180", x=912, y=86, width=164, height=148, z=7, query_ref="Data_Agregados.SVG Ring Trading7180"),
                    svg_ring_card("Ring_6_Trading7277", x=1080, y=86, width=170, height=148, z=8, query_ref="Data_Agregados.SVG Ring Trading7277"),

                    # ---------------------------------------------------------
                    # 4. FILA MEDIA: CARD "Countries" (IZQUIERDA) + "Countries Graphic" (DERECHA)
                    # ---------------------------------------------------------
                    # 4.1 CARD IZQUIERDA: "Countries" (DONUT GOBERNANZA / RED)
                    {
                        "x": 234, "y": 250, "z": 0, "width": 500, "height": 224,
                        "config": json.dumps({
                            "name": "Card_Countries_Frame",
                            "layouts": [{"id": 0, "position": {"x": 234, "y": 250, "z": 0, "width": 500, "height": 224}}],
                            "singleVisual": {
                                "visualType": "shape",
                                "objects": {
                                    "shape": [{"properties": {"tileShape": {"expr": {"Literal": {"Value": "'rectangle'"}}}}}],
                                    "fill": [{"properties": {"fillColor": {"solid": {"color": {"expr": {"Literal": {"Value": "'#FFFFFF'"}}}}}}}],
                                    "line": [{"properties": {"lineColor": {"solid": {"color": {"expr": {"Literal": {"Value": "'#DDE4F0'"}}}}}}}],
                                    "roundCorners": [{"properties": {"radius": {"expr": {"Literal": {"Value": "4D"}}}}}]
                                }
                            }
                        })
                    },
                    *card_header_bar(name="Header_Countries", x=234, y=250, width=500, height=24, z=1, title_text="Countries  (Distribución de Gobernanza por Nivel)"),
                    
                    {
                        "x": 240, "y": 278, "z": 3, "width": 488, "height": 190,
                        "config": json.dumps({
                            "name": "Donut_Countries_Gobernanza",
                            "layouts": [{"id": 0, "position": {"x": 240, "y": 278, "z": 3, "width": 488, "height": 190}}],
                            "singleVisual": {
                                "visualType": "donutChart",
                                "projections": {
                                    "Category": [{"queryRef": "Data_Agregados.Nivel Autorización / Alerta"}],
                                    "Y": [{"queryRef": "Data_Agregados.Total Rutas"}]
                                },
                                "prototypeQuery": {
                                    "Version": 2,
                                    "From": [{"Name": "d", "Entity": "Data_Agregados", "Type": 0}],
                                    "Select": [
                                        {"Column": {"Expression": {"SourceRef": {"Source": "d"}}, "Property": "Nivel Autorización / Alerta"}, "Name": "Data_Agregados.Nivel Autorización / Alerta"},
                                        {"Measure": {"Expression": {"SourceRef": {"Source": "d"}}, "Property": "Total Rutas"}, "Name": "Data_Agregados.Total Rutas"}
                                    ]
                                },
                                "objects": {
                                    "legend": [{"properties": {"show": {"expr": {"Literal": {"Value": "true"}}}, "position": {"expr": {"Literal": {"Value": "'Right'"}}}}}],
                                    "labels": [{"properties": {"show": {"expr": {"Literal": {"Value": "true"}}}, "labelStyle": {"expr": {"Literal": {"Value": "'Both'"}}}}}]
                                },
                                "vcObjects": {
                                    "border": [{"properties": {"show": {"expr": {"Literal": {"Value": "false"}}}}}],
                                    "background": [{"properties": {"show": {"expr": {"Literal": {"Value": "true"}}}, "color": {"solid": {"color": {"expr": {"Literal": {"Value": "'#FFFFFF'"}}}}}}}],
                                    "dropShadow": [{"properties": {"show": {"expr": {"Literal": {"Value": "false"}}}}}]
                                }
                            }
                        })
                    },

                    # 4.2 CARD DERECHA: "Countries Graphic" (RANKING BARRAS HORIZONTALES)
                    {
                        "x": 744, "y": 250, "z": 0, "width": 516, "height": 224,
                        "config": json.dumps({
                            "name": "Card_CountriesGraphic_Frame",
                            "layouts": [{"id": 0, "position": {"x": 744, "y": 250, "z": 0, "width": 516, "height": 224}}],
                            "singleVisual": {
                                "visualType": "shape",
                                "objects": {
                                    "shape": [{"properties": {"tileShape": {"expr": {"Literal": {"Value": "'rectangle'"}}}}}],
                                    "fill": [{"properties": {"fillColor": {"solid": {"color": {"expr": {"Literal": {"Value": "'#FFFFFF'"}}}}}}}],
                                    "line": [{"properties": {"lineColor": {"solid": {"color": {"expr": {"Literal": {"Value": "'#DDE4F0'"}}}}}}}],
                                    "roundCorners": [{"properties": {"radius": {"expr": {"Literal": {"Value": "4D"}}}}}]
                                }
                            }
                        })
                    },
                    *card_header_bar(name="Header_CountriesGraphic", x=744, y=250, width=516, height=24, z=1, title_text="Countries Graphic  (Ranking Top Materiales MOP %)"),
                    
                    {
                        "x": 750, "y": 278, "z": 3, "width": 504, "height": 190,
                        "config": json.dumps({
                            "name": "Bar_CountriesGraphic_Materiales",
                            "layouts": [{"id": 0, "position": {"x": 750, "y": 278, "z": 3, "width": 504, "height": 190}}],
                            "singleVisual": {
                                "visualType": "barChart",
                                "projections": {
                                    "Category": [{"queryRef": "Data_Agregados.Denominación"}],
                                    "Y": [{"queryRef": "Data_Agregados.MOP Promedio Terceros"}]
                                },
                                "prototypeQuery": {
                                    "Version": 2,
                                    "From": [{"Name": "d", "Entity": "Data_Agregados", "Type": 0}],
                                    "Select": [
                                        {"Column": {"Expression": {"SourceRef": {"Source": "d"}}, "Property": "Denominación"}, "Name": "Data_Agregados.Denominación"},
                                        {"Measure": {"Expression": {"SourceRef": {"Source": "d"}}, "Property": "MOP Promedio Terceros"}, "Name": "Data_Agregados.MOP Promedio Terceros"}
                                    ],
                                    "OrderBy": [{"Direction": 2, "Expression": {"Measure": {"Expression": {"SourceRef": {"Source": "d"}}, "Property": "MOP Promedio Terceros"}}}]
                                },
                                "objects": {
                                    "categoryAxis": [{"properties": {"showAxisTitle": {"expr": {"Literal": {"Value": "false"}}}, "fontSize": {"expr": {"Literal": {"Value": "8D"}}}}}],
                                    "valueAxis": [{"properties": {"showAxisTitle": {"expr": {"Literal": {"Value": "false"}}}}}],
                                    "dataPoint": [{"properties": {"fill": {"solid": {"color": {"expr": {"Literal": {"Value": "'#1B2E6F'"}}}}}}}],
                                    "labels": [{"properties": {"show": {"expr": {"Literal": {"Value": "true"}}}, "labelPosition": {"expr": {"Literal": {"Value": "'OutsideEnd'"}}}, "fontSize": {"expr": {"Literal": {"Value": "8D"}}}}}],
                                    "plotArea": [{"properties": {"transparency": {"expr": {"Literal": {"Value": "0D"}}}}}]
                                },
                                "vcObjects": {
                                    "border": [{"properties": {"show": {"expr": {"Literal": {"Value": "false"}}}}}],
                                    "background": [{"properties": {"show": {"expr": {"Literal": {"Value": "true"}}}, "color": {"solid": {"color": {"expr": {"Literal": {"Value": "'#FFFFFF'"}}}}}}}],
                                    "dropShadow": [{"properties": {"show": {"expr": {"Literal": {"Value": "false"}}}}}]
                                }
                            }
                        })
                    },

                    # ---------------------------------------------------------
                    # 5. FILA INFERIOR: CARD "Calendar" (IZQUIERDA) + "Weekly Access" (DERECHA)
                    # ---------------------------------------------------------
                    # 5.1 CARD IZQUIERDA: "Calendar" (FILTRO TEMPORAL Y SOCIEDADES)
                    {
                        "x": 234, "y": 482, "z": 0, "width": 360, "height": 226,
                        "config": json.dumps({
                            "name": "Card_Calendar_Frame",
                            "layouts": [{"id": 0, "position": {"x": 234, "y": 482, "z": 0, "width": 360, "height": 226}}],
                            "singleVisual": {
                                "visualType": "shape",
                                "objects": {
                                    "shape": [{"properties": {"tileShape": {"expr": {"Literal": {"Value": "'rectangle'"}}}}}],
                                    "fill": [{"properties": {"fillColor": {"solid": {"color": {"expr": {"Literal": {"Value": "'#FFFFFF'"}}}}}}}],
                                    "line": [{"properties": {"lineColor": {"solid": {"color": {"expr": {"Literal": {"Value": "'#DDE4F0'"}}}}}}}],
                                    "roundCorners": [{"properties": {"radius": {"expr": {"Literal": {"Value": "4D"}}}}}]
                                }
                            }
                        })
                    },
                    *card_header_bar(name="Header_Calendar", x=234, y=482, width=360, height=24, z=1, title_text="Calendar  (Segmentación por Sociedad SAP)"),
                    
                    {
                        "x": 240, "y": 510, "z": 3, "width": 348, "height": 192,
                        "config": json.dumps({
                            "name": "Slicer_Calendar_Sociedad",
                            "layouts": [{"id": 0, "position": {"x": 240, "y": 510, "z": 3, "width": 348, "height": 192}}],
                            "singleVisual": {
                                "visualType": "slicer",
                                "projections": {"Values": [{"queryRef": "Data_Agregados.Sociedad"}]},
                                "prototypeQuery": {
                                    "Version": 2,
                                    "From": [{"Name": "d", "Entity": "Data_Agregados", "Type": 0}],
                                    "Select": [{"Column": {"Expression": {"SourceRef": {"Source": "d"}}, "Property": "Sociedad"}, "Name": "Data_Agregados.Sociedad"}]
                                },
                                "objects": {
                                    "header": [{"properties": {"show": {"expr": {"Literal": {"Value": "false"}}}}}],
                                    "selection": [{"properties": {"selectAllCheckboxEnabled": {"expr": {"Literal": {"Value": "true"}}}}}]
                                },
                                "vcObjects": {
                                    "border": [{"properties": {"show": {"expr": {"Literal": {"Value": "false"}}}}}],
                                    "background": [{"properties": {"show": {"expr": {"Literal": {"Value": "true"}}}, "color": {"solid": {"color": {"expr": {"Literal": {"Value": "'#FFFFFF'"}}}}}}}],
                                    "dropShadow": [{"properties": {"show": {"expr": {"Literal": {"Value": "false"}}}}}]
                                }
                            }
                        })
                    },

                    # 5.2 CARD DERECHA: "Weekly Access" (MATRIZ DE AUDITORÍA Y MICROVISUALES)
                    {
                        "x": 604, "y": 482, "z": 0, "width": 656, "height": 226,
                        "config": json.dumps({
                            "name": "Card_WeeklyAccess_Frame",
                            "layouts": [{"id": 0, "position": {"x": 604, "y": 482, "z": 0, "width": 656, "height": 226}}],
                            "singleVisual": {
                                "visualType": "shape",
                                "objects": {
                                    "shape": [{"properties": {"tileShape": {"expr": {"Literal": {"Value": "'rectangle'"}}}}}],
                                    "fill": [{"properties": {"fillColor": {"solid": {"color": {"expr": {"Literal": {"Value": "'#FFFFFF'"}}}}}}}],
                                    "line": [{"properties": {"lineColor": {"solid": {"color": {"expr": {"Literal": {"Value": "'#DDE4F0'"}}}}}}}],
                                    "roundCorners": [{"properties": {"radius": {"expr": {"Literal": {"Value": "4D"}}}}}]
                                }
                            }
                        })
                    },
                    *card_header_bar(name="Header_WeeklyAccess", x=604, y=482, width=656, height=24, z=1, title_text="Weekly Access  (Auditoría de Rutas y Microvisuales SVG)"),
                    
                    {
                        "x": 610, "y": 510, "z": 3, "width": 644, "height": 192,
                        "config": json.dumps({
                            "name": "Table_WeeklyAccess_Auditoria",
                            "layouts": [{"id": 0, "position": {"x": 610, "y": 510, "z": 3, "width": 644, "height": 192}}],
                            "singleVisual": {
                                "visualType": "tableEx",
                                "projections": {
                                    "Values": [
                                        {"queryRef": "Data_Agregados.Centro"},
                                        {"queryRef": "Data_Agregados.Destino"},
                                        {"queryRef": "Data_Agregados.Denominación"},
                                        {"queryRef": "Data_Agregados.Importe MP"},
                                        {"queryRef": "Data_Agregados.MP Compra (Costo Material)"},
                                        {"queryRef": "Data_Agregados.SVG MOP Bullet Bar"},
                                        {"queryRef": "Data_Agregados.SVG Status Pill"}
                                    ]
                                },
                                "prototypeQuery": {
                                    "Version": 2,
                                    "From": [{"Name": "d", "Entity": "Data_Agregados", "Type": 0}],
                                    "Select": [
                                        {"Column": {"Expression": {"SourceRef": {"Source": "d"}}, "Property": "Centro"}, "Name": "Data_Agregados.Centro"},
                                        {"Column": {"Expression": {"SourceRef": {"Source": "d"}}, "Property": "Destino"}, "Name": "Data_Agregados.Destino"},
                                        {"Column": {"Expression": {"SourceRef": {"Source": "d"}}, "Property": "Denominación"}, "Name": "Data_Agregados.Denominación"},
                                        {"Column": {"Expression": {"SourceRef": {"Source": "d"}}, "Property": "Importe MP"}, "Name": "Data_Agregados.Importe MP"},
                                        {"Column": {"Expression": {"SourceRef": {"Source": "d"}}, "Property": "MP Compra (Costo Material)"}, "Name": "Data_Agregados.MP Compra (Costo Material)"},
                                        {"Measure": {"Expression": {"SourceRef": {"Source": "d"}}, "Property": "SVG MOP Bullet Bar"}, "Name": "Data_Agregados.SVG MOP Bullet Bar"},
                                        {"Measure": {"Expression": {"SourceRef": {"Source": "d"}}, "Property": "SVG Status Pill"}, "Name": "Data_Agregados.SVG Status Pill"}
                                    ]
                                },
                                "objects": {
                                    "values": [{
                                        "properties": {
                                            "fontSize": {"expr": {"Literal": {"Value": "8.5D"}}},
                                            "imageHeight": {"expr": {"Literal": {"Value": "18L"}}},
                                            "imageWidth": {"expr": {"Literal": {"Value": "95L"}}}
                                        }
                                    }],
                                    "columnHeaders": [{
                                        "properties": {
                                            "fontColor": {"solid": {"color": {"expr": {"Literal": {"Value": "'#FFFFFF'"}}}}},
                                            "backColor": {"solid": {"color": {"expr": {"Literal": {"Value": "'#1B2E6F'"}}}}},
                                            "fontSize": {"expr": {"Literal": {"Value": "8.5D"}}},
                                            "bold": {"expr": {"Literal": {"Value": "true"}}}
                                        }
                                    }]
                                },
                                "vcObjects": {
                                    "border": [{"properties": {"show": {"expr": {"Literal": {"Value": "false"}}}}}],
                                    "background": [{"properties": {"show": {"expr": {"Literal": {"Value": "true"}}}, "color": {"solid": {"color": {"expr": {"Literal": {"Value": "'#FFFFFF'"}}}}}}}],
                                    "dropShadow": [{"properties": {"show": {"expr": {"Literal": {"Value": "false"}}}}}]
                                }
                            }
                        })
                    }
                ]
            }
        ]
    }
    with open(os.path.join(rep_dir, "report.json"), "w", encoding="utf-8") as f:
        json.dump(report_json, f, indent=2, ensure_ascii=False)

    logging.info(f"¡Proyecto Power BI (.pbip) creado exitosamente en: {pbip_file}!")
    return pbip_file

if __name__ == "__main__":
    script_dir = os.path.dirname(os.path.abspath(__file__))
    out_dir = os.path.join(script_dir, "_salidas_integradas", "PowerBI_Project")
    
    pbi_files = glob.glob(os.path.join(script_dir, "_salidas_integradas", "Base_PowerBI_Agregados_*.xlsx"))
    data_file = pbi_files[-1] if pbi_files else os.path.join(script_dir, "_salidas_integradas", "Base_PowerBI_Agregados.xlsx")
    
    crear_proyecto_pbip(out_dir, data_file)
