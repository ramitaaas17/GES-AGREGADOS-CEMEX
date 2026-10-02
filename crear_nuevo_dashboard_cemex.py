"""
=============================================================================
GENERADOR DEL NUEVO DASHBOARD CEMEX AGREGADOS (.PBIP)
Totalmente nuevo, limpio, sin dependencias viejas ni archivos heredados.
=============================================================================
"""

import os
import json
import logging

logging.basicConfig(level=logging.INFO, format="%(asctime)s - %(levelname)s - %(message)s")

def build_new_cemex_dashboard():
    base_dir = r"C:\Users\coazo\ramitaaas17 GES-AGREGADOS-CEMEX main MatrizVentas\_salidas_integradas"
    output_dir = os.path.join(base_dir, "CEMEX_Dashboard_Nuevo")
    excel_dataset_path = os.path.join(base_dir, "Base_PowerBI_Agregados_20260929_153844.xlsx")
    
    os.makedirs(output_dir, exist_ok=True)
    project_name = "CEMEX_Dashboard_Nuevo"
    
    pbip_file = os.path.join(output_dir, f"{project_name}.pbip")
    sm_dir = os.path.join(output_dir, f"{project_name}.SemanticModel")
    rep_dir = os.path.join(output_dir, f"{project_name}.Report")
    
    shared_themes_dir = os.path.join(rep_dir, "StaticResources", "SharedResources", "BaseThemes")
    reg_themes_dir = os.path.join(rep_dir, "StaticResources", "RegisteredResources")
    
    os.makedirs(os.path.join(sm_dir, ".pbi"), exist_ok=True)
    os.makedirs(os.path.join(rep_dir, ".pbi"), exist_ok=True)
    os.makedirs(shared_themes_dir, exist_ok=True)
    os.makedirs(reg_themes_dir, exist_ok=True)
    
    data_path_escaped = excel_dataset_path.replace("\\", "\\\\")
    
    # -------------------------------------------------------------------------
    # 1. ARCHIVO .pbip
    # -------------------------------------------------------------------------
    pbip_content = {
        "version": "1.0",
        "artifacts": [{"report": {"path": f"{project_name}.Report"}}],
        "settings": {"enableAutoRecovery": True}
    }
    with open(pbip_file, "w", encoding="utf-8") as f:
        json.dump(pbip_content, f, indent=2, ensure_ascii=False)
        
    # -------------------------------------------------------------------------
    # 2. DEFINITIONS (.pbir y .pbism)
    # -------------------------------------------------------------------------
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

    with open(os.path.join(sm_dir, ".pbi", "localSettings.json"), "w", encoding="utf-8") as f:
        json.dump({"version": "1.0"}, f, indent=2)
    with open(os.path.join(rep_dir, ".pbi", "localSettings.json"), "w", encoding="utf-8") as f:
        json.dump({"version": "1.0"}, f, indent=2)

    # -------------------------------------------------------------------------
    # 3. TEMA CEMEX CORPORATIVO
    # -------------------------------------------------------------------------
    base_theme_content = {
        "name": "CY23SU04",
        "version": "5.46",
        "type": 2
    }
    with open(os.path.join(shared_themes_dir, "CY23SU04.json"), "w", encoding="utf-8") as f:
        json.dump(base_theme_content, f, indent=2)

    theme_registered_name = "CEMEX_Corporate_Theme.json"
    cemex_theme = {
        "name": "CEMEX_Corporate_Theme",
        "dataColors": [
            "#002D62",  # CEMEX Navy
            "#00C2FF",  # Cyan Acento
            "#00A859",  # Verde Vertua
            "#D32027",  # Rojo Alerta
            "#3B5CA8",  # Azul Secundario
            "#F59E0B",  # Ámbar Advertencia
            "#64748B",  # Gris Pizarra
            "#1E293B"   # Texto Oscuro
        ],
        "background": "#F4F6F9",
        "foreground": "#1E293B",
        "tableAccent": "#002D62",
        "good": "#00A859",
        "neutral": "#00C2FF",
        "bad": "#D32027",
        "textClasses": {
            "callout": {
                "fontFace": "Segoe UI, Outfit, Arial, sans-serif",
                "fontSize": 20,
                "color": "#002D62"
            },
            "title": {
                "fontFace": "Segoe UI, Outfit, Arial, sans-serif",
                "fontSize": 12,
                "color": "#002D62",
                "fontWeight": "bold"
            },
            "header": {
                "fontFace": "Segoe UI, Outfit, Arial, sans-serif",
                "fontSize": 10,
                "color": "#1E293B",
                "fontWeight": "bold"
            },
            "label": {
                "fontFace": "Segoe UI, Outfit, Arial, sans-serif",
                "fontSize": 9,
                "color": "#475569"
            }
        },
        "visualStyles": {
            "*": {
                "*": {
                    "background": [{"color": {"solid": {"color": "#FFFFFF"}}, "transparency": 0}],
                    "border": [{"show": True, "color": {"solid": {"color": "#E2E8F0"}}, "radius": 6}],
                    "dropShadow": [{"show": True, "color": {"solid": {"color": "#000000"}}, "transparency": 92, "blur": 4}]
                }
            },
            "page": {
                "*": {
                    "background": [{"color": {"solid": {"color": "#F4F6F9"}}, "transparency": 0}]
                }
            }
        }
    }
    with open(os.path.join(reg_themes_dir, theme_registered_name), "w", encoding="utf-8") as f:
        json.dump(cemex_theme, f, indent=2)

    # -------------------------------------------------------------------------
    # 4. MODELO SEMÁNTICO (model.bim)
    # -------------------------------------------------------------------------
    columns_config = [
        ("Concat1", "string", "text"),
        ("Concat2", "string", "text"),
        ("Sociedad", "string", "text"),
        ("Ship From", "string", "text"),
        ("Nombre SF", "string", "text"),
        ("Centro", "string", "text"),
        ("Desc. Centro", "string", "text"),
        ("Destino", "string", "text"),
        ("Nombre Destino", "string", "text"),
        ("Material", "string", "text"),
        ("Denominación", "string", "text"),
        ("PV", "double", "number"),
        ("Inicio Vigencia", "string", "text"),
        ("Fin Vigencia", "string", "text"),
        ("Modalidad Venta", "string", "text"),
        ("Clase Cond. MP", "string", "text"),
        ("Importe MP", "double", "number"),
        ("UM Venta", "string", "text"),
        ("Importe Flete", "double", "number"),
        ("No. Contrato Compra", "string", "text"),
        ("Inicio Vigencia Compra", "string", "text"),
        ("Fin Vigencia Compra", "string", "text"),
        ("Costo Total TRAOPE", "double", "number"),
        ("Flete Compra", "double", "number"),
        ("MP Compra (Costo Material)", "double", "number"),
        ("Margen Material (MOP %)", "double", "number"),
        ("UM Costo", "string", "text"),
        ("Tipo Operación", "string", "text"),
        ("Nivel Autorización / Alerta", "string", "text"),
        ("Validacion 1", "double", "number"),
        ("Validacion 2", "double", "number"),
        ("Semaforo", "string", "text"),
        ("No. Contrato Venta", "string", "text"),
        ("UM Contrato", "string", "text"),
        ("Precio Contrato", "double", "number")
    ]

    bim_columns = []
    type_transforms = []
    for col_name, data_type, m_type in columns_config:
        col_def = {
            "name": col_name,
            "dataType": data_type,
            "sourceColumn": col_name,
            "summarizeBy": "none" if data_type == "string" else "sum"
        }
        if data_type == "double":
            col_def["formatString"] = "$#,##0.00;($#,##0.00);$0.00" if "Importe" in col_name or "Costo" in col_name or "Precio" in col_name else "#,##0.00"
        bim_columns.append(col_def)
        type_transforms.append(f'        {{"{col_name}", type {m_type}}}')

    transforms_str = ",\n".join(type_transforms)

    dax_measures = [
        {
            "name": "Total Rutas",
            "expression": "COUNTROWS('Data_Agregados')",
            "formatString": "#,##0",
            "displayFolder": "KPIs"
        },
        {
            "name": "Rutas Validadas OK",
            "expression": "CALCULATE(COUNTROWS('Data_Agregados'), 'Data_Agregados'[Semaforo] = \"OK\" || ISBLANK('Data_Agregados'[Semaforo]))",
            "formatString": "#,##0",
            "displayFolder": "KPIs"
        },
        {
            "name": "% Rutas Validadas",
            "expression": "DIVIDE([Rutas Validadas OK], [Total Rutas], 0)",
            "formatString": "0.0%",
            "displayFolder": "KPIs"
        },
        {
            "name": "Rutas con Alerta",
            "expression": "CALCULATE(COUNTROWS('Data_Agregados'), 'Data_Agregados'[Semaforo] = \"DIFERENCIA\" || 'Data_Agregados'[Semaforo] = \"DISCREPANCIA_ORG\")",
            "formatString": "#,##0",
            "displayFolder": "KPIs"
        },
        {
            "name": "Precio Promedio MP",
            "expression": "AVERAGE('Data_Agregados'[Importe MP])",
            "formatString": "$#,##0.00;($#,##0.00);$0.00",
            "displayFolder": "Financiero"
        },
        {
            "name": "Costo Promedio TRAOPE",
            "expression": "AVERAGE('Data_Agregados'[Costo Total TRAOPE])",
            "formatString": "$#,##0.00;($#,##0.00);$0.00",
            "displayFolder": "Financiero"
        },
        {
            "name": "MOP Promedio Unitario",
            "expression": "[Precio Promedio MP] - [Costo Promedio TRAOPE]",
            "formatString": "$#,##0.00;($#,##0.00);$0.00",
            "displayFolder": "Financiero"
        }
    ]

    model_bim = {
        "name": "Model",
        "compatibilityLevel": 1550,
        "model": {
            "culture": "es-MX",
            "dataSources": [
                {
                    "type": "structured",
                    "name": "ExcelAgregados",
                    "connectionDetails": {
                        "protocol": "file",
                        "address": {"path": excel_dataset_path}
                    },
                    "credential": {
                        "AuthenticationKind": "Anonymous",
                        "kind": "File",
                        "path": excel_dataset_path
                    }
                }
            ],
            "tables": [
                {
                    "name": "Data_Agregados",
                    "columns": bim_columns,
                    "partitions": [
                        {
                            "name": "Data_Agregados",
                            "mode": "import",
                            "source": {
                                "type": "m",
                                "expression": [
                                    "let",
                                    f'    Source = Excel.Workbook(File.Contents("{data_path_escaped}"), null, true),',
                                    '    Data_Sheet = Source{[Item="Data_Agregados",Kind="Sheet"]}[Data],',
                                    '    #"Promoted Headers" = Table.PromoteHeaders(Data_Sheet, [PromoteAllScalars=true]),',
                                    '    #"Changed Type" = Table.TransformColumnTypes(#"Promoted Headers",{\n' + transforms_str + '\n    })',
                                    "in",
                                    '    #"Changed Type"'
                                ]
                            }
                        }
                    ],
                    "measures": dax_measures
                }
            ]
        }
    }

    with open(os.path.join(sm_dir, "model.bim"), "w", encoding="utf-8") as f:
        json.dump(model_bim, f, indent=2, ensure_ascii=False)

    # -------------------------------------------------------------------------
    # 5. DEFINICIÓN DEL REPORTE (report.json)
    # -------------------------------------------------------------------------
    report_json = {
        "config": json.dumps({
            "version": "5.46",
            "themeCollection": {
                "baseTheme": {"name": "CY23SU04", "reportVersionAtImport": "5.46", "type": 2},
                "customTheme": {"name": "CEMEX_Corporate_Theme", "reportVersionAtImport": "5.46", "type": 1}
            },
            "activeSectionName": "Section_Gobernanza"
        }),
        "layoutOptimization": 0,
        "resourcePackages": [
            {
                "resourcePackage": {
                    "disabled": False,
                    "items": [
                        {
                            "name": "CY23SU04.json",
                            "path": "BaseThemes/CY23SU04.json",
                            "type": 202
                        },
                        {
                            "name": theme_registered_name,
                            "path": f"RegisteredResources/{theme_registered_name}",
                            "type": 201
                        }
                    ],
                    "name": "SharedResources",
                    "type": 2
                }
            }
        ],
        "sections": [
            {
                "displayName": "Panel de Gobernanza CEMEX",
                "name": "Section_Gobernanza",
                "ordinal": 0,
                "width": 1280,
                "height": 720,
                "visualContainers": [
                    # 1. HEADER BANNER
                    {
                        "x": 20, "y": 12, "width": 1240, "height": 58,
                        "config": json.dumps({
                            "name": "Header_Banner",
                            "singleVisual": {
                                "visualType": "textbox",
                                "objects": {
                                    "general": [{
                                        "properties": {
                                            "paragraphs": [
                                                {
                                                    "textRuns": [
                                                        {
                                                            "value": "CEMEX MÉXICO  |  Panel de Control y Gobernanza de Agregados\n",
                                                            "textStyle": {"fontWeight": "bold", "fontSize": "13pt", "color": "#002D62", "fontFamily": "Segoe UI"}
                                                        },
                                                        {
                                                            "value": "Auditoría Integral de Precios VK13 vs Costos de Compra TRAOPE y Contratos ZSDD4501",
                                                            "textStyle": {"fontSize": "9pt", "color": "#64748B", "fontFamily": "Segoe UI"}
                                                        }
                                                    ]
                                                }
                                            ]
                                        }
                                    }],
                                    "background": [{"properties": {"color": {"solid": {"color": "#FFFFFF"}}, "transparency": {"expr": {"Literal": {"Value": "0D"}}}}}]
                                }
                            }
                        })
                    },
                    # 2. SLICER 1: CEDIS (Centro)
                    {
                        "x": 20, "y": 80, "width": 220, "height": 190,
                        "config": json.dumps({
                            "name": "Slicer_Centro",
                            "singleVisual": {
                                "visualType": "slicer",
                                "projections": {
                                    "Values": [{"queryRef": "Data_Agregados.Centro"}]
                                },
                                "prototypeQuery": {
                                    "Version": 2,
                                    "From": [{"Name": "d", "Entity": "Data_Agregados", "Type": 0}],
                                    "Select": [{"Column": {"Expression": {"SourceRef": {"Source": "d"}}, "Property": "Centro"}, "Name": "Data_Agregados.Centro"}]
                                },
                                "objects": {
                                    "header": [{"properties": {"title": {"expr": {"Literal": {"Value": "'Filtrar por CEDIS'"}}}}}],
                                    "background": [{"properties": {"color": {"solid": {"color": "#FFFFFF"}}}}]
                                }
                            }
                        })
                    },
                    # 3. SLICER 2: SEMÁFORO / ESTATUS
                    {
                        "x": 20, "y": 280, "width": 220, "height": 180,
                        "config": json.dumps({
                            "name": "Slicer_Semaforo",
                            "singleVisual": {
                                "visualType": "slicer",
                                "projections": {
                                    "Values": [{"queryRef": "Data_Agregados.Semaforo"}]
                                },
                                "prototypeQuery": {
                                    "Version": 2,
                                    "From": [{"Name": "d", "Entity": "Data_Agregados", "Type": 0}],
                                    "Select": [{"Column": {"Expression": {"SourceRef": {"Source": "d"}}, "Property": "Semaforo"}, "Name": "Data_Agregados.Semaforo"}]
                                },
                                "objects": {
                                    "header": [{"properties": {"title": {"expr": {"Literal": {"Value": "'Estatus Semáforo'"}}}}}],
                                    "background": [{"properties": {"color": {"solid": {"color": "#FFFFFF"}}}}]
                                }
                            }
                        })
                    },
                    # 4. SLICER 3: SOCIEDAD
                    {
                        "x": 20, "y": 470, "width": 220, "height": 230,
                        "config": json.dumps({
                            "name": "Slicer_Sociedad",
                            "singleVisual": {
                                "visualType": "slicer",
                                "projections": {
                                    "Values": [{"queryRef": "Data_Agregados.Sociedad"}]
                                },
                                "prototypeQuery": {
                                    "Version": 2,
                                    "From": [{"Name": "d", "Entity": "Data_Agregados", "Type": 0}],
                                    "Select": [{"Column": {"Expression": {"SourceRef": {"Source": "d"}}, "Property": "Sociedad"}, "Name": "Data_Agregados.Sociedad"}]
                                },
                                "objects": {
                                    "header": [{"properties": {"title": {"expr": {"Literal": {"Value": "'Sociedad / Razón Social'"}}}}}],
                                    "background": [{"properties": {"color": {"solid": {"color": "#FFFFFF"}}}}]
                                }
                            }
                        })
                    },
                    # 5. KPI CARD 1: TOTAL RUTAS
                    {
                        "x": 255, "y": 80, "width": 235, "height": 95,
                        "config": json.dumps({
                            "name": "Card_Total_Rutas",
                            "singleVisual": {
                                "visualType": "card",
                                "projections": {"Fields": [{"queryRef": "Data_Agregados.Total Rutas"}]},
                                "prototypeQuery": {
                                    "Version": 2,
                                    "From": [{"Name": "d", "Entity": "Data_Agregados", "Type": 0}],
                                    "Select": [{"Measure": {"Expression": {"SourceRef": {"Source": "d"}}, "Property": "Total Rutas"}, "Name": "Data_Agregados.Total Rutas"}]
                                },
                                "objects": {
                                    "labels": [{"properties": {"color": {"solid": {"color": "#002D62"}}, "fontSize": {"expr": {"Literal": {"Value": "22D"}}}}}],
                                    "categoryAxis": [{"properties": {"show": {"expr": {"Literal": {"Value": "true"}}}}}]
                                }
                            }
                        })
                    },
                    # 6. KPI CARD 2: % RUTAS VALIDADAS
                    {
                        "x": 505, "y": 80, "width": 235, "height": 95,
                        "config": json.dumps({
                            "name": "Card_Pct_Validadas",
                            "singleVisual": {
                                "visualType": "card",
                                "projections": {"Fields": [{"queryRef": "Data_Agregados.% Rutas Validadas"}]},
                                "prototypeQuery": {
                                    "Version": 2,
                                    "From": [{"Name": "d", "Entity": "Data_Agregados", "Type": 0}],
                                    "Select": [{"Measure": {"Expression": {"SourceRef": {"Source": "d"}}, "Property": "% Rutas Validadas"}, "Name": "Data_Agregados.% Rutas Validadas"}]
                                },
                                "objects": {
                                    "labels": [{"properties": {"color": {"solid": {"color": "#00A859"}}, "fontSize": {"expr": {"Literal": {"Value": "22D"}}}}}],
                                    "categoryAxis": [{"properties": {"show": {"expr": {"Literal": {"Value": "true"}}}}}]
                                }
                            }
                        })
                    },
                    # 7. KPI CARD 3: MOP PROMEDIO
                    {
                        "x": 755, "y": 80, "width": 235, "height": 95,
                        "config": json.dumps({
                            "name": "Card_MOP_Promedio",
                            "singleVisual": {
                                "visualType": "card",
                                "projections": {"Fields": [{"queryRef": "Data_Agregados.MOP Promedio Unitario"}]},
                                "prototypeQuery": {
                                    "Version": 2,
                                    "From": [{"Name": "d", "Entity": "Data_Agregados", "Type": 0}],
                                    "Select": [{"Measure": {"Expression": {"SourceRef": {"Source": "d"}}, "Property": "MOP Promedio Unitario"}, "Name": "Data_Agregados.MOP Promedio Unitario"}]
                                },
                                "objects": {
                                    "labels": [{"properties": {"color": {"solid": {"color": "#002D62"}}, "fontSize": {"expr": {"Literal": {"Value": "22D"}}}}}],
                                    "categoryAxis": [{"properties": {"show": {"expr": {"Literal": {"Value": "true"}}}}}]
                                }
                            }
                        })
                    },
                    # 8. KPI CARD 4: RUTAS CON ALERTA
                    {
                        "x": 1005, "y": 80, "width": 255, "height": 95,
                        "config": json.dumps({
                            "name": "Card_Rutas_Alerta",
                            "singleVisual": {
                                "visualType": "card",
                                "projections": {"Fields": [{"queryRef": "Data_Agregados.Rutas con Alerta"}]},
                                "prototypeQuery": {
                                    "Version": 2,
                                    "From": [{"Name": "d", "Entity": "Data_Agregados", "Type": 0}],
                                    "Select": [{"Measure": {"Expression": {"SourceRef": {"Source": "d"}}, "Property": "Rutas con Alerta"}, "Name": "Data_Agregados.Rutas con Alerta"}]
                                },
                                "objects": {
                                    "labels": [{"properties": {"color": {"solid": {"color": "#D32027"}}, "fontSize": {"expr": {"Literal": {"Value": "22D"}}}}}],
                                    "categoryAxis": [{"properties": {"show": {"expr": {"Literal": {"Value": "true"}}}}}]
                                }
                            }
                        })
                    },
                    # 9. BAR CHART: RUTAS POR CEDIS
                    {
                        "x": 255, "y": 190, "width": 485, "height": 240,
                        "config": json.dumps({
                            "name": "BarChart_CEDIS",
                            "singleVisual": {
                                "visualType": "barChart",
                                "projections": {
                                    "Category": [{"queryRef": "Data_Agregados.Centro"}],
                                    "Y": [{"queryRef": "Data_Agregados.Total Rutas"}]
                                },
                                "prototypeQuery": {
                                    "Version": 2,
                                    "From": [{"Name": "d", "Entity": "Data_Agregados", "Type": 0}],
                                    "Select": [
                                        {"Column": {"Expression": {"SourceRef": {"Source": "d"}}, "Property": "Centro"}, "Name": "Data_Agregados.Centro"},
                                        {"Measure": {"Expression": {"SourceRef": {"Source": "d"}}, "Property": "Total Rutas"}, "Name": "Data_Agregados.Total Rutas"}
                                    ]
                                },
                                "objects": {
                                    "title": [{"properties": {"text": {"expr": {"Literal": {"Value": "'Volumen de Rutas Activas por CEDIS'"}}}, "fontSize": {"expr": {"Literal": {"Value": "11D"}}}, "fontColor": {"solid": {"color": "#002D62"}}}}]
                                }
                            }
                        })
                    },
                    # 10. DONUT CHART: ESTATUS SEMÁFORO
                    {
                        "x": 755, "y": 190, "width": 505, "height": 240,
                        "config": json.dumps({
                            "name": "Donut_Semaforo",
                            "singleVisual": {
                                "visualType": "donutChart",
                                "projections": {
                                    "Category": [{"queryRef": "Data_Agregados.Semaforo"}],
                                    "Y": [{"queryRef": "Data_Agregados.Total Rutas"}]
                                },
                                "prototypeQuery": {
                                    "Version": 2,
                                    "From": [{"Name": "d", "Entity": "Data_Agregados", "Type": 0}],
                                    "Select": [
                                        {"Column": {"Expression": {"SourceRef": {"Source": "d"}}, "Property": "Semaforo"}, "Name": "Data_Agregados.Semaforo"},
                                        {"Measure": {"Expression": {"SourceRef": {"Source": "d"}}, "Property": "Total Rutas"}, "Name": "Data_Agregados.Total Rutas"}
                                    ]
                                },
                                "objects": {
                                    "title": [{"properties": {"text": {"expr": {"Literal": {"Value": "'Distribución de Gobernanza y Semáforos'"}}}, "fontSize": {"expr": {"Literal": {"Value": "11D"}}}, "fontColor": {"solid": {"color": "#002D62"}}}}]
                                }
                            }
                        })
                    },
                    # 11. DETAILED AUDIT TABLE
                    {
                        "x": 255, "y": 445, "width": 1005, "height": 255,
                        "config": json.dumps({
                            "name": "Table_Auditoria_Detallada",
                            "singleVisual": {
                                "visualType": "tableEx",
                                "projections": {
                                    "Values": [
                                        {"queryRef": "Data_Agregados.Centro"},
                                        {"queryRef": "Data_Agregados.Material"},
                                        {"queryRef": "Data_Agregados.Denominación"},
                                        {"queryRef": "Data_Agregados.Destino"},
                                        {"queryRef": "Data_Agregados.Importe MP"},
                                        {"queryRef": "Data_Agregados.Importe Flete"},
                                        {"queryRef": "Data_Agregados.Costo Total TRAOPE"},
                                        {"queryRef": "Data_Agregados.Validacion 1"},
                                        {"queryRef": "Data_Agregados.Validacion 2"},
                                        {"queryRef": "Data_Agregados.Semaforo"}
                                    ]
                                },
                                "prototypeQuery": {
                                    "Version": 2,
                                    "From": [{"Name": "d", "Entity": "Data_Agregados", "Type": 0}],
                                    "Select": [
                                        {"Column": {"Expression": {"SourceRef": {"Source": "d"}}, "Property": "Centro"}, "Name": "Data_Agregados.Centro"},
                                        {"Column": {"Expression": {"SourceRef": {"Source": "d"}}, "Property": "Material"}, "Name": "Data_Agregados.Material"},
                                        {"Column": {"Expression": {"SourceRef": {"Source": "d"}}, "Property": "Denominación"}, "Name": "Data_Agregados.Denominación"},
                                        {"Column": {"Expression": {"SourceRef": {"Source": "d"}}, "Property": "Destino"}, "Name": "Data_Agregados.Destino"},
                                        {"Column": {"Expression": {"SourceRef": {"Source": "d"}}, "Property": "Importe MP"}, "Name": "Data_Agregados.Importe MP"},
                                        {"Column": {"Expression": {"SourceRef": {"Source": "d"}}, "Property": "Importe Flete"}, "Name": "Data_Agregados.Importe Flete"},
                                        {"Column": {"Expression": {"SourceRef": {"Source": "d"}}, "Property": "Costo Total TRAOPE"}, "Name": "Data_Agregados.Costo Total TRAOPE"},
                                        {"Column": {"Expression": {"SourceRef": {"Source": "d"}}, "Property": "Validacion 1"}, "Name": "Data_Agregados.Validacion 1"},
                                        {"Column": {"Expression": {"SourceRef": {"Source": "d"}}, "Property": "Validacion 2"}, "Name": "Data_Agregados.Validacion 2"},
                                        {"Column": {"Expression": {"SourceRef": {"Source": "d"}}, "Property": "Semaforo"}, "Name": "Data_Agregados.Semaforo"}
                                    ]
                                },
                                "objects": {
                                    "title": [{"properties": {"text": {"expr": {"Literal": {"Value": "'Auditoría Renglón por Renglón: Precios VK13 vs Costos TRAOPE y Semáforo'"}}}, "fontSize": {"expr": {"Literal": {"Value": "11D"}}}, "fontColor": {"solid": {"color": "#002D62"}}}}]
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

    logging.info(f"✨ ¡Nuevo Proyecto Power BI (.pbip) creado con éxito en: {pbip_file}!")
    return pbip_file

if __name__ == "__main__":
    build_new_cemex_dashboard()
