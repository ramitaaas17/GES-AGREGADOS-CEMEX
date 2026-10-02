"""
=============================================================================
ACTUALIZADOR DE DATOS PARA EL DASHBOARD DE POWER BI CEMEX
Conecta el dashboard al archivo Master Consolidado:
Matriz_Precios_TODOS_Master_Consolidado_C2024_20261002_151706.xlsx (5,504 filas)
=============================================================================
"""

import os
import re
import pandas as pd
import logging

logging.basicConfig(level=logging.INFO, format="%(asctime)s - %(levelname)s - %(message)s")

def actualizar_datos():
    base_dir = r"C:\Users\coazo\ramitaaas17 GES-AGREGADOS-CEMEX main MatrizVentas\_salidas_integradas"
    master_file = os.path.join(base_dir, "Matriz_Precios_TODOS_Master_Consolidado_C2024_20261002_151706.xlsx")
    
    if not os.path.exists(master_file):
        logging.error(f"No se encontró el archivo: {master_file}")
        return

    logging.info(f"Leyendo datos desde Master: {master_file}...")
    df = pd.read_excel(master_file, sheet_name="Matriz", header=1)
    logging.info(f"Filas leídas: {len(df)}, Columnas: {len(df.columns)}")

    # Normalizar nombres de columnas para compatibilidad total con DAX y Power Query
    col_map = {
        'Validación 1': 'Validacion 1',
        'Validación 2': 'Validacion 2',
        'Semáforo': 'Semaforo',
    }
    for old_c in list(df.columns):
        for k, v in col_map.items():
            if k.lower() in str(old_c).lower():
                df.rename(columns={old_c: v}, inplace=True)

    # 1. Actualizar el archivo al que ya apuntaba Power BI (para que con botón 'Actualizar' funcione de inmediato)
    old_pbi_file = os.path.join(base_dir, "Base_PowerBI_Agregados_20260929_153844.xlsx")
    with pd.ExcelWriter(old_pbi_file, engine="openpyxl") as writer:
        df.to_excel(writer, sheet_name="Data_Agregados", index=False)
    logging.info(f"Archivo base actualizado con {len(df)} filas: {old_pbi_file}")

    # 2. Actualizar también en CEMEX_Dashboard_Nuevo
    nuevo_pbi_file = os.path.join(base_dir, "CEMEX_Dashboard_Nuevo", "Base_Datos_CEMEX.xlsx")
    if os.path.exists(os.path.dirname(nuevo_pbi_file)):
        with pd.ExcelWriter(nuevo_pbi_file, engine="openpyxl") as writer:
            df.to_excel(writer, sheet_name="Data_Agregados", index=False)
        logging.info(f"Archivo base CEMEX_Dashboard_Nuevo actualizado: {nuevo_pbi_file}")

    # 3. Actualizar la definición TMDL del SemanticModel de PowerBI_Project
    tmdl_path = os.path.join(base_dir, "PowerBI_Project", "Dashboard_CEMEX_Agregados.SemanticModel", "definition", "tables", "Data_Agregados.tmdl")
    if os.path.exists(tmdl_path):
        with open(tmdl_path, "r", encoding="utf-8") as f:
            tmdl_content = f.read()

        # Reemplazar la partición para que apunte directamente al Master Consolidado
        nueva_particion = """\tpartition Data_Agregados = m
\t\tmode: import
\t\tsource =
\t\t\t\tlet
\t\t\t\t    RutaArchivo = "C:\\\\Users\\\\coazo\\\\ramitaaas17 GES-AGREGADOS-CEMEX main MatrizVentas\\\\_salidas_integradas\\\\Matriz_Precios_TODOS_Master_Consolidado_C2024_20261002_151706.xlsx",
\t\t\t\t    Origen = Excel.Workbook(File.Contents(RutaArchivo), null, true),
\t\t\t\t    Matriz_Sheet = Origen{[Item="Matriz",Kind="Sheet"]}[Data],
\t\t\t\t    FilasOmitidas = Table.Skip(Matriz_Sheet, 1),
\t\t\t\t    Encabezados = Table.PromoteHeaders(FilasOmitidas, [PromoteAllScalars=true]),
\t\t\t\t    RenombrarCols = Table.RenameColumns(Encabezados, {
\t\t\t\t        {"Validación 1", "Validacion 1"},
\t\t\t\t        {"Validación 2", "Validacion 2"},
\t\t\t\t        {"Semáforo", "Semaforo"}
\t\t\t\t    }, MissingField.Ignore),
\t\t\t\t    TiposCambiados = Table.TransformColumnTypes(RenombrarCols,{
\t\t\t\t        {"Concat1", type text}, {"Concat2", type text}, {"Sociedad", type text}, {"Ship From", type text}, {"Nombre SF", type text},
\t\t\t\t        {"Centro", type text}, {"Desc. Centro", type text}, {"Destino", type text}, {"Nombre Destino", type text}, {"Material", type text}, {"Denominación", type text},
\t\t\t\t        {"PV", type number}, {"Inicio Vigencia", type text}, {"Fin Vigencia", type text}, {"Modalidad Venta", type text}, {"Clase Cond. MP", type text},
\t\t\t\t        {"Importe MP", Currency.Type}, {"UM Venta", type text}, {"Importe Flete", Currency.Type},
\t\t\t\t        {"No. Contrato Compra", type text}, {"Inicio Vigencia Compra", type text}, {"Fin Vigencia Compra", type text},
\t\t\t\t        {"Costo Total TRAOPE", Currency.Type}, {"Flete Compra", Currency.Type}, {"MP Compra (Costo Material)", Currency.Type},
\t\t\t\t        {"Margen Material (MOP %)", Percentage.Type}, {"UM Costo", type text}, {"Tipo Operación", type text}, {"Nivel Autorización / Alerta", type text},
\t\t\t\t        {"Validacion 1", Currency.Type}, {"Validacion 2", Currency.Type}, {"Semaforo", type text},
\t\t\t\t        {"No. Contrato Venta", type text}, {"UM Contrato", type text}, {"Precio Contrato", Currency.Type}
\t\t\t\t    }, "es-MX")
\t\t\t\tin
\t\t\t\t    TiposCambiados
"""
        # Reemplazar la sección de partition
        pattern = re.compile(r"\tpartition Data_Agregados = m.*?(?=\n\n\n|\Z)", re.DOTALL)
        if pattern.search(tmdl_content):
            tmdl_content = pattern.sub(nueva_particion, tmdl_content)
            with open(tmdl_path, "w", encoding="utf-8") as f:
                f.write(tmdl_content)
            logging.info(f"TMDL actualizado para apuntar a: {master_file}")
        else:
            logging.warning("No se pudo ubicar el patrón de la partición en Data_Agregados.tmdl")

    logging.info("¡Actualización completada exitosamente!")

if __name__ == "__main__":
    actualizar_datos()
