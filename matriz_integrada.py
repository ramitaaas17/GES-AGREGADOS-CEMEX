import os
import sys
import glob
import time
import pickle
import argparse
import logging
from datetime import datetime
from pathlib import Path
import pandas as pd
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils.dataframe import dataframe_to_rows

# Configuración de Logging
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

# =============================================================================
# COLORES CEMEX Y ESTILOS OPENPYXL
# =============================================================================
CLR_NAVY     = '002D72'  # Azul CEMEX
CLR_VENTA    = 'BDD7EE'  # Azul claro - columnas de venta
CLR_COSTO    = 'FCE4D6'  # Salmón - columnas de costo
CLR_VAL      = 'E2EFDA'  # Verde claro - validaciones
CLR_CONTRATO = 'FFF2CC'  # Amarillo claro - contratos
CLR_WARN     = 'FFFF00'  # Amarillo - diferencia > 1
CLR_RED      = 'FF0000'  # Rojo - error grave
CLR_ORANGE   = 'FFC000'  # Naranja - dato faltante
CLR_HEADER   = 'D9D9D9'  # Gris - encabezados normales
CLR_WHITE    = 'FFFFFF'

FILL_NAVY     = PatternFill(start_color=CLR_NAVY, end_color=CLR_NAVY, fill_type='solid')
FILL_VENTA    = PatternFill(start_color=CLR_VENTA, end_color=CLR_VENTA, fill_type='solid')
FILL_COSTO    = PatternFill(start_color=CLR_COSTO, end_color=CLR_COSTO, fill_type='solid')
FILL_VAL      = PatternFill(start_color=CLR_VAL, end_color=CLR_VAL, fill_type='solid')
FILL_CONTRATO = PatternFill(start_color=CLR_CONTRATO, end_color=CLR_CONTRATO, fill_type='solid')
FILL_WARN     = PatternFill(start_color=CLR_WARN, end_color=CLR_WARN, fill_type='solid')
FILL_RED      = PatternFill(start_color=CLR_RED, end_color=CLR_RED, fill_type='solid')
FILL_ORANGE   = PatternFill(start_color=CLR_ORANGE, end_color=CLR_ORANGE, fill_type='solid')
FILL_HEADER   = PatternFill(start_color=CLR_HEADER, end_color=CLR_HEADER, fill_type='solid')
FILL_WHITE    = PatternFill(start_color=CLR_WHITE, end_color=CLR_WHITE, fill_type='solid')

FONT_WHITE_BOLD = Font(color=CLR_WHITE, bold=True, name='Calibri', size=10)
FONT_BLACK_BOLD = Font(color='000000', bold=True, name='Calibri', size=10)
FONT_NORMAL = Font(color='000000', name='Calibri', size=9)

ALIGN_CENTER = Alignment(horizontal='center', vertical='center')
ALIGN_LEFT = Alignment(horizontal='left', vertical='center')

THIN_BORDER = Border(
    left=Side(style='thin'), 
    right=Side(style='thin'), 
    top=Side(style='thin'), 
    bottom=Side(style='thin')
)

# =============================================================================
# FUNCIONES DE CÁLCULO
# =============================================================================

def calc_validacion1(um_venta, um_costo, cond_exp, importe_mp, importe_flete, pv):
    """
    Reglas de negocio CEMEX:
    - Si UM_venta=TN y UM_costo=TN: resultado = MP + Flete
    - Si Cond.Expedicion=1: resultado = PV * MP
    - Else: resultado = (MP + Flete) * PV
    - Si falta algún dato requerido por la rama: devolver None (NO inventar 0 ni 1)
    """
    mp = importe_mp if pd.notna(importe_mp) else None
    flete = importe_flete if pd.notna(importe_flete) else None
    
    if mp is None:
        return None
        
    um_v = str(um_venta).strip().upper() if pd.notna(um_venta) else ""
    um_c = str(um_costo).strip().upper() if pd.notna(um_costo) else ""
    
    if um_v == 'TN' and um_c == 'TN':
        return mp + (flete if flete is not None else 0)
    
    # Ramas que necesitan PV
    if pd.isna(pv):  # NO inventar PV=0 ni PV=1
        return None
        
    try:
        cond = str(int(cond_exp)).strip() if pd.notna(cond_exp) else ''
    except:
        cond = str(cond_exp).strip() if pd.notna(cond_exp) else ''
    
    if cond == '1':
        return pv * mp
    else:
        return (mp + (flete if flete is not None else 0)) * pv

def calc_validacion2(importe_costo, validacion1):
    """NUNCA devolver número inventado: si falta costo real, devolver None."""
    if pd.isna(importe_costo) or pd.isna(validacion1):
        return None
    return importe_costo - validacion1

def eval_semaforo(row):
    """
    Una fila tiene anomalía si:
    - Importe_MP es None o <= 0
    - Importe_MP > 0 pero Importe_Flete is None (hueco real, marcar como 'SIN_FLETE')
    - Importe_Costo is None (hueco real, marcar como 'SIN_COSTO')
    - PV is None y la rama de Val1 lo necesita (marcar como 'SIN_PV')
    - Validacion2 is not None and abs(Validacion2) > 1 (marcar como 'DIFERENCIA')
    - Clase condición no corresponde a org de ventas
    """
    # 1. MP
    mp = row.get('Importe_MP')
    if pd.isna(mp) or mp <= 0:
        return 'SIN_MP' # Error no listado pero deducido
    
    # 2. Flete
    if pd.isna(row.get('Importe_Flete')):
        return 'SIN_FLETE'
        
    # 3. Costo
    if pd.isna(row.get('Importe_Costo')):
        return 'SIN_COSTO'
        
    # 4. PV (Solo si validacion 1 es nan, puede ser porque falta PV y no es todo en TN)
    if pd.isna(row.get('Validacion 1')):
        um_v = str(row.get('UM Venta', '')).strip().upper()
        um_c = str(row.get('UM Costo', '')).strip().upper()
        if not (um_v == 'TN' and um_c == 'TN'):
            if pd.isna(row.get('PV')):
                return 'SIN_PV'
                
    # 5. Diferencia
    v2 = row.get('Validacion 2')
    if pd.notna(v2) and abs(v2) > 1:
        return 'DIFERENCIA'
        
    # 6. Clase Cond / Org Ventas
    org = str(row.get('Sociedad', '')).strip()
    clase_mp = str(row.get('Clase Cond. MP', '')).strip()
    # 7180 con conds de 7100 es VÁLIDO.
    # Anomalia si org 7100 usa condiciones de 7180
    if org == '7100' and (clase_mp in ['ZMA6', 'ZMP1']):
        return 'DISCREPANCIA_ORG'
        
    return ''

# =============================================================================
# MANEJO DE ARCHIVOS Y PARSEO
# =============================================================================

def str_val(v):
    if v is None or (isinstance(v, float) and pd.isna(v)):
        return ''
    try:
        f = float(v)
        if f == int(f):
            return str(int(f))
        return str(v)
    except (ValueError, TypeError):
        return str(v).strip()

def cache_data(cache_file, data):
    with open(cache_file, 'wb') as f:
        pickle.dump(data, f)

def load_cached_data(cache_file, source_file=None):
    if not os.path.exists(cache_file):
        return None
    if source_file and os.path.exists(source_file):
        if os.path.getmtime(source_file) > os.path.getmtime(cache_file):
            return None # Invalidar
    with open(cache_file, 'rb') as f:
        return pickle.load(f)

def cargar_excel(archivo, force_refresh=False):
    logging.info(f"Cargando archivo Excel: {archivo}")
    cache_file = f"{archivo}.mci_cache_.pkl"
    
    if not force_refresh:
        data = load_cached_data(cache_file, archivo)
        if data is not None:
            logging.info("Datos cargados desde caché.")
            return data
            
    try:
        xls = pd.ExcelFile(archivo)
    except Exception as e:
        logging.error(f"Error al abrir el archivo Excel: {e}")
        sys.exit(1)
        
    hojas_requeridas = ['PVTA_MAT VK13', 'PVTA_FTE VK13', 'CONT_COMPRA TRAOPE', 'PESO_VOL', 'CONT_VTA ZSDD4501']
    hojas_disponibles = xls.sheet_names
    
    faltantes = [h for h in hojas_requeridas if h not in hojas_disponibles]
    if faltantes:
        logging.error(f"Faltan hojas en el archivo Excel: {faltantes}")
        logging.error(f"Hojas disponibles: {hojas_disponibles}")
        sys.exit(1)
        
    data = {}
    for hoja in hojas_requeridas:
        logging.info(f"Parseando hoja {hoja}...")
        data[hoja] = pd.read_excel(xls, sheet_name=hoja, dtype=str) # Cargar como string primero
        
    cache_data(cache_file, data)
    return data

def normalizar_etiq(texto):
    """Normaliza etiqueta: sin tildes, minúsculas, sin espacios extra."""
    import unicodedata
    sin_tildes = "".join(
        c for c in unicodedata.normalize("NFD", texto)
        if unicodedata.category(c) != "Mn"
    )
    return sin_tildes.strip().lower()

def fecha_ddmmyyyy_a_num(fecha_str):
    """Convierte 'DD.MM.YYYY' o '31.12.9999' a entero YYYYMMDD para comparar."""
    try:
        partes = fecha_str.strip().split('.')
        if len(partes) == 3:
            dd, mm, yyyy = partes
            return int(f"{yyyy}{mm.zfill(2)}{dd.zfill(2)}")
    except Exception:
        pass
    return 0

def parse_txt(archivo_txt):
    """Parsea el formato bloque multi-línea separado por tabs (Modo B — SAP VK13 crudo)."""
    logging.info(f"Parseando txt: {archivo_txt}")
    registros = []

    # Probar codificaciones: SAP exporta cp1252, los mocks suelen ser utf-8
    for encoding in ('utf-8-sig', 'utf-8', 'cp1252', 'latin-1'):
        try:
            with open(archivo_txt, 'r', encoding=encoding) as f:
                lines = f.readlines()
            break
        except UnicodeDecodeError:
            continue
    else:
        logging.error(f"No se pudo decodificar el archivo: {archivo_txt}")
        return pd.DataFrame()

    current_reg = {}
    for line in lines:
        line = line.rstrip('\n').rstrip('\r')
        if not line.strip():
            continue

        campos = line.split('\t')
        etiq = normalizar_etiq(campos[0])

        if etiq.startswith("clase de condicion"):
            if current_reg.get('Clase Cond.'):
                registros.append(current_reg.copy())
                current_reg = {}
            vals = [c.strip() for c in campos[1:] if c.strip()]
            current_reg['Clase Cond.'] = vals[0] if vals else ''

        elif etiq.startswith("organizacion ventas") or etiq.startswith("org. ventas"):
            vals = [c.strip() for c in campos[1:] if c.strip()]
            current_reg['Org. Ventas'] = vals[0] if vals else ''

        elif etiq == "acreedor":
            vals = [c.strip() for c in campos[1:] if c.strip()]
            current_reg['Shipfrom'] = vals[0] if vals else ''

        elif etiq == "centro":
            vals = [c.strip() for c in campos[1:] if c.strip()]
            current_reg['Centro'] = vals[0] if vals else ''
            if len(vals) > 1:
                current_reg['Desc. Centro'] = vals[-1]

        elif etiq.startswith("destinatario mcia") or etiq.startswith("destinatario"):
            vals = [c.strip() for c in campos[1:] if c.strip()]
            current_reg['Destinatario'] = vals[0] if vals else ''

        elif etiq == "material":
            vals = [c.strip() for c in campos[1:] if c.strip()]
            current_reg['Material'] = vals[0] if vals else ''
            if len(vals) > 1:
                current_reg['Denominacion'] = vals[-1]

        elif etiq.startswith("periodo de validez"):
            # Buscar fechas DD.MM.YYYY en los campos
            fechas = [c.strip() for c in campos if len(c.strip()) == 10
                      and c.strip()[2] == '.' and c.strip()[5] == '.']
            if fechas:
                current_reg['Inicio Validez'] = fechas[0]
                current_reg['Valido a'] = fechas[-1]
                current_reg['Valido a_num'] = fecha_ddmmyyyy_a_num(fechas[-1])

        elif "MXN" in campos:
            mxn_idx = campos.index("MXN")
            # Importe: buscar hacia atrás desde MXN
            importe = None
            for c in reversed(campos[:mxn_idx]):
                val = c.strip().replace(',', '').replace(' ', '')
                try:
                    importe = float(val)
                    break
                except (ValueError, TypeError):
                    pass
            # UM: último campo no vacío después de MXN
            um = None
            for c in reversed(campos[mxn_idx + 1:]):
                if c.strip() and c.strip().upper() not in ('MXN', 'POR'):
                    um = c.strip()
                    break

            current_reg['Importe'] = importe
            current_reg['Unidad'] = um if um else ''
            if current_reg.get('Clase Cond.'):
                registros.append(current_reg.copy())
            current_reg = {}

    if current_reg.get('Clase Cond.'):
        registros.append(current_reg)

    df = pd.DataFrame(registros)

    # Aplicar modalidad según org/clase
    if not df.empty:
        df['Modalidad'] = df.apply(lambda r: _asignar_modalidad(
            r.get('Org. Ventas', ''), r.get('Clase Cond.', '')), axis=1)
        # Filtro de vigencia: solo 31.12.9999
        if 'Valido a_num' in df.columns:
            df = df[df['Valido a_num'] >= 99991231]
        logging.info(f"  → {len(df)} registros vigentes tras filtro 9999")

    return df

def _asignar_modalidad(org, clase):
    """Devuelve 'MATERIAL' o 'FLETE' según org de ventas y clase de condición."""
    clase_u = str(clase).upper()
    if org == '7100':
        if clase_u == 'ZMAH': return 'MATERIAL'
        if clase_u == 'ZMPH': return 'FLETE'
    elif org == '7180':
        if clase_u == 'ZMA6': return 'MATERIAL'
        if clase_u == 'ZMP1': return 'FLETE'
    return clase_u  # Devuelve la clase como texto si no coincide



# =============================================================================
# PROCESAMIENTO
# =============================================================================

def procesar_datos(data, cedis=None, modo_a=True):
    logging.info("Cruzando datos y calculando...")
    
    if modo_a:
        df_mp = data['PVTA_MAT VK13'].copy()
        df_flete = data['PVTA_FTE VK13'].copy()
        df_traope = data['CONT_COMPRA TRAOPE'].copy()
        df_pv = data['PESO_VOL'].copy()
        df_contratos = data['CONT_VTA ZSDD4501'].copy()
        
        # Filtros obligatorios
        # Fechas as YYYYMMDD numeric
        df_mp['Valido a_num'] = pd.to_numeric(df_mp['Valido a'], errors='coerce')
        df_mp = df_mp[df_mp['Valido a_num'] >= 99991231]
        
        df_flete['Fin Validez_num'] = pd.to_numeric(df_flete['Fin Validez'], errors='coerce')
        df_flete = df_flete[df_flete['Fin Validez_num'] >= 99991231]
        
        if 'Validez a' in df_traope.columns:
            df_traope['Validez a_num'] = pd.to_numeric(df_traope['Validez a'].str.replace('.', '').str.replace('-', ''), errors='coerce')
            # 20291231
            df_traope = df_traope[df_traope['Validez a_num'] >= 20291231]
            
        if 'Fecha Fin validez' in df_contratos.columns:
            df_contratos['Fecha Fin validez_num'] = pd.to_numeric(df_contratos['Fecha Fin validez'], errors='coerce')
            df_contratos = df_contratos[df_contratos['Fecha Fin validez_num'] >= 20291231]
            
        # Limpieza de llaves
        for df in [df_mp, df_flete]:
            df['Shipfrom'] = df['Shipfrom'].apply(str_val)
            df['Centro'] = df['Centro'].apply(str_val)
            df['Destinatario'] = df['Destinatario'].apply(str_val)
            df['Material'] = df['Material'].apply(str_val)
            df['Concat1'] = df['Shipfrom'] + '-' + df['Centro'] + '-' + df['Destinatario'] + '-' + df['Material']
            df['Concat2'] = df['Shipfrom'] + '-' + df['Centro'] + '-' + df['Material']
            
        # Filtro de CEDIS
        if cedis:
            df_mp = df_mp[df_mp['Centro'] == cedis]
            
        if df_mp.empty:
            logging.warning("El archivo está vacío o sin datos vigentes para MP.")
            
        # Importes numéricos
        df_mp['Importe'] = pd.to_numeric(df_mp['Importe'], errors='coerce')
        df_flete['Importe'] = pd.to_numeric(df_flete['Importe'], errors='coerce')
        
        # PV: convertir kg/m3 a TN
        df_pv['Cant. UMB'] = pd.to_numeric(df_pv['Cant. UMB'], errors='coerce')
        df_pv['PV_calc'] = df_pv['Cant. UMB'] / 1000.0
        df_pv['Material'] = df_pv['Material'].apply(str_val)
        
        # TRAOPE llaves
        df_traope['Ship From'] = df_traope['Ship From'].apply(str_val)
        df_traope['Centro'] = df_traope['Centro'].apply(str_val)
        
        # Ajuste de columna Destino si existe
        col_destino = 'Destino' if 'Destino' in df_traope.columns else None
        if col_destino:
            df_traope[col_destino] = df_traope[col_destino].apply(str_val)
        
        df_traope['Material'] = df_traope['Material'].apply(str_val)
        
        if col_destino:
            df_traope['Concat1'] = df_traope['Ship From'] + '-' + df_traope['Centro'] + '-' + df_traope[col_destino] + '-' + df_traope['Material']
        else:
            df_traope['Concat1'] = 'NA'
            
        df_traope['Concat2'] = df_traope['Ship From'] + '-' + df_traope['Centro'] + '-' + df_traope['Material']
        
        # Hay casos donde Destino en traope es nulo. Agrupamos por Concat1 y Concat2 para quedarnos con uno
        traope_c1 = df_traope[df_traope['Concat1'] != 'NA'].drop_duplicates('Concat1').set_index('Concat1')
        traope_c2 = df_traope.drop_duplicates('Concat2').set_index('Concat2')
        
        # Contratos venta llaves
        if 'Ruta A' in df_contratos.columns:
            df_contratos['Llave_Contrato'] = df_contratos['Ruta A'].apply(str_val)
        else:
            df_contratos['Llave_Contrato'] = 'NA'
            
        contratos_idx = df_contratos.drop_duplicates('Llave_Contrato').set_index('Llave_Contrato')
        
        # Fletes: SOLO cruce exacto por Concat1 (el flete depende estrictamente del destino)
        flete_c1 = df_flete.drop_duplicates('Concat1').set_index('Concat1')
        
        # PV
        pv_idx = df_pv.drop_duplicates('Material').set_index('Material')
        
        # =====================================================================
        # CONSTRUIR LA MATRIZ
        # =====================================================================
        
        filas = []
        for _, row in df_mp.iterrows():
            c1 = row['Concat1']
            c2 = row['Concat2']
            mat = row['Material']
            
            # 1. Flete: ÚNICAMENTE por Concat1 exacto
            f_row = None
            if c1 in flete_c1.index:
                f_row = flete_c1.loc[c1]
                
            # 2. PV
            pv_val = None
            pv_desc = ''
            if mat in pv_idx.index:
                pv_val = pv_idx.loc[mat, 'PV_calc']
                pv_desc = pv_idx.loc[mat, 'Texto de material']
                
            # 3. TRAOPE
            t_row = None
            if c1 in traope_c1.index:
                t_row = traope_c1.loc[c1]
            elif c2 in traope_c2.index:
                t_row = traope_c2.loc[c2]
                
            # 4. Contratos
            co_row = None
            # La llave de contrato es Centro-Shipfrom-Destino-Material = Centro + '-' + Shipfrom + '-' + Destino + '-' + Material
            llave_co = f"{row['Centro']}-{row['Shipfrom']}-{row['Destinatario']}-{mat}"
            if llave_co in contratos_idx.index:
                co_row = contratos_idx.loc[llave_co]
                
            # Importes
            imp_mp = row['Importe']
            imp_flete = f_row['Importe'] if f_row is not None else None
            
            # Buscar UM y precio en traope
            imp_costo = None
            um_costo = None
            if t_row is not None:
                # Precio neto pedido
                if 'Precio neto pedido' in t_row.index:
                    imp_costo = pd.to_numeric(t_row['Precio neto pedido'], errors='coerce')
                # UM Precio Pedido - asumiendo columna UM
                if 'UM Precio Pedido' in t_row.index:
                    um_costo = str(t_row['UM Precio Pedido'])
                elif 'UM' in t_row.index:
                    um_costo = str(t_row['UM'])
                elif 'UMB' in t_row.index:
                    um_costo = str(t_row['UMB'])
                
            um_venta = row.get('Unidad', '')
            cond_exp = f_row['Cond. Expedición'] if f_row is not None and 'Cond. Expedición' in f_row.index else ''
            
            # Validaciones
            val1 = calc_validacion1(um_venta, um_costo, cond_exp, imp_mp, imp_flete, pv_val)
            val2 = calc_validacion2(imp_costo, val1)
            
            fila_dict = {
                'Concat1': c1,
                'Concat2': c2,
                'Sociedad': row.get('Org. Ventas', ''),
                'Clase Cond. MP': row.get('Clase Cond.', ''),
                'Clase Cond. Flete': f_row['Clase Cond.'] if f_row is not None else '',
                'Ship From': row.get('Shipfrom', ''),
                'Nombre SF': t_row['Nombre SF'] if t_row is not None and 'Nombre SF' in t_row.index else '',
                'Centro': row.get('Centro', ''),
                'Desc. Centro': t_row['Descripción Centro'] if t_row is not None and 'Descripción Centro' in t_row.index else '',
                'Destino': row.get('Destinatario', ''),
                'Cond. Expedición': cond_exp,
                'Material': mat,
                'Denominación': pv_desc,
                'PV': pv_val,
                'Inicio Vigencia': row.get('Inicio Validez', ''),
                'Fin Vigencia': row.get('Valido a', ''),
                'Importe_MP': imp_mp,
                'Importe_Flete': imp_flete,
                'UM Venta': um_venta,
                'Importe_Costo': imp_costo,
                'UM Costo': um_costo,
                'Validacion 1': val1,
                'Validacion 2': val2,
                'No. Contrato Venta': co_row['Llave'] if co_row is not None and 'Llave' in co_row.index else (co_row['Tipo Documento'] if co_row is not None and 'Tipo Documento' in co_row.index else ''),
                'UM Contrato': co_row['UM'] if co_row is not None and 'UM' in co_row.index else '',
                'Precio Contrato': pd.to_numeric(co_row['Precio Neto'], errors='coerce') if co_row is not None and 'Precio Neto' in co_row.index else None
            }
            
            # Semáforo
            fila_dict['Semaforo'] = eval_semaforo(fila_dict)
            
            filas.append(fila_dict)
            
        matriz_df = pd.DataFrame(filas)
        return matriz_df, df_mp, df_flete, df_traope, df_contratos
        
    else:
        # Modo B
        logging.warning("Modo B seleccionado. Funcionalidad reducida (sin cruce completo).")
        df_mp = data['MP']
        df_flete = data['Flete']
        # Mismo procedimiento simplificado
        return pd.DataFrame(), None, None, None, None

# =============================================================================
# ESCRITURA EN EXCEL
# =============================================================================

def apply_header_style(cell, text, fill, font):
    cell.value = text
    cell.fill = fill
    cell.font = font
    cell.alignment = ALIGN_CENTER
    cell.border = THIN_BORDER

def escribir_excel(df_matriz, cedis, out_dir, raw_data_frames):
    if df_matriz.empty:
        logging.error("DataFrame vacío, no se generará Excel.")
        return
        
    fecha = datetime.now().strftime('%Y%m%d_%H%M%S')
    cedis_str = cedis if cedis else "TODOS"
    os.makedirs(out_dir, exist_ok=True)
    out_file = os.path.join(out_dir, f"Matriz_Precios_Integral_{cedis_str}_{fecha}.xlsx")
    
    logging.info(f"Escribiendo Excel: {out_file}")
    
    wb = openpyxl.Workbook()
    
    # ==========================
    # Hoja 1: Matriz
    # ==========================
    ws_matriz = wb.active
    ws_matriz.title = "Matriz"
    
    # Definición de encabezados y grupos
    headers = [
        ('Concat1', FILL_HEADER, FONT_BLACK_BOLD, 1),
        ('Concat2', FILL_HEADER, FONT_BLACK_BOLD, 1),
        ('Sociedad / Org Ventas', FILL_HEADER, FONT_BLACK_BOLD, 1),
        ('Clase Cond. MP', FILL_HEADER, FONT_BLACK_BOLD, 1),
        ('Clase Cond. Flete', FILL_HEADER, FONT_BLACK_BOLD, 1),
        ('Ship From', FILL_HEADER, FONT_BLACK_BOLD, 1),
        ('Nombre SF', FILL_HEADER, FONT_BLACK_BOLD, 1),
        ('Centro', FILL_HEADER, FONT_BLACK_BOLD, 1),
        ('Desc. Centro', FILL_HEADER, FONT_BLACK_BOLD, 1),
        ('Destino', FILL_HEADER, FONT_BLACK_BOLD, 1),
        ('Cond. Expedición', FILL_HEADER, FONT_BLACK_BOLD, 1),
        ('Material', FILL_HEADER, FONT_BLACK_BOLD, 1),
        ('Denominación', FILL_HEADER, FONT_BLACK_BOLD, 1),
        ('PV', FILL_HEADER, FONT_BLACK_BOLD, 1),
        ('Inicio Vigencia', FILL_HEADER, FONT_BLACK_BOLD, 1),
        ('Fin Vigencia', FILL_HEADER, FONT_BLACK_BOLD, 1),
        
        # === VENTA ===
        ('Importe MP', FILL_VENTA, FONT_BLACK_BOLD, 1),
        ('Importe Flete', FILL_VENTA, FONT_BLACK_BOLD, 1),
        ('UM Venta', FILL_VENTA, FONT_BLACK_BOLD, 1),
        
        # === COSTO ===
        ('Importe Costo', FILL_COSTO, FONT_BLACK_BOLD, 1),
        ('UM Costo', FILL_COSTO, FONT_BLACK_BOLD, 1),
        
        # === VALIDACIONES ===
        ('Validación 1', FILL_VAL, FONT_BLACK_BOLD, 1),
        ('Validación 2', FILL_VAL, FONT_BLACK_BOLD, 1),
        ('Semáforo', FILL_VAL, FONT_BLACK_BOLD, 1),
        
        # === CONTRATO ===
        ('No. Contrato Venta', FILL_CONTRATO, FONT_BLACK_BOLD, 1),
        ('UM Contrato', FILL_CONTRATO, FONT_BLACK_BOLD, 1),
        ('Precio Contrato', FILL_CONTRATO, FONT_BLACK_BOLD, 1)
    ]
    
    # Fila 1: Grupos
    # Venta: col 17 a 19
    ws_matriz.merge_cells(start_row=1, start_column=17, end_row=1, end_column=19)
    apply_header_style(ws_matriz.cell(row=1, column=17), "=== VENTA ===", FILL_NAVY, FONT_WHITE_BOLD)
    
    ws_matriz.merge_cells(start_row=1, start_column=20, end_row=1, end_column=21)
    apply_header_style(ws_matriz.cell(row=1, column=20), "=== COSTO ===", FILL_NAVY, FONT_WHITE_BOLD)
    
    ws_matriz.merge_cells(start_row=1, start_column=22, end_row=1, end_column=24)
    apply_header_style(ws_matriz.cell(row=1, column=22), "=== VALIDACIONES ===", FILL_NAVY, FONT_WHITE_BOLD)
    
    ws_matriz.merge_cells(start_row=1, start_column=25, end_row=1, end_column=27)
    apply_header_style(ws_matriz.cell(row=1, column=25), "=== CONTRATO ===", FILL_NAVY, FONT_WHITE_BOLD)
    
    # Fila 2: Encabezados de columnas
    for col_idx, (col_name, fill, font, width) in enumerate(headers, start=1):
        cell = ws_matriz.cell(row=2, column=col_idx)
        cell.value = col_name
        cell.fill = fill
        cell.font = font
        cell.border = THIN_BORDER
        cell.alignment = ALIGN_CENTER
        ws_matriz.column_dimensions[openpyxl.utils.get_column_letter(col_idx)].width = 15
        
    # Congelar
    ws_matriz.freeze_panes = "A3"
    
    # Escribir Datos
    df_out = df_matriz[[
        'Concat1', 'Concat2', 'Sociedad', 'Clase Cond. MP', 'Clase Cond. Flete',
        'Ship From', 'Nombre SF', 'Centro', 'Desc. Centro', 'Destino',
        'Cond. Expedición', 'Material', 'Denominación', 'PV',
        'Inicio Vigencia', 'Fin Vigencia', 'Importe_MP', 'Importe_Flete', 'UM Venta',
        'Importe_Costo', 'UM Costo', 'Validacion 1', 'Validacion 2', 'Semaforo',
        'No. Contrato Venta', 'UM Contrato', 'Precio Contrato'
    ]]
    
    for r_idx, row in enumerate(df_out.itertuples(index=False), start=3):
        for c_idx, val in enumerate(row, start=1):
            cell = ws_matriz.cell(row=r_idx, column=c_idx)
            # Manejo None / Nan (pd.isna falla con strings, usar try/except)
            try:
                is_na = pd.isna(val)
            except (TypeError, ValueError):
                is_na = False
            if is_na:
                cell.value = ""
            else:
                cell.value = val
                
            cell.font = FONT_NORMAL
            cell.border = THIN_BORDER
            
            # Formatos especiales por columna
            if c_idx in [17, 18, 20, 22, 23, 27]: # Importes
                cell.number_format = '#,##0.00'
                
            # Semáforo Col 24
            if c_idx == 24:
                sem = cell.value
                if sem == 'SIN_FLETE' or sem == 'SIN_PV' or sem == 'SIN_COSTO' or sem == 'SIN_MP':
                    cell.fill = FILL_ORANGE
                elif sem == 'DIFERENCIA':
                    cell.fill = FILL_WARN
                elif sem == 'DISCREPANCIA_ORG':
                    cell.fill = FILL_RED
                    
            # Validacion 2 amarilla si abs > 1
            if c_idx == 23 and isinstance(val, (int, float)) and pd.notna(val):
                if abs(val) > 1:
                    cell.fill = FILL_WARN
                    
    # ==========================
    # Hoja 2: Dashboard
    # ==========================
    ws_dash = wb.create_sheet(title="Dashboard")
    ws_dash.merge_cells("A1:F2")
    apply_header_style(ws_dash.cell(row=1, column=1), "DASHBOARD CEMEX - MATRIZ INTEGRADA", FILL_NAVY, Font(color=CLR_WHITE, bold=True, size=16))
    
    # KPIs
    totales = len(df_matriz)
    vigentes = len(df_matriz[df_matriz['Fin Vigencia'] == '99991231']) # Asumiendo string
    imp_prom = df_matriz['Importe_MP'].mean() if not df_matriz['Importe_MP'].empty else 0
    con_dif = len(df_matriz[df_matriz['Semaforo'] == 'DIFERENCIA'])
    sin_flete = len(df_matriz[df_matriz['Semaforo'] == 'SIN_FLETE'])
    sin_costo = len(df_matriz[df_matriz['Semaforo'] == 'SIN_COSTO'])
    
    pct_flete = (sin_flete / totales * 100) if totales > 0 else 0
    pct_costo = (sin_costo / totales * 100) if totales > 0 else 0
    
    kpis = [
        ("Total Registros", totales),
        ("Registros Vigentes", vigentes),
        ("Importe MP Promedio", round(imp_prom, 2)),
        ("Registros c/ Dif>1", con_dif),
        ("% Sin Flete", f"{pct_flete:.1f}%"),
        ("% Sin Costo", f"{pct_costo:.1f}%")
    ]
    
    for i, (k, v) in enumerate(kpis):
        r = 4 + i
        ws_dash.cell(row=r, column=2, value=k).font = FONT_BLACK_BOLD
        ws_dash.cell(row=r, column=3, value=v).font = FONT_NORMAL
        
    # Top anomalías
    ws_dash.cell(row=11, column=2, value="Últimas Anomalías Detectadas").font = FONT_BLACK_BOLD
    anomalias = df_matriz[df_matriz['Semaforo'] != ''].head(50)
    
    if not anomalias.empty:
        for c_idx, col in enumerate(['Concat1', 'Centro', 'Material', 'Semaforo', 'Validacion 2'], start=2):
            cell = ws_dash.cell(row=12, column=c_idx)
            cell.value = col
            cell.fill = FILL_HEADER
            cell.font = FONT_BLACK_BOLD
            
        for r_idx, row in enumerate(anomalias.itertuples(), start=13):
            ws_dash.cell(row=r_idx, column=2, value=row.Concat1)
            ws_dash.cell(row=r_idx, column=3, value=row.Centro)
            ws_dash.cell(row=r_idx, column=4, value=row.Material)
            ws_dash.cell(row=r_idx, column=5, value=row.Semaforo)
            ws_dash.cell(row=r_idx, column=6, value=row._23 if hasattr(row, '_23') else getattr(row, 'Validacion 2', ''))
            
    ws_dash.cell(row=70, column=2, value=f"Generado el: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    
    # ==========================
    # Hojas Ocultas
    # ==========================
    if raw_data_frames:
        for name, df in zip(['_MP', '_Flete', '_TRAOPE', '_Contratos'], raw_data_frames):
            if df is not None and not df.empty:
                ws_h = wb.create_sheet(title=name)
                for r in dataframe_to_rows(df, index=False, header=True):
                    ws_h.append(r)
                ws_h.sheet_state = 'hidden'
                
    wb.save(out_file)
    logging.info(f"¡Excel guardado exitosamente!")

# =============================================================================
# CLI
# =============================================================================

def main():
    parser = argparse.ArgumentParser(description="Orquestador Unificado CEMEX Matriz Ventas")
    parser.add_argument('--fuente', type=str, help="Ruta al archivo Excel .xlsm (Modo A)")
    parser.add_argument('--mp', type=str, help="Ruta al archivo TXT de Material (Modo B)")
    parser.add_argument('--flete', type=str, help="Ruta al archivo TXT de Flete (Modo B)")
    parser.add_argument('--cedis', type=str, help="Filtro opcional por código de Centro (CEDIS)")
    parser.add_argument('--output', type=str, default="./_salidas_integradas", help="Carpeta de salida")
    parser.add_argument('--refresh-cache', action='store_true', help="Ignorar caché y reprocesar el Excel")
    
    args = parser.parse_args()
    
    if len(sys.argv) == 1:
        parser.print_help()
        sys.exit(0)
        
    inicio = time.time()
        
    if args.fuente:
        # Modo A
        # Manejo de wildcards o paths complejos
        files = glob.glob(args.fuente)
        if not files:
            logging.error(f"No se encontró el archivo: {args.fuente}")
            sys.exit(1)
        fuente = files[0]
        
        data_raw = cargar_excel(fuente, force_refresh=args.refresh_cache)
        df_matriz, df_mp, df_flete, df_traope, df_contratos = procesar_datos(data_raw, cedis=args.cedis, modo_a=True)
        escribir_excel(df_matriz, args.cedis, args.output, [df_mp, df_flete, df_traope, df_contratos])
        
    elif args.mp and args.flete:
        # Modo B
        data_raw = {
            'MP': parse_txt(args.mp),
            'Flete': parse_txt(args.flete)
        }
        df_matriz, _, _, _, _ = procesar_datos(data_raw, cedis=args.cedis, modo_a=False)
        logging.info("Modo B no genera el Excel de la misma forma, funcionalidad de ejemplo.")
    else:
        logging.error("Debe proveer --fuente (Modo A) o --mp y --flete (Modo B)")
        
    logging.info(f"Listo en {time.time() - inicio:.2f}s")

if __name__ == "__main__":
    main()
