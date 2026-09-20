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
from openpyxl.utils import get_column_letter

# Configuración de Logging
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

# =============================================================================
# COLORES CEMEX Y ESTILOS OPENPYXL
# =============================================================================
CLR_NAVY        = '002D72'  # Azul CEMEX Primario
CLR_VENTA       = 'BDD7EE'  # Azul claro - columnas de venta
CLR_COSTO       = 'FCE4D6'  # Salmón claro - columnas de costo
CLR_GOB         = 'E7E6E6'  # Gris plata - gobernanza y autorizaciones
CLR_VAL         = 'E2EFDA'  # Verde claro - validaciones de margen
CLR_CONTRATO    = 'FFF2CC'  # Amarillo claro - contratos de venta
CLR_WARN        = 'FFFF00'  # Amarillo - diferencia > 1
CLR_RED         = 'FFC7CE'  # Rojo suave - alerta nacional / fuera de rango
CLR_ORANGE      = 'FCE4D6'  # Naranja suave - dato faltante
CLR_GREEN       = 'C6EFCE'  # Verde éxito - Champion
CLR_REGIONAL    = 'FFF2CC'  # Amarillo suave - Regional
CLR_HEADER      = 'D9D9D9'  # Gris - encabezados normales
CLR_WHITE       = 'FFFFFF'

FILL_NAVY       = PatternFill(start_color=CLR_NAVY, end_color=CLR_NAVY, fill_type='solid')
FILL_VENTA      = PatternFill(start_color=CLR_VENTA, end_color=CLR_VENTA, fill_type='solid')
FILL_COSTO      = PatternFill(start_color=CLR_COSTO, end_color=CLR_COSTO, fill_type='solid')
FILL_GOB        = PatternFill(start_color=CLR_GOB, end_color=CLR_GOB, fill_type='solid')
FILL_VAL        = PatternFill(start_color=CLR_VAL, end_color=CLR_VAL, fill_type='solid')
FILL_CONTRATO   = PatternFill(start_color=CLR_CONTRATO, end_color=CLR_CONTRATO, fill_type='solid')
FILL_WARN       = PatternFill(start_color=CLR_WARN, end_color=CLR_WARN, fill_type='solid')
FILL_RED        = PatternFill(start_color=CLR_RED, end_color=CLR_RED, fill_type='solid')
FILL_ORANGE     = PatternFill(start_color=CLR_ORANGE, end_color=CLR_ORANGE, fill_type='solid')
FILL_GREEN      = PatternFill(start_color=CLR_GREEN, end_color=CLR_GREEN, fill_type='solid')
FILL_REGIONAL   = PatternFill(start_color=CLR_REGIONAL, end_color=CLR_REGIONAL, fill_type='solid')
FILL_HEADER     = PatternFill(start_color=CLR_HEADER, end_color=CLR_HEADER, fill_type='solid')
FILL_WHITE      = PatternFill(start_color=CLR_WHITE, end_color=CLR_WHITE, fill_type='solid')

FONT_WHITE_BOLD = Font(color=CLR_WHITE, bold=True, name='Calibri', size=10)
FONT_BLACK_BOLD = Font(color='000000', bold=True, name='Calibri', size=10)
FONT_NORMAL     = Font(color='000000', name='Calibri', size=9)

ALIGN_CENTER    = Alignment(horizontal='center', vertical='center')
ALIGN_LEFT      = Alignment(horizontal='left', vertical='center')
ALIGN_RIGHT     = Alignment(horizontal='right', vertical='center')

THIN_BORDER = Border(
    left=Side(style='thin', color='D9D9D9'), 
    right=Side(style='thin', color='D9D9D9'), 
    top=Side(style='thin', color='D9D9D9'), 
    bottom=Side(style='thin', color='D9D9D9')
)

# =============================================================================
# FUNCIONES DE CÁLCULO DE MÁRGENES Y VALIDACIONES
# =============================================================================

def calc_validacion1(um_venta, um_costo, cond_exp, importe_mp, importe_flete, pv):
    """
    Costo Teórico / Margen Esperado (Validación 1):
    - Si UM_venta=TN y UM_costo=TN: MP + Flete (ya están en TN).
    - Si Cond.Expedicion=1 (Entregado): PV * MP.
    - Else: (MP + (Flete or 0)) * PV.
    - Si falta PV en ramas que lo necesitan: None.
    """
    mp = importe_mp if pd.notna(importe_mp) else None
    flete = importe_flete if pd.notna(importe_flete) else None
    
    if mp is None:
        return None
        
    um_v = str(um_venta).strip().upper() if pd.notna(um_venta) else ""
    um_c = str(um_costo).strip().upper() if pd.notna(um_costo) else ""
    
    if um_v == 'TN' and um_c == 'TN':
        return mp + (flete if flete is not None else 0)
    
    if pd.isna(pv):
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
    """Desviación Real vs Teórico: Costo_TRAOPE - Validación_1."""
    if pd.isna(importe_costo) or pd.isna(validacion1):
        return None
    return importe_costo - validacion1

def calc_mop(importe_mp, um_venta, importe_costo, um_costo, pv):
    """
    Margen Operativo de Material (MOP %):
    MOP = (Precio Venta MP - Costo Compra MP Homologado) / Precio Venta MP
    Fórmula oficial de Excel:
    =@SI.CONJUNTO(
        [UM COMP]="M3", SI([MP COMP]=0, 0, ([MP VTA]-[MP COMP])/[MP VTA]),
        [UM COMP]="TN", SI([MP COMP]=0, 0, ([MP VTA]-[MP COMP])/[MP VTA])
    )
    """
    if pd.isna(importe_mp) or importe_mp <= 0:
        return None
    if pd.isna(importe_costo) or importe_costo <= 0:
        return None
        
    um_v = str(um_venta).strip().upper() if pd.notna(um_venta) else ""
    um_c = str(um_costo).strip().upper() if pd.notna(um_costo) else ""
    
    costo_homo = importe_costo
    if um_v == um_c:
        costo_homo = importe_costo
    elif um_v == 'TN' and um_c == 'M3':
        if pd.notna(pv) and pv > 0:
            costo_homo = importe_costo / pv
        else:
            return None
    elif um_v == 'M3' and um_c == 'TN':
        if pd.notna(pv) and pv > 0:
            costo_homo = importe_costo * pv
        else:
            return None
            
    return (importe_mp - costo_homo) / importe_mp

def eval_autorizacion(tipo_operacion, precio_venta, precio_referencia, mop):
    """
    Niveles de Autorización y Alertas de Gobernanza:
    
    1. CANTERAS PROPIAS:
       - Si Precio_Venta >= Precio_Referencia -> 'Autoriza: Champion'
       - Si Descuento <= 5% -> 'Autoriza: Regional | Requiere Justificacion'
       - Si Descuento > 5%  -> 'Alerta: Requiere Vo.Bo. Nacional | Requiere Justificacion'
       
    2. TRADING:
       - Si Margen_Material > 8% -> 'Autoriza: Champion'
       - Si Margen_Material entre 5% y 7.99% -> 'Autoriza: Regional'
       - Si Margen_Material < 5% -> 'Alerta Fuera de Rango: Requiere revision puntual con Nacional'
    """
    if tipo_operacion == 'TRADING':
        if mop is None or pd.isna(mop):
            return "Pendiente de Costo TRAOPE"
        if mop > 0.08:
            return "Autoriza: Champion"
        elif mop >= 0.05:
            return "Autoriza: Regional"
        else:
            return "Alerta Fuera de Rango: Requiere revision puntual con Nacional"
            
    elif tipo_operacion == 'CANTERAS PROPIAS':
        if pd.isna(precio_venta) or precio_venta <= 0:
            return "Sin Precio MP"
        if pd.isna(precio_referencia) or precio_referencia <= 0:
            return "Autoriza: Champion"
            
        descuento = (precio_referencia - precio_venta) / precio_referencia
        if precio_venta >= precio_referencia or descuento <= 0:
            return "Autoriza: Champion"
        elif descuento <= 0.05:
            return "Autoriza: Regional | Requiere Justificacion"
        else:
            return "Alerta: Requiere Vo.Bo. Nacional | Requiere Justificacion"
            
    return ""

def eval_semaforo(row):
    """
    Semáforo de inconsistencias operativas:
    - SIN_MP: falta precio de venta de material
    - SIN_FLETE: en modalidad entregada (01/04), no se encontró flete
    - SIN_COSTO: en trading, no se encontró orden de compra
    - SIN_PV: falta factor de peso volumétrico
    - DIFERENCIA: desviación mayor a $1 entre costo real y teórico
    - DISCREPANCIA_ORG: condición inválida por sociedad
    """
    mp = row.get('Importe MP') if 'Importe MP' in row else row.get('Importe_MP')
    if pd.isna(mp) or mp <= 0:
        return 'SIN_MP'
    
    cond_exp = str(row.get('Cond. Expedición', '')).strip()
    flete = row.get('Importe Flete') if 'Importe Flete' in row else row.get('Importe_Flete')
    
    # En 02 (Recogido), el cliente manda camiones -> NO lleva flete, NO es anomalía
    if cond_exp in ('1', '01', '4', '04'):
        if pd.isna(flete):
            return 'SIN_FLETE'
            
    tipo_op = row.get('Tipo Operación', '')
    costo = row.get('Costo Total TRAOPE') if 'Costo Total TRAOPE' in row else row.get('MP Compra (Costo Material)')
    if costo is None or pd.isna(costo):
        costo = row.get('Importe Costo')
    if tipo_op == 'TRADING' and (pd.isna(costo) or costo is None or costo <= 0):
        return 'SIN_COSTO'
        
    if pd.isna(row.get('Validacion 1')):
        um_v = str(row.get('UM Venta', '')).strip().upper()
        um_c = str(row.get('UM Costo', '')).strip().upper()
        if not (um_v == 'TN' and um_c == 'TN'):
            if pd.isna(row.get('PV')):
                return 'SIN_PV'
                
    v2 = row.get('Validacion 2')
    if pd.notna(v2) and abs(v2) > 1:
        return 'DIFERENCIA'
        
    org = str(row.get('Sociedad', '')).strip()
    clase_mp = str(row.get('Clase Cond. MP', '')).strip()
    if org == '7100' and (clase_mp in ['ZMA6', 'ZMP1']):
        return 'DISCREPANCIA_ORG'
        
    return 'OK'

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
            return None
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
        data[hoja] = pd.read_excel(xls, sheet_name=hoja, dtype=str)
        
    cache_data(cache_file, data)
    return data

def normalizar_etiq(texto):
    import unicodedata
    sin_tildes = "".join(
        c for c in unicodedata.normalize("NFD", texto)
        if unicodedata.category(c) != "Mn"
    )
    return sin_tildes.strip().lower()

def fecha_ddmmyyyy_a_num(fecha_str):
    try:
        partes = fecha_str.strip().split('.')
        if len(partes) == 3:
            dd, mm, yyyy = partes
            return int(f"{yyyy}{mm.zfill(2)}{dd.zfill(2)}")
    except Exception:
        pass
    return 0

def parse_txt(archivo_txt):
    logging.info(f"Parseando txt: {archivo_txt}")
    registros = []
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
            fechas = [c.strip() for c in campos if len(c.strip()) == 10 and c.strip()[2] == '.' and c.strip()[5] == '.']
            if fechas:
                current_reg['Inicio Validez'] = fechas[0]
                current_reg['Valido a'] = fechas[-1]
                current_reg['Valido a_num'] = fecha_ddmmyyyy_a_num(fechas[-1])
        elif "MXN" in campos:
            mxn_idx = campos.index("MXN")
            importe = None
            for c in reversed(campos[:mxn_idx]):
                val = c.strip().replace(',', '').replace(' ', '')
                try:
                    importe = float(val)
                    break
                except (ValueError, TypeError):
                    pass
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
    if not df.empty:
        if 'Valido a_num' in df.columns:
            df = df[df['Valido a_num'] >= 99991231]
        logging.info(f"  → {len(df)} registros vigentes tras filtro 9999")

    return df

# =============================================================================
# PROCESAMIENTO Y CRUCE UNIFICADO
# =============================================================================

def procesar_datos(data, cedis=None, modo_a=True):
    logging.info("Cruzando datos y calculando...")
    
    if modo_a:
        df_mp = data['PVTA_MAT VK13'].copy()
        df_flete = data['PVTA_FTE VK13'].copy()
        df_traope = data['CONT_COMPRA TRAOPE'].copy()
        df_pv = data['PESO_VOL'].copy()
        df_contratos = data['CONT_VTA ZSDD4501'].copy()
        
        # Filtros de vigencia abierta
        df_mp['Valido a_num'] = pd.to_numeric(df_mp['Valido a'], errors='coerce')
        df_mp = df_mp[df_mp['Valido a_num'] >= 99991231]
        
        df_flete['Fin Validez_num'] = pd.to_numeric(df_flete['Fin Validez'], errors='coerce')
        df_flete = df_flete[df_flete['Fin Validez_num'] >= 99991231]
        
        if 'Validez a' in df_traope.columns:
            df_traope['Validez a_num'] = pd.to_numeric(df_traope['Validez a'].str.replace('.', '').str.replace('-', ''), errors='coerce')
            df_traope = df_traope[df_traope['Validez a_num'] >= 20291231]
            
        if 'Fecha Fin validez' in df_contratos.columns:
            df_contratos['Fecha Fin validez_num'] = pd.to_numeric(df_contratos['Fecha Fin validez'], errors='coerce')
            df_contratos = df_contratos[df_contratos['Fecha Fin validez_num'] >= 20291231]
            
        # Normalización de llaves
        for df in [df_mp, df_flete]:
            df['Shipfrom'] = df['Shipfrom'].apply(str_val)
            df['Centro'] = df['Centro'].apply(str_val)
            df['Destinatario'] = df['Destinatario'].apply(str_val)
            df['Material'] = df['Material'].apply(str_val)
            df['Concat1'] = df['Shipfrom'] + '-' + df['Centro'] + '-' + df['Destinatario'] + '-' + df['Material']
            df['Concat2'] = df['Shipfrom'] + '-' + df['Centro'] + '-' + df['Material']
            
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
        pv_idx = df_pv.drop_duplicates('Material').set_index('Material')
        
        # TRAOPE: preparación de llaves Concat1 y Concat2
        df_traope['Ship From'] = df_traope['Ship From'].apply(str_val)
        df_traope['Centro'] = df_traope['Centro'].apply(str_val)
        col_destino = 'Destino' if 'Destino' in df_traope.columns else None
        if col_destino:
            df_traope[col_destino] = df_traope[col_destino].apply(str_val)
            df_traope['Concat1'] = df_traope['Ship From'] + '-' + df_traope['Centro'] + '-' + df_traope[col_destino] + '-' + df_traope['Material'].apply(str_val)
        else:
            df_traope['Concat1'] = 'NA'
            
        df_traope['Material'] = df_traope['Material'].apply(str_val)
        df_traope['Concat2'] = df_traope['Ship From'] + '-' + df_traope['Centro'] + '-' + df_traope['Material']
        
        # Índices TRAOPE
        traope_c1 = df_traope[df_traope['Concat1'] != 'NA'].drop_duplicates('Concat1').set_index('Concat1')
        traope_c2 = df_traope.drop_duplicates('Concat2').set_index('Concat2')
        
        # Fletes: SOLO cruce exacto por Concat1
        flete_c1 = df_flete.drop_duplicates('Concat1').set_index('Concat1')
        
        # Contratos de venta: indexados por Ruta A
        if 'Ruta A' in df_contratos.columns:
            df_contratos['Llave_Contrato'] = df_contratos['Ruta A'].apply(str_val)
        else:
            df_contratos['Llave_Contrato'] = 'NA'
        contratos_idx = df_contratos.drop_duplicates('Llave_Contrato').set_index('Llave_Contrato')
        
        # Precios de referencia base por Material para Canteras Propias (promedio ponderado/media)
        precios_ref_mat = df_mp.groupby('Material')['Importe'].mean().to_dict()
        
        # =====================================================================
        # CONSTRUCCIÓN DE LA MATRIZ FILA POR FILA
        # =====================================================================
        filas = []
        for _, row in df_mp.iterrows():
            c1 = row['Concat1']
            c2 = row['Concat2']
            mat = row['Material']
            
            # 1. Flete (match exacto)
            f_row = flete_c1.loc[c1] if c1 in flete_c1.index else None
            
            # 2. PV
            pv_val = pv_idx.loc[mat, 'PV_calc'] if mat in pv_idx.index else None
            pv_desc = pv_idx.loc[mat, 'Texto de material'] if mat in pv_idx.index else ''
            
            # 3. Contratos de Venta
            llave_co = f"{row['Centro']}-{row['Shipfrom']}-{row['Destinatario']}-{mat}"
            co_row = contratos_idx.loc[llave_co] if llave_co in contratos_idx.index else None
            
            # Condición de Expedición (de Flete o del Contrato)
            cond_exp = ''
            if f_row is not None and 'Cond. Expedición' in f_row.index and pd.notna(f_row['Cond. Expedición']):
                cond_exp = str(f_row['Cond. Expedición']).strip()
            elif co_row is not None and 'Condición Exp.' in co_row.index and pd.notna(co_row['Condición Exp.']):
                cond_exp = str(co_row['Condición Exp.']).strip()
                
            # 4. TRAOPE (Costo de Compra)
            # Match Concat1 exacto; si no, Concat2 de respaldo si no es entrega directa estricta
            t_row = None
            if c1 in traope_c1.index:
                t_row = traope_c1.loc[c1]
            elif c2 in traope_c2.index:
                t_row = traope_c2.loc[c2]
                
            imp_mp = row['Importe']
            imp_flete = f_row['Importe'] if f_row is not None else None
            um_venta = row.get('Unidad', '')
            
            imp_costo_total = None
            imp_flete_compra = None
            mp_compra = None
            um_costo = None
            
            if t_row is not None:
                if 'Precio neto pedido' in t_row.index:
                    imp_costo_total = pd.to_numeric(t_row['Precio neto pedido'], errors='coerce')
                if 'Importe Condición' in t_row.index and pd.notna(t_row['Importe Condición']):
                    imp_flete_compra = pd.to_numeric(t_row['Importe Condición'], errors='coerce')
                    
                if 'UM Precio Pedido' in t_row.index and pd.notna(t_row['UM Precio Pedido']):
                    um_costo = str(t_row['UM Precio Pedido']).strip()
                elif 'UM' in t_row.index and pd.notna(t_row['UM']):
                    um_costo = str(t_row['UM']).strip()
                elif 'UMB' in t_row.index and pd.notna(t_row['UMB']):
                    um_costo = str(t_row['UMB']).strip()
                    
                # Calcular Costo MP Compra puro (descontando componente de flete/condición)
                if imp_costo_total is not None:
                    if imp_flete_compra is not None and imp_flete_compra > 0 and imp_costo_total > imp_flete_compra:
                        mp_compra = imp_costo_total - imp_flete_compra
                    else:
                        mp_compra = imp_costo_total
                        
            # Clasificación Tipo de Operación
            nombre_sf_upper = str(t_row['Nombre SF'] if t_row is not None and 'Nombre SF' in t_row.index else '').upper()
            if (imp_costo_total is not None and imp_costo_total > 0) or 'TP-' in nombre_sf_upper or 'SF TP' in nombre_sf_upper or 'TC-' in nombre_sf_upper:
                tipo_operacion = 'TRADING'
            else:
                tipo_operacion = 'CANTERAS PROPIAS'
                
            # Cálculos de Márgenes y Gobernanza (MOP exclusivamente sobre Material)
            mop = calc_mop(imp_mp, um_venta, mp_compra, um_costo, pv_val)
            precio_ref = precios_ref_mat.get(mat, imp_mp)
            nivel_aut = eval_autorizacion(tipo_operacion, imp_mp, precio_ref, mop)
            
            # Validaciones Integrales
            val1 = calc_validacion1(um_venta, um_costo, cond_exp, imp_mp, imp_flete, pv_val)
            val2 = calc_validacion2(imp_costo_total, val1)
            
            # Armado de fila con precio junto a cada condición
            fila_dict = {
                'Concat1': c1,
                'Concat2': c2,
                'Sociedad': row.get('Org. Ventas', ''),
                'Ship From': row.get('Shipfrom', ''),
                'Nombre SF': t_row['Nombre SF'] if t_row is not None and 'Nombre SF' in t_row.index else '',
                'Centro': row.get('Centro', ''),
                'Desc. Centro': t_row['Descripción Centro'] if t_row is not None and 'Descripción Centro' in t_row.index else '',
                'Destino': row.get('Destinatario', ''),
                'Material': mat,
                'Denominación': pv_desc,
                'PV': pv_val,
                'Inicio Vigencia': row.get('Inicio Validez', ''),
                'Fin Vigencia': row.get('Valido a', ''),
                
                # VENTA: Precio junto a condición
                'Clase Cond. MP': row.get('Clase Cond.', ''),
                'Importe MP': imp_mp,
                'UM Venta': um_venta,
                'Clase Cond. Flete': f_row['Clase Cond.'] if f_row is not None else '',
                'Importe Flete': imp_flete,
                'Cond. Expedición': cond_exp,
                
                # COSTO: Desglosado Total, Flete Compra y Costo Material Puro
                'Costo Total TRAOPE': imp_costo_total,
                'Flete Compra': imp_flete_compra,
                'MP Compra (Costo Material)': mp_compra,
                'UM Costo': um_costo,
                'Margen Material (MOP %)': mop,
                
                # GOBERNANZA Y AUTORIZACIÓN
                'Tipo Operación': tipo_operacion,
                'Precio Referencia': precio_ref,
                'Nivel Autorización / Alerta': nivel_aut,
                
                # VALIDACIÓN MARGEN INTEGRAL
                'Validacion 1': val1,
                'Validacion 2': val2,
                'Semaforo': '',
                
                # CONTRATO VENTA
                'No. Contrato Venta': co_row['Llave'] if co_row is not None and 'Llave' in co_row.index else '',
                'UM Contrato': co_row['UM'] if co_row is not None and 'UM' in co_row.index else '',
                'Precio Contrato': pd.to_numeric(co_row['Precio Neto'], errors='coerce') if co_row is not None and 'Precio Neto' in co_row.index else None
            }
            
            # Evaluación del Semáforo
            fila_dict['Semaforo'] = eval_semaforo(fila_dict)
            filas.append(fila_dict)
            
        matriz_df = pd.DataFrame(filas)
        return matriz_df, df_mp, df_flete, df_traope, df_contratos
    else:
        return pd.DataFrame(), None, None, None, None

# =============================================================================
# ESCRITURA EN EXCEL FORMATEADO
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
    
    # -------------------------------------------------------------------------
    # Hoja 1: Matriz
    # -------------------------------------------------------------------------
    ws_matriz = wb.active
    ws_matriz.title = "Matriz"
    
    # Definición de 31 columnas organizadas lógicamente
    headers = [
        ('Concat1', FILL_HEADER, FONT_BLACK_BOLD, 30),
        ('Concat2', FILL_HEADER, FONT_BLACK_BOLD, 25),
        ('Sociedad', FILL_HEADER, FONT_BLACK_BOLD, 10),
        ('Ship From', FILL_HEADER, FONT_BLACK_BOLD, 12),
        ('Nombre SF', FILL_HEADER, FONT_BLACK_BOLD, 22),
        ('Centro', FILL_HEADER, FONT_BLACK_BOLD, 10),
        ('Desc. Centro', FILL_HEADER, FONT_BLACK_BOLD, 20),
        ('Destino', FILL_HEADER, FONT_BLACK_BOLD, 12),
        ('Material', FILL_HEADER, FONT_BLACK_BOLD, 12),
        ('Denominación', FILL_HEADER, FONT_BLACK_BOLD, 30),
        ('PV', FILL_HEADER, FONT_BLACK_BOLD, 10),
        ('Inicio Vigencia', FILL_HEADER, FONT_BLACK_BOLD, 14),
        ('Fin Vigencia', FILL_HEADER, FONT_BLACK_BOLD, 14),
        
        # === CONDICIONES DE VENTA (Precio al lado de Condición) ===
        ('Clase Cond. MP', FILL_VENTA, FONT_BLACK_BOLD, 14),
        ('Importe MP', FILL_VENTA, FONT_BLACK_BOLD, 14),
        ('UM Venta', FILL_VENTA, FONT_BLACK_BOLD, 10),
        ('Clase Cond. Flete', FILL_VENTA, FONT_BLACK_BOLD, 15),
        ('Importe Flete', FILL_VENTA, FONT_BLACK_BOLD, 14),
        ('Cond. Expedición', FILL_VENTA, FONT_BLACK_BOLD, 14),
        
        # === CONDICIONES DE COMPRA (Desglosado) ===
        ('Costo Total TRAOPE', FILL_COSTO, FONT_BLACK_BOLD, 16),
        ('Flete Compra', FILL_COSTO, FONT_BLACK_BOLD, 14),
        ('MP Compra (Costo Material)', FILL_COSTO, FONT_BLACK_BOLD, 18),
        ('UM Costo', FILL_COSTO, FONT_BLACK_BOLD, 10),
        ('Margen Material (MOP %)', FILL_COSTO, FONT_BLACK_BOLD, 18),
        
        # === GOBERNANZA Y AUTORIZACIÓN ===
        ('Tipo Operación', FILL_GOB, FONT_BLACK_BOLD, 16),
        ('Precio Referencia Base', FILL_GOB, FONT_BLACK_BOLD, 16),
        ('Nivel Autorización / Alerta', FILL_GOB, FONT_BLACK_BOLD, 36),
        
        # === VALIDACIÓN DE MARGEN INTEGRAL ===
        ('Validación 1', FILL_VAL, FONT_BLACK_BOLD, 14),
        ('Validación 2', FILL_VAL, FONT_BLACK_BOLD, 14),
        ('Semáforo', FILL_VAL, FONT_BLACK_BOLD, 16),
        
        # === CONTRATO DE VENTA ===
        ('No. Contrato Venta', FILL_CONTRATO, FONT_BLACK_BOLD, 18),
        ('UM Contrato', FILL_CONTRATO, FONT_BLACK_BOLD, 12),
        ('Precio Contrato', FILL_CONTRATO, FONT_BLACK_BOLD, 14)
    ]
    
    # Fila 1: Grupos Superiores Fusionados
    # 1. RUTA (1-13)
    ws_matriz.merge_cells(start_row=1, start_column=1, end_row=1, end_column=13)
    apply_header_style(ws_matriz.cell(row=1, column=1), "=== DATOS DE RUTA Y MATERIAL ===", FILL_NAVY, FONT_WHITE_BOLD)
    
    # 2. VENTA (14-19)
    ws_matriz.merge_cells(start_row=1, start_column=14, end_row=1, end_column=19)
    apply_header_style(ws_matriz.cell(row=1, column=14), "=== CONDICIONES DE VENTA ===", FILL_NAVY, FONT_WHITE_BOLD)
    
    # 3. COSTO (20-24)
    ws_matriz.merge_cells(start_row=1, start_column=20, end_row=1, end_column=24)
    apply_header_style(ws_matriz.cell(row=1, column=20), "=== CONDICIONES DE COMPRA (COSTO) ===", FILL_NAVY, FONT_WHITE_BOLD)
    
    # 4. GOBERNANZA (25-27)
    ws_matriz.merge_cells(start_row=1, start_column=25, end_row=1, end_column=27)
    apply_header_style(ws_matriz.cell(row=1, column=25), "=== GOBERNANZA Y AUTORIZACIÓN ===", FILL_NAVY, FONT_WHITE_BOLD)
    
    # 5. VALIDACIÓN (28-30)
    ws_matriz.merge_cells(start_row=1, start_column=28, end_row=1, end_column=30)
    apply_header_style(ws_matriz.cell(row=1, column=28), "=== VALIDACIÓN DE MARGEN ===", FILL_NAVY, FONT_WHITE_BOLD)
    
    # 6. CONTRATOS (31-33)
    ws_matriz.merge_cells(start_row=1, start_column=31, end_row=1, end_column=33)
    apply_header_style(ws_matriz.cell(row=1, column=31), "=== CONTRATO DE VENTA ===", FILL_NAVY, FONT_WHITE_BOLD)
    
    # Fila 2: Encabezados individuales
    for col_idx, (col_name, fill, font, width) in enumerate(headers, start=1):
        cell = ws_matriz.cell(row=2, column=col_idx)
        cell.value = col_name
        cell.fill = fill
        cell.font = font
        cell.border = THIN_BORDER
        cell.alignment = ALIGN_CENTER
        ws_matriz.column_dimensions[get_column_letter(col_idx)].width = width
        
    ws_matriz.freeze_panes = "A3"
    
    # Orden exacto de columnas para el volcado (33 columnas)
    cols_order = [
        'Concat1', 'Concat2', 'Sociedad', 'Ship From', 'Nombre SF',
        'Centro', 'Desc. Centro', 'Destino', 'Material', 'Denominación',
        'PV', 'Inicio Vigencia', 'Fin Vigencia',
        'Clase Cond. MP', 'Importe MP', 'UM Venta', 'Clase Cond. Flete', 'Importe Flete', 'Cond. Expedición',
        'Costo Total TRAOPE', 'Flete Compra', 'MP Compra (Costo Material)', 'UM Costo', 'Margen Material (MOP %)',
        'Tipo Operación', 'Precio Referencia', 'Nivel Autorización / Alerta',
        'Validacion 1', 'Validacion 2', 'Semaforo',
        'No. Contrato Venta', 'UM Contrato', 'Precio Contrato'
    ]
    df_out = df_matriz[cols_order]
    
    for r_idx, row in enumerate(df_out.itertuples(index=False), start=3):
        for c_idx, val in enumerate(row, start=1):
            cell = ws_matriz.cell(row=r_idx, column=c_idx)
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
            
            # Formatos numéricos y alineaciones
            # Importes monetarios
            if c_idx in [15, 18, 20, 21, 22, 26, 28, 29, 33]: 
                cell.number_format = '$#,##0.00'
                cell.alignment = ALIGN_RIGHT
            # Margen MOP % (Col 24)
            elif c_idx == 24:
                cell.number_format = '0.0%'
                cell.alignment = ALIGN_RIGHT
                if isinstance(val, (int, float)) and pd.notna(val):
                    if val < 0.05:
                        cell.fill = FILL_RED
                    elif val > 0.08:
                        cell.fill = FILL_GREEN
                    else:
                        cell.fill = FILL_REGIONAL
            # PV (Densidad, Col 11)
            elif c_idx == 11:
                cell.number_format = '#,##0.000'
                cell.alignment = ALIGN_RIGHT
            # Códigos y textos cortos centrados
            elif c_idx in [3, 4, 6, 8, 9, 12, 13, 14, 16, 17, 19, 23, 25, 30, 31, 32]:
                cell.alignment = ALIGN_CENTER
            else:
                cell.alignment = ALIGN_LEFT
                
            # Alertas de Gobernanza (Col 27)
            if c_idx == 27 and isinstance(val, str):
                if 'Champion' in val:
                    cell.fill = FILL_GREEN
                elif 'Regional' in val:
                    cell.fill = FILL_REGIONAL
                elif 'Alerta' in val or 'Nacional' in val:
                    cell.fill = FILL_RED
                    cell.font = FONT_BLACK_BOLD
                    
            # Semáforo (Col 30)
            if c_idx == 30 and isinstance(val, str):
                if val in ('SIN_FLETE', 'SIN_PV', 'SIN_COSTO', 'SIN_MP'):
                    cell.fill = FILL_ORANGE
                elif val == 'DIFERENCIA':
                    cell.fill = FILL_WARN
                elif val == 'DISCREPANCIA_ORG':
                    cell.fill = FILL_RED
                elif val == 'OK':
                    cell.fill = FILL_GREEN
                    
            # Validación 2 amarilla si abs > 1 (Col 29)
            if c_idx == 29 and isinstance(val, (int, float)) and pd.notna(val):
                if abs(val) > 1:
                    cell.fill = FILL_WARN
                    
    # -------------------------------------------------------------------------
    # Hoja 2: Dashboard Ejecutivo
    # -------------------------------------------------------------------------
    ws_dash = wb.create_sheet(title="Dashboard")
    ws_dash.merge_cells("A1:G2")
    apply_header_style(ws_dash.cell(row=1, column=1), "DASHBOARD CEMEX - MATRIZ INTEGRAL Y GOBERNANZA DE PRECIOS", FILL_NAVY, Font(color=CLR_WHITE, bold=True, size=15))
    
    totales = len(df_matriz)
    vigentes = len(df_matriz[df_matriz['Fin Vigencia'] == '99991231'])
    imp_prom = df_matriz['Importe MP'].mean() if 'Importe MP' in df_matriz and not df_matriz['Importe MP'].empty else 0
    con_dif = len(df_matriz[df_matriz['Semaforo'] == 'DIFERENCIA'])
    sin_flete = len(df_matriz[df_matriz['Semaforo'] == 'SIN_FLETE'])
    sin_costo = len(df_matriz[df_matriz['Semaforo'] == 'SIN_COSTO'])
    
    auth_champion = len(df_matriz[df_matriz['Nivel Autorización / Alerta'].str.contains('Champion', na=False)])
    auth_regional = len(df_matriz[df_matriz['Nivel Autorización / Alerta'].str.contains('Regional', na=False)])
    auth_nacional = len(df_matriz[df_matriz['Nivel Autorización / Alerta'].str.contains('Nacional', na=False)])
    
    kpis = [
        ("Total Rutas Analizadas", totales),
        ("Rutas Vigentes (Fin 9999)", vigentes),
        ("Importe MP Promedio", f"${imp_prom:,.2f}"),
        ("Rutas con Diferencia Margen > $1", con_dif),
        ("Rutas Sin Flete Requerido", sin_flete),
        ("Rutas Sin Costo TRAOPE (Trading)", sin_costo),
        ("--- GOBERNANZA Y AUTORIZACIONES ---", ""),
        ("Nivel Champion (Margen > 8% / Sin Descuento)", auth_champion),
        ("Nivel Regional (Margen 5-8% / Desc <= 5%)", auth_regional),
        ("Alerta Nivel Nacional (Margen < 5% / Desc > 5%)", auth_nacional)
    ]
    
    for i, (k, v) in enumerate(kpis):
        r = 4 + i
        c_k = ws_dash.cell(row=r, column=2, value=k)
        c_v = ws_dash.cell(row=r, column=3, value=v)
        c_k.font = FONT_BLACK_BOLD
        c_v.font = FONT_NORMAL
        if "Alerta" in k:
            c_v.fill = FILL_RED
            c_v.font = FONT_BLACK_BOLD
        elif "Champion" in k:
            c_v.fill = FILL_GREEN
            
    ws_dash.cell(row=16, column=2, value="Top 50 Rutas con Alerta o Desviación").font = FONT_BLACK_BOLD
    anomalias = df_matriz[(df_matriz['Semaforo'] != 'OK') | (df_matriz['Nivel Autorización / Alerta'].str.contains('Alerta|Nacional', na=False))].head(50)
    
    if not anomalias.empty:
        for c_idx, col in enumerate(['Concat1', 'Centro', 'Material', 'MOP %', 'Nivel Autorización', 'Semáforo', 'Validación 2'], start=2):
            cell = ws_dash.cell(row=17, column=c_idx)
            cell.value = col
            cell.fill = FILL_HEADER
            cell.font = FONT_BLACK_BOLD
            cell.alignment = ALIGN_CENTER
            
        for r_idx, row in enumerate(anomalias.itertuples(), start=18):
            ws_dash.cell(row=r_idx, column=2, value=getattr(row, 'Concat1', ''))
            ws_dash.cell(row=r_idx, column=3, value=getattr(row, 'Centro', ''))
            ws_dash.cell(row=r_idx, column=4, value=getattr(row, 'Material', ''))
            
            c_mop = ws_dash.cell(row=r_idx, column=5, value=getattr(row, '_22', getattr(row, 'MOP %', '')))
            if isinstance(c_mop.value, (int, float)):
                c_mop.number_format = '0.0%'
                
            ws_dash.cell(row=r_idx, column=6, value=getattr(row, '_25', getattr(row, 'Nivel Autorización / Alerta', '')))
            ws_dash.cell(row=r_idx, column=7, value=getattr(row, 'Semaforo', ''))
            ws_dash.cell(row=r_idx, column=8, value=getattr(row, '_27', getattr(row, 'Validacion 2', '')))
            
    ws_dash.cell(row=72, column=2, value=f"Generado automáticamente el: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    ws_dash.column_dimensions['B'].width = 38
    ws_dash.column_dimensions['C'].width = 18
    ws_dash.column_dimensions['D'].width = 18
    ws_dash.column_dimensions['E'].width = 16
    ws_dash.column_dimensions['F'].width = 36
    ws_dash.column_dimensions['G'].width = 18
    ws_dash.column_dimensions['H'].width = 16
    
    # -------------------------------------------------------------------------
    # Hojas Ocultas de Respaldo
    # -------------------------------------------------------------------------
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
    parser = argparse.ArgumentParser(description="Orquestador Unificado CEMEX Matriz Ventas y Gobernanza")
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
        files = glob.glob(args.fuente)
        if not files:
            logging.error(f"No se encontró el archivo: {args.fuente}")
            sys.exit(1)
        fuente = files[0]
        
        data_raw = cargar_excel(fuente, force_refresh=args.refresh_cache)
        df_matriz, df_mp, df_flete, df_traope, df_contratos = procesar_datos(data_raw, cedis=args.cedis, modo_a=True)
        escribir_excel(df_matriz, args.cedis, args.output, [df_mp, df_flete, df_traope, df_contratos])
        
    elif args.mp and args.flete:
        data_raw = {
            'MP': parse_txt(args.mp),
            'Flete': parse_txt(args.flete)
        }
        df_matriz, _, _, _, _ = procesar_datos(data_raw, cedis=args.cedis, modo_a=False)
        logging.info("Modo B ejecutado exitosamente.")
    else:
        logging.error("Debe proveer --fuente (Modo A) o --mp y --flete (Modo B)")
        
    logging.info(f"Listo en {time.time() - inicio:.2f}s")

if __name__ == "__main__":
    main()
