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
# PALETA DE COLORES CORPORATIVA CEMEX
# =============================================================================
CLR_NAVY        = '002D72'  # Azul CEMEX Principal
CLR_BLUE_ACCENT = '005A9C'  # Azul Secundario
CLR_VENTA       = 'BDD7EE'  # Azul claro - columnas de venta
CLR_COSTO       = 'FCE4D6'  # Salmón claro - columnas de costo
CLR_GOB         = 'E7E6E6'  # Gris plata - gobernanza y autorizaciones
CLR_VAL         = 'E2EFDA'  # Verde claro - validaciones de margen
CLR_CONTRATO    = 'FFF2CC'  # Amarillo claro - contratos de venta
CLR_WARN        = 'FFFF00'  # Amarillo alerta - diferencia > 1
CLR_RED         = 'FFC7CE'  # Rojo suave - alerta nacional / fuera de rango
CLR_DARK_RED    = '9C0006'  # Texto rojo oscuro
CLR_ORANGE      = 'FCE4D6'  # Naranja suave - dato faltante
CLR_GREEN       = 'C6EFCE'  # Verde éxito - Champion
CLR_DARK_GREEN  = '006100'  # Texto verde oscuro
CLR_REGIONAL    = 'FFF2CC'  # Amarillo suave - Regional
CLR_DARK_YEL    = '9C6500'  # Texto amarillo oscuro
CLR_HEADER      = 'D9D9D9'  # Gris - encabezados normales
CLR_CARD_BG     = 'F8FAFC'  # Fondo tarjetas KPI
CLR_CARD_BOR    = 'CBD5E1'  # Borde tarjetas
CLR_WHITE       = 'FFFFFF'

FILL_NAVY       = PatternFill(start_color=CLR_NAVY, end_color=CLR_NAVY, fill_type='solid')
FILL_BLUE_ACC   = PatternFill(start_color=CLR_BLUE_ACCENT, end_color=CLR_BLUE_ACCENT, fill_type='solid')
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
FILL_CARD_BG    = PatternFill(start_color=CLR_CARD_BG, end_color=CLR_CARD_BG, fill_type='solid')
FILL_WHITE      = PatternFill(start_color=CLR_WHITE, end_color=CLR_WHITE, fill_type='solid')

FONT_WHITE_HERO = Font(color=CLR_WHITE, bold=True, name='Segoe UI', size=13)
FONT_WHITE_SUB  = Font(color='CBD5E1', bold=False, name='Segoe UI', size=8.5)
FONT_WHITE_BOLD = Font(color=CLR_WHITE, bold=True, name='Calibri', size=10)
FONT_BLACK_BOLD = Font(color='000000', bold=True, name='Calibri', size=10)
FONT_NORMAL     = Font(color='000000', name='Calibri', size=9)
FONT_KPI_TITLE  = Font(color='475569', bold=True, name='Segoe UI', size=7.5)
FONT_KPI_VAL    = Font(color=CLR_NAVY, bold=True, name='Segoe UI', size=16)
FONT_KPI_SUB    = Font(color='64748B', bold=False, name='Segoe UI', size=7)

ALIGN_CENTER    = Alignment(horizontal='center', vertical='center')
ALIGN_LEFT      = Alignment(horizontal='left', vertical='center')
ALIGN_RIGHT     = Alignment(horizontal='right', vertical='center')

THIN_BORDER = Border(
    left=Side(style='thin', color='CBD5E1'), 
    right=Side(style='thin', color='CBD5E1'), 
    top=Side(style='thin', color='CBD5E1'), 
    bottom=Side(style='thin', color='CBD5E1')
)

CARD_BORDER = Border(
    left=Side(style='thin', color=CLR_CARD_BOR), 
    right=Side(style='thin', color=CLR_CARD_BOR), 
    top=Side(style='thin', color=CLR_CARD_BOR), 
    bottom=Side(style='thin', color=CLR_CARD_BOR)
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
        res = mp + (flete if flete is not None else 0)
        return round(res, 2)
    
    if pd.isna(pv):
        return None
        
    try:
        cond = str(int(cond_exp)).strip() if pd.notna(cond_exp) else ''
    except:
        cond = str(cond_exp).strip() if pd.notna(cond_exp) else ''
    
    if cond == '1':
        res = pv * mp
    else:
        res = (mp + (flete if flete is not None else 0)) * pv
    return round(res, 2)

def calc_validacion2(importe_costo, validacion1, sociedad=None):
    """
    Desviación Real vs Teórico: Costo_TRAOPE - Validación_1.
    - Exclusivo para Sociedad 7100 (Filiales / Intercompañía), donde las transferencias son al costo.
    - Para Sociedades 7180 y 7277 (Trading / Terceros), se devuelve None (N/A) ya que el margen comercial
      se audita mediante la columna Margen Material (MOP %) y la Gobernanza de Autorizaciones.
    """
    soc = str(sociedad).strip() if sociedad is not None else ""
    if soc in ('7180', '7277') or (soc != '7100' and soc != ""):
        return None
        
    if pd.isna(importe_costo) or pd.isna(validacion1):
        return None
    diff = round(importe_costo - validacion1, 2)
    if abs(diff) < 0.05:
        return 0.0
    return diff

def calc_mop(importe_mp, um_venta, mp_compra, um_costo, pv):
    """
    Margen Operativo de Material (MOP %):
    MOP = (Precio Venta MP - MP Compra Puro Homologado) / Precio Venta MP
    Fórmula oficial:
    =@SI.CONJUNTO(
        [UM COMP]="M3", SI([MP COMP]=0, 0, ([MP VTA]-[MP COMP])/[MP VTA]),
        [UM COMP]="TN", SI([MP COMP]=0, 0, ([MP VTA]-[MP COMP])/[MP VTA])
    )
    """
    if pd.isna(importe_mp) or importe_mp <= 0:
        return None
    if pd.isna(mp_compra) or mp_compra <= 0:
        return None
        
    um_v = str(um_venta).strip().upper() if pd.notna(um_venta) else ""
    um_c = str(um_costo).strip().upper() if pd.notna(um_costo) else ""
    
    costo_homo = mp_compra
    if um_v == um_c:
        costo_homo = mp_compra
    elif um_v == 'TN' and um_c == 'M3':
        if pd.notna(pv) and pv > 0:
            costo_homo = mp_compra / pv
        else:
            return None
    elif um_v == 'M3' and um_c == 'TN':
        if pd.notna(pv) and pv > 0:
            costo_homo = mp_compra * pv
        else:
            return None
            
    mop_val = (importe_mp - costo_homo) / importe_mp
    mop_round = round(mop_val, 4)
    if abs(mop_round) < 0.0005:
        return 0.0
    return mop_round

def es_proximo_a_vencer(fecha_val, dias_margen=60):
    """
    Verifica si una fecha de fin de vigencia está próxima a vencer (+- 60 días / 2 meses desde hoy).
    """
    if pd.isna(fecha_val) or not fecha_val:
        return False
    fecha_limpia = str(fecha_val).strip().split('.')[0].replace('-', '').replace('/', '')
    if len(fecha_limpia) == 8 and fecha_limpia.isdigit():
        try:
            yyyy = int(fecha_limpia[:4])
            mm = int(fecha_limpia[4:6])
            dd = int(fecha_limpia[6:8])
            f_dt = datetime(yyyy, mm, dd)
            diff = (f_dt - datetime.now()).days
            return -dias_margen <= diff <= dias_margen
        except Exception:
            pass
    return False

def eval_autorizacion(sociedad, tipo_operacion, precio_venta, precio_referencia, mop):
    """
    Niveles de Autorización y Alertas de Gobernanza:
    
    1. SOCIEDAD 7100 (Filial / Intercompañía):
       - No Aplica (Filial)
       
    2. TRADING / SOCIEDADES 7180 y 7277 (Terceros):
       - Si Margen_Material > 8% -> 'Autoriza: Champion'
       - Si Margen_Material entre 5% y 8% -> 'Autoriza: Regional'
       - Si Margen_Material < 5% -> 'Alerta Fuera de Rango: Requiere revision puntual con Nacional'
       - Si MOP es None -> 'Pendiente de Costo TRAOPE'
    """
    soc = str(sociedad).strip()
    if soc == '7100':
        return "No Aplica (Filial)"
        
    if tipo_operacion == 'TRADING' or soc in ('7180', '7277'):
        if mop is None or pd.isna(mop):
            return "Pendiente de Costo TRAOPE"
        if mop > 0.08:
            return "Autoriza: Champion"
        elif mop >= 0.05:
            return "Autoriza: Regional"
        else:
            return "Alerta Fuera de Rango: Requiere revision puntual con Nacional"
            
    return "No Aplica (Filial)"

def eval_semaforo(row):
    """
    Semáforo de inconsistencias operativas y cumplimiento de negocio:
    - SIN_MP: falta precio de venta de material
    - SIN_FLETE: en modalidad entregada (01/04), no se encontró flete
    - SIN_COSTO: en trading, no se encontró orden de compra
    - SIN_PV: falta factor de peso volumétrico
    - DIFERENCIA: desviación mayor a $1 entre costo real y teórico (exclusivo para 7100 filiales al costo)
    - ALERTA_MARGEN: en trading (7180/7277), margen menor al 5% (requiere revisión Nacional)
    - DISCREPANCIA_ORG: condición inválida por sociedad
    - OK: registro correcto y alineado
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
    sociedad = str(row.get('Sociedad', '')).strip()
    costo = row.get('Costo Total TRAOPE') if 'Costo Total TRAOPE' in row else row.get('MP Compra (Costo Material)')
    if costo is None or pd.isna(costo):
        costo = row.get('Importe Costo')
    if (tipo_op == 'TRADING' or sociedad in ('7180', '7277')) and (pd.isna(costo) or costo is None or costo <= 0):
        return 'SIN_COSTO'
        
    if pd.isna(row.get('Validacion 1')):
        um_v = str(row.get('UM Venta', '')).strip().upper()
        um_c = str(row.get('UM Costo', '')).strip().upper()
        if not (um_v == 'TN' and um_c == 'TN'):
            if pd.isna(row.get('PV')):
                return 'SIN_PV'
                
    # Para Sociedad 7100 (Filial al costo): auditar si hay desviación contable
    if sociedad == '7100':
        v2 = row.get('Validacion 2')
        if pd.notna(v2) and abs(v2) > 1:
            return 'DIFERENCIA'
    else:
        # Para Trading (7180/7277): auditar si el margen cae por debajo del 5%
        mop = row.get('Margen Material (MOP %)')
        if pd.notna(mop) and mop < 0.05:
            return 'ALERTA_MARGEN'
        
    org = sociedad
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
# PROCESAMIENTO Y CRUCE UNIFICADO MULTI-CEDIS
# =============================================================================

def procesar_datos(data, cedis=None, modo_a=True, filtro_traope='2024'):
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
            df_traope['Validez a_str'] = df_traope['Validez a'].astype(str).str.split('.').str[0].str.replace('-', '').str.replace('.', '')
            df_traope['Validez a_num'] = pd.to_numeric(df_traope['Validez a_str'], errors='coerce')
            
            # Opción configurable de vigencia de contratos TRAOPE
            if str(filtro_traope).strip() == '2029':
                # Modo estricto legado (>= 20291231)
                df_traope = df_traope[df_traope['Validez a_num'] >= 20291231]
            elif str(filtro_traope).strip() == 'todos':
                # Sin filtro de fecha
                pass
            else:
                # Contratos (Vigencia >= 2024) [Recomendado]
                df_traope = df_traope[(df_traope['Validez a_num'] >= 20240101) | (df_traope['Validez a_num'].isna())]
                
            df_traope = df_traope.sort_values(by='Validez a_num', ascending=False)
            
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
            
        # Filtro Multi-CEDIS flexible
        if cedis:
            if isinstance(cedis, (list, tuple)):
                cedis_list = [str(c).strip().upper() for c in cedis if str(c).strip().upper() not in ('TODOS', '')]
            elif isinstance(cedis, str):
                cedis_list = [str(c).strip().upper() for c in cedis.replace(',', ' ').split() if str(c).strip().upper() not in ('TODOS', '')]
            else:
                cedis_list = [str(cedis).strip().upper()]
                
            if cedis_list:
                df_mp = df_mp[df_mp['Centro'].isin(cedis_list)]
                logging.info(f"Filtrando por CEDIS seleccionados ({len(cedis_list)}): {cedis_list}")
            
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
        
        # Catálogos Maestros Globales de Snowflake para Nombres de SF y Descripción de Centros
        col_sf_traope = 'Ship From' if 'Ship From' in df_traope.columns else 'Shipfrom'
        col_nom_sf = 'Nombre SF' if 'Nombre SF' in df_traope.columns else None
        col_desc_centro = 'Descripción Centro' if 'Descripción Centro' in df_traope.columns else 'Desc. Centro'
        
        map_shipfrom_nombre = {}
        if col_nom_sf and col_nom_sf in df_traope.columns:
            map_shipfrom_nombre = (
                df_traope.dropna(subset=[col_sf_traope, col_nom_sf])
                .drop_duplicates(col_sf_traope)
                .set_index(col_sf_traope)[col_nom_sf]
                .to_dict()
            )
            
        map_centro_desc = {}
        if col_desc_centro and col_desc_centro in df_traope.columns:
            map_centro_desc = (
                df_traope.dropna(subset=['Centro', col_desc_centro])
                .drop_duplicates('Centro')
                .set_index('Centro')[col_desc_centro]
                .to_dict()
            )

        # Identificar columnas de Condición de Expedición en cada DataFrame
        col_f_exp = next((c for c in df_flete.columns if 'exp' in str(c).lower()), None)
        col_v_exp = next((c for c in df_contratos.columns if 'exp' in str(c).lower()), None)
        col_t_exp = next((c for c in df_traope.columns if 'cond. exp' in str(c).lower()), None)

        # Normalizar llaves en Contratos de Venta para búsqueda en cascada
        if 'Shipfrom' in df_contratos.columns and 'Centro' in df_contratos.columns and 'Material' in df_contratos.columns:
            df_contratos['Shipfrom_clean'] = df_contratos['Shipfrom'].apply(str_val)
            df_contratos['Centro_clean'] = df_contratos['Centro'].apply(str_val)
            df_contratos['Material_clean'] = df_contratos['Material'].apply(str_val)
            col_dest_v = 'Destino' if 'Destino' in df_contratos.columns else None
            if col_dest_v:
                df_contratos['Destino_clean'] = df_contratos[col_dest_v].apply(str_val)
                df_contratos['Concat1'] = df_contratos['Shipfrom_clean'] + '-' + df_contratos['Centro_clean'] + '-' + df_contratos['Destino_clean'] + '-' + df_contratos['Material_clean']
            else:
                df_contratos['Concat1'] = 'NA'
            df_contratos['Concat2'] = df_contratos['Shipfrom_clean'] + '-' + df_contratos['Centro_clean'] + '-' + df_contratos['Material_clean']

        # Mapeos globales para resolución en cascada de Condición de Expedición
        f_map_c1 = df_flete.dropna(subset=[col_f_exp]).drop_duplicates('Concat1').set_index('Concat1')[col_f_exp].to_dict() if col_f_exp else {}
        f_map_c2 = df_flete.dropna(subset=[col_f_exp]).drop_duplicates('Concat2').set_index('Concat2')[col_f_exp].to_dict() if col_f_exp else {}

        v_map_c1 = df_contratos.dropna(subset=[col_v_exp]).drop_duplicates('Concat1').set_index('Concat1')[col_v_exp].to_dict() if (col_v_exp and 'Concat1' in df_contratos.columns) else {}
        v_map_c2 = df_contratos.dropna(subset=[col_v_exp]).drop_duplicates('Concat2').set_index('Concat2')[col_v_exp].to_dict() if (col_v_exp and 'Concat2' in df_contratos.columns) else {}

        t_map_c1 = df_traope.dropna(subset=[col_t_exp]).drop_duplicates('Concat1').set_index('Concat1')[col_t_exp].to_dict() if (col_t_exp and 'Concat1' in df_traope.columns) else {}
        t_map_c2 = df_traope.dropna(subset=[col_t_exp]).drop_duplicates('Concat2').set_index('Concat2')[col_t_exp].to_dict() if (col_t_exp and 'Concat2' in df_traope.columns) else {}

        # Catálogo de Descripciones de Material de respaldo (TRAOPE y Contratos)
        col_desc_traope = next((c for c in df_traope.columns if 'descrip' in str(c).lower()), None)
        map_mat_desc_traope = {}
        if col_desc_traope:
            map_mat_desc_traope = (
                df_traope.dropna(subset=['Material', col_desc_traope])
                .drop_duplicates('Material')
                .set_index('Material')[col_desc_traope]
                .to_dict()
            )
            
        col_desc_vta = next((c for c in df_contratos.columns if 'desc' in str(c).lower()), None)
        map_mat_desc_vta = {}
        if col_desc_vta and 'Material_clean' in df_contratos.columns:
            map_mat_desc_vta = (
                df_contratos.dropna(subset=['Material_clean', col_desc_vta])
                .drop_duplicates('Material_clean')
                .set_index('Material_clean')[col_desc_vta]
                .to_dict()
            )

        # =====================================================================
        # CONSTRUCCIÓN DE LA MATRIZ FILA POR FILA
        # =====================================================================
        filas = []
        for _, row in df_mp.iterrows():
            c1 = row['Concat1']
            c2 = row['Concat2']
            mat = row['Material']
            sf_str = str(row.get('Shipfrom', '')).strip()
            centro_str = str(row.get('Centro', '')).strip()
            
            # 1. Flete (match exacto)
            f_row = flete_c1.loc[c1] if c1 in flete_c1.index else None
            
            # 2. PV y Denominación (con respaldo de TRAOPE y Contratos si falta en PESO_VOL)
            pv_val = pv_idx.loc[mat, 'PV_calc'] if mat in pv_idx.index else None
            pv_desc = pv_idx.loc[mat, 'Texto de material'] if mat in pv_idx.index else ''
            if not pv_desc:
                pv_desc = map_mat_desc_traope.get(mat, '') or map_mat_desc_vta.get(mat, '')
            
            # 3. Contratos de Venta
            llave_co = f"{row['Centro']}-{row['Shipfrom']}-{row['Destinatario']}-{mat}"
            co_row = contratos_idx.loc[llave_co] if llave_co in contratos_idx.index else None
            
            # 4. TRAOPE (Costo de Compra)
            t_row = None
            if c1 in traope_c1.index:
                t_row = traope_c1.loc[c1]
            elif c2 in traope_c2.index:
                t_row = traope_c2.loc[c2]
                
            # Condición de Expedición con Búsqueda en Cascada Multifuente
            cond_exp = ''
            if f_row is not None and col_f_exp and col_f_exp in f_row.index and pd.notna(f_row[col_f_exp]):
                cond_exp = str(f_row[col_f_exp]).strip()
            elif co_row is not None and col_v_exp and col_v_exp in co_row.index and pd.notna(co_row[col_v_exp]):
                cond_exp = str(co_row[col_v_exp]).strip()
            elif t_row is not None and col_t_exp and col_t_exp in t_row.index and pd.notna(t_row[col_t_exp]):
                cond_exp = str(t_row[col_t_exp]).strip()
            elif c1 in f_map_c1:
                cond_exp = str(f_map_c1[c1]).strip()
            elif c1 in v_map_c1:
                cond_exp = str(v_map_c1[c1]).strip()
            elif c1 in t_map_c1:
                cond_exp = str(t_map_c1[c1]).strip()
            elif c2 in f_map_c2:
                cond_exp = str(f_map_c2[c2]).strip()
            elif c2 in v_map_c2:
                cond_exp = str(v_map_c2[c2]).strip()
            elif c2 in t_map_c2:
                cond_exp = str(t_map_c2[c2]).strip()
                
            # Nombres de Ship From y Descripción de Centro con catálogo maestro global
            nombre_sf = ''
            if t_row is not None and 'Nombre SF' in t_row.index and pd.notna(t_row['Nombre SF']):
                nombre_sf = str(t_row['Nombre SF']).strip()
            if not nombre_sf:
                nombre_sf = map_shipfrom_nombre.get(sf_str, '')

            desc_centro = ''
            if t_row is not None and 'Descripción Centro' in t_row.index and pd.notna(t_row['Descripción Centro']):
                desc_centro = str(t_row['Descripción Centro']).strip()
            if not desc_centro:
                desc_centro = map_centro_desc.get(centro_str, '')
                
            imp_mp = row['Importe']
            imp_flete = f_row['Importe'] if f_row is not None else None
            um_venta = row.get('Unidad', '')
            
            imp_costo_total = None
            imp_flete_compra = None
            mp_compra = None
            um_costo = None
            contrato_compra = ''
            fin_vigencia_compra = ''
            
            if t_row is not None:
                if 'Contrato_Pos' in t_row.index and pd.notna(t_row['Contrato_Pos']):
                    contrato_compra = str(t_row['Contrato_Pos']).strip()
                elif 'Doc. Compras' in t_row.index and pd.notna(t_row['Doc. Compras']):
                    pos_str = str(t_row.get('Pos', '')).strip()
                    contrato_compra = f"{str(t_row['Doc. Compras']).strip()}-{pos_str}" if pos_str else str(t_row['Doc. Compras']).strip()
                    
                if 'Validez a' in t_row.index and pd.notna(t_row['Validez a']):
                    fin_vigencia_compra = str(t_row['Validez a']).strip().split('.')[0]
                elif 'Fin período validez' in t_row.index and pd.notna(t_row['Fin período validez']):
                    fin_vigencia_compra = str(t_row['Fin período validez']).strip().split('.')[0]

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
                        
            # Clasificación Oficial Tipo de Operación:
            # - 7100 = CANTERAS PROPIAS (100% de los casos)
            # - 7180 y 7277 = TRADING (Terceros)
            # - Otras sociedades: si tiene costo o prefijo TP/TC es TRADING, si no CANTERAS PROPIAS
            sociedad_str = str(row.get('Org. Ventas', '')).strip()
            nombre_sf_upper = str(t_row['Nombre SF'] if t_row is not None and 'Nombre SF' in t_row.index else '').upper()
            
            if sociedad_str == '7100':
                tipo_operacion = 'CANTERAS PROPIAS'
            elif sociedad_str in ('7180', '7277'):
                tipo_operacion = 'TRADING'
            elif (imp_costo_total is not None and imp_costo_total > 0) or 'TP-' in nombre_sf_upper or 'SF TP' in nombre_sf_upper or 'TC-' in nombre_sf_upper:
                tipo_operacion = 'TRADING'
            else:
                tipo_operacion = 'CANTERAS PROPIAS'
                
            # Cálculos de Márgenes y Gobernanza (MOP exclusivamente sobre Material)
            # En Sociedad 7100 (Filial / Canteras Propias), el MOP es siempre 0.0%
            if sociedad_str == '7100':
                mop = 0.0
            else:
                mop = calc_mop(imp_mp, um_venta, mp_compra, um_costo, pv_val)
                
            precio_ref = precios_ref_mat.get(mat, imp_mp)
            nivel_aut = eval_autorizacion(sociedad_str, tipo_operacion, imp_mp, precio_ref, mop)
            
            # Validaciones Integrales
            val1 = calc_validacion1(um_venta, um_costo, cond_exp, imp_mp, imp_flete, pv_val)
            val2 = calc_validacion2(imp_costo_total, val1, sociedad_str)
            
            # Armado de fila con precio junto a cada condición
            fila_dict = {
                'Concat1': c1,
                'Concat2': c2,
                'Sociedad': row.get('Org. Ventas', ''),
                'Ship From': row.get('Shipfrom', ''),
                'Nombre SF': nombre_sf,
                'Centro': row.get('Centro', ''),
                'Desc. Centro': desc_centro,
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
                'No. Contrato Compra': contrato_compra,
                'Fin Vigencia Compra': fin_vigencia_compra,
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
# CONSTRUCCIÓN DEL DASHBOARD EJECUTIVO PREMIUM (ESTILO CEMEX UI)
# =============================================================================

def apply_header_style(cell, text, fill, font):
    cell.value = text
    cell.fill = fill
    cell.font = font
    cell.alignment = ALIGN_CENTER
    cell.border = THIN_BORDER

def dibujar_tarjeta_kpi(ws, col_ini, col_fin, titulo, valor_str, subtitulo, accent_fill, val_font=None):
    """Dibuja una tarjeta KPI corporativa con barra de acento superior y sombra sutil."""
    # Fila 5: Barra de Acento Superior (3.5 pt)
    for c in range(col_ini, col_fin + 1):
        cell_acc = ws.cell(row=5, column=c)
        cell_acc.fill = accent_fill
        cell_acc.border = CARD_BORDER
    ws.merge_cells(start_row=5, start_column=col_ini, end_row=5, end_column=col_fin)
    
    # Fila 6: Título
    for c in range(col_ini, col_fin + 1):
        cell_t = ws.cell(row=6, column=c)
        cell_t.fill = FILL_CARD_BG
        cell_t.border = CARD_BORDER
    ws.merge_cells(start_row=6, start_column=col_ini, end_row=6, end_column=col_fin)
    c_tit = ws.cell(row=6, column=col_ini, value=titulo)
    c_tit.font = FONT_KPI_TITLE
    c_tit.alignment = ALIGN_CENTER
    
    # Fila 7: Valor Principal (Grande)
    for c in range(col_ini, col_fin + 1):
        cell_v = ws.cell(row=7, column=c)
        cell_v.fill = FILL_CARD_BG
        cell_v.border = CARD_BORDER
    ws.merge_cells(start_row=7, start_column=col_ini, end_row=7, end_column=col_fin)
    c_val = ws.cell(row=7, column=col_ini, value=valor_str)
    c_val.font = val_font or FONT_KPI_VAL
    c_val.alignment = ALIGN_CENTER
    
    # Fila 8: Subtítulo
    for c in range(col_ini, col_fin + 1):
        cell_s = ws.cell(row=8, column=c)
        cell_s.fill = FILL_CARD_BG
        cell_s.border = CARD_BORDER
    ws.merge_cells(start_row=8, start_column=col_ini, end_row=8, end_column=col_fin)
    c_sub = ws.cell(row=8, column=col_ini, value=subtitulo)
    c_sub.font = FONT_KPI_SUB
    c_sub.alignment = ALIGN_CENTER

def construir_dashboard_ejecutivo(wb, df_matriz, cedis_str, fecha_str, filtro_traope='2024'):
    ws_dash = wb.create_sheet(title="Dashboard", index=0)
    ws_dash.views.sheetView[0].showGridLines = True
    
    # Anchos de columna estilo App Web
    ws_dash.column_dimensions['A'].width = 3.5    # Sidebar
    for c in ['B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M']:
        ws_dash.column_dimensions[c].width = 16.5
        
    # Alturas de fila
    ws_dash.row_dimensions[1].height = 6
    ws_dash.row_dimensions[2].height = 24
    ws_dash.row_dimensions[3].height = 18
    ws_dash.row_dimensions[4].height = 10
    ws_dash.row_dimensions[5].height = 4.5  # Acento KPI
    ws_dash.row_dimensions[6].height = 16   # Título KPI
    ws_dash.row_dimensions[7].height = 28   # Valor KPI
    ws_dash.row_dimensions[8].height = 16   # Subtítulo KPI
    ws_dash.row_dimensions[9].height = 12   # Separador
    ws_dash.row_dimensions[10].height = 22  # Header Secciones
    ws_dash.row_dimensions[18].height = 12  # Separador
    ws_dash.row_dimensions[19].height = 22  # Header Auditoría
    
    # --- 1. SIDEBAR AZUL CEMEX ---
    for r in range(1, 60):
        c_side = ws_dash.cell(row=r, column=1)
        c_side.fill = FILL_NAVY
        
    # --- 2. HERO BANNER ---
    for c in range(2, 14):
        for r in (1, 2, 3):
            ws_dash.cell(row=r, column=c).fill = FILL_NAVY
    ws_dash.merge_cells("B2:M2")
    ws_dash.merge_cells("B3:M3")
    
    c_hero_t = ws_dash.cell(row=2, column=2, value="CEMEX AGREGADOS  |  Matriz de Precios Venta vs Costo & Gobernanza")
    c_hero_t.font = FONT_WHITE_HERO
    c_hero_t.alignment = ALIGN_LEFT
    
    tag_ft_desc = "Contratos (Vigencia >= 2029)" if str(filtro_traope).strip() == '2029' else "Contratos (Vigencia >= 2024)"
    sub_txt = f"  CEDIS: {cedis_str}   |   TRAOPE: {tag_ft_desc}   |   Fuente: Snowflake 2026   |   {fecha_str}"
    c_hero_s = ws_dash.cell(row=3, column=2, value=sub_txt)
    c_hero_s.font = FONT_WHITE_SUB
    c_hero_s.alignment = ALIGN_LEFT
    
    # --- 3. MÉTRICAS CLAVE ---
    totales = len(df_matriz)
    df_trading = df_matriz[df_matriz['Sociedad'].astype(str).str.strip() != '7100']
    tot_trading = len(df_trading)
    tot_filial = totales - tot_trading
    
    mop_validos = df_trading['Margen Material (MOP %)'].dropna() if not df_trading.empty else df_matriz['Margen Material (MOP %)'].dropna()
    mop_prom = mop_validos.mean() if not mop_validos.empty else 0
    
    auth_champion = len(df_trading[df_trading['Nivel Autorización / Alerta'].str.contains('Champion', na=False)])
    auth_regional = len(df_trading[df_trading['Nivel Autorización / Alerta'].str.contains('Regional', na=False)])
    auth_nacional = len(df_trading[df_trading['Nivel Autorización / Alerta'].str.contains('Nacional|Alerta', na=False)])
    
    pct_champ = (auth_champion / tot_trading * 100) if tot_trading > 0 else 0
    pct_reg = (auth_regional / tot_trading * 100) if tot_trading > 0 else 0
    pct_nac = (auth_nacional / tot_trading * 100) if tot_trading > 0 else 0
    
    # 4 TARJETAS KPI MODERNAS
    dibujar_tarjeta_kpi(ws_dash, 2, 4, "TOTAL RUTAS ANALIZADAS", f"{totales:,}", f"Trading: {tot_trading:,} | Filial: {tot_filial:,}", FILL_NAVY)
    dibujar_tarjeta_kpi(ws_dash, 5, 7, "MARGEN MATERIAL (TRADING)", f"{mop_prom*100:.1f}%", "Ponderado sobre material puro", FILL_GREEN, Font(color=CLR_DARK_GREEN, bold=True, size=16))
    dibujar_tarjeta_kpi(ws_dash, 8, 10, "AUTORIZACIÓN CHAMPION", f"{auth_champion} ({pct_champ:.1f}%)", "Margen > 8% (Trading)", FILL_BLUE_ACC, Font(color=CLR_BLUE_ACCENT, bold=True, size=16))
    dibujar_tarjeta_kpi(ws_dash, 11, 13, "ALERTAS NIVEL NACIONAL", f"{auth_nacional} ({pct_nac:.1f}%)", "Margen < 5% (Trading)", FILL_RED, Font(color=CLR_DARK_RED, bold=True, size=16))
    
    # --- 4. SECCIONES ANALÍTICAS LADO A LADO ---
    
    # BLOQUE IZQUIERDO: Mix de Gobernanza
    ws_dash.merge_cells("B10:F10")
    apply_header_style(ws_dash.cell(row=10, column=2), "ESTATUS DE GOBERNANZA Y AUTORIZACIÓN", FILL_NAVY, FONT_WHITE_BOLD)
    
    headers_gob = [("Nivel Autorización", 2), ("Criterio Objetivo", 3), ("Rutas", 5), ("% Trading", 6)]
    ws_dash.merge_cells("C11:D11")
    for txt, col in headers_gob:
        c = ws_dash.cell(row=11, column=col, value=txt)
        c.fill = FILL_HEADER
        c.font = FONT_BLACK_BOLD
        c.alignment = ALIGN_CENTER
        c.border = THIN_BORDER
        
    filas_gob = [
        ("Nivel Champion", "Margen > 8% (Trading)", auth_champion, pct_champ, FILL_GREEN, FONT_BLACK_BOLD),
        ("Nivel Regional", "Margen 5% - 8% (Trading)", auth_regional, pct_reg, FILL_REGIONAL, FONT_NORMAL),
        ("Alerta Nacional", "Margen < 5% (Trading)", auth_nacional, pct_nac, FILL_RED, FONT_BLACK_BOLD),
        ("No Aplica (Filial)", "Transferencia 0% (Intercompañía)", tot_filial, (tot_filial/totales*100) if totales>0 else 0, FILL_WHITE, FONT_NORMAL)
    ]
    
    for idx, (niv, crit, cant, pct, fill, f_font) in enumerate(filas_gob, start=12):
        ws_dash.merge_cells(start_row=idx, start_column=3, end_row=idx, end_column=4)
        c1 = ws_dash.cell(row=idx, column=2, value=niv)
        c2 = ws_dash.cell(row=idx, column=3, value=crit)
        c3 = ws_dash.cell(row=idx, column=5, value=cant)
        c4 = ws_dash.cell(row=idx, column=6, value=f"{pct:.1f}%")
        
        for c_col in [c1, c2, c3, c4]:
            c_col.fill = fill
            c_col.font = f_font
            c_col.border = THIN_BORDER
        c1.alignment = ALIGN_LEFT
        c2.alignment = ALIGN_LEFT
        c3.alignment = ALIGN_CENTER
        c4.alignment = ALIGN_CENTER
        
    # BLOQUE DERECHO: Top 5 Materiales por Margen
    ws_dash.merge_cells("H10:M10")
    apply_header_style(ws_dash.cell(row=10, column=8), "TOP 5 MATERIALES CON MAYOR MARGEN MATERIAL (MOP %)", FILL_NAVY, FONT_WHITE_BOLD)
    
    headers_mat = [("Material", 8), ("Denominación", 9), ("Precio Vta Prom.", 11), ("MP Compra Prom.", 12), ("MOP %", 13)]
    ws_dash.merge_cells("I11:J11")
    for txt, col in headers_mat:
        c = ws_dash.cell(row=11, column=col, value=txt)
        c.fill = FILL_HEADER
        c.font = FONT_BLACK_BOLD
        c.alignment = ALIGN_CENTER
        c.border = THIN_BORDER
        
    # Top 5 materiales con mayor margen (excluyendo filiales 7100 donde MOP es 0%)
    df_top_candidates = df_matriz[df_matriz['Margen Material (MOP %)'].notna() & (df_matriz['Sociedad'].astype(str).str.strip() != '7100')]
    if df_top_candidates.empty:
        df_top_candidates = df_matriz[df_matriz['Margen Material (MOP %)'].notna()]
        
    df_top_mat = df_top_candidates.groupby(['Material', 'Denominación']).agg({
        'Importe MP': 'mean',
        'MP Compra (Costo Material)': 'mean',
        'Margen Material (MOP %)': 'mean'
    }).reset_index().dropna().sort_values(by='Margen Material (MOP %)', ascending=False).head(5)
    
    for idx, (_, row_m) in enumerate(df_top_mat.iterrows(), start=12):
        ws_dash.merge_cells(start_row=idx, start_column=9, end_row=idx, end_column=10)
        c_mat = ws_dash.cell(row=idx, column=8, value=row_m['Material'])
        c_den = ws_dash.cell(row=idx, column=9, value=row_m['Denominación'])
        c_vta = ws_dash.cell(row=idx, column=11, value=row_m['Importe MP'])
        c_cmp = ws_dash.cell(row=idx, column=12, value=row_m['MP Compra (Costo Material)'])
        c_mop = ws_dash.cell(row=idx, column=13, value=row_m['Margen Material (MOP %)'])
        
        for cell_item in [c_mat, c_den, c_vta, c_cmp, c_mop]:
            cell_item.border = THIN_BORDER
            cell_item.font = FONT_NORMAL
            
        c_mat.alignment = ALIGN_CENTER
        c_den.alignment = ALIGN_LEFT
        c_vta.number_format = '$#,##0.00'
        c_vta.alignment = ALIGN_RIGHT
        c_cmp.number_format = '$#,##0.00'
        c_cmp.alignment = ALIGN_RIGHT
        c_mop.number_format = '0.0%'
        c_mop.alignment = ALIGN_RIGHT
        mop_val_top = row_m['Margen Material (MOP %)']
        c_mop.fill = FILL_GREEN if pd.notna(mop_val_top) and mop_val_top > 0.08 else FILL_REGIONAL
        
    # --- 5. TABLA DE AUDITORÍA: TOP RUTAS CON ALERTA O DESVIACIÓN ---
    ws_dash.merge_cells("B19:M19")
    apply_header_style(ws_dash.cell(row=19, column=2), "AUDITORÍA DE RUTAS: DESVIACIONES DE MARGEN Y ALERTAS NACIONALES (TOP 50)", FILL_NAVY, FONT_WHITE_BOLD)
    
    headers_aud = [
        ("Concat1 (Ruta)", 2), ("CEDIS", 3), ("Destino", 4), ("Material", 5),
        ("Denominación", 6), ("Precio Vta", 8), ("MP Compra", 9), ("MOP %", 10),
        ("Nivel Autorización", 11), ("Semáforo", 12), ("Diferencia ($)", 13)
    ]
    ws_dash.merge_cells("F20:G20")
    for txt, col in headers_aud:
        c = ws_dash.cell(row=20, column=col, value=txt)
        c.fill = FILL_HEADER
        c.font = FONT_BLACK_BOLD
        c.alignment = ALIGN_CENTER
        c.border = THIN_BORDER
        
    anomalias = df_matriz[(df_matriz['Semaforo'] != 'OK') | (df_matriz['Nivel Autorización / Alerta'].str.contains('Alerta|Nacional', na=False))].head(50)
    
    for r_idx, (_, row) in enumerate(anomalias.iterrows(), start=21):
        ws_dash.merge_cells(start_row=r_idx, start_column=6, end_row=r_idx, end_column=7)
        c1 = ws_dash.cell(row=r_idx, column=2, value=row.get('Concat1', ''))
        c2 = ws_dash.cell(row=r_idx, column=3, value=row.get('Centro', ''))
        c3 = ws_dash.cell(row=r_idx, column=4, value=row.get('Destino', ''))
        c4 = ws_dash.cell(row=r_idx, column=5, value=row.get('Material', ''))
        c5 = ws_dash.cell(row=r_idx, column=6, value=row.get('Denominación', ''))
        
        c6 = ws_dash.cell(row=r_idx, column=8, value=row.get('Importe MP'))
        c7 = ws_dash.cell(row=r_idx, column=9, value=row.get('MP Compra (Costo Material)'))
        c8 = ws_dash.cell(row=r_idx, column=10, value=row.get('Margen Material (MOP %)'))
        c9 = ws_dash.cell(row=r_idx, column=11, value=row.get('Nivel Autorización / Alerta', ''))
        c10 = ws_dash.cell(row=r_idx, column=12, value=row.get('Semaforo', ''))
        c11 = ws_dash.cell(row=r_idx, column=13, value=row.get('Validacion 2'))
        
        for c_item in [c1, c2, c3, c4, c5, c6, c7, c8, c9, c10, c11]:
            c_item.border = THIN_BORDER
            c_item.font = FONT_NORMAL
            
        c1.alignment = ALIGN_LEFT
        c2.alignment = ALIGN_CENTER
        c3.alignment = ALIGN_CENTER
        c4.alignment = ALIGN_CENTER
        c5.alignment = ALIGN_LEFT
        
        c6.number_format = '$#,##0.00'
        c6.alignment = ALIGN_RIGHT
        c7.number_format = '$#,##0.00'
        c7.alignment = ALIGN_RIGHT
        
        if isinstance(c8.value, (int, float)):
            c8.number_format = '0.0%'
            c8.alignment = ALIGN_RIGHT
            if c8.value < 0.05:
                c8.fill = FILL_RED
            elif c8.value > 0.08:
                c8.fill = FILL_GREEN
                
        c9.alignment = ALIGN_CENTER
        if 'Nacional' in str(c9.value) or 'Alerta' in str(c9.value):
            c9.fill = FILL_RED
            c9.font = FONT_BLACK_BOLD
        elif 'Regional' in str(c9.value):
            c9.fill = FILL_REGIONAL
            
        c10.alignment = ALIGN_CENTER
        if c10.value == 'DIFERENCIA':
            c10.fill = FILL_WARN
        elif c10.value in ('SIN_FLETE', 'SIN_COSTO', 'SIN_PV'):
            c10.fill = FILL_ORANGE
        elif c10.value == 'OK':
            c10.fill = FILL_GREEN
            
        if isinstance(c11.value, (int, float)):
            c11.number_format = '$#,##0.00'
            c11.alignment = ALIGN_RIGHT
            if abs(c11.value) > 1:
                c11.fill = FILL_WARN

# =============================================================================
# ESCRITURA EN EXCEL FORMATEADO
# =============================================================================

def escribir_excel(df_matriz, cedis, out_dir, raw_data_frames, filtro_traope='2024'):
    if df_matriz.empty:
        logging.error("DataFrame vacío, no se generará Excel.")
        return
        
    fecha_dt = datetime.now()
    fecha_str = fecha_dt.strftime('%Y%m%d_%H%M%S')
    fecha_legible = fecha_dt.strftime('%d/%m/%Y %H:%M')
    
    if isinstance(cedis, (list, tuple)):
        cedis_tag = "_".join(cedis[:3]) + (f"_y_{len(cedis)-3}_mas" if len(cedis)>3 else "")
        cedis_label = ", ".join(cedis)
    elif cedis:
        cedis_tag = str(cedis).replace(' ', '_').replace(',', '_')
        cedis_label = str(cedis)
    else:
        cedis_tag = "TODOS"
        cedis_label = "TODOS LOS CEDIS"
        
    tag_ft_file = "Contratos2029" if str(filtro_traope).strip() == '2029' else "Contratos2024"
    os.makedirs(out_dir, exist_ok=True)
    out_file = os.path.join(out_dir, f"Matriz_Precios_Integral_{cedis_tag}_{tag_ft_file}_{fecha_str}.xlsx")
    
    logging.info(f"Escribiendo Excel: {out_file}")
    wb = openpyxl.Workbook()
    
    # -------------------------------------------------------------------------
    # Hoja 1: Matriz de Datos Detallada
    # -------------------------------------------------------------------------
    ws_matriz = wb.active
    ws_matriz.title = "Matriz"
    
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
        ('No. Contrato Compra', FILL_COSTO, FONT_BLACK_BOLD, 18),
        ('Fin Vigencia Compra', FILL_COSTO, FONT_BLACK_BOLD, 16),
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
    ws_matriz.merge_cells(start_row=1, start_column=1, end_row=1, end_column=13)
    apply_header_style(ws_matriz.cell(row=1, column=1), "DATOS DE RUTA Y MATERIAL", FILL_NAVY, FONT_WHITE_BOLD)
    
    ws_matriz.merge_cells(start_row=1, start_column=14, end_row=1, end_column=19)
    apply_header_style(ws_matriz.cell(row=1, column=14), "CONDICIONES DE VENTA", FILL_NAVY, FONT_WHITE_BOLD)
    
    ws_matriz.merge_cells(start_row=1, start_column=20, end_row=1, end_column=26)
    apply_header_style(ws_matriz.cell(row=1, column=20), "CONDICIONES DE COMPRA (COSTO)", FILL_NAVY, FONT_WHITE_BOLD)
    
    ws_matriz.merge_cells(start_row=1, start_column=27, end_row=1, end_column=29)
    apply_header_style(ws_matriz.cell(row=1, column=27), "GOBERNANZA Y AUTORIZACIÓN", FILL_NAVY, FONT_WHITE_BOLD)
    
    ws_matriz.merge_cells(start_row=1, start_column=30, end_row=1, end_column=32)
    apply_header_style(ws_matriz.cell(row=1, column=30), "VALIDACIÓN DE MARGEN", FILL_NAVY, FONT_WHITE_BOLD)
    
    ws_matriz.merge_cells(start_row=1, start_column=33, end_row=1, end_column=35)
    apply_header_style(ws_matriz.cell(row=1, column=33), "CONTRATO DE VENTA", FILL_NAVY, FONT_WHITE_BOLD)
    
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
    
    # Orden exacto de columnas para el volcado (35 columnas)
    cols_order = [
        'Concat1', 'Concat2', 'Sociedad', 'Ship From', 'Nombre SF',
        'Centro', 'Desc. Centro', 'Destino', 'Material', 'Denominación',
        'PV', 'Inicio Vigencia', 'Fin Vigencia',
        'Clase Cond. MP', 'Importe MP', 'UM Venta', 'Clase Cond. Flete', 'Importe Flete', 'Cond. Expedición',
        'No. Contrato Compra', 'Fin Vigencia Compra', 'Costo Total TRAOPE', 'Flete Compra', 'MP Compra (Costo Material)', 'UM Costo', 'Margen Material (MOP %)',
        'Tipo Operación', 'Precio Referencia', 'Nivel Autorización / Alerta',
        'Validacion 1', 'Validacion 2', 'Semaforo',
        'No. Contrato Venta', 'UM Contrato', 'Precio Contrato'
    ]
    df_out = df_matriz[cols_order]
    
    for r_idx, row in enumerate(df_out.itertuples(index=False), start=3):
        # row mapping:
        # row[2]: Sociedad, row[20]: Fin Vigencia Compra, row[26]: Tipo Operacion
        soc_val = str(row[2]).strip()
        tipo_op_val = str(row[26]).strip()
        fin_vig_compra = row[20]
        es_expirando = es_proximo_a_vencer(fin_vig_compra) if (tipo_op_val == 'TRADING' or soc_val in ('7180', '7277')) else False
        
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
            if c_idx in [15, 18, 22, 23, 24, 28, 30, 31, 35]: 
                cell.number_format = '$#,##0.00'
                cell.alignment = ALIGN_RIGHT
            elif c_idx == 26: # Margen MOP %
                cell.number_format = '0.0%'
                cell.alignment = ALIGN_RIGHT
                if isinstance(val, (int, float)) and pd.notna(val):
                    if soc_val == '7100':
                        cell.fill = FILL_WHITE
                    elif val < 0.05:
                        cell.fill = FILL_RED
                    elif val > 0.08:
                        cell.fill = FILL_GREEN
                    else:
                        cell.fill = FILL_REGIONAL
            elif c_idx == 11: # PV
                cell.number_format = '#,##0.000'
                cell.alignment = ALIGN_RIGHT
            elif c_idx in [3, 4, 6, 8, 9, 12, 13, 14, 16, 17, 19, 20, 21, 25, 27, 32, 33, 34]:
                cell.alignment = ALIGN_CENTER
            else:
                cell.alignment = ALIGN_LEFT
                
            # Alerta Próximo a Vencer (+- 2 meses / 60 días) en Contratos de Compra Terceros (Cols 20 y 21)
            if c_idx in [20, 21] and es_expirando:
                cell.fill = FILL_WARN
                
            # Alertas de Gobernanza (Col 29)
            if c_idx == 29 and isinstance(val, str):
                if 'Champion' in val:
                    cell.fill = FILL_GREEN
                elif 'Regional' in val:
                    cell.fill = FILL_REGIONAL
                elif 'Alerta' in val or 'Nacional' in val:
                    cell.fill = FILL_RED
                    cell.font = FONT_BLACK_BOLD
                    
            # Semáforo (Col 32)
            if c_idx == 32 and isinstance(val, str):
                if val in ('SIN_FLETE', 'SIN_PV', 'SIN_COSTO', 'SIN_MP'):
                    cell.fill = FILL_ORANGE
                elif val == 'DIFERENCIA':
                    cell.fill = FILL_WARN
                elif val == 'ALERTA_MARGEN':
                    cell.fill = FILL_RED
                    cell.font = FONT_BLACK_BOLD
                elif val == 'DISCREPANCIA_ORG':
                    cell.fill = FILL_RED
                elif val == 'OK':
                    cell.fill = FILL_GREEN
                    
            # Validación 2 amarilla si abs > 1 (Col 31) (Exclusivo para Filiales 7100)
            if c_idx == 31 and isinstance(val, (int, float)) and pd.notna(val):
                if abs(val) > 1 and soc_val == '7100':
                    cell.fill = FILL_WARN
                    
    # -------------------------------------------------------------------------
    # Hoja 2: Dashboard Ejecutivo Premium (Hoja Inicial)
    # -------------------------------------------------------------------------
    construir_dashboard_ejecutivo(wb, df_matriz, cedis_label, fecha_legible, filtro_traope=filtro_traope)
    
    # -------------------------------------------------------------------------
    # Hojas Ocultas de Respaldo para Auditoría
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
    return out_file

# =============================================================================
# CLI Y EJECUCIÓN
# =============================================================================

def main():
    parser = argparse.ArgumentParser(description="Orquestador Unificado CEMEX Matriz Ventas y Gobernanza")
    parser.add_argument('--fuente', type=str, help="Ruta al archivo Excel .xlsm (Modo A)")
    parser.add_argument('--mp', type=str, help="Ruta al archivo TXT de Material (Modo B)")
    parser.add_argument('--flete', type=str, help="Ruta al archivo TXT de Flete (Modo B)")
    parser.add_argument('--cedis', nargs='*', default=None, help="Uno o varios centros CEDIS (ej. D836 D838 DW66 o TODOS)")
    parser.add_argument(
        '--filtro-traope',
        type=str,
        default='2024',
        choices=['2024', '2029', 'todos'],
        help="Filtro de vigencia para contratos TRAOPE: '2024' (Vigencia >= 2024, recomendado), '2029' (Vigencia >= 2029), 'todos' (sin filtro)"
    )
    script_dir = os.path.dirname(os.path.abspath(__file__))
    default_out = os.path.join(script_dir, "_salidas_integradas")
    parser.add_argument('--output', type=str, default=default_out, help="Carpeta de salida")
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
        df_matriz, df_mp, df_flete, df_traope, df_contratos = procesar_datos(
            data_raw, cedis=args.cedis, modo_a=True, filtro_traope=args.filtro_traope
        )
        escribir_excel(
            df_matriz, args.cedis, args.output, [df_mp, df_flete, df_traope, df_contratos], filtro_traope=args.filtro_traope
        )
        
    elif args.mp and args.flete:
        data_raw = {
            'MP': parse_txt(args.mp),
            'Flete': parse_txt(args.flete)
        }
        df_matriz, _, _, _, _ = procesar_datos(data_raw, cedis=args.cedis, modo_a=False, filtro_traope=args.filtro_traope)
        logging.info("Modo B ejecutado exitosamente.")
    else:
        logging.error("Debe proveer --fuente (Modo A) o --mp y --flete (Modo B)")
        
    logging.info(f"Listo en {time.time() - inicio:.2f}s")

if __name__ == "__main__":
    main()
