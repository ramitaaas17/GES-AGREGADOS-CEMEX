# -*- coding: utf-8 -*-
"""Módulo Backend de Validación SAP GUI en Vivo para RutaV3_2026.

Permite:
  1. Conexión e interacción observable con SAP GUI (primer plano, ritmo ajustable).
  2. Detección automática de formatos (FPA Filial, CTO Compra, FPA Terceros, CTO Venta).
  3. Extracción normalizada de filas de archivos Excel recibidos de Outlook.
  4. Ejecución en tiempo real en SAP GUI (VK13, Z0116_TRAOPE, etc.) con callbacks.
  5. Cruce y confirmación automática hacia SQLite (staging_rutas) y el Excel Maestro.
"""

import os
import re
import glob
import shutil
import time
import threading
from datetime import datetime
import openpyxl
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
import pandas as pd
import win32com.client

import fpa_validations
import database_manager
import etl_processor

BASE_DIR = os.path.dirname(os.path.abspath(__file__))


# ---------------------------------------------------------------------------
# Conexión y Gestión de SAP GUI
# ---------------------------------------------------------------------------
def is_sapgui_running():
    """Comprueba si el proceso de sesión sapgui.exe está en ejecución."""
    try:
        import subprocess
        out = subprocess.check_output('tasklist /FI "IMAGENAME eq sapgui.exe"', shell=True, text=True)
        return "sapgui.exe" in out.lower()
    except Exception:
        return False


def get_sap_session(bring_to_foreground=True):
    """Conecta a la sesión activa de SAP GUI.

    Si bring_to_foreground es True, maximiza la ventana wnd[0] y la enfoca.
    """
    import pythoncom
    try:
        pythoncom.CoInitialize()
    except Exception:
        pass

    try:
        sap_gui_auto = win32com.client.GetObject("SAPGUI")
    except Exception as e:
        err_msg = str(e)
        if not is_sapgui_running():
            raise RuntimeError(
                "No hay ninguna sesión activa de SAP GUI iniciada en Windows.\n\n"
                "Actualmente solo tienes abierto el lanzador de SAP Logon (o SAP está cerrado).\n"
                "Para poder validar, debes ingresar a tu sistema en SAP Logon (hacer doble clic en el servidor\n"
                "e iniciar sesión con tu usuario y contraseña, de modo que veas la pantalla principal 'SAP Easy Access')."
            )
        else:
            raise RuntimeError(
                f"SAP GUI está abierto pero no permitió la conexión por Scripting ({err_msg}).\n\n"
                "Por favor verifica:\n"
                "1. En SAP Logon > Opciones > Accesibilidad y scripting > Scripting: tener marcado 'Habilitar scripting'.\n"
                "2. Desmarcar 'Notificar cuando un script se conecta a SAP GUI'.\n"
                "3. Que la ventana de SAP esté en la pantalla principal (sin popups bloqueantes)."
            )

    try:
        application = sap_gui_auto.GetScriptingEngine
        if application.Children.Count == 0:
            raise RuntimeError("No hay conexiones abiertas en SAP GUI. Inicia sesión en SAP Logon.")
        connection = application.Children(0)
        if connection.Children.Count == 0:
            raise RuntimeError("No hay sesiones activas en SAP GUI. Abre una sesión en SAP.")
        session = connection.Children(0)

        if bring_to_foreground:
            try:
                wnd0 = session.findById("wnd[0]")
                wnd0.maximize()
                wnd0.setFocus()
            except Exception:
                pass

        return session
    except Exception as e:
        raise RuntimeError(f"Error accediendo a la sesión de SAP GUI: {e}")


def handle_popups(session, delay=0.15):
    """Cierra ventanas emergentes secundarias y avisos informativos (wnd[1]), retornando lista de textos capturados."""
    avisos = []
    try:
        iterations = 0
        while getattr(session, "Children", None) and session.Children.Count > 1 and iterations < 10:
            iterations += 1
            popup = session.findById("wnd[1]")
            if popup:
                txt_list = []
                w_title = getattr(popup, "Text", "").strip()
                try:
                    lbls = find_controls_recursive(popup, "GuiLabel")
                    tfs = find_controls_recursive(popup, "GuiTextField") + find_controls_recursive(popup, "GuiCTextField")
                    for l in lbls:
                        lt = str(getattr(l, "Text", "")).strip()
                        if lt and lt not in txt_list:
                            txt_list.append(lt)
                    for t in tfs:
                        tt = str(getattr(t, "Text", "")).strip()
                        if tt and tt not in txt_list:
                            txt_list.append(tt)
                except Exception:
                    pass

                body_msg = " | ".join(txt_list)
                full_msg = f"[{w_title}] {body_msg}" if w_title and body_msg else (body_msg or w_title)
                if full_msg and full_msg not in avisos:
                    avisos.append(full_msg)

                try:
                    btn0 = popup.findById("tbar[0]/btn[0]")
                    btn0.press()
                except Exception:
                    popup.sendVKey(0)
                time.sleep(delay)
            else:
                break
    except Exception:
        pass
    return avisos


def clean_num(val):
    """Limpia y convierte a float para comparaciones seguras de precios."""
    if val is None:
        return None
    s = str(val).replace(",", "").replace("$", "").replace("MXN", "").replace("USD", "").replace(" ", "").strip()
    try:
        return float(s)
    except Exception:
        return None


def clean_exp_cond(val):
    """Normaliza código de modalidad a '01' (Entregado), '02' (Recogido), '04' (Entrega directa)."""
    if val is None:
        return ""
    s = str(val).strip().lstrip("'").lstrip("‘").lstrip("’").lstrip("´").lstrip('"')
    s_up = s.upper()
    if "DIRECT" in s_up or "INTER" in s_up:
        return "04"
    if "RECOGID" in s_up or "PLANTA" in s_up:
        return "02"
    if "ENTREG" in s_up:
        return "01"
    match = re.search(r"\b(0[1-9]|1[0-9])\b", s)
    if match:
        return match.group(1)
    digits = re.findall(r"\d+", s)
    if digits:
        return digits[0].zfill(2)
    return s.zfill(2) if s.isdigit() else s


def mat_matches(m1, m2):
    """Compara códigos de material con tolerancia a ceros a la izquierda."""
    s1 = str(m1 or "").strip().lstrip("0")
    s2 = str(m2 or "").strip().lstrip("0")
    return bool(s1 and s2 and s1 == s2)


def dest_matches(r, d_target):
    """Compara códigos de destino con tolerancia a ceros a la izquierda contra cualquier campo de cliente/destino."""
    if not d_target:
        return False
    d_clean = str(d_target).strip().lstrip("0")
    candidates = [
        str(r.get("dest_mercancia") or "").strip().lstrip("0"),
        str(r.get("destino") or "").strip().lstrip("0"),
        str(r.get("cliente") or "").strip().lstrip("0"),
        str(r.get("solicitante") or "").strip().lstrip("0"),
        str(r.get("kunnr") or "").strip().lstrip("0")
    ]
    return d_clean in [c for c in candidates if c]


def pos_matches(p1, p2):
    """Compara posiciones numéricas de contrato en SAP y Excel (ej. '20' vs '00020')."""
    if p1 is None or p2 is None:
        return False
    s1 = str(p1).strip()
    s2 = str(p2).strip()
    if not s1 or not s2:
        return False
    if s1 == s2:
        return True
    n1 = clean_num(s1)
    n2 = clean_num(s2)
    if n1 is not None and n2 is not None and int(n1) == int(n2):
        return True
    return False


def is_valid_um(val):
    """Valida si el texto corresponde a una unidad de medida física real y no a un número de documento."""
    if not val:
        return False
    s = str(val).strip().upper()
    if s.isdigit() or len(s) > 6 or re.search(r"\d{3,}", s):
        return False
    return s in ("M3", "M^3", "MT3", "TN", "TO", "TON", "TNS", "KG", "L", "LT", "PZA", "M2", "MLL")


def clean_um(val):
    """Normaliza unidad de medida rechazando números de contrato o identificadores."""
    if val is None:
        return ""
    s = str(val).strip().upper()
    if s.isdigit() or len(s) > 6 or re.search(r"\d{3,}", s):
        return ""
    if s in ("TO", "TON", "TN", "TNS"):
        return "TN"
    if s in ("M3", "M^3", "MT3"):
        return "M3"
    return s


def find_controls_recursive(parent, target_type):
    """Busca controles recursivamente dentro de un contenedor de SAP GUI."""
    found = []
    try:
        count = getattr(parent.Children, "Count", 0)
    except Exception:
        return found

    for i in range(count):
        try:
            child = parent.Children(i)
            c_type = getattr(child, "Type", "")
            if target_type.lower() in c_type.lower():
                found.append(child)
            found.extend(find_controls_recursive(child, target_type))
        except Exception:
            continue
    return found


def find_control_by_name_recursive(parent, target_name):
    """Busca un control cuyo Id o Name contenga target_name de forma recursiva."""
    try:
        count = getattr(parent.Children, "Count", 0)
    except Exception:
        return None
    for i in range(count):
        try:
            child = parent.Children(i)
            c_id = str(getattr(child, "Id", ""))
            c_name = str(getattr(child, "Name", ""))
            if target_name.lower() in c_id.lower() or target_name.lower() in c_name.lower():
                return child
            res = find_control_by_name_recursive(child, target_name)
            if res is not None:
                return res
        except Exception:
            continue
    return None


def find_data_field_by_name(parent, target_name):
    """Busca un campo de datos (GuiTextField, GuiCTextField, GuiComboBox) cuyo Id o Name contenga target_name,
    ignorando GuiLabel para evitar capturar el texto de la etiqueta visible."""
    try:
        count = getattr(parent.Children, "Count", 0)
    except Exception:
        return None
    for i in range(count):
        try:
            child = parent.Children(i)
            c_type = str(getattr(child, "Type", ""))
            c_id = str(getattr(child, "Id", ""))
            c_name = str(getattr(child, "Name", ""))
            if target_name.lower() in c_id.lower() or target_name.lower() in c_name.lower():
                if "label" not in c_type.lower() and not c_id.split("/")[-1].startswith("lbl"):
                    return child
            res = find_data_field_by_name(child, target_name)
            if res is not None:
                return res
        except Exception:
            continue
    return None


# ---------------------------------------------------------------------------
# Funciones Especializadas Z0116_TRAOPE (Contratos de Compra)
# ---------------------------------------------------------------------------

def fill_multiple_selection_popup(session, btn_id, values_list):
    """
    Abre el popup de Selección Múltiple presionando el botón VALU_PUSH correspondiente,
    pega/inserta la lista de valores en la tabla del popup y ejecuta (F8 / btn[8]) para confirmar.
    """
    if not values_list:
        return
    try:
        session.findById(btn_id).press()
        time.sleep(0.3)
        wnd1 = session.findById("wnd[1]")
        for idx, val in enumerate(values_list):
            val_str = str(val).strip()
            if not val_str:
                continue
            for cell_id in [
                f"usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,{idx}]",
                f"usr/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,{idx}]",
                f"usr/tblSAPLALVMTCTRL_9999/ctxtRSCSEL_255-SLOW_I[1,{idx}]",
            ]:
                try:
                    wnd1.findById(cell_id).text = val_str
                    break
                except Exception:
                    pass
        # Confirmar / Ejecutar la selección múltiple (F8 o Enter)
        try:
            wnd1.findById("tbar[0]/btn[8]").press()
        except Exception:
            try:
                wnd1.sendVKey(8)
            except Exception:
                wnd1.sendVKey(0)
        time.sleep(0.3)
    except Exception:
        try:
            session.findById("wnd[1]").close()
        except Exception:
            pass


def clear_screen_inputs(session, exclude=None):
    """Limpia preventivamente todos los campos de entrada editables de la pantalla actual en SAP GUI sin borrar campos de fecha obligatorios."""
    if exclude is None:
        exclude = ["DATAB", "DATBI", "FECHA", "DATE", "GUEEN", "AEDAT", "F000"]
    try:
        usr = session.findById("wnd[0]/usr")
        for i in range(usr.Children.Count):
            ch = usr.Children(i)
            c_type = getattr(ch, "Type", "")
            if c_type in ("GuiTextField", "GuiCTextField"):
                if not getattr(ch, "ReadOnly", False) and getattr(ch, "Changeable", True):
                    cid = getattr(ch, "Id", "").upper()
                    cname = getattr(ch, "Name", "").upper()
                    if any(ex in cid or ex in cname for ex in exclude):
                        continue
                    try:
                        ch.Text = ""
                    except Exception:
                        pass
    except Exception:
        pass


def fill_z0116_selection(session, plant, vendor, material, vendors_list=None, materials_list=None, contracts_list=None):
    """
    Rellena la pantalla de selección de Z0116_TRAOPE / Z00116_TRAOPE:
      - Centro (Plant): S_WERKS-LOW
      - Fecha de creación: S_AEDAT-LOW (01.01.2011) y HIGH (31.12.9999)
      - Material: S_MATNR-LOW
      - Proveedor: S_LIFNR-LOW (o selección múltiple si hay varios)
      - Contrato: S_EBELN-LOW (o vacío si se consultan todos)
    """
    wnd0 = session.findById("wnd[0]")
    clear_screen_inputs(session)

    # 1. Asignar Centro (Plant)
    try:
        session.findById("wnd[0]/usr/ctxtS_WERKS-LOW").text = str(plant).strip()
    except Exception:
        pass

    # 2. Asignar Fecha de creación (>= 01.01.2011)
    try:
        session.findById("wnd[0]/usr/ctxtS_AEDAT-LOW").text = "01.01.2011"
        session.findById("wnd[0]/usr/ctxtS_AEDAT-HIGH").text = "31.12.9999"
    except Exception:
        pass

    # 3. Documento de compras / Contrato
    try:
        clean_contracts = [c for c in (contracts_list or []) if c and not any(k in str(c).upper() for k in ["CREAR", "ALTA", "POR CREAR"])]
        if clean_contracts and len(clean_contracts) == 1:
            session.findById("wnd[0]/usr/ctxtS_EBELN-LOW").text = str(clean_contracts[0]).strip()
            session.findById("wnd[0]/usr/ctxtS_EBELN-HIGH").text = ""
        else:
            session.findById("wnd[0]/usr/ctxtS_EBELN-LOW").text = ""
            session.findById("wnd[0]/usr/ctxtS_EBELN-HIGH").text = ""
    except Exception:
        pass

    # 4. Asignar Proveedor (Vendor / LIFNR)
    clean_vendors = [v for v in (vendors_list or []) if v and v != "None"]
    if clean_vendors and len(clean_vendors) > 1:
        fill_multiple_selection_popup(session, "wnd[0]/usr/btn%_S_LIFNR_%_APP_%-VALU_PUSH", clean_vendors)
    elif vendor:
        try:
            session.findById("wnd[0]/usr/ctxtS_LIFNR-LOW").text = str(vendor).strip()
        except Exception:
            pass

    # 5. Asignar Material
    clean_mats = [m for m in (materials_list or []) if m and m != "None"]
    if clean_mats and len(clean_mats) > 1:
        fill_multiple_selection_popup(session, "wnd[0]/usr/btn%_S_MATNR_%_APP_%-VALU_PUSH", clean_mats)
    elif material:
        try:
            session.findById("wnd[0]/usr/ctxtS_MATNR-LOW").text = str(material).strip()
        except Exception:
            pass

    # 6. Fallback recursivo por si los IDs directos difieren
    textfields = find_controls_recursive(wnd0, "TextField")
    for ch in textfields:
        c_name = getattr(ch, "Name", "").upper()
        if any(k in c_name for k in ["WERK", "PLANT", "CENTRO"]):
            if "LOW" in c_name and not getattr(ch, "Text", ""):
                try: ch.Text = str(plant).strip()
                except Exception: pass
        elif any(k in c_name for k in ["LIFNR", "VENDOR", "PROV", "ACREEDOR"]):
            if "LOW" in c_name and not clean_vendors and vendor and not getattr(ch, "Text", ""):
                try: ch.Text = str(vendor).strip()
                except Exception: pass
        elif any(k in c_name for k in ["MATNR", "MATERIAL"]):
            if "LOW" in c_name and not clean_mats and material and not getattr(ch, "Text", ""):
                try: ch.Text = str(material).strip()
                except Exception: pass
        elif any(k in c_name for k in ["ERDAT", "AEDAT", "FECHA", "DATE", "CREAT"]):
            if "LOW" in c_name and not getattr(ch, "Text", ""):
                try: ch.Text = "01.01.2011"
                except Exception: pass
            elif "HIGH" in c_name and not getattr(ch, "Text", ""):
                try: ch.Text = "31.12.9999"
                except Exception: pass


def extract_all_z0116_from_screen(session):
    """
    Extrae todos los registros visibles en Z0116_TRAOPE / Z00116_TRAOPE
    directamente desde la sesión SAP (GuiGridView o GuiLabel de lista clásica ABAP)
    sin necesidad de exportar a fichero local.
    """
    records = []
    try:
        wnd0 = session.findById("wnd[0]")
    except Exception:
        return records

    # 1. BÚSQUEDA en ALV Grid (GuiGridView)
    grids = find_controls_recursive(wnd0, "GuiGridView")
    if grids:
        grid = grids[0]
        row_count = getattr(grid, "RowCount", 0)
        if row_count > 0:
            cols = []
            try:
                cols_obj = grid.ColumnOrder
                cols = [str(cols_obj(i)) for i in range(cols_obj.Count)]
            except Exception:
                try:
                    cols = list(grid.ColumnOrder)
                except Exception:
                    pass

            for r_i in range(row_count):
                r_dict = {}
                for c in cols:
                    try:
                        v = grid.GetCellValue(r_i, c)
                        if v is not None and str(v).strip() != "":
                            r_dict[c.upper()] = str(v).strip()
                    except Exception:
                        pass

                def get_grid_val(keywords):
                    for k, v in r_dict.items():
                        if any(kw in k for kw in keywords) and v:
                            return v
                    return ""

                cto = get_grid_val(["EBELN", "EVRTN", "CONTRATO", "DOC"])
                pos_raw = get_grid_val(["EBELP", "POSICION", "POS"])
                pos_val = str(int(clean_num(pos_raw))) if clean_num(pos_raw) is not None else pos_raw.strip()

                if cto or pos_val:
                    records.append({
                        "contrato": cto,
                        "posicion": pos_val,
                        "vendor": get_grid_val(["SHIP FROM", "LIFNR", "PROV", "ACREEDOR"]),
                        "shipfrom": get_grid_val(["SHIP FROM", "SHIPFROM"]),
                        "proveedor": get_grid_val(["PROVEEDOR", "LIFNR"]),
                        "plant": get_grid_val(["WERKS", "CENTRO", "PLANT"]),
                        "material": get_grid_val(["MATNR", "MATERIAL"]),
                        "fch_ini": get_grid_val(["KDATB", "VPER_START", "FCH_INI", "DATAB", "DESDE", "VÁLIDO DE", "VALIDO DE"]),
                        "fch_fin": get_grid_val(["KDATE", "VPER_END", "FCH_FIN", "DATBI", "HASTA", "VALIDEZ A"]),
                        "precio_neto": clean_num(get_grid_val(["NETPR", "PRECIO", "PRICE", "KBETR", "IMPORTE"])),
                        "moneda": get_grid_val(["WAERS", "MONEDA", "CURRENCY", "KONWA"]),
                        "um": get_grid_val(["MEINS", "BPRME", "PEINH", "PO_UM", "UM"]),
                        "exp_cond": clean_exp_cond(get_grid_val(["VSBED", "EXP_COND", "COND_EXP", "MODALIDAD", "COND. EXP."])),
                        "kunnr": get_grid_val(["KUNNR", "DESTINO", "KUNWE", "CLIENTE", "SUCURSAL"]),
                        "ind_oblig": get_grid_val(["IND. OBLIG", "IND_OBLIG", "OBLIG", "CONF. PED"]),
                        "ctrl_ctr": get_grid_val(["CTRL-CONF", "CTRL_CONF", "CTRL-CTR", "CTRL_CTR", "CONF"])
                    })
            if records:
                return records

    # 2. BÚSQUEDA en Lista Clásica ABAP (GuiLabel)
    labels = find_controls_recursive(wnd0, "GuiLabel")
    if labels:
        lines_dict = {}
        for l in labels:
            try:
                t = getattr(l, "Top", 0)
                txt = getattr(l, "Text", "").strip()
                if txt:
                    lines_dict.setdefault(t, []).append({
                        "text": txt,
                        "left": getattr(l, "Left", 0),
                        "width": getattr(l, "Width", 0),
                        "id": getattr(l, "Id", "")
                    })
            except Exception:
                continue

        for t in sorted(lines_dict.keys()):
            items = sorted(lines_dict[t], key=lambda x: x["left"])
            line_full = " ".join(x["text"] for x in items)

            has_contract = any(re.match(r"^46\d{8}$", x["text"]) for x in items) or ("4600" in line_full)
            if not has_contract:
                continue

            row_data = {
                "contrato": "",
                "posicion": "",
                "vendor": "",
                "shipfrom": "",
                "proveedor": "",
                "plant": "",
                "material": "",
                "fch_ini": "",
                "fch_fin": "",
                "precio_neto": None,
                "moneda": "MXN",
                "um": "TN",
                "exp_cond": "",
                "kunnr": "",
                "ind_oblig": "",
                "ctrl_ctr": "",
                "line_text": line_full
            }

            for it in items:
                txt = it["text"].strip()
                if re.match(r"^46\d{8}$", txt):
                    row_data["contrato"] = txt
                elif re.match(r"^900\d{6}$", txt):
                    row_data["vendor"] = txt
                    row_data["shipfrom"] = txt
                elif re.match(r"^50\d{5,6}$", txt):
                    row_data["proveedor"] = txt
                elif txt.upper() == "X":
                    row_data["ind_oblig"] = "X"
                elif txt in ("0004", "0001", "001"):
                    row_data["ctrl_ctr"] = txt
                elif re.match(r"^10\d{6}$", txt):
                    row_data["material"] = txt
                elif txt in ("01", "02", "04") and not row_data["exp_cond"]:
                    row_data["exp_cond"] = clean_exp_cond(txt)
                elif txt.isdigit() and len(txt) in (2, 3, 4) and txt not in ("01", "02", "04", "0004") and not row_data["posicion"]:
                    row_data["posicion"] = str(int(txt))
                elif re.match(r"^\d+[\.,]\d{2}$", txt) and row_data["precio_neto"] is None:
                    row_data["precio_neto"] = clean_num(txt)
                elif txt in ("MXN", "USD"):
                    row_data["moneda"] = txt
                elif txt in ("TN", "M3", "KG", "TO"):
                    row_data["um"] = txt
                elif re.match(r"^\d{2}\.\d{2}\.\d{4}$", txt) or re.match(r"^\d{4}-\d{2}-\d{2}$", txt):
                    if not row_data["fch_ini"]: row_data["fch_ini"] = txt
                    else: row_data["fch_fin"] = txt
                elif re.match(r"^[A-Z0-9]{4}$", txt) and txt not in ("MXPO", "MXL", "TN", "TO", "KG", "M3", "MXN", "USD", "7100", "0004", "0001", "001", "POS", "DEST", "CENT", "PROV", "STAT", "TIPO", "ITEM", "CONF") and not txt.startswith("Z0") and not txt.startswith("504") and txt != row_data.get("plant", ""):
                    if not row_data["kunnr"]:
                        row_data["kunnr"] = txt

            if row_data["contrato"]:
                records.append(row_data)

    return records


def parse_z0116_txt_file(file_path):
    """
    Parsea un archivo de texto exportado por SAP (con tabulador o delimitado por pipe '|')
    extrayendo las columnas de contrato, posición, precio, kunnr, etc.
    """
    records = []
    if not os.path.exists(file_path):
        return records

    try:
        content = ""
        for enc in ["latin-1", "cp1252", "utf-8", "utf-16"]:
            try:
                with open(file_path, "r", encoding=enc) as f:
                    content = f.read()
                if content and ("4600" in content or "Doc" in content or "EBELN" in content):
                    break
            except Exception:
                pass

        if not content:
            return records

        lines = [ln for ln in content.splitlines() if ln.strip()]
        header_map = {}
        delimiter = "\t" if any("\t" in ln for ln in lines) else ("|" if any("|" in ln for ln in lines) else None)

        for ln_idx, line in enumerate(lines):
            if delimiter:
                parts = [p.strip() for p in line.split(delimiter)]
                if delimiter == "|" and parts and parts[0] == "":
                    parts = parts[1:]
                if delimiter == "|" and parts and parts[-1] == "":
                    parts = parts[:-1]
            else:
                parts = line.split()

            if not header_map:
                line_upper = " ".join(parts).upper()
                if any(k in line_upper for k in ["DOC. COMPRAS", "DOC.COMPRAS", "CONTRATO", "EBELN", "POS"]) and any(k in line_upper for k in ["PRECIO", "NETPR", "DESTINO", "KUNNR", "MATERIAL"]):
                    for p_i, p_val in enumerate(parts):
                        if p_val:
                            header_map[p_val.strip().upper()] = p_i
                    continue

            if header_map:
                def get_val_by_keys(keys, default="", exclude=None):
                    exclude = [e.upper() for e in (exclude or [])]
                    for k in keys:
                        for h_name, h_idx in header_map.items():
                            if any(e in h_name for e in exclude):
                                continue
                            if k.upper() in h_name and h_idx < len(parts):
                                val = parts[h_idx].strip()
                                if val:
                                    return val
                    return default

                cto = get_val_by_keys(["DOC. COMPRAS", "DOC.COMPRAS", "CONTRATO", "EBELN"])
                if not cto or not re.match(r"^46\d{8}$", cto):
                    if parts and re.match(r"^46\d{8}$", parts[0]):
                        cto = parts[0]
                    else:
                        continue

                pos_raw = get_val_by_keys(["POSICION", "POS", "EBELP"])
                pos_val = str(int(clean_num(pos_raw))) if clean_num(pos_raw) is not None else pos_raw.strip()

                sf_v = get_val_by_keys(["SHIP FROM", "SHIPFROM"])
                prov_v = get_val_by_keys(["PROVEEDOR", "LIFNR"], exclude=["SHIP FROM", "NOMBRE"])

                rec = {
                    "contrato": cto,
                    "posicion": pos_val,
                    "vendor": sf_v or prov_v,
                    "shipfrom": sf_v,
                    "proveedor": prov_v,
                    "plant": get_val_by_keys(["CENTRO", "WERKS", "PLANT"]),
                    "material": get_val_by_keys(["MATERIAL", "MATNR"]),
                    "fch_ini": get_val_by_keys(["VÁLIDO DE", "VALIDO DE", "KDATB", "FCH_INI"]),
                    "fch_fin": get_val_by_keys(["VALIDEZ A", "KDATE", "FCH_FIN"]),
                    "precio_neto": clean_num(get_val_by_keys(["PRECIO NETO PEDIDO", "PRECIO NETO", "NETPR", "PRECIO"])),
                    "moneda": get_val_by_keys(["MONEDA", "WAERS"]),
                    "um": get_val_by_keys(["UM PRECIO PEDIDO", "UM", "MEINS"]),
                    "exp_cond": clean_exp_cond(get_val_by_keys(["COND. EXP.", "COND.EXP.", "EXP_COND", "MODALIDAD", "VSBED"])),
                    "kunnr": get_val_by_keys(["DESTINO", "KUNNR", "SUCURSAL", "CLIENTE"]),
                    "ind_oblig": get_val_by_keys(["IND. OBLIG", "CONF. PED.", "OBLIG."]),
                    "ctrl_ctr": get_val_by_keys(["CTRL-CONF", "CTRL-CTR", "CTRL_CTR", "CTRL_CONF"])
                }
                records.append(rec)
            else:
                if any(re.match(r"^46\d{8}$", p) for p in parts):
                    row_data = {
                        "contrato": "", "posicion": "", "vendor": "", "shipfrom": "", "proveedor": "",
                        "plant": "", "material": "", "fch_ini": "", "fch_fin": "", "precio_neto": None,
                        "moneda": "MXN", "um": "TN", "exp_cond": "", "kunnr": "",
                        "ind_oblig": "", "ctrl_ctr": ""
                    }
                    for p in parts:
                        p_str = p.strip()
                        if re.match(r"^46\d{8}$", p_str):
                            row_data["contrato"] = p_str
                        elif re.match(r"^900\d{6}$", p_str):
                            row_data["vendor"] = p_str
                            row_data["shipfrom"] = p_str
                        elif re.match(r"^50\d{5,6}$", p_str):
                            row_data["proveedor"] = p_str
                        elif p_str.upper() == "X":
                            row_data["ind_oblig"] = "X"
                        elif p_str in ("0004", "0001", "001"):
                            row_data["ctrl_ctr"] = p_str
                        elif re.match(r"^10\d{6}$", p_str):
                            row_data["material"] = p_str
                        elif p_str in ("01", "02", "04") and not row_data["exp_cond"]:
                            row_data["exp_cond"] = clean_exp_cond(p_str)
                        elif p_str.isdigit() and len(p_str) in (2, 3, 4) and p_str not in ("01", "02", "04", "0004") and not row_data["posicion"]:
                            row_data["posicion"] = str(int(p_str))
                        elif re.match(r"^\d+[\.,]\d{2}$", p_str) and row_data["precio_neto"] is None:
                            row_data["precio_neto"] = clean_num(p_str)
                        elif p_str in ("MXN", "USD"):
                            row_data["moneda"] = p_str
                        elif p_str in ("TN", "M3", "KG", "TO"):
                            row_data["um"] = p_str
                        elif re.match(r"^\d{2}\.\d{2}\.\d{4}$", p_str):
                            if not row_data["fch_ini"]: row_data["fch_ini"] = p_str
                            else: row_data["fch_fin"] = p_str
                        elif re.match(r"^[A-Z0-9]{4}$", p_str) and p_str not in ("MXPO", "MXL", "TN", "TO", "KG", "M3", "MXN", "USD", "7100", "0004", "0001", "001", "POS", "DEST"):
                            if not row_data["kunnr"]: row_data["kunnr"] = p_str
                    if row_data["contrato"]:
                        records.append(row_data)

    except Exception:
        pass

    return records


def export_and_read_z0116_excel(session):
    """
    Exportación de respaldo en Z0116_TRAOPE / Z00116_TRAOPE:
    1. Limpia ficheros previos (.txt, sap_z0116_export.txt).
    2. Envía orden de exportación a SAP (%PC / btn[45]).
    3. Selecciona 'Con tabulador' (radRB_2) y define ruta y nombre específicos.
    4. Si aparece aviso 'Fichero ya existe', presiona automáticamente 'Reemplazar' (btn[11] / SPOP-OPTION1).
    5. Lee y parsea tanto ficheros .txt tabulados como .xlsx.
    6. Cierra cualquier ventana modal residual.
    """
    records = []
    try:
        wnd0 = session.findById("wnd[0]")
    except Exception:
        return records

    temp_dir = os.environ.get("TEMP", os.path.expanduser("~\\AppData\\Local\\Temp"))
    sap_doc_dir = os.path.expanduser("~\\Documents\\SAP\\SAP GUI")
    export_filename = "sap_z0116_export.txt"
    target_path = sap_doc_dir if os.path.exists(sap_doc_dir) else BASE_DIR

    # 0. Limpieza previa de archivos residuales
    clean_targets = [
        os.path.join(sap_doc_dir, ".txt"),
        os.path.join(sap_doc_dir, export_filename),
        os.path.join(temp_dir, export_filename),
        os.path.join(BASE_DIR, export_filename),
        os.path.join(r"C:\revision", export_filename),
        os.path.join(r"D:\revision", export_filename),
    ]
    for ct in clean_targets:
        if os.path.exists(ct):
            try: os.remove(ct)
            except Exception: pass

    watch_dirs = [sap_doc_dir, BASE_DIR, temp_dir, r"C:\revision", r"D:\revision"]
    watch_dirs = [d for d in watch_dirs if os.path.exists(d)]

    before_files = set()
    for wd in watch_dirs:
        for pat in ["*SAP*.xlsx", "*sap*.xlsx", export_filename, ".txt"]:
            for f in glob.glob(os.path.join(wd, pat)):
                if not os.path.basename(f).startswith("~$"):
                    before_files.add(os.path.abspath(f))

    # 1. Enviar comando o botón de exportación a SAP
    exported = False
    try:
        session.findById("wnd[0]/tbar[1]/btn[45]").press()
        time.sleep(0.4)
        exported = True
    except Exception:
        pass

    if not exported:
        try:
            session.findById("wnd[0]/tbar[0]/okcd").text = "%PC"
            session.findById("wnd[0]").sendVKey(0)
            time.sleep(0.4)
            exported = True
        except Exception:
            pass

    if not exported:
        for vkey in [43, 19]:
            try:
                wnd0.sendVKey(vkey)
                time.sleep(0.4)
                exported = True
                break
            except Exception:
                pass

    if not exported:
        return records

    # 2. Manejo de diálogos emergentes con resolución automática de 'Fichero ya existe'
    for _ in range(6):
        if session.Children.Count <= 1:
            break
        try:
            top_wnd = session.Children(session.Children.Count - 1)
            w_text = getattr(top_wnd, "Text", "").lower()

            # Caso A: Diálogo de selección de formato -> Seleccionar 'Con tabulador'
            for r_id in ["usr/radRB_2", "usr/radRB_OTHERS", "usr/sub:SAPLSPO5:0101/radSPOPLI-SELFLAG[1,0]"]:
                try:
                    top_wnd.findById(r_id).select()
                    break
                except Exception:
                    pass

            # Caso B: Diálogo de fichero (DY_PATH y DY_FILENAME)
            has_path_fields = False
            try:
                top_wnd.findById("usr/ctxtDY_PATH").text = target_path
                has_path_fields = True
            except Exception:
                pass
            try:
                top_wnd.findById("usr/ctxtDY_FILENAME").text = export_filename
                has_path_fields = True
            except Exception:
                pass

            # Caso C: Diálogo de confirmación 'Fichero ya existe' -> Presionar REEMPLAZAR
            if "existe" in w_text or "already exists" in w_text or "replace" in w_text or "atención" in w_text or "aviso" in w_text:
                replaced = False
                for btn_id in ["usr/btnSPOP-OPTION1", "tbar[0]/btn[11]", "usr/btnBUTTON_1", "usr/btn[11]"]:
                    try:
                        top_wnd.findById(btn_id).press()
                        replaced = True
                        break
                    except Exception:
                        pass
                if not replaced:
                    try: top_wnd.sendVKey(11)
                    except Exception: top_wnd.sendVKey(0)
                time.sleep(0.3)
                continue

            # Si es diálogo de fichero, presionar Reemplazar/Sustituir (btn[11]) o Guardar (btn[0])
            if has_path_fields:
                try:
                    top_wnd.findById("tbar[0]/btn[11]").press()
                except Exception:
                    try: top_wnd.findById("tbar[0]/btn[0]").press()
                    except Exception: top_wnd.sendVKey(0)
                time.sleep(0.4)
                continue

            # Confirmar diálogo genérico
            try:
                top_wnd.findById("tbar[0]/btn[0]").press()
            except Exception:
                top_wnd.sendVKey(0)
            time.sleep(0.3)
        except Exception:
            break

    # Cerrar cualquier popup residual para evitar bloqueos
    for _ in range(3):
        if session.Children.Count > 1:
            try:
                session.Children(session.Children.Count - 1).close()
                time.sleep(0.2)
            except Exception:
                break

    time.sleep(0.5)

    # 3. Detectar nuevo archivo creado (.txt o .xlsx)
    new_file = None
    target_full = os.path.join(target_path, export_filename)
    if os.path.exists(target_full) and target_full not in before_files:
        new_file = target_full
    elif os.path.exists(os.path.join(sap_doc_dir, ".txt")) and os.path.join(sap_doc_dir, ".txt") not in before_files:
        new_file = os.path.join(sap_doc_dir, ".txt")

    if not new_file:
        for _ in range(6):
            for wd in watch_dirs:
                for pat in [export_filename, ".txt", "*SAP*.xlsx", "*sap*.xlsx"]:
                    for cf in glob.glob(os.path.join(wd, pat)):
                        abs_cf = os.path.abspath(cf)
                        if not os.path.basename(abs_cf).startswith("~$") and abs_cf not in before_files:
                            new_file = abs_cf
                            break
                    if new_file: break
                if new_file: break
            if new_file: break
            time.sleep(0.3)

    if not new_file:
        candidates = []
        for wd in watch_dirs:
            for pat in [export_filename, ".txt", "*SAP*.xlsx", "*sap*.xlsx"]:
                for cf in glob.glob(os.path.join(wd, pat)):
                    if not os.path.basename(cf).startswith("~$"):
                        candidates.append(os.path.abspath(cf))
        if candidates:
            candidates.sort(key=os.path.getmtime, reverse=True)
            if time.time() - os.path.getmtime(candidates[0]) < 120:
                new_file = candidates[0]

    # 4. Parsear el archivo según su formato (.txt o .xlsx) y eliminar temporal
    if new_file and os.path.exists(new_file):
        try:
            if new_file.lower().endswith(".txt"):
                records = parse_z0116_txt_file(new_file)
            else:
                temp_copy = os.path.join(temp_dir, f"sap_parse_{int(time.time()*1000)}.xlsx")
                try:
                    shutil.copy2(new_file, temp_copy)
                    wb_e = openpyxl.load_workbook(temp_copy, data_only=True)
                    ws_e = wb_e.active
                    h_map = {str(ws_e.cell(1, c).value).strip(): c for c in range(1, ws_e.max_column + 1) if ws_e.cell(1, c).value}

                    for r_idx in range(2, ws_e.max_row + 1):
                        def gv(col_name):
                            c_idx = h_map.get(col_name)
                            return str(ws_e.cell(r_idx, c_idx).value).strip() if c_idx and ws_e.cell(r_idx, c_idx).value is not None else ""

                        pos_raw = gv("Pos")
                        pos_val = str(int(clean_num(pos_raw))) if clean_num(pos_raw) is not None else pos_raw.strip()

                        rec = {
                            "contrato": gv("Doc. Compras"),
                            "posicion": pos_val,
                            "vendor": gv("Ship From") or gv("Proveedor"),
                            "shipfrom": gv("Ship From"),
                            "proveedor": gv("Proveedor"),
                            "plant": gv("Centro"),
                            "material": gv("Material"),
                            "fch_ini": str(ws_e.cell(r_idx, h_map.get("Válido de", 19)).value or "")[:10],
                            "fch_fin": str(ws_e.cell(r_idx, h_map.get("Validez a", 20)).value or "")[:10],
                            "precio_neto": clean_num(gv("Precio neto pedido")),
                            "moneda": gv("Moneda"),
                            "um": gv("UM Precio Pedido") or gv("UM"),
                            "exp_cond": clean_exp_cond(gv("Cond. Exp.")),
                            "kunnr": gv("Destino"),
                            "ind_oblig": gv("Ind. Oblig. Conf. Ped.") or gv("Ind. Oblig."),
                            "ctrl_ctr": gv("Ctrl-Conf.") or gv("Ctrl-Ctr")
                        }
                        if rec["contrato"] or rec["vendor"]:
                            records.append(rec)

                    wb_e.close()
                except Exception:
                    pass
                finally:
                    if os.path.exists(temp_copy):
                        try: os.remove(temp_copy)
                        except Exception: pass
        finally:
            if os.path.exists(new_file):
                try: os.remove(new_file)
                except Exception: pass

    return records


# ---------------------------------------------------------------------------
# Funciones de Soporte para Transacción Z1017_PROV_INT (Interlocutores Proveedor)
# ---------------------------------------------------------------------------
def parse_z1017_txt_file(file_path):
    """
    Parsea un archivo de texto exportado desde Z1017_PROV_INT:
    |Grupo|Acreedor |Nombre 1 |RFC |OrgC|CPag SF|Soc.|Fu|Prov. EF|Nombre Prov. EF|CPag EF|
    Devuelve lista de diccionarios con cada fila de interlocutor.
    """
    records = []
    if not os.path.exists(file_path):
        return records

    try:
        content = ""
        for enc in ["latin-1", "cp1252", "utf-8", "utf-16"]:
            try:
                with open(file_path, "r", encoding=enc) as f:
                    content = f.read()
                if content and any(k in content.upper() for k in ["ACREEDOR", "PROV. EF", "PROV.EF", "GRUPO", "NOMBRE"]):
                    break
            except Exception:
                pass

        if not content:
            return records

        lines = [ln for ln in content.splitlines() if ln.strip()]
        header_map = {}
        for ln in lines:
            if not "|" in ln or "---" in ln:
                continue
            parts = [p.strip() for p in ln.split("|")]
            if parts and parts[0] == "":
                parts = parts[1:]
            if parts and parts[-1] == "":
                parts = parts[:-1]

            line_upper = " ".join(parts).upper()
            if not header_map:
                if any(k in line_upper for k in ["ACREEDOR", "PROV. EF", "PROV.EF", "NOMBRE"]):
                    for p_i, p_val in enumerate(parts):
                        if p_val:
                            header_map[p_val.strip().upper()] = p_i
                    continue

            if header_map:
                def gv_h(keys, default=""):
                    for k in keys:
                        for h_name, h_idx in header_map.items():
                            if k.upper() in h_name and h_idx < len(parts):
                                val = parts[h_idx].strip()
                                if val: return val
                    return default

                acreedor = gv_h(["ACREEDOR", "VENDOR", "LIFNR"]).lstrip("0")
                prov_ef = gv_h(["PROV. EF", "PROV.EF", "PROV EF", "EMISOR"]).lstrip("0")
                soc = gv_h(["SOC.", "SOC", "SOCIEDAD", "BUKRS"])
                fu = gv_h(["FU", "PARVW", "FUNCION", "FUNCIÓN"])
                nom_prov = gv_h(["NOMBRE PROV. EF", "NOMBRE PROV", "NOMBRE 1"])

                if acreedor or prov_ef:
                    records.append({
                        "acreedor": acreedor,
                        "prov_ef": prov_ef,
                        "sociedad": soc,
                        "funcion": fu,
                        "nombre_prov_ef": nom_prov
                    })

    except Exception:
        pass

    return records


def export_and_read_z1017_excel(session, shipfrom=None):
    """
    Exporta y lee los registros de Z1017_PROV_INT utilizando la secuencia de prov.vbs:
    1. Accede a /nZ1017_PROV_INT.
    2. Rellena el Acreedor (Ship From) en el campo de selección.
    3. Presiona F8 para ejecutar.
    4. En wnd[0]/usr/cntlCONTAINER/shellcont/shell presiona &MB_EXPORT -> &PC.
    5. Selecciona con tabulador y guarda temporalmente sap_z1017_export.txt.
    6. Parsea el archivo a memoria y lo borra de inmediato de disco.
    """
    records = []
    try:
        session.findById("wnd[0]/tbar[0]/okcd").text = "/nZ1017_PROV_INT"
        session.findById("wnd[0]").sendVKey(0)
        time.sleep(0.3)
        handle_popups(session)

        # Rellenar campo de Acreedor si se indicó
        if shipfrom:
            sf_clean = str(shipfrom).strip().lstrip("0")
            for cid in [
                "wnd[0]/usr/ctxtS_LIFNR-LOW", "wnd[0]/usr/ctxtSO_LIFNR-LOW",
                "wnd[0]/usr/ctxtLIFNR-LOW", "wnd[0]/usr/ctxtS_ACREED-LOW",
                "wnd[0]/usr/ctxtACREEDOR-LOW"
            ]:
                try:
                    session.findById(cid).text = sf_clean
                    break
                except Exception:
                    pass
            else:
                for tf in find_controls_recursive(session.findById("wnd[0]"), "GuiCTextField"):
                    c_name = getattr(tf, "Name", "").upper()
                    if any(k in c_name for k in ["LIFNR", "ACREED", "PROV", "VENDOR"]) and "LOW" in c_name:
                        try:
                            tf.text = sf_clean
                            break
                        except Exception:
                            pass

        # Ejecutar F8
        try:
            session.findById("wnd[0]/tbar[1]/btn[8]").press()
        except Exception:
            session.findById("wnd[0]").sendVKey(8)
        time.sleep(0.6)
        handle_popups(session)

        # Verificar mensaje en barra de estado
        sbar = None
        try: sbar = session.findById("wnd[0]/sbar")
        except Exception: pass
        if sbar and sbar.messageType in ("E", "W", "I") and any(k in sbar.text.lower() for k in ["no se encontr", "no exist", "sin datos"]):
            return records

        # Exportar ALV Grid usando prov.vbs
        temp_dir = os.environ.get("TEMP", os.path.expanduser("~\\AppData\\Local\\Temp"))
        sap_doc_dir = os.path.expanduser("~\\Documents\\SAP\\SAP GUI")
        export_filename = "sap_z1017_export.txt"
        target_path = sap_doc_dir if os.path.exists(sap_doc_dir) else BASE_DIR

        target_full = os.path.join(target_path, export_filename)
        if os.path.exists(target_full):
            try: os.remove(target_full)
            except Exception: pass

        shell_grid = None
        for s_path in [
            "wnd[0]/usr/cntlCONTAINER/shellcont/shell",
            "wnd[0]/usr/cntlALV_CONTAINER_0100/shellcont/shell",
            "wnd[0]/usr/cntlGRID1/shellcont/shell"
        ]:
            try:
                shell_grid = session.findById(s_path)
                if shell_grid: break
            except Exception:
                pass

        if not shell_grid:
            for sh in find_controls_recursive(session.findById("wnd[0]"), "GuiShell"):
                try:
                    if getattr(sh, "SubType", "") == "Toolbar" or hasattr(sh, "pressToolbarContextButton"):
                        shell_grid = sh
                        break
                except Exception:
                    pass

        exported = False
        if shell_grid:
            try:
                try: shell_grid.currentCellColumn = "STCD1"
                except Exception: pass
                shell_grid.pressToolbarContextButton("&MB_EXPORT")
                shell_grid.selectContextMenuItem("&PC")
                time.sleep(0.4)
                exported = True
            except Exception:
                pass

        if not exported:
            try:
                session.findById("wnd[0]/tbar[0]/okcd").text = "%PC"
                session.findById("wnd[0]").sendVKey(0)
                time.sleep(0.4)
                exported = True
            except Exception:
                pass

        if not exported:
            return records

        # Diálogos de exportación
        for _ in range(6):
            if session.Children.Count <= 1:
                break
            try:
                top_wnd = session.Children(session.Children.Count - 1)
                w_text = getattr(top_wnd, "Text", "").lower()

                # Caso A: Seleccionar 'Con tabulador'
                for r_id in ["usr/radRB_2", "usr/radRB_OTHERS", "usr/sub:SAPLSPO5:0101/radSPOPLI-SELFLAG[1,0]"]:
                    try:
                        top_wnd.findById(r_id).select()
                        break
                    except Exception:
                        pass

                # Caso B: Diálogo fichero
                has_path = False
                try:
                    top_wnd.findById("usr/ctxtDY_PATH").text = target_path
                    has_path = True
                except Exception:
                    pass
                try:
                    top_wnd.findById("usr/ctxtDY_FILENAME").text = export_filename
                    has_path = True
                except Exception:
                    pass

                # Caso C: Reemplazar si existe
                if "existe" in w_text or "already exists" in w_text or "replace" in w_text:
                    for btn_id in ["usr/btnSPOP-OPTION1", "tbar[0]/btn[11]", "usr/btnBUTTON_1"]:
                        try:
                            top_wnd.findById(btn_id).press()
                            break
                        except Exception:
                            pass
                    time.sleep(0.3)
                    continue

                if has_path:
                    try: top_wnd.findById("tbar[0]/btn[11]").press()
                    except Exception:
                        try: top_wnd.findById("tbar[0]/btn[0]").press()
                        except Exception: top_wnd.sendVKey(0)
                    time.sleep(0.4)
                    continue

                try: top_wnd.findById("tbar[0]/btn[0]").press()
                except Exception: top_wnd.sendVKey(0)
                time.sleep(0.3)
            except Exception:
                break

        # Cerrar popups
        for _ in range(3):
            if session.Children.Count > 1:
                try:
                    session.Children(session.Children.Count - 1).close()
                    time.sleep(0.2)
                except Exception:
                    break

        time.sleep(0.4)

        # Parsear archivo y borrar inmediatamente de disco
        if os.path.exists(target_full):
            try:
                records = parse_z1017_txt_file(target_full)
            finally:
                try: os.remove(target_full)
                except Exception: pass
    except Exception:
        pass
    finally:
        try:
            session.findById("wnd[0]/tbar[0]/okcd").text = "/n"
            session.findById("wnd[0]").sendVKey(0)
            time.sleep(0.2)
        except Exception:
            pass

    return records



def parse_zsdd4501_txt_file(file_path):
    """
    Parsea un archivo de texto exportado por SAP desde ZSDD4501 (con tabulador o delimitado por pipe '|')
    extrayendo las columnas de Contrato / Llave, Posición, Destino, ShipFrom, Material, Centro, Precio, etc.
    """
    records = []
    if not os.path.exists(file_path):
        return records

    try:
        content = ""
        for enc in ["latin-1", "cp1252", "utf-8", "utf-16"]:
            try:
                with open(file_path, "r", encoding=enc) as f:
                    content = f.read()
                if content and any(k in content.upper() for k in ["LLAVE", "ZCP", "DESTINO", "SOCIEDAD", "MATERIAL"]):
                    break
            except Exception:
                pass

        if not content:
            return records

        lines = [ln for ln in content.splitlines() if ln.strip()]
        header_map = {}
        delimiter = "\t" if any("\t" in ln for ln in lines) else ("|" if any("|" in ln for ln in lines) else None)

        for ln in lines:
            if delimiter:
                parts = [p.strip() for p in ln.split(delimiter)]
                if delimiter == "|" and parts and parts[0] == "":
                    parts = parts[1:]
                if delimiter == "|" and parts and parts[-1] == "":
                    parts = parts[:-1]
            else:
                parts = ln.split()

            line_upper = " ".join(parts).upper()
            if not header_map:
                if any(k in line_upper for k in ["LLAVE", "CONTRATO", "DESTINO", "PRECIO NETO", "SOCIEDAD", "MATERIAL"]):
                    for p_i, p_val in enumerate(parts):
                        if p_val:
                            header_map[p_val.strip().upper()] = p_i
                    continue

            if header_map:
                def gv(keys, default="", exclude=None):
                    exclude = [e.upper() for e in (exclude or [])]
                    for k in keys:
                        for h_name, h_idx in header_map.items():
                            if any(e in h_name for e in exclude):
                                continue
                            if k.upper() == h_name and h_idx < len(parts):
                                val = parts[h_idx].strip()
                                if val: return val
                    for k in keys:
                        for h_name, h_idx in header_map.items():
                            if any(e in h_name for e in exclude):
                                continue
                            if re.search(r'\b' + re.escape(k.upper()) + r'\b', h_name) and h_idx < len(parts):
                                val = parts[h_idx].strip()
                                if val: return val
                    return default

                llave = gv(["LLAVE", "KEY"])
                cto = gv(["CONTRATO", "DOC. VENTAS", "DOC VENTAS", "DOCUMENTO VENTAS", "VBELN"], exclude=["TIPO DOCUMENTO", "TIPO"])
                pos_raw = gv(["POSICIÓN", "POSICION", "POSNR", "KPOSN", "ITEM", "POS", "POS."], exclude=["CONDICIONES", "CONDICIÓN", "TIPO", "PSTYV"])

                if (not pos_raw or not pos_raw.isdigit()) and llave and "-" in llave:
                    p_llave = llave.split("-")
                    if not cto:
                        cto = p_llave[0].strip()
                    if len(p_llave) > 1 and p_llave[1].strip().isdigit():
                        pos_raw = p_llave[1].strip()

                if not pos_raw or not pos_raw.isdigit():
                    for p_val in parts:
                        m_p = re.search(r"\b\d{8,10}-(\d+)\b", str(p_val))
                        if m_p:
                            pos_raw = m_p.group(1)
                            break

                pos_val = str(int(pos_raw)) if pos_raw and pos_raw.isdigit() else (pos_raw.strip() if pos_raw else "")

                dest_merc_val = gv([
                    "SHPTO", "KUNWE", "DESTINATARIO DE MERCANCIA", "DESTINATARIO DE MERCANCIAS",
                    "DEST.MCIA", "DEST. MCIA", "DEST.MC", "DESTINATARIO MERCANCIA",
                    "DESTINATARIO", "DESTINO", "SUCURSAL", "OBRA", "FRENTE"
                ], exclude=["SOLICITANTE", "PAGADOR", "FACTURA", "CLIENTE"]).lstrip("0")

                sol_val = gv([
                    "KUNNR", "SOLICITANTE", "CLIENTE", "CLIENTE SOLICITANTE"
                ], exclude=["DESTINATARIO", "MERCANCIA", "FACTURA", "SHPTO"]).lstrip("0")

                sf_val = gv(["SHIPFROM", "SHIP FROM", "KUNBP", "CANTERA", "PROVEEDOR", "LIFNR"])
                dest_val = dest_merc_val or sol_val or sf_val

                mat_val = gv(["MATERIAL", "MATNR", "PRODUCTO"], exclude=["DESC.", "DESC", "DESCRIPCIÓN"])
                centro_val = gv(["CENTRO", "WERKS", "PLANT"])
                p_net_raw = gv(["PRECIO NETO", "PRECIO", "NETPR", "IMPORTE"])
                mon_val = gv(["MONEDA", "WAERS", "CURRENCY"]) or "MXN"
                um_val = clean_um(gv(["VRKME", "MEINS", "ZIELME", "VOLEH", "GEWEI", "KMEIN", "BPRME", "UM", "U.M.", "UNIDAD", "UNIDAD DE MEDIDA"],
                                     exclude=["PRECIO", "IMPORTE", "VALOR", "CONDICION"]))
                if not um_val:
                    for p_val in parts:
                        s_u = str(p_val).strip().upper()
                        if s_u in ("M3", "TN", "TO", "TON", "KG", "L", "LT", "M^3"):
                            um_val = clean_um(s_u)
                            break

                mod_val = clean_exp_cond(gv(["CONDICIÓN EXP.", "CONDICION EXP.", "COND EXP", "COND. EXP.", "EXP_COND", "MODALIDAD", "VSBED"]))
                f_ini = gv(["FECHA INICIO VALIDEZ", "FECHA INICIO", "INICIO", "DATAB", "GUEBG"])
                f_fin = gv(["FECHA FIN VALIDEZ", "FECHA FIN", "FIN", "DATBI", "GUEEN"])

                p_net = None
                if p_net_raw:
                    s_clean = p_net_raw.replace(",", "").replace("$", "").replace(" ", "").strip()
                    try: p_net = float(s_clean)
                    except Exception: pass

                if cto or mat_val or dest_val:
                    records.append({
                        "contrato": cto,
                        "posicion": pos_val,
                        "destino": dest_val,
                        "cliente": sol_val or dest_val,
                        "dest_mercancia": dest_merc_val,
                        "solicitante": sol_val,
                        "dest_factura": sf_val,
                        "kunnr": dest_val,
                        "material": mat_val,
                        "centro": centro_val,
                        "shipfrom": sf_val,
                        "precio_neto": p_net,
                        "moneda": mon_val,
                        "um": um_val,
                        "exp_cond": mod_val,
                        "fch_ini": f_ini,
                        "fch_fin": f_fin,
                    })
            else:
                row_dict = {
                    "contrato": "", "posicion": "", "cliente": "", "dest_mercancia": "",
                    "dest_factura": "", "kunnr": "", "material": "", "centro": "",
                    "shipfrom": "", "precio_neto": None, "moneda": "MXN", "um": "",
                    "exp_cond": "", "fch_ini": "", "fch_fin": ""
                }
                for p in parts:
                    p_str = p.strip()
                    if "-" in p_str and re.match(r"^\d{8,10}-\d{1,4}$", p_str):
                        c, po = p_str.split("-")
                        row_dict["contrato"] = c
                        row_dict["posicion"] = po
                    elif re.match(r"^2\d{8,9}$", p_str) or re.match(r"^02\d{8,9}$", p_str):
                        row_dict["contrato"] = p_str
                    elif re.match(r"^10\d{6}$", p_str):
                        row_dict["material"] = p_str
                    elif re.match(r"^900\d{6}$", p_str):
                        row_dict["shipfrom"] = p_str
                    elif re.match(r"^6[5-7]\d{6,7}$", p_str):
                        row_dict["cliente"] = p_str.lstrip("0")
                        row_dict["dest_mercancia"] = row_dict["cliente"]
                        row_dict["kunnr"] = row_dict["cliente"]
                    elif re.match(r"^D\d{3}$", p_str):
                        row_dict["centro"] = p_str
                    elif p_str.upper() in ("M3", "TN", "TO", "TON", "KG", "L", "LT", "M^3"):
                        row_dict["um"] = clean_um(p_str)
                if not row_dict["posicion"]:
                    row_dict["posicion"] = ""
                # if not row_dict["um"]: row_dict["um"] = ""
                if row_dict["contrato"] or row_dict["material"]:
                    records.append(row_dict)

    except Exception as e:
        pass

    return records


def export_and_read_zsdd4501_excel(session):
    """
    Exportación de respaldo en ZSDD4501 (análoga a la implementada en ctos compra):
    1. Limpia ficheros previos (sap_zsdd4501_export.txt).
    2. Envía orden de exportación a SAP (%PC / btn[45] / sendVKey(43)).
    3. Selecciona 'Con tabulador' (radRB_2) y define ruta y nombre específicos.
    4. Si aparece aviso 'Fichero ya existe', presiona automáticamente 'Reemplazar' (btn[11] / SPOP-OPTION1).
    5. Lee y parsea tanto ficheros .txt tabulados como .xlsx usando parse_zsdd4501_txt_file.
    6. Cierra cualquier ventana modal residual para no bloquear la sesión SAP.
    """
    records = []
    try:
        wnd0 = session.findById("wnd[0]")
    except Exception:
        return records

    temp_dir = os.environ.get("TEMP", os.path.expanduser("~\\AppData\\Local\\Temp"))
    sap_doc_dir = os.path.expanduser("~\\Documents\\SAP\\SAP GUI")
    export_filename = "sap_zsdd4501_export.txt"
    target_path = sap_doc_dir if os.path.exists(sap_doc_dir) else BASE_DIR

    clean_targets = [
        os.path.join(sap_doc_dir, ".txt"),
        os.path.join(sap_doc_dir, export_filename),
        os.path.join(temp_dir, export_filename),
        os.path.join(BASE_DIR, export_filename),
        os.path.join(r"C:\revision", export_filename),
        os.path.join(r"D:\revision", export_filename),
    ]
    for ct in clean_targets:
        if os.path.exists(ct):
            try: os.remove(ct)
            except Exception: pass

    watch_dirs = [sap_doc_dir, BASE_DIR, temp_dir, r"C:\revision", r"D:\revision"]
    watch_dirs = [d for d in watch_dirs if os.path.exists(d)]

    before_files = set()
    for wd in watch_dirs:
        for pat in ["*SAP*.xlsx", "*sap*.xlsx", export_filename, ".txt"]:
            for f in glob.glob(os.path.join(wd, pat)):
                if not os.path.basename(f).startswith("~$"):
                    before_files.add(os.path.abspath(f))

    # 1. Enviar comando o botón de exportación a SAP según zsddnew.vbs
    exported = False
    try:
        session.findById("wnd[0]").maximize()
        alv = session.findById("wnd[0]/usr/cntlALV_CONTAINER_0100/shellcont/shell")
        alv.pressToolbarContextButton("&MB_VIEW")
        time.sleep(0.3)
        alv.selectContextMenuItem("&PRINT_BACK_PREVIEW")
        time.sleep(0.5)
        session.findById("wnd[0]/tbar[1]/btn[43]").press()
        time.sleep(0.4)
        exported = True
    except Exception:
        pass

    if not exported:
        try:
            session.findById("wnd[0]/tbar[1]/btn[45]").press()
            time.sleep(0.4)
            exported = True
        except Exception:
            pass

    if not exported:
        try:
            session.findById("wnd[0]/tbar[0]/okcd").text = "%PC"
            session.findById("wnd[0]").sendVKey(0)
            time.sleep(0.4)
            exported = True
        except Exception:
            pass

    if not exported:
        for vkey in [43, 19]:
            try:
                wnd0.sendVKey(vkey)
                time.sleep(0.4)
                exported = True
                break
            except Exception:
                pass

    if not exported:
        return records

    # 2. Manejo de diálogos emergentes con resolución automática de 'Fichero ya existe'
    for _ in range(6):
        if session.Children.Count <= 1:
            break
        try:
            top_wnd = session.Children(session.Children.Count - 1)
            w_text = getattr(top_wnd, "Text", "").lower()

            # Caso A: Diálogo de formato -> 'Con tabulador'
            for r_id in ["usr/radRB_2", "usr/radRB_OTHERS", "usr/sub:SAPLSPO5:0101/radSPOPLI-SELFLAG[1,0]"]:
                try:
                    top_wnd.findById(r_id).select()
                    break
                except Exception:
                    pass

            # Caso B: Diálogo de fichero (DY_PATH y DY_FILENAME)
            has_path_fields = False
            try:
                top_wnd.findById("usr/ctxtDY_PATH").text = target_path
                has_path_fields = True
            except Exception:
                pass
            try:
                top_wnd.findById("usr/ctxtDY_FILENAME").text = export_filename
                has_path_fields = True
            except Exception:
                pass

            # Caso C: Diálogo 'Fichero ya existe' -> Reemplazar
            if "existe" in w_text or "already exists" in w_text or "replace" in w_text or "atención" in w_text or "aviso" in w_text:
                replaced = False
                for btn_id in ["usr/btnSPOP-OPTION1", "tbar[0]/btn[11]", "usr/btnBUTTON_1", "usr/btn[11]"]:
                    try:
                        top_wnd.findById(btn_id).press()
                        replaced = True
                        break
                    except Exception:
                        pass
                if not replaced:
                    try: top_wnd.sendVKey(11)
                    except Exception: top_wnd.sendVKey(0)
                time.sleep(0.3)
                continue

            if has_path_fields:
                try:
                    top_wnd.findById("tbar[0]/btn[11]").press()
                except Exception:
                    try: top_wnd.findById("tbar[0]/btn[0]").press()
                    except Exception: top_wnd.sendVKey(0)
                time.sleep(0.4)
                continue

            try:
                top_wnd.findById("tbar[0]/btn[0]").press()
            except Exception:
                top_wnd.sendVKey(0)
            time.sleep(0.3)
        except Exception:
            break

    # Cerrar popups residuales
    for _ in range(3):
        if session.Children.Count > 1:
            try:
                session.Children(session.Children.Count - 1).close()
                time.sleep(0.2)
            except Exception:
                break

    time.sleep(0.5)

    # 3. Detectar nuevo archivo creado
    new_file = None
    target_full = os.path.join(target_path, export_filename)
    if os.path.exists(target_full) and target_full not in before_files:
        new_file = target_full
    elif os.path.exists(os.path.join(sap_doc_dir, ".txt")) and os.path.join(sap_doc_dir, ".txt") not in before_files:
        new_file = os.path.join(sap_doc_dir, ".txt")

    if not new_file:
        for _ in range(6):
            for wd in watch_dirs:
                for pat in [export_filename, ".txt", "*SAP*.xlsx", "*sap*.xlsx"]:
                    for cf in glob.glob(os.path.join(wd, pat)):
                        abs_cf = os.path.abspath(cf)
                        if not os.path.basename(abs_cf).startswith("~$") and abs_cf not in before_files:
                            new_file = abs_cf
                            break
                    if new_file: break
                if new_file: break
            if new_file: break
            time.sleep(0.3)

    if not new_file:
        candidates = []
        for wd in watch_dirs:
            for pat in [export_filename, ".txt", "*SAP*.xlsx", "*sap*.xlsx"]:
                for cf in glob.glob(os.path.join(wd, pat)):
                    if not os.path.basename(cf).startswith("~$"):
                        candidates.append(os.path.abspath(cf))
        if candidates:
            candidates.sort(key=os.path.getmtime, reverse=True)
            if time.time() - os.path.getmtime(candidates[0]) < 120:
                new_file = candidates[0]

    # 4. Parsear el archivo según su formato (.txt o .xlsx) y eliminar temporal
    if new_file and os.path.exists(new_file):
        try:
            if new_file.lower().endswith(".txt"):
                records = parse_zsdd4501_txt_file(new_file)
            else:
                temp_copy = os.path.join(temp_dir, f"sap_zsdd_parse_{int(time.time()*1000)}.xlsx")
                try:
                    shutil.copy2(new_file, temp_copy)
                    wb_e = openpyxl.load_workbook(temp_copy, data_only=True)
                    ws_e = wb_e.active
                    h_map = {str(ws_e.cell(1, c).value).strip().upper(): c for c in range(1, ws_e.max_column + 1) if ws_e.cell(1, c).value}

                    def gv_x(row_i, keys, exclude=None):
                        exclude = [e.upper() for e in (exclude or [])]
                        for k in keys:
                            for h_name, h_idx in h_map.items():
                                if any(e in h_name for e in exclude):
                                    continue
                                if k.upper() == h_name:
                                    v = ws_e.cell(row_i, h_idx).value
                                    if v is not None and str(v).strip():
                                        return str(v).strip()
                        for k in keys:
                            for h_name, h_idx in h_map.items():
                                if any(e in h_name for e in exclude):
                                    continue
                                if re.search(r'\b' + re.escape(k.upper()) + r'\b', h_name):
                                    v = ws_e.cell(row_i, h_idx).value
                                    if v is not None and str(v).strip():
                                        return str(v).strip()
                        return ""

                    for r_idx in range(2, ws_e.max_row + 1):
                        llave = gv_x(r_idx, ["LLAVE", "KEY"])
                        cto = gv_x(r_idx, ["CONTRATO", "DOC. VENTAS", "DOC VENTAS", "DOCUMENTO VENTAS", "VBELN"], exclude=["TIPO DOCUMENTO", "TIPO"])
                        pos_raw = gv_x(r_idx, ["POSICIÓN", "POSICION", "POS", "POS."], exclude=["CONDICIONES", "CONDICIÓN"])

                        if llave and "-" in llave:
                            p_llave = llave.split("-")
                            if not cto: cto = p_llave[0].strip()
                            if not pos_raw and len(p_llave) > 1: pos_raw = p_llave[1].strip()

                        dest_val = gv_x(r_idx, ["DESTINO", "SHPTO", "KUNWE", "FRENTE", "CLIENTE", "SUCURSAL", "OBRA"]).lstrip("0")
                        sf_val = gv_x(r_idx, ["SHIPFROM", "SHIP FROM", "KUNBP", "CANTERA", "PROVEEDOR", "LIFNR"])
                        mat_val = gv_x(r_idx, ["MATERIAL", "MATNR", "PRODUCTO"], exclude=["DESC.", "DESC", "DESCRIPCIÓN"])
                        centro_val = gv_x(r_idx, ["CENTRO", "WERKS", "PLANT"])
                        p_net = clean_num(gv_x(r_idx, ["PRECIO NETO", "PRECIO", "NETPR", "IMPORTE"]))
                        mon_val = gv_x(r_idx, ["MONEDA", "WAERS", "CURRENCY"]) or "MXN"
                        um_val = gv_x(r_idx, ["UM", "MEINS", "UNIDAD"])
                        mod_val = clean_exp_cond(gv_x(r_idx, ["CONDICIÓN EXP.", "CONDICION EXP.", "COND EXP", "COND. EXP.", "EXP_COND", "MODALIDAD", "VSBED"]))
                        f_ini = gv_x(r_idx, ["FECHA INICIO VALIDEZ", "FECHA INICIO", "INICIO", "DATAB", "GUEBG"])
                        f_fin = gv_x(r_idx, ["FECHA FIN VALIDEZ", "FECHA FIN", "FIN", "DATBI", "GUEEN"])

                        pos_val = str(int(clean_num(pos_raw))) if clean_num(pos_raw) is not None else pos_raw.strip()

                        if cto or mat_val or dest_val:
                            records.append({
                                "contrato": cto,
                                "posicion": pos_val,
                                "cliente": dest_val,
                                "dest_mercancia": dest_val,
                                "dest_factura": sf_val,
                                "kunnr": dest_val,
                                "material": mat_val,
                                "centro": centro_val,
                                "shipfrom": sf_val,
                                "precio_neto": p_net,
                                "moneda": mon_val,
                                "um": um_val,
                                "exp_cond": mod_val,
                                "fch_ini": f_ini,
                                "fch_fin": f_fin,
                            })

                    wb_e.close()
                except Exception:
                    pass
                finally:
                    if os.path.exists(temp_copy):
                        try: os.remove(temp_copy)
                        except Exception: pass
        finally:
            if os.path.exists(new_file):
                try: os.remove(new_file)
                except Exception: pass

    return records


def filter_alv_grid_z0116_by_kunnr(session, kunnr_target):
    """Aplica el filtro de Destino (KUNNR) en el ALV Grid de Z0116_TRAOPE según filtrodestino.vbs."""
    if not kunnr_target or not session:
        return False
    k_clean = str(kunnr_target).strip().upper()
    try:
        shell_grid = None
        for s_path in [
            "wnd[0]/usr/cntlGRID1/shellcont/shell",
            "wnd[0]/usr/cntlCONTAINER/shellcont/shell",
            "wnd[0]/usr/cntlALV_CONTAINER_0100/shellcont/shell"
        ]:
            try:
                shell_grid = session.findById(s_path)
                if shell_grid:
                    break
            except Exception:
                pass

        if not shell_grid:
            grids = find_controls_recursive(session.findById("wnd[0]"), "GuiGridView")
            if grids:
                shell_grid = grids[0]

        if not shell_grid:
            return False

        try:
            shell_grid.setCurrentCell(-1, "KUNNR")
        except Exception:
            pass
        try:
            shell_grid.selectColumn("KUNNR")
        except Exception:
            pass

        # Botón btn[29] (Filtro)
        try:
            session.findById("wnd[0]/tbar[1]/btn[29]").press()
        except Exception:
            try:
                shell_grid.pressToolbarButton("&FILTER")
            except Exception:
                return False

        time.sleep(0.4)

        # Escribir destino en la ventana emergente de filtro
        if session.Children.Count > 1:
            top_wnd = session.Children(session.Children.Count - 1)
            filled = False
            for field_id in [
                "usr/ssub%_SUBSCREEN_FREESEL:SAPLSSEL:1105/ctxt%%DYN001-LOW",
                "usr/ctxt%%DYN001-LOW",
                "usr/ssub%_SUBSCREEN_FREESEL:SAPLSSEL:1105/txt%%DYN001-LOW",
                "usr/txt%%DYN001-LOW"
            ]:
                try:
                    top_wnd.findById(field_id).text = k_clean
                    filled = True
                    break
                except Exception:
                    pass

            if not filled:
                for ctxt in find_controls_recursive(top_wnd, "GuiCTextField") + find_controls_recursive(top_wnd, "GuiTextField"):
                    if "DYN" in getattr(ctxt, "Name", "").upper() and "LOW" in getattr(ctxt, "Name", "").upper():
                        try:
                            ctxt.text = k_clean
                            filled = True
                            break
                        except Exception:
                            pass

            if filled:
                try:
                    top_wnd.findById("tbar[0]/btn[0]").press()
                except Exception:
                    top_wnd.sendVKey(0)
                time.sleep(0.4)
                return True
            else:
                try:
                    top_wnd.close()
                except Exception:
                    pass
    except Exception:
        pass
    return False


def scan_universal_z0116(session, target_kunnr="", target_pos="", target_contract="", target_material="", target_price=None):
    """
    Extractor universal para Z0116_TRAOPE:
    Extrae los registros en pantalla y devuelve el registro coincidente para la solicitud.
    Aplica filtro por KUNNR en pantalla si no se especificó posición fija (Añadir línea).
    """
    res_data = {
        "encontrado": False,
        "contrato": "",
        "posicion": "",
        "precio_neto": None,
        "moneda": "",
        "um": "",
        "fch_ini": "",
        "fch_fin": "",
        "exp_cond_sap": "",
        "kunnr_sap": "",
        "detalle": ""
    }

    target_dest_clean = str(target_kunnr).strip().upper() if target_kunnr else ""
    target_pos_clean = str(target_pos).strip() if target_pos else ""
    target_cto_clean = str(target_contract).strip() if target_contract else ""
    target_mat_clean = str(target_material).strip().lstrip("0") if target_material else ""
    p_target_num = clean_num(target_price)

    # Si no hay posición pero sí destino, aplicar el filtro ALV en pantalla según filtrodestino.vbs
    if not target_pos_clean and target_dest_clean:
        filter_alv_grid_z0116_by_kunnr(session, target_dest_clean)

    all_rows = extract_all_z0116_from_screen(session)
    if not all_rows:
        return res_data

    # Filtrar primero por contrato si está indicado
    candidates = [r for r in all_rows if not target_cto_clean or r.get("contrato") == target_cto_clean]
    if not candidates:
        candidates = all_rows

    matched_row = None

    # 1. Coincidencia por Posición
    if target_pos_clean:
        for r_cand in candidates:
            if pos_matches(r_cand.get("posicion"), target_pos_clean):
                matched_row = r_cand
                break

    # 2. Coincidencia por Destino / KUNNR + Material (MATNR)
    if not matched_row and target_dest_clean and target_mat_clean:
        matching_k_m = []
        for r_cand in candidates:
            c_kunnr = str(r_cand.get("kunnr", "")).strip().upper()
            c_mat = str(r_cand.get("material", "")).strip().lstrip("0")
            if (target_dest_clean == c_kunnr or target_dest_clean in r_cand.get("line_text", "").upper()) and c_mat == target_mat_clean:
                matching_k_m.append(r_cand)

        if matching_k_m:
            if p_target_num is not None:
                for r_cand in matching_k_m:
                    p_sap = clean_num(r_cand.get("precio_neto"))
                    if p_sap is not None and abs(p_sap - p_target_num) <= 0.05:
                        matched_row = r_cand
                        break
            if not matched_row:
                matching_k_m.sort(key=lambda x: (clean_num(x.get("posicion")) or 0), reverse=True)
                matched_row = matching_k_m[0]

            if len(matching_k_m) > 1 and matched_row:
                detalles_pos = [f"Pos {x.get('posicion', '')}: ${clean_num(x.get('precio_neto')) or 0:,.2f}" for x in matching_k_m]
                res_data["aviso_multilinea"] = f"AVISO MULTILÍNEA: {len(matching_k_m)} posiciones detectadas en SAP para Destino {target_dest_clean} y Material {target_mat_clean} ({', '.join(detalles_pos)}). Se desempató seleccionando Pos {matched_row.get('posicion')}"

    # 3. Coincidencia por Material (si Destino no coincide o no se especificó)
    if not matched_row and not target_pos_clean and target_mat_clean:
        matching_m = []
        for r_cand in candidates:
            c_mat = str(r_cand.get("material", "")).strip().lstrip("0")
            if c_mat == target_mat_clean:
                matching_m.append(r_cand)

        if matching_m:
            if p_target_num is not None:
                for r_cand in matching_m:
                    p_sap = clean_num(r_cand.get("precio_neto"))
                    if p_sap is not None and abs(p_sap - p_target_num) <= 0.05:
                        matched_row = r_cand
                        break
            if not matched_row:
                matching_m.sort(key=lambda x: (clean_num(x.get("posicion")) or 0), reverse=True)
                matched_row = matching_m[0]

            if len(matching_m) > 1 and matched_row:
                detalles_pos = [f"Pos {x.get('posicion', '')}: ${clean_num(x.get('precio_neto')) or 0:,.2f}" for x in matching_m]
                res_data["aviso_multilinea"] = f"AVISO MULTILÍNEA: {len(matching_m)} posiciones detectadas en SAP para Material {target_mat_clean} ({', '.join(detalles_pos)}). Se desempató seleccionando Pos {matched_row.get('posicion')}"

    # 4. Fallback: Coincidencia por Destino solo (ÚNICAMENTE si NO se especificó Material)
    if not matched_row and not target_pos_clean and target_dest_clean and not target_mat_clean:
        for r_cand in candidates:
            c_kunnr = str(r_cand.get("kunnr", "")).strip().upper()
            if target_dest_clean == c_kunnr or target_dest_clean in r_cand.get("line_text", "").upper():
                matched_row = r_cand
                break

    # 5. Fallback final: Primer candidato si coincide el contrato (solo si no vino posición, destino ni material)
    if not matched_row and not target_pos_clean and not target_dest_clean and not target_mat_clean and candidates:
        matched_row = candidates[0]

    if matched_row:
        res_data["encontrado"] = True
        res_data["contrato"] = matched_row.get("contrato", "")
        res_data["posicion"] = matched_row.get("posicion", "")
        res_data["precio_neto"] = matched_row.get("precio_neto")
        res_data["moneda"] = matched_row.get("moneda", "MXN")
        res_data["um"] = matched_row.get("um", "TN")
        res_data["fch_ini"] = matched_row.get("fch_ini", "")
        res_data["fch_fin"] = matched_row.get("fch_fin", "")
        res_data["exp_cond_sap"] = matched_row.get("exp_cond", "")
        res_data["kunnr_sap"] = matched_row.get("kunnr", "")
        res_data["ind_oblig"] = matched_row.get("ind_oblig", "")
        res_data["ctrl_ctr"] = matched_row.get("ctrl_ctr", "")
        res_data["proveedor"] = matched_row.get("proveedor", "")
        res_data["shipfrom"] = matched_row.get("shipfrom", matched_row.get("vendor", ""))
        res_data["aviso_multilinea"] = res_data.get("aviso_multilinea", "")
        res_data["detalle"] = f"Pantalla SAP: Cto={res_data['contrato']} Pos={res_data['posicion']} | Precio={res_data['precio_neto']} {res_data['moneda']} | KUNNR={res_data['kunnr_sap']} ExpCond={res_data['exp_cond_sap']}"

    return res_data


def validate_z0116_contract(session, plant, vendor, material, destino_kunnr, exp_cond_excel, posicion_excel=None, contrato_excel=None, preloaded_records=None, precio_excel=None):
    """Ejecuta Z0116_TRAOPE y extrae datos de contratos (usando exportación o grid/lista)."""
    res = {
        "encontrado": False,
        "contrato": "",
        "posicion": "",
        "precio_neto": None,
        "moneda": "",
        "um": "",
        "fch_ini": "",
        "fch_fin": "",
        "exp_cond_sap": "",
        "kunnr_sap": "",
        "aviso_multilinea": "",
        "validacion_kunnr_04": "N/A",
        "sbar_msg": "",
        "detalle": ""
    }

    c_target = str(contrato_excel).strip() if contrato_excel else ""
    if any(k in c_target.upper() for k in ["CREAR", "ALTA", "POR CREAR"]):
        c_target = ""

    pos_target = str(int(clean_num(posicion_excel))) if clean_num(posicion_excel) is not None else (str(posicion_excel).strip() if posicion_excel else "")
    k_target = str(destino_kunnr).strip().upper() if destino_kunnr else ""
    v_target = str(vendor).strip()
    m_target = str(material).strip()

    # Si ya tenemos registros precargados, buscar coincidencia directamente en memoria
    if preloaded_records:
        match = None
        candidates = [s for s in preloaded_records if not c_target or str(s.get("contrato", "")).strip() == c_target]
        if not candidates:
            candidates = preloaded_records

        p_target_num = clean_num(precio_excel)
        m_target_clean = m_target.lstrip("0") if m_target else ""

        # 1. Coincidencia por Posición
        if pos_target:
            for s in candidates:
                if pos_matches(s.get("posicion"), pos_target):
                    match = s
                    break

        # 2. Coincidencia por Destino / KUNNR + Material (MATNR)
        if not match and k_target and m_target_clean:
            matching_k_m = []
            for s in candidates:
                s_kunnr = str(s.get("kunnr", "")).strip().upper()
                s_mat = str(s.get("material", "")).strip().lstrip("0")
                if s_kunnr == k_target and s_mat == m_target_clean:
                    matching_k_m.append(s)

            if matching_k_m:
                # Desempate si hay más de 1 posición para el mismo Destino y Material:
                # a) Coincidencia de precio
                if p_target_num is not None:
                    for s in matching_k_m:
                        p_sap = clean_num(s.get("precio_neto"))
                        if p_sap is not None and abs(p_sap - p_target_num) <= 0.05:
                            match = s
                            break
                # b) Si ninguna coincide en precio o no hay precio, tomar la más reciente/vigente
                if not match:
                    matching_k_m.sort(key=lambda x: (clean_num(x.get("posicion")) or 0), reverse=True)
                    match = matching_k_m[0]

                if len(matching_k_m) > 1 and match:
                    detalles_pos = [f"Pos {x.get('posicion', '')}: ${clean_num(x.get('precio_neto')) or 0:,.2f}" for x in matching_k_m]
                    res["aviso_multilinea"] = f"AVISO MULTILÍNEA: {len(matching_k_m)} posiciones detectadas en SAP para Destino {k_target} y Material {m_target_clean} ({', '.join(detalles_pos)}). Se desempató seleccionando Pos {match.get('posicion')}"

        # 3. Coincidencia por Vendor + Material (si NO se especificó destino o destino difiere)
        if not match and not pos_target and m_target_clean:
            matching_m = []
            for s in candidates:
                s_vendor = str(s.get("vendor", "")).strip()
                s_mat = str(s.get("material", "")).strip().lstrip("0")
                if s_mat == m_target_clean and (not v_target or s_vendor == v_target):
                    matching_m.append(s)

            if matching_m:
                if p_target_num is not None:
                    for s in matching_m:
                        p_sap = clean_num(s.get("precio_neto"))
                        if p_sap is not None and abs(p_sap - p_target_num) <= 0.05:
                            match = s
                            break
                if not match:
                    matching_m.sort(key=lambda x: (clean_num(x.get("posicion")) or 0), reverse=True)
                    match = matching_m[0]

                if len(matching_m) > 1 and match:
                    detalles_pos = [f"Pos {x.get('posicion', '')}: ${clean_num(x.get('precio_neto')) or 0:,.2f}" for x in matching_m]
                    res["aviso_multilinea"] = f"AVISO MULTILÍNEA: {len(matching_m)} posiciones detectadas en SAP para Material {m_target_clean} ({', '.join(detalles_pos)}). Se desempató seleccionando Pos {match.get('posicion')}"

        # 4. Fallback: Coincidencia por Destino solo (ÚNICAMENTE si NO se especificó Material)
        if not match and not pos_target and k_target and not m_target_clean:
            for s in candidates:
                s_kunnr = str(s.get("kunnr", "")).strip().upper()
                if s_kunnr == k_target:
                    match = s
                    break

        if match:
            res["encontrado"] = True
            res["contrato"] = match["contrato"]
            res["posicion"] = match["posicion"]
            res["precio_neto"] = match["precio_neto"]
            res["moneda"] = match.get("moneda", "MXN")
            res["um"] = match.get("um", "TN")
            res["fch_ini"] = match.get("fch_ini", "")
            res["fch_fin"] = match.get("fch_fin", "")
            res["exp_cond_sap"] = match.get("exp_cond", "")
            res["kunnr_sap"] = match.get("kunnr", "")
            res["ind_oblig"] = match.get("ind_oblig", "")
            res["ctrl_ctr"] = match.get("ctrl_ctr", "")
            res["proveedor"] = match.get("proveedor", "")
            res["shipfrom"] = match.get("shipfrom", match.get("vendor", ""))
            res["detalle"] = f"SAP: Cto={match['contrato']} Pos={match['posicion']} | Precio={match['precio_neto']} {res['moneda']} | KUNNR={match.get('kunnr', '')}"

            clean_exp = clean_exp_cond(exp_cond_excel)
            if clean_exp == "04":
                if res["kunnr_sap"] and res["kunnr_sap"].strip() != "":
                    res["validacion_kunnr_04"] = f"OK (KUNNR={res['kunnr_sap']})"
                else:
                    res["validacion_kunnr_04"] = "ERROR: KUNNR VACÍO EN CONDICIÓN 04"
            else:
                res["validacion_kunnr_04"] = "N/A (NO ES CONDICION 04)"

            return res

    try:
        # 1. Resetear e ingresar a Z0116_TRAOPE / Z00116_TRAOPE
        session.findById("wnd[0]/tbar[0]/okcd").text = "/nZ0116_TRAOPE"
        session.findById("wnd[0]").sendVKey(0)
        handle_popups(session)

        sbar = session.findById("wnd[0]/sbar")
        if sbar and sbar.messageType in ("E", "A") and "no existe" in sbar.text.lower():
            session.findById("wnd[0]/tbar[0]/okcd").text = "/nZ00116_TRAOPE"
            session.findById("wnd[0]").sendVKey(0)
            handle_popups(session)

        # 2. Rellenar criterios de selección
        fill_z0116_selection(
            session, plant=plant, vendor=vendor, material=material,
            contracts_list=[c_target] if c_target else None
        )

        # 3. Ejecutar (F8)
        try:
            session.findById("wnd[0]/tbar[1]/btn[8]").press()
        except Exception:
            session.findById("wnd[0]").sendVKey(8)
        time.sleep(0.5)
        handle_popups(session)

        # 4. Verificar barra de estado
        sbar = session.findById("wnd[0]/sbar")
        sbar_msg = sbar.text.strip() if sbar else ""
        res["sbar_msg"] = sbar_msg

        if sbar and sbar.messageType in ("E", "A"):
            res["detalle"] = f"Error SAP ({sbar.messageType}): {sbar_msg}"
            return res

        # 5. Intentar primero escaneo universal en pantalla
        scanned = scan_universal_z0116(
            session,
            target_kunnr=destino_kunnr,
            target_pos=posicion_excel,
            target_contract=c_target,
            target_material=material,
            target_price=clean_num(precio_excel)
        )
        if scanned["encontrado"]:
            res.update(scanned)
        else:
            # 6. Fallback a exportación automática si no se detectó en pantalla
            exp_recs = export_and_read_z0116_excel(session)
            if exp_recs:
                return validate_z0116_contract(
                    session, plant=plant, vendor=vendor, material=material,
                    destino_kunnr=destino_kunnr, exp_cond_excel=exp_cond_excel,
                    posicion_excel=posicion_excel, contrato_excel=contrato_excel,
                    preloaded_records=exp_recs, precio_excel=precio_excel
                )
            else:
                res["detalle"] = f"Sin registros encontrados en pantalla ni exportación ({sbar_msg})"
                return res

        # 7. Regla estricta: Condición 04 exige KUNNR
        clean_exp = clean_exp_cond(exp_cond_excel)
        if clean_exp == "04":
            if res["kunnr_sap"] and res["kunnr_sap"].strip() != "":
                res["validacion_kunnr_04"] = f"OK (KUNNR={res['kunnr_sap']})"
            else:
                res["validacion_kunnr_04"] = "ERROR: KUNNR VACÍO EN CONDICIÓN 04"
        else:
            res["validacion_kunnr_04"] = "N/A (NO ES CONDICION 04)"

        return res

    except Exception as e:
        res["detalle"] = f"Excepción en Z0116: {str(e)}"
        return res


# ---------------------------------------------------------------------------
# Funciones de Soporte para Transacción ZSDD4501 (Contratos de Venta SD)
# ---------------------------------------------------------------------------
def fill_zsdd4501_selection(session, vkorg="7180", destino="", material="", centro="", contrato="", shipfrom="", dest_factura=""):
    """Rellena los campos de selección en la transacción ZSDD4501 (Contratos de Venta SD).
    
    Reglas oficiales especificadas por el usuario:
    1. Sector (SPART): Siempre va '04'. Dynpro: S_SPART-LOW
    2. Fecha fin de validez (GUEEN): Mayor a la fecha actual (desde hoy hasta 31.12.9999). Dynpro: S_GUEEN-LOW / HIGH
    3. Destinatario factura (KUNBP): Aquí va el SHIPFROM (Cantera/Proveedor/SF). Dynpro: S_KUNBP-LOW
    4. Destinatario de mercancía (SHPTO): Aquí va el DESTINO / CLIENTE. Dynpro: S_SHPTO-LOW
    5. Org. Ventas (VKORG), Centro (WERKS), Material (MATNR), Contrato (VBELN si aplica).
    """
    wnd0 = session.findById("wnd[0]")
    clear_screen_inputs(session)
    fch_hoy = datetime.now().strftime("%d.%m.%Y")

    # 1. Sector (SPART) siempre va 04
    for fid in ["wnd[0]/usr/ctxtS_SPART-LOW", "wnd[0]/usr/ctxtP_SPART", "wnd[0]/usr/txtS_SPART-LOW", "wnd[0]/usr/txtP_SPART"]:
        try:
            session.findById(fid).text = "04"
            break
        except Exception:
            pass

    # 2. Fecha Fin de Validez (GUEEN): Mayor a la fecha actual (desde hoy hasta 31.12.9999)
    for fid_low, fid_high in [
        ("wnd[0]/usr/ctxtS_GUEEN-LOW", "wnd[0]/usr/ctxtS_GUEEN-HIGH"),
        ("wnd[0]/usr/txtS_GUEEN-LOW", "wnd[0]/usr/txtS_GUEEN-HIGH")
    ]:
        try:
            session.findById(fid_low).text = fch_hoy
            session.findById(fid_high).text = "31.12.9999"
            break
        except Exception:
            pass

    for fid in ["wnd[0]/usr/ctxtP_GUEEN", "wnd[0]/usr/txtP_GUEEN"]:
        try:
            session.findById(fid).text = fch_hoy
            break
        except Exception:
            pass

    # 3. Destinatario Factura (S_KUNBP-LOW): Aquí va el SHIPFROM (Cantera/Proveedor/SF)
    factura_val = str(shipfrom or dest_factura or "").strip()
    set_dest_fact = False
    if factura_val:
        for fid in [
            "wnd[0]/usr/ctxtS_KUNBP-LOW", "wnd[0]/usr/txtS_KUNBP-LOW",
            "wnd[0]/usr/ctxtP_KUNBP", "wnd[0]/usr/txtP_KUNBP",
            "wnd[0]/usr/ctxtS_KUNRE-LOW", "wnd[0]/usr/txtS_KUNRE-LOW"
        ]:
            try:
                session.findById(fid).text = factura_val
                set_dest_fact = True
                break
            except Exception:
                pass
        for fid_high in ["wnd[0]/usr/ctxtS_KUNBP-HIGH", "wnd[0]/usr/txtS_KUNBP-HIGH", "wnd[0]/usr/ctxtS_KUNRE-HIGH"]:
            try:
                session.findById(fid_high).text = ""
            except Exception:
                pass
    else:
        for fid in [
            "wnd[0]/usr/ctxtS_KUNBP-LOW", "wnd[0]/usr/txtS_KUNBP-LOW",
            "wnd[0]/usr/ctxtP_KUNBP", "wnd[0]/usr/txtP_KUNBP"
        ]:
            try:
                session.findById(fid).text = ""
            except Exception:
                pass

    # 4. Destinatario de Mercancía (S_SHPTO-LOW): Aquí va el DESTINO / CLIENTE (ej. 67393641)
    dest_val = str(destino).strip()
    set_dest_merc = False
    if dest_val:
        for fid in [
            "wnd[0]/usr/ctxtS_SHPTO-LOW", "wnd[0]/usr/txtS_SHPTO-LOW",
            "wnd[0]/usr/ctxtP_SHPTO", "wnd[0]/usr/txtP_SHPTO",
            "wnd[0]/usr/ctxtS_KUNWE-LOW", "wnd[0]/usr/txtS_KUNWE-LOW",
            "wnd[0]/usr/ctxtP_KUNWE", "wnd[0]/usr/txtP_KUNWE",
            "wnd[0]/usr/ctxtS_DEST-LOW", "wnd[0]/usr/ctxtS_DESTINO-LOW",
            "wnd[0]/usr/ctxtS_OBRA-LOW", "wnd[0]/usr/ctxtS_SUCURSAL-LOW"
        ]:
            try:
                session.findById(fid).text = dest_val
                set_dest_merc = True
                break
            except Exception:
                pass
        for fid_high in ["wnd[0]/usr/ctxtS_SHPTO-HIGH", "wnd[0]/usr/txtS_SHPTO-HIGH", "wnd[0]/usr/ctxtS_KUNWE-HIGH"]:
            try:
                session.findById(fid_high).text = ""
            except Exception:
                pass
    else:
        for fid in [
            "wnd[0]/usr/ctxtS_SHPTO-LOW", "wnd[0]/usr/txtS_SHPTO-LOW",
            "wnd[0]/usr/ctxtP_SHPTO", "wnd[0]/usr/ctxtS_KUNWE-LOW",
            "wnd[0]/usr/ctxtS_DEST-LOW", "wnd[0]/usr/ctxtS_DESTINO-LOW"
        ]:
            try:
                session.findById(fid).text = ""
            except Exception:
                pass

    # 5. Organización de Ventas (VKORG)
    if vkorg:
        for fid in ["wnd[0]/usr/ctxtS_VKORG-LOW", "wnd[0]/usr/ctxtP_VKORG", "wnd[0]/usr/ctxtS_BUKRS-LOW"]:
            try:
                session.findById(fid).text = str(vkorg).strip()
                break
            except Exception:
                pass

    # 6. Material (MATNR)
    if material:
        mat_clean = str(material).strip()
        for fid in ["wnd[0]/usr/ctxtS_MATNR-LOW", "wnd[0]/usr/ctxtP_MATNR", "wnd[0]/usr/txtS_MATNR-LOW"]:
            try:
                session.findById(fid).text = mat_clean
                break
            except Exception:
                pass
    else:
        for fid in ["wnd[0]/usr/ctxtS_MATNR-LOW", "wnd[0]/usr/ctxtP_MATNR", "wnd[0]/usr/txtS_MATNR-LOW"]:
            try:
                session.findById(fid).text = ""
            except Exception:
                pass

    # 7. Centro (WERKS)
    if centro:
        centro_clean = str(centro).strip()
        for fid in ["wnd[0]/usr/ctxtS_WERKS-LOW", "wnd[0]/usr/ctxtP_WERKS", "wnd[0]/usr/txtS_WERKS-LOW"]:
            try:
                session.findById(fid).text = centro_clean
                break
            except Exception:
                pass

    # 8. Contrato SAP (VBELN) si viene indicado
    if contrato and not any(k in str(contrato).upper() for k in ["CREAR", "ALTA", "POR CREAR"]):
        cto_clean = str(contrato).strip()
        for fid in ["wnd[0]/usr/ctxtS_VBELN-LOW", "wnd[0]/usr/ctxtP_VBELN", "wnd[0]/usr/txtS_VBELN-LOW"]:
            try:
                session.findById(fid).text = cto_clean
                break
            except Exception:
                pass
    else:
        for fid in ["wnd[0]/usr/ctxtS_VBELN-LOW", "wnd[0]/usr/ctxtP_VBELN", "wnd[0]/usr/txtS_VBELN-LOW"]:
            try:
                session.findById(fid).text = ""
            except Exception:
                pass

    # 9. Clase de contrato (AUART = ZCP)
    for fid in ["wnd[0]/usr/ctxtS_AUART-LOW", "wnd[0]/usr/ctxtP_AUART"]:
        try:
            session.findById(fid).text = "ZCP"
            break
        except Exception:
            pass

    # 10. Búsqueda y llenado dinámico por inspección de controles
    try:
        textfields = find_controls_recursive(wnd0, "TextField")
        labels = find_controls_recursive(wnd0, "GuiLabel")

        # Localizar el campo de Destinatario de Factura (KUNBP) por etiqueta si no se llenó antes
        if factura_val and not set_dest_fact:
            for lbl in labels:
                ltxt = getattr(lbl, "Text", "").lower()
                if any(w in ltxt for w in ["factura", "fact.", "kunbp", "bill-to", "pagador"]):
                    top_l = getattr(lbl, "Top", -1)
                    for tf in textfields:
                        if abs(getattr(tf, "Top", -99) - top_l) < 5 and not getattr(tf, "ReadOnly", False):
                            try:
                                if not getattr(tf, "Text", "").strip():
                                    tf.Text = factura_val
                                    set_dest_fact = True
                                    break
                            except Exception:
                                pass
                    if set_dest_fact:
                        break

        # Localizar el campo de Destinatario de Mercancía (SHPTO) por etiqueta si no se llenó antes
        if dest_val and not set_dest_merc:
            for lbl in labels:
                ltxt = getattr(lbl, "Text", "").lower()
                if any(w in ltxt for w in ["mercanc", "merc.", "destinatario merc", "dest.merc", "shpto", "ship-to"]):
                    top_l = getattr(lbl, "Top", -1)
                    for tf in textfields:
                        if abs(getattr(tf, "Top", -99) - top_l) < 5 and not getattr(tf, "ReadOnly", False):
                            try:
                                if not getattr(tf, "Text", "").strip():
                                    tf.Text = dest_val
                                    set_dest_merc = True
                                    break
                            except Exception:
                                pass
                    if set_dest_merc:
                        break

        # Llenar otros campos por nombre/id si quedaron vacíos
        for ch in textfields:
            if getattr(ch, "ReadOnly", False):
                continue
            cid = getattr(ch, "Id", "").upper()
            cname = getattr(ch, "Name", "").upper()
            curr_val = getattr(ch, "Text", "").strip()

            if ("KUNBP" in cid or "KUNBP" in cname) and factura_val and not curr_val:
                if "HIGH" not in cid and "HIGH" not in cname:
                    try: ch.Text = factura_val
                    except Exception: pass
            elif ("SHPTO" in cid or "SHPTO" in cname) and dest_val and not curr_val:
                if "HIGH" not in cid and "HIGH" not in cname:
                    try: ch.Text = dest_val
                    except Exception: pass
            elif "SPART" in cid or "SPART" in cname:
                if not curr_val:
                    try: ch.Text = "04"
                    except Exception: pass
            elif "GUEEN" in cid or "GUEEN" in cname:
                if not curr_val:
                    if "HIGH" in cid or "HIGH" in cname:
                        try: ch.Text = "31.12.9999"
                        except Exception: pass
                    else:
                        try: ch.Text = fch_hoy
                        except Exception: pass
            elif not curr_val and ("LOW" in cid or "LOW" in cname):
                if any(k in cname or k in cid for k in ["VKORG", "ORG_VENTAS"]) and vkorg:
                    try: ch.Text = str(vkorg).strip()
                    except Exception: pass
                elif any(k in cname or k in cid for k in ["WERKS", "CENTRO"]) and centro:
                    try: ch.Text = str(centro).strip()
                    except Exception: pass
                elif any(k in cname or k in cid for k in ["MATNR", "MATERIAL"]) and material:
                    try: ch.Text = str(material).strip()
                    except Exception: pass
    except Exception:
        pass


def extract_um_via_detail_toolbar(session, grid, row_index):
    """
    Extrae la Unidad de Medida (UM) presionando el botón de la Lupa (&DETAIL) en la barra de herramientas del ALV Grid.
    Inspecciona la ventana modal Dynpro 0500 (estructura LVC_S_DETA, campo VALUE / elemento LVC_VALUE).
    Únicamente toma el valor real de SAP sin inventar ni usar fallbacks.
    """
    try:
        grid.selectedRows = str(row_index)
    except Exception:
        pass
    try:
        grid.setCurrentCell(row_index, "")
    except Exception:
        pass

    opened = False
    for btn_code in ["&DETAIL", "&ETA", "&OET"]:
        try:
            grid.pressToolbarButton(btn_code)
            time.sleep(0.4)
            if session.Children.Count > 1:
                opened = True
                break
        except Exception:
            pass

    um_extraida = ""
    if opened:
        try:
            wnd1 = session.findById("wnd[1]")
            grids_d = find_controls_recursive(wnd1, "GuiGridView")
            if not grids_d:
                grids_d = [ch for ch in find_controls_recursive(wnd1, "GuiShell") if hasattr(ch, "RowCount")]

            def _safe_cell(g, r, col):
                try:
                    v = g.GetCellValue(r, col)
                    return str(v or "").strip()
                except Exception:
                    return ""

            if grids_d:
                d_grid = grids_d[0]
                d_rows = getattr(d_grid, "RowCount", 0)

                # Obtener lista de columnas existentes en el grid de detalles de forma segura
                d_cols = []
                try:
                    c_names = d_grid.GetAllColumnNames()
                    try:
                        d_cols = [str(c_names(i)) for i in range(c_names.Count)]
                    except Exception:
                        d_cols = [str(c) for c in c_names]
                except Exception:
                    pass

                for d_r in range(d_rows):
                    try:
                        fname = ""
                        for c_fn in ["FIELDNAME", "FIELD"]:
                            v = _safe_cell(d_grid, d_r, c_fn)
                            if v:
                                fname = v.upper()
                                break

                        ftext = ""
                        for c_ft in ["FIELDTEXT", "REPTEXT", "SCRTEXT_L", "SCRTEXT_M", "COLTEXT", "SELTEXT"]:
                            v = _safe_cell(d_grid, d_r, c_ft)
                            if v:
                                ftext = v.upper()
                                break

                        # En Dynpro 0500 LVC_S_DETA el campo valor es VALUE
                        val = ""
                        for c_vl in ["VALUE", "VAL", "LVC_VALUE", "VALOR"]:
                            v = _safe_cell(d_grid, d_r, c_vl)
                            if v:
                                val = v.strip()
                                break

                        if not val and d_cols:
                            for c_name in d_cols:
                                if c_name.upper() in ("VALUE", "VAL", "LVC_VALUE", "VALOR") or "VAL" in c_name.upper():
                                    v = _safe_cell(d_grid, d_r, c_name)
                                    if v:
                                        val = v.strip()
                                        break

                        # Si es el campo de Unidad de Medida (VRKME, MEINS, ZIELME o texto descriptivo)
                        is_um_field = (
                            fname in ("VRKME", "MEINS", "ZIELME", "GEWEI", "VOLEH") or
                            any(k in ftext for k in ["UNIDAD DE MEDIDA", "UNIDAD MEDIDA", "UNIDAD VTA", "UNID.MEDIDA", "UNID. MEDIDA", "UNIDAD"])
                        )

                        if is_um_field and val:
                            c_val = clean_um(val)
                            if c_val and is_valid_um(c_val):
                                um_extraida = c_val
                                break
                    except Exception:
                        pass

            if not um_extraida:
                labels = find_controls_recursive(wnd1, "GuiLabel")
                textfields = find_controls_recursive(wnd1, "TextField")
                for lbl in labels:
                    ltxt = getattr(lbl, "Text", "").upper()
                    if any(k in ltxt for k in ["UNIDAD DE MEDIDA", "UNIDAD MEDIDA", "VRKME", "MEINS", "UNIDAD"]):
                        lbl_top = getattr(lbl, "Top", -1)
                        for tf in textfields:
                            if abs(getattr(tf, "Top", -99) - lbl_top) < 6:
                                c_val = clean_um(getattr(tf, "Text", ""))
                                if c_val and is_valid_um(c_val):
                                    um_extraida = c_val
                                    break
                        if um_extraida:
                            break
        finally:
            for close_cmd in [
                lambda: session.findById("wnd[1]/tbar[0]/btn[12]").press(),
                lambda: session.findById("wnd[1]").sendVKey(12),
                lambda: session.findById("wnd[1]/tbar[0]/btn[0]").press(),
                lambda: session.findById("wnd[1]").sendVKey(0)
            ]:
                try:
                    if session.Children.Count > 1:
                        close_cmd()
                        time.sleep(0.2)
                    else:
                        break
                except Exception:
                    pass

    return um_extraida


def extract_all_zsdd4501_from_screen(session):
    """Extrae todos los registros de contratos de venta desde la tabla VBAK / resultados de ZSDD4501.
    
    1. Búsqueda exhaustiva en ALV Grid (GuiGridView, GuiShell con SubType GridView o con RowCount).
    2. Mapeo dual por Technical Names (VRKME, POSNR, etc.) y por Column Titles (UM, Posición).
    3. Extracción robusta de Posición (descomponiendo Llave 'CONTRATO-POS') y UM (VRKME, MEINS, M3).
    4. Búsqueda en TableControl y Lista clásica ABAP (GuiLabel).
    5. Respaldo automático a exportación nativa de texto si no hay filas.
    """
    records = []
    try:
        wnd0 = session.findById("wnd[0]")
    except Exception:
        return records

    import unicodedata
    def norm_str(s):
        if not s: return ""
        s = unicodedata.normalize("NFKD", str(s)).encode("ASCII", "ignore").decode("utf-8")
        return s.strip().upper()

    # 1. BÚSQUEDA UNIVERSAL en ALV Grid
    def find_all_grids(parent):
        found = []
        try:
            for i in range(parent.Children.Count):
                ch = parent.Children(i)
                c_type = getattr(ch, "Type", "")
                c_subtype = getattr(ch, "SubType", "")
                if "gridview" in c_type.lower() or "gridview" in c_subtype.lower():
                    found.append(ch)
                elif "shell" in c_type.lower() and hasattr(ch, "RowCount"):
                    found.append(ch)
                if hasattr(ch, "Children") and ch.Children.Count > 0:
                    found.extend(find_all_grids(ch))
        except Exception:
            pass
        return found

    grids = find_all_grids(wnd0)
    if grids:
        for grid in grids:
            row_count = getattr(grid, "RowCount", 0)
            if row_count <= 0:
                continue
            cols = []
            try:
                co = grid.ColumnOrder
                cols = [str(co(i)) for i in range(co.Count)]
            except Exception:
                try: cols = list(grid.ColumnOrder)
                except Exception: pass

            # Mapear títulos de columnas si están disponibles
            col_titles_map = {}
            for c in cols:
                try:
                    ct_list = grid.GetColumnTitles(c)
                    if ct_list:
                        titles = []
                        try:
                            for idx_t in range(ct_list.Count):
                                titles.append(str(ct_list(idx_t)))
                        except Exception:
                            titles = [str(t) for t in ct_list]
                        for t in titles:
                            if t and t.strip():
                                col_titles_map[norm_str(t)] = c
                except Exception:
                    pass

            for r_i in range(row_count):
                r_dict = {}
                for c in cols:
                    try:
                        v = grid.GetCellValue(r_i, c)
                        if v is not None and str(v).strip() != "":
                            v_str = str(v).strip()
                            r_dict[norm_str(c)] = v_str
                            r_dict[c.upper()] = v_str
                    except Exception:
                        pass

                # Incorporar mapeo de títulos
                for title_norm, c_id in col_titles_map.items():
                    if norm_str(c_id) in r_dict:
                        r_dict[title_norm] = r_dict[norm_str(c_id)]

                def get_grid_val(keywords, exclude=None):
                    exclude = [norm_str(e) for e in (exclude or [])]
                    for k in keywords:
                        k_norm = norm_str(k)
                        for col_k, col_v in r_dict.items():
                            if any(e in col_k for e in exclude): continue
                            if k_norm == col_k and col_v:
                                return col_v
                    for k in keywords:
                        k_norm = norm_str(k)
                        for col_k, col_v in r_dict.items():
                            if any(e in col_k for e in exclude): continue
                            if k_norm in col_k and col_v:
                                return col_v
                    return ""

                llave = get_grid_val(["LLAVE", "KEY"])
                cto = get_grid_val(["CONTRATO", "DOC. VENTAS", "DOC VENTAS", "DOCUMENTO VENTAS", "VBELN", "DOC"], exclude=["TIPO DOCUMENTO", "TIPO"])
                pos_raw = get_grid_val(["POSICION", "POSNR", "POS", "POS.", "KPOSN", "ITEM", "PARTIDA", "POS.DOC", "POSICION DOC"], exclude=["CONDICIONES", "CONDICION", "TIPO", "PSTYV"])

                # Búsqueda en catálogo técnico completo (GetAllColumnNames) si la columna de Posición o Llave no estaba visible en el layout
                if not pos_raw or not pos_raw.isdigit():
                    try:
                        all_cols = []
                        try:
                            all_cols = [str(grid.GetAllColumnNames()(ci)) for ci in range(grid.GetAllColumnNames().Count)]
                        except Exception:
                            all_cols = list(grid.GetAllColumnNames())
                        for tc in ["POSNR", "KPOSN", "POS", "POSICION", "VBELP", "ITEM"]:
                            if tc in all_cols:
                                c_cand = str(grid.GetCellValue(r_i, tc) or "").strip()
                                if c_cand and c_cand.isdigit():
                                    pos_raw = c_cand
                                    break
                        if not pos_raw or not pos_raw.isdigit():
                            for tc in ["LLAVE", "KEY"]:
                                if tc in all_cols:
                                    c_cand = str(grid.GetCellValue(r_i, tc) or "").strip()
                                    if "-" in c_cand:
                                        parts = [p.strip() for p in c_cand.split("-")]
                                        if not cto and parts[0].isdigit():
                                            cto = parts[0]
                                        if len(parts) > 1 and parts[1].isdigit():
                                            pos_raw = parts[1]
                                            break
                    except Exception:
                        pass

                if (not pos_raw or not pos_raw.isdigit()) and llave and "-" in llave:
                    parts = [p.strip() for p in llave.split("-")]
                    if not cto and parts[0].isdigit():
                        cto = parts[0]
                    if len(parts) > 1 and parts[1].isdigit():
                        pos_raw = parts[1]

                if not pos_raw or not pos_raw.isdigit():
                    for cell_v in r_dict.values():
                        m_p = re.search(r"\b\d{8,10}-(\d+)\b", str(cell_v))
                        if m_p:
                            pos_raw = m_p.group(1)
                            break

                if not pos_raw or not pos_raw.isdigit():
                    for c_k, cell_v in r_dict.items():
                        s_v = str(cell_v).strip()
                        if any(k in norm_str(c_k) for k in ["POS", "ITM", "PARTIDA"]) and s_v.isdigit():
                            pos_raw = s_v
                            break

                pos_val = str(int(pos_raw)) if pos_raw and pos_raw.isdigit() else (pos_raw.strip() if pos_raw else "")

                dest_merc = get_grid_val([
                    "SHPTO", "KUNWE", "DESTINATARIO DE MERCANCIA", "DESTINATARIO DE MERCANCIAS",
                    "DEST.MCIA", "DEST. MCIA", "DEST.MC", "DESTINATARIO MERCANCIA",
                    "DESTINATARIO", "DESTINO", "SUCURSAL", "OBRA", "FRENTE"
                ], exclude=["SOLICITANTE", "PAGADOR", "FACTURA", "CLIENTE"]).lstrip("0")

                solicitante = get_grid_val([
                    "KUNNR", "SOLICITANTE", "CLIENTE", "CLIENTE SOLICITANTE"
                ], exclude=["DESTINATARIO", "MERCANCIA", "FACTURA", "SHPTO"]).lstrip("0")

                dest_fact = get_grid_val([
                    "KUNBP", "FACTURA", "PAGADOR", "DESTINATARIO DE FACTURA", "DEST.FACT"
                ], exclude=["DESTINATARIO MERCANCIA", "MERCANCIA", "SHPTO", "SOLICITANTE"]).lstrip("0")

                material = get_grid_val(["MATNR", "MATERIAL", "PRODUCTO"], exclude=["DESC"])
                centro = get_grid_val(["WERKS", "CENTRO", "PLANT"])
                sf = get_grid_val(["SHIPFROM", "SHIP FROM", "KUNBP", "CANTERA", "LIFNR", "PROVEEDOR"])
                p_net = clean_num(get_grid_val(["NETPR", "PRECIO", "PRICE", "NETWR", "KBETR", "IMPORTE"]))
                mon = get_grid_val(["WAERK", "WAERS", "MONEDA", "CURRENCY", "KONWA"]) or "MXN"

                # 1. Búsqueda estricta de UM en columnas visibles (excluyendo terminantemente columnas de documento/contrato)
                um_val = ""
                for c_k, c_v in r_dict.items():
                    if any(e in c_k for e in ["DOCUMENTO", "NUMERO", "DOC", "NUM", "CONTRATO", "PRECIO", "IMPORTE", "VALOR"]):
                        continue
                    if c_k in ("VRKME", "MEINS", "ZIELME", "GEWEI", "VOLEH") or re.search(r"\bU\.?M\.?\b", c_k) or "UNIDAD" in c_k:
                        c_cand = clean_um(c_v)
                        if c_cand and is_valid_um(c_cand):
                            um_val = c_cand
                            break

                # 2. Búsqueda directa en valores de celda visibles (ej. 'M3', 'TN')
                if not um_val:
                    for cell_v in r_dict.values():
                        s_u = str(cell_v).strip().upper()
                        if s_u in ("M3", "TN", "TO", "TON", "KG", "L", "LT", "M^3"):
                            c_cand = clean_um(s_u)
                            if c_cand and is_valid_um(c_cand):
                                um_val = c_cand
                                break

                # 3. Búsqueda en catálogo técnico completo (GetAllColumnNames) si la columna no estaba visible en el layout
                if not um_val:
                    try:
                        all_cols = []
                        try: all_cols = [str(grid.GetAllColumnNames()(ci)) for ci in range(grid.GetAllColumnNames().Count)]
                        except Exception: all_cols = list(grid.GetAllColumnNames())
                        for tc in ["VRKME", "MEINS", "ZIELME", "GEWEI"]:
                            if tc in all_cols:
                                c_cand = clean_um(grid.GetCellValue(r_i, tc))
                                if c_cand and is_valid_um(c_cand):
                                    um_val = c_cand
                                    break
                    except Exception:
                        pass

                # 4. Extracción mediante el botón de la Lupa (&DETAIL) si la UM no estuvo visible
                if not um_val and r_i < 5:
                    um_val = extract_um_via_detail_toolbar(session, grid, r_i)

                f_ini = get_grid_val(["GUEBG", "DATAB", "VPER_START", "FCH_INI", "INICIO", "FECHA INICIO", "FECHA INICIO VALIDEZ"])
                f_fin = get_grid_val(["GUEEN", "DATBI", "VPER_END", "FCH_FIN", "FIN", "FECHA FIN", "FECHA FIN VALIDEZ"])
                exp_c = clean_exp_cond(get_grid_val(["VSBED", "COND_EXP", "EXP_COND", "MODALIDAD", "COND. EXP.", "CONDICION EXP."]))

                dest_final = dest_merc or solicitante or dest_fact
                sf_final = sf or dest_fact

                if cto or material or dest_final:
                    records.append({
                        "contrato": cto,
                        "posicion": pos_val,
                        "destino": dest_final,
                        "cliente": solicitante or dest_final,
                        "dest_mercancia": dest_merc,
                        "solicitante": solicitante,
                        "shipfrom": sf_final,
                        "dest_factura": sf_final,
                        "vendor": sf_final,
                        "kunnr": dest_final,
                        "material": material,
                        "centro": centro,
                        "plant": centro,
                        "precio_neto": p_net,
                        "moneda": mon,
                        "um": um_val,
                        "fch_ini": f_ini,
                        "fch_fin": f_fin,
                        "exp_cond": exp_c,
                        "_grid_row": r_i,
                    })

            if records:
                return records

    # 2. BÚSQUEDA en TableControl
    table_controls = find_controls_recursive(wnd0, "TableControl")
    if table_controls:
        for tbl in table_controls:
            try:
                col_names = {}
                col_count = getattr(tbl, "Columns", []).Count if hasattr(tbl, "Columns") else 0
                for c_i in range(col_count):
                    try:
                        col_obj = tbl.Columns(c_i)
                        c_title = getattr(col_obj, "Title", "") or getattr(col_obj, "Name", "")
                        col_names[c_i] = norm_str(c_title)
                    except Exception:
                        pass

                cells_by_row = {}
                for ch_i in range(tbl.Children.Count):
                    cell = tbl.Children(ch_i)
                    cid = getattr(cell, "Id", "")
                    cname = norm_str(getattr(cell, "Name", ""))
                    ctext = getattr(cell, "Text", "").strip()

                    match = re.search(r"\[(\d+),(\d+)\]", cid)
                    if match:
                        col_idx = int(match.group(1))
                        row_idx = int(match.group(2))
                        cells_by_row.setdefault(row_idx, {})
                        field_key = col_names.get(col_idx, cname)
                        cells_by_row[row_idx][field_key] = ctext
                        if cname:
                            cells_by_row[row_idx][cname] = ctext

                for r_idx in sorted(cells_by_row.keys()):
                    r_data = cells_by_row[r_idx]

                    def get_tc_val(keywords, exclude=None):
                        exclude = [norm_str(e) for e in (exclude or [])]
                        for k in keywords:
                            k_norm = norm_str(k)
                            for col_k, col_v in r_data.items():
                                if any(e in col_k for e in exclude): continue
                                if k_norm in col_k and col_v:
                                    return col_v
                        return ""

                    llave = get_tc_val(["LLAVE", "KEY"])
                    cto = get_tc_val(["CONTRATO", "DOC. VENTAS", "DOC VENTAS", "VBELN", "DOC"], exclude=["TIPO"])
                    pos_val = get_tc_val(["POSICION", "POSNR", "POS", "POS."], exclude=["CONDICIONES"])

                    if (not pos_val or not pos_val.isdigit()) and llave and "-" in llave:
                        parts = [p.strip() for p in llave.split("-")]
                        if not cto and parts[0].isdigit(): cto = parts[0]
                        if len(parts) > 1 and parts[1].isdigit(): pos_val = parts[1]

                    if not pos_val or not pos_val.isdigit():
                        pos_val = ""

                    dest_merc = get_tc_val([
                        "SHPTO", "KUNWE", "DESTINATARIO DE MERCANCIA", "DESTINATARIO DE MERCANCIAS",
                        "DEST.MCIA", "DEST. MCIA", "DEST.MC", "DESTINATARIO MERCANCIA",
                        "DESTINATARIO", "DESTINO", "SUCURSAL", "OBRA", "FRENTE"
                    ], exclude=["SOLICITANTE", "PAGADOR", "FACTURA", "CLIENTE"]).lstrip("0")

                    solicitante = get_tc_val([
                        "KUNNR", "SOLICITANTE", "CLIENTE", "CLIENTE SOLICITANTE"
                    ], exclude=["DESTINATARIO", "MERCANCIA", "FACTURA", "SHPTO"]).lstrip("0")

                    dest_fact = get_tc_val(["KUNBP", "FACTURA", "DESTINATARIO FACTURA", "DEST.FACT", "PAGADOR"]).lstrip("0")
                    dest_final = dest_merc or solicitante or dest_fact

                    material = get_tc_val(["MATNR", "MATERIAL", "PRODUCTO"], exclude=["DESC"])
                    centro = get_tc_val(["WERKS", "CENTRO", "PLANT"])
                    sf = get_tc_val(["SHIPFROM", "SHIP FROM", "KUNBP", "LIFNR", "CANTERA"])
                    p_net = clean_num(get_tc_val(["NETPR", "PRECIO", "NETWR", "KBETR", "IMPORTE"]))
                    mon = get_tc_val(["WAERK", "WAERS", "MONEDA"]) or "MXN"
                    um_val = clean_um(get_tc_val(["VRKME", "MEINS", "ZIELME", "UNIDAD DE MEDIDA", "UNIDAD MEDIDA"], exclude=["DOCUMENTO", "NUMERO", "DOC", "NUM", "CONTRATO", "PRECIO", "IMPORTE"]))
                    exp_c = clean_exp_cond(get_tc_val(["VSBED", "COND_EXP", "EXP_COND", "MODALIDAD"]))
                    f_ini = get_tc_val(["GUEBG", "DATAB", "FCH_INI", "INICIO"])
                    f_fin = get_tc_val(["GUEEN", "DATBI", "FCH_FIN", "FIN"])

                    if cto or material or dest_final:
                        records.append({
                            "contrato": cto,
                            "posicion": pos_val,
                            "destino": dest_final,
                            "cliente": solicitante or dest_final,
                            "dest_mercancia": dest_merc,
                            "solicitante": solicitante,
                            "dest_factura": dest_fact,
                            "kunnr": dest_final,
                            "material": material,
                            "centro": centro,
                            "shipfrom": sf,
                            "precio_neto": p_net,
                            "moneda": mon,
                            "um": um_val,
                            "exp_cond": exp_c,
                            "fch_ini": f_ini,
                            "fch_fin": f_fin,
                        })

                if records:
                    return records
            except Exception:
                pass

    # 3. BÚSQUEDA en Lista Clásica ABAP (GuiLabel)
    labels = find_controls_recursive(wnd0, "GuiLabel")
    if labels:
        lines_dict = {}
        for l in labels:
            try:
                t = getattr(l, "Top", 0)
                txt = getattr(l, "Text", "").strip()
                if txt:
                    lines_dict.setdefault(t, []).append({
                        "text": txt,
                        "left": getattr(l, "Left", 0),
                        "width": getattr(l, "Width", 0),
                    })
            except Exception:
                continue

        for t in sorted(lines_dict.keys()):
            items = sorted(lines_dict[t], key=lambda x: x["left"])
            line_full = " ".join(x["text"] for x in items)

            row_data = {
                "contrato": "", "posicion": "", "cliente": "", "dest_mercancia": "",
                "dest_factura": "", "kunnr": "", "material": "", "centro": "",
                "shipfrom": "", "precio_neto": None, "moneda": "MXN", "um": "",
                "exp_cond": "", "fch_ini": "", "fch_fin": "", "line_text": line_full
            }

            for it in items:
                txt = it["text"].strip()
                if "-" in txt and re.match(r"^\d{8,10}-\d{1,4}$", txt):
                    c, po = txt.split("-")
                    row_data["contrato"] = c
                    row_data["posicion"] = po
                elif re.match(r"^2\d{8,9}$", txt) or re.match(r"^02\d{8,9}$", txt):
                    row_data["contrato"] = txt
                elif re.match(r"^10\d{6}$", txt):
                    row_data["material"] = txt
                elif re.match(r"^6[5-7]\d{6,7}$", txt):
                    row_data["cliente"] = txt.lstrip("0")
                    row_data["dest_mercancia"] = row_data["cliente"]
                    row_data["kunnr"] = row_data["cliente"]
                elif re.match(r"^D\d{3}$", txt):
                    row_data["centro"] = txt
                elif re.match(r"^900\d{6}$", txt):
                    row_data["shipfrom"] = txt
                elif txt in ("01", "02", "04") and not row_data["exp_cond"]:
                    row_data["exp_cond"] = clean_exp_cond(txt)
                elif txt.isdigit() and len(txt) in (2, 3) and not row_data["posicion"]:
                    row_data["posicion"] = str(int(txt))
                elif re.match(r"^\d+[\.,]\d{2}$", txt) and row_data["precio_neto"] is None:
                    row_data["precio_neto"] = clean_num(txt)
                elif txt in ("MXN", "USD"):
                    row_data["moneda"] = txt
                elif txt.upper() in ("TN", "M3", "KG", "TO"):
                    row_data["um"] = clean_um(txt)

            if not row_data["posicion"]:
                row_data["posicion"] = ""
            # if not row_data["um"]: row_data["um"] = ""

            if row_data["contrato"] or row_data["material"]:
                records.append(row_data)

    # 4. Fallback automático a exportación si la inspección de pantalla no obtuvo registros
    if not records:
        records = export_and_read_zsdd4501_excel(session)

    return records

def escribir_resultados_en_excel(ruta_excel_origen, resultados, ruta_salida=None):
    """Escribe las columnas de auditoría y colorea el archivo Excel según los resultados de validación SAP.
    
    Genera un archivo con sufijo _VALIDADO.xlsx con las columnas de auditoría:
    - Verde: Validado Correcto / Coincide
    - Rojo: Discrepancia detectada
    - Amarillo: Alta pendiente o no encontrado
    """
    if not os.path.exists(ruta_excel_origen):
        return None

    if not ruta_salida:
        directorio = os.path.dirname(ruta_excel_origen)
        base = os.path.splitext(os.path.basename(ruta_excel_origen))[0]
        ruta_salida = os.path.join(directorio, f"{base}_VALIDADO.xlsx")

    try:
        shutil.copy2(ruta_excel_origen, ruta_salida)
    except Exception:
        ts = datetime.now().strftime("%H%M%S")
        directorio = os.path.dirname(ruta_excel_origen)
        base = os.path.splitext(os.path.basename(ruta_excel_origen))[0]
        ruta_salida = os.path.join(directorio, f"{base}_VALIDADO_{ts}.xlsx")
        shutil.copy2(ruta_excel_origen, ruta_salida)

    wb = openpyxl.load_workbook(ruta_salida)

    header_font = Font(name="Calibri", size=10, bold=True, color="FFFFFF")
    header_fill = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
    header_align = Alignment(horizontal="center", vertical="center", wrap_text=True)

    fill_ok = PatternFill(start_color="C6EFCE", end_color="C6EFCE", fill_type="solid")
    font_ok = Font(color="006100", bold=True)
    fill_warn = PatternFill(start_color="FFF2CC", end_color="FFF2CC", fill_type="solid")
    font_warn = Font(color="B25900", bold=True)
    fill_err = PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid")
    font_err = Font(color="9C0006", bold=True)

    por_hoja = {}
    for item, res in resultados:
        hoja = item.get("hoja", "")
        if hoja not in por_hoja:
            por_hoja[hoja] = []
        por_hoja[hoja].append((item, res))

    for hoja, items_res in por_hoja.items():
        if hoja not in wb.sheetnames:
            continue
        ws = wb[hoja]
        tipo_seccion = items_res[0][0].get("tipo_seccion", "")

        # -------------------------------------------------------------
        # HOJA FPA (FILIAL / TERCEROS)
        # -------------------------------------------------------------
        if tipo_seccion == "FPA" or "FILIAL" in hoja.upper() or "TERCEROS" in hoja.upper():
            h_row = 5
            for test_r in (5, 4, 6):
                val_c = str(ws.cell(test_r, 1).value or "").upper()
                if "SOLICITUD" in val_c:
                    h_row = test_r
                    break

            col_start = 41
            headers_fpa = [
                "VK13_MAT_EXISTE", "VK13_MAT_IMPORTE", "VK13_MAT_UM", "VK13_MAT_MONEDA",
                "VK13_MAT_INICIO", "VK13_MAT_FIN", "VK13_MAT_VALIDACION",
                "VK13_FTE_EXISTE", "VK13_FTE_IMPORTE", "VK13_FTE_UM", "VK13_FTE_MONEDA",
                "VK13_FTE_MODALIDAD", "VK13_FTE_INICIO", "VK13_FTE_FIN", "VK13_FTE_VALIDACION",
                "ESTADO_GLOBAL_FILIAL", "DETALLE_AUDITORIA"
            ]
            for c_i, t in enumerate(headers_fpa):
                cell = ws.cell(row=h_row, column=col_start + c_i, value=t)
                cell.font = header_font
                cell.fill = header_fill
                cell.alignment = header_align

            for item, res in items_res:
                r = int(item.get("idx_fila", item.get("fila_origen", h_row + 1)))
                d_sap = res.get("datos_sap", {})
                estado = res.get("estado", "")

                mat_ex = "SI" if d_sap.get("material_encontrado") else ("N/A" if "SOLO_FLETES" in str(res.get("mensaje", "")) else "NO")
                fte_ex = "SI" if d_sap.get("flete_encontrado") else "NO"
                estado_txt = "VALIDADO CORRECTO" if estado == "OK" else ("DISCREPANCIA DETECTADA" if estado == "DISCREPANCIA" else "ALTA PENDIENTE")

                vals = [
                    mat_ex, d_sap.get("precio_mat_sap"), d_sap.get("um_mat_sap"), d_sap.get("moneda_mat_sap"),
                    d_sap.get("datab_mat_sap"), d_sap.get("datbi_mat_sap"), d_sap.get("status_mat", ""),
                    fte_ex, d_sap.get("precio_fte_sap"), d_sap.get("um_fte_sap"), d_sap.get("moneda_fte_sap"),
                    d_sap.get("modalidad_fte_sap"), d_sap.get("datab_fte_sap"), d_sap.get("datbi_fte_sap"), d_sap.get("status_fte", ""),
                    estado_txt, res.get("mensaje", "")
                ]

                row_fill = fill_ok if estado == "OK" else (fill_err if estado == "DISCREPANCIA" else fill_warn)
                row_font = font_ok if estado == "OK" else (font_err if estado == "DISCREPANCIA" else font_warn)

                for c_i, val in enumerate(vals):
                    c_cell = ws.cell(row=r, column=col_start + c_i, value=val)
                    c_cell.alignment = Alignment(horizontal="center")
                    if c_i == 15:
                        c_cell.fill = row_fill
                        c_cell.font = row_font

        # -------------------------------------------------------------
        # HOJA CTO COMPRA (Z0116_TRAOPE)
        # -------------------------------------------------------------
        elif tipo_seccion == "CTO_COMPRA" or "COMPRA" in hoja.upper():
            col_start = 35
            headers_cto = [
                "SAP_CONTRATO", "SAP_POSICION", "SAP_PRECIO_NETO", "SAP_MONEDA", "SAP_UM",
                "SAP_FCH_INI", "SAP_FCH_FIN", "SAP_EXP_COND", "SAP_KUNNR",
                "VALIDACION_KUNNR_04", "ESTADO_VALIDACION_CTO", "DETALLE_AUDITORIA"
            ]
            for c_i, t in enumerate(headers_cto):
                cell = ws.cell(row=1, column=col_start + c_i, value=t)
                cell.font = header_font
                cell.fill = header_fill
                cell.alignment = header_align

            for item, res in items_res:
                r = int(item.get("idx_fila", item.get("fila_origen", 2)))
                d_sap = res.get("datos_sap", {})
                estado = res.get("estado", "")
                msg_txt = res.get("mensaje", "")

                es_baja = ("BAJA" in str(item.get("solicitud", "")).upper() or 
                           "BAJA" in str(item.get("comentario", "")).upper() or 
                           "BAJA" in msg_txt.upper() or 
                           "dada de baja" in msg_txt.lower())

                if estado == "OK":
                    if es_baja:
                        estado_txt = "BAJA CORRECTA (SIN REGISTRO)"
                    else:
                        estado_txt = "CONTRATO VALIDADO OK"
                    row_fill = fill_ok
                    row_font = font_ok
                elif estado == "DISCREPANCIA":
                    estado_txt = "DISCREPANCIA / ERROR"
                    row_fill = fill_err
                    row_font = font_err
                else:
                    estado_txt = "NO ENCONTRADO EN SAP"
                    row_fill = fill_warn
                    row_font = font_warn

                vals = [
                    d_sap.get("contrato", "") or item.get("contrato", ""),
                    d_sap.get("posicion", "") or item.get("posicion", ""),
                    d_sap.get("precio", item.get("precio_mat")),
                    d_sap.get("moneda", "MXN"),
                    d_sap.get("um", "TN"),
                    d_sap.get("fch_ini", ""),
                    d_sap.get("fch_fin", ""),
                    d_sap.get("exp_cond", "") or item.get("modalidad", ""),
                    d_sap.get("destino", "") or d_sap.get("kunnr", "") or item.get("destino", ""),
                    d_sap.get("res_cto_raw", {}).get("validacion_kunnr_04", "N/A"),
                    estado_txt,
                    msg_txt
                ]

                for c_i, val in enumerate(vals):
                    c_cell = ws.cell(row=r, column=col_start + c_i, value=val)
                    c_cell.alignment = Alignment(horizontal="center")
                    if c_i == 10:
                        c_cell.fill = row_fill
                        c_cell.font = row_font

        # -------------------------------------------------------------
        # HOJA CTO VENTA (VA43 / ZCP)
        # -------------------------------------------------------------
        elif tipo_seccion == "CTO_VENTA" or "VENTA" in hoja.upper() or "VTA" in hoja.upper():
            # Buscar si ya existen columnas de auditoría previas para sobreescribir limpiamente
            col_start = None
            for c in range(1, ws.max_column + 1):
                c_val = str(ws.cell(1, c).value or "").upper()
                if "SAP_CONTRATO" in c_val:
                    col_start = c
                    break
            if not col_start:
                col_start = ws.max_column + 1

            headers_vta = [
                "SAP_CONTRATO_VENTA", "SAP_POSICION", "SAP_PRECIO_NETO", "SAP_MONEDA", "SAP_UM",
                "SAP_FCH_INI", "SAP_FCH_FIN", "SAP_EXP_COND", "SAP_DESTINO", "SAP_SHIPFROM",
                "SAP_MATERIAL", "SAP_CENTRO", "ESTADO_VALIDACION_VENTA", "DETALLE_AUDITORIA"
            ]
            for c_i, t in enumerate(headers_vta):
                cell = ws.cell(row=1, column=col_start + c_i, value=t)
                cell.font = header_font
                cell.fill = header_fill
                cell.alignment = header_align

            for item, res in items_res:
                r_rows = [int(item.get("idx_fila", 2))]
                if "-" in str(item.get("fila_origen", "")):
                    try:
                        r_rows.append(int(item.get("fila_origen").split("-")[1]))
                    except Exception:
                        pass

                d_sap = res.get("datos_sap", {})
                estado = res.get("estado", "")
                msg_txt = res.get("mensaje", "")
                es_baja = ("BAJA" in str(item.get("solicitud", "")).upper() or
                           "BAJA" in str(item.get("comentario", "")).upper() or
                           "BAJA" in msg_txt.upper() or "dada de baja" in msg_txt.lower())

                if estado == "OK":
                    if es_baja:
                        estado_txt = "BAJA CORRECTA (SIN REGISTRO)"
                    elif "ALTA" in str(item.get("solicitud", "")).upper() and not d_sap.get("contrato"):
                        estado_txt = "ALTA PENDIENTE"
                    else:
                        estado_txt = "CONTRATO VALIDADO OK"
                    row_fill = fill_ok
                    row_font = font_ok
                elif estado == "DISCREPANCIA":
                    estado_txt = "DISCREPANCIA / ERROR"
                    row_fill = fill_err
                    row_font = font_err
                else:
                    estado_txt = "NO LOCALIZADO EN SAP"
                    row_fill = fill_warn
                    row_font = font_warn

                pos_out = d_sap.get("posicion") or item.get("posicion") or "10"
                um_out = d_sap.get("um") or item.get("um") or "M3"

                vals = [
                    d_sap.get("contrato", "") or item.get("contrato", ""),
                    pos_out,
                    d_sap.get("precio", item.get("precio_mat")),
                    d_sap.get("moneda", item.get("moneda", "MXN")),
                    um_out,
                    d_sap.get("fch_ini", item.get("fch_ini", "")),
                    d_sap.get("fch_fin", item.get("fch_fin", "")),
                    d_sap.get("exp_cond", "") or item.get("modalidad", ""),
                    d_sap.get("destino", "") or item.get("destino", ""),
                    d_sap.get("shipfrom", "") or item.get("sf", ""),
                    d_sap.get("material", "") or item.get("material", ""),
                    d_sap.get("centro", "") or item.get("cedis", ""),
                    estado_txt,
                    msg_txt
                ]

                for r_target in set(r_rows):
                    for c_i, val in enumerate(vals):
                        c_cell = ws.cell(row=r_target, column=col_start + c_i, value=val)
                        c_cell.alignment = Alignment(horizontal="center")
                        if c_i == len(vals) - 2:  # ESTADO_VALIDACION_VENTA
                            c_cell.fill = row_fill
                            c_cell.font = row_font
                        elif c_i == len(vals) - 1 and estado == "DISCREPANCIA":
                            c_cell.font = font_err

        # -------------------------------------------------------------
        # HOJA MATRIZ PRECIOS (MATRIZ_CONSOLIDADA)
        # -------------------------------------------------------------
        elif tipo_seccion == "MATRIZ_CONSOLIDADA" or "MATRIZ" in hoja.upper():
            h_row2 = [str(c.value or "").strip().upper() for c in ws[2]] if ws.max_row >= 2 else []
            has_cond_fte = any("FLETE" in h and ("COND" in h or "CLASE" in h) for h in h_row2)
            col_start = 37 if (has_cond_fte or (ws.max_column and ws.max_column >= 36)) else 36
            # 1. Cabeceras de Grupo en Fila 1
            groups_new = [
                (col_start, col_start + 3, "AUDITORÍA VK13 (VENTA)"),
                (col_start + 4, col_start + 8, "AUDITORÍA CONTRATO COMPRA (TRAOPE)"),
                (col_start + 9, col_start + 13, "AUDITORÍA CONTRATO VENTA (ZCP)"),
                (col_start + 14, col_start + 15, "DICTAMEN GLOBAL DE GOBERNANZA")
            ]
            for c_s, c_e, g_title in groups_new:
                for c_col in range(c_s, c_e + 1):
                    c_cell = ws.cell(row=1, column=c_col, value=g_title)
                    c_cell.font = header_font
                    c_cell.fill = header_fill
                    c_cell.alignment = header_align

            # 2. Subcabeceras en Fila 2
            subheaders_mat = [
                (col_start, "SAP_VK13_PRECIO_MAT"), (col_start + 1, "SAP_VK13_PRECIO_FTE"), (col_start + 2, "SAP_VK13_ESTADO"), (col_start + 3, "SAP_VK13_DETALLE"),
                (col_start + 4, "SAP_COMPRA_CONTRATO"), (col_start + 5, "SAP_COMPRA_POS"), (col_start + 6, "SAP_COMPRA_PRECIO"), (col_start + 7, "SAP_COMPRA_ESTADO"), (col_start + 8, "SAP_COMPRA_DETALLE"),
                (col_start + 9, "SAP_VENTA_CONTRATO"), (col_start + 10, "SAP_VENTA_POS"), (col_start + 11, "SAP_VENTA_PRECIO"), (col_start + 12, "SAP_VENTA_ESTADO"), (col_start + 13, "SAP_VENTA_DETALLE"),
                (col_start + 14, "ESTADO_GLOBAL_FILA"), (col_start + 15, "RESUMEN_DISCREPANCIAS")
            ]
            sub_fill = PatternFill(start_color="2F5597", end_color="2F5597", fill_type="solid")
            for col_idx, s_title in subheaders_mat:
                c_cell = ws.cell(row=2, column=col_idx, value=s_title)
                c_cell.font = header_font
                c_cell.fill = sub_fill
                c_cell.alignment = header_align

            # 3. Filas de Resultados
            for item, res in items_res:
                r = int(item.get("fila_origen", 3))
                d_sap = res.get("datos_sap", {})
                vk_sap = d_sap.get("vk13", {})
                cp_sap = d_sap.get("compra", {})
                vt_sap = d_sap.get("venta", {})
                res_vk = res.get("res_vk13", {})
                res_cp = res.get("res_compra", {})
                res_vt = res.get("res_venta", {})

                est_global = res.get("estado", "")
                if est_global == "OK":
                    row_fill = fill_ok
                    row_font = font_ok
                elif est_global == "DISCREPANCIA":
                    row_fill = fill_err
                    row_font = font_err
                else:
                    row_fill = fill_warn
                    row_font = font_warn

                row_vals = [
                    vk_sap.get("precio_mat_sap"),
                    vk_sap.get("precio_fte_sap"),
                    res_vk.get("estado", ""),
                    res_vk.get("mensaje", ""),
                    cp_sap.get("contrato", ""),
                    cp_sap.get("posicion", ""),
                    cp_sap.get("precio"),
                    res_cp.get("estado", ""),
                    res_cp.get("mensaje", ""),
                    vt_sap.get("contrato", ""),
                    vt_sap.get("posicion", ""),
                    vt_sap.get("precio"),
                    res_vt.get("estado", ""),
                    res_vt.get("mensaje", ""),
                    est_global,
                    res.get("mensaje", "")
                ]

                for c_i, v in enumerate(row_vals):
                    cell = ws.cell(row=r, column=col_start + c_i, value=v)
                    cell.alignment = Alignment(horizontal="center")
                    if c_i == 14: # ESTADO_GLOBAL_FILA
                        cell.fill = row_fill
                        cell.font = row_font
                    elif c_i == 15 and est_global == "DISCREPANCIA":
                        cell.font = font_err

    wb.save(ruta_salida)
    wb.close()
    return ruta_salida


# ---------------------------------------------------------------------------
# Detección Automática de Formatos Excel
# ---------------------------------------------------------------------------
def detectar_tipo_archivo(ruta_excel):
    """Analiza todas las hojas del archivo Excel para detectar todos los formatos presentes.

    Soporta archivos con una sola hoja o múltiples hojas combinadas:
      - FPA FILIAL (VK13 ZMAH / ZMPH)
      - FPA TERCEROS (VK13 ZMA6 / ZMP1)
      - CTO COMPRA (Z0116_TRAOPE / ME31K / ME33K)
      - CTO VENTA (VA43 / VA45 / ZCP)

    Retorna un diccionario con:
      - tipo: 'MULTI_HOJA' si hay más de 1 hoja relevante, o el tipo específico
      - hoja: nombre de la primera hoja detectada
      - subtipo: 'COMPLETO' o 'SOLO_FLETES'
      - descripcion: descripción legible combinada
      - hojas_detectadas: lista de dicts con info de cada hoja
    """
    if not os.path.exists(ruta_excel):
        raise FileNotFoundError(f"Archivo no encontrado: {ruta_excel}")

    wb = openpyxl.load_workbook(ruta_excel, read_only=True, data_only=True)
    sheet_names = wb.sheetnames
    wb.close()

    nombre_archivo = os.path.basename(ruta_excel).lower()
    hojas_detectadas = []

    # 1. Búsqueda exhaustiva por nombres de hojas
    for s in sheet_names:
        s_low = s.strip().lower()
        if "fpa filial" in s_low or "filial" in s_low:
            subtipo = _analizar_subtipo_fpa(ruta_excel, s)
            desc_sub = "Sólo Fletes" if subtipo == "SOLO_FLETES" else "Material + Fletes"
            hojas_detectadas.append({
                "tipo": "FPA_FILIAL",
                "hoja": s,
                "subtipo": subtipo,
                "descripcion": f"FPA Filial ({desc_sub})"
            })
        elif "fpa terceros" in s_low or "terceros" in s_low:
            hojas_detectadas.append({
                "tipo": "FPA_TERCEROS",
                "hoja": s,
                "subtipo": "COMPLETO",
                "descripcion": f"FPA Terceros ({s})"
            })
        elif "cto compra" in s_low or "contrato de compra" in s_low or "compra" in s_low or "traope" in s_low or "z0116" in s_low:
            hojas_detectadas.append({
                "tipo": "CTO_COMPRA",
                "hoja": s,
                "subtipo": "COMPLETO",
                "descripcion": f"Contrato de Compra ({s})"
            })
        elif "cto vta" in s_low or "cto venta" in s_low or "contrato de venta" in s_low or "venta" in s_low:
            hojas_detectadas.append({
                "tipo": "CTO_VENTA",
                "hoja": s,
                "subtipo": "COMPLETO",
                "descripcion": f"Contrato de Venta ({s})"
            })
        elif ("matriz" in s_low or "consolidad" in s_low) and "dashboard" not in s_low:
            hojas_detectadas.append({
                "tipo": "MATRIZ_CONSOLIDADA",
                "hoja": s,
                "subtipo": "COMPLETO",
                "descripcion": f"Matriz Precios Consolidada ({s})"
            })

    # 2. Fallbacks si las hojas tienen nombres genéricos
    if not hojas_detectadas:
        if "fpa" in nombre_archivo and "filial" in nombre_archivo:
            hoja = sheet_names[0]
            subtipo = _analizar_subtipo_fpa(ruta_excel, hoja)
            hojas_detectadas.append({
                "tipo": "FPA_FILIAL",
                "hoja": hoja,
                "subtipo": subtipo,
                "descripcion": f"FPA Filial ({'Sólo Fletes' if subtipo == 'SOLO_FLETES' else 'Material + Fletes'})"
            })
        elif "compra" in nombre_archivo or "cto" in nombre_archivo:
            hojas_detectadas.append({
                "tipo": "CTO_COMPRA",
                "hoja": sheet_names[0],
                "subtipo": "COMPLETO",
                "descripcion": "Contrato de Compra (Z0116_TRAOPE)"
            })
        elif "matriz" in nombre_archivo:
            hojas_detectadas.append({
                "tipo": "MATRIZ_CONSOLIDADA",
                "hoja": sheet_names[0],
                "subtipo": "COMPLETO",
                "descripcion": "Matriz Precios Consolidada"
            })
        else:
            hojas_detectadas.append({
                "tipo": "DESCONOCIDO",
                "hoja": sheet_names[0] if sheet_names else "",
                "subtipo": "COMPLETO",
                "descripcion": "Formato no identificado automáticamente"
            })

    desc_total = " + ".join(h["descripcion"] for h in hojas_detectadas)
    es_multi = len(hojas_detectadas) > 1

    return {
        "tipo": "MULTI_HOJA" if es_multi else hojas_detectadas[0]["tipo"],
        "hoja": hojas_detectadas[0]["hoja"],
        "subtipo": hojas_detectadas[0].get("subtipo", "COMPLETO"),
        "descripcion": desc_total,
        "hojas_detectadas": hojas_detectadas
    }


def _analizar_subtipo_fpa(ruta_excel, hoja):
    """Detecta si un archivo FPA contiene sólo fletes o material + fletes."""
    try:
        wb = openpyxl.load_workbook(ruta_excel, read_only=True, data_only=True)
        ws = wb[hoja]
        for row in ws.iter_rows(max_row=10, values_only=True):
            headers = [str(c).upper().strip() for c in row if c is not None]
            if any("FLETE" in h for h in headers) and not any("PRECIO MATERIAL" in h for h in headers):
                wb.close()
                return "SOLO_FLETES"
        wb.close()
    except Exception:
        pass
    return "COMPLETO"


# ---------------------------------------------------------------------------
# Extracción de Datos Normalizados desde el Excel (Multi-Hoja)
# ---------------------------------------------------------------------------
def extraer_filas_archivo(ruta_excel, info_tipo=None):
    """Extrae y normaliza las filas de todas las hojas identificadas en el Excel.

    Soporta extracción conjunta de:
      - FPA FILIAL / FPA TERCEROS
      - CTO COMPRA (Z0116_TRAOPE)
      - CTO VENTA (ZCP - Doble fila Cabecera H / Posición I)
    """
    if info_tipo is None:
        info_tipo = detectar_tipo_archivo(ruta_excel)

    hojas_a_procesar = info_tipo.get("hojas_detectadas", [info_tipo])
    wb = openpyxl.load_workbook(ruta_excel, data_only=True)
    items = []
    detected_vkorgs_by_cedis = {}
    default_detected_vkorg = ""

    for h_info in hojas_a_procesar:
        hoja = h_info["hoja"]
        tipo = h_info["tipo"]
        if hoja not in wb.sheetnames:
            continue
        ws = wb[hoja]

        # -------------------------------------------------------------------
        # 1. FPA FILIAL / FPA TERCEROS
        # -------------------------------------------------------------------
        if tipo in ("FPA_FILIAL", "FPA_TERCEROS"):
            # Buscar dinámicamente la fila de encabezados (contiene SOLICITUD y CENTRO)
            h_row = None
            for r in range(1, min(ws.max_row + 1, 12)):
                txts = [str(ws.cell(r, c).value or "").upper() for c in range(1, min(ws.max_column + 1, 25))]
                if any("SOLICITUD" in t for t in txts) and any("CENTRO" in t for t in txts):
                    h_row = r
                    break
            if not h_row:
                h_row = 5

            col_map = {}
            for c in range(1, ws.max_column + 1):
                v_head = str(ws.cell(h_row, c).value or "").strip().upper()
                v_prev = str(ws.cell(h_row - 1, c).value or "").strip().upper() if h_row > 1 else ""
                combo = f"{v_prev} {v_head}"

                if "SOLICITUD" in combo and "solicitud" not in col_map:
                    col_map["solicitud"] = c
                elif any(k in combo for k in ["ORG.VENTAS", "ORG. VENTAS", "ORG_VENTAS"]) and "vkorg" not in col_map:
                    col_map["vkorg"] = c
                elif ("SHIPFROM" in combo or "SHIP FROM" in combo) and "DESCRIP" not in combo and "sf" not in col_map:
                    col_map["sf"] = c
                elif "CENTRO" in combo and "DESCRIP" not in combo and "cedis" not in col_map:
                    col_map["cedis"] = c
                elif any(k in combo for k in ["DESTINO", "DEST"]) and ("NO." in combo or "SAP" in combo or "KUNNR" in combo) and "DESCRIP" not in combo and "destino" not in col_map:
                    col_map["destino"] = c
                elif "MATERIAL" in combo and "DESCRIP" not in combo and "PRECIO" not in combo and "COND" not in combo and "material" not in col_map:
                    col_map["material"] = c
                elif "COND" in combo and "MATERIAL" in combo and "kschl_mat" not in col_map:
                    col_map["kschl_mat"] = c
                elif ("PRECIO MATERIAL" in combo or ("PRECIO" in combo and "MATERIAL" in combo and "COND" not in combo)) and "precio_mat" not in col_map:
                    col_map["precio_mat"] = c
                elif "COND" in combo and "FLETE" in combo and "kschl_fte" not in col_map:
                    col_map["kschl_fte"] = c
                elif ("PRECIO FLETE" in combo or ("PRECIO" in combo and "FLETE" in combo and "COND" not in combo)) and "precio_flete" not in col_map:
                    col_map["precio_flete"] = c
                elif any(k in combo for k in ["MODALIDAD", "EXPEDICI"]) and "modalidad" not in col_map:
                    col_map["modalidad"] = c
                elif "UM" in combo and "um" not in col_map:
                    col_map["um"] = c
                elif "PV" in combo and "pv" not in col_map:
                    col_map["pv"] = c

            # Fallbacks estándar
            if "solicitud" not in col_map: col_map["solicitud"] = 1
            if "vkorg" not in col_map: col_map["vkorg"] = 6
            if "sf" not in col_map: col_map["sf"] = 7
            if "cedis" not in col_map: col_map["cedis"] = 9
            if "destino" not in col_map: col_map["destino"] = 12
            if "material" not in col_map: col_map["material"] = 14
            if "kschl_mat" not in col_map: col_map["kschl_mat"] = 16
            if "precio_mat" not in col_map: col_map["precio_mat"] = 17
            if "um" not in col_map: col_map["um"] = 18
            if "kschl_fte" not in col_map: col_map["kschl_fte"] = 19
            if "precio_flete" not in col_map: col_map["precio_flete"] = 20
            if "modalidad" not in col_map: col_map["modalidad"] = 21

            for r in range(h_row + 1, ws.max_row + 1):
                def _val(col_name):
                    c_idx = col_map.get(col_name)
                    if not c_idx:
                        return ""
                    v = ws.cell(r, c_idx).value
                    return str(v).strip() if v is not None else ""

                solicitud = _val("solicitud").upper()
                vkorg = _val("vkorg") or "7100"
                sf = _val("sf")
                cedis = _val("cedis")
                if vkorg:
                    default_detected_vkorg = vkorg.strip()
                    if cedis:
                        detected_vkorgs_by_cedis[cedis.strip().upper()] = vkorg.strip()
                destino = _val("destino")
                material = _val("material")
                modalidad = clean_exp_cond(_val("modalidad"))
                kschl_mat = _val("kschl_mat") or ("ZMA6" if tipo == "FPA_TERCEROS" else "ZMAH")
                kschl_fte = _val("kschl_fte") or ("ZMP1" if tipo == "FPA_TERCEROS" else "ZMPH")
                precio_mat = clean_num(ws.cell(r, col_map["precio_mat"]).value if "precio_mat" in col_map else None)
                precio_flete = clean_num(ws.cell(r, col_map["precio_flete"]).value if "precio_flete" in col_map else None)
                um = clean_um(_val("um"))
                pv = clean_num(ws.cell(r, col_map["pv"]).value if "pv" in col_map else None)

                if not sf and not cedis and not material:
                    continue
                if any(k in sf.upper() for k in ["SHIP", "ESQUEMA", "SOLICITUD"]):
                    continue

                clave_str = f"[{hoja}] Org {vkorg} | Centro {cedis} | SF {sf} | Destino {destino} | Mat {material} | Mod {modalidad or '01'}"
                precios_desc = []
                if precio_mat is not None:
                    precios_desc.append(f"${precio_mat:,.2f} Mat")
                if precio_flete is not None:
                    precios_desc.append(f"${precio_flete:,.2f} Flete")
                precio_str = " / ".join(precios_desc) if precios_desc else "-"

                items.append({
                    "idx_fila": r,
                    "fila_origen": str(r),
                    "hoja": hoja,
                    "tipo_seccion": "FPA",
                    "tipo_fpa": tipo,
                    "subtipo": h_info.get("subtipo", "COMPLETO"),
                    "solicitud": solicitud,
                    "vkorg": vkorg,
                    "sf": sf,
                    "cedis": cedis,
                    "destino": destino,
                    "material": material,
                    "modalidad": modalidad or "01",
                    "kschl_mat": kschl_mat,
                    "kschl_fte": kschl_fte,
                    "precio_mat": precio_mat,
                    "precio_flete": precio_flete,
                    "contrato": "",
                    "posicion": "",
                    "proveedor": "",
                    "um": um,
                    "pv": pv,
                    "clave_str": clave_str,
                    "precio_str": precio_str,
                    "raw_dict": {
                        "fila": r,
                        "solicitud": solicitud,
                        "vkorg": vkorg,
                        "sf": sf,
                        "cedis": cedis,
                        "destino": destino,
                        "planta_obra": destino,
                        "material": material,
                        "articulo": material,
                        "modalidad": modalidad or "01",
                        "precio_mat": precio_mat,
                        "precio_flete": precio_flete,
                        "um": um,
                        "pv": pv,
                    },
                })

        # -------------------------------------------------------------------
        # 2. CTO COMPRA (Z0116_TRAOPE)
        # -------------------------------------------------------------------
        elif tipo == "CTO_COMPRA":
            col_c = {}
            for c in range(1, ws.max_column + 1):
                h = str(ws.cell(1, c).value or "").strip().upper()
                if not h:
                    continue
                if "CONTRATO" in h and "contrato" not in col_c: col_c["contrato"] = c
                elif "PROVEEDOR" in h and "proveedor" not in col_c: col_c["proveedor"] = c
                elif any(k in h for k in ["VENDOR", "PROV", "LIFNR"]) and "vendor" not in col_c: col_c["vendor"] = c
                elif any(k in h for k in ["PLANT", "CENTRO", "WERK"]) and "plant" not in col_c: col_c["plant"] = c
                elif "MATERIAL" in h and "material" not in col_c: col_c["material"] = c
                elif "PRICE" in h and "price" not in col_c: col_c["price"] = c
                elif any(k in h for k in ["EXP_COND", "MODALIDAD", "VSBED"]) and "exp_cond" not in col_c: col_c["exp_cond"] = c
                elif any(k in h for k in ["KUNNR", "DESTINO"]) and "kunnr" not in col_c: col_c["kunnr"] = c
                elif any(k in h for k in ["POSICION", "POS", "EBELP"]) and "pos" not in col_c: col_c["pos"] = c
                elif "COMPANY" in h and "company" not in col_c: col_c["company"] = c
                elif any(k in h for k in ["COMENTARIO", "OBSERVACION", "STATUS", "SOLICITUD"]) and "comentario" not in col_c: col_c["comentario"] = c

            if "contrato" not in col_c: col_c["contrato"] = 2
            if "vendor" not in col_c: col_c["vendor"] = 3
            if "plant" not in col_c: col_c["plant"] = 10
            if "material" not in col_c: col_c["material"] = 12
            if "price" not in col_c: col_c["price"] = 15
            if "exp_cond" not in col_c: col_c["exp_cond"] = 23
            if "kunnr" not in col_c: col_c["kunnr"] = 24
            if "comentario" not in col_c and ws.max_column >= 33: col_c["comentario"] = 33
            if "pos" not in col_c: col_c["pos"] = 34

            for r in range(2, ws.max_row + 1):
                def _cv(name):
                    ci = col_c.get(name)
                    if not ci: return ""
                    v = ws.cell(r, ci).value
                    return str(v).strip() if v is not None else ""

                contrato = _cv("contrato")
                vendor = _cv("vendor")
                plant = _cv("plant")
                material = _cv("material")
                price_net = clean_num(ws.cell(r, col_c["price"]).value if "price" in col_c else None)
                exp_cond = clean_exp_cond(_cv("exp_cond"))
                kunnr = _cv("kunnr")
                pos = _cv("pos")
                comentario = _cv("comentario")
                es_baja = "BAJA" in comentario.upper()
                solicitud = "BAJA" if es_baja else ("ALTA" if not contrato else "VALIDAR")
                prov_col = _cv("proveedor")
                proveedor_final = prov_col if prov_col else vendor

                # Extracción de COMPANY (Sociedad / Org. Ventas en CEMEX)
                company = _cv("company")
                if company:
                    try:
                        company = str(int(float(company)))
                    except Exception:
                        company = str(company).strip()

                plant_clean = plant.strip().upper() if plant else ""
                vkorg_compra = company or detected_vkorgs_by_cedis.get(plant_clean, "") or default_detected_vkorg
                if not company and vkorg_compra:
                    company = vkorg_compra

                if not contrato and not vendor and not plant:
                    continue
                if any(k in contrato.upper() for k in ["CONTRATO", "DOCUMENTO"]):
                    continue

                c_label = contrato if contrato else "(Por crear)"
                clave_str = f"[{hoja}] Cto {c_label} | Prov {vendor} | Centro {plant} | Dest {kunnr} | Mat {material}"
                precio_str = f"${price_net:,.2f} Neto" if price_net is not None else "-"

                items.append({
                    "idx_fila": r,
                    "fila_origen": str(r),
                    "hoja": hoja,
                    "tipo_seccion": "CTO_COMPRA",
                    "solicitud": solicitud,
                    "comentario": comentario,
                    "sf": vendor,
                    "cedis": plant,
                    "destino": kunnr,
                    "material": material,
                    "modalidad": exp_cond or "01",
                    "precio_mat": price_net,
                    "precio_flete": None,
                    "contrato": contrato,
                    "posicion": pos,
                    "proveedor": proveedor_final,
                    "company": company,
                    "vkorg": vkorg_compra,
                    "org_ventas": vkorg_compra,
                    "sociedad": company,
                    "um": "",
                    "pv": None,
                    "clave_str": clave_str,
                    "precio_str": precio_str,
                    "raw_dict": {
                        "fila": r,
                        "sf": vendor,
                        "cedis": plant,
                        "destino": kunnr,
                        "planta_obra": kunnr,
                        "material": material,
                        "articulo": material,
                        "modalidad": exp_cond or "01",
                        "precio_mat": price_net,
                        "comentario": comentario,
                        "company": company,
                        "vkorg": vkorg_compra,
                        "org_ventas": vkorg_compra,
                        "sociedad": company,
                    },
                })

        # -------------------------------------------------------------------
        # 3. CTO VENTA (VA43 / ZCP)
        # -------------------------------------------------------------------
        # -------------------------------------------------------------------
        # 3. CTO VENTA (VA43 / ZCP / ZSDD4501)
        # -------------------------------------------------------------------
        elif tipo == "CTO_VENTA":
            # Mapeo dinámico de encabezados en fila 1 con fallbacks
            col_v = {}
            for c in range(1, ws.max_column + 1):
                txt = str(ws.cell(1, c).value or "").strip().upper()
                if not txt: continue
                if "CONTRATO" in txt and "contrato" not in col_v: col_v["contrato"] = c
                elif "CLASE" in txt and "clase" not in col_v: col_v["clase"] = c
                elif "ORG" in txt and "vkorg" not in col_v: col_v["vkorg"] = c
                elif ("SUCURSAL" in txt or "OBRA" in txt or "DESTINO" in txt) and "destino" not in col_v: col_v["destino"] = c
                elif any(k in txt for k in ["CLIENTE NOMBRE", "CLIENTE", "FACTURA", "PAGADOR"]) and "dest_fact" not in col_v: col_v["dest_fact"] = c
                elif "INICIO" in txt and "fch_ini" not in col_v: col_v["fch_ini"] = c
                elif "FIN" in txt and "fch_fin" not in col_v: col_v["fch_fin"] = c
                elif any(k in txt for k in ["COND EXP", "EXP", "MODALIDAD"]) and "modalidad" not in col_v: col_v["modalidad"] = c
                elif any(k in txt for k in ["PRECIO", "VALOR", "NETO", "IMPORTE"]) and "precio" not in col_v: col_v["precio"] = c
                elif any(k in txt for k in ["MONEDA", "WAERS", "CURR"]) and "moneda" not in col_v: col_v["moneda"] = c
                elif ("PRODUCTO" in txt or "MATERIAL" in txt or "ARTICULO" in txt) and "material" not in col_v: col_v["material"] = c
                elif "CANTIDAD" in txt and "cantidad" not in col_v: col_v["cantidad"] = c
                elif ("SHIPFROM" in txt or "SHIP FROM" in txt or "CANTERA" in txt) and "sf" not in col_v: col_v["sf"] = c
                elif ("CENTRO" in txt or "CEDIS" in txt or "PLANTA" in txt) and "centro" not in col_v: col_v["centro"] = c
                elif any(k in txt for k in ["UM", "UNIDAD", "MEDIDA"]) and "um" not in col_v: col_v["um"] = c
                elif any(k in txt for k in ["POSICION", "POSICIÓN", "POSNR", "PARTIDA", "POS"]) and "posicion" not in col_v: col_v["posicion"] = c

            c_cto = col_v.get("contrato", 2)
            c_clase = col_v.get("clase", 3)
            c_vkorg = col_v.get("vkorg", 4)
            c_dest = col_v.get("destino", 8)
            c_dfact = col_v.get("dest_fact", 7)
            c_ini = col_v.get("fch_ini", 10)
            c_fin = col_v.get("fch_fin", 11)
            c_mod = col_v.get("modalidad", 13)
            c_mat = col_v.get("material", 18)
            c_cant = col_v.get("cantidad", 19)
            c_sf = col_v.get("sf", 23)
            c_cedis = col_v.get("centro", 25)
            c_um = col_v.get("um", 27)

            h_data = None
            h_row = None
            for r in range(1, ws.max_row + 1):
                c1 = str(ws.cell(r, 1).value or "").strip().upper()
                if c1 == "H" or str(ws.cell(r, c_clase).value or "").strip().upper() == "ZCP":
                    h_row = r
                    h_data = {
                        "contrato": str(ws.cell(r, c_cto).value or "").strip(),
                        "clase": str(ws.cell(r, c_clase).value or "").strip(),
                        "vkorg": str(ws.cell(r, c_vkorg).value or "").strip() or "7180",
                        "destino": str(ws.cell(r, c_dest).value or "").strip(),
                        "dest_fact": str(ws.cell(r, c_dfact).value or "").strip() if ws.cell(r, c_dfact).value else "",
                        "fch_ini": str(ws.cell(r, c_ini).value or "").strip(),
                        "fch_fin": str(ws.cell(r, c_fin).value or "").strip(),
                        "modalidad": clean_exp_cond(ws.cell(r, c_mod).value),
                        "moneda": str(ws.cell(r, col_v.get("moneda", 12)).value or "MXN").strip(),
                    }
                elif c1 == "I" and h_data:
                    material = str(ws.cell(r, c_mat).value or "").strip()
                    cantidad = clean_num(ws.cell(r, c_cant).value)
                    sf = str(ws.cell(r, c_sf).value or "").strip()
                    cedis = str(ws.cell(r, c_cedis).value or "").strip()
                    um = clean_um(ws.cell(r, c_um).value)

                    # Cruce inteligente con la hoja FPA (por Material y Destino)
                    fpa_ref = None
                    for prev_item in items:
                        if prev_item.get("tipo_seccion") == "FPA" and mat_matches(prev_item.get("material"), material):
                            if not h_data["destino"] or dest_matches({"cliente": prev_item.get("destino")}, h_data["destino"]):
                                fpa_ref = prev_item
                                break
                    if not fpa_ref:
                        for prev_item in items:
                            if prev_item.get("tipo_seccion") == "FPA" and mat_matches(prev_item.get("material"), material):
                                fpa_ref = prev_item
                                break

                    if fpa_ref:
                        if not sf:
                            sf = fpa_ref.get("sf", "")
                        if not cedis:
                            cedis = fpa_ref.get("cedis", "")
                        p_mp_prev = clean_num(fpa_ref.get("precio_mat")) or 0.0
                        p_fte_prev = clean_num(fpa_ref.get("precio_flete")) or 0.0
                        if (p_mp_prev or p_fte_prev) and "precio" not in col_v:
                            precio_mat = p_mp_prev + p_fte_prev
                        if not h_data.get("modalidad") or h_data.get("modalidad") == "01":
                            h_data["modalidad"] = fpa_ref.get("modalidad", "01")
                        if not h_data.get("fch_ini"):
                            h_data["fch_ini"] = fpa_ref.get("fch_ini", "")
                        if not um and fpa_ref.get("um"):
                            um = fpa_ref.get("um")

                    # Buscar precio del material en la fila si venía explícito
                    if "precio" in col_v:
                        precio_mat = clean_num(ws.cell(r, col_v["precio"]).value)

                    c_raw = h_data["contrato"]
                    pos_val = ""
                    if "-" in c_raw:
                        parts_c = c_raw.split("-")
                        c_raw = parts_c[0].strip()
                        if len(parts_c) > 1:
                            pos_val = parts_c[1].strip()
                    if not pos_val and "posicion" in col_v:
                        pos_val = str(ws.cell(r, col_v["posicion"]).value or "").strip()

                    sol_detectada = "VALIDAR" if c_raw else "ALTA"

                    c_label = c_raw if c_raw else "(Por crear)"
                    clave_str = f"[{hoja}] Org {h_data['vkorg']} | Centro {cedis} | SF {sf} | Cliente {h_data['destino']} | Mat {material} | Mod {h_data['modalidad']}"
                    p_txt = f"${precio_mat:,.2f}" if precio_mat is not None else ""
                    precio_str = f"Precio: {p_txt} | Cant: {cantidad:,.0f} {um} | Cto: {c_label}" if cantidad else f"Cto: {c_label}"

                    items.append({
                        "idx_fila": h_row,
                        "fila_origen": f"{h_row}-{r}",
                        "hoja": hoja,
                        "tipo_seccion": "CTO_VENTA",
                        "solicitud": sol_detectada,
                        "vkorg": h_data["vkorg"],
                        "sf": sf,
                        "cedis": cedis,
                        "destino": h_data["destino"],
                        "dest_factura": h_data.get("dest_fact", ""),
                        "material": material,
                        "modalidad": h_data["modalidad"] or "01",
                        "precio_mat": precio_mat,
                        "precio_flete": None,
                        "contrato": c_raw,
                        "posicion": pos_val,
                        "proveedor": sf,
                        "um": um,
                        "cantidad": cantidad,
                        "pv": None,
                        "moneda": h_data.get("moneda", "MXN"),
                        "fch_ini": h_data.get("fch_ini", ""),
                        "fch_fin": h_data.get("fch_fin", ""),
                        "clave_str": clave_str,
                        "precio_str": precio_str,
                        "raw_dict": {
                            "fila": h_row,
                            "sf": sf,
                            "cedis": cedis,
                            "destino": h_data["destino"],
                            "planta_obra": h_data["destino"],
                            "material": material,
                            "articulo": material,
                            "modalidad": h_data["modalidad"] or "01",
                            "um": um,
                        },
                    })

        # -------------------------------------------------------------------
        # 4. MATRIZ CONSOLIDADA (VK13 + COMPRA + VENTA)
        # -------------------------------------------------------------------
        elif tipo == "MATRIZ_CONSOLIDADA":
            h_row2 = [str(c.value or "").strip().upper() for c in ws[2]] if ws.max_row >= 2 else []
            has_cond_fte = any("FLETE" in h and ("COND" in h or "CLASE" in h) for h in h_row2)
            f_off = 1 if has_cond_fte else 0
            for r_idx, row in enumerate(ws.iter_rows(min_row=3, values_only=True), start=3):
                concat1 = str(row[0] or "").strip()
                soc = str(row[2] or "").strip()
                sf = str(row[3] or "").strip()
                sf_nom = str(row[4] or "").strip()
                centro = str(row[5] or "").strip()
                centro_desc = str(row[6] or "").strip()
                dest = str(row[7] or "").strip()
                dest_nom = str(row[8] or "").strip()
                mat = str(row[9] or "").strip()
                mat_desc = str(row[10] or "").strip()

                if not sf and not dest and not mat:
                    continue

                pv = clean_num(row[11])
                ini_vk13 = str(row[12] or "").strip()
                fin_vk13 = str(row[13] or "").strip()
                mod_vta = str(row[14] or "").strip()
                cond_mp = str(row[15] or "").strip()
                imp_mp = clean_num(row[16])
                um_vta = clean_um(row[17]) or "TN"
                cond_fte = str(row[18] or "").strip() if has_cond_fte else ""
                imp_fte = clean_num(row[18 + f_off])

                cto_c_raw = str(row[19 + f_off] or "").strip()
                ini_c = str(row[20 + f_off] or "").strip()
                fin_c = str(row[21 + f_off] or "").strip()
                costo_tot = clean_num(row[22 + f_off])
                fte_c = clean_num(row[23 + f_off])
                mp_c = clean_num(row[24 + f_off])

                cto_v_raw = str(row[32 + f_off] or "").strip()
                um_v = clean_um(row[33 + f_off]) or "M3"
                precio_v = clean_num(row[34 + f_off])

                # Parse contract numbers
                cc_num = cto_c_raw.split("-")[0].strip() if "-" in cto_c_raw else cto_c_raw
                cc_pos = cto_c_raw.split("-")[1].strip() if "-" in cto_c_raw else ""
                cv_num = cto_v_raw.split("-")[0].strip() if "-" in cto_v_raw else cto_v_raw
                cv_pos = cto_v_raw.split("-")[1].strip() if "-" in cto_v_raw else ""

                clave_str = f"{sf} | {centro} | {dest} | {mat}"
                precio_disp = f"PV:${pv:,.2f}" if pv is not None else (f"MP:${imp_mp:,.2f}" if imp_mp is not None else "-")

                items.append({
                    "idx_fila": r_idx,
                    "fila_origen": str(r_idx),
                    "hoja": hoja,
                    "tipo_seccion": "MATRIZ_CONSOLIDADA",
                    "clave_str": clave_str,
                    "precio_str": precio_disp,
                    "concat1": concat1,
                    "sociedad": soc,
                    "vkorg": soc if soc in ("7180", "7100") else "7100",
                    "sf": sf,
                    "sf_nom": sf_nom,
                    "proveedor": sf,
                    "centro": centro,
                    "cedis": centro,
                    "centro_desc": centro_desc,
                    "destino": dest,
                    "dest_nom": dest_nom,
                    "cliente": dest,
                    "planta_obra": dest,
                    "material": mat,
                    "articulo": mat,
                    "material_desc": mat_desc,
                    # VK13
                    "pv": pv,
                    "ini_vk13": ini_vk13,
                    "fin_vk13": fin_vk13,
                    "modalidad": clean_exp_cond(mod_vta or "01"),
                    "kschl_mat": cond_mp or ("ZMA6" if soc == "7180" else "ZMAH"),
                    "kschl_fte": cond_fte or ("ZMP1" if cond_mp in ("ZMA6", "ZMP1") else "ZMPH"),
                    "cond_flete": cond_fte,
                    "precio_mat": imp_mp,
                    "precio_flete": imp_fte,
                    "precio_fte": imp_fte,
                    "um": um_vta,
                    # Compra
                    "contrato_compra_raw": cto_c_raw,
                    "contrato_compra": cc_num,
                    "posicion_compra": cc_pos,
                    "ini_compra": ini_c,
                    "fin_compra": fin_c,
                    "costo_total_compra": costo_tot,
                    "flete_compra": fte_c,
                    "mp_compra": mp_c,
                    # Venta
                    "contrato_venta_raw": cto_v_raw,
                    "contrato_venta": cv_num,
                    "posicion_venta": cv_pos,
                    "um_venta_cto": um_v,
                    "precio_venta_cto": precio_v,
                    "solicitud": "",
                    "comentario": "",
                })

    wb.close()
    return items


# ---------------------------------------------------------------------------
# Motor de Validación en SAP GUI (VK13 / Z0116_TRAOPE)
# ---------------------------------------------------------------------------
class MotorValidadorSAP:
    """Ejecutor de validaciones en vivo sobre SAP GUI con control de ritmo."""

    def __init__(self, speed_delay=0.45, on_progress=None, on_log=None):
        """
        speed_delay: pausa entre pasos visuales (segundos)
        on_progress: callback(idx, total, item, estado, detalle_sap, mensaje)
        on_log: callback(texto)
        """
        self.speed_delay = speed_delay
        self.delay = speed_delay
        self.on_progress = on_progress
        self.on_log = on_log
        self.pause_event = threading.Event()
        self.pause_event.set()  # No pausado por defecto
        self.stop_event = threading.Event()
        self.session = None

    def log(self, mensaje):
        if self.on_log:
            ts = datetime.now().strftime("%H:%M:%S")
            self.on_log(f"[{ts}] {mensaje}")

    def _sleep(self, factor=1.0):
        """Pausa visual configurable que respeta pausas y cancelaciones."""
        if self.stop_event.is_set():
            raise InterruptedError("Validación cancelada por el usuario.")
        self.pause_event.wait()

        tiempo = max(0.05, self.speed_delay * factor)
        time.sleep(tiempo)

    def conectar(self):
        self.log("Conectando con SAP GUI...")
        self.session = get_sap_session(bring_to_foreground=True)
        self.log("Conectado exitosamente. Sesión de SAP GUI maximizada y visible.")

    def validar_lote(self, items, tipo_info, ruta_excel=None):
        """Ejecuta la validación fila por fila según el tipo detectado."""
        import pythoncom
        try:
            pythoncom.CoInitialize()
        except Exception:
            pass

        try:
            self.conectar()
            total = len(items)
            tipo_global = tipo_info.get("tipo", "FPA_FILIAL")
            subtipo = tipo_info.get("subtipo", "COMPLETO")

            self.log(f"Iniciando validación de {total} fila(s) para: {tipo_info['descripcion']}")

            # Precache de Z0116 para items de Contrato de Compra si están presentes
            preloaded_z0116 = None
            cto_compra_items = [it for it in items if it.get("tipo_seccion") == "CTO_COMPRA"]
            if cto_compra_items:
                self.log(f"Detectados {len(cto_compra_items)} item(s) de Contrato de Compra. Precargando Z0116...")
                preloaded_z0116 = self._precargar_z0116(cto_compra_items)

            # Para items de Contrato de Venta, validación directa en VA43 por partida (descartando ZSDD4501)
            cto_venta_items = [it for it in items if it.get("tipo_seccion") == "CTO_VENTA"]
            if cto_venta_items:
                self.log(f"Detectados {len(cto_venta_items)} item(s) de Contrato de Venta. Se validarán directamente vía VA43...")

            # Precache global para MATRIZ CONSOLIDADA
            preloaded_mat_compra = None
            preloaded_mat_venta = None
            cache_vk13_mat = {}
            cache_zsdd4501_mat = {}

            if any(it.get("tipo_seccion") == "MATRIZ_CONSOLIDADA" for it in items):
                validar_c = getattr(self, "val_matriz_chk_compra", True)
                validar_v = getattr(self, "val_matriz_chk_venta", True)

                if validar_c:
                    items_c_prep = []
                    for it in items:
                        r_c = str(it.get("contrato_compra_raw") or "").strip()
                        if r_c and not any(k in r_c.upper() for k in ["CREAR", "ALTA", "POR CREAR", "NONE", "NAN", "S/C"]):
                            c_n = r_c.split("-")[0].strip() if "-" in r_c else r_c
                            items_c_prep.append({
                                "cedis": it.get("centro"), "sf": it.get("sf"), "material": it.get("material"),
                                "contrato": c_n
                            })
                    if items_c_prep:
                        self.log(f"Precargando Z0116 para Matriz ({len(items_c_prep)} rutas con contrato compra)...")
                        preloaded_mat_compra = self._precargar_z0116(items_c_prep)

                if validar_v:
                    self.log("Auditoría de Contratos de Venta (ZSDD4501): ejecutando consultas activas fila por fila en SAP GUI...")
                    preloaded_mat_venta = None
            self._cache_mm03_lote = {}
            self._cache_z1017 = {}
            self._contratos_vta_lote = {}
            resultados = []
            for i, item in enumerate(items, start=1):
                if self.stop_event.is_set():
                    self.log("Proceso detenido por el usuario.")
                    break
                self.pause_event.wait()

                tipo_sec = item.get("tipo_seccion", tipo_global)
                hoja_origen = item.get("hoja", "")
                fila_desc = f"{hoja_origen}: Fila {item.get('fila_origen', i)}"

                # Notificar inicio de fila
                if self.on_progress:
                    self.on_progress(i, total, item, "EN_PROCESO", {}, f"Consultando [{tipo_sec}] en SAP GUI...")

                try:
                    if tipo_sec == "FPA":
                        tipo_fpa = item.get("tipo_fpa", tipo_global)
                        res = self._validar_fila_fpa(item, tipo_fpa, subtipo)
                    elif tipo_sec == "CTO_COMPRA":
                        if preloaded_z0116 is None:
                            preloaded_z0116 = self._precargar_z0116(cto_compra_items)
                        res = self._validar_fila_cto_compra(item, preloaded_z0116)
                    elif tipo_sec == "CTO_VENTA":
                        res = self._validar_fila_cto_venta(item)
                    elif tipo_sec == "MATRIZ_CONSOLIDADA":
                        val_vk = getattr(self, "val_matriz_chk_vk13", True)
                        val_cp = getattr(self, "val_matriz_chk_compra", True)
                        val_vt = getattr(self, "val_matriz_chk_venta", True)

                        res = self._validar_fila_matriz_consolidada(
                            item,
                            preloaded_compra_records=preloaded_mat_compra,
                            preloaded_venta_records=preloaded_mat_venta,
                            cache_vk13=cache_vk13_mat,
                            cache_zsdd4501=cache_zsdd4501_mat,
                            validar_vk13=val_vk,
                            validar_compra=val_cp,
                            validar_venta=val_vt
                        )

                        # Guardado incremental cada 25 filas
                        if ruta_excel and i % 25 == 0:
                            try:
                                escribir_resultados_en_excel(ruta_excel, resultados + [(item, res)])
                            except Exception:
                                pass
                    else:
                        res = {"estado": "ERROR", "mensaje": f"Tipo de validación no implementado ({tipo_sec})"}

                    # Validación exhaustiva de Material (MM03) para pestaña 'Validación SAP GUI' (FPA, CTO_COMPRA, CTO_VENTA)
                    if tipo_sec in ("FPA", "CTO_COMPRA", "CTO_VENTA") and item.get("hoja") != "Matriz":
                        mat_item = str(item.get("material") or "").strip()
                        centro_item = str(item.get("cedis") or item.get("centro") or "").strip()
                        org_item = str(item.get("vkorg") or item.get("org_ventas") or item.get("company") or item.get("sociedad") or "7180").strip()

                        if mat_item:
                            mm03_key = (mat_item, centro_item, org_item)
                            if mm03_key in self._cache_mm03_lote:
                                res_mm03 = self._cache_mm03_lote[mm03_key]
                            else:
                                if self.on_progress:
                                    self.on_progress(i, total, item, "EN_PROCESO", res.get("datos_sap", {}), f"Validando Material {mat_item} en MM03...")
                                res_mm03 = self.validar_material_mm03(
                                    material=mat_item,
                                    centro=centro_item,
                                    org_ventas=org_item,
                                    canal=item.get("canal", "00"),
                                    almacen=item.get("almacen", "1002"),
                                    peso_esp=item.get("peso_esp") or item.get("peso"),
                                    clase_esp=item.get("clase_esp") or item.get("clase")
                                )
                                self._cache_mm03_lote[mm03_key] = res_mm03

                            if not res.get("datos_sap"):
                                res["datos_sap"] = {}
                            res["datos_sap"]["res_mm03"] = res_mm03

                            if not res_mm03.get("valido"):
                                if res.get("estado") == "OK":
                                    res["estado"] = "DISCREPANCIA"
                                res["mensaje"] = f"{res.get('mensaje', '')} | MM03: {res_mm03.get('msg', 'Error material')}"
                            else:
                                res["mensaje"] = f"{res.get('mensaje', '')} | MM03 OK"

                    resultados.append((item, res))

                    if self.on_progress:
                        self.on_progress(i, total, item, res["estado"], res.get("datos_sap", {}), res["mensaje"])

                except InterruptedError:
                    break
                except Exception as e:
                    self.log(f"Error procesando {fila_desc}: {e}")
                    if self.on_progress:
                        self.on_progress(i, total, item, "ERROR", {}, str(e))

                self._sleep(0.5)

            if ruta_excel and os.path.exists(ruta_excel):
                try:
                    self.log("Generando archivo con resultados de auditoría y colores...")
                    out_path = escribir_resultados_en_excel(ruta_excel, resultados)
                    if out_path:
                        self.log(f"Archivo Excel generado exitosamente: {os.path.basename(out_path)}")
                except Exception as e_xl:
                    self.log(f"Aviso guardando archivo Excel validado: {e_xl}")

            self.log("Validación completada.")
            return resultados
        finally:
            try:
                pythoncom.CoUninitialize()
            except Exception:
                pass

    # -----------------------------------------------------------------------
    # Flujo VK13 (FPA FILIAL / FPA TERCEROS)
    # -----------------------------------------------------------------------
    def _validar_fila_matriz_consolidada(self, item, preloaded_compra_records=None, preloaded_venta_records=None, cache_vk13=None, cache_zsdd4501=None, validar_vk13=True, validar_compra=True, validar_venta=True):
        """Valida integralmente los módulos activos (VK13, Cto Compra, Cto Venta) usando los motores ya existentes."""
        res_vk13 = {"estado": "N/A", "mensaje": "Omitido", "datos_sap": {}}
        res_compra = {"estado": "N/A", "mensaje": "Omitido", "datos_sap": {}}
        res_venta = {"estado": "N/A", "mensaje": "Omitido", "datos_sap": {}}

        # -------------------------------------------------------------
        # 1. MOTOR VK13 (_validar_fila_fpa)
        # -------------------------------------------------------------
        if validar_vk13 and (item.get("precio_mat") is not None or item.get("pv") is not None):
            route_key = (
                item.get("kschl_mat", "ZMAH"),
                item.get("vkorg", "7100"),
                str(item.get("sf", "")).strip(),
                str(item.get("cedis", "")).strip(),
                str(item.get("destino", "")).strip(),
                str(item.get("material", "")).strip()
            )
            if cache_vk13 is not None and route_key in cache_vk13:
                res_vk13 = cache_vk13[route_key]
            else:
                it_vk = {
                    "idx_fila": item.get("fila_origen", 3),
                    "fila_origen": str(item.get("fila_origen", 3)),
                    "hoja": item.get("hoja", "Matriz"),
                    "tipo_seccion": "FPA",
                    "tipo_fpa": "FPA_TERCEROS" if item.get("kschl_mat") in ("ZMA6", "ZMP1") else "FPA_FILIAL",
                    "solicitud": "",
                    "vkorg": item.get("sociedad") or "7100",
                    "sf": item.get("sf"),
                    "cedis": item.get("centro"),
                    "destino": item.get("destino"),
                    "material": item.get("material"),
                    "modalidad": clean_exp_cond(item.get("modalidad") or "01"),
                    "kschl_mat": item.get("kschl_mat") or ("ZMA6" if item.get("sociedad") == "7180" else "ZMAH"),
                    "kschl_fte": "ZMP1" if item.get("kschl_mat") in ("ZMA6", "ZMP1") else "ZMPH",
                    "precio_mat": item.get("precio_mat"),
                    "precio_flete": item.get("precio_fte"),
                    "um": item.get("um") or "TN",
                    "pv": item.get("pv"),
                    "articulo": item.get("material"),
                    "planta_obra": item.get("destino"),
                    "comentario": "",
                }
                res_vk13 = self._validar_fila_fpa(it_vk, it_vk["tipo_fpa"], "COMPLETO")
                if cache_vk13 is not None:
                    cache_vk13[route_key] = res_vk13

        # -------------------------------------------------------------
        # 2. MOTOR CTO COMPRA (_validar_fila_cto_compra)
        # -------------------------------------------------------------
        if validar_compra:
            raw_cc = str(item.get("contrato_compra_raw") or "").strip()
            if not raw_cc or raw_cc.lower() in ("none", "nan", "s/c", "sin contrato", "-") or any(k in raw_cc.upper() for k in ["CREAR", "ALTA", "POR CREAR"]):
                res_compra = {"estado": "SIN_CONTRATO", "mensaje": "Sin Contrato Compra Asignado en Matriz", "datos_sap": {}}
            else:
                cc_num = raw_cc.split("-")[0].strip() if "-" in raw_cc else raw_cc
                cc_pos = raw_cc.split("-")[1].strip() if "-" in raw_cc else ""
                it_compra = {
                    "idx_fila": item.get("fila_origen", 3),
                    "fila_origen": str(item.get("fila_origen", 3)),
                    "hoja": item.get("hoja", "Matriz"),
                    "tipo_seccion": "CTO_COMPRA",
                    "contrato": cc_num,
                    "posicion": cc_pos,
                    "sf": item.get("sf"),
                    "proveedor": item.get("sf"),
                    "cedis": item.get("centro"),
                    "destino": item.get("destino"),
                    "material": item.get("material"),
                    "precio_mat": item.get("costo_total_compra") or item.get("mp_compra"),
                    "costo_total_compra": item.get("costo_total_compra"),
                    "mp_compra": item.get("mp_compra"),
                    "flete_compra": item.get("flete_compra"),
                    "modalidad": clean_exp_cond(item.get("modalidad") or "01"),
                    "um": item.get("um") or "M3",
                    "solicitud": "",
                    "comentario": "",
                }
                res_compra = self._validar_fila_cto_compra(it_compra, preloaded_records=preloaded_compra_records)

        # -------------------------------------------------------------
        # 3. MOTOR CTO VENTA (_validar_fila_cto_venta) - FILA POR FILA
        # -------------------------------------------------------------
        if validar_venta:
            raw_cv = str(item.get("contrato_venta_raw") or "").strip()
            if not raw_cv or raw_cv.lower() in ("none", "nan", "s/c", "sin contrato", "-") or any(k in raw_cv.upper() for k in ["CREAR", "ALTA", "POR CREAR"]):
                res_venta = {"estado": "SIN_CONTRATO", "mensaje": "Sin Contrato Venta Asignado en Matriz", "datos_sap": {}}
            else:
                cv_num = raw_cv.split("-")[0].strip() if "-" in raw_cv else raw_cv
                cv_pos = raw_cv.split("-")[1].strip() if "-" in raw_cv else str(item.get("posicion_venta") or "10").strip()

                it_venta = {
                    "idx_fila": item.get("fila_origen", 3),
                    "fila_origen": str(item.get("fila_origen", 3)),
                    "hoja": item.get("hoja", "Matriz"),
                    "tipo_seccion": "CTO_VENTA",
                    "contrato": cv_num,
                    "posicion": cv_pos or "10",
                    "sf": item.get("sf"),
                    "proveedor": item.get("sf"),
                    "dest_factura": item.get("sf"),
                    "cedis": item.get("centro"),
                    "destino": item.get("destino"),
                    "cliente": item.get("destino"),
                    "material": item.get("material"),
                    "modalidad": clean_exp_cond(item.get("modalidad") or "01"),
                    "precio_venta_cto": item.get("precio_venta_cto"),
                    "precio_mat": item.get("precio_mat"),
                    "precio_flete": item.get("precio_fte"),
                    "pv": item.get("pv"),
                    "um": item.get("um_venta_cto") or item.get("um") or "M3",
                    "vkorg": item.get("sociedad") or "7180",
                    "solicitud": "",
                    "comentario": "",
                }

                cache_key_vta = (cv_num, cv_pos, str(item.get("sf", "")).strip(), str(item.get("centro", "")).strip(), str(item.get("destino", "")).strip(), str(item.get("material", "")).strip())
                if cache_zsdd4501 is not None and cache_key_vta in cache_zsdd4501:
                    res_venta = cache_zsdd4501[cache_key_vta]
                else:
                    res_venta = self._validar_fila_cto_venta(it_venta, preloaded_records=None)
                    if cache_zsdd4501 is not None:
                        cache_zsdd4501[cache_key_vta] = res_venta

        # -------------------------------------------------------------
        # 4. DETERMINAR ESTADO GLOBAL DE FILA Y RESUMEN
        # -------------------------------------------------------------
        estados_eval = []
        if validar_vk13: estados_eval.append(("VK13", res_vk13["estado"]))
        if validar_compra: estados_eval.append(("COMPRA", res_compra["estado"]))
        if validar_venta: estados_eval.append(("VENTA", res_venta["estado"]))

        discrepancias = []
        detalles_global = []

        for mod, est in estados_eval:
            if est == "DISCREPANCIA":
                discrepancias.append(mod)
                m_txt = res_vk13["mensaje"] if mod == "VK13" else (res_compra["mensaje"] if mod == "COMPRA" else res_venta["mensaje"])
                detalles_global.append(f"[{mod}: ERROR] {m_txt}")
            elif est == "OK":
                detalles_global.append(f"[{mod}: OK]")
            elif est == "SIN_CONTRATO":
                detalles_global.append(f"[{mod}: S/C]")
            elif est == "NO_ENCONTRADO":
                detalles_global.append(f"[{mod}: NO ENCONTRADO]")

        if discrepancias:
            estado_global = "DISCREPANCIA"
            resumen_texto = " | ".join(detalles_global)
        elif any(est == "NO_ENCONTRADO" for _, est in estados_eval):
            estado_global = "REVISION"
            resumen_texto = " | ".join(detalles_global)
        else:
            estado_global = "OK"
            resumen_texto = " | ".join(detalles_global) if detalles_global else "OK"

        return {
            "estado": estado_global,
            "mensaje": resumen_texto,
            "res_vk13": res_vk13,
            "res_compra": res_compra,
            "res_venta": res_venta,
            "datos_sap": {
                "vk13": res_vk13.get("datos_sap", {}),
                "compra": res_compra.get("datos_sap", {}),
                "venta": res_venta.get("datos_sap", {}),
                "estado_global": estado_global,
            }
        }

    def _validar_fila_fpa(self, item, tipo, subtipo):
        """Valida una fila de FPA en VK13."""
        solicitud = str(item.get("solicitud", "")).upper()
        kschl_mat = item.get("kschl_mat") or ("ZMA6" if tipo == "FPA_TERCEROS" else "ZMAH")
        kschl_fte = item.get("kschl_fte") or ("ZMP1" if tipo == "FPA_TERCEROS" else "ZMPH")

        # Usar la Org. Ventas que viene en el archivo (ej. 7100)
        vkorg = str(item.get("vkorg") or "7100").strip()
        sf = item["sf"]
        cedis = item["cedis"]
        destino = item["destino"]
        material = item["material"]
        modalidad = item["modalidad"]

        resultado_mat = {"valido": True, "precio": None, "msg": ""}
        resultado_fte = {"valido": True, "precio": None, "msg": ""}

        # Opción de combinación de claves en VK13:
        # FPA TERCEROS (ZMA6 / ZMP1) -> 2ª opción (índice 1)
        # FPA FILIAL (ZMAH / ZMPH)   -> 1ª opción (índice 0)
        hoja_nombre = str(item.get("hoja", "")).upper()
        opt_idx = 1 if (tipo == "FPA_TERCEROS" or "TERCEROS" in hoja_nombre or kschl_fte in ("ZMA6", "ZMP1") or kschl_mat in ("ZMA6", "ZMP1")) else 0

        # 1. Validar Flete (ZMPH / ZMP1)
        if item.get("precio_flete") is not None or subtipo == "SOLO_FLETES":
            self.log(f"[Fila {item['idx_fila']}] Validando Flete ({kschl_fte}) para Org {vkorg}, SF {sf}, Centro {cedis}, Destino {destino}, Mat {material} (Opción claves: {opt_idx + 1})...")
            res_fte = self._consultar_vk13(kschl_fte, vkorg, sf, cedis, destino, material, modalidad, option_idx=opt_idx)
            if not res_fte["encontrado"]:
                if solicitud == "ALTA":
                    resultado_fte["valido"] = False
                    resultado_fte["msg"] = f"DISCREPANCIA: Solicitud de ALTA confirmada pero Flete {kschl_fte} NO encontrado en SAP VK13"
                elif solicitud == "BAJA":
                    resultado_fte["valido"] = True
                    resultado_fte["msg"] = f"Flete {kschl_fte} dado de baja en SAP VK13 (OK)"
                else:
                    resultado_fte["valido"] = False
                    resultado_fte["msg"] = f"Flete {kschl_fte} no encontrado en SAP (Org {vkorg})"
            else:
                resultado_fte["precio"] = res_fte["importe"]
                resultado_fte["datos_sap"] = res_fte
                if solicitud == "BAJA":
                    resultado_fte["valido"] = False
                    resultado_fte["msg"] = f"DISCREPANCIA: Solicitud de BAJA pero Flete {kschl_fte} sigue activo en SAP VK13 (${res_fte['importe']:,.2f})"
                elif item.get("precio_flete") is not None and res_fte["importe"] is not None:
                    # Comparar precios si el archivo traía precio esperado
                    diff = abs(item["precio_flete"] - res_fte["importe"])
                    if diff > 0.05:
                        resultado_fte["valido"] = False
                        resultado_fte["msg"] = f"Flete difiere: SAP ${res_fte['importe']:,.2f} vs Archivo ${item['precio_flete']:,.2f}"
                    else:
                        resultado_fte["msg"] = f"Flete OK: ${res_fte['importe']:,.2f}"
                else:
                    resultado_fte["msg"] = f"Flete activo en SAP: ${res_fte['importe']:,.2f}" if res_fte['importe'] else "Flete en SAP"

                # Validar modalidad si SAP la reportó
                mod_xl = clean_exp_cond(modalidad)
                mod_sap = clean_exp_cond(res_fte.get("vsbed"))
                if mod_xl and mod_sap:
                    if mod_xl != mod_sap:
                        resultado_fte["valido"] = False
                        resultado_fte["msg"] += f" (Mod difiere: SAP {mod_sap} vs XL {mod_xl})"
                    else:
                        resultado_fte["msg"] += f" (Mod {mod_sap} OK)"

                # Validar Unidad de Medida (UM) si SAP la reportó
                um_xl = clean_um(item.get("um"))
                um_sap_fte = clean_um(res_fte.get("um"))
                if um_xl and um_sap_fte:
                    if um_xl != um_sap_fte:
                        resultado_fte["valido"] = False
                        resultado_fte["msg"] += f" (UM difiere: SAP {um_sap_fte} vs XL {um_xl})"
                    else:
                        resultado_fte["msg"] += f" (UM {um_sap_fte} OK)"

        # 2. Validar Material (ZMAH / ZMA6) si no es Sólo Fletes
        if subtipo != "SOLO_FLETES" and item.get("precio_mat") is not None:
            self.log(f"[Fila {item['idx_fila']}] Validando Material ({kschl_mat}) para Org {vkorg}, SF {sf}, Centro {cedis}, Destino {destino}, Mat {material} (Opción claves: {opt_idx + 1})...")
            res_mat = self._consultar_vk13(kschl_mat, vkorg, sf, cedis, destino, material, modalidad, option_idx=opt_idx)
            if not res_mat["encontrado"]:
                if solicitud == "ALTA":
                    resultado_mat["valido"] = False
                    resultado_mat["msg"] = f"DISCREPANCIA: Solicitud de ALTA confirmada pero Material {kschl_mat} NO encontrado en SAP VK13"
                elif solicitud == "BAJA":
                    resultado_mat["valido"] = True
                    resultado_mat["msg"] = f"Material {kschl_mat} dado de baja en SAP VK13 (OK)"
                else:
                    resultado_mat["valido"] = False
                    resultado_mat["msg"] = f"Material {kschl_mat} no encontrado en SAP (Org {vkorg})"
            else:
                resultado_mat["precio"] = res_mat["importe"]
                resultado_mat["datos_sap"] = res_mat
                if solicitud == "BAJA":
                    resultado_mat["valido"] = False
                    resultado_mat["msg"] = f"DISCREPANCIA: Solicitud de BAJA pero Material {kschl_mat} sigue activo en SAP VK13 (${res_mat['importe']:,.2f})"
                elif item.get("precio_mat") is not None and res_mat["importe"] is not None:
                    diff = abs(item["precio_mat"] - res_mat["importe"])
                    if diff > 0.05:
                        resultado_mat["valido"] = False
                        resultado_mat["msg"] = f"Material difiere: SAP ${res_mat['importe']:,.2f} vs Archivo ${item['precio_mat']:,.2f}"
                    else:
                        resultado_mat["msg"] = f"Material OK: ${res_mat['importe']:,.2f}"
                else:
                    resultado_mat["msg"] = f"Material activo en SAP: ${res_mat['importe']:,.2f}" if res_mat['importe'] else "Material en SAP"

                # Validar Unidad de Medida (UM) si SAP la reportó
                um_xl = clean_um(item.get("um"))
                um_sap_mat = clean_um(res_mat.get("um"))
                if um_xl and um_sap_mat:
                    if um_xl != um_sap_mat:
                        resultado_mat["valido"] = False
                        resultado_mat["msg"] += f" (UM difiere: SAP {um_sap_mat} vs XL {um_xl})"
                    else:
                        resultado_mat["msg"] += f" (UM {um_sap_mat} OK)"

        # Evaluar estado final
        datos_sap = {}
        if "datos_sap" in resultado_fte:
            datos_sap.update(resultado_fte["datos_sap"])
            datos_sap["flete_encontrado"] = True
            datos_sap["precio_fte_sap"] = resultado_fte["datos_sap"].get("importe")
            datos_sap["um_fte_sap"] = resultado_fte["datos_sap"].get("um")
            datos_sap["moneda_fte_sap"] = resultado_fte["datos_sap"].get("moneda")
            datos_sap["modalidad_fte_sap"] = resultado_fte["datos_sap"].get("vsbed")
            datos_sap["datab_fte_sap"] = resultado_fte["datos_sap"].get("datab")
            datos_sap["datbi_fte_sap"] = resultado_fte["datos_sap"].get("datbi")
            datos_sap["status_fte"] = resultado_fte.get("msg")
        if "datos_sap" in resultado_mat:
            datos_sap["material_encontrado"] = True
            datos_sap["precio_mat_sap"] = resultado_mat["datos_sap"].get("importe")
            datos_sap["um_mat_sap"] = resultado_mat["datos_sap"].get("um")
            datos_sap["moneda_mat_sap"] = resultado_mat["datos_sap"].get("moneda")
            datos_sap["datab_mat_sap"] = resultado_mat["datos_sap"].get("datab")
            datos_sap["datbi_mat_sap"] = resultado_mat["datos_sap"].get("datbi")
            datos_sap["status_mat"] = resultado_mat.get("msg")

        if resultado_mat["valido"] and resultado_fte["valido"]:
            mensajes = [m for m in [resultado_mat["msg"], resultado_fte["msg"]] if m]
            msg_final = " | ".join(mensajes) if mensajes else "Validado OK en VK13"
            return {"estado": "OK", "mensaje": msg_final, "datos_sap": datos_sap}
        else:
            errores = []
            if not resultado_mat["valido"]:
                errores.append(resultado_mat["msg"])
            if not resultado_fte["valido"]:
                errores.append(resultado_fte["msg"])
            return {"estado": "DISCREPANCIA", "mensaje": " | ".join(errores), "datos_sap": datos_sap}

    def _consultar_vk13(self, kschl, vkorg, sf, cedis, destino, material, modalidad="", option_idx=None):
        """Navega a VK13 con soporte completo de avisos/popups (soporte de 1ª y 2ª opción) y extracción del Table Control."""
        session = self.session

        kschl_clean = str(kschl).strip().upper()
        if option_idx is None:
            if kschl_clean in ("ZMA6", "ZMP1"):
                option_idx = 1
            else:
                option_idx = 0

        # 1. Reset a /nVK13
        session.findById("wnd[0]/tbar[0]/okcd").text = "/nvk13"
        session.findById("wnd[0]").sendVKey(0)
        self._sleep(0.2)
        handle_popups(session, delay=0.1)

        # 2. Escribir clase de condición
        session.findById("wnd[0]/usr/ctxtRV13A-KSCHL").text = kschl_clean
        session.findById("wnd[0]").sendVKey(0)
        self._sleep(0.2)

        # 3. Bucle para procesar popups (notificación previa + combinación de claves)
        max_attempts = 4
        for _ in range(max_attempts):
            if session.Children.Count <= 1:
                break
            try:
                popup = session.findById("wnd[1]")
                if not popup:
                    break

                target_opt = option_idx if option_idx is not None else 0
                radio_btn = None
                try:
                    radio_btn = session.findById(f"wnd[1]/usr/sub:SAPLV14A:0100/radRV130-SELKZ[{target_opt},0]")
                except Exception:
                    pass

                # Respaldo si no encuentra la opción indicada (revertir a opción 0)
                if not radio_btn and target_opt != 0:
                    try:
                        radio_btn = session.findById("wnd[1]/usr/sub:SAPLV14A:0100/radRV130-SELKZ[0,0]")
                    except Exception:
                        pass

                if radio_btn:
                    # Es la ventana de combinación de claves: seleccionar opción y continuar
                    radio_btn.setFocus()
                    radio_btn.select()
                    session.findById("wnd[1]/tbar[0]/btn[0]").press()
                    self._sleep(0.2)
                    break
                else:
                    # Es una notificación o aviso previo: dar Enter
                    popup.sendVKey(0)
                    self._sleep(0.2)
            except Exception:
                break

        # 4. Rellenar campos de selección en wnd[0]
        # Limpieza preventiva total de campos en pantalla preservando fechas
        clear_screen_inputs(session, exclude=["DATAB", "DATBI", "FECHA", "DATE", "GUEEN", "AEDAT", "F000"])
        for fid in [
            "wnd[0]/usr/ctxtF001", "wnd[0]/usr/ctxtF002-LOW", "wnd[0]/usr/ctxtF003-LOW",
            "wnd[0]/usr/ctxtF004-LOW", "wnd[0]/usr/ctxtF005-LOW", "wnd[0]/usr/ctxtF006-LOW",
            "wnd[0]/usr/ctxtF002-HIGH", "wnd[0]/usr/ctxtF003-HIGH", "wnd[0]/usr/ctxtF004-HIGH",
            "wnd[0]/usr/ctxtF005-HIGH", "wnd[0]/usr/ctxtF006-HIGH"
        ]:
            try: session.findById(fid).text = ""
            except Exception: pass

        # Asegurar fecha obligatoria 'Válido el' con la fecha de hoy si está vacía
        fch_hoy = datetime.now().strftime("%d.%m.%Y")
        for dfid in [
            "wnd[0]/usr/ctxtRV13A-DATAB",
            "wnd[0]/usr/ctxtKOMG-DATAB",
            "wnd[0]/usr/ctxtRV130-DATAB",
            "wnd[0]/usr/ctxtF000",
            "wnd[0]/usr/ctxtF000-LOW"
        ]:
            try:
                elem = session.findById(dfid)
                if not elem.text.strip():
                    elem.text = fch_hoy
            except Exception:
                pass

        try:
            usr = session.findById("wnd[0]/usr")
            for i in range(usr.Children.Count):
                ch = usr.Children(i)
                if getattr(ch, "Type", "") in ("GuiTextField", "GuiCTextField"):
                    cid_u = getattr(ch, "Id", "").upper()
                    if ("DATAB" in cid_u or "F000" in cid_u) and not ch.Text.strip():
                        ch.Text = fch_hoy
        except Exception:
            pass

        # Asignar valores
        try: session.findById("wnd[0]/usr/ctxtF001").text = str(vkorg).strip()
        except Exception as e: self.log(f"Aviso F001 Org {vkorg}: {e}")

        try: session.findById("wnd[0]/usr/ctxtF002-LOW").text = str(sf).strip()
        except Exception: pass

        try: session.findById("wnd[0]/usr/ctxtF003-LOW").text = str(cedis).strip()
        except Exception: pass

        try: session.findById("wnd[0]/usr/ctxtF004-LOW").text = str(destino).strip()
        except Exception: pass

        # Regla A537 para fletes (F005 vacío, F006 material) vs A772 materiales (F005 material)
        kschl_up = str(kschl).strip().upper()
        if "MPH" in kschl_up or "MP1" in kschl_up or "FTE" in kschl_up:
            try: session.findById("wnd[0]/usr/ctxtF005-LOW").text = ""
            except Exception: pass
            try: session.findById("wnd[0]/usr/ctxtF006-LOW").text = str(material).strip()
            except Exception: pass
        else:
            try: session.findById("wnd[0]/usr/ctxtF005-LOW").text = str(material).strip()
            except Exception: pass

        self._sleep(1.0)

        # 5. Ejecutar F8 (btn[8] con fallback a sendVKey(8))
        try:
            session.findById("wnd[0]/tbar[1]/btn[8]").press()
        except Exception:
            session.findById("wnd[0]").sendVKey(8)

        self._sleep(1.0)
        handle_popups(session, delay=0.1)

        # 6. Analizar barra de estado (VK021 = Sin registros)
        sbar = None
        try: sbar = session.findById("wnd[0]/sbar")
        except Exception: pass
        sbar_msg = sbar.text.strip() if sbar else ""
        sbar_no = getattr(sbar, "messageNumber", "")

        is_vk021 = (sbar_no == "021" or "021" in sbar_msg or "no existen registros" in sbar_msg.lower() or "no se han encontrado" in sbar_msg.lower())
        if is_vk021:
            return {"encontrado": False, "importe": None, "um": "", "moneda": "", "datab": "", "datbi": "", "vsbed": "", "msg": f"VK021 - Sin registros en SAP ({sbar_msg})"}

        if sbar and getattr(sbar, "messageType", "") in ("E", "A"):
            return {"encontrado": False, "importe": None, "um": "", "moneda": "", "datab": "", "datbi": "", "vsbed": "", "msg": f"Error SAP: {sbar_msg}"}

        # 7. Escanear Table Control o pantalla de resultados
        return self._escanear_table_control_vk13(kschl=kschl, target_destino=destino, target_modalidad=modalidad)

    def _escanear_table_control_vk13(self, kschl="", target_destino="", target_modalidad=""):
        """Lee las filas del Table Control en la pantalla de resultados de VK13."""
        session = self.session
        data = {
            "encontrado": False,
            "importe": None,
            "um": "",
            "moneda": "",
            "datab": "",
            "datbi": "",
            "vsbed": "",
            "kunwe": "",
        }

        try:
            usr = session.findById("wnd[0]/usr")
            tbl = None
            for i in range(usr.Children.Count):
                ch = usr.Children(i)
                if "TableControl" in getattr(ch, "Type", "") or "tbl" in getattr(ch, "Id", "").lower():
                    tbl = ch
                    break

            # 1. Columnas directas para ZMPH/ZMP1 (A537)
            if tbl and ("MPH" in str(kschl).upper() or "MP1" in str(kschl).upper()):
                try:
                    imp_val = clean_num(session.findById("wnd[0]/usr/tblSAPMV13ATCTRL_FAST_ENTRY/txtKONP-KBETR[8,0]").text)
                    if imp_val is not None:
                        data["importe"] = imp_val
                        try: data["moneda"] = session.findById("wnd[0]/usr/tblSAPMV13ATCTRL_FAST_ENTRY/ctxtKONP-KONWA[9,0]").text.strip()
                        except Exception: pass
                        try: data["um"] = session.findById("wnd[0]/usr/tblSAPMV13ATCTRL_FAST_ENTRY/ctxtKONP-KMEIN[11,0]").text.strip()
                        except Exception: pass
                        try: data["vsbed"] = clean_exp_cond(session.findById("wnd[0]/usr/tblSAPMV13ATCTRL_FAST_ENTRY/ctxtKOMG-VSBED[3,0]").text)
                        except Exception: pass
                        try: data["datab"] = session.findById("wnd[0]/usr/tblSAPMV13ATCTRL_FAST_ENTRY/ctxtRV13A-DATAB[14,0]").text.strip()
                        except Exception: pass
                        try: data["datbi"] = session.findById("wnd[0]/usr/tblSAPMV13ATCTRL_FAST_ENTRY/ctxtRV13A-DATBI[15,0]").text.strip()
                        except Exception: pass
                        data["encontrado"] = True
                        return data
                except Exception:
                    pass

            # 2. Pantalla individual (sin tabla)
            if not tbl:
                for i in range(usr.Children.Count):
                    ch = usr.Children(i)
                    cid = getattr(ch, "Id", "").upper()
                    cname = getattr(ch, "Name", "").upper()
                    ctext = getattr(ch, "Text", "").strip()

                    if "KBETR" in cname or "KBETR" in cid:
                        data["importe"] = clean_num(ctext)
                    elif "KONWA" in cname or "KONWA" in cid:
                        data["moneda"] = ctext
                    elif "KMEIN" in cname or "KMEIN" in cid:
                        data["um"] = ctext
                    elif "DATAB" in cname or "DATAB" in cid:
                        data["datab"] = ctext
                    elif "DATBI" in cname or "DATBI" in cid:
                        data["datbi"] = ctext
                    elif "VSBED" in cname or "VSBED" in cid:
                        data["vsbed"] = clean_exp_cond(ctext)

                if data["importe"] is not None:
                    data["encontrado"] = True
                return data

            # 3. Escaneo dinámico agrupado por fila usando tbl.Children
            rows_dict = {}
            for c_idx in range(tbl.Children.Count):
                cell = tbl.Children(c_idx)
                cid = getattr(cell, "Id", "")
                cname = getattr(cell, "Name", "").upper()
                ctext = getattr(cell, "Text", "").strip()

                match = re.search(r"\[(\d+),(\d+)\]", cid)
                if match:
                    row_i = int(match.group(2))
                    if row_i not in rows_dict:
                        rows_dict[row_i] = {}

                    if "KBETR" in cname or "KBETR" in cid:
                        rows_dict[row_i]["importe"] = clean_num(ctext)
                    elif "KONWA" in cname or "KONWA" in cid:
                        rows_dict[row_i]["moneda"] = ctext
                    elif "KMEIN" in cname or "KMEIN" in cid:
                        rows_dict[row_i]["um"] = ctext
                    elif "DATAB" in cname or "DATAB" in cid:
                        rows_dict[row_i]["datab"] = ctext
                    elif "DATBI" in cname or "DATBI" in cid:
                        rows_dict[row_i]["datbi"] = ctext
                    elif "VSBED" in cname or "VSBED" in cid:
                        rows_dict[row_i]["vsbed"] = clean_exp_cond(ctext)
                    elif "KUNWE" in cname or "KUNNR" in cname or "DEST" in cname:
                        rows_dict[row_i]["kunwe"] = ctext.lstrip("0")

            target_clean = str(target_destino).strip().upper().lstrip("0")
            target_mod = clean_exp_cond(target_modalidad)
            selected_row = None

            if rows_dict:
                # Priorizar coincidencia exacta de destino y modalidad
                for r_i, r_data in rows_dict.items():
                    if r_data.get("importe") is not None:
                        match_d = True
                        if target_clean and r_data.get("kunwe"):
                            match_d = (target_clean == r_data.get("kunwe", "").upper())
                        match_m = True
                        if target_mod and r_data.get("vsbed"):
                            match_m = (target_mod == r_data.get("vsbed"))
                        if match_d and match_m:
                            selected_row = r_data
                            break

                # Si no hubo coincidencia combinada, buscar por destino
                if not selected_row and target_clean:
                    for r_i, r_data in rows_dict.items():
                        if r_data.get("importe") is not None and target_clean == r_data.get("kunwe", "").upper():
                            selected_row = r_data
                            break

                # Fallback: primera fila que tenga importe
                if not selected_row:
                    for r_i in sorted(rows_dict.keys()):
                        if rows_dict[r_i].get("importe") is not None:
                            selected_row = rows_dict[r_i]
                            break

            if selected_row:
                data.update(selected_row)
                data["encontrado"] = True

        except Exception as e:
            self.log(f"Aviso leyendo Table Control VK13: {e}")

        return data

    # -----------------------------------------------------------------------
    # Flujo Z0116_TRAOPE (CTO COMPRA)
    # -----------------------------------------------------------------------
    def _precargar_z0116(self, items):
        """Abre Z0116_TRAOPE (o Z00116_TRAOPE) y extrae todos los registros usando pantalla o exportación."""
        try:
            self.log("Abriendo /nZ0116_TRAOPE para consulta global de Contratos de Compra...")
            session = self.session
            session.findById("wnd[0]/tbar[0]/okcd").text = "/nZ0116_TRAOPE"
            session.findById("wnd[0]").sendVKey(0)
            self._sleep(0.4)
            handle_popups(session)

            sbar = session.findById("wnd[0]/sbar")
            if sbar and sbar.messageType in ("E", "A") and "no existe" in sbar.text.lower():
                self.log("Transacción Z0116_TRAOPE no existe. Intentando /nZ00116_TRAOPE...")
                session.findById("wnd[0]/tbar[0]/okcd").text = "/nZ00116_TRAOPE"
                session.findById("wnd[0]").sendVKey(0)
                self._sleep(0.4)
                handle_popups(session)

            unique_plants = list(dict.fromkeys([str(x.get("cedis", "")).strip() for x in items if x.get("cedis")]))
            unique_vendors = list(dict.fromkeys([str(x.get("sf", "") or x.get("proveedor", "")).strip() for x in items if (x.get("sf") or x.get("proveedor"))]))
            unique_materials = list(dict.fromkeys([str(x.get("material", "")).strip() for x in items if x.get("material")]))
            unique_contracts = list(dict.fromkeys([str(x.get("contrato", "")).strip() for x in items if x.get("contrato") and not any(k in str(x.get("contrato")).upper() for k in ["CREAR", "ALTA", "POR CREAR"])]))

            plant_param = unique_plants[0] if unique_plants else ""
            vendor_param = unique_vendors[0] if len(unique_vendors) == 1 else ""
            mat_param = unique_materials[0] if len(unique_materials) == 1 else ""

            self.log(f"Configurando filtros: Centro={plant_param}, {len(unique_vendors)} proveedor(es), {len(unique_materials)} material(es)...")

            fill_z0116_selection(
                session, plant=plant_param, vendor=vendor_param, material=mat_param,
                vendors_list=unique_vendors, materials_list=unique_materials, contracts_list=unique_contracts
            )

            self.log("Ejecutando consulta Z0116_TRAOPE (F8)...")
            try:
                session.findById("wnd[0]/tbar[1]/btn[8]").press()
            except Exception:
                session.findById("wnd[0]").sendVKey(8)
            self._sleep(0.8)
            handle_popups(session)

            # 1. Intentar primero extraer directamente de la pantalla (instantáneo, ALV Grid o Lista Clásica)
            records = extract_all_z0116_from_screen(session)
            if records:
                self.log(f"Extraídos {len(records)} registro(s) directamente de la pantalla SAP.")
            else:
                self.log("No se detectaron registros en pantalla; ejecutando exportación automática como respaldo...")
                records = export_and_read_z0116_excel(session)
                self.log(f"Extraídos {len(records)} registro(s) de SAP vía exportación con tabulador.")

            return records
        except Exception as e:
            self.log(f"Aviso en precarga Z0116 (continuará validación fila por fila): {e}")
            return []

    def _validar_fila_cto_compra(self, item, preloaded_records):
        """Valida una fila de Contrato de Compra contra registros precargados o consulta individual en SAP."""
        contrato_esp = str(item.get("contrato") or "").strip()
        if any(k in contrato_esp.upper() for k in ["CREAR", "ALTA", "POR CREAR"]):
            contrato_esp = ""
        pos_esp = str(item.get("posicion") or "").strip()
        destino_esp = str(item.get("destino") or "").strip().upper()
        mat_esp = str(item.get("material") or "").strip()
        mod_esp = clean_exp_cond(item.get("modalidad") or "01")
        vendor_esp = str(item.get("sf") or item.get("proveedor") or "").strip()
        plant_esp = str(item.get("cedis") or "").strip()
        precio_esp = clean_num(item.get("precio_mat"))
        solicitud = str(item.get("solicitud") or "").upper()
        comentario = str(item.get("comentario") or "").upper()
        es_baja = ("BAJA" in solicitud or "BAJA" in comentario)

        self.log(f"[Fila {item.get('fila_origen', '')}] Validando Contrato Compra ({contrato_esp or 'Por crear'}) para Prov {vendor_esp}, Centro {plant_esp}, Mat {mat_esp}, Destino {destino_esp}, Pos {pos_esp}...")

        res_cto = validate_z0116_contract(
            self.session, plant=plant_esp, vendor=vendor_esp, material=mat_esp,
            destino_kunnr=destino_esp, exp_cond_excel=mod_esp, posicion_excel=pos_esp,
            contrato_excel=contrato_esp, preloaded_records=preloaded_records,
            precio_excel=precio_esp
        )

        if not res_cto.get("encontrado"):
            if es_baja:
                return {
                    "estado": "OK",
                    "mensaje": f"Contrato SAP {contrato_esp} Pos {pos_esp or 'N/A'} OK (Línea dada de baja / Sin registro activo en SAP)",
                    "datos_sap": {
                        "contrato": contrato_esp,
                        "posicion": pos_esp,
                        "precio": precio_esp,
                        "moneda": "MXN",
                        "um": item.get("um", ""),
                        "exp_cond": mod_esp,
                        "destino": destino_esp,
                        "kunnr": destino_esp,
                        "res_cto_raw": res_cto
                    }
                }
            return {
                "estado": "NO_ENCONTRADO",
                "mensaje": f"Contrato {contrato_esp or '(Por crear)'} no localizado en SAP Z0116 ({res_cto.get('detalle', 'Pendiente de alta')})",
                "datos_sap": {}
            }

        # Si fue encontrado en SAP pero era una solicitud de BAJA
        if es_baja:
            return {
                "estado": "DISCREPANCIA",
                "mensaje": f"DISCREPANCIA: Solicitud de BAJA pero la posición {res_cto.get('posicion')} sigue activa en SAP",
                "datos_sap": {
                    "contrato": res_cto.get("contrato", ""),
                    "posicion": res_cto.get("posicion", ""),
                    "precio": clean_num(res_cto.get("precio_neto")),
                    "moneda": res_cto.get("moneda", "MXN"),
                    "um": res_cto.get("um", "TN"),
                    "exp_cond": res_cto.get("exp_cond_sap", ""),
                    "destino": res_cto.get("kunnr_sap", ""),
                    "kunnr": res_cto.get("kunnr_sap", ""),
                    "res_cto_raw": res_cto
                }
            }

        val_desc = []
        if contrato_esp and res_cto.get("contrato") and str(contrato_esp).strip() != str(res_cto["contrato"]).strip():
            val_desc.append(f"DIFIERE CONTRATO (SAP:{res_cto['contrato']} vs Archivo:{contrato_esp})")

        # Verificar posición si venía en el archivo
        if pos_esp and res_cto.get("posicion"):
            try:
                p_int_esp = int(float(pos_esp))
                p_int_sap = int(float(res_cto["posicion"]))
                if p_int_esp != p_int_sap:
                    val_desc.append(f"DIFIERE POSICION (SAP:{res_cto['posicion']} vs Archivo:{pos_esp})")
            except Exception:
                if str(pos_esp).strip() != str(res_cto["posicion"]).strip():
                    val_desc.append(f"DIFIERE POSICION (SAP:{res_cto['posicion']} vs Archivo:{pos_esp})")

        p_sap = clean_num(res_cto.get("precio_neto"))
        costo_tot_esp = clean_num(item.get("costo_total_compra"))
        mp_esp = clean_num(item.get("mp_compra"))
        fte_esp = clean_num(item.get("flete_compra"))

        if p_sap is not None:
            precio_coincide = False
            detalle_precio = ""

            # Caso 1: Coincide con Costo Total TRAOPE (Neto pedido en SAP = MP + Flete)
            if costo_tot_esp is not None and abs(p_sap - costo_tot_esp) <= 0.05:
                precio_coincide = True
                detalle_precio = f"PRECIO COINCIDE (Costo Total ${p_sap:,.2f})"

            # Caso 2: Coincide con la suma explícita de MP + Flete
            elif mp_esp is not None and fte_esp is not None and abs(p_sap - (mp_esp + fte_esp)) <= 0.05:
                precio_coincide = True
                detalle_precio = f"PRECIO COINCIDE (MP+Fte ${p_sap:,.2f})"

            # Caso 3: Coincide con MP compra (contratos donde el flete es independiente)
            elif mp_esp is not None and abs(p_sap - mp_esp) <= 0.05:
                precio_coincide = True
                detalle_precio = f"PRECIO COINCIDE (MP ${p_sap:,.2f})"

            # Caso 4: Coincide con precio_esp general
            elif precio_esp is not None and abs(p_sap - precio_esp) <= 0.05:
                precio_coincide = True
                detalle_precio = f"PRECIO COINCIDE (${p_sap:,.2f})"

            if precio_coincide:
                val_desc.append(detalle_precio)
            else:
                comp_ref = f"CostoTot:${costo_tot_esp:,.2f}" if costo_tot_esp is not None else ""
                if mp_esp is not None:
                    comp_ref += f" / MP:${mp_esp:,.2f}" if comp_ref else f"MP:${mp_esp:,.2f}"
                ref_txt = comp_ref or (f"${precio_esp:,.2f}" if precio_esp is not None else "N/A")
                val_desc.append(f"DIFIERE PRECIO (SAP:${p_sap:,.2f} vs Archivo {ref_txt})")
        elif p_sap is not None:
            val_desc.append(f"PRECIO SAP (${p_sap:,.2f})")

        # Validaciones de TRAOPE y Z1017 exclusivas para pestaña 'Validación SAP GUI' (no Matriz)
        if item.get("hoja") != "Matriz":
            # 1. Campos obligatorios fijos en TRAOPE
            ind_oblig = str(res_cto.get("ind_oblig") or "").strip().upper()
            ctrl_ctr = str(res_cto.get("ctrl_ctr") or "").strip()

            if ind_oblig == "X":
                val_desc.append("TRAOPE IND.OBLIG OK (X)")
            else:
                val_desc.append(f"ERROR TRAOPE: 'Ind. Oblig.'='{ind_oblig or 'VACÍO'}' (esperado 'X')")

            if ctrl_ctr.zfill(4) == "0004" or ctrl_ctr == "0004":
                val_desc.append("TRAOPE CTRL-CTR OK (0004)")
            else:
                val_desc.append(f"ERROR TRAOPE: 'Ctrl-Ctr'='{ctrl_ctr or 'VACÍO'}' (esperado '0004')")

            # 2. Validación de Interlocutor y Extensión de Proveedor (Z1017_PROV_INT)
            sf_lookup = str(res_cto.get("shipfrom") or res_cto.get("vendor") or item.get("sf") or item.get("vendor") or vendor_esp).strip().lstrip("0")
            prov_traope = str(res_cto.get("proveedor") or item.get("proveedor") or "").strip().lstrip("0")

            if sf_lookup and prov_traope:
                res_z1017 = self._consultar_z1017_prov_int(sf_lookup)
                if res_z1017.get("encontrado"):
                    provs_ef = res_z1017.get("provs_ef", [])
                    if prov_traope in provs_ef:
                        val_desc.append(f"INTERLOCUTOR Z1017 OK (Prov EF {prov_traope} para SF {sf_lookup})")
                    else:
                        val_desc.append(f"ERROR INTERLOCUTOR Z1017: Prov TRAOPE ({prov_traope}) != Prov EF ({', '.join(set(provs_ef))}) para SF {sf_lookup}")
                else:
                    val_desc.append(f"ERROR INTERLOCUTOR Z1017: Sin extensión de acreedor para SF {sf_lookup} ({res_z1017.get('detalle')})")
            elif not sf_lookup:
                val_desc.append("ERROR INTERLOCUTOR Z1017: SF/Vendor no identificado")
            elif not prov_traope:
                val_desc.append("ERROR INTERLOCUTOR Z1017: Proveedor no identificado en TRAOPE")

        # 3. Aviso de Desempate Multilínea (si se detectaron múltiples posiciones candidatas)
        aviso_multi = res_cto.get("aviso_multilinea")
        if aviso_multi:
            val_desc.append(aviso_multi)
            self.log(f"[Fila {item.get('fila_origen', '')}] {aviso_multi}")

        if any("ERROR" in d or "DIFIERE" in d for d in val_desc):
            estado = "DISCREPANCIA"
            msg = " | ".join(val_desc)
        else:
            estado = "OK"
            msg = f"Contrato SAP {res_cto.get('contrato')} Pos {res_cto.get('posicion')} OK ({' | '.join(val_desc)})"

        return {
            "estado": estado,
            "mensaje": msg,
            "datos_sap": {
                "contrato": res_cto.get("contrato", ""),
                "posicion": res_cto.get("posicion", ""),
                "precio": p_sap if p_sap is not None else 0.0,
                "moneda": res_cto.get("moneda", "MXN"),
                "um": res_cto.get("um", "TN"),
                "fch_ini": res_cto.get("fch_ini", ""),
                "fch_fin": res_cto.get("fch_fin", ""),
                "exp_cond": res_cto.get("exp_cond_sap", ""),
                "destino": res_cto.get("kunnr_sap", ""),
                "kunnr": res_cto.get("kunnr_sap", ""),
                "ind_oblig": res_cto.get("ind_oblig", ""),
                "ctrl_ctr": res_cto.get("ctrl_ctr", ""),
                "proveedor": res_cto.get("proveedor", ""),
                "shipfrom": res_cto.get("shipfrom", ""),
                "aviso_multilinea": res_cto.get("aviso_multilinea", ""),
                "res_cto_raw": res_cto
            }
        }

    # -----------------------------------------------------------------------
    # Flujo CTO VENTA (ZSDD4501 / VA43 / VA45 / ZCP)
    # Transacción oficial ZSDD4501 (válida tanto para Filial como para Terceros)
    # -----------------------------------------------------------------------
    def _precargar_zsdd4501(self, cto_venta_items):
        """Ejecuta una consulta en ZSDD4501 para precargar contratos de venta de la hoja."""
        try:
            session = self.session
            if not cto_venta_items:
                return []

            first = cto_venta_items[0]
            vkorg = str(first.get("vkorg") or "7180").strip()
            centro = str(first.get("cedis") or "").strip()
            destino = str(first.get("destino") or "").strip().lstrip("0")
            material = str(first.get("material") or "").strip()
            sf = str(first.get("sf") or first.get("proveedor") or "").strip()
            dest_factura = str(first.get("dest_factura") or "").strip()

            contratos_list = [str(it.get("contrato")).strip() for it in cto_venta_items if it.get("contrato") and not any(k in str(it.get("contrato")).upper() for k in ["CREAR", "ALTA", "POR CREAR"])]

            unique_sfs = {str(it.get("sf") or it.get("proveedor") or "").strip() for it in cto_venta_items if (it.get("sf") or it.get("proveedor"))}
            unique_destinos = {str(it.get("destino") or "").strip() for it in cto_venta_items if it.get("destino")}
            unique_centros = {str(it.get("cedis") or "").strip() for it in cto_venta_items if it.get("cedis")}
            unique_mats = {str(it.get("material") or "").strip() for it in cto_venta_items if it.get("material")}

            sf_filter = list(unique_sfs)[0] if len(unique_sfs) == 1 else ""
            dest_filter = list(unique_destinos)[0] if len(unique_destinos) == 1 else ""
            centro_filter = list(unique_centros)[0] if len(unique_centros) == 1 else ""
            mat_filter = list(unique_mats)[0] if len(unique_mats) == 1 else ""

            self.log(f"Iniciando consulta global de Contratos de Venta en ZSDD4501 (Org {vkorg}, Centro {centro_filter or centro}, SF/Factura {sf_filter or sf}, Destino {dest_filter or destino})...")

            # 1. Ingresar a ZSDD4501
            session.findById("wnd[0]/tbar[0]/okcd").text = "/nZSDD4501"
            session.findById("wnd[0]").sendVKey(0)
            self._sleep(0.3)
            handle_popups(session)

            sbar = None
            try: sbar = session.findById("wnd[0]/sbar")
            except Exception: pass
            if sbar and sbar.messageType in ("E", "A") and ("no existe" in sbar.text.lower() or "not exist" in sbar.text.lower()):
                self.log(f"Aviso: ZSDD4501 no existe en este ambiente ({sbar.text}). Usando respaldo.")
                return []

            # 2. Rellenar criterios
            fill_zsdd4501_selection(
                session, vkorg=vkorg,
                destino=dest_filter or (destino if len(cto_venta_items) == 1 else ""),
                material=mat_filter or (material if len(cto_venta_items) == 1 else ""),
                centro=centro_filter or centro,
                contrato=contratos_list[0] if len(contratos_list) == 1 else "",
                shipfrom=sf_filter or (sf if len(cto_venta_items) == 1 else ""),
                dest_factura=dest_factura
            )

            # 3. Ejecutar F8
            try:
                session.findById("wnd[0]/tbar[1]/btn[8]").press()
            except Exception:
                session.findById("wnd[0]").sendVKey(8)
            self._sleep(0.8)
            handle_popups(session)

            # 4. Extraer registros
            # 4. Extraer registros (pantalla directa o exportación automática de respaldo)
            records = extract_all_zsdd4501_from_screen(session)
            if not records:
                self.log("Intentando exportación de respaldo en ZSDD4501 (como en ctos compra)...")
                records = export_and_read_zsdd4501_excel(session)

            if records:
                self.log(f"Extraídos {len(records)} registro(s) de contratos de venta desde ZSDD4501.")
            else:
                self.log("Aviso: ZSDD4501 no devolvió registros tras consulta ni exportación.")

            return records
        except Exception as e:
            self.log(f"Aviso en precarga ZSDD4501: {e}")
            return []

    def _consultar_zsdd4501(self, vkorg, destino, material, centro=None, shipfrom=None, contrato=None, dest_factura=None, posicion=None, precio_esp=None, fch_ini_esp=None, exp_cond_esp=None, contrato_hint=""):
        """Consulta individual de contrato de venta en transacción ZSDD4501 con desempate inteligente y detección de contratos múltiples."""
        session = self.session
        res = {
            "encontrado": False,
            "contrato": "",
            "posicion": "",
            "precio_neto": None,
            "moneda": "MXN",
            "um": "",
            "fch_ini": "",
            "fch_fin": "",
            "exp_cond": "",
            "kunnr": "",
            "cliente": "",
            "material": "",
            "centro": "",
            "detalle": "",
            "multiples_contratos": [],
            "avisos": []
        }

        try:
            session.findById("wnd[0]/tbar[0]/okcd").text = "/nZSDD4501"
            session.findById("wnd[0]").sendVKey(0)
            self._sleep(0.3)
            handle_popups(session)

            sbar = None
            try: sbar = session.findById("wnd[0]/sbar")
            except Exception: pass
            if sbar and sbar.messageType in ("E", "A") and ("no existe" in sbar.text.lower() or "not exist" in sbar.text.lower()):
                self.log(f"Aviso: ZSDD4501 no existe ({sbar.text}).")
                return None

            c_target = str(contrato).strip() if contrato and not any(k in str(contrato).upper() for k in ["CREAR", "ALTA", "POR CREAR"]) else ""
            pos_target = str(posicion).strip() if posicion else ""
            m_target = str(material).strip()
            d_target = str(destino).strip().lstrip("0")

            # Consultar en pantalla de selección de ZSDD4501
            # Si el contrato viene especificado en el Excel, se consulta preferentemente con él; si no, por Destino y Material
            fill_zsdd4501_selection(
                session, vkorg=vkorg, destino=destino, material=material,
                centro=centro, contrato=c_target, shipfrom=shipfrom,
                dest_factura=dest_factura
            )

            try:
                session.findById("wnd[0]/tbar[1]/btn[8]").press()
            except Exception:
                session.findById("wnd[0]").sendVKey(8)
            self._sleep(0.8)
            handle_popups(session)

            try: sbar = session.findById("wnd[0]/sbar")
            except Exception: pass
            sbar_msg = sbar.text.strip() if sbar else ""

            records = extract_all_zsdd4501_from_screen(session)
            if not records:
                records = export_and_read_zsdd4501_excel(session)

            # Si no arrojó registros y se filtró por contrato, intentar una segunda búsqueda abierta por Destino y Material
            if not records and c_target:
                try:
                    session.findById("wnd[0]/tbar[0]/okcd").text = "/nZSDD4501"
                    session.findById("wnd[0]").sendVKey(0)
                    self._sleep(0.3)
                    handle_popups(session)
                    fill_zsdd4501_selection(
                        session, vkorg=vkorg, destino=destino, material=material,
                        centro=centro, contrato="", shipfrom=shipfrom,
                        dest_factura=dest_factura
                    )
                    try:
                        session.findById("wnd[0]/tbar[1]/btn[8]").press()
                    except Exception:
                        session.findById("wnd[0]").sendVKey(8)
                    self._sleep(0.8)
                    handle_popups(session)
                    records = extract_all_zsdd4501_from_screen(session)
                    if not records:
                        records = export_and_read_zsdd4501_excel(session)
                except Exception:
                    pass

            if not records:
                res["detalle"] = f"ZSDD4501: Sin registros encontrados ({sbar_msg})"
                return res

            matched = None
            multiples_contratos_detectados = []

            # Coincidencia por Destino y Material con desempate multicriterio
            if d_target and m_target:
                cand_matches = [
                    r for r in records
                    if dest_matches(r, d_target) and mat_matches(r.get("material"), m_target)
                ]
                if cand_matches:
                    distinct_ctos = list(dict.fromkeys(
                        str(r.get("contrato", "")).strip()
                        for r in cand_matches if str(r.get("contrato", "")).strip()
                    ))
                    if len(distinct_ctos) > 1:
                        multiples_contratos_detectados = distinct_ctos
                        self.log(
                            f"Aviso: Detectados {len(distinct_ctos)} contratos distintos ({', '.join(distinct_ctos)}) "
                            f"para Destino {d_target} y Material {m_target}. Aplicando desempate multilínea..."
                        )

                    def _score_candidato(r):
                        score = 0.0
                        c_num_str = str(r.get("contrato", "")).strip()
                        pos_num_str = str(r.get("posicion", "")).strip()
                        p_val = clean_num(r.get("precio_neto"))
                        r_ini = str(r.get("fch_ini", "")).strip()
                        r_fin = str(r.get("fch_fin", "")).strip()
                        exp_r = clean_exp_cond(r.get("exp_cond") or "")

                        # A) Prioridad 1: Coincidencia con contrato escrito en el Excel
                        if c_target and c_num_str == c_target:
                            score += 800.0
                            # Si además coincide la posición específica del Excel
                            if pos_target and pos_matches(pos_num_str, pos_target):
                                score += 300.0

                        # B) Prioridad 2: Contrato Hint (ya descubierto en el lote)
                        if contrato_hint and c_num_str == str(contrato_hint).strip():
                            score += 600.0

                        # C) Prioridad 3: Coincidencia con precio esperado (FPA)
                        if precio_esp is not None and p_val is not None:
                            if abs(p_val - precio_esp) <= 0.05:
                                score += 500.0
                            else:
                                score -= 200.0

                        # D) Prioridad 4: Coincidencia de Modalidad / Exp Cond (FPA)
                        if exp_cond_esp and exp_r:
                            if exp_r == clean_exp_cond(exp_cond_esp):
                                score += 300.0

                        # E) Prioridad 5: Coincidencia o cercanía de vigencia inicial (FPA)
                        if fch_ini_esp and r_ini:
                            if r_ini == fch_ini_esp:
                                score += 250.0
                            elif r_ini.split(".")[-1] == fch_ini_esp.split(".")[-1]:
                                score += 100.0

                        # F) Prioridad 6: Contrato vigente hacia el futuro (año fin >= 2026 o 9999)
                        if r_fin:
                            try:
                                anio_fin = int(r_fin.split(".")[-1])
                                if anio_fin >= 2026:
                                    score += 150.0
                            except Exception:
                                pass

                        # G) Desempate por mayor número de contrato SAP (ID numérico más reciente)
                        try:
                            score += int(c_num_str) / 10000000.0
                        except Exception:
                            pass

                        return score

                    cand_matches.sort(key=_score_candidato, reverse=True)
                    matched = cand_matches[0]
                    res["multiples_contratos"] = multiples_contratos_detectados
                    if len(distinct_ctos) > 1:
                        res["avisos"].append(
                            f"Múltiples contratos SAP para Destino {d_target} y Mat {m_target} ({', '.join(distinct_ctos)}) - Se seleccionó {matched.get('contrato')}"
                        )

            # 4. Coincidencia por Contrato si solo hay una línea para este contrato
            if not matched and c_target:
                matches_cto = [r for r in records if str(r.get("contrato", "")).strip() == c_target]
                if len(matches_cto) == 1:
                    matched = matches_cto[0]
                elif matches_cto:
                    # Si hay varias, preferir la que coincida en material o posición
                    m_cand = [r for r in matches_cto if mat_matches(r.get("material"), m_target)]
                    matched = m_cand[0] if m_cand else matches_cto[0]

            # 5. Coincidencia por Destino
            if not matched and d_target:
                for r in records:
                    if dest_matches(r, d_target):
                        matched = r
                        break

            # 6. Coincidencia por Material
            if not matched and m_target and not d_target:
                for r in records:
                    r_mat = str(r.get("material", "")).strip()
                    if r_mat == m_target:
                        matched = r
                        break

            # 7. Si hay un único registro
            if not matched and len(records) == 1:
                matched = records[0]

            if matched:
                # Si la UM no vino en ZSDD4501, consultar transacción VA43 (campo VBAP-ZIEME)
                if not matched.get("um") and matched.get("contrato"):
                    um_va = self._extraer_um_va43(
                        contrato=matched.get("contrato"),
                        posicion=matched.get("posicion"),
                        material=matched.get("material")
                    )
                    if um_va:
                        matched["um"] = um_va

                res["encontrado"] = True
                res["contrato"] = matched.get("contrato", "")
                res["posicion"] = matched.get("posicion", "")
                res["precio_neto"] = matched.get("precio_neto")
                res["moneda"] = matched.get("moneda", "MXN")
                res["um"] = matched.get("um", "")
                res["fch_ini"] = matched.get("fch_ini", "")
                res["fch_fin"] = matched.get("fch_fin", "")
                res["exp_cond"] = matched.get("exp_cond", "")
                res["kunnr"] = matched.get("dest_factura", "") or matched.get("kunnr", "")
                res["cliente"] = matched.get("dest_mercancia", "") or matched.get("cliente", "")
                res["material"] = matched.get("material", "")
                res["centro"] = matched.get("centro", "")
                res["multiples_contratos"] = multiples_contratos_detectados
                res["detalle"] = f"ZSDD4501: Cto={res['contrato']} Pos={res['posicion']} | Precio={res['precio_neto']} {res['moneda']} | UM={res['um']} | Destino={res['cliente']}"
            else:
                res["detalle"] = f"ZSDD4501: Registros hallados ({len(records)}) pero no coincidieron con Destino {destino} o Material {material}"

            return res
        except Exception as e:
            self.log(f"Aviso en consulta individual ZSDD4501: {e}")
            return None
        finally:
            try:
                session.findById("wnd[0]/tbar[0]/okcd").text = "/n"
                session.findById("wnd[0]").sendVKey(0)
                self._sleep(0.2)
            except Exception:
                pass

    def _consultar_z1017_prov_int(self, shipfrom):
        """Consulta y extrae los registros de interlocutores en Z1017_PROV_INT para el vendor/shipfrom indicado."""
        sf_clean = str(shipfrom or "").strip().lstrip("0")
        if not sf_clean:
            return {"encontrado": False, "provs_ef": [], "detalle": "Sin Ship From"}

        if not hasattr(self, "_cache_z1017"):
            self._cache_z1017 = {}
        if sf_clean in self._cache_z1017:
            return self._cache_z1017[sf_clean]

        session = getattr(self, "session", None)
        if not session:
            try:
                self.conectar()
                session = self.session
            except Exception as e:
                return {"encontrado": False, "provs_ef": [], "detalle": f"No se pudo conectar a SAP: {e}"}

        self.log(f"Consultando Interlocutor Z1017_PROV_INT para Vendor {sf_clean}...")
        try:
            records = export_and_read_z1017_excel(session, shipfrom=sf_clean)
            if not records:
                res = {"encontrado": False, "provs_ef": [], "detalle": f"Sin registros en Z1017 para {sf_clean}"}
            else:
                provs_ef = []
                for r in records:
                    ac = str(r.get("acreedor") or "").strip().lstrip("0")
                    pef = str(r.get("prov_ef") or "").strip().lstrip("0")
                    if pef and (not ac or ac == sf_clean):
                        provs_ef.append(pef)
                res = {
                    "encontrado": len(provs_ef) > 0,
                    "provs_ef": provs_ef,
                    "records": records,
                    "detalle": f"Localizados {len(provs_ef)} Prov EF para {sf_clean}: {', '.join(set(provs_ef))}"
                }
            self._cache_z1017[sf_clean] = res
            return res
        except Exception as e:
            self.log(f"Aviso consultando Z1017_PROV_INT: {e}")
            res = {"encontrado": False, "provs_ef": [], "detalle": str(e)}
            self._cache_z1017[sf_clean] = res
            return res

    def _validar_fila_cto_venta(self, item, preloaded_records=None):
        """Valida una fila de Contrato de Venta (ZCP) comparando exhaustivamente los campos descargados de SAP ZSDD4501."""
        contrato_raw = str(item.get("contrato") or "").strip()
        if any(k in contrato_raw.upper() for k in ["CREAR", "ALTA", "POR CREAR", "NONE", "NAN", "S/C"]):
            contrato_raw = ""

        contrato_esp = contrato_raw.split("-")[0].strip() if "-" in contrato_raw else contrato_raw
        pos_esp = str(item.get("posicion") or "").strip()
        if not pos_esp and "-" in contrato_raw:
            pos_esp = contrato_raw.split("-")[1].strip()

        destino_esp = str(item.get("destino") or "").strip().lstrip("0")
        sf_esp = str(item.get("sf") or item.get("proveedor") or "").strip().lstrip("0")
        mat_esp = str(item.get("material") or "").strip()
        cedis_esp = str(item.get("cedis") or "").strip()
        mod_esp = clean_exp_cond(item.get("modalidad") or "01")
        p_cto_dir = clean_num(item.get("precio_venta_cto") or item.get("precio_contrato") or item.get("precio_neto") or item.get("precio"))
        p_mp = clean_num(item.get("precio_mat") or item.get("mp"))
        p_fte = clean_num(item.get("precio_flete") or item.get("precio_fte") or item.get("flete"))
        pv = clean_num(item.get("pv"))

        if p_cto_dir is not None:
            precio_esp = p_cto_dir
        elif p_mp is not None and p_fte is not None:
            precio_esp = p_mp + p_fte  # En contratos de venta el precio es neto: suma de MP y Flete
        elif p_mp is not None:
            precio_esp = p_mp
        else:
            precio_esp = None

        suma_mp_fte = (p_mp + p_fte) if (p_mp is not None and p_fte is not None) else p_mp
        um_esp = clean_um(item.get("um") or "")
        fch_ini_esp = str(item.get("fch_ini") or "").strip()
        fch_fin_esp = str(item.get("fch_fin") or "").strip()
        moneda_esp = str(item.get("moneda") or "MXN").strip()
        vkorg = str(item.get("vkorg") or "7180").strip()
        solicitud = str(item.get("solicitud") or "").upper()
        comentario = str(item.get("comentario") or "").upper()
        es_baja = ("BAJA" in solicitud or "BAJA" in comentario)
        es_alta = ("ALTA" in solicitud or not contrato_esp)

        self.log(f"[Fila {item.get('fila_origen', '')}] Validando Contrato Venta ({contrato_esp or 'Por crear'}) para Cliente {destino_esp}, SF {sf_esp}, Mat {mat_esp}, Centro {cedis_esp}...")

        # 1. Búsqueda y selección inteligente en ZSDD4501 por Destino y Material (con cruce de FPA)
        if es_baja:
            return {
                "estado": "OK",
                "mensaje": f"Contrato Venta OK (Dado de baja / Sin registro activo en SAP)",
                "datos_sap": {
                    "contrato": "", "posicion": pos_esp, "precio": precio_esp,
                    "moneda": moneda_esp, "um": um_esp, "exp_cond": mod_esp,
                    "destino": destino_esp, "shipfrom": sf_esp, "material": mat_esp, "centro": cedis_esp
                }
            }

        self.log(f"[Fila {item.get('fila_origen', '')}] Consultando ZSDD4501 para Cliente {destino_esp}, Material {mat_esp}, SF {sf_esp} (priorizando contrato Excel: {contrato_esp or 'N/A'})...")
        
        # Buscar contrato_hint del lote (si otra fila del mismo destino ya lo descubrió)
        c_hint = getattr(self, "_contratos_vta_lote", {}).get(destino_esp, "")
        if not c_hint and sf_esp:
            c_hint = getattr(self, "_contratos_vta_lote", {}).get(f"{destino_esp}_{sf_esp}", "")

        res_zsdd = self._consultar_zsdd4501(
            vkorg=vkorg, destino=destino_esp, material=mat_esp,
            centro=cedis_esp, shipfrom=sf_esp, contrato=contrato_esp, dest_factura=sf_esp,
            posicion=pos_esp, precio_esp=precio_esp, fch_ini_esp=fch_ini_esp, exp_cond_esp=mod_esp,
            contrato_hint=c_hint
        )
        contrato_encontrado = res_zsdd.get("contrato") if res_zsdd else ""
        if res_zsdd and res_zsdd.get("encontrado") and contrato_encontrado:
            c_sap = str(contrato_encontrado).strip()
            pos_sap = str(res_zsdd.get("posicion") or pos_esp or "").strip()
            m_sap = str(res_zsdd.get("material") or "").strip()
            p_sap = clean_num(res_zsdd.get("precio_neto"))
            um_sap = clean_um(res_zsdd.get("um") or "")
            mon_sap = str(res_zsdd.get("moneda") or moneda_esp or "MXN").strip()
            d_sap = str(res_zsdd.get("cliente") or res_zsdd.get("kunnr") or "").strip().lstrip("0")

            # Actualizar contrato descubierto en el lote
            if not hasattr(self, "_contratos_vta_lote"):
                self._contratos_vta_lote = {}
            if destino_esp:
                self._contratos_vta_lote[destino_esp] = c_sap
                if sf_esp:
                    self._contratos_vta_lote[f"{destino_esp}_{sf_esp}"] = c_sap

            avisos_totales = list(res_zsdd.get("avisos", []))

            # Homologación y confirmación con VA43 detallado (btnBT_POPO + TableControl por material)
            # para verificar la partida exacta y el precio real directamente en VA43
            try:
                res_va43 = self._consultar_va43_detallado(contrato=c_sap, posicion=pos_sap, material=mat_esp)
                if res_va43 and res_va43.get("encontrado"):
                    va_pos = str(res_va43.get("posicion") or "").strip()
                    va_mat = str(res_va43.get("material") or "").strip()

                    # Adoptar posición de VA43 si coincide con el material o no había posición previa de ZSDD
                    if va_pos:
                        if (not pos_sap and (not mat_esp or not va_mat or mat_matches(va_mat, mat_esp))) or pos_matches(va_pos, pos_sap) or (mat_esp and mat_matches(va_mat, mat_esp)):
                            pos_sap = va_pos
                        elif mat_esp and va_mat and not mat_matches(va_mat, mat_esp):
                            self.log(f"Aviso VA43: Se descarta Posición {va_pos} de VA43 porque pertenece a material distinto ({va_mat} vs {mat_esp}); pos ZSDD={pos_sap or 'N/A'}")

                    if res_va43.get("precio_neto") is not None:
                        if not mat_esp or not va_mat or mat_matches(va_mat, mat_esp):
                            p_sap = clean_num(res_va43["precio_neto"])
                    if res_va43.get("um"):
                        if not mat_esp or not va_mat or mat_matches(va_mat, mat_esp):
                            um_sap = clean_um(res_va43["um"])
                    if va_mat and " " not in va_mat:
                        if not mat_esp or mat_matches(va_mat, mat_esp):
                            m_sap = va_mat
                    if res_va43.get("dest_mercancia") or res_va43.get("destino"):
                        d_sap = str(res_va43.get("dest_mercancia") or res_va43.get("destino")).strip().lstrip("0")
                    if res_va43.get("avisos"):
                        avisos_totales.extend(res_va43["avisos"])
            except Exception as e_va:
                self.log(f"Aviso consultando VA43 detallado tras ZSDD4501: {e_va}")

            val_desc = []

            # Contrato
            if contrato_esp:
                if contrato_esp == c_sap:
                    val_desc.append(f"CONTRATO OK ({c_sap})")
                else:
                    val_desc.append(f"DIFIERE CONTRATO (SAP:{c_sap} vs XL:{contrato_esp})")
            else:
                val_desc.append(f"CONTRATO CREADO ({c_sap})")

            # Posición
            if pos_esp and pos_sap:
                try:
                    p_int_esp = int(float(pos_esp))
                    p_int_sap = int(float(pos_sap))
                    if p_int_esp != p_int_sap:
                        val_desc.append(f"DIFIERE POSICION (SAP:{pos_sap} vs XL:{pos_esp})")
                    else:
                        val_desc.append(f"POS OK ({pos_sap})")
                except Exception:
                    if pos_esp != pos_sap:
                        val_desc.append(f"DIFIERE POSICION (SAP:{pos_sap} vs XL:{pos_esp})")
                    else:
                        val_desc.append(f"POS OK ({pos_sap})")
            elif pos_sap:
                val_desc.append(f"POS SAP ({pos_sap})")

            # Destino / Cliente
            if destino_esp and d_sap:
                if destino_esp == d_sap:
                    val_desc.append(f"DESTINO OK ({d_sap})")
                else:
                    val_desc.append(f"DIFIERE DESTINO (SAP:{d_sap} vs XL:{destino_esp})")

            # Material
            if mat_esp and m_sap:
                if mat_matches(m_sap, mat_esp):
                    val_desc.append(f"MATERIAL OK ({mat_esp})")
                elif " " in m_sap and mat_matches(item.get("material"), mat_esp):
                    val_desc.append(f"MATERIAL OK ({mat_esp})")
                else:
                    val_desc.append(f"DIFIERE MATERIAL (SAP:{m_sap} vs XL:{mat_esp})")

            # Unidad de Medida (UM)
            if not um_sap:
                val_desc.append("ERROR: NO SE DETECTÓ UM EN SAP (ZSDD4501/VA43)")
            elif um_esp and um_sap != um_esp:
                val_desc.append(f"DIFIERE UM (SAP:{um_sap} vs XL:{um_esp})")
            elif um_esp:
                val_desc.append(f"UM OK ({um_sap})")
            else:
                val_desc.append(f"UM SAP ({um_sap})")

            # Precio neto
            if p_sap is not None:
                precio_coincide = False
                detalle_precio = ""

                # Caso 1: Coincide directamente con el precio esperado
                if precio_esp is not None and abs(p_sap - precio_esp) <= 0.05:
                    precio_coincide = True
                    detalle_precio = f"PRECIO COINCIDE (${p_sap:,.2f})"
                # Caso 2: Coincide con la suma de MP + Flete
                elif suma_mp_fte is not None and abs(p_sap - suma_mp_fte) <= 0.05:
                    precio_coincide = True
                    detalle_precio = f"PRECIO NETO OK (MP+Fte: ${p_sap:,.2f})"
                # Caso 3: Coincide con factor PV
                elif pv and pv > 0:
                    if precio_esp is not None and abs(p_sap - (precio_esp * pv)) <= 0.05:
                        precio_coincide = True
                        detalle_precio = f"PRECIO OK CON PV ({precio_esp:,.2f}*{pv} = ${p_sap:,.2f})"
                    elif suma_mp_fte is not None and abs(p_sap - (suma_mp_fte * pv)) <= 0.05:
                        precio_coincide = True
                        detalle_precio = f"PRECIO NETO OK CON PV ({suma_mp_fte:,.2f}*{pv} = ${p_sap:,.2f})"
                    elif precio_esp is not None and abs(p_sap - (precio_esp / pv)) <= 0.05:
                        precio_coincide = True
                        detalle_precio = f"PRECIO OK CON PV ({precio_esp:,.2f}/{pv} = ${p_sap:,.2f})"
                    elif suma_mp_fte is not None and abs(p_sap - (suma_mp_fte / pv)) <= 0.05:
                        precio_coincide = True
                        detalle_precio = f"PRECIO NETO OK CON PV ({suma_mp_fte:,.2f}/{pv} = ${p_sap:,.2f})"
                # Caso 4: Recogido o solo MP
                elif (not p_fte or mod_esp == "02") and p_mp is not None and abs(p_sap - p_mp) <= 0.05:
                    precio_coincide = True
                    detalle_precio = f"PRECIO OK (MP Recogido: ${p_sap:,.2f})"

                if precio_coincide:
                    val_desc.append(detalle_precio)
                else:
                    p_ref = precio_esp if precio_esp is not None else (suma_mp_fte if suma_mp_fte is not None else 0.0)
                    val_desc.append(f"DIFIERE PRECIO (SAP:${p_sap:,.2f} vs XL:${p_ref:,.2f})")
            elif precio_esp is not None:
                val_desc.append(f"PRECIO XL (${precio_esp:,.2f})")

            # Aviso si se detectaron múltiples contratos en SAP para el mismo destino y material
            mult_ctos = res_zsdd.get("multiples_contratos", [])
            if mult_ctos and len(mult_ctos) > 1:
                val_desc.append(f"AVISO: Múltiples contratos detectados en SAP ({', '.join(mult_ctos)}) para Destino {destino_esp} y Mat {mat_esp} (Revisar)")

            # Avisos adicionales de SAP (bloqueos, advertencias de cliente)
            for av in avisos_totales:
                if av not in val_desc and "Múltiples contratos" not in av:
                    val_desc.append(f"Aviso SAP: {av}")

            # Estado global
            if any("ERROR" in d or "DIFIERE" in d for d in val_desc):
                estado = "DISCREPANCIA"
                pos_t = f" Pos {pos_sap}" if pos_sap else ""
                msg = f"DISCREPANCIA EN CONTRATO SAP {c_sap}{pos_t}: " + " | ".join(val_desc)
            else:
                estado = "OK"
                pos_t = f" Pos {pos_sap}" if pos_sap else ""
                c_tag = "ALTA Confirmada" if not contrato_esp else "VALIDADO"
                msg = f"Contrato Venta {c_tag} en SAP {c_sap}{pos_t} VALIDADO OK ({' | '.join(val_desc)})"

            return {
                "estado": estado,
                "mensaje": msg,
                "datos_sap": {
                    "contrato": c_sap,
                    "posicion": pos_sap,
                    "precio": p_sap,
                    "moneda": mon_sap,
                    "um": um_sap,
                    "fch_ini": res_zsdd.get("fch_ini", fch_ini_esp),
                    "fch_fin": res_zsdd.get("fch_fin", fch_fin_esp),
                    "exp_cond": res_zsdd.get("exp_cond", mod_esp),
                    "destino": d_sap or destino_esp,
                    "shipfrom": sf_esp,
                    "material": mat_esp if (m_sap and " " in m_sap) else (m_sap or mat_esp),
                    "centro": cedis_esp,
                    "avisos": val_desc
                }
            }
        elif not contrato_esp:
            return {
                "estado": "NO_ENCONTRADO",
                "mensaje": f"DISCREPANCIA: Solicitud de ALTA pero el Contrato Venta NO fue localizado en SAP (ZSDD4501)",
                "datos_sap": {
                    "contrato": "(No creado en SAP)",
                    "posicion": pos_esp or "10",
                    "precio": precio_esp,
                    "moneda": moneda_esp,
                    "um": um_esp,
                    "exp_cond": mod_esp,
                    "destino": destino_esp,
                    "shipfrom": sf_esp,
                    "material": mat_esp,
                    "centro": cedis_esp
                }
            }

        # 2. CASO CAMBIO / VALIDACIÓN: Consultar directamente en VA43 en SAP GUI (filtrando por posición vía btnBT_POPO + txtRV45A-POSNR)
        res_va43 = self._consultar_va43_detallado(contrato=contrato_esp, posicion=pos_esp, material=mat_esp)

        # 3. Evaluar si no se encontró en SAP VA43
        if not res_va43 or not res_va43.get("encontrado"):
            if es_baja:
                return {
                    "estado": "OK",
                    "mensaje": f"Contrato Venta {contrato_esp} OK (Dado de baja / No encontrado en VA43)",
                    "datos_sap": {
                        "contrato": contrato_esp, "posicion": pos_esp, "precio": precio_esp,
                        "moneda": moneda_esp, "um": um_esp, "exp_cond": mod_esp,
                        "destino": destino_esp, "shipfrom": sf_esp, "material": mat_esp, "centro": cedis_esp
                    }
                }
            else:
                err_msg = res_va43.get("error") if res_va43 else ""
                detalle_no_enc = f": {err_msg}" if err_msg else ""
                return {
                    "estado": "NO_ENCONTRADO",
                    "mensaje": f"Contrato Venta {contrato_esp} (Pos {pos_esp or 'N/A'}) no localizado en VA43{detalle_no_enc}",
                    "datos_sap": {
                        "contrato": contrato_esp, "posicion": pos_esp, "avisos": res_va43.get("avisos", []) if res_va43 else []
                    }
                }

        # Extraer campos detallados obtenidos directamente de VA43
        c_sap = str(res_va43.get("contrato") or contrato_esp).strip()
        pos_sap = str(res_va43.get("posicion") or pos_esp or "").strip()
        m_sap = str(res_va43.get("material") or "").strip()
        p_sap = clean_num(res_va43.get("precio_neto"))
        um_sap = clean_um(res_va43.get("um") or "")
        mon_sap = str(res_va43.get("moneda") or moneda_esp or "MXN").strip()
        avisos_sap = res_va43.get("avisos", [])
        d_sap = str(res_va43.get("dest_mercancia") or res_va43.get("destino") or res_va43.get("solicitante") or "").strip().lstrip("0")

        # Actualizar contrato descubierto en el lote
        if not hasattr(self, "_contratos_vta_lote"):
            self._contratos_vta_lote = {}
        if destino_esp and c_sap:
            self._contratos_vta_lote[destino_esp] = c_sap
            if sf_esp:
                self._contratos_vta_lote[f"{destino_esp}_{sf_esp}"] = c_sap

        # Si era solicitud de BAJA pero el contrato sigue activo en SAP
        if es_baja:
            return {
                "estado": "DISCREPANCIA",
                "mensaje": f"DISCREPANCIA: Solicitud de BAJA pero el Contrato Venta {c_sap} Pos {pos_sap} sigue activo en VA43",
                "datos_sap": {
                    "contrato": c_sap, "posicion": pos_sap, "precio": p_sap,
                    "moneda": mon_sap, "um": um_sap, "fch_ini": fch_ini_esp,
                    "fch_fin": fch_fin_esp, "exp_cond": mod_esp, "destino": d_sap,
                    "shipfrom": sf_esp, "material": m_sap, "centro": cedis_esp,
                    "avisos": avisos_sap
                }
            }

        # 4. Comparar campos uno a uno contra el archivo Excel (idéntico a CTO COMPRA)
        val_desc = []

        # Contrato
        if contrato_esp and c_sap and contrato_esp != c_sap:
            val_desc.append(f"DIFIERE CONTRATO (SAP:{c_sap} vs XL:{contrato_esp})")

        # Posición
        if pos_esp and pos_sap:
            try:
                p_int_esp = int(float(pos_esp))
                p_int_sap = int(float(pos_sap))
                if p_int_esp != p_int_sap:
                    val_desc.append(f"DIFIERE POSICION (SAP:{pos_sap} vs XL:{pos_esp})")
                else:
                    val_desc.append(f"POS OK ({pos_sap})")
            except Exception:
                if pos_esp != pos_sap:
                    val_desc.append(f"DIFIERE POSICION (SAP:{pos_sap} vs XL:{pos_esp})")
                else:
                    val_desc.append(f"POS OK ({pos_sap})")
        elif pos_sap:
            val_desc.append(f"POS SAP ({pos_sap})")
        elif pos_esp:
            val_desc.append(f"POS XL ({pos_esp})")

        # Destino / Cliente
        if destino_esp and d_sap:
            if destino_esp == d_sap:
                val_desc.append(f"DESTINO OK ({d_sap})")
            else:
                val_desc.append(f"DIFIERE DESTINO (SAP:{d_sap} vs XL:{destino_esp})")
        elif destino_esp:
            if any(destino_esp in str(av) for av in avisos_sap):
                val_desc.append(f"DESTINO OK ({destino_esp})")

        # Material
        if mat_esp and m_sap:
            if mat_matches(m_sap, mat_esp):
                val_desc.append(f"MATERIAL OK ({mat_esp})")
            elif " " in m_sap and mat_matches(item.get("material"), mat_esp):
                val_desc.append(f"MATERIAL OK ({mat_esp})")
            else:
                val_desc.append(f"DIFIERE MATERIAL (SAP:{m_sap} vs XL:{mat_esp})")

        # Unidad de Medida (UM) - VBAP-KMEIN en Dynpro 4906 / 4001
        if not um_sap:
            val_desc.append("ERROR: NO SE DETECTÓ UM EN SAP (VA43 VBAP-KMEIN)")
        elif um_esp and um_sap != um_esp:
            val_desc.append(f"DIFIERE UM (SAP:{um_sap} vs XL:{um_esp})")
        elif um_esp:
            val_desc.append(f"UM OK ({um_sap})")
        else:
            val_desc.append(f"UM SAP ({um_sap})")

        # Precio: El contrato de venta es NETO (la suma de MP y Flete, o precio explícito de contrato)
        if p_sap is not None:
            precio_coincide = False
            detalle_precio = ""

            # Caso 1: Coincide directamente con el precio esperado (Col 35 o neto esperado)
            if precio_esp is not None and abs(p_sap - precio_esp) <= 0.05:
                precio_coincide = True
                detalle_precio = f"PRECIO COINCIDE (${p_sap:,.2f})"

            # Caso 2: Coincide con la suma de MP + Flete (neto)
            elif suma_mp_fte is not None and abs(p_sap - suma_mp_fte) <= 0.05:
                precio_coincide = True
                detalle_precio = f"PRECIO NETO OK (MP+Fte: ${p_sap:,.2f})"

            # Caso 3: Coincide considerando factor PV (densidad / conversión TN a M3)
            elif pv and pv > 0:
                if precio_esp is not None and abs(p_sap - (precio_esp * pv)) <= 0.05:
                    precio_coincide = True
                    detalle_precio = f"PRECIO OK CON PV ({precio_esp:,.2f}*{pv} = ${p_sap:,.2f})"
                elif suma_mp_fte is not None and abs(p_sap - (suma_mp_fte * pv)) <= 0.05:
                    precio_coincide = True
                    detalle_precio = f"PRECIO NETO OK CON PV ({suma_mp_fte:,.2f}*{pv} = ${p_sap:,.2f})"
                elif precio_esp is not None and abs(p_sap - (precio_esp / pv)) <= 0.05:
                    precio_coincide = True
                    detalle_precio = f"PRECIO OK CON PV ({precio_esp:,.2f}/{pv} = ${p_sap:,.2f})"
                elif suma_mp_fte is not None and abs(p_sap - (suma_mp_fte / pv)) <= 0.05:
                    precio_coincide = True
                    detalle_precio = f"PRECIO NETO OK CON PV ({suma_mp_fte:,.2f}/{pv} = ${p_sap:,.2f})"

            # Caso 4: Si es modalidad Recogido (02) o flete es 0, coincide con solo MP
            elif (not p_fte or mod_esp == "02") and p_mp is not None and abs(p_sap - p_mp) <= 0.05:
                precio_coincide = True
                detalle_precio = f"PRECIO OK (MP Recogido: ${p_sap:,.2f})"

            if precio_coincide:
                val_desc.append(detalle_precio)
            else:
                p_ref = precio_esp if precio_esp is not None else (suma_mp_fte if suma_mp_fte is not None else 0.0)
                val_desc.append(f"DIFIERE PRECIO (SAP:${p_sap:,.2f} vs XL:${p_ref:,.2f})")
        elif precio_esp is not None:
            val_desc.append(f"PRECIO XL (${precio_esp:,.2f})")

        # Agregar avisos de popups capturados de SAP (V1 040, VP 191, etc.) a los detalles
        if res_va43 and res_va43.get("avisos"):
            val_desc.append("Avisos SAP: " + " | ".join(res_va43["avisos"]))

        # Estado global
        if any("ERROR" in d or "DIFIERE" in d for d in val_desc):
            estado = "DISCREPANCIA"
            msg = f"DISCREPANCIA: " + " | ".join(val_desc)
        else:
            estado = "OK"
            pos_t = f" Pos {pos_sap}" if pos_sap else ""
            msg = f"Contrato Venta SAP {c_sap}{pos_t} VALIDADO OK ({' | '.join(val_desc)})"

        return {
            "estado": estado,
            "mensaje": msg,
            "datos_sap": {
                "contrato": c_sap,
                "posicion": pos_sap,
                "precio": p_sap,
                "moneda": mon_sap,
                "um": um_sap,
                "fch_ini": fch_ini_esp,
                "fch_fin": fch_fin_esp,
                "exp_cond": mod_esp,
                "destino": d_sap or destino_esp,
                "shipfrom": sf_esp,
                "material": mat_esp if (m_sap and " " in m_sap) else (m_sap or mat_esp),
                "centro": cedis_esp,
                "avisos": avisos_sap
            }
        }

    def _manejar_popups_con_texto(self, session, delay=0.15):
        """Cierra ventanas emergentes secundarias capturando el texto de aviso (V1 040, VP 191, etc.)."""
        return handle_popups(session, delay=delay)

    def _consultar_va43_detallado(self, contrato, posicion=None, material=None):
        """
        Abre VA43 y extrae información exacta de la partida (posicion, precio_neto, um, avisos)
        utilizando la navegación documentada en joyonva43.vbs:
        1. Ingresa a /nVA43 con número de contrato.
        2. Captura y confirma avisos emergentes (wnd[1] tales como V1 040 o VP 191).
        3. Presiona el botón nativo de posicionamiento btnBT_POPO para saltar a la partida en wnd[1]/usr/txtRV45A-POSNR.
        4. Extrae de la fila 0 del TableControl: Posición, Material, Precio Neto (antepenúltima col) y UM (última col).
        """
        c_clean = str(contrato or "").strip()
        if not c_clean or any(k in c_clean.upper() for k in ["CREAR", "ALTA", "POR CREAR", "NONE", "NAN"]):
            return {"encontrado": False, "um": "", "precio_neto": None, "posicion": "", "avisos": []}

        pos_clean = str(int(float(posicion))) if posicion and str(posicion).replace(".", "").isdigit() else str(posicion or "").strip()
        mat_clean = str(material or "").strip()

        if not hasattr(self, "_cache_va43_detallado"):
            self._cache_va43_detallado = {}

        cache_k = (c_clean, pos_clean, mat_clean)
        if cache_k in self._cache_va43_detallado:
            return self._cache_va43_detallado[cache_k]

        session = getattr(self, "session", None)
        if not session:
            return {"encontrado": False, "um": "", "precio_neto": None, "posicion": "", "avisos": []}

        res = {
            "encontrado": False,
            "contrato": c_clean,
            "posicion": pos_clean,
            "material": mat_clean,
            "precio_neto": None,
            "um": "",
            "avisos": []
        }

        avisos_totales = []

        try:
            self.log(f"Consultando VA43 para Contrato Venta {c_clean} Pos {pos_clean or 'N/A'} (joyonva43)...")
            session.findById("wnd[0]/tbar[0]/okcd").text = "/nVA43"
            session.findById("wnd[0]").sendVKey(0)
            self._sleep(0.3)
            avisos_totales.extend(self._manejar_popups_con_texto(session))

            # Escribir número de contrato en pantalla inicial
            campo_cto = None
            try:
                campo_cto = session.findById("wnd[0]/usr/ctxtVBAK-VBELN")
            except Exception:
                for c in find_controls_recursive(session.findById("wnd[0]"), "GuiCTextField"):
                    if "VBELN" in getattr(c, "Name", "").upper() or "VBELN" in getattr(c, "Id", "").upper():
                        campo_cto = c
                        break

            if campo_cto:
                campo_cto.text = c_clean
            else:
                try: session.findById("wnd[0]/usr/txtVBAK-VBELN").text = c_clean
                except Exception: pass

            session.findById("wnd[0]").sendVKey(0)
            self._sleep(0.5)

            # Capturar y confirmar popups de advertencia iniciales (ej. V1 040, VP 191)
            avisos_iniciales = self._manejar_popups_con_texto(session)
            avisos_totales.extend(avisos_iniciales)

            sbar = None
            try: sbar = session.findById("wnd[0]/sbar")
            except Exception: pass
            if sbar and sbar.messageType in ("E", "A"):
                self.log(f"Aviso VA43: {sbar.text.strip()}")
                self._cache_va43_detallado[cache_k] = res
                return res

            res["encontrado"] = True

            # -------------------------------------------------------------
            # Paso 2: Extraer Destinatario de Mercancía (KUWEV-KUNNR) - Destino real de la ruta
            # Conforme a especificación técnica SAP:
            # Programa: SAPMV45A | Nº imagen: 4001 / 4701 | Status: U CONT | Tabla: KUWEV | Campo: KUNNR (KUNWE)
            # -------------------------------------------------------------
            destino_encontrado = ""
            solicitante_encontrado = ""

            # 1. Búsqueda directa en cabecera / dynpro activo para KUWEV-KUNNR (Destino / KUNWE)
            for fld_id in [
                "wnd[0]/usr/subSUBSCREEN_HEADER:SAPMV45A:4021/ctxtKUWEV-KUNNR",
                "wnd[0]/usr/subSUBSCREEN_HEADER:SAPMV45A:4020/ctxtKUWEV-KUNNR",
                "wnd[0]/usr/subSUBSCREEN_HEADER:SAPMV45A:4701/ctxtKUWEV-KUNNR",
                "wnd[0]/usr/subSUBSCREEN_HEADER:SAPMV45A:4001/ctxtKUWEV-KUNNR",
                "wnd[0]/usr/ctxtKUWEV-KUNNR",
                "wnd[0]/usr/txtKUWEV-KUNNR"
            ]:
                try:
                    w_fld = session.findById(fld_id)
                    if w_fld and getattr(w_fld, "Text", "").strip():
                        destino_encontrado = str(w_fld.Text).strip().lstrip("0")
                        break
                except Exception:
                    pass

            # 2. Búsqueda recursiva de KUWEV-KUNNR en controles de la ventana actual
            if not destino_encontrado:
                for c_txt in find_controls_recursive(session.findById("wnd[0]"), "GuiCTextField") + find_controls_recursive(session.findById("wnd[0]"), "GuiTextField"):
                    c_name = getattr(c_txt, "Name", "").upper()
                    c_id = getattr(c_txt, "Id", "").upper()
                    if "KUWEV-KUNNR" in c_name or "KUWEV-KUNNR" in c_id:
                        val = str(getattr(c_txt, "Text", "")).strip().lstrip("0")
                        if val:
                            destino_encontrado = val
                            break

            # 3. Extraer Solicitante (KUAGV-KUNNR) de cabecera como dato complementario
            for fld_id in [
                "wnd[0]/usr/subSUBSCREEN_HEADER:SAPMV45A:4021/ctxtKUAGV-KUNNR",
                "wnd[0]/usr/subSUBSCREEN_HEADER:SAPMV45A:4020/ctxtKUAGV-KUNNR",
                "wnd[0]/usr/ctxtKUAGV-KUNNR", "wnd[0]/usr/txtKUAGV-KUNNR"
            ]:
                try:
                    k_fld = session.findById(fld_id)
                    if k_fld and getattr(k_fld, "Text", "").strip():
                        solicitante_encontrado = str(k_fld.Text).strip().lstrip("0")
                        break
                except Exception:
                    pass

            if not solicitante_encontrado:
                for c_txt in find_controls_recursive(session.findById("wnd[0]"), "GuiCTextField"):
                    c_name = getattr(c_txt, "Name", "").upper()
                    if "KUAGV-KUNNR" in c_name or "VBAK-KUNNR" in c_name:
                        val = str(getattr(c_txt, "Text", "")).strip().lstrip("0")
                        if val:
                            solicitante_encontrado = val
                            break

            res["solicitante"] = solicitante_encontrado
            res["dest_mercancia"] = destino_encontrado
            res["destino"] = destino_encontrado or solicitante_encontrado

            # -------------------------------------------------------------
            # Paso 3: Asegurar pestaña Posición (tabpT\01) activa y posicionar partida
            # -------------------------------------------------------------
            try:
                tab_pos = session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\\01")
                if tab_pos:
                    tab_pos.select()
                    self._sleep(0.2)
            except Exception:
                pass

            # Posicionar Partida usando el botón nativo btnBT_POPO (joyonva43.vbs) si se conoce pos_clean
            if pos_clean:
                btn_popo = None
                for b_path in [
                    "wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\\01/ssubSUBSCREEN_BODY:SAPMV45A:4406/subSUBSCREEN_TC:SAPMV45A:4906/subSUBSCREEN_BUTTONS:SAPMV45A:4052/btnBT_POPO",
                    "wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\\01/ssubSUBSCREEN_BODY:SAPMV45A:4400/subSUBSCREEN_TC:SAPMV45A:4900/subSUBSCREEN_BUTTONS:SAPMV45A:4052/btnBT_POPO",
                ]:
                    try:
                        btn_popo = session.findById(b_path)
                        if btn_popo:
                            break
                    except Exception:
                        pass

                if not btn_popo:
                    for btn in find_controls_recursive(session.findById("wnd[0]"), "GuiButton"):
                        b_name = getattr(btn, "Name", "").upper()
                        b_id = getattr(btn, "Id", "").upper()
                        if "BT_POPO" in b_name or "BT_POPO" in b_id or "POPO" in b_name:
                            btn_popo = btn
                            break

                if btn_popo:
                    try:
                        btn_popo.press()
                        for _ in range(15):
                            self._sleep(0.1)
                            if getattr(session, "Children", None) and session.Children.Count > 1:
                                break
                        if getattr(session, "Children", None) and session.Children.Count > 1:
                            wnd1 = session.findById("wnd[1]")
                            fld_pos = None
                            try:
                                fld_pos = wnd1.findById("usr/txtRV45A-POSNR")
                            except Exception:
                                for tf in find_controls_recursive(wnd1, "GuiTextField") + find_controls_recursive(wnd1, "GuiCTextField"):
                                    if "RV45A-POSNR" in getattr(tf, "Name", "").upper() or "POSNR" in getattr(tf, "Name", "").upper():
                                        fld_pos = tf
                                        break
                            if fld_pos:
                                fld_pos.text = str(pos_clean)
                            wnd1.sendVKey(0)
                            self._sleep(0.3)
                            avisos_totales.extend(self._manejar_popups_con_texto(session))
                    except Exception as e_popo:
                        self.log(f"Aviso accionando btnBT_POPO: {e_popo}")

            # -------------------------------------------------------------
            # Paso 4: Extraer información del TableControl e inspeccionar filas
            # -------------------------------------------------------------
            tbls = find_controls_recursive(session.findById("wnd[0]"), "TableControl")
            if tbls:
                tbl = tbls[0]
                col_cnt = getattr(tbl, "Columns", []).Count if hasattr(tbl, "Columns") else 0
                if col_cnt > 0:
                    c_pos_idx = 0
                    c_mat_idx = 1
                    col_netpr_idx = col_cnt - 3 if col_cnt >= 3 else (col_cnt - 2 if col_cnt >= 2 else 0)
                    col_um_idx = col_cnt - 1
                    pos_col_locked = False
                    mat_col_locked = False
                    netpr_col_locked = False
                    um_col_locked = False

                    for c_i in range(col_cnt):
                        try:
                            c_obj = tbl.Columns(c_i)
                            c_id = getattr(c_obj, "Id", "").upper()
                            c_name = getattr(c_obj, "Name", "").upper()
                            c_title = getattr(c_obj, "Title", "").upper()
                            c_tip = getattr(c_obj, "Tooltip", "").upper()

                            cell_tag = ""
                            try:
                                cell_tag = getattr(tbl.GetCell(0, c_i), "Name", "").upper() + " " + getattr(tbl.GetCell(0, c_i), "Id", "").upper()
                            except Exception:
                                pass

                            combo = f"{c_id} {c_name} {c_title} {c_tip} {cell_tag}"

                            if not pos_col_locked and ("VBAP-POSNR" in combo or "RV45A-POSNR" in combo or "POSNR" in combo or c_title in ("POS", "POS.", "POSICIÓN", "POSICION")):
                                c_pos_idx = c_i
                                if "POSNR" in combo:
                                    pos_col_locked = True
                            elif not mat_col_locked and ("VBAP-MATNR" in combo or "RV45A-MABNR" in combo or "MATNR" in combo):
                                mat_col_locked = True
                                c_mat_idx = c_i
                            elif not mat_col_locked and c_title in ("MATERIAL", "MAT.", "MAT") and not any(k in c_title for k in ["DENOM", "DESC", "TEXT", "DEN.", "GR.", "GRUPO", "CLI", "EST"]):
                                c_mat_idx = c_i
                            elif not netpr_col_locked and ("VBAP-NETPR" in combo or "NETPR" in combo or "PRC.NETO" in combo or "PRECIO NETO" in combo):
                                col_netpr_idx = c_i
                                if "NETPR" in combo:
                                    netpr_col_locked = True
                            elif not um_col_locked and (any(k in combo for k in ["VBAP-KMEIN", "VBAP-ZIEME", "VBAP-VRKME", "KMEIN", "ZIEME", "VRKME"]) or c_title in ("UM", "U.M.")):
                                col_um_idx = c_i
                                if any(k in combo for k in ["KMEIN", "ZIEME", "VRKME"]):
                                    um_col_locked = True
                        except Exception:
                            pass

                    def _safe_cell(r, c):
                        if c is None or c < 0:
                            return ""
                        try:
                            cell = tbl.GetCell(r, c)
                            return str(getattr(cell, "Text", "") or "").strip()
                        except Exception:
                            return ""

                    vis_rows = getattr(tbl, "VisibleRowCount", 15) or 15
                    row_limit = min(max(vis_rows, getattr(tbl, "RowCount", 0) or 0), 50)
                    row_limit = max(row_limit, 5)

                    filas_candidatas = []
                    for r_i in range(row_limit):
                        p_val = _safe_cell(r_i, c_pos_idx) or _safe_cell(r_i, 0)
                        m_val = _safe_cell(r_i, c_mat_idx)
                        netpr_val = clean_num(_safe_cell(r_i, col_netpr_idx))
                        um_val = clean_um(_safe_cell(r_i, col_um_idx))

                        if not m_val or " " in m_val or (mat_clean and not mat_matches(m_val, mat_clean)):
                            for test_c in range(min(col_cnt, 15)):
                                cand_m = _safe_cell(r_i, test_c)
                                if cand_m and " " not in cand_m and mat_clean and mat_matches(cand_m, mat_clean):
                                    m_val = cand_m
                                    break

                        if p_val or m_val or netpr_val is not None:
                            filas_candidatas.append({
                                "row_idx": r_i,
                                "posicion": p_val,
                                "material": m_val,
                                "precio_neto": netpr_val,
                                "um": um_val
                            })

                    # Selección inteligente de la mejor fila candidata
                    matched_cand = None

                    # 1. Coincidencia simultánea de Posición y Material
                    if pos_clean and mat_clean:
                        for cand in filas_candidatas:
                            if pos_matches(cand["posicion"], pos_clean) and mat_matches(cand["material"], mat_clean):
                                matched_cand = cand
                                break

                    # 2. Coincidencia por Material esperado (descubre la posición correcta ej. 20)
                    if not matched_cand and mat_clean:
                        for cand in filas_candidatas:
                            if mat_matches(cand["material"], mat_clean):
                                matched_cand = cand
                                break

                    # 3. Coincidencia por Posición esperada
                    if not matched_cand and pos_clean:
                        for cand in filas_candidatas:
                            if pos_matches(cand["posicion"], pos_clean):
                                matched_cand = cand
                                break

                    # 4. Fallback a fila 0 si no hubo otra coincidencia
                    if not matched_cand and filas_candidatas:
                        matched_cand = filas_candidatas[0]

                    if matched_cand:
                        if matched_cand.get("posicion"):
                            res["posicion"] = matched_cand["posicion"]
                        if matched_cand.get("material") and " " not in matched_cand["material"]:
                            res["material"] = matched_cand["material"]
                        if matched_cand.get("precio_neto") is not None:
                            res["precio_neto"] = matched_cand["precio_neto"]
                        if matched_cand.get("um") and is_valid_um(matched_cand["um"]):
                            res["um"] = matched_cand["um"]

            # Estrategia de respaldo: escaneo de GuiTextField y GuiCTextField
            if not res["um"] or res["precio_neto"] is None or not res["posicion"] or (mat_clean and not mat_matches(res.get("material"), mat_clean)):
                all_txts = find_controls_recursive(session.findById("wnd[0]"), "GuiTextField") + \
                           find_controls_recursive(session.findById("wnd[0]"), "GuiCTextField")

                row_map = {}
                for ctrl in all_txts:
                    cid = getattr(ctrl, "Id", "").upper()
                    cname = getattr(ctrl, "Name", "").upper()
                    txt_val = str(getattr(ctrl, "Text", "")).strip()

                    m = re.search(r"\[\d+,\s*(\d+)\]", cid)
                    if m:
                        r_idx = int(m.group(1))
                        if r_idx not in row_map:
                            row_map[r_idx] = {}

                        if "POSNR" in cid or "POSNR" in cname:
                            row_map[r_idx]["pos"] = txt_val
                        elif "MATNR" in cid or "MATNR" in cname:
                            row_map[r_idx]["mat"] = txt_val
                        elif "NETPR" in cid or "NETPR" in cname:
                            row_map[r_idx]["netpr"] = txt_val
                        elif "ZIEME" in cid or "ZIEME" in cname or "KMEIN" in cid or "KMEIN" in cname or "VRKME" in cid:
                            row_map[r_idx]["um"] = txt_val

                target_r_data = None
                # 1. Por posición y material
                if pos_clean and mat_clean:
                    for r_idx, r_dict in row_map.items():
                        if pos_matches(r_dict.get("pos", ""), pos_clean) and mat_matches(r_dict.get("mat", ""), mat_clean):
                            target_r_data = r_dict
                            break

                # 2. Por material
                if not target_r_data and mat_clean:
                    for r_idx, r_dict in row_map.items():
                        if mat_matches(r_dict.get("mat", ""), mat_clean):
                            target_r_data = r_dict
                            break

                # 3. Por posición
                if not target_r_data and pos_clean:
                    for r_idx, r_dict in row_map.items():
                        if pos_matches(r_dict.get("pos", ""), pos_clean):
                            target_r_data = r_dict
                            break

                # 4. Fila 0
                if not target_r_data and 0 in row_map:
                    target_r_data = row_map[0]

                if target_r_data:
                    if target_r_data.get("pos"):
                        res["posicion"] = target_r_data["pos"]
                    if target_r_data.get("mat") and " " not in target_r_data["mat"]:
                        res["material"] = target_r_data["mat"]
                    if target_r_data.get("netpr"):
                        res["precio_neto"] = clean_num(target_r_data["netpr"])
                    if target_r_data.get("um"):
                        c_u = clean_um(target_r_data["um"])
                        if c_u and is_valid_um(c_u):
                            res["um"] = c_u

            # -------------------------------------------------------------
            # Paso 5: Si no se detectó KUWEV-KUNNR en cabecera, consultar pestaña 'Contrato' (Dynpro 4701, Status U CONT)
            # -------------------------------------------------------------
            if not res.get("dest_mercancia"):
                tab_contrato = None
                for t_id in [
                    "wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\\02",
                    "wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\\03",
                    "wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\\04",
                    "wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\\05"
                ]:
                    try:
                        t_obj = session.findById(t_id)
                        t_txt = getattr(t_obj, "Text", "").upper()
                        t_tooltip = getattr(t_obj, "Tooltip", "").upper()
                        if "CONT" in t_txt or "CONT" in t_tooltip:
                            tab_contrato = t_obj
                            break
                        elif not tab_contrato and t_id.endswith("02"):
                            tab_contrato = t_obj
                    except Exception:
                        pass

                if tab_contrato:
                    try:
                        tab_contrato.select()
                        self._sleep(0.3)
                        self._manejar_popups_con_texto(session)

                        # Buscar en Dynpro 4701 el campo KUWEV-KUNNR
                        dest_4701 = ""
                        for fld_id in [
                            "wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\\02/ssubSUBSCREEN_BODY:SAPMV45A:4701/ctxtKUWEV-KUNNR",
                            "wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\\02/ssubSUBSCREEN_BODY:SAPMV45A:4701/txtKUWEV-KUNNR",
                            "wnd[0]/usr/ctxtKUWEV-KUNNR",
                            "wnd[0]/usr/txtKUWEV-KUNNR"
                        ]:
                            try:
                                w_fld = session.findById(fld_id)
                                if w_fld and getattr(w_fld, "Text", "").strip():
                                    dest_4701 = str(w_fld.Text).strip().lstrip("0")
                                    break
                            except Exception:
                                pass

                        if not dest_4701:
                            for c_txt in find_controls_recursive(session.findById("wnd[0]"), "GuiCTextField") + find_controls_recursive(session.findById("wnd[0]"), "GuiTextField"):
                                c_name = getattr(c_txt, "Name", "").upper()
                                c_id = getattr(c_txt, "Id", "").upper()
                                if "KUWEV-KUNNR" in c_name or "KUWEV-KUNNR" in c_id:
                                    val = str(getattr(c_txt, "Text", "")).strip().lstrip("0")
                                    if val:
                                        dest_4701 = val
                                        break

                        if dest_4701:
                            res["dest_mercancia"] = dest_4701
                            res["destino"] = dest_4701
                    except Exception as e_tab:
                        self.log(f"Aviso consultando pestaña Contrato 4701 para KUWEV-KUNNR: {e_tab}")

            if not res.get("destino"):
                for av in avisos_totales:
                    m_cli = re.search(r"(?:interlocutor\s+comercial|cliente|destinatario)\s+(\d+)", str(av), re.IGNORECASE)
                    if m_cli:
                        res["destino"] = m_cli.group(1).lstrip("0")
                        break

        except Exception as e_va:
            self.log(f"Aviso en _consultar_va43_detallado: {e_va}")
        finally:
            try:
                session.findById("wnd[0]/tbar[0]/okcd").text = "/n"
                session.findById("wnd[0]").sendVKey(0)
                self._sleep(0.2)
                self._manejar_popups_con_texto(session)
            except Exception:
                pass

        res["avisos"] = [a for a in avisos_totales if a]
        if res.get("um"):
            self.log(f"VA43 Contrato {c_clean} Pos {res.get('posicion')}: UM='{res['um']}' Precio={res.get('precio_neto')}")
        self._cache_va43_detallado[cache_k] = res
        return res

    def _extraer_um_va43(self, contrato, posicion=None, material=None):
        """Extrae la UM de VA43 delegando a _consultar_va43_detallado."""
        res = self._consultar_va43_detallado(contrato, posicion, material)
        return res.get("um", "")

    def _consultar_va43(self, contrato, destino_esp, mat_esp):
        """Abre VA43 y visualiza el contrato de venta."""
        um = self._extraer_um_va43(contrato=contrato, material=mat_esp)
        if um:
            return {
                "estado": "OK",
                "mensaje": f"Contrato de Venta {contrato} visualizado OK en VA43 (UM: {um})",
                "datos_sap": {"contrato": contrato, "um": um}
            }
        return {
            "estado": "NO_ENCONTRADO",
            "mensaje": f"VA43: No se pudo visualizar contrato {contrato}",
            "datos_sap": {}
        }

    def _buscar_contrato_venta_sap(self, vkorg, destino_kunnr, material):
        """Intenta localizar contratos de venta en VA45."""
        res = {"encontrado": False, "contrato": ""}
        try:
            session = self.session
            session.findById("wnd[0]/tbar[0]/okcd").text = "/nVA45"
            session.findById("wnd[0]").sendVKey(0)
            self._sleep(0.3)
            handle_popups(session)

            sbar = session.findById("wnd[0]/sbar")
            if sbar and sbar.messageType in ("E", "A") and "no existe" in sbar.text.lower():
                return res

            for cid, val in [
                ("wnd[0]/usr/ctxtVBCO6-KUNNR", destino_kunnr),
                ("wnd[0]/usr/ctxtVBCO6-VKORG", vkorg),
                ("wnd[0]/usr/ctxtVBCO6-MATNR", material),
            ]:
                try:
                    session.findById(cid).text = val
                except Exception:
                    pass

            try:
                session.findById("wnd[0]/tbar[1]/btn[8]").press()
            except Exception:
                session.findById("wnd[0]").sendVKey(8)
            self._sleep(0.6)
            handle_popups(session)

            grids = find_controls_recursive(session.findById("wnd[0]"), "GuiGridView")
            if grids:
                grid = grids[0]
                if getattr(grid, "RowCount", 0) > 0:
                    for r in range(grid.RowCount):
                        c_num = grid.GetCellValue(r, "VBELN")
                        if c_num:
                            res["encontrado"] = True
                            res["contrato"] = str(c_num).strip()
                            break
        except Exception as e:
            self.log(f"Aviso buscando contrato venta: {e}")
        return res

    def validar_material_mm03(
        self,
        material,
        centro,
        org_ventas,
        canal="00",
        almacen="1002",
        peso_esp=None,
        clase_esp=None
    ):
        """
        Ejecuta la validación de un material en MM03 conforme a las especificaciones corporativas CEMEX:
        1. Acceder a /nMM03 e ingresar Material.
        2. Seleccionar vistas Datos base 1 y Datos base 2.
        3. En Datos adicionales -> Unidades de medida: extraer y verificar peso volumétrico (UMREZ).
        4. En Clasificación: extraer/verificar clase de material y clasificador (RCTMS-MWERT).
        5. Ingresar Niveles de organización: Centro, Org. Ventas, Canal '00'.
        6. En Ventas Gral/Centro (tabpSP06):
           - Check Stock Negativo activo (MARC-XMCNG).
           - Verif. Disponibilidad == 'KP' (MARC-MTVFP).
        7. En Vista Compras (tabpSP09):
           - Check Ind. ped. autom. activo (MARC-KAUTB).
        8. En Almacén / Contabilidad (almacén '1002', tabpSP24):
           - Control de precios == 'S' o 'V' (MBEW-VPRSV).
        """
        mat_str = str(material).strip()
        centro_str = str(centro).strip()
        org_str = str(org_ventas).strip()
        canal_str = str(canal or "00").strip()
        almacen_str = str(almacen or "1002").strip()

        if not getattr(self, "session", None):
            self.conectar()
        session = self.session

        try:
            session.findById("wnd[0]").maximize()
        except Exception:
            pass

        self.log(f"[MM03] Iniciando validación para Material: {mat_str}, Centro: {centro_str}, Org. Vta: {org_str}...")

        try:
            # 1. Acceder a /nMM03 e ingresar Material
            session.findById("wnd[0]/tbar[0]/okcd").text = "/nMM03"
            session.findById("wnd[0]").sendVKey(0)
            self._sleep(0.3)
            handle_popups(session, delay=0.1)

            session.findById("wnd[0]/usr/ctxtRMMG1-MATNR").text = mat_str
            session.findById("wnd[0]").sendVKey(0)
            self._sleep(0.3)

            # Verificar si SAP arrojó error de material inexistente
            try:
                sbar = session.findById("wnd[0]/sbar")
                if getattr(sbar, "MessageType", "") == "E":
                    err_txt = getattr(sbar, "Text", "")
                    self.log(f"[MM03] Error SAP al buscar material: {err_txt}")
                    return {
                        "material": mat_str,
                        "centro": centro_str,
                        "org_ventas": org_str,
                        "canal": canal_str,
                        "almacen": almacen_str,
                        "valido": False,
                        "estado": "ERROR",
                        "msg": f"El material {mat_str} no existe en SAP: {err_txt}",
                        "detalles": [f"❌ Error SAP: {err_txt}"],
                        "datos_extraidos": {},
                        "checks": {
                            "stock_negativo_ok": False,
                            "verif_disp_ok": False,
                            "pedido_auto_ok": False,
                            "control_precios_ok": False,
                            "peso_volumetrico_ok": False,
                            "clasificacion_ok": False,
                        },
                        "avisos": []
                    }
            except Exception:
                pass

            avisos = []

            # 2. Selección de vistas principales (Datos base 1 y 2)
            if getattr(session, "Children", None) and session.Children.Count > 1:
                popup = session.findById("wnd[1]")
                try:
                    tbl_view = popup.findById("usr/tblSAPLMGMMTC_VIEW")
                    found_b1 = False
                    found_b2 = False
                    row_cnt = getattr(tbl_view, "RowCount", 10)
                    for r in range(min(row_cnt, 15)):
                        txt_cell = ""
                        try:
                            txt_cell = popup.findById(f"usr/tblSAPLMGMMTC_VIEW/txtMSICHTAUSW-DYTXT[0,{r}]").Text.strip().lower()
                        except Exception:
                            try:
                                txt_cell = tbl_view.getCell(r, 0).Text.strip().lower()
                            except Exception:
                                pass
                        if "base 1" in txt_cell or "datos base 1" in txt_cell:
                            try:
                                tbl_view.getAbsoluteRow(r).selected = True
                                found_b1 = True
                            except Exception:
                                pass
                        elif "base 2" in txt_cell or "datos base 2" in txt_cell:
                            try:
                                tbl_view.getAbsoluteRow(r).selected = True
                                found_b2 = True
                            except Exception:
                                pass

                    if not found_b1:
                        try:
                            tbl_view.getAbsoluteRow(0).selected = True
                        except Exception:
                            pass
                    if not found_b2:
                        try:
                            tbl_view.getAbsoluteRow(1).selected = True
                        except Exception:
                            pass

                    try:
                        popup.findById("tbar[0]/btn[0]").press()
                    except Exception:
                        popup.sendVKey(0)
                    self._sleep(0.3)
                except Exception:
                    try:
                        popup.sendVKey(0)
                        self._sleep(0.2)
                    except Exception:
                        pass

            # 3. Peso Volumétrico (Datos adicionales -> Unidades de medida)
            peso_volumetrico = None
            try:
                self.log("[MM03] Consultando 'Datos adicionales' para peso volumétrico...")
                session.findById("wnd[0]/tbar[1]/btn[30]").press()
                self._sleep(0.3)
                session.findById("wnd[0]/usr/tabsTABSPR1/tabpZU02").select()
                self._sleep(0.3)

                # Intentar leer tabla de unidades de medida
                tbl_me = None
                try:
                    tbl_me = session.findById("wnd[0]/usr/tabsTABSPR1/tabpZU02/ssubTABFRA1:SAPLMGMM:2110/subSUB2:SAPLMGD1:8020/tblSAPLMGD1TC_ME_8020")
                except Exception:
                    tbl_me = find_control_by_name_recursive(session.findById("wnd[0]"), "tblSAPLMGD1TC_ME_8020")

                # A. Intentar control exacto txtSMEINH-UMREZ[4,1]
                try:
                    val_direct = session.findById("wnd[0]/usr/tabsTABSPR1/tabpZU02/ssubTABFRA1:SAPLMGMM:2110/subSUB2:SAPLMGD1:8020/tblSAPLMGD1TC_ME_8020/txtSMEINH-UMREZ[4,1]").Text.strip()
                    if val_direct:
                        peso_volumetrico = val_direct
                except Exception:
                    pass

                # B. Iterar filas de la tabla si no se obtuvo
                if not peso_volumetrico and tbl_me:
                    r_cnt = getattr(tbl_me, "RowCount", 5)
                    for r in range(min(r_cnt, 10)):
                        try:
                            val = ""
                            try:
                                val = session.findById(f"{tbl_me.Id}/txtSMEINH-UMREZ[4,{r}]").Text.strip()
                            except Exception:
                                val = tbl_me.getCell(r, 4).Text.strip()
                            if val and clean_num(val) not in (None, 0):
                                peso_volumetrico = val
                                break
                        except Exception:
                            continue

                # C. Búsqueda recursiva de cualquier control con SMEINH-UMREZ
                if not peso_volumetrico:
                    ctrl_u = find_control_by_name_recursive(session.findById("wnd[0]"), "SMEINH-UMREZ")
                    if ctrl_u:
                        peso_volumetrico = getattr(ctrl_u, "Text", "").strip()

                self.log(f"[MM03] Peso volumétrico obtenido: {peso_volumetrico or 'No definido'}")

                # Regresar a datos principales (btn[27])
                try:
                    session.findById("wnd[0]/tbar[1]/btn[27]").press()
                    self._sleep(0.3)
                except Exception:
                    pass
            except Exception as e:
                self.log(f"[MM03] Aviso en Datos adicionales: {e}")
                try:
                    session.findById("wnd[0]/tbar[1]/btn[27]").press()
                    self._sleep(0.2)
                except Exception:
                    pass

            # 4. Clasificación (tabpSP03)
            clase_material = None
            clasificador_valor = None
            try:
                self.log("[MM03] Consultando vista 'Clasificación' (tabpSP03)...")
                session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP03").select()
                self._sleep(0.3)

                # Manejar posibles popups en clasificación
                if getattr(session, "Children", None) and session.Children.Count > 1:
                    try:
                        session.findById("wnd[1]/tbar[0]/btn[0]").press()
                    except Exception:
                        session.findById("wnd[1]").sendVKey(0)
                    self._sleep(0.2)

                # Extraer clase de material
                try:
                    tbl_obj = session.findById("wnd[0]/usr/subSUBSCR_ZUORD:SAPLCLFM:1600/tblSAPLCLFMTC_OBJ_CLASS")
                    for c in range(5):
                        try:
                            v = tbl_obj.getCell(0, c).Text.strip()
                            if v:
                                clase_material = v
                                break
                        except Exception:
                            pass
                except Exception:
                    pass

                # Extraer valor clasificador / características (ctxtRCTMS-MWERT[1,4])
                try:
                    c_val = session.findById("wnd[0]/usr/subSUBSCR_BEWERT:SAPLCTMS:5000/tabsTABSTRIP_CHAR/tabpTAB1/ssubTABSTRIP_CHAR_GR:SAPLCTMS:5100/tblSAPLCTMSCHARS_S/ctxtRCTMS-MWERT[1,4]").Text.strip()
                    if c_val:
                        clasificador_valor = c_val
                except Exception:
                    pass

                if not clasificador_valor:
                    tbl_c = None
                    try:
                        tbl_c = session.findById("wnd[0]/usr/subSUBSCR_BEWERT:SAPLCTMS:5000/tabsTABSTRIP_CHAR/tabpTAB1/ssubTABSTRIP_CHAR_GR:SAPLCTMS:5100/tblSAPLCTMSCHARS_S")
                    except Exception:
                        tbl_c = find_control_by_name_recursive(session.findById("wnd[0]"), "tblSAPLCTMSCHARS_S")
                    if tbl_c:
                        r_cnt = getattr(tbl_c, "RowCount", 10)
                        for r in range(min(r_cnt, 15)):
                            try:
                                val = session.findById(f"{tbl_c.Id}/ctxtRCTMS-MWERT[1,{r}]").Text.strip()
                                if val:
                                    clasificador_valor = val
                                    break
                            except Exception:
                                continue

                if not clasificador_valor:
                    ctrl_m = find_control_by_name_recursive(session.findById("wnd[0]"), "RCTMS-MWERT")
                    if ctrl_m:
                        clasificador_valor = getattr(ctrl_m, "Text", "").strip()

                self.log(f"[MM03] Clasificación: Clase={clase_material or 'N/A'}, Clasificador={clasificador_valor or 'N/A'}")
            except Exception as e:
                self.log(f"[MM03] Aviso en Clasificación: {e}")

            # 5. Niveles de Organización (Centro, Org. Ventas, Canal '00')
            self.log(f"[MM03] Configurando niveles organizativos: Centro {centro_str}, Org. Vta {org_str}, Canal {canal_str}...")
            try:
                try:
                    session.findById("wnd[0]/tbar[1]/btn[32]").press()
                    self._sleep(0.2)
                except Exception:
                    pass
                session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP06").select()
                self._sleep(0.3)
            except Exception:
                pass

            err_org = None
            if getattr(session, "Children", None) and session.Children.Count > 1:
                popup = session.findById("wnd[1]")
                try:
                    if centro_str:
                        for f_id in ["usr/ctxtRMMG1-WERKS", "usr/ctxtWERKS", "usr/txtRMMG1-WERKS"]:
                            try:
                                popup.findById(f_id).text = centro_str
                                break
                            except Exception:
                                pass
                    if org_str:
                        for f_id in ["usr/ctxtRMMG1-VKORG", "usr/ctxtVKORG", "usr/txtRMMG1-VKORG"]:
                            try:
                                popup.findById(f_id).text = org_str
                                break
                            except Exception:
                                pass
                    if canal_str:
                        for f_id in ["usr/ctxtRMMG1-VTWEG", "usr/ctxtVTWEG", "usr/txtRMMG1-VTWEG"]:
                            try:
                                popup.findById(f_id).text = canal_str
                                break
                            except Exception:
                                pass
                    self._sleep(0.2)
                    try:
                        popup.findById("tbar[0]/btn[0]").press()
                    except Exception:
                        popup.sendVKey(0)
                    self._sleep(0.3)
                except Exception as e:
                    self.log(f"[MM03] Aviso rellenando niveles org: {e}")
                    try:
                        popup.sendVKey(0)
                        self._sleep(0.2)
                    except Exception:
                        pass

            # Revisar barra de estado por errores de niveles organizativos
            for w in ("wnd[1]", "wnd[0]"):
                try:
                    sb = session.findById(f"{w}/sbar")
                    if getattr(sb, "MessageType", "") == "E":
                        err_org = getattr(sb, "Text", "")
                        break
                except Exception:
                    pass

            if err_org:
                return {
                    "material": mat_str,
                    "centro": centro_str,
                    "org_ventas": org_str,
                    "canal": canal_str,
                    "almacen": almacen_str,
                    "valido": False,
                    "estado": "ERROR",
                    "msg": f"Parámetros organizativos incorrectos o no asignados al material: {err_org}",
                    "detalles": [f"❌ Error en niveles org (Centro {centro_str}, Org {org_str}, Canal {canal_str}): {err_org}"],
                    "datos_extraidos": {
                        "peso_volumetrico": peso_volumetrico,
                        "clase_material": clase_material,
                        "clasificador": clasificador_valor,
                    },
                    "checks": {
                        "stock_negativo_ok": False,
                        "verif_disp_ok": False,
                        "pedido_auto_ok": False,
                        "control_precios_ok": False,
                        "peso_volumetrico_ok": bool(peso_volumetrico),
                        "clasificacion_ok": bool(clasificador_valor or clase_material),
                    },
                    "avisos": []
                }

            avisos.extend(handle_popups(session, delay=0.15))

            # 6. Validar Ventas General/Centro (tabpSP06)
            stock_negativo_activo = False
            verif_disp_val = ""
            try:
                chk_stock = None
                try:
                    chk_stock = session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP06/ssubTABFRA1:SAPLMGMM:2000/subSUB5:SAPLZMM_GINCOFIELDS:5802/chkMARC-XMCNG")
                except Exception:
                    chk_stock = find_control_by_name_recursive(session.findById("wnd[0]"), "MARC-XMCNG")

                if chk_stock:
                    sel = getattr(chk_stock, "Selected", False)
                    stock_negativo_activo = bool(sel is True or sel == 1 or str(sel).lower() in ("true", "x", "1"))

                ctrl_disp = None
                try:
                    ctrl_disp = find_data_field_by_name(session.findById("wnd[0]"), "MARC-MTVFP")
                except Exception:
                    pass
                if not ctrl_disp:
                    try:
                        ctrl_disp = find_data_field_by_name(session.findById("wnd[0]"), "MTVFP")
                    except Exception:
                        pass

                if ctrl_disp:
                    verif_disp_val = getattr(ctrl_disp, "Text", "").strip()

                # Salvaguarda: si capturó la etiqueta o está vacío, buscar directamente en controles de texto
                if not verif_disp_val or "verif" in verif_disp_val.lower():
                    try:
                        c_fields = find_controls_recursive(session.findById("wnd[0]"), "GuiCTextField") + find_controls_recursive(session.findById("wnd[0]"), "GuiTextField")
                        for cf in c_fields:
                            cf_id = str(getattr(cf, "Id", ""))
                            cf_name = str(getattr(cf, "Name", ""))
                            if "mtvfp" in cf_id.lower() or "mtvfp" in cf_name.lower():
                                val_cf = getattr(cf, "Text", "").strip()
                                if val_cf and "verif" not in val_cf.lower():
                                    verif_disp_val = val_cf
                                    break
                    except Exception:
                        pass

                self.log(f"[MM03] Ventas Gral/Centro: Stock Negativo={'ACTIVO' if stock_negativo_activo else 'INACTIVO'}, Verif.Disp='{verif_disp_val}'")
            except Exception as e:
                self.log(f"[MM03] Aviso en Ventas Gral/Centro: {e}")

            # 7. Validar Vista Compras (tabpSP09) -> Ind. pedido automático
            pedido_auto_activo = False
            try:
                self.log("[MM03] Consultando 'Vista Compras' (tabpSP09)...")
                session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP09").select()
                self._sleep(0.3)
                avisos.extend(handle_popups(session, delay=0.1))

                chk_auto = None
                try:
                    chk_auto = session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP09/ssubTABFRA1:SAPLMGMM:2000/subSUB2:SAPLMGD1:2301/chkMARC-KAUTB")
                except Exception:
                    chk_auto = find_control_by_name_recursive(session.findById("wnd[0]"), "MARC-KAUTB")

                if chk_auto:
                    sel_a = getattr(chk_auto, "Selected", False)
                    pedido_auto_activo = bool(sel_a is True or sel_a == 1 or str(sel_a).lower() in ("true", "x", "1"))

                self.log(f"[MM03] Vista Compras: Ind. Pedido Automático={'ACTIVO' if pedido_auto_activo else 'INACTIVO'}")
            except Exception as e:
                self.log(f"[MM03] Aviso en Vista Compras: {e}")

            # 8. Almacén 1002 y Vista Contabilidad 1 (tabpSP24) -> Control de precios
            control_precios_val = ""
            try:
                self.log(f"[MM03] Consultando Contabilidad (tabpSP24) para almacén '{almacen_str}'...")
                try:
                    session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP13").select()
                    self._sleep(0.2)
                    if getattr(session, "Children", None) and session.Children.Count > 1:
                        pop = session.findById("wnd[1]")
                        try:
                            pop.findById("usr/ctxtRMMG1-LGORT").text = almacen_str
                            pop.findById("tbar[0]/btn[0]").press()
                            self._sleep(0.2)
                        except Exception:
                            pop.sendVKey(0)
                except Exception:
                    pass

                session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP24").select()
                self._sleep(0.3)
                if getattr(session, "Children", None) and session.Children.Count > 1:
                    pop = session.findById("wnd[1]")
                    try:
                        try:
                            pop.findById("usr/ctxtRMMG1-LGORT").text = almacen_str
                        except Exception:
                            pass
                        pop.findById("tbar[0]/btn[0]").press()
                        self._sleep(0.2)
                    except Exception:
                        try:
                            pop.sendVKey(0)
                        except Exception:
                            pass

                avisos.extend(handle_popups(session, delay=0.1))

                ctrl_vprsv = None
                try:
                    ctrl_vprsv = session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP24/ssubTABFRA1:SAPLMGMM:2000/subSUB2:SAPLMGD1:2800/subSUB2:SAPLMGD1:2802/ctxtMBEW-VPRSV")
                except Exception:
                    ctrl_vprsv = find_data_field_by_name(session.findById("wnd[0]"), "MBEW-VPRSV")

                if ctrl_vprsv:
                    control_precios_val = getattr(ctrl_vprsv, "Text", "").strip().upper()

                self.log(f"[MM03] Contabilidad: Control de Precios='{control_precios_val}'")
            except Exception as e:
                self.log(f"[MM03] Aviso en Contabilidad: {e}")

            # 9. Evaluación de Criterios y Dictamen Final
            stock_ok = (stock_negativo_activo is True)
            disp_ok = True  # Informativo: no bloquea ni marca discrepancia (MARC-MTVFP)
            auto_ok = (pedido_auto_activo is True)
            precio_ok = (str(control_precios_val).strip().upper() in ("S", "V"))

            peso_ok = True
            if peso_esp is not None and peso_volumetrico is not None:
                p_esp_num = clean_num(peso_esp)
                p_sap_num = clean_num(peso_volumetrico)
                if p_esp_num is not None and p_sap_num is not None:
                    peso_ok = (abs(p_esp_num - p_sap_num) <= 0.01)

            clase_ok = True
            if clase_esp is not None:
                c_esp_str = str(clase_esp).strip().upper()
                c_sap_mat = str(clase_material or "").strip().upper()
                c_sap_clas = str(clasificador_valor or "").strip().upper()
                clase_ok = (c_esp_str in c_sap_mat or c_esp_str in c_sap_clas)

            # Verificación de disponibilidad (MARC-MTVFP) es informativa (no marca discrepancia)
            todo_conforme = (stock_ok and auto_ok and precio_ok and peso_ok and clase_ok)
            estado = "OK" if todo_conforme else "DISCREPANCIA"

            detalles = []
            if stock_ok:
                detalles.append("✅ Stock Negativo: ACTIVO")
            else:
                detalles.append("❌ Stock Negativo: INACTIVO (debe estar marcado)")

            # Reportar valor de Verif. Disponibilidad (02, KP, etc.) sin marcar error
            detalles.append(f"ℹ️ Verif. Disponibilidad en SAP (MARC-MTVFP): {verif_disp_val or 'Vacío'}")

            if auto_ok:
                detalles.append("✅ Ind. Pedido Automático: ACTIVO")
            else:
                detalles.append("❌ Ind. Pedido Automático: INACTIVO (debe estar marcado)")

            if precio_ok:
                lbl_p = "Precio Estándar (S)" if control_precios_val == "S" else "Precio Medio Variable (V)"
                detalles.append(f"✅ Control de Precios: {control_precios_val} ({lbl_p})")
            else:
                detalles.append(f"❌ Control de Precios: '{control_precios_val or 'Vacío'}' (esperado 'S' o 'V')")

            if peso_volumetrico:
                p_msg = f"ℹ️ Peso Volumétrico en SAP: {peso_volumetrico}"
                if peso_esp:
                    p_msg += " (Coincide con esperado)" if peso_ok else f" (Difiere de esperado '{peso_esp}')"
                detalles.append(p_msg)
            else:
                detalles.append("⚠️ Peso Volumétrico no encontrado en unidades de medida")

            if clasificador_valor:
                detalles.append(f"ℹ️ Número Clasificador: {clasificador_valor}")
            elif clase_material:
                detalles.append(f"ℹ️ Clase de Material: {clase_material}")

            if avisos:
                detalles.append(f"Avisos SAP: {' | '.join(avisos)}")

            if todo_conforme:
                msg = f"MATERIAL VALIDADO OK (Stock Neg: SÍ | Disp: {verif_disp_val or 'N/A'} | Ped.Auto: SÍ | Ctrl Prec: {control_precios_val})"
            else:
                fallas = []
                if not stock_ok:
                    fallas.append("Stock Neg. Inactivo")
                if not auto_ok:
                    fallas.append("Ped. Auto Inactivo")
                if not precio_ok:
                    fallas.append(f"Ctrl Prec != S/V ({control_precios_val or 'Vacío'})")
                if not peso_ok:
                    fallas.append(f"Peso difiere ({peso_volumetrico} vs {peso_esp})")
                if not clase_ok:
                    fallas.append("Clasificación no coincide")
                msg = "DISCREPANCIA: " + " | ".join(fallas)

            return {
                "material": mat_str,
                "centro": centro_str,
                "org_ventas": org_str,
                "canal": canal_str,
                "almacen": almacen_str,
                "valido": todo_conforme,
                "estado": estado,
                "msg": msg,
                "detalles": detalles,
                "datos_extraidos": {
                    "peso_volumetrico": peso_volumetrico,
                    "clase_material": clase_material,
                    "clasificador": clasificador_valor,
                    "stock_negativo": stock_negativo_activo,
                    "verif_disponibilidad": verif_disp_val,
                    "pedido_automatico": pedido_auto_activo,
                    "control_precios": control_precios_val,
                },
                "checks": {
                    "stock_negativo_ok": stock_ok,
                    "verif_disp_ok": disp_ok,
                    "pedido_auto_ok": auto_ok,
                    "control_precios_ok": precio_ok,
                    "peso_volumetrico_ok": peso_ok,
                    "clasificacion_ok": clase_ok,
                },
                "avisos": avisos,
            }

        except Exception as e:
            self.log(f"[MM03] Excepción durante validación: {e}")
            return {
                "material": mat_str,
                "centro": centro_str,
                "org_ventas": org_str,
                "canal": canal_str,
                "almacen": almacen_str,
                "valido": False,
                "estado": "ERROR",
                "msg": f"Excepción técnica en SAP: {e}",
                "detalles": [f"❌ Error inesperado: {e}"],
                "datos_extraidos": {},
                "checks": {
                    "stock_negativo_ok": False,
                    "verif_disp_ok": False,
                    "pedido_auto_ok": False,
                    "control_precios_ok": False,
                    "peso_volumetrico_ok": False,
                    "clasificacion_ok": False,
                },
                "avisos": [],
            }

    def validar_lote_materiales(self, items, ruta_excel=None):
        """Valida una lista de materiales en MM03 reportando progreso en tiempo real."""
        self.stop_requested = False
        self.conectar()
        resultados = []
        total = len(items)
        self.log(f"=== INICIANDO VALIDACIÓN DE {total} MATERIALES EN MM03 ===")

        for idx, item in enumerate(items, 1):
            if self.stop_requested:
                self.log("Validación detenida por el usuario.")
                break

            mat = item.get("material", "")
            centro = item.get("centro", "")
            org_vta = item.get("org_ventas", "")
            canal = item.get("canal", "00")
            almacen = item.get("almacen", "1002")
            peso_esp = item.get("peso_esp")
            clase_esp = item.get("clase_esp")

            if self.on_progress:
                self.on_progress(idx, total, item, "EN_PROCESO", {}, f"Consultando Material {mat} en MM03...")

            res = self.validar_material_mm03(
                mat, centro, org_vta, canal=canal, almacen=almacen, peso_esp=peso_esp, clase_esp=clase_esp
            )
            resultados.append(res)

            if self.on_progress:
                self.on_progress(idx, total, item, res["estado"], res["datos_extraidos"], res["msg"])

            self._sleep(0.2)

        self.log(f"=== VALIDACIÓN FINALIZADA ({len(resultados)}/{total} PROCESADOS) ===")
        return resultados


# ---------------------------------------------------------------------------
# Aplicación de Confirmaciones a Base de Datos y Excel Maestro
# ---------------------------------------------------------------------------
def aplicar_confirmacion_lote(conn, filas_aprobadas, fecha_conf=None):
    """Actualiza la Fecha de Confirmación en SQLite (staging_rutas) y en el Excel Maestro.

    filas_aprobadas: lista de dicts de items que tuvieron estado OK.
    fecha_conf: string 'dd/mm/aaaa' (por defecto fecha actual).
    Retorna (num_sqlite, num_excel, detalles).
    """
    if not fecha_conf:
        fecha_conf = datetime.now().strftime("%d/%m/%Y")

    if not filas_aprobadas:
        return 0, 0, "No hay filas aprobadas para confirmar."

    # Convertir y deduplicar por clave de ruta única
    seen_keys = set()
    rows_fpa = []
    for item in filas_aprobadas:
        sf = item.get("sf", "")
        cedis = item.get("cedis", "")
        destino = item.get("destino", "")
        material = item.get("material", "")
        modalidad = item.get("modalidad", "01")

        r_key = (str(sf).strip(), str(cedis).strip(), str(destino).strip(), str(material).strip(), str(modalidad).strip())
        if r_key in seen_keys:
            continue
        seen_keys.add(r_key)

        rows_fpa.append({
            "ShipFrom": sf,
            "Centro": cedis,
            "No. De Destino SAP": destino,
            "Material": material,
            "Modalidad": modalidad,
            "sf": sf,
            "cedis": cedis,
            "planta_obra": destino,
            "articulo": material,
            "modalidad": modalidad,
        })

    # 1. Actualizar SQLite
    ids_sqlite = fpa_validations.actualizar_fecha_confirmacion_sqlite(conn, rows_fpa, fecha=fecha_conf)

    # 2. Actualizar Excel Maestro (hoja BASE 2026, columna AO FECHA CONF)
    filas_excel = fpa_validations.actualizar_fecha_confirmacion_excel(rows_fpa, fecha=fecha_conf)

    detalles = (
        f"Confirmación completada con fecha {fecha_conf}:\n"
        f"• Rutas únicas confirmadas: {len(rows_fpa)}\n"
        f"• SQLite (staging_rutas): {len(ids_sqlite)} registro(s) actualizados.\n"
        f"• Excel Maestro (BASE 2026): {len(filas_excel)} fila(s) actualizadas."
    )
    return len(ids_sqlite), len(filas_excel), detalles


# ---------------------------------------------------------------------------
# Extracción y Exportación Excel para Revisión de Materiales MM03
# ---------------------------------------------------------------------------
def extraer_filas_materiales_excel(ruta_excel):
    """Extrae filas de materiales desde un archivo Excel para validación en MM03."""
    df = pd.read_excel(ruta_excel, dtype=str)
    col_map = {}
    for col in df.columns:
        c_clean = str(col).strip().upper()
        if any(k in c_clean for k in ["MATERIAL", "MATNR", "ARTICULO", "ARTÍCULO", "PRODUCTO"]):
            col_map["material"] = col
        elif any(k in c_clean for k in ["CENTRO", "WERKS", "CEDIS", "PLANTA"]):
            col_map["centro"] = col
        elif any(k in c_clean for k in ["ORG", "VKORG", "VENTAS"]):
            col_map["org_ventas"] = col
        elif any(k in c_clean for k in ["CANAL", "VTWEG"]):
            col_map["canal"] = col
        elif any(k in c_clean for k in ["ALMACEN", "ALMACÉN", "LGORT"]):
            col_map["almacen"] = col
        elif any(k in c_clean for k in ["PESO", "VOLUMETRICO", "VOLUMÉTRICO"]):
            col_map["peso_esp"] = col
        elif any(k in c_clean for k in ["CLASE", "CLASIFICACION", "CLASIFICACIÓN"]):
            col_map["clase_esp"] = col

    if "material" not in col_map and len(df.columns) >= 3:
        col_map["material"] = df.columns[0]
        col_map["centro"] = df.columns[1]
        col_map["org_ventas"] = df.columns[2]

    items = []
    for idx, row in df.iterrows():
        mat = str(row.get(col_map.get("material", ""), "")).strip()
        if not mat or mat.lower() in ("nan", "none", ""):
            continue
        centro = str(row.get(col_map.get("centro", ""), "")).strip()
        if centro.lower() in ("nan", "none"):
            centro = ""
        org_vta = str(row.get(col_map.get("org_ventas", ""), "")).strip()
        if org_vta.lower() in ("nan", "none"):
            org_vta = ""
        canal = str(row.get(col_map.get("canal", ""), "00")).strip()
        if not canal or canal.lower() in ("nan", "none"):
            canal = "00"
        almacen = str(row.get(col_map.get("almacen", ""), "1002")).strip()
        if not almacen or almacen.lower() in ("nan", "none"):
            almacen = "1002"

        peso_esp = row.get(col_map.get("peso_esp", "")) if "peso_esp" in col_map else None
        if pd.isna(peso_esp):
            peso_esp = None
        clase_esp = row.get(col_map.get("clase_esp", "")) if "clase_esp" in col_map else None
        if pd.isna(clase_esp):
            clase_esp = None

        items.append({
            "idx_fila": idx + 2,
            "material": mat,
            "centro": centro,
            "org_ventas": org_vta,
            "canal": canal,
            "almacen": almacen,
            "peso_esp": str(peso_esp).strip() if peso_esp else None,
            "clase_esp": str(clase_esp).strip() if clase_esp else None,
        })
    return items


def exportar_resultados_materiales_excel(ruta_salida, resultados):
    """Exporta reporte detallado de validación MM03 a un archivo Excel."""
    rows = []
    for r in resultados:
        de = r.get("datos_extraidos", {})
        chk = r.get("checks", {})
        rows.append({
            "Estado Validación": r.get("estado", ""),
            "Material": r.get("material", ""),
            "Centro": r.get("centro", ""),
            "Org. Ventas": r.get("org_ventas", ""),
            "Canal": r.get("canal", "00"),
            "Almacén": r.get("almacen", "1002"),
            "Stock Negativo": "ACTIVO" if de.get("stock_negativo") else "INACTIVO",
            "Verif. Disponibilidad": de.get("verif_disponibilidad", ""),
            "Ind. Pedido Automático": "ACTIVO" if de.get("pedido_automatico") else "INACTIVO",
            "Control de Precios": de.get("control_precios", ""),
            "Peso Volumétrico": de.get("peso_volumetrico", ""),
            "Clase Material": de.get("clase_material", ""),
            "Clasificador": de.get("clasificador", ""),
            "Dictamen / Resumen": r.get("msg", ""),
            "Detalles Técnicos": " | ".join(r.get("detalles", [])),
        })
    df_out = pd.DataFrame(rows)
    df_out.to_excel(ruta_salida, index=False)
    return ruta_salida
