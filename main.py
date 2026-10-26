"""RutaV3_2026 - Sistema de gestión corporativa CEMEX México."""

import os
import threading
from datetime import date, datetime

import customtkinter as ctk
from customtkinter import CTkFont
import tkinter.messagebox as messagebox
from tkinter import BooleanVar, StringVar, filedialog, ttk

import pandas as pd
from tkinterdnd2 import TkinterDnD

import database_manager
import etl_processor
import fpa_validations
import sync_manager
import sap_validator_engine

import cell_grid

# ---------------------------------------------------------------------------
# Rutas relativas al directorio del script (portables entre equipos)
# ---------------------------------------------------------------------------
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
DB_FILE = os.path.join(BASE_DIR, 'staging_rutas.db')
EXPORT_DIR = os.path.join(BASE_DIR, 'Exportaciones_Rutas')

# ---------------------------------------------------------------------------
# Identidad visual CEMEX Digital Design System
# ---------------------------------------------------------------------------
CEMEX_BLUE = "#023185"
CEMEX_BLUE_DARK = "#01245F"
CEMEX_BLUE_LIGHT = "#EAF0F8"
CEMEX_RED = "#E3303D"
CEMEX_RED_DARK = "#B02630"
CEMEX_GRAY_BG = "#F4F6F9"
CEMEX_WHITE = "#FFFFFF"
CEMEX_TEXT = "#1A1A1A"
CEMEX_LIGHT_GRAY = "#DDE1E6"
CEMEX_SIDEBAR = "#023185"
CEMEX_SIDEBAR_DARK = "#01245F"
CEMEX_SIDEBAR_MUTED = "#C7D3E8"
CEMEX_NEUTRAL = "#6C757D"
CEMEX_NEUTRAL_DARK = "#5A6268"

# ---------------------------------------------------------------------------
# Campos del formulario
# ---------------------------------------------------------------------------
FIELDS_IDENTIFICACION = [
    ("org_vta", "ORG VTA", ""),
    ("cedis", "CEDIS", ""),
    ("nombre_cedis", "NOMBRE CEDIS", ""),
    ("sf", "SF", ""),
    ("nombre_sf", "NOMBRE SF", ""),
    ("planta_obra", "NO PLANTA / OBRA", ""),
    ("nombre_planta", "NOMBRE PLANTA / OBRA", ""),
    ("articulo", "ARTÍCULO", ""),
    ("desc_articulo", "DESCRIPCIÓN ARTÍCULO", ""),
    ("pv", "PV", ""),
    ("modalidad", "MODALIDAD", "01"),
    ("um_vta", "UM VTA", "M3"),
]

FIELDS_COSTOS = [
    ("mp_vta_m3", "MP VTA (M3)", ""),
    ("flete_vta_m3", "FLETE VTA (M3)", ""),
    ("mp_comp_m3", "MP COMP (M3)", ""),
    ("ca_comp_m3", "CA COMP (M3)", ""),
    ("flete_comp_m3", "FLETE COMP (M3)", ""),
    ("integral_comp_m3", "INTEGRAL COMP (M3)", ""),
    ("tipo", "TIPO", "ALTA"),
    ("fecha_procesamiento", "FECHA DE PROCESAMIENTO", ""),
    ("tipo_cliente", "TIPO CLIENTE", "TERCEROS"),
    ("observacion", "OBSERVACIÓN", ""),
    ("fecha_solicitud", "FECHA SOLICITUD", ""),
    ("incidencia", "INCIDENCIA?", "NO"),
    ("motivo", "MOTIVO", ""),
    ("mes", "MES", ""),
    ("fecha_conf", "FECHA CONF", ""),
    ("formato_destino", "FORMATO DESTINO", ""),
]

# Orden estricto de los 11 índices al pegar desde Excel (Ctrl+V)
PASTE_FIELDS = [
    "org_vta", "cedis", "nombre_cedis", "sf", "nombre_sf",
    "planta_obra", "nombre_planta", "articulo", "desc_articulo", "pv", "modalidad",
]

# Grupos de campos de costos para pegado inteligente (izquierda -> derecha)
COST_PASTE_GROUPS = {
    "mp_vta_m3": ("mp_vta_m3", "flete_vta_m3"),
    "flete_vta_m3": ("flete_vta_m3",),
    "mp_comp_m3": ("mp_comp_m3", "ca_comp_m3", "flete_comp_m3"),
    "ca_comp_m3": ("ca_comp_m3", "flete_comp_m3"),
    "flete_comp_m3": ("flete_comp_m3",),
}

# Columnas visibles del Treeview de filas pendientes (Excel)
PENDING_COLUMNS = (
    "ESTADO", "ORG VTA", "CEDIS", "NOMBRE CEDIS", "SF", "NOMBRE SF",
    "NO PLANTA", "NOMBRE PLANTA", "ARTÍCULO", "DESC ARTÍCULO", "PV", "MODALIDAD"
)

# Columnas visibles del Treeview de Validaciones / SLA
SLA_COLUMNS = (
    "ID", "RUTA", "ORGVTA", "CEDIS", "SF", "NO PLANTA", "ARTÍCULO",
    "MODALIDAD", "FECHA SOLICITUD", "FECHA PROCESAMIENTO", "FECHA CONF",
    "DÍAS HÁBILES", "DÍAS TRANSCURRIDOS", "ESTADO"
)


APP_VERSION = "1.2.1"
APP_BUILD = "FACV · BETA"


def normalizar_fecha(fecha_str):
    """Normaliza una fecha ingresada a dd/mm/yyyy.

    Reemplaza puntos y guiones por barras y estandariza el orden de día/mes/año.
    Si no puede parsear el valor, devuelve la cadena original limpia (con
    separadores ya convertidos a '/'), de modo que nunca rompa la consulta.
    """
    if not fecha_str:
        return ""
    try:
        s = str(fecha_str).strip()
        s = s.replace(".", "/").replace("-", "/")
        from datetime import datetime
        for fmt in ("%d/%m/%Y", "%Y/%m/%d", "%d/%m/%y"):
            try:
                return datetime.strptime(s, fmt).strftime("%d/%m/%Y")
            except ValueError:
                continue
        return s
    except Exception:
        return str(fecha_str).strip()


class App(ctk.CTk, TkinterDnD.DnDWrapper):
    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        self.TkdndVersion = TkinterDnD._require(self)

        self.title(f"RutaV3_2026 | {APP_BUILD} v{APP_VERSION} | CEMEX")
        self.geometry("1400x980")
        self.configure(fg_color=CEMEX_GRAY_BG)

        # Estado
        self.vars = {}
        self.entries = {}
        self.formato_checkboxes = {}
        self.last_calculated = None
        self.sap_data = None
        self.conn = None
        self.batch_rows = []
        self.pending_rows = []
        self.pending_selected_index = None

        self.active_view = "captura"
        self.nav_buttons = {}
        self.sla_search_vars = {}

        # Estado del Área de Validaciones (FPA)
        self.fpa_rows = []
        self.fpa_selected_index = None
        self.fpa_db_matches = []
        self.fpa_excel_matches = []
        self.fpa_selected_request = None
        self.fpa_excel_row_seleccionada = None

        # Estado del Validador Visual SAP GUI
        self.val_sap_items = []
        self.val_sap_file_path = ""
        self.val_sap_tipo_info = None
        self.val_sap_motor = None
        self.val_sap_thread = None
        self.val_sap_approved_items = []

        # Variables para Matriz Precios
        self.val_matriz_items = []
        self.val_matriz_filtered_items = []
        self.val_matriz_file_path = ""
        self.val_matriz_out_path = ""
        self.val_matriz_tipo_info = None
        self.val_matriz_motor = None

        # Variables para Validación de Materiales (MM03)
        self.val_mat_items = []
        self.val_mat_file_path = ""
        self.val_mat_motor = None
        self.val_mat_resultados = []

        self._configure_ttk_style()
        self._build_layout()
        self._init_db()
        self._refresh_sla_tree()
        self._refresh_sync_status()

        # Habilitar Drag & Drop de archivos Excel.
        self._habilitar_drag_drop()

        # Atajos globales
        self.bind("<Return>", lambda e: self.calculate_row())
        self.bind("<Control-v>", self._on_global_paste)
        self.bind("<F4>", lambda e: self.toggle_mantener_mp())
        self.bind("<F5>", lambda e: self.toggle_mantener_flete())
        self.bind("<F6>", lambda e: self.fpa_buscar())
        self.bind("<F8>", lambda e: self.fpa_validar())

    # ------------------------------------------------------------------
    # Estilo de tablas (ttk.Style) - identidad CEMEX
    # ------------------------------------------------------------------
    def _configure_ttk_style(self):
        style = ttk.Style(self)
        try:
            style.theme_use("clam")
        except Exception:
            pass

        style.configure(
            "Cemex.Treeview",
            background=CEMEX_WHITE,
            fieldbackground=CEMEX_WHITE,
            foreground=CEMEX_TEXT,
            borderwidth=0,
            relief="flat",
            rowheight=30,
            font=("Segoe UI", 11),
        )
        style.configure(
            "Cemex.Treeview.Heading",
            background=CEMEX_BLUE,
            foreground=CEMEX_WHITE,
            font=("Segoe UI", 11, "bold"),
            borderwidth=0,
            relief="flat",
            padding=(6, 8),
        )
        style.map(
            "Cemex.Treeview.Heading",
            background=[("active", CEMEX_BLUE_DARK)],
            relief=[("active", "flat"), ("pressed", "flat")],
        )
        style.map(
            "Cemex.Treeview",
            background=[("selected", CEMEX_BLUE)],
            foreground=[("selected", CEMEX_WHITE)],
        )
        style.configure(
            "Cemex.Treeview.Even",
            background=CEMEX_WHITE,
            foreground=CEMEX_TEXT,
        )
        style.configure(
            "Cemex.Treeview.Odd",
            background=CEMEX_GRAY_BG,
            foreground=CEMEX_TEXT,
        )

    # ------------------------------------------------------------------
    # Construcción de la GUI
    # ------------------------------------------------------------------
    def _build_layout(self):
        self.grid_columnconfigure(0, weight=0, minsize=260)
        self.grid_columnconfigure(1, weight=1)
        self.grid_rowconfigure(0, weight=1)

        self._build_sidebar()
        self._build_content()

    def _build_sidebar(self):
        sidebar = ctk.CTkFrame(
            self,
            fg_color=CEMEX_SIDEBAR,
            corner_radius=0,
            width=260,
        )
        self.sidebar_frame = sidebar
        sidebar.grid(row=0, column=0, sticky="nsw")
        sidebar.grid_propagate(False)
        sidebar.grid_rowconfigure(0, weight=0)
        sidebar.grid_rowconfigure(1, weight=0)
        sidebar.grid_rowconfigure(2, weight=1)
        sidebar.grid_rowconfigure(3, weight=0)

        # Bloque de marca
        brand = ctk.CTkFrame(sidebar, fg_color="transparent")
        brand.grid(row=0, column=0, padx=24, pady=(30, 26), sticky="ew")
        brand.grid_columnconfigure(0, weight=1)

        ctk.CTkLabel(
            brand,
            text="RutaV3_2026",
            font=CTkFont("Segoe UI", 22, "bold"),
            text_color=CEMEX_WHITE,
            anchor="w",
        ).grid(row=0, column=0, sticky="ew")

        ctk.CTkLabel(
            brand,
            text="CEMEX · Gestión Corporativa",
            font=CTkFont("Segoe UI", 11),
            text_color=CEMEX_SIDEBAR_MUTED,
            anchor="w",
        ).grid(row=1, column=0, pady=(4, 0), sticky="ew")

        # Acento rojo bajo la marca
        accent = ctk.CTkFrame(brand, fg_color=CEMEX_RED, height=3, corner_radius=2)
        accent.grid(row=2, column=0, pady=(16, 0), sticky="ew")
        accent.grid_propagate(False)

        # Separador
        divider = ctk.CTkFrame(sidebar, fg_color=CEMEX_SIDEBAR_MUTED, height=1)
        divider.grid(row=1, column=0, padx=24, sticky="ew")

        # Contenedor scrolleable para NAVEGACIÓN + SINCRONIZACIÓN.
        scroll_body = ctk.CTkScrollableFrame(
            sidebar,
            fg_color="transparent",
            scrollbar_button_color=CEMEX_SIDEBAR_MUTED,
            scrollbar_button_hover_color=CEMEX_WHITE,
        )
        scroll_body.grid(row=2, column=0, sticky="nsew")
        scroll_body.grid_columnconfigure(0, weight=1)

        # Navegación
        nav = ctk.CTkFrame(scroll_body, fg_color="transparent")
        nav.grid(row=0, column=0, padx=16, pady=(24, 0), sticky="ew")
        nav.grid_columnconfigure(0, weight=1)

        ctk.CTkLabel(
            nav,
            text="NAVEGACIÓN",
            font=CTkFont("Segoe UI", 10, "bold"),
            text_color=CEMEX_SIDEBAR_MUTED,
            anchor="w",
        ).grid(row=0, column=0, padx=8, pady=(0, 10), sticky="ew")

        self.nav_buttons["captura"] = ctk.CTkButton(
            nav,
            text="Mesa de Captura",
            command=lambda: self._switch_view("captura"),
            anchor="w",
            height=42,
            corner_radius=8,
            fg_color="transparent",
            hover_color=CEMEX_SIDEBAR_DARK,
            text_color=CEMEX_WHITE,
            font=CTkFont("Segoe UI", 13, "bold"),
        )
        self.nav_buttons["captura"].grid(row=1, column=0, pady=(0, 8), sticky="ew")

        self.nav_buttons["validaciones"] = ctk.CTkButton(
            nav,
            text="Validaciones y SLA",
            command=lambda: self._switch_view("validaciones"),
            anchor="w",
            height=42,
            corner_radius=8,
            fg_color="transparent",
            hover_color=CEMEX_SIDEBAR_DARK,
            text_color=CEMEX_WHITE,
            font=CTkFont("Segoe UI", 13, "bold"),
        )
        self.nav_buttons["validaciones"].grid(row=2, column=0, pady=(0, 8), sticky="ew")

        self.nav_buttons["fpa"] = ctk.CTkButton(
            nav,
            text="Área de Validaciones (FPA)",
            command=lambda: self._switch_view("fpa"),
            anchor="w",
            height=42,
            corner_radius=8,
            fg_color="transparent",
            hover_color=CEMEX_SIDEBAR_DARK,
            text_color=CEMEX_WHITE,
            font=CTkFont("Segoe UI", 13, "bold"),
        )
        self.nav_buttons["fpa"].grid(row=3, column=0, pady=(0, 8), sticky="ew")

        self.nav_buttons["val_sap"] = ctk.CTkButton(
            nav,
            text="Validación SAP GUI",
            command=lambda: self._switch_view("val_sap"),
            anchor="w",
            height=42,
            corner_radius=8,
            fg_color="transparent",
            hover_color=CEMEX_SIDEBAR_DARK,
            text_color=CEMEX_WHITE,
            font=CTkFont("Segoe UI", 13, "bold"),
        )
        self.nav_buttons["val_sap"].grid(row=4, column=0, pady=(0, 8), sticky="ew")

        self.nav_buttons["val_matriz"] = ctk.CTkButton(
            nav,
            text="Validar Matriz Precios",
            command=lambda: self._switch_view("val_matriz"),
            anchor="w",
            height=42,
            corner_radius=8,
            fg_color="transparent",
            hover_color=CEMEX_SIDEBAR_DARK,
            text_color=CEMEX_WHITE,
            font=CTkFont("Segoe UI", 13, "bold"),
        )
        self.nav_buttons["val_matriz"].grid(row=5, column=0, pady=(0, 8), sticky="ew")

        self.nav_buttons["val_material"] = ctk.CTkButton(
            nav,
            text="Revisión Materiales (MM03)",
            command=lambda: self._switch_view("val_material"),
            anchor="w",
            height=42,
            corner_radius=8,
            fg_color="transparent",
            hover_color=CEMEX_SIDEBAR_DARK,
            text_color=CEMEX_WHITE,
            font=CTkFont("Segoe UI", 13, "bold"),
        )
        self.nav_buttons["val_material"].grid(row=6, column=0, pady=(0, 8), sticky="ew")

        # Bloque de sincronización manual (drive)
        sync = ctk.CTkFrame(scroll_body, fg_color="transparent")
        sync.grid(row=1, column=0, padx=16, pady=(8, 16), sticky="ew")
        sync.grid_columnconfigure(0, weight=1)

        ctk.CTkLabel(
            sync,
            text="SINCRONIZACIÓN",
            font=CTkFont("Segoe UI", 10, "bold"),
            text_color=CEMEX_SIDEBAR_MUTED,
            anchor="w",
        ).grid(row=0, column=0, padx=8, pady=(0, 6), sticky="ew")

        self.upload_btn = ctk.CTkButton(
            sync,
            text="Subir al drive",
            command=self.sync_upload,
            height=36,
            corner_radius=8,
            fg_color=CEMEX_WHITE,
            hover_color=CEMEX_BLUE_LIGHT,
            text_color=CEMEX_BLUE,
            font=CTkFont("Segoe UI", 12, "bold"),
        )
        self.upload_btn.grid(row=1, column=0, pady=(0, 6), sticky="ew")

        self.download_btn = ctk.CTkButton(
            sync,
            text="Cargar datos",
            command=self.sync_download,
            height=36,
            corner_radius=8,
            fg_color=CEMEX_WHITE,
            hover_color=CEMEX_BLUE_LIGHT,
            text_color=CEMEX_BLUE,
            font=CTkFont("Segoe UI", 12, "bold"),
        )
        self.download_btn.grid(row=2, column=0, pady=(0, 6), sticky="ew")

        self.recuperacion_btn = ctk.CTkButton(
            sync,
            text="Sincronización de Recuperación",
            command=self.sync_recuperacion,
            height=36,
            corner_radius=8,
            fg_color=CEMEX_RED,
            hover_color=CEMEX_RED_DARK,
            text_color=CEMEX_WHITE,
            font=CTkFont("Segoe UI", 12, "bold"),
        )
        self.recuperacion_btn.grid(row=3, column=0, pady=(0, 6), sticky="ew")

        self.sync_status_label = ctk.CTkLabel(
            sync,
            text="",
            font=CTkFont("Segoe UI", 9),
            text_color=CEMEX_SIDEBAR_MUTED,
            anchor="w",
            justify="left",
            wraplength=210,
        )
        self.sync_status_label.grid(row=4, column=0, padx=8, pady=(2, 6), sticky="ew")

        self.sql_terminal_btn = ctk.CTkButton(
            sync,
            text="Terminal SQL",
            command=self._abrir_terminal_sql,
            height=32,
            corner_radius=8,
            fg_color=CEMEX_NEUTRAL,
            hover_color=CEMEX_NEUTRAL_DARK,
            text_color=CEMEX_WHITE,
            font=CTkFont("Segoe UI", 11, "bold"),
        )
        self.sql_terminal_btn.grid(row=5, column=0, pady=(0, 6), sticky="ew")

        # Pie de barra lateral
        footer = ctk.CTkFrame(sidebar, fg_color="transparent")
        footer.grid(row=3, column=0, padx=24, pady=20, sticky="sew")

        ctk.CTkLabel(
            footer,
            text=f"{APP_BUILD} v{APP_VERSION}",
            font=CTkFont("Segoe UI", 10, "bold"),
            text_color=CEMEX_SIDEBAR_MUTED,
            anchor="w",
        ).grid(row=0, column=0, sticky="w")

        ctk.CTkLabel(
            footer,
            text="Sistema de Diseño CEMEX",
            font=CTkFont("Segoe UI", 10),
            text_color=CEMEX_SIDEBAR_MUTED,
            anchor="w",
        ).grid(row=1, column=0, pady=(2, 0), sticky="w")

    def _build_content(self):
        content = ctk.CTkFrame(self, fg_color=CEMEX_GRAY_BG, corner_radius=0)
        content.grid(row=0, column=1, sticky="nsew")
        content.grid_rowconfigure(0, weight=0)
        content.grid_columnconfigure(0, weight=1)

        # Botón hamburguesa fijo (fuera del sidebar, para poder re-expandirlo
        # cuando esté colapsado). Colocado en su propia fila superior.
        self.hamburger_btn = ctk.CTkButton(
            content, text="☰", width=36, height=32,
            command=self._toggle_sidebar,
            fg_color=CEMEX_NEUTRAL, hover_color=CEMEX_NEUTRAL_DARK,
            font=CTkFont("Segoe UI", 16, "bold"),
        )
        self.hamburger_btn.grid(row=0, column=0, padx=8, pady=6, sticky="nw")

        # Contenedor de vistas en la fila inferior
        views_host = ctk.CTkFrame(content, fg_color=CEMEX_GRAY_BG, corner_radius=0)
        views_host.grid(row=1, column=0, sticky="nsew")
        views_host.grid_rowconfigure(0, weight=1)
        views_host.grid_columnconfigure(0, weight=1)
        content.grid_rowconfigure(1, weight=1)

        self.captura_view = ctk.CTkFrame(views_host, fg_color=CEMEX_GRAY_BG, corner_radius=0)
        self.captura_view.grid(row=0, column=0, sticky="nsew")

        self.validaciones_view = ctk.CTkFrame(views_host, fg_color=CEMEX_GRAY_BG, corner_radius=0)
        self.validaciones_view.grid(row=0, column=0, sticky="nsew")
        self.validaciones_view.grid_remove()

        self.fpa_view = ctk.CTkFrame(views_host, fg_color=CEMEX_GRAY_BG, corner_radius=0)
        self.fpa_view.grid(row=0, column=0, sticky="nsew")
        self.fpa_view.grid_remove()

        self.val_sap_view = ctk.CTkFrame(views_host, fg_color=CEMEX_GRAY_BG, corner_radius=0)
        self.val_sap_view.grid(row=0, column=0, sticky="nsew")
        self.val_sap_view.grid_remove()

        self.val_matriz_view = ctk.CTkFrame(views_host, fg_color=CEMEX_GRAY_BG, corner_radius=0)
        self.val_matriz_view.grid(row=0, column=0, sticky="nsew")
        self.val_matriz_view.grid_remove()

        self.val_material_view = ctk.CTkFrame(views_host, fg_color=CEMEX_GRAY_BG, corner_radius=0)
        self.val_material_view.grid(row=0, column=0, sticky="nsew")
        self.val_material_view.grid_remove()

        self._build_tab_captura()
        self._build_tab_validaciones()
        self._build_tab_fpa()
        self._build_tab_val_sap()
        self._build_tab_val_matriz()
        self._build_tab_val_material()
        self._style_nav_buttons()

    def _switch_view(self, name):
        self.active_view = name
        self.captura_view.grid_remove()
        self.validaciones_view.grid_remove()
        self.fpa_view.grid_remove()
        self.val_sap_view.grid_remove()
        self.val_matriz_view.grid_remove()
        self.val_material_view.grid_remove()
        if name == "captura":
            self.captura_view.grid()
        elif name == "validaciones":
            self.validaciones_view.grid()
        elif name == "fpa":
            self.fpa_view.grid()
        elif name == "val_sap":
            self.val_sap_view.grid()
        elif name == "val_matriz":
            self.val_matriz_view.grid()
        elif name == "val_material":
            self.val_material_view.grid()
        self._style_nav_buttons()

    def _toggle_sidebar(self):
        """Colapsa/expande la barra lateral para dar más espacio a las tablas."""
        if self.sidebar_frame.winfo_ismapped():
            self.sidebar_frame.grid_remove()
            self.grid_columnconfigure(0, weight=0, minsize=0)
        else:
            self.sidebar_frame.grid()
            self.grid_columnconfigure(0, weight=0, minsize=260)

    def _style_nav_buttons(self):
        for name, btn in self.nav_buttons.items():
            if name == self.active_view:
                btn.configure(
                    fg_color=CEMEX_WHITE,
                    text_color=CEMEX_BLUE,
                    hover_color=CEMEX_BLUE_LIGHT,
                )
            else:
                btn.configure(
                    fg_color="transparent",
                    text_color=CEMEX_WHITE,
                    hover_color=CEMEX_SIDEBAR_DARK,
                )

    # ------------------------------------------------------------------
    # Pestaña: Mesa de Captura
    # ------------------------------------------------------------------
    def _build_tab_captura(self):
        tab = self.captura_view
        tab.grid_columnconfigure(0, weight=1)
        tab.grid_rowconfigure(0, weight=1)

        scroll = ctk.CTkScrollableFrame(tab, fg_color=CEMEX_GRAY_BG)
        scroll.grid(row=0, column=0, sticky="nsew")
        scroll.grid_columnconfigure((0, 1), weight=1)

        # Frame 1: Datos de Identificación
        self.frame_id = self._create_card(scroll, "Datos de Identificación")
        self.frame_id.grid(row=0, column=0, padx=15, pady=15, sticky="nsew")
        self._build_form_section(self.frame_id, FIELDS_IDENTIFICACION)

        # Frame 2: Captura de Costos
        self.frame_costos = self._create_card(scroll, "Captura de Costos")
        self.frame_costos.grid(row=0, column=1, padx=15, pady=15, sticky="nsew")
        self._build_form_section(self.frame_costos, FIELDS_COSTOS)
        self._build_mantener_precio_row(self.frame_costos)

        # Frame 3: Panel de Acciones
        self.frame_acciones = self._create_card(scroll, "Panel de Acciones y Lote Actual")
        self.frame_acciones.grid(row=1, column=0, columnspan=2, padx=15, pady=(0, 15), sticky="ew")
        self._build_actions_section(self.frame_acciones)

        # Frame 4: HUD horizontal de filas pendientes desde Excel
        self.frame_pending = self._create_card(scroll, "Lote pendiente desde Excel (Ctrl+V)")
        self.frame_pending.grid(row=2, column=0, columnspan=2, padx=15, pady=(0, 15), sticky="nsew")
        self._build_pending_tree(self.frame_pending)

    def _create_card(self, parent, title):
        card = ctk.CTkFrame(
            parent,
            fg_color=CEMEX_WHITE,
            corner_radius=12,
            border_width=1,
            border_color=CEMEX_LIGHT_GRAY,
        )
        card.grid_columnconfigure((0, 1), weight=1)

        ctk.CTkLabel(
            card,
            text=title.upper(),
            font=CTkFont("Segoe UI", 14, "bold"),
            text_color=CEMEX_BLUE,
        ).grid(row=0, column=0, columnspan=2, padx=18, pady=(14, 8), sticky="w")
        return card

    def _build_form_section(self, parent, fields):
        for i, (key, label, default) in enumerate(fields, start=1):
            ctk.CTkLabel(
                parent,
                text=label,
                font=CTkFont("Segoe UI", 12),
                text_color=CEMEX_TEXT,
                width=220,
                anchor="w",
            ).grid(row=i, column=0, padx=(18, 8), pady=5, sticky="w")

            var = StringVar(value=default)
            self.vars[key] = var

            if key == "formato_destino":
                widget = ctk.CTkFrame(parent, fg_color="transparent")
                widget.grid_columnconfigure((0, 1), weight=1)
                opciones = [
                    ("FPA Filial", "FPA Filial"),
                    ("FPA Terceros", "FPA Terceros"),
                    ("CTO Venta", "CTO Venta"),
                    ("CTO Compra", "CTO Compra"),
                ]
                for idx, (text, value) in enumerate(opciones):
                    cb = ctk.CTkCheckBox(
                        widget,
                        text=text,
                        font=CTkFont("Segoe UI", 11),
                        text_color=CEMEX_TEXT,
                        checkbox_width=18,
                        checkbox_height=18,
                        fg_color=CEMEX_BLUE,
                        hover_color=CEMEX_BLUE_DARK,
                        border_color=CEMEX_LIGHT_GRAY,
                    )
                    cb.grid(row=idx // 2, column=idx % 2, padx=6, pady=4, sticky="w")
                    self.formato_checkboxes[value] = cb
                widget.grid(row=i, column=1, padx=(8, 18), pady=5, sticky="ew")
            elif key in ("modalidad", "um_vta", "tipo", "tipo_cliente", "incidencia"):
                ctk.CTkOptionMenu(
                    parent,
                    variable=var,
                    values=self._options_for(key),
                    font=CTkFont("Segoe UI", 12),
                    dropdown_font=CTkFont("Segoe UI", 11),
                    width=280,
                    fg_color=CEMEX_BLUE,
                    button_color=CEMEX_BLUE,
                    button_hover_color=CEMEX_BLUE_DARK,
                    dropdown_hover_color=CEMEX_BLUE_LIGHT,
                    text_color=CEMEX_WHITE,
                ).grid(row=i, column=1, padx=(8, 18), pady=5, sticky="ew")
            else:
                entry = ctk.CTkEntry(
                    parent,
                    textvariable=var,
                    font=CTkFont("Segoe UI", 12),
                    width=280,
                    fg_color=CEMEX_WHITE,
                    border_color=CEMEX_LIGHT_GRAY,
                    text_color="#1A1A1A",
                    placeholder_text_color="#8C92AC",
                )
                entry.grid(row=i, column=1, padx=(8, 18), pady=5, sticky="ew")
                self.entries[key] = entry
                if key in COST_PASTE_GROUPS:
                    entry.bind("<Control-v>", lambda e, k=key: self._smart_paste_cost(k, e))

    def _build_mantener_precio_row(self, parent):
        """Fila inferior del card de costos con las casillas de precios persistentes."""
        row = ctk.CTkFrame(parent, fg_color="transparent")
        row.grid(row=len(FIELDS_COSTOS) + 1, column=0, columnspan=2,
                 padx=(18, 8), pady=(4, 14), sticky="w")

        self.mantener_mp_var = ctk.BooleanVar(value=False)
        self.mantener_flete_var = ctk.BooleanVar(value=False)

        self.mantener_mp_cb = ctk.CTkCheckBox(
            row,
            text="Mantener MP",
            variable=self.mantener_mp_var,
            font=CTkFont("Segoe UI", 12, "bold"),
            text_color=CEMEX_BLUE,
            checkbox_width=18,
            checkbox_height=18,
            fg_color=CEMEX_BLUE,
            hover_color=CEMEX_BLUE_DARK,
            border_color=CEMEX_LIGHT_GRAY,
        )
        self.mantener_mp_cb.grid(row=0, column=0, padx=(0, 18), pady=2, sticky="w")

        self.mantener_flete_cb = ctk.CTkCheckBox(
            row,
            text="Mantener Flete",
            variable=self.mantener_flete_var,
            font=CTkFont("Segoe UI", 12, "bold"),
            text_color=CEMEX_BLUE,
            checkbox_width=18,
            checkbox_height=18,
            fg_color=CEMEX_BLUE,
            hover_color=CEMEX_BLUE_DARK,
            border_color=CEMEX_LIGHT_GRAY,
        )
        self.mantener_flete_cb.grid(row=0, column=1, padx=(0, 18), pady=2, sticky="w")

    def _build_actions_section(self, parent):
        parent.grid_columnconfigure((0, 1), weight=1)

        # Botón Calcular
        self.calc_button = ctk.CTkButton(
            parent,
            text="Calcular Fila (Enter)",
            command=self.calculate_row,
            fg_color=CEMEX_BLUE,
            hover_color=CEMEX_BLUE_DARK,
            font=CTkFont("Segoe UI", 13, "bold"),
            height=38,
        )
        self.calc_button.grid(row=1, column=0, padx=(18, 8), pady=(10, 8), sticky="ew")

        # Contador de lote
        self.counter_label = ctk.CTkLabel(
            parent,
            text="Filas en la solicitud actual: 0",
            font=CTkFont("Segoe UI", 14, "bold"),
            text_color=CEMEX_BLUE,
            anchor="w",
        )
        self.counter_label.grid(row=1, column=1, padx=18, pady=(10, 8), sticky="w")

        # Vista previa
        self.preview_text = ctk.CTkTextbox(
            parent,
            state="disabled",
            wrap="none",
            font=CTkFont("Consolas", 11),
            height=180,
            fg_color="#FAFAFA",
            text_color=CEMEX_TEXT,
            border_width=1,
            border_color=CEMEX_LIGHT_GRAY,
        )
        self.preview_text.grid(row=4, column=0, columnspan=2, padx=18, pady=8, sticky="nsew")

        # Botón Agregar a Lote
        self.add_batch_button = ctk.CTkButton(
            parent,
            text="Agregar Fila a Solicitud Actual",
            command=self.agregar_fila_a_lote,
            fg_color=CEMEX_BLUE,
            hover_color=CEMEX_BLUE_DARK,
            font=CTkFont("Segoe UI", 13, "bold"),
            height=38,
            state="disabled",
        )
        self.add_batch_button.grid(row=2, column=0, padx=(18, 8), pady=8, sticky="ew")

        # Botón Generar Archivos de Solicitud (CTA principal)
        self.generate_button = ctk.CTkButton(
            parent,
            text="Generar Archivos de Solicitud",
            command=self.generar_archivos_solicitud,
            fg_color=CEMEX_RED,
            hover_color=CEMEX_RED_DARK,
            font=CTkFont("Segoe UI", 13, "bold"),
            height=38,
        )
        self.generate_button.grid(row=2, column=1, padx=(8, 18), pady=8, sticky="ew")

        # Botón ETL histórico
        self.etl_button = ctk.CTkButton(
            parent,
            text="Procesar ETL Histórico",
            command=self.run_etl,
            fg_color=CEMEX_NEUTRAL,
            hover_color=CEMEX_NEUTRAL_DARK,
            font=CTkFont("Segoe UI", 12),
            height=32,
        )
        self.etl_button.grid(row=3, column=0, padx=(18, 8), pady=8, sticky="w")

    def _build_pending_tree(self, parent):
        parent.grid_columnconfigure(0, weight=1)
        parent.grid_rowconfigure(2, weight=1)

        ctk.CTkLabel(
            parent,
            text="Haz clic en una fila para cargarla en el formulario y capturar costos.",
            font=CTkFont("Segoe UI", 11),
            text_color=CEMEX_TEXT,
        ).grid(row=1, column=0, padx=18, pady=(6, 4), sticky="w")

        frame = ctk.CTkFrame(parent, fg_color=CEMEX_WHITE, corner_radius=0)
        frame.grid(row=2, column=0, padx=18, pady=(0, 14), sticky="nsew")
        frame.grid_columnconfigure(0, weight=1)
        frame.grid_rowconfigure(0, weight=1)

        self.pending_tree = ttk.Treeview(
            frame,
            columns=PENDING_COLUMNS,
            show="headings",
            height=6,
            style="Cemex.Treeview",
        )
        self.pending_tree.tag_configure("even", background=CEMEX_WHITE, foreground=CEMEX_TEXT)
        self.pending_tree.tag_configure("odd", background=CEMEX_GRAY_BG, foreground=CEMEX_TEXT)
        self.pending_tree.grid(row=0, column=0, sticky="nsew")

        vsb = ttk.Scrollbar(frame, orient="vertical", command=self.pending_tree.yview)
        vsb.grid(row=0, column=1, sticky="ns")
        hsb = ttk.Scrollbar(frame, orient="horizontal", command=self.pending_tree.xview)
        hsb.grid(row=1, column=0, sticky="ew")
        self.pending_tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)

        for col in PENDING_COLUMNS:
            self.pending_tree.heading(col, text=col)
            self.pending_tree.column(col, width=110, anchor="center")

        self.pending_tree.bind("<<TreeviewSelect>>", self._on_pending_select)
        self._bind_copy_excel(self.pending_tree)

    # ------------------------------------------------------------------
    # Pestaña: Panel de Validaciones y SLA
    # ------------------------------------------------------------------
    def _build_tab_validaciones(self):
        tab = self.validaciones_view
        tab.grid_columnconfigure(0, weight=1)
        tab.grid_rowconfigure(0, weight=1)

        scroll = ctk.CTkScrollableFrame(tab, fg_color=CEMEX_GRAY_BG)
        scroll.grid(row=0, column=0, sticky="nsew")
        scroll.grid_columnconfigure(0, weight=1)

        # Card de búsqueda rápida (pegar fila A-R o editar manualmente)
        search_card = self._create_card(scroll, "Búsqueda rápida de solicitudes (pegar fila A-R)")
        search_card.grid(row=0, column=0, padx=15, pady=(15, 8), sticky="nsew")
        search_card.grid_columnconfigure(0, weight=1)
        self._build_sla_search(search_card)

        card = self._create_card(scroll, "Solicitudes en Base de Datos y Estado SLA")
        card.grid(row=1, column=0, padx=15, pady=(0, 15), sticky="nsew")
        card.grid_columnconfigure(0, weight=1)

        # Ruta absoluta de BD
        ctk.CTkLabel(
            card,
            text=f"BD: {DB_FILE}",
            font=CTkFont("Segoe UI", 10),
            text_color=CEMEX_TEXT,
        ).grid(row=1, column=0, padx=18, pady=(6, 0), sticky="w")

        # Treeview SLA
        frame = ctk.CTkFrame(card, fg_color=CEMEX_WHITE, corner_radius=0)
        frame.grid(row=2, column=0, padx=18, pady=(8, 12), sticky="nsew")
        frame.grid_columnconfigure(0, weight=1)
        frame.grid_rowconfigure(0, weight=1)

        self.sla_tree = ttk.Treeview(
            frame,
            columns=SLA_COLUMNS,
            show="headings",
            height=14,
            style="Cemex.Treeview",
        )
        self.sla_tree.tag_configure("even", background=CEMEX_WHITE, foreground=CEMEX_TEXT)
        self.sla_tree.tag_configure("odd", background=CEMEX_GRAY_BG, foreground=CEMEX_TEXT)
        self.sla_tree.grid(row=0, column=0, sticky="nsew")

        vsb = ttk.Scrollbar(frame, orient="vertical", command=self.sla_tree.yview)
        vsb.grid(row=0, column=1, sticky="ns")
        hsb = ttk.Scrollbar(frame, orient="horizontal", command=self.sla_tree.xview)
        hsb.grid(row=1, column=0, sticky="ew")
        self.sla_tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)

        for col in SLA_COLUMNS:
            self.sla_tree.heading(col, text=col)
            self.sla_tree.column(col, width=110, anchor="center")
        self.sla_tree.column("ID", width=50, anchor="center")

        self.sla_tree.bind("<<TreeviewSelect>>", self._on_sla_select)
        self._bind_copy_excel(self.sla_tree)

        # Panel de edición de fechas (fila 1: campos, fila 2: botones)
        edit_frame = ctk.CTkFrame(card, fg_color="transparent")
        edit_frame.grid(row=3, column=0, padx=18, pady=(0, 12), sticky="w")

        ctk.CTkLabel(
            edit_frame,
            text="Fecha Solicitud:",
            font=CTkFont("Segoe UI", 12),
            text_color=CEMEX_TEXT,
        ).grid(row=0, column=0, padx=(0, 8), pady=5)
        self.sla_fecha_solicitud = ctk.CTkEntry(
            edit_frame,
            font=CTkFont("Segoe UI", 12),
            width=150,
            fg_color=CEMEX_WHITE,
            border_color=CEMEX_LIGHT_GRAY,
        )
        self.sla_fecha_solicitud.grid(row=0, column=1, padx=(0, 18), pady=5)

        ctk.CTkLabel(
            edit_frame,
            text="Fecha Confirmación:",
            font=CTkFont("Segoe UI", 12),
            text_color=CEMEX_TEXT,
        ).grid(row=0, column=2, padx=(0, 8), pady=5)
        self.sla_fecha_conf = ctk.CTkEntry(
            edit_frame,
            font=CTkFont("Segoe UI", 12),
            width=150,
            fg_color=CEMEX_WHITE,
            border_color=CEMEX_LIGHT_GRAY,
        )
        self.sla_fecha_conf.grid(row=0, column=3, padx=(0, 18), pady=5)

        self.save_dates_button = ctk.CTkButton(
            edit_frame,
            text="Guardar Fechas Seleccionadas",
            command=self._guardar_fechas_sla,
            fg_color=CEMEX_BLUE,
            hover_color=CEMEX_BLUE_DARK,
            font=CTkFont("Segoe UI", 12, "bold"),
            height=32,
        )
        self.save_dates_button.grid(row=0, column=4, padx=(0, 12), pady=5)

        self.refresh_button = ctk.CTkButton(
            edit_frame,
            text="Refrescar",
            command=self._refresh_sla_tree,
            fg_color=CEMEX_NEUTRAL,
            hover_color=CEMEX_NEUTRAL_DARK,
            font=CTkFont("Segoe UI", 12),
            height=32,
        )
        self.refresh_button.grid(row=0, column=5, padx=(0, 12), pady=5)

        # Fila 2 de acciones (para que ningún botón quede cortado).
        action_frame = ctk.CTkFrame(card, fg_color="transparent")
        action_frame.grid(row=4, column=0, padx=18, pady=(0, 14), sticky="w")

        self.sla_cellgrid_btn = ctk.CTkButton(
            action_frame,
            text="Selección por celdas",
            command=lambda: self._abrir_cell_grid(self.sla_tree, "SLA — Selección por celdas"),
            fg_color=CEMEX_BLUE,
            hover_color=CEMEX_BLUE_DARK,
            font=CTkFont("Segoe UI", 12, "bold"),
            height=32,
        )
        self.sla_cellgrid_btn.grid(row=0, column=0, padx=(0, 12), pady=5)

        self.generate_sla_button = ctk.CTkButton(
            action_frame,
            text="Generar Archivos desde BD",
            command=self._generar_archivos_desde_bd,
            fg_color=CEMEX_RED,
            hover_color=CEMEX_RED_DARK,
            font=CTkFont("Segoe UI", 12, "bold"),
            height=32,
        )
        self.generate_sla_button.grid(row=0, column=1, padx=(0, 12), pady=5)

        self.select_excel_button = ctk.CTkButton(
            action_frame,
            text="Seleccionar Excel Maestro",
            command=self._seleccionar_excel_maestro,
            fg_color=CEMEX_NEUTRAL,
            hover_color=CEMEX_NEUTRAL_DARK,
            font=CTkFont("Segoe UI", 12),
            height=32,
        )
        self.select_excel_button.grid(row=0, column=2, padx=(0, 12), pady=5)

        # Etiqueta con la ruta actual del Excel Maestro (fila propia, ancha).
        self.excel_path_label = ctk.CTkLabel(
            action_frame,
            text=f"Excel: {etl_processor.get_excel_maestro_path()}",
            font=CTkFont("Segoe UI", 10),
            text_color=CEMEX_NEUTRAL,
            anchor="w",
        )
        self.excel_path_label.grid(row=1, column=0, columnspan=3, padx=(0, 12), pady=(8, 0), sticky="w")

    def _habilitar_drag_drop(self):
        """Registra Drag & Drop de archivos en la ventana y en el área de
        validaciones."""
        try:
            # Ya heredamos de DnDWrapper, así que registramos el drop
            # directamente sobre `self` sin instanciar nada manualmente.
            self.drop_target_register("DND_Files")
            self.dnd_bind("<<Drop>>", self._on_drop_excel)
        except Exception as e:
            print(f"AVISO: no se pudo habilitar Drag & Drop: {e}")

    def _on_drop_excel(self, event):
        """Maneja el evento <Drop> de archivos .xlsx/.xlsm.

        Lee la ruta del archivo, extrae las hojas FPA/CTO con pandas y muestra
        el resultado limpio en el Treeview SLA (o CellGrid si está abierto).
        """
        try:
            data = event.data
            if not data:
                return
            # tkinterdnd entrega la ruta entre llaves en Windows.
            ruta = data.strip()
            if ruta.startswith("{") and ruta.endswith("}"):
                ruta = ruta[1:-1]
            ruta = os.path.normpath(ruta)
            if not os.path.exists(ruta):
                messagebox.showerror("Drag & Drop", f"No se encontró el archivo:\n{ruta}")
                return
            if not ruta.lower().endswith((".xlsx", ".xlsm")):
                messagebox.showwarning("Drag & Drop", "Solo se aceptan archivos .xlsx o .xlsm.")
                return

            # Si estamos en la pestaña de Validación SAP GUI, procesar ahí
            if getattr(self, "active_view", None) == "val_sap":
                self._cargar_archivo_val_sap(ruta)
                return

            # Si el archivo parece un formato FPA/CTO y estamos en otra pestaña, preguntar
            nombre_base = os.path.basename(ruta).lower()
            if any(k in nombre_base for k in ["fpa", "filial", "terceros", "cto", "compra", "venta", "dw"]):
                if messagebox.askyesno(
                    "Validación SAP GUI",
                    f"Se detectó el archivo:\n{os.path.basename(ruta)}\n\n¿Deseas abrirlo en 'Validación SAP GUI' para auditarlo en vivo con SAP?"
                ):
                    self._switch_view("val_sap")
                    self._cargar_archivo_val_sap(ruta)
                    return

            df = self._extraer_hoja_excel(ruta)
            if df is None or df.empty:
                messagebox.showinfo("Drag & Drop", "No se encontró una hoja FPA/CTO con datos.")
                return
            self._mostrar_dataframe(df)
        except Exception as e:
            messagebox.showerror("Drag & Drop", f"Error al procesar el archivo:\n{e}")

    def _extraer_hoja_excel(self, ruta):
        """Lee todas las hojas y busca las variaciones de FPA/CTO.

        Retorna el DataFrame limpio de la primera hoja que coincida, o None.
        """
        variaciones = (
            "fpa filial", "fpa terceros", "cto venta", "cto compra",
        )
        xls = pd.read_excel(ruta, sheet_name=None, header=None)
        for nombre_hoja, df in xls.items():
            if not any(v in str(nombre_hoja).lower() for v in variaciones):
                continue
            # Localizar la fila de encabezados: la primera fila con texto en
            # al menos una celda se usa como header.
            header_row = None
            for idx, row in df.iterrows():
                if row.notna().any():
                    header_row = idx
                    break
            if header_row is None:
                continue
            # Releer esa hoja usando la fila detectada como encabezado.
            df_limpio = pd.read_excel(ruta, sheet_name=nombre_hoja, header=header_row)
            df_limpio = df_limpio.dropna(how="all").fillna("")
            if not df_limpio.empty:
                return df_limpio
        return None

    def _mostrar_dataframe(self, df):
        """Muestra un DataFrame en el Treeview SLA de forma dinámica."""
        # Limpiar el árbol y reconfigurar columnas según el DataFrame.
        for item in self.sla_tree.get_children():
            self.sla_tree.delete(item)

        cols = [str(c) for c in df.columns]
        self.sla_tree["columns"] = cols
        for col in cols:
            self.sla_tree.heading(col, text=col)
            self.sla_tree.column(col, width=120, anchor="center", stretch=True)

        filas = df.values.tolist()
        for i, fila in enumerate(filas):
            tag = "even" if i % 2 == 0 else "odd"
            self.sla_tree.insert(
                "", "end", iid=str(i),
                values=tuple("" if pd.isna(v) else v for v in fila),
                tags=(tag,),
            )

        self.ajustar_columnas_inteligente(self.sla_tree, filas)
        self._fpa_show_message(f"Archivo cargado: {len(df)} fila(s), {len(cols)} columna(s).")

    def _build_sla_search(self, parent):
        """Construye la zona de búsqueda rápida en Validaciones y SLA.

        Permite pegar una fila del Excel (columnas A-R) o editar manualmente
        los campos de la clave: CEDIS, SF, Planta/Obra, Artículo, Modalidad.
        """
        parent.grid_columnconfigure(0, weight=1)

        # Zona de pegado
        ctk.CTkLabel(
            parent,
            text="Pega aquí una fila copiada de Excel (columnas A-R) y presiona Ctrl+V.",
            font=CTkFont("Segoe UI", 11),
            text_color=CEMEX_TEXT,
        ).grid(row=0, column=0, padx=18, pady=(6, 4), sticky="w")

        self.sla_paste_box = ctk.CTkTextbox(
            parent,
            height=50,
            wrap="none",
            font=CTkFont("Consolas", 11),
            fg_color="#FFF8E1",
            text_color=CEMEX_TEXT,
            border_width=1,
            border_color=CEMEX_LIGHT_GRAY,
        )
        self.sla_paste_box.grid(row=1, column=0, padx=18, pady=(0, 6), sticky="ew")
        self.sla_paste_box.insert("1.0", "Zona de pegado (Ctrl+V)")
        self.sla_paste_box.bind("<Control-v>", self._sla_paste_search)
        self.sla_paste_box.bind("<Button-1>", lambda e: self._sla_clear_paste_hint())

        # Campos editables de la clave
        fields = [
            ("CEDIS", "cedis"),
            ("SF", "sf"),
            ("Planta/Obra", "planta_obra"),
            ("Artículo", "articulo"),
            ("Modalidad", "modalidad"),
        ]
        for i, (label, key) in enumerate(fields):
            col = i * 2
            ctk.CTkLabel(
                parent, text=label, font=CTkFont("Segoe UI", 12, "bold"),
                text_color=CEMEX_TEXT, anchor="w",
            ).grid(row=2, column=col, padx=(18, 4), pady=(4, 4), sticky="w")

            var = StringVar(value="")
            self.sla_search_vars[key] = var
            entry = ctk.CTkEntry(
                parent, textvariable=var, font=CTkFont("Segoe UI", 12),
                width=120, fg_color=CEMEX_WHITE,
                border_color=CEMEX_LIGHT_GRAY, text_color=CEMEX_TEXT,
            )
            entry.grid(row=2, column=col + 1, padx=(0, 8), pady=(4, 4), sticky="w")
            entry.bind("<KeyRelease>", self._sla_filter_live)

        # Botones: limpiar y buscar
        btns = ctk.CTkFrame(parent, fg_color="transparent")
        btns.grid(row=3, column=0, padx=18, pady=(4, 8), sticky="ew")
        btns.grid_columnconfigure(0, weight=1)

        ctk.CTkButton(
            btns, text="Limpiar búsqueda", command=self._sla_clear_search,
            fg_color=CEMEX_NEUTRAL, hover_color=CEMEX_NEUTRAL_DARK,
            font=CTkFont("Segoe UI", 12, "bold"), height=32,
        ).grid(row=0, column=0, padx=(0, 8), pady=4, sticky="ew")

        ctk.CTkButton(
            btns, text="Buscar ahora", command=self._sla_filter_live,
            fg_color=CEMEX_BLUE, hover_color=CEMEX_BLUE_DARK,
            font=CTkFont("Segoe UI", 12, "bold"), height=32,
        ).grid(row=0, column=1, padx=(8, 0), pady=4, sticky="ew")

        ctk.CTkButton(
            btns, text="Búsqueda Avanzada", command=self._abrir_busqueda_avanzada,
            fg_color=CEMEX_RED, hover_color=CEMEX_RED_DARK,
            font=CTkFont("Segoe UI", 12, "bold"), height=32,
        ).grid(row=0, column=2, padx=(8, 0), pady=4, sticky="ew")

    def _sla_clear_paste_hint(self):
        try:
            current = self.sla_paste_box.get("1.0", "end").strip()
            if current == "Zona de pegado (Ctrl+V)":
                self.sla_paste_box.delete("1.0", "end")
        except Exception:
            pass

    def _sla_paste_search(self, event=None):
        """Lee una fila pegada y llena los campos de la clave según columnas A-R."""
        try:
            raw = self.sla_paste_box.get("1.0", "end")
            if raw.strip() == "Zona de pegado (Ctrl+V)":
                raw = self.clipboard_get()
        except Exception:
            raw = self.clipboard_get()

        line = raw.splitlines()[0] if raw.splitlines() else ""
        values = [v.strip() for v in line.split("\t")]

        # Columnas del Excel maestro (1-based) para la clave:
        #   C=3 CEDIS, E=5 SF, G=7 NO PLANTA/OBRA, I=9 ARTICULO, L=12 MODALIDAD
        mapping = {"cedis": 3, "sf": 5, "planta_obra": 7, "articulo": 9, "modalidad": 12}

        for key, col in mapping.items():
            if col - 1 < len(values):
                self.sla_search_vars[key].set(values[col - 1])

        # Reflejar el texto en la caja de pegado
        try:
            self.sla_paste_box.delete("1.0", "end")
            self.sla_paste_box.insert("1.0", line)
        except Exception:
            pass

        self._sla_filter_live()
        return "break"

    def _sla_construir_where(self, criterios):
        """Construye (where_clause, params) tolerante para la búsqueda SLA.

        Normaliza modalidad ('1' <-> '01'), IDs numéricos ('900006225.0' ->
        '900006225') y fechas (separadores . - /), de modo que la búsqueda no
        falle por diferencias de formato al pegar desde Excel.
        """
        clauses = []
        params = []

        for key, val in criterios.items():
            if not val:
                continue
            if key == "modalidad":
                canon = fpa_validations._norm_modalidad_db(val)
                corto = fpa_validations._norm_modalidad(val)
                if canon and corto and canon != corto:
                    clauses.append('(modalidad = ? OR modalidad = ?)')
                    params.extend([canon, corto])
                elif canon:
                    clauses.append('modalidad = ?')
                    params.append(canon)
                else:
                    clauses.append('modalidad = ?')
                    params.append(val)
            elif key in ("cedis", "sf", "planta_obra", "articulo", "org_vta"):
                clauses.append(f'"{key}" = ?')
                params.append(fpa_validations._norm_id(val))
            elif key in ("fecha_solicitud", "fecha_conf", "fecha_procesamiento"):
                fecha = normalizar_fecha(val)
                if not fecha:
                    clauses.append(f'"{key}" = ?')
                    params.append(val)
                else:
                    clauses.append(
                        f'REPLACE(REPLACE("{key}", \'-\', \'/\'), \'.\', \'/\') = ?'
                    )
                    params.append(fecha)
            else:
                clauses.append(f'"{key}" = ?')
                params.append(val)

        return clauses, params

    def _sla_filter_live(self, event=None):
        """Filtra la tabla SLA en vivo según los campos de búsqueda."""
        if not hasattr(self, "sla_tree") or not self.conn:
            return
        criterios = {k: v.get().strip() for k, v in self.sla_search_vars.items() if v.get().strip()}

        if not criterios:
            self._refresh_sla_tree()
            return

        clauses, params = self._sla_construir_where(criterios)
        if not clauses:
            self._refresh_sla_tree()
            return

        where = " AND ".join(clauses)
        sql = (
            "SELECT id, org_vta, cedis, sf, nombre_sf, planta_obra, nombre_planta, "
            "articulo, modalidad, fecha_solicitud, fecha_procesamiento, fecha_conf "
            f"FROM staging_rutas WHERE {where} ORDER BY id DESC"
        )
        cur = self.conn.cursor()
        cur.execute(sql, params)
        rows = cur.fetchall()

        self._sla_render_rows(rows)

    def _sla_render_rows(self, rows):
        """Puebla el Treeview SLA con las filas crudas devueltas por una consulta.

        `rows` es una lista de tuplas en el orden:
        (id, org_vta, cedis, sf, nombre_sf, planta_obra, nombre_planta,
         articulo, modalidad, fecha_solicitud, fecha_procesamiento, fecha_conf).
        """
        for item in self.sla_tree.get_children():
            self.sla_tree.delete(item)

        for idx, row in enumerate(rows):
            (row_id, org_vta, cedis, sf, nombre_sf, planta, nombre_planta,
             articulo, modalidad, fecha_solicitud, fecha_procesamiento, fecha_conf) = row
            ruta = "-".join(str(x) for x in [cedis, sf, planta, articulo, modalidad] if x)
            dias_habiles, dias_transcurridos, estado = etl_processor.calcular_estado_sla(
                fecha_procesamiento, fecha_conf
            )
            dias_habiles_str = str(dias_habiles) if dias_habiles is not None else ""
            dias_trans_str = str(dias_transcurridos) if dias_transcurridos is not None else ""
            tag = "even" if idx % 2 == 0 else "odd"
            self.sla_tree.insert(
                "", "end", iid=str(row_id),
                values=(row_id, ruta, org_vta, cedis, sf, planta, articulo, modalidad,
                        fecha_solicitud or "", fecha_procesamiento or "", fecha_conf or "",
                        dias_habiles_str, dias_trans_str, estado),
                tags=(tag,)
            )

    def _abrir_busqueda_avanzada(self):
        """Abre un popup independiente de Búsqueda Avanzada con campos adicionales.

        Permite filtrar por CEDIS, SF, Planta/Obra, Artículo, Modalidad,
        Fecha Solicitud y Fecha Confirmación sin saturar la pestaña principal.
        """
        if not self.conn:
            messagebox.showerror("Búsqueda Avanzada", "No hay conexión a SQLite.")
            return

        top = ctk.CTkToplevel(self)
        top.title("Búsqueda Avanzada")
        top.geometry("520x430")
        top.configure(fg_color=CEMEX_GRAY_BG)
        top.grab_set()

        campos = [
            ("CEDIS", "cedis"),
            ("SF", "sf"),
            ("Planta/Obra", "planta_obra"),
            ("Artículo", "articulo"),
            ("Modalidad", "modalidad"),
            ("Fecha Solicitud", "fecha_solicitud"),
            ("Fecha Confirmación", "fecha_conf"),
        ]
        vars_ = {}
        for i, (label, key) in enumerate(campos):
            ctk.CTkLabel(
                top, text=label, font=CTkFont("Segoe UI", 12, "bold"),
                text_color=CEMEX_TEXT, anchor="w",
            ).grid(row=i, column=0, padx=(20, 8), pady=6, sticky="w")
            var = StringVar(value="")
            vars_[key] = var
            ctk.CTkEntry(
                top, textvariable=var, font=CTkFont("Segoe UI", 12),
                width=220, fg_color=CEMEX_WHITE,
                border_color=CEMEX_LIGHT_GRAY, text_color=CEMEX_TEXT,
            ).grid(row=i, column=1, padx=(0, 20), pady=6, sticky="w")

        def _buscar():
            criterios = {k: v.get().strip() for k, v in vars_.items() if v.get().strip()}
            if not criterios:
                messagebox.showwarning("Búsqueda Avanzada", "Escribe al menos un criterio.", parent=top)
                return

            clauses, params = self._sla_construir_where(criterios)
            if not clauses:
                messagebox.showwarning("Búsqueda Avanzada", "Escribe al menos un criterio.", parent=top)
                return

            where = " AND ".join(clauses)
            sql = (
                "SELECT id, org_vta, cedis, sf, nombre_sf, planta_obra, nombre_planta, "
                "articulo, modalidad, fecha_solicitud, fecha_procesamiento, fecha_conf "
                f"FROM staging_rutas WHERE {where} ORDER BY id DESC"
            )
            try:
                cur = self.conn.cursor()
                cur.execute(sql, params)
                rows = cur.fetchall()
            except Exception as e:
                messagebox.showerror("Búsqueda Avanzada", f"Error:\n{e}", parent=top)
                return

            self._sla_render_rows(rows)
            top.destroy()

        btns = ctk.CTkFrame(top, fg_color="transparent")
        btns.grid(row=len(campos), column=0, columnspan=2, padx=20, pady=(12, 16), sticky="ew")
        btns.grid_columnconfigure(0, weight=1)

        ctk.CTkButton(
            btns, text="Buscar", command=_buscar,
            fg_color=CEMEX_BLUE, hover_color=CEMEX_BLUE_DARK,
            font=CTkFont("Segoe UI", 12, "bold"), height=36,
        ).grid(row=0, column=0, padx=(0, 6), sticky="ew")

        ctk.CTkButton(
            btns, text="Cerrar", command=top.destroy,
            fg_color=CEMEX_NEUTRAL, hover_color=CEMEX_NEUTRAL_DARK,
            font=CTkFont("Segoe UI", 12, "bold"), height=36,
        ).grid(row=0, column=1, padx=(6, 0), sticky="ew")

    def _sla_clear_search(self):
        for var in self.sla_search_vars.values():
            var.set("")
        try:
            self.sla_paste_box.delete("1.0", "end")
            self.sla_paste_box.insert("1.0", "Zona de pegado (Ctrl+V)")
        except Exception:
            pass
        self._refresh_sla_tree()

    def _options_for(self, key):
        opts = {
            "modalidad": ["01", "02", "04"],
            "um_vta": ["M3", "TN"],
            "tipo": ["ALTA", "CAMBIO", "BAJA"],
            "tipo_cliente": ["FILIAL", "TERCEROS"],
            "incidencia": ["NO", "SI"],
        }
        return opts.get(key, [])

    # ------------------------------------------------------------------
    # Pestaña: Área de Validaciones (FPA)
    # ------------------------------------------------------------------
    def _build_tab_fpa(self):
        tab = self.fpa_view
        tab.grid_columnconfigure(0, weight=1)
        tab.grid_rowconfigure(0, weight=1)

        scroll = ctk.CTkScrollableFrame(tab, fg_color=CEMEX_GRAY_BG)
        scroll.grid(row=0, column=0, sticky="nsew")
        scroll.grid_columnconfigure(0, weight=1)

        # Card: Pegado masivo / identificación FPA
        card_paste = self._create_card(scroll, "Área de Pegado Masivo (Identificación FPA)")
        card_paste.grid(row=0, column=0, padx=15, pady=15, sticky="nsew")
        card_paste.grid_columnconfigure(0, weight=1)

        ctk.CTkLabel(
            card_paste,
            text="1) Haz clic en la ZONA DE PEGADO y presiona Ctrl+V para pegar las 24 columnas desde Excel.",
            font=CTkFont("Segoe UI", 12, "bold"),
            text_color=CEMEX_RED,
        ).grid(row=1, column=0, padx=18, pady=(8, 4), sticky="w")

        ctk.CTkLabel(
            card_paste,
            text="2) El resultado aparece en la tabla inferior. Selecciona una fila y usa F6/F7/F8/F9.",
            font=CTkFont("Segoe UI", 11),
            text_color=CEMEX_TEXT,
        ).grid(row=2, column=0, padx=18, pady=(0, 6), sticky="w")

        # Filtro opcional: Fecha de Procesamiento (columna AF / fecha_procesamiento)
        filtro_frame = ctk.CTkFrame(card_paste, fg_color="transparent")
        filtro_frame.grid(row=3, column=0, padx=18, pady=(0, 8), sticky="ew")
        filtro_frame.grid_columnconfigure(1, weight=1)

        ctk.CTkLabel(
            filtro_frame,
            text="Fecha de Procesamiento (opcional, AF):",
            font=CTkFont("Segoe UI", 12, "bold"),
            text_color=CEMEX_TEXT,
        ).grid(row=0, column=0, padx=(0, 8), sticky="w")
        self.fpa_fecha_proc_var = StringVar(value="")
        self.fpa_fecha_proc_entry = ctk.CTkEntry(
            filtro_frame,
            textvariable=self.fpa_fecha_proc_var,
            font=CTkFont("Segoe UI", 12),
            width=160,
            placeholder_text="dd/mm/aaaa",
            fg_color=CEMEX_WHITE,
            border_color=CEMEX_LIGHT_GRAY,
            text_color=CEMEX_TEXT,
        )
        self.fpa_fecha_proc_entry.grid(row=0, column=1, padx=(0, 8), sticky="w")

        self.fpa_migrar_btn = ctk.CTkButton(
            filtro_frame,
            text="Migrar Excel → SQLite",
            command=self.fpa_migrar_excel,
            fg_color=CEMEX_NEUTRAL,
            hover_color=CEMEX_NEUTRAL_DARK,
            font=CTkFont("Segoe UI", 12, "bold"),
            height=32,
        )
        self.fpa_migrar_btn.grid(row=0, column=2, padx=(8, 0), sticky="e")

        # Zona de pegado dedicada y claramente visible
        self.fpa_paste_box = ctk.CTkTextbox(
            card_paste,
            height=80,
            wrap="none",
            font=CTkFont("Consolas", 11),
            fg_color="#FFF8E1",
            text_color=CEMEX_TEXT,
            border_width=2,
            border_color=CEMEX_RED,
        )
        self.fpa_paste_box.grid(row=4, column=0, padx=18, pady=(0, 10), sticky="ew")
        self.fpa_paste_box.insert("1.0", "ZONA DE PEGADO: haz clic aquí y presiona Ctrl+V")
        self.fpa_paste_box.bind("<Control-v>", self._fpa_paste)
        self.fpa_paste_box.bind("<Button-1>", self._on_fpa_paste_box_click)

        frame_paste = ctk.CTkFrame(card_paste, fg_color=CEMEX_WHITE, corner_radius=0)
        frame_paste.grid(row=5, column=0, padx=18, pady=(0, 12), sticky="nsew")
        frame_paste.grid_columnconfigure(0, weight=1)
        frame_paste.grid_rowconfigure(0, weight=1)

        self.fpa_tree = ttk.Treeview(
            frame_paste,
            columns=fpa_validations.FPA_COLUMNS,
            show="headings",
            height=6,
            style="Cemex.Treeview",
        )
        self.fpa_tree.tag_configure("even", background=CEMEX_WHITE, foreground=CEMEX_TEXT)
        self.fpa_tree.tag_configure("odd", background=CEMEX_GRAY_BG, foreground=CEMEX_TEXT)
        self.fpa_tree.grid(row=0, column=0, sticky="nsew")

        vsb = ttk.Scrollbar(frame_paste, orient="vertical", command=self.fpa_tree.yview)
        vsb.grid(row=0, column=1, sticky="ns")
        hsb = ttk.Scrollbar(frame_paste, orient="horizontal", command=self.fpa_tree.xview)
        hsb.grid(row=1, column=0, sticky="ew")
        self.fpa_tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)

        for col in fpa_validations.FPA_COLUMNS:
            self.fpa_tree.heading(col, text=col)
            self.fpa_tree.column(col, width=120, anchor="center")

        self.fpa_tree.bind("<Control-v>", self._fpa_paste)
        self.fpa_tree.bind("<<TreeviewSelect>>", self._on_fpa_select)
        self.fpa_tree.bind("<Delete>", self._fpa_delete_selected)
        self._bind_copy_excel(self.fpa_tree)

        # El resultado de búsqueda/proyección se coloca en una segunda card (más abajo)

        # Botones F6/F8
        btns = ctk.CTkFrame(card_paste, fg_color="transparent")
        btns.grid(row=6, column=0, padx=18, pady=(0, 14), sticky="w")
        btns.grid_columnconfigure((0, 1), weight=1)

        self.fpa_buscar_btn = ctk.CTkButton(
            btns, text="Buscar fila (F6)",
            command=self.fpa_buscar,
            fg_color=CEMEX_BLUE, hover_color=CEMEX_BLUE_DARK,
            font=CTkFont("Segoe UI", 12, "bold"), height=36,
        )
        self.fpa_buscar_btn.grid(row=0, column=0, padx=(0, 8), pady=6, sticky="ew")

        self.fpa_validar_btn = ctk.CTkButton(
            btns, text="Validar fila(s) (F8)",
            command=self.fpa_validar,
            fg_color=CEMEX_BLUE, hover_color=CEMEX_BLUE_DARK,
            font=CTkFont("Segoe UI", 12, "bold"), height=36,
        )
        self.fpa_validar_btn.grid(row=0, column=1, padx=(8, 0), pady=6, sticky="ew")

        # Botones de procesamiento masivo (toda la solicitud pegada)
        self.fpa_buscar_solicitud_btn = ctk.CTkButton(
            btns, text="Buscar solicitud",
            command=self.fpa_buscar_solicitud,
            fg_color=CEMEX_BLUE, hover_color=CEMEX_BLUE_DARK,
            font=CTkFont("Segoe UI", 12, "bold"), height=36,
        )
        self.fpa_buscar_solicitud_btn.grid(row=1, column=0, padx=(0, 8), pady=6, sticky="ew")

        self.fpa_validar_solicitud_btn = ctk.CTkButton(
            btns, text="Validar solicitud",
            command=self.fpa_validar_solicitud,
            fg_color=CEMEX_RED, hover_color=CEMEX_RED_DARK,
            font=CTkFont("Segoe UI", 12, "bold"), height=36,
        )
        self.fpa_validar_solicitud_btn.grid(row=1, column=1, padx=(8, 0), pady=6, sticky="ew")

        self.fpa_cellgrid_btn = ctk.CTkButton(
            btns, text="Selección por celdas",
            command=lambda: self._abrir_cell_grid(self.fpa_tree, "FPA — Selección por celdas"),
            fg_color=CEMEX_NEUTRAL, hover_color=CEMEX_NEUTRAL_DARK,
            font=CTkFont("Segoe UI", 12, "bold"), height=36,
        )
        self.fpa_cellgrid_btn.grid(row=2, column=0, columnspan=2, padx=(0, 8), pady=6, sticky="ew")

        # Card: resultados de búsqueda
        card_result = self._create_card(scroll, "Resultados de Búsqueda")
        card_result.grid(row=1, column=0, padx=15, pady=(0, 15), sticky="nsew")
        card_result.grid_columnconfigure(0, weight=1)

        self.fpa_result_text = ctk.CTkTextbox(
            card_result,
            state="disabled",
            wrap="none",
            font=CTkFont("Consolas", 11),
            height=140,
            fg_color="#FAFAFA",
            text_color=CEMEX_TEXT,
            border_width=1,
            border_color=CEMEX_LIGHT_GRAY,
        )
        self.fpa_result_text.grid(row=1, column=0, padx=18, pady=(8, 14), sticky="nsew")

    def _on_fpa_paste_box_click(self, event=None):
        """Limpia el texto de ayuda cuando el usuario hace clic en la zona de pegado."""
        try:
            current = self.fpa_paste_box.get("1.0", "end").strip()
            if current == "ZONA DE PEGADO: haz clic aquí y presiona Ctrl+V":
                self.fpa_paste_box.delete("1.0", "end")
        except Exception:
            pass
        return None

    # ------------------------------------------------------------------
    # NUEVA PESTAÑA: Validación Visual SAP GUI & Relleno de Formatos
    # ------------------------------------------------------------------
    def _build_tab_val_sap(self):
        tab = self.val_sap_view
        tab.grid_columnconfigure(0, weight=1)
        tab.grid_rowconfigure(0, weight=1)

        scroll = ctk.CTkScrollableFrame(tab, fg_color=CEMEX_GRAY_BG)
        scroll.grid(row=0, column=0, sticky="nsew")
        scroll.grid_columnconfigure(0, weight=1)

        # 1. Card: Zona de Carga (Drag & Drop + Selector)
        card_drop = self._create_card(scroll, "1. Carga de Formato para Validación (Outlook / Sinergia / BSO)")
        card_drop.grid(row=0, column=0, padx=15, pady=(15, 10), sticky="nsew")
        card_drop.grid_columnconfigure(0, weight=1)

        drop_zone = ctk.CTkFrame(card_drop, fg_color="#F8FAFC", corner_radius=10, border_width=2, border_color=CEMEX_BLUE)
        drop_zone.grid(row=1, column=0, padx=18, pady=(8, 12), sticky="ew")
        drop_zone.grid_columnconfigure(0, weight=1)

        ctk.CTkLabel(
            drop_zone,
            text="📂 ARRASTRA AQUÍ EL ARCHIVO DE OUTLOOK (.XLSX / .XLSM) O HAZ CLIC EN EXAMINAR",
            font=CTkFont("Segoe UI", 13, "bold"),
            text_color=CEMEX_BLUE,
        ).grid(row=0, column=0, pady=(14, 4))

        ctk.CTkLabel(
            drop_zone,
            text="Soporta FPA Filial (VK13), CTO Compra (Z0116_TRAOPE), FPA Terceros (VK13) y CTO Venta (VA43 / ZSDD4501)",
            font=CTkFont("Segoe UI", 11),
            text_color=CEMEX_TEXT,
        ).grid(row=1, column=0, pady=(0, 10))

        btn_browse = ctk.CTkButton(
            drop_zone,
            text="Examinar Archivo...",
            command=self._val_sap_examinar_archivo,
            fg_color=CEMEX_BLUE,
            hover_color=CEMEX_BLUE_DARK,
            font=CTkFont("Segoe UI", 12, "bold"),
            height=32,
            width=180,
        )
        btn_browse.grid(row=2, column=0, pady=(0, 14))

        try:
            drop_zone.drop_target_register("DND_Files")
            drop_zone.dnd_bind("<<Drop>>", self._on_val_sap_drop)
        except Exception as e:
            print(f"Aviso registrando DND en zona drop: {e}")

        # Badge de Archivo y Formato Detectado
        info_frame = ctk.CTkFrame(card_drop, fg_color="transparent")
        info_frame.grid(row=2, column=0, padx=18, pady=(0, 12), sticky="ew")
        info_frame.grid_columnconfigure(1, weight=1)

        self.val_sap_lbl_archivo = ctk.CTkLabel(
            info_frame,
            text="Archivo: (Ningún archivo cargado)",
            font=CTkFont("Segoe UI", 11, "bold"),
            text_color=CEMEX_TEXT,
            anchor="w",
        )
        self.val_sap_lbl_archivo.grid(row=0, column=0, sticky="w")

        self.val_sap_lbl_badge = ctk.CTkLabel(
            info_frame,
            text="Formato: Pendiente",
            font=CTkFont("Segoe UI", 11, "bold"),
            text_color=CEMEX_NEUTRAL,
            anchor="e",
        )
        self.val_sap_lbl_badge.grid(row=0, column=1, sticky="e")

        # 2. Card: Controles de Validación en SAP GUI y Confirmación
        card_ctrl = self._create_card(scroll, "2. Parámetros de Validación en SAP GUI y Confirmación")
        card_ctrl.grid(row=1, column=0, padx=15, pady=(0, 10), sticky="nsew")
        card_ctrl.grid_columnconfigure(0, weight=1)

        bar_frame = ctk.CTkFrame(card_ctrl, fg_color="transparent")
        bar_frame.grid(row=1, column=0, padx=18, pady=(8, 12), sticky="ew")
        bar_frame.grid_columnconfigure(6, weight=1)

        # Selector de Velocidad Visual
        ctk.CTkLabel(
            bar_frame,
            text="Velocidad SAP GUI:",
            font=CTkFont("Segoe UI", 11, "bold"),
            text_color=CEMEX_TEXT,
        ).grid(row=0, column=0, padx=(0, 6), sticky="w")

        self.val_sap_speed_var = StringVar(value="👁️ Humano (0.45s)")
        self.val_sap_speed_menu = ctk.CTkOptionMenu(
            bar_frame,
            variable=self.val_sap_speed_var,
            values=["👁️ Humano (0.45s)", "⚡ Rápido (0.15s)", "⏸️ Paso a Paso"],
            width=160,
            fg_color=CEMEX_BLUE,
            button_color=CEMEX_BLUE_DARK,
            font=CTkFont("Segoe UI", 11),
        )
        self.val_sap_speed_menu.grid(row=0, column=1, padx=(0, 16), sticky="w")

        # Campo Fecha de Confirmación
        ctk.CTkLabel(
            bar_frame,
            text="Fecha Confirmación:",
            font=CTkFont("Segoe UI", 11, "bold"),
            text_color=CEMEX_TEXT,
        ).grid(row=0, column=2, padx=(0, 6), sticky="w")

        self.val_sap_fecha_conf_var = StringVar(value=datetime.now().strftime("%d/%m/%Y"))
        self.val_sap_fecha_conf_entry = ctk.CTkEntry(
            bar_frame,
            textvariable=self.val_sap_fecha_conf_var,
            width=120,
            font=CTkFont("Segoe UI", 11),
            fg_color=CEMEX_WHITE,
            border_color=CEMEX_LIGHT_GRAY,
            text_color=CEMEX_TEXT,
        )
        self.val_sap_fecha_conf_entry.grid(row=0, column=3, padx=(0, 16), sticky="w")

        # Botón Iniciar
        self.val_sap_btn_iniciar = ctk.CTkButton(
            bar_frame,
            text="▶ Iniciar Validación en SAP GUI",
            command=self._iniciar_validacion_sap,
            fg_color=CEMEX_BLUE,
            hover_color=CEMEX_BLUE_DARK,
            font=CTkFont("Segoe UI", 12, "bold"),
            height=34,
            state="disabled",
        )
        self.val_sap_btn_iniciar.grid(row=0, column=4, padx=(0, 8), sticky="w")

        # Botón Detener
        self.val_sap_btn_detener = ctk.CTkButton(
            bar_frame,
            text="⏹ Detener",
            command=self._detener_validacion_sap,
            fg_color=CEMEX_RED,
            hover_color=CEMEX_RED_DARK,
            font=CTkFont("Segoe UI", 12, "bold"),
            height=34,
            width=90,
            state="disabled",
        )
        self.val_sap_btn_detener.grid(row=0, column=5, padx=(0, 8), sticky="w")

        # Botón Confirmar y Rellenar
        self.val_sap_btn_confirmar = ctk.CTkButton(
            bar_frame,
            text="💾 Confirmar y Rellenar Formato Maestro",
            command=self._confirmar_y_rellenar_maestro,
            fg_color="#2E7D32",
            hover_color="#1B5E20",
            font=CTkFont("Segoe UI", 12, "bold"),
            height=34,
            state="disabled",
        )
        self.val_sap_btn_confirmar.grid(row=0, column=6, sticky="e")

        # 3. Card: Tabla de Resultados en Tiempo Real
        card_tabla = self._create_card(scroll, "3. Rutas y Resultados de Auditoría en SAP GUI")
        card_tabla.grid(row=2, column=0, padx=15, pady=(0, 10), sticky="nsew")
        card_tabla.grid_columnconfigure(0, weight=1)

        frame_tree = ctk.CTkFrame(card_tabla, fg_color=CEMEX_WHITE, corner_radius=0)
        frame_tree.grid(row=1, column=0, padx=18, pady=(8, 12), sticky="nsew")
        frame_tree.grid_columnconfigure(0, weight=1)
        frame_tree.grid_rowconfigure(0, weight=1)

        val_columns = ("ESTADO", "FILA", "RUTA / CLAVE", "SOLICITADO (ARCHIVO)", "RESULTADO EN SAP", "OBSERVACIÓN / AUDITORÍA")
        self.val_sap_tree = ttk.Treeview(
            frame_tree,
            columns=val_columns,
            show="headings",
            height=10,
            style="Cemex.Treeview",
        )
        self.val_sap_tree.tag_configure("ok", background="#E8F5E9", foreground="#1B5E20")
        self.val_sap_tree.tag_configure("error", background="#FFEBEE", foreground="#B71C1C")
        self.val_sap_tree.tag_configure("running", background="#FFF9C4", foreground="#F57F17")
        self.val_sap_tree.tag_configure("even", background=CEMEX_WHITE, foreground=CEMEX_TEXT)
        self.val_sap_tree.tag_configure("odd", background=CEMEX_GRAY_BG, foreground=CEMEX_TEXT)

        self.val_sap_tree.heading("ESTADO", text="Estado SAP")
        self.val_sap_tree.column("ESTADO", width=120, anchor="center")

        self.val_sap_tree.heading("FILA", text="Fila")
        self.val_sap_tree.column("FILA", width=50, anchor="center")

        self.val_sap_tree.heading("RUTA / CLAVE", text="Ruta / Clave Compuesta")
        self.val_sap_tree.column("RUTA / CLAVE", width=340, anchor="w")

        self.val_sap_tree.heading("SOLICITADO (ARCHIVO)", text="Solicitado en Archivo")
        self.val_sap_tree.column("SOLICITADO (ARCHIVO)", width=170, anchor="center")

        self.val_sap_tree.heading("RESULTADO EN SAP", text="Encontrado en SAP")
        self.val_sap_tree.column("RESULTADO EN SAP", width=170, anchor="center")

        self.val_sap_tree.heading("OBSERVACIÓN / AUDITORÍA", text="Detalle de Auditoría")
        self.val_sap_tree.column("OBSERVACIÓN / AUDITORÍA", width=300, anchor="w")

        vsb = ttk.Scrollbar(frame_tree, orient="vertical", command=self.val_sap_tree.yview)
        vsb.grid(row=0, column=1, sticky="ns")
        hsb = ttk.Scrollbar(frame_tree, orient="horizontal", command=self.val_sap_tree.xview)
        hsb.grid(row=1, column=0, sticky="ew")
        self.val_sap_tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)
        self.val_sap_tree.grid(row=0, column=0, sticky="nsew")

        # 4. Card: Consola de Auditoría en Tiempo Real
        card_log = self._create_card(scroll, "4. Consola de Registro y Comunicación en Vivo con SAP")
        card_log.grid(row=3, column=0, padx=15, pady=(0, 15), sticky="nsew")
        card_log.grid_columnconfigure(0, weight=1)

        self.val_sap_log_box = ctk.CTkTextbox(
            card_log,
            height=110,
            wrap="none",
            font=CTkFont("Consolas", 10),
            fg_color="#1E1E1E",
            text_color="#00FF66",
        )
        self.val_sap_log_box.grid(row=1, column=0, padx=18, pady=(8, 14), sticky="ew")
        self.val_sap_log_box.insert("1.0", "[LISTO] Arrastra un archivo o haz clic en Examinar para comenzar.\n")

    # ------------------------------------------------------------------
    # Métodos de Carga, Validación SAP y Confirmación
    # ------------------------------------------------------------------
    def _val_sap_examinar_archivo(self):
        """Abre un diálogo de selección para cargar un archivo Excel."""
        path = filedialog.askopenfilename(
            title="Seleccionar archivo de validación",
            filetypes=[("Archivos de Excel", "*.xlsx *.xlsm"), ("Todos los archivos", "*.*")],
        )
        if path:
            self._cargar_archivo_val_sap(path)

    def _on_val_sap_drop(self, event):
        """Maneja el evento Drop de archivos en la pestaña de Validación SAP."""
        try:
            data = event.data
            if not data:
                return
            ruta = data.strip()
            if ruta.startswith("{") and ruta.endswith("}"):
                ruta = ruta[1:-1]
            ruta = os.path.normpath(ruta)
            if not os.path.exists(ruta):
                messagebox.showerror("Error", f"No se encontró el archivo:\n{ruta}")
                return
            if not ruta.lower().endswith((".xlsx", ".xlsm")):
                messagebox.showwarning("Formato no compatible", "Solo se admiten archivos .xlsx o .xlsm.")
                return
            self._cargar_archivo_val_sap(ruta)
        except Exception as e:
            messagebox.showerror("Error Drop", f"No se pudo cargar el archivo:\n{e}")

    def _cargar_archivo_val_sap(self, path):
        """Carga el archivo, detecta su formato y llena la tabla de previsualización."""
        try:
            self._val_sap_log(f"Cargando archivo: {os.path.basename(path)}...")
            tipo_info = sap_validator_engine.detectar_tipo_archivo(path)
            items = sap_validator_engine.extraer_filas_archivo(path, tipo_info)

            if not items:
                messagebox.showwarning("Archivo vacío", "No se encontraron rutas con datos para validar.")
                return

            self.val_sap_file_path = path
            self.val_sap_tipo_info = tipo_info
            self.val_sap_items = items
            self.val_sap_approved_items = []

            # Actualizar etiquetas de cabecera
            self.val_sap_lbl_archivo.configure(text=f"Archivo: {os.path.basename(path)}")
            n_hojas = len(tipo_info.get("hojas_detectadas", [1]))
            badge_text = f"Formato: {tipo_info['descripcion']} ({len(items)} registros)"
            self.val_sap_lbl_badge.configure(text=badge_text, text_color=CEMEX_BLUE)

            # Llenar Treeview con filas pendientes
            for item in self.val_sap_tree.get_children():
                self.val_sap_tree.delete(item)

            for i, it in enumerate(items):
                tag = "even" if i % 2 == 0 else "odd"
                hoja_origen = it.get("hoja", "")
                fila_origen = it.get("fila_origen", it.get("idx_fila", ""))
                fila_label = f"{hoja_origen}: {fila_origen}" if hoja_origen else str(fila_origen)

                self.val_sap_tree.insert(
                    "", "end",
                    iid=str(i),
                    values=(
                        "⏳ Pendiente",
                        fila_label,
                        it["clave_str"],
                        it["precio_str"],
                        "-",
                        "Listo para consultar en SAP GUI"
                    ),
                    tags=(tag,)
                )

            self.val_sap_btn_iniciar.configure(state="normal")
            self.val_sap_btn_confirmar.configure(state="disabled")
            self._val_sap_log(f"Archivo cargado exitosamente. Se identificaron {len(items)} registros en {n_hojas} hoja(s): {tipo_info['descripcion']}.")

        except Exception as e:
            messagebox.showerror("Error al leer archivo", f"No se pudo procesar el formato del archivo:\n{e}")
            self._val_sap_log(f"Error procesando archivo: {e}")

    def _val_sap_log(self, texto):
        """Escribe un mensaje con marca de tiempo en la consola de auditoría."""
        ts = datetime.now().strftime("%H:%M:%S")
        self.val_sap_log_box.insert("end", f"[{ts}] {texto}\n")
        self.val_sap_log_box.see("end")

    def _iniciar_validacion_sap(self):
        """Inicia la validación en vivo en un hilo secundario."""
        if not self.val_sap_items:
            messagebox.showwarning("Sin datos", "Por favor arrastra o carga un archivo primero.")
            return

        # Probar conexión previa a SAP GUI
        try:
            sap_validator_engine.get_sap_session(bring_to_foreground=True)
        except Exception as e:
            messagebox.showerror(
                "SAP GUI No Disponible",
                str(e)
            )
            return

        # Calcular delay según opción seleccionada
        vel_str = self.val_sap_speed_var.get()
        if "Rápido" in vel_str:
            delay = 0.15
        elif "Paso" in vel_str:
            delay = 1.0
        else:
            delay = 0.45

        self.val_sap_btn_iniciar.configure(state="disabled")
        self.val_sap_btn_detener.configure(state="normal")
        self.val_sap_btn_confirmar.configure(state="disabled")
        self.val_sap_approved_items = []

        self.val_sap_motor = sap_validator_engine.MotorValidadorSAP(
            speed_delay=delay,
            on_progress=self._on_sap_progress_bridge,
            on_log=self._on_sap_log_bridge,
        )

        self._val_sap_log("Iniciando validación en vivo en SAP GUI. Observa la ventana de SAP...")
        threading.Thread(target=self._run_sap_validation_worker, daemon=True).start()

    def _detener_validacion_sap(self):
        """Cancela la ejecución de la validación en vivo."""
        if self.val_sap_motor:
            self.val_sap_motor.stop_event.set()
            self._val_sap_log("Solicitud de detención enviada al motor SAP...")
        self.val_sap_btn_detener.configure(state="disabled")

    def _run_sap_validation_worker(self):
        """Bucle de validación ejecutado en segundo plano."""
        import pythoncom
        try:
            pythoncom.CoInitialize()
        except Exception:
            pass
        try:
            resultados = self.val_sap_motor.validar_lote(
                self.val_sap_items, self.val_sap_tipo_info, ruta_excel=self.val_sap_file_path
            )
            self.after(0, self._on_sap_validation_finished, resultados, None)
        except Exception as e:
            self.after(0, self._on_sap_validation_finished, [], e)
        finally:
            try:
                pythoncom.CoUninitialize()
            except Exception:
                pass

    def _on_sap_progress_bridge(self, idx, total, item, estado, datos_sap, msg):
        """Puente thread-safe hacia el hilo principal de la GUI."""
        self.after(0, self._update_val_sap_tree_row, idx, total, item, estado, datos_sap, msg)

    def _update_val_sap_tree_row(self, idx, total, item, estado, datos_sap, msg):
        """Actualiza la fila correspondiente en el Treeview con estados y colores."""
        iid = str(idx - 1)
        if not self.val_sap_tree.exists(iid):
            return

        if estado == "EN_PROCESO":
            estado_texto = "🔄 Validando..."
            tag = "running"
            res_sap = "Consultando..."
        elif estado == "OK":
            estado_texto = "✅ OK"
            tag = "ok"
            # Extraer precio o contrato para la columna de resultado
            if "contrato" in datos_sap:
                if datos_sap.get("contrato") == "(Por crear)":
                    res_sap = "Sin duplicados (ZSDD4501)"
                else:
                    pos = datos_sap.get("posicion", "")
                    pos_str = f" Pos {pos}" if pos else ""
                    p = datos_sap.get("precio")
                    p_str = f" (${p:,.2f})" if (isinstance(p, (int, float)) and p is not None) else ""
                    res_sap = f"Cto {datos_sap.get('contrato')}{pos_str}{p_str}" if (datos_sap.get('contrato') or pos_str) else "Coincide"
            elif "importe" in datos_sap:
                res_sap = f"${datos_sap.get('importe', 0):,.2f}"
            else:
                res_sap = "Coincide"
        elif estado == "DISCREPANCIA":
            estado_texto = "❌ Discrepancia"
            tag = "error"
            if "contrato" in datos_sap and datos_sap.get("contrato") and datos_sap.get("contrato") != "(Por crear)":
                pos = datos_sap.get("posicion", "")
                pos_str = f" Pos {pos}" if pos else ""
                p = datos_sap.get("precio")
                p_str = f" (${p:,.2f})" if (isinstance(p, (int, float)) and p is not None) else ""
                res_sap = f"Cto {datos_sap.get('contrato')}{pos_str}{p_str}"
            elif "importe" in datos_sap:
                res_sap = f"${datos_sap.get('importe', 0):,.2f}"
            else:
                res_sap = "Difiere"
        elif estado == "NO_ENCONTRADO":
            estado_texto = "❌ No Encontrado"
            tag = "error"
            res_sap = "No existe en SAP"
        else:
            estado_texto = "⚠️ Error"
            tag = "error"
            res_sap = "Error SAP"

        hoja_origen = item.get("hoja", "")
        fila_origen = item.get("fila_origen", item.get("idx_fila", ""))
        fila_label = f"{hoja_origen}: {fila_origen}" if hoja_origen else str(fila_origen)

        self.val_sap_tree.item(
            iid,
            values=(
                estado_texto,
                fila_label,
                item["clave_str"],
                item["precio_str"],
                res_sap,
                msg
            ),
            tags=(tag,)
        )
        self.val_sap_tree.see(iid)

    def _on_sap_log_bridge(self, text):
        """Puente thread-safe para los logs de auditoría."""
        self.after(0, self._val_sap_log, text)

    def _on_sap_validation_finished(self, resultados, error):
        """Llamado al finalizar la validación por completo."""
        self.val_sap_btn_iniciar.configure(state="normal")
        self.val_sap_btn_detener.configure(state="disabled")

        if error:
            messagebox.showerror("Error en Validación SAP", f"Ocurrió un problema durante la ejecución:\n{error}")
            self._val_sap_log(f"Finalizado con error: {error}")
            return

        aprobados = [it for it, res in resultados if res.get("estado") == "OK"]
        self.val_sap_approved_items = aprobados

        # Ubicar archivo validado si se generó
        archivo_validado = ""
        if self.val_sap_file_path:
            directorio = os.path.dirname(self.val_sap_file_path)
            base = os.path.splitext(os.path.basename(self.val_sap_file_path))[0]
            cand = os.path.join(directorio, f"{base}_VALIDADO.xlsx")
            if os.path.exists(cand):
                archivo_validado = cand

        if aprobados:
            self.val_sap_btn_confirmar.configure(state="normal")
            msg_log = f"¡Éxito! Se validaron correctamente {len(aprobados)} de {len(resultados)} rutas."
            if archivo_validado:
                msg_log += f" Archivo generado: {os.path.basename(archivo_validado)}"
            self._val_sap_log(msg_log)

            msg_info = (
                f"Validación finalizada con éxito:\n\n"
                f"• Rutas validadas correctamente en SAP: {len(aprobados)}\n"
                f"• Total de rutas procesadas: {len(resultados)}\n"
            )
            if archivo_validado:
                msg_info += f"• Archivo con auditoría generado: {os.path.basename(archivo_validado)}\n\n"
            msg_info += "Ahora puedes hacer clic en 'Confirmar y Rellenar Formato Maestro' para actualizar la base de datos y el Excel Maestro."

            messagebox.showinfo("Validación SAP Concluida", msg_info)
        else:
            self.val_sap_btn_confirmar.configure(state="disabled")
            self._val_sap_log("La validación concluyó pero no hubo rutas con estado OK para confirmar.")
            messagebox.showwarning(
                "Sin Coincidencias Exitosas",
                "Se procesaron las rutas pero ninguna cumplió los criterios de validación exitosa (OK). Revisa los detalles de auditoría en la tabla."
            )

    def _confirmar_y_rellenar_maestro(self):
        """Aplica la Fecha de Confirmación a las filas aprobadas en SQLite y en el Excel Maestro."""
        if not self.val_sap_approved_items:
            messagebox.showwarning("Sin aprobados", "No hay rutas validadas OK para confirmar.")
            return

        fecha = self.val_sap_fecha_conf_var.get().strip()
        if not fecha:
            messagebox.showwarning("Fecha requerida", "Por favor ingresa una Fecha de Confirmación válida.")
            return

        resp = messagebox.askyesno(
            "Confirmar Rutas",
            f"Se aplicará la Fecha de Confirmación: '{fecha}'\n"
            f"a {len(self.val_sap_approved_items)} ruta(s) aprobada(s) en la base de datos local y en el Excel Maestro.\n\n"
            f"¿Deseas continuar?"
        )
        if not resp:
            return

        self.val_sap_btn_confirmar.configure(state="disabled", text="Aplicando...")

        try:
            n_sqlite, n_excel, detalles = sap_validator_engine.aplicar_confirmacion_lote(
                self.conn, self.val_sap_approved_items, fecha_conf=fecha
            )
            self._val_sap_log(detalles)
            self._refresh_sla_tree()
            messagebox.showinfo("Confirmación Aplicada", detalles)
        except Exception as e:
            messagebox.showerror("Error al Confirmar", f"No se pudo completar la confirmación:\n{e}")
            self._val_sap_log(f"Error confirmando: {e}")
        finally:
            self.val_sap_btn_confirmar.configure(state="normal", text="💾 Confirmar y Rellenar Formato Maestro")

    # ---------------------------------------------------------------------------
    # VISTA: VALIDACIÓN DE MATRIZ DE PRECIOS CONSOLIDADA (VK13 / COMPRA / VENTA)
    # ---------------------------------------------------------------------------
    def _build_tab_val_matriz(self):
        tab = self.val_matriz_view
        tab.grid_columnconfigure(0, weight=1)
        tab.grid_rowconfigure(0, weight=1)

        scroll = ctk.CTkScrollableFrame(tab, fg_color=CEMEX_GRAY_BG)
        scroll.grid(row=0, column=0, sticky="nsew")
        scroll.grid_columnconfigure(0, weight=1)

        # 1. Card: Zona de Carga de Matriz
        card_drop = self._create_card(scroll, "1. Carga de Matriz de Precios Consolidada (VK13, Compra y Venta)")
        card_drop.grid(row=0, column=0, padx=15, pady=(15, 10), sticky="nsew")
        card_drop.grid_columnconfigure(0, weight=1)

        drop_zone = ctk.CTkFrame(card_drop, fg_color="#F8FAFC", corner_radius=10, border_width=2, border_color=CEMEX_BLUE)
        drop_zone.grid(row=1, column=0, padx=18, pady=(8, 12), sticky="ew")
        drop_zone.grid_columnconfigure(0, weight=1)

        ctk.CTkLabel(
            drop_zone,
            text="📊 ARRASTRA AQUÍ EL ARCHIVO 'Matriz_Precios_...xlsx' O HAZ CLIC EN EXAMINAR",
            font=CTkFont("Segoe UI", 13, "bold"),
            text_color=CEMEX_BLUE,
        ).grid(row=0, column=0, pady=(14, 4))

        ctk.CTkLabel(
            drop_zone,
            text="Audita integralmente cada fila: Precios de Venta (VK13), Contratos de Compra (TRAOPE) y Contratos de Venta (ZCP)",
            font=CTkFont("Segoe UI", 11),
            text_color=CEMEX_TEXT,
        ).grid(row=1, column=0, pady=(0, 10))

        btn_browse = ctk.CTkButton(
            drop_zone,
            text="Examinar Archivo Matriz...",
            command=self._val_matriz_examinar_archivo,
            fg_color=CEMEX_BLUE,
            hover_color=CEMEX_BLUE_DARK,
            font=CTkFont("Segoe UI", 12, "bold"),
            height=32,
            width=220,
        )
        btn_browse.grid(row=2, column=0, pady=(0, 14))

        try:
            drop_zone.drop_target_register("DND_Files")
            drop_zone.dnd_bind("<<Drop>>", self._on_val_matriz_drop)
        except Exception as e:
            print(f"Aviso registrando DND en zona matriz: {e}")

        info_frame = ctk.CTkFrame(card_drop, fg_color="transparent")
        info_frame.grid(row=2, column=0, padx=18, pady=(0, 12), sticky="ew")
        info_frame.grid_columnconfigure(1, weight=1)

        self.val_matriz_lbl_archivo = ctk.CTkLabel(
            info_frame,
            text="Archivo: (Ningún archivo cargado)",
            font=CTkFont("Segoe UI", 11, "bold"),
            text_color=CEMEX_TEXT,
            anchor="w",
        )
        self.val_matriz_lbl_archivo.grid(row=0, column=0, sticky="w")

        self.val_matriz_lbl_badge = ctk.CTkLabel(
            info_frame,
            text="Formato: Pendiente de carga",
            font=CTkFont("Segoe UI", 11, "bold"),
            text_color=CEMEX_NEUTRAL,
            anchor="e",
        )
        self.val_matriz_lbl_badge.grid(row=0, column=1, sticky="e")

        # 2. Card: Parámetros y Filtros de Selección
        card_ctrl = self._create_card(scroll, "2. Parámetros de Auditoría y Filtros de Lote")
        card_ctrl.grid(row=1, column=0, padx=15, pady=(0, 10), sticky="nsew")
        card_ctrl.grid_columnconfigure(0, weight=1)

        # Fila 0: Selección de Módulos (Motores a Auditar)
        f_mods = ctk.CTkFrame(card_ctrl, fg_color="transparent")
        f_mods.grid(row=1, column=0, padx=18, pady=(8, 4), sticky="ew")

        ctk.CTkLabel(
            f_mods,
            text="Motores SAP a Auditar:",
            font=CTkFont("Segoe UI", 11, "bold"),
            text_color=CEMEX_TEXT,
        ).grid(row=0, column=0, padx=(0, 12), sticky="w")

        self.val_matriz_chk_vk13_var = BooleanVar(value=True)
        self.val_matriz_chk_vk13 = ctk.CTkCheckBox(
            f_mods,
            text="VK13 (Precios Venta)",
            variable=self.val_matriz_chk_vk13_var,
            font=CTkFont("Segoe UI", 11),
            fg_color=CEMEX_BLUE,
        )
        self.val_matriz_chk_vk13.grid(row=0, column=1, padx=(0, 16), sticky="w")

        self.val_matriz_chk_compra_var = BooleanVar(value=True)
        self.val_matriz_chk_compra = ctk.CTkCheckBox(
            f_mods,
            text="Contrato Compra (Z0116_TRAOPE)",
            variable=self.val_matriz_chk_compra_var,
            font=CTkFont("Segoe UI", 11),
            fg_color=CEMEX_BLUE,
        )
        self.val_matriz_chk_compra.grid(row=0, column=2, padx=(0, 16), sticky="w")

        self.val_matriz_chk_venta_var = BooleanVar(value=True)
        self.val_matriz_chk_venta = ctk.CTkCheckBox(
            f_mods,
            text="Contrato Venta (VA43)",
            variable=self.val_matriz_chk_venta_var,
            font=CTkFont("Segoe UI", 11),
            fg_color=CEMEX_BLUE,
        )
        self.val_matriz_chk_venta.grid(row=0, column=3, padx=(0, 16), sticky="w")

        f_filters = ctk.CTkFrame(card_ctrl, fg_color="transparent")
        f_filters.grid(row=2, column=0, padx=18, pady=(4, 8), sticky="ew")

        ctk.CTkLabel(
            f_filters,
            text="Modo de Selección:",
            font=CTkFont("Segoe UI", 11, "bold"),
            text_color=CEMEX_TEXT,
        ).grid(row=0, column=0, padx=(0, 6), pady=4, sticky="w")

        self.val_matriz_modo_var = StringVar(value="Todo (Todas las filas)")
        self.val_matriz_modo_menu = ctk.CTkOptionMenu(
            f_filters,
            variable=self.val_matriz_modo_var,
            values=["Todo (Todas las filas)", "Filtrar por Centro", "Filtrar por Sociedad", "Rango de Prueba (Filas)"],
            command=self._on_val_matriz_modo_changed,
            width=180,
            fg_color=CEMEX_BLUE,
            button_color=CEMEX_BLUE_DARK,
            font=CTkFont("Segoe UI", 11),
        )
        self.val_matriz_modo_menu.grid(row=0, column=1, padx=(0, 16), pady=4, sticky="w")

        ctk.CTkLabel(
            f_filters,
            text="Centro:",
            font=CTkFont("Segoe UI", 11, "bold"),
            text_color=CEMEX_TEXT,
        ).grid(row=0, column=2, padx=(0, 6), pady=4, sticky="w")

        self.val_matriz_centro_var = StringVar(value="Todos")
        self.val_matriz_centro_combo = ctk.CTkComboBox(
            f_filters,
            variable=self.val_matriz_centro_var,
            values=["Todos"],
            width=120,
            font=CTkFont("Segoe UI", 11),
            command=lambda _: self._aplicar_filtro_matriz(),
        )
        self.val_matriz_centro_combo.grid(row=0, column=3, padx=(0, 16), pady=4, sticky="w")

        ctk.CTkLabel(
            f_filters,
            text="Sociedad:",
            font=CTkFont("Segoe UI", 11, "bold"),
            text_color=CEMEX_TEXT,
        ).grid(row=0, column=4, padx=(0, 6), pady=4, sticky="w")

        self.val_matriz_soc_var = StringVar(value="Todas")
        self.val_matriz_soc_combo = ctk.CTkComboBox(
            f_filters,
            variable=self.val_matriz_soc_var,
            values=["Todas", "7100", "7180", "7277"],
            width=100,
            font=CTkFont("Segoe UI", 11),
            command=lambda _: self._aplicar_filtro_matriz(),
        )
        self.val_matriz_soc_combo.grid(row=0, column=5, padx=(0, 16), pady=4, sticky="w")

        # Fila 2 de controles: Rango y Velocidad
        f_row2 = ctk.CTkFrame(card_ctrl, fg_color="transparent")
        f_row2.grid(row=3, column=0, padx=18, pady=(0, 8), sticky="ew")

        ctk.CTkLabel(
            f_row2,
            text="Rango Filas: Desde:",
            font=CTkFont("Segoe UI", 11, "bold"),
            text_color=CEMEX_TEXT,
        ).grid(row=0, column=0, padx=(0, 4), sticky="w")
        self.val_matriz_desde_var = StringVar(value="1")
        self.val_matriz_desde_entry = ctk.CTkEntry(
            f_row2,
            textvariable=self.val_matriz_desde_var,
            width=60,
            font=CTkFont("Segoe UI", 11),
        )
        self.val_matriz_desde_entry.grid(row=0, column=1, padx=(0, 10), sticky="w")

        ctk.CTkLabel(
            f_row2,
            text="Hasta:",
            font=CTkFont("Segoe UI", 11, "bold"),
            text_color=CEMEX_TEXT,
        ).grid(row=0, column=2, padx=(0, 4), sticky="w")
        self.val_matriz_hasta_var = StringVar(value="50")
        self.val_matriz_hasta_entry = ctk.CTkEntry(
            f_row2,
            textvariable=self.val_matriz_hasta_var,
            width=60,
            font=CTkFont("Segoe UI", 11),
        )
        self.val_matriz_hasta_entry.grid(row=0, column=3, padx=(0, 14), sticky="w")

        btn_aplicar_filtro = ctk.CTkButton(
            f_row2,
            text="🔍 Aplicar Filtro",
            command=self._aplicar_filtro_matriz,
            width=110,
            height=30,
            fg_color=CEMEX_SIDEBAR_DARK,
            hover_color=CEMEX_SIDEBAR,
            font=CTkFont("Segoe UI", 11, "bold"),
        )
        btn_aplicar_filtro.grid(row=0, column=4, padx=(0, 20), sticky="w")

        ctk.CTkLabel(
            f_row2,
            text="Velocidad SAP:",
            font=CTkFont("Segoe UI", 11, "bold"),
            text_color=CEMEX_TEXT,
        ).grid(row=0, column=5, padx=(0, 4), sticky="w")
        self.val_matriz_speed_var = StringVar(value="👁️ Humano (0.45s)")
        self.val_matriz_speed_menu = ctk.CTkOptionMenu(
            f_row2,
            variable=self.val_matriz_speed_var,
            values=["👁️ Humano (0.45s)", "⚡ Rápido (0.15s)", "⏸️ Paso a Paso"],
            width=150,
            fg_color=CEMEX_BLUE,
            button_color=CEMEX_BLUE_DARK,
            font=CTkFont("Segoe UI", 11),
        )
        self.val_matriz_speed_menu.grid(row=0, column=6, padx=(0, 16), sticky="w")

        # Fila 3: Botones de Acción
        f_actions = ctk.CTkFrame(card_ctrl, fg_color="transparent")
        f_actions.grid(row=4, column=0, padx=18, pady=(4, 14), sticky="ew")

        self.val_matriz_btn_iniciar = ctk.CTkButton(
            f_actions,
            text="▶ Iniciar Auditoría Masiva SAP",
            command=self._iniciar_validacion_matriz,
            fg_color=CEMEX_BLUE,
            hover_color=CEMEX_BLUE_DARK,
            font=CTkFont("Segoe UI", 12, "bold"),
            height=36,
            state="disabled",
        )
        self.val_matriz_btn_iniciar.grid(row=0, column=0, padx=(0, 10), sticky="w")

        self.val_matriz_btn_detener = ctk.CTkButton(
            f_actions,
            text="⏹ Detener",
            command=self._detener_validacion_matriz,
            fg_color=CEMEX_RED,
            hover_color=CEMEX_RED_DARK,
            font=CTkFont("Segoe UI", 12, "bold"),
            height=36,
            width=100,
            state="disabled",
        )
        self.val_matriz_btn_detener.grid(row=0, column=1, padx=(0, 10), sticky="w")

        self.val_matriz_btn_abrir = ctk.CTkButton(
            f_actions,
            text="📂 Abrir Excel _VALIDADO.xlsx",
            command=self._abrir_excel_matriz_validado,
            fg_color="#2E7D32",
            hover_color="#1B5E20",
            font=CTkFont("Segoe UI", 12, "bold"),
            height=36,
            state="disabled",
        )
        self.val_matriz_btn_abrir.grid(row=0, column=2, sticky="e")

        # 3. Card: Tabla de Monitoreo Multimódulo
        card_tabla = self._create_card(scroll, "3. Monitoreo Multimódulo en Tiempo Real (VK13 / Compra / Venta)")
        card_tabla.grid(row=2, column=0, padx=15, pady=(0, 10), sticky="nsew")
        card_tabla.grid_columnconfigure(0, weight=1)

        frame_tree = ctk.CTkFrame(card_tabla, fg_color=CEMEX_WHITE, corner_radius=0)
        frame_tree.grid(row=1, column=0, padx=18, pady=(8, 12), sticky="nsew")
        frame_tree.grid_columnconfigure(0, weight=1)
        frame_tree.grid_rowconfigure(0, weight=1)

        mat_cols = ("ESTADO", "FILA", "RUTA / CLAVE", "VK13 (SOL/SAP)", "COMPRA (CTO/PRC)", "VENTA (CTO/PRC)", "DICTAMEN GLOBAL / OBSERVACIÓN")
        self.val_matriz_tree = ttk.Treeview(
            frame_tree,
            columns=mat_cols,
            show="headings",
            height=11,
            style="Cemex.Treeview",
        )
        self.val_matriz_tree.tag_configure("ok", background="#E8F5E9", foreground="#1B5E20")
        self.val_matriz_tree.tag_configure("error", background="#FFEBEE", foreground="#B71C1C")
        self.val_matriz_tree.tag_configure("warning", background="#FFF9C4", foreground="#F57F17")
        self.val_matriz_tree.tag_configure("running", background="#E3F2FD", foreground="#0D47A1")
        self.val_matriz_tree.tag_configure("even", background=CEMEX_WHITE, foreground=CEMEX_TEXT)
        self.val_matriz_tree.tag_configure("odd", background=CEMEX_GRAY_BG, foreground=CEMEX_TEXT)

        self.val_matriz_tree.heading("ESTADO", text="Semáforo")
        self.val_matriz_tree.column("ESTADO", width=110, anchor="center")

        self.val_matriz_tree.heading("FILA", text="Fila")
        self.val_matriz_tree.column("FILA", width=55, anchor="center")

        self.val_matriz_tree.heading("RUTA / CLAVE", text="Ruta / Clave Matriz")
        self.val_matriz_tree.column("RUTA / CLAVE", width=290, anchor="w")

        self.val_matriz_tree.heading("VK13 (SOL/SAP)", text="VK13 (Venta)")
        self.val_matriz_tree.column("VK13 (SOL/SAP)", width=150, anchor="center")

        self.val_matriz_tree.heading("COMPRA (CTO/PRC)", text="Contrato Compra")
        self.val_matriz_tree.column("COMPRA (CTO/PRC)", width=160, anchor="center")

        self.val_matriz_tree.heading("VENTA (CTO/PRC)", text="Contrato Venta")
        self.val_matriz_tree.column("VENTA (CTO/PRC)", width=150, anchor="center")

        self.val_matriz_tree.heading("DICTAMEN GLOBAL / OBSERVACIÓN", text="Dictamen Global de Auditoría")
        self.val_matriz_tree.column("DICTAMEN GLOBAL / OBSERVACIÓN", width=340, anchor="w")

        vsb = ttk.Scrollbar(frame_tree, orient="vertical", command=self.val_matriz_tree.yview)
        vsb.grid(row=0, column=1, sticky="ns")
        hsb = ttk.Scrollbar(frame_tree, orient="horizontal", command=self.val_matriz_tree.xview)
        hsb.grid(row=1, column=0, sticky="ew")
        self.val_matriz_tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)
        self.val_matriz_tree.grid(row=0, column=0, sticky="nsew")

        # 4. Card: Progreso y Consola
        card_log = self._create_card(scroll, "4. Progreso y Consola de Auditoría Masiva en Vivo")
        card_log.grid(row=3, column=0, padx=15, pady=(0, 15), sticky="nsew")
        card_log.grid_columnconfigure(0, weight=1)

        prog_frame = ctk.CTkFrame(card_log, fg_color="transparent")
        prog_frame.grid(row=1, column=0, padx=18, pady=(4, 6), sticky="ew")
        prog_frame.grid_columnconfigure(0, weight=1)

        self.val_matriz_progress_bar = ctk.CTkProgressBar(prog_frame, fg_color=CEMEX_LIGHT_GRAY, progress_color=CEMEX_BLUE)
        self.val_matriz_progress_bar.grid(row=0, column=0, sticky="ew")
        self.val_matriz_progress_bar.set(0)

        self.val_matriz_lbl_progreso = ctk.CTkLabel(
            prog_frame,
            text="Listo para iniciar.",
            font=CTkFont("Segoe UI", 10),
            text_color=CEMEX_TEXT,
        )
        self.val_matriz_lbl_progreso.grid(row=1, column=0, sticky="w", pady=(2, 4))

        self.val_matriz_log_box = ctk.CTkTextbox(
            card_log,
            height=110,
            wrap="none",
            font=CTkFont("Consolas", 10),
            fg_color="#1E1E1E",
            text_color="#00FF66",
        )
        self.val_matriz_log_box.grid(row=2, column=0, padx=18, pady=(0, 12), sticky="ew")

    def _val_matriz_log(self, texto):
        ts = datetime.now().strftime("%H:%M:%S")
        self.val_matriz_log_box.insert("end", f"[{ts}] {texto}\n")
        self.val_matriz_log_box.see("end")

    def _val_matriz_examinar_archivo(self):
        path = filedialog.askopenfilename(
            title="Seleccionar Matriz de Precios",
            filetypes=[("Archivos Excel", "*.xlsx *.xlsm")],
        )
        if path:
            self._cargar_archivo_val_matriz(path)

    def _on_val_matriz_drop(self, event):
        try:
            ruta = event.data.strip()
            if ruta.startswith("{") and ruta.endswith("}"):
                ruta = ruta[1:-1]
            if ruta.lower().endswith((".xlsx", ".xlsm")):
                self._cargar_archivo_val_matriz(ruta)
            else:
                messagebox.showwarning("Formato no compatible", "Solo se admiten archivos .xlsx o .xlsm.")
        except Exception as e:
            messagebox.showerror("Error Drop", f"No se pudo cargar el archivo:\n{e}")

    def _cargar_archivo_val_matriz(self, path):
        try:
            self._val_matriz_log(f"Cargando Matriz: {os.path.basename(path)}...")
            tipo_info = sap_validator_engine.detectar_tipo_archivo(path)
            items = sap_validator_engine.extraer_filas_archivo(path, tipo_info)

            if not items:
                messagebox.showwarning("Archivo vacío", "No se encontraron rutas con datos en la Matriz Precios.")
                return

            self.val_matriz_file_path = path
            self.val_matriz_tipo_info = tipo_info
            self.val_matriz_items = items

            # Extraer centros y sociedades únicos para los filtros
            centros_unicos = sorted(list({str(x.get("cedis", "")).strip() for x in items if x.get("cedis")}))
            sociedades_unicas = sorted(list({str(x.get("sociedad", "")).strip() for x in items if x.get("sociedad")}))

            self.val_matriz_centro_combo.configure(values=["Todos"] + centros_unicos)
            self.val_matriz_centro_var.set("Todos")

            self.val_matriz_soc_combo.configure(values=["Todas"] + sociedades_unicas)
            self.val_matriz_soc_var.set("Todas")

            self.val_matriz_hasta_var.set(str(min(50, len(items))))

            self.val_matriz_lbl_archivo.configure(text=f"Archivo: {os.path.basename(path)}")
            self.val_matriz_lbl_badge.configure(
                text=f"Matriz: {len(items):,} registros | {len(centros_unicos)} Centros",
                text_color=CEMEX_BLUE
            )

            self._aplicar_filtro_matriz()
            self._val_matriz_log(f"Matriz cargada exitosamente: {len(items):,} registros totales listos para auditar.")

        except Exception as e:
            messagebox.showerror("Error al leer matriz", f"No se pudo procesar la matriz:\n{e}")
            self._val_matriz_log(f"Error cargando matriz: {e}")

    def _on_val_matriz_modo_changed(self, choice=None):
        self._aplicar_filtro_matriz()

    def _aplicar_filtro_matriz(self):
        if not self.val_matriz_items:
            return

        modo = self.val_matriz_modo_var.get()
        c_filtro = self.val_matriz_centro_var.get().strip()
        s_filtro = self.val_matriz_soc_var.get().strip()

        filtered = list(self.val_matriz_items)

        if "Centro" in modo and c_filtro != "Todos":
            filtered = [x for x in filtered if str(x.get("cedis", "")).strip() == c_filtro]
        elif "Sociedad" in modo and s_filtro != "Todas":
            filtered = [x for x in filtered if str(x.get("sociedad", "")).strip() == s_filtro]
        elif "Rango" in modo:
            try:
                d_idx = max(1, int(self.val_matriz_desde_var.get().strip()))
                h_idx = min(len(filtered), int(self.val_matriz_hasta_var.get().strip()))
                filtered = filtered[d_idx - 1:h_idx]
            except Exception:
                pass

        self.val_matriz_filtered_items = filtered

        # Llenar Treeview
        for item in self.val_matriz_tree.get_children():
            self.val_matriz_tree.delete(item)

        for i, it in enumerate(filtered[:200]): # Mostrar hasta 200 en preview
            tag = "even" if i % 2 == 0 else "odd"
            f_num = it.get("fila_origen", i + 1)
            vk_txt = f"MP:${it.get('precio_mat', 0):,.2f}" if it.get("precio_mat") is not None else "-"
            cp_txt = f"Cto {it.get('contrato_compra_raw', '')}" if it.get("contrato_compra_raw") else "S/C"
            vt_txt = f"Cto {it.get('contrato_venta_raw', '')}" if it.get("contrato_venta_raw") else "S/C"

            self.val_matriz_tree.insert(
                "", "end",
                iid=str(i),
                values=(
                    "⏳ Pendiente",
                    str(f_num),
                    it.get("clave_str", ""),
                    vk_txt,
                    cp_txt,
                    vt_txt,
                    "Listo para auditar en SAP GUI"
                ),
                tags=(tag,)
            )

        self.val_matriz_btn_iniciar.configure(state="normal")
        self.val_matriz_lbl_progreso.configure(text=f"Filas seleccionadas para auditar: {len(filtered):,} de {len(self.val_matriz_items):,}")
        self._val_matriz_log(f"Filtro aplicado: {len(filtered):,} fila(s) seleccionada(s).")

    def _iniciar_validacion_matriz(self):
        if not self.val_matriz_filtered_items:
            messagebox.showwarning("Sin datos", "Por favor carga un archivo de matriz primero.")
            return

        try:
            sap_validator_engine.get_sap_session(bring_to_foreground=True)
        except Exception as e:
            messagebox.showerror("SAP GUI No Disponible", str(e))
            return

        vel_str = self.val_matriz_speed_var.get()
        delay = 0.15 if "Rápido" in vel_str else (1.0 if "Paso" in vel_str else 0.45)

        self.val_matriz_btn_iniciar.configure(state="disabled")
        self.val_matriz_btn_detener.configure(state="normal")
        self.val_matriz_btn_abrir.configure(state="disabled")

        self.val_matriz_motor = sap_validator_engine.MotorValidadorSAP(
            speed_delay=delay,
            on_progress=self._on_matriz_progress_bridge,
            on_log=self._val_matriz_log,
        )
        self.val_matriz_motor.val_matriz_chk_vk13 = self.val_matriz_chk_vk13_var.get()
        self.val_matriz_motor.val_matriz_chk_compra = self.val_matriz_chk_compra_var.get()
        self.val_matriz_motor.val_matriz_chk_venta = self.val_matriz_chk_venta_var.get()

        self._val_matriz_log(f"Iniciando auditoría de {len(self.val_matriz_filtered_items):,} filas en SAP GUI...")
        threading.Thread(target=self._run_matriz_validation_worker, daemon=True).start()

    def _detener_validacion_matriz(self):
        if self.val_matriz_motor:
            self.val_matriz_motor.stop_event.set()
            self._val_matriz_log("Deteniendo auditoría por solicitud del usuario...")
        self.val_matriz_btn_detener.configure(state="disabled")

    def _run_matriz_validation_worker(self):
        import pythoncom
        try:
            pythoncom.CoInitialize()
        except Exception:
            pass
        try:
            resultados = self.val_matriz_motor.validar_lote(
                self.val_matriz_filtered_items,
                self.val_matriz_tipo_info,
                ruta_excel=self.val_matriz_file_path
            )
            self.after(0, self._on_matriz_validation_finished, resultados, None)
        except Exception as e:
            self.after(0, self._on_matriz_validation_finished, [], e)
        finally:
            try:
                pythoncom.CoUninitialize()
            except Exception:
                pass

    def _on_matriz_progress_bridge(self, idx, total, item, estado, datos_sap, msg):
        self.after(0, self._update_matriz_tree_row, idx, total, item, estado, datos_sap, msg)

    def _update_matriz_tree_row(self, idx, total, item, estado, datos_sap, msg):
        pct = (idx / total) if total > 0 else 0
        self.val_matriz_progress_bar.set(pct)
        c_centro = item.get("cedis", "")
        self.val_matriz_lbl_progreso.configure(
            text=f"Procesando {idx:,} de {total:,} ({pct*100:.1f}%) · Centro: {c_centro}"
        )

        iid = str(idx - 1)
        if not self.val_matriz_tree.exists(iid):
            return

        if estado == "EN_PROCESO":
            estado_texto = "🔄 Validando..."
            tag = "running"
        elif estado == "OK":
            estado_texto = "✅ OK"
            tag = "ok"
        elif estado == "DISCREPANCIA":
            estado_texto = "❌ Discrepancia"
            tag = "error"
        else:
            estado_texto = "⚠️ Revisión"
            tag = "warning"

        vk_info = f"VK13: {datos_sap.get('vk13', {}).get('precio_mat_sap', '-')}" if datos_sap.get('vk13') else "-"
        cp_info = f"Cto: {datos_sap.get('compra', {}).get('contrato', '-')}" if datos_sap.get('compra') else "-"
        vt_info = f"Cto: {datos_sap.get('venta', {}).get('contrato', '-')}" if datos_sap.get('venta') else "-"

        self.val_matriz_tree.item(
            iid,
            values=(
                estado_texto,
                str(item.get("fila_origen", idx)),
                item.get("clave_str", ""),
                vk_info,
                cp_info,
                vt_info,
                msg
            ),
            tags=(tag,)
        )
        self.val_matriz_tree.see(iid)

    def _on_matriz_validation_finished(self, resultados, error):
        self.val_matriz_btn_iniciar.configure(state="normal")
        self.val_matriz_btn_detener.configure(state="disabled")

        if error:
            messagebox.showerror("Error en Auditoría", f"Ocurrió un error:\n{error}")
            self._val_matriz_log(f"Error en auditoría: {error}")
            return

        # Localizar archivo _VALIDADO.xlsx
        directorio = os.path.dirname(self.val_matriz_file_path)
        base = os.path.splitext(os.path.basename(self.val_matriz_file_path))[0]
        out_cand = os.path.join(directorio, f"{base}_VALIDADO.xlsx")
        if os.path.exists(out_cand):
            self.val_matriz_out_path = out_cand
            self.val_matriz_btn_abrir.configure(state="normal")
            self._val_matriz_log(f"Auditoría completada. Archivo guardado: {os.path.basename(out_cand)}")
            messagebox.showinfo(
                "Auditoría Masiva SAP Concluida",
                f"Se auditaron exitosamente {len(resultados):,} filas.\n\n"
                f"Archivo con columnas de auditoría y colores generado:\n{out_cand}\n\n"
                f"Puedes abrirlo directamente con el botón verde 'Abrir Excel _VALIDADO.xlsx'."
            )
        else:
            self._val_matriz_log("Auditoría finalizada.")

    def _abrir_excel_matriz_validado(self):
        if self.val_matriz_out_path and os.path.exists(self.val_matriz_out_path):
            try:
                os.startfile(self.val_matriz_out_path)
            except Exception as e:
                messagebox.showerror("Error al abrir archivo", f"No se pudo abrir el archivo:\n{e}")
        else:
            messagebox.showwarning("Archivo no encontrado", "No se encontró el archivo generado _VALIDADO.xlsx.")

    # ------------------------------------------------------------------
    # Pestaña: Revisión de Materiales en SAP GUI (MM03)
    # ------------------------------------------------------------------
    def _build_tab_val_material(self):
        tab = self.val_material_view
        tab.grid_columnconfigure(0, weight=1)
        tab.grid_rowconfigure(0, weight=1)

        scroll = ctk.CTkScrollableFrame(tab, fg_color=CEMEX_GRAY_BG)
        scroll.grid(row=0, column=0, sticky="nsew")
        scroll.grid_columnconfigure(0, weight=1)

        # Variables de Formulario Individual
        self.val_mat_material_var = StringVar(value="")
        self.val_mat_centro_var = StringVar(value="")
        self.val_mat_org_var = StringVar(value="")
        self.val_mat_canal_var = StringVar(value="00")
        self.val_mat_almacen_var = StringVar(value="1002")
        self.val_mat_peso_esp_var = StringVar(value="")
        self.val_mat_clase_esp_var = StringVar(value="")
        self.val_mat_speed_var = StringVar(value="👁️ Humano (0.45s)")

        # --------------------------------------------------------------
        # 1. Card: Validación Individual de Material (MM03)
        # --------------------------------------------------------------
        card_ind = self._create_card(scroll, "1. Validación Individual de Material en SAP GUI (MM03)")
        card_ind.grid(row=0, column=0, padx=15, pady=(15, 10), sticky="nsew")
        card_ind.grid_columnconfigure(0, weight=1)

        frame_form = ctk.CTkFrame(card_ind, fg_color="transparent")
        frame_form.grid(row=1, column=0, padx=18, pady=(4, 10), sticky="ew")
        for col_i in range(6):
            frame_form.grid_columnconfigure(col_i, weight=1)

        # Fila 1: Material, Centro, Org. Ventas, Canal, Almacén, Velocidad
        ctk.CTkLabel(frame_form, text="Material SAP *:", font=CTkFont("Segoe UI", 11, "bold"), text_color=CEMEX_TEXT).grid(row=0, column=0, padx=6, pady=(4, 2), sticky="w")
        ctk.CTkEntry(frame_form, textvariable=self.val_mat_material_var, placeholder_text="Ej. 10147586", font=CTkFont("Segoe UI", 11), fg_color=CEMEX_WHITE, border_color=CEMEX_LIGHT_GRAY).grid(row=1, column=0, padx=6, pady=(0, 8), sticky="ew")

        ctk.CTkLabel(frame_form, text="Centro *:", font=CTkFont("Segoe UI", 11, "bold"), text_color=CEMEX_TEXT).grid(row=0, column=1, padx=6, pady=(4, 2), sticky="w")
        ctk.CTkEntry(frame_form, textvariable=self.val_mat_centro_var, placeholder_text="Ej. D825", font=CTkFont("Segoe UI", 11), fg_color=CEMEX_WHITE, border_color=CEMEX_LIGHT_GRAY).grid(row=1, column=1, padx=6, pady=(0, 8), sticky="ew")

        ctk.CTkLabel(frame_form, text="Org. Ventas *:", font=CTkFont("Segoe UI", 11, "bold"), text_color=CEMEX_TEXT).grid(row=0, column=2, padx=6, pady=(4, 2), sticky="w")
        ctk.CTkEntry(frame_form, textvariable=self.val_mat_org_var, placeholder_text="Ej. 7277", font=CTkFont("Segoe UI", 11), fg_color=CEMEX_WHITE, border_color=CEMEX_LIGHT_GRAY).grid(row=1, column=2, padx=6, pady=(0, 8), sticky="ew")

        ctk.CTkLabel(frame_form, text="Canal Distribución:", font=CTkFont("Segoe UI", 11, "bold"), text_color=CEMEX_TEXT).grid(row=0, column=3, padx=6, pady=(4, 2), sticky="w")
        ctk.CTkEntry(frame_form, textvariable=self.val_mat_canal_var, font=CTkFont("Segoe UI", 11), fg_color=CEMEX_WHITE, border_color=CEMEX_LIGHT_GRAY).grid(row=1, column=3, padx=6, pady=(0, 8), sticky="ew")

        ctk.CTkLabel(frame_form, text="Almacén Contab.:", font=CTkFont("Segoe UI", 11, "bold"), text_color=CEMEX_TEXT).grid(row=0, column=4, padx=6, pady=(4, 2), sticky="w")
        ctk.CTkEntry(frame_form, textvariable=self.val_mat_almacen_var, font=CTkFont("Segoe UI", 11), fg_color=CEMEX_WHITE, border_color=CEMEX_LIGHT_GRAY).grid(row=1, column=4, padx=6, pady=(0, 8), sticky="ew")

        ctk.CTkLabel(frame_form, text="Velocidad SAP GUI:", font=CTkFont("Segoe UI", 11, "bold"), text_color=CEMEX_TEXT).grid(row=0, column=5, padx=6, pady=(4, 2), sticky="w")
        ctk.CTkOptionMenu(
            frame_form,
            variable=self.val_mat_speed_var,
            values=["👁️ Humano (0.45s)", "⚡ Rápido (0.15s)", "⏸️ Paso a Paso"],
            fg_color=CEMEX_BLUE,
            button_color=CEMEX_BLUE_DARK,
            font=CTkFont("Segoe UI", 11),
        ).grid(row=1, column=5, padx=6, pady=(0, 8), sticky="ew")

        # Fila 2: Parámetros opcionales y botones de acción
        frame_act = ctk.CTkFrame(card_ind, fg_color="transparent")
        frame_act.grid(row=2, column=0, padx=18, pady=(0, 12), sticky="ew")
        frame_act.grid_columnconfigure(0, weight=1)
        frame_act.grid_columnconfigure(1, weight=1)
        frame_act.grid_columnconfigure(2, weight=0)
        frame_act.grid_columnconfigure(3, weight=0)
        frame_act.grid_columnconfigure(4, weight=0)

        ctk.CTkEntry(frame_act, textvariable=self.val_mat_peso_esp_var, placeholder_text="Peso Vol. Esperado (Opcional)", font=CTkFont("Segoe UI", 11), fg_color=CEMEX_WHITE, border_color=CEMEX_LIGHT_GRAY).grid(row=0, column=0, padx=6, pady=2, sticky="ew")
        ctk.CTkEntry(frame_act, textvariable=self.val_mat_clase_esp_var, placeholder_text="Clase/Clasif. Esperada (Opcional)", font=CTkFont("Segoe UI", 11), fg_color=CEMEX_WHITE, border_color=CEMEX_LIGHT_GRAY).grid(row=0, column=1, padx=6, pady=2, sticky="ew")

        self.val_mat_btn_validar_ind = ctk.CTkButton(
            frame_act,
            text="▶ Validar Material en SAP GUI",
            command=self._iniciar_val_material_individual,
            fg_color=CEMEX_BLUE,
            hover_color=CEMEX_BLUE_DARK,
            font=CTkFont("Segoe UI", 12, "bold"),
            height=34,
            width=230,
        )
        self.val_mat_btn_validar_ind.grid(row=0, column=2, padx=6, pady=2)

        self.val_mat_btn_detener_ind = ctk.CTkButton(
            frame_act,
            text="⏹ Detener",
            command=self._detener_val_material,
            fg_color=CEMEX_RED,
            hover_color=CEMEX_RED_DARK,
            font=CTkFont("Segoe UI", 12, "bold"),
            height=34,
            width=90,
            state="disabled",
        )
        self.val_mat_btn_detener_ind.grid(row=0, column=3, padx=6, pady=2)

        ctk.CTkButton(
            frame_act,
            text="🧹 Limpiar",
            command=self._limpiar_formulario_val_material,
            fg_color=CEMEX_NEUTRAL,
            hover_color=CEMEX_NEUTRAL_DARK,
            font=CTkFont("Segoe UI", 11, "bold"),
            height=34,
            width=80,
        ).grid(row=0, column=4, padx=6, pady=2)

        # Panel de Resultados Visuales Individuales (Banner + 6 Badges + Detalle)
        frame_res_ind = ctk.CTkFrame(card_ind, fg_color="#F8FAFC", corner_radius=10, border_width=1, border_color=CEMEX_LIGHT_GRAY)
        frame_res_ind.grid(row=3, column=0, padx=18, pady=(0, 14), sticky="ew")
        frame_res_ind.grid_columnconfigure(0, weight=1)

        # Banner de Estado Global
        self.val_mat_lbl_banner = ctk.CTkLabel(
            frame_res_ind,
            text="ESTADO: PENDIENTE · Ingrese Material, Centro y Org. Ventas, luego presione 'Validar Material'",
            font=CTkFont("Segoe UI", 12, "bold"),
            fg_color="#ECEFF1",
            text_color=CEMEX_TEXT,
            corner_radius=6,
            height=32,
        )
        self.val_mat_lbl_banner.grid(row=0, column=0, padx=12, pady=(10, 8), sticky="ew")

        # Grid de 6 Tarjetas/Badges de Validación
        badges_grid = ctk.CTkFrame(frame_res_ind, fg_color="transparent")
        badges_grid.grid(row=1, column=0, padx=12, pady=(0, 8), sticky="ew")
        for i in range(6):
            badges_grid.grid_columnconfigure(i, weight=1)

        # Badge 1: Stock Negativo
        b1 = ctk.CTkFrame(badges_grid, fg_color=CEMEX_WHITE, corner_radius=6, border_width=1, border_color=CEMEX_LIGHT_GRAY)
        b1.grid(row=0, column=0, padx=4, pady=4, sticky="nsew")
        ctk.CTkLabel(b1, text="Stock Negativo", font=CTkFont("Segoe UI", 10, "bold"), text_color=CEMEX_TEXT).pack(pady=(4, 1))
        self.val_mat_badge_stock = ctk.CTkLabel(b1, text="[ PENDIENTE ]", font=CTkFont("Segoe UI", 11, "bold"), text_color=CEMEX_NEUTRAL)
        self.val_mat_badge_stock.pack(pady=(0, 4))

        # Badge 2: Verif. Disponibilidad
        b2 = ctk.CTkFrame(badges_grid, fg_color=CEMEX_WHITE, corner_radius=6, border_width=1, border_color=CEMEX_LIGHT_GRAY)
        b2.grid(row=0, column=1, padx=4, pady=4, sticky="nsew")
        ctk.CTkLabel(b2, text="Verif. Disponibilidad", font=CTkFont("Segoe UI", 10, "bold"), text_color=CEMEX_TEXT).pack(pady=(4, 1))
        self.val_mat_badge_disp = ctk.CTkLabel(b2, text="[ PENDIENTE ]", font=CTkFont("Segoe UI", 11, "bold"), text_color=CEMEX_NEUTRAL)
        self.val_mat_badge_disp.pack(pady=(0, 4))

        # Badge 3: Pedido Automático
        b3 = ctk.CTkFrame(badges_grid, fg_color=CEMEX_WHITE, corner_radius=6, border_width=1, border_color=CEMEX_LIGHT_GRAY)
        b3.grid(row=0, column=2, padx=4, pady=4, sticky="nsew")
        ctk.CTkLabel(b3, text="Pedido Automático", font=CTkFont("Segoe UI", 10, "bold"), text_color=CEMEX_TEXT).pack(pady=(4, 1))
        self.val_mat_badge_auto = ctk.CTkLabel(b3, text="[ PENDIENTE ]", font=CTkFont("Segoe UI", 11, "bold"), text_color=CEMEX_NEUTRAL)
        self.val_mat_badge_auto.pack(pady=(0, 4))

        # Badge 4: Control de Precios
        b4 = ctk.CTkFrame(badges_grid, fg_color=CEMEX_WHITE, corner_radius=6, border_width=1, border_color=CEMEX_LIGHT_GRAY)
        b4.grid(row=0, column=3, padx=4, pady=4, sticky="nsew")
        ctk.CTkLabel(b4, text="Control Precios (S/V)", font=CTkFont("Segoe UI", 10, "bold"), text_color=CEMEX_TEXT).pack(pady=(4, 1))
        self.val_mat_badge_precio = ctk.CTkLabel(b4, text="[ PENDIENTE ]", font=CTkFont("Segoe UI", 11, "bold"), text_color=CEMEX_NEUTRAL)
        self.val_mat_badge_precio.pack(pady=(0, 4))

        # Badge 5: Peso Volumétrico
        b5 = ctk.CTkFrame(badges_grid, fg_color=CEMEX_WHITE, corner_radius=6, border_width=1, border_color=CEMEX_LIGHT_GRAY)
        b5.grid(row=0, column=4, padx=4, pady=4, sticky="nsew")
        ctk.CTkLabel(b5, text="Peso Volumétrico", font=CTkFont("Segoe UI", 10, "bold"), text_color=CEMEX_TEXT).pack(pady=(4, 1))
        self.val_mat_badge_peso = ctk.CTkLabel(b5, text="[ PENDIENTE ]", font=CTkFont("Segoe UI", 11, "bold"), text_color=CEMEX_NEUTRAL)
        self.val_mat_badge_peso.pack(pady=(0, 4))

        # Badge 6: Número Clasificador
        b6 = ctk.CTkFrame(badges_grid, fg_color=CEMEX_WHITE, corner_radius=6, border_width=1, border_color=CEMEX_LIGHT_GRAY)
        b6.grid(row=0, column=5, padx=4, pady=4, sticky="nsew")
        ctk.CTkLabel(b6, text="Nº Clasificador", font=CTkFont("Segoe UI", 10, "bold"), text_color=CEMEX_TEXT).pack(pady=(4, 1))
        self.val_mat_badge_clase = ctk.CTkLabel(b6, text="[ PENDIENTE ]", font=CTkFont("Segoe UI", 11, "bold"), text_color=CEMEX_NEUTRAL)
        self.val_mat_badge_clase.pack(pady=(0, 4))

        # Texto de Detalle de Auditoría Individual
        self.val_mat_lbl_detalle = ctk.CTkLabel(
            frame_res_ind,
            text="",
            font=CTkFont("Consolas", 11),
            text_color=CEMEX_TEXT,
            anchor="w",
            justify="left",
        )
        self.val_mat_lbl_detalle.grid(row=2, column=0, padx=14, pady=(0, 10), sticky="w")

        # --------------------------------------------------------------
        # 2. Card: Carga y Validación Masiva por Lote (Excel)
        # --------------------------------------------------------------
        card_lote = self._create_card(scroll, "2. Carga y Auditoría Masiva por Lote (Archivo Excel)")
        card_lote.grid(row=1, column=0, padx=15, pady=(0, 10), sticky="nsew")
        card_lote.grid_columnconfigure(0, weight=1)

        drop_zone = ctk.CTkFrame(card_lote, fg_color="#F8FAFC", corner_radius=10, border_width=2, border_color=CEMEX_BLUE)
        drop_zone.grid(row=1, column=0, padx=18, pady=(6, 10), sticky="ew")
        drop_zone.grid_columnconfigure(0, weight=1)

        ctk.CTkLabel(
            drop_zone,
            text="📂 ARRASTRA AQUÍ ARCHIVO EXCEL (.XLSX / .XLSM) O HAZ CLIC EN EXAMINAR",
            font=CTkFont("Segoe UI", 12, "bold"),
            text_color=CEMEX_BLUE,
        ).grid(row=0, column=0, pady=(12, 2))

        ctk.CTkLabel(
            drop_zone,
            text="Columnas requeridas: Material, Centro, Org. Ventas (opcionales: Canal, Almacén, Peso Vol., Clase)",
            font=CTkFont("Segoe UI", 11),
            text_color=CEMEX_TEXT,
        ).grid(row=1, column=0, pady=(0, 8))

        btn_browse = ctk.CTkButton(
            drop_zone,
            text="Examinar Archivo...",
            command=self._val_material_examinar_archivo,
            fg_color=CEMEX_BLUE,
            hover_color=CEMEX_BLUE_DARK,
            font=CTkFont("Segoe UI", 11, "bold"),
            height=30,
            width=160,
        )
        btn_browse.grid(row=2, column=0, pady=(0, 12))

        try:
            drop_zone.drop_target_register("DND_Files")
            drop_zone.dnd_bind("<<Drop>>", self._on_val_material_drop)
        except Exception as e:
            print(f"Aviso registrando DND en val_material: {e}")

        # Barra de Control Lote
        bar_lote = ctk.CTkFrame(card_lote, fg_color="transparent")
        bar_lote.grid(row=2, column=0, padx=18, pady=(0, 12), sticky="ew")
        bar_lote.grid_columnconfigure(1, weight=1)

        self.val_mat_lbl_archivo = ctk.CTkLabel(
            bar_lote,
            text="Archivo: (Ningún archivo cargado)",
            font=CTkFont("Segoe UI", 11, "bold"),
            text_color=CEMEX_TEXT,
            anchor="w",
        )
        self.val_mat_lbl_archivo.grid(row=0, column=0, sticky="w")

        self.val_mat_lbl_conteo = ctk.CTkLabel(
            bar_lote,
            text="Total materiales: 0",
            font=CTkFont("Segoe UI", 11, "bold"),
            text_color=CEMEX_NEUTRAL,
            anchor="w",
        )
        self.val_mat_lbl_conteo.grid(row=0, column=1, padx=16, sticky="w")

        self.val_mat_btn_validar_lote = ctk.CTkButton(
            bar_lote,
            text="▶ Validar Lote en SAP GUI",
            command=self._iniciar_val_material_lote,
            fg_color=CEMEX_BLUE,
            hover_color=CEMEX_BLUE_DARK,
            font=CTkFont("Segoe UI", 12, "bold"),
            height=34,
            state="disabled",
        )
        self.val_mat_btn_validar_lote.grid(row=0, column=2, padx=(0, 8), sticky="e")

        self.val_mat_btn_detener_lote = ctk.CTkButton(
            bar_lote,
            text="⏹ Detener",
            command=self._detener_val_material,
            fg_color=CEMEX_RED,
            hover_color=CEMEX_RED_DARK,
            font=CTkFont("Segoe UI", 12, "bold"),
            height=34,
            width=90,
            state="disabled",
        )
        self.val_mat_btn_detener_lote.grid(row=0, column=3, padx=(0, 8), sticky="e")

        self.val_mat_btn_exportar = ctk.CTkButton(
            bar_lote,
            text="💾 Exportar a Excel",
            command=self._exportar_val_material_excel,
            fg_color="#2E7D32",
            hover_color="#1B5E20",
            font=CTkFont("Segoe UI", 12, "bold"),
            height=34,
            state="disabled",
        )
        self.val_mat_btn_exportar.grid(row=0, column=4, sticky="e")

        # --------------------------------------------------------------
        # 3. Card: Tabla de Monitoreo de Resultados en Tiempo Real
        # --------------------------------------------------------------
        card_tabla = self._create_card(scroll, "3. Monitoreo y Resultados de Auditoría en Tiempo Real")
        card_tabla.grid(row=2, column=0, padx=15, pady=(0, 10), sticky="nsew")
        card_tabla.grid_columnconfigure(0, weight=1)

        frame_tree = ctk.CTkFrame(card_tabla, fg_color=CEMEX_WHITE, corner_radius=0)
        frame_tree.grid(row=1, column=0, padx=18, pady=(6, 12), sticky="nsew")
        frame_tree.grid_columnconfigure(0, weight=1)
        frame_tree.grid_rowconfigure(0, weight=1)

        val_columns = (
            "ESTADO", "MATERIAL", "CENTRO", "ORG_VTA",
            "STOCK_NEG", "VERIF_DISP", "PED_AUTO", "CTRL_PREC",
            "PESO_VOL", "DETALLES"
        )
        self.val_mat_tree = ttk.Treeview(
            frame_tree,
            columns=val_columns,
            show="headings",
            height=9,
            style="Cemex.Treeview",
        )
        self.val_mat_tree.tag_configure("ok", background="#E8F5E9", foreground="#1B5E20")
        self.val_mat_tree.tag_configure("error", background="#FFEBEE", foreground="#B71C1C")
        self.val_mat_tree.tag_configure("running", background="#FFF9C4", foreground="#F57F17")
        self.val_mat_tree.tag_configure("even", background=CEMEX_WHITE, foreground=CEMEX_TEXT)
        self.val_mat_tree.tag_configure("odd", background=CEMEX_GRAY_BG, foreground=CEMEX_TEXT)

        self.val_mat_tree.heading("ESTADO", text="Estado SAP")
        self.val_mat_tree.column("ESTADO", width=110, anchor="center")

        self.val_mat_tree.heading("MATERIAL", text="Material")
        self.val_mat_tree.column("MATERIAL", width=95, anchor="center")

        self.val_mat_tree.heading("CENTRO", text="Centro")
        self.val_mat_tree.column("CENTRO", width=65, anchor="center")

        self.val_mat_tree.heading("ORG_VTA", text="Org. Vta")
        self.val_mat_tree.column("ORG_VTA", width=75, anchor="center")

        self.val_mat_tree.heading("STOCK_NEG", text="Stock Neg.")
        self.val_mat_tree.column("STOCK_NEG", width=85, anchor="center")

        self.val_mat_tree.heading("VERIF_DISP", text="Verif. Disp")
        self.val_mat_tree.column("VERIF_DISP", width=80, anchor="center")

        self.val_mat_tree.heading("PED_AUTO", text="Ped. Auto")
        self.val_mat_tree.column("PED_AUTO", width=80, anchor="center")

        self.val_mat_tree.heading("CTRL_PREC", text="Ctrl Prec")
        self.val_mat_tree.column("CTRL_PREC", width=75, anchor="center")

        self.val_mat_tree.heading("PESO_VOL", text="Peso Vol.")
        self.val_mat_tree.column("PESO_VOL", width=85, anchor="center")

        self.val_mat_tree.heading("DETALLES", text="Detalles de Auditoría")
        self.val_mat_tree.column("DETALLES", width=340, anchor="w")

        vsb = ttk.Scrollbar(frame_tree, orient="vertical", command=self.val_mat_tree.yview)
        vsb.grid(row=0, column=1, sticky="ns")
        hsb = ttk.Scrollbar(frame_tree, orient="horizontal", command=self.val_mat_tree.xview)
        hsb.grid(row=1, column=0, sticky="ew")
        self.val_mat_tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)
        self.val_mat_tree.grid(row=0, column=0, sticky="nsew")

        # --------------------------------------------------------------
        # 4. Card: Consola de Auditoría en Tiempo Real
        # --------------------------------------------------------------
        card_log = self._create_card(scroll, "4. Consola de Registro y Comunicación en Vivo con SAP (MM03)")
        card_log.grid(row=3, column=0, padx=15, pady=(0, 15), sticky="nsew")
        card_log.grid_columnconfigure(0, weight=1)

        self.val_mat_log_box = ctk.CTkTextbox(
            card_log,
            height=100,
            wrap="none",
            font=CTkFont("Consolas", 10),
            fg_color="#1E1E1E",
            text_color="#00FF66",
        )
        self.val_mat_log_box.grid(row=1, column=0, padx=18, pady=(6, 12), sticky="ew")
        self.val_mat_log_box.insert("1.0", "[LISTO] Módulo MM03 preparado. Ingrese material individual o cargue un archivo Excel.\n")

    # ------------------------------------------------------------------
    # Métodos de Lógica y Ejecución para Revisión de Materiales MM03
    # ------------------------------------------------------------------
    def _val_mat_log(self, texto):
        """Escribe un mensaje con marca de tiempo en la consola de MM03."""
        hora = datetime.now().strftime("%H:%M:%S")
        self.val_mat_log_box.insert("end", f"[{hora}] {texto}\n")
        self.val_mat_log_box.see("end")

    def _limpiar_formulario_val_material(self):
        """Limpia los campos del formulario individual y restablece los badges."""
        self.val_mat_material_var.set("")
        self.val_mat_centro_var.set("")
        self.val_mat_org_var.set("")
        self.val_mat_canal_var.set("00")
        self.val_mat_almacen_var.set("1002")
        self.val_mat_peso_esp_var.set("")
        self.val_mat_clase_esp_var.set("")

        self.val_mat_lbl_banner.configure(
            text="ESTADO: PENDIENTE · Ingrese Material, Centro y Org. Ventas, luego presione 'Validar Material'",
            fg_color="#ECEFF1",
            text_color=CEMEX_TEXT,
        )
        self.val_mat_badge_stock.configure(text="[ PENDIENTE ]", text_color=CEMEX_NEUTRAL)
        self.val_mat_badge_disp.configure(text="[ PENDIENTE ]", text_color=CEMEX_NEUTRAL)
        self.val_mat_badge_auto.configure(text="[ PENDIENTE ]", text_color=CEMEX_NEUTRAL)
        self.val_mat_badge_precio.configure(text="[ PENDIENTE ]", text_color=CEMEX_NEUTRAL)
        self.val_mat_badge_peso.configure(text="[ PENDIENTE ]", text_color=CEMEX_NEUTRAL)
        self.val_mat_badge_clase.configure(text="[ PENDIENTE ]", text_color=CEMEX_NEUTRAL)
        self.val_mat_lbl_detalle.configure(text="")

    def _iniciar_val_material_individual(self):
        mat = self.val_mat_material_var.get().strip()
        centro = self.val_mat_centro_var.get().strip()
        org_vta = self.val_mat_org_var.get().strip()
        canal = self.val_mat_canal_var.get().strip() or "00"
        almacen = self.val_mat_almacen_var.get().strip() or "1002"
        peso_esp = self.val_mat_peso_esp_var.get().strip() or None
        clase_esp = self.val_mat_clase_esp_var.get().strip() or None

        if not mat or not centro or not org_vta:
            messagebox.showwarning(
                "Campos obligatorios requeridos",
                "Por favor ingrese Material, Centro y Organización de Ventas para poder consultar en SAP."
            )
            return

        spd_choice = self.val_mat_speed_var.get()
        delay = 0.45
        if "0.15s" in spd_choice:
            delay = 0.15
        elif "Paso a Paso" in spd_choice:
            delay = 1.0

        self.val_mat_btn_validar_ind.configure(state="disabled")
        self.val_mat_btn_detener_ind.configure(state="normal")
        self.val_mat_lbl_banner.configure(
            text=f"🔄 VALIDANDO MATERIAL {mat} EN SAP GUI (MM03)...",
            fg_color="#FFF9C4",
            text_color="#F57F17",
        )

        self.val_mat_motor = sap_validator_engine.MotorValidadorSAP(
            speed_delay=delay,
            on_log=self._val_mat_log,
        )

        threading.Thread(
            target=self._run_val_material_individual_worker,
            args=(mat, centro, org_vta, canal, almacen, peso_esp, clase_esp),
            daemon=True,
        ).start()

    def _run_val_material_individual_worker(self, mat, centro, org_vta, canal, almacen, peso_esp, clase_esp):
        import pythoncom
        try:
            pythoncom.CoInitialize()
        except Exception:
            pass
        try:
            res = self.val_mat_motor.validar_material_mm03(
                mat, centro, org_vta, canal=canal, almacen=almacen, peso_esp=peso_esp, clase_esp=clase_esp
            )
            self.after(0, self._on_val_material_individual_finished, res, None)
        except Exception as e:
            self.after(0, self._on_val_material_individual_finished, None, e)
        finally:
            try:
                pythoncom.CoUninitialize()
            except Exception:
                pass

    def _on_val_material_individual_finished(self, res, error):
        self.val_mat_btn_validar_ind.configure(state="normal")
        self.val_mat_btn_detener_ind.configure(state="disabled")

        if error:
            self._val_mat_log(f"Error en validación: {error}")
            self.val_mat_lbl_banner.configure(
                text=f"⚠️ ERROR: {error}",
                fg_color="#FFEBEE",
                text_color="#B71C1C",
            )
            messagebox.showerror("Error en Validación SAP", f"Ocurrió un error al consultar SAP GUI:\n{error}")
            return

        estado = res.get("estado", "ERROR")
        de = res.get("datos_extraidos", {})
        chk = res.get("checks", {})

        # Actualizar Banner
        if estado == "OK":
            self.val_mat_lbl_banner.configure(
                text=f"✅ MATERIAL {res['material']} CONFORME · CUMPLE POLÍTICAS CORPORATIVAS CEMEX",
                fg_color="#E8F5E9",
                text_color="#1B5E20",
            )
        elif estado == "DISCREPANCIA":
            self.val_mat_lbl_banner.configure(
                text=f"❌ MATERIAL {res['material']} CON DISCREPANCIA · REVISAR PARÁMETROS OBSERVADOS",
                fg_color="#FFEBEE",
                text_color="#B71C1C",
            )
        else:
            self.val_mat_lbl_banner.configure(
                text=f"⚠️ {res.get('msg', 'Error en consulta')}",
                fg_color="#FFF3E0",
                text_color="#E65100",
            )

        # Actualizar Badges
        # 1. Stock Negativo
        if chk.get("stock_negativo_ok"):
            self.val_mat_badge_stock.configure(text="ACTIVO ✅", text_color="#1B5E20")
        else:
            self.val_mat_badge_stock.configure(text="INACTIVO ❌", text_color="#B71C1C")

        # 2. Verif. Disponibilidad (Informativo: muestra el valor tal cual esté en SAP: 02, KP, etc.)
        disp_txt = de.get("verif_disponibilidad") or "Vacío"
        self.val_mat_badge_disp.configure(
            text=f"{disp_txt}",
            text_color=CEMEX_BLUE if disp_txt != "Vacío" else CEMEX_NEUTRAL
        )

        # 3. Pedido Automático
        if chk.get("pedido_auto_ok"):
            self.val_mat_badge_auto.configure(text="ACTIVO ✅", text_color="#1B5E20")
        else:
            self.val_mat_badge_auto.configure(text="INACTIVO ❌", text_color="#B71C1C")

        # 4. Control de Precios
        pr_txt = de.get("control_precios") or "Vacío"
        if chk.get("control_precios_ok"):
            self.val_mat_badge_precio.configure(text=f"{pr_txt} (OK) ✅", text_color="#1B5E20")
        else:
            self.val_mat_badge_precio.configure(text=f"{pr_txt} (No S/V) ❌", text_color="#B71C1C")

        # 5. Peso Volumétrico
        peso_txt = de.get("peso_volumetrico") or "No def."
        self.val_mat_badge_peso.configure(text=str(peso_txt), text_color=CEMEX_BLUE if peso_txt != "No def." else CEMEX_NEUTRAL)

        # 6. Clasificación (Número clasificador directamente)
        clasif_num = de.get("clasificador") or "N/A"
        self.val_mat_badge_clase.configure(
            text=str(clasif_num),
            text_color=CEMEX_BLUE if clasif_num != "N/A" else CEMEX_NEUTRAL
        )

        # Detalle de texto
        detalles_str = "\n".join(f"• {d}" for d in res.get("detalles", []))
        self.val_mat_lbl_detalle.configure(text=detalles_str)

        # Agregar fila al Treeview para trazabilidad
        tag = "ok" if estado == "OK" else ("error" if estado == "DISCREPANCIA" else "odd")
        self.val_mat_tree.insert(
            "",
            "end",
            values=(
                f"✅ {estado}" if estado == "OK" else f"❌ {estado}",
                res.get("material", ""),
                res.get("centro", ""),
                res.get("org_ventas", ""),
                "ACTIVO" if de.get("stock_negativo") else "INACTIVO",
                de.get("verif_disponibilidad", ""),
                "ACTIVO" if de.get("pedido_automatico") else "INACTIVO",
                de.get("control_precios", ""),
                de.get("peso_volumetrico", ""),
                res.get("msg", ""),
            ),
            tags=(tag,),
        )
        self.val_mat_btn_exportar.configure(state="normal")
        self.val_mat_resultados.append(res)

    def _val_material_examinar_archivo(self):
        file_types = [("Archivos Excel", "*.xlsx;*.xlsm"), ("Todos los archivos", "*.*")]
        path = filedialog.askopenfilename(title="Seleccionar Archivo de Materiales para MM03", filetypes=file_types)
        if path:
            self._cargar_archivo_val_material(path)

    def _on_val_material_drop(self, event):
        raw = event.data.strip()
        if raw.startswith("{") and raw.endswith("}"):
            raw = raw[1:-1]
        if os.path.isfile(raw) and raw.lower().endswith((".xlsx", ".xlsm")):
            self._cargar_archivo_val_material(raw)
        else:
            messagebox.showwarning("Archivo no válido", "Por favor arrastra un archivo de Excel válido (.xlsx / .xlsm).")

    def _cargar_archivo_val_material(self, path):
        try:
            self.val_mat_file_path = path
            items = sap_validator_engine.extraer_filas_materiales_excel(path)
            self.val_mat_items = items
            self.val_mat_resultados = []

            for item in self.val_mat_tree.get_children():
                self.val_mat_tree.delete(item)

            nom = os.path.basename(path)
            self.val_mat_lbl_archivo.configure(text=f"Archivo: {nom}")
            self.val_mat_lbl_conteo.configure(text=f"Total materiales detectados: {len(items)}")

            for idx, it in enumerate(items):
                iid = str(idx)
                tag = "even" if idx % 2 == 0 else "odd"
                self.val_mat_tree.insert(
                    "",
                    "end",
                    iid=iid,
                    values=(
                        "⏳ Pendiente",
                        it.get("material", ""),
                        it.get("centro", ""),
                        it.get("org_ventas", ""),
                        "-", "-", "-", "-", "-",
                        "Listo para consultar en SAP",
                    ),
                    tags=(tag,),
                )

            self.val_mat_btn_validar_lote.configure(state="normal" if items else "disabled")
            self.val_mat_btn_exportar.configure(state="disabled")
            self._val_mat_log(f"Cargado archivo '{nom}' con {len(items)} materiales para auditar.")
        except Exception as e:
            messagebox.showerror("Error al leer archivo", f"No se pudo procesar el archivo Excel:\n{e}")

    def _iniciar_val_material_lote(self):
        if not self.val_mat_items:
            return

        spd_choice = self.val_mat_speed_var.get()
        delay = 0.45
        if "0.15s" in spd_choice:
            delay = 0.15
        elif "Paso a Paso" in spd_choice:
            delay = 1.0

        self.val_mat_btn_validar_lote.configure(state="disabled")
        self.val_mat_btn_detener_lote.configure(state="normal")
        self.val_mat_btn_validar_ind.configure(state="disabled")

        self.val_mat_motor = sap_validator_engine.MotorValidadorSAP(
            speed_delay=delay,
            on_progress=self._on_val_material_progress_bridge,
            on_log=self._val_mat_log,
        )

        threading.Thread(target=self._run_val_material_lote_worker, daemon=True).start()

    def _run_val_material_lote_worker(self):
        import pythoncom
        try:
            pythoncom.CoInitialize()
        except Exception:
            pass
        try:
            resultados = self.val_mat_motor.validar_lote_materiales(self.val_mat_items, ruta_excel=self.val_mat_file_path)
            self.after(0, self._on_val_material_lote_finished, resultados, None)
        except Exception as e:
            self.after(0, self._on_val_material_lote_finished, [], e)
        finally:
            try:
                pythoncom.CoUninitialize()
            except Exception:
                pass

    def _on_val_material_progress_bridge(self, idx, total, item, estado, datos_sap, msg):
        self.after(0, self._update_val_material_tree_row, idx, total, item, estado, datos_sap, msg)

    def _update_val_material_tree_row(self, idx, total, item, estado, datos_sap, msg):
        iid = str(idx - 1)
        if not self.val_mat_tree.exists(iid):
            return

        if estado == "EN_PROCESO":
            estado_texto = "🔄 Validando..."
            tag = "running"
            st_neg = "-"
            v_disp = "-"
            p_aut = "-"
            c_prc = "-"
            p_vol = "-"
        elif estado == "OK":
            estado_texto = "✅ OK"
            tag = "ok"
            st_neg = "ACTIVO" if datos_sap.get("stock_negativo") else "INACTIVO"
            v_disp = str(datos_sap.get("verif_disponibilidad") or "")
            p_aut = "ACTIVO" if datos_sap.get("pedido_automatico") else "INACTIVO"
            c_prc = str(datos_sap.get("control_precios") or "")
            p_vol = str(datos_sap.get("peso_volumetrico") or "")
        elif estado == "DISCREPANCIA":
            estado_texto = "❌ Discrepancia"
            tag = "error"
            st_neg = "ACTIVO" if datos_sap.get("stock_negativo") else "INACTIVO"
            v_disp = str(datos_sap.get("verif_disponibilidad") or "")
            p_aut = "ACTIVO" if datos_sap.get("pedido_automatico") else "INACTIVO"
            c_prc = str(datos_sap.get("control_precios") or "")
            p_vol = str(datos_sap.get("peso_volumetrico") or "")
        else:
            estado_texto = f"⚠️ {estado}"
            tag = "error"
            st_neg = "-"
            v_disp = "-"
            p_aut = "-"
            c_prc = "-"
            p_vol = "-"

        self.val_mat_tree.item(
            iid,
            values=(
                estado_texto,
                item.get("material", ""),
                item.get("centro", ""),
                item.get("org_ventas", ""),
                st_neg,
                v_disp,
                p_aut,
                c_prc,
                p_vol,
                msg,
            ),
            tags=(tag,),
        )
        self.val_mat_tree.see(iid)

    def _on_val_material_lote_finished(self, resultados, error):
        self.val_mat_btn_validar_lote.configure(state="normal")
        self.val_mat_btn_detener_lote.configure(state="disabled")
        self.val_mat_btn_validar_ind.configure(state="normal")
        self.val_mat_resultados = resultados

        if error:
            self._val_mat_log(f"Error en validación por lote: {error}")
            messagebox.showerror("Error en Validación SAP", f"Ocurrió un error en la validación por lote:\n{error}")
            return

        if resultados:
            self.val_mat_btn_exportar.configure(state="normal")
            oks = sum(1 for r in resultados if r.get("estado") == "OK")
            discs = sum(1 for r in resultados if r.get("estado") == "DISCREPANCIA")
            errs = len(resultados) - oks - discs
            self._val_mat_log(f"Auditoría MM03 completada: {oks} Conformes, {discs} Discrepancias, {errs} Errores.")
            messagebox.showinfo(
                "Auditoría MM03 Finalizada",
                f"Validación de materiales finalizada exitosamente.\n\n"
                f"• Conformes (OK): {oks}\n"
                f"• Con Discrepancia: {discs}\n"
                f"• Errores / Inexistentes: {errs}\n\n"
                f"Puede exportar el reporte detallado con el botón 'Exportar a Excel'."
            )

    def _detener_val_material(self):
        if self.val_mat_motor:
            self.val_mat_motor.stop_requested = True
            self._val_mat_log("Se solicitó detener la validación...")

    def _exportar_val_material_excel(self):
        if not self.val_mat_resultados:
            messagebox.showwarning("Sin resultados", "No hay resultados para exportar.")
            return

        def_name = f"Auditoria_Materiales_MM03_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
        salida = filedialog.asksaveasfilename(
            title="Guardar Reporte de Auditoría de Materiales",
            defaultextension=".xlsx",
            initialfile=def_name,
            filetypes=[("Archivos Excel", "*.xlsx")],
        )
        if salida:
            try:
                sap_validator_engine.exportar_resultados_materiales_excel(salida, self.val_mat_resultados)
                messagebox.showinfo("Exportación Exitosa", f"Reporte guardado exitosamente en:\n{salida}")
            except Exception as e:
                messagebox.showerror("Error al guardar", f"No se pudo guardar el archivo Excel:\n{e}")

    def _fpa_paste(self, event=None):
        """Pega filas FPA desde el portapapeles en el Treeview, sin tocar la
        Mesa de Captura."""
        try:
            raw = self.clipboard_get()
        except Exception:
            return "break"

        rows = fpa_validations.parse_clipboard_fpa(raw)
        if not rows:
            self._fpa_show_message(
                "No se detectaron filas tabuladas. Copia un rango de Excel con las 24 columnas FPA."
            )
            return "break"

        self.fpa_rows = rows
        self.fpa_selected_index = None
        self.fpa_selected_request = None
        self._refresh_fpa_tree()

        # Auto-seleccionar la primera fila para que F6/F7/F8/F9 funcionen de inmediato
        items = self.fpa_tree.get_children()
        if items:
            self.fpa_tree.selection_set(items[0])
            self.fpa_tree.focus(items[0])
            self._on_fpa_select()

        # Reflejar el texto pegado en la zona de pegado (solo como confirmación visual)
        try:
            self.fpa_paste_box.delete("1.0", "end")
            self.fpa_paste_box.insert("1.0", raw)
        except Exception:
            pass

        self._fpa_show_message(
            f"Se importaron {len(rows)} filas FPA desde Excel.\n"
            "Selecciona una fila en la tabla inferior y usa F6/F7/F8/F9."
        )
        return "break"

    def _refresh_fpa_tree(self):
        for item in self.fpa_tree.get_children():
            self.fpa_tree.delete(item)
        for i, row in enumerate(self.fpa_rows):
            values = tuple(row.get(c, "") for c in fpa_validations.FPA_COLUMNS)
            tag = "even" if i % 2 == 0 else "odd"
            self.fpa_tree.insert("", "end", iid=str(i), values=values, tags=(tag,))
        self.auto_ajustar_columnas(self.fpa_tree, self.fpa_rows)

    def _fpa_sync_rows_from_tree(self):
        """Reconstruye `self.fpa_rows` a partir de las filas que quedan en el
        Treeview (los iid son índices secuenciales en `self.fpa_rows`), y
        reindexa el árbol. Reutilizado por el borrado y la validación."""
        remaining = []
        for iid in self.fpa_tree.get_children():
            try:
                remaining.append(self.fpa_rows[int(iid)])
            except (ValueError, IndexError):
                continue
        self.fpa_rows = remaining
        self._refresh_fpa_tree()

    def _fpa_delete_selected(self, event=None):
        """Elimina visualmente las filas seleccionadas del Treeview FPA al
        presionar la tecla Suprimir, manteniendo sincronizada la lista interna
        `self.fpa_rows` que alimenta a F6/F7/F8/F9."""
        seleccionadas = list(self.fpa_tree.selection())
        if not seleccionadas:
            return None

        # Eliminar primero del Treeview
        for iid in seleccionadas:
            self.fpa_tree.delete(iid)

        # Reconstruir self.fpa_rows a partir de las filas que quedan en el árbol
        self._fpa_sync_rows_from_tree()

        # Limpiar selección actual
        self.fpa_selected_index = None
        self.fpa_selected_request = None

        self._fpa_show_message(
            f"Se eliminaron {len(seleccionadas)} fila(s). Quedan {len(self.fpa_rows)} fila(s)."
        )
        return None

    def _on_fpa_select(self, event=None):
        selected = self.fpa_tree.selection()
        if not selected:
            self.fpa_selected_index = None
            self.fpa_selected_request = None
            return
        iid = selected[0]
        try:
            self.fpa_selected_index = int(iid)
            self.fpa_selected_request = self.fpa_rows[self.fpa_selected_index]
        except (ValueError, IndexError):
            self.fpa_selected_index = None
            self.fpa_selected_request = None

    def _fpa_current_row(self):
        if self.fpa_selected_request is not None:
            return self.fpa_selected_request
        selected = self.fpa_tree.selection()
        if not selected:
            return None
        try:
            return self.fpa_rows[int(selected[0])]
        except (ValueError, IndexError):
            return None

    def _fpa_show_message(self, text):
        self.fpa_result_text.configure(state="normal")
        self.fpa_result_text.delete("1.0", "end")
        self.fpa_result_text.insert("end", text)
        self.fpa_result_text.configure(state="disabled")

    def fpa_buscar(self):
        row = self._fpa_current_row()
        if row is None:
            messagebox.showwarning("FPA", "Seleccione una fila en el área de pegado.")
            return

        fecha_proc = normalizar_fecha(self.fpa_fecha_proc_var.get())

        # Búsqueda en SQLite
        db_matches = fpa_validations.buscar_en_sqlite(self.conn, row, fecha_proc)
        self.fpa_db_matches = db_matches

        # Búsqueda en Excel (en hilo para no bloquear la GUI)
        self.fpa_buscar_btn.configure(state="disabled", text="Buscando...")

        def _run():
            try:
                excel_matches = fpa_validations.buscar_en_excel(row, fecha_proc)
                self.after(0, self._on_fpa_buscar_done, db_matches, excel_matches, None)
            except Exception as e:
                self.after(0, self._on_fpa_buscar_done, db_matches, [], e)

        threading.Thread(target=_run, daemon=True).start()

    def _on_fpa_buscar_done(self, db_matches, excel_matches, error):
        self.fpa_buscar_btn.configure(state="normal", text="Buscar fila (F6)")
        if error:
            messagebox.showerror("Error de búsqueda", str(error))
            return
        self.fpa_excel_matches = excel_matches
        fecha_proc = normalizar_fecha(self.fpa_fecha_proc_var.get())
        header = f"Fecha de procesamiento: {fecha_proc or 'sin filtro'}"
        lines = [header, f"SQLite: {len(db_matches)} coincidencia(s)."]
        for m in db_matches:
            lines.append(
                f"  id={m.get('id')} | {m.get('cedis')} | {m.get('sf')} | "
                f"{m.get('articulo')} | {m.get('modalidad')} | {m.get('fecha_procesamiento')}"
            )
        lines.append(f"Excel: {len(excel_matches)} coincidencia(s).")
        for m in excel_matches:
            lines.append(
                f"  fila={m.get('_excel_row')} | {m.get('Centro')} | {m.get('ShipFrom')} | "
                f"{m.get('Material')} | {m.get('Modalidad')} | {m.get('FechaProcesamiento')}"
            )
        self._fpa_show_message("\n".join(lines))

    def fpa_buscar_solicitud(self):
        """Busca en SQLite/Excel cada fila de self.fpa_rows y muestra un
        conteo sumario total en el área de resultados."""
        if not self.fpa_rows:
            messagebox.showwarning("FPA", "No hay filas pegadas para buscar.")
            return

        fecha_proc = normalizar_fecha(self.fpa_fecha_proc_var.get())
        self.fpa_buscar_solicitud_btn.configure(state="disabled", text="Buscando...")

        def _run():
            total_db = 0
            total_excel = 0
            errores = 0
            lineas = [f"Búsqueda masiva sobre {len(self.fpa_rows)} fila(s):", ""]
            for i, row in enumerate(self.fpa_rows):
                try:
                    db_matches = fpa_validations.buscar_en_sqlite(self.conn, row, fecha_proc)
                    total_db += len(db_matches)
                except Exception:
                    db_matches = []
                    errores += 1
                try:
                    excel_matches = fpa_validations.buscar_en_excel(row, fecha_proc)
                    total_excel += len(excel_matches)
                except Exception:
                    excel_matches = []
                    errores += 1
                lineas.append(
                    f"Fila {i + 1}: SQLite={len(db_matches)} | Excel={len(excel_matches)}"
                )
            lineas.append("")
            lineas.append(f"TOTAL SQLite: {total_db}")
            lineas.append(f"TOTAL Excel: {total_excel}")
            if errores:
                lineas.append(f"Filas con error: {errores}")
            self.after(0, self._on_fpa_buscar_solicitud_done, "\n".join(lineas))

        threading.Thread(target=_run, daemon=True).start()

    def _on_fpa_buscar_solicitud_done(self, texto):
        self.fpa_buscar_solicitud_btn.configure(state="normal", text="Buscar solicitud")
        self._fpa_show_message(texto)

    def fpa_validar_solicitud(self):
        """Abre un modal pidiendo una única Fecha de Confirmación y evidencias,
        y al finalizar procesa TODAS las filas de self.fpa_rows en bucle."""
        if not self.fpa_rows:
            messagebox.showwarning("FPA", "No hay filas pegadas para validar.")
            return
        self._abrir_modal_validar_solicitud()

    def _abrir_modal_validar_solicitud(self):
        top = ctk.CTkToplevel(self)
        top.title("Validar Solicitud (todas las filas)")
        top.geometry("460x420")
        top.configure(fg_color=CEMEX_GRAY_BG)
        top.grab_set()
        top.resizable(True, True)

        imagenes = []

        container = ctk.CTkScrollableFrame(top, fg_color="transparent")
        container.pack(fill="both", expand=True, padx=8, pady=8)
        container.grid_columnconfigure(0, weight=1)
        container.grid_columnconfigure(1, weight=1)

        ctk.CTkLabel(
            container,
            text=f"Se validarán {len(self.fpa_rows)} fila(s) de la solicitud.",
            font=CTkFont("Segoe UI", 15, "bold"), text_color=CEMEX_BLUE,
        ).grid(row=0, column=0, columnspan=2, padx=20, pady=(12, 8), sticky="w")

        ctk.CTkLabel(
            container, text="Evidencias FPA (pega imágenes desde SAP con Ctrl+V)",
            font=CTkFont("Segoe UI", 14, "bold"), text_color=CEMEX_BLUE,
        ).grid(row=1, column=0, columnspan=2, padx=20, pady=(8, 4), sticky="w")

        status = ctk.CTkLabel(
            container, text="Pega una imagen con Ctrl+V.",
            font=CTkFont("Segoe UI", 12), text_color=CEMEX_TEXT,
        )
        status.grid(row=2, column=0, columnspan=2, padx=20, pady=4, sticky="w")

        fecha_var = StringVar(value=datetime.now().strftime("%d/%m/%Y"))
        ctk.CTkLabel(
            container, text="Fecha de Confirmación:",
            font=CTkFont("Segoe UI", 12, "bold"), text_color=CEMEX_TEXT,
        ).grid(row=3, column=0, padx=20, pady=(8, 4), sticky="w")
        fecha_entry = ctk.CTkEntry(
            container, textvariable=fecha_var,
            font=CTkFont("Segoe UI", 12), width=160,
            fg_color=CEMEX_WHITE, border_color=CEMEX_LIGHT_GRAY,
            text_color=CEMEX_TEXT,
        )
        fecha_entry.grid(row=3, column=1, padx=(8, 20), pady=(8, 4), sticky="w")

        def _refresh_preview():
            fecha = fecha_var.get().strip() or datetime.now().strftime("%d/%m/%Y")
            partes = [f"Ok {fecha}"]
            for i in range(1, len(imagenes) + 1):
                partes.append(f"[Imagen {i}]")
            preview.configure(state="normal")
            preview.delete("1.0", "end")
            preview.insert("end", "\n".join(partes))
            preview.configure(state="disabled")

        ctk.CTkLabel(
            container, text="Vista previa del mensaje para Outlook:",
            font=CTkFont("Segoe UI", 12, "bold"), text_color=CEMEX_TEXT,
        ).grid(row=4, column=0, columnspan=2, padx=20, pady=(6, 4), sticky="w")
        preview = ctk.CTkTextbox(
            container, height=100, wrap="word",
            font=CTkFont("Consolas", 11),
            fg_color="#FAFAFA", text_color=CEMEX_TEXT,
            border_width=1, border_color=CEMEX_LIGHT_GRAY,
            state="disabled",
        )
        preview.grid(row=5, column=0, columnspan=2, padx=20, pady=(0, 8), sticky="ew")

        _refresh_preview()
        fecha_var.trace_add("write", lambda *_: _refresh_preview())

        def on_paste(event=None):
            img = fpa_validations.capturar_imagen_portapapeles()
            if img is None:
                status.configure(text="No se encontró imagen en el portapapeles.")
                return "break"
            imagenes.append(img)
            status.configure(text=f"Imagen {len(imagenes)} capturada.")
            _refresh_preview()
            return "break"

        top.bind("<Control-v>", on_paste)

        def on_finalizar():
            fecha = fecha_var.get().strip() or datetime.now().strftime("%d/%m/%Y")
            self._fpa_ejecutar_validacion_solicitud(top, fecha, imagenes)

        ctk.CTkButton(
            container, text="Finalizar", command=on_finalizar,
            fg_color=CEMEX_BLUE, hover_color=CEMEX_BLUE_DARK,
            font=CTkFont("Segoe UI", 13, "bold"), height=40,
        ).grid(row=6, column=0, columnspan=2, padx=20, pady=(8, 4), sticky="ew")

        ctk.CTkButton(
            container, text="Cerrar", command=top.destroy,
            fg_color=CEMEX_NEUTRAL, hover_color=CEMEX_NEUTRAL_DARK,
            font=CTkFont("Segoe UI", 12, "bold"), height=36,
        ).grid(row=7, column=0, columnspan=2, padx=20, pady=(4, 16), sticky="ew")

    def _fpa_ejecutar_validacion_solicitud(self, top, fecha, imagenes):
        """Actualiza SQLite y Excel en bucle para TODAS las filas de
        self.fpa_rows, genera el portapapeles de Outlook, limpia la tabla y
        muestra un resumen global."""
        fecha_str = fecha or datetime.now().strftime("%d/%m/%Y")
        filas = list(self.fpa_rows)
        total = len(filas)
        if total == 0:
            messagebox.showwarning("FPA", "No hay filas para validar.", parent=top)
            return

        ok = 0
        errores = 0

        for row in filas:
            try:
                fpa_validations.actualizar_fecha_confirmacion_sqlite(self.conn, [row], fecha_str)
                fpa_validations.actualizar_fecha_confirmacion_excel([row], fecha_str)
                ok += 1
            except Exception:
                errores += 1

        ok_clip = fpa_validations.preparar_portapapeles_outlook(f"Ok {fecha_str}", imagenes)

        # Limpiar la solicitud completa
        self._fpa_clear_rows()

        resumen = (
            f"Validación de solicitud finalizada.\n"
            f"Filas procesadas: {ok}\n"
            f"Filas con error: {errores}\n"
            f"Total: {total}\n"
            f"Portapapeles: {'OK' if ok_clip else 'Error'}"
        )
        self._fpa_show_message(resumen)
        messagebox.showinfo("Validar Solicitud", resumen, parent=top)
        top.destroy()

    def fpa_migrar_excel(self):
        """Sincroniza el Excel Maestro con SQLite (inserta nuevas, actualiza existentes)."""
        self.fpa_migrar_btn.configure(state="disabled", text="Migrando...")

        def _run():
            try:
                insertadas, actualizadas, error = fpa_validations.migrar_excel_a_sqlite(self.conn)
                self.after(0, self._on_fpa_migrar_done, insertadas, actualizadas, error)
            except Exception as e:
                self.after(0, self._on_fpa_migrar_done, 0, 0, str(e))

        threading.Thread(target=_run, daemon=True).start()

    def _on_fpa_migrar_done(self, insertadas, actualizadas, error):
        self.fpa_migrar_btn.configure(state="normal", text="Migrar Excel → SQLite")
        if error:
            messagebox.showerror("Error de migración", error)
            return
        self._refresh_sla_tree()
        self._fpa_show_message(
            f"Migración completada.\nInsertadas: {insertadas}\nActualizadas: {actualizadas}"
        )
        messagebox.showinfo(
            "Migración Excel → SQLite",
            f"Filas nuevas insertadas: {insertadas}\nFilas existentes actualizadas: {actualizadas}",
        )

    def _abrir_seleccion_duplicados(self, duplicados, fecha_conf, on_elegir):
        """Ventana emergente para elegir entre filas duplicadas del Excel.

        `duplicados` es una lista de dict con al menos '_excel_row' y campos
        diferenciadores. Al elegir una opción y presionar 'Finalizar', se
        ejecuta on_elegir(num_fila_excel).
        """
        top = ctk.CTkToplevel(self)
        top.title("Seleccionar Fila Duplicada")
        top.geometry("640x500")
        top.configure(fg_color=CEMEX_GRAY_BG)
        top.grab_set()

        ctk.CTkLabel(
            top, text="Se encontraron varias filas coincidentes",
            font=CTkFont("Segoe UI", 15, "bold"), text_color=CEMEX_BLUE,
        ).grid(row=0, column=0, padx=20, pady=(16, 4), sticky="w")

        ctk.CTkLabel(
            top, text="Elige cuál fila del Excel deseas afectar:",
            font=CTkFont("Segoe UI", 12), text_color=CEMEX_TEXT,
        ).grid(row=1, column=0, padx=20, pady=(0, 8), sticky="w")

        seleccion = StringVar(value="")

        container = ctk.CTkScrollableFrame(top, fg_color="transparent", height=320)
        container.grid(row=2, column=0, padx=20, pady=(0, 8), sticky="ew")
        container.grid_columnconfigure(0, weight=1)

        for i, d in enumerate(duplicados):
            num_fila = d.get("_excel_row", "?")
            texto = (
                f"Fila {num_fila} | Centro: {d.get('Centro','')} | "
                f"ShipFrom: {d.get('ShipFrom','')} | Material: {d.get('Material','')} | "
                f"Modalidad: {d.get('Modalidad','')}"
            )
            ctk.CTkRadioButton(
                container,
                text=texto,
                variable=seleccion,
                value=str(num_fila),
                font=CTkFont("Segoe UI", 12),
                text_color=CEMEX_TEXT,
                fg_color=CEMEX_BLUE,
                hover_color=CEMEX_BLUE_DARK,
                wraplength=560,
            ).grid(row=i, column=0, padx=10, pady=6, sticky="w")

        def finalizar():
            valor = seleccion.get()
            if not valor:
                messagebox.showwarning("Selección", "Elige una fila duplicada.", parent=top)
                return
            top.destroy()
            on_elegir(int(valor))

        ctk.CTkButton(
            top, text="Finalizar", command=finalizar,
            fg_color=CEMEX_BLUE, hover_color=CEMEX_BLUE_DARK,
            font=CTkFont("Segoe UI", 13, "bold"), height=40,
        ).grid(row=3, column=0, padx=20, pady=(8, 16), sticky="ew")

    def fpa_proyectar_m3(self):
        row = self._fpa_current_row()
        if row is None:
            messagebox.showwarning("FPA", "Seleccione una fila en el área de pegado.")
            return
        mp_m3, flete_m3 = fpa_validations.proyectar_m3(row)
        um = fpa_validations._s(row.get("UM")).upper()
        msg = (
            "Proyección M3 (solo lectura, no altera la data cruda)\n"
            "-----------------------------------------------------\n"
            f"UM capturada: {um}\n"
            f"Precio Material (M3): {mp_m3:.4f}\n"
            f"Precio Flete (M3): {flete_m3:.4f}\n"
        )
        if um == "TN":
            msg += "Nota: TN -> M3 multiplicando por PV.\n"
        self._fpa_show_message(msg)

    def fpa_validar(self):
        row = self._fpa_current_row()
        if row is None:
            messagebox.showwarning("FPA", "Seleccione una fila en el área de pegado.")
            return
        self._abrir_modal_validar(row)

    def _abrir_modal_validar(self, row):
        self._abrir_modal_validar_evidencias(row)

    def fpa_abrir_evidencias(self):
        row = self._fpa_current_row()
        if row is None:
            messagebox.showwarning("FPA", "Seleccione una fila antes de validar.")
            return
        self._abrir_modal_validar_evidencias(row)

    def _abrir_modal_validar_evidencias(self, row):
        # Modo dinámico: si hay varias filas seleccionadas, se abre una versión
        # "mini" (solo fecha + evidencias); si hay una sola, el detalle completo.
        seleccion = list(self.fpa_tree.selection())
        multiple = len(seleccion) > 1

        top = ctk.CTkToplevel(self)
        if multiple:
            top.title(f"Validar {len(seleccion)} Solicitudes")
            top.geometry("440x400")
        else:
            top.title("Validar y Evidencias FPA")
            top.geometry("760x720")
        top.configure(fg_color=CEMEX_GRAY_BG)
        top.grab_set()
        top.resizable(True, True)

        imagenes = []

        # Contenedor con scroll para que quepa resumen + evidencias + botones
        container = ctk.CTkScrollableFrame(top, fg_color="transparent")
        container.pack(fill="both", expand=True, padx=8, pady=8)
        container.grid_columnconfigure(0, weight=1)
        container.grid_columnconfigure(1, weight=1)

        # Barra superior con botón minimizar
        bar = ctk.CTkFrame(container, fg_color="transparent")
        bar.grid(row=0, column=0, columnspan=2, padx=20, pady=(0, 4), sticky="ew")
        ctk.CTkButton(
            bar, text="—", width=36, height=28,
            command=lambda: top.iconify(),
            fg_color=CEMEX_NEUTRAL, hover_color=CEMEX_NEUTRAL_DARK,
            font=CTkFont("Segoe UI", 12, "bold"),
        ).grid(row=0, column=0, sticky="e")

        status_var = StringVar(value="")

        def _copiar(valor):
            texto = str(valor)
            top.clipboard_clear()
            top.clipboard_append(texto)
            status_var.set(f"Copiado: {texto}")

        if multiple:
            # Versión mini: solo conteo + fecha + evidencias + botones
            ctk.CTkLabel(
                container,
                text=f"Se procesarán {len(seleccion)} solicitudes seleccionadas.",
                font=CTkFont("Segoe UI", 15, "bold"), text_color=CEMEX_BLUE,
            ).grid(row=1, column=0, columnspan=2, padx=20, pady=(12, 8), sticky="w")
            base_row = 2
        else:
            # Versión detallada: resumen de la solicitud (valores clicables).
            ctk.CTkLabel(
                container, text="Resumen de la Solicitud (clic en un valor para copiarlo)",
                font=CTkFont("Segoe UI", 16, "bold"), text_color=CEMEX_BLUE,
            ).grid(row=1, column=0, columnspan=2, padx=20, pady=(12, 8), sticky="w")

            # Si UM viene en TN, mostrar además los precios en M3 (TN * PV)
            um_val = fpa_validations._s(row.get("UM")).upper()
            mp_m3, flete_m3 = fpa_validations.proyectar_m3(row)

            info = [
                ("Org. Ventas", row.get("Org.Ventas", "")),
                ("ShipFrom", row.get("ShipFrom", "")),
                ("Centro", row.get("Centro", "")),
                ("No. Cuenta Destino SAP", row.get("No. De Destino SAP", "")),
                ("Modalidad", row.get("Modalidad", "")),
                ("Condición Expedición", row.get("Descripción Modalidad", "")),
                ("Material", row.get("Material", "")),
                ("Condición precio material", row.get("Condición precio material", "")),
                ("Precio Material", row.get("Precio Material", "")),
                ("Condición precio flete", row.get("Condición precio flete", "")),
                ("Precio Flete", row.get("Precio Flete", "")),
                ("UM", row.get("UM", "")),
            ]

            if um_val == "TN":
                info.append(("Precio Material (M3)", f"{mp_m3:.4f}"))
                info.append(("Precio Flete (M3)", f"{flete_m3:.4f}"))

            for i, (label, value) in enumerate(info, start=2):
                ctk.CTkLabel(
                    container, text=f"{label}:", font=CTkFont("Segoe UI", 12, "bold"),
                    text_color=CEMEX_TEXT, anchor="w", width=200,
                ).grid(row=i, column=0, padx=(20, 8), pady=4, sticky="w")
                ctk.CTkButton(
                    container, text=str(value), command=lambda v=value: _copiar(v),
                    font=CTkFont("Segoe UI", 12),
                    fg_color="transparent", text_color=CEMEX_TEXT,
                    hover_color=CEMEX_BLUE_LIGHT,
                    anchor="w", height=30, corner_radius=6,
                ).grid(row=i, column=1, padx=(8, 20), pady=4, sticky="ew")

            base_row = len(info) + 2

            ctk.CTkLabel(
                container, textvariable=status_var, font=CTkFont("Segoe UI", 11),
                text_color=CEMEX_BLUE, anchor="w",
            ).grid(row=base_row, column=0, columnspan=2, padx=20, pady=(8, 4), sticky="w")

        # Separador visual
        sep = ctk.CTkFrame(container, fg_color=CEMEX_LIGHT_GRAY, height=1)
        sep.grid(row=base_row + 1, column=0, columnspan=2, padx=20, pady=(4, 8), sticky="ew")

        # Evidencias FPA
        ctk.CTkLabel(
            container, text="Evidencias FPA (pega imágenes desde SAP con Ctrl+V)",
            font=CTkFont("Segoe UI", 16, "bold"), text_color=CEMEX_BLUE,
        ).grid(row=base_row + 2, column=0, columnspan=2, padx=20, pady=(8, 4), sticky="w")

        status = ctk.CTkLabel(
            container, text="Pega una imagen con Ctrl+V.",
            font=CTkFont("Segoe UI", 12), text_color=CEMEX_TEXT,
        )
        status.grid(row=base_row + 3, column=0, columnspan=2, padx=20, pady=4, sticky="w")

        # Fecha de Confirmación
        fecha_var = StringVar(value=datetime.now().strftime("%d/%m/%Y"))
        ctk.CTkLabel(
            container, text="Fecha de Confirmación:",
            font=CTkFont("Segoe UI", 12, "bold"), text_color=CEMEX_TEXT,
        ).grid(row=base_row + 4, column=0, padx=20, pady=(8, 4), sticky="w")
        fecha_entry = ctk.CTkEntry(
            container, textvariable=fecha_var,
            font=CTkFont("Segoe UI", 12), width=160,
            fg_color=CEMEX_WHITE, border_color=CEMEX_LIGHT_GRAY,
            text_color=CEMEX_TEXT,
        )
        fecha_entry.grid(row=base_row + 4, column=1, padx=(8, 20), pady=(8, 4), sticky="w")

        # Vista previa del mensaje
        ctk.CTkLabel(
            container, text="Vista previa del mensaje para Outlook:",
            font=CTkFont("Segoe UI", 12, "bold"), text_color=CEMEX_TEXT,
        ).grid(row=base_row + 5, column=0, columnspan=2, padx=20, pady=(6, 4), sticky="w")
        preview = ctk.CTkTextbox(
            container, height=140, wrap="word",
            font=CTkFont("Consolas", 11),
            fg_color="#FAFAFA", text_color=CEMEX_TEXT,
            border_width=1, border_color=CEMEX_LIGHT_GRAY,
            state="disabled",
        )
        preview.grid(row=base_row + 6, column=0, columnspan=2, padx=20, pady=(0, 8), sticky="ew")

        def _refresh_preview():
            fecha = fecha_var.get().strip() or datetime.now().strftime("%d/%m/%Y")
            texto = f"Ok {fecha}"
            partes = [texto]
            for i in range(1, len(imagenes) + 1):
                partes.append(f"[Imagen {i}]")
            preview.configure(state="normal")
            preview.delete("1.0", "end")
            preview.insert("end", "\n".join(partes))
            preview.configure(state="disabled")

        _refresh_preview()
        fecha_var.trace_add("write", lambda *_: _refresh_preview())

        def on_paste(event=None):
            img = fpa_validations.capturar_imagen_portapapeles()
            if img is None:
                status.configure(text="No se encontró imagen en el portapapeles.")
                return "break"
            imagenes.append(img)
            status.configure(text=f"Imagen {len(imagenes)} capturada.")
            _refresh_preview()
            return "break"

        top.bind("<Control-v>", on_paste)

        def on_finalizar():
            fecha = fecha_var.get().strip() or datetime.now().strftime("%d/%m/%Y")
            self._fpa_ejecutar_finalizacion(top, fecha, imagenes)

        # Botones al final
        ctk.CTkButton(
            container, text="Finalizar", command=on_finalizar,
            fg_color=CEMEX_BLUE, hover_color=CEMEX_BLUE_DARK,
            font=CTkFont("Segoe UI", 13, "bold"), height=40,
        ).grid(row=base_row + 7, column=0, columnspan=2, padx=20, pady=(8, 4), sticky="ew")

        ctk.CTkButton(
            container, text="Cerrar", command=top.destroy,
            fg_color=CEMEX_NEUTRAL, hover_color=CEMEX_NEUTRAL_DARK,
            font=CTkFont("Segoe UI", 12, "bold"), height=36,
        ).grid(row=base_row + 8, column=0, columnspan=2, padx=20, pady=(4, 16), sticky="ew")

    def _fpa_ejecutar_finalizacion(self, top, fecha, imagenes):
        """Procesa las filas SELECCIONADAS del Treeview FPA una por una.

        Por cada fila: alerta de sobrescritura si ya tiene fecha_conf, guarda
        en SQLite y en Excel, y la elimina visualmente. El mensaje de resumen
        global se muestra solo al terminar todo el bucle. Las filas no
        seleccionadas quedan intactas.
        """
        seleccionadas = list(self.fpa_tree.selection())
        if not seleccionadas:
            messagebox.showwarning("FPA", "No hay filas seleccionadas.", parent=top)
            return

        resumen_ok = []
        resumen_error = []
        fecha_str = fecha or datetime.now().strftime("%d/%m/%Y")

        def _procesar_una(idx_pos):
            if idx_pos >= len(seleccionadas):
                # Bucle terminado: preparar portapapeles y mostrar resumen global
                ok_clip = fpa_validations.preparar_portapapeles_outlook(
                    f"Ok {fecha_str}", imagenes
                )
                partes = [
                    f"Filas procesadas: {len(resumen_ok)}",
                    f"Filas omitidas/error: {len(resumen_error)}",
                    f"Portapapeles: {'OK' if ok_clip else 'Error'}",
                ]
                messagebox.showinfo("Validar y Evidencias FPA", "\n".join(partes), parent=top)

                # Sincronizar self.fpa_rows con lo que queda en el árbol
                self._fpa_sync_rows_from_tree()
                self.fpa_selected_index = None
                self.fpa_selected_request = None
                self._fpa_show_message(
                    "Validación finalizada.\n"
                    f"Procesadas: {len(resumen_ok)} | Omitidas/error: {len(resumen_error)}"
                )
                top.destroy()
                return

            iid = seleccionadas[idx_pos]
            try:
                row_idx = int(iid)
                row = self.fpa_rows[row_idx]
            except (ValueError, IndexError):
                resumen_error.append(iid)
                self.after(10, _procesar_una, idx_pos + 1)
                return

            # Alerta de sobrescritura: SELECT a SQLite por la clave de la fila
            try:
                matches = fpa_validations.buscar_en_sqlite(self.conn, row)
                tiene_fecha_previa = any(m.get("fecha_conf") for m in matches)
            except Exception:
                tiene_fecha_previa = False

            if tiene_fecha_previa:
                if not messagebox.askyesno(
                    "Sobrescribir Fecha de Confirmación",
                    "La solicitud ya tiene fecha_conf registrada. ¿Deseas sobrescribirla?",
                    parent=top,
                ):
                    resumen_error.append(iid)
                    self.after(10, _procesar_una, idx_pos + 1)
                    return

            # Guardar en SQLite (una fila) y en Excel (una fila)
            try:
                fpa_validations.actualizar_fecha_confirmacion_sqlite(
                    self.conn, [row], fecha_str
                )
                fpa_validations.actualizar_fecha_confirmacion_excel([row], fecha_str)
            except Exception as e:
                resumen_error.append(iid)
                self._fpa_show_message(f"Error procesando la fila {iid}: {e}")
                self.after(10, _procesar_una, idx_pos + 1)
                return

            resumen_ok.append(iid)
            # Eliminar visualmente la fila procesada
            self.fpa_tree.delete(iid)

            self.after(10, _procesar_una, idx_pos + 1)

        _procesar_una(0)

    def _fpa_clear_rows(self):
        """Borra todas las filas de la solicitud activa en el Área FPA."""
        self.fpa_rows = []
        self.fpa_selected_index = None
        self.fpa_selected_request = None
        self.fpa_db_matches = []
        self.fpa_excel_matches = []
        self._refresh_fpa_tree()
        try:
            self.fpa_paste_box.delete("1.0", "end")
            self.fpa_paste_box.insert("1.0", "ZONA DE PEGADO: haz clic aquí y presiona Ctrl+V")
        except Exception:
            pass
        self._fpa_show_message("Solicitud finalizada. Zona de pegado lista para una nueva solicitud.")

    # ------------------------------------------------------------------
    # Base de datos
    # ------------------------------------------------------------------
    def _init_db(self):
        self.conn = database_manager.create_connection(DB_FILE)
        if self.conn:
            database_manager.migrate_if_needed(self.conn)
        else:
            messagebox.showerror("Error de BD", "No se pudo conectar a SQLite.")

    # ------------------------------------------------------------------
    # Sincronización manual (subir/cargar datos)
    # ------------------------------------------------------------------
    def _cerrar_conn(self):
        """Cierra la conexión SQLite para liberar el archivo .db."""
        if getattr(self, "conn", None):
            try:
                self.conn.close()
            except Exception:
                pass
            self.conn = None

    def _reabrir_conn(self):
        """Reabre la conexión SQLite y refresca la tabla SLA."""
        self.conn = database_manager.create_connection(DB_FILE)
        if self.conn:
            database_manager.migrate_if_needed(self.conn)
        self._refresh_sla_tree()

    def _refresh_sync_status(self):
        """Actualiza el texto de la última actividad de sincronización."""
        try:
            texto = sync_manager.get_last_activity()
        except Exception:
            texto = ""
        self.sync_status_label.configure(text=texto or "Sin actividad registrada")

    def sync_upload(self):
        if not sync_manager.hay_conexion_internet():
            messagebox.showwarning("Sin conexión", "No hay conexión a internet. Revisa tu red e intenta de nuevo.")
            return
        # Liberar el .db para que no esté bloqueado al copiar
        self._cerrar_conn()
        ok, msg = sync_manager.upload_to_drive()
        self._reabrir_conn()
        self._refresh_sync_status()
        if ok:
            messagebox.showinfo("Sincronización", msg)
        elif "conflicto" in msg.lower() or "recuperación" in msg.lower():
            messagebox.showwarning("Sincronización bloqueada", msg)
        else:
            messagebox.showerror("Sincronización", msg)

    def sync_recuperacion(self):
        if not sync_manager.hay_conexion_internet():
            messagebox.showwarning("Sin conexión", "No hay conexión a internet. Revisa tu red e intenta de nuevo.")
            return
        if not self.conn:
            messagebox.showerror("Sincronización de Recuperación", "No hay conexión a SQLite.")
            return
        if not messagebox.askyesno(
            "Sincronización de Recuperación",
            "Esto descargará el Excel actualizado del drive y le aplicará tus "
            "cambios locales de hoy (UPDATE/INSERT) en segundo plano.\n\n"
            "¿Continuar?",
        ):
            return

        self.recuperacion_btn.configure(state="disabled", text="Recuperando...")

        def _run():
            def _done(ok, msg):
                self.after(0, self._on_recuperacion_done, ok, msg)
            try:
                sync_manager.sincronizar_recuperacion(self.conn, on_done=_done)
            except Exception as e:
                # Si la sincronización revienta fuera de su manejo interno,
                # garantizamos que el botón se restaure y se muestre el error.
                self.after(0, self._on_recuperacion_done, False, str(e))

        threading.Thread(target=_run, daemon=True).start()

    def _on_recuperacion_done(self, ok, msg):
        self.recuperacion_btn.configure(state="normal", text="Sincronización de Recuperación")
        self._refresh_sla_tree()
        self._refresh_sync_status()
        if ok:
            messagebox.showinfo("Sincronización de Recuperación", msg)
        else:
            messagebox.showerror("Sincronización de Recuperación", msg)

    def sync_download(self):
        if not sync_manager.hay_conexion_internet():
            messagebox.showwarning("Sin conexión", "No hay conexión a internet. Revisa tu red e intenta de nuevo.")
            return
        if not messagebox.askyesno(
            "Cargar datos",
            "Esto SOBRESCRIBIRÁ tu base de datos y Excel locales con la "
            "versión del drive.\n\nSe creará una copia de seguridad antes.\n"
            "¿Continuar?",
        ):
            return
        # Liberar el .db para poder sobrescribirlo
        self._cerrar_conn()
        ok, msg = sync_manager.download_from_drive()
        self._reabrir_conn()
        self._refresh_sync_status()

        # Limpiar cualquier filtro de búsqueda para que se vean todos los datos
        try:
            self._sla_clear_search()
        except Exception:
            pass

        if ok:
            # Sincronizar SQLite con lo que venga en el Excel descargado (upsert)
            migra_msg = ""
            try:
                ins, act, migra_err = fpa_validations.migrar_excel_a_sqlite(self.conn)
                if migra_err:
                    migra_msg = f"\n\nAviso al sincronizar Excel→SQLite: {migra_err}"
                else:
                    migra_msg = (
                        f"\n\nSincronización Excel→SQLite:\n"
                        f"  Nuevas: {ins}\n  Actualizadas: {act}"
                    )
                self._refresh_sla_tree()
            except Exception as e:
                migra_msg = f"\n\nError sincronizando Excel→SQLite: {e}"

            messagebox.showinfo("Sincronización", f"{msg}{migra_msg}")
        else:
            messagebox.showerror("Sincronización", msg)

    # ------------------------------------------------------------------
    # Acción: Calcular Fila
    # ------------------------------------------------------------------
    def calculate_row(self):
        manual = {key: var.get().strip() for key, var in self.vars.items()}

        required = ["cedis", "sf", "planta_obra", "articulo", "modalidad", "um_vta"]
        if any(not self.vars[k].get().strip() for k in required):
            messagebox.showwarning(
                "Campos incompletos",
                "CEDIS, SF, Planta/Obra, Artículo, Modalidad y UM VTA son obligatorios."
            )
            return

        try:
            if self.sap_data is None:
                self.sap_data = etl_processor.load_sap_data()

            self.last_calculated = etl_processor.calculate_row(manual, self.sap_data)
            self._render_preview(self.last_calculated)
            self.add_batch_button.configure(state="normal")
        except Exception as e:
            messagebox.showerror("Error al calcular", f"No se pudo calcular la fila:\n{e}")

    # ------------------------------------------------------------------
    # Hotkeys de "Mantener Precio"
    # ------------------------------------------------------------------
    def toggle_mantener_mp(self):
        nuevo = not self.mantener_mp_var.get()
        self.mantener_mp_var.set(nuevo)
        if nuevo:
            self.mantener_mp_cb.select()
        else:
            self.mantener_mp_cb.deselect()

    def toggle_mantener_flete(self):
        nuevo = not self.mantener_flete_var.get()
        self.mantener_flete_var.set(nuevo)
        if nuevo:
            self.mantener_flete_cb.select()
        else:
            self.mantener_flete_cb.deselect()

    # ------------------------------------------------------------------
    # Pegado inteligente en campos de costos (Ctrl+V)
    # ------------------------------------------------------------------
    def _smart_paste_cost(self, key, event=None):
        if event is not None:
            try:
                raw = self.clipboard_get()
            except Exception:
                return "break"
        else:
            try:
                raw = self.clipboard_get()
            except Exception:
                return

        if raw is None or not raw.strip():
            return "break" if event is not None else None

        values = raw.split("\t")
        targets = COST_PASTE_GROUPS.get(key, (key,))

        for i, target_key in enumerate(targets):
            if i < len(values):
                self.vars[target_key].set(self._clean_money(values[i]))

        return "break" if event is not None else None

    @staticmethod
    def _clean_money(value):
        """Limpia formato de dinero a un número con punto decimal para captura.

        Acepta '$1,234.56', '1.234,56', '1,234', '123.45'. Devuelve una cadena
        lista para guardar (usa '.' como separador decimal)."""
        s = str(value).strip()
        if s == "":
            return ""
        s = s.replace("$", "").replace("MXN", "").replace("USD", "").replace(" ", "")
        s = s.replace("mxn", "").replace("usd", "")
        if s == "":
            return ""

        # Ambos separadores: el de la derecha es decimal, el otro miles.
        if "," in s and "." in s:
            if s.rfind(",") > s.rfind("."):
                s = s.replace(".", "").replace(",", ".")
            else:
                s = s.replace(",", "")
        elif "," in s and "." not in s:
            if len(s.rsplit(",", 1)[1]) == 3 and s.count(",") == 1:
                s = s.replace(",", "")
            else:
                s = s.replace(",", ".")
        return s

    # ------------------------------------------------------------------
    # Auto-llenado / pegado masivo desde portapapeles (Excel -> GUI)
    # ------------------------------------------------------------------
    def _on_global_paste(self, event=None):
        """Enruta Ctrl+V: si estamos en el Área FPA, pega en su Treeview;
        en cualquier otro caso, usa el flujo de la Mesa de Captura."""
        if getattr(self, "active_view", None) == "fpa":
            return self._fpa_paste(event)
        return self.paste_from_clipboard(event)

    def paste_from_clipboard(self, event=None):
        try:
            raw = self.clipboard_get()
        except Exception:
            return

        if not raw.strip():
            return

        lines = [line for line in raw.splitlines() if line.strip()]
        if not lines:
            return

        if len(lines) == 1:
            # Pegado de una sola fila: llenar directamente el formulario
            values = lines[0].split('\t')
            for i, key in enumerate(PASTE_FIELDS):
                if i < len(values):
                    self.vars[key].set(values[i].strip())
            return "break"

        # Pegado masivo de múltiples filas: insertar en Treeview como Pendiente
        for line in lines:
            values = line.split('\t')
            row_values = values[:11] + [''] * (11 - len(values))
            self.pending_rows.append({key: row_values[i].strip() for i, key in enumerate(PASTE_FIELDS)})

        self._refresh_pending_tree()

        # Auto-seleccionar la primera fila para cargar los datos al formulario
        items = self.pending_tree.get_children()
        if items:
            self.pending_tree.selection_set(items[0])
            self.pending_tree.focus(items[0])
            self._on_pending_select()

        messagebox.showinfo("Pegado masivo", f"Se importaron {len(lines)} filas pendientes desde Excel.")
        return "break"

    def _refresh_pending_tree(self):
        for item in self.pending_tree.get_children():
            self.pending_tree.delete(item)
        for i, row in enumerate(self.pending_rows):
            values = ("Pendiente",) + tuple(row.get(k, '') for k in PASTE_FIELDS)
            tag = "even" if i % 2 == 0 else "odd"
            self.pending_tree.insert("", "end", iid=str(i), values=values, tags=(tag,))

    def _on_pending_select(self, event=None):
        selected = self.pending_tree.selection()
        if not selected:
            return
        iid = selected[0]
        try:
            row = self.pending_rows[int(iid)]
        except (ValueError, IndexError):
            return
        self.pending_selected_index = int(iid)
        for key in PASTE_FIELDS:
            self.vars[key].set(row.get(key, ''))

    def _seleccionar_siguiente_pendiente(self):
        """Selecciona automáticamente la primera fila con estado Pendiente/Pendiente de Costos."""
        for item in self.pending_tree.get_children():
            values = self.pending_tree.item(item, "values")
            if values and values[0] in ("Pendiente", "Pendiente de Costos"):
                self.pending_tree.selection_set(item)
                self.pending_tree.focus(item)
                self._on_pending_select()
                return True
        return False

    # ------------------------------------------------------------------
    # Vista previa
    # ------------------------------------------------------------------
    def _render_preview(self, fila: dict):
        self.preview_text.configure(state="normal")
        self.preview_text.delete("1.0", "end")

        lines = []
        for col in "ABCDEFGHIJKLMNOPQRSTUVWXYZ":
            prefix = col + " "
            for key, val in fila.items():
                if key.startswith(prefix):
                    lines.append(f"{key:<35} = {val}")
                    break
        for key, val in fila.items():
            if key[:2] in ("AA", "AB", "AC", "AD", "AE", "AF", "AG", "AH", "AI",
                            "AJ", "AK", "AL", "AM", "AN", "AO", "AP", "AQ", "AR"):
                lines.append(f"{key:<35} = {val}")

        seen = set()
        unique_lines = []
        for line in lines:
            if line not in seen:
                seen.add(line)
                unique_lines.append(line)

        self.preview_text.insert("end", "\n".join(unique_lines))
        self.preview_text.configure(state="disabled")

    # ------------------------------------------------------------------
    # Helper: leer FORMATO DESTINO desde checkboxes
    # ------------------------------------------------------------------
    def get_formato_destino(self):
        seleccionados = [op for op, cb in self.formato_checkboxes.items() if cb.get()]
        return ", ".join(seleccionados)

    # ------------------------------------------------------------------
    # Acción: Agregar Fila a Lote
    # ------------------------------------------------------------------
    def agregar_fila_a_lote(self):
        if self.last_calculated is None:
            messagebox.showwarning("Sin calcular", "Primero presione 'Calcular Fila'.")
            return

        formato = self.get_formato_destino()
        if not formato:
            messagebox.showwarning(
                "Formato destino requerido",
                "Seleccione al menos un FORMATO DESTINO antes de agregar."
            )
            return

        # Sincronizar formato_destino y asignar fecha de procesamiento si está vacía
        self.vars["formato_destino"].set(formato)
        if not self.vars["fecha_procesamiento"].get().strip():
            self.vars["fecha_procesamiento"].set(datetime.now().strftime("%d/%m/%Y"))
        manual = {key: var.get().strip() for key, var in self.vars.items()}

        # Persistir en SQLite
        if self.conn:
            orden = ["org_vta", "cedis", "nombre_cedis", "sf", "nombre_sf",
                     "planta_obra", "nombre_planta", "articulo", "desc_articulo", "pv",
                     "modalidad", "um_vta",
                     "mp_vta_m3", "flete_vta_m3",
                     "mp_comp_m3", "ca_comp_m3", "flete_comp_m3", "integral_comp_m3",
                     "tipo", "fecha_procesamiento", "tipo_cliente", "observacion",
                     "fecha_solicitud", "incidencia", "motivo", "mes", "fecha_conf",
                     "formato_destino"]
            data = tuple(self.vars[k].get().strip() for k in orden)
            try:
                database_manager.insert_data(self.conn, data)
            except Exception as e:
                messagebox.showerror("Error de BD", f"No se pudo guardar en SQLite:\n{e}")
                return

        # Agregar a lote en memoria
        self.batch_rows.append(manual)
        self._update_counter()

        # Si la fila provenía del pegado masivo, removerla del área pendiente
        if self.pending_selected_index is not None:
            try:
                self.pending_rows.pop(self.pending_selected_index)
                self._refresh_pending_tree()
            except IndexError:
                pass
            self.pending_selected_index = None

        messagebox.showinfo("Lote", f"Fila agregada. Filas en solicitud actual: {len(self.batch_rows)}")

        # Reset para siguiente captura
        self.last_calculated = None
        self.add_batch_button.configure(state="disabled")
        self.clear_fields()

        # Auto-navegación: seleccionar la siguiente fila pendiente y cargarla
        self.after(100, self._seleccionar_siguiente_pendiente)

    # ------------------------------------------------------------------
    # Acción: Generar Archivos de Solicitud
    # ------------------------------------------------------------------
    def generar_archivos_solicitud(self):
        if not self.batch_rows:
            messagebox.showwarning("Lote vacío", "No hay filas en la solicitud actual.")
            return

        self.generate_button.configure(state="disabled", text="Procesando...")

        def _run():
            try:
                folder = etl_processor.procesar_lote(self.batch_rows)
                self.after(0, self._on_generar_complete, folder, None)
            except Exception as e:
                self.after(0, self._on_generar_complete, None, e)

        threading.Thread(target=_run, daemon=True).start()

    def _on_generar_complete(self, folder, error):
        self.generate_button.configure(state="normal", text="Generar Archivos de Solicitud")
        if isinstance(error, etl_processor.ExcelBloqueadoError):
            messagebox.showwarning("Excel ocupado", str(error))
        elif isinstance(error, PermissionError):
            messagebox.showerror("Archivo bloqueado", str(error))
        elif error:
            messagebox.showerror("Error de exportación", f"Fallo al generar los archivos:\n{error}")
        elif folder:
            messagebox.showinfo("Solicitud generada", f"Se generaron los archivos en:\n{folder}")
            self.batch_rows.clear()
            self._update_counter()

    # ------------------------------------------------------------------
    # Proceso ETL histórico (todos los registros en SQLite)
    # ------------------------------------------------------------------
    def run_etl(self):
        self.etl_button.configure(state="disabled", text="Procesando...")

        def _run():
            try:
                df = etl_processor.process_and_validate_data(DB_FILE, self.sap_data)
                self.after(0, self._on_etl_complete, df, None)
            except Exception as e:
                self.after(0, self._on_etl_complete, None, e)

        threading.Thread(target=_run, daemon=True).start()

    def _on_etl_complete(self, df, error):
        self.etl_button.configure(state="normal", text="Procesar ETL Histórico")
        if isinstance(error, etl_processor.ExcelBloqueadoError):
            messagebox.showwarning("Excel ocupado", str(error))
        elif isinstance(error, PermissionError):
            messagebox.showerror("Archivo bloqueado", str(error))
        elif error:
            messagebox.showerror("Error de exportación", f"Fallo al inyectar datos en el Excel Maestro:\n{error}")
        elif df is not None and not df.empty:
            messagebox.showinfo(
                "ETL completado",
                f"Datos procesados: {len(df)} filas."
            )
        else:
            messagebox.showwarning("ETL", "No hay datos en staging para procesar.")

    # ------------------------------------------------------------------
    # Validaciones y SLA
    # ------------------------------------------------------------------
    def _refresh_sla_tree(self):
        if not self.conn:
            return

        cur = self.conn.cursor()
        cur.execute(
            """
            SELECT id, org_vta, cedis, sf, nombre_sf, planta_obra, nombre_planta,
                   articulo, modalidad, fecha_solicitud, fecha_procesamiento, fecha_conf
            FROM staging_rutas
            ORDER BY id DESC
            """
        )
        rows = cur.fetchall()

        for item in self.sla_tree.get_children():
            self.sla_tree.delete(item)

        for idx, row in enumerate(rows):
            (row_id, org_vta, cedis, sf, nombre_sf, planta, nombre_planta,
             articulo, modalidad, fecha_solicitud, fecha_procesamiento, fecha_conf) = row

            ruta = "-".join(str(x) for x in [cedis, sf, planta, articulo, modalidad] if x)
            dias_habiles, dias_transcurridos, estado = etl_processor.calcular_estado_sla(
                fecha_procesamiento, fecha_conf
            )
            dias_habiles_str = str(dias_habiles) if dias_habiles is not None else ""
            dias_trans_str = str(dias_transcurridos) if dias_transcurridos is not None else ""

            tag = "even" if idx % 2 == 0 else "odd"
            self.sla_tree.insert(
                "", "end",
                iid=str(row_id),
                values=(
                    row_id, ruta, org_vta, cedis, sf, planta, articulo, modalidad,
                    fecha_solicitud or "", fecha_procesamiento or "", fecha_conf or "",
                    dias_habiles_str, dias_trans_str, estado
                ),
                tags=(tag,)
            )

    def _on_sla_select(self, event=None):
        selected = self.sla_tree.selection()
        if not selected:
            return
        values = self.sla_tree.item(selected[0], "values")
        # values: ID, RUTA, ORGVTA, CEDIS, SF, NO PLANTA, ARTICULO, MODALIDAD,
        #         FECHA SOLICITUD, FECHA PROCESAMIENTO, FECHA CONF,
        #         DÍAS HÁBILES, DÍAS TRANSCURRIDOS, ESTADO
        self.sla_fecha_solicitud.delete(0, "end")
        self.sla_fecha_solicitud.insert(0, values[8])
        self.sla_fecha_conf.delete(0, "end")
        self.sla_fecha_conf.insert(0, values[10])

    def _guardar_fechas_sla(self):
        selected = self.sla_tree.selection()
        if not selected:
            messagebox.showwarning("Selección", "Seleccione una solicitud en la tabla.")
            return

        row_id = int(selected[0])
        fecha_solicitud = self.sla_fecha_solicitud.get().strip()
        fecha_conf = self.sla_fecha_conf.get().strip()

        try:
            cur = self.conn.cursor()
            # Obtener clave compuesta + discriminadores para ubicarla en Excel
            cur.execute(
                """SELECT cedis, sf, planta_obra, articulo, modalidad,
                          fecha_procesamiento, tipo_cliente, fecha_solicitud, um_vta
                   FROM staging_rutas WHERE id=?""",
                (row_id,),
            )
            clave = cur.fetchone()
            cur.execute(
                "UPDATE staging_rutas SET fecha_solicitud=?, fecha_conf=? WHERE id=?",
                (fecha_solicitud if fecha_solicitud else None,
                 fecha_conf if fecha_conf else None,
                 row_id)
            )
            self.conn.commit()

            # Escribir FECHA CONF también en el Excel Maestro (búsqueda por ramas)
            excel_msg = ""
            if clave and fecha_conf:
                cedis, sf, planta, articulo, modalidad, fecha_proc, tipo_cliente, fecha_sol, um_vta = clave
                excel_msg = self._resolver_y_actualizar_excel(
                    cedis, sf, planta, articulo, modalidad, fecha_conf,
                    fecha_procesamiento=fecha_proc,
                    tipo_cliente=tipo_cliente,
                    fecha_solicitud=fecha_sol,
                    um_vta=um_vta,
                )
            elif not fecha_conf:
                excel_msg = " | Excel: sin fecha de confirmación, no se actualizó"

            self._refresh_sla_tree()
            messagebox.showinfo("Actualizado", f"Fechas guardadas correctamente.{excel_msg}")
        except Exception as e:
            messagebox.showerror("Error de BD", f"No se pudieron guardar las fechas:\n{e}")

    def _resolver_y_actualizar_excel(self, cedis, sf, planta, articulo, modalidad,
                                     fecha_conf, fecha_procesamiento=None,
                                     tipo_cliente=None, fecha_solicitud=None, um_vta=None):
        """Busca candidatos en el Excel por ramas y pregunta si hay empate.

        Retorna el mensaje de estado para mostrar en el popup.
        """
        candidatos = etl_processor.buscar_candidatos_excel_sla(
            cedis, sf, planta, articulo, modalidad,
            fecha_procesamiento=fecha_procesamiento,
            tipo_cliente=tipo_cliente,
            fecha_solicitud=fecha_solicitud,
            um_vta=um_vta,
        )

        if candidatos is None:
            return " | Excel: no se pudo leer"

        if len(candidatos) == 1:
            r = etl_processor.actualizar_fecha_conf_excel_por_fila(
                candidatos[0]["_excel_row"], fecha_conf
            )
            return " | Excel: 1 fila actualizada" if r else " | Excel: no se pudo actualizar"

        if len(candidatos) == 0:
            return " | Excel: sin coincidencias"

        # Hay varios candidatos: preguntar cuál actualizar
        elegido = self._preguntar_candidato_excel(candidatos)
        if elegido is None:
            return " | Excel: operación cancelada"
        r = etl_processor.actualizar_fecha_conf_excel_por_fila(elegido, fecha_conf)
        return " | Excel: 1 fila actualizada" if r else " | Excel: no se pudo actualizar"

    def _preguntar_candidato_excel(self, candidatos):
        """Muestra un diálogo para elegir entre filas duplicadas del Excel.

        Retorna el número de fila elegido (int) o None si cancela.
        """
        from tkinter import simpledialog

        lineas = ["Se encontraron varias filas coincidentes en el Excel:\n"]
        for i, c in enumerate(candidatos, 1):
            lineas.append(
                f"{i}) Fila Excel {c['_excel_row']} | "
                f"CEDIS {c['cedis']} | SF {c['sf']} | {c['planta_obra']} | "
                f"{c['articulo']} | Mod {c['modalidad']}"
            )
            if c.get("fecha_procesamiento"):
                lineas[-1] += f" | F.Proc {c['fecha_procesamiento']}"
            if c.get("tipo_cliente"):
                lineas[-1] += f" | {c['tipo_cliente']}"
        lineas.append("\nEscribe el número de la fila a actualizar (0 para cancelar):")

        resp = simpledialog.askstring(
            "Seleccionar fila del Excel",
            "\n".join(lineas),
            parent=self,
        )
        if resp is None:
            return None
        try:
            idx = int(str(resp).strip())
        except ValueError:
            messagebox.showwarning("Inválido", "Debes escribir un número.")
            return None
        if idx < 1 or idx > len(candidatos):
            return None
        return candidatos[idx - 1]["_excel_row"]

    def _seleccionar_excel_maestro(self):
        """Abre el selector de archivo para elegir el Excel Maestro y guarda la ruta."""
        ruta = filedialog.askopenfilename(
            title="Seleccionar Excel Maestro",
            initialdir=BASE_DIR,
            filetypes=[("Excel habilitado para macros", "*.xlsm *.xlsb"), ("Excel", "*.xlsx *.xls"), ("Todos", "*.*")],
        )
        if not ruta:
            return
        etl_processor.set_excel_maestro_path(ruta)
        self.excel_path_label.configure(text=f"Excel: {ruta}")
        messagebox.showinfo("Excel Maestro", f"Ruta guardada:\n{ruta}")

    def _generar_archivos_desde_bd(self):
        if not self.conn:
            messagebox.showerror("Error", "No hay conexión a la base de datos.")
            return

        try:
            cur = self.conn.cursor()
            cur.execute("SELECT * FROM staging_rutas")
            cols = [d[0] for d in cur.description]
            rows = cur.fetchall()
        except Exception as e:
            messagebox.showerror("Error de BD", f"No se pudo leer la base de datos:\n{e}")
            return

        if not rows:
            messagebox.showwarning("Vacío", "No hay registros en la base de datos.")
            return

        manual_rows = []
        for row in rows:
            d = dict(zip(cols, row))
            manual_rows.append(d)

        self.generate_sla_button.configure(state="disabled", text="Procesando...")

        def _run():
            try:
                folder = etl_processor.procesar_lote(manual_rows, base_output_dir=EXPORT_DIR)
                self.after(0, self._on_generar_bd_complete, folder, None)
            except Exception as e:
                self.after(0, self._on_generar_bd_complete, None, e)

        threading.Thread(target=_run, daemon=True).start()

    def _on_generar_bd_complete(self, folder, error):
        self.generate_sla_button.configure(state="normal", text="Generar Archivos desde BD")
        if isinstance(error, etl_processor.ExcelBloqueadoError):
            messagebox.showwarning("Excel ocupado", str(error))
        elif isinstance(error, PermissionError):
            messagebox.showerror("Archivo bloqueado", str(error))
        elif error:
            messagebox.showerror("Error de exportación", f"Fallo al generar los archivos:\n{error}")
        elif folder:
            messagebox.showinfo("Generado", f"Archivos generados en:\n{folder}")

    # ------------------------------------------------------------------
    # Terminal SQL interna
    # ------------------------------------------------------------------
    def _abrir_terminal_sql(self):
        """Abre una terminal SQL cruda contra staging_rutas.db.

        Los SELECT se muestran en un Treeview dinámico; los UPDATE/DELETE/
        INSERT muestran un mensaje de éxito/error. Diseñada como ventana
        independiente para no saturar la interfaz principal.
        """
        if not self.conn:
            messagebox.showerror("Terminal SQL", "No hay conexión a SQLite.")
            return

        top = ctk.CTkToplevel(self)
        top.title("Terminal SQL — staging_rutas.db")
        top.geometry("980x640")
        top.configure(fg_color=CEMEX_GRAY_BG)
        top.grab_set()

        top.grid_columnconfigure(0, weight=1)
        top.grid_rowconfigure(2, weight=1)

        ctk.CTkLabel(
            top, text="Sentencia SQL cruda (sobre staging_rutas):",
            font=CTkFont("Segoe UI", 13, "bold"), text_color=CEMEX_BLUE,
        ).grid(row=0, column=0, padx=14, pady=(14, 6), sticky="w")

        sql_box = ctk.CTkTextbox(
            top, height=110, wrap="none",
            font=CTkFont("Consolas", 12),
            fg_color="#FFF8E1", text_color=CEMEX_TEXT,
            border_width=1, border_color=CEMEX_LIGHT_GRAY,
        )
        sql_box.grid(row=1, column=0, padx=14, pady=(0, 8), sticky="ew")

        results_frame = ctk.CTkFrame(top, fg_color=CEMEX_WHITE, corner_radius=0)
        results_frame.grid(row=2, column=0, padx=14, pady=(0, 8), sticky="nsew")
        results_frame.grid_columnconfigure(0, weight=1)
        results_frame.grid_rowconfigure(0, weight=1)

        # Estado de la última consulta SELECT (para exportar).
        current_tree = {"tree": None, "cols": [], "rows": []}

        def _ejecutar():
            sql = sql_box.get("1.0", "end").strip()
            if not sql:
                messagebox.showwarning("Terminal SQL", "Escribe una sentencia SQL.", parent=top)
                return

            # Limpiar cualquier Treeview anterior
            for child in results_frame.winfo_children():
                child.destroy()
            current_tree["tree"] = None
            current_tree["cols"] = []
            current_tree["rows"] = []

            try:
                cur = self.conn.cursor()
                cur.execute(sql)
                es_select = sql.lstrip().upper().startswith("SELECT")
                if es_select:
                    cols = [d[0] for d in cur.description] if cur.description else []
                    rows = cur.fetchall()
                    if not cols:
                        messagebox.showinfo(
                            "Terminal SQL", "Consulta ejecutada sin columnas para mostrar.", parent=top
                        )
                        return

                    tree = ttk.Treeview(results_frame, columns=cols, show="headings", style="Cemex.Treeview")
                    tree.tag_configure("even", background=CEMEX_WHITE, foreground=CEMEX_TEXT)
                    tree.tag_configure("odd", background=CEMEX_GRAY_BG, foreground=CEMEX_TEXT)
                    tree.grid(row=0, column=0, sticky="nsew")
                    vsb = ttk.Scrollbar(results_frame, orient="vertical", command=tree.yview)
                    vsb.grid(row=0, column=1, sticky="ns")
                    hsb = ttk.Scrollbar(results_frame, orient="horizontal", command=tree.xview)
                    hsb.grid(row=1, column=0, sticky="ew")
                    tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)
                    for col in cols:
                        tree.heading(col, text=col)
                        tree.column(col, width=120, anchor="center")
                    for i, r in enumerate(rows):
                        tag = "even" if i % 2 == 0 else "odd"
                        tree.insert("", "end", values=tuple(r), tags=(tag,))

                    current_tree["tree"] = tree
                    current_tree["cols"] = cols
                    current_tree["rows"] = rows
                else:
                    self.conn.commit()
                    n = cur.rowcount if cur.rowcount is not None and cur.rowcount >= 0 else 0
                    messagebox.showinfo(
                        "Terminal SQL", f"Ejecutado correctamente.\n{'' if n < 0 else f'{n} fila(s) afectada(s).'}",
                        parent=top,
                    )
            except Exception as e:
                try:
                    self.conn.rollback()
                except Exception:
                    pass
                messagebox.showerror("Terminal SQL", f"Error:\n{e}", parent=top)

        def _exportar():
            tree = current_tree["tree"]
            cols = current_tree["cols"]
            rows = current_tree["rows"]
            if tree is None or not cols:
                messagebox.showwarning("Terminal SQL", "No hay resultados SELECT para exportar.", parent=top)
                return

            path = filedialog.asksaveasfilename(
                title="Exportar resultados",
                defaultextension=".xlsm",
                filetypes=[("Libro de Excel con macros", "*.xlsm"), ("Libro de Excel", "*.xlsx")],
            )
            if not path:
                return

            try:
                import pandas as pd
                df = pd.DataFrame([list(r) for r in rows], columns=cols)
                df.to_excel(path, index=False, engine="openpyxl")
            except Exception as e:
                messagebox.showerror("Terminal SQL", f"Error al exportar:\n{e}", parent=top)
                return
            messagebox.showinfo("Terminal SQL", f"Resultados exportados a:\n{path}", parent=top)

        btns = ctk.CTkFrame(top, fg_color="transparent")
        btns.grid(row=3, column=0, padx=14, pady=(0, 14), sticky="ew")
        btns.grid_columnconfigure(0, weight=1)

        ctk.CTkButton(
            btns, text="Ejecutar", command=_ejecutar,
            fg_color=CEMEX_BLUE, hover_color=CEMEX_BLUE_DARK,
            font=CTkFont("Segoe UI", 12, "bold"), height=34,
        ).grid(row=0, column=0, padx=(0, 6), sticky="ew")

        ctk.CTkButton(
            btns, text="Exportar a xlsm", command=_exportar,
            fg_color=CEMEX_RED, hover_color=CEMEX_RED_DARK,
            font=CTkFont("Segoe UI", 12, "bold"), height=34,
        ).grid(row=0, column=1, padx=(0, 6), sticky="ew")

        ctk.CTkButton(
            btns, text="Cerrar", command=top.destroy,
            fg_color=CEMEX_NEUTRAL, hover_color=CEMEX_NEUTRAL_DARK,
            font=CTkFont("Segoe UI", 12, "bold"), height=34,
        ).grid(row=0, column=2, padx=(6, 0), sticky="ew")

    # ------------------------------------------------------------------
    # Utilidades
    # ------------------------------------------------------------------
    def _bind_copy_excel(self, tree):
        """Vincula Ctrl+C a un Treeview para copiar las filas seleccionadas
        en formato estilo Excel (celdas separadas por tabulaciones, filas por
        saltos de línea), listas para pegar en Excel/Google Sheets/editor."""
        def _copiar_excel(event=None):
            seleccion = tree.selection()
            if not seleccion:
                return "break"
            lineas = []
            for iid in seleccion:
                valores = tree.item(iid, "values")
                lineas.append("\t".join(str(v) if v is not None else "" for v in valores))
            texto = "\n".join(lineas)
            self.clipboard_clear()
            self.clipboard_append(texto)
            return "break"
        tree.bind("<Control-c>", _copiar_excel)

    def auto_ajustar_columnas(self, tree, datos):
        """Compatibilidad: delega en el ajuste inteligente."""
        self.ajustar_columnas_inteligente(tree, datos)

    def ajustar_columnas_inteligente(self, tree, datos):
        """Ajusta el ancho de las columnas según el tipo de dato.

        - Texto largo: ancho máximo fijo (se corta, no desborda).
        - Números/fechas: ancho dinámico exacto al valor más largo + padding.
        - Resto: ancho estándar de 100px.

        `datos` puede ser lista de dict (claves = columnas) o lista de listas.
        """
        import tkinter.font as tkfont

        cols_texto_largo = [
            "nombre", "descripcion", "cliente", "observacion", "planta",
            "obra", "cedis", "motivo", "material",
        ]
        cols_numeros_fechas = [
            "precio", "costo", "m3", "monto", "total", "fecha", "id",
            "folio", "tarifa",
        ]

        font_cell = tkfont.Font(font=("Segoe UI", 11))
        cols = list(tree["columns"])
        for col_idx, col in enumerate(cols):
            col_lower = col.lower()
            if any(k in col_lower for k in cols_texto_largo):
                tree.column(col, width=180, stretch=False)
                continue
            if any(k in col_lower for k in cols_numeros_fechas):
                max_len = len(col)
                for row in datos:
                    if isinstance(row, dict):
                        valor = row.get(col, "")
                    elif col_idx < len(row):
                        valor = row[col_idx]
                    else:
                        valor = ""
                    texto = "" if valor is None else str(valor)
                    max_len = max(max_len, len(texto))
                tree.column(col, width=max(60, font_cell.measure("0" * max_len) + 20), stretch=False)
                continue
            tree.column(col, width=100, stretch=True)

    def _abrir_cell_grid(self, tree, titulo="Selección por celdas"):
        """Abre una ventana con un CellGrid (selección por celda real) cargado
        con los encabezados y filas de un Treeview.

        Permite: clic simple (celda), Ctrl+clic (agregar/quitar), Shift+clic o
        arrastre (rango rectangular), flechas para mover la celda activa,
        Shift+flechas para extender la selección horizontal/vertical, y
        Ctrl+C para copiar las celdas seleccionadas al portapapeles.
        """
        headers = list(tree["columns"])
        rows = [list(tree.item(iid, "values")) for iid in tree.get_children()]
        # Aplicar el auto-ajuste inteligente y leer los anchos resultantes.
        self.ajustar_columnas_inteligente(tree, rows)
        col_widths = [int(tree.column(col, "width")) for col in headers]

        top = ctk.CTkToplevel(self)
        top.title(titulo)
        top.geometry("1000x600")
        top.configure(fg_color=CEMEX_GRAY_BG)
        top.grab_set()
        # Pantalla completa en Windows (equivalente a maximizar el Toplevel).
        try:
            top.state("zoomed")
        except Exception:
            pass

        top.grid_columnconfigure(0, weight=1)
        top.grid_rowconfigure(1, weight=1)

        ctk.CTkLabel(
            top,
            text="Clic: celda · Ctrl+clic: agregar/quitar · Shift+clic/arrastrar: rango · "
                 "Flechas: mover · Shift+flechas: extender · Ctrl+C: copiar",
            font=CTkFont("Segoe UI", 11), text_color=CEMEX_TEXT, anchor="w",
        ).grid(row=0, column=0, padx=12, pady=(10, 6), sticky="w")

        body = ctk.CTkFrame(top, fg_color=CEMEX_WHITE, corner_radius=0)
        body.grid(row=1, column=0, padx=12, pady=(0, 12), sticky="nsew")
        body.grid_columnconfigure(0, weight=1)
        body.grid_rowconfigure(0, weight=1)

        grid = cell_grid.CellGrid(body, headers=headers, rows=rows, col_widths=col_widths)
        grid.grid(row=0, column=0, sticky="nsew")

        btns = ctk.CTkFrame(top, fg_color="transparent")
        btns.grid(row=2, column=0, padx=12, pady=(0, 12), sticky="ew")
        btns.grid_columnconfigure(0, weight=1)

        def _copiar():
            texto = grid.get_selected_text()
            if not texto:
                messagebox.showwarning("Selección por celdas", "No hay celdas seleccionadas.", parent=top)
                return
            grid.clipboard_clear()
            grid.clipboard_append(texto)
            messagebox.showinfo("Selección por celdas", "Celdas copiadas al portapapeles.", parent=top)

        ctk.CTkButton(
            btns, text="Copiar selección (Ctrl+C)", command=_copiar,
            fg_color=CEMEX_BLUE, hover_color=CEMEX_BLUE_DARK,
            font=CTkFont("Segoe UI", 12, "bold"), height=34,
        ).grid(row=0, column=0, padx=(0, 6), sticky="ew")

        ctk.CTkButton(
            btns, text="Cerrar", command=top.destroy,
            fg_color=CEMEX_NEUTRAL, hover_color=CEMEX_NEUTRAL_DARK,
            font=CTkFont("Segoe UI", 12, "bold"), height=34,
        ).grid(row=0, column=1, padx=(6, 0), sticky="ew")

    def _update_counter(self):
        self.counter_label.configure(text=f"Filas en la solicitud actual: {len(self.batch_rows)}")

    def clear_fields(self):
        mantener_mp = bool(self.mantener_mp_var.get())
        mantener_flete = bool(self.mantener_flete_var.get())

        # Campos que sobreviven al reset si la casilla correspondiente está activa
        persist_mp = {"mp_vta_m3", "mp_comp_m3"}
        persist_flete = {"flete_vta_m3", "flete_comp_m3"}
        persist = set()
        if mantener_mp:
            persist.update(persist_mp)
        if mantener_flete:
            persist.update(persist_flete)

        for key, var in self.vars.items():
            if key in persist:
                continue
            if key == "modalidad":
                var.set("01")
            elif key == "um_vta":
                var.set("M3")
            elif key == "tipo":
                var.set("ALTA")
            elif key == "tipo_cliente":
                var.set("TERCEROS")
            elif key == "incidencia":
                var.set("NO")
            elif key == "formato_destino":
                var.set("")
            else:
                var.set("")
        for cb in self.formato_checkboxes.values():
            cb.deselect()
        self.pending_selected_index = None

    def on_closing(self):
        if self.conn:
            self.conn.close()
        self.destroy()


if __name__ == "__main__":
    app = App()
    app.protocol("WM_DELETE_WINDOW", app.on_closing)
    app.mainloop()
