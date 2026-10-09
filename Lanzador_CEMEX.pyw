import os
import sys
import glob
import time
import json
import threading
import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import pandas as pd
import openpyxl

try:
    import win32com.client
    import pythoncom
    HAS_WIN32COM = True
except ImportError:
    HAS_WIN32COM = False

# Importar motor de procesamiento unificado
import matriz_integrada

# =============================================================================
# CONSTANTES DE DISEÑO CORPORATIVO CEMEX
# =============================================================================
CLR_NAVY        = "#002D72"  # Azul CEMEX Principal
CLR_BLUE_LIGHT  = "#005A9C"  # Azul Secundario
CLR_BG          = "#F1F5F9"  # Fondo gris claro suave
CLR_WHITE       = "#FFFFFF"
CLR_GREEN       = "#28A745"  # Verde Acción
CLR_GREEN_HOVER = "#218838"
CLR_TEXT        = "#1E293B"
CLR_MUTED       = "#64748B"
CLR_BORDER      = "#CBD5E1"
CLR_SNOW_BG     = "#E2E8F0"  # Fondo barra Snowflake
CLR_SNOW_OK     = "#166534"  # Verde texto Snowflake sincronizado
CLR_SNOW_WARN   = "#B45309"  # Ámbar advertencia Snowflake

FONT_TITLE = ("Segoe UI", 12, "bold")
FONT_SUB   = ("Segoe UI", 8, "normal")
FONT_BODY  = ("Segoe UI", 9, "normal")
FONT_BOLD  = ("Segoe UI", 9, "bold")
FONT_BTN   = ("Segoe UI", 10, "bold")

class LanzadorCEMEXApp:
    def __init__(self, root):
        self.root = root
        self.root.title("CEMEX | Sistema Integral de Matriz de Precios y Gobernanza")
        self.root.geometry("600x730")
        self.root.minsize(500, 540)
        self.root.configure(bg=CLR_BG)
        
        # Centrar ventana en pantalla
        self.centrar_ventana(600, 730)
        
        self.base_dir = os.path.dirname(os.path.abspath(__file__))
        self.cfg_compartida = self.cargar_config_compartida()
        self.source_file = self.detectar_archivo_fuente()
        self.all_cedis = self.cargar_lista_cedis()
        
        # Estado de sincronización Snowflake
        self.is_syncing_snowflake = False
        self.auto_sync_snowflake = self.cfg_compartida.get("auto_sync_snowflake", True)
        self.ultimo_sync_str = self.cfg_compartida.get("ultimo_sync_snowflake", "Sin registrar")
        self.ultimo_error_snowflake = ""
        
        # Diccionario para almacenar el estado booleano de cada CEDIS
        self.cedis_vars = {}
        for c in self.all_cedis:
            var = tk.BooleanVar(value=False)
            if c in ["D836", "D838"]:
                var.set(True)
            self.cedis_vars[c] = var
            
        self.checkbox_widgets = {}
        self.is_processing = False
        
        self.construir_interfaz()
        self.actualizar_contador()
        
        self.root.protocol("WM_DELETE_WINDOW", self.al_cerrar_ventana)
        
        # Iniciar sincronización con Snowflake en segundo plano al arrancar
        if self.auto_sync_snowflake:
            self.root.after(400, self.iniciar_sincronizacion_snowflake)

    def centrar_ventana(self, ancho, alto):
        self.root.update_idletasks()
        sw = self.root.winfo_screenwidth()
        sh = self.root.winfo_screenheight()
        x = max(0, (sw - ancho) // 2)
        y = max(0, (sh - alto) // 2)
        self.root.geometry(f"{ancho}x{alto}+{x}+{y}")

    def cargar_config_compartida(self):
        cfg_path = os.path.join(self.base_dir, "config_compartida.json")
        default_dir = os.path.join(self.base_dir, "_salidas_integradas", "Carpeta_Compartida_Champions")
        if os.path.exists(cfg_path):
            try:
                with open(cfg_path, "r", encoding="utf-8") as f:
                    cfg = json.load(f)
                    if "ruta_compartida" in cfg and cfg["ruta_compartida"]:
                        return cfg
            except Exception:
                pass
        return {"ruta_compartida": default_dir, "organizar_subcarpetas": True, "auto_sync_snowflake": True}

    def guardar_config_compartida(self, cfg):
        cfg_path = os.path.join(self.base_dir, "config_compartida.json")
        try:
            with open(cfg_path, "w", encoding="utf-8") as f:
                json.dump(cfg, f, indent=2, ensure_ascii=False)
        except Exception:
            pass

    def obtener_ruta_display(self):
        ruta = self.cfg_compartida.get("ruta_compartida", "")
        if not ruta:
            return "No configurada (clic en Cambiar)"
        if len(ruta) > 55:
            return "..." + ruta[-52:]
        return ruta

    def seleccionar_carpeta_compartida(self):
        dir_actual = self.cfg_compartida.get("ruta_compartida", self.base_dir)
        nueva_ruta = filedialog.askdirectory(
            initialdir=dir_actual if os.path.exists(dir_actual) else self.base_dir,
            title="Seleccionar Carpeta Compartida para Champions (SharePoint / Teams / Red)"
        )
        if nueva_ruta:
            self.cfg_compartida["ruta_compartida"] = os.path.normpath(nueva_ruta)
            self.guardar_config_compartida(self.cfg_compartida)
            if hasattr(self, "lbl_ruta_comp"):
                self.lbl_ruta_comp.config(text=self.obtener_ruta_display())

    def detectar_archivo_fuente(self):
        patron = os.path.join(self.base_dir, "*2026*.xlsm")
        files = glob.glob(patron)
        if files:
            return files[0]
        return os.path.join(self.base_dir, "Extracción de Datos Agregados _ Formatos Para Alta de Ruta (ultimo)_2026.xlsm")

    def cargar_lista_cedis(self):
        try:
            if os.path.exists(self.source_file):
                wb = openpyxl.load_workbook(self.source_file, read_only=True, data_only=True)
                ws = wb['PVTA_MAT VK13']
                headers = next(ws.iter_rows(max_row=1, values_only=True))
                c_idx = headers.index('Centro')
                c_set = set()
                for r in ws.iter_rows(min_row=2, values_only=True):
                    val = r[c_idx]
                    if val:
                        c_set.add(str(val).strip().upper())
                wb.close()
                return sorted(list(c_set))
        except Exception:
            pass
        return ['D836', 'D838', 'DW66', 'D285', 'D847', 'D865', 'D881', 'DW74', 'DW75', 'DW88']

    def construir_interfaz(self):
        # 1. HEADER HERO (ARRIBA)
        header_frame = tk.Frame(self.root, bg=CLR_NAVY, padx=14, pady=8)
        header_frame.pack(fill="x", side="top")
        
        lbl_brand = tk.Label(header_frame, text="CEMEX AGREGADOS", font=("Segoe UI", 8, "bold"), fg="#93C5FD", bg=CLR_NAVY)
        lbl_brand.pack(anchor="w")
        
        lbl_title = tk.Label(header_frame, text="Sistema Integral de Matriz y Gobernanza", font=FONT_TITLE, fg=CLR_WHITE, bg=CLR_NAVY)
        lbl_title.pack(anchor="w")
        
        lbl_sub = tk.Label(header_frame, text="Cálculo automático de MOP %, Desglose de Flete y Dashboard Ejecutivo", font=FONT_SUB, fg="#CBD5E1", bg=CLR_NAVY)
        lbl_sub.pack(anchor="w", pady=(1, 0))

        # 1.1 BANNER DE SINCRONIZACIÓN SNOWFLAKE (SUB-HEADER)
        snow_frame = tk.Frame(self.root, bg=CLR_SNOW_BG, padx=14, pady=4)
        snow_frame.pack(fill="x", side="top")
        
        self.lbl_snow_status = tk.Label(
            snow_frame,
            text=f"❄️ Snowflake: Listo  (Última sinc: {self.ultimo_sync_str})",
            font=("Segoe UI", 8, "normal"),
            fg=CLR_TEXT,
            bg=CLR_SNOW_BG,
            anchor="w"
        )
        self.lbl_snow_status.pack(side="left", fill="x", expand=True)
        self.lbl_snow_status.bind("<Button-1>", self.mostrar_detalle_error_snowflake)

        self.snow_pbar = ttk.Progressbar(snow_frame, mode="indeterminate", length=80)

        self.auto_sync_var = tk.BooleanVar(value=self.auto_sync_snowflake)
        cb_auto = tk.Checkbutton(
            snow_frame,
            text="Auto al abrir",
            variable=self.auto_sync_var,
            font=("Segoe UI", 8, "normal"),
            bg=CLR_SNOW_BG,
            fg="#475569",
            selectcolor=CLR_WHITE,
            cursor="hand2",
            command=self.guardar_opcion_auto_sync
        )
        cb_auto.pack(side="right", padx=(4, 0))

        self.btn_sync_snow = tk.Button(
            snow_frame,
            text="🔄 Sincronizar",
            font=("Segoe UI", 8, "bold"),
            bg=CLR_WHITE,
            fg=CLR_NAVY,
            relief="groove",
            padx=6,
            pady=1,
            cursor="hand2",
            command=self.iniciar_sincronizacion_snowflake
        )
        self.btn_sync_snow.pack(side="right", padx=2)

        # 2. FOOTER FIJO (SIEMPRE VISIBLE ABAJO)
        footer_frame = tk.Frame(self.root, bg=CLR_BG, padx=14, pady=6)
        footer_frame.pack(fill="x", side="bottom")

        self.progress_bar = ttk.Progressbar(footer_frame, mode="indeterminate")
        self.lbl_status = tk.Label(footer_frame, text="", font=FONT_SUB, fg=CLR_BLUE_LIGHT, bg=CLR_BG)

        self.btn_generar = tk.Button(
            footer_frame,
            text="🚀  GENERAR REPORTE LOCAL (EXCEL COMPLETO)",
            font=FONT_BTN,
            bg=CLR_GREEN,
            fg=CLR_WHITE,
            activebackground=CLR_GREEN_HOVER,
            activeforeground=CLR_WHITE,
            relief="flat",
            pady=6,
            cursor="hand2",
            command=self.iniciar_generacion
        )
        self.btn_generar.pack(fill="x", pady=(0, 4))

        self.btn_publicar = tk.Button(
            footer_frame,
            text="📤  PUBLICAR MATRICES A CHAMPIONS (CARPETA COMPARTIDA)",
            font=FONT_BTN,
            bg=CLR_NAVY,
            fg=CLR_WHITE,
            activebackground="#001D4A",
            activeforeground=CLR_WHITE,
            relief="flat",
            pady=6,
            cursor="hand2",
            command=self.iniciar_publicacion_champions
        )
        self.btn_publicar.pack(fill="x")

        # 3. CONTENEDOR PRINCIPAL (CENTRO CON AUTO-AJUSTE)
        main_frame = tk.Frame(self.root, bg=CLR_BG, padx=14, pady=6)
        main_frame.pack(fill="both", expand=True)

        # 4. BUSCADOR Y ACCIONES RÁPIDAS
        lbl_busc = tk.Label(main_frame, text="🔍 Buscar y seleccionar centros (CEDIS):", font=FONT_BOLD, fg=CLR_TEXT, bg=CLR_BG)
        lbl_busc.pack(anchor="w", pady=(0, 2))
        
        search_box = tk.Frame(main_frame, bg=CLR_BG)
        search_box.pack(fill="x", pady=(0, 4))
        
        self.txt_search = ttk.Entry(search_box, font=FONT_BODY)
        self.txt_search.pack(side="left", fill="x", expand=True, padx=(0, 6))
        self.txt_search.bind("<KeyRelease>", self.filtrar_lista)

        btn_all = tk.Button(search_box, text="☑ Todos", font=FONT_SUB, bg=CLR_WHITE, fg=CLR_TEXT, relief="groove", command=self.marcar_todos, padx=6, pady=1, cursor="hand2")
        btn_all.pack(side="left", padx=2)
        
        btn_none = tk.Button(search_box, text="☐ Ninguno", font=FONT_SUB, bg=CLR_WHITE, fg=CLR_TEXT, relief="groove", command=self.desmarcar_todos, padx=6, pady=1, cursor="hand2")
        btn_none.pack(side="left", padx=2)
        
        btn_demo = tk.Button(search_box, text="🎯 Demo", font=FONT_SUB, bg=CLR_WHITE, fg=CLR_BLUE_LIGHT, relief="groove", command=self.marcar_demo, padx=6, pady=1, cursor="hand2")
        btn_demo.pack(side="left", padx=2)

        # 5. LISTA SCROLLABLE DE CHECKBOXES
        list_card = tk.Frame(main_frame, bg=CLR_WHITE, bd=1, relief="solid", highlightthickness=0)
        list_card.pack(fill="both", expand=True, pady=(2, 4))
        
        self.canvas = tk.Canvas(list_card, bg=CLR_WHITE, height=130, highlightthickness=0)
        self.scrollbar = ttk.Scrollbar(list_card, orient="vertical", command=self.canvas.yview)
        self.scrollable_frame = tk.Frame(self.canvas, bg=CLR_WHITE)

        self.scrollable_frame.bind(
            "<Configure>",
            lambda e: self.canvas.configure(scrollregion=self.canvas.bbox("all"))
        )

        self.canvas_window = self.canvas.create_window((0, 0), window=self.scrollable_frame, anchor="nw")
        self.canvas.bind('<Configure>', self.ajustar_ancho_canvas)
        self.canvas.configure(yscrollcommand=self.scrollbar.set)

        self.canvas.pack(side="left", fill="both", expand=True)
        self.scrollbar.pack(side="right", fill="y")
        self.canvas.bind_all("<MouseWheel>", self.on_mousewheel)

        self.dibujar_checkboxes(self.all_cedis)

        # 6. CONTADOR DE SELECCIÓN
        self.lbl_count = tk.Label(main_frame, text="", font=FONT_SUB, fg=CLR_MUTED, bg=CLR_BG)
        self.lbl_count.pack(anchor="w", pady=(0, 4))

        # 7. FILTRO DE CONTRATOS DE COMPRA (TRAOPE) EN FORMATO COMPACTO HORIZONTAL
        frame_opc = tk.LabelFrame(main_frame, text="  Vigencia Contratos Compra (TRAOPE):  ", font=FONT_BOLD, fg=CLR_TEXT, bg=CLR_BG, padx=8, pady=2, relief="groove")
        frame_opc.pack(fill="x", pady=(0, 4))
        
        self.filtro_traope_var = tk.StringVar(value="2024")
        
        rb_2024 = tk.Radiobutton(
            frame_opc,
            text="Vigencia >= 2024 [Recomendado]",
            variable=self.filtro_traope_var,
            value="2024",
            font=FONT_BODY,
            bg=CLR_BG,
            fg=CLR_TEXT,
            activebackground=CLR_BG,
            selectcolor=CLR_WHITE,
            cursor="hand2"
        )
        rb_2024.pack(side="left", padx=(0, 16), pady=1)
        
        rb_2029 = tk.Radiobutton(
            frame_opc,
            text="Vigencia >= 2029 [Modo Estricto]",
            variable=self.filtro_traope_var,
            value="2029",
            font=FONT_BODY,
            bg=CLR_BG,
            fg=CLR_TEXT,
            activebackground=CLR_BG,
            selectcolor=CLR_WHITE,
            cursor="hand2"
        )
        rb_2029.pack(side="left", padx=0, pady=1)

        # 8. ORGANIZACIÓN DE HOJAS / PESTAÑAS EN EXCEL (CUADRÍCULA 2x2 COMPACTA)
        frame_fmt = tk.LabelFrame(main_frame, text="  Organización de Pestañas en Excel:  ", font=FONT_BOLD, fg=CLR_TEXT, bg=CLR_BG, padx=8, pady=2, relief="groove")
        frame_fmt.pack(fill="x", pady=(0, 2))
        
        self.formato_hojas_var = tk.StringVar(value="por_cedis")
        
        grid_fmt = tk.Frame(frame_fmt, bg=CLR_BG)
        grid_fmt.pack(fill="x", expand=True)

        rb_fmt1 = tk.Radiobutton(
            grid_fmt,
            text="📑 Pestaña x CEDIS [Recomendado]",
            variable=self.formato_hojas_var,
            value="por_cedis",
            font=FONT_BODY,
            bg=CLR_BG,
            fg=CLR_TEXT,
            activebackground=CLR_BG,
            selectcolor=CLR_WHITE,
            cursor="hand2"
        )
        rb_fmt1.grid(row=0, column=0, sticky="w", padx=(0, 10), pady=1)
        
        rb_fmt2 = tk.Radiobutton(
            grid_fmt,
            text="📊 1 Sola Hoja Consolidada",
            variable=self.formato_hojas_var,
            value="consolidado",
            font=FONT_BODY,
            bg=CLR_BG,
            fg=CLR_TEXT,
            activebackground=CLR_BG,
            selectcolor=CLR_WHITE,
            cursor="hand2"
        )
        rb_fmt2.grid(row=0, column=1, sticky="w", pady=1)

        rb_fmt3 = tk.Radiobutton(
            grid_fmt,
            text="🏢 Pestaña x Sociedad (7100 / 7180-7277 Terceros)",
            variable=self.formato_hojas_var,
            value="por_sociedad",
            font=FONT_BODY,
            bg=CLR_BG,
            fg=CLR_TEXT,
            activebackground=CLR_BG,
            selectcolor=CLR_WHITE,
            cursor="hand2"
        )
        rb_fmt3.grid(row=1, column=0, sticky="w", padx=(0, 10), pady=1)

        rb_fmt4 = tk.Radiobutton(
            grid_fmt,
            text="🌟 Híbrido (Global + x CEDIS)",
            variable=self.formato_hojas_var,
            value="hibrido",
            font=FONT_BODY,
            bg=CLR_BG,
            fg=CLR_TEXT,
            activebackground=CLR_BG,
            selectcolor=CLR_WHITE,
            cursor="hand2"
        )
        rb_fmt4.grid(row=1, column=1, sticky="w", pady=1)

        # 9. ENTREGABLES A GENERAR (MASTER, CLIENTE FINAL, POWER BI)
        frame_ent = tk.LabelFrame(main_frame, text="  Entregables y Reportes a Generar:  ", font=FONT_BOLD, fg=CLR_TEXT, bg=CLR_BG, padx=8, pady=2, relief="groove")
        frame_ent.pack(fill="x", pady=(2, 2))
        
        self.gen_master_var = tk.BooleanVar(value=True)
        self.gen_cliente_var = tk.BooleanVar(value=True)
        self.gen_pbi_var = tk.BooleanVar(value=True)
        
        cb_master = tk.Checkbutton(
            frame_ent, text="📘 Versión Master (Interna)", variable=self.gen_master_var,
            font=FONT_BODY, bg=CLR_BG, fg=CLR_TEXT, selectcolor=CLR_WHITE, cursor="hand2"
        )
        cb_master.pack(side="left", padx=(0, 6), pady=1)
        
        cb_cli = tk.Checkbutton(
            frame_ent, text="📗 Versión Cliente Final", variable=self.gen_cliente_var,
            font=FONT_BODY, bg=CLR_BG, fg=CLR_TEXT, selectcolor=CLR_WHITE, cursor="hand2"
        )
        cb_cli.pack(side="left", padx=6, pady=1)
        
        cb_pbi = tk.Checkbutton(
            frame_ent, text="📊 Power BI", variable=self.gen_pbi_var,
            font=FONT_BODY, bg=CLR_BG, fg=CLR_TEXT, selectcolor=CLR_WHITE, cursor="hand2"
        )
        cb_pbi.pack(side="left", padx=6, pady=1)

        # 10. CARPETA COMPARTIDA PARA CHAMPIONS (SHAREPOINT / TEAMS / RED)
        frame_comp = tk.LabelFrame(main_frame, text="  📁 Carpeta Compartida Champions (SharePoint / Teams):  ", font=FONT_BOLD, fg=CLR_TEXT, bg=CLR_BG, padx=8, pady=3, relief="groove")
        frame_comp.pack(fill="x", pady=(2, 2))
        
        self.lbl_ruta_comp = tk.Label(frame_comp, text=self.obtener_ruta_display(), font=FONT_SUB, fg=CLR_MUTED, bg=CLR_BG, anchor="w")
        self.lbl_ruta_comp.pack(side="left", fill="x", expand=True, padx=(0, 4))
        
        btn_sel_comp = tk.Button(frame_comp, text="📂 Cambiar", font=FONT_SUB, bg=CLR_WHITE, fg=CLR_TEXT, relief="groove", command=self.seleccionar_carpeta_compartida, padx=6, pady=1, cursor="hand2")
        btn_sel_comp.pack(side="right")

    def ajustar_ancho_canvas(self, event):
        self.canvas.itemconfig(self.canvas_window, width=event.width)

    def on_mousewheel(self, event):
        if self.canvas.winfo_exists():
            self.canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")

    def dibujar_checkboxes(self, lista_cedis):
        for widget in self.scrollable_frame.winfo_children():
            widget.destroy()
        self.checkbox_widgets.clear()

        for c in lista_cedis:
            if c not in self.cedis_vars:
                self.cedis_vars[c] = tk.BooleanVar(value=False)
            var = self.cedis_vars[c]
            cb = tk.Checkbutton(
                self.scrollable_frame,
                text=f"  CEDIS {c}",
                variable=var,
                font=FONT_BODY,
                bg=CLR_WHITE,
                fg=CLR_TEXT,
                activebackground=CLR_WHITE,
                selectcolor=CLR_WHITE,
                anchor="w",
                padx=8,
                pady=2,
                command=self.actualizar_contador
            )
            cb.pack(fill="x", anchor="w")
            self.checkbox_widgets[c] = cb

    def filtrar_lista(self, event=None):
        query = self.txt_search.get().strip().upper()
        if not query:
            lista_filtrada = self.all_cedis
        else:
            lista_filtrada = [c for c in self.all_cedis if query in c]
        self.dibujar_checkboxes(lista_filtrada)

    def marcar_todos(self):
        for var in self.cedis_vars.values():
            var.set(True)
        self.actualizar_contador()

    def desmarcar_todos(self):
        for var in self.cedis_vars.values():
            var.set(False)
        self.actualizar_contador()

    def marcar_demo(self):
        for c, var in self.cedis_vars.items():
            if c in ["D836", "D838"]:
                var.set(True)
            else:
                var.set(False)
        self.actualizar_contador()

    def actualizar_contador(self):
        total_marcados = sum(1 for var in self.cedis_vars.values() if var.get())
        total_total = len(self.all_cedis)
        if total_marcados == total_total:
            self.lbl_count.config(text=f"✔ Se procesarán TODOS los {total_total} CEDIS del país")
        else:
            self.lbl_count.config(text=f"Centros seleccionados: {total_marcados} de {total_total}")

    # =========================================================================
    # LÓGICA DE SINCRONIZACIÓN AUTOMÁTICA CON SNOWFLAKE (SEGUNDO PLANO)
    # =========================================================================
    def guardar_opcion_auto_sync(self):
        self.auto_sync_snowflake = self.auto_sync_var.get()
        self.cfg_compartida["auto_sync_snowflake"] = self.auto_sync_snowflake
        self.guardar_config_compartida(self.cfg_compartida)

    def iniciar_sincronizacion_snowflake(self):
        if self.is_syncing_snowflake:
            return
        if not HAS_WIN32COM:
            self.lbl_snow_status.config(text="⚠️ Módulo win32com no disponible para automatizar Excel.", fg=CLR_SNOW_WARN)
            return
        if not os.path.exists(self.source_file):
            self.lbl_snow_status.config(text="⚠️ Archivo fuente .xlsm no encontrado.", fg=CLR_SNOW_WARN)
            return

        self.is_syncing_snowflake = True
        self.btn_sync_snow.config(state="disabled", text="⏳ Sincronizando...")
        self.lbl_snow_status.config(text="🔄 Conectando a Snowflake y actualizando consultas en segundo plano...", fg=CLR_BLUE_LIGHT)
        self.snow_pbar.pack(side="left", padx=6, before=self.btn_sync_snow)
        self.snow_pbar.start(10)

        hilo = threading.Thread(target=self._hilo_sincronizar_snowflake, daemon=True)
        hilo.start()

    def mostrar_detalle_error_snowflake(self, event=None):
        if self.ultimo_error_snowflake:
            messagebox.showinfo(
                "Diagnóstico de Conexión Snowflake",
                f"Detalle técnico reportado por Excel / Snowflake:\n\n{self.ultimo_error_snowflake}\n\n"
                f"📌 Motivos habituales:\n"
                f"• Equipo fuera de la red CEMEX / VPN corporativa desconectada.\n"
                f"• Credenciales de Snowflake pendientes de autenticación en este equipo.\n"
                f"• Archivo .xlsm bloqueado o abierto por otra aplicación.\n\n"
                f"El sistema continúa operando con la información local preexistente sin interrupción."
            )

    def _hilo_sincronizar_snowflake(self):
        if HAS_WIN32COM:
            try:
                pythoncom.CoInitialize()
            except Exception:
                pass
        try:
            excel = win32com.client.Dispatch("Excel.Application")
            excel.Visible = False
            excel.DisplayAlerts = False
            excel.ScreenUpdating = False
            excel.EnableEvents = False
            excel.AskToUpdateLinks = False
            try:
                excel.AutomationSecurity = 1  # msoAutomationSecurityLow
            except Exception:
                pass

            wb = None
            try:
                wb = excel.Workbooks.Open(
                    Filename=os.path.abspath(self.source_file),
                    UpdateLinks=0,
                    ReadOnly=False,
                    IgnoreReadOnlyRecommended=True
                )
                for i in range(1, wb.Connections.Count + 1):
                    conn = wb.Connections.Item(i)
                    if hasattr(conn, "OLEDBConnection"):
                        conn.OLEDBConnection.BackgroundQuery = False

                wb.RefreshAll()
                excel.CalculateUntilAsyncQueriesDone()
                wb.Save()
            finally:
                if wb:
                    try:
                        wb.Close(SaveChanges=False)
                    except Exception:
                        pass
                try:
                    excel.Quit()
                except Exception:
                    pass

            # Invalidar archivo de caché para forzar re-lectura limpia
            cache_file = self.source_file + ".mci_cache_.pkl"
            if os.path.exists(cache_file):
                try:
                    os.remove(cache_file)
                except Exception:
                    pass

            self.root.after(0, self._finalizar_sincronizacion_snowflake, True, "")
        except Exception as e:
            self.root.after(0, self._finalizar_sincronizacion_snowflake, False, str(e))
        finally:
            if HAS_WIN32COM:
                try:
                    pythoncom.CoUninitialize()
                except Exception:
                    pass

    def _finalizar_sincronizacion_snowflake(self, exito, error_msg):
        self.is_syncing_snowflake = False
        self.snow_pbar.stop()
        self.snow_pbar.pack_forget()
        self.btn_sync_snow.config(state="normal", text="🔄 Sincronizar")

        if exito:
            self.ultimo_error_snowflake = ""
            hora_act = time.strftime("%H:%M:%S")
            self.ultimo_sync_str = hora_act
            self.cfg_compartida["ultimo_sync_snowflake"] = hora_act
            self.guardar_config_compartida(self.cfg_compartida)
            
            # Recargar lista de CEDIS desde el archivo actualizado
            self.recargar_cedis_tras_sincronizacion()
            self.lbl_snow_status.config(
                text=f"✅ Snowflake sincronizado ({hora_act}) | {len(self.all_cedis)} CEDIS detectados",
                fg=CLR_SNOW_OK,
                cursor=""
            )
        else:
            self.ultimo_error_snowflake = error_msg
            try:
                log_path = os.path.join(self.base_dir, "snowflake_sync.log")
                with open(log_path, "a", encoding="utf-8") as f:
                    f.write(f"[{time.strftime('%Y-%m-%d %H:%M:%S')}] ERROR: {error_msg}\n")
            except Exception:
                pass

            self.lbl_snow_status.config(
                text="⚠️ Sincronización Snowflake no completada (clic para ver detalle)",
                fg=CLR_SNOW_WARN,
                cursor="hand2"
            )

    def recargar_cedis_tras_sincronizacion(self):
        try:
            nuevos_cedis = self.cargar_lista_cedis()
            if nuevos_cedis:
                for c in nuevos_cedis:
                    if c not in self.cedis_vars:
                        self.cedis_vars[c] = tk.BooleanVar(value=True if c in ["D836", "D838"] else False)
                self.all_cedis = nuevos_cedis
                self.filtrar_lista()
                self.actualizar_contador()
        except Exception:
            pass

    def al_cerrar_ventana(self):
        self.root.destroy()

    # =========================================================================
    # GENERACIÓN Y PUBLICACIÓN
    # =========================================================================
    def iniciar_generacion(self):
        if self.is_processing:
            return

        if self.is_syncing_snowflake:
            proceder = messagebox.askyesno(
                "Sincronización Snowflake en Curso",
                "Snowflake se está actualizando en segundo plano para obtener los datos más recientes.\n\n"
                "¿Desea continuar de todas formas con los datos locales existentes?\n\n"
                "(Elija 'No' para esperar unos segundos a que termine la sincronización)."
            )
            if not proceder:
                return

        seleccionados = [c for c, var in self.cedis_vars.items() if var.get()]
        if not seleccionados:
            messagebox.showwarning(
                "Selección Requerida",
                "Por favor seleccione al menos un CEDIS de la lista o presione 'Marcar Todos'."
            )
            return

        if len(seleccionados) == len(self.all_cedis):
            arg_cedis = None  # Indica procesar todos
            txt_cedis = "TODOS LOS CEDIS"
        else:
            arg_cedis = seleccionados
            txt_cedis = ", ".join(seleccionados[:4]) + (f" (+{len(seleccionados)-4} más)" if len(seleccionados)>4 else "")

        # Bloquear UI y mostrar barra de progreso
        self.is_processing = True
        self.btn_generar.config(state="disabled", text="⏳  PROCESANDO DATOS...", bg=CLR_MUTED)
        self.btn_publicar.config(state="disabled")
        self.progress_bar.pack(fill="x", pady=(0, 4), before=self.btn_generar)
        self.progress_bar.start(10)
        self.lbl_status.pack(pady=(0, 4), before=self.btn_generar)
        self.lbl_status.config(text=f"Procesando {txt_cedis}...")

        # Ejecutar procesamiento en un hilo secundario para no congelar la UI
        hilo = threading.Thread(target=self.ejecutar_motor, args=(arg_cedis,), daemon=True)
        hilo.start()

    def ejecutar_motor(self, arg_cedis):
        out_dir = os.path.join(self.base_dir, "_salidas_integradas")
        os.makedirs(out_dir, exist_ok=True)
        
        filtro_sel = self.filtro_traope_var.get()
        fmt_sel = self.formato_hojas_var.get()
        g_master = self.gen_master_var.get()
        g_cli = self.gen_cliente_var.get()
        g_pbi = self.gen_pbi_var.get()
        
        try:
            # 1. Cargar datos maestros (usa caché rápido)
            data_raw = matriz_integrada.cargar_excel(self.source_file, force_refresh=False)
            
            # 2. Procesar cruce, MOP puro, gobernanza y validaciones con la vigencia seleccionada
            df_matriz, df_mp, df_flete, df_traope, df_contratos = matriz_integrada.procesar_datos(
                data_raw, cedis=arg_cedis, modo_a=True, filtro_traope=filtro_sel
            )
            
            # 3. Construir libro final Master, Cliente Final y Power BI
            archivo_generado = matriz_integrada.escribir_excel(
                df_matriz, arg_cedis, out_dir, [df_mp, df_flete, df_traope, df_contratos],
                filtro_traope=filtro_sel, formato_hojas=fmt_sel,
                generar_master=g_master, generar_cliente=g_cli, exportar_powerbi=g_pbi
            )
            
            # Notificar éxito a la UI principal
            self.root.after(0, self.finalizar_exito, archivo_generado)
            
        except Exception as e:
            self.root.after(0, self.finalizar_error, str(e))

    def iniciar_publicacion_champions(self):
        if self.is_processing:
            return

        if self.is_syncing_snowflake:
            proceder = messagebox.askyesno(
                "Sincronización Snowflake en Curso",
                "Snowflake se está actualizando en segundo plano para obtener los datos más recientes.\n\n"
                "¿Desea continuar de todas formas con los datos locales existentes?\n\n"
                "(Elija 'No' para esperar unos segundos a que termine la sincronización)."
            )
            if not proceder:
                return

        seleccionados = [c for c, var in self.cedis_vars.items() if var.get()]
        if not seleccionados:
            messagebox.showwarning(
                "Selección Requerida",
                "Por favor seleccione al menos un CEDIS de la lista para publicar a los Champions, o presione '☑ Todos'."
            )
            return

        ruta_comp = self.cfg_compartida.get("ruta_compartida", "")
        if not ruta_comp:
            messagebox.showinfo(
                "Configurar Carpeta Compartida",
                "Por favor seleccione la carpeta compartida (SharePoint, Teams o Red) donde se publicarán los archivos para los Champions."
            )
            self.seleccionar_carpeta_compartida()
            ruta_comp = self.cfg_compartida.get("ruta_compartida", "")
            if not ruta_comp:
                return

        # Si no existe la carpeta, crearla o verificar acceso
        try:
            os.makedirs(ruta_comp, exist_ok=True)
        except Exception as e:
            messagebox.showerror("Error de Acceso", f"No se pudo acceder o crear la carpeta compartida:\n{ruta_comp}\n\nDetalle: {e}")
            return

        if len(seleccionados) == len(self.all_cedis):
            arg_cedis = None
            txt_cedis = f"TODOS LOS CEDIS ({len(self.all_cedis)} centros)"
        else:
            arg_cedis = seleccionados
            txt_cedis = f"{len(seleccionados)} CEDIS seleccionados"

        confirm = messagebox.askyesno(
            "Confirmar Publicación a Champions",
            f"Se generarán matrices individuales en versión 'Cliente Final' (con costos TRAOPE y márgenes protegidos) para:\n\n"
            f"• Centros: {txt_cedis}\n"
            f"• Destino: {ruta_comp}\n\n"
            f"¿Desea iniciar la publicación ahora?"
        )
        if not confirm:
            return

        # Bloquear UI y mostrar barra de progreso
        self.is_processing = True
        self.btn_generar.config(state="disabled")
        self.btn_publicar.config(state="disabled", text="⏳  PUBLICANDO A CHAMPIONS...", bg=CLR_MUTED)
        self.progress_bar.pack(fill="x", pady=(0, 4), before=self.btn_generar)
        self.progress_bar.start(10)
        self.lbl_status.pack(pady=(0, 4), before=self.btn_generar)
        self.lbl_status.config(text=f"Publicando {txt_cedis} en carpeta compartida...")

        hilo = threading.Thread(target=self.ejecutar_publicacion, args=(arg_cedis, ruta_comp), daemon=True)
        hilo.start()

    def ejecutar_publicacion(self, arg_cedis, ruta_comp):
        filtro_sel = self.filtro_traope_var.get()
        try:
            data_raw = matriz_integrada.cargar_excel(self.source_file, force_refresh=False)
            df_matriz, _, _, _, _ = matriz_integrada.procesar_datos(
                data_raw, cedis=arg_cedis, modo_a=True, filtro_traope=filtro_sel
            )
            res = matriz_integrada.publicar_matrices_champions(
                df_matriz, ruta_comp, cedis_list=arg_cedis, subcarpetas=True
            )
            self.root.after(0, self.finalizar_publicacion_exito, res)
        except Exception as e:
            self.root.after(0, self.finalizar_error, str(e))

    def finalizar_publicacion_exito(self, res):
        self.progreso_terminado()
        ruta = res['ruta']
        total_c = res['total_cedis']
        total_r = res['total_rutas']
        
        abrir = messagebox.askyesno(
            "¡Publicación Completada con Éxito!",
            f"Se han publicado las matrices para los Champions en la carpeta compartida.\n\n"
            f"• Centros actualizados: {total_c} CEDIS\n"
            f"• Total de rutas distribuidas: {total_r}\n"
            f"• Formato: Cliente Final (Costos TRAOPE y márgenes internos protegidos)\n"
            f"• Ubicación: {ruta}\n\n"
            f"¿Desea abrir la carpeta compartida ahora?"
        )
        if abrir:
            try:
                os.startfile(ruta)
            except Exception:
                pass

    def finalizar_exito(self, archivo_generado):
        self.progreso_terminado()
        
        # Abrir automáticamente el archivo generado en Microsoft Excel
        if archivo_generado and os.path.exists(archivo_generado):
            try:
                os.startfile(archivo_generado)
            except Exception:
                pass
            
            nombre_arc = os.path.basename(archivo_generado)
            messagebox.showinfo(
                "¡Proceso Completado!",
                f"Matriz y Dashboard generados exitosamente.\n\n"
                f"Archivo: {nombre_arc}\n\n"
                f"Se ha abierto el reporte en Microsoft Excel."
            )
        else:
            messagebox.showinfo(
                "Proceso Terminado",
                "Proceso completado. Los reportes se guardaron en la carpeta _salidas_integradas."
            )

    def finalizar_error(self, mensaje_error):
        self.progreso_terminado()
        messagebox.showerror(
            "Error en Generación",
            f"Ocurrió un inconveniente durante el cálculo:\n\n{mensaje_error}"
        )

    def progreso_terminado(self):
        self.is_processing = False
        self.progress_bar.stop()
        self.progress_bar.pack_forget()
        self.lbl_status.pack_forget()
        self.btn_generar.config(
            state="normal",
            text="🚀  GENERAR REPORTE LOCAL (EXCEL COMPLETO)",
            bg=CLR_GREEN
        )
        self.btn_publicar.config(
            state="normal",
            text="📤  PUBLICAR MATRICES A CHAMPIONS (CARPETA COMPARTIDA)",
            bg=CLR_NAVY
        )

# =============================================================================
# ENTRADA PRINCIPAL
# =============================================================================
if __name__ == "__main__":
    root = tk.Tk()
    app = LanzadorCEMEXApp(root)
    root.mainloop()
