import os
import sys
import glob
import time
import threading
import tkinter as tk
from tkinter import ttk, messagebox, simpledialog, font as tkfont
import filtros_cedis
import pandas as pd
import openpyxl

# Importar motor de procesamiento unificado
import matriz_integrada

# =============================================================================
# CONSTANTES DE DISEÑO CORPORATIVO CEMEX
# =============================================================================
CLR_NAVY = "#12356B"
CLR_BLUE_LIGHT = "#245BC4"
CLR_BG = "#EDF2FA"
CLR_ROSE = "#FFF3F4"
CLR_WHITE = "#FFFFFF"
CLR_RED = "#D52E42"
CLR_RED_HOVER = "#B92134"
CLR_TEXT = "#172D50"
CLR_MUTED = "#62728B"
CLR_BORDER = "#DFE6F1"
CLR_SELECTED = "#EAF1FF"
FONT_TITLE = ("Segoe UI Variable Display", 23, "bold")
FONT_SUB = ("Segoe UI", 10)
FONT_BODY = ("Segoe UI", 11)
FONT_BOLD = ("Segoe UI", 11, "bold")
FONT_BTN = ("Segoe UI", 12, "bold")


class RoundedCard(tk.Frame):
    """Contenedor con esquinas suaves; conserva los controles nativos accesibles."""
    def __init__(self, parent, fill=CLR_WHITE, **kwargs):
        super().__init__(parent, bg=CLR_BG, **kwargs)
        self.fill = fill
        self.surface = tk.Canvas(self, highlightthickness=0, bg=CLR_BG, takefocus=0)
        self.surface.place(x=0, y=0, relwidth=1, relheight=1)
        self.body = tk.Frame(self, bg=fill)
        self.body.pack(fill="both", expand=True, padx=16, pady=10)
        self.bind("<Configure>", self.redraw)

    def redraw(self, event):
        w, h, r = event.width - 1, event.height - 1, 12
        self.surface.delete("all")
        self.surface.create_polygon(
            r, 1, w-r, 1, w, 1, w, r, w, h-r, w, h,
            w-r, h, r, h, 1, h, 1, h-r, 1, r, 1, 1,
            smooth=True, splinesteps=24, fill=self.fill, outline=CLR_BORDER)


class LanzadorCEMEXApp:
    def __init__(self, root):
        self.root = root
        self.animation_jobs = {}
        self.motion_enabled = True
        if os.name == 'nt':
            try:
                import ctypes
                enabled = ctypes.c_int(1)
                if ctypes.windll.user32.SystemParametersInfoW(0x1042, 0, ctypes.byref(enabled), 0):
                    self.motion_enabled = bool(enabled.value)
            except (AttributeError, OSError):
                pass
        self.root.bind('<Destroy>', self.detener_animaciones, add='+')
        self.root.title("Generador de Matrices | CEMEX")
        self.root.geometry("820x860")
        self.root.minsize(700, 660)
        self.root.configure(bg=CLR_BG)
        
        # Centrar ventana en pantalla
        self.centrar_ventana(min(820, self.root.winfo_screenwidth() - 64), min(860, self.root.winfo_screenheight() - 80))
        
        self.base_dir = os.path.dirname(os.path.abspath(__file__))
        self.assets_dir = os.path.join(self.base_dir, "assets")
        self.logo_image = None
        try:
            logo = tk.PhotoImage(file=os.path.join(self.assets_dir, "cemex-logo.png"))
            factor = max(1, (logo.width() + 199) // 200)
            self.logo_image = logo.subsample(factor, factor)
            self.window_icon = tk.PhotoImage(file=os.path.join(self.assets_dir, "generador-matrices.png"))
            self.root.iconphoto(True, self.window_icon)
            if os.name == "nt":
                self.root.iconbitmap(os.path.join(self.assets_dir, "generador-matrices.ico"))
        except (tk.TclError, OSError):
            pass  # El texto de marca permite abrir la app si se copió sin los recursos.
        self.source_file = self.detectar_archivo_fuente()
        self.all_cedis = self.cargar_lista_cedis()
        
        # Diccionario para almacenar el estado booleano de cada CEDIS
        self.cedis_vars = {}
        for c in self.all_cedis:
            var = tk.BooleanVar(value=False)
            self.cedis_vars[c] = var
            
        self.checkbox_widgets = {}
        self.is_processing = False
        self.ruta_filtros = filtros_cedis.ruta_filtros()
        try:
            self.filtros = filtros_cedis.cargar(self.ruta_filtros)
        except (OSError, ValueError):
            self.filtros = {'filtros': {}, 'predeterminado': ''}
            self.root.after(0, lambda: messagebox.showwarning('Filtros guardados',
                'No se pudieron leer tus filtros. Puedes seguir seleccionando centros manualmente.', parent=self.root))
        inicial = self.filtros['filtros'].get(self.filtros['predeterminado'], [])
        for c in inicial:
            if c in self.cedis_vars:
                self.cedis_vars[c].set(True)
        
        self.construir_interfaz()
        self.actualizar_contador()
        if inicial:
            self.filtro_nombre.set(self.filtros['predeterminado'])
        self.aparecer_suavemente()

    def detener_animaciones(self, event=None):
        if event is not None and event.widget != self.root:
            return
        for job in self.animation_jobs.values():
            self.root.after_cancel(job)
        self.animation_jobs.clear()

    def animar(self, clave, duracion, actualizar):
        anterior = self.animation_jobs.pop(clave, None)
        if anterior:
            self.root.after_cancel(anterior)
        if not self.motion_enabled:
            actualizar(1.0)
            return
        inicio = time.monotonic()

        def frame():
            self.animation_jobs.pop(clave, None)
            progreso = min(1.0, (time.monotonic() - inicio) / duracion)
            suave = 1 - (1 - progreso) ** 3
            actualizar(suave)
            if progreso < 1:
                self.animation_jobs[clave] = self.root.after(16, frame)
        frame()

    def aparecer_suavemente(self):
        if not self.motion_enabled:
            return
        def aplicar(progreso):
            try:
                self.root.attributes('-alpha', 0.88 + 0.12 * progreso)
            except tk.TclError:
                pass  # La aplicación funciona también sin soporte de transparencia.
        self.animar('entrada', 0.24, aplicar)

    def animar_boton(self, destino):
        if self.is_processing:
            return
        origen = self.btn_generar.winfo_rgb(self.btn_generar.cget('bg'))
        final = self.btn_generar.winfo_rgb(destino)
        def aplicar(progreso):
            canales = [round((a + (b - a) * progreso) / 257) for a, b in zip(origen, final)]
            self.btn_generar.configure(bg='#{:02x}{:02x}{:02x}'.format(*canales))
        self.animar('boton', 0.16, aplicar)

    def centrar_ventana(self, ancho, alto):
        self.root.update_idletasks()
        sw = self.root.winfo_screenwidth()
        sh = self.root.winfo_screenheight()
        x = max(0, (sw - ancho) // 2)
        y = max(0, (sh - alto) // 2)
        self.root.geometry(f"{ancho}x{alto}+{x}+{y}")

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
        families = set(tkfont.families(self.root))
        display = "Segoe UI Variable Display" if "Segoe UI Variable Display" in families else "Segoe UI"
        style = ttk.Style(self.root)
        style.theme_use("clam")
        style.configure("Cedis.TEntry", padding=9, fieldbackground=CLR_BG,
                        bordercolor=CLR_BORDER, lightcolor=CLR_BORDER, darkcolor=CLR_BORDER)
        style.map("Cedis.TEntry", bordercolor=[("focus", CLR_BLUE_LIGHT)])
        style.configure('Filtros.TCombobox', padding=6, bordercolor=CLR_BORDER,
                        arrowcolor=CLR_BLUE_LIGHT, arrowsize=14)
        style.map('Filtros.TCombobox', fieldbackground=[('readonly', CLR_BG)],
                  foreground=[('readonly', CLR_NAVY)], selectbackground=[('readonly', CLR_BG)],
                  selectforeground=[('readonly', CLR_NAVY)])
        style.configure("Cedis.Horizontal.TProgressbar", background=CLR_BLUE_LIGHT,
                        troughcolor=CLR_SELECTED, borderwidth=0)
        style.configure("Vertical.TScrollbar", background=CLR_BORDER, troughcolor=CLR_WHITE,
                        borderwidth=0, arrowsize=12)
        header = tk.Frame(self.root, bg=CLR_BG, padx=24, pady=12)
        header.pack(fill="x")
        brand = tk.Frame(header, bg=CLR_BG)
        brand.pack(fill="x")
        tk.Label(brand, text="Generador de matrices", font=(display, 20, "bold"),
                 fg=CLR_NAVY, bg=CLR_BG).pack(side="left", anchor="center")
        tk.Label(brand, image=self.logo_image, text="CEMEX" if self.logo_image is None else "",
                 font=(display, 13, "bold"), fg=CLR_NAVY, bg=CLR_BG).pack(side="right")
        tk.Label(header, text="Prepara tu matriz de precios, fletes y gobernanza en Excel.",
                 font=FONT_BODY, fg=CLR_MUTED, bg=CLR_BG).pack(anchor="w", pady=(5, 3))

        footer = tk.Frame(self.root, bg=CLR_BG, padx=24, pady=8)
        footer.pack(fill="x", side="bottom")
        self.progress_bar = ttk.Progressbar(footer, mode="indeterminate", style="Cedis.Horizontal.TProgressbar")
        self.lbl_status = tk.Label(footer, text="", font=FONT_SUB, fg=CLR_BLUE_LIGHT, bg=CLR_BG)
        self.btn_generar = tk.Button(footer, text="Generar matriz y dashboard", font=FONT_BTN,
            bg=CLR_RED, fg=CLR_WHITE, activebackground=CLR_RED_HOVER, activeforeground=CLR_WHITE,
            disabledforeground=CLR_WHITE, relief="flat", bd=0, pady=12, cursor="hand2",
            highlightthickness=2, highlightbackground=CLR_BG, highlightcolor=CLR_NAVY,
            command=self.iniciar_generacion)
        self.btn_generar.pack(fill="x")
        self.btn_generar.bind("<Enter>", lambda e: self.animar_boton(CLR_RED_HOVER))
        self.btn_generar.bind("<Leave>", lambda e: self.animar_boton(CLR_RED))
        tk.Label(footer, text="Tu reporte se guardará y abrirá en Excel.",
                 font=FONT_SUB, fg=CLR_MUTED, bg=CLR_BG).pack(pady=(6, 0))

        main = tk.Frame(self.root, bg=CLR_BG, padx=24)
        main.pack(fill="both", expand=True)
        # Se reservan las opciones inferiores antes del área flexible de centros.
        options = tk.Frame(main, bg=CLR_BG)
        options.pack(side="bottom", fill="x", pady=(10, 0))
        options.columnconfigure(0, weight=1, uniform="options")
        options.columnconfigure(1, weight=1, uniform="options")
        vigencia = RoundedCard(options, fill=CLR_SELECTED)
        vigencia.grid(row=0, column=0, sticky="nsew", padx=(0, 5))
        formato = RoundedCard(options, fill=CLR_ROSE)
        formato.grid(row=0, column=1, sticky="nsew", padx=(5, 0))
        self.section_title(vigencia.body, "02", "Vigencia de contratos")
        tk.Label(vigencia.body, text="Contratos de compra · TRAOPE", font=FONT_SUB,
                 bg=CLR_SELECTED, fg=CLR_MUTED).pack(anchor="w", pady=(0, 8))
        self.filtro_traope_var = tk.StringVar(value="2024")
        self.formato_hojas_var = tk.StringVar(value="por_cedis")
        self.option_widgets = []
        for value, text in [("2024", "Desde 2024 · recomendado"), ("2029", "Desde 2029 · estricto")]:
            self.option_radio(vigencia.body, self.filtro_traope_var, value, text)
        self.section_title(formato.body, "03", "Formato de Excel")
        for value, text in [("por_cedis", "Una pestaña por CEDIS"), ("consolidado", "Una hoja consolidada"),
                            ("por_sociedad", "Una pestaña por sociedad"), ("hibrido", "Global + por CEDIS")]:
            self.option_radio(formato.body, self.formato_hojas_var, value, text)

        centers = RoundedCard(main)
        centers.pack(fill="both", expand=True)
        self.section_title(centers.body, "01", "Selecciona tus centros")
        search = tk.Frame(centers.body, bg=CLR_WHITE)
        search.pack(fill="x", pady=(4, 8))
        tk.Label(search, text="Buscar CEDIS", font=FONT_SUB, bg=CLR_WHITE, fg=CLR_MUTED).pack(side="left", padx=(0, 10))
        self.search_var = tk.StringVar()
        self.txt_search = ttk.Entry(search, textvariable=self.search_var, font=FONT_BODY, style="Cedis.TEntry")
        self.txt_search.pack(side="left", fill="x", expand=True, padx=(0, 8))
        self.search_var.trace_add("write", lambda *args: self.filtrar_lista())
        self.root.bind("<Control-f>", lambda e: self.txt_search.focus_set())
        actions = tk.Frame(centers.body, bg=CLR_SELECTED, padx=10, pady=8)
        actions.pack(fill="x", pady=(2, 10))
        self.action_widgets = []
        for label, command in [("Todos", self.marcar_todos), ("Limpiar", self.desmarcar_todos)]:
            outline = tk.Frame(search, bg=CLR_BORDER, bd=0)
            outline.pack(side='left', padx=(0, 6))
            button = tk.Button(outline, text=label, command=command, font=FONT_SUB, relief="flat", bd=0,
                bg=CLR_WHITE, fg=CLR_BLUE_LIGHT, activebackground=CLR_SELECTED, activeforeground=CLR_NAVY,
                padx=10, pady=5, cursor="hand2")
            button.pack(padx=1, pady=1)
            self.action_widgets.append(button)
        self.filtro_nombre = tk.StringVar(value='Filtros guardados')
        self.combo_filtros = ttk.Combobox(actions, textvariable=self.filtro_nombre,
            values=sorted(self.filtros['filtros']), state='readonly', width=17, font=FONT_SUB, style='Filtros.TCombobox')
        self.combo_filtros.pack(side='left', fill='x', expand=True, padx=(0, 10))
        self.combo_filtros.bind('<<ComboboxSelected>>', self.aplicar_filtro)
        for label, command in [('Crear filtro', self.crear_filtro), ('Eliminar', self.eliminar_filtro)]:
            button = tk.Button(actions, text=label, command=command, font=FONT_SUB,
                bg=CLR_NAVY if label == 'Crear filtro' else CLR_ROSE,
                fg=CLR_WHITE if label == 'Crear filtro' else CLR_RED,
                activebackground=CLR_SELECTED, relief='flat', bd=0, padx=8, pady=5, cursor='hand2')
            button.pack(side='left', padx=(0, 4))
            self.action_widgets.append(button)
        self.lbl_count = tk.Label(centers.body, text="", font=FONT_BOLD, fg=CLR_BLUE_LIGHT, bg=CLR_WHITE)
        self.lbl_count.pack(side="bottom", anchor="w", pady=(8, 0))
        list_area = tk.Frame(centers.body, bg=CLR_WHITE)
        list_area.pack(fill="both", expand=True)
        self.canvas = tk.Canvas(list_area, bg=CLR_WHITE, height=170, highlightthickness=0)
        self.scrollbar = ttk.Scrollbar(list_area, orient="vertical", command=self.canvas.yview)
        self.scrollable_frame = tk.Frame(self.canvas, bg=CLR_WHITE)
        self.scrollable_frame.bind("<Configure>", lambda e: self.canvas.configure(scrollregion=self.canvas.bbox("all")))
        self.canvas_window = self.canvas.create_window((0, 0), window=self.scrollable_frame, anchor="nw")
        self.canvas.bind("<Configure>", self.ajustar_ancho_canvas)
        self.canvas.configure(yscrollcommand=self.scrollbar.set)
        self.canvas.pack(side="left", fill="both", expand=True)
        self.scrollbar.pack(side="right", fill="y")
        self.root.bind("<MouseWheel>", self.on_mousewheel)
        self.dibujar_checkboxes(self.all_cedis)

    def section_title(self, parent, number, title):
        background = parent.cget('bg')
        row = tk.Frame(parent, bg=background)
        row.pack(fill="x", pady=(0, 5))
        tk.Label(row, text=number, bg=background, fg=CLR_RED, font=FONT_BOLD).pack(side="left", padx=(0, 8))
        tk.Label(row, text=title, bg=background, fg=CLR_NAVY, font=FONT_BOLD).pack(side="left")

    def option_radio(self, parent, variable, value, text):
        button = tk.Radiobutton(parent, text=text, variable=variable, value=value, font=FONT_SUB,
            bg=parent.cget('bg'), fg=CLR_TEXT, activebackground=CLR_SELECTED, activeforeground=CLR_NAVY,
            selectcolor=CLR_SELECTED, anchor="w", cursor="hand2", pady=0, relief="flat")
        button.pack(fill="x")
        self.option_widgets.append(button)

    def ajustar_ancho_canvas(self, event):
        self.canvas.itemconfig(self.canvas_window, width=event.width)

    def on_mousewheel(self, event):
        if self.canvas.winfo_exists() and str(event.widget).startswith(str(self.canvas)):
            self.canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")

    def dibujar_checkboxes(self, lista_cedis):
        for widget in self.scrollable_frame.winfo_children():
            widget.destroy()
        self.checkbox_widgets.clear()
        self.group_labels = []

        if not lista_cedis:
            tk.Label(self.scrollable_frame, text="No encontramos ese centro. Prueba otro código.",
                     font=FONT_BODY, bg=CLR_WHITE, fg=CLR_MUTED, pady=16).pack(anchor="w")
        for col in range(3):
            self.scrollable_frame.columnconfigure(col, weight=1, uniform="cedis")
        for index, c in enumerate(lista_cedis):
            var = self.cedis_vars[c]
            outline = tk.Frame(self.scrollable_frame, bg=CLR_BORDER, bd=0)
            outline.grid(row=index // 3, column=index % 3, sticky='ew', padx=4, pady=4)
            cb = tk.Checkbutton(
                outline,
                text=f"+  {c}",
                indicatoron=False,
                relief='flat',
                offrelief='flat',
                overrelief='flat',
                borderwidth=0,
                highlightthickness=0,
                highlightbackground=CLR_BORDER,
                highlightcolor=CLR_BLUE_LIGHT,
                variable=var,
                font=FONT_BODY,
                bg=CLR_WHITE,
                fg=CLR_TEXT,
                activebackground=CLR_WHITE,
                selectcolor=CLR_WHITE,
                anchor="w",
                padx=8,
                pady=8,
                cursor="hand2",
                state="disabled" if self.is_processing else "normal",
                command=self.actualizar_contador
            )
            cb.pack(fill='x', padx=1, pady=1)
            cb.bind('<FocusIn>', lambda e, frame=outline: frame.configure(bg=CLR_BLUE_LIGHT))
            cb.bind('<FocusOut>', lambda e: self.refrescar_seleccion())
            self.checkbox_widgets[c] = cb
        self.refrescar_seleccion()
        self.ordenar_seleccionados()

    def ordenar_seleccionados(self):
        for label in self.group_labels:
            label.destroy()
        self.group_labels = []
        row = 0
        for selected, title in [(True, 'SELECCIONADOS'), (False, 'DISPONIBLES')]:
            codes = sorted(c for c in self.checkbox_widgets if self.cedis_vars[c].get() == selected)
            if not codes:
                continue
            label = tk.Label(self.scrollable_frame, text=f'{title}  ·  {len(codes)}',
                font=('Segoe UI', 9), fg=CLR_BLUE_LIGHT if selected else CLR_MUTED, bg=CLR_WHITE)
            label.grid(row=row, column=0, columnspan=3, sticky='w', padx=3, pady=(5, 4))
            self.group_labels.append(label)
            row += 1
            for index, code in enumerate(codes):
                self.checkbox_widgets[code].master.grid_configure(row=row + index // 3, column=index % 3, pady=4, padx=4)
            row += (len(codes) + 2) // 3

    def filtrar_lista(self, event=None):
        query = self.search_var.get().strip().upper()
        if not query:
            lista_filtrada = self.all_cedis
        else:
            lista_filtrada = [c for c in self.all_cedis if query in c]
        self.dibujar_checkboxes(lista_filtrada)
        self.canvas.yview_moveto(0)

    def marcar_todos(self):
        for var in self.cedis_vars.values():
            var.set(True)
        self.actualizar_contador()

    def desmarcar_todos(self):
        for var in self.cedis_vars.values():
            var.set(False)
        self.actualizar_contador()

    def refrescar_seleccion(self):
        for code, widget in self.checkbox_widgets.items():
            selected = self.cedis_vars[code].get()
            widget.master.configure(bg='#ABC3EC' if selected else CLR_BORDER)
            widget.configure(bg=CLR_SELECTED if selected else CLR_BG,
                             text=f'✓  {code}' if selected else f'+  {code}',
                             highlightbackground=CLR_BLUE_LIGHT if selected else CLR_BORDER,
                             fg=CLR_BLUE_LIGHT if selected else CLR_TEXT,
                             activebackground=CLR_SELECTED, selectcolor=CLR_SELECTED,
                             font=FONT_BOLD if selected else FONT_BODY)

    def actualizar_contador(self):
        count = sum(var.get() for var in self.cedis_vars.values())
        self.lbl_count.config(text=f"{count} de {len(self.all_cedis)} centros seleccionados")
        self.refrescar_seleccion()
        self.ordenar_seleccionados()
        self.canvas.yview_moveto(0)
        if hasattr(self, 'filtro_nombre'):
            self.filtro_nombre.set('Selección personalizada')

    def bloquear_opciones(self, disabled):
        state = "disabled" if disabled else "normal"
        for widget in self.option_widgets + self.action_widgets + list(self.checkbox_widgets.values()):
            widget.configure(state=state)
        self.txt_search.configure(state=state)
        self.combo_filtros.configure(state='disabled' if disabled else 'readonly')

    def persistir_filtros(self, datos):
        try:
            filtros_cedis.guardar(self.ruta_filtros, datos)
        except OSError as error:
            messagebox.showerror('No se pudo guardar', str(error), parent=self.root)
            return False
        self.filtros = datos
        self.combo_filtros.configure(values=sorted(datos['filtros']))
        return True

    def crear_filtro(self):
        seleccion = sorted(c for c, var in self.cedis_vars.items() if var.get())
        if not seleccion:
            messagebox.showinfo('Crear filtro', 'Selecciona al menos un centro antes de guardar.', parent=self.root)
            return
        nombre = simpledialog.askstring('Crear filtro',
            'Nombre de esta selección (se recuperará al abrir la aplicación):', parent=self.root)
        if not nombre or not nombre.strip():
            return
        nombre = nombre.strip()[:60]
        if nombre in self.filtros['filtros'] and not messagebox.askyesno('Actualizar filtro',
                f'¿Reemplazar la selección de «{nombre}»?', parent=self.root):
            return
        datos = {'filtros': dict(self.filtros['filtros']), 'predeterminado': nombre}
        datos['filtros'][nombre] = seleccion
        if self.persistir_filtros(datos):
            self.filtro_nombre.set(nombre)

    def aplicar_filtro(self, event=None):
        nombre = self.filtro_nombre.get()
        if nombre not in self.filtros['filtros']:
            return
        seleccion = self.filtros['filtros'][nombre]
        for c, var in self.cedis_vars.items():
            var.set(c in seleccion)
        self.search_var.set('')
        self.actualizar_contador()
        self.filtro_nombre.set(nombre)
        self.persistir_filtros({'filtros': dict(self.filtros['filtros']), 'predeterminado': nombre})
        faltantes = sorted(set(seleccion) - set(self.cedis_vars))
        if faltantes:
            messagebox.showwarning('Centros no disponibles',
                'Estos centros ya no aparecen en el archivo fuente: ' + ', '.join(faltantes), parent=self.root)

    def eliminar_filtro(self):
        nombre = self.filtro_nombre.get()
        if nombre not in self.filtros['filtros']:
            return
        if not messagebox.askyesno('Eliminar filtro', f'¿Eliminar «{nombre}» de tus filtros guardados?', parent=self.root):
            return
        datos = {'filtros': dict(self.filtros['filtros']), 'predeterminado': self.filtros['predeterminado']}
        del datos['filtros'][nombre]
        if datos['predeterminado'] == nombre:
            datos['predeterminado'] = ''
        if self.persistir_filtros(datos):
            self.filtro_nombre.set('Selección personalizada')

    def iniciar_generacion(self):
        if self.is_processing:
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
        job = self.animation_jobs.pop('boton', None)
        if job:
            self.root.after_cancel(job)
        self.bloquear_opciones(True)
        self.btn_generar.config(state="disabled", text="Preparando tu reporte…", bg=CLR_NAVY)
        self.progress_bar.pack(fill="x", pady=(0, 4), before=self.btn_generar)
        self.progress_bar.start(10)
        self.lbl_status.pack(pady=(0, 4), before=self.btn_generar)
        self.lbl_status.config(text=f"Procesando {txt_cedis}...")

        # Ejecutar procesamiento en un hilo secundario para no congelar la UI
        hilo = threading.Thread(target=self.ejecutar_motor, args=(arg_cedis, self.filtro_traope_var.get(), self.formato_hojas_var.get()), daemon=True)
        hilo.start()

    def ejecutar_motor(self, arg_cedis, filtro_sel, fmt_sel):
        out_dir = os.path.join(self.base_dir, "_salidas_integradas")
        
        
        try:
            os.makedirs(out_dir, exist_ok=True)
            # 1. Cargar datos maestros (usa caché rápido)
            data_raw = matriz_integrada.cargar_excel(self.source_file, force_refresh=False)
            
            # 2. Procesar cruce, MOP puro, gobernanza y validaciones con la vigencia seleccionada
            df_matriz, df_mp, df_flete, df_traope, df_contratos = matriz_integrada.procesar_datos(
                data_raw, cedis=arg_cedis, modo_a=True, filtro_traope=filtro_sel
            )
            
            # 3. Construir libro final con Dashboard y Matriz según formato de hojas
            archivo_generado = matriz_integrada.escribir_excel(
                df_matriz, arg_cedis, out_dir, [df_mp, df_flete, df_traope, df_contratos],
                filtro_traope=filtro_sel, formato_hojas=fmt_sel
            )
            
            # Notificar éxito a la UI principal
            self.root.after(0, self.finalizar_exito, archivo_generado)
            
        except Exception as e:
            self.root.after(0, self.finalizar_error, str(e))

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
        self.bloquear_opciones(False)
        self.progress_bar.stop()
        self.progress_bar.pack_forget()
        self.lbl_status.pack_forget()
        self.btn_generar.config(
            state="normal",
            text="Generar matriz y dashboard",
            bg=CLR_RED
        )

# =============================================================================
# ENTRADA PRINCIPAL
# =============================================================================
if __name__ == "__main__":
    if os.name == "nt":
        try:
            import ctypes
            ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID("CEMEX.GeneradorMatrices")
        except (AttributeError, OSError):
            pass
    root = tk.Tk()
    app = LanzadorCEMEXApp(root)
    root.mainloop()
