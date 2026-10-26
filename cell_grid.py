"""CeldaGrid: grilla editable de solo lectura con selección por celdas.

Componente autocontenido (basado en tk.Canvas) que replica la selección de
celdas estilo Excel/SAP, superando la limitación de ttk.Treeview de seleccionar
únicamente filas completas.

Características:
  - Resaltado real a nivel de celda individual (fondo + borde).
  - Clic simple: selecciona una sola celda.
  - Ctrl + clic: agrega/quita una celda individual.
  - Shift + clic: rango rectangular desde la celda ancla.
  - Arrastrar (B1-Motion): rango rectangular continuo.
  - Flechas: mueven la selección de celda en celda.
  - Shift + flechas: extienden la selección horizontal/vertical.
  - Ctrl + C: copia las celdas seleccionadas (ordenadas fila/columna,
    separadas por tabulaciones y saltos de línea) al portapapeles.

No crea un widget por celda (usa Canvas + virtualización de filas visibles),
por lo que soporta miles de filas sin degradar el rendimiento. El texto de cada
celda se trunca al ancho real de su columna (con elipsis), evitando que se
desborde sobre las columnas vecinas.
"""

import tkinter as tk
import tkinter.font as tkfont


class CellGrid(tk.Canvas):
    def __init__(self, master, headers=None, rows=None,
                 col_widths=None, row_height=28, col_width=110,
                 header_height=30, header_bg="#023185", header_fg="#FFFFFF",
                 cell_bg="#FFFFFF", alt_bg="#F4F6F9",
                 sel_bg="#BBD3F5", sel_border="#023185",
                 grid_color="#DDE1E6", text_color="#1A1A1A",
                 font=("Segoe UI", 11), header_font=("Segoe UI", 11, "bold"),
                 cell_pad=6, **kwargs):
        kwargs.setdefault("background", cell_bg)
        kwargs.setdefault("highlightthickness", 0)
        super().__init__(master, **kwargs)

        self.headers = list(headers or [])
        self.rows = [list(r) for r in (rows or [])]
        self._row_height = row_height
        self._default_col_width = col_width
        self._header_height = header_height
        self._header_bg = header_bg
        self._header_fg = header_fg
        self._cell_bg = cell_bg
        self._alt_bg = alt_bg
        self._sel_bg = sel_bg
        self._sel_border = sel_border
        self._grid_color = grid_color
        self._text_color = text_color
        self._font = font
        self._header_font = header_font
        self._cell_pad = cell_pad

        # Anchuros por columna (índice -> píxeles). Si no se indica, todas
        # usan `col_width`; si la lista es más corta, se rellena con el default.
        self._col_widths = list(col_widths or [])

        # Fuentes Tk para medir texto de forma precisa.
        self._tk_font = tkfont.Font(font=font)
        self._tk_header_font = tkfont.Font(font=header_font)

        # Selección por celdas: set de (row_idx, col_idx).
        self._selected = set()
        self._anchor = None   # (row, col) usado por Shift/arrastre.
        self._cursor = [0, 0]  # celda activa para navegación con flechas.

        # Desplazamiento (píxeles) del área de datos.
        self._scroll_x = 0
        self._scroll_y = 0

        # Barras de scroll.
        self._vsb = tk.Scrollbar(master, orient="vertical", command=self._on_vscroll)
        self._hsb = tk.Scrollbar(master, orient="horizontal", command=self._on_hscroll)
        self._vsb.grid(row=0, column=1, sticky="ns")
        self._hsb.grid(row=1, column=0, sticky="ew")
        self.grid(row=0, column=0, sticky="nsew")

        self._bind_events()
        self._update_scrollregion()
        self._render()

    # ------------------------------------------------------------------
    # API pública
    # ------------------------------------------------------------------
    def set_data(self, headers, rows, col_widths=None):
        """Reemplaza encabezados, filas y anchos, limpiando la selección."""
        self.headers = list(headers)
        self.rows = [list(r) for r in rows]
        if col_widths is not None:
            self._col_widths = list(col_widths)
        self._selected.clear()
        self._anchor = None
        self._cursor = [0, 0]
        self._update_scrollregion()
        self._render()

    def get_selected_cells(self):
        """Retorna lista de (row_idx, col_idx, valor) ordenada por fila/columna."""
        return sorted(
            ((r, c, self._cell_value(r, c)) for r, c in self._selected),
            key=lambda t: (t[0], t[1]),
        )

    def get_selected_text(self):
        """Texto de las celdas seleccionadas, alineado por el rectángulo mínimo
        que las contiene (filas y columnas contiguas), dejando celdas vacías
        donde no hay selección, para que al pegar en Excel se conserve la
        estructura."""
        celdas = self.get_selected_cells()
        if not celdas:
            return ""
        por_fila = {}
        for r, c, v in celdas:
            por_fila.setdefault(r, {})[c] = "" if v is None else str(v)

        min_r = min(por_fila)
        max_r = max(por_fila)
        min_c = min(c for fila in por_fila.values() for c in fila)
        max_c = max(c for fila in por_fila.values() for c in fila)

        lineas = []
        for r in range(min_r, max_r + 1):
            fila = por_fila.get(r, {})
            lineas.append("\t".join(fila.get(c, "") for c in range(min_c, max_c + 1)))
        return "\n".join(lineas)

    # ------------------------------------------------------------------
    # Geometría / utilidades
    # ------------------------------------------------------------------
    def _col_count(self):
        return max(1, len(self.headers), *(len(r) for r in self.rows) if self.rows else (1,))

    def _row_count(self):
        return len(self.rows)

    def _col_width_at(self, c):
        """Ancho en píxeles de la columna `c`."""
        if c < len(self._col_widths):
            w = self._col_widths[c]
            if w and w > 0:
                return w
        return self._default_col_width

    def _cell_value(self, row, col):
        if row < len(self.rows) and col < len(self.rows[row]):
            return self.rows[row][col]
        return ""

    def _col_x(self, c):
        """Coordenada x inicial de la columna `c` (sin restar scroll)."""
        x = 0
        for i in range(c):
            x += self._col_width_at(i)
        return x

    def _total_width(self):
        return sum(self._col_width_at(i) for i in range(self._col_count()))

    def _total_height(self):
        return self._header_height + self._row_count() * self._row_height

    def _update_scrollregion(self):
        self.configure(scrollregion=(0, 0, self._total_width(), self._total_height()))

    def _visible_range(self):
        """Retorna (row_start, row_end, col_start, col_end) visibles.

        Con anchos variables, col_start/col_end se calculan acumulando anchos.
        """
        w = max(1, self.winfo_width())
        h = max(1, self.winfo_height())
        ncols = self._col_count()

        row_start = max(0, int(self._scroll_y // self._row_height))
        row_end = min(self._row_count(), int((self._scroll_y + h - self._header_height) // self._row_height) + 1)

        # Columnas visibles por acumulación de anchos.
        acc = 0
        col_start = 0
        col_end = 0
        for c in range(ncols):
            cw = self._col_width_at(c)
            if acc + cw > self._scroll_x and col_start == 0 and c == 0:
                pass
            if acc + cw <= self._scroll_x:
                col_start = c + 1
            if acc < self._scroll_x + w:
                col_end = c + 1
            acc += cw

        return row_start, row_end, col_start, col_end

    def _truncate(self, text, font, max_width):
        """Trunca `text` con elipsis para que no exceda `max_width` píxeles."""
        if not text:
            return ""
        if font.measure(text) <= max_width:
            return text
        # Buscar el prefijo más largo que quepa con "…".
        suffix = "…"
        lo, hi = 0, len(text)
        while lo < hi:
            mid = (lo + hi + 1) // 2
            candidate = text[:mid] + suffix
            if font.measure(candidate) <= max_width:
                lo = mid
            else:
                hi = mid - 1
        if lo == 0:
            return suffix if font.measure(suffix) <= max_width else ""
        return text[:lo] + suffix

    # ------------------------------------------------------------------
    # Render
    # ------------------------------------------------------------------
    def _render(self):
        self.delete("all")

        w = max(1, self.winfo_width())
        h = max(1, self.winfo_height())

        # Fondo del header (fijo).
        self.create_rectangle(0, 0, w, self._header_height, fill=self._header_bg, outline="")

        # Encabezados de columnas visibles.
        _, _, col_start, col_end = self._visible_range()
        for c in range(col_start, col_end):
            cw = self._col_width_at(c)
            x0 = self._col_x(c) - self._scroll_x
            x1 = x0 + cw
            texto = self.headers[c] if c < len(self.headers) else f"Col {c + 1}"
            texto = self._truncate(texto, self._tk_header_font, cw - 2 * self._cell_pad)
            self.create_text(
                x0 + self._cell_pad, self._header_height / 2,
                text=texto, anchor="w",
                font=self._header_font, fill=self._header_fg,
            )
            self.create_line(x1, 0, x1, self._header_height, fill=self._grid_color)

        # Filas visibles.
        row_start, row_end, _, _ = self._visible_range()
        for r in range(row_start, row_end):
            y0 = self._header_height + r * self._row_height - self._scroll_y
            y1 = y0 + self._row_height
            # Fondo alternado.
            base = self._cell_bg if r % 2 == 0 else self._alt_bg
            self.create_rectangle(0, y0, w, y1, fill=base, outline="")
            for c in range(col_start, col_end):
                cw = self._col_width_at(c)
                x0 = self._col_x(c) - self._scroll_x
                x1 = x0 + cw
                valor = self._cell_value(r, c)
                if valor is None:
                    valor = ""
                # Resaltado de celda seleccionada.
                if (r, c) in self._selected:
                    self.create_rectangle(x0, y0, x1, y1,
                                          fill=self._sel_bg, outline=self._sel_border, width=1)
                # Texto truncado al ancho real de la celda.
                texto = self._truncate(str(valor), self._tk_font, cw - 2 * self._cell_pad)
                self.create_text(
                    x0 + self._cell_pad, (y0 + y1) / 2,
                    text=texto, anchor="w",
                    font=self._font, fill=self._text_color,
                )
                # Línea de grilla vertical.
                self.create_line(x1, y0, x1, y1, fill=self._grid_color)
            # Línea de grilla horizontal.
            self.create_line(0, y1, w, y1, fill=self._grid_color)

    # ------------------------------------------------------------------
    # Eventos
    # ------------------------------------------------------------------
    def _bind_events(self):
        self.bind("<Button-1>", self._on_click)
        self.bind("<B1-Motion>", self._on_drag)
        self.bind("<Configure>", lambda e: self._render())
        self.bind("<Up>", lambda e: self._move(0, -1, e.state))
        self.bind("<Down>", lambda e: self._move(0, 1, e.state))
        self.bind("<Left>", lambda e: self._move(-1, 0, e.state))
        self.bind("<Right>", lambda e: self._move(1, 0, e.state))
        self.bind("<Control-c>", self._copy)
        self.bind("<Control-C>", self._copy)

    def _cell_at(self, x, y):
        """Devuelve (row, col) para una posición en píxeles, o None si no cae
        en una celda de datos."""
        if y < self._header_height:
            return None
        row = int((y + self._scroll_y - self._header_height) // self._row_height)
        if row < 0 or row >= self._row_count():
            return None
        # Columna por acumulación de anchos.
        px = x + self._scroll_x
        acc = 0
        for c in range(self._col_count()):
            acc += self._col_width_at(c)
            if px < acc:
                return row, c
        return None

    def _on_click(self, event):
        self.focus_set()
        cell = self._cell_at(event.x, event.y)
        if cell is None:
            return
        r, c = cell
        ctrl = bool(event.state & 0x0004)
        shift = bool(event.state & 0x0001)

        if ctrl:
            if (r, c) in self._selected:
                self._selected.discard((r, c))
            else:
                self._selected.add((r, c))
            self._anchor = (r, c)
        elif shift and self._anchor is not None:
            ar, ac = self._anchor
            r1, r2 = sorted((ar, r))
            c1, c2 = sorted((ac, c))
            for rr in range(r1, r2 + 1):
                for cc in range(c1, c2 + 1):
                    self._selected.add((rr, cc))
        else:
            self._selected.clear()
            self._selected.add((r, c))
            self._anchor = (r, c)

        self._cursor = [r, c]
        self._render()

    def _on_drag(self, event):
        cell = self._cell_at(event.x, event.y)
        if cell is None or self._anchor is None:
            return
        r, c = cell
        ar, ac = self._anchor
        r1, r2 = sorted((ar, r))
        c1, c2 = sorted((ac, c))
        # En arrastre, reemplaza la selección (sin Ctrl) por el rango ancla->actual.
        self._selected.clear()
        for rr in range(r1, r2 + 1):
            for cc in range(c1, c2 + 1):
                self._selected.add((rr, cc))
        self._cursor = [r, c]
        self._render()

    def _move(self, dc, dr, state):
        if self._row_count() == 0:
            return
        r = max(0, min(self._row_count() - 1, self._cursor[0] + dr))
        c = max(0, min(self._col_count() - 1, self._cursor[1] + dc))
        shift = bool(state & 0x0001)

        if shift and self._anchor is None:
            self._anchor = tuple(self._cursor)

        if shift:
            self._selected.add((r, c))
        else:
            self._selected.clear()
            self._selected.add((r, c))
            self._anchor = (r, c)

        self._cursor = [r, c]
        self._ensure_visible(r, c)
        self._render()
        return "break"

    def _ensure_visible(self, r, c):
        """Ajusta el scroll para que la celda (r, c) quede visible."""
        y = self._header_height + r * self._row_height
        x = self._col_x(c)
        cw = self._col_width_at(c)
        h = max(1, self.winfo_height())
        w = max(1, self.winfo_width())

        if y < self._scroll_y + self._header_height:
            self._scroll_y = max(0, y - self._header_height)
        elif y + self._row_height > self._scroll_y + h:
            self._scroll_y = y + self._row_height - h

        if x < self._scroll_x:
            self._scroll_x = x
        elif x + cw > self._scroll_x + w:
            self._scroll_x = x + cw - w

        self._sync_scrollbars()

    def _copy(self, event=None):
        texto = self.get_selected_text()
        if texto:
            self.clipboard_clear()
            self.clipboard_append(texto)
        return "break"

    # ------------------------------------------------------------------
    # Scroll
    # ------------------------------------------------------------------
    def _on_vscroll(self, *args):
        if args[0] == "moveto":
            self._scroll_y = int(float(args[1]) * max(0, self._total_height() - self.winfo_height()))
        elif args[0] == "scroll":
            self._scroll_y += int(args[1]) * (self._row_height * 3 if args[2] == "units" else self.winfo_height())
        self._scroll_y = max(0, min(self._scroll_y, max(0, self._total_height() - self.winfo_height())))
        self._sync_scrollbars()
        self._render()

    def _on_hscroll(self, *args):
        if args[0] == "moveto":
            self._scroll_x = int(float(args[1]) * max(0, self._total_width() - self.winfo_width()))
        elif args[0] == "scroll":
            self._scroll_x += int(args[1]) * (self._col_width * 3 if args[2] == "units" else self.winfo_width())
        self._scroll_x = max(0, min(self._scroll_x, max(0, self._total_width() - self.winfo_width())))
        self._sync_scrollbars()
        self._render()

    def _sync_scrollbars(self):
        h = max(1, self.winfo_height())
        w = max(1, self.winfo_width())
        total_h = max(1, self._total_height())
        total_w = max(1, self._total_width())

        if total_h > h:
            self._vsb.set(self._scroll_y / total_h, (self._scroll_y + h) / total_h)
        else:
            self._vsb.set(0, 1)
        if total_w > w:
            self._hsb.set(self._scroll_x / total_w, (self._scroll_x + w) / total_w)
        else:
            self._hsb.set(0, 1)
