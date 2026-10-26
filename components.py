"""Componentes reutilizables del sistema de diseño CEMEX (CustomTkinter).

Mantiene la API de CTkFrame/CTkButton pero aplica tokens MD3, corner_radius=16,
botones de 44px y variantes primaria/acento/outline sin tocar la lógica de negocio.
"""

import customtkinter as ctk
from customtkinter import CTkFont

from design_tokens import DESIGN_TOKENS

C = DESIGN_TOKENS["colors"]
SP = DESIGN_TOKENS["spacing"]
RD = DESIGN_TOKENS["radius"]
TP = DESIGN_TOKENS["typography"]
SZ = DESIGN_TOKENS["sizes"]


def _font(key):
    return CTkFont(*TP[key])


class CEMEXCard(ctk.CTkFrame):
    """Tarjeta con esquinas redondeadas, borde sutil y padding interno de 24px."""

    def __init__(self, master, title=None, **kwargs):
        kwargs.setdefault("fg_color", C["surface_variant"])
        kwargs.setdefault("corner_radius", RD["lg"])
        kwargs.setdefault("border_width", 1)
        kwargs.setdefault("border_color", C["outline"])
        super().__init__(master, **kwargs)

        self.grid_columnconfigure((0, 1), weight=1)

        if title:
            ctk.CTkLabel(
                self,
                text=title.upper(),
                font=_font("subtitle"),
                text_color=C["primary"],
                anchor="w",
            ).grid(row=0, column=0, padx=SP["lg"], pady=(SP["lg"], SP["sm"]), sticky="ew")
            self._body_row = 1
        else:
            self._body_row = 0

    @property
    def body_row(self):
        """Fila donde los widgets hijos deben empezar a colocarse."""
        return self._body_row


class CEMEXButton(ctk.CTkButton):
    """Botón de 44px con tres variantes: primary, accent y outline."""

    VARIANTS = {
        "primary": {
            "fg_color": C["primary"],
            "hover_color": C["primary_hover"],
            "text_color": C["on_primary"],
            "border_width": 0,
        },
        "accent": {
            "fg_color": C["accent"],
            "hover_color": C["accent_hover"],
            "text_color": C["on_primary"],
            "border_width": 0,
        },
        "outline": {
            "fg_color": C["surface_variant"],
            "hover_color": C["primary_container"],
            "text_color": C["primary"],
            "border_width": 1,
            "border_color": C["outline"],
        },
        "neutral": {
            "fg_color": C["neutral"],
            "hover_color": C["neutral_hover"],
            "text_color": C["on_primary"],
            "border_width": 0,
        },
    }

    def __init__(self, master, text="", variant="primary", **kwargs):
        style = dict(self.VARIANTS[variant])
        style.update(kwargs)
        style.setdefault("height", SZ["button_height"])
        style.setdefault("corner_radius", RD["md"])
        style.setdefault("font", _font("button"))
        super().__init__(master, text=text, **style)


class CEMEXSidebarButton(ctk.CTkButton):
    """Botón de navegación transparente con hover azul al 10% de opacidad."""

    def __init__(self, master, text="", **kwargs):
        kwargs.setdefault("anchor", "w")
        kwargs.setdefault("height", SZ["button_height"])
        kwargs.setdefault("corner_radius", RD["md"])
        kwargs.setdefault("fg_color", "transparent")
        kwargs.setdefault("hover_color", C["primary_container"])
        kwargs.setdefault("text_color", C["on_surface_variant"])
        kwargs.setdefault("font", _font("subtitle"))
        super().__init__(master, text=text, **kwargs)

    def set_active(self, active: bool):
        if active:
            self.configure(
                fg_color=C["primary_container"],
                text_color=C["primary"],
                hover_color=C["primary_container"],
            )
        else:
            self.configure(
                fg_color="transparent",
                text_color=C["on_surface_variant"],
                hover_color=C["primary_container"],
            )
