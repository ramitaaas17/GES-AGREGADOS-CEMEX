"""Sistema de diseño CEMEX Digital — Material Design 3 (light).

Tokens centralizados de color, espaciado, radios y tipografía. Todos los
paddings/márgenes son múltiplos de 8px para mantener una retícula consistente.
"""

DESIGN_TOKENS = {
    "colors": {
        "primary": "#023185",
        "primary_variant": "#0241A8",
        "primary_hover": "#01245F",
        "primary_container": "#EAF0F8",
        "on_primary": "#FFFFFF",
        "accent": "#E3303D",
        "accent_hover": "#B02630",
        "surface": "#F4F6F9",
        "surface_variant": "#FFFFFF",
        "on_surface": "#1A1A1A",
        "on_surface_variant": "#5F6B7A",
        "outline": "#E0E0E0",
        "outline_variant": "#DDE1E6",
        "neutral": "#6C757D",
        "neutral_hover": "#5A6268",
        "elevation_low": "#00000008",
        "elevation_medium": "#0000001A",
    },
    "spacing": {
        "xs": 4,
        "sm": 8,
        "md": 16,
        "lg": 24,
        "xl": 32,
        "xxl": 48,
    },
    "radius": {
        "sm": 6,
        "md": 12,
        "lg": 16,
    },
    "typography": {
        "font_family": "Segoe UI",
        "display": ("Segoe UI", 28, "bold"),
        "headline": ("Segoe UI", 22, "bold"),
        "title": ("Segoe UI", 16, "bold"),
        "subtitle": ("Segoe UI", 14, "bold"),
        "body": ("Segoe UI", 14, "normal"),
        "body_small": ("Segoe UI", 12, "normal"),
        "label": ("Segoe UI", 11, "bold"),
        "button": ("Segoe UI", 14, "bold"),
    },
    "sizes": {
        "button_height": 44,
        "input_height": 38,
        "sidebar_width": 280,
        "row_height": 32,
    },
}


def token(path: str):
    """Acceso por ruta de puntos: token('colors.primary') -> '#023185'."""
    node = DESIGN_TOKENS
    for part in path.split("."):
        node = node[part]
    return node
