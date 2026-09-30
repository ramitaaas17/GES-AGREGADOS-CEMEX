"""Selecciones personales de centros; no se guardan en la carpeta compartida."""
import json
import os
from pathlib import Path
import shutil
import tempfile


def ruta_filtros():
    return Path(os.environ.get('LOCALAPPDATA', str(Path.home()))) / 'CEMEX' / 'GeneradorMatrices' / 'filtros.json'


def cargar(ruta):
    if not ruta.exists():
        return {'filtros': {}, 'predeterminado': ''}
    datos = json.loads(ruta.read_text(encoding='utf-8'))
    if (not isinstance(datos, dict) or not isinstance(datos.get('filtros'), dict)
            or not isinstance(datos.get('predeterminado'), str)
            or any(not isinstance(k, str) or not isinstance(v, list)
                   or any(not isinstance(c, str) for c in v) for k, v in datos['filtros'].items())):
        raise ValueError('Formato de filtros inválido')
    return datos


def guardar(ruta, datos):
    ruta.parent.mkdir(parents=True, exist_ok=True)
    if ruta.exists():
        shutil.copy2(ruta, ruta.with_suffix('.json.bak'))
    temporal = None
    try:
        with tempfile.NamedTemporaryFile(mode='w', encoding='utf-8', dir=ruta.parent, delete=False) as archivo:
            temporal = Path(archivo.name)
            json.dump(datos, archivo, ensure_ascii=False, indent=2)
        os.replace(temporal, ruta)
    finally:
        if temporal is not None and temporal.exists():
            temporal.unlink()
