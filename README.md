# RutaV3_2026 — Rediseño UI (CEMEX Digital · Material Design 3)

Interfaz de escritorio para análisis financiero, automatización ETL y
consolidación CEMEX. Este rediseño aplica un sistema de diseño basado en
tokens (MD3 light) sobre **CustomTkinter**, sin alterar la lógica de negocio.

## Instalación

Las rutas de datos, base de datos y Excel maestro se resuelven dinámicamente
con `os.path.dirname(os.path.abspath(__file__))`, por lo que la aplicación es
portable y funciona desde cualquier carpeta (por ejemplo `C:\RutaV3_2026`).

```powershell
cd C:\RutaV3_2026
python -m venv .venv
.\.venv\Scripts\Activate.ps1
python -m pip install --upgrade pip
python -m pip install -r requirements.txt
```

## Sincronización manual (sin servidor)

Cada usuario trabaja en su copia local (`staging_rutas.db` + Excel maestro) y
usa una carpeta "drive" como punto de intercambio:

- **Subir al drive**: copia la BD y el Excel locales a la carpeta drive.
- **Cargar datos**: sobrescribe la copia local con la versión del drive
  (crea un backup con fecha antes de sobrescribir).

La ruta del drive se configura en `sync_config.json`:

```json
{
  "drive_dir": "C:\\Users\\TU_USUARIO\\OneDrive - CEMEX\\RutaV3_2026_Drive"
}
```

Cada equipo tiene su propia ruta OneDrive, por lo que ese archivo debe
configurarse por usuario.

## Ejecución

```powershell
python main.py
```

## Estructura del frontend

| Archivo            | Responsabilidad                                             |
| ------------------ | ---------------------------------------------------------- |
| `design_tokens.py` | Colores, espaciado, radios, tipografía y tamaños centralizados |
| `components.py`    | `CEMEXCard`, `CEMEXButton`, `CEMEXSidebarButton`           |
| `main.py`          | Layout con sidebar, header, grid de cards y tablas          |
| `etl_processor.py` | Motor ETL: cálculo, lookups SAP, inyección en Excel y reportes |
| `database_manager.py` | Conexión SQLite, creación/migración de esquema e inserciones |
| `fpa_validations.py` | Backend del Área de Validaciones (FPA): pegado 24 cols, búsqueda, proyección M3, evidencias y portapapeles Outlook |

## Sistema de diseño

- **Colores**: azul `#023185`, acento rojo `#E3303D`, superficie `#F4F6F9`.
- **Espaciado**: todos los paddings/márgenes en múltiplos de 8px (4/8/16/24/32).
- **Cards**: `corner_radius=16`, borde sutil `#E0E0E0`.
- **Botones**: altura mínima de 44px, variantes `primary`, `accent`, `outline`, `neutral`.
- **Tipografía**: Segoe UI en toda la aplicación.

## Atajos de teclado

| Tecla       | Acción                                       |
| ----------- | -------------------------------------------- |
| `F4`        | Alterna la casilla "Mantener MP"             |
| `F5`        | Alterna la casilla "Mantener Flete"          |
| `F6`        | Busca solicitudes FPA (SQLite + Excel)       |
| `F7`        | Proyección a M3 (TN × PV)                    |
| `F8`        | Abre modal de Validar (solo lectura)         |
| `F9`        | Abre modal de Evidencias FPA                 |
| `Ctrl+V`    | Pegado masivo desde Excel (autocompleta)     |
| `Enter`     | Calcula la fila actual                       |

## Área de Validaciones (FPA)

- Pega filas de Excel con las 24 columnas FPA exactas usando `Ctrl+V` sobre el Treeview.
- **Buscar (F6)** localiza solicitudes en SQLite y en el Excel Maestro por clave compuesta (SF, CEDIS, NO PLANTA, ARTICULO, MODALIDAD), y opcionalmente refina por **Fecha de Procesamiento** (columna `AF` en Excel, `fecha_procesamiento` en SQLite).
- **Migrar Excel → SQLite** lee la tabla `Solicitudes` completa del Excel Maestro y carga en SQLite solo las filas que aún no existen (evita duplicados).
- **Proyección M3 (F7)** muestra los precios convertidos a M3 cuando `UM == TN`, multiplicando por PV. Es solo lectura: no modifica la data cruda.
- **Validar (F8)** abre un modal con el resumen de la fila seleccionada.
- **Evidencias FPA (F9)** captura imágenes pegadas desde SAP (`Ctrl+V`), actualiza "Fecha de Confirmación" (columna `fecha_conf` en SQLite y `AO FECHA CONF` en Excel), y arma el portapapeles mixto (texto + imágenes) listo para Outlook. Incluye campo editable de fecha de confirmación (por defecto la fecha actual). Al finalizar, borra las filas de la solicitud activa para ingresar una nueva.

## Notas de compatibilidad

- Se fuerza `appearance_mode = "light"`.
- No se modifican `xlwings`, `threading`, ni los handlers de `etl_processor`.
- Las tablas usan `ttk.Treeview` con estilo CEMEX (zebra striping y cabecera azul).
- Excel debe tener habilitado "Confiar en el acceso al modelo de objetos de proyectos de VBA" (Centro de confianza) para que xlwings escriba sin errores.
- La base de datos `staging_rutas.db` se crea automáticamente si no existe.
