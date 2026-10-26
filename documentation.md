# Herramienta de Captura y Consolidación de Datos para Excel (RutaV3_2026)

## Introducción

Esta herramienta optimiza el proceso de captura y validación inicial de datos para el modelo de consolidación financiera `RutaV3_2026`. El método original dependía de múltiples `XLOOKUPs` en Excel sobre tablas SAP pesadas (VK13 con >21,000 filas y TRAOPE), lo que lo hacía lento y propenso a errores. Esta solución separa el ingreso de datos del cálculo pesado mediante:

*   Una **interfaz gráfica ligera** (`CustomTkinter`) para captura manual o pegado rápido desde Excel.
*   Una **base de datos local SQLite** (`staging_rutas.db`) como zona de staging.
*   Un **motor ETL en pandas** que replica las fórmulas del Reporte Maestro de Fórmulas (PDF de especificación técnica) y exporta CSVs listos para Power Query.

---

## Estructura del Proyecto

| Archivo | Descripción |
|---------|-------------|
| `main.py` | GUI de captura (CustomTkinter). Maneja la entrada manual, el pegado desde Excel (Ctrl+V), la previsualización de la fila calculada (Columnas A-AR) y la confirmación para guardar en SQLite. |
| `etl_processor.py` | Motor ETL con `pandas`: calcula las columnas formuladas, ejecuta los cruces contra los catálogos SAP simulados y exporta los CSVs. |
| `database_manager.py` | Operaciones SQLite: conexión, creación de tabla, migración de esquema e inserción de registros. |
| `staging_rutas.db` | Base de datos local que almacena las filas confirmadas (campos manuales). |
| `PVTA_MAT VK13.csv` | Catálogo SAP simulado de precios de venta de materia prima (VK13). |
| `PVTA_FTE VK13.csv` | Catálogo SAP simulado de precios de venta de flete (VK13). |
| `CONT_COMPRA TRAOPE.csv` | Catálogo SAP simulado de contratos de compra (TRAOPE). |
| `input_powerquery.csv` | **Salida principal**: filas completas calculadas y validadas, listas para Power Query. |
| `cto_venta.csv` | Reporte de catálogo de costos de venta (simula la macro CTO VTA). |
| `cto_compra.csv` | Reporte de catálogo de costos de compra (simula la macro CTO COMPRA). |
| `fpa_filial.csv` | Reporte FPA para clientes **FILIAL** (simula la macro FPA FILIAL). |
| `fpa_terceros.csv` | Reporte FPA para clientes **TERCEROS** (simula la macro FPA TERCEROS). |

---

## Flujo de Trabajo

```
Excel / Captura manual → GUI (Ctrl+V o teclado) → Calcular Fila (Enter)
→ Previsualización A-AR → Confirmar y Guardar a SQLite
→ Procesar ETL → input_powerquery.csv + cto_venta.csv + cto_compra.csv
                + fpa_filial.csv + fpa_terceros.csv → Power Query en Excel
```

---

## 1. Interfaz de Captura Rápida (`main.py`)

### 1.1 Zona de Entrada Manual

La GUI solicita **únicamente los campos sin fórmula** (Columnas B a AO):

*   **Identificación**: ORG VTA, CEDIS, NOMBRE CEDIS, SF, NOMBRE SF, NO PLANTA/OBRA, NOMBRE PLANTA/OBRA, ARTICULO, DESC ARTÍCULO, PV, MODALIDAD.
*   **Costos de venta (M3)**: MP VTA (M3), FLETE VTA (M3).
*   **Costos de compra (M3)**: MP COMP (M3), CA COMP (M3), FLETE COMP (M3), INTEGRAL COMP (M3).
*   **Gestión**: TIPO, FECHA DE PROCESAMIENTO, TIPO CLIENTE, OBSERVACION, FECHA SOLICITUD, INCIDENCIA?, MOTIVO, MES, FECHA CONF.
*   **FORMATO DESTINO** (menú desplegable): `CTO COMPRA`, `CTO VENTA`, `FPA`. **Obligatorio** para calcular y guardar la fila. Este campo **no se auto-llena con Ctrl+V**: el usuario lo elige manualmente porque determina a qué reporte irá la fila.

Los campos con menú desplegable (`MODALIDAD`, `UM VTA`, `TIPO`, `TIPO CLIENTE`, `INCIDENCIA`, `FORMATO DESTINO`) validan los valores permitidos.

### 1.2 Auto-llenado por Portapapeles (Excel → GUI) — **NUEVO**

Al presionar **`Ctrl+V`** en cualquier parte de la ventana:

1.  La app lee el portapapeles.
2.  Si el texto contiene **tabulaciones** (`\t`), asume que es una fila copiada de Excel.
3.  Toma la primera línea y la separa por tabulaciones.
4.  **Auto-llena los campos B a L** (ORG VTA, CEDIS, NOMBRE CEDIS, SF, NOMBRE SF, NO PLANTA/OBRA, NOMBRE PLANTA/OBRA, ARTICULO, DESC ARTÍCULO, PV, MODALIDAD) mapeando cada columna de Excel a su campo correspondiente.
5.  La columna A (RUTA) se ignora porque se recalcula automáticamente.

Esto permite copiar una fila desde el Excel maestro y pegarla directamente, reduciendo errores de captura. La escritura manual sigue disponible en todo momento.

### 1.3 Cálculo y Previsualización

*   **Botón "Calcular Fila"** (o tecla `Enter`): procesa las fórmulas internamente y muestra en pantalla las **Columnas A a AR** completas en formato vertical (tabla de 2 columnas: nombre de columna = valor).
*   **Botón "Confirmar y Guardar a SQLite"**: se habilita **solo después** de calcular la fila. Al presionarlo, inserta los campos manuales en `staging_rutas.db` y limpia el formulario.
*   **Botón "Procesar ETL"**: lanza el pipeline completo que lee SQLite, calcula todas las columnas y exporta los CSVs.

---

## 2. Motor de Cálculo y Validación (`etl_processor.py`)

### 2.1 Construcción de la Llave Compuesta (Columna A)

```
RUTA = CEDIS + "-" + SF + "-" + NO PLANTA/OBRA + "-" + ARTICULO + "-" + último dígito de MODALIDAD
```

### 2.2 Reglas Matemáticas Corregidas con PV (Columnas N a AD) — **CORREGIDO**

El **PV (Peso Volumétrico, Columna K)** es el factor de conversión entre Metros Cúbicos (M3) y Toneladas (TN). Según las fórmulas reales del Excel (fila 2946), la conversión se hace **dividiendo entre PV**:

| Columna | Fórmula en Excel | Implementación en Python |
|---------|------------------|--------------------------|
| N `$ MP VTA (TN)` | `=Q2946/K2946` | `MP VTA (M3) / PV` |
| O `$ FLETE VTA (TN)` | `=[@FLETE VTA (M3)]/[@PV]` | `FLETE VTA (M3) / PV` |
| P `$ INTEGRA VTA (TN)` | `=+$O+N` | `FLETE VTA (TN) + MP VTA (TN)` |
| S `INTEGRAL VTA (M3)` | `=SUMA(Q:R)` | `MP VTA (M3) + FLETE VTA (M3)` |
| X `INTEGRAL COMP (M3)` | `=+U+V+W` | `MP COMP (M3) + CA COMP (M3) + FLETE COMP (M3)` |
| Y `MP COMP (TN)` | `=U2946/K2946` | `MP COMP (M3) / PV` |
| Z `CA COMPR (TN)` | `=+V2946/K2946` | `CA COMP (M3) / PV` |
| AA `FLETE COMP (TN)` | `=+W2946/K2946` | `FLETE COMP (M3) / PV` |
| AB `INTEGRAL COMP (TN)` | `=+AA+Z+Y` | `FLETE COMP (TN) + CA COMP (TN) + MP COMP (TN)` |

**Nota importante**: la conversión M3 → TN es **división** (`M3 / PV`), no multiplicación. Esto replica exactamente el comportamiento del Excel.

### 2.3 Márgenes Operativos (Columnas AC, AD, AI, AJ)

Según la `UM COMP [Solicitud]` (M3 o TN):

*   **AC `$ MOP INTEGRAL`**: `INTEGRAL VTA - INTEGRAL COMP` (en M3 o TN según la UM).
*   **AD `% MOP INTEGRAL`**: `(INTEGRAL VTA - INTEGRAL COMP) / INTEGRAL VTA`.
*   **AI `$ MOP/M3`**: `MP VTA - MP COMP` (en M3 o TN según la UM).
*   **AJ `MOP %`**: `(MP VTA - MP COMP) / MP VTA`.

### 2.4 Tiempos de Solicitud (Columnas AP, AQ, AR)

*   **AP `DÍAS HÁBILES`**: emula `NETWORKDAYS.INTL(fecha_procesamiento, fecha_conf, 1) - 1` (cuenta días laborables lun-vie entre ambas fechas).
*   **AQ `DÍAS TRANSCURRIDOS`**: `fecha_conf - fecha_procesamiento` en días calendario.
*   **AR `ESTADO DE SOLICITUD`**:
    *   Si existe `FECHA CONF` → `CONFIRMADO`.
    *   Si `DÍAS TRANSCURRIDOS` entre 0 y 3 → `VIGENTE`.
    *   Si es 4 → `POR VENCER`.
    *   Si es ≥ 5 → `VENCIDO`.

### 2.5 Cruces con SAP/Snowflake (Columnas AS a BE)

*   **`PVTA_MAT VK13`**: lookup por RUTA para obtener `MP VTA SNOV` y `UM MAT`.
*   **`PVTA_FTE VK13`**: lookup por RUTA para obtener `FLETE VTA SNOV` y `UM FTE` (por defecto `RECOGIDO` si no existe).
*   **`CONT_COMPRA TRAOPE`**: lookup condicional — si la MODALIDAD termina en `4`, la llave es la RUTA completa; en caso contrario, `CEDIS-SF-ARTICULO`.
*   Validaciones: `var MAT`, `var UMM`, `var FTE`, `vs UMF`, `Validar Costo Compra vs Solicitud` (tolerancia `±0.1` → `ENTREGADO`/`RECOGIDO`, en caso contrario `VERIFICAR`), `Variación Compra`, `Valida UM`.

Los catálogos se cargan con **tolerancia a fallos**: si un CSV no se encuentra, el proceso continúa con una advertencia y los cruces devuelven valores neutros (0 / `RECOGIDO`).

---

## 3. Migración de Macros VBA a Pandas — Enrutamiento por FORMATO DESTINO

La función `generar_reportes_fpa_cto(df_validado)` reemplaza las macros de Excel (VBA) y **enruta cada fila según el campo `FORMATO DESTINO`** capturado en la GUI:

| FORMATO DESTINO | Regla de enrutamiento | Archivo de salida |
|-----------------|----------------------|-------------------|
| `CTO VENTA` | Exporta las filas directamente | `cto_venta.csv` |
| `CTO COMPRA` | Aplica mapeo base→target + valores fijos | `cto_compra.csv` |
| `FPA` + `ORG VTA = 7100` | Mapeo FPA + valores fijos FILIAL | `fpa_filial.csv` |
| `FPA` + `ORG VTA = 7180` o `7277` | Mapeo FPA + valores fijos TERCEROS | `fpa_terceros.csv` |

### 3.1 CTO COMPRA (mapeo de macros)

Mapeo de columnas (base → target):

| Col target | Fuente en base |
|------------|----------------|
| C | SF |
| E | ORG VTA |
| H | FECHA DE PROCESAMIENTO |
| J | CEDIS |
| L | ARTICULO |
| W | MODALIDAD |
| N, Q, AA | UM COMP [Solicitud] |

Valores fijos: `D="MXPO"`, `F="MXL"`, `G="MXN"`, `I="31.12.2030"`, `K="'001"`, `M="999999999"`, `P="1"`, `R="V4"`, `T="K"`, `Z="1"`.

Lógica adicional: si `MODALIDAD` (Col L base) es `"04"` o `"4"`, `Col_X = NO PLANTA / OBRA`; en caso contrario, `Col_X` queda vacío.

### 3.2 FPA (mapeo de macros)

Mapeo común FPA (base → target):

| Col target | Fuente en base |
|------------|----------------|
| A | TIPO |
| F | ORG VTA |
| G | SF |
| H | NOMBRE SF |
| I | CEDIS |
| J | NOMBRE CEDIS |
| L | NO PLANTA / OBRA |
| M | NOMBRE PLANTA / OBRA |
| N | ARTICULO |
| O | DESC ARTÍCULO |
| R | UM VTA |
| U | MODALIDAD |
| W | FECHA DE PROCESAMIENTO |
| AL | $ MOP INTEGRAL |
| AM | % MOP INTEGRAL |
| AN | PV |

Valores fijos:
*   **FPA FILIAL**: `P="ZMAH"`, `S="ZMPH"`, `X="31.12.9999"`.
*   **FPA TERCEROS**: `P="ZMA6"`, `S="ZMP1"`, `X="31.12.9999"`.

Lógica UM FPA:
*   `Col_Q`: si `Col_R` (UM VTA) es `TN` → `$ MP VTA (TN)`; si es `M3` → `MP VTA (M3)`.
*   `Col_T`: si `Col_R` es `TN` → `$ FLETE VTA (TN)`; si es `M3` → `FLETE VTA (M3)`.

Los nombres de columna se normalizan (se eliminan los prefijos de letra tipo `A `, `B `) para que los CSV sean legibles en Excel.

---

## 4. Exportación y Manejo de Errores — **NUEVO**

La exportación a `input_powerquery.csv` usa:

```python
result.to_csv(OUTPUT_CSV, index=False, mode="w", encoding="utf-8")
```

*   `index=False`: no se escribe el índice de pandas.
*   `mode="w"`: sobrescribe el archivo limpiamente desde la raíz del proyecto en cada ejecución.
*   `encoding="utf-8"`: compatible con caracteres acentuados (Á, Í, etc.).

### Diagnóstico de errores de escritura

Si `to_csv` falla, la aplicación:

1.  **Imprime en la terminal** el error completo con `traceback` (tipo de excepción, mensaje y archivo destino) para facilitar el diagnóstico.
2.  **Muestra una alerta visible en la GUI** (CustomTkinter): `"Fallo al guardar el CSV: {str(e)}"`.

---

## Instrucciones para Ejecutar el Script

1.  **Asegúrese de tener Python instalado** (3.8 o superior).
2.  **Instale las bibliotecas necesarias**:
    ```bash
    pip install customtkinter pandas
    ```
3.  **Coloque los archivos** en la misma carpeta (los 3 módulos Python y los CSV simulados de SAP).
4.  **Navegue a la carpeta** y ejecute:
    ```bash
    cd C:\Users\coazo\Downloads\proyeceto
    python main.py
    ```

### Uso rápido

1.  Copie una fila del Excel maestro (o capture a mano los campos).
2.  Presione **`Ctrl+V`** para auto-llenar los campos B-L desde el portapapeles.
3.  Complete los costos (M3) y datos de gestión si no vinieron del pegado.
4.  Presione **`Calcular Fila`** (Enter) y revise la previsualización A-AR.
5.  Si es correcta, presione **`Confirmar y Guardar a SQLite`**.
6.  Al finalizar la captura, presione **`Procesar ETL`** para generar `input_powerquery.csv` y los 4 reportes (`cto_venta.csv`, `cto_compra.csv`, `fpa_filial.csv`, `fpa_terceros.csv`).
7.  Ingeste `input_powerquery.csv` en Power Query dentro del Excel maestro.

---

## Notas Técnicas

*   **Conversión TN ↔ M3**: siempre `M3 / PV`. Si `PV` es 0 o está vacío, el valor TN se deja en 0 para evitar divisiones inválidas.
*   **UM COMP [Solicitud]** se toma igual a `UM VTA` según el modelo.
*   **Tolerancia de validación**: una diferencia de costos `≤ 0.1` entre Snowflake y la solicitud devuelve `ENTREGADO` o `RECOGIDO`; en caso contrario, `VERIFICAR`.
*   **Migración de esquema**: `database_manager.migrate_if_needed()` migra automáticamente la tabla `staging_rutas` a la versión más reciente (18 columnas, incluyendo `formato_destino`). Soporta los esquemas v0 (8 cols), v1 (17 cols) y v2 (18 cols).
*   **FORMATO DESTINO**: campo obligatorio que determina a qué reporte se enruta cada fila (`CTO COMPRA`, `CTO VENTA` o `FPA`). Se almacena en SQLite y se propaga al DataFrame validado.

---

## Changelog

### v1.2.0 — Concurrencia, recuperación y mejoras operativas

*   **Concurrencia (semáforo de colisiones)**: `sync_manager.py` guarda la huella (`mtime`) de los archivos del drive al descargar y, antes de subir, bloquea la subida si detecta que otro usuario subió una versión más nueva. En ese caso sugiere usar la Sincronización de Recuperación.
*   **Auditoría SQLite**: nueva columna `timestamp_modificacion` (migrada automáticamente) que registra `datetime.now()` en cada INSERT/UPDATE de `staging_rutas`.
*   **Sincronización de Recuperación**: botón que, en segundo plano, descarga el Excel actualizado del drive, filtra los registros trabajados hoy y con una instancia de Excel invisible (`visible=False`) hace UPDATE de filas existentes e INSERT de filas nuevas al final de la tabla, sin perder el trabajo local del día.
*   **Terminal SQL interna**: ventana independiente para ejecutar SQL crudo contra `staging_rutas.db` (SELECT en Treeview dinámico; UPDATE/DELETE con mensaje de éxito/error).
*   **Pegado FPA robusto**: `parse_clipboard_fpa()` acepta rangos con encabezados (pandas + `reindex` sobre la lista maestra de 24 columnas, rellenando con vacío las omitidas) y conserva el formato posicional de 26 columnas con celdas combinadas/ocultas.
*   **Copiado estilo Excel en Treeviews**: `Ctrl+C` copia las filas seleccionadas (tabulaciones + saltos de línea) listas para pegar en Excel.
*   **Menú lateral colapsable**: botón hamburguesa que oculta/expande la barra lateral.
*   **Finalizar FPA línea por línea**: procesa las filas seleccionadas una a una, alerta antes de sobrescribir `fecha_conf`, elimina visualmente cada fila procesada y muestra el resumen global solo al final.
*   **Ventana dinámica de finalización**: popup "mini" (solo fecha + evidencias) al validar múltiples filas; popup detallado al validar una sola.
*   **Búsqueda Avanzada**: popup independiente con más criterios de filtrado sin saturar la pestaña de Validaciones y SLA.

### v1.2.1 — Selección por celdas, recuperación robusta y Drag & Drop

*   **Selección por celdas (estilo Excel/SAP)**: nuevo componente `cell_grid.py` que permite seleccionar celdas individuales (clic simple, `Ctrl+clic`, `Shift+clic`/arrastre, flechas y `Shift+flechas`) y copiarlas con `Ctrl+C`. Se abre desde los botones "Selección por celdas" de SLA y FPA.
*   **Recuperación de sincronización robusta**: la manipulación de Excel se aisló en `recovery_worker.py` (subproceso con timeout), usando una instancia visible de Excel en lugar de invisible para evitar procesos `EXCEL.EXE` zombi y cuelgues.
*   **Drag & Drop de Excel**: la ventana acepta soltar archivos `.xlsx`/`.xlsm` (`tkinterdnd2`). Busca hojas FPA/CTO (`fpa filial`, `fpa terceros`, `cto venta`, `cto compra`), extrae la tabla con pandas (encabezado detectado, `dropna(how='all')`, `fillna("")`) y la muestra en el Treeview SLA.
*   **Fechas robustas**: `normalizar_fecha()` en `main.py` y `_norm_fecha()` en `fpa_validations.py` aceptan separadores `.`, `-` y `/`, y normalizan a `dd/mm/yyyy` antes de consultar SQLite o Excel.
