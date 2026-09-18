# Contexto del Proyecto — Matriz Ventas (Agregados CEMEX)

> Este documento existe para que cualquier sesión futura (humana o de IA) pueda retomar
> este proyecto sin tener que re-explicar todo desde cero. Cubre el objetivo de negocio,
> el estado anterior, todo lo que se ha hecho hasta ahora, los bugs encontrados (corregidos
> y pendientes), las decisiones de arquitectura discutidas, y lo que falta. Se actualizó por
> última vez tras encontrar un bug pendiente de corregir en las columnas de Validación
> (ver sección "PROBLEMA ABIERTO — sin resolver" al final, es lo más urgente).

---

## 1. Qué es esto y por qué existe

CEMEX (línea de negocio Agregados) necesita generar, para cada CEDIS (centro de
distribución), un archivo llamado **"Base Cedis DW##"** que cruza:

- **Precios de venta de materia prima (MP)** — cuánto se cobra por material/destino.
- **Fletes** — cuánto cuesta transportar ese material a ese destino.
- **Costos de compra (TRAOPE)** — cuánto costó comprarlo (condiciones de compra/orden).
- **Catálogo de materiales** — nombre y factor de conversión Peso/Volumen (PV) de cada
  material.

Con esos 4 insumos cruzados, el archivo resultante calcula automáticamente una
**"Validación de margen"**: compara lo que debería costar (fórmula de negocio) contra lo
que realmente costó, y resalta en amarillo las filas donde la diferencia es sospechosa
(mayor a 1). Este archivo es el que usa el equipo comercial/finanzas para decisiones de
precio y para detectar errores de carga en SAP.

## 2. Cómo funcionaba antes (proceso viejo)

Los 4 insumos se extraían **directamente de transacciones de SAP**, en 4 archivos Excel
crudos exportados manualmente, **uno por CEDIS**:

| Archivo | Se identifica por (nombre) | Contenido |
|---|---|---|
| MP / VK13 | `mp`, `precios de venta`, `VK13` (sin "flete") | Precio de venta de materia prima |
| Fletes | `flete(s)`, `VK13 flete` | Importe de flete |
| TRAOPE | `traope` | Condiciones de compra / precio neto pedido |
| Materiales | `material(es)` | Catálogo con PV (Peso/Volumen) |

Ya existía un script en Python (`matrizVentas.py`) que:
1. Detecta los 4 archivos por nombre en una carpeta (no importa el orden).
2. Los parsea (soporta un formato "columnar" viejo y uno "vertical apilado" más nuevo
   que a veces exportaba SAP).
3. Cruza todo usando llaves `Concat1` (ShipFrom+Centro+Destino+Material) y `Concat2`
   (ShipFrom+Centro+Material, sin destino, como respaldo cuando no hay match exacto).
4. Genera 2 archivos de salida por CEDIS:
   - **`Base Cedis DW## (Con Formulas).xlsx`** — con fórmulas VLOOKUP/BUSCARV reales
     (para poder auditar/verificar a mano en Excel).
   - **`Base Cedis DW##.xlsx`** — mismos datos pero con valores ya calculados en Python
     (más rápido de abrir), con el resaltado amarillo de validación.
5. Organiza todo en `MatrizVentas_Generado\<CEDIS>\entradas\` y `\salidas\`.

Este proceso funcionaba, pero era manual: alguien tenía que exportar los 4 archivos de
SAP a mano, por cada CEDIS, cada vez.

## 3. Qué cambió — la nueva fuente de datos (Snowflake)

Ahora la empresa tiene acceso a **tablas de Snowflake**, y existe un archivo Excel que ya
jala de esas tablas los datos más importantes y los resume — **con la misma estructura
de campos que los archivos de SAP de siempre**, pero:

- Es **un solo archivo `.xlsm`** con varias hojas (en vez de 4 archivos sueltos).
- Trae **TODOS los CEDIS mezclados** en las mismas hojas (antes cada archivo era de un
  solo CEDIS). Se identificó que hay ~100-102 CEDIS distintos en el archivo.
- Distingue dos líneas de negocio por "clase de condición" SAP, mezcladas en las mismas
  hojas:
  - **Sacos/Intergiros** → usa códigos `ZMAH` (precio MP) / `ZMPH` (flete).
  - **Agregados/Terceros** → usa códigos `ZMA6` (precio MP) / `ZMP1` (flete).
  - Los Centros de Sacos/Intergiros suelen tener formato `DW##` (ej. `DW68`) y los de
    Agregados/Terceros formato `D###` (ej. `D836`) — aunque esto es una observación, no
    una regla garantizada al 100%.
- El archivo Excel de origen (una vez que viva en SharePoint) se **refresca
  automáticamente** desde Snowflake — eso tiene que reflejarse en las bases generadas.

**Archivo de prueba usado durante todo este trabajo** (ruta local, en la computadora de
pruebas del usuario):
```
C:\Users\twaaa\Downloads\Extracción de Datos Agregados _ Formatos Para Alta de Ruta (ultimo)_2026.xlsm
```

### 3.1 Las 4 hojas relevantes y su mapeo exacto de columnas

El usuario dio este mapeo (confirmado 1:1 contra el archivo real vía Excel COM):

**`CONT_COMPRA TRAOPE`** (equivale a TRAOPE):
| Campo | Columna |
|---|---|
| Ship From | F |
| Nombre SF | G |
| Material | L |
| Centro | N |
| Destino | Z |
| Nombre Destino | AA |
| Precio neto pedido | T |
| UM Precio Pedido | V |

(Esta hoja trae 38 columnas en total, encabezados reales en la fila 1, sin filas basura
antes.)

**`PVTA_MAT VK13`** (equivale a MP) — 12 columnas, encabezados reales:
`Clase Cond.`, `Org. Ventas`, `Shipfrom`, `Centro`, `Destinatario`, `Material`,
`Inicio Validez`, `Valido a`, `Modif_Date`, `Importe`, `Unidad`, `Ruta`.
Se mapean a los mismos nombres normalizados internos que ya usaba el pipeline
(`sf`, `cedis`, `destino`, `material`, `importe_mp`, `um`, `inicio_vig`, `fin_vig`, etc.)
No trae Nombre Destino ni Nombre Material (esos se resuelven después vía TRAOPE y
Materiales respectivamente).

**`PVTA_FTE VK13`** (equivale a Flete) — 13 columnas, encabezados reales:
`Clase Cond.`, `Org. Ventas`, `Shipfrom`, `Centro`, `Destinatario`,
`Cond. Expedición`, `Material`, `Inicio Validez`, `Fin Validez`, `Modif_Date`, `Importe`,
`Unidad`, `Ruta`. A diferencia del archivo viejo, esta hoja ya trae `Cond. Expedición`
directo por fila (no hay que inferirla).

**`PESO_VOL`** (equivale a Materiales) — 16 columnas. Columna clave: **columna E
("Cant. UMB")** ya trae el **PV (Peso/Volumen) calculado y listo para usar tal cual**
— instrucción explícita del usuario: **no recalcular** como hacía el script viejo
(que calculaba `PV = Contador / Denom` y filtraba `Denom = 1000`). También trae
`Texto de material` (nombre) y `Material` (clave).

## 4. El objetivo final — SharePoint (todavía sin definir del todo)

La idea de fondo (el usuario aún no tiene la decisión 100% cerrada) es que esta
automatización termine viviendo **dentro de un SharePoint**. Se plantearon dos
propuestas a evaluar (o una mezcla, o algo mejor):

**Propuesta A — una sección/carpeta por CEDIS:**
- Debe poder catalogarse/clasificarse (Sacos/Intergiros vs Agregados/Terceros, ver
  arriba).
- Todas las bases ahí, y se actualizan cuando alguien pica un botón de "Refresh".
- Pros: simple, calza con la estructura de carpetas que ya usa el script
  (`MatrizVentas_Generado\<CEDIS>\salidas\`), fácil dar permisos por región.
- Contras: con 100+ CEDIS son 200+ archivos, duplicando datos globales (ej. el catálogo
  de Materiales se repetiría en cada uno); cualquier refresh implica regenerar el lote
  completo.

**Propuesta B — barra de búsqueda en SharePoint:**
- La persona busca lo que necesita y el sistema genera esa vista/archivo bajo demanda.
- Pros: no hay que pre-generar ni guardar 100+ archivos; más "tipo app".
- Contras: Python no corre nativo dentro de SharePoint/Power Platform — habría que
  portar la lógica a Power Automate/Office Scripts, o exponerla vía una Azure Function;
  y si cada búsqueda tuviera que releer el archivo fuente completo desde cero, no sería
  instantáneo.

**Recomendación dada (pendiente de que el usuario decida):** una **mezcla híbrida** —
un refresh centralizado (pesado, corre una sola vez cuando el origen cambia) que deja
los datos ya limpios y calculados en un solo lugar (no 100+ archivos sueltos, sino algo
tipo tabla/lista centralizada), y encima una búsqueda/vista filtrable que solo lee esos
datos ya calculados — así la respuesta al usuario final es instantánea, y el trabajo
pesado se aísla y controla aparte.

**Esto NO se ha empezado a construir.** Es solo la discusión/recomendación de
arquitectura. Ver sección 7 para el estado real de "qué está construido hoy".

## 5. Decisión de alcance/fases (importante — para no perder el hilo)

El usuario fue explícito en que:
1. **Por ahora NO hace falta SharePoint** — eso es el objetivo final, no el paso actual.
2. **Lo primero es adaptar el código existente** para que jale los datos directo del
   archivo consolidado (Snowflake) y genere el archivo de salida correcto seguido la
   plantilla ya usada, con los campos correctos.
3. **Mientras tanto, todo puede quedarse en su computadora** — él mismo corre el script
   a mano por ahora — pero el código **sí tiene que quedar preparado** para eventualmente
   engancharse a la solución de SharePoint.
4. **Se resolvió la instalación de Python** en esta máquina de pruebas (no había).
5. Más adelante el usuario aclaró algo importante de logística: **esta computadora es
   solo de pruebas**. El plan real es llevar todo esto a **otra computadora** — esa otra
   sí tiene el archivo Excel ya configurado con la sincronización en vivo a Snowflake
   (queries/conexiones reales) — y ahí es donde eventualmente se configurará el
   disparador automático (Propuesta A de arriba, con el botón de refresh). Pero antes de
   pasarlo hay que dejar **todo pulido y probado aquí**, sin sorpresas.
6. El usuario explícitamente pidió **no configurar todavía** el disparador automático
   (Task Scheduler / file watcher) — eso queda pendiente para cuando se mude a la otra
   computadora.

## 6. Ubicación de archivos clave

| Qué | Ruta |
|---|---|
| Script principal | `C:\Users\twaaa\Documents\GES-AGREGADOS-CEMEX\MatrizVentas\matrizVentas.py` |
| README técnico (uso del script) | `C:\Users\twaaa\Documents\GES-AGREGADOS-CEMEX\MatrizVentas\README.md` |
| Este documento de contexto | `C:\Users\twaaa\Documents\GES-AGREGADOS-CEMEX\MatrizVentas\CONTEXTO_PROYECTO.md` |
| Dependencias fijadas | `C:\Users\twaaa\Documents\GES-AGREGADOS-CEMEX\MatrizVentas\requirements.txt` |
| Archivo fuente de prueba (Snowflake/consolidado) | `C:\Users\twaaa\Downloads\Extracción de Datos Agregados _ Formatos Para Alta de Ruta (ultimo)_2026.xlsm` |
| Cache del archivo fuente ya parseado | `C:\Users\twaaa\Downloads\.matrizventas_cache\` (se crea junto al archivo fuente) |
| Salidas de prueba (algunos CEDIS sueltos) | `C:\Users\twaaa\Documents\GES-AGREGADOS-CEMEX\MatrizVentas\_test_out\` |
| Salidas de prueba (los 102 CEDIS completos) | `C:\Users\twaaa\Documents\GES-AGREGADOS-CEMEX\MatrizVentas\_test_out_all\` |

Las carpetas `_test_out*` son resultados de pruebas que se fueron generando durante este
trabajo — sirven para revisar, pero no son "la" salida oficial de producción todavía.

## 7. Qué se construyó/cambió en el código (`matrizVentas.py`) — detallado

El script viejo (4 archivos SAP, un CEDIS a la vez) **se dejó intacto** — sigue
funcionando igual que antes, sin tocarlo. Todo lo nuevo es **aditivo**, activado con la
bandera `--fuente`.

### 7.1 Entorno

No había Python instalado en esta máquina (solo el stub de Microsoft Store). Se instaló:
- Python 3.12.10 (vía `winget install Python.Python.3.12`).
- `pandas==3.0.5`, `openpyxl==3.1.5`, `xlrd==2.0.2` (vía pip; versiones fijadas en
  `requirements.txt` para reproducir igual en otra máquina).

### 7.2 Nuevo modo `--fuente` (lee directo del archivo consolidado)

Funciones nuevas agregadas (todas en `matrizVentas.py`):

- `EXTRACCION_SHEET_NAMES` — diccionario que mapea nombre lógico → nombre real de hoja
  (`CONT_COMPRA TRAOPE`, `PVTA_MAT VK13`, `PVTA_FTE VK13`, `PESO_VOL`).
- `load_extraccion_raw(path)` — abre el `.xlsm` UNA sola vez (`pd.ExcelFile`) y lee las
  4 hojas. Si falta alguna hoja, lanza un error claro listando cuáles hojas SÍ existen.
- `parse_traope_extraccion(raw)` — parsea TRAOPE por nombre de columna (no por
  posición, a diferencia del parser viejo). Agrega columnas `concat1`/`concat2` y guarda
  en `.attrs` del DataFrame las posiciones de columna reales (`sf_ci`, `nombre_sf_ci`,
  `mat_ci`, `centro_ci`, `destino_ci`, `nombre_destino_ci`, `precio_ci`, `um_precio_ci`,
  `col_offset=0`, `is_extraccion=True`) — esto es clave para que las fórmulas VLOOKUP del
  archivo "Con Formulas" sepan a qué columna apuntar sin tener hardcodeada una posición
  fija (ver 7.4).
- `parse_mp_extraccion(raw)` / `parse_flete_extraccion(raw)` — mapean las columnas reales
  de las hojas a los nombres normalizados internos que ya esperaba el resto del
  pipeline (`sf`, `cedis`, `destino`, `material`, `importe_mp`/`importe_flete`, `um`,
  etc.), para no tener que tocar el resto del código de cruce/escritura.
- `parse_materiales_extraccion(raw)` — usa la columna `Cant. UMB` (columna E) como PV
  directo, sin recalcular (a diferencia del parser viejo de SAP).
- `write_traope_sheet_extraccion(ws, traope_df)` — escribe el sub-sheet TRAOPE
  directamente desde el DataFrame ya en memoria (sin releer el archivo de disco, a
  diferencia del escritor viejo).
- `run_extraccion(fuente, output, cedis_filter, cache_dir, force_refresh)` — orquesta
  todo: carga los datos (con cache, ver 7.5), determina qué CEDIS generar (uno solo si
  se pasa `--cedis`, o todos los distintos encontrados si no), filtra cada DataFrame por
  CEDIS, y genera los 2 archivos de salida por CEDIS en
  `MatrizVentas_Generado\<CEDIS>\salidas\` (mismo layout de carpetas que el modo viejo).
- Nuevos argumentos CLI: `--fuente ARCHIVO`, `--cache-dir DIR`, `--refresh-cache`.

Ejemplos de uso:
```
python matrizVentas.py --fuente "Extracción....xlsm"                      # genera los 102 CEDIS
python matrizVentas.py --fuente "Extracción....xlsm" --cedis D836          # solo un CEDIS
python matrizVentas.py --fuente "Extracción....xlsm" --refresh-cache       # fuerza releer todo
```

### 7.3 Validación de columnas (mensajes de error claros)

Se agregó `_require_columns(df, required, sheet_label)` — antes de usar las columnas de
cada hoja, valida que existan; si falta alguna, lanza un `ValueError` claro en español
diciendo cuál columna falta y cuáles sí se encontraron (en vez de un `KeyError` críptico
de pandas). Se aplicó en `parse_mp_extraccion`, `parse_flete_extraccion`,
`parse_materiales_extraccion`, y una validación equivalente (con nombres lógicos) dentro
de `parse_traope_extraccion`. Se probó simulando una columna faltante/renombrada y el
mensaje sale correcto. Esto es importante porque el archivo real (en la otra
computadora) viene de un query de Snowflake que, si se re-configura, podría cambiar
ligeramente los encabezados.

### 7.4 Fórmulas VLOOKUP dinámicas (generalización necesaria)

El archivo consolidado tiene un layout de columnas distinto al TRAOPE viejo de SAP. Las
fórmulas VLOOKUP del archivo "Con Formulas" (`build_main_sheet_formulas`) originalmente
tenían **posiciones de columna hardcodeadas** (ej. `TRAOPE!$F:$G`, `TRAOPE!$A:$V,22`,
`TRAOPE!$AD:$AE`) que solo eran válidas para el layout del TRAOPE viejo. Se generalizó
con la función `_traope_out_col(traope_df, attr_key)`, que calcula la posición real de
cada columna (Nombre SF, Precio neto, UM Precio Pedido, Nombre Destino) a partir de los
`.attrs` guardados al parsear TRAOPE (funciona igual para el archivo viejo que para el
nuevo). Esto se necesitó modificar en 4 fórmulas: `NOMBRE_SF`, `NOMBRE_DESTINO`,
`IMPORTE_COSTO`, `UM_COSTO`.

### 7.5 Cache del archivo fuente (para que sea rápido)

Leer y parsear las ~47,000 filas del archivo consolidado tomaba ~20 segundos, sin
importar si al final solo se pedía un CEDIS. Se agregó:
- `_source_fingerprint(path)` — tamaño + fecha de modificación del archivo (sin
  hashear el contenido, para que revisar "¿cambió?" sea instantáneo).
- `_cache_path(fuente, cache_dir)` — ruta del cache (`.matrizventas_cache\` junto al
  archivo fuente, o donde diga `--cache-dir`).
- `load_extraccion_cached(fuente, cache_dir, force_refresh)` — si existe un cache válido
  (mismo tamaño/fecha que el archivo fuente actual) lo reusa; si no, parsea todo de
  nuevo y actualiza el cache. Usa `pickle` (seguro aquí porque es un archivo que el
  propio script crea y controla, no un archivo externo/no confiable).
- `EXTRACCION_CACHE_VERSION` — se sube manualmente si algún día cambia cómo se parsean
  los datos, para que un cache viejo con estructura distinta nunca se reuse por error.
- `--refresh-cache` — fuerza ignorar el cache (equivalente al botón "Refresh" que se
  imagina para SharePoint).

Resultado probado: pedir un CEDIS pasó de ~20-25s a ~3s cuando el cache ya existe y es
válido. El cache se invalida solo (sin que nadie tenga que acordarse) en cuanto el
archivo fuente cambia de tamaño/fecha.

### 7.6 Optimización de rendimiento (bug de performance, no de datos)

Al correr el batch completo (102 CEDIS) la primera vez, tardó **26 minutos 16
segundos** — demasiado lento. Se perfiló con `cProfile` y se encontró la causa real:
las funciones auxiliares de escritura (`wv()`/`wf()` dentro de
`build_main_sheet_formulas`/`build_main_sheet_values`) creaban un objeto `Font(size=10)`
**nuevo en cada celda** en vez de reusar uno compartido, y la función `_fill()` hacía lo
mismo con `PatternFill` — esto obliga a openpyxl a comparar/deduplicar estilos por hash
en cada celda, lo cual escala pésimo (a partir de unos miles de filas, esto solo ya
tomaba más de un minuto). **Esto ya existía en el script original** (no algo introducido
por el trabajo de este proyecto), solo que antes nadie lo notaba porque se corría un
CEDIS a la vez de forma manual e interactiva.

**Corrección:** se agregó una constante compartida `FONT_DATA = Font(size=10)` (reusada
en vez de instanciar una nueva por celda) y se cacheó `_fill()` con
`@functools.lru_cache`. Resultado: el CEDIS más grande (D836, ~2,982 filas) pasó de 118s
a ~28s: el batch completo de 102 CEDIS pasó de 26m16s a **6m21s** (~4.1x más rápido).

### 7.7 Bugs de fórmulas encontrados y corregidos (en el archivo "Con Formulas")

Al correr contra datos reales, ~63% de las filas de D836 mostraban `#VALOR!`
(`#VALUE!`) en las columnas de Validación. Se investigó y **se confirmó que NO era un
bug de cruce de llaves** (se verificó a mano contra pandas: de verdad hay MUCHAS
combinaciones Ship From+Destino+Material en MP que no tienen un registro de Flete
correspondiente — es un hueco real de datos, esperado en el archivo nuevo).

El problema real era de diseño de fórmulas, existente desde antes:

- **Bug 1 (corregido):** cuando el VLOOKUP de Flete o PV no encuentra match, su
  `IFERROR(...)` devuelve `""` (texto vacío) en vez de un número. La fórmula de
  `Validación 1` hacía aritmética directa con ese `""` (`MP + Flete`), lo cual en Excel
  truena en `#VALOR!` (sumar texto con número). **Fix:** se envolvió con `N(...)` (que en
  Excel convierte `""` a `0`), consistente con lo que el lado Python (`calc_validation1`)
  ya hacía silenciosamente.
- **Bug 2 (corregido):** al arreglar el Bug 1, aparecieron errores `#N/D` (`#N/A`) en su
  lugar: el patrón `IFERROR(VLOOKUP(concat1,...), VLOOKUP(concat2,...))` para
  `Importe Costo`/`UM Costo` no tenía un `IFERROR` final — si NINGUNA de las dos llaves
  encontraba match en TRAOPE (otro hueco real de datos: hay filas de MP sin condición de
  compra asociada), el `#N/A` del segundo VLOOKUP se propagaba sin capturar. **Fix:** se
  envolvió todo en un `IFERROR(...,"")` adicional, y la fórmula de `Validación 2` se
  cambió a `=IF(Costo="","",Costo-Validación1)` para que, si no hay costo real con qué
  comparar, la celda quede en blanco (dato faltante, no un error de Excel).

Verificado con Excel COM (abrir el archivo, forzar recálculo completo, y usar
`Range.SpecialCells` para detectar celdas con error real — no solo texto con "#" visual,
que puede ser un falso positivo por columnas angostas tipo `########`): **0 celdas con
error real** en los CEDIS probados (D836, DW68, D285 —este último con 0 registros de
TRAOPE y Flete, caso extremo—, DW62).

### 7.8 Validación cruzada manual de los números

Se tomaron filas puntuales del archivo generado y se compararon a mano contra las hojas
fuente originales (usando pandas) para confirmar que:
- El cruce por `Concat1` (SF+Centro+Destino+Material) y el respaldo por `Concat2`
  (SF+Centro+Material, cuando TRAOPE tiene una condición de compra sin destino
  específico — caso real encontrado: `Destino = NaN` en TRAOPE) funcionan correctamente.
- La normalización de tipos (`str_val`, que convierte floats tipo `10000521.0` a
  `"10000521"` igual que un int `10000521`) hace que las llaves de cruce coincidan
  correctamente entre hojas que traen el mismo campo con distinto tipo de dato
  (ej. Material como float64 en PESO_VOL vs int64 en MP).
- El PV y Nombre de Material jalados desde `PESO_VOL` coinciden con el archivo fuente.

### 7.9 Pruebas de robustez (para la mudanza a la otra computadora)

- **Archivo abierto en Excel:** se probó (con una instancia real de Excel, no solo COM
  efímero) que el script puede leer el archivo fuente **mientras está abierto en
  Excel** — Windows permite lectura compartida, no hubo problema.
- **Caveat honesto, NO probado:** no se pudo simular el escenario de leer el archivo
  **mientras Power Query/Snowflake lo está refrescando activamente en ese instante
  exacto** (solo se probó "abierto y quieto"). Recomendación: cuando se configure el
  disparador automático en la otra computadora, que corra unos minutos **después** de un
  refresh conocido, no exactamente durante.
- Se re-corrió el pipeline completo después de cada cambio de robustez para confirmar
  cero regresión (mismos resultados correctos, 0 errores de fórmula).

## 8. Arquitectura interna del sheet principal (para quien edite el código)

- Cada fila del sheet principal corresponde a un registro de **MP** (es la "columna
  vertebral" — cada fila de MP genera una fila de salida).
- Llaves de cruce: `Concat1 = ShipFrom+Centro+Destino+Material`,
  `Concat2 = ShipFrom+Centro+Material` (sin destino, respaldo).
- Columnas de **Venta**: `Importe MP`, `Importe Flete`, `UM Venta`.
- Columnas de **Costo**: `Importe Costo` (de TRAOPE), `UM Costo`.
- **Validación 1** (fórmula de margen esperado):
  - Si `UM Venta = TN` y `UM Costo = TN` → `MP + Flete`.
  - Si `Cond. Exp. = 1` → `PV × MP`.
  - Si no (el "resto") → `(MP + Flete) × PV`.
- **Validación 2** = `Importe Costo - Validación 1` (diferencia real vs esperado);
  se resalta en amarillo si `abs(diferencia) > 1`.
- Dos versiones del sheet: con fórmulas VLOOKUP reales (`build_main_sheet_formulas`,
  para auditar en Excel) y con valores ya calculados en Python
  (`build_main_sheet_values`, la que debería ser más rápida/confiable de abrir).

## 9. Decisiones pendientes / preguntas abiertas

1. **¿Propuesta A, B, o la mezcla híbrida para SharePoint?** — sin decidir aún.
2. **¿Qué disparador automático usar en la otra computadora?** — se ofrecieron 3 opciones
   (programado con Task Scheduler, vigilancia en tiempo real con file watcher, o
   trigger de Power Automate al detectar cambio en SharePoint) — el usuario pidió NO
   configurar nada de esto todavía; primero quiere todo pulido aquí.
3. **¿102 archivos por CEDIS, o un modelo centralizado (una sola fuente de verdad +
   vistas filtradas)?** — parte de la discusión de SharePoint, sin resolver.

## 10. PROBLEMA ABIERTO — sin resolver (lo más urgente a retomar)

El usuario, revisando manualmente algunas de las "bases" generadas, reportó:
> "las validaciones están mal, hay bastantes campos vacíos" — y pidió investigar por qué
> los datos no se estaban extrayendo correctamente.

Se empezó a investigar (con el archivo `Base Cedis D836.xlsx`, el archivo de
**valores/limpio**, no el "Con Formulas") y **se encontraron 2 problemas reales, pero la
corrección NO se llegó a aplicar al código** (el usuario canceló la sesión de trabajo
justo cuando se estaba diagnosticando, antes de escribir el fix). El código today
(`matrizVentas.py`) **todavía tiene este bug sin corregir**:

### 10.1 Bug confirmado: `Validación 2` calcula un número inventado (no error, PEOR)

En el archivo **limpio** (`Base Cedis <CEDIS>.xlsx`, no el "Con Formulas"), la función
`build_main_sheet_values` escribe:
```
Validación 1 → SIGUE siendo una fórmula de Excel (no un valor calculado en Python)
Validación 2 → fórmula simple: =ImporteCosto - Validación1   (SIN blindaje contra blanco)
```

El campo `Importe Costo` en este archivo se escribe como un **valor de Python** (no
fórmula): si TRAOPE no tiene match para esa fila, el valor es `None` → **celda
verdaderamente en blanco** en Excel. El problema: en Excel, una celda en blanco se trata
como **0** en una resta. Entonces, cuando `Importe Costo` está vacío, en vez de que
`Validación 2` salga en blanco (que sería lo correcto — "no se puede validar, falta el
costo real"), la fórmula calcula `0 - Validación1`, dando un **número negativo grande,
con apariencia de resultado válido**, que en realidad es basura. Esto es mucho peor que
un error de Excel (`#VALOR!`) porque **no se nota a simple vista** — parece un dato real.

**Se confirmó que esto afecta muchas filas:** en D836 (2,982 filas), **356 filas (~12%)**
tienen `Importe Costo` en blanco — cada una de esas probablemente muestra un
`Validación 2` fabricado/incorrecto en vez de estar en blanco.

Este bug **NO existe** en el archivo "Con Formulas" (ahí ya se blindó con
`=IF(Costo="","",Costo-Validación1)`, ver sección 7.7) — el bug es específico del
archivo de **valores/limpio**, que además es probablemente el que más usa la gente
porque es "más rápido de abrir".

### 10.2 Inconsistencia sospechada (identificada, no confirmada al 100%, no corregida)

Comparando las dos formas de calcular `Validación 1`:

- **Lado Python** (`calc_validation1`, usado hoy solo para decidir el resaltado
  amarillo, su resultado se descarta): si `PV` es `None` (no se encontró en
  `PESO_VOL`/Materiales), **asume `PV = 1.0`** como valor por default.
- **Lado fórmula de Excel** (`validation_formula`, usada en ambos archivos hoy): tras el
  fix de la sección 7.7, si `PV` no se encuentra, `N(PV)` lo convierte a **`0`**.

O sea: **cuando falta el PV de un material, Python asume que el factor de conversión es
1, y Excel asume que es 0** — dos supuestos distintos, ambos arbitrarios/inventados, que
dan resultados de `Validación 1` distintos para la misma fila según qué mecanismo la
calculó. Ninguno de los dos es realmente "correcto": si no sabemos el PV real, no
deberíamos inventar un número — lo más honesto sería dejar `Validación 1`/`Validación 2`
en blanco para esa fila (igual que ya se hace cuando falta `Importe Costo`), en vez de
adivinar 0 o 1.

### 10.3 Fix propuesto (diagnosticado, NO implementado todavía)

1. **Hacer que `Validación 1` y `Validación 2` en el archivo LIMPIO sean valores
   calculados en Python de verdad** (usando la función `calc_validation1` que YA existe
   y ya se calcula — hoy nada más se usa para decidir el amarillo y se descarta —, en
   vez de escribir una fórmula de Excel ahí). Esto:
   - Elimina el riesgo de que aparezcan en blanco si Excel no recalcula automáticamente
     (ej. si el modo de cálculo de Excel de alguien está en manual).
   - Permite manejar el caso "falta el dato" con precisión en Python (`None` real) en
     vez de que Excel trate un blanco como 0 en una resta.
2. **Unificar el criterio de "falta PV"**: decidir que, igual que con `Importe Costo`,
   si falta el PV y la rama de la fórmula lo necesita (la rama `Cond.Exp.=1` y la rama
   "resto" sí lo usan; la rama `UM=TN,UM=TN` NO usa PV), el resultado debe quedar en
   blanco/`None`, no inventar 0 ni 1. Aplicar esto tanto en `calc_validation1` (Python)
   como en `validation_formula` (string de fórmula de Excel, para que el archivo "Con
   Formulas" siga siendo consistente con el archivo limpio).
3. Después de aplicar el fix, **repetir todo el proceso de verificación ya establecido**:
   recompilar, correr contra el archivo real, abrir con Excel COM y usar
   `SpecialCells` para confirmar 0 errores, y volver a contar cuántas filas quedan
   correctamente en blanco (vs. antes, cuántas mostraban el número fabricado) para
   confirmar que el bug realmente desapareció.

**Este es el punto exacto donde retomar la próxima sesión.**

## 11. Notas / advertencias generales a tener en cuenta

- El script viejo (4 archivos SAP) sigue funcionando igual, sin cambios.
- El cache (`--fuente`) es solo para acelerar la LECTURA del archivo fuente; el tiempo
  de ESCRIBIR un CEDIS grande sigue siendo proporcional a su tamaño (unos 8-10s para el
  CEDIS más grande, D836).
- Es **normal y esperado** que algunas filas de MP no tengan Flete o Costo asociado —
  es un hueco real de datos en SAP/Snowflake, no un bug del cruce. Lo que SÍ es un bug
  es que ese hueco se muestre como un número inventado en vez de en blanco (ver sección
  10).
- No se ha tocado nada relacionado con SharePoint todavía — todo lo construido corre
  100% local, en la computadora de pruebas.
- No se ha configurado ningún disparador automático (ni Task Scheduler ni file watcher)
  — pendiente, a propósito, hasta que se decida en la otra computadora.
