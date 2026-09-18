# Matriz Ventas

Genera la "Base Cedis DW##" a partir de 4 archivos de entrada de SAP, cruzando precios de
materia prima, fletes, condiciones de compra (TRAOPE) y el catálogo de materiales.

## Requisitos (para instalar en una computadora nueva)

```
python -m pip install -r requirements.txt
```

Necesita Python 3.10+ y las librerías fijadas en `requirements.txt` (mismas versiones con
las que se probó todo). No hace falta nada más — el script no necesita que Excel esté
instalado ni cerrado (puede leer el archivo fuente aunque este abierto en Excel).

## Entradas (4 archivos, en una misma carpeta)

El script detecta cada archivo por su nombre (no importa el orden ni el nombre exacto,
basta con que contenga la palabra clave):

| Archivo | Palabra clave en el nombre | Contenido |
|---|---|---|
| MP / precios de venta | `mp`, `precios de venta`, `VK13` (sin "flete") | Precios de materia prima por CEDIS/destino/material |
| Fletes | `flete(s)`, `VK13 flete` | Importes de flete por CEDIS/destino/material |
| TRAOPE | `traope` | Órdenes de compra / condiciones de precio neto por Ship From |
| Materiales | `material(es)` | Catálogo de materiales (PV = Contador / Denom, solo Denom=1000) |

Soporta tanto el formato "columnar" clásico como el formato "vertical apilado" más nuevo
que a veces exporta SAP para MP y Fletes (se detecta automáticamente).

## Organización de carpetas

Cada CEDIS tiene su propia carpeta en `MatrizVentas_Generado\<CEDIS>\`, con dos
subcarpetas: **`entradas\`** (los 4 archivos crudos de ese CEDIS) y **`salidas\`** (los 2
archivos "Base Cedis" que genera el proceso) — para no mezclar insumos con resultados.

## Salidas (2 archivos)

- **`Base Cedis DW## (Con Formulas).xlsx`** — con fórmulas VLOOKUP/BUSCARV reales apuntando
  a las hojas TRAOPE / Flete / Materiales (útil para auditar/verificar a mano en Excel).
- **`Base Cedis DW##.xlsx`** — la misma información pero con los valores ya calculados en
  Python (más rápido de abrir, y con las validaciones de margen resaltadas en amarillo
  cuando la diferencia es mayor a 1).

El código CEDIS (`DW##`) se detecta automáticamente del nombre de los archivos; si no lo
encuentra, usa `DW00`.
## Cómo correrlo

**Desde el menú** (`python ..\menu.py`, opción 1): solo pide la carpeta con los 4 archivos
y dónde guardar el resultado.

**Directo por terminal (4 archivos crudos de SAP, un CEDIS a la vez):**
```
python matrizVentas.py "C:\ruta\a\la\carpeta"
python matrizVentas.py --mp MP.xlsx --flete Fletes.xlsx --traope TRAOPE.xlsx --materiales MAT.xlsx --output "C:\salida" --cedis DW88
```

**Directo por terminal (archivo consolidado de Snowflake, todos los CEDIS o uno solo):**
```
python matrizVentas.py --fuente "Extracción de Datos Agregados....xlsm"
python matrizVentas.py --fuente "Extracción de Datos Agregados....xlsm" --cedis D836 --output "C:\salida"
```

## Fuente Snowflake (archivo consolidado)

Además de los 4 archivos crudos de SAP, el script puede leer directamente un único
archivo `.xlsm`/`.xlsx` que ya trae, en un solo libro, las 4 hojas que alimenta el
archivo que se arma con datos de Snowflake:

| Hoja del archivo consolidado | Equivale a |
|---|---|
| `CONT_COMPRA TRAOPE` | TRAOPE |
| `PVTA_MAT VK13` | MP |
| `PVTA_FTE VK13` | Flete |
| `PESO_VOL` | Materiales (usa el PV ya calculado en la columna "Cant. UMB", no se recalcula) |

Este archivo trae **todos los CEDIS a la vez** (Sacos/Intergiros con clases de condición
ZMAH/ZMPH y Agregados/Terceros con ZMA6/ZMP1, mezclados). Por eso, al usar `--fuente` sin
`--cedis`, el script genera automáticamente un "Base Cedis" por cada CEDIS distinto que
encuentra, guardando cada uno en su propia carpeta
`MatrizVentas_Generado\<CEDIS>\salidas\`, igual que en el flujo de archivos crudos.

Es normal que algunas filas de MP no tengan Flete o Costo (TRAOPE) asociado — es un hueco
real de datos (esa combinación Ship From/Destino/Material aún no tiene flete o precio de
compra cargado en SAP), no un error del cruce. En ese caso el archivo "Con Formulas" deja
esas celdas en blanco en vez de mostrar un error de Excel.

### Cache del archivo fuente

Leer y parsear el archivo consolidado completo (~47,000 filas entre las 4 hojas) toma
unos ~20 segundos. Para no pagar ese costo cada vez que alguien pide un solo CEDIS, el
script guarda los datos ya parseados en un cache en disco (`.matrizventas_cache\` junto
al archivo fuente, o donde indique `--cache-dir`). Mientras el archivo fuente no cambie
de tamaño/fecha de modificación, las corridas siguientes reusan ese cache — pedir un
CEDIS pasa de ~20s a ~3s.

- El cache se invalida solo en cuanto el archivo fuente se vuelve a guardar/actualizar
  (se compara tamaño + fecha de modificación).
- `--refresh-cache` fuerza a ignorar el cache y releer todo, por si se necesita forzar
  una actualización con certeza (equivalente al botón "Refresh" de la idea de SharePoint).
- Nota: el cache acelera la LECTURA del archivo fuente; el tiempo de ESCRIBIR el Base
  Cedis de un CEDIS grande (miles de filas) sigue siendo proporcional a su tamaño.

```
python matrizVentas.py --fuente "Extracción....xlsm" --cedis D836   # usa cache si existe y es valido
python matrizVentas.py --fuente "Extracción....xlsm" --refresh-cache # fuerza releer todo
```

## Flujo interno (resumen)

1. Detecta y parsea los 4 archivos (`parse_mp`, `parse_flete`, `parse_traope`,
   `parse_materiales`).
2. Construye diccionarios de búsqueda en Python (`build_traope_lookups`,
   `build_flete_lookups`, `build_mat_lookups`) para no depender de que Excel recalcule.
3. Arma el sheet principal (Concat1/Concat2 como llaves de cruce SF+Cedis+Destino+Material),
   con columnas de Venta (Importe MP, Importe Flete) y Costo (Importe Costo, UM Costo).
4. Calcula "Validación 1" (fórmula de margen: TN usa MP+Flete; Cond.Exp=1 usa PV×MP; el resto
   usa (MP+Flete)×PV) y "Validación 2" (diferencia contra el costo real), resaltando en
   amarillo cuando la diferencia supera 1.
5. Escribe las hojas de apoyo (TRAOPE, Flete, Materiales, MP) tal cual, para que las fórmulas
   del archivo "Con Formulas" tengan de dónde jalar.
