# Matriz Ventas

Genera la "Base Cedis DW##" a partir de 4 archivos de entrada de SAP, cruzando precios de
materia prima, fletes, condiciones de compra (TRAOPE) y el catálogo de materiales.

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

**Directo por terminal:**
```
python matrizVentas.py "C:\ruta\a\la\carpeta"
python matrizVentas.py --mp MP.xlsx --flete Fletes.xlsx --traope TRAOPE.xlsx --materiales MAT.xlsx --output "C:\salida" --cedis DW88
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
