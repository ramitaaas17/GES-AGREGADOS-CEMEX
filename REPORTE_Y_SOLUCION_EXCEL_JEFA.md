# REPORTE Y SOLUCIÓN – Excel de la jefa (`Extracción de Datos Agregados _ Formatos Para Alta de Ruta (ultimo)_2026.xlsm`)

> Análisis **solo lectura** sobre una copia. SHA256 verificado antes/después. Pico de RAM del script: ~26 MB (streaming con `zipfile` + `iterparse`).
> Evidencia completa en `diagnostico_excel\` (`diagnostico_principal.txt`, `vba_extracted.bas`, `olevba_report.txt`, `diag__recuperacion_tmp_*.txt`).

## 0. Mapa del libro

| Hoja | codeName (VBA) | Notas |
|---|---|---|
| BASE 2026 | `Sheet4` | 3,941 filas × 67 cols, **154,281 fórmulas**, XML de 45.9 MB (de ~92 MB descomprimidos) |
| FPA FILIAL | `Hoja1` | Filas 1-5 encabezados, **fila 6 = fila modelo con formato**, fila 7 **vacía**, datos desde fila 8 |
| FPA TERCEROS | `Hoja6` | Misma estructura |
| CTO COMPRA | `Hoja8` | |
| PVTA_MAT, PVTA_FTE, CONT_VTA, CONT_COMPRA | – | Tablas cargadas por Power Query (`connections.xml`, 4 `queryTables`), ~37 MB de XML |

El flujo Python equivalente (`etl_processor.py` L1418-1422) genera FPA FILIAL con `ORGVTA == "7100"`, `ZMAH/ZMPH`, `Valido hasta = '31.12.9999'`, 24 columnas. La macro VBA hace lo mismo a mano.

---

## 1. BUG A – Formato en 7100 (`Hoja1.cls`, `ActualizarFilialSolo7100`)

### Causa raíz (confirmada en el XML de la hoja)
```vba
FilaDestino = 8                                   ' L143
...
wsFilial.Rows(FilaDestino).Insert Shift:=xlDown   ' L149
```
`Insert` sin `CopyOrigin` usa `xlFormatFromLeftOrAbove` → la fila nueva **hereda el formato de la fila 7**. En `sheet2.xml` la **fila 7 no existe** (celdas General, sin estilo). El formato real (estilos 243/263/266/268/269/274/276/277/244, `%`, `$`, merges `A:B`/`C:D`, formato condicional de F/G/H) vive en la **fila 6 (plantilla)** y en la fila 8 actual. Resultado: cada fila nueva queda en General, sin formato condicional y sin números/porcentajes/moneda; la fila 8 formateada solo se "empuja" hacia abajo.

### Fallas adicionales en la misma macro
| Línea | Problema |
|---|---|
| L162 `If ValorBase <> 0 And Trim(ValorBase) <> ""` | Con texto da **Error 13 (Type mismatch)** y la macro se interrumpe a medias (con `ScreenUpdating=False` y sin `On Error`). |
| L116/L200 | Sin `On Error`: si falla, deja `ScreenUpdating` apagado. |
| L172 `X = "31.12.9999"` en celda General | Depende de configuración regional; en filas con formato fecha se convierte en fecha. Debe ir como texto. |
| L149 dentro del bucle | Una inserción de fila por registro: lento. |
| L130 `OrgVentas = 7100` | Si la celda es texto `"7100"` la comparación varía según el tipo; se unifica con `CStr`. |
| `Hoja6.cls` (7180/7277) | **Misma falla** (L276). Se corrige con el mismo helper. |

Los CF de FPA FILIAL están anclados a celdas sueltas (`F6,G6,H6,F8,G8,H8`) en vez de a un rango, lo que confirma el copiado manual de formato.

**Solución:** `vba_corregido.bas` (Sección A/B): inserta el bloque completo, copia **solo formatos** desde la fila 6 con `PasteSpecial xlPasteFormats` (incluye CF, merges y NumberFormat) y escribe los valores con validación segura.

---

## 2. BUG B – `#¿NOMBRE?` y "pide recuperación"

### Hallazgos (workbook.xml / sheets)
| # | Hallazgo | Evidencia |
|---|---|---|
| 1 | **Funciones modernas en 154k fórmulas**: `_xlfn.XLOOKUP` ×51,181; `_xlfn.TEXTJOIN` ×27,559; `_xlfn._xlws.FILTER` ×11,811; `_xlfn.IFS` ×8,769 | `BASE 2026` (cols A, AD, AJ, AR, AS, AT, AW, AX, BA, BC, BF…) |
| 2 | **Nombres ocultos corruptos**: `_xleta.IF`, `_xleta.IFERROR`, `_xleta.TODAY` = `#NAME?` (`xlm="1"`), y `_xlnm._FilterDatabase` de BASE 2026 = `#REF!` | `workbook.xml` `<definedNames>` (también en las 3 copias `_recuperacion_tmp`) |
| 3 | Fórmulas con `#REF!` | `BASE 2026!X1381, Z3687, AA3687`; `CTO VTA!D3,H3,J3,M3…` = `'FPA TERCEROS'!#REF!`; en la copia de recuperación `BO364:BO368…` = `XLOOKUP(…=#REF!…)` |
| 4 | Rangos de 27,651 / 29,987 filas en fórmulas con solo 3,941 filas de datos | 11,607 fórmulas con `$27651` |
| 5 | Partes exclusivas de Excel 365: `threadedComments`, `persons`, `namedSheetViews`, `metadata.xml` (matrices dinámicas), `workbookPr updateLinks="always"` | listado del ZIP |
| 6 | `calcChain.xml` 3.81 MB; `dxfs` 238 (estilos CF acumulados); sin enlaces externos (`externalLinks`=0) | `diagnostico_principal.txt` |
| 7 | **Python escribe sobre el maestro**: `xlwings` con `app.calculation="manual"` + `wb.save()` (`etl_processor.py` L526/L639, `fpa_validations.py` L857/L894, `recovery_worker.py` L125) y `sync_manager.py` copia el archivo con `shutil.copy2/copyfile` | código del proyecto |

### Diagnóstico
- **`#¿NOMBRE?`**: el XML guarda `_xlfn.XLOOKUP/FILTER/TEXTJOIN/IFS`. Un Excel que no los conoce (2016/2019, Excel Online antiguo o un equipo sin actualizar) muestra `#¿NOMBRE?`, y **al guardar desde ahí queda persistido**. A esto se suman los nombres `_xleta.*` = `#NAME?` (restos de funciones tipo LAMBDA/η) y los `#REF!`.
  - El VBA **no** inyecta funciones modernas: sus `.Formula =` son simples (`=Q5/K5`) → no es la causa; sí conviene usar `Formula2` si se agregan funciones nuevas por código.
- **Corrupción al abrir** (hipótesis más probable, a confirmar con la jefa): guardado interrumpido o copia del archivo **mientras Excel lo tiene abierto/guardando** (xlwings `save` + `sync_manager` copiando el binario + archivos `~$…` presentes en la carpeta del proyecto), con el libro de 9.7 MB / 45 MB de XML en cálculo manual. Las 3 copias `_recuperacion_tmp` ya arrastran los mismos defectos (nombres rotos).
- **Pendiente de confirmar**: versión de Excel de la jefa (`Archivo > Cuenta`). Si es 2016/2019, ese es el origen del `#¿NOMBRE?`.

**Solución (VBA):** `SoportaFuncionesModernas` + `Workbook_Open` (avisa) + `Workbook_BeforeSave` (bloquea guardado desde un Excel incompatible), `LimpiarNombresRotos`, `RepararNombreEnFormulas` (re-introduce con `Formula2`). Ver `vba_corregido.bas`.

---

## 3. Rendimiento y peso (datos reales)

| Concepto | Valor | Impacto |
|---|---|---|
| `BASE 2026` XML | 45.9 MB (2.84 MB zip) | Calcula y abre lento |
| XLOOKUP con rango `$5:$27651` ×2 matrices | 11.6k fórmulas | Coste O(n × 27k) por fórmula |
| `TEXTJOIN` repetido para formar la llave | 27.5k | Se recalcula en cada XLOOKUP |
| PVTA_MAT / PVTA_FTE / CONT_VTA / CONT_COMPRA | 37 MB XML, ~4.7 MB zip | Datos duplicados (tabla + query con `saveData=1`) |
| `calcChain.xml` | 3.8 MB | Excel lo reconstruye si se elimina |
| Celdas vacías con estilo en BASE | 19,173 | Menor |
| `UsedRange` vs real | Coincide, salvo CONT_VTA (`V` vs real `S`) | Menor |
| CF | 7 reglas en BASE, 0-6 en otras | **No es problema** |
| `Worksheet_Change` de BASE 2026 | Recorre ~200k celdas por edición | **Lentitud al editar** |

---

## 4. Plan de optimización (3 pasos, interfaz intacta)

1. **Limpieza sin cambio funcional (hacer en copia)**: aplicar `LimpiarNombresRotos`, corregir los `#REF!` listados, reemplazar `Worksheet_Change` por la versión de la Sección D, y guardar con *Guardar como* limpio (regenera `calcChain`). Cambiar los rangos `$5:$27651` por referencias de tabla (`Solicitudes[…]`) — misma lógica, ~7× menos celdas por XLOOKUP.
2. **Conectar los CSV de RutaV3_2026**: en Power Query, apuntar las consultas `PVTA_MAT`, `PVTA_FTE`, `CONT_VTA`, `CONT_COMPRA` a los CSV (`Data > Get Data > From Text/CSV`, ruta fija en carpeta compartida o parámetro `RutaCSV`). Las tablas y hojas se conservan con el mismo nombre/columnas, así los XLOOKUP no cambian. Desactivar *Refrescar al abrir* y dejar un botón/atajo para actualizar. `fpa_filial.csv` e `input_powerquery.csv` se pueden consumir igual como consulta para validar contra FPA FILIAL, sin pegar a mano. *(Verificar en el Editor de Power Query el origen actual de cada consulta: está en `DataMashup`, que este análisis no abrió.)*
3. **Proceso de guardado seguro**: que Python **no** escriba/copie el maestro mientras Excel lo tiene abierto (`~$…`), usar una copia de trabajo y reemplazo atómico; evitar `calculation="manual"` persistente (restaurar a automático en `finally`). Mantener `.xlsm`; `.xlsb` solo si la jefa acepta el cambio de extensión (reduce peso ~50%).

---

## 5. Código VBA corregido
Archivo listo para copiar: **`vba_corregido.bas`** (instrucciones de pegado en la cabecera). Contiene `CargarFPA` (7100 y 7180/7277), herramientas anti `#NOMBRE?`, `Workbook_Open/BeforeSave` y `Worksheet_Change` optimizado.

> [!WARNING]
> No se pudo ejecutar VBA (no se modificó el Excel). Probar en una copia: 1) correr la macro 7100 con 2-3 filas nuevas, 2) verificar formatos de las 24 columnas, merges y CF, 3) abrir/guardar 3 veces.
