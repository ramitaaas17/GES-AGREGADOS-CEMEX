# CEMEX Agregados — Matriz de Precios Venta vs Costo & Gobernanza

Sistema automatizado de conciliación comercial, cálculo de márgenes operativos (**MOP %**), auditoría de costos de compra (**TRAOPE**) y matriz de gobernanza y autorizaciones para **CEMEX Agregados**.

---

## 📑 Tabla de Contenidos
1. [¿Qué hace este proyecto?](#-qué-hace-este-proyecto)
2. [Evolución y Actualizaciones Recientes](#-evolución-y-actualizaciones-recientes)
3. [Arquitectura de Archivos (¿Qué hace cada cosa?)](#-arquitectura-de-archivos-qué-hace-cada-cosa)
4. [Cómo se Ejecuta](#-cómo-se-ejecuta)
   - [Opción A: 1-Clic desde Excel (Recomendado para Usuarios Finales)](#opción-a-1-clic-desde-excel-panel_control_cemexxlsm)
   - [Opción B: Vía Consola Python](#opción-b-vía-consola-python)
5. [Explicación de Negocio: Validación de Margen y MOP](#-explicación-de-negocio-validación-de-margen-y-mop)
   - [1. Margen Operativo de Material (MOP %) Puro](#1-margen-operativo-de-material-mop--puro)
   - [2. ¿Qué es Validación 1 (Costo Teórico Esperado)?](#2-qué-es-validación-1-costo-teórico-esperado)
   - [3. ¿Qué es Validación 2 (Brecha Costo Real vs Venta)?](#3-qué-es-validación-2-brecha-costo-real-vs-venta)
   - [4. Matriz de Gobernanza y Autorizaciones (Trading vs Filial)](#4-matriz-de-gobernanza-y-autorizaciones)
6. [Nota de Revisión y Colaboración](#-nota-de-revisión-y-colaboración)

---

## 🎯 ¿Qué hace este proyecto?

El sistema cruza de manera automatizada las bases de datos extraídas de **Snowflake / SAP** para los 102 CEDIS de CEMEX Agregados en México:
* **`PVTA_MAT VK13`**: Precios de venta de materia prima por ruta (Sociedad 7100 / 7180).
* **`PVTA_FTE VK13`**: Precios de flete de venta y condición de expedición (`1` = Entregado, `2` = Recogido).
* **`CONT_COMPRA TRAOPE`**: Contratos y órdenes de compra con proveedores / canteras (Costo Total y Flete de Compra).
* **`PESO_VOL`**: Factores de densidad / peso volumétrico ($PV = kg/m^3 \rightarrow TN$).
* **`CONT_VTA ZSDD4501`**: Contratos de venta vigentes y precios pactados con clientes.

Genera un libro consolidado en Excel con un **Dashboard Ejecutivo Corporativo** (KPIs, estatus de autorizaciones, materiales más rentables y tabla de auditoría) y una **Hoja Matriz** de 33 columnas completamente desglosada y validada.

---

## 🚀 Evolución y Actualizaciones Recientes

Desde el inicio de este sprint, transformamos el flujo manual en una plataforma ejecutiva:

| Componente | Antes | Ahora (Versión Actual) |
|---|---|---|
| **Motor de Cálculo** | Script monolítico con fórmulas lentas de Excel | Motor optimizado en Python (`matriz_integrada.py`) con valores directos y caché en memoria (~10s). |
| **Cálculo de MOP** | Consideraba el costo total de compra sin descontar fletes | **MOP Puro sobre Material:** descuenta el flete de compra (`TRAOPE`) para evaluar margen exclusivamente sobre la roca/arena. |
| **Gobernanza Filial (`7100`)** | Generaba alertas erróneas de margen bajo | Clasificado como **`No Aplica (Filial)`** con `MOP = 0.0%` (transferencia intercompañía a costo). |
| **Gobernanza Trading (`7180`)** | Reglas no estandarizadas | Matriz oficial: **Champion** ($>8\%$), **Regional** ($5\%-8\%$) y **Alerta Nacional** ($<5\%$). |
| **Interfaz de Usuario** | Comandos manuales en terminal | **Panel de Control en Excel (`.xlsm`)** con botón interactivo y ventana emergente (*UserForm*) con buscador en tiempo real. |
| **Dashboard** | No existía / Formato plano | **Dashboard Ejecutivo Premium** estilo CEMEX UI como primera pestaña del reporte. |
| **Compatibilidad Excel** | Errores de "archivo dañado" por caracteres `=` | Encabezados limpios y sanitizados, abriendo al instante sin advertencias ni reparaciones. |

---

## 📁 Arquitectura de Archivos (¿Qué hace cada cosa?)

```
GES-AGREGADOS-CEMEX/
│
├── PANEL_CONTROL_CEMEX.xlsm      # Interfaz principal de Excel para usuarios finales (con macros y UserForm)
├── matriz_integrada.py           # Orquestador unificado en Python (Procesamiento, MOP, Gobernanza y Dashboard)
├── build_panel_control_xlsm.py   # Script automatizado que compila y genera PANEL_CONTROL_CEMEX.xlsm
├── ModSelectorCEDIS.bas          # Módulo VBA exportable del selector multi-CEDIS
│
├── Extracción de Datos...xlsm    # Archivo fuente con las 4 hojas maestras de Snowflake/SAP
├── requirements.txt              # Dependencias de Python (pandas, openpyxl, pywin32)
└── _salidas_integradas/          # Carpeta donde se guardan automáticamente los reportes generados
```

---

## 💻 Cómo se Ejecuta

### Opción A: 1-Clic desde Excel (`PANEL_CONTROL_CEMEX.xlsm`)
*Recomendado para usuarias y directivos (Ana, Lorena, etc.) sin necesidad de tocar código:*

1. Abre el archivo [`PANEL_CONTROL_CEMEX.xlsm`](./PANEL_CONTROL_CEMEX.xlsm).
2. Haz clic en el botón verde **`▶ SELECCIONAR CEDIS Y GENERAR MATRIZ`**.
3. Se abrirá la ventana emergente (*UserForm*):
   - Puedes buscar o filtrar en la caja de texto.
   - Marcar uno o varios CEDIS con las casillas.
   - O usar botones rápidos: **`🎯 Demo D836/D838`** o **`☑ Marcar Todos`**.
4. Pulsa **`🚀 GENERAR MATRIZ Y DASHBOARD`**.
5. El sistema procesará los datos en segundo plano y abrirá automáticamente el reporte final generado en `_salidas_integradas\`.

### Opción B: Vía Consola Python
Para ejecuciones automatizadas por línea de comandos:

```bash
# Procesar CEDIS específicos (ej. D836 y D838)
python matriz_integrada.py --fuente "Extracción de Datos Agregados _ Formatos Para Alta de Ruta (ultimo)_2026.xlsm" --cedis D836 D838

# Procesar todos los 102 CEDIS del país
python matriz_integrada.py --fuente "Extracción de Datos Agregados _ Formatos Para Alta de Ruta (ultimo)_2026.xlsm"

# Forzar recarga ignorando la memoria caché
python matriz_integrada.py --fuente "Extracción de Datos Agregados _ Formatos Para Alta de Ruta (ultimo)_2026.xlsm" --refresh-cache
```

---

## 🧠 Explicación de Negocio: Validación de Margen y MOP

*Esta sección detalla la lógica comercial y operativa detrás de cada cálculo implementado en el sistema:*

```
                               ┌──────────────────────────────────────────────┐
                               │       FLUJO COMERCIAL CEMEX AGREGADOS        │
                               └──────────────────────┬───────────────────────┘
                                                      │
                       ┌──────────────────────────────┴──────────────────────────────┐
                       ▼                                                             ▼
         【 SOCIEDAD 7100: FILIAL 】                                   【 SOCIEDAD 7180: TRADING 】
         • Canteras Propias CEMEX                                       • Compra a Proveedores Terceros
         • Transferencia Intercompañía                                  • Margen Comercial Directo
         • MOP = 0.0% (Operación a Costo)                               • MOP % = (MP Venta - MP Compra) / MP Venta
         • Gobernanza: "No Aplica (Filial)"                             • Gobernanza: Champion / Regional / Nacional
```

---

### 1. Margen Operativo de Material (MOP %) Puro

El **MOP %** mide la rentabilidad que le queda a CEMEX por comercializar el agregado mineral, **aislando el costo del transporte (flete)**.

$$\text{MOP \%} = \frac{\text{Precio Venta Material (MP)} - \text{Costo Compra Material Puro}}{\text{Precio Venta Material (MP)}}$$

#### ¿Por qué descontamos el flete de compra en TRAOPE?
En las órdenes de compra (`CONT_COMPRA TRAOPE`), el campo `Precio neto pedido` a menudo incluye el servicio de flete si el proveedor entrega en planta. Si no descontáramos ese flete (`Importe Condición`), estaríamos comparando el precio del material de venta contra un costo de compra inflado por transporte, arrojando márgenes falsamente negativos o distorsionados.

#### Homologación de Unidades de Medida ($TN \leftrightarrow M3$):
Los clientes a veces compran en **Toneladas (TN)** pero los proveedores cobran en **Metros Cúbicos (M3)** (o viceversa). Para poder calcular el margen:
* Si Venta es en $TN$ y Compra en $M3$: $\text{Costo Homologado} = \frac{\text{Costo } M3}{PV}$
* Si Venta es en $M3$ y Compra en $TN$: $\text{Costo Homologado} = \text{Costo } TN \times PV$
*(donde $PV$ es la densidad de la roca en $TN/m^3$, ej. 1.262)*.

---

### 2. ¿Qué es Validación 1 (Costo Teórico Esperado)?

Representa **cuánto debería costar o valer la ruta completa** según las condiciones de venta pactadas con el cliente y la modalidad logística de entrega:

$$\text{Validación 1} = \begin{cases} 
\text{MP Venta} + \text{Flete Venta} & \text{si UM Venta y Costo son en } TN \\
\text{PV} \times \text{MP Venta} & \text{si Condición de Expedición = 1 (Entregado)} \\
(\text{MP Venta} + \text{Flete Venta}) \times \text{PV} & \text{en las demás modalidades mixtas}
\end{cases}$$

* **Condición de Expedición 1 (Entregado en Obra):** CEMEX gestiona la logística integral.
* **Condición de Expedición 2 (Recogido en Cantera):** El cliente envía sus propios camiones; por lo tanto, no se suma flete de venta.

---

### 3. ¿Qué es Validación 2 (Brecha Costo Real vs Venta)?

Representa la **desviación económica** entre lo que nos cuesta comprar el material en SAP vs lo que proyectamos venderle al cliente:

$$\text{Validación 2} = \text{Costo Total Compra (TRAOPE)} - \text{Validación 1}$$

* **$\text{Validación 2} = 0.00$:** Alineación perfecta entre compra y venta.
* **$\text{Validación 2} > 1.00$ (Alerta Amarilla `DIFERENCIA`):** El costo de compra cargado en SAP supera al costo teórico de venta en más de $1.00 peso. Esto alerta una posible fuga de margen o un contrato de compra desactualizado respecto a la lista de precios de venta.

---

### 4. Matriz de Gobernanza y Autorizaciones

| Sociedad | Clasificación | Regla de Margen MOP | Nivel de Autorización | Acción Requerida |
| :---: | :---: | :---: | :---: | :--- |
| **`7100`** | Canteras Propias | $MOP = 0.0\%$ | ⚪ **No Aplica (Filial)** | Operación intercompañía estándar a costo. |
| **`7180`** | Trading Terceros | $MOP > 8.0\%$ | 🟢 **Champion** | Aprobación comercial directa en sistema. |
| **`7180`** | Trading Terceros | $5.0\% \le MOP \le 8.0\%$ | 🟡 **Regional** | Requiere justificación documentada y Vo.Bo. Regional. |
| **`7180`** | Trading Terceros | $MOP < 5.0\%$ | 🔴 **Alerta Nacional** | Margen crítico; requiere Vo.Bo. de Dirección Nacional. |

---

## 🤝 Nota de Revisión y Colaboración

> **Para el compañero / equipo de desarrollo y negocio:**
> 
> Esta versión unificada mantiene intacta la lógica de cálculo que venías trabajando para `Validación 1` y `Validación 2`, pero ahora con:
> 1. Eliminación de micro-residuos decimales (redondeo a 2 decimales y tolerancia limpia).
> 2. Separación explícita de `7100` (Filial) vs `7180` (Trading) para que el Dashboard no se contamine con falsas alertas.
> 
> 💡 **Punto a revisar juntos:** Si detectas que en alguna zona/planta específica la fórmula de `Validación 1` deba tomar alguna consideración adicional para modalidades especiales de flete (ej. fletes escalonados o flete incluido en precio neto), avísame para incorporarlo en la matriz de reglas.
