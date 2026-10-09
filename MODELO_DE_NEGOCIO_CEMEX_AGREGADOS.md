# Documento Maestro: Modelo de Negocio, Operación y Gobernanza Comercial
## CEMEX Agregados México

---

## 1. Visión General de la Línea de Negocio

La división de **Agregados de CEMEX** es la columna vertebral del suministro de materiales pétreos (arenas, gravas, balasto, tepetates, bases hidráulicas y calizas trituradas) tanto para la infraestructura nacional como para la cadena de suministro interna del grupo.

El negocio opera bajo dos grandes objetivos estratégicos:
1. **Abastecimiento Interno (Autoconsumo Estratégico):** Garantizar materia prima pétrea homogénea, certificada y a costo eficiente para las plantas de concreto premezclado de CEMEX en todo el país.
2. **Comercialización y Trading a Terceros (Venta Comercial):** Vender agregados a clientes externos, grandes constructoras, proyectos de infraestructura pública, desarrolladores y distribuidores de materiales de construcción, maximizando el margen operativo del recurso no renovable extraído.

---

## 2. Ecosistema de Sociedades y Modelos Operativos

En los sistemas SAP y Snowflake, la operación se segrega principalmente en tres **Organizaciones de Ventas / Sociedades**, cada una con reglas de negocio y estructuras financieras distintas:

```
                      ┌────────────────────────────────────────┐
                      │        CEMEX AGREGADOS MÉXICO          │
                      └──────────────────┬─────────────────────┘
                                         │
         ┌───────────────────────────────┼───────────────────────────────┐
         ▼                               ▼                               ▼
┌─────────────────────────┐  ┌─────────────────────────┐  ┌─────────────────────────┐
│     SOCIEDAD 7100       │  │     SOCIEDAD 7180       │  │     SOCIEDAD 7277       │
│  (CEMEX Concretos /     │  │  (CEMEX Agregados /     │  │  (Trading y Soluciones  │
│   Filiales Internas)    │  │   Trading Comercial)    │  │   Terceros Especial)    │
├─────────────────────────┤  ├─────────────────────────┤  ├─────────────────────────┤
│ • Modelo: Costo puro    │  │ • Modelo: Margen (MOP%) │  │ • Modelo: Intermediación│
│ • Margen comercial: 0%  │  │ • Clientes: Externos    │  │ • Logística de terceros │
│ • Clases Condición:     │  │ • Clases Condición:     │  │ • Proyectos especiales  │
│   - Material: ZMAH      │  │   - Material: ZMA6      │  │ • Suministro directo de │
│   - Flete: ZMPH         │  │   - Flete: ZMP1         │  │   canteras aliadas      │
│ • Nomenclatura común:   │  │ • Nomenclatura común:   │  │                         │
│   Centros tipo DW##     │  │   Centros tipo D###     │  │                         │
└─────────────────────────┘  └─────────────────────────┘  └─────────────────────────┘
```

### 2.1. Sociedad 7100 — CEMEX Concretos (Filiales Internas)
- **Propósito:** Transferencia entre plantas y canteras propias para producción de concreto.
- **Lógica de Precios:** No se busca generar utilidad entre divisiones de la misma empresa; el precio de venta replica el costo de compra/producción ($MOP = 0\%$).
- **Regla Operativa:** Las condiciones de 7100 registradas en centros 7180 son completamente válidas cuando una cantera interna surte proyectos del área comercial.

### 2.2. Sociedad 7180 — Agregados Comercial / Trading
- **Propósito:** Venta en el mercado abierto a constructoras y clientes externos.
- **Lógica de Precios:** Fijación de precios por mercado y zona geográfica, buscando capturar el máximo **Margen Operativo Puro (MOP %)** sobre la materia prima.

### 2.3. Sociedad 7277 — Trading Terceros
- **Propósito:** Compra de agregados en bancos de materiales de terceros (donde CEMEX no tiene cantera propia) para distribuirlos directamente a grandes clientes a través de fleteros certificados.

---

## 3. Cadena de Suministro y Red Logística (Supply Chain)

La distribución física de agregados depende de una relación de 4 nodos clave:

```
┌─────────────────┐       ┌─────────────────┐       ┌─────────────────┐       ┌─────────────────┐
│   SHIP FROM     │ ────> │  CENTRO / CEDIS │ ────> │ COND. EXPEDICIÓN│ ────> │ DESTINO FINAL   │
│ (Mina / Cantera)│       │(Nodo Logístico) │       │   (1 Recolección│       │(Obra / Cliente /│
│  ID: 90000XXXX  │       │   DW## / D###   │       │   / 2 Entrega)  │       │  Planta Concreto)
└─────────────────┘       └─────────────────┘       └─────────────────┘       └─────────────────┘
```

### 3.1. Nodos de la Red
1. **Ship From (Origen / Mina / Cantera):** Identificador numérico del acreedor o proveedor que opera el banco de extracción (ej. `900002370`, `900000064`). Distingue entre canteras propiedad de CEMEX y canteras de proveedores terceros.
2. **Centro / CEDIS (Centro de Distribución):** Instalación intermedia o planta receptora. Existen más de 100 centros a nivel nacional que articulan la logística regional.
3. **Destinatario / Destino:** Identificador de la ubicación exacta de descarga (obra, patio de acopio, planta satélite).

### 3.2. Modalidades de Entrega (Condición de Expedición)
- **Condición 1 (Entregado / Entrega Estándar):** CEMEX coordina el flete hasta el punto de entrega. El flete es obligatorio y forma parte del precio final al cliente.
- **Condición 2 o vacío sin flete (Recogido / FOB - Free On Board):** El cliente envía sus propios camiones a recoger el material a la cantera o CEDIS. En esta modalidad, **el flete pagado por CEMEX es $0.00** y no se exige flete en el semáforo.
- **Condición 4 (Entrega directa / CIF):** Entrega directa en obra gestionada por fleteros de CEMEX. El flete es obligatorio en auditoría y semáforo.

---

## 4. El Factor de Conversión Peso/Volumen (PV)

Los agregados presentan densidades variables según la geología de la cantera y la humedad:

$$\text{Factor PV (TN/m}^3) = \frac{\text{Cant. UMB (kg/m}^3)}{1000}$$

- **Unidad de Medida Base:** Los contratos de compra y las ventas a menudo se cotizan en **Toneladas (TN)** o en **Metros Cúbicos (M³)**.
- **Relevancia Financiera:** Si un cliente compra en $M^3$ pero a la cantera se le paga en $TN$, un error en el factor $PV$ en SAP puede pulverizar el margen de la ruta o hacer que CEMEX quede fuera de precio de mercado.

---

## 5. Arquitectura Financiera: Costos, Precios y Margen (MOP %)

### 5.1. Desglose del Precio de Venta
El precio total cobrado al cliente se compone de:
$$\text{Precio Total Venta} = \text{Precio Material (MP)} + \text{Tarifa de Flete}$$

### 5.2. Aislamiento del Margen Material (MOP Puro)
En el negocio de agregados, **el flete es un costo logístico transferible (*pass-through*)**, no una fuente de margen del producto pétreo. Mezclar el flete en el cálculo del margen distorsiona la rentabilidad real de la cantera.

El **MOP % (Margen Operativo Puro)** mide exclusivamente la rentabilidad de la roca extraída o comprada:

$$\text{Margen Material (\$)} = \text{Precio Venta MP (SAP VK13)} - \text{Costo Compra MP (TRAOPE)}$$

$$\text{MOP \%} = \frac{\text{Margen Material (\S)}}{\text{Precio Venta MP}} \times 100$$

### 5.3. Fórmulas de Validación de Negocio (Validación 1 y 2)
El sistema compara el costo esperado según la regla de negocio contra el contrato de compra real (`TRAOPE`):

```
                       ┌──────────────────────────────────────┐
                       │    ¿UM Venta y Costo son 'TN'?       │
                       └──────────────────┬───────────────────┘
                                          │
                        ┌─────────────────┴─────────────────┐
                     SÍ │                                   │ NO
                        ▼                                   ▼
          ┌───────────────────────────┐    ┌──────────────────────────────────┐
          │  Validación 1 =           │    │     ¿Condición Expedición = 1?   │
          │  Importe MP + Flete       │    └────────────────┬─────────────────┘
          └───────────────────────────┘                     │
                                             ┌──────────────┴──────────────┐
                                          SÍ │                             │ NO
                                             ▼                             ▼
                               ┌───────────────────────────┐  ┌───────────────────────────┐
                               │  Validación 1 =           │  │  Validación 1 =           │
                               │  PV × Importe MP          │  │  (Importe MP + Flete) × PV│
                               └───────────────────────────┘  └───────────────────────────┘
```

- **Validación 2 (Desviación Financiera):**
  $$\text{Validación 2} = \text{Importe Costo (TRAOPE)} - \text{Validación 1}$$
  - Si $\lvert \text{Validación 2} \rvert > \$1.00\text{ MXN}$ en operaciones de Filiales 7100, se enciende una alerta por inconsistencia de tarifas.

---

## 6. Matriz de Gobernanza y Niveles de Autorización

Para controlar la rentabilidad nacional, la fijación de precios en el negocio de Trading (7180/7277) se rige por umbrales estrictos de autorización comercial:

| Rango de Margen (MOP %) | Semáforo Visual | Nivel de Gobernanza / Aprobación | Acción Requerida |
| :--- | :---: | :--- | :--- |
| **MOP > 8.0%** | 🟢 **Verde** | **Champion Comercial / Dirección de Zona** | Operación saludable. Venta estándar autorizada. |
| **5.0% ≤ MOP ≤ 8.0%** | 🟡 **Amarillo** | **Gerencia Regional de Precios** | Margen ajustado. Requiere justificación de volumen o cliente estratégico. |
| **MOP < 5.0% o Negativo** | 🔴 **Rojo** | **Alerta Nacional / Dirección General** | Riesgo de pérdida operativa. Venta bloqueada hasta revisión de tarifas en SAP. |

---

## 7. Ecosistema de Datos y Tecnologías de Soporte

```
┌────────────────────────────────────────────────────────────────────────┐
│                          FUENTES TRANSACCIONALES                       │
│                                                                        │
│   ┌─────────────────────┐  ┌─────────────────────┐  ┌────────────────┐ │
│   │     SAP VK13        │  │     SAP TRAOPE      │  │  SAP ZSDD4501  │ │
│   │ Precios Venta MP y  │  │ Contratos Compra y  │  │ Contratos Venta│ │
│   │ Flete (ZMAH, ZMA6,  │  │ Tarifas de Mina /   │  │ y Condiciones  │ │
│   │  ZMPH, ZMP1)        │  │ Fleteros Terceros   │  │ Comerciales    │ │
│   └──────────┬──────────┘  └──────────┬──────────┘  └────────┬───────┘ │
└──────────────┼────────────────────────┼──────────────────────┼─────────┘
               │                        │                      │
               ▼                        ▼                      ▼
┌────────────────────────────────────────────────────────────────────────┐
│                         SNOWFLAKE DATA CLOUD                           │
│  Centralización analítica diaria de todas las rutas y CEDIS del país   │
│  Archivo: Extracción de Datos Agregados _ Formatos Para Alta de Ruta   │
└───────────────────────────────────┬────────────────────────────────────┘
                                    │
                                    ▼
┌────────────────────────────────────────────────────────────────────────┐
│                   MOTOR DE AUDITORÍA Y GOBERNANZA                      │
│                                                                        │
│  • Cruce multinivel por Llaves exactas (Concat1) y respaldo (Concat2)  │
│  • Cálculo instantáneo de MOP %, Desglose de Flete y Validaciones      │
│  • Generación de Dashboards Ejecutivos y Matrices multi-pestaña        │
└────────────────────────────────────────────────────────────────────────┘
```

### 7.1. Llaves Maestras de Cruce de Información
- **Concat 1 (Ruta Exacta Completa):** `ShipFrom - Centro - Destino - Material`
- **Concat 2 (Ruta General de Respaldo):** `ShipFrom - Centro - Material` (usada cuando el contrato de compra aplica a nivel cantera general sin restricción de destino).
- **Ruta A (Cruce con Contratos de Venta):** `Centro - ShipFrom - Destino - Material`.

---

## 8. Resumen de Términos y Glosario Clave

- **CEDIS:** Centro de Distribución o nodo receptor de agregados.
- **Ship From (SF):** Cantera, mina o banco de material de origen.
- **Destinatario:** Punto final de consumo o descarga de la mercancía.
- **ZMAH / ZMPH:** Clases de condición SAP para Material y Flete en Filiales Concretos (7100).
- **ZMA6 / ZMP1:** Clases de condición SAP para Material y Flete en Trading Agregados (7180).
- **TRAOPE:** Sistema y módulo de contratos marco de compras de materia prima.
- **ZSDD4501:** Reporte/Transacción de contratos comerciales de venta con clientes.
- **MOP (Material Operating Profit):** Margen de utilidad porcentual atribuible exclusivamente a la roca, descontando el flete.
- **PV:** Peso Volumétrico (relación tonelada/metro cúbico para conversión de densidad).
