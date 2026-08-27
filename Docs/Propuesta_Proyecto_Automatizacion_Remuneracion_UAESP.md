# Propuesta de Proyecto

## Automatización de la Remuneración Quincenal del Servicio de Aseo

**UAESP — Unidad Administrativa Especial de Servicios Públicos**

---

| Campo | Valor |
|---|---|
| **Proyecto** | Automatización Remuneración Quincenal — Servicio de Aseo |
| **Entidad** | UAESP |
| **Versión** | 1.0 |
| **Fecha** | Agosto 2026 |
| **Estado** | Propuesta |

---

## Contenido

1. [Resumen Ejecutivo](#1-resumen-ejecutivo)
2. [Contexto y Problema](#2-contexto-y-problema)
3. [Objetivo](#3-objetivo)
4. [Alcance — Fase 1](#4-alcance--fase-1)
5. [Insumos Requeridos](#5-insumos-requeridos)
6. [Arquitectura de la Aplicación](#6-arquitectura-de-la-aplicación)
7. [Diagrama de Bloques](#7-diagrama-de-bloques)
8. [Stack Tecnológico](#8-stack-tecnológico)
9. [Plan de Desarrollo](#9-plan-de-desarrollo)
10. [Criterios de Aceptación](#10-criterios-de-aceptación)
11. [Riesgos y Mitigaciones](#11-riesgos-y-mitigaciones)
12. [Próximos Pasos](#12-próximos-pasos)

---

## 1. Resumen Ejecutivo

Se propone el desarrollo de una aplicación de escritorio que automatice el proceso de remuneración quincenal del servicio de aseo en Bogotá, reemplazando el proceso manual actual basado en instrucciones para IA y diligenciamiento manual de plantillas Excel.

La aplicación will leer los archivos fuente de recaudo por concesionario (ASE), aplicar las reglas de cálculo definidas en la Resolución UAESP 27 de 2018, y generar el archivo consolidado de remuneración con validaciones de coherencia entre hojas.

**Fase 1:** Prototipo con un solo ASE para validar los cálculos y la arquitectura.

---

## 2. Contexto y Problema

### 2.1 Situación Actual

El proceso de remuneración quincenal se ejecuta dos veces al mes y involucra:

- **5 operadores de aseo (ASE):** Promoambiental, LIME, Ciudad Limpia, Bogotá Limpia, Área Limpia.
- **5+ empresas de facturación conjunta:** EAAB Reciprocidad, EAAB + Ciudad Limpia, ENEL, ENERBIT, Occidente (facturación directa).
- **Múltiples reportes fuente** por ASE, generados en formato Excel, con estructura de nombres basada en rango de fechas y timestamp.
- **15+ hojas de cálculo** en el archivo de plantilla que deben ser diligenciadas con valores extraídos de los reportes fuente.
- **Fórmulas existentes** en la plantilla que deben ser preservadas al reemplazar valores.

### 2.2 Problema

| Problema | Impacto |
|---|---|
| Proceso manual propenso a errores | Cálculos incorrectos que afectan pagos a operadores |
| Dependencia de IA para interpretar instrucciones | Resultados inconsistentes entre ejecuciones |
| Alta complejidad operativa | 5 ASE × 15+ hojas = decenas de operaciones manuales por quincena |
| Sin trazabilidad | Difícil auditar qué se hizo en cada ejecución |
| Repetitividad | Mismo proceso cada 15 días sin variación significativa |

### 2.3 Marco Normativo

La **Resolución UAESP 27 de 2018** (Reglamento Comercial y Financiero) define las reglas de distribución del recaudo entre operadores, los componentes tarifarios y los procedimientos de remuneración. La aplicación debe respetar estas reglas en su totalidad.

---

## 3. Objetivo

### 3.1 Objetivo General

Desarrollar una aplicación de escritorio que automatice la elaboración de la remuneración quincenal del servicio de aseo, distribuyendo el recaudo entre los concesionales y generando el archivo consolidado de salida.

### 3.2 Objetivos Específicos

1. Leer y parsear los archivos fuente de recaudo (R1, R2, R4) por ASE.
2. Aplicar las reglas de cálculo del CONSOLIDADO_TOTAL RECAUDO (TOT_OPT, R2 Total Oportuno, EXTEMP, Reversión R4, AJUSTES).
3. Escribir los valores calculados en la plantilla de remuneración preservando las fórmulas existentes.
4. Ejecutar validaciones cruzadas entre hojas para garantizar coherencia.
5. Generar el archivo de salida con la nomenclatura estándar.
6. Proporcionar trazabilidad completa de cada ejecución.

---

## 4. Alcance — Fase 1

### 4.1 Alcance del Prototipo

| Concepto | Alcance |
|---|---|
| **ASE a procesar** | **Un (1) ASE** Selection para prueba de concepto |
| **Archivos fuente obligatorios** | 3 por ASE: Recaudoporcomponente (R1), RerpoteDetalleSaldosaFavor (R2), ReversiónPorComponente (R4) |
| **Cálculos del CONSOLIDADO** | TOT_OPT (D9:D13), R2 Total Oportuno (D28:D32), EXTEMP (D47:D51), Reversión R4 (D66:D70), Total por ASE (D104:D108) |
| **Hojas a diligenciar** | CONSOLIDADO_TOTAL RECAUDO + hojas de detalle (R1, R2, R4) |
| **Validaciones** | Básicas: coherencia entre hojas, fórmulas preservadas |
| **UI** | Windows Forms — selección de carpetas, progreso, resultados |

### 4.2 Fuera del Alcance — Fase 1

- Los otros 4 ASE (se procesarán en Fase 2).
- Hojas: SALDOS POR NOTA, RETRIBUCION NEGATIVA, AJUSTES-SF-T (solo 2.ª quincena).
- Hojas: BCE SC POR FACT, INTERVENTORIA, REPORTE RECAUDO x BANCO.
- Validaciones avanzadas (VALIDACION_TOTAL, VALIDACION_RECIP).
- Generación automática de la hoja DetRetri.

### 4.3 Justificación del Alcance

Focusing on a single ASE first allows the team to:

- Validate that the application's calculations match known results from manual processing.
- Refine the architecture and Excel handling before scaling to all 5 ASE.
- Identify edge cases (e.g., ASE4 que no tiene columna "Especiales") in a controlled manner.
- Deliver a working proof of concept quickly.

---

## 5. Insumos Requeridos

> **NOTA IMPORTANTE:** Para construir la aplicación sin ambigüedades y con datos reales, se requieren los siguientes insumos. Sin ellos, los cálculos y posiciones de celdas serían suposiciones.

### 5.1 Archivos Fuente por ASE

Cada ASE tiene una carpeta numerada (1 a 5) dentro de la carpeta del período. Cada carpeta contiene como mínimo 3 archivos fuente obligatorios:

| # | Reporte | Nombre Exacto del Archivo | Obligatorio | Periodicidad | Qué Extrae |
|---|---|---|---|---|---|
| R1 | Recaudo por componente | `Recaudoporcomponente_to_date{INI}ddMMyyyy_to_date{FIN}ddMMyyyy___{timestamp}.xlsx` | **SÍ** | Ambas quincenas | TOT_OPT (fila col1="Componente", col2="Total", col. F), EXTEMP (fila col2="Mes" en sección Extemporáneo, col. F) |
| R2 | Detalle saldos a favor | `RerpoteDetalleSaldosaFavor_to_date{INI}ddMMyyyy_to_date{FIN}ddMMyyyy___{timestamp}.xlsx` | **SÍ** | Ambas quincenas | Grand Total (col. E última fila), SERV_ESP_K (col. 11 según encabezado) |
| R4 | Reversión por componente | `ReversiónPorComponente_to_date{INI}ddMMyyyy_to_date{FIN}ddMMyyyy___{timestamp}.xlsx` | **SÍ** | Ambas quincenas | Total reversiones (última fila, col. 4, valor negativo) |
| — | Pago por banco | `ReportePagosxBanco_to_date{INI}ddMMyyyy_to_date{FIN}ddMMyyyy___{timestamp}.xlsx` | Opcional | Ambas quincenas | Recaudo por fuente/banco |
| — | Balance SC | `R4-BalanceSubsidioyContribuciones_to_date{INI}ddMMyyyy_to_date{FIN}ddMMyyyy___{timestamp}.xlsx` | Opcional | Ambas quincenas | Subsidio (col. D) y contribución (col. E) |
| — | Saldos por nota | `SaldosaFavorAplicadosPorNotas_to_date{INI}ddMMyyyy_to_date{FIN}ddMMyyyy___{timestamp}.xlsx` | **Solo 2.ª quincena** | 2.ª quincena | AJUSTES-SF-T |
| — | Retribución negativa | `RetribuciónNegativa_to_date{INI}ddMMyyyy_to_date{FIN}ddMMyyyy___{timestamp}.xlsx` | **Solo 2.ª quincena** | 2.ª quincena | AJUSTES-SF-T |

**Donde:**
- `{INI}` = Fecha inicio de la quincena en formato `ddMMyyyy` (ej. `01052026`)
- `{FIN}` = Fecha fin de la quincena en formato `ddMMyyyy` (ej. `15052026`)
- `{timestamp}` = Número aleatorio generado por el sistema fuente (ej. `202651916254664`)

### 5.2 Estructura de Carpetas por Período

```
REMUNERACION AAAAMM#/
├── 1-Promoambiental/
│   ├── Recaudoporcomponente_to_date01052026ddMMyyyy_to_date15052026ddMMyyyy___202651916254664.xlsx
│   ├── RerpoteDetalleSaldosaFavor_to_date01052026ddMMyyyy_to_date15052026ddMMyyyy___20265191641033.xlsx
│   └── ReversiónPorComponente_to_date01052026ddMMyyyy_to_date15052026ddMMyyyy___202651916425722.xlsx
├── 2-Lime/
│   └── (misma estructura)
├── 3-Ciudad Limpia/
│   └── (misma estructura)
├── 4-Bogotá Limpia/
│   └── (misma estructura)
└── 5-Área Limpia/
    └── (misma estructura)
```

### 5.3 Archivo de Plantilla

| Archivo | Nombre | Descripción |
|---|---|---|
| Plantilla | `Remuneración AAAAMM-# Total.xlsx` | Archivo Excel con 15+ hojas, fórmulas predefinidas, estructura fija. Es el archivo base que se diligencia con los valores calculados. |

### 5.4 Insumos Críticos para el Desarrollo

| # | Insumo | Propósito | Urgencia |
|---|---|---|---|
| 1 | **Archivo de plantilla real** ya utilizado | Verificar posiciones exactas de celdas, encabezados, fórmulas | **CRÍTICO** |
| 2 | **1 juego completo de archivos fuente de un período** (mínimo R1, R2, R4 de un ASE) | Validar cálculos contra resultado conocido | **CRÍTICO** |
| 3 | **El resultado final de ese mismo período** (archivo consolidado ya diligenciado) | Tener el "golden test" — comparar salida vs. resultado real | **CRÍTICO** |
| 4 | **Nombre del ASE Selection para Fase 1** | Definir qué concesionario se procesa primero | **ALTO** |

---

## 6. Arquitectura de la Aplicación

### 6.1 Principios de Diseño

| Principio | Descripción |
|---|---|
| **Separación de capas** | Lógica de negocio (Core), acceso a datos (Infrastructure), interfaz (WinForms) |
| **Migrabilidad** | La lógica de negocio no depende de WinForms; puede migrar a WPF, Blazor, API Web |
| **Preservación de fórmulas** | Nunca se alteran las fórmulas existentes de la plantilla |
| **Lectura directa de fuente** | Siempre se leen valores calculados de los archivos fuente, nunca de fórmulas almacenadas |
| **Trazabilidad** | Cada paso se registra para auditoría |

### 6.2 Estructura de Capas

```
┌─────────────────────────────────────────────────┐
│            CAPA DE PRESENTACIÓN                  │
│         Remuneracion.WinForms (.NET 10)          │
│                                                  │
│  • Formulario de selección de período            │
│  • Selector de carpetas de fuentes               │
│  • Selector de archivo plantilla                 │
│  • Barra de progreso                             │
│  • Panel de resultados y validaciones            │
│  • Log de ejecución                              │
└──────────────────────┬──────────────────────────┘
                       │ depende de
                       ▼
┌─────────────────────────────────────────────────┐
│            CAPA DE INFRAESTRUCTURA               │
│     Remuneracion.Infrastructure (.NET 10)        │
│                                                  │
│  ┌────────────────────┐ ┌────────────────────┐  │
│  │  ExcelDataReader   │ │   OpenXML SDK      │  │
│  │  (Lectura)         │ │   (Escritura)      │  │
│  │  Licencia: MIT     │ │   Licencia: MIT    │  │
│  └────────────────────┘ └────────────────────┘  │
│                                                  │
│  • RecaudoReader: lee archivos fuente R1/R2/R4  │
│  • PlantillaWriter: escribe en plantilla        │
│  • ArchivoFuenteLocator: resuelve rutas         │
└──────────────────────┬──────────────────────────┘
                       │ usa
                       ▼
┌─────────────────────────────────────────────────┐
│            CAPA DE NEGOCIO (CORE)                │
│         Remuneracion.Core (.NET 10)              │
│                                                  │
│  ┌─────────────────────────────────────────┐    │
│  │  MODELOS DE DATOS                       │    │
│  │  • Periodo (AAAAMM, quincena)           │    │
│  │  • Ase (id, nombre, carpeta)            │    │
│  │  • RecaudoComponente (R1)               │    │
│  │  • SaldosFavor (R2)                     │    │
│  │  • ReversionPago (R4)                   │    │
│  │  • ConsolidadoAse (valores calculados)  │    │
│  │  • ResultadoRemuneracion                │    │
│  └─────────────────────────────────────────┘    │
│                                                  │
│  ┌─────────────────────────────────────────┐    │
│  │  SERVICIOS DE CÁLCULO                   │    │
│  │  • CalculoTotOpt(): TOT_OPT desde R1   │    │
│  │  • CalculoR2Total(): R2 desde R2        │    │
│  │  • CalculoExtemp(): EXTEMP desde R1     │    │
│  │  • CalculoReversion(): R4 desde R4      │    │
│  │  • CalculoTotalAse(): suma components   │    │
│  │  • CalculoGranTotal(): total período    │    │
│  └─────────────────────────────────────────┘    │
│                                                  │
│  ┌─────────────────────────────────────────┐    │
│  │  SERVICIOS DE VALIDACIÓN                │    │
│  │  • ValidarCoherenciaHojas()             │    │
│  │  • ValidarFormulas()                    │    │
│  │  • ValidarDiferencias() (≤ ±0.5)       │    │
│  └─────────────────────────────────────────┘    │
└─────────────────────────────────────────────────┘
```

---

## 7. Diagrama de Bloques

### 7.1 Flujo General de la Aplicación

```
┌─────────────────────────────────────────────────────────────────────────┐
│                         REMUNERACION UAESP                              │
│                   .NET 10 — Windows Forms (GUI)                         │
└─────────────────────────────────┬───────────────────────────────────────┘
                                  │
                                  ▼
┌─────────────────────────────────────────────────────────────────────────┐
│                      ENTRADA DE DATOS                                   │
│                                                                         │
│  ┌──────────────┐    ┌──────────────┐    ┌──────────────┐             │
│  │  Seleccionar │    │  Seleccionar │    │  Seleccionar │             │
│  │  Período     │    │  Carpeta de  │    │  Archivo de  │             │
│  │  (AAAAMM-#)  │    │  Fuentes     │    │  Plantilla   │             │
│  │              │    │  (5 carpetas │    │  Excel       │             │
│  │              │    │   ASE)       │    │              │             │
│  └──────┬───────┘    └──────┬───────┘    └──────┬───────┘             │
│         └───────────────────┼───────────────────┘                     │
└─────────────────────────────┼─────────────────────────────────────────┘
                              │
                              ▼
┌─────────────────────────────────────────────────────────────────────────┐
│                      LECTURA DE FUENTES                                 │
│                   (ExcelDataReader — MIT)                               │
│                                                                         │
│  Para cada ASE Selection (Fase 1: un solo ASE):                        │
│                                                                         │
│  ┌─────────────────────┐                                               │
│  │  Recaudoporcomponente│  →  Buscar fila col1="Componente"            │
│  │  (R1)                │      col2="Total", extraer col. F → TOT_OPT │
│  │                      │  →  Buscar fila col2="Mes" en Extemporáneo  │
│  │                      │      extraer col. F → EXTEMP                │
│  └──────────┬──────────┘                                               │
│             │                                                          │
│  ┌──────────┴──────────┐                                               │
│  │  RerpoteDetalle     │  →  Última fila: Grand Total (col. E)        │
│  │  SaldosFavor (R2)   │  →  Encabezado col. 11: ¿"Especiales"?      │
│  │                     │      SÍ → SERV_ESP_K = valor col. 11         │
│  │                     │      NO → SERV_ESP_K = 0                     │
│  │                     │  →  R2 Total Oportuno = E - SERV_ESP_K       │
│  └──────────┬──────────┘                                               │
│             │                                                          │
│  ┌──────────┴──────────┐                                               │
│  │  ReversiónPor       │  →  Última fila, columna 4                   │
│  │  Componente (R4)    │      (valor negativo) → R4 Reversión         │
│  └─────────────────────┘                                               │
└─────────────────────────────┼─────────────────────────────────────────┘
                              │
                              ▼
┌─────────────────────────────────────────────────────────────────────────┐
│                    MOTOR DE CÁLCULO                                     │
│                                                                         │
│  CONSOLIDADO_TOTAL RECAUDO — Celdas D por ASE:                        │
│                                                                         │
│  ┌───────────────────────────────────────────────────────────────────┐ │
│  │  D9  = TOT_OPT (ASE1)     ← de Recaudoporcomponente, col. F    │ │
│  │  D10 = TOT_OPT (ASE2)     ← (solo 1 ASE en Fase 1)            │ │
│  │  D11 = TOT_OPT (ASE3)                                              │ │
│  │  D12 = TOT_OPT (ASE4)                                              │ │
│  │  D13 = TOT_OPT (ASE5)                                              │ │
│  └───────────────────────────────────────────────────────────────────┘ │
│  ┌───────────────────────────────────────────────────────────────────┐ │
│  │  D28 = R2 Total Oportuno (ASE1) ← Grand Total (E) - SERV_ESP_K │ │
│  │  D29-D32 = ... (resto de ASE)                                   │ │
│  └───────────────────────────────────────────────────────────────────┘ │
│  ┌───────────────────────────────────────────────────────────────────┐ │
│  │  D47 = EXTEMP (ASE1)        ← Recaudoporcomponente, Extemporáneo│ │
│  │  D48-D51 = ... (resto de ASE)                                   │ │
│  └───────────────────────────────────────────────────────────────────┘ │
│  ┌───────────────────────────────────────────────────────────────────┐ │
│  │  D66 = Reversión R4 (ASE1)  ← ReversiónPorComponente, últ. fila│ │
│  │  D67-D70 = ... (resto de ASE)                                   │ │
│  └───────────────────────────────────────────────────────────────────┘ │
│  ┌───────────────────────────────────────────────────────────────────┐ │
│  │  D85 = AJUSTES-SF-T (ASE1)  ← 0 (Fase 1, sin 2.ª quincena)    │ │
│  │  D86-D89 = ...                                                 │ │
│  └───────────────────────────────────────────────────────────────────┘ │
│  ┌───────────────────────────────────────────────────────────────────┐ │
│  │  D104 = D9 + D28 + D47 + D66 + D85  → Total ASE1               │ │
│  │  D105-D108 = ... (resto de ASE)                                 │ │
│  │  D109 = SUMA(D104:D108)          → Gran Total del período       │ │
│  └───────────────────────────────────────────────────────────────────┘ │
└─────────────────────────────┼─────────────────────────────────────────┘
                              │
                              ▼
┌─────────────────────────────────────────────────────────────────────────┐
│                   ESCRITURA EN PLANTILLA                                │
│                (OpenXML SDK — MIT)                                      │
│                                                                         │
│  ┌─────────────────────────────────────────────────────────────────┐   │
│  │  HOJAS DE DETALLE (por ASE)                                    │   │
│  │                                                                 │   │
│  │  • REPORTE COMPONENTES R1 → pega valores R1                    │   │
│  │  • REM. ANTICIPOS R2      → pega valores R2                    │   │
│  │  • REVERSION PAGOS R4     → pega valores R4                    │   │
│  │  • Recaudo EAAB/ENEL/etc → pega conciliación por empresa      │   │
│  └─────────────────────────────────────────────────────────────────┘   │
│  ┌─────────────────────────────────────────────────────────────────┐   │
│  │  HOJA CONSOLIDADO_TOTAL RECAUDO                                │   │
│  │                                                                 │   │
│  │  • Escribe D9:D13, D28:D32, D47:D51, D66:D70, D85:D89        │   │
│  │  • Escribe D104:D108 (Total por ASE)                           │   │
│  │  • Escribe D109 (Gran Total)                                   │   │
│  │  • FÓRMULAS NO SE ALTERAN                                      │   │
│  └─────────────────────────────────────────────────────────────────┘   │
│  ┌─────────────────────────────────────────────────────────────────┐   │
│  │  HOJA DetRetri[período]                                        │   │
│  │  • Columna D con valores enteros redondeados                   │   │
│  └─────────────────────────────────────────────────────────────────┘   │
└─────────────────────────────┼─────────────────────────────────────────┘
                              │
                              ▼
┌─────────────────────────────────────────────────────────────────────────┐
│                     VALIDACIÓN                                          │
│                                                                         │
│  ┌─────────────────────────────────────────────────────────────────┐   │
│  │  ✅  Diferencia CONSOLIDADO vs fuentes ≤ ±0.5 (redondeo)      │   │
│  │  ✅  D109 = SUMA(D104:D108)                                    │   │
│  │  ✅  Fórmulas de la plantilla preservadas                      │   │
│  │  ✅  Columnas correctas según encabezados                      │   │
│  │  ✅  DetRetri con valores enteros                              │   │
│  │  ✅  Total Oportuno en Reporte R1 = CONSOLIDADO D9:D13         │   │
│  └─────────────────────────────────────────────────────────────────┘   │
└─────────────────────────────┼─────────────────────────────────────────┘
                              │
                              ▼
┌─────────────────────────────────────────────────────────────────────────┐
│                          SALIDA                                         │
│                                                                         │
│  ┌─────────────────────────────────────────────────────────────────┐   │
│  │  📄 Archivo: "Remuneración AAAAMM-# Total.xlsx"                │   │
│  │                                                                 │   │
│  │  📊 Panel de resultados en pantalla                            │   │
│  │     • Totales por ASE                                         │   │
│  │     • Diferencias detectadas (si las hay)                     │   │
│  │     • Estado de validaciones                                  │   │
│  │                                                                 │   │
│  │  📋 Log de ejecución detallado                                │   │
│  │     • Timestamp de cada operación                             │   │
│  │     • Archivos leídos                                         │   │
│  │     • Valores extraídos y escritos                            │   │
│  │     • Warnings e informativos                                 │   │
│  └─────────────────────────────────────────────────────────────────┘   │
└─────────────────────────────────────────────────────────────────────────┘
```

---

## 8. Stack Tecnológico

| Componente | Tecnología | Versión | Licencia | Propósito |
|---|---|---|---|---|
| **Framework** | .NET | 10 (LTS) | MIT | Runtime de la aplicación |
| **UI** | Windows Forms | — | — | Interfaz de escritorio (prototipo migrable) |
| **Lectura Excel** | ExcelDataReader | 3.x+ | MIT | Lectura rápida de archivos fuente |
| **Escritura Excel** | OpenXML SDK | 2.5+ | MIT | Escritura en plantilla preservando fórmulas |
| **Logging** | Serilog | 3.x+ | Apache 2.0 | Trazabilidad de ejecución |
| **Testing** | xUnit + FluentAssertions | — | Apache 2.0 | Validación de cálculos |

### 8.1 Justificación de OpenXML SDK

Se selecciona **DocumentFormat.OpenXml** (Microsoft OpenXML SDK) por las siguientes razones:

1. **Licencia MIT** — gratuita para uso comercial sin restricciones.
2. **Preserva fórmulas** — permite escribir valores en celdas sin alterar las fórmulas existentes.
3. **Soporte oficial** — mantenida por Microsoft, compatibilidad garantizada con .NET 10.
4. **Control total** — acceso a nivel de nodo XML para operaciones complejas.
5. **Sin dependencias externas** — no requiere Microsoft Office instalado.

### 8.2 Justificación de ExcelDataReader

Se selecciona **ExcelDataReader** para lectura de archivos fuente porque:

1. **Licencia MIT** — gratuita para uso comercial.
2. **Alto rendimiento** — optimizado para lectura masiva de datos.
3. **Lectura de valores** — extrae valores numéricos directamente (no necesita fórmulas de los fuente).
4. **API simple** — reduce código de parsing.

---

## 9. Plan de Desarrollo

### Fase 1 — Prototipo (1 ASE)

| Iteración | Entregable | Descripción |
|---|---|---|
| 1.1 | Estructura del proyecto | Crear solución .NET 10 con 3 capas (Core, Infrastructure, WinForms) |
| 1.2 | Modelos de datos | Definir entidades: Periodo, Ase, RecaudoComponente, SaldosFavor, ReversionPago, ConsolidadoAse |
| 1.3 | Lector de fuentes R1 | Leer Recaudoporcomponente, extraer TOT_OPT y EXTEMP |
| 1.4 | Lector de fuentes R2 | Leer RerpoteDetalleSaldosaFavor, extraer Grand Total y SERV_ESP_K |
| 1.5 | Lector de fuentes R4 | Leer ReversiónPorComponente, extraer total reversiones |
| 1.6 | Motor de cálculo | Implementar cálculo del CONSOLIDADO para 1 ASE |
| 1.7 | Escritor de plantilla | Escribir valores en plantilla con OpenXML SDK preservando fórmulas |
| 1.8 | Validaciones básicas | Coherencia entre hojas, fórmulas preservadas |
| 1.9 | UI WinForms | Formulario de selección, progreso, resultados, log |
| 1.10 | Testing | Validar contra resultado real conocido (golden test) |

### Fase 2 — Escalamiento (5 ASE)

| Iteración | Entregable | Descripción |
|---|---|---|
| 2.1 | Procesamiento multi-ASE | Iterar las 5 carpetas ASE |
| 2.2 | Hojas de conciliación | Diligenciar hojas por empresa de facturación (EAAB, ENEL, etc.) |
| 2.3 | REPORTE RECAUDO x BANCO | Hoja de consolidado por banco |
| 2.4 | BCE SC POR FACT | Balance de subsidios y contribuciones |
| 2.5 | AJUSTES-SF-T | SALDOS POR NOTA + RETRIBUCION NEGATIVA (2.ª quincena) |
| 2.6 | DetRetri | Hoja de detalle con valores enteros redondeados |
| 2.7 | Validaciones cruzadas | VALIDACION_TOTAL, VALIDACION_RECIP, etc. |

### Fase 3 — Producción

| Iteración | Entregable | Descripción |
|---|---|---|
| 3.1 | Manejo de errores | Errores de archivo faltante, formato incorrecto, etc. |
| 3.2 | Log completo | Auditoría detallada de cada ejecución |
| 3.3 | Modo automático | Ejecución sin intervención manual (CLI) |
| 3.4 | Documentación | Manual de usuario y technical |

---

## 10. Criterios de Aceptación

### 10.1 Fase 1 (Prototipo — 1 ASE)

| # | Criterio | Verificación |
|---|---|---|
| CA-1 | La aplicación lee correctamente los 3 archivos fuente (R1, R2, R4) de un ASE | Comparación visual de valores extraídos vs. archivo original |
| CA-2 | Los valores TOT_OPT, R2 Total Oportuno, EXTEMP y Reversión R4 coinciden con el cálculo manual | Prueba con datos conocidos |
| CA-3 | El archivo de salida tiene los valores correctos en CONSOLIDADO D9:D13, D28:D32, D47:D51, D66:D70, D104:D109 | Comparación contra resultado real |
| CA-4 | Las fórmulas de la plantilla NO fueron alteradas | Verificación de celdas con fórmulas |
| CA-5 | Las validaciones básicas pasan sin errores | Ejecución exitosa |
| CA-6 | El log registra cada paso de la ejecución | Revisión del log |
| CA-7 | La UI permite seleccionar período, carpeta y plantilla | Prueba funcional |

### 10.2 Criterio de Éxito General

> La aplicación se considera exitosa cuando los valores del CONSOLIDADO_TOTAL RECAUDO generados automáticamente **coinciden exactamente** (o con diferencia ≤ ±0.5 por redondeo) con los valores del resultado manual conocido para el mismo período.

---

## 11. Riesgos y Mitigaciones

| # | Riesgo | Probabilidad | Impacto | Mitigación |
|---|---|---|---|---|
| 1 | **No se dispone del archivo de plantilla real** | Alta | Crítico | Solicitar a UAESP un archivo de plantilla ya utilizado |
| 2 | **No se dispone de archivos fuente de ejemplo** | Alta | Crítico | Solicitar un período completo ya procesado |
| 3 | **Posiciones de celdas varían entre plantillas** | Media | Alto | Diseñar la app para detectar encabezados dinámicamente, noHardcodear posiciones fijas |
| 4 | **La plantilla tiene fórmulas que dependen de filas dinámicas** | Media | Alto | Usar OpenXML SDK que preserva fórmulas; pegar solo valores |
| 5 | **Errores de parsing en archivos fuente** | Media | Medio | Implementar validación de estructura antes de procesar |
| 6 | **Cambio en formato de reportes fuente** | Baja | Medio | Diseñar parsers configurables por archivo |
| 7 | **OpenXML SDK con curva de aprendizaje alta** | Media | Bajo | Crear wrappers que simplifiquen las operaciones comunes |

---

## 12. Próximos Pasos

| # | Acción | Responsable | Estado |
|---|---|---|---|
| 1 | **Aprobar propuesta** | UAESP / Equipo de desarrollo | Pendiente |
| 2 | **Entregar archivo de plantilla real** | UAESP | **Requerido** |
| 3 | **Entregar 1 período completo de ejemplo** (5 carpetas ASE + resultado final) | UAESP | **Requerido** |
| 4 | **Seleccionar ASE para Fase 1** | UAESP | Pendiente |
| 5 | **Configurar entorno de desarrollo** (.NET 10, OpenXML SDK, ExcelDataReader) | Equipo de desarrollo | Pendiente |
| 6 | **Iniciar Fase 1** | Equipo de desarrollo | Pendiente |

---

## Apéndice A — Glosario

| Término | Definición |
|---|---|
| **ASE** | Administrador del Servicio de Aseo (concesionario/operador) |
| **UAESP** | Unidad Administrativa Especial de Servicios Públicos |
| **EFC** | Empresa de Facturación Conjunta |
| **Remuneración** | Mecanismo de distribución del recaudo entre operadores |
| **Quincena** | Período de 15 días (1-15 o 16-31 del mes) |
| **Componente tarifario** | Subcategoría del servicio de aseo (subbolsa) |
| **Recaudo** | Pagos recibidos por el servicio de aseo |
| **Recaudo oportuno** | Pago dentro del período esperado |
| **Recaudo extemporáneo** | Pago fuera del período esperado |
| **Saldos a favor** | Recaudo remanente a favor del usuario |
| **Reversión** | Anulación de un pago registrado |
| **Consolidado** | Archivo final que integra todos los reportes de remuneración |

---

## Apéndice B — Operadores por Fuente de Recaudo

| ASE | ENEL | EAAB Recip. | EAAB+CiudLimp | ENERBIT | Occidente |
|---|---|---|---|---|---|
| ASE1 Promoambiental | X | — | — | — | X |
| ASE2 LIME | X | X | — | — | — |
| ASE3 Ciudad Limpia | X | — | — | — | X |
| ASE4 Bogotá Limpia | X | — | X | — | — |
| ASE5 Área Limpia | X | — | — | X | X |
