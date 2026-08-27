# Proyecto de Automatización — Remuneración Quincenal del Servicio de Aseo (UAESP)

> **Norma de referencia:** Resolución UAESP 27 de 2018 — Reglamento Comercial y Financiero

---

## 1. Descripción General

Este proyecto automatiza la elaboración de la **remuneración quincenal** del servicio de aseo en Bogotá, distribuyendo el recaudo recibido y aplicado entre los cinco operadores (ASE) y asignándolo a cada componente tarifario. El resultado es un archivo Excel consolidado que integra todos los reportes de recaudo, anticipos, reversiones y validaciones.

---

## 2. Estructura del Repositorio

```
Automatización/
├── Docs/
│   ├── Detalle de plantilla.docx          # Anexo operativo: instrucciones por hoja
│   ├── Proceso de Recaudo.docx            # Documentación del proceso de recepción de pagos
│   ├── Prompt_Maestro_Proyecto_Remuneracion_UAESP _ Vo .docx
│   │                                      # Prompt maestro ejecutable para IA
│   └── Propuesta_Proyecto_Automatizacion_Remuneracion_UAESP.md
│                                          # Propuesta formal del proyecto de automatización
└── README.md
```

---

## 3. Documentos del Proyecto

### 3.1 Detalle de Plantilla

**¿Qué es?** Anexo de referencia con el funcionamiento de cada hoja del archivo de plantilla Excel.

**¿Para qué sirve?** Explica en detalle qué datos contiene cada hoja, de qué fuente provienen, qué fórmulas tiene y cómo se alimentan entre sí. Es el manual técnico para quien ejecuta la remuneración.

**Puntos principales:**

| Hoja | Función |
|---|---|
| `CONSOLIDADO_TOTAL RECAUDO` | Resumen maestro: distribución de recaudo por ASE y componente. Reporte final. |
| `Reporte Componentes R1` | Recaudo oportuno por componente tarifario (fuente: *Recaudoporcomponente*) |
| `Rem. Anticipos R2` | Anticipos / saldos a favor aplicados (fuente: *RerpoteDetalleSaldosaFavor*) |
| `Reversion Pagos R4` | Pagos revertidos o anulados (fuente: *ReversiónPorComponente*) |
| `REPORTE RECAUDO x BANCO` | Consolidado de recaudo aplicado por banco y canal |
| `ANT EXT-REV` | Recaudo anulado o reversado durante la quincena |
| `ANTICIPOS USUARIOS` | Anticipos generados y aplicados a facturación (formulado automáticamente) |
| `BCE SC POR FACT` | Balance de subsidios y contribuciones por facturación |
| `SALDOS POR NOTA` | Saldos a favor aplicados por nota (solo 2.ª quincena) |
| `RETRIBUCION NEGATIVA` | Reintegros a usuarios como valores negativos (solo 2.ª quincena) |
| `AJUSTES - SF-T` | Combina SALDOS POR NOTA + RETRIBUCION NEGATIVA → alimenta R5 |
| `INTERVENTORIA` | Costo de interventoría por ASE (rango L25:N31) |
| `Recaudo EAAB/ENEL/ENERBIT/etc.` | Conciliación de recaudo por empresa de facturación conjunta |

**Regla crítica:** Toda información se pega **en valores**, nunca alterando las fórmulas existentes.

---

### 3.2 Proceso de Recaudo

**¿Qué es?** Documentación del flujo de recepción de pagos del servicio de aseo, desde las empresas de facturación hasta el cargue en el sistema comercial.

**¿Para qué sirve?** Entender de dónde vienen los datos, cómo se estructuran las carpetas de archivos fuente y cómo se procesan antes de llegar a la plantilla de remuneración.

**Puntos principales:**

- **Empresas de facturación conjunta (EFC):** ENEL, EAAB, ENERBIT — remiten archivos diarios en formato TXT identificados por banco y canal.
- **Facturación directa:** Banco de Occidente — recaudo por canales presenciales (O23), corresponsales no bancarios (H23) y PSE.
- **Estructura de carpetas:** Cada EFC genera subcarpetas por banco/canal (ej. `OCC. BLTO`, `OCC. MED-`).
- **Proceso posterior:** Los archivos se convierten a formato Asobancaria 98, se cargan al sistema comercial y se aplican a cada cuenta contrato.
- **Reversiones:** Al finalizar la quincena se genera un archivo plano con todos los pagos anulados/revertidos.
- **Validación:** Se generan archivos planos acumulados y diarios para cruzar el recaudo cargado vs. el recibido.

---

### 3.3 Prompt Maestro — Proyecto de Remuneración

**¿Qué es?** Documento ejecutable e integral que contiene el prompt listo para entregar a una IA, más el manual completo de instrucciones para elaborar la remuneración.

**¿Para qué sirve?** Es el heart del proyecto: define el rol, los cálculos, las reglas de negocio, el proceso secuencial y las validaciones. Con él, una IA puede ejecutar la remuneración de forma autónoma y verificable.

**Puntos principales:**

- **Rol:** Especialista en remuneración quincenal del servicio de aseo de la UAESP.
- **Objetivo:** Distribuir el recaudo entre 5 ASE, asignarlo a componentes tarifarios, producir un archivo consolidado.
- **Alcance:** Instrucciones funcionales, operativas y de cálculo en un único procedimiento.

#### 3.3.1 Los 5 Operadores (ASE)

| ASE | Operador |
|---|---|
| ASE1 | Promoambiental |
| ASE2 | LIME (Limpieza Metropolitana) |
| ASE3 | Ciudad Limpia |
| ASE4 | Bogotá Limpia |
| ASE5 | Área Limpia |

#### 3.3.2 Estructura de Carpetas por Período

```
REMUNERACION AAAAMM#/
├── 1-Promoambiental/
│   ├── Recaudoporcomponente_*.xlsx
│   ├── RerpoteDetalleSaldosaFavor_*.xlsx
│   └── ReversiónPorComponente_*.xlsx
├── 2-Lime/
├── 3-Ciudad Limpia/
├── 4-Bogotá Limpia/
└── 5-Área Limpia/
```

#### 3.3.3 Fórmulas Clave del CONSOLIDADO

| Celdas | Concepto | Fuente |
|---|---|---|
| `D9:D13` | TOT_OPT (Recaudo Oportuno R1) | *Recaudoporcomponente* → fila col1="Componente", col2="Total", col. F |
| `D28:D32` | R2 Total Oportuno (Anticipos) | *RerpoteDetalleSaldosaFavor* → Grand Total (E) − SERV_ESP_K |
| `D47:D51` | EXTEMP (Recaudo Extemporáneo R1) | *Recaudoporcomponente* → fila col2="Mes" en sección Extemporáneo, col. F |
| `D66:D70` | Reversión Pagos R4 | *ReversiónPorComponente* → última fila, columna 4 (valor negativo) |
| `D85:D89` | AJUSTES-SF-T | Hojas SALDOS POR NOTA + RETRIBUCION NEGATIVA (2.ª quincena) |
| `D104:D108` | Total por ASE | D9 + D28 + D47 + D66 + D85 |
| `D109` | Gran Total | SUMA(D104:D108) |

#### 3.3.4 Errores Comunes

| Error | Causa | Solución |
|---|---|---|
| TOT_OPT incompleto | Se usa fila col2="Mes" en lugar de col1="Componente" + col2="Total" | Buscar ambas condiciones |
| R2 Total Oportuno incorrecto | Fórmula referencia fila fija que cambia con subfilas | Leer Grand Total (E) directo y restar SERV_ESP_K |
| R4 con valores viejos | Se leen fórmulas almacenadas, no recalculadas | Leer última fila, col. 4, de la fuente |
| SERV_ESP_K erróneo en ASE4 | ASE4 no tiene columna "Especiales" | Verificar encabezado col. 11; si no es "Especiales", usar 0 |
| Decimales en DetRetri | Los valores del CONSOLIDADO traen centavos | Redondear a entero antes de escribir |

---

### 3.4 Propuesta Formal del Proyecto

**¿Qué es?** Documento técnico que define la propuesta formal para construir una aplicación .NET que automatice la remuneración quincenal, reemplazando el proceso manual actual.

**¿Para qué sirve?** Establece el alcance, los insumos necesarios, la arquitectura de la aplicación, el stack tecnológico y el plan de desarrollo. Es el contrato técnico entre el equipo de desarrollo y la UAESP para ejecutar el proyecto.

**Alcance del documento:**

| Sección | Contenido |
|---|---|
| Resumen Ejecutivo | Visión general del proyecto y su propósito |
| Contexto y Problema | Situación actual, problemas identificados y marco normativo (Resolución 27 de 2018) |
| Alcance — Fase 1 | Prototipo con un solo ASE, 3 archivos fuente obligatorios (R1, R2, R4), cálculos del CONSOLIDADO |
| Insumos Requeridos | Tabla detallada con nombres exactos de archivos fuente, estructura de carpetas y los 4 insumos críticos para el desarrollo |
| Arquitectura | 3 capas: Core (negocio), Infrastructure (Excel), WinForms (UI) — diseñada para migración futura |
| Diagrama de Bloques | Flujo visual completo: Entrada → Lectura → Cálculo → Escritura → Validación → Salida |
| Stack Tecnológico | .NET 10, Windows Forms, ExcelDataReader (MIT), OpenXML SDK (MIT) |
| Plan de Desarrollo | 3 fases: Prototipo (1 ASE) → Escalamiento (5 ASE) → Producción |
| Criterios de Aceptación | 7 criterios verificables para validar el prototipo |
| Riesgos | 7 riesgos identificados con mitigaciones |

**Stack tecnológico definido:**

| Componente | Tecnología | Licencia |
|---|---|---|
| Framework | .NET 10 | MIT |
| UI (prototipo) | Windows Forms | — |
| Lectura Excel | ExcelDataReader | MIT (gratis) |
| Escritura Excel | OpenXML SDK | MIT (gratis) |

**Insumos críticos para iniciar el desarrollo:**

1. Archivo de plantilla real ya utilizado (`Remuneración AAAAMM-# Total.xlsx`).
2. 1 juego completo de archivos fuente de un período (mínimo R1, R2, R4 de un ASE).
3. El resultado final de ese mismo período (archivo consolidado ya diligenciado).
4. Selección del ASE para la Fase 1.

---

## 4. Flujo del Proceso (Resumen)

```
┌─────────────────────────┐
│  Archivos fuente (5 ASE)│  Recaudoporcomponente, SaldosaFavor, ReversiónPorComponente, etc.
└───────────┬─────────────┘
            │
            ▼
┌─────────────────────────┐
│  Cálculo por ASE        │  TOT_OPT, R2, EXTEMP, R4, AJUSTES → para cada concesionario
└───────────┬─────────────┘
            │
            ▼
┌─────────────────────────┐
│  Diligenciamiento       │  Se pega en valores en cada hoja de la plantilla
│  de hojas de plantilla  │  (R1, R2, R4, recaudo por empresa, banco, etc.)
└───────────┬─────────────┘
            │
            ▼
┌─────────────────────────┐
│  CONSOLIDADO_TOTAL      │  D9:D13, D28:D32, D47:D51, D66:D70, D85:D89
│  RECAUDO                │  → D104:D108 (Total/ASE), D109 (Gran Total)
└───────────┬─────────────┘
            │
            ▼
┌─────────────────────────┐
│  Validaciones cruzadas  │  Coherencia entre hojas, diferencias ≤ ±0.5
│  + hojas de validación  │  VALIDACION_TOTAL, VALIDACION_RECIP, etc.
└───────────┬─────────────┘
            │
            ▼
┌─────────────────────────┐
│  Entregable final       │  "Remuneración AAAAMM-# Total.xlsx"
└─────────────────────────┘
```

---

## 5. Principios Rectores

1. **Lectura directa de la fuente** — Siempre leer archivos fuente del período; nunca confiar en valores previos.
2. **Conservación de la estructura** — Mantener la distribución de hojas del archivo de plantilla.
3. **Integridad de fórmulas** — No alterar fórmulas existentes; reemplazar solo valores.
4. **Coherencia entre hojas** — Validar consistencia de totales, diferencias y reversiones.
5. **Entregable único** — Un solo archivo final con nomenclatura `Remuneración AAAAMM-# Total.xlsx`.

---

## 6. Nomenclatura y Períodos

| Parámetro | Quincena 1 | Quincena 2 |
|---|---|---|
| Período | Días 1–15 | Días 16–31 |
| Código | AAAAMM1 | AAAAMM2 |
| Archivo | Remuneración AAAAMM-1 Total.xlsx | Remuneración AAAAMM-2 Total.xlsx |

---

## 7. Notas Importantes

- La hoja **"Informe AFaseo Recaudo"** se ignora para efectos del consolidado.
- Los reportes de **SALDOS POR NOTA** y **RETRIBUCION NEGATIVA** solo aplican a la **segunda quincena** del mes.
- La **Interventoría** se consulta en el rango L25:N31 de su hoja específica.
- La tolerancia de diferencia entre el CONSOLIDADO y los reportes de control de UAESP es de **±0.5** por redondeo.
