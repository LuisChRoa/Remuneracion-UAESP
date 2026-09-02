# Requerimientos — Fase 1: Prototipo de Remuneración Quincenal (Libro Completo, 1 ASE)

> **Proyecto:** Automatización de la Remuneración Quincenal del Servicio de Aseo (UAESP)
> **Documento:** Guía de requerimientos para el desarrollo — Fase 1
> **Versión:** 1.0 — Agosto 2026
> **Alcance:** Prototipo funcional que elabora el libro consolidado completo de **un (1) ASE**, a partir de las 8 fuentes y el archivo de plantilla.
> **Trazabilidad:** derivado de `Docs/Propuesta_Proyecto_Automatizacion_Remuneracion_UAESP.md` (en adelante *Propuesta*) y `Docs/Prompt_Maestro_Proyecto_Remuneracion_UAESP _ Vo .docx` (en adelante *Prompt Maestro*).

---

## 1. Introducción y propósito

Este documento es la **guía única de desarrollo** de la Fase 1. Define qué se construye, con qué insumos, bajo qué reglas de negocio y cómo se valida. Toda implementación, prueba y criterio de aceptación debe trazarse a una de las historias de usuario de la Sección 5.

El objetivo de la Fase 1 es reemplazar el proceso manual (hoy ejecutado por IA a partir del Prompt Maestro) por una aplicación de escritorio .NET 10 que, para un ASE, lea las fuentes, calcule y diligencie **todas las hojas** del libro de remuneración, preservando fórmulas y validando coherencia.

---

## 2. Alcance de la Fase 1

| Concepto | Alcance de esta fase |
|---|---|
| **ASE a procesar** | Un (1) ASE — seleccionable (se recomienda comenzar por **ASE1 Promoambiental**; decisión final de UAESP). |
| **Período base recomendado** | **2.ª quincena** (ej. `202607-2`) para cubrir las hojas `AJUSTES-SF-T` (solo aplicables en 2.ª quincena). |
| **Fuentes obligatorias** | Las 8 descritas en la Sección 3. |
| **Hojas a diligenciar** | Todas las del libro: `CONSOLIDADO_TOTAL RECAUDO`, `Reporte Componentes R1`, `Rem. Anticipos R2`, `Reversion Pagos R4`, `DetRetri`, `Recaudo EAAB Reciprocidad`, `EAAB + Ciud Limp`, `ENEL`, `ENERBIT`, `Directa Occidente`, `REPORTE RECAUDO x BANCO`, `ANT EXT-REV`, `ANTICIPOS USUARIOS`, `BCE SC POR FACT`, `SALDOS POR NOTA`, `RETRIBUCION NEGATIVA`, `AJUSTES - SF-T`, `INTERVENTORIA` (rango L25:N31). |
| **Validaciones** | Básicas de coherencia + validaciones cruzadas (`VALIDACION_TOTAL`, `VALIDACION_RECIP`). |
| **UI** | Windows Forms: selección de período/carpeta/plantilla, progreso, resultados, log. |
| **Fuera de alcance (aún)** | Los otros 4 ASE (Fase 2), modo CLI (Fase 3), documentación de usuario final (Fase 3). |

> **Nota de mentoría:** Esta Fase 1 está **ampliada** respecto a la *Propuesta* §4.1 (que limitaba a CONSOLIDADO + hojas R1/R2/R4). La ampliación fue decidida para entregar el libro completo de 1 ASE y solicitar los insumos completos de una vez.

---

## 3. Prerrequisitos — Insufos a solicitar a UAESP (CRÍTICOS)

Sin estos insumos no es posible verificar la Fase 1. Estado actual del repositorio: `Docs/Insumos/` contiene `Remuneracion 202607-1 Total.xlsx`, `Remuneracion 202607-2 Total.xlsx` y `A1 _ R4-BalanceSubsidioyContribuciones_...xlsx`. **Faltan** los fuentes obligatorios y el consolidado golden.

| # | Insumo | Nombre exacto (patrón) | Estado | Para qué se requiere |
|---|---|---|---|---|
| 1 | Plantilla real ya utilizada | `Remuneración AAAAMM-# Total.xlsx` | Parcial (hay 202607-1/2, pero sin fuentes para validar) | Verificar posiciones de celdas, encabezados y fórmulas |
| 2 | Fuente R1 | `Recaudoporcomponente_to_date{INI}ddMMyyyy_to_date{FIN}ddMMyyyy___{ts}.xlsx` | **FALTA** | TOT_OPT, EXTEMP |
| 3 | Fuente R2 | `ReporteDetalleSaldosaFavor_to_date{INI}...___{ts}.xlsx` | **FALTA** | Anticipos / saldos a favor |
| 4 | Fuente R4 (Reversión) | `ReversiónPorComponente_to_date{INI}...___{ts}.xlsx` | **FALTA** | Reversión de pagos |
| 5 | Reporte por banco | `ReportePagosxBanco_to_date{INI}...___{ts}.xlsx` | **FALTA** | `REPORTE RECAUDO x BANCO`, `ANTICIPOS USUARIOS` |
| 6 | Recaudos reversados | `RecaudosReversados_to_date{INI}...___{ts}.xlsx` | **FALTA** | `ANT EXT-REV` |
| 7 | Balance SC | `R4-BalanceSubsidioyContribuciones_to_date{INI}...___{ts}.xlsx` | Presente (202607) | `BCE SC POR FACT` |
| 8 | Saldos por nota | `SaldosaFavorAplicadosPorNotas_to_date{INI}...___{ts}.xlsx` | **FALTA** | `AJUSTES - SF-T` (solo 2.ª quincena) |
| 9 | Retribución negativa | `RetribuciónNegativa_to_date{INI}...___{ts}.xlsx` | **FALTA** | `AJUSTES - SF-T` (solo 2.ª quincena) |
| 10 | Consolidado ya diligenciado (golden) | `Remuneración AAAAMM-# Total.xlsx` del mismo período | **FALTA** | Golden test (comparar salida vs resultado real) |

> ⚠️ **Solicitud recomendada:** un período de **2.ª quincena** completo (5 carpetas ASE con R1/R2/R4 mínimo, y preferiblemente los 8 fuentes), más el consolidado manual ya diligenciado de ese mismo período como golden. Ej. período `202607-2`.

---

## 4. Mapa de insumos → historia

| Documento | Historia(s) | Cuándo aplica |
|---|---|---|
| `Recaudoporcomponente_*.xlsx` (R1) | HU-02, HU-04 | Ambas quincenas |
| `ReporteDetalleSaldosaFavor_*.xlsx` (R2) | HU-02, HU-04 | Ambas quincenas |
| `ReversiónPorComponente_*.xlsx` (R4) | HU-02, HU-04 | Ambas quincenas |
| `ReportePagosxBanco_*.xlsx` | HU-02, HU-04 (`REPORTE RECAUDO x BANCO`, `ANTICIPOS USUARIOS`) | Ambas quincenas |
| `RecaudosReversados_*.xlsx` | HU-02, HU-04 (`ANT EXT-REV`) | Ambas quincenas |
| `R4-BalanceSubsidioyContribuciones_*.xlsx` | HU-02, HU-04 (`BCE SC POR FACT`) | Ambas quincenas |
| `SaldosaFavorAplicadosPorNotas_*.xlsx` | HU-02, HU-04 (`AJUSTES - SF-T`) | **Solo 2.ª quincena** |
| `RetribuciónNegativa_*.xlsx` | HU-02, HU-04 (`AJUSTES - SF-T`) | **Solo 2.ª quincena** |
| `Remuneración AAAAMM-# Total.xlsx` (plantilla) | HU-04 (template), HU-05 (golden) | Entrada y validación |

---

## 5. Historias de usuario

### HU-01 — Andamiaje, dominio y localizador de fuentes
- **Alcance:** Solución de software estructurada en tres capas sobre .NET 10 (núcleo de dominio, acceso a Excel e interfaz de escritorio), compilable sin errores ni advertencias, con un localizador de archivos que encuentra automáticamente las carpetas de los 5 ASE y los 8 reportes fuente por su nombre.
- **Entregable:** Una solución de software estructurada en tres capas sobre .NET 10 (núcleo de dominio, acceso a Excel e interfaz de escritorio), compilable sin errores ni advertencias, con un localizador de archivos que encuentra automáticamente las carpetas de los 5 ASE y los 8 reportes fuente por su nombre.
- **Insumos / Dependencias:** .NET 10 SDK; paquetes ya referenciados (ExcelDataReader 3.9.0, OpenXML SDK 3.5.1, Serilog 4.4.0/7.0.0). ⚠️ Confirmar nombres reales de los 8 archivos (`Rerpote` vs `Reporte`, prefijos exactos) para no hardcodear.
- **Criterios de aceptación:**
  - La solución compila con 0 warnings (`dotnet build` limpio).
  - `ArchivoFuenteLocator` resuelve la carpeta de un ASE dado y localiza cada patrón de archivo fuente por prefijo.
- **Trazabilidad:** *Propuesta* §6 (Arquitectura), §9 (Plan 1.1–1.2).

### HU-02 — Lectura de las 8 fuentes (readers)
- **Alcance:** Componentes lectores (*readers*) de las ocho fuentes, que abren cada Excel, interpretan sus encabezados dinámicamente y extraen los valores numéricos correctos por ASE (recaudo oportuno/extemporáneo, anticipos, reversiones, recaudo por banco, anulados, subsidios/contribuciones y ajustes de nota/retribución negativa), listos para ser consumidos por el cálculo.
- **Entregable:** Los componentes lectores (*readers*) de las ocho fuentes, que abren cada Excel, interpretan sus encabezados dinámicamente y extraen los valores numéricos correctos por ASE (recaudo oportuno/extemporáneo, anticipos, reversiones, recaudo por banco, anulados, subsidios/contribuciones y ajustes de nota/retribución negativa), listos para ser consumidos por el cálculo.
- **Insumos / Dependencias:** Los 8 archivos fuente (**CRÍTICOS, faltan en repo**). Dep: HU-01.
- **Criterios de aceptación:**
  - Dado un archivo R1 del ASE seleccionado, el reader extrae TOT_OPT y EXTEMP coincidiendo con la fila `col1=Componente`/`col2=Total` y la sección Extemporáneo, dentro de tolerancia ±0,5 vs valor manual.
  - Para R2, el Grand Total y `SERV_ESP_K` se calculan según encabezado real de col. 11 (ASE4 → 0).
  - Para R4 (Reversión), se lee la última fila, columna 4 (valor negativo).
  - Los readers de banco, reversados, balance y ajustes (Q2) extraen sus valores en las columnas correctas.
- **Trazabilidad:** *Prompt Maestro* §5.4, §7.1–7.5, Part II; *Propuesta* §5.1.

### HU-03 — Motor de cálculo del CONSOLIDADO + hojas derivadas (1 ASE)
- **Alcance:** Motor de cálculo (`ICalculoRemuneracion`) que, a partir de los datos leídos, calcula todos los montos del CONSOLIDADO y de las hojas derivadas (totales por ASE, gran total, `DetRetri` redondeado, anticipos, ajustes y balance), aplicando las reglas de negocio y la tolerancia de ±0,5.
- **Entregable:** El motor de cálculo (`ICalculoRemuneracion`) que, a partir de los datos leídos, calcula todos los montos del CONSOLIDADO y de las hojas derivadas (totales por ASE, gran total, `DetRetri` redondeado, anticipos, ajustes y balance), aplicando las reglas de negocio y la tolerancia de ±0,5.
- **Insumos / Dependencias:** Reglas del *Prompt Maestro* §7 (Partes I y II). Dep: HU-02.
- **Criterios de aceptación:**
  - `D9:D13` = TOT_OPT, `D28:D32` = R2 Total Oportuno, `D47:D51` = EXTEMP, `D66:D70` = Reversión R4, `D85:D89` = AJUSTES-SF-T (0 si no hay instrucciones / Q1), `D104:D108` = suma de bloques por ASE, `D109` = SUMA(D104:D108).
  - `DetRetri` columna D = ROUND(D104:D108, 0).
  - Las filas "Total Oportuno" de `Reporte Componentes R1` coinciden con D9:D13.
- **Trazabilidad:** *Prompt Maestro* §7.1–7.6, §8; *Propuesta* §7 (Diagrama).

### HU-04 — Escritura de plantilla (OpenXML) — TODAS las hojas
- **Alcance:** Escritor de plantilla (`IPlantillaWriter`) que vuelca todos los valores calculados en cada hoja del libro de remuneración, en las celdas correctas y **sin alterar ninguna fórmula existente**, produciendo el archivo `Remuneración AAAAMM-# Total.xlsx` completo.
- **Entregable:** El escritor de plantilla (`IPlantillaWriter`) que vuelca todos los valores calculados en cada hoja del libro de remuneración, en las celdas correctas y sin alterar ninguna fórmula existente, produciendo el archivo `Remuneración AAAAMM-# Total.xlsx` completo.
- **Insumos / Dependencias:** Plantilla real `Remuneración AAAAMM-# Total.xlsx` (**CRÍTICO**; en `Docs/Insumos` hay 202607-1/2 pero sin fuentes para validar); los 8 archivos ya leídos en HU-02. Dep: HU-03.
- **Criterios de aceptación:**
  - Todas las hojas listadas en §2 contienen los valores calculados en las celdas correctas.
  - Ninguna fórmula de la plantilla es modificada (solo se escriben valores).
  - La hoja `INTERVENTORIA` se diligencia desde el rango L25:N31 de la fuente correspondiente.
  - Se ignora `Informe AFaseo Recaudo`.
- **Trazabilidad:** *Prompt Maestro* Part II (detalle por hoja); *Propuesta* §7 (Escritura).

> **Rebase 2026-09-02:** la HU-05 vigente de implementación es `plans/05 - HU-05 WorkbookLeafInputs y escritura real.md` (leaf inputs + escritura real). La historia original de validación/UI/golden queda como **HU-06** y se ejecuta después de certificar la escritura leaf.

### HU-05 legado — Validación, UI y pruebas golden (1 ASE) — ahora HU-06
- **Alcance:** Aplicación de escritorio ejecutable (WinForms) con selección de período/carpeta/plantilla, barra de progreso y registro de auditoría, acompañada de una prueba automatizada (xUnit) que contrasta la salida del sistema contra el consolidado real pre-diligenciado para certificar que coinciden dentro de la tolerancia.
- **Entregable:** La aplicación de escritorio ejecutable (WinForms) con selección de período/carpeta/plantilla, barra de progreso y registro de auditoría, acompañada de una prueba automatizada (xUnit) que contrasta la salida del sistema contra el consolidado real pre-diligenciado para certificar que coinciden dentro de la tolerancia.
- **Insumos / Dependencias:** Consolidado ya diligenciado del mismo período (**golden, CRÍTICO, falta**). Dep: HU-04.
- **Criterios de aceptación:**
  - D109 = SUMA(D104:D108) y validaciones `VALIDACION_TOTAL`/`VALIDACION_RECIP` sin discrepancias.
  - Diferencia del consolidado vs golden ≤ ±0,5.
  - `DetRetri` en valores enteros.
  - El log registra cada paso (timestamp, archivos leídos, valores extraídos/escritos).
  - La UI permite seleccionar período, carpeta fuente y plantilla, y muestra progreso/resultados.
- **Trazabilidad:** *Propuesta* §10 (Criterios de Aceptación), §8 (Stack/Serilog).

---

## 6. Reglas de negocio transversales

1. **TOT_OPT (R1):** fila donde `col1="Componente"` **Y** `col2="Total"` → columna F. No usar la fila `col2="Mes"` (omite "Aplicación nuevos x reversión").
2. **R2 Total Oportuno:** `Grand Total` (col E, última fila) − `SERV_ESP_K`. `SERV_ESP_K` se toma de la columna 11 solo si su encabezado es `"Especiales"`; en ASE4 la columna 11 es `"Componente TCS"` → `SERV_ESP_K = 0`.
3. **EXTEMP (R1):** fila `col2="Mes"` en la sección Extemporáneo → columna F.
4. **Reversión R4:** última fila, columna 4 (valor negativo).
5. **AJUSTES-SF-T:** normalmente 0; solo se diligencia en 2.ª quincena con `SALDOS POR NOTA` + `RETRIBUCION NEGATIVA`.
6. **Totales:** `D104 = D9 + D28 + D47 + D66 + D85` por ASE; `D109 = SUMA(D104:D108)`.
7. **DetRetri:** columna D = ROUND(valor, 0) (enteros).
8. **Consistencia interna:** filas "Total Oportuno" de `Reporte Componentes R1` = D9:D13.
9. **Tolerancia:** diferencia máxima tolerable ±0,5 por redondeo.
10. **Integridad de fórmulas:** nunca alterar fórmulas; pegar solo valores.
11. **Parsing dinámico:** detectar encabezados y posiciones por contenido, no hardcodear números de fila/columna fijos.
12. **Ignorar** la hoja `Informe AFaseo Recaudo` para efectos del consolidado.
13. **Interventoría:** rango L25:N31 de su hoja.

---

## 7. Criterios de aceptación generales (Fase 1)

| # | Criterio | Verificación |
|---|---|---|
| CA-1 | La app lee correctamente los 8 archivos fuente del ASE seleccionado | Comparación visual de valores extraídos vs archivo original |
| CA-2 | TOT_OPT, R2 Total Oportuno, EXTEMP, Reversión R4 y AJUSTES-SF-T coinciden con cálculo manual | Prueba con datos conocidos |
| CA-3 | El archivo de salida tiene valores correctos en CONSOLIDADO D9:D13, D28:D32, D47:D51, D66:D70, D85:D89, D104:D109 | Comparación contra golden |
| CA-4 | Las fórmulas de la plantilla NO fueron alteradas | Verificación de celdas con fórmulas |
| CA-5 | Las validaciones básicas y cruzadas pasan sin errores | Ejecución exitosa |
| CA-6 | El log registra cada paso de la ejecución | Revisión del log |
| CA-7 | La UI permite seleccionar período, carpeta y plantilla | Prueba funcional |
| CA-8 | `DetRetri` se escribe en valores enteros y las filas "Total Oportuno" R1 = D9:D13 | Revisión de hoja y validación |

---

## 8. Riesgos de Fase 1

| # | Riesgo | Prob. | Impacto | Mitigación |
|---|---|---|---|---|
| R1 | No se dispone de la plantilla real | Alta | Crítico | Solicitar a UAESP un archivo ya utilizado (Insumo #1) |
| R2 | No se dispone de las fuentes de ejemplo | Alta | Crítico | Solicitar un período completo ya procesado (Insumos #2–#9) |
| R3 | Posiciones de celdas varían entre plantillas | Media | Alto | Detectar encabezados dinámicamente, no hardcodear |
| R4 | Fórmulas dependen de filas dinámicas | Media | Alto | OpenXML preserva fórmulas; pegar solo valores |
| R5 | Errores de parsing en fuentes | Media | Medio | Validación de estructura antes de procesar |
| R6 | Cambio en formato de reportes fuente | Baja | Medio | Parsers configurables por archivo |
| R7 | Curva de aprendizaje OpenXML SDK | Media | Bajo | Wrappers que simplifiquen operaciones comunes |

---

## 9. Próximos pasos

1. **Aprobar este documento** y confirmar ASE y período base (recomendado: ASE1, 2.ª quincena).
2. **Solicitar a UAESP** los insumos críticos faltantes (Sección 3).
3. Derivar a `sdd-planner` la generación del plan SDD en `plans/` a partir de estas historias.
4. Implementar por historia, con golden test en HU-05.

> **Control de cambios:** cualquier ajuste a alcance, insumos o reglas de negocio debe actualizar este documento y registrarse en su control de versiones.
