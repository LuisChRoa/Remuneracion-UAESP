# Plan 33 — T0 Evidencia Agosto-2 (Fase 2, T4) — Veredicto H1/H2/H3

> **Tarea:** T4 (Plan 33), evidencia read-only. **Cero código de producción** (`git diff` de producción vacío para T4).
> **Síntoma:** `Espejo R1 (Plan 32/T2-interior): ASE 2 Reporte Componentes R1!F199 [SUB_EMP]: faltan las anclas [Mes0] del sub-bloque; no se escribe parcial.` (aborta el pipeline de agosto-2026082, `ERR-PLANTILLA`).
> **Método:** reproducción instrumentada in-process (lectura ZIP+XML BCL, sin Excel/COM): se abren los 5 bloques R1 de la fuente de agosto con `ExcelDataReaderWorkbookLeafInputReader.LeerEspejoR1` y se invoca `OpenXmlEspejoR1Mutador.AjustarEnWorkbook` sobre la base canónica; se captura la excepción y se inspecciona el workbook mutado en memoria (dimensionado + visibles recompuestos, antes del `Save`).
> **Rutas de disco (post-reorg `a867706`):** la evidencia originalmente citada como `Docs/Prueba2/...` vive hoy en `Docs/Prueba Agosto-2/...` (rename R100, mismos bytes). Se cita la ruta vigente.

---

## 0. Fuentes (disco)

| Rol | Archivo | Hoja / celda |
|---|---|---|
| Base canónica (la que usa el pipeline) | `Docs/Prueba Agosto-2/Plantilla_Remuneracion_2026082.xlsx` | `Reporte Componentes R1` (`sheet10.xml`) |
| Oráculo del administrativo | `Docs/Prueba Agosto-2/Resultado/Resultado Manual por el administrativo/Remuneracion 202608-2 Total_7721.xlsx` | `Reporte Componentes R1` |
| Fuente R1 ASE2 (LIME) agosto | `Docs/Prueba Agosto-2/Insumos/2-Lime/Recaudoporcomponente_to_date16082026ddMMyyyy_to_date31082026ddMMyyyy___202691163346577.xlsx` | `Sheet1` (C = empresa, D = `Total`, B vacía) |
| Mapa de producción | `Remuneracion.Infrastructure/Excel/OpenXmlEspejoR1Mutador.cs` → `CeldasInterioresAgosto[2]` | `F199 SUB_EMP`, `F204 SUB_EMP`, `F214 SUB_EMP` |

## 1. Geometría real post-dimensionado (app) vs manual

Tras `AjustarEnWorkbook` (capturado antes del throw):

| Ancla | App post-dimensionado | Manual (oráculo) |
|---|---|---|
| Nombre bloque ASE2 (`B='LIME'`) | fila **83** | fila **83** |
| Visible `TOTAL/OPORTUNO` (F) | fila **194** | fila **194** (`F194`) |
| `C199` (empresa del `F199 SUB_EMP`) | **RECIPROCIDAD** | **RECIPROCIDAD** |
| `C204` (empresa del `F204 SUB_EMP`) | **ENEL** | **ENEL** |
| `C214` (empresa del `F214 SUB_EMP`) | **OCCIDENTE** | **OCCIDENTE** |

→ **La geometría y las celdas del mapa COINCIDEN con el manual.** `CeldasInterioresAgosto[2]` no está desalineado en sus referencias.

## 2. Empresas de la fuente R1 ASE2-agosto

Filas-dato de empresa (`C` no vacía ∧ `B` vacía ∧ `D=Total`) de la fuente LIME:

`ENEL, OCCIDENTE, ENEL, OCCIDENTE, ENEL, NUEVO ESQUEMA, OCCIDENTE, ENEL, ENEL, NUEVO ESQUEMA, OCCIDENTE.`

**Conjunto:** `{ENEL, OCCIDENTE, NUEVO ESQUEMA}`. **No existe `RECIPROCIDAD`.**

Sello del manual: `Reporte Componentes R1!F199 = F140+F114-L114`, y las filas `F140`/`F114` del manual son subtotales-dato **`NUEVO ESQUEMA`** (ver manual filas 114 y 140). Es decir: el subtotal rotulado `RECIPROCIDAD` en el template/manual resume datos de **`NUEVO ESQUEMA`**.

## 3. Anclaje en negocio (los 3 .docx)

- **`Docs/Detalle de plantilla.docx`** (hoja `REPORTE RECAUDO x BANCO`, definición de títulos): *«**NUEVO ESQUEMA: Corresponde al recaudo "EAAB Reciprocidad", reportado por la empresa de facturación conjunta EAAB, en el esquema del servicio de aseo vigente**»*; y *«en los casos de **LIME** y BOGOTA LIMPIA se encontrará recaudo con nombres de columnas de **ENEL, NUEVO ESQUEMA y OCCIDENTE**»*. Además (hoja `REPORTE COMPONENTES R1`): *«…teniendo la validación por fuente de recaudo considerando que **no en todas las quincenas se encuentra la misma información**»*.
- **`Docs/Proceso de Recaudo.docx`** (matriz EFC×ASE, tabla 1 y 2): facturación conjunta `{ENEL, EAAB, ENERBIT}` → **ASE 2 = ENEL + EAAB**; facturación directa `OCCIDENTE` → todos los ASE. `RECIPROCIDAD` es la marca histórica de **EAAB**.
- **`Docs/Prompt_Maestro...Vo.docx`** (§Proceso 4–5): *«Conserva todas las fórmulas existentes; solo reemplaza información donde corresponda en valores»* y *«Si el número de filas cambia en un reporte fuente, inserta o elimina filas preservando las fórmulas»*.

**Conclusión de negocio:** `RECIPROCIDAD` (template, legado) y `NUEVO ESQUEMA` (fuente, vigente) son **la misma EFC (EAAB Reciprocidad)**. La fuente trae el nombre vigente.

## 4. Veredicto por hipótesis

| Hipótesis | Veredicto | Evidencia |
|---|---|---|
| **H1 — nombre divergente en `C199` vs fuente** *(causa raíz)* | **CONFIRMADA** | `C199` (app y manual) = **`RECIPROCIDAD`**; la fuente LIME trae **`NUEVO ESQUEMA`** (= EAAB Reciprocidad, `Detalle de plantilla`). La recomposición del interior resuelve la empresa **por el nombre del template** (`DerivarFilasPorFirma` → `R1FirmaInterior.EsDatoEmpresa(fila, empresa)`, `empresa = C199`) y no halla filas de `RECIPROCIDAD` en la fuente → `faltan las anclas [Mes0]`. |
| **H2 — filas oportunas ausentes en la fuente** | **DESCARTADA como causa raíz** | La fuente SÍ trae las filas oportunas del sub-bloque, bajo `{ENEL, NUEVO ESQUEMA, OCCIDENTE}` (11 filas-dato de empresa). Solo el rótulo legado `RECIPROCIDAD` no tiene coincidencia; es un efecto de H1, no una ausencia real de datos. |
| **H3 — `CeldasInterioresAgosto[2]` stale para 2026082** | **DESCARTADA** | La celda del mapa `F199` (y `F204`/`F214`) corresponde a la geometría real de agosto: post-dimensionado `B=LIME`=83, `TOTAL/OPORTUNO`=194 y `C199/C204/C214 = RECIPROCIDAD/ENEL/OCCIDENTE`, idénticos al manual. El mapa no está desalineado en sus referencias; el problema es el **rótulo** leído del template, no la celda del mapa. |

**Veredicto:** **H1 (nombre divergente, sinónimo legado `RECIPROCIDAD` ↔ vigente `NUEVO ESQUEMA`)**. H2 y H3 descartadas como causa raíz. El aborto es determinista por el filtro por-nombre contra un rótulo que la fuente ya no usa.

## 5. Hallazgos secundarios (para el diseño de T5, no son la causa del aborto)

1. **Mapa incompleto para ASE2:** el manual trae 5 subtotales-empresa en el interior de ASE2 (`C`= `RECIPROCIDAD` 199, `ENEL` 204, `ENERBIT` 209, `OCCIDENTE` 214, `CIUDAD LIMPIA-ACUEDUCTO` 219). `CeldasInterioresAgosto[2]` solo lista `F199/F204/F214`; **omite `F209` (ENERBIT) y `F219` (CIUDAD LIMPIA-ACUEDUCTO)** (ambos `F=0`, clase L-1 del TODO(T2-full) del Plan 32). No causa el aborto, pero queda como brecha del interior de ASE2.
2. **Doctrina declarada vs implementada:** `R1FirmaInterior` documenta alinear el sub-bloque fuente ↔ filas finales **por ORDEN** («nunca por enésima ocurrencia», D-B), pero `DerivarFilasPorFirma` filtra **por nombre** (`empresa = C<celda>`). La divergencia sinónimo-legado↔vigente rompe el filtro por nombre.
3. **Dirección acotada del fix (NO prescrita aquí, es de T5):** resolver el sinónimo rótulo-template ↔ empresa-fuente (o alinear por la secuencia de la fuente) para que `RECIPROCIDAD` y `NUEVO ESQUEMA` ocupen el mismo slot del sub-bloque. **T5 no se implementa en este documento** (Plan 33 Fase 2 lo deja condicional a este veredicto; el diseño vive en la tarea T5).

## 6. Estado

- T4: **CERO código de producción** (`git diff` de `Remuneracion.*` = solo el fix de Fase 1/T2, ajeno a Fase 2).
- Las citas a `Docs/Prueba2/...` del plan quedaron **stale** por el reorg `a867706`; se citan las rutas vigentes `Docs/Prueba Agosto-2/...`.
