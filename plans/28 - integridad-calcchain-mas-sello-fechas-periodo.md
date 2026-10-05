# Plan 28 — Integridad estructural del workbook de salida (calcChain) + sello de fechas del período

> **Alcance:** dos unidades sin acoplamiento aritmético. **Unidad S (YA APROBADA en sesión, Alternativa A):** el workbook de salida deja de producir el diálogo de reparación de Excel ("Registros quitados: Fórmula de /xl/calcChain.xml") eliminando `CalculationChainPart` al guardar + `FullCalculationOnLoad=true`, centralizado en un punto único de guardado para que NINGÚN camino quede sin la limpieza; test nuevo de integridad estructural como blindaje. **Unidad F (núcleo a ratificación junto a este plan):** sellar las fechas del período (`CONSOLIDADO_TOTAL RECAUDO` G7 = Fecha Desde, K7/J7 = Fecha Hasta — columna exacta la fija T0) desde `DetRetriInputs` (fechas ya leídas del R10), SOLO valores sobre celdas estáticas, con fail-fast si el R10 no trae fechas legibles; la plantilla en ceros queda agnóstica de período y reutilizable.
> **Rector (IRRENUNCIABLE):** `README.md` + `Docs/Propuesta_Proyecto_Automatizacion_Remuneracion_UAESP.md` + `Docs/Detalle de plantilla.docx` + `Docs/Prompt_Maestro_Proyecto_Remuneracion_UAESP _ Vo .docx`. **Invariantes del proyecto:** NUNCA sobrescribir fórmulas (solo valores; G7/K7 son celdas estáticas — a confirmar en T0; si fueran fórmula → fail-fast, nunca escritura); tolerancia ±0.5; ASE4 sin columna "Especiales" (0 si ausente); Q2 aplica SALDOS POR NOTA/RETRIBUCION NEGATIVA (`Periodo.NumeroQuincena == 2`, `AjustesSfT = 0` en Q1); redondear a entero antes de DetRetri/columna D; build 0 warnings; finales de línea CRLF; fail-fast nombrando archivo/celda; tests SOLO con insumos reales (`Docs/Insumos`, `Docs/Prueba2`), NO fixtures sintéticas; sin commits (los hace el Ingeniero con `#commit`); sin emojis.
> **Continuidad:** HU-01..HU-20 cerradas; Planes 21 (espejo R1 + `OpenXmlEspejoR1Mutador`), 23 (opcionalidad 2.5), 25 (roles por firma + DetRetri 5/5 vs R10), 26 (preflight) y 27 (firma + hoja-por-nombre, suite **339/339**) cerrados. Este plan NO reabre ninguna semántica de lectura, cálculo, escritura de valores ni validación fuera de lo declarado: **cero cambios de fórmulas de negocio (`<f>` ni `<v>` se tocan); Q1/Q2-julio intactos por construcción; goldens julio ±0.5 intactos; DetRetri 5/5 vs R10 intacto; cero cambios de plumbing UI/CLI.**
> **Insumos canónicos (no se re-fijan, no se inventan datos):** plantilla en ceros `Docs/Insumos/REMUNERACION 2026072/Plantilla_ Remuneracion 202607-2.xlsx`; goldens `Docs/Insumos/Remuneracion 202607-1 Total.xlsx` / `Docs/Insumos/Remuneracion 202607-2 Total.xlsx`; R10 `Docs/Insumos/REMUNERACION 2026071/R10_Remuneracion_2026071.xlsx` (01–15/07), `Docs/Insumos/REMUNERACION 2026072/R10_Remuneracion_2026072.xlsx` (16–31/07), `Docs/Prueba2/Insumos/R10_Remuneracion_2026082.xlsx` (16–31/08, a re-verificar en T0); período agosto `Docs/Prueba2/Insumos` (5 ASE + `Conciliaciones/`).
> **Numeración:** `plans/` 01..27 ocupados; este plan toma el primer correlativo libre, **28**.
> **Estado:** IMPLEMENTADO — T0 + Unidad S + Unidad F + tests + docs ejecutados y verificados (build 0/0; suite 345/345; salida agosto regenerada desde ceros sin calcChain + fullCalcOnLoad + G7/K7=16/08–31/08; DetRetri 5/5 vs R10). Sin commit (lo hace el Ingeniero). Ver §6.
> **Fecha:** 2026-10-05

---

## 0. Clarification Gate

**Unidad S ya decidida en sesión (Alternativa A aprobada; se incluye como primera unidad sin fork). Unidad F con diseño pedido explícito y una sola decisión con fork real (D-F en §0.3: alcance del sello — solo G7/K7 vs G7/K7 + otros rótulos estáticos que T0 encuentre): no hay preguntas bloqueantes.** El plan viene redactado con las decisiones D-A..D-G (§0.3); si el Ingeniero discrepa de alguna, solo cambia la tarea acotada que la implementa. Punto de bloqueo respetado: NO se implementa nada en este documento.

### 0.1 Qué está verificado y qué NO (honestidad técnica)

**Verificado en esta planificación** (lectura directa de código, 2026-10-05):

| # | Hecho | Método | Consecuencia |
|---|---|---|---|
| V1 | Hoja consolidado = `CONSOLIDADO_TOTAL RECAUDO` (`WorkbookLeafCellMap.cs:9`, re-exportada en `PorAse`/`Q2`) | Lectura código | El sello F escribe sobre esa hoja; el nombre no se inventa |
| V2 | Dos caminos de `Save()` en el writer + un tercer camino espejo standalone: single-ASE `GenerarWorkbook` L145-147 (`workbookXml.Save()`), multi-ASE L285-287 (mismo), con `AjustarEnWorkbook` en el medio (L259); `EscribirEspejoR1ConResultado` → `OpenXmlEspejoR1Mutador.Ajustar` (origen→salida, `OpenXmlPlantillaWriter.cs:353-361`, `OpenXmlEspejoR1Mutador.cs:46-84`) | Lectura código | La limpieza calcChain debe centralizarse para que NINGÚN camino quede sin ella (D-B); T0 re-verifica el inventario exacto en disco |
| V3 | CERO manejo de `CalculationChain`/`calcPr`/`FullCalculationOnLoad` en el código (grep sin matches) | Grep código | El hueco es real al símbolo: el mutador inserta/borra filas (doctrina Plan 21) y nada sanea la cadena de cálculo |
| V4 | `DetRetriInputs.FechaDesde/FechaHasta` existen (modelo) y SE LEEN del R10 (`ExcelDataReaderDetRetriR10Reader.cs:88-89`, `BuscarFecha` G7/J7 por etiqueta), pero grep `FechaDesde\|FechaHasta` = solo modelo + lector + tests → NADA las escribe al workbook | Lectura + grep | El sello F es escritura nueva acotada a 2 celdas; sin riesgo de doble-escritura |
| V5 | La quincena (selector C59) YA se escribe por otro camino: `EscribirValorNumerico(HojaBanco, "C59", leaf.ReporteBanco.Quincena)` (`OpenXmlPlantillaWriter.cs:1200-1201`, "dominio, nunca fuente") | Lectura código | C59 no se duplica en el sello F; T0 solo documenta su gobierno (D-F) |
| V6 | Guard anti-fórmula vigente: `EscribirValorNumerico` L1261-1268 lanza `ERR-PLANTILLA` si la celda destino es fórmula | Lectura código | G7/K7 pasan por ese guard por construcción: si fueran fórmula → fail-fast honesto, nunca sobrescritura (D-D) |
| V7 | R10 julio con fechas legibles (tests `GoldenResumenMesYR10Tests.cs:119-120` Q1 = 01–15/07; `:137-138` Q2 = 16–31/07); plantilla en ceros y R10-2026082 existen en disco (`Docs/Insumos/REMUNERACION 2026072/`, `Docs/Prueba2/Insumos/`) | Lectura código + listado disco | Base real para T0 (fechas agosto a re-verificar) y para los tests |
| V8 | `ProcesadorPeriodo` ya lee el R10 (L294-300) y valida DetRetri-vs-R10 en AMBAS quincenas (L314-326) ANTES de `GenerarWorkbook` (L330) | Lectura código | El sello F se enchufa tras la lectura del R10 y antes de generar el workbook, sin mover el orden del flujo |
| V9 | `plans/` 01..27 ocupados; suite base **339/339** (cierre Plan 27 §6.3) | Listado + Plan 27 | Numeración 28; red ciega 339/339 |
| V10 | Discrepancia a fijar en T0: el encargo nombra `CONSOLIDADO_TOTAL RECAUDO` G7/K7, pero `DetRetriInputs`/lector documentan el rango del R10 en G7/**J7** | Lectura código vs encargo | T0 fija la columna real de "Fecha Hasta" en LA PLANTILLA (K7 vs J7) y en el R10; la SPEC usa `K7` como alias hasta el veredicto T0 |

**Aportado por la sesión (NO re-verificado en disco en esta planificación; T0 lo re-verifica antes de implementar — ver TASKS T0):** salida agosto con 1003 entradas calcChain inconsistentes (julio Δ=0 → 0 inconsistencias; agosto Δ=[-3,-9,-6,+6,+8]); `xl/workbook.xml` con `<calcPr calcId="191029"/>` sin `fullCalcOnLoad`; calcChain byte-idéntico a la golden julio (12921 entradas vs 12920 de la plantilla en ceros); G7/K7 = 46219/46234 (16/07–31/07) como valores estáticos con propagación por fórmula; fechas R10 agosto = 16–31/08/2026.

### 0.2 Mapeo al Rector (in vs out)

**Entra:**

| Requisito | Superficie de cambio |
|---|---|
| Unidad S: eliminar `CalculationChainPart` al guardar + `FullCalculationOnLoad=true` en `CalculationProperties`, centralizado en punto único de guardado (D-A/D-B) | `OpenXmlPlantillaWriter` (helper interno, invocado en los 2 `Save()` + camino espejo) o `OpenXmlEspejoR1Mutador` según veredicto T0; sin tocar `<f>` ni `<v>` |
| Unidad S: test nuevo de integridad estructural (sin calcChain o calcChain↔fórmulas consistente + `fullCalcOnLoad="1"`) | `Remuneracion.IntegrationTests` (xUnit + FluentAssertions), con insumos reales |
| Unidad F: escribir FechaDesde→G7 y FechaHasta→K7/J7 de `CONSOLIDADO_TOTAL RECAUDO` desde `DetRetriInputs`, SOLO valores sobre celdas estáticas, con fail-fast si el R10 no trae fechas legibles o si G7/K7 son fórmula (D-C/D-D/D-E) | Writer (1 método de sello, tras lectura R10) + `ProcesadorPeriodo` (paso del `r10` ya leído); guard anti-fórmula existente como red |
| Trazabilidad doctrinal mínima | Diff acotado en `.opencode/project-context.md` (+ manual de usuario: correr SIEMPRE desde la plantilla en ceros) |
| Tests con insumos reales; goldens julio ±0.5 intactos; DetRetri 5/5 vs R10 intacto; salida agosto regenerada desde ceros sin reparación y con G7/K7 = 16/08–31/08/2026 | `Remuneracion.IntegrationTests`; suite 339/339 como red ciega |

**Sale (EXPLÍCITO):** tocar `<f>` o `<v>` (prohibido: goldens ±0.5); regenerar/recalcular la calcChain (se elimina o se deja recalcular a Excel, nunca se reescribe a mano); cambios a finders, readers (salvo pasar `r10` ya leído), cálculo, validaciones o R10-oráculo; reinterpretar valores de agosto; duplicar C59 (ya gobernado por dominio, §V5); bloqueo automático ante "plantilla = salida trabajada" sin evidencia T0 (D-G: solo documentación + evaluación de guardrail barato); paquetes NuGet; commits (los hace el Ingeniero con `#commit`).

### 0.3 Decisiones ejecutivas (aprobación = aceptar estas)

| ID | Decisión |
|---|---|
| D-A | **Unidad S, Alternativa A YA APROBADA en sesión:** al guardar, eliminar `CalculationChainPart` y setear `FullCalculationOnLoad=true` en `CalculationProperties`. No se toca ninguna fórmula ni valor cacheado (`<f>`/`<v>` intactos; goldens ±0.5 por construcción). |
| D-B | **Punto único de guardado:** la limpieza S vive en UN helper invocado por TODOS los caminos de guardado (los 2 `Save()` de `GenerarWorkbook` + el flujo espejo — inventario exacto re-verificado en T0). Ningún camino de guardado queda sin la limpieza; si T0 encuentra un tercer camino (tests/herramientas fuera del writer), se declara y se cubre o se excluye con justificación. |
| D-C | **Unidad F (a ratificación):** escribir FechaDesde→G7 y FechaHasta→K7/J7 (columna exacta según T0, §V10) de `CONSOLIDADO_TOTAL RECAUDO` desde `DetRetriInputs` (fechas del R10 ya leídas en `ProcesadorPeriodo` L300). SOLO valores, en la MISMA pasada atómica, con el guard anti-fórmula existente. |
| D-D | **G7/K7 estáticas o fail-fast:** T0 confirma que G7/K7 son valores estáticos en la plantilla en ceros. Si alguna fuera fórmula → el sello NO se escribe y el writer lanza `ERR-PLANTILLA` nombrando hoja+celda (nunca sobrescritura silenciosa; invariante NUNCA-fórmulas intacta). |
| D-E | **R10 sin fechas legibles → fail-fast:** si `FechaDesde/FechaHasta` llegan en `default` (el lector devuelve `default` cuando no encuentra/no parsea, L88-89), el período aborta con `CalculoInvalidoException` nombrando período + archivo R10 + celda (nada de fechas silenciosas de otro período). |
| D-F | **Alcance del sello (único fork):** base = SOLO G7/K7. Si T0 encuentra OTROS rótulos estáticos de fecha/período sin sellar, se sellan SOLO si son valores estáticos del mismo origen (R10) y sin fórmulas; C59 queda documentado como ya gobernado (dominio, §V5) y no se duplica. Recomendación: base + lo que T0 evidencie, sin expandir sin justificar. |
| D-G | **Plantilla en ceros como guía de proceso (recomendación, no bloqueo):** documentar en `project-context.md` + manual que SIEMPRE se corre desde la plantilla en ceros; T0 evalúa con evidencia si un guardrail barato es viable (p. ej. advertencia/log ante plantilla sospechosa) o si queda como non-goal documentado. NO expandir alcance sin justificar. |

---

## 1. PROPOSE

### 1.1 Intent

Que el workbook de salida abra en Excel SIN diálogo de reparación (cadena de cálculo saneada en todo camino de guardado, sin tocar una sola fórmula) y que cada salida quede sellada con las fechas de SU período (G7/K7 desde el R10, nunca arrastradas de la plantilla), con la plantilla en ceros como base agnóstica de período documentada en el proceso.

### 1.2 In Scope

- Unidad S: helper de saneamiento + invocación en todos los caminos de guardado + test de integridad estructural.
- Unidad F: sello G7/K7 desde `DetRetriInputs` + fail-fast por fechas ilegibles o celdas-fórmula + verificación de propagación por fórmula tras recálculo.
- T0 obligatorio (§4) que re-verifica en disco (a)..(e) del encargo antes de implementar.
- Tests con insumos reales + salida agosto regenerada desde ceros (sin reparación estructural + G7/K7 = 16/08–31/08/2026).
- Docs: `project-context.md` (doctrina calcChain + sello + plantilla-en-ceros) + manual de usuario.

### 1.3 Non-Goals

Ver §0.2 "Sale". En particular: NO se reconstruye la calcChain a mano; NO se valida contenido de fechas más allá de legibilidad (el R10-oráculo sigue validando DetRetri como hoy); NO se cambia ningún mensaje existente salvo los fail-fast nuevos D-D/D-E; NO se toca C59; NO hay bloqueo automático por "salida trabajada como plantilla" salvo veredicto T0/D-G.

### 1.4 Riesgos

| ID | Riesgo | Mitigación |
|---|---|---|
| R-CALC | Eliminar la cadena de cálculo degrada la apertura (Excel recalcula todo al abrir) | Es el comportamiento diseñado de `fullCalcOnLoad`: Excel recalcula una vez con fórmulas intactas; goldens ±0.5 y DetRetri 5/5 como red aritmética |
| R-CAMINO | Un camino de guardado queda sin limpieza y el bug reaparece en otro flujo (single-ASE, espejo standalone, herramientas) | D-B + inventario T0 + test estructural sobre salidas de AMBOS overloads |
| R-FECHA-SILENCIOSA | Fechas de otro período selladas sin que nadie lo note (el riesgo que motiva F) | D-E (fail-fast sin fechas) + D-D (fail-fast si fórmula) + test G7/K7 por período |
| R-PLANTILLA | La "plantilla en ceros" trae ella misma fechas/la calcChain de otro período | T0 la inspecciona (bytes + G7/K7 + calcChain base); la guía de proceso la declara canónica con evidencia |
| R-SOBREDISENO | Convertir el sello en framework de metadatos del período | Alcance congelado D-F; cualquier rótulo extra requiere evidencia T0 + mismo origen R10 |
| R-FALSO-GUARDRAIL | Un detector de "salida trabajada" bloquea plantillas legítimas | D-G: solo con evidencia T0; por defecto documentación, no bloqueo |

---

## 2. DESIGN

### 2.1 Enfoque: sanear al guardar, sellar al generar

**Unidad S (Alternativa A).** Al guardar el workbook de salida, ANTES del `Save()` final de cada camino:

1. `workbookPart.DeletePart(workbookPart.CalculationChainPart)` si existe (o el equivalente OpenXML 3.5.1 del proyecto);
2. `workbookPart.Workbook.CalculationProperties.FullCalculationOnLoad = true` (creando `CalculationProperties` si la plantilla no lo trae; `calcId` se conserva).

Dónde: UN helper interno (p. ej. `SanearCadenaCalculo(WorkbookPart)`, nombre a elección) invocado en los 2 `Save()` de `GenerarWorkbook` (L147/L287) y en el camino espejo (`Ajustar`/`AjustarEnWorkbook` si persiste parte propia al guardar — veredicto T0). Sin tocar `<f>` ni `<v>`: los valores cacheados quedan y Excel los recalcula al abrir por `fullCalcOnLoad`.

**Test de integridad estructural (el blindaje).** Sobre workbooks generados en temp con insumos reales:

- (a) `CalculationChainPart` ausente — o, si el diseño final decide conservarlo en algún camino, consistencia calcChain↔fórmulas con mapeo por `sheetId` (T0 + implementación fijan cuál de las dos ramas, con evidencia; recomendación: ausencia, la rama simple de la Alternativa A);
- (b) `Workbook.CalculationProperties.FullCalculationOnLoad == true` (`fullCalcOnLoad="1"` en el XML).

**Unidad F (sello de fechas).** En `ProcesadorPeriodo`, tras leer el R10 (L300, `r10` ya en memoria) y antes de `GenerarWorkbook` (L330): el writer escribe `r10.FechaDesde → CONSOLIDADO_TOTAL RECAUDO!G7` y `r10.FechaHasta → K7/J7` (columna según T0) como valores fecha (serial OADate, estilo de fecha preservado de la plantilla — sin crear estilos nuevos si es evitable), vía `EscribirValorNumerico` (guard anti-fórmula §V6 intacto: si G7/K7 es fórmula → `ERR-PLANTILLA`). Si alguna fecha es `default` → `CalculoInvalidoException` nombrando período + archivo R10 + celda ("Fecha Desde"/"Fecha Hasta" del R10). La propagación a las demás hojas ocurre por las fórmulas existentes al recalcular con `fullCalcOnLoad` (T0 verifica el grafo de referencias G7/K7; el plan NO agrega fórmulas).

**Discrepancia G7/K7 vs G7/J7 (§V10):** T0 abre la plantilla en ceros y el R10 y fija las columnas reales de "Fecha Desde"/"Fecha Hasta" en AMBOS archivos. La implementación usa esas columnas; este plan usa `K7` como alias hasta el veredicto.

### 2.2 Alternativas y trade-offs

| Eje | Elegida | Descartada |
|---|---|---|
| CalcChain: eliminar + `fullCalcOnLoad` (A, aprobada) | Simple, sin tocar fórmulas/valores; Excel recalcula al abrir; goldens intactos | Reconstruir la calcChain consistente a mano: frágil (1003 entradas, mapeo `sheetId`, reanclaje espejo la invalida de nuevo en cada Δ); conservar calcChain stale: reproduce el bug |
| Dónde sanear | Helper único invocado en cada `Save()` (D-B): un solo sitio, todo camino cubierto | Sanear solo en el multi-ASE: el single-ASE y el espejo standalone quedarían con el bug latente |
| Sello: 2 celdas desde R10 ya leído | Origen único ya validado (R10 oráculo G3); sin I/O nuevo; sin plumbing UI/CLI | Leer fechas del nombre de carpeta/período o pedirlas por UI: duplica fuentes de verdad; el R10 ya las trae y ya se lee |
| Formato de fecha | Serial OADate con estilo existente de la celda | Texto "16/08/2026": rompería las fórmulas que propagan G7/K7 si esperan fecha serial |
| Detección "salida trabajada" | D-G: documentar + evaluar guardrail barato en T0 (por defecto non-goal) | Bloqueo automático sin evidencia: riesgo R-FALSO-GUARDRAIL (falsos positivos sobre plantillas legítimas) |

### 2.3 Por qué NO hay desplazamiento, reanclaje ni riesgo geométrico (tratamiento explícito)

Unidad S no mueve filas ni toca referencias: elimina una parte (`calcChain.xml`) y setea un flag (`calcPr`). Unidad F escribe 2 celdas de valor fuera de todo bloque espejo, sin insert/delete, sin reanclaje. La prueba dura: el diff no toca `OpenXmlEspejoR1Mutador`, ningún mapa de celdas existente, ninguna fórmula; solo agrega el helper de saneamiento, el método de sello (2 celdas) y su llamada. Los goldens Q1/Q2 no pueden moverse por construcción (ningún valor de negocio cambia; las fechas G7/K7 de julio selladas desde su propio R10 son idénticas a las de la plantilla: 16/07–31/07).

### 2.4 Contratos afectados (resumen)

| Contrato | Cambio |
|---|---|
| `OpenXmlPlantillaWriter` (nuevo helper interno `SanearCadenaCalculo` + método de sello `EscribirFechasPeriodo`) | Creación + invocación en cada `Save()`; G7/K7 vía `EscribirValorNumerico` (guard intacto) |
| `OpenXmlEspejoR1Mutador` (si T0 confirma camino propio de guardado) | Invocación al mismo helper; sin cambios de mutación/reanclaje |
| `ProcesadorPeriodo.Ejecutar` | Pasa `r10.FechaDesde/FechaHasta` (ya leídos) al writer; fail-fast D-E si `default` |
| `ProcesadorRemuneracion`, `CatalogoErrores`, `Form1`, `EjecutorCli`, finders, readers, cálculo, `IValidador`, R10-oráculo | Sin cambios (cero plumbing UI/CLI) |
| `Remuneracion.IntegrationTests` | Tests nuevos S + F con insumos reales; suite 339/339 como red ciega |
| `.opencode/project-context.md` + `Docs/Manual-Usuario-Remuneracion-UAESP.md` | Doctrina calcChain + sello + plantilla-en-ceros (diff acotado) |

### 2.5 Architecture Validation Certificate

| Principio | Veredicto |
|---|---|
| **SRP** | ✅ Una sola razón de cambio por costura: el helper sanea la cadena, el método de sello escribe 2 fechas, el procesador pasa el R10 ya leído, los tests demuestran. Sin ramas `if agosto` (la regla vale en todo período). |
| **OCP** | ✅ El próximo rótulo estático de período se absorbe por parámetro (celda + valor), no bifurcando el writer; la cardinalidad futura de bloques espejo no afecta al saneamiento (la ausencia de calcChain es independiente de Δ). |
| **DIP** | ✅ Cambios detrás de `IWorkbookLeafWriter`/`IEspejoR1Writer` y del servicio de R10 existente; UI/CLI no conocen la regla. |
| **Best practices** | ✅ Fail-fast con período+archivo+celda nombrados (D-D/D-E); quincena por dominio (C59 intacto); fórmulas nunca tocadas (guard como red, no como adorno); tolerancia ±0.5 y goldens intactos por construcción; mensaje para personas, detalle técnico al log. |
| **Performance** | ✅ Saneamiento O(1) por guardado (borrar una parte + setear un flag); sello = 2 celdas; el `fullCalcOnLoad` traslada el recálculo a la apertura en Excel (una vez), no al pipeline. |

---

## 3. SPEC

### 3.1 Requirements

| ID | Requisito | Grupo | Criterio de aceptación |
|---|---|---|---|
| R-S-1 | Al guardar, el workbook sale SIN `CalculationChainPart` y CON `fullCalcOnLoad="1"`, en TODOS los caminos de guardado (inventario T0, D-B) | Unidad S (aprobada) | Test estructural verde sobre salidas single-ASE y 5-ASE con insumos reales |
| R-S-2 | Ninguna fórmula ni valor cacheado se toca (`<f>`/`<v>` intactos) | Unidad S | Goldens Q1+Q2 ±0.5 intactos; DetRetri 5/5 vs R10 intacto |
| R-S-3 | Salida de agosto regenerada desde la plantilla en ceros: verificación estructural sin reparación (calcChain ausente/consistente + `fullCalcOnLoad`; sin Excel en el pipeline) | Unidad S | Test dedicado; la salida que hoy dispara el diálogo queda sana |
| R-F-1 | `FechaDesde→G7`, `FechaHasta→K7/J7` (columna T0) de `CONSOLIDADO_TOTAL RECAUDO` desde `DetRetriInputs`, SOLO valores, en la misma pasada atómica | Unidad F | Salida agosto con G7/K7 = 16/08–31/08/2026 como seriales fecha; julio intacto (16/07–31/07) |
| R-F-2 | G7/K7 fórmula → `ERR-PLANTILLA` nombrando hoja+celda, sin escritura (D-D) | Unidad F | Test con celda-fórmula en copia temp (insumo real copiado, nunca fixture en repo) |
| R-F-3 | R10 sin fechas legibles → fail-fast nombrando período + archivo + celda (D-E) | Unidad F | Test con R10 copiado a temp sin fechas (o fechas removidas en la copia) |
| R-F-4 | Propagación verificada: las fórmulas que referencian G7/K7 se resuelven tras recálculo (T0 inventaría el grafo; el test aserta las referencias, no el recálculo) | Unidad F | T0 documenta el grafo; sin fórmulas nuevas agregadas |
| R-F-5 | Otros rótulos estáticos de fecha/período (si T0 los encuentra): sellados solo con evidencia + mismo origen, o declarados fuera con justificación; C59 documentado, no duplicado (D-F) | Unidad F | Checklist T0 (d) cerrado sin ítems ambiguos |
| R-R-1 | Suite base 339/339 verde antes y después (`dotnet build` 0 warnings) | Transversal | Build + suite completos en cada WU |
| R-R-2 | Goldens Q1+Q2 ±0.5 intactos; DetRetri 5/5 vs R10 intacto (julio y agosto) | Transversal | Capa A + `Regresion2026082Tests` verdes |
| R-R-3 | Guía de proceso documentada: SIEMPRE desde la plantilla en ceros (`project-context.md` + manual); veredicto T0/D-G sobre guardrail barato | Transversal | Diff acotado a doctrina/reglas; sin código de bloqueo salvo veredicto |
| R-R-4 | Q1 intacto (`AjustesSfT = 0`; sello con fechas Q1 01–15/07 idénticas a plantilla) | Transversal | Golden Q1 ±0.5 verde |

### 3.2 Scenarios (Given/When/Then)

- **S1 (agosto sin reparación, Unidad S):** Given `Docs/Prueba2/Insumos` + plantilla en ceros, When `ProcesadorPeriodo` Q2 genera la salida, Then el xlsx resultante NO contiene `xl/calcChain.xml` (o la trae consistente) y `xl/workbook.xml` trae `fullCalcOnLoad="1"`; verificación por inspección ZIP, sin Excel.
- **S2 (julio intacto, Unidad S):** Given períodos Q1/Q2 julio, When se regeneran, Then test estructural verde + goldens ±0.5 (Δ=0: la limpieza es no-op observable salvo el flag).
- **S3 (sello agosto, Unidad F):** Given R10-2026082 con fechas 16–31/08, When flujo 5-ASE, Then `CONSOLIDADO_TOTAL RECAUDO` G7/K7 = 16/08/2026 y 31/08/2026 (seriales), DetRetri 5/5 vs R10 intacto.
- **S4 (sello julio, Unidad F):** Given R10-2026072 (16–31/07), When flujo Q2, Then G7/K7 = 46219/46234 y golden ±0.5 (sello idempotente respecto de la plantilla).
- **S5 (R10 sin fechas, D-E):** Given copia temp del período con R10 sin fechas legibles, When `Ejecutar`, Then fail-fast previo a escribir que nombra período + archivo + celda; ninguna salida creada.
- **S6 (G7 fórmula, D-D):** Given copia temp de la plantilla con G7 como fórmula, When `GenerarWorkbook`, Then `ERR-PLANTILLA` nombrando hoja+celda; plantilla original intacta.
- **S7 (cero-geometría):** Given el diff, When se audita, Then ningún mapa, ninguna fórmula y ningún valor de negocio cambió salvo el helper S + 2 celdas F.
- **S8 (guía de proceso):** Given el manual actualizado, When un operador prepara un período nuevo, Then encuentra la instrucción "correr SIEMPRE desde la plantilla en ceros" + qué riesgo corre si usa una salida trabajada.

---

## 4. TASKS

| ID | Tarea | Grupo | Depende de | Verificación |
|---|---|---|---|---|
| T0 | **Evidencia T0 OBLIGATORIA (Fase 0, NO es PR):** re-verificar en disco (a) inconsistencias calcChain de la salida actual de agosto (conteo calcChain↔fórmulas, `calcId`, ausencia de `fullCalcOnLoad`; julio como control Δ=0); (b) inventario exacto de caminos de guardado con `Save()` (writer L147/L287 + flujo espejo + terceros si los hay); (c) G7/K7 estáticos en la plantilla en ceros + grafo de propagación por fórmula + columna real de "Fecha Hasta" (K7 vs J7, §V10); (d) barrido de OTROS rótulos estáticos de fecha/quincena sin sellar (gobierno actual de cada uno; C59 por dominio); (e) fechas legibles en los 3 R10 (2026071: 01–15/07, 2026072: 16–31/07, 2026082: 16–31/08). Veredicto por rama S/F + fijación de columnas + D-G con evidencia | Fase 0 | — | §0.1 (a)..(e) cerrados con bytes/celdas citados; sin NEEDS_CONTEXT o con recorte explícito |
| T1 | Unidad S: helper `SanearCadenaCalculo` + invocación en TODOS los caminos T0 (D-A/D-B); `<f>`/`<v>` intactos | Fix-S | T0 | R-S-1/R-S-2; build 0 warnings |
| T2 | Unidad S: test de integridad estructural (a)+(b) sobre salidas single-ASE y 5-ASE con insumos reales + salida agosto regenerada desde ceros sin reparación (R-S-3, S1/S2) | Fix-S | T1 | SPEC S completa; suite ≥339/339 |
| T3 | Unidad F: sello G7/K7 desde `DetRetriInputs` (columnas T0) + fail-fast D-D/D-E + paso de `r10` en `ProcesadorPeriodo`; C59 intacto | Fix-F | T0 | R-F-1/R-F-2/R-F-3; S3/S4/S5/S6; build 0 warnings |
| T4 | Tests F con insumos reales (propagación S4, Q1 R-R-4, checklist R-F-5) + regresión DetRetri 5/5 vs R10 intacta (julio y agosto) | Transversal | T3 | R-F-4/R-F-5 + R-R-2; `dotnet build Remuneracion.WinForms/Remuneracion.WinForms.slnx` 0 warnings; suite ≥339/339 |
| T5 | Docs: `project-context.md` (doctrina calcChain + sello + plantilla-en-ceros) + manual de usuario (guía SIEMPRE-desde-ceros + veredicto D-G) | Transversal | T2, T4 | R-R-3; diff acotado; S7/S8 citados |

**Orden sugerido:** T0 (bloqueante) → T1 → T2 → T3 → T4 → T5. T1+T2 pueden shippear solas (WU-1, Unidad S aprobada); T3+T4+T5 en WU-2 (Unidad F ratificada). Build 0 warnings antes de marcar cada tarea completa. Sin commits (los hace el Ingeniero).

---

## 5. Resumen ejecutivo, riesgos y decisiones

- **Resumen:** dos defectos de distinta capa con una sola corrida de evidencia. Unidad S (aprobada): el mutador espejo mueve filas pero la calcChain se preserva byte-idéntica → Excel repara al abrir; la Alternativa A (eliminar cadena + `fullCalcOnLoad`, sin tocar fórmulas) lo cierra en un punto único de guardado, con test estructural como blindaje del hueco. Unidad F: agosto se generó sobre la salida trabajada de julio porque no había otra plantilla; sellando G7/K7 desde el R10 (ya leído, nunca escrito) la plantilla en ceros queda agnóstica de período y reutilizable, con fail-fast ante fechas ausentes o celdas-fórmula. Con julio el flujo es bit-idéntico (fechas idénticas a la plantilla, Δ=0).
- **Riesgo principal:** R-CAMINO (un camino sin limpieza reabre el bug) + R-FECHA-SILENCIOSA (fechas de otro período) + R-FALSO-GUARDRAIL (bloqueo de plantillas legítimas). Contenidos por D-B + inventario T0 + test estructural, D-D/D-E + tests S5/S6, y D-G (documentar por defecto).
- **Decisión para el Ingeniero:** ratificar D-A..D-G (§0.3). D-A ya aprobada en sesión (WU-1 shippea sin esperar a F). **Preguntas con recomendación (ninguna bloqueante):** (1) alcance del sello D-F — recomendación: base G7/K7 + lo que T0 evidencie con mismo origen, sin expandir sin justificar; (2) guardrail "salida trabajada como plantilla" D-G — recomendación: documentación + evaluación T0, sin bloqueo salvo evidencia de un heurístico sin falsos positivos; (3) columna real de "Fecha Hasta" (§V10) — la fija T0, el plan usa `K7` como alias.

### DoD (Definition of Done)

1. `dotnet build Remuneracion.WinForms/Remuneracion.WinForms.slnx` con 0 warnings.
2. Suite completa verde sin regresión (base 339/339) + goldens Capa A Q1+Q2 en dif ±0.5 intactos.
3. Test de integridad estructural verde (S1/S2): calcChain ausente/consistente + `fullCalcOnLoad="1"` en salidas single-ASE y 5-ASE.
4. Salida de agosto regenerada desde la plantilla en ceros: verificación estructural sin reparación + G7/K7 = 16/08–31/08/2026 (S1/S3) + DetRetri 5/5 vs R10 intacto.
5. Fail-fast nuevos verdes (S5/S6): sin fechas → período+archivo+celda; G7/K7 fórmula → hoja+celda; ninguna fórmula sobrescrita (S7).
6. `project-context.md` + manual actualizados (guía plantilla-en-ceros, S8); sin datos inventados; sin commits del agente; CRLF; sin emojis.

### Follow-ups explícitos (fuera de esta HU)

- R4 espejo pendiente del Plan 21 (motor solo R1) y R4-por-empresa Q2 (recorte HU-20/G2-D1): intactos, no tocados.
- Validación de NOMBRES de hoja en preflight (follow-up del Plan 27): intacto.
- Si T0/D-G produce un heurístico viable de detección de "salida trabajada", su bloqueo (si lo hay) va en su propia HU con su propio T0.

---

## 6. Estado de ejecución (T0 + WU-1 + WU-2)

> Cierre 2026-10-05. Ejecutado por `implementer`. Sin commit.

### 6.1 T0 — Evidencia re-verificada en disco (parser ZIP+XML, sin Excel)

| Ítem | Veredicto T0 | Evidencia |
|---|---|---|
| (a) calcChain salida agosto | **CONFIRMADO 1003 inconsistencias** | `Docs/Prueba2/Resultado/Remuneración 202608-2 Total.xlsx`: 12921 entradas calcChain / 12921 fórmulas / **1003 inconsist** (mapeo `i`=sheetId); `<calcPr calcId="191029">` SIN `fullCalcOnLoad`. Controles: plantilla ceros 12920/12920=0; golden julio Q1 12683/0; golden julio Q2 12921/0. |
| (b) caminos de `Save()` | **3 de producción + 1 externo** | `OpenXmlPlantillaWriter.GenerarWorkbook` L147 (single) / L287 (multi); `OpenXmlEspejoR1Mutador.Ajustar` L85 (standalone, consumido por `EscribirEspejoR1`). Herramienta `Herramientas/VerificadorRecaudo` L450 = FUERA (diagnóstico). Tests guardan fixtures propios (fuera de producción). |
| (c) G7/K7 plantilla en ceros | **VALORES estáticos + columna real K7** | `F7`="Fecha Desde:"(ss27) `G7`=46219 estático (`s=9`, numFmt 14 fecha); `J7`="Feha Hasta:"(ss28, typo de la fuente) `K7`=46234 estático (`s=9`) → **"Fecha Hasta" REAL = K7** (J7 es rótulo). Propagación por fórmula confirmada: `AJUSTES - SF-T`, `REMUNERACION_*`, `DetRetri2026072`, `DetValiRetri2026072`, `GERENTES_*` referencian `'CONSOLIDADO_TOTAL RECAUDO'!G7/K7`. |
| (d) otros rótulos estáticos | **documentados, fuera de scope** | `C6`="Fecha de Proceso:"/`D6`=46238 (sello de generación, no rango) y `N3`=2026072 (código remuneración; también en nombres de hoja). `C59` (quincena) ya gobernado por dominio. Ninguno del mismo origen R10 que G7/K7 → no se sellan (D-F, sin sobre-diseño). |
| (e) fechas R10 | **legibles en los 3** | `R10_2026071`: G7=01/07/2026, J7=15/07/2026. `R10_2026072`: G7=16/07/2026, J7=31/07/2026. `R10_2026082`: **G7=16/08/2026, J7=31/08/2026**. |

**Sin desmentidos de premisa.** La única fijación fue la columna de "Fecha Hasta": el encargo nombró G7/K7 (correcto para el template) y el modelo/R10 documentaban G7/J7 (correcto para el R10). El sello escribe template `G7←r10.FechaDesde`, `K7←r10.FechaHasta`; el reader ya leía bien el R10 (G7/J7).

### 6.2 Unidades ejecutadas

- **Unidad S (D-A/D-B):** nuevo `SaneadorCadenaCalculo.Sanear(WorkbookPart)` (elimina `CalculationChainPart` si existe + `FullCalculationOnLoad=true` conservando `calcId`), invocado en los 3 caminos de guardado. `<f>`/`<v>` intactos.
- **Unidad F (D-C/D-D/D-E):** `ResultadoRemuneracion.FechaDesde/FechaHasta`; `ProcesadorPeriodo` valida fechas del R10 (fail-fast `ERR-FORMATO-FUENTE`) y las propaga; el writer sella `CONSOLIDADO_TOTAL RECAUDO!G7`/`!K7` (seriales OADate, solo valores, guard anti-fórmula intacto). C59 sin tocar; `N3`/`C6`/`D6` fuera de scope.
- **T5 (D-G):** doctrina en `.opencode/project-context.md` + manual (§2.3/§3.4/§10); **sin guardrail de bloqueo** (documentación, no falso positivo).

### 6.3 Tests y verificación (evidencia fresca)

- **Tests nuevos** (`Remuneracion.IntegrationTests/IntegridadCalcChainYSelloFechasTests.cs`, 6): agosto desde ceros (sin calcChain + `fullCalcOnLoad` + G7/K7 agosto), julio Q2 y Q1 (sello idempotente + estructura), **single-ASE** (estructura sana; no sella fechas, G7/K7 idénticas a plantilla), R10 sin fechas → fail-fast nombrando período+archivo+G7/J7 sin salida, G7 fórmula → `ERR-PLANTILLA` nombrando hoja+celda.
- **Build:** `dotnet build Remuneracion.WinForms/Remuneracion.WinForms.slnx` → **0 warnings, 0 errors**.
- **Suite:** `dotnet test Remuneracion.IntegrationTests/...` → **345/345** (339 base + 6 nuevos).
- **Salida agosto regenerada desde la plantilla en ceros** (CLI, temp con insumos reales Prueba2 + conciliaciones Q2), verificada por parser (ZIP+XML): `xl/calcChain.xml` **ausente**; `xl/workbook.xml` con **`fullCalcOnLoad="1"`** (`calcId=181029` preservado); `CONSOLIDADO_TOTAL RECAUDO!G7=46250` (16/08/2026) y `!K7=46265` (31/08/2026) con estilo fecha `s=9`; `DetRetri2026072!D9:D13` = **18378829331 / 21599709648 / 16369059896 / 8372092112 / 12137660178** = R10 `DetRetri2026082` D9:D13 exacto (D14 76857351165). Sin Excel en el pipeline.

### 6.4 Riesgos residuales

- R-CALC: Excel recalcula una vez al abrir (fullCalcOnLoad) con fórmulas intactas; goldens ±0.5 y DetRetri 5/5 como red (verdes).
- R-CAMINO: los 3 caminos de guardado quedan saneados; el helper es idempotente/tolerante.
- R-FECHA-SILENCIOSA: cerrado por D-E + D-D + test por período.
- Follow-ups intactos: R4 espejo (Plan 21), R4-por-empresa Q2, validación de nombres de hoja en preflight.
