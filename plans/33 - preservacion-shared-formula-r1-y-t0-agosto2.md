# Plan 33 — Preservación de shared-formula R1 (Julio-2) + T0 bloqueante Agosto-2

> **Alcance:** dos fases en orden estricto. **Fase 1 (referencia):** fix del shared-formula huérfano en `Reporte Componentes R1` que rompe la apertura en Excel en Julio-2 (2026072) — causa cerrada por debug-agent, no se re-diagnostica. **Fase 2 (secuenciada):** T0 de evidencia OBLIGATORIO para el fallo de Agosto-2 `ASE 2 Reporte Componentes R1!F199 [SUB_EMP]: faltan las anclas [Mes0] del sub-bloque` — este plan prescribe el T0 que decide el fix, NO prescribe el fix de agosto. Cambio de producción esperado: unas líneas en `Remuneracion.Infrastructure/Excel/OpenXmlEspejoR1Mutador.cs` (+ tests + gate). Cero cambios de semántica de negocio.
> **Rector (IRRENUNCIABLE):** `README.md` + `Docs/Propuesta_Proyecto_Automatizacion_Remuneracion_UAESP.md` + `Docs/Detalle de plantilla.docx` + `Docs/Proceso de Recaudo.docx` + `Docs/Prompt_Maestro_Proyecto_Remuneracion_UAESP _ Vo .docx` + `.opencode/project-context.md`. **Invariantes del proyecto:** tolerancia ±0.5 en comparaciones golden; NUNCA usar cachés `<v>` como oráculo; doctrina calcChain (Plan 28) intacta — no se toca `<f>` salvo lo declarado, no se toca `<v>` nunca; build 0 warnings antes de marcar cualquier tarea completa; finales de línea CRLF sin mezclar; NUNCA inventar recursos de UI (N/A aquí); NUNCA git commit/push (los hace el Ingeniero con `#commit`/`#push`); la implementación posterior la ejecuta `implementer` usando SOLO este plan aprobado; `code-reviewer` verifica spec compliance + 0 warnings. **Quincena por dominio, nunca por contenido.**
> **Continuidad:** Planes 21 (espejo R1), 25 (roles por firma), 26 (preflight), 27 (firma + hoja-por-nombre), 28 (calcChain + sello G7/K7), 29 (paridad app-vs-manual), 30 (naming dinámico + base canónica), 31 (recomposición visibles R1 por firma — `RecomponerVisibles`), 32 (recomposición interior R1 — `RecomponerInterior` + `CeldasInterioresAgosto` + `CompositorInteriorR1`) vigentes. Este plan NO reabre lectura, cálculo, espejo-dimensional, preflight, sellos, naming, calcChain ni paridad fuera de lo declarado: **julio = identidad por construcción; agosto-actual = rojo que se cierra.**
> **Insumos canónicos (no se re-fijan, no se inventan datos):** plantilla `Docs/Prueba Julio-2/Plantilla_Remuneracion.xlsx`; salidas sanas pre-commit `Resultado1/Resultado2` (commit previo a `73de7ea`); salida corrupta regenerada post-commit (fresca, a temp); R1 fuente ASE por período (`Recaudoporcomponente_*`); base canónica agosto `Docs/Prueba2/Plantilla_Remuneracion_2026082.xlsx` (Plan 30) para la Fase 2; manuales administrativos como oráculo de paridad. Reproducción determinista vía `Remuneracion.Cli`: `--periodo 2026072 --carpeta "Docs/Prueba Julio-2/Insumos" --plantilla "Docs/Prueba Julio-2/Plantilla_Remuneracion.xlsx"` hacia una ruta temp FRESCA (nunca sobre una salida existente).
> **Numeración:** `plans/` 01..32 ocupados; este plan toma el primer correlativo libre, **33**.
> **Estado:** APROBADO por el Ingeniero 2026-10-08 (con enmienda D-H cero-quema-de-períodos) — listo para `implementer` usando SOLO este documento.
> **Fecha:** 2026-10-08

---

## 0. Clarification Gate

**Sin preguntas bloqueantes para la Fase 1: la causa viene cerrada por debug-agent (§0.1) y el fix recomendado viene redactado en la ruta recomendada con su verificación de cierre dentro de la tarea. Para la Fase 2 hay un punto NO negociable: el T0 de evidencia (T4) es BLOQUEANTE — ningún fix de agosto se prescribe ni se implementa antes de su veredicto.** Punto de bloqueo respetado: NO se implementa nada en este documento (planificación read-only: cero cambios a código de producción en este documento).

### 0.1 Diagnóstico adoptado (causa cerrada, no re-diagnosticar)

Fuente: debug-agent (veredicto confirmado, se adopta tal cual como fuente única de verdad de la causa):

| # | Afirmación del diagnóstico | Reutilización en este plan | Re-verificación (el test, no la mano) |
|---|---|---|---|
| E1 | Síntoma: salida Julio-2 (2026072) reporta OK pero Excel abre con diálogo de reparación `Registros quitados: Formula compartida de /xl/worksheets/sheet10.xml`. `sheet10.xml` = `Reporte Componentes R1`. | Síntoma que el gate T1 debe reproducir en rojo con regen fresca y que T3 debe cerrar con apertura sana. | T1: regen fresca a temp → parse zip+XML muestra `si` huérfanos; T3: regen post-fix → invariante verde + apertura manual sana (aceptación final por el Ingeniero). |
| E2 | Causa raíz: `RecomponerVisibles` (Plan 31/T2) en `Remuneracion.Infrastructure/Excel/OpenXmlEspejoR1Mutador.cs` escribe fórmulas visibles con `celda.CellFormula = new CellFormula(texto)` (`EscribirFormulaVisible` ~L1514; mismo patrón `EscribirFormulaInterior` ~L1209). Esto DESCARTA `FormulaType=Shared` / `Reference` / `SharedIndex` cuando la celda destino es un MASTER de shared-formula en el template. | Causa raíz: no se re-investiga. El fix vive exactamente en esos dos puntos de escritura. | T1 congela el mapa shared esperado (fixture); T2 muta `.Text` in place y el gate lo demuestra. |
| E3 | Afectadas: `G53` (ASE1, `si=0` `ref=G53:AP53`) y `G468` (ASE4, `si=20` `ref=G468:AO468`); sus seguidoras (p. ej. `N53` `si=0`, `U468` `si=20`) quedan huérfanas. | Conjunto mínimo de asserts explícitos del gate: `G53` y `G468` con `ref`+`si` preservados + existencia de master por cada `si` con seguidoras. | T1: asserts `G53`/`G468` en rojo con regen corrupta; T2/T3: en verde. |
| E4 | Regresión del commit `73de7ea` (10/8): `Resultado1/2` (pre-commit) sanos con 1070 shared intactos; regen fresca post-commit muestra 1068 shared / 33 masters (exactamente los 2 masters perdidos, texto idéntico). | Delimitador de regresión: el diff de producción de este plan NO debe tocar semántica — solo preservación de atributos shared. Julio-identidad incluye texto idéntico + atributos intactos. | T1: conteos congelados como fixture (35 masters, rango `si`, ~1035 seguidoras — a congelar exacto desde el disco en T1 si el redondeo difiere; el fixture manda); T3: regen post-fix = texto idéntico + masters restaurados. |
| E5 | Repro determinista vía `Remuneracion.Cli`: `--periodo 2026072 --carpeta "Docs/Prueba Julio-2/Insumos" --plantilla "Docs/Prueba Julio-2/Plantilla_Remuneracion.xlsx"` a ruta temp FRESCA. | Procedimiento canónico de regen para T1 (rojo) y T3 (verde). Prohibido regenerar sobre salidas existentes. | T1/T3 usan ese comando + parse zip+XML BCL (sin Excel/COM en pipeline ni tests). |
| E6 | Fix recomendado: preserve-and-rewrite — cuando la celda ya trae `CellFormula`, mutar `.Text` in place en vez de `new CellFormula`, preservando tipo shared / `ref` / `si`; `Reanclar` ya corrige rangos `ref`. Alternativas: expand-to-plain (compleja, riesgosa), skip-when-delta-zero (enmascara, NO es causa raíz — rechazar o diferir explícitamente). | Decisión D-A de este plan. | T2 implementa preserve-and-rewrite; el plan documenta el rechazo de las alternativas (§2.2). |

**Nota de arrastre a agosto (E7, del encargo):** `G468` está ADEMÁS en `CeldasInterioresAgosto[4]` como `("G468","EXT_INT")` — el pase interior de agosto puede golpear MÁS celdas shared además de las 2 visibles de julio. Por eso el gate T1 incluye un caso agosto (workbook-wide, no solo `G53`/`G468`) y el fix T2 cubre ambos puntos de escritura (visible + interior).

### 0.2 Mapeo al Rector (Sale / No sale)

**Sale (EXPLÍCITO):**
- Re-diagnosticar la causa E1–E6 o re-abrir lectura, cálculo, espejo-dimensional, preflight, sellos, naming, calcChain o paridad fuera de lo declarado.
- Tocar cachés `<v>` como oráculo en cualquier gate o comparador (prohibido por doctrina Plan 29: stale por diseño; solo texto-`<f>` + atributos + literales).
- Tocar `<f>` de negocio fuera de la preservación de atributos declarada (el texto de fórmula NO cambia en este plan; solo se preservan `t`/`ref`/`si`).
- Expand-to-plain (expandir shared a fórmulas planas) o skip-when-delta-zero como fix (rechazadas en D-A; la segunda solo podría volver como optimización futura con su propio T0, nunca como fix).
- Prescribir el fix de Agosto-2 antes del veredicto T4 (la Fase 2 prescribe el T0, no el fix).
- Inventar el mapa shared esperado o las filas-ancla de agosto (todo se congela desde el disco: template + `Resultado1` + regen fresca).
- Paquetes NuGet; commits/push (los hace el Ingeniero con `#commit`/`#push`).

**No sale (entra, alcance congelado en 2 fases):**
| Fase | Superficie de cambio |
|---|---|
| F1 — gate + fixture shared | Nuevo gate workbook-wide de invariante shared (`sheet10.xml` = `Reporte Componentes R1`) + fixture del mapa shared esperado (35 masters, rango `si`, ~1035 seguidoras; asserts explícitos `G53`/`G468` ref+si) + tests TDD rojo-primero |
| F1 — fix preserve-and-rewrite | Unas líneas en `OpenXmlEspejoR1Mutador.cs` (`EscribirFormulaVisible` ~L1514 + `EscribirFormulaInterior` ~L1209): mutar `.Text` in place cuando la celda ya trae `CellFormula`; log auditado; `Reanclar` intacto |
| F1 — cierre | Regen fresca post-fix: invariante verde + goldens ±0.5 + suite verde + apertura manual sana en Excel (paso manual del Ingeniero como aceptación final) |
| F2 — T0 Agosto-2 (BLOQUEANTE) | Evidencia que arbitra 3 hipótesis para `ASE 2 Reporte Componentes R1!F199 [SUB_EMP]: faltan las anclas [Mes0] del sub-bloque`: (H1) nombre divergente en template `C199` vs fuente R1 ASE2, (H2) filas oportunas ausentes en la fuente, (H3) mapa `CeldasInterioresAgosto[2]` stale para 2026082 — anclado en los 3 .docx (§4/T4). Cero código de producción en T4. |
| F2 — fix agosto (CONDICIONAL) | Solo tras veredicto T4, con su propio diseño acotado; este plan NO lo prescribe. |
| D — docs | `project-context.md` (doctrina preservación-shared + invariante + lección Julio→Agosto) + manual (proceso: regen a temp fresca + cómo leer un FAIL del gate) |

**Fuera de scope explícito (ni Fase 1 ni Fase 2):** espejo-motor R4; cobertura `conditionalFormatting`/`dataValidations` (W-4 vivo); semántica DetRetri (redondeo/columna D intactos); UI/CLI (cero plumbing; la CLI solo se usa como runner de repro); validación de contenido en preflight (follow-up vivo); cualquier reescritura de `<f>` de negocio.

### 0.3 Decisiones ejecutivas (aprobación = aceptar estas)

| ID | Decisión |
|---|---|
| D-A | **Fix = preserve-and-rewrite (recomendado por debug-agent).** Cuando la celda destino ya trae `CellFormula`, mutar `.Text` in place preservando `FormulaType`/`Reference`/`SharedIndex`; solo crear `new CellFormula(texto)` cuando la celda NO traía fórmula. `Reanclar` sigue corrigiendo rangos `ref`. Descartado: expand-to-plain (reescritura masiva de seguidoras, riesgo de divergencia f-vs-f contra el manual y contra goldens) y skip-when-delta-zero (enmascara el defecto en julio y lo deja latente para cualquier delta futuro — NO es causa raíz). |
| D-B | **El fix cubre AMBOS puntos de escritura** (`EscribirFormulaVisible` ~L1514 y `EscribirFormulaInterior` ~L1209 — mismo patrón, mismo defecto). `G468` lo exige (E7: visible de julio + interior de agosto). El `EscribirCeroVisible`/`EscribirLiteralInterior` (literal 0 → `CellFormula = null`) NO se toca: un literal no es master shared y su semántica (Plan 31/32) queda intacta. |
| D-C | **Gate workbook-wide de invariante shared (nuevo, TDD rojo-primero).** Por cada `si` con seguidoras en `sheet10.xml`, existe un master con ese `si` en la misma hoja; además asserts explícitos `G53` (`si=0`, `ref=G53:AP53`) y `G468` (`si=20`, `ref=G468:AO468`) con texto idéntico al template. Rojo con regen fresca corrupta, verde con `Resultado1` (control sano) y con regen post-fix. Incluye caso agosto (E7: el interior puede golpear más celdas — el gate es workbook-wide por construcción, no lista cerrada de 2 celdas). |
| D-D | **Julio-identidad extendida a atributos.** Con geometría julio el pase reproduce el texto canónico (doctrina Plan 31 D-E) Y preserva los atributos shared (`t`/`ref`/`si`). Si el post-fix difiere en texto o en atributos, es regresión: se revierte el pase, no el golden. |
| D-E | **Doctrina calcChain intacta (Plan 28).** Este plan no toca `CalculationChainPart`, `calcPr` ni `fullCalcOnLoad`; no toca `<v>`; no toca `<f>` de negocio (solo preservación de atributos en las celdas del contrato de recomposición). El saneador `SaneadorCadenaCalculo.Sanear` y sus 3 caminos de invocación quedan intactos. |
| D-F | **Fase 2 = T0 bloqueante antes de cualquier fix de agosto.** El T0 arbitra H1/H2/H3 con evidencia en disco + cita a los 3 .docx de negocio; su veredicto define el diseño del fix (que sale como tarea condicional T5, no prescita aquí). Sin veredicto no hay T5. |
| D-G | **Lección Julio→Agosto como método.** Todo gate nuevo debe leer el WORKBOOK generado (texto-`<f>` + atributos + literales), no solo el dominio C# — la lección del Plan 31 (§lección) se extiende a atributos shared: el dominio nunca vio el `si` huérfano porque nunca leyó el workbook. |
| D-H | **Cero quema de períodos (enmienda aprobada por el Ingeniero 2026-10-08).** El diff de producción no contiene ningún literal ni rama por período (`2026072`, `2026082`, "julio", "agosto" ni equivalentes). Los períodos existen SOLO como datos de prueba (fixtures, goldens, comandos de repro). Criterio de aceptación: `grep` de literales de período en el diff de producción = 0 y el gate workbook-wide pasa sin conocer el período. La generalización del interior por-período (sucesor de `CeldasInterioresAgosto`) queda como follow-up declarado, no como deuda silenciosa. |

---

## 1. PROPOSE

### 1.1 Intent

Que la salida Julio-2 abra en Excel sin diálogo de reparación (masters shared preservados byte-idénticos en atributos, texto idéntico), que ningún `si` huérfano vuelva a entregarse en silencio (gate workbook-wide que hoy fallaría en rojo), y que el fallo de Agosto-2 (`F199 [SUB_EMP]` sin ancla `Mes0`) quede arbitrado por evidencia (T0 bloqueante) antes de prescribir su fix — con julio byte-idéntico en texto y atributos.

### 1.2 In Scope

- Gate de invariante shared + fixture del mapa esperado + tests TDD rojo-primero (Fase 1).
- Fix preserve-and-rewrite en los 2 puntos de escritura + log auditado (Fase 1).
- Cierre Fase 1: regen fresca verde + goldens ±0.5 + suite verde + apertura manual sana (paso del Ingeniero).
- T0 de evidencia Agosto-2 con veredicto H1/H2/H3 + cita a negocio (Fase 2, bloqueante, cero código).
- Docs mínimos: doctrina preservación-shared + invariante + lección (§D-G) en `project-context.md` + manual (proceso).

### 1.3 Non-Goals

Ver §0.2 "Sale" + fuera de scope explícito. En particular: NO se prescribe el fix de agosto (solo su T0); NO se generaliza la preservación a otras hojas (el patrón solo existe en R1; si otra hoja lo necesitara, su propio T0); NO se toca semántica numérica (el texto de fórmula no cambia: los valores post-recálculo son idénticos por construcción).

### 1.4 Riesgos

| ID | Riesgo | Mitigación |
|---|---|---|
| R-SHARED-OCULTO | Hay MÁS masters afectados que `G53`/`G468` (E7 advierte el interior de agosto) y el fix de 2 celdas deja huérfanos vivos | D-C: gate workbook-wide (todo `si` con seguidoras), no lista cerrada; T1 incluye caso agosto; D-B cubre ambos puntos de escritura |
| R-TEXTO-DIVERGE | Al mutar `.Text` in place se altera el texto además de preservar atributos y se rompe f-vs-f contra el manual | D-D: julio-identidad extendida (texto idéntico exigido); fixture congela textos; comparador f-vs-f + goldens ±0.5 como red por tarea |
| R-FALSO-VERDE | El gate pasa con cachés `<v>` stale o sin resolver `ref` y entrega huérfanos en silencio | Prohibido `<v>`; el gate lee atributos `t`/`ref`/`si` + texto-`<f>` del XML; fail-fast con hoja+celda+`si` si un `si` no tiene master |
| R-JULIO | El pase altera julio (texto o atributos) y rompe goldens | D-D + goldens Q1+Q2 ±0.5 + 0-diff R1 (texto + atributos) por tarea; ante diff se revierte el pase |
| R-AGOSTO-PRESUROSO | Se prescribe el fix de agosto sin evidencia y se fija lo equivocado (mapa vs fuente vs template) | D-F: T4 bloqueante con 3 hipótesis arbitradas; T5 condicional al veredicto, sin diseño prescrito aquí |
| R-CALCCHAIN | El fix interfiere con la doctrina Plan 28 (cadena de cálculo / `fullCalcOnLoad`) | D-E: el diff no toca saneador ni flags; test estructural existente (`IntegridadCalcChainYSelloFechasTests`) como red ciega en T3/T7 |
| R-REGEN-SUCIA | La evidencia se contamina regenerando sobre salidas existentes en vez de temp fresca | Procedimiento canónico E5 (temp FRESCA) exigido en T1/T3/T4; regen sobre salida existente = evidencia inválida |

---

## 2. DESIGN

### 2.1 Enfoque: preservar, no reescribir; gatear el workbook, no el dominio; evidenciar agosto, no prescribirlo

**Fase 1 — gate (T1).** Nuevo gate de post-escritura en `Remuneracion.IntegrationTests` (BCL `System.IO.Compression` + `System.Xml`, sin Excel/COM — mismo método que el comparador Plan 29 y los T0 de Planes 31/32): abre el xlsx generado, localiza `sheet10.xml` (`Reporte Componentes R1` — a congelar la resolución hoja↔sheetId en T1 desde el disco, no a asumir), y verifica (a) invariante workbook-wide: todo `si` que aparece en seguidoras (`t="shared"` con `si` y sin `ref`) tiene un master (`t="shared"` con ese `si` y con `ref`) en la misma hoja; (b) asserts explícitos: `G53` master `si=0` `ref=G53:AP53` + `G468` master `si=20` `ref=G468:AO468` con texto idéntico al template; (c) fixture del mapa esperado: 35 masters, rango `si`, ~1035 seguidoras (conteos del diagnóstico E4 — T1 los congela exactos desde `Resultado1` + template; si difieren en ±1 por redondeo de conteo, el fixture manda y E4 se enmienda). El gate corre sobre TRES corpus: regen fresca corrupta (ROJO esperado), `Resultado1` pre-commit (VERDE esperado — control sano), y regen post-fix (VERDE esperado). Incluye caso agosto (E7): el mismo invariante sobre la salida agosto actual — si el interior golpeó más celdas, el gate lo delata en rojo antes del fix y lo cierra después.

**Fase 1 — fix (T2).** En `EscribirFormulaVisible` (~L1514) y `EscribirFormulaInterior` (~L1209): si `celda.CellFormula` ya existe, mutar `celda.CellFormula.Text = texto` (preserva `FormulaType`, `Reference`, `SharedIndex` del template); si no existe, `celda.CellFormula = new CellFormula(texto)` (comportamiento actual para celdas literales que nacen con fórmula, clase L-2 del Plan 32). `Reanclar` (que ya corrige rangos `ref`) intacto; `EscribirCeroVisible`/`EscribirLiteralInterior` intactos (literal 0 no es master). Log Debug existente extendido con `t`/`ref`/`si` preservados (ASE+celda+etiqueta+texto+atributos). Diff esperado: unas líneas + tests + gate.

**Fase 2 — T0 (T4, BLOQUEANTE).** Evidencia en disco (mismo método zip+XML BCL) que arbitra para `ASE 2 Reporte Componentes R1!F199 [SUB_EMP]` (`faltan las anclas [Mes0]`): (H1) ¿la celda template `C199` (empresa del sub-bloque) trae un nombre divergente del que la fuente R1 ASE2 trae en sus filas (p. ej. `NUEVO ESQUEMA` vs `ENEL/OCCIDENTE` — el disco del Plan 32 ya trae `NUEVO ESQUEMA` en C, y `Proceso de Recaudo.docx` fija qué EFC factura a cada ASE: ASE2 = LIME con ENEL+EAAB)?; (H2) ¿la fuente `Recaudoporcomponente` ASE2-agosto trae CERO filas oportunas del sub-bloque que `F199` resume (recorte real — la doctrina espejo Plan 21 dice ausente = se suprime, pero el subtotal sin datos no puede componerse)?; (H3) ¿el mapa `CeldasInterioresAgosto[2]` está stale para 2026082 (la celda o su etiqueta no corresponden a la geometría real del template de agosto)? Anclaje en negocio: `Prompt_Maestro §5–8` (proceso secuencial: lectura directa de la fuente, insert/delete preservando fórmulas, `Reporte Componentes R1` desde `Recaudoporcomponente`), `Detalle de plantilla` (instrucción 5: si el nº de filas cambia, insertar/eliminar preservando fórmulas; `Reporte Componentes R1` por empresa facturadora), `Proceso de Recaudo` (matriz EFC×ASE: qué empresas pueden aparecer por ASE — arbitra si un nombre en C es legítimo o divergente). Veredicto por hipótesis (confirmada/descartada) + diseño acotado del fix resultante (que se ejecuta como T5, fuera del diseño de este plan).

### 2.2 Dónde vive el fix y trade-offs

| Eje | Elegida | Descartada |
|---|---|---|
| Qué se muta | `.Text` in place cuando `CellFormula` existe (D-A: preserva `t`/`ref`/`si`; `Reanclar` ya fija rangos) | `new CellFormula(texto)` siempre (el bug actual: descarta atributos del master) |
| Alcance de puntos | Ambos (`EscribirFormulaVisible` + `EscribirFormulaInterior` — mismo patrón, E7 exige agosto) | Solo visibles (dejaría el interior de agosto con el mismo defecto latente) |
| expand-to-plain | Rechazada: reescribir masters+seguidoras como fórmulas planas rompe la estructura del template, invalida f-vs-f contra el manual y agranda el diff órdenes de magnitud | — |
| skip-when-delta-zero | Rechazada como fix (enmascara: julio delta-0 pasaría pero cualquier delta futuro reabre el defecto); diferida como posible optimización futura con su propio T0, nunca como corrección | — |
| Dónde corre el gate | Tests + post-`AjustarEnWorkbook` donde corren `ExigirFormaRecompuesta`/`ExigirFormaRecompuestaInterior` (el punto donde el workbook ya está dimensionado y recompuesto) | Solo dominio C# (ciego a atributos shared por construcción — la trampa que este plan cierra, D-G) |
| Fixture shared | Congelado del disco (`Resultado1` + template + regen fresca; T1 lo fija y manda) | Conteos asumidos del encargo sin re-leer (si E4 difiere en ±1, el disco manda) |

### 2.3 Contratos afectados (resumen)

| Contrato | Cambio |
|---|---|
| `OpenXmlEspejoR1Mutador.EscribirFormulaVisible` (~L1514) | Preserve-and-rewrite: `CellFormula?.Text = texto` si existe; `new CellFormula(texto)` si no; log + atributos |
| `OpenXmlEspejoR1Mutador.EscribirFormulaInterior` (~L1209) | Idéntico (mismo patrón, E7) |
| `Reanclar` / dimensionado 5→1 / `ComponerTotOpt`/`ComponerTdf`/`ComponerExtemp` / `CompositorInteriorR1` / `DerivarFilasPorFirma` / `CeldasInterioresAgosto` | Intactos (el texto compuesto no cambia; solo se preservan atributos al escribirlo) |
| `SaneadorCadenaCalculo` / `CalculationProperties` / sellos / naming / preflight / readers / cálculo | Intactos (D-E) |
| Gate nuevo invariante-shared (+ fixture mapa shared) | Nuevo en `Remuneracion.IntegrationTests` (BCL zip+XML); corre en regresión + post-`AjustarEnWorkbook` donde corren los gates R1 |
| `ComparadorSalidaVsManualTests.BrechaDeclarada` | Sin cambios (el defecto no era divergencia de valores sino integridad estructural; el comparador sigue f-vs-f + literales) |
| `project-context.md` + manual | Doctrina preservación-shared (D-A/D-B) + invariante workbook-wide (D-C) + lección D-G + proceso (temp fresca + lectura de FAIL) |

### 2.4 Architecture Validation Certificate

| Principio | Veredicto |
|---|---|
| **SRP** | ✅ Una razón de cambio por costura: el mutador preserva atributos al escribir, el gate contrasta atributos del workbook, el fixture fija el mapa, el T0 arbitra agosto. Ninguna rama `if julio/agosto` en el fix (la preservación vale en todo período). |
| **OCP** | ✅ Un master nuevo futuro o una celda interior nueva se preservan sin código nuevo (la regla es por existencia de `CellFormula`, no por lista de celdas); julio-identidad garantiza que lo cerrado no se mueve. |
| **DIP** | ✅ El gate consume XML del workbook generado + fixture de tests; no crea dependencias nuevas (BCL ya presente; sin Excel/COM). El fix no agrega dependencias (OpenXML 3.5.1 ya referenciado). |
| **Best practices** | ✅ TDD rojo-primero (gate rojo con regen corrupta, verde con `Resultado1`); fail-fast con hoja+celda+`si` nombrados; quincena por dominio; `<f>` de negocio y `<v>` intactos; tolerancia ±0.5; insumos reales; sin Excel/COM en pipeline ni tests. |
| **Performance** | ✅ Fix O(1) por celda (una asignación vs una construcción); gate O(fórmulas de la hoja) lineal en lectura; cero impacto en julio (mismo texto, atributos preservados). |

---

## 3. SPEC

### 3.1 Requirements

| ID | Requisito | Fase | Criterio de aceptación |
|---|---|---|---|
| R-G-1 | Gate de invariante shared workbook-wide en `sheet10.xml` (`Reporte Componentes R1`): todo `si` con seguidoras tiene master con ese `si` en la hoja; fail-fast nombrando hoja+celda+`si` | F1 | TDD rojo-primero: FAIL con regen fresca corrupta (huérfanos `si=0`/`si=20`), PASS con `Resultado1` (control sano) |
| R-G-2 | Asserts explícitos `G53` (master `si=0`, `ref=G53:AP53`) y `G468` (master `si=20`, `ref=G468:AO468`) con texto idéntico al template | F1 | Rojo-corrupta / verde-`Resultado1` / verde-post-fix, por celda |
| R-G-3 | Fixture del mapa shared esperado (35 masters, rango `si`, ~1035 seguidoras — T1 congela exactos desde el disco; el fixture manda) | F1 | Test-drift re-lee template+`Resultado1` y pasa; si el disco difiere de E4, el fixture fija y E4 se enmienda |
| R-G-4 | Caso agosto en el gate (E7: `G468` + workbook-wide sobre salida agosto — el interior puede golpear más celdas) | F1 | El gate delata en rojo toda celda agosto con `si` huérfano pre-fix y cierra en verde post-fix (no solo `G53`/`G468`) |
| R-B-1 | Fix preserve-and-rewrite en ambos puntos (`EscribirFormulaVisible`, `EscribirFormulaInterior`): `CellFormula` existente → mutar `.Text`; inexistente → `new CellFormula` | F1 | Diff mínimo (unas líneas); `Reanclar` y compositores intactos (verificado por diff); log con ASE+celda+etiqueta+texto+atributos |
| R-B-2 | Julio-identidad extendida: regen post-fix 2026072 = template/`Resultado1` en texto de fórmula Y en atributos (`t`/`ref`/`si`) para toda la hoja | F1 | 0-diff R1 julio (texto + atributos) + goldens Q1+Q2 ±0.5 + DetRetri-D 5/5 vs R10 ambos períodos |
| R-B-3 | CalcChain y cachés intactos: `SaneadorCadenaCalculo`, `fullCalcOnLoad`, `<v>` sin tocar | F1 | Test estructural existente verde; grep: ningún `<v>` tocado; ningún flag de cálculo cambiado por este plan |
| R-T-1 | T0 Agosto-2 bloqueante: veredicto H1/H2/H3 para `ASE 2 Reporte Componentes R1!F199 [SUB_EMP]` (`faltan [Mes0]`) con evidencia en disco + cita a los 3 .docx; cero código de producción | F2 | Tabla de veredicto por hipótesis (confirmada/descartada) con archivo+hoja+celda/fila citados; `git diff --stat` vacío en producción para T4 |
| R-T-2 | T5 condicional: fix de agosto SOLO tras veredicto T4, con diseño acotado al veredicto (no prescrito aquí) | F2 | T5 no inicia sin T4 cerrado; su diseño cita el veredicto |
| R-R-1 | Suite base verde antes y después + build 0 warnings en cada tarea; goldens julio ±0.5; DetRetri-D 5/5 vs R10 (julio y agosto donde aplique) | Transversal | `dotnet build Remuneracion.WinForms/Remuneracion.WinForms.slnx` + `dotnet test` por tarea |
| R-R-2 | Apertura manual sana en Excel (paso del Ingeniero): salida Julio-2 post-fix abre SIN diálogo de reparación | F1 cierre | Aceptación final manual (el pipeline demuestra invariante verde; Excel lo confirma el humano) |
| R-R-3 | `project-context.md` (doctrina + invariante + lección D-G) + manual (temp fresca + lectura de FAIL) | Transversal | Diff acotado; sin código en el mismo commit |

### 3.2 Scenarios (Given/When/Then)

- **S1 (TDD rojo — reproduce el huérfano):** Given regen fresca 2026072 post-commit a temp (E5), When el gate verifica `sheet10.xml`, Then FAIL con `Reporte Componentes R1!N53 [si=0]: sin master` (y `U468 [si=20]`), mientras `Resultado1` da PASS con el mismo gate.
- **S2 (atributos explícitos):** Given regen corrupta, When el gate verifica `G53`/`G468`, Then FAIL con `G53: master si=0 ref=G53:AP53 ausente` (y análogo `G468`); Given `Resultado1`, Then PASS con texto idéntico al template.
- **S3 (fix que cierra):** Given regen post-fix a temp fresca, When el gate + goldens, Then invariante verde workbook-wide (julio y caso agosto), texto idéntico, goldens ±0.5, DetRetri-D 5/5, y Excel abre sin reparación (paso manual R-R-2).
- **S4 (agosto interior):** Given salida agosto actual, When el gate workbook-wide, Then delata en rojo TODO `si` huérfano (incl. celdas del pase interior si las hay, E7); post-fix, Then verde.
- **S5 (T0 agosto):** Given template `C199` + fuente R1 ASE2-agosto + `CeldasInterioresAgosto[2]`, When T4, Then veredicto H1/H2/H3 con citas (una confirmada o combinación explicitada) y diseño acotado resultante para T5; el diff de producción de T4 es vacío.
- **S6 (no-regresión estructural):** Given el diff F1, When se audita, Then ningún `<v>`, ningún flag calcChain, ningún texto `<f>` de negocio y ningún mapa cambiado salvo las líneas de preservación (grep de cierre).

---

## 4. TASKS

| ID | Tarea | Fase | Depende de | Verificación |
|---|---|---|---|---|
| T1 | Gate-que-falla-primero (TDD) + fixture shared. (a) Resolver `sheet10.xml`↔`Reporte Componentes R1` desde el disco y congelar fixture (35 masters, rango `si`, ~1035 seguidoras; asserts `G53`/`G468` ref+si+texto) desde template + `Resultado1` (+ test-drift). (b) Nuevo gate workbook-wide (BCL zip+XML, prohibido `<v>`, fail-fast hoja+celda+`si`) con tests S1/S2 (rojo regen corrupta E5 a temp, verde `Resultado1`) + caso agosto S4 en rojo. | F1 | — | R-G-1..R-G-4 (S1/S2/S4 pre-fix en rojo + verde-control); build 0 warnings; resto de la suite verde (solo lo nuevo en rojo) |
| T2 | Fix preserve-and-rewrite (mínimo). Mutar `.Text` in place en `EscribirFormulaVisible` (~L1514) y `EscribirFormulaInterior` (~L1209) cuando `CellFormula` existe; `new CellFormula` solo si no existe; log con atributos; resto intacto (`Reanclar`, compositores, literales-0, saneador, sellos). Tests: unitarios del helper por rama (con/sin `CellFormula` previa) + S3 en verde sobre regen fresca. | F1 | T1 | R-B-1 (S3 parcial); diff = unas líneas + tests + gate; build 0 warnings |
| T3 | Cierre Fase 1. Regen fresca 2026072 post-fix a temp (E5): invariante verde workbook-wide + caso agosto verde + 0-diff texto+atributos vs `Resultado1`/template + goldens Q1+Q2 ±0.5 + DetRetri-D 5/5 vs R10 + test estructural calcChain verde (R-B-3/S6) + grep de cierre (ningún `<v>`, ningún texto `<f>` de negocio, ningún flag cambiado). Paso manual del Ingeniero: abrir en Excel y confirmar SIN reparación (R-R-2). | F1 | T2 | R-B-2/R-B-3/R-R-1/R-R-2 (S3/S4/S6); suite completa verde |
| T4 | **T0 Agosto-2 BLOQUEANTE (cero código).** Arbitrar H1/H2/H3 para `ASE2 F199 [SUB_EMP] [Mes0]` con evidencia en disco: (H1) `C199` template vs filas-empresa fuente R1 ASE2-agosto (¿nombre divergente? — matriz EFC×ASE de `Proceso de Recaudo` + catálogo abierto incl. `NUEVO ESQUEMA` del Plan 32 como marco); (H2) filas oportunas del sub-bloque en la fuente (¿ausentes? — doctrina espejo Plan 21: ausente=suprimir, pero subtotal sin datos no componible); (H3) `CeldasInterioresAgosto[2]` vs geometría real del template agosto (¿stale?). Citas a `Prompt_Maestro §5–8`, `Detalle de plantilla` (instr. 5 + R1 por empresa), `Proceso de Recaudo` (matriz EFC×ASE). Veredicto + diseño acotado resultante para T5. | F2 | T3 (secuencia ordenada por el Ingeniero; no usa su código) | R-T-1 (S5); `git diff --stat` vacío en producción; sin datos inventados |
| T5 | Fix Agosto-2 CONDICIONAL (diseño NO prescrito aquí — lo define el veredicto T4). Solo inicia con T4 cerrado; vive en el mismo pase final del mutador por defecto salvo que T4 demuestre otro punto; con TDD rojo-primero y julio-identidad como red. | F2 | T4 | R-T-2; goldens ±0.5; DetRetri-D 5/5; build 0 warnings |
| T6 | Docs: `project-context.md` (doctrina preservación-shared D-A/D-B + invariante D-C + lección D-G + nota E7) + manual (regen SIEMPRE a temp fresca + cómo leer un FAIL del gate S1/S4). | Transversal | T3 (F1), T4 (cita del veredicto, sin adelantar T5) | R-R-3; diff acotado; sin código en el mismo commit |
| T7 | Cierre global: build 0 warnings + suite completa verde + goldens ±0.5 + DetRetri-D 5/5 ambos períodos + grep de cierre global (diff producción F1 = líneas de preservación; T4 sin diff; T5 solo lo que su veredicto autorice) + CRLF + sin commits. | Transversal | T1..T6 | R-R-1; DoD §5 |

**Orden sugerido (estricto por mandato):** T1 (rojo) → T2 (fix) → T3 (cierre F1 + aceptación manual Excel) → T4 (T0 bloqueante agosto) → T5 (condicional al veredicto) → T6 → T7. Build 0 warnings antes de marcar cada tarea completa. Sin commits (los hace el Ingeniero con `#commit`/`#push`).

---

## 5. Resumen ejecutivo, riesgos y decisiones

- **Resumen:** el debug-agent cerró la causa (E2: `new CellFormula` descarta `t`/`ref`/`si` de los masters `G53`/`G468`; E4: regresión `73de7ea` con 2 masters perdidos y texto idéntico; E7: `G468` también en el pase interior de agosto). La Fase 1 convierte esa evidencia en: gate workbook-wide que hoy fallaría (T1, con `Resultado1` como control verde), fix mínimo preserve-and-rewrite en ambos puntos de escritura (T2) y cierre con regen fresca + goldens + apertura manual sana (T3). La Fase 2 secuencia el fallo de agosto (`F199 [SUB_EMP]` sin `Mes0`) como T0 bloqueante (T4) que arbitra nombre-divergente vs filas-ausentes vs mapa-stale con cita a los 3 .docx de negocio, y deja su fix (T5) condicional al veredicto — no prescrito aquí por mandato.
- **Riesgo principal:** R-SHARED-OCULTO (más masters afectados que los 2 visibles, vía el interior de agosto). Contenido por triple red: gate workbook-wide (no lista cerrada), fix en ambos puntos de escritura, y caso agosto incluido en el gate desde T1.
- **Decisión para el Ingeniero:** ratificar D-A..D-G (§0.3). **Forks con recomendación:** D-A (preserve-and-rewrite vs expand-to-plain vs skip-when-delta-zero — recomendado preserve: mínimo, preserva estructura del template y f-vs-f; las otras dos rechazadas con motivo) y D-F (T0 bloqueante antes de prescribir agosto — sin alternativa recomendada: es mandato). **Sin preguntas bloqueantes pendientes para Fase 1; Fase 2 bloqueada hasta T4 por diseño.**

### DoD (Definition of Done)

1. `dotnet build Remuneracion.WinForms/Remuneracion.WinForms.slnx` con 0 warnings.
2. Suite completa verde sin regresión (base + nuevos) + goldens Q1+Q2 ±0.5 + DetRetri 5/5 vs R10 (julio y agosto donde aplique) + 0-diff R1 julio (texto + atributos shared).
3. Julio-2 post-fix regenerado a temp fresca: invariante shared verde workbook-wide (+ caso agosto verde) y apertura manual en Excel SIN reparación (paso del Ingeniero, R-R-2).
4. T4 con veredicto H1/H2/H3 publicado y citado (negocio + disco); T5 solo lo que el veredicto autorice; `project-context.md` + manual actualizados.
5. Grep de cierre: diff producción F1 = líneas de preservación en los 2 puntos (ningún `<v>`, ningún texto `<f>` de negocio, ningún flag calcChain, ningún mapa); sin commits del agente; CRLF; sin emojis.

### Follow-ups explícitos (fuera de este plan)

- Espejo-motor R4, R4-por-empresa Q2, CF/DV del reanclaje (W-4), `R-EXTRA-CONCEPTO` (Plan 23), validación de nombres de hoja en preflight, remesh R2/R4 (F-T4-2), re-encendido SALDOS/AJUSTES (F-T4-3), `INTERVENTORIA!R26` (F-T4-4): vivos, intactos.
- Toda optimización tipo skip-when-delta-zero futura: su propio T0 antes de cualquier línea (D-A).
- Retiro/archivo de evidencias: `Resultado1/2` se conservan como control sano; las regens de evidencia viven en temp, no se versionan.
