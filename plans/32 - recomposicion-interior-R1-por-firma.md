# Plan 32 — Recomposición por firma del INTERIOR de `Reporte Componentes R1` (sub-visibles por empresa) + fixture congelado + gate-v2 estructural

> **Alcance:** extender la recomposición por firma del Plan 31 (15 visibles: TOT_OPT F, TDF G, EXTEMP F por ASE) al **INTERIOR de `Reporte Componentes R1`** — los sub-visibles con offsets de anclaje mecánico julio (ej. `F60` app 3 términos vs manual 2; `F195` app `F150` vs manual `F148`), par a par con la recomposición de los visibles del Plan 31/T2. Superficie: fixture congelado del interior desde el manual + recomposición por firma del interior en el pase final del mutador (misma disciplina: `<f>` compuesta, log auditado, julio-identidad, guard ajustado solo en el pase del mutador) + tests (texto-vs-manual del interior, gate-v2 estructural, suite) + docs mínimos. **Cero otros cambios.**
> **Rector (IRRENUNCIABLE):** `README.md` + `Docs/Propuesta_Proyecto_Automatizacion_Remuneracion_UAESP.md` + `Docs/Detalle de plantilla.docx` + `Docs/Prompt_Maestro_Proyecto_Remuneracion_UAESP _ Vo .docx` + `.opencode/project-context.md`. **Invariantes del proyecto:** NUNCA sobrescribir fórmulas —salvo la excepción legítima, explícita y acotada de este plan (extensión del contrato T2 del Plan 31 al interior, D-D)—; tolerancia ±0.5; ASE4 sin columna "Especiales" (0 si ausente); Q2 aplica SALDOS POR NOTA/RETRIBUCION NEGATIVA (`Periodo.NumeroQuincena == 2`, `AjustesSfT = 0` en Q1); redondear a entero antes de DetRetri/columna D; build 0 warnings; finales de línea CRLF; fail-fast nombrando archivo+hoja+celda; tests SOLO con insumos reales (`Docs/Prueba2/Insumos`, `Docs/Prueba Julio-2/Insumos`); sin commits (los hace el Ingeniero con `#commit`); sin emojis. **Quincena por dominio, nunca por contenido.**
> **Continuidad:** Plan 31 CERRADO (suite 443/444; único rojo intencional: comparador agosto = 96 inesperadas TODAS en `Reporte Componentes R1` interior, R-D-5). Gate S1 (`ValidadorTotalesR1Workbook`) verde ambos períodos para los 15 visibles. Recomposición activa en `OpenXmlEspejoR1Mutador` (contrato visibles: TOT_OPT F, TDF G, EXTEMP F; 0-Aplic→literal 0 auditado; julio-identidad). Este plan NO reabre visibles, lectura, cálculo, espejo-dimensional, preflight, sellos, naming, calcChain ni paridad fuera del interior R1: **julio = byte-idéntico por construcción.**
> **Insumos canónicos (no se re-fijan, no se inventan datos):** T0 de este plan (§0.1, método zip+XML BCL, sin Excel/COM); manual agosto `Docs/Prueba2/Resultado/Resultado Manual por el administrativo/Remuneracion 202608-2 Total_7721.xlsx`; app agosto `Docs/Prueba2/Resultado/Resultado3/Remuneración 202608-2 Total.xlsx` (R3, estado pre-T2: conserva el anclaje viejo también en el interior — útil como negativo); manual julio `Docs/Prueba Julio-2/Resultado/Remuneracion 202607-2 Total Administrativo.xlsx`; app julio `Docs/Prueba Julio-2/Resultado/Resultado2/Remuneración 202607-2 Total.xlsx`; base canónica agosto `Docs/Prueba2/Plantilla_Remuneracion_2026082.xlsx`; evidencia previa `plans/31-T0-Evidencia-R1.md` (causa + visibles) y `plans/31-T4-Veredictos-R2R4-soporte.md` (R-D-5/F-T4-1 + conteos).
> **Numeración:** `plans/` 01..31 ocupados (31 + anexos `31-T0-Evidencia-R1.md` y `31-T4-Veredictos-R2R4-soporte.md`); este plan toma el primer correlativo libre, **32**.
> **Estado:** PLANIFICADO (T0 cerrado en disco, §0.1) — planificado desde T0 propio, sin implementar.
> **Fecha:** 2026-10-08

---

## 0. Clarification Gate

**Sin preguntas bloqueantes: el plan viene redactado con las decisiones D-A..D-G (§0.3) y T0 ya cerrado en disco (§0.1).** Los dos puntos con recomendación (dónde vive la reescritura interior, D-A; promesa del gate S1 —extender vs comparador—, D-F) vienen redactados en la ruta recomendada con su verificación de cierre dentro de la tarea. Punto de bloqueo respetado: NO se implementa nada en este documento.

### 0.1 T0 — inventario clasificado del interior (causa cerrada para visibles, causa abierta para el interior; no re-investigar lo cerrado)

Método (BLOQUEANTE cumplido, BCL zip+XML, sin Excel/COM): volcado fórmula-vs-fórmula app-R3 vs manual-agosto en `Reporte Componentes R1` (213 diffs brutos) + etiquetas A–E del manual en filas-ancla + volcado app-Resultado2 vs manual-julio (0 diffs). Reconciliación con el rojo vivo del comparador (T4 §1.0/1.5: 96 inesperadas post-T2, todas R1-interior): `213 brutos − 15 visibles (fijados por T2) − ~102 seguidores de fórmula compartida con texto vacío (excluidos por H1/T0g, familia AA..Z + H50/H456-style) ≈ 96`. El desglose estático sustantivo suma ~100 celdas (el vivo absorbe 4 bordes shared/empty: `G355`, `H464`, `H584`-style); **T1 congela el conjunto exacto desde el rojo vivo**, con la tabla de abajo como base (sobre-incluir es seguro: el assert es texto-vs-manual; sub-incluir es el riesgo).

#### T0a — inventario por ASE/bloque con patrón del offset (la fórmula del MANUAL es la fuente de verdad del objetivo)

Firma de filas (leída del manual, misma normalización que el gate): subtotal-empresa `C=<empresa> ∧ D='OPORTUNO' ∧ E='Total'`; EXTEMP-interior `D='EXTEMPORANEO' ∧ E='Total'`; Subs-interior `E='Subsidio(-)/Contribucion(+)'`; AFaseo `D='AFaseo' ∧ E='Total'`. Anclas: mixtas (ver T0c).

| ASE | Celda interior (F, muestra; G/H espejan) | App R3 (anclaje julio) | Manual (objetivo) | Patrón del offset | Clase |
|---|---|---|---|---|---|
| 1 | `F60` [ENEL/OPORTUNO] | `F40+F21+F8-L8-L21` (3 térm.) | `F37+F18-L18` (2 térm.) | App conserva FORMA julio; anclas viejas + término de más | R-1 |
| 1 | `F70` [OCCIDENTE/OPORTUNO] | `F47+F31+F11-L11-L31` | `F44+F28+F8-L8-L28` | Offset −3/−3/−3 (recorte en cabeza, igual que visibles) | R-1 |
| 1 | `F62` [EXTEMP-interior] | `F36+F16-L16` | `F33+F13-L13` | Offset −3/−3/−3 | R-1 |
| 1 | `F51/F53/F61/F63/F71` [Subs] | `F47/F37/F40/F36/F47` | `F45/F34/F37/F33/F44` | Single-ref desplazada (−2/−3) | R-1 |
| 1 | `F80` [AFaseo] | `F13` | `F10` | Single-ref −3 | R-1 |
| 1 | `G60/G70/G62` (+`H60/H62/H52`) | `G21+G8` / `G31+G11` / `G16` | `G18` / `G28+G8` / `G13` | TDF-interior sigue regla E2 (ΣG menos última, cardinalidad variable: 2→1 y 2→2) | R-1 |
| 2 | `F195/F197` [Subs] | `F150/F139` | `F148/F130` | Single-ref (−2/−9); ancla = fila `Mes` de bloque (firma existente ✓) | R-1 |
| 2 | `F199` [RECIPROCIDAD] | `F149+F118-L118` | `F140+F114-L114` | Offset mecánico (−9/−4/−4) | R-1 |
| 2 | `F204` [ENEL] | `F142+F108+F87--L87-L108` | `F133+F105+F87--L87-L105` | Offset + quirk `--L87` en AMBAS (ver H-DOBLE-SIGNO) | R-1 |
| 2 | `F206` [EXTEMP-int.] | `F133+F96-L96` | `F129+F97-L97` | Términos cruzados (F−4/+1): re-derivar por firma, no por delta | R-1 |
| 2 | `F214` [OCCIDENTE] | `F150+F128+F90-L90-L128` | `F147+F124+F92-L92-L124` | Términos cruzados (−3/−4/+2) | R-1 |
| 2 | `F216` [EXTEMP-int.] | `F138+F103-L103` (3 térm.) | `F100-L100` (2 térm.) | App 3 térm. vs man 2 (el 3.er térm. apunta a fila eliminada) | R-1 |
| 2 | `F217` [Subs] | `F138` (fórmula) | literal (sin `<f>`) | **Caso F217: el manual trae literal** (línea Subs de sub-bloque podado) | L-1 |
| 2 | `F224` [AFaseo] | `F91` | `F93` | Single-ref con signo contrario (+2: la fila correcta está ABAJO) | R-1 |
| 3 | `F335` [ENEL] | `F261+F242-L242` (2 térm.) | `F255+F240+F231-L231-L240` (3 térm.) | Man 3 térm. vs app 2 (término NUEVO por geometría agosto) | R-1 |
| 3 | `F345` [OCCIDENTE] | `F265+F252+F231-L231-L252` | `F262+F250+F234-L234-L250` | Offset mecánico (−3/−2/+3) | R-1 |
| 3 | `F326/F336/F346/F355` [Subs/AFaseo] | `F265/F261/F265/F232` | `F263/F255/F262/F235` | Single-refs (−2/−6/−3/+3) | R-1 |
| 3 | `F327/F328` + `F337/F338` (+`G327/G337`,`H327/H337`) [EXTEMP-int./Subs] | fórmula (`F258+F238-L238`, `F258`, …) | **sin `<f>`** | Filas Aplic-interior que el manual ELIMINÓ (ASE3-agosto 0 filas `Aplicacion`: el interior EXTEMP no existe en el manual) | L-1 |
| 4 | `F461` [RECIPROCIDAD] | `F422+F393-L393` | `F428+F396-L396` | Offset +6/+3 (delta>0: fuente agosto +6 filas) | R-1 |
| 4 | `F463` [EXTEMP-int.] | `F374-L374` (2 térm.) | `F414+F377-L377` (3 térm.) | App pierde un término (fila insertada sin fórmula apta) | R-1 |
| 4 | `F464` [Subs] | **sin `<f>`** | `F414` | **App SIN fórmula donde el manual la trae** (fila insertada, clon vacío — ver T0c) | L-2 |
| 4 | `F466/F476` (+`G466/G476`,`L569`) | `F415+F383+F362-L362-L383`, `F429+F403+F365-L365-L403` | `F421+F386+F362-L362-L386`, `F435+F406+F365-L365-L406` | Offsets +6/+3 con término fijo común (`F362`,`L362`,`F365`) | R-1 |
| 5 | `F564/F569/F574` [ENEL/ENERBIT/OCCIDENTE] | `F529+F504-L504`, `F532+F508-L508`, `F518+F539-L518` | `F537+F510-L510`, `F540+F514-L514`, `F524+F547+F493-L493-L524` | Offsets +8/+6; `F574` man 3 térm. vs app 2 (término nuevo) | R-1 |
| 5 | `F584`/`G584` [AFaseo] | **sin `<f>`** | `F495` / `G495` | App sin fórmula (delta +8, clon vacío) | L-2 |

Clases: **R-1** recomponible-por-firma (mismo patrón que los visibles: re-derivar anclas por firma del sub-bloque; el manual es el objetivo textual); **L-1** literal-en-manual (app trae fórmula huérfana de filas que el manual eliminó —podar a literal/vacío según el manual, caso `F217` incluido); **L-2** fórmula-ausente-en-app (fila insertada por delta>0 cuyo clon no trae `<f>` —componerla como R-1). G/H espejan F en su columna (misma form
a; `H` con seguidores shared vacíos ya excluidos por H1/T0g). `L569` (`L508`→`L514`) espeja en L.

**Hallazgo H-DOBLE-SIGNO:** `F204` (y `F574`) traen `--L87` en AMBAS (app y manual): quirk textual del template preservado por los dos reanclajes. La recomposición interior debe reproducir el texto EXACTO del manual (el fixture lo fija); si la emisión canónica (`…+F87-L87…`) difiere del disco (`…+F87--L87…`), T1 lo delata y el compositor deriva el patrón de signos de la fórmula del template base para esa forma (no inventa).

#### T0b — fixture `InterioresR1Esperados` (mismo método del Plan 31)

Se congelan del MANUAL (solo lectura zip+XML) los textos esperados del interior: **todas las celdas sustantivas del inventario T0a (≈100: F/G/H/L interiores con `<f>` no-vacía en el manual + literales L-1 con cita)**, como fixture `InterioresR1Esperados` (celda+texto+cita `archivo+hoja+celda`), estables julio/agosto (julio: el manual-julio fija las formas canónicas; agosto: el manual-agosto fija las recompuestas). Test espejo de `TotalesR1EsperadosTests`: re-lee el disco y delata drift del manual antes de que el compositor se confunda.

#### T0c — ¿alcanza la firma existente? NO: se necesita nueva firma para filas de sub-bloque

Evidencia (etiquetas A–E del manual en filas-ancla, 2026-10-08): los anclajes interiores son MIXTOS. Parte resuelve con firma existente (`F195→F148`, `F224→F93`, `F494/525/548`: `B='Mes' ∧ C='Total'` = `EsMesTotal` ✓; `F196→F130/F101`: `B` contiene `Aplicacion…` = `EsAplicacionTotal` ✓). Pero los términos de empresa resuelven a filas con firma **`C=<empresa> ∧ D='Total'`** (`F18/F37`=[ENEL/Total]→`F60`; `F140/F114`=[NUEVO ESQUEMA/Total]→`F199`; `F100`=[OCCIDENTE/Total]→`F216`; `F133/F105`=[ENEL/Total]→`F204`) — firma que NINGÚN predicado vigente cubre (`EsMesTotal` exige `B='Mes'`; `EsCodigoD` mira la columna D de códigos, no C). **Veredicto: se necesita nueva firma** (`EsDatoEmpresa`: `C` no-vacía ∧ `D='Total'` ∧ `B` vacía, con el catálogo de empresas ABIERTO —el disco trae `NUEVO ESQUEMA` además de ENEL/OCCIDENTE/RECIPROCIDAD/ENERBIT—; más localizadores de fila-subtotal por `(C,D,E)`), a congelar en T1 desde manual+fuente. La fuente `Recaudoporcomponente` trae las mismas A–E (el lector las copia textual, `:717-723`), así que la secuencia fuente ALCANZA como especificación del sub-bloque por orden, igual que en visibles.

Mecanismo dual confirmado (explica L-1/L-2): con delta<0 (ASE1/2/3) la app conserva fórmulas reancladas en filas que el manual podó (huérfanas → L-1); con delta>0 (ASE4/ASE5) `ClonarFilaVacia` crea celdas SIN `<f>` y el manual sí las trae (→ L-2). Ambos se cierran componiendo por firma en el pase final (podar = componer literal/vacío según el manual; insertada = componer la forma del sub-bloque).

#### T0d — julio: 0 fórmulas del interior difieren app-vs-manual HOY

Volcado app-Resultado2 vs manual-julio en `Reporte Componentes R1`: **0 divergencias de fórmula (1155/1155 `<f>`)**. No hay residuales de julio: el invariante julio = byte-idéntico SIGUE siendo la red de reversión (si el pase interior altera julio, es regresión y se revierte el pase, no el golden). La derogación documentada de `R1Q2ProtectedPorAse` (Plan 31/T2) se EXTIENDE al interior en este plan; el guard anti-fórmula genérico sigue aplicando (ajuste SOLO en el pase del mutador, D-D).

#### T0e — promesa del gate S1: extender con gate-v2 ESTRUCTURAL (recomendado D-F)

El dominio C# NO trae agregados por empresa (`WorkbookLeafInputsR1` solo: `TotalOportunoEsperadoPorAse`/`ExtemporaneoEsperadoPorAse`; `ValidarSigmaEmpresasPorAse` valida Σ-empresas de conciliación, no del R1-interior): un gate numérico interior exigiría INVENTAR agregados de dominio — prohibido. **El comparador f-vs-f basta como red numérica** (texto-idéntico + mismos inputs ⇒ mismo post-recálculo, método Plan 29), y el gate se extiende SOLO estructuralmente: `ExigirFormaRecompuestaInterior` verifica que cada subtotal-interior referencia exactamente las filas-firma de su sub-bloque (conjunto de refs, sin evaluar valores). Sin gate numérico nuevo.

### 0.2 Mapeo al Rector (Sale / No sale)

**Sale (EXPLÍCITO):**
- Re-investigar visibles o causa E1 del Plan 31 (cerradas) fuera de la re-verificación T1.
- Tocar el lector (`MapearR1Q2`, `LeerEspejoR1`), el cálculo, `CONSOLIDADO`, R2/R4-detalle, preflight, sellos, naming, calcChain, BCE, columnas por quincena.
- Reescribir `<f>` fuera del interior R1 declarado en el contrato T2 (visibles + interior; el resto intacto; la autorización vive SOLO aquí, §2.3/D-D).
- Inventar agregados de dominio por empresa para un gate numérico (T0e: prohibido; la red numérica es el comparador f-vs-f + fixture).
- Normalizar el quirk `--L` (H-DOBLE-SIGNO: el disco manda; se reproduce, no se corrige).
- Julio con diff ≠ 0 en R1-interior (identidad; si difiere, es regresión).
- Paquetes NuGet; commits (los hace el Ingeniero con `#commit`).

**No sale (entra, alcance congelado):**
| Frente | Superficie de cambio |
|---|---|
| F — fixture | Nuevo `InterioresR1Esperados` (+ tests espejo de `TotalesR1EsperadosTests`) congelado del manual julio+agosto |
| B — recomposición interior | Pase final en `OpenXmlEspejoR1Mutador` (extensión de `RecomponerVisibles` al interior por sub-bloque/empresa; clases R-1/L-1/L-2; log ASE+celda+texto) + extensión documentada de la derogación `R1Q2ProtectedPorAse` |
| G — gate-v2 | `ExigirFormaRecompuestaInterior` (estructural por firma, sin valores) + tests; el comparador pierde el rojo R-D-5 al cerrar (0 inesperadas) |
| D — docs | `project-context.md` (doctrina interior-por-firma + nueva firma + H-DOBLE-SIGNO + clases L) + manual (proceso) |

### 0.3 Decisiones ejecutivas (aprobación = aceptar estas)

| ID | Decisión |
|---|---|
| D-A | **La reescritura interior vive en el mismo pase final post-5→1 del mutador** (extensión de `RecomponerVisibles`, no pase nuevo ni writer). El mutador posee la geometría final; el writer con `omitirR1` no la conoce. Descartado: componer por sub-bloque dentro de `AjustarBloque` (los `Reanclar` posteriores lo reescribirían) y en el writer/validación. |
| D-B | **Firma nueva para sub-bloque (T0c), sin conteo ni borde.** Localizar cada fila-subtotal por `(C,D,E)`; alinear sub-bloque fuente ↔ filas finales por orden (misma cardinalidad por construcción del delta); componer `<f>` con las filas-dato del sub-bloque por firma (`EsDatoEmpresa` + existentes). El catálogo de empresas en C es ABIERTO (el disco trae `NUEVO ESQUEMA`); prohibido `if ENEL/OCCIDENTE`. |
| D-C | **Clases L-1/L-2 se componen, no se podan a ciegas.** L-1 (manual sin `<f>`): componer literal/vacío EXACTO del manual (caso `F217` y filas `F327/F328/F337/F338` de ASE3-0Aplic —el manual no trae esas filas-fórmula); L-2 (app sin `<f>` por clon vacío): componer la forma del sub-bloque (`F464→F414`, `F584→F495`). Lo que diga el disco para esa celda MANDA (fixture). |
| D-D | **Excepción `<f>` extendida al interior, con la misma auditoría término a término.** Solo celdas del contrato T2-interior (§2.3); DOS asserts por celda (texto = manual + forma = firma del sub-bloque); log Debug ASE+celda+texto antes/después; guard anti-fórmula ajustado SOLO en el pase del mutador (fuera de él, intacto). |
| D-E | **Julio-identidad como invariante de reversión (T0d: 0 diffs hoy).** Con geometría julio el pase interior reproduce byte-idénticas las fórmulas canónicas; goldens ±0.5 + 0-diff R1 (f-vs-f + literales) como red por tarea. Si julio difiere, se revierte el pase. |
| D-F | **Gate-v2 ESTRUCTURAL (recomendado) + comparador como red numérica (T0e).** `ExigirFormaRecompuestaInterior`: conjunto-de-refs == filas-firma del sub-bloque (sin evaluar valores, sin agregados inventados). La paridad numérica la demuestra el comparador f-vs-f (0 inesperadas) + fixture texto-vs-manual. Descartado: gate numérico interior (exigiría agregar dominio por empresa: dato inventado). |
| D-G | **El quirk `--L` se reproduce (H-DOBLE-SIGNO).** El compositor deriva el patrón de signos de la fórmula del template base para esa forma; el fixture lo fija. Nunca normalizar a `-L`. |

---

## 1. PROPOSE

### 1.1 Intent

Que TODA fórmula de `Reporte Componentes R1` —visibles (Plan 31) e interior (este plan)— se derive de las filas reales del período por firma, que ningún subtotal mal anclado vuelva a entregarse en silencio (gate-v2 que hoy fallaría + comparador en verde), con julio byte-idéntico.

### 1.2 In Scope

- Fixture `InterioresR1Esperados` congelado del manual (julio+agosto) + test-drift del disco.
- Recomposición interior por sub-bloque/empresa en el pase final (clases R-1/L-1/L-2, incl. `--L` y caso `F217`).
- Gate-v2 estructural + cierre del rojo R-D-5 (96→0 inesperadas).
- Docs: `project-context.md` (doctrina interior + firma nueva + H-DOBLE-SIGNO + clases L) + manual (proceso).

### 1.3 Non-Goals

Ver §0.2 "Sale". En particular: NO se agregan agregados de dominio por empresa; NO se toca el dominio numérico (ya correcto); NO se generaliza a otras hojas; NO se corrige el `--L`.

### 1.4 Riesgos

| ID | Riesgo | Mitigación |
|---|---|---|
| R-FORMA-INT | La forma interior recompuesta diverge del manual en un borde (cardinalidad de sub-bloque, `--L`, L-1/L-2) y la excepción `<f>` crea divergencia silenciosa | Fixture congelado ANTES de componer (T1); doble assert por celda (texto + forma); gate-v2 + comparador f-vs-f + julio-identidad en cada tarea |
| R-FIRMA | La firma nueva no cubre una forma de sub-bloque futura (otra empresa en C, sub-bloque sin EXTEMP) y el pase falla o compone mal | Catálogo C abierto por construcción (D-B); fail-fast que nombra ASE+celda+firma si un sub-bloque no resuelve (nunca 0 silencioso); T1 congela TODAS las formas del disco |
| R-JULIO | El pase interior altera julio | D-E: 0-diff R1 julio (f-vs-f + literales) + goldens ±0.5 por tarea; ante diff se revierte el pase |
| R-SIGNO | El compositor normaliza `--L` y rompe texto-vs-manual | D-G: patrón de signos del template base; fixture con `F204/F574` como guardianes explícitos |
| R-LITERAL | L-1/L-2 se implementan como "podar a 0" genérico y mienten donde el manual trae fórmula (o viceversa) | D-C: lo que diga el disco para ESA celda manda (fixture por celda, no por clase); asserts por celda |

---

## 2. DESIGN

### 2.1 Enfoque: el mismo pase, un nivel más adentro; la firma crece, el método no cambia

**Frente F (fixture).** Nuevo `InterioresR1Esperados` en `Remuneracion.IntegrationTests` (misma forma que `TotalesR1Esperados`: record `InteriorR1(AseId, Celda, Etiqueta, Formula|null, Cita)` + listas `Agosto`/`Julio` + tests-drift que re-leen el disco). Etiquetas: `SUB_EMP` (subtotal empresa F/G/H), `SUB_TDF` (G/H de empresa), `EXT_INT` (EXTEMP-interior), `SUBS` (single-ref Subsidio), `AFASEO`. `Formula null` = el manual trae literal (clase L-1: se congela también el valor literal esperado).

**Frente B (recomposición).** Extensión de `RecomponerVisibles` en `OpenXmlEspejoR1Mutador`: tras los 3 visibles por ASE, por cada fila-subtotal interior (localizada por firma `(C,D,E)` en el workbook ya dimensionado): (a) determinar su sub-bloque (filas entre el subtotal anterior y este, o entre invariantes de zona); (b) alinear con la secuencia fuente del mismo rango por orden; (c) componer `<f>` por firma de las filas-dato del sub-bloque (nueva `EsDatoEmpresa` + `EsMesTotal`/`EsAplicacionTotal` donde el ancla es de bloque); (d) L-1 → escribir literal/vacío exacto del fixture; L-2 → componer normal (la fila existe, solo le faltaba `<f>`). Formas (idénticas en espíritu al dominio, con cardinalidad real del sub-bloque): `SUB_EMP_F = ΣF(datos-emp) − ΣL(datos-emp menos último)`; `SUB_TDF_G/H = conjunto E2 (G/H de datos-emp menos último)`; `EXT_INT_F = ΣF(Aplic-sub) − L(primera)` (0 Aplic-sub → literal según fixture, cf. ASE3); `SUBS = F(ancla)`; `AFASEO = F(ancla)`. Orden de términos y patrón de signos: el del template base para esa forma (D-G), verificado por fixture.

**Frente G (gate-v2).** `ValidadorTotalesR1Workbook.ExigirFormaRecompuestaInterior`: por ASE y subtotal-interior, el conjunto de refs de `<f>` == celdas de las filas-firma del sub-bloque (L-1: exige literal/vacío según fixture). Sin evaluar valores, sin dominio nuevo. Corre donde corre `ExigirFormaRecompuesta` (post-`AjustarEnWorkbook`, rama espejo) + tests de regresión.

### 2.2 Dónde vive la reescritura interior (mutador vs mapa) y trade-offs

| Eje | Elegida | Descartada |
|---|---|---|
| Dónde | Mismo pase final post-5→1 (`RecomponerVisibles` extendido: posee geometría final + secuencias + `Reanclar` ya corrido) | Mapa de celdas por rol/ocurrencia (el patrón que falló dos veces: L-menores HU-12, `Aplicacion` agosto, visibles Plan 31) / writer (`omitirR1`: no conoce filas) |
| Cómo se localiza | Por firma `(C,D,E)` en el workbook mutado + alineación por orden con la fuente del sub-bloque | Por dirección congelada (se mueve con el delta) / por enésima ocurrencia (frágil a cardinalidad) |
| L-1/L-2 | Componer lo que el disco dice por celda (fixture manda) | Regla genérica "0 si falta" (mentiría donde el manual trae fórmula y viceversa) |
| `--L` | Reproducir patrón del template base | Normalizar (rompería f-vs-f contra el manual) |

### 2.3 Composición por sub-bloque (cardinalidad variable) + auditoría + contratos

**Regla de alcance del sub-bloque:** las filas-dato de un subtotal-empresa son las filas con firma de dato (`EsDatoEmpresa` o `EsMesTotal`/`EsAplicacionTotal` según el ancla que el fixture congele) físicamente entre el subtotal previo (o inicio de zona) y la fila-subtotal, en el workbook ya dimensionado; alineadas por orden con la secuencia fuente del mismo rango (misma cardinalidad por construcción del delta —si difiere, fail-fast ASE+celda, nunca composición parcial).

**Auditoría (D-D operativa):** cada celda interior recompuesta queda cubierta por DOS asserts: (i) texto `=` manual (fixture, f-vs-f); (ii) forma `=` firma del sub-bloque (gate-v2). Log Debug por celda (ASE + celda + texto antes/después), mismo formato que visibles T2.

| Contrato | Cambio |
|---|---|
| `OpenXmlEspejoR1Mutador.RecomponerVisibles` | Extensión al interior (sub-bloques por `(C,D,E)` + clases R-1/L-1/L-2 + log); dimensionado/`Reanclar` intactos |
| `FilaEspejoR1` (Core, puro) | Nueva firma `EsDatoEmpresa` (+ localizadores de subtotal por `(C,D,E)` si el diseño T1 los pide en Core); `EsMesTotal`/`EsAplicacionTotal` intactos |
| `WorkbookLeafCellMapQ2.R1Q2ProtectedPorAse` | Extensión documentada de la derogación Plan 31/T2 al interior (los fragmentos congelados a julio dejan de describir sub-bloques recompuestos); julio sigue validándose igual (D-E) |
| `ValidadorTotalesR1Workbook` | Nuevo `ExigirFormaRecompuestaInterior` (estructural, sin valores); S1 intacto |
| `InterioresR1Esperados` (+tests) | Nuevo fixture + drift-del-disco (espejo de `TotalesR1EsperadosTests`) |
| `ComparadorSalidaVsManualTests.BrechaDeclarada` | Sin cambios de código (el rojo R-D-5 se CIERRA por construcción: 96→0 inesperadas); si alguna interior exige divergencia permanente, se versiona con cita (no blanket) |
| `project-context.md` + manual | Doctrina interior-por-firma (D-A/D-B) + firma nueva (T0c) + H-DOBLE-SIGNO (D-G) + clases L (D-C) + gate-v2 (D-F) |

### 2.4 Architecture Validation Certificate

| Principio | Veredicto |
|---|---|
| **SRP** | ✅ Una razón de cambio por costura: el mutador compone geometría (visibles + interior), el gate-v2 contrasta forma, el fixture fija texto, el comparador clasifica. Ninguna rama `if agosto`: la cardinalidad sale de la fuente. |
| **OCP** | ✅ Una empresa nueva en C o un término nuevo en un sub-bloque componen sin código nuevo (catálogo C abierto, forma = función de la secuencia); julio-identidad garantiza que lo cerrado no se mueve. |
| **DIP** | ✅ El gate-v2 consume firmas puras de Core (+ fixture de tests); no crea dependencias nuevas (BCL/OpenXML ya presentes); cero agregados de dominio inventados. |
| **Best practices** | ✅ TDD rojo-primero (gate-v2 + comparador fallan en estado actual para el interior); fail-fast ASE+celda+firma; quincena por dominio; `<f>` intacto fuera del contrato; tolerancia ±0.5; insumos reales; sin Excel/COM. |
| **Performance** | ✅ Pase interior O(filas del bloque × subtotales) por ASE + verificación lineal; cero impacto en julio (delta 0: componer reproduce el texto existente). |

---

## 3. SPEC

### 3.1 Requirements

| ID | Requisito | Grupo | Criterio de aceptación |
|---|---|---|---|
| R-F-1 | Fixture `InterioresR1Esperados` congelado del manual (julio+agosto): celda+texto+cita por cada interior sustantivo (≈100) + literales L-1 con valor | F | Test-drift re-lee el disco y pasa; si el manual cambia, el test lo delata antes que el compositor |
| R-B-1 | Recomposición interior R-1 por sub-bloque: texto idéntico al manual + forma = firma del sub-bloque, 5 ASE × ambos períodos (incl. `F60` 3→2, `F204/F574` con `--L`, `F216` 3→2, `F335/F574` 2→3) | B | Texto-vs-manual por celda + gate-v2 por celda; agosto cierra, julio idéntico |
| R-B-2 | Caso `F217` y filas L-1 (ASE3 `F327/F328/F337/F338` + `G327/G337`/`H327/H337`): componer literal/vacío EXACTO del manual (el manual es quien no trae `<f>`) | B | Assert por celda: sin `<f>` + literal = fixture; gate-v2 exige literal donde el fixture dice literal |
| R-B-3 | Filas L-2 (`F464→F414`, `F584→F495`, `G584→G495`): componer la forma del sub-bloque aunque la app actual no traiga `<f>` | B | Texto-vs-manual por celda verde (la celda nace con la fórmula correcta) |
| R-B-4 | Julio-identidad interior: con geometría julio el pase reproduce byte-idénticas las fórmulas canónicas | B | 0-diff R1-interior julio (f-vs-f + literales) vs manual/golden; goldens ±0.5 intactos |
| R-B-5 | Excepción `<f>` acotada al contrato T2-interior y auditada término a término | B | Grep: ninguna `<f>` fuera del contrato cambia (diff de `<f>` por hoja = visibles Plan 31 + interior de este plan); log con ASE+celda+texto |
| R-G-1 | Gate-v2 `ExigirFormaRecompuestaInterior`: refs == filas-firma del sub-bloque (o literal según fixture); FAIL en agosto-actual, PASS en julio y post-T2 | G | TDD rojo-primero; fail-fast con ASE + celda + reales-vs-esperadas |
| R-G-2 | Comparador agosto: 96→0 divergencias inesperadas (cierre de R-D-5); julio sigue en 0 | G | `Salida5AseQ2_vs_Manual` verde ambos períodos sin nuevas brechas (o versionadas con cita si alguna interior lo exige) |
| R-R-1 | Suite base 443/444 verde antes y después (el único rojo previo se vuelve verde) + build 0 warnings; goldens julio ±0.5; DetRetri-D 5/5 vs R10 ambos períodos | Transversal | `dotnet build …slnx` + `dotnet test` en cada tarea |
| R-R-2 | `project-context.md` (doctrina + firma + H-DOBLE-SIGNO + clases L + gate-v2) + manual (proceso) | Transversal | Diff acotado; sin código en el mismo commit |

### 3.2 Scenarios (Given/When/Then)

- **S1 (TDD rojo — reproduce el interior):** Given la salida agosto actual (interior con anclaje julio, ej. `F60=F40+F21+F8-L8-L21`), When el gate-v2 verifica `F60`, Then FAIL con `ASE 1 Reporte Componentes R1!F60 [SUB_EMP]: refs [F8,F21,F40,L8,L21] != esperadas [F18,F37,L18]` (y PASS en julio con la forma canónica).
- **S2 (cardinalidad variable interior):** Given sub-bloque ENEL-ASE3 agosto, When el pase recompone `F335`, Then `<f>` = `F255+F240+F231-L231-L240` (texto = manual) y su forma = firma del sub-bloque.
- **S3 (`F217` literal):** Given `F217` (manual sin `<f>`), When el pase interior, Then la celda queda sin fórmula y con el literal del fixture; el gate-v2 la exige literal (fórmula → FAIL).
- **S4 (`--L`):** Given `F204`, When recomposición, Then `<f>` = `F133+F105+F87--L87-L105` (doble signo reproducido, texto = manual).
- **S5 (L-2):** Given `F464` (app actual sin `<f>`), When recomposición, Then `<f>` = `F414` (texto = manual).
- **S6 (julio identidad):** Given período 2026072 + base julio, When flujo 5-ASE Q2 + pase interior, Then diff R1-interior vs manual = 0 (fórmulas y literales) y suite goldens ±0.5 verde.
- **S7 (cierre del rojo):** Given `BrechaDeclarada` intacta (sin eximir R1-interior), When regresión agosto post-T2-interior, Then 0 inesperadas (R-D-5 cerrada); pre-T2-interior, Then las 96 en rojo (S1 a escala suite).

---

## 4. TASKS

| ID | Tarea | Grupo | Depende de | Verificación |
|---|---|---|---|---|
| T1 | Fixture-que-delata-primero (TDD) + gate-v2 en rojo. (a) Congelar del MANUAL los textos del interior (`InterioresR1Esperados`: F/G/H/L sustantivas + literales L-1 con valor, julio+agosto, con cita; si el vivo difiere de T0a, el fixture manda y T0a se enmienda). (b) Congelar firmas de sub-bloque (`EsDatoEmpresa` + formas por `(C,D,E)`) desde manual+fuente. (c) Nuevo `ExigirFormaRecompuestaInterior` con tests S1 (rojo agosto-actual, verde julio). El comparador sigue en rojo R-D-5 (INTENCIONAL hasta T2). | F+G | — | R-F-1, R-G-1 (S1); build 0 warnings; resto de la suite 443/444 (solo lo nuevo en rojo + 96 vivas) |
| T2 | Recomposición interior por firma + sucesor del gate. Extensión de `RecomponerVisibles` (sub-bloques por `(C,D,E)`; R-1/L-1/L-2; patrón `--L` del template base; log ASE+celda+texto) + derogación documentada de `R1Q2ProtectedPorAse` para el interior + gate-v2 en el pipeline donde corre `ExigirFormaRecompuesta`. Tests: texto-vs-manual por celda (agosto, incl. S2/S3/S4/S5) + identidad julio (S6). | B | T1 | R-B-1..R-B-5 (S2/S3/S4/S5/S6); S7 post-T2 verde (96→0); goldens ±0.5; DetRetri-D 5/5 ambos períodos |
| T3 | Docs: `project-context.md` (doctrina interior-por-firma D-A/D-B + firma nueva T0c + H-DOBLE-SIGNO D-G + clases L D-C + gate-v2 D-F) + manual (cómo leer un FAIL del gate-v2). | Transversal | T2 | R-R-2; diff acotado; sin código en el mismo commit |
| T4 | Cierre: build 0 warnings + suite completa (443/444 base + nuevos, TODO verde) + goldens julio ±0.5 + DetRetri-D 5/5 vs R10 ambos períodos + grep de cierre (diff de `<f>` = visibles Plan 31 + interior de este plan; cero literales nuevos de período/empresa) + auditoría S6/S7 + CRLF + sin commits. | Transversal | T1..T3 | R-R-1; DoD §5 |

**Orden sugerido (TDD):** T1 (rojo que reproduce el interior + fixture que manda) → T2 (recomposición que lo cierra) → T3 → T4. Build 0 warnings antes de marcar cada tarea completa. Sin commits (los hace el Ingeniero).

---

## 5. Resumen ejecutivo, riesgos y decisiones

- **Resumen:** el T0 (§0.1) inventarió el interior completo desde el disco: ~100 celdas sustantivas en 3 clases (R-1 offset-recomponible —la gran mayoría, mismo patrón que los visibles—; L-1 literal-en-manual con `F217` como caso testigo y las filas `F327/F328/F337/F338` de ASE3-0Aplic; L-2 fórmula-ausente-en-app por clon vacío con `F464`/`F584` como testigos), con dos hallazgos no obvios: la firma existente NO alcanza (anclas mixtas: `EsMesTotal`/`EsAplicacionTotal` cubren parte; las filas `C=<empresa> ∧ D='Total'` exigen firma nueva con catálogo abierto —el disco trae `NUEVO ESQUEMA`—) y el quirk `--L` (`F204/F574`) debe reproducirse, no normalizarse. Julio trae 0 diffs de interior (la red de reversión está intacta). El plan convierte esa evidencia en: fixture que manda (T1), recomposición en el mismo pase final (T2), gate-v2 estructural + comparador como red numérica (D-F), y docs mínimos.
- **Riesgo principal:** R-FORMA-INT (la forma interior recompuesta diverge del manual en un borde y la excepción `<f>` crea divergencia silenciosa). Contenido por triple red: fixture congelado antes de componer, doble assert por celda (texto + forma), y julio-identidad como invariante de reversión.
- **Decisión para el Ingeniero:** ratificar D-A..D-G (§0.3). **Forks con recomendación:** D-A (mismo pase final del mutador vs pase/mapa nuevo —recomendado mismo pase: posee la geometría; el writer con `omitirR1` no la conoce) y D-F (gate-v2 estructural + comparador como red numérica vs gate numérico interior —recomendado estructural: el dominio no trae agregados por empresa e inventarlos violaría el rector). **Sin preguntas bloqueantes pendientes.**

### DoD (Definition of Done)

1. `dotnet build Remuneracion.WinForms/Remuneracion.WinForms.slnx` con 0 warnings.
2. Suite completa verde sin regresión (base 443/444 + nuevos, TODO verde: el rojo R-D-5 se cierra) + goldens Capa A Q1+Q2 ±0.5 + DetRetri 5/5 vs R10 (julio y agosto) + 0-diff R1 julio (visibles + interior).
3. Agosto 2026082 end-to-end con interior R1 recompuesto = manual en texto (5 ASE, incl. cardinalidades 2↔3, `--L`, `F217` literal, L-2 nacidas) y forma = firma del sub-bloque (S1/S2/S3/S4/S5/S7).
4. `project-context.md` + manual actualizados; grep de cierre: diff de `<f>` = visibles Plan 31 + interior de este plan.
5. Sin commits del agente; finales de línea CRLF; sin emojis.

### Follow-ups explícitos (fuera de este plan)

- F-T4-2 (remesh R2/R4), F-T4-3 (re-encender SALDOS/AJUSTES/CONSOLIDADO/REMUNERACION_*), F-T4-4 (`INTERVENTORIA!R26`): vivos, intactos.
- R4 espejo-motor, R4-por-empresa Q2, CF/DV del reanclaje (W-4), `R-EXTRA-CONCEPTO` (Plan 23), validación de nombres de hoja en preflight: vivos, intactos.
- Todo frente con veredicto "requiere código": su propio T0 antes de cualquier línea.
- Retiro/archivo de evidencias T0 (este §0.1 vive en el plan; los dumps estáticos `%TEMP%\opencode\dump*.ps1` no se versionan).
