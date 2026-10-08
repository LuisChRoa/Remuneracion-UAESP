# Plan 31 — Recomposición por firma de los totales visibles R1 al dimensionar el espejo + gate workbook-vs-dominio + sello D6/D7 + decisión R2/R4-soporte

> **Alcance:** cuatro frentes en orden de impacto: (1) **recomposición por firma** de los totales visibles de `Reporte Componentes R1` al dimensionar el espejo (Opción B del T0; Opción A como red mínima); (2) **gate workbook-vs-dominio sobre R1** (Opción C) + corrección de la exención falsa del comparador en agosto; (3) **sello `DetRetri/DetValiRetri D6/D7`** (Fecha de Proceso/Hora) extendiendo el mecanismo vigente de fechas; (4) **decisión R2/R4-remesh + hojas de soporte** (solo documenta y decide; el código que salga es follow-up con su propio T0). Cero cambios de fórmulas de negocio fuera de la recomposición declarada; goldens julio ±0.5, DetRetri-D 5/5 vs R10 y suite 406/406 intactos.
> **Rector (IRRENUNCIABLE):** `README.md` + `Docs/Propuesta_Proyecto_Automatizacion_Remuneracion_UAESP.md` + `Docs/Detalle de plantilla.docx` + `Docs/Prompt_Maestro_Proyecto_Remuneracion_UAESP _ Vo .docx` + `.opencode/project-context.md`. **Invariantes del proyecto:** NUNCA sobrescribir fórmulas —salvo la excepción legítima, explícita y acotada de este plan (recomposición de geometría del espejo, D-B/D-D)—; tolerancia ±0.5; ASE4 sin columna "Especiales" (0 si ausente); Q2 aplica SALDOS POR NOTA/RETRIBUCION NEGATIVA (`Periodo.NumeroQuincena == 2`, `AjustesSfT = 0` en Q1); redondear a entero antes de DetRetri/columna D; build 0 warnings; finales de línea CRLF; fail-fast nombrando archivo+celda; tests SOLO con insumos reales (`Docs/Prueba2/Insumos`, `Docs/Prueba Julio-2/Insumos`); sin commits (los hace el Ingeniero con `#commit`); sin emojis. **Quincena por dominio, nunca por contenido.**
> **Continuidad:** HU-01..HU-23 cerradas; Planes 21 (espejo R1), 23, 25 (roles por firma + totOpt generalizado), 26 (preflight), 27 (firma + hoja-por-nombre), 28 (calcChain + sello G7/K7), 29 (paridad app-vs-manual), 30 (naming dinámico + base canónica agosto, suite 406/406) cerrados. Este plan NO reabre lectura, cálculo, espejo-dimensional, preflight, sello G7/K7 ni paridad fuera de lo declarado: **julio = identidad (0-diff R1) por construcción.**
> **Insumos canónicos (no se re-fijan, no se inventan datos):** T0 cerrado `plans/31-T0-Evidencia-R1.md` (276 líneas, método zip+XML BCL, sin Excel/COM); base canónica agosto `Docs/Prueba2/Plantilla_Remuneracion_2026082.xlsx` (Plan 30/T3); manual agosto `Docs/Prueba2/Resultado/Resultado Manual por el administrativo/Remuneracion 202608-2 Total_7721.xlsx`; app agosto `Docs/Prueba2/Resultado/Resultado3/Remuneración 202608-2 Total.xlsx` (R2 y raíz con la misma fórmula); canónico julio `Docs/Plantilla_Remuneracion.xlsx` + golden `Docs/Insumos/Remuneracion 202607-2 Total.xlsx`; fuentes `Docs/Prueba2/Insumos/{1..5}-*/Recaudoporcomponente_*_16082026_*.xlsx`; R10 `Docs/Prueba2/Insumos/R10_Remuneracion_2026082.xlsx` (oráculo); T0b `plans/29-T0-Evidencia-B.md` (deriva de malla R2/R4 + orígenes DetRetri).
> **Numeración:** `plans/` 01..30 ocupados; este plan toma el primer correlativo libre, **31** (el informe `31-T0-Evidencia-R1.md` es evidencia previa, no el plan).
> **Estado:** CERRADO (2026-10-07) — T1→T2→T3→T4→T5 ejecutados y verificados: build 0 warnings, suite 443/444 (único rojo intencional: comparador agosto con 96 R-D-5), gate S1 verde ambos períodos, sello D6/D7 como texto, veredictos R-D-1..R-D-7 + follow-ups F-T4-1..4, docs actualizados, sin commits. — planificado desde T0 cerrado, sin implementar.
> **Fecha:** 2026-10-07

---

## 0. Clarification Gate

**Sin preguntas bloqueantes: el plan viene redactado con las decisiones D-A..D-G (§0.3) y T0 ya cerrado (§0.1).** Los dos puntos con recomendación (dónde vive la recomposición, D-A; origen del sello D6/D7, D-F) vienen redactados en la ruta recomendada con su verificación de cierre dentro de la tarea. Punto de bloqueo respetado: NO se implementa nada en este documento.

### 0.1 T0 reutilizado (causa cerrada, no re-investigar) + qué se re-verifica

Cada fila: afirmación del informe → se reutiliza tal cual → qué la re-verifica en este plan (el test, no la mano).

| # | Afirmación T0 (cita) | Reutilización | Re-verificación en este plan |
|---|---|---|---|
| E1 | El espejo dimensiona por conteo (`delta`, `:157`), borra en el PIE con `primeraBorrada = totalRowIdx - m` (`:183-211`) y reancla solo `r >= primeraBorrada`; la fuente recorta en la CABEZA (Mes 12/32/48→9/29/45); datos por orden bien, fórmulas ancladas a julio (§0, §1.4) | Causa raíz: no se re-investiga | T1: el gate reproduce el anclaje viejo (F50=F32+F47+F12) y falla; T2 lo cierra |
| E2 | Tabla de fórmulas TOT_OPT app vs manual por ASE (§1.2 ASE1, §1.5 ASE2-5) + G (`G50=G29+G9`, v=975168318.98) + EXTEMP (`F52=F34+F14-L14`) | Textos del manual = valores esperados de los tests T1/T2 (transcripción exacta, sin recalcular) | T1 congela lo NO tabulado: celda G y celda EXTEMP por ASE 2-5 + G-ASE5-3Mes + sub-visibles de empresa (M50/M52-style, §4.5) leídos del manual; si difieren, el test lo delata |
| E3 | Los DATOS son idénticos App vs Manual fila a fila (§1.3: F29=17.977.519.904,88 en ambas; es anclaje, no lectura) | Aislante del bug: la recomposición NO toca valores, solo `<f>` de visibles | T2: el gate evalúa las fórmulas recompuestas contra literales ya escritos (resuelven, no son 0) |
| E4 | Los 4 gates ciegos (§2): validación salteada con `espejoDesplazado=true` (`OpenXmlPlantillaWriter.cs:277-286, 330-339`); coherencia dominio-a-dominio (`WorkbookLeafCoherence.cs:22-58, 85-113, 180-220`); oráculo sin R1 (`ValidacionOracleReader.cs:28-69, 98-128`); exención blanket del comparador (`ComparadorSalidaVsManualTests.cs:172-185`) con premisa falsa (geometrías idénticas, §1.3) | Mapa de huecos: cada uno se cierra nominalmente | T1: gate nuevo FAIL-agosto/PASS-julio + R1 de vuelta al comparador (delata en rojo hasta T2) |
| E5 | ASE5-agosto trae 3 `Mes` (494/525/548) vs 2 términos en plantilla (`F558=F552+F531-L531`): el reanclaje mecánico es insuficiente (§1.5, Opción B) | Justificación de la recomposición por firma como fix real (no solo Opción A) | T2/S2: `F554=F548+F525+F494-L494-L525` texto + evaluación = totOpt dominio ±0.5 |
| E6 | Dominio correcto aislado (§1.7): `DetRetri2026082 D9:D14` app = manual al entero; `CONSOLIDADO!D9='R1'!F50` idéntica en texto; `MapearR1Q2` (`ExcelDataReaderWorkbookLeafInputReader.cs:1083-1090`) agrega por firma | Agregados `totOpt/extemp` = oráculo del gate y de la recomposición (fuente única, Plan 25) | T1/T2: sin cambios en el reader; el gate consume `leaf.R1.TotalOportunoEsperadoPorAse` / `ExtemporaneoEsperadoPorAse` |
| E7 | D6/D7 Det/DetVali: base 2026082 trae `D6=04/08/2026`, `D7=10:15 AM` (julio) vs manual `02/09/2026` / `07:42 AM`; `CONSOLIDADO G7/J7` ya bien; `CONSOLIDADO D6`=46267 por T3 (`BaseConsistente2026082Tests.cs:113-114`) | Dato de partida del frente 3 (aporta el encargo; fechas 02/09 y 04/08 coinciden con `CONSOLIDADO D6` 46267/46238 del Plan 30/V7 — hipótesis: mismo sello de proceso) | T3 mini-T0 (NUEVA verificación, lectura zip+XML): literal-vs-`<f>` en base+manual+R10 y congelamiento del origen; sin origen legible no hay sello (fail-fast, no invento) |
| E8 | T0b: deriva de malla R2/R4 (`29-T0-Evidencia-B.md` §2: R2-detalle agosto solo {1,3,5}; R4-detalle agosto solo {2,5}; columnas por encabezado) + manual que recompone geometría | Tabla de partida del frente 4 | T4: veredicto por frente (adoptar / divergencia / fuera-de-alcance), solo docs; lo que pida código sale como follow-up con su propio T0 |

**NEEDS_CONTEXT heredados del T0 §4 (no bloquean; quedan documentados):** (1) números del encargo "Recaudo Base / Remuneración Final" no localizados en ningún xlsx — fuera de alcance, la evidencia verificada es el anclaje R1; (2) plantilla exacta de la corrida R3 (alta confianza: base 2026082, sin log que lo confirme); (3) forma exacta de los visibles — SE CIERRA en T1 (E2); (4) CF/DV del bloque mutado (W-4, follow-up vivo); (5) barrido de consistencia post-fix en hojas que referencian visibles R1 (`REMUNERACION_*`, `GERENTES_*`, `VALIDACION_*`, `DetValiRetri`) — entra como aserto del gate T1/T2 (si el recálculo arrastra, el comparador lo muestra); (6) mapa aritmético exacto por ASE — innecesario (la recomposición no lo usa).

### 0.2 Mapeo al Rector (Sale / No sale)

**Sale (EXPLÍCITO):**
- Re-investigar la causa (E1 cerrada) o re-leer geometrías ya tabuladas (E2) fuera de la re-verificación E2/T1.
- Tocar el lector (`MapearR1Q2`, firmas `FilaEspejoR1`, `BloqueEspejoAseInputs`), el cálculo (`CalculoRemuneracion`, `ProcesadorPeriodo` salvo el paso del sello), el sello G7/K7, `N3`/`C6`/`CONSOLIDADO D6`, columnas por quincena, BCE, preflight, naming dinámico, calcChain.
- Reescribir `<f>` fuera de las celdas visibles R1 declaradas en el contrato T2 (el resto del workbook sigue intacto; la autorización vive SOLO aquí, §2.3/D-D).
- Depender de cachés `<v>` en el gate o el comparador (método D-E del Plan 29: stale por diseño; solo texto-fórmula + literales).
- Inventar el origen de D6/D7 o de cualquier frente-4 (fail-fast o divergencia documentada, nunca dato simulado).
- Julio con diff ≠ 0 en R1 (la recomposición en julio es identidad; si difiere, es regresión).
- Paquetes NuGet; commits (los hace el Ingeniero con `#commit`).

**No sale (entra, alcance congelado en 4 frentes):**
| Frente | Superficie de cambio |
|---|---|
| G — gate + comparador | Nuevo gate workbook-vs-dominio R1 (Infrastructure, BCL/OpenXML en lectura) + corrección de la exención `ComparadorSalidaVsManualTests.cs:172-185` (R1-agosto vuelve a compararse; resto de hojas según veredicto T4) + tests |
| B — recomposición por firma | Pase final en `OpenXmlEspejoR1Mutador` (post-5→1) que recompone `<f>` de visibles R1 por firma + sucesor del gate de protegidas salteado + actualización/relajación documentada de `R1Q2ProtectedPorAse` + tests |
| S — sello D6/D7 | Mini-T0 de origen + `ProcesadorPeriodo`/`ResultadoRemuneracion` + extensión de `EscribirFechasPeriodo` a 4 celdas (nombres por `NombresHojaPeriodo`) + tests |
| D4 — decisión R2/R4+soporte | Solo docs: veredicto por frente (adoptar-geometría / declarar-divergencia / fuera-de-alcance) + follow-ups con T0 propio; cero diff de código |

### 0.3 Decisiones ejecutivas (aprobación = aceptar estas)

| ID | Decisión |
|---|---|
| D-A | **La recomposición vive en el mutador, en un pase final post-5→1** (no en el writer, no por bloque). El mutador posee la geometría final; el writer escribe con `omitirR1` y no la conoce. El pase final re-localiza cada bloque, alinea secuencia-fuente ↔ filas finales por orden (misma cardinalidad por construcción) y compone `<f>` por firma. Descartado: recomponer por bloque dentro de `AjustarBloque` (los `Reanclar` posteriores de bloques superiores reescribirían lo recompuesto) y en el writer (no posee ni filas ni firmas). |
| D-B | **Opción B (recomposición por firma) es el fix real; Opción A (reanclaje mecánico de borde) queda como red mínima.** El `Reanclar` con mapa `primeraBorrada` (`:210`) se conserva para todo lo NO-visible (refs externas tipo `CONSOLIDADO!D9='R1'!F50`, nombres, merges): la fila del total se movió bien (53→50) y esas refs son desplazamientos físicos correctos. Las visibles se recomponen DESPUÉS, por firma, con cardinalidad real (ASE5-3Mes imposible por mapa). |
| D-C | **Gate workbook-vs-dominio (Opción C) obligatorio como red, con TDD rojo-primero.** El gate lee el workbook generado, evalúa cada visible (texto-fórmula + literales, nunca `<v>` stale) y lo contrasta contra el dominio por firma ±0.5: debe FALLAR en agosto-actual y PASAR en julio. La R1-agosto vuelve al comparador con exclusiones versionadas legítimas (shared-formula H1/T0g, token-período, columnas de otra quincena V3), no blanket. El rojo intencional T1→T2 es evidencia, no regresión. |
| D-D | **Excepción legítima y acotada al "NUNCA sobrescribir fórmulas".** La reescritura aquí es composición de geometría del espejo (el informe la fundamenta: sin ella ASE5 pierde un término), limitada a las celdas visibles del contrato T2, con forma `ΣF(Mes)−ΣL(menos última)` / `ΣF(Aplic)−Esp(primera)` idéntica al dominio (Plan 25), y auditada término a término (texto vs manual + evaluación vs dominio). Fuera de esas celdas, el invariante sigue intacto. |
| D-E | **Julio-identidad como invariante de cierre.** Con geometría julio (delta 0 efectivo) la recomposición reproduce byte-idéntica la fórmula canónica (`F53=F32+F48+F12-L12-L32`, …): goldens ±0.5 y 0-diff R1 julio son la red en cada tarea. Si julio difiere, se revierte el pase, no el golden. |
| D-F | **Sello D6/D7 por el mecanismo vigente de fechas, con origen congelado por mini-T0.** Se extiende `EscribirFechasPeriodo` (valores OADate, estilo preservado, guard anti-fórmula) a `DetRetri*/DetValiRetri* D6/D7` (nombres por `NombresHojaPeriodo`, fail-fast `ERR-FORMATO-FUENTE`/`ERR-PLANTILLA` si la fuente no trae fecha legible o la celda es fórmula). Hipótesis de trabajo (a congelar, no a asumir): mismo sello de proceso que `CONSOLIDADO D6`. `CONSOLIDADO D6` no se duplica (ya 46267). |
| D-G | **Frente 4 solo decide y documenta.** Por frente (R2-remesh, R4-remesh, SALDOS POR NOTA detalle, ANT EXT-REV, ANTICIPOS, DetRetri E..O por cadena): adoptar-geometría / declarar-divergencia / dejar-fuera-de-alcance, con cita a T0b y sin inventar datos. Todo lo que pida código sale como follow-up con su propio T0 (no entra ni una línea de producción en T4). |

---

## 1. PROPOSE

### 1.1 Intent

Que los totales visibles de `Reporte Componentes R1` se deriven de las filas reales del período (firma, no conteo ni borde), que ningún total mal anclado vuelva a entregarse en silencio (gate que hoy fallaría), que el sello de proceso cubra `DetRetri/DetValiRetri D6/D7`, y que el remesh R2/R4+soporte quede decidido por frente — con julio byte-idéntico.

### 1.2 In Scope

- Gate workbook-vs-dominio R1 (lectura del generado + evaluación BCL + contraste ±0.5) con TDD rojo-primero; R1-agosto de vuelta al comparador.
- Recomposición por firma de visibles R1 (TOT_OPT F + total TDF G + EXTEMP F por ASE, incl. cardinalidad variable y 0-Aplic) + sucesor del gate de protegidas + contrato de forma.
- Sello D6/D7 (mini-T0 de origen + 4 celdas + fail-fast).
- Tabla de veredictos frente-4 + follow-ups.
- Docs: `project-context.md` (doctrina recomposición-por-firma + excepción `<f>` + gate + sello extendido + veredictos) + manual (proceso).

### 1.3 Non-Goals

Ver §0.2 "Sale". En particular: NO se generaliza la recomposición a otras hojas (R2/R4-detalle, CONSOLIDADO, validaciones: sus `<f>` siguen intactos); NO se toca el dominio numérico (ya correcto, E6); NO se valida contenido en preflight (follow-up vivo).

### 1.4 Riesgos

| ID | Riesgo | Mitigación |
|---|---|---|
| R-FORMA | La forma recompuesta diverge de la del manual en un borde (G-ASE5-3Mes, EXTEMP 0-Aplic, sub-visibles de empresa) y la excepción `<f>` crea una divergencia silenciosa nueva | E2 congela lo no tabulado antes de componer; gate C + comparador f-vs-f + julio-identidad en cada tarea; 0-Aplic = literal 0 auditado (decisión T2, no improvisación) |
| R-JULIO | El pase final altera julio (delta 0) y rompe goldens | D-E: identidad byte-idéntica exigida; goldens Q1+Q2 ±0.5 + 0-diff R1 como red por tarea; ante diff se revierte el pase |
| R-GATE-CIEGO | El gate evalúa cachés stale o refs sin resolver y pasa en falso | Método D-E Plan 29: prohibido `<v>`; evaluación desde texto-fórmula + literales con refs resueltas y fail-fast si una ref no resuelve (archivo+celda) |
| R-SELLO | El origen de D6/D7 no es el supuesto y se sella un dato inventado | D-F: mini-T0 congela el origen; sin origen legible → fail-fast, nunca sello por defecto |
| R-BLANKET | Al re-encender R1 en el comparador aparecen diffs legítimos no versionados y se re-exime en bloque | D-C: solo exclusiones versionadas con cita (H1/T0g, token-período, V3); cada diff nuevo se clasifica o se fija, nunca blanket |
| R-DERIVA-4 | El frente 4 presiona a "arreglar de paso" con código sin T0 | D-G: T4 con cero diff de producción (verificado por diff); follow-ups con T0 propio |

---

## 2. DESIGN

### 2.1 Enfoque: recomponer, no reanclar; gatear, no eximir; sellar por mecanismo, no por_LITERAL; decidir, no parchar

**Frente G (gate + comparador).** Nuevo gate de post-escritura en Infrastructure (lector BCL/OpenXML en modo lectura, sin Excel/COM): por ASE, (a) lee el texto `<f>` de los visibles R1 del workbook generado; (b) parsea refs de celda (mismo regex de `ComparadorSalidaVsManualTests:41`); (c) resuelve cada ref a su literal en el mismo workbook (fail-fast archivo+celda si no resuelve o si la celda es fórmula); (d) evalúa la suma lineal (`+F…−L…`) y la contrasta contra `leaf.R1.TotalOportunoEsperadoPorAse` (F) y `ExtemporaneoEsperadoPorAse` (EXTEMP) ±0.5; (e) para G verifica que el conjunto de refs == celdas G de las filas `EsMesTotal` (firma, sin agregado de dominio que lo cubra). El gate corre donde hoy se salta la validación (post-`AjustarEnWorkbook`, rama `espejoDesplazado`) y en los tests de regresión. En paralelo, `BrechaDeclarada` pierde el blanket R1-agosto (`:179-184`): `REPORTE COMPONENTES R1` vuelve a compararse f-vs-f + literales; el resto de hojas del blanket pasa a veredicto T4 (se re-enciende lo adoptado, se versiona con causa lo divergente).

**Frente B (recomposición).** Tras el loop 5→1 de `AjustarEnWorkbook` (`OpenXmlEspejoR1Mutador.cs:99-120`) y antes de `hoja.Save()` (`:117`), pase final por ASE (1→5, orden irrelevante: ya no hay más `Reanclar`): re-localizar el bloque (`LocalizarBloque`, `:225-288`), alinear `bloque.Filas[i] ↔ dataRows[i]` por orden (misma cardinalidad por construcción del delta) y componer el texto `<f>` de cada visible del contrato con las filas reales: TOT_OPT F = `ΣF(Mes_k) − ΣL(Mes_k, k<último)`; total TDF G = `ΣG(Mes_k)`; EXTEMP F = `ΣF(Aplic_k) − Esp(Aplic_0)` (0 Aplic → literal `0` auditado, coherente con "ausente = 0 explícito" del Plan 25). Las filas-ancla se identifican por `FilaEspejoR1.EsMesTotal` (`:77-79`) / `EsAplicacionTotal` (`:86-88`) sobre la secuencia fuente (patrón Plan 25, NO conteo ni borde). El `Reanclar` mecánico (`:568-642`) se conserva intacto para el resto (refs externas, nombres, merges). Tras recomponer, el sucesor del gate salteado valida las protegidas por FORMA recompuesta (fragmentos = filas reales, no `R1Q2ProtectedPorAse` congelado a julio).

**Frente S (sello).** Mini-T0 zip+XML (base + manual + R10, ambos períodos): ¿D6/D7 de `DetRetri*/DetValiRetri*` son literales o `<f>`? ¿De dónde sale 02/09/2026 + 07:42 AM? Con origen congelado, `ProcesadorPeriodo` propaga (mismo fail-fast `ERR-FORMATO-FUENTE` del Plan 28 si ilegible) y `EscribirFechasPeriodo` (`OpenXmlPlantillaWriter.cs:1634-1652`) sella las 4 celdas como seriales OADate (D6 fecha, D7 hora-fracción) con estilo preservado y guard anti-fórmula. Nombres de hoja por `NombresHojaPeriodo` (Plan 30); single-ASE sin fuente = no-op (doctrina Plan 28).

**Frente D4 (decisión).** Re-lectura de T0b §2 + estado actual del código/docs por frente, veredicto en tabla (adoptar-geometría / declarar-divergencia / fuera-de-alcance) con cita y consecuencia (comparador: re-encender o versionar; código: follow-up con T0 propio o nada).

### 2.2 Dónde vive la recomposición (mutador vs writer) y trade-offs

| Eje | Elegida | Descartada |
|---|---|---|
| Dónde | Pase final en `OpenXmlEspejoR1Mutador.AjustarEnWorkbook` (posee geometría final + secuencias + `Reanclar`; corre en ambos caminos de guardado que mutan + `Ajustar` standalone L85) | Writer (`EscribirCeldasLeafPorAse` con `omitirR1`: no conoce filas finales ni firmas; duplicaría localización) / por bloque dentro de `AjustarBloque` (los `Reanclar` de bloques superiores aún pendientes reescribirían lo recompuesto) |
| Cuándo | Después del loop 5→1, antes de `hoja.Save()`; el gate C corre después (post-escritura, donde hoy se salta) | Recomponer antes de dimensionar (las filas aún no existen) / en validación (validar no escribe) |
| Qué refs se reescriben | Solo visibles del contrato T2; el resto sigue el mapa mecánico (D-B: red mínima, refs externas correctas por desplazamiento físico) | Re-derivar TODO por firma (rework del mutador, riesgo en nombres/merges; innecesario: fuera del bloque nada cambió de cardinalidad) |
| 0-Aplic | Literal `0` auditado en EXTEMP (dominio = 0; fórmula con refs a no-Aplic mentiría) | Conservar refs del template (apuntan a filas de otro rol: el bug actual) / fórmula vacía (inválida) |

### 2.3 Composición de términos (cardinalidad variable) + auditoría de `<f>` + contratos

**Plantilla de composición** (idéntica al dominio `MapearR1Q2 :1083-1090`, por construction no por coincidencia):
- `TOT_OPT_F(ase) = +F(m0)+F(m1)[+F(m2)…] −L(m0)[−L(m1)…]` (todos los `Mes` en F; `L` de todos menos el último; ASE5-julio con 2 Mes y `Lmes1=0` ya cierra; ASE5-agosto con 3 Mes compone 3+2 = manual).
- `TOT_TDF_G(ase) = +G(m0)+G(m1)[+G(m2)…]` (sin términos L: verificado E2; G-ASE5-3Mes se congela en T1).
- `EXTEMP_F(ase) = +F(a0)[+F(a1)…] −Esp(a0)` (todas las `Aplicacion` en F menos Especiales de la primera; 0 filas → literal `0`).
- ASE4 (Bogotá Limpia, sin `Especiales`): la forma es la misma; los literales L resuelven 0 (doctrina vigente, sin rama especial en el texto).
- Orden de términos: el del template/manual (F ascendente por fila: `F29+F45+F9` se escribe en el orden canónico que T1 congele del manual, no en orden de descubrimiento).

**Auditoría de `<f>` (D-D operativa):** cada visible recompuesto queda cubierto por DOS asserts independientes: (i) igualdad de texto de fórmula vs manual (mismo método f-vs-f del comparador, término a término); (ii) evaluación BCL de ese texto contra literales == dominio por firma ±0.5 (el gate). La recomposición además se registra en el log de corrida (ASE + celda + texto, nivel Debug existente) para trazabilidad sin abrir el xlsx.

| Contrato | Cambio |
|---|---|
| `OpenXmlEspejoR1Mutador` | Pase final post-5→1 (re-localizar + alinear por orden + componer `<f>` por firma) + log; `Reanclar`/dimensionado intactos |
| `WorkbookLeafCellMapQ2.R1Q2ProtectedPorAse` (`:115-143`) | Sus fragmentos congelados a julio dejan de describir agosto: se documenta la derogación para visibles recompuestos y el sucesor valida por forma recompuesta (fragmentos = filas reales); julio sigue validándose igual (D-E) |
| `OpenXmlPlantillaWriter` (gates `:283-286, :336-339`) | En rama `espejoDesplazado`, la validación salteada se sustituye por gate C (workbook-vs-dominio) + validación de forma recompuesta; rama `!espejoDesplazado` intacta |
| `WorkbookLeafCoherence` / `ValidacionOracleReader` | Sin cambios (siguen dominio-a-dominio; el gate C es pieza nueva, no parche de ellos) |
| `ComparadorSalidaVsManualTests.BrechaDeclarada` (`:172-185`) | Fin del blanket R1-agosto; resto de hojas del blanket → veredicto T4 |
| `ProcesadorPeriodo` / `ResultadoRemuneracion` / `EscribirFechasPeriodo` | Propagación + sello D6/D7 (solo Frente S) |
| `project-context.md` + manual | Doctrina recomposición-por-firma + excepción `<f>` + gate C + sello extendido + veredictos D4 |

### 2.4 Architecture Validation Certificate

| Principio | Veredicto |
|---|---|
| **SRP** | ✅ Una razón de cambio por costura: el mutador compone geometría, el gate contrasta workbook-vs-dominio, el reader calcula dominio, el comparador clasifica. Ninguna rama `if agosto`: la cardinalidad sale de la fuente. |
| **OCP** | ✅ Un 4.º `Mes` o un 3.er `Aplicacion` futuro compone sin código nuevo (la forma es función de la secuencia); julio-identidad garantiza que lo ya cerrado no se mueve. |
| **DIP** | ✅ El gate consume agregados de dominio existentes (`TotalOportunoEsperadoPorAse`/`ExtemporaneoEsperadoPorAse`) y firmas puras de Core (`EsMesTotal`/`EsAplicacionTotal`); no crea dependencias nuevas (BCL/OpenXML ya presentes). |
| **Best practices** | ✅ TDD rojo-primero (el gate falla en estado actual); fail-fast archivo+celda en refs no resueltas y fechas ilegibles; quincena por dominio; `<f>` intacto fuera del contrato; tolerancia ±0.5; insumos reales; sin Excel/COM en pipeline ni tests. |
| **Performance** | ✅ Pase final O(filas del bloque × visibles) por ASE + evaluación lineal del gate; cero impacto en el path julio (delta 0: componer reproduce el texto existente). |

---

## 3. SPEC

### 3.1 Requirements

| ID | Requisito | Grupo | Criterio de aceptación |
|---|---|---|---|
| R-G-1 | Gate workbook-vs-dominio R1: evalúa visibles (texto-fórmula + literales, prohibido `<v>`) vs dominio por firma ±0.5; FAIL en agosto-actual, PASS en julio | G | Test S1 rojo-agosto / verde-julio; fail-fast con ASE + celda + esperado-vs-evaluado si una ref no resuelve |
| R-G-2 | R1-agosto vuelve al comparador con exclusiones versionadas legítimas (H1/T0g, token-período, V3), no blanket | G | `BrechaDeclarada` sin el bloque R1-agosto; el comparador delata F50/G50/EXTEMP hasta T2 y cierra tras T2 |
| R-B-1 | Recomposición por firma de TOT_OPT F + total TDF G + EXTEMP F por ASE, con cardinalidad real (ASE5-3Mes → 3+2; ASE3-0Aplic → EXTEMP literal 0) | B | Texto idéntico al manual (E2) + evaluación = dominio ±0.5, 5 ASE × ambos períodos |
| R-B-2 | Julio-identidad: con geometría julio el pase reproduce byte-idéntica la fórmula canónica | B | 0-diff R1 julio (f-vs-f + literales) vs golden/manual; goldens ±0.5 intactos |
| R-B-3 | Reanclaje mecánico conservado como red para lo no-visible; visibles recompuestos en pase final post-5→1 | B | Refs externas (`CONSOLIDADO!D9='R1'!F50` y cadena) intactas en texto; visibles con forma recompuesta |
| R-B-4 | Excepción `<f>` explícita, acotada al contrato y auditada término a término | B | Grep: ninguna `<f>` fuera del contrato cambia (diff de `<f>` por hoja = solo visibles R1); log con ASE+celda+texto |
| R-S-1 | Origen de D6/D7 congelado por mini-T0 (literal-vs-`<f>`, base+manual+R10, ambos períodos) | S | Tabla de origen con archivo+hoja+celda; sin origen legible no hay sello |
| R-S-2 | Sello `DetRetri*/DetValiRetri* D6/D7` (nombres por `NombresHojaPeriodo`) como valores OADate, estilo preservado, guard anti-fórmula; fail-fast si fuente ilegible | S | Salida agosto con D6/D7 = sello del manual ±redondeo de hora; fuente ilegible → `ERR-FORMATO-FUENTE` archivo+celda |
| R-S-3 | `CONSOLIDADO D6` intacto (ya 46267) | S | Test de no-duplicación: el sello no escribe `CONSOLIDADO!D6` |
| R-D-1..R-D-6 | Veredicto por frente (R2-remesh, R4-remesh, SALDOS POR NOTA detalle, ANT EXT-REV, ANTICIPOS, DetRetri E..O por cadena): adoptar / divergencia / fuera-de-alcance, con cita T0b y consecuencia | D4 | Tabla en el plan + `project-context.md`; cero diff de producción en T4 (verificado por diff) |
| R-R-1 | Suite base 406/406 verde antes y después + build 0 warnings; goldens julio ±0.5; DetRetri-D 5/5 vs R10 ambos períodos | Transversal | `dotnet build …slnx` + `dotnet test` en cada tarea |
| R-R-2 | `project-context.md` (doctrina + excepción + gate + sello + veredictos) + manual (proceso: qué base, qué sella cada corrida, cómo leer un FAIL del gate) | Transversal | Diff acotado; sin código en el mismo commit |

### 3.2 Scenarios (Given/When/Then)

- **S1 (TDD rojo — reproduce el bug):** Given la salida agosto actual (`F50=F32+F47+F12-L12-L32`), When el gate evalúa TOT_OPT ASE1, Then FAIL con `ASE 1 Reporte Componentes R1!F50: evaluado <…> vs dominio 18298670992.96 ±0.5` (y PASS en julio con la misma fórmula canónica).
- **S2 (ASE5-3Mes):** Given bloque ASE5 agosto (3 `Mes` 494/525/548), When el pase final recompone, Then `<f>` = `F548+F525+F494-L494-L525` (orden canónico T1, texto = manual) y su evaluación = totOpt dominio (12.105.458.586,04-style Plan 25) ±0.5.
- **S3 (comparador que delata):** Given `BrechaDeclarada` sin blanket R1, When regresión agosto pre-T2, Then divergencias F50/G50/EXTEMP fuera de brechas (rojo intencional); post-T2, Then cero divergencias R1 fuera de versionadas.
- **S4 (julio identidad):** Given período 2026072 + base julio, When flujo 5-ASE Q2 + pase final, Then diff R1 vs golden/manual = 0 (fórmulas y literales) y suite goldens ±0.5 verde.
- **S5 (EXTEMP 0-Aplic):** Given ASE3-agosto (0 filas `Aplicacion`, dominio extemp = 0), When recomposición, Then visible EXTEMP = literal `0` (auditado en log + test) y gate PASS ±0.5.
- **S6 (sello):** Given base 2026082 + fuente con fecha legible (origen R-S-1), When flujo Q2, Then `DetRetri2026082!D6/D7` y `DetValiRetri2026082!D6/D7` = sello (fecha 02/09/2026-style + hora 07:42-style según origen); Given fuente ilegible, When flujo, Then `ERR-FORMATO-FUENTE` con período + archivo + celda, sin escritura parcial.
- **S7 (decisión sin código):** Given T0b §2 + estado actual, When T4, Then tabla R-D-1..R-D-6 con veredicto y consecuencia por frente; el diff de producción de T4 es vacío.

---

## 4. TASKS

| ID | Tarea | Grupo | Depende de | Verificación |
|---|---|---|---|---|
| T1 | Gate-que-falla-primero (TDD) + fin del blanket R1. (a) Congelar E2 pendiente: G y EXTEMP por ASE 2-5 + G-ASE5-3Mes + sub-visibles de empresa desde el manual (solo lectura zip+XML; si difiere de lo supuesto, el test lo fija). (b) Nuevo gate workbook-vs-dominio (lectura + evaluación BCL, prohibido `<v>`, fail-fast ref-no-resuelta) con tests S1 (rojo agosto-actual, verde julio). (c) Quitar R1-agosto de `BrechaDeclarada` (`:172-185` → solo versionadas H1/T0g, token-período, V3): la regresión agosto queda en rojo INTENCIONAL hasta T2. | G | — | R-G-1 (S1), R-G-2 (S3 pre-T2 en rojo); build 0 warnings; resto de la suite 406/406 (solo los 2 соч nuevos en rojo + regresión agosto) |
| T2 | Recomposición por firma + sucesor del gate salteado. Pase final post-5→1 en `OpenXmlEspejoR1Mutador` (contrato §2.3: TOT_OPT F, total TDF G, EXTEMP F; 0-Aplic → literal 0; log ASE+celda+texto) + validación de forma recompuesta donde hoy se salta (`:283-286, :330-339` rama `espejoDesplazado`) + derogación documentada de `R1Q2ProtectedPorAse` para visibles recompuestos. Tests: texto-vs-manual + evaluación-vs-dominio por ASE (agosto) + identidad julio (S4) + S5. | B | T1 | R-B-1..R-B-4 (S2/S4/S5); S3 post-T2 verde (cierra el rojo T1); goldens ±0.5; DetRetri-D 5/5 ambos períodos |
| T3 | Sello D6/D7. (a) Mini-T0: literal-vs-`<f>` y origen de D6/D7 en base+manual+R10 (ambos períodos) → tabla R-S-1. (b) Propagación (`ProcesadorPeriodo`/`ResultadoRemuneracion`, fail-fast `ERR-FORMATO-FUENTE`) + extensión de `EscribirFechasPeriodo` a las 4 celdas (nombres por `NombresHojaPeriodo`, OADate, estilo preservado, guard anti-fórmula). Tests S6 + no-duplicación `CONSOLIDADO!D6`. | S | T1 (puede ir en paralelo a T2) | R-S-1..R-S-3 (S6); suite + goldens intactos |
| T4 | Decisión R2/R4+soporte (SOLO docs). Re-lectura T0b §2 + estado actual por frente → veredictos R-D-1..R-D-6 (adoptar-geometría / declarar-divergencia / fuera-de-alcance) con cita y consecuencia (comparador: re-encender o versionar; código: follow-up con T0 propio). Cero diff de producción (verificado por `git diff --stat` vacío en código). | D4 | T1 (lecturas), T2/T3 (no bloquean el veredicto) | R-D-1..R-D-6 (S7); sin código; sin datos inventados |
| T5 | Docs: `project-context.md` (doctrina recomposición-por-firma D-A/D-B + excepción `<f>` D-D + gate C + sello extendido D-F + veredictos D4) + manual (qué base usar, qué sella cada corrida, cómo leer un FAIL del gate S1). | Transversal | T2, T3, T4 | R-R-2; diff acotado; sin código en el mismo commit |
| T6 | Cierre: build 0 warnings + suite completa (406 base + nuevos, todo verde) + goldens julio ±0.5 + DetRetri-D 5/5 vs R10 ambos períodos + grep de cierre (ninguna `<f>` fuera del contrato; cero literales nuevos de período) + auditoría S4/S6 + CRLF + sin commits. | Transversal | T1..T5 | R-R-1; DoD §5 |

**Orden sugerido (TDD):** T1 (rojo que reproduce el bug) → T2 (recomposición que lo cierra) → T3 (sello, en paralelo a T2 si se quiere) → T4 (decisión) → T5 → T6. Build 0 warnings antes de marcar cada tarea completa. Sin commits (los hace el Ingeniero).

---

## 5. Resumen ejecutivo, riesgos y decisiones

- **Resumen:** el T0 cerró la causa (E1: dimensionado por conteo + borrado en pie + fuente que recorta en cabeza; E5: ASE5 exige recomponer, no reanclar) y el mapa de cegueras (E4). El plan convierte esa evidencia en 4 frentes: gate que hoy fallaría (TDD rojo-primero, T1), recomposición por firma en pase final del mutador (T2, con Opción A como red y julio-identidad como invariante), sello D6/D7 con origen congelado por mini-T0 (T3) y veredictos documentados sin código para R2/R4+soporte (T4). La única reescritura de `<f>` del proyecto queda explícita, acotada al contrato y auditada término a término (D-D).
- **Riesgo principal:** R-FORMA (la forma recompuesta diverge del manual en un borde y la excepción `<f>` crea divergencia silenciosa). Contenido por triple red: E2 congelado antes de componer, gate C + comparador f-vs-f en cada tarea, y julio-identidad como invariante de reversión.
- **Decisión para el Ingeniero:** ratificar D-A..D-G (§0.3). **Forks con recomendación:** D-A (pase final en mutador vs writer — recomendado mutador: posee la geometría; el writer con `omitirR1` no la conoce) y D-F (origen del sello a congelar en mini-T0 — hipótesis: mismo sello de proceso que `CONSOLIDADO D6`; si el mini-T0 la refuta, el sello se re-enfoca sin inventar). **Sin preguntas bloqueantes pendientes.**

### DoD (Definition of Done)

1. `dotnet build Remuneracion.WinForms/Remuneracion.WinForms.slnx` con 0 warnings.
2. Suite completa verde sin regresión (base 406/406 + nuevos) + goldens Capa A Q1+Q2 ±0.5 + DetRetri 5/5 vs R10 (julio y agosto) + 0-diff R1 julio.
3. Agosto 2026082 end-to-end con visibles R1 recompuestos = manual en texto (5 ASE, incl. ASE5-3Mes) y evaluación = dominio ±0.5 (S1/S2/S3/S5).
4. `DetRetri*/DetValiRetri* D6/D7` sellados desde origen congelado (S6); fuente ilegible → fail-fast archivo+celda; `CONSOLIDADO!D6` intacto.
5. Tabla R-D-1..R-D-6 publicada; comparador sin blanket R1 (solo versionadas con cita); `project-context.md` + manual actualizados.
6. Grep de cierre: diff de `<f>` = solo visibles R1 del contrato; sin commits del agente; finales de línea CRLF; sin emojis.

### Follow-ups explícitos (fuera de este plan)

- R4 espejo-motor, R4-por-empresa Q2, CF/DV del reanclaje (W-4), `R-EXTRA-CONCEPTO` (Plan 23), validación de nombres de hoja en preflight: vivos, intactos.
- Todo frente D4 con veredicto "requiere código": su propio T0 antes de cualquier línea (D-G).
- Números del encargo "Recaudo Base / Remuneración Final" (NEEDS_CONTEXT T0 §4.1): si el Ingeniero aporta el origen (pantalla UI u otro archivo), se abre T0 puntual; no se rastrea en este plan.
- Retiro/archivo de la evidencia T0 (`31-T0-Evidencia-R1.md` se conserva como fuente de E1..E8).
