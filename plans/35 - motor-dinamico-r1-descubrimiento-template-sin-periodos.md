# Plan 35 — Motor dinámico R1: descubrimiento por firma + forma-del-template, cero períodos

> **Alcance:** diseñar el MOTOR DINÁMICO R1 — sin hardcodeo de períodos, sin mapas congelados por mes — EN VEZ de un fix puntual F463. Este plan **SUBSUME** el fix del clasificador de frontera F463 como parte del motor genérico (no existe un fix puntual separado). Cambio de producción esperado: mínimo y neto-negativo en líneas (se ELIMINAN `CeldasInterioresAgosto` y `EsPeriodoJulio`; se agregan clasificador por frontera de sección + descubrimiento por firma + composición por forma-del-template + tests + gates). Cero cambios de semántica de negocio: cada subtotal interior sigue siendo la misma forma del manual con las filas reales del período.
> **Rector (IRRENUNCIABLE):** `README.md` + `Docs/Propuesta_Proyecto_Automatizacion_Remuneracion_UAESP.md` + `Docs/Detalle de plantilla.docx` + `Docs/Proceso de Recaudo.docx` + `Docs/Prompt_Maestro_Proyecto_Remuneracion_UAESP _ Vo .docx` + `.opencode/project-context.md`. **Invariantes del proyecto:** tolerancia ±0.5 en comparaciones golden; NUNCA usar cachés `<v>` como oráculo (doctrina Planes 28/29); doctrina calcChain (Plan 28) intacta — no se toca `CalculationChainPart`, `calcPr`/`fullCalcOnLoad`, ni `<f>`/`<v>` fuera de lo declarado; preservación shared-formula (Plan 33, `EscribirFormulaPreservando`) intacta — incl. `G53` (`si=0`) y `G468` (`si=20`); build 0 warnings antes de marcar cualquier tarea completa; finales de línea CRLF sin mezclar; NUNCA git commit/push (los hace el Ingeniero con `#commit`/`#push`); la implementación posterior la ejecuta `implementer` usando SOLO este plan aprobado; `code-reviewer` verifica spec compliance + 0 warnings. **Quincena por dominio, nunca por contenido.**
> **Continuidad:** Planes 21 (espejo), 25 (roles por firma), 26/27 (preflight + firma + hoja-por-nombre), 28 (calcChain + sello), 29 (paridad + comparador f-vs-f), 30 (naming dinámico + base canónica), 31 (visibles por firma), 32 (interior por firma + catálogo C ABIERTO D-B + quirk `--L` D-G), 33 (shared-formula + T0 agosto H1) + anexo `33-T0-Evidencia-Agosto2.md`, 34 (sinonimia `RECIPROCIDAD`↔`NUEVO ESQUEMA` como DATOS + D-H cero-quema-de-períodos + catálogo abierto) vigentes. Este plan NO reabre lectura, cálculo, espejo-dimensional, preflight, sellos, naming, calcChain, visibles, shared-preservation ni sinonimia fuera de lo declarado: **julio = identidad por construcción; agosto ASE4-F463 + ASE2-silencioso = rojos que se cierran con el motor.**
> **Insumos canónicos (no se re-fijan, no se inventan datos):** base canónica julio + base canónica agosto `Docs/Prueba Agosto-2/Plantilla_Remuneracion_2026082.xlsx` (rutas vigentes post-reorg `a867706`); fuentes R1 por ASE (`Recaudoporcomponente_*`); oráculos manuales del administrativo (julio + `Remuneracion 202608-2 Total_7721.xlsx` — `Reporte Componentes R1`); `InterioresR1Esperados` + `TotalesR1Esperados` + `SharedFormulaR1Esperados` como fixtures que mandan. Reproducción determinista vía `Remuneracion.Cli` hacia rutas temp FRESCAS (nunca sobre salidas existentes).
> **Numeración:** `plans/` 01..34 ocupados (33 + anexo T0, 34); este plan toma el primer correlativo libre, **35**.
> **Estado:** APROBADO por el Ingeniero 2026-10-08 (D-H cero-quema-de-períodos, catálogo-C abierto y Plan 34 vigentes) — listo para `implementer` usando SOLO este documento.
> **Fecha:** 2026-10-08

---

## 0. Clarification Gate

**Sin preguntas bloqueantes: la evidencia viene ADOPTADA por encargo (no se re-diagnostica), la dirección de diseño viene APROBADA (descubrimiento + forma-del-template + faseo con gates), y las restricciones vienen cerradas (D-H cero-quema-de-períodos + catálogo-C abierto + Plan 34 en vigor).** Los puntos con recomendación (rol de `CompositorInteriorR1`, D-C; alcance del gate manual-oráculo, D-G) vienen redactados en la ruta recomendada con su verificación de cierre dentro de la tarea. Punto de bloqueo respetado: NO se implementa nada en este documento.

### 0.1 Diagnóstico adoptado (causa cerrada, no re-diagnosticar)

| # | Afirmación adoptada | Reutilización en este plan | Re-verificación (el test, no la mano) |
|---|---|---|---|
| E1 | **F463 causa raíz H2 (CONFIRMADA):** el predicado de adyacencia `Filas[i+1].EsAplicacionTotal` en `MesOportuno` (§L1064-1083) y `FilasAplic` (§L1090-1109) solo funciona cuando la empresa es la ÚLTIMA de su sección. Si la empresa-dato no cierra la sección, su fila oportuna se clasifica mal (o su fila aplicación se pierde) aunque los datos existan. | El motor reemplaza el predicado de adyacencia por el clasificador de frontera de sección (D-A). `MesOportuno` y `FilasAplic` comparten el clasificador (un solo punto de verdad). | T2: test puro del clasificador con la secuencia ASE4-agosto (fronteras reales) + gates manual-oráculo T3 que reproducen el manual término a término. |
| E2 | **Regla correcta (derivada del oráculo):** una fila-dato de empresa (`EsDatoEmpresa`) es **Aplicacion** ssi la próxima fila debajo de ella que sea `EsMesTotal \|\| EsAplicacionTotal` es `EsAplicacionTotal`; si es `EsMesTotal`, la fila es **Oportuno**. Reproduce el manual administrativo exactamente en ASE4 (`F461`/`F463`/`F466`/`F468`/`F476`/`F478`) y ASE2 (`F199`/`F204`/`F206`). | Especificación normativa del clasificador D-A; el fixture congela las fronteras esperadas por bloque desde el disco. | T2/T3: asserts por celda `F461`, `F463`, `F466`, `F468`, `F476`, `F478` (ASE4) + `F199`, `F204`, `F206` (ASE2) con texto `=` manual. |
| E3 | **H1 (0-Aplic→literal) DESCARTADA para ASE4:** el sub-bloque de F463 SÍ trae filas aplicación reales; componer literal era enmascarar el defecto de clasificación. **Pero** el mecanismo `filasAplicInexistentes→literal` (parámetro de `CompositorInteriorR1.ComponerInterior`) **queda pendiente** para celdas-cero genuinas clase L-1 (ASE2 `F201`, ASE3, ASE5 ENERBIT/CIUDAD, E6 `F209`/`F219` como follow-up). | El motor NO usa el camino literal para F463; el camino literal sobrevive SOLO para L-1 genuino declarado, con su propio diseño futuro. | T3: F463 compone fórmula (no literal); el gate delata como FAIL cualquier literal donde el manual trae `<f>`. |
| E4 | **ASE2 MISCOMPOSICIÓN SILENCIOSA:** `F204` app 4 términos vs manual 3; `F206` app 1 término vs manual 2 — el pipeline "pasa" pero compone MAL. **"Passes" ≠ "correct".** | Los gates de aceptación comparan texto compuesto vs oráculo manual para TODOS los ASE, no solo la celda que aborta (D-G). | T1/T3: gate texto-vs-manual workbook-wide del interior (todas las celdas sustantivas, ambos meses); F204/F206 como guardianes explícitos. |
| E5 | **Quemados a ELIMINAR (pre-existentes, no introducidos por planes 33/34):** (1) mapa congelado `CeldasInterioresAgosto` (`OpenXmlEspejoR1Mutador.cs` ~L854-897); (2) early-return NO-OP `EsPeriodoJulio()` (~L906-930, detección por hojas `DetRetri2026071/2026072`). | T5 los DELETEA tras las pruebas de equivalencia (T1) e identidad-julio (T4). Sin esas pruebas, la eliminación está PROHIBIDA. | T5: `grep` de cierre — 0 ocurrencias de `CeldasInterioresAgosto` y `EsPeriodoJulio` en producción; suite verde sin ellos. |
| E6 | **Invariantes que DEBEN sostener:** doctrina calcChain (Plan 28) intacta; preservación shared-formula (Plan 33, `EscribirFormulaPreservando`) intacta incl. `G53`/`G468`; prohibido `<v>` como oráculo; tolerancia ±0.5; 0 warnings; CRLF; sin commits. | Redes ciegas en T6: test estructural calcChain + `SharedFormulaR1Gate` + goldens + DetRetri-D. | T6: las tres redes verdes + greps de cierre (ningún `<v>` tocado, ningún flag calcChain cambiado). |
| E7 | **Plan 34 `SinonimosEmpresaR1` sigue en vigor:** las equivalencias de empresas viven como FILAS citadas de DATOS, nunca como ramas por nombre (doctrina Plan 32 D-B, reafirmada). | El clasificador y el descubrimiento consumen `SonMismaEmpresa`; ninguna rama `if EMPRESA` nueva. | T6: `grep` — 0 comparaciones por nombre de empresa fuera de las filas citadas de la tabla. |

### 0.2 Mapeo al Rector (Sale / No sale)

**Sale (EXPLÍCITO):**
- Re-diagnosticar H1/H2, el quirk `--L`, la sinonimia `RECIPROCIDAD`↔`NUEVO ESQUEMA`, o reabrir lectura, cálculo, espejo-dimensional, preflight, sellos, naming, calcChain, visibles, shared-preservation o paridad fuera de lo declarado.
- Cualquier fix puntual F463 (celda, fila o ASE hardcodeados; `if F463…`, `if ASE4…`, `if agosto…`) — este plan lo SUBSUME por mandato; un diff puntual se rechaza en review.
- Literales o ramas por período en el diff de producción (cualquier capitalización de períodos, meses o códigos de hoja de período fuera de `NombresHojaPeriodo`) — restricción D-H; los períodos existen SOLO como datos de prueba.
- Ramas por nombre de empresa en producción (`== "ENEL"` / `"OCCIDENTE"` / `"RECIPROCIDAD"` / …) fuera de las FILAS citadas de `SinonimosEmpresaR1` — doctrina catálogo-abierto D-B/Plan 34.
- Tocar `SaneadorCadenaCalculo`, `CalculationProperties`/`fullCalcOnLoad`, `<v>`, `EscribirFormulaPreservando`/`Reanclar` (salvo consumo intacto), sellos, naming, preflight, readers o cálculo.
- Tocar cachés `<v>` como oráculo en cualquier gate o comparador; normalizar el quirk `--L`; renombrar rótulos del template.
- Paquetes NuGet; commits/push.

**No sale (entra, alcance congelado):**
| Frente | Superficie de cambio |
|---|---|
| B1 — clasificador de frontera | Predicado puro en Core (p. ej. `R1FirmaInterior.EsAplicacionDeEmpresa` / `ClasificarFrontera`) con la regla E2 + consumo en `MesOportuno`/`FilasAplic` (único punto de verdad) |
| B2 — descubrimiento por firma | Localizador de subtotales interiores en el workbook por firma (`EsSubtotalEmpresa` / `EsExtemporaneoInterior` / `EsSubsidioInterior` / `EsAFaseoInterior`, todas en `R1FirmaInterior.cs`) + delimitación de cada sub-bloque (subtotal previo..este); reemplaza `CeldasInterioresAgosto` |
| B3 — composición por forma-del-template | Parser del CÓMO desde la fórmula ORIGINAL del template de cada celda (orden de términos, signos, espejo de columnas G/H/L, quirk `--L`) + composición del QUÉ FILAS desde fuente vía B1 + sinonimia Plan 34; `CompositorInteriorR1` queda como validador/fallback según D-C |
| G — gates TDD | Equivalencia descubrimiento==mapa (T1) + clasificador puro (T2) + texto-vs-manual workbook-wide todos los ASE (T3) + identidad-julio byte-idéntica texto+atributos shared (T4) + regresión global (T6); rutas explícitas, prohibido `Insumos.Raiz()` y `<v>` |
| R — regresión | 5 ASE × ambos meses, goldens ±0.5, DetRetri-D 5/5 vs R10, `SharedFormulaR1Gate` + test calcChain como redes ciegas |
| D — docs | `project-context.md` (doctrina motor-dinámico + clasificador + descubrimiento + forma-del-template + supervisión de mes nuevo) + manual (procedimiento de mes nuevo supervisado) |

**Fuera de scope explícito (ni diseño ni código en este plan):** motor-espejo R4; cobertura `conditionalFormatting`/`dataValidations` (W-4); semántica DetRetri (redondeo/columna D intactos); plumbing UI/CLI (la CLI solo como runner de repro); reconciliación de fixture-root de los 511 rojos pre-existentes (dependencia declarada: lo nuevo resuelve fixtures por ruta explícita, nunca `Docs/Insumos` raíz); completar mapa ASE2 `F209` ENERBIT / `F219` CIUDAD LIMPIA-ACUEDUCTO (follow-up con su propio diseño, clase L-1 del TODO(T2-full) Plan 32).

### 0.3 Decisiones ejecutivas (aprobación = aceptar estas)

| ID | Decisión |
|---|---|
| D-A | **Clasificador de frontera de sección (H2→regla E2).** Una fila-dato de empresa es Aplicacion ssi la próxima `EsMesTotal \|\| EsAplicacionTotal` debajo es `EsAplicacionTotal`. Predicado puro en Core, consumido por `MesOportuno` Y `FilasAplic` (un solo punto de verdad — el defecto vivía duplicado en ambos). Descarta H1 para F463 (las filas aplicación existen; el literal era enmascaramiento) sin retirar el camino literal para L-1 genuino. |
| D-B | **Descubrimiento desde el template, no mapa congelado.** Los subtotales interiores se localizan en el workbook ya dimensionado por firma (`R1FirmaInterior`: `EsSubtotalEmpresa` C=`<empresa>`/D=OPORTUNO/E=Total; `EsExtemporaneoInterior` D=EXTEMPORANEO/E=Total; `EsSubsidioInterior`; `EsAFaseoInterior`) y cada sub-bloque se delimita (subtotal previo..este). **Contrato:** en los templates vigentes el descubrimiento debe hallar EXACTAMENTE las celdas del mapa actual — probado primero con test de equivalencia de cero-cambio-de-comportamiento (T1). Recién entonces se conmuta. |
| D-C | **Forma-del-template manda en el CÓMO; `CompositorInteriorR1` queda como validador + fallback acotado (recomendado).** El orden de términos, los signos, el espejo de columnas G/H/L y el quirk `--L` se parsean de la fórmula ORIGINAL de ESA celda en el template (no de formas hardcodeadas por etiqueta). Las QUÉ FILAS vienen de la fuente vía D-A + sinonimia Plan 34. Cardinalidad igual → remap puro (mismas posiciones relativas, filas nuevas); cardinalidad distinta → composición con la forma parseada. `CompositorInteriorR1` NO se borra: (a) **validador** — su forma canónica debe coincidir en clase con la forma parseada (mismatch = fail-fast con ambos textos, gana el template); (b) **fallback** — single-refs `SUBS`/`AFASEO` y camino literal L-1 (`filasAplicInexistentes`) donde la forma-del-template es trivial. Descartado: borrar el compositor (se pierde la red canónica y el camino L-1) y descartado: mantener formas hardcodeadas como primario (reintroduce el hardcodeo por etiqueta que este plan elimina). |
| D-D | **Puerta de identidad-julio + borrado de rama (condicional duro).** Se demuestra que la recomposición sobre geometría julio reproduce el template byte-idénticamente (texto `<f>` + atributos shared `t`/`ref`/`si`) y RECÉN se DELETEA el early-return `EsPeriodoJulio`. Sin esa prueba verde, el borrado está PROHIBIDO; la prueba es la aceptación de la tarea de borrado. |
| D-E | **Sinonimia intacta (Plan 34 en vigor).** El clasificador y el descubrimiento comparan empresas SOLO vía `SinonimosEmpresaR1.SonMismaEmpresa`. Añadir una equivalencia futura = añadir una FILA citada (sin cita no hay fila). Cero ramas por empresa. |
| D-F | **D-H cero-quema-de-períodos (irrenunciable).** El diff de producción contiene CERO literales/ramas por período; el motor no recibe ningún período (descubrimiento por firma + forma-del-template son agnósticos al mes). Criterio: `grep` de literales de período en el diff = 0. |
| D-G | **"Passes" ≠ "correct": gate manual-oráculo para TODOS los ASE.** La aceptación compara texto compuesto vs manual administrativo celda a celda en todo el interior sustantivo (ambos meses), no solo la celda que aborta. F204 (4-vs-3) y F206 (1-vs-2) son guardianes explícitos: si el gate no los cubriera, el defecto silencioso volvería. |
| D-H | **Doctrina de mes nuevo supervisado.** La primera corrida de un mes inédito es SUPERVISADA: se compara contra el archivo del administrativo de ese mes y se congela como golden. Sin golden congelado no hay corrida desatendida del mes. El procedimiento queda documentado (qué comparar, dónde congelar, quién aprueba). |

---

## 1. PROPOSE

### 1.1 Intent

Que el interior de `Reporte Componentes R1` se recomponga para CUALQUIER mes sin mapas congelados ni ramas por período: el workbook dice QUÉ subtotales hay (descubrimiento por firma), el template dice CÓMO es cada fórmula (forma parseada de la celda), la fuente dice QUÉ FILAS la alimentan (clasificador de frontera + sinonimia) — con F463 y el silencioso ASE2 cerrados por construcción, julio byte-idéntico, y el mapa `CeldasInterioresAgosto` + la rama `EsPeriodoJulio` eliminados del código.

### 1.2 In Scope

- Clasificador de frontera E2 (puro, Core) + consumo en `MesOportuno`/`FilasAplic`.
- Descubrimiento de subtotales interiores por firma + delimitación de sub-bloques (reemplazo del mapa).
- Composición por forma-del-template (parse del CÓMO) + remap/composición de filas + `CompositorInteriorR1` como validador/fallback (D-C).
- Gates TDD en faseo (equivalencia → switch → identidad-julio → borrado → regresión) con oráculo manual para todos los ASE.
- Eliminación de `CeldasInterioresAgosto` y `EsPeriodoJulio` tras sus pruebas (T1/T4).
- Docs: doctrina motor-dinámico + procedimiento de mes nuevo supervisado.

### 1.3 Non-Goals

Ver §0.2 "Sale" + fuera de scope explícito. En particular: NO se prescribe fix puntual F463 alternativo (subsumido); NO se completa `F209`/`F219` (follow-up L-1); NO se generaliza a R4/CF-DV/DetRetri/UI-CLI; NO se reconcilian los 511 rojos de fixture-root.

### 1.4 Riesgos

| ID | Riesgo | Mitigación |
|---|---|---|
| R-CLASIF | El clasificador E2 falla en una frontera no vista (sección con orden atípico) y mueve dinero entre Oportuno/Aplicación | Regla derivada del oráculo + fixture de fronteras por bloque desde el disco + gate texto-vs-manual todos los ASE (D-G); fail-fast con bloque+filas si una frontera no resuelve (nunca 0 silencioso) |
| R-DESCUBRIMIENTO | El descubrimiento halla de más (falsos subtotales) o de menos (subtotal real no firmado) y el pase compone celdas fuera de contrato o deja huecos | Contrato exacto T1 (descubrimiento == mapa vigente en templates actuales) + fail-fast por celda no resuelta + gate workbook-wide; sobre-incluir delata por texto-vs-manual, sub-incluir por celda sin recomponer |
| R-FORMA | El parse de la forma-del-template malinterpreta una fórmula atípica (orden, signos, `--L`, espejo G/H/L) y compone texto divergente | D-C: `CompositorInteriorR1` como validador de clase (mismatch = fail-fast, gana el template) + fixture texto-vs-manual por celda + guardianes F204/F206/F463 explícitos |
| R-JULIO | Al borrar `EsPeriodoJulio` el pase altera julio | D-D: puerta dura — borrado PROHIBIDO sin identidad byte-idéntica (texto + shared attrs) verde; ante diff se revierte el pase, no el golden |
| R-ATRIBUCION | Descubrimiento+sinonimia atribuyen filas a la empresa equivocada | D-E (solo `SonMismaEmpresa`, filas citadas) + assert texto `=` manual término a término + goldens ±0.5; empresas sin divergencia = identidad |
| R-ALCANCE | El motor se expande a R4/CF-DV/DetRetri/UI y el diff deja de ser mínimo | §0.2 OUT explícito + grep de cierre (diff = motor R1-interior + tests + gates + docs); cada frente OUT reabierto exige su propio T0 |
| R-FIXTURE-ROOT | Gates nuevos heredan los 511 rojos (resolución raíz rota post-reorg) | Rutas explícitas estilo `SharedFormulaR1GateTests` (`Docs/Prueba Julio-2/…`, `Docs/Prueba Agosto-2/…`); prohibido `Insumos.Raiz()` en lo nuevo |

---

## 2. DESIGN

### 2.1 Enfoque: firma descubre, template forma, fuente alimenta — en faseo con puertas que prohíben avanzar en rojo

**Fase 0 — equivalencia (T1, cero cambio de comportamiento).** Nuevo descubridor puro (p. ej. `DescubridorInteriorR1.Descubrir(sheetData, workbookPart)` → lista `(celda, etiqueta)` por ASE usando `R1FirmaInterior`) corriendo EN PARALELO al mapa vigente SIN consumirlo: el test compara conjunto-descubierto vs `CeldasInterioresAgosto` en ambas bases canónicas (julio + agosto) y exige IGUALDAD EXACTA (celda + etiqueta + ASE). Cero producción conmutada; si difiere, el descubridor se ajusta (o el contrato se enmienda con cita al disco), nunca el mapa.

**Fase 1 — clasificador (T2).** Predicado puro de frontera en Core (regla E2) + sustitución del predicado de adyacencia en `MesOportuno` y `FilasAplic` (mismo punto de verdad, ambas ramas global/empresa). Tests puros con secuencias reales (ASE4-agosto con empresa NO-última + ASE2-agosto + julio-identidad del clasificador). F463 compone fórmula real (no literal); H1 queda descartada para ASE4 sin retirar el parámetro literal L-1.

**Fase 2 — conmutación a descubrimiento + forma-del-template (T3).** `RecomponerInterior` itera lo DESCUBIERTO (no el mapa); cada celda compone con forma parseada de su fórmula ORIGINAL del template (orden/signos/espejo/`--L`) y filas de la fuente vía D-A + sinonimia. Cardinalidad igual → remap puro; distinta → composición. `CompositorInteriorR1` valida la clase (D-C) y cubre single-refs/literales. Gates: texto-vs-manual workbook-wide TODOS los ASE ambos meses (F463 + F204 + F206 guardianes) + goldens + DetRetri-D + shared.

**Fase 3 — identidad-julio + borrado (T4→T5).** T4 demuestra recomposición sobre geometría julio = template byte-idéntico (texto `<f>` + atributos `t`/`ref`/`si`, celda a celda del contrato descubierto en julio). T5 DELETEA `EsPeriodoJulio` + `CeldasInterioresAgosto` (difunto tras T3) con grep de cierre. T5 sin T4 verde = PROHIBIDO.

**Fase 4 — regresión global + docs + supervisión (T6→T7).** 5 ASE × ambos meses end-to-end a temp FRESCA: 0-diff julio, agosto `=` manual en interior, goldens ±0.5, DetRetri-D 5/5 vs R10, `SharedFormulaR1Gate` + calcChain estructurales verdes, greps D-H/catálogo-abierto/calcChain/shared. T7 documenta doctrina + procedimiento de mes nuevo supervisado (comparar vs admin del mes → congelar golden → aprobar).

### 2.2 Dónde vive cada pieza y trade-offs

| Eje | Elegida | Descartada |
|---|---|---|
| Clasificador | Predicado puro en Core (regla E2: próxima `EsMesTotal\|\|EsAplicacionTotal` debajo), consumido por `MesOportuno` + `FilasAplic` (D-A) | Parche solo en `FilasAplic` (dejaría `MesOportuno` con el mismo defecto latente — la trampa E7/Plan 33 repetida) / fix puntual F463 (viola el mandato) |
| Descubrimiento | Por firma en el workbook dimensionado (`R1FirmaInterior`, catálogo abierto, sin períodos) con contrato de igualdad exacta vs mapa (D-B) | Mapa por mes/período (el quemado que se elimina) / detección por contenido de fuente (la fuente no trae subtotales: el template manda en QUÉ hay) |
| CÓMO de la fórmula | Parse de la fórmula ORIGINAL del template de ESA celda (D-C) | Formas hardcodeadas por etiqueta (reintroduce hardcodeo; no escala a meses nuevos) / borrar `CompositorInteriorR1` (se pierde red canónica + camino L-1) |
| Validador/fallback | `CompositorInteriorR1` valida clase + fallback single-ref/literal (D-C) | Doble compositor compitiendo (divergencias silenciosas) / sin validador (forma mal parseada compone en silencio — R-FORMA) |
| Puerta julio | Identidad byte-idéntica (texto + shared attrs) como precondición del borrado (D-D) | Borrar la rama "porque ya funciona en agosto" (sin prueba = regresión julio latente) / conservar la rama "por seguridad" (perpetúa la quema de períodos) |
| Mes nuevo | Supervisado con golden congelado del admin del mes (D-H) | Auto-aceptación del primer mes (un cambio de layout compone en silencio) / congelar otro mapa por mes (recrear el quemado) |

### 2.3 Contratos afectados (resumen)

| Contrato | Cambio |
|---|---|
| `R1FirmaInterior` (Core puro) | NUEVO predicado de frontera (E2) + NUEVO descubridor de subtotales interiores por firma + delimitador de sub-bloque; firmas existentes intactas; consumo de `SonMismaEmpresa` donde se compare empresa (D-E) |
| `OpenXmlEspejoR1Mutador.MesOportuno` / `FilasAplic` (§L1064-1109) | Predicado de adyacencia → clasificador de frontera D-A (único cambio de lógica de clasificación) |
| `OpenXmlEspejoR1Mutador.RecomponerInterior` (§L923-983) | Itera lo DESCUBIERTO (no el mapa); compone con forma-del-template + filas D-A + sinonimia; `CompositorInteriorR1` como validador/fallback (D-C); throw enriquecido D-G intacto |
| `OpenXmlEspejoR1Mutador.CeldasInterioresAgosto` (~L854-897) | ELIMINADO en T5 (difunto tras T3; contrato T1 prueba que nada lo extraña) |
| `OpenXmlEspejoR1Mutador.EsPeriodoJulio` (~L906-930) | ELIMINADO en T5 (difunto tras T4; puerta D-D) |
| `CompositorInteriorR1` / `ValidadorTotalesR1Workbook` / `SaneadorCadenaCalculo` / `EscribirFormulaPreservando` / `Reanclar` / sellos / naming / preflight / readers / cálculo | Intactos salvo consumo (D-C: validador + fallback single-ref/literal; resto sin cambios — grep de cierre) |
| Gates nuevos (+ fixtures que mandan) | Equivalencia descubrimiento==mapa (T1) + clasificador puro (T2) + texto-vs-manual workbook-wide todos los ASE (T3) + identidad-julio texto+attrs (T4); en `Remuneracion.IntegrationTests`, BCL/lectura in-process, rutas explícitas, prohibido `<v>` e `Insumos.Raiz()` |
| `ComparadorSalidaVsManualTests.BrechaDeclarada` | Sin cambios de código (el interior cierra por construcción; si una celda exigiera divergencia permanente se versiona con cita, no blanket) |
| `project-context.md` + manual | Doctrina motor-dinámico (D-A..D-H) + procedimiento mes-nuevo-supervisado |

### 2.4 Architecture Validation Certificate

| Principio | Veredicto |
|---|---|
| **SRP** | ✅ Una razón de cambio por costura: Core clasifica (frontera) y descubre (firma), el template da la forma, el mutador orquesta, cada gate contrasta una puerta del faseo, el fixture fija el disco. Ninguna rama por período ni por empresa. |
| **OCP** | ✅ Un mes nuevo o una empresa nueva componen sin código nuevo (firma abierta + forma parseada + filas por frontera); lo no tabulado es identidad; julio-identidad garantiza que lo cerrado no se mueve. |
| **DIP** | ✅ Clasificador/descubridor/parseador puros (sin I/O ni OpenXML en Core); el mutador depende de firmas de Core; tests con insumos reales + BCL ya presente; cero dependencias nuevas (OpenXML 3.5.1 ya referenciado). |
| **Best practices** | ✅ TDD con puertas que prohíben avanzar en rojo (equivalencia → switch → identidad → borrado → regresión); fail-fast con ASE+celda+frontera nombrados; quincena por dominio; `<f>` de negocio y `<v>` intactos; tolerancia ±0.5; sin Excel/COM; Visual Design Intent N/A (cero UI). |
| **Performance** | ✅ Descubrimiento O(filas del bloque) por ASE + clasificación O(filas) por sub-bloque + parse acotado a celdas descubiertas; julio delta-0 reproduce texto existente (costo = verificación). |

---

## 3. SPEC

### 3.1 Requirements

| ID | Requisito | Grupo | Criterio de aceptación |
|---|---|---|---|
| R-B-1 | Clasificador de frontera E2 puro en Core, consumido por `MesOportuno` Y `FilasAplic` | B | Unitarios puros con secuencias reales: empresa NO-última clasifica Aplicacion/Oportuno por próxima `EsMesTotal\|\|EsAplicacionTotal`; empresa última conserva comportamiento; julio-identidad del clasificador |
| R-B-2 | Descubrimiento por firma halla EXACTAMENTE el mapa vigente en templates actuales | B | Test de equivalencia T1: conjunto descubierto == `CeldasInterioresAgosto` (celda+etiqueta+ASE) en base julio y base agosto, cero conmutación aún |
| R-B-3 | Composición por forma-del-template: CÓMO parseado de la fórmula original de cada celda; QUÉ FILAS de la fuente (D-A + sinonimia) | B | Remap puro con cardinalidad igual (texto re-anclado, misma forma); composición con cardinalidad distinta; quirk `--L` y espejos G/H/L preservados por celda |
| R-B-4 | `CompositorInteriorR1` como validador de clase + fallback single-ref/literal (D-C) | B | Mismatch de clase = fail-fast con ambos textos (gana el template); `SUBS`/`AFASEO` y L-1 genuino por el fallback; grep: ninguna forma hardcodeada como primario fuera del validador |
| R-B-5 | F463 + ASE2-silencioso cerrados por construcción (subsumidos, no puntuales) | B | `F461`/`F463`/`F466`/`F468`/`F476`/`F478` (ASE4) y `F199`/`F204`/`F206` (ASE2) con texto `=` manual término a término; F204 3 términos (no 4), F206 2 términos (no 1); F463 fórmula (no literal) |
| R-B-6 | D-H + catálogo abierto en el diff | B | `grep`: 0 literales/ramas por período y 0 comparaciones por nombre de empresa fuera de filas citadas; el motor no recibe ningún período |
| R-G-1 | Gate equivalencia (T1) rojo→verde sin conmutar producción | G | FAIL si descubrimiento ≠ mapa (con el defecto inyectado de prueba); PASS exacto en ambas bases; resto de la suite sin nuevos rojos |
| R-G-2 | Gate texto-vs-manual workbook-wide TODOS los ASE ambos meses (D-G) | G | FAIL pre-T3 con F463 + F204 + F206 divergentes citados; PASS post-T3 celda a celda del interior sustantivo; prohibido `<v>`; rutas explícitas |
| R-G-3 | Puerta identidad-julio byte-idéntica (texto + shared attrs) | G | Recomposición sobre geometría julio = template en `<f>` y `t`/`ref`/`si` por celda descubierta; sin esto T5 PROHIBIDO |
| R-R-1 | Julio 0-diff + goldens ±0.5 + DetRetri-D 5/5 ambos períodos + shared/calcChain intactos | R | 0-diff R1 julio (f-vs-f + literales); goldens verdes; DetRetri 5/5 vs R10; `SharedFormulaR1Gate` verde (`G53`/`G468`); test calcChain verde |
| R-R-2 | Borrado de quemados sin regresión | R | `grep`: 0 `CeldasInterioresAgosto` / `EsPeriodoJulio` en producción; suite verde sin ellos; diff neto-negativo documentado |
| R-R-3 | Suite + build + docs + supervisión | R | `dotnet build …slnx` 0 warnings + `dotnet test` por tarea; `project-context.md` + manual con doctrina y procedimiento de mes nuevo; regens a temp FRESCA; sin commits |

### 3.2 Scenarios (Given/When/Then)

- **S1 (equivalencia — cero cambio):** Given bases julio + agosto, When el descubridor lista subtotales, Then conjunto == `CeldasInterioresAgosto` exacto (celda+etiqueta+ASE) en ambas; producción aún consume el mapa.
- **S2 (clasificador — empresa no-última):** Given secuencia ASE4-agosto (empresa-dato seguida de sección que NO cierra en Aplicacion), When clasificador E2, Then la fila es Oportuno/Aplicacion según la próxima `EsMesTotal||EsAplicacionTotal` (no según el vecino inmediato); `F463` compone fórmula con las filas aplicación reales.
- **S3 (silencioso ASE2):** Given bloque ASE2-agosto, When recomposición T3, Then `F204` = 3 términos del manual (no 4) y `F206` = 2 términos (no 1), texto `=` manual.
- **S4 (remap puro):** Given sub-bloque con cardinalidad igual al template, When recomposición, Then misma forma con filas del período (solo re-anclaje), texto verificado contra el manual.
- **S5 (identidad julio):** Given geometría julio, When pase completo del motor, Then cada celda descubierta = template en texto `<f>` + atributos `t`/`ref`/`si`; cualquier diff = FAIL que bloquea T5.
- **S6 (borrado):** Given T1+T4 verdes, When se eliminan mapa + rama, Then `grep` 0 ocurrencias + suite verde + diff neto-negativo; sin T4 verde el borrado no existe.
- **S7 (no-atribución cruzada):** Given bloque multi-empresa, When matching con empresa X, Then solo filas X (vía `SonMismaEmpresa`); las demás empresas del bloque componen como antes salvo lo que el oráculo autorice.
- **S8 (mes nuevo):** Given mes inédito sin golden, When primera corrida, Then procedimiento supervisado (comparar vs admin del mes → congelar golden → aprobar) antes de cualquier corrida desatendida.

---

## 4. TASKS

| ID | Tarea | Grupo | Depende de | Verificación |
|---|---|---|---|---|
| T1 | Equivalencia descubrimiento==mapa (cero conmutación). (a) Descubridor puro por firma (`EsSubtotalEmpresa`/`EsExtemporaneoInterior`/`EsSubsidioInterior`/`EsAFaseoInterior` + delimitador de sub-bloque) corriendo en paralelo SIN consumirlo. (b) Test que exige igualdad EXACTA vs `CeldasInterioresAgosto` (celda+etiqueta+ASE) en base julio y base agosto + fixture de firmas esperadas que manda. Rutas explícitas, prohibido `<v>` e `Insumos.Raiz()`. | G | — | R-B-2/R-G-1 (S1); build 0 warnings; resto de la suite sin nuevos rojos |
| T2 | Clasificador de frontera E2 (puro) + consumo en ambos caminos. (a) Predicado en Core con regla E2. (b) `MesOportuno` + `FilasAplic` lo consumen (adyacencia eliminada de ambos). (c) Unitarios puros con secuencias reales ASE4/ASE2/julio + `F463` compone fórmula (no literal, H1 descartada para ASE4) + camino literal L-1 intacto para genuinos. | B | T1 | R-B-1/R-B-5 parcial (S2); build 0 warnings |
| T3 | Conmutación a descubrimiento + forma-del-template. (a) `RecomponerInterior` itera lo descubierto; CÓMO parseado de la fórmula original de cada celda (orden/signos/espejo G/H/L/`--L`); QUÉ FILAS vía T2 + sinonimia Plan 34 (intacta). (b) `CompositorInteriorR1` como validador de clase + fallback single-ref/literal (D-C). (c) Gate texto-vs-manual workbook-wide TODOS los ASE ambos meses con guardianes F463/F204/F206 (D-G) + goldens + DetRetri-D + shared como redes. | B+G | T2 | R-B-3/R-B-4/R-B-5/R-G-2 (S2/S3/S4/S7); build 0 warnings |
| T4 | Puerta identidad-julio byte-idéntica. Recomposición sobre geometría julio == template en texto `<f>` + atributos `t`/`ref`/`si` por celda descubierta. Sin este verde, T5 PROHIBIDO (D-D documentado en el gate). | G | T3 | R-G-3 (S5); build 0 warnings |
| T5 | Borrado de quemados (solo con T1+T4 verdes). DELETE `CeldasInterioresAgosto` + `EsPeriodoJulio` + cableado muerto asociado. Grep de cierre + suite verde + diff neto-negativo. | B | T4 | R-R-2/R-B-6 (S6); build 0 warnings |
| T6 | Regresión global end-to-end. Regens FRESCAS a temp (5 ASE × ambos meses): julio 0-diff R1; agosto interior `=` manual; goldens Q1+Q2 ±0.5; DetRetri-D 5/5 vs R10 ambos períodos; `SharedFormulaR1Gate` verde incl. `G53`/`G468`; test calcChain verde; greps D-H + catálogo-abierto + ningún `<v>` tocado + ningún flag calcChain. Fixture-root: solo se listan los 511 rojos pre-existentes (dependencia, no se tocan). | R | T5 | R-R-1 (S5/S6); suite completa verde salvo dependencia declarada |
| T7 | Docs + doctrina supervisada. `project-context.md` (motor-dinámico D-A..D-H + clasificador + descubrimiento + forma-del-template + D-C + lección "passes≠correct") + manual (cómo leer un FAIL de cada gate + procedimiento de mes nuevo supervisado S8). Sin código en el mismo commit. | D | T6 | R-R-3 (S8); diff acotado; CRLF |
| T8 | Cierre global: build 0 warnings + suite (base + nuevos) verde con la única excepción declarada de fixture-root + goldens ±0.5 + DetRetri-D 5/5 + greps globales + CRLF + sin commits. | Transversal | T1..T7 | R-R-3; DoD §5 |

**Orden sugerido (puertas duras, TDD estricto):** T1 (equivalencia, cero cambio) → T2 (clasificador) → T3 (switch a descubrimiento+forma, F463 + silencioso ASE2 en verde manual-oráculo) → T4 (identidad julio) → T5 (borrado, prohibido sin T4) → T6 (regresión global) → T7 → T8. Build 0 warnings antes de marcar cada tarea completa. Sin commits (los hace el Ingeniero con `#commit`/`#push`).

---

## 5. Resumen ejecutivo, riesgos y decisiones

- **Resumen:** por mandato no hay fix puntual F463: el defecto de adyacencia (H2) se corrige dentro de un motor R1-interior dinámico en tres piezas — clasificador de frontera E2 (una fila-dato es Aplicacion ssi la próxima `EsMesTotal||EsAplicacionTotal` debajo es Aplicacion), descubrimiento de subtotales por firma en el workbook (contrato de igualdad exacta vs el mapa vigente antes de conmutar), y composición con el CÓMO parseado del template de cada celda (las QUÉ FILAS vienen de la fuente vía clasificador + sinonimia Plan 34, intacta). El faseo con puertas (equivalencia → switch con oráculo-manual para todos los ASE → identidad-julio byte-idéntica → borrado → regresión → docs) contiene el cambio de comportamiento en todos los ASE: F463 compone fórmula real, F204/F206 dejan de miscomponerse en silencio, julio queda byte-idéntico, y el mapa congelado + la rama por período se eliminan solo tras sus pruebas.
- **Riesgo principal:** R-CLASIF + R-DESCUBRIMIENTO (cambio de comportamiento en todos los ASE a la vez). Contenido por cuádruple red faseada: equivalencia exacta pre-conmutación (T1), oráculo-manual celda a celda para todos los ASE con guardianes explícitos (T3/D-G), identidad-julio como puerta del borrado (T4/D-D), y regresión global con goldens + DetRetri-D + shared + calcChain ciegos (T6).
- **Decisión para el Ingeniero:** ratificar D-A..D-H (§0.3). **Forks con recomendación:** D-C (`CompositorInteriorR1` como validador+fallback con forma-del-template primaria — recomendado: conserva la red canónica y el camino L-1 sin reintroducir hardcodeo) y alcance (motor R1-interior + `F209`/`F219`, R4, CF/DV, DetRetri, UI/CLI y fixture-root como OUT/dependencias declaradas — recomendado: dif mínimo, cada frente OUT con su propio T0). **Sin preguntas bloqueantes pendientes.**

### DoD (Definition of Done)

1. `dotnet build Remuneracion.WinForms/Remuneracion.WinForms.slnx` con 0 warnings.
2. Puertas en orden: T1 equivalencia exacta (descubrimiento == mapa, sin conmutar) → T3 texto `=` manual en todo el interior sustantivo ambos meses (F463 + F204 + F206 guardianes) → T4 identidad-julio byte-idéntica (texto + `t`/`ref`/`si`) → T5 borrado con grep 0 + suite verde.
3. Julio 0-diff R1 (f-vs-f + literales) + goldens Q1+Q2 ±0.5 + DetRetri-D 5/5 vs R10 (ambos meses) + `SharedFormulaR1Gate` verde incl. `G53`/`G468` + test calcChain verde.
4. Greps de cierre: 0 literales/ramas por período en el diff (D-H); 0 comparaciones por nombre de empresa fuera de filas citadas (catálogo abierto); `CeldasInterioresAgosto`/`EsPeriodoJulio` 0 ocurrencias; ningún `<v>`, ningún flag calcChain, `EscribirFormulaPreservando`/`Reanclar`/saneador intactos salvo consumo declarado.
5. `project-context.md` + manual actualizados (incl. procedimiento mes-nuevo-supervisado); regens siempre a temp FRESCA; sin commits del agente; CRLF; sin emojis.

### Follow-ups explícitos (fuera de este plan)

- E6/`F209` ENERBIT / `F219` CIUDAD LIMPIA-ACUEDUCTO (clase L-1, TODO(T2-full) Plan 32) + resto L-1 genuino (ASE2 `F201`, ASE3, ASE5 ENERBIT/CIUDAD): su propio diseño (camino `filasAplicInexistentes→literal` ya reservado, no expandido aquí).
- Motor-espejo R4, R4-por-empresa Q2, CF/DV del reanclaje (W-4), `R-EXTRA-CONCEPTO` (Plan 23), validación de nombres de hoja en preflight, remesh R2/R4, re-encendido SALDOS/AJUSTES, `INTERVENTORIA!R26`: vivos, intactos.
- Reconciliación de rutas raíz de fixtures (511 rojos pre-existentes post-reorg `a867706`): dependencia declarada; lo nuevo usa rutas explícitas y nunca `Docs/Insumos` raíz.
- Retiro/archivo de evidencias: las regens de evidencia viven en temp, no se versionan; los goldens congelados (incl. el primer golden de cada mes nuevo supervisado) sí.
