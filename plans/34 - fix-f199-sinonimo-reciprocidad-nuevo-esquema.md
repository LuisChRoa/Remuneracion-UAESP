# Plan 34 — Fix T5 Agosto: sinónimo `RECIPROCIDAD` ↔ `NUEVO ESQUEMA` en el interior R1 (ASE2 `F199 [SUB_EMP]`)

> **Alcance:** SOLO el fix T5 de agosto para el camino que aborta `Espejo R1 (Plan 32/T2-interior): ASE 2 Reporte Componentes R1!F199 [SUB_EMP]: faltan las anclas [Mes0] del sub-bloque` — rama `SUB_EMP` de `DerivarFilasPorFirma` (empresa leída del template `C`) + matching `MesOportuno` / `EsDatoEmpresa` en `OpenXmlEspejoR1Mutador.cs` + `R1FirmaInterior.cs`. Causa adoptada sin re-diagnóstico: veredicto T4 (`plans/33-T0-Evidencia-Agosto2.md`, H1 CONFIRMADA, H2/H3 DESCARTADAS). Cambio de producción esperado: mínimo (tabla de sinonimia como DATOS en Core + matcher puro consumido por el overload existente + tests + gate). Cero cambios de semántica de negocio: el subtotal `F199` sigue siendo `ΣF(datos-emp) − ΣL(datos-emp menos último)`; solo aprende que `RECIPROCIDAD` y `NUEVO ESQUEMA` son la misma empresa.
> **Rector (IRRENUNCIABLE):** `README.md` + `Docs/Propuesta_Proyecto_Automatizacion_Remuneracion_UAESP.md` + `Docs/Detalle de plantilla.docx` + `Docs/Proceso de Recaudo.docx` + `Docs/Prompt_Maestro_Proyecto_Remuneracion_UAESP _ Vo .docx` + `.opencode/project-context.md`. **Invariantes del proyecto:** tolerancia ±0.5 en goldens; NUNCA usar cachés `<v>` como oráculo; doctrina calcChain (Plan 28) intacta; preservación shared-formula (Plan 33, `EscribirFormulaPreservando`) intacta — `G468` es además celda interior agosto (`EXT_INT`) y NO debe regresionar; build 0 warnings antes de marcar cualquier tarea completa; finales de línea CRLF sin mezclar; NUNCA git commit/push (los hace el Ingeniero con `#commit`/`#push`); la implementación posterior la ejecuta `implementer` usando SOLO este plan aprobado; `code-reviewer` verifica spec compliance + 0 warnings. **Quincena por dominio, nunca por contenido.**
> **Continuidad:** Planes 21 (espejo), 25 (roles por firma), 26/27 (preflight + firma + hoja-por-nombre), 28 (calcChain + sello), 29 (paridad), 30 (naming + base canónica), 31 (visibles), 32 (interior + catálogo C ABIERTO, D-B), 33 (shared-formula + T0 agosto con veredicto H1) vigentes. Este plan NO reabre lectura, cálculo, espejo-dimensional, preflight, sellos, naming, calcChain, visibles ni el mapa `CeldasInterioresAgosto` fuera de lo declarado: **julio = identidad por construcción; agosto ASE2-F199 = rojo que se cierra.**
> **Insumos canónicos (no se re-fijan, no se inventan datos):** base canónica agosto `Docs/Prueba Agosto-2/Plantilla_Remuneracion_2026082.xlsx` (rutas vigentes post-reorg `a867706`; las citas a `Docs/Prueba2/...` están stale); fuente R1 ASE2 (LIME) `Docs/Prueba Agosto-2/Insumos/2-Lime/Recaudoporcomponente_to_date16082026ddMMyyyy_to_date31082026ddMMyyyy___202691163346577.xlsx`; oráculo manual `Docs/Prueba Agosto-2/Resultado/Resultado Manual por el administrativo/Remuneracion 202608-2 Total_7721.xlsx` (`Reporte Componentes R1!F199 = F140+F114-L114`, filas `F140`/`F114` = subtotales-dato `NUEVO ESQUEMA`).
> **Numeración:** `plans/` 01..33 ocupados (33 + anexo `33-T0-Evidencia-Agosto2.md`); este plan toma el primer correlativo libre, **34**.
> **Estado:** APROBADO por el Ingeniero 2026-10-08 (con D-H cero-quema-de-períodos y catálogo-C abierto vigentes) — listo para `implementer` usando SOLO este documento.
> **Fecha:** 2026-10-08

---

## 0. Clarification Gate

**Sin preguntas bloqueantes: los veredictos vienen adoptados por encargo (H1 confirmada, H2/H3 descartadas — no se re-diagnostica) y las restricciones de diseño vienen cerradas (D-H cero-quema-de-períodos + catálogo-C abierto por doctrina Plan 32 D-B).** Los dos puntos con recomendación (mecanismo de equivalencia, D-A; alcance del gate rojo-primero, D-C) vienen redactados en la ruta recomendada con su verificación de cierre dentro de la tarea. Punto de bloqueo respetado: NO se implementa nada en este documento (planificación read-only: cero cambios a código de producción en este documento).

### 0.1 Diagnóstico adoptado (causa cerrada, no re-diagnosticar)

Fuente: T4 (`plans/33-T0-Evidencia-Agosto2.md`), se adopta tal cual como fuente única de verdad:

| # | Afirmación del veredicto | Reutilización en este plan | Re-verificación (el test, no la mano) |
|---|---|---|---|
| E1 | Síntoma: `ASE 2 Reporte Componentes R1!F199 [SUB_EMP]: faltan las anclas [Mes0] del sub-bloque; no se escribe parcial` — aborta el pipeline agosto-2026082 (`ERR-PLANTILLA`). | Síntoma que el gate T1 debe reproducir en rojo y que T2/T3 deben cerrar en verde. | T1: repro in-process con template 2026082 + fuente ASE2 real → throw con `[Mes0]`; T2/T3: misma corrida compone `F199` sin throw. |
| E2 | Causa raíz H1: la celda template `C199` trae `RECIPROCIDAD` (rótulo legado) mientras la fuente R1 ASE2-agosto trae la empresa como `NUEVO ESQUEMA` (nombre vigente); el filtro por nombre (`DerivarFilasPorFirma` → `R1FirmaInterior.EsDatoEmpresa(fila, empresa)`, `empresa = C199`) halla 0 filas oportunas → `MesOportuno` vacío → `AnclasFaltantes` reporta `Mes0`. | Causa raíz: no se re-investiga. El fix vive exactamente en ese matching. | T1 congela ambas caras del disco (C199 del template + conjunto fuente `{ENEL, NUEVO ESQUEMA, OCCIDENTE}`); T2 hace que el matcher las reconcilie. |
| E3 | H2 DESCARTADA: la fuente SÍ trae las filas oportunas del sub-bloque (11 filas-dato de empresa bajo `{ENEL, NUEVO ESQUEMA, OCCIDENTE}`). | El fix NO inventa filas ni relaja `AnclasFaltantes`: los datos existen, solo cambia la equivalencia de nombres. | T2: `F199` compuesto desde filas reales `NUEVO ESQUEMA`; ningún literal simulado. |
| E4 | H3 DESCARTADA: `CeldasInterioresAgosto[2]` (`F199`/`F204`/`F214`) coincide con la geometría real post-dimensionado (`B=LIME` fila 83, `TOTAL/OPORTUNO` fila 194, `C199/C204/C214 = RECIPROCIDAD/ENEL/OCCIDENTE` idénticos al manual). | El mapa NO se toca en este plan (ni sus celdas ni sus etiquetas). | T3: grep — diff vacío en `CeldasInterioresAgosto`. |
| E5 | Negocio: `Detalle de plantilla` («NUEVO ESQUEMA: Corresponde al recaudo "EAAB Reciprocidad"…»; «en LIME… ENEL, NUEVO ESQUEMA y OCCIDENTE»); `Proceso de Recaudo` (matriz EFC×ASE: ASE2 = ENEL+EAAB; facturación directa OCCIDENTE); `Prompt_Maestro §4–5` (conservar fórmulas / insert-delete preservando). | Anclaje normativo de la equivalencia: la fila de la tabla de sinonimia cita estas fuentes, no la intuición. | T4: cada fila de la tabla lleva su cita; sin cita no hay fila. |
| E6 | Secundario (NO causa del aborto): mapa ASE2 incompleto — el manual trae 5 subtotales-empresa (`RECIPROCIDAD` 199, `ENEL` 204, `ENERBIT` 209, `OCCIDENTE` 214, `CIUDAD LIMPIA-ACUEDUCTO` 219) y el mapa solo lista `F199/F204/F214` (omite `F209`/`F219`, ambos `F=0`, clase L-1 del TODO(T2-full) del Plan 32). | Follow-up declarado, NO entra en T5 (el plan lo deja explícitamente fuera para no expandir el diff). | T5-cierre: grep confirma que el mapa no creció en este diff. |
| E7 | Secundario: `G468` es celda interior agosto (`CeldasInterioresAgosto[4]`, `EXT_INT`) ADEMÁS de master shared `si=20` (Plan 33). | Coordinación obligatoria: T5 no toca `EscribirFormulaPreservando`/`Reanclar` y T3 re-verifica el invariante shared incl. `G468`. | T3: gate `SharedFormulaR1Gate` verde + asserts `G53`/`G468` intactos. |

### 0.2 Mapeo al Rector (Sale / No sale)

**Sale (EXPLÍCITO):**
- Re-diagnosticar H1/H2/H3 o re-abrir lectura, cálculo, espejo-dimensional, preflight, sellos, naming, calcChain, visibles o paridad fuera de lo declarado.
- `if (empresa == "RECIPROCIDAD") …` (o cualquier rama por nombre de empresa) en código de producción — viola la doctrina Plan 32 D-B (catálogo C ABIERTO: nunca `if ENEL/OCCIDENTE/…`); el encargo lo prohíbe explícitamente también como aliasing hardcodeado.
- Literales o ramas por período en el diff de producción (`2026072`/`2026082`/julio/agosto o equivalentes) — restricción D-H del Ingeniero; los períodos existen SOLO como datos de prueba.
- Tocar `CeldasInterioresAgosto` (E4: no está stale), `CompositorInteriorR1` (la forma `SUB_EMP` ya es correcta), `EscribirFormulaPreservando`/`Reanclar`/saneador calcChain (Plan 33/28 intactos).
- Renombrar `C199` en el template a `NUEVO ESQUEMA` — el template es del negocio y el rótulo legado es legítimo (histórico EAAB; el manual-oráculo también trae `RECIPROCIDAD`).
- Completar el mapa ASE2 (`F209` ENERBIT / `F219` CIUDAD LIMPIA-ACUEDUCTO, E6) — follow-up con su propio diseño, no deuda silenciosa (§5).
- Tocar cachés `<v>` como oráculo; tocar `<f>` de negocio fuera del `F199` recompuesto por firma; paquetes NuGet; commits/push.
- Inventar agregados de dominio por empresa para un gate numérico nuevo (la red numérica es comparador f-vs-f + fixture + goldens, doctrina Planes 29/32).

**No sale (entra, alcance congelado):**
| Frente | Superficie de cambio |
|---|---|
| B — equivalencia | Tabla de sinonimia como DATOS en Core + matcher puro `SonMismaEmpresa` + consumo en el overload `EsDatoEmpresa(fila, empresa)` (único punto de matching template↔fuente del interior); `MesOportuno`/`FilasAplic`/`FilaAnclaSubs` heredan el fix sin tocarse (ya filtran por ese overload) |
| G — gate TDD | Repro rojo-primero del aborto `F199 [Mes0]` con template 2026082 + fuente ASE2 reales (rutas explícitas, sin `Insumos.Raiz()`), verde post-fix con texto `=` manual |
| R — regresión | ASE1/3/4/5 + julio-identidad + goldens ±0.5 + DetRetri-D 5/5 + invariante shared (`G53`/`G468`) intacto |
| D — docs | `project-context.md` (doctrina sinonimia + reafirmación catálogo-abierto + lección rótulo-legado) |

**Fuera de scope explícito (ni diseño ni código en este plan):** espejo-motor R4; cobertura `conditionalFormatting`/`dataValidations` (W-4); semántica DetRetri (redondeo/columna D intactos); UI/CLI plumbing (la CLI solo como runner de repro manual); reconciliación de rutas fixture para los 511 rojos pre-existentes (dependencia declarada / follow-up: los gates de ESTE plan resuelven fixtures por ruta explícita como el gate shared del Plan 33, y NO dependen de la resolución raíz `Docs/Insumos`).

### 0.3 Decisiones ejecutivas (aprobación = aceptar estas)

| ID | Decisión |
|---|---|
| D-A | **Equivalencia por tabla de sinonimia como DATOS + matcher puro (recomendado).** Nueva tabla estática en Core (p. ej. `SinonimosEmpresaR1`: clases de equivalencia `{RECIPROCIDAD, NUEVO ESQUEMA}` con cita a `Detalle de plantilla` + `Proceso de Recaudo`) + `SonMismaEmpresa(a, b)` (normalización existente + pertenencia a la misma clase; desconocido = se compara consigo mismo) consumido por el overload `R1FirmaInterior.EsDatoEmpresa(fila, empresa)`. Añadir una divergencia futura = añadir una FILA con su cita de negocio, cero ramas, cero cambios de lógica. Descartados: aliasing hardcodeado en el mutador (viola D-B), alineación por orden/posición (atribuiría dinero a la empresa equivocada si el orden diverge — ver §2.2), normalización ciega sola (no resuelve strings disjuntos), renombrar el template (el rótulo legado es del negocio). |
| D-B | **El fix vive en UN punto (el overload), no en tres.** `MesOportuno` (§L1054), `FilasAplic` (§L1080) y `FilaAnclaSubs` (vía `MesOportuno`/`FilasAplic`) ya filtran por `EsDatoEmpresa(fila, empresa)`; al enseñar equivalencia al overload, los tres caminos (`SUB_EMP`, `EXT_INT`, `SUBS` con contexto empresa) quedan cubiertos sin tocar el mutador. `EsSubtotalEmpresa(fila, empresa)` (localizador lado-template) NO se toca: ambos lados del template hablan el mismo rótulo. |
| D-C | **Gate rojo-primero sobre el aborto real (TDD).** T1 reproduce `ASE 2 …!F199 [SUB_EMP]: faltan las anclas [Mes0]` con el template 2026082 + la fuente ASE2 reales (mismo método in-process del T4: `LeerEspejoR1` + `AjustarEnWorkbook` capturando el throw; o corrida `ProcesadorPeriodo` a temp FRESCA). Verde post-fix: `F199` compuesto con texto `=` manual (`F140+F114-L114` en filas reales del período) y forma `=` firma del sub-bloque. |
| D-D | **Julio-identidad como invariante de reversión.** En julio no existe divergencia (el pase interior es NO-OP por `EsPeriodoJulio` + geometría julio canónica); el diff T5 debe producir julio byte-idéntico (f-vs-f + literales). Si julio difiere, se revierte el pase, no el golden. |
| D-E | **Coordinación shared-formula (E7).** T5 no toca `EscribirFormulaPreservando`, `Reanclar`, ni el saneador calcChain; T3 re-corre el gate `SharedFormulaR1Gate` (workbook-wide + `G53`/`G468`) como red ciega. |
| D-F | **D-H cero-quema-de-períodos (restricción del Ingeniero, irrenunciable).** El diff de producción contiene CERO literales/ramas por período; los períodos existen SOLO como datos de prueba (rutas de fixtures, comandos de repro). Criterio de aceptación: `grep` de literales de período en el diff = 0 y el mecanismo no conoce ningún período (la tabla es de empresas, no de períodos). `EsPeriodoJulio` y `CeldasInterioresAgosto` quedan intactos (ni se mejoran ni se empeoran en este plan). |
| D-G | **Fail-fast enriquecido, no silenciado.** Si la equivalencia aún deja 0 filas (empresa futura sin fila de sinonimia), el throw sigue existiendo pero nombra AMBOS lados (rótulo template + etiquetas distintas observadas en la fuente del bloque) para que el próximo T0 arbitre con evidencia en vez de adivinar. Nunca 0 silencioso, nunca composición parcial. |

---

## 1. PROPOSE

### 1.1 Intent

Que el pipeline agosto-2026082 deje de abortar en `ASE 2 Reporte Componentes R1!F199 [SUB_EMP]` porque el matcher interior aprende —por datos, no por ramas— que el rótulo legado del template `RECIPROCIDAD` y el nombre vigente de la fuente `NUEVO ESQUEMA` son la misma empresa (EAAB Reciprocidad), componiendo `F199` desde las filas reales con texto idéntico al manual; que cualquier divergencia legado↔vigente futura de CUALQUIER empresa en CUALQUIER mes se resuelva añadiendo una fila citada a la tabla (sin cambios de lógica); con julio byte-idéntico y el resto de ASE intactos.

### 1.2 In Scope

- Tabla de sinonimia + matcher puro en Core + consumo en el overload `EsDatoEmpresa(fila, empresa)` (D-A/D-B).
- Gate TDD rojo-primero del aborto `F199 [Mes0]` con insumos reales agosto (D-C).
- Cierre: `F199` texto `=` manual + forma `=` firma + regresión 5 ASE + julio-identidad + goldens ±0.5 + DetRetri-D 5/5 + shared intacto (D-D/D-E).
- Fail-fast enriquecido (D-G) + docs mínimos.

### 1.3 Non-Goals

Ver §0.2 "Sale" + fuera de scope explícito. En particular: NO se generaliza el interior por-período (sucesor de `CeldasInterioresAgosto`, con su TODO(T2-full) vivo); NO se completa el mapa ASE2 (`F209`/`F219`); NO se toca `ObtenerLabelFuente` (camino del leaf-reader por `empresaId`, distinto del matching interior por rótulo — se deja intacto para no expandir el diff); NO se reconcilian las rutas raíz de los 511 rojos pre-existentes.

### 1.4 Riesgos

| ID | Riesgo | Mitigación |
|---|---|---|
| R-SINONIMO-ABUSO | La tabla se usa como cajón de alias inventados sin cita de negocio y encubre divergencias reales de la fuente | D-A: cada fila exige cita (`Detalle de plantilla` / `Proceso de Recaudo` / manual); sin cita no hay fila; el gate congela la tabla (drift-test) |
| R-ATRIBUCION | La equivalencia atribuye filas a la empresa equivocada (dinero movido entre empresas) | La equivalencia solo UNE rótulos de la misma EFC probada por el negocio (E5); el assert es texto `=` manual término a término (`F199` contra `F140+F114-L114` del oráculo) + goldens ±0.5; las demás empresas del bloque (`ENEL`, `OCCIDENTE`) siguen con matching exacto |
| R-JULIO | El overload con equivalencia altera julio (donde no hay divergencia) | D-D: equivalencia es conservadora (idéntica sin filas de sinonimia aplicables) + 0-diff R1 julio (f-vs-f + literales) por tarea; ante diff se revierte |
| R-MAPA-OCULTO | `F209`/`F219` (E6) u otra celda fuera del mapa impide el verde end-to-end aunque `F199` cierre | Alcance declarado: el gate T1/T2 corre a nivel sub-bloque `F199` (no end-to-end ciego); el cierre end-to-end se mide contra el estado pre-T5 (el aborto) — si otra celda aborta después, su propio throw nombrado lo delata y sale como follow-up, no como regresión de T5 |
| R-SHARED | T5 regresiona `G468` (interior + master `si=20`) u otro master shared | D-E: `EscribirFormulaPreservando` intacto (grep) + gate shared verde en T3 como red ciega |
| R-PERIODO | Se cuela un literal/rama de período en el diff (violación D-H) | D-F: grep de cierre sobre el diff (cero literales `2026xxx`/julio/agosto); la tabla es de empresas, el matcher no recibe ningún período |
| R-FIXTURE-ROOT | Los gates nuevos dependen de `Insumos.Raiz()` (roto post-reorg `a867706`) y heredan los 511 rojos | Resolución por ruta explícita estilo `SharedFormulaR1GateTests.Raiz()` (por `AGENTS.md`) + `Docs/Prueba Agosto-2/…` directo; prohibido `Insumos.Raiz()` en lo nuevo |

---

## 2. DESIGN

### 2.1 Enfoque: enseñar equivalencia al overload, no ramas al mutador; gatear el aborto, no el dominio

**Frente B (fix).** Nuevo tipo puro en Core (p. ej. `Remuneracion.Core/Models/SinonimosEmpresaR1.cs`): clases de equivalencia como DATOS — `IReadOnlyList<IReadOnlySet<string>>` con la única fila inicial `{ "RECIPROCIDAD", "NUEVO ESQUEMA" }`, cada fila con comentario de cita (`Detalle de plantilla`: «NUEVO ESQUEMA = recaudo EAAB Reciprocidad»; `Proceso de Recaudo`: ASE2 = ENEL+EAAB+OCCIDENTE; oráculo manual `F199=F140+F114-L114` con filas `NUEVO ESQUEMA`) — + `SonMismaEmpresa(string a, string b)`: normaliza (mayúsculas invariantes, sin diacríticos — misma regla que `Normalizar` del mutador / `OrdinalIgnoreCase` de las firmas) y devuelve `true` si son iguales o pertenecen a la misma clase; cualquier rótulo no tabulado solo se iguala a sí mismo (catálogo abierto por construcción). Consumo: el overload `R1FirmaInterior.EsDatoEmpresa(fila, empresa)` pasa de `string.Equals(fila.C, empresa, OrdinalIgnoreCase)` a `SinonimosEmpresaR1.SonMismaEmpresa(fila.C, empresa)`. Ese overload es el único punto donde el rótulo-template se confronta con el rótulo-fuente en el interior (`MesOportuno` §L1054-1073, `FilasAplic` §L1080-1099 rama no-global, `FilaAnclaSubs` vía ambos) — D-B: el mutador NO se toca salvo el fail-fast enriquecido (D-G, mismo `throw` de `RecomponerInterior` §L954-960 con el detalle añadido). `EsSubtotalEmpresa(fila, empresa)` intacto (ambos lados hablan rótulo-template).

**Frente G (gate).** Nuevo test TDD en `Remuneracion.IntegrationTests` (mismo método in-process del T4 + resolución de rutas explícita del gate shared Plan 33): (a) abre los bloques R1 de la fuente agosto real con `ExcelDataReaderWorkbookLeafInputReader.LeerEspejoR1` y corre `OpenXmlEspejoR1Mutador.AjustarEnWorkbook` sobre la base canónica 2026082 — pre-fix: throw `…F199 [SUB_EMP]: faltan las anclas [Mes0]…` (ROJO que reproduce E1); (b) post-fix: sin throw, `F199` con texto `=` manual y forma `=` firma del sub-bloque `NUEVO ESQUEMA` (VERDE). Fixture que manda: `C199=RECIPROCIDAD` (template), conjunto fuente ASE2 `{ENEL, NUEVO ESQUEMA, OCCIDENTE}`, `F199` manual `F140+F114-L114`. Prohibido `<v>`; prohibido `Insumos.Raiz()` (R-FIXTURE-ROOT).

**Frente R (regresión).** Post-fix, corrida 5-ASE agosto a temp FRESCA + julio a temp FRESCA: julio 0-diff R1 (f-vs-f + literales) vs golden/manual (D-D); agosto ASE1/3/4/5 sin cambios de texto salvo lo que el veredicto autorice (solo ASE2-F199 y sus espejos de columna G/H del mismo sub-bloque si aplican — el fixture T1 los congela); goldens ±0.5; DetRetri-D 5/5 vs R10 ambos períodos; `SharedFormulaR1Gate` verde (`G53`/`G468`, D-E).

### 2.2 Dónde vive la equivalencia y trade-offs

| Eje | Elegida | Descartada |
|---|---|---|
| Qué aprende el sistema | Tabla de sinonimia como DATOS (filas citadas) + matcher puro en Core; añadir divergencia futura = añadir fila, cero lógica (D-A) | `if (empresa == "RECIPROCIDAD") …` en el mutador: rama por nombre, viola Plan 32 D-B y el encargo; cada divergencia futura exigiría otra rama |
| Dónde se consume | Overload `EsDatoEmpresa(fila, empresa)` (único punto template↔fuente; cubre `SUB_EMP`/`EXT_INT`/`SUBS` sin tocar el mutador, D-B) | Parche solo en la rama `SUB_EMP` de `DerivarFilasPorFirma`: dejaría `EXT_INT`/`SUBS` del mismo sub-bloque con el mismo defecto latente (misma trampa que E7 advirtió para visible/interior) |
| Alternativa posicional | Rechazada: alinear k-ésimo `SUB_EMP` del template ↔ k-ésima empresa distinta de la fuente por orden (sin mirar nombres) | Parece "más abierta", pero el orden template↔fuente no es contractual (el mapa ASE2 ya omite `F209`/`F219`: el k-ésimo no corresponde) y un desorden movería dinero entre empresas en silencio; la equivalencia nominal mantiene la atribución auditada término a término contra el manual |
| Normalización ciega sola | Rechazada como suficiente (se conserva como capa base): `Normalizar` (mayúsculas + sin diacríticos) NO une `RECIPROCIDAD`↔`NUEVO ESQUEMA` (disjuntos); necesaria pero no suficiente | — |
| Renombrar template | Rechazado: `C199` es del negocio (el manual-oráculo también dice `RECIPROCIDAD`); reescribir el template canónico rompería f-vs-f contra el manual y la doctrina "el template manda en rótulos" | — |
| Extender `ObtenerLabelFuente` | Rechazado en este plan: camino distinto (leaf-reader por `empresaId`, ya mapea `1+ASE2/4→NUEVO ESQUEMA` por ids); tocarlo expande el diff y mezcla dos doctrinas (ids cerrados vs rótulos abiertos) | Se deja intacto y se documenta la diferencia en T4-docs |

### 2.3 Contratos afectados (resumen)

| Contrato | Cambio |
|---|---|
| `SinonimosEmpresaR1` (nuevo, Core puro, sin deps) | Tabla citada + `SonMismaEmpresa(a, b)`; drift-test del contenido |
| `R1FirmaInterior.EsDatoEmpresa(fila, empresa)` | `Equals` → `SonMismaEmpresa` (único cambio de lógica de producción; `EsDatoEmpresa(fila)` sin empresa y `EsSubtotalEmpresa*` intactos) |
| `OpenXmlEspejoR1Mutador` (`MesOportuno`/`FilasAplic`/`FilaAnclaSubs`/`CeldasInterioresAgosto`/`Reanclar`/`EscribirFormulaInterior`) | Intactos, salvo el mensaje del throw §L954-960 (D-G: añade rótulo-template + etiquetas fuente observadas del bloque) |
| `CompositorInteriorR1` / `ValidadorTotalesR1Workbook` / `SaneadorCadenaCalculo` / sellos / naming / preflight / readers / cálculo | Intactos |
| Gate nuevo `F199SinonimoTests` (+ fixture de citas) | Nuevo en `Remuneracion.IntegrationTests` (BCL/lectura in-process; rutas explícitas; prohibido `<v>` e `Insumos.Raiz()`) |
| `ComparadorSalidaVsManualTests.BrechaDeclarada` | Sin cambios (R-D-5 sigue su curso del Plan 32; T5 solo cierra el aborto, no versiona brechas) |
| `project-context.md` | Doctrina sinonimia (D-A/D-B) + reafirmación catálogo-abierto + lección rótulo-legado↔nombre-vigente + nota E6/E7 |

### 2.4 Architecture Validation Certificate

| Principio | Veredicto |
|---|---|
| **SRP** | ✅ Una razón de cambio por costura: Core declara equivalencia (datos + predicado puro), el overload la consume, el gate reproduce el aborto, el fixture fija las citas, el comparador/goldens vigilan. Ninguna rama por empresa ni por período en el fix. |
| **OCP** | ✅ Una divergencia futura se resuelve añadiendo una fila citada a la tabla, sin tocar lógica ni reabrir este plan; un rótulo no tabulado se comporta exactamente como hoy (igualdad consigo mismo) — julio y las demás empresas son identidad por construcción. |
| **DIP** | ✅ El matcher es puro (strings → bool, sin I/O ni OpenXML); el mutador sigue dependiendo de la firma de Core, sin dependencias nuevas; los tests consumen insumos reales + BCL ya presente. |
| **Best practices** | ✅ TDD rojo-primero sobre el aborto real con insumos del disco; fail-fast enriquecido que nombra ambos lados (D-G); quincena por dominio; `<f>` de negocio y `<v>` intactos; tolerancia ±0.5; sin Excel/COM en pipeline ni tests; Visual Design Intent N/A (cero UI). |
| **Performance** | ✅ Matching O(clases × miembros) despreciable frente al pase; la tabla inicial tiene 1 clase × 2 miembros; cero impacto en julio (mismo camino, igualdad rápida primero). |

---

## 3. SPEC

### 3.1 Requirements

| ID | Requisito | Grupo | Criterio de aceptación |
|---|---|---|---|
| R-B-1 | Tabla `SinonimosEmpresaR1` con la clase `{RECIPROCIDAD, NUEVO ESQUEMA}` citada (E5) + `SonMismaEmpresa` (normalización + clase; desconocido = identidad) | B | Drift-test verde: la tabla contiene exactamente las clases citadas; `SonMismaEmpresa("RECIPROCIDAD","NUEVO ESQUEMA")==true`, `("ENEL","OCCIDENTE")==false`, `("ENEL","ENEL")==true`, caso/acentos cubiertos |
| R-B-2 | Overload `EsDatoEmpresa(fila, empresa)` usa equivalencia; resto de `R1FirmaInterior` intacto | B | Unitarios puros por rama (sin I/O): fila `NUEVO ESQUEMA` matchea empresa `RECIPROCIDAD` y viceversa; `ENEL` no matchea `OCCIDENTE`; `EsSubtotalEmpresa` sin cambios (diff) |
| R-B-3 | `F199` agosto-ASE2 se compone desde filas reales `NUEVO ESQUEMA` con texto `=` manual y forma `=` firma; sin throw; sin literales simulados | B | Gate D-C verde: `F199` texto `=` `F140+F114-L114` del oráculo (en filas reales del período) + `MesOportuno(RECIPROCIDAD)` no vacío + `AnclasFaltantes(SUB_EMP,…)` vacío |
| R-B-4 | D-H: cero literales/ramas por período en el diff de producción | B | `grep` sobre el diff: 0 ocurrencias de `2026071/2026072/2026082`/julio/agosto (cualquier capitalización) en `Remuneracion.Core/` + `Remuneracion.Infrastructure/`; el matcher no recibe ningún período |
| R-B-5 | Catálogo abierto: ninguna rama por nombre de empresa en producción | B | `grep`: el diff no introduce `== "ENEL"/"OCCIDENTE"/"RECIPROCIDAD"/"NUEVO ESQUEMA"/…` fuera de las FILAS de la tabla (datos citados, no ramas) |
| R-G-1 | Gate rojo-primero reproduce `F199 [SUB_EMP] [Mes0]` pre-fix con template 2026082 + fuente ASE2 reales | G | T1 en rojo con el mensaje exacto del throw (ASE + hoja + celda + `[Mes0]`); post-fix el mismo test en verde (S1/S2) |
| R-G-2 | Fixtures del gate por ruta explícita (estilo Plan 33), independientes de `Insumos.Raiz()` | G | Los tests resuelven `Docs/Prueba Agosto-2/…` + `Docs/Prueba Julio-2/…` directo; `grep`: cero usos de `Insumos.Raiz()` en lo nuevo; regen SIEMPRE a temp FRESCA |
| R-R-1 | Julio-identidad: julio byte-idéntico (f-vs-f + literales R1) + goldens ±0.5 + DetRetri-D 5/5 ambos períodos | R | 0-diff R1 julio vs golden/manual; suite goldens verde; DetRetri 5/5 vs R10 (julio y agosto donde aplique) |
| R-R-2 | ASE1/3/4/5 sin cambios de texto atribuibles a T5 + shared intacto (`G53` `si=0`, `G468` `si=20`, workbook-wide) | R | Diff de textos R1 agosto: solo celdas del sub-bloque `RECIPROCIDAD`/`NUEVO ESQUEMA` de ASE2 se mueven (hacia el manual); `SharedFormulaR1Gate` verde |
| R-R-3 | Suite verde + build 0 warnings por tarea; fail-fast enriquecido (D-G) sin silenciar | R | `dotnet build …slnx` 0 warnings + `dotnet test` por tarea; el throw enriquecido nombra rótulo-template + etiquetas fuente del bloque |
| R-R-4 | Docs: doctrina sinonimia + catálogo-abierto + lección en `project-context.md` | R | Diff acotado de docs; cada fila de la tabla remite a su cita de negocio |

### 3.2 Scenarios (Given/When/Then)

- **S1 (TDD rojo — reproduce el aborto):** Given base canónica 2026082 + fuente R1 ASE2 real, When repro in-process (`LeerEspejoR1` + `AjustarEnWorkbook`), Then throw `ASE 2 Reporte Componentes R1!F199 [SUB_EMP]: faltan las anclas [Mes0]…` (pre-fix en rojo; post-fix este throw desaparece).
- **S2 (verde — equivalencia compone):** Given lo mismo post-fix, When se inspecciona `F199`, Then texto `=` manual (`F140+F114-L114` en filas reales) y `MesOportuno(bloque, "RECIPROCIDAD")` = filas `NUEVO ESQUEMA` del bloque (no vacío).
- **S3 (no-atribución cruzada):** Given bloque ASE2, When matching con empresa `ENEL`, Then solo filas `ENEL` (la equivalencia no arrastra `NUEVO ESQUEMA` ni `OCCIDENTE`); `F204`/`F214` componen como antes salvo re-anclaje ya vigente.
- **S4 (julio identidad):** Given período julio + base julio, When flujo 5-ASE + pase interior, Then diff R1 vs golden/manual = 0 (fórmulas y literales) y goldens ±0.5 verdes.
- **S5 (fail-fast enriquecido):** Given empresa futura sin fila de sinonimia y rótulos divergentes, When el sub-bloque queda en 0 filas, Then throw con rótulo-template + etiquetas fuente observadas (no genérico mudo, no composición parcial).
- **S6 (D-H):** Given el diff de producción T5, When `grep` de literales/ramas por período, Then 0 ocurrencias; y When `grep` de comparaciones por nombre de empresa fuera de la tabla, Then 0 ocurrencias.
- **S7 (shared intacto):** Given salida agosto post-fix, When `SharedFormulaR1Gate`, Then verde workbook-wide con `G53` (`si=0`, `ref=G53:AP53`) y `G468` (`si=20`, `ref=G468:AO468`) preservados.

---

## 4. TASKS

| ID | Tarea | Grupo | Depende de | Verificación |
|---|---|---|---|---|
| T1 | Gate-que-falla-primero (TDD) + fixtures que mandan. (a) Congelar del disco: `C199=RECIPROCIDAD` (template 2026082), conjunto fuente ASE2 `{ENEL, NUEVO ESQUEMA, OCCIDENTE}` (11 filas-dato), `F199` manual `F140+F114-L114` con filas `NUEVO ESQUEMA`, citas E5. (b) Nuevo test de repro in-process (rutas explícitas estilo Plan 33, temp FRESCA, prohibido `<v>` e `Insumos.Raiz()`) que exige el throw `F199 [SUB_EMP] [Mes0]` pre-fix (S1 en rojo) + drift-test de la tabla futura (vacío hoy: documenta el formato de fila citada). | G | — | R-G-1/R-G-2 (S1 rojo; resto de la suite sin nuevos rojos salvo este); build 0 warnings |
| T2 | Fix mínimo (datos + 1 overload + throw enriquecido). (a) `SinonimosEmpresaR1` en Core (1 clase citada + `SonMismaEmpresa` puro). (b) Overload `EsDatoEmpresa(fila, empresa)` → equivalencia (único cambio de lógica). (c) Throw §L954-960 enriquecido (D-G: + rótulo-template + etiquetas fuente del bloque). Unitarios puros R-B-1/R-B-2 + gate S1→verde y S2/S3 verdes. `CeldasInterioresAgosto`, `CompositorInteriorR1`, `EscribirFormulaInterior`, `Reanclar`, saneador: intactos (grep). | B | T1 | R-B-1..R-B-5 (S2/S3/S5/S6 parcial-B); diff = tabla + overload + mensaje + tests + gate; build 0 warnings |
| T3 | Cierre de regresión. Regens FRESCAS a temp (agosto 5-ASE + julio 5-ASE): `F199` texto `=` manual + forma `=` firma; ASE1/3/4/5 sin movimientos ajenos al veredicto; julio 0-diff R1; goldens Q1+Q2 ±0.5; DetRetri-D 5/5 vs R10 ambos períodos; `SharedFormulaR1Gate` verde incl. `G468` (S7); test estructural calcChain verde; greps de cierre (D-H + sin ramas por empresa + ningún `<v>` tocado + mapa intacto). | R | T2 | R-R-1/R-R-2/R-R-3 (S4/S6/S7); suite completa verde salvo los 511 rojos pre-existentes de fixture-root (se listan, no se tocan) |
| T4 | Docs: `project-context.md` (doctrina sinonimia D-A/D-B + catálogo-abierto reafirmado + lección rótulo-legado↔nombre-vigente con citas E5 + notas E6/E7 como follow-ups/coordinación) — sin código en el mismo commit. | D | T3 | R-R-4; diff acotado; CRLF |
| T5 | Cierre global: build 0 warnings + suite (base + nuevos) verde con la única excepción declarada de los 511 rojos de fixture-root (dependencia, §0.2) + goldens ±0.5 + DetRetri-D 5/5 + grep D-H global + CRLF + sin commits. | Transversal | T1..T4 | R-R-3; DoD §5 |

**Orden sugerido (TDD estricto):** T1 (rojo que reproduce `F199 [Mes0]` + fixtures que mandan) → T2 (equivalencia que lo cierra) → T3 (regresión 5 ASE + julio-identidad + shared) → T4 → T5. Build 0 warnings antes de marcar cada tarea completa. Sin commits (los hace el Ingeniero con `#commit`/`#push`).

---

## 5. Resumen ejecutivo, riesgos y decisiones

- **Resumen:** el T4 cerró la causa (H1: `C199=RECIPROCIDAD` legado vs fuente `NUEVO ESQUEMA` vigente = misma EFC EAAB Reciprocidad; H2/H3 descartadas; mapa y geometría sanos). Este plan convierte ese veredicto en: una tabla de sinonimia como DATOS con cita de negocio + un matcher puro consumido por el único punto donde el rótulo-template se confronta con el rótulo-fuente (overload `EsDatoEmpresa(fila, empresa)` — con lo que `SUB_EMP`/`EXT_INT`/`SUBS` del sub-bloque quedan cubiertos sin tocar el mutador), un gate TDD que hoy fallaría con el throw exacto, y un cierre con julio-identidad + shared intacto (`G468`) + goldens ±0.5. El diff de producción no conoce ningún período (D-H) ni enumera empresas en ramas (D-B/D-A).
- **Riesgo principal:** R-ATRIBUCION (mover dinero entre empresas con una equivalencia mal citada). Contenido por triple red: una sola clase inicial con doble cita de negocio (E5) + assert texto `=` manual término a término + goldens ±0.5; el resto del bloque sigue con matching exacto.
- **Decisión para el Ingeniero:** ratificar D-A..D-G (§0.3). **Forks con recomendación:** D-A (tabla-como-datos vs `if RECIPROCIDAD` vs alineación posicional vs normalización sola vs renombrar template — recomendada tabla: única que respeta catálogo-abierto + D-H y generaliza sin cambios de lógica) y alcance (T5 solo `F199`-camino + equivalencia; `F209`/`F219` y sucesor por-período del mapa quedan follow-ups, no deuda silenciosa). **Sin preguntas bloqueantes pendientes.**

### DoD (Definition of Done)

1. `dotnet build Remuneracion.WinForms/Remuneracion.WinForms.slnx` con 0 warnings.
2. Gate nuevo: rojo pre-fix (throw `F199 [SUB_EMP] [Mes0]` con insumos reales) → verde post-fix (`F199` texto `=` manual, forma `=` firma, S1/S2/S3/S5).
3. Julio 0-diff R1 (f-vs-f + literales) + goldens Q1+Q2 ±0.5 + DetRetri-D 5/5 vs R10 (julio y agosto donde aplique) + `SharedFormulaR1Gate` verde incl. `G468` (S4/S7).
4. Greps de cierre: 0 literales/ramas por período en el diff (D-H, S6); 0 comparaciones por nombre de empresa fuera de las filas citadas de la tabla (R-B-5); mapa/ `CompositorInteriorR1`/ `EscribirFormulaPreservando`/saneador intactos.
5. `project-context.md` actualizado; regens siempre a temp FRESCA; sin commits del agente; CRLF; sin emojis.

### Follow-ups explícitos (fuera de este plan)

- E6: completar mapa ASE2 (`F209` ENERBIT / `F219` CIUDAD LIMPIA-ACUEDUCTO) — su propio diseño (clase L-1 / TODO(T2-full) del Plan 32).
- Sucesor por-período de `CeldasInterioresAgosto` (generalización declarada en Plan 33 D-H) — su propio T0.
- Reconciliación de rutas raíz de fixtures (los 511 rojos pre-existentes por `Docs/Insumos` post-reorg `a867706`) — dependencia declarada, no tocada aquí; lo nuevo usa rutas explícitas.
- Espejo-motor R4, R4-por-empresa Q2, CF/DV del reanclaje (W-4), `R-EXTRA-CONCEPTO` (Plan 23), validación de nombres de hoja en preflight, remesh R2/R4, re-encendido SALDOS/AJUSTES (F-T4-2/3), `INTERVENTORIA!R26` (F-T4-4): vivos, intactos.
