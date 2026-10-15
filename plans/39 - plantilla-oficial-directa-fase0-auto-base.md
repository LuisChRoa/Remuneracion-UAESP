# Plan 39 — Plantilla oficial directa: Phase 0 automática de base + L9:L13 por quincena (R-INTERV-Q)

> **Alcance:** la plantilla oficial `Docs/Plantilla estandarizada/Formato Remuneracion 2026xx-x Total.xlsx` pasa a ser el INPUT DIRECTO que consume la App para TODOS los períodos — se elimina el paso manual intermedio de `Docs/BasesPreparadas/` del flujo de producción. Al arrancar, el pipeline ejecuta una **Phase 0 automática de preparación de base** sobre la copia de trabajo (el oficial NUNCA se modifica in place): renombra las 3 hojas de período, reescribe los 11 tokens `<f>` de período, sella N3, Y selecciona la columna L9:L13 según la quincena (Q1→N, Q2→M). Reframing preciso del invariante: la preparación de base (mecánica, tokens de período + selección de columna L, fase explícita y auditada) SÍ puede tocar `<f>`; el PROCESAMIENTO de negocio (slot writer values-only y todo lo downstream) NUNCA. Cambio de producción esperado: **mínimo y acotado** (una fase nueva que MUEVE —no reescribe— el código probado de `PreparadorBasePeriodo --plantilla-unica`, + selección L por quincena parametrizada por dominio, + gates; cero cambios de semántica de negocio, cero cambios de flujo de lectura/cálculo/escritura).
> **Rector (IRRENUNCIABLE):** `README.md` + `Docs/Propuesta_Proyecto_Automatizacion_Remuneracion_UAESP.md` + `Docs/Detalle de plantilla.docx` + `Docs/Proceso de Recaudo.docx` + `Docs/Prompt_Maestro_Proyecto_Remuneracion_UAESP _ Vo .docx` + `.opencode/project-context.md`. **Invariantes:** tolerancia ±0.5 en goldens/comparaciones numéricas; NUNCA usar cachés `<v>` como oráculo (aquí `<f>` debe quedar byte-idéntica a la variante esperada por quincena — se aserta); doctrina calcChain Plan 28 intacta (delete calcChain + fullCalcOnLoad); slot doctrine Plan 36 intacta en procesamiento (NUNCA insert/delete filas, NUNCA escribir `<f>` en path de negocio); build 0 warnings antes de marcar cualquier tarea completa; CRLF sin mezclar; NUNCA git commit/push (los hace el Ingeniero con `#commit`/`#push`); la implementación la ejecuta `implementer` usando SOLO este plan aprobado; `code-reviewer` verifica spec compliance + 0 warnings. **Quincena por dominio, nunca por contenido.**
> **Mandato (aprobado por el Ingeniero 2026-10-14):** (1) el template oficial es el input directo de la App para todos los períodos (adiós paso manual `BasesPreparadas`); (2) manejo dinámico de hojas de período gobernado por la selección del usuario (AñoMes + Quincena → `Informe AFaseo Recaudo {AAAAMM}-{Q}`, `DetRetri{AAAAMMQ}`, `DetValiRetri{AAAAMMQ}` vía `NombresHojaPeriodo`); (3) `CONSOLIDADO_TOTAL RECAUDO!L9:L13` dependiente de quincena: Q1 → `INTERVENTORIA!N26:N30` (Primera quincena), Q2 → `INTERVENTORIA!M26:M30` (Segunda quincena).
> **Continuidad:** Planes 21 (espejo — retirado para R1 por Plan 36), 25 (roles por firma — vigente), 26/27 (preflight + firma + hoja-por-nombre — vigentes), 28 (calcChain + sello + regla "SIEMPRE plantilla en ceros" — vigente; su D-G sin-guardrail se ENDURECE aquí a guardrail automático Phase 0), 29 (paridad + comparador f-vs-f — vigente), 30 (naming dinámico + base canónica + herramienta offline — vigente; su herramienta se MUEVE al pipeline como Phase 0), 31/32/35-partes de escritura (retiradas por Plan 36), 33 (shared-preservation — intacta), 34 (sinonimia como DATOS + D-H + catálogo abierto — vigentes), 36 (slot doctrine + writer values-only + `ValidadorBasePeriodo` preflight + `PreparadorBasePeriodo --plantilla-unica` — vigente; su §2.3 opción A offline se PROMUEVE a Phase 0 automática en-pipeline), 38 (recongelamiento geométrico + drift gate — vigente; su T0 se extiende con las 2 variantes L). Este plan NO reabre lectura R1/R2/R4-detalle, cálculo CONSOLIDADO, preflight de insumos, sellos G7/K7/D6/D7, naming, calcChain, paridad, ni el writer values-only fuera de lo declarado.
> **Insumos canónicos:** plantilla oficial `Docs/Plantilla estandarizada/Formato Remuneracion 2026xx-x Total.xlsx` (oráculo geométrico + input directo); `Herramientas/PreparadorBasePeriodo/Program.cs` (modo `--plantilla-unica`: código probado a MOVER, exit 0 con self-verification para 2026071/2026072/2026082); `NombresHojaPeriodo.cs` (fuente única de naming); `ValidadorBasePeriodo.cs` (preflight runtime sobre la copia auto-preparada); `OpenXmlPlantillaWriter.cs` (save paths + wiring `ValidarNoInPlace` + `ValidadorBasePeriodo.Exigir`); `WorkbookLeafCellMap*.cs` (+ `InterventoriaDeclarada.cs`); `Form1.cs` (campos Año/Mes/Quincena + `txtPlantilla` ya existen — se cablean, no se rediseñan); `Remuneracion.Cli` (`OpcionesCli`/`EjecutorCli`: `--plantilla` acepta el oficial directo, `--periodo` decide hojas/tokens/columna-L); los 3 `.docx` de negocio según necesidad. Regens SIEMPRE a rutas temp FRESCAS.
> **Numeración:** `plans/` 01..38 ocupados; este plan toma el primer correlativo libre, **39**.
> **Estado:** APROBADO por el Ingeniero 2026-10-14 (a: plantilla oficial = input directo para todos los períodos; b: Phase 0 automática moviendo —sin reescribir— el código probado + selección L por quincena Q1→N/Q2→M con el invariante re-encuadrado base-prep-SÍ/procesamiento-NUNCA; c: `Docs/BasesPreparadas/` SALE del flujo de producción y SE ELIMINA una vez que ningún consumidor la use — "lo que no se usa, borrar"). EN EJECUCIÓN.
> **Fecha:** 2026-10-14

---

## 0. Clarification Gate

**Evidencia adoptada, no re-diagnosticada (punteros del encargo; T0 congela sus propios números desde el disco).** El planificador verificó read-only en disco (2026-10-14): existe la plantilla única + las 3 bases `Docs/BasesPreparadas/Plantilla_Remuneracion_{2026071,2026072,2026082}.xlsx`; existe `PreparadorBasePeriodo/Program.cs` con modo `--plantilla-unica` (detección de período base por forma `DetRetri{7dígitos}`, reemplazo de tokens, sello N3, escritura atómica tmp+verificación+move, autoverificación embebida); existe `ValidadorBasePeriodo.cs` (hojas por `NombresHojaPeriodo` + N3==`CodigoCompleto` + sellos sin fórmula + 0 tokens stale en `<f>`); `OpenXmlPlantillaWriter` tiene los 3 `ValidarNoInPlace` (L127/L220/L395: copia origen→salida, nunca in-place) y el wiring `ValidadorBasePeriodo.Exigir` pre-escritura (L274-278); `Form1.cs` ya trae `cmbAnio/cmbMes/cmbQuincena` + `PeriodoSeleccionado` (`{anio}{mes}{quincena}`) + `txtPlantilla` + guard salida==plantilla; CLI ya trae `--periodo/--plantilla/--carpeta/--salida/--quincena`; `InterventoriaDeclarada.cs` declara M (Segunda) y N (Primera) por ASE. **Lo que T0 debe aún congelar desde bytes (no verificado por este plan):** el texto exacto de cada `<f>` L9:L13 en ambas variantes, cada celda-dato vs celda-fórmula del sweep §2.3, y cada referencia con fragmento Q1/Q2-fijo.

### 0.1 Evidencia adoptada (punteros — el fixture T0 es la autoridad)

| # | Afirmación adoptada | Reutilización en este plan | Re-verificación (el test, no la mano) |
|---|---|---|---|
| E1 | Template oficial congelado 2026-10-13, base interna 2026092: `CONSOLIDADO!L9:L13 = INTERVENTORIA!M26:M30` como `<f>` con cachés `<v>` 189185988/266580256/154788536/120391083/128990446 (= Segunda quincena por ASE); `N3 = 2026092` como VALOR plano (último dígito = quincena); headers `INTERVENTORIA` M='Segunda quincena' N='Primera quincena'; M31=859936309, N31=859936305, K31=1719872614 (M+N=mes, split con off-by-one por redondeo por ASE: 189185988 vs 189185987 etc. — la atribución de quincena es material, no cosmética). | Regla R-INTERV-Q (§2.2): Q1→N, Q2→M. T0 congela AMBAS variantes (texto-`<f>` + cachés esperados por variante). | T0: fixture con las 2 variantes citadas celda-a-celda; T-gates: drift acepta la variante de la quincena del run; tercera forma = FAIL. Prohibido `<v>` como oráculo de fórmulas (el caché solo documenta la variante, nunca la demuestra). |
| E2 | `PreparadorBasePeriodo --plantilla-unica` probado (exit 0 + self-verification para 2026071/2026072/2026082): reescribe tokens de período una vez + sella N3 + renombra las 3 hojas + actualiza defined names/fórmulas/app.xml + escritura atómica. | Se MUEVE (no se reescribe) a una fase compartida del pipeline (§2.1). La herramienta CLI queda como wrapper delgado del mismo código. | T1: paridad byte-a-byte herramienta-vs-fase (misma entrada → mismos bytes salvo la columna L parametrizada); tests existentes de la herramienta en verde sin duplicar. |
| E3 | Arsenal a reutilizar tal cual: `NombresHojaPeriodo` (única fuente de naming), `ValidadorBasePeriodo` (preflight runtime), guard `ValidarNoInPlace` (oficial NUNCA modificado in place), slot writer + firma/sinonimia/clasificador, slot doctrine (procesamiento NUNCA escribe `<f>`, NUNCA inserta/borra filas). | Phase 0 vive ANTES del preflight y ANTES de cualquier escritura de negocio; el preflight corre sobre la copia auto-preparada (red ciega). | T-gates: grep de cierre (0 escrituras `<f>` fuera de Phase 0; 0 insert/delete en procesamiento). |
| E4 | UX: los campos de período/plantilla existen en WinForms (logs lo confirman) y el CLI ya acepta `--plantilla/--periodo/--quincena`. | Se cablea, no se rediseña (§2.4). | T-gates: el período seleccionado por el usuario (no el contenido del workbook) gobierna hojas/tokens/columna-L en ambos frontends. |

### 0.2 Gate de entrada — pregunta única al Ingeniero (BLOQUEANTE para T1..Tn)

| ID | Pregunta | Efecto según respuesta |
|---|---|---|
| Q1 | ¿Ratifica que (a) la plantilla oficial es el input directo para todos los períodos (adiós paso manual `BasesPreparadas` en producción), (b) la preparación de base se automatiza como Phase 0 en-pipeline moviendo —sin reescribir— el código probado de `PreparadorBasePeriodo --plantilla-unica` + selección L9:L13 por quincena (Q1→N, Q2→M), con el invariante re-encuadrado "base-prep SÍ puede tocar `<f>` (fase mecánica auditada) / procesamiento NUNCA", y (c) `Docs/BasesPreparadas/` sale del flujo de producción (se declara: conservar como artefactos de debug o retirar)? | "Sí, proceder" = puerta a T0..T6. Si (c) pide conservar: entran como fixtures de referencia versionados con su README (nunca como input de producción); si pide retirar: su propio commit del Ingeniero fuera de este plan (el código nunca los referencia). |

### 0.3 Mapeo al Rector (Sale / No sale)

**Sale (EXPLÍCITO):**
- Re-diagnosticar E1..E4 o reabrir "template por mes vs plantilla única" (mandato #2067 + mandato 2026-10-14 lo cierran).
- Reescribir la lógica de `PreparadorBasePeriodo` (se MUEVE; cualquier reescritura exige su propio T0 + justificación).
- Escribir `<f>` en runtime fuera de Phase 0 (slot writer y todo lo downstream siguen values-only); insert/delete filas en procesamiento.
- Modificar el oficial in place (el guard `ValidarNoInPlace` se ENDURECE, nunca se relaja).
- Restaurar fórmulas viejas donde el template trae 0 intencional (allow-list Plan 36 intacta); tocar mirror/hidden/labels fuera de contrato.
- Ramas por nombre de empresa fuera de las FILAS citadas de `SinonimosEmpresaR1`; literales/ramas por período en el diff de producción (D-H); cachés `<v>` como oráculo; fórmula runtime `SI()` para elegir quincena (se descarta en §2.5); paquetes NuGet; commits/push.
- Renombrar rótulos del template en runtime.

**No sale (entra, alcance congelado):**
| Frente | Superficie de cambio |
|---|---|
| T0 — evidencia | Congelamiento desde bytes de AMBAS variantes L9:L13 (Q1→N / Q2→M, texto-`<f>` + cachés por variante) + sweep BCL de OTRAS refs duras de quincena (§2.3: cada hallazgo = intencional-por-quincena con su regla, o invariante) + N3/consumidores; cero código de producción |
| B1 — Phase 0 | Nueva fase de pipeline que MUEVE el código del preparador (3 hojas + 11 tokens `<f>` + N3 + columna-L por quincena) + copia oficial→trabajo + endurecimiento `ValidarNoInPlace`; la herramienta queda como wrapper delgado |
| B2 — R-INTERV-Q | Regla de negocio quincena-dependiente L + drift gate por variante + `ValidadorBasePeriodo` extendido (acepta la variante de la quincena, rechaza tercera forma) |
| B3 — cableado UX | WinForms (campos existentes → Phase 0) + CLI (`--plantilla` = oficial directo, `--periodo` decide hojas/tokens/L); sin rediseño |
| G — gates | `<f>`-idéntica-a-variante + DetRetri-D vs R10 ±0.5 ambos períodos + shared/calcChain ciegos + healthy-open + greps de cierre |
| D — docs | `project-context.md` (doctrina Phase 0 + R-INTERV-Q + sweep) + manual (corre desde el oficial: procedimiento nuevo) + destino declarado de `Docs/BasesPreparadas/` |

**Fuera de scope explícito (OUT):** espejo R4; cobertura `conditionalFormatting`/`dataValidations` (W-4); semántica DetRetri (redondeo/columna D intactos); F209/F219 (follow-up con su propio T0); reconciliación de fixture-root (dependencia declarada; lo nuevo usa rutas explícitas); restaurar fórmulas viejas; tocar mirror/hidden/labels fuera de contrato.

### 0.4 Decisiones ejecutivas (aprobación = aceptar estas)

| ID | Decisión |
|---|---|
| D-A | **El oficial es el input directo.** Ningún flujo de producción exige base preparada a mano. `Docs/BasesPreparadas/` sale del flujo (destino declarado en Q1). |
| D-B | **Phase 0 automática y auditada.** Copia oficial→trabajo + preparación mecánica (3 hojas + tokens `<f>` + N3 + columna-L) parametrizada SOLO por el dominio (`Periodo`: `CodigoCompleto`, `CodigoAAAAMM`, `NumeroQuincena` vía `NombresHojaPeriodo`). El oficial nunca se toca (guard endurecido). |
| D-C | **MOVER, no reescribir.** El código probado del preparador se traslada a una fase compartida (`Infrastructure`, BCL `ZipArchive`+`Regex`, sin OpenXML en la mutación mecánica); la herramienta CLI queda como wrapper delgado del mismo código. Paridad byte-a-byte exigida. |
| D-D | **Invariante re-encuadrado con precisión quirúrgica:** base-preparation (mecánica, tokens de período + selección L, fase explícita Phase 0) PUEDE tocar `<f>`; business PROCESSING (slot writer values-only y todo lo downstream) NUNCA toca `<f>` ni inserta/borra filas. El grep de cierre distingue ambas fases por namespace/clase. |
| D-E | **R-INTERV-Q como regla de negocio de primera clase.** L9:L13 es quincena-dependiente por dominio (no por contenido): Q1→N, Q2→M. T0 congela ambas variantes; el drift acepta la del run; tercera forma = FAIL. |
| D-F | **Sweep sin supuestos silenciosos.** Cada ref dura de quincena hallada por el barrido BCL se congela como intencional-por-quincena (con su regla citada) o se declara invariante. Ningún "debe ser igual en ambas" sin evidencia. |
| D-G | **D-H + catálogo abierto irrenunciables.** Cero literales/ramas por período y cero ramas por empresa en el diff; Phase 0 no recibe literales: recibe `Periodo`. Empresas no tabuladas = identidad. |

---

## 1. PROPOSE

### 1.1 Intent

Que la remuneración de CUALQUIER quincena se produzca corriendo la App directamente sobre la plantilla oficial — con la base preparada automáticamente (hojas + tokens + N3 + columna-L de la quincena), apertura sana en Excel, `<f>` byte-idénticas a la variante esperada, DetRetri-D validado contra R10 ±0.5, y fail-fast accionable cuando la base auto-preparada no corresponda al período — sin paso manual intermedio y sin tocar el oficial.

### 1.2 In Scope

- T0 de evidencia: ambas variantes L9:L13 + sweep de refs duras de quincena + N3/consumidores, todo citado celda-a-celda desde bytes (con Q1 respondida).
- Phase 0 (B1): mover el preparador al pipeline + columna-L por quincena + copia oficial→trabajo + `ValidarNoInPlace` endurecido.
- R-INTERV-Q (B2): regla + drift por variante + extensión de `ValidadorBasePeriodo`.
- Cableado UX (B3): WinForms + CLI sin rediseño.
- Regresión (T5): ambos períodos end-to-end DESDE EL OFICIAL (sin base preparada como input), salidas frescas a temp, healthy-open, `<f>`-idéntica-a-variante, DetRetri-D vs R10 ±0.5, shared/calcChain ciegos, preflight verde sobre la copia auto-preparada.
- Docs (T6): doctrina Phase 0 + R-INTERV-Q + sweep + manual "corre desde el oficial" + destino de `BasesPreparadas/`.

### 1.3 Non-Goals

Ver §0.3 OUT. En particular: NO se reescribe el preparador; NO se escribe `<f>` fuera de Phase 0; NO hay `SI()` de quincena en el template (§2.5); NO se restaura ninguna fórmula vieja; NO se tocan R4-espejo/CF-DV/DetRetri-semántica/F209-F219/fixture-root.

### 1.4 Riesgos

| ID | Riesgo | Mitigación |
|---|---|---|
| R-MOVER | Al mover el preparador se introduce una divergencia silenciosa (byte distinto a la herramienta) | D-C + paridad byte-a-byte (misma entrada → mismos bytes salvo columna-L parametrizada); tests existentes de la herramienta en verde; la herramienta queda como wrapper delgado (una sola implementación viva) |
| R-ATRIBUCION | La columna L equivocada atribuye dinero de la quincena contraria (material: off-by-one por ASE) | R-INTERV-Q: T0 congela ambas variantes con cachés citados; Phase 0 decide SOLO por `Periodo.NumeroQuincena` (dominio); drift acepta solo la variante del run; DetRetri-D vs R10 ±0.5 por ASE como red |
| R-SWEEP | Queda una ref dura de quincena sin congelar y un período futuro la hereda mal | Sweep BCL exhaustivo §2.3 (M/N, fragmentos Q1/Q2-fijos, J9/J-columnas, AJUSTES, R26, C59, N3-consumidores); cada hallazgo con regla o invariante; drift gate lo convierte en FAIL futuro, no en silencio |
| R-INPLACE | Un path escribe sobre el oficial (pérdida del contrato dorado) | `ValidarNoInPlace` endurecido (todos los caminos de guardado + Phase 0 con assert origen≠trabajo + oficial read-only por atributo en tests); grep de cierre |
| R-DOBLE-F | Phase 0 y procesamiento pisan la misma `<f>` (el re-encuadre se vuelve excusa) | D-D quirúrgico: allow-list EXPLÍCITA de `<f>` tocables solo por Phase 0 (11 tokens de período + 5 L9:L13); grep distingue fases; cualquier `<f>` fuera de la allow-list tocada por cualquier fase = FAIL |
| R-UX | El usuario selecciona período en UI pero el pipeline usa otro (hojas/tokens/L de período cruzado) | B3: una sola fuente (`PeriodoSeleccionado`→`Periodo.Parse`→`NombresHojaPeriodo`+Phase 0); CLI igual (`--periodo`→`Periodo`); `ValidadorBasePeriodo` verde sobre la copia como red ciega |
| R-BASES | `BasesPreparadas/` sigue usándose como input por costumbre y diverge del oficial | Q1 declara su destino (debug versionado con README o retiro por commit del Ingeniero); ningún código nuevo las referencia; el manual nuevo dice "corre desde el oficial" |
| R-JULIO | Al automatizar se altera un período que ya abría sano | Puerta de identidad: Q1 y Q2 desde el OFICIAL reproducen las salidas de referencia (DetRetri-D vs R10 ±0.5 + `<f>`-idéntica-a-variante); ante diff se revierte el pase, no el golden |
| R-FIXTURE-ROOT | Gates nuevos heredan rojos de resolución raíz | Rutas explícitas (oficial + fuentes por ASE por ruta explícita + temp FRESCA); prohibido `Insumos.Raiz()` y `Docs/Insumos` raíz en lo nuevo |

---

## 2. DESIGN

### 2.1 Enfoque: Phase 0 mecánica y auditada ANTES del preflight, procesamiento intacto DESPUÉS

**Fase 0 — preparación automática de base (T1/T2, código movido, no reescrito).** Orden del pipeline multi-ASE (y single-ASE donde aplique):

```
1. Resolver Periodo (dominio: UI cmbAño/Mes/Quincena → Periodo.Parse; CLI --periodo [--quincena])
2. ValidarNoInPlace ENDURECIDO (oficial ≠ trabajo; el oficial jamás se abre en modo escritura)
3. Copiar oficial → ruta de trabajo (File.Copy a rutaSalida o a temp intermedio si salida==destino final)
4. Phase 0 sobre la COPIA (BCL ZipArchive+Regex, sin OpenXML en la mutación mecánica):
   a. renombrar 3 <sheet name> al período (vía NombresHojaPeriodo: DetRetri{AAAAMMQ},
      DetValiRetri{AAAAMMQ}, Informe AFaseo Recaudo {AAAAMM}-{Q}) — rIds/sheetIds intactos
   b. reescribir los 11 tokens <f> de período (base detectada por forma DetRetri{7dígitos} o --base)
   c. sellar N3 = CodigoCompleto (valor plano; si N3 trae <f> → FAIL, nunca sobrescritura)
   d. NUEVO: setear L9:L13 → M-columna si quincena==2, N-columna si quincena==1
      (texto-<f> exacto de la variante T0; si L trae forma inesperada → FAIL)
   e. actualizar docProps/app.xml + defined names (código movido, intacto)
   f. autoverificación embebida (0 token viejo, N hojas, rIds intactos, N3 sellado,
      L-variante == esperada para la quincena) + escritura atómica (tmp+verificación+move)
5. ValidadorBasePeriodo.Exigir (EXTENDIDO: acepta la variante L de la quincena; tercera forma = FAIL)
6. Procesamiento INTACTO (slot writer values-only + hojas + sellos G7/K7/D6/D7 + calcChain) — NUNCA toca <f>
```

**Costura (D-C):** el cuerpo mecánico actual de `PreparadorBasePeriodo/Program.cs` (modo `--plantilla-unica`: `EjecutarPlantillaUnica` + `VerificarBaseUnica` + helpers `SetValorNumerico/TokenReplace/EscribirZipAtomico`) se traslada a una clase compartida de `Infrastructure` (p. ej. `Excel/PreparadorBasePipeline.cs`, BCL puro, sin dependencia OpenXML/ExcelDataReader), parametrizada por `(rutaOficial, periodo, rutaTrabajo)` — el período viaja como `Periodo`, nunca como literales (D-H). La herramienta queda como wrapper delgado (parseo de args + llamada a la fase + consola). La selección L (paso 4d) es el ÚNICO código nuevo de la fase, tabulado desde el fixture T0 (textos-`<f>` por variante, sin `if "M26"` literales sueltos: tabla `VarianteLPorQuincena` indexada por `NumeroQuincena`).

**Dónde se invoca:** en `OpenXmlPlantillaWriter.GenerarWorkbook` (ambos overloads multi-ASE + single-ASE) y `EscribirEspejoR1`, ENTRE el `File.Copy` y el `ValidadorBasePeriodo.Exigir` / `ValidarFormulasProtegidas`. Alternativa considerada y DESCARTADA: invocar desde `ProcesadorPeriodo` antes del writer — se descarta porque el writer es el dueño de las 3 rutas de guardado + atomicidad + `ValidarNoInPlace` (un hook en el procesador dejaría `EscribirEspejoR1` sin Phase 0). El procesador NO cambia (su `ValidadorBasePeriodo.Exigir` posterior actúa como red ciega sobre la copia auto-preparada).

### 2.2 Regla R-INTERV-Q (nueva regla de negocio de primera clase)

> **R-INTERV-Q:** `CONSOLIDADO_TOTAL RECAUDO!L9:L13` es quincena-dependiente por dominio. Si `Periodo.NumeroQuincena == 1` → `L9:L13 = INTERVENTORIA!N26:N30` (Primera quincena). Si `== 2` → `L9:L13 = INTERVENTORIA!M26:M30` (Segunda quincena). Cualquier otra forma en la copia auto-preparada o en la salida = FAIL (`ERR-PLANTILLA` nombrando hoja+celda+variante esperada).

Variantes a congelar por T0 (punteros E1, el fixture manda):

| Variante | Texto-`<f>` (a congelar exacto por T0) | Cachés `<v>` esperados por ASE (documentan, no demuestran) |
|---|---|---|
| Q1 → N | `L9:L13 = INTERVENTORIA!N26:N30` (texto exacto por T0) | 189185987 / 266580255 / 154788535 / 120391083 / 128990445 (= `PrimeraQuincenaPorAse`) |
| Q2 → M | `L9:L13 = INTERVENTORIA!M26:M30` (texto exacto por T0, = forma actual del oficial base 2026092) | 189185988 / 266580256 / 154788536 / 120391083 / 128990446 (= `SegundaQuincenaPorAse`) |

Notas normativas: (a) ASE4 es idéntico en ambas (120391083) — el gate NO puede distinguir Q1/Q2 por ASE4 solo; la aserción es por texto-`<f>`, no por valor. (b) El off-by-one ASE1/2/3/5 (…988 vs …987) es split de redondeo mensual M+N=K±1 — por eso la atribución es material: pegar la columna contraria mueve ~1 por ASE y rompe DetRetri-D vs R10 en el centavo equivocado. (c) `InterventoriaDeclarada` (M/N por ASE) se REUTILIZA como oráculo de cachés por variante, no se duplica. (d) La hoja `INTERVENTORIA` en sí sigue PROTEGIDA intacta (D2b vigente): Phase 0 solo re-apunta las `<f>` L del CONSOLIDADO, jamás escribe valores en INTERVENTORIA.

### 2.3 Sweep obligatorio de refs duras de quincena (deber de mentor — sin supuestos silenciosos)

Barrido BCL (`ZipArchive` + XML, sin Excel/COM) sobre la plantilla oficial, ejecutado por T0, que lista TODA referencia con olor a quincena. Cada hallazgo sale con veredicto citado: **intencional-por-quincena** (con su regla, p. ej. R-INTERV-Q) o **invariante** (igual en ambas, con motivo). Superficie mínima a barrer:

| # | Superficie | Qué buscar | Veredicto esperado (a confirmar por T0 con bytes) |
|---|---|---|---|
| S-MN | Refs columna M/N en todo el workbook | Toda `<f>` que mencione `!M` / `!N` + columna M/N (INTERVENTORIA y cualquier otra hoja) | L9:L13 = R-INTERV-Q; M31/N31/K32 = invariantes SUM (protegidas vigentes); resto citado uno a uno |
| S-VAL | Fragmentos Q1/Q2-fijos en `VALIDACION_*` | Literales de hoja/celda con sufijo de período o rama de quincena | Intencional (parametrizado por período vía mapas) o invariante — citado |
| S-J | `DetRetri` J9 / J-columnas | Refs a J (BCE/validación) con forma fija | Citado: ¿invariante o por-quincena? (NO se toca, solo se congela el veredicto) |
| S-DV | `DetValiRetri` (D21/D29/I/Z, J9 AJUSTE A LA DECENA externo) | Refs con forma fija de quincena | Invariante declarado (J9 SALE Plan 29) o intencional — citado |
| S-AJ | Cadenas AJUSTES (`AJUSTES - SF-T`, `SALDOS POR NOTA`, `RETRIBUCION NEGATIVA`, `CONSOLIDADO D85:D89`) | Refs/columnas con rama Q1/Q2 | Q1: `AjustesSfT=0` por dominio (invariante de cálculo); resto citado |
| S-R26 | `INTERVENTORIA!R26` (`=F15` vs `=H15`, veredicto R-D-6 Plan 31) | Texto actual en el oficial | Declarar-divergencia cosmética versionada o invariante — citado, no re-abierto |
| S-C59 | `C59` hoja banco (gate `banco.Quincena == NumeroQuincena`) | ¿La plantilla trae valor fijo o fórmula? | Invariante de escritura por dominio (el runtime la sella por dominio, nunca por contenido) — citado |
| S-N3 | Consumidores de `N3` (quién lee el código de período) | Toda `<f>` que referencie `!N3` o `N3` | Citado: Phase 0 sella N3=valor plano; consumidores leen, nunca escriben — invariante |

**Regla de cierre del sweep:** ningún hallazgo queda en "probablemente igual". Si T0 no puede arbitrar un hallazgo con bytes + citas, el hallazgo se eleva como follow-up con su propio T0 (no bloquea T1..T4 salvo que toque el camino L/N3/hojas).

### 2.4 UX: cablear, no rediseñar

| Frontend | Estado actual (verificado) | Cambio (mínimo) |
|---|---|---|
| WinForms `Form1` | `cmbAnio/cmbMes/cmbQuincena` + `PeriodoSeleccionado` (`{anio}{mes}{quincena}`) + `txtPlantilla` (file picker) + guard salida==plantilla existen | El usuario selecciona Año/Mes/Quincena + el OFICIAL en `txtPlantilla`; el run ejecuta Phase 0 con `Periodo.Parse(PeriodoSeleccionado)`. Sin controles nuevos salvo etiqueta de ayuda ("Seleccione la plantilla oficial; la base se prepara sola"). `lblSubtitulo` ya muestra período+quincena — intacto |
| CLI | `--periodo/--quincena/--plantilla/--carpeta/--salida` existen; `--plantilla` = "plantilla de origen (se copia; jamás se modifica)" | `--plantilla` acepta el OFICIAL directamente (cero cambio de flag, solo docs + ayuda); `--periodo` gobierna hojas/tokens/L. Texto de `Uso()` actualizado ("puede ser la plantilla oficial; la base del período se prepara automáticamente") |
| `Docs/BasesPreparadas/` | 3 bases existen; eran input manual | Q1 declara destino: (i) conservar como artefactos de debug versionados con README que diga "NO son input de producción", o (ii) retiro por commit del Ingeniero. Ningún código las referencia tras este plan |

### 2.5 Trade-offs (por qué NO las alternativas)

| Eje | Elegida | Descartada + por qué |
|---|---|---|
| Dónde vive la preparación | Phase 0 en-pipeline (mover código probado) | Mantener paso manual offline: perpetúa el defecto "Agosto 11/116 tokens stale" y la divergencia base-por-costumbre; el mandato lo cierra |
| Cómo elegir M vs N en L | Reescritura mecánica en Phase 0 por dominio (`NumeroQuincena` → variante T0) | **Fórmula runtime `SI()` en el template** (p. ej. `=SI(...;M26;N26)`): contamina el contrato dorado con lógica de quincena (el oficial dejaría de ser el oráculo byte-idéntico), rompe `<f>`-idéntica-a-variante, exige tocar `<f>` en CADA corrida de negocio (viola D-D), y Excel evaluaría la rama contraria como caché stale confuso. La quincena es dominio, no fórmula |
| Dónde enganchar Phase 0 | En el writer (dueño de las 3 rutas de guardado + atomicidad + NoInPlace) | En `ProcesadorPeriodo`: dejaría `EscribirEspejoR1` sin Phase 0 (dos caminos divergentes); el procesador queda intacto y su preflight actúa como red ciega |
| BCL vs OpenXML en Phase 0 | BCL `ZipArchive`+`Regex` (código movido tal cual) | Reimplementar con OpenXML SDK: reescribe lo probado (viola D-C), introduce riesgo de re-serialización (namespaces, shared strings) y acopla la fase mecánica al SDK de negocio |
| `BasesPreparadas/` | Declarar destino en Q1 (debug versionado o retiro) | Seguir aceptándolas como input alternativo: dos fuentes de base = divergencia garantizada; el mandato dice input ÚNICO oficial |

### 2.6 Contratos afectados (resumen)

| Contrato | Cambio |
|---|---|
| Fase nueva `PreparadorBasePipeline` (Infrastructure/Excel, BCL puro) | Código MOVIDO del preparador + paso 4d (columna-L por quincena tabulado desde T0) + autoverificación extendida; parametrizada por `Periodo`; cero literales de período/empresa (D-H) |
| `Herramientas/PreparadorBasePeriodo` | Pasa a wrapper delgado de la fase (mismo CLI, misma salida); cero lógica propia |
| `OpenXmlPlantillaWriter` (solo wiring) | Invoca Phase 0 entre `File.Copy` y `ValidadorBasePeriodo.Exigir` en las 3 rutas; `ValidarNoInPlace` endurecido (assert en Phase 0 + tests origen≠trabajo); misma severidad/mensajes; procesamiento intacto |
| `ValidadorBasePeriodo` | Extendido: acepta la variante L de la quincena del run (texto-`<f>` T0), rechaza tercera forma con `ERR-PLANTILLA` (hoja+celda+variante esperada); resto intacto |
| `WorkbookLeafCellMap*` + `InterventoriaDeclarada` | Solo lo que el sweep §2.3 demuestre stale (direcciones/fragmentos, cero lógica); `InterventoriaDeclarada` como oráculo de cachés por variante |
| `NombresHojaPeriodo` / `Periodo` / slot writer / sellos / calcChain / preflight insumos | INTACTOS (Phase 0 los consume, no los cambia) |
| `Form1` / `Remuneracion.Cli` | Solo cableado + textos de ayuda; sin controles/flags nuevos |
| Fixture T0 (nuevo, tests, cero producción) | Ambas variantes L (texto-`<f>` + cachés por variante) + tabla del sweep con veredicto por hallazgo + hash/conteos de drift |
| `project-context.md` + manual | Doctrina Phase 0 (D-A..D-G) + R-INTERV-Q + tabla del sweep + procedimiento "corre desde el oficial" + destino de `BasesPreparadas/` |

### 2.7 Architecture Validation Certificate

| Principio | Veredicto |
|---|---|
| **SRP** | ✅ Una razón de cambio por costura: T0 declara variantes+veredictos (datos), Phase 0 prepara mecánicamente, R-INTERV-Q decide por dominio, cada gate contrasta una regla, UX solo cablea. Ninguna rama por período ni por empresa. |
| **OCP** | ✅ Un período nuevo compone sin código nuevo (mismo helper de naming + misma Phase 0 parametrizada); empresa no tabulada = identidad; quincena futura imposible por construcción (`Periodo.Parse` solo 1/2). |
| **DIP** | ✅ Phase 0 BCL pura (sin OpenXML/ExcelDataReader); depende de `Periodo`+`NombresHojaPeriodo`+fixture T0; tests con insumos reales + BCL ya presente; cero dependencias nuevas (OpenXML 3.5.1 ya referenciado). |
| **Best practices** | ✅ TDD con puertas (evidencia → fase → regla → cableado → regresión); fail-fast que nombra hoja+celda+variante; quincena por dominio; procesamiento values-only intacto; tolerancia ±0.5; sin Excel/COM en pipeline ni tests; Visual Design Intent N/A (cero rediseño UI). |
| **Performance** | ✅ Phase 0 O(partes del zip) una vez por corrida; drift O(`<f>` de la hoja) lineal; cero costo en procesamiento (la fase sustituye el paso manual, no agrega pases de negocio). |

---

## 3. SPEC

### 3.1 Requirements

| ID | Requisito | Grupo | Criterio de aceptación |
|---|---|---|---|
| R-T-1 | T0 congela desde bytes AMBAS variantes L9:L13 (Q1→N / Q2→M: texto-`<f>` exacto + cachés por variante con citas) + tabla del sweep §2.3 (cada hallazgo = intencional-por-quincena con regla, o invariante) + N3/consumidores | T0 | Fixture que manda; Q1 respondida; revisión del Ingeniero; `git diff` de producción vacío para T0 |
| R-T-2 | Allow-list explícita de `<f>` tocables SOLO por Phase 0 (11 tokens de período + 5 L9:L13); cualquier otra `<f>` tocada por cualquier fase = FAIL | T0 | Lista citada en el fixture; grep de cierre la respeta por namespace/clase |
| R-B-1 | Phase 0 automática: oficial→trabajo (oficial nunca modificado) + 3 hojas + 11 tokens + N3 + columna-L por quincena + app.xml/defined-names + autoverificación + escritura atómica; código MOVIDO (paridad byte-a-byte con la herramienta salvo columna-L) | B | Misma entrada → mismos bytes (herramienta vs fase); `ValidarNoInPlace` verde en las 3 rutas; build 0 warnings |
| R-B-2 | R-INTERV-Q: Phase 0 decide SOLO por `Periodo.NumeroQuincena`; drift acepta la variante del run; tercera forma = `ERR-PLANTILLA` (hoja+celda+variante esperada) | B | Tests por quincena con el OFICIAL como entrada: Q1→N, Q2→M, forma corrupta→FAIL con los 3 campos |
| R-B-3 | `ValidadorBasePeriodo` extendido corre sobre la copia auto-preparada y está verde en ambos períodos (red ciega) | B | Regens desde el oficial: preflight verde Q1 y Q2 sin intervención manual |
| R-B-4 | D-H + catálogo abierto en el diff | B | Grep: 0 literales/ramas por período; 0 comparaciones por empresa fuera de filas citadas; Phase 0 recibe `Periodo`, nunca strings de período/hoja |
| R-B-5 | Cableado UX sin rediseño: WinForms (campos existentes) + CLI (`--plantilla`=oficial, `--periodo` decide) | B | Corrida Q1 y Q2 desde ambos frontends con el oficial como `--plantilla`/`txtPlantilla`; la quincena del run = la seleccionada (nunca la del contenido) |
| R-G-1 | `<f>` byte-idénticas a la variante esperada workbook-wide (texto + `t`/`ref`/`si`) + `<v>` solo en destinos del contrato | G | FAIL ante cualquier `<f>` divergente de la variante o `<v>` fuera de contrato; prohibido `<v>` como oráculo |
| R-G-2 | DetRetri-D 5/5 vs R10 ±0.5 AMBOS períodos + shared/calcChain ciegos + healthy-open (sin diálogo de reparación) | G | Regens FRESCAS a temp DESDE EL OFICIAL (sin base preparada como input); Excel abre sano (paso manual del Ingeniero como aceptación final) |
| R-R-1 | Suite verde salvo dependencias declaradas + build 0 warnings por tarea; `BasesPreparadas/` con destino declarado (Q1) y cero referencias nuevas en código | R | `dotnet build …slnx` 0 warnings + `dotnet test` por tarea; grep: 0 referencias a `BasesPreparadas` en producción nueva; prohibido `Insumos.Raiz()` y `Docs/Insumos` raíz en lo nuevo |
| R-R-2 | Docs: doctrina Phase 0 + R-INTERV-Q + sweep + manual "corre desde el oficial" + lectura de cada FAIL | R | Diff acotado; sin código en el mismo commit |

### 3.2 Scenarios (Given/When/Then)

- **S1 (oficial directo Q2):** Given el oficial + período Q2 (p. ej. 2026072), When run end-to-end, Then la copia de trabajo trae las 3 hojas Q2 + 11 tokens Q2 + N3=2026072 + L9:L13=M-variante; DetRetri-D 5/5 vs R10 ±0.5; healthy-open.
- **S2 (oficial directo Q1):** Given el oficial + período Q1 (p. ej. 2026071), When run end-to-end, Then lo mismo con hojas/tokens/N3 Q1 + L9:L13=N-variante; DetRetri-D vs R10 ±0.5 (oráculo de validación según G3); healthy-open.
- **S3 (tercera forma):** Given una copia cuya L9:L13 no es ni M- ni N-variante, When preflight/drift, Then `ERR-PLANTILLA` nombrando hoja+celda+variante esperada; cero escritura parcial.
- **S4 (in-place denegado):** Given salida == ruta del oficial, When cualquier camino de guardado o Phase 0, Then denegación `ERR-PLANTILLA` antes de tocar un byte; el oficial queda intacto (hash idéntico).
- **S5 (quincena por dominio):** Given selección Q1 con workbook base-interna Q2 (o viceversa), When run, Then hojas/tokens/L/N3 = los de la SELECCIÓN (el contenido previo se sobrescribe mecánicamente, nunca se hereda).
- **S6 (paridad mover-no-reescribir):** Given misma entrada (oficial + período), When herramienta CLI vs Phase 0 en-pipeline (fijando columna-L igual), Then bytes idénticos (salvo metadatos ZIP esperados: timestamps — lista explícita).
- **S7 (sweep cerrado):** Given el oficial, When barrido §2.3, Then cada hallazgo tiene veredicto citado (intencional-con-regla o invariante-con-motivo); ningún "probablemente igual".
- **S8 (regresión):** Given ambos períodos desde el OFICIAL a temp FRESCA, When regens, Then healthy-open + `<f>`-idéntica-a-variante + DetRetri-D 5/5 vs R10 ±0.5 + shared/calcChain verdes + preflight verde.

---

## 4. TASKS

| ID | Tarea | Grupo | Depende de | Verificación |
|---|---|---|---|---|
| T0 | Evidencia congelada (cero código). Leer el oficial (BCL ZIP+XML): ambas variantes L9:L13 (texto-`<f>` exacto + cachés por variante con citas celda-a-celda) + sweep §2.3 (S-MN/S-VAL/S-J/S-DV/S-AJ/S-R26/S-C59/S-N3, cada hallazgo con veredicto) + allow-list `<f>`-tocables-por-Phase-0 + hash/conteos de drift. Con Q1 respondida y revisión del Ingeniero. | T0 | Q1 | R-T-1/R-T-2 (S7 parcial-fixture); `git diff --stat` vacío en producción; el fixture manda |
| T1 | Phase 0 — mover el preparador (TDD). Trasladar el cuerpo mecánico a `Infrastructure` parametrizado por `Periodo` + paso 4d (columna-L por quincena tabulado desde T0) + autoverificación extendida (variante L) + herramienta como wrapper delgado. Tests: paridad byte-a-byte herramienta-vs-fase (S6) + Q1→N / Q2→M / forma-corrupta→FAIL (S3 parcial) con el oficial por ruta explícita a temp FRESCA. | B+G | T0 | R-B-1/R-B-2/R-B-4 (S3/S6 parcial); build 0 warnings; 0 literales de período en el diff |
| T2 | Wiring + R-INTERV-Q + preflight extendido. Invocar Phase 0 en las 3 rutas del writer (entre `File.Copy` y validaciones) + `ValidarNoInPlace` endurecido (S4) + `ValidadorBasePeriodo` acepta-variante-del-run / rechaza-tercera-forma + drift gate por variante. Tests: S3/S4/S5 por quincena desde el oficial. | B+G | T1 | R-B-2/R-B-3 (S3/S4/S5); build 0 warnings; grep: 0 escrituras `<f>` fuera de Phase 0, 0 insert/delete en procesamiento |
| T3 | Sweep a veredictos congelados. Cada hallazgo §2.3 con regla o invariante: ajustes mínimos de direcciones/fragmentos donde T0 demuestre stale (cero lógica) + drift gate por familia. Si un hallazgo exige diseño propio → follow-up con T0 propio (no bloquea T4 salvo camino L/N3/hojas). | B | T2 | R-T-1 (S7 completo); grep 0 direcciones legacy en el camino tocado |
| T4 | Cableado UX. WinForms: oficial en `txtPlantilla` + período seleccionado → Phase 0 (textos de ayuda, sin controles nuevos) + CLI: `Uso()` actualizado (`--plantilla` = oficial directo). Tests: S1/S2/S5 desde ambos frontends (in-process donde aplique). + Destino declarado de `Docs/BasesPreparadas/` según Q1 (README-debug o retiro fuera del plan; cero referencias nuevas en código). | B | T2 | R-B-5/R-R-1 (S1/S2/S5 parcial); build 0 warnings |
| T5 | Regresión global desde el OFICIAL. Regens FRESCAS a temp (Q1+Q2 × 5 ASE, sin base preparada como input): healthy-open + `<f>`-idéntica-a-variante (S1/S2) + DetRetri-D vs R10 ±0.5 + shared/calcChain ciegos + preflight verde + comparador sin exenciones nuevas. Solo rutas explícitas. | R | T3, T4 | R-G-1/R-G-2/R-R-1 (S8); suite verde salvo dependencias declaradas |
| T6 | Docs + cierre. `project-context.md` (doctrina Phase 0 D-A..D-G + R-INTERV-Q + tabla del sweep + regla de re-congelamiento) + manual ("corre desde el oficial" + lectura de cada FAIL S3/S4/S5) + build 0 warnings + greps globales (D-H, catálogo-abierto, sin-`<f>`-fuera-de-Phase-0, sin-insert/delete-procesamiento, 0 refs `BasesPreparadas` en producción nueva, oficial intacto) + CRLF + sin commits. | D+Transversal | T0..T5 | R-R-2; DoD §5 |

**Orden sugerido (puertas duras):** Q1 → T0 (evidencia, sin código) → T1 (mover fase + paridad) → T2 (wiring + R-INTERV-Q + preflight, prohibido sin T1 verde) → T3 (sweep a veredictos) + T4 (UX, en paralelo tras T2) → T5 (regresión global desde el oficial) → T6. Build 0 warnings antes de marcar cada tarea completa. Sin commits (los hace el Ingeniero con `#commit`/`#push`).

---

## 5. Resumen ejecutivo, riesgos y decisiones

- **Resumen:** por mandato 2026-10-14 la plantilla oficial pasa a ser el input directo para todos los períodos: una Phase 0 automática (código MOVIDO del preparador probado, no reescrito) copia el oficial a la ruta de trabajo y prepara la base —3 hojas por `NombresHojaPeriodo`, 11 tokens `<f>`, N3 sellado, Y columna L9:L13 por quincena (R-INTERV-Q: Q1→N, Q2→M)— antes del preflight y de cualquier escritura de negocio. El invariante se re-encuadra con precisión quirúrgica (base-prep mecánica SÍ toca `<f>` en fase explícita auditada; procesamiento NUNCA), el oficial nunca se modifica in place (guard endurecido), y un sweep BCL cierra sin supuestos silenciosos toda otra ref dura de quincena. La regresión corre ambos períodos end-to-end DESDE EL OFICIAL a temp fresca.
- **Riesgo principal:** R-ATRIBUCION + R-MOVER (columna de quincena cruzada mueve dinero real por el off-by-one ASE1/2/3/5; mover código puede divergir en silencio). Contenido por quíntuple red: T0 congela ambas variantes con cachés citados (T0), paridad byte-a-byte herramienta-vs-fase (T1), decisión SOLO por dominio + drift que acepta-únicamente-la-variante-del-run (T2), preflight verde como red ciega (T2/T5), y DetRetri-D vs R10 ±0.5 por ASE en regresión desde el oficial (T5).
- **Por qué NO fórmula runtime `SI()` ni `SI()` en template:** contaminaría el contrato dorado con lógica de quincena (el oficial dejaría de ser oráculo byte-idéntico), rompería `<f>`-idéntica-a-variante, exigiría tocar `<f>` en cada corrida de negocio (viola D-D), y dejaría cachés stale de la rama contraria. La quincena es dominio (`Periodo.NumeroQuincena`), no fórmula — se resuelve una vez en Phase 0 mecánica, no en cada recálculo de Excel.
- **Por qué NO reescribir el preparador:** tiene exit 0 probado con self-verification para 3 períodos; reescribirlo reintroduce el riesgo que ya está cerrado. MOVER + paridad byte-a-byte + wrapper delgado = una sola implementación viva.
- **Decisión para el Ingeniero:** ratificar D-A..D-G (§0.4) + responder Q1 (§0.2, con el destino de `Docs/BasesPreparadas/`). **Forks con recomendación:** hook en writer vs procesador (recomendado writer: dueño de las 3 rutas + atomicidad + NoInPlace); `BasesPreparadas/` como debug-versionado vs retiro (recomendado: lo que diga Q1, sin código que las referencie en ningún caso); alcance OUT (§0.3, recomendado: diff mínimo). **Sin preguntas técnicas bloqueantes pendientes; gate de entrada = Q1.**

### DoD (Definition of Done)

1. `dotnet build Remuneracion.WinForms/Remuneracion.WinForms.slnx` con 0 warnings.
2. Puertas en orden: T0 congelado del disco (Q1 + revisión) → T1 fase movida con paridad byte-a-byte → T2 wiring + R-INTERV-Q + preflight extendido (S3/S4/S5) → T3 sweep con veredicto por hallazgo → T4 UX cableado → T5 regresión Q1+Q2 desde el OFICIAL (healthy-open + `<f>`-idéntica-a-variante + DetRetri-D 5/5 vs R10 ±0.5 + shared/calcChain verdes).
3. Greps de cierre: 0 literales/ramas por período (D-H); 0 ramas por empresa fuera de filas citadas; 0 escrituras `<f>` fuera de Phase 0 (allow-list T0); 0 insert/delete en procesamiento; 0 referencias a `BasesPreparadas` en producción nueva; oficial intacto (hash); ningún `<v>` como oráculo.
4. `project-context.md` + manual actualizados (doctrina Phase 0 + R-INTERV-Q + sweep + "corre desde el oficial"); regens siempre a temp FRESCA con rutas explícitas y el OFICIAL como única entrada; sin commits del agente; CRLF; sin emojis.

### Follow-ups explícitos (fuera de este plan)

- Cambio futuro del oficial por el negocio: su propio T0 (procedimiento §T0 + drift gate como detector; re-congelar AMBAS variantes L).
- Hallazgo del sweep que exija diseño propio (T0 + plan hijo).
- F209/F219 como celdas (su propio T0); motor-espejo R4; CF/DV (W-4); DetRetri-semántica; reconciliación de fixture-root (dependencia declarada; lo nuevo usa rutas explícitas).
- Retiro físico de `Docs/BasesPreparadas/` si Q1 lo pide (commit del Ingeniero, fuera del diff de código).
