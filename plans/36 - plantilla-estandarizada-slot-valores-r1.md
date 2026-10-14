# Plan 36 — Plantilla estandarizada única: escritor de valores por slot R1 (retiro del motor insert/delete + recomposición)

> **Alcance:** cambiar el paradigma de escritura R1: de espejo dimensional + recomposición de `<f>` (Planes 21/31/32/35) a **escritor values-only por slot sobre la plantilla estandarizada única** `Docs/Plantilla estandarizada/Formato Remuneracion 2026xx-x Total.xlsx`. La variabilidad mensual vive en el DATO, nunca en el template. La app NUNCA inserta/borra filas, NUNCA escribe `<f>`; solo pega VALORES (`<v>`) en el slot de datos que corresponde por firma+clasificador+sinonimia; lo no entregado queda en cero como lo trae el template; fuente que trae empresa/componente sin slot → fail-fast (nunca remap silencioso, nunca insert). Cambio de producción esperado: **mínimo y neto-negativo** (se ELIMINAN dimensionado del espejo + reanclaje por delta + `RecomponerVisibles`/`RecomponerInterior` que escriben `<f>`; se AGREGA contrato de slots congelado desde disco + writer values-only + gates). Cero cambios de semántica de negocio fuera de lo declarado.
> **Rector (IRRENUNCIABLE):** `README.md` + `Docs/Propuesta_Proyecto_Automatizacion_Remuneracion_UAESP.md` + `Docs/Detalle de plantilla.docx` + `Docs/Proceso de Recaudo.docx` + `Docs/Prompt_Maestro_Proyecto_Remuneracion_UAESP _ Vo .docx` + `.opencode/project-context.md`. **Invariantes:** tolerancia ±0.5 en goldens; NUNCA usar cachés `<v>` como oráculo (aquí `<f>` debe quedar byte-idéntica al template — se aserta); doctrina calcChain Plan 28 intacta (delete calcChain + fullCalcOnLoad); build 0 warnings antes de marcar cualquier tarea completa; CRLF sin mezclar; NUNCA git commit/push (los hace el Ingeniero con `#commit`/`#push`); la implementación la ejecuta `implementer` usando SOLO este plan aprobado; `code-reviewer` verifica spec compliance + 0 warnings. **Quincena por dominio, nunca por contenido.**
> **Mandato (aprobado por el Ingeniero, Engram decision #2067 topic `plantilla-estandarizada`):** plantilla oficial única para TODOS los períodos; slot doctrine; labels nuevos en columna C son el ESTÁNDAR y se respetan tal cual; ceros donde Julio/Agosto traían fórmulas reales son defaults intencionales de negocio — la app NO restaura fórmulas viejas; se reinvierten como inputs de colocación firma→slot, `SinonimosEmpresaR1`, clasificador de frontera `EsAplicacionPorFrontera` (Plan 35-T2 aprobado); se RETIRA el engine insert/delete/resize + recomposición de fórmulas R1 (dimensionamiento del espejo, Reanclar por delta, `RecomponerVisibles`/`RecomponerInterior` que escriben `<f>` de Planes 21-dinámica/31/32); se CONSERVAN signature reading, classifier, synonymy, calcChain doctrine, shared-preservation donde aplique, D-H cero-literales-de-período en diff de producción, open-company-catalog (nunca ramas `if ENEL/OCCIDENTE`), sellos G7/K7/D6/D7 vía `NombresHojaPeriodo`.
> **Continuidad:** Planes 21 (espejo), 25 (roles por firma), 26/27 (preflight + firma + hoja-por-nombre), 28 (calcChain + sello), 29 (paridad + comparador f-vs-f), 30 (naming dinámico + base canónica), 31 (visibles), 32 (interior + catálogo C ABIERTO + quirk `--L`), 33 (shared-formula + T0 agosto H1), 34 (sinonimia como DATOS + D-H), 35 (motor dinámico: T2-clasificador APROBADO y vigente) vigentes en lo no retirado. Este plan NO reabre lectura R2/R4-detalle, cálculo CONSOLIDADO, preflight, sellos, naming, calcChain, paridad fuera de lo declarado.
> **Insumos canónicos:** plantilla única `Docs/Plantilla estandarizada/Formato Remuneracion 2026xx-x Total.xlsx` (base de slots; período base interno 2026092 aunque el filename diga `2026xx-x`); bases julio/agosto y oráculos manuales solo como referencia de regresión (julio-2 abre sano aceptado; agosto-2 Resultado5 con 495 divergencias `<f>`-texto en 20 filas R1 agrupadas C1..C4 — ver §0.1); fuentes R1 por ASE; `R1FirmaInterior.cs` (+`EsAplicacionPorFrontera`), `CompositorInteriorR1.cs` (solo como referencia de firma, no como escritor), `SinonimosEmpresaR1.cs`, `NombresHojaPeriodo.cs`, `ValidadorTotalesR1Workbook.cs`, `ComparadorSalidaVsManual.cs`, `WorkbookLeafCellMapQ2.cs`, `Herramientas/PreparadorBasePeriodo/Program.cs`. Regens SIEMPRE a rutas temp FRESCAS.
> **Numeración:** `plans/` 01..35 ocupados; este plan toma el primer correlativo libre, **36**.
> **Estado:** EN EJECUCIÓN — T0+T1 ACEPTADOS por code-reviewer 2026-10-13 (build 0 warnings, 22/22 tests verdes; Q1..Q4 negocio siguen PENDING). Puerta a T2 abierta.
> **Fecha:** 2026-10-13

---

## 0. Clarification Gate

**Sin preguntas bloqueantes de diseño: la evidencia estructural viene ADOPTADA por encargo (no se re-diagnostica) y las restricciones vienen cerradas (slot doctrine + D-H + catálogo abierto + Plan 34 en vigor).** El único gate de entrada es de NEGOCIO (§0.2: cuestionario de 4 puntos), no técnico. Punto de bloqueo respetado: NO se implementa nada en este documento (planificación read-only).

### 0.1 Evidencia adoptada (no re-diagnosticar; se congela como contrato)

| # | Afirmación adoptada | Reutilización en este plan | Re-verificación (el test, no la mano) |
|---|---|---|---|
| E1 | Template único, 40 hojas mismo orden; base período 2026092 (`Informe AFaseo Recaudo 202609-2`, `DetRetri2026092`, `DetValiRetri2026092`) aunque filename diga `2026xx-x`; bloques por ASE +22..+32 filas con 111 filas `hidden="1"` (213-251, 356-412, 552-566) como espacio reservado. | El mapa de slots se congela desde estos bytes (T0): cada slot = (ASE × empresa × tipo-fila → celda fija). Las filas hidden son RESERVA: el writer values-only NUNCA las toca ni las desoculta. | T0: fixture del slot map leído del disco; T-gates: grep + asserts de que ninguna fila hidden cambia `hidden` ni recibe `<v>`. |
| E2 | Slots `NUEVO ESQUEMA` 5→8 (C135/C160/C182/C200 ASE2, C475/C500/C522/C540 ASE4); labels de empresa de sub-área movidos A→C (Julio A8=ENEL/A11=OCCIDENTE vs NEW C8/C15 + slots C21/C31/C36/C46/C51/C58/C62/C69). | Los labels NUEVOS en C son el estándar: el descubrimiento de slots matchea por firma C (no se normaliza a layout julio). | T0: asserts C8/C15 + slots citados; T3: texto-vs-template de labels intacto (el writer no escribe labels). |
| E3 | Interiores mantienen forma Sigma-minus-last (ASE1 F85: F62+F36+F8-L8-L36; quirk `--L` intacto F264); shared masters 35→37 (`si=0..36`, nuevos si=22 F445:Z445, si=31 G598:Z598, si=36 G710:AA710, si=6 F105:AP105; seguidoras 1035→1101); celdas L-1 ahora literal `<v>0</v>` PERO perdieron fórmulas reales de julio (F209=F142, F217=F145, F219=F136 → 0 en NEW); `CONSOLIDADO!D111=D109+N109` solo-en-NEW; `DetRetri2026092!D9:D14` editable-cero; sellos G7=16/09/2026 K7=30/09/2026 D6=02/10/2026 D7≈10:43AM N3=2026092; defectos: typo `RECIPRCIDAD` en `ANTICIPOS USUARIOS!E1`, placeholder `OJO/// DILIGENCIAR` en `REPORTE RECAUDO x BANCO!D59` (modificado 2026-10-13T16:41Z, retocado a mano); template Agosto solo 11/116 tokens de período actualizados (105 stale 2026072) — NO es base preparada. | Tres consecuencias normativas: (a) ceros L-1 son defaults intencionales → allow-list de ceros que el writer NUNCA "completa" con fórmulas viejas; (b) grupos mirror G:AP shared = regla NO-TOUCH total (ni `<v>` ni `<f>` en esas columnas por este plan); (c) base canónica de cada período sale de la plantilla ÚNICA vía procedimiento §2.3, nunca de Agosto-stale ni de salidas trabajadas. | T0: allow-list L-1 congelada desde disco + lista de celdas mirror NO-TOUCH; T-gates: `<f>` byte-idéntica al template en todo R1 + `<v>` solo en slots de datos. |
| E4 | Runtime previo: Julio-2 abre sano (aceptado); Agosto-2 Resultado5 abre OK con 495 divergencias `<f>`-texto en 20 filas R1: C1 ancla-0 AFASEO (5 ASE F0: fuente A=`AFaseo` B=`Total` vs predicado D=`AFASEO` E=`Total`), C2 L-1, C3 F574 orden, C4 grupos mirror H..AP stale con diffs reales de valor. | Bajo slot doctrine: C2/C3 se DISUELVEN (ceros respetados, orden irrelevante si no se toca); C1 (descubrimiento de slot AFaseo cubriendo AMBAS variantes de firma) y C4 (NO tocar mirror groups — el writing values-only lo fija por construcción) entran como ítems del mapa de slots, no como bugs de recomposición. | T0: AFaseo dual-firma en el contrato; T-gates: oráculo `<f>`-idéntica + DetRetri-D vs R10 ±0.5. |
| E5 | Arsenal vigente a REINVERTIR: firma→slot mapping (predicados `R1FirmaInterior`), tabla `SinonimosEmpresaR1` (`{RECIPROCIDAD, NUEVO ESQUEMA}` como filas de DATO citadas), clasificador de frontera `EsAplicacionPorFrontera` (Plan 35-T2 aprobado). A RETIRAR para R1: dimensionamiento del espejo, Reanclar por delta, `RecomponerVisibles`/`RecomponerInterior` que escriben `<f>` (planes 21-dinámica/31/32 partes de escritura). A CONSERVAR: signature reading, classifier, synonymy, calcChain doctrine (Plan 28), shared-preservation donde aplique, D-H, open-company-catalog, sellos vía `NombresHojaPeriodo`. Anclas: `OpenXmlEspejoR1Mutador.cs` (dimensionado, Reanclar, RecomponerVisibles/Interior, `CeldasInterioresAgosto` ~L854-897, `EsPeriodoJulio` ~L906-930, `EscribirFormulaPreservando` ~L1578), `OpenXmlPlantillaWriter.cs` (save paths + `SaneadorCadenaCalculo` + shared-follower awareness), `Herramientas/PreparadorBasePeriodo/Program.cs` (11 refs token `<f>`; debe manejar placeholder `2026xx-x` + base 2026092). | Delimitación exacta del diff: el retiro es de ESCRITURA geométrica y de `<f>`; la LECTURA por firma y la clasificación sobreviven como resolutores de slot. | T-gates: grep de cierre (0 `InsertRow`/`DeleteRow`/dimensionado R1, 0 `new CellFormula` en path R1, 0 ramas por empresa/período). |

### 0.2 Gate de entrada — cuestionario de negocio (BLOQUEANTE para T1..Tn)

La primera corrida de la plantilla única es SUPERVISADA. Antes de congelar el contrato T0 el Ingeniero responde (una línea cada una basta):

| ID | Pregunta | Efecto en el plan si la respuesta es "sí/no" |
|---|---|---|
| Q1 | Typo `RECIPRCIDAD` en `ANTICIPOS USUARIOS!E1`: ¿se deja tal cual (estándar) o se corrige en la plantilla única? | Si se corrige: la corrección es cambio A LA PLANTILLA (fuera del diff de código), se re-congela T0 y se versiona el byte nuevo. El código nunca la "corrige" en runtime. |
| Q2 | Placeholder `OJO/// DILIGENCIAR` en `REPORTE RECAUDO x BANCO!D59`: ¿dato pendiente del administrativo o celda que la app debe poblar? | Si es de la app: entra como slot con fuente citada (su propio mini-T0); si no: queda fuera del contrato y el gate aserta que D59 no se toca. |
| Q3 | Ceros ex-fórmula (F209/F217/F219 y clase L-1): ¿confirmado que son defaults intencionales que la app NO debe restaurar con fórmulas viejas? | Confirmación = allow-list T0 queda sellada; cualquier "completar fórmula vieja" futuro exige su propio T0 + aprobación. |
| Q4 | Filas hidden de reserva (213-251, 356-412, 552-566): ¿confirmado que son espacio reservado que la app NUNCA usa (ni desoculta, ni dimensiona)? | Confirmación = regla NO-TOUCH con gate; si un mes futuro las necesitara, su propio T0 (nunca uso silencioso). |

### 0.3 Mapeo al Rector (Sale / No sale)

**Sale (EXPLÍCITO):**
- Re-diagnosticar la evidencia E1..E5 o reabrir el debate "template por mes vs template único" (mandato #2067 lo cierra).
- Cualquier `InsertRow`/`DeleteRow`/resize/dimensionado por delta en R1; cualquier escritura de `<f>` en R1 (`new CellFormula`, `.Text` de fórmula, Reanclar de fórmulas por delta); restaurar fórmulas viejas donde el template trae 0 intencional.
- Remap silencioso (fuente sin slot → colocar en "el parecido"), composición parcial, o 0 silencioso donde el contrato exige fail-fast con ASE+reporte+celda+slot.
- Tocar columnas mirror G:AP shared (ni `<v>` ni `<f>`), filas hidden de reserva, labels (A o C), sellos fuera de G7/K7/D6/D7, tokens de período dentro de `<f>` en runtime.
- Ramas por nombre de empresa (`if ENEL/OCCIDENTE/RECIPROCIDAD/…`) fuera de las FILAS citadas de `SinonimosEmpresaR1`; literales/ramas por período en el diff (D-H); cachés `<v>` como oráculo; paquetes NuGet; commits/push.
- Renombrar rótulos del template en runtime (el template manda en rótulos; Q1 es cambio de plantilla, no de código).

**No sale (entra, alcance congelado):**
| Frente | Superficie de cambio |
|---|---|
| T0 — contrato de slots | Mapa congelado desde disco (ASE × empresa × tipo-fila → celda fija) + allow-list L-1 + AFaseo dual-firma + NO-TOUCH mirror + techos de capacidad con códigos fail-fast; cero código de producción (fixture + doc) |
| B1 — writer values-only | Nuevo path que escribe SOLO `<v>` en slots resueltos por firma+clasificador+sinonimia, con assert pre-write "celda destino NO tiene `<f>`" (guard anti-fórmula como REGLA, no excepción); cero insert/delete |
| B2 — retiro | DELETE dimensionado espejo R1 + Reanclar por delta R1 + `RecomponerVisibles`/`RecomponerInterior` que escriben `<f>` (+ `CeldasInterioresAgosto` + `EsPeriodoJulio` si aún viven tras Plan 35); conservar lectura por firma, `EsAplicacionPorFrontera`, `SinonimosEmpresaR1`, `EscribirFormulaPreservando` donde aplique, saneador calcChain, sellos |
| B3 — base de período | Decisión PreparadorBasePeriodo-extendido vs naming-solo-runtime (§2.3) + sellos G7/K7/D6/D7 vía `NombresHojaPeriodo`; runtime NUNCA reescribe tokens `<f>` de período |
| G — gates | `<f>`-idéntica workbook-wide + `<v>`-solo-en-slots + invariante shared + calcChain estructural + DetRetri-D vs R10 ±0.5 + healthy-open; fixtures por ruta explícita, prohibido `Docs/Insumos` raíz y `<v>` como oráculo de fórmulas |
| D — docs | `project-context.md` (doctrina slot) + manual (procedimiento mes-nuevo supervisado + cómo leer cada FAIL) |

**Fuera de scope explícito (OUT):** motor-espejo R4; cobertura `conditionalFormatting`/`dataValidations` (W-4); cambios de semántica DetRetri (redondeo/columna D intactos); plumbing UI/CLI (CLI solo como runner de repro); reconciliación de fixture-root de los 511 rojos pre-existentes (dependencia declarada, no se toca; lo nuevo usa rutas explícitas); completar F209/F219 como follow-up separado salvo que el slot map T0 las subsuma (si el template trae slot real para ellas, entran; si son L-1 genuino-cero, quedan en allow-list, no se "completan").

### 0.4 Decisiones ejecutivas (aprobación = aceptar estas)

| ID | Decisión |
|---|---|
| D-A | **Slot doctrine como invariante de escritura.** La app solo escribe `<v>` en celdas-dato del contrato T0; pre-write assert "sin `<f>` en destino"; sin slot → fail-fast `ERR-SIN-SLOT` nombrando ASE+reporte+celda+slot esperado (nunca remap/insert). |
| D-B | **Un solo template; la variabilidad mensual es DATO.** Ningún cambio geométrico ni de `<f>` por mes en código ni en runtime. Q1..Q4 (§0.2) es el único gate de entrada. |
| D-C | **Reinversión acotada del arsenal.** Firma (predicados `R1FirmaInterior`), `SinonimosEmpresaR1` (filas citadas, catálogo abierto), `EsAplicacionPorFrontera` (Plan 35-T2) sobreviven como RESOLUTORES de slot. El engine de escritura geométrica + recomposición `<f>` se retira (net-negative). |
| D-D | **Ceros del template se respetan; fórmulas viejas no se restauran.** Allow-list L-1 congelada desde disco; el gate falla si aparece `<f>` donde el template trae literal-0 intencional (es regresión, no mejora). |
| D-E | **Mirror groups G:AP y filas hidden = NO-TOUCH.** Ni lectura para escribir, ni `<v>` ni `<f>`; el gate los audita byte-idénticos (fórmulas + valores + atributos `t`/`ref`/`si` + `hidden`). C1/C4 de E4 se cierran por esta regla + AFaseo dual-firma. |
| D-F | **D-H + catálogo abierto siguen irrenunciables.** Cero literales/ramas por período y cero ramas por empresa en el diff; el motor no recibe ningún período; empresas no tabuladas = identidad. |
| D-G | **Base por procedimiento, no por código runtime.** El runtime NUNCA reescribe tokens `<f>` de período; la base canónica de cada período sale del procedimiento §2.3 (decisión T3 con trade-offs). Sellos G7/K7 (serial OADate, Plan 28) + D6/D7 (texto, Plan 31) vía `NombresHojaPeriodo`, con guard anti-fórmula. |

---

## 1. PROPOSE

### 1.1 Intent

Que la remuneración de CUALQUIER quincena se produzca desde la plantilla estandarizada única pegando solo valores en slots fijos — con apertura sana en Excel, `<f>` byte-idénticas al template, DetRetri-D validado contra R10 ±0.5, y fail-fast accionable cuando la fuente traiga lo que el template no contempla — retirando del código todo dimensionamiento y recomposición de fórmulas R1.

### 1.2 In Scope

- Contrato de slots T0 congelado desde disco + allow-list L-1 + AFaseo dual-firma + NO-TOUCH mirror + techos fail-fast (con Q1..Q4 respondidas).
- Writer values-only B1 + retiro B2 + base B3 con sellos.
- Gates de aceptación (§3) con oráculo `<f>`-idéntica + DetRetri-D vs R10 + shared/calcChain ciegos + healthy-open.
- Docs: doctrina slot + procedimiento de mes nuevo supervisado.

### 1.3 Non-Goals

Ver §0.3 OUT. En particular: NO se toca R4-espejo, CF/DV, DetRetri-semántica, UI/CLI-plumbing, fixture-root; NO se restaura ninguna fórmula vieja; NO se "completa" F209/F219 salvo slot real en T0.

### 1.4 Riesgos

| ID | Riesgo | Mitigación |
|---|---|---|
| R-SLOT | El slot map congelado omite un slot real (empresa×tipo que sí trae el template) y meses futuros fallan en fail-fast legítimo | T0 congela desde disco celda-a-celda con fixture que manda + revisión del Ingeniero; fail-fast nombra slot esperado (el próximo T0 lo arbitra con evidencia, no se adivina) |
| R-ATRIBUCION | Firma+sinonimia colocan valor en el slot de la empresa equivocada (dinero movido) | Catálogo abierto + solo `SonMismaEmpresa` (filas citadas) + asserts texto-slot vs fuente término a término + DetRetri-D vs R10 ±0.5 como red por ASE |
| R-ALL-ASE | El switch de comportamiento golpea los 5 ASE a la vez (bloques +22..+32, NUEVO ESQUEMA 5→8, labels C) | Faseo con puertas por ASE (§4: T0 por ASE, T1 con gate por ASE, T4 regresión 5 ASE × ambos meses); cada ASE tiene su techo fail-fast; ningún ASE avanza en rojo |
| R-AFASEO | Solo una variante de firma AFaseo resuelve y la otra aborta (C1 reaparece) | Contrato dual-firma (A=`AFaseo` B=`Total` AND D=`AFASEO` E=`Total`) con tests puros por variante + gate workbook-wide que cita F0 por ASE |
| R-MIRROR | Un path legacy escribe `<v>` en G:AP o toca `hidden` y rompe shared/mirror (C4 reaparece) | D-E + gate NO-TOUCH (fórmulas+valores+attrs+hidden byte-idénticos) + grep de cierre (0 escrituras fuera de slots) |
| R-CERO | Un cero intencional se "completa" con fórmula vieja (F209/F217/F219) creyendo que es mejora | D-D + allow-list sellada + gate `<f>`-idéntica (cualquier `<f>` donde el template trae 0 = FAIL) |
| R-BASE | Se corre desde Agosto-stale (105 tokens 2026072) o salida trabajada y se heredan sellos/calcChain ajenos | §2.3: procedimiento de base + preflight de base (N3/G7/K7/D6/D7 + conteo de tokens del período) con fail-fast antes de escribir |
| R-JULIO | Al retirar el engine se altera julio (que abría sano) | Puerta de identidad-julio adaptada: julio contra NEW-base = valores en slots + `<f>` idéntica + DetRetri-D vs R10; ante diff se revierte el pase, no el golden |
| R-FIXTURE-ROOT | Gates nuevos heredan los 511 rojos (resolución raíz rota) | Rutas explícitas estilo `SharedFormulaR1GateTests` (plantilla única + `Docs/Prueba Julio-2/…`, `Docs/Prueba Agosto-2/…` solo como fuentes de dato); prohibido `Insumos.Raiz()` y `Docs/Insumos` raíz en lo nuevo |

---

## 2. DESIGN

### 2.1 Enfoque: congelar slots, pegar valores, borrar geometría — con puertas que prohíben avanzar en rojo

**Fase 0 — contrato (T0, cero código de producción).** Lectura ZIP+XML BCL del template único (sin Excel/COM): por cada ASE, lista `(empresa × tipo-fila → celda-dato fija)` donde celda-dato = celda con `<v>` (o literal-0) y SIN `<f>`, incluyendo AFaseo en AMBAS variantes de firma, slots `NUEVO ESQUEMA` 5→8 por ASE, y techos de capacidad por empresa×tipo (nº de filas-dato que el slot acepta; exceso → `ERR-SIN-SLOT`). Se congelan además: allow-list L-1 (celdas literal-0 intencionales que NUNCA reciben `<f>`), lista NO-TOUCH mirror (columnas G:AP shared + filas hidden de reserva con sus `t`/`ref`/`si`/`hidden`), y lista de labels estándar (columna C por sub-área, byte-idénticos). Fixture que manda: bytes del disco, no este documento.

**Fase 1 — writer (T1).** Nuevo path (p. ej. `EscritorValoresSlotR1.Escribir(workbookPart, bloquesEspejoLeídos, contratoT0)`): por cada fila-dato de la fuente (leída por `LeerEspejoR1` intacto), resuelve slot por firma (`R1FirmaInterior` — mismos predicados que hoy, ahora como localizadores) + tipo por frontera (`EsAplicacionPorFrontera`, Plan 35-T2 intacto) + empresa por sinonimia (`SonMismaEmpresa`, Plan 34 intacto, catálogo abierto). Escribe SOLO `<v>` (número redondeado según contrato; texto donde el slot sea texto) con assert pre-write: si la celda destino trae `<f>` → `ERR-PLANTILLA` (hoja+celda+slot). Cero `InsertRow`/`DeleteRow`, cero `new CellFormula`, cero `.Text` de fórmula, cero Reanclar. Lo no entregado no se escribe (queda el 0 del template). Fuente sin slot → `ERR-SIN-SLOT` (ASE+reporte+celda+slot esperado + etiquetas observadas). `CONSOLIDADO!D` y `DetRetri!D` se alimentan del dominio calculado bottom-up como hoy (solo `<v>`, redondeo a entero en `DetRetri`, Plan 29 Unidad D intacta), sin tocar sus `<f>` vecinas.

**Fase 2 — retiro (T2).** DELETE en `OpenXmlEspejoR1Mutador.cs`: dimensionado del espejo (inserta/borra 5→1 + deltas por ASE), `Reanclar` por delta para R1 (se conserva solo si otro consumidor no-R1 lo usa — si no, se elimina; el grep de cierre lo demuestra), `RecomponerVisibles` + `RecomponerInterior` en sus partes que ESCRIBEN `<f>`, `CeldasInterioresAgosto`, `EsPeriodoJulio` (si Plan 35 no los eliminó ya). Se CONSERVAN: `MesOportuno`/`FilasAplic` como clasificadores (lectura, no escritura), `EscribirFormulaPreservando` donde aplique a otros paths (disciplina shared intacta), `SaneadorCadenaCalculo` + `fullCalcOnLoad` (Plan 28, 3 caminos de guardado), sellos, naming, preflight, readers, cálculo. Diff neto-negativo exigido por el DoD.

**Fase 3 — base (T3).** Decisión con trade-offs:

| Eje | Opción A (recomendada): extender `PreparadorBasePeriodo` offline | Opción B: naming-solo en runtime |
|---|---|---|
| Qué hace | La herramienta one-shot (BCL `ZipArchive`+`Regex`, fuera del pipeline) acepta la plantilla única como entrada: resuelve placeholder `2026xx-x` → período canónico (`DetRetri{AAAAMMQ}` etc. vía `NombresHojaPeriodo`), reescribe los tokens `<f>` de período UNA vez offline, actualiza `docProps/app.xml`, sella metadatos base; el runtime solo valida la base (preflight de base) y NUNCA toca `<f>` | El runtime renombra hojas al período en memoria sin tocar `<f>` de período |
| Pro | Mantiene el invariante "runtime NUNCA toca `<f>`" (Planes 28/30); la base queda internamente consistente (11+ tokens `<f>` + `N3` + nombres de hoja del mismo período); reutiliza doctrina Plan 30 probada | Menos piezas (sin herramienta) |
| Contra | Un paso offline más (documentado en el manual) | Los tokens `<f>` con `2026092`/stale quedarían inconsistentes con el período real (el runtime no puede reescribirlos sin violar el invariante); se hereda el defecto "Agosto 11/116" por diseño |
| Veredicto | **A**, salvo que T3 demuestre con bytes que B deja la base consistente (carga de la prueba en B) | — |

Sellos en runtime (ambas opciones): `CONSOLIDADO G7/K7` (serial OADate desde R10, Plan 28) + `DetRetri/DetValiRetri D6/D7` (texto, Plan 31) por `NombresHojaPeriodo`, con guard anti-fórmula; `CONSOLIDADO!D6` NO se duplica (ya sellado en base); `N3` lo fija la base (procedimiento), no el runtime.

**Fase 4 — regresión (T4) + docs (T5) + cierre (T6).** Regens FRESCAS a temp (Julio + Agosto contra NEW-base + fuentes reales por ruta explícita): healthy-open (sin diálogo de reparación), `<f>` byte-idénticas al template en todo R1 (texto + `t`/`ref`/`si`), `<v>` solo en slots del contrato, DetRetri-D 5/5 vs R10 ±0.5 ambos períodos, `SharedFormulaR1Gate` + test calcChain verdes como redes ciegas, comparador sin exenciones nuevas (C2/C3 disueltas por construcción; si una celda exigiera divergencia permanente se versiona con cita, no blanket).

### 2.2 Dónde vive cada pieza y trade-offs

| Eje | Elegida | Descartada |
|---|---|---|
| Resolución de slot | Firma (`R1FirmaInterior`) + frontera (`EsAplicacionPorFrontera`) + sinonimia (`SonMismaEmpresa`, filas citadas) como localizadores puros en Core; el writer solo pega `<v>` | Rama por empresa/período/celda (`if F463…`, `if ASE4…`, `if agosto…`): viola catálogo-abierto + D-H; cada mes nuevo exigiría otra rama |
| Escritura | Values-only con assert pre-write sin-`<f>` (D-A: el guard pasa de excepción a REGLA) | Recomposición de `<f>` por firma/forma (Planes 31/32/35-escritura): reintroduce el engine que el mandato retira; orden de términos (C3/F574) volvería a importar |
| Cardinalidad distinta fuente-vs-slot | Fail-fast `ERR-SIN-SLOT` con techo por empresa×tipo (nunca insert, nunca remap, nunca 0 silencioso donde hay dato sin slot) | Dimensionar el espejo (Plan 21): prohibido por slot doctrine; las filas hidden de reserva NO son capacidad elástica (Q4) |
| AFaseo | Dual-firma en contrato (ambas variantes resuelven al mismo slot F0 por ASE; test puro por variante) | Normalizar a una sola variante: la otra abortaría (C1 reaparece) o se "mapearía" en silencio |
| Mirror G:AP | NO-TOUCH total auditado (D-E) | "Pegar valores también en mirror": rompería shared followers (1035→1101) y reabriría C4 |
| L-1 ceros | Allow-list sellada (D-D) | Restaurar F209/F217/F219 viejas: prohibido (son defaults de negocio; Q3) |
| `CompositorInteriorR1` | Queda como REFERENCIA de firma documentada, no como escritor (su forma canónica sirvió para clasificar; el writer no compone formas) | Mantenerlo como escritor primario o fallback: reintroduce escritura `<f>` por etiqueta |
| Validador numérico | `ValidadorTotalesR1Workbook` adaptado a values-only: evalúa `<f>` del template (intactas) + literales/`<v>` en slots vs dominio ±0.5 (el gate lee el WORKBOOK, lección Planes 31→33) | Gate solo-dominio C#: ciego a colocación (la trampa que este plan cierra) |

### 2.3 Contratos afectados (resumen)

| Contrato | Cambio |
|---|---|
| Contrato T0 (nuevo fixture+doc, cero producción) | Slot map por ASE + allow-list L-1 + AFaseo dual + NO-TOUCH mirror/hidden + techos fail-fast, congelado del disco |
| `EscritorValoresSlotR1` (nuevo, Infrastructure) | Solo `<v>` en slots T0; assert sin-`<f>`; `ERR-SIN-SLOT` con ASE+reporte+celda+slot; cero insert/delete, cero `<f>` |
| `OpenXmlEspejoR1Mutador` (dimensionado, Reanclar-por-delta-R1, `RecomponerVisibles`/`RecomponerInterior`-escritura-`<f>`, `CeldasInterioresAgosto`, `EsPeriodoJulio`) | ELIMINADO (difunto; grep de cierre). `MesOportuno`/`FilasAplic` sobreviven como clasificadores puros si el writer los consume; si no, se mueven a Core junto al clasificador |
| `R1FirmaInterior` + `EsAplicacionPorFrontera` + `SinonimosEmpresaR1` | INTACTOS como resolutores (cero lógica nueva salvo dual-firma AFaseo si el predicado actual no cubre ambas variantes — cambio mínimo con tests por variante) |
| `EscribirFormulaPreservando` / `SaneadorCadenaCalculo` / `fullCalcOnLoad` / sellos / naming / preflight / readers / cálculo / DetRetri-redondeo | INTACTOS (grep de cierre); sellos por `NombresHojaPeriodo` |
| `ValidadorTotalesR1Workbook` | Adaptado a values-only (misma severidad, nuevo objeto comparado: `<f>`-template + `<v>`-slots vs dominio) |
| `ComparadorSalidaVsManual` | Sin exenciones nuevas; C2/C3 disueltas por construcción; C1/C4 como ítems de slot, no de recomposición |
| `Herramientas/PreparadorBasePeriodo` | Extendido a plantilla única (opción A §2.3) + preflight de base en runtime |
| `project-context.md` + manual | Doctrina slot (D-A..D-G) + procedimiento mes-nuevo supervisado + cómo leer cada FAIL |

### 2.4 Architecture Validation Certificate

| Principio | Veredicto |
|---|---|
| **SRP** | ✅ Una razón de cambio por costura: T0 declara slots (datos), el writer pega valores, el retiro borra geometría, cada gate contrasta una regla, el fixture fija el disco. Ninguna rama por período ni por empresa. |
| **OCP** | ✅ Un mes nuevo compone sin código nuevo (slots fijos + dato variable); empresa no tabulada = identidad; lo no entregado = 0 del template por construcción. |
| **DIP** | ✅ Resolutores puros (firma/frontera/sinonimia, sin I/O ni OpenXML en Core); writer depende de contrato T0 + firmas de Core; tests con insumos reales + BCL ya presente; cero dependencias nuevas (OpenXML 3.5.1 ya referenciado). |
| **Best practices** | ✅ TDD con puertas que prohíben avanzar en rojo (contrato → writer → retiro → base → regresión); fail-fast con ASE+reporte+celda+slot; quincena por dominio; `<f>` intactas; tolerancia ±0.5; sin Excel/COM en pipeline ni tests. |
| **Performance** | ✅ Escritura O(slots tocados) sin dimensionado ni reanclaje; gates O(fórmulas de la hoja) lineales; julio/agosto sin costo geométrico. |

---

## 3. SPEC

### 3.1 Requirements

| ID | Requisito | Grupo | Criterio de aceptación |
|---|---|---|---|
| R-T-1 | Contrato T0 congelado del disco único: slot map por ASE (empresa×tipo→celda, incl. AFaseo dual-firma y NUEVO ESQUEMA 5→8), allow-list L-1, labels-C estándar, NO-TOUCH mirror/hidden, techos por empresa×tipo | T0 | Fixture que manda (bytes del disco); Q1..Q4 respondidas; revisión del Ingeniero; `git diff` de producción vacío para T0 |
| R-T-2 | AFaseo dual-firma: A=`AFaseo` B=`Total` AND D=`AFASEO` E=`Total` resuelven al mismo slot F0 por ASE | T0 | Tests puros por variante + gate workbook-wide que cita F0 por ASE (C1 cierra) |
| R-B-1 | Writer values-only: SOLO `<v>` en slots T0; assert pre-write sin-`<f>`; cero insert/delete; lo no entregado intacto (0 del template) | B | Gate `<v>`-solo-en-slots verde; grep: 0 `InsertRow`/`DeleteRow`/resize y 0 `new CellFormula`/`.Text`-fórmula en path R1 |
| R-B-2 | Sin slot → `ERR-SIN-SLOT` fail-fast (ASE+reporte+celda+slot esperado + etiquetas observadas); nunca remap/insert/0-silencioso | B | Repro con empresa/componente sin slot aborta con el código y los 4 campos; ningún valor colocado fuera de slot |
| R-B-3 | Retiro neto-negativo: fuera dimensionado espejo R1 + Reanclar-por-delta R1 + escritura `<f>` R1 + mapas/ramas por período del engine | B | Grep de cierre: 0 ocurrencias en producción R1; diff neto-negativo documentado (`git diff --stat`) |
| R-B-4 | D-H + catálogo abierto en el diff | B | Grep: 0 literales/ramas por período; 0 comparaciones por nombre de empresa fuera de filas citadas; el writer no recibe ningún período |
| R-B-5 | Base por procedimiento §2.3-A + preflight de base + sellos G7/K7/D6/D7 vía `NombresHojaPeriodo`; runtime NUNCA toca `<f>` de período | B | Base del período consistente (tokens + N3 + hojas); preflight aborta base equivocada antes de escribir; sellos verificados contra R10 |
| R-G-1 | `<f>` byte-idénticas al template en todo R1 (texto + `t`/`ref`/`si`); labels, mirror, hidden intactos | G | FAIL si cualquier `<f>`/label/mirror/hidden difiere (C2/C3/C4 cierran por construcción); prohibido `<v>` como oráculo de fórmulas |
| R-G-2 | DetRetri-D 5/5 vs R10 ±0.5 ambos períodos + goldens + `SharedFormulaR1Gate` + calcChain estructural verdes + healthy-open sin reparación | G | Regens FRESCAS a temp (Julio + Agosto vs NEW-base); Excel abre sano (paso manual del Ingeniero como aceptación final) |
| R-R-1 | Regresión 5 ASE × ambos meses por ruta explícita; suite + build 0 warnings por tarea; 511 rojos de fixture-root solo listados (dependencia, no se tocan) | R | `dotnet build …slnx` 0 warnings + `dotnet test` por tarea; prohibido `Insumos.Raiz()` y `Docs/Insumos` raíz en lo nuevo |
| R-R-2 | Docs: doctrina slot + procedimiento mes-nuevo supervisado + lectura de cada FAIL | R | Diff acotado; sin código en el mismo commit |

### 3.2 Scenarios (Given/When/Then)

- **S1 (contrato manda):** Given template único en disco, When T0 lista slots, Then cada slot = celda-dato sin `<f>` con (ASE, empresa-C, tipo) citados; AFaseo aparece en ambas variantes; mirror/hidden en lista NO-TOUCH; L-1 en allow-list.
- **S2 (pegar valores):** Given fuente R1 real + contrato T0, When writer, Then `<v>` solo en slots resueltos por firma+frontera+sinonimia; `<f>` idénticas; lo no entregado = 0 del template.
- **S3 (sin slot):** Given fuente con empresa/componente sin slot, When writer, Then `ERR-SIN-SLOT` con ASE+reporte+celda+slot + etiquetas observadas; nada colocado fuera de slot.
- **S4 (anti-fórmula):** Given slot cuya celda destino trae `<f>`, When writer, Then `ERR-PLANTILLA` (hoja+celda+slot) y cero escritura parcial.
- **S5 (C1 AFaseo):** Given filas AFaseo en variante A=`AFaseo`/B=`Total` y en variante D=`AFASEO`/E=`Total`, When resolución, Then ambas al mismo slot F0 del ASE; F0 con valor esperado.
- **S6 (C4 mirror):** Given salida post-writer, When auditoría mirror, Then G:AP + hidden byte-idénticos al template (fórmulas, valores, `t`/`ref`/`si`, `hidden`).
- **S7 (ceros intencionales):** Given celdas allow-list L-1, When writer + gates, Then siguen literal-0 y sin `<f>`; cualquier `<f>` ahí = FAIL.
- **S8 (base):** Given plantilla única + período P, When procedimiento §2.3-A, Then base P consistente (hojas `DetRetriP`/tokens/N3/sellos-base); runtime con base equivocada → fail-fast antes de escribir.
- **S9 (regresión):** Given Julio + Agosto vs NEW-base, When regens frescas, Then healthy-open + `<f>`-idéntica + DetRetri-D 5/5 vs R10 ±0.5 + shared/calcChain verdes.

---

## 4. TASKS

| ID | Tarea | Grupo | Depende de | Verificación |
|---|---|---|---|---|
| T0 | Contrato de slots congelado (cero código). Leer template único (BCL ZIP+XML): slot map por ASE (empresa×tipo→celda, AFaseo dual, NUEVO ESQUEMA 5→8, labels-C, techos fail-fast) + allow-list L-1 + NO-TOUCH mirror/hidden + citas E1..E5. Con Q1..Q4 respondidas y revisión del Ingeniero. | T0 | Q1..Q4 | R-T-1/R-T-2 (S1/S5 parcial-fixture); `git diff --stat` vacío en producción; el fixture manda |
| T1 | Writer values-only (TDD). Nuevo path solo-`<v>` con resolutores intactos (firma+frontera+sinonimia) + assert sin-`<f>` + `ERR-SIN-SLOT` (D-A/D-B/D-F). Tests: S2/S3/S4/S5 por ASE con fuentes reales por ruta explícita a temp FRESCA. | B+G | T0 | R-B-1/R-B-2/R-B-4/R-T-2 (S2/S3/S4/S5); build 0 warnings; grep sin insert/delete ni `<f>`-escritura en path R1 |
| T2 | Retiro del engine (net-negative). DELETE dimensionado + Reanclar-por-delta-R1 + escritura-`<f>` R1 + `CeldasInterioresAgosto`/`EsPeriodoJulio` (si viven). Conservar clasificadores/lectura/sinonimia/calcChain/sellos. | B | T1 | R-B-3/R-B-4 (grep de cierre + `git diff --stat` neto-negativo); suite sin ellos en verde |
| T3 | Base + sellos (§2.3-A). Extender `PreparadorBasePeriodo` a plantilla única (placeholder `2026xx-x` + base 2026092) + preflight de base en runtime + sellos G7/K7/D6/D7 vía `NombresHojaPeriodo`. Runtime sin tocar `<f>` de período. | B | T2 | R-B-5 (S8); build 0 warnings |
| T4 | Regresión 5 ASE × ambos meses vs NEW-base. Healthy-open + `<f>`-idéntica (S6/S7) + DetRetri-D vs R10 ±0.5 + shared/calcChain ciegos + comparador sin exenciones nuevas. Solo rutas explícitas. | R | T3 | R-G-1/R-G-2/R-R-1 (S6/S7/S9); suite completa verde salvo 511 dependencia listada |
| T5 | Docs: `project-context.md` (doctrina slot D-A..D-G + Allow-list + NO-TOUCH + qubit AFaseo + lección C1..C4) + manual (mes-nuevo supervisado + lectura de cada FAIL S3/S4/S8). Sin código en el mismo commit. | D | T4 | R-R-2; diff acotado; CRLF |
| T6 | Cierre global: build 0 warnings + suite verde (salvo dependencia declarada) + goldens ±0.5 + DetRetri-D 5/5 + greps globales (D-H, catálogo-abierto, sin-`<f>`-escritura-R1, sin-mirror-touch) + CRLF + sin commits. | Transversal | T0..T5 | R-R-1; DoD §5 |

**Orden sugerido (puertas duras):** Q1..Q4 → T0 (contrato, sin código) → T1 (writer values-only) → T2 (retiro, prohibido sin T1 verde) → T3 (base+sellos) → T4 (regresión global) → T5 → T6. Build 0 warnings antes de marcar cada tarea completa. Sin commits (los hace el Ingeniero con `#commit`/`#push`).

---

## 5. Resumen ejecutivo, riesgos y decisiones

- **Resumen:** por mandato #2067 la plantilla estandarizada única sustituye a las bases por mes; la escritura R1 pasa de espejo-dimensional + recomposición `<f>` a pegado de valores en slots fijos resueltos por el arsenal vigente (firma, frontera Plan 35-T2, sinonimia Plan 34). El faseo con puertas (contrato congelado → writer → retiro neto-negativo → base por procedimiento → regresión 5 ASE → docs) contiene el switch en todos los ASE: C1 se cierra con AFaseo dual-firma, C4 con NO-TOUCH mirror por construcción, C2/C3 se disuelven (ceros respetados, orden irrelevante), y los ceros ex-fórmula quedan sellados como defaults (allow-list). La base de cada período sale de la herramienta offline extendida, nunca del runtime.
- **Riesgo principal:** R-ALL-ASE (cambio de comportamiento en los 5 ASE a la vez: +22..+32 filas por bloque, NUEVO ESQUEMA 5→8, labels C). Contenido por quíntuple red faseada: contrato por ASE con techos fail-fast (T0), writer con gates por ASE (T1), retiro prohibido sin T1 verde (T2), base validada por preflight (T3), y regresión 5 ASE × ambos meses con `<f>`-idéntica + DetRetri-D vs R10 + shared/calcChain ciegos + healthy-open (T4).
- **Qué pasa con 21/31/32/35-T1:** el espejo-DIMENSIONAL del Plan 21 queda RETIRADO para R1 (su doctrina "la fuente define la forma" se sustituye por slot doctrine: el template define los slots, la fuente define los valores); las partes de ESCRITURA de 31/32 (`RecomponerVisibles`/`RecomponerInterior` que escriben `<f>`, dimensionamiento, Reanclar por delta) quedan RETIRADAS y borradas en T2; sus partes de LECTURA/clasificación (firmas, roles, `EsAplicacionPorFrontera`, sinonimia, quirk `--L` como dato de plantilla) se CONSERVAN como resolutores; el T1 de equivalencia del Plan 35 (descubrimiento==mapa) queda SUBSUMIDO por T0 (el descubrimiento se reorienta a firma→slot, no a geometría→recomposición) mientras su T2-clasificador queda APROBADO y vigente.
- **Decisión para el Ingeniero:** ratificar D-A..D-G (§0.4) + responder Q1..Q4 (§0.2). **Forks con recomendación:** §2.3 base-A vs base-B (recomendada A: única que preserva "runtime NUNCA toca `<f>`"); alcance (F209/F219 y R4/CF-DV/DetRetri/UI/fixture-root como OUT/dependencia — recomendado: diff mínimo, cada frente con su propio T0). **Sin preguntas técnicas bloqueantes pendientes; gate de entrada = Q1..Q4.**

### DoD (Definition of Done)

1. `dotnet build Remuneracion.WinForms/Remuneracion.WinForms.slnx` con 0 warnings.
2. Puertas en orden: T0 congelado del disco (Q1..Q4 + revisión) → T1 values-only (S2/S3/S4/S5) → T2 retiro con grep 0 + diff neto-negativo → T3 base consistente + sellos → T4 regresión 5 ASE × ambos meses (healthy-open + `<f>`-idéntica + DetRetri-D 5/5 vs R10 ±0.5 + shared/calcChain verdes).
3. Greps de cierre: 0 literales/ramas por período (D-H); 0 ramas por empresa fuera de filas citadas; 0 insert/delete/resize y 0 escritura `<f>` en path R1; mirror/hidden byte-idénticos; ningún `<v>` como oráculo de fórmulas.
4. `project-context.md` + manual actualizados (incl. procedimiento mes-nuevo supervisado); regens siempre a temp FRESCA con rutas explícitas; sin commits del agente; CRLF; sin emojis.

### Follow-ups explícitos (fuera de este plan)

- F209/F219 (y resto L-1 genuino): su propio diseño si Q3/T0 revelan slot real; si no, permanecen en allow-list sellada.
- Motor-espejo R4, R4-por-empresa Q2, CF/DV (W-4), `R-EXTRA-CONCEPTO` (Plan 23), validación de nombres de hoja en preflight, remesh R2/R4, re-encendido SALDOS/AJUSTES, `INTERVENTORIA!R26`: vivos, intactos.
- Reconciliación de fixture-root (511 rojos post-reorg): dependencia declarada; lo nuevo usa rutas explícitas y nunca `Docs/Insumos` raíz.
- Retiro/archivo de evidencias: regens en temp (no versionadas); goldens congelados (incl. primer golden de cada mes nuevo supervisado) sí.
