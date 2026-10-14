# Plan 38 — Recongelamiento de mapas a la geometría nueva de la plantilla estandarizada (cierre del bloqueador Plan 36-T4)

> **Alcance:** cerrar el bloqueador confirmado del Plan 36-T4 — la familia de mapas de celdas congelados (cell-map family) y las validaciones de fórmula protegida siguen ancladas a la geometría LEGACY (templates por mes) mientras el runtime ya escribe values-only sobre la plantilla estandarizada única — re-congelando TODA la familia al contrato nuevo derivado por evidencia desde los bytes del disco, migrando (re-congelar o retirar con motivo) los tests estructurales que congelan geometría legacy, y definiendo el gate de colocación subsidio vs no-subsidio que el implementer del Plan 36 dejó como riesgo pendiente. Cambio de producción esperado: **mínimo y acotado a direcciones/cadenas** (se mueven literales de celda y fragmentos de fórmula esperados; cero cambios de semántica de negocio, cero cambios de flujo). La escritura R1 values-only del Plan 36 queda INTACTA.
> **Rector (IRRENUNCIABLE):** `README.md` + `Docs/Propuesta_Proyecto_Automatizacion_Remuneracion_UAESP.md` + `Docs/Detalle de plantilla.docx` + `Docs/Proceso de Recaudo.docx` + `Docs/Prompt_Maestro_Proyecto_Remuneracion_UAESP _ Vo .docx` + `.opencode/project-context.md`. **Invariantes:** tolerancia ±0.5 en goldens/comparaciones numéricas; NUNCA usar cachés `<v>` como oráculo (y aquí `<f>` debe quedar byte-idéntica al template — se aserta); doctrina calcChain Plan 28 intacta (delete calcChain + fullCalcOnLoad); slot doctrine Plan 36 intacta (runtime NUNCA inserta/borra filas, NUNCA escribe `<f>` en R1); build 0 warnings antes de marcar cualquier tarea completa; CRLF sin mezclar; NUNCA git commit/push (los hace el Ingeniero con `#commit`/`#push`); la implementación la ejecuta `implementer` usando SOLO este plan aprobado; `code-reviewer` verifica spec compliance + 0 warnings. **Quincena por dominio, nunca por contenido.**
> **Mandato (este plan):** la plantilla estandarizada única (`Docs/Plantilla estandarizada/Formato Remuneracion 2026xx-x Total.xlsx`, período base interno 2026092) es el ÚNICO contrato geométrico. Toda dirección congelada en código o en tests que no coincida con sus bytes es un defecto, sin excepción. El fixture manda: bytes del disco, no este documento (los números citados aquí son punteros, no autoridad).
> **Continuidad:** Planes 21 (espejo — retirado para R1 por Plan 36), 25 (roles por firma — vigente como resolutor), 26/27 (preflight + firma + hoja-por-nombre — vigentes), 28 (calcChain + sello — intactos), 29 (paridad + comparador f-vs-f — vigente), 30 (naming dinámico + base canónica — vigente; `NombresHojaPeriodo` manda en nombres de hoja), 31/32/35-partes de escritura (retiradas por Plan 36), 33 (shared-preservation — intacta), 34 (sinonimia como DATOS + D-H + catálogo abierto — vigentes), 36 (slot doctrine + writer values-only + bases por período en `Docs/BasesPreparadas/` — vigente; su T4 deja el bloqueador que ESTE plan cierra y el riesgo de semántica subsidio que ESTE plan gatea). Este plan NO reabre lectura R1/R2/R4, cálculo CONSOLIDADO, preflight, sellos, naming, calcChain, paridad, ni el writer values-only fuera de lo declarado.
> **Insumos canónicos:** plantilla única `Docs/Plantilla estandarizada/Formato Remuneracion 2026xx-x Total.xlsx` (oráculo geométrico); bases `Docs/BasesPreparadas/Plantilla_Remuneracion_{2026071,2026072,2026082}.xlsx` (verificado en disco: existen las tres); fuentes R1/R2/R4 por ASE por ruta explícita; oráculos manuales julio-2/agosto-2 solo como referencia de regresión; `ContratoSlotR1T0Tests.cs` (ya congela geometría nueva parcial: 37 masters `si=0..36`, 1101 seguidoras, `DetRetri2026092`/`DetValiRetri2026092`, `D111=D109+N109`, N3=2026092 — se REUTILIZA, no se duplica); familia `WorkbookLeafCellMap*.cs` + `WorkbookLeafCoherence.cs` + `OpenXmlPlantillaWriter.cs` + `ValidadorTotalesR1Workbook.cs` + `ValidadorInsumosPeriodo.cs`; tests estructurales `MapaQ2T0Tests`, `Golden*Q2Tests`, `InterventoriaTests`, `BasePeriodoPreflightTests`. Regens SIEMPRE a rutas temp FRESCAS.
> **Numeración:** `plans/` 01..37 ocupados (37 + anexos T0); este plan toma el primer correlativo libre, **38**.
> **Estado:** APROBADO por el Ingeniero 2026-10-14 (ratifica: la plantilla única es el único contrato geométrico para todos los períodos; toda dirección legacy es defecto; tests legacy se re-congelan o se retiran con motivo, nunca se relajan en silencio; el gate de subsidio se define aquí con prueba al tener goldens comparables). EN EJECUCIÓN.
> **Fecha:** 2026-10-14

---

## 0. Clarification Gate

**Bloqueador adoptado, no re-diagnosticado (evidencia del encargo).** 72 rojos de la suite (de 96 totales) fallan con `La celda 'F53' de 'Reporte Componentes R1' debería seguir siendo fórmula; se detectó un valor` (misma clase para `F46`) más aserciones de cadena (`CONSOLIDADO_TOTAL RECAUDO!D9 debió ser fórmula con fragmentos [F53, Reporte Componentes R1]`, `DetRetri2026072!D32 debió ser fórmula con fragmentos [D23,F53,E43,D73,F55,AJUSTES,D47,D9]`, `Rem. Anticipos R2!E43 debió ser fórmula con fragmentos [E17,E28,K17]`, `SALDOS POR NOTA!C9 debió ser fórmula`, `Reporte Componentes R1!F265 (ASE3) debió ser VALOR editable`, `T0_VisiblesR1Q2_SonFormulasConFragmentosExactos`, `T0_M1Estructural_*`, `CapaA_Q2_*`, `W1_*`, `D3b_*`). Causa raíz adoptada: la familia de mapas congelados + validaciones de fórmula protegida + tests estructurales anclan geometría legacy. El T0 de ESTE plan la re-congela desde el disco; ningún T-gate re-diagnostica la causa.

### 0.1 Punteros medidos (NO autoridad — el T0 congela sus propios números desde el disco)

| # | Puntero (a re-congelar por T0 con cita celda-a-celda) | Consumidor actual en legacy |
|---|---|---|
| P1 | Bloques de datos R1 por ASE: ASE1 6-72, ASE2 110-210, ASE3 289-355, ASE4 450-550, ASE5 603-676 (contrato T0 del Plan 36) | `WorkbookLeafCellMapQ2` (visibles `F53/F206/F343/F468/F558`…), `WorkbookLeafCellMapPorAse`, tests `T0_VisiblesR1Q2_*`, `T0_M1Estructural_*` |
| P2 | `CONSOLIDADO!D9:D13` → R1 `F75/F254/F415/F568/F680`; `D28:D32` → `Rem. Anticipos R2!E47/E144/E262/E362/E446`; `D66:D70` → `Reversion Pagos R4!D73/D174/D211/D326/D361`; NUEVA `D111=D109+N109` | `WorkbookLeafCellMap.PorDefecto` (verificado: `(HojaR1,"F46",…)`, `(HojaConsolidado,"D9",["F46",…])`), `WorkbookLeafCellMapPorAse`, `OpenXmlPlantillaWriter.ValidarCadenaFormulaR1R2R4` + cadena CONSOLIDADO (verificado: `F46/F48`, `E41`, `D67`, `D9→F46`) |
| P3 | Masters shared 35→37 (`si=0..36`); seguidoras 1035→1101 | `SharedFormulaR1Esperados` / gate shared (Plan 33); ya re-congelado parcial en `ContratoSlotR1T0Tests` (verificado en disco) |
| P4 | 11 tokens `<f>` de período; base interna 2026092; hojas `DetRetri2026092` / `DetValiRetri2026092` / `Informe AFaseo Recaudo 202609-2` | `PreparadorBasePeriodo`, `BasePeriodoPreflightTests`, `NombresHojaPeriodo`, `WorkbookLeafCellMapDetRetri` |
| P5 | Referencias `DetRetri/DetValiRetri/INTERVENTORIA/VALIDACION` movidas (cadenas con `D23/F53/E43/D73/F55/AJUSTES/D47/D9` legacy como ejemplo) | `WorkbookLeafCellMapDetRetri`, `WorkbookLeafCellMapInterventoria`, `WorkbookLeafCellMapValidaciones`, `WorkbookLeafCellMapAjustesSfT`, `WorkbookLeafCellMapReporteBanco`, `WorkbookLeafCellMapBalanceSc`, `WorkbookLeafCellMapDetalleR2R4` |

Verificación propia del planificador (read-only, 2026-10-14): existen en disco `WorkbookLeafCellMap.cs`, `WorkbookLeafCellMapQ2.cs`, `WorkbookLeafCellMapPorAse.cs`, `WorkbookLeafCellMapPorEmpresa.cs`, `WorkbookLeafCellMapDetRetri.cs`, `WorkbookLeafCellMapBalanceSc.cs`, `WorkbookLeafCellMapAjustesSfT.cs`, `WorkbookLeafCellMapReporteBanco.cs`, `WorkbookLeafCellMapInterventoria.cs`, `WorkbookLeafCellMapDetalleR2R4.cs`, `WorkbookLeafCellMapValidaciones.cs`, `WorkbookLeafCoherence.cs` (todos en `Remuneracion.Infrastructure/Excel/`); `OpenXmlPlantillaWriter.cs` contiene los call sites `ValidarCeldaTieneFormula` (L75-76 R1 `F46/F48`, L94 R2 `E41`, L111 R4 `D67`, L454-465 cadena CONSOLIDADO, L474-483 `ValidarCadenaFormulaR1R2R4`, L705+ listas protegidas Q2) y `WorkbookLeafCellMapQ2.cs` congela visibles legacy (`F53=[F32,F48,F12,L12,L32]`, `F468`, `F558/F560`, CONSOLIDADO `D9→F53`, DetRetri `D32→[D23,F53,E43,D73,F55,AJUSTES,D47,D9]`); plantilla única + 3 bases preparadas existen en disco; `ContratoSlotR1T0Tests.cs` ya congela P3/P4 parcial. **Lo que T0 debe aún congelar desde bytes (no verificado por este plan):** el contenido exacto de cada `<f>` nueva, cada celda-dato vs celda-fórmula por bloque ASE, y cada referencia DetRetri/DetValiRetri/INTERVENTORIA/VALIDACION.

### 0.2 Gate de entrada — pregunta única al Ingeniero (BLOQUEANTE para T1..Tn)

| ID | Pregunta | Efecto según respuesta |
|---|---|---|
| Q1 | ¿Se ratifica que la plantilla única es el ÚNICO contrato geométrico y que toda dirección legacy en código/tests es defecto (este plan la re-congela), y que los tests estructurales legacy se re-congelan a la base nueva o se retiran con motivo declarado — nunca se silencian ni se relajan? | "Sí" = puerta a T0..T6. "No / con matices" = el Ingeniero lista qué test legacy debe sobrevivir intacto y con qué justificación; ese test entra como exclusión versionada con cita (no blanket) y el resto avanza. |

### 0.3 Mapeo al Rector (Sale / No sale)

**Sale (EXPLÍCITO):**
- Re-diagnosticar la causa del bloqueador (adoptada) o reabrir "template por mes vs plantilla única" (mandato Plan 36 #2067 lo cierra).
- Cambiar semántica de negocio, flujos de lectura/cálculo/escritura, el writer values-only, sellos, naming, preflight, calcChain, shared-preservation, sinonimia o tolerancia ±0.5.
- Cualquier `InsertRow`/`DeleteRow`/resize en R1; cualquier escritura de `<f>` en el path R1 (`new CellFormula`, `.Text` de fórmula fuera de preservación shared declarada); restaurar fórmulas legacy donde el template trae 0 intencional.
- Relajar o silenciar un test estructural legacy sin motivo declarado y versionado con cita (prohibido "lo marco Skip", prohibido borrar asserts para que pase).
- Ramas por nombre de empresa fuera de las FILAS citadas de `SinonimosEmpresaR1`; literales/ramas por período en el diff de producción (D-H); cachés `<v>` como oráculo; paquetes NuGet; commits/push.
- Renombrar rótulos del template en runtime.

**No sale (entra, alcance congelado):**
| Frente | Superficie de cambio |
|---|---|
| T0 — evidencia | Congelamiento desde bytes (BCL ZIP+XML) de direcciones/cadenas nuevas por familia de mapa + listas protegidas + refs DetRetri/DetValiRetri/INTERVENTORIA/VALIDACION, con citas celda-a-celda; cero código de producción |
| B1 — validaciones protegidas | `OpenXmlPlantillaWriter` (call sites `ValidarCeldaTieneFormula` + `ValidarCadenaFormulaR1R2R4` + listas protegidas) al contrato nuevo |
| B2 — familia de mapas | `WorkbookLeafCellMap*.cs` (+ coherencia/validadores que consumen direcciones) a geometría nueva, con drift gate |
| B3 — tests estructurales | Re-congelar a la base nueva o retirar con motivo declarado (nunca relajar en silencio) |
| B4 — gate subsidio | Definición del gate de colocación subsidio vs no-subsidio (riesgo Plan 36-T4) + prueba cuando los goldens sean comparables |
| G — gates | `<f>`-idéntica-al-template + DetRetri-D vs R10 ±0.5 ambos períodos + shared/calcChain ciegos + healthy-open + greps de cierre |
| D — docs | `project-context.md` (contrato geométrico nuevo + procedimiento de re-congelamiento ante cambio de plantilla) |

**Fuera de scope explícito (OUT):** motor-espejo R4; cobertura `conditionalFormatting`/`dataValidations` (W-4); semántica DetRetri (redondeo/columna D intactos); plumbing UI/CLI (CLI solo como runner de repro); completar F209/F219 como celdas (follow-up con su propio T0 salvo que T0 las subsuma como slots reales).

### 0.4 Decisiones ejecutivas (aprobación = aceptar estas)

| ID | Decisión |
|---|---|
| D-A | **El disco manda.** Toda dirección/cadena congelada se deriva de los bytes de la plantilla única (BCL ZIP+XML, sin Excel/COM); el fixture T0 cita hoja+celda+texto y es la fuente única de verdad. Los punteros P1..P5 (§0.1) no son autoridad. |
| D-B | **Híbrido firma-donde-hay-firma / congelado-donde-es-estructural (§2.2).** Las direcciones de datos R1 con firma (empresa×tipo-fila) se resuelven por firma+clasificador+sinonimia (doctrina vigente, nada que congelar por celda); las direcciones ESTRUCTURALES (visibles R1, cadenas CONSOLIDADO, refs DetRetri/DetValiRetri/INTERVENTORIA/VALIDACION, listas protegidas) se re-congelan como datos tabulares con drift gate. |
| D-C | **Re-congelar o retirar con motivo — nunca relajar.** Cada test estructural legacy sale en verde contra la base nueva o se retira con razón citada y versionada. Un `Skip`, un assert borrado o una tolerancia ampliada sin cita es regresión, no migración. |
| D-D | **D-H + catálogo abierto irrenunciables.** Cero literales/ramas por período y cero ramas por empresa en el diff; el mecanismo de migración no recibe ningún período; empresas no tabuladas = identidad. |
| D-E | **La plantilla única es contrato fijo legítimo para congelar.** Congelar direcciones es correcto SI el congelamiento es derivado por evidencia (D-A) y lleva drift gate que falla ante cualquier cambio de plantilla (el próximo cambio exige su propio T0, nunca edición silenciosa). |
| D-F | **Gate de subsidio ahora, prueba cuando haya goldens comparables.** La definición del gate (qué separa subsidio vs no-subsidio en el slot writer) se entrega en este plan; su prueba de colocación corre en cuanto los goldens sean comparables, con trigger explícito (§4-T4). |

---

## 1. PROPOSE

### 1.1 Intent

Que la suite vuelva a verde sobre la plantilla estandarizada única — con la familia de mapas, las validaciones de fórmula protegida y los tests estructurales anclados a la geometría REAL del template (derivada por evidencia, con drift gate), sin tocar la escritura values-only ni ninguna semántica de negocio — y con el riesgo de colocación subsidio del Plan 36 convertido en gate definido con trigger de prueba explícito.

### 1.2 In Scope

- T0 de evidencia: congelamiento celda-a-celda desde bytes para cada familia de mapa + listas protegidas + refs DetRetri/DetValiRetri/INTERVENTORIA/VALIDACION, con fixture que manda y revisión del Ingeniero.
- Migración B1 (validaciones protegidas) + B2 (familia de mapas, data-driven donde aplique + drift gate) + B3 (tests estructurales: re-congelar o retirar con motivo).
- Gate de subsidio B4 (definición + trigger de prueba).
- Gates de aceptación §3 + docs mínimos.

### 1.3 Non-Goals

Ver §0.3 OUT. En particular: NO se toca el writer values-only salvo sus listas de direcciones esperadas; NO se restaura ninguna fórmula legacy; NO se "completan" F209/F219; NO se reabre ningún frente OUT con el pretexto de la migración.

### 1.4 Riesgos

| ID | Riesgo | Mitigación |
|---|---|---|
| R-DERIVA | El re-congelamiento copia un número mal (mano, no disco) y congela un defecto nuevo | D-A: el fixture lee bytes (BCL ZIP+XML); el test compara código-vs-fixture, no documento-vs-código; revisión del Ingeniero sobre el diff T0 |
| R-DRIFT-FUTURO | La plantilla cambia (negocio) y el código congelado miente en silencio | Drift gate: hash + conteo de `<f>`/masters/hojas; cualquier cambio → FAIL que exige nuevo T0 (procedimiento documentado) |
| R-ATRIBUCION | Al mover direcciones se atribuye dinero a la celda/empresa equivocada | Firma+sinonimia intactas como resolutores; DetRetri-D vs R10 ±0.5 por ASE como red; `<f>`-idéntica workbook-wide |
| R-SILENCIO | Un test legacy se "migra" relajándolo y el defecto que cubría vuelve | D-C: prohibido relajar sin cita; cada retiro lleva motivo versionado; `code-reviewer` verifica que ningún assert perdió severidad sin cita |
| R-SUBSIDIO | Firma sola no separa subsidio vs no-subsidio y el slot writer coloca en el slot equivocado | B4: gate definido en este plan con trigger explícito; hasta su prueba en verde, la colocación de esas filas queda marcada como NO-probada (no como correcta) |
| R-ALCANCE | La migración se expande a semántica DetRetri/R4/CF-DV/UI | §0.3 OUT + grep de cierre (diff = direcciones/cadenas + tests + gates + docs) |
| R-FIXTURE-ROOT | Gates nuevos heredan rojos de resolución raíz | Rutas explícitas (plantilla única + `Docs/BasesPreparadas/` + fuentes por ASE); prohibido `Insumos.Raiz()` y `Docs/Insumos` raíz en lo nuevo |

---

## 2. DESIGN

### 2.1 Enfoque: evidencia primero, migración tabular con drift gate, tests re-congelados o retirados con motivo

**Fase 0 — evidencia (T0, cero código de producción).** Lectura BCL (`ZipArchive` + XML, sin Excel/COM) de la plantilla única: por cada familia (§2.3), tabla `(hoja, celda, clase: dato|fórmula-protegida|valor-editable, texto-<f> o literal, fragmentos de cadena)` con cita al byte. Incluye: bloques R1 por ASE (celdas-dato vs celdas-fórmula, visibles con su texto exacto), cadenas CONSOLIDADO D (D9:D13, D28:D32, D47:D51, D66:D70, D104:D109, D111), DetRetri/DetValiRetri (D32 y clase con fragmentos), R2 (`E43` y clase), SALDOS POR NOTA (`C9`), INTERVENTORIA/VALIDACION/BCE/banco/AJUSTES-SF-T protegidas, celda `F265`-clase (ASE3 valor editable) y análogas por ASE, masters shared (`si=0..36`) + conteo de seguidoras, tokens `<f>` de período (11) + N3 + nombres de hoja del período base. Fixture que manda + revisión del Ingeniero (Q1).

**Fase 1 — migración (T1..T3).** Las direcciones/cadenas legacy se sustituyen por las T0 como DATOS tabulares (mismo tipo, misma severidad, mismas claves; solo cambian los literales de celda y los fragmentos esperados). Donde la dirección ya es resoluble por firma (datos R1), se elimina el literal y se deriva (D-B). Cada mapa/test migrado lleva su drift gate: el test falla si el byte del disco deja de coincidir con lo congelado.

**Fase 2 — gate subsidio (T4).** Definición + prueba cuando los goldens sean comparables (trigger §4-T4).

**Fase 3 — regresión + docs + cierre (T5..T6).**

### 2.2 Decisión de diseño con trade-offs: re-congelar vs derivar en runtime

| Eje | Opción A: re-congelar todo a datos tabulares | Opción B: derivar todo en runtime por firma/label | **Elegida: híbrido D-B** |
|---|---|---|---|
| Qué hace | Cada dirección vive en una tabla congelada desde T0; el runtime las consume tal cual | Cada dirección se descubre en runtime buscando firmas/labels en el workbook | **Datos R1 con firma → derivación (ya existe: firma+clasificador+sinonimia); estructural sin firma (visibles, cadenas CONSOLIDADO, DetRetri/VALIDACION/INTERVENTORIA, listas protegidas) → tabla congelada T0** |
| Pro | Determinista; diff mínimo; auditable celda-a-celda; barato | Cero hardcodeo; robusto ante reordenamientos del template | Lo derivable ya se deriva (cero costo nuevo); lo estructural queda determinista y auditado |
| Contra | Ante cambio de plantilla hay que re-congelar (costo conocido) | Costo alto: cada dirección estructural necesita su firma (muchas no la tienen: `D111`, cadenas CONSOLIDADO, `J9`-clase); un descubridor nuevo es un motor nuevo con sus propios falsos positivos/negativos; reintroduce la clase de fragilidad que los Planes 21-35 retiraron | Exige disciplina: la frontera firma-vs-estructural debe quedar documentada por celda (el T0 la fija) |
| Riesgo | R-DRIFT-FUTURO | Un label movido por el negocio cambia el comportamiento en silencio (peor que un FAIL) | Contenido: drift gate convierte el cambio silencioso en FAIL accionable |
| Veredicto | Parcial: solo donde es estructural | Parcial: solo donde hay firma real | **Híbrido D-B: derivar donde hay firma, congelar donde es estructural, drift gate en ambos** |

Descartado puro-A-total: congelar también lo que ya se deriva por firma sería regresar (reintroduciría literales donde la doctrina vigente ya derivó). Descartado puro-B-total: derivar direcciones estructurales sin firma exige inventar firmas (heurística) y convierte un contrato fijo y auditable en un descubrimiento frágil. Nota de legitimidad (D-E): la plantilla única es contrato fijo por mandato #2067; congelar es correcto SI es derivado por evidencia + drift gate.

### 2.3 Contratos afectados (resumen)

| Contrato | Cambio |
|---|---|
| Fixture T0 (nuevo, tests, cero producción) | Tablas congeladas desde bytes por familia + citas + hash/conteos de drift (masters, seguidoras, `<f>`, hojas, N3, D111) |
| `OpenXmlPlantillaWriter` (solo literales de direcciones/fragmentos) | `ValidarCeldaTieneFormula` call sites (R1 `F46/F48`→nuevas, R2 `E41`→nueva, R4 `D67`→nueva), `ValidarCadenaFormulaR1R2R4`, cadena CONSOLIDADO (`D9→F46`→`D9→F75…` por ASE), listas protegidas Q2 (R1 visibles, R2, R4, AJUSTES-SF-T, banco, BCE, INTERVENTORIA, validaciones); misma severidad, mismos mensajes |
| `WorkbookLeafCellMap` / `PorAse` / `Q2` | Visibles R1 legacy (`F53/F206/F343/F468/F558` y cadenas) → visibles nuevos (`F75/F254/F415/F568/F680` y clase, a confirmar por T0); `D9→F46` → `D9→F75…`; `F46/F48/F519/F521`-como-valor Q2 revisado contra T0 |
| `WorkbookLeafCellMapDetRetri` / `BalanceSc` / `AjustesSfT` / `ReporteBanco` / `Interventoria` / `DetalleR2R4` / `Validaciones` | Refs y listas protegidas a geometría nueva (incl. `D28:D32→E47/…`, `D66:D70→D73/…`, DetRetri `D32`-clase, `E43`-clase, `C9`-clase, `F265`-clase valor-editable por ASE) |
| `WorkbookLeafCoherence` / `ValidadorTotalesR1Workbook` / `ValidadorInsumosPeriodo` | Solo lo que consuma direcciones legacy (mismo predicado, nuevas celdas); sin cambios de lógica |
| Tests estructurales (`MapaQ2T0Tests`, `Golden*Q2Tests` CapaA, `W1_*`, `D3b_*`, `T0_M1Estructural_*`, compartidos) | Re-congelar a base nueva (mismo assert, nuevas celdas) o retiro con motivo citado versionado; `ContratoSlotR1T0Tests` (ya nuevo) se reutiliza/extiende, no se duplica |
| `project-context.md` | Contrato geométrico nuevo + regla de re-congelamiento (drift gate + procedimiento ante cambio de plantilla) + frontera firma-vs-estructural |

### 2.4 Architecture Validation Certificate

| Principio | Veredicto |
|---|---|
| **SRP** | ✅ Una razón de cambio por costura: T0 declara geometría (datos), B1/B2 mueven literales, B3 migra tests, B4 gatea subsidio, cada gate contrasta una regla. Ninguna rama por período ni por empresa. |
| **OCP** | ✅ Un cambio futuro de plantilla se absorbe con un nuevo T0 (procedimiento documentado), sin cambios de lógica; empresa no tabulada = identidad. |
| **DIP** | ✅ Tablas como datos puros; resolutores (firma/frontera/sinonimia) intactos y puros; tests con insumos reales + BCL ya presente; cero dependencias nuevas (OpenXML 3.5.1 ya referenciado). |
| **Best practices** | ✅ TDD con puertas (evidencia → migración → regresión); fail-fast que nombra hoja+celda+fragmentos; quincena por dominio; `<f>` intactas; tolerancia ±0.5; sin Excel/COM en pipeline ni tests; Visual Design Intent N/A (cero UI). |
| **Performance** | ✅ Migración O(direcciones) en tiempo de carga/test; drift gate O(`<f>` de la hoja) lineal; cero costo runtime nuevo (las tablas sustituyen literales, no agregan pases). |

---

## 3. SPEC

### 3.1 Requirements

| ID | Requisito | Grupo | Criterio de aceptación |
|---|---|---|---|
| R-T-1 | T0 congela desde bytes la geometría nueva por familia (R1 bloques+visibles, CONSOLIDADO D, DetRetri/DetValiRetri, R2, R4, SALDOS/NOTA, AJUSTES-SF-T, banco, BCE, INTERVENTORIA, VALIDACION) con citas celda-a-celda + hash/conteos de drift | T0 | Fixture que manda; Q1 respondida; revisión del Ingeniero; `git diff` de producción vacío para T0 |
| R-T-2 | Frontera firma-vs-estructural documentada por celda (qué se deriva, qué se congela) | T0 | Tabla en el fixture; ninguna celda estructural sin cita; ninguna celda con firma congelada por literal |
| R-B-1 | Listas protegidas del writer re-congeladas (R1/R2/R4 + CONSOLIDADO + Q2) con misma severidad y mensajes | B | Los 72 rojos clase F46/F53 en verde; FAIL sigue nombrando hoja+celda+fragmentos |
| R-B-2 | Familia `WorkbookLeafCellMap*` + coherencia/validadores a geometría nueva, data-driven donde aplique + drift gate | B | Cadenas `D9/D28/D32/D66/D111`, `E43`, `C9`, `F265`-clase y análogas en verde; grep: 0 direcciones legacy en producción |
| R-B-3 | D-H + catálogo abierto en el diff | B | Grep: 0 literales/ramas por período; 0 comparaciones por empresa fuera de filas citadas; el mecanismo no recibe ningún período |
| R-B-4 | Tests estructurales migrados: verde contra base nueva o retiro con motivo citado | B | `T0_VisiblesR1Q2_*`, `T0_M1Estructural_*`, `CapaA_Q2_*`, `W1_*`, `D3b_*` en verde o retirados con razón versionada; prohibido relajar sin cita |
| R-G-1 | `<f>` byte-idénticas al template workbook-wide (texto + `t`/`ref`/`si`) + `<v>` solo en slots/destinos del contrato | G | FAIL ante cualquier `<f>` divergente o `<v>` fuera de contrato; prohibido `<v>` como oráculo |
| R-G-2 | DetRetri-D 5/5 vs R10 ±0.5 ambos períodos + goldens + shared/calcChain ciegos + healthy-open | G | Regens FRESCAS a temp (bases `Docs/BasesPreparadas/` + fuentes reales por ruta explícita); Excel abre sano (paso manual del Ingeniero como aceptación final) |
| R-S-1 | Gate de subsidio definido (B4) con trigger de prueba explícito | B | Definición entregada (qué distingue subsidio vs no-subsidio: firma + clasificador + frontera + allow-list de slots); prueba en verde cuando los goldens sean comparables, o follow-up con su propio T0 si la evidencia lo exige |
| R-R-1 | Suite verde salvo dependencias declaradas + build 0 warnings por tarea | R | `dotnet build …slnx` 0 warnings + `dotnet test` por tarea; los 72 F46/F53-clase cerrados; el resto (24) triageado: verde o dependencia declarada con motivo |

### 3.2 Scenarios (Given/When/Then)

- **S1 (evidencia manda):** Given plantilla única en disco, When T0 lista una familia, Then cada fila = (hoja, celda, clase, texto-`<f>`/literal, fragmentos) citada al byte; el test código-vs-fixture pasa y el diff de producción está vacío.
- **S2 (protegida re-congelada):** Given salida post-migración, When validación protegida, Then `F75` (y análogas) es fórmula con sus fragmentos nuevos; `F265`-clase (ASE3) es valor editable; CONSOLIDADO `D9` es fórmula con fragmentos `[F75, Reporte Componentes R1]` (por ASE según T0).
- **S3 (cadenas):** Given salida Q2, When cadena DetRetri, Then `D32` es fórmula con los fragmentos nuevos de T0 (no los legacy `[D23,F53,E43,D73,F55,…]`); `E43`/`C9` análogas.
- **S4 (drift):** Given plantilla modificada (un `<f>`, un master, una hoja), When drift gate, Then FAIL que nombra qué cambió y exige nuevo T0 (nunca verde silencioso).
- **S5 (no-relajación):** Given test estructural legacy, When migrado, Then mismo assert con celdas nuevas (verde) o retiro con motivo citado; un assert con severidad reducida sin cita = FAIL de review.
- **S6 (subsidio):** Given filas de fuente en secciones subsidio vs no-subsidio con la misma firma, When slot writer, Then cada valor cae en su slot T0 (gate B4 lo demuestra con goldens comparables); hasta entonces el estado es NO-probado, nunca asumido.
- **S7 (regresión):** Given ambos períodos vs bases preparadas, When regens frescas, Then healthy-open + `<f>`-idéntica + DetRetri-D 5/5 vs R10 ±0.5 + shared/calcChain verdes + 72 F46/F53-clase cerrados.

---

## 4. TASKS

| ID | Tarea | Grupo | Depende de | Verificación |
|---|---|---|---|---|
| T0 | Evidencia congelada (cero código). Leer plantilla única (BCL ZIP+XML): tablas por familia (R1 bloques+visibles, CONSOLIDADO D incl. D111, DetRetri/DetValiRetri, R2, R4, SALDOS/NOTA, AJUSTES-SF-T, banco, BCE, INTERVENTORIA, VALIDACION) + frontera firma-vs-estructural + hash/conteos drift (masters 37, seguidoras 1101, `<f>` período 11, N3=2026092, hojas 40). Con Q1 respondida y revisión del Ingeniero. Reutiliza/extiende `ContratoSlotR1T0Tests` donde ya congele (no duplicar). | T0 | Q1 | R-T-1/R-T-2 (S1); `git diff --stat` vacío en producción; el fixture manda |
| T1 | Validaciones protegidas al contrato nuevo (TDD). Migrar en `OpenXmlPlantillaWriter`: `ValidarCadenaFormulaR1R2R4` + call sites R1/R2/R4 + cadena CONSOLIDADO + listas Q2 (misma severidad/mensajes, nuevos literales T0). Tests: S2/S3 por ASE con bases preparadas + fuentes reales por ruta explícita a temp FRESCA. | B+G | T0 | R-B-1 (S2/S3); build 0 warnings; 72 F46/F53-clase en verde |
| T2 | Familia de mapas a geometría nueva. Migrar `WorkbookLeafCellMap*.cs` + ajustes mínimos en `WorkbookLeafCoherence`/`ValidadorTotalesR1Workbook`/`ValidadorInsumosPeriodo` (solo direcciones, cero lógica) + drift gate por familia. Donde haya firma real, derivar en vez de congelar (D-B, con cita a la frontera T0). | B | T1 | R-B-2/R-B-3 (S2/S3/S4); grep 0 direcciones legacy + 0 literales de período en el diff |
| T3 | Tests estructurales: re-congelar o retirar con motivo. `MapaQ2T0Tests`, `Golden*Q2Tests` CapaA, `W1_*`, `D3b_*`, `T0_M1Estructural_*` y hermanos: mismo assert con celdas T0, o retiro con razón citada versionada (lista explícita en el commit del Ingeniero). | B | T2 | R-B-4 (S5); build 0 warnings; ningún `Skip`/tolerancia ampliada sin cita |
| T4 | Gate subsidio (riesgo Plan 36-T4). (a) Definir qué distingue subsidio vs no-subsidio en el slot writer (firma + `EsAplicacionPorFrontera` + frontera de sección + allow-list de slots por ASE) con citas al disco. (b) Prueba de colocación **trigger:** corre cuando los goldens sean comparables (T5 verde en DetRetri-D); si la evidencia muestra que exige su propio diseño, se abre follow-up con T0 propio en vez de forzarla aquí. Hasta su verde: estado NO-probado declarado. | B | T2 | R-S-1 (S6); definición citada; trigger explícito |
| T5 | Regresión global. Regens FRESCAS a temp (ambos períodos × 5 ASE vs bases preparadas): healthy-open + `<f>`-idéntica (S6/S7) + DetRetri-D vs R10 ±0.5 + shared/calcChain ciegos + comparador sin exenciones nuevas. Triage de los 24 no-F46/F53: verde o dependencia declarada con motivo. Solo rutas explícitas. | R | T3 | R-G-1/R-G-2/R-R-1 (S7); suite verde salvo dependencias declaradas |
| T6 | Docs + cierre. `project-context.md` (contrato geométrico nuevo + frontera D-B + procedimiento de re-congelamiento ante cambio de plantilla + estado del gate subsidio) + build 0 warnings + greps globales (D-H, catálogo-abierto, sin insert/delete, sin escritura `<f>` en path R1, 0 direcciones legacy) + CRLF + sin commits. | D+Transversal | T0..T5 | R-R-1; DoD §5 |

**Orden sugerido (puertas duras):** Q1 → T0 (evidencia, sin código) → T1 (protegidas) → T2 (mapas + drift) → T3 (tests: verde o retiro motivado) → T4 (subsidio: definición + prueba por trigger) → T5 (regresión global) → T6. Build 0 warnings antes de marcar cada tarea completa. Sin commits (los hace el Ingeniero con `#commit`/`#push`).

---

## 5. Resumen ejecutivo, riesgos y decisiones

- **Resumen:** el Plan 36 cambió el paradigma a plantilla única + values-only pero dejó la familia de mapas/protegidas/tests anclada a geometría legacy (72/96 rojos). Este plan la re-congela desde los bytes: T0 de evidencia por familia (el fixture manda), migración tabular de validaciones (T1) y mapas (T2) con drift gate, migración de tests estructurales con regla re-congelar-o-retirar-con-motivo (T3), gate del riesgo subsidio con trigger explícito (T4), y regresión global con triage declarado del resto (T5). Decisión de diseño: híbrido firma-donde-hay-firma / congelado-donde-es-estructural — derivar lo estructural sin firma sería inventar heurística frágil; congelar lo que ya se deriva por firma sería regresar. D-H y catálogo abierto intactos; la plantilla única como contrato fijo hace legítimo congelar SI es por evidencia + drift gate.
- **Riesgo principal:** R-DERIVA + R-SILENCIO (congelar mal un número; relajar un test para que pase). Contenido por triple red: fixture que manda desde bytes + revisión del Ingeniero (T0), drift gate que convierte cambios futuros en FAIL (T2), y prohibición de relajar sin cita con verificación de `code-reviewer` (T3/D-C).
- **Qué cierra cada clase de rojo:** F46/F53-clase (72) → T1+T2 (S2/S3); cadenas DetRetri/R2/SALDOS (`D32/E43/C9`) → T1+T2 (S3); `F265`-clase valor-editable → T2 (S2); `T0_VisiblesR1Q2/T0_M1Estructural/CapaA_Q2/W1/D3b` → T3 (S5); resto (24) → T5 triage verde-o-dependencia-declarada (S7); subsidio → T4 (S6, NO-probado hasta su verde).
- **Decisión para el Ingeniero:** ratificar D-A..D-F (§0.4) + responder Q1 (§0.2). **Forks con recomendación:** híbrido D-B vs puro-A vs puro-B (recomendado híbrido: §2.2); T4 como definición+trigger dentro de este plan vs follow-up separado (recomendado dentro, con escape a follow-up si la evidencia lo exige); alcance OUT (§0.3, recomendado: diff mínimo). **Pregunta exacta de confirmación:** «¿Ratifica que la plantilla única es el único contrato geométrico (toda dirección legacy es defecto), que los tests legacy se re-congelan o se retiran con motivo — nunca se relajan en silencio — y que el gate de subsidio se define en este plan con prueba al tener goldens comparables (D-A..D-F)? Responda "sí, proceder" o liste las excepciones con su justificación.»

### DoD (Definition of Done)

1. `dotnet build Remuneracion.WinForms/Remuneracion.WinForms.slnx` con 0 warnings.
2. Puertas en orden: T0 congelado del disco (Q1 + revisión) → T1 protegidas (72 F46/F53-clase en verde) → T2 mapas + drift gate → T3 tests verdes o retirados con motivo → T4 subsidio definido (prueba por trigger) → T5 regresión global.
3. `<f>`-idéntica workbook-wide + DetRetri-D 5/5 vs R10 ±0.5 ambos períodos + shared/calcChain verdes + healthy-open.
4. Greps de cierre: 0 direcciones legacy en producción; 0 literales/ramas por período (D-H); 0 ramas por empresa fuera de filas citadas; 0 insert/delete/resize y 0 escritura `<f>` en path R1; ningún `<v>` como oráculo.
5. `project-context.md` actualizado (contrato nuevo + procedimiento de re-congelamiento + estado subsidio); regens siempre a temp FRESCA con rutas explícitas; sin commits del agente; CRLF; sin emojis.

### Follow-ups explícitos (fuera de este plan)

- Cambio futuro de la plantilla por el negocio: su propio T0 (procedimiento §T0 + drift gate como detector).
- Gate subsidio si la evidencia de T4 exige diseño propio (T0 + plan hijo).
- F209/F219 como celdas (su propio T0 si no son slots T0); motor-espejo R4; CF/DV (W-4); DetRetri-semántica; UI/CLI-plumbing; reconciliación de fixture-root (dependencia declarada; lo nuevo usa rutas explícitas).
