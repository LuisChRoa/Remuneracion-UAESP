# Plan 29 — Cierre de brecha app vs manual (BCE + desglose R2/R4 + DetRetri + proceso)

> **Alcance:** cerrar la brecha verificable entre la salida de la app y el archivo manual del administrativo en 2026072 y 2026082, con T0 bloqueante que arbitra cada categoría. **Veredicto anticipado de esta planificación (evidencia zip+XML en §0.1): NO hay una única causa** — el encargo suponía una; el disco demuestra **cuatro causas distintas más dos no-brechas** (artefactos de comparación). El plan congela ese alcance: Unidad B (BCE), Unidad R (desglose-detalle R2/R4), Unidad D (desglose DetRetri, condicionada a T0), Unidad P (proceso + cosmética + comparador de regresión). Cero cambios de fórmulas; DetRetri-D 5/5 vs R10 y goldens julio ±0.5 intactos por construcción.
> **Rector (IRRENUNCIABLE):** `README.md` + `Docs/Propuesta_Proyecto_Automatizacion_Remuneracion_UAESP.md` + `Docs/Detalle de plantilla.docx` + `Docs/Prompt_Maestro_Proyecto_Remuneracion_UAESP _ Vo .docx` + `.opencode/project-context.md`. **Invariantes del proyecto:** NUNCA sobrescribir fórmulas (solo valores); tolerancia ±0.5; ASE4 sin columna "Especiales" (0 si ausente); Q2 aplica SALDOS POR NOTA/RETRIBUCION NEGATIVA (`Periodo.NumeroQuincena == 2`, `AjustesSfT = 0` en Q1); redondear a entero antes de DetRetri/columna D; build 0 warnings; finales de línea CRLF; fail-fast nombrando archivo/celda; tests SOLO con insumos reales; sin commits (los hace el Ingeniero con `#commit`); sin emojis.
> **Continuidad:** HU-01..HU-23 cerradas; Planes 21 (espejo R1), 23 (opcionalidad 2.5), 25 (roles por firma), 26 (preflight), 27 (firma + hoja-por-nombre, suite 339/339), 28 (calcChain + sello fechas, suite 345/345) cerrados. Este plan NO reabre ninguna semántica de lectura, cálculo, espejo, preflight ni sello fuera de lo declarado: **cero cambios de fórmulas de negocio (`<f>` intacto en todo el diff); Q1/Q2-julio intactos por construcción.**
> **Insumos canónicos (no se re-fijan, no se inventan datos):** plantilla canónica `Docs/Plantilla_Remuneracion.xlsx` (672 KB, per. 2026072, 40 hojas); app/manual julio `Docs/Prueba Julio-2/Resultado/Remuneración 202607-2 Total.xlsx` / `Remuneracion 202607-2 Total Administrativo.xlsx` + insumos `Docs/Prueba Julio-2/Insumos` (5 ASE + `Conciliaciones/` + `R10_Remuneracion_2026072.xlsx`); app/manual agosto `Docs/Prueba2/Resultado/Remuneración 202608-2 Total.xlsx` / `Resultado Manual por el administrativo/Remuneracion 202608-2 Total_7721.xlsx` + insumos `Docs/Prueba2/Insumos` (+ `R10_Remuneracion_2026082.xlsx`, oráculo D9:D13/D14).
> **Numeración:** `plans/` 01..28 ocupados; este plan toma el primer correlativo libre, **29**.
> **Estado:** CERRADO (2026-10-07) — T0→T1→T2→T3→T4→T5+T6 ejecutados y verificados por el arquitecto. DoD: T0a..T0g en `plans/29-T0-Evidencia-A.md`/`-B.md`; build 0 warnings; suite 399/399 (363 base + 25 escritura-detalle + 6 desglose + 3 F10 + 2 regresión comparador); BCE 20/20 + F invariante; detalle R2/R4 ±0.5 vs fuente y manual; DetRetri trazable 12/12+11/12 (SALE: DetValiRetri J9); comparador verde julio+agosto desde ceros; cero `<f>` tocadas (verificado en diff); CRLF; sin emojis; sin commits (los hace el Ingeniero). Recortes aplicados: Unidad R por encabezado (ratificado), micro-fix 0-vacío NO aplicado (T0e), DetValiRetri J9 SALE, julio-ASE5 documentado como divergencia-del-manual.
> **Fecha:** 2026-10-07

---

## 0. Clarification Gate

**Sin preguntas bloqueantes: el plan viene redactado con las decisiones D-A..D-F (§0.3) y T0-bloqueante que solo puede RECORTAR alcance (Unidad D y sub-caso julio-ASE5), nunca expandirlo.** Si el Ingeniero discrepa de alguna decisión, solo cambia la unidad acotada que la implementa. Punto de bloqueo respetado: NO se implementa nada en este documento.

### 0.1 Qué está verificado y qué NO (honestidad técnica — evidencia T0 previa, parser zip+XML BCL, sin Excel)

**Verificado en esta planificación** (inspección directa de bytes en disco, 2026-10-07; `T0lib/T0q` en `%TEMP%\opencode`, reutilizables por el implementador):

| # | Hecho | Método | Consecuencia |
|---|---|---|---|
| V1 | La plantilla trae las hojas leaf como **literales-0 de captura** (`Rem. Anticipos R2` E3:R3, `Reversion Pagos R4` D3, `Recaudo ENEL` D3:G7, `BCE` D3:E7 = `<v>0</v>` sin `<f>`; `DetRetri2026072` D9:O14 todo literal 0) y los consolidados como **fórmulas** (`CONSOLIDADO D28='Rem. Anticipos R2'!E43`, `D66='Reversion Pagos R4'!D73`, `D9='Reporte Componentes R1'!F53`, `D104=D9+D28+D47+D66+D85`, `J13='BCE SC POR FACT.'!F7`; `REMUNERACION_*`/`GERENTES_*`/`VALIDACION_*` 100% fórmulas hacia R1/CONSOLIDADO) | zip+XML plantilla | La app solo puede cerrar la brecha **escribiendo literales** en las celdas de captura; todo lo demás recalcula solo. `ANT EXT-REV` C3/N3/M3 son fórmulas (`C11+C19+…`) → sin gap propio (T0 lo cierra al símbolo) |
| V2 | **R2/R4 agregados: app == manual al centavo.** App-julio `Rem. Anticipos R2` E17=126791992.42/E28=-843439.3/K17=0, E396=39162866.71/E410=4810642.25/K396=990006.55; `Reversion Pagos R4` D15=-16536713.09/P15=0, D352=-3092375.53 — idénticos en el manual (E17, D15, P15 incluidos). Agosto ASE1 idéntico (E17=93477927.32, E28=-13972083.45, K17=144941.68 ambos). Post-recalc E43=E17+E28-K17=125948553.12 = manual, D73=D15-P15=-16536713.09 = manual | zip+XML app vs manual, ambos períodos | **Refuta el encargo (A) y (c):** las hojas leaf R2/R4 NO están vacías; el encargo miró direcciones Q1 (`E15/E26/D3/D9`, literal-0 por diseño) en vez de las Q2 (`E17/E28/D15`, mapa `WorkbookLeafCellMapQ2` activo), y cachés `<v>0</v>` pre-recalc en vez de valores post-recalc. El mapa "E41/D67" del encargo es el mapa **Q1** (`WorkbookLeafCellMap` base); en Q2 rigen E43/D73-fórmula + celdas de captura Q2 — el writer SÍ las usa |
| V3 | **Recaudo por empresa: app == manual en columnas Q2.** App-julio `Recaudo ENEL` F3=16467129562/G3=273797 … F7=11684535987/G7=242618, idénticos al manual; D3:E7 (VALOR/N°REG **1°Q**) = 0 en app vs con datos en manual (D3=16020970408). Mismo patrón en agosto | zip+XML app vs manual | **No-brecha por diseño:** cada corrida escribe SOLO sus columnas (`Periodo.NumeroQuincena`: D/E en Q1, F/G en Q2 — project-context §Columnas). El manual acumula Q1+Q2 en un workbook; la app-Q2 nunca debe tocar D/E. Comparar D3 es comparar quincenas distintas. Va a Unidad P (proceso/documentación, cero código salvo veredicto) |
| V4 | **BCE ASE1-4: swap D/E puro, suma idéntica.** App-julio D3=+2452057655.23/E3=-6041512320.78 vs manual D3=-6041512320.78/E3=+2452057655.23 (F3=D3+E3=-3589454665.55 ambos). Agosto ASE1 y ASE5 igual (D7/E7 magnitudes idénticas, orden invertido; F7 idéntico). Header plantilla: D2=`SUBSIDIO`, E2=`CONTRIBUCION`; fuente: E=Subsidio (negativo), F=Contribución (positivo). La app escribe Contribución→D, Subsidio→E | zip+XML + lectura `WorkbookLeafCellMapBalanceSc.cs:44-56` (`[1]=("D3","E3")` con `Contribucion` primero) + `OpenXmlPlantillaWriter.cs:1239-1240` | Causa 1 (Unidad B): el veredicto T0-0.2 del Plan 10 ("D=CONTRIBUCION") contradice el header y al manual; como F=D+E es conmutativa, el "±0.5 vs golden" que la probó **no podía detectar el swap**. Aguas abajo neutro (J/K/M consumen F), pero D/E quedan visiblemente invertidas |
| V5 | **BCE julio-ASE5: NO es swap, son otros valores.** App D7=+2322461234.81/E7=-3615845886.78 (F7=-1293384651.97) vs manual D7=-1138650714.67/E7=+1055691691.01 (F7=-82959023.66). La fuente ASE5 (`R4-BalanceSubsidioyContribuciones-Optimizado…`, `TOTAL GENERAL` R39: E39=-3615845886.78/F39=+2322461234.81/G39=-1293384651.97) **respalda a la app al centavo** (D7←F39, E7←E39, F7=G39). Los valores del manual no aparecen en esa fila | zip+XML fuente + app + manual | Causa 2 (Unidad B, arbitraje T0): o el manual trae un error propio, o bebe de otra fuente (PDF del par, otro total). T0-bloqueante lo arbitra buscando D7/E7-manual en todos los insumos ASE5-julio; **si no aparecen → no hay cambio de código** (se documenta; la app es fuente-fiel). Explica J13 (-1.29e9 vs -82.9e6), REMUNERACION J/K/M/O, DetRetri-J y VALIDACION aguas abajo de ASE5 |
| V6 | **Desglose-detalle R2: el manual lo trae, la app no, y la fuente lo tiene.** Manual-julio `Rem. Anticipos R2` E3=104754634.94, F3:R3 con desglose por componente; app E3:R3 = 0. Fuente `RerpoteDetalleSaldosaFavor` ASE1 R5: E5=104754634.94/F5=6975298.69/… = manual E3/F3:… al centavo (con `Especiales` K=0 insertada y corrimiento TCS→L…R). Igual en agosto (E3=68638686.24). R4-detalle análogo (manual D3=-15794348.21/D9=-762167.27 literales; app 0; D15 agregado idéntico ambos) | zip+XML fuente + app + manual | Causa 3 (Unidad R): ningún mapa declara esas filas de detalle (Q1 y Q2 escriben solo agregados); el reader ni las extrae. No mueve DetRetri-D (los agregados cierran), pero es la brecha visible más grande en celdas |
| V7 | **DetRetri/DetValiRetri C..O: app solo D, manual todo literal.** App-julio D9=17450228673/E9:O9=0; manual D9=17450228673, E9=907089032 … J9=-3589454666 (=BCE F3), O9=15238755634 — todo `<v>` literal. DetRetri-D 5/5 vs R10 intacto en ambos períodos (cifras del encargo reproducidas en el R10 por Plan 28 §6.3) | zip+XML app vs manual | Causa 4 (Unidad D, **condicionada**): la fuente de cada columna C..O NO está identificada (E9-manual=907089032 ≠ `REMUNERACION_ENEL`!E104-manual=807895299; J9=BCE.F3 sí es trazable). Si T0 no demuestra derivabilidad workbook-interna por columna → SALE (paso administrativo manual, se documenta y no se inventa) |
| V8 | `VALIDACION_TOTAL` C15 = fórmula `C14=C9`, caché `1` (True) en **ambos** archivos julio; app C9-caché=0 (pre-recalc), manual C9=72497949967. `REPORTE RECAUDO x BANCO`: app escribe literal `0`, manual deja vacío (verificado en estructura; el writer escribe los 4 valores de `ReporteBancoInputs` sin saltar ceros) | zip+XML + lectura `OpenXmlPlantillaWriter.cs:1199-1211` | El `False` del encargo solo puede ser post-recalc (cachés stale no prueban nada) → T0 lo evalúa por igualdad de texto-fórmula + literales (método §2.1), no por caché. El `0`-vs-vacío va a Unidad P (micro-fix cosmético) |

**Aportado por el encargo como hipótesis de trabajo (NO verificado en disco en esta planificación; T0 del §4 lo cierra antes de implementar):** H1 = G7/K7-agosto sellados (46250/46265) y metadatos heredados 2026072 (N3, nombres de hoja) por correr agosto sobre la salida de julio — coherente con Plan 28 §6.3 pero no re-inspeccionado aquí; H2 = GERENTES E9 ~0.97% (898290673.52 vs 907089032.16) — sin localizar la celda-fuente del desvío (candidatos: `Reporte Componentes R1`!G63 app vs manual vs fuente, o arrastre de H1/BCE-ASE5); H3 = log de agosto ("veredicto D/E T0 hipótesis líder") como contexto del swap V4; H4 = valores DetRetri-D por período citados (oráculo R10 vigente, Plan 28 §6.3).

### 0.2 Mapeo al Rector (Sale / No sale)

**Sale (EXPLÍCITO — no-brechas y non-goals blindados por V1..V8):**
- Reescribir agregados R2/R4 o columnas Q2 de Recaudo (V2/V3: idénticos; tocarlos es regresión, no fix).
- Tocar `<f>` en CUALQUIER hoja (todas las brechas son literales de captura; las fórmulas ya encadenan bien).
- Cambiar la regla Q1/Q2 por columna en Recaudo (V3) o el sello G7/K7 y N3/C6/D6 (Plan 28 D-F: fuera de scope salvo H1 → proceso).
- Unidad D sin veredicto T0 de derivabilidad por columna (V7: prohibido inventar el desglose).
- Validación de contenido anticipada en preflight, espejo R4-motor, R4-por-empresa Q2, CF/DV, INTERVENTORIA-valores: follow-ups vivos de planes 21/25/27, intactos.
- Paquetes NuGet; commits (los hace el Ingeniero con `#commit`).

**No sale (entra, alcance congelado en 4 unidades):**
| Unidad | Brecha | Superficie de cambio |
|---|---|---|
| B — BCE D/E + arbitraje ASE5-julio | V4 (swap ASE1-4 ambos períodos) + V5 (ASE5-julio diverge de fuente) | `WorkbookLeafCellMapBalanceSc` (1 tabla) + `ExcelDataReaderWorkbookLeafInputReader.LeerBalanceSc` si T0 demuestra mislectura + tests. Si T0 no encuentra D7/E7-manual en insumos → cero código, solo docs |
| R — desglose-detalle R2/R4 | V6 (filas E3:R3-style y D3/D9-style en 0) | Reader (extraer matriz por componente) + 2 mapas (filas detalle por ASE/empresa) + writer (misma pasada atómica, guard anti-fórmula intacto) + tests. Sin insert/delete (las filas ya existen en ceros) |
| D — desglose DetRetri/DetValiRetri C..O | V7 (E9:O9 en 0) | **Condicionada:** solo las columnas cuyo origen workbook-interno demuestre T0; resto = SALE documentado |
| P — proceso + cosmética + comparador | V3 (Q1/Q2 acumulado), H1 (base por período), V8 (`0`-vs-vacío, VALIDACION), verificador reutilizable | Writer (1 micro-fix: no escribir `0` donde la plantilla trae vacío — solo si T0 confirma vacío-sin-fórmula) + docs proceso + `ComparadorSalidaVsManual` (BCL zip+XML) + tests |

### 0.3 Decisiones ejecutivas (aprobación = aceptar estas)

| ID | Decisión |
|---|---|
| D-A | **Causa múltiple, no única (refuta el encargo con evidencia):** V2/V3 son no-brechas (prohibido "arreglarlas"); V4/V5/V6/V7 son 4 causas independientes con 4 unidades. El "numerador final" ya cierra (DetRetri-D 5/5 vs R10): lo que se parifica es la trazabilidad visible, no el dinero. |
| D-B | **BCE header-manda:** D=SUBSIDIO←E-fuente, E=CONTRIBUCIÓN←F-fuente (texto del doc base + header + manual, 10/10 celdas ambos períodos). F=D+E idéntico antes/después → goldens y DetRetri-D intactos por construcción. Julio-ASE5: la app es fuente-fiel (V5); sin contra-evidencia T0 no se toca el reader. |
| D-C | **Unidad R sin geometría:** las filas de detalle existen en la plantilla en ceros (V1); el fix escribe valores, nunca inserta/borra ni reancla (derogación Plan 21 innecesaria aquí). El corrimiento `Especiales K=0` replica el patrón documentado de SALDOS (Plan 23 V1: template inserta `Especiales`, fuente no la trae). |
| D-D | **Unidad D con gate de derivabilidad:** T0 publica por columna C..O (DetRetri) y D..O (DetValiRetri) su origen (celda-fórmula interna exacta o "manual-externo"); solo las trazables entran. J (=BCE.F, V7) entra por construcción tras Unidad B. |
| D-E | **Verificador sin recálculo:** el comparador demuestra paridad por (a) literales ±0.5 y (b) **igualdad de texto de fórmula** (misma fórmula + mismos inputs ⇒ mismo resultado post-recalc; Plan 28 garantiza `fullCalcOnLoad`). Prohibido depender de cachés `<v>` (V8: stale por diseño) y de Excel/COM en el pipeline. |
| D-F | **Proceso sobre código para H1/V3:** cada quincena corre desde su base propia (Q1→Q2 encadenado en el mismo workbook si se quiere acumulado, o bases separadas si se compara por quincena); la guía vive en `project-context.md` + manual. Sin guardrail de bloqueo (doctrina Plan 28 D-G). |

---

## 1. PROPOSE

### 1.1 Intent

Que la salida de la app sea indistinguible del archivo manual en todo lo que el workbook puede derivar de los insumos (BCE por header, detalle R2/R4 por componente desde la fuente, DetRetri trazable, ceros cosméticos), con cada afirmación respaldada por T0 en disco, el dinero (DetRetri-D vs R10 5/5) intacto, y un comparador BCL que lo demuestre en cada corrida futura sin Excel.

### 1.2 In Scope

- T0 bloqueante (§4, T0a..T0g) que cierra H1/H2/H3 y arbitra V5/V7/VALIDACION antes de escribir una línea de producción. [CERRADO 2026-10-07: veredictos en `plans/29-T0-Evidencia-A.md` + `-B.md` — T0a DIVERGENCIA-DEL-MANUAL, T0b RECORTE→ampliado por decisión del Ingeniero a mapeo por encabezado, T0d 12/12+11/12 trazable (SALE: DetValiRetri J9), T0e micro-fix no aplica, T0f F10-poblar, T0g H1/H3 documentados]
- Unidad B: swap BCE + veredicto ASE5-julio + tests por ASE contra `TOTAL GENERAL` de la fuente.
- Unidad R: matriz de detalle R2 (filas Vlr Servicio/Intereses/Total por empresa-bloque, cols E..R) + detalle R4 (D3/D9-style), con tabla de filas congelada por T0. **AMPLIADA (ratificada 2026-10-07): mapeo por encabezado/label en vez de corrimiento fijo — cubre también la malla derivada de agosto (LIME/BOGOTA filas + Especiales en K).**
- Unidad D: solo columnas trazables (gate D-D) + J por arrastre de B.
- Unidad P: ~~micro-fix `0`-vs-vacío~~ [T0e: no aplica] + guía Q1→Q2 y base-por-período, comparador `ComparadorSalidaVsManual` + test de regresión julio+agosto + **poblar `'Valida - Control Recaudo'!F10`** (T0f: el `C15=False` es REAL por este literal faltante).
- Docs: `project-context.md` (doctrina paridad + tabla de causas V2..V7) + manual (proceso Q1→Q2).

### 1.3 Non-Goals

Ver §0.2 "Sale". En particular: NO se tocan agregados que ya cierran (V2/V3); NO se inventa desglose sin fuente (V6 tiene fuente; V7 parcial); NO se usa Excel/COM en pipeline ni tests; NO se reabre Plan 28 (sello/metadatos) salvo documentar H1.

### 1.4 Riesgos

| ID | Riesgo | Mitigación |
|---|---|---|
| R-MANUAL-MANDA | El manual de julio-ASE5 resulta ser el correcto (otra fuente legítima) y la app la divergente | T0a arbitra ANTES: si aparece la fuente, el fix va al reader, no al mapa; si no aparece, D-B blinda no tocar |
| R-DESGLOSE-DERIVA | La malla de detalle R2/R4 varía por período como varió R1 (Plan 21) y el mapa fijo se rompe en el próximo período | T0b congela la tabla por ASE con invariantes de cierre; si la forma varía entre julio/agosto → la unidad se recorta a "agregados + documentar" (fail-fast honesto, sin espejo nuevo en este plan) |
| R-CACHE | Comparar cachés `<v>` reintroduce el falso "vacío" (V2) | D-E: el comparador IGNORA `<v>` donde hay `<f>`; goldens y tests asertan literales + texto-fórmula |
| R-SOBREDISENO | Convertir el comparador en un mini-Excel (evaluador de fórmulas) | Alcance congelado: literales ±0.5 + texto-fórmula idéntico + reporte de divergencias; sin evaluar ni una fórmula |
| R-FALSO-POSITIVO | El comparador marca Q1-columnas o metadatos de período como brechas (V3/H1) | Lista de exclusiones declarada y versionada (columnas de otra quincena, N3/C6/D6/nombres de hoja salvo proceso P) |

---

## 2. DESIGN

### 2.1 Enfoque: escribir lo que falta, demostrar lo que cierra, documentar lo que es proceso

**Unidad B (2 líneas + arbitraje).** `EditablesPorAse` pasa de `(Contribucion=D, Subsidio=E)` a `(Subsidio=D, Contribucion=E)` —o se permutan las dos llamadas de `EscribirCeldasBalanceSc` (L1239-1240); T0a fija cuál para no romper el nombre del log `BCE.Contribucion.ASE{n}`. El reader (`ColumnaSubsidioFuente=4/Contribucion=5`) ya es header-correcto: solo el destino estaba invertido. Test: por ASE y período, D=SUBSIDIO=E-fuente-TOTAL-GENERAL y E=CONTRIBUCIÓN=F-fuente (10/10 julio + 10/10 agosto); F invariante ⇒ goldens ±0.5 y DetRetri-D intactos por construcción. Julio-ASE5: si T0a encuentra la fuente del manual → el fix va al reader (misma tabla, otro origen); si no → sin código.

**Unidad R (lectura + mapa + escritura, sin geometría).** El reader nuevo (`LeerDetalleR2/R4` o extensión del leaf-reader; T0b fija si convive con `LeerR2/LeerR4` del `IRecaudoReader` o los extiende) extrae por bloque-empresa la matriz Vlr Servicio/Vlr Intereses × componentes E..P-fuente; el mapa declara destino E..R con `Especiales K=0` + corrimiento (TCS→L, TLU→M, TBL→N, TRT→O, CCSA-NoAprov→P, DebCred→Q/R); el writer escribe en la misma pasada atómica tras los agregados (guard anti-fórmula existente: si un destino-detalle fuera fórmula → `ERR-PLANTILLA`, nunca sobrescritura). ASE4: columna `Especiales` ausente en fuente → 0 (invariante). R4-detalle: misma mecánica sobre las filas D3/D9-style (tabla T0b). Cero insert/delete: D-C.

**Unidad D (trazabilidad, no cálculo).** Por columna trazable, el writer copia el valor workbook-interno ya recalculable (p. ej. J←BCE.F del mismo workbook; E..I/K..O←celdas `REMUNERACION_*`/`GERENTES_*` que T0d identifique 10/10) como **literal redondeado** en la misma pasada (el manual los deja literales, V7). Columna sin origen 10/10 → SALE documentado con su nombre. DetValiRetri igual (origen probable `VALIDACION_*`/R1; T0d lo fija o lo saca).

**Unidad P (proceso + 1 micro-fix + comparador).** Micro-fix: en `EscribirValorNumerico` (o en `EscribirCeldasBanco` L1199-1211) no escribir cuando el valor es 0 **y** la plantilla trae la celda vacía-sin-fórmula (T0e confirma celda por celda; el guard anti-fórmula ya distingue fórmula/vacío). Comparador (`Remuneracion.Infrastructure`, BCL puro `System.IO.Compression` + `System.Xml`, sin OpenXML/ExcelDataReader): por hoja×celda con contenido en alguno de los dos archivos — (i) literal-vs-literal: |Δ|≤0.5 (numéricos) o igualdad (texto); (ii) fórmula: igualdad de texto normalizado (espacios fuera); (iii) literal-vs-fórmula o presencia unilateral fuera de exclusiones: divergencia reportada con hoja+celda+ambos valores. Exclusiones versionadas (R-FALSO-POSITIVO). El test de regresión corre el flujo 5-ASE Q2 julio y agosto desde la plantilla en ceros (insumos reales) y compara contra app-actual≈manual salvo exclusiones + brechas declaradas por unidad.

### 2.2 Alternativas y trade-offs

| Eje | Elegida | Descartada |
|---|---|---|
| BCE: permutar destino | 2 líneas, F invariante, 0 riesgo aritmético | Reinterpretar la fuente (columnas E/F ya correctas; V5 lo prueba) |
| Julio-ASE5: arbitrar antes | Evita "arreglar" la app hacia un manual sin fuente | Asumir manual-manda (rompería la fidelidad-fuente probada V5) |
| R: mapa fijo de detalle | Filas ya existen (V1); julio+agosto dirán si la malla es estable (T0b) | Espejo-con-mutación estilo Plan 21 (innecesario: D-C; prohibido expandir sin evidencia R-DESGOSE-DERIVA) |
| D: solo trazable | No inventa ni un número (V7) | Rellenar C..O desde el motor C# (duplicaría la cadena de fórmulas fuera del workbook: doble fuente de verdad) |
| Comparador sin evaluar | Demuestra paridad sin motor de cálculo (D-E) | Recalcular con Excel/COM (frágil, fuera del pipeline, ya descartado por el proyecto) |

### 2.3 Contratos afectados (resumen)

| Contrato | Cambio |
|---|---|
| `WorkbookLeafCellMapBalanceSc.EditablesPorAse` (+ `ObtenerEditables`) | Permuta D/E (Unidad B); comentarios T0-0.2 corregidos con veredicto V4 |
| `ExcelDataReaderWorkbookLeafInputReader` (detalle R2/R4; `LeerBalanceSc` solo si T0a) | Métodos nuevos de detalle + modelos `DetalleR2AseInputs/DetalleR4AseInputs` (Core puros); `LeerBalanceSc` intacto salvo veredicto |
| `WorkbookLeafInputs` (+ `R1/R2/R4` o nuevos miembros) | Matrices de detalle por ASE (diccionario celda→valor, mismo patrón `Celdas`) |
| `OpenXmlPlantillaWriter` (detalle + DetRetri-C..O trazable + micro-fix `0`-vacío) | Escrituras nuevas en la misma pasada atómica; guard anti-fórmula intacto como red |
| `ComparadorSalidaVsManual` (nuevo, Infrastructure BCL) + `Remuneracion.IntegrationTests` | Comparador + regresión julio/agosto; suite 345/345 como red ciega |
| `ProcesadorPeriodo/Remuneracion`, `IValidador`, R10-oráculo, preflight, espejo, sello | Sin cambios |
| `.opencode/project-context.md` + manual | Doctrina paridad (tabla V2..V7) + proceso Q1→Q2/base-por-período + exclusiones del comparador |

### 2.4 Architecture Validation Certificate

| Principio | Veredicto |
|---|---|
| **SRP** | ✅ Una razón de cambio por costura: mapa BCE (destino), reader-detalle (lectura), writer (tres escrituras nuevas independientes), comparador (reporte). Ninguna rama `if julio/agosto`. |
| **OCP** | ✅ El detalle nuevo entra por tabla (mapa celda→valor) + reader por encabezado, no por direcciones congeladas por período; el comparador absorbe futuras hojas por exclusión versionada, no por `if`. |
| **DIP** | ✅ Cambios detrás de `IWorkbookLeafInputReader`/`IPlantillaWriter`; el comparador depende de BCL, no de Excel; UI/CLI no conocen la regla. |
| **Best practices** | ✅ Fail-fast con archivo+celda (guard anti-fórmula como red, no adorno); quincena por dominio (V3 intacto); fórmulas nunca tocadas; tolerancia ±0.5; mensaje administrativo + detalle al log; sin Excel en pipeline. |
| **Performance** | ✅ BCE O(1); detalle = una pasada de lectura + N escrituras (N = celdas-detalle/ASE, acotado por T0b); comparador O(celdas-con-contenido) sobre ZIP, una vez por regresión. |

---

## 3. SPEC

### 3.1 Requirements

| ID | Requisito | Grupo | Criterio de aceptación |
|---|---|---|---|
| R-T0 | T0a..T0g cerrados con bytes/celdas citados; sin NEEDS_CONTEXT o con recorte explícito (D-C/D-D) | T0 | §4 verificable celda por celda |
| R-B-1 | BCE D=SUBSIDIO (E-fuente), E=CONTRIBUCIÓN (F-fuente) en 5 ASE × {julio, agosto} | Unidad B | 20/20 literales ±0.5 vs fuente `TOTAL GENERAL`; F=D+E invariante vs app-actual |
| R-B-2 | Julio-ASE5 arbitrado: fuente del manual encontrada → fix en reader; no encontrada → sin código + nota documentada | Unidad B | T0a veredicto escrito; app D7/E7 siguen = fuente R39 salvo contra-evidencia |
| R-B-3 | Goldens Q1+Q2 ±0.5 + DetRetri 5/5 vs R10 (julio y agosto) intactos | Transversal | Capa A + `Regresion2026082Tests` verdes; suite ≥345/345 |
| R-R-1 | Detalle R2 escrito (filas Vlr Servicio/Intereses/Total × componentes, `Especiales`=0, corrimiento T0b) 5 ASE × {julio, agosto} | Unidad R | Literales ±0.5 vs fuente (E..P→E..R) y vs manual E3:R3-style; agregados E43/D28 intactos |
| R-R-2 | Detalle R4 escrito (tabla T0b) con misma regla | Unidad R | ±0.5 vs fuente y vs manual D3/D9-style; D73/D66 intactos |
| R-R-3 | Ningún destino-detalle es fórmula (si lo fuera → `ERR-PLANTILLA`, no escritura) | Unidad R | Guard verde; test de celda-fórmula en copia temp |
| R-D-1 | DetRetri/DetValiRetri: solo columnas con origen 10/10 demostrado (J por arrastre B + las que T0d fije), como literal redondeado | Unidad D | ±0.5 vs manual por columna incluida; columnas excluidas listadas con motivo |
| R-P-1 | Micro-fix `0`-vs-vacío solo donde T0e confirme vacío-sin-fórmula en plantilla | Unidad P | `REPORTE RECAUDO x BANCO` D13:F14-style: vacío donde el manual vacía; valores ≠0 intactos |
| R-P-2 | Comparador verde en julio y agosto (salida regenerada desde ceros vs manual, con exclusiones versionadas) | Unidad P | Literales ±0.5 + texto-fórmula idéntico fuera de exclusiones; divergencias residuales = solo las declaradas por unidad pendiente |
| R-P-3 | VALIDACION_* post-paridad en True (o divergencia explicada celda por celda) | Unidad P | T0f + aserto del comparador sobre C15-style tras B+R+D |
| R-R-4 | Build 0 warnings; suite completa sin regresión (base 345/345) | Transversal | `dotnet build …slnx` + `dotnet test` en cada WU |

### 3.2 Scenarios (Given/When/Then)

- **S1 (BCE):** Given fuentes balance julio+agosto, When flujo 5-ASE Q2, Then BCE D3:E7 = (E,F)-fuente-`TOTAL GENERAL` por ASE y F idéntico al app-actual (swap invisible aguas abajo, visible en D/E).
- **S2 (ASE5-julio):** Given T0a sin fuente para D7/E7-manual, When se implementa B, Then D7/E7 siguen fuente-fieles y el plan lo declara divergencia-del-manual, no defecto.
- **S3 (detalle R2):** Given fuente R2 ASE1-julio, When flujo Q2, Then `Rem. Anticipos R2` E3=104754634.94/F3=6975298.69/…/K3=0/L3=8065770.61 (±0.5 vs fuente y manual).
- **S4 (detalle R4):** Given fuente R4 ASE1-julio, When flujo Q2, Then D3=-15794348.21/D9=-762167.27 (±0.5) con D15/P15 intactos.
- **S5 (DetRetri trazable):** Given T0d con J←BCE.F, When flujo Q2, Then DetRetri J9=-3589454665 (julio) como literal; columnas no-trazables en 0 con motivo documentado.
- **S6 (comparador):** Given salida regenerada desde ceros + manual, When `ComparadorSalidaVsManual`, Then cero divergencias fuera de exclusiones + brechas declaradas; informe nombra hoja+celda+ambos valores.
- **S7 (proceso):** Given manual actualizado, When un operador prepara Q2, Then encuentra "Q1→Q2 encadenado o bases separadas por quincena" + por qué D3: E3 en 0 es correcto en una corrida Q2 aislada (V3).
- **S8 (cero-geometría):** Given el diff, When se audita, Then ningún `<f>`, ningún mapa vigente y ningún agregado cambió salvo lo declarado (B + detalle + C..O-trazable + micro-fix).

---

## 4. TASKS

| ID | Tarea | Grupo | Depende de | Verificación |
|---|---|---|---|---|
| T0 | **Evidencia T0 OBLIGATORIA (Fase 0, NO es PR):** (a) arbitrar julio-ASE5: buscar -1138650714.67/+1055691691.01 en TODOS los insumos ASE5-julio (xlsx por zip+XML incl. `RecaudosReversados`, `ReportePagosxBanco`, conciliaciones; PDFs por texto extraíble si es barato, si no se declara límite) → veredicto fuente-del-manual o divergencia-del-manual; (b) congelar tabla de detalle R2/R4 por ASE (filas-fuente ↔ celdas-destino, corrimiento Especiales, invariantes de cierre) en julio+agosto; si la malla varía → recorte R-DESGOSE-DERIVA; (c) localizar H2 (GERENTES E9): R1!G63 app vs manual vs fuente + cadena E104→GERENTES; (d) origen por columna DetRetri C..O / DetValiRetri D..O (celda-fórmula interna exacta 10/10 o "manual-externo"); (e) `REPORTE BANCO` vacío-vs-`0` celda por celda en plantilla + manual; (f) evaluar VALIDACION_* por texto-fórmula + literales (método D-E) e identificar la(s) celda(s) que la ponen en False post-recalc; (g) cerrar H1 (metadatos agosto) y H3 (log D/E) al símbolo | Fase 0 | — | Veredictos T0a..T0g escritos con celdas citadas; sin cierre no empieza T1 |
| T1 | Unidad B: permuta destino BCE (mapa o llamadas L1239-1240, según T0a) + corrección comentarios T0-0.2 + tests 20/20 vs fuente + F-invariante | Fix-B | T0 | R-B-1/R-B-2/R-B-3; build 0 warnings |
| T2 | Unidad R-lectura: detalle R2/R4 en reader + modelos Core (tabla T0b; `Especiales`=0; ASE4 sin columna → 0) | Fix-R | T0 | R-R-1/R-R-2 (lectura); goldens intactos |
| T3 | Unidad R-escritura: mapas destino-detalle + escritura en pasada atómica + guard `ERR-PLANTILLA` + tests ±0.5 vs fuente y manual (julio+agosto) | Fix-R | T2 | R-R-1/R-R-2/R-R-3; S3/S4; suite ≥345/345 |
| T4 | Unidad D: escritura C..O trazables (gate T0d; J por arrastre B) como literal redondeado + lista de excluidas con motivo + tests | Fix-D | T0, T1 | R-D-1; S5; DetRetri-D 5/5 intacto |
| T5 | Unidad P: micro-fix `0`-vacío (gate T0e) + comparador BCL + regresión julio/agosto desde ceros + VALIDACION en True (R-P-3) | Fix-P | T0, T1..T4 | R-P-1/R-P-2/R-P-3; S6; suite ≥345/345 |
| T6 | Docs: `project-context.md` (doctrina paridad + tabla V2..V7 + exclusiones) + manual (proceso Q1→Q2, base-por-período, V3/H1) | Transversal | T1..T5 | R-R-4; S7/S8 citados; diff acotado |

**Orden sugerido:** T0 (bloqueante) → T1 → T2 → T3 → T4 → T5 → T6. T1 puede shippear sola (WU-1, 2 líneas, riesgo nulo); T2+T3 (WU-2); T4 (WU-3, solo si T0d da trazables — si no, se cierra como SALE documentado sin código); T5+T6 (WU-4). Build 0 warnings antes de marcar cada tarea completa. Sin commits (los hace el Ingeniero).

---

## 5. Resumen ejecutivo, riesgos y decisiones

- **Resumen:** el disco refuta la "única causa": los agregados R2/R4 y las columnas Q2 de Recaudo ya son idénticos (V2/V3 — el encargo comparó direcciones Q1 y cachés stale), y el dinero final ya cierra (DetRetri-D 5/5 vs R10). La brecha real son cuatro causas: (1) BCE D/E invertido (más julio-ASE5 donde la app es fuente-fiel y el manual diverge), (2) detalle por componente R2/R4 nunca escrito aunque la fuente lo trae fila a fila, (3) desglose DetRetri C..O nunca escrito y parcialmente sin origen identificado, (4) proceso (Q1/Q2 acumulado, base por período, `0`-vs-vacío). El plan las cierra una a una sin tocar una sola fórmula, con T0 que puede recortar (D, ASE5) pero nunca expandir.
- **Riesgo principal:** R-MANUAL-MANDA (T0a arbitra primero) + R-DESGOSE-DERIVA (T0b congela o recorta) + R-CACHE (D-E prohíbe cachés como evidencia). Contenidos por gates, no por confianza.
- **Decisión para el Ingeniero:** ratificar D-A..D-F (§0.3). **Preguntas con recomendación (ninguna bloqueante):** (1) D-B header-manda en BCE — recomendado: sí (10/10 + header + manual; F invariante); (2) D-D gate de Unidad D — recomendado: solo trazables, resto SALE; (3) julio-ASE5 — recomendado: salvo que T0a encuentre la fuente del manual, la app queda fuente-fiel y se documenta la divergencia.

### DoD (Definition of Done)

1. T0a..T0g publicados con celdas citadas (bloquea todo lo demás).
2. `dotnet build Remuneracion.WinForms/Remuneracion.WinForms.slnx` con 0 warnings.
3. Suite completa verde sin regresión (base 345/345) + goldens Capa A Q1+Q2 ±0.5 + DetRetri 5/5 vs R10 (julio y agosto).
4. BCE 20/20 vs fuente + F invariante (S1/S2); detalle R2/R4 ±0.5 vs fuente y manual (S3/S4); DetRetri trazable ±0.5 (S5).
5. Comparador verde en julio+agosto salvo exclusiones + brechas declaradas (S6); VALIDACION en True o explicada (R-P-3).
6. `project-context.md` + manual actualizados (S7); cero `<f>` tocados (S8); sin datos inventados; sin commits del agente; CRLF; sin emojis.

### Follow-ups explícitos (fuera de este plan)

- Espejo R4-motor (Plan 21), R4-por-empresa Q2 (HU-20/G2-D1), validación de nombres de hoja en preflight (Plan 27), CF/DV del reanclaje (W-4), INTERVENTORIA-valores (HU-16).
- Si T0b demuestra deriva de malla de detalle entre períodos, el espejo-detalle va en su propia HU con su propio T0 (no se improvisa aquí).
- Si T0a encuentra fuente legítima para julio-ASE5-manual, el fix del reader va acotado a ese origen con su test.
