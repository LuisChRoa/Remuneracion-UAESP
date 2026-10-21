---
titulo: Plan 49 - Cerrar falso verde gate 2.7 (gate absoluto anti-cancelacion)
estado: CERRADO 2026-10-19 (Clarification Gate §7 ratificado 5/5; desviacion W-1 del operando ratificada; H1/H2/H3 ejecutados; code-review APROBADO CON OBSERVACIONES — ver §9.6)
fecha: 2026-10-19
numero: 49
planes_ocupados: 01..48 ocupados; este es el 49
padre: Docs/T0-EXTEMP-Ase5-Plan49.md (T0 read-only EJECUTADO y verificado 2026-10-19; fuente unica de verdad del diagnostico)
madre_doctrina: plans/40 - validacion-total-evalua-sella.md (W-1 evalua+sella, E1-1..E1-5) + plans/42 - divergencia-declarada-reconocimiento-explicito.md (clasificacion, ultima puerta) + plans/48 - adopcion-plantilla-v2-paridad-manual-julio-2.md (frente que dejo el falso verde)
gate_entrada:
  t0_previo: Docs/T0-EXTEMP-Ase5-Plan49.md verificado (Q1 causa raiz, Q2 falso verde julio-2, Q3 barrido 77 celdas); no se re-investiga
  plantilla_v1: NO SE TOCA (historica; SHA256 `3086AC317595C2F8E84ED24B1C5EE3E52565DD5193394E3062878E7FA49C5214` intacto antes/despues)
  plantilla_v2: NO SE TOCA (SHA256 `FF5CF7A819110A1E1E2CA7BDAD553DD7A25141ADD472B1904BD496F082642632` intacto antes/despues; la correccion de plantilla depende de una v3 de la propietaria, fuera de este plan)
  alcance_aprobado: A1 (gate absoluto) + A2 (test de mutacion) + A3 (doctrina y reclasificacion); no se corrige el EXTEMP ni el R2 (se detectan y declaran)
  suite_base: p48-h3-suite.trx = 1059 Total / 1021 Passed / 38 Failed / 0 Skipped; referencia o43-final-failed.txt (set-diff 0 altas / 1 baja heredada)
  sin_commits: NUNCA `git commit`/`push`
---

# Plan 49 - Cerrar falso verde gate 2.7 (gate absoluto anti-cancelacion)

> SDD completo PROPOSE -> DESIGN -> SPEC -> TASKS. Fuente unica de verdad para la implementacion una vez aprobado ESTE documento. No se implementa sin aprobacion del Ingeniero (punto de bloqueo: sin su OK no se abre H1).
> Lectura obligatoria previa: `Docs/T0-EXTEMP-Ase5-Plan49.md` (completo; es la fuente de verdad) + `.opencode/project-context.md` (doctrinas Planes 21/25/26/27/28/29/30/31/33/34/39/40/41/42/45/46/47/48, invariantes, allow-list Phase 0 de 68 celdas) + `plans/48` (frente que dejo el falso verde) + `plans/46`/`plans/47` (genealogia R1/R2) + `Remuneracion.Core/Services/ValidadorBasico.cs` (gates vigentes) + `Remuneracion.Infrastructure/Excel/ValidacionOracleReader.cs` + `Remuneracion.Infrastructure/Excel/ResolutorValorValidacion.cs`.
> Entorno: Windows, repo con espacios en la ruta; plantillas CERRADAS en Excel al correr tests (FileShare).
> Visual Design Intent: N/A (cero UI en este plan).

---

## 0. Contexto adoptado del T0 (no re-investigar; cada hecho cita bytes del T0)

### 0.1 Los tres numeros (§2.1 del T0)

| Celda (VALIDACION_TOTAL) | Nuestra salida (agosto-2) | Manual | Δ |
|---|---|---|---|
| `D34` (`D33='CONSOLIDADO_TOTAL RECAUDO'!D109`) | `76.856.113.817,23` | `76.857.351.164,97` | `1.237.347,74` |
| `D34` lado ASE (`D109`) | `76.856.101.354,97` | `76.857.351.164,97` | `1.249.810,00` |
| `F34`, `O25` | Δ identico | — | `12.462,26` |

Descomposicion exacta del Δ (verificada por celdas, §2.2): EXTEMPORANEO ASE5 (`R1!F682`/`F702` = literal 0) pierde `1.249.810,00` en **ambos** lados → **cancela** (invisible); `R2` ASE4 grupo `A/EAB` (`R2!E307 = −12.462,26`) pierde solo en un lado → **rompe** (visible, ya reportado a la propietaria).

### 0.2 Cadena causal (§3 + §6.1 del T0)

```text
fuente R1 ASE5 (F20=1.666.506,54 · F48=−416.696,54)
  → ExcelDataReaderRecaudoReader (lectura OK)
  → RecaudoComponenteR1 / espejo (dominio OK: Extemp = 1.249.810; lo prueba DetRetri2026082!D13 identico al manual)
  → EscritorValoresSlotR1 (pega F629 y F660, filas 603..676; RangoDatosAse(5) = 603..676, EscritorValoresSlotR1.cs:89-97)
  → [FALTA: la formula-agregado F682 = F660+F629−L629 NO existe en la plantilla v2: <c r="F682" s="234"><v>0</v></c>]
  → CONSOLIDADO!D51 = 0 y REMUNERACION_OCCIDENTE!D51 = 0
  → D109 y D33 pierden 1.249.810 → el gate D34 cancela esa perdida (no la ve)
```

El defecto NO lo introdujo el Plan 48: la plantilla oficial legacy tambien trae `F682`/`F702` como literales (heredado, §3.4). La omision del escritor es **correcta por doctrina** (values-only, fuera de su contrato por diseño; `OpenXmlPlantillaWriter.cs:290-302` nunca toca la cola 677..710).

### 0.3 Falso verde de julio-2 (§4 del T0)

La corrida promovida de julio-2 tiene `CONSOLIDADO!D51 = 0` frente a `441.849` del manual (oraculo v2: `R1!F682 = F660+F629-L629`), y el gate no lo detecto. El T0 del Plan 48 **observo** las 9 formulas faltantes y las **congelo como esperadas** ("el delta es del manual, no de la base", `AdopcionPlantillaV2T0Plan48Tests.cs:421-450`); `GeometriaPlantillaEstandarizadaT0Tests.cs:230-231` lo elevo a contrato. Esa clasificacion es **incorrecta**: esas celdas son **anclas consumidoras** (`CONSOLIDADO!D51/E51/F51`, `DetRetri!D36`, `DetValiRetri!D28`, `REMUNERACION_OCCIDENTE!D51`, `REMUNERACION_ENEL!D51`), no cosmetica.

### 0.4 Alcance del defecto (§7 del T0)

Hojas: `Reporte Componentes R1` (77 celdas-ancla, filas 682/683/692/693/702/703) + `Rem. Anticipos R2` (3 filas ASE4 sin slot, otra causa). ASE: solo ASE5 en R1 (ASE1-4 tienen formula). Periodos: julio-2 (`441.849`) y agosto-2 (`1.249.810`); afecta a **cualquier** periodo con EXTEMPORANEO en ASE5. Lo NO afectado: lectura, calculo de dominio, colocacion interior R1, R4, AJUSTES, BCE, DetRetri-Q2 escrito.

### 0.5 Restricciones duras (vinculantes, del encargo)

- **NO** modificar `Docs/Plantilla estandarizada/v2/Formato Remuneracion 2026xx-x Total vf.xlsx` (SHA `FF5CF7…2632`) ni la v1 historica (SHA `3086AC31…C5214`). La correccion de plantilla depende de una **v3 de la propietaria** — fuera de este plan.
- **NO** debilitar tests: 0 `Skip`, 0 `Assert.True(true)`, 0 borrados. Los 5 T0 del §8.1 del T0 se **reclasifican como deuda declarada**, no se tocan sus asserts.
- **CERO** cambios de formula en plantillas; values-only.
- Build 0 warnings; CRLF; sin `git commit`/`push`.
- El plan **no corrige** el EXTEMP ni el R2 (son de la propietaria): los **detecta** y los **declara**.

---

## 1. PROPOSE

### 1.1 Intent

Cerrar la **clase de error** "igualdad entre dos derivaciones del mismo workbook" (T0 §6.3): una perdida comun a ambas derivaciones es invisible por construccion al gate 2.7 (`VALIDACION_TOTAL` `C15/D25/O25/D34/F34` en `ValidadorBasico.ValidarGatesValidacionesCruzadasPorAse`, `Remuneracion.Core/Services/ValidadorBasico.cs:472`). Se cierra con un gate **absoluto** que no puede cancelarse (un operando es el **dominio**, no otra celda del workbook), un **test de mutacion** que garantiza permanentemente la deteccion, y la **reclasificacion** de los 5 T0 que hoy asertan el defecto como contrato.

### 1.2 In-scope EXPLICITO (A1 + A2 + A3; prohibido ampliar)

| ID | Entregable | Criterio de salida |
|---|---|---|
| **A1-nucleo** | Gate absoluto: `CONSOLIDADO_TOTAL RECAUDO!D47:D51` resuelto desde lo escrito == `ConsolidadoAse.Extemp` de cada ASE en `resultado` (el dominio YA lo tiene: `1.249.810` en agosto-2). Falla con el defecto actual. | Gate bloqueante en `ValidadorBasico` con evidencia del reader; julio-2 y agosto-2 lo disparan con Δ exactos (`441.849` / `1.249.810`) |
| **A1-interior** | Gate interior: `R1!F682` (ASE5) vs suma resuelta de sus secciones (`F660+F629−L629`); `F702` analogo. Veredicto DESIGN §2.3: **SI, acotado** (solo columna F, solo anclas conocidas, mapa como DATO + T0 de anclas previo) | Nombra hoja R1 + celda ancla cuando el nucleo dispara; falla con el defecto actual |
| **A2** | Test de mutacion anti-cancelacion: workbook/snapshot sintetico donde `D33` y `D109` pierden el **mismo** componente → el gate nuevo **falla** (hoy el gate pasa). | Test verde que prueba el fallo; sin dependencia de disco (sintetico, determinista) |
| **A3** | Reclasificacion de los 5 T0 (§8.1 del T0) como **deuda declarada** + regla nueva en `project-context.md`: **ninguna celda referenciada por otra hoja puede congelarse como "delta del manual" sin demostrar que el valor no altera el consumidor.** | 5 archivos marcados con region/comentario citado; asserts intactos; doctrina escrita |

### 1.3 Non-goals / out (vinculante)

- Corregir el EXTEMPORANEO (formula `F682`/`F702`) o el R2 ASE4 (`E307`/`A/EAB`): son de la propietaria (v3). Este plan los detecta y declara; la via de correccion de plantilla (opciones A/B del T0 §8.2/§8.3) queda como follow-up fuera del plan.
- Tocar plantillas (v1/v2), sellar cachés nuevos, re-mapear R1/R2, tocar `DivergenciasDeclaradasCatalogo` (la entrada OCCIDENTE-ASE5 retirada sigue retirada; el Δ `12.462,26` del R2 NO se declara aqui: requiere su propio T0 de dueño con los 6 campos citados — ver riesgo R-ASIMETRIA).
- Q1: fuera de alcance como frente (sigue con v1 historica en e2e); el gate nuevo corre en Q1 sin exencion (§2.6) pero ningun comportamiento Q1 se cambia.
- Commits, UI, CLI, docs de usuario.

### 1.4 Riesgos con mitigacion

| ID | Riesgo | Mitigacion |
|---|---|---|
| R-CANCEL | El gate nuevo vuelve a comparar dos derivaciones del workbook y cancela | Un operando es el **dominio** (`ConsolidadoAse.Extemp`, `ConsolidadoAse.cs:29`); el otro es lo escrito resuelto. Por construccion no cancela (§2.2) |
| R-JULIO | Julio-2 promovida pasa a fallar y la suite queda ambigua | §2.5: rojo esperado DOCUMENTADO; altas autorizadas en lista cerrada; re-congelamiento con evidencia, nunca reverdecido silencioso |
| R-GEOM | El gate interior codifica geometria y se pudre con una v3 | Mapa de anclas como DATO citado + T0 de anclas previo; tercera forma = `ERR-PLANTILLA`; re-congelar por T0, nunca edicion silenciosa |
| R-ASIMETRIA | Arreglar/detectar solo el EXTEMP deja el gate en rojo por `12.462,26` y se lee como "el plan no cerro" | Declarado en SPEC: el exito de A1 es **disparar con el Δ exacto**, no un gate verde e2e; el `12.462,26` sigue su via propia con dueño UAESP |
| R-FALSO-POSITIVO | El gate absoluto falla en corridas sanas (Q1, ASE sin EXTEMP) | Compara contra el dominio del periodo (0 si no hay Aplicacion); sin rama por quincena; caso sano en A2 prueba el no-disparo |
| R-FIXTURE | Reclasificar los 5 T0 se lee como "debilitar tests" | Asserts intactos; solo region/comentario de deuda citada; grep de cierre lo verifica |
| R-OTRO-CANCEL | Otra perdida comun (otra hoja) sigue invisible | A2 cubre la clase (mutacion generica D33/D109); el barrido a otras hojas es follow-up con su T0 |

---

## 2. DESIGN

### 2.1 Donde vive el gate (decision)

**Decision: el JUICIO vive en `ValidadorBasico` (Core); la EVIDENCIA vive en `ValidacionOracleReader` (Infrastructure).**

| Candidata | Veredicto | Motivo |
|---|---|---|
| `ValidadorBasico.ValidarGatesValidacionesCruzadasPorAse` (`ValidadorBasico.cs:472`, invocada en `:266`) | **JUICIO — elegida** | Ya recibe `resultado` (dominio: `Consolidados[].Extemp`) y `snapshot` (workbook). Todos los gates cruzados HU-08..HU-12 viven aqui con agregacion de errores por lista (Plan 13 §2.8 / Plan 14 §9.7: moverlos a otro sitio cambiaria la semantica). `ClasificarDivergenciasDeclaradas` (`:308`, `IValidador.cs:87`, invocada en `ProcesadorPeriodo.cs:451`) reutiliza `Validar(...)` → el gate absoluto queda **automaticamente dentro de la clasificacion declarada/no-declarada** (ultima puerta, Plan 42 D-F) sin codigo extra. |
| `ValidacionOracleReader` (`Infrastructure/Excel/ValidacionOracleReader.cs:27`) | **EVIDENCIA — elegida (extension minima)** | El reader es oraculo read-only: NO juzga. Solo expone en el snapshot el EXTEMP resuelto desde lo escrito por ASE (via `ResolutorValorValidacion`, `:27`). No conoce `ResultadoRemuneracion` y no debe conocerlo (SRP del reader, doctrina W-1 Plan 40). |
| Juicio dentro del reader | **Descartada** | Violaria el SRP del reader y la semantica de agregacion del validador; ademas el reader no tiene el dominio contra el que comparar. |

**Precedente insuficiente (no re-litigar, solo situar):** `ValidadorBasico.cs:716-720` ya gatea `leaf.R1.F25` vs `consolidado.Extemp`, pero AMBOS operandos derivan de la **lectura de fuentes**, ninguno del **workbook generado** — por eso tampoco ve la perdida. A1 es cobertura genuinamente nueva: dominio vs workbook.

### 2.2 Gate absoluto A1-nucleo (spec de diseno)

- **Operando izquierdo (workbook):** `CONSOLIDADO_TOTAL RECAUDO!D{47+k}` resuelto desde lo escrito con `ResolutorValorValidacion` (texto-`<f>` + literales; **prohibido `<v>` stale**, doctrina Plan 28/29), expuesto por el reader en el snapshot (campo nuevo, p. ej. `ExtempResueltoPorAse`; el nombre exacto lo fija H1 citando `ValidacionCruzadaSnapshot.cs:130`). `k` = indice del ASE (ASE1→D47 … ASE5→D51).
- **Operando derecho (dominio):** `resultado.Consolidados.Single(c => c.Ase.Id).Extemp` (matcheo estricto por `Ase.Id`, nunca fallback).
- **Regla:** `|resuelto − dominio| > 0.5` → error bloqueante con trazabilidad §2.7. Sin rama por quincena (§2.6).
- **Comportamiento probado:** agosto-2 dispara en ASE5 con `resuelto=0` vs `dominio=1.249.810`; julio-2 dispara con `resuelto=0` vs `dominio=441.849`. Con la perdida actual, D34 sigue pasando (cancelacion intacta) y A1 falla: esa es exactamente la firma del falso verde, ahora visible.
- **Metodo nuevo:** `ValidarExtempAbsolutoPorAse(errores, snapshot, resultado)` (privado, junto al gate 2.7, `ValidadorBasico.cs:472-578`), invocado desde `ValidarGatesValidacionesCruzadasPorAse`. `ClasificarDivergenciasDeclaradas` lo hereda sin cambios (reutiliza `Validar`).

### 2.3 Gate interior A1-secundario (decision SI, acotada)

| Aspecto | Veredicto |
|---|---|
| ¿Conviene? | **SI, acotada a columna F + anclas conocidas.** El nucleo dice "el CONSOLIDADO no trae el dinero"; el interior dice **donde se perdio** (hoja R1, celda ancla). Sin el, el diagnostico futuro vuelve a exigir un T0 de bytes a mano. |
| Trade-off | Pro: localiza la hoja/celda; no depende de ningun oraculo externo (compara interior de la propia hoja); detectaria aunque el dominio estuviera mal. Contra: **codifica geometria** (`F660+F629−L629`); requiere T0 de anclas; acoplamiento a v2 (una v3 exige re-congelar); solo cubre columna F (G/H y demas columnas quedan fuera por diseño declarado). |
| Alcance | Solo ASE5, solo columna F: `F682 ?= F660+F629−L629`, `F702 ?= F659+F628−L628` (formulas del manual v2-julio citadas en T0 §3.4). ASE1-4 tienen formula en plantilla y el resolutor ya las evalua: no se gatean (cero geometria inventada donde la plantilla calcula sola). |
| Forma | Mapa como **DATO** (cero ramas por empresa/periodo, D-H), p. ej. `AnclasExtempR1`: `{ ASE5: [(F682, [F660, F629, −L629]), (F702, [F659, F628, −L628])] }` en Core (Models o Rules, H1 decide citando). Tercera forma (ancla ausente o con `<f>` inesperado) = `ERR-PLANTILLA` nombrando hoja+celda, nunca 0 silencioso ni fallback posicional. |
| Requisito previo | **T0 de anclas** (H1 primera tarea): verifica por bytes en la v2 que `F660/F629/L629/F659/F628/L628` existen y que `F682/F702` son literales; congela el mapa. Si aparece forma nueva → STOP, no rama. |

### 2.4 Test de mutacion A2 (diseno)

- **Nivel:** unitario contra `ValidadorBasico` con **snapshot + resultado sinteticos** (cero disco, determinista; evita la fragilidad de fixtures con rutas `%TEMP%` declarada en Plan 48 §9.7).
- **Caso 1 (el que hoy pasa y debe fallar):** snapshot sintetico con `D34` TRUE (ambos operandos reducidos en `X = 1.249.810`: simula la cancelacion) + `ExtempResueltoPorAse[5] = 0` + `resultado` con `Consolidados[ASE5].Extemp = 1.249.810`. Aserta: `Validar(...)` contiene el error A1 del ASE5 con el Δ exacto, y `ClasificarDivergenciasDeclaradas(...).HayNoDeclaradas == true` (falla en duro aunque la declarada este reconocida — Plan 42 D-F). Con el codigo actual este test **falla** (el gate no existe): es rojo→verde honesto.
- **Caso 2 (no falso positivo):** resuelto == dominio (±0.5) → cero errores A1. Protege contra un gate que falle siempre.
- **Archivo nuevo:** `Remuneracion.IntegrationTests/GateAbsolutoAntiCancelacionTests.cs` (cero produccion tocada por el test salvo el ctor con catalogo sintetico ya previsto en `ValidadorBasico.cs:33-36`; el fixture ejercita la implementacion real, patron Plan 46 review W-1).
- **Nota:** el caso 1 usa el booleano `D34=TRUE` entregado (snapshots sin evidencia conservan el gate del booleano, `ValidadorBasico.cs:572-577`) + el campo nuevo de evidencia A1: asi se prueba exactamente la configuracion "gate viejo verde + perdida real".

### 2.5 Impacto en la suite: que pasa con julio-2 (decision)

**Decision: rojo esperado DOCUMENTADO, nunca reverdecido.**

1. Con A1, toda corrida julio-2/agosto-2 sobre la v2 **falla honestamente** en el ASE5 (`D51 resuelto=0` vs dominio `441.849`/`1.249.810`). Ese rojo es **el exito del plan** (el falso verde deja de ser verde), no una regresion.
2. H2 inventaria (grep `2026072` e2e + goldens `D51`) los tests que hoy esperan verde sobre esas salidas. Cada uno se **re-congela con evidencia**: su expectativa pasa a exigir el fallo A1 con el Δ exacto (fail honesto citado), o queda documentado como rojo-de-deuda con la condicion de retiro (v3 de la propietaria). **Ningun test se toca para que vuelva a verde.**
3. Criterio de suite (§3-S0): set-diff vs base `p48-h3-suite.trx` (38 failed) con **altas = lista cerrada y citada** (solo las del A1 sobre julio-2/agosto-2 + los verdes nuevos de A2); **cero altas fuera de esa lista**; 0 `Skip`, 0 `Assert.True(true)`, 0 borrados. Si un test fuera de la lista se mueve → STOP con causa.

### 2.6 Q1 vs Q2 (decision)

**El gate corre en AMBAS quincenas, sin rama por quincena.** Compara el dominio **del periodo** contra lo escrito: en Q1 el `Extemp` del dominio con fuentes sin filas `Aplicacion` es el que el reader calcule (tolerancia ±0.5; 0 si no hay) — el gate lo compara igual. Eximir Q1 seria reintroducir la clase de error por la puerta de atras. Intacto: `AjustesSfT = 0` en Q1 (`ValidadorBasico.cs:142-153`); el e2e Q1 sigue con la **v1 historica** (doctrina Plan 48) hasta autorizacion separada; ningun comportamiento Q1 se cambia. Si el gate dispara en Q1 con un Δ real, ese rojo es honesto y bienvenido (no se apaga).

### 2.7 Trazabilidad (formato prescrito, en la linea del proyecto)

```text
[ERR-VALIDACION] ASE {n} · CONSOLIDADO_TOTAL RECAUDO · D{47+k}: EXTEMP resuelto desde lo escrito ({resuelto}) != Extemp del dominio ({esperado}). Diferencia={diff} > ±0.5. (Plan 49 A1: gate absoluto anti-cancelacion.)
[ERR-VALIDACION] ASE 5 · Reporte Componentes R1 · F682: ancla EXTEMP resuelta ({resuelto}) != suma de secciones F660+F629−L629 ({esperado}). Diferencia={diff} > ±0.5. (Plan 49 A1-interior.)
```

Cada error nombra **ASE + hoja + celda + esperado vs obtenido + diferencia**. Jerga de codigo prohibida en el texto (doctrina Plan 26 D-E); el codigo de error viaja en el prefijo `[ERR-VALIDACION]`.

### 2.8 Reclasificacion A3 (diseno)

En cada uno de los 5 archivos, envolver/senalar el assert citado con region + comentario de deuda (aserts **intactos**):

```csharp
#region DEUDA-DECLARADA Plan 49 (T0-EXTEMP-Ase5-Plan49 §4.3/§8.1)
// La celda F682/F702 es ancla consumidora (CONSOLIDADO!D51). El assert congela el
// defecto de la plantilla v2 (literal 0), NO un comportamiento esperado.
// Condicion de retiro: v3 de la propietaria con la formula restaurada + T0 nuevo.
#endregion
```

Archivos: `Remuneracion.IntegrationTests/GeometriaPlantillaEstandarizadaT0Tests.cs:104,230-231` · `Remuneracion.IntegrationTests/T0ResidualOccidenteAse5Tests.cs:30-48` · `Remuneracion.IntegrationTests/AdopcionPlantillaV2T0Plan48Tests.cs:446-450` · `Remuneracion.IntegrationTests/R1FuenteCanalConceptoT0Tests.cs:209-210` · `Remuneracion.IntegrationTests/DetRetriQ2Tests.cs:38-40`. Verificacion: grep de cierre prueba asserts byte-identicos (solo se anaade region/comentario).

### 2.9 Arquitectura del cambio (contract-first)

- **Core (puro):** `AnclasExtempR1` (dato citado) + `ValidarExtempAbsolutoPorAse` + mensaje §2.7. Cero ramas por periodo/empresa (D-H): el mapa es dato, el ASE es clave.
- **Infrastructure (lectura):** `ValidacionOracleReader` expone `ExtempResueltoPorAse` (+ evidencia interior `AnclaExtempR1Resuelta`) en el snapshot via `ResolutorValorValidacion`. Cero `<f>` tocada, cero sellado nuevo (el sellado W-1 del Plan 40 intacto).
- **Tests:** `GateAbsolutoAntiCancelacionTests` (nuevo, sintetico) + re-congelamiento H2 con evidencia + regiones A3.
- **Docs:** `project-context.md` (regla celda-referenciada + doctrina gate absoluto) + nota en `plans/48` (falso verde cerrado por este plan) sin reescribir historia.

---

## 3. SPEC (criterios numerados + gates no negociables)

### S0 — Gates transversales (toda T)

- G-BUILD: `dotnet build Remuneracion.WinForms/Remuneracion.WinForms.slnx` → **0 warnings / 0 errores**.
- G-SUITE: suite completa con trx + **set-diff vs linea base `p48-h3-suite.trx` (38 Failed / 1021 Passed / 1059 Total)**: altas == lista cerrada H2 (solo A1 sobre julio-2/agosto-2 + verdes A2); **rojo nuevo fuera de lista = STOP con causa**. 0 `Skip`, 0 `Assert.True(true)`, 0 tests borrados.
- G-SHA: v1 `3086AC31…C5214` + v2 `FF5CF7…2632` intactos antes/despues.
- G-PROC: cero commits; CRLF sin mezclar en lo tocado.
- G-D-H: cero literales/ramas por periodo o empresa en produccion (grep de cierre; el mapa de anclas es DATO citado).
- G-WORKBOOK: todo gate de negocio lee el workbook generado (texto-`<f>` + literales; prohibido `<v>` stale; `cmpnum -Tol 0.5` como oraculo numerico).
- G-FORMULA: `<f>` byte-identica en todo el diff (cero `new CellFormula` fuera del preservador Plan 33; cero `InsertRow`/`DeleteRow`).
- G-TRAZA: cada error nuevo cumple el formato §2.7 (ASE + hoja + celda + esperado vs obtenido).

### S1 — Criterios H1 (gate absoluto A1)

- S1.1: T0 de anclas previo verde (bytes v2: `F660/F629/L629/F659/F628/L628` existen; `F682/F702` literales 0) y mapa `AnclasExtempR1` congelado como dato citado.
- S1.2: `ValidarExtempAbsolutoPorAse` implementado y llamado desde `ValidarGatesValidacionesCruzadasPorAse`; `ClasificarDivergenciasDeclaradas` lo hereda sin cambios (prueba: clasificacion con catalogo sintetico).
- S1.3: reproduccion agosto-2: A1 dispara en ASE5 con `resuelto=0` vs `dominio=1.249.810` (Δ exacto); reproduccion julio-2: dispara con `441.849`.
- S1.4: corrida sana (resuelto == dominio ±0.5) → cero errores A1 (caso 2 de A2 en verde).
- S1.5: build 0 warnings + greps D-H + G-FORMULA + G-SHA.

### S2 — Criterios H2 (mutacion A2 + impacto julio-2)

- S2.1: `GateAbsolutoAntiCancelacionTests` verde: caso 1 falla-con-el-gate (rojo→verde honesto) + caso 2 no-falso-positivo.
- S2.2: inventario H2 de tests julio-2/agosto-2 e2e + goldens `D51` completo; cada afectado re-congelado con evidencia del fallo A1 (Δ exacto citado) o documentado como rojo-de-deuda con condicion de retiro (v3).
- S2.3: set-diff cumple S0 (altas == lista cerrada; cero fuera de lista) + Q1 intacto (gate corre sin exencion pero ningun comportamiento Q1 cambia) + SHAs intactos.

### S3 — Criterios H3 (docs/cierre)

- S3.1: doctrina en `.opencode/project-context.md` (gate absoluto + regla celda-referenciada + mapa de anclas); nota de cierre del falso verde en `plans/48` sin reescribir historia; desviaciones con fecha/motivo/evidencia.
- S3.2: diff de docs citado; cero codigo de produccion en H3.
- S3.3: code-review independiente (spec compliance + calidad): trazabilidad cada cambio → A1/A2/A3; ningun test debilitado; CRLF; SHAs.

---

## 4. TASKS (fases H1..H3 con puertas)

### H1/F1 — Gate absoluto A1 (puerta: aprobacion §7; sin OK no se abre)

1. H1.1 T0 de anclas (read-only, BCL ZIP+XML): verificar en la v2 `F660/F629/L629/F659/F628/L628` (existen, valores) + `F682/F702` (literales 0) + consumidores ya citados (T0 §5.2). Congelar mapa `AnclasExtempR1` como dato. Si forma nueva → STOP.
2. H1.2 Core: nuevo `ValidarExtempAbsolutoPorAse` en `Remuneracion.Core/Services/ValidadorBasico.cs` (junto a `:472`), llamado desde `ValidarGatesValidacionesCruzadasPorAse`; mensajes formato §2.7 (nucleo + interior). Incluye mapa `AnclasExtempR1` como dato citado (Models o Rules, con cita T0).
3. H1.3 Infrastructure (lectura): `ValidacionOracleReader` (`Remuneracion.Infrastructure/Excel/ValidacionOracleReader.cs:27`) expone `ExtempResueltoPorAse` (+ evidencia interior) en `ValidacionCruzadaSnapshot` (`Remuneracion.Core/Models/ValidacionCruzadaSnapshot.cs:130`) via `ResolutorValorValidacion` (`ResolutorValorValidacion.cs:27`). Cero juicio en el reader; cero `<f>` tocada.
4. H1.4 Reproduccion: agosto-2 dispara (`0` vs `1.249.810`) + julio-2 dispara (`0` vs `441.849`) con Δ exactos; corrida sana no dispara. Build 0 warnings + greps D-H/G-FORMULA + SHAs.
- **Puerta H1→H2:** S1.1..S1.5 verdes. Sin esto no se abre H2.

### H2/F2 — Mutacion A2 + impacto julio-2

1. H2.1 Nuevo `Remuneracion.IntegrationTests/GateAbsolutoAntiCancelacionTests.cs`: caso 1 (cancelacion simulada `D34=TRUE` + `Extemp` dominio `1.249.810` vs resuelto `0` → error A1 ASE5 + `HayNoDeclaradas == true`) + caso 2 (sano → cero errores A1). Sintetico, sin disco; ctor con catalogo sintetico (`ValidadorBasico.cs:33-36`).
2. H2.2 Inventario: grep tests e2e `2026072`/`2026082` + goldens `D51`/`441849`/`1249810`; tabla test → veredicto (re-congelar con evidencia A1 / rojo-de-deuda con retiro v3 / intacto con motivo).
3. H2.3 Re-congelamiento con evidencia citada por test (expectativa = fallo A1 honesto con Δ exacto); ningun test reverdecido. Suite trx + set-diff vs base (altas == lista cerrada) + Q1 intacto + SHAs.
4. H2.4 Code-review independiente (spec compliance + calidad).
- **Puerta H2→H3:** review APROBADO (o APROBADO-CON-OBSERVACIONES cerradas) + set-diff S0 + G-Q1 + SHAs.

### H3/F3 — Doctrina A3 + cierre

1. H3.1 Reclasificar los 5 T0 con region `DEUDA-DECLARADA Plan 49` (§2.8); asserts byte-identicos (grep de cierre: solo region/comentario anadido).
2. H3.2 Doctrina en `.opencode/project-context.md`: gate absoluto anti-cancelacion + regla celda-referenciada + mapa de anclas; nota de cierre en `plans/48` (§9.7 sin reescribir historia).
3. H3.3 `Docs/Instructivo-Capa-B.md` si aplica (que significa el error A1 para el administrativo: defecto de plantilla v2, dueño propietaria, retiro v3).
- **Puerta H3:** diff de docs citado; cero codigo de produccion en H3.

---

## 5. Riesgos (tabla consolidada + gates por fase)

Ver §1.4 (R-CANCEL, R-JULIO, R-GEOM, R-ASIMETRIA, R-FALSO-POSITIVO, R-FIXTURE, R-OTRO-CANCEL). Gates por fase:

- **Gate H0→H1 (aprobacion):** Clarification Gate §7 ratificado. Sin OK no se implementa.
- **Gate H1→H2:** S1 verde (T0 anclas + reproducciones con Δ exactos + build 0w + greps + SHAs).
- **Gate H2→H3:** set-diff S0 + lista cerrada de altas + Q1 intacto + SHAs + review aprobado.
- **Gate H3:** diff de docs citado; cero codigo de produccion en H3.

---

## 6. Architecture Validation Certificate

- **SOLID:** SRP (reader = evidencia read-only / validador = juicio con agregacion / mapa de anclas = dato citado; el test ejercita la implementacion real via ctor con catalogo); OCP (nueva ancla o periodo = fila citada en el dato + T0, cero ramas nuevas; D-H intacto); LSP/DIP (validador depende de `IValidador`/`ResultadoRemuneracion`/snapshot, no de concretos de Infrastructure; el reader no conoce el dominio); ISP (contratos `IValidador`/`IValidacionOracleReader` intactos: el campo nuevo del snapshot es aditivo).
- **Best Practices:** contract-first (H1.1 congela anclas por bytes antes del juicio); fail-fast accionable (`ERR-PLANTILLA` en tercera forma; `[ERR-VALIDACION]` con ASE+hoja+celda+esperado vs obtenido); gates que leen el workbook (texto-`<f>` + literales, nunca `<v>`); aceptacion con Δ exactos (`1.249.810` / `441.849`) + set-diff contra 38/999…(base 1059/1021/38/0) + doble SHA256; CRLF; sin commits.
- **Performance:** el gate es O(5 ASE) sobre valores ya resueltos por el reader (una lectura por celda ancla; sin re-lectura por celda, sin abrir workbooks extra). El test A2 es sintetico en memoria (sin disco, sin Excel/COM).

---

## 7. Clarification Gate (requiere aprobacion del Ingeniero; sin esto no se abre H1)

1. **Ubicacion y doble gate:** ratificar juicio en `ValidadorBasico` + evidencia en `ValidacionOracleReader` (§2.1) y el doble gate (nucleo obligatorio + interior acotado a F/ASE5, §2.3) frente a la alternativa "solo nucleo".
2. **Julio-2:** ratificar rojo esperado documentado + lista cerrada de altas (§2.5) frente a cualquier reverdecimiento.
3. **Q1:** ratificar que el gate corre en Q1 sin exencion y sin cambiar comportamiento Q1 (§2.6).
4. **Doctrina:** ratificar la regla celda-referenciada y la reclasificacion como deuda (aserts intactos) de los 5 T0 (§2.8).
5. **R2 asimetrico:** ratificar que el `12.462,26` NO se declara en este plan (requiere su T0 de dueño) y que el exito de A1 se mide por **disparar con el Δ exacto**, no por un gate e2e verde.
6. **Puerta:** sin 1–5 no se abre H1; sin H1 verde no hay H2; sin review no hay H3.

**RATIFICACION 2026-10-19 (Ingeniero) — APROBADO 5/5:** (1) doble gate: nucleo obligatorio + interior acotado a F/ASE5; (2) julio-2: rojo esperado documentado + lista cerrada de altas, prohibido reverdecer; (3) el gate corre en Q1 sin exencion; (4) regla celda-referenciada + reclasificacion de los 5 T0 como deuda declarada con asserts intactos; (5) el `12.462,26` del R2 queda fuera de este plan y el exito de A1 se mide por **disparar con el Δ exacto**. **H1 HABILITADO.**

**RATIFICACION W-1 (2026-10-19, Ingeniero) — OPERANDO DERECHO APROBADO:** se ratifica comparar contra `WorkbookLeafInputsR1.ExtemporaneoEsperadoPorAse` (agregado del dominio por firma) en lugar de `ConsolidadoAse.Extemp`, porque este ultimo esta mal alimentado por el lector legacy (primera fila `Mes`): julio-2 ASE5 `13.084.233.937,44` y agosto-2 ASE5 `-4.485,01`, frente a `441.849,00` / `1.249.810,00` reales. Evidencia: §9.2 y `p49-h2-setdiff.txt`. **Plan habilitado a cierre.**

---

## 8. Estado

**CERRADO 2026-10-19.** Clarification Gate §7 ratificado 5/5; desviacion W-1 del operando ratificada. H1/H2/H3 ejecutados; suite `p49-h2-suite.trx` = 1061/1012/49/0 (11 altas en la lista cerrada, 0 bajas); code-review **APROBADO CON OBSERVACIONES** (§9.6). La correccion de plantilla (EXTEMP ASE5 y R2 ASE4) queda FUERA: v3 de la propietaria.

---

## 9. EJECUCION H1/H2/H3-docs (2026-10-19)

### 9.1 H1.1 - T0 de anclas (read-only, BCL ZIP+XML) — VERDE

Bytes de la v2 (SHA `FF5CF7A8…2632`) confirmados:

- `Reporte Componentes R1!F660/F629/L629/F659/F628/L628`: **existen** como literales (`<v>0</v>`, sin `<f>`).
- `Reporte Componentes R1!F682/F702`: **literales `0` sin `<f>`** (defecto heredado).
- `CONSOLIDADO_TOTAL RECAUDO!D47:D51`: fórmulas `='Reporte Componentes R1'!F77/F256/F417/F570/F682`.
- El mismo patron se repite en la v1 histórica (`F682/F702` literales) — no se toca.

Mapa congelado como DATO citado: `Remuneracion.Core.Models.AnclasExtempR1` (ASE5 → `F682: F660+F629-L629`, `F702: F659+F628-L628`).

### 9.2 H1.2/H1.3 - implementacion (contract-first)

| Pieza | Archivo | Traza |
|---|---|---|
| Mapa de anclas (dato) | `Remuneracion.Core/Models/AnclasExtempR1.cs` | S1.1 / §2.3 |
| Evidencia en snapshot (`ExtempResuelto`, `AnclasExtempR1`) | `Remuneracion.Core/Models/ValidacionCruzadaSnapshot.cs` | S1.2 / D-A |
| Juicio A1-nucleo + A1-interior | `Remuneracion.Core/Services/ValidadorBasico.cs` (`ValidarExtempAbsolutoPorAse`, invocado desde `ValidarGatesValidacionesCruzadasPorAse`) | S1.2 / §2.2 / §2.7 |
| Evidencia read-only del workbook | `Remuneracion.Infrastructure/Excel/ValidacionOracleReader.cs` (`ResolverExtempConsolidado` / `ResolverAnclasExtempR1`) | S1.2 / D-A |
| Helper de existencia | `Remuneracion.Infrastructure/Excel/ResolutorValorValidacion.cs` (`ExisteCelda`) | S1.2 |
| Test de mutacion A2 | `Remuneracion.IntegrationTests/GateAbsolutoAntiCancelacionTests.cs` | S2.1 / §2.4 |

**Desviacion ratificable (H1, 2026-10-19) — operando derecho del nucleo:** el plan §2.2 nombra `ConsolidadoAse.Extemp`. En el codebase ese campo lo alimenta `ExcelDataReaderRecaudoReader.LeerR1` con la **primera fila `Mes`** de la fuente, NO el agregado EXTEMP: julio-2 ASE5 = `13.084.233.937,44` y agosto-2 ASE5 = `-4.485,01` (ASE1-4 tambien divergentes), frente a `441.849,00` / `1.249.810,00` del manual. Usarlo habria disparado **5/5 ASE con Δ falsos** (rojo fuera de la lista cerrada). El operando que produce EXACTAMENTE los Δ del plan es el **agregado del dominio por firma** `WorkbookLeafInputsR1.ExtemporaneoEsperadoPorAse`, que coincide al centavo con `CONSOLIDADO!D47:D51` del manual en julio-2 (`49580 / 56190070 / 1307680 / 4343101 / 441849`) y agosto-2 (`63570 / 221630 / 0 / 4030410 / 1249810`). Se implemento ese operando; `ConsolidadoAse.Extemp` queda **intacto** (fuera de alcance: no se corrige el EXTEMP). Evidencia en `Remuneracion.IntegrationTests/TestResults/p49-h2-setdiff.txt` §Delta.

### 9.3 H1.4 - reproduccion con Δ exactos (S1.3)

- **julio-2:** `D51 = 0` vs dominio `441.849,00` (Δ `441.849,00`); `F682 = 0` vs `441.849,00`; `F702 = 0` vs `441.849,00`. ASE1-4 sin error.
- **agosto-2:** `D51 = 0` vs dominio `1.249.810,00` (Δ `1.249.810,00`); `F682 = 0` vs `1.249.810,00`; `F702 = 0` vs `1.249.810,00`. ASE1-4 exactos.
- **corrida sana (A2 caso 2):** resuelto == dominio ±0.5 → cero errores A1.

### 9.4 H2 - suite y set-diff (S0/S2.3)

- `p49-h2-suite.trx`: **1061 Total / 1012 Passed / 49 Failed / 0 Skipped**.
- Base `p48-h3-suite.trx`: 1059 / 1021 / 38 / 0. **Altas: 11; bajas: 0.**
- Las 11 altas son **solo A1 sobre julio-2** (0 fuera de la lista cerrada); detalle test→evidencia en `Remuneracion.IntegrationTests/TestResults/p49-h2-setdiff.txt`.
- 0 `Skip`, 0 `Assert.True(true)`, 0 tests borrados, 0 reverdecidos. **No se re-congelo ningun test**: las 11 quedan como **rojo-de-deuda declarado** con condicion de retiro **v3 de la propietaria** (restaurar `F682/F702` + T0). Re-congelarlos habria vaciado su proposito (clasificacion declarada, observabilidad, UI/CLI en dos pasadas) y §2.5 permite la segunda rama.
- Q1: el gate corre sin exencion; ningun comportamiento Q1 cambio (los rojos Q1 preexistentes siguen por su causa `ERR-SIN-SLOT`).
- SHAs v1 `3086AC31…C5214` y v2 `FF5CF7…2632` intactos antes/despues.

### 9.5 H3 - doctrina y cierre documental

- `.opencode/project-context.md`: doctrina "gate absoluto anti-cancelacion + regla celda-referenciada" (incluye la desviacion del operando y el estado de suite).
- `plans/48` §9.7: addendum 2026-10-19 que cierra el falso verde (sin reescribir la historia).
- `Docs/Instructivo-Capa-B.md`: nota A1 para el administrativo (que significa, dueno, retiro v3; corrige el punto 4 de la seccion v2 que llamaba "delta del manual" a las anclas).

### 9.6 Code-review independiente (H2.4/S3.3) y cierre — APROBADO CON OBSERVACIONES

- Veredicto: **APROBADO CON OBSERVACIONES**, 0 bloqueantes (`B-x`).
- Verificado por el revisor: build 0 warnings; desviacion del operando demostrada con codigo y bytes (`ExcelDataReaderRecaudoReader.LeerR1:42-50` → `CalculoRemuneracion.cs:39`); A2 2/2 verde re-ejecutado; 0 `Skip`/`Assert.True(true)`/borrados; los 5 T0 con asserts byte-identicos; SHAs v1/v2 intactos; CRLF; sin commits.
- **W-1 (ratificado 2026-10-19):** operando derecho = `WorkbookLeafInputsR1.ExtemporaneoEsperadoPorAse` (ver §7).
- **W-2 (limitacion declarada):** el revisor no pudo re-correr la suite completa (>20 min; fixture de `%TEMP%` no reproducible). Se acepta el trx `p49-h2-suite.trx` + set-diff declarado.
- Sugerencias no bloqueantes: S-1 (`SingleOrDefault` → `TryFirst`), S-2 (Caso1 muta en cadena), S-3 (la regla celda-referenciada es doctrina, no mecanismo automatico), S-4 (resultados 7/10 untracked como evidencia de cierre).

### 9.7 Deuda declarada (fuera de este plan)

- Correccion de plantilla (EXTEMP ASE5 + R2 ASE4 `12.462,26`): v3 de la propietaria + T0 de dueno.
- `ConsolidadoAse.Extemp` mal alimentado (primera fila `Mes`) y gate preexistente `F25 vs Extemp` tautologico (`ValidadorBasico.cs:801`): frente propio; no alimenta la escritura.

**Estado de ejecucion: H1 VERDE · H2 VERDE (set-diff dentro de la lista cerrada) · H3-docs EJECUTADO · code-review APROBADO CON OBSERVACIONES · W-1 ratificado. PLAN CERRADO 2026-10-19.**
