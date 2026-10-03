# Plan 25 — Roles R1-Q2 leaf resueltos por firma/etiqueta sobre la secuencia espejo (fin del mapa congelado por conteo)

> **Alcance:** eliminar el defecto de **rol** del mapa leaf R1-Q2 (`WorkbookLeafCellMapQ2.R1Q2EditablesPorAse` + `MapearR1Q2`): las celdas destino dejan de resolverse por enésima ocurrencia de roles congelados y pasan a resolverse por **firma de etiquetas (A/B/C/D/E)** dentro de la **secuencia espejo ya cargada** (`LeerEspejoR1` / `BloqueEspejoAseInputs.Filas`), con regla de ausencia por rol (core obligatorio con fail-fast; `Aplic0/Aplic1` opcionales con 0 explícito; agregado EXTEMP = suma de TODAS las filas `Aplicacion`). Cero mutación de geometría (sin espejo nuevo, sin insert/delete, sin reanclaje): el espejo R1 del Plan 21 sigue gobernando la escritura; esta HU corrige la **lectura de roles** y el **agregado** que alimenta el DetRetri.
> **Rector (IRRENUNCIABLE):** `README.md` + `Docs/Propuesta_Proyecto_Automatizacion_Remuneracion_UAESP.md` + `Docs/Detalle de plantilla.docx` + `Docs/Prompt_Maestro_Proyecto_Remuneracion_UAESP _ Vo .docx`. **Invariantes del proyecto:** NUNCA sobrescribir fórmulas (solo valores); tolerancia ±0.5; ASE4 sin columna "Especiales" (0 si ausente); Q2 aplica SALDOS POR NOTA/RETRIBUCION NEGATIVA (`Periodo.NumeroQuincena == 2`, `AjustesSfT = 0` en Q1); redondear a entero antes de DetRetri/columna D; build 0 warnings; finales de línea CRLF; fail-fast nombrando ASE+reporte+celda/fila; tests SOLO con insumos reales (`Docs/Insumos`, `Docs/Prueba2`), NO fixtures sintéticos.
> **Continuidad:** HU-01..HU-20 cerradas; Plan 21 cerrado (espejo estructural R1: `OpenXmlEspejoR1Mutador`, `LeerEspejoR1`, deltas por bloque; ver `plans/21 - espejo-estructural-reporte-componentes-r1.md` + `plans/21 - T0 Evidencia.md`); Plan 23 (opcionalidad `Vlr Intereses` en SALDOS/RETRI, WU-1 aplicada, suite base 303/303, sin commit; ver `plans/23 - opcionalidad-vlr-intereses-saldos-retribucion.md`). Este plan NO reabre la semántica del espejo R1 ni la opcionalidad 2.5 ni ninguna aritmética fuera de EXTEMP: **cero cambios de fórmulas de negocio; Q1/Q2-julio intactos por construcción.**
> **Fuente única de verdad:** `plans/24 - T0 Evidencia.md` — **T0 ya cerrado sin NEEDS_CONTEXT** (veredictos T0a..T0e congelados; este plan NO los re-deduce ni los contradice). Los planes 21/23 son solo antecedente.
> **Insumos canónicos (no se re-fijan, no se inventan datos):** plantilla canónica `Docs/Insumos/REMUNERACION 2026072/Plantilla_ Remuneracion 202607-2.xlsx`; goldens `Remuneracion 202607-1 Total.xlsx` / `Remuneracion 202607-2 Total.xlsx` (+ cachés `%TEMP%\opencode\`); fuentes Q2-julio `Docs/Insumos/REMUNERACION 2026072/{1..5}-*/Recaudoporcomponente_*1607*_31072026*.xlsx`; fuentes agosto `Docs/Prueba2/Insumos/{1..5}-*/Recaudoporcomponente_*1608*_31082026*.xlsx` + `R10_Remuneracion_2026082.xlsx` (hoja `DetRetri2026082`, D9:D13 por ASE, único oráculo de agosto).
> **Numeración:** el 24 quedó reservado por la evidencia T0 (`plans/24 - T0 Evidencia.md`); este plan toma el primer correlativo libre, **25**.
> **Estado:** CERRADO — WU-1 (T1+T2) y WU-2 (T3+T4) implementadas y verificadas; build 0 warnings; suite 310/310. Sin commit (lo hace el Ingeniero con `#commit`). Ver §6 (cierre).
> **Fecha:** 2026-10-02

---

## 0. Clarification Gate

**T0 está cerrado y las cinco decisiones de diseño ya fueron ratificadas por el Ingeniero (ver §0.3 D-A..D-E): no hay preguntas bloqueantes pendientes.** El plan viene redactado en la ruta aprobada (estrategia ii: resolver roles por firma/etiqueta sobre la secuencia espejo; plan B = parche mínimo del mapa, acotado a T-alt). Punto de bloqueo respetado: NO se implementa nada en este documento.

### 0.1 Qué está verificado y qué NO (honestidad técnica)

**Verificado — hechos congelados del T0 (fuente: `plans/24 - T0 Evidencia.md`, NO re-deducidos aquí):**

| # | Hecho congelado | Origen T0 | Consecuencia para el plan |
|---|---|---|---|
| V-T0-1 | Agosto: solo ASE3 (de 5) no trae filas `Aplicacion nuevos x reversion` (0 filas); julio Q2: 5/5 las traen. Q1 no usa este mapa (`MapearR1`, EXTEMP estático 0 en ASE2/ASE5) | T0a (§2.1) | La ausencia es específica de ASE3-agosto; la regla de ausencia se declara solo para roles `Aplic` |
| V-T0-2 | **Rol mal asignado:** `F37` (ASE1) y `F270` (ASE3) bindeados a `Subs0` cuando son filas `Aplicacion`. En julio `Subs0==Aplic1` por coincidencia (bloques monocompañía); en agosto ASE3 `Subs0 = −1.087.996.443,09` y el desvío de DetRetri sería ~1.088e9 | T0 §§1/3.3/3.4 | D-A: `Subs0 → Aplic1` en F37/F270; la opcionalidad pura sin corrección de rol está rechazada por datos |
| V-T0-3 | **Regla validada vs oráculo R10 de agosto** (ASE3 RECAUDO TOTAL 16.369.059.896; extemp implícito 0,33): `EXTEMP = Aplic0 + Aplic1` cierra ±0.5 en **10/10** (5 ASE x {Q2-julio, agosto}); `Subs0` falla en ASE2/ASE4/ASE5 y ASE3-agosto | T0b (§3.4, tabla 10 combinaciones) | Agregado congelado: suma de TODAS las filas `Aplicacion` (0 si ninguna); D-C |
| V-T0-4 | El espejo (Plan 21) supersede la ESCRITURA R1 en 5-ASE (`omitirR1`) pero **NO** el agregado `extemp` (`ExtemporaneoEsperadoPorAse` → `TotalD104` → DetRetri) ni `totOpt` (gate Σ empresas) ni `F25` (gate HU-02) | T0c (§4.1/4.2) | Esta HU corrige lectura+agregado; la escritura espejo no se toca |
| V-T0-5 | Tabla congelada roles → firmas → celdas (§6): `Aplic` opcionales (0 explícito), `Mes` obligatorios; invariantes duras `Componente/Total`, `Subs/Cont/Total`, `Total` final 15/15 | T0e (§6) | Regla de ausencia por rol de la SPEC (R-F-3) |

**Re-verificado en esta planificación** (lectura directa de código, sin cambiar nada):

| # | Hecho | Método | Consecuencia |
|---|---|---|---|
| V1 | `R1Q2EditablesPorAse[1]` bindea `("F37", Subs0)` y `[3]` bindea `("F270", Subs0)`; el resto de ASE bindea `Aplic0/Aplic1/LAplic0` | Lectura `WorkbookLeafCellMapQ2.cs:57-94` | Confirma V-T0-2 al símbolo: el cambio toca esos dos bindeos y la resolución |
| V2 | `MapearR1Q2` filtra `filasMes` (B=`Mes`∧C=`Total`), `filasAplicacion` (B contiene `Aplicacion nuevos x reversion`∧C=`Total`), `filasSubsidio` (E contiene `Subsidio(-)/Contribucion(+)`); `RolFila` resuelve por `ElementAtOrDefault(n)` y el `foreach` lanza `CalculoInvalidoException("ASE {id}: la fuente R1-Q2 no trae la fila del rol {fuente} ... {celda}")` si es null; `extemp += Subs0/Aplic0/Aplic1`, `extemp -= LAplic0`; `F25/F41/L25` conservan semántica Q1 | Lectura `ExcelDataReaderWorkbookLeafInputReader.cs:966-1073` | Sitio exacto del cambio: resolución + fail-fast + acumulado de `extemp` |
| V3 | `LeerEspejoR1` ya carga por ASE la secuencia `List<FilaEspejoR1>` (A/B/C/D/E + `ValoresPorColumna` por encabezado, `SERVICIO ESPECIALES` dinámico) delimitada por etiquetas (`Vlr Servicio` .. `Total` final) con invariantes duras T0e; `FilaEspejoR1` es modelo puro con `Firma`, `Valor(encabezado)`, `EsComponenteTotal/EsSubsContTotal/EsTotalFinal` | Lectura `ExcelDataReaderWorkbookLeafInputReader.cs:537-623`, `FilaEspejoR1.cs`, `BloqueEspejoAseInputs.cs` | La secuencia espejo existe y es reutilizable como fuente de resolución (D-B); solo faltan los predicados de rol `Mes/Aplicacion/Subsidio` sobre ella |
| V4 | `ProcesadorPeriodo` ya adjunta `EspejoR1 = LeerEspejoR1(ase, rutaR1)` y `OpenXmlPlantillaWriter` escribe con `omitirR1: espejoR1` en 5-ASE; el flujo single-ASE (`ProcesadorRemuneracion`) comparte el reader sin espejo adjunto | Lectura `ProcesadorPeriodo.cs:143`, `OpenXmlPlantillaWriter.cs:218-274`, grep `ProcesadorRemuneracion` | D-D es viable: el reader es compartido; la resolución por firma debe funcionar con y sin bloque espejo adjunto |
| V5 | `plans/` 01..24 ocupados (24 = evidencia T0); este plan es el **25**, primer número libre | Listado `plans/` | Numeración del documento |

**Aportado como hechos congelados del encargo (NO re-verificado en disco en esta planificación; se acredita en tests):**

| # | Límite | Tratamiento en el plan |
|---|---|---|
| N1 | Suite base **303/303** verde tras WU-1 del Plan 23 (sin commit) | Red ciega de cada tarea: 303/303 antes y después; DoD-2 |
| N2 | Oráculo R10 de agosto: único oráculo de agosto (`DetRetri2026082` D9:D13); sin golden de agosto de UAESP | Regresión 2026082 = DetRetri-vs-R10 por ASE + invariantes + deltas (D-E); sin asserts ±0.5 contra salida de agosto distinta del R10; prohibido inventar valores |
| N3 | Regla de agregado con más de 2 filas `Aplicacion` sin evidencia (observado: siempre 0 o 2) | D-C la fija por decisión ejecutiva (suma de TODAS); test con insumos reales cubre 0 y 2; el caso >2 queda especificado sin fixture sintético |

### 0.2 Mapeo al Rector (in vs out)

**Entra:**

| Requisito | Superficie de cambio |
|---|---|
| Roles R1-Q2 resueltos por firma/etiqueta sobre la secuencia espejo ya cargada: cada celda destino se llena desde la fila fuente cuya firma (A/B/C/D/E) corresponde a su rol, en orden de aparición; fin del `ElementAtOrDefault(n)` sobre listas filtradas por raw | `ExcelDataReaderWorkbookLeafInputReader.MapearR1Q2` (resolución + fail-fast + acumulado) + predicados de firma sobre `FilaEspejoR1`/`BloqueEspejoAseInputs` (Core puro, sin I/O) |
| Corrección de rol D-A: `F37/F270: Subs0 → Aplic1` (ASE1/ASE3); el miembro `Subs0` deja de usarse en Q2 | `WorkbookLeafCellMapQ2.R1Q2EditablesPorAse[1]/[3]` (dos bindeos) + rama `Subs0` en `MapearR1Q2` (acumulado y switch de valor) |
| Regla de ausencia por rol (T0e §6): core obligatorio con fail-fast que nombra ASE+reporte+celda (`Mes0/1/2`, `Lmes0/1(/2)`); `Aplic0/Aplic1/LAplic0` opcionales con 0 explícito; agregado EXTEMP = suma de TODAS las filas `Aplicacion` (F) menos Especiales de la primera (0 si ninguna) | Mismo lector (rama tolerante + sumatoria); `totOpt`, `F25/F41/L25`, gates de coherencia intactos |
| Alcance 5-ASE **y** single-ASE (el reader es compartido; `ProcesadorRemuneracion` no adjunta espejo) | La resolución debe operar sobre la secuencia del período en ambos paths (firma la fija §2.1); sin bifurcación por período/ASE |
| Goldens julio Q1+Q2 ±0.5 intactos + regresión 2026082 con DetRetri-vs-R10 por ASE (5/5) + invariantes + deltas espejo | `Remuneracion.IntegrationTests` (xUnit + FluentAssertions): regresión continua + test dedicado con insumos reales |
| Trazabilidad doctrinal mínima | Diff acotado en `.opencode/project-context.md` (fin del mapa congelado por conteo en R1-Q2 + regla de ausencia por rol) |

**Sale (EXPLÍCITO):** nueva mutación de geometría (sin insert/delete, sin reanclaje; el espejo del Plan 21 no se toca); cambios a la aritmética fuera de EXTEMP (`totOpt`, R2, R4, AJUSTES-SF-T, banco, BCE, conciliaciones, INTERVENTORIA-valores, `F25`); regla general "ausencia de cualquier rol = 0" (rechazada: `Mes` y gates exigen el dato); reinterpretar valores de agosto (dato, no cálculo); Q1 (no usa este mapa; `AjustesSfT = 0` intacto); UI WinForms y CLI salvo el resumen log existente; paquetes NuGet; commits (los hace el Ingeniero con `#commit`).

### 0.3 Decisiones ejecutivas (ratificadas por el Ingeniero — aprobación = aceptar estas)

| ID | Decisión |
|---|---|
| D-A | **`Subs0 → Aplic1` en F37 (ASE1) y F270 (ASE3).** Evidencia: la fila destino es `Aplicacion` (T0 §4.4) y `Aplic0+Aplic1` casa 10/10 con el R10 mientras `Subs0` falla en ASE2/ASE4/ASE5 y ASE3-agosto. |
| D-B | **Estrategia = resolver roles por firma/etiqueta sobre la secuencia espejo** (opción ii del T0d). Plan B = parche mínimo del mapa (opción iii: solo rebindeo + opcionalidad), acotado a T-alt y activable solo si la implementación demuestra que (ii) no es expresable sin ensuciar el modelo. |
| D-C | **Agregado de `Aplicacion` = suma de TODAS las filas con esa firma** (0 si ninguna). Evidencia actual: siempre 0 o 2; la regla queda fijada por decisión para >2 sin fixture sintético. |
| D-D | **Alcance = 5-ASE y single-ASE.** `ProcesadorRemuneracion` comparte el reader; el fix vive en el reader, no en el orquestador. |
| D-E | **Regresión 2026082 = asertar DetRetri de los 5 ASE contra el R10** (único oráculo de agosto; ASE3 = 16.369.059.896) **+ invariantes + deltas.** Sin golden de agosto: sin asserts ±0.5 contra salida de agosto fuera del R10. |
| D-F | **`totOpt` generalizado = `ΣF(todas las Mes) − Especiales de todas MENOS la última`** (ratificada por el Ingeniero como parte del patrón por firma). Surge como forzante del propio DoD (D-E): ASE5-agosto trae 3 filas `Mes/Total` (T0a §2.1) mientras el mapa Q2 estaba congelado con 2 (julio). Idéntico al mapa congelado en ASE1-4 (3 filas) y ASE5-julio (2 filas, Lmes1 = 0); generaliza ASE5-agosto al mismo 5-término visible del template. Guardián: ASE5-julio = 12.033.011.685,71 (== golden D13); ASE5-agosto = 12.105.458.586,04. |

---

## 1. PROPOSE

### 1.1 Intent

Que el mapa leaf R1-Q2 deje de depender de ocurrencias congeladas que ya fallaron dos veces (L-menores HU-12, `Aplicacion` agosto): cada celda destino R1-Q2 se llena desde la fila fuente identificada por su **firma de etiquetas** dentro de la secuencia espejo del período en curso, con la corrección de rol `Subs0 → Aplic1` y la regla de ausencia por rol del T0e, de modo que agosto (ASE3 sin `Aplicacion`) produzca EXTEMP = 0 y el DetRetri de los 5 ASE cierre contra el R10, con julio Q1+Q2 intactos ±0.5.

### 1.2 In Scope

- Resolución por firma sobre la secuencia espejo en `MapearR1Q2` (firmas `Mes/Total`, `Aplicacion/Total`, `Subsidio`; valores F por encabezado `Total`, L por `SERVICIO ESPECIALES` con 0 si ausente), con D-A (rebindeo F37/F270) y D-C (sumatoria de todas las `Aplicacion`).
- Regla de ausencia por rol: core fail-fast intacto (mensaje actual con ASE+reporte+celda); `Aplic` ausente = 0 explícito en sus celdas + agregado.
- Conservación verificada: `totOpt`, `F25/F41/L25`, gates `Σ empresas` y `F25 vs Extemporáneo HU-02`, escritura espejo (`omitirR1`), Q1 sin tocar.
- Tests con insumos reales: goldens Capa A Q1+Q2 (±0.5) + regresión 2026082 (DetRetri-vs-R10 5/5 + invariantes + deltas) + Q1 verde con `AjustesSfT=0`.
- Diff acotado en `.opencode/project-context.md` (fin del conteo congelado en R1-Q2 + regla por rol + trampa documentada).

### 1.3 Non-Goals

Ver §0.2 "Sale". En particular: NO se agrega ningún rol/firma nuevo sin evidencia; NO se emite valor distinto de 0 para un rol ausente; NO se silencia el fail-fast de core; NO se cambia el mapeo de columnas por encabezado; NO se tocan `R1Q2ProtectedPorAse`, visibles `F343/F345`, puentes AJUSTES/CONSOLIDADO ni nada aguas abajo; NO se reimplementa el espejo.

### 1.4 Riesgos

| ID | Riesgo | Mitigación |
|---|---|---|
| R-ROL-OCULTO | Otro bindeo celda→rol con coincidencia numérica latente (`Subs0==Aplic1` enmascaró este) que julio no delata | La resolución por firma elimina la clase (el valor sale de la fila con la firma correcta, no del rol declarado); la regresión R10 5/5 es el detector aritmético |
| R-AGREGADO-N | Futura fuente con 1 o 3+ filas `Aplicacion` donde la semántica "suma de todas" no aplique | D-C la fija por decisión; T0e §6 la congela; si UAESP cambia la semántica, requiere su propio T0 con fuente real (no se especula aquí) |
| R-DOBLE-FUENTE | Divergencia entre la lista raw (`filas`) y la secuencia espejo (`Filas`) si la resolución usa filtros distintos | Fuente única: los predicados de firma viven en Core sobre `FilaEspejoR1` y el reader los usa para ambas (misma normalización de etiquetas que `LeerEspejoR1`); grep de cierre sin filtros paralelos |
| R-REGRESION-JULIO | La nueva resolución altera el path julio (donde todo está presente y debe leerse idéntico) | Goldens Q1+Q2 ±0.5 como red ciega en cada tarea; julio no toca la rama tolerante (todas las filas presentes) |
| R-ORACULO-AGOSTO | Sin golden de agosto: solo el R10 acredita negocio | D-E: DetRetri-vs-R10 5/5 + invariantes + deltas; prohibido inventar golden o valores |
| R-SOBREDISENO | Convertir la firma puntual en framework general de roles opcionales | Alcance congelado a R1-Q2 `Aplic` (D-A..D-C); cualquier otro rol opcional futuro requiere su propio T0/HU |
| R-TRAMPA-OPC | **Trampa documentada:** aplicar opcionalidad (`Aplic0` ausente=0) SIN la corrección de rol D-A desplaza el fallo a la validación DetRetri-vs-R10 con error ~1.088e9 en ASE3 (T0b §3.4) | La SPEC exige D-A + opcionalidad juntas (R-F-1); el test de regresión aserta el DetRetri final, no solo la ausencia de throw |

---

## 2. DESIGN

### 2.1 Enfoque: la secuencia observada resuelve los roles; el mapa declara celdas, no ocurrencias

**Modelo (Core puro, sin I/O).** Sobre `FilaEspejoR1` se agregan los predicados de rol que hoy solo existen como filtros raw en `MapearR1Q2`, con la MISMA normalización de etiquetas que `LeerEspejoR1` (nunca índices fijos):

- `EsMesTotal`: B=`Mes` ∧ C=`Total` (equivale a `filasMes`).
- `EsAplicacionTotal`: B contiene `Aplicacion nuevos x reversion` ∧ C=`Total` (equivale a `filasAplicacion`).
- `EsSubsidio`: E contiene `Subsidio(-)/Contribucion(+)` (equivale a `filasSubsidio`; se conserva solo como predicado de cierre/documentación, ya sin uso en el agregado Q2 tras D-A).
- Valores: F = `Valor("Total")`, L = `Valor("SERVICIO ESPECIALES") ?? 0` (ASE4/ausencia → 0, invariante del proyecto).

La firma exacta (propiedades bool en `FilaEspejoR1` o helpers en `BloqueEspejoAseInputs`: p. ej. `FilasMes()`, `FilasAplicacion()`) la elige el implementador; condición: **un solo sitio** define cada firma (R-DOBLE-FUENTE), y `MapearR1Q2` deja de filtrar `List<object?[]>` en paralelo.

**Lectura (único cambio de comportamiento, en `MapearR1Q2`).** El método opera sobre la secuencia del período (`BloqueEspejoAseInputs.Filas` cuando el orquestador la adjunta; en single-ASE se deriva del mismo `LeerEspejoR1` del reader compartido — D-D, sin duplicar filtros):

1. `mes = Filas.Where(EsMesTotal)` en orden; `aplic = Filas.Where(EsAplicacionTotal)` en orden.
2. Por cada `(celda, fuente)` de `ObtenerR1Q2Editables(ase.Id)` (con D-A aplicado: F37/F270 ← `Aplic1`):
   - `Mes0/1/2` ← `mes[0/1/2]` (F); `Lmes0/1(/2)` ← Especiales de la misma fila; **ausente → throw actual intacto** (core).
   - `Aplic0` ← `aplic[0]` (F), `Aplic1` ← `aplic[1]` (F), `LAplic0` ← Especiales de `aplic[0]`; **ausente → 0 explícito** en la celda (T0e §6.1).
   - `Subs0` deja de aparecer en el mapa Q2 (tras D-A no hay consumidor; si el enum subsiste por compatibilidad, su rama lanza o se retira con grep de cierre — lo fija la implementación dentro de D-A).
3. Agregados (D-C): `totOpt = ΣF(mes) − ΣL(mes presentes en el mapa)` (idéntico a hoy, incluso ASE5 sin Mes2); **`extemp = ΣF(todas las filas aplic) − (aplic.Any ? Esp(aplic[0]) : 0)`**. Con 2 filas coincide con `Aplic0+Aplic1` del T0b 10/10; con 0 filas es 0 (ASE3-agosto → extemp implícito 0,33 ±0.5 vs R10).
4. `F25/F41/L25` (primera/segunda fila `Mes`) y `F30/F10/L10 = 0` intactos; mensajes de throw intactos (ASE+reporte+celda).

**Por qué 0 explícito y no "no escribir":** las celdas leaf R1 en 5-ASE las supersede el espejo (`omitirR1`), pero el **agregado** `extemp` vive en `WorkbookLeafInputsR1` y alimenta `TotalD104`/DetRetri en ambos flujos; el 0 explícito hace la invariante `EXTEMP(ASE3-agosto)=0` asertable a nivel de dominio sin depender del estado de la plantilla. En single-ASE el writer escribe esos 0 como cualquier valor (guard anti-fórmula vigente).

**Escritura: sin cambios.** `OpenXmlEspejoR1Mutador`, `omitirR1`, `R1Q2ProtectedPorAse`, visibles `F343/F345` y puentes intactos. Doctrina D-B del Plan 21 vigente: fuera de los bloques espejo el fail-fast "slot ausente ≠ 0" sigue intacto; dentro de esta HU solo los roles `Aplic` son tolerantes (T0e).

### 2.2 Alternativas y trade-offs

| Eje | A — Firma sobre secuencia espejo (recomendada, D-B) | B — Parche mínimo del mapa (plan B, T-alt) |
|---|---|---|
| Qué se hace | Predicados en Core + `MapearR1Q2` resuelve sobre `Filas`; rebindeo F37/F270; sumatoria D-C | Solo rebindeo `Subs0→Aplic1` en las dos entradas + `Aplic` opcional (= 0) sobre los filtros raw actuales |
| Diff | Medio: nuevo predicado + reescritura de resolución (un solo sitio) | Mínimo: 2 bindeos + 1 rama `if` |
| Defecto de rol | Eliminado por construcción (la firma manda) | Corregido solo donde se vio (riesgo R-ROL-OCULTO subsiste) |
| Doble fuente | Eliminada (una sola definición de firma) | Subsiste (filtros raw + espejo por caminos distintos, divergen) |
| Veredicto | **Principal (D-B)** | **Fallback acotado (T-alt)** si la implementación demuestra que (A) no es expresable sin ensuciar el modelo |

Descartadas explícitamente: opcionalidad pura sin corrección de rol (rechazada por datos: desvío ~1.088e9, T0d §5.1, trampa R-TRAMPA-OPC); regla general "todo-ausente = 0" (debilita fail-fast de `Mes`/gates); 0 silencioso para core; fabricar valores (prohibido, T0 §6.2 del Plan 23 por analogía).

### 2.3 Por qué NO hay desplazamiento ni reanclaje (tratamiento explícito)

Porque no se mueve ninguna fila ni cambia ninguna dirección del template: las celdas (`F12..F552`, `L*`), las protegidas (`F53/F55`, `F343/F345`, …), los 292 refs aguas abajo y los 7 mapas de anclaje quedan idénticos en anclaje. La prueba dura: el diff no toca `R1Q2ProtectedPorAse`, `Filas*` de otros mapas ni el mutador; solo `R1Q2EditablesPorAse[1]/[3]` (dos bindeos), `MapearR1Q2` y los predicados nuevos. El inventario T0c del Plan 24 (§4) se usa como checklist de no-toque.

### 2.4 Contratos afectados (resumen)

| Contrato | Cambio |
|---|---|
| `WorkbookLeafCellMapQ2` (Infrastructure): `R1Q2EditablesPorAse[1]/[3]` | Solo rebindeo D-A (`F37/F270: Subs0→Aplic1`); resto de entradas, `R1Q2ProtectedPorAse`, filas y fórmulas intactos |
| `FilaEspejoR1` / `BloqueEspejoAseInputs` (Core puro) | Solo predicados de firma `Mes/Aplicacion(/Subsidio)` + acceso a secuencias en orden; invariantes duras T0e intactas |
| `IWorkbookLeafInputReader` con `ExcelDataReaderWorkbookLeafInputReader.MapearR1Q2` | Resolución sobre la secuencia (firma+orden), fail-fast de core intacto, `Aplic` ausente = 0 explícito, agregado D-C; `totOpt`/`F25/F41/L25`/mensajes intactos |
| `ProcesadorPeriodo` / `ProcesadorRemuneracion`, `OpenXmlPlantillaWriter`, `IValidador`, R10 | Sin cambios (el fix vive en el reader compartido; D-D por construcción) |
| `Remuneracion.IntegrationTests` (xUnit + FluentAssertions) | Regresión julio + regresión 2026082 con DetRetri-vs-R10 5/5; suite base 303/303 como red ciega |
| `.opencode/project-context.md` | Diff acotado: fin del conteo congelado en R1-Q2 + regla de ausencia por rol + trampa R-TRAMPA-OPC |

### 2.5 Architecture Validation Certificate

| Principio | Veredicto |
|---|---|
| **SRP** | ✅ Una sola razón de cambio por costura: Core declara firmas, el mapa declara celdas, el lector resuelve y agrega, los tests demuestran. Sin ramas `if agosto` ni `if período` (la regla es por firma, válida en todo Q2 presente y futuro). |
| **OCP** | ✅ El mapa deja de congelar ocurrencias por período: la cardinalidad futura (0, 2 o N filas `Aplicacion`) se absorbe por parámetro (secuencia del período + sumatoria D-C), no con un mapa nuevo por quincena. Esta HU es la última que debería necesitar un bindeo por conteo para R1-Q2. |
| **DIP** | ✅ Cambios detrás de `IWorkbookLeafInputReader` y modelo puro de Core; `ProcesadorPeriodo`/`ProcesadorRemuneracion` y el writer no conocen la regla ni se tocan. |
| **Best practices** | ✅ Fuente-define-la-forma (doctrina D-B del Plan 21 extendida a la lectura de roles); fail-fast preservado donde protege (core nombra ASE+reporte+celda); tolerancia solo donde la evidencia la autoriza (único rol variable en 15 fuentes reales); fórmulas nunca sobrescritas; columnas por encabezado; tolerancia ±0.5; Q1 excluido por dominio (`NumeroQuincena`), no por parche. |
| **Performance** | ✅ Sin impacto: la secuencia espejo ya se carga en 5-ASE; en single-ASE una lectura etiquetada más por ASE sin I/O adicional fuera de la fuente ya abierta; tests reutilizan `Remuneracion.IntegrationTests` con insumos reales existentes. |

---

## 3. SPEC

### 3.1 Requirements

| ID | Requisito | Grupo | Criterio de aceptación |
|---|---|---|---|
| R-F-1 | Corrección de rol D-A + resolución por firma: `F37/F270 ← Aplic1` (segunda fila `Aplicacion/Total` en orden); `F17/F250 ← Aplic0`; `L17/L250 ← Especiales de Aplic0`; valores F por encabezado `Total` | Fix | Goldens Q2 ASE1/ASE3 ±0.5 con valores `Aplicacion` reales (julio: F37/F270 = 40.364,31 / 627.141,97, no `Subs0`); `grep Subs0` sin usos Q2 |
| R-F-2 | Agregado D-C: `extemp = ΣF(todas las Aplicacion) − Esp(primera)`; 0 filas ⇒ 0; `LAplic0` sin `Aplic0` ⇒ 0 | Fix | 10/10 combinaciones ±0.5 vs extemp implícito del R10 (tabla T0b §3.4); ASE3-agosto `extemp = 0` (implícito 0,33) |
| R-F-3 | Regla de ausencia por rol: `Mes0/1/2` y `Lmes` ausentes con fail-fast actual intacto (mensajes con ASE+reporte+celda); `Aplic0/Aplic1/LAplic0` ausentes con 0 explícito | Fix | Sin cambio de mensajes (`grep "la fuente R1-Q2 no trae la fila del rol"`); S3/S6 |
| R-F-4 | Julio intacto: con todas las filas presentes la resolución por firma produce los mismos valores que el mapa actual (incl. `Subs0==Aplic1` en ASE1/ASE3-julio y subtotales multi-compañía en ASE2/4/5) | Fix | Goldens Q2 5/5 ±0.5; extemp ASE2/4/5 con `Aplic0+Aplic1` (no `Subs0`) cierra contra R10-julio implícito |
| R-F-5 | Conservación: `totOpt`, `F25/F41/L25`, gates `Σ empresas` y `F25 vs HU-02`, escritura espejo (`omitirR1`), protegidas `F343/F345`, Q1 sin este path (`AjustesSfT=0`) | Conservación | Goldens Q1+Q2 ±0.5; Q1 verde; diff no toca protegidas/mutador |
| R-R-1 | Suite base 303/303 verde antes y después (red ciega, `dotnet build` 0 warnings) | Transversal | Build + suite completos en cada WU |
| R-R-2 | Regresión 2026082 (D-E): flujo 5-ASE con `Docs/Prueba2/Insumos` supera `MapearR1Q2` sin throw, genera workbook con fórmulas preservadas, y el DetRetri calculado de los 5 ASE cierra ±0.5 post-redondeo contra `R10_Remuneracion_2026082` (D9:D13); invariantes de cierre + deltas espejo {-3,-9,-6,+6,+8} sin regresión | Transversal | Test dedicado; ASE3 = 16.369.059.896 ±0.5; sin asserts de negocio fuera del R10 |
| R-R-3 | Cero fórmulas sobrescritas (assert estructural) + sin datos inventados (todo 0 remite a ausencia evidenciada T0a; todo valor remite a fuente real o R10) | Transversal | Guard anti-fórmula + checklist §2.3 verdes |
| R-R-4 | `project-context.md` con fin-del-conteo en R1-Q2 + regla por rol + trampa R-TRAMPA-OPC (diff acotado, sin código en el mismo commit) | Transversal | Diff solo doctrina/reglas |

### 3.2 Scenarios (Given/When/Then)

- **S1 (agosto-ASE3, el blocker):** Given fuente agosto ASE3 sin `Aplicacion` (0 filas, `Docs/Prueba2/Insumos/3-Ciudad Limpia/`), When `LeerLeafInputs(ase3, ...)` Q2, Then no lanza; celdas F250/F270/L250 en 0 explícito; `ExtemporaneoEsperadoPorAse = 0`.
- **S2 (julio-ASE3, no-regresión del rol):** Given fuente Q2-julio ASE3 (Aplic0/Aplic1 = R14/R34), When se lee, Then F250/F270 traen esos valores (627.141,97 en F270, no `Subs0` como etiqueta aunque coincida el número) y el golden Q2 cierra ±0.5.
- **S3 (core-ausente, fail-fast intacto):** Given secuencia sin una fila `Mes` (acreditado por rama + mensajes intactos; sin sintéticos), When se lee, Then `CalculoInvalidoException` nombra ASE+reporte+celda.
- **S4 (multi-compañía, ASE2/4/5):** Given fuentes Q2-julio y agosto de ASE2/4/5 (subtotales, no ecos), When se lee, Then `extemp = Aplic0+Aplic1−LAplic0` cierra ±0.5 contra el implícito del R10 en las 8 combinaciones (4 ASE x 2 períodos).
- **S5 (single-ASE, D-D):** Given `ProcesadorRemuneracion` (sin espejo adjunto) con fuente agosto ASE3, When se leen los leaf inputs, Then mismo resultado que en 5-ASE (0 explícito, sin throw) y las celdas R1 se escriben con esos valores.
- **S6 (regresión-2026082, D-E):** Given `Docs/Prueba2/Insumos` (5 ASE + R10), When `ProcesadorPeriodo` corre Q2 completo, Then supera `MapearR1Q2`, genera workbook con fórmulas preservadas, DetRetri 5/5 = R10 D9:D13 ±0.5 post-redondeo, invariantes duras presentes y deltas {-3,-9,-6,+6,+8}.
- **S7 (Q1-intacto):** Given período Q1, When se procesa, Then nunca entra a `MapearR1Q2`, `AjustesSfT = 0`, goldens Q1 ±0.5 verdes.
- **S8 (cero-geometría):** Given el diff de la HU, When se audita contra el inventario T0c del Plan 24 (§4), Then ninguna fila fija, fórmula, puente (AJUSTES, CONSOLIDADO D85:D89/D47:D51, visibles F343/F345) o mapa de anclaje cambió salvo los dos bindeos D-A y los predicados nuevos.

---

## 4. TASKS

| ID | Tarea | Grupo | Depende de | Verificación |
|---|---|---|---|---|
| T0 | **CERRADO — Discovery evidencia (NO es PR nuevo):** veredictos T0a..T0e en `plans/24 - T0 Evidencia.md` (dumps por ASE x período, coincidencia `Subs0==Aplic1`, validación R10 10/10, impacto en celdas, veredicto (ii)+(iii) plan B, tabla congelada §6) | Fase 0 | — | DONE sin NEEDS_CONTEXT; este plan lo consume como hechos congelados |
| T1 | Core: predicados de firma `EsMesTotal` / `EsAplicacionTotal` (/ `EsSubsidio` solo cierre) sobre `FilaEspejoR1` + acceso en orden en `BloqueEspejoAseInputs`; mismo criterio de normalización que `LeerEspejoR1`; sin I/O, sin deps nuevas | Fix | T0 | Compila en `Remuneracion.Core`; S2/S4 a nivel de secuencia (insumos reales); build 0 warnings |
| T2 | Mapa + lector: rebindeo D-A (`[1] F37→Aplic1`, `[3] F270→Aplic1`); `MapearR1Q2` resuelve sobre la secuencia (firma+orden), `Aplic` ausente = 0 explícito, core ausente = throw actual, agregado D-C; rama `Subs0` sin usos Q2 (retiro o lanzamiento, con grep de cierre) | Fix | T1 | S1/S2/S3/S4/S5; `grep Subs0` acotado a lo declarado; goldens julio ±0.5; build 0 warnings |
| T-alt | *(Solo si la implementación demuestra que T2 no es expresable sin ensuciar el modelo)* Plan B: solo rebindeo D-A + `Aplic` opcional sobre filtros raw, sin secuencia; sustituye la resolución de T2, mantiene D-A/D-C/D-E | Fix-alt | T1 | S1/S2/S4/S6 por parche mínimo; goldens julio ±0.5; R-ROL-OCULTO/R-DOBLE-FUENTE quedan como deuda registrada |
| T3 | Tests con insumos reales: (a) regresión goldens Capa A Q1+Q2 ±0.5 (suite 303/303 verde); (b) regresión 2026082 5-ASE con DetRetri-vs-R10 5/5 + invariantes + deltas; (c) single-ASE agosto ASE3 (D-D); (d) Q1 verde con `AjustesSfT=0`; (e) assert estructural cero-fórmulas-sobrescritas | Transversal | T2 (o T-alt) | R-R-1/R-R-2/R-R-3 + S6/S7/S8; `dotnet build Remuneracion.WinForms/Remuneracion.WinForms.slnx` 0 warnings; suite completa verde |
| T4 | Docs: `project-context.md` (fin del mapa congelado por conteo en R1-Q2; regla de ausencia por rol T0e §6; trampa R-TRAMPA-OPC como follow-up/doctrina) | Transversal | T3 | Diff acotado a doctrina/reglas; sin código en el mismo commit; S8 checklist citado |

**Orden sugerido:** T0 (cerrado) → T1 → T2 (o T-alt) → T3 → T4. T3 corre como regresión continua después de T1 y T2 (julio verde antes y después de cada tarea). Build 0 warnings antes de marcar cada tarea completa. Sin commits (los hace el Ingeniero). **Work units encadenadas en PRs (Fase 0 ya hecha, no es PR):** WU-1 = T1+T2 (predicados + lector/mapa); WU-2 = T3+T4 (regresión 2026082 + docs). PRs secuenciales, cada uno con gates DoD 1-2. T-alt, si se activa, vive dentro de WU-1 y se declara en el PR.

---

## 5. Resumen ejecutivo, riesgos y decisiones

- **Resumen:** el blocker de agosto no era un "slot ausente" sino un **rol mal asignado** enmascarado por una coincidencia numérica de julio (`Subs0==Aplic1` en bloques monocompañía). La solución ratificada (D-A..D-E) corrige el rol (F37/F270 → `Aplic1`), resuelve cada celda por **firma sobre la secuencia espejo** en vez de por ocurrencia congelada, fija el agregado como suma de todas las `Aplicacion` (0 si ninguna, validado 10/10 contra el R10) y lo acredita con DetRetri-vs-R10 5/5 en 2026082, con julio y Q1 intactos por construcción y sin mover una sola fila ni tocar una sola fórmula.
- **Riesgo principal:** R-ROL-OCULTO (otro rol enmascarado que julio no delata) + R-ORACULO-AGOSTO (solo el R10 acredita agosto) + R-TRAMPA-OPC (opcionalidad sin D-A desplaza el fallo al DetRetri con ~1.088e9 de error). Contenidos por la resolución por firma (elimina la clase), D-E (detector aritmético 5/5) y la SPEC que exige D-A + opcionalidad juntas.
- **Decisión para el Ingeniero:** ratificar D-A..D-E (§0.3). Aprobación del plan = aceptación de esas cinco decisiones. **Sin preguntas bloqueantes pendientes.**

### DoD (Definition of Done)

1. `dotnet build Remuneracion.WinForms/Remuneracion.WinForms.slnx` con 0 warnings.
2. Suite completa verde sin regresión (base 303/303) + goldens Capa A Q1+Q2 en dif ±0.5 intactos.
3. Regresión 2026082 verde: supera `MapearR1Q2` sin throw, genera workbook con fórmulas preservadas, DetRetri 5/5 = R10 D9:D13 ±0.5 post-redondeo (ASE3 = 16.369.059.896), invariantes duras presentes, deltas espejo {-3,-9,-6,+6,+8} sin regresión; Q1 verde con `AjustesSfT=0`.
4. Ninguna fórmula sobrescrita (guard anti-fórmula + assert estructural + checklist S8 verdes); diff geométrico acotado a los dos bindeos D-A y los predicados nuevos (§2.1).
5. `project-context.md` actualizado (fin-del-conteo R1-Q2 + regla por rol + trampa R-TRAMPA-OPC); sin datos inventados (todo 0 remite a ausencia evidenciada, todo valor a fuente real o R10).
6. Sin commits del agente (los hace el Ingeniero con `#commit`); finales de línea CRLF; sin emojis.

### Follow-ups explícitos (fuera de esta HU)

- R-AGREGADO-N: si UAESP trae 1 o 3+ filas `Aplicacion`, la sumatoria D-C queda validada por decisión pero sin evidencia runtime; requiere su propio T0 con fuente real si la semántica cambia.
- R4 espejo pendiente del Plan 21 (motor solo R1) y R4-por-empresa Q2 (recorte HU-20/G2-D1); R-EXTRA-CONCEPTO del Plan 23: intactos, no tocados por esta HU.

---

## 6. CIERRE (WU-1 + WU-2) — 2026-10-02

> **Estado: CERRADO.** Sin commit (los hace el Ingeniero con `#commit`). Build 0 warnings; suite **310/310** verde.

### 6.1 WU-1 (T1+T2) — implementada y verificada

| Qué | Archivo |
|---|---|
| Predicados de firma de rol + acceso en orden + agregado D-C | `Remuneracion.Core/Models/FilaEspejoR1.cs` (`EsMesTotal`, `EsAplicacionTotal`, `EsSubsidio`; el resto de firmas espejo), `Remuneracion.Core/Models/BloqueEspejoAseInputs.cs` (`FilasMes`, `FilasAplicacion`, `FilasSubsidio`, `SumarAplicacion`) |
| D-A: rebindeo `F37/F270` de `Subs0`→`Aplic1`; enum `Subs0` retirado | `Remuneracion.Infrastructure/Excel/WorkbookLeafCellMapQ2.cs` |
| `MapearR1Q2` resuelve por firma sobre la secuencia (misma lectura, R-DOBLE-FUENTE); `Aplic` ausente = 0 explícito; core ausente = fail-fast intacto; agregado D-C | `Remuneracion.Infrastructure/Excel/ExcelDataReaderWorkbookLeafInputReader.cs` |
| Fixtures R1/R2/R4 de agosto + tests por firma | `Remuneracion.IntegrationTests/Insumos.cs`, `Remuneracion.IntegrationTests/R1Q2ResolucionPorFirmaTests.cs` |

**Hallazgo forzante (D-F):** ASE5-agosto trae **3** filas `Mes/Total` (T0a §2.1) y el mapa Q2 congelado asumía 2 (julio). El `totOpt` por mapa (2 filas) daba 13.546.787.188,77 y el DetRetri de ASE5 divergía +1.441.328.603 del R10; se generalizó el `totOpt` a la fórmula visible del template. Sin ello, el propio DoD (DetRetri 5/5 vs R10) no cerraba.

### 6.2 WU-2 (T3+T4) — implementada y verificada

| Qué | Archivo | Trazabilidad |
|---|---|---|
| Regresión 2026082 end-to-end: DetRetri 5/5 vs R10 D9:D13 (DIF=0; ASE3 = 16.369.059.896), invariantes de cierre T0e, deltas {-3,-9,-6,+6,+8}, fórmulas preservadas | `Remuneracion.IntegrationTests/Regresion2026082Tests.cs` | R-R-2, S6, D-E |
| Guardián `totOpt` generalizado: ASE5-agosto (3 filas) = 12.105.458.586,04; ASE5-julio (2 filas) = 12.033.011.685,71 (idéntico al mapa congelado); mismo ASE, distinta cardinalidad, ambos contra su valor | `Remuneracion.IntegrationTests/TotOptGeneralizacionTests.cs` | R-F-5 (punto A), D-F |
| Doctrina: fin del conteo congelado en R1-Q2, regla por rol, agregado D-C, `totOpt` generalizado, estados finales, trampa R-TRAMPA-OPC, follow-ups vivos | `.opencode/project-context.md` | R-R-4 |

### 6.3 DoD (verificación)

1. `dotnet build Remuneracion.WinForms/Remuneracion.WinForms.slnx` -> **0 warnings, 0 errors**. OK
2. Suite completa verde: **310/310** (base 307 + 3 tests del punto A). OK
3. Regresión 2026082: corre END-TO-END sin throw (`fallo is null`), DetRetri 5/5 = R10 ±0.5 post-redondeo, ASE3 = 16.369.059.896, invariantes duras presentes, deltas {-3,-9,-6,+6,+8}. OK
4. Julio bit-comparable: goldens Q1+Q2 ±0.5 verdes (A7 Q1 y `DetRetriQ2Tests` Q2); ASE5-julio `totOpt` idéntico al mapa congelado. OK
5. Sin datos inventados: todo 0 remite a ausencia evidenciada (T0a); todo valor a fuente real o R10. OK
6. Sin commits del agente; CRLF; sin emojis. OK

### 6.4 Decisiones ratificadas

D-A..D-F (§0.3). **D-F** ratificada por el Ingeniero como parte del patrón por firma (el `totOpt` generalizado absorbe la cardinalidad futura de filas `Mes` sin mapa nuevo, coherente con OCP).
