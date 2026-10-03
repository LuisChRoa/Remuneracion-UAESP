# Plan 23 — Opcionalidad de `Vlr Intereses` en SALDOS POR NOTA + RETRIBUCION NEGATIVA (HU-22, paso 2.5)

> **Alcance:** eliminar el blocker del flujo 5-ASE de agosto (2026082, HU-11/Q2 paso 2.5) haciendo **opcional puntual** el concepto `Vlr Intereses` en los readers de `SALDOS POR NOTA` y `RETRIBUCION NEGATIVA` (ausente = 0, nunca fail-fast), con conceptos core obligatorios y **cero mutación de geometría** (sin espejo, sin insert/delete, sin reanclaje).
> **Rector (IRRENUNCIABLE):** `README.md` + `Docs/Propuesta_Proyecto_Automatizacion_Remuneracion_UAESP.md` + `Docs/Detalle de plantilla.docx` + `Docs/Prompt_Maestro_Proyecto_Remuneracion_UAESP _ Vo .docx`. **Invariantes del proyecto:** NUNCA sobrescribir fórmulas (solo valores); tolerancia ±0.5; Q2 aplica SALDOS POR NOTA/RETRIBUCION NEGATIVA (`Periodo.NumeroQuincena == 2`, `AjustesSfT = 0` en Q1); redondear a entero antes de DetRetri/columna D; build 0 warnings; finales de línea CRLF; fail-fast nombrando ASE+reporte+celda/fila; tests SOLO con insumos reales (`Docs/Insumos`, `Docs/Prueba2`), NO fixtures sintéticos.
> **Continuidad:** HU-01..HU-20 cerradas; Plan 21 cerrado (espejo R1, suite 299/299, PR 4 con blocker declarado fuera de scope). Este plan construye encima de: `WorkbookLeafCellMapAjustesSfT.ConceptosSaldosPorAse/ConceptosRetribucionPorAse` (listas por `Ase.Id`), `ExcelDataReaderWorkbookLeafInputReader.LeerSaldosNotas` (L321-414) / `LeerRetribucionNegativa` (L417+), `ProcesadorPeriodo` paso 2.5 (L196-215, invocación L212), y `Regresion2026082Tests` (captura el fail-fast por nombre, L41/87). Este plan NO reabre la semántica del espejo R1 ni ninguna aritmética de negocio: **cero cambios de fórmulas; Q1/Q2-julio intactos por construcción.**
> **Fuente única de verdad:** `plans/22 - T0 Evidencia.md` — **T0 ya cerrado sin NEEDS_CONTEXT** (veredictos T0a..T0e congelados; este plan NO los re-deduce ni los contradice). El Plan 21 (`plans/21 - espejo-estructural-reporte-componentes-r1.md` + su T0) es solo antecedente: **HU-22 NO es espejo**.
> **Insumos canónicos (no se re-fijan, no se inventan datos):** plantilla canónica `Docs/Insumos/REMUNERACION 2026072/Plantilla_ Remuneracion 202607-2.xlsx`; goldens `Remuneracion 202607-1 Total.xlsx` / `Remuneracion 202607-2 Total.xlsx`; fuentes Q2-julio `Docs/Insumos/REMUNERACION 2026072/{1..5}-*/` (`SaldosaFavorAplicadosPorNotas_*`, `RetribuciónNegativa_*`); fuentes agosto `Docs/Prueba2/Insumos/{1..5}-*/` (mismo patrón, rango 01082026-31082026) + `R10_Remuneracion_2026082.xlsx`. **Sin golden de agosto de UAESP** (R-ORACULO-AGOSTO): verificación estructural/smoke únicamente.
> **Numeración:** el 22 quedó reservado por la evidencia T0 (`plans/22 - T0 Evidencia.md`); este plan toma el primer correlativo libre, **23**.
> **Estado:** PENDIENTE DE APROBACIÓN — no implementar hasta OK del Ingeniero.
> **Fecha:** 2026-10-02

---

## 0. Clarification Gate

**T0 está cerrado y la decisión de diseño ya fue tomada por el Ingeniero este turno (ver §0.3): no hay preguntas bloqueantes pendientes.** El plan viene redactado en la ruta aprobada (opcionalidad puntual de `Vlr Intereses`, ambos readers, sin espejo/insert/delete). Punto de bloqueo respetado: NO se implementa nada en este documento.

### 0.1 Qué está verificado y qué NO (honestidad técnica)

**Verificado — hechos congelados del T0 (fuente: `plans/22 - T0 Evidencia.md`, NO re-deducidos aquí):**

| # | Hecho congelado | Origen T0 | Consecuencia para el plan |
|---|---|---|---|
| V1 | Destino `SALDOS POR NOTA` = geometría FIJA por ASE (`FilasSaldosNotasPorAse`) + filas de fórmula propias (`C9=C7-I7` y análogas); template inserta columna `Especiales` en I que la fuente no trae; columnas por TÍTULO (`MapeoColumnasPorTitulo`) | T0a/T0b (§2-3) | El patrón espejo del Plan 21 NO aplica; el fix NO muta geometría, sin reanclaje |
| V2 | Fuente Q2-julio conceptos por ASE = 5/6/5/6/0; agosto = 5/**5**/5/6/0 (ASE2 sin `Vlr Intereses`, delta destino−fuente +1 en ASE2) | T0a (§2.2) | El síntoma exacto del blocker; `Vlr Intereses` es la única fila de concepto que varía (opcional evidenciado) |
| V3 | Causa raíz: `ConceptosSaldosPorAse[2]` exige `(16,"Vlr Intereses")` yendo a `BuscarFilaConcepto` = -1 con throw en `LeerSaldosNotas` (~L376), invocado desde `ProcesadorPeriodo.cs:212`, antes de escribir (`:305`) | T0d (§5.1) | Punto de cambio mínimo y único en el path SALDOS |
| V4 | `RETRIBUCION NEGATIVA` NO bloquea agosto (5/5 fuentes solo R1 con early-return 0) pero comparte el defecto latente en `ConceptosRetribucionPorAse` (ASE1-4 exigen `Vlr Intereses`) | T0d (§5.2-5.3) | El fix cubre ambos readers con la MISMA regla aunque RETRI no bloquee hoy |
| V5 | Aguas abajo: AJUSTES-SF-T (462 fórmulas) desde SALDOS + RETRI; 292 refs externas en 6 hojas (CONSOLIDADO D85:D89, DetRetri, DetValiRetri, REMUNERACION_*); 7 mapas fijan filas absolutas; ninguna fórmula referencia `Vlr Intereses` (fila de entrada pura) | T0c (§4) | Escribir `Vlr Intereses=0` no altera ninguna fórmula; mutar geometría costaría unos 432 refs + 7 mapas, prohibido |
| V6 | La aritmética cierra: agosto ASE2 trae `Componente = Vlr Servicio` (931073.72) y `Total = Componente + Subsidio` (506855.36); visible `C22=C20-I20` = 506855.36 con `Vlr Intereses=0` | T0e (§6.1) | Ausente = 0 es aritméticamente neutro |

**Re-verificado en esta planificación** (lectura directa de código, sin cambiar nada):

| # | Hecho | Método | Consecuencia |
|---|---|---|---|
| V7 | `ConceptosSaldosPorAse` ([1]=5 sin intereses, [2]=6 con `(16,"Vlr Intereses")`, [3]=5, [4]=6 con `(41,...)`, [5]=6 con `(54,...)`); `ConceptosRetribucionPorAse` ([1..4] con `Vlr Intereses` primera, [5] sin) | Lectura `WorkbookLeafCellMapAjustesSfT.cs:116-177` | Confirma V3/V4 al símbolo: el cambio toca esas dos tablas y nada más en el mapa |
| V8 | `LeerSaldosNotas` itera conceptos del mapa y lanza `CalculoInvalidoException` con ASE+concepto si `BuscarFilaConcepto < 0`; early-return 0 si sin headers; `Deb/Cred` opcional; `I{fila}=0` siempre | Lectura `ExcelDataReaderWorkbookLeafInputReader.cs:321-414` | El `foreach` L370-397 es el sitio del `if` opcional; el resto del método queda intacto |
| V9 | `LeerRetribucionNegativa` replica la estructura (early-return 0, throw equivalente L467-468) | Lectura `ExcelDataReaderWorkbookLeafInputReader.cs:417-480` | Misma regla, segundo sitio de cambio |

**NO verificado / declarado como riesgo (NO se finge evidencia):**

| # | Límite | Tratamiento en el plan |
|---|---|---|
| N1 | Sin golden/oráculo de agosto de UAESP (R-ORACULO-AGOSTO) | Test 2026082 = smoke estructural con invariante `Vlr Intereses=0` en SALDOS ASE2; sin asserts ±0.5 contra salida de agosto, sin inventar valores |
| N2 | R-EXTRA-CONCEPTO sin evidencia (ninguna fuente Q2/agosto trae concepto extra desconocido) | Se declara el comportamiento actual (iterar mapa, extra se ignora) y queda igual con evidencia de grep; sin test runtime (prohibido fixture sintético) |
| N3 | Golden Q1 con 6 filas/ASE (geometría distinta de Q2) pero sin fuentes Q1 de la cadena | Irrelevante hoy (Q1 nunca llega a 2.5); test Q1 debe seguir verde (R-GEOMETRIA-Q1) |

### 0.2 Mapeo al Rector (in vs out)

**Entra:**

| Requisito | Superficie de cambio |
|---|---|
| Opcionalidad puntual de `Vlr Intereses`: ausente = 0 explícito, nunca fail-fast; core (`Vlr Servicio`, `Componente`, `Subsidio(-)/Contribucion(+)`, `Subs/Cont`, `Total`) sigue obligatorio con fail-fast que nombra ASE+reporte | `WorkbookLeafCellMapAjustesSfT` (marca de opcionalidad en las entradas `Vlr Intereses` de ambas tablas) + `ExcelDataReaderWorkbookLeafInputReader.LeerSaldosNotas` / `LeerRetribucionNegativa` (rama tolerante solo para el concepto opcional) |
| Paridad: la MISMA regla en ambos readers (RETRI por defecto latente, T0 §5.3) | Segundo sitio de cambio, idéntica semántica; sin divergencia SALDOS vs RETRI |
| Fuente vacía (solo R1, ASE5) = 0 legítimo; `Deb/Cred` opcional; `Especiales` (I) = 0; columnas por título | Conductas existentes que se conservan intactas y se re-aseguran con regresión (no se tocan) |
| Goldens julio Q1+Q2 ±0.5 intactos + smoke 2026082 con invariante nueva + Q1 intacto (`AjustesSfT=0`) | `Remuneracion.IntegrationTests` (xUnit + FluentAssertions): regresión continua + smoke dedicado con insumos reales |
| Trazabilidad doctrinal mínima | Diff acotado en `.opencode/project-context.md` (regla de opcionalidad 2.5 + follow-up R-EXTRA-CONCEPTO) |

**Sale (EXPLÍCITO):** patrón espejo / insert-delete de filas / reanclaje de fórmulas (prohibido: §2.3); cambios a aritmética de negocio (CONSOLIDADO, AJUSTES-SF-T, DetRetri, banco, BCE, conciliaciones, INTERVENTORIA-valores); regla general "ausencia de cualquier concepto = 0" (rechazada: debilita el fail-fast de core); fail-fast nuevo para concepto extra (rechazado en esta HU: sin evidencia, amplía blast radius); reinterpretar valores de agosto (dato, no cálculo); Q1 (no se abre paso 2.5); UI WinForms y CLI; paquetes NuGet; commits (los hace el Ingeniero con `#commit`).

### 0.3 Decisiones ejecutivas (ya tomadas por el Ingeniero — aprobación = ratificar estas)

| ID | Decisión |
|---|---|
| D-A | **Opcionalidad puntual de `Vlr Intereses`** (ausente = 0) como único concepto opcional evidenciado. Conceptos core obligatorios: `Vlr Servicio`, `Componente`, `Subsidio(-)/Contribucion(+)`, `Subs/Cont`, `Total`. Recomendación T0 adoptada; regla general "todo-ausente = 0" rechazada. |
| D-B | **Cero mutación de geometría:** sin espejo, sin insert/delete, sin reanclar fórmulas. `FilasSaldosNotasPorAse` / `FilasRetribucionNegativaPorAse` y las 7 tablas de filas absolutas quedan intactas. |
| D-C | **Alcance = ambos readers** (SALDOS y RETRIBUCION) con la MISMA regla, aunque RETRIBUCION no bloquee hoy (defecto latente T0 §5.3). |
| D-D | **Plantilla canónica** = `Docs/Insumos/REMUNERACION 2026072/Plantilla_ Remuneracion 202607-2.xlsx`. El nombre citado en el comentario del mapa ("Plantilla 8 agos 2026 …") ya no existe en disco y no gobierna. |
| D-E | **Test 2026082 = smoke con invariante `Vlr Intereses=0` en SALDOS ASE2** (sin golden de agosto, sin inventar valores). Rama end-to-end completa: aserta fin del blocker 2.5 + invariantes R1 ya capturadas (delta = {-3,-9,-6,+6,+8}) sin regresión. |
| D-F | **Q1 intacto:** no se abre paso 2.5 en Q1 (sin fuentes; `AjustesSfT = 0`); test Q1 sigue verde. |

---

## 1. PROPOSE

### 1.1 Intent

Que el período 2026082 deje de detenerse en el paso 2.5: cuando la fuente `SaldosaFavorAplicadosPorNotas` (o futuras `RetribuciónNegativa` con headers) no traiga la fila `Vlr Intereses`, el lector la trate como **0 explícito** en la fila fija del template en vez de lanzar, manteniendo obligatorio todo concepto core, sin mover una sola fila ni tocar una sola fórmula, y demostrándolo con los goldens de julio intactos más el smoke de agosto con la invariante nueva.

### 1.2 In Scope

- Marca de opcionalidad en `ConceptosSaldosPorAse` / `ConceptosRetribucionPorAse` solo para entradas `Vlr Intereses` (forma exacta la fija la implementación dentro de D-A/D-B; §2.1 recomienda la mínima).
- Rama tolerante en `LeerSaldosNotas` + `LeerRetribucionNegativa`: concepto opcional ausente con fila en 0 explícito; concepto core ausente con fail-fast actual intacto (nombra ASE+reporte).
- Conservación verificada: early-return fuente vacía = 0, `Deb/Cred` opcional, `Especiales` I = 0, columnas por título, Q1 sin paso 2.5.
- Tests con insumos reales: goldens Capa A Q1+Q2 (regresión ±0.5) + smoke 2026082 (invariante `Vlr Intereses=0` SALDOS ASE2, RETRI = 0) + Q1 verde.
- Diff acotado en `.opencode/project-context.md` (regla 2.5 + R-EXTRA-CONCEPTO como follow-up).

### 1.3 Non-Goals

Ver §0.2 "Sale". En particular: NO se agrega ningún concepto nuevo al mapa sin evidencia; NO se emite valor distinto de 0 para un concepto ausente (prohibido inventar intereses); NO se silencia el fail-fast de core; NO se cambia el mapeo de columnas por título; NO se toca la fila visible `Cn-In` ni el puente AJUSTES-SF-T ni nada aguas abajo.

### 1.4 Riesgos

| ID | Riesgo | Mitigación |
|---|---|---|
| R-EXTRA-CONCEPTO (declarado por T0) | Concepto en fuente que el mapa no conoce y hoy se ignora silencioso (el lector itera el mapa, no la fuente). Si UAESP agrega un concepto con valor real, se perdería sin aviso | En esta HU queda IGUAL con evidencia (grep + §2.1/§3.1 declarados); se registra como follow-up explícito (§5). No se agrega fail-fast nuevo sin evidencia para no ampliar blast radius |
| R-GEOMETRIA-Q1 (declarado por T0) | Golden Q1 tiene 6 filas/ASE (geometría distinta de Q2); un cambio descuidado en el mapa podría romper Q1 aunque Q1 no lea la cadena | Q1 no entra al código cambiado en tiempo de ejecución (paso 2.5 solo Q2); test Q1 verde es gate explícito de cada PR (DoD-3) |
| R-ORACULO-AGOSTO (declarado por T0) | Sin golden de UAESP para agosto: el smoke no puede certificar ±0.5 de negocio | Smoke estructural únicamente (flujo completo + invariante 0 + invariantes R1); sin asserts de valores de negocio de agosto; prohibido inventar golden |
| R-REGRESION-JULIO | La rama tolerante altera accidentalmente el path julio (donde `Vlr Intereses` SÍ está y debe seguir leyéndose) | Goldens Q1+Q2 ±0.5 como red ciega en cada tarea; la rama solo se activa con `filaFuente < 0` + concepto marcado opcional; julio no toca la rama (valores reales escritos) |
| R-SOBREDISENO | Convertir la opcionalidad puntual en framework general de conceptos opcionales | Alcance congelado a `Vlr Intereses` (D-A); cualquier otro concepto opcional futuro requiere su propio T0/HU |

---

## 2. DESIGN

### 2.1 Enfoque: concepto opcional evidenciado, geometría fija

**Modelo.** Las tablas `ConceptosSaldosPorAse` / `ConceptosRetribucionPorAse` dejan de ser "todas obligatorias": cada entrada pasa a llevar su obligatoriedad. Implementación mínima recomendada (equivalente a la alternativa (b) del T0 §6.1, pero con menor diff): marcar `Opcional` las entradas `Vlr Intereses` existentes (SALDOS: ASE 2/4/5; RETRI: ASE 1-4) sin mover filas ni renombrar conceptos. La firma exacta (tupla con flag, enum, o tabla paralela de opcionales) la elige el implementador; condición: el mapa sigue siendo explícito por `Ase.Id` y las filas del template no cambian.

**Lectura (único cambio de comportamiento).** En el `foreach` de conceptos de `LeerSaldosNotas` (hoy L370-397) y su gemelo de `LeerRetribucionNegativa` (L462+):

1. `filaFuente = BuscarFilaConcepto(filas, concepto)` como hoy.
2. Si `filaFuente < 0` y el concepto está marcado opcional, emitir la fila completa en **0 explícito** (todas las columnas del `MapeoColumnasPorTitulo` + `I{fila}=0`; columna `O/DebCred` se omite como hoy cuando la fuente no la trae, conservando el 0 de plantilla) y continuar sin lanzar. Determinista y testeable sin depender del estado de la plantilla.
3. Si `filaFuente < 0` y el concepto es core, throw actual intacto (`ASE {id}: no se encontró la fila de concepto '{c}' en la fuente SALDOS POR NOTA / RETRIBUCION NEGATIVA.`).
4. Si `filaFuente >= 0` (caso julio), path idéntico al actual, valor real leído por título de columna. `Total`/`ServEspK` se derivan como hoy (la fila `Total` es core y siempre está en fuentes no vacías).

**Por qué 0 explícito y no "no escribir":** ambas son aritméticamente equivalentes (T0 V6: ninguna fórmula referencia la fila; el visible no se mueve), pero el 0 explícito hace la invariante `Vlr Intereses=0` asertable a nivel de `SaldosNotasAseInputs.Celdas` sin depender de los ceros preexistentes de la plantilla. Costo: cero (mismo diccionario, mismos writers).

**Escritura: sin cambios.** El writer sigue escribiendo valores a celdas fijas con guard anti-fórmula vigente; los 0 explícitos son valores como cualquier otro. Ninguna celda protegida (`Protegidas`: visibles, AJUSTES D9:D13/D28:D32/D47:D51, CONSOLIDADO D85:D89, INTERVENTORIA, ANT EXT-REV) se toca.

**R-EXTRA-CONCEPTO (comportamiento declarado, sin cambio):** el lector itera el mapa, no la fuente; un concepto extra en la fuente se ignora. Queda igual en esta HU por ausencia de evidencia y prohibición de fixtures sintéticos (no hay cómo testearlo con insumos reales). Evidencia de cierre: grep que demuestre dirección de iteración + revisión.

### 2.2 Alternativas y trade-offs

| Eje | A — Marcar `Opcional` en el mapa (recomendada) | B — Bucle dirigido por la fuente (alternativa T0 §6.1) |
|---|---|---|
| Qué se hace | Las entradas `Vlr Intereses` llevan flag opcional; el `foreach` tolera solo esas | Se invierte el bucle: por cada fila de concepto hallada en la fuente se resuelve su fila fija; lo no hallado y no-core se omite/0 |
| Diff | Mínimo: 2 tablas + 2 ramas `if` | Medio: reescribe el bucle de ambos readers; más riesgo de tocar el path core |
| Fail-fast de core | Intacto por construcción (default = obligatorio) | Intacto solo si se lista core explícito (misma tabla, invertida) |
| R-EXTRA-CONCEPTO | Igual en ambas (se ignora; el mapa sigue mandando en filas destino) | Podría detectarlo (la fuente manda), pero sin evidencia no hay spec que lo exija |
| Veredicto | **Principal (D-A/D-B):** menor blast radius, simétrica en ambos readers | Descartada salvo que la implementación demuestre que A no expresa la regla sin ensuciar el mapa |

Descartadas explícitamente: espejo/insert/delete/reanclaje (costo aprox. 432 refs + 7 mapas, T0 §4.4); regla general "todo-ausente = 0" (debilita fail-fast, D-A); 0 silencioso para core (prohibido, T0 §6.2); fabricar valor de intereses (prohibido, T0 §6.2).

### 2.3 Por qué NO hay desplazamiento ni reanclaje (tratamiento explícito)

Porque no se mueve ninguna fila: las filas destino (`3..7`, `15..20`, `28..32`, `40..45`, `53..58` en SALDOS; análogas en RETRI) conservan dirección, longitud de bloque y fórmulas vecinas. Las 462 fórmulas de AJUSTES-SF-T, los 70+70 refs SALDOS/RETRI, los 292 refs aguas abajo y los 7 mapas de filas absolutas quedan idénticos en anclaje. La prueba dura: el inventario T0c (§4) se usa como checklist de no-toque (grep de que ningún archivo de esos mapas se modifica en el diff, salvo las dos tablas de conceptos del §2.1).

### 2.4 Contratos afectados (resumen)

| Contrato | Cambio |
|---|---|
| `WorkbookLeafCellMapAjustesSfT` (Infrastructure): `ConceptosSaldosPorAse`, `ConceptosRetribucionPorAse` | Solo marca de opcionalidad en entradas `Vlr Intereses`; filas, conceptos, `Filas*`, `FilaTotal*`, `MapeoColumnasPorTitulo`, `Protegidas`, `ColumnaTemplateDebCred` intactos |
| `IWorkbookLeafInputReader` con `ExcelDataReaderWorkbookLeafInputReader.LeerSaldosNotas` / `LeerRetribucionNegativa` | Rama tolerante para opcional ausente (= 0 explícito); core y vacía y headers y Deb/Cred intactos; mensajes de throw intactos |
| `ProcesadorPeriodo` paso 2.5, `OpenXmlPlantillaWriter`, `IValidador`, R10 | Sin cambios (el paso 2.5 deja de lanzar para agosto por el lector, no por orquestación) |
| `Remuneracion.IntegrationTests` (xUnit + FluentAssertions) | Smoke 2026082 extendido con invariante nueva + regresión Q1/Q2; suite base 299/299 como red ciega |
| `.opencode/project-context.md` | Diff acotado: regla de opcionalidad 2.5 + follow-up R-EXTRA-CONCEPTO |

### 2.5 Architecture Validation Certificate

| Principio | Veredicto |
|---|---|
| **SRP** | ✅ Una sola razón de cambio por costura: el mapa declara obligatoriedad, el lector la interpreta, los tests la demuestran. Sin ramas `if agosto` ni `if período` en el lector (la regla es por concepto, válida en todo Q2 presente y futuro). |
| **OCP** | ✅ La tabla de conceptos se extiende con un flag en vez de bifurcarse por período: el próximo concepto variable (si T0 futuro lo evidencia) se marca, no se duplica el reader. |
| **DIP** | ✅ Cambios detrás de `IWorkbookLeafInputReader`; `ProcesadorPeriodo` y el writer no se tocan ni conocen la regla. |
| **Best practices** | ✅ Fail-fast preservado donde protege (core nombra ASE+reporte); tolerancia solo donde la evidencia la autoriza (único concepto variable en 15 fuentes reales); fórmulas nunca sobrescritas; columnas por título; tolerancia ±0.5; Q1 excluido por dominio (`NumeroQuincena`), no por parche. |
| **Performance** | ✅ Sin impacto: mismo número de lecturas; la rama tolerante evita un throw, no agrega I/O; tests reutilizan `Remuneracion.IntegrationTests` con insumos reales existentes. |

---

## 3. SPEC

### 3.1 Requirements

| ID | Requisito | Grupo | Criterio de aceptación |
|---|---|---|---|
| R-F-1 | `Vlr Intereses` opcional en SALDOS: si `BuscarFilaConcepto` < 0 para esa entrada, fila en 0 explícito, sin throw | Fix | Smoke agosto ASE2: `SaldosNotasAseInputs.Celdas` de la fila 16 en 0; flujo 2.5 no lanza |
| R-F-2 | `Vlr Intereses` opcional en RETRIBUCION con la MISMA regla (ASE1-4) | Fix | Rama simétrica + revisión; goldens RETRI=0 intactos; grep de simetría entre ambos readers (sin fuente real que ejercite el caso, prohibido sintético) |
| R-F-3 | Core obligatorio intacto: `Vlr Servicio`, `Componente`, `Subsidio(-)/Contribucion(+)`, `Subs/Cont`, `Total` ausente con fail-fast que nombra ASE+reporte (mensajes actuales) | Fix | Sin cambio de mensajes; `grep "no se encontró la fila de concepto"` sigue en ambos readers |
| R-F-4 | Julio intacto: `Vlr Intereses` presente y se lee el valor real (p. ej. ASE2-Q2 = 172.61), no 0 | Fix | Goldens Q2 ASE2/ASE4 con intereses reales ±0.5 |
| R-F-5 | Conservación: fuente vacía = 0 (ASE5), `Deb/Cred` opcional, `Especiales` I = 0, columnas por título, Q1 sin paso 2.5 (`AjustesSfT=0`) | Conservación | Goldens Q1+Q2 ±0.5; Q1 verde; ningún cambio en esos branches (diff lo demuestra) |
| R-F-6 | Cero mutación geométrica: filas fijas, fórmulas, puentes y mapas de anclaje intactos | Conservación | Diff no toca `Filas*`/`FilaTotal*`/`MapeoColumnasPorTitulo`/`Protegidas`; checklist T0c verificado |
| R-R-1 | Suite base 299/299 verde antes y después (red ciega, `dotnet build` 0 warnings) | Transversal | Build + suite completos en cada PR |
| R-R-2 | Smoke 2026082: flujo 5-ASE con `Docs/Prueba2/Insumos` completa el paso 2.5 y genera workbook; invariante `Vlr Intereses=0` en SALDOS ASE2; RETRI = 0; invariantes R1 (delta = {-3,-9,-6,+6,+8}) sin regresión | Transversal | Test dedicado; sin asserts de negocio ±0.5 sobre agosto (R-ORACULO-AGOSTO) |
| R-R-3 | R-EXTRA-CONCEPTO declarado: documentado que extra-se-ignora queda igual + follow-up | Transversal | §2.1 + §5 + entrada en `project-context.md`; grep de dirección de iteración |
| R-R-4 | `project-context.md` con la regla de opcionalidad 2.5 (diff acotado, sin código en el mismo commit) | Transversal | Diff solo doctrina/reglas |

### 3.2 Scenarios (Given/When/Then)

- **S1 (agosto-ASE2, el blocker):** Given fuente agosto ASE2 sin `Vlr Intereses` (5 conceptos, `Docs/Prueba2/Insumos/2-Lime/`), When `LeerSaldosNotas(ase2, ruta)`, Then no lanza y la fila 16 queda en 0 explícito; `Total` = 506855.36 (fila `Total`, no afectado por la ausencia).
- **S2 (julio-ASE2, no-regresión):** Given fuente Q2-julio ASE2 con `Vlr Intereses=172.61`, When se lee, Then la fila 16 trae los valores reales por título de columna (no 0) y el golden Q2 cierra ±0.5.
- **S3 (core-ausente, fail-fast intacto):** Given fuente con headers pero sin un concepto core (acreditado por rama simétrica + mensajes intactos; sin sintéticos), When se lee, Then `CalculoInvalidoException` nombra ASE+reporte+concepto.
- **S4 (RETRI-agosto, conservación):** Given las 5 fuentes RETRI de agosto (solo R1), When se leen, Then early-return 0 en las 5; golden AJUSTES D28:D32 = 0 intacto.
- **S5 (ASE5-vacía, conservación):** Given fuente SALDOS ASE5 (solo R1) en Q2 y agosto, When se lee, Then 0 legítimo sin throw, ambos períodos.
- **S6 (smoke-2026082):** Given `Docs/Prueba2/Insumos` (5 ASE + R10), When `ProcesadorPeriodo` corre Q2 completo, Then supera el paso 2.5 (L212), genera workbook con fórmulas preservadas, `Vlr Intereses=0` en SALDOS ASE2, y las invariantes R1 delta = {-3,-9,-6,+6,+8} se mantienen.
- **S7 (Q1-intacto):** Given período Q1, When se procesa, Then nunca entra a 2.5, `AjustesSfT = 0`, goldens Q1 ±0.5 verdes.
- **S8 (cero-geometría):** Given el diff de la HU, When se audita contra el inventario T0c, Then ninguna fila fija, fórmula, puente (AJUSTES D9:D13/D28:D32/D47:D51/D52, CONSOLIDADO D85:D89) o mapa de anclaje cambió salvo las dos tablas de conceptos del §2.1.

---

## 4. TASKS

| ID | Tarea | Grupo | Depende de | Verificación |
|---|---|---|---|---|
| T0 | **CERRADO — Discovery evidencia (NO es PR nuevo):** veredictos T0a..T0e en `plans/22 - T0 Evidencia.md` (dumps, geometría NO-espejo, inventario 462/70+70/292/7 mapas, causa raíz, veredicto de diseño + riesgos + invenciones prohibidas) | Fase 0 | — | DONE sin NEEDS_CONTEXT; este plan lo consume como hechos congelados |
| T1 | Mapa: marcar opcionalidad puntual de `Vlr Intereses` en `ConceptosSaldosPorAse` (ASE 2/4/5) y `ConceptosRetribucionPorAse` (ASE 1-4); filas y conceptos intactos; resto del mapa intacto | Fix | T0 | Compila; `grep ConceptosSaldosPorAse/ConceptosRetribucionPorAse` acotado a esas entradas; S8 (cero-geometría) |
| T2 | Lector: rama tolerante en `LeerSaldosNotas` (opcional ausente con 0 explícito; core ausente con throw actual) + MISMA rama en `LeerRetribucionNegativa`; vacía/DebCred/Especiales/títulos intactos | Fix | T1 | S1/S2/S3/S4/S5; `grep "no se encontró la fila de concepto"` en ambos readers; build 0 warnings |
| T3 | Tests con insumos reales: (a) regresión goldens Capa A Q1+Q2 ±0.5 (suite 299/299 verde); (b) smoke 2026082 5-ASE con invariante `Vlr Intereses=0` SALDOS ASE2 + RETRI=0 + invariantes R1; (c) Q1 verde con `AjustesSfT=0` | Transversal | T2 | R-R-1/R-R-2 + S6/S7; `dotnet build Remuneracion.WinForms/Remuneracion.WinForms.slnx` 0 warnings; suite completa verde |
| T4 | Docs: `project-context.md` (regla opcionalidad 2.5: único opcional evidenciado, core obligatorio, R-EXTRA-CONCEPTO como follow-up) | Transversal | T3 | Diff acotado a doctrina/reglas; sin código en el mismo commit; S8 checklist citado |

**Orden sugerido:** T0 (cerrado) → T1 → T2 → T3 → T4. T3 corre como regresión continua después de T1 y T2 (julio verde antes y después de cada tarea). Build 0 warnings antes de marcar cada tarea completa. Sin commits (los hace el Ingeniero). **Work units encadenadas en PRs (Fase 0 ya hecha, no es PR):** WU-1 = T1+T2 (fix mapa+lector); WU-2 = T3+T4 (tests + docs). PRs secuenciales, cada uno con gates DoD 1-2.

---

## 5. Resumen ejecutivo, riesgos y decisiones

- **Resumen:** fix mínimo y simétrico para el blocker de agosto: el mapa exigía `Vlr Intereses` en ASE2 y la fuente real no la trae (delta +1). La solución aprobada (D-A..D-F) tolera ese único concepto como 0 explícito en ambos readers, conserva core obligatorio con fail-fast, y no mueve filas ni fórmulas (checklist T0c: 462 + 70+70 + 292 refs y 7 mapas intactos). Julio (con intereses reales) y Q1 (sin paso 2.5) quedan intactos por construcción; agosto se acredita con smoke estructural porque UAESP no entregó golden.
- **Riesgo principal:** R-ORACULO-AGOSTO (sin golden, solo smoke) + R-EXTRA-CONCEPTO (extra-se-ignora queda igual, follow-up registrado) + R-GEOMETRIA-Q1 (test Q1 verde como guardián). Contenidos por gates y por la prohibición de inventar datos/valores.
- **Decisión para el Ingeniero:** ratificar D-A..D-F (§0.3). Aprobación del plan = aceptación de esas seis decisiones. **Sin preguntas bloqueantes pendientes.**

### DoD (Definition of Done)

1. `dotnet build Remuneracion.WinForms/Remuneracion.WinForms.slnx` con 0 warnings.
2. Suite completa verde sin regresión (base 299/299) + goldens Capa A Q1+Q2 en dif ±0.5 intactos.
3. Smoke 2026082 verde: supera el paso 2.5, genera workbook con fórmulas preservadas, invariante `Vlr Intereses=0` en SALDOS ASE2, RETRI = 0, invariantes R1 sin regresión; Q1 verde con `AjustesSfT=0`.
4. Ninguna fórmula sobrescrita (guard anti-fórmula + checklist S8 verdes); diff geométrico acotado a las dos tablas de conceptos (§2.1).
5. `project-context.md` actualizado (regla 2.5 + follow-up R-EXTRA-CONCEPTO); sin datos inventados (toda celda nueva remite a fuente real o a 0 explícito por ausencia evidenciada).
6. Sin commits del agente (los hace el Ingeniero con `#commit`); finales de línea CRLF; sin emojis.

### Follow-ups explícitos (fuera de esta HU)

- R-EXTRA-CONCEPTO: si UAESP introduce un concepto nuevo con valor, el lector actual lo ignorará en silencio. Requiere su propio T0 (con fuente real que lo evidencie) para decidir entre fail-fast de extra-desconocido vs absorción al mapa.
- R4 espejo pendiente del Plan 21 (motor solo R1) y R4-por-empresa Q2 (recorte HU-20/G2-D1): intactos, no tocados por esta HU.
