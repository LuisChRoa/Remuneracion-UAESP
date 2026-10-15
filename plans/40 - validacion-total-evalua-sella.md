# Plan 40 — W-1 evalúa + writer sella: fin del caché stale como oráculo (C15) + F-NEW-1

> **Alcance:** cerrar el incidente real medido 2026-10-14 — la corrida end-to-end sobre la PLANTILLA OFICIAL para 2026072 (Julio-2, 5 ASE, Q2) produjo un archivo válido pero terminó en `ERR-VALIDACION` para los 5 ASE porque el gate W-1 lee cachés `<v>` stale heredados de la plantilla en vez de verificar nada. El gate W-1 (`ValidadorBasico` + `ValidacionOracleReader`) deja de leer cachés y **evalúa** los 5 sub-bloques booleanos de `VALIDACION_TOTAL` desde los valores realmente escritos en la salida (con assert previo del texto `<f>` esperado); el writer **sella** los `<v>` computables por read-back de lo efectivamente escrito (jamás `<f>`); y la familia F-NEW-1 del Plan 39-T3 recibe su propio T0 con regla honesta por celda (reescritura Phase 0 por quincena o invariante declarado — prohibido el sellado cosmético). Cambio de producción esperado: **mínimo y acotado** (un evaluador puro en Core + un sellado values-only en Infrastructure + asserts de texto-`<f>`; cero cambios de semántica de negocio, cero cambios de lectura R1/R2/R4, cálculo CONSOLIDADO, naming, preflight, sellos, calcChain, sinonimia).
> **Rector (IRRENUNCIABLE):** `README.md` + `Docs/Propuesta_Proyecto_Automatizacion_Remuneracion_UAESP.md` + `Docs/Detalle de plantilla.docx` + `Docs/Proceso de Recaudo.docx` + `Docs/Prompt_Maestro_Proyecto_Remuneracion_UAESP _ Vo .docx` + `.opencode/project-context.md`. **Invariantes:** tolerancia ±0.5 en goldens/comparaciones numéricas; NUNCA sobrescribir fórmulas — solo valores (`<v>`); NUNCA `InsertRow`/`DeleteRow` en procesamiento; **NUNCA usar cachés `<v>` como oráculo de fórmulas — el gate evalúa y el writer sella (doctrina ratificada, §0.4 D-A)**; doctrina calcChain Plan 28 intacta (delete calcChain + fullCalcOnLoad); slot doctrine Plan 36 intacta en procesamiento; Phase 0 Plan 39 intacta (allow-list 16 celdas, hoy solo tokens de período + L9:L13); D-H (cero literales/ramas por período en el diff de producción — los períodos llegan por parámetro) + catálogo abierto (cero ramas por empresa fuera de `SinonimosEmpresaR1`/`EmpresaFacturacion`); el archivo OFICIAL nunca se modifica (hash SHA256 antes/después); build 0 warnings antes de marcar cualquier tarea completa; CRLF sin mezclar; NUNCA git commit/push (los hace el Ingeniero con `#commit`/`#push`); la implementación la ejecuta `implementer` usando SOLO este plan aprobado; `code-reviewer` verifica spec compliance + 0 warnings. **Quincena por dominio, nunca por contenido.**
> **Mandato (decisión del Ingeniero 2026-10-14 — "evaluar + sellar", en este orden de obligación):** (1) **EVALUAR (verificación real):** el gate W-1 deja de leer cachés y evalúa el booleano desde las celdas realmente escritas en el workbook de salida; antes de evaluar, aserta el TEXTO `<f>` esperado de cada sub-bloque y de sus dependencias (si la plantilla cambia la fórmula, el gate falla ruidosamente en vez de verificar la fórmula equivocada); sigue nombrando `ASE n · VALIDACION_TOTAL · <celda>` y sigue siendo fail-fast. (2) **SELLAR (artefacto autoconsistente):** el writer escribe los `<v>` de la cadena y de los sub-bloques que pueda computar, calculados por READ-BACK de los valores efectivamente escritos (nunca desde la intención del dominio), de modo que el caché no pueda mentir y Excel recalcule al abrir coincidiendo. (3) **F-NEW-1 "de paso" (misma familia):** los auto-chequeos que hardcodean `INTERVENTORIA!M26:M31` reciben su propio T0 que determina la regla honesta por celda (reescritura Phase 0 por quincena con allow-list ampliada, o invariante declarado con motivo) — prohibido sellar TRUE donde la divergencia es real.
> **Continuidad:** Planes 21 (espejo — vigente), 25 (roles por firma — vigente), 26/27 (preflight + firma + hoja-por-nombre — vigentes), 28 (calcChain + sello + `fullCalcOnLoad` — vigente; su saneamiento es lo que hace válido el archivo entregado pese al gate rojo), 29 (paridad + comparador f-vs-f — vigente; su T0f localizó el culpable `F10` y su §2.1 fijó el método "igualdad de texto-fórmula + literales, no caché" que ESTE plan convierte en doctrina del gate), 30 (naming dinámico + `NombresHojaPeriodo` — vigente), 31 (gate workbook-vs-dominio + veredictos R-D-1..R-D-7 — vigentes; su lección "un gate debe leer el WORKBOOK generado" es la que W-1 violaba), 33 (shared-preservation — intacta), 34 (sinonimia como DATOS + D-H + catálogo abierto — vigentes), 36 (slot doctrine + writer values-only — vigente; el sellado de ESTE plan es values-only y la extiende, no la excepciona), 38 (recongelamiento geométrico + drift gate — vigente), 39 (plantilla oficial directa + Phase 0 + R-INTERV-Q + F-NEW-1 pendiente — vigente; su T3 `Docs/T0-Plan39-Fase0-L9L13.md` §6 es el insumo de F-NEW-1 y su allow-list de 16 celdas es la que T3 de ESTE plan puede proponer ampliar). Este plan NO reabre lectura R1/R2/R4-detalle, cálculo CONSOLIDADO, preflight de insumos, sellos G7/K7/D6/D7, naming, calcChain, paridad, sinonimia, ni el writer values-only fuera del sellado declarado.
> **Insumos canónicos:** plantilla oficial `Docs/Plantilla estandarizada/Formato Remuneracion 2026xx-x Total.xlsx` (oráculo; NUNCA se modifica — hash SHA256 antes/después); salida real de hoy `Docs/Prueba Julio-2/Resultado/Resultado5/Remuneración 202607-2 Total.xlsx` (incidente vivo); salida legacy `Resultado4` (caso de falso-verde por herencia); `plans/29-T0-Evidencia-B.md` (§T0f y línea 30 V8 + línea 125 T0f); `plans/14 - HU-14-robustez-observabilidad.md` (§2.5 W-1); `Docs/T0-Plan39-Fase0-L9L13.md` §6 (sweep de quincena + F-NEW-1); `Remuneracion.Core/Services/ValidadorBasico.cs` (gate W-1, L~410-419); `Remuneracion.Core/Models/ValidacionCruzadaSnapshot.cs` (`SubBloquesValidacionTotal`); `Remuneracion.Infrastructure/Excel/ValidacionOracleReader.cs` (lector-oráculo, L~130-138 + `LeerValorBooleano` L229); `Remuneracion.Infrastructure/Excel/WorkbookLeafCellMapValidaciones.cs` (`SubBloquesValidacionTotal` = ["C15","D25","O25","D34","F34"], L51); `Remuneracion.Infrastructure/Excel/OpenXmlPlantillaWriter.cs` (`EscribirControlRecaudoTotal` L1642-1663, `EscribirValorNumerico` L1803-1819 con guard anti-fórmula L1808-1815); `Remuneracion.Infrastructure/Excel/PreparadorBasePipeline.cs` (Phase 0, allow-list 16 celdas); `Remuneracion.IntegrationTests/TestResults/p39-final.trx` (línea base 36 Failed / 782 Passed / 818 Total). Regens SIEMPRE a rutas temp FRESCAS. ENTORNO: si la plantilla oficial está abierta en Excel, la suite infla con fallos de bloqueo `FileShare.Read` — no es defecto del código.
> **Numeración:** `plans/` 01..39 ocupados; este plan toma el primer correlativo libre, **40**.
> **Estado:** APROBADO por el Ingeniero 2026-10-14 (Q1 = "Sí, proceder": W-1 evaluador con assert de texto-`<f>` + sellado `<v>` por read-back + F-NEW-1 con T0 propio y regla honesta; D-A..D-G ratificadas; sellado cosmético prohibido). EN EJECUCIÓN.
> **Fecha:** 2026-10-14

---

## 0. Clarification Gate

**Incidente adoptado, no re-diagnosticado (punteros del encargo; T0 congela sus propios números desde el disco).** El planificador verificó read-only en código (2026-10-14): `ValidadorBasico.cs` L410-419 itera `snapshot.SubBloquesValidacionTotal` y falla con `ASE {id} · VALIDACION_TOTAL · {celda}: sub-bloque booleano es falso en el workbook (debe ser TRUE exacto; W-1)`; `ValidacionOracleReader.cs` L130-138 puebla esos booleanos vía `LeerValorBooleano` (L229) sobre la celda del workbook de salida; `WorkbookLeafCellMapValidaciones.cs` L51 fija la lista `["C15","D25","O25","D34","F34"]`; `OpenXmlPlantillaWriter.cs` L1642-1663 (`EscribirControlRecaudoTotal`) puebla `'Valida - Control Recaudo'!F10` como literal desde los `Recaudo *` ya leídos, y L1803-1819 (`EscribirValorNumerico`) tiene el guard anti-fórmula intacto (celda-fórmula → `ERR-PLANTILLA`, nunca sobrescritura). **Lo que T0 debe aún congelar desde bytes (no verificado por este plan):** el texto exacto de cada `<f>` de los 5 sub-bloques + sus dependencias en la plantilla oficial, cada caché `<v>` del oficial y de la salida real de hoy celda-a-celda, el diferencial oficial→salida de TODAS las celdas de fórmula con caché viejo (la clase, no solo C15), y la allow-list Phase 0 actual (16 celdas) + propuesta de ampliación para F-NEW-1.

### 0.1 Evidencia adoptada (punteros — el fixture T0 es la autoridad)

| # | Afirmación adoptada | Reutilización en este plan | Re-verificación (el test, no la mano) |
|---|---|---|---|
| E1 | Cadena C15 medida: `VALIDACION_TOTAL!C15 = <f>C14=C9</f> t="b"`; `C14 = <f>'Valida - Control Recaudo'!F10</f>`; `C9 = <f>SUM(C3:C8)</f>`; `C3:C8` = suma de las 5 hojas `Recaudo *` en `F21:F26`. Plantilla OFICIAL: `F10` = **73981716599** (literal residual del ejemplo base 2026092), `C9` cache = 0 → `C15` cache = **0 (FALSE)**. `calcPr calcId="191029"` SIN fullCalcOnLoad; calcChain presente. | T0 congela texto-`<f>` + dependencias + cachés del oficial; T1 aserta ese texto antes de evaluar (R-E-2). | T0: fixture celda-a-celda desde bytes (BCL ZIP+XML); T-gates: assert de texto-`<f>` byte-idéntico al fixture, tercera forma = FAIL. |
| E2 | Salida de hoy (Resultado5, 2026072 Q2): la app SÍ escribió `F10 = 72497949967` (idéntico dígito a dígito al manual, `plans/29-T0-Evidencia-B.md` línea 125), pero `C14` cache = 73981716599, `C9` cache = 0, `C15` cache = 0 quedaron VIEJOS (heredados). `calcPr fullCalcOnLoad="1"`, sin calcChain → Excel recalcula al abrir y C15 queda TRUE. **El archivo entregado es válido; el gate es el que falla.** | Doctrina D-A/D-B (§0.4): el archivo es autoconsistente post-recálculo; el defecto está 100% en el gate. T2 sella para que además sea autoconsistente en caché. | T4: regresión end-to-end desde el OFICIAL termina en ÉXITO (exit 0 / sin ERR-VALIDACION) + `C15` y los 4 sub-bloques consistentes; healthy-open manual como aceptación final. |
| E3 | Salida legacy (Resultado4): `C14` cache 0, `C9` cache 0, `C15` cache **1** → el gate pasaba POR HERENCIA (esa base había sido recalculada por Excel alguna vez). Falso-verde confirmado. | R-E-4 / R-S-3: los tests prueban que un FALSE REAL se detecta (nunca más un verde heredado). | T1: test con FALSE genuino (divergencia real inyectada) → gate FAIL nombrando `ASE n · VALIDACION_TOTAL · <celda>`; test con TRUE evaluado (no heredado) → PASS. |
| E4 | Barrido COMPLETO de gateados en el oficial: **C15 es el ÚNICO en FALSE**. En TRUE por herencia: `VALIDACION_TOTAL` D25/O25/D34/F34 y O3:P9; `VALIDACION_*` ×5 O/P filas 3..7; `DetValiRetri2026092` D24:D29; `DetRetri2026092` D32:D36. Fórmulas observadas: `D25=<f>D24=C9</f>`, `O25=<f>O24=N9</f>`, `D34=<f>D33='CONSOLIDADO_TOTAL RECAUDO'!D109</f>`, `F34=<f>F33='CONSOLIDADO_TOTAL RECAUDO'!D109+'CONSOLIDADO_TOTAL RECAUDO'!N109</f>`. Los 4 sub-bloques restantes son falsos-verdes latentes. | T0 congela los 5 (texto + dependencias + caché); T1 evalúa los 5; T5 da veredicto por celda de TODA la clase. | T5: tabla celda → (se evalúa / se sella / invariante / requiere decisión); cero supuestos silenciosos. |
| E5 | Doctrina Plan 29 (`29-T0-Evidencia-B.md` §T0f y línea 30 V8): "cachés stale no prueban nada; evaluar por igualdad de texto-fórmula + literales (método §2.1), no por caché". El gate W-1 la contradice hoy. | D-A la eleva a invariante del gate y del writer (§0.4). | T-gates: ningún `<v>` como oráculo en el diff (grep + tests). |
| E6 | F-NEW-1 (`Docs/T0-Plan39-Fase0-L9L13.md` §6): `DetRetri{per}!L32:L36`, `DetRetri{per}!O32:O36`, `DetValiRetri{per}!I24:I29`, `INTERVENTORIA!S26:S31` hardcodean `M26:M31` y evaluarían FALSE en Q1 tras R-INTERV-Q. Aquí sellar TRUE sería MENTIR (la divergencia es real). | T3 con su propio T0: regla honesta por celda (Phase 0 reescribe `<f>` por quincena con allow-list ampliada, o invariante declarado con motivo). | T3: veredicto citado por celda; si reescritura → `<f>`-idéntica-a-variante por quincena; si invariante → motivo versionado. |
| E7 | Línea base de suite: **36 Failed / 782 Passed / 818 Total** (`p39-final.trx`); los 36 rojos incluyen el clúster C15. | Puerta de regresión: set-diff contra 36/782/818 por tarea (T-gates). | T-gates: focused por tarea + suite con set-diff; ningún rojo nuevo sin dependencia declarada. |

### 0.2 Gate de entrada — pregunta única al Ingeniero (BLOQUEANTE para T1..Tn)

| ID | Pregunta | Efecto según respuesta |
|---|---|---|
| Q1 | ¿Ratifica que (a) el gate W-1 deja de leer cachés `<v>` y pasa a **evaluar** los 5 sub-bloques desde los valores escritos (con assert previo del texto `<f>`, manteniendo el nombre `ASE n · VALIDACION_TOTAL · <celda>` y el fail-fast), (b) el writer **sella** los `<v>` computables por read-back de lo escrito (jamás `<f>`, jamás desde la intención del dominio), y (c) F-NEW-1 recibe su propio T0 con regla honesta por celda — reescritura Phase 0 por quincena (ampliando la allow-list de 16) o invariante declarado con motivo — quedando **prohibido el sellado cosmético** donde la divergencia sea real? | "Sí, proceder" = puerta a T0..T6. "No / con matices" = el Ingeniero lista qué sub-bloque o celda F-NEW-1 queda fuera y con qué justificación; eso entra como exclusión versionada con cita (no blanket) y el resto avanza. |

### 0.3 Mapeo al Rector (Sale / No sale)

**Sale (EXPLÍCITO):**
- Re-diagnosticar la cadena C15/F10 o reabrir "caché vs evaluación" (mandato 2026-10-14 + D-A lo cierran).
- Sobrescribir cualquier `<f>` en runtime fuera de Phase 0 (el sellado escribe SOLO `<v>`; el guard anti-fórmula de `EscribirValorNumerico` se ENDURECE, nunca se relaja); cualquier `InsertRow`/`DeleteRow` en procesamiento.
- Sellar TRUE (o cualquier `<v>`) donde la divergencia sea real — en particular F-NEW-1 en Q1 (mentir en caché es peor que el FALSE honesto).
- Computar el sellado desde la intención del dominio (el sellado es READ-BACK de lo escrito, nunca re-cálculo del dominio).
- Usar `<v>` como oráculo en cualquier gate nuevo o migrado; debilitar un test para que pase (si un test falla por premisa obsoleta, se retargetea con justificación citada).
- Literales/ramas por período o mes/año en el diff de producción (D-H); ramas por empresa fuera del catálogo citado; modificar el oficial in place (hash SHA256 antes/después); paquetes NuGet; commits/push.
- Renombrar rótulos del template en runtime; tocar mirror/hidden/labels fuera de contrato.

**No sale (entra, alcance congelado):**
| Frente | Superficie de cambio |
|---|---|
| T0 — evidencia | Congelamiento desde bytes de los 5 sub-bloques W-1 (texto-`<f>` + dependencias + caché oficial) + estado F10/C9/C14/C15 + diferencial oficial→salida de TODA la clase caché-viejo + allow-list Phase 0 actual + propuesta de ampliación F-NEW-1; cero código de producción |
| B1 — gate evaluador | `ValidacionOracleReader` + `ValidadorBasico` (W-1): asserts de texto-`<f>` + evaluación desde valores escritos; retirada la lectura de caché para sub-bloques; mismo nombre de error, mismo fail-fast |
| B2 — sellado | `OpenXmlPlantillaWriter`: sellado `<v>` por read-back con guard anti-fórmula intacto + allow-list del writer documentada |
| B3 — F-NEW-1 | T0 propio + regla honesta por celda (Phase 0 reescribe `<f>` por quincena con allow-list ampliada, o invariante declarado) |
| G — gates | Assert-`<f>` + evaluación ±0.5 + DetRetri-D vs R10 ±0.5 + shared/calcChain ciegos + healthy-open + prueba de FALSE-real-falla + greps de cierre |
| D — docs | `project-context.md` (doctrina "los cachés no son oráculo; el gate evalúa y el writer sella") + manual si aplica + lectura de cada FAIL |

**Fuera de scope explícito (OUT):** motor-espejo R4; cobertura `conditionalFormatting`/`dataValidations` (W-4); semántica DetRetri (redondeo/columna D intactos); F209/F219 (su propio T0); reconciliación de fixture-root (dependencia declarada; lo nuevo usa rutas explícitas); semántica de los auto-chequeos F-NEW-1 más allá de su regla honesta (no se rediseña el negocio de validación).

### 0.4 Decisiones ejecutivas (aprobación = aceptar estas)

| ID | Decisión |
|---|---|
| D-A | **Los cachés de fórmulas no son oráculo; el gate evalúa y el writer sella.** Ningún gate lee `<v>` de una celda de fórmula para decidir TRUE/FALSE o igualdad; evalúa desde valores escritos (texto-`<f>` + literales). Ningún `<v>` sellado se computa desde la intención del dominio. Es invariante de cierre (grep + tests). |
| D-B | **El archivo entregado era válido; el defecto era 100% del gate.** `fullCalcOnLoad="1"` + sin calcChain (Plan 28) hacen que Excel recalcule al abrir y C15 quede TRUE. El sellado no "arregla" el archivo: lo hace autoconsistente en caché para que ningún lector futuro herede una mentira. |
| D-C | **Falsa verificación → verificación real + artefacto autoconsistente.** W-1 pasaba por herencia (Resultado4) y fallaba por herencia (Resultado5): nunca verificó nada. Tras este plan, un TRUE significa "evaluado TRUE desde lo escrito" y un FALSE significa "divergencia real" — en ambos sentidos, con prueba. |
| D-D | **Assert de texto-`<f>` antes de evaluar.** Si la plantilla cambia una fórmula, el gate falla ruidosamente (`ERR-PLANTILLA` nombrando hoja+celda+texto esperado) en vez de evaluar la fórmula equivocada en silencio. El fixture T0 es la fuente de los textos esperados. |
| D-E | **Sellado por read-back, values-only, con allow-list documentada.** Solo `<v>`, solo celdas cuya cadena completa está escrita por la app, solo valores leídos del workbook de salida tras la escritura. El guard anti-fórmula (`EscribirValorNumerico` L1808-1815) cubre cada escritura del sellado. |
| D-F | **F-NEW-1 honesto o nada.** Si el T0 concluye que una celda debe variar por quincena, Phase 0 reescribe ese `<f>` (allow-list ampliada más allá de 16, invariante Plan 39 actualizado con cita) y el gate la evalúa por variante; si concluye invariante, veredicto motivado versionado. Prohibido sellar para callar un FALSE real. |
| D-G | **D-H + catálogo abierto irrenunciables.** Cero literales/ramas por período y cero ramas por empresa en el diff; el evaluador y el sellador reciben `Periodo` y mapas, nunca strings de período/hoja/empresa. |

---

## 1. PROPOSE

### 1.1 Intent

Que la corrida end-to-end desde la plantilla oficial termine en ÉXITO cuando el workbook es correcto — con el gate W-1 evaluando los 5 sub-bloques desde lo realmente escrito (assert de `<f>` + evaluación ±0.5), el archivo de salida autoconsistente en caché (sellado por read-back, Excel recalcula coincidiendo al abrir), un FALSE genuino haciendo fallar el gate con el mismo nombre accionable, y la familia F-NEW-1 con regla honesta por celda — sin tocar el oficial y sin mentir jamás en un `<v>`.

### 1.2 In Scope

- T0 de evidencia: 5 sub-bloques W-1 + F10/C9/C14/C15 + diferencial de clase caché-viejo oficial→salida + allow-list Phase 0 + propuesta F-NEW-1, todo citado celda-a-celda desde bytes, con Q1 respondida.
- Gate W-1 evaluador (T1): asserts de texto-`<f>` + evaluación Core; retirada la lectura de caché para sub-bloques; tests de FALSE-real-detectado.
- Sellado de cachés (T2): writer por read-back + guard anti-fórmula + allow-list del writer; test de coincidencia sellado-vs-evaluación.
- F-NEW-1 (T3): T0 propio + regla (reescritura Phase 0 por quincena o invariante motivado).
- Regresión Q1+Q2 end-to-end desde el OFICIAL con ÉXITO + prueba de falso-positivo-negativo (T4); barrido de clase con veredicto por celda (T5); docs + cierre (T6).

### 1.3 Non-Goals

Ver §0.3 OUT. En particular: NO se cambia ninguna semántica de negocio ni ningún valor escrito hoy (`F10 = 72497949967` y el resto del writer quedan intactos); NO se reescribe ninguna `<f>` fuera de lo que T3 demuestre para F-NEW-1 en Phase 0; NO se relaja ningún gate existente (O9/P9, DetRetri-D vs R10, shared, calcChain); NO se "arregla" el archivo de hoy (ya es válido).

### 1.4 Riesgos

| ID | Riesgo | Mitigación |
|---|---|---|
| R-MENTIRA | El sellado escribe un `<v>` que no coincide con lo que Excel recalcularía (caché mentiroso nuevo, peor que el viejo) | D-E: read-back de lo escrito + test de coincidencia sellado-vs-evaluación (T2) + Excel recalcula al abrir y debe coincidir (T4 healthy-open); ante duda la celda NO se sella (veredicto T5) |
| R-FORMULA-EQUIVOCADA | El evaluador verifica la fórmula vieja tras un cambio de plantilla (verde silencioso sobre la fórmula equivocada) | D-D: assert de texto-`<f>` previo a toda evaluación; tercera forma = `ERR-PLANTILLA`; drift gate del Plan 38 como segunda red |
| R-COSMETICO | F-NEW-1 se "cierra" sellando TRUE en Q1 donde la divergencia es real | D-F + Q1(c): prohibido por invariante; T3 exige T0 propio y el veredicto por celda; `code-reviewer` lo verifica como spec compliance |
| R-ALCANCE | El evaluador crece a un motor de fórmulas general (scope creep: parser completo de Excel) | Solo se evalúan las formas congeladas por T0 (igualdades y sumas de la cadena W-1 + F-NEW-1; tabla explícita, §2.2); cualquier forma fuera de tabla = FAIL, no "soportar más sintaxis" |
| R-DOBLE-F | Sellado y Phase 0 pisan la misma `<f>`/`<v>` (el re-encuadre del Plan 39 se vuelve excusa) | Frontera quirúrgica: Phase 0 toca `<f>` (allow-list T0+ampliación T3); el sellado toca SOLO `<v>` y jamás una celda con `<f>` distinta de la esperada (guard + assert previo); grep distingue fases por namespace/clase |
| R-INPLACE | Un path del sellado escribe sobre el oficial | `ValidarNoInPlace` intacto + el sellado corre sobre la copia de trabajo/salida, nunca sobre el oficial; hash SHA256 del oficial antes/después en T-gates |
| R-JULIO | Al tocar W-1/sellado se altera un período que ya abría sano | Puerta de identidad: Q1 y Q2 desde el OFICIAL reproducen las salidas de referencia (DetRetri-D vs R10 ±0.5 + `<f>`-idéntica); ante diff se revierte el pase, no el golden |
| R-FIXTURE-ROOT | Gates nuevos heredan rojos de resolución raíz | Rutas explícitas (oficial + fuentes por ASE por ruta explícita + temp FRESCA); prohibido `Insumos.Raiz()` y `Docs/Insumos` raíz en lo nuevo |
| R-ENTORNO | La suite infla rojos por bloqueo `FileShare.Read` con el oficial abierto en Excel | Puerta de entorno en T-gates: oficial cerrado antes de correr; un rojo `FileShare`/`IOException` por archivo abierto se triagea como entorno, nunca como defecto (pero nunca se usa como excusa para un rojo real) |

---

## 2. DESIGN

### 2.1 Enfoque: evidencia primero, evaluar en Core, sellar en Infrastructure, F-NEW-1 con T0 propio

**Fase 0 — evidencia (T0, cero código de producción).** Lectura BCL (`ZipArchive` + XML, sin Excel/COM) del oficial + la salida real de hoy: por cada uno de los 5 sub-bloques, tabla `(celda, texto-`<f>` exacto, dependencias transitivas con su texto, caché oficial, escrito-por-la-app SÍ/NO, caché en salida-hoy)`; estado F10/C9/C14/C15; diferencial oficial→salida de TODAS las celdas de fórmula cuyo `<v>` cambió-sin-`<f>` o quedó viejo (la clase); allow-list Phase 0 actual (16 celdas citadas) + propuesta de ampliación para F-NEW-1 (celdas + textos por quincena). Fixture que manda + revisión del Ingeniero (Q1).

**Fase 1 — gate evaluador (T1, Core).** `ValidacionOracleReader` deja de poblar `SubBloquesValidacionTotal` por `LeerValorBooleano` (caché) y entrega la evidencia para evaluar: textos-`<f>` + valores literales leídos del workbook de salida. `ValidadorBasico` (W-1) aserta el texto esperado (fixture T0) y evalúa el booleano desde los valores escritos (±0.5 donde haya aritmética). Mismo nombre de error (`ASE n · VALIDACION_TOTAL · <celda>`), mismo fail-fast, misma severidad. `LeerValorBooleano`/`LeerValorNumericoExigido` siguen vigentes para celdas que SÍ son oráculo-legítimo (O9/P9 y resto fuera de sub-bloques) — solo se retira su uso como veredicto de sub-bloque.

**Fase 2 — sellado (T2, Infrastructure).** Tras la escritura de negocio y antes del saneamiento calcChain, el writer recorre la allow-list del writer (T0: cadena C3:C9→C14→C15 + D25/O25/D34/F34 donde computable) y escribe cada `<v>` calculado por READ-BACK de las celdas efectivamente escritas en el workbook de salida. Cada escritura pasa por `EscribirValorNumerico`-o-equivalente-con-guard (celda con `<f>` inesperada → `ERR-PLANTILLA`, nunca sobrescritura; y aquí la celda destino del `<v>` ES la celda de fórmula — el sellado escribe `CellValue` preservando `CellFormula` intacta, lo que el guard actual prohibiría: T2 implementa el camino `SellarCacheFormula` separado, con assert previo de texto-`<f>` + preservación byte de `<f>` verificada por test). Si una dependencia no está escrita por la app (no hay read-back posible), la celda NO se sella.

**Fase 3 — F-NEW-1 (T3, su propio T0 + regla).** Mini-T0 BCL por celda (`DetRetri{per}!L32:L36`, `O32:O36`, `DetValiRetri{per}!I24:I29`, `INTERVENTORIA!S26:S31`): texto-`<f>` exacto, columna referenciada (M vs N), valor en Q1 vs Q2 tras R-INTERV-Q. Veredicto por celda: (a) debe variar por quincena → Phase 0 reescribe ese `<f>` (tabla `VarianteFNew1PorQuincena` indexada por `NumeroQuincena`, allow-list 16 → 16+N con citas, invariante Plan 39 actualizado) y el gate la evalúa por variante; o (b) invariante → motivo versionado y la celda queda fuera del sellado y de la evaluación quincenal. Ningún "probablemente igual".

**Fase 4 — regresión + clase + docs (T4..T6).**

### 2.2 Tabla de evaluación (formas que el gate sabe evaluar — cerrada por T0)

| Celda | Texto-`<f>` esperado (puntero E1/E4, el fixture manda) | Evaluación desde lo escrito |
|---|---|---|
| `VALIDACION_TOTAL!C15` | `C14=C9` (`t="b"`) | `Igualdad(C14_eval, C9_eval)` |
| `VALIDACION_TOTAL!C14` | `'Valida - Control Recaudo'!F10` | read-back del literal `F10` escrito por la app (Unidad P) |
| `VALIDACION_TOTAL!C9` | `SUM(C3:C8)` | `Σ read-back C3:C8` (±0.5) |
| `VALIDACION_TOTAL!C3:C8` | `Σ 'Recaudo *'!F21:F26` por fila | read-back de los literales `Recaudo *` escritos |
| `VALIDACION_TOTAL!D25` | `D24=C9` | `Igualdad(D24_eval, C9_eval)` |
| `VALIDACION_TOTAL!O25` | `O24=N9` | `Igualdad(O24_eval, N9_eval)` |
| `VALIDACION_TOTAL!D34` | `D33='CONSOLIDADO_TOTAL RECAUDO'!D109` | `Igualdad(D33_eval, read-back D109)` |
| `VALIDACION_TOTAL!F34` | `F33='CONSOLIDADO_TOTAL RECAUDO'!D109+'CONSOLIDADO_TOTAL RECAUDO'!N109` | `Igualdad(F33_eval, D109+N109)` (±0.5) |
| F-NEW-1 (4 familias) | Textos por T3 (M-fijo hoy) | Por variante Q1/Q2 según veredicto T3 |

Regla de cierre: cualquier `<f>` real cuyo texto no esté en esta tabla (tras normalizar espacios) = `ERR-PLANTILLA` (hoja+celda+texto esperado + texto observado). El evaluador NO parsea Excel general: compara el texto contra la tabla y aplica el operador tabulado (igualdad / suma, ±0.5).

### 2.3 Trade-offs (por qué NO las alternativas)

| Eje | Elegida | Descartada + por qué |
|---|---|---|
| Dónde evaluar | En Core, desde valores leídos del workbook de salida (texto-`<f>` + literales) | Seguir leyendo `<v>`: es la falsa verificación que causó el incidente (E2/E3). Recalcular con un motor Excel embebido: dependencia nueva + semántica duplicada (dos motores que pueden divergir); prohibido por Sale |
| Qué sellar | Solo `<v>` computable por read-back de lo escrito; lo no-escrito no se sella | Sellar todo lo stale (T5-clase completa): sellar sin read-back exige inventar el valor desde el dominio (viola D-E) o copiar el caché viejo (perpetúa la mentira). Lo no-sellable queda con veredicto explícito, no con caché heredado silencioso |
| Cómo escribir el `<v>` de una fórmula | Camino `SellarCacheFormula` separado: assert de texto + preserva `<f>` byte-idéntica + escribe `CellValue`/`DataType` | Reutilizar `EscribirValorNumerico` tal cual: su guard (L1808-1815) prohibe tocar celdas-fórmula por diseño — relajarlo rompería la red W-2.3 para todo el writer. El camino separado mantiene ambas garantías |
| F-NEW-1 | T0 propio + regla honesta (reescritura Phase 0 por quincena o invariante) | Sellar TRUE en Q1: MENTIRA (la divergencia tras R-INTERV-Q es real). Ignorarlo: deja 4 familias en FALSE-real en cada Q1 (el incidente se repite por otra puerta). Excluirlo en silencio: viola T5-cero-supuestos |
| Dónde enganchar el sellado | En el writer, tras escritura de negocio y antes del saneamiento calcChain | En el procesador: el writer es dueño de los 3 caminos de guardado + atomicidad + `ValidarNoInPlace`; un hook fuera dejaría `EscribirEspejoR1` sin sellado (misma razón que Plan 39 §2.1) |

### 2.4 Contratos afectados (resumen)

| Contrato | Cambio |
|---|---|
| Fixture T0 (nuevo, tests, cero producción) | 5 sub-bloques (texto + dependencias + cachés) + F10/C9/C14/C15 + diferencial de clase + allow-list 16 + propuesta F-NEW-1; hash SHA256 del oficial |
| `ValidacionOracleReader` (L130-138) | Sub-bloques: de `LeerValorBooleano` (caché) a entrega de evidencia evaluable (textos + literales); `ObtenerCeldaFormula` + asserts de presencia intactos; `LeerValorBooleano` sigue para O9/P9 y resto no-sub-bloque |
| `ValidadorBasico` (L410-419) | W-1: assert de texto-`<f>` (fixture) + evaluación desde valores (±0.5); mismo mensaje `ASE n · VALIDACION_TOTAL · <celda>`, mismo fail-fast, misma severidad; `SubBloquesValidacionTotal` del snapshot pasa a resultado evaluado (documentado en `ValidacionCruzadaSnapshot`) |
| `OpenXmlPlantillaWriter` | Nuevo paso de sellado (camino `SellarCacheFormula`, guard anti-`<f>` propio) + allow-list del writer documentada; `EscribirControlRecaudoTotal` y `EscribirValorNumerico` INTACTOS |
| `PreparadorBasePipeline` (solo si T3 lo exige) | Ampliación de allow-list 16 → 16+N (textos F-NEW-1 por quincena, tabla por `NumeroQuincena`); `ValidadorBasePeriodo` acepta la variante del run; invariante Plan 39 actualizado con cita |
| `WorkbookLeafCellMapValidaciones` | Solo lo que T0 demuestre stale (direcciones/fragmentos de la cadena W-1, cero lógica); la lista `SubBloquesValidacionTotal` conserva sus 5 miembros (ahora evaluados, no leídos) |
| `NombresHojaPeriodo` / `Periodo` / lectura / cálculo / sellos / calcChain / preflight / sinonimia | INTACTOS (el evaluador y el sellador los consumen, no los cambian) |
| `project-context.md` + manual | Doctrina "los cachés no son oráculo; el gate evalúa y el writer sella" + tabla de la clase + regla F-NEW-1 + lectura de cada FAIL |

### 2.5 Architecture Validation Certificate

| Principio | Veredicto |
|---|---|
| **SRP** | ✅ Una razón de cambio por costura: T0 declara textos+dependencias (datos), el evaluador verifica (Core puro), el sellador escribe cachés (Infrastructure), F-NEW-1 arbitra su regla (T3), cada gate contrasta una regla. Ninguna rama por período ni por empresa. |
| **OCP** | ✅ Un cambio futuro de plantilla se absorbe con un nuevo T0 (textos re-congelados; el evaluador no cambia salvo formas nuevas tabuladas); empresa no tabulada = identidad; quincena futura imposible por construcción (`Periodo.Parse` solo 1/2). |
| **DIP** | ✅ Evaluador puro (texto + decimales, sin OpenXML nuevo); sellador sobre el writer existente; tests con insumos reales + BCL ya presente; cero dependencias nuevas (OpenXML 3.5.1 ya referenciado). |
| **Best practices** | ✅ TDD con puertas (evidencia → evaluación → sellado → F-NEW-1 → regresión); fail-fast que nombra `ASE n · hoja · celda`; quincena por dominio; procesamiento values-only intacto + sellado con guard propio; tolerancia ±0.5; sin Excel/COM en pipeline ni tests; Visual Design Intent N/A (cero UI). |
| **Performance** | ✅ Evaluación O(celdas de la cadena) por ASE; sellado O(allow-list) una vez por guardado; cero costo en lectura/cálculo (el sellado sustituye la herencia del caché, no agrega pases de negocio). |

---

## 3. SPEC

### 3.1 Requirements

| ID | Requisito | Grupo | Criterio de aceptación |
|---|---|---|---|
| R-T-1 | T0 congela desde bytes los 5 sub-bloques (texto-`<f>` exacto + dependencias transitivas con texto + caché oficial + escrito-por-la-app SÍ/NO + caché en salida-hoy) + estado F10/C9/C14/C15 + diferencial oficial→salida de TODA la clase caché-viejo + allow-list Phase 0 (16 citadas) + propuesta de ampliación F-NEW-1 | T0 | Fixture que manda; Q1 respondida; revisión del Ingeniero; `git diff` de producción vacío para T0; SHA256 del oficial registrado antes/después idéntico |
| R-E-1 | W-1 evalúa (no lee caché): para cada sub-bloque, assert de texto-`<f>` + evaluación del booleano desde valores escritos; mensaje y fail-fast intactos (`ASE n · VALIDACION_TOTAL · <celda>`, W-1) | B | 5/5 ASE 2026072 en verde con valores reales; `LeerValorBooleano` ya no decide ningún sub-bloque (grep) |
| R-E-2 | Tercera forma = `ERR-PLANTILLA` (hoja+celda+texto esperado+texto observado); ninguna forma fuera de la tabla §2.2 se evalúa en silencio | B | Test con `<f>` mutada → FAIL con los 4 campos; cero escritura parcial |
| R-E-3 | Tolerancia ±0.5 en toda comparación aritmética del evaluador; booleanos puros por igualdad exacta post-tolerancia | B | Tests en el borde (±0.4 PASS, ±0.6 FAIL donde aplique) |
| R-E-4 | Un FALSE REAL se detecta: divergencia genuina inyectada (valor escrito que rompe la igualdad) → gate FAIL nombrando la celda; y un TRUE evaluado (no heredado) → PASS aunque el caché heredado diga FALSE | B | Par de tests rojo→verde que habrían fallado con el lector viejo en AMBAS direcciones (falso-verde Resultado4 y falso-rojo Resultado5) |
| R-S-1 | Sellado por read-back: cada `<v>` = función de los valores efectivamente escritos en la salida (nunca del dominio); lo no-escrito no se sella | B | Test: sellado == evaluación independiente; test: dependencia no-escrita → celda intacta (caché heredado + veredicto T5, nunca valor inventado) |
| R-S-2 | Camino `SellarCacheFormula` separado con assert de texto-`<f>` + preservación byte de `<f>`; el guard de `EscribirValorNumerico` (L1808-1815) intacto para todo lo demás | B | Test: `<f>` byte-idéntica antes/después del sellado workbook-wide en las celdas selladas; grep: 0 escrituras `<f>` fuera de Phase 0/ampliación T3 |
| R-S-3 | Allow-list del writer documentada (qué celdas se sellan y por qué cada una es computable); fuera de allow-list nada se sella | B | Lista citada en código + fixture; test que falla si el sellado toca una celda fuera de lista |
| R-F-1 | F-NEW-1: mini-T0 por celda (texto + columna M/N + valor Q1 vs Q2) + veredicto honesto: reescritura Phase 0 por quincena (allow-list ampliada, invariante Plan 39 actualizado) o invariante motivado versionado | B | Tabla celda → veredicto con citas; prohibido sellado cosmético (test: Q1 con divergencia real sigue FALSE hasta que la regla la resuelva) |
| R-G-1 | `<f>` byte-idénticas al template workbook-wide tras sellado (texto + `t`/`ref`/`si`) salvo ampliación T3 aprobada; `<v>` solo en destinos del contrato + celdas selladas R-S-3 | G | FAIL ante cualquier `<f>` divergente o `<v>` fuera de contrato+allow-list; prohibido `<v>` como oráculo |
| R-G-2 | DetRetri-D 5/5 vs R10 ±0.5 en Q1 y Q2 + shared/calcChain ciegos + healthy-open (sin diálogo de reparación) + app en ÉXITO (exit 0 / sin ERR-VALIDACION) con `C15` y los 4 sub-bloques consistentes | G | Regens FRESCAS a temp DESDE EL OFICIAL (Q1+Q2 × 5 ASE); Excel abre sano y recalcula coincidiendo con lo sellado (paso manual del Ingeniero como aceptación final) |
| R-R-1 | Suite con set-diff vs línea base 36/782/818 por tarea + build 0 warnings; oficial intacto (SHA256); greps de cierre (D-H, catálogo, sin-`<f>`-fuera-de-fase, sin-insert/delete, sin-`<v>`-como-oráculo) | R | `dotnet build …slnx` 0 warnings + `dotnet test` por tarea; ningún rojo nuevo sin dependencia declarada; oficial hash idéntico |
| R-R-2 | Docs: doctrina evaluar+sellar + tabla de la clase + regla F-NEW-1 + lectura de cada FAIL | R | Diff acotado; sin código en el mismo commit |

### 3.2 Scenarios (Given/When/Then)

- **S1 (falso-rojo de hoy, Q2):** Given el oficial + 2026072, When run end-to-end, Then ÉXITO (sin ERR-VALIDACION); `C15` y los 4 sub-bloques evaluados TRUE desde lo escrito; el sellado deja `C14/F10/C9/C15` autoconsistentes; Excel abre sano y recalcula coincidiendo.
- **S2 (falso-verde legacy):** Given una salida con caché `C15=1` heredado pero divergencia real inyectada (p. ej. `F10` que rompe `C14=C9`), When gate W-1, Then FAIL `ASE n · VALIDACION_TOTAL · C15` (el verde heredado ya no existe).
- **S3 (tercera forma):** Given una `<f>` de la cadena mutada (p. ej. `C15` con otro texto), When gate o sellado, Then `ERR-PLANTILLA` con hoja+celda+esperado+observado; cero escritura parcial.
- **S4 (in-place denegado):** Given salida == ruta del oficial, When cualquier camino (incluido el sellado), Then denegación antes de tocar un byte; SHA256 del oficial idéntico.
- **S5 (Q1 honesto):** Given el oficial + 2026071, When run, Then ÉXITO con la regla F-NEW-1 que T3 fije; ninguna celda F-NEW-1 con divergencia real aparece sellada en TRUE.
- **S6 (sellado coincide):** Given salida post-escritura, When sellado + evaluación independiente, Then cada `<v>` sellado == valor evaluado; cada `<f>` sellada byte-idéntica a antes del sellado.
- **S7 (clase cerrada):** Given el diferencial T0, When T5, Then cada celda de fórmula con caché viejo tiene veredicto (se evalúa / se sella / invariante / requiere decisión); ningún "probablemente igual".
- **S8 (regresión):** Given Q1+Q2 × 5 ASE desde el OFICIAL a temp FRESCA, When regens, Then ÉXITO + `<f>`-idéntica + DetRetri-D 5/5 vs R10 ±0.5 + shared/calcChain verdes + set-diff de suite sin rojos nuevos.

---

## 4. TASKS

| ID | Tarea | Grupo | Depende de | Verificación |
|---|---|---|---|---|
| T0 | Evidencia congelada (cero código). BCL ZIP+XML sobre oficial + salida-hoy (Resultado5) + legacy (Resultado4): 5 sub-bloques (texto-`<f>` + dependencias + cachés) + F10/C9/C14/C15 + diferencial de clase oficial→salida + allow-list 16 + propuesta F-NEW-1 + SHA256 del oficial. Con Q1 respondida y revisión del Ingeniero. | T0 | Q1 | R-T-1 (S7 parcial-fixture); `git diff --stat` vacío en producción; el fixture manda |
| T1 | Gate W-1 evaluador (TDD, Core). `ValidacionOracleReader`: sub-bloques a evidencia evaluable (textos + literales; retirar `LeerValorBooleano` para sub-bloques); `ValidadorBasico`: assert de texto + evaluación ±0.5, mismo mensaje/fail-fast. Tests: S1-parcial (TRUE evaluado con caché heredado FALSE), S2 (FALSE real → FAIL), S3 (tercera forma → `ERR-PLANTILLA`), S-borde (R-E-3). | B+G | T0 | R-E-1/R-E-2/R-E-3/R-E-4 (S1/S2/S3 parcial); build 0 warnings; grep: 0 `LeerValorBooleano` en camino de sub-bloque |
| T2 | Sellado de cachés (TDD, Infrastructure). Camino `SellarCacheFormula` (assert texto + preserva `<f>` + escribe `CellValue`) + paso post-escritura/pre-calcChain en las 3 rutas + allow-list del writer documentada. Tests: S6 (sellado==evaluación, `<f>` intacta), S4 (in-place denegado), no-escrito-no-sellado (R-S-1). | B+G | T1 | R-S-1/R-S-2/R-S-3 (S4/S6 parcial); build 0 warnings; grep: 0 escrituras `<f>` fuera de Phase 0/ampliación T3 |
| T3 | F-NEW-1 con su T0 y su regla. Mini-T0 BCL por celda (texto + M/N + Q1-vs-Q2) → veredicto por celda: (a) reescritura Phase 0 por quincena (tabla por `NumeroQuincena`, allow-list 16→16+N, `ValidadorBasePeriodo` por variante, invariante Plan 39 actualizado) con `<f>`-idéntica-a-variante, o (b) invariante motivado versionado. Tests: S5 (Q1 honesto; prohibido TRUE cosmético). | B | T0 (puede correr en paralelo a T1/T2; su wiring entra tras su veredicto) | R-F-1 (S5); build 0 warnings; 0 literales de período en el diff |
| T4 | Regresión end-to-end desde el OFICIAL (Q1+Q2 × 5 ASE a temp FRESCA): ÉXITO (exit 0 / sin ERR-VALIDACION), `C15` + 4 sub-bloques consistentes, DetRetri-D vs R10 ±0.5, shared/calcChain ciegos, healthy-open con recálculo coincidente, S8. Solo rutas explícitas; oficial cerrado en Excel (R-ENTORNO). | R | T1, T2, T3 | R-G-2/R-R-1 (S1/S5/S8); suite con set-diff vs 36/782/818 |
| T5 | Barrido de la CLASE completa. Cada celda del diferencial T0 con veredicto (se evalúa / se sella / invariante / requiere decisión con su propio T0). Si un veredicto exige diseño propio → follow-up con T0 propio (no bloquea el cierre salvo camino W-1/F-NEW-1). | R | T4 | R-G-1 (S7 completo); tabla sin supuestos silenciosos |
| T6 | Docs + cierre. `project-context.md` (doctrina evaluar+sellar + tabla de la clase + regla F-NEW-1 + lectura de cada FAIL S2/S3/S4) + manual si aplica + build 0 warnings + greps globales (D-H, catálogo-abierto, sin-`<f>`-fuera-de-fase, sin-insert/delete, sin-`<v>`-como-oráculo, 0 refs nuevas indebidas) + SHA256 oficial + CRLF + sin commits. | D+Transversal | T0..T5 | R-R-2; DoD §5 |

**Orden sugerido (puertas duras):** Q1 → T0 (evidencia, sin código) → T1 (evaluador; prohibido sin T0 verde) → T2 (sellado; prohibido sin T1 verde) → T3 (F-NEW-1, en paralelo a T1/T2 desde T0; su wiring Phase 0 entra tras su veredicto) → T4 (regresión Q1+Q2 desde el oficial) → T5 (clase con veredicto por celda) → T6. Build 0 warnings antes de marcar cada tarea completa. Sin commits (los hace el Ingeniero con `#commit`/`#push`).

---

## 5. Resumen ejecutivo, riesgos y decisiones

- **Resumen:** el incidente 2026-10-14 demostró que W-1 nunca verificó nada: leía cachés `<v>` heredados de la plantilla — falló en falso sobre un archivo válido (Resultado5: `F10 = 72497949967` correcto, cachés viejos `C14 = 73981716599`/`C9 = 0`/`C15 = 0`) y pasaba en falso sobre herencia (Resultado4: `C15 = 1` heredado). Este plan lo convierte en verificación real (evaluar desde lo escrito con assert de texto-`<f>`) + artefacto autoconsistente (sellar `<v>` por read-back), con F-NEW-1 arbitrado honestamente por su propio T0. Cambio acotado: evaluador puro en Core + sellado values-only en Infrastructure + asserts; cero semántica de negocio tocada.
- **Riesgo principal:** R-MENTIRA + R-COSMETICO (sellar un valor que Excel no reproduciría; callar un FALSE real de F-NEW-1). Contenido por triple red: read-back obligatorio + test sellado==evaluación + recálculo coincidente al abrir (T2/T4); prohibición de sellado cosmético con verificación de `code-reviewer` (T3/D-F/Q1c).
- **Por qué NO seguir leyendo caché ni "tolerar" el ERR:** el caché proviene de la plantilla de entrada, no de la salida — cualquier umbral sobre él seguiría siendo falsa verificación en ambas direcciones (E2/E3). Y el archivo no necesita arreglo (Excel ya recalcula TRUE): necesita un gate que verifique y un caché que no mienta al próximo lector.
- **Por qué NO un motor de fórmulas general:** la tabla §2.2 (igualdades + sumas de la cadena W-1) cubre la clase medida; un parser general es un motor nuevo con sus propios falsos, y su divergencia con Excel sería el próximo incidente. Forma fuera de tabla = FAIL, no "más sintaxis".
- **Decisión para el Ingeniero:** ratificar D-A..D-G (§0.4) + responder Q1 (§0.2, con el alcance de F-NEW-1). **Forks con recomendación:** sellar solo-lo-computable vs sellar-toda-la-clase (recomendado solo-lo-computable: D-E/R-S-1); F-NEW-1 reescritura-Phase-0 vs invariante (recomendado: lo que el mini-T0 demuestre, sin a priori — T3); alcance OUT (§0.3, recomendado: diff mínimo). **Sin preguntas técnicas bloqueantes pendientes; gate de entrada = Q1.**

### DoD (Definition of Done)

1. `dotnet build Remuneracion.WinForms/Remuneracion.WinForms.slnx` con 0 warnings.
2. Puertas en orden: T0 congelado del disco (Q1 + revisión + SHA256 oficial) → T1 evaluador (S2/S3: FALSE-real-falla, tercera-forma-FAIL) → T2 sellado (S6: sellado==evaluación, `<f>` intacta) → T3 F-NEW-1 con veredicto por celda → T4 regresión Q1+Q2 desde el OFICIAL en ÉXITO (S1/S5/S8) → T5 clase con veredicto por celda.
3. `<f>`-idéntica workbook-wide (salvo ampliación T3 aprobada) + DetRetri-D 5/5 vs R10 ±0.5 ambos períodos + shared/calcChain verdes + healthy-open con recálculo coincidente.
4. Greps de cierre: 0 literales/ramas por período (D-H); 0 ramas por empresa fuera de filas citadas; 0 escrituras `<f>` fuera de Phase 0/ampliación T3; 0 insert/delete en procesamiento; ningún `<v>` como oráculo (0 `LeerValorBooleano` en camino de sub-bloque); oficial intacto (hash); ningún test debilitado sin justificación citada.
5. `project-context.md` + manual actualizados (doctrina evaluar+sellar + tabla de la clase + F-NEW-1); regens siempre a temp FRESCA con rutas explícitas y el OFICIAL como única entrada; sin commits del agente; CRLF; sin emojis.

### Follow-ups explícitos (fuera de este plan)

- Cambio futuro del oficial por el negocio: su propio T0 (textos re-congelados; drift gate del Plan 38 como detector).
- Veredicto T5 "requiere decisión": cada uno con su propio T0 + plan hijo (no se sellan por defecto).
- Motor-espejo R4; CF/DV (W-4); DetRetri-semántica; F209/F219 (su propio T0); reconciliación de fixture-root (dependencia declarada).
- Puerta de entorno documentada: oficial cerrado en Excel antes de suite/regens (rojos `FileShare.Read` = entorno, nunca defecto ni excusa).

---

## §0.2 Pregunta única de confirmación (cierre — BLOQUEANTE para T1..Tn)

| ID | Pregunta única | Opciones |
|---|---|---|
| Q1 | ¿Ratifica (a) W-1 evaluador — adiós lectura de caché, assert de texto-`<f>` + evaluación desde lo escrito, mismo nombre `ASE n · VALIDACION_TOTAL · <celda>` y fail-fast —, (b) sellado por read-back — solo `<v>`, solo computable, jamás `<f>`, jamás desde el dominio —, y (c) F-NEW-1 con T0 propio y regla honesta — reescritura Phase 0 por quincena (allow-list 16→16+N) o invariante motivado — con el sellado cosmético prohibido donde la divergencia sea real (D-A..D-G)? | **A. Sí, proceder** (puerta a T0..T6) · **B. Sí, con exclusiones** (liste sub-bloque/celda F-NEW-1 excluido + justificación; entra como exclusión versionada con cita) · **C. No / con matices** (indique qué decisión D-A..D-G no ratifica y el plan se re-encuadra) |

**Respuesta del Ingeniero (2026-10-14): A — "Sí, proceder".**

---

## Enmienda E-1 (2026-10-14) — Criterio de éxito estricto: probar el cierre de los 5

**Origen:** al cerrar T0/T1/T2 el implementador marcó `O25`/`D34`/`F34` con `SeGatea=false` alegando que su cadena pasa por el espejo de remuneración (`VALIDACION_*!N9 → REMUNERACION_*!D109 → Reporte Componentes R1`) y que "la app no lo puebla", con una divergencia reportada de 61,1B vs 73,1B. El orquestador **no aceptó esa desviación como final** (reducción de alcance de un invariante ratificado) y elevó el fork al Ingeniero.

**Decisión ratificada por el Ingeniero (2026-10-14): criterio de éxito ESTRICTO — la app debe PROBAR el cierre de los 5 sub-bloques.**

| ID | Decisión enmendada |
|---|---|
| E1-1 | **Prohibido `SeGatea=false` como forma de pasar.** El evaluador debe resolver la cadena COMPLETA de los 5 sub-bloques desde los valores que la app escribió (incluido el espejo de remuneración / `Reporte Componentes R1`), y emitir veredicto TRUE/FALSE por celda. "No evaluable" solo es admisible si se DEMUESTRA, con números y la cadena citada, que la resolución no es posible desde lo escrito — y aun así el veredicto queda como hallazgo, no como verde. |
| E1-2 | **Un sub-bloque que no cierre es un DEFECTO REAL, no un gate a apagar.** Si la evaluación honesta da FALSE, el archivo entregado mostraría validaciones en FALSE tras el recálculo de Excel: se reporta con evidencia celda-a-celda y se abre su **frente propio** (familia espejo R1 / remuneración). Prohibido sellar, apagar o re-etiquetar para que la suite pase. |
| E1-3 | **T4 (regresión) NO puede declarar ÉXITO si algún sub-bloque quedó sin evaluar.** El ÉXITO exige los 5 evaluados TRUE (o el defecto E1-2 declarado y escalado, con el plan detenido en ese punto). |
| E1-4 | **Bloqueante Q1 incorporado a la familia T3:** el oficial referencia `VALIDACION_TOTAL!C3:C8 → 'Recaudo *'!F21:F26` FIJO, pero en Q1 el recaudo se escribe en la columna **D** ⇒ `C9 = 0 ≠ F10` ⇒ `C15` FALSE en Q1. Es la misma familia que F-NEW-1 (fórmula del template que debe variar por quincena y no varía): T3 lo incluye en su T0 y su regla (Phase 0 reescribe el `<f>` por `NumeroQuincena` — dominio, nunca contenido), ampliando la allow-list más allá de las 16 celdas. |
| E1-5 | La tabla de evaluación §2.2 y el veredicto por celda de T5 quedan **vinculados a E1-1/E1-2**: el veredicto "requiere decisión" deja de ser una salida válida para los 5 sub-bloques de W-1 (sí lo sigue siendo para el resto de la clase que T5 barra). |

**Estado tras la enmienda:** T0/T1/T2 ejecutados (suite 32/802/834; 4 rojos cerrados; 0 nuevos) con la desviación `SeGatea=false` **revertida por E1-1**. Siguiente tramo: evaluación estricta de los 5 (incluida la cadena del espejo) + T3 (F-NEW-1 + bloqueante Q1) → T4/T5/T6.
