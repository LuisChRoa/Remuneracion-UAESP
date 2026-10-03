# Plan 21 — Espejo estructural Reporte Componentes R1 (fin de mapas absolutos)

> **Alcance:** reemplazar el enfoque "mapa absoluto por (`Ase.Id`, quincena) + ocurrencias congeladas" por un modelo de **espejo estructural fila-a-fila 1:1** fuente→destino para la hoja `Reporte Componentes R1` (y `Reversion Pagos R4` donde T0 lo confirme), dimensionando cada bloque ASE a la forma de la **fuente del período actual** (insertar/borrar filas + reanclar fórmulas), con T0 bloqueante que re-deriva y congela la tabla espejo por ASE.
> **Rector (IRRENUNCIABLE):** `README.md` + `Docs/Propuesta_Proyecto_Automatizacion_Remuneracion_UAESP.md` + `Docs/Detalle de plantilla.docx` ("se deben adicionar filas o suprimir según el número de filas que traiga el reporte fuente"; "cuando el número de filas de un reporte fuente cambie, inserta o elimina filas preservando las fórmulas") + `Docs/Prompt_Maestro_Proyecto_Remuneracion_UAESP _ Vo .docx` ("las fórmulas pueden desajustarse si el número de filas cambia entre períodos; lee siempre de la fuente"). **Invariantes del proyecto:** NUNCA sobrescribir fórmulas (solo valores); tolerancia ±0.5; ASE4 sin columna "Especiales" (0 si ausente); Q2 aplica SALDOS POR NOTA/RETRIBUCIÓN NEGATIVA; redondear a entero antes de DetRetri/columna D; build 0 warnings; finales de línea `\r\n`; fail-fast nombrando ASE+reporte+celda.
> **Continuidad:** HU-01..HU-20 cerradas. Este plan construye encima de: `ExcelDataReaderWorkbookLeafInputReader.LeerLEspecialesMenores` (líneas ~505-617, rol/ocurrencia sobre zona Componente), `WorkbookLeafCellMapInterventoria.LMenoresPorAseQ2` (L20 = `TotalD_E` ocurrencia 2), `OpenXmlPlantillaWriter.ObtenerOCrearCelda` + `EscribirCeldasLEspecialesMenores` (escritura a celdas fijas, sin mover filas, con guard anti-fórmula), y los 5 mapas absolutos (`PorAse`, `Q2`, `PorEmpresa`, `Interventoria`, `Validaciones`). Este plan NO reabre su semántica de cálculo: **cero cambios de fórmulas de negocio; Q1/Q2-julio intactos por construcción.**
> **Insumos canónicos (no se re-fijan, no se inventan datos):** `Docs/Insumos/REMUNERACION 2026071`, `Docs/Insumos/REMUNERACION 2026072`, `Docs/Insumos/Remuneracion 202607-1 Total.xlsx`, `Docs/Insumos/Remuneracion 202607-2 Total.xlsx`. **Caso de regresión:** período 2026082 Q2, insumos en `Docs/Prueba2/Insumos` (5 carpetas ASE + `R10_Remuneracion_2026082.xlsx`) — el fallo que motiva esta HU.
> **Estado:** PENDIENTE DE APROBACIÓN — no implementar hasta OK del Ingeniero.
> **Fecha:** 2026-10-02

---

## 0. Clarification Gate

**Una sola pregunta con fork real para el Ingeniero (decisión D-A en §0.3):** tras el veredicto T0, si ambas alternativas resultan viables, ¿**A (mutar filas + reanclar fórmulas, recomendado)** o **B (regenerar la hoja desde fuente + plantilla de fórmulas por período)**? El plan viene redactado con A como ruta principal y B como fallback acotado (tarea T-alt); si el Ingeniero elige B de entrada, solo cambia la fase de escritura (§2.2, tareas T4/T-alt). Todo lo demás —T0, modelo de filas tipadas, tabla espejo congelada, gates de evidencia— es idéntico en ambas rutas. Punto de bloqueo respetado: NO se implementa nada en este documento.

### 0.1 Qué está verificado y qué NO (honestidad técnica)

**Verificado en esta planificación** (lectura directa de código/disco):

| # | Hecho | Método | Consecuencia |
|---|---|---|---|
| V1 | `LeerLEspecialesMenores` resuelve cada celda del mapa por rol+ocurrencia 0-based dentro de la zona Componente (filas anteriores a A=`Componente` B=`Total`) y lanza `ERR-VALIDACION` (`CalculoInvalidoException`) si `candidatos.ElementAtOrDefault(ocurrencia)` es null — mensaje literal `la fuente R1 no trae la fila del rol {rol} (ocurrencia {ocurrencia}) requerida por el mapa T0-0.5 para la celda {HojaR1}!{celda}` | Lectura `ExcelDataReaderWorkbookLeafInputReader.cs:505-617` | El fallo 2026082 (`TotalD_E` ocurrencia 2 para L20, ASE 1) es estructural: la fuente de agosto trae menos bloques `TotalD_E` que el golden julio-Q2 del que se congeló la ocurrencia |
| V2 | `LMenoresPorAseQ2[1]` contiene `("L20", TotalD_E, 2)`; `ObtenerLMenores(aseId, quincena)` despacha Q1/Q2 por `NumeroQuincena` y lanza si el ASE no está soportado | Lectura `WorkbookLeafCellMapInterventoria.cs:164-226` | El mapa confirma direcciones absolutas + ocurrencias congeladas derivadas de julio; cualquier período con distinta cardinalidad de bloques rompe |
| V3 | `EscribirCeldasLEspecialesMenores` escribe a celdas fijas del mapa vía `EscribirValorNumerico`; `ObtenerOCrearCelda` solo crea la celda/fila adresada — **no inserta ni borra filas, no desplaza rangos, no reancla fórmulas**; el guard anti-fórmula lanza `ERR-PLANTILLA` si el destino es fórmula | Lectura `OpenXmlPlantillaWriter.cs:1158-1229` | El escritor hoy es incapaz de absorber un bloque de distinto tamaño: la otra mitad de la fragilidad (lectura §V1 + escritura §V3) |
| V4 | Existen 9 mapas en `Remuneracion.Infrastructure/Excel/`: `WorkbookLeafCellMap`, `...AjustesSfT`, `...BalanceSc`, `...Interventoria`, `...PorAse`, `...PorEmpresa`, `...Q2`, `...ReporteBanco`, `...Validaciones` — todos con direcciones absolutas | Listado de directorio | El riesgo de desplazamiento apilado (§2.3) afecta a los 5 mapas hoja-R1 aguas abajo, no solo al de Interventoria |
| V5 | `Docs/Prueba2/Insumos/` contiene 5 carpetas ASE (`1-Promoambiental`, `2-Lime`, `3-Ciudad Limpia`, `4-Bogota Limpia`, `5-Area Limpia`) + `R10_Remuneracion_2026082.xlsx`; `Docs/Prueba2/Resultado/` existe (vacía); `Docs/Insumos/` contiene `REMUNERACION 2026071`, `REMUNERACION 2026072` y ambos `Remuneracion 202607-{1,2} Total.xlsx` | Listado de directorio | Insumos del caso de regresión y de los goldens disponibles en disco para T0 |
| V6 | `plans/` 01..20 ocupados (20 = conciliaciones RESUMEN MES Q2 + R10); este plan es el **21**, primer número libre | Listado `plans/` | Numeración del documento |
| V7 | La doctrina vigente prohíbe offsets e insert/delete de filas ("si la plantilla no alcanza: fail-fast honesto", planes 07/16) | Planes 07/16 | Esta HU **deroga puntualmente** esa prohibición SOLO para los bloques espejo R1/R4 (§0.3 D-C): es el cambio de doctrina central y por eso exige T0 + reanclaje + gates |

**Aportado por el Ingeniero como evidencia medida (NO re-verificado en disco en esta planificación; se re-verifica en T0 antes de implementar):** la hoja `Reporte Componentes R1` de salida es espejo 1:1 fila-a-fila de `Recaudoporcomponente` (mismo orden de firmas A–E, desplazamiento +1 por encabezado de bloque), validado 10/10 (5 ASE × Q1 `REMUNERACION 2026071` + Q2 `REMUNERACION 2026072`); conteos por ASE Q1 = 38/52/46/69/44, Q2 = 45/75/43/73/52; `Reversion Pagos R4` es espejo 1:1 de `ReversiónPorComponente`; la variación entre períodos viene de la cardinalidad de bloques (agosto pierde el bloque ENEL negativo; otros ASE traen `ENERBIT`, `NUEVO ESQUEMA`, códigos 1/2/5).

### 0.2 Mapeo al Rector (in vs out)

**Entra:**

| Requisito | Superficie de cambio |
|---|---|
| Espejo R1: leer fuente como secuencia de filas tipadas por firma (A/B/C/D/E) + valores por columna | Nuevo lector espejo en `ExcelDataReaderWorkbookLeafInputReader` (o clase adyacente; T0 decide si sustituye o convive con `LeerLEspecialesMenores`) + modelo en Core (`FilaEspejoR1`: firma + valores por columna) |
| Espejo R1/R4: dimensionar el bloque destino a la fuente del período actual (insert/borrar filas) | `OpenXmlPlantillaWriter`: nueva capacidad de mutación de filas + reanclaje de fórmulas que referencian el bloque (hoy inexistente, §V3) |
| Reanclaje: fórmulas que referencian rangos del bloque + filas de resumen/validación relativas al bloque | Mapa de fórmulas por bloque + ajuste de referencias tras mutación; filas TOTAL/OPORTUNO/EXTEMPORÁNEO reposicionadas relativas |
| Regla de ausencia: la fuente define la forma (fila ausente = fila suprimida, nunca 0 simulado) + invariantes de cierre (`Componente/Total`, `Mes`, `Subs/Cont`, `Total`) | Lector espejo (sin fail-fast por cardinalidad) + asserts de cierre en validador |
| T0 bloqueante: re-derivar y congelar la tabla espejo por ASE con evidencia Q1+Q2 (+ agosto como regresión) | Fase 0 (§4, PR 1): dumps fuente-vs-destino por ASE, tabla de firmas, veredicto A/B, inventario de fórmulas afectadas |
| Goldens Q1+Q2 ±0.5 + caso Prueba2/agosto que hoy falla | `Remuneracion.IntegrationTests` (xUnit + FluentAssertions): goldens existentes verdes + nuevo test de regresión 2026082 modo 5 ASE |

**Sale (EXPLÍCITO):** cambios a la aritmética de negocio (CONSOLIDADO, DetRetri, AJUSTES-SF-T, banco, BCE, conciliaciones, INTERVENTORIA-valores); reinterpretar valores de agosto (dato, no cálculo); UI WinForms y CLI salvo el resumen log mínimo si T0 lo exige; paquetes NuGet; commits (los hace el Ingeniero con `#commit`).

### 0.3 Decisiones ejecutivas (aprobación = aceptar estas)

| ID | Decisión |
|---|---|
| D-A | **Ruta A (mutar + reanclar) como principal, B como fallback.** Justificación: preserva la plantilla y todas sus fórmulas validadas (menor blast radius sobre lo certificado HU-07..HU-20); B se activa solo si T0 demuestra que el reanclaje es intratable (ver §2.2). |
| D-B | **La fuente del período actual define la forma; prohibido simular filas ausentes con 0.** El fail-fast por cardinalidad (`slot ausente`, doctrina HU-12/HU-16) se retira SOLO dentro de los bloques espejo; fuera de ellos sigue vigente. |
| D-C | **Derogación puntual de la prohibición insert/delete:** autorizada exclusivamente para los bloques espejo R1 (y R4 si T0 lo confirma), con reanclaje obligatorio y gates de evidencia. Fuera de esos bloques, la prohibición sigue intacta. |
| D-D | **T0 bloquea todo:** ninguna firma, ningún rango de bloque y ninguna fórmula entra al código sin evidencia Q1+Q2 (+ agosto). Si T0 refuta el espejo 1:1 en algún ASE/hoja → **NEEDS_CONTEXT con recorte, nunca invención**. |
| D-E | **Desplazamiento apilado tratado explícitamente:** insertar/borrar en el bloque de un ASE desplaza los bloques inferiores; el plan exige que T0 inventaríe todas las celdas absolutas aguas abajo y que la implementación las reubique por bloque, no por dirección (§2.3). |
| D-F | **Una sola llamada de escritura** por proceso (mismo overload lista HU-07..HU-12, patrón G5 de HU-16). Sin segundo pase. |

---

## 1. PROPOSE

### 1.1 Intent

Que el motor deje de romperse cada vez que la fuente R1 trae un número distinto de bloques: leer `Recaudoporcomponente` (y `ReversiónPorComponente` donde aplique) como lo que realmente son —secuencias de filas tipadas—, reproducir esa secuencia exacta en el bloque del ASE dimensionando a la fuente del período en curso, reanclar las fórmulas afectadas sin tocar su semántica, y demostrarlo con los goldens de julio (Q1+Q2, dif ±0.5) más el período 2026082 de agosto que hoy falla con `ERR-VALIDACION`.

### 1.2 In Scope

- T0 discovery bloqueante (§4 Fase 0) + tabla espejo congelada por ASE con evidencia Q1+Q2+agosto.
- Modelo de filas tipadas en Core + lector espejo en Infrastructure (R1; R4 supeditado a veredicto T0).
- Capacidad de mutación de filas + reanclaje de fórmulas en `OpenXmlPlantillaWriter`, acotada a los bloques espejo.
- Reposicionamiento relativo de filas de resumen/validación del bloque; invariantes de cierre.
- Goldens Capa A Q1+Q2 (regresión) + test de regresión 2026082 modo 5 ASE (hoy rojo → debe quedar verde).
- Actualización de `.opencode/project-context.md` (doctrina espejo + derogación puntual D-C).

### 1.3 Non-Goals

Ver §0.2 "Sale". En particular: NO se recalcula ningún total de negocio; NO se tocan los mapas fuera de los bloques espejo salvo su reubicación mecánica por desplazamiento (§2.3); NO se agregan firmas/roles nuevos sin evidencia T0.

### 1.4 Riesgos

| ID | Riesgo | Mitigación |
|---|---|---|
| R1 | El espejo 1:1 no se sostiene en algún ASE/hoja (la evidencia aportada es 10/10 pero no re-verificada) | T0 lo re-deriva desde disco; si refuta → NEEDS_CONTEXT con recorte (D-D), sin invención |
| R2 | Desplazamiento apilado: mutar el bloque ASE-N rompe todas las celdas absolutas de los bloques N+1..5 y de los mapas aguas abajo | §2.3: inventario T0 de celdas afectadas + reubicación por bloque; goldens 5-ASE lo cubren |
| R3 | Reanclaje de fórmulas OpenXML intratable (referencias cruzadas, rangos con nombre, fórmulas fuera del bloque que apuntan dentro) | T0 inventaría fórmulas afectadas; si intratable → fallback B (D-A); guard anti-fórmula existente sigue protegiendo |
| R4 | Regresión julio: el dimensionado dinámico altera salidas que hoy cierran | Goldens Q1+Q2 deben seguir en dif ±0.5 antes y después (tarea de regresión explícita en cada paso) |
| R5 | R4-por-empresa Q2 sigue divergiendo (motivo del recorte HU-20/G2-D1) y contamina el espejo R4 | El espejo R4 está supeditado a veredicto T0 propio; si diverge, R4 queda en follow-up y esta HU cierra solo R1 |
| R6 | Sobredimensionado: convertir toda la escritura a espejo de una vez | Alcance acotado a `Reporte Componentes R1` (+ `Reversion Pagos R4` si T0 lo confirma); resto de hojas intactas |

---

## 2. DESIGN

### 2.1 Enfoque: modelo espejo

**Lectura (nuevo).** El lector espejo produce por ASE una secuencia ordenada de `FilaEspejoR1 { Firma, ValoresPorColumna }`, donde `Firma` = tupla normalizada de etiquetas (cols A/B/D/E) que identifica el rol de la fila (p. ej. `E='Vlr Servicio'`, `D='E'∧E='Total'`, `A='Componente'∧B='Total'`, `A='Total'∧B=∅`), y `ValoresPorColumna` = valores numéricos por encabezado de columna (detección por header, nunca por índice fijo — invariante del proyecto; ASE4: columna `Especiales` ausente → 0). La zona Componente y la fila de cierre se delimitan por etiqueta (`Componente/Total`, `Total` final), igual que hoy (§V1), pero **sin exigir cardinalidades**: la secuencia observada ES la especificación del período.

**Tabla espejo congelada (T0).** Para cada ASE, T0 publica: (a) firmas observadas en Q1-julio, Q2-julio y agosto-Prueba2; (b) correspondencia fuente-fila → destino-fila (desplazamiento +1 por encabezado de bloque, a confirmar); (c) filas de resumen/validación con fórmula y sus rangos referenciados; (d) invariantes de cierre presentes en las tres muestras (`Componente/Total`, `Mes`, `Subs/Cont`, `Total`). Solo lo presente en la evidencia entra al código.

**Escritura (ruta A).** Por bloque ASE: (1) comparar longitud actual del bloque (plantilla dimensionada al período de referencia) vs longitud de la secuencia fuente del período actual; (2) insertar/borrar filas del bloque preservando estilos y **preservando fórmulas** (mandato del Detalle de plantilla); (3) escribir valores en las celdas de la secuencia (mismo orden, desplazamiento de bloque confirmado por T0); (4) **reanclar** toda fórmula cuyo rango intersecte el bloque mutado (desplazamiento de referencias por delta de filas); (5) reposicionar las filas de resumen/validación relativas al fin del bloque del período. El guard anti-fórmula existente (`EscribirValorNumerico`) sigue vigente: si una celda destino de valor es fórmula tras la mutación → `ERR-PLANTILLA`, nunca sobrescritura silenciosa.

**Regla de ausencia.** Fila ausente en la fuente ⇒ fila suprimida en el destino (no 0). Fila presente con valor 0 ⇒ se escribe 0 (distinguir "ausente" de "cero", doctrina HU-12, invertida de fail-fast a supresión solo dentro del espejo). Cierre: asserts de presencia de las filas invariantes; si falta una invariante → fail-fast nombrando ASE+reporte+fila esperada.

### 2.2 Alternativas y trade-offs (riesgo exigido por el encargo)

| Eje | A — Mutar filas + reanclar (recomendada) | B — Regenerar hoja desde fuente + plantilla de fórmulas |
|---|---|---|
| Qué se hace | Insert/delete de filas OpenXML dentro de cada bloque ASE + desplazamiento de referencias de fórmulas afectadas | Reconstruir `Reporte Componentes R1` por período: filas de datos desde la fuente + filas de fórmula desde una plantilla versionada por (forma del período) |
| Preservación | Máxima: la plantilla y sus fórmulas validadas HU-07..HU-20 se conservan; solo se mueven referencias | Media: las fórmulas viven en una plantilla de fórmulas que hay que mantener y versionar por forma |
| Blast radius | Acotado a bloques espejo + reanclaje; resto de hojas intacto | Toda la hoja se regenera cada período; cualquier divergencia de plantilla es regresión total |
| Complejidad | Reanclaje OpenXML (desplazar referencias en fórmulas que cruzan el bloque) — T0 debe inventariarlas | Generador de hoja + versionado de plantillas de fórmulas — más código nuevo, más superficie de invención |
| Cuándo gana | Bloques apilados con fórmulas locales y referencias inventariables (hipótesis de trabajo) | Reanclaje intratable (refs cruzadas masivas, rangos con nombre, fórmulas externas al bloque) |
| Veredicto | **Principal (D-A)** | **Fallback acotado (tarea T-alt)** si T0 demuestra R3 |

### 2.3 El desplazamiento apilado (tratamiento explícito)

Los 5 bloques ASE están apilados verticalmente en `Reporte Componentes R1`. Insertar/borrar filas en el bloque del ASE-N desplaza en ±Δ todas las filas físicas de los bloques N+1..5 y, con ellas, **toda dirección absoluta** de los 5 mapas aguas abajo que caiga por debajo del punto de mutación. Por eso:

1. T0 publica el **inventario de celdas absolutas afectadas** por bloque (qué celdas de `PorAse/Q2/PorEmpresa/Interventoria/Validaciones` caen dentro o debajo de cada bloque ASE).
2. La implementación procesa los bloques **de abajo hacia arriba (ASE 5 → ASE 1)** o recomputa offsets acumulados, de modo que cada mutación vea direcciones ya estabilizadas — la tarea T4 fija una de las dos técnicas, sin mezclarlas.
3. Las filas de resumen/validación con fórmula se direccionan **relativas al fin del bloque del período**, nunca absolutas.
4. Los goldens 5-ASE Q1+Q2 son el detector: cualquier desplazamiento no compensado rompe la paridad ±0.5 aguas abajo.

### 2.4 Contratos afectados (resumen)

| Contrato | Cambio |
|---|---|
| Core (nuevo): `FilaEspejoR1` / `BloqueEspejoAseInputs` (secuencia de firmas + valores) | Modelo puro, sin I/O; firma exacta la fija T0 |
| `IWorkbookLeafInputReader`: lector espejo R1 (+ R4 si T0 confirma) | Método nuevo; `LeerLEspecialesMenores` convive durante la migración y se retira solo cuando el espejo lo absorba (tarea con grep de cierre) |
| `OpenXmlPlantillaWriter`: mutación de filas + reanclaje + escritura por secuencia | Capacidad nueva acotada a bloques espejo; `ObtenerOCrearCelda`/`EscribirValorNumerico` intactos (guard anti-fórmula vigente) |
| `IValidador`/`ValidadorBasico`: asserts de cierre espejo + inventario de fórmulas reancladas | Nuevas validaciones; validaciones 2.7 existentes intactas |
| `ProcesadorPeriodo`: orquestación espejo por ASE (arriba→abajo en lectura, 5→1 o con offsets en escritura según T4) | Integración; composición DetRetri HU-12 intacta |
| `Remuneracion.IntegrationTests`: goldens + regresión 2026082 | Tests nuevos; 259/259 + 24/24 como red ciega |

### 2.5 Architecture Validation Certificate

| Principio | Veredicto |
|---|---|
| **SRP** | ✅ Lectura (secuencia tipada), mutación (filas), reanclaje (fórmulas) y validación (cierre) son costuras separadas; el lector espejo es método/clase nueva, no una rama `if agosto` dentro de `LeerLEspecialesMenores`. |
| **OCP** | ✅ Los bloques espejo se dimensionan por parámetro (la forma de la fuente), no por mapas nuevos por período: esta HU es la última que debería necesitar un mapa congelado por quincena para R1. |
| **DIP** | ✅ Cambios vía interfaces (`IWorkbookLeafInputReader`, `IPlantillaWriter`/`OpenXmlPlantillaWriter`, `IValidador`); `ProcesadorPeriodo` sigue dependiendo de abstracciones. |
| **Best practices** | ✅ Fuente-define-la-forma (mandato del Detalle + Prompt Vo); período-dominio-gobierna intacto; fail-fast con ASE+reporte+celda fuera del espejo; fórmulas preservadas, solo valores escritos. |
| **Performance** | ✅ Sin impacto: mismas fuentes, una pasada de lectura + una mutación por bloque; T0 y goldens reutilizan `Remuneracion.IntegrationTests`. |

---

## 3. SPEC

### 3.1 Requirements

| ID | Requisito | Grupo | Criterio de aceptación |
|---|---|---|---|
| R-T0-1 | T0 re-deriva la tabla espejo por ASE (firmas Q1-julio, Q2-julio, agosto-Prueba2) con dumps fuente-vs-destino | T0 | Tabla publicada en el repo (evidencia) con conteos 38/52/46/69/44 (Q1) y 45/75/43/73/52 (Q2) reproducidos desde disco, más conteos agosto |
| R-T0-2 | T0 publica el veredicto espejo R4 (`Reversion Pagos R4` 1:1 de `ReversiónPorComponente`) | T0 | Veredicto escrito: R4 entra a esta HU o pasa a follow-up (R5) |
| R-T0-3 | T0 publica el inventario de fórmulas que referencian cada bloque + veredicto A/B | T0 | Veredicto escrito D-A confirmado o T-alt activada |
| R-T0-4 | T0 publica el inventario de celdas absolutas aguas abajo afectadas por bloque | T0 | §2.3 puntos 1-2 implementables sin adivinar |
| R-E-1 | El lector espejo reproduce la secuencia fuente exacta (orden de firmas A–E, valores por header) sin exigir cardinalidades | Espejo | 10/10 combinaciones ASE×quincena-julio dif ±0.5 vs goldens; agosto sin `ERR-VALIDACION` |
| R-E-2 | El bloque destino queda dimensionado a la fuente del período actual (insert/borrar con estilos y fórmulas preservadas) | Espejo | Longitud del bloque por ASE = longitud fuente del período; goldens julio bit-compatibles ±0.5 |
| R-E-3 | Toda fórmula que referencie el bloque mutado queda reanclada con su semántica intacta | Espejo | Assert estructural de fórmulas reancladas + goldens; ningún `ERR-PLANTILLA` espurio en julio ni agosto |
| R-E-4 | Filas de resumen/validación posicionadas relativas al bloque del período | Espejo | TOTAL/OPORTUNO/EXTEMPORÁNEO cierran ±0.5 en las 10 combinaciones + agosto |
| R-E-5 | Regla de ausencia: fila ausente = suprimida (nunca 0); invariantes de cierre presentes o fail-fast | Espejo | Agosto procesa sin simular el bloque ENEL ausente; invariante ausente → error que nombra ASE+reporte+fila |
| R-E-6 | ASE4 sin columna `Especiales` → 0; resto de columnas por header dinámico | Espejo | Golden ASE4 Q1+Q2 + agosto verdes sin columna |
| R-R-1 | Goldens Capa A Q1+Q2 verdes (regresión, dif ±0.5) | Transversal | `Remuneracion.IntegrationTests` verdes; 259/259 + 24/24 como red ciega |
| R-R-2 | Caso 2026082 modo 5 ASE verde (hoy rojo con `ERR-VALIDACION` en ASE 1 L20) | Transversal | Test de regresión dedicado: procesa `Docs/Prueba2/Insumos`, sin excepción de cardinalidad, valores L validados vs salida manual cuando exista |
| R-R-3 | `project-context.md` documenta doctrina espejo + derogación puntual D-C | Transversal | Diff acotado; sin cambios de código en el mismo commit |

### 3.2 Scenarios (Given/When/Then)

- **S1 (T0-espejo):** Given las fuentes R1 de julio-Q1/Q2 y agosto por ASE, When T0 vuelca firmas y las difiere contra `Reporte Componentes R1` de los goldens, Then publica la tabla espejo por ASE con conteos reproducidos (38/52/46/69/44; 45/75/43/73/52; + agosto).
- **S2 (lectura-julio):** Given período 202607-1 o 202607-2, When se lee el espejo R1, Then la secuencia por ASE reproduce la fuente exacta y los goldens cierran ±0.5.
- **S3 (lectura-agosto):** Given período 2026082 (`Docs/Prueba2/Insumos`), When se lee el espejo R1 del ASE 1, Then NO lanza `ERR-VALIDACION` por `TotalD_E` ocurrencia 2: la secuencia observada (sin bloque ENEL negativo) es la especificación.
- **S4 (escritura-dimensionado):** Given bloque ASE dimensionado a julio y fuente de agosto más corta/larga, When se escribe, Then el bloque queda de la longitud de agosto con estilos/fórmulas preservadas y resumen relativo cerrando ±0.5.
- **S5 (reanclaje):** Given fórmulas que referencian el bloque mutado (inventario T0), When se muta, Then las referencias se desplazan por el delta y su semántica se conserva (assert estructural + goldens).
- **S6 (ausencia):** Given fuente sin una fila no-invariante, When se escribe, Then la fila se suprime sin 0 simulado y el cierre se mantiene.
- **S7 (invariante-ausente):** Given fuente sin `Componente/Total` (o `Mes`/`Subs/Cont`/`Total` según T0), When se procesa, Then fail-fast que nombra ASE+reporte+fila esperada.
- **S8 (regresión):** Given goldens julio verdes antes del cambio, When se aplica la HU completa, Then siguen verdes ±0.5 y el caso 2026082 pasa de rojo a verde.

---

## 4. TASKS

| ID | Tarea | Grupo | Depende de | Verificación |
|---|---|---|---|---|
| T0 | **PR 1 bloqueante — Discovery espejo:** (a) dumps fuente-vs-destino `Recaudoporcomponente` ↔ `Reporte Componentes R1` por ASE en Q1-julio, Q2-julio y agosto-Prueba2 (firmas A/B/D/E, conteos, desplazamiento de bloque); (b) veredicto espejo R4 con dumps `ReversiónPorComponente` ↔ `Reversion Pagos R4`; (c) inventario de fórmulas que referencian cada bloque + veredicto A/B (§2.2); (d) inventario de celdas absolutas aguas abajo por bloque (§2.3); (e) tabla espejo congelada por ASE + invariantes de cierre | — | — | Veredictos escritos T0a..T0e; si (a) refuta el 1:1 o (c) demuestra R3 → NEEDS_CONTEXT con recorte (D-D) o T-alt; sin veredictos no empieza T1 |
| T1 | Modelo Core `FilaEspejoR1`/`BloqueEspejoAseInputs` (firma + valores por columna, puro, sin I/O) según tabla T0 | Espejo | T0 | Compila en `Remuneracion.Core` sin deps nuevas; build 0 warnings |
| T2 | Lector espejo R1 (+ R4 si T0b lo incluye): secuencia por firma con headers dinámicos, `Especiales` opcional (ASE4 → 0), sin exigir cardinalidades; convive con `LeerLEspecialesMenores` | Espejo | T0, T1 | S2/S3; goldens julio ±0.5; agosto sin `ERR-VALIDACION` |
| T3 | Mutación de filas + reanclaje en `OpenXmlPlantillaWriter` acotados a bloques espejo (técnica 5→1 u offsets acumulados, §2.3 punto 2 — una sola, sin mezclar) + reposicionamiento relativo de resumen/validación; guard anti-fórmula vigente | Espejo | T0c/d | S4/S5; ningún `ERR-PLANTILLA` espurio en julio |
| T4 | Integración en `ProcesadorPeriodo` (orquestación espejo por ASE) + asserts de cierre e inventario (R-E-5) + reubicación mecánica de celdas absolutas aguas abajo según inventario T0d | Espejo | T2, T3 | S6/S7; goldens 5-ASE julio ±0.5 |
| T-alt | *(Solo si T0c demuestra R3)* Regenerar hoja desde fuente + plantilla de fórmulas versionada por forma; sustituye a T3 | Espejo-alt | T0c | S4/S5 por regeneración; goldens julio ±0.5 |
| T5 | Retiro de `LeerLEspecialesMenores`/mapa rol-ocurrencia una vez absorbido por el espejo (o su acotación documentada si algún slot no-espejo subsiste) + grep de cierre | Espejo | T4 | `grep "Ocurrencia" + "LMenoresPorAse"` acotado a lo que T0 declare subsistente; sin código muerto |
| T6 | Tests: goldens Capa A Q1+Q2 (regresión continua) + test de regresión 2026082 modo 5 ASE en `Remuneracion.IntegrationTests` (xUnit + FluentAssertions) | Transversal | T2-T5 | R-R-1/R-R-2; `dotnet build` 0 warnings; 259/259 + 24/24 como red ciega |
| T7 | Actualizar `.opencode/project-context.md` (doctrina espejo, derogación puntual D-C, tabla de bloques por ASE) | Transversal | T0-T5 | Diff acotado a doctrina/reglas; sin código en el mismo commit |

**Orden sugerido:** T0 → T1 → T2 → T3 (o T-alt) → T4 → T5 → T6 → T7. T6 corre en cada paso (regresión continua: julio verde antes y después de cada tarea). Build 0 warnings antes de marcar cada tarea completa. Sin commits (los hace el Ingeniero). **Work units encadenadas en PRs:** PR 1 = T0 (bloquea todo); PR 2 = T1+T2 (lectura); PR 3 = T3 (+T-alt si aplica) + T4 (escritura+integración); PR 4 = T5+T6+T7 (cierre, tests, docs).

---

## 5. Resumen ejecutivo, riesgos y decisiones

- **Resumen:** el plan elimina la fragilidad estructural en su raíz: los mapas absolutos con ocurrencias congeladas de julio no pueden sobrevivir a fuentes cuya cardinalidad de bloques cambia cada quincena (agosto lo demuestra con el `ERR-VALIDACION` en ASE 1/L20). El modelo espejo invierte la dependencia —la fuente del período actual define la forma— y trata explícitamente el costo real del cambio (desplazamiento apilado + reanclaje de fórmulas) en vez de esconderlo bajo otro mapa congelado.
- **Riesgo principal:** R2 (desplazamiento apilado) + R3 (reanclaje intratable) — contenidos por T0 (inventarios + veredicto A/B) y por la ruta fallback B ya acotada (T-alt).
- **Decisión para el Ingeniero:** D-A ruta A vs B (recomendado: A; el plan viene redactado para A, T-alt es el único cambio si T0 demuestra R3 o si el Ingeniero elige B de entrada). Aprobación del plan = aceptación de D-A..D-F (§0.3).

### DoD (Definition of Done)

1. T0 publicado con tabla espejo + veredictos (bloquea todo lo demás).
2. `dotnet build Remuneracion.WinForms/Remuneracion.WinForms.slnx` con 0 warnings.
3. Goldens Capa A Q1+Q2 verdes (dif ±0.5) + 259/259 + 24/24 sin regresión.
4. Test de regresión 2026082 modo 5 ASE verde (el caso que hoy falla).
5. Ninguna fórmula sobrescrita (guard anti-fórmula + assert estructural verdes).
6. `project-context.md` actualizado; sin datos inventados (toda celda nueva remite a evidencia T0).

### Estado de cierre — PR 4 (T5+T6+T7), 2026-10-02

- **PR 1 (T0)** ✅ — evidencia en `plans/21 - T0 Evidencia.md`.
- **PR 2 (T1+T2)** ✅ — modelo Core + lector espejo.
- **PR 3 (T3+T4)** ✅ — mutación/reanclaje + integración; CRITICAL #1 corregido en auditoría.
- **PR 4 (T5+T6+T7)** ✅ con un bloqueo declarado fuera de scope:
  - **T5** ✅ — path legado rol/ocurrencia retirado (`LeerLEspecialesMenores`, `LMenoresPorAse/Q2`, `RolLMenor`, `ObtenerLMenores`, `LEspecialesMenores`, `LEspecialesMenoresAseInputs`, `EscribirCeldasLEspecialesMenores`, carve-out W-5). Grep de cierre: 0 usos activos (solo historial en comentarios). Absorción probada 15/15 en `EspejoR1AbsorcionTests` (celdas L del mapa caen en el bloque; agosto: espejo gobierna toda la columna L).
  - **T6** ✅ — goldens Capa A Q1+Q2 ±0.5 verdes; regresión end-to-end 2026082 vía `ProcesadorPeriodo` con insumos reales de `Docs/Prueba2/Insumos`: aserta que ya NO falla por el mapa L-menores (`ERR-VALIDACION`) y **captura el blocker de SALDOS POR NOTA de ASE2** (HU-11/Q2). Flujo 5-ASE con espejo Δ≠0 real ({−3,−9,−6,+6,+8}) verifica workbook generado, fórmulas preservadas e invariantes. Suite **299/299**, build 0 warnings.
  - **T7** ✅ — doctrina espejo + derogación puntual D-C + tabla de deltas + follow-ups documentados en `.opencode/project-context.md`.
- **Blocker declarado (fuera del alcance del Plan 21):** el período 2026082 trae otra divergencia estructural en `SaldosaFavorAplicadosPorNotas` (HU-11/Q2): el mapa por `Ase.Id` exige `Vlr Intereses` para ASE2 y la fuente de agosto no la trae. El flujo completo de `ProcesadorPeriodo` se detiene ahí (fail-fast 2.5), ANTES de escribir; no es el defecto que el Plan 21 resuelve. Requiere su propio T0/HU (patrón análogo: mapa por ASE vs forma de la fuente). No se inventan datos para sortearlo.
- **Follow-ups:** R4 por empresa Q2 (recorte HU-20/G2-D1); implementación del espejo para `Reversion Pagos R4` (declarado "entra" por T0b, motor solo R1); CF/DV (`conditionalFormatting`/`dataValidations`, W-4); W-5 end-to-end (no rearmable de forma estable por el blocker de SALDOS POR NOTA).
