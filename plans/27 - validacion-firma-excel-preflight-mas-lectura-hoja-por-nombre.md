# Plan 27 — Validación de firma Excel en el preflight + lectura de RESUMEN MES por nombre de hoja

> **Alcance:** dos unidades sin acoplamiento. **Unidad A (YA APROBADA en sesión 2026-10-05, se incluye como primera unidad):** `LeerRecaudoEmpresa` deja de leer la primera hoja a ciegas y resuelve `RESUMEN MES` por nombre (sobrecarga existente `LeerFilas(ruta, nombreHoja)`). **Unidad B (núcleo del plan, a aprobación):** el preflight del Plan 26 (`ValidadorInsumosPeriodo`) valida la **firma** de TODOS los insumos que ya resuelve (R1/R2/R4 por ASE + optativos Q2 + 5 conciliaciones + R10): un archivo no-Excel (MHTML/HTML renombrado, corrupto) se detecta ANTES de procesar, con mensaje administrativo en la MISMA lista numerada de faltantes (numeración única, QUÉ/DÓNDE/QUÉ HACER, sin jerga).
> **Rector (IRRENUNCIABLE):** `README.md` + `Docs/Propuesta_Proyecto_Automatizacion_Remuneracion_UAESP.md` + `Docs/Detalle de plantilla.docx` + `Docs/Prompt_Maestro_Proyecto_Remuneracion_UAESP _ Vo .docx`. **Invariantes del proyecto:** NUNCA sobrescribir fórmulas (solo valores); tolerancia ±0.5; ASE4 sin columna "Especiales" (0 si ausente); Q2 aplica SALDOS POR NOTA/RETRIBUCION NEGATIVA (`Periodo.NumeroQuincena == 2`, `AjustesSfT = 0` en Q1); redondear a entero antes de DetRetri/columna D; build 0 warnings; finales de línea CRLF; fail-fast nombrando archivo/carpeta; tests SOLO con insumos reales (`Docs/Insumos`, `Docs/Prueba2`), NO fixtures sintéticas salvo MHTML generado en temp (nunca en el repo); sin commits (los hace el Ingeniero con `#commit`); sin emojis.
> **Continuidad:** HU-01..HU-20 cerradas; Planes 21 (espejo R1), 23 (opcionalidad 2.5), 25 (roles por firma) y 26 (preflight, suite **318/318**) cerrados. Este plan NO reabre ninguna semántica de lectura, cálculo, escritura ni validación fuera de lo declarado: **cero cambios de fórmulas de negocio; Q1/Q2-julio intactos por construcción. Los finders runtime quedan intactos; el plumbing UI/CLI no se toca; la doctrina Plan 26 se extiende, no se reemplaza.**
> **Insumos canónicos (no se re-fijan, no se inventan datos):** `Docs/Insumos/REMUNERACION 2026071`, `Docs/Insumos/REMUNERACION 2026072` (goldens ±0.5), `Docs/Prueba2/Insumos` (período 2026082 Q2 con `Conciliaciones/` YA corregida por el usuario + `R10_Remuneracion_2026082.xlsx`).
> **Numeración:** `plans/` 01..26 ocupados; este plan toma el primer correlativo libre, **27**.
> **Estado:** IMPLEMENTADO — CERRADO (2026-10-05). Unidad A (T1) + Unidad B (T2) implementadas; T3 (tests) y T4 (docs) cerradas. Build 0/0; suite **339/339** (base real 316/318 — ver §6). Sin commits (los hace el Ingeniero con `#commit`).
> **Fecha:** 2026-10-05

---

## 0. Clarification Gate

**Encargo ya decidido en sesión (Unidad A aprobada; Unidad B con diseño pedido explícito): no hay preguntas bloqueantes.** El plan viene redactado con las decisiones D-A..D-G (§0.3); si el Ingeniero discrepa de alguna, solo cambia la tarea acotada que la implementa. Punto de bloqueo respetado: NO se implementa nada en este documento.

### 0.1 Qué está verificado y qué NO (honestidad técnica)

**Verificado en esta planificación** (lectura directa de código/disco, 2026-10-05):

| # | Hecho | Método | Consecuencia |
|---|---|---|---|
| V1 | `LeerRecaudoEmpresa` lee la primera hoja a ciegas: `ExcelWorksheetNavigator.LeerFilas(rutaConciliacion)` (`ExcelDataReaderWorkbookLeafInputReader.cs:890`) | Lectura código | El defecto runtime es real al símbolo: con 3 hojas cae en `Oportuno` (detalle crudo) y no encuentra los 3 bloques |
| V2 | Ya existe la sobrecarga correcta `LeerFilas(ruta, nombreHoja)` (`ExcelWorksheetNavigator.cs:52`): match case-insensitive, fail-fast que nombra archivo+hoja ("nunca lee la primera hoja a ciegas", doctrina R10 HU-20/G3) | Lectura código | Unidad A = 1 línea + comentario; sin código nuevo de I/O |
| V3 | Los 15 archivos de conciliación en disco son xlsx válidos (`50 4B 03 04`, ZIP/PK) y los 15 contienen una hoja de nombre exacto `RESUMEN MES` (sin necesidad de normalización de espacios; el match case-insensitive ya cubre) | Inspección ZIP (`xl/workbook.xml`) de los 15 archivos | Unidad A funciona en julio (hoja única, primera posición) y agosto (2a/3a posición); no hace falta trim/normalización |
| V4 | Agosto trae multi-hoja donde julio trae una sola: `Conjunta Recip -082026` = `Oportuno \| EXTEMPORANEO \| RESUMEN MES`; `Conjunta Otros -082026` = `Oportuno \| RESUMEN MES \| RESUMEN`; `Directa-082026` = `OPORTUNO \| EXTEMPORANEOS \| RESUMEN MES` (nótese el plural en las hojas de detalle; `RESUMEN MES` exacto en las 5) | Inspección ZIP | Confirma que solo el match por nombre exacto es robusto; leer "la segunda/tercera" sería otro mapa congelado |
| V5 | Ningún `.xlsx` inválido queda en `Docs/Prueba2/Insumos/Conciliaciones/` (los 5 abren como ZIP válido; el MHTML renombrado ya fue corregido por el usuario) | Firmas §V3 + apertura ZIP | El caso "no-Excel" de los tests se construye con MHTML real en carpeta temporal (nunca en el repo) |
| V6 | Superficie exacta de `ValidadorInsumosPeriodo.Validar` (`ValidadorInsumosPeriodo.cs:34-111`): por ASE (1..5) R1, R2, R4 (con fallback `Reversión/Reversion`), banco, balance, + saldos-notas/retribución SOLO Q2; de período: 5 conciliaciones SIEMPRE + R10 SIEMPRE; `RecaudosReversados` fuera | Lectura código | Superficie de firma = exactamente esos paths resueltos (ni uno más ni uno menos) |
| V7 | `ValidadorInsumosPeriodo` ya hace I/O filesystem BCL (`Directory.Exists`, L121); Core NO referencia Infrastructure (solo BCL + dominio) | Lectura código + csproj | Leer los primeros bytes con `FileStream` BCL en Core NO rompe la estratificación (mismo nivel que el `Directory.Exists` existente); no se necesita ni ExcelDataReader ni OpenXML ni el localizador para la firma |
| V8 | Punto de integración: `ProcesadorPeriodo.Ejecutar` L106-121 (preflight tras resolver carpetas ASE, antes del loop; `Log.Warning` + throw único `ERR-FUENTE-NO-ENCONTRADA` con `FormateadorInsumosFaltantes.Mensaje`); guardrail single-ASE `File.Exists` en `ProcesadorRemuneracion.Ejecutar` | Lectura código | La firma se enchufa en el MISMO throw y formato; costo UI/CLI = 0 |
| V9 | `ExcelReaderFactory.CreateReader(stream)` genérico en ambas sobrecargas del Navigator (sin switch de formato); proyecto en ExcelDataReader 3.9.0 (`Remuneracion.Infrastructure.csproj`) | Lectura código | El reader auto-detecta xls/xlsx: aceptar OLE legacy como válido tiene evidencia en el propio código |
| V10 | `FormateadorInsumosFaltantes` es el sitio único del mensaje (D-E Plan 26): `Componer` numera `1)..n)` con QUÉ/DÓNDE/QUÉ HACER; hay factorías por caso (carpeta ASE, reporte ASE, conciliación, R10) | Lectura código | Los ítems "archivo inválido" se agregan como factorías nuevas en el mismo sitio, misma numeración |
| V11 | `plans/` 01..26 ocupados; suite base **318/318** (cierre Plan 26) | Listado + Plan 26 §cierre | Numeración 27; red ciega 318/318 |

**Aportado por la sesión (NO re-verificado en disco en esta planificación; se acredita en tests):** cadena de eventos (MHTML renombrado → `HeaderException — Invalid file signature` tardía; corrida siguiente → `CalculoInvalidoException — No se encontraron los 3 bloques` en Recip agosto); layout del RESUMEN MES de agosto idéntico al de julio (verificado fila a fila con tooling por el Ingeniero); valor testigo ASE2 OPORTUNO col F = 18.262.430.

### 0.2 Mapeo al Rector (in vs out)

**Entra:**

| Requisito | Superficie de cambio |
|---|---|
| Unidad A (aprobada): `LeerRecaudoEmpresa` resuelve `RESUMEN MES` por nombre | 1 línea (`ExcelDataReaderWorkbookLeafInputReader.cs:890`) + actualización del comentario T0-0.6 (ya no "hoja única") |
| Unidad B: clasificador de firma puro (bytes → veredicto) + lectura acotada a ≤512 B por archivo resuelto | Helper nuevo en Core (BCL puro, sin deps nuevas); función pura testeable sin archivos |
| Unidad B: el preflight valida la firma de cada path resuelto y agrega ítems "archivo inválido" a la MISMA lista numerada, mismo código `ERR-FUENTE-NO-ENCONTRADA`, sin procesar nada | `ValidadorInsumosPeriodo` (chequeo tras cada finder no-nulo) + 2 factorías en `FormateadorInsumosFaltantes` (web-copia vs formato desconocido) |
| Tests con insumos reales + MHTML en temp; goldens julio ±0.5 intactos | `Remuneracion.IntegrationTests` (xUnit + FluentAssertions); suite 318/318 como red ciega |
| Trazabilidad doctrinal mínima | Diff acotado en `.opencode/project-context.md` (+ fila del manual si el mensaje nuevo lo exige) |

**Sale (EXPLÍCITO):** validación de CONTENIDO/headers/estructura (sigue en runtime como hoy); validación de NOMBRES de hoja en el preflight (abriría workbooks — non-goal, ver D-G); cambios a finders, readers (salvo L890), cálculo, escritura, validaciones o R10-oráculo; reinterpretar valores de agosto; paquetes NuGet; commits (los hace el Ingeniero con `#commit`).

### 0.3 Decisiones ejecutivas (aprobación = aceptar estas)

| ID | Decisión |
|---|---|
| D-A | **Unidad A YA APROBADA en sesión:** `LeerFilas(ruta)` → `LeerFilas(ruta, "RESUMEN MES")` en `LeerRecaudoEmpresa`. Julio: mismo resultado (hoja única con ese nombre, §V3). Agosto: lee la hoja correcta. Hoja ausente → fail-fast existente que nombra archivo+hoja. Alcance: SOLO conciliaciones; los demás blind reads (R1/R2/R4/banco/balance de fuente mono-hoja) no se tocan. |
| D-B | **La firma se valida en el preflight, sobre los paths ya resueltos, DESPUÉS del chequeo de existencia.** Cada finder no-nulo → 1 lectura acotada (≤512 B). Faltante e inválido conviven en la MISMA lista con numeración única y orden estable (el ítem inválido ocupa el lugar del archivo). Mismo código `ERR-FUENTE-NO-ENCONTRADA`: un archivo ilegible equivale a un insumo no utilizable; costo UI/CLI = 0 por construcción (§V8). |
| D-C | **La lectura de bytes vive en Core como helper BCL puro** (p. ej. `InspectorFirmaExcel`, nombre a elección): `File.OpenRead` + buffer ≤512 B, sin ExcelDataReader, sin OpenXML, sin referencia a Infrastructure. Justificación: el validador ya hace I/O BCL (§V7); la alternativa "el localizador confirma y el validador recibe veredictos" partiría la enumeración en dos sitios y duplicaría la superficie (rechazada). El clasificador `byte[] → veredicto` es función pura (testeable sin archivos). |
| D-D | **xls legacy (OLE `D0 CF 11 E0`) se ACEPTA como Excel válido.** Evidencia: `CreateReader` genérico del proyecto lo soporta (§V9). Rechazarlo inventaría un requisito que el runtime no tiene. |
| D-E | **Dos mensajes administrativos, un solo formato.** Variante web-copia: *"El archivo '...' no es un Excel válido (parece una copia de una página web). Solicite el archivo real al área encargada, colóquelo en la carpeta '...' y vuelva a ejecutar."* Variante formato desconocido: *"El archivo '...' no es un Excel válido. ..."* (misma acción). Prohibido en el mensaje al usuario: "magic bytes", "MIME", "firma binaria", "header", "MHTML", códigos hex. El detalle técnico (bytes detectados, firma clasificada) va SOLO al `Log.Warning` técnico. |
| D-F | **Single-ASE incluido:** el guardrail de `ProcesadorRemuneracion` (R1/R2/R4) valida firma con el mismo helper y formateador (`MensajeAse`); nunca exige banco/balance/conciliaciones/R10 (doctrina 26 vigente). |
| D-G | **Sin acoplamiento A↔B y sin sheet-check en preflight:** la firma es preflight (existencia+formato, sin abrir workbooks); la hoja por nombre es runtime (Unidad A). El preflight NO verifica nombres de hoja (requeriría abrir workbooks: non-goal explícito, posible follow-up con su propio T0). |

---

## 1. PROPOSE

### 1.1 Intent

Que un archivo no-Excel (una captura web renombrada, un corrupto) falle en segundos y de una sola vez —en el preflight, antes de abrir el primer workbook— con un mensaje que el usuario administrativo entiende y puede accionar (qué archivo, dónde va el real, qué hacer), en vez de descubrirse tras procesar los 5 ASE con un `HeaderException` técnico; y que la lectura de conciliaciones deje de depender del orden de hojas del archivo (RESUMEN MES por nombre, no "la primera").

### 1.2 In Scope

- Unidad A: fix por nombre + tests con el insumo real de agosto (3 bloques leídos, col F correcta) + goldens julio intactos.
- Unidad B: helper de firma en Core + extensión del validador (superficie §V6) + 2 factorías de mensaje (D-E) + integración al throw existente + `Log.Warning` técnico con la firma detectada.
- Tests: 15 conciliaciones reales = firma válida; Prueba2 con archivos corregidos → preflight pasa y el flujo avanza más allá de la lectura de conciliaciones; MHTML/HTML en temp → mensaje administrativo, cero procesamiento.
- Docs: `project-context.md` (doctrina firma + hoja-por-nombre) + fila del manual si aplica.

### 1.3 Non-Goals

Ver §0.2 "Sale". En particular: NO se abre ningún workbook en el preflight (la firma son bytes, no contenido); NO se cambia ningún mensaje existente; NO se tocan finders ni el resto de readers; NO se crea ningún MHTML en el repo.

### 1.4 Riesgos

| ID | Riesgo | Mitigación |
|---|---|---|
| R-PERF | 41 archivos × apertura en Q2 (5 ASE × 7 + 5 conciliaciones + R10) | Solo ≤512 B por archivo, sin parseo (milisegundos); el path feliz ya abre esos mismos archivos completos después |
| R-FALSO-POSITIVO | Un xlsx válido marcado inválido (firma parcial, stream cifrado con firma distinta) | Regla mínima: PK (`50 4B 03 04`) u OLE (`D0 CF 11 E0`) = válido; todo lo demás = inválido con mensaje accionable (el usuario revalida con el área encargada); los 15 reales + goldens son la red |
| R-DERIVA | Runtime pasa a resolver un insumo que el preflight no firma-chequea (o viceversa) | D-B: el chequeo vive DENTRO del validador sobre cada path resuelto (no lista paralela); checklist de paridad preflight↔runtime en SPEC |
| R-LENGUAJE | Jerga técnica filtrada al mensaje administrativo | Doctrina D-E + test que aserta la lista prohibida (`magic bytes`, `MIME`, `firma binaria`, `header`, `MHTML`, `0x`) sobre el texto de ítems, análogo al test de lenguaje del Plan 26 |
| R-SOBREDISENO | Convertir la firma en framework de validación de contenido | Alcance congelado a magic bytes (D-C); cualquier validación de contenido futura requiere su propio T0/HU |

---

## 2. DESIGN

### 2.1 Enfoque: bytes en el preflight, nombre en el runtime

**Unidad A (1 línea).** `ExcelDataReaderWorkbookLeafInputReader.cs:890`:

```csharp
var filas = ExcelWorksheetNavigator.LeerFilas(rutaConciliacion, "RESUMEN MES");
```

El comentario T0-0.6 (L874-887, "tienen UNA sola hoja") se actualiza: julio = hoja única con ese nombre; agosto = multi-hoja con ese nombre en 2a/3a posición (§V3/V4). Detección de bloques, pares de columnas por quincena y fail-fast de bloques ausentes: intactos.

**Unidad B (puerta de bytes con los mismos paths).** Modelo: el clasificador puro

- `FirmaExcelValida` (PK `50 4B 03 04` — xlsx; OLE `D0 CF 11 E0` — xls, D-D),
- `NoExcel_PareceCopiaWeb` (prefijo de texto con marcadores: `<`, `<!doctype`, `<html`, `MIME-Version`, `Saved by`, `From:` — case-insensitive sobre el prefijo ASCII),
- `NoExcel_FormatoDesconocido` (resto),
- (defensivo) `Ilegible` (I/O error al leer → se trata como desconocido, nunca throw propio: el runtime ya tiene su fail-fast).

Lectura: UN `File.OpenRead` por path resuelto, buffer único ≤512 B (constante; nunca el archivo entero), cierre inmediato (sin bloqueo para el runtime posterior). Dónde: helper Core BCL (D-C), invocado por `ValidadorInsumosPeriodo` tras cada finder no-nulo —incluidas las 5 conciliaciones y el R10— y por el guardrail single-ASE (D-F). Cada `NoExcel_*` → 1 ítem vía factoría nueva del formateador (D-E), con `QueFalta` = el archivo con su nombre real, `DondeDebeIr` = la carpeta donde debe ir el real, `QueHacer` = solicitar el real + reejecutar. Orden estable idéntico al actual (el ítem inválido ocupa la posición del archivo). `Log.Warning` previo al throw con `{Archivo}`, `{FirmaDetectada}` y primeros bytes en hex (jerga solo al log).

**Mensaje congelado (plantilla, §3.1 R-F-2):** encabezado vigente del Plan 26 + ítems numerados. Ejemplo variante web (tono del encargo): `2) El archivo 'Conjunta Recip -082026 - Sigab 7 de sept. 2026.xlsx' no es un Excel válido (parece una copia de una página web). Solicite el archivo real al área encargada, colóquelo en la carpeta "Conciliaciones" y vuelva a ejecutar.` Ejemplo variante desconocida: `2) El archivo '...' no es un Excel válido. Solicite el archivo real al área encargada, colóquelo en la carpeta "..." y vuelva a ejecutar.`

### 2.2 Alternativas y trade-offs

| Eje | Elegida | Descartada |
|---|---|---|
| Dónde leer bytes | Helper Core BCL (D-C): mismo nivel que el `Directory.Exists` existente, cero deps nuevas | Localizador Infrastructure que "confirma": partiría la enumeración, el localizador devuelve paths no veredictos, y el mensaje vive en Core |
| Cuánto leer | ≤512 B una vez (magia binaria + sniff de texto acotado) | 8 bytes estrictos: insuficientes para distinguir web-copia de desconocido sin segunda lectura; archivo entero: prohibido por el encargo |
| Código de error | Reutilizar `ERR-FUENTE-NO-ENCONTRADA` (D-B): mismo throw, UI/CLI intactos | Código nuevo: exigiría rama en `CatalogoErrores` + `CodigoSalidaPara` + matriz de tests por cero ganancia observable (el usuario ve el mismo formato) |
| xls OLE | Aceptar (D-D, evidencia §V9) | Rechazar: inventaría un requisito contra el runtime que sí lo lee |
| Sheet-check en preflight | NO (D-G): abriría workbooks | Hacerlo: viola el non-goal del Plan 26 y su costo (bloqueo, lentitud, duplicar readers) |

### 2.3 Por qué NO hay desplazamiento, reanclaje ni riesgo geométrico (tratamiento explícito)

Porque no se toca ningún workbook, ninguna fila, ninguna fórmula ni ningún mapa: Unidad A cambia QUÉ hoja se lee (mismo layout, misma detección de bloques); Unidad B es enumeración de filesystem + ≤512 B por archivo + un throw antes del loop. La prueba dura: con insumos válidos el flujo es bit-idéntico al actual (goldens Q1/Q2 no pueden moverse por construcción).

### 2.4 Contratos afectados (resumen)

| Contrato | Cambio |
|---|---|
| `ExcelDataReaderWorkbookLeafInputReader.LeerRecaudoEmpresa` (L890) + comentario T0-0.6 | Solo hoja por nombre (D-A); bloques, columnas por quincena, fail-fast intactos |
| Core (nuevo): helper de firma (clasificador puro + lectura acotada BCL) | Creación; sin deps nuevas |
| `ValidadorInsumosPeriodo` + `FormateadorInsumosFaltantes` | Chequeo de firma por path resuelto + 2 factorías (D-E); existencia y orden intactos |
| `ProcesadorRemuneracion` (guardrail) | Misma firma-chequeo en R1/R2/R4 (D-F) |
| `ProcesadorPeriodo`, `CatalogoErrores`, `Form1`, `EjecutorCli`, finders, readers, writer | Sin cambios |
| `Remuneracion.IntegrationTests` | Tests nuevos con insumos reales + MHTML en temp; suite 318/318 como red ciega |
| `.opencode/project-context.md` (+ manual si aplica) | Doctrina firma + hoja-por-nombre |

### 2.5 Architecture Validation Certificate

| Principio | Veredicto |
|---|---|
| **SRP** | ✅ Una sola razón de cambio por costura: el helper clasifica bytes, el validador enumera, el formateador redacta, el lector resuelve la hoja. Sin ramas `if agosto` (la regla es por nombre/firma, válida en todo período). |
| **OCP** | ✅ La firma se absorbe por path resuelto (parámetro), no por tipo de insumo: el próximo insumo que el runtime pase a resolver la hereda sin bifurcar el validador. |
| **DIP** | ✅ Cambios detrás de `IWorkbookLeafInputReader` y del servicio de preflight existente; UI/CLI no conocen la regla. |
| **Best practices** | ✅ Fail-fast al inicio con archivo/carpeta nombrados (doctrina extendida a formato); quincena por dominio; fórmulas nunca tocadas; mensaje para personas, detalle técnico al log; sin fixtures sintéticas en el repo. |
| **Performance** | ✅ Preflight: 41 aperturas de ≤512 B (milisegundos, sin parseo); path con inválidos ahorra el procesamiento completo de 5 ASE (el desperdicio de la sesión). |

---

## 3. SPEC

### 3.1 Requirements

| ID | Requisito | Grupo | Criterio de aceptación |
|---|---|---|---|
| R-A-1 | `LeerRecaudoEmpresa` lee `RESUMEN MES` por nombre (D-A) | Unidad A (aprobada) | Con `Docs/Prueba2/Insumos/Conciliaciones/Conjunta Recip -082026...xlsx`: 3 bloques, valores col F correctos (testigo ASE2 OPORTUNO F=18.262.430); julio bit-idéntico |
| R-A-2 | Hoja ausente → fail-fast que nombra archivo+hoja | Unidad A | Sobrecarga existente sin cambios; sin test con fixture (rama ya cubierta por doctrina HU-20/G3) |
| R-F-1 | Preflight valida la firma de CADA path resuelto (superficie §V6: R1/R2/R4+banco+balance siempre, +saldos/retri solo Q2, 5 conciliaciones, R10) antes de procesar | Unidad B | Con Prueba2 corregida: CERO ítems inválidos; con 1 MHTML en temp (patrón real `Saved by Blink`): 1 ítem variante web |
| R-F-2 | Mensaje administrativo en la MISMA lista numerada (numeración única, QUÉ/DÓNDE/QUÉ HACER, D-E); cero jerga | Unidad B | `grep -i "magic\|MIME\|firma binaria\|header\|MHTML\|0x"` vacío sobre ítems; cada ítem con verbo accionable + reejecutar |
| R-F-3 | Con inválidos: UN `ERR-FUENTE-NO-ENCONTRADA`, cero workbooks abiertos, cero salidas, misma salida UI/CLI (código 2, línea grepable) | Unidad B | Test: MHTML en temp → throw previo al primer `LeerR1`, sin archivo de salida creado |
| R-F-4 | xls OLE aceptado como válido (D-D); PK aceptado; texto/otros rechazados con variante correcta | Unidad B | Clasificador puro con tests de bytes (sin archivos): PK→válido, OLE→válido, `<html`/MIME→web, resto→desconocido |
| R-F-5 | Single-ASE: firma en R1/R2/R4 con mismo código y formateador (D-F); nunca exige otros insumos | Unidad B | Test: ruta con HTML en temp → mismo código + `MensajeAse` |
| R-F-6 | Paridad preflight↔runtime: todo path que el runtime resuelve pasa por firma-chequeo y viceversa | Conservación | Grep de cierre; finders y throws runtime sin diff funcional |
| R-R-1 | Suite base 318/318 verde antes y después (`dotnet build` 0 warnings) | Transversal | Build + suite completos en cada WU |
| R-R-2 | Goldens Q1+Q2 ±0.5 intactos; Q1 no exige reportes Q2 | Transversal | Capa A verde; preflight Q1/Q2 julio = CERO faltantes + CERO inválidos |
| R-R-3 | Prueba2 corregida: preflight pasa y el flujo avanza más allá de la lectura de conciliaciones | Transversal | Test de avance (supera `LeerRecaudosEmpresa` sin `CalculoInvalidoException` de bloques) |
| R-R-4 | `project-context.md` con doctrina firma + hoja-por-nombre (diff acotado, sin código en el mismo commit) | Transversal | Diff solo doctrina/reglas |

### 3.2 Scenarios (Given/When/Then)

- **S1 (agosto-Recip, Unidad A):** Given `Conjunta Recip -082026...xlsx` (3 hojas), When `LeerRecaudosEmpresa` Q2, Then 3 bloques leídos desde `RESUMEN MES`, col F correcta (testigo 18.262.430).
- **S2 (julio intacto, Unidad A):** Given las 10 conciliaciones de julio (hoja única), When se leen, Then valores idénticos a hoy + goldens ±0.5.
- **S3 (MHTML en preflight, Unidad B):** Given período con 1 conciliación = MHTML real en temp (patrón `Saved by Blink`), When `ProcesadorPeriodo.Ejecutar`, Then UN `ERR-FUENTE-NO-ENCONTRADA` previo al primer `LeerR1` con el ítem variante web en la lista numerada; ninguna salida creada.
- **S4 (formato desconocido):** Given 1 R1 con bytes no-Excel/no-texto en temp, When preflight, Then ítem variante desconocida (sin "página web").
- **S5 (todo válido):** Given `Docs/Prueba2/Insumos` corregida, When preflight Q2, Then lista vacía y el flujo supera la lectura de conciliaciones.
- **S6 (single-ASE):** Given `ProcesadorRemuneracion` con R2 = HTML en temp, When `Ejecutar`, Then mismo código + ítem en `MensajeAse`, antes de leer nada.
- **S7 (lenguaje):** Given cualquier lista con inválidos, When se lee el detalle, Then sin jerga prohibida; el log técnico sí registra firma/bytes.
- **S8 (cero-geometría):** Given el diff, When se audita, Then ningún mapa, fórmula ni mensaje existente cambió salvo L890+comentario y las factorías nuevas.

---

## 4. TASKS

| ID | Tarea | Grupo | Depende de | Verificación |
|---|---|---|---|---|
| T0 | **Evidencia T0 (CERRADA en planificación, NO es PR):** firmas PK 15/15, `RESUMEN MES` exacto 15/15 (§V3/V4), superficie del validador (§V6), sobrecarga por nombre (§V2), punto de integración (§V8) | Fase 0 | — | §0.1 V1..V11; sin NEEDS_CONTEXT |
| T1 | Unidad A: L890 → `LeerFilas(ruta, "RESUMEN MES")` + comentario T0-0.6 actualizado + tests S1/S2 (agosto 3 bloques + col F; julio goldens ±0.5) | Fix-A | T0 | R-A-1/R-A-2; build 0 warnings |
| T2 | Unidad B: helper firma Core (clasificador puro + lectura ≤512 B) + chequeo en `ValidadorInsumosPeriodo` (superficie §V6) + 2 factorías D-E + guardrail single-ASE (D-F) + `Log.Warning` técnico; throws runtime intactos | Fix-B | T0 | Compila en Core sin deps nuevas; S3/S4/S6 a nivel de validador; build 0 warnings |
| T3 | Tests con insumos reales: (a) 15 conciliaciones = firma válida, Prueba2 CERO inválidos (R-F-1, S5); (b) MHTML/HTML en temp → variante web + throw previo a `LeerR1` sin salida (R-F-3, S3); (c) bytes desconocidos → variante desconocida (S4); (d) clasificador puro PK/OLE/texto/resto (R-F-4); (e) single-ASE S6; (f) avance Prueba2 más allá de conciliaciones (R-R-3); (g) lenguaje S7 + paridad R-F-6 + goldens (R-R-1/R-R-2) | Transversal | T1, T2 | SPEC completa; `dotnet build Remuneracion.WinForms/Remuneracion.WinForms.slnx` 0 warnings; suite ≥318/318 |
| T4 | Docs: `project-context.md` (doctrina firma-preflight + hoja-por-nombre + jerga solo al log) + fila del manual si el mensaje nuevo lo exige | Transversal | T3 | Diff acotado; S8 citado |

**Orden sugerido:** T0 (cerrada) → T1 → T2 → T3 → T4. T1 puede shippear sola (WU-1, ya aprobada); T2+T3+T4 en WU-2 secuencial. Build 0 warnings antes de marcar cada tarea completa. Sin commits (los hace el Ingeniero).

**Estado de ejecución (2026-10-05):** T0 ✅ (re-verificada, §6.1) · T1 ✅ (Unidad A) · T2 ✅ (Unidad B) · T3 ✅ (21 tests nuevos + 2 reparados) · T4 ✅ (docs). Build 0/0; suite 339/339.

---

## 5. Resumen ejecutivo, riesgos y decisiones

- **Resumen:** la sesión dejó dos defectos de distinta capa con una sola corrida de evidencia. Unidad A (aprobada): leer "la primera hoja" funcionó mientras hubo una sola; agosto trae 3 y el fix es usar la sobrecarga por nombre que la doctrina HU-20/G3 ya exige (1 línea, julio bit-idéntico por §V3). Unidad B: el MHTML renombrado se descubrió tarde porque la existencia no implica formato; la puerta del Plan 26 se extiende con una lectura de ≤512 B por insumo resuelto (Core BCL, sin abrir workbooks, sin deps) que clasifica PK/OLE/web/desconocido y reporta en la misma lista numerada con el mismo código. Con insumos válidos el flujo es bit-idéntico (cero-geometría §2.3).
- **Riesgo principal:** R-FALSO-POSITIVO (regla mínima PK/OLE + red de 15 reales y goldens) + R-DERIVA (chequeo dentro del validador, checklist R-F-6) + R-LENGUAJE (test de jerga prohibida análogo al Plan 26).
- **Decisión para el Ingeniero:** ratificar D-A..D-G (§0.3). D-A ya aprobada en sesión (shippea en WU-1 sin esperar a B). **Sin preguntas bloqueantes pendientes; una sola no-bloqueante con recomendación:** el umbral de 512 B y los marcadores de texto exactos (§2.1) — recomendación: fijarlos en implementación y traerlos a revisión en el PR de WU-2.

### DoD (Definition of Done)

1. `dotnet build Remuneracion.WinForms/Remuneracion.WinForms.slnx` con 0 warnings.
2. Suite completa verde sin regresión (base 318/318) + goldens Capa A Q1+Q2 en dif ±0.5 intactos.
3. Prueba2 corregida: preflight CERO inválidos y el flujo avanza más allá de la lectura de conciliaciones (R-R-3, S1/S5).
4. MHTML/HTML en temp → UN `ERR-FUENTE-NO-ENCONTRADA` previo a todo I/O Excel, variante web en la lista numerada, sin salida creada (S3/S7).
5. Ninguna fórmula sobrescrita ni leída; diff acotado a L890+comentario, helper Core, validador y 2 factorías (S8).
6. `project-context.md` actualizado; ningún MHTML creado en el repo; sin commits del agente; CRLF; sin emojis.

### Follow-ups explícitos (fuera de esta HU)

- Validación de NOMBRES de hoja en preflight (requeriría abrir workbooks): posible HU futura con su propio T0.
- Si UAESP entrega xls legacy real, la rama OLE queda validada end-to-end (hoy aceptada por evidencia de `CreateReader` + test del clasificador).

---

## 6. Cierre de implementación (2026-10-05)

### 6.1 Re-verificación T0 — la base NO era 318/318

El plan asumía suite base **318/318**. La realidad medida en implementación fue **316/318**: dos tests del Plan 26 (`Preflight_Prueba2SinConciliaciones_ListaLosCincoDeUnaVez_YNoInicia` y `Preflight_Mensaje_SinJergaTecnica_ConNombresReconocibles`) fallaban porque la corrección de la carpeta `Conciliaciones/` de `Docs/Prueba2/Insumos` por el usuario (la misma corrección que `§V5` declara) invalidó su premisa "Prueba2 SIN Conciliaciones". **V5 re-verificado en disco:** los 5 archivos de `Docs/Prueba2/Insumos/Conciliaciones/` son xlsx válidos (firma PK) y `Conjunta Recip -082026...` trae 3 hojas (`Oportuno | EXTEMPORANEO | RESUMEN MES`). Se preservó la intención de ambos tests reproduciendo el escenario incompleto sobre una **copia temporal** de Prueba2 (insumos reales, nunca fixture en repo). Sin este ajuste la suite no podía quedar verde.

### 6.2 Detalles menores fijados en implementación (§5, no-bloqueante)

- **Umbral de lectura: 512 B** (`InspectorFirmaExcel.BytesALeer`, constante). Firma binaria (4 B) + sniff de texto acotado para MHTML real (`From:` / `MIME-Version:` / `<html`).
- **Marcadores de copia web** (case-insensitive sobre el prefijo): `<` (prefijo recortado, tras BOM/espacios), `<!doctype`, `<html`, `mime-version`, `saved by`, `from:`.
- **Ilegible** (error de I/O al leer la firma) se trata como **formato desconocido** (ítem administrativo), nunca throw propio del inspector.
- **Dos factorías** en `FormateadorInsumosFaltantes` (`ArchivoNoEsExcelPareceWeb` / `ArchivoNoEsExcelFormatoDesconocido`) sobre un texto base único; el ítem no-Excel usa `Alcance` = el del archivo que ocupa (ASE 1..5 / Período).

### 6.3 Evidencia fresca (DoD)

| Check | Resultado |
|---|---|
| `dotnet build Remuneracion.WinForms/Remuneracion.WinForms.slnx` | **0 warnings, 0 errors** |
| `dotnet test Remuneracion.IntegrationTests` | **339/339 verde** (316 base + 2 reparados + 21 nuevos) |
| Golden Q1/Q2 ±0.5 | intactos (suite completa verde; cero-geometría) |
| Prueba2 2026082 corregida | preflight **CERO faltantes/inválidos**; el flujo **completa END-TO-END** (supera la lectura de conciliaciones y cierra con el R10) |
| MHTML en temp | UN `ERR-FUENTE-NO-ENCONTRADA` previo a todo I/O Excel; variante web en la lista; sin salida creada |
| 15 conciliaciones reales | firma válida 15/15 |
| Archivos tocados | `ExcelDataReaderWorkbookLeafInputReader.cs` (L890 + comentario), `InspectorFirmaExcel.cs` (nuevo), `ValidadorInsumosPeriodo.cs`, `FormateadorInsumosFaltantes.cs`, `ProcesadorRemuneracion.cs`, tests + docs |

### 6.4 Prueba2 — recorrido real

Con `Docs/Prueba2/Insumos` corregida, el preflight Q2 pasa (0 faltantes / 0 inválidos) y el período 2026082 corre hasta el final sin error, contra su propio R10 (`R10_Remuneracion_2026082.xlsx`). No reaparece el `CalculoInvalidoException` de los 3 bloques de `Conjunta Recip -082026` (Unidad A lo resuelve leyendo `RESUMEN MES` por nombre). Detalle en los tests `LecturaResumenMesPorNombre_RecipAgosto_LeiTresBloquesYColF` y `FlujoPrueba2Corregida_AvanzaMasAlaDeConciliaciones`.
