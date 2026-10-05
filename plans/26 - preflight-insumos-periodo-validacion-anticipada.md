# Plan 26 — Preflight de insumos del período: validación completa ANTES de procesar, con reporte agregado en lenguaje administrativo

> **Alcance:** un preflight de insumos (fail-fast al inicio, sin abrir ningún workbook) que verifica la existencia de TODOS los archivos fuente del período —por ASE y de período— ANTES de procesar el primer ASE, y ante faltantes lanza UN solo error `ERR-FUENTE-NO-ENCONTRADA` con la lista numerada completa en lenguaje claro para el usuario administrativo (qué falta, dónde debe ir, qué hacer). Cubre `ProcesadorPeriodo.Ejecutar` (caso que motiva la HU) y `ProcesadorRemuneracion.Ejecutar` con la misma doctrina, acotado a lo que cada flujo realmente consume.
> **Rector (IRRENUNCIABLE):** `README.md` + `Docs/Propuesta_Proyecto_Automatizacion_Remuneracion_UAESP.md` + `Docs/Detalle de plantilla.docx` + `Docs/Prompt_Maestro_Proyecto_Remuneracion_UAESP _ Vo .docx`. **Invariantes del proyecto:** NUNCA sobrescribir fórmulas (solo valores); tolerancia ±0.5; ASE4 sin columna "Especiales" (0 si ausente); Q2 aplica SALDOS POR NOTA/RETRIBUCION NEGATIVA (`Periodo.NumeroQuincena == 2`, `AjustesSfT = 0` en Q1); redondear a entero antes de DetRetri/columna D; build 0 warnings; finales de línea CRLF; fail-fast nombrando archivo/carpeta; tests SOLO con insumos reales (`Docs/Insumos`, `Docs/Prueba2`), NO fixtures sintéticas; sin commits (los hace el Ingeniero con `#commit`); sin emojis.
> **Continuidad:** HU-01..HU-20 cerradas; Planes 21 (espejo R1), 23 (opcionalidad `Vlr Intereses` 2.5) y 25 (roles R1-Q2 por firma) cerrados, suite base **310/310** sin commit. Este plan NO reabre ninguna semántica de lectura, cálculo, escritura ni validación: **cero cambios de fórmulas de negocio; Q1/Q2-julio intactos por construcción. El preflight solo verifica existencia en filesystem; los finders y readers runtime quedan intactos.**
> **Insumos canónicos (no se re-fijan, no se inventan datos):** `Docs/Insumos/REMUNERACION 2026071`, `Docs/Insumos/REMUNERACION 2026072` (períodos completos: preflight debe reportar CERO faltantes), `Remuneracion 202607-1 Total.xlsx` / `Remuneracion 202607-2 Total.xlsx` (goldens ±0.5), `Docs/Prueba2/Insumos` (caso real que motiva la HU: 5 carpetas ASE + `R10_Remuneracion_2026082.xlsx`, SIN carpeta `Conciliaciones/`).
> **Numeración:** `plans/` 01..25 ocupados (24 = evidencia T0 del Plan 25); este plan toma el primer correlativo libre, **26**.
> **Estado:** CERRADO (WU-1 + WU-2, 2026-10-05) — ruta A ratificada (D-B); implementado y verificado (build 0 warnings, suite 318/318). Sin commit (lo hace el Ingeniero con `#commit`).
> **Fecha:** 2026-10-05

---

## 0. Clarification Gate

**Una sola decisión con fork real para el Ingeniero (D-B en §0.3):** ante faltantes, ¿**A (reutilizar `ERR-FUENTE-NO-ENCONTRADA` con detalle numerado, recomendado)** o **B (código nuevo del catálogo, p. ej. `ERR-INSUMOS-FALTANTES`)**? El plan viene redactado en la ruta A (costo UI/CLI = 0, trazado con evidencia en §2.2); si el Ingeniero elige B, solo cambia la entrada del catálogo + mapeo de salida + matriz de tests (§3.1 R-F-5, tarea T-alt). Todo lo demás —validador, superficie verificada, mensaje en lenguaje claro, gates— es idéntico en ambas rutas. Punto de bloqueo respetado: NO se implementa nada en este documento.

### 0.1 Qué está verificado y qué NO (honestidad técnica)

**Verificado en esta planificación** (lectura directa de código/disco):

| # | Hecho | Método | Consecuencia |
|---|---|---|---|
| V1 | `ProcesadorPeriodo.Ejecutar` procesa los 5 ASE completos (loop L103-243: R1/R2/R4, espejo, conciliación Q1, banco, balance, saldos/retri Q2, DetRetri) y RECIÉN en L250-254 lee las hojas `Recaudo *` desde `Conciliaciones/` vía `LeerRecaudosEmpresa` | Lectura `ProcesadorPeriodo.cs:103-258` | El defecto reportado es real al símbolo: un insumo de período conocible en el segundo 0 se descubre tras procesar 5 ASE que luego se descartan |
| V2 | `LeerRecaudosEmpresa` itera `EmpresaFacturacion.Catalogo` y ante el primer `BuscarConciliacion(...) == null` lanza `ArchivoFuenteNoEncontradoException` con el mensaje literal `No se encontró el archivo de conciliación de {empresa.Nombre} (prefijo '{PrefijoConciliacion}') en la carpeta Conciliaciones/ del período.` | Lectura `ExcelDataReaderWorkbookLeafInputReader.cs:94-115` | Hoy se descubre UN faltante por corrida (el primero del catálogo en orden); confirma el segundo síntoma del encargo |
| V3 | La lectura `Recaudo *` (L250-254) está FUERA de toda rama por quincena; la conciliación por empresa (`LeerConciliacionEmpresas`, derivada de los R1/R2/R4 ya leídos, sin archivo extra) es la que está acotada a Q1 (L161-171, HU-20-T0b) | Lectura `ProcesadorPeriodo.cs:154-171,245-258` | **Los 5 archivos de conciliación se exigen SIEMPRE (Q1 y Q2)** — punto que el encargo pedía re-verificar: confirmado, sin ambigüedad |
| V4 | `ArchivoFuenteLocator`: `BuscarArchivo` (L38, prefijo case-insensitive, `TopDirectoryOnly`), `BuscarConciliacion` (L62: `Path.Combine(carpetaPeriodo, "Conciliaciones")`, match por prefijo, `TopDirectoryOnly`, `null` si no existe la carpeta), `ObtenerCarpetasAse` (prefijos `1-`..`5-`), finders `BuscarR10` (prefijo `R10_`), `BuscarReporteBanco`, `BuscarBalance` (base + variante `-Optimizado_`), `BuscarSaldosNotas` / `BuscarRetribucionNegativa` (matcher normalizado agnóstico a diacríticos) | Lectura `ArchivoFuenteLocator.cs` + `ILocalizadorArchivosAse.cs` | El preflight puede reutilizar TODOS los finders vía la interfaz existente (sin duplicar lógica de match); la superficie verificable por filesystem es exactamente la resuelta por estos finders |
| V5 | `EmpresaFacturacion.Catalogo` = 5 empresas con `PrefijoConciliacion` (`Conjunta Recip`, `Conjunta ENEL`, `Conjunta ENERBIT`, `Directa`, `Conjunta Otros`) | Lectura `EmpresaFacturacion.cs:54-111` | Las 5 claves de búsqueda de conciliaciones; base de los nombres legibles del mensaje (§2.1) |
| V6 | Catálogo: `CodigoError.FuenteNoEncontrada = "ERR-FUENTE-NO-ENCONTRADA"`; `CatalogoErrores.Para(FuenteNoEncontrada, ex)` → título `Archivo fuente faltante` + guía `Falta un archivo fuente del período: {ex.Message}...`; `CodigoSalidaPara` → `FuenteOPlantilla` (2). UI: `Form1.MostrarErrorUx` muestra `MessageBox($"{guia} Código: {codigo}.")` + log `ERROR [{codigo}]: ...` con formato `[{HH:mm:ss}]` (HU-14); CLI: `stderr $"[{codigo}] {guia}"` + línea grepable `RESULTADO ERROR codigo=... salida=... runId=...` (HU-15) | Lectura `CodigoError.cs`, `CatalogoErrores.cs`, `Form1.cs:348-352,485-520`, `EjecutorCli.cs:99-108,180-186` | **Reutilizar el código existente = costo UI/CLI cero** (evidencia para D-B, §2.2); el detalle numerado viaja en `ex.Message` y ambas guías lo envuelven sin cambios |
| V7 | `ProcesadorRemuneracion.Ejecutar` (single-ASE) recibe `RutaR1/R2/R4` explícitas, valida no-vacías y lee INMEDIATAMENTE (L55-69); NO consume banco, balance, saldos, retri, conciliaciones ni R10 (flujo: leer → calcular → leaf → validar → escribir) | Lectura `ProcesadorRemuneracion.cs:37-101` | El preflight single-ASE es solo guardarrail de existencia de esos 3 paths; la resolución por locator YA es fail-fast previo en `Form1.EjecutarModoUnAse` (L534-536) y `EjecutorCli.EjecutarUnAse` (L152-158) |
| V8 | `Docs/Prueba2/Insumos/` = 5 carpetas ASE + `R10_Remuneracion_2026082.xlsx`, SIN carpeta `Conciliaciones/`; `1-Promoambiental/` trae el set completo por ASE (R1, R2, R4, banco, balance, saldos-notas, retribución-negativa + PDFs) | Listado de directorio | Caso real reproducible: el preflight debe listar los faltantes de `Conciliaciones/` de una vez sin tocar ningún workbook (test R-F-3) |
| V9 | `RecaudosReversados` NO es consumido por el runtime: sin método en `ILocalizadorArchivosAse`, sin lectura en ningún procesador (grep: solo menciones en `ANT EXT-REV` protegida y docs de requerimientos) | Grep `RecaudosReversados\|LeerRecaudos\|ANT EXT` en `*.cs` | **Corrección al encargo (§0.3 D-D):** el preflight NO debe exigir `RecaudosReversados`; exigirlo sería inventar un requisito que el runtime no resuelve |
| V10 | `plans/` 01..25 ocupados; este plan es el **26**, primer número libre | Listado `plans/` | Numeración del documento |

**Re-verificación en implementación (WU-1, 2026-10-05):** V1..V10 quedaron constados contra el código final. La puerta construida es `ValidadorInsumosPeriodo` (Core, reutiliza los finders vía `ILocalizadorArchivosAse`) + `FormateadorInsumosFaltantes` (sitio único del mensaje administrativo) + interconexión fail-fast en `ProcesadorPeriodo.Ejecutar` (tras resolver las 5 carpetas ASE, antes del loop) y guardrail `File.Exists` en `ProcesadorRemuneracion.Ejecutar`. Se confirmó lo previsto: `RecaudosReversados` quedó FUERA (V9), las 5 conciliaciones se exigen en AMBAS quincenas (V3) y los `?? throw` runtime permanecen intactos (D-C). Los tests nuevos viven en `Remuneracion.IntegrationTests/PreflightInsumosTests.cs` (Capa A + CLI + paridad finder↔runtime). Cierre: build **0 warnings / 0 errors**; suite **318/318** verde.

**Aportado por el Ingeniero (NO re-verificado en disco en esta planificación; se acredita en tests):** suite base 310/310 verde en el working tree; caso real 2026082 con 5 ASE procesados y descartados por falta de `Conciliaciones/`; tono esperado del mensaje (ejemplo del encargo).

### 0.2 Mapeo al Rector (in vs out)

**Entra:**

| Requisito | Superficie de cambio |
|---|---|
| Validador de insumos de período (Core service + contrato en Core) que, a partir de carpeta del período + quincena + catálogo de empresas y usando `ILocalizadorArchivosAse`, produce la lista COMPLETA de faltantes (solo existencia en filesystem, sin abrir workbooks) | Clase nueva en `Remuneracion.Core/Services/` + modelo de faltante en `Remuneracion.Core/Models/`; depende solo de `ILocalizadorArchivosAse` (DIP intacto) |
| Conexión al inicio de `Ejecutar` en `ProcesadorPeriodo` (tras guards de carpeta/plantilla + resolución de carpetas ASE, antes del loop L103); guardrail de existencia en `ProcesadorRemuneracion` | 2 puntos de llamada; el proceso NO inicia (no lee, no calcula, no escribe) si hay faltantes |
| UN solo error `ERR-FUENTE-NO-ENCONTRADA` con detalle numerado en lenguaje administrativo, integrado a la guía de pantalla existente UI/CLI sin tocar `CatalogoErrores`, `Form1` ni `EjecutorCli` | Solo el `ex.Message` del throw; formato `[HH:mm:ss]`, prefijo de código y línea grepable se conservan por construcción (§V6) |
| Tests con insumos reales: preflight sobre `Docs/Prueba2/Insumos` (lista TODOS los faltantes de `Conciliaciones/` de una vez) + preflight sobre `Docs/Insumos` Q1/Q2 (CERO faltantes) + Q1 no exige reportes Q2 | `Remuneracion.IntegrationTests` (xUnit + FluentAssertions); suite 310/310 como red ciega |
| Trazabilidad doctrinal mínima | Diff acotado en `.opencode/project-context.md` (+ 1 fila en `Docs/Manual-Uso-App-Administrativa.md`, que ya documenta el error de conciliación simple) |

**Sale (EXPLÍCITO):** validación de CONTENIDO (headers, estructura, coherencia numérica — siguen fallando en runtime como hoy); cambios a finders, readers, cálculo, escritura, validaciones o R10-oráculo; `RecaudosReversados` como insumo exigido (§V9); nuevo código del catálogo salvo T-alt si el Ingeniero elige B; reinterpretar valores de agosto; paquetes NuGet; commits (los hace el Ingeniero con `#commit`).

### 0.3 Decisiones ejecutivas (aprobación = aceptar estas)

| ID | Decisión |
|---|---|
| D-A | **Fail-fast al inicio, antes de cualquier I/O Excel.** En `ProcesadorPeriodo`, el preflight corre tras los guards de carpeta/plantilla y la resolución de las 5 carpetas ASE, y antes del loop por ASE. En `ProcesadorRemuneracion`, guardrail de existencia (`File.Exists`) de las 3 rutas explícitas al inicio. Sin faltantes → el flujo continúa idéntico; con faltantes → un solo throw, cero workbooks abiertos, cero salidas parciales. |
| D-B | **Un solo error con el código EXISTENTE `ERR-FUENTE-NO-ENCONTRADA` y detalle numerado en lenguaje claro (ruta A).** Justificación con evidencia: el catálogo, la guía (`Archivo fuente faltante`), el log `ERROR [{codigo}]`, el MessageBox, el stderr `[{codigo}]` y la línea `RESULTADO ERROR` ya existen y ya mapean a salida 2; un código nuevo exigiría rama en `CatalogoErrores.Para` + `CodigoSalidaPara` + matriz de tests + verificación UI/CLI (§2.2). T-alt acota el costo si el Ingeniero elige B. |
| D-C | **El preflight REUTILIZA los finders vía `ILocalizadorArchivosAse`; no duplica lógica de match ni reemplaza la resolución runtime.** Mismo prefijo, mismo case-insensitive, misma normalización de diacríticos, mismo `TopDirectoryOnly`. Los `throw` intermedios actuales quedan como red de seguridad (defensa en profundidad); el preflight es la puerta, no el reemplazo. |
| D-D | **Superficie verificada = superficie que el runtime resuelve, ni un archivo más ni uno menos.** Por ASE: R1, R2, R4, banco, balance siempre; + saldos-notas y retribución-negativa SOLO en Q2 (`NumeroQuincena == 2`, gobierno por dominio, nunca por detección de contenido). De período: 5 conciliaciones SIEMPRE (Q1 y Q2, §V3) + R10 SIEMPRE (oráculo obligatorio HU-20/G3 en ambas quincenas). `RecaudosReversados` explícitamente FUERA (§V9). Single-ASE: solo R1/R2/R4 (§V7). |
| D-E | **Doctrina del mensaje (lenguaje administrativo, congelada en §3.1):** encabezado con período + quincena + "No se procesó ningún ASE"; lista numerada donde cada ítem dice QUÉ falta (nombre que el usuario reconoce: ASE, reporte, carpeta), DÓNDE debe ir (ruta/carpeta con el nombre exacto esperado) y QUÉ hacer (pedir/generar el archivo, colocarlo, reejecutar). Prohibida jerga de código (sin "prefijo", "matcher", "finder", "TopDirectoryOnly", nombres de clases o códigos internos salvo el código del catálogo que agrega el frontend). |
| D-F | **Si falta la carpeta `Conciliaciones/` completa, UN solo ítem la nombra + enumera los 5 archivos esperados dentro** (no 6 líneas redundantes). Si la carpeta existe, un ítem por cada archivo de empresa faltante. Misma regla para carpetas ASE ausentes (un ítem nombra la carpeta + los reportes que debería contener). |

---

## 1. PROPOSE

### 1.1 Intent

Que un período con insumos incompletos falle en segundos y de una sola vez —antes de abrir el primer workbook— con un mensaje que el usuario administrativo entiende y puede accionar (qué falta, dónde va, qué hacer), en vez de procesar los 5 ASE para descubrir al final un faltante por corrida con lenguaje técnico (`prefijo 'Conjunta ENEL'`, `carpeta Conciliaciones/`).

### 1.2 In Scope

- `ValidadorInsumosPeriodo` (nombre orientativo) en Core: enumera faltantes por ASE (1..5) y de período usando solo `ILocalizadorArchivosAse`, gobernado por `Periodo.NumeroQuincena`.
- Integración fail-fast en `ProcesadorPeriodo.Ejecutar` + guardrail en `ProcesadorRemuneracion.Ejecutar`, con el error único D-B/D-E.
- Tests con insumos reales: `Docs/Prueba2/Insumos` (agregación completa, lenguaje claro), `Docs/Insumos` Q1/Q2 (cero faltantes), Q1 sin exigencia Q2.
- Docs: `project-context.md` (doctrina preflight) + 1 fila del manual administrativo.

### 1.3 Non-Goals

Ver §0.2 "Sale". En particular: NO se valida contenido (un xlsx corrupto o con headers rotos sigue fallando en su reader con su mensaje actual); NO se cambia ningún mensaje de error existente fuera del nuevo throw agregado; NO se toca la resolución runtime (los `?? throw` actuales permanecen).

### 1.4 Riesgos

| ID | Riesgo | Mitigación |
|---|---|---|
| R-COMPAT | Tests existentes que asertan el texto exacto de un fail-fast intermedio (p. ej. regresión que "captura el fail-fast por nombre") podrían romperse si ahora el throw agregado llega antes | El tipo de excepción y el `Codigo` son idénticos (D-B); la tarea T3 incluye grep de asserts sobre esos mensajes y suite completa verde como gate; solo se ajusta el assert si el cambio de texto es el comportamiento especificado (lista agregada), nunca la lógica |
| R-DERIVA | El preflight y la resolución runtime divergen con el tiempo (nuevo insumo que el runtime exige pero el preflight no lista, o viceversa) | D-C: el preflight llama a los MISMOS finders (no lógica paralela); SPEC exige checklist de paridad finder↔runtime + grep de cierre; `RecaudosReversados` documentado como fuera por la misma regla |
| R-RUIDO | Listas muy largas (p. ej. período vacío: 5 carpetas + decenas de archivos) ilegibles en un MessageBox | D-F (agregación por carpeta ausente) + orden estable (carpetas ASE 1..5, luego período: conciliaciones, R10) + encabezado que declara el total; el log conserva el texto completo |
| R-FALSO-VERDE | Preflight verde pero runtime falla igual (contenido roto): el usuario percibe "validación que no valida" | Non-goal explícito + encabezado honesto del mensaje ("faltan insumos", nunca "insumos correctos"); el contenido sigue validado por sus readers |
| R-SOBREDISENO | Convertir el preflight en framework general de validación de contenido/reglas de negocio | Alcance congelado a existencia en filesystem (D-D); cualquier validación de contenido futura requiere su propio T0/HU |

---

## 2. DESIGN

### 2.1 Enfoque: puerta filesystem con los mismos finders, mensaje para personas

**Modelo (Core puro, sin I/O directo).** Un `InsumoFaltante { Alcance, QueFalta, DondeDebeIr, QueHacer }` (nombres exactos a elección del implementador) + formateador único que produce el detalle numerado. Condición: **un solo sitio** compone el texto en lenguaje administrativo (D-E); el servicio solo enumera hechos (qué finder devolvió `null`).

**Servicio.** `Validar(carpetaPeriodo, periodo, localizador) → IReadOnlyList<InsumoFaltante>` (firma orientativa):

1. Carpetas ASE 1..5 vía `ObtenerCarpetasAse` (misma resolución por `StartsWith(id)` que el loop L106-110; carpeta ausente → 1 ítem D-F, y se omiten sus archivos para no duplicar ruido).
2. Por cada carpeta ASE presente: `BuscarArchivo(R1/R2/R4 con el fallback Reversión/Reversion)`, `BuscarReporteBanco`, `BuscarBalance` (ambos prefijos), y SOLO si `NumeroQuincena == 2`: `BuscarSaldosNotas`, `BuscarRetribucionNegativa`. Cada `null` → 1 ítem con nombre administrativo del reporte (`Recaudo por componente (R1)`, `Detalle de saldos a favor (R2)`, `Reversión por componente (R4)`, `Reporte de recaudo por banco`, `Balance de subsidios y contribuciones`, `Saldos a favor aplicados por notas`, `Retribución negativa`).
3. De período: `BuscarConciliacion` por cada empresa del catálogo (D-F si falta la carpeta: se detecta porque los 5 devuelven `null` + `!Directory.Exists(Conciliaciones)` — el servicio puede consultar el filesystem o el locator expone el hecho; la implementación elige, sin lógica de match propia) + `BuscarR10`.
4. Orden estable: ASE 1..5 (carpeta, luego archivos en el orden del paso 2), luego `Conciliaciones/`, luego R10.

**Integración.** `ProcesadorPeriodo.Ejecutar`, tras L90-96 (carpetas resueltas) y antes del loop L103: `var faltantes = preflight.Validar(...); if (faltantes.Count > 0) throw new ArchivoFuenteNoEncontradoException(CodigoError.FuenteNoEncontrada, FormatearMensaje(...))`. El validador se inyecta por constructor (depende de `ILocalizadorArchivosAse`, el mismo ya inyectado — puede construirse dentro del procesador con el locator existente para no ampliar la firma de composición de UI/CLI; la SPEC lo deja a elección con DIP intacto). `ProcesadorRemuneracion.Ejecutar`, tras el guard L55-63: `File.Exists` de las 3 rutas → ítems agregados con el mismo formateador y código. `Log.Warning` previo al throw (preflight = hito observable, doctrina HU-14 D4) + `progreso?.Report("Verificando insumos del período...")`.

**Mensaje congelado (plantilla, §3.1 R-F-2):** `Faltan insumos para el período {CodigoCompleto} (quincena {1|2}). No se procesó ningún ASE.` + líneas `1) ...`, `2) ...`. Ejemplo (tono del encargo): `1) La carpeta 'Conciliaciones' no existe dentro de la carpeta del período ({carpeta}); debe contener los 5 archivos de conciliación (Reciprocidad EAAB, ENEL, ENERBIT, Occidente Directa, Otros). Pida los archivos al área encargada, cree la carpeta con ese nombre exacto y vuelva a ejecutar.` Cada ítem de archivo: `Falta {reporte} del ASE {n} ({carpetaAse}); el archivo debe estar dentro de '{carpetaAse}' y su nombre debe empezar por '{prefijo visible}' (p. ej. el archivo mensual que entrega {origen}). Colóquelo ahí y vuelva a ejecutar.` (Los `{prefijo visible}`/`{origen}` exactos los fija la implementación dentro de D-E; prohibido inventar orígenes: si no consta en docs, se nombra solo el inicio de nombre esperado.)

### 2.2 Alternativas y trade-offs (ruta A vs B del Clarification Gate)

| Eje | A — Reutilizar `ERR-FUENTE-NO-ENCONTRADA` (recomendada, D-B) | B — Código nuevo (T-alt) |
|---|---|---|
| Qué se hace | El detalle numerado viaja en `ex.Message`; `CatalogoErrores.Para` lo envuelve tal cual (`Falta un archivo fuente del período: {detalle}...`) | Nueva constante + rama en `Para` (título/guía propios) + rama en `CodigoSalidaPara` (¿salida 2 igual? entonces el código solo cambia el texto) |
| Diff | Solo Core (servicio + 2 llamadas); UI/CLI/tests de contrato intactos | + `CodigoError`, `CatalogoErrores` (2 ramas), matriz `CatalogoErroresTests`/`CodigosSalidaCliTests`, verificación MessageBox/stderr |
| Grepabilidad | Idéntica (`ERROR [ERR-FUENTE-NO-ENCONTRADA]`, `[ERR-FUENTE-NO-ENCONTRADA]`, `RESULTADO ERROR codigo=ERR-FUENTE-NO-ENCONTRADA`) | Nueva cadena a indexar en runbooks; el manual ya documenta la actual |
| Cuándo ganaría B | Si el Ingeniero quiere distinguir "faltante previo" de "faltante runtime" en métricas/log o guías distintas por caso | — |
| Veredicto | **Principal (D-B)** | **Fallback acotado (T-alt)** solo si el Ingeniero elige B |

Descartados explícitamente: validar contenido en el preflight (abre workbooks = lento, bloquea archivos, duplica readers; non-goal); un throw por faltante (reproduce el defecto actual); warning-y-continuar (viola atomicidad fail-fast del procesador); nuevo proyecto/assembly (el servicio es Core puro con la interfaz ya existente).

### 2.3 Por qué NO hay desplazamiento, reanclaje ni riesgo geométrico (tratamiento explícito)

Porque no se toca ningún workbook, ninguna fila, ninguna fórmula ni ningún mapa: el preflight es enumeración de filesystem (`Directory`/`EnumerateFiles` vía locator) + un throw antes del loop. La prueba dura: el diff no modifica ningún archivo bajo `Remuneracion.Infrastructure/Excel/` salvo (si acaso) cero líneas — solo agrega servicio/modelo en Core y 2 llamadas en procesadores. Los goldens Q1/Q2 no pueden moverse por construcción (ningún valor cambia cuando no hay faltantes: el flujo continúa idéntico).

### 2.4 Contratos afectados (resumen)

| Contrato | Cambio |
|---|---|
| Core (nuevo): modelo `InsumoFaltante` + servicio de preflight + formateador | Creación; puro, sin deps nuevas (usa `ILocalizadorArchivosAse`, `Periodo`, `EmpresaFacturacion`, `CarpetasAse`) |
| `ProcesadorPeriodo` / `ProcesadorRemuneracion` | Solo la llamada fail-fast al inicio (D-A); guards, loop, cálculo, escritura, validaciones intactos |
| `ILocalizadorArchivosAse` / finders / readers / writer / `IValidador` / R10 | Sin cambios (D-C; defensa en profundidad: sus `throw` permanecen) |
| `CatalogoErrores` / `CodigoError` / `CodigosSalida` / `Form1` / `EjecutorCli` | Sin cambios en ruta A (D-B); T-alt acota B |
| `Remuneracion.IntegrationTests` (xUnit + FluentAssertions) | Tests nuevos con insumos reales; suite 310/310 como red ciega |
| `.opencode/project-context.md` + `Docs/Manual-Uso-App-Administrativa.md` | Diff acotado: doctrina preflight + fila del mensaje agregado |

### 2.5 Architecture Validation Certificate

| Principio | Veredicto |
|---|---|
| **SRP** | ✅ Una sola razón de cambio por costura: el servicio enumera faltantes, el formateador redacta, el procesador aborta, los tests demuestran. Sin ramas `if período` ni `if agosto` (la regla es por quincena del dominio + existencia, válida en todo período presente y futuro). |
| **OCP** | ✅ El próximo insumo que el runtime pase a resolver se agrega llamando a su finder en la lista (parámetro), no bifurcando el procesador; la cardinalidad futura (nuevas empresas, nuevos reportes Q) se absorbe por enumeración, no con código por período. |
| **DIP** | ✅ El servicio depende de `ILocalizadorArchivosAse` (abstracción existente); los procesadores dependen del servicio (o lo componen con el locator ya inyectado); UI/CLI no conocen la regla. |
| **Best practices** | ✅ Fail-fast al inicio con archivo/carpeta nombrados (doctrina del proyecto, ahora agregada); quincena gobernada por `Periodo.NumeroQuincena` (G2-D2); fórmulas nunca tocadas (ni siquiera leídas); tolerancia ±0.5 y goldens intactos por construcción; mensaje para personas, detalle técnico al log. |
| **Performance** | ✅ Sin impacto en el path feliz: solo enumeraciones de directorio (milisegundos, sin abrir workbooks); en el path con faltantes ahorra el procesamiento completo de 5 ASE (el desperdicio que motiva la HU). Tests reutilizan `Remuneracion.IntegrationTests` con insumos reales existentes. |

---

## 3. SPEC

### 3.1 Requirements

| ID | Requisito | Grupo | Criterio de aceptación |
|---|---|---|---|
| R-F-1 | Preflight al inicio de `ProcesadorPeriodo.Ejecutar` (tras guards+carpetas, antes del loop): con faltantes lanza `ArchivoFuenteNoEncontradoException(ERR-FUENTE-NO-ENCONTRADA)` y NO abre ningún workbook, no calcula, no escribe | Fix | Con `Docs/Prueba2/Insumos`: throw antes del primer `LeerR1` (evidencia: log sin `leyendo R1` + ningún archivo de salida creado); sin faltantes: flujo idéntico |
| R-F-2 | Error único con lista numerada COMPLETA en lenguaje administrativo (D-E/D-F): encabezado período+quincena+`No se procesó ningún ASE`; cada ítem con QUÉ/DÓNDE/QUÉ HACER; cero jerga de código | Fix | Mensaje del caso Prueba2 contiene los 5 archivos de conciliación (o el ítem carpeta D-F con los 5 nombrados) + nombres que el usuario reconoce; `grep -i "prefijo\|matcher\|finder\|TopDirectoryOnly"` vacío sobre el texto de ítems |
| R-F-3 | Caso real Prueba2 (2026082, Q2): el preflight lista TODOS los faltantes de `Conciliaciones/` DE UNA VEZ (hoy solo el primero tras procesar 5 ASE) | Fix | Test dedicado: `Ejecutar` lanza en < umbral (sin I/O Excel), `ex.Message` contiene las 5 empresas (nombres administrativos) y la carpeta `Conciliaciones`; el código es `ERR-FUENTE-NO-ENCONTRADA` |
| R-F-4 | Período completo (`Docs/Insumos` Q1 y Q2): CERO faltantes; goldens ±0.5 intactos; Q1 NO exige saldos-notas ni retribución-negativa | Fix | Tests Q1/Q2: lista vacía + goldens Capa A verdes; test Q1 con carpeta sin esos dos reportes (las fuentes reales Q1) no los reclama |
| R-F-5 | Single-ASE (`ProcesadorRemuneracion`): guardrail de existencia de R1/R2/R4 con el mismo código y formateador; NUNCA exige banco/balance/conciliaciones/R10/saldos | Fix | Test: ruta inexistente → mismo código + ítem en lenguaje claro; flujo normal con rutas reales intacto |
| R-F-6 | Finders intactos + defensa en profundidad: ningún `?? throw` runtime se retira ni se modifica; checklist de paridad preflight↔runtime (cada archivo que el runtime resuelve está en la lista del §2.1 y viceversa, salvo `RecaudosReversados` documentado fuera) | Conservación | Grep de cierre: finders y throws runtime sin diff funcional; `RecaudosReversados` ausente de la lista con justificación §V9 citada |
| R-R-1 | Suite base 310/310 verde antes y después (red ciega, `dotnet build` 0 warnings) | Transversal | Build + suite completos en cada WU |
| R-R-2 | Mismo formato observable HU-14/HU-15: log `ERROR [ERR-FUENTE-NO-ENCONTRADA]`, MessageBox con guía + código, stderr `[{codigo}]`, `RESULTADO ERROR codigo=... salida=2` | Transversal | Test CLI in-process con carpeta incompleta → salida 2 + línea grepable (ruta A: sin cambios en `EjecutorCli`) |
| R-R-3 | `project-context.md` con doctrina preflight (puerta filesystem, superficie D-D, mensaje D-E) + fila del manual administrativo actualizada | Transversal | Diff acotado a doctrina/reglas; sin código en el mismo commit |

### 3.2 Scenarios (Given/When/Then)

- **S1 (Prueba2, el caso real):** Given `Docs/Prueba2/Insumos` (5 ASE + R10, sin `Conciliaciones/`), When `ProcesadorPeriodo.Ejecutar` Q2, Then lanza ANTES del primer `LeerR1` con UN `ERR-FUENTE-NO-ENCONTRADA` cuyo detalle enumera los 5 archivos de conciliación (o ítem carpeta D-F) en lenguaje claro; ningún workbook abierto, ninguna salida creada.
- **S2 (Q1 completo):** Given `Docs/Insumos/REMUNERACION 2026071`, When preflight Q1, Then lista vacía (no reclama saldos-notas/retribución-negativa) y el golden Q1 cierra ±0.5.
- **S3 (Q2 completo):** Given `Docs/Insumos/REMUNERACION 2026072`, When preflight Q2, Then lista vacía y el golden Q2 cierra ±0.5.
- **S4 (múltiples faltantes mixtos):** Given período con 1 carpeta ASE ausente + 1 R1 ausente en otro ASE + R10 ausente (acreditado con directorios reales de prueba —copias de estructura con nombres reales, nunca xlsx sintéticos— o con los casos que el disco ya ofrece), When preflight, Then UN error con los 3 ítems numerados en orden estable (ASE 1..5, luego período).
- **S5 (single-ASE):** Given `ProcesadorRemuneracion` con `RutaR2` inexistente, When `Ejecutar`, Then mismo código + ítem que nombra el reporte y la ruta esperada, antes de leer nada.
- **S6 (carpeta existe, archivo falta):** Given `Conciliaciones/` con 4 de 5 archivos, When preflight, Then 1 ítem que nombra la empresa faltante (nombre administrativo) + dónde debe ir el archivo.
- **S7 (lenguaje):** Given cualquier lista de faltantes, When se lee el detalle, Then sin "prefijo/matcher/finder/TopDirectoryOnly/nombres de clase"; cada ítem contiene verbo accionable (crear/colocar/pedir + reejecutar).
- **S8 (cero-geometría):** Given el diff de la HU, When se audita, Then ningún archivo bajo `Infrastructure/Excel/`, ningún mapa, ninguna fórmula y ningún mensaje existente cambió salvo las 2 llamadas de integración.

---

## 4. TASKS

| ID | Tarea | Grupo | Depende de | Verificación | Estado |
|---|---|---|---|---|---|
| T1 | Core: modelo `InsumoFaltante` + servicio de preflight (pasos §2.1 1-4, orden estable, D-D/D-F) + formateador único D-E; solo `ILocalizadorArchivosAse` como dependencia externa; sin I/O Excel, sin deps nuevas | Fix | — | Compila en `Remuneracion.Core`; S7 a nivel de formateador; build 0 warnings | COMPLETADA (WU-1) |
| T2 | Integración: llamada fail-fast en `ProcesadorPeriodo.Ejecutar` (tras L90-96, antes del loop; `Log.Warning` + `progreso?.Report`) + guardrail `File.Exists` en `ProcesadorRemuneracion.Ejecutar`; mismo código y formateador; throws runtime intactos | Fix | T1 | S1/S5; `grep "?? throw"` runtime sin diff; build 0 warnings | COMPLETADA (WU-1) |
| T-alt | *(Solo si el Ingeniero elige B)* Código nuevo del catálogo + ramas `Para`/`CodigoSalidaPara` + matriz de tests + verificación UI/CLI; sustituye el código del throw T2, mantiene servicio/mensaje/WUs | Fix-alt | T1 | S1/S7 con código nuevo; R-R-2 con la nueva cadena; manual actualizado | NO APLICABLE (ruta A) |
| T3 | Tests con insumos reales: (a) Prueba2 2026082: throw previo a `LeerR1`, 5 empresas en el detalle, código, sin salida creada (R-F-1/2/3, S1); (b) Q1/Q2 `Docs/Insumos`: lista vacía + goldens ±0.5 (R-F-4, S2/S3); (c) agregación múltiple S4 + carpeta-parcial S6 (con estructura real, sin xlsx sintéticos); (d) single-ASE S5; (e) CLI salida 2 + línea grepable (R-R-2); (f) S8 checklist + grep paridad preflight↔runtime (R-F-6) | Transversal | T2 (o T-alt) | R-F-1..R-F-6 + R-R-1/R-R-2; `dotnet build Remuneracion.WinForms/Remuneracion.WinForms.slnx` 0 warnings; suite 310/310 + nuevos sin regresión | COMPLETADA (WU-1) |
| T4 | Docs: `project-context.md` (doctrina preflight: puerta filesystem, superficie D-D con `RecaudosReversados` fuera por §V9, mensaje D-E, conciliaciones siempre §V3) + mensaje de preflight en la sección de errores del manual de usuario (`Docs/Manual-Usuario-Remuneracion-UAESP.md` §6.1; el plan nombraba `Docs/Manual-Uso-App-Administrativa.md`, ajustado por la instrucción de WU-2) | Transversal | T3 | Diff acotado a doctrina/reglas; sin código en el mismo commit; S8 citado | COMPLETADA (WU-2) |

**Orden sugerido:** T1 → T2 (o T-alt) → T3 → T4. T3 corre como regresión continua después de T1 y T2 (Q1/Q2 verdes antes y después de cada tarea). Build 0 warnings antes de marcar cada tarea completa. Sin commits (los hace el Ingeniero). **Work units encadenadas en PRs:** WU-1 = T1+T2 (servicio + integración); WU-2 = T3+T4 (tests con insumos reales + docs). PRs secuenciales, cada uno con gates DoD 1-2. T-alt, si se activa, vive dentro de WU-1 y se declara en el PR.

---

## 5. Resumen ejecutivo, riesgos y decisiones

- **Resumen:** el desperdicio es estructural y está localizado al símbolo: el procesador invierte el orden natural (procesa 5 ASE y pregunta por los insumos de período al final, L250-254) y además pregunta de a un faltante por corrida (§V1/V2). La solución aprobada (D-A..D-F) invierte el orden con una puerta filesystem que reutiliza los finders existentes (costo de divergencia = 0 por D-C), reporta TODO de una vez con UN código ya integrado a UI/CLI (costo de catálogo = 0 por D-B) y habla el idioma del usuario administrativo (D-E). Con insumos completos el flujo es bit-idéntico al actual (cero-geometría §2.3); con insumos incompletos falla en segundos sin abrir workbooks.
- **Riesgo principal:** R-COMPAT (asserts existentes sobre textos de fail-fast intermedios) + R-DERIVA (paridad futura preflight↔runtime). Contenidos por tipo+código idénticos, suite 310/310 como red ciega en cada WU, checklist de paridad y la regla D-C (mismos finders, nunca lógica paralela).
- **Decisión para el Ingeniero:** ratificar D-A..D-F (§0.3); la única con fork es D-B (recomendado: A; T-alt acota B a catálogo+tests si se elige). Aprobación del plan = aceptación de esas seis decisiones. **Corrección verificada al encargo:** `RecaudosReversados` queda FUERA de la superficie (D-D, evidencia §V9) y las conciliaciones se exigen en AMBAS quincenas (evidencia §V3). **Sin preguntas bloqueantes pendientes.**
- **Pregunta no-bloqueante con recomendación (solo si el Ingeniero quiere afinar el tono):** los `{prefijo visible}`/`{origen}` por reporte del §2.1 — recomendación: fijarlos en implementación desde `Docs/` (proceso de recaudo/detalle de plantilla) y traerlos a revisión en el PR de WU-1, sin bloquear la aprobación del plan.

### DoD (Definition of Done)

1. `dotnet build Remuneracion.WinForms/Remuneracion.WinForms.slnx` con 0 warnings.
2. Suite completa verde sin regresión (base 310/310) + goldens Capa A Q1+Q2 en dif ±0.5 intactos.
3. Test de preflight con `Docs/Prueba2/Insumos` verde: UN error con TODOS los faltantes de `Conciliaciones/` de una vez, en lenguaje claro (S1+S7); sin workbooks abiertos ni salida creada.
4. Tests con período completo Q1/Q2 verdes: CERO faltantes; Q1 no exige reportes exclusivos de Q2 (S2/S3).
5. Ninguna fórmula sobrescrita ni leída por el preflight; diff geométrico = solo servicio/modelo Core + 2 llamadas (S8); finders y mensajes existentes intactos.
6. `project-context.md` (+ fila del manual) actualizado; sin datos inventados (todo ítem remite a un finder `null` real; todo nombre administrativo a catálogo o docs).
7. Sin commits del agente (los hace el Ingeniero con `#commit`); finales de línea CRLF; sin emojis.

### Follow-ups explícitos (fuera de esta HU)

- Validación de CONTENIDO anticipada (headers/estructura antes de procesar): requiere su propio T0 (abre workbooks, bloquea archivos, duplica readers — descartado aquí por diseño, no por olvido).
- Si UAESP incorpora un insumo nuevo que el runtime resuelva, agregarlo a la lista del §2.1 con su finder (D-C); si deja de consumir uno, retirarlo (checklist R-F-6 lo detecta).

### Estado de cierre (WU-1 + WU-2, 2026-10-05)

**Ruta elegida:** A (D-B) — se reutilizó `ERR-FUENTE-NO-ENCONTRADA` con el detalle numerado en `ex.Message`; `CatalogoErrores`, `Form1` y `EjecutorCli` sin cambios y la salida 2 se conserva por construcción. **Ninguna tarea quedó pendiente** (T-alt no aplica en la ruta A).

**Hecho:**

- **T1 (WU-1):** `Remuneracion.Core/Models/InsumoFaltante.cs`, `Remuneracion.Core/Models/ReporteInsumoAse.cs`, `Remuneracion.Core/Services/ValidadorInsumosPeriodo.cs` y `Remuneracion.Core/Services/FormateadorInsumosFaltantes.cs`. Superficie D-D: R1/R2/R4 + banco + balance siempre; saldos-notas/retribución solo Q2; 5 conciliaciones + R10 siempre; `RecaudosReversados` fuera.
- **T2 (WU-1):** preflight fail-fast en `ProcesadorPeriodo.Ejecutar` (tras resolver las 5 carpetas ASE y antes del loop; `progreso?.Report` + `Log.Warning`) y guardrail `File.Exists` de R1/R2/R4 en `ProcesadorRemuneracion.Ejecutar`. Los `?? throw` runtime quedaron intactos.
- **T3 (WU-1):** `Remuneracion.IntegrationTests/PreflightInsumosTests.cs` con insumos reales: Prueba2 2026082 (todos los faltantes de `Conciliaciones/` de una vez, sin abrir workbooks), Q1/Q2 `Docs/Insumos` (cero faltantes + goldens ±0.5), single-ASE, CLI (salida 2 + línea grepable) y paridad finder↔runtime.
- **T4 (WU-2):** doctrina en `.opencode/project-context.md` (sección "Preflight de insumos del período") + mensaje de preflight en la sección de errores del manual de usuario (`Docs/Manual-Usuario-Remuneracion-UAESP.md` §6.1 y FAQ §10).

**Verificación:** `dotnet build Remuneracion.WinForms/Remuneracion.WinForms.slnx` → **0 warnings / 0 errors**; `dotnet test Remuneracion.IntegrationTests/Remuneracion.IntegrationTests.csproj` → **318/318** verde. Sin commit (lo hace el Ingeniero con `#commit`); archivos tocados en CRLF; sin emojis.

**Follow-ups vivos (fuera de esta HU):**

- Validación de CONTENIDO anticipada (headers/estructura antes de procesar): requiere su propio T0 (abre workbooks, bloquea archivos, duplica readers).
- Paridad futura preflight↔runtime: al agregar o retirar un insumo que el runtime resuelva, actualizar la lista del validador con su finder (D-C / R-F-6).
