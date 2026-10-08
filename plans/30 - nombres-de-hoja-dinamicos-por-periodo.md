# Plan 30 — Nombres de hoja dinámicos por período (DetRetri/DetValiRetri/Informe) + base canónica agosto 2026082

> **Alcance:** dos piezas: (1) **código:** el escritor, los mapas y el oráculo resuelven `DetRetri*`/`DetValiRetri*` (y el sustituto Q1→Q2 de fragmentos "Protegidas") a partir de `Periodo.CodigoCompleto`, en vez de los literales 2026071/2026072; para agosto debe quedar `DetRetri2026082` y la app debe buscarlo dinámicamente según mes+quincena. (2) **base de agosto:** transformar `Docs/Prueba2/Plantilla_Remuneracion.xlsx` en una base canónica consistente con el manual (`DetRetri2026082`, `DetValiRetri2026082`, `Informe AFaseo Recaudo 202608-2`, referencias de fórmula reescritas, metadatos del manual). Cero cambios de fórmulas de negocio en runtime; goldens julio ±0.5 y DetRetri-D 5/5 vs R10 intactos.
> **Rector (IRRENUNCIABLE):** `README.md` + `Docs/Propuesta_Proyecto_Automatizacion_Remuneracion_UAESP.md` + `Docs/Detalle de plantilla.docx` + `Docs/Prompt_Maestro_Proyecto_Remuneracion_UAESP _ Vo .docx` + `.opencode/project-context.md`. **Invariantes del proyecto:** NUNCA sobrescribir fórmulas (solo valores); tolerancia ±0.5; ASE4 sin columna "Especiales" (0 si ausente); Q2 aplica SALDOS POR NOTA/RETRIBUCION NEGATIVA (`Periodo.NumeroQuincena == 2`, `AjustesSfT = 0` en Q1); redondear a entero antes de DetRetri/columna D; build 0 warnings; finales de línea CRLF; fail-fast nombrando archivo/celda; tests SOLO con insumos reales; sin commits (los hace el Ingeniero con `#commit`); sin emojis. **Autorización explícita y única de este plan:** la reescritura MECÁNICA del token de nombre de hoja en `<f>` corre UNA vez, offline, en la herramienta de preparación de la base de agosto (T3) — el runtime jamás toca `<f>`.
> **Continuidad:** HU-01..HU-23 cerradas; Planes 21 (espejo R1), 23 (opcionalidad 2.5), 25 (roles por firma), 26 (preflight), 27 (firma + hoja-por-nombre, suite 339/339), 28 (calcChain + sello fechas, suite 345/345), 29 (paridad app-vs-manual, suite 399/399) cerrados. Este plan NO reabre ninguna semántica de lectura, cálculo, espejo, preflight, sello ni paridad: **cero cambios de fórmulas de negocio; Q1/Q2-julio intactos por construcción (el naming dinámico debe resolver exactamente los mismos nombres que hoy para 2026071/2026072).**
> **Insumos canónicos (no se re-fijan, no se inventan datos):** `Periodo.CodigoCompleto` (`Remuneracion.Core/Models/Periodo.cs:32`); base agosto `Docs/Prueba2/Plantilla_Remuneracion.xlsx` (40 hojas, sufijo 2026072); manual agosto `Docs/Prueba2/Resultado/Resultado Manual por el administrativo/Remuneracion 202608-2 Total_7721.xlsx` (40 hojas, sufijo 2026082 — fuente de metadatos/valores permitidos); julio-base `Docs/Prueba Julio-2/Plantilla_Remuneracion.xlsx` + manual `Remuneracion 202607-2 Total Administrativo.xlsx` (2026072, no se tocan); R10 `Docs/Prueba2/Insumos/R10_Remuneracion_2026082.xlsx` (oráculo, ya dinámico).
> **Numeración:** `plans/` 01..29 ocupados; este plan toma el primer correlativo libre, **30**.
> **Estado:** CERRADO (2026-10-07) — aprobado y ejecutado T1-T5, verificado: build 0 warnings, suite 406/406, agosto E2E con hojas 2026082, julio intacto, docs actualizados, sin commits. — T0 verificado en disco (§0.1), sin implementar.
> **Fecha:** 2026-10-07

---

## 0. Clarification Gate

**Sin preguntas bloqueantes: el plan viene redactado con las decisiones D-A..D-E (§0.3) y T0 ya cerrado en disco (§0.1).** El único fork con recomendación es D-B (herramienta de preparación one-shot vs paso de la app); el plan viene redactado en la ruta recomendada (one-shot offline). Si el Ingeniero elige la otra ruta, solo cambia T3. Punto de bloqueo respetado: NO se implementa nada en este documento.

### 0.1 Qué está verificado y qué NO (honestidad técnica — evidencia T0, parser zip+XML BCL + lectura directa, sin Excel)

**Verificado en esta planificación** (disco + código, 2026-10-07):

| # | Hecho | Método | Consecuencia |
|---|---|---|---|
| V1 | `Periodo.CodigoCompleto => $"{CodigoAAAAMM}{NumeroQuincena}"` (`Periodo.cs:32`); `Periodo.Parse("2026071")` → AAAAMM+quincena; UI (`Form1.cs:83-93`) compone el período desde cmbAño/Mes/Quincena (`$"{anio}{mes}{quincena}"`) | Lectura `Periodo.cs` + `Form1.cs:83-115,199-207` | El sufijo de hoja del período ES `CodigoCompleto` por construcción (2026071, 2026072, 2026082); la fuente del naming dinámico ya existe y la UI ya selecciona mes+quincena |
| V2 | Literales de producción: `OpenXmlPlantillaWriter.cs:41-42` (`SufijoHojasQ1="2026071"`, `SufijoHojasQ2="2026072"`); `WorkbookLeafCellMapQ2.cs:37-38` (consts `HojaDetRetri/HojaDetValiRetri` 2026072); `WorkbookLeafCellMapValidaciones.cs:54-61` (`SufijoHojasDetRetri` Q1/Q2 + `HojaDetRetri/HojaDetValiRetri(quincena)`); `WorkbookLeafCellMapDetRetri.cs:37-40` (re-export const Q2); `WorkbookLeafCellMapBalanceSc.cs:80-96,146-173` (fragmentos "Protegidas" anclados a `DetRetri2026071`/`DetValiRetri2026071`); `Form1.cs:691-707` (strings de log `DetRetri2026072`, cosmético) | Grep `2026071\|2026072\|2026082\|DetRetri2026\|DetValiRetri2026\|Informe AFaseo` en Core/Infrastructure/WinForms + lectura | Inventario completo de producción (lista cerrada, §4 T1/T2). En Core solo hay comentarios/docs con esos tokens (sin runtime). `ComparadorSalidaVsManual.cs:280-286` (`NormalizarHoja` quita el token de período) y `ExcelDataReaderDetRetriR10Reader.cs:32` (`$"DetRetri{periodo.CodigoCompleto}"`) YA son dinámicos: son el patrón a seguir, no se tocan |
| V3 | Sustituto Q1→Q2 actual: `ProtegidasBceParaPeriodo(bool esQuincena2)` (`OpenXmlPlantillaWriter.cs:894-909`) aplica `ReemplazarSufijo` (`Replace(SufijoHojasQ1, SufijoHojasQ2)`) sobre Hoja+Fragmentos del mapa HU-10 (que NO se toca). El mismo `Replace("2026071","2026072")` vive en tests (`GoldenDetRetriQ2Tests.cs:293-295`, `MapaQ2T0Tests.cs:309-311`) | Lectura L894-909 + grep | El mecanismo a generalizar es "mapa anclado a Q1 + Replace del sufijo por período": basta con que el sufijo destino sea `periodo.CodigoCompleto` en vez de la const Q2 (D-A) |
| V4 | Falla actual con base consistente 2026082: `ObtenerHoja` (`OpenXmlPlantillaWriter.cs:551-557`) lanza `ERR-PLANTILLA`: `La hoja 'DetRetri2026072' no existe en el workbook para {operacion}.` Igual `ValidacionOracleReader.cs:44-49` (pide `HojaDetValiRetri(quincena)` = 2026072 en Q2) y Unidad D (`OpenXmlPlantillaWriter.cs:1132-1133` vía mapa DetRetri const) | Lectura L551-563, L1106-1133 + `ValidacionOracleReader.cs:28-49` | **Corrección al encargo:** el mensaje NO es críptico, es engañoso — nombra la hoja esperada equivocada (la hardcodeada, no la del período). El fix lo vuelve honesto (D-C): debe nombrar la hoja esperada del período + operación |
| V5 | Base agosto (40 hojas): `DetRetri2026072` + `DetValiRetri2026072` + `Informe AFaseo Recaudo 202607-2`, calcChain 12920 entradas. Manual agosto (40 hojas): `DetRetri2026082` + `DetValiRetri2026082` + `Informe AFaseo Recaudo 202608-2`, calcChain 12837. Julio-base y manual-julio: todo 2026072 (consistente, no se toca) | zip+XML `xl/workbook.xml` ambos + `Docs/Prueba Julio-2` (plantilla, R10 y manual admin) | La base de agosto es una copia con sufijo de julio; el manual prueba que el workbook real de agosto es internamente consistente en 2026082 |
| V6 | Referencias de fórmula exactas: base = **11 `<f>` con `DetRetri2026072`** (6 en `sheet9.xml` = `BCE SC POR FACT.` — H3:H7+H10-style — y 5 en `sheet37.xml` = `DetValiRetri2026072` que referencia la hoja DetRetri); manual = 11 idénticas en `DetRetri2026082`. **Cero `<f>` referencian `DetValiRetri*`** (solo aparece como nombre de hoja en `workbook.xml` + `docProps/app.xml`) y **cero `<f>` referencian `Informe AFaseo*`** (nombre de hoja solamente) | zip+XML por `xl/worksheets/sheet*.xml` + `workbook.xml` + `sharedStrings.xml` + `calcChain.xml` (0 refs Det) + `definedNames` (119/119, 0 con Det) | **Corrección al encargo (12+1):** son 11 `<f>` + declaraciones de hoja (`workbook.xml` rId36/37, sheetId 93/94 — se preservan al renombrar) + `TitlesOfParts` en `app.xml` (cosmético; la herramienta lo actualiza por higiene). Renombrar `<sheet name>` + Replace del token en esos 11 `<f>` es suficiente; gate por regex en T3 |
| V7 | Metadatos base vs manual (celdas literales): `CONSOLIDADO_TOTAL RECAUDO!N3` 2026072→**2026082**; `D6` 46238 (04/08/2026)→**46267 (02/09/2026)**; `G7` 46219 (16/07/2026)→**46250 (16/08/2026)**; `C6`=17 idéntico | zip+XML `sheet18.xml` ambos + `DateTime.FromOADate` | Valores del encargo confirmados al día; fuente = manual (no se inventa nada). `G7/K7` los sella la corrida desde el R10 (Plan 28) igualmente; la base los trae consistentes por higiene |
| V8 | `Informe AFaseo Recaudo`: **cero referencias en código** (grep en `*.cs` vacío; solo README/requirements/planes/docs: "se ignora para el consolidado") | Grep repo completo | El Informe NO entra al runtime: su rename es solo consistencia de la base (D-D), sin tarea de código |
| V9 | Tests con literales: 20 archivos (conteos: `GoldenResumenMesYR10Tests` 14, `OpcionesCliTests` 10, `Regresion2026082Tests` 9, `CodigosSalidaCliTests` 8, `GoldenDetRetriQ2Tests`/`MapaQ2T0Tests`/`Insumos` 7, `IntegridadCalcChainYSelloFechasTests`/`MapaValidacionesT0Tests` 6, `PreflightInsumosTests`/`ProcesadorPeriodoTests` 5, `DetRetriDesgloseTrazableTests` 4, `ParidadCliTests` 3, resto 1-2; suite 53 archivos). `DetRetriDesgloseTrazableTests` y `Regresion2026082Tests` ya conocen `DetRetri2026082` (R10, dinámico por V2) | `Select-String` por archivo en `Remuneracion.IntegrationTests` | La mayoría son períodos-válidos de CLI/preflight/R10 (no se tocan); el plan solo parametriza los que asertan nombres de hoja del workbook (lista cerrada en T4) |
| V10 | `Docs/Prueba2/Insumos/` hoy SÍ trae `Conciliaciones/` (el defecto del Plan 26 ya no reproduce en disco) + `R10_Remuneracion_2026082.xlsx` cuya hoja R10 es `DetRetri2026082` (leída dinámica por V2, regresión verde) | Listado + `Regresion2026082Tests` | El oráculo R10 no entra al plan; la regresión agosto es la red que protege T1/T2 |

**Aportado por el encargo (verificado arriba, no asumido):** sufijo = `CodigoCompleto` exacto; naming `DetRetri{AAAAAMMQ}` / `DetValiRetri{AAAAAMMQ}` / `Informe AFaseo Recaudo {AAAAAM}-{Q}`; conteo de fórmulas (corregido en V6: 11, no 12+1); metadatos (confirmados en V7); Q1-columnas (criterio en D-E).

### 0.2 Mapeo al Rector (Sale / No sale)

**Sale (EXPLÍCITO):**
- Reescribir `<f>` en runtime o en cualquier camino de la app (la autorización de reescritura mecánica vive SOLO en la herramienta offline T3).
- Tocar el lector R10 (`ExcelDataReaderDetRetriR10Reader`, ya dinámico V2), el comparador (`NormalizarHoja`, ya dinámico V2), el espejo R1, el preflight, el sello G7/K7, `N3` en código (es dato de la base, no del runtime).
- Cambiar semántica Q1/Q2, columnas por quincena (V3-Plan29), agregados, BCE, detalle R2/R4, Unidad D (trazables), validaciones o R10-oráculo.
- Julio-base y manual-julio (V5: ya consistentes; solo los cubre la regresión).
- Paquetes NuGet; commits (los hace el Ingeniero con `#commit`).

**No sale (entra, alcance congelado en 2 piezas):**
| Pieza | Superficie de cambio |
|---|---|
| 1 — Código dinámico | Helper Core `NombresHojaPeriodo` (D-A) + `OpenXmlPlantillaWriter` (consts→período, `ProtegidasBceParaPeriodo(periodo)`, fail-fast honesto D-C) + mapas (`Q2`, `Validaciones`, `DetRetri`, comentarios `BalanceSc`) + strings cosméticos `Form1` + tests que asertan hojas (lista T4) |
| 2 — Base agosto | Herramienta offline one-shot (BCL zip+XML, fuera del runtime) que produce la base canónica 2026082 (rename 3 hojas + 11 `<f>` + `app.xml` + metadatos del manual + columnas Q1 por D-E) + test de consistencia que la fija + regresión agosto apuntando a la nueva base |

### 0.3 Decisiones ejecutivas (aprobación = aceptar estas)

| ID | Decisión |
|---|---|
| D-A | **Helper en Core `NombresHojaPeriodo` (puro, sin I/O): `DetRetri(CodigoCompleto)`, `DetValiRetri(CodigoCompleto)`, `InformeAFaseo(CodigoAAAAMM, quincena)`.** Los mapas/writer/oráculo lo consumen; el sustituto Q1→Q2 existente pasa a `Replace(ancla-Q1, periodo.CodigoCompleto)`. OCP: el próximo período (2026091, …) funciona con cero cambios de código; Q1/Q2-julio resuelven byte-idéntico a hoy (regresión por goldens). |
| D-B | **La base de agosto se prepara con herramienta offline one-shot (recomendada), NO como paso de la app.** Preserva "plantilla canónica por período" (doctrina Plan 28/29) y el invariante runtime-nunca-toca-`<f>`. La herramienta vive fuera del pipeline (utilidad de preparación versionada o script documentado; T3 fija cuál), corre una vez, su salida se revisa contra el manual y queda fijada por test. Descartado: rename+rewrite en runtime (viola NUNCA-sobrescribir-fórmulas y convertiría cada corrida en mutación de estructura). |
| D-C | **Fail-fast honesto, no tolerancia.** Si la plantilla no trae la hoja esperada del período → `ERR-PLANTILLA` nombrando hoja esperada + período + operación (p. ej. `La hoja 'DetRetri2026082' no existe en el workbook para escritura DetRetri-Q2 (período 2026082).`). Sin fallback a hojas de otro período (un fallback silencioso escribiría el dinero del período en la hoja del período equivocado). |
| D-D | **Informe: cero código.** El runtime lo ignora por spec (V8); solo se renombra en la base por consistencia con el manual. Ninguna fórmula lo referencia (V6), así que su rename es solo `<sheet name>` (+ `app.xml`). |
| D-E | **Criterio Q1-columnas y metadatos de la base agosto: copia exacta del manual.** La app-Q2 nunca escribe columnas D/E ni `N3`/`D6` (V3-Plan29); por tanto la base las hereda del manual tal cual (proceso, no cálculo): `N3`=2026082, `D6`=02/09/2026, `G7`=16/08/2026 (+`J7`/resto de metadatos idénticos al manual celda por celda, verificado por el test T3). Nada se inventa ni se recalcula. |

---

## 1. PROPOSE

### 1.1 Intent

Que el período seleccione sus hojas por construcción (`DetRetri2026082` para agosto sin tocar código) y que agosto corra desde una base internamente consistente como la del manual — con un fail-fast que, si algo falta, nombre la hoja del período (no la de julio).

### 1.2 In Scope

- Helper `NombresHojaPeriodo` + consumo en writer/mapas/oráculo + sustituto generalizado a `CodigoCompleto`.
- Fail-fast D-C en los 3 puntos de resolución de hoja Det (escritura Q2, Unidad D, oráculo DetValiRetri).
- Herramienta offline + base canónica agosto + test de consistencia + regresión agosto sobre la nueva base.
- Docs: `project-context.md` (doctrina naming dinámico + base por período) + manual (proceso agosto).

### 1.3 Non-Goals

Ver §0.2 "Sale". En particular: NO se deriva ningún valor nuevo (la base copia metadatos/Q1 del manual; el runtime no calcula nombres, los compone del dominio); NO se valida contenido de hojas en preflight (follow-up vivo del Plan 26/27).

### 1.4 Riesgos

| ID | Riesgo | Mitigación |
|---|---|---|
| R-JULIO | El naming dinámico resuelve distinto que el literal para 2026071/2026072 y rompe goldens | D-A exige igualdad byte-exacta para esos sufijos; goldens Capa A Q1+Q2 ±0.5 + DetRetri 5/5 como red en cada tarea; T1 incluye test parametrizado 2026071/2026072/2026082 |
| R-REF | El rewrite de la base deja un `<f>`/`definedName`/`calcChain` con el token viejo → #REF! o reparación al abrir | V6 cierra el inventario (11 `<f>`, 0 definedNames, 0 calcChain); T3 verifica por regex post-rewrite (0 apariciones del token viejo en todo el zip salvo historial) + apertura sin reparación + comparador vs manual |
| R-DERIVA | Futuras hojas con sufijo de período heredan literales nuevos | Grep de cierre (T6): cero `2026xxx`/nombres-con-sufijo hardcodeados en producción fuera de `NombresHojaPeriodo` + comentarios que citen el helper |
| R-BASE-DOBLE | Dos bases de agosto divergen (la vieja 2026072 y la nueva 2026082) y alguien corre con la equivocada | La nueva base tiene nombre propio (`..._2026082.xlsx`, T3); el fail-fast D-C convierte "base equivocada" en error accionable inmediato; el manual declara cuál usar |

---

## 2. DESIGN

### 2.1 Enfoque: componer, no detectar; preparar, no mutar en runtime

**Pieza 1 (código).** Nuevo `NombresHojaPeriodo` en `Remuneracion.Core` (puro, testeable sin Excel): `DetRetri(string codigoCompleto) => $"DetRetri{codigoCompleto}"`, `DetValiRetri(...)` idem, `InformeAFaseo(string codigoAaaamm, int quincena) => $"Informe AFaseo Recaudo {codigoAaaamm}-{quincena}"`. Consumos: `WorkbookLeafCellMapQ2` (consts → `HojaDetRetri(Periodo)`/`HojaDetValiRetri(Periodo)` o el helper directo; `WorkbookLeafCellMapDetRetri` re-exporta), `WorkbookLeafCellMapValidaciones.SufijoHojasDetRetri(quincena)` → delega al helper (firma por `Periodo` donde haya dominio a mano; donde solo llegue la quincena —oráculo— se mantiene la sobrecarga pero sin literales: el helper también expone `SufijoPorQuincena` solo como puente documentado… no: mejor el oráculo recibe `Periodo` que ya tiene —`LeerSnapshots(ruta, periodo)`— y compone con `CodigoCompleto`), `OpenXmlPlantillaWriter`: consts `SufijoHojasQ1/Q2` → ancla-Q1 + `periodo.CodigoCompleto`; `ProtegidasBceParaPeriodo(Periodo)` con `Replace(ancla, periodo.CodigoCompleto)`; los 3 sitios de escritura Det reciben el nombre resuelto + fail-fast D-C en `ObtenerHoja`-llamante (el mensaje incluye hoja esperada, período y operación). `Form1` L691-707: strings al nombre resuelto (higiene de log).

**Pieza 2 (base agosto).** Herramienta offline BCL (`System.IO.Compression` + `System.Xml`, sin OpenXML/Excel — mismo método del T0): sobre copia de la base: (1) renombra los 3 `<sheet name>` en `xl/workbook.xml` (rIds/sheetIds intactos); (2) `Replace(2026072→2026082)` restringido a `<f>` que contienen `DetRetri2026072` (11, V6) — prohibido Replace global; (3) actualiza `TitlesOfParts` en `docProps/app.xml`; (4) sella metadatos + columnas Q1 desde el manual celda por celda (D-E); (5) autoverificación: 0 apariciones de `2026072` en `worksheets/`+`workbook.xml`, 11 `<f>` con `2026082`, definedNames sin token viejo, y diff vs manual en hojas/`<f>`-texto = solo valores esperados (el comparador `ComparadorSalidaVsManual` con `NormalizarHoja` ya tolera el token, V2). Salida: `Docs/Prueba2/Plantilla_Remuneracion_2026082.xlsx` (nombre propio, R-BASE-DOBLE). La base vieja se conserva como evidencia del defecto H1 (no se borra en este plan).

### 2.2 Alternativas y trade-offs

| Eje | Elegida | Descartada |
|---|---|---|
| Dónde vive el naming | Helper Core puro (testeable, OCP, sin I/O) | Resolver por enumeración de hojas del workbook ("buscar la que empiece por DetRetri") — detección por contenido, frágil ante hojas duplicadas/heredadas (justo el defecto H1: habría encontrado la de julio) |
| Cuándo se reescribe la base | One-shot offline (D-B) | Paso de la app al inicio (mutaría `<f>` en runtime; viola el rector y haría cada corrida dependiente del estado de la base) |
| Hoja ausente | Fail-fast D-C con nombre esperado | Fallback/creación de la hoja (inventaría estructura; el dinero iría a una hoja sin fórmulas encadenadas) |
| Informe en código | Nada (V8/D-D) | Resolverlo dinámicamente en runtime (código muerto: nadie lo lee ni lo escribe) |

### 2.3 Contratos afectados (resumen)

| Contrato | Cambio |
|---|---|
| Core (nuevo): `NombresHojaPeriodo` | Creación; puro, sin deps |
| `WorkbookLeafCellMapQ2` / `WorkbookLeafCellMapDetRetri` / `WorkbookLeafCellMapValidaciones` | Consts→resolución por período (literales 2026071/2026072 eliminados del runtime; el ancla-Q1 del mapa `BalanceSc` queda como dato + comentario que cita el helper) |
| `OpenXmlPlantillaWriter` | `SufijoHojasQ1/Q2`→ancla+`CodigoCompleto`; `ProtegidasBceParaPeriodo(Periodo)`; fail-fast D-C; comentarios con literales actualizados |
| `ValidacionOracleReader` | Sin cambio de firma (`periodo` ya entra): resuelve por helper (queda dinámico por T1) |
| `Form1` (logs) | Strings al nombre resuelto (cosmético) |
| `ComparadorSalidaVsManual`, R10-reader, espejo, preflight, sello, `IValidador` | Sin cambios (ya dinámicos o fuera de superficie) |
| Tests | Parametrización de asserts de hoja (lista T4) + 3 tests nuevos (helper, base-consistente, fail-fast) + regresión agosto sobre nueva base |
| `project-context.md` + manual | Doctrina naming dinámico + base canónica por período (agosto como ejemplo) |

### 2.4 Architecture Validation Certificate

| Principio | Veredicto |
|---|---|
| **SRP** | ✅ Una razón de cambio por costura: Core compone nombres, mapas declaran, writer escribe, herramienta prepara la base. Ninguna rama `if agosto`. |
| **OCP** | ✅ Futuros períodos = cero código (el nombre se compone del dominio); la base por período absorbe futuros meses por el mismo procedimiento T3. |
| **DIP** | ✅ El helper es Core puro; Infrastructure lo consume; UI/CLI no conocen la regla (ya reciben `Periodo`). |
| **Best practices** | ✅ Fail-fast con hoja+período+operación (D-C); quincena por dominio (V1); `<f>` intacto en runtime (V6/R-REF); tolerancia ±0.5 y goldens intactos por construcción; sin Excel/COM en herramienta ni tests. |
| **Performance** | ✅ O(1) por resolución (interpolación); la herramienta corre una vez offline; cero impacto en el path de corrida. |

---

## 3. SPEC

### 3.1 Requirements

| ID | Requisito | Grupo | Criterio de aceptación |
|---|---|---|---|
| R-C-1 | `NombresHojaPeriodo` compone `DetRetri/DetValiRetri/InformeAFaseo` desde el dominio; para 2026071/2026072 resuelve byte-idéntico a los literales actuales | Pieza 1 | Test parametrizado (2026071, 2026072, 2026082 + caso futuro p. ej. 2026091): `DetRetri2026071`, `DetRetri2026072`, `DetRetri2026082`, `Informe AFaseo Recaudo 202608-2` |
| R-C-2 | Writer/mapas/oráculo resuelven por período; cero literales `2026071/2026072/2026082` en runtime de producción | Pieza 1 | Grep de cierre en `Remuneracion.{Core,Infrastructure,WinForms}/**/*.cs` (salvo el helper + comentarios que lo citan + ejemplo-doc de `Periodo`); goldens Q1+Q2 ±0.5 y DetRetri 5/5 verdes |
| R-C-3 | Hoja del período ausente → `ERR-PLANTILLA` con hoja esperada + período + operación (D-C), en los 3 puntos Det | Pieza 1 | Test: plantilla sin `DetRetri2026082` (copia temp con hoja renombrada) → mensaje contiene `DetRetri2026082` + `2026082`; sin rastro del literal viejo en el mensaje |
| R-B-1 | Base `..._2026082.xlsx`: 3 hojas renombradas + 11 `<f>` con `DetRetri2026082` + 0 `2026072` en `worksheets/`+`workbook.xml` + definedNames limpios + metadatos/Q1 = manual (D-E) | Pieza 2 | Test de consistencia zip+XML (mismo método T0): conteos V5/V6/V7 + diff vs manual solo en valores de corrida |
| R-B-2 | Regresión agosto corre sobre la nueva base y cierra DetRetri 5/5 vs R10 (oráculo `DetRetri2026082`) | Pieza 2 | `Regresion2026082Tests` verde apuntando a la nueva base; salida con hojas 2026082 |
| R-R-1 | Suite base 399/399 verde antes y después + build 0 warnings; julio intacto (goldens, R10, DetRetri-D) | Transversal | `dotnet build …slnx` + `dotnet test` en cada tarea |
| R-R-2 | `project-context.md` (doctrina naming + base por período) + manual (proceso agosto: qué base usar) | Transversal | Diff acotado; sin código en el mismo commit |

### 3.2 Scenarios (Given/When/Then)

- **S1 (agosto dinámico):** Given período 2026082 + nueva base, When flujo 5-ASE Q2, Then escribe `DetRetri2026082!D9:D13` y el oráculo valida contra R10 `DetRetri2026082` 5/5; la salida trae las 3 hojas 2026082.
- **S2 (julio intacto):** Given 2026071/2026072 + bases actuales, When flujos Q1/Q2, Then nombres resueltos idénticos a hoy y goldens ±0.5 verdes (R-JULIO).
- **S3 (base equivocada):** Given período 2026082 + base vieja (2026072), When flujo Q2, Then `ERR-PLANTILLA` con `DetRetri2026082` + período + operación (SALE honesto, sin escritura parcial).
- **S4 (base consistente):** Given la nueva base, When test zip+XML, Then conteos V6 (11 `<f>` nuevos, 0 viejos) + metadatos V7 + `Informe AFaseo Recaudo 202608-2` presente y sin `<f>` que lo referencien.
- **S5 (futuro):** Given período hipotético 2026091, When test del helper, Then `DetRetri2026091`/`DetValiRetri2026091`/`Informe AFaseo Recaudo 202609-1` sin cambios de producción.
- **S6 (cero-geometría runtime):** Given el diff de Pieza 1, When se audita, Then ningún `<f>`, ningún agregado y ningún mapa vigente cambió salvo la resolución de nombres (R-C-2 grep + suite verde).

---

## 4. TASKS

| ID | Tarea | Grupo | Depende de | Verificación |
|---|---|---|---|---|
| T1 | Core `NombresHojaPeriodo` + consumo en mapas (`Q2`, `Validaciones`, `DetRetri`) y writer (ancla-Q1 + `CodigoCompleto`, `ProtegidasBceParaPeriodo(Periodo)`); `ValidacionOracleReader` por helper; higiene `Form1` L691-707; comentarios `BalanceSc`/writer actualizados citando el helper | Pieza 1 | — | R-C-1/R-C-2 (test helper parametrizado S5 incluido); build 0 warnings; suite 399/399 |
| T2 | Fail-fast D-C en los 3 puntos Det (escritura Q2 D9:D13/D14, Unidad D C..O trazable, oráculo DetValiRetri D21/D29) con hoja esperada + período + operación + test S3 (copia temp sin la hoja) | Pieza 1 | T1 | R-C-3; S3; suite 399/399 |
| T3 | Herramienta offline one-shot (BCL zip+XML, fuera del runtime/pipeline) + base `Docs/Prueba2/Plantilla_Remuneracion_2026082.xlsx` (rename 3 hojas + 11 `<f>` + `app.xml` + metadatos/Q1 del manual D-E) + test de consistencia R-B-1 (conteos V5/V6/V7) | Pieza 2 | — (paralela a T1/T2) | R-B-1; S4; apertura sin reparación; comparador vs manual solo con divergencias de corrida |
| T4 | Tests: parametrizar asserts de hoja con literales (`DetRetriCapacidadTests`, `GoldenDetRetriQ2Tests:293-295`, `MapaQ2T0Tests:109-126,298-311`, `MapaValidacionesT0Tests:162-166`, `GoldenResumenMesYR10Tests`, `IntegridadCalcChainYSelloFechasTests`, `DetRetriDesgloseTrazableTests`, `Insumos` donde fije hoja) + regresión agosto sobre la nueva base (R-B-2) | Transversal | T1, T3 | R-B-2/R-R-1; S1/S2; goldens ±0.5; DetRetri 5/5 ambos períodos |
| T5 | Docs: `project-context.md` (doctrina naming dinámico D-A + base canónica por período D-B/D-E + fail-fast D-C) + manual (qué base usar por mes, S3 como guía de error) | Transversal | T1..T4 | R-R-2; diff acotado |
| T6 | Cierre: grep cero-literales (R-C-2) + build 0 warnings + suite completa + auditoría S6 (cero `<f>` runtime; T3 como única reescritura, offline y versionada como utilidad, no como paso) | Transversal | T1..T5 | R-R-1; S6; CRLF; sin emojis; sin commits |

**Orden sugerido:** T1 → T2 → T4 (código) en paralelo con T3 (base) → T5 → T6. T3 no bloquea T1/T2 (usa el método T0 ya verificado). Build 0 warnings antes de marcar cada tarea completa. Sin commits (los hace el Ingeniero).

---

## 5. Resumen ejecutivo, riesgos y decisiones

- **Resumen:** el período ya existe como dominio (`CodigoCompleto`) y la UI ya elige mes+quincena (V1); solo los nombres de hoja quedaron congelados en julio. El T0 corrige dos supuestos del encargo: las fórmulas a reescribir son **11 (no 12+1)** —`DetValiRetri` e `Informe` solo aparecen como nombres de hoja (V6)— y el fallo actual no es críptico sino **engañoso** (nombra la hoja de julio, V4). El plan compone los nombres desde el dominio (Pieza 1, OCP hacia futuros períodos) y prepara la base de agosto offline una sola vez copiando metadatos/Q1 del manual (Pieza 2), con el runtime sin tocar jamás una fórmula.
- **Riesgo principal:** R-JULIO (resolución distinta para 2026071/2026072) + R-REF (#REF! por rewrite incompleto). Contenidos por test parametrizado del helper, goldens ±0.5 en cada tarea y verificación regex post-rewrite (0 token viejo).
- **Decisión para el Ingeniero:** ratificar D-A..D-E (§0.3). **Fork con recomendación (D-B):** herramienta offline one-shot (redactada) vs paso de la app — recomendado one-shot (preserva plantilla-canónica-por-período y el invariante `<f>`-intacto-en-runtime); si se elige paso-de-app, solo cambia T3 y debe justificarse contra el rector. **Sin preguntas bloqueantes pendientes.**

### DoD (Definition of Done)

1. `dotnet build Remuneracion.WinForms/Remuneracion.WinForms.slnx` con 0 warnings.
2. Suite completa verde sin regresión (base 399/399) + goldens Capa A Q1+Q2 ±0.5 + DetRetri 5/5 vs R10 (julio y agosto).
3. Agosto 2026082 end-to-end sobre la nueva base con hojas 2026082 (S1); julio byte-idéntico en nombres (S2).
4. Base equivocada → `ERR-PLANTILLA` con hoja del período (S3); base nueva con conteos V6/V7 (S4).
5. Grep de cierre: cero literales de período en runtime fuera del helper (S6); `project-context.md` + manual actualizados.
6. Sin commits del agente; finales de línea CRLF; sin emojis.

### Follow-ups explícitos (fuera de este plan)

- Validación de nombres de hoja en preflight (abriría workbooks; follow-up vivo Plan 26/27).
- Si un futuro período trae hojas con otra forma (no solo otro sufijo), su propio T0 decide (este plan solo compone el sufijo).
- Retiro de la base vieja de agosto (conservarla como evidencia H1 hasta que el Ingeniero la archive).
