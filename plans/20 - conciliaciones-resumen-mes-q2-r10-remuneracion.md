# Plan 20 — Conciliaciones RESUMEN MES por quincena + R10 por periodo (Q1+Q2)

> **Alcance:** G1 localizador Conciliaciones a la nueva ruta `{periodo}/Conciliaciones/`; G2 lector RESUMEN MES por quincena (Q1→D/E, Q2→F/G) + levantar el recorte T0-0.6 (Conciliacion+Recaudos en Q2); G3 R10 por periodo (localizador + reader DetRetri) con decisión de diseño DetRetri. Goldens Capa A Q1+Q2, fail-fast que nombran ASE+empresa+archivo, y actualización de `.opencode/project-context.md`.
> **Rector (IRRENUNCIABLE):** `README.md` + `Docs/Propuesta_Proyecto_Automatizacion_Remuneracion_UAESP.md` + `Docs/Detalle de plantilla.docx` + `Docs/Prompt_Maestro_Proyecto_Remuneracion_UAESP _ Vo .docx`. **Invariantes del proyecto:** NUNCA sobrescribir fórmulas (solo valores); filas dinámicas por encabezado; tolerancia ±0.5; redondear a entero antes de DetRetri/columna D; build 0 warnings; finales de línea `\r\n`; fail-fast nombrando ASE y empresa.
> **Continuidad:** HU-01..HU-19 cerradas. Este plan construye encima de: `ArchivoFuenteLocator.BuscarConciliacion`, `ExcelDataReaderWorkbookLeafInputReader.LeerRecaudoEmpresa/LeerRecaudosEmpresa`, `ProcesadorPeriodo` (recorte T0-0.6 líneas ~138-154 y 225-238), `OpenXmlPlantillaWriter.EscribirCeldasEmpresa`, `EmpresaFacturacion.Catalogo`, composición DetRetri-Q2 HU-12 (`DetRetriQ2Inputs` = ROUND(D104:D108,0)).
> **Estado:** PENDIENTE DE APROBACIÓN — no implementar hasta OK del Ingeniero.
> **Fecha:** 2026-09-22

---

## 0. Clarification Gate

Una sola pregunta con fork real para el Ingeniero (decisión G1-D1 en §0.3): **migrar** (nueva ruta únicamente, rompe `Consolidado/Conciliaciones/`) o **fallback** (nueva ruta primero, vieja como respaldo). El plan **recomienda migrar** por ser el estándar nuevo definido por el Ingeniero, y viene redactado para migrar; si el Ingeniero elige fallback, solo cambia la tarea T1 (ver alternativa acotada en §4). Todo lo demás son decisiones ejecutivas que la aprobación del plan fija. Punto de bloqueo respetado: NO se implementa nada en este documento.

### 0.1 Qué está verificado y qué NO (honestidad técnica)

**Verificado en esta planificación** (lectura directa de código):

| # | Hecho | Método | Consecuencia |
|---|---|---|---|
| V1 | `ArchivoFuenteLocator.BuscarConciliacion` resuelve `Path.Combine(carpetaPeriodo, "Consolidado", "Conciliaciones")` y devuelve `null` si no existe | Lectura `ArchivoFuenteLocator.cs:61-75` | G1: cambiar el segmento de ruta; el `null` actual permite fail-fast aguas arriba |
| V2 | `LeerRecaudoEmpresa` lee fijo columnas índice 3/4 (D/E = VALOR 1°Q) y mapea a celdas `D{3..27}/E{3..27}` (bloques OPORTUNO D3:D9, EXTEMP D12:D18, TOTAL D21:D27); comentario T0-0.6 confirma hoja única RESUMEN MES y detección de 3 bloques por corrida ASE 1..5 en col C | Lectura `ExcelDataReaderWorkbookLeafInputReader.cs:729-794` | G2: parametrizar par de columnas por quincena (Q1→3/4, Q2→5/6); la detección de bloques se reutiliza intacta |
| V3 | `ProcesadorPeriodo` omite `leaf.Conciliacion` (líneas 144-154) y `leaf.Recaudos` (líneas 228-238) en Q2 con `esQuincena2`; motivo documentado: layout R4-por-empresa Q2 diverge (ASE2 trae ENEL+OCCIDENTE, sin fila RECIPROCIDAD/"NUEVO ESQUEMA"; template Q2 sin esa fila); mapa HU-08 congelado para Q1 | Lectura `ProcesadorPeriodo.cs:138-154, 225-238` | G2: levantar el recorte exige T0 que re-verifique ese layout Q2 antes de tocar el mapa |
| V4 | Mensaje fail-fast en `LeerRecaudosEmpresa` dice literal `en Consolidado/Conciliaciones` (`ExcelDataReaderWorkbookLeafInputReader.cs:104`); XML-docs de `EmpresaFacturacion`, `ILocalizadorArchivosAse`, `IWorkbookLeafInputReader` también referencian `Consolidado/Conciliaciones` | `grep` + lecturas | G1: actualizar mensaje + XML-docs junto con el locator (grep-verificable: 0 matches de `Consolidado/Conciliaciones` al cierre) |
| V5 | `EmpresaFacturacion.Catalogo` = 5 empresas con `PrefijoConciliacion` (`Conjunta Recip`, `Conjunta ENEL`, `Conjunta ENERBIT`, `Directa`, `Conjunta Otros`) + `HojaRecaudo` por empresa | Lectura `EmpresaFacturacion.cs:54-111` | G1/G2: el catálogo NO cambia (los 5 prefijos siguen siendo las 5 claves de búsqueda en la nueva carpeta) |
| V6 | HU-12 vigente: `leaf.DetRetriQ2.TotalD104` = Σ visibles leaf Q2 + AJUSTES-SF-T (composición congelada, `ProcesadorPeriodo.cs:202-213`) | Lectura `ProcesadorPeriodo.cs:199-213` | G3: la decisión de diseño parte de esta composición como opción A (cálculo) vs opción B (pegado R10) |
| V7 | Existe `WorkbookLeafCellMapQ2` (referenciado en `OpenXmlPlantillaWriter.cs:632`) y el bloque final cambia de filas 103-108 (Q1) a 104-109 (Q2) según evidencia | `grep` + evidencia del encargo | G2/G3: el writer ya distingue Q1/Q2 en protegido; las tareas deben extender el mismo patrón a Recaudo */DetRetri |
| V8 | `plans/` 01-19 ocupados (19 = HU-19 UI premium colapsable); este plan es el **20**, primer número libre | Lectura directorio `plans/` | Numeración del documento |

**Aportado por el Ingeniero como evidencia medida (NO re-verificado en disco en esta planificación; se re-verifica en T0 antes de implementar):** organización `{periodo}/Conciliaciones/` + `{periodo}/R10_Remuneracion_AAAAMMQ.xlsx`; rutas `Docs/Insumos/REMUNERACION 2026071 y 2026072`; copia 1:1 RESUMEN MES→Recaudo * (Q1→D/E, Q2→F/G, dif 0); DetRetri = copia R10 (Q1 58.210.094.822 rango 01-15/07; Q2 72.441.209.168 rango 16-31/07); cierre CONSOLIDADO vs R10 en centavos; agregado conciliaciones +237M sobre R10 en Q1 (anulado/reversado, ANT EXT-REV); documentos fundacionales respaldan las 5 hojas Recaudo * en ambas quincenas (selector C59 1/2).

### 0.2 Mapeo al Rector (in vs out)

**Entra:**

| Requisito | Superficie de cambio |
|---|---|
| G1: Conciliaciones en `{periodo}/Conciliaciones/` | `ArchivoFuenteLocator.BuscarConciliacion` + mensaje fail-fast + XML-docs (`EmpresaFacturacion`, `ILocalizadorArchivosAse`, `IWorkbookLeafInputReader`) |
| G2: RESUMEN MES por quincena (Q1→D/E, Q2→F/G según `Periodo.NumeroQuincena`) | `LeerRecaudoEmpresa` (par de columnas por quincena) + firma `LeerRecaudosEmpresa` (recibir periodo/quincena) + `RecaudoEmpresaInputs` si el modelo ata celdas a D/E |
| G2: habilitar Conciliacion+Recaudos en Q2 | `ProcesadorPeriodo` (levantar recorte líneas ~138-154 y 225-238) + mapa R4-por-empresa Q2 (T0 previo obligatorio) + `EscribirCeldasEmpresa` columnas destino por quincena |
| G3: R10 por periodo | Nuevo `BuscarR10(carpetaPeriodo)` + reader DetRetri/DetValiRetri + decisión A/B (§0.3) + validación contra R10 |
| Transversal | Goldens Capa A Q1+Q2 (G2/G3); fail-fast ASE+empresa+archivo; actualización `.opencode/project-context.md` (rutas nuevas) |

**Sale (EXPLÍCITO):** cambios a R1/R2/R4, banco, BCE, AJUSTES-SF-T, INTERVENTORIA, L-Especiales, UI WinForms, CLI, logging, paquetes NuGet; reinterpretar el +237M Q1 (dato, no cálculo); commits (los hace el Ingeniero con `#commit`).

### 0.3 Decisiones ejecutivas (aprobación = aceptar estas)

| ID | Decisión |
|---|---|
| G1-D1 | **Migrar, no fallback (recomendado):** `BuscarConciliacion` resuelve solo `{periodo}/Conciliaciones/`. Justificación: es el estándar nuevo definido por el Ingeniero y la carpeta `Consolidado/` fue eliminada (mantener la vieja es código muerto que invita a usar la ruta equivocada). Si el Ingeniero prefiere fallback, aplicar alternativa T1-alt (§4) sin tocar el resto del plan. |
| G2-D1 | **T0 obligatorio antes de levantar el recorte:** re-verificar el layout R4-por-empresa Q2 (el caso ASE2 ENEL+OCCIDENTE que motivó T0-0.6) contra los archivos reales de `{periodo}/Conciliaciones/`. Si el layout Q2 sigue divergiendo, el levantamiento se acota a Recaudos (RESUMEN MES, layout uniforme verificado) y Conciliación-Q2 pasa a follow-up con su propio T0 — NO se reescribe el mapa HU-08 a ciegas. |
| G2-D2 | **Quincena = dominio, nunca fuente:** la selección D/E vs F/G la gobierna `Periodo.NumeroQuincena` (parámetro), jamás la detección de contenido del xlsx. Mismo principio que C59 (Requirement 4 citado en `IWorkbookLeafInputReader`). |
| G3-D1 | **DetRetri = cálculo D104:D108 + validación contra R10 (opción A).** Justificación con evidencia: el CONSOLIDADO cierra vs R10 en centavos bottom-up (R1..R5), y el agregado de conciliaciones supera a R10 en +237M Q1 por anulan/reversan (ANT EXT-REV) — prueba de que las fuentes hoja-a-hoja y el R10 NO son idénticos por construcción. Pegar R10 exacto (opción B) rompería la trazabilidad bottom-up y ocultaría divergencias reales; calcular + validar con tolerancia las expone. R10 es **oráculo de validación**, no fuente de escritura. |
| G3-D2 | **R10 fail-fast nombra periodo+archivo** (es insumo de periodo, no de ASE): `No se encontró R10_Remuneracion_{AAAAMMQ} en '{carpetaPeriodo}'`. Sin R10 no hay validación DetRetri → error, no warning silencioso. |
| G4-D1 | **Cierre grep-verificable G1:** al terminar, `grep "Consolidado/Conciliaciones"` = 0 matches en `Remuneracion.Core/` + `Remuneracion.Infrastructure/` (código, mensajes y XML-docs). |

---

## 1. PROPOSE

### 1.1 Intent

Que el motor lea las conciliaciones donde realmente están (`{periodo}/Conciliaciones/`), lea de cada RESUMEN MES la quincena que corresponde al periodo procesado (Q1→D/E, Q2→F/G), vuelva a producir Conciliación por empresa y hojas Recaudo * también en Q2, y valide el DetRetri calculado contra el R10 del periodo — con goldens Q1+Q2 que lo demuestren y errores que digan exactamente qué ASE, qué empresa y qué archivo falló.

### 1.2 In Scope

- G1: migración del localizador + mensajes + docs + `project-context.md`.
- G2: columnas por quincena en lector RESUMEN MES; levantamiento del recorte Q2 (con T0); escritura Recaudo * por quincena (verificar columnas destino Q2 F/G en writer).
- G3: localizador R10 + reader + validación DetRetri-vs-R10 con tolerancia; goldens.
- Transversal: fail-fast ASE+empresa+archivo; goldens Capa A Q1+Q2; actualización `project-context.md`.

### 1.3 Non-Goals

Ver §0.2 "Sale". En particular: NO se recalcula ni se reinterpreta el +237M Q1; NO se tocan hojas de conciliación del template más allá de Recaudo *; NO se agregan empresas al catálogo.

### 1.4 Riesgos

| ID | Riesgo | Mitigación |
|---|---|---|
| R1 | El layout R4-por-empresa Q2 sigue divergiendo (motivo original del recorte) | T0 previo (G2-D1); levantamiento acotado a Recaudos si diverge |
| R2 | Columnas destino del writer para Recaudo * Q2 (F/G) no verificadas en código | T0 incluye lectura de `EscribirCeldasEmpresa` + mapa Q2; tarea dedicada |
| R3 | Layout del R10 desconocido (nuevo insumo, sin reader previo) | T0 de layout R10 antes del reader; fail-fast si no hay hoja/celda esperada |
| R4 | Regresión Q1 al parametrizar columnas (hoy D/E fijo funciona) | Golden Q1 existente debe seguir verde antes y después (tarea de regresión explícita) |
| R5 | Bloque final desplazado Q1 (103-108) vs Q2 (104-109) | Reutilizar patrón `WorkbookLeafCellMapQ2`; golden Q2 lo cubre |

---

## 2. DESIGN

### 2.1 Enfoque por grupo

**G1 — Localizador (migrar).**
`BuscarConciliacion`: `Path.Combine(carpetaPeriodo, "Conciliaciones")`. Resto del método intacto (match por `PrefijoConciliacion`, `TopDirectoryOnly`, `null` si no existe). Actualizar: mensaje en `LeerRecaudosEmpresa` ("en Consolidado/Conciliaciones" → "en Conciliaciones/"), XML-docs de `EmpresaFacturacion` (líneas ~6-8, ~35-39), `ILocalizadorArchivosAse.BuscarConciliacion`, `IWorkbookLeafInputReader.LeerRecaudosEmpresa`. Cierre con grep G4-D1. Alternativa fallback (solo si el Ingeniero la elige): probar nueva ruta y, si `!Directory.Exists`, caer a la vieja con `Log.Warning` — tarea T1-alt sustituye a T1.

**G2 — RESUMEN MES por quincena.**
`LeerRecaudoEmpresa(empresa, ruta, quincena)`: la detección de los 3 bloques por corrida ASE 1..5 (col C) NO cambia; solo se parametriza el par de columnas leídas: Q1→índices (3,4), Q2→índices (5,6). Las claves de `Celdas` deben reflejar la columna destino real (`F/G` en Q2) o documentar que `D/E` son claves lógicas — decisión de implementación dentro de la tarea, con la restricción de que `EscribirCeldasEmpresa` + goldens Q2 cierren en dif 0 contra las salidas manuales. `LeerRecaudosEmpresa` recibe el periodo (o la quincena) y lo propaga; el `Func<EmpresaFacturacion,string?>` de `ProcesadorPeriodo` no cambia. `Periodo.NumeroQuincena` gobierna (G2-D2); fail-fast de bloques ausentes nombra empresa+archivo (ya lo hace: mantener el formato).

**G2 — Levantar recorte Q2.**
`ProcesadorPeriodo`: eliminar las dos ramas `if (!esQuincena2)` (conciliación ~138-154 y recaudos ~225-238) tras el veredicto T0. Si T0 confirma divergencia R4-Q2 persistente: se levantan solo Recaudos y Conciliación-Q2 queda como follow-up documentado (sin reescribir mapa HU-08). Writer: verificar/extender `EscribirCeldasEmpresa` para destino F/G en Q2 + bloque final 104-109 (patrón `WorkbookLeafCellMapQ2`).

**G3 — R10.**
Nuevo `BuscarR10(string carpetaPeriodo)` (prefijo `R10_`, patrón `R10_Remuneracion_AAAAMMQ.xlsx`; idealmente contrastar `AAAAMMQ` contra `Periodo.CodigoCompleto`). Nuevo reader `LeerDetRetri(periodo, rutaR10)` → `DetRetriInputs` (valores DetRetri + DetValiRetri/rango del periodo). Uso: **validación** del DetRetri calculado (HU-12) contra R10 con tolerancia ±0.5 (regla inviolable) tras redondeo a entero; divergencia → error fail-fast con periodo+archivo+ambos valores. El +237M Q1 conciliaciones-vs-R10 NO se valida (son niveles distintos: Recaudo * ≠ R10 por anulados/reversados).

### 2.2 Contratos afectados (resumen)

| Contrato | Cambio |
|---|---|
| `ILocalizadorArchivosAse.BuscarConciliacion` | Solo docs (ruta nueva); firma intacta. Nuevo `BuscarR10` |
| `IWorkbookLeafInputReader.LeerRecaudosEmpresa` | Nuevo parámetro periodo/quincena; docs (ruta nueva, columnas por quincena) |
| `RecaudoEmpresaInputs` | Revisar si ata claves a D/E; ajustar solo si impide F/G |
| `ProcesadorPeriodo` | Levantar recorte Q2; integrar validación DetRetri-vs-R10 |
| `IValidador`/`ValidadorBasico` | Nueva validación cruzada DetRetri-vs-R10 (tolerancia ±0.5) |
| `OpenXmlPlantillaWriter` | Destino Recaudo * por quincena (verificar F/G Q2); no tocar fórmulas |

### 2.3 Architecture Validation Certificate

| Principio | Veredicto |
|---|---|
| **SRP** | ✅ Cada grupo toca su costura: locator (G1), reader (G2), orquestador+validador (G2/G3). `BuscarR10` es método nuevo, no sobrecarga de `BuscarConciliacion`. |
| **OCP** | ✅ Columnas por quincena como parámetro, no `if Q2` duplicando `LeerRecaudoEmpresa`. Mapa Q2 extiende el patrón existente (`WorkbookLeafCellMapQ2`). |
| **DIP** | ✅ Cambios vía interfaces (`ILocalizadorArchivosAse`, `IWorkbookLeafInputReader`, `IValidador`); `ProcesadorPeriodo` sigue dependiendo de abstracciones. |
| **Best practices** | ✅ Quincena-dominio-gobierna (G2-D2); fail-fast con ASE+empresa+archivo; R10 como oráculo, no como escritura (G3-D1 preserva trazabilidad bottom-up). |
| **Performance** | ✅ Sin impacto: mismos archivos, una pasada; R10 es 1 archivo de periodo. Goldens reutilizan `Remuneracion.IntegrationTests` (golden Capa A existente). |

---

## 3. SPEC

### 3.1 Requirements

| ID | Requisito | Grupo | Criterio de aceptación |
|---|---|---|---|
| R-G1-1 | `BuscarConciliacion` resuelve `{periodo}/Conciliaciones/` | G1 | Los 5 prefijos del catálogo localizan sus xlsx en la nueva ruta; `grep Consolidado/Conciliaciones` = 0 en Core+Infrastructure |
| R-G1-2 | Mensajes y docs mencionan la ruta nueva | G1 | Fail-fast `LeerRecaudosEmpresa` + XML-docs actualizados |
| R-G2-1 | RESUMEN MES Q1 lee D/E | G2 | Golden Q1 Recaudo * dif 0 vs salidas manuales (regresión) |
| R-G2-2 | RESUMEN MES Q2 lee F/G según `NumeroQuincena==2` | G2 | Golden Q2 Recaudo * dif 0 vs salidas manuales |
| R-G2-3 | Conciliacion+Recaudos se producen en Q2 | G2 | Leafs Q2 con `Conciliacion`/`Recaudos` poblados; validador/writer los consumen sin rama especial |
| R-G2-4 | T0 re-verifica layout R4-por-empresa Q2 | G2 | Veredicto escrito: levantar todo o acotar a Recaudos (sin reescritura ciega del mapa HU-08) |
| R-G3-1 | `BuscarR10` localiza `{periodo}/R10_Remuneracion_AAAAMMQ.xlsx` | G3 | Fail-fast nombra periodo+archivo si falta |
| R-G3-2 | Reader R10 expone DetRetri del periodo | G3 | Q1=58.210.094.822 (01-15/07), Q2=72.441.209.168 (16-31/07) en goldens |
| R-G3-3 | DetRetri calculado se valida vs R10 (±0.5, post-redondeo) | G3 | Divergencia → error con periodo+archivo+ambos valores; R10 nunca se escribe al template |
| R-T-1 | Goldens Capa A Q1+Q2 cubren G2/G3 | Transversal | Verdes en `Remuneracion.IntegrationTests` |
| R-T-2 | `project-context.md` describe rutas nuevas | Transversal | Rutas `{periodo}/Conciliaciones/` y `R10_*.xlsx` documentadas |

### 3.2 Scenarios (Given/When/Then)

- **S1 (G1):** Given `{periodo}/Conciliaciones/` con los 5 Conjunta/Directa, When se procesa Q1 o Q2, Then cada empresa resuelve su archivo y el procesamiento continúa.
- **S2 (G1-negativo):** Given carpeta `Conciliaciones/` inexistente, When se procesa, Then fail-fast que nombra empresa+prefijo+ruta esperada.
- **S3 (G2-Q1):** Given periodo 202607-1, When se leen Recaudos, Then valores = columnas D/E del RESUMEN MES (golden dif 0).
- **S4 (G2-Q2):** Given periodo 202607-2, When se leen Recaudos, Then valores = columnas F/G del RESUMEN MES (golden dif 0).
- **S5 (G2-recorte):** Given periodo Q2 con T0 favorable, When se procesa, Then `leaf.Conciliacion` y `leaf.Recaudos` poblados y hojas Recaudo * escritas.
- **S6 (G3):** Given R10 del periodo presente, When se calcula DetRetri, Then se valida vs R10 (±0.5) y el valor escrito es el calculado bottom-up.
- **S7 (G3-negativo):** Given R10 ausente o DetRetri calculado diverge >±0.5, When se procesa, Then error fail-fast (periodo+archivo+valores), sin escritura parcial silenciosa.

---

## 4. TASKS

| ID | Tarea | Grupo | Depende de | Verificación |
|---|---|---|---|---|
| T0 | Re-verificación de evidencia en disco: (a) 5 xlsx en `{periodo}/Conciliaciones/` con hoja RESUMEN MES y VALOR 2°Q en F/G; (b) layout R4-por-empresa Q2 (¿persiste la divergencia ASE2?); (c) layout `R10_*.xlsx` (hoja/celdas DetRetri y rango); (d) columnas destino Recaudo * Q2 en `EscribirCeldasEmpresa` + mapa Q2 y bloque 104-109 | — | — | Veredicto escrito T0a..T0d; si (b) diverge → alcance acotado (G2-D1) |
| T1 | Migrar `BuscarConciliacion` a `{periodo}/Conciliaciones/` + mensaje fail-fast + XML-docs + grep de cierre (=0) | G1 | T0a | `grep Consolidado/Conciliaciones` = 0 en Core+Infrastructure |
| T1-alt | *(Solo si el Ingeniero elige fallback)* nueva ruta primero, vieja como respaldo con `Log.Warning` | G1 | T0a | Ambas rutas localizan; warning cuando usa la vieja |
| T2 | Parametrizar `LeerRecaudoEmpresa` por quincena (Q1→3/4, Q2→5/6) + propagar periodo en `LeerRecaudosEmpresa` + ajustar `RecaudoEmpresaInputs` si ata D/E | G2 | T0a | Goldens Recaudo * Q1 y Q2 dif 0 |
| T3 | Levantar recorte Q2 en `ProcesadorPeriodo` (conciliación + recaudos) según veredicto T0b | G2 | T0b, T2 | S5; sin ramas `!esQuincena2` para Recaudos (o acotado documentado) |
| T4 | Writer Recaudo * por quincena (destino F/G Q2, bloque 104-109) sin tocar fórmulas | G2 | T0d, T3 | Golden Q2 hojas Recaudo * dif 0; fórmulas intactas (assert estructural existente) |
| T5 | `BuscarR10` + reader DetRetri/R10 + validación calculado-vs-R10 (±0.5, post-redondeo, fail-fast periodo+archivo) | G3 | T0c | S6/S7; R-G3-2 con valores canónicos 58.210.094.822 / 72.441.209.168 |
| T6 | Goldens Capa A Q1+Q2 para G2/G3 en `Remuneracion.IntegrationTests` + regresión Q1 previa (verde antes y después) | Transversal | T2-T5 | `dotnet build` 0 warnings + goldens verdes |
| T7 | Actualizar `.opencode/project-context.md` (rutas `{periodo}/Conciliaciones/`, `R10_*.xlsx`, columnas por quincena, DetRetri-vs-R10 como validación) | Transversal | T1-T5 | Diff acotado a rutas/reglas; sin cambios de código en el mismo commit |

**Orden sugerido:** T0 → T1 → T2 → T3 → T4 → T5 → T6 → T7. T6 corre en cada paso (regresión continua). Build 0 warnings antes de marcar cada tarea completa. Sin commits (los hace el Ingeniero).

---

## 5. Resumen ejecutivo, riesgos y decisiones

- **Resumen:** el plan alinea el código con la organización estándar ya definida por el Ingeniero (Conciliaciones por periodo + R10 por periodo), completa la lectura por quincena que hoy está a medias (solo D/E = Q1) y cierra el recorte honesto T0-0.6 en Q2 con la re-verificación que el propio recorte exigía. La decisión de diseño central (G3-D1: calcular + validar, no pegar R10) se justifica con la evidencia más fuerte del encargo: el cierre en centavos es bottom-up y el +237M Q1 demuestra que Recaudo * y R10 viven en niveles distintos.
- **Riesgo principal:** R1 (divergencia R4-Q2 persistente) — contenido por T0 + levantamiento acotado, sin reescritura ciega.
- **Decisión para el Ingeniero:** G1-D1 migrar vs fallback (recomendado: migrar; el plan viene redactado para migrar, T1-alt es el único cambio si elige fallback).
