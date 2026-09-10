# Plan 17 — HU-17: Manual de usuario + entrega a pruebas (Fase 3, it. 3.4)

> **Historia:** cerrar el proyecto con la documentación que convierte el 94 % técnico en entrega: manual de usuario, instructivo Capa B protocolizado, paquete de entrega con casos y criterio de pase, más el saneamiento final de la deuda documental y micro-fixes acotados. Última HU del proyecto (17/17).
> **Rector (IRRENUNCIABLE):** `Docs/Propuesta_Proyecto_Automatizacion_Remuneracion_UAESP.md` — §9 Fase 3 it. **3.4** ("Documentación | Manual de usuario y technical"), §6 (trazabilidad: "Cada paso se registra"), §10 CA-4/CA-6/§10.2 (fórmulas intactas, log por paso, éxito = CONSOLIDADO ±0.5 vs resultado manual). **Entra SOLO 3.4 + deuda listada en §0.1.**
> **Origen funcional:** `Detalle de plantilla.docx` (pasos 1–10 de ejecución, nomenclatura `Remuneración AAAAMM-# Total.xlsx`), `Prompt Maestro Vo.docx` (pegado en valores, AJUSTES-SF-T D85:D89, C59, E59:E80). `Proceso de Recaudo.docx` **NO aplica** (flujo bancario, cubierto en HU-09).
> **Continuidad:** HU-01..HU-16 cerradas (build 0/0, tests 264/264, harness 24/24; UI + CLI funcionales; Q1/Q2 certificables en Capa A). Este plan NO reabre su semántica: **cero cambios de cálculo/escritura/validación; Q1/Q2 intactos por construcción; sin golden nuevo.**
> **Estado:** PENDIENTE DE APROBACIÓN — no implementar hasta OK del Ingeniero.
> **Fecha:** 2026-09-09

---

## 0. Clarification Gate

No hay pregunta bloqueante para el Ingeniero. Las decisiones de formato y disposición de deuda se dictaminan ejecutivamente en §0.3; la aprobación del plan las fija. No hay T0 de workbook en esta HU: ningún micro-fix exige semántica de celda no congelada (los valores D2b/D3a/D3b y la doctrina A1/A2 se reutilizan, no se re-descubren).

### 0.1 Qué está verificado y qué NO (honestidad técnica)

**Verificado en esta planificación** (código, disco y hashes propios del orquestador):

| # | Hecho | Método | Consecuencia |
|---|---|---|---|
| V1 | Propuesta §9 it. 3.4 = "Manual de usuario y technical"; §10.2 = éxito si CONSOLIDADO ±0.5 vs resultado manual; CA-4 fórmulas intactas; CA-6 log por paso | Lectura Propuesta:474-502 | Rector normativo de entrega; el criterio de pase de §3 lo cita literal |
| V2 | SHA256 fijados hoy (acta W-1 HU-16): Q1 `Docs/Insumos/Remuneracion 202607-1 Total.xlsx` = `0B090E9C…C096F`; Q2 golden `Docs/Insumos/Remuneracion 202607-2 Total.xlsx` = `58431010…21DCB1`; Q2 canónico `REMUNERACION 2026072/Plantilla 8 agos 2026 _ …` = `95825422…E0B8C` (coincide con HU-11); Q2 control `Plantilla  _ …` = `509BF210…57529B` | `Get-FileHash` propio | Ancla de integridad del instructivo Capa B: si un SHA difiere, se detiene y se escala (los insumos son untracked en git) |
| V3 | CLI real: flags `--periodo/--quincena/--carpeta/--plantilla/--salida/--ase N/--cinco-ase/--sobrescribir/--help/-h` (`OpcionesCli.cs:85-132`); códigos 0 OK / 1 validación / 2 fuente-plantilla / 3 escritura / 4 inesperado / 5 cancelado (contrato HU-14, corridas vivas HU-15) | Lectura código + reportes HU-15 | El manual documenta exactamente estos flags y códigos, sin inventar ninguno |
| V4 | Deuda pendiente real: W-1 HU-16 (acta SHA + checkboxes §4 HU-16); S-1 HU-16 (corte `row>700` en `InterventoriaTests.cs:183`/`GoldenInterventoriaTests.cs:178`); S-2 (~90 líneas helpers OpenXML duplicadas en ambos test files); S-3 (`LeerCeldaNumerica` texto→0 en `OpenXmlPlantillaWriter.cs:741-752`); S-4 (triple log D2b); W-2.2 HU-14 (re-parseo regex + `InvariantCulture` en `Form1.cs:489-540`); S-4 HU-14 (`ParseAse` lanza tipo archivo-inexistente en `Form1.cs:573`); S-5 (25 lecturas W-1 misma hoja); S-1 newlines (5 archivos sin newline final + 10 creados sin él); W-3 HU-15 (`.gitignore` cubre `remuneracion_log.txt` exacto, el rolling `remuneracion_log_*.txt` queda untracked); S-1 HU-15 (`Program.cs:12-17` Main sin catch externo); S-2 HU-15 (sin dispose tras reconfigurar Serilog en `EjecutorCli.cs:57`); S-3 (`temp/head_program.cs` tracked, único archivo bajo `temp/`) | Lectura revisiones HU-14/15/16 + `git ls-files temp/` + `.gitignore` | Tabla de disposición §2.5: cada ítem entra con micro-fix + test o con nota documentada; nada queda flotando |
| V5 | Plan 12 dice "sheets 93/94" pero la verdad canónica es **36/37** (revisión HU-12 con `workbook.xml`; el código sigue la verdad) | Revisión HU-12 | Fe de erratas en §9 del Plan 12 (edición del plan, como se hizo en Plan 15 §9) |
| V6 | M2 HU-12: `TieneColumnaEspeciales = false` fijado por código (evidencia T0-0.5: 5/5 sin columna); L1 HU-12: fuente vacía → bloque no se limpia (inocuo hoy: canónico Q2 en blanco) | Revisión HU-12 | M2 = nota de aceptación en Plan 12 §9 + manual; L1 = nota en instructivo Capa B + pinning test del comportamiento |
| V7 | `Docs/` hoy: Propuesta + 3 base + 3 `.html` de exposición (Comparativo/Exposicion/Flujo, aporte del Ingeniero); **ningún manual** | Glob `Docs/` | El manual y el instructivo son documentos nuevos, no ediciones |
| V8 | Harness `Herramientas/VerificadorRecaudo` (fuera del `.slnx`) emite 6 warnings CS8602/8603/8604 preexistentes; suite 24/24 verde | Corrida propia HU-16 | Se sanea mecánicamente con red 24/24, o se registra si rompe (rollback: harness intacto) |

**NO verificado (y por eso NO entra):** instalador/setup (no pedido; §2.6 dictamina `dotnet publish` framework-dependent + nota); motor Excel/COM en CI (prohibido por doctrina); validez de los `.html` de exposición (aporte del usuario, no se tocan).

> **Regla de hierro del plan:** ningún micro-fix cambia valores, gates, mapas, goldens ni tolerancia ±0.5. Lo que rompa la red 264/264 + 24/24 = NEEDS_CONTEXT con rollback del micro-fix, nunca maquillaje de tests.

### 0.2 Mapeo al Rector (in vs out)

**Entra porque §9 it. 3.4 / §6 / §10 lo piden ahora:**

| Rector | Qué cubre HU-17 |
|---|---|
| §9 it. 3.4 | Manual de usuario (`Docs/Manual-Usuario-Remuneracion-UAESP.md`) + nota técnica (ADR Serilog + geometría Q2 + D2b, §2.2) |
| §10 CA-4/CA-6/§10.2 | Instructivo Capa B (`Docs/Instructivo-Capa-B.md`): protocolo post-Excel con criterio de pase literal del Rector |
| §9 Fase 3 (cierre) | Paquete de entrega (`Docs/Paquete-Entrega-Pruebas.md`): contenido + casos CT-Q1-5A/CT-Q1-1A/CT-Q2-5A/CT-NEG + criterio de pase |
| Deuda §0.1 V4 | Tabla §2.5: micro-fixes + notas + `.gitattributes` + gitignore + `temp/head_program.cs` + erratas de planes |

**Sale (EXPLÍCITO):** cambios de cálculo/escritura/validación; INTERVENTORIA lógica (cerrada HU-16); CLI flags/códigos nuevos; instalador/setup; DI framework; restyle UI; tocar goldens/canónicos (solo lectura + hash); commits (los hace el Ingeniero con `#commit`).

### 0.3 Decisiones ejecutivas (aprobación = aceptar estas)

| ID | Decisión |
|---|---|
| G1 | **Tres documentos nuevos en `Docs/`**: manual + instructivo Capa B + paquete de entrega. Nada en wiki externa ni en `.html` existentes. |
| G2 | **Capa B como procedimiento ejecutable**, no como declaración: pasos numerados por quincena con celdas exactas, tolerancia y registro de evidencia; el Ingeniero la ejecuta, el instructivo la guía. |
| G3 | **SHAs V2 como ancla de integridad** en el instructivo: ante divergencia, detener y escalar (no "ajustar" el hash). |
| G4 | **Micro-fix = código + test o nota**: cada ítem §2.5 sale con test dedicado o con nota de aceptación documentada; prohibido el cambio silencioso. |
| G5 | **`.gitattributes` repo-wide**: `*.cs/csproj/slnx/md → text eol=crlf`; `*.xlsx/docx/pdf → binary`. Cierra S-1 newlines + regla #8 a nivel repo. |
| G6 | **`temp/head_program.cs` sale del repo** (`git rm` en working tree; el commit lo hace el Ingeniero). `temp/` queda sin tracked. |
| G7 | **Harness con rollback garantizado**: se sanea solo si 24/24 sigue verde; si un fix lo rompe, se revierte el archivo y se registra. |
| G8 | **Q1/Q2 intactos por construcción**; red 264/264 + 24/24 en cada unidad; build 0 warnings incl. CLI; CRLF; sin commits. |

---

## 1. PROPOSE

### 1.1 Intent

Entregar al Ingeniero el paquete cerrado de pruebas: qué ejecutar (UI/CLI), qué esperar (goldens ±0.5, fórmulas intactas, códigos 0–5), cómo certificarlo (Capa B paso a paso con evidencia) y con qué criterio se pasa — más el saneamiento final que deja el repo sin deuda flotante.

### 1.2 In Scope

- `Docs/Manual-Usuario-Remuneracion-UAESP.md` (instalación, UI, CLI, códigos, catálogo, log/RunId, fallos, INTERVENTORIA-insumo-externo, preguntas frecuentes).
- `Docs/Instructivo-Capa-B.md` (protocolo post-Excel por quincena + acta SHA + registro de evidencia + notas L1/M2).
- `Docs/Paquete-Entrega-Pruebas.md` (contenido del paquete + casos CT + criterio de pase §10.2).
- Micro-fixes y notas de la tabla §2.5 + `.gitattributes` + gitignore + salida de `temp/head_program.cs` + erratas en planes 12 (§9) y 15 (si aplica).
- Saneamiento mecánico del harness con red 24/24.

### 1.3 Out of Scope

Todo §0.2 (cálculo/escritura/validación, INTERVENTORIA lógica, flags/códigos nuevos, setup, DI, restyle, goldens, commits).

### 1.4 Resultado de negocio

El Ingeniero ejecuta el instructivo Capa B con el CLI, llena el registro de evidencia y declara pase/falla por caso con el criterio del Rector; el repo queda con 17 planes, 0 deuda flotante y convenciones fijadas a nivel repo.

### 1.5 Base documental (origen funcional — citas por sección)

| Documento | Sección / instrucción | Qué aporta a 3.4 |
|---|---|---|
| `Detalle de plantilla.docx` | Pasos 1–10 (identificar período, verificar reportes, diligenciar hojas, conservar fórmulas, insertar/eliminar filas, CONSOLIDADO, DetRetri enteros, validaciones, ignorar AFaseo, entregar + reportar) | Esqueleto del manual (procedimiento operativo que la app automatiza) |
| `Prompt Maestro Vo.docx` | Nomenclatura `Remuneración AAAAMM-# Total.xlsx`; pegado en valores; AJUSTES-SF-T D85:D89; C59/E59:E80 | Convenciones que el manual fija para el operador |
| Rector Propuesta | §9 it. 3.4 (alcance), §6 (trazabilidad), §10 CA + §10.2 (criterio de pase) | Rector normativo |
| `Proceso de Recaudo.docx` | — | **NO aplica** (flujo bancario, HU-09) |

---

## 2. DESIGN

### 2.1 Estructura del manual (`Docs/Manual-Usuario-Remuneracion-UAESP.md`)

1. Qué hace la app y qué NO hace (automatiza pegado en valores; Excel recalcula; INTERVENTORIA es insumo externo anual).
2. Requisitos (`dotnet` 10, Excel para Capa B) + `dotnet publish` framework-dependent (nota §2.6).
3. Modo UI paso a paso (período, carpeta `REMUNERACION AAAAMM{Q}`, plantilla canónica, salida, modo 1-ASE/5-ASE, quincena, sobrescribir).
4. Modo CLI: flags exactos V3 + ejemplos Q1-5A / Q1-1A / Q2-5A + línea `RESULTADO OK/ERROR` grepable.
5. Códigos de salida 0–5 con causa y acción (tabla del catálogo HU-14; salida 4 = uso, 5 = cancelado/sin sobrescribir).
6. Catálogo de errores (código → título → guía) + RunId: cómo filtrar una ejecución en el log.
7. El log (`remuneracion_log_*.txt`, rolling 30 días, niveles, propiedades `Hoja/AseId/Validacion/RunId`).
8. Casos que NO son error (E59:E80 ≠ 0 esperado; D21/D9:D14/J9:J14 excluidos; retribuciones vacías → 0 legítimo; Q2-1ASE fail-fastea por diseño).
9. INTERVENTORIA: origen del costo (insumo externo anual, cómo diligenciarlo), assert estructural, guía ante divergencia.
10. FAQ + "si el SHA difiere, detente y escala".

### 2.2 Nota técnica (capítulo del manual, it. 3.4 "technical")

ADR Serilog-en-Core (alcance confinado a orquestadores); geometría Q2 ≠ Q1 (tabla de visibles F53/F206/… por período); D2b INTERVENTORIA; D3a/D3b L-Especiales; veredicto D/E del BCE; golden honesto Capa A (OpenXML no recalcula).

### 2.3 Instructivo Capa B (`Docs/Instructivo-Capa-B.md`)

Acta SHA (V2) → generar salida a ruta distinta del golden → abrir en Excel → recalcular → comparar por bloque vs golden ±0.5:
- Q1: CONSOLIDADO D9:D109 cadena completa; DetRetri = ROUND (39148067547 ASE1); REMUNERACION_* bloques; BCE D/E; VALIDACION O=0/P=TRUE; INTERVENTORIA 26..30 + K31/K32 fórmulas.
- Q2: D85:D89 AJUSTES (973693.46 / 216025.77 / 104231.83 / 35954.44 / 0); DetRetri-D enteros (total 72441209168); D104:D109; INTERVENTORIA idéntica a Q1 (tabla anual).
- Fórmulas intactas (spot-check de protegidas) + registro de evidencia (tabla fecha/quincena/caso/resultado/SHA salida) + notas L1 (bloque vacío vs fuente vacía: verificar, no asumir) y M2 (Especiales ausente en fuentes saldos-nota).
- Criterio: todo cierra ±0.5 + fórmulas intactas = PASA; cualquier divergencia = FALLA con celda y valores (nunca se "ajusta" el golden).

### 2.4 Paquete de entrega (`Docs/Paquete-Entrega-Pruebas.md`)

Contenido (binarios publicados UI+CLI, plantillas canónicas, goldens como referencia solo-lectura, los 3 docs, registro de evidencia en blanco) + casos:
- **CT-Q1-5A**: CLI 5-ASE Q1 → exit 0; Capa B vs golden Q1.
- **CT-Q1-1A**: CLI `--ase 3` Q1 → exit 0; bloque ASE3 vs golden.
- **CT-Q2-5A**: CLI 5-ASE Q2 → exit 0; AJUSTES + DetRetri + D104:D109 vs golden Q2.
- **CT-NEG**: carpeta inexistente → 2; salida==plantilla → 2; sin `--sobrescribir` → 5; `--periodo 2026` → 4.
- Pase = §10.2 literal + CA-4 + CA-6.

### 2.5 Disposición de deuda (cada ítem sale con test o nota)

| # | Deuda | Disposición | Test / evidencia |
|---|---|---|---|
| W-1 HU-16 | Acta SHA + checkboxes | Doc: SHAs V2 en instructivo + §4 HU-16 | Revisión del doc |
| S-1 HU-16 | Corte `row>700` | Code: derivar de `Dimension` | Test con fila fantasma >700 |
| S-2 HU-16 | Helpers OpenXML duplicados | Code: `TestHelpers` compartido (solo tests) | Suite verde sin duplicado |
| S-3 HU-16 | Texto→0 miente | Code: fail-fast "no numérico" nombrando celda | Test texto en L=Id |
| S-4 HU-16 | Triple log D2b | Code: lectura a `Debug` | Suite Observabilidad verde |
| W-2.2 HU-14 | Re-parseo cultura Form1 | Code: `CurrentCulture` o eliminar re-parseo (usa evento estructurado) | Test con cultura `es-CO` vs invariante |
| S-4 HU-14 | `ParseAse` tipo erróneo | Code: `ArgumentException` (1 línea) | Test tipo de excepción |
| S-5 | 25 lecturas misma hoja | Medir una vez; código solo si supera 1 s por hoja | Nota + medida en el plan de cierre |
| S-1 nl | 15 archivos sin newline final | Normalizar + convención §9 | Scan de bytes en la tarea |
| Attr | `.gitattributes` repo-wide | Crear (`*.cs/csproj/slnx/md → eol=crlf`; `*.xlsx/docx/pdf → binary`) | Build + scan |
| W-3 HU-15 | Rolling sin gitignore | `remuneracion_log_*.txt` al `.gitignore` | `git status` limpio tras corrida |
| S-1 HU-15 | Main sin catch externo | try/catch → 4 con flush | Test arg inválido extremo |
| S-2 HU-15 | Sin dispose Serilog | Guard dispose antes de reconfigurar | Test de doble configuración |
| S-3 | `temp/head_program.cs` tracked | `git rm` (working tree; commit del Ingeniero) | `git ls-files temp/` vacío |
| Plan12 | "sheets 93/94" → 36/37 + M2 nota | Editar §9 Plan 12 | Diff del plan |
| L1 HU-12 | Bloque vacío no se limpia | Nota Capa B + pinning test del comportamiento actual | Test + instructivo |
| HU-11 | S-1/S-2/W-2 opcionales | Barrido: cerrar o registrar con evidencia | Nota de cierre |
| Harness | 6 warnings CS860x | Saneamiento mecánico + 24/24; si rompe, revert + registro | Harness verde |

> **Estado de ejecución HU-17 (U2, verificado):**
> - **W-1 HU-16** ✅ doc: SHAs V2 en `Docs/Instructivo-Capa-B.md` §2 + Manual §12 + acta en Plan 16 §4 (checkboxes marcados).
> - **S-1 HU-16** ✅ `TestHelpers.LimiteFilaDesdeDimension` (deriva de `SheetDimension`); test `S1_Hu16_EnumerarCeldasL_DerivaLimiteDeDimension_FilaFantasma` (L730 dentro de A1:N750 se lee; L780 fuera se ignora).
> - **S-2 HU-16** ✅ `TestHelpers` compartido; `InterventoriaTests`/`GoldenInterventoriaTests` sin duplicado; suite verde.
> - **S-3 HU-16** ✅ `OpenXmlPlantillaWriter.LeerCeldaNumerica` fail-fast "NO numérico" nombrando `hoja!celda`; test `S3_Hu16_Writer_TextoEnLCeldaId_FailFastNombraCelda`.
> - **S-4 HU-16** ✅ lectura K/M/N + L-Especiales por ASE → `Debug` en `ProcesadorPeriodo` y `Form1` (hitos siguen `Information`); test `S4_Hu16_InterventoriaLecturaPorAse_Debug_HitoInformation`.
> - **W-2.2 HU-14** ✅ se eliminó el re-parseo con `InvariantCulture` en `Form1` (log del texto capturado); test `W22_Hu14_Validaciones_CulturaEsCo_LineaYEventoCoherentes`.
> - **S-4 HU-14** ✅ `Form1.ParseAse` lanza `ArgumentException`; test `S4_Hu14_SeleccionAseInvalida_TipoArgumentException` (familia).
> - **S-5** ✅ medido con proyecto .NET real (OpenXML 3.5.1): 25 lecturas de `INTERVENTORIA` (golden Q1) = **1028.1 ms total → 41.12 ms/hoja** — muy por debajo del umbral de 1 s/hoja → **nota, sin código**.
> - **S-1 nl** ✅ 79 archivos normalizados a CRLF + newline final (scan: 0 sin newline final); `.gitattributes` fija la convención.
> - **Attr** ✅ `.gitattributes` repo-wide creado (`*.cs/csproj/slnx/md/jsonc/json → text eol=crlf`; `*.xlsx/docx/pdf → binary`).
> - **W-3 HU-15** ✅ `remuneracion_log_*.txt` agregado a `.gitignore`.
> - **S-1 HU-15** ✅ `Main` con try/catch externo → código 4 + flush; test subproceso `S1_Hu15_MainCatchExterno_ArgInvalidoExtremo_Exit4_SinCrash`.
> - **S-2 HU-15** ✅ `ConfigurarSerilog` con `Log.CloseAndFlush()` previo (dispose del sink); test `S2_Hu15_ConfigurarSerilog_DobleConfiguracion_NoLanzaYSigueEscribiendo`.
> - **S-3** ✅ `git rm temp/head_program.cs` (working tree); `git ls-files temp/` vacío.
> - **Plan12** ✅ fe de erratas "sheets 93/94" → 36/37 + nota M2 en §9.
> - **L1 HU-12** ✅ nota en `Docs/Instructivo-Capa-B.md` §8 + pinning test `L1_Hu12_BloqueVacio_NoSeLimpia_OutputIgualTemplate`.
> - **HU-11** ✅ **nota de cierre:** los ítems S-1/S-2/W-2 opcionales de la revisión HU-11 no quedaron persistidos en texto (revisión efímera). El barrido con el código actual no identifica ítems accionables pendientes: los patrones asociados (log filtrable por hoja/validación, mensajes no redundantes, cultura de parseo) están cubiertos por HU-14 (RunId + propiedad `Validacion` + catálogo), HU-15 (W-2.1..W-2.4) y HU-17 (W-2.2 re-parseo eliminado). Sin deuda flotante.
> - **Harness** ✅ 6 warnings CS860x saneados mecánicamente (null-forgiveness, cero cambio de comportamiento) con **24/24 verde + 0 warnings** (rollback no necesario).

### 2.6 Publicación (dictamen: sin instalador)

`dotnet publish` framework-dependent para UI + CLI + nota de despliegue en el manual. Instalador/setup queda explícitamente fuera (no pedido; se re-evalúa solo si pruebas lo exigen).

---

## 3. SPEC

| Req | Requisito | Aceptación | Trazabilidad |
|---|---|---|---|
| R1 | Manual con capítulos §2.1 + nota §2.2 | Revisión del Ingeniero: cada flag/código del manual existe en código (grep); cero instrucciones inventadas | V3, Rector §9-3.4 |
| R2 | Instructivo Capa B ejecutable | El Ingeniero puede certificar CT-Q1-5A solo con el instructivo; acta SHA presente; registro en blanco | §10 CA-4/§10.2, G2/G3 |
| R3 | Paquete + casos + pase | Casos CT ejecutables con comandos literales; pase = §10.2 + CA-4 + CA-6 | §2.4 |
| R4 | Deuda §2.5 saldada | Cada fila con test verde o nota documentada; `git ls-files temp/` vacío; `.gitattributes` efectivo | G4/G5/G6 |
| R5 | Regresión intacta | 264/264 + 24/24 verdes; slnx 0/0; CRLF; SHAs V2 intactos tras todo | G8 |

---

## 4. TASKS (PRs encadenados)

- [x] **U0 — Baseline con gate**: build 0/0, 264/264, 24/24, SHAs V2 re-verificados. Sin verde no se empieza.
- [x] **U1 — Documentos**: manual + instructivo + paquete (§2.1–§2.4) + erratas en planes 12 (§9) y acta W-1. Revisión propia contra código (R1).
- [x] **U2 — Micro-fixes + repo**: tabla §2.5 en orden (tests primero que código donde aplique); `.gitattributes`; gitignore; `git rm temp/head_program.cs`; normalización newlines; harness con rollback.
- [x] **U3 — Cierre**: regresión completa + scan CRLF + SHAs intactos + `git status` explicado línea por línea (qué es producto HU-17, qué es base sin commit previa). Sin commits. Reporte final con la lista §2.5 tildada.

---

## 5. Verificación y Capa B

- **Capa A (CI)**: R5 en U0 y U3; los micro-fixes agregan sus tests a la suite (264 + N).
- **Capa B (manual, acción del Ingeniero)**: el instructivo §2.3 es el protocolo; esta HU lo escribe, no lo ejecuta. El primer CT-Q1-5A ejecutado por el Ingeniero con el instructivo es la prueba de fuego de R2 (si el instructivo no alcanza, se enmienda en esta misma HU antes del cierre).

## 8. Rollback

- Revertir U2→U1 en orden inverso; los docs son aditivos (borrar 3 archivos + revertir §9 Plan 12 deja el árbol como estaba).
- Micro-fix que rompa red: revert del archivo + registro en §2.5 (nunca maquillaje).
- `temp/head_program.cs`: se restaura con `git checkout -- temp/` si el Ingeniero lo pide.

## 9. Decisiones fijadas por este plan

1. Rector = Propuesta §9 it. 3.4 + §6 + §10. Entran manual + Capa B + paquete + tabla §2.5; cálculo/escritura/INTERVENTORIA-lógica/flags/setup/commits quedan fuera.
2. SHAs V2 como ancla de integridad (tabla §0.1); ante divergencia se detiene y se escala.
3. Tres docs nuevos en `Docs/`; `.html` de exposición intactos.
4. Convención repo-wide: CRLF + newline final (`§2.5` + `.gitattributes`); binarios Excel/Word como `binary`.
5. `temp/` sin tracked; rolling logs ignorados.
6. `dotnet publish` framework-dependent; sin instalador.
7. PRs U0→U3; red 264/264 + 24/24; slnx 0 warnings; CRLF; sin commits.
8. **Fe de erratas de cierre (2026-09-09, orquestador):** la red citada es el baseline U0 (264/264); los 9 tests de micro-fixes de esta misma HU la llevaron a **273/273** (paquete corregido en consecuencia). W-1/W-2 del review cerrados como micro-fix post-revisión.

**Listo para aprobación del Ingeniero. No implementar hasta OK.**
