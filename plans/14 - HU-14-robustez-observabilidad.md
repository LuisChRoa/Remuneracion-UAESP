# Plan 14 — HU-14: Robustez y observabilidad (Fase 3, it. 3.1 + 3.2)

> **Historia:** cerrar formalmente el manejo de errores (3.1) y el logging (3.2): catálogo de errores con códigos, mensajes UX consistentes, contrato de códigos de salida (para HU-15 CLI), política de niveles y propiedades de log buscables, y correlación log↔hilo de ejecución (RunId). Absorbe la deuda HU-13 (W-1/W-2/W-3, S-1..S-4), verificada una por una contra el código real (§0.1).
> **Rector (IRRENUNCIABLE):** `Docs/Propuesta_Proyecto_Automatizacion_Remuneracion_UAESP.md` — §9 Fase 3 it. **3.1** ("Errores de archivo faltante, formato incorrecto, etc.") y **3.2** ("Auditoría detallada de cada ejecución"), §6 (trazabilidad: "Cada paso se registra"), §10 CA-6 ("El log registra cada paso"). **Entran SOLO 3.1 + 3.2 + deuda listada.**
> **Continuidad:** HU-01..HU-13 cerradas (fail-fast honesto desde HU-07 que nombra ASE+reporte+celda; Serilog por paso/ASE/empresa/hoja en ~50 sitios; build 0 warnings; tests 160/160 + harness 24/24 regresión ciega). Este plan NO reabre su semántica: **cero cambios de comportamiento de cálculo/escritura; Q1/Q2 intactos por construcción; golden/Capa B no cambian (sin cambios de valores).**
> **Estado:** PENDIENTE DE APROBACIÓN — no implementar hasta OK del Ingeniero.
> **Fecha:** 2026-09-09

---

## 0. Clarification Gate

No hay pregunta bloqueante para el Ingeniero. La única incógnita aparente —si alguna deuda HU-13 ya estaba saldada— **se verificó contra el código real** (§0.1): las 7 están vigentes, ninguna se declara saldada. No hay T0 de workbook en esta HU: ningún gate nuevo exige semántica de celda no congelada (W-1 reutiliza la evidencia T0-0.4 de HU-13; S-4 es decisión de código, no de celda).

### 0.1 Qué está verificado y deuda HU-13 (evidencia por ítem)

**Verificado en esta planificación** (código leído):

| # | Hecho | Método | Consecuencia |
|---|---|---|---|
| V1 | Fail-fast honesto existe: `ProcesadorPeriodo` nombra ASE+reporte en cada fuente faltante (líneas 78-88, 115-126, 135-140) y `ValidadorBasico` nombra ASE+validación+celda en cada gate | Lectura `ProcesadorPeriodo.cs` + `ValidadorBasico.cs` | 3.1 = **cierre formal** (catálogo + códigos + UX), no invención del fail-fast |
| V2 | Solo **2 excepciones de dominio**: `ArchivoFuenteNoEncontradoException`, `CalculoInvalidoException` (mensajes ad-hoc, sin código, sin jerarquía por causa) | Glob `Core/Exceptions/` | Base del catálogo D1: extender sin romper (compat con 160 tests) |
| V3 | Serilog en ~50 sitios pero **config mínima**: `ConfigurarSerilog` = solo `.WriteTo.File("remuneracion_log.txt")` (sin nivel mínimo, sin rolling, sin template con propiedades) | Lectura `Form1.cs:485-490` | 3.2 = política de niveles + propiedades + rolling (§2.4) |
| V4 | Todo el detalle va a `Log.Information` (hitos, por-empresa, por-celda, cancelaciones); **cero `Warning`/`Debug`**; el único `catch` de UI mete todo a `Log.Error` + `MessageBox(ex.Message)` crudo | Grep `Log\.(Information\|Warning\|Error\|Debug)` en `Form1.cs` | Política de niveles D4 + mensajes UX D2 |
| V5 | `catch` desnudos de `OpenXmlPlantillaWriter:152,231` son **atomicidad** (borran salida parcial + `throw`), no swallowing | Lectura `OpenXmlPlantillaWriter.cs:130-240` | No es deuda; el catálogo 3.1 lo nombra como patrón a conservar |
| V6 | WinForms **sin códigos de salida** (correcto para GUI); `Herramientas/VerificadorRecaudo` sí usa `Environment.Exit(0/1)` pero es herramienta auxiliar, no el producto | Grep `Environment\.Exit` | 3.1 define el **contrato** `CodigosSalida` en Core para HU-15; Form1 lo registra, no lo emite |
| V7 | Sin correlación: ningún RunId/Periodo/AseId como propiedad estructurada sistemática; `Log.Information(mensaje)` plano en el `Progress<string>` (línea 207) | Lectura `Form1.cs:204-209` | Correlación D5 (RunId por ejecución) |

**Deuda HU-13 — verificación contra código (todas VIGENTES, ninguna saldada):**

| ID | Declaración | Evidencia | Veredicto |
|---|---|---|---|
| W-1 | Sub-booleanos VALIDACION_TOTAL (C15/D25/O25/D34:F34) declarados en docstring sin assert ni gate | `WorkbookLeafCellMapValidaciones.cs:19-20` los declara TRUE; `ProtegidasValidacionesParaPeriodo` solo cubre O3..O9/P3..P9 (líneas 87-92); `ValidadorBasico:313-322` solo gatea O9/P9 | **VIGENTE** → §2.5 (assert en mapa protegido + lectura + gate TRUE exacto; reutiliza T0-0.4 HU-13, sin T0 nuevo) |
| W-2 | Desviación de placement: `ValidarCruzadasContraResultado` vive en `ValidadorBasico`, no en `WorkbookLeafCoherence` | `WorkbookLeafCoherence.cs` (221 líneas) **no contiene** ningún método de cruzadas; el gate vive en `ValidadorBasico.ValidarGatesValidacionesCruzadasPorAse:249-323` (privado) | **VIGENTE como nota de aceptación** → §2.5 veredicto documentado: NO mover (evidencia: todos los gates cruzados HU-08..HU-12 — `ValidarSigmaEmpresasPorAse`, `ValidarReporteBanco`, `ValidarBalanceSc` — viven en `ValidadorBasico` con agregación de errores por lista; mover a `Coherence` (throw al primer fallo) cambiaría la semántica de agregación). Se salda con docstring + test de pinning, no con move |
| W-3 | `ForContext("Validacion","cruzada")` constante; el plan pedía el nombre de la validación | `Form1.cs:425-426`: `Log.ForContext("Validacion", "cruzada").Information("{Linea}", linea)` — literal único para las ~5 validaciones × 5 ASE | **VIGENTE** → §2.4 (propiedad `Validacion` = nombre real: `VALIDACION_ENEL`, `DetValiRetri`, `VALIDACION_TOTAL`, …) |
| S-1 | `ValidacionOracleReader` trata literal `"false"` como TRUE | `ValidacionOracleReader.cs:186-189`: con `DataType==Boolean`, `return texto != "0"` → `"false" != "0"` = **TRUE**; `"FALSE"` igual | **VIGENTE** → §2.5 (endurecer: `1`/`true`→TRUE; `0`/`false`→FALSE case-insensitive; otro texto → fail-fast nombrado) |
| S-2 | Gate VALIDACION_TOTAL duplicado ×5 (misma O9/P9 en los 5 snapshots) | `ValidacionOracleReader.cs:110-112` lee O9/P9 **dentro** de `LeerSnapshotAse` (×5, misma celda global); `ValidadorBasico:313-322` gatea por snapshot (×5 errores idénticos ante un fallo) | **VIGENTE** → §2.5 (leer una vez, gate único + assert de igualdad entre snapshots) |
| S-3 | `CrearAse` duplicado en 3 sitios productivos | `ProcesadorPeriodo.CrearAse:274-286` (deriva de `CarpetasAse.Prefijos` ✓), `ValidacionOracleReader.CrearAse:207-217` (**array hardcodeado** `{"Promoambiental","Lime",…}` ✗ — fuente divergente), `Form1.ParseAse:456-467` (parsea texto del combo) | **VIGENTE** → §2.5 (factoría única `AseFactory.DesdeId` en Core desde `CarpetasAse.Prefijos`; evidencia de drift latente: el array del reader duplica los nombres) |
| S-4 | `LeerValorNumerico` fail-open ante `<f>` sin `<v>` (retorna 0) | `ValidacionOracleReader.cs:163-168`: `CellValue` null/vacío → `return 0m` — para celdas-O críticas (fórmula sin caché = workbook nunca recalculado) el gate O=0±0.5 **pasa en falso** | **VIGENTE** → §2.5 decisión: celdas críticas de gate **exigen caché** (fail-fast nombra hoja+celda); solo controles informativos `Valida -*` conservan el 0 tolerado |

### 0.2 Mapeo al Rector (in vs out)

**Entra porque §9 it. 3.1 / 3.2 / §6 / §10 CA-6 lo piden ahora:**

| Rector | Qué cubre HU-14 |
|---|---|
| §9 it. 3.1 | Catálogo de errores (códigos por causa: fuente faltante, formato, validación, escritura/plantilla, inesperado), mensajes UX consistentes por categoría, contrato de códigos de salida para HU-15 |
| §9 it. 3.2 | Política de niveles (Information/Warning/Error/Debug), propiedades buscables (`RunId/Periodo/AseId/Hoja/Validacion/Empresa`), correlación log↔ejecución, Serilog con rolling + template estructurado |
| §6 trazabilidad | Cada paso registrado con RunId; cancelaciones y decisiones del usuario también logueadas como `Warning` |
| §10 CA-6 | El log registra cada paso — ahora con niveles y propiedades auditables |

**Sale porque §9 lo asigna a otras HUs (EXPLÍCITO):**

- 3.3 modo CLI (HU-15) — aquí solo el **contrato** de códigos de salida, ninguna CLI.
- INTERVENTORIA L25:N31 + filas L-Especiales (HU-16).
- 3.4 manual + entrega (HU-17).
- Cambios de comportamiento de cálculo/escritura (prohibidos; Q1/Q2 intactos; golden/Capa B no cambian).
- DI framework, restyle UI.

### 0.3 Decisiones ejecutivas (aprobación = aceptar estas)

| ID | Decisión |
|---|---|
| G1 | **Sin T0 de workbook**: W-1 reutiliza la evidencia T0-0.4 de HU-13 (sub-bloques TRUE en ambos goldens); S-4 es decisión de código. Ningún gate nuevo exige semántica de celda no congelada. |
| G2 | **Cero cambios de valores**: todo lo que altere un número escrito o un gate existente que pase a fallar contra los goldens actuales es NEEDS_CONTEXT, no "robustez". La red 160/160 + 24/24 debe seguir verde sin tocar goldens. |
| G3 | **PRs encadenados** (§4 Work Units): dominio (catálogo + factory + snapshot) → validador/readers (S-1/S-2/S-4/W-1) → UI+Serilog (niveles, UX, RunId, W-3) → regresión total. |
| G4 | **W-2 se salda sin mover código**: el placement en `ValidadorBasico` es consistente con todos los gates HU-08..HU-12; se registra el veredicto + pinning, no un move que cambiaría agregación→throw. |
| G5 | **S-4 endurece solo celdas de gate** (O/P, D-diferencias, O9/P9); los controles informativos conservan el 0 tolerado documentado. Si algún golden actual trae `<f>` sin `<v>` en celda de gate, el primer run lo revela → NEEDS_CONTEXT con recorte, nunca silenciar. |

---

## 1. PROPOSE

### 1.1 Intent

Cerrar Fase 3 (3.1 + 3.2) convirtiendo el fail-fast honesto y el logging existente en un sistema **formal y auditable**: cada fallo posible tiene un código, cada código tiene un mensaje UX consistente y un nivel de log, cada ejecución tiene un RunId que correlaciona todos sus eventos, y la deuda HU-13 queda saldada con evidencia — sin mover un solo valor de cálculo ni una sola celda escrita.

### 1.2 In Scope

- Catálogo de errores en Core: `CodigoError` + extensión no-rompiente de las 2 excepciones de dominio con `Codigo` (mensajes ganan prefijo `[CÓDIGO] … ASE n … reporte … celda … acción`).
- Mensajes UX centralizados por categoría (fuente / formato / validación / escritura / inesperado) + títulos de `MessageBox` consistentes en `Form1` (el box muestra guía accionable; el detalle técnico queda en el log).
- Contrato `CodigosSalida` en Core (0 OK; 1 validación; 2 fuente/plantilla; 3 escritura; 4 inesperado; 5 cancelado por usuario) — Form1 lo **registra** (status + log); HU-15 lo **emitirá**.
- Política de niveles Serilog + propiedades buscables (`RunId`, `Periodo`, `Quincena`, `Modo`, `AseId`, `Hoja`, `Validacion`, `Empresa`) + `ConfigurarSerilog` con rolling diario, nivel mínimo y template estructurado (misma base `remuneracion_log.txt` por convención).
- Correlación: RunId (Guid) por ejecución en ambos procesadores + UI, propagado por `LogContext`/`ForContext` (solo Serilog core, sin paquetes nuevos).
- Deuda: W-1 (assert+gate sub-booleanos), W-2 (veredicto documentado + pinning), W-3 (nombre real en `Validacion`), S-1 (booleano estricto), S-2 (gate TOTAL único), S-3 (factoría `AseFactory`), S-4 (`<v>` exigido en celdas de gate).
- Regresión 160/160 + harness 24/24 verde; build 0 warnings; CRLF; sin commits.

### 1.3 Out of Scope

Todo §0.2 (CLI, INTERVENTORIA, manual 3.4). Además: nuevos tipos de validación de negocio; cambiar tolerancia ±0.5; tocar mapas HU-07..HU-13 salvo las adiciones W-1/S-2/S-4; reescribir mensajes de progreso existentes (solo se les suma nivel/propiedades); paquetes NuGet nuevos; `Environment.Exit` en WinForms.

### 1.4 Resultado de negocio

El Ingeniero ejecuta como hoy y obtiene: (a) ante cualquier fallo, un **código + mensaje accionable** ("falta R2 del ASE 3 en …; genere el reporte … y reintente") en vez de un `ex.Message` crudo; (b) un log con niveles donde puede filtrar por `RunId`, `AseId` o `Validacion` para auditar una ejecución; (c) la certeza de que las 7 deudas HU-13 están cerradas con tests que lo demuestran; (d) el contrato de salida que HU-15 necesita para la CLI.

---

## 2. DESIGN

### 2.1 Tablas verificadas (contratables desde el día uno)

**Catálogo de errores 3.1 (causas reales del código actual):**

| Código | Causa | Dónde nace hoy | Mensaje UX (guía) | Nivel |
|---|---|---|---|---|
| `ERR-FUENTE-NO-ENCONTRADA` | Carpeta ASE / R1/R2/R4 / banco / balance / notas / retribución ausente | `ProcesadorPeriodo:78-88,115-140`, `Form1:246-248` | Qué reporte falta, de qué ASE, en qué carpeta; qué generar/reintentar | Error |
| `ERR-PLANTILLA` | Plantilla ausente, salida==plantilla, salida sin WorkbookPart | `Form1:176-182`, `OpenXmlPlantillaWriter:ValidarArchivo/NoInPlace` | Ruta esperada vs recibida; no usar la plantilla como salida | Error |
| `ERR-FORMATO-FUENTE` | Estructura inesperada (header ausente, celda no numérica, hoja faltante en oráculo) | Readers + `ObtenerCeldaFormula` (W2) | Archivo + hoja + celda + qué se esperaba encontrar | Error |
| `ERR-VALIDACION` | Gate de dominio/coherencia que no cierra (ya nombra ASE+validación) | `ValidadorBasico`, `WorkbookLeafCoherence` | Validación + ASE + valor vs esperado ±0.5 | Error |
| `ERR-ESCRITURA` | Fallo durante escritura (con atomicidad: salida parcial borrada, patrón V5) | `OpenXmlPlantillaWriter:152,231` | Ruta salida + causa + "no quedó archivo parcial" | Error |
| `ERR-INESPERADO` | Cualquier otra excepción | `Form1:223` catch-all | Mensaje genérico + RunId para buscar en el log | Error |
| `WARN-CANCELADO` | Usuario cancela (no sobrescribe salida) | `Form1:192-199` | Informativo (hoy `Information`, pasa a `Warning`) | Warning |

**Niveles 3.2 (política; hoy todo es `Information`, V4):**

| Nivel | Uso | Ejemplos |
|---|---|---|
| `Information` | Hitos de ejecución (inicio/fin por paso y ASE, GranTotal, resumen) | "Iniciando proceso del período", "ASE 3: leyendo R1, R2 y R4…", "Proceso completado" |
| `Warning` | Decisiones del usuario + divergencias conocidas documentadas (D21, J9:J14) | Cancelación por no-sobrescritura, D21 excluida (D6 HU-13) |
| `Error` | Todo fail-fast con excepción + código del catálogo | `ERR-FUENTE-NO-ENCONTRADA … ASE 3 …` |
| `Debug` | Detalle por celda/empresa/hoja (hoy `Information`, inunda) | Líneas por empresa R1/R2/R4, banco por empresa, BCE por ASE, VALIDACIONES por empresa |

**Propiedades buscables (todas vía `ForContext`/`LogContext`; W-3 fija `Validacion` real):**

| Propiedad | Valores | Origen |
|---|---|---|
| `RunId` | Guid por ejecución (ambos modos) | D5, generado en `Form1.btnEjecutar` / procesadores |
| `Periodo` / `Quincena` / `Modo` | `2026071`, `1`, `5 ASE` | Ya en mensajes; pasan a propiedades |
| `AseId` | 1..5 | Ya interpolado; pasa a propiedad |
| `Hoja` | `REPORTE RECAUDO x BANCO`, `BCE SC POR FACT.`, … (existentes, se conservan) | `Form1` actual |
| `Validacion` | **Nombre real**: `VALIDACION_ENEL`, `VALIDACION_RECIP`, `DetValiRetri`, `VALIDACION_TOTAL`, `Valida -Remunera`, … (W-3) | Nuevo (hoy literal `"cruzada"`) |
| `Empresa` | `ENEL`, `RECIPROCIDAD`, … | Ya interpolada; pasa a propiedad |

### 2.2 Decisiones de arquitectura

| ID | Opción elegida | Descartada | Por qué |
|---|---|---|---|
| D1 | `CodigoError` (enum/static de códigos string estables) + propiedad `Codigo` con default en las 2 excepciones existentes (constructores actuales intactos) | Nueva jerarquía de excepciones por causa | 160 tests referencian los tipos actuales; romper constructores o tipos es churn sin valor. El código viaja en la excepción, no en el tipo |
| D2 | Mensajes UX centralizados en Core (`CatalogoErrores`: código → título + plantilla de guía) consumidos por `Form1`; el `MessageBox` muestra guía, el log conserva el técnico | Formatear mensajes en el catch de UI | La guía es regla de negocio ("qué reporte generar"), no texto de UI; testeable in-memory sin WinForms |
| D3 | `CodigosSalida` (static Core: 0,1,2,3,4,5) + `Form1` registra `UltimoCodigoSalida` (status/log) sin `Environment.Exit` | Emitir exit codes desde WinForms | WinForms es GUI (V6); emitirlos rompería el ciclo de vida del form. HU-15 consumirá el contrato tal cual |
| D4 | Niveles por uso (§2.1) + `txtLog` sigue mostrando `IProgress<string>` (hitos), el detalle Debug va solo a archivo | Bajar todo el detalle a `txtLog` con filtros UI | Restyle/filtros UI = fuera de alcance; el archivo es el canal auditable (CA-6) |
| D5 | RunId vía `Serilog.Context.LogContext.PushProperty` al inicio de `Ejecutar` (ambos procesadores) + `Form1`; sin paquetes nuevos (LogContext es Serilog core) | ThreadStatic/AsyncLocal propio o `CorrelationId` en cada modelo | LogContext ya resuelve el flujo async (`Task.Run` existente); tocar modelos por observabilidad contamina el dominio |
| D6 | Serilog: rolling diario (`remuneracion_log_.txt`, 30 días), `MinimumLevel.Debug` (archivo) y template con `{RunId}{Periodo}{AseId}{Hoja}{Validacion}` | Cambiar de sink o de ruta base | Convención vigente (`./remuneracion_log.txt`, gitignored); el rolling conserva el prefijo y evita el archivo infinito |
| D7 | W-1: sub-booleanos como **assert protegido** (presencia `<f>` en mapa) + **lectura** (reader) + **gate TRUE exacto** (validador) con snapshot extendido; amparado en T0-0.4 HU-13 (TRUE en ambos goldens) | Solo assert sin gate | La deuda pide explícitamente assert **y** gate; la evidencia T0-0.4 elimina el riesgo de falsos positivos contra goldens |
| D8 | S-2: O9/P9 se leen **una vez** en `LeerSnapshots` (no por ASE); el gate corre **una vez** + assert interno de que los 5 snapshots portan el mismo valor (detecta bug del reader) | Mover TOTAL a un tipo global nuevo | Menor churn: el snapshot conserva los campos (compat), el reader los puebla desde una sola lectura, el validador gatea `[0]` + igualdad |

### 2.3 Flujo 3.1 + 3.2 (sin tocar cálculo/escritura)

```text
btnEjecutar → genera RunId → LogContext.PushProperty(RunId, Periodo, Modo)
  └─► IProgress<string> → txtLog (hitos, igual que hoy) + Serilog con nivel/propiedades
        └─► ProcesadorPeriodo.Ejecutar (lanza con Codigo + mensaje [CÓDIGO] ASE/reporte/celda/acción)
              └─► ValidadorBasico (errores agregados; cada uno con código ERR-VALIDACION)
                    └─► Writer atómico (falla → ERR-ESCRITURA, parcial borrado, V5)
  └─► catch (Exception ex): Codigo = ex.Codigo ?? ERR-INESPERADO
        → Log.Error(ex, "[{Codigo}] …", con RunId/props)
        → MessageBox.Show(CatalogoErrores.Guia(codigo, ex), títuloPorCategoria)  // UX, no ex.Message crudo
        → UltimoCodigoSalida = mapa(codigo); status "Error"
```

### 2.4 Dominio (Core, sin deps)

```csharp
public static class CodigoError  // D1: strings estables (no rompen serialización de mensajes)
{
    public const string FuenteNoEncontrada = "ERR-FUENTE-NO-ENCONTRADA";
    public const string Plantilla = "ERR-PLANTILLA";
    public const string FormatoFuente = "ERR-FORMATO-FUENTE";
    public const string Validacion = "ERR-VALIDACION";
    public const string Escritura = "ERR-ESCRITURA";
    public const string Inesperado = "ERR-INESPERADO";
    // Warning (no es fallo): "WARN-CANCELADO"
}

public static class CodigosSalida  // D3: contrato para HU-15
{
    public const int Ok = 0; public const int Validacion = 1;
    public const int FuenteOPlantilla = 2; public const int Escritura = 3;
    public const int Inesperado = 4; public const int CanceladoPorUsuario = 5;
}

public static class CatalogoErrores  // D2: código → (título box, plantilla de guía)
{
    public static (string Titulo, string Guia) Para(string codigo, Exception ex);
    public static int CodigoSalidaPara(string codigo);  // mapa código → CodigosSalida
}

public static class AseFactory  // S-3: fuente única desde CarpetasAse.Prefijos
{
    public static Ase DesdeId(int id);  // "N-Nombre" → Ase; id fuera de 1..5 = fail-fast
}
```

Extensiones no-rompientes: `ArchivoFuenteNoEncontradoException` y `CalculoInvalidoException` ganan `string Codigo { get; }` (default `ERR-FUENTE-NO-ENCONTRADA` / `ERR-VALIDACION`) + overload con `codigo` explícito; constructores actuales intactos. `ValidacionCruzadaSnapshot` gana `IReadOnlyDictionary<string,bool> SubBloquesValidacionTotal` (W-1: C15/D25/O25/D34/F34) — aditivo, default vacío = HU-13 puro.

### 2.5 Cierre de deuda HU-13 (diseño por ítem)

1. **W-1** (D7): mapa `ProtegidasValidacionesParaPeriodo` suma `(VALIDACION_TOTAL, C15/D25/O25/D34/F34)` con presencia `<f>`; `ValidacionOracleReader` los lee vía `ObtenerCeldaFormula` + `LeerValorBooleano` estricto (tras S-1); `ValidadorBasico` gatea TRUE exacto por ASE con error que nombra `ASE n · VALIDACION_TOTAL · C15`. Amparo: T0-0.4 HU-13 (TRUE en ambos goldens) — sin T0 nuevo (G1). Riesgo residual: si un golden trae alguno en FALSE, el gate lo revela en regresión → NEEDS_CONTEXT (G5), nunca silenciar.
2. **W-2**: veredicto documentado (no mover): docstring de `ValidarGatesValidacionesCruzadasPorAse` cita la razón (consistencia con `ValidarSigmaEmpresasPorAse`/`ValidarReporteBanco`/`ValidarBalanceSc` + agregación por lista vs throw) + test de pinning (el gate sigue en `ValidadorBasico`, `WorkbookLeafCoherence` sigue sin cruzadas). La nota de aceptación del Plan 13 §2.8/§4-3.1 queda enmendada en §9.7.
3. **W-3**: `Form1` bloque VALIDACIONES emite `ForContext("Validacion", nombreReal)` por línea (empresa → `VALIDACION_{EMPRESA}`, DetValiRetri → `DetValiRetri`, total → `VALIDACION_TOTAL`) + `AseId` como propiedad; se conserva `Hoja` donde existe. Mensajes con template estructurado (no solo `{Linea}` preformateada).
4. **S-1**: `LeerValorBooleano` estricto — `t="b"`: solo `"1"`=TRUE, `"0"`=FALSE (cualquier otro texto → fail-fast nombrado); sin `t="b"`: `"1"/"true"`→TRUE, `"0"/"false"`→FALSE (case-insensitive), numérico ≠0→TRUE, vacío→FALSE (caché ausente, conserva S-4 para informativos), resto→fail-fast nombrado. `"false"` literal → FALSE en todas las ramas.
5. **S-2** (D8): `LeerSnapshots` lee O9/P9 una vez (fuera del loop por ASE), puebla los 5 snapshots con el mismo valor + el validador gatea una vez (snapshot[0]) y aserta igualdad en los 5 (divergencia = bug del reader, fail-fast). Conteo de errores ante TOTAL roto: 1, no 5.
6. **S-3**: `AseFactory.DesdeId` (Core, desde `CarpetasAse.Prefijos`); `ProcesadorPeriodo.CrearAse` y `ValidacionOracleReader.CrearAse` delegan (se elimina el array hardcodeado); `Form1.ParseAse` conserva el parseo del combo (UI) pero valida contra `CarpetasAse.Prefijos`. Helpers de tests (`GoldenLeafPathTests`, `WorkbookLeafInputsTests`) migran a la factory (cambio mecánico).
7. **S-4**: `LeerValorNumerico` se desdobla — `LeerValorNumericoExigido` (celdas de gate: O/P por empresa, D-diferencias/verificaciones, O9/P9, D21/D29: `<v>` ausente → `CalculoInvalidoException` que nombra hoja+celda + "fórmula sin caché; recalcule el workbook") vs `LeerValorNumericoSiExiste` (controles informativos: conserva el 0). Decisión registrada: un gate que lee 0 donde no hay caché es peor que un fallo nombrado (doctrina W2: nunca 0 silencioso en gates).

### 2.6 UI — Visual Design Intent (delta mínimo)

Densidad Balanced, mismos controles, sin restyle/colores/iconos. Cambios: (a) `MessageBox` por categoría del catálogo (título = `CatalogoErrores.Titulo`, texto = guía + código; detalle técnico solo en log); (b) cancelación por no-sobrescritura loguea `Warning` con `WARN-CANCELADO`; (c) `txtLog` sin cambios visuales (el detalle Debug no inunda el box); (d) status strip muestra el código de salida registrado (`Error ERR-… (salida 2)`). Sin nuevos controles.

### 2.7 Golden y regresión (sin cambios de valores)

No hay golden nuevo (ningún valor cambia). La red es: 160/160 + harness 24/24 verdes **sin modificar** + suites nuevas (§5) que solo agregan: catálogo/UX in-memory, niveles/propiedades (Serilog TestSink o appender en memoria — sin paquetes nuevos: `Serilog.Sinks.File` ya referenciado; para asserts se usa un `ILogEventSink` propio de test), RunId propagado, W-1/S-1/S-2/S-4 contra goldens reales (read-only), y pinning W-2/S-3.

### 2.8 File changes previstos

| Archivo | Acción | Motivo |
|---|---|---|
| `Remuneracion.Core/Errors/CodigoError.cs` (+ `CodigosSalida`, `CatalogoErrores`) | Crear | Catálogo 3.1 + contrato CLI (D1/D2/D3) |
| `Remuneracion.Core/Exceptions/*.cs` (2 archivos) | Modificar | Propiedad `Codigo` + overload, constructores intactos (D1) |
| `Remuneracion.Core/Services/AseFactory.cs` | Crear | Fuente única S-3 (desde `CarpetasAse.Prefijos`) |
| `Remuneracion.Core/Models/ValidacionCruzadaSnapshot.cs` | Modificar | `SubBloquesValidacionTotal` aditivo (W-1) |
| `Remuneracion.Core/Services/ValidadorBasico.cs` | Modificar | Prefijo `[CÓDIGO]` en errores (ERR-VALIDACION), gate W-1, gate TOTAL único (S-2), docstring W-2 |
| `Remuneracion.Core/Services/ProcesadorPeriodo.cs` + `ProcesadorRemuneracion.cs` | Modificar | Códigos en fail-fasts, `AseFactory`, RunId/LogContext (D5) |
| `Remuneracion.Infrastructure/Excel/WorkbookLeafCellMapValidaciones.cs` | Modificar | Protegidas W-1 (C15/D25/O25/D34/F34) |
| `Remuneracion.Infrastructure/Excel/ValidacionOracleReader.cs` | Modificar | S-1 booleano estricto, S-2 lectura única O9/P9, S-4 caché exigido en gates, `AseFactory`, lectura W-1 |
| `Remuneracion.WinForms/Form1.cs` | Modificar | UX por catálogo (D2), niveles+props (W-3, D4), RunId (D5), `UltimoCodigoSalida` (D3) |
| `Remuneracion.IntegrationTests/CatalogoErroresTests.cs` | Crear | Código→título/guía/salida; defaults de excepciones; compat constructores |
| `Remuneracion.IntegrationTests/ObservabilidadTests.cs` | Crear | Niveles por evento, props `RunId/AseId/Validacion`, W-3 nombre real, RunId correlaciona inicio→fin |
| `Remuneracion.IntegrationTests/RobustezDeudaHu13Tests.cs` | Crear | W-1 (gate TRUE + goldens), S-1 (`false`→FALSE), S-2 (1 error, no 5), S-4 (`<f>` sin `<v>` falla nombrando), S-3 (factory única), W-2 (pinning placement) |
| Tests helpers (`GoldenLeafPathTests`, `WorkbookLeafInputsTests`) | Modificar | Migración mecánica a `AseFactory` (S-3) |

**No tocar:** cálculo/escritura/mapas HU-07..HU-13 (salvo adiciones listadas); tolerancia ±0.5; `requirements/` legado; goldens; `Herramientas/VerificadorRecaudo` (auxiliar).

---

## 3. SPEC / Acceptance Criteria (trazabilidad Propuesta §9–§10)

### Requirement 1 — Catálogo con códigos (3.1; CA-5)

Todo fail-fast del path período **MUST** portar un `Codigo` del catálogo en el mensaje (`[ERR-…]` + ASE/reporte/celda/acción). **MUST NOT** romper constructores existentes.

- GIVEN carpeta ASE 3 sin R2 → WHEN `Ejecutar` → THEN `ArchivoFuenteNoEncontradoException` con `Codigo == ERR-FUENTE-NO-ENCONTRADA` y mensaje que nombra ASE 3 + R2 + carpeta.
- GIVEN gate DetValiRetri roto → THEN error con `[ERR-VALIDACION]` que nombra ASE+celda (formato actual + prefijo).
- GIVEN suites 160 existentes → THEN verdes sin cambios (compat).

### Requirement 2 — Mensajes UX consistentes (3.1; CA-7)

El `MessageBox` **MUST** mostrar título+guía del catálogo (accionable), **MUST NOT** mostrar `ex.Message` crudo como texto principal.

- GIVEN `ERR-FUENTE-NO-ENCONTRADA` en UI → THEN título "Archivo fuente faltante" + guía con reporte/ASE/carpeta + código visible.
- GIVEN excepción inesperada → THEN guía genérica + RunId ("búsquelo en remuneracion_log…").

### Requirement 3 — Contrato de códigos de salida (3.1; base HU-15)

`CodigosSalida` **MUST** existir en Core con el mapa código→salida; `Form1` **MUST** registrar `UltimoCodigoSalida` en cada cierre (OK/cancelado/fallo). **MUST NOT** llamar `Environment.Exit` en WinForms.

- GIVEN fallo `ERR-VALIDACION` → THEN `UltimoCodigoSalida == 1` + log con el código.
- GIVEN cancelación del usuario → THEN salida 5 + `Warning WARN-CANCELADO`.

### Requirement 4 — Niveles y propiedades buscables (3.2; CA-6)

Cada evento **MUST** salir con el nivel §2.1 y las propiedades aplicables; `Validacion` **MUST** ser el nombre real (W-3). **MUST NOT** quedar ningún `Log.Information` con detalle por celda/empresa (pasa a `Debug`).

- GIVEN ejecución 5 ASE → WHEN filtrar por `RunId` → THEN todos sus eventos (inicio→GranTotal o Error).
- GIVEN bloque VALIDACIONES → THEN eventos con `Validacion == "VALIDACION_ENEL"` (no `"cruzada"`) + `AseId`.
- GIVEN cancelación → THEN `Warning` (no `Information`).

### Requirement 5 — Serilog operativo (3.2; CA-6)

`ConfigurarSerilog` **MUST** usar rolling diario (mismo prefijo, 30 días), `MinimumLevel.Debug` y template con `{Timestamp} [{Level}] (RunId/Periodo/AseId/Hoja/Validacion) mensaje`. **MUST NOT** cambiar la ruta base fuera de `./remuneracion_log*` ni agregar paquetes.

### Requirement 6 — W-1: sub-booleanos con assert y gate (CA-4/CA-5)

C15/D25/O25/D34:F34 de VALIDACION_TOTAL **MUST** estar en el mapa protegido (presencia `<f>`) y gateados TRUE exacto por ASE contra ambos goldens (amparo T0-0.4 HU-13).

- GIVEN goldens Q1/Q2 → THEN gate verde (son TRUE).
- GIVEN C15=FALSE en snapshot in-memory → THEN error `ASE n · VALIDACION_TOTAL · C15`.

### Requirement 7 — S-1/S-2/S-4 del reader-oráculo (CA-1/CA-2)

El reader **MUST**: `"false"`→FALSE en toda rama (S-1); leer O9/P9 una vez con gate único (S-2: un TOTAL roto = 1 error); exigir `<v>` en celdas de gate con fallo nombrado (S-4). **MUST NOT** cambiar el 0 tolerado de controles informativos.

- GIVEN celda-P booleana con texto `"false"` → THEN `VerificacionP == false` (hoy TRUE).
- GIVEN O9=2.0 → THEN exactamente 1 error `VALIDACION_TOTAL (O9)` en 5 snapshots.
- GIVEN `<f>` sin `<v>` en O-empresa → THEN `CalculoInvalidoException` nombra hoja+celda (no 0 silencioso).

### Requirement 8 — S-3/W-2 estructurales (CA-5)

`AseFactory.DesdeId` **MUST** ser la única fuente (los 3 sitios delegan; el array hardcodeado del reader desaparece). El gate de cruzadas **MUST** permanecer en `ValidadorBasico` (pinning test; `WorkbookLeafCoherence` sin cruzadas por diseño documentado).

| CA §10 | HU-14 |
|---|---|
| CA-6 | Niveles + props + RunId + rolling; cada paso auditable por `RunId` |
| CA-5 | Catálogo con códigos, gates W-1/S-2 con matcheo estricto, S-3 factory, W-2 pinning |
| CA-7 | Mismo flujo 5-ASE + UX por catálogo (sin snapshot = HU-13 puro intacto) |
| CA-1/CA-2 | S-1/S-4 endurecen lectura-oráculo sin cambiar semántica de validación |
| CA-3/CA-4 | Sin cambios de valores (no hay golden nuevo); protegidas extendidas W-1 |

---

## 4. TASKS

## Review Workload Forecast

| Field | Value |
|---|---|
| Estimated changed lines | 500–800 |
| 400-line budget risk | Medium-High |
| Chained PRs recommended | Yes |
| Suggested split | PR1 dominio (catálogo+factory+snapshot) → PR2 validador+readers (S-1/S-2/S-4/W-1) → PR3 UI+Serilog (niveles/UX/RunId/W-3) → PR4 tests+regresión |
| Delivery strategy | ask-on-risk |
| Chain strategy | pending |

Decision needed before apply: Yes. Chained PRs recommended: Yes. Chain strategy: pending.

### Suggested Work Units

| Unit | Goal | Likely PR | Notes |
|---|---|---|---|
| 1 | Dominio: catálogo, salida, factory, snapshot W-1 | PR 1 | Sin deps; tests in-memory desde el día uno |
| 2 | Validador + oracle-reader (S-1/S-2/S-4/W-1, docstring W-2) | PR 2 | Depende de PR 1; read-only contra goldens |
| 3 | UI + Serilog (D2/D3/D4/D5/D6, W-3) | PR 3 | Depende de PR 1–2 |
| 4 | Suites nuevas + regresión 160/160 + 24/24 + 0 warnings | PR 4 | Depende de PR 1–3 |

### Phase 1 — Dominio (catálogo + factory + snapshot)

- [ ] 1.1 `CodigoError` + `CodigosSalida` + `CatalogoErrores` (D1/D2/D3: títulos, guías, mapa a salida) + `AseFactory.DesdeId` (S-3, desde `CarpetasAse.Prefijos`, id 1..5 o fail-fast).
- [ ] 1.2 `Codigo` en las 2 excepciones (defaults + overload; constructores actuales intactos) + `SubBloquesValidacionTotal` en `ValidacionCruzadaSnapshot` (W-1, default vacío).
- [ ] 1.3 `CatalogoErroresTests`: código→título/guía/salida por categoría; defaults; compat constructores (regresión de tipos).

### Phase 2 — Validador + oracle-reader (deuda funcional)

- [ ] 2.1 S-1: `LeerValorBooleano` estricto (§2.5.4) + casos (`"false"/"FALSE"/"0"`→FALSE; basura→fail-fast nombrado).
- [ ] 2.2 S-4: `LeerValorNumericoExigido` en celdas de gate (O/P, D-diffs/verifs, O9/P9, D21/D29) + informativos intactos; mensaje "fórmula sin caché; recalcule".
- [ ] 2.3 S-2: O9/P9 una lectura en `LeerSnapshots` + gate único + assert de igualdad ×5 (1 error, no 5).
- [ ] 2.4 W-1: protegidas C15/D25/O25/D34/F34 + lectura + gate TRUE exacto nombrando ASE·celda (amparo T0-0.4 HU-13 citado en test).
- [ ] 2.5 `ProcesadorPeriodo`/`ProcesadorRemuneracion`: códigos en fail-fasts + `AseFactory` (S-3) + RunId/LogContext (D5); docstring W-2 en `ValidadorBasico` (veredicto, sin mover).
- [ ] 2.6 `ValidacionOracleReader.CrearAse` → delega a `AseFactory` (muere el array hardcodeado).

### Phase 3 — UI + Serilog (observabilidad)

- [ ] 3.1 `ConfigurarSerilog` (D6: rolling, Debug, template con props) + niveles §2.1 (hitos `Information`, detalle `Debug`, cancelación `Warning`, fallos `Error`).
- [ ] 3.2 `MessageBox` por catálogo (D2) + `UltimoCodigoSalida` (D3) + RunId por ejecución (D5) + bloque VALIDACIONES con `Validacion` real + `AseId` prop (W-3).

### Phase 4 — Pruebas y regresión

- [ ] 4.1 `ObservabilidadTests` (Req 4–5: niveles, props, W-3, correlación RunId inicio→fin).
- [ ] 4.2 `RobustezDeudaHu13Tests` (Req 6–8: W-1 ambos goldens, S-1, S-2 conteo=1, S-4 fail nombrado, S-3 única fuente, W-2 pinning).
- [ ] 4.3 Migración mecánica de helpers de tests a `AseFactory`; 160/160 + 24/24 verdes; build 0 warnings; CRLF; sin commit.

---

## 5. Test Strategy

| Capa | Qué | Cómo |
|---|---|---|
| Unidad | Catálogo (código→título/guía/salida), defaults de excepciones, `AseFactory` (ids 1..5 + fuera de rango) | In-memory, sin I/O |
| Unidad | S-1 (matriz `"1"/"0"/"true"/"false"/"FALSE"/vacío/basura"` × con/sin `t="b"`) | Celdas OpenXML sintéticas en memoria |
| Unidad | S-2 (O9 roto = 1 error), W-1 (C15 FALSE = error nombrado), W-2 (gate en `ValidadorBasico` vía reflexión/llamada) | Snapshots in-memory |
| Integración | W-1/S-2/S-4 contra goldens Q1+Q2 reales (read-only, sin salida) | `ValidacionOracleReader` + `ValidadorBasico` |
| Integración | Observabilidad (niveles/props/RunId con sink en memoria propio, sin paquetes) | Ejecución período en temp + asserts de eventos |
| Regresión | 160/160 + harness 24/24 sin cambios | Suites existentes intactas |
| UI | `MessageBox` por categoría + status con código | Funcional manual (sin harness) |

### 5.1 Casos negativos obligatorios (nombran código + ASE/celda)

Carpeta ASE 3 sin R2 (`ERR-FUENTE-NO-ENCONTRADA`); salida==plantilla (`ERR-PLANTILLA`); hoja oráculo ausente (`ERR-FORMATO-FUENTE` W2 intacto); O-empresa=2.0 (`ERR-VALIDACION`); P con `"false"` literal (S-1: FALSE, no TRUE); O9=2.0 (S-2: 1 error); `<f>` sin `<v>` en O3 (S-4: falla nombrando); C15=FALSE (W-1); `AseFactory.DesdeId(6)` (S-3); cancelación usuario (salida 5 + `Warning`); excepción genérica (`ERR-INESPERADO` + RunId).

---

## 6. Architecture Validation Certificate

### 6.1 SOLID

| Principio | Veredicto |
|---|---|
| **S** | Observabilidad = extensión transversal (catálogo, niveles, RunId); cálculo/validación/escritura conservan su rol HU-01..HU-13. |
| **O** | Excepciones extendidas por propiedad, no por jerarquía; snapshot extendido por campo aditivo; niveles sin tocar mensajes de progreso. |
| **L** | Constructores actuales intactos: todo código que captura los tipos actuales sigue compilando y comportándose igual + `Codigo`. |
| **I** | `CatalogoErrores`/`AseFactory`/`CodigosSalida` estáticos sin deps; ningún contrato existente cambia de forma. |
| **D** | Core define catálogo/factory/códigos; Infrastructure/WinForms componen. Cero paquetes nuevos. |

### 6.2 Best Practices

- La verdad del workbook manda: ningún gate nuevo inventa semántica (W-1 amparado en T0-0.4 HU-13; S-4 endurece lectura, no redefine celdas).
- Nunca 0 silencioso en gates (doctrina W2 HU-13, extendida por S-4 a `<v>` ausente).
- Un solo código por fallo, un solo gate TOTAL (S-2), una sola factory (S-3), un solo nombre de validación por evento (W-3).
- UX accionable en el box, técnico en el log (D2); GUI no emite exit codes (V6; contrato para HU-15).
- Golden honesto: sin cambios de valores no hay golden nuevo; la red 160/160 + 24/24 es el gate.

### 6.3 Performance

5 ASE × lectura única O9/P9 (menos I/O que hoy) + gates in-memory + un `Guid` por ejecución. Irrelevante a esta escala; `Task.Run` existente intacto.

**Veredicto:** APROBADO como 3.1 + 3.2 con absorción W-1/W-2/W-3, S-1..S-4 **si** se respeta G2 (cualquier gate nuevo que falle contra goldens actuales = NEEDS_CONTEXT, no robustez) y G4 (W-2 sin mover código).

---

## 7. Riesgos

| Riesgo | Likelihood | Mitigación |
|---|---|---|
| W-1 revela un sub-booleano en FALSE en algún golden (T0-0.4 citado, no re-verificado aquí) | Baja | G5: NEEDS_CONTEXT con recorte (gate a los verificados), nunca silenciar; el primer run de PR2 lo revela gratis |
| S-4 revela `<f>` sin `<v>` en celda de gate en goldens actuales (workbook nunca recalculado) | Media | G5: NEEDS_CONTEXT con recorte a las celdas con caché; la alternativa (seguir retornando 0) perpetúa el fail-open |
| Bajar detalle a `Debug` "oculta" info que el Ingeniero lee en el archivo | Baja | El archivo conserva todo (`MinimumLevel.Debug`); solo el `txtLog` respira; documentado en HU-17 (manual) |
| Rolling cambia el nombre efectivo del log y rompe un hábito (`remuneracion_log.txt` exacto) | Baja | Prefijo conservado (`remuneracion_log_*.txt`); el día 1 coexisten; se documenta en el manual HU-17 |
| Inflar a HU-15 (CLI) dentro de esta HU (el contrato tienta) | Media | §0.2 out-of-scope explícito; rechazar PRs que emitan `Environment.Exit` o parseen args |
| `CatalogoErrores` crece con textos de negocio no validados por el Ingeniero | Media | Plantillas de guía mínimas (qué falta + dónde + reintente); el tono final lo fija la revisión del PR3 |
| S-3 rompe tests que construyen `Ase` a mano con nombres divergentes | Baja | Migración mecánica en PR4; la factory acepta exactamente los 5 nombres de `CarpetasAse.Prefijos` |

---

## 8. Rollback

- Revertir PRs en orden inverso (regresión → UI/Serilog → validador/readers → dominio).
- HU-01..HU-13 intactas sin esta HU: excepciones sin `Codigo` = comportamiento actual; Serilog mínimo actual; `ForContext("Validacion","cruzada")` actual.
- `Docs/Insumos/` untracked: jamás como destino; pisada se restaura desde backup.
- Si W-1/S-4 revelan divergencia contra goldens (G5), recorte con rebase a lo verificado, no silenciamiento.

---

## 9. Decisiones fijadas por este plan

1. Rector = Propuesta §9 it. 3.1/3.2 + §6 trazabilidad + §10 CA-6. Entran 3.1 + 3.2 + deuda listada; 3.3 (HU-15), INTERVENTORIA (HU-16), 3.4 (HU-17) quedan fuera.
2. Cierre formal sin invención: el fail-fast (V1) y el logging (V3/V4) existen; 3.1 les pone catálogo+códigos+UX, 3.2 niveles+propiedades+correlación+rolling.
3. Sin T0 de workbook (G1): W-1 amparado en T0-0.4 HU-13; S-4 es decisión de código; ningún gate exige semántica nueva de celda.
4. G2 (hierro): cero cambios de valores; gate nuevo que falle contra goldens = NEEDS_CONTEXT.
5. Deuda HU-13: las 7 vigentes con evidencia (§0.1); W-1 assert+gate, W-3 nombre real, S-1/S-2/S-4 en reader+validador, S-3 factory única.
6. W-2 se salda **sin mover código** (G4): placement en `ValidadorBasico` consistente con todos los gates HU-08..HU-12 (agregación por lista); pinning test + docstring; enmienda a la nota del Plan 13 §2.8/§4-3.1.
7. S-4 desdobla lectura: caché exigido en gates, 0 tolerado solo en informativos (nunca 0 silencioso en gates).
8. Contrato `CodigosSalida` en Core para HU-15; WinForms registra, no emite (V6).
9. PRs encadenados 1→4 (§4); regresión 160/160 + 24/24 como red; build 0 warnings; CRLF; sin commits.

**Listo para aprobación del Ingeniero. No implementar hasta OK.**
