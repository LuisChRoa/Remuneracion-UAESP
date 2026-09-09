# Plan 15 — HU-15: Modo CLI (Fase 3, it. 3.3)

> **Historia:** ejecución desatendida sin intervención manual (Propuesta §9 Fase 3 it. 3.3: "Ejecución sin intervención manual (CLI)"): proyecto consola que replica el flujo WinForms (período/carpeta/plantilla/salida/modo), consume el contrato `CodigosSalida` 0-5 de HU-14 con `Environment.Exit`, y habilita pruebas repetibles/desatendidas (objetivo declarado de HU-15 en el mapa de cierre).
> **Rector (IRRENUNCIABLE):** `Docs/Propuesta_Proyecto_Automatizacion_Remuneracion_UAESP.md` — §9 Fase 3 it. **3.3** (una sola fila: "Modo automático | Ejecución sin intervención manual (CLI)"), §6 (trazabilidad: "Cada paso se registra"), §10 CA-6 (el log registra cada paso). **Entran SOLO 3.3 + deuda listada en §0.1.**
> **Continuidad:** HU-01..HU-13 cerradas; HU-14 (3.1 + 3.2) implementada en working tree sin commit (catálogo `CodigoError`/`CatalogoErrores`, contrato `CodigosSalida`, `AseFactory`, RunId por ejecución, niveles/propiedades Serilog). Este plan NO reabre su semántica: **cero cambios de comportamiento de cálculo/escritura/validación; Q1/Q2 intactos por construcción; sin golden nuevo (paridad = mismos valores, no nuevo oráculo).**
> **Estado:** PENDIENTE DE APROBACIÓN — no implementar hasta OK del Ingeniero.
> **Fecha:** 2026-09-09

---

## 0. Clarification Gate

No hay pregunta bloqueante para el Ingeniero. Las ambigüedades aparentes (forma de los args, parseo manual vs paquete, default de modo, `--salida` como carpeta o archivo, mapeo de errores de uso, destino de S-2..S-5/S-1-newlines) **se dictaminan ejecutivamente** en §0.1–§0.3 con trade-offs; la aprobación del plan las fija. No hay T0 de workbook en esta HU: ningún gate nuevo exige semántica de celda no congelada.

### 0.1 Qué está verificado y deuda HU-14 (evidencia por ítem)

**Verificado en esta planificación** (código leído del working tree):

| # | Hecho | Método | Consecuencia |
|---|---|---|---|
| V1 | `CodigosSalida` 0-5 existe en Core (`Ok/Validacion/FuenteOPlantilla/Escritura/Inesperado/CanceladoPorUsuario`); `CatalogoErrores.CodigoSalidaPara` mapea cada código del catálogo | Lectura `CodigosSalida.cs` + `CatalogoErrores.cs` | El CLI los **CONSUME tal cual** (D1); no se redefine ningún valor |
| V2 | Único `Environment.Exit` del repo vive en `Herramientas/VerificadorRecaudo` (harness auxiliar HU-02, fuera del `.slnx`, no es producto) | Grep `Environment\.Exit` | El `Environment.Exit` del CLI será la **primera y única salida con Exit del producto** (slnx); el harness no se toca |
| V3 | Ambos procesadores generan `RunId = Guid.NewGuid()` **internamente** (`ProcesadorPeriodo.cs:56`, `ProcesadorRemuneracion.cs:42`); `SolicitudProcesoPeriodo`/`SolicitudProcesoAse` **no tienen** propiedad RunId | Lectura procesadores + solicitudes | W-2.1 VIGENTE → §2.5 (propiedad `RunId` nullable; el procesador respeta el inyectado) |
| V4 | `OpenXmlPlantillaWriter.ValidarArchivo`/`ValidarRutaSalida` lanzan `ArchivoFuenteNoEncontradoException` **sin código explícito** (default `ERR-FUENTE-NO-ENCONTRADA`); `ValidarNoInPlace` lanza `CalculoInvalidoException` sin código (**default `ERR-VALIDACION` → salida 1, errónea**: es un problema de plantilla, salida 2); todos los `CalculoInvalidoException` de estructura en `GenerarWorkbook` (hoja ausente, fórmula ausente, valor fijo) portan default `ERR-VALIDACION`; los `catch` de atomicidad re-lanzan sin envolver (un `IOException` en copia/escritura saldría como `ERR-INESPERADO`, nunca `ERR-ESCRITURA`) | Lectura `OpenXmlPlantillaWriter.cs:111-240,335-364` + defaults de excepciones | W-2.3 VIGENTE → §2.5 (códigos explícitos en el writer; `ERR-ESCRITURA` real con inner preservado) |
| V5 | La denegación salida==plantilla de `Form1` (líneas 196-205) **no emite ningún `Log.Error`**: fija `UltimoCodigoSalida`, muestra el box y retorna — sin traza auditable | Lectura `Form1.cs:196-205` | W-2.4 VIGENTE → §2.5 (`Log.Error` en `Form1` + en la denegación equivalente del CLI) |
| V6 | S-1 "false": `LeerValorBooleano` ya es estricto (`"1"/"true"→TRUE`, `"0"/"false"→FALSE` case-insensitive en toda rama, resto fail-fast nombrado; `ValidacionOracleReader.cs:220-261`) | Lectura + grep | **SALDADO con evidencia — no se replanifica** |
| V7 | S-2 (O9/P9 una lectura): implementado (`ValidacionOracleReader.cs:51-57`, fuera del loop); S-3 (`AseFactory`): existe y delegan `ProcesadorPeriodo:93` + `Form1:577`; S-4 (`LeerValorNumericoExigido`/`SiExiste`): implementado (líneas 178-201) | Lectura + grep | **SALDADAS con evidencia — no se replanifican** |
| V8 | Composición idéntica lista para reutilizar: `Program.cs:15-40` (lector, leafReader, cálculo, validador, writer, locator, oráculo) + `SolicitudProcesoPeriodo`/`SolicitudProcesoAse` + `Periodo.Parse("2026071")` + `Periodo.NombreArchivo` | Lectura `Program.cs`, `Periodo.cs` | El CLI compone los mismos servicios; `--salida` = carpeta + `NombreArchivo` (igual que `Form1:181`) |
| V9 | `.slnx` con 4 proyectos (`WinForms`, `Core`, `Infrastructure`, `IntegrationTests`); test project `net10.0` con refs a Core+Infra; Serilog 4.4.0 + Sinks.File 7.0.0 ya en el árbol | Lectura `.slnx` + `.csproj` | Proyecto nuevo sin paquetes nuevos (D2); se suma al `.slnx` como 5.º proyecto |
| V10 | No existe S-5 en el Plan 14 (S-1..S-4) ni archivo de revisión HU-14 en el repo; el working tree HU-14 = código + 3 suites nuevas, sin commit | Glob `plans/`, `git status` | S-2..S-5 "de suggestions" y S-1-newlines → §0.2 + Task 0 (pre-vuelo con gate) |

**Deuda que entra a HU-15 (toda VIGENTE, verificada arriba):** W-2.1 (RunId inyectable), W-2.3 (códigos del writer), W-2.4 (denegación auditable). S-1-"false", S-2, S-3, S-4 del Plan 14 están **saldadas** (V6/V7) y no entran.

### 0.2 Mapeo al Rector (in vs out) + deuda derivada

**Entra porque §9 it. 3.3 / §6 / §10 CA-6 lo piden ahora:**

| Rector | Qué cubre HU-15 |
|---|---|
| §9 it. 3.3 | Proyecto consola `net10.0` en el `.slnx` con parseo de args (período, carpeta, plantilla, salida, modo 1-ASE/5-ASE, quincena) + `--help` + validación con mensajes del catálogo HU-14 |
| §9 it. 3.3 + §6 | `Environment.Exit` con `CodigosSalida` 0-5 (primera y única salida con Exit del producto); RunId único por ejecución que el procesador respeta (W-2.1); denegación salida==plantilla con `Log.Error` (W-2.4) |
| §10 CA-6 | Misma ejecución, mismos valores: test de paridad CLI↔WinForms ante mismos insumos (sin golden nuevo); pruebas repetibles/desatendidas |

**Deuda HU-14 restante — veredicto:**

| ID | Veredicto |
|---|---|
| S-1 "false" | Saldado (V6). No entra. |
| S-2/S-3/S-4 del Plan 14 | Saldadas (V7). No entran. |
| S-2..S-5 "de suggestions" (revisión HU-14) | **No localizadas en el repo** (V10). Task 0 las reclama al orquestador con gate: solo entra lo que afecte corrección del CLI o del path compartido; cosmética/estilo → HU-17; lo ya saldado → evidencia y cierre. Nada fuera de §0.2 entra sin Clarification. |
| S-1 newlines (convención única) | **Decidido aquí (G7):** HU-15 mantiene la convención vigente (archivos nuevos en CRLF Windows, sin `.gitattributes`/`.editorconfig`, sin mezclar — verificado en Task 4.3). Crear `.gitattributes` repo-wide es entrega → HU-17 (manual + paquete), no esta HU. |

**Sale porque §9 lo asigna a otras HUs (EXPLÍCITO):** INTERVENTORIA L25:N31 + filas L-Especiales (HU-16); manual + instructivo Capa B + paquete + casos (HU-17); cambios de cálculo/escritura/validación (Q1/Q2 intactos); instalador/setup; DI framework; restyle UI; tocar `Herramientas/VerificadorRecaudo`.

### 0.3 Decisiones ejecutivas (aprobación = aceptar estas)

| ID | Decisión |
|---|---|
| G1 | **Sin T0 de workbook.** Ningún gate nuevo exige semántica de celda no congelada; la paridad compara valores entre dos ejecuciones del mismo código, no contra oráculo nuevo. |
| G2 | **Cero cambios de valores (hierro).** Todo lo que altere un número escrito o un gate existente que pase a fallar contra la red actual es NEEDS_CONTEXT, no "CLI". La red 208/208 + 24/24 debe seguir verde sin tocar goldens. |
| G3 | **PRs encadenados** (§4 Work Units): path compartido (W-2.1/W-2.3/W-2.4 + `CodigoDe`) → proyecto CLI → paridad + regresión. HU-14 vive en working tree sin commit; HU-15 asume su base. Sin commits en esta HU (`#commit` explícito del Ingeniero). |
| G4 | **Parseo manual, sin System.CommandLine** (D2: dictamen con trade-offs en §2.2; re-evaluar solo si HU-16+ exige subcomandos). |
| G5 | **`--salida` es siempre carpeta** (igual que `txtCarpetaSalida` en `Form1:181`): salida = `Path.Combine(salida, periodo.NombreArchivo)`. Sin modo "archivo exacto": la paridad con WinForms es literal. |
| G6 | **Errores de uso (flags) → stderr + ayuda + salida 4; errores de dominio → catálogo.** No se crea ningún código nuevo: el contrato 0-5 queda intacto. Sin args → ayuda + 0 (descubrible, inofensivo). |
| G7 | **Newlines:** CRLF en archivos nuevos, verificado; `.gitattributes` → HU-17 (§0.2). |

---

## 1. PROPOSE

### 1.1 Intent

Darle al Ingeniero la misma ejecución que hoy corre con clicks, pero invocable desde un script: `Remuneracion.Cli --periodo 2026071 --carpeta … --plantilla … --salida … [--ase N|--cinco-ase] [--sobrescribir]` devuelve 0-5 al sistema, deja la misma salida byte-a-valor y el mismo log auditable por `RunId` — y cierra tres cabos sueltos de HU-14 (RunId inyectable, códigos del writer, denegación auditable) en el path que ambos frontends comparten.

### 1.2 In Scope

- Proyecto `Remuneracion.Cli` (`net10.0`, `OutputType Exe`, consola): `Program.Main` + `EjecutorCli.Ejecutar(args) → int` + `OpcionesCli.Parse` puro + composición idéntica a `Program.cs` (mismos servicios, mismo oráculo de lectura en modo período).
- Flags: `--periodo AAAAMM[Q]` (+ `--quincena 1|2` como override/complemento), `--carpeta`, `--plantilla`, `--salida` (carpeta, G5), `--ase N` xor `--cinco-ase` (default 5 ASE), `--sobrescribir`, `--help`/`-h`. Validación de args con mensajes del catálogo HU-14 donde son dominio; uso → stderr + salida 4 (G6).
- `Environment.Exit(codigo)` solo en `Main`, con `Log.CloseAndFlush()` previo (primera y única salida con Exit del producto, V2).
- W-2.1: `Guid? RunId` aditivo en ambas solicitudes; procesadores usan el inyectado o generan (`?? Guid.NewGuid()`); el CLI genera uno y lo pasa (cierra W-2.1).
- W-2.3: códigos explícitos en el writer (`ERR-PLANTILLA` en `ValidarArchivo`/`ValidarRutaSalida`/`ValidarNoInPlace`/estructura de plantilla; `ERR-ESCRITURA` con inner en fallos de I/O durante escritura, preservando atomicidad borrar-parcial).
- W-2.4: `Log.Error` en la denegación salida==plantilla de `Form1` + denegación equivalente del CLI.
- `CatalogoErrores.CodigoDe(Exception)` en Core (código de una excepción sin UI); `Form1.ObtenerCodigoError` pasa a delegar (mecánico, sin cambio de comportamiento).
- Paridad CLI↔WinForms: test que ejecuta ambas vías in-process ante mismos insumos y demuestra igualdad de valores (sin golden nuevo). Regresión 208/208 + 24/24 verde; build 0 warnings incluido el proyecto nuevo; CRLF; sin commits.

### 1.3 Out of Scope

Todo §0.2 (INTERVENTORIA HU-16, manual/paquete HU-17). Además: cambiar valores/gates/tolerancia ±0.5; tocar mapas HU-07..HU-13; nuevo oráculo/golden; `Environment.Exit` en WinForms; DI framework; instalador/setup; `System.CommandLine` u otro paquete (D2); `.gitattributes`; tocar el harness auxiliar.

### 1.4 Resultado de negocio

El Ingeniero puede programar la quincena (`Remuneracion.Cli … --cinco-ase` → `%ERRORLEVEL%` 0-5), encadenarla en scripts (`if %ERRORLEVEL% neq 0 …`), y auditarla por `RunId` en el mismo log; y obtiene la prueba de que el CLI produce exactamente los mismos valores que la UI ante los mismos insumos.

---

## 2. DESIGN

### 2.1 Contrato CLI (tela verificable desde el día uno)

```text
Remuneracion.Cli --periodo 2026071 --carpeta <dir-periodo> --plantilla <xlsx>
                 --salida <dir> [--ase 3 | --cinco-ase] [--quincena 2] [--sobrescribir] [--help]
```

| Flag | Regla | Error si se viola → salida |
|---|---|---|
| `--periodo` | `AAAAMM` (6) o `AAAAMMQ` (7, Q∈{1,2}); `--quincena` opcional: si el período ya trae Q y difiere → uso; si trae solo AAAAMM, `--quincena` es obligatoria (el alcance pide quincena explícita) | Uso → stderr + ayuda + 4 |
| `--carpeta`, `--plantilla` | Requeridos; existencia verificada por los procesadores/locator (dominio, no parseo) | `ERR-FUENTE-NO-ENCONTRADA` → 2 / `ERR-PLANTILLA` → 2 |
| `--salida` | Requerida, siempre carpeta (G5); salida = `Combine(salida, periodo.NombreArchivo)` | Uso si ausente → 4 |
| `--ase N` / `--cinco-ase` | Mutuamente excluyentes; default `--cinco-ase`; N∈1..5 (`AseFactory.DesdeId`, fail-fast nombrado) | Ambos o N inválido → 4 |
| `--sobrescribir` | Sin él y salida existe → denegación no-interactiva (espejo de "No" en el box): `Warning WARN-CANCELADO` + salida 5 | 5 |
| `--help`, `-h` | Imprime uso a stdout, ignora el resto | 0 |
| flag desconocido | Uso a stderr + hint de `--help` | 4 |
| salida == plantilla | Denegación con `Log.Error` (W-2.4) + guía del catálogo a stderr | `ERR-PLANTILLA` → 2 |

**Línea final grepable (stdout):** `RESULTADO OK codigo=0 salida=<ruta> runId=<guid>` o `RESULTADO ERROR codigo=<ERR-…> salida=<N> runId=<guid>`. Hitos a stdout; detalle técnico al log (misma doctrina HU-14 D2/D4).

### 2.2 Decisiones de arquitectura

| ID | Opción elegida | Descartada | Por qué |
|---|---|---|---|
| D1 | Consumir `CodigosSalida`/`CatalogoErrores` tal cual (V1); `Main` = única llamada a `Environment.Exit` del producto, tras `Log.CloseAndFlush()` | `static int Main` sin `Exit` | El alcance lo pide explícito ("con `Environment.Exit`"); el flush previo evita perder el evento final (Exit no corre `finally`) |
| D2 | **Parseo manual** (`OpcionesCli.Parse(string[])`, ~80 líneas, puro y testeable) | `System.CommandLine` | Trade-offs: el paquete suma dependencia + API en estabilización + curva para 6 flags estables sin subcomandos; el manual es 0 paquetes (prohibidos salvo justificación, Plan 14 §1.2), testeable in-memory y suficiente. Re-evaluar solo si HU-16+ exige subcomandos |
| D3 | `EjecutorCli.Ejecutar(args, salida?, errores?) → int` separable de `Main` (exit solo en `Main`) | Lógica en `Main` | Testabilidad: la matriz de códigos y la paridad corren in-process sin subprocesos ni `Exit` en tests |
| D4 | `Guid? RunId` aditivo en `SolicitudProcesoPeriodo`/`SolicitudProcesoAse` (default `null` = genera, compat con 208 tests); procesadores: `solicitud.RunId ?? Guid.NewGuid()` | Parametro extra en `Ejecutar` o `AsyncLocal` propio | Aditivo = cero churn en firmas `IProcesador*` y tests; `LogContext` existente lo propaga (HU-14 D5) |
| D5 | Códigos explícitos en el writer, preservando tipos y atomicidad: `ValidarArchivo`/`ValidarRutaSalida` → `ArchivoFuenteNoEncontradoException(CodigoError.Plantilla, …)`; `ValidarNoInPlace` → mismo tipo+código (**corrige salida 1→2**, V4); estructura de plantilla (hoja/fórmula/valor-fijo) → `CalculoInvalidoException(CodigoError.Plantilla, …)`; `catch` no-codificado en `GenerarWorkbook` → `CalculoInvalidoException(CodigoError.Escritura, …, inner)` tras borrar el parcial | Nueva jerarquía de excepciones o tocar `WorkbookLeafCoherence` (gates → siguen `ERR-VALIDACION`) | Los gates de coherencia SON validación (salida 1, correcto); lo que miente hoy es la fase de escritura con default `ERR-VALIDACION`. El `inner` preserva la causa real para el log |
| D6 | `CatalogoErrores.CodigoDe(Exception)` en Core (los 2 tipos de dominio → `Codigo`, resto → `Inesperado`); `Form1` delega, CLI consume | Duplicar el `switch` en el CLI | Core sigue puro (los tipos son Core); elimina la 3.ª copia del mapeo y lo hace testeable in-memory |
| D7 | Sin runner compartido UI↔CLI: el CLI compone los mismos servicios con ~30 líneas propias de mapeo arg→solicitud; la paridad testea la divergencia | Extraer `EjecutarModo*` de `Form1` a clase compartida | Están acoplados a UI (`txtLog`, `progressBar`, `MessageBox`); extraerlos es churn sobre archivos HU-14 por ~30 líneas. La paridad (§5) es el guardián anti-divergencia, más barata que la extracción |
| D8 | Config Serilog duplicada en el CLI (12 líneas, mismo rolling/template que `Form1.ConfigurarSerilog`, con docstring "canónico en `Form1`") | Helper compartido en Core o Infrastructure | `WriteTo.File` exige `Serilog.Sinks.File`, que Core no referencia por ADR HU-14 (Core = Serilog core para `LogContext`) e Infrastructure no debe ganar por un helper. Divergencia improbable y visible en diff |

### 2.3 Flujo CLI (misma ejecución, otro frontend)

```text
Main(args) → EjecutorCli.Ejecutar → OpcionesCli.Parse (uso → stderr + 4)
  └─► genera RunId → ConfigurarSerilog (D8) → LogContext(RunId, Periodo, Modo)
        └─► salida==plantilla? → Log.Error + guía catálogo (W-2.4) → 2
        └─► existe salida && !--sobrescribir? → Log.Warning WARN-CANCELADO → 5
        └─► solicitud.RunId = runId (W-2.1, D4) → procesador (5 ASE | 1 ASE vía locator, igual que Form1)
              └─► writer con códigos explícitos (W-2.3, D5)
        └─► catch: codigo = CatalogoErrores.CodigoDe(ex) (D6) → Log.Error + RESULTADO ERROR → mapa a 0-5
Main: Log.CloseAndFlush(); Environment.Exit(codigo)  // D1, único Exit del producto
```

Modo 1-ASE resuelve carpeta+R1/R2/R4 con `ArchivoFuenteLocator` exactamente como `Form1.ObtenerCarpetaAse` + `EjecutarModoUnAse` (mismo orden de prefijos, mismo fallback `Reversión/Reversion`).

### 2.4 Dominio (Core, sin deps nuevas)

```csharp
public sealed class SolicitudProcesoPeriodo
{
    // ... existentes intactos ...
    /// HU-15 (W-2.1, D4): RunId inyectado por el frontend; null = el procesador genera uno.
    public Guid? RunId { get; set; }
}
// idem SolicitudProcesoAse.

// HU-15 (D6): código del catálogo para una excepción, sin UI (Form1 delega aquí).
public static string CodigoDe(Exception ex) => ex switch
{
    ArchivoFuenteNoEncontradoException archivo => archivo.Codigo,
    CalculoInvalidoException calculo => calculo.Codigo,
    _ => CodigoError.Inesperado
};
```

Procesadores: `var runId = solicitud.RunId ?? Guid.NewGuid();` (resto idéntico). Excepciones: sin cambios (defaults HU-14 intactos).

### 2.5 Cierre de deuda HU-14 en el path compartido (diseño por ítem)

1. **W-2.1** (V3, D4): propiedad aditiva + `?? Guid.NewGuid()` en ambos procesadores; el CLI genera el suyo al arrancar y lo pasa; `Form1` puede seguir generando el suyo (su `LogContext` UI + el del procesador anidado comparten valor solo si se pasa — se pasa: `Form1` pone su `runId` en la solicitud; unifica la correlación UI→procesador que hoy son dos Guid distintos).
2. **W-2.3** (V4, D5): códigos explícitos listados en D5 + `inner` en `ERR-ESCRITURA`; la atomicidad (borrar parcial) se conserva intacta; los gates de coherencia NO se tocan (siguen `ERR-VALIDACION` → 1).
3. **W-2.4** (V5): `Log.Error("[{Codigo}] …", CodigoError.Plantilla, …)` en el early-return de `Form1:196-205` (una línea + contexto; el box/status quedan igual) y en la denegación equivalente del CLI (stderr + log).

### 2.6 Proyecto consola (integración al `.slnx`)

```xml
<!-- Remuneracion.Cli/Remuneracion.Cli.csproj -->
<Project Sdk="Microsoft.NET.Sdk">
  <PropertyGroup>
    <OutputType>Exe</OutputType>
    <TargetFramework>net10.0</TargetFramework>   <!-- consola, NO -windows: testeable y desatendida -->
    <Nullable>enable</Nullable>
    <ImplicitUsings>enable</ImplicitUsings>
    <RootNamespace>Remuneracion.Cli</RootNamespace>
  </PropertyGroup>
  <ItemGroup>
    <ProjectReference Include="..\Remuneracion.Core\Remuneracion.Core.csproj" />
    <ProjectReference Include="..\Remuneracion.Infrastructure\Remuneracion.Infrastructure.csproj" />
  </ItemGroup>
  <ItemGroup>
    <!-- Mismas versiones del árbol (D2: sin paquetes nuevos) -->
    <PackageReference Include="Serilog" Version="4.4.0" />
    <PackageReference Include="Serilog.Sinks.File" Version="7.0.0" />
  </ItemGroup>
</Project>
```

`.slnx`: 5.ª línea `<Project Path="..\Remuneracion.Cli\Remuneracion.Cli.csproj" />`. Compila y se empaqueta junto al WinForms con el build existente (`dotnet build Remuneracion.WinForms/Remuneracion.WinForms.slnx`); ejecución: `dotnet run --project Remuneracion.Cli -- --help` (el WinForms sigue `WinExe` sin consola). Sin instalador/setup (out).

Archivos nuevos: `Remuneracion.Cli/Program.cs` (Main + wiring idéntico a `Program.cs` V8), `Remuneracion.Cli/EjecutorCli.cs`, `Remuneracion.Cli/OpcionesCli.cs` (Parse puro + uso), `Remuneracion.Cli/ConfigurarSerilog` (privado en `Program`, D8).

### 2.7 Paridad y regresión (sin golden nuevo)

Paridad = **mismos valores, no nuevo oráculo**: ante mismos insumos, la vía CLI in-process (`EjecutorCli.Ejecutar`, sin `Exit`) y la vía directa (procesador, como la usa `Form1`) producen iguales `Resultado` (GranTotal + totales por ASE) y workbooks con iguales valores de celda (comparados vía `ValidacionOracleReader`, no por bytes — los metadatos OpenXML llevan timestamps). Casos: 5-ASE Q1 una vez (camino principal) + 1-ASE Q1/Q2 (cubre AJUSTES-SF-T/DetRetri en Q2). Regresión 208/208 + 24/24 intacta + suites nuevas (§5).

### 2.8 File changes previstos

| Archivo | Acción | Motivo |
|---|---|---|
| `Remuneracion.Cli/*.csproj, Program.cs, EjecutorCli.cs, OpcionesCli.cs` | Crear (+ línea en `.slnx`) | Proyecto consola §2.6 |
| `Remuneracion.Core/Models/SolicitudProceso{Periodo,Ase}.cs` | Modificar (aditivo) | `Guid? RunId` (W-2.1, D4) |
| `Remuneracion.Core/Services/Procesador{Periodo,Remuneracion}.cs` | Modificar (1 línea c/u + docstring) | Respetar `RunId` inyectado (W-2.1) |
| `Remuneracion.Core/Errors/CatalogoErrores.cs` | Modificar (aditivo) | `CodigoDe(Exception)` (D6) |
| `Remuneracion.Infrastructure/Excel/OpenXmlPlantillaWriter.cs` | Modificar | Códigos explícitos (W-2.3, D5); atomicidad intacta |
| `Remuneracion.WinForms/Form1.cs` | Modificar (2 puntos) | `Log.Error` en salida==plantilla (W-2.4); delegar a `CodigoDe` (D6); pasar `runId` en solicitudes (W-2.1) |
| `Remuneracion.IntegrationTests/OpcionesCliTests.cs` | Crear | Parse: matriz de flags, defaults, exclusión, errores de uso |
| `Remuneracion.IntegrationTests/CodigosSalidaCliTests.cs` | Crear | Matriz código→salida in-process (2,1,3,4,5 + 0 + help) |
| `Remuneracion.IntegrationTests/ParidadCliTests.cs` | Crear | Igualdad de valores CLI↔directo (§2.7) + RunId respetado |
| `Remuneracion.IntegrationTests/EscrituraCodigosTests.cs` | Crear | W-2.3: plantilla ausente/in-place/estructura→`ERR-PLANTILLA`; I/O→`ERR-ESCRITURA` con inner |

**No tocar:** cálculo/validación/mapas/goldens/tolerancia; `Program.cs` WinForms; harness auxiliar; `requirements/` legado.

---

## 3. SPEC / Acceptance Criteria (trazabilidad Propuesta §9–§10)

### Requirement 1 — CLI ejecuta desatendido con args (3.3)

El binario **MUST** aceptar §2.1 y ejecutar ambos modos sin intervención; **MUST NOT** pedir input interactivo nunca (sin salida existente + sin `--sobrescribir` = denegación 5, no prompt).

- GIVEN `--periodo 2026071 --carpeta <P> --plantilla <T> --salida <S> --cinco-ase --sobrescribir` válido → WHEN `Ejecutar` → THEN `0` + workbook en `S/Remuneración 202607-1 Total.xlsx` + `RESULTADO OK … runId=…`.
- GIVEN sin `--ase` ni `--cinco-ase` → THEN modo 5 ASE (default documentado).
- GIVEN `--ase 3` → THEN solo ASE 3 (mismas fuentes que `Form1.EjecutarModoUnAse`).

### Requirement 2 — `--help` y validación de args (3.3; catálogo HU-14)

`--help` **MUST** imprimir uso y devolver 0; flag desconocido / período malformado / `--ase` fuera de 1..5 / `--ase`+`--cinco-ase` juntos / `--quincena` contradictoria **MUST** ir a stderr + ayuda + salida 4 (G6). **MUST NOT** crear códigos nuevos.

- GIVEN `--help` → THEN uso a stdout + `0`.
- GIVEN `--periodo 2026` → THEN stderr con uso + `4`.
- GIVEN `--periodo 202607 --quincena 3` → THEN `4`; GIVEN `--periodo 202607` sin `--quincena` → THEN `4` (quincena obligatoria si el período no la trae).

### Requirement 3 — Códigos de salida 0-5 con `Environment.Exit` (3.3; contrato HU-14)

`Main` **MUST** ser el único `Environment.Exit` del producto (V2), tras `Log.CloseAndFlush()` (D1); `EjecutorCli` **MUST** retornar el `int` sin salir (D3). **MUST NOT** aparecer `Exit` en WinForms/Core/Infrastructure/tests.

- GIVEN R2 del ASE 3 ausente → THEN `2` + stderr `[ERR-FUENTE-NO-ENCONTRADA] … ASE 3 …`.
- GIVEN gate que no cierra → THEN `1`; GIVEN fallo de escritura I/O → THEN `3`; GIVEN excepción genérica → THEN `4`; GIVEN salida existente sin `--sobrescribir` → THEN `5` + `Warning WARN-CANCELADO`.
- GIVEN grep `Environment\.Exit` en el producto → THEN solo `Remuneracion.Cli/Program.cs`.

### Requirement 4 — RunId único respetado (W-2.1; §6 trazabilidad)

El `RunId` generado por el CLI **MUST** ser el que correla todos los eventos (UI→procesador→writer) y **MUST** imprimirse en la línea final; el procesador **MUST** usar `solicitud.RunId` cuando viene (D4). **MUST NOT** generar un segundo Guid en el procesador cuando se inyecta.

- GIVEN ejecución CLI → WHEN filtrar el log por el `runId` impreso → THEN todos sus eventos (inicio→GranTotal o Error).
- GIVEN `solicitud.RunId = G` fijo → THEN el evento de inicio del procesador porta `G` (sink en memoria, patrón `ObservabilidadTests`).

### Requirement 5 — Códigos del writer (W-2.3; 3.1 en path compartido)

Cada fallo del writer **MUST** portar el código explícito D5 con mensaje que nombra plantilla/salida/hoja-celda; el parcial **MUST** seguir borrándose (atomicidad intacta). **MUST NOT** quedar ningún `throw` del writer con default `ERR-VALIDACION` que sea problema de plantilla/escritura.

- GIVEN plantilla ausente → THEN `ERR-PLANTILLA` (→2); GIVEN salida==plantilla en writer → THEN `ERR-PLANTILLA` (antes `ERR-VALIDACION`→1, V4).
- GIVEN hoja sin WorkbookPart / fórmula ausente / valor fijo en plantilla → THEN `ERR-PLANTILLA`.
- GIVEN `IOException` durante copia/escritura → THEN `ERR-ESCRITURA` con `InnerException` preservada (→3) y sin archivo parcial.

### Requirement 6 — Denegación auditable (W-2.4; CA-6)

La denegación salida==plantilla **MUST** emitir `Log.Error` con el código en ambas superficies (UI + CLI). **MUST NOT** cambiar el comportamiento visible (mismo box/status en UI; misma salida 2 en CLI).

- GIVEN salida==plantilla en UI → THEN `Log.Error` con `ERR-PLANTILLA` (además del box actual).
- GIVEN salida==plantilla en CLI → THEN stderr con guía del catálogo + `Log.Error` + `2`.

### Requirement 7 — Paridad CLI↔WinForms (3.3; CA-2/CA-6)

Ante mismos insumos, ambas vías **MUST** producir iguales valores (GranTotal + 5 totales + celdas-oráculo iguales vía reader). **MUST NOT** crear golden/oráculo nuevo ni comparar bytes.

- GIVEN insumos Q1 5 ASE → THEN `Resultado` CLI == directo (tolerancia exacta, mismos decimales) y snapshots-oráculo de ambas salidas iguales.
- GIVEN Q2 1-ASE → THEN igualdad incluyendo `AjustesSfT`/`DetRetriQ2` (cubre el path Q2 sin golden nuevo).

| CA §10 | HU-15 |
|---|---|
| CA-2 | Paridad demuestra mismos valores (tolerancia exacta entre vías, no ±0.5 contra manual) |
| CA-6 | RunId + niveles HU-14 en CLI; cada ejecución auditable por `runId` impreso |
| CA-5/CA-7 | Mismo flujo 5-ASE/1-ASE + UX por catálogo a stderr; cero cambios de cálculo |

---

## 4. TASKS

## Review Workload Forecast

| Field | Value |
|---|---|
| Estimated changed lines | 600–900 (CLI ~350 nuevo + path compartido ~100 + tests ~300) |
| 400-line budget risk | Medium-High |
| Chained PRs recommended | Yes |
| Suggested split | PR1 path compartido (W-2.1/W-2.3/W-2.4 + CodigoDe) → PR2 proyecto CLI → PR3 tests paridad + regresión |
| Delivery strategy | ask-on-risk |
| Chain strategy | pending |

Decision needed before apply: Yes. Chained PRs recommended: Yes. Chain strategy: pending.

### Suggested Work Units

| Unit | Goal | Likely PR | Notes |
|---|---|---|---|
| 0 | Pre-vuelo: baseline + reclamo S-2..S-5 | — | Gate: sin verde no se empieza; suggestions fuera de §0.2 → HU-17 |
| 1 | Path compartido (Core + writer + Form1 2 puntos) | PR 1 | Sin deps; tests in-memory/temp desde el día uno |
| 2 | Proyecto CLI en el slnx (parse + ejecutor + Main) | PR 2 | Depende de PR 1 (CodigoDe, RunId, writer) |
| 3 | Paridad + matriz de salidas + regresión total | PR 3 | Depende de PR 1–2 |

### Phase 0 — Pre-vuelo (gate)

- [ ] 0.1 Baseline: `dotnet build Remuneracion.WinForms/Remuneracion.WinForms.slnx` 0 warnings + suites 208/208 + harness 24/24 verdes (HU-14 en working tree asumida como base; si algo falla → NEEDS_CONTEXT, no se empieza).
- [ ] 0.2 Reclamar al orquestador las suggestions S-2..S-5 de la revisión HU-14 (V10): verificar una por una contra código; entra solo lo que afecte corrección del CLI/path compartido; cosmética → HU-17; saldada → evidencia. Registrar veredicto en §9 antes de PR1.

### Phase 1 — Path compartido (W-2.1 / W-2.3 / W-2.4 + CodigoDe)

- [ ] 1.1 `Guid? RunId` en ambas solicitudes (D4) + `solicitud.RunId ?? Guid.NewGuid()` en ambos procesadores + `Form1` pasa su `runId` (unifica correlación UI→procesador, §2.5.1).
- [ ] 1.2 `CatalogoErrores.CodigoDe` (D6) + `Form1.ObtenerCodigoError` delega (mecánico) + casos in-memory (2 tipos con/sin código explícito, genérica→Inesperado, código desconocido→Inesperado).
- [ ] 1.3 W-2.3 writer (D5): códigos explícitos + `ERR-ESCRITURA` con inner + atomicidad intacta + `EscrituraCodigosTests` (Req 5).
- [ ] 1.4 W-2.4: `Log.Error` en `Form1:196-205` (una línea, box/status intactos).

### Phase 2 — Proyecto CLI

- [ ] 2.1 `Remuneracion.Cli.csproj` (§2.6) + línea en `.slnx` + build 0 warnings del slnx completo.
- [ ] 2.2 `OpcionesCli.Parse` puro (tabla §2.1: período/quincena, exclusión de modo, `--help`, flag desconocido) + `OpcionesCliTests` (Req 1–2; sin I/O).
- [ ] 2.3 `EjecutorCli` (§2.3: RunId, Serilog D8, denegaciones W-2.4/5, composición V8, línea final grepable, catch con `CodigoDe`) + `Program.Main` (wiring + flush + único `Exit`, D1).
- [ ] 2.4 `CodigosSalidaCliTests` in-process (Req 3: matriz 0/1/2/3/4/5 + help; asserts de stdout/stderr, nunca subproceso).

### Phase 3 — Paridad y regresión

- [ ] 3.1 `ParidadCliTests` (Req 7: 5-ASE Q1 + 1-ASE Q1/Q2, igualdad exacta de valores + snapshots-oráculo iguales; asserts de `runId` impreso == correlación del log).
- [ ] 3.2 Regresión total 208/208 + 24/24 + build 0 warnings (incluido CLI) + CRLF verificado en archivos nuevos (G7; sin `.gitattributes`).
- [ ] 3.3 Grep final: `Environment\.Exit` solo en `Remuneracion.Cli/Program.cs` (producto); cálculo/validación/mapas/goldens intactos (diff revisado).

---

## 5. Test Strategy

| Capa | Qué | Cómo |
|---|---|---|
| Unidad | `OpcionesCli.Parse` (matriz §2.1: 6 flags, defaults, exclusiones, 5 usos inválidos) | Puro, sin I/O |
| Unidad | `CodigoDe` (dominio con/sin código, genérica, desconocido) + `CodigoSalidaPara` intacto | In-memory |
| Unidad/Integ | RunId respetado (inyectado fijo → evento con ese Guid; null → Guid generado ≠ entre corridas) | Sink en memoria (patrón `ObservabilidadTests`, sin paquetes) |
| Integración | W-2.3 (temp: plantilla ausente/in-place/estructura corrupta/I-O bloqueada) | Copias en temp; `Docs/Insumos/` jamás destino |
| Integración | CLI in-process (matriz 0-5 + help; stdout/stderr capturados) | Temp dirs + fixtures existentes |
| Integración | Paridad CLI↔directo (valores exactos + snapshots-oráculo) | Goldens read-only + salidas en temp |
| Regresión | 208/208 + harness 24/24 sin cambios | Suites existentes intactas |
| Manual | `%ERRORLEVEL%` real tras `dotnet run --project Remuneracion.Cli -- …` (único `Exit` verificable solo fuera de proceso) | Un comando por código 0/2/5 |

### 5.1 Casos negativos obligatorios (nombran código + salida)

`--periodo 2026` (uso→4); `--ase 6` (uso→4); `--ase 2 --cinco-ase` (uso→4); `--periodo 202607` sin `--quincena` (uso→4); carpeta sin R2 del ASE 3 (`ERR-FUENTE-NO-ENCONTRADA`→2); salida==plantilla (`ERR-PLANTILLA`→2 + `Log.Error`); plantilla corrupta sin hoja (`ERR-PLANTILLA`→2); salida bloqueada I/O (`ERR-ESCRITURA`→3 + inner); gate roto (`ERR-VALIDACION`→1); salida existente sin `--sobrescribir` (`WARN-CANCELADO`→5).

---

## 6. Architecture Validation Certificate

### 6.1 SOLID

| Principio | Veredicto |
|---|---|
| **S** | El CLI es un frontend más (parse + composición + exit); cálculo/validación/escritura conservan su rol HU-01..HU-14. W-2.3 fija códigos, no mueve responsabilidades. |
| **O** | Solicitudes extendidas por propiedad nullable (firmas intactas); writer extendido por código, no por tipos; `CodigoDe` aditivo. |
| **L** | Constructores y firmas `IProcesador*` intactos: todo consumidor actual compila y se comporta igual (`RunId null` = HU-14 puro). |
| **I** | `OpcionesCli`/`EjecutorCli` consumen contratos existentes; ningún contrato cambia de forma. |
| **D** | Core define códigos/solicitudes; CLI compone Core+Infrastructure como `Program.cs`. Cero paquetes nuevos. |

### 6.2 Best Practices

- El contrato manda: 0-5 intactos, ningún código nuevo (G6); el CLI consume, no redefine (V1).
- Nunca 0 silencioso en gates (doctrina HU-13/14): la paridad compara valores exactos entre vías, no contra oráculo nuevo.
- Un solo `Exit` en el producto (V2, D1) con flush previo; tests siempre in-process (D3).
- Desatendido honesto: sin prompt posible; la negativa a sobrescribir es salida 5, no espera infinita (Req 1).
- La verdad del repo manda: S-1/S-2/S-3/S-4 saldadas con evidencia (V6/V7), no re-trabajo; S-2..S-5.unknown tras Task 0; newlines por convención vigente (G7).

### 6.3 Performance

Parseo O(args) + una ejecución idéntica a la UI; un `Guid` y un flush extra. Irrelevante a esta escala; `Task.Run` no aplica (consola sincrónica; los procesadores son síncronos).

**Veredicto:** APROBADO como it. 3.3 con absorción W-2.1/W-2.3/W-2.4 **si** se respeta G2 (cualquier cambio de valor/gate = NEEDS_CONTEXT) y G4 (sin `System.CommandLine` sin re-dictamen).

---

## 7. Riesgos

| Riesgo | Likelihood | Mitigación |
|---|---|---|
| HU-14 sin commit: la base se mueve bajo los pies (rebase/commit parcial cambia líneas citadas) | Media | Phase 0 fija el baseline con hash; si la base cambia, re-verificar V1..V9 antes del PR1 |
| S-2..S-5 unknown traen alcance escondido (quieren entrar a HU-15) | Media | Task 0 con gate explícito; nada fuera de §0.2 sin Clarification |
| Duplicación Serilog/mapeo arg→solicitud diverge de `Form1` con el tiempo | Baja | D8 documenta el canónico; la paridad (§2.7) detecta divergencia de valores en cada run |
| `ERR-ESCRITURA` con inner rompe un test existente que espera el tipo/mensaje crudo del writer | Baja | El tipo se conserva (`CalculoInvalidoException`); solo se suma `Codigo`+inner; la regresión 208 lo revela gratis en PR1 |
| Salida Q1 5-ASE en temp para paridad es lenta (42 hitos + I/O real) | Baja | Una sola corrida 5-ASE; el resto 1-ASE; aceptado por única vez (no es suite de ciclo corto) |
| Scripts que parsean `%ERRORLEVEL%` 4 como "inesperado" cuando fue error de uso | Baja | G6 documentado + línea `RESULTADO ERROR` con el detalle; el manual HU-17 lo fija para operadores |
| Inflar a HU-16/HU-17 (INTERVENTORIA/manual/setup) dentro de esta HU | Media | §0.2 out explícito; rechazar PRs con hojas nuevas, docs de usuario o setup |

---

## 8. Rollback

- Revertir PRs en orden inverso (paridad → CLI → path compartido); el producto sin esta HU = UI actual + HU-14 intacta (solicitudes sin `RunId` = generan; writer con defaults; `Form1` sin `Log.Error` extra).
- Eliminar `Remuneracion.Cli/` + línea del `.slnx` deja el build idéntico al actual.
- `Docs/Insumos/` untracked: jamás como destino; pisada se restaura desde backup.
- Si W-2.3 revela un `throw` del writer que la regresión esperaba crudo, recorte al mínimo (código explícito solo donde el test lo permite), nunca silenciamiento.

---

## 9. Decisiones fijadas por este plan

1. Rector = Propuesta §9 it. 3.3 + §6 + §10 CA-6. Entran 3.3 + W-2.1/W-2.3/W-2.4; HU-16, HU-17 y cambios de cálculo/escritura quedan fuera.
2. Deuda HU-14 verificada contra código (§0.1): W-2.1/W-2.3/W-2.4 vigentes con evidencia; S-1-"false" y S-2/S-3/S-4 saldadas con evidencia (no replanificar); S-2..S-5 de review tras Task 0; newlines por G7 (CRLF ahora, `.gitattributes` en HU-17).
3. Sin T0 de workbook (G1): paridad = mismos valores entre vías, sin golden nuevo.
4. G2 (hierro): cero cambios de valores; lo que rompa la red actual = NEEDS_CONTEXT.
5. Parseo manual con dictamen (G4/D2); `System.CommandLine` solo tras re-dictamen si hay subcomandos.
6. `--salida` siempre carpeta (G5); errores de uso → 4 sin códigos nuevos (G6); sin args → ayuda + 0.
7. `Environment.Exit` solo en `Main` del CLI con flush previo (D1); tests in-process (D3); sin runner compartido UI↔CLI (D7); Serilog duplicado documentado (D8).
8. PRs encadenados 0→3 (§4); regresión 208/208 + 24/24 como red; build 0 warnings; CRLF; sin commits.

**Listo para aprobación del Ingeniero. No implementar hasta OK.**
