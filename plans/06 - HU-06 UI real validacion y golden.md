# Plan 06 — HU-06: UI real, validación básica y golden sobre la ruta leaf

> **Historia:** cerrar el prototipo Fase 1 (1 ASE) con orquestación real, validaciones básicas y golden honesto.
> **Fuente rectora:** `Docs/Propuesta_Proyecto_Automatizacion_Remuneracion_UAESP.md` (§4.1, §4.2, §6, §9 it. 1.8–1.10, §10 CA-1..CA-7).
> **Continuidad:** `plans/05 - HU-05 WorkbookLeafInputs y escritura real.md` (cerrada). HU-04 validation-only **intacta**.
> **No rector de alcance:** `requirements/Fase1-Requerimientos.md` HU-05-legado (8 archivos, `VALIDACION_TOTAL`/`VALIDACION_RECIP`, `DetRetri` como objetivo de escritura). Eso es **Fase 2 / HU posteriores**. Citar §4.2 al excluirlo.
> **Estado:** PENDIENTE DE APROBACIÓN — no implementar hasta OK del Ingeniero.
> **Fecha:** 2026-09-02

---

## 0. Clarification Gate

No hay pregunta bloqueante. Decisiones resueltas abajo; el apply espera aprobación.

### 0.1 Por qué esta HU existe ahora

| Hecho verificado | Consecuencia |
|---|---|
| HU-01..HU-05 cerradas | Lectura, cálculo, leaf reader y `GenerarWorkbook` existen. |
| `Form1.btnEjecutar_Click` sigue en stub (delay 2s) | La app **no automatiza** el proceso. |
| `Program.cs` = `new Form1()` | No hay composition root. |
| `IValidador` sin implementación | Falta el servicio de validación básica de §6 / it. 1.8. |
| OpenXML **no recalcula** fórmulas | Comparar cache de `D9` del archivo generado vs golden **es evidencia falsa**. |

### 0.2 Mapeo a la Propuesta (in vs out)

**Entra porque §4.1 / §9 / §10 lo piden ahora:**

| Rector | Qué cubre HU-06 |
|---|---|
| §4.1 UI | Carpetas, progreso, resultados — **ejecución real**, no skeleton. |
| §4.1 Validaciones básicas | Coherencia entre hojas + fórmulas preservadas. |
| §9 it. 1.8 | Validaciones básicas. |
| §9 it. 1.9 | UI WinForms de selección, progreso, resultados, log. |
| §9 it. 1.10 | Testing contra resultado conocido (golden **honesto**). |
| §10 CA-1..CA-7 | Se ejercitan por la ruta leaf. CA-3 **parcial y declarado**. |
| §6 | WinForms → Infrastructure → Core; Serilog; no mutar fórmulas. |

**Sale porque §4.2 lo saca de Fase 1:**

- Los otros 4 ASE.
- `SALDOS POR NOTA`, `RETRIBUCION NEGATIVA`, `AJUSTES-SF-T`.
- `BCE SC POR FACT`, `INTERVENTORIA`, `REPORTE RECAUDO x BANCO`.
- `VALIDACION_TOTAL` / `VALIDACION_RECIP`.
- Generación automática de `DetRetri` (enteros como objetivo de escritura).

El HU-05-legado de `requirements/Fase1-Requerimientos.md` **infla** Fase 1 (8 fuentes, CA-3 con D85:D89, CA-8 DetRetri). **No es el alcance de esta HU.**

### 0.3 Verdad del workbook que HU-06 NO puede desaprender

Promoambiental 202607-1:

| Cantidad | Valor | Origen |
|---|---|---|
| `R1!F46` / `CONSOLIDADO!D9` | `16704332434.57` | `F25+F41-L25` — **no** HU-02 `TotOpt` `19556118465.99` |
| `R1!F48` / `CONSOLIDADO!D47` | `11673020` | `F30+F10-L10` — **no** HU-02 `Extemp` |
| HU-02 `Extemp` | `19549786950.62` | Leaf `F25` |
| `R2!E41` / `D28` | `54216385.68` | `E15+E26-K15` = HU-02 R2 TotalOportuno |
| `R4!D67` / `D66` | `-12054255.65` | `D9-P9` = HU-02 TotalReversiones |

Nunca sobrescribir fórmulas. Nunca mutar la plantilla in-place.

### 0.4 Decisiones ejecutivas (aprobación = aceptar estas)

| ID | Decisión |
|---|---|
| G1 | **Golden automatizado ≠ CA-3 cache.** Ver §2.4. |
| G2 | **Salida = carpeta + `Periodo.NombreArchivo`.** Ver §2.3. |
| G3 | **Orquestador en Core; Form1 flaco.** Ver §2.2. |
| G4 | **`IValidador` cubre lo que falta en dominio; no reabre HU-04/HU-05.** Ver §2.5. |
| G5 | **UI delta mínimo:** una fila de salida + resumen en `txtLog`. Sin restyle. |
| G6 | Recálculo Excel = protocolo **manual/opcional**, no CI. |

---

## 1. PROPOSE

### 1.1 Intent

Reemplazar el stub de WinForms por el flujo real de 1 ASE (Promoambiental Q1 como fixture) y certificar esa ruta con validación básica + golden **sin fingir recálculo Excel**.

### 1.2 In Scope

- Composition root en `Program.cs` (new-up, **sin** contenedor DI).
- Servicio de aplicación en Core: leer R1/R2/R4 → calcular → leaf → validar → `GenerarWorkbook`.
- Implementar `IValidador` (validaciones básicas de dominio).
- UI: destino de salida, orquestación real, progreso por pasos, log Serilog + `txtLog`, resumen de resultados.
- Fail-fast con excepciones existentes.
- Golden automatizado: leaf escritos vs golden; visibles de **dominio** vs cache golden; fórmulas intactas; hash de plantilla origen.
- Protocolo manual documentado para CA-3 post-Excel.

### 1.3 Out of Scope

Todo §4.2. Además:

- Reabrir semántica de `IPlantillaWriter` (HU-04) o de `IWorkbookLeafWriter` (HU-05).
- Motor Excel / COM / interop en CI.
- Multi-ASE (el combo existe; **se procesa solo el ASE seleccionado**).
- Restyle / tokens / colores nuevos.
- Framework DI.
- Ampliar `ValidarContraResultado` a matching multi-ASE.

### 1.4 Resultado de negocio

El Ingeniero elige período, carpeta `REMUNERACION 2026071`, plantilla, ASE Promoambiental y carpeta de salida; pulsa **Ejecutar**; obtiene un `.xlsx` copia con leaf escritos; el log audita cada paso; las pruebas demuestran leaf+visibles de dominio contra golden. CA-3 de celdas fórmula **después de Excel** queda explícito como residual.

---

## 2. DESIGN

### 2.1 Decisiones de arquitectura

| ID | Opción elegida | Descartada | Por qué |
|---|---|---|---|
| D1 | `IProcesadorRemuneracion` + `ProcesadorRemuneracion` en Core | God-object `Form1` | §6.1 migrabilidad: el caso de uso no vive en el formulario. |
| D2 | Rutas R1/R2/R4 **ya resueltas** entran al procesador | `IArchivoFuenteLocator` en Core | El locator ya existe en Infrastructure; no inflar contratos. WinForms resuelve con `ArchivoFuenteLocator`. |
| D3 | Composition root manual en `Program.cs` | Autofac/MS.DI | Tres implementaciones concretas; un contenedor no se justifica. |
| D4 | Carpeta de salida + `Periodo.NombreArchivo` | SaveFileDialog de archivo | §5.3 ya nombra el archivo; el writer **prohíbe** in-place; el patrón visual ya es “carpeta + examinar”. |
| D5 | `IValidador` dominio + overload con leaf | Reparsear OpenXML en el validador | Fórmulas las certifica el writer (HU-04/HU-05). El validador no duplica I/O. |
| D6 | Golden en dos capas (auto vs Excel) | “Output D9 cache = golden D9” | OpenXML no calcula. Si la plantilla **es** el golden, esa igualdad es **falso positivo**. |
| D7 | UI: una fila en `grpArchivos` + resumen en log | Rediseño / grid de resultados | OPA = Ejecutar. Densidad Balanced. Designer actual es la paleta. |

### 2.2 Orquestación (Form1 flaco)

```text
Form1  →  ArchivoFuenteLocator (solo resolución de rutas)
       →  IProcesadorRemuneracion.Ejecutar(solicitud, IProgress<string>)
                ├─ IRecaudoReader.LeerR1/R2/R4
                ├─ ICalculoRemuneracion.CalcularConsolidado   (1 tupla)
                ├─ IWorkbookLeafInputReader.LeerLeafInputs
                ├─ IValidador.Validar(resultado, leaf)        // vacío ⇒ OK
                └─ IWorkbookLeafWriter.GenerarWorkbook        // copia + leaf
```

Core **no** referencia WinForms. Serilog queda en el form: cada `IProgress<string>` se escribe en `txtLog` **y** `Log.Information`.

Contrato propuesto (nombres ajustables, semántica no):

```csharp
public sealed class SolicitudProcesoAse
{
    public Ase Ase { get; set; } = new();
    public Periodo Periodo { get; set; } = new();
    public string RutaR1 { get; set; } = "";
    public string RutaR2 { get; set; } = "";
    public string RutaR4 { get; set; } = "";
    public string RutaPlantilla { get; set; } = "";
    public string RutaSalida { get; set; } = "";
}

public sealed class ResultadoProcesoAse
{
    public ResultadoRemuneracion Resultado { get; set; } = new();
    public WorkbookLeafInputs Leaf { get; set; } = new();
    public string RutaSalida { get; set; } = "";
}

public interface IProcesadorRemuneracion
{
    ResultadoProcesoAse Ejecutar(SolicitudProcesoAse solicitud, IProgress<string>? progreso = null);
}
```

`Form1` solo: validar controles, resolver carpeta ASE + 3 prefijos, armar `RutaSalida`, `Task.Run` del procesador, marshal al UI, fail-fast visual.

Prefijos de locator (Propuesta §5.1): `Recaudoporcomponente`, `RerpoteDetalleSaldosaFavor`, `ReversiónPorComponente`. Falta cualquiera → `ArchivoFuenteNoEncontradoException`.

### 2.3 Destino de salida

GAP actual: el form no tiene output y `GenerarWorkbook` aborta si origen == destino.

| Control nuevo | Igual que | Comportamiento |
|---|---|---|
| `txtCarpetaSalida` (ReadOnly) + `btnSeleccionarSalida` | carpeta fuentes | `FolderBrowserDialog` |
| Nombre de archivo | no editable | `Periodo.NombreArchivo` → `Remuneración {AAAAMM}-{Q} Total.xlsx` |

Reglas:

1. Ejecutar exige carpeta fuentes, plantilla **y** carpeta salida.
2. `RutaSalida = Path.Combine(carpetaSalida, periodo.NombreArchivo)`.
3. Si `RutaSalida` equivale a la plantilla → no ejecutar (el writer también lo impide).
4. Si el archivo existe → `MessageBox` Sí/No de overwrite; No aborta sin llamar al procesador.
5. La plantilla en `Docs/Insumos/` **nunca** se pisa.

### 2.4 Golden — honestidad CA-3

**Hecho:** OpenXML no recalcula. Tras escribir leaf, el cache de `F46`/`F48`/`E41`/`D67`/`D9`/`D28`/`D47`/`D66` sigue siendo el de la **plantilla copiada** hasta que Excel abre el archivo.

**Trampa:** HU-05 usa `Docs/Insumos/Remuneracion 202607-1 Total.xlsx` como plantilla. Ese archivo **es el golden diligenciado**. Copiarlo y comparar cache de `D9` vs golden `D9` **siempre pasa**, aunque los leaf estén mal.

#### Capa A — automatizada (in-repo, sin Excel) — aceptación de HU-06

| # | Qué | Contra qué | Tolerancia |
|---|---|---|---|
| A1 | Leaf escritos en la **salida**: R1 `F25,F41,L25,F30,F10,L10`; R2 `E15,E26,K15`; R4 `D9,P9` | Mismas celdas **leaf** del golden | ±0.5 |
| A2 | Visibles de **dominio** `F25+F41-L25`, `F30+F10-L10`, `E15+E26-K15`, `D9-P9` | Cache golden `F46`,`F48`,`E41`,`D67` y `D9`,`D47`,`D28`,`D66` | ±0.5 |
| A3 | Fórmulas protegidas del cell-map | Siguen siendo fórmula en la salida | n/a |
| A4 | SHA256 de la plantilla origen | Igual antes/después | n/a |
| A5 | `TotOpt`/`Extemp` HU-02 **no** se comparan con `F46`/`F48` | — | prohibido |

#### Capa B — fuera de CI

Abrir la salida en Excel, recalcular, comparar cache CONSOLIDADO vs golden. Protocolo en §5.3. **No** inventar motor Excel.

**Veredicto CA-3:** HU-06 certifica la **ruta leaf + aritmética de visibles**. La prueba de celdas fórmula recalculadas queda residual documentada (manual o HU siguiente). Afirmar CA-3 completo sería humo.

### 2.5 `IValidador`: qué es “básico” **ahora**

Ya cubierto (no reimplementar, no reabrir):

| Capacidad | Dónde |
|---|---|
| Fórmulas protegidas | HU-04 `IPlantillaWriter` + HU-05 post-write |
| Gate leaf vs agregado (`F25` vs Extemp; R2/R4 vs agregados) | `WorkbookLeafCoherence` en reader/writer |
| No in-place | `GenerarWorkbook` |

**Falta (esta HU)** — implementar `IValidador` en Core:

Mantener `Validar(ResultadoRemuneracion) → List<string>`. Agregar overload `Validar(ResultadoRemuneracion, WorkbookLeafInputs)`.

El overload **MUST** fallar (mensaje en la lista; el procesador trata lista no vacía como bloqueo) si:

1. No hay exactamente 1 consolidado (Fase 1).
2. `AjustesSfT != 0` (Fase 1 / Q1; AJUSTES está en §4.2).
3. `TotalAse` no cuadra con la suma de componentes (el getter ya suma; se valida coherencia de inputs: `TotOpt`/`R2`/`Extemp`/`ReversionR4` presentes).
4. `GranTotal` ≠ único `TotalAse` (±0.5).
5. Gate de dominio (misma regla HU-05, **sin** TotOpt vs F46):
   - `leaf.R1.F25` vs `consolidado.Extemp`
   - `leaf.R2.TotalOportunoEsperado` vs `consolidado.R2TotalOportuno`
   - `leaf.R4.TotalReversionEsperada` vs `consolidado.ReversionR4`

El writer **conserva** su gate interno (defensa). El validador no abre `.xlsx`.

### 2.6 UI — Visual Design Intent

- Densidad: **Balanced**. GroupBoxes actuales.
- OPA: `btnEjecutar`.
- Delta: tercera fila en `grpArchivos` (“Carpeta salida:”) con los **mismos** tamaños que fuentes (`420×23`, botón `110×28`). Empujar `grpEjecucion` / `txtLog` / altura del form lo mínimo.
- Resultados: bloque final en `txtLog` (ASE, ruta salida, visibles **esperados post-Excel** F46/D9, F48/D47, E41/D28, D67/D66, estado validación). **No** mostrar `ConsolidadoAse.TotOpt` como si fuera D9.
- Progreso: pasar marquee a pasos (`ProgressBarStyle.Continuous`) con los hitos del procesador.
- **Prohibido:** colores, fuentes, iconos, “cards”, restyle. `ui-ux-specialist` solo si después se pide pulido visual.

### 2.7 Residuales HU-05 (opcionales, no bloquean)

- Visibles F46/F48 vs golden son más fuertes en harness que en `WorkbookLeafCoherence` runtime — la Capa A2 de golden cubre el hueco.
- Fallback `FirstOrDefault` si no matchea `Ase.Id` — aceptable a 1 ASE; no expandir.

### 2.8 File changes

| Archivo | Acción | Motivo |
|---|---|---|
| `Remuneracion.Core/Interfaces/IProcesadorRemuneracion.cs` | Crear | Caso de uso |
| `Remuneracion.Core/Models/SolicitudProcesoAse.cs` | Crear | Input del caso de uso |
| `Remuneracion.Core/Models/ResultadoProcesoAse.cs` | Crear | Output para UI/log |
| `Remuneracion.Core/Services/ProcesadorRemuneracion.cs` | Crear | Orquestación |
| `Remuneracion.Core/Services/ValidadorBasico.cs` | Crear | `IValidador` |
| `Remuneracion.Core/Interfaces/IValidador.cs` | Modificar | Overload con leaf |
| `Remuneracion.WinForms/Program.cs` | Modificar | Composition root |
| `Remuneracion.WinForms/Form1.cs` | Modificar | Stub → orquestación real |
| `Remuneracion.WinForms/Form1.Designer.cs` | Modificar | Fila carpeta salida + progreso por pasos |
| `Remuneracion.IntegrationTests/ValidadorBasicoTests.cs` | Crear | Dominio, sin Excel |
| `Remuneracion.IntegrationTests/GoldenLeafPathTests.cs` | Crear | Matriz Capa A |
| `Herramientas/VerificadorRecaudo/Program.cs` | Modificar | 1–2 casos Capa A (opcional si el test xUnit ya cubre) |

**No tocar:** semántica de `OpenXmlPlantillaWriter.Escribir*` (HU-04) ni el cell-map de escritura leaf (HU-05), salvo bug blocker.

---

## 3. SPEC / Acceptance Criteria

### Requirement 1 — Orquestación real de 1 ASE

El sistema **MUST** ejecutar el flujo locator → R1/R2/R4 → cálculo → leaf → validador → `GenerarWorkbook` para el ASE seleccionado y **MUST NOT** mutar la plantilla origen.

#### Scenario: Happy path Promoambiental Q1

- GIVEN carpeta `Docs/Insumos/REMUNERACION 2026071`, plantilla del período, ASE `1 - Promoambiental`, carpeta salida distinta
- WHEN el usuario pulsa Ejecutar
- THEN se genera `Remuneración 202607-1 Total.xlsx` en la carpeta salida
- AND el log (UI + Serilog) registra cada paso con timestamp
- AND la plantilla origen conserva su hash

#### Scenario: Falta R4

- GIVEN la carpeta del ASE sin `ReversiónPorComponente*`
- WHEN Ejecutar
- THEN `ArchivoFuenteNoEncontradoException` (o equivalente ya existente)
- AND no se certifica archivo de salida
- AND los controles se rehabilitan

### Requirement 2 — Destino de salida obligatorio

El sistema **MUST** exigir carpeta de salida y **MUST NOT** escribir in-place.

#### Scenario: Sin carpeta salida

- GIVEN fuentes y plantilla elegidas, salida vacía
- WHEN Ejecutar
- THEN se bloquea en UI (mismo patrón de validación actual) y no hay I/O

#### Scenario: Salida = plantilla

- GIVEN ruta calculada igual a la plantilla
- WHEN Ejecutar
- THEN no se llama escritura / falla fail-fast

### Requirement 3 — Validaciones básicas de dominio

`IValidador` **MUST** bloquear incoherencias de Fase 1 **antes** de certificar. **MUST NOT** abrir Excel. **MUST NOT** comparar TotOpt HU-02 vs F46.

#### Scenario: Q1 con AjustesSfT ≠ 0

- GIVEN un resultado con `AjustesSfT != 0`
- WHEN `Validar`
- THEN la lista no está vacía y el procesador no escribe

#### Scenario: Gate leaf vs consolidado

- GIVEN `leaf.R2.TotalOportunoEsperado` fuera de ±0.5 del consolidado
- WHEN `Validar(resultado, leaf)`
- THEN bloquea

### Requirement 4 — Golden Capa A (sin recálculo)

Las pruebas **MUST** ejecutar la matriz §2.4 A1–A5. **MUST NOT** usar cache de celdas fórmula de la **salida** como prueba de CA-3.

#### Scenario: Leaf vs golden

- GIVEN salida generada desde fuentes Promoambiental 202607-1
- WHEN se leen las 11 celdas leaf por OpenXML
- THEN coinciden con el golden ±0.5

#### Scenario: Visibles de dominio vs cache golden

- GIVEN los leaf leídos de fuente
- WHEN se calculan `TotalOportunoEsperado` / `ExtemporaneoEsperado` / R2 / R4
- THEN coinciden con cache golden `F46/D9`, `F48/D47`, `E41/D28`, `D67/D66` ±0.5

#### Scenario: Prohibición CA-3 falsa

- GIVEN la plantilla de trabajo es el propio golden
- WHEN las pruebas corren
- THEN **ningún** assert compara cache de `D9` (u otra fórmula) **de la salida** contra golden

### Requirement 5 — UI mínima y resultados honestos

La UI **MUST** permitir período, carpeta fuentes, plantilla, ASE, carpeta salida, progreso y log. El resumen **MUST** etiquetar F46/D9 como visible esperado **post-Excel**, no como `ConsolidadoAse.TotOpt`.

#### Scenario: Resumen no miente D9

- GIVEN proceso OK
- WHEN se imprime el resumen
- THEN aparece `16704332434.57` (±0.5) como visible esperado D9/F46
- AND no se presenta `19556118465.99` como valor de CONSOLIDADO D9

### AC resumidos vs Propuesta §10

| CA | HU-06 |
|---|---|
| CA-1 | UI real lee los 3 archivos del ASE elegido |
| CA-2 | Cálculo + visibles de dominio (honestos; TotOpt fuente ≠ D9) |
| CA-3 | **Parcial:** Capa A. Cache CONSOLIDADO post-Excel = residual |
| CA-4 | Reassert: fórmulas del cell-map intactas; plantilla no mutada |
| CA-5 | `IValidador` + gates HU-05 |
| CA-6 | Serilog + `txtLog` por paso |
| CA-7 | Período, carpetas, plantilla, **más salida** |

---

## 4. TASKS

## Review Workload Forecast

| Field | Value |
|---|---|
| Estimated changed lines | 480–720 |
| 400-line budget risk | High |
| Chained PRs recommended | Yes |
| Suggested split | PR1 Core procesador+validador → PR2 UI/composition → PR3 golden Capa A |
| Delivery strategy | ask-on-risk |
| Chain strategy | pending |

Decision needed before apply: Yes
Chained PRs recommended: Yes
Chain strategy: pending
400-line budget risk: High

### Suggested Work Units

| Unit | Goal | Likely PR | Notes |
|---|---|---|---|
| 1 | `IValidador` + `ProcesadorRemuneracion` + tests dominio | PR 1 | Sin UI |
| 2 | Composition + Form1 real + carpeta salida | PR 2 | Depende de PR 1 |
| 3 | `GoldenLeafPathTests` matriz A1–A5 | PR 3 | Depende de PR 1; puede paralelo a PR 2 |

### Phase 1 — Core

- [ ] 1.1 Overload `IValidador.Validar(ResultadoRemuneracion, WorkbookLeafInputs)` sin romper el método actual.
- [ ] 1.2 Crear `ValidadorBasico` con las reglas §2.5 (1 ASE, Ajustes=0, gran total, gate leaf; **no** TotOpt vs F46).
- [ ] 1.3 Crear `SolicitudProcesoAse`, `ResultadoProcesoAse`, `IProcesadorRemuneracion`, `ProcesadorRemuneracion` (pasos + `IProgress<string>`).
- [ ] 1.4 Fail-fast: lista de validación no vacía → `CalculoInvalidoException`; no dejar salida certificada.

### Phase 2 — UI y composition

- [ ] 2.1 `Program.cs`: new-up reader, leaf reader, `CalculoRemuneracion`, `ValidadorBasico`, `OpenXmlPlantillaWriter`, `ProcesadorRemuneracion`; inyectar en `Form1`. Locator en el form o helper WinForms.
- [ ] 2.2 Designer: fila “Carpeta salida” (mismos tamaños); desplazar ejecución/log; progreso Continuous.
- [ ] 2.3 Reemplazar stub de `btnEjecutar_Click`: validar 3 rutas, resolver ASE+prefijos, overwrite, `Task.Run`, log dual, resumen honesto D9≠TotOpt.
- [ ] 2.4 Catch de excepciones existentes; rehabilitar controles; no tragar errores.

### Phase 3 — Golden y evidencia

- [ ] 3.1 `ValidadorBasicoTests`: happy path Q1; Ajustes≠0; mismatch R2; no compara TotOpt vs F46.
- [ ] 3.2 `GoldenLeafPathTests`: A1 leaf vs golden; A2 dominio vs cache golden; A3 fórmulas; A4 hash plantilla; A5 ningún assert de cache de fórmula **en la salida**.
- [ ] 3.3 Fixtures: golden `Docs/Insumos/Remuneracion 202607-1 Total.xlsx`; fuentes `Docs/Insumos/REMUNERACION 2026071/1-Promoambiental/` (mismos nombres que IntegrationTests actuales). Salida en temp. **0 warnings**, CRLF, sin commit.

### Phase 4 — Documental

- [ ] 4.1 Dejar en el log de cierre de HU el residual CA-3 Excel (§5.3).
- [ ] 4.2 No reescribir `requirements/Fase1-Requerimientos.md` HU-05-legado como si fuera este alcance.

---

## 5. Test Strategy

| Capa | Qué | Cómo |
|---|---|---|
| Unidad | `ValidadorBasico` | Modelos in-memory, sin Excel |
| Integración | `ProcesadorRemuneracion` 1 ASE | Insumos reales, salida temp |
| Golden Capa A | Matriz §2.4 | OpenXML read-only + aritmética de dominio |
| UI | Ejecutar real | Prueba funcional manual (no hay harness WinForms); el form no se “testea” con click automatizado en esta HU |
| Capa B | CA-3 cache CONSOLIDADO | Manual Excel — §5.3 |

### 5.1 Fixtures

- Golden / plantilla de estructura: `Docs/Insumos/Remuneracion 202607-1 Total.xlsx`
- R1/R2/R4: carpeta `1-Promoambiental` (archivos ya referenciados en `WorkbookLeafInputsTests`)

### 5.2 Protocolo manual CA-3 (no CI)

1. Generar salida con la UI o el procesador a ruta **distinta** del golden.
2. Abrir la salida en Excel y forzar recálculo.
3. Comparar `D9,D28,D47,D66,D104,D109` vs golden ±0.5.
4. Esperado Q1 Promoambiental: D9=`16704332434.57`, D28=`54216385.68`, D47=`11673020`, D66=`-12054255.65`.
5. Adjuntar evidencia (captura o nota) si se quiere cerrar CA-3 de negocio. **No es gate de merge de HU-06.**

---

## 6. Architecture Validation Certificate

### 6.1 SOLID

| Principio | Veredicto |
|---|---|
| **S** | Form1 = UI; procesador = caso de uso; validador = reglas; writer = I/O. |
| **O** | Se agrega orquestación/validación sin reinterpretar HU-04/HU-05. |
| **L** | `OpenXmlPlantillaWriter` sigue siendo validation-only en `IPlantillaWriter` y leaf writer en el otro contrato. |
| **I** | Contratos chicos; locator no se mete a Core. |
| **D** | Core define `IProcesadorRemuneracion` / `IValidador`; WinForms e Infrastructure implementan/componen. |

### 6.2 Best Practices

- Un solo ASE, tres fuentes, ruta leaf ya certificada.
- Honestidad de golden por encima de “poner verde CA-3”.
- Sin DI framework, sin AI-slop visual.
- Auditoría por paso (Serilog).

### 6.3 Performance

- Un ASE, tres lecturas, una copia de workbook, escritura de 11 celdas. Irrelevante a escala Fase 1.
- `Task.Run` para no congelar el form.

**Veredicto:** APROBADO como cierre de prototipo §4.1 **si** se acepta CA-3 parcial.

---

## 7. Riesgos

| Riesgo | Likelihood | Mitigación |
|---|---|---|
| Alguien “cierra” CA-3 comparando cache de la copia del golden | Alta | A5 explícito en tests; este plan lo prohíbe |
| Form1 se vuelve god-object | Media | Procesador en Core; Form1 < orquestación de UI |
| Resumen UI muestra TotOpt HU-02 como D9 | Alta | AC Requirement 5; copy del log fijado |
| Inflar a 8 archivos / VALIDACION_* / DetRetri | Media | §4.2 en Out of Scope; rechazar PRs que lo metan |
| Plantilla = golden pisada por el usuario | Media | Carpeta salida distinta + hash A4 + MessageBox overwrite |
| Recálculo Excel nunca se hace | Media | Residual explícito; no fingir cierre de §10.2 |

---

## 8. Rollback

- Revertir PRs en orden inverso (UI → golden tests → Core).
- HU-04/HU-05 siguen válidas sin esta HU: el stub de Form1 puede volver.
- `Docs/Insumos/` no se versiona como mutado; cualquier pisada se restaura desde backup/git si estuviera trackeado (hoy untracked: **no usar esos xlsx como destino**).

---

## 9. Decisiones fijadas por este plan

1. Rector = Propuesta §4.1/§4.2/§6/§9/§10. El HU-05-legado de `requirements/` **no** manda.
2. Orquestador en Core; composition en `Program.cs`; Form1 flaco.
3. Salida = carpeta + `Periodo.NombreArchivo`; jamás in-place.
4. `IValidador` = dominio básico Fase 1; writer sigue siendo el guardián de fórmulas.
5. Golden HU-06 = Capa A (leaf + visibles de dominio + fórmulas + hash). Capa B Excel = manual.
6. CA-3 de §10 queda **parcial** y declarado. §10.2 (éxito general post-Excel) no se finge.
7. Apply espera aprobación + estrategia de PRs encadenados.

**Listo para aprobación del Ingeniero. No implementar hasta OK.**
