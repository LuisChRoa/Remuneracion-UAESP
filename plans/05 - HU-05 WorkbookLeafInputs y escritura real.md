# Plan 05 — HU-05: WorkbookLeafInputs y escritura real

> **Historia nueva:** prerrequisito formal para habilitar escritura funcional real del workbook
> **Fuente rectora:** `Docs/Propuesta_Proyecto_Automatizacion_Remuneracion_UAESP.md` (§3, §4.1, §6, §7, §9, §10)
> **Contexto de continuidad:** `requirements/Fase1-Requerimientos.md` (HU-04/HU-05 legado) y `plans/04 - HU-04 Escritura OpenXML.md`
> **Estado:** REBASE 2026-09-02 — AC-5/AC-6 alineados al workbook real (aprobación del usuario: "Si por fa")
> **Fecha:** 2026-09-02

---

## 0. Clarification Gate

### 0.1 Qué cambió y por qué HU-04 NO alcanzaba

| Punto | Resolución del plan | Evidencia ya verificada |
|---|---|---|
| HU-04 prometía escritura, ¿pero de qué tipo? | **Solo validación estructural honesta.** No escritura funcional. | `OpenXmlPlantillaWriter` hoy valida fórmulas y rechaza payload insuficiente. |
| ¿Por qué no bastan los agregados actuales? | Porque el workbook real consume **inputs editables hoja-a-hoja**, no solo resultados finales. | `D9 -> R1!F46`, `D47 -> R1!F48`, `D28 -> R2!E41`, `D66 -> R4!D67`; esas celdas upstream también son fórmulas. |
| ¿Esto es un parche escondido de HU-04? | **No.** Es una historia nueva con contrato nuevo de dominio y de escritura. | HU-04 cerró bien como *validation-only*; la capacidad faltante es otra. |
| ¿Impacta el roadmap original? | **Sí.** Esta HU se inserta antes del HU-05 legado. | Sin leaf inputs no existe base certificable para “golden test + UI + validación final” de escritura real. |

### 0.2 Decisión ejecutiva

La Propuesta sigue mandando: **preservar fórmulas** y diligenciar la plantilla real. El descubrimiento del workbook no contradice eso; lo que demuestra es que la arquitectura actual se quedó un nivel arriba. Antes de certificar “escritura real”, hay que modelar y leer los **leaf editable inputs** que alimentan las fórmulas profundas.

### 0.3 Rebase del roadmap

Secuencia corregida de Fase 1:

1. HU-01 — cerrado
2. HU-02 — cerrado
3. HU-03 — cerrado
4. HU-04 — cerrado como **validación OpenXML no destructiva**
5. **HU-05 (nuevo)** — `WorkbookLeafInputs + escritura real`
6. **HU-06 (antes HU-05)** — validación final, UI y pruebas golden sobre la ruta real de escritura

> Si no hacemos este rebase, el equipo estaría intentando certificar una capacidad que TODAVÍA no existe. Eso sería humo técnico. No va.

### 0.4 Rebase de coherencia R1 (post-implementación)

Dump OpenXML de `Docs/Insumos/Remuneracion 202607-1 Total.xlsx` (Promoambiental) demostró que **los agregados HU-02 de R1 no son los visibles del workbook**:

| Cantidad | Valor verificado | Origen |
|---|---|---|
| HU-02 `RecaudoComponenteR1.TotalOportuno` | `19556118465.99` | Fuente R1: fila `Componente/Total` col F |
| Visible workbook `R1!F46` / consolidado `D9` | `16704332434.57` | Fórmula `F25+F41-L25` |
| HU-02 `RecaudoComponenteR1.Extemporaneo` | `19549786950.62` | Fuente R1: 1.ª fila `Mes/Total` col F = **leaf F25** |
| Visible workbook `R1!F48` / consolidado `D47` | `11673020.00` | Fórmula `F30+F10-L10` |
| HU-02 `SaldosFavorR2.TotalOportuno` | `54216385.68` | Coincide con `E15+E26-K15` |
| HU-02 `ReversionR4.TotalReversiones` | `-12054255.65` | Coincide con `D9-P9` |

El AC-5 original (`r1.TotalOportuno` vs `leaf.R1.TotalOportunoEsperado` y `r1.Extemporaneo` vs `leaf.R1.ExtemporaneoEsperado`) es **imposible**: forzar esa comparación haría fallar el gate siempre. El workbook manda. El gate queda rebaseado en §2.5.

AC-6 original pedía recálculo Excel embebido. El repo **no** tiene motor Excel. HU-05 certifica escritura estructural + aritmética de visibles esperados. El recálculo Excel contra golden queda para **HU-06**.

---

## 1. PROPOSE

### 1.1 Intent

Introducir un contrato write-side explícito para representar los inputs editables reales del workbook (`WorkbookLeafInputs`) y habilitar que `OpenXmlPlantillaWriter` escriba **solo celdas leaf**, dejando que las fórmulas existentes propaguen los resultados hacia R1/R2/R4 visibles y luego al `CONSOLIDADO_TOTAL RECAUDO`.

### 1.2 Por qué esta HU es necesaria

- La Propuesta exige diligenciar hojas reales **sin romper fórmulas**.
- El workbook verificado demuestra que las celdas visibles hoy usadas como objetivo (`F46`, `F48`, `E41`, `D67`, `D9`, `D28`, `D47`, `D66`) **no son editables**; son salidas derivadas.
- Los modelos actuales (`RecaudoComponenteR1`, `SaldosFavorR2`, `ReversionR4`) describen el **resultado agregado** útil para cálculo, pero NO el **input granular** necesario para escritura.
- Sin ese contrato granular, cualquier intento de “escritura real” sería o destructivo o engañoso.

### 1.3 Scope

#### In Scope

- Nuevo concepto de dominio/contrato `WorkbookLeafInputs` para R1, R2 y R4.
- Lectura de leaf inputs desde fuentes reales existentes.
- Escritura OpenXML real de celdas editables leaf en hojas `Reporte Componentes R1`, `Rem. Anticipos R2` y `Reversion Pagos R4`.
- Verificación de coherencia entre leaf inputs y agregados calculados.
- Validación posterior de preservación de fórmulas y propagación esperada.

#### Out of Scope

- Multi-ASE.
- UI WinForms nueva.
- `AJUSTES-SF-T`, `SALDOS POR NOTA`, `RETRIBUCION NEGATIVA`.
- Resto de hojas ampliadas de la Fase 1 extendida.
- Motor de cálculo Excel interno dentro del repo.

### 1.4 Trazabilidad al workbook real

| Resultado visible | Fórmula verificada | Implicación arquitectónica |
|---|---|---|
| `CONSOLIDADO_TOTAL RECAUDO!D9` | `'Reporte Componentes R1'!F46` | No escribir D9; escribir inputs leaf de R1 |
| `Reporte Componentes R1!F46` | `F25 + F41 - L25` | R1 necesita 3 leaf inputs para TOT_OPT visible |
| `CONSOLIDADO_TOTAL RECAUDO!D47` | `'Reporte Componentes R1'!F48` | No escribir D47; escribir inputs leaf de EXTEMP |
| `Reporte Componentes R1!F48` | `F30 + F10 - L10` | R1 necesita 3 leaf inputs para EXTEMP visible |
| `CONSOLIDADO_TOTAL RECAUDO!D28` | `'Rem. Anticipos R2'!E41` | No escribir D28; escribir inputs leaf de R2 |
| `Rem. Anticipos R2!E41` | `E15 + E26 - K15` | R2 necesita 3 leaf inputs |
| `CONSOLIDADO_TOTAL RECAUDO!D66` | `'Reversion Pagos R4'!D67` | No escribir D66; escribir inputs leaf de R4 |
| `Reversion Pagos R4!D67` | `D9 - P9` | R4 necesita 2 leaf inputs |

### 1.5 Resultado esperado de negocio

El libro final se genera escribiendo inputs reales de detalle; las fórmulas del workbook permanecen intactas y producen los visibles de R1/R2/R4 y el consolidado esperado cuando el archivo es recalculado por Excel.

---

## 2. DESIGN

### 2.1 Decisiones de arquitectura

| ID | Decisión | Alternativa descartada | Rationale |
|---|---|---|---|
| D1 | Mantener modelos agregados y agregar modelos write-side paralelos | Inflar `RecaudoComponenteR1` / `SaldosFavorR2` / `ReversionR4` con responsabilidades de escritura | Separa cálculo de escritura; evita mezclar “qué total dio” con “qué celdas editables lo producen”. |
| D2 | Crear `IWorkbookLeafInputReader` nuevo | Romper `IRecaudoReader` con más métodos y más semántica | Respeta ISP; el reader actual sigue siendo aggregate-first. |
| D3 | Crear `IWorkbookLeafWriter` nuevo y dejar `IPlantillaWriter` como contrato legacy de validación | Reinterpretar `IPlantillaWriter` otra vez | HU-04 ya cerró como validation-only; la escritura real merece contrato propio y honesto. |
| D4 | Escribir solo leaf cells protegidas por un cell-map explícito | Escribir celdas visibles resumidas | Preserva fórmulas y respeta el workbook real. |
| D5 | Agregar gate de coherencia `leaf -> visible -> aggregate` | Confiar a ciegas en el writer | Si leaf inputs y agregados divergen, hay bug de parsing o mapping. Debe fallar antes de certificar. |
| D6 | Rebaselinar roadmap y pruebas golden después de esta HU | Seguir con HU-05 legado sin cambios | Las golden tests solo sirven cuando ya existe escritura real. |

### 2.2 Modelos nuevos propuestos

Seguir la convención actual: **archivos flat dentro de `Remuneracion.Core/Models/`**.

#### 2.2.1 Modelo raíz

`Remuneracion.Core/Models/WorkbookLeafInputs.cs`

```csharp
public sealed class WorkbookLeafInputs
{
    public Ase Ase { get; set; } = new();
    public Periodo Periodo { get; set; } = new();
    public WorkbookLeafInputsR1 R1 { get; set; } = new();
    public WorkbookLeafInputsR2 R2 { get; set; } = new();
    public WorkbookLeafInputsR4 R4 { get; set; } = new();
}
```

#### 2.2.2 Submodelos por hoja

- `Remuneracion.Core/Models/WorkbookLeafInputsR1.cs`
- `Remuneracion.Core/Models/WorkbookLeafInputsR2.cs`
- `Remuneracion.Core/Models/WorkbookLeafInputsR4.cs`

Regla de naming:

1. **Primero** intentar nombre semántico basado en label visible del workbook.
2. **Si el label no es estable o no existe**, usar slot técnico controlado por mapa (`SlotF25`, `SlotF41`, etc.) **solo como fallback temporal**, nunca como lenguaje de negocio definitivo.

Cada submodelo DEBE exponer además sus visibles esperados como propiedades computadas para coherencia:

- `WorkbookLeafInputsR1.TotalOportunoEsperado`
- `WorkbookLeafInputsR1.ExtemporaneoEsperado`
- `WorkbookLeafInputsR2.TotalOportunoEsperado`
- `WorkbookLeafInputsR4.TotalReversionEsperada`

### 2.3 Contratos nuevos propuestos

#### `Remuneracion.Core/Interfaces/IWorkbookLeafInputReader.cs`

Responsabilidad: poblar `WorkbookLeafInputs` desde las fuentes reales R1/R2/R4.

```csharp
public interface IWorkbookLeafInputReader
{
    WorkbookLeafInputs LeerLeafInputs(Ase ase, Periodo periodo, string rutaR1, string rutaR2, string rutaR4);
}
```

#### `Remuneracion.Core/Interfaces/IWorkbookLeafWriter.cs`

Responsabilidad: generar un workbook de salida copiando plantilla y escribiendo únicamente celdas leaf editables.

```csharp
public interface IWorkbookLeafWriter
{
    void GenerarWorkbook(string rutaPlantillaOrigen, string rutaSalida, ResultadoRemuneracion resultado, WorkbookLeafInputs leafInputs);
}
```

### 2.4 Estrategia de implementación en Infrastructure

#### Readers

- **No** duplicar parsing a lo loco.
- Extraer helpers comunes desde `ExcelDataReaderRecaudoReader` a un helper interno compartido, por ejemplo:
  - `Remuneracion.Infrastructure/Excel/ExcelWorksheetNavigator.cs`
- Crear lector dedicado:
  - `Remuneracion.Infrastructure/Excel/ExcelDataReaderWorkbookLeafInputReader.cs`

Este lector debe:

1. abrir los mismos archivos reales ya usados por HU-02;
2. localizar filas/columnas por headers y labels, nunca por fila fija ciega;
3. poblar leaf slots mínimos de R1/R2/R4;
4. derivar visibles esperados;
5. comparar contra agregados (`LeerR1/LeerR2/LeerR4`) dentro de tolerancia ±0.5.

#### Writer

`Remuneracion.Infrastructure/Excel/OpenXmlPlantillaWriter.cs` pasa a implementar también `IWorkbookLeafWriter`.

Flujo de escritura:

```text
Plantilla origen
  └─► copiar a ruta de salida
        └─► validar hojas y fórmulas protegidas
              └─► escribir solo leaf cells R1/R2/R4
                    └─► guardar
                          └─► revalidar que fórmulas resumen sigan intactas
```

Protecciones obligatorias:

- `D9`, `D28`, `D47`, `D66`, `D104`, `D109` siguen siendo fórmula.
- `F46`, `F48`, `E41`, `D67` siguen siendo fórmula.
- Si cualquier fórmula protegida desaparece o es alterada, se aborta.
- La plantilla original nunca se muta in-place.

### 2.5 Coherencia entre agregados y leaf inputs

Este punto es CLAVE.

Los agregados existentes siguen siendo útiles para:

- `CalculoRemuneracion`
- validaciones de negocio
- comparación con tolerancia

Los leaf inputs nuevos sirven para:

- escritura real del workbook
- trazabilidad hoja-a-hoja

**Gate obligatorio antes de escribir (rebase 2026-09-02):**

| Comparación | Regla | Por qué |
|---|---|---|
| `leaf.R1.F25` vs `r1.Extemporaneo` | diferencia ≤ ±0.5 | En el workbook, F25 **es** la 1.ª fila Mes/Total col F (agregado HU-02 Extemp). |
| `leaf.R2.TotalOportunoEsperado` vs `r2.TotalOportuno` | diferencia ≤ ±0.5 | `E15+E26-K15` coincide con GrandTotal − Especiales. |
| `leaf.R4.TotalReversionEsperada` vs `r4.TotalReversiones` | diferencia ≤ ±0.5 | `D9-P9` coincide con el total R4. |

**NO comparar** (haría fallar el gate siempre):

| Comparación prohibida | Evidencia |
|---|---|
| `r1.TotalOportuno` vs `leaf.R1.TotalOportunoEsperado` | `19556118465.99` ≠ `16704332434.57` (`F46`) |
| `r1.Extemporaneo` vs `leaf.R1.ExtemporaneoEsperado` | `19549786950.62` ≠ `11673020` (`F48`) |

Los visibles `TotalOportunoEsperado` (`F46`) y `ExtemporaneoEsperado` (`F48`) se certifican contra el **workbook de referencia**, no contra los agregados HU-02.

Si el gate rebaseado no cuadra, no se escribe. Primero se corrige el mapping.

### 2.6 File changes previstos

| Archivo | Acción | Motivo |
|---|---|---|
| `Remuneracion.Core/Models/WorkbookLeafInputs.cs` | Crear | Modelo raíz write-side |
| `Remuneracion.Core/Models/WorkbookLeafInputsR1.cs` | Crear | Leaf inputs R1 |
| `Remuneracion.Core/Models/WorkbookLeafInputsR2.cs` | Crear | Leaf inputs R2 |
| `Remuneracion.Core/Models/WorkbookLeafInputsR4.cs` | Crear | Leaf inputs R4 |
| `Remuneracion.Core/Interfaces/IWorkbookLeafInputReader.cs` | Crear | Contrato de lectura granular |
| `Remuneracion.Core/Interfaces/IWorkbookLeafWriter.cs` | Crear | Contrato de escritura real |
| `Remuneracion.Infrastructure/Excel/ExcelWorksheetNavigator.cs` | Crear | Helper compartido para parsing dinámico |
| `Remuneracion.Infrastructure/Excel/ExcelDataReaderWorkbookLeafInputReader.cs` | Crear | Lectura de leaf inputs |
| `Remuneracion.Infrastructure/Excel/OpenXmlPlantillaWriter.cs` | Modificar | Soportar generación real de workbook |
| `Herramientas/VerificadorRecaudo/Program.cs` | Modificar | Evidencia con workbook real y guardas de fórmula |
| `Remuneracion.WinForms/Remuneracion.WinForms.slnx` | Modificar | Incluir proyecto de pruebas si se aprueba |
| `Remuneracion.IntegrationTests/Remuneracion.IntegrationTests.csproj` | Crear | Pruebas de integración con insumos reales |

### 2.7 Visual Design Intent

No aplica. Esta HU no toca UI.

---

## 3. SPEC / Acceptance Criteria

### Requirement 1 — Lectura granular para escritura real

El sistema **MUST** poder construir `WorkbookLeafInputs` para R1, R2 y R4 a partir de archivos fuente reales y **MUST** derivar desde esos leaf inputs los visibles esperados del workbook.

#### Scenario: R1/R2/R4 producen leaf inputs coherentes

- GIVEN las fuentes reales de `Docs/Insumos/REMUNERACION 2026071/1-Promoambiental/`
- WHEN se invoca `LeerLeafInputs(...)`
- THEN se obtienen leaf inputs no vacíos para R1, R2 y R4
- AND `leaf.R1.F25` coincide con `r1.Extemporaneo` dentro de ±0.5
- AND `leaf.R2.TotalOportunoEsperado` coincide con `r2.TotalOportuno` dentro de ±0.5
- AND `leaf.R4.TotalReversionEsperada` coincide con `r4.TotalReversiones` dentro de ±0.5
- AND `leaf.R1.TotalOportunoEsperado` coincide con el visible de referencia `R1!F46` / `D9` (`16704332434.57` en Promoambiental 202607-1) dentro de ±0.5
- AND `leaf.R1.ExtemporaneoEsperado` coincide con el visible de referencia `R1!F48` / `D47` (`11673020` en Promoambiental 202607-1) dentro de ±0.5

### Requirement 2 — Escritura real preservando fórmulas

El sistema **MUST** escribir únicamente celdas leaf editables y **MUST NOT** sobrescribir fórmulas protegidas del workbook.

#### Scenario: Se escriben leaf cells y no celdas derivadas

- GIVEN una plantilla real y un `WorkbookLeafInputs` válido
- WHEN se genera el workbook de salida
- THEN cambian solo las celdas leaf mapeadas de R1/R2/R4
- AND `F46`, `F48`, `E41`, `D67`, `D9`, `D28`, `D47`, `D66`, `D104`, `D109` permanecen como fórmula

### Requirement 3 — Coherencia previa a escritura

El sistema **MUST** bloquear la escritura si el gate rebaseado de §2.5 diverge por encima de la tolerancia.

#### Scenario: Mismatch entre leaf y aggregate

- GIVEN leaf inputs cuyo `R2.TotalOportunoEsperado` o `R4.TotalReversionEsperada` o `R1.F25` no coincide con el agregado correspondiente (`SaldosFavorR2.TotalOportuno`, `ReversionR4.TotalReversiones`, `RecaudoComponenteR1.Extemporaneo`)
- WHEN se solicita `GenerarWorkbook(...)`
- THEN la operación falla con mensaje descriptivo
- AND no se genera un archivo certificado como válido

### Requirement 4 — Propagación certificable sin motor Excel embebido

El repo **NO** tiene motor de recálculo Excel. HU-05 **MUST** certificar:

1. fórmulas protegidas intactas tras la escritura (OpenXML);
2. valores leaf escritos coinciden con el payload;
3. aritmética de visibles esperados (`F25+F41-L25`, `F30+F10-L10`, `E15+E26-K15`, `D9-P9`) coincide con el libro de referencia dentro de ±0.5.

El recálculo real en Excel contra golden **queda fuera de HU-05** y entra en **HU-06**.

#### Scenario: Escritura estructural y aritmética de visibles

- GIVEN un workbook generado con leaf inputs reales
- WHEN se inspecciona por OpenXML (sin recálculo Excel)
- THEN las fórmulas protegidas siguen presentes
- AND los valores leaf escritos coinciden con el payload
- AND los visibles esperados calculados en dominio coinciden con el libro de referencia dentro de ±0.5

### Acceptance Criteria resumidos

| ID | Criterio |
|---|---|
| AC-1 | Existe contrato `WorkbookLeafInputs` para R1/R2/R4 |
| AC-2 | Existe `IWorkbookLeafInputReader` y puebla inputs reales desde fuentes |
| AC-3 | Existe `IWorkbookLeafWriter` y copia plantilla a salida sin mutar origen |
| AC-4 | Solo se escriben celdas leaf; fórmulas protegidas permanecen intactas |
| AC-5 | Gate rebaseado §2.5: `F25` vs Extemp HU-02; R2/R4 vs agregados; visibles `F46`/`F48` vs workbook de referencia. **No** comparar TotOpt/Extemp HU-02 contra `TotalOportunoEsperado`/`ExtemporaneoEsperado`. |
| AC-6 | OpenXML certifica fórmulas intactas + leaf escritos + aritmética de visibles. Recálculo Excel/golden = HU-06. |
| AC-7 | `OpenXmlPlantillaWriter` conserva el modo validation-only legado para HU-04 |

---

## 4. TASKS

## Review Workload Forecast

| Field | Value |
|---|---|
| Estimated changed lines | 520–780 |
| 400-line budget risk | High |
| Chained PRs recommended | Yes |
| Suggested split | PR1 contratos/modelos → PR2 readers/coherencia → PR3 writer/verificación |
| Delivery strategy | ask-on-risk |
| Chain strategy | pending |

Decision needed before apply: Yes
Chained PRs recommended: Yes
Chain strategy: pending
400-line budget risk: High

### Suggested Work Units

| Unit | Goal | Likely PR | Notes |
|---|---|---|---|
| 1 | Contratos y modelos `WorkbookLeafInputs` | PR 1 | Base segura; sin tocar escritura aún |
| 2 | Lectura granular + gates de coherencia | PR 2 | Depende de PR 1 |
| 3 | Escritura real OpenXML + verificación | PR 3 | Depende de PR 2 |

### Phase 1: Contratos Core

- [ ] 1.1 Crear `Remuneracion.Core/Models/WorkbookLeafInputs.cs` y submodelos `WorkbookLeafInputsR1.cs`, `WorkbookLeafInputsR2.cs`, `WorkbookLeafInputsR4.cs`.
- [ ] 1.2 Crear `Remuneracion.Core/Interfaces/IWorkbookLeafInputReader.cs` e `IWorkbookLeafWriter.cs`.
- [ ] 1.3 Documentar en XML comments qué celdas visibles protegen esos modelos y qué propiedades computadas derivan.

### Phase 2: Lectura granular desde fuentes

- [ ] 2.1 Extraer helpers comunes de `Remuneracion.Infrastructure/Excel/ExcelDataReaderRecaudoReader.cs` a `Remuneracion.Infrastructure/Excel/ExcelWorksheetNavigator.cs`.
- [ ] 2.2 Crear `Remuneracion.Infrastructure/Excel/ExcelDataReaderWorkbookLeafInputReader.cs` para poblar leaf inputs R1/R2/R4.
- [ ] 2.3 Implementar el gate rebaseado §2.5: `F25` vs Extemp HU-02; R2/R4 vs agregados; visibles `F46`/`F48` vs workbook de referencia.

### Phase 3: Escritura real segura

- [ ] 3.1 Ampliar `Remuneracion.Infrastructure/Excel/OpenXmlPlantillaWriter.cs` para implementar `IWorkbookLeafWriter`.
- [ ] 3.2 Centralizar un cell-map de hojas/celdas editables y fórmulas protegidas en `Remuneracion.Infrastructure/Excel/WorkbookLeafCellMap.cs`.
- [ ] 3.3 Generar copia de salida y escribir solo celdas leaf de `Reporte Componentes R1`, `Rem. Anticipos R2` y `Reversion Pagos R4`.
- [ ] 3.4 Revalidar después del guardado que `F46`, `F48`, `E41`, `D67`, `D9`, `D28`, `D47`, `D66`, `D104` y `D109` siguen siendo fórmula.

### Phase 4: Evidencia y pruebas

- [ ] 4.1 Actualizar `Herramientas/VerificadorRecaudo/Program.cs` para verificar leaf inputs, protección de fórmulas y generación de workbook real.
- [ ] 4.2 Crear `Remuneracion.IntegrationTests/Remuneracion.IntegrationTests.csproj` y pruebas con `Docs/Insumos/` como fixtures reales.
- [ ] 4.3 Agregar el proyecto de pruebas a `Remuneracion.WinForms/Remuneracion.WinForms.slnx`.

### Phase 5: Rebase documental

- [ ] 5.1 Actualizar trazabilidad para que el HU-05 legado pase a HU-06 en documentos vivos del proyecto.
- [ ] 5.2 Dejar explícito en el siguiente plan que las golden/UI validan la ruta real de escritura, no la ruta validation-only.

---

## 5. Test Strategy

### 5.1 Fuentes de prueba reales

Usar como fixtures mínimos reales:

- `Docs/Insumos/Remuneracion 202607-1 Total.xlsx`
- `Docs/Insumos/REMUNERACION 2026071/1-Promoambiental/Recaudoporcomponente_*.xlsx`
- `Docs/Insumos/REMUNERACION 2026071/1-Promoambiental/RerpoteDetalleSaldosaFavor_*.xlsx`
- `Docs/Insumos/REMUNERACION 2026071/1-Promoambiental/ReversiónPorComponente_*.xlsx`

### 5.2 Qué debe probarse

| Capa | Evidencia |
|---|---|
| Reader granular | Los leaf inputs correctos se extraen desde R1/R2/R4 por labels |
| Coherencia (gate rebaseado) | `F25` vs Extemp HU-02; R2/R4 vs agregados; `F46`/`F48` aritméticos vs workbook de referencia |
| Writer estructural | Solo cambian celdas leaf; fórmulas protegidas siguen presentes |
| Writer funcional HU-05 | Aritmética de visibles esperados coincide con el libro de referencia. Recálculo Excel = HU-06 |

### 5.3 Protocolo de validación HU-05 (sin motor Excel)

1. Generar copia nueva del workbook desde plantilla real.
2. Escribir leaf inputs R1/R2/R4.
3. Verificar por OpenXML que las fórmulas protegidas siguen intactas.
4. Verificar que los valores leaf escritos coinciden con el payload.
5. Verificar en dominio que `F25+F41-L25`, `F30+F10-L10`, `E15+E26-K15` y `D9-P9` coinciden con el libro de referencia dentro de ±0.5.

> El recálculo real en Excel contra golden **no** es aceptación de HU-05. Queda para HU-06. OpenXML no calcula fórmulas; no se finge que sí.

---

## 6. Architecture Validation Certificate

### 6.1 SOLID

| Principio | Veredicto |
|---|---|
| **S** — SRP | Aprobado: agregados siguen en cálculo; leaf inputs viven en write-side; writer no vuelve a fingir cálculo. |
| **O** — OCP | Aprobado: se agregan contratos y modelos nuevos sin romper HU-04 cerrada. |
| **L** — LSP | Aprobado: `OpenXmlPlantillaWriter` mantiene compatibilidad con el contrato legacy y suma uno nuevo para escritura real. |
| **I** — ISP | Aprobado: `IRecaudoReader`, `IPlantillaWriter`, `IWorkbookLeafInputReader` e `IWorkbookLeafWriter` quedan separados por responsabilidad. |
| **D** — DIP | Aprobado: Core define contratos; Infrastructure implementa OpenXML/ExcelDataReader sin contaminar el dominio. |

### 6.2 Best Practices

- La verdad del workbook manda sobre el diseño previo.
- Se escribe en los leaves, no en las fórmulas.
- Se evita mutar la plantilla original.
- Se agrega coherencia interna antes de escribir.
- Se rebaselina el roadmap en vez de esconder deuda.

### 6.3 Performance

- Lectura focalizada sobre pocas hojas y slots concretos.
- Escritura puntual de celdas leaf; no recorre todo el workbook sin necesidad.
- Revalidación acotada a fórmulas protegidas y visibles canónicos.

**Veredicto:** **APROBADO** como prerequisito arquitectónico obligatorio para habilitar escritura real certificable.

---

## 7. Riesgos y mitigaciones

| Riesgo | Likelihood | Mitigación |
|---|---|---|
| Nombrar leaf slots con semántica incorrecta | Media | Congelar nombres públicos solo después de verificar labels reales del workbook |
| Duplicar lógica entre aggregate reader y leaf reader | Alta | Extraer helper compartido antes de implementar lectura granular |
| Escribir una celda derivada por error | Alta | Cell-map con fórmulas protegidas + pruebas negativas |
| Certificar sin recálculo real | Alta | Protocolo explícito de aceptación con workbook recalculado |
| Dejar HU-05 legado sin rebase | Alta | Documentar desde ya el corrimiento a HU-06 |

---

## 8. Rollback Plan

- Si el nuevo contrato no es aprobado, HU-04 permanece como única capacidad vigente de `OpenXmlPlantillaWriter` (validación-only).
- Los cambios futuros de implementación deben poder revertirse por PR slice: contratos, reader granular y writer real por separado.
- La plantilla real nunca debe quedar dañada porque la nueva escritura trabajará sobre copia de salida.

---

## 9. Decisiones que quedan fijadas por este plan

1. `WorkbookLeafInputs` es un **nuevo contrato formal**, no una extensión oportunista de los agregados actuales.
2. `IPlantillaWriter` NO vuelve a reinterpretarse; la escritura real entra por un contrato nuevo.
3. `OpenXmlPlantillaWriter` seguirá existiendo, pero con dos modos honestos: validación legacy y escritura leaf real.
4. El roadmap de Fase 1 se reordena: primero leaf inputs + escritura real, después golden/UI/validación final.
5. La aprobación de este plan deja lista la historia para implementación, pero **requiere decisión de estrategia de PR encadenados antes del apply**.
