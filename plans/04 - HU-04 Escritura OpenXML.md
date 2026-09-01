# Plan 04 — HU-04: OpenXML honesto sobre workbook profundamente formula-driven

> **Historia:** HU-04 — Escritura de plantilla (OpenXML)
> **Fuente rectora:** `Docs/Propuesta_Proyecto_Automatizacion_Remuneracion_UAESP.md` (§4.1, §4.2, §5.1, §6, §7, §9, §10)
> **Fuente complementaria con conflicto detectado:** `requirements/Fase1-Requerimientos.md` (HU-04 amplía a “todas las hojas”, pero esa ampliación NO prevalece sobre la Propuesta)
> **Estado:** REFORMULADO OTRA VEZ — listo para nueva aprobación
> **Fecha:** 2026-09-01
> **Motivo de reformulación:** el workbook real `Docs/Insumos/Remuneracion 202607-1 Total.xlsx` confirmó que el consolidado y sus primeras celdas upstream visibles también son fórmulas; por lo tanto, el payload actual NO alcanza para una escritura funcional honesta de R1/R2/R4.

---

## 0. Clarification Gate

| Pregunta / tensión | Resolución honesta del plan | Evidencia verificada |
|---|---|---|
| ¿`CONSOLIDADO_TOTAL RECAUDO` era escribible? | **No.** Se mantiene como vista derivada y solo se certifica validación estructural. | `D9`, `D28`, `D47`, `D66`, `D85`, `D104`, `D109` son fórmula o dependen de fórmulas en el workbook real. |
| ¿`Reporte Componentes R1!F46/F48`, `Rem. Anticipos R2!E41` y `Reversion Pagos R4!D67` eran inputs escribibles? | **Tampoco.** Eran un supuesto falso. Son celdas derivadas por fórmula y no deben sobrescribirse. | `F46 = F25 + F41 - L25`; `F48 = F30 + F10 - L10`; `E41 = E15 + E26 - K15`; `D67 = D9 - P9`. |
| Entonces, ¿HU-04 sigue siendo de escritura? | **No en sentido funcional completo.** Con contratos actuales, HU-04 queda reducido a validación OpenXML no destructiva + introspección segura + explicitación de prerequisitos para futura escritura real. | Los modelos actuales solo exponen agregados y no los inputs hoja-a-hoja necesarios para reconstruir las fórmulas profundas. |
| ¿`EscribirDetalleR1/R2/R4(...)` pueden prometer algo útil? | **Sí, pero no “escribir valores reales”**. Su alcance certificable pasa a ser preflight estructural por hoja, detección de fórmulas, trazabilidad de cadena y fail-fast explicativo si se intenta escribir con payload insuficiente. | Contrato `IPlantillaWriter` existe, pero el workbook real invalida la semántica anterior. |
| ¿`DetallePorComponente` sigue siendo deuda útil? | **Sí, pero solo parcial.** Sirve como insumo para futura expansión de R1, NO basta por sí solo para poblar los verdaderos leaf inputs que alimentan `F46/F48`. | Aunque `DetallePorComponente` mejora granularidad, las fórmulas de R1 combinan más celdas/zonas que no están modeladas hoy. |
| ¿Qué pasa con `SharedStrings`? | **Sigue siendo obligatorio** para cualquier matching textual en OpenXML. | El workbook real contiene `xl/sharedStrings.xml`. |

### Decisión de alcance

Esta reformulación elimina la falsa promesa. **HU-04 queda en modo validation-first**:

1. **NO** se escriben celdas del consolidado.
2. **NO** se escriben `F46`, `F48`, `E41` ni `D67`.
3. `EscribirConsolidado(...)` queda como **validación estructural** del consolidado derivado.
4. `EscribirDetalleR1/R2/R4(...)` quedan como **validación estructural por hoja + fail-fast honesto** hasta que exista payload profundo.
5. Se deja especificado el **prerequisito exacto** para una futura HU de escritura real.

### Conclusión ejecutiva

La versión anterior seguía siendo demasiado optimista. Con la nueva verdad del workbook, **HU-04 ya no es parcialmente escribible de forma certificable con los modelos actuales**. El entregable honesto y ejecutable es **validación-only**, no escritura funcional.

---

## 1. PROPOSE — Propuesta

### 1.1 Intent

Convertir HU-04 en una entrega realista y verificable: inspeccionar el workbook OpenXML sin dañarlo, certificar que la cadena de fórmulas clave sigue intacta y dejar trazados los contratos faltantes para habilitar escritura real en una historia posterior.

### 1.2 Scope

#### In Scope

- Validación estructural de `CONSOLIDADO_TOTAL RECAUDO` como hoja derivada.
- Validación estructural de `Reporte Componentes R1`, `Rem. Anticipos R2` y `Reversion Pagos R4`.
- Introspección OpenXML segura con soporte de `SharedStrings`.
- Verificación explícita de la cadena de fórmulas conocida:
  - `D9 -> R1!F46 -> F25 + F41 - L25`
  - `D47 -> R1!F48 -> F30 + F10 - L10`
  - `D28 -> R2!E41 -> E15 + E26 - K15`
  - `D66 -> R4!D67 -> D9 - P9`
- Fail-fast descriptivo cuando el contrato de escritura solicitado no puede cumplirse con el payload actual.
- Documento de prerequisitos para expansión futura de escritura real.

#### Out of Scope

- Escribir valores funcionales en `CONSOLIDADO_TOTAL RECAUDO`.
- Escribir valores funcionales en `F46`, `F48`, `E41` o `D67`.
- Simular escritura con “última fórmula visible” o con hardcodes de celdas.
- Implementar `AJUSTES - SF-T`, `SALDOS POR NOTA` o `RETRIBUCION NEGATIVA`.
- Prometer libro completo o escritura real de R1/R2/R4 con los modelos actuales.

### 1.3 Capabilities

| Capability | Tipo | Descripción |
|---|---|---|
| `openxml-structural-validation` | Modificada | HU-04 pasa de “writer funcional” a validador estructural no destructivo |
| `formula-chain-verification` | Nueva | Verifica fórmulas del consolidado y de sus primeras hojas upstream |
| `openxml-worksheet-text-resolution` | Nueva | Matching textual seguro resolviendo `SharedStrings` |
| `future-write-contract-prerequisites` | Nueva | Define el payload/reader/contrato faltante para escritura real posterior |

### 1.4 Approach

La estrategia correcta es **dejar de escribir donde solo vemos resultados derivados**. Infrastructure debe comportarse como inspector del workbook real: abrir, resolver textos, detectar fórmulas, seguir dependencias conocidas y rechazar cualquier intento de escritura cuya semántica no esté respaldada por inputs verdaderamente modelados.

### 1.5 Success Criteria

- [ ] `OpenXmlPlantillaWriter` deja de tratar `F46`, `F48`, `E41` y `D67` como celdas escribibles.
- [ ] `EscribirConsolidado(...)` verifica estructura y fórmulas sin sobrescribir nada.
- [ ] `EscribirDetalleR1/R2/R4(...)` fallan de manera honesta y explicativa si se pretende escritura funcional con payload insuficiente.
- [ ] La introspección OpenXML resuelve `SharedStrings`, `InlineString`, números y fórmulas.
- [ ] Queda documentado el prerequisito exacto para una futura HU de escritura real.

---

## 2. DESIGN — Diseño técnico

### 2.1 Decisiones de arquitectura

| ID | Decisión | Alternativa descartada | Rationale |
|---|---|---|---|
| D1 | HU-04 se reduce a **validation-only** | Insistir en “partially writable” sin evidencia de leaf inputs | Sería vender humo: las celdas visibles de R1/R2/R4 también son fórmula. |
| D2 | `EscribirConsolidado(...)` se conserva como validador estructural | Borrarlo o seguir escribiendo D-cells | El contrato existe y puede seguir siendo útil sin mentir. |
| D3 | `EscribirDetalleR1/R2/R4(...)` pasan a preflight estructural + fail-fast | Seguir escribiendo `F46/F48/E41/D67` | Eso violaría la regla crítica de preservar fórmulas. |
| D4 | El plan explicita un **contrato futuro de inputs profundos** | Ampliar a ciegas los modelos actuales | La escritura real exige modelar inputs verdaderos, no resultados agregados. |
| D5 | `DetallePorComponente` se mantiene como deuda útil pero insuficiente | Marcarlo como resuelto | Ayuda a R1, pero no cubre todas las hojas/celdas leaf requeridas por las fórmulas profundas. |
| D6 | `SharedStrings` y matching por contenido son obligatorios | Matching por índices/celdas fijas | El workbook real usa textos compartidos y puede variar layout. |

### 2.2 Nueva semántica contractual de `IPlantillaWriter`

#### `EscribirConsolidado(string rutaPlantilla, ResultadoRemuneracion resultado)`

**Responsabilidad certificable:**

1. abrir el workbook;
2. localizar `CONSOLIDADO_TOTAL RECAUDO`;
3. validar que `D9`, `D28`, `D47`, `D66`, `D85`, `D104` y `D109` sigan siendo fórmulas o dependencias derivadas esperadas;
4. validar que las referencias upstream esperadas sigan presentes;
5. fallar si el workbook fue alterado o si alguna fórmula fue reemplazada por valor.

**No hará:**

- escribir valores en el consolidado;
- recalcular negocio;
- asumir que `resultado` contiene todos los insumos de hoja necesarios para escribir workbooks reales.

#### `EscribirDetalleR1(string rutaPlantilla, Ase ase, RecaudoComponenteR1 datos)`

**Responsabilidad certificable temporal:**

- validar que `Reporte Componentes R1` exista y que `F46/F48` sean fórmulas;
- verificar la cadena conocida `F46 = F25 + F41 - L25` y `F48 = F30 + F10 - L10`;
- rechazar escritura funcional con excepción explícita que indique insuficiencia del payload actual.

#### `EscribirDetalleR2(string rutaPlantilla, Ase ase, SaldosFavorR2 datos)`

**Responsabilidad certificable temporal:**

- validar que `Rem. Anticipos R2` exista y que `E41` sea fórmula;
- verificar la cadena `E41 = E15 + E26 - K15`;
- rechazar escritura funcional por insuficiencia del payload actual.

#### `EscribirDetalleR4(string rutaPlantilla, Ase ase, ReversionR4 datos)`

**Responsabilidad certificable temporal:**

- validar que `Reversion Pagos R4` exista y que `D67` sea fórmula;
- verificar la cadena `D67 = D9 - P9`;
- rechazar escritura funcional por insuficiencia del payload actual.

### 2.3 Prerequisito exacto para escritura real futura

La futura HU de escritura real necesita **un contrato nuevo o ampliado por hoja que modele los verdaderos inputs editables leaf** del workbook, no solo agregados. Eso implica:

1. **R1:** un modelo que exponga semánticamente los inputs que alimentan `F25`, `F41`, `L25`, `F30`, `F10`, `L10` y cualquier otra celda editable real relacionada con TOT_OPT y EXTEMP.
2. **R2:** un modelo que exponga los inputs reales detrás de `E15`, `E26`, `K15`.
3. **R4:** un modelo que exponga los inputs reales detrás de `D9`, `P9`.
4. **Readers/Fuentes:** lectores capaces de poblar esos nuevos modelos desde archivos fuente reales o desde nuevas fuentes aún no representadas.
5. **Trazabilidad:** mapping semántico hoja-campo-celda validado contra workbook real, no solo por referencia A1.

**Nombre sugerido del prerequisito:** `WorkbookLeafInputs`.

### 2.4 Flujo de datos reformulado

```text
Workbook real ──► OpenXmlPlantillaWriter
      │                  │
      │                  ├──► Resuelve SharedStrings y hojas requeridas
      │                  ├──► Verifica fórmulas del consolidado
      │                  ├──► Verifica fórmulas en R1/R2/R4
      │                  └──► Falla honestamente si el payload no cubre leaf inputs
      │
      └──► Futuro cambio: nuevos readers/modelos poblarán WorkbookLeafInputs
```

### 2.5 Helpers / contratos internos recomendados

```csharp
private static WorksheetPart ObtenerHojaRequerida(WorkbookPart workbookPart, string nombreHoja)
private static Cell ObtenerCeldaExistente(Worksheet worksheet, string referencia)
private static string LeerTextoCelda(Cell cell, WorkbookPart workbookPart)
private static void ValidarCeldaTieneFormula(Cell cell, string hoja, string referencia)
private static void ValidarFormulaContiene(Cell cell, string hoja, string referencia, params string[] fragmentos)
private static void LanzarPayloadInsuficiente(string operacion, string hoja, string detalle)
```

### 2.6 File Changes

| Archivo | Acción | Descripción |
|---|---|---|
| `Remuneracion.Infrastructure/Excel/OpenXmlPlantillaWriter.cs` | Modificar | Reorientar a validación-only, quitar falsas escrituras y agregar fail-fast honesto |
| `Herramientas/VerificadorRecaudo/Program.cs` | Modificar | Validar cadenas de fórmula profundas y casos negativos de payload insuficiente |
| `Remuneracion.Core/Interfaces/IPlantillaWriter.cs` | Sin cambios | Se conserva el contrato; cambia la semántica documentada de esta HU |
| `Remuneracion.Core/Models/RecaudoComponenteR1.cs` | Sin cambios en HU-04 | `DetallePorComponente` se mantiene como deuda útil, no como solución final |
| `Remuneracion.Core/Models/SaldosFavorR2.cs` | Sin cambios en HU-04 | Modelo actual sigue siendo agregado e insuficiente |
| `Remuneracion.Core/Models/ReversionR4.cs` | Sin cambios en HU-04 | Modelo actual sigue siendo agregado e insuficiente |

### 2.7 Visual Design Intent

No aplica. Esta HU no toca UI.

---

## 3. SPEC — Requisitos y aceptación

### 3.1 Requisitos

#### Requirement: Validación no destructiva del workbook

El sistema **MUST** inspeccionar el workbook sin sobrescribir fórmulas ni convertir celdas derivadas en valores fijos.

#### Scenario: Consolidado derivado intacto

- GIVEN un workbook real compatible
- WHEN se invoca `EscribirConsolidado(...)`
- THEN las celdas canónicas del consolidado permanecen como fórmulas
- AND la operación falla si alguna fue alterada

#### Requirement: Verificación de primera cadena upstream

El sistema **MUST** validar que las celdas resumen de R1, R2 y R4 también son derivadas por fórmula y **MUST NOT** tratarlas como inputs editables.

#### Scenario: R1/R2/R4 visibles pero no escribibles

- GIVEN el workbook real 202607-1
- WHEN se inspeccionan `F46`, `F48`, `E41` y `D67`
- THEN cada celda se reconoce como fórmula
- AND el writer rechaza escritura funcional sobre ellas

#### Requirement: Payload insuficiente debe fallar explícitamente

El sistema **MUST** reportar cuando el contrato actual no provee los leaf inputs requeridos para escritura real.

#### Scenario: Intento de escritura con agregados

- GIVEN `RecaudoComponenteR1`, `SaldosFavorR2` o `ReversionR4` actuales
- WHEN se solicita escribir el workbook real
- THEN se lanza una excepción descriptiva de payload insuficiente
- AND el archivo no se modifica destructivamente

#### Requirement: Resolución textual OpenXML segura

El sistema **MUST** resolver `SharedStrings` al buscar hojas, headers o labels.

#### Scenario: Texto almacenado como índice compartido

- GIVEN una celda con `SharedStringTable`
- WHEN el writer lee su texto visible
- THEN usa el valor resuelto y no el índice raw

### 3.2 Acceptance Criteria

| ID | Escenario | Resultado esperado |
|---|---|---|
| AC-1 | `EscribirConsolidado(...)` sobre workbook real | Valida fórmulas y no escribe valores directos |
| AC-2 | `EscribirDetalleR1(...)` sobre workbook real | Detecta `F46/F48` como fórmulas y rechaza escritura funcional |
| AC-3 | `EscribirDetalleR2(...)` sobre workbook real | Detecta `E41` como fórmula y rechaza escritura funcional |
| AC-4 | `EscribirDetalleR4(...)` sobre workbook real | Detecta `D67` como fórmula y rechaza escritura funcional |
| AC-5 | Workbook con fórmula reemplazada por valor | Se lanza excepción descriptiva |
| AC-6 | Matching textual con `SharedStrings` | La localización funciona correctamente |
| AC-7 | Intento de usar payload agregado para escritura real | Se informa prerequisito `WorkbookLeafInputs` |
| AC-8 | Build de solución | `dotnet build Remuneracion.WinForms/Remuneracion.WinForms.slnx` con 0 warnings |

### 3.3 Limitación explícita

**HU-04 NO certifica escritura funcional de R1/R2/R4**. Eso queda diferido hasta contar con:

- payload profundo por hoja,
- readers que lo alimenten,
- mapping semántico de verdaderos inputs editables,
- y aprobación de una historia nueva o expansión formal de HU-04.

---

## 4. TASKS — Desglose de implementación

## Review Workload Forecast

| Field | Value |
|---|---|
| Estimated changed lines | 170–280 |
| 400-line budget risk | Low |
| Chained PRs recommended | No |
| Suggested split | single PR |
| Delivery strategy | single-pr |
| Chain strategy | size-exception |

Decision needed before apply: Yes
Chained PRs recommended: No
Chain strategy: size-exception
400-line budget risk: Low

### Phase 1: Corrección semántica del writer

- [ ] 1.1 Eliminar de `OpenXmlPlantillaWriter.cs` toda suposición de escritura en `D9:D109`, `F46`, `F48`, `E41` y `D67`.
- [ ] 1.2 Centralizar helpers de lectura de fórmula, resolución de `SharedStrings` y validación de hoja/celda.
- [ ] 1.3 Documentar en código la semántica validation-only de HU-04.

### Phase 2: Validación estructural por hoja

- [ ] 2.1 Implementar verificación de fórmulas en `CONSOLIDADO_TOTAL RECAUDO`.
- [ ] 2.2 Implementar verificación de fórmulas profundas en `Reporte Componentes R1`.
- [ ] 2.3 Implementar verificación de fórmulas profundas en `Rem. Anticipos R2` y `Reversion Pagos R4`.
- [ ] 2.4 Hacer fail-fast descriptivo por payload insuficiente en `EscribirDetalleR1/R2/R4(...)`.

### Phase 3: Evidencia ejecutable

- [ ] 3.1 Actualizar `Herramientas/VerificadorRecaudo/Program.cs` para probar cadenas `D9→F46`, `D47→F48`, `D28→E41`, `D66→D67`.
- [ ] 3.2 Agregar casos negativos donde la fórmula sea reemplazada por valor y donde el writer reciba payload agregado insuficiente.
- [ ] 3.3 Ejecutar `dotnet build Remuneracion.WinForms/Remuneracion.WinForms.slnx` con 0 warnings.

### Phase 4: Prerequisito para historia futura

- [ ] 4.1 Documentar el contrato candidato `WorkbookLeafInputs` y sus submodelos por hoja.
- [ ] 4.2 Dejar trazadas las celdas/semánticas faltantes que deberá poblar HU futura de escritura real.

### Reglas de implementación

1. No sobrescribir nunca una celda con fórmula.
2. No llamar “escritura exitosa” a una validación estructural.
3. No prometer inputs reales donde hoy solo existen agregados.
4. `DetallePorComponente` se preserva como deuda útil, no como cobertura total.
5. No ignorar `SharedStrings`.

---

## 5. Architecture Validation Certificate

### 5.1 SOLID

| Principio | Cumplimiento |
|---|---|
| **S** — SRP | Infrastructure valida estructura OpenXML; Core sigue siendo dueño del dominio y del cálculo. |
| **O** — OCP | Se corrige la semántica sin romper el contrato público actual. |
| **L** — LSP | `OpenXmlPlantillaWriter` sigue cumpliendo `IPlantillaWriter`, ahora sin afirmar capacidades inexistentes. |
| **I** — ISP | No se agregan métodos ficticios; se explicita el contrato faltante en vez de inflar la interfaz actual. |
| **D** — DIP | Core permanece ajeno a OpenXML y al layout físico del workbook. |

### 5.2 Best Practices

- Prioriza verdad del workbook real sobre supuestos del código.
- Preserva fórmulas como regla máxima del negocio.
- Reduce alcance antes que mentir sobre capacidad técnica.
- Separa deuda útil (`DetallePorComponente`) de suficiencia contractual real.
- Usa evidencia verificable y fail-fast descriptivo.

### 5.3 Performance

- Validación focalizada en un conjunto chico de hojas y fórmulas canónicas.
- Sin recalcular el libro completo.
- Sin recorridos masivos innecesarios fuera de celdas/zonas trazadas.

**Veredicto:** **APROBADO COMO VALIDATION-ONLY.** No aprobado como writer funcional real con los contratos actuales.

---

## 6. Riesgos

| Riesgo | Likelihood | Mitigación |
|---|---|---|
| Volver a tratar fórmulas como celdas editables | Alta | Invalidar explícitamente `F46/F48/E41/D67` como targets de escritura |
| Confundir “valida” con “escribe” | Alta | Excepciones y documentación honestas; acceptance criteria alineados |
| Sobreestimar `DetallePorComponente` | Media | Declararlo útil pero insuficiente |
| Workbook cambie layout o referencias | Media | Validación estructural y fallas descriptivas |
| Ignorar `SharedStrings` y romper matching | Alta | Helper centralizado obligatorio |

---

## 7. Rollback Plan

- Revertir `OpenXmlPlantillaWriter.cs` y el verificador al estado anterior si la reformulación no es aprobada.
- No hay cambios de contrato público ni UI.
- El rollback queda acotado a Infrastructure + harness de verificación.

---

> **Control de cambios:** esta reformulación reemplaza otra premisa inválida. La nueva fuente de verdad es clara: el workbook real no solo tiene consolidado formula-driven, sino también resumen upstream formula-driven. Por eso HU-04 queda **reducida a validación OpenXML no destructiva** y la escritura funcional real se difiere hasta modelar los verdaderos `WorkbookLeafInputs`.
