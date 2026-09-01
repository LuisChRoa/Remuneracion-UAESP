# Plan 03 — HU-03: Motor de cálculo del CONSOLIDADO (1 ASE) + resultado global

> **Historia:** HU-03 — Motor de cálculo del CONSOLIDADO (1 ASE) + resultado global
> **Fuente rectora:** `Docs/Propuesta_Proyecto_Automatizacion_Remuneracion_UAESP.md` (§4.1, §6, §7, §9, §10) + `requirements/Fase1-Requerimientos.md` (HU-03, §6 reglas 5–9)
> **Estado:** PLAN — pendiente de aprobación para implementación
> **Fecha:** 2026-09-01
> **Alcance:** Implementar el motor de cálculo en `Remuneracion.Core` contra los contratos existentes de `ICalculoRemuneracion`, `ConsolidadoAse` y `ResultadoRemuneracion`. **NO** incluye escritura OpenXML ni materialización física de hojas.

---

## 0. Clarification Gate

| Pregunta / tensión real | Resolución del plan | Fuente de verificación |
|---|---|---|
| `requirements/Fase1-Requerimientos.md` amplía HU-03 a “CONSOLIDADO + hojas derivadas”, pero la **Propuesta** Fase 1 limita el prototipo a **1 ASE + CONSOLIDADO + hojas R1/R2/R4** | **Se adopta la Propuesta como rector arquitectónico.** HU-03 calculará los **valores canónicos de negocio** que alimentan el CONSOLIDADO (`TotOpt`, `R2TotalOportuno`, `Extemp`, `ReversionR4`, `AjustesSfT`, `TotalAse`, `GranTotal`). La materialización por hoja queda separada. | Propuesta §4.1, §4.2, §9 iteración 1.6 |
| El requerimiento menciona `DetRetri`, anticipos, ajustes y balance, pero los contratos actuales **no modelan payloads por hoja** | **No se expanden contratos en HU-03.** `DetRetri` se trata como **regla de negocio de redondeo** reutilizable desde Core, pero su escritura física queda en HU-04. Anticipos/balance detallado también quedan en HU-04/HU-05. | Contratos actuales + Propuesta §4.2 (“DetRetri fuera de alcance Fase 1”), §7 escritura |
| `ICalculoRemuneracion.Calcular(...)` no recibe `Periodo`; por sí solo no puede inferir lógica Q1/Q2 | `Calcular(...)` será una proyección pura R1/R2/R4 → `ConsolidadoAse`, con `AjustesSfT = 0m` en el alcance actual. La validación de quincena vive en `CalcularConsolidado(...)`, que sí recibe `Periodo`. | `ICalculoRemuneracion.cs` |
| ¿Quincena 2 debe “adivinarse” con `AjustesSfT = 0`? | **No.** Para evitar un cálculo silenciosamente incompleto, HU-03 queda **certificado para quincena 1**. Si `Periodo.NumeroQuincena == 2`, `CalcularConsolidado(...)` debe fallar con `CalculoInvalidoException` explicando que SALDOS POR NOTA / RETRIBUCION NEGATIVA no están modelados aún. | Propuesta §4.2, requisitos §6 regla 5, contratos sin inputs Q2 |
| ¿HU-03 debe escribir celdas D9:D109 o validar fórmulas? | **No.** HU-03 calcula valores de dominio; la asignación a celdas, preservación de fórmulas y hojas derivadas pertenece a `OpenXmlPlantillaWriter` en HU-04. | Propuesta §7 bloque de escritura, `IPlantillaWriter` |

**Decisión de alcance:** HU-03 implementa el **motor de cálculo canónico del CONSOLIDADO** y el **resultado agregado** usando los modelos actuales. **No cambia interfaces públicas** y **no mezcla concerns de escritura**. `DetRetri` queda preservado como regla de redondeo reusable, no como salida escrita.

Sin ambigüedades bloqueantes: el plan queda listo para implementación con ese recorte profesional de alcance.

---

## 1. PROPOSE — Propuesta

### 1.1 Por qué HU-03 es el siguiente paso crítico

HU-02 ya cerró la lectura real de R1/R2/R4 con evidencia PASS. El siguiente cuello de botella de Fase 1 ya no es “leer”, sino **transformar esos datos en los valores de negocio que el libro manual usa en el CONSOLIDADO**. Sin HU-03:

- no existe implementación concreta de `ICalculoRemuneracion`;
- no puede calcularse `TotalAse` ni `GranTotal` desde datos reales;
- HU-04 no tendría una fuente limpia y estable para escribir la plantilla.

### 1.2 Trazabilidad

| Fuente | Referencia | Aporte al plan |
|---|---|---|
| `Docs/Propuesta_Proyecto_Automatizacion_Remuneracion_UAESP.md` | §4.1 | El prototipo Fase 1 es **1 ASE** y se centra en cálculos del CONSOLIDADO |
| Propuesta | §4.2 | `DetRetri` automático y hojas Q2 quedan fuera de alcance de Fase 1 |
| Propuesta | §6.1 / §6.2 | La lógica de negocio vive en `Remuneracion.Core`; separación de capas |
| Propuesta | §7 | El flujo correcto es: leer fuentes → calcular dominio → escribir plantilla |
| Propuesta | §9 iteración 1.6 | HU-03 = “Motor de cálculo” |
| Propuesta | §10.1 CA-2 / CA-3 | Valores del CONSOLIDADO deben coincidir con cálculo manual conocido |
| `requirements/Fase1-Requerimientos.md` | HU-03 + §6 reglas 5–9 | Fórmulas esperadas para D85, D104, D109, `DetRetri`, tolerancia ±0.5 |

### 1.3 Alcance

**In (scope):**
- Implementación concreta de `ICalculoRemuneracion` en `Remuneracion.Core`.
- Cálculo canónico por ASE: `TotOpt`, `R2TotalOportuno`, `Extemp`, `ReversionR4`, `AjustesSfT`, `TotalAse`.
- Cálculo agregado: `ResultadoRemuneracion.Consolidados`, `GranTotal`, `Exitoso`, `Mensajes`, `Timestamp`.
- Regla de redondeo Excel-compatible para `DetRetri` como helper de dominio reutilizable.
- Verificación pragmática con harness temporal y fixture real Promoambiental 202607-1.

**Out of scope (explícitamente NO):**
- Escritura de celdas `D9:D109` y cualquier hoja del workbook (HU-04).
- `OpenXmlPlantillaWriter`.
- Implementación de `IValidador`.
- Inputs Q2 (`SALDOS POR NOTA`, `RETRIBUCION NEGATIVA`) y cálculos avanzados de balance/conciliación.
- UI, logging de WinForms y pruebas xUnit/golden formales (HU-05).

### 1.4 Criterios de éxito

- [ ] Existe una implementación concreta de `ICalculoRemuneracion` sin cambiar sus contratos públicos.
- [ ] Para Promoambiental 202607-1, el motor produce exactamente los valores canónicos esperados del CONSOLIDADO.
- [ ] `GranTotal` para resultado single-ASE coincide con `TotalAse` esperado: `39148067546.64m`.
- [ ] Quincena 2 no se procesa silenciosamente con datos incompletos: el motor falla de forma explícita.
- [ ] Build limpio con **0 warnings**.

### 1.5 Riesgos

| Riesgo | Prob. | Mitigación |
|---|---|---|
| Expandir HU-03 hasta mezclar cálculo + escritura | Alta | Mantener frontera Core vs Infrastructure documentada y verificable |
| Interpretar `DetRetri` como output del modelo y no como regla de dominio | Media | Preservar regla en helper de Core, sin tocar contratos existentes |
| Soportar Q2 “por default en 0” y producir resultados falsamente exitosos | Alta | Fail-fast en `CalcularConsolidado(...)` para quincena 2 |
| Repetir lógica de redondeo distinta a Excel | Media | Encapsular `ROUND(...,0)` con semántica compatible (`AwayFromZero`) |

---

## 2. DESIGN — Diseño técnico

### 2.1 Decisiones de diseño (registradas)

| # | Decisión | Justificación | Consecuencia |
|---|---|---|---|
| **D1** | Crear `Remuneracion.Core/Services/CalculoRemuneracion.cs` como implementación de `ICalculoRemuneracion` | No existe carpeta de servicios aún; la Propuesta §6.2 ubica explícitamente “Servicios de cálculo” en Core | Se crea `Services/` como primera convención de servicios de dominio |
| **D2** | **No cambiar** `ICalculoRemuneracion`, `ConsolidadoAse` ni `ResultadoRemuneracion` | Los contratos ya modelan exactamente el CONSOLIDADO canónico; expandirlos ahora mezclaría responsabilidades de hoja | Superficie mínima; HU-04 consumirá el resultado tal cual |
| **D3** | `Calcular(...)` será puro y determinista: mapea R1/R2/R4 a `ConsolidadoAse` | La firma no recibe `Periodo`; no debe inventar lógica contextual | `AjustesSfT = 0m` en esta HU y alcance certificado solo para Q1 |
| **D4** | `CalcularConsolidado(...)` valida la quincena y coordina el resultado global | Es el único método con `Periodo`; ahí debe vivir la regla de soporte Q1/Q2 | Quincena 2 → excepción explícita; quincena 1 → cálculo exitoso |
| **D5** | Preservar `DetRetri` con helper de dominio `Remuneracion.Core/Rules/DetRetriRounder.cs` | El redondeo es regla de negocio, no detalle de UI ni de OpenXML | HU-04 podrá reutilizar la misma regla sin duplicarla |
| **D6** | Reutilizar el harness temporal existente `Herramientas/VerificadorRecaudo` para HU-03 | Ya existe, ya apunta a fixtures reales y evita crear más ruido estructural antes de HU-05 | Se agregan casos de cálculo al `Program.cs`; xUnit sigue diferido |

### 2.2 Comportamiento esperado de la implementación

#### 2.2.1 `Calcular(Ase ase, RecaudoComponenteR1 r1, SaldosFavorR2 r2, ReversionR4 r4)`

Responsabilidad: construir el `ConsolidadoAse` de un ASE **sin side effects**.

Mapeo esperado:

```csharp
TotOpt            = r1.TotalOportuno;
R2TotalOportuno   = r2.TotalOportuno;      // GrandTotal - ServEspK (ya computado por el modelo)
Extemp            = r1.Extemporaneo;
ReversionR4       = r4.TotalReversiones;   // ya viene negativo; NO re-negar
AjustesSfT        = 0m;                    // alcance HU-03 / Q1
```

Validaciones:
- `ArgumentNullException` para `ase`, `r1`, `r2`, `r4` nulos.
- **No** validar escritura de celdas ni fórmulas aquí.
- **No** intentar inferir Q2; ese contexto no existe en esta firma.

#### 2.2.2 `CalcularConsolidado(Periodo periodo, List<(Ase ase, RecaudoComponenteR1 r1, SaldosFavorR2 r2, ReversionR4 r4)> datos)`

Responsabilidad: orquestar el cálculo agregado y proteger el alcance real.

Validaciones mínimas:
- `ArgumentNullException` para `periodo` / `datos` nulos.
- `CalculoInvalidoException` si `datos.Count == 0`.
- `CalculoInvalidoException` si hay ASE duplicados en la colección.
- `CalculoInvalidoException` si `NumeroQuincena` no es 1 ni 2.
- `CalculoInvalidoException` si `NumeroQuincena == 2`, porque HU-03 no recibe aún inputs de ajustes Q2.

Flujo:
1. Validar inputs y alcance.
2. Ejecutar `Calcular(...)` por cada tupla.
3. Construir `ResultadoRemuneracion`.
4. Dejar `GranTotal` como getter calculado del modelo.

Asignación de estado:
- `Exitoso = true` cuando el método completa sin excepción.
- `Mensajes = []` en éxito limpio; no usarlo para ocultar fallos de negocio.
- `Timestamp = DateTime.Now` explícito para dejar evidencia del momento de cálculo y ser coherente con el modelo actual.

### 2.3 `AjustesSfT` en Q1 vs Q2

| Contexto | Comportamiento planificado |
|---|---|
| **Quincena 1** | `AjustesSfT = 0m` por regla actual y por ausencia legítima de insumos Q2 |
| **Quincena 2** | **No soportado en HU-03**. Debe fallar explícitamente porque el contrato no recibe `SALDOS POR NOTA` ni `RETRIBUCION NEGATIVA` |

### 2.4 Regla `DetRetri` sin mezclar escritura

La regla **sí pertenece al dominio**, pero la escritura de la hoja **no**. Por eso el plan propone:

- helper puro en Core: `DetRetriRounder.Round(decimal totalAse)`;
- semántica compatible con Excel: `ROUND(valor, 0)` → `MidpointRounding.AwayFromZero`;
- **sin** agregar propiedad `DetRetri` al modelo `ConsolidadoAse`;
- HU-04 reutiliza ese helper al escribir la columna D de la hoja `DetRetri`.

### 2.5 Estructura resultante

```
Remuneracion.Core
├── Interfaces/
│   └── ICalculoRemuneracion.cs          // sin cambios
├── Models/
│   ├── ConsolidadoAse.cs                // sin cambios
│   └── ResultadoRemuneracion.cs         // sin cambios
├── Rules/
│   └── DetRetriRounder.cs               // nuevo
└── Services/
    └── CalculoRemuneracion.cs           // nuevo
```

### 2.6 Visual Design Intent

No aplica: HU-03 no contiene UI ni cambios visuales.

---

## 3. Estrategia de pruebas

> No hay proyecto xUnit todavía. Se usa evidencia ejecutable pragmática, igual que HU-02, reutilizando el harness temporal ya existente hasta que HU-05 formalice pruebas automáticas.

### 3.1 Estrategia elegida

- **Reutilizar** `Herramientas/VerificadorRecaudo/Program.cs`.
- Agregar bloque HU-03 que:
  1. lea R1/R2/R4 reales con el reader ya validado;
  2. instancie `CalculoRemuneracion`;
  3. ejecute `Calcular(...)` y `CalcularConsolidado(...)`;
  4. compare contra los valores reales conocidos del período `202607-1`.

**Por qué esta estrategia:** da evidencia inmediata con archivos reales, evita inventar datos sintéticos antes de HU-05 y no obliga a abrir un proyecto de tests todavía.

### 3.2 Valores de referencia (Promoambiental 202607-1)

| Caso | Valor esperado |
|---|---:|
| `TotOpt` | `19556118465.99m` |
| `R2TotalOportuno` | `54216385.68m` |
| `Extemp` | `19549786950.62m` |
| `ReversionR4` | `-12054255.65m` |
| `AjustesSfT` | `0m` |
| `TotalAse` | `39148067546.64m` |
| `GranTotal` (single ASE) | `39148067546.64m` |
| `DetRetri` redondeado | `39148067547m` |

### 3.3 Casos mínimos del harness

| ID | Escenario | Verificación |
|---|---|---|
| T6 | `Calcular(...)` con fixture Promoambiental 202607-1 | Todos los campos del `ConsolidadoAse` coinciden con la tabla 3.2 |
| T7 | `CalcularConsolidado(...)` con 1 ASE | `Consolidados.Count = 1`, `GranTotal = 39148067546.64m`, `Exitoso = true`, `Mensajes` vacío |
| T8 | `DetRetriRounder.Round(39148067546.64m)` | Devuelve `39148067547m` |
| T9 | `CalcularConsolidado(...)` con quincena 2 | Lanza `CalculoInvalidoException` explicando alcance no soportado |
| T10 | `CalcularConsolidado(...)` con lista vacía o ASE duplicado | Lanza `CalculoInvalidoException` |

### 3.4 Criterio de comparación

- Para el **servicio puro**: comparación exacta de `decimal`, porque la lógica es suma/resta directa.
- Para la **ruta end-to-end con archivos reales**: se mantiene como respaldo la tolerancia de negocio **±0.5**.

Comando de evidencia:

```bash
dotnet run --project Herramientas/VerificadorRecaudo/Herramientas.VerificadorRecaudo.csproj
```

---

## 4. SPEC — Criterios de aceptación y trazabilidad

### 4.1 Escenarios de aceptación HU-03

| ID | Escenario (dado/cuando/entonces) | Traza |
|---|---|---|
| AC-1 | Dado un `Ase` y datos válidos R1/R2/R4, cuando `Calcular(...)`, entonces `ConsolidadoAse` refleja `TotOpt`, `R2TotalOportuno`, `Extemp`, `ReversionR4` y `AjustesSfT=0` | Requisitos HU-03, reglas 5–6; Propuesta §7 |
| AC-2 | Dado el fixture Promoambiental 202607-1, cuando se calcula el consolidado del ASE, entonces `TotalAse = 39148067546.64m` | CA-2 general; reglas 1–6 |
| AC-3 | Dado un período de quincena 1 y una lista con 1 ASE, cuando `CalcularConsolidado(...)`, entonces `GranTotal = 39148067546.64m`, `Exitoso = true` y `Mensajes` queda vacío | Propuesta §10.1 CA-2 |
| AC-4 | Dado un período de quincena 2, cuando `CalcularConsolidado(...)`, entonces el motor rechaza el cálculo con `CalculoInvalidoException` por falta de inputs Q2 | Propuesta §4.2; requisitos §6 regla 5 |
| AC-5 | Dado `DetRetri`, cuando se redondea el total del ASE, entonces se aplica semántica equivalente a `ROUND(valor, 0)` de Excel | Requisitos HU-03; regla 7 |
| AC-6 | Dado `datos` vacío o con ASE duplicados, cuando `CalcularConsolidado(...)`, entonces el motor falla explícitamente | Robustez de Core; alcance agregado |
| AC-7 | El motor no escribe celdas ni depende de OpenXML / WinForms | Propuesta §6.1, §7 |
| AC-8 | La solución compila con **0 warnings** y el harness de HU-03 pasa | Convención del proyecto |

### 4.2 Matriz de trazabilidad

| Entregable | Requisito | Propuesta | Evidencia |
|---|---|---|---|
| Implementación `ICalculoRemuneracion` | HU-03 | §6.2 “Servicios de cálculo”, §9 iteración 1.6 | T6, T7 |
| Regla `AjustesSfT = 0` en Q1 | Regla 5 | §4.1 / §7 | T6 |
| `TotalAse` y `GranTotal` | Regla 6 | §7, §10.1 CA-2 | T6, T7 |
| Regla `DetRetri` | Regla 7 | §7 bloque DetRetri | T8 |
| Fail-fast Q2 | Regla 5 + tensión contractual | §4.2 fuera de alcance | T9 |

---

## 5. TASKS — Desglose de implementación

### 5.1 Forecast

- **Superficie estimada:** ~140–240 líneas cambiadas.
- **Riesgo 400-line budget:** Bajo.
- **Chained PRs:** No necesario.

### 5.2 Tareas

| # | Tarea | Depende de | Archivos | Criterio de salida |
|---|---|---|---|---|
| T1 | Crear carpeta `Remuneracion.Core/Services/` e implementar `CalculoRemuneracion.cs` contra `ICalculoRemuneracion` | — | `Remuneracion.Core/Services/CalculoRemuneracion.cs` | Compila y no cambia contratos públicos |
| T2 | Implementar `Calcular(...)` como proyección pura R1/R2/R4 → `ConsolidadoAse` | T1 | ídem | `TotOpt`, `R2TotalOportuno`, `Extemp`, `ReversionR4`, `AjustesSfT=0` correctos |
| T3 | Implementar `CalcularConsolidado(...)` con validaciones de null, lista vacía, duplicados e invalidación Q2 | T1 | ídem | Retorna `ResultadoRemuneracion` exitoso en Q1 y falla explícito en Q2 |
| T4 | Crear helper `DetRetriRounder` con semántica Excel-compatible (`AwayFromZero`) | T1 | `Remuneracion.Core/Rules/DetRetriRounder.cs` | Redondea `39148067546.64m` a `39148067547m` |
| T5 | Extender harness temporal para cubrir HU-03 con fixture real Promoambiental 202607-1 | T2, T3, T4 | `Herramientas/VerificadorRecaudo/Program.cs` | Casos T6–T10 PASS |
| T6 | Ejecutar build limpio y documentar evidencia del harness | T5 | `Remuneracion.WinForms/Remuneracion.WinForms.slnx`, `Herramientas/VerificadorRecaudo/` | `dotnet build` 0 warnings + `dotnet run` PASS |

**Secuencia:** `T1 → (T2, T3, T4) → T5 → T6`.

### 5.3 Reglas de implementación

1. **No tocar** `ICalculoRemuneracion`, `ConsolidadoAse`, `ResultadoRemuneracion` salvo que aparezca una incompatibilidad real verificada.
2. **No mezclar** lógica OpenXML o referencias WinForms en Core.
3. Quincena 2 **no se simula** ni se “completa en cero” silenciosamente.
4. `ReversionR4` se usa **tal como viene**: negativa.
5. `DetRetri` se preserva como **regla de negocio reusable**, no como escritura.
6. **0 warnings** al cerrar la HU.

---

## 6. Architecture Validation Certificate

### 6.1 SOLID

| Principio | Cumplimiento |
|---|---|
| **S** — SRP | `CalculoRemuneracion` calcula dominio; `DetRetriRounder` redondea; el writer seguirá escribiendo en HU-04 |
| **O** — OCP | La implementación agrega servicios sin romper contratos existentes |
| **L** — LSP | `CalculoRemuneracion` sustituye correctamente a `ICalculoRemuneracion` |
| **I** — ISP | No se ensancha la interfaz con responsabilidades de hoja o writer |
| **D** — DIP | `WinForms`/`Infrastructure` dependen del contrato Core; Core no depende de OpenXML ni UI |

### 6.2 Best Practices

- Implementación **contract-first** y con superficie mínima.
- Validación explícita de inputs inválidos y de alcance no soportado.
- `DateTime.Now` coherente con el modelo actual `ResultadoRemuneracion.Timestamp`.
- Regla Excel de redondeo centralizada en un helper único.
- Verificación con fixture real antes de declarar la HU lista.

### 6.3 Performance

- Operaciones puramente en memoria sobre `decimal`; costo O(n) sobre la lista de ASE.
- Para Fase 1 (1 ASE), el costo es despreciable.
- No se agregan dependencias ni capas intermedias innecesarias.

**Veredicto de arquitectura:** **APROBADO.** El diseño respeta el corte por capas, mantiene el Core limpio y resuelve la tensión de alcance SIN contaminar HU-03 con responsabilidades de escritura.

---

## 7. Rollback

- El cambio es local a `Remuneracion.Core` y al harness temporal.
- Revertir = eliminar `Services/CalculoRemuneracion.cs`, `Rules/DetRetriRounder.cs` y los casos HU-03 agregados al verificador.
- Como no hay cambios de contrato ni de esquema persistente, el rollback no impacta UI, Infrastructure ni archivos Excel.

---

> **Control de cambios:** este plan cubre exclusivamente el motor de cálculo canónico del CONSOLIDADO para Fase 1 (1 ASE) y el resultado agregado. Escritura OpenXML, hojas derivadas físicas, `DetRetri` en workbook, validaciones cruzadas y soporte real de quincena 2 corresponden a HU-04/HU-05.
