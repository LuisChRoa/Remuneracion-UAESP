# Plan 13 — HU-13: Validaciones cruzadas (Fase 2, it. 2.7 — última de Fase 2)

> **Historia:** implementar la lógica de validaciones cruzadas en dominio (`IValidador`/Core) como ORÁCULO DE LECTURA: se validan, no se escriben — las hojas de validación siguen protegidas. Cierra Fase 2 (it. 2.7) y absorbe la deuda HU-12 (W1 + W2).
> **Rector (IRRENUNCIABLE):** `Docs/Propuesta_Proyecto_Automatizacion_Remuneracion_UAESP.md` — §9 Fase 2 it. **2.7 validaciones cruzadas**, §6 arquitectura (capas, preservación de fórmulas, Serilog, lectura directa de fuente), §10 CA + tolerancia ±0.5. **Entra SOLO 2.7.** Nada de Fase 3 (CLI, manual) entra aquí.
> **Origen funcional:** `Detalle de plantilla.docx` + `Prompt Maestro Vo.docx` en lo que digan de validaciones / `DetValiRetri` / `VALIDACION_*`. `Proceso de Recaudo.docx` **NO aplica** salvo cita directa que lo exija.
> **Continuidad:** HU-01..HU-12 cerradas (ruta leaf multi-ASE atómica single-write, mapas explícitos por `Ase.Id` + período, validador estricto, UI modo 5-ASE, Golden Capa A, build 0/0, tests 128/128 —HU-01..HU-12— + harness 24/24 regresión ciega). Este plan NO reabre su semántica.
> **Estado actual:** las hojas de validación hoy están PROTEGIDAS (nunca escritas, nunca implementada su lógica): `VALIDACION_TOTAL`, `VALIDACION_RECIP/ENEL/ENERBIT/OCCIDENTE/EAAB-CL` (O3=H3-N3, P3=INT(O3)=0), `DetValiRetri2026071/2026072` (D16:D21 = D9:D14 vs CONSOLIDADO U104:U109; D24:D29 con refs M/L/J/AJUSTES/U por ASE — ver `WorkbookLeafCellMapQ2.cs:257-268`), `GERENTES_*`, `Valida -Remunera` / `-Anticipos` / `-Control Recaudo`. `DetRetri` J9:J14 y `DetValiRetri` D9:D14 del golden tienen valores ≠ ROUND (hallazgo HU-12, fuera de escritura).
> **Estado:** PENDIENTE DE APROBACIÓN — no implementar hasta OK del Ingeniero.
> **Fecha:** 2026-09-09

---

## 0. Clarification Gate

No hay pregunta bloqueante para el Ingeniero. La única incógnita real —la semántica exacta de cada validación (qué compara contra qué, en qué hojas, en Q1 y en Q2)— **no se asume**: la dicta el T0 contra ambos canónicos (§4 Fase 0), con veredicto por validación. El apply espera aprobación explícita.

### 0.1 Qué está verificado y qué NO (honestidad técnica)

**Verificado en esta planificación** (código leído + contexto del orquestador):

| # | Hecho | Método | Consecuencia |
|---|---|---|---|
| V1 | `VALIDACION_*` compara Recaudo * vs `REMUNERACION_*` por empresa con patrón `O3=H3-N3`, `P3=INT(O3)=0` (diferencia debe ser 0) | Plan 08 V4 (dump `<f>` sheet25) | Semántica esperada del gate por empresa; T0 la confirma celda-por-celda en ambos períodos |
| V2 | `DetValiRetri` Q2 protegido ya mapeado: D16:D21 = D9:D14 vs `CONSOLIDADO_TOTAL RECAUDO` U104:U109; D24:D29 composición M/L/J/AJUSTES/U por ASE; D29 agregada multi-bloque | `WorkbookLeafCellMapQ2.cs:257-268` (leído) | Base del oráculo DetValiRetri-Q2; falta el espejo Q1 + verificación de existencia explícita (W2) |
| V3 | `IValidador` tiene 3 sobrecargas (single, leaf, multi-ASE lista); `ValidadorBasico` usa matcheo estricto `Single` por `Ase.Id`, tolerancia ±0.5, y NO abre `.xlsx` | Lectura `IValidador.cs` + `ValidadorBasico.cs:1-80` | 2.7 = **extensión por overload/gate nuevo**, no reescritura; el oráculo `.xlsx` vive en Infrastructure (read-only), nunca en el validador |
| V4 | `DetRetriRounder.Round` existe (`Core/Rules/DetRetriRounder.cs`, ROUND/AwayFromZero, puro) | Plan 12 V0.7 | Reutilizado para el gate DetValiRetri-D; prohibido duplicar la regla |
| V5 | Mapas hermanos por período existen: `WorkbookLeafCellMapPorAse` (Q1), `WorkbookLeafCellMapQ2` (Q2), `PorEmpresa`, `AjustesSfT`, `ReporteBanco`, `BalanceSc` | Glob Infrastructure/Excel | 2.7 = **nuevo mapa de protegidas/oráculo**, no parche de los existentes; Q1/Q2 intactos por construcción |
| V6 | Modelos de dominio por bloque existen (`ConciliacionEmpresaInputs`, `ReporteBancoInputs`, `BalanceScInputs`, `AjustesSfTInputs`, `DetRetriQ2Inputs`) + `WorkbookLeafInputs` extensible por listas/nullables | Glob Core/Models | Patrón de extensión aditiva para el snapshot de validación (nullable/default = HU-12 intacto) |
| V7 | Deuda HU-12 declarada: W1 (negativas §4.5 del Plan 12 + guarda de capacidad `DetRetri` D9:D14 ante plantilla futura) y W2 (asserts de existencia explícita en T0) | Revisión HU-12 (orquestador) | W1+W2 se absorben en este plan (§2.9, §4 Fases 4–5); sin rebase |
| V8 | Corrección menor del Plan 12: "sheets 93/94" → verdad **36/37** | Revisión HU-12 (orquestador) | Este plan usa 36/37 en todo mapa/test nuevo; el Plan 12 no se reescribe (se registra la fe de erratas en §9) |
| V9 | `DetRetri` J9:J14 y `DetValiRetri` D9:D14 del golden tienen valores ≠ ROUND (hallazgo HU-12, fuera de escritura) | Orquestador | El gate 2.7 **MUST NOT** exigir ROUND en J9:J14/D9:D14; se trata como divergencia conocida documentada (ver D6) |

**NO verificado (y por eso T0 es bloqueante, §4 Fase 0):**

1. Semántica exacta por validación y por período: `VALIDACION_RECIP/ENEL/ENERBIT/OCCIDENTE/EAAB-CL` (¿O3/P3 en ambas quincenas con mismas refs H/N? ¿qué filas por ASE?).
2. `VALIDACION_TOTAL`: layout, fórmula por celda, qué agrega (¿Σ empresas? ¿Σ ASE? ¿gran total?).
3. `DetValiRetri` Q1: espejo de `WorkbookLeafCellMapQ2.cs:257-268` (¿mismas refs M/L/J/AJUSTES/U o variantes Q1? ¿U104:U109 en Q1?).
4. `Valida -Remunera` / `-Anticipos` / `-Control Recaudo`: layout por hoja (¿una fila por ASE? ¿qué columnas comparan? ¿booleanas o diferencias?).
5. `GERENTES_*`: confirmar fórmulas puras en ambos canónicos (pendiente de Plan 08, T0-0.5).
6. Cuantificación exacta de la divergencia J9:J14/D9:D14 vs ROUND por ASE y período (para el veredicto D6, no para "arreglarla").
7. Si alguna celda oráculo depende de un externalLink roto.
8. Barrido de existencia explícita (W2): cada hoja/celda del oráculo debe assertar existencia antes de comparar.

> **Regla de hierro del plan:** ninguna semántica de validación entra al código sin pasar por T0 contra ambos canónicos. Lo ya verificado arriba (V1–V9) sí es contratable desde el día uno. Si una validación no cierra ±0.5 contra el golden: NEEDS_CONTEXT con recorte, nunca invención.

### 0.2 Mapeo al Rector (in vs out)

**Entra porque §9 it. 2.7 / §6 / §10 lo piden ahora:**

| Rector | Qué cubre HU-13 |
|---|---|
| §9 it. 2.7 | Lógica de validaciones cruzadas en dominio: Recaudo * vs `REMUNERACION_*` por empresa (= 0, patrón O3/P3), `DetValiRetri` D16:D21/D24:D29 vs dominio ±0.5, `VALIDACION_TOTAL`, `Valida -*` |
| §6 | Oráculo de lectura (fórmulas/visibles se validan, no se escriben; hojas siguen protegidas); Serilog; lectura directa |
| §10 CA-1/CA-2 | Leer los visibles/fórmulas-oráculo por ASE y período + coherencia vs dominio ±0.5 |
| §10 CA-3/CA-4 | Golden Capa A de validaciones en ambos períodos; fórmulas intactas (mapa protegido extendido) |
| §10 CA-5/CA-6/CA-7 | Gates por ASE con matcheo estricto, Serilog por validación, UI sin cambios salvo resumen |

**Sale porque §9 lo asigna a Fase 3 u otras HUs (EXPLÍCITO):**

- Escritura de hojas de validación (siguen protegidas — escribirlas es prohibido por test).
- `INTERVENTORIA` (sin HU asignada — protegida, no se toca).
- `ANT EXT-REV` (protegida, no se toca).
- 2.6 ya cerrada (cadena DetRetri certificada, no se reabre).
- Fase 3 (CLI, manual de usuario).
- Insert/delete de filas (prohibidos; si el oráculo no alcanza → NEEDS_CONTEXT, nunca filas nuevas).

### 0.3 Decisiones ejecutivas (aprobación = aceptar estas)

| ID | Decisión |
|---|---|
| G1 | **T0-evidencia primero y T0 bloquea el mapa**: la semántica exacta de cada validación la dicta T0 contra ambos canónicos, nunca se asume (§4 Fase 0). |
| G2 | **Oráculo de lectura, cero escritura**: las hojas de validación entran al mapa protegido extendido; el writer falla si alguna deja de ser fórmula. La lógica vive en Core como gates de comparación. |
| G3 | **PRs encadenados con T0 al frente** (§4 Work Units): Unidad 0 (T0 + oráculo congelado) → dominio → oracle-reader → coherencia + golden → W1+W2 + regresión. |
| G4 | **Golden honesto Capa A** en ambos períodos (leaf salida vs leaf golden; visibles de dominio vs caché golden; nunca caché de salida vs golden). Capa B manual residual del usuario. |
| G5 | **Matcheo estricto por `Ase.Id`** (`Single`, nunca fallback) en todo gate nuevo; Q1/Q2 intactos por construcción (dispatch por `NumeroQuincena`). |
| G6 | **Deuda HU-12 absorbida aquí**: W1 (negativas §4.5 Plan 12 + guarda de capacidad `DetRetri` D9:D14) + W2 (asserts de existencia explícita en T0). |
| G7 | **Fe de erratas Plan 12**: "sheets 93/94" → verdad 36/37; este plan usa 36/37 en todo artefacto nuevo. |
| G8 | **Si una validación no cierra ±0.5**: NEEDS_CONTEXT con recorte declarado, nunca invención. |

---

## 1. PROPOSE

### 1.1 Intent

Cerrar Fase 2 implementando la lógica de validaciones cruzadas como gates de dominio que comparan el cálculo propio contra el oráculo del workbook (fórmulas/visibles leídos read-only), certificando Recaudo-vs-Remuneración por empresa, `DetValiRetri`, `VALIDACION_TOTAL` y `Valida -*` en ambos períodos — sin escribir jamás una hoja de validación y sin mover una coma de HU-01..HU-12.

### 1.2 In Scope

- Discovery T0 bloqueante (§4 Fase 0) + `WorkbookLeafCellMapValidaciones` (oráculo + protegidas) congelado solo con evidencia en ambos canónicos.
- Modelos snapshot-oráculo (`ValidacionCruzadaSnapshot` por ASE/período) + reader read-only `IValidacionOracleReader` (Infrastructure, OpenXML sin recalcular).
- Gates 2.7 en `IValidador`/`ValidadorBasico` + `WorkbookLeafCoherence` (por empresa=O3/P3, DetValiRetri D16:D21/D24:D29, VALIDACION_TOTAL, Valida -*), tolerancia ±0.5, matcheo estricto.
- Mapa protegido extendido (todas las hojas de validación; el writer las rechaza como destino).
- W1: tests negativos §4.5 Plan 12 + guarda de capacidad `DetRetri` D9:D14 ante plantilla futura.
- W2: asserts de existencia explícita en T0 (cada hoja/celda oráculo assertada antes de comparar).
- Golden Capa A de validaciones en ambos períodos + regresión 128/128 + harness 24/24 intacta.
- Resumen de validaciones por ASE en log/Serilog (delta mínimo UI).

### 1.3 Out of Scope

Todo §0.2 (escritura de validaciones, INTERVENTORIA, ANT EXT-REV, Fase 3, insert/delete). Además: reescritura HU-08 en Q2 (recorte 1 HU-11, se mantiene); cadena 2.5/2.6 (certificadas, no se reabren); motor Excel/COM en CI; restyle UI; framework DI; "corregir" la divergencia J9:J14/D9:D14 vs ROUND (hallazgo documentado, fuera de escritura).

### 1.4 Resultado de negocio

El Ingeniero ejecuta el modo 5 ASE como hoy sobre Q1 o Q2; además de la salida certificada HU-12, obtiene un veredicto de validaciones cruzadas por ASE y período (qué validación cierra = 0 ±0.5, cuál diverge y con qué recorte); las hojas de validación siguen intactas como fórmulas; Fase 2 queda cerrada y la deuda W1/W2 saldada; el log audita cada validación.

### 1.5 Base documental (origen funcional — citas por sección)

| Documento | Sección / instrucción | Qué aporta a 2.7 |
|---|---|---|
| `Detalle de plantilla.docx` | Inst. hojas de validación / `DetValiRetri` / `VALIDACION_*` | Layout objetivo y semántica de cada validación; base del mapa T0 |
| `Prompt Maestro Vo.docx` | Pasos de validación por empresa y de DetValiRetri | Qué se compara contra qué (oráculo esperado); base de los gates |
| Rector Propuesta | §9 it. 2.7 (alcance), §6 (preservación, lectura directa, trazabilidad), §10 CA + tolerancia ±0.5 | Rector normativo |
| `plans/08` V4 | `O3=H3-N3`, `P3=INT(O3)=0` en `VALIDACION_*` | Patrón del gate por empresa (a confirmar por T0 en ambos períodos) |
| `WorkbookLeafCellMapQ2.cs:257-268` | Fragmentos DetValiRetri D16:D21/D24:D29 Q2 | Base del oráculo DetValiRetri-Q2; el espejo Q1 lo dicta T0 |
| `Proceso de Recaudo.docx` | — | **NO aplica** salvo cita directa que lo exija |

---

## 2. DESIGN

### 2.1 Tablas verificadas (contratables desde el día uno)

**Oráculo esperado por validación (forma; semántica exacta la congela T0):**

| Validación | Hojas | Patrón esperado | Gate de dominio |
|---|---|---|---|
| Empresa ×5 | `VALIDACION_RECIP/ENEL/ENERBIT/OCCIDENTE/EAAB-CL` | `O3=H3-N3` (Recaudo * vs `REMUNERACION_*`), `P3=INT(O3)=0` | O3 = 0 ±0.5 y P3 = 0 exacto por empresa (5 empresas; ceros legítimos donde aplique) |
| DetValiRetri | `DetValiRetri2026071/2026072` | D16:D21 = D9:D14 vs CONSOLIDADO U104:U109; D24:D29 composición M/L/J/AJUSTES/U por ASE (Q2 en `WorkbookLeafCellMapQ2.cs:257-268`) | D16:D21 = 0 ±0.5 por ASE; D24:D29 verificado según semántica T0 (booleana o diferencia) |
| Total | `VALIDACION_TOTAL` | Agregación (T0: ¿Σ empresas? ¿Σ ASE? ¿gran total?) | Total = 0 ±0.5 o composición T0 exacta |
| Control ×3 | `Valida -Remunera` / `-Anticipos` / `-Control Recaudo` | Layout T0 (¿fila por ASE? ¿columna por concepto?) | Cada control = 0 ±0.5 o booleano según T0 |
| Gerentes | `GERENTES_*` ×5 | Fórmulas puras (esperado, a confirmar T0) | Protegidas (fallan la escritura si dejan de ser fórmula); sin gate numérico salvo que T0 demuestre composición comparable |

**Contratos que NO cambian:**

| Elemento | Estado | Tratamiento 2.7 |
|---|---|---|
| `IValidador` 3 sobrecargas + `ValidadorBasico` (`Single` por Id, ±0.5, NO abre `.xlsx`) | Certificado HU-01..HU-12 | Intacto; 2.7 suma overload/gate con snapshot-oráculo (contrato aditivo) |
| `DetRetriRounder.Round` (ROUND/AwayFromZero, puro) | Existe | Reutilizado en el gate DetValiRetri-D; prohibido duplicar |
| Mapas Q1/Q2/Empresa/Ajustes/Banco/Balance + writer single-write + hash A4 | Certificados | Intactos; 2.7 agrega mapa-oráculo hermano, nunca los parchea |
| 128 tests + harness 24/24 | Verdes | Red de regresión ciega a 2.7 |
| J9:J14/D9:D14 ≠ ROUND (hallazgo HU-12) | Documentado | Divergencia conocida: el gate la excluye y la documenta (D6); prohibido "cerrarla" inventando origen |

### 2.2 Decisiones de arquitectura

| ID | Opción elegida | Descartada | Por qué |
|---|---|---|---|
| D1 | Snapshot-oráculo como modelo de dominio (`ValidacionCruzadaSnapshot` por ASE: O3/P3 por empresa, D16:D21, D24:D29, total, controles) + reader read-only en Infrastructure | Que `ValidadorBasico` abra el `.xlsx` | El validador NO abre `.xlsx` (doctrina HU-06..HU-12). La lectura vive en Infrastructure; Core solo compara números (DIP, testeabilidad in-memory) |
| D2 | `WorkbookLeafCellMapValidaciones`: mapa hermano explícito por (`Ase.Id`, período, validación) — refs oráculo + fragmentos de fórmula protegida | Reutilizar mapas BCE/Ajustes/Q2 como oráculo | Cada validación tiene su propia aritmética y sus propias hojas; mezclarlas rompería Q1/Q2 y el hash A4 |
| D3 | Gates 2.7 aditivos por quincena: si snapshot es null → comportamiento HU-12 intacto (sin gates nuevos); si presente → gates T0 por período | Unificar reglas sin distinguir quincena/período | Q1/Q2 tienen layouts-oráculo potencialmente distintos (lección Plan 12 v1: F53≠F46). Dispatch por `NumeroQuincena` blinda la regresión |
| D4 | Validación protegida extendida: TODAS las hojas de validación al mapa protegido (el writer falla si alguna deja de ser fórmula o si se intenta escribir en ellas) | Proteger solo las tocadas por 2.7 | Blindan Fase 3 contra escritura accidental desde cualquier HU; costo cero en runtime |
| D5 | T0 con doble veredicto por validación: (a) cierra ±0.5 → gate activo; (b) no cierra → NEEDS_CONTEXT con recorte declarado, la validación queda en "divergencia conocida" sin bloquear el resto | Forzar todas las validaciones a gate activo | Honestidad §0.1: si una validación no cierra, inventar su semántica rompería el golden. El recorte es el mecanismo honesto (precedente: recorte 1 HU-11) |
| D6 | Divergencia J9:J14/D9:D14 vs ROUND: excluida de los gates, cuantificada en T0 y documentada como hallazgo (test que demuestra la divergencia y la excluye a propósito) | Exigir ROUND en J9:J14/D9:D14 o "corregir" el golden | El hallazgo HU-12 es fuera de escritura; convertirlo en gate rompería 128/128 contra el golden real |
| D7 | W2 como asserts de existencia en el propio mapa-oráculo: cada entrada (hoja, celda) assertada existente antes de comparar (hoja existe, celda existe, `<f>` presente donde T0 lo exige) | Existencia implícita (leer y comparar directo) | La deuda W2 es exactamente esa: sin asserts explícitos, una plantilla futura con hoja faltante daría "0" silencioso en vez de fallo nombrado |
| D8 | W1-guarda de capacidad: `DetRetri` D9:D14 con guarda de rango (si la plantilla futura trae más/menos filas ASE, fail-fast que nombra la hoja, nunca truncado silencioso) + negativas §4.5 Plan 12 portadas como tests de este plan | Dejar W1 para Fase 3 | La deuda es de HU-12 y esta es la última HU de Fase 2: Fase 2 no cierra con deuda abierta |

### 2.3 Flujo 2.7 (sin escritura nueva — solo lectura-oráculo + gates)

```text
File.Copy plantilla → salida (una vez, igual que HU-07..HU-12)
  └─► ValidarFormulasProtegidas (mapa HU-07..HU-12 + mapa 2.7: VALIDACION_*×5,
      VALIDACION_TOTAL, DetValiRetri Q1/Q2, Valida -*, GERENTES_* — todas protegidas)
        └─► EscribirCeldasLeaf HU-07..HU-12 (intactas; solo activas según período)
              └─► revalidar fórmulas protegidas → guardar
Lectura-oráculo (read-only, post-guardado o sobre el golden según test):
  IValidacionOracleReader.LeerSnapshot(rutaWorkbook, periodo)
    └─► ValidadorBasico.Validar(resultado, leafs, snapshot) — gates D5 por validación
          └─► veredicto por ASE × validación (cierra / diverge-con-recorte / falla-nombrando)
```

Ante cualquier fallo de escritura: borrar salida parcial (patrón existente). Plantilla origen jamás mutada (hash A4 extendido al mapa 2.7). Sin capacidad de filas en el oráculo: NEEDS_CONTEXT honesto (D5b), nunca insert/delete.

### 2.4 Dominio (Core, sin deps)

```csharp
public sealed class ValidacionEmpresaSnapshot  // una empresa: Recaudo * vs REMUNERACION_*
{
    public string Empresa { get; set; } = "";  // label catálogo EmpresaFacturacion
    public decimal DiferenciaO3 { get; set; }  // H3-N3 leído del caché-oráculo
    public decimal VerificacionP3 { get; set; } // INT(O3) leído (debe ser 0)
}

public sealed class DetValiRetriSnapshot  // un ASE
{
    public Ase Ase { get; set; } = new();
    public IReadOnlyDictionary<string, decimal> D16_D21 { get; set; } = ...; // celda→valor caché
    public IReadOnlyDictionary<string, decimal> D24_D29 { get; set; } = ...; // semántica T0
}

public sealed class ValidacionCruzadaSnapshot  // un ASE en un período
{
    public Ase Ase { get; set; } = new();
    public IReadOnlyList<ValidacionEmpresaSnapshot> PorEmpresa { get; set; } = ...;
    public DetValiRetriSnapshot? DetValiRetri { get; set; }
    public decimal ValidacionTotal { get; set; }
    public IReadOnlyDictionary<string, decimal> Controles { get; set; } = ...; // Valida -*
}
```

`IValidador`: overload aditivo `Validar(ResultadoRemuneracion, IReadOnlyList<WorkbookLeafInputs>, IReadOnlyList<ValidacionCruzadaSnapshot>)` (contrato intacto, implementación extendida; snapshot ausente/vacío = HU-12 puro). `ProcesadorPeriodo`: tras la escritura, lee el snapshot de la salida (read-only) y evalúa los gates con fail-fast que nombra ASE **+ validación**; una escritura, como siempre (G-principio single-write). Modo sin snapshot (Q1/Q2 regresión): pasos 2.7 se omiten.

### 2.5 `IValidador` 2.7 (sin reabrir HU-04..HU-12)

1. Todo lo HU-07..HU-12 intacto (matcheo estricto por `Ase.Id`, `GranTotal = Σ`, gates leaf/empresa/banco/balance/ajustes/DetRetri).
2. Nuevo, **solo si hay snapshot para el ASE** (matcheo `Single` por Id; sin snapshot → sin gates nuevos): por cada empresa, `O3 = 0 ±0.5` y `P3 = 0` exacto; `D16:D21 = 0 ±0.5` por ASE; `D24:D29` según semántica T0; `VALIDACION_TOTAL = 0 ±0.5` o composición T0; controles `Valida -*` según T0.
3. J9:J14/D9:D14 vs ROUND: **excluidos** (D6; test dedicado que documenta la exclusión).
4. El validador NO abre `.xlsx` (igual que HU-06..HU-12). La semántica D85:D89 y DetRetri-D van a Capa A, no se re-abren.
5. Coherencia `WorkbookLeafCoherence`: `ValidarCruzadasContraResultado` con matcheo estricto (misma regla HU-07, extendida).

### 2.6 UI — Visual Design Intent (delta mínimo)

Densidad Balanced, mismos GroupBoxes, sin restyle/colores/iconos. El `txtLog` agrega, por cada ASE, un bloque VALIDACIONES (por empresa: O3/P3; DetValiRetri: D16:D21/D24:D29; TOTAL; controles — "cierra/diverge" con valor). Serilog: mismos eventos HU-07..HU-12 con propiedad `Validacion` (nombre de la validación). Sin nuevos controles.

### 2.7 Golden Capa A de validaciones en ambos períodos (honestidad HU-06..HU-12)

| # | Qué | Contra qué | Tol |
|---|---|---|---|
| A1 | Snapshot-oráculo leído de la **salida** (read-only, OpenXML sin recalcular) | Mismas celdas del golden canónico correspondiente (Q1 / Q2) | ±0.5 |
| A2 | Gates de **dominio** por ASE × validación (O3/P3, D16:D21/D24:D29, TOTAL, controles) | Caché golden de las hojas de validación (ambos períodos) | ±0.5 (P3 exacto 0) |
| A3 | Todas las hojas de validación siguen siendo fórmula en la salida + intento de escritura en ellas falla a propósito | Estructura | n/a |
| A4 | SHA256 plantillas (Q1 y Q2-canónico) igual antes/después | — | n/a |
| A5 | **Prohibido** comparar caché de fórmula de la salida vs golden; **prohibido** usar agregados HU-02 o totales Q1 como oráculo de validación | — | prohibido |
| A6 | `TotalAse`/`GranTotal` de dominio vs caché golden D104:D109 en ambos períodos (re-assert HU-12, sin duplicar suite) | Caché golden ambos canónicos | ±0.5 |
| A7 | Regresión: 128/128 + harness 24/24 verdes con snapshot ausente (= HU-12 puro) | Suites existentes | n/a |
| A8 | Fe de erratas: todo mapa/test nuevo referencia hojas DetRetri/DetValiRetri como **36/37** (no 93/94) | Estructura canónica | n/a |

Capa B (manual Excel: abrir, recalcular, comparar `VALIDACION_*`/`DetValiRetri`/`VALIDACION_TOTAL`/`Valida -*` vs golden) fuera de CI, protocolo §5.3. **Capa B manual queda como acción del usuario.**

### 2.8 File changes previstos

| Archivo | Acción | Motivo |
|---|---|---|
| `Remuneracion.Core/Models/ValidacionCruzadaSnapshot.cs` (+ `ValidacionEmpresaSnapshot`, `DetValiRetriSnapshot`) | Crear | Snapshot-oráculo por ASE (D1; dominio puro) |
| `Remuneracion.Core/Interfaces/IValidacionOracleReader.cs` | Crear | Contrato read-only del oráculo (D1) |
| `Remuneracion.Core/Interfaces/IValidador.cs` | Modificar | Overload aditivo con snapshots (contrato intacto) |
| `Remuneracion.Core/Services/ValidadorBasico.cs` | Modificar | Gates §2.5 (por empresa/D16:D21/D24:D29/TOTAL/controles; J9:J14 excluidos D6) |
| `Remuneracion.Core/Services/ProcesadorPeriodo.cs` | Modificar | Leer snapshot post-escritura (read-only) + evaluar gates; fail-fast ASE+validación |
| `Remuneracion.Infrastructure/Excel/WorkbookLeafCellMapValidaciones.cs` | Crear | Mapa-oráculo hermano por (`Ase.Id`, período, validación) + protegidas (congelado T0; hojas 36/37, no 93/94) |
| `Remuneracion.Infrastructure/Excel/ValidacionOracleReader.cs` | Crear | OpenXML read-only: fórmulas como oráculo + caché `<v>` (nunca escribe; asserts W2) |
| `Remuneracion.Infrastructure/Excel/WorkbookLeafCoherence.cs` | Modificar | `ValidarCruzadasContraResultado`, matcheo estricto |
| `Remuneracion.Infrastructure/Excel/OpenXmlPlantillaWriter.cs` | Modificar | Mapa protegido extendido (todas las validaciones; rechaza escritura en ellas) |
| `Remuneracion.WinForms/Form1.cs` | Modificar | Bloque VALIDACIONES por ASE en log (§2.6) |
| `Remuneracion.IntegrationTests/GoldenValidacionesCruzadasTests.cs` | Crear | Capa A ambos períodos (§2.7) |
| `Remuneracion.IntegrationTests/ValidacionesCruzadasTests.cs` | Crear | Dominio in-memory: gates por validación, mismatch nombra ASE+validación, exclusión J9:J14 (D6) |
| `Remuneracion.IntegrationTests/ValidacionOracleReaderTests.cs` | Crear | Reader read-only: snapshot vs golden ±0.5, asserts W2, divergencia documentada |
| `Remuneracion.IntegrationTests/ProcesadorPeriodoTests.cs` | Modificar | Casos 2.7 (con snapshot → veredicto; sin snapshot → HU-12 puro) |
| `Remuneracion.IntegrationTests/DetRetriCapacidadTests.cs` | Crear | **W1-guarda**: plantilla futura con D9:D14 de otra capacidad → fail-fast nombra hoja (nunca truncado) |
| `Remuneracion.IntegrationTests/Insumos.cs` | Modificar | Helpers `SnapshotValidacion(aseId, quincena)`, goldens Q1 + Q2-canónico |

**No tocar (salvo bug blocker):** `IPlantillaWriter`/validation-only (HU-04); agregados HU-02; coherencia `F25`-Extemp HU-05; mapas HU-08/HU-09/HU-10/HU-11/HU-12; `DetRetriRounder` (se reutiliza); reglas Q1/Q2 de `ValidadorBasico`; cadenas 2.5/2.6 certificadas; recorte 1 HU-11 (HU-08 fuera de Q2); `requirements/` legado.

### 2.9 Absorción de deuda HU-12 (W1 + W2)

| Deuda | Contenido | Dónde se salda en este plan |
|---|---|---|
| W1-negativas | Casos negativos §4.5 del Plan 12 (slot ausente ASE5+R1, `Replace` que no matchea, origen-equivocado como DetRetri, salida==plantilla, sin filas DetRetri, Q1 con rama Q2, HU-08 en Q2) | Portados a §5.2 como regresión obligatoria (no se re-descubren; se re-ejecutan verdes) |
| W1-guarda | Guarda de capacidad `DetRetri` D9:D14 ante plantilla futura (más/menos filas ASE) | `DetRetriCapacidadTests` + guarda en el writer/reader (§2.8, §4 Fase 4): capacidad distinta → fail-fast que nombra la hoja |
| W2 | Asserts de existencia explícita en T0 (hoja existe, celda existe, `<f>` presente donde aplique) | T0-0.7 + asserts en `ValidacionOracleReader` (D7): sin asserts no hay snapshot; lectura que no encuentra → fallo nombrado, nunca 0 silencioso |
| Fe de erratas | Plan 12 "sheets 93/94" → verdad 36/37 | §9.11 + A8: todo artefacto nuevo usa 36/37 |

---

## 3. SPEC / Acceptance Criteria (trazabilidad Propuesta §10 + base docs)

### Requirement 1 — Oráculo de lectura por empresa con P3 exacto (CA-1/CA-2; Plan 08 V4)

El sistema **MUST** comparar Recaudo * vs `REMUNERACION_*` por empresa mediante el snapshot-oráculo (O3 = 0 ±0.5, P3 = 0 exacto) en ambos períodos, con matcheo estricto por `Ase.Id`. **MUST NOT** escribir jamás una hoja `VALIDACION_*`; **MUST NOT** asumir refs H/N sin T0.

- GIVEN snapshot con O3=0.3/P3=0 en ENEL-ASE1 → WHEN `Validar(resultado, leafs, snapshots)` → THEN sin errores de esa empresa.
- GIVEN O3=2.0 en OCCIDENTE-ASE3 → THEN error que nombra ASE3+OCCIDENTE, sin salida certificada como "válida".
- GIVEN `VALIDACION_*` sin snapshot (regresión) → THEN comportamiento HU-12 intacto.

### Requirement 2 — DetValiRetri D16:D21/D24:D29 vs dominio (CA-2/CA-3; `WorkbookLeafCellMapQ2.cs:257-268`)

El sistema **MUST** verificar D16:D21 = 0 ±0.5 por ASE y D24:D29 según semántica T0 en ambos períodos (espejo Q1 dictado por T0, no copiado de Q2). **MUST NOT** exigir ROUND en D9:D14 (D6).

- GIVEN D16:D21 = 0 ±0.5 en los 5 ASE Q2 → THEN gate verde.
- GIVEN D18 = 1.2 en ASE3 → THEN error que nombra ASE3+D18.
- GIVEN D9:D14 ≠ ROUND (hallazgo) → THEN el gate lo excluye a propósito (test dedicado).

### Requirement 3 — VALIDACION_TOTAL y Valida -* (CA-2/CA-3)

El sistema **MUST** evaluar `VALIDACION_TOTAL` y `Valida -Remunera/-Anticipos/-Control Recaudo` según semántica T0 por período. Si alguna no cierra ±0.5 contra el golden → NEEDS_CONTEXT con recorte (D5b), **MUST NOT** inventar su semántica.

- GIVEN TOTAL = 0 ±0.5 ambos períodos → THEN verde.
- GIVEN control `-Anticipos` con divergencia no explicada → THEN recorte declarado + NEEDS_CONTEXT, resto de gates intactos.

### Requirement 4 — Hojas de validación siempre protegidas (CA-4)

`VALIDACION_*`, `DetValiRetri`, `VALIDACION_TOTAL`, `Valida -*`, `GERENTES_*` **MUST** estar en el mapa protegido en ambos períodos (fallan la escritura si alguna deja de ser fórmula); intentar escribir en ellas **MUST** fallar a propósito en test.

### Requirement 5 — Golden Capa A ambos períodos + regresión (CA-3/CA-5)

Las pruebas **MUST** ejecutar la matriz §2.7 en Q1 y Q2-canónico. **MUST NOT** comparar caché de salida vs golden (A5). Los 128 tests + harness 24/24 **MUST** seguir verdes con snapshot ausente.

### Requirement 6 — Deuda W1+W2 saldada (CA-4/CA-5)

Las negativas §4.5 Plan 12 **MUST** re-ejecutarse verdes; la guarda de capacidad D9:D14 **MUST** fallar nombrando la hoja ante plantilla futura; cada celda-oráculo **MUST** assertar existencia (W2) antes de comparar.

| CA §10 | HU-13 |
|---|---|
| CA-1 | Lee el oráculo de validación por ASE y período (fail-fast nombra ASE+validación; asserts W2) |
| CA-2 | Gates cruzados = 0 ±0.5 por validación (P3 exacto; J9:J14 excluidos D6) |
| CA-3 | Capa A ambos períodos; validaciones correctas post-Excel (Capa B residual) |
| CA-4 | Reassert mapa protegido extendido + plantillas no mutadas (hash) + W1-guarda + fe de erratas 36/37 |
| CA-5 | Gates 2.7 aditivos + coherencia cruzada con `Single` por Id |
| CA-6 | Serilog + log VALIDACIONES por ASE |
| CA-7 | Mismo flujo 5-ASE + bloque validaciones (sin snapshot = HU-12 puro) |

---

## 4. TASKS

## Review Workload Forecast

| Field | Value |
|---|---|
| Estimated changed lines | 700–1100 |
| 400-line budget risk | High |
| Chained PRs recommended | Yes |
| Suggested split | PR1 T0 + oráculo congelado → PR2 dominio snapshots + gates → PR3 oracle-reader read-only → PR4 coherencia + writer protegido + período/UI → PR5 golden ambos períodos + W1 + W2 + regresión |
| Delivery strategy | ask-on-risk |
| Chain strategy | pending |

Decision needed before apply: Yes. Chained PRs recommended: Yes. Chain strategy: pending.
400-line budget risk: High.

### Suggested Work Units

| Unit | Goal | Likely PR | Notes |
|---|---|---|---|
| 0 | Discovery T0 + mapa-oráculo congelado (ambos canónicos) | PR 1 | Bloquea todo; solo lectura + datos |
| 1 | Dominio snapshots + gates `ValidadorBasico` | PR 2 | Depende de PR 1 |
| 2 | Oracle-reader read-only + asserts W2 | PR 3 | Depende de PR 1 |
| 3 | Coherencia + writer protegido + período + UI | PR 4 | Depende de PR 1–3 |
| 4 | Golden ambos períodos + W1 + regresión 128/128 + 24/24 | PR 5 | Depende de PR 2–4 |

### Phase 0 — Discovery T0 (BLOQUEANTE, solo lectura, ambos canónicos)

- [ ] 0.1 Fijar canónicos de trabajo (Q1 golden + Q2-canónico HU-11/HU-12 reutilizados); registrar SHA256 de trabajo. No se re-fijan.
- [ ] 0.2 `VALIDACION_RECIP/ENEL/ENERBIT/OCCIDENTE/EAAB-CL`: dumpear O3/P3 por empresa (fórmula + `<v>` + refs H/N) en ambos canónicos; congelar semántica (¿O3=H3-N3 y P3=INT(O3) en Q1 y Q2? ¿filas por ASE?).
- [ ] 0.3 `DetValiRetri` Q1: espejo de `WorkbookLeafCellMapQ2.cs:257-268` (D16:D21 vs ¿U104:U109 en Q1?; D24:D29 refs M/L/J/AJUSTES/U o variantes Q1); clasificar valor-vs-fórmula celda-por-celda.
- [ ] 0.4 `VALIDACION_TOTAL`: dumpear layout + fórmulas + `<v>` en ambos canónicos; congelar composición (¿Σ empresas? ¿Σ ASE?).
- [ ] 0.5 `Valida -Remunera/-Anticipos/-Control Recaudo`: dumpear layout por hoja en ambos canónicos (filas por ASE, columnas comparadas, booleanas vs diferencias); congelar semántica.
- [ ] 0.6 `GERENTES_*` ×5: verificar fórmulas puras en ambos canónicos (cierra el pendiente Plan 08 T0-0.5); cuantificar J9:J14/D9:D14 vs ROUND por ASE y período (veredicto D6, no corrección).
- [ ] 0.7 Congelar `WorkbookLeafCellMapValidaciones` (oráculo + protegidas por `Ase.Id`/período/validación) con **asserts de existencia explícita W2** por entrada + corrección 36/37 (no 93/94) + barrido de externalLinks en celdas-oráculo. **Nada entra al código sin esta tabla.**
- [ ] 0.8 Confirmar baseline HU-12 (128/128 + 24/24) intacto antes del apply.

### Phase 1 — Dominio (snapshots + gates)

- [ ] 1.1 `ValidacionCruzadaSnapshot.cs` (+ empresa/DetValiRetri) e `IValidacionOracleReader.cs` (§2.4; dominio puro, sin deps).
- [ ] 1.2 Overload `IValidador.Validar(resultado, leafs, snapshots)` + gates §2.5 en `ValidadorBasico` (O3 ±0.5, P3 exacto, D16:D21, D24:D29-T0, TOTAL, controles; J9:J14 excluidos; `Single` por Id; sin snapshot = HU-12 puro).
- [ ] 1.3 Tests in-memory `ValidacionesCruzadasTests`: gate verde, mismatch nombra ASE+validación, P3≠0 falla exacto, exclusión J9:J14 dedicada, sin snapshot = HU-12 intacto.

### Phase 2 — Lectura-oráculo (read-only + W2)

- [ ] 2.1 `WorkbookLeafCellMapValidaciones.cs` (datos T0-0.7; hojas 36/37) + `ValidacionOracleReader.cs` (OpenXML read-only, caché `<v>` + verificación `<f>`; asserts W2 por celda: hoja/celda/fórmula existen o fallo nombrado).
- [ ] 2.2 `ValidacionOracleReaderTests`: snapshot vs golden ±0.5 ambos períodos; hoja/celda ausente → fallo nombrado (W2); divergencia J9:J14 documentada; el reader nunca escribe (hash salida intacto tras lectura).

### Phase 3 — Orquestación + UI delta mínimo

- [ ] 3.1 `WorkbookLeafCoherence.ValidarCruzadasContraResultado` (matcheo estricto) + `OpenXmlPlantillaWriter` con mapa protegido extendido (toda validación; escritura en ellas falla a propósito por test).
- [ ] 3.2 `ProcesadorPeriodo`: leer snapshot post-escritura (read-only) + evaluar gates; fail-fast ASE+validación; sin snapshot = HU-12 puro.
- [ ] 3.3 `Form1`: bloque VALIDACIONES por ASE en log (§2.6) + Serilog `Validacion`.

### Phase 4 — Pruebas, golden y deuda W1

- [ ] 4.1 `GoldenValidacionesCruzadasTests`: matriz §2.7 en ambos períodos (A1–A8; A5 prohibido a propósito por test).
- [ ] 4.2 **W1-negativas** (§4.5 Plan 12) portadas y verdes como regresión obligatoria.
- [ ] 4.3 **W1-guarda** `DetRetriCapacidadTests`: plantilla futura con D9:D14 de otra capacidad → fail-fast nombra la hoja (nunca truncado silencioso).
- [ ] 4.4 `ProcesadorPeriodoTests`: con snapshot → veredicto por validación; sin snapshot → HU-12 puro; divergencia no-T0 → NEEDS_CONTEXT (el test declara el recorte, no lo inventa).
- [ ] 4.5 Los 128 tests + harness 24/24 existentes verdes; build 0 warnings; CRLF; sin commit.

### Phase 5 — Documental

- [ ] 5.1 Capa B manual §5.3 ejecutada una vez por período y evidenciada (acción del usuario; sin fingirla como gate de merge).
- [ ] 5.2 Cierre deja explícito que Fase 2 queda cerrada (2.1–2.7) y el frente siguiente es Fase 3.

---

## 5. Test Strategy

| Capa | Qué | Cómo |
|---|---|---|
| Unidad | Gates 2.7 (O3/P3, D16:D21/D24:D29, TOTAL, controles; `Single` por Id) | In-memory con snapshots T0 |
| Unidad | Exclusión J9:J14/D9:D14 vs ROUND (D6) | Test dedicado que demuestra la divergencia y la excluye |
| Unidad | Oracle-reader (read-only, asserts W2, nunca escribe) | Goldens reales, sin salida |
| Integración | Período con snapshot (salida temp, fail-fast ASE+validación) | Insumos Q1 + Q2, canónicos |
| Golden Capa A | Matriz §2.7 ambos períodos | OpenXML read-only + aritmética dominio (A5 aplica) |
| Regresión | 128/128 + harness 24/24 sin snapshot | Suites existentes, sin cambios |
| Deuda W1 | Negativas Plan 12 §4.5 + guarda capacidad D9:D14 | Tests dedicados obligatorios |
| UI | Bloque VALIDACIONES | Funcional manual (sin harness) |
| Capa B | Validaciones post-Excel ambos períodos | Manual — §5.3 |

### 5.1 Fixtures

- Goldens/plantillas: Q1 (`Remuneracion 202607-1 Total.xlsx`) + Q2-canónico HU-11/HU-12 (sin re-fijar).
- Fuentes: `Docs/Insumos/REMUNERACION 2026071/{1..5}-*/` + `Docs/Insumos/REMUNERACION 2026072/{1..5}-*/`.
- Referencia: semántica T0-0.7 + fragmentos `WorkbookLeafCellMapQ2.cs:257-268` (Q2) + espejo Q1 (T0), tolerancia ±0.5 (P3 exacto 0).

### 5.2 Casos negativos obligatorios (nombran ASE y validación)

O3≠0 en empresa-ASE (falla ASE+empresa); P3≠0 exacto (falla aunque O3 esté en tolerancia); D18≠0 (falla ASE3+D18); hoja/celda oráculo ausente (W2: falla nombrando hoja+celda, nunca 0 silencioso); escritura en hoja de validación (falla a propósito); J9:J14≠ROUND (excluido a propósito, no falla); plantilla futura con D9:D14 de otra capacidad (W1-guarda: falla nombrando la hoja); snapshot ausente (HU-12 puro, no falla); salida == plantilla (no in-place); divergencia no-T0 (NEEDS_CONTEXT con recorte, nunca invención); HU-08 en Q2 (fuera de alcance — no se toca); W1-negativas Plan 12 §4.5 (todas verdes).

### 5.3 Protocolo manual Capa B (no CI — acción del usuario)

1. Generar salidas Q1 y Q2 a rutas distintas de los goldens. 2. Abrir en Excel, recalcular. 3. Comparar `VALIDACION_*` (O3/P3), `DetValiRetri` (D16:D21/D24:D29), `VALIDACION_TOTAL`, `Valida -*` vs goldens ±0.5. 4. Evidencia adjunta. No es gate de merge.

---

## 6. Architecture Validation Certificate

### 6.1 SOLID

| Principio | Veredicto |
|---|---|
| **S** | Validaciones = extensión de lectura/comparación (mapa-oráculo + snapshots + gates); cálculo/período/writer conservan su rol HU-01..HU-12. |
| **O** | Se agregan modelos/mapa/overload; sin snapshot el path HU-12 funciona sin modificarse (abierto sin modificar). |
| **L** | `ValidadorBasico` suma overload sin cambiar el comportamiento de los 3 existentes ni sus fail-fasts. |
| **I** | `IValidacionOracleReader` separado de `IWorkbookLeafInputReader`/`IRecaudoReader`; reader/validador crecen por overload. |
| **D** | Core define snapshots/gates; Infrastructure/WinForms componen. Sin nuevas deps; `DetRetriRounder` reutilizado. |

### 6.2 Best Practices

- La verdad del workbook manda: las validaciones son fórmulas ⇒ cero escritura directa (Rector §6 preservación; doctrina BCE-H HU-10, D2 Plan 11).
- La semántica exacta se prueba contra ambos canónicos, no se asume (honestidad §0.1; lección Plan 12 v1).
- Mapa explícito por `Ase.Id` + período + validación, no offsets ni filas/columnas fijas (Rector §11.3).
- Oráculo read-only con asserts de existencia (W2): sin asserts no hay snapshot.
- Una escritura atómica; plantillas nunca mutadas; hash A4 extendido.
- Golden honesto ambos períodos (OpenXML no recalcula; A5/A6/A7/A8).
- Fail-fast nombra ASE+validación; sin salida "válida" ante fallo; sin insert/delete de filas.
- Deuda saldada en la última HU de Fase 2 (W1+W2 aquí, no en Fase 3).
- Fe de erratas 36/37 aplicada en todo artefacto nuevo.

### 6.3 Performance

- 5 ASE × lecturas-oráculo read-only (decenas de celdas por validación) + gates in-memory, fuera de la sesión de escritura. Irrelevante a esta escala; `Task.Run` existente para no congelar el form.

**Veredicto:** APROBADO como it. 2.7 (última de Fase 2) **si** T0 congela el mapa-oráculo con evidencia en ambos canónicos (§4 Fase 0) y se acepta CA-3 parcial (Capa A en CI, Capa B manual como acción del usuario).

---

## 7. Riesgos

| Riesgo | Likelihood | Mitigación |
|---|---|---|
| Alguna validación no cierra ±0.5 contra ninguna hipótesis (semántica no entendida) | Media | D5b: NEEDS_CONTEXT con recorte declarado; la validación queda en divergencia conocida sin bloquear el resto; prohibido inventar semántica |
| `VALIDACION_TOTAL` / `Valida -*` traen layouts distintos Q1-vs-Q2 | Media | T0-0.4/0.5 con dispatch por período (D3); si el hueco es real, recorte a las validaciones verificadas |
| `GERENTES_*` no son fórmulas puras en algún período | Baja | T0-0.6; si traen valores, entran como protegidas-valor (no gate numérico) sin rebase |
| Divergencia J9:J14/D9:D14 contamina un gate por error | Alta | D6 + test de exclusión dedicada; T0-0.6 la cuantifica sin "arreglarla" |
| Hoja oráculo faltante en plantilla futura da "0" silencioso | Media | W2 (D7): asserts de existencia por celda; ausente = fallo nombrado |
| Plantilla futura cambia la capacidad D9:D14 | Media | W1-guarda (D8): fail-fast nombrando la hoja; nunca truncado silencioso |
| Inflar a Fase 3 dentro de esta HU (CLI/manual tientan) | Media | §0.2 out-of-scope explícito; rechazar PRs que lo metan |
| Reabrir 2.6/2.5 o HU-08-en-Q2 por inercia | Media | Fuera de alcance explícito; rechazar PRs que lo metan |
| Comparar caché de salida vs golden y "cerrar" en falso | Alta | A5 + canónicos fijos; este plan lo prohíbe |
| Q1/Q2 regresión rota por el overload nuevo | Media | 128/128 + 24/24 como red en PR5; sin snapshot = HU-12 puro por construcción |
| Ceros legítimos confundidos con "falta de lectura" en el oráculo | Alta | W2 distingue "leído 0" de "celda ausente" (fallo); tests de ceros explícitos |

---

## 8. Rollback

- Revertir PRs en orden inverso (golden+W1 → período/UI/coherencia → oracle-reader → dominio → mapa/T0).
- HU-01..HU-12 intactas sin esta HU: sin snapshot = comportamiento HU-12 puro; Q1/Q2 siguen certificadas.
- `Docs/Insumos/` untracked: jamás como destino; pisada se restaura desde backup.
- Si T0 demuestra semántica que no cierra (Riesgo 1), recorte con rebase a las validaciones verificadas, no invención.

---

## 9. Decisiones fijadas por este plan

1. Rector = Propuesta §9 it. 2.7 / §6 / §10 (±0.5) + Detalle de plantilla y Prompt Maestro como origen funcional (§1.5). Escritura de validaciones, INTERVENTORIA, ANT EXT-REV, Fase 3 quedan fuera. Proceso de Recaudo NO aplica salvo cita directa.
2. Lógica de validaciones cruzadas en dominio como ORÁCULO DE LECTURA (Recaudo * vs `REMUNERACION_*` = 0 patrón O3/P3; DetValiRetri D16:D21/D24:D29 vs dominio ±0.5; `VALIDACION_TOTAL`; `Valida -*`): se validan, no se escriben; hojas siguen protegidas.
3. T0-evidencia primero y T0 bloquea el mapa: semántica exacta por validación dictada por T0 contra ambos canónicos, nunca asumida; si no cierra ±0.5 → NEEDS_CONTEXT con recorte.
4. PRs encadenados con T0 al frente (Unidad 0–4, §4); golden honesto Capa A ambos períodos; Capa B manual residual del usuario; matcheo estricto por Id; Q1/Q2 intactos por construcción.
5. Deuda HU-12 absorbida: W1 (negativas §4.5 Plan 12 + guarda de capacidad `DetRetri` D9:D14) + W2 (asserts de existencia explícita en T0).
6. J9:J14/D9:D14 ≠ ROUND (hallazgo HU-12) excluido de los gates y documentado (D6); prohibido "cerrarlo".
7. Una sola escritura atómica; sin insert/delete de filas (NEEDS_CONTEXT honesto si el oráculo no alcanza).
8. UI delta mínimo + Serilog por validación; OPA = Ejecutar.
9. Regresión 128/128 + harness 24/24 intacta (sin snapshot = HU-12 puro).
10. Fase 2 cierra con esta HU (2.1–2.7); el frente siguiente es Fase 3.
11. Fe de erratas del Plan 12: "sheets 93/94" → verdad **36/37**; todo artefacto nuevo de este plan usa 36/37.

**Listo para aprobación del Ingeniero. No implementar hasta OK.**
