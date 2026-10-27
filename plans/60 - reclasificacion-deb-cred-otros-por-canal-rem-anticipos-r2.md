---
titulo: Plan 60 - Reclasificacion Deb/Cred a Deb/Cred-otros por canal en Rem. Anticipos R2
estado: PROPUESTO — pendiente Clarification Gate §6.5 y APROBADO del Ingeniero; sin aprobación no se abre H0
fecha: 2026-10-27
numero: 60
planes_ocupados: 01..59 ocupados; este es el 60
padre: plans/59 - eliminacion-coordenadas-quemadas-geometria-derivada-v3.md (§12: partición Deb/Cred vs Deb/Cred-otros declarada como pendiente de negocio sin regla derivable)
madre_doctrina: .opencode/project-context.md (plantilla autoridad; values-only; guard anti-fórmula; tolerancia ±0,5; prohibido <v> stale como oráculo; 0 warnings; sin valores quemados por período/ASE/canal; nunca sobrescribir fórmulas) + plans/55 (conciliación por empresa) + plans/56 (identidad A/B, cero documentado) + plans/47 (canal efectivo R46-1) + plans/40 (W-1 evalúa+sella) + plans/31 (gate workbook-vs-dominio)
caso_rojo: corrida real 2026101 --cinco-ase — divergencia declarada en `Rem. Anticipos R2` ASE2: la app escribe el canal O en Q (54,00 / 6.759.294,81 / 6.759.348,81) con R=0 en toda la hoja; el oráculo administrativo mueve ese bloque a R y ajusta OCCIDENTE/Componente/Total. Una sola causa, no cinco: falta la regla de reclasificación por canal.
gate_entrada:
  plantilla_v4: Docs/Plantilla estandarizada/v4/Formato Remuneracion 2026xx-x Total v4.xlsx — VIGENTE. Huella semántica (helper canónico `HuellaSemanticaPlantilla`) a7476ee2e80dfa13c9e39430d1ea4d77f0c52e3ca0a3d11a526d5140b8608c55; SHA256 de archivo 5FEDFD316980276945DF2AF29784EFFC5FE178C5CF1AA16F3C5D5F6F0A4E38E5 (705.805 bytes, dato INFORMATIVO). `Rem. Anticipos R2` hoja11: fórmulas en Q57/R57, Q63/R63, Q69/R69, Q178/R178, Q181/R181, Q184/R184, Q190/R190, Q299/R299, Q305/R305, Q311/R311, Q423/R423, Q426/R426, Q429/R429, Q435/R435, Q510/R510, Q516/R516, Q519/R519, Q522/R522 (cuadro OPORTUNO — ver §0.2)
  suite_base: la vigente al abrir H0 (trx de entrada citado en H0; 0 Skip; set-diff: rojo nuevo fuera de lista = STOP)
  sin_commits: NUNCA git commit / git push; NUNCA git checkout / git reset
  excel_cerrado: Excel debe estar cerrado durante las corridas
---

# Plan 60 — Reclasificación Deb/Cred → Deb/Cred-otros por canal en `Rem. Anticipos R2`

> SDD completo PROPOSE → DESIGN → SPEC → TASKS. Fuente única de verdad para la implementación una vez aprobado ESTE documento. No se implementa sin aprobación del Ingeniero (punto de bloqueo: sin su OK no se abre H0; estado inicial PROPUESTO).
> Lectura obligatoria previa: `plans/59` (§12: la partición Q/R como pendiente de negocio; §14: rebase v4; §15: Q2 por empresa) + `plans/47` (canal efectivo) + `plans/56` (identidad A/B) + `.opencode/project-context.md` (doctrina madre).
> Entorno: Windows, repo con espacios en la ruta; plantillas CERRADAS en Excel al correr tests (FileShare). Stack .NET 10 (Core/Infrastructure/Cli/WinForms) + Excel; persistencia por archivos Excel, NO base de datos. Español; CRLF sin mezclar.
> Visual Design Intent: N/A (cero UI en este plan).

---

## §0. Evidencia T0 (medida por el orquestador; el plan la cita, el H0 la re-verifica)

### §0.1 Regla del Ingeniero (texto literal, 6 puntos — fuente normativa, ver §10)

1. Cuando el componente **Deb/Cred** contiene algún concepto `Total` de algún **canal de recaudo** con valor **mayor a 500.000**, se debe mover ese valor de `Deb/Cred` a `Deb/Cred-otros`.
2. Se trasladan los valores del canal: los conceptos que hacen parte de ese canal (en el ejemplo: `Vlr Servicio`, `Vlr Intereses`, `Total`). No es un ajuste puntual de una celda: es el bloque completo del canal.
3. El total de la fuente asociado al canal trasladado (en el ejemplo "OCCIDENTE") debe quedar consistente: aplicar en `Deb/Cred-otros` el total respectivo y modificar el total de `Deb/Cred` de manera consistente.
4. "Lo que es fórmula no se debe tocar" (las filas resumen que son fórmula se recalculan solas).
5. Debe servir de evidencia/auditoría del resultado administrativo.
6. **No debe quedar quemado** para Occidente ni para el canal "O": es una regla general de la pestaña **`Rem. Anticipos R2`**.

### §0.2 Dónde vive la regla (geometría de la hoja)

| Dato | Valor medido |
|---|---|
| Hoja | `Rem. Anticipos R2` |
| Columna `Total` | E |
| Columnas de componentes | F..P |
| Columna `Deb/Cred` | **Q** |
| Columna `Deb/Cred-otros` | **R** |
| Fila de encabezados (plantilla v4) | **2** |
| Bloque ASE2 (LIME) en la v4 | cabecera de bloque `75`, encabezados de columna `76`, detalle `77..147` |
| Agregados del bloque ASE2 en la v4 | `OCCIDENTE` **116**, `Componente` **117**, `Subs/Cont` **146**, `Total` **147** |
| Canales declarados por la v4 en el bloque (12) | A · E · 1 · 2 · 3 · 4 · 5 · 6 · 9 · H · O · T (la fuente de octubre trae 9: **no asumir lista de canales**) |
| Filas del cuadro OPORTUNO que son **fórmula** en la v4 (hoja11) | `Q57/R57`, `Q63/R63`, `Q69/R69`, `Q178/R178`, `Q181/R181`, `Q184/R184`, `Q190/R190`, `Q299/R299`, `Q305/R305`, `Q311/R311`, `Q423/R423`, `Q426/R426`, `Q429/R429`, `Q435/R435`, `Q510/R510`, `Q516/R516`, `Q519/R519`, `Q522/R522` → **no se escriben, se recalculan** |
| Filas de agregado DENTRO del bloque de detalle (`OCCIDENTE`, `Componente`, `Subs/Cont`, `Total`) | **literales 0** en la plantilla v4 → **sí las escribe la app** y por eso deben ajustarse |

### §0.3 Fuente ASE2 (Lime) — el único disparo de octubre-1

Ruta: `Docs/Prueba Octubre-1/Insumos/2-Lime/RerpoteDetalleSaldosaFavor_..._2026101617441360.xlsx` (nombre completo exacto: **a verificar en implementación** contra el disco — el plan no congela el timestamp), hoja `Sheet1`, headers en fila **4**. Bloque completo = filas **5..57** (la última es `A=Total`).

| Fila (fuente) | Firma A\|B\|C\|D | Q (`Deb/Cred`) medido |
|---|---|---|
| 26 | `\| \|O\|Vlr Servicio` (detalle del canal O) | **54,00** |
| 27 | `\| \|O\|Vlr Intereses` (detalle del canal O) | **6.759.294,81** |
| 28 | `\| \|O\|Total` — **cierre de canal (leaf) O** | **6.759.348,81** |
| 32 | `\|OCCIDENTE\|Total\|` (empresa) | **6.759.342,66** (= H + O + T, medido) |
| 33 | `Componente\|Total\|\|` (sección) | **6.759.371,77** (= los 9 cierres de canal, medido) |
| 57 | `Total\|\|\|` (gran total) | **6.759.371,77** |
| resto de cierres de canal del ASE2 | A / E / 1 / 2 / 3 / 5 / H / T | 5,17 · 31,40 · 12,95 · 2,44 · 4,90 · -27,75 · -21,95 · 15,80 → **ninguno supera el umbral** |

Hechos estructurales medidos:

- La fuente **NO trae columna `Deb/Cred-otros`** (todos los R del modelo son ausentes → 0). El modelo del reader sigue siendo "observación fiel de la fuente": el 0 de R es observado, no inventado.
- ASE1, ASE4 y ASE5 **sí traen valores** en `Deb/Cred`, todos por debajo del umbral: máximo `|Q|` entre cierres de canal = **160,05** (ASE1) · **98,04** (ASE4) · **124,42** (ASE5); ASE3 = **2.232,90**. **En octubre-1 la regla dispara una sola vez (ASE2 · canal O).** (Corregido en H0: la versión inicial de este plan decía "sin valores" porque la columna `Deb/Cred` **no ocupa la misma posición en todos los ASE** — en ASE1/4/5 está en `P` y en ASE2/3 en `Q`; el reader la resuelve por ENCABEZADO, nunca por índice. Lección: prohibido escanear las fuentes por índice fijo de columna.)
- La etiqueta `O` aparece **dos veces** en la secuencia (cuadro `Componente` y cuadro `Subs/Cont`): un `Map<string,…>` por etiqueta colapsa ambos y rompe el cálculo (error ya reproducido en el prototipo — ver §3.3).

### §0.4 Oráculo administrativo (qué exige el resultado manual)

Ruta: `Docs/Prueba Octubre-1/Resultados/Resultado administrativo/Remuneracion 202610-1 Total.xlsx`, misma hoja `Rem. Anticipos R2`.

| Fila (por etiqueta, NO por número — ver §0.6) | Q (`Deb/Cred`) oráculo | R (`Deb/Cred-otros`) oráculo |
|---|---|---|
| canal O `Vlr Servicio` | **0** | **54,00** |
| canal O `Vlr Intereses` | **0** | **6.759.294,81** |
| canal O `Total (C=O)` | **0** | **6.759.348,81** |
| `B=OCCIDENTE C=Total` | **-6,15** | **6.759.348,81** |
| `A=Componente B=Total` | **22,96** | **6.759.348,81** |
| `A=Total` | **22,96** | **6.759.348,81** |
| `B=OCCIDENTE C=OPORTUNO D=Total` | **-6,15** | **6.759.348,81** — **FÓRMULA en v4 → no se escribe, se recalcula** |
| `B=TOTAL C=OPORTUNO D=Total` | **22,96** | **6.759.348,81** — **FÓRMULA en v4 → no se escribe, se recalcula** |

### §0.5 Salida actual de la app (la divergencia declarada)

Ruta: `Docs/Prueba Octubre-1/Resultados/5/Remuneración 202610-1 Total.xlsx`.

| Fila (por etiqueta) | Q actual | R actual |
|---|---|---|
| canal O (54,00 / 6.759.294,81 / 6.759.348,81) | **sin mover** (los 3 valores en Q) | **0 en toda la hoja** |
| `OCCIDENTE` | **6.759.342,66** | 0 |
| `Componente` | **6.759.371,77** | 0 |

Esta tabla ES la divergencia declarada que el plan cierra. El gate actual (`DetalleR2R4EscritoPlan59Tests.Q1_2026101_…`) la documenta comparando el **par** Q+R (coincide al centavo) con comentario explícito; este plan **cierra ese gate**: comparar ambas columnas **por separado** contra el oráculo.

### §0.6 Advertencia de geometría (normativa)

El oráculo administrativo está sobre una **geometría de plantilla anterior** (las filas del bloque no coinciden con la v4: p. ej. `OCCIDENTE` en **90** en el oráculo vs **116** en la v4). **Toda comparación debe ser por etiqueta/nivel, nunca por número de fila.** Un test que aserte `Q90` o `R116` es un quemado de coordenada y viola la doctrina madre.

### §0.7 Regresión histórica (ya verificado en los 4 oráculos administrativos)

En `2026071` / `2026072` / `2026082` ningún canal supera el umbral en ningún ASE (medido por el orquestador en los 4 oráculos: **R=0 en toda la hoja**). Estos tres períodos son la no-regresión del plan: tras el cambio deben salir **byte-idénticos en Q/R** (ver §6, gate H3.3).

---

## §1. Problema / Objetivo

### 1.1 Intent

Que la app aplique la regla administrativa de reclasificación `Deb/Cred → Deb/Cred-otros` por canal como **regla general, pura y sin quemados** (sin `if (ase==2)`, sin `if (canal=="O")`, sin `if (empresa=="OCCIDENTE")`, sin literales de fila/celda/período): cuando el cierre de un canal trae en `Deb/Cred` un importe significativo (> umbral), el **bloque completo del canal** se mueve a `Deb/Cred-otros` y los **agregados estructurales que son literales** se re-propagan de forma consistente, dejando intacta toda celda con fórmula. Al cerrar, el gate `DetalleR2R4EscritoPlan59Tests` compara **Q y R por separado** contra el oráculo y octubre-1 coincide al centavo; los períodos históricos no cambian.

### 1.2 In Scope (H0 + H1 + H2 + H3, prohibido ampliar sin nuevo OK del Ingeniero)

| ID | Entregable | Criterio de salida |
|---|---|---|
| **H0/T0** | Evidencia congelada por bytes: re-medición de §0 (fuente ASE2, oráculo, salida actual, fórmulas v4 del cuadro OPORTUNO, R=0 histórico) | Fixture verde + doc de evidencia; puerta H0→H1 |
| **H1** | Núcleo puro en Core: `ReclasificadorDebCredR2` + validación de pertenencia (§4.5) + traza de auditoría con coordenada + unitarios sintéticos | CA-R1/CA-R8 en suite unitaria, 0 warnings |
| **H2** | Punto único en el writer (`EscribirDetalleR2`, antes de `EmparejarBloque`) + actualización del XML doc de `ColumnasR2` + `ERR-RECLASIFICACION` si aplica | CA-R2 en corrida real octubre-1 |
| **H3** | Cierre del gate oráculo (Q y R por separado) + blast radius + fórmulas intactas + regresión histórica + review | CA-R3/CA-R4/CA-R5 + review APROBADO |
| **Candado** | Extensión del candado: falla si aparece en producción cualquier literal de canal/empresa/ASE/período/quincena o literal de fila/celda/columna Q/R asociado a esta regla | CA-R6 |
| **Gate** | Corridas CLI octubre-1 (`2026101`), julio-1 (`2026071`), julio-2 (`2026072`), agosto-2 (`2026082`) + suite completa con set-diff de rojos | CA-R7 |

### 1.3 Non-goals / out (vinculante — ver §9)

- No tocar la lógica de cálculo (`CalculoRemuneracion`), ni el CONSOLIDADO, ni `DetRetri`, ni `VALIDACION_TOTAL`.
- No tocar la plantilla (SHA/huella intactos; cero `<f>`, cero `InsertRow`/`DeleteRow`).
- No reabrir Plan 59, Plan 47 (canal efectivo), Plan 55/56, ni Plan 58-P2.
- Commits, UI, CLI nueva, docs de usuario. NUNCA `git commit` / `git push`. NUNCA `git checkout` / `git reset`.

### 1.4 Riesgos con mitigación (resumen; detalle en §7)

| ID | Riesgo | Mitigación |
|---|---|---|
| R-UMBRAL | Umbral mágico sin fuente documentada | Constante nombrada + comentario de procedencia + D1 con alternativa en una línea |
| R-SIN-TRAMO | Cierre cuya suma no reconcilia con ningún tramo de cierres | Fail-fast `ERR-RECLASIFICACION` solo para filas de cierre; filas de detalle sin match no se tocan |
| R-FORMULA-FUTURA | Una plantilla futura convierte Q/R del bloque de detalle en fórmula | El guard `ERR-PLANTILLA` ya aborta; declararlo, no rodearlo |
| R-ORACULO-FILA | Comparar contra el oráculo por fila (geometría distinta) | Prohibido; solo por etiqueta/nivel |
| R-DOBLE-APLICACION | Aplicar la regla dos veces (writer + processor) | Punto único en el writer + test de idempotencia |

---

## §2. Decisiones (D1..D6 — con recomendación y alternativa)

### D1 — Signo del umbral: `|Q| > 500.000` (recomendada) vs `Q > 500.000` (literal)

- **Recomendación del orquestador (a formalizar):** comparar **valor absoluto** (`Math.Abs(Q) > Umbral`), porque un canal con -6M en `Deb/Cred` es igual de significativo y quedaría mal clasificado con la lectura literal.
- **Alternativa literal:** `Q > 500.000` (solo positivos). Cambio en **una sola línea** si la administradora lo confirma (ver §4.4).
- **Constante única:** `UmbralReclasificacionDebCred = 500_000m` en un único punto con comentario de procedencia (regla de la administradora, 2026-10-27). Prohibido duplicar el literal.
- **Marca:** D1 queda como **decisión pendiente de confirmación explícita del Ingeniero** (opción recomendada vs literal). El Clarification Gate §6.5 la pregunta; los tests cubren ambas lecturas donde divergen (negativo grande — ver §6, fila U3).

### D2 — Propagación a agregados: obligatoria (evidencia del oráculo)

- El oráculo exige `OCCIDENTE/Componente/Total` ajustados en ambas columnas. Sin propagación el bloque cierra en el detalle pero el entregable sigue divergiendo en los agregados.
- **Contrato de "tramo de cierres":** para cada fila NO-leaf con `Q != 0`, se busca el tramo contiguo **más largo** de cierres de canal anteriores cuya suma sea igual a su `Q` (tolerancia 0,01). Si existe: `Q' = Q − Σ(movidos del tramo)`, `R' = R + Σ(movidos del tramo)`; si no existe, la fila **no se toca** (ver §4.3).
- **Fallo honesto:** si una **fila de cierre** (ver §4.1: `A ∈ {Componente, Subs/Cont, Total}`, **o** `B` no vacío con `C == "Total"` — cierre de empresa) con `Q != 0` no reconcilia con ningún tramo ⇒ **fail-fast `ERR-RECLASIFICACION`** nombrando ASE + fila + etiqueta + importe. **Nunca inconsistencia silenciosa.** Las filas no estructurales sin match simplemente no se tocan.

### D3 — Alcance: solo `Rem. Anticipos R2`; `Reversion Pagos R4` fuera

- `Reversion Pagos R4` **no tiene columnas Deb/Cred** (su última columna es `Otros → R`) → se declara **fuera de alcance con evidencia** (su `ColumnasR4` no contiene `Deb/Cred`; ver `WorkbookLeafCellMapDetalleR2R4.ColumnasR4`). El reclasificador no corre sobre R4. Test de alcance lo fija (ver §6, fila G6).

### D4 — Idempotencia: aplicar la regla a su propia salida es no-op

- Segunda aplicación sobre filas ya movidas (`Q = 0`, `R = Q_original`) no dispara (el cierre ya no supera el umbral en Q) y los agregados ya no reconcilian un segundo movimiento (test §6, fila U7). Garantiza que el punto único no duplique el efecto aunque el writer se re-ejecute sobre el mismo modelo en memoria.

### D5 — Origen bajo el umbral: ASE1/ASE3/ASE4/ASE5 con valores pequeños ⇒ R=0 explícito, sin cambios

- Con `|Q|` máximo de 2.232,90 (ASE3), 160,05 (ASE1), 124,42 (ASE5) y 98,04 (ASE4) no hay disparo: el writer deja **R=0 explícito** (el encabezado `null` de `ColumnasR2` ya lo hace hoy) y no toca Q. Tests §6, filas U8/I3.

### D6 — Validación de pertenencia por concepto: canal · fuente · subsección · sección (requisito del Ingeniero, 2026-10-27)

- **Norma:** ningún valor se mueve si antes no se puede nombrar su **coordenada completa** y demostrar que pertenece al canal que disparó, a la fuente R2 del ASE, a la subsección (empresa) y a la sección (`Componente`/`Subs/Cont`). Sin coordenada ⇒ `ERR-RECLASIFICACION`.
- **Diferencia con D2:** D2 valida la **aritmética** del agregado (que el tramo sume exacto); D6 valida la **pertenencia** del concepto (que el importe movido es de ese canal, de esa sección, de esa subsección y de esa fuente).
- **Evidencia:** la coordenada viaja en la traza por cada concepto movido y por cada agregado reconciliado (§4.5) y se escribe al log con el RunId (punto 5 de la regla). Test U13 + I4.
- **Prohibido:** inferir la coordenada por posición de fila, por nombre de ASE, por período o por etiqueta suelta; se deriva de los cierres de la propia secuencia (misma doctrina que `DelimitadorBloquesHoja`).

---

## §3. Diseño

### 3.1 Dónde vive (punto único de aplicación)

```
ProcesadorPeriodo / ProcesadorRemuneracion
        │  (construyen DetalleR2AseInputs = observación fiel de la fuente)
        ▼
OpenXmlPlantillaWriter.GenerarWorkbook
        │  EscribirCeldasDetalleR2R4
        ▼
EscribirDetalleR2(WorkbookPart, DetalleR2AseInputs)   ← PUNTO ÚNICO (≈línea 2049)
        │  (1) ReclasificadorDebCredR2.Aplicar(detalle)  [NUEVO, ANTES de EmparejarBloque]
        │  (2) EmparejarBloque (firma + canal efectivo — intacto)
        │  (3) EscribirValorNumerico por (encabezado → columna Q/R — intacto + guard anti-fórmula)
        ▼
VerificarDetalleR2R4 (gate existente — intacto)
```

- **Por qué el writer y no el reader ni el processor:** el modelo del reader debe seguir siendo "observación fiel de la fuente" (doctrina `DetalleR2AseInputs`); la reclasificación es **presentación administrativa auditada**, no lectura. Ponerla en el writer —antes de `EmparejarBloque`— garantiza que **ninguna ruta se salte la regla**: la usan tanto `ProcesadorPeriodo` como `ProcesadorRemuneracion` vía `GenerarWorkbook`. Punto único ⇒ R-DOBLE-APLICACION imposible por construcción + test de idempotencia.
- **Por qué NO depende de Infrastructure (§3.4):** la regla no necesita geometría de plantilla (filas, columnas Q/R, bloques por ASE). Opera sobre la secuencia tipada `DetalleR2Fila` (A/B/C/D + `ValoresPorComponente`). La justificación formal de que no depende del predicado de canal de Infrastructure está en §3.4.

### 3.2 Componente Core (nuevo, puro, testeable in-memory)

Nuevo tipo de dominio (ubicación propuesta: `Remuneracion.Core/Rules/` junto a las demás reglas puras; nombre de archivo exacto **a verificar en implementación** contra la convención del directorio):

```csharp
// Firma propuesta (pseudocódigo mínimo — el cuerpo completo vive en la implementación,
// no en este plan). Puro: sin I/O, sin período, sin ASE, sin plantilla.
public sealed record CoordenadaConcepto(
    string Hoja,          // "Rem. Anticipos R2"
    string Fuente,        // archivo R2 del ASE (ruta del insumo) — trazabilidad documental
    int AseId,
    string Seccion,       // A: Componente | Subs/Cont | Total
    string Subseccion,    // B: empresa del bloque (EAB/ENEL/NUEVO ESQUEMA/OCCIDENTE/…)
    string Canal,         // C: rótulo del cierre del grupo
    string Concepto);     // D: Vlr Servicio | Vlr Intereses | Subsidio(-)/Contribucion(+) | Total

public sealed record MovimientoReclasificacion(
    int IndiceFila,          // posición en DetalleR2AseInputs.Filas
    string Firma,            // A|B|C|D para auditoría
    CoordenadaConcepto Coordenada, // canal + fuente + subsección + sección + concepto (ver §4.5)
    decimal MovidoQ,         // importe sacado de Deb/Cred
    string Motivo);          // "cierre canal <C> |Q|>umbral" | "agregado tramo [<cierres>]"

public sealed record ResultadoReclasificacion(
    DetalleR2AseInputs Clasificado,                       // NUEVA instancia; la de ENTRADA no se muta
    IReadOnlyList<MovimientoReclasificacion> Movimientos); // traza de auditoría

public static class ReclasificadorDebCredR2
{
    public const decimal UmbralReclasificacionDebCred = 500_000m;

    public static ResultadoReclasificacion Aplicar(
        DetalleR2AseInputs detalle,
        decimal umbral = UmbralReclasificacionDebCred);   // el único call site productivo usa el default
}
```

Efecto (normativo, ver §4): **función pura sin efectos sobre la entrada** — devuelve un `DetalleR2AseInputs` NUEVO con filas NUEVAS (`ValoresPorComponente` reconstruido con las claves `Deb/Cred` y `Deb/Cred-otros`) tal que para cada fila afectada `Q' = 0`, `R' = R + Q_original`, y para cada agregado reconciliado `Q' = Q − Σ`, `R' = R + Σ`. Mutar el modelo de entrada queda **prohibido** (el writer puede re-ejecutarse y los gates comparan el modelo del reader). Devuelve además la **traza de auditoría** (lista de movimientos con firma+motivo) que el writer escribe al log (nivel Debug/Information con RunId — convención existente, sin nuevo formato de log) para cumplir el punto 5 de la regla (evidencia/auditoría).

### 3.3 Algoritmo (prototipo ya ejecutado sobre la fuente ASE2 — se formaliza, no se re-descubre)

Pseudocódigo mínimo (la implementación lo viste con tipos y tolerancias de §4):

```
1. Recorrer la secuencia en orden físico; detectar CIERRES de canal (D == "Total" && C no vacío).
2. Grupo = tramo (cierre_anterior_exclusive .. cierre_actual_inclusive] (instancia, NO etiqueta).
   Las filas posteriores al último cierre (OCCIDENTE/Componente/Subs.Cont/Total) no pertenecen a ningún grupo.
3. Si |Q_cierre| > umbral → marcar grupo: para cada fila del grupo Q' = 0, R' = R + Q.
4. Para cada fila NO-leaf con Q != 0 (en orden físico):
     buscar el tramo contiguo MÁS LARGO de cierres anteriores con Σ == Q (±0,01).
     si existe → Q' = Q − Σ(movidos del tramo ∩ grupos marcados), R' = R + lo mismo.
     si no existe y la fila es CIERRE (A∈{Componente,Subs/Cont,Total} o B!="" con C=="Total") → ERR-RECLASIFICACION.
     si no existe y no es cierre (fila de detalle) → no tocar.
```

Por qué la identidad es la **instancia** y no la etiqueta: la etiqueta `O` aparece dos veces (cuadro `Componente` y cuadro `Subs/Cont`); agrupar por `Map<C,…>` colapsa ambos y el tramo del segundo cuadro hereda el movimiento del primero (error ya reproducido en el prototipo). El algoritmo indexa cierres por **posición**, nunca por diccionario de etiqueta.

Resultado del prototipo sobre ASE2 (idéntico al oráculo — el H1 lo re-mide como test):

| Fila | Q antes → Q' | R antes → R' | Tramo usado |
|---|---|---|---|
| grupo O (Vlr Servicio, Vlr Intereses, C=O) | 54,00 / 6.759.294,81 / 6.759.348,81 → **0 / 0 / 0** | 0 → **54,00 / 6.759.294,81 / 6.759.348,81** | disparo directo |
| `OCCIDENTE` | 6.759.342,66 → **-6,15** | 0 → **+6.759.348,81** | H,O,T |
| `Componente` | 6.759.371,77 → **22,96** | 0 → **+6.759.348,81** | los 9 cierres |
| `Total` | 6.759.371,77 → **22,96** | 0 → **+6.759.348,81** | los 9 cierres |
| `EAB`/`ENEL`/`NUEVO ESQUEMA` | sin cambio | sin cambio | tramo sin el canal movido |
| detalle no-cierre (`Vlr Servicio`, `Vlr Intereses` sueltos) | sin cambio | sin cambio | sin tramo (no son agregados) |

### 3.4 Por qué la regla NO depende de Infrastructure (justificación exigida)

`CanalEfectivoDetalleR2R4` (Infrastructure-referenciado-desde-Core, predicado puro) resuelve **dónde** va cada fila en la plantilla (slot por firma + canal efectivo + segundo cierre). La reclasificación resuelve **qué valor** porta cada fila en Q/R **antes** de colocar: opera solo sobre `(A,B,C,D, Q, R)` de la secuencia del dominio, sin leer bloques, columnas Q/R, ni fórmulas. No consume `WorkbookLeafCellMapDetalleR2R4`, `DelimitadorBloquesHoja`, ni `OpenXmlPlantillaWriter`. Dependencia: Core → nada (puro). El writer la consume; ella no consume al writer. Por eso vive en `Remuneracion.Core` (dominio) y no en `Infrastructure`.

### 3.5 Archivos que el plan cita y toca (lista cerrada; nada fuera de ella sin nuevo OK)

| Archivo | Toque |
|---|---|
| `Remuneracion.Core/Models/DetalleR2AseInputs.cs` | **Cambio aditivo**: nueva propiedad `Fuente` (ruta del R2 del ASE) para la coordenada de auditoría §4.5. Sin tocar `DetalleR2Fila`, `ValoresPorComponente` ni `Firma`. |
| `Remuneracion.Infrastructure/Excel/ExcelDataReaderWorkbookLeafInputReader.cs` → `ConstruirDetalleR2` | **Cambio de una línea**: `Fuente = rutaR2` al construir el modelo (sigue siendo observación fiel: la ruta ES el origen). Sin cambios de lectura/valores. |
| `Remuneracion.Core/Models/CanalEfectivoDetalleR2R4.cs` | Referencia del predicado de canal; el plan justifica en §3.4 por qué la regla NO lo consume (se ejecuta antes de `EmparejarBloque`). Sin cambios. |
| `Remuneracion.Core/Rules/ReclasificadorDebCredR2.cs` (nuevo) | Núcleo puro + `UmbralReclasificacionDebCred = 500_000m` + traza. |
| `Remuneracion.Infrastructure/Excel/OpenXmlPlantillaWriter.cs` → `EscribirDetalleR2` (~línea 2049) | **Punto único**: invocar `Aplicar(...)` antes de `EmparejarBloque`. Sin cambiar firma pública ni cálculo ni CONSOLIDADO. |
| `Remuneracion.Infrastructure/Excel/WorkbookLeafCellMapDetalleR2R4.cs` (`ColumnasR2`) | Actualizar el XML doc que hoy declara R como invariante en 0 (`"Deb/Cred-otros R"`): R pasa a ser destino de reclasificación; Q/R se escriben por valor movido. Cero cambios de columnas. |
| `Remuneracion.Core/Errors/CodigoError.cs` + `CatalogoErrores.cs` | Nuevo `ERR-RECLASIFICACION` (mensaje UX + código de salida) **solo si el plan lo requiere**: se requiere (D2 fail-fast). Mensaje nombra ASE + fila + etiqueta + importe. Código de salida: el mismo de fuente/plantilla (contrato HU-15) — valor exacto **a verificar en implementación** contra `CodigosSalida`. |
| `Remuneracion.IntegrationTests/DetalleR2R4EscritoPlan59Tests.cs` | **Cerrar el gate**: comparar Q y R **por separado** contra el oráculo (hoy compara el par con comentario de divergencia). El comentario de divergencia se retira o se reescribe como cierre. |
| `Remuneracion.IntegrationTests/Insumos.cs` | Fixtures; `PlantillaOficialV4()` (sin cambios; solo consumo). Nuevos fixtures reutilizan `BaseDelPeriodoV4`. |

### Architecture Validation Certificate

| Principio | Veredicto |
|---|---|
| **SRP** | ✅ Reclasificador = qué valor (Q/R) por fila; EmparejarBloque = dónde va cada fila; Writer = escribir valores; Verificador = gate. Cada pieza una sola razón de cambio |
| **OCP** | ✅ Nuevo canal/empresa/ASE/período se absorbe sin código nuevo (agrupación por instancia + tramo por suma, nunca por etiqueta literal); el umbral cambia en un punto |
| **LSP/DIP** | ✅ Contratos `IPlantillaWriter`/`IWorkbookLeafWriter` intactos; `DetalleR2AseInputs` intacto como observación; cambios detrás de las mismas abstracciones |
| **Best practices** | ✅ Contract-first (§4 fija cierre/grupo/tramo antes del código); fail-fast accionable (`ERR-RECLASIFICACION` con ASE+fila+etiqueta+importe); tolerancia ±0,5 solo en comparación y 0,01 solo en reconciliación de tramo (ver §4.6); values-only; guard anti-fórmula intacto; CRLF; sin commits |
| **Performance** | ✅ O(n·m) acotado al bloque del ASE (n = filas, m = cierres; m ≤ ~12 en lo medido); una pasada por ASE; cero workbooks extra en runtime; sin Excel en gates unitarios |

---

## §4. Contrato e invariantes

### 4.1 Definiciones (normativas — todo test las cita)

- **Cierre de canal (leaf):** fila con `D == "Total"` (comparación exacta tras `Trim`, `OrdinalIgnoreCase` — misma convención que `CanalEfectivoDetalleR2R4.EsLiteralTotal`; **a verificar en implementación** que no exista variante con espacios/caso distinto en las fuentes) **y** `C` no vacío (tras `Trim`). El "canal" es la etiqueta `C`.
- **Grupo:** tramo contiguo `(cierre_anterior, cierre_actual]` en orden físico, donde el cierre actual es el leaf que lo termina. **Identidad = instancia del cierre, no la etiqueta.** Las filas posteriores al último leaf (`OCCIDENTE`/`Componente`/`Subs.Cont`/`Total` y equivalentes — vocabulario exacto **a verificar en implementación** contra la fuente y `DelimitadorBloquesHoja`) **no pertenecen a ningún grupo**.
- **Tramo de cierres:** secuencia contigua de cierres de canal (por posición) anteriores a la fila agregada. Contrato: para una fila NO-leaf con `Q != 0`, el tramo válido es el **más largo** cuya `Σ == Q` con tolerancia **0,01** (ver §4.6). Si varios tramos empatan en longitud (caso no observado), la implementación debe fallar con `ERR-RECLASIFICACION` (ambigüedad = no tocar en silencio) — comportamiento exacto **a verificar en implementación** y fijado por test.
- **Umbral:** `UmbralReclasificacionDebCred = 500_000m`. Disparo (recomendado D1): `Math.Abs(Q_cierre) > Umbral`. Alternativa literal: `Q_cierre > Umbral` (una línea).
- **Fila de cierre (agregado):** fila con `A ∈ {Componente, Subs/Cont, Total}` **o** `B` no vacío con `C == "Total"` (cierre de empresa: `EAB`/`ENEL`/`NUEVO ESQUEMA`/`OCCIDENTE` en la fuente medida). Comparación `Trim` + `OrdinalIgnoreCase`; lista/taxonomía **a re-verificar en H0** contra la fuente y la plantilla (si aparece otro rol, el T0 lo fija; prohibido congelar números de fila). Las filas de detalle (`Vlr Servicio`, `Vlr Intereses`, `Subsidio(-)/Contribucion(+)`) **no** son cierres.

### 4.2 Movimiento (qué cambia y qué no)

- Fila de grupo marcado: `Q' = 0`, `R' = R + Q_original` (donde `R` es el actual del modelo: 0 en todas las fuentes medidas, pero la fórmula lo suma por corrección).
- Agregado reconciliado: `Q' = Q − Σ_movidos`, `R' = R + Σ_movidos` (Σ = suma de los `Q_original` de los cierres del tramo que pertenecen a grupos marcados).
- Fila sin match (fila de detalle) o con `Q == 0`: **no se toca**.
- Filas del cuadro OPORTUNO con fórmula: **no se tocan nunca** (el guard las protege; se recalculan solas — punto 4 de la regla).

### 4.3 Conservación del par (invariante que cierra CONSOLIDADO/VALIDACION)

Para **toda fila tocada**: `Q' + R' == Q + R` (al centavo; tolerancia de aserto ±0,5 según doctrina). Corolario: la suma del bloque en Q+R no cambia ⇒ `VALIDACION_TOTAL` y `CONSOLIDADO` (que leen el bloque) **siguen cerrando**. El gate de blast radius (§6, fila G3) lo exige.

### 4.4 Idempotencia (invariante)

`Aplicar(Aplicar(x)) == Aplicar(x)` fila por fila. Segunda pasada: ningún cierre supera el umbral en `Q'` (todos los movidos quedaron en 0) ⇒ ningún grupo nuevo; agregados ya ajustados no re-reconcilian un segundo movimiento. Test §6 fila U7.

### 4.5 Coordenada de auditoría y VALIDACIÓN DE PERTENENCIA (requisito del Ingeniero, 2026-10-27)

Antes de mover **un solo valor**, cada concepto debe quedar clasificado en su coordenada completa y esa coordenada debe validarse. No se mueve nada que no se pueda nombrar.

| Dimensión | De dónde se deriva (R2) | Regla |
|---|---|---|
| **Hoja** | constante `Rem. Anticipos R2` | siempre |
| **Fuente** | ruta del insumo R2 del ASE, expuesta por el modelo (`DetalleR2AseInputs.Fuente`, la fija el reader al construir) | no vacía en producción; en modelos sintéticos de test se declara explícitamente |
| **ASE** | `DetalleR2AseInputs.Ase` | el bloque destino es el del mismo ASE (lo garantiza el writer) |
| **Sección** | `A` del **cierre de sección** que encierra el grupo (`A="Componente"` / `"Subs/Cont"`) | obligatoria; ausente ⇒ `ERR-RECLASIFICACION` |
| **Subsección** | `B` del **primer cierre de empresa** por debajo del grupo (`B` no vacío, `C="Total"`) — R2 no declara nivel de subsección propio; la empresa ocupa `B` (`DelimitadorBloquesHoja.R2`, nivel `EmpresaBloque`) | obligatoria; ausente ⇒ `ERR-RECLASIFICACION` |
| **Canal** | `C` del cierre del grupo | obligatorio; es la clave del disparo |
| **Concepto** | `D` de la fila (`Vlr Servicio`, `Vlr Intereses`, `Subsidio(-)/Contribucion(+)`, `Total`) | obligatorio; no vacío |

> **Nota de vocabulario (R1 vs R2) — obligatoria para no confundir al implementador:** en R1 las columnas son `A=sección`, `B=subsección`, `C=fuente`, `D=canal`, `E=concepto`. En **R2** la hoja reusa `A/B/C/D` con otro rol: `A=sección` (`Componente`/`Subs/Cont`), `B=empresa/recaudador` (lo que en R1 es la "fuente": EAB/ENEL/ENERBIT/OCCIDENTE), `C=canal` (A/E/1/2/3/5/H/O/T) y `D=concepto`. Por eso la coordenada registra la empresa en `Subseccion` y, además, `Fuente` = archivo R2 del ASE: **no se inventa una columna "fuente" que R2 no tiene**. La tabla de arriba fija el mapeo; cualquier lectura distinta es `ERR-RECLASIFICACION`.

**Validaciones duras (se ejecutan por grupo antes de mover):**

1. El grupo es **contiguo** y termina exactamente en su cierre de canal (`D="Total"`, `C` no vacío).
2. **Todas** las filas del grupo pertenecen al canal del cierre (ninguna fila del tramo trae una etiqueta de canal distinta en `C` ni una etiqueta de sección/empresa que la expulse del grupo).
3. El grupo tiene **sección** y **subsección (empresa)** derivables según la tabla anterior.
4. Cada fila del grupo tiene **concepto** no vacío.
5. La **fuente** está declarada y corresponde al ASE del bloque.
6. El importe movido por fila es el `Q` de esa fila (no se reparte ni se prorratea: el movimiento es por concepto, al centavo).

Si (1)–(5) fallan ⇒ **`ERR-RECLASIFICACION`** con mensaje accionable que nombra: hoja · ASE · fuente · coordenada parcial disponible (sección/subsección/canal/concepto) · fila y firma · motivo. Nunca se mueve "lo que se pueda" ni se omite la fila en silencio.

**La traza de auditoría (`MovimientoReclasificacion`) lleva la coordenada completa por CADA concepto movido y por CADA agregado reconciliado**, y se escribe al log con el RunId. Eso es lo que permite responder, para cualquier valor del entregable, "este importe salió de esta fuente, de esta sección, de esta subsección, de este canal y de este concepto", y es la evidencia del punto 5 de la regla.

**Caso octubre-1 (fija el test):** el grupo del canal `O` del ASE2 debe producir exactamente tres movimientos con coordenada `(Rem. Anticipos R2 · <fuente R2 ASE2> · ASE 2 · Componente · OCCIDENTE · O · Vlr Servicio | Vlr Intereses | Total)` y los agregados `OCCIDENTE`, `Componente` y `Total` reconciliados en el mismo registro. El grupo homónimo `O` del cuadro `Subs/Cont` (segunda instancia, Q=0) **no** debe producir movimiento: demuestra que la coordenada distingue sección+subsección y no colapsa por etiqueta.

---

### 4.6 Tolerancias (dos, no una — no mezclar)

- **Comparación de montos contra oráculo/fuente:** ±0,5 (doctrina madre, redondeo).
- **Reconciliación de tramo (`Σ == Q`):** 0,01 (centavo; los agregados medidos cierran al centavo: OCCIDENTE = H+O+T exacto). Un tramo que solo cierra con holgura mayor ⇒ no es tramo ⇒ fila no reconcilia ⇒ D2.

---

## §5. Fases de implementación (H0..H3 con criterio de salida medible)

> Orden sugerido: H0 → H1 → H2 → H3. Regresión continua (suite + SHAs/huella en cada paso). Build 0 warnings antes de marcar cada tarea completa. Sin commits (los hace el Ingeniero). Excel cerrado durante las corridas.

### H0/F0 — T0 read-only (bloqueante; cero código de producción)

Fixture: `Remuneracion.IntegrationTests/ReclasificacionDebCredT0Plan60Tests.cs` (BCL ZIP+XML + ExcelDataReader solo-lectura sobre copias; rutas explícitas a la fuente ASE2, al oráculo y a la salida actual §0.3–§0.5; prohibido `Insumos.Raiz()` en lo nuevo salvo el helper `Ruta` ya existente). Evidencia: `Docs/T0-Plan60-Reclasificacion-DebCred.md` (corto: tablas §0 re-medidas + veredicto).

| ID | Tarea | Verifica |
|---|---|---|
| T0.1 | Re-medir fuente ASE2 en disco: filas 5..57, headers fila 4, canal O (26/27/28), OCCIDENTE (32), Componente (33), Total (57), resto de cierres bajo umbral, R ausente | §0.3 |
| T0.2 | Re-medir oráculo por etiqueta (canal O, OCCIDENTE, Componente, Total a Q/R del §0.4) + confirmar que la comparación por fila es inválida (geometría distinta a v4) | §0.4 + §0.6 |
| T0.3 | Re-medir salida actual §0.5 (Q sin mover, R=0) + fórmulas v4 del cuadro OPORTUNO (serie `Q57..Q190` en ASE1/ASE2 y `Q299..Q522` en ASE3/4/5 — lista exacta ya fijada en §0.2 y re-verificada contra la v4 en disco) | §0.5 + §0.2 |
| T0.4 | Re-medir R=0 histórico en los 4 oráculos (2026071/2026072/2026082 + 2026101-resto-de-ASE) + R4 sin columnas Deb/Cred | §0.7 + D3 |
| T0.5 | Lista cerrada de Δ de suite + huella v4 antes/después de H0 | Gate entrada |

- **Puerta H0→H1:** fixture verde + T0.1–T0.5 cerrados + APROBADO del Ingeniero + D1 confirmada (recomendada vs literal). Si algún número difiere de §0 ⇒ STOP (la evidencia del plan no es la del disco; se re-mide, no se "ajusta" el plan en silencio).

### H1/F1 — Núcleo puro en Core (traza a §2/§3/§4)

| ID | Tarea | Traza |
|---|---|---|
| H1.1 | `ReclasificadorDebCredR2.Aplicar(...)` + `UmbralReclasificacionDebCred` + `MovimientoReclasificacion`/`CoordenadaConcepto` (traza) en `Remuneracion.Core/Rules/` | D1, D2, D6, §3.2, §4 |
| H1.2 | Unitarios sintéticos §6 (U1..U13: disparo, borde 500.000 exacto, negativo grande, dos canales, etiqueta repetida, conservación, idempotencia, no-op, coordenada incompleta, grupo no contiguo, traza por concepto) | §4 + §6 |
| H1.3 | `ERR-RECLASIFICACION` en `CodigoError` + `CatalogoErrores` (mensaje UX + código de salida; nombra hoja+ASE+fuente+coordenada+fila+importe) | D2, D6, §3.5 |
| H1.4 | `Fuente` en `DetalleR2AseInputs` + `ConstruirDetalleR2` la fija con `rutaR2` (cambio aditivo, sin tocar valores) | D6, §3.5 |
| H1.5 | Greps de cierre G-QUEMADO (parcial H1) + huella v4 + 0 warnings | §6 G-gates |

- **Puerta H1→H2:** U1..U13 verdes + 0w + huella intacta + D1 ya confirmada cableada (una sola comparación).

### H2/F2 — Punto único en el writer + XML doc

| ID | Tarea | Traza |
|---|---|---|
| H2.1 | Invocación en `EscribirDetalleR2` antes de `EmparejarBloque` + traza al log (RunId; convención existente) | §3.1, punto 5 de la regla |
| H2.2 | Actualizar XML doc de `ColumnasR2` (R deja de ser "invariante en 0"; Q/R son destino de reclasificación por valor movido) | §3.5 |
| H2.3 | e2e octubre-1 ASE2: canal O, OCCIDENTE, Componente, Total en Q y R por separado contra §0.4 | §0.4 + §6 I1 |
| H2.4 | Greps de cierre G-QUEMADO (total H2: sin `if canal/empresa/ase/periodo`, sin literales Q/R/fila) + huella + 0 warnings | §6 |

- **Puerta H2→H3:** I1 verde + 0w + huella + greps + ninguna `<f>` tocada.

### H3/F3 — Cierre oráculo + blast radius + regresión + review

| ID | Tarea | Traza |
|---|---|---|
| H3.1 | Cerrar `DetalleR2R4EscritoPlan59Tests.Q1_2026101_…`: Q y R por separado por etiqueta (I1/I2); retirar el comentario de divergencia como cierre documentado | §0.5, D2 |
| H3.2 | Blast radius 2026101 (G3): diff de celdas = exactamente Q/R del grupo movido + sus agregados; `VALIDACION_TOTAL`/`CONSOLIDADO` cierran | §4.3 |
| H3.3 | Regresión histórica 2026071/2026072/2026082 sin cambios en Q/R + fórmulas intactas (G4/G5) | §0.7 |
| H3.4 | Suite trx + set-diff vs base de entrada; fuera de lista = STOP | G-gates |
| H3.5 | Code-review independiente (`code-reviewer`, este plan como spec): **APROBADO** requerido | §8 |

- **Puerta H3→cierre:** CA-R1..CA-R8 verdes + review APROBADO + huella v4 intacta.

---

## §6. Gates y tests

### G-gates transversales (toda H)

- G-BUILD: `dotnet build Remuneracion.WinForms/Remuneracion.WinForms.slnx` → **0 warnings / 0 errores**.
- G-SUITE: `dotnet test Remuneracion.IntegrationTests/Remuneracion.IntegrationTests.csproj` + set-diff de rojos vs base de entrada: altas/bajas una por una con causa citada; **rojo nuevo fuera de lista ⇒ STOP**. 0 `Skip`, 0 tests borrados, 0 tolerancias ampliadas; re-anclajes solo con evidencia.
- G-HUELLA: huella v4 `a7476ee2…8c55` intacta antes/después de cada H (el SHA de archivo es informativo).
- G-FORMULA: ninguna celda con `<f>` escrita por nosotros (guard `EscribirValorNumerico` intacto; el cuadro OPORTUNO sigue siendo fórmula).
- G-QUEMADO: cero literales de canal/empresa/ASE/período/quincena y cero literales de fila/celda/columna Q/R en los caminos tocados —incluida la forma escondida en mapas y mensajes— (grep de cierre + candado §5-H3).
- G-WORKBOOK: todo gate de negocio lee el workbook generado (texto-`<f>` + literales; prohibido `<v>` stale) o la fuente ya leída; **cero workbooks extra en runtime**; **sin Excel en el gate**.
- G-TRAZA: cada error nuevo nombra concepto + ASE + fila/etiqueta + importe donde aplique.
- G-PROC: cero commits; cero `git checkout` / `git reset`; CRLF sin mezclar en lo tocado; Excel cerrado durante las corridas.

### Tabla test → qué prueba → archivo

| ID | Test (nombre propuesto) | Qué prueba | Archivo |
|---|---|---|---|
| T0 | `ReclasificacionDebCredT0Plan60Tests` (5–8 casos) | Re-mide §0 en disco (fuente/oráculo/salida/fórmulas/R-histórico/R4-sin-DebCred) | `Remuneracion.IntegrationTests/ReclasificacionDebCredT0Plan60Tests.cs` + `Docs/T0-Plan60-Reclasificacion-DebCred.md` |
| U1 | `DisparoSobreUmbral_MueveGrupoCompleto` | Cierre con Q=6.759.348,81 → las 3 filas del grupo a Q'=0, R'=Q; conservación Q+R por fila | Nuevo `ReclasificadorDebCredR2Tests.cs` (Core, sintético, sin I/O) |
| U2 | `Exactamente500000_NoDispara` | Cierre con Q=500.000 → no-op (borde: "mayor a", no "mayor o igual") | ídem |
| U3 | `NegativoGrande_DisparaSegunD1` | Cierre con Q=-6.000.000 → dispara con la recomendada D1; documenta que la literal no dispararía (fija D1 en código) | ídem |
| U4 | `DosCanalesSobreUmbral_MueveAmbos` | Dos cierres > umbral en el mismo bloque → ambos grupos movidos; agregado = Q − Σ ambos | ídem |
| U5 | `EtiquetaRepetidaEnDosCuadros_NoColapsa` | Etiqueta `O` en dos cuadros (dos instancias) → solo se mueve el grupo cuya instancia dispara; el otro intacto | ídem |
| U6 | `Conservacion_QMasR_Invariante` | Para toda fila tocada `Q'+R' == Q+R` | ídem |
| U7 | `Idempotencia_SegundaAplicacionNoOp` | `Aplicar(Aplicar(x)) == Aplicar(x)` | ídem |
| U8 | `SinCanalesSobreUmbral_NoOp_R0Explicito` | Modelo vacío/bajo umbral (ASE1/3/4/5) → sin cambios, R=0 | ídem |
| U9 | `AgregadoSinTramo_FilaDeDetalle_NoSeToca` | Fila de detalle sin tramo → intacta (sin throw) | ídem |
| U10 | `CierreSinTramo_FailFast` | Fila de cierre (empresa/Componente/Subs.Cont/Total) sin tramo → `ERR-RECLASIFICACION` con ASE+fila+etiqueta+importe (inyección: Q que no es suma de ningún tramo) | ídem |
| U11 | `CoordenadaIncompleta_FailFast` | Grupo sin sección derivable, sin empresa (subsección) o con `Fuente` vacía en producción → `ERR-RECLASIFICACION` nombrando la coordenada parcial (nunca se mueve "lo que se pueda") | ídem |
| U12 | `GrupoNoContiguoOFilaDeOtroCanal_FailFast` | Tramo del grupo con una fila que trae otra etiqueta de canal/sección → `ERR-RECLASIFICACION` (valida pertenencia §4.5) | ídem |
| U13 | `TrazaPorConcepto_LlevaCoordenadaCompleta` | El grupo movido emite **un movimiento por concepto** con `(Hoja, Fuente, AseId, Seccion, Subseccion, Canal, Concepto)` y los agregados se registran con su tramo | ídem |
| U14 | `SinDisparo_AgregadoQueNoReconcilia_NoFalla` | Sin ningún canal sobre el umbral, un cierre estructural que no reconcilia **no** lanza: no-op exacto (la validación de tramos es post-movimiento) | ídem |
| U15 | `DetalleQueCoincideConTramo_NoSeToca` | Fila de detalle cuyo `Q` coincide por casualidad con un tramo de cierres anteriores → **intacta**; solo se reconcilian cierres | ídem |
| I1 | `Q1_2026101_ASE2_CanalO_Reclasificado` (cierra el gate Plan59) | Salida vs oráculo **por separado**: canal O (Q=0/R=movido), OCCIDENTE (−6,15/6.759.348,81), Componente y Total (22,96/6.759.348,81), por etiqueta | `DetalleR2R4EscritoPlan59Tests.cs` (evolución del test actual) |
| I2 | `Q1_2026101_RestoDeFilas_R0` | R=0 en todas las demás filas del bloque ASE2 + ASE1/3/4/5 intactos (Q y R) | ídem |
| I3 | `Q1_2026101_ASE1345_Intactos` | ASE1/3/4/5 con `|Q|` bajo el umbral (máx 160,05 · 2.232,90 · 98,04 · 124,42) → R=0 en todo su bloque y Q/R idénticos al oráculo | ídem |
| I4 | `Auditoria_2026101_TrazaExacta` | La traza del ASE2 trae **exactamente** 3 movimientos de concepto con coordenada `(Componente · OCCIDENTE · O · Vlr Servicio / Vlr Intereses / Total)` + los agregados `OCCIDENTE`/`Componente`/`Total`; el grupo `O` del cuadro `Subs/Cont` (Q=0) **no** genera movimiento | ídem |
| G3 | `BlastRadius_2026101_SoloCeldasEsperadas` | Diff salida-con-regla vs salida-sin-regla = exactamente Q/R del grupo O + OCCIDENTE/Componente/Total (+ OPORTUNO recalculado por fórmula); `VALIDACION_TOTAL`/`CONSOLIDADO` cierran | Nuevo o extensión del fixture de blast radius (ruta temp fresca) |
| G4 | `FormulasIntactas_OportunoSigueSiendoFormula` | Ninguna celda del cuadro OPORTUNO con fórmula fue escrita (texto-`<f>` idéntico antes/después) | ídem / gate existente |
| G5 | `RegresionHistorica_SinCambios` | 2026071/2026072/2026082: Q/R byte-idénticos (ningún canal supera el umbral) | ídem (3 períodos) |
| G6 | `Alcance_R4SinDebCred_NoCorre` | R4 no contiene `Deb/Cred` (evidencia `ColumnasR4`) y la regla no lo toca | Unitario de alcance |
| G7 | `CandadoSinQuemados_Reclasificacion` | Falla si aparece literal de canal/empresa/ASE/período/quincena o de fila/celda Q/R en el camino tocado (+ injertos que deben enrojecer) | Nuevo `CandadoReclasificacionPlan60Tests.cs` (patrón Plan59-H5) |

### §6.5 Clarification Gate (requiere aprobación del Ingeniero; sin esto no se abre H0)

**Confirmado por el Ingeniero el 2026-10-27 (mensaje "procede"):** alcance solo `Rem. Anticipos R2` (3), oráculo por etiqueta (4), fail-fast de cierre (2), la puerta (5) y la **validación de pertenencia por concepto** (D6/§4.5). Queda **una sola decisión abierta**:

1. **D1 — signo del umbral:** ¿se confirma `|Q| > 500.000` (recomendada) frente a la literal `Q > 500.000`? (Si la administradora exige la literal, el cambio es de una línea.)

---

## §7. Riesgos

| ID | Riesgo | Probabilidad / Impacto | Mitigación + gate que lo detecta |
|---|---|---|---|
| R-UMBRAL | Umbral 500.000 sin fuente documentada; alguien lo "ajusta" a ojo | Media / Alto | Constante única `UmbralReclasificacionDebCred` + comentario de procedencia (2026-10-27) + D1 con alternativa en una línea; candado G-QUEMADO prohíbe duplicar el literal (U2 fija el borde) |
| R-SIN-TRAMO | Fuente futura con agregado que no es suma de ningún tramo (estructura distinta) | Media / Alto | D2: estructurales ⇒ `ERR-RECLASIFICACION` accionable (U10); no-estructurales ⇒ no tocar (U9). Nunca silencio ni remapeo |
| R-FORMULA-FUTURA | Plantilla futura convierte Q/R del bloque de detalle en fórmula | Baja / Alto | Declarado: el guard `ERR-PLANTILLA` ya aborta la escritura (G4). No rodearlo, no "actualizar la lista" |
| R-ORACULO-FILA | Comparar por fila contra el oráculo (geometría anterior) produce falsos rojos/verdes | Alta si no se prohíbe / Alto | Prohibido normativo (§0.6); I1/I2 resuelven por etiqueta + `DelimitadorBloquesHoja`; review lo exige |
| R-DOBLE-APLICACION | La regla se aplica dos veces (writer + processor, o re-ejecución) | Baja / Alto | Punto único §3.1 + U7 (idempotencia). Prohibido invocar el reclasificador fuera de `EscribirDetalleR2` |
| R-PERTENENCIA | Mover un importe que no pertenece al canal/sección/subsección/fuente identificados (o mover "lo que se pueda" cuando falta un dato) | Media / Crítico | D6 + §4.5: coordenada completa obligatoria y validada antes de mover; U11/U12 (fail-fast) + U13/I4 (traza por concepto). Prohibido inferir por posición o por etiqueta suelta |
| R-ETIQUETA-COLAPSO | `Map<C,…>` colapsa las dos instancias de `O` (ya reproducido) | Media / Alto | Identidad = instancia (§4.1); U5 lo fija con etiqueta repetida |
| R-CALCULO-CONTAGIO | Alguien "aprovecha" para tocar cálculo/CONSOLIDADO | Baja / Crítico | §9 lo prohíbe; G-SUITE + G3 (CONSOLIDADO/VALIDACION cierran) lo detectan |
| R-ORACULO-STALE | Usar `<v>` caché de celdas con fórmula del oráculo como verdad | Media / Alto | G-WORKBOOK: texto-`<f>` + literales; `<v>` de fórmula nunca es oráculo |

---

## §8. Criterios de aceptación (vinculantes)

- **CA-R0:** todo número de criterio viaja con su derivación (`fuente × etiqueta v4` + hoja). Un número sin derivación no se acepta.
- **CA-R1 (núcleo):** U1..U13 verdes; `UmbralReclasificacionDebCred` en un único punto; 0 warnings. Derivación: modelos sintéticos §6 (incluye borde 500.000 y etiqueta repetida).
- **CA-R2 (punto único):** la regla corre solo en `EscribirDetalleR2` antes de `EmparejarBloque`; `ProcesadorPeriodo` y `ProcesadorRemuneracion` la heredan vía `GenerarWorkbook` sin código propio. Derivación: grep de invocaciones (una sola) + U7.
- **CA-R3 (oráculo):** octubre-1 ASE2 coincide **por separado** en Q y R, por etiqueta: canal O (0 / movido), OCCIDENTE (−6,15 / 6.759.348,81), Componente y Total (22,96 / 6.759.348,81); R=0 en el resto; ASE1/3/4/5 intactos. Derivación: fuente §0.3 × oráculo §0.4.
- **CA-R4 (blast radius):** la corrida 2026101 difiere de la salida previa **exactamente** en Q/R del grupo O + OCCIDENTE/Componente/Total (+ OPORTUNO recalculado por fórmula); `VALIDACION_TOTAL`/`CONSOLIDADO` cierran (par Q+R conservado). Derivación: §4.3 × diff de celdas.
- **CA-R5 (no-regresión + fórmulas):** 2026071/2026072/2026082 sin cambios en Q/R; ninguna `<f>` escrita; el cuadro OPORTUNO sigue siendo fórmula. Derivación: §0.7 × G4/G5.
- **CA-R6 (candado):** candado §6-G7 verde, incluidos injertos en rojo (literal de canal/empresa/ASE/período, literal de fila/celda Q/R).
- **CA-R7 (gates):** build 0 warnings; suite con set-diff exacto (rojo nuevo = STOP); huella v4 intacta; review independiente **APROBADO**.
- **CA-R8 (pertenencia, D6):** ningún valor se mueve sin coordenada completa validada (hoja · fuente · ASE · sección · subsección/empresa · canal · concepto); coordenada incompleta o grupo no contiguo ⇒ `ERR-RECLASIFICACION` (U11/U12); la traza de octubre-1 trae exactamente los 3 conceptos del canal O con su coordenada y ningún movimiento del grupo homónimo del cuadro `Subs/Cont` (U13/I4). Derivación: §4.5 × fuente §0.3.

---

## §9. No-objetivos (NO implementar)

1. No tocar la lógica de cálculo (`CalculoRemuneracion`), el CONSOLIDADO, `DetRetri`/`DetValiRetri`, `VALIDACION_TOTAL`, ni la cadena de fórmulas.
2. No tocar la plantilla (cero `<f>`, cero `InsertRow`/`DeleteRow`, cero geometría; SHAs/huella intactos).
3. No extender la regla a `Reversion Pagos R4` (D3, fuera con evidencia) ni a ninguna otra hoja.
4. No reabrir el predicado de canal (`CanalEfectivoDetalleR2R4`), la identidad A/B (Plan 56), la delimitación por cierres (Plan 59), ni Plan 58-P2.
5. No re-congelar etiquetas, canales, empresas, filas ni períodos bajo ninguna forma (ni mapa, ni allow-list, ni "solo este ASE").
6. No incluir código productivo completo en este plan (solo firmas/pseudocódigo §3). No UI, no CLI nueva, no docs de usuario. No commits.

---

## §10. Trazabilidad con la regla del Ingeniero (punto 1..6 → dónde se cumple)

| Punto de la regla (§0.1) | Dónde se cumple (sección + test) |
|---|---|
| 1. `Total` de canal con valor > 500.000 en `Deb/Cred` se mueve a `Deb/Cred-otros` | D1 (§2) + cierre/grupo/umbral (§4.1) + U1/U2/U3 + I1 |
| 2. Se traslada el bloque completo del canal (`Vlr Servicio`, `Vlr Intereses`, `Total`), no una celda | Grupo por instancia (§3.3, §4.1) + U1/U4/U5 + I1 (las 3 filas del canal O) |
| 3. El total de la fuente del canal (OCCIDENTE) queda consistente en ambas columnas | D2 + tramo de cierres (§3.3, §4.1) + U4/U9/U10 + I1 (OCCIDENTE −6,15 / +6.759.348,81; Componente/Total 22,96) |
| 4. Lo que es fórmula no se toca (se recalcula solo) | §0.2 + §3.1 (guard intacto) + G4 (OPORTUNO sigue fórmula; OCCIDENTE/Componente/Total del bloque de detalle sí se escriben porque son literales) |
| 5. Sirve de evidencia/auditoría del resultado administrativo | Traza `MovimientoReclasificacion` con `CoordenadaConcepto` por concepto y agregado (§3.2, §4.5) + I1/I2 (Q y R por separado contra el oráculo) + I4 (traza exacta del caso) |
| 6. No quemado para Occidente ni para el canal "O": regla general de `Rem. Anticipos R2` | Punto único sin ramas por canal/empresa/ASE/período (§3.1) + U5 (etiqueta repetida) + G7 (candado) + D3 (solo R2, general) |
| **Requisito adicional (2026-10-27): validar canal · fuente · subsección · sección de cada concepto antes de reclasificar** | **D6 + §4.5** (coordenada completa obligatoria, derivada de los cierres, validada antes de mover) + **U11/U12** (fail-fast si falta o no es contiguo) + **U13/I4** (la traza prueba la pertenencia) + **CA-R8** |

---

## §11. Estado

**IMPLEMENTADO 2026-10-27.** H0..H3.5 ejecutados. Evidencia: build 0 warnings; `ReclasificadorDebCredR2Tests` U1..U15 (17 casos) verde; `DetalleR2R4EscritoPlan59Tests` I1..I4 + G3..G6 verde; candado G7 verde; corrida real 2026101 (`--cinco-ase`) `RESULTADO OK` con canal O `Q=0 / R=54,00 · 6.759.294,81 · 6.759.348,81`, OCCIDENTE `−6,15 / 6.759.348,81`, Componente y Total `22,96 / 6.759.348,81`, `R≠0` solo en esas 6 filas y las 36 celdas Q/R del cuadro OPORTUNO con `<f>` intactas; blast radius = 12 celdas, todas en `Rem. Anticipos R2`; regresión 2026071/2026072/2026082 sin cambios. Review independiente **APROBADO** (H3.5), sin bloqueantes. Pendiente de gate final CA-R7: set-diff de la suite completa en background.

**Dos endurecimientos surgidos en la revisión (ya en código y test):** (1) sin ningún canal sobre el umbral la regla es **no-op inmediato** y no valida tramos (U14), evitando un falso `ERR-RECLASIFICACION` en períodos que debían salir intactos; (2) la propagación aplica **solo a cierres** — una fila de detalle cuyo importe coincida por casualidad con un tramo no se toca (U15).

**Deuda aceptable documentada (follow-up, no bloquea):** `MarcaTotal` duplicada Core/Infrastructure (consecuencia de DIP); filtros de log por `StartsWith(Motivo)` (el prefijo es constante pura); T0.2 lee el cuadro OPORTUNO del oráculo por caché `<v>` (evidencia congelada de un archivo recalculado; los gates de negocio no usan `<v>`).
