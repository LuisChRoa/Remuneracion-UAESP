---
titulo: Plan 63 - Autochequeos M32-M36 de interventoria por quincena (DetRetri)
estado: APROBADO 2026-10-29 — aprobación completa del Ingeniero (H0→H3+Gate); mapeo Q1→N26:N30 / Q2→M26:M30
fecha: 2026-10-29
numero: 63
planes_ocupados: 01..62 ocupados; este es el 63
padre: brief del Ingeniero 2026-10-29 (mapeo exacto M32:M36 + T0 ya medido)
madre_doctrina: .opencode/project-context.md (plantilla autoridad; values-only; guard anti-fórmula; tolerancia ±0,5; prohibido `<v>` stale como oráculo; 0 warnings; quincena = dominio, nunca contenido; 73 celdas Phase 0) + plans/62 (plantilla de estilo + antecedente directo: pegado R10, E1-4-ext, R5-GATE) + plans/40 (familia F-NEW-1 E1-4, Phase 0 por dominio, R-INTERV-Q) + plans/39 (sweep T5, F-NEW-1 requiere decisión)
caso_rojo: salida 2026091 (Q1) — `DetRetri{per}!M32:M36` evalúan FALSO/FALSO/FALSO/VERDADERO(M35 ASE4)/FALSO porque restan `INTERVENTORIA!F15:F19` (redondeado) en vez de `N26:N30` (crudo); ASE4 cierra porque F18=N28. En Q2 el manual canónico ya cierra con F15:F19 (= M26:M30 redondeado) ⇒ sin impacto numérico, solo tema de referencia.
gate_entrada:
  plantilla_oficial: `Insumos.PlantillaOficialV4` (VIGENTE); medir M32:M36 y confirmar forma F15:F19 + AJUSTES (assert T0). Huella distinta ⇒ STOP
  salida_q1: Docs/Prueba Septiembre 1/Resultados/7/Remuneración 202609-1 Total.xlsx (oráculo FALSO×4+VERDADERO)
  manual_q1: manual Administrativo 2026091 (oráculo N26:N30, todos VERDADERO)
  manual_q2: Docs/Prueba Julio-2/Resultado/Remuneracion 202607-2 Total Administrativo.xlsx (oráculo F15:F19 + AJUSTES, todos VERDADERO)
  suite_base: la vigente al abrir H0 (trx de entrada citado en H0; 0 Skip; set-diff: rojo nuevo fuera de lista = STOP)
  sin_commits: NUNCA git commit / git push; NUNCA git checkout / git reset
  excel_cerrado: Excel debe estar cerrado durante las corridas
---

# Plan 63 — Autochequeos `M32:M36` de interventoría por quincena

> SDD completo PROPOSE → DESIGN → SPEC → TASKS → GATES. Fuente única de verdad para la implementación una vez aprobado ESTE documento. Ratificación directa del Ingeniero 2026-10-29 (mapeo exacto, §8): NO hay Clarification Gate abierto; H0 re-verifica el T0 ya medido con asserts blindados.
> Lectura obligatoria previa: `plans/62` (estilo + antecedente: pegado R10, extensión E1-4, R5-GATE shared-formula) + `plans/40` (familia E1-4 F-NEW-1, mecánica L32:O32/DetValiRetri por dominio) + `.opencode/project-context.md` (R-INTERV-Q, regla F-NEW-1, allow-list 73).
> Entorno: Windows, repo con espacios en la ruta; plantillas CERRADAS en Excel al correr tests (FileShare). Stack .NET 10 (Core/Infrastructure/Cli/WinForms) + Excel; persistencia por archivos Excel, NO base de datos. Español técnico; CRLF sin mezclar.
> Visual Design Intent: N/A (cero UI en este plan; WinForms/CLI solo heredan el estado y el código de salida).

---

## §0. Breadcrumb de patrones y citas (qué se reutiliza, qué cambia)

| # | Patrón / doctrina previa | Uso en este plan |
|---|---|---|
| P-40-E1-4 | `PreparadorBasePipeline.cs` (familia E1-4: `CeldasRecaudoValidacion` + `ColumnaRecaudoPorQuincena`/`ColumnaRecaudoDe(periodo)` + `AplicarVariante*` + verificación `ValidadorBasePeriodo`; R5-GATE shared `si`/`ref` del plan 62) | **EXTENDIDA** con la misma mecánica a `CeldasAutoChequeoDetRetriM` (M32→26 … M36→30, Q1→N, Q2→M). R5-GATE intacto: si existieran fórmulas compartidas bajo la región, se preservan |
| P-40-F-NEW-1 | T5: `DetRetri{per}!L32:L36` / `O32:O36`, `DetValiRetri{per}!I24:I29` ya reescritas por quincena; `INTERVENTORIA!S26:S31` pendiente con T0 propio | M32:M36 sale de "pendiente F-NEW-1" a reescrita; S26:S31 queda como **único pendiente** de la clase si sale FALSO (fuera de alcance) |
| P-62-D-G3 | Pegado R10: DetRetri/DetValiRetri son FUENTE en ambas quincenas; los auto-chequeos consumen lo pegado | El pegado NO se toca; solo cambia la referencia interventoría de M32:M36 (el resto del texto `<f>` byte-idéntico) |
| P-62-R5 | R-1…R-5 + N-1/N-5/H3-SEMANTIC: preservar `si`/`ref`, read-back semántico, signo solo inicial | Aplican tal cual a la sub-familia M (T2.2/T4.2) |
| P-29-D-E | `ComparadorSalidaVsManual` (f-vs-f + literales ±0.5, prohibido `<v>` stale) + catálogo `DivergenciasDeclaradas` | La divergencia M32-con-vs-manual puede seguir divergiendo en TEXTO (filas del layout) pero el SELL coincide ⇒ se declara así, no se versiona como exclusión nueva |

---

## §1. PROPOSE — Problema / Objetivo

### 1.1 Intent

Que los 5 auto-chequeos `DetRetri{per}!M32:M36` resten la interventoría **cruda de la quincena del run** en vez de la columna fija `F15:F19` (redondeada), con la misma mecánica de dominio ya probada en la familia R-INTERV-Q / E1-4 del Plan 40:

- **Q1:** `M32:M36` restan `INTERVENTORIA!N26:N30` (crudo, por ASE).
- **Q2:** `M32:M36` restan `INTERVENTORIA!M26:M30`.
- **El resto de la fórmula NO se toca:** el término `+'AJUSTES - SF-T'!M47:M51` se conserva byte-idéntico (en Q1 vale 0). Los cierres R1/R2/R4/BCE quedan con las filas del template canónico.

### 1.2 In Scope (H0 + H1 + H2 + H3 + Gate, prohibido ampliar sin nuevo OK del Ingeniero)

| ID | Entregable | Criterio de salida |
|---|---|---|
| **H0/T0** | Verificación read-only: re-verificar con asserts blindados los hechos ya medidos del brief (NO re-medir a mano) | Fixture verde + doc `Docs/T0-Plan63-M32-36.md`; puerta H0→H1 |
| **H1** | Sub-familia Phase 0 `CeldasAutoChequeoDetRetriM` (M32→26, M33→27, M34→28, M35→29, M36→30) + reescritura dirigida por `NumeroQuincena` (Q1→N, Q2→M) + verificación de base en `ValidadorBasePeriodo` (tercera forma ⇒ `ERR-PLANTILLA`) | CA-a/CA-b en texto-`<f>` |
| **H2** | Corridas 2026091 (Q1) y 2026072 (Q2): M32:M36 = VERDADERO×5 tras recálculo; resto del workbook sin deltas | CA-a/CA-b/CA-c |
| **H3** | Doctrina (`project-context.md`: M32:M36 pasa de "pendiente F-NEW-1" a reescrita; S26:S31 único pendiente si sale FALSO) + reclasificación comparador julio/agosto Q2 (0 altas) + doc T0 + fe de errata al plan 62 si aplica | CA-d/CA-e + suite verde con set-diff declarado |
| **Gate** | Corridas + suite completa con set-diff + build 0 warnings | CA (a–e) |

### 1.3 Non-goals / out (vinculante)

- El término `AJUSTES - SF-T!M47:M51` (se conserva; en Q1 vale 0); filas de R1/R2/R4; `INTERVENTORIA!S26:S31` (pendiente declarado, misma clase que la que sale — no se amplía); `DetValiRetri!I24:I29` (ya reescrito por el Plan 40, solo se confirma).
- Tocar la plantilla oficial (huella intacta; cero `<f>` nuevo, cero `InsertRow`/`DeleteRow` fuera de la allow-list).
- Commits, docs fuera de lo citado. NUNCA `git commit` / `git push`. NUNCA `git checkout` / `git reset`.

### 1.4 Riesgos con mitigación (resumen; detalle en §7)

| ID | Riesgo | Mitigación |
|---|---|---|
| R-PARCIAL | Reescribir de más (tocar el término AJUSTES o el comparando) y romper el cierre | D-REEMPLAZO-DIRIGIDO: solo se sustituye la referencia `INTERVENTORIA!F1x` → `INTERVENTORIA!{N,M}2x`; diff del resto del texto = ∅ (assert H1) |
| R-SHARED | Romper geometría de fórmulas compartidas si la región trajera masters | R5-GATE del plan 62: preservar `t`/`ref`/`si`; gate en T4.2-style |
| R-TERCERA | Tercera forma en base (ni F15 ni la variante del run) pasa en silencio | `ValidadorBasePeriodo` exige la variante del run; tercera forma = `ERR-PLANTILLA` |
| R-ORACULO-STALE | Comparar con cachés `<v>` de la salida Q1 | G-WORKBOOK: texto-`<f>` + evaluación BCL; `<v>` de fórmula nunca oráculo |
| R-ALCANCE | "Ya que estamos" tocar S26:S31 o DetValiRetri I24:I29 | Fuera de alcance vinculante (§1.3); S26:S31 solo se MIDE y se declara |

---

## §T0. Evidencia medida (el brief la trae medida; H0 la verifica con asserts blindados — NO re-medir a mano lo ya medido)

> Orígenes: salida `Docs/Prueba Septiembre 1/Resultados/7/Remuneración 202609-1 Total.xlsx` (Q1) + manual Administrativo 2026091 + manual Administrativo 2026072 (Q2, `Docs/Prueba Julio-2/Resultado/Remuneracion 202607-2 Total Administrativo.xlsx`) + plantilla v4 (`Insumos.PlantillaOficialV4`). H0 = fixture `AutoChequeoM32M36T0Plan63Tests` (BCL puro `ZipArchive` + `X.Linq`; recálculo Excel SOLO donde el fixture lo permita; rutas explícitas; prohibido `Insumos.Raiz()` en lo nuevo). Cero código de producción en H0.

### T0a. Salida 2026091 Q1 — el defecto (asserts T0-01…T0-03)

- **T0-01 (forma actual):** `M32:M36` traen `((...))-INTERVENTORIA!F15..F19)+M23)+'AJUSTES - SF-T'!M47` comparadas contra su `M9` (forma citada; filas del layout de salida). Assert: texto-`<f>` con `F15..F19` + término AJUSTES presente.
- **T0-02 (evaluación):** `M32:M36` = FALSO/FALSO/FALSO/VERDADERO(M35 ASE4)/FALSO. `INTERVENTORIA!F15:F19` = 189.185.988 / 266.580.256 / 154.788.536 / 120.391.083 / 128.990.446 (REDONDEADO) vs `N26:N30` = ...987 / ...255 / ...535 / ...083 / ...445 (CRUDO). ASE4: F18=N28 ⇒ VERDADERO. Assert: evaluación (si el fixture la permite) + read-back crudo-vs-redondeado por ASE.
- **T0-03 (invariante):** los demás términos numéricos son IDÉNTICOS a los del manual (solo cambian las filas por layout). Assert: diff de operandos no-interventoría vs manual = ∅ salvo filas.

### T0b. Manuales — oráculos de paridad (asserts T0-04…T0-05)

- **T0-04 (manual 2026091):** `M32:M36` restan `INTERVENTORIA!N26:N30`, sin término AJUSTES, referencias a filas del layout manual (47/49…), todos VERDADERO. Assert: texto con `N26:N30` + evaluación VERDADERO×5.
- **T0-05 (manual 2026072 Q2):** `M32:M36` restan **`INTERVENTORIA!F15:F19`** (redondeado) + AJUSTES, todos VERDADERO. En Q2 la plantilla canónica ya cierra: `INTERVENTORIA!M26:M31` (redondeado, = F) con `N26:N31` crudo. Con el mapeo pedido, Q2 pasa de `F15:F19` a `M26:M30` = MISMOS valores ⇒ sin impacto numérico en Q2; cambia solo el tema de referencia. Assert: `M26:M30 ≡ F15:F19` valor a valor + evaluación VERDADERO×5.

### T0c. Plantilla v4 + familias vecinas (asserts T0-06…T0-08)

- **T0-06 (plantilla v4):** medir `M32:M36` y confirmar que traen la forma F15:F19 + AJUSTES. Assert: texto-`<f>` con `F15..F19` y término `AJUSTES - SF-T!M47..M51`.
- **T0-07 (DetValiRetri I24:I29, Plan 40):** confirmar en plantilla/salida que usa la columna del run y VERDADERO — observar en salida Q1 si quedó `N26:N31` (medir y documentar). NO se toca. Assert: texto con la columna del run + evaluación.
- **T0-08 (INTERVENTORIA S26:S31, tercera familia F-NEW-1 del plan 39):** medir en la salida 2026091 (¿verde?): si quedó FALSO, se DECLARA como pendiente (fuera de alcance de este plan, misma clase que la que sale) — no se amplía. Assert: evaluación documentada + veredicto pendiente.

**Puerta H0→H1:** T0-01…T0-08 verdes + doc `Docs/T0-Plan63-M32-36.md`. Si un número del §T0 difiere ⇒ STOP (se re-mide, no se "ajusta" el plan en silencio).

---

## §Doctrina (cambios normativos que este plan introduce)

### D-M32M36 — La familia M32:M36 sale de "pendiente F-NEW-1" (REESCRITA por quincena)

**Texto vigente a retirar** (`project-context.md`, regla F-NEW-1 + Plan 40 §veredictos): la familia figura implícitamente entre los auto-chequeos con columna fija que "requieren decisión / T0 propio".

**Texto de reescritura para `project-context.md`** (regla F-NEW-1, allow-list):

> **Auto-chequeos `DetRetri{per}!M32:M36` por quincena (Plan 63):** restan la interventoría cruda de la quincena del run — Q1 → `INTERVENTORIA!N26:N30`, Q2 → `INTERVENTORIA!M26:M30` (misma mecánica R-INTERV-Q/E1-4 del Plan 40: `ColumnaLPorQuincena` {1→N, 2→M} + fila 26..30 por `M32→26 … M36→30`; el término `+'AJUSTES - SF-T'!M47..M51` se preserva byte-idéntico). **Allow-list Phase 0 vigente: 78 celdas** (73 + 5 de esta familia). `ValidadorBasePeriodo` exige la variante del run; tercera forma = `ERR-PLANTILLA`. La divergencia M32-con-vs-manual puede seguir divergiendo en TEXTO (filas del layout) pero el SELL del auto-check (V) coincide — declarado así, sin exclusión nueva en el comparador.

### D-S26PENDIENTE — `INTERVENTORIA!S26:S31` queda como único pendiente de la clase (si T0-08 sale FALSO)

Si la medición lo confirma, la doctrina lo declara pendiente con T0 propio futuro (misma clase que la que sale: auto-chequeo con columna fija). No se implementa, no se versiona como exclusión, no se amplía el alcance.

---

## §2. DESIGN

### D-H0. Método T0 por bytes (read-only bloqueante, cero código de producción)

Forma normativa: BCL puro (`ZipArchive` + `X.Linq`) para plantilla v4 / salida 2026091 / manuales 2026091 y 2026072; recálculo Excel SOLO donde el fixture lo permita (evaluación de M32:M36 y de I24:I29/S26:S31); rutas explícitas a `Docs/Prueba Septiembre 1/...` y `Docs/Prueba Julio-2/...`. Prohibido `Insumos.Raiz()` en lo nuevo. Cero código de producción en H0: solo fixture `AutoChequeoM32M36T0Plan63Tests` + doc `Docs/T0-Plan63-M32-36.md`. Ocho asserts (T0-01…T0-08).

| Candidata | Veredicto | Motivo |
|---|---|---|
| **M-A: sub-familia Phase 0 M32:M36 + reescritura dirigida por dominio + verificación de base** | **Elegida** | Única que cierra Q1 (FALSO→VERDADERO×5) con la mecánica ya probada R-INTERV-Q/E1-4, sin tocar el resto de la fórmula |
| M-B: reescribir la fórmula completa de M32:M36 por quincena | Descartada | Riesgo de drift en el término AJUSTES y el comparando; el reemplazo dirigido preserva el resto byte-idéntico (D-REEMPLAZO-DIRIGIDO) |
| M-C: resolver la columna por contenido (detectar N o M escritos) | Descartada | Viola D-H (dominio, nunca contenido); la columna la decide `NumeroQuincena` |
| M-D: incluir S26:S31 e I24:I29 en el mismo cambio | Descartada | I24:I29 ya está (Plan 40); S26:S31 requiere su T0 + decisión — ampliar sería la clase de defecto "ya que estamos" |

### D-H1. Sub-familia Phase 0 (decisión)

1. **Tabla nueva, misma mecánica:** `CeldasAutoChequeoDetRetriM = [(M32,26),(M33,27),(M34,28),(M35,29),(M36,30)]` en `PreparadorBasePipeline` (geometría T0-06) + `AplicarVarianteAutoChequeoDetRetriM` (reutiliza la columna por quincena: Q1→N, Q2→M; reutiliza `AsegurarVariante*`/`ReconstruirFormula*` o su generalización sin duplicar) + verificación base en `ValidadorBasePeriodo` (variante del run; tercera forma = `ERR-PLANTILLA`).
2. **Reemplazo dirigido:** SOLO se sustituye la referencia `INTERVENTORIA!F15` → `INTERVENTORIA!N26` (Q1) / `INTERVENTORIA!M26` (Q2), … `F19` → `N30`/`M30`. El resto del texto de `<f>` (término AJUSTES, comparando, paréntesis) queda byte-idéntico; assert H1: diff del resto = ∅.
3. **R5-GATE:** si la región trajera fórmulas compartidas, se preservan `t`/`ref`/`si` (misma costura del plan 62); gate en H1.
4. **Alcance:** SOLO `M32:M36` de `DetRetri{per}`. Ni una celda más de `<f>`.

### D-H2. Corridas y verificación (decisión)

1. **Q1 (2026091):** corrida desde el oficial + read-back: texto-`<f>` con `N26:N30` y evaluación VERDADERO×5 tras recálculo (read-back de texto + evaluación si el fixture lo permite, H3-SEMANTIC del plan 62).
2. **Q2 (2026072):** corrida + read-back: texto con `M26:M30`, VERDADERO×5; valores idénticos a F15:F19 (sin impacto numérico — solo cambia el tema de referencia).
3. **Resto del workbook sin deltas:** diff de `<f>` fuera de las 5 celdas = ∅; goldens ±0.5 intactos.

### D-H3. Doctrina + comparador (decisión)

1. **`project-context.md`** (textos §Doctrina D-M32M36/D-S26PENDIENTE) + allow-list 73→78 con grep de cierre Phase-0 vs values-only.
2. **Comparador julio/agosto Q2:** no empeora (0 altas). Los rojos pre-existentes pueden mejorar en COUNT (M32 texto sigue divergiendo por filas del layout) — solo set-diff declarado, nunca reverdecido a ciegas.
3. **Doc** `Docs/T0-Plan63-M32-36.md` + fe de errata al plan 62 si aplica (§10).

---

## §3. SPEC — Tasks trazadas con AC

### H0 — T0 read-only (bloqueante)

| Task | AC |
|---|---|
| T0.1 Fixture `AutoChequeoM32M36T0Plan63Tests` (T0-01…T0-03: forma F15:F19+AJUSTES, evaluación FALSO×4+VERDADERO, invariante de operandos) | 3/3 verdes; rutas explícitas; cero producción |
| T0.2 Asserts manuales (T0-04…T0-05: N26:N30 sin AJUSTES todos V; F15:F19+AJUSTES todos V; M26:M30 ≡ F15:F19) | 2/2 verdes |
| T0.3 Asserts plantilla v4 + vecinas (T0-06…T0-08: forma v4; I24:I29 columna del run; S26:S31 medida y veredicto pendiente) | 3/3 verdes |
| T0.4 Doc `Docs/T0-Plan63-M32-36.md` (matriz salida-vs-manuales + forma v4 + veredictos I24:I29 y S26:S31) | Doc existe y cita los 8 asserts |

### H1 — Sub-familia Phase 0

| Task | AC |
|---|---|
| T1.1 `CeldasAutoChequeoDetRetriM` + `AplicarVarianteAutoChequeoDetRetriM` (reutiliza columna por quincena; sin literales de período/empresa) | Compila; 0 warnings; Q1→N{26:30}, Q2→M{26:30} en texto-`<f>` |
| T1.2 Resto byte-idéntico (diff del texto fuera de la referencia = ∅) + R5-GATE (`t`/`ref`/`si` preservados; fixture con master injertado debe enrojecer si se rompe) | Verdes |
| T1.3 `ValidadorBasePeriodo`: exige la variante del run; tercera forma ⇒ `ERR-PLANTILLA` nombrando hoja+celda+variante esperada | Tercera forma sintética enrojece |
| T1.4 Allow-list 73 → 78 + grep de cierre Phase-0 (`PreparadorBasePipeline`) vs values-only | Grep verde |

### H2 — Corridas

| Task | AC |
|---|---|
| T2.1 Corrida 2026091 (Q1): `M32:M36` = VERDADERO×5 tras recálculo (read-back texto N26:N30 + evaluación) | CA-a |
| T2.2 Corrida 2026072 (Q2): `M32:M36` = VERDADERO×5 y texto con `M26:M30`; valores ≡ F15:F19 | CA-b |
| T2.3 Resto del workbook sin deltas (diff `<f>` fuera de las 5 = ∅; goldens ±0.5) | CA-c |

### H3 — Doctrina + comparador

| Task | AC |
|---|---|
| T3.1 Actualizar `project-context.md` (D-M32M36/D-S26PENDIENTE, allow-list 78) | Diff citado en el Gate |
| T3.2 Comparador julio/agosto Q2: 0 altas (set-diff declarado; COUNT puede mejorar, nunca empeorar) | CA-d |
| T3.3 Declarar divergencia M32-texto-vs-manual (SELL coincide) sin exclusión nueva | CA-e |

### Gate — Cierre

| Task | AC |
|---|---|
| G.1 Corridas 2026091/2026072 + suite completa con set-diff + build 0 warnings | CA (a–e) |

---

## §4. Criterios de aceptación mínimos

- **CA-a (Q1):** salida 2026091: `M32:M36` = VERDADERO×5 tras recálculo (read-back de texto-`N26:N30` + evaluación si el fixture lo permite).
- **CA-b (Q2):** salida 2026072: `M32:M36` = VERDADERO×5 y texto con `M26:M30` (valores ≡ `F15:F19`, sin impacto numérico).
- **CA-c (integridad):** resto del workbook sin deltas (diff `<f>` fuera de las 5 = ∅).
- **CA-d (comparador):** julio/agosto Q2 no empeora (0 altas; solo set-diff declarado).
- **CA-e (doctrina):** la divergencia M32-con-vs-manual puede seguir divergiendo en TEXTO (filas del layout) pero el SELL del auto-check (V) coincide — declarado así en la doctrina.

---

## §5. Architecture Validation Certificate

- **SOLID:** SRP (sub-familia M en tabla + aplicador + validador separados); OCP (nueva tabla citada, no ramas; columna por quincena reutilizada); LSP/ISP (extiende la mecánica E1-4 sin ensanchar interfaces); DIP (Core/dominio decide la columna; Infrastructure la materializa).
- **Best Practices:** reemplazo dirigido (resto byte-idéntico; diff = ∅ fuera de la referencia); guard anti-fórmula como red; localización por dominio (`NumeroQuincena`, cero literales de período/empresa en producción nueva); prohibido `<v>` stale como oráculo; R5-GATE shared-formula preservado.
- **Performance:** 5 celdas por corrida sobre la copia de trabajo (no sobre el oficial); sin impacto en el hot path R1/R2/R4; el T0 es read-only acotado (BCL + recálculo solo donde el fixture lo permite).

---

## §6. Ratificación directa (sin Clarification Gate abierto)

El Ingeniero ratificó el 2026-10-29 el mapeo exacto (Q1→`N26:N30`, Q2→`M26:M30`, resto de la fórmula intacto incluido AJUSTES). NO hay preguntas abiertas; H0 verifica (no re-mide) el T0 ya medido. Si un assert T0 difiere ⇒ STOP y reporte, no ajuste silencioso.

---

## §7. Riesgos y regresiones (detalle)

| ID | Riesgo / regresión | Disparo | Mitigación + test que lo fija |
|---|---|---|---|
| R-PARCIAL | Tocar AJUSTES o el comparando al reescribir | Reemplazo manual amplio | D-REEMPLAZO-DIRIGIDO (T1.2: diff del resto = ∅) |
| R-SHARED | Romper `si`/`ref` si hubiera shared en la región | `new CellFormula` ciego | R5-GATE (T1.2; fixture injertado enrojece) |
| R-TERCERA | Base con tercera forma pasa en silencio | Plantilla futura | T1.3 (`ERR-PLANTILLA` con hoja+celda+variante) |
| R-ORACULO-STALE | `<v>` como oráculo de M32:M36 | Costumbre | G-WORKBOOK en T2.1/T2.2 (texto + evaluación) |
| R-ALCANCE | Tocar S26:S31 o I24:I29 "de paso" | Inercia | §1.3 vinculante; S26:S31 solo medida (T0-08) |
| R-Q2-REF | Q2 cambia de valor al pasar F15→M26 | Si M≠F en algún ASE | T0-05 fija `M26:M30 ≡ F15:F19`; T2.2 lo re-verifica |

**Tests afectados (espectro a reclasificar en H3, lista explícita):** `BaseConsistente` (variante M32:M36 por quincena), `FNew1`/T0 de validación E1-4 (allow-list 73→78), `ComparadorSalidaVsManual` (julio/agosto Q2: 0 altas; M32 texto sigue divergiendo por layout — set-diff declarado), goldens Q1/Q2 (desplazamientos declarados: M32:M36 FALSO→VERDADERO en Q1; Q2 solo tema de referencia).

---

## §8. Traza de decisiones arbitrales (elecciones propias — para el Ingeniero)

| # | Decisión | Fundamento | Alternativa descartada |
|---|---|---|---|
| D-REEMPLAZO-DIRIGIDO | Sustituir SOLO la referencia `INTERVENTORIA!F1x` → `INTERVENTORIA!{N,M}2x`, resto byte-idéntico | El término AJUSTES y el comparando ya cierran; reescribir la fórmula completa abre drift sin beneficio | Reescritura completa de la fórmula por quincena (doble fuente de verdad del texto) |
| D-TEXTO-VS-SELL | La divergencia M32-con-vs-manual se declara en TEXTO (filas del layout) con SELL coincidente (V), sin exclusión nueva en el comparador | El auto-check ya prueba lo que debe (booleano); versionar la divergencia de filas taparía una regresión real de layout | Exclusión versionada M32 (ancha: taparía regresiones) o forzar filas del manual (rompería la geometría canónica) |
| D-S26-FUERA | `INTERVENTORIA!S26:S31` solo se MIDE (T0-08) y se declara pendiente; no entra al alcance aunque sea la misma clase | Requiere su T0 de la tabla J + decisión propia (Q1→N26:N30); ampliarlo aquí viola el STOP del Plan 39 y el §1.3 | "Ya que estamos" reescribir S26:S31 (alcance elástico, sin T0 ni decisión) |

---

## §9. Fuera de alcance (vinculante)

Término AJUSTES; filas de R1/R2/R4/BCE; `INTERVENTORIA!S26:S31` (pendiente declarado); `DetValiRetri!I24:I29` (ya reescrito, solo confirmación); cualquier otra hoja; cambios de negocio fuera de `M32:M36`; deuda `R-SLOT-Q1-ENERBIT`; F-NEW-1 hardcode M residual; RecaudosReversados; re-congelamientos fuera del espectro §7/§8; commits/push/checkout/reset.

## §10. Entregables y reporte

- **Ruta del plan (este documento):** `plans/63 - autochequeos-m32-36-interventoria-por-quincena.md`.
- **Al aprobar:** H0 produce `Docs/T0-Plan63-M32-36.md` + fixture `AutoChequeoM32M36T0Plan63Tests` (8 asserts T0-01…T0-08); H1–H3 producen código + tests citados; el Gate produce trx + set-diff + corridas 2026091/2026072.
- **Fe de errata al plan 62:** no se requiere corrección (ningún assert del Plan 62 midió `M32:M36` ni fijó su variante); solo extensión normativa: allow-list Phase 0 73 → **78** y salida de M32:M36 del estado "pendiente F-NEW-1". Si T0-08 sale FALSO, la doctrina D-S26PENDIENTE nombra a S26:S31 como único pendiente de la clase.
- **Estado:** PROPUESTO-RATIFICADO (mapeo ratificado; H0 pendiente de ejecución).
