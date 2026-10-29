# Plan 63 — Cierre H0 / H1 / H2 / Gate

**Plan fuente (única verdad):** `plans/63 - autochequeos-m32-36-interventoria-por-quincena.md` (APROBADO 2026-10-29).
**Ejecutor:** subagente `implementer`. **Fecha de cierre:** 2026-10-29.
**Alcance ejecutado:** H0 (T0 read-only) → H1 (sub-familia Phase 0 + preflight) → H2 (corridas y contrato de texto) → Gate (corridas + suite + set-diff).
**Reglas respetadas:** cero `git commit` / `push` / `checkout` / `reset`; cero escritura sobre `Docs/**/Insumos`, `Docs/**/Resultados`, manuales ni plantilla (todo corrió sobre **copias temporales**); Excel cerrado durante las corridas; CRLF.

---

## 1. Resumen ejecutivo

| Fase | Resultado | Evidencia |
|---|---|---|
| **H0** | **PASA** — 8/8 asserts T0 verdes | `TestResults/p63-h0-t0.trx` · `Docs/T0-Plan63-M32-36.md` |
| **H1** | **PASA** — fase 0 reescribe `M32:M36` por dominio; resto byte-idéntico; tercera forma = `ERR-PLANTILLA` | `TestResults/p63-h1.trx` (10/10) |
| **H2** | **PASA** — contrato de texto Q1 `N26:N30` / Q2 `M26:M30`; allow-list 73→78; integridad | `TestResults/p63-h2.trx` (5/5) |
| **Gate** | Corridas 2026091 y 2026072 **exit 0**; `M32:M36` = **VERDADERO×5** post-recálculo en ambas; **0 deltas** fuera de las 5 celdas; build **0 warnings** | §5 |
| **Suite** | **0 altas / 0 bajas** vs `p62-gate-suite.trx`: 1.799 tests · 1.776 verdes · 23 rojos (los 23 del baseline); **3 bajas legítimas** nombradas frente a la medición intermedia (§5.5) | `TestResults/p63-final-suite.trx` |

**Veredicto del plan: PASA / CERRADO.** Los criterios CA-a…CA-e se cumplen con medición (§6). El
hallazgo de datos (§7.1) quedó **RESUELTO por decisión del negocio**: el manual del administrativo se
mantiene **corregido** (ratificación del Ingeniero 2026-10-29) y las dos entradas de errata del
catálogo se **retiraron** con su cita de retiro; la errata del peso ya no existe en ninguna fuente.
No quedan hallazgos abiertos que afecten al alcance del plan.

---

## 2. H0 — T0 read-only (bloqueante)

**Entregables:** `Remuneracion.IntegrationTests/AutoChequeoM32M36T0Plan63Tests.cs` (8 asserts
T0-01…T0-08, BCL puro `ZipArchive` + `XDocument`, rutas explícitas, cero código de producción) y
`Docs/T0-Plan63-M32-36.md`.

**Medición central (bytes):**

| Elemento | Medido |
|---|---|
| Geometría | `M32→F15/26`, `M33→F16/27`, `M34→F17/28`, `M35→F18/29`, `M36→F19/30`; ref de interventoría **1×** por celda; término `+'AJUSTES - SF-T'!M47:M51` presente; comparando `=M9:M13` |
| Salida Q1 (`Resultados/7`) | `M32:M36` = **FALSO/FALSO/FALSO/VERDADERO/FALSO** |
| Interventoría | `F15:F19` (redondeado) − `N26:N30` (crudo) = **+1** en ASE 1/2/3/5, **0** en ASE 4 |
| Manual Q1 | resta `N26:N30`, sin AJUSTES, **VERDADERO×5** |
| Manual Q2 | resta `F15:F19` + AJUSTES, **VERDADERO×5**; `M26:M30 ≡ F15:F19` valor a valor |
| `DetValiRetri!I24:I29` (salida Q1) | columna **del run** (`N26:N31`), **VERDADERO×6** — no se toca |
| `INTERVENTORIA!S26:S31` (salida Q1) | **VERDADERO×6** con `R26 = N26` ⇒ **no queda pendiente** |

**Precisiones declaradas** (doc §6): signo del Δ (`+1`), fila `N31`, exclusión del término AJUSTES en
la comparación normalizada, y el veredicto verde de `S26:S31`.

---

## 3. H1 — Sub-familia Phase 0 + preflight

### 3.1 Código (producción)

| Archivo | Cambio |
|---|---|
| `Remuneracion.Infrastructure/Excel/PreparadorBasePipeline.cs` | tabla `CeldasAutoChequeoDetRetriM = [(M32,26)…(M36,30)]`; constante `DesplazamientoFilaInterventoriaM32M36 = 11`; `AplicarVarianteAutoChequeoDetRetriM` + `AsegurarVarianteAutoChequeoDetRetriM` (**reemplazo dirigido** con conteo exacto y frontera numérica); `VerificarAutoChequeoDetRetriM` (autoverificación embebida); `FragmentosAutoChequeoEsperados` ahora devuelve también la referencia **prohibida**; `CeldasAutoChequeoDeDetRetriM` (geometría pública) |
| `Remuneracion.Infrastructure/Excel/ValidadorBasePeriodo.cs` | exige la variante del run y **rechaza la referencia redondeada/mixta** con `ERR-PLANTILLA` nombrando hoja + celda + variante esperada (tercera forma) |

**Costura reutilizada (no duplicada):** `SustituirTextoFormula` (Plan 58/62) preserva `t`/`ref`/`si`
de la `<f>`; una seguidora *shared* sin texto propio hace **fail-fast** en vez de reescribirse.

### 3.2 Gates (10/10 verde — `p63-h1.trx`)

| Test | Fija |
|---|---|
| `T1_1_Phase0_M32M36_RestaLaInterventoriaCrudaDelRun` (Q1→`N`, Q2→`M`) | texto-`<f>` destino, sin la referencia redondeada, AJUSTES y comparando intactos |
| `T1_2_Phase0_RestoDeLaFormulaByteIdentico` | fórmula producida **==** fórmula del template con **solo** el token sustituido; diff de `<f>` en toda la hoja ⊆ {M32:M36} ∪ {familia L/O del Plan 40} |
| `T1_2_R5Gate_MasterInjertado_ConservaTyRefYSi` | master *shared* injertado: `t`/`ref`/`si` preservados + `ValidarPreparacion` (Plan 58) pasa |
| `T1_2_R5Gate_SeguidoraCompartida_FailFastSinBase` | seguidora sin texto ⇒ `ERR-PLANTILLA` nombrando la celda, **sin base escrita** |
| `T1_3_BasePreparada_PasaElPreflightDeSuQuincena` (Q1/Q2) | control positivo |
| `T1_3_TerceraForma_ValidadorBase_AbortaConErrPlantilla` | `F15` donde se espera `N26` ⇒ `ERR-PLANTILLA` con hoja+celda+variante |
| `T1_3_FormaMixta_ValidadorBase_AbortaConErrPlantilla` | fórmula con `N28` **y** `F17` ⇒ rechazo |

---

## 4. H2 — Contrato de texto e integridad (5/5 verde — `p63-h2.trx`)

| Test | Fija |
|---|---|
| `T4_Q1_TextoN26N30_YEquivalenciaDeCuerpoConElManual` | Q1: `N26:N30` **1×**, sin `F15:F19`, AJUSTES + comparando intactos, cuerpo normalizado **≡ manual Q1** (cuyo SELL es VERDADERO) |
| `T4_Q2_TextoM26M30_SinImpactoNumerico` | Q2: `M26:M30` **1×**, sin `F`, y `M26:M30 ≡ F15:F19` valor a valor |
| `T4_RestoDelWorkbookSinDeltas` (Q1/Q2) | hoja `AJUSTES - SF-T` byte-idéntica (módulo token de período); `L145`/`M142` conservan `t`/`ref`/`si` y texto; `M9:M13` sin cambios |
| `T1_4_AllowListFase0_CreceDe73A78ConLas5CeldasM32M36` | `73 + 5 = 78` contra la **fuente única** de la geometría; las 5 celdas existen como `<f>` en la v4 |

**Grep de cierre:** `FNew1Q1BlockerT0Tests.EsCeldaDeAllowList` incorpora `M32:M36`
(`T3_Phase0_AllowList_SoloCeldasDeclaradas` Q1/Q2 **verde** en la corrida de fase ⇒ 0 `<f>` tocadas
fuera del inventario).

**Doctrina (T3.1):** `.opencode/project-context.md` actualizado — allow-list **78**, regla
`D-M32M36` (Q1→`N26:N30`, Q2→`M26:M30`, AJUSTES intacto, divergencia declarada en TEXTO con SELL
coincidente) y `D-S26PENDIENTE` **cerrado por medición** (`S26:S31` VERDADERO×6 en Q1).

---

## 5. Gate

### 5.1 Preparación e integridad de insumos

| Hecho | Evidencia |
|---|---|
| `Get-Process EXCEL` | **0 procesos** (Excel cerrado) |
| Copias | `temp/p63-gate2/insumos-2026091` (**36** archivos), `temp/p63-gate2/insumos-2026072` (**46**), `temp/p63-gate2/plantilla/v4.xlsx` |
| SHAs de insumos (antes y después, **sin cambios**) | v4 `5FEDFD316980276945DF2AF29784EFFC5FE178C5CF1AA16F3C5D5F6F0A4E38E5` · v1 `3086AC317595C2F8E84ED24B1C5EE3E52565DD5193394E3062878E7FA49C5214` · R10 2026091 `DD0EAA7B6C03180F1B360F3D83EDADD7B4CE18B921478C6AAC775C9F50E91122` · R10 2026072 `68360DADC326E0DCA8C065DF0D7AC8563904959EA6A8B07776E18ADDE85125D3` |
| Salidas | `temp/p63-gate2/out-2026091` y `out-2026072` (jamás `Resultados`) |

### 5.2 Corridas

```
dotnet run --project Remuneracion.Cli/Remuneracion.Cli.csproj --no-build -- \
  --periodo <2026091|2026072> --carpeta temp/p63-gate2/insumos-<periodo> \
  --plantilla temp/p63-gate2/plantilla/v4.xlsx --salida temp/p63-gate2/out-<periodo>
```

| Período | Exit | runId | Duración | Archivo |
|---|---|---|---|---|
| 2026091 | **0** | `adc9135f-c6d6-4e8a-a7a5-fe680235e6aa` | 20 s | `Remuneración 202609-1 Total.xlsx` (622.109 B), sin sufijo `(sin validar)` ni TXT de novedades |
| 2026072 | **0** | `0618d072-cb87-4d53-a502-37698443f062` | 16 s | `Remuneración 202607-2 Total.xlsx` (626.688 B), idem |

### 5.3 Resultado real de `M32:M36` (fuente explícita)

**Fuente usada: recálculo con Excel (COM) sobre una COPIA en `temp/`, no el caché del archivo.**
Procedimiento: copia → `Excel.Application.CalculateFullRebuild()` → `Save` → lectura del `<v>` con
BCL. **El caché del archivo recién emitido NO es oráculo**: la plantilla lo trae en `1` y viaja
*stale* (medido: pre-recálculo `M32:M36` = `1,1,1,1,1`; por eso la cifra de abajo es post-recálculo).

**Control de que Excel realmente recalculó:** `CONSOLIDADO_TOTAL RECAUDO!E142` pasa de `0` (caché
heredado) a `17.482.962.516` (valor real) en el mismo guardado ⇒ el libro se recalculó completo.

| Período | Texto-`<f>` (celdas) | `M32:M36` post-recálculo | Veredicto |
|---|---|---|---|
| 2026091 (Q1) | `…-INTERVENTORIA!N{26..30})…` | `1,1,1,1,1` | **VERDADERO×5** |
| 2026072 (Q2) | `…-INTERVENTORIA!M{26..30})…` | `1,1,1,1,1` | **VERDADERO×5** |

**Contraste “antes” (mismo cálculo, fórmula previa):** la salida de la app previa al plan
(`Docs/Prueba Septiembre 1/Resultados/7/…`, `INTERVENTORIA!F15:F19`) da `0,0,0,1,0`
(FALSO×4 + VERDADERO en ASE4) — §2.

**Sin impacto numérico en Q2:** `INTERVENTORIA!M26:M30 ≡ F15:F19` (189.185.988 / 266.580.256 /
154.788.536 / 120.391.083 / 128.990.446) verificado en la salida recalculada.

### 5.4 Deltas del workbook (CA-c)

Comparador BCL propio (fórmulas + datos semánticos: texto exacto / montos ±0,5) sobre las 40 hojas:

| Comparación | `<f>` distintas | Datos distintos | Lectura |
|---|---|---|---|
| **2026091 nuevo vs salida previa al plan** (`Resultados/7`) | **5** = `DetRetri2026091!M32:M36` | **0** | el plan SOLO cambia esas 5 fórmulas |
| oficial v4 vs **base Q1** (Phase 0) | 68 (todas en familias declaradas: tokens `BCE`/`DetValiRetri!K`, `L9:L13`, `E142:E146`, `VALIDACION_*!C3:C8`, `DetRetri!L/O`, `DetRetri!M32:M36`, `INTERVENTORIA!R26`) | 4 (sellos de período: `CONSOLIDADO!N3`, `DetRetri!M3`, `DetValiRetri!K3`, `Valida - Control Recaudo!C2`) | cero celdas fuera del inventario |
| oficial v4 vs **base Q2** | 27 (mismas familias; celdas ya en la variante del run no se reescriben — byte-preserving) | 4 (sellos) | idem |

### 5.5 Suite completa y set-diff

**Corrida de cierre ÚNICA (2026-10-29 14:45:32 → 14:56:02), con el árbol ya cerrado y sin otra
suite en ejecución** (verificado con `Get-Process testhost` antes de lanzar; la corrida anterior
terminó 14:43:35).

```
dotnet build Remuneracion.WinForms/Remuneracion.WinForms.slnx   → 0 Warning(s) / 0 Error(s)
# logs de bin limpiados antes de contar: 0 archivos remuneracion_log_*.txt
dotnet test Remuneracion.IntegrationTests/Remuneracion.IntegrationTests.csproj --no-build \
  --logger "trx;LogFileName=p63-final-suite.trx"
powershell -NoProfile -ExecutionPolicy Bypass -File Herramientas/set-diff-trx.ps1 `
  Remuneracion.IntegrationTests/TestResults/p62-gate-suite.trx `
  Remuneracion.IntegrationTests/TestResults/p63-final-suite.trx
```

| Métrica | Baseline `p62-gate-suite.trx` | Cierre `p63-final-suite.trx` |
|---|---|---|
| Total | 1.776 | **1.799** (+23 = los 3 fixtures del Plan 63, todos verdes: T0 8 + H1 10 + H2 5) |
| Verdes | 1.753 | **1.776** |
| Rojos | 23 | **23** |
| Skip | 0 | **0** |
| Duración | — | 10 min 30 s (14:45:32 → 14:56:02) |

Salida real del set-diff:

```
baseline : 1776 tests / 1753 verdes / 23 rojos
nuevo    : 1799 tests / 1776 verdes / 23 rojos
ALTAS (rojo nuevo) : 0
BAJAS (rojo cerrado): 0
```

**0 altas y 0 bajas contra el baseline:** el espectro de rojos es **idéntico** (los 23
pre-existentes de los Planes 39–61 + `ComparadorSalidaVsManualTests` julio/agosto, deuda declarada).
Estabilidad del árbol **dentro del lapso de la corrida**: SHA256 de los 296 `.cs` de
Core/Infrastructure/IntegrationTests iguales antes y después (0 cambios) — por eso el candado
`CandadoPeriodoQuemadoPlan59Tests.H5_Produccion_SinLiteralesDePeriodoNiHojasConSufijoDePeriodo`,
que lee el **código fuente en disco**, mide el mismo árbol que el resto de la suite (**6/6 verde**).

**Bajas legítimas de alto nivel (nominales, frente a la medición intermedia de 14:00→14:24, que midió
un árbol transitorio con la edición en disco y el binario del baseline):**

```
baseline : 1799 tests / 1773 verdes / 26 rojos      (p63-gate-suite.trx, medición intermedia)
nuevo    : 1799 tests / 1776 verdes / 23 rojos      (p63-final-suite.trx, corrida de cierre)
ALTAS (rojo nuevo) : 0
BAJAS (rojo cerrado): 3
  - ComparadorPegadoH5Plan62Tests.T5_1_Septiembre1_LaSuperficieDelPegadoSoloTraeLasDeclaradas
  - CandadoPeriodoQuemadoPlan59Tests.H5_Produccion_SinLiteralesDePeriodoNiHojasConSufijoDePeriodo
  - EspecialesR4H2Plan61Tests.H21_2026091_Ase3_ElCasoDelPeso_ErrataHumanaCorregidaEnElR10
```

Las tres son **cierres**, no relajaciones: las dos primeras por el **retiro ratificado** de las
entradas de errata (`D11`/`D14`) al corregirse el dato en la fuente (§7.1), la tercera por la
corrección del literal de período que había introducido la cita de `C16` (§7.6). Los asertos siguen
midiendo: el fixture de septiembre-1 espera **exactamente 1** declarada vigente (`C16`) y verifica que
las dos retiradas **no apliquen**.

> **Timeline de mediciones (obligatorio para leer los TRX).** El árbol de trabajo recibió **a las
> 13:55:44** la edición que aplica el retiro del catálogo (§7.6). Cada TRX mide un estado distinto:
>
> | TRX | Ventana | Árbol medido | Resultado |
> |---|---|---|---|
> | `p63-fase-gate.trx` | 13:38 | Plan 63 sobre baseline | 131/133 verde; 2 rojos: `S5_2026071` (pre-existente) y `ComparadorPegadoH5…T5_1_Septiembre1` (deriva del manual, §7.1) |
> | `p63-gate-suite.trx` | 14:00→14:24 | binario del baseline + edición concurrente **en disco** | 1.773/1.799; 26 rojos = 23 baseline + 2 deriva del manual + **candado** (lee el **código fuente**, que ya traía el literal de la edición) |
> | `p63-gate-suite-final.trx` | 14:27→14:43 | árbol final (Plan 63 + retiro + corrección del literal) | **1.799 / 1.776 / 23 · 0 altas · 0 bajas** |
> | **`p63-final-suite.trx`** | **14:45→14:56** | **árbol cerrado (idéntico al anterior; 296 `.cs` con SHA igual antes/después)** | **1.799 / 1.776 / 23 · 0 altas · 0 bajas · 3 bajas legítimas vs el intermedio** |

---

## 6. Criterios de aceptación (veredicto con evidencia)

| CA | Enunciado | Evidencia | Veredicto |
|---|---|---|---|
| **CA-a** | Salida Q1: `M32:M36` = VERDADERO×5 | §5.3 (recálculo Excel sobre copia) + texto `N26:N30` (§5.3/H2) | **CUMPLE** |
| **CA-b** | Salida Q2: `M32:M36` = VERDADERO×5 y texto `M26:M30`, valores ≡ `F15:F19` | §5.3 + `T4_Q2_TextoM26M30_SinImpactoNumerico` | **CUMPLE** |
| **CA-c** | Resto del workbook sin deltas | §5.4 (5 fórmulas / 0 datos vs salida previa; base: solo familias declaradas) | **CUMPLE** |
| **CA-d** | Comparador julio/agosto Q2 no empeora (0 altas) | set-diff §5.5: `ComparadorSalidaVsManualTests` (julio/agosto) siguen en el mismo estado del baseline | **CUMPLE** (deuda pre-existente declarada) |
| **CA-e** | Divergencia M32-con-vs-manual declarada en TEXTO con SELL coincidente, sin exclusión nueva | `.opencode/project-context.md` (`D-M32M36`) + `ComparadorPegadoH5Plan62Tests.T5_1_Catalogo_*` verde (no se añadió entrada) | **CUMPLE** |
| **Gate 62-bis (CA-1)** | septiembre-1: `DetRetri`/`DetValiRetri` ≡ manual (±0,5 / texto) **salvo declaradas**, cero no declaradas | `p63-final-suite.trx`: `T5_1_Septiembre1_LaSuperficieDelPegadoSoloTraeLasDeclaradas` **verde** — 0 inesperadas en la superficie del pegado y **1** declarada vigente aplicada (`C16`); al corregirse el manual, las 2 errata dejaron de ser divergencia (§7.1) | **CUMPLE (mejorado: 3 → 1 declarada, 0 inesperadas)** |

---

## 7. Hallazgos

### 7.1 HALLAZGO de datos (RESUELTO): el manual del administrativo quedó **corregido por el negocio**

- **Qué pasó:** `Docs/Prueba Septiembre 1/Resultados/Administrativo/Remuneracion 202609-1 Total.xlsx`
  está **modificado en el working tree** (no por este plan): `git status` lo marca `M`, `LastWriteTime`
  2026-10-29 **11:50:41** (antes de esta sesión) y su SHA pasó de
  `3CE1B869…E420916` (el que citó el cierre del Plan 62) a `9EF8F023…87383AC6`.
- **Efecto medido:** el manual ya trae `DetRetri!D11 = 10.891.338.739` y `D14 = 63.145.073.622`
  (idénticos al R10). La versión de `HEAD` (extraída con `git archive` a temp) trae `…740` / `…623`.
  Es decir: la **errata de 1 peso del manual desapareció** porque el archivo fue corregido.
- **Veredicto ratificado por el Ingeniero (2026-10-29): el manual queda CORREGIDO POR EL NEGOCIO.**
  No se restaura la versión de `HEAD`: el dato vigente es el bueno (coincide con el R10 y con la
  salida de la app). En consecuencia se **RETIRAN** las dos entradas de errata del catálogo
  `DivergenciasDeclaradasComparador` (`P62-DETRETRI-D11-ERRATA-PESO-ASE3` y
  `P62-DETRETRI-D14-TOTAL-ERRATA-PESO`): su condición de retiro quedó **CUMPLIDA** (el dato fue
  corregido en la fuente).
- **Terminación del retiro (mecanismo auditable, mismo patrón que `DivergenciasDeclaradasCatalogo`):**
  cada entrada retirada conserva `Retirada = true` + `FechaRetiro` + `MotivoRetiro` +
  `EvidenciaRetiro` (fecha, motivo y evidencia **citados**; `RetiroDocumentado` lo verifica) y **ya no
  aplica** ni en `Buscar` ni en `ComparadorSalidaVsManual.DivergenciasDeclaradas` (que expone
  `Vigentes`). Estado final: **`Vigentes` = 1** (`P62-DETRETRI-C16-PIE-INTERVENTORIA`) y
  **`Retiradas` = 2**. No se declaró ninguna entrada ancha: la vigente mantiene **celda + par exacto**
  (predicado del pie de interventoría) y su **valor manual re-verificado hoy** es
  `DetRetri!C16 = "* Costos de Interventoría: REMUNERACIÓN 202604-1"` (pie heredado; ningún valor
  medido se alteró).
- **Fixtures que cinchaban `D11`/`D14` del manual de septiembre-1, reclasificados con cita:**
  1. `ComparadorPegadoH5Plan62Tests.T5_1_Septiembre1_LaSuperficieDelPegadoSoloTraeLasDeclaradas` —
     expectativa **3 → 1** declarada aplicada (`C16`) + anti-recurrencia del retiro
     (`Retiradas.Count == 2`, `RetiroDocumentado`, `FechaRetiro = 2026-10-29`, ninguna vigente en
     `D11`/`D14`, y `EsDivergenciaDeclarada(D11, 739, 740) == false`).
  2. `ComparadorPegadoH5Plan62Tests.T5_1_Declaracion_EsExclusivaDelParMedido` — el par de 1 peso ya
     **no** está declarado (antes `True`, ahora `False` con la cita del retiro).
  3. `EspecialesR4H2Plan61Tests.H21_2026091_Ase3_ElCasoDelPeso_ErrataHumanaCorregidaEnElR10` — el
     manual pasa de `…740` (Δ +1) a `…739` (Δ 0) y el aserto se invierte a Δ = 0 **en las dos
     fuentes**, como anti-recurrencia. El **nombre del test se conservó** (la métrica es la misma;
     solo cambió el dato de la fuente) para no mover el espectro del TRX.
- **Restricción respetada:** el archivo del manual **no se tocó** en ningún momento (la verificación
  fue de lectura: `Get-FileHash` + lectura de celdas) y no se relajó ningún aserto.
- **Impacto en el plan:** ninguno sobre CA-a…CA-e; el T0 midió contra los bytes vigentes y sus asserts
  pasan con ambas versiones del manual.

### 7.2 `INTERVENTORIA!S26:S31` — veredicto: **VERDE**, no queda pendiente

Medición en la salida 2026091 recalculada: `S26:S31` = **VERDADERO×6** (`S26 = Q26=R26`,
`S27`, `S28` maestro *shared* `ref=S28:S31 si=2`), con `R26 = N26` (Phase 0 del Plan 41/T3) y
`Q26 = CONSOLIDADO!L104`. **No aplica** la sección “PENDIENTE declarado” de `D-S26PENDIENTE`: la
familia no queda como deuda de la clase y, aun así, el alcance **no se amplió** (no se modificó
`S26:S31`).

### 7.3 R-ORACULO-STALE confirmado en el entregable

El `<v>` de `M32:M36` en el archivo recién emitido es **`1` stale** (heredado del template), igual que
antes del fix: quien mire solo el caché **no distingue** el defecto. La demostración válida es
texto-`<f>` + recálculo (§5.3). Refuerza la doctrina “prohibido `<v>` como oráculo”.

### 7.4 Hallazgo menor (proceso): la corrida de suite anterior se interrumpió

La primera corrida completa de la suite (sesión previa) fue abortada por un reinicio del servidor. La
evidencia que manda es la **corrida de cierre** `p63-final-suite.trx` (§5.5), fresca y única, lanzada
con `Get-Process testhost` en 0 justo antes de arrancar; `p63-gate-suite.trx` y
`p63-gate-suite-final.trx` quedan como mediciones del tránsito del árbol (§5.5, timeline).

### 7.5 Red pre-existente tocada por la corrida de fase

`BasePeriodoPreflightTests.S5_OficialDirecto_EndToEnd_L9L13PorQuincena(2026071)` sigue **rojo** con el
mismo mensaje del baseline (`ERR-PLANTILLA … ASE5 … F682 extemporáneo`): deuda Q1 declarada desde el
Plan 42, no una regresión.

### 7.6 Cierre del retiro del catálogo durante el Gate (edición concurrente, ahora ratificada)

- **Qué pasó:** mientras corría la suite intermedia (14:00), el árbol recibió **a las 13:55:44** la
  edición que **aplica el retiro** de las dos entradas de errata en cuatro archivos:
  `Remuneracion.Infrastructure/Excel/DivergenciasDeclaradasComparador.cs`,
  `Remuneracion.Infrastructure/Excel/ComparadorSalidaVsManual.cs`,
  `Remuneracion.IntegrationTests/ComparadorPegadoH5Plan62Tests.cs` y
  `Remuneracion.IntegrationTests/EspecialesR4H2Plan61Tests.cs`. Esa edición es exactamente la
  terminación del retiro ratificado en §7.1; en su momento se documentó como concurrencia porque el
  Gate había arrancado antes.
- **Efecto inmediato:** el **candado** (`CandadoPeriodoQuemadoPlan59Tests`) pasó a **rojo** porque el
  texto de cita de `P62-DETRETRI-C16-PIE-INTERVENTORIA` quedó con un **literal de período en código**
  (`REMUNERACIÓN 202604-1`). El candado lee el **código fuente en disco**, así que enrojeció incluso
  con el binario anterior compilado (lo captura `p63-gate-suite.trx`).
- **Corrección aplicada (quirúrgica, sin cambio de comportamiento):** se sustituyó **solo ese
  fragmento de cita** por una descripción equivalente que remite a este documento ("C16 conserva el
  pie de interventoría con el período heredado (valor y cita medidos en `Docs/Plan63-Cierre.md`
  §7.1)"). El valor medido (`202604-1`) **vive en la documentación, no en el código**: es la forma
  correcta, porque el candado prohíbe literales de período en producción. No se tocó el mecanismo de
  retiro ni ningún aserto ajeno.
- **Verificación en la corrida de cierre:** candado **6/6 verde** y, con la instantánea SHA256 de los
  296 `.cs` idéntica antes y después de la suite, la lectura de fuente del candado corresponde al
  mismo árbol que el resto de la prueba.
- **Consecuencia sobre la evidencia:** los dos rojos de deriva (§7.1) quedaron **verdes** por el
  retiro, que ahora es parte del cierre del Plan 63 (decisión del Ingeniero) y no una edición ajena.

---

## 8. Archivos del plan

**Producción**
- `Remuneracion.Infrastructure/Excel/PreparadorBasePipeline.cs` (tabla + aplicador dirigido + autoverificación + geometría pública)
- `Remuneracion.Infrastructure/Excel/ValidadorBasePeriodo.cs` (variante del run + rechazo de tercera forma)
- `Remuneracion.Infrastructure/Excel/DivergenciasDeclaradasComparador.cs` (cierre §7.1/§7.6: `Retirada`/`FechaRetiro`/`MotivoRetiro`/`EvidenciaRetiro` + `Vigente`/`Vigentes`/`Retiradas`; retiro de `D11`/`D14`; `Buscar` ignora retiradas; cita de `C16` sin literal de período)
- `Remuneracion.Infrastructure/Excel/ComparadorSalidaVsManual.cs` (`DivergenciasDeclaradas => Vigentes`)

**Tests**
- `Remuneracion.IntegrationTests/AutoChequeoM32M36T0Plan63Tests.cs` (8)
- `Remuneracion.IntegrationTests/AutoChequeoM32M36H1Plan63Tests.cs` (10)
- `Remuneracion.IntegrationTests/AutoChequeoM32M36H2Plan63Tests.cs` (5)
- `Remuneracion.IntegrationTests/FNew1Q1BlockerT0Tests.cs` (allow-list `M32:M36`)
- `Remuneracion.IntegrationTests/ComparadorPegadoH5Plan62Tests.cs` (expectativa 3 → 1 + anti-recurrencia; par de la errata ya no declarado)
- `Remuneracion.IntegrationTests/EspecialesR4H2Plan61Tests.cs` (H21: manual `…739`, Δ 0 en ambas fuentes)

**Documentación**
- `Docs/T0-Plan63-M32-36.md`, `Docs/Plan63-Cierre.md` (este documento)
- `.opencode/project-context.md` (allow-list 78 + `D-M32M36` + `D-S26PENDIENTE` cerrado)

**Evidencia**
- `Remuneracion.IntegrationTests/TestResults/p63-h0-t0.trx`, `p63-h1.trx`, `p63-h2.trx`, `p63-h2-fase.trx`, `p63-fase-gate.trx`, `p63-gate-suite.trx` (medición intermedia, §5.5), `p63-post-edicion.trx`, `p63-gate-suite-final.trx`, **`p63-final-suite.trx` (corrida de cierre)**

---

## 9. Limpieza e integridad

| Acción | Resultado |
|---|---|
| `temp/p63-gate` y `temp/p63-gate2` (copias, salidas, recálculos) | **conservados** como evidencia de esta sesión (fuera de `Docs/`, no versionados); se pueden borrar sin afectar el repo |
| Logs de bin de tests | limpiados **antes** de la corrida de cierre: **0 archivos** `remuneracion_log_*.txt` |
| Insumos re-verificados al cierre (SHA256, idénticos a §5.1) | v4 `5FEDFD31…4E38E5` · v1 `3086AC31…9C5214` · R10 2026091 `DD0EAA7B…E91122` · R10 2026072 `68360DAD…5125D3` |
| Manual del administrativo (lectura, jamás escritura) | SHA256 `9EF8F023…7383AC6` antes y después; `DetRetri!C16` = `REMUNERACIÓN 202604-1`, `D11` = 10.891.338.739, `D14` = 63.145.073.622 (leídos hoy) |
| Estabilidad del árbol durante la suite | 296 `.cs` (Core/Infrastructure/IntegrationTests) con SHA256 idéntico antes/después: **0 cambios** (`temp/p63-pre-suite-hashes.txt` vs `…-post-suite-hashes.txt`) |
| `Docs/**/Insumos`, `Docs/**/Resultados`, manuales, plantilla v1/v4 | **intactos** |
| Commits / checkout / reset | **ninguno** |

---

## 10. Meliorates (mejoras de estilo y proceso)

Lecciones que esta fase dejó medidas; ninguna amplía el alcance del plan.

| # | Meliorate | Por qué (hecho medido) |
|---|---|---|
| M-1 | **Un candado que lee fuente debe correr con instantánea de hashes.** | El candado de período enrojeció por un **literal en una cita** mientras el binario era otro. La instantánea SHA256 pre/post de los `.cs` convierte "verde" en un enunciado verificable (§5.5/§7.6). |
| M-2 | **Las citas de datos medidos van a la documentación, no al código.** | `REMUNERACIÓN 202604-1` en un `string` de producción es un literal de período prohibido; el mismo texto en `Docs/` es evidencia legible. Regla: el código cita el **documento**, el documento cita el **valor**. |
| M-3 | **Mecanismo único de retiro para todo catálogo declarado.** | Se replicó el patrón de `DivergenciasDeclaradasCatalogo` (`Retirada`+fecha+motivo+evidencia, `Vigentes`/`Retiradas`) en lugar de borrar filas: la historia queda auditable y `Buscar` no puede "revivir" una divergencia cerrada. |
| M-4 | **Expectativa de test = cardinalidad del catálogo vigente, nunca lista literal de tres.** | El aserto pasó a `Assert.Equal(["P62-DETRETRI-C16-…"], …)` + anti-recurrencia: si mañana entra una declarada nueva, el test la nombra en vez de romperse por conteo. |
| M-5 | **No renombrar tests al re-clasificar datos.** | El nombre de `H21_…ErrataHumanaCorregidaEnElR10` se conservó con el aserto invertido a Δ = 0; renombrarlo habría generado altas/bajas espurias en el set-diff. El nombre describe la **intención**, el aserto mide el **dato**. |
| M-6 | **Una sola suite «de cierre», con nombre propio y con la puerta cerrada.** | `Get-Process testhost` en 0 + limpieza de logs + `p63-final-suite.trx`: la medición es reproducible y no se confunde con las intermedias del timeline (§5.5). |
