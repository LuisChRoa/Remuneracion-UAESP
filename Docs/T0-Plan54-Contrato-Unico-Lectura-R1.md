# T0 — Plan 54 — Contrato único de lectura R1 (Q1/Q2)

**Estado H0:** T0.1–T0.5 **VERDE**. T0.6 **cerrado** con **Δ-1 = GAP que BLOQUEA la puerta H0→H1**.
**Fecha:** 2026-10-22 (post-enmienda 1)
**Plan:** `plans/54 - contrato-unico-lectura-r1-q1-q2.md` (APROBADO + enmienda 1)
**Fixture:** `Remuneracion.IntegrationTests/ContratoUnicoLecturaR1T0Tests.cs` — **58 casos, 0 rojos, 0 skip**.
**Tipo:** evidencia por bytes (BCL ZIP + XML crudo). Cero código de producción tocado en H0.

> Invariante aplicada: ninguna celda con `<f>` se usa por su `<v>` cacheado. Los números de criterio
> viajan siempre con su derivación `fuente × <f>` + hoja/celda (CA0).

---

## 0. G-SHA — plantillas intactas (antes / después de H0)

| Plantilla | SHA-256 medido | Veredicto |
|---|---|---|
| v2 `Docs/Plantilla estandarizada/v2/Formato Remuneracion 2026xx-x Total vf.xlsx` | `0F192AD2721C77986B6A8CB85D1A92F782DCD2F69989F362E066FA17027F2DAB` | **IDÉNTICO** (antes y después) |
| v1 `Docs/Plantilla estandarizada/Formato Remuneracion 2026xx-x Total.xlsx` | `3086AC317595C2F8E84ED24B1C5EE3E52565DD5193394E3062878E7FA49C5214` | **IDÉNTICO** (antes y después) |

Nota de trazabilidad: el literal del SHA v2 en el frontmatter del plan (`…DCD2F699F362E…`) tiene 62
caracteres (le falta `89`); el valor de 64 hex medido en disco es el de la tabla y es el que aserta el
fixture. El archivo no cambió: el typo es de transcripción del plan, no de bytes.

---

## 1. T0.1 — Geometría v2 R1 por bytes (S1.1) — **VERDE** (F6 adoptado sin corrección)

Hoja `Reporte Componentes R1` = `xl/worksheets/sheet10.xml` (rId10).

| ASE | Slots editables (`<v>0</v>` SIN `<f>`) | TOT_OPT `<f>` | EXTEMP `<f>` |
|---|---|---|---|
| 1 | F16, L16, F32, L32, F47, L47, F59, F70 | `F75=F47+F70+F16-L16-L47` | `F77=F59+F32-L32` |
| 2 | F120, L120, F146, L146, F171, L171, F190, F208 | `F254=F171+F208-L171+F120-L120` | `F256=F190+F146-L146` |
| 3 | F299, L299, F315, L315, F330, L330, F342, F353 | `F415=F330+F353+F299-L299-L330` | `F417=F342+F315-L315` |
| 4 | F460, L460, F486, L486, F511, L511, F530, F548 | `F568=F511+F548+F460-L460-L511` | `F570=F530+F486-L486` |
| 5 | F613, L613, F629, L629, F648, F660 (**6**) | `F680=F674+F648+F613-L613-L648` | `F682=F660+F629-L629` |

- Las **10** `<f>` de §0.2-F6 se **adoptan sin corrección**, incluidas `F415`/`F417` (ASE3) y
  `F680`/`F682` (ASE5).
- Los slots son literales `0` sin fórmula ⇒ escribibles. Simetría exacta con
  `WorkbookLeafCellMapQ2.R1Q2EditablesPorAse` (aserción en el fixture).
- El "literal 0 sin fórmula" del T0 v1 **ya no aplica** a los visibles.
- ASE5: `F674` **no** está en el mapa (forma de 2 filas) y vale literal `0`; `F680` lo consume como
  término nulo (D6). Asertado aparte.

---

## 2. T0.2 — Censo octubre-1 por bytes (S1.2) — **VERDE**

Fuente: `Docs/Prueba Octubre-1/Insumos/<N-ASE>/Recaudoporcomponente_…` hoja `Sheet1`.
Firma `Mes/Total` = `B='Mes' ∧ C='Total'`; `Aplicacion` = `B ⊃ 'Aplicacion nuevos x reversion' ∧ C='Total'`;
cierre = `A='Total' ∧ B` vacío.

| ASE | `Mes/Total` | `Aplicacion` | Cierre (col F) | Σ `Mes` | Σ `Aplic` |
|---|---|---|---|---|---|
| 1 | 8, 33, 52 | 18, 41 | 54 → `14.835.602.537,78` | `14.834.953.667,78` | `648.870,00` |
| 2 | 11, 38, 57 | 23, 46 | 59 → `11.238.640.614,90` | `11.228.935.640,90` | `9.704.974,00` |
| 3 | 8, 24, 36 | **ninguna** | 38 → `8.830.777.878,00` | `8.830.777.878,00` | `0,00` |
| 4 | 11, 46, 70 | 31, 59 | 72 → `10.504.168.126,00` | `10.482.610.949,00` | `21.557.177,00` |
| 5 | 23, 38 (**2**) | **ninguna** | 40 → `7.556.633.218,00` | `7.556.633.218,00` | `0,00` |

Hecho estructural (gobierna todo T0.3): **la fila `A=Total` de la fuente NO es un visible** — es
`Σ Mes + Σ Aplic` y **no resta** `SERVICIO ESPECIALES`. Verificado por identidad en los 5 ASE.
F5 del plan queda confirmado íntegro.

---

## 3. T0.3 — CA0/CA2/CA3/CA7 por derivación (S1.3) — **VERDE**

Semántica única (`MapearR1Q2`, Plan 25 D-C/D-F — no se re-arbitra):
`D9 = Σ F(Mes) − Σ Esp(Mes sin la última)` · `D28 = E(Componente) + E(Subs/Cont) − K(Especiales)`
(`K = 0` si no hay encabezado `Especiales`) · `D47 = Σ F(Aplic) − Esp(Aplic[0])` ·
`D66 = D(Total R4)` · `D104 = D9 + D28 + D47 + D66 + D85` (Q1 ⇒ `D85 = 0`).

| ASE | `D9` TOT_OPT | `D28` R2 | `D47` EXTEMP | `D66` R4 | `D104` | `ROUND` | R10 oficial | Δ |
|---|---|---|---|---|---|---|---|---|
| 1 | `14.815.021.974,75` | `30.588.071,02` | `648.870,00` | `-3.833.780,00` | `14.842.425.135,77` | `14.842.425.136` | `$14.842.425.136` | **0** |
| 2 | `11.219.101.305,09` | `95.634.974,32` | `9.704.974,00` | `-21.803.648,45` | `11.302.637.604,96` | `11.302.637.605` | `$11.302.637.605` | **0** |
| 3 | `8.811.472.160,34` | `14.574.412,03` | `0,00` | `-2.259.272,00` | `8.823.787.300,37` | `8.823.787.300` | `$8.823.787.300` | **0** |
| 4 | `10.473.967.813,42` | `29.206.357,07` | `21.557.177,00` | `-41.558.585,40` | `10.483.172.762,09` | `10.483.172.762` | `$10.483.172.762` | **0** |
| 5 | `7.544.253.279,94` | `10.379.542,05` | `0,00` | `-3.964.157,00` | `7.550.668.664,99` | `7.550.668.665` | `$7.550.668.665` | **0** |

Oráculo: `Docs/Prueba Octubre-1/Insumos/R10_Remuneracion_2026101.xlsx`, hoja `DetRetri2026101`,
celdas `D9:D13` (columna `RECAUDO TOTAL`), emitido por PROCERASEO 18/10/2026 16:03. **5/5 exactos al
peso, sin tolerancia.**

- **CA3** — ASE3: `L330 = 19.305.717,66` **sí se resta** ⇒ `F415 = 8.811.472.160,34` (no `8.830.777.878`,
  que es la fila `A=Total` de la fuente). EXTEMP `0,00` en ASE3 (`F417`) y ASE5 (`F682`) por **cero**
  filas `Aplicacion`. ASE5 `TOT_OPT = 7.544.253.279,94` (`L613 = 12.379.938,06`).
- **CA7** — ASE1: `Aplic` 18 = `636.047,07`, 41 = `12.822,93` ⇒ EXTEMP `648.870,00`; el
  `subsidio[0]` (fila 35 = `−32.099,64`) queda **fuera** (término v1 erradicado).
- **Valores v1 descartados** (`18.052.547.615` / `11.292.932.631`, del log de la corrida fallida):
  reconstruidos al centavo como semántica v1 y **no** se aceptan como criterio (CA0).

Evidencia cruzada adicional (script BCL, no fixture): la misma derivación reproduce el R10 oficial de
**julio-1** (5/5: `16.758.167.585` / `12.227.292.053` / `10.127.195.599` / `10.553.776.376` /
`8.543.663.209`) y de **julio-2** (5/5 incluyendo `D85` de Q2). Los tres períodos cuadran con la
semántica única ⇒ la ruta que el plan adopta es la que el proceso manual ejecuta.

---

## 4. T0.4 — Caracterización R2/R4 Q1-vs-Q2 (S1.4) — **VERDE**

### 4.1 Valores idénticos sobre las MISMAS fuentes

Método: `LeerLeafInputs` con `Periodo` Q1 (`2026071`) y Q2 (`2026072`) sobre el mismo triple R1/R2/R4;
se compara `E15/E26/K15/CeldasPorAse` (R2) y `D9/P9/CeldasPorAse` (R4). La única variable es la quincena.

- **8 de 10 combos** (5 ASE × julio-1 + julio-2): **igualdad exacta, sin tolerancia**.
- **2 combos con rama no viable** (declarados como dato en el fixture, no como excepción silenciosa):
  | Combo | Rama | Causa (ordinal congelado, ajena a R2/R4) |
  |---|---|---|
  | ASE1 + julio-1 | **Q2** | `MapearR1Q2` exige rol core `Mes2` y la fuente trae 2 filas `Mes` |
  | ASE5 + julio-2 | **Q1** | `MapearR1 case 5` exige 3 filas `Mes` y la fuente trae 2 |
  El fixture exige que el conjunto de abortos **⊆** esa lista: cualquier aborto nuevo = rojo.
- **Celdas destino literales**: los 15 destinos R2 (`E21/E32/K21`, `E101/E123/K101`, `E193/E204/K193`,
  `E327/E349/K327`, `E426/E440/K426`) y los 10 R4 (`D15/P15` … `D358/P358`) son literales `0` sin `<f>`
  en la v2 ⇒ escribibles. Ninguna es fórmula (no hay clase `R-MAPA-V2-R2-ASE4` en el mapa por ASE).

### 4.2 Veredicto

`MapearR2`/`MapearR2Q2` y `MapearR4`/`MapearR4Q2` **no divergen en valores**: comparten predicados
(`Componente/Total`, `Subs/Cont/Total`, `Total` con `B` vacío), índices de columna (E = 4, D = 3) y
tabla destino idéntica. El colapso de D4/T3 queda **habilitado**.

---

## 5. T0.5 — Inventario de consumidores y forma-fuente (S1.5) — **CERRADO**

### 5.1 `WorkbookLeafCellMapPorAse.EditableLeafCellsPorAse` (geometría v1)

| Consumidor | Cita | Veredicto |
|---|---|---|
| `ValidarFormulasProtegidasMultiAse` — usa **solo `.Keys`** (lista de ASE, no celdas) | `OpenXmlPlantillaWriter.cs:759` | **Re-apuntar** a `R1Q2EditablesPorAse.Keys` |
| `ObtenerEditables(aseId)` (helper) | `WorkbookLeafCellMapPorAse.cs:141` | **Aislar LEGACY (geometría v1)** |
| `EscribirCeldasLeafPorAse` rama Q1 | `OpenXmlPlantillaWriter.cs:1052-1080` (helper en `:1060`) | **Re-apuntar a `R1Q2EditablesPorAse`** y colapsar el `if (esQuincena2)` |

Evidencia de que **no** es la geometría productiva: `EditableLeafCellsPorAse[1]` =
`F25/F41/L25/F30/F10/L10` y su intersección con la v2 (`F16/L16/F32/L32/F47/L47/F59/F70`) es **vacía**;
`F75` (el visible real) no menciona `F25` ni `F41`. Asertado en el fixture.

### 5.2 `F30` / `F10` / `L10` (legacy)

| Consumidor | Cita | Veredicto |
|---|---|---|
| Lector ruta única (Q2 vigente) → `0m` | `ExcelDataReaderWorkbookLeafInputReader.cs:1305-1307` | **Conservar** (D2: constante, nunca rama) |
| Lector rama Q1 (`ObtenerF30`/`ObtenerF10`) | `…:1575-1577` + `…:2042-2071` | **BORRAR** con la rama Q1 |
| `EscribirDetalleR1` (path legacy) | `OpenXmlPlantillaWriter.cs:1017-1019` | **Aislar LEGACY-only, fuera de paridad** |
| `WorkbookLeafCellMap.EditableLeafCells` (lista v1) | `WorkbookLeafCellMap.cs:19-21` | No productivo: alimenta solo `EscribirCeldasLeaf` (`:1010-1044`) |
| Herramienta de diagnóstico | `Herramientas/VerificadorRecaudo/Program.cs:367` | **Intacto** |
| Paridad CLI | `ParidadCliTests.cs:135-137` | **Re-anclar** (H1.6) |
| Fixtures sintéticos | `AjustesSfTTests.cs:310`, `ValidadorBasicoTests.cs:77-79` | **Re-anclar** solo si el valor cambia (hoy ya usan `0m`/valores fijos) |

Grep de cierre (H1): cero lectores **nuevos** de `F30/F10/L10`.

### 5.3 Forma-fuente R1 por período (dato duro, no supuesto)

| Período | ASE1 | ASE2 | ASE3 | ASE4 | ASE5 |
|---|---|---|---|---|---|
| julio-1 | **2 Mes** / 2 Aplic | 3 / 0 | 3 / 2 | 3 / 2 | 3 / 0 |
| julio-2 | 3 / 2 | 3 / 2 | 3 / 2 | 3 / 2 | **2** / 2 |
| octubre-1 | 3 / 2 | 3 / 2 | 3 / **0** | 3 / 2 | **2** / 0 |
| agosto-2 | 3 / 2 | 3 / 2 | 3 / 0 | 3 / 2 | 3 / 2 |

Slots `Mes` que declara el mapa v2: ASE1..4 = 3, ASE5 = 2.
⇒ **Único desajuste estructural del set**: `ASE1 + julio-1` (mapa 3, fuente 2). Ver Δ-1.

---

## 6. T0.6 — Lista cerrada de Δ (S1.6)

### Δ-1 — **GAP BLOQUEANTE**: ASE1-julio-1 no es legible por la ruta única

- **Hecho:** julio-1 ASE1 trae **2** filas `Mes/Total` (24 y 40; no existe el bloque `AFaseo`) y la
  geometría v2 le declara `Mes2` → `F70`. `MapearR1Q2` trata `Mes2` como rol **core** y aborta:
  `CalculoInvalidoException — ASE 1: la fuente R1-Q2 no trae la fila del rol Mes2 requerida por el
  mapa T0 para la celda F70 (reporte Recaudoporcomponente).`
- **Impacto:** `CA4` ("`2026071` verde **por la misma ruta**") es irrealizable con la política actual;
  tras H1.1 decenas de tests del período julio-1 (default de `Insumos`) enrojecerían.
- **La aritmética NO es el problema** (evidencia en el fixture,
  `T0_6_Ase1Julio1_DosFilasMes_ReproduceSuR10ConMes2EnCero`): con el slot `Mes2` ausente tratado como
  `0`, la ruta única reproduce el R10 oficial de julio-1 ASE1 **al peso**:
  `D9 = 16.704.332.434,57` · `D28 = 54.216.385,68` · `D47 = 11.673.020,00` · `D66 = -12.054.255,65`
  ⇒ `D104 = 16.758.167.584,60` ⇒ `ROUND = 16.758.167.585` = `DetRetri2026071!D9`.
  Además el `TotalOportunoEsperado` v1 de esa fuente (`F25+F41−L25 = 16.704.332.434,57`) coincide con
  el agregado Σ ⇒ **cero cambio de valor**.
- **Propuesta de enmienda 2 (una línea de política):** `Mes2`/`Lmes2` pasan a **opcionales → `0m`**
  (mismo tratamiento que `Aplic0/Aplic1/LAplic0`), conservando `Mes0`/`Mes1` como core fail-fast y el
  **texto del mensaje byte-idéntico**. Consecuencias: (a) CA4 alcanzable; (b) cero cambio aritmético
  (los agregados son Σ, D6.3); (c) el caso rojo del plan (ASE3-octubre) sigue resolviendo igual;
  (d) `F70` se escribe `0` (el template ya está en `0`).
- **Sin esta ratificación, H1 no se abre** (implementar la política por cuenta propia sería inventar
  alcance: D1.1 congela el fail-fast de `Mes0..2`).

### Δ-2 — Abortos Q1 que desaparecen con la rama (`MapearR1`, los 8 E1-6)

| # | `case` | Aborto | Nueva causa en la ruta única |
|---|---|---|---|
| 1 | 1 | `filasMes.Count < 2` (F25/F41) | `Mes0/Mes1` core: sigue siendo fail-fast (mensaje propio) |
| 2 | 1 | sin fila `Aplicacion` (F10) | **Aplic opcional → 0** |
| 3 | 1 | sin fila `Subsidio(-)/Contribucion(+)` (F30) | la v2 **no lee el subsidio**; nada que exigir |
| 4 | 2, 3, 4, 5 | `filasMes.Count < 3` (TOT_OPT 5 términos) | Σ + `Mes2` opcional (Δ-1) |
| 5 | 3 | sin fila `Aplicacion` (F223) — **EL ROJO 2026101** | **Aplic opcional → 0** |
| 6 | 3 | sin fila Subsidio (F243) | la v2 no lee el subsidio |
| 7 | 4 | `filasAplicacion.Count < 2` (F401/F369) | **Aplic opcional → 0** |
| 8 | default | `ASE {id} no está soportado` | el mapa v2 cubre 1..5 |

### Δ-3 — Re-anclajes (sin tolerancia, sin `Skip`, sin borrados)

| Test | Cita | Δ esperado |
|---|---|---|
| `WorkbookLeafInputsTests` | `:16-27` | `TotalOportunoEsperado`/`ExtemporaneoEsperado` (getters v1) **no cambian** en julio-1 ASE1 (verificado: `16.704.332.434,57` / `11.673.020,00` = ruta única) |
| `ParidadCliTests` | `:135-137` | `F30/F10/L10` del leaf pasan a `0m` (D2): re-anclar contra el contrato legacy |
| `ValidadorBasicoTests` | `:77-79` | fixture sintético: sin cambio (ya son `0m`) |
| `AjustesSfTTests` | `:310` | fixture sintético: sin cambio |
| `VerificadorRecaudo` | `Program.cs:367` | herramienta: compila intacta |

### Δ-4 — Altas (esperadas en el set-diff de H2.3)

`ContratoUnicoLecturaR1T0Tests`: **58 casos** (T0.1×17, T0.2×5, T0.3×9, T0.4×11, T0.5×15, T0.6×1).
Son la única alta: no se borró ni se renombró ningún test existente.

---

## 7. Puerta H0→H1

| Criterio | Estado |
|---|---|
| Fixture T0 verde | ✅ 58/58 |
| S1.1 geometría v2 | ✅ |
| S1.2 censo octubre-1 | ✅ |
| S1.3 CA0/CA2/CA3/CA7 | ✅ (5/5 contra R10 oficial) |
| S1.4 caracterización R2/R4 | ✅ (8/10 igualdad + 2 abortos declarados; destinos literales) |
| S1.5 inventario de consumidores | ✅ |
| S1.6 lista cerrada de Δ | ✅ **con Δ-1 = GAP bloqueante** |
| Build 0 warnings | ✅ |
| G-SHA antes/después | ✅ |

**Veredicto: PUERTA CERRADA.** T0.4 no dispara STOP (R2/R4 son idénticos), pero **Δ-1 sí**: la ruta
única no puede leer julio-1 ASE1 y CA4 exige que lo haga. Se escala al Ingeniero con la enmienda 2
propuesta (`Mes2`/`Lmes2` opcionales → `0m`). Sin ratificación no se abre H1.

---

## 8. Evidencia de regresión (H0, sin cambios de producción)

`dotnet test Remuneracion.IntegrationTests` → **1346 Total / 1279 Passed / 67 Failed / 0 Skipped**
(`TestResults/p54-H0-suite.trx`).

Set-diff contra la línea base ratificada `p53-h2-suite.trx` (1289/1222/67/0):

- Altas: **+57** (todas verdes; el caso 58 es posterior a esa corrida).
- Bajas: **0**.
- **Rojos nuevos: 0.** **Rojos desaparecidos: 0.** (67 = 67, mismos nombres).
