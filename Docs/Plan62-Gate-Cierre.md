# Plan 62 — Gate de cierre (G.1…G.5)

**Fecha de ejecución:** 2026-10-29 · **Verificador:** subagente `implementer` (Gate en sesión aparte) ·
**Plan fuente:** `plans/62 - pegado-r10-det-detvali-y-validacion-recaudo-q1.md` §3 (Gate) + §4 (CA)
**Estado del working tree al medir:** H0…H5 cerradas; `dotnet build` de la solución **0 warnings / 0 errors**.
**Doctrina de trabajo:** nunca se escribió sobre `Docs/Insumos`, `Docs/**/Resultados`, el manual del
administrativo ni la plantilla oficial: **todo corrió sobre COPIAS** en `temp/p62-gate/` (borradas en G.5).
**Sin commits / push / checkout / reset.**

> **Snapshot del árbol.** El Gate arrancó con el árbol en el estado de H5 y, a media ejecución, otra
> sesión del implementer corrigió un rojo que H5 había introducido (`DivergenciasDeclaradasComparador.cs`
> escribía códigos de período en literales de producción y encendía `CandadoPeriodoQuemadoPlan59Tests`;
> corregido a las 10:39:23, solo texto de `Motivo`/`Cita`, semántica del catálogo intacta). **La medición
> final de la suite (10:58→11:10) y del comparador es POSTERIOR a esa corrección** y es la que se reporta.
> Las tres corridas de producción se repitieron después con los binarios finales y resultaron
> **semánticamente idénticas** (el comparador entre la salida v1 y v2 da 0 divergencias en los 3 períodos)
> — ver §6.2.

---

## 1. G.1 — Corrida de producción 2026091 (5 ASE)

### 1.1 Insumos, bloqueo de Excel/OneDrive y copias

| Paso | Resultado |
|---|---|
| `Get-Process EXCEL` | **sin procesos** (Excel cerrado) |
| Copia `Docs/Prueba Septiembre 1/Insumos` → `temp/p62-gate/insumos-2026091` | **OK** (36 archivos, 0 bloqueos) |
| Copia plantilla `Docs/Plantilla estandarizada/v4/… v4.xlsx` → `temp/p62-gate/plantilla/v4.xlsx` | **OK** |
| Salida | `temp/p62-gate/out2-2026091/` (carpeta temporal del repo, jamás `Resultados`/`Insumos`) |

No hubo ningún archivo bloqueado por OneDrive/Excel; no fue necesario reintento ni se reporta stop.

### 1.2 Comando y código de salida

```
dotnet run --project Remuneracion.Cli/Remuneracion.Cli.csproj --no-build -- \
  --periodo 2026091 \
  --carpeta <repo>/temp/p62-gate/insumos-2026091 \
  --plantilla <repo>/temp/p62-gate/plantilla/v4.xlsx \
  --salida <repo>/temp/p62-gate/out2-2026091
```

| Hecho | Evidencia |
|---|---|
| Código de salida | **0** (`RESULTADO OK codigo=0`, `runId=f1e17c0e-1a6c-48f0-b256-0e98d42ae227`) |
| Archivo entregado | `Remuneración 202609-1 Total.xlsx` (622.117 B) — **sin** el sufijo ` (sin validar)` |
| Reporte de novedades | **no se emitió** (una corrida con novedades SIEMPRE emite TXT y sale 7 — Plan 61 H4.2) |
| Gates del período en stdout | `DetValiRetri D16:D20 cierra (máx\|dif\| = 0)` y `VALIDACION_TOTAL O9 = 0; P9 = TRUE` en los 5 ASE |
| Duración | 36,8 s |

**Conclusión (c):** la salida **NO** trae la novedad R10 `−1` — no hay reporte de novedades ni marca de
entrega "sin validar"; el gate de total calculado-vs-R10 ya no existe (D-61R10) y la hoja ES el R10.

### 1.3 Read-back DetRetri/DetValiRetri ≡ R10 (CA-6)

Auditor BCL independiente (Zip/OpenXML, sólo lectura) sobre la salida y el `R10_Remuneracion_2026091.xlsx`:

| Hoja | Rango | Celdas comparadas | Diferencias |
|---|---|---|---|
| `DetRetri2026091` | `A3:O19` (17×15) | 255 | **0** |
| `DetValiRetri2026091` | `A3:O14` (12×15) | 180 | **0** |

Regla: montos ±0,5 (el R10 se interpreta con el parser ÚNICO `ParseoMonetarioR10`), textos exactos,
vacías = vacías, forma-agnóstico en `B9:B13`. Incluye los encabezados que viajan pegados
(`PROCERASEO`, NIT, `M=2026091`, `D6=17/09/2026`, `D7=08:44 AM`, rango 01/09–15/09, pie de interventoría
`C16 = * Costos de Interventoría: REMUNERACIÓN 202609-1`).

### 1.4 Paridad contra el manual del administrativo (comparador binario)

`ComparadorSalidaVsManual(quincena:1).Comparar(salida, manual, incluirDivergenciasDeclaradas: true)`
contra `Docs/Prueba Septiembre 1/Resultados/Administrativo/Remuneracion 202609-1 Total.xlsx`.

| Superficie | Divergencias | Declaradas | **NO declaradas** |
|---|---|---|---|
| **Bloque pegado** `DetRetri` filas 3–19 | 3 | **3** | **0** |
| **Bloque pegado** `DetValiRetri` filas 3–14 | 0 | 0 | **0** |
| Bajo el pegado `DetRetri` filas 20+ | 51 | 0 | 51 |
| Bajo el pegado `DetValiRetri` filas 15+ | 62 | 0 | 62 |
| Workbook completo (fuera de exclusiones versionadas) | 17.976 | 3 | 17.973 |

**Declaradas aplicadas (catálogo `DivergenciasDeclaradasComparador`, las 3 con Id + cita + retiro):**

| Id | Celda | app (R10 pegado) | manual | Δ |
|---|---|---|---|---|
| `P62-DETRETRI-C16-PIE-INTERVENTORIA` | `DetRetri!C16` | `…REMUNERACIÓN 202609-1` | `…REMUNERACIÓN 202604-1` | pie heredado del manual |
| `P62-DETRETRI-D11-ERRATA-PESO-ASE3` | `DetRetri!D11` | 10.891.338.739 | 10.891.338.740 | **1 peso** |
| `P62-DETRETRI-D14-TOTAL-ERRATA-PESO` | `DetRetri!D14` | 63.145.073.622 | 63.145.073.623 | **1 peso** (arrastre) |

**Lectura honesta de las NO declaradas:** las 113 celdas de `DetRetri`(51)+`DetValiRetri`(62) caen
**todas bajo el bloque pegado** y son **texto de fórmula** (`D32:D36`, `K32:K36`, `M32:M36`, `D16:O21`,
etc.) que difieren por la **malla heredada del manual** (R-TRAMPA-ANCLA / clase H5-3): la app sigue la
geometría canónica de la plantilla y el manual trae sus propias anclas. **No** se declaran (declararlas
taparía una regresión de geometría) y **ninguna** cae dentro de la superficie pegada. Es exactamente la
misma clase que ya estaba en el baseline (julio/agosto) y no la introduce el Plan 62.

---

## 2. G.2 — Corridas 2026071 / 2026072 (Julio Q1/Q2) y goldens

### 2.1 Corridas (cero regresión)

Carpetas localizadas: `Docs/Prueba Julio-1/Insumos` (Q1) y `Docs/Prueba Julio-2/Insumos` (Q2); copiadas a
`temp/p62-gate/insumos-2026071` / `insumos-2026072` (66 y 46 archivos).

| Período | Comando (idéntico, cambiando periodo/carpeta) | Exit | Salida |
|---|---|---|---|
| 2026071 | `dotnet run --project Remuneracion.Cli -- --periodo 2026071 --carpeta … --plantilla v4.xlsx --salida …` | **0** (`runId=eac245e9…`) | `Remuneración 202607-1 Total.xlsx` (620.510 B) |
| 2026072 | idem `2026072` | **0** (`runId=a9b39a54…`) | `Remuneración 202607-2 Total.xlsx` (626.696 B) |

`RESULTADO OK codigo=0` en ambas; 27,2 s y 29,4 s. Read-back del pegado en julio:

| Período | `DetRetri` A3:O19 | `DetValiRetri` A3:O14 |
|---|---|---|
| 2026071 | 255 celdas · **0 dif** | 180 celdas · **0 dif** |
| 2026072 | 255 celdas · **0 dif** | 180 celdas · **0 dif** |

### 2.2 Set REAL de goldens (corrido, no inferido)

Del TRX `Remuneracion.IntegrationTests/TestResults/p62-gate-suite.trx` (filtro `testName ~ Golden`):

| Métrica | Baseline `p62-h4-suite.trx` | Gate `p62-gate-suite.trx` |
|---|---|---|
| Tests golden | 53 | **53** |
| Verdes | 46 | **46** |
| Rojos | 7 | **7** (mismo conjunto, 0 altas) |

Rojos golden (idénticos al baseline, todos en la lista de 25):

1. `GoldenBalanceScTests.ProcesadorPeriodo_CapaA2_4_BceContraGolden`
2. `GoldenConciliacionEmpresaTests.ProcesadorPeriodo_CapaA2_2_EmpresasYRecaudosContraGolden`
3. `GoldenInterventoriaTests.CapaA_InterventoriaYLEspeciales_AmbosCanonicos`
4. `GoldenLeafPathTests.Procesador_CapaA_LeafVisiblesFormulasYHashSinUsarCacheDeSalida`
5. `GoldenMultiAseTests.ProcesadorPeriodo_CapaAX5_LeafVisiblesFormulasHashYSumas`
6. `GoldenReporteBancoTests.ProcesadorPeriodo_CapaA2_3_BloquesYConsolidadoContraGolden`
7. `GoldenResumenMesYR10Tests.GoldenQ1_S3_ResumenMes_DE_EsCopia1a1DeLosRecaudos`

`GoldenDetRetriQ2Tests` (**6/6 verde**, incluye `CapaA_Q2_A1…A7`) y `DetRetriQ2Tests` (5/5) confirman que
el pegado no movió el detalle Q2 de julio.

### 2.3 Desplazamientos DECLARADOS de Julio (doctrina nueva: las hojas vienen del R10)

| # | Superficie | Antes | Ahora (Plan 62) | ¿Aserto actualizado? | Evidencia real |
|---|---|---|---|---|---|
| D-1 | `DetRetri{2026071/2026072}` `A3:O19` y `DetValiRetri{…}` `A3:O14` | asiento calculado por el dominio (Plan 59 H4) solo en Q2 | **bloque completo del R10 del período, en AMBAS quincenas** | Sí: `GoldenDetRetriQ2Tests` y `DetRetriCapacidadTests` adjuntan el bloque real (`LeerBloqueDetRetri/DetValiRetri`); `AsientoDetRetriEscritoPlan59Tests` y `DesgloseTrazableAmbasQuincenasPlan59Tests` **eliminados** (su objeto se retiró) | read-back 435 celdas × 2 períodos = **0 diferencias** |
| D-2 | `DetValiRetri!J9:J14` (AJUSTE A LA DECENA) | 0 de plantilla (SALE Plan 29) | **dato pegado del R10** | Sí: `JulioQ2_DetValiRetri_J_EsSale_QuedaEnCero` → `JulioQ2_DetValiRetri_J_TraeElDatoDelR10Pegado` | bloque `J` del R10 ≡ salida; `DetValiRetri` bloque vs manual = **0 divergencias** |
| D-3 | `DetRetri!J/L`, `DetValiRetri!I/O` | derivados del dominio | **pegados del R10**; en julio coinciden con el manual en los 5 ASE (la antigua divergencia ASE5 desaparece) | Sí: `DetRetriDesgloseTrazableTests` reescrito (J ≡ bloque ≡ manual, 5/5) | fixture **verde** |
| D-4 | Sello `DetRetri/DetValiRetri!D6/D7` | escritura separada (Plan 31) | **llega pegado**; cero doble escritura | Sí: `IntegridadCalcChainYSelloFechasTests.SelloProceso_D6Formula_LanzaErrPlantillaSinSobrescribir` → `…_NoSeToca_ElSelloDeRangoNoEscribeD6` | `D6/D7` de la salida ≡ R10 (read-back) |
| D-5 | Comparador: exclusión `DetValiRetri!J` | SALE versionada | **retirada**; `J` se compara como cualquier celda y las diferencias legítimas van por catálogo | Sí (H5/T5.1) | julio: **15.623 inesperadas de 18.861** — **el mismo conteo de inesperadas que el baseline** (15.623 de 18.856); el total se mueve +5 por la preparación de base, las inesperadas NO |
| D-6 | `E142:E146` (Q1→`D{21:25}`, Q2→`F{21:25}`) | Q1 leía `F` (bloque en 0) | columna por quincena | Sí (T4) | §3 |
| D-7 | `ComparadorSalidaVsManualTests(julio)` y `(agosto)` | rojo | **sigue rojo** (pre-existente: mesh del manual) | No aplica: se conserva el aserto anti-falso-verde | 15.623 / 7.254 inesperadas |

**No hay ningún golden enredado a la vieja geometría que se haya quedado en silencio**: los únicos
asertos tocados fueron los que el plan declaró (§7) y quedaron verdes; el resto conserva su rojo
pre-existente con el **mismo mensaje y el mismo conteo**.

---

## 3. G.3 — `CONSOLIDADO!E142:E146` por quincena + `fullCalcOnLoad`

Texto-`<f>` leído de las 3 salidas reales (nunca cachés `<v>`):

| Salida | Columna referida en `E142:E146` | Filas |
|---|---|---|
| 2026091 (Q1) | **`D`** | 21,22,23,24,25 (E142→21 … E146→25) |
| 2026071 (Q1) | **`D`** | 21…25 |
| 2026072 (Q2) | **`F`** | 21…25 |

Forma exacta (2026072): `E142 = 'Recaudo EAAB Reciprocidad'!F21+'Recaudo EAAB + Ciud Limp'!F21+'Recaudo Directa Occidente'!F21+'Recaudo ENEL'!F21+'Recaudo ENERBIT'!F21` (y análogas 143→22 … 146→25).

| Chequeo workbook | 2026091 | 2026071 | 2026072 |
|---|---|---|---|
| `calcPr fullCalcOnLoad` | **True** | **True** | **True** |
| `calcChain` part | **ausente (saneado)** | ausente | ausente |
| `<f>` dentro del pegado (`DetRetri` filas 3–19) | **0** | 0 | 0 |
| `<f>` en `DetValiRetri` filas 3–19 | 48 (= `D16:O19`, consumidoras, fuera del rango pegado `A3:O14`) | 48 | 48 |

"Abre sin reparación" se garantiza por el saneador (`SaneadorCadenaCalculo.Sanear` en los 3 caminos de
guardado) y su malla de tests: `IntegridadCalcChainYSelloFechasTests.SalidaJulioQ1/Q2_SinCalcChainYFullCalcOnLoad…`
y `RecaudoValidacionConsolidado142H4Plan62Tests.T4_FullCalcAlcanzable_SaneadorPasaTrasLaReescritura` —
**verdes** en el TRX del Gate.

---

## 4. G.4 — Suite completa FINAL y set-diff

Comando (con el directorio de logs **vacío antes de contar**):

```
# 1) limpieza: 0 archivos remuneracion_log_*.txt en Remuneracion.IntegrationTests/bin/Debug/net10.0
# 2) dotnet build Remuneracion.WinForms/Remuneracion.WinForms.slnx   → 0 Warning(s), 0 Error(s)
dotnet test Remuneracion.IntegrationTests/Remuneracion.IntegrationTests.csproj --no-build \
  --logger "trx;LogFileName=p62-gate-suite.trx"
```

| Métrica | Baseline `p62-h4-suite.trx` | **Gate `p62-gate-suite.trx`** |
|---|---|---|
| Total | 1.769 | **1.776** (+7 = `ComparadorPegadoH5Plan62Tests`, todos verdes) |
| Verdes | 1.744 | **1.753** |
| Rojos | 25 | **23** |
| Skip | 0 | **0** |
| Duración | — | 11 min 32 s (10:58:34 → 11:10:08) |

`Herramientas/set-diff-trx.ps1 p62-h4-suite.trx p62-gate-suite.trx`:

```
baseline : 1769 tests / 1744 verdes / 25 rojos
nuevo    : 1776 tests / 1753 verdes / 23 rojos
ALTAS (rojo nuevo) : 0
BAJAS (rojo cerrado): 2
  - Remuneracion.IntegrationTests.EspecialesR4H2Plan61Tests.H21_2026091_Ase2_R4EspecialesYDetRetri_ValoresMedidos
  - Remuneracion.IntegrationTests.EspecialesR4H2Plan61Tests.H21_2026091_Ase3_ElCasoDelPeso_R10EsErrataHumana
```

- **Altas: 0** (condición `rojo nuevo fuera de lista = STOP` cumplida).
- **Bajas: 2**, ambas de `EspecialesR4H2Plan61Tests` (H5 las reclasificó al dato vigente del R10:
  `t="s"` en vez de `t="str"` y ASE3 = 739 con la errata 740 solo en el manual). Quedan **verdes**.
- `ParidadCliTests` (3 casos) **verde** — el directorio de logs quedó limpio y no apareció ningún
  `_00N` (0 archivos con sufijo de secuencia), es decir **no hubo escritor concurrente** durante la ventana
  del Gate.
- **Confirmación explícita del punto que el brief pedía verificar:** `ComparadorSalidaVsManualTests`
  **julio** y **agosto** SÍ están en la lista de 25 del baseline (`p62-h4-suite.trx`, posiciones 9 y 19) y
  **siguen rojos** en el Gate (posiciones 9 y 10 del set final): no son un alta ni un cierre, son deuda
  pre-existente (mesh del manual). El enunciado que los excluía era incorrecto; el TRX manda.
- Build final: `dotnet build Remuneracion.WinForms/Remuneracion.WinForms.slnx` → **0 warnings / 0 errors**.

---

## 5. Veredicto por criterio de aceptación (con evidencia real)

| CA | Enunciado | Comando / evidencia | Veredicto |
|---|---|---|---|
| **CA-1** | septiembre-1 sin novedad R10 `−1`; `DetRetri/DetValiRetri` ≡ manual (±0,5/texto) salvo declaradas | corrida 2026091 → exit **0**, sin TXT de novedades, sin `(sin validar)`; read-back **435/435** celdas; comparador: en el bloque pegado **3 divergencias, las 3 declaradas, 0 no declaradas** | **CUMPLE** |
| **CA-2** | Julio Q1/Q2: sólo desplazamientos DECLARADOS, cero regresión no explicada | corridas 2026071/2026072 exit **0**; read-back **435 celdas × 2 = 0 dif**; goldens **53 tests: 46 verdes / 7 rojos = mismo set del baseline**; comparador julio **15.623 inesperadas (idéntico al baseline)**; tabla §2.3 | **CUMPLE** |
| **CA-3** | Q1 `E142:E146` = Σ`Recaudo *`!`D{21:25}`; Q2 = `F{21:25}`; `fullCalcOnLoad` alcanzable | texto-`<f>` §3: 2026091 y 2026071 → `D21:D25`; 2026072 → `F21:F25`; `fullCalcOnLoad=True` y `calcChain` ausente en los 3 | **CUMPLE** |
| **CA-4** | R10 ausente o sin las 2 hojas ⇒ fail-fast sin escritura parcial | `BloqueR10H1Plan62Tests.T1_3_R10Inexistente_FallaNombrandoLaRuta`, `T1_3_R10SinLaHojaDelPeriodo_FallaNombrandoArchivoYHoja`, `T1_3_R10SinUnMonto_FallaNombrandoLaCelda`, `PegadoR10H2Plan62Tests.T2_2_R10ConMontoEnBlanco_FallaSinEscrituraParcial` — **verdes** (0 rojos en los 118 tests Plan 62) | **CUMPLE** |
| **CA-5** | build 0 warnings; suite con espectro de reclasificación explícito | build **0/0**; suite 1.776 tests, **0 altas**, 2 bajas nombradas; espectro en `Docs/T5-Plan62-Reclasificacion-Rojos.md` (25/25 clasificados) | **CUMPLE** (con la deuda pre-existente de 23 rojos, declarada) |
| **CA-6** | read-back del pegado con anchors por hoja a la fuente R10 | auditor independiente: `DetRetri A3:O19` 255 celdas, `DetValiRetri A3:O14` 180 celdas → **0 diferencias en los 3 períodos**; más `PegadoR10H2Plan62Tests.T2_3_ReadBack_*` y `GatesPegadoH3Plan62Tests.T3_3_ReadBack_EnrojeceAntePegadoCorrupto` verdes | **CUMPLE** |

**Veredicto global del Gate: PASA.** Ningún CA queda en rojo; la deuda de 23 rojos es pre-existente
(Planes 39–49 + mesh del manual) y no creció.

---

## 6. Hallazgos, riesgos y notas de proceso (para el informe al Ingeniero)

### 6.1 Hallazgo H-GATE-1 (rojo nuevo introducido en H5 y corregido en el Gate)

H5 agregó `Remuneracion.Infrastructure/Excel/DivergenciasDeclaradasComparador.cs` con los períodos
reales dentro de los literales de `Motivo`/`Cita`. Eso encendió `CandadoPeriodoQuemadoPlan59Tests`
(los códigos de período en **código** están prohibidos; los comentarios sí están permitidos). La corrida
previa de la suite (10:11→10:30, `p62-h5-suite.trx`) mostró **1 alta** (`CandadoPeriodoQuemado…`):
lo correcto era **STOP**, no silencio. Otra sesión del implementer reescribió **solo el texto** de las
citas (movió los valores medidos a `Docs/T5-Plan62-Reclasificacion-Rojos.md` y `Docs/T0-Plan62-Pegado-R10.md`)
a las **10:39:23**; el candado pasó a verde y el catálogo conserva Ids, celdas y predicados. **La medición
del Gate es posterior** y el set-diff final da **0 altas**.

> Recomendación de proceso: que el candado se corra en cada fase (no solo al final) — un `[Fact]` de
> 40 ms hubiera evitado dos corridas completas de 20 min.

### 6.2 Determinismo del entregable

Dos corridas idénticas del mismo período producen archivos con **el mismo tamaño y contenido** pero
**SHA256 distinto** (metadatos de empaquetado/timestamps del ZIP, no celdas): el comparador entre la
salida v1 y v2 da **0 divergencias** en los 3 períodos. Consecuencia para el proceso: **no usar el SHA
del entregable como sello de reproducibilidad**; usar el comparador celda-a-celda.

### 6.3 Deuda que el Gate deja a la vista (no es del Plan 62)

- 23 rojos pre-existentes: 7 goldens de geometría heredada, 4 de la familia BCE, 4 de espejo-R1, 2 de
  base/metadatos de agosto, 1 de tabla J, 2 del comparador con manual, 3 de plantilla/P-N.
- Las 113 celdas de `DetRetri`/`DetValiRetri` **bajo** el bloque pegado difieren del manual por la malla
  heredada (R-TRAMPA-ANCLA / H5-3). No se declaran: hacerlo taparía una regresión de geometría real.

### 6.4 Riesgos observados en la ejecución

- Durante el Gate hubo **otras sesiones corriendo la suite en paralelo** (3 corridas completas entre
  10:11 y 10:52). La suite del Gate se lanzó con el entorno quieto, el directorio de logs vacío y sin
  `_00N` posteriores: la evidencia es limpia, pero el árbol cambió a las 10:39 (§6.1). Para futuras
  mediciones: **una sola corrida a la vez** por repositorio.

---

## 7. G.5 — Limpieza e integridad

| Acción | Resultado |
|---|---|
| `temp/p62-gate/` (copias de insumos, plantilla, salidas, harness BCL, reportes) | **borrado** |
| Logs `remuneracion_log_20261029*.txt` del bin de tests | borrados antes y después de la corrida (0 restantes) |
| `Docs/**/Insumos`, `Docs/**/Resultados`, manual del administrativo, plantilla oficial v1/v4 | **intactos** (SHA256 antes/después en §7.1) |
| Commits / checkout / reset | **ninguno** |
| Evidencia conservada | `Remuneracion.IntegrationTests/TestResults/p62-gate-suite.trx` + este documento |

### 7.1 Huellas (antes y después de todas las corridas — sin cambios)

| Rol | Archivo | SHA256 |
|---|---|---|
| Plantilla VIGENTE (pipeline) | `Docs/Plantilla estandarizada/v4/Formato Remuneracion 2026xx-x Total v4.xlsx` | `5FEDFD316980276945DF2AF29784EFFC5FE178C5CF1AA16F3C5D5F6F0A4E38E5` |
| Plantilla v1 (histórica, T0) | `Docs/Plantilla estandarizada/Formato Remuneracion 2026xx-x Total.xlsx` | `3086AC317595C2F8E84ED24B1C5EE3E52565DD5193394E3062878E7FA49C5214` |
| R10 del caso | `Docs/Prueba Septiembre 1/Insumos/R10_Remuneracion_2026091.xlsx` | `DD0EAA7B6C03180F1B360F3D83EDADD7B4CE18B921478C6AAC775C9F50E91122` |
| Manual (paridad) | `Docs/Prueba Septiembre 1/Resultados/Administrativo/Remuneracion 202609-1 Total.xlsx` | `3CE1B86967CD709520272224D60686875AD1AD2EE5C510A531BE09726E420916` |

Huella de **árbol** de los insumos (orden por ruta + SHA por archivo → SHA256 del conjunto):

| Árbol | Archivos | TREE-SHA256 (antes = después) |
|---|---|---|
| `Docs/Prueba Septiembre 1/Insumos` | 36 | `C60A5FF50663D6363D999303D114A41F0341E7F7C0E3DDA4BC2BCD060CD49DBA` |
| `Docs/Prueba Julio-1/Insumos` | 66 | `23D5A71C36A498818CFDEA9EAD523EC70196DB3864280661926DD5768D73239C` |
| `Docs/Prueba Julio-2/Insumos` | 46 | `9BC977A8E38879399E4E0062550EC5FD7D79B994AFE2E4E8973CE1F068E25459` |

R10 de julio (sin cambios durante el Gate): `2026071` = `DC01A1BC…1BEDA8`; `2026072` = `68360DAD…5125D3`.
Manual julio Q2: `584310105AC7223CE840833E3CB64E26F95B61907C9ECD959A0A632EEC21DCB1`.

### 7.2 Traza a los CA del plan

`CA-1 → §1` · `CA-2 → §2` · `CA-3 → §3` · `CA-4 → §5` · `CA-5 → §4` · `CA-6 → §1.3/§2.1` ·
`Gate §3 (G.1…G.4) → §1…§4` · `G.5 → §7`.
