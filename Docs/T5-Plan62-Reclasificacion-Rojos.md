# T5 — Reclasificación del espectro de rojos (Plan 62, H5)

> Fase: **H5 / T5.2**. Fuente del espectro: `Remuneracion.IntegrationTests/TestResults/p62-h4-suite.trx`
> (base de set-diff del Gate): **1.769 tests · 1.744 verdes · 25 rojos · 0 skip**.
> Regla aplicada (plan 62 §3 T5.2): para CADA rojo se determina raíz causal y acción, con tres
> desenlaces posibles: **(a) fixture actualizado** al dato vigente conservando la intención que
> blindaba; **(b) fragilidad log / deuda pre-existente**: NO se toca (queda de proceso, con dueño
> declarado); **(c) aserto anti-falso-verde que sigue válido**: se conserva.
> Cero corridas de producción: la evidencia sale del trx, del mensaje de fallo y de la lectura del
> fixture.

## 1. Matriz de preexistencia (evidencia objetiva, 5 baselines)

`R` = rojo en ese trx, `.` = verde. Columnas: `p59-h5-cierre` · `p61-h0-suite` · `p61-h3-suite` ·
`p62-h2-baseline` (baseline de entrada de H5) · `p62-h4-suite` (espectro vigente).

```
RRRRR  GoldenBalanceScTests.ProcesadorPeriodo_CapaA2_4_BceContraGolden
RRRRR  AdopcionPlantillaV2T0Plan48Tests.T0P48_P6_TablaParidad_40Hojas_ConteosReales
RRRRR  InterventoriaTests.D3b_LEspecialesMenores_TodaLNumericaFueraDeMapas_EsCero_EnAmbosCanonicos
RRRRR  BalanceScPermutaTests.JulioAse5Bce_ManualDivergeDeLaFuente_ConstanteDocumentada
RRRRR  GoldenReporteBancoTests.ProcesadorPeriodo_CapaA2_3_BloquesYConsolidadoContraGolden
RRRRR  GoldenInterventoriaTests.CapaA_InterventoriaYLEspeciales_AmbosCanonicos
RRRRR  WorkbookLeafInputsTests.GenerarWorkbook_EscribeSoloLeafYPreservaFormulas
RRRRR  ValidadorInteriorR1GateTests.S1_Julio2026072_FormaInteriorVerde
RRRRR  ComparadorSalidaVsManualTests.Salida5AseQ2_vs_Manual_CeroDivergenciasFueraDeBrechasDeclaradas(julio)
RRRRR  GoldenResumenMesYR10Tests.GoldenQ1_S3_ResumenMes_DE_EsCopia1a1DeLosRecaudos
RRRRR  BalanceScTests.J9Consolidado_NoEsOraculoDeValoresBce
RRRRR  BaseConsistente2026082Tests.BaseAgosto2026082_MetadatosCanonicosEsperados
...RR  EspecialesR4H2Plan61Tests.H21_2026091_Ase2_R4EspecialesYDetRetri_ValoresMedidos
RRRRR  BasePeriodoPreflightTests.S5_OficialDirecto_EndToEnd_L9L13PorQuincena(2026071, Julio-1, N)
RRRRR  BaseConsistente2026082Tests.BaseAgosto2026082_EsInternamenteConsistenteConElManual
RRRRR  Regresion2026082Tests.Ejecutar_Periodo2026082_Modo5Ase_EndToEnd_DetRetri5De5VsR10
RRRRR  AdopcionPlantillaV2T0Plan48Tests.T0P48_P3_R1_Manual_492_Formulas_DosPatrones
RRRRR  GoldenMultiAseTests.ProcesadorPeriodo_CapaAX5_LeafVisiblesFormulasHashYSumas
RRRRR  ComparadorSalidaVsManualTests.Salida5AseQ2_vs_Manual_CeroDivergenciasFueraDeBrechasDeclaradas(agosto)
RRRRR  SharedFormulaR1GateTests.S1_RegenFrescaJulio_MaestrosSharedPreservados
RRRRR  GoldenConciliacionEmpresaTests.ProcesadorPeriodo_CapaA2_2_EmpresasYRecaudosContraGolden
RRRRR  BalanceScPermutaTests.JulioQ2_PermutaBce_SubsidioAD_ContribucionAE_VsFuente_E2E
RRRRR  ValidadorInteriorR1GateTests.S1_Agosto2026082_FormaInteriorVerde
...RR  EspecialesR4H2Plan61Tests.H21_2026091_Ase3_ElCasoDelPeso_R10EsErrataHumana
RRRRR  GoldenLeafPathTests.Procesador_CapaA_LeafVisiblesFormulasYHashSinUsarCacheDeSalida
```

**Lectura del set-diff:** **23 de 25** ya estaban rojos antes del Plan 62 (rojos desde el cierre del
Plan 59 ⇒ deuda acumulada, no regresión de N1/N2). **2** rojos son **data drift** del R10 reemitido
(aparecen en `p62-h2-baseline`, no en `p61-h3-suite`). Un rojo del Plan 59
(`DetRetriDesgloseTrazableTests.JulioQ2_DesgloseTrazable_DetRetri_JyL_Y_DetVali_IyO`) se reclasificó y
quedó **verde** en H2 (el pegado cubre el desglose: D-59H4).

## 2. Veredicto por rojo (25/25)

| # | Test | Mensaje del baseline (resumen) | Raíz causal | Veredicto | Acción H5 |
|---|---|---|---|---|---|
| 1 | `GoldenBalanceScTests…BceContraGolden` | Δ = −3.256.235.169,97 (= −D del golden ASE1) | El dominio lee `Contribucion = 0` para ASE1-julio-1: la fuente BCE del período no entrega esa columna en la geometría que el golden espera (familia **F-BCE-0**, lectura vs golden v1/v2) | (b) pre-existente | NO SE TOCA |
| 2 | `AdopcionPlantillaV2T0Plan48Tests…P6_TablaParidad_40Hojas` | `Reporte Componentes R1: fDiff esperado 0, observado 111` | Paridad v1↔v2 de 40 hojas: las 492 fórmulas del manual que la v2 no trae (deuda declarada Plan 48/49, `DEUDA-DECLARADA`) | (b) pre-existente | NO SE TOCA |
| 3 | `InterventoriaTests.D3b_LEspecialesMenores…` | `Reporte Componentes R1!L132` = 53.454.221,28 (debe ser 0) | L-Especiales fuera del set mapeado por rol/ocurrencia en la geometría v4 (familia R1-espejo; deuda D3b) | (b) pre-existente | NO SE TOCA |
| 4 | `BalanceScPermutaTests.JulioAse5Bce…` | Δ = 2.477.195.172,11 (= E del golden ASE5) | Misma familia **F-BCE-0** (ASE5-julio-Q1) | (b) pre-existente | NO SE TOCA |
| 5 | `GoldenReporteBancoTests…BloquesYConsolidado` | Δ = 16.194.517.449 | Golden por celda contra la salida: el golden conserva caché/geometría de otra versión de plantilla (familia **F-GOLDEN-GEOM**) | (b) pre-existente | NO SE TOCA |
| 6 | `GoldenInterventoriaTests…InterventoriaYLEspeciales` | Δ = 31.737.181,81 | Ídem anterior (interventoría/especiales: `INTERVENTORIA` tabla J + L-Especiales) | (b) pre-existente | NO SE TOCA |
| 7 | `WorkbookLeafInputsTests.GenerarWorkbook…` | `Assert.True` sin mensaje (uno de: `F46`/`CONSOLIDADO!D9`/`D109` debe seguir siendo fórmula) | El espejo R1 dimensiona filas y la celda ancla deja de ser fórmula (familia **F-ESPEJO-R1**, E1-2) | (b) pre-existente | NO SE TOCA |
| 8 | `ValidadorInteriorR1GateTests.S1_Julio2026072` | `Gate R1: la forma interior … ASE 1 F54 [SUBS]: esperado [F48], real sin fórmula` | El interior R1 conserva el anclaje de julio (anclas `F45..F55` desplazadas); follow-up **F-T4-1** | (b) pre-existente | NO SE TOCA |
| 9 | `ComparadorSalidaVsManualTests(… "julio")` | 15.623 divergencias INESPERADAS de 18.856; primeras en `AJUSTES - SF-T!AA10 app=[=f('SALDOS POR NOTA'!O23)] manual=[=f(…!O22)]` | **Mesh del manual heredado** (R-TRAMPA-ANCLA): la app sigue la geometría canónica y el manual desplaza filas en `SALDOS POR NOTA`/`AJUSTES - SF-T`; julio no tenía esa clase declarada (solo agosto, T4) | (b) pre-existente + **reclasificado en H5**: la exclusión `DetValiRetri!J` **SALIÓ** del catálogo (D-J9); el rojo NO viene de J | Doc del fixture actualizado; el aserto se conserva (anti-falso-verde) |
| 10 | `GoldenResumenMesYR10Tests.GoldenQ1_S3` | `Assert.Equal` esperado 0, observado 29.345.550 | El lector entrega un registro que la base canónica Q1 (copiada del manual) tiene en 0: **base vs fuente** (deuda Plan 30) | (b) pre-existente | NO SE TOCA |
| 11 | `BalanceScTests.J9Consolidado_NoEsOraculoDeValoresBce` | `NotEqual` 0 vs 0 | Familia **F-BCE-0** (la celda BCE usada como contraste es 0) | (b) pre-existente | NO SE TOCA |
| 12 | `BaseConsistente2026082Tests…MetadatosCanonicos` | `"46267"` vs `"46297"` (serial 02/09 vs 02/10) | Metadato de período de la base de agosto ≠ manual (fecha del bloque) — familia **F-METADATOS-BASE** | (b) pre-existente | NO SE TOCA |
| 13 | `EspecialesR4H2Plan61Tests.H21_2026091_Ase2…` | `El R10 del período es una copia formateada (t="str")` | **DATA DRIFT**: el R10 reemitido trae los montos como texto monetario en `sharedStrings` (`t="s"`), no como `t="str"`; el helper además leía el `<v>` crudo (= índice de la cadena) | **(a) fixture actualizado** | `EsTexto` = celda de texto (`s`/`str`/`inlineStr`); `Monto()` parsea el texto RESUELTO. Intención intacta: el R10 es copia formateada, no oráculo independiente. **VERDE** |
| 14 | `BasePeriodoPreflightTests.S5_OficialDirecto…(2026071, N)` | `ERR-PLANTILLA: … F682 de 'Reporte Componentes R1' … la plantilla trae un valor fijo` | Familia E1-6 / `R-ANCLA-EXTEMP`: el contrato de sección Q1 vs slots del template (deuda Planes 41/46, `R-SLOT-Q1-ENERBIT`) | (b) pre-existente | NO SE TOCA |
| 15 | `BaseConsistente2026082Tests…EsInternamenteConsistente` | `"46267"` vs `"46297"` | Ídem #12 | (b) pre-existente | NO SE TOCA |
| 16 | `Regresion2026082Tests.Ejecutar…DetRetri5De5VsR10` | `Assert.Equal` esperado 42, observado 81 (filas del bloque R1 del ASE1) | El espejo R1 dimensiona 81 filas donde el contrato de agosto espera 42 (familia **F-ESPEJO-R1**, E1-2) | (b) pre-existente | NO SE TOCA |
| 17 | `AdopcionPlantillaV2T0Plan48Tests…P3_R1_Manual_492_Formulas` | `Assert.Equal` esperado 492, observado 415 | Las 492 fórmulas del manual R1 no adoptadas (deuda declarada Plan 48/49, P5a pendiente) | (b) pre-existente | NO SE TOCA |
| 18 | `GoldenMultiAseTests…LeafVisiblesFormulasHashYSumas` | Δ = 932.978,37 vs golden | Familia **F-GOLDEN-GEOM** (suma/hoja golden de otra geometría) | (b) pre-existente | NO SE TOCA |
| 19 | `ComparadorSalidaVsManualTests(… "agosto")` | 7.254 divergencias de 18.906; `Informe AFaseo Recaudo 202608-2!D12 app=[…F119…] manual=[…F80…]` | **Mesh del manual de agosto** (T4/R-D-3..R-D-6 cubren `SALDOS POR NOTA`/`AJUSTES`/`REMUNERACION_*`, no los `Informe AFaseo Recaudo`): misma clase, frente distinto | (b) pre-existente + reclasificado en H5 (J fuera del catálogo) | Doc del fixture actualizado; aserto conservado |
| 20 | `SharedFormulaR1GateTests.S1_RegenFrescaJulio` | `T3 (tabla J): INTERVENTORIA!R26 trae la forma inesperada 'F15' (se esperaba 'N26' o 'J15')` | Tabla J de `INTERVENTORIA` sin variante Q1: requiere **T0 propio** (declarado en project-context, Plan 40) | (b) pre-existente | NO SE TOCA |
| 21 | `GoldenConciliacionEmpresaTests…EmpresasYRecaudos` | Δ = 19.104.244.534,3 | Familia **F-GOLDEN-GEOM** | (b) pre-existente | NO SE TOCA |
| 22 | `BalanceScPermutaTests.JulioQ2_PermutaBce…E2E` | Δ = 1.210.425.628,31 | Familia **F-BCE-0** (julio ASE5) | (b) pre-existente | NO SE TOCA |
| 23 | `ValidadorInteriorR1GateTests.S1_Agosto2026082` | `Gate interior R1 de agosto en ROJO (intencional en T1 …)` + `F51/F52/F53 sin fórmula` | **Anclaje julio** del interior R1 en la malla de agosto (F-T4-1 declarado) | (b) pre-existente | NO SE TOCA |
| 24 | `EspecialesR4H2Plan61Tests.H21_2026091_Ase3_ElCasoDelPeso…` | `El R10 del período es una copia formateada (t="str")…` | **DATA DRIFT**: el R10 reemitido ya trae el **739** correcto (la errata **740** quedó solo en el manual del administrativo) | **(a) fixture actualizado** (renombrado a `…_ErrataHumanaCorregidaEnElR10`) | Assert: R10 = 739 ≡ `Detalle` (Δ = 0); **manual = 740** (Δ manual−app = 1, declarado como `P62-DETRETRI-D11-ERRATA-PESO-ASE3` y `P62-DETRETRI-D14-TOTAL-ERRATA-PESO`). **VERDE** |
| 25 | `GoldenLeafPathTests.Procesador_CapaA…` | Δ = 19.549.786.950,62 | Familia **F-GOLDEN-GEOM** (leaf vs golden de otra versión de plantilla) | (b) pre-existente | NO SE TOCA |

**Balance del espectro:** 2 rojos cerrados por **fixture actualizado** (data drift del R10), 23 rojos
**pre-existentes** conservados como deuda declarada (9 de ellos son goldens/geometría heredada, 4 de
la familia BCE, 4 de la familia espejo-R1, 3 de metadatos/base de agosto, 1 de tabla J, 2 del
comparador con manual). **Cero altas nuevas** por la implementación de H5.

**Clase (b) "fragilidad log" (ParidadCli/MicroFixes, dependencia de limpieza de
`remuneracion_log_*.txt`):** **ninguno de los 25** rojos pertenece a esa clase — no se tocó nada por
ese motivo y el gotcha sigue declarado como de proceso.

## 3. Asertos anti-falso-verde del Plan 61: qué se conservó

- `EntregaConNovedadesH3Plan61Tests` (**H34/H42**) — **verdes** en el baseline: no se tocan (la
  novedad de TOTAL-R10 dejó de emitirse en H3; la política de entrega con novedades sigue intacta).
- `GatesPegadoH3Plan62Tests` — verdes; el gate total ausente y el read-back de integridad siguen
  siendo la malla de la hoja.
- `ComparadorSalidaVsManualTests` — se conserva el aserto "cero divergencias fuera de exclusiones +
  brechas declaradas" (es el anti-falso-verde del comparador). Lo único que cambió es la doctrina
  citada en su encabezado y el catálogo del comparador: **`DetValiRetri!J` ya no se excluye** y las
  diferencias legítimas se declaran por celda + par de valores (nunca por columna).

## 4. Hallazgos y STOP

- **H5-1 (data drift, no defecto):** el R10 2026091 vigente cambió dos veces de forma frente a la
  narrativa del Plan 61: (i) montos como **texto compartido** (`t="s"`) en vez de `t="str"`;
  (ii) **ASE3 = 739** (errata 740 corregida en el dato). Los dos fixtures se actualizaron sin perder
  lo que blindaban.
- **H5-2 (divergencia declarada nueva, medida):** además del pie de interventoría (`C16`), el
  comparador de septiembre-1 arroja el **peso de la errata** en `DETRETRI!D11` y `D14`. Los tres
  quedan en `DivergenciasDeclaradasComparador` con Id y condición de retiro.
- **H5-3 (alcance honesto):** en las hojas `DetRetri`/`DetValiRetri` las fórmulas de las filas 20+
  difieren del manual por **geometría heredada** (mesh R1/R2/R4 por anclas distintas). NO se
  declaran: declararlas taparía una regresión de geometría. Pertenecen a la clase R-TRAMPA-ANCLA y
  siguen en rojo en los fixtures del comparador con manual (#9 y #19).
- **H5-4 (falso-verde cerrado en el fixture del comparador):** la brecha "Unidad D (T4/T0d)" de
  `ComparadorSalidaVsManualTests` declaraba como "no escrita" TODA la superficie de las dos hojas
  (D/J/L de `DetRetri`, D..N de `DetValiRetri`). Con el pegado esa premisa es falsa y habría tapado
  cualquier diferencia del bloque copiado (incluida `DetValiRetri!J`, la que D-J9 quiere comparar):
  se acotó al **bloque de validación bajo el pegado** (`DetRetri` fila ≥ 20, `DetValiRetri` fila ≥ 15).
  Medición: julio/agosto conservan EXACTAMENTE los mismos inesperados (15.623 y 7.254) — la superficie
  pegada de esos períodos coincide con su manual salvo el pie de interventoría (declarado, p. ej.
  julio: app `REMUNERACIÓN 2026072` vs manual `remuneración 2026031`).
- **Sin STOP:** ningún hallazgo contradice el plan. Los 23 rojos pre-existentes son deuda de los
  planes 39–49 y del mesh del manual; el Gate los mantiene en el set-diff (0 altas).

## 5. Evidencia de cierre de H5 (re-corrida fresca 2026-10-29)

| Verificación | Comando / artefacto | Resultado |
|---|---|---|
| Build | `dotnet build Remuneracion.WinForms/Remuneracion.WinForms.slnx` | **0 warnings / 0 errors** |
| Fixtures de fase (H1–H5 + tocados) | `p62-h5-fases.trx` (filtro por `PegadoR10*`, `BloqueR10*`, `GatesPegado*`, `RecaudoValidacion…142*`, `ComparadorPegadoH5*`, `EspecialesR4H2*`, `DetRetri*`, `GoldenDetRetriQ2*`, `EntregaConNovedadesH3*`, `FNew1*`, `FailFastHojaDetRetri*`, `IntegridadCalcChain*`, `DivergenciaDeclarada*`) | **154 / 154 verdes** |
| Suite completa | `p62-h5-suite.trx` | **1.776 tests · 1.753 verdes · 23 rojos · 0 skip** |
| Set-diff vs `p62-h4-suite.trx` | `Herramientas/set-diff-trx.ps1` | **0 altas · 2 bajas** (`H21_2026091_Ase2…` y `H21_2026091_Ase3…`, cerrados por fixture actualizado al dato vigente) |
| Candado anti-quemado | `CandadoPeriodoQuemadoPlan59Tests` (6/6) | **verde** — detectó literales de período en la PRIMERA versión del catálogo nuevo; se corrigieron (la cita apunta al doc, no quema el período). Sus pruebas de injerto siguen enrojeciendo |
| Motivos de los 23 rojos | comparación mensaje-a-mensaje baseline vs H5 | **21 idénticos**; los 2 del comparador conservan el MISMO número de inesperadas (15.623 julio / 7.254 agosto) y solo cambia el TOTAL comparado (18.856→18.861 y 18.906→18.911: las 5 celdas de `DetValiRetri!J` que dejaron de excluirse, dentro de la brecha declarada del bloque de validación) |
| Insumos (huellas) | SHA256 | plantilla v1 `3086AC31…9C5214` · R10 `DD0EAA7B…0E91122` · manual `3CE1B869…E420916` · plantilla v4 (la que consume el pipeline) `5FEDFD31…4E38E5` — **intactas** |

**Gotcha operativo (no del producto):** correr dos suites en paralelo (o borrar
`remuneracion_log_*.txt` **durante** una corrida) enrojece la clase de fragilidad log
(`ParidadCliTests`, `MicroFixesHu17Tests.S2`, `CandadoPeriodoQuemadoPlan59Tests.H5_Produccion`). Es el
gotcha ya declarado: la suite se corre **sola** y el log se limpia **antes** de arrancar.
