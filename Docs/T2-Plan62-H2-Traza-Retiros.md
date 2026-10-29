# Plan 62 — H2: traza de retiros, cota de orden y reclasificación

> Evidencia de implementación de la fase **H2** del `plans/62 - pegado-r10-det-detvali-y-validacion-recaudo-q1.md`
> (T2.1…T2.5 + refinamientos §12: N-1, N-5, R5-GATE, H3-SEMANTIC). Los números de línea son los del
> árbol en el momento de la entrega de H2.

---

## 1. Cota de orden del pipeline (T2.1)

**Decisión definitiva (verificada en código, no en memoria):** el pegado ocurre **dentro de la pasada
atómica del writer**, después del cálculo del CONSOLIDADO y de la Phase 0, y **antes** de los gates de
datos y del oráculo de validaciones.

| Paso | Archivo:línea | Qué prueba |
|---|---|---|
| 1. Cálculo CONSOLIDADO del período | `Remuneracion.Core/Services/ProcesadorPeriodo.cs:316-317` | `D104:D108`/`U104:AG108` ya están calculados |
| 2. Lectura del R10 + de los DOS bloques del contrato | `ProcesadorPeriodo.cs:339-340` | fail-fast **antes** de escribir (R10 ausente / sin las 2 hojas / bloque incompleto) |
| 3. Escritura del workbook (una sola pasada) | `ProcesadorPeriodo.cs:402` | `_workbookLeafWriter.GenerarWorkbook(...)` |
| 3.a Phase 0 de la base (renombres/`L9:L13`) | `Remuneracion.Infrastructure/Excel/OpenXmlPlantillaWriter.cs:290` | el pegado corre sobre la base ya preparada |
| 3.b escritura leaf/empresa/banco/BCE/ajustes | `OpenXmlPlantillaWriter.cs:329` | CONSOLIDADO escrito antes del pegado |
| 3.c **PEGADO** (`PegarBloqueR10`) | `OpenXmlPlantillaWriter.cs:347` (multi-ASE) y `:165` (single-ASE) | values-only sobre `DetRetri`/`DetValiRetri` |
| 3.d sellado de cachés + saneamiento | `OpenXmlPlantillaWriter.cs:361` y `:364` | `fullCalcOnLoad` intacto (no se toca `<f>`) |
| 3.e **read-back de integridad del pegado** | `OpenXmlPlantillaWriter.cs:383` (y `:199`) | `VerificarIntegridadPegadoR10` sobre los bytes guardados |
| 4. Gate DetRetri-vs-R10 (per-ASE; total → H3) | `ProcesadorPeriodo.cs:455` | lee el R10 y el dominio, ya con la hoja pegada |
| 5. Oráculo de validaciones cruzadas (lee la SALIDA) | `ProcesadorPeriodo.cs:480` | evalúa `DetValiRetri D16:D20` desde lo pegado |

**Consecuencia:** `DetValiRetri!D16:O20` (fórmulas que consumen `D9:D13` + `CONSOLIDADO!U104:AG104`)
evalúan lo pegado y lo calculado en el mismo punto; ningún gate ve el estado intermedio.

---

## 2. Retiros con traza de consumidores (T2.5)

| Retiro | Dónde vivía | Consumidores verificados (grep) | Dónde vive hoy el valor |
|---|---|---|---|
| `EscribirCeldasDetRetri` (asiento `D9:D14`, AMBAS quincenas, Plan 59 H4) | `OpenXmlPlantillaWriter` (método retirado) | `VerificarAsientoDetRetri` (retirado con él); fixtures `AsientoDetRetriEscritoPlan59Tests` (reclasificado), `DetRetriCapacidadTests`/`FailFastHojaDetRetriPeriodoTests`/`GoldenDetRetriQ2Tests` (ahora aportan el bloque del R10) | `leaf.DetRetriQ2.Detalle` sigue en el DOMINIO y lo consume el gate `ValidarDetRetriContraR10` (`ProcesadorPeriodo.cs:455`); la hoja la llena el R10 pegado |
| `VerificarAsientoDetRetri` (gate de valores del asiento) | idem | `EntregaConNovedadesH3Plan61Tests` (reclasificado a `VerificarIntegridadPegadoR10` para fijar la firma de `ConservaParcialAlFallar`) | Sustituido por `VerificarIntegridadPegadoR10` (T2.3) |
| `EscribirDesgloseTrazableDetRetri` + `EscribirCeldaDeDesglose` + `DesgloseDeclaradoPorElDominio` + `ColumnasDelDesglose` + `ValorDeclarado` + `BalanceSubsidioContribuciones` | idem (rótulos `Costo de la Interventoría` / `Validación retribución` / `Balance de subsidio o contribuciones`) | solo se consumían entre sí; fixtures `DesgloseTrazableAmbasQuincenasPlan59Tests` (retirado, su aserto pasa al read-back) y `DetRetriDesgloseTrazableTests` (reclasificado en J) | `leaf.BalanceSc` y `InterventoriaDeclarada.PorQuincena` siguen en el DOMINIO (los loguea `ProcesadorPeriodo.cs:425-429`); la hoja la llena el R10 pegado |
| `VerificarDesgloseTrazableDetRetri` + `VerificarCeldaDeDesglose` + `LeerValorDelAsiento` | idem | únicamente los métodos retirados | sustituido por el read-back del pegado |
| `EscribirValorNumericoEnCeldaExistente` (guarda de capacidad `D9:D14`) | idem | única consumidora: el desglose/asiento retirado | la guarda de capacidad vive ahora en el plan-then-write del pegado (`PlanificarBloqueR10`: celda destino ausente ⇒ `ERR-PLANTILLA` nombrando hoja+celda) |
| `SellarProcesoEnDetRetri` + `FormatearHoraProceso` + `EscribirValorTexto` (sello `D6/D7`, Plan 31) | idem; llamado desde `EscribirFechasPeriodo` | ninguno más (grep de los 3 símbolos); la UI no parsea el sello | `D6/D7` viajan **pegados** en el bloque del R10 (T2.4). La **lectura** `FechaProceso`/`HoraProceso` y su fail-fast `ERR-FORMATO-FUENTE` siguen intactas en `ProcesadorPeriodo.cs:372-390` |
| Comentario «las 23 celdas» de la v2 (`WorkbookLeafCellMapQ2.DetRetriProtected`) | `OpenXmlPlantillaWriter.cs` (bloque de validación de fórmulas protegidas) | el mapa es LEGACY: solo lo usan fixtures históricos (`GoldenDetRetriQ2Tests`, `MapaQ2T0Tests`) | la verificación de forma del detalle la gobierna `DetallePeriodoDerivado.Validar`; el comentario ahora declara que el pegado **no toca** esas 23 celdas (R5-GATE) |

**Grep de cierre:** `SellarProcesoEnDetRetri|EscribirCeldasDetRetri|VerificarAsientoDetRetri|EscribirDesgloseTrazableDetRetri|VerificarDesgloseTrazableDetRetri|EscribirValorNumericoEnCeldaExistente|EscribirValorTexto|FormatearHoraProceso` → **0 referencias** fuera de los comentarios de retiro.

**Formatos de log que consume la UI (verificación explícita, el retiro NO los toca):** los únicos
formatos que la UI parsea son las líneas de `ResultadoProcesoPeriodo.Validaciones`
(`ASE n VALIDACIONES:`, `  {Empresa}: O (…) = …; P (INT(O)=0) = …`, `  VALIDACION_TOTAL O9 = …; P9 = …`,
`  DetValiRetri D16:D20 cierra …`) que siguen emitiéndose en `ProcesadorPeriodo.cs:550-582`; `Form1.cs:936`
las re-emite a Serilog con `ForContext("Validacion", "DetValiRetri" | "VALIDACIONES")` y
`Form1.cs:922-931` reconoce `VALIDACION_TOTAL` (`%Validacion%`). Ninguno de los métodos retirados
emitía log propio (`SellarProcesoEnDetRetri`/`EscribirCeldasDetRetri`/`EscribirDesgloseTrazableDetRetri`
escribían celdas, no log), así que no desapareció ningún formato consumido por la UI ni por el
`FormDetalleLog`.

**Conjunto físico muerto:** los dos fixtures cuyo OBJETO era la escritura retirada se eliminaron y su
aserto se trasladó al read-back (`AsientoDetRetriEscritoPlan59Tests`,
`DesgloseTrazableAmbasQuincenasPlan59Tests` → `PegadoR10H2Plan62Tests.T2_2_*` y `T2_3_*`).

---

## 3. Pegado: contrato implementado (T2.2/T2.3)

- Punto único: `OpenXmlPlantillaWriter.PegarBloqueR10(workbookPart, periodo, bloqueDet, bloqueVali)`
  (`internal`, costura de prueba) — hojas por `NombresHojaPeriodo`, rango del contrato `A3:O19` /
  `A3:O14`, plan-then-write (valida todo y recién entonces escribe).
- **Guard anti-fórmula**: `GuardEscrituraFormula.ExigirSinFormula` por celda ⇒ `ERR-PLANTILLA`
  nombrando hoja + celda (misma familia única del writer, sin duplicar el predicado).
- **N-5 (sin doble estándar)**: el `Numerico` del bloque se escribe exacto
  (`39148067546.64` queda `39148067546.64`). El redondeo a entero sigue siendo regla de CODA del
  CONSOLIDADO (`DetRetriRounder`), no del pegado.
- **R5-GATE**: no se toca ninguna `<f>` (ni `t`/`ref`/`si`), cero `InsertRow`/`DeleteRow`; el test
  barre el conjunto completo de fórmulas de las dos hojas antes/después y exige igualdad exacta.
- **Vacías**: la celda del bloque vacía deja el destino VACÍO (se limpia `<v>`/inline; nunca `0`).
- **Textos**: texto exacto crudo como inline string; estilo `s` del destino intacto.
- **Read-back** (`VerificarIntegridadPegadoR10`): montos ±0,5, textos exactos, vacías vacías y
  **forma-agnóstico** en `B9:B13` (H3-SEMANTIC: texto `"1"` vs numérico `1` = mismo valor). La
  divergencia es `ERR-ESCRITURA` nombrando hoja + celda + esperado + escrito.

---

## 4. N-1 — hardening del signo en `ParseoMonetarioR10`

`TryParseTexto` acepta `-` **solo** en posición inicial (al frente o detrás de `$`) y rechaza
cualquier otro carácter que no sea dígito, `.` o `,` después del signo. Casos adversariales fijados:
`"01/09/2026"` ⇒ false · `"R10-2026091"` ⇒ false · `"01-09"` ⇒ false · `"01-09-2026"` ⇒ false. Los
10 casos del fixture original siguen verdes y los montos del R10 2026091 no se mueven
(`BloqueR10H1Plan62Tests` y `PegadoR10T0Plan62Tests` re-corridos).

---

## 5. Espectro de reclasificación tocado por H2 (resto: H5/T5.2)

| Fixture | Cambio | Motivo |
|---|---|---|
| `AsientoDetRetriEscritoPlan59Tests` | **eliminado** | su objeto (asiento escrito por el dominio) se retiró; sus asertos viven en `PegadoR10H2Plan62Tests` |
| `DesgloseTrazableAmbasQuincenasPlan59Tests` | **eliminado** | idem (desglose trazable) |
| `EntregaConNovedadesH3Plan61Tests.H3_D10_…` | reclasificado a `VerificarIntegridadPegadoR10` | el gate de valores del asiento ya no existe |
| `IntegridadCalcChainYSelloFechasTests.SelloProceso_*` | reclasificado: `EscribirFechasPeriodo` ya no escribe `D6/D7` | T2.4 (D-31SELLO) |
| `DetRetriCapacidadTests`, `FailFastHojaDetRetriPeriodoTests`, `GoldenDetRetriQ2Tests` | aportan el bloque del R10 | la hoja Det ya la llena el pegado |
| `DetRetriDesgloseTrazableTests.JulioQ2_DetValiRetri_J_EsSale_QuedaEnCero` | reclasificado a `…_TraeElDatoDelR10Pegado` | D-J9/T0-09: J entra al pegado |
| `EntregaConNovedadesH3Plan61Tests.H34/H35/H42` | **siguen rojos** (ya lo estaban) | el pegado elimina la novedad R10 `-1` de septiembre-1 (CA-3); el aserto de novedad se reescribe en H3/T3.2 |
| `DetRetriDesgloseTrazableTests.JulioQ2_DesgloseTrazable_…` | **sigue rojo** (ya lo estaba) | el desglose ya no lo escribe el dominio (D-59H4); lo reescribe H5/T5.2 |

---

## 6. Evidencia de la entrega

- `dotnet build Remuneracion.WinForms/Remuneracion.WinForms.slnx` ⇒ **0 warnings / 0 errors**.
- `PegadoR10H2Plan62Tests` + `PegadoR10T0Plan62Tests` + `BloqueR10H1Plan62Tests` ⇒ **55/55** (incluye
  los 12 casos adversariales de N-1 y la no-regresión de los 10 casos es-CO de H1).
- **Suite completa**: 29 rojos = la MISMA lista de 29 pre-existentes (25 de §9.1 del Plan 61 −
  `ParidadCliTests` + 5 de data-drift del commit `63e0fc8`); **set-diff vacío (0 altas / 0 bajas)**.
- **Hallazgo de entorno (no de código)**: `ParidadCliTests.Paridad_5Ase_Q1_CliVsDirecto_…` falla
  cuando quedan `remuneracion_log_YYYYMMDD_00N.txt` de corridas previas en `bin/Debug/net10.0`: el
  test elige el archivo por `OrderByDescending(nombre)` y `…_00N.txt` gana sobre `…YYYYMMDD.txt`, que
  es donde escribe el sink. Con el directorio de logs limpio el test **pasa** (verificado 1/1). Es la
  fragilidad por la que la lista declarada excluye `ParidadCliTests`.
- SHA256 intactos: plantilla oficial `3086AC…C5214`; R10 2026091 `DD0EAA7B…1122`; manual
  Administrativo `3CE1B869…0916` (el set de 3 huellas es el declarado; en el brief las dos últimas
  venían con la etiqueta intercambiada).
