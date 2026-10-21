---
titulo: Plan 48 - Adopcion plantilla v2 y paridad con manual julio-2
estado: CERRADO 2026-10-19 (Clarification Gate §7 ratificado; review APROBADO-CON-OBSERVACIONES → W-1/S-1/S-2/S-3 cerradas) — H0 (T0 21/21) · H1 (R2 v2 + control recaudo; 492 fórmulas del manual excluidas) · H1b (7 altas cerradas) · H2 (corrida fresca v2 `codigo=0`; R2 `litVsLit=0`; `F10/C14/C15` OK) · H3 (docs). Suite `p48-h3-suite.trx` = 1059/1021/38/0; set-diff vs `o43-final-failed.txt` (39) = 0 altas / 1 baja heredada; SHAs v1/v2 intactos; cero commits.
fecha: 2026-10-19
numero: 48
planes_ocupados: 01..47 ocupados; este es el 48
padre: plans/44 - llenado-r1-fuente-canal-concepto-componentes.md (F1 DONE + F2 Via A/B disuelta por el 46 + F3/paridad-auditoria: la auditoria LLEGO 2026-10-19 como manual v2)
madre_doctrina: plans/46 - correccion-mapeo-interior-r1-canal-efectivo-y-seccion-por-cierre.md (CERRADO 2026-10-17) + plans/47 - revision-demas-pestanas-lente-r1.md (CERRADO 2026-10-17; 38/999/1037 en p47-warnings.trx)
gate_entrada:
  decision_ingeniero_2026-10-19: adoptar la v2 como nueva base de trabajo y lograr paridad con el manual v2, con F0 = T0 read-only obligatorio en el que nada se toca hasta clasificar cada diferencia en bytes
  q1: fuera de alcance (2026071 no se toca; solo se preserva/mide neutralidad)
  plantilla_v1: NO SE TOCA (historica; SHA256 `3086AC317595C2F8E84ED24B1C5EE3E52565DD5193394E3062878E7FA49C5214` intacto antes/despues)
  plantilla_v2: nueva base (SHA256 `FF5CF7A819110A1E1E2CA7BDAD553DD7A25141ADD472B1904BD496F082642632`; verificar intacta antes/despues; 40 hojas, mismo orden y nombres que la v1)
  oraculo_nuevo: ejecucion manual de julio-2 con la v2 (`Docs/Prueba Julio-2/Resultado/Resultado Administrativo con plantilla estandar v2/Formato Remuneracion 202607-2 Total.xlsx`) — NUEVO ORACULO de paridad; reemplaza al anterior (`Remuneracion 202607-2 Total Administrativo.xlsx`, medicion de la v1)
  suite_base: 38 Failed / 999 Passed / 1037 Total / 0 Skipped (`p47-warnings.trx`; referencia de fallos `o43-final-failed.txt` via `p46-cierre.trx`/`p47-warnings-setdiff.txt`); 0 Skip; tests no se debilitan
  sin_commits: NUNCA `git commit`/`push`
---

# Plan 48 - Adopcion plantilla v2 y paridad con manual julio-2

> SDD completo PROPOSE -> DESIGN -> SPEC -> TASKS. Fuente unica de verdad para la implementacion una vez aprobado ESTE documento. No se implementa sin aprobacion del Ingeniero (punto de bloqueo: sin su OK no se abre H1).
> Lectura obligatoria previa: plan 44 completo (incluidas Enmienda E-44.1, nota F2/Via A/B disuelta y R-PLANTILLA-nueva-version) + plan 45 completo (incluida la nota de supersesion por el 46) + plan 46 completo (§0_FP, §2.1/§2.2, §9 cierre con opcion A) + plan 47 completo (§0, §9 T0, §10 H1/H2, §11 W-1..S-1) + `Docs/T0-Plan47-Revision-Pestanas-LenteR1.md` + `.opencode/project-context.md` (doctrinas Planes 21-47, invariantes, allow-list Phase 0 de 68 celdas) + mapas/escritores vigentes (`WorkbookLeafCellMapDetalleR2R4.cs`, `WorkbookLeafCellMapAjustesSfT.cs`, `WorkbookLeafCellMapReporteBanco.cs`, `WorkbookLeafCellMapBalanceSc.cs`, `WorkbookLeafCellMapDetRetri.cs`, `WorkbookLeafCellMapQ2.cs`, `WorkbookLeafCellMapValidaciones.cs`, `WorkbookLeafCellMap.cs`, `EscritorValoresSlotR1.cs`, `CanalEfectivoDetalleR2R4.cs`, `CerosDocumentadosDetalleR2R4.cs`, `ContratoSeccionPeriodFreeR1.cs`, `OpenXmlPlantillaWriter.cs`) + `Docs/T0-Plan45-Contrato-Seccion-PeriodFree.md` (con erratas §12).
> Entorno: Windows, repo con espacios en la ruta; ambas plantillas deben estar CERRADAS en Excel al correr tests (FileShare). Oficial v1 historica: `Docs/Plantilla estandarizada/Formato Remuneracion 2026xx-x Total.xlsx`. Nueva base v2: `Docs/Plantilla estandarizada/v2/Formato Remuneracion 2026xx-x Total vf.xlsx`. Manual v2 (oraculo): `Docs/Prueba Julio-2/Resultado/Resultado Administrativo con plantilla estandar v2/Formato Remuneracion 202607-2 Total.xlsx` (+ `R4-BalanceSubsidioyContribuciones_…xlsx` en la misma carpeta). Insumos: `Docs/Prueba Julio-2/Insumos` (+ subcarpetas por ASE). Corrida de referencia vigente: `p47-h2c` (salida del pipeline sobre la v1).
> Visual Design Intent: N/A (cero UI en este plan).

---

## 0. Contexto verificado en bytes (adoptado del pedido; el H0/T0 lo re-verifica por bytes, no se inventa geometria)

### 0.1 Lo que llego (auditoria de la propietaria, 2026-10-19)

1. **Plantilla v2** `Docs/Plantilla estandarizada/v2/Formato Remuneracion 2026xx-x Total vf.xlsx` — SHA256 `FF5CF7A819110A1E1E2CA7BDAD553DD7A25141ADD472B1904BD496F082642632` (40 hojas, mismo orden y nombres que la v1). **La v1 queda intacta como historica** (`3086AC317595C2F8E84ED24B1C5EE3E52565DD5193394E3062878E7FA49C5214`).
2. **Ejecucion manual de julio-2 con la v2** `Docs/Prueba Julio-2/Resultado/Resultado Administrativo con plantilla estandar v2/Formato Remuneracion 202607-2 Total.xlsx` (+ `R4-BalanceSubsidioyContribuciones_…xlsx` en la misma carpeta). **Este manual v2 es el NUEVO ORACULO de paridad** (reemplaza al anterior `Remuneracion 202607-2 Total Administrativo.xlsx`, que era medicion de la v1).

### 0.2 Evidencia ya levantada por el orquestador (read-only; el T0 la re-verifica y completa)

**A) El oraculo CONFIRMA los Planes 46/47 (validado celula a celula):**

- R1 ASE1 del manual v2 = nuestra corrida `p47-h2c` (R6-R20: `-7.108,82` E; `0` H; `-4.108,55` O; `0` T; `-11.217,37` Mes/AFaseo…).
- R1 ASE5: `603..614=0`, `619/621=52.706,16`, `622-624=49.921,06+402,40=50.323,46`, `625-627=306.792,77+7.697,23=314.490`, `628/629=417.519,62`.
- R2 ASE1: `R3/R4=0` y `R7/R8=104.754.634,94 / 15.816,77` (canal E); LIME: `R67/68=0`, `R71/72=71.702.408,71 / 95.038,85`.

**B) Que cambio la v2 vs v1 (diff numerico real con tolerancia ±0,5):**

- `Reporte Componentes R1`: **16 celdas** — todas la **fila 649 → 0** (los residuos conocidos, limpiados).
- `Rem. Anticipos R2`: **0 diferencias numericas**, pero **+1.323 celdas nuevas**: la v2 **agrega el cuadro "OPORTUNO" por empresa** en cada bloque ASE (ASE1: R47 TOTAL `[F:E21+E32-K21]`, R48 Subsidio `[F:E32]`, R50 RECIPROCIDAD, R53 ENEL `[F:E24+E10+E6-K6-K10]`, R56 ENERBIT, R59 OCCIDENTE `[F:E20+E31-K20]`, R62 CIUDAD LIMPIA-ACUEDUCTO — con y sin formula segun la empresa exista en el ASE; ASE5: R468/474/477/480 con formula, R471/483 sin). Esto **desplaza +11 filas** (ASE2/3) y **+22** (ASE4/5) y los consumidores ya apuntan a las filas nuevas (`'Rem. Anticipos R2'!I144→I155`, `H362→H384`, `E447→E469`… en `CONSOLIDADO_TOTAL RECAUDO`, `DetRetri2026092`, `DetValiRetri2026092`, `REMUNERACION_*`).
- `Valida - Control Recaudo`: la v2 trae **formulas nuevas** (`E7=SUM(D7:D9)`, `E10=SUM(D10)`, `F10=E7+E10`) donde la v1 tenia valores (caches en 0).
- `VALIDACION_*` (ENEL/ENERBIT/OCCIDENTE/RECIP/TOTAL): `C15: 0→1` prefijado y `C14` sin cache.
- Resto de hojas: **sin diferencias numericas reales**.

**C) Quien cambio que (diffs decisivos):**

- `v1 → nuestra corrida`: **fDiff=0** (la App NO toca formulas — invariante intacto).
- `v2 → manual`: **fDiff=492 en R1** ⇒ **la propietaria cambio 492 formulas en su ejecucion manual**: patron `G→H` (p. ej. `P570: G486→H486`, `J95: G46+G15→H46+H15`, `I264: G150+G112→H150+H112`) y **formulas nuevas** en `W682/S682/AN682` (=`H629`) y **`F702 = F659+F628-L628`** (zona del residual 24.329,38). En R2, `fDiff=0`.
- **La plantilla v2 `vf` NO trae esas 492 formulas** → punto a preguntar a la propietaria (¿la referencia oficial es la v2 tal cual o la v2+correcciones del manual?).

**D) Metodologia y herramientas (read-only, en `%TEMP%\opencode\`):** `dumpxlsx.ps1` (`-Path -Sheet -From -To -Cols -List -OnlyValues`), `diffv.ps1` (formulas con *shared* expandidas + valores crudos), `cmpnum.ps1` (**comparador numerico con `-Tol` (0.5)** que clasifica numDiff/textDiff/mixto/onlyA/onlyB). **Advertencias obligatorias:** (i) el diff textual crudo es **ruido de representacion** (`-335335691.51` vs `-335335691.50999999` = mismo numero); (ii) los **caches de formula** (area `<v>` de celdas con `<f>`) difieren entre nuestro libro (preservado) y el del Excel de ella (recalculado) — **jamas son «contenido escrito»**.

### 0.3 Geometria y contratos congelados contra la v1 (adoptados; el T0 verifica que falla y donde)

- Anclas R1 vigentes (v1): ASE1 cabecera R4, bloque 6-72, fronteras 16/32/47/59/70; ASE5 cabecera R601, bloque 603-676, fronteras 613/629/648/660/674; secciones oficiales citadas en el 46 §0.2.
- Bloques R2/R4 vigentes (v1): `BloquesR2` ASE1 3-33, ASE2 67-113, ASE3 164-194, ASE4 282-328, ASE5 382-419; `BloquesR4` ASE1 3-15, ASE2 93-111, ASE3 194-206, ASE4 231-249, ASE5 346-358 (Plan 47 §0.3). Con la v2 estos rangos **se desplazan** (§0.2-B: +11/+22) y el T0 debe re-derivarlos por bytes.
- Contrato R-PLANTILLA-nueva-version (Plan 44 §1.4 R-PLANTILLA): contrato por etiquetas, no por direcciones; `ContratoSlotR1T0Tests` + `GeometriaPlantillaEstandarizadaT0Tests` rompen primero (fail-fast `ERR-PLANTILLA`); re-congelamiento por T0 nuevo, nunca edicion silenciosa. Este plan ES ese T0 nuevo para la v2.
- Frente R1 del 46 y R2/R4 del 47 **ya cerrados y validados por el oraculo** (§0.2-A): no re-litigar; solo lo que la v2/manual exijan.

### 0.4 Linea base y puertas heredadas (no re-litigar; cada gate las re-verifica)

- Linea base de suite a la entrada: `p47-warnings.trx` = **1037 Total / 999 Passed / 38 Failed / 0 Skipped**; referencia de fallos `o43-final-failed.txt` via `p46-cierre.trx`/`p47-warnings-setdiff.txt`. Criterio vinculante: **set-diff sin altas**; toda baja documentada una por una con causa byte-probada; **rojo nuevo = STOP con causa** (ver §3-S0).
- Puertas intactas y referenciadas, no reabiertas: Q1 fuera de alcance; v1 historica intacta; v2 intacta antes/despues; cero commits; D-H (cero literales/ramas por periodo/empresa en produccion); values-only estricto; cero insert/delete; `<f>` intacta (el unico modulo autorizado a tocar formulas es el preservador de *shared formulas*); allow-list Phase 0 de 68 celdas intacta; tolerancia ±0.5 solo en comparacion; CRLF sin mezclar; `<v>` nunca es oraculo; declarada = ultima puerta con doble huella y fail-closed (Plan 42); prohibido fallback posicional (Plan 41 E1-1) y prohibido sellar/apagar/re-etiquetar gates (Plan 40 E1-1..E1-5).

---

## 1. PROPOSE

### 1.1 Intent

Adoptar la v2 como nueva base de trabajo y lograr paridad con el manual v2, con **F0 = T0 read-only obligatorio** en el que **nada se toca** hasta que quede clasificado con bytes. Cada pestaña/hoja termina en uno de dos estados, sin tercero: **(a) delta probado y corregido**, o **(b) sin delta con evidencia**.

### 1.2 In-scope EXPLICITO (lo que el T0 del Plan 48 DEBE resolver; prohibido suponer)

| ID | Pregunta vinculante | Criterio de salida |
|---|---|---|
| **P1 — matriz v1→v2 completa** | Por hoja: estructura (filas/columnas ganadas o movidas, con la lista de bloques y sus rangos nuevos), formulas (texto, con *shared* expandida), literales. Foco especial: R2 (bloques OPORTUNO, rangos nuevos por ASE y que filas son **formula vs escritura**), R1 (fila 649 y zona 682/702), `Valida - Control Recaudo`, `VALIDACION_*` | Matriz congelada en fixture read-only (§2.2, §4-H0); sin matriz verde no hay H1 |
| **P2 — clasificacion nuestra-corrida vs manual v2** | Clasificar TODAS las diferencias entre `nuestra corrida (p47-h2c, v1)` y `el manual v2`: (i) **geometria v2** (esperadas), (ii) **caches de formula** (no contenido), (iii) **contenido real** (literales distintos / faltantes / sobrantes) → estas ultimas son las brechas a cerrar | Herramientas: `cmpnum -Tol 0.5` + `diffv` + comparacion de **literales contra literales** (celda sin `<f>` en ambos). Cada numDiff/textDiff real cita causa |
| **P3 — delta del manual** | Inventario de lo que la propietaria escribio/corrigio al ejecutar (las 492 formulas de R1 + literales por hoja) | Lista citada por hoja (celda + antes/desde v2 + despues/manual + patron G→H o formula nueva); sin inventario no hay veredicto P5 |
| **P4 — veredicto de factibilidad de adopcion** | ¿La App puede construir sobre la v2? (preflight/anclas/contratos actuales congelados contra la v1 → ¿que falla y donde? El R-PLANTILLA-nueva-version del Plan 44 ya preve fail-fast `ERR-PLANTILLA` + re-congelamiento por T0) | Lista de que falla y donde + que mapas/contratos hay que re-congelar (R1, R2/R4 detalle, CerosDocumentados, cell maps del CONSOLIDADO/DetRetri/Validaciones, anclas, allow-list Phase 0) |
| **P5 — preguntas para la propietaria** | (a) ¿la referencia oficial de R1 son las **492 formulas corregidas del manual** o la v2 tal cual?; (b) el cuadro **OPORTUNO** de R2: ¿sus filas son solo formula o alguna se llena con valores?; (c) ¿la v2 es la base desde ahora para todos los periodos (y la v1 queda como historica)?; (d) el `R4-Balance…` de la carpeta, ¿es fuente o anexo? | Documentadas en el plan para que el Ingeniero las traslade; las respuestas que cambien geometria/filtros entran como errata fechada, nunca como edicion silenciosa |
| **P6 — tabla de paridad por hoja** | Las 40 hojas: numDiff reales/literales, clasificados por causa | Tabla con veredicto por hoja (delta probado vs sin delta con evidencia) |

### 1.3 Non-goals / out (vinculante)

- **Q1 (2026071): fuera de alcance.** No se cambia comportamiento Q1; Q1 solo se preserva/mide (neutralidad). Cualquier cambio de comportamiento Q1 exige autorizacion explicita separada.
- **Plantilla v1: NO SE TOCA** (historica; SHA256 intacto). **Plantilla v2: NO SE EDITA** (ni por residuos ni por hueco aparente); si el T0 encuentra hueco, se reporta al dueño, no se parchea.
- **Sin commits.** NUNCA `git commit`/`push`.
- **R1 interior (46) y R2/R4-detalle (47): cerrados y validados por el oraculo** (§0.2-A). No se re-litigan; solo se re-congelan contra la geometria v2 donde el T0 pruebe desplazamiento.
- **Gates/evaluadores y agregados no se "pegan":** `VALIDACION_TOTAL` se evalua (Plan 40), no se pega; `CONSOLIDADO_TOTAL RECAUDO` / `AJUSTES - SF-T` / `REMUNERACION_*` / `GERENTES_*` son arrastre — solo se tocan si el T0 prueba arrastre roto con bytes (y entonces el fix vive en el origen, no en el agregado).
- **Frente R1 del 46 vs las 492 formulas del manual:** no se re-mapea nada por intuicion; el veredicto P5(a) decide si esas formulas entran al contrato (y como) o quedan como delta documentado del manual.

### 1.4 Riesgos con mitigacion

| ID | Riesgo | Mitigacion |
|---|---|---|
| R-GEOM | Los rangos congelados contra la v1 (R2/R4, anclas, cell maps) apuntan a filas viejas en la v2 (+11/+22) y el pipeline falla o escribe cruzado | P4 bloqueante: preflight/anclas/contratos re-verificados contra la v2 en H0; fallo declarado = STOP con `ERR-PLANTILLA` que nombra hoja+ASE+celda; re-congelamiento SOLO con evidencia citada |
| R-ATRIBUCION | Atribuir a colocacion lo que es geometria v2, cache de formula o delta del manual (o viceversa) | P2 con triple clasificacion (geometria / caches / contenido real) + literales-contra-literales; `cmpnum -Tol 0.5` + `diffv`; prohibido `<v>` stale como oraculo; sin veredicto citado no hay fix |
| R-RUIDO | El diff textual crudo y los caches recalculados se leen como «contenido escrito» | Doctrina §0.2-D: tolerancia ±0.5 numerica; caches jamas son contenido; `fDiff` con *shared* expandida; `values_changed` como unica evidencia de contenido |
| R-MANUAL | Las 492 formulas del manual se adoptan o ignoran sin criterio (G→H, W682/S682/AN682, F702) | P3 + P5(a): inventario citado + pregunta explicita; sin respuesta, el T0 declara el tratamiento (contrato vs delta documentado) y H1 no lo toca sin autorizacion |
| R-OPORTUNO | El cuadro OPORTUNO de R2 se trata como escritura cuando es formula (o viceversa) | P1/P5(b): por cada fila del cuadro, citar si tiene `<f>` o es literal escribible; el writer solo toca literales (guard anti-formula intacto) |
| R-FALSO-HUECO | Mantener una omision que la v2/manual revelan como contraparte real (precedente del 46) | P2/P4: ninguna omitida sobrevive sin prueba de ausencia en bytes contra la v2 y el manual |
| R-OTRA-LENTE | Aplicar la lente R1 a una hoja de otra lente (banco × disposicion; agregados; controles) | P6 con criterio citado por hoja; fuera = fuera con criterio, no por intuicion |
| R-PERIODO | Tocar hojas de otro periodo (`DetRetri2026092`/septiembre) corriendo julio-2 (Plan 30) | La explicacion P5 del 47 sigue vigente: base/periodo equivocado = `ERR-PLANTILLA`, nunca escritura cruzada |
| R-FALSO-VERDE | Apagar/sellar/re-etiquetar un gate para pasar (E1-1..E1-5 Plan 40) | Prohibido `SeGatea=false`; sub-bloque no evaluable = hallazgo/defecto con frente propio, nunca verde |
| R-Q1 | Contagio a Q1 al re-congelar contra la v2 | Q1 fuera de alcance + gate G-Q1 (§3-S0); cualquier movimiento Q1 = STOP con causa byte-probada |
| R-FIXTURE | Fixtures que pinchan la geometria v1 actualizados por conveniencia | Actualizacion SOLO con evidencia citada (celda v2 + rango nuevo + causa); cada assert cambiado cita bytes; R-FIXTURE vigente |
| R-CAMBIO-EXTRA | Cambios de celda corrida-nueva→fresca sin justificacion trazada | Todo cambio justificado por P1/P2/P3; cambio inexplicado = STOP |

---

## 2. DESIGN

### 2.1 Tabla de alcance por hoja (el T0 la confirma o corrige con bytes; sin T0 verde no hay H1)

| Hoja | Lente a la entrada | Veredicto esperado (a confirmar por T0) | Criterio |
|---|---|---|---|
| `Reporte Componentes R1` | **Directa** (fuente/canal/concepto/componentes; reglas 46 §2.1/§2.2) | **DENTRO** (re-congelar fila 649 + zona 682/702 + delta 492 formulas del manual) | Pegado fila-a-fila por firma + encabezado; sin canal explicito rige canal efectivo |
| `Rem. Anticipos R2` | **Directa** (C=fuente, D=concepto, firma A\|B\|C\|D; canal efectivo compuesto del 47) | **DENTRO** (re-congelar bloques OPORTUNO + desplazamientos +11/+22; formula vs escritura por fila) | Pegado por (firma + canal efectivo) + encabezado E..R |
| `Reversion Pagos R4` | **Directa** (A=fuente, B=rol, C=concepto, firma A\|B\|C) | **DENTRO si el T0 prueba desplazamiento** (los bloques se mueven con R2; `Otros` por encabezado intacto) | Idem R2; columnas D..R por encabezado |
| `SALDOS POR NOTA` | Parcial (cuadro por concesionario; mapeo por titulo + `Vlr Intereses` opcional, Plan 47) | **Medir** (¿la v2 la mueve? — lo decide el T0 con bytes) | Sin suponer derivacion: el T0 cita insumo→celda o declara identidad |
| `Valida - Control Recaudo` | Control (formulas nuevas `E7/E10/F10` en v2) | **Medir** (¿arrastre roto? — el T0 lo prueba) | Control: se evalua, no se pega; el fix viviria en el origen |
| `VALIDACION_*` (ENEL/ENERBIT/OCCIDENTE/RECIP/TOTAL) | Gate/evaluador (`C15: 0→1`, `C14` sin cache en v2) | **FUERA salvo hallazgo T0** | Se evalua (Plan 40), no se pega |
| `CONSOLIDADO_TOTAL RECAUDO` (consumidores `I144→I155`, `H362→H384`, `E447→E469`…) | Arrastre (ya apunta a filas nuevas en v2) | **Medir** (¿arrastre consistente con los bloques nuevos? — lo prueba el T0) | Agregado; el fix viviria en el origen, no en el agregado |
| `DetRetri2026092` / `DetValiRetri2026092` / `REMUNERACION_*` / `GERENTES_*` | Arrastre del periodo / metadatos (ya apuntan a filas nuevas) | **FUERA salvo arrastre roto probado** | Periodo por dominio (Plan 30); prohibida escritura cruzada |
| `REPORTE RECAUDO x BANCO` (30) / `BCE SC POR FACT.` (10) | Otra lente / agregado | **FUERA salvo hallazgo T0** (precedente 47: banco alineado por etiqueta; BCE agregado) | Lente distinta / agregado |
| `Recaudo *` (36/36/18/8) | Pegado real desde `Conciliaciones/RESUMEN MES` (Plan 47 P5) | **Medir** (¿la v2 las mueve? — lo decide el T0) | Fuente citada antes de tocar |
| `INTERVENTORIA`, `RETRIBUCION NEGATIVA`, `ANT EXT-REV`, `ANTICIPOS USUARIOS`, resto con 0 cambios | Intactas en v1 | **FUERA salvo falso hueco probado** (P2/P4) | Cero cambios ≠ cero prueba: el T0 las barre contra v2 y manual |

### 2.2 Reglas a aplicar (cuando apliquen; el T0 decide el "cuando")

- **R46-1 canal efectivo + R46-2 seccion por cierre (frente R1):** vigentes e intactas; se re-aplican sobre la geometria v2 (fila 649 en 0, zona 682/702 con `F702=F659+F628-L628` del manual como dato a clasificar, no como spec).
- **Canal efectivo (analogo R2/R4 del 47) + canal compuesto por segundo cierre (W-3):** vigentes; se re-derivan sobre los bloques desplazados (+11/+22) y sobre el cuadro OPORTUNO (que filas componen canal y con que segundo cierre).
- **Firma/clave por hoja (P6 del 47):** el T0 congela por pestana la clave efectiva (R2 `(C-fuente, D-concepto, firma-A|B|C|D, seccion/bloque-ASE)`; R4 `(A-fuente, B-rol, C-concepto, firma-A|B|C, seccion/bloque-ASE)`) con regla de desempate citada cuando falte etiqueta. **Prohibido casar por orden fisico** (Plan 41 E1-1): sin clave no hay colocacion; fila sin slot probado = omision probada o fail-fast accionable.
- **Columnas por encabezado, nunca por indice:** R1 F..Z, R2 E..R, R4 D..R ya se resuelven por encabezado; el T0 re-verifica cada par (encabezado-fuente → columna-destino) contra la v2 y lista los sin-contraparte (¿0 explicito legitimo? ¿falso hueco?).
- **Formulas:** el writer solo toca literales (celda sin `<f>`). Cuadro OPORTUNO y `Valida - Control Recaudo`: cada fila citada como formula o escribible; celda destino con `<f>` ⇒ `ERR-PLANTILLA`, nunca sobrescritura.

### 2.3 Arquitectura del cambio (solo si el T0 prueba desvio; contract-first)

- **Core (puro, sin I/O, sin periodo/empresa — D-H):** predicados puros existentes (`ContratoSeccionPeriodFreeR1.SeccionPorCierre`/`.CanalEfectivo`, `CanalEfectivoDetalleR2R4`, `CerosDocumentadosR2/R4`) re-congelados contra la geometria v2 (rangos nuevos, bloques OPORTUNO, fila 649, zona 682/702). Reutilizar los contratos puros donde apliquen; nada de parches por periodo/empresa. `CerosDocumentados*` solo crece con motivo y evidencia citados (doctrina E-44.1).
- **Infrastructure (values-only):** consumo en el writer/mapa que el T0 señale (`WorkbookLeafCellMapDetalleR2R4` / `AjustesSfT` / resto + `OpenXmlPlantillaWriter` + `EscritorValoresSlotR1`), preservando guard anti-formula, decimal exacto, cero `InsertRow`/`DeleteRow`, preservacion de *shared formulas* (Plan 33), allow-list Phase 0 intacta.
- **Compatibilidad:** `SeccionEspejoR1.PorFronteraArriba`, `R1FirmaInterior.EsAplicacionPorFrontera` / `EsDatoEmpresa`, `SinonimosEmpresaR1`, consumidor Plan 35-T2 y todo el frente R1 del 46 quedan **intactos** salvo re-congelamiento geometrico citado.
- **Neutralidad Q1:** Q1 fuera de alcance; el H1 preserva plan-then-write / comportamiento historico donde exista y el gate G-Q1 lo verifica.

### 2.4 Fixtures y evidencia que el T0 debe congelar

- Fixture read-only nuevo (BCL ZIP+XML, cero produccion; gemelo de `ContratoSeccionPeriodFreeT0Plan46Tests` y `RevisionPestanasLenteR1T0Plan47Tests`): matriz v1→v2 por hoja × ASE con cada celda citada (fila-fuente + celda-slot + firma/clave + canal efectivo + frontera de cierre + formula-vs-literal) + lista de omitidas probadas + tabla P6 de paridad por hoja (40) + veredictos P1..P5.
- Referencias de conteo a re-verificar: R1 16 (fila 649) · R2 +1.323 celdas nuevas (OPORTUNO) con 0 numDiff · resto §0.2-B; `fDiff` v1→corrida = 0 (R1) y v2→manual = 492 (R1) / 0 (R2).

---

## 3. SPEC (criterios numerados + gates no negociables)

### S0 — Gates transversales (toda T)

- G-BUILD: `dotnet build Remuneracion.WinForms/Remuneracion.WinForms.slnx` → **0 warnings / 0 errores**.
- G-SUITE: suite completa con trx + **set-diff vs linea base `p47-warnings.trx` (38/999/1037; referencia `o43-final-failed.txt`)**: altas/bajas una por una. **Rojo nuevo = STOP con causa citada** (no se maquilla). 0 `Skip`, 0 `Assert.True(true)`, 0 tests borrados.
- G-Q1: Q1 intacto (fuera de alcance); **cualquier movimiento Q1 = STOP con causa byte-probada**.
- G-E2E: corrida fresca a temp (regen del pipeline sobre `Docs/Prueba Julio-2/Insumos` con **la v2 como unica entrada**) + verificacion fila a fila §S1; **todo cambio justificado por P1/P2/P3**; cambio inexplicado = STOP.
- G-SHA: v1 SHA256 intacto (`3086AC31…C5214`) + v2 SHA256 intacto (`FF5CF7…2632`) antes/despues.
- G-PROC: cero commits; CRLF sin mezclar en lo tocado.
- G-H0: T0 read-only deriva la matriz §2.1/§2.2 y la congela en fixture read-only ANTES de tocar produccion (puerta H0→H1). **Sin T0 verde no hay H1.**
- G-D-H: cero literales/ramas por periodo o empresa en produccion (grep de cierre).
- G-WORKBOOK: todo gate de negocio lee el workbook generado (texto-`<f>` + literales; prohibido `<v>` stale; diff de `<f>` con *shared* expandida; solo `values_changed` con `cmpnum -Tol 0.5` es evidencia de contenido).
- G-FORMULA: `<f>` byte-identica en todo el diff salvo allow-list Phase 0 y salvo lo que el veredicto P5(a) autorice explícitamente (grep: cero `new CellFormula` fuera del preservador; cero `InsertRow`/`DeleteRow`).

### S1 — Criterios H0/T0 (read-only, cero codigo de produccion)

- S1.1: matriz v1→v2 por hoja (P1) congelada en fixture read-only con firma/clave por fila y formula-vs-literal citada (foco: R2 OPORTUNO por ASE con rangos nuevos, R1 fila 649 y zona 682/702, `Valida - Control Recaudo`, `VALIDACION_*`).
- S1.2: clasificacion P2 completa (geometria v2 / caches / contenido real) con `cmpnum -Tol 0.5` + `diffv` + literales-contra-literales; cada numDiff/textDiff real cita causa.
- S1.3: inventario P3 del delta del manual (492 formulas R1 + literales por hoja) con patron citado (G→H, W682/S682/AN682, F702).
- S1.4: veredicto P4 de factibilidad (que falla y donde; que mapas/contratos re-congelar: R1, R2/R4 detalle, CerosDocumentados, cell maps CONSOLIDADO/DetRetri/Validaciones, anclas, allow-list Phase 0).
- S1.5: preguntas P5 documentadas para la propietaria (a/b/c/d) + tabla P6 de paridad por hoja (las 40) con veredicto por hoja.
- S1.6: medicion fila a fila del oraculo §0.2-A reproducida (R6-R20 ASE1, R603-R629 ASE5, R3/R4 + R7/R8 y R67/R68 + R71/R72 de R2) como ancla de paridad.

### S2 — Criterios H1 (adopcion de la v2, solo lo que el T0 probo)

- S2.1: cada re-congelamiento cita T0 (rango nuevo + celda + firma/clave + canal/seccion + formula-vs-literal); reutiliza contratos puros de Core; cero parches por periodo/empresa.
- S2.2: el pipeline construye sobre la v2 (preflight + anclas + mapas resuelven la geometria nueva); fallo declarado = STOP con `ERR-PLANTILLA` accionable (hoja+ASE+celda). Hojas sin delta **intactas** (cero cambios de produccion en ellas; el fixture T0 es su prueba).
- S2.3: frente R1 del 46 y R2/R4 del 47 intactos salvo geometria v2 citada; consumidor Plan 35-T2 intacto; `DivergenciasDeclaradasCatalogo` intacto (aqui no se re-decide nada de declaradas).
- S2.4: `<f>` intacta salvo veredicto P5(a); Y/Exonera-* y columnas sin contraparte intactas salvo falso hueco probado; decimal exacto; guard anti-formula intacto.
- S2.5: build 0 warnings + greps D-H + G-FORMULA + G-SHA.

### S3 — Criterios H2 (paridad: corrida fresca sobre la v2 vs manual v2)

- S3.1: corrida fresca sobre la v2 a temp + matriz fila a fila verde + `<f>` byte-identica (salvo P5a) + set-diff vs 38/999/1037 + Q1 intacto + SHAs intactos.
- S3.2: tabla corrida-fresca→manual-v2 con justificacion trazada por celda (P1/P2/P3; geometria/caches/contenido); cambio inexplicado = STOP. Brechas de contenido cerradas por regla geometrica (nada de parches por periodo/empresa).
- S3.3: numeros frescos capturados como insumo (sin re-decidir nada fuera de este plan).

### S4 — Criterios H3 (docs/cierre)

- S4.1: doctrina del plan en `.opencode/project-context.md` (solo lo probado: v2 como base, OPORTUNO, fila 649, zona 682/702, veredicto 492 formulas).
- S4.2: estados en `plans/44` (R-PLANTILLA-nueva-version ejercida; paridad-auditoria atendida con el manual v2) + nota de adopcion en `plans/46`/`plans/47` (re-congelamiento geometrico, sin re-litigar frentes) sin reescribir historia; desviaciones declaradas con fecha/motivo/evidencia.
- S4.3: diff de docs citado; cero codigo de produccion en H3.

---

## 4. TASKS (fases H0..H3 con puertas)

### H0/F0 — T0 read-only (bloqueante; cero codigo de produccion; nada se toca)

1. H0.1 Re-verificar por bytes (BCL ZIP+XML; `dumpxlsx.ps1` / `diffv.ps1` / `cmpnum.ps1 -Tol 0.5`): SHAs v1/v2; 40 hojas (orden/nombres); matriz v1→v2 por hoja (estructura/formulas/literales; foco §2.1); que escribe el runtime en cada hoja y desde que insumo; reproducir el oraculo §0.2-A como ancla.
2. H0.2 Clasificacion P2 + inventario P3 (492 formulas + literales) + veredicto P4 (que falla y donde) + preguntas P5 + tabla P6 (40 hojas) con veredicto por celda/hoja.
3. H0.3 Congelar en fixture read-only nuevo (cero produccion) + doc T0 con tablas y veredictos 1–6; medicion fila a fila vs insumo, vs corrida `p47-h2c` y vs manual v2.
- **Puerta H0→H1:** fixture read-only verde + matriz citada + veredictos P1..P6. Sin esto no se abre H1. Si la matriz no absorbe la geometria v2 sin ramas, STOP declarado.

### H1/F1 — Adopcion de la v2 (solo donde el T0 probo desvio)

1. H1.1 Core: re-congelar predicados/contratos puros contra la geometria v2 (D-H; reutilizar contratos; `CerosDocumentados*` solo con evidencia).
2. H1.2 Consumo values-only en el writer/mapa señalado (guard anti-formula, decimal exacto, encabezado-no-indice, cero insert/delete); preflight/anclas/cell maps resuelven la v2.
3. H1.3 Tests focalizados por frente re-congelado (verdes) + build 0 warnings + greps D-H + G-FORMULA + G-SHA.
- **Puerta H1→H2:** S2.1..S2.5 verdes. Sin esto no se abre H2.

### H2/F2 — Paridad (corrida fresca sobre la v2 vs manual v2)

1. H2.1 Corrida fresca a temp desde la v2 + matriz fila a fila verde + `<f>` byte-identica (salvo P5a) + set-diff vs 38/999/1037 + Q1 intacto + SHAs intactos.
2. H2.2 Tabla corrida-fresca→manual-v2 justificada por celda (P1/P2/P3); brechas de contenido cerradas por regla geometrica; cambios inexplicados = STOP.
3. H2.3 Code-review independiente (spec compliance + calidad): trazabilidad cada cambio → P1/P2/P3; ningun test debilitado; CRLF; SHAs.
- **Puerta H2→H3:** review APROBADO (o APROBADO CON OBSERVACIONES cerradas) + set-diff sin rojos nuevos + G-Q1 + SHAs.

### H3/F3 — Docs/cierre

1. H3.1 Doctrina probada en `.opencode/project-context.md`; estados en `plans/44` (+ notas en `46`/`47` si aplica); desviaciones declaradas con fecha/motivo/evidencia.
2. H3.2 `Docs/Instructivo-Capa-B.md` si aplica (auditoria fila a fila con las claves nuevas + OPORTUNO + `ERR-SIN-SLOT`/`ERR-PLANTILLA`).
- **Puerta H3:** diff de docs citado en el resumen de cierre.

---

## 5. Riesgos (tabla consolidada + gates por fase)

Ver §1.4 (R-GEOM, R-ATRIBUCION, R-RUIDO, R-MANUAL, R-OPORTUNO, R-FALSO-HUECO, R-OTRA-LENTE, R-PERIODO, R-FALSO-VERDE, R-Q1, R-FIXTURE, R-CAMBIO-EXTRA). Gates por fase:

- **Gate H0→H1:** fixture T0 verde + matriz citada + P1..P6 resueltos. Sin T0 verde no hay H1.
- **Gate H1→H2:** S2 verde (solo geometria probada) + build 0w + greps D-H + G-FORMULA + G-SHA.
- **Gate H2→H3:** set-diff sin altas + Q1 intacto + SHAs intactos + tabla fresca→manual justificada + review aprobado.
- **Gate H3:** diff de docs citado; cero codigo de produccion en H3.

---

## 6. Architecture Validation Certificate

- **SOLID:** SRP (T0 read-only / predicados puros en Core / pegado values-only en Infrastructure / matrices como DATOS del T0); OCP (hoja o bloque nuevo = fila citada en T0 nuevo, cero ramas nuevas; periodo nuevo = cero cambios por D-H y geometria pura); LSP/DIP (writers dependen de mapas + predicados + `IPlantillaWriter`/`IWorkbookLeafWriter`, no de concretos de periodo); ISP (contratos `IRecaudoReader`/`IPlantillaWriter`/`IWorkbookLeafWriter`/`IEspejoR1Writer`/`IDetRetriR10Reader` intactos).
- **Best Practices:** contract-first (H0 congela bytes de la v2 antes de H1); fail-fast accionable (`ERR-PLANTILLA`/`ERR-SIN-SLOT` con hoja+ASE+fila/firma+canal); gates que leen el workbook (texto-`<f>` + literales, nunca `<v>`); aceptacion fila a fila + `cmpnum -Tol 0.5` + set-diff contra 38/999/1037 + G-Q1 + doble SHA256; CRLF; sin commits.
- **Performance:** etiquetado/clave una vez por fila/slot (O(filas + slots), sin re-lectura por celda); agrupacion por (firma, seccion/bloque, canal) con cursor por grupo donde aplique; lectura acotada en H0 (BCL ZIP+XML); regens a temp fresca.

---

## 7. Clarification Gate (requiere aprobacion del Ingeniero; sin esto no se abre H1)

1. **Alcance P1 + P4:** ratificar la tabla §2.1 (que hojas entran/salen y con que criterio) y el veredicto de factibilidad (que falla y donde; que mapas/contratos se re-congelan).
2. **Reglas y clasificacion P2/P3:** ratificar la triple clasificacion (geometria / caches / contenido real) y el inventario del delta del manual (492 formulas + literales) antes de tocar nada.
3. **Preguntas P5 para la propietaria:** ratificar que el Ingeniero traslada (a) referencia oficial de R1 (¿492 del manual o v2 tal cual?), (b) OPORTUNO (¿formula o escritura?), (c) v2 como base desde ahora (v1 historica), (d) `R4-Balance…` (¿fuente o anexo?).
4. **Puerta H0→H1:** sin H0 verde (matriz citada + fixture read-only + P1..P6) no se implementa. Si H0 declara STOP, no se abre H1.
5. **Q1 / plantillas / auditoria / commits / suite:** ratificar que Q1 sigue fuera, la v1 intacta, la v2 intacta, el manual v2 como nuevo oraculo, cero commits, y 38/999/1037 como base del set-diff (rojo nuevo = STOP).

Sin estas aprobaciones el plan no avanza de H0 a H1 (1, 2, 4, 5) ni cierra H2 (toda la regresion + §S3) ni H3 (4, 5).

---

## 8. Estado

**APROBADO 2026-10-19 (Clarification Gate §7 ratificado). H0 EJECUTADO (T0 read-only verde, fixture 21/21). H1 EJECUTADO (alcance probado; las 492 formulas del manual quedan EXCLUIDAS). H2 VERIFICADO (corrida fresca sobre v2 `codigo=0` + tabla vs manual). H3 PENDIENTE.** Code-review H2→H3 **NO lanzado**: lo comisiona el Ingeniero al verificar la evidencia.

---

## 9. EJECUCION H0/H1/H2 (2026-10-19)

### 9.1 H0 — T0 read-only

Ver `Docs/T0-Plan48-Adopcion-Plantilla-v2.md`: fixture `AdopcionPlantillaV2T0Plan48Tests` (21/21), matriz v1→v2 P1..P6, sonda P4 = `ERR-PLANTILLA` accionable en `Rem. Anticipos R2!E144`, SHAs intactos, cero commits.

### 9.2 H1 — adopcion (solo lo probado; 492 EXCLUIDAS)

| # | Cambio | Archivo:linea | Traza T0 |
|---|---|---|---|
| 1 | `BloquesR2` → v2 (3-33, 67-124, 175-205, 293-350, 404-441) | `Remuneracion.Infrastructure/Excel/WorkbookLeafCellMapDetalleR2R4.cs:85` | §2.2 / §4.2 |
| 2 | `CerosDocumentados` → **3 claves A/EAB** (retira las 9 del falso hueco) | `Remuneracion.Core/Models/CerosDocumentadosDetalleR2R4.cs:54` | §2.2 / §2.4 |
| 3 | Cadena R2 Q2 (editables/protegidas) → `E101/E123/K101`…`E426/E440/K426`; visibles `E155/E273/E384/E468` | `Remuneracion.Infrastructure/Excel/WorkbookLeafCellMapQ2.cs:166-182` | §1.3 / §4.2 |
| 4 | CONSOLIDADO `D29:D32` → `E155/E273/E384/E468` | `WorkbookLeafCellMapQ2.cs:225-228` + `OpenXmlPlantillaWriter.cs:1016-1019` | §1.7 |
| 5 | DetRetri/DetValiRetri cadenas → `E155/E273/E384/E468` y `L155/L273/L384/L468` | `WorkbookLeafCellMapQ2.cs:283-298` | §1.7 |
| 6 | Mapa multi-ASE (`WorkbookLeafCellMapPorAse`) → mismas celdas v2 | `WorkbookLeafCellMapPorAse.cs:37-60,88-129` | §1.3 |
| 7 | Reader `MapearR2` → celdas v2 | `ExcelDataReaderWorkbookLeafInputReader.cs:1655-1680` | §1.3 |
| 8 | `Valida - Control Recaudo`: `D7:D10` por ETIQUETA desde `Recaudo *!F27`; `F10` solo si **no** es formula | `OpenXmlPlantillaWriter.cs:1717-1782`; tabla `WorkbookLeafCellMapValidaciones.cs:57` | §1.5 / §4.2 |
| 9 | Resolutor soporta `SUM(celda)` (`E10=SUM(D10)`) | `ResolutorValorValidacion.cs:124` | §1.5 |
| 10 | `Insumos.PlantillaOficial()` → v2 (la v1 queda historica) | `Remuneracion.IntegrationTests/Insumos.cs:42` | §4.3 |
| 11 | R1 649: **sin cambio de mapa**; assert de no regresion (el mirror la escribe) | `AdopcionPlantillaV2T0Plan48Tests` | §1.4 |
| 12 | R1 682/702 **EXCLUIDO** por decision del Ingeniero (no se toca, no se simula) | — | §5 / P5(a) |

**Tests re-congelados (R-FIXTURE, evidencia T0 citada):** `DetalleR2R4EscrituraTests` (bloque `67..124`; el VI del canal 2 **se escribe**; ASE5 `408/410/412/414/425`), `RevisionPestanasLenteR1T0Plan47Tests` (ceros → 3 claves + slots v2), `GeometriaPlantillaEstandarizadaT0Tests` (v1 historico vs v2 geometrico; `D111` ausente), `AdopcionPlantillaV2T0Plan48Tests` (P2 con retiro), `Insumos`.

**Fuera de H1 (sin cambio):** `WorkbookLeafCellMapPorEmpresa.EditablesR2PorEmpresa` (path por-empresa Q1/omitido Q2), R1 682/702, `DivergenciasDeclaradasCatalogo`, frente R1 del 46 (salvo assert 649), `SALDOS`/banco/BCE/AJUSTES/`Recaudo *`/columnas excluidas de `DetRetri`.

### 9.3 H2 — corrida fresca sobre v2 vs manual v2

- CLI `--periodo 2026072 --carpeta Docs/Prueba Julio-2/Insumos --plantilla <v2> --salida <temp>` → **RESULTADO OK `codigo=0`**.
- `litVsLit` (contenido real, `cmplit`): **R2 = 0**, **R1 = 0**, **R4 = 0**, **SALDOS = 0**, **banco = 0**, **AJUSTES = 0**, **REMUNERACION_\* = 0**, **GERENTES_\* = 0**.
- **Brecha R2 cerrada:** suma por bloque `ΔE ≈ 5,9e-8` (antes 53.283.379,41); las 9 filas del falso hueco se escriben (ASE2 `E78/E79/E80 = 1.616.173,63 / 406,42 / 1.616.580,05`; ASE4 `E313/E314/E315 = 20.152.773,11 / -2.377,86 / 20.150.395,25`).
- **Gate `C14`:** `VALIDACION_TOTAL!C14 = 72.497.949.967` y `C15 = 1` (sellado por read-back).
- **Deltas conocidos/aceptados (no contenido de negocio):** R1 492 formulas/682-702 (manual, NO adoptadas); BCE V5 (`E7/D7` = 2 celdas); `CONSOLIDADO!D6` metadato; `Recaudo *` Q1 acumulada (V3); `DetRetri/DetValiRetri` columnas excluidas; `ANT EXT-REV` protegida.
- **Deltas del ORACULO a declarar (no brechas de la App):** (i) `Valida - Control Recaudo!E7/E10/F10`: la App preserva las **formulas** de la v2, el manual las **convirtio a valores** (los valores coinciden: 68.213.380.915 / 4.284.569.052 / 72.497.949.967) y añadio `F9=1` (bandera de control manual); (ii) `CONSOLIDADO!Y111` formula vs valor; (iii) `VALIDACION_EAAB-CL` 5 formulas con columna `D` (Q1) en el manual vs `F` (Q2) de la App; (iv) `Informe AFaseo` (hoja que la App no escribe). Ninguna es contenido derivable de los insumos.
- **`<f>`:** fresh vs `p47-h2c` → `fDiff` solo en R2 (352), CONSOLIDADO (61), DetRetri (40), DetValiRetri (51), REMUNERACION_* (56-60) y `Valida - Control Recaudo` (3); **cero** en el resto. La App no crea ni modifica formulas fuera de la geometria ya congelada.

### 9.4 Puertas

- **H1→H2 (S2.1..S2.5): CUMPLIDA.** Re-congelamiento con T0 citado (S2.1); pipeline sobre v2 con fail-fast accionable corregido y run `codigo=0` (S2.2); frente R1/46 y R2/R4/47 intactos salvo geometria v2 + assert 649 (S2.3); `<f>` intacta, guard anti-formula y decimal exacto (S2.4); build 0 warnings + greps (S2.5).
- **H2→H3:** suite, Q1, SHAs, build y cero commits (ver §9.5). **Code-review NO lanzado** (lo comisiona el Ingeniero).

### 9.5 Gates observados

- **Build:** 0 warnings / 0 errores.
- **Suite H1b:** `p48-h1b-suite.trx` = **1059 Total / 1021 Passed / 38 Failed / 0 Skipped**. Set-diff vs `o43-final-failed.txt` (39) = **0 altas / 1 baja heredada** (`ProcesadorPeriodoTests.Ejecutar_Periodo2026072_ConLectorOracle_VeredictosValidacionesNoVacios`, causa Plan 46); vs `p47-warnings.trx` (38) = 0 altas / 0 bajas.
- **Q1:** `SeccionEspejoR1Tests.T1_Ase2Q1_NuevoEsquemaEnPrimeraSeccion_ErrSinSlotSinFallback` → **Passed** (mismo set base).
- **SHAs:** v1 `3086AC31…C5214` y v2 `FF5CF7…2632` intactos.
- **Proceso:** cero commits (`HEAD = c438c8c`), CRLF, cero `InsertRow`/`DeleteRow`, D-H (la tabla de etiquetas de control es DATO citado, no rama por empresa).

### 9.6 H1b — cierre de las 7 altas del STOP (2026-10-19)

La suite H1 (`p48-h1-suite.trx`, 1058/1013/45/0) dejó **7 altas** vs la base de 38. El Ingeniero
autorizó H1b con este alcance; todas quedaron cerradas.

| # | Rojo | Causa raíz (byte-probada) | Cambio (archivo:linea) |
|---|---|---|---|
| 1 | `ValidaControlRecaudoF10Tests.JulioQ2/AgostoQ2` (F10=0) | En la v2 `F10=E7+E10` es **fórmula**; el writer escribía `D7:D10` pero **no sellaba** los cachés de `E7/E10/F10`. | `OpenXmlPlantillaWriter.cs` → `SellarCachesControlRecaudo` tras `EscribirControlRecaudoTotal` en **ambos** `GenerarWorkbook` (single ~L160 y multi ~L334); sellado por read-back, jamás inventado. |
| 2 | `ValidaControlRecaudoF10Tests.F10_DestinoFormula_…` | El test anclaba el guard a `F10`, que en la v2 es fórmula legítima (la App no la escribe). | Re-basado a `F10_DestinoFormula_NoSeSobrescribe_YSelladoEvaluado`: (a) el texto `<f>` mutado se conserva; (b) el sellado deja el valor evaluado. Se añadió `D7_DestinoFormula_LanzaErrPlantilla` para conservar el guard anti-fórmula en el destino que sí se escribe. |
| 3 | `BasePeriodoPreflightTests.S5(2026072)` → `E155` "valor fijo" | **No había ningún writer de producción escribiendo `E155`** (grep: solo aparece en las listas protegidas). El test corría el pipeline v2 contra la **plantilla v1**: en v1 la fila 155 no existe/está vacía y el guard de la cadena v2 (`E155`) la veía como "valor fijo". En v2 `E155=E101+E123-K101` (fórmula). | `BasePeriodoPreflightTests`: `PlantillaUnica()` → v2 (base vigente); `S5` parametriza **Q2→v2 / Q1→v1 histórica**; assert nuevo de que el cuadro OPORTUNO conserva sus fórmulas tras la corrida. |
| 4 | `DivergenciaDeclaradaT0Tests.T0_Oficial_Sha256Intacto` | El test pedía el SHA v1 contra la v2 (ahora oficial). | Verifica **AMBOS**: v1 histórica `3086AC31…C5214` y v2 `FF5CF7…2632`. |
| 5 | `DivergenciaDeclaradaT0Tests.T0_CorridaVivaQ2_…` | Igual: SHA v1 anclado. El escenario ya es **sintético** (Plan 46/47). | SHA v2; escenario sintético intacto. |
| 6 | `DivergenciaDeclaradaCierreT4Tests.T4_NoDeclaradaInyectada…` | El decorador corrompía el **caché** de `F10`; en v2 `F10` es fórmula y el evaluador la resuelve ignorando el caché (correcto), así que la inyección no producía divergencia. | Inyecta el **literal** `D7` (control de recaudo que alimenta `E7→F10→C14`): la divergencia NO declarada vuelve a bloquear. |
| 7 | `BasePeriodoPreflightTests.S5(2026071)` | El caso Q1 quedó corriendo con v2 al cambiar `PlantillaUnica()`; Q1 está fuera de alcance y su rojo es **heredado** de la base 38. | Q1 corre con la **v1 histórica** (dato del test, no producción); el caso vuelve a su rojo base (sin alta nueva). |

**Evidencia H1b:** corrida fresca v2 `codigo=0`; `F10 = 72.497.949.967` (julio-2) y `77.064.563.298`
(agosto-2); `VALIDACION_TOTAL!C14 = 72.497.949.967`, `C15 = 1`; cuadro OPORTUNO con fórmulas intactas
(`E47=E21+E32-K21`, `E155=E101+E123-K101`, `E273`, `E384`, `E468`).

### 9.7 H3 — cierre documental (2026-10-19)

**Observaciones del review APROBADO-CON-OBSERVACIONES — CERRADAS:**

- **W-1 (Regla #8):** `Remuneracion.IntegrationTests/AdopcionPlantillaV2T0Plan48Tests.cs` normalizado
  a **CRLF** (`CR=914 / LF=914 / LF_solo=0`), contenido intacto.
- **S-1 (punto único del sellado):** el trío `{E7, E10, F10}` es **DATO citado** en
  `WorkbookLeafCellMapValidaciones.CeldasSelladoControlRecaudo`; `SellarCachesControlRecaudo` lo
  consume (se retiró el arreglo inline). Cero cambio de comportamiento.
- **S-2 (SHA v2):** corregido a `FF5CF7A819110A1E1E2CA7BDAD553DD7A25141ADD472B1904BD496F082642632`
  en front-matter y §0.2 (coincide con `Docs/T0-Plan48…` y los fixtures).
- **S-3 (punto único del CONSOLIDADO):** `WorkbookLeafCellMapQ2.ConsolidadoProtected` es la única
  fuente; `ValidarFormulasProtegidasMultiAse` consume ese mapa y se retiró la copia inline del
  writer (`WorkbookLeafCellMapPorAseProtectedConsolidado`). Cero cambio de comportamiento.

**Docs H3 tocados:** `.opencode/project-context.md` (doctrina v2), `plans/44` (§8 R-PLANTILLA +
paridad), `plans/46`/`plans/47` (nota de cierre), `Docs/Instructivo-Capa-B.md` (R2/R4-detalle v2 +
control de recaudo), `plans/48` (§9.7), `Docs/T0-Plan48-Adopcion-Plantilla-v2.md` (§9.6).

**Riesgos residuales declarados:**
1. El fixture `AdopcionPlantillaV2T0Plan48Tests` depende de `%TEMP%\opencode\p47-h2c`; en máquina
   limpia falla con mensaje accionable (dependencia declarada).
2. `Insumos.BaseDelPeriodo` deriva **todos** los períodos de la v2 y la propietaria solo auditó
   julio-2 (**P5c RESUELTA 2026-10-19: la v2 es la referencia para TODOS los periodos; v1 historica; la cobertura de evidencia sigue limitada a julio-2**)` — nota: la ratificacion deja abierto el frente Q1 (procesar una Q1 con la v2 hoy falla con `ERR-SIN-SLOT` en R1 ASE2 `||NUEVO ESQUEMA|Total|`), que requiere autorizacion separada..
3. `Remuneracion.IntegrationTests/TestResults/p48-setdiff.txt` es **scratch**; la evidencia
   vinculante es el set-diff de la suite contra `o43-final-failed.txt` (abajo). No se cita como gate.
4. Acoplamiento a geometría/rótulos v2: una eventual **v3** exigirá re-congelar por T0 con fail-fast
   visible (`ERR-PLANTILLA`), nunca edición silenciosa.

**Estado final: CERRADO 2026-10-19.** H0 (T0 21/21) · H1 (R2 v2 + control de recaudo; 492 fórmulas
del manual excluidas) · H1b (7 altas cerradas) · H2 (`codigo=0`; R2 `litVsLit=0`; `F10/C14/C15` OK) ·
H3 (docs). Review APROBADO-CON-OBSERVACIONES → observaciones cerradas.
Suite `p48-h3-suite.trx`: **1059 Total / 1021 Passed / 38 Failed / 0 Skipped**; set-diff vs
`o43-final-failed.txt` (39): **0 altas / 1 baja heredada**
(`ProcesadorPeriodoTests.Ejecutar_Periodo2026072_ConLectorOracle_VeredictosValidacionesNoVacios`);
vs `p47-warnings.trx` (38): 0 altas / 0 bajas. SHAs v1/v2 intactos; cero commits.

#### Addendum 2026-10-19 — el falso verde de julio-2 queda CERRADO por el Plan 49 (no se reescribe la historia)

El Plan 49 (`plans/49 - cerrar-falso-verde-gate-2.7.md`, APROBADO 2026-10-19) probó con bytes
(`Docs/T0-EXTEMP-Ase5-Plan49.md`) que este cierre convivía con un **falso verde**: la corrida
promovida de julio-2 tiene `CONSOLIDADO_TOTAL RECAUDO!D51 = 0` frente a `441.849` del manual, y el
gate `D34/F34/O25` **no lo veía** porque la pérdida del EXTEMPORÁNEO ASE5 entraba en **ambos**
operandos y cancelaba. La clasificación de §9.7 ("las 9 fórmulas faltantes de 682/692/702 son delta
del manual") queda **incorrecta**: esas celdas son **anclas consumidoras** (`CONSOLIDADO!D51`,
`DetRetri!D36`, `DetValiRetri!D28`, `REMUNERACION_OCCIDENTE!D51`), no cosmética.

Qué cambió (solo Plan 49, este documento no se reescribe):

- Gate **absoluto** `ValidadorBasico.ValidarExtempAbsolutoPorAse` (núcleo workbook-vs-dominio +
  interior R1 `F682/F702` vs secciones), evidencia en `ValidacionOracleReader`
  (`ExtempResuelto`/`AnclasExtempR1`). Corre en Q1 y Q2 sin exención.
- Los 5 T0 de §8.1 del T0-Plan49 quedan **DEUDA DECLARADA** (región, asserts byte-idénticos);
  retiro cuando llegue la **v3** de la propietaria con las fórmulas restauradas.
- La suite pasa a `p49-h2-suite.trx`: **1061/1012/49/0** con **11 altas, 0 bajas** (todas `Plan 49
  A1` sobre julio-2). Esas altas son el **rojo honesto** del plan; **prohibido reverdecarlas**.
- La corrección de plantilla (EXTEMP ASE5 y R2 ASE4 `12.462,26`) sigue **FUERA**: depende de la v3
  y de un T0 de dueño propio.
