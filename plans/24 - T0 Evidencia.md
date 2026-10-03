# Plan 24 — T0 Evidencia (Fase 0, blocker agosto en el mapa leaf R1-Q2 de ASE3)

> Documento de evidencia del Discovery T0 del blocker de agosto del mapa leaf R1-Q2
> (`WorkbookLeafCellMapQ2.R1Q2EditablesPorAse[3]`), descubierto al superar el blocker 2.5 (HU-22,
> Plan 23). **Solo evidencia: cero cambios de código productivo o de tests.** Toda afirmación remite
> a disco (archivo/hoja/fila). Replica el formato del `plans/21 - T0 Evidencia.md` y del
> `plans/22 - T0 Evidencia.md`.
>
> **Fecha:** 2026-10-02
> **Origen del blocker:** `Remuneracion.IntegrationTests/Regresion2026082Tests.cs` (L37-L42, L88-L92)
> — al quedar resuelto el blocker de SALDOS POR NOTA (HU-22/Plan 23, `Vlr Intereses` opcional), el
> flujo 5-ASE de agosto avanza y muere en `ExcelDataReaderWorkbookLeafInputReader.MapearR1Q2`
> (`Readers`→`LeerLeafInputs`) con `CalculoInvalidoException("ASE 3: la fuente R1-Q2 no trae la fila
> del rol ...")` ANTES de escribir.
> **Insumos Q1-julio:** `Docs/Insumos/REMUNERACION 2026071/{1..5}-*/Recaudoporcomponente_*0107*_15072026*.xlsx`.
> **Insumos Q2-julio:** `Docs/Insumos/REMUNERACION 2026072/{1..5}-*/Recaudoporcomponente_*1607*_31072026*.xlsx`.
> **Insumos agosto:** `Docs/Prueba2/Insumos/{1..5}-*/Recaudoporcomponente_*1608*_31082026*.xlsx`.
> **Oráculo DetRetri (R10):** `Docs/Prueba2/Insumos/R10_Remuneracion_2026082.xlsx` (hoja
> `DetRetri2026082`, col D = `RECAUDO TOTAL`, filas 9..13 = ASE 1..5), leído con `mcp-excel`.
> **Goldens/destinos:** `%TEMP%\opencode\Remuneracion 202607-1 Total.xlsx` y
> `...-2 Total.xlsx` (cachés; el original en `Docs` puede estar tomado por Excel), y la plantilla
> canónica `Docs/Insumos/REMUNERACION 2026072/Plantilla_ Remuneracion 202607-2.xlsx`.

---

## 0. Conclusión de una línea

La causa raíz NO es "faltó marcar `Aplicacion nuevos x reversion` como opcional": el mapa R1-Q2
tiene **mal el rol** de un operando EXTEMP de ASE1 y ASE3 (asigna la fila `Subsidio` donde la celda
destino es una fila `Aplicacion`), y ese error quedó **enmascarado por una coincidencia numérica de
julio** (`Subs0 == Aplic1`). El oráculo R10 de agosto lo prueba: el EXTEMP correcto de ASE3 es
**0** (`R10 − base = 0.33`), mientras que la resolución actual produciría `Subs0 = −1.087.996.443,09`
(un desvío de ~1.088e9 en el DetRetri). La regla que casa con el oráculo en 10/10 (5 ASE × {julio-Q2,
agosto}) es **`EXTEMP = Aplic0 + Aplic1`** (las dos filas `Aplicacion`, ausente = 0).

---

## 1. Causa raíz de agosto (evidencia dura)

El fail-fast de 2026082 es **estructural y reproducible desde disco**:

| Período | ASE3 fuente (`Recaudoporcomponente_*`) | filas `Mes/Total` | filas `Aplicacion nuevos x reversion` | Resultado |
|---|---|---|---|---|
| Q2-julio | `...3-Ciudad Limpia\...16072026...` | 3 (R8, R29, R45) | **2** (R14, R34) | OK (mapa casa) |
| agosto | `...3-Ciudad Limpia\...16082026...` | 3 (R11, R27, R39) | **0** | **fail-fast antes de escribir** |

Disparador exacto (`ExcelDataReaderWorkbookLeafInputReader.cs`, `MapearR1Q2`, L1009-L1016): el bucle
recorre `R1Q2EditablesPorAse[3]` y, al no existir `filasAplicacion.ElementAtOrDefault(0)`, lanza:

```text
ASE 3: la fuente R1-Q2 no trae la fila del rol Aplic0 requerida por el mapa T0 para la celda F250 (reporte Recaudoporcomponente).
```

El test `Regresion2026082Tests` (L91-L92) lo captura por nombre (`"la fuente R1-Q2 no trae la fila del
rol"` + `"ASE 3"`). Reproducción del dump crudo de la fuente breaker (Sheet1, columnas A–E + F):
ver §2.4.

**Hallazgo clave (va más allá del "slot ausente"):** la celda `F270` de ASE3 está bindeada en el mapa
a `R1Q2Fuente.Subs0` (`WorkbookLeafCellMapQ2.cs` L79: `("F270", R1Q2Fuente.Subs0)`), pero en el
destino esa fila es una fila `Aplicacion nuevos x reversion` (§4.4). En julio `Subs0` (primera fila
`Subsidio`) coincide numéricamente con `Aplic1` (segunda fila `Aplicacion`), por eso el golden no lo
delató; en agosto, con las filas `Aplicacion` suprimidas pero la fila `Subsidio` presente, la
coincidencia se rompe (§3.3, §3.4).

---

## 2. T0a — Dump por ASE × período de los roles que exige `R1Q2EditablesPorAse`

Método: lectura OpenXML del `Sheet1` de cada `Recaudoporcomponente_*.xlsx`; los roles se resuelven
con las MISMAS firmas del reader (`filasMes`: B=`Mes`∧C=`Total`; `filasAplicacion`: B contiene
`Aplicacion nuevos x reversion`∧C=`Total`; `filasSubsidio`: E contiene `Subsidio(-)/Contribucion(+)`);
la columna `SERVICIO ESPECIALES` se localiza por su encabezado.

### 2.1 Inventario por período × ASE (fuente real)

| Período | ASE | Mes/Total | Aplicacion | Subsidio | Especiales (L) | Roles que exige el mapa | Faltantes |
|---|---|---:|---:|---:|:--:|---|---|
| Q1-julio | 1 | 2 | 2 | 5 | L | Q1 usa `MapearR1` (rol distinto) | n/a (Q1 no usa este mapa) |
| Q1-julio | 2 | 3 | **0** | 6 | L | Q1 usa `MapearR1` (EXTEMP estático 0) | n/a |
| Q1-julio | 3 | 3 | 2 | 5 | L | Q1 usa `MapearR1` | n/a |
| Q1-julio | 4 | 3 | 2 | 9 | L | Q1 usa `MapearR1` | n/a |
| Q1-julio | 5 | 3 | **0** | 5 | L | Q1 usa `MapearR1` (EXTEMP estático 0) | n/a |
| Q2-julio | 1 | 3 | 2 | 5 | L | Mes0/1/2, Lmes0/1, Subs0, Aplic0, LAplic0 | ninguno |
| Q2-julio | 2 | 3 | 2 | 10 | L | Mes0/1/2, Lmes0/1, Aplic0, Aplic1, LAplic0 | ninguno |
| Q2-julio | 3 | 3 | 2 | 5 | L | Mes0/1/2, Lmes0/1, Subs0, Aplic0, LAplic0 | ninguno |
| Q2-julio | 4 | 3 | 2 | 9 | L | Mes0/1/2, Lmes0/1, Aplic0, Aplic1, LAplic0 | ninguno |
| Q2-julio | 5 | **2** | 2 | 7 | L | Mes0/1, Lmes0, Aplic0, Aplic1, LAplic0 | ninguno |
| agosto | 1 | 3 | 2 | 5 | L | idem Q2 | ninguno |
| agosto | 2 | 3 | 2 | 8 | L | idem Q2 | ninguno |
| agosto | 3 | 3 | **0** | 4 | L | idem Q2 | **Aplic0, Aplic1, LAplic0** |
| agosto | 4 | 3 | 2 | 10 | L | idem Q2 | ninguno |
| agosto | 5 | 3 | 2 | 8 | L | idem Q2 | ninguno |

**Lectura T0a:** en Q2 el mapa solo se ejercita en la 2.ª quincena (`MapearR1Q2`,
`ExcelDataReaderWorkbookLeafInputReader.cs` L48). En Q2-julio **5/5 ASE** traen todas las filas
`Aplicacion` (por eso el mapa nunca falló). En agosto **4/5 ASE** las traen; la excepción es
**únicamente ASE3**. La versión Q1 del mapa (`MapearR1`) no aplica (Q1 no usa este path).

### 2.2 Firma real de los roles (definición operativa del reader)

| Rol (`R1Q2Fuente`) | Firma en la fuente | Columna de valor |
|---|---|---|
| `Mes0/Mes1/Mes2` | B=`Mes` y C=`Total`, por orden de aparición (0-based) | F (index 5) |
| `Lmes0/Lmes1/Lmes2` | fila `Mes` de la ocurrencia correspondiente | columna `SERVICIO ESPECIALES` (L) |
| `Subs0` | primera fila E contiene `Subsidio(-)/Contribucion(+)` | F |
| `Aplic0/Aplic1` | B contiene `Aplicacion nuevos x reversion` y C=`Total`, por orden | F |
| `LAplic0` | fila `Aplic0` | columna `SERVICIO ESPECIALES` (L) |

### 2.3 Requerimientos congelados por ASE (`WorkbookLeafCellMapQ2.cs` L57-L94)

| ASE | celdas ← rol (mapa actual) | Observación |
|---|---|---|
| 1 | F12←Mes0, L12←Lmes0, F32←Mes1, L32←Lmes1, F48←Mes2, **F37←Subs0**, F17←Aplic0, L17←LAplic0 | F37 destino es fila `Aplicacion` (§4.4) |
| 2 | F94←Mes0, L94←Lmes0, F132←Mes1, L132←Lmes1, F160←Mes2, F142←Aplic1, F107←Aplic0, L107←LAplic0 | consistente |
| 3 | F244←Mes0, L244←Lmes0, F265←Mes1, L265←Lmes1, F281←Mes2, **F270←Subs0**, F250←Aplic0, L250←LAplic0 | F270 destino es fila `Aplicacion` (§4.4) |
| 4 | F384←Mes0, L384←Lmes0, F422←Mes1, L422←Lmes1, F448←Mes2, F430←Aplic1, F397←Aplic0, L397←LAplic0 | consistente |
| 5 | F531←Mes0, L531←Lmes0, F552←Mes1, F538←Aplic1, F512←Aplic0, L512←LAplic0 | consistente (sin Mes2) |

### 2.4 Dump crudo de la fuente breaker — ASE3 agosto (`Sheet1`, A–E + F)

Archivo: `Docs\Prueba2\Insumos\3-Ciudad Limpia\Recaudoporcomponente_to_date16082026ddMMyyyy_to_date31082026ddMMyyyy___202691194245429.xlsx`

```text
R1  | A=Recaudo Desde: 16/08/2026 Hasta: 31/08/2026
R4  | (headers) ... E=Total ... F=Total ... L=SERVICIO ESPECIALES
R5  | E=Vlr Servicio                             F=-3933.31
R6  | D=E    E=Total                             F=-3933.31
R7  | C=ENEL D=Total                             F=-3933.31
R8  | E=Vlr Servicio                             F=-522.8
R9  | D=O    E=Total                             F=-522.8
R10 | C=OCCIDENTE D=Total                        F=-522.8
R11 | B=Mes  C=Total                             F=-4456.11      <- Mes0 (F244)
R12 | A=AFaseo B=Total                           F=-4456.11
R13 | E=Vlr Servicio                             F=16203110832.58
R14 | E=Vlr Intereses                            F=9783329.82
R15 | D=E    E=Total                             F=16212894162.4
R16 | C=ENEL D=Total                             F=16212894162.4
R17 | E=Vlr Servicio                             F=56129769.21
R18 | E=Vlr Intereses                            F=445748.01
R19 | D=H    E=Total                             F=56575517.22
R20 | E=Vlr Servicio                             F=69786246.86
R21 | E=Vlr Intereses                            F=-1435565.73
R22 | D=O    E=Total                             F=68350681.13
R23 | E=Vlr Servicio                             F=754857948.81
R24 | E=Vlr Intereses                            F=17926857.18
R25 | D=T    E=Total                             F=772784805.99
R26 | C=OCCIDENTE D=Total                        F=897711004.34
R27 | B=Mes  C=Total                             F=17110605166.74 <- Mes1 (F265)
R28 | A=Componente B=Total                       F=17110605166.74
R29 | E=Subsidio(-)/Contribucion(+)              F=-1087996443.09 <- Subs0
R30 | D=E    E=Total                             F=-1087996443.09
R31 | C=ENEL D=Total                             F=-1087996443.09
R32 | E=Subsidio(-)/Contribucion(+)              F=2327352.78
R33 | D=H    E=Total                             F=2327352.78
R34 | E=Subsidio(-)/Contribucion(+)              F=30014496.44
R35 | D=O    E=Total                             F=30014496.44
R36 | E=Subsidio(-)/Contribucion(+)              F=328146742.01
R37 | D=T    E=Total                             F=328146742.01
R38 | C=OCCIDENTE D=Total                        F=360488591.23
R39 | B=Mes  C=Total                             F=-727507851.86 <- Mes2 (F281)
R40 | A=Subs/Cont B=Total                        F=-727507851.86
R41 | A=Total                                    F=16383092858.77
```

> **Lectura:** las filas `B=Mes C=Total` son R11, R27 y R39 (**3** filas) → Mes0/Mes1/Mes2, que es
> exactamente lo que resuelve `filasMes` (L970-L973). La fila `A=AFaseo B=Total` (R12) NO es `Mes` y
> no entra al filtro. Las filas `Aplicacion nuevos x reversion` (Aplic0/Aplic1) NO existen en este
> archivo (0 filas): ése es el fail-fast. Nota: `sharedStrings` de esta fuente se verificó sin la
> cadena `"Aplicacion nuevos"` (memoria discovery/r1q2-ase3-agosto-blocker).

---

## 3. T0b — Semántica de los roles ausentes (con datos)

### 3.1 Qué falta en agosto, por ASE

Solo **ASE3** pierde filas `Aplicacion nuevos x reversion` en agosto (0 filas). ASE1, ASE2, ASE4 y
ASE5 traen las 2 esperadas. Por tanto la ausencia **NO es generalizada** en Q2: es **específica de
ASE3** en agosto. (En Q1 sí es común — ASE2/ASE5 sin `Aplicacion` — pero Q1 usa `MapearR1`, con
EXTEMP estático 0 para esos dos ASE; no ejerce este mapa.)

### 3.2 Qué es la fila `Aplicacion nuevos x reversion` (semántica deducida de la estructura)

En las fuentes, la fila `Aplicacion nuevos x reversion` es el **subtotal del sub-bloque de
componentes inmediatamente anterior**:

| Período | ASE | `Aplicacion` (F) | Sub-bloque previo | ¿coincide con la fila `Total` previa? |
|---|---|---|---|---|
| Q2-julio | 2 | R24 = 30.393.431,63 | R13..R23 (ENEL 30.144.259,37 + O 189.277,37 + T 59.894,89) | **suma**, no eco de una sola fila |
| Q2-julio | 4 | R24 = 4.600.575,16 | bloque multi-compañía | suma |
| Q2-julio | 5 | R14 = 417.519,62 | bloque multi-compañía | suma |
| Q1/Q2/ago | 1 | Aplic0/Aplic1 = (R9,R29) / (R16,R36) / (R13,R33) | bloque monocompañía (ENEL u OCCIDENTE) | **eco** de la fila `Total` previa |
| Q1/Q2 | 3 | Aplic0/Aplic1 = (R17,R37) / (R14,R34) | bloque monocompañía | **eco** de la fila `Total` previa |

Verificación programática (`%TEMP%\opencode\p24echo.ps1`, `Aplicacion-echo: 11/24`): coinciden las 6
filas de ASE1, las 4 de ASE3 y 1 de ASE2-agosto (R50); NO coinciden las de ASE2/ASE4/ASE5 (porque son
subtotales de bloques multi-compañía). Conclusión: `Aplicacion` es un **subtotal real**; en ASE1/ASE3,
por ser bloques monocompañía, coincide numéricamente con la fila `Total` previa —y, en el bloque de
`Subsidio`, con la primera fila `Subsidio` (`Subs0`).

### 3.3 La coincidencia que enmascaró el defecto (`Subs0 == Aplic1`)

| Período | ASE | `Subs0` (F) | `Aplic1` (F) | ¿Iguales? |
|---|---|---:|---:|:--:|
| Q1-julio | 1 | 5.341.504,63 (R26) | 5.341.504,63 (R29) | sí |
| Q1-julio | 3 | 4.499.898,22 (R34) | 4.499.898,22 (R37) | sí |
| Q2-julio | 1 | 40.364,31 (R33) | 40.364,31 (R36) | sí |
| Q2-julio | 3 | 627.141,97 (R31) | 627.141,97 (R34) | sí |
| agosto | 1 | 21.200,58 (R30) | 21.200,58 (R33) | sí |
| agosto | 3 | **−1.087.996.443,09 (R29)** | **AUSENTE (0 filas)** | **NO** |

Por eso el mapa de ASE1/ASE3 usa `Subs0` y "funcionó" en julio: `Subs0 + Aplic0 == Aplic1 + Aplic0`.
En agosto ASE3 la igualdad se rompe (Aplic1 no existe, Subs0 sí).

### 3.4 Semántica de la ausencia: 0 legítimo, NO valor a recuperar

El oráculo R10 de agosto (`DetRetri2026082`, col D `RECAUDO TOTAL`) fija el DetRetri esperado de
ASE3 = **16.369.059.896**. La composición del DetRetri está congelada
(`ProcesadorPeriodo.cs` L223-L231): `TotalD104 = R1.TotalOportunoEsperadoPorAse + R2.TotalOportunoEsperado
+ R1.ExtemporaneoEsperadoPorAse + R4.TotalReversionEsperada + AjustesSfT.TotalAjustes`, y
`DetRetri = ROUND(TotalD104)`. Despejando el EXTEMP:

`extemp_implícito = R10 − (totOpt + R2 + R4 + Ajustes)`.

Verificación completa (10 combinaciones; método en §9, script `p24oracle.ps1`).

| Período | ASE | R10 (RECAUDO TOTAL) | base = totOpt+R2+R4+Ajustes | extemp implícito | `Subs0+Aplic0` (dif) | **`Aplic1+Aplic0`** (dif) |
|---|---|---:|---:|---:|---:|---:|
| Q2-julio | 1 | 17.450.228.673 | 17.450.179.093,35 | 49.579,65 | 49.580 (−0,35) | **49.580 (−0,35)** |
| Q2-julio | 2 | 20.516.143.970 | 20.459.953.899,65 | 56.190.070,35 | 56.328.032,26 (−137.961,91) | **56.190.070 (+0,35)** |
| Q2-julio | 3 | 15.221.896.467 | 15.220.588.787,45 | 1.307.679,55 | 1.307.680 (−0,45) | **1.307.680 (−0,45)** |
| Q2-julio | 4 | 7.179.595.396 | 7.175.252.294,67 | 4.343.101,33 | 4.374.297,09 (−31.195,76) | **4.343.101 (+0,33)** |
| Q2-julio | 5 | 12.073.344.662 | 12.072.902.812,59 | 441.849,41 | 409.613,46 (+32.235,95) | **441.849 (+0,41)** |
| agosto | 1 | 18.378.829.331 | 18.378.765.761,13 | 63.569,87 | 63.570 (−0,13) | **63.570 (−0,13)** |
| agosto | 2 | 21.599.709.648 | 21.599.488.017,53 | 221.630,47 | 221.630 (+0,47) | **221.630 (+0,47)** |
| agosto | 3 | 16.369.059.896 | 16.369.059.895,67 | **0,33** | −1.087.996.443,09 (+1.087.996.443,42) | **0 (+0,33)** |
| agosto | 4 | 8.372.092.112 | 8.368.061.702,47 | 4.030.409,53 | 4.078.364,25 (−47.954,72) | **4.030.410 (−0,47)** |
| agosto | 5 | 12.137.660.178 | 12.136.410.368,17 | 1.249.809,83 | 1.455.151,60 (−205.341,77) | **1.249.810 (−0,17)** |

**Veredicto T0b:**
- La ausencia de `Aplicacion` en ASE3-agosto significa **EXTEMP = 0 para esas filas** (0 legítimo,
  NO fila que jamás puede faltar): el oráculo R10 cierra con extemp implícito = **0,33** (±0,5).
- La resolución correcta es **`EXTEMP = Aplic0 + Aplic1`** (suma de las filas `Aplicacion`;
  ausentes ⇒ 0). Casa con el oráculo en **10/10** combinaciones.
- La resolución actual (`Subs0 + Aplic0`) casa solo donde `Subs0 == Aplic1` por coincidencia
  (ASE1, ASE3-julio); **falla en ASE2/ASE4/ASE5 y en ASE3-agosto**.
- **Hacer `Aplic0` opcional (ausente=0) SIN corregir el rol `Subs0` de `F270/F37` NO arregla la
  aritmética**: ASE3-agosto daría extemp = −1.087.996.443,09 y el DetRetri escrito quedaría
  1.087.996.443,42 por debajo del R10 → `ERR-VALIDACION DetRetri-vs-R10` (no llega a escribir).

---

## 4. T0c — Impacto en celdas destino

### 4.1 Consumidores reales de `MapearR1Q2`

| Consumidor | Archivo:línea | Qué usa | ¿Depende de `Aplic`? |
|---|---|---|---|
| DetRetri Q2 (escritura) | `ProcesadorPeriodo.cs` L223-L231 | `TotalD104 = totOpt + R2 + extemp + R4 + Ajustes` | **SÍ** (`ExtemporaneoEsperadoPorAse`) |
| Gate `Σ empresas = visible bloque` | `WorkbookLeafCoherence.cs` L101-L104 | `TotalOportunoEsperadoPorAse` (totOpt) | no (solo Mes/Lmes) |
| Gate `F25 vs Extemporáneo HU-02` | `WorkbookLeafCoherence.cs` L33-L36 y L182-L185 | `F25` (= Mes0) | no |
| Escritura celdas leaf R1 (single-ASE) | `OpenXmlPlantillaWriter.cs` L1019-L1030 | `R1.CeldasPorAse` (incl. F250/F270) | **SÍ** (pero ver §4.2) |
| UI/log | `Form1.cs` L545, L575, L797-L799 | `ExtemporaneoEsperado`/`…PorAse` | SÍ (informativo) |

### 4.2 El espejo R1 supersede la escritura R1 del mapa leaf

En el flujo 5-ASE, `GenerarWorkbook` llama `EscribirCeldasLeafPorAse(..., omitirR1: espejoR1)`
(`OpenXmlPlantillaWriter.cs` L267-L274) y `espejoR1 = bloquesEspejo.Count == 5` (L218-L223). Con
`omitirR1 = true`, el bloque R1 del mapa leaf **no se escribe** (L1019); el espejo
(`OpenXmlEspejoR1Mutador`) reescribe la secuencia completa por ENCABEZADO (L294-L331). Es decir, en el
flujo 5-ASE las celdas `F250/F270/...` del R1 NO dependen del mapa leaf para su valor. **Pero
`ExtemporaneoEsperadoPorAse` sí se sigue consumiendo** (DetRetri) y, aguas abajo, el visible
`F345 = F270+F250−L250` (golden, §4.4) alimenta `CONSOLIDADO D49`. En el flujo single-ASE
(`ProcesadorRemuneracion.cs` L77-L90, sin `EspejoR1`), el mapa leaf SÍ escribe las celdas R1.

### 4.3 Conteo `Mes/Total` de ASE3-agosto (verificado, sin riesgo)

`filasMes` (B=`Mes` ∧ C=`Total`) en ASE3-agosto = R11, R27, R39 (**3** filas; la fila
`A=AFaseo B=Total` NO matchea). El mapa pide Mes0/Mes1/Mes2 → R11/R27/R39, que es lo que resuelve el
reader, y `Lmes0`/`Lmes1` = `SERVICIO ESPECIALES` de R11 (0) y R27 (37.341.348,91). Resultado:

`totOpt = F(R11) + F(R27) + F(R39) − L(R11) − L(R27) = −4.456,11 + 17.110.605.166,74 − 727.507.851,86 − 0 − 37.341.348,91 = 16.345.751.509,86`.

Es exactamente el `totOpt` que usa la verificación de §3.4 y cierra contra el oráculo R10. **Sin
riesgo: el conteo de `Mes` no es un defecto adicional.**

### 4.4 Celdas visibles protegidas y su fórmula (golden Q2, bloque ASE3)

| Celda | Fórmula (golden) | Valor golden | Fragmentos congelados (`R1Q2ProtectedPorAse[3]`) |
|---|---|---:|---|
| R1!F343 | `F265+F281+F244-L244-L265` | 15.226.209.207,13 | `["F265","F281","F244","L244","L265"]` |
| R1!F345 | `F270+F250-L250` | 1.307.680 | `["F270","F250","L250"]` |
| R1!F344 | `F281` | −692.112.419,66 | (no listado en el mapa) |

`F345` referencia exactamente las celdas que el mapa bindea (`F270`, `F250`, `L250`): confirma que el
rol de `F270` debe ser una fila `Aplicacion` (la fila destino R270 del golden es
`Aplicacion nuevos x reversion`; ver §3.2). En agosto, `espejoDesplazado = true` ⇒ la validación de
protegidas (`ValidarFormulasProtegidasMultiAseQ2`) NO corre (`OpenXmlPlantillaWriter.cs` L256-L265),
y `F343/F345` se reanclan por el espejo (riesgo R-DELTA-NEGATIVO ya declarado en el Plan 21 §8.1).

### 4.5 Aristas: `L250` (Especiales de `Aplic0`)

La fórmula usa `−L250`; en todas las filas `Aplicacion` observadas el `SERVICIO ESPECIALES` es 0
(`LAplic0 = 0`). Si faltara `Aplic0`, `L250` también debe quedar 0 (no fail-fast), consistente con la
opcionalidad del par.

---

## 5. T0d — Veredicto de diseño

El mapa R1-Q2 leaf por roles congelados es el mismo patrón que ya falló dos veces (L-menores en HU-12,
`Aplicacion` en agosto). Evaluación con datos:

### 5.1 Opción (i) — "rol opcional (ausente = 0), como Plan 23"

Marcar `Aplic0`/`LAplic0` (y la `Aplic1` que faltara) como opcionales, ausente = 0, **manteniendo el
resto del mapa**.

- **Rechazada por datos:** para ASE3-agosto, `F270` sigue bindeado a `Subs0` (presente = −1.087.996.443,09).
  El extemp resultante NO es 0 → el DetRetri se desvía 1.087.996.443,42 del oráculo R10 (§3.4).
- Únicamente marcar `Aplic0` opcional resolvería el fail-fast, pero rompería la validación
  DetRetri-vs-R10 (peor: un error silencioso que emerge en el paso de validación, no en el de lectura).
- La opcionalidad SÍ es necesaria como parte de la solución, pero **no suficiente**.

### 5.2 Opción (ii) — "resolver el rol por etiqueta/firma dinámica dentro de la secuencia espejo"

El `LeerEspejoR1` (Plan 21, `ExcelDataReaderWorkbookLeafInputReader.cs` L537-L623) ya carga la
secuencia del período como `List<FilaEspejoR1>` con A/B/C/D/E y valores por encabezado. Resolver los
roles de la misma firma que hoy usa el reader (`Mes/Total`, `Aplicacion/Total`, `Subsidio`) sobre esa
secuencia, en vez de un mapa celda→rol congelado de ocurrencia.

- **Casa con el oráculo 10/10**: `EXTEMP = Aplic0 + Aplic1` (las filas `Aplicacion`; ausentes = 0).
- Coherente con la doctrina del Plan 21 (D-B/R-E-5: "la secuencia observada ES la especificación").
- Elimina de raíz el defecto de rol (`Subs0` vs `Aplic1`): el valor sale de la MISMA fila que el
  espejo ya escribe.
- **Trade-off:** cambia el modelo de resolución del mapa leaf (de roles congelados a firma+orden).
  Requiere definir el agregado con >2 filas `Aplicacion` (evidencia actual: siempre 2; regla
  propuesta: suma de TODAS las filas `Aplicacion`).

### 5.3 Opción (iii) — corrección mínima del mapa congelado + opcionalidad

Cambiar SOLO los roles erróneos (`ASE1 F37: Subs0→Aplic1`; `ASE3 F270: Subs0→Aplic1`) y marcar las
filas `Aplicacion` como opcionales (ausente = 0), conservando el resto del mapa.

- ✅ Casa con el oráculo 10/10 (idéntico resultado a (ii) en los datos observados).
- ✅ Blast-radius mínimo; no toca la arquitectura de mapas ni el espejo.
- ⚠️ Conserva el patrón "mapa congelado por rol" que ya falló 2 veces: la corrección sigue dependiendo
  de que ninguna coincidencia futura (`Subs0 == Aplic1`) vuelva a enmascarar otro desajuste de rol.
- ⚠️ Requiere declarar la regla de ausencia por rol (obligatorio vs. opcional) en el mapa.

### 5.4 Recomendación T0d

**Recomendación: (ii) resolver por firma/etiqueta sobre la secuencia (con (iii) como plan B de menor
riesgo).** Razones basadas en datos:
1. El defecto es de **rol**, no de ocurrencia; un fix de opcionalidad no lo cubre.
2. La firma `Aplicacion/Total` casa con el oráculo en 10/10 (julio y agosto, 5 ASE), mientras que el
   rol congelado `Subs0` (ASE1/ASE3) casa solo donde `Subs0 == Aplic1` por coincidencia y falla en
   ASE3-agosto.
3. El espejo ya tiene la secuencia; derivar de ella elimina la duplicación de "fuente de verdad"
   (hoy el mapa leaf y el espejo resuelven la misma hoja por caminos distintos, y divergen).
4. El caso de agosto (ausencia real) es exactamente el escenario que la doctrina espejo del Plan 21
   anticipó.

**NO forzar una regla de opcionalidad genérica "ausencia de concepto = 0"** para toda la hoja: los
roles `Mes` (totOpt) y los gates de coherencia exigen el dato; solo las filas `Aplicacion` son las
que varían (evidencia: 0 en ASE3-agosto, 2 en el resto).

---

## 6. T0e — Tabla congelada roles → firmas → celdas (lista para SPEC/TASKS)

### 6.1 Operandos EXTEMP (corregidos) por ASE

| ASE | Rol correcto | Celda destino | Firma en fuente | Obligatoriedad | Regla de ausencia |
|---|---|---|---|---|---|
| 1 | Aplic0 | F17 | B contiene `Aplicacion nuevos x reversion` ∧ C=`Total`, ocurrencia 0 | opcional | 0 explícito |
| 1 | **Aplic1** | **F37** | idem, ocurrencia 1 (hoy el mapa dice `Subs0`) | opcional | 0 explícito |
| 1 | LAplic0 | L17 | `SERVICIO ESPECIALES` de la fila Aplic0 | opcional | 0 |
| 2 | Aplic0 | F107 | idem, ocurrencia 0 | opcional | 0 |
| 2 | Aplic1 | F142 | idem, ocurrencia 1 | opcional | 0 |
| 2 | LAplic0 | L107 | idem | opcional | 0 |
| 3 | Aplic0 | F250 | idem, ocurrencia 0 | opcional | 0 |
| 3 | **Aplic1** | **F270** | idem, ocurrencia 1 (hoy el mapa dice `Subs0`) | opcional | 0 explícito |
| 3 | LAplic0 | L250 | idem | opcional | 0 |
| 4 | Aplic0 | F397 | idem, ocurrencia 0 | opcional | 0 |
| 4 | Aplic1 | F430 | idem, ocurrencia 1 | opcional | 0 |
| 4 | LAplic0 | L397 | idem | opcional | 0 |
| 5 | Aplic0 | F512 | idem, ocurrencia 0 | opcional | 0 |
| 5 | Aplic1 | F538 | idem, ocurrencia 1 | opcional | 0 |
| 5 | LAplic0 | L512 | idem | opcional | 0 |

### 6.2 Operandos TOT_OPT (no dependen de `Aplic`; obligatorios)

| ASE | Roles | Celdas | Firma | Obligatorio |
|---|---|---|---|---|
| 1 | Mes0, Mes1, Mes2, Lmes0, Lmes1 | F12/F32/F48, L12/L32 | B=`Mes`∧C=`Total` (F) y `SERVICIO ESPECIALES` (L) | sí (Mes2 también) |
| 2 | Mes0, Mes1, Mes2, Lmes0, Lmes1 | F94/F132/F160, L94/L132 | idem | sí |
| 3 | Mes0, Mes1, Mes2, Lmes0, Lmes1 | F244/F265/F281, L244/L265 | idem | sí (ver §4.3) |
| 4 | Mes0, Mes1, Mes2, Lmes0, Lmes1 | F384/F422/F448, L384/L422 | idem | sí |
| 5 | Mes0, Mes1, Lmes0 | F531/F552, L531 | idem (sin Mes2) | sí (2 filas) |

### 6.3 Invariantes de cierre ya vigentes (T0e del Plan 21)

| Invariante (firma) | Presencia | Consecuencia |
|---|---|---|
| A=`Componente` ∧ B=`Total` | 15/15 | dura (corte de zona) |
| A=`Subs/Cont` ∧ B=`Total` | 15/15 | dura |
| A=`Total` ∧ B vacío | 15/15 | dura |

Las filas `Aplicacion` y `Mes` NO son invariantes de cierre: su ausencia es legítima y se resuelve
por rol (EXTEMP=0 / ocurrencia disponible).

---

## 7. Veredictos T0a..T0e (resumen)

| ID | Entregable | Veredicto | Estado |
|---|---|---|---|
| T0a | Dump roles por ASE × período + dump A–E ASE3-agosto | Q2-julio: 5/5 ASE traen `Aplicacion`; agosto: 4/5 (solo **ASE3** sin `Aplicacion`); Q1 no usa este mapa (EXTEMP estático 0 en ASE2/ASE5); ASE3-agosto `Mes/Total` = 3 (R11/R27/R39), sin riesgo | DONE |
| T0b | Semántica de ausencia | `Aplicacion` = subtotal del sub-bloque previo (eco en ASE1/ASE3, subtotal en ASE2/4/5). Ausencia ASE3-agosto ⇒ EXTEMP = **0 legítimo** (oráculo R10: implícito 0,33). Regla que casa 10/10: **`EXTEMP = Aplic0 + Aplic1`** | DONE |
| T0c | Impacto en celdas destino | `ExtemporaneoEsperadoPorAse` → `TotalD104` (DetRetri); `totOpt` → gate Σ empresas; `F25` → gate HU-02; en 5-ASE el espejo supersede la ESCRITURA R1 pero NO el agregado `extemp` | DONE |
| T0d | Veredicto de diseño | (i) opcional puro **rechazada** (desvío 1.088e9 vs R10); (ii) firma sobre secuencia **recomendada**; (iii) corrección mínima del mapa = plan B | DONE |
| T0e | Tabla congelada | §6: roles `Aplic` corregidos (Subs0→Aplic1 en ASE1/ASE3) + opcionalidad; `Mes` obligatorios | DONE |

No hay `NEEDS_CONTEXT` de T0a/T0b. Las decisiones que requieren aprobación están en §8.

---

## 8. Decisiones que el Ingeniero debe aprobar (antes del plan de HU)

1. **Rol de `F37/F270` (ASE1/ASE3):** ¿se confirma el cambio `Subs0 → Aplic1` (evidencia: R10 casa
   10/10 con `Aplic0+Aplic1`; la fila destino es `Aplicacion`)? Recomendación: **sí**.
2. **Estrategia de resolución:** ¿(ii) firma/etiqueta sobre la secuencia espejo, o (iii) corrección
   mínima del mapa congelado? Recomendación: **(ii)**, con (iii) como plan B.
3. **Regla de agregado con ≠2 filas `Aplicacion`:** ¿suma de TODAS las filas `Aplicacion`
   (ausente/cero filas ⇒ 0)? Evidencia actual: siempre 2. Recomendación: **suma de todas**.
4. **Alcance del fix:** ¿solo el path Q2 5-ASE, o también el single-ASE (`ProcesadorRemuneracion`,
   que hoy no adjunta `EspejoR1` y escribe las celdas leaf R1)? Recomendación: **ambos**, porque el
   reader es compartido.
5. **Criterio de regresión 2026082:** al resolver el blocker, ¿la rama end-to-end debe asertar que el
   DetRetri de los 5 ASE coincide con el R10 (16.369.059.896 en ASE3) además de las invariantes y Δ?
   Recomendación: **sí** (el R10 existe; es el único oráculo de agosto).

---

## 9. Reproducibilidad (método, todo en disco)

Scripts de medición en `%TEMP%\opencode\` (no versionados; invocados con `-File` y `Get-Location`
para evitar el encoding de acentos en `Automatización`):

- `p24roles.ps1` — dump de roles R1Q2 por ASE × período (Mes/Aplicacion/Subsidio + Especiales) y
  resolución rol→celda (salida `p24roles_out.txt`).
- `p24echo.ps1` — verifica si `Aplicacion` coincide con la fila `Total` previa (11/24).
- `p24oracle.ps1` — recomputa `totOpt`, R2, R4, Ajustes por ASE/período y despeja `extemp` del R10
  (`p24oracle_out.txt`); control julio-Q2 + agosto.
- `sheetsig.ps1`, `dumpwb.ps1` — herramientas existentes (secuencia A–E + F; fórmulas/valores por hoja).

Comandos (PowerShell 5.1, cwd = raíz del repo):

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File "$env:TEMP\opencode\p24roles.ps1"
powershell -NoProfile -ExecutionPolicy Bypass -File "$env:TEMP\opencode\p24echo.ps1"
powershell -NoProfile -ExecutionPolicy Bypass -File "$env:TEMP\opencode\p24oracle.ps1"
```

Archivos crudos de evidencia en `%TEMP%\opencode\`: `p24roles_out.txt`, `p24oracle_out.txt`.

Insumos usados (nombres exactos por patrón):
- Q1-julio: `Docs/Insumos/REMUNERACION 2026071/{1..5}-*/Recaudoporcomponente_*_01072026_*_15072026_*.xlsx`.
- Q2-julio: `Docs/Insumos/REMUNERACION 2026072/{1..5}-*/Recaudoporcomponente_*_16072026_*_31072026_*.xlsx`.
- Agosto: `Docs/Prueba2/Insumos/{1..5}-*/Recaudoporcomponente_*_16082026_*_31082026_*.xlsx`
  (+ `RerpoteDetalleSaldosaFavor_*`, `Reversi*PorComponente_*`, `SaldosaFavorAplicadosPorNotas_*`,
  `Retribuci*Negativa_*` para R2/R4/Ajustes).
- Oráculo: `Docs/Prueba2/Insumos/R10_Remuneracion_2026082.xlsx`, hoja `DetRetri2026082`, D9:D13.
- Goldens: `%TEMP%\opencode\Remuneracion 202607-1 Total.xlsx` y `...-2 Total.xlsx`.
