# T0 Plan 50 — Fuentes de recaudo opcionales (H0 read-only)

> **Alcance:** H0 del `plans/50 - opcionalidad-fuentes-recaudo-y-unificacion-q1-q2.md` (APROBADO 2026-10-21).
> **Método:** lectura de BYTES con BCL (`System.IO.Compression.ZipArchive` + `System.Xml.Linq`), sin Excel/COM,
> sin OpenXml SDK y sin ExcelDataReader. Cero código de producción. Cero escritura sobre fuentes ni plantillas.
> **Fixture que congela este documento:** `Remuneracion.IntegrationTests/FuentesRecaudoOpcionalesT0Plan50Tests.cs`.
> **Fecha:** 2026-10-21. **Autor:** implementer (H0).

---

## 0. Veredicto H0 (puerta H0 → H1)

| Ítem | Veredicto |
|---|---|
| S1.1 inventario de ausencias | **CONGELADO** — 8 ausencias totales de label en las 3 muestras (octubre-1, julio-2, agosto-2); el resto de labels de los mapas trae filas. |
| S1.2 granularidad | **ABSORBIBLE por O-C** — las 8 ausencias son *ausencia total del label* (0 filas) → 0 + huella sin ramas. **1 caso parcial** (octubre-1 ASE3 R1 OCCIDENTE) queda **fail-fast** por la regla ratificada §7.3 y se declara como **residuo con su propio T0** (no se parchea el mapa). |
| S1.3 A3 | **ARITMÉTICA NO VIABLE en R1 y R2** (medido con bytes). En R4 cierra por construcción (sin poder de detección propio). → Se implementa el **control mínimo alternativo (§2.5): control de etiquetas no reconocidas**, con este T0 propio. |
| S1.4 residuo Q2 | **AUSENCIAS CUBIERTAS por A1; 2 residuos declarados.** (a) `R-SIGMA-Q2`: el gate Σ empresas-vs-bloque no es válido en Q2 (14 divergencias medidas, §4.2) → acotado a Q1. (b) `R-MAPA-V2-R2-ASE4`: la celda destino `Rem. Anticipos R2!E285` es **fórmula** en la v2 (era literal en v1) → **A2 NO se habilita** (§4.3). |
| S1.5 matriz de mensajes | **CONGELADA** (§5). |
| Puerta | **H0 VERDE con residuos declarados** → H1 habilitado. G-SHA: v1 intacta; v2 **discrepa** del SHA congelado por un commit previo a este plan (§6). |

**STOP no declarado** por una razón: la política O-C absorbe el 100 % de las ausencias observadas con *cero*
ramas por empresa/período (el label no aparece en el reporte → los operandos del mapa valen 0). Los dos
residuos no son ausencias de fuente: (a) el caso parcial de octubre-1 es una diferencia de *forma* del bloque
de OCCIDENTE y la decisión ratificada (§7.3) es dejarlo en fail-fast; (b) `R-SIGMA-Q2` es un gate de HU-08
que no descompone el visible de bloque en Q2 y se acota a Q1 conservando el comportamiento Q2 actual.

---

## 1. S1.1 — Inventario de ausencia por (período × ASE × reporte × label)

Criterio de localización = espejo del reader (§2.1 del plan):

| Reporte | Etiqueta de la fuente | Fila de cierre |
|---|---|---|
| R1 `Recaudoporcomponente` | col **C** | col **D** = `Total` |
| R2 `RerpoteDetalleSaldosaFavor` | col **B** | col **C** = `Total` |
| R4 `ReversiónPorComponente` | col **A** | col **B** = `Total` |

Las empresas con operando declarado por ASE son las de `WorkbookLeafCellMapPorEmpresa`
(`EditablesR1/R2/R4PorEmpresa`): R1/R2 ASE1 `{ENEL, OCCIDENTE}`; ASE2 `{NUEVO ESQUEMA, ENEL, OCCIDENTE}`;
ASE3 `{ENEL, OCCIDENTE}`; ASE4 `{NUEVO ESQUEMA, ENEL, OCCIDENTE}`; ASE5 `{ENEL, ENERBIT, OCCIDENTE}` (R2 ASE5
solo `{ENEL, OCCIDENTE}`); R4 ASE1 `{OCCIDENTE}`, ASE2 `{NUEVO ESQUEMA, ENEL}`, ASE3 `{ENEL, OCCIDENTE}`,
ASE4 `{NUEVO ESQUEMA, ENEL, OCCIDENTE}`, ASE5 `{ENEL, OCCIDENTE}`.

### 1.1 Ausencias totales (0 filas del label) — las 8 que absorbe A1

| # | Período | ASE | Reporte | Label | Evidencia (archivo · hoja `Sheet1` · filas del reporte) |
|---|---|---|---|---|---|
| 1 | 2026101 (oct-1) | 2 LIME | R1 | `NUEVO ESQUEMA` | `Recaudoporcomponente_…_20261016174350151.xlsx`: 0 filas con `C=NUEVO ESQUEMA & D=Total`; etiquetas observadas `ENEL` (filas 7,16,27,42,49) y `OCCIDENTE` (10,22,37,45,56); invariante `Componente/Total` fila 39 (`F39 = 14.508.429.002,77`) y `Total` fila 59 |
| 2 | 2026101 | 2 LIME | R4 | `NUEVO ESQUEMA` | `ReversiónPorComponente_…_20261016174430392.xlsx`: solo `ENEL` (fila 7, `D7 = -20.858.585,45`) y `OCCIDENTE` (fila 13, `D13 = -945.063`); `Total` fila 14 = `-21.803.648,45` |
| 3 | 2026101 | 3 CIUDAD LIMPIA | R4 | `OCCIDENTE` | `ReversiónPorComponente_…_20261016174833303.xlsx`: solo `ENEL` (fila 7, `D7 = -2.259.272`); `Total` fila 8 |
| 4 | 2026101 | 5 ÁREA LIMPIA | R4 | `OCCIDENTE` | `ReversiónPorComponente_…_202610161752242.xlsx`: solo `ENEL` (fila 7, `D7 = -3.964.157`); `Total` fila 8 |
| 5 | 2026072 (jul-2) | 2 LIME | R4 | `NUEVO ESQUEMA` | `ReversiónPorComponente_…_2026847164046.xlsx`: `ENEL` fila 7 (`-70.779.009`) y `OCCIDENTE` fila 13 (`-290.750`); `Total` fila 14 = `-71.069.759` |
| 6 | 2026082 (ago-2) | 1 PROMOAMBIENTAL | R4 | `OCCIDENTE` | `ReversiónPorComponente_…_20269119382545.xlsx`: solo `ENEL` (fila 6, `-73.090`); `Total` fila 7 |
| 7 | 2026082 | 2 LIME | R4 | `NUEVO ESQUEMA` | `ReversiónPorComponente_…_202691194047838.xlsx`: `ENEL` fila 7 (`-13.447.592`) y `OCCIDENTE` fila 13 (`-24.890`); `Total` fila 14 = `-13.472.482` |
| 8 | 2026082 | 3 CIUDAD LIMPIA | R4 | `OCCIDENTE` | `ReversiónPorComponente_…_20269119433567.xlsx`: solo `ENEL` (fila 7, `-4.617.354`); `Total` fila 8 |

**Observaciones de bytes relevantes:**

- Las 8 ausencias son **ausencia total del label** (0 filas). No hay ninguna ausencia "parcial sin filas" que
  exija distinguir bloques: la rama `filasTotal.Count + filasSubs.Count == 0` (R1) / `filasEmpresa.Count == 0`
  (R2) / "sin fila `Total`" (R4) es la condición exacta.
- **julio-1 (2026071)** se leyó también como contramuestra (no exigida por el plan): **0 ausencias** en los
  tres reportes para los 5 ASE.
- El caso que motivó el plan (rojo `2026101` ASE2) es el **#1**: `NUEVO ESQUEMA` no viene en R1 (sí viene en R2,
  filas 22 y 48) — confirma la doctrina "R1/R2/R4 son reportes independientes".

### 1.2 Presencia (todo lo demás)

La matriz completa (período × ASE × reporte → labels presentes/ausentes) está congelada en
`MatrizAusenciasData` del fixture (45 líneas `período|ASE|reporte|presentes|ausentes`). Cualquier cambio de
forma de una fuente nueva rompe ese fixture → exige T0 nuevo (patrón `ContratoSeccionPeriodFreeT0Tests`).

---

## 2. S1.2 — Granularidad: total vs parcial

### 2.1 Ausencias totales → 0 + huella (A1)

Las 8 de §1.1. No hay ambigüedad de bloques: el label no existe en el reporte.

### 2.2 Parcial con label presente → fail-fast (residuo declarado)

| Caso | Período | ASE | Reporte | Empresa | Evidencia | Operando que no resuelve |
|---|---|---|---|---|---|---|
| **R-PARCIAL-OCT-ASE3-R1-OCC** | 2026101 | 3 CIUDAD LIMPIA | R1 | OCCIDENTE | `Recaudoporcomponente_…_20261016174753997.xlsx` hoja `Sheet1`: **una sola** fila `Total` de OCCIDENTE en el bloque Componente (fila 23, `F23 = 286.711.907,53`) y **una sola** en Subs/Cont (fila 35, `F35 = 145.489.038,47`) | `FuenteR1.TotalMedio` = índice 1 de la lista → no existe (`ElementAtOrDefault(1)` = null) |

- **Regla aplicada (ratificada §7.3):** bloque parcial con label presente = **fail-fast vigente**; sólo un T0 de
  negocio con bytes puede moverlo a "0 documentado". Este documento **no** lo mueve: no hay oráculo del manual
  para octubre-1 que permita decidir si `F216/F222/F242` (TotalMenor/TotalMedio/SubsMenor) valen 0 cuando la
  fuente trae un único bloque de OCCIDENTE.
- **Consecuencia operativa (declarada):** una vez A1 elimine el throw de ASE2 en `2026101`, la corrida
  **seguirá deteniéndose en el ASE3** con el mensaje vigente
  `ASE 3: no se encontró la fila Total (bloque medio) de la empresa OccidenteDirecta en la fuente.`
  Ese rojo **no es una regresión de A1**: es el residuo pre-existente, enmascarado hoy por el aborto temprano en ASE2.
- **Dueño / condición de retiro:** T0 de negocio propio (¿la remuneración de OCCIDENTE con un solo bloque debe
  tomar TotalMenor/TotalMedio = 0?). Fuera de A1/A2/A3; **prohibido parchear el mapa** (G2-D1).

---

## 3. S1.3 — Veredicto A3 (control anti-falso-0)

### 3.1 La aritmética propuesta NO es viable en R1

Operandos medidos: `Σ` de **todas** las filas-etiqueta del bloque `Componente` (col F) vs fila invariante
`A='Componente' B='Total'` (col F).

| Período | ASE1 | ASE2 | ASE3 | ASE4 | ASE5 |
|---|---:|---:|---:|---:|---:|
| 2026101 | 2.015,07 | 243.714,65 | 558,05 | 30.180,94 | 0 |
| 2026072 | 11.217,37 | 405.906,75 | 16.458,46 | 439.956,52 | 0 |
| 2026082 | 9.123,90 | 1.443.509,98 | 4.456,11 | 51.713,94 | 4.485,01 |

**Causa (bytes):** `Componente/Total` **no** es la suma de las filas-etiqueta. Está compuesto por la fila
`Mes/Total` de la última sección **más** la fila frontera `Aplicacion nuevos x reversion/Total`; las filas de la
*sección primera* (p. ej. `ENEL` fila 7 de octubre-1 ASE3 = `-558,05`) quedan **fuera** del total. Reproducir esa
composición exige un modelo de secciones que el reader no tiene → **no viable sin reescribir el mapa (prohibido)**.

### 3.2 La aritmética NO es viable en R2 (etiqueta real sin destino)

`Σ` de **todas** las filas-etiqueta del bloque `Componente` (col E) **sí** coincide exacto con
`Componente/Total` en 20/20 casos (es una tautología: la fila de cierre es la suma de esas filas).
Al restringir el `Σ` a las etiquetas **reconocidas** (las que tienen destino en el mapa) aparece Δ ≠ 0 por la
etiqueta **`EAB`**, que es una fuente real **sin destino** en el mapa (deuda declarada del Plan 49: R2 ASE4
`A/EAB = 12.462,26`):

| Período | ASE | `Σ` reconocidas | `Componente/Total` | Δ | Etiqueta excluida |
|---|---|---:|---:|---:|---|
| 2026101 | 2 | 116.261.282,80 | 116.410.420,67 | **149.137,87** | `EAB` (fila 7) |
| 2026072 | 3 | 23.633.735,10 | 23.707.258,20 | **73.523,10** | `EAB` (fila 4) |
| 2026082 | 3 | 31.159.700,50 | 31.231.390,27 | **71.689,77** | `EAB` (fila 4) |
| 2026082 | 4 | 126.710.780,17 | 126.782.453,43 | **71.673,26** | `EAB` (fila 4) |
| 2026101 | 1 | 31.276.742,42 | 31.276.742,42 | 0 | — |

→ Un control aritmético sobre R2 dispararía **falso rojo** en corridas sanas (el dinero de `EAB` no es una
"etiqueta cambiada": es una fuente conocida sin dueño en el mapa). **No viable como anti-falso-0.**

### 3.3 R4 cierra, pero por construcción

`Σ` filas-etiqueta (col D) == fila `A='Total'` (col D) en **20/20** casos (Δ = 0 exacto, todas las muestras).
Como la fila `Total` **es** la suma de las filas-etiqueta, la igualdad es una tautología; sólo aporta poder de
detección si se restringe a etiquetas reconocidas — y en ese caso la señal es exactamente la misma que la del
control de etiquetas (el dinero de una etiqueta desconocida aparece como Δ).

### 3.4 Inventario de etiquetas no reconocidas (insumo del control)

"Conocida" = pertenece al **vocabulario de fuentes del negocio** (`EmpresaFacturacion.Catalogo`: rótulos
`RECIPROCIDAD`, `ENEL`, `ENERBIT`, `OCCIDENTE`, `CIUDAD LIMPIA-ACUEDUCTO`, más el rótulo vigente por ASE
`NUEVO ESQUEMA`, unificados por la sinonimia citada del Plan 34). Con ese vocabulario el inventario
observado en las tres muestras es **una sola etiqueta**:

| Período | ASE | Reporte | Etiqueta | Filas (evidencia) | Origen (deuda ajena, no se toca) |
|---|---|---|---|---|---|
| 2026101 | 2 | R2 | `EAB` | 7 y 33 (`E7 = 149.137,87`, `E33 = -18.613,87`) | fuente real sin destino en el mapa R2 (Plan 49: `A/EAB`) |
| 2026072 | 3 | R2 | `EAB` | 4 (`E4 = 73.523,10`) | ídem |
| 2026082 | 3 | R2 | `EAB` | 4 (`E4 = 71.689,77`) | ídem |
| 2026082 | 4 | R2 | `EAB` | 4 y 32 (`E4 = 71.673,26`, `E32 = -12.462,26`) | ídem (el `12.462,26` es deuda declarada del Plan 49) |

`ENEL` sin celda destino en R4 de ASE1/3/5 y `ENERBIT` sin celda destino en R2 de ASE5 **no** son
"etiquetas no reconocidas" (pertenecen al vocabulario): son **huecos de cobertura del mapa**, que el gate
Σ de Q1 detecta como dinero no atribuido (§4.2) y que este plan **no** corrige (deuda del mapa HU-08,
fuera de A1/A2/A3).

**Ninguna** de las etiquetas no reconocidas coexiste con una ausencia **en su mismo reporte** (las
ausencias están en R1/R4) → el control no dispara en ninguna corrida sana. Ésta es la evidencia que
sustenta el acotamiento por reporte del juicio A3.

### 3.5 Veredicto y control elegido
**Se implementa el control mínimo alternativo de §2.5: control de ETIQUETAS NO RECONOCIDAS** (T0 propio = este
documento), aplicado a R1, R2 y R4 con una sola costura:

- **Insumo (evidencia, sin juicio):** por (ASE, reporte), la lista de etiquetas observadas en la columna de
  fuente de las filas de cierre (`R1:C` con `D=Total`, `R2:B` con `C=Total`, `R4:A` con `B=Total`) que **no**
  resuelven a ninguna empresa del mapa de ese (ASE, reporte) ni por la sinonimia citada
  (`SinonimosEmpresaR1`: `RECIPROCIDAD` ↔ `NUEVO ESQUEMA`). Se reporta como `EtiquetasNoReconocidas`.
- **Juicio (Core, junto a los gates 2.7):** si en un (ASE, reporte) hay **al menos una etiqueta no reconocida**
  **y** al menos una **fuente ausente declarada** (la huella de A1 en ese mismo reporte) → **fail-fast** con el
  formato de §2.6 (ASE + reporte + etiquetas + ausentes). Es la firma de "fuente presente no reconocida"
  (cambio de etiqueta): hay dinero en el reporte que ninguna fuente reconocida explica.
- **Sin umbral numérico:** el T0 declara no viable la vía aritmética; el umbral ±0.5 no aplica a este control
  (no hay resta). Queda documentado como desviación justificada de §2.5 (que sólo prevé umbral si la aritmética
  es viable).

---

## 4. S1.4 — Inventario de divergencias Q2 restantes

### 4.1 Lo que A1/A2 absorben

| Divergencia Q2 documentada | Evidencia de bytes | Veredicto |
|---|---|---|
| R4-por-empresa Q2 (`ProcesadorPeriodo.cs:205-214`, causa del recorte) | julio-2 ASE2 R4 sin `NUEVO ESQUEMA`; agosto-2 ASE1 R4 sin `OCCIDENTE`, ASE2 R4 sin `NUEVO ESQUEMA`, ASE3 R4 sin `OCCIDENTE` | **A1 lo cubre** (ausencia total → 0 + huella). El mapa HU-08 **no se reescribe**. |
| R1 por empresa en Q2 | julio-2/agosto-2: todos los labels de R1 presentes; ningún parcial (OCCIDENTE ASE3 tiene 2 filas Total en ambos períodos) | **Sin residuo** |
| R2 por empresa en Q2 | julio-2/agosto-2: todos los labels de R2 presentes (más `EAB`/`ENERBIT`, sin destino) | **Sin residuo de ausencia** |
| `leaf.Recaudos` (RESUMEN MES F/G en Q2) | Ya levantado por el Plan 20 (G2) | **Intacto** — A1/A2 no lo tocan |
| Escrituras Q2 nuevas | Las hojas por empresa reciben los mismos operandos que en Q1 (0 donde ausente) | **Sin celdas nuevas** fuera de los mapas por empresa |

### 4.2 RESIDUO Q2 DECLARADO — `R-SIGMA-Q2`: el gate Σ no es válido en Q2

Si A2 puebla `leaf.Conciliacion` en Q2, el gate `ValidarSigmaEmpresasPorAse` (Σ visibles-empresa ==
visible de bloque, HU-08 §2.5) se activaría por primera vez en Q2. **Medido en bytes: no cierra.** El gate
se definió sobre los mapas Q1 y no es una descomposición válida del visible de bloque en Q2:

| Período | ASE | Reporte | Σ empresas | Visible de bloque | Δ | Causa (bytes) |
|---|---|---|---:|---:|---:|---|
| 2026072 | 1 | R4 | -1.002.751,39 | -16.536.713,09 | **15.533.961,70** | `ENEL` (fila 7, `-15.533.961,70`) no tiene celda destino en el mapa R4 de ASE1 |
| 2026072 | 2 | R4 | -70.779.009 | -71.069.759 | **290.750** | `OCCIDENTE` (fila 13) sin destino en el mapa R4 de ASE2 |
| 2026072 | 3 | R2 | 26.512.805,19 | 26.586.328,29 | **73.523,10** | etiqueta `EAB` sin destino (deuda Plan 49) |
| 2026072 | 5 | R2 | 36.741.099,93 | 42.983.502,41 | **6.242.402,48** | etiqueta `ENERBIT` sin destino en el mapa R2 de ASE5 |
| 2026072 | 1 | R1 | 17.339.804.777,23 | 17.339.793.559,86 | **11.217,37** | el visible R1-Q2 se deriva por firma (espejo), no de estos operandos |
| 2026072 | 2 | R1 | 20.365.128.471,20 | 20.352.268.540,72 | **12.859.930,48** | ídem (+ `Block5Total` del bloque 5) |
| 2026072 | 3 | R1 | 15.226.889.745,16 | 15.226.209.207,13 | **680.538,03** | ídem (fila de la sección primera) |
| 2026072 | 5 | R1 | 24.856.625.673,59 | 12.033.011.685,71 | **12.823.613.987,88** | ídem: con una sola fila `Total`, `TotalMain` y `TotalMenor` son la MISMA fila (doble conteo) |
| 2026082 | 1 | R4 | 0 | -73.090 | **73.090** | ídem (`OCCIDENTE` ausente + `ENEL` sin destino) |
| 2026082 | 2 | R4 | -13.447.592 | -13.472.482 | **24.890** | ídem |
| 2026082 | 3 | R2 | 27.735.175,20 | 27.806.864,97 | **71.689,77** | `EAB` |
| 2026082 | 4 | R2 | 111.058.724,44 | 111.117.935,44 | **59.211** | `EAB` |
| 2026082 | 5 | R2 | 28.908.130,83 | 34.963.153,78 | **6.055.022,95** | `ENERBIT` |
| 2026082 | 1/2/5 | R1 | — | — | 9.123,90 / 11.506.205,90 / 13.260.342.962,39 | ídem (firma vs operandos) |

**Veredicto:** el gate Σ queda **acotado a Q1** en sus DOS costuras (`ValidadorBasico.ValidarSigmaEmpresasPorAse`
y el gemelo `WorkbookLeafCoherence.ValidarSigmaEmpresas` invocado por el writer). En Q2 se mantiene el
comportamiento actual: no se evalúa (hoy tampoco se evaluaba, porque la lista de conciliación llegaba
vacía). A2 sí puebla la conciliación y escribe los operandos por empresa en Q2; el control de Q2 sigue
siendo el gate workbook-vs-dominio (W-1, Plan 40) y el DetRetri-vs-R10 (G3). Corregir el gate Σ en Q2 exige
el **contrato nuevo** (descomposición por firma + cobertura completa del mapa: celdas destino para `ENEL` en
R4 ASE1/3/5, `ENERBIT` en R2 ASE5 y `EAB`), que es un frente distinto con su propio T0 — **prohibido
parchear el mapa aquí** (§1.3, G2-D1).

**Consecuencia declarada para el período objetivo `2026101`:** aun con A1/A2, la corrida **no completará**:

1. ASE3 R1 OCCIDENTE → parcial `TotalMedio` (§2.2) — aborta en lectura, antes de validar;
2. ASE1 R4 (Δ 3.698.870 = `ENEL` sin destino) y ASE2 R2 (Δ 130.524 = `EAB`) fallarían el gate Σ de Q1 si
   la corrida llegara a la validación.

Ninguno de los dos es causado por A1/A2: son defectos pre-existentes del mapa HU-08 (deuda con dueño
propio) que el aborto temprano en ASE2 enmascaraba. **S2.3 se verifica por ASE** (el ASE2 deja de lanzar
`CalculoInvalidoException` por la ausencia de `NUEVO ESQUEMA`), no por corrida completa.

**Residuo Q2: `R-SIGMA-Q2` (declarado, con este T0 como evidencia).** El único residuo Q1 es el parcial de
§2.2.

---

### 4.3 BLOQUEO DE A2 EN v2 — `R-MAPA-V2-R2-ASE4` (la celda destino es fórmula)

Al reactivar la escritura de la conciliación por empresa en Q2 (A2) la corrida Q2 aborta en el guard
anti-fórmula del writer:

```
CalculoInvalidoException: La celda leaf R2.ReciprocidadEaab.E285 (Rem. Anticipos R2!E285) es fórmula
en la plantilla. HU-05 no puede sobrescribir fórmulas.
```

**Causa (bytes):** la plantilla v2 convirtió en FÓRMULA la fila-total de la empresa del bloque R2 de ASE4:

| Celda | v1 (`3086AC31…`) | v2 (`0F192AD2…`) | Mapa HU-08 |
|---|---|---|---|
| `Rem. Anticipos R2!E285` | literal `0` | **fórmula** `E192+E203-K192` | `EditablesR2PorEmpresa[(1,4)]` → `("E285", Total)` |
| `Rem. Anticipos R2!K285` | literal `0` | **fórmula** (familia shared) | `EditablesR2PorEmpresa[(1,4)]` → `("K285", Especiales)` |

Es drift **mapa↔plantilla** de la clase que el Plan 48 se comprometió a re-congelar por T0 ("una v3 exige
re-congelar el mapa con su T0, nunca edición silenciosa"). El defecto estaba **latente**: Q1 nunca llega a
escribir (aborta antes en R1: `ERR-SIN-SLOT` de Q1 ASE2) y Q2 omitía la conciliación (recorte HU-20/G2-D1).

**Veredicto:** A2 **NO se habilita** en este plan. El recorte Q2 se mantiene y su mensaje se re-etiqueta con
este residuo. Reactivarlo exige **re-congelar el mapa R2 (1,4) contra la v2** (T0 propio, con la
propietaria), nunca editar el mapa a ciegas ni relajar el guard anti-fórmula (W-A/Plan 28/48).

**Consecuencia:** `leaf.Conciliacion` se puebla en Q1 (con la política A1) y en Q2 sigue vacío (recorte).
S2.2 queda **cumplido en la lectura period-free del reader** (`S2_2_Reader_LeeLaConciliacionPorEmpresa_EnQ1yQ2_PeriodFree`)
pero **NO en el levantamiento del recorte** — desviación declarada.

---

### 4.4 e2e del período objetivo `2026101` — bloqueos pre-existentes que el rojo de ASE2 enmascaraba

Corrida CLI real (read-only sobre fuentes; salida a carpeta temp fresca, plantilla v2 oficial):

```
dotnet run --project Remuneracion.Cli -- --periodo 2026101 --carpeta "Docs\Prueba Octubre-1\Insumos" \
  --plantilla "Docs\Plantilla estandarizada\v2\Formato Remuneracion 2026xx-x Total vf.xlsx" --salida <temp> --cinco-ase

=> ASE 1: ... conciliación por empresa OK ... DetRetri esperado 18.052.547.615
=> ASE 2 (Lime): leyendo conciliación por empresa (R1/R2/R4)...   <-- A1: la ausencia de NUEVO ESQUEMA
                                                                       en R1 ya NO aborta
=> RESULTADO ERROR codigo=ERR-VALIDACION salida=1 runId=24f9d7ee-…
   [ERR-VALIDACION] ASE 2: la columna fuente 2 del Resumen debería ser 'NUEVO ESQUEMA' pero dice
   'OCCIDENTE'. Mapa T0-0.3 no coincide con la fuente.
```

**Lectura:** el rojo que motivó el plan (`ASE 2: no se encontró la fila Total de la empresa
ReciprocidadEaab en la fuente`) **desapareció**. La corrida avanza y muere en un bloqueo **distinto y
pre-existente**: el mapa **posicional** T0-0.3 de `ReportePagosxBanco` (`LeerReporteBanco`,
`ExcelDataReaderWorkbookLeafInputReader.cs:251-262`) espera `NUEVO ESQUEMA` en la columna 2 del Resumen y
octubre-1 trae `OCCIDENTE` — es la deuda "mapeo por ETIQUETA, nunca por posición" en HU-09, ajena a
A1/A2/A3.

**Bloqueos pre-existentes del período objetivo (todos con dueño propio, ninguno se toca aquí):**

| # | Bloqueo | Evidencia | Dueño |
|---|---|---|---|
| 1 | `ReportePagosxBanco` ASE2: columnas por POSICIÓN (T0-0.3) | mensaje arriba (`:251-262`) | HU-09 (mapa por encabezado) |
| 2 | R1 ASE3 OCCIDENTE parcial (`TotalMedio` sin fila) | §2.2 | T0 de negocio |
| 3 | Gate Σ Q1: ASE1 R4 `ENEL` sin celda destino (Δ 3.698.870) y ASE2 R2 `EAB` (Δ 130.524) | §4.2 / sonda §4.2 | mapa HU-08 (cobertura) |
| 4 | `ERR-SIN-SLOT` del contrato de sección R1 en fuentes Q1 | pre-existente (Q1 julio-1) | Plan 41/46 (Q1) |

**Consecuencia:** S2.3 se verifica **por ASE** (el ASE2 deja de lanzar por la ausencia) y S3.2 (tabla
e2e Q1 completa) **no es alcanzable** con A1/A2/A3: el período objetivo sigue sin poder completarse por
bloqueos ajenos al alcance aprobado. Se declara como desviación con evidencia.

---

## 5. S1.5 — Matriz de mensajes residuales (qué throw queda)

| Caso | ¿Aborta? | Mensaje / huella |
|---|---|---|
| Fuente ausente (label con 0 filas) en R1/R2/R4 | **No** | `Log.Warning` estructurado (`AseId`, `Reporte`, `Fuente`, `Empresa`, `CeldasEnCero`) + entrada `R1:NUEVO ESQUEMA` / `R4:OCCIDENTE` en `FuentesAusentes` (resultado) |
| Bloque parcial con label presente (p. ej. `TotalMedio` sin fila) | **Sí** | `ASE {id}: no se encontró la fila {concepto} de la empresa {nombre} en la fuente.` (vigente, `ValorDe`) / `… para sus Especiales.` (`EspecialesDe`) / `… no se encontró la fila Total de la empresa {nombre} en la fuente R4.` (solo si el label NO trae NINGUNA fila → pasa a A1) |
| Invariante de cierre R1 (`A='Componente'`) ausente | **Sí** | `ASE {id}: no se encontró la fila 'Componente' en la fuente R1 para localizar los bloques por empresa.` (intacto) |
| Invariantes T0e del espejo (`Componente/Total`, `Subs/Cont/Total`, `Total`) | **Sí** | `Espejo estructural: la fuente R1 del ASE {id} ({nombre}) no trae las filas invariantes de cierre en el reporte 'Reporte Componentes R1': {faltantes}.` (intacto, `ValidarInvariantesEspejoR1`) |
| Etiqueta no reconocida **con** fuente ausente en el mismo (ASE, reporte) | **Sí** | `[ERR-VALIDACION] ASE {n} · {reporte} · control de fuentes: etiqueta(s) no reconocida(s) [{X}] con fuente(s) ausente(s) [{Y}]. (Plan 50 A3: posible fuente presente no reconocida.)` |
| Etiqueta no reconocida **sin** ausencias en el mismo (ASE, reporte) | **No** | Sólo evidencia (`EtiquetasNoReconocidas` en el resultado). Hoy: `EAB` en R2 (`2026101` ASE2; `2026072` ASE3; `2026082` ASE3/ASE4) y `ENERBIT` en R2 ASE5 (las tres muestras) |
| Escritura de una celda del mapa sin valor | **Sí** | `No hay valor R1/R2/R4 por empresa para {nombre} ({hoja}!{celda}) del ASE {id}.` (writer intacto, W-A) |

---

## 6. Gate G-SHA — hallazgo de entrada (⚠️ pre-existente, no causado por este plan)

| Plantilla | Ruta | SHA256 congelado | SHA256 en el working tree | Veredicto |
|---|---|---|---|---|
| v1 histórica | `Docs/Plantilla estandarizada/Formato Remuneracion 2026xx-x Total.xlsx` | `3086AC31…C5214` | `3086AC317595C2F8E84ED24B1C5EE3E52565DD5193394E3062878E7FA49C5214` | **INTACTA** ✅ |
| v2 | `Docs/Plantilla estandarizada/v2/Formato Remuneracion 2026xx-x Total vf.xlsx` | `FF5CF7A819110A1E1E2CA7BDAD553DD7A25141ADD472B1904BD496F082642632` | `0F192AD2721C77986B6A8CB85D1A92F782DCD2F69989F362E066FA17027F2DAB` | **DISCREPA** ⚠️ |

**Causa localizada (git):** el commit `d55e621 "Ajustes nueva plantilla"` (2026-10-21 12:01) reemplazó el
binario de la v2 (`689.790 → 693.400` bytes) **y** dos fuentes R2 (`…202691153449402.xlsx`,
`…202683203011528.xlsx`) en el mismo commit que la suite base `p49-h2-suite.trx`. El archivo **no** fue tocado
por este plan (H0 es read-only): el desfase existe **antes** de H1.

**Consecuencias declaradas (no se corrigen aquí — no está en el alcance de A1/A2/A3):**

1. Los fixtures que congelan el SHA viejo (`AdopcionPlantillaV2T0Plan48Tests.ShaV2`, `DivergenciaDeclaradaT0Tests`)
   quedan **rojos por diseño** en el árbol actual.
2. La suite base `p49-h2-suite.trx` (1061/1012/49/0) se produjo **antes** del reemplazo de la v2, por lo que
   el set-diff de H2 debe compararse también contra una **línea base re-medida en el árbol actual**
   (`p50-h0-baseline.trx`), para no atribuir a este plan rojos pre-existentes.
3. El gate G-SHA del plan 50 se cumple en su parte exigible: **la v2 no se modifica por este plan**
   (SHA antes == SHA después de la implementación). La re-congelación del nuevo SHA exige T0 propio (patrón
   Plan 48) y decisión del Ingeniero.

---

## 7. Evidencia reproducible

| Qué | Cómo |
|---|---|
| Matriz completa de labels por bytes | `FuentesRecaudoOpcionalesT0Plan50Tests` (46 líneas de matriz + parcial + aritmética + etiquetas) |
| Ausencias | `T0P50_MatrizAusenciaTotal_PorPeriodoAseReporte` |
| Parcial octubre-1 ASE3 | `T0P50_Parcial_OctubreAse3R1Occidente_UnaSolaFilaTotal_TotalMedioSinFila` |
| Aritmética A3 | `T0P50_A3_R1_…`, `T0P50_A3_R2_…`, `T0P50_A3_R4_…` |
| Etiquetas no reconocidas | `T0P50_EtiquetasNoReconocidas_…`, `T0P50_EtiquetasNoReconocidas_NoCoincidenConAusenciasEnElMismoReporte` |
| Residuo Q2 | `T0P50_Q2_Residuo_UnicaDivergenciaEsAusenciaR4PorEmpresa` |
