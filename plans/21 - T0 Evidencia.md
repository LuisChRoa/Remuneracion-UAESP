# Plan 21 — T0 Evidencia (Fase 0, PR 1 bloqueante)

> Documento de evidencia del Discovery T0 del `plans/21 - espejo-estructural-reporte-componentes-r1.md`
> (§2.1, §2.3, §3.1 R-T0-*, §4 T0). **Solo evidencia: cero cambios de código productivo o de tests.**
> Toda afirmación remite a disco (archivo/hoja/fila). Método reproducible al final (§9).
>
> **Fecha:** 2026-10-02
> **Insumos Q1/Q2:** `Docs/Insumos/REMUNERACION 2026071`, `Docs/Insumos/REMUNERACION 2026072`,
> `Remuneracion 202607-1 Total.xlsx`, `Remuneracion 202607-2 Total.xlsx`.
> **Caso de regresión:** `Docs/Prueba2/Insumos` (período 2026082 Q2; 5 carpetas ASE + `R10_Remuneracion_2026082.xlsx`).
> **Cachés golden usados:** copias en `%TEMP%\opencode\` (los originales en `Docs/Insumos/` pueden
> estar tomados por Excel).

---

## 0. Conclusión de una línea

El espejo estructural R1 **se sostiene 1:1 en 10/10 combinaciones julio** (conteos 38/52/46/69/44
y 45/75/43/73/52 reproducidos desde disco) y R4 es espejo **de fila** 1:1 en 10/10; la causa raíz
de agosto queda probada en disco (`TotalD_E` ocurrencia 2 ausente en ASE1); el reanclaje de la
ruta A **es viable** (R3 no se activa) pero su superficie es **mucho mayor** que la hipótesis de
trabajo del plan (miles de referencias cruzadas). No hay `NEEDS_CONTEXT` de T0a/T0b; sí hay 2
decisiones bloqueantes de alcance (ver §8).

---

## 1. Causa raíz de agosto (evidencia dura, S3)

El fallo `ERR-VALIDACION` de 2026082 es **estructural y reproducible desde disco**. El mapa
`WorkbookLeafCellMapInterventoria.LMenoresPorAseQ2[1]` exige `L20 = RolLMenor.TotalD_E, ocurrencia 2`
(0-based) — leído del golden julio-Q2 (`Remuneracion 202607-2 Total.xlsx`).

| Período | ASE1 `TotalD_E` (occs válidas) | L20 pide occ | Resultado |
|---|---|---|---|
| Q2-julio (`REMUNERACION 2026072/1-Promoambiental/Recaudoporcomponente_*_16072026_*.xlsx`) | **3** (0,1,2) | 2 | OK |
| agosto-Prueba2 (`Docs/Prueba2/Insumos/1-Promoambiental/Recaudoporcomponente_*_16082026_*.xlsx`) | **2** (0,1) | 2 | **slot ausente → `ERR-VALIDACION`** |

La fuente de agosto pierde el bloque `TotalD_E` (ENEL) adicional que sí traen Q1/Q2. Complementario:
Q1 ASE1 tiene `TotalD_E=1` y en Q1 el mapa (`LMenoresPorAse[1]`, roll Q1) pide `L13 = TotalD_E occ 0`
→ por eso Q1 no rompe aunque su cardinalidad también difiere de Q2. **La ocurrencia congelada es
el defecto, no el dato.**

---

## 2. T0a — Dumps fuente-vs-destino `Recaudoporcomponente` ↔ `Reporte Componentes R1`

Método: lectura OpenXML de `Recaudoporcomponente_*.xlsx` (hoja `Sheet1`) como secuencia de firmas
`A|B|C|D|E`, y del bloque espejo del golden (`Reporte Componentes R1`) desde la fila del nombre del
ASE (`B=<nombre ASE>`) hasta la primera fila `A='Total'` con `B` vacío. Comparación posición a
posición (case-sensitive) y delta por par alineado.

### 2.1 Conteos y alineación (reproducidos desde disco)

| Período | ASE | Fuente filas | Destino filas | 1:1 | delta (dst−src) | nameRow | totalRow |
|---|---|---:|---:|:--:|---:|---:|---:|
| Q1-julio | 1 PROMOAMBIENTAL | 38 | 38 | Sí | +1 | 3 | 43 |
| Q1-julio | 2 LIME | 52 | 52 | Sí | +76 | 79 | 132 |
| Q1-julio | 3 CIUDAD LIMPIA | 46 | 46 | Sí | +206 | 209 | 256 |
| Q1-julio | 4 BOGOTA LIMPIA | 69 | 69 | Sí | +346 | 349 | 419 |
| Q1-julio | 5 AREA LIMPIA | 44 | 44 | Sí | +467 | 470 | 515 |
| Q2-julio | 1 PROMOAMBIENTAL | 45 | 45 | Sí | +1 | 3 | 50 |
| Q2-julio | 2 LIME | 75 | 75 | Sí | +83 | 86 | 162 |
| Q2-julio | 3 CIUDAD LIMPIA | 43 | 43 | Sí | +236 | 239 | 283 |
| Q2-julio | 4 BOGOTA LIMPIA | 73 | 73 | Sí | +373 | 376 | 450 |
| Q2-julio | 5 AREA LIMPIA | 52 | 52 | Sí | +498 | 501 | 554 |

- **Q1 = 38/52/46/69/44 coincide exacto con lo aportado y medido.** ✅
- **Q2 = 45/75/43/73/52 coincide exacto.** ✅
- El **delta es constante dentro de cada bloque** (no es siempre +1): +1 solo para ASE1; en ASE2..5
  acumula el tamaño de las secciones previas del destino. El `+1 por encabezado de bloque` que
  menciona el plan es la relación intra-bloque del primer bloque; la regla real es *"dstRow − srcRow
  constante por bloque"*. La secuencia de firmas A–E es **idéntica** fuente↔destino en las 10
  combinaciones (ver reproducibilidad §6, `compare-r1.ps1`).

### 2.2 Agosto-Prueba2 (solo fuente; sin destino en disco)

En `Docs/Prueba2` **no existe destino** `Reporte Componentes R1`: `Docs/Prueba2/Resultado/` está
vacío y `Docs/Prueba2/Insumos/R10_Remuneracion_2026082.xlsx` es un workbook **solo con las hojas
`DetRetri2026082` y `DetValiRetri2026082`** (verificado por OpenXML), no la plantilla. La plantilla
de destino la elige el usuario en la UI (`Form1.txtPlantilla`) — ver §7.

| Período | ASE | Fuente filas | Destino |
|---|---|---:|---|
| agosto-Prueba2 | 1 PROMOAMBIENTAL | 42 | sin artefacto en disco |
| agosto-Prueba2 | 2 LIME | 66 | sin artefacto en disco |
| agosto-Prueba2 | 3 CIUDAD LIMPIA | 37 | sin artefacto en disco |
| agosto-Prueba2 | 4 BOGOTA LIMPIA | 79 | sin artefacto en disco |
| agosto-Prueba2 | 5 AREA LIMPIA | 60 | sin artefacto en disco |

### 2.3 Inventario de roles por ASE (zona Componente / lecturas del reader)

Conteo de roles que resuelve `ExcelDataReaderWorkbookLeafInputReader` (firmas col D/E) en la zona
Componente (filas anteriores a `A='Componente' B='Total'`). Es la tabla que explica la fragilidad.

| Período | ASE | VlrServicio | VlrIntereses | TotalD_E | TotalD_H | TotalD_O | TotalD_T |
|---|---|---:|---:|---:|---:|---:|---:|
| Q1 | 1 | 5 | 5 | 1 | 1 | 1 | 2 |
| Q1 | 2 | 9 | 6 | 2 | 1 | 2 | 1 |
| Q1 | 3 | 7 | 5 | 2 | 1 | 2 | 2 |
| Q1 | 4 | 11 | 8 | 3 | 1 | 3 | 2 |
| Q1 | 5 | 7 | 5 | 3 | 1 | 2 | 1 |
| Q2 | 1 | 7 | 4 | **3** | 1 | 2 | 1 |
| Q2 | 2 | 12 | 10 | 3 | 1 | 3 | 2 |
| Q2 | 3 | 6 | 5 | 2 | 1 | 2 | 1 |
| Q2 | 4 | 12 | 9 | 3 | 1 | 3 | 1 |
| Q2 | 5 | 8 | 7 | 2 | 2 | 2 | 2 |
| agosto | 1 | 6 | 4 | **2** | 1 | 2 | 1 |
| agosto | 2 | 12 | 6 | 3 | 2 | 3 | 1 |
| agosto | 3 | 6 | 4 | 2 | 1 | 2 | 1 |
| agosto | 4 | 13 | 10 | 3 | 1 | 3 | 1 |
| agosto | 5 | 9 | 8 | 2 | 2 | 3 | 2 |

**Veredicto T0a: espejo 1:1 CONFIRMADO** (R-T0-1 cumplido). La fuente del período define la
secuencia; las ocurrencias congeladas del mapa T0-0.5 quedan refutadas como estrategia.

---

## 3. T0b — Veredicto espejo R4 (`ReversiónPorComponente` ↔ `Reversion Pagos R4`)

Método: firma estructural `A|B|C` (las columnas D/E de R4 son valores, no etiquetas), desde la
fila de nombre del bloque hasta `A='Total'`.

| Período | ASE | Fuente filas | Destino filas | 1:1 (A/B/C) | delta (dst−src) | nameRow | totalRow |
|---|---|---:|---:|:--:|---:|---:|---:|
| Q1 | 1 | 7 | 7 | Sí | −2 | 1 | 9 |
| Q1 | 2 | 12 | 12 | Sí | +82 | 84 | 98 |
| Q1 | 3 | 13 | 13 | Sí | +176 | 178 | 193 |
| Q1 | 4 | 19 | 19 | Sí | +213 | 215 | 236 |
| Q1 | 5 | 13 | 13 | Sí | +327 | 329 | 344 |
| Q2 | 1 | 13 | 13 | Sí | −2 | 1 | 15 |
| Q2 | 2 | 13 | 13 | Sí | +88 | 90 | 105 |
| Q2 | 3 | 13 | 13 | Sí | +183 | 185 | 200 |
| Q2 | 4 | 19 | 19 | Sí | +220 | 222 | 243 |
| Q2 | 5 | 13 | 13 | Sí | +335 | 337 | 352 |
| agosto | 1 | 6 | — | — | — | — | — |
| agosto | 2 | 13 | — | — | — | — | — |
| agosto | 3 | 7 | — | — | — | — | — |
| agosto | 4 | 18 | — | — | — | — | — |
| agosto | 5 | 13 | — | — | — | — | — |

**Matiz obligatorio (no inventar 1:1 de celdas):** el espejo R4 es **de fila/bloque**, no de
columnas. El destino `Reversion Pagos R4` tiene `P=SERVICIO ESPECIALES` (valor 0) que **no tiene
contraparte en la fuente**; la fuente trae `F=Componente TTL` y `P` vacío. El mapeo de columnas
sigue siendo el del mapa HU-07/HU-12 (`D←Total`, `P=0`). El espejo solo garantiza **misma cantidad
y mismo orden de filas/labels A/B/C**.

**Veredicto T0b: R4 (bloque principal `Reversion Pagos R4`) ENTRA a esta HU** a nivel de forma de
bloque (filas), en ambos períodos, 10/10. El **R4-por-empresa Q2** (divergencia documentada en el
recorte HU-20/G2-D1, riesgo R5) queda **fuera** y pasa a follow-up.

---

## 4. T0c — Inventario de fórmulas por bloque + veredicto A/B

### 4.1 Fórmulas internas de `Reporte Componentes R1`

Celdas con fórmula (masters) por sección-ASE:

| Período | ASE1 | ASE2 | ASE3 | ASE4 | ASE5 |
|---|---:|---:|---:|---:|---:|
| Q1 | 190 | 159 | 203 | 215 | 148 |
| Q2 | 227 | 258 | 206 | 242 | 222 |

Familias (Q1, textos exactos del golden):

| Familia | Celda(s) | Fórmula(s) representativas |
|---|---|---|
| TOT_OPT visible | F46, F176, F316, F437, F519 | `F46=F25+F41-L25`; `F176=F113+F130-L113+F90-L90`; `F316=F238+F254+F217-L217-L238`; `F437=F391+F417+F357-L357-L391`; `F519=F513+F498+F478-L478-L498` |
| EXTEMP visible | F48, F178, F318, F439, F521 | `F48=F30+F10-L10`; `F318=F243+F223-L223`; `F439=F401+F369-L369`; **F178/F521 = valor 0 sin fórmula** (HU-07 T0-0.2) |
| OPORTUNO por empresa | F56, F66, F68, F181, F186, F196, F326, F336, F338, F447, F449, F457, F459, F529, F534, F539 | `F56=F33+F14-L14`; `F68=F29+F9-L9`; `F181=F102+F122+F86-L86-L102`; `F186=F117+F95+F83-L83-L95`; `F196=F129+F112+F89-L89-L112` |
| AFaseo | F206, F346, F467 | `F206=F91`; `F346=F217`; `F467=F357` |
| Resto de columnas (G..AP) | p.ej. `G46=G25`, `G176=G113+G90`, `G519=G498+G478` | fórmulas por columna que referencian filas del bloque/sección |

La totalidad del volcado (915 celdas-fórmula en R1-Q1) está en disco (§9). Todas las referencias
internas son **A1 relativas/absolutas simples**; no hay funciones volátiles.

### 4.2 Referencias externas a R1/R4 (dependencias aguas abajo)

Conteo de ocurrencias explícitas de referencia en texto de fórmula (masters; los followers
`t="shared"` sin texto no suman — sesgo conservador):

| Período | Refs externas a R1 | Refs externas a R4 | Refs R1 por hoja (top) |
|---|---:|---:|---|
| Q1 | 1061 | 556 | DetValiRetri=670, DetRetri=260, CONSOLIDADO=220, REMUNERACION_* (5 hojas ≈220 c/u; EAAB-CL 161), Informe AFaseo=55 |
| Q2 | 1061 | 556 | DetValiRetri=670, DetRetri=260, CONSOLIDADO=220, REMUNERACION_* (idem) |

### 4.3 Nombres definidos y funciones especiales

- Único nombre definido que apunta a R1: **`PROMOAMBIENTAL_1 = 'Reporte Componentes R1'!$A$5:$X$23`**,
  que cae **dentro del bloque de datos de ASE1**. **No se usa en ninguna fórmula** (0 usos), pero
  debe auditarse/reanclarse por doctrina.
- `OFFSET`, `INDIRECT`, `INDEX`, `MATCH`, `VLOOKUP`, `HLOOKUP`: **0 ocurrencias**.
- Enlaces externos `[...]` dentro de fórmulas: **0**. `#REF!` existentes: **3** (`X346`, `Y346`, `Z346`).
- Fórmulas con referencias cruzadas entre hojas: **mecánicas A1** (`'Hoja'!F46`), sin rangos con nombre en uso.

### 4.4 Veredicto A/B (R-T0-3)

**VEREDICTO: A (mutar filas + reanclar fórmulas). R3 NO se activa; T-alt NO se activa.**

Justificación basada en evidencia:
- El reanclaje es **mecánico** (solo referencias A1; sin volátiles, sin rangos con nombre en uso,
  sin vínculos externos). No hay intratabilidad técnica.
- Pero la superficie es **mucho mayor** que la hipótesis del plan: ~450 refs externas por sección-ASE
  en R1 y ~163 en R4, más 150–260 fórmulas internas por sección, más el nombre `PROMOAMBIENTAL_1`.
- Condición para A: **paso dedicado de reanclaje** (formulas + definedNames) con red de goldens
  5-ASE y test de regresión 2026082. Si el Ingeniero prioriza reducir blast-radius sobre preservar
  la plantilla, T-alt sigue disponible, pero **también** tendría que reescribir las referencias
  externas (los visibles cambian de fila), por lo que no elimina el costo principal.

---

## 5. T0d — Inventario de celdas absolutas aguas abajo por bloque

### 5.1 Secciones y refs externas por ASE

Sección = desde la fila de nombre del ASE hasta la víspera del siguiente (incluye datos + resumen).
Refs externas = ocurrencias que apuntan a celdas de esa sección.

| Período | ASE | Sección R1 | Refs externas R1 | Sección R4 | Refs externas R4 |
|---|---|---|---:|---|---:|
| Q1 | 1 | R3..R78 | 447 | R1..R83 | 164 |
| Q1 | 2 | R79..R208 | 451 | R84..R177 | 163 |
| Q1 | 3 | R209..R348 | 450 | R178..R214 | 163 |
| Q1 | 4 | R349..R469 | 448 | R215..R328 | 163 |
| Q1 | 5 | R470..R549 | 450 | R329..R368 | 163 |
| Q2 | 1 | R3..R85 | 447 | R1..R89 | 164 |
| Q2 | 2 | R86..R238 | 445 | R90..R184 | 163 |
| Q2 | 3 | R239..R375 | 456 | R185..R221 | 163 |
| Q2 | 4 | R376..R500 | 448 | R222..R336 | 163 |
| Q2 | 5 | R501..R588 | 450 | R337..R376 | 163 |

Desglose Q1 ASE1 (representativo): DetValiRetri=134, DetRetri=52, CONSOLIDADO=44,
REMUNERACION_ENEL=44, REMUNERACION_OCCIDENTE=44, REMUNERACION_Reciprocidad=44,
REMUNERACION_ENERBIT=43, REMUNERACION_EAAB-CL=31, Informe AFaseo=11.

**Tratamiento del desplazamiento apilado (§2.3):** mutar el bloque del ASE-N desplaza
`Σ(refs de secciones N..5)`. Por eso T4 debe procesar **5→1** o recomputar offsets acumulados
(§2.3 punto 2). Con la tabla anterior el costo por bloque es calculable sin adivinar.

### 5.2 Celdas absolutas de los mapas que caen dentro/debajo del bloque

| Mapa | Hoja | Celdas absolutas (por bloque) | Relación con el bloque |
|---|---|---|---|
| `WorkbookLeafCellMapPorAse` (Q1) — `EditableLeafCellsPorAse` | R1 | ASE1 `F25,F41,L25,F30,F10,L10`; ASE2 `F113,F130,L113,F90,L90`; ASE3 `F238,F254,F217,L217,L238,F243,F223,L223`; ASE4 `F391,F417,F357,L357,L391,F401,F369,L369`; ASE5 `F513,F498,F478,L478,L498` | **Dentro** del bloque (destino de escritura); se mueven con la fila |
| `WorkbookLeafCellMapPorAse` — `ProtectedFormulasPorAse` | R1 | ASE1 `F46,F48`; ASE2 `F176`; ASE3 `F316,F318`; ASE4 `F437,F439`; ASE5 `F519` | **Debajo** del bloque (visibles); se mueven y reanclan |
| `WorkbookLeafCellMapQ2` — `R1Q2EditablesPorAse` / `R1Q2ProtectedPorAse` | R1 | ASE1 `F12,L12,F32,L32,F48,F37,F17,L17` / `F53,F55`; ASE2 `F94..F160` / `F206,F208`; ASE3 `F244..F281` / `F343,F345`; ASE4 `F384..F448` / `F468,F470`; ASE5 `F531..F552` / `F558,F560` | Idem Q1 en Q2 |
| `WorkbookLeafCellMapInterventoria` — `LMenoresPorAse` / `LMenoresPorAseQ2` | R1 | 8–12 celdas `L*` por ASE (columna L) resueltas por rol+ocurrencia | **Dentro/debajo** del bloque (las L caen en filas del bloque/sección) |
| `WorkbookLeafCellMapPorEmpresa` | R1 | `EditablesR1PorEmpresa` (p.ej. ASE1 `F33,F14,L14,F40,F24,L24,F29,F9,L9`) + `REMUNERACION_*` D9:D13/D47:D51 | Dentro (operandos) y refs externas (visibles) |
| `WorkbookLeafCellMapValidaciones` / `WorkbookLeafCellMapQ2` | DetRetri / DetValiRetri | D9:D13 (`Remuneracion`) / D23:D28 / D16:D21 | Referencian visibles R1 (`F46..`, `F176..`) vía CONSOLIDADO |

Notas de evidencia:
- El nombre definido `PROMOAMBIENTAL_1` (`$A$5:$X$23`, ASE1) es la única celda absoluta **no-mapa**
  que cae dentro de un bloque y **debe entrar al inventario de reanclaje de nombres**.
- La totalidad exhaustiva de celdas por mapa vive en los `.cs` citados; esta tabla ubica **qué
  cae dentro/debajo** de cada bloque y cuantifica las refs externas (§5.1).

**Veredicto T0d: inventario publicado** (R-T0-4 cumplido); §2.3 puntos 1–2 quedan implementables
sin adivinar.

---

## 6. T0e — Tabla espejo congelada + invariantes de cierre

### 6.1 Tabla espejo congelada (julio; agosto solo fuente)

| Período | ASE | nameRow | totalRow | Fuente | Destino | delta | Invariantes |
|---|---|---:|---:|---:|---:|---:|---|
| Q1 | 1 | 3 | 43 | 38 | 38 | +1 | Comp, Subs, Total |
| Q1 | 2 | 79 | 132 | 52 | 52 | +76 | Comp, Subs, Total (+Mes/AFaseo) |
| Q1 | 3 | 209 | 256 | 46 | 46 | +206 | Comp, Subs, Total (+Mes/AFaseo) |
| Q1 | 4 | 349 | 419 | 69 | 69 | +346 | Comp, Subs, Total (+Mes/AFaseo) |
| Q1 | 5 | 470 | 515 | 44 | 44 | +467 | Comp, Subs, Total (+Mes) |
| Q2 | 1 | 3 | 50 | 45 | 45 | +1 | Comp, Subs, Total (+Mes/AFaseo) |
| Q2 | 2 | 86 | 162 | 75 | 75 | +83 | Comp, Subs, Total (+Mes/AFaseo) |
| Q2 | 3 | 239 | 283 | 43 | 43 | +236 | Comp, Subs, Total (+Mes/AFaseo) |
| Q2 | 4 | 376 | 450 | 73 | 73 | +373 | Comp, Subs, Total (+Mes/AFaseo) |
| Q2 | 5 | 501 | 554 | 52 | 52 | +498 | Comp, Subs, Total (+Mes) |
| agosto | 1 | — | — | 42 | — | — | Comp, Subs, Total (+Mes/AFaseo) |
| agosto | 2 | — | — | 66 | — | — | Comp, Subs, Total (+Mes/AFaseo) |
| agosto | 3 | — | — | 37 | — | — | Comp, Subs, Total (+Mes/AFaseo) |
| agosto | 4 | — | — | 79 | — | — | Comp, Subs, Total (+Mes/AFaseo) |
| agosto | 5 | — | — | 60 | — | — | Comp, Subs, Total (+Mes/AFaseo) |

### 6.2 Invariantes de cierre observadas en las 3 muestras (15 combinaciones)

| Invariante (firma) | Presencia | Conclusión |
|---|---|---|
| `A='Componente' B='Total'` | 15/15, exactamente 1 vez | **Invariante dura** (corte de zona Componente) |
| `A='Subs/Cont' B='Total'` | 15/15, exactamente 1 vez | **Invariante dura** (cierre Subs/Cont) |
| `A='Total' B` vacío (fila Total final) | 15/15, exactamente 1 vez | **Invariante dura** (cierre de bloque) |
| `B='Mes' C='Total'` | 2 o 3 veces según ASE/período | **NO invariante** (la forma la define la fuente) |
| `A='AFaseo' B='Total'` | 0 o 1 vez | **NO invariante** (ASE1/ASE5 sin AFaseo en julio) |

Regla de ausencia resultante (D-B / R-E-5): fila ausente = **suprimida**; si falta una invariante
dura → fail-fast nombrando ASE+reporte+fila esperada. Ejemplos medidos: ASE1-Q1 `Mes=2, AFaseo=0`;
ASE1-Q2 `Mes=3, AFaseo=1`; ASE5-Q1/Q2 `Mes=2, AFaseo=0`; ASE5-agosto `Mes=3, AFaseo=1`.

**Veredicto T0e: tabla congelada publicada; solo `Componente/Total`, `Subs/Cont/Total` y `Total`
final son invariantes de cierre.**

---

## 7. Veredictos T0a..T0e (resumen)

| ID | Entregable | Veredicto | Estado |
|---|---|---|---|
| T0a | Espejo R1 fuente↔destino (Q1/Q2/agosto) | 1:1 CONFIRMADO 10/10; conteos 38/52/46/69/44 y 45/75/43/73/52 reproducidos; agosto fuente 42/66/37/79/60 (sin destino en disco); causa raíz `TotalD_E` probada | DONE |
| T0b | Veredicto espejo R4 | R4 bloque principal **ENTRA** (espejo de fila A/B/C 10/10); R4-por-empresa Q2 **follow-up**; columnas no 1:1 (documentado) | DONE |
| T0c | Inventario de fórmulas + veredicto A/B | **A** (mutar + reanclar); R3 no se activa; reanclaje mecánico pero de gran volumen (~450 refs externas/sección R1) | DONE |
| T0d | Celdas absolutas aguas abajo por bloque | Inventario publicado por sección + mapas; nombre `PROMOAMBIENTAL_1` dentro de ASE1 | DONE |
| T0e | Tabla espejo congelada + invariantes | Tabla por ASE + 3 invariantes duras; `Mes`/`AFaseo` NO invariantes | DONE |

No hay `NEEDS_CONTEXT` de T0a/T0b (el 1:1 no se refutó). Quedan 2 decisiones bloqueantes (§8).

---

## 8. Preguntas bloqueantes (antes de T1)

1. **Plantilla destino de 2026082:** en `Docs/Prueba2` no hay plantilla (`Resultado/` vacío; `R10_*`
   es solo DetRetri). La UI exige seleccionar plantilla. ¿Cuál se usa como plantilla de destino del
   test de regresión 2026082 (copia del golden Q2 de julio, u otra)? Sin esto, R-R-2 no es ejecutable.
2. **Alcance de la ruta A dado el volumen:** T0c confirma A técnicamente viable, pero el reanclaje
   mueve ~2500 refs externas R1 + ~800 R4 + ~1000 fórmulas internas + 1 nombre definido. ¿El
   Ingeniero confirma A con un PR dedicado al reanclaje, o prefiere evaluar T-alt (que no elimina el
   costo de reescribir las refs externas, pero cambia el modelo de la hoja R1)?

No continúo con T1 hasta respuesta.

---

### 8.1 Riesgos declarados y working list (auditoría code-review PR3)

Estos puntos NO cambian el alcance de este pase; se documentan como limitaciones conscientes para
que el próximo T0/PR no los trate como hallazgos nuevos.

- **R-DELTA-NEGATIVO (W-3) — riesgo declarado para agosto.** Con Δ<0 (el bloque destino es más
  largo que la fuente del período), las referencias A1 que apuntaban a una fila suprimida dentro
  de `[primeraBorrada, totalRow]` no se redirigen a "su" dato (ya no existe): el mutador
  (`OpenXmlEspejoR1Mutador.AjustarBloque`, rama Δ<0) las reancla a `primeraBorrada` = la fila
  `Total` del bloque del período, garantizando que jamás quede un `#REF!`. Es un compromiso
  deliberado: la aritmética de negocio no se recalcula (fuera de alcance por D-A/§0.2) y Excel
  recalcula las fórmulas al abrir el workbook. Evidencia de agosto: Δ = -3/-9/-6 para ASE1/2/3
  (fuente 42/66/37 vs plantilla Q2 45/75/43), con 2428/1930/1445 referencias reancladas
  respectivamente (instrumentación W-2). A revisar en el T0 de agosto si alguna referencia a fila
  suprimida es semánticamente relevante: las visibles (TOT_OPT/EXTEMP) viven debajo del bloque y
  solo se desplazan, no caen en el rango suprimido.

- **R-CF-DV (W-4) — cobertura pendiente de conditional formatting / data validations.** El espejo
  actual reancla fórmulas, `ref` de shared formulas, celdas combinadas y nombres definidos, pero
  NO toca ni verifica `conditionalFormatting` ni `dataValidations` cuyos `sqref` caigan dentro del
  bloque mutado. El T0 original no inventarió estas colecciones (no se inventan coberturas sin
  evidencia). Quedan como working list para el T0 de agosto, que debe volcar los `sqref`
  afectados por bloque antes de decidir si requieren reanclaje. **No hay cambio de código en este
  pase.**

- **W-2 — evidencia de reanclaje ya instrumentada (no es deuda).**
  `EspejoR1MutacionResultado` expone el Δ y las referencias reancladas por bloque;
  `EspejoR1MutadorTests.Agosto_Reanclaje_DeltasYReferenciasPorBloque_ContraExpectativaT0a` aserta
  Δ = -3/-9/-6/+6/+8 y el total de 7188 referencias contra la expectativa T0a/T0c, y verifica por
  spot-check que las referencias de CONSOLIDADO a los visibles de R1 no queden en fila equivocada
  (F53→F50, F558→F554). Reemplaza la cota arbitraria de fórmulas por igualdad exacta.


---

## 9. Reproducibilidad (método, todo en disco)

Scripts de medición en `%TEMP%\opencode\` (no versionados):

- `compare-r1.ps1` — compara secuencia fuente↔golden por ASE (da 10/10; Q1/Q2).
- `t0a.ps1` — conteos, alineación y delta por ASE; `-Mode R1` y `-Mode R4` para Q1/Q2/agosto.
- `t0roles.ps1 -Mode R1` — inventario de roles en la zona Componente por ASE/período.
- `dumpwb.ps1` — volcado OpenXML de hojas/fórmulas/filas de un `.xlsx`.
- `t0buckets.ps1` — secciones por ASE y refs externas por período/hoja.

Comandos (PowerShell 5.1, cwd = raíz del repo; los `.ps1` se invocan con `-File` y usan
`Get-Location` para evitar problemas de encoding de acentos):

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File "$env:TEMP\opencode\compare-r1.ps1"
powershell -NoProfile -ExecutionPolicy Bypass -File "$env:TEMP\opencode\t0a.ps1" -Mode R1
powershell -NoProfile -ExecutionPolicy Bypass -File "$env:TEMP\opencode\t0a.ps1" -Mode R4
powershell -NoProfile -ExecutionPolicy Bypass -File "$env:TEMP\opencode\t0roles.ps1" -Mode R1
powershell -NoProfile -ExecutionPolicy Bypass -File "$env:TEMP\opencode\t0buckets.ps1"
```

Archivos de evidencia cruda generados: `allf_q1.txt`, `allf_q2.txt`, `extref_r1_q1.txt`,
`extref_r1_q2.txt`, `extref_r4_q1.txt`, `extref_r4_q2.txt`, `r1_q1_formulas.txt`, `r1_q1_rows.txt`,
`r4_q1_rows.txt` (todos en `%TEMP%\opencode\`).

Insumos usados (nombres exactos por patrón):
- Q1: `Docs/Insumos/REMUNERACION 2026071/{1..5}-*/Recaudoporcomponente_*_01072026_*_15072026_*.xlsx`
  y `Reversi*PorComponente_*_01072026_*_15072026_*.xlsx`.
- Q2: `Docs/Insumos/REMUNERACION 2026072/{1..5}-*/Recaudoporcomponente_*_16072026_*_31072026_*.xlsx`
  y `Reversi*PorComponente_*_16072026_*_31072026_*.xlsx`.
- Agosto-Prueba2: `Docs/Prueba2/Insumos/{1..5}-*/Recaudoporcomponente_*_16082026_*_31082026_*.xlsx`
  y `Reversi*PorComponente_*_16082026_*_31082026_*.xlsx`.
- Goldens: cachés `%TEMP%\opencode\Remuneracion 202607-1 Total.xlsx` y `...-2 Total.xlsx`.
