# T0 — Plan 60: reclasificación `Deb/Cred` → `Deb/Cred-otros` por canal (`Rem. Anticipos R2`)

**Fecha:** 2026-10-27 · **Fase:** H0/F0 (read-only) · **Plan:** `plans/60 - reclasificacion-deb-cred-otros-por-canal-rem-anticipos-r2.md`
**Fixture:** `Remuneracion.IntegrationTests/ReclasificacionDebCredT0Plan60Tests.cs` (5/5 verde)
**Naturaleza:** medición sobre el disco. Ningún archivo del repo se modificó para producir esta evidencia; el fixture abre los `.xlsx` en modo lectura.

---

## 0. Veredicto

La evidencia de §0 del plan **se reproduce en el disco**. El caso es uno solo (octubre-1 · ASE2 · canal `O`) y la divergencia declarada (§0.5) es exactamente la que el plan describe. Hay **una diferencia de redacción** en §0.3 (ASE1/ASE4/ASE5 no traen «celdas vacías» en `Deb/Cred`: traen importes pequeños bajo el umbral) que **no cambia ninguna conclusión ni el contrato** — se declara en §6 de este documento.

| Medición | Plan | Disco | Veredicto |
|---|---|---|---|
| Fuente R2 ASE2 · bloque | 5..57, headers en 4 | 5..57, headers en 4 | ✅ |
| Canal `O` (Vlr Servicio / Vlr Intereses / Total) | 54,00 / 6.759.294,81 / 6.759.348,81 | idéntico | ✅ |
| `OCCIDENTE` / `Componente` / gran `Total` | 6.759.342,66 / 6.759.371,77 / 6.759.371,77 | idéntico | ✅ |
| Otros 8 cierres de canal | 5,17 · 31,40 · 12,95 · 2,44 · 4,90 · −27,75 · −21,95 · 15,80 | idéntico | ✅ |
| Columna `Deb/Cred-otros` en la fuente | ausente | ausente (todos los R vacíos) | ✅ |
| Etiqueta `O` en la secuencia | 2 instancias | 2 instancias | ✅ |
| Oráculo: canal `O` | Q=0 / R=movido | Q=0 / R=movido | ✅ |
| Oráculo: `OCCIDENTE` / `Componente` / `Total` | −6,15 / 22,96 / 22,96 en Q | idéntico | ✅ |
| Salida actual (§0.5) | sin mover, R=0 | sin mover, R=0 en todo el bloque | ✅ |
| Fórmulas del cuadro OPORTUNO en la v4 | 36 celdas (Q/R de 18 filas) | exactamente esas 36 | ✅ |
| Histórico julio-1 / julio-2 / agosto-2 | ningún canal sobre el umbral | ningún canal sobre el umbral | ✅ |
| `Reversion Pagos R4` | sin `Deb/Cred` | sin `Deb/Cred` | ✅ |
| Huella semántica v4 | `a7476ee2…8c55` | `a7476ee2…8c55` | ✅ |

---

## 1. T0.1 — Fuente R2 del ASE2 (octubre-1)

Ruta (nombre exacto, verificado):

```
Docs/Prueba Octubre-1/Insumos/2-Lime/RerpoteDetalleSaldosaFavor_to_date01102026ddMMyyyy_to_date15102026ddMMyyyy___2026101617441360.xlsx
```

Hoja `Sheet1`; headers de componente en la fila **4** (`Deb/Cred` en la columna que la propia fila declara); bloque de detalle **5..57** (la fila 57 es `A=Total`).

| Fila | Firma A\|B\|C\|D | `Deb/Cred` medido |
|---|---|---|
| 26 | `\| \|O\|Vlr Servicio` | 54,00 |
| 27 | `\| \|O\|Vlr Intereses` | 6.759.294,81 |
| 28 | `\| \|O\|Total` (cierre del canal O) | 6.759.348,81 |
| 32 | `\|OCCIDENTE\|Total\|` (empresa) | 6.759.342,66 |
| 33 | `Componente\|Total\|\|` (sección) | 6.759.371,77 |
| 57 | `Total\|\|\|` (gran total) | 6.759.371,77 |

Cierres de canal del cuadro `Componente`, medidos: `A` 5,17 · `E` 31,40 · `1` 12,95 · `2` 2,44 · `3` 4,90 · `5` −27,75 · `H` −21,95 · `O` 6.759.348,81 · `T` 15,80. **Un único cierre supera el umbral** (`|Q| > 500.000`): el del canal `O`.

Hechos estructurales confirmados:
- La fuente **no trae** `Deb/Cred-otros` (no hay encabezado y las 53 filas del bloque tienen la celda vacía) ⇒ el `0` de `R` es observado, no inventado.
- La etiqueta `O` aparece **dos veces** (cuadro `Componente` y cuadro `Subs/Cont`): la identidad tiene que ser la **instancia**, nunca la etiqueta.

---

## 2. T0.2 — Oráculo administrativo, **por etiqueta**

Ruta: `Docs/Prueba Octubre-1/Resultados/Resultado administrativo/Remuneracion 202610-1 Total.xlsx`, hoja `Rem. Anticipos R2`.

| Etiqueta (nivel) | `Deb/Cred` | `Deb/Cred-otros` |
|---|---|---|
| canal `O` · `Vlr Servicio` | 0 | 54,00 |
| canal `O` · `Vlr Intereses` | 0 | 6.759.294,81 |
| canal `O` · `Total` | 0 | 6.759.348,81 |
| `B=OCCIDENTE C=Total` | −6,15 | 6.759.348,81 |
| `A=Componente B=Total` | 22,96 | 6.759.348,81 |
| `A=Total` (gran total) | 22,96 | 6.759.348,81 |
| cuadro OPORTUNO · `OCCIDENTE` | −6,15 | 6.759.348,81 |
| cuadro OPORTUNO · `TOTAL` | 22,96 | 6.759.348,81 |

**Geometría (normativo §0.6).** La fila de `OCCIDENTE` del bloque **no** coincide entre el oráculo y la plantilla vigente: el fixture mide ambas y exige que sean **distintas**. Por eso toda comparación va **por etiqueta/nivel**, nunca por número de fila. Un test que asertara `Q90` o `R116` sería un quemado de coordenada.

Las dos últimas filas del cuadro OPORTUNO son **fórmula** en la v4 ⇒ no se escriben: se recalculan solas (punto 4 de la regla).

---

## 3. T0.3 — Salida actual (§0.5) y fórmulas de la v4

Salida pre-regla: `Docs/Prueba Octubre-1/Resultados/5/Remuneración 202610-1 Total.xlsx`.

| Etiqueta | `Deb/Cred` | `Deb/Cred-otros` |
|---|---|---|
| canal `O` (3 filas) | 54,00 · 6.759.294,81 · 6.759.348,81 (**sin mover**) | 0 |
| `OCCIDENTE` | 6.759.342,66 | 0 |
| `Componente` | 6.759.371,77 | 0 |

`Deb/Cred-otros` = 0 en **todas** las filas del bloque: esa tabla **es** la divergencia declarada que el plan cierra.

**Fórmulas del cuadro OPORTUNO en la plantilla vigente (36 celdas = Q/R de 18 filas):**

```
Q57/R57 · Q63/R63 · Q69/R69                                  (cuadro del ASE1)
Q178/R178 · Q181/R181 · Q184/R184 · Q190/R190                (ASE2)
Q299/R299 · Q305/R305 · Q311/R311                           (ASE3)
Q423/R423 · Q426/R426 · Q429/R429 · Q435/R435               (ASE4)
Q510/R510 · Q516/R516 · Q519/R519 · Q522/R522               (ASE5)
```

El fixture exige que el conjunto medido sea **exactamente** ese, y que las celdas del par dentro del bloque de detalle **no** tengan `<f>` (son celdas de valor que el writer sí escribe).

---

## 4. T0.4 — Regresión histórica y alcance R4

| Período | Resultado medido |
|---|---|
| `2026071` (julio-1) | ningún cierre de canal supera el umbral en ninguno de los 5 ASE |
| `2026072` (julio-2) | idem |
| `2026082` (agosto-2) | idem |
| `2026101` (octubre-1) | dispara **solo** ASE2 · canal `O` (6.759.348,81). ASE3 máximo 2.232,90. ASE1 160,05 · ASE4 98,04 · ASE5 124,42 (todos ≤ umbral) |
| `Reversion Pagos R4` | su bloque no declara `Deb/Cred` ni `Deb/Cred-otros` (última columna de valor: `Otros`) |

Los manuales históricos (julio-2 `Remuneracion 202607-2 Total Administrativo.xlsx` y agosto-2 `Remuneracion 202608-2 Total_7721.xlsx`) publican `Deb/Cred-otros = 0` en los cinco bloques de detalle.

---

## 5. T0.5 — Huella de la plantilla

`HuellaSemanticaPlantilla.Calcular(Docs/Plantilla estandarizada/v4/Formato Remuneracion 2026xx-x Total v4.xlsx)` = `a7476ee2e80dfa13c9e39430d1ea4d77f0c52e3ca0a3d11a526d5140b8608c55` — idéntica al gate de entrada del plan.

---

## 6. Discrepancias declaradas (no se «ajusta» el plan)

**D-1 · §0.3 dice «ASE1, ASE4 y ASE5: su R2 no trae valores en `Deb/Cred` (columna presente, celdas vacías)».**
Medido en el disco (máximo `|Deb/Cred|` de los **cierres de canal** de octubre-1):

| ASE | Máx. `\|Deb/Cred\|` en cierres de canal |
|---|---|
| 1 · PROMOAMBIENTAL | 160,05 |
| 2 · LIME | 6.759.348,81 |
| 3 · CIUDAD LIMPIA | 2.232,90 |
| 4 · BOGOTA LIMPIA | 98,04 |
| 5 · ÁREA LIMPIA | 124,42 |

Es decir: **sí traen valores**, todos muy por debajo del umbral. La conclusión del plan —«en octubre-1 la regla dispara una sola vez (ASE2 · canal O)»— **se conserva intacta**, y el umbral sigue teniendo un único disparo. Lo que queda desmentido es solo la palabra «vacías». El T0 fija los cuatro números medidos (no el «0») y deja la nota escrita; **el plan no se edita**.

No se detectó ninguna otra diferencia. Ninguna geometría del plan quedó desmentida.

---

## 7. Evidencia de la corrida real (anexo de cierre H2/H3)

Corridas CLI `--cinco-ase` con la plantilla **v4** (rutas temp frescas, `git` intacto):

| Período | Exit | Veredicto | `DetRetri!D9:D14` |
|---|---|---|---|
| `2026101` | 0 | `RESULTADO OK` · `VALIDACION_TOTAL O9=0; P9=TRUE` | 14.842.425.136 / 11.302.637.605 / 8.823.787.300 / 10.483.172.762 / 7.550.668.665 / **53.002.691.468** |
| `2026071` | 0 | `RESULTADO OK` | sin cambios (diff = 0 celdas) |
| `2026072` | 0 | `RESULTADO OK` | sin cambios (diff = 0 celdas) |
| `2026082` | 0 | `RESULTADO OK` | sin cambios (diff = 0 celdas) |

**Blast radius medido** (`2026101`, diff celda a celda de las 40 hojas entre la salida **pre**-regla y la **post**-regla — 12 celdas, todas en `Rem. Anticipos R2`, todas del par):

```
Q110 54 → 0                    R110 0 → 54
Q111 6.759.294,81 → 0          R111 0 → 6.759.294,81
Q112 6.759.348,81 → 0          R112 0 → 6.759.348,81
Q116 6.759.342,66 → −6,15      R116 0 → 6.759.348,81
Q117 6.759.371,77 → 22,96      R117 0 → 6.759.348,81
Q147 6.759.371,77 → 22,96      R147 0 → 6.759.348,81
```

**Regresión** (`2026071` / `2026072` / `2026082`): **0 celdas distintas** en todo el libro (40 hojas, valores + texto de fórmula) entre la salida pre-regla y la post-regla.

**Traza de auditoría en el log** (`remuneracion_log_20261027.txt`, RunId `930271f3-eca0-4c5b-aa5e-3b3064557970`): una línea `INF` de resumen («3 concepto(s) movidos por 13.518.697,62 y 3 agregado(s) reconciliados») y seis líneas `DBG` con la coordenada completa —`Componente · OCCIDENTE · O · Vlr Servicio / Vlr Intereses / Total` y los agregados `OCCIDENTE` (`H,O,T`), `Componente` y gran `Total`.
