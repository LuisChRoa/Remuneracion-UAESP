# T0 Plan 51 — Reporte banco: headers de fuente por ASE (H0 read-only)

> **Alcance:** H0 del `plans/51 - reporte-banco-fuentes-opcionales-por-header.md` (APROBADO 2026-10-22, Clarification Gate §7 ratificado 6/6).
> **Método:** lectura de BYTES con BCL (`System.IO.Compression.ZipArchive` + `System.Xml.Linq`), sin Excel/COM,
> sin OpenXml SDK y sin ExcelDataReader. Cero código de producción, cero escritura sobre fuentes ni plantillas
> (las fuentes abiertas se copian a `%TEMP%` antes de leer).
> **Fixture que congela este documento:** `Remuneracion.IntegrationTests/ReporteBancoHeadersT0Plan51Tests.cs` (105 verdes).
> **Muestras:** 20 (5 ASE × 4 períodos) de `ReportePagosxBanco_*.xlsx`, hoja `Sheet1`.
> **Fecha:** 2026-10-22. **Autor:** implementer (H0).

---

## 0. Veredicto H0 (puerta H0 → H1)

| Ítem | Veredicto |
|---|---|
| **S1.1** matriz de headers por ASE × período | **CONGELADA** (§1). Los headers observados en las 20 muestras son, en todos los casos, un **subconjunto** de las claves de `WorkbookLeafCellMapReporteBanco.EditablesPorAse[ase]`. |
| **S1.2** `Total` vs fuente | **CONGELADO**: exactamente **una** columna `Total` por header; **ninguna** fuente real normaliza a `total`; la exclusión por normalizado `== "total"` es suficiente y no colisiona. |
| **S1.3** conceptos 1/2/3/7 + fila `Total` | **CONGELADO**: se leen igual por prefijo normalizado; la fila `7-APLICADOS…` **no viene en ninguna** de las 20 muestras (0 demostrable); la fila `Total` existe en las 20 y aporta valor a todas las columnas de fuente. |
| **S1.4** Σ conceptos == `TotalFuente` | **CONGELADO**: Δ = `0,00` en las **40** combinaciones (empresa presente × muestra). El gate (i) no se altera. |
| **S1.5** ASE2 vs ASE4/ASE5 en octubre-1 | **CONGELADO con una desviación declarada**: ASE2 trae **2** fuentes (`NUEVO ESQUEMA` ausente) como esperaba el plan; **ASE4 también trae 2** (dato medido, el plan esperaba 3); ASE5 trae 3 = 3. Las dos ausencias son la **misma familia** (esperada no observada) y se absorben con la regla ratificada **sin ramas**. |
| **S1.6** headers no reconocidos | **VACÍO en las 20 muestras** → **no hay STOP**. |
| **S1.7** fixture read-only | **VERDE** (`ReporteBancoHeadersT0Plan51Tests`: 105/105, 0 skip). |
| **Puerta** | **H0 VERDE** → H1 habilitado. G-SHA: v1 `3086AC31…C5214` y v2 `0F192AD2…2DAB` verificadas antes de H1. |

**STOP no declarado.** No apareció ningún header desconocido, ninguna fuente reconocida sin celdas destino
(desync) y ninguna geometría no absorbible por la regla ratificada: *esperado = claves de `EditablesPorAse[ase]`;
observado ⊆ esperado; ausente = esperado − observado → se omite + huella; cero reconocidas → fail-fast*.

---

## 1. S1.1 — Matriz de headers por (período × ASE)

Criterio de localización (espejo del reader HU-09, `T0-0.7`): etiqueta `Resumen Recaudo Aplicado Por Servicio`
buscada **desde el final** del `Sheet1` por prefijo normalizado; la fila de headers es la inmediatamente siguiente;
los conceptos se localizan por prefijo normalizado en la columna `A`; la fila `Total` por normalizado `== "total"`.

Índice físico de columna **0-based** (`[1]` = columna `B`; la columna `A` es la etiqueta de concepto, nunca una fuente).
Los números de fila son **filas reales de Excel** (atributo `r` del XML), evidencia y no contrato.

| Período | ASE | `Resumen` (fila) | Headers (texto crudo · índice físico) | Presentes | Ausentes | Conceptos observados |
|---|---|---|---|---|---|---|
| 2026071 (jul-1) | 1 PROMOAMBIENTAL | 414 | `[1]=ENEL [2]=OCCIDENTE [3]=Total` | ENEL, OCCIDENTE | — | 1, 2, 3 |
| 2026071 | 2 LIME | 440 | `[1]=ENEL [2]=NUEVO ESQUEMA [3]=OCCIDENTE [4]=Total` | ENEL, NUEVO ESQUEMA, OCCIDENTE | — | 1, 2 |
| 2026071 | 3 CIUDAD LIMPIA | 406 | `[1]=ENEL [2]=OCCIDENTE [3]=Total` | ENEL, OCCIDENTE | — | 1, 2, 3 |
| 2026071 | 4 BOGOTA LIMPIA | 433 | `[1]=ENEL [2]=NUEVO ESQUEMA [3]=OCCIDENTE [4]=Total` | ENEL, NUEVO ESQUEMA, OCCIDENTE | — | 1, 2, 3 |
| 2026071 | 5 AREA LIMPIA | 418 | `[1]=ENEL [2]=ENERBIT [3]=OCCIDENTE [4]=Total` | ENEL, ENERBIT, OCCIDENTE | — | 1, 2, 3 |
| 2026072 (jul-2) | 1 | 451 | `[1]=ENEL [2]=OCCIDENTE [3]=Total` | ENEL, OCCIDENTE | — | 1, 2, 3 |
| 2026072 | 2 | 468 | `[1]=ENEL [2]=NUEVO ESQUEMA [3]=OCCIDENTE [4]=Total` | ENEL, NUEVO ESQUEMA, OCCIDENTE | — | 1, 2 |
| 2026072 | 3 | 434 | `[1]=ENEL [2]=OCCIDENTE [3]=Total` | ENEL, OCCIDENTE | — | 1, 2, 3 |
| 2026072 | 4 | 456 | `[1]=ENEL [2]=NUEVO ESQUEMA [3]=OCCIDENTE [4]=Total` | ENEL, NUEVO ESQUEMA, OCCIDENTE | — | 1, 2, 3 |
| 2026072 | 5 | 448 | `[1]=ENEL [2]=ENERBIT [3]=OCCIDENTE [4]=Total` | ENEL, ENERBIT, OCCIDENTE | — | 1, 2, 3 |
| 2026082 (ago-2) | 1 | 450 | `[1]=ENEL [2]=OCCIDENTE [3]=Total` | ENEL, OCCIDENTE | — | 1, 2, 3 |
| 2026082 | 2 | 473 | `[1]=ENEL [2]=NUEVO ESQUEMA [3]=OCCIDENTE [4]=Total` | ENEL, NUEVO ESQUEMA, OCCIDENTE | — | 1, 2 |
| 2026082 | 3 | 427 | `[1]=ENEL [2]=OCCIDENTE [3]=Total` | ENEL, OCCIDENTE | — | 1, 2, 3 |
| 2026082 | 4 | 445 | `[1]=ENEL [2]=NUEVO ESQUEMA [3]=OCCIDENTE [4]=Total` | ENEL, NUEVO ESQUEMA, OCCIDENTE | — | 1, 2, 3 |
| 2026082 | 5 | 434 | `[1]=ENEL [2]=ENERBIT [3]=OCCIDENTE [4]=Total` | ENEL, ENERBIT, OCCIDENTE | — | 1, 2, 3 |
| **2026101 (oct-1)** | 1 | 404 | `[1]=ENEL [2]=OCCIDENTE [3]=Total` | ENEL, OCCIDENTE | — | 1, 2, 3 |
| **2026101** | **2 LIME** | **391** | `[1]=ENEL [2]=OCCIDENTE [3]=Total` | ENEL, OCCIDENTE | **NUEVO ESQUEMA** | 1, 2, 3 |
| 2026101 | 3 | 378 | `[1]=ENEL [2]=OCCIDENTE [3]=Total` | ENEL, OCCIDENTE | — | 1, 2, 3 |
| **2026101** | **4 BOGOTA LIMPIA** | **394** | `[1]=ENEL [2]=OCCIDENTE [3]=Total` | ENEL, OCCIDENTE | **NUEVO ESQUEMA** | 1, 2, 3 |
| 2026101 | 5 | 409 | `[1]=ENEL [2]=ENERBIT [3]=OCCIDENTE [4]=Total` | ENEL, ENERBIT, OCCIDENTE | — | 1, 2, 3 |

**Lectura de la matriz:** el orden físico de las columnas **no** es estable entre ASE (ASE1/3 tienen `OCCIDENTE`
en `[2]`; ASE2/4 lo tienen en `[3]` porque insertan `NUEVO ESQUEMA` en `[2]`; ASE5 pone `ENERBIT` en `[2]`), y en
octubre-1 el `NUEVO ESQUEMA` **desaparece** de ASE2 y ASE4 dejando `OCCIDENTE` en `[2]`. Un mapa posicional
(`T0-0.3`, `ColumnasFuentePorAse`) no puede representar esa forma: es exactamente el rojo del plan (§0.1).

---

## 2. S1.2 — `Total` se distingue de una fuente

- En las 20 muestras hay **exactamente una** celda de header cuyo normalizado es `total` (columna `[3]` en ASE1/3,
  `[4]` en ASE2/4/5 y en octubre-1 ASE2/ASE4).
- Ninguna clave de `EditablesPorAse` normaliza a `total` (las claves son `ENEL`, `OCCIDENTE`, `NUEVO ESQUEMA`, `ENERBIT`).
- Conclusión: la exclusión explícita `normalizado == "total"` **antes** del vocabulario (regla §2.3.1 del plan) es
  suficiente; `Total` nunca se emitirá como empresa ni como evidencia de etiqueta desconocida.

---

## 3. S1.3 — Conceptos 1/2/3/7 y fila `Total`

- **Concepto 1** (`1-APLICADOS A FACTURACION`), **2** (`2-SALDOS A FAVOR GENERADOS`) y **3**
  (`3-APLICADOS A FINANCIACIONES NUEVAS`) se localizan por prefijo normalizado en las 20 muestras; ASE2 **no trae**
  la fila 3 en ninguno de los 4 períodos (0 demostrable, comportamiento ya cubierto por `ReporteBancoTests`).
- **Concepto 7** (`7-APLICADOS A RECIBOS SERVICIOS ESPECIALES`) **no viene en ninguna** de las 20 muestras → 0
  demostrable intacto (`ReporteBancoInputs.cs:34-38`).
- La fila `Total` existe en las 20 muestras y aporta valor a **todas** las columnas de fuente presentes
  (p. ej. octubre-1 ASE2 `R396`: `B=10.309.260.573`, `C=971.950.730`, `D=11.281.211.303`).

Evidencia por muestra (fila `Total` y conceptos):

| Período | ASE | Conceptos (filas) | `Total` (fila) |
|---|---|---|---|
| jul-1 | 1 | 416, 417, 418 | 419 |
| jul-1 | 2 | 442, 443 | 444 |
| jul-1 | 3 | 408, 409, 410 | 411 |
| jul-1 | 4 | 435, 436, 437 | 438 |
| jul-1 | 5 | 420, 421, 422 | 423 |
| jul-2 | 1 | 453, 454, 455 | 456 |
| jul-2 | 2 | 470, 471 | 472 |
| jul-2 | 3 | 436, 437, 438 | 439 |
| jul-2 | 4 | 458, 459, 460 | 461 |
| jul-2 | 5 | 450, 451, 452 | 453 |
| ago-2 | 1 | 452, 453, 454 | 455 |
| ago-2 | 2 | 475, 476 | 477 |
| ago-2 | 3 | 429, 430, 431 | 432 |
| ago-2 | 4 | 447, 448, 449 | 450 |
| ago-2 | 5 | 436, 437, 438 | 439 |
| oct-1 | 1 | 406, 407, 408 | 409 |
| oct-1 | 2 | 393, 394, 395 | 396 |
| oct-1 | 3 | 380, 381, 382 | 383 |
| oct-1 | 4 | 396, 397, 398 | 399 |
| oct-1 | 5 | 411, 412, 413 | 414 |

---

## 4. S1.4 — Σ conceptos == `TotalFuente` por empresa presente

**40/40 combinaciones (20 muestras × empresas presentes) con Δ = `0,00`.** El gate D5(i)
(`WorkbookLeafCoherence.ValidarContraFuentesBanco` + `ValidadorBasico.ValidarReporteBanco`) no cambia: la empresa
ausente **no entra** al gate (se omite; no hay `TotalFuente = 0` que confundir con un total real).

| Muestra | Σ ENEL | Σ NUEVO ESQUEMA | Σ ENERBIT | Σ OCCIDENTE |
|---|---|---|---|---|
| jul-1 ASE1 | 16.244.803.518 | — | — | 540.475.692,00 |
| jul-1 ASE2 | 11.833.210.912 | 29.345.550 | — | 422.221.285 |
| jul-1 ASE3 | 9.685.219.461 | — | — | 472.519.065,00 |
| jul-1 ASE4 | 10.376.694.382 | 15.240.310 | — | 239.142.386,00 |
| jul-1 ASE5 | 8.278.977.659 | — | 129.536.656 | 181.733.562,00 |
| jul-2 ASE1 | 16.660.688.191 | — | — | 805.313.717,00 |
| jul-2 ASE2 | 19.092.079.577 | 24.231.760 | — | 1.425.946.178 |
| jul-2 ASE3 | 14.054.241.942 | — | — | 1.245.170.519,00 |
| jul-2 ASE4 | 6.652.822.320 | 15.885.640 | — | 486.639.517,00 |
| jul-2 ASE5 | 11.726.152.837 | — | 28.645.323 | 321.499.121,00 |
| ago-2 ASE1 | 17.485.468.108 | — | — | 910.584.918,00 |
| ago-2 ASE2 | 19.313.863.587 | 19.061.190 | — | 2.386.806.166 |
| ago-2 ASE3 | 15.145.312.302 | — | — | 1.268.918.628,00 |
| ago-2 ASE4 | 7.396.103.941 | 8.844.380 | — | 977.975.927,00 |
| ago-2 ASE5 | 11.781.230.026 | — | 32.556.620 | 341.210.575,00 |
| oct-1 ASE1 | 14.510.944.162 | — | — | 370.866.743,00 |
| oct-1 ASE2 | 10.309.260.573,00 | — (ausente) | — | 971.950.730,00 |
| oct-1 ASE3 | 8.407.422.012 | — | — | 437.110.235,00 |
| oct-1 ASE4 | 9.932.648.720 | — (ausente) | — | 599.663.169,00 |
| oct-1 ASE5 | 7.248.898.705 | — | 140.424.760 | 175.905.753,00 |

(Valores `Total` de la fila del reporte por empresa; en las 40 celdas la Σ de los conceptos coincide con el total
al centavo, tolerancia ±0,5 no consumida.)

---

## 5. S1.5 — Veredicto octubre-1 y paridad global esperado↔observado

| ASE | Esperadas (`EditablesPorAse`) | Observadas oct-1 | Veredicto |
|---|---|---|---|
| 1 | ENEL, OCCIDENTE | ENEL, OCCIDENTE | paridad |
| **2 LIME** | ENEL, **NUEVO ESQUEMA**, OCCIDENTE | ENEL, OCCIDENTE | **ausencia legítima** de `NUEVO ESQUEMA` → omisión + huella `BANCO:NUEVO ESQUEMA` |
| 3 | ENEL, OCCIDENTE | ENEL, OCCIDENTE | paridad |
| **4 BOGOTA LIMPIA** | ENEL, **NUEVO ESQUEMA**, OCCIDENTE | ENEL, OCCIDENTE | **ausencia legítima** de `NUEVO ESQUEMA` → omisión + huella (dato medido) |
| 5 | ENEL, ENERBIT, OCCIDENTE | ENEL, ENERBIT, OCCIDENTE | paridad |

### 5.1 Desviación declarada respecto del plan

El §3-S1.5 del plan esperaba que **solo ASE2** llegara con 2 fuentes en octubre-1 y que **ASE4/ASE5** trajeran 3.
La medición de bytes dice: **ASE4 de octubre-1 también trae 2** (`NUEVO ESQUEMA` ausente). No es un STOP:

1. No hay header desconocido (S1.6 vacío) ni fuente reconocida sin celdas destino (observado ⊆ esperado).
2. La ausencia es **del label completo** (no hay bloque parcial): la misma familia que ASE2 y que las 8 ausencias
   del Plan 50, absorbida por la regla genérica ratificada (**cero ramas** por período/empresa/fuente).
3. El efecto es el previsto por el plan: la fuente no se emite, sus celdas destino (`F41:F44` del bloque ASE4)
   quedan en el 0 de la plantilla y la huella viaja como `ASE 4 · BANCO:NUEVO ESQUEMA`.

El fixture congela este veredicto en `T0P51_Octubre1_Ase4_TambienOmiteNuevoEsquema_VeredictoMedido`.

---

## 6. S1.6 — Headers no reconocidos (inventario)

**VACÍO en las 20 muestras.** Vocabulario aplicado: los 5 `LabelTemplate` del catálogo
(`RECIPROCIDAD`, `ENEL`, `ENERBIT`, `OCCIDENTE`, `CIUDAD LIMPIA-ACUEDUCTO`) + el rótulo vigente por ASE
(empresa 1 en ASE2/ASE4 → `NUEVO ESQUEMA`) + la sinonimia citada `RECIPROCIDAD ↔ NUEVO ESQUEMA` (Plan 34).
No hay `EAB`, `TOTAL`-como-fuente ni ninguna otra etiqueta en la fila de headers.

---

## 7. G-SHA y notas de método

- Plantilla **v2** `Docs/Plantilla estandarizada/v2/Formato Remuneracion 2026xx-x Total vf.xlsx`:
  `0F192AD2721C77986B6A8CB85D1A92F782DCD2F69989F362E066FA17027F2DAB` (verificada antes de H0).
- Plantilla **v1** `Docs/Plantilla estandarizada/Formato Remuneracion 2026xx-x Total.xlsx`:
  `3086AC317595C2F8E84ED24B1C5EE3E52565DD5193394E3062878E7FA49C5214` (verificada antes de H0).
- Las fuentes se copiaron a `%TEMP%` antes de leer (Excel puede tenerlas abiertas); ni fuentes ni plantillas se
  modificaron.
- El fixture no consume `WorkbookLeafCellMapReporteBanco.ColumnasFuentePorAse` (mapa posicional que este plan
  retira): la observación es BCL pura y la expectativa por ASE sale de `EditablesPorAse[ase].Keys`.
