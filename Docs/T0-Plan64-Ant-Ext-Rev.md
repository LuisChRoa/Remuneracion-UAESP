# T0 — Plan 64 · ANT EXT-REV desde `RecaudosReversados`

**Fase:** H0/F0 — T0 read-only (bloqueante, cero código de producción).
**Plan rector:** `plans/64 - ant-ext-rev-recaudos-reversados.md` (fuente única de verdad).
**Fixture:** `Remuneracion.IntegrationTests/AntExtRevT0Plan64Tests.cs`.
**Fecha de medición:** 2026-10-30.
**Método:** BCL puro (`ZipArchive` + `System.Xml.Linq`) sobre la plantilla v4 y los 4 manuales; ExcelDataReader 3.9.0 **solo-lectura sobre copias** temporales de los 30 `RecaudosReversados_*.xlsx`. Ningún archivo se abrió en escritura. Rutas **explícitas** (no se usó `Insumos.Raiz()`).

> **Regla aplicada:** cada número de este documento es el BYTE medido. Donde el §1 del plan difiere del disco, se re-midió y se declara en §8 como **enmienda requerida**; el texto del §1 **no** se convirtió en contrato.

---

## §1. Rutas medidas

| Rol | Ruta |
|---|---|
| Plantilla vigente v4 | `Docs/Plantilla estandarizada/v4/Formato Remuneracion 2026xx-x Total v4.xlsx` |
| Insumos Julio-1 | `Docs/Prueba Julio-1/Insumos/{1-Promoambiental,2-Lime,3-Ciudad Limpia,4-Bogota Limpia,5-Área Limpia}` |
| Insumos Julio-2 | `Docs/Prueba Julio-2/Insumos/…` |
| Insumos Agosto-2 | `Docs/Prueba Agosto-2/Insumos/…` |
| Insumos Septiembre-1 | `Docs/Prueba Septiembre 1/Insumos/…` |
| Insumos Septiembre-2 | `Docs/Prueba Septiembre-2/Insumos/…` |
| Insumos Octubre-1 | `Docs/Prueba Octubre-1/Insumos/…` |
| Manual Julio-2 | `Docs/Prueba Julio-2/Resultado/Resultado Administrativo con plantilla estandar v2/Formato Remuneracion 202607-2 Total.xlsx` |
| Manual Septiembre-1 | `Docs/Prueba Septiembre 1/Resultados/Administrativo/Remuneracion 202609-1 Total.xlsx` |
| Manual Septiembre-2 | `Docs/Prueba Septiembre-2/Resultados/Administrativo/Remuneracion 202609-2 Total.xlsx` |
| Manual Octubre-1 | `Docs/Prueba Octubre-1/Resultados/Resultado administrativo/Remuneracion 202610-1 Total.xlsx` |
| Docx rector | `Docs/Detalle de plantilla.docx` |

Conteo cerrado: **30 insumos** = 6 períodos × 5 ASE, cada uno con exactamente un `RecaudosReversados_*.xlsx` (además existen `.pdf` homónimos que no se leen). **4 manuales** presentes.

---

## §2. Frente (a) — Geometría v4 de `ANT EXT-REV` (REQ-01)

Hoja localizada por **nombre** (`xl/worksheets/sheet7.xml` vía `workbook.xml` + rels). Dimensión declarada: **`A1:P45`**.

### §2.1 Bloque superior (filas 1–6) — texto de `<f>` medido

| Celda | Fórmula efectiva (shared expandida) |
|---|---|
| `C3`..`L3` | `X11+X19+X27+X35+X43` (por columna X) |
| `C4`..`L4` | `X12+X20+X28+X36+X44` |
| `C5`..`L5` | `SUM(X3:X4)` |
| `C6`..`L6` | `X5=(X13+X21+X29+X37+X45)` |
| `M3` / `N3` | `C3+E3+G3+I3+K3` / `D3+F3+H3+J3+L3` |
| `M4` / `N4` | `C4+E4+G4+I4+K4` / `D4+F4+H4+J4+L4` |
| `M5` / `N5` | `SUM(M3:M4)` / `SUM(N3:N4)` |
| `M6` / `N6` | `M5=(M13+M21+M29+M37+M45)` / `N5=(N13+N21+N29+N37+N45)` |
| `P3` (t=`b`) | `N3=(-VALIDACION_TOTAL!M9)` — caché `1` |
| `P4` (t=`b`) | `N4=(-'CONSOLIDADO_TOTAL RECAUDO'!L131)` — caché `1` |

Forma de las compartidas (evidencia de bytes): `si=0 ref=E3:K3` (fila 3), `si=1 ref=E4:K4` (fila 4), `si=2 ref=D5:L5` (fila 5), `si=4 ref=C6:K6` (fila 6).

### §2.2 Las 5 etiquetas de bloque (columna B)

| Etiqueta | Fila | ASEs | REPORTE 3/4 | APLICADO | ANTICIPOS | TOTAL |
|---|---|---|---|---|---|---|
| `BANCO DE OCCIDENTE` | 8 | 9 | 10 | 11 | 12 | 13 |
| `EAAB RECIPROCIDAD` | 16 | 17 | 18 | 19 | 20 | 21 |
| `EAAB CIUDAD LIMPIA` | 24 | 25 | 26 | 27 | 28 | 29 |
| `ENEL` | 32 | 33 | 34 | 35 | 36 | 37 |
| `ENERBIT` | 40 | 41 | 42 | 43 | 44 | 45 |

Filas de ASE: `C/E/G/I/K = ASE 1..5` y `M = TOTAL`. Fila de reportes: alternan `REPORTE 3`/`REPORTE 4` en **C..N** (12 columnas), no solo C..L (ver §8-M2).

### §2.3 Rejilla de valores y fila TOTAL por bloque

- Filas `APLICADO` y `ANTICIPOS`: `M = C+E+G+I+K` y `N = D+F+H+J+L` (suma horizontal de las 5 columnas de ASE).
- Fila `TOTAL`: `C..L = SUM(Xap:Xan)`. La forma de `M`/`N` es **heterogénea** (medida, congelada):

| Bloque | `M` del TOTAL | `N` del TOTAL |
|---|---|---|
| BANCO DE OCCIDENTE | `SUM(M11:M12)` | `SUM(N11:N12)` |
| EAAB RECIPROCIDAD | `C21+E21+G21+I21+K21` | `SUM(N19:N20)` |
| EAAB CIUDAD LIMPIA | `SUM(M27:M28)` | `SUM(N27:N28)` |
| ENEL | `SUM(M35:M36)` | `SUM(N35:N36)` |
| ENERBIT | `SUM(M43:M44)` | `SUM(N43:N44)` |

### §2.4 Naturaleza de las 100 celdas de valor

**Medido:** 5 bloques × 10 columnas (`C..L`) × 2 filas (`APLICADO`, `ANTICIPOS`) = **100 celdas**, todas con `<v>0</v>`, **sin** `<f>`, **sin** `si`, y **sin atributo `t`** ⇒ son **numéricas en cero**, no texto ni vacías (ver §8-M1).

### §2.5 "Protección" — precisión de bytes

La hoja **no trae `<sheetProtection>`**. Lo que el código vigente llama "protegidas" (`WorkbookLeafCellMapQ2.cs:402-405`, `WorkbookLeafCellMapAjustesSfT.cs:248-255`) son celdas **con `<f>` que no son destino** (guard anti-fórmula). Conteo medido en `ANT EXT-REV`: **130 celdas con `<f>`** (48 del bloque superior + 2 controles + 80 de las filas M/N y TOTAL de los 5 bloques), y **0** en las 100 celdas de valor (ver §8-M3).

---

## §3. Frente (b) — `ANTICIPOS USUARIOS` y `CONSOLIDADO_TOTAL RECAUDO!L125:L130` (REQ-01)

### §3.1 `ANTICIPOS USUARIOS` es 100 % fórmulas — CONFIRMADO

- Rejilla de datos (columnas `C,F,I,L,O,R` × filas `3..7`, `12..16`, `21..25`): **90/90 celdas con `<f>`**.
- Filas de suma (`8`, `17`, `26`): **18/18 `SUM(Xn:Xn+4)`**. Varias son **seguidoras de compartidas** (`si=3`): `O26` **sí** es fórmula (`SUM(O21:O25)` tras expandir la compartida) — una lectura que ignore `si` la vería "vacía".

### §3.2 Referencias a `ANT EXT-REV` (medidas)

`72` celdas de `ANTICIPOS USUARIOS` referencian `'ANT EXT-REV'`. Mapa medido (fila del template → celda del insumo):

| Bloque del template | Columna ANTI | Celdas insumo (ASE 1..5) |
|---|---|---|
| "anticipos generados por extemporáneos" (filas 12..16) | `C..` de la fila `ANTICIPOS` (4) | `C4`, `E4`, `G4`, `I4`, `K4` |
| "anticipos que cargan por reversiones" (filas 21..25) | `C..` del reporte 4 (fila 4) | `D4`, `F4`, `H4`, `J4`, `L4` — **incluye los `D4`/`M4`/`D20`/`M20` citados por el §1.1** |

Auto-chequeos medidos: `C9='REPORTE RECAUDO x BANCO'!I3` (bloque de pagos), `C18='ANT EXT-REV'!M4`, `C27='ANT EXT-REV'!N4` (y sus gemelas por columna).

### §3.3 `CONSOLIDADO_TOTAL RECAUDO!L125:L130`

| Celda | Contenido medido |
|---|---|
| `D125:D129` | `'ANTICIPOS USUARIOS'!C3..C7` |
| `K125:K129` | `'ANTICIPOS USUARIOS'!C12..C16` |
| `L125:L129` | `-'ANTICIPOS USUARIOS'!C21..C25` |
| `L130` | **VACÍA** (sin `<f>` y sin `<v>`) |
| `L131` | `SUM(L125:L130)` — consumida por `ANT EXT-REV!P4` |

### §3.4 Los 4 manuales conservan las fórmulas

Plantilla v4 y los 4 manuales: **130 celdas con `<f>`** en `ANT EXT-REV`, y `P3 = P4 = TRUE` en los 5 archivos. Leído como **estado del archivo al guardarse** (nunca como oráculo de la fórmula: doctrina Plan 28/29/40). Evidencia de que el administrativo diligenció **a valores** sin tocar `<f>`.

---

## §4. Frente (c) — Estructura `Sheet1` de los 30 insumos (REQ-02)

Una única hoja por archivo, **siempre `Sheet1`** (30/30). Marcadores localizados por texto normalizado (mayúsculas, sin diacríticos, trim, espacios colapsados):

- `Pagos Pendientes por Aplicar` (bloque 2) — siempre **antes** de
- `Resumen Recaudo Aplicado Por Servicio` (bloque 3). **0/30 invertidos.**

### §4.1 Encabezado del bloque 2 (HUE-3)

Medido en **30/30**: fila del marcador `A='Pagos Pendientes por Aplicar'`, `B='Banco'`, `C='Fecha Recaudo'`, `D='Cantidad'`, `E='Valor'`; dos filas después, la cabecera real de la tabla de fuentes: `E='ASEO'`, `F='Total'`; y en la fila siguiente `D='66'` (texto), `E='Total'`.

⚠️ Ese encabezado **no nombra** las columnas de valor: `D='Cantidad'` pertenece a la tabla de *pendientes* (filas vacías) y **no** describe la columna `D` de las filas de concepto. En **todas** las filas de concepto y de total de canal: **`D = E = F`** (30/30, Δ ≤ 0,0000001). Ver §8-M4.

### §4.2 Filas del bloque 2 (30/30)

| Tipo de fila | Firma medida |
|---|---|
| Concepto | `C ∈ {1-APLICADOS A FACTURACION, 2-SALDOS A FAVOR GENERADOS, 9-REVERSION PAGOS A COMPONENTES, 9-REVERSION SALDOS A FAVOR}`, `A` y `B` vacías, `D=E=F` |
| Total de canal | `B = canal`, `C = 'Total'`, `D=E=F` |
| Total de fuente | `A = fuente`, `B = 'Total'`, `D=E=F` — **cierra el grupo** (la etiqueta de fuente va al final) |
| Gran total | `A = 'Total'`, `B` vacía |

Orden real por grupo: filas de concepto(s) → total de canal que las cierra → … → total de fuente. `0` filas no reconocidas en los 30 archivos.

### §4.3 Tabla congelada 30/30 (evidencia, **no** contrato posicional)

| Período | ASE | marcador 2 | marcador 3 | Fuentes del bloque 2 (total) | Columnas del bloque 3 | Total bloque 2 | Total bloque 3 |
|---|---|---|---|---|---|---|---|
| 2026071 | 1 | 11 | 33 | OCCIDENTE=-7402760 | RECAUDO OCCIDENTE\|TOTAL | -7402760 | -7402760 |
| 2026071 | 2 | 54 | 104 | ENEL=-9917803 \| NUEVO ESQUEMA=-620493 | RECAUDO OTROS\|TOTAL | -10538296 | -10538296 |
| 2026071 | 3 | 52 | 104 | ENEL=-5627802 \| OCCIDENTE=-10051780 | RECAUDO OCCIDENTE\|RECAUDO OTROS\|TOTAL | -15679582 | -15679582 |
| 2026071 | 4 | 248 | 342 | ENEL=-3475400 \| NUEVO ESQUEMA=0 \| OCCIDENTE=-522720 | RECAUDO OCCIDENTE\|RECAUDO OTROS\|TOTAL | -3998120 | -3998120 |
| 2026071 | 5 | 43 | 92 | ENEL=-5470988 \| OCCIDENTE=-1386390 | RECAUDO OCCIDENTE\|RECAUDO OTROS\|TOTAL | -6857378 | -6857378 |
| 2026072 | 1 | 60 | 110 | ENEL=-15501271 \| OCCIDENTE=-1086060 | RECAUDO OCCIDENTE\|RECAUDO OTROS\|TOTAL | -16587331 | -16587331 |
| 2026072 | 2 | 65 | 124 | ENEL=-9472909 \| OCCIDENTE=-178170 | RECAUDO OCCIDENTE\|RECAUDO OTROS\|TOTAL | -9651079 | -9651079 |
| 2026072 | 3 | 47 | 94 | ENEL=-30776490 \| OCCIDENTE=-2650410 | RECAUDO OCCIDENTE\|RECAUDO OTROS\|TOTAL | -33426900 | -33426900 |
| 2026072 | 4 | 293 | 389 | ENEL=-14910059 \| NUEVO ESQUEMA=0 \| OCCIDENTE=-14926610 | RECAUDO OCCIDENTE\|RECAUDO OTROS\|TOTAL | -29836669 | -29836669 |
| 2026072 | 5 | 30 | 78 | ENEL=-2114880 \| OCCIDENTE=-663150 | RECAUDO OCCIDENTE\|RECAUDO OTROS\|TOTAL | -2778030 | -2778030 |
| 2026082 | 1 | 10 | 36 | ENEL=0 | RECAUDO OTROS\|TOTAL | 0 | 0 |
| 2026082 | 2 | 70 | 134 | ENEL=-10736662 \| OCCIDENTE=0 | RECAUDO OCCIDENTE\|RECAUDO OTROS\|TOTAL | -10736662 | -10736662 |
| 2026082 | 3 | 32 | 69 | ENEL=-4617354 | RECAUDO OTROS\|TOTAL | -4617354 | -4617354 |
| 2026082 | 4 | 237 | 326 | ENEL=-3442420 \| NUEVO ESQUEMA=0 \| OCCIDENTE=0 | RECAUDO OCCIDENTE\|RECAUDO OTROS\|TOTAL | -3442420 | -3442420 |
| 2026082 | 5 | 80 | 131 | ENEL=-1454550 \| OCCIDENTE=-66010 | RECAUDO OCCIDENTE\|RECAUDO OTROS\|TOTAL | -1520560 | -1520560 |
| 2026091 | 1 | 18 | 57 | ENEL=-3462950 \| OCCIDENTE=0 | RECAUDO OCCIDENTE\|RECAUDO OTROS\|TOTAL | -3462950 | -3462950 |
| 2026091 | 2 | 30 | 72 | ENEL=-8739913 | RECAUDO OTROS\|TOTAL | -8739913 | -8739913 |
| 2026091 | 3 | 40 | 91 | ENEL=-6939290 \| OCCIDENTE=0 | RECAUDO OCCIDENTE\|RECAUDO OTROS\|TOTAL | -6939290 | -6939290 |
| 2026091 | 4 | 346 | 433 | ENEL=-3196930 \| NUEVO ESQUEMA=0 \| OCCIDENTE=-64000 | RECAUDO OCCIDENTE\|RECAUDO OTROS\|TOTAL | -3260930 | -3260930 |
| 2026091 | 5 | 31 | 72 | ENEL=-8838800 \| OCCIDENTE=0 | RECAUDO OCCIDENTE\|RECAUDO OTROS\|TOTAL | -8838800 | -8838800 |
| 2026092 | 1 | 26 | 67 | ENEL=-690790 \| OCCIDENTE=0 | RECAUDO OCCIDENTE\|RECAUDO OTROS\|TOTAL | -690790 | -690790 |
| 2026092 | 2 | 642 | 735 | ENEL=-15212817 \| NUEVO ESQUEMA=0 \| OCCIDENTE=0 | RECAUDO OCCIDENTE\|RECAUDO OTROS\|TOTAL | -15212817 | -15212817 |
| 2026092 | 3 | 40 | 84 | ENEL=-6595914 | RECAUDO OTROS\|TOTAL | -6595914 | -6595914 |
| 2026092 | 4 | 2242 | 2358 | ENEL=-4393374 \| NUEVO ESQUEMA=0 \| OCCIDENTE=0 | RECAUDO OCCIDENTE\|RECAUDO OTROS\|TOTAL | -4393374 | -4393374 |
| 2026092 | 5 | 28 | 69 | ENEL=-4790099 | RECAUDO OTROS\|TOTAL | -4790099 | -4790099 |
| 2026101 | 1 | 30 | 79 | ENEL=-3184910 \| OCCIDENTE=0 | RECAUDO OCCIDENTE\|RECAUDO OTROS\|TOTAL | -3184910 | -3184910 |
| 2026101 | 2 | 340 | 426 | ENEL=-6297940 \| OCCIDENTE=0 | RECAUDO OCCIDENTE\|RECAUDO OTROS\|TOTAL | -6297940 | -6297940 |
| 2026101 | 3 | 26 | 61 | ENEL=-2259272 | RECAUDO OTROS\|TOTAL | -2259272 | -2259272 |
| 2026101 | 4 | 746 | 851 | ENEL=-1581684 \| NUEVO ESQUEMA=0 \| OCCIDENTE=0 | RECAUDO OCCIDENTE\|RECAUDO OTROS\|TOTAL | -1581684 | -1581684 |
| 2026101 | 5 | 26 | 68 | ENEL=-4005397 | RECAUDO OTROS\|TOTAL | -4005397 | -4005397 |

**Bloque 3 = control, nunca fuente.** Columnas variables medidas: **22 archivos** con `RECAUDO OCCIDENTE | RECAUDO OTROS | Total` y **8 archivos** con solo `RECAUDO OTROS | Total`:

`2026071 ASE2`, `2026082 ASE1`, `2026082 ASE3`, `2026091 ASE2`, `2026092 ASE3`, `2026092 ASE5`, `2026101 ASE3`, `2026101 ASE5`.

Es decir: el caso "ASE2 Sept-1 sin OCCIDENTE" del §1.3 se confirma, pero **no es el único** (ver §8-M5). En los 30 archivos el total del bloque 3 **cierra exactamente** contra el total del bloque 2 (Δ = 0): el bloque 3 no aporta ni quita dinero, solo agrega.

### §4.4 Invariante Σ-grupo == total de fuente

En los 30 archivos, por cada fuente presente: `Σ(4 conceptos) == total de fuente` con Δ ≤ 1e-9 (tolerancia de comparación ±0,5). Cada total de canal cierra contra los conceptos que lo preceden. El gran total del bloque 2 == Σ totales de fuente.

---

## §5. Frente (d) — Universo de fuentes y paridad §1.4 (REQ-02, REQ-03, HUE-1)

### §5.1 Universo (30/30)

| Etiqueta de fuente (bloque 2) | Archivos en que aparece | Total en los 4 manuales |
|---|---|---|
| `ENEL` | 29 / 30 (falta solo en `2026071 ASE1`) | ≠ 0 (con datos) |
| `OCCIDENTE` | 22 / 30 | ≠ 0 (con datos) |
| `NUEVO ESQUEMA` | 8 / 30 (`2026071 ASE2/ASE4`, `2026072 ASE4`, `2026082 ASE4`, `2026091 ASE4`, `2026092 ASE2/ASE4`, `2026101 ASE4`) | ≠ 0 / 0 |
| `EAAB CIUDAD LIMPIA` | **0 / 30** | **0** en los 4 manuales (bloque B24, 5 ASE × 2 filas × 2 reportes) |
| `ENERBIT` | **0 / 30** | **0** en los 4 manuales (bloque B40) |

HUE-1 queda **sostenido por bytes**: sin fila de catálogo para CIUDAD LIMPIA / ENERBIT; si aparecieran, `ERR-FORMATO-FUENTE` (nunca 0 silencioso).

### §5.2 Los 4 casos citados por el §1.4 — re-derivados y confirmados

| Período | ASE | Fuente | Aplicado/R3 | Anticipos/R3 | Aplicado/R4 | Anticipos/R4 | Destino |
|---|---|---|---|---|---|---|---|
| 2026091 | 1 | OCCIDENTE | 42 680 | 40 630 | −83 310 | 0 | `C11`, `C12`, `D11`, `D12` |
| 2026091 | 1 | ENEL | 0 | 0 | −3 461 197,29 | −1 752,71 | `D35`, `D36` |
| 2026091 | 4 | OCCIDENTE | 239 834 | 1 438 701 | −1 742 535 | 0 | `I11`, `I12`, `J11`, `J12` |
| 2026091 | 4 | NUEVO ESQUEMA | 208 970 | 53 030 | −262 000 | 0 | `I19`, `I20`, `J19`, `J20` |
| 2026091 | 4 | ENEL | 8 186 910 | 9 321 147 | −20 245 524,78 | −459 462,22 | `I35`, `I36`, `J35`, `J36` |

Los cinco casos coinciden **exactamente** con el §1.4 (Δ < 1e-6) y con el manual del administrativo.

### §5.3 Paridad completa `fuente × etiqueta` (20 archivos con manual)

`2026072`, `2026091`, `2026092`, `2026101` × 5 ASE: **todas** las celdas destino del §1.2 para **todas** las fuentes presentes coinciden con el manual ±0,5 (Δ máx. medida < 1e-6; los residuos son ruido de coma flotante del `.xlsx`). Esto valida en bytes:

- el catálogo etiqueta-de-bloque ← rótulo-de-fuente (`BANCO DE OCCIDENTE ← OCCIDENTE`, `EAAB RECIPROCIDAD ← NUEVO ESQUEMA`, `ENEL ← ENEL`);
- el catálogo concepto → (fila `APLICADO`/`ANTICIPOS`, reporte `REPORTE 3`/`REPORTE 4`);
- el signo "tal cual" (los R4 ya vienen negativos; la plantilla no re-signa);
- la columna del ASE derivada de la etiqueta `ASE n` del template (nunca por índice fijo);
- la regla D-04 (fuente ausente = 0 del template; concepto ausente en fuente presente = 0).

**Única celda vacía** en el universo de manuales: `2026092` hoja `ANT EXT-REV` celda **`J20`** (EAAB RECIPROCIDAD, ASE4, `ANTICIPOS`×`REPORTE 4`): el concepto `9-REVERSION SALDOS A FAVOR` está ausente en el insumo y el manual la deja **en blanco** en lugar de `0` explícito. Se compara como 0 (diferencia 0) y se declara para H2/H4 (§8-H1).

---

## §6. Frente (e) — Divergencia docx-vs-encargo (D-10), con cita exacta

Texto extraído de `Docs/Detalle de plantilla.docx` (`word/document.xml`, párrafos en orden; comillas tipográficas normalizadas):

1. Encabezado de sección: **`ANT EXT-REV`**.
2. *«Se utiliza la información "Resumen Recaudo Aplicado Por Servicio ", la cual se encuentra al final del reporte de recaudos reversados generado para cada concesionario o ASE.»*
3. *«Para las primeras pruebas de la ejecución de este proyecto de remuneración, se entregará para cada periodo esta hoja "ANT EXT-REV" diligenciada. Posteriormente, daremos el instructivo detallado para que sea diligenciada con la ejecución del reporte.»*
4. *«La información de esta hoja "ANT EXT-REV" se utiliza o alimenta la hoja de "ANTICIPOS USUARIOS".»*

**Veredicto por bytes (sustenta D-10):**

- El docx rector nombra el **"Resumen Recaudo Aplicado Por Servicio"** (bloque 3) como insumo.
- El bloque 3 medido **no discrimina por fuente**: sus columnas son `RECAUDO OCCIDENTE`, `RECAUDO OTROS`, `Total` — nunca `ENEL`, `NUEVO ESQUEMA` ni `ENERBIT` —, y en 8 de los 30 archivos no hay siquiera columna `OCCIDENTE`. `RECAUDO OTROS` agrega ENEL + NUEVO ESQUEMA (y ENERBIT si viniera) sin separarlos.
- El **bloque 2** sí trae una **fila de total por fuente** rotulada con la etiqueta de la fuente, y sus `Σ-conceptos` cierran contra ese total.
- ⇒ Usar el bloque 3 como fuente obligaría a **inventar el reparto** entre ENEL / NUEVO ESQUEMA / ENERBIT. Veredicto D-10: **bloque 2 = fuente, bloque 3 = control**, con HUE-1 declarado.

No se encontró en el docx mapeo alguno para EAAB CIUDAD LIMPIA ni ENERBIT: **HUE-1 confirmado** (sin evidencia ⇒ sin fila de catálogo).

---

## §7. Frente (f) — Δ de suite y huella v4

### §7.1 Huella v4 (G-HUELLA)

| Momento | Huella semántica (`HuellaSemanticaPlantilla.Calcular`) |
|---|---|
| Antes de H0 (citada por el plan §0 gate) | `a7476ee2e80dfa13c9e39430d1ea4d77f0c52e3ca0a3d11a526d5140b8608c55` |
| Medida por el fixture H0 (`T0_07_HuellaV4_Intacta`) | `a7476ee2e80dfa13c9e39430d1ea4d77f0c52e3ca0a3d11a526d5140b8608c55` |
| Después de H0 | idéntica (H0 no abre la plantilla en escritura) |

**Veredicto: huella INTACTA** — no se "ajustó" nada. El SHA-256 del archivo se reporta solo como dato informativo.

### §7.2 Delta de suite (set-diff)

| Corrida | Comando | Resultado |
|---|---|---|
| Base (suite **sin** el fixture nuevo) | `dotnet test Remuneracion.IntegrationTests/Remuneracion.IntegrationTests.csproj --filter "FullyQualifiedName!~AntExtRevT0Plan64Tests" --logger "trx;LogFileName=p64-h0-base.trx"` | **1799** tests · 1776 Passed · **23 Failed** · **0 Skipped** · 17 m 13 s → `TestResults/p64/p64-h0-base.trx` |
| H0 (suite **completa**) | `dotnet test Remuneracion.IntegrationTests/Remuneracion.IntegrationTests.csproj --logger "trx;LogFileName=p64-h0-final.trx"` | **1817** tests · 1794 Passed · **23 Failed** · **0 Skipped** · 21 m 03 s → `TestResults/p64/p64-h0-final.trx` |

**Set-diff exacto:** `1817 − 1799 = 18` tests nuevos, **todos `Passed`**; **0 rojos nuevos** y **0 rojos desaparecidos** (los mismos 23 en ambas corridas).

### §7.3 Rojos de entrada (baseline declarado, NO causados por H0)

H0 **no tocó código de producción** (`git status --porcelain` → solo el fixture nuevo y este documento), así que estos 23 rojos son el **estado de entrada** de la suite:

| # | Test |
|---|---|
| 1 | `AdopcionPlantillaV2T0Plan48Tests.T0P48_P3_R1_Manual_492_Formulas_DosPatrones` |
| 2 | `AdopcionPlantillaV2T0Plan48Tests.T0P48_P6_TablaParidad_40Hojas_ConteosReales` |
| 3 | `BalanceScPermutaTests.JulioAse5Bce_ManualDivergeDeLaFuente_ConstanteDocumentada` |
| 4 | `BalanceScPermutaTests.JulioQ2_PermutaBce_SubsidioAD_ContribucionAE_VsFuente_E2E` |
| 5 | `BalanceScTests.J9Consolidado_NoEsOraculoDeValoresBce` |
| 6 | `BaseConsistente2026082Tests.BaseAgosto2026082_EsInternamenteConsistenteConElManual` |
| 7 | `BaseConsistente2026082Tests.BaseAgosto2026082_MetadatosCanonicosEsperados` |
| 8 | `BasePeriodoPreflightTests.S5_OficialDirecto_EndToEnd_L9L13PorQuincena(codigo: "2026071", carpetaRel: "Docs/Prueba Julio-1/Insumos", columna: "N")` |
| 9 | `ComparadorSalidaVsManualTests.Salida5AseQ2_vs_Manual_CeroDivergenciasFueraDeBrechasDeclaradas(etiqueta: "agosto")` |
| 10 | `ComparadorSalidaVsManualTests.Salida5AseQ2_vs_Manual_CeroDivergenciasFueraDeBrechasDeclaradas(etiqueta: "julio")` |
| 11 | `GoldenBalanceScTests.ProcesadorPeriodo_CapaA2_4_BceContraGolden` |
| 12 | `GoldenConciliacionEmpresaTests.ProcesadorPeriodo_CapaA2_2_EmpresasYRecaudosContraGolden` |
| 13 | `GoldenInterventoriaTests.CapaA_InterventoriaYLEspeciales_AmbosCanonicos` |
| 14 | `GoldenLeafPathTests.Procesador_CapaA_LeafVisiblesFormulasYHashSinUsarCacheDeSalida` |
| 15 | `GoldenMultiAseTests.ProcesadorPeriodo_CapaAX5_LeafVisiblesFormulasHashYSumas` |
| 16 | `GoldenReporteBancoTests.ProcesadorPeriodo_CapaA2_3_BloquesYConsolidadoContraGolden` |
| 17 | `GoldenResumenMesYR10Tests.GoldenQ1_S3_ResumenMes_DE_EsCopia1a1DeLosRecaudos` |
| 18 | `InterventoriaTests.D3b_LEspecialesMenores_TodaLNumericaFueraDeMapas_EsCero_EnAmbosCanonicos` |
| 19 | `Regresion2026082Tests.Ejecutar_Periodo2026082_Modo5Ase_EndToEnd_DetRetri5De5VsR10` |
| 20 | `SharedFormulaR1GateTests.S1_RegenFrescaJulio_MaestrosSharedPreservados` |
| 21 | `ValidadorInteriorR1GateTests.S1_Agosto2026082_FormaInteriorVerde` |
| 22 | `ValidadorInteriorR1GateTests.S1_Julio2026072_FormaInteriorVerde` |
| 23 | `WorkbookLeafInputsTests.GenerarWorkbook_EscribeSoloLeafYPreservaFormulas` |

Muestras del `Error Message` de dos de ellos (para trazabilidad, sin diagnosticarlos aquí):

- #8: `CalculoInvalidoException : ERR-PLANTILLA: bloque ASE5 · sección '(sin sección)' · cuadro de totales · extemporáneo (Σ 'Aplicacion nuevos x reversion' − especiales de todos menos el último) debería ser una fórmula del template (celda hallada F682 de 'Reporte Componentes R1') y la plantilla trae un valor fijo.`
- #17: `Assert.Equal() Failure: Expected: 0 · Actual: 29345550`.

⚠️ **G-SUITE:** el set-diff es **exacto** (0 rojos nuevos), pero el plan asumía `suite_base` **sin rojos** (`0 Skip`; "rojo nuevo fuera de lista = STOP"). El baseline real **no es verde**. H0 no puede arreglarlos (sería código de producción fuera de alcance): se declaran y se escala la decisión (ver §9).

### §7.4 Build

`dotnet build Remuneracion.WinForms/Remuneracion.WinForms.slnx` → **Build succeeded · 0 Warning(s) · 0 Error(s)** (Core, Infrastructure, Cli, IntegrationTests y WinForms compilados con el fixture nuevo incluido).

---

## §8. Hallazgos y enmiendas requeridas al §1 del plan

> Ninguno de estos puntos se convirtió en contrato del fixture más allá del **byte medido**. Cada uno requiere enmienda del plan (o ratificación) antes de H1, según la regla del §1 ("si un número del §1 difiere en disco ⇒ STOP + enmienda").

| ID | §1 dice | Medido en disco (evidencia) | Enmienda requerida |
|---|---|---|---|
| **M1** | §1.1 (línea 113): "Template: **texto '0' o vacías**" para las 100 celdas de valor | Las **100** son `t` ausente + `<v>0</v>` (= **numéricas**, no texto, no vacías): `xl/worksheets/sheet7.xml`, filas 11/12/19/20/27/28/35/36/43/44, columnas C..L | Reescribir la naturaleza: "numéricas en 0, sin `<f>`". Impacto en H2: el writer debe escribir números (no strings) y el read-back compara decimales |
| **M2** | §1.1 (línea 112): "fila+2 = REPORTE 3 / REPORTE 4 alternados **C..L**" | Alternan en **C..N** (12 columnas: `C2='REPORTE 3'`, `D2='REPORTE 4'`, … `M2='REPORTE 3'`, `N2='REPORTE 4'`); `C..L` es el rango de las **columnas de valor** (5 ASE × 2 reportes) | Precisar: rejilla de valores `C..L`; encabezados de reporte `C..N` |
| **M3** | §1.1 (línea 109): "hoy **100 % protegida**, sin escritura" | No existe `<sheetProtection>` en la hoja. "Protegida" = las 130 celdas con `<f>` no son destino (guard anti-fórmula); las 100 celdas de valor son libres | Reescribir: "sin escritura; las celdas con `<f>` están fuera de destino (guard), no hay protección de hoja" |
| **M4** | §2.4/paso 3 y REQ-04: "columnas **D/E/F por ENCABEZADO** (nunca por índice)" | En 30/30 el encabezado del bloque 2 es `Pagos Pendientes por Aplicar | Banco | Fecha Recaudo | Cantidad | Valor` y la cabecera real es `E='ASEO'`, `F='Total'`, `D='66'`; **no hay nombres semánticos** para D/E/F. La única propiedad estable es **`D = E = F`** en toda fila de dato (Δ ≤ 1e-7) | Enmendar la regla: D/E/F se identifican como **tres copias del mismo monto** con fail-fast si divergen > 0,5 (D-06); "por encabezado" solo aplica a `C` (concepto) y a las etiquetas de canal/fuente. **HUE-3 queda cerrado**: no "manda" una columna; se exige igualdad |
| **M5** | §1.3 (línea 146): "ASE2 Sept-1 no trae OCCIDENTE" (como caso único) | **8** archivos sin columna `RECAUDO OCCIDENTE` en el bloque 3 (lista en §4.3) | Ampliar el caso testigo: la ausencia de OCCIDENTE en el bloque 3 es **frecuente y por archivo**, no anecdótica |
| **M6** | §1.1 (línea 116): "`CONSOLIDADO_TOTAL RECAUDO!L125:L130` referencia ANTICIPOS USUARIOS" | `L125:L129` sí (`-'ANTICIPOS USUARIOS'!C21..C25`); **`L130` está vacía**; `L131 = SUM(L125:L130)`; además `D125:D129` y `K125:K129` también referencian ANTICIPOS USUARIOS | Precisar el rango que referencia (`L125:L129`) y que `L131` es la consumidora de bloque |
| **M7** | §1.1 (línea 116): "`ANTICIPOS USUARIOS` 100 % fórmulas" | **CONFIRMADO** (90/90 rejilla + 18/18 sumas). Matiz de implementación: `O26` es **seguidora de compartida** (`si=3`); leerla sin expandir `si` la reportaría como "vacía/0 literal" | Sin enmienda de negocio; **sí** nota de implementación (el lector debe expandir fórmulas compartidas) |
| **M8** | §1.4 (líneas 155-162): paridad citada | **Confirmada exacta** en los 5 casos (incluido `Sept-1 ASE1 ENEL`) y ampliada a 20 archivos por `fuente × etiqueta` | Ninguna (fortalece el §1.4) |

### §8-H1 (hallazgo de manuales, para H2/H4)

`2026092 · ANT EXT-REV · J20` está **vacía** en el manual. La regla D-04 dice "concepto ausente en fuente presente = 0 explícito": H2 escribirá `0` donde el manual dejó blanco. El golden del comparador (H4) debe tratar `vacío ≡ 0` en esa celda, o declarar la diferencia; **no** se debe inventar el valor.

---

## §9. Veredicto por tarea y puerta H0→H1

| Tarea | Contenido | Veredicto |
|---|---|---|
| **T0.1** | Geometría v4 por bytes: bloque superior `<f>` (C3..N3/C4..N4/fila 5/fila 6), controles P3/P4, 5 etiquetas B8/B16/B24/B32/B40, rejilla C..L, naturaleza de las 100 celdas | **VERDE** (`T0_01a..d`) — con enmiendas M1/M2/M3 |
| **T0.2** | 100 %-fórmulas de `ANTICIPOS USUARIOS` + referencias (`D4/M4/D20/M20…`) + `CONSOLIDADO_TOTAL RECAUDO!L125:L130` como consumidora | **VERDE** (`T0_02a..d`) — con precisión M6 |
| **T0.3** | `Sheet1` por archivo 30/30: marcadores, encabezados, concepto/canal/fuente, bloque 3 con columnas variables | **VERDE** (`T0_03`, `T0_03b`) — con enmiendas M4/M5 |
| **T0.4** | Universo de fuentes 30/30 + ausencia de CIUDAD LIMPIA/ENERBIT + 0 en los 4 manuales | **VERDE** (`T0_04`) |
| **T0.5** | Paridad §1.4 re-derivada + invariante Σ-conceptos == total-fuente ±0,5 | **VERDE** (`T0_05a`, `T0_05b`) — 0 mismatches |
| **T0.6** | Divergencia docx-vs-encargo (D-10) con cita exacta | **VERDE** (`T0_06`) |
| **T0.7** | Lista cerrada de Δ de suite + huella v4 antes/después | **VERDE** (`T0_07` + §7.2/§7.3) |

### Puerta H0→H1

| Criterio | Evidencia | Estado |
|---|---|---|
| T0.1–T0.7 verdes | 18/18 tests del fixture `Passed` (0 Skip, 0 tautologías, 0 borrados) | ✅ |
| 0 warnings | `dotnet build …slnx` → 0 Warning(s) / 0 Error(s) | ✅ |
| Huella v4 intacta | `a7476ee2…8c55` antes y después | ✅ |
| G-QUEMADO parcial | `git status --porcelain`: solo el fixture nuevo y este documento; **0 literales nuevos en producción** | ✅ |
| G-H0-READONLY | cero líneas de producción modificadas; ningún `.xlsx` escrito | ✅ |
| G-SUITE set-diff | 0 rojos nuevos / 0 desaparecidos; +18 verdes | ✅ (con 23 rojos de entrada, ver §7.3) |
| `Docs/T0-Plan64-Ant-Ext-Rev.md` | este documento | ✅ |

**Estado de la puerta: ABIERTA CON DOS CONDICIONES** (no es "verde limpio"):

1. **Enmienda del plan (M1–M6)** antes de H1. Son correcciones de texto del §1/§2 contra bytes medidos, no cambios de alcance. M4 es la más importante: **"columnas D/E/F por encabezado" no es implementable** con estos bytes (no hay encabezados semánticos); la regla honesta es `D = E = F` con fail-fast.
2. **Ratificación del baseline de suite**: 23 rojos de entrada ajenos a este plan (§7.3). Hay que decidir si se corrigen en un plan hijo **antes** de H1/H2 (recomendado: #8 `ERR-PLANTILLA F682` y #9/#10 `ComparadorSalidaVsManual` tocan dinero y plantilla, y podrían enmascarar el gate H2/H4) o si se ratifican como baseline declarado.

**No hay ningún frente con "literal/rama/mapeo inventado"**: los seis frentes cerraron con bytes; lo que falta es decisión del Ingeniero.

> **Nota de forma:** el frontmatter del plan (§0.1) sigue diciendo `estado: PROPUESTO` y "H0 NO habilitado". La aprobación del Ingeniero que habilitó esta fase está registrada como decisión de proyecto (*"Plan 64 aprobado: abre H0 ANT EXT-REV"*). Conviene actualizar el frontmatter del plan al abrir H1 para que el documento no se contradiga con su propia ejecución.

---

## §10. Lo que NO se hizo (vinculante, G-H0-READONLY)

1. **Cero código de producción**: no se tocó `Remuneracion.Core`, `Remuneracion.Infrastructure`, `Remuneracion.Cli` ni `Remuneracion.WinForms`. Solo se agregó `Remuneracion.IntegrationTests/AntExtRevT0Plan64Tests.cs` (fixture `*T0Plan64*`) y este documento.
2. No se escribió ningún `.xlsx`: la plantilla y los 4 manuales se abrieron con `ZipArchive` en solo lectura; los 30 insumos con ExcelDataReader sobre **copias** temporales que se borran al terminar.
3. No se usó `Insumos.Raiz()` (raíz propia por `AGENTS.md` + rutas explícitas).
4. No se agregaron features fuera del plan, ni se crearon modelos/servicios/catálogos (eso es H1).
5. No se "ajustó" ningún número del §1 para que cuadrara: los desvíos están declarados en §8.
6. No se ejecutó `git commit` / `git push` / `git checkout` / `git reset`.
