# Plan 29 — T0 Evidencia A (T0a arbitraje julio-ASE5 BCE + T0g cierre H1/H3)

> Documento de evidencia de la Fase 0 del `plans/29 - cierre-brecha-app-vs-manual.md`.
> **Solo evidencia: cero cambios de codigo productivo, cero modificacion de `Docs/`
> (solo lectura), cero Excel/COM/ExcelDataReader/OpenXML de inspeccion.** Toda la
> inspeccion es zip+XML BCL (`System.IO.Compression` + `System.Xml`), inclusion de
> los PDF por inflado de streams Deflate (texto extraible, coste barato).
>
> **Fecha:** 2026-10-07
> **Insumos:** `Docs/Prueba Julio-2/Insumos` (5 ASE + `Conciliaciones/`), resultados
> `Docs/Prueba Julio-2/Resultado`, `Docs/Prueba2/Resultado`, plantilla
> `Docs/Plantilla_Remuneracion.xlsx` y golden `%TEMP%\opencode\Remuneracion 202607-2 Total.xlsx`.
> **Scripts (no versionados, en `%TEMP%\opencode\`):** reutilizados de la Fase T0
> (`t0a.ps1`, `t0a2.ps1`, `t0a-pdf.ps1`, `t0a-pdf2.ps1`) y dos ajustes de esta corrida
> (`t0a-pdfall.ps1` barrido de PDF; `t0a-src5.ps1` fuente ASE5; `t0g_h1.ps1` metadatos).
> Libreria `t0help.ps1` / `T0q.ps1`.

---

## 0. Veredictos (resumen)

| ID | Veredicto | Archivo + celda | Consecuencia |
|---|---|---|---|
| **T0a** | **DIVERGENCIA-DEL-MANUAL** (buscado en 88 archivos / sin aparicion) | Manual `Docs/Prueba Julio-2/Resultado/Remuneracion 202607-2 Total Administrativo.xlsx` -> `BCE SC POR FACT.!D7`=-1138650714.67 y `E7`=+1055691691.01. Fuente ASE5 `.../5-area Limpia/R4-BalanceSubsidioyContribuciones-Optimizado_..._20268465755383.xlsx` -> `Sheet1!E39`=-3615845886.78, `F39`=+2322461234.81 (`R39` = `TOTAL GENERAL`) | T1 = **mapa-vs-reader**: la app es fuente-fiel (D7/E7-calculados == E39/F39 al centavo); el unico fix de la Unidad B es la permuta de destino del mapa/llamadas (`Subsidio->D`, `Contribucion->E`). **NO se toca `LeerBalanceSc` por ASE5.** La divergencia del manual se documenta (R-B-2 / S2) |
| **T0g (H1)** | **BASE-JULIO-HEREDADA** (confirmado con bytes) | `Docs/Prueba2/Resultado/Remuneración 202608-2 Total.xlsx` -> `CONSOLIDADO_TOTAL RECAUDO!N3`=**2026072** (literal) y hojas `DetRetri2026072`, `DetValiRetri2026072`, `Informe AFaseo Recaudo 202607-2`; identicas a la base julio (`Docs/Prueba Julio-2/Resultado/Remuneración 202607-2 Total.xlsx` y `Docs/Plantilla_Remuneracion.xlsx`, mismo `N3` y mismas hojas) | **Unidad P (proceso)**: una corrida Q2 sobre la salida de otro periodo arrastra el sello de periodo `N3` y los nombres de hoja. Guia: correr cada quincena desde base propia del periodo (D-F). **Sin guardrail de bloqueo ni cambio de codigo por esto** |
| **T0g (H3)** | **LOCALIZADO** (linea citada) | `Remuneracion.WinForms/bin/Debug/net10.0-windows/remuneracion_log_20261005.txt` lineas **1701** y **2367** | Contexto del swap BCE (hipotesis lider): la traza de esperados BCE por ASE ya asumia el arreglo D/E. Se cita como antecedente de la Unidad B; sin cambio de codigo |

**Nota de hecho H1 (matiz):** el `N3`/nombres de hoja **no** se sellan por periodo y por eso
se heredan; en cambio `G7`/`K7`/`D6` **si** se sellan con las fechas del periodo corrido
(agosto: `G7`=46250, `K7`=46265, `D6`=46238; julio: `G7`=46219, `K7`=46234) — coherente con
Plan 28 (sello de fechas) y con la hipotesis H1 del Plan 29 (solo los metadatos de periodo
quedan heredados).

---

## 1. Evidencia

### 1.1 T0a — Arbitraje julio-ASE5 (BCE)

**Valores objetivo del manual (leidos por zip+XML):**

| Celda (manual, `BCE SC POR FACT.`) | Valor | Tipo |
|---|---|---|
| `D2` | `SUBSIDIO` | literal texto (encabezado) |
| `E2` | `CONTRIBUCION` | literal texto (encabezado) |
| `D7` | **-1138650714.67** | literal numerico |
| `E7` | **+1055691691.01** | literal numerico |
| `F7` | -82959023.66 | formula `D7+E7` + cache |

**App julio (`Docs/Prueba Julio-2/Resultado/Remuneración 202607-2 Total.xlsx`, `BCE SC POR FACT.`):**
`D7`=+2322461234.81, `E7`=-3615845886.78, `F7`= formula `D7+E7` (cache 0 pre-recalculo).
Los dos literales de la app coinciden, permutados, con el `TOTAL GENERAL` de la fuente:

**Fuente ASE5 (`Sheet1`, fila 39 = `TOTAL GENERAL`; encabezados R1: `E1`=Subsidio, `F1`=Contribucion, `G1`=Valor):**

| Celda fuente | Valor | App julio | Match |
|---|---|---|---|
| `E39` (Subsidio) | -3615845886.78 | `E7` app | == al centavo |
| `F39` (Contribucion) | +2322461234.81 | `D7` app | == al centavo |
| `G39` (Valor = E39+F39) | -1293384651.97 | `F7`=D7+E7 | == (pre-recalculo, 0 como cache) |

Es decir: la app escribe Contribucion->D y Subsidio->E (permuta V4); el manual de ASE5
trae **otros valores** (`D7`=-1.138e9, `E7`=+1.055e9) que **no provienen de esta fuente**.

**Barrido de los valores del manual (tolerancia ±0.01, mas estricta que ±0.5):**

| Universo | Archivos | Buscado | Resultado |
|---|---|---|---|
| Todos los `.xlsx` de `Insumos` (recursivo) + `Resultado` | **48** | -1138650714.67 / +1055691691.01 / +1138650714.67 / 1055691691 / -82959023.66 | **0 apariciones** en insumos. `-82959023.66` aparece solo **dentro del propio manual** (`BCE!E23`/`F7`, `CONSOLIDADO_TOTAL RECAUDO!J108`/`J13`, `GERENTES_ENEL!J13`, `REMUNERACION_ENEL!J108`/`J13`, `Valida -Remunera!J47`/`J13`) |
| Todos los `.pdf` de `Insumos` (recursivo), texto inflado normalizado a digitos | **40** | 1138650714 / 1055691691 / 82959023 | **0 apariciones**. El PDF de ASE5 `R4-BalanceSubsidioyContribuciones-Optimizado_..._20268465751283.pdf` **si** contiene `3615845886` y `2322461234` (valores de la app) |

Cobertura adicional: el script `t0a2.ps1` escanea cada workbook completo (todas las hojas,
todas las celdas con valor numerico) y el `t0a.ps1` dirigido a `5-*` + `Conciliaciones`
(archivos `R4-BalanceSubsidioyContribuciones-Optimizado*`, `RecaudosReversados*`,
`ReportePagosxBanco*`, `RetribuciónNegativa*`, `RerpoteDetalleSaldosaFavor*`,
`SaldosaFavorAplicadosPorNotas*`, `ReversiónPorComponente*`, `Recaudoporcomponente*`,
`Conjunta ENEL/ENERBIT/Otros/Recip-072026`, `Directa-072026`) tambien devuelve 0.
El par del manual tampoco aparece en ningun otro ASE ni conciliacion.

**Extracto crudo (t0a.ps1):**

```
===== Manual julio BCE D3:F7 =====
D2: [LITERAL] disp=[SUBSIDIO]
E2: [LITERAL] disp=[CONTRIBUCION]
D7: [LITERAL] v=[-1138650714.6700001]
E7: [LITERAL] v=[1055691691.01]
F7: [FORMULA+cached] f=[D7+E7] v=[-82959023.660000086]
===== App julio BCE D3:F7 =====
D7: [LITERAL] v=[2322461234.81]
E7: [LITERAL] v=[-3615845886.78]
F7: [FORMULA+cached] f=[D7+E7] v=[0]
########## BUSCANDO -1138650714.67 ##########   (0 hits)
########## BUSCANDO 1055691691.01 ##########    (0 hits)
```

**Extracto crudo (t0a-src5.ps1) fuente ASE5:**

```
FILE: R4-BalanceSubsidioyContribuciones-Optimizado_..._20268465755383.xlsx
--- SHEETS --- [Sheet1] -> xl/worksheets/Sheet1.xml
R1  : ... E1[v=Subsidio] F1[v=Contribución] G1[v=Valor]
R39 : A39[v=TOTAL GENERAL] E39[v=-3615845886.78] F39[v=2322461234.81] G39[v=-1293384651.97]
```

**Veredicto T0a: DIVERGENCIA-DEL-MANUAL.** Buscado el par `D7/E7` del manual en 48 `.xlsx`
(todos los insumos incluidos ASE5, RecaudosReversados y ReportePagosxBanco, mas
Conciliaciones) y 40 `.pdf` (texto extraible): **sin aparicion**. La fuente respalda a la
app al centavo. => Sin cambio de codigo en el reader; la Unidad B se limita a la permuta de
destino (mapa-vs-reader) y la divergencia del manual se documenta.

### 1.2 T0g-H1 — Metadatos de agosto (base heredada)

Metodo: `t0g_h1.ps1` compara zip+XML de la salida agosto contra la salida julio y la
plantilla canonica.

**Hojas clave y sello de periodo (`CONSOLIDADO_TOTAL RECAUDO!N3`):**

| Archivo | `N3` | `DetRetri*` | `DetValiRetri*` | `Informe AFaseo*` | `G7` / `K7` / `D6` |
|---|---|---|---|---|---|
| `Docs/Prueba2/Resultado/Remuneración 202608-2 Total.xlsx` (agosto app) | **2026072** | `DetRetri2026072` | `DetValiRetri2026072` | `Informe AFaseo Recaudo 202607-2` | 46250 / 46265 / 46238 |
| `Docs/Prueba Julio-2/Resultado/Remuneración 202607-2 Total.xlsx` (julio app) | 2026072 | `DetRetri2026072` | `DetValiRetri2026072` | `Informe AFaseo Recaudo 202607-2` | 46219 / 46234 / 46238 |
| `Docs/Plantilla_Remuneracion.xlsx` (canonica) | 2026072 | `DetRetri2026072` | `DetValiRetri2026072` | `Informe AFaseo Recaudo 202607-2` | 46219 / 46234 / 46238 |
| `%TEMP%/opencode/Remuneracion 202607-2 Total.xlsx` (golden) | 2026072 | `DetRetri2026072` | `DetValiRetri2026072` | `Informe AFaseo Recaudo 202607-2` | 46219 / 46234 / 46238 |

Los 4 archivos tienen 40 hojas con exactamente los mismos nombres. La salida de agosto
**hereda** `N3`=2026072 y los nombres `DetRetri2026072`/`DetValiRetri2026072`/
`Informe AFaseo Recaudo 202607-2` de la base julio (donde ya estaban esos valores), mientras
que `G7`/`K7` si quedan sellados con las fechas de agosto (46250/46265). Confirma H1:
metadata de periodo heredado; fechas de proceso/corte selladas por corrida.

**Veredicto T0g-H1: BASE-JULIO-HEREDADA (confirmado).** Consecuencia: Unidad P (proceso),
correr Q2 desde base propia del periodo; documentar en `project-context.md` + manual.

### 1.3 T0g-H3 — Log de agosto (veredicto D/E)

Archivo: `Remuneracion.WinForms/bin/Debug/net10.0-windows/remuneracion_log_20261005.txt`.

Linea literal buscada ("veredicto D/E T0 hipotesis lider") localizada 2 veces:

```
1701: 2026-10-05 13:06:48.305 -05:00 [INF] (RunId=8b569e88-207b-440e-8460-48c1a866c5de Periodo=2026082 AseId= Hoja=BCE SC POR FACT. Validacion=) BCE SC POR FACT.: esperados post-Excel por ASE (veredicto D/E T0 hipótesis líder).
2367: 2026-10-05 14:39:13.923 -05:00 [INF] (RunId=27ebe055-8448-4036-bea4-cfc3707ae3b3 Periodo=2026082 AseId= Hoja=BCE SC POR FACT. Validacion=) BCE SC POR FACT.: esperados post-Excel por ASE (veredicto D/E T0 hipótesis líder).
```

RunIds citados: `8b569e88-207b-440e-8460-48c1a866c5de` (2026-10-05 13:06:48.305 -05:00) y
`27ebe055-8448-4036-bea4-cfc3707ae3b3` (2026-10-05 14:39:13.923 -05:00), ambos `Periodo=2026082`.

**Veredicto T0g-H3: LOCALIZADO.** Antecedente del swap BCE (la traza de esperados asumia el
orden D/E); se documenta, sin cambio de codigo.

---

## 2. Reproducibilidad

```powershell
# cwd = raiz del repo; PowerShell 5.1; -File (evita problemas de encoding de acentos)
powershell -NoProfile -ExecutionPolicy Bypass -File "$env:TEMP\opencode\t0a.ps1"
powershell -NoProfile -ExecutionPolicy Bypass -File "$env:TEMP\opencode\t0a2.ps1"
powershell -NoProfile -ExecutionPolicy Bypass -File "$env:TEMP\opencode\t0a-pdf2.ps1"
powershell -NoProfile -ExecutionPolicy Bypass -File "$env:TEMP\opencode\t0a-pdfall.ps1"
powershell -NoProfile -ExecutionPolicy Bypass -File "$env:TEMP\opencode\t0a-src5.ps1"
powershell -NoProfile -ExecutionPolicy Bypass -File "$env:TEMP\opencode\t0g_h1.ps1"
```

Metodo: apertura por `ZipArchive` + parseo `XmlDocument` con namespace
`.../spreadsheetml/2006/main`; se leen `f` (formula) y `v` (cache/literal) por celda, con
`sharedStrings` resuelto. PDF: extraccion del texto de streams `FlateDecode` y normalizacion
a digitos para busqueda. Ninguna corrida usa Excel, COM, ExcelDataReader ni OpenXML.

---

## 3. Cierre

- T0a: arbitraje cerrado -> DIVERGENCIA-DEL-MANUAL (app fuente-fiel; sin cambio de reader).
- T0g: H1 cerrado (metadatos de agosto heredados de la base julio) y H3 cerrado (log D/E citado).
- Sin `NEEDS_CONTEXT`. Sin cambios de codigo productivo ni de `Docs/`.
