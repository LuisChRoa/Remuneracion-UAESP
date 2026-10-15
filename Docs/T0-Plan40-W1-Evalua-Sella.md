# T0 (Plan 40) — W-1 evalúa + writer sella: evidencia congelada desde bytes

> **Fixture que manda:** `Remuneracion.IntegrationTests/GeometriaPlantillaEstandarizadaT0Tests.cs`
> (extiende el T0 del Plan 38/39; reutiliza `ContratoSlotR1T0Tests`). Lectura SOLO con BCL
> (`System.IO.Compression.ZipArchive` + `System.Xml.Linq`): sin Excel/COM, sin DocumentFormat.OpenXml,
> sin ExcelDataReader. **Cero código de producción en T0.**
> **Oráculo geométrico:** `Docs/Plantilla estandarizada/Formato Remuneracion 2026xx-x Total.xlsx`
> (período interno base **2026092**). **Salida real de hoy:**
> `Docs/Prueba Julio-2/Resultado/Resultado5/Remuneración 202607-2 Total.xlsx`.
> **Salida legacy (falso-verde por herencia):** `.../Resultado4/Remuneración 202607-2 Total.xlsx`.
> Los números citados son citas del disco (dump BCL 2026-10-14); el test compara código-vs-fixture.

## 0. Hashes (R-INPLACE — el oficial NUNCA se modifica)

| Archivo | SHA256 |
|---|---|
| Oficial (antes y después del T0) | `3086AC317595C2F8E84ED24B1C5EE3E52565DD5193394E3062878E7FA49C5214` |
| Resultado5 (salida real de hoy) | `1CE66248E9963E9C0E7503DE6C8F66A7AE59189718C9203C70C4414D90A74F5E` |
| Resultado4 (legacy) | `5A87AA933509ED7FB79774922D718363F22A2D4CAD5A818FA4E3067811151C65` |

## 1. Los 5 sub-bloques W-1 de `VALIDACION_TOTAL` (texto-`<f>` EXACTO)

| Celda | Texto-`<f>` (oficial, byte-fiel) | Caché `<v>` oficial | Caché `<v>` hoy | Escrito por la app |
|---|---|---|---|---|
| `C15` | `C14=C9` (`t="b"`) | `0` (FALSE) | `0` (FALSE) | no (fórmula) |
| `D25` | `D24=C9` (`t="b"`) | `1` (TRUE) | `1` (TRUE) | no (fórmula) |
| `O25` | `O24=N9` (`t="b"`) | `1` (TRUE) | `1` (TRUE) | no (fórmula) |
| `D34` | `D33='CONSOLIDADO_TOTAL RECAUDO'!D109` (`t="b"`) | `1` (TRUE) | `1` (TRUE) | no (fórmula) |
| `F34` | `F33='CONSOLIDADO_TOTAL RECAUDO'!D109+'CONSOLIDADO_TOTAL RECAUDO'!N109` (`t="b"`) | `1` (TRUE) | `1` (TRUE) | no (fórmula) |

### 1.1 Cadenas de dependencia (congeladas)

Cadena del incidente (recaudo, 100 % escrita por la app):

| Hoja | Celda | Texto-`<f>` | Naturaleza |
|---|---|---|---|
| `VALIDACION_TOTAL` | `C14` | `'Valida - Control Recaudo'!F10` | literal escrito (F10) |
| `VALIDACION_TOTAL` | `C9` | `SUM(C3:C8)` | suma |
| `VALIDACION_TOTAL` | `C3` | `'Recaudo EAAB Reciprocidad'!F21+'Recaudo EAAB + Ciud Limp'!F21+'Recaudo Directa Occidente'!F21+'Recaudo ENEL'!F21+'Recaudo ENERBIT'!F21` | suma de recaudos |
| `VALIDACION_TOTAL` | `C4`..`C7` | análogas con `F22`..`F25` | suma de recaudos |
| `VALIDACION_TOTAL` | `C8` | `...!F26+...` (5 refs) | suma de recaudos |
| `VALIDACION_TOTAL` | `D24` | `SUM(D20:D23)`; `D20`..`D23` = `VALIDACION_{emp}!C9` → `SUM(C3:C8)` → `'Recaudo *'!F21:F26` | recaudo |
| `VALIDACION_TOTAL` | `N9` | `SUM(N3:N8)`; `N3` = `SUM(I3:M3)` → `'CONSOLIDADO_TOTAL RECAUDO'!D104…` | consolidado (recaudo) |
| `VALIDACION_TOTAL` | `D33` | `SUM(D29:D32)`; `D29` = `'REMUNERACION …'!D109` | **remuneración** |
| `VALIDACION_TOTAL` | `F33` | `SUM(F29:F32)`; `F29` = `D29+E29` | **remuneración** |
| `CONSOLIDADO_TOTAL RECAUDO` | `D109` | `SUM(D104:D108)` (maestro shared si=61, `ref=D109:N109`) | consolidado |
| `VALIDACION_TOTAL` | `O24` | shared si=5 (maestro `H24`, `SUM(H20:H23)`, `ref=H24:O24`) → efectiva `SUM(O20:O23)` | **remuneración** |
| `CONSOLIDADO_TOTAL RECAUDO` | `N109` | follower de si=61 → efectiva `SUM(N104:N108)` | consolidado |

### 1.2 Estado de la cadena C15 (oficial vs hoy)

| Celda | Oficial | Hoy (Resultado5) |
|---|---|---|
| `'Valida - Control Recaudo'!F10` | literal `73981716599` (residuo del ejemplo base 2026092) | literal `72497949967` (**escrito por la app** = Σ `Recaudo *`!F21:F26) |
| `C14` | `=F10`; caché `73981716599` | `=F10`; caché `73981716599` (**viejo**) |
| `C9` | `=SUM(C3:C8)`; caché `0` | `=SUM(C3:C8)`; caché `0` (**viejo**) |
| `C15` | `=C14=C9`; caché `0` (FALSE) | `=C14=C9`; caché `0` (**viejo, FALSE**) |

**Lectura:** la app SÍ escribió `F10 = 72497949967` (= Σ recaudo), pero los cachés de `C14`/`C9`/`C15`
quedaron heredados de la plantilla → el gate viejo (que leía el caché) fallaba en falso. La
evaluación desde lo escrito da `C15 = (72497949967 = 72497949967) = TRUE`.

## 2. Diferencial de clase (oficial → salida de hoy, alineado por índice de hoja)

Se comparan TODAS las celdas de fórmula de las 40 hojas (el período renombra 3 hojas, por eso la
alineación es por índice):

| Clase | Conteo | Detalle |
|---|---|---|
| `<f>` DISTINTA | **11** | las 11 `<f>` del token de período reescritas por Phase 0: `BCE SC POR FACT.!H3:H7` (`DetRetri2026092!J9:J13` → `…2026072…`), `H10`, `DetValiRetri!K24:K28` |
| `<f>` IDÉNTICA con `<v>` distinta | **103** | consumidores de `N3` resellados por Phase 0 (`REMUNERACION_*!N3/N22/N41/N60/N79/N98/M117/AD3`, `GERENTES_*`, `AJUSTES - SF-T!N3/N22/N41/AC3/AC22/AC41`, `CONSOLIDADO!N3/N22/…/M117/AD3/AD22/…`, `Valida -Remunera!N121/AD121`): `2026092` → `2026072` |

**Nuance decisiva (D-A):** la clase "caché stale" del incidente (C14/C9/C15) **NO aparece en este
diferencial** porque la app nunca recalcula: `C15` es `0` en oficial y en hoy (idéntico). Por eso el
diferencial oficial→salida NO basta para localizar cachés mentirosos — la doctrina "el gate EVALÚA"
(y no lee caché) es la única red honesta. Los 114 cambios del diferencial son todos de Phase 0
(reescritura de token + resello de `N3`), no divergencias de negocio.

## 3. Allow-list Phase 0 (16 celdas) + propuesta de ampliación F-NEW-1

- **Vigente (Plan 39 T0):** 16 celdas `<f>` = 11 tokens de período (§2 del T0 del Plan 39) + 5
  `CONSOLIDADO_TOTAL RECAUDO!L9:L13` (R-INTERV-Q). Congelada por `T0_AllowListFase0_16Celdas`.
- **Propuesta F-NEW-1 (T3, fuera de esta corrida):** 4 familias de auto-chequeo que hardcodean
  `INTERVENTORIA!M26:M31` y evaluarían FALSE en Q1 tras R-INTERV-Q:
  - `DetRetri{per}!L32:L36` = `INTERVENTORIA!M26:M30 + L23:L27 = L9:L13` (5 celdas)
  - `DetRetri{per}!O32:O36` = `(… − INTERVENTORIA!M26:M30) … = O9:O13` (5 celdas)
  - `DetValiRetri{per}!I24:I29` = `−INTERVENTORIA!M26:M31 + I16:I21 = I9:I14` (6 celdas)
  - `INTERVENTORIA!S26:S31` = `Q26=R26` (maestro shared) (6 celdas)
  - Ampliación propuesta: **16 → 16 + 22** (sujeta al T0 propio de T3; prohibido sellar TRUE).

## 4. Hallazgo de T0 que el plan no anticipó (se reporta, no se oculta)

El T0, al congelar las cadenas, resolvió los 5 sub-bloques desde los bytes con un resolutor BCL
acotado (sumas + referencias + fórmulas compartidas). Resultado **verificado contra el manual del
administrativo recalculado por Excel** (`Remuneracion 202607-2 Total Administrativo.xlsx`):

| Sub-bloque | Cadena | Manual (Excel) | Salida real (app) |
|---|---|---|---|
| `C15` | recaudo | TRUE | **TRUE** (evaluado; caché 0) |
| `D25` | recaudo | TRUE | **TRUE** |
| `O25` | remuneración (`O24` vs `N9`) | TRUE | **FALSE** (`61.117.131.901,59` vs `73.137.739.247,63`) |
| `D34` | remuneración (`D33` vs `D109`) | TRUE | **FALSE** (`60.433.731.796,50` vs `72.439.437.413,21`) |
| `F34` | remuneración (`F33` vs `D109+N109`) | TRUE | **FALSE** (`60.561.923.648,84` vs `72.582.530.994,88`) |

**Interpretación (honesta):** `C15`/`D25` son 100 % de RECAUDO (la app las escribe) → el gate las
**evalúa**. `O25`/`D34`/`F34` traen el **espejo de REMUNERACIÓN** (`VALIDACION_*!N9` →
`REMUNERACION_*!D104` → `'Reporte Componentes R1'!F85 = F62+F36+F8-L8-L36`), que la app **no
puebla** (la salida real deja la remuneración corta frente al consolidado, mientras el manual del
administrativo SÍ cierra). Por tanto **no son evaluables "desde lo escrito"** y quedan con veredicto
**requiere decisión (T5)** — NO gateadas ni selladas (sellar/gatear en falso sería la mentira que
D-F/R-COSMETICO prohíben). El diccionario `SubBloquesValidacionTotal` expone el **veredicto del gate**
y la evidencia cruda (operandos + `Computable`) queda visible: la divergencia NO se oculta.

## 5. Doctrina instalada (D-A)

**Los cachés de fórmulas no son oráculo; el gate EVALÚA y el writer SELLA.**

- `ValidacionOracleReader` (T1): los 5 sub-bloques ya no se leen del caché `<v>`; se asertó su
  texto-`<f>` contra la tabla `FormasValidacionTotal` y se evaluaron los operandos resueltos desde lo
  escrito por un resolutor BCL acotado (sumas/referencias/compartidas).
- `ValidadorBasico` (T1, W-1): aserta el texto y gatea `Igualdad(izq, der)` ±0.5; misma severidad y
  mensaje `ASE n · VALIDACION_TOTAL · <celda>`.
- `OpenXmlPlantillaWriter` (T2): sella por read-back la cadena `C3:C8→C9→C14→C15` (allow-list del
  writer), preservando el `<f>` byte-idéntico, antes del saneamiento calcChain.
