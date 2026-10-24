# T0 — Plan 57: anclas de validación derivadas de la plantilla

**Fecha:** 2026-10-23 · **Fase:** H0 (read-only, cero código de producción) · **Fixture:**
`Remuneracion.IntegrationTests/AnclasDerivadasT0Plan57Tests.cs` (7/7 verde) · **Método:** BCL puro
(`System.IO.Compression.ZipArchive` + `System.Xml.Linq`) sobre rutas explícitas. Sin Excel/COM, sin
OpenXML SDK, sin ExcelDataReader, sin `Insumos.Raiz()`. Cero escritura.

**Plantillas (G-SHA, verificadas antes y después):**

| Plantilla | SHA256 |
|---|---|
| v2 `Docs/Plantilla estandarizada/v2/Formato Remuneracion 2026xx-x Total vf.xlsx` | `0F192AD2721C77986B6A8CB85D1A92F782DCD2F69989F362E066FA17027F2DAB` |
| v1 `Docs/Plantilla estandarizada/Formato Remuneracion 2026xx-x Total.xlsx` | `3086AC317595C2F8E84ED24B1C5EE3E52565DD5193394E3062878E7FA49C5214` |

---

## 1. Inventario de lo congelado (S1.1)

`AnclasExtempR1` (Plan 49, medición 2026-10-19) es la **única** tabla de anclas congelada del
contrato del oráculo:

| Entrada | Naturaleza declarada (Plan 49) | Naturaleza real en la v2 | Veredicto |
|---|---|---|---|
| `Mapa[5] = "F682"` (`F660+F629-L629`) | literal | **`<f>F660+F629-L629</f>`** | desactualizada → causa del `ERR-PLANTILLA` |
| `Mapa[5] = "F702"` (`F659+F628-L628`) | literal | **`<f>F659+F628-L628</f>`** | desactualizada |
| `Mapa[1..4]` | (no existen: "ASE1-4 traen `<f>`") | `<f>` | correcto por omisión, no por derivación |
| `PrimeraFilaConsolidadoExtemp = 47` + `CeldaConsolidadoExtemp(aseId)` | dirección | `D47..D51` del CONSOLIDADO | **geometría de publicación**, no ancla del interior: se conserva (es la celda del núcleo) |

**Causa raíz medida (por qué la foto nació mal y no solo «envejeció»).** El Plan 49 midió las anclas sobre la
**plantilla v1** (`Docs/Plantilla estandarizada/Formato Remuneracion 2026xx-x Total.xlsx`), donde
`F682`/`F702` sí son literales (`<c r="F682" s="234"><v>0</v></c>`, verificado por bytes). Pero desde el
Plan 48 el pipeline usa la **v2** (`…/v2/…vf.xlsx`), donde la misma celda trae
`<f>F660+F629-L629</f>`. La tabla siguió apuntando a la naturaleza de OTRO archivo: el error era
inevitable en toda corrida 5-ASE. Re-congelarla con la fórmula de la v2 habría repetido el defecto con
el próximo cambio de plantilla, de ahí el retiro (§2-D2.3).


Cotas revisadas (T0.1): `R1Q2ProtectedPorAse` (10 `<f>` TOT_OPT/EXTEMP, Plan 54 H1.7) es un
inventario de **fórmulas protegidas del escritor**, no un mapa de anclas del oráculo: no se toca
(prohibido por §1.3). `WorkbookLeafCellMapQ2.R1Q2ProtectedPorAse` se verificó simétrico con los
bytes (5 ASE × {TOT_OPT, EXTEMP}).

**Número exacto de anclas congeladas en el oráculo: 2** (`F682`, `F702`), ambas de ASE5.

## 2. Naturaleza real en la v2 por bloque (S1.1 — veredicto de homogeneidad)

Ancla EXTEMP del bloque = celda que declara el `<f>` de `CONSOLIDADO_TOTAL RECAUDO!D{46+aseId}`
(§2-D3.1: la posición se **lee**, no se declara):

| Bloque | Publicación | Ancla (hoja!celda) | Naturaleza | Fórmula (evidencia, **no** se congela) |
|---|---|---|---|---|
| ASE1 | `CONSOLIDADO!D47` | `Reporte Componentes R1!F77` | `<f>` | `F59+F32-L32` |
| ASE2 | `CONSOLIDADO!D48` | `Reporte Componentes R1!F256` | `<f>` | `F190+F146-L146` |
| ASE3 | `CONSOLIDADO!D49` | `Reporte Componentes R1!F417` | `<f>` | `F342+F315-L315` |
| ASE4 | `CONSOLIDADO!D50` | `Reporte Componentes R1!F570` | `<f>` | `F530+F486-L486` |
| ASE5 | `CONSOLIDADO!D51` | `Reporte Componentes R1!F682` | `<f>` | `F660+F629-L629` |

**Veredicto: HOMOGÉNEO.** Las 5 anclas son fórmula, con la misma forma (`Σ aplicaciones ±
especiales`, cf. Plan 54 F6) y sólo referencias y signos — exactamente lo que
`ResolutorValorValidacion` sabe evaluar. **No hay STOP**: el patrón es derivable.

Familia secundaria (§2-D3.2: consumidores `<f>` de las hojas `REMUNERACION_*`), publicada en la
**misma fila** del bloque (D47..D51), 6 destinos por bloque en la v2:

| Bloque | Destinos declarados en la fila | Naturaleza |
|---|---|---|
| ASE1 (fila 47) | `F77` `F82` `F87` `F92` `F97` `F102` | fórmula: `F77 F87 F92 F97`; literal: `F82 F102` |
| ASE2 (fila 48) | `F256` `F259` `F264` `F269` `F274` `F279` | fórmula: `F256 F259 F264 F274`; literal: `F269 F279` |
| ASE3 (fila 49) | `F415` `F420` `F425` `F430` `F435` `F440` | fórmula: `F415 F425 F435`; literal: `F420 F430 F440` |
| ASE4 (fila 50) | `F568` `F573` `F578` `F583` `F588` `F593` | fórmula: `F568 F573 F578 F588`; literal: `F583 F593` |
| ASE5 (fila 51) | `F682` `F687` `F692` `F697` `F702` `F707` | fórmula: `F682 F692 F702`; literal: `F687 F697 F707` |

La familia secundaria **mezcla** fórmula y literal. Eso **no** es heterogeneidad no prevista: es
justo el caso que §2-D2.3 manda resolver en runtime (si trae `<f>` → evaluar con los valores del
workbook; si es literal → leerlo), sin ramas por bloque ni por quincena. El consumidor estable de la
familia existe y se leyó por bytes: `REMUNERACION_OCCIDENTE_ Directa!D51 = 'Reporte Componentes
R1'!F702` (aserción del fixture) → **vía elegida: §2-D3.1/D3.2 por referencia declarada**, no por
etiqueta.

## 3. Cierre contra el R10 de los tres períodos (S1.3)

DetRetri leído por bytes de `R10_Remuneracion_{período}.xlsx` (hoja `DetRetri{período}`, columna D,
filas ASE 1..5) y comparado al peso con la tabla CA-A3 del plan:

| Período | ASE1 | ASE2 | ASE3 | ASE4 | ASE5 |
|---|---|---|---|---|---|
| 2026071 | 16.758.167.585 | 12.227.292.053 | 10.127.195.599 | 10.553.776.376 | 8.543.663.209 |
| 2026072 | 17.450.228.673 | 20.516.143.970 | 15.221.896.467 | 7.179.595.396 | 12.073.344.662 |
| 2026101 | 14.842.425.136 | 11.302.637.605 | 8.823.787.300 | 10.483.172.762 | 7.550.668.665 |

15/15 al peso: la tabla CA-A3 se re-deriva del R10 sin residuo. La cadena de derivación de cada
valor (`fuente × <f>`: R1 por firma → `Reporte Componentes R1!F{ancla}` → `CONSOLIDADO!D{46+aseId}` →
`D104:D108` → R10) la ejercita end-to-end la corrida CLI de H1 (CA-A3), que debe reproducir estos
mismos 15 números sin mover un peso.

## 4. Consumidores de `AnclasExtempR1` (S2.4 anticipado)

| Consumidor | Uso | Veredicto |
|---|---|---|
| `ValidacionOracleReader.ResolverAnclasExtempR1` | `PorAse` (ancla + expresión) | **retirar** (objeto del plan) |
| `ValidacionOracleReader.ResolverExtempConsolidado` | `HojaConsolidado`, `CeldaConsolidadoExtemp` | conservar (celda del núcleo, §2-D3.1) |
| `ValidadorBasico.ValidarExtempAbsolutoPorAse` | `Hoja`, `HojaConsolidado`, `CeldaConsolidadoExtemp` | conservar el núcleo; el interior pasa a consumir la evidencia derivada (H0: `ValidacionOracleReader.cs:193` → `ValidadorBasico.cs:662-671`) |
| Tests | sólo comentarios/documentación | re-anclar el comentario |

## 5. Delta de suite esperado (S1.4)

H0 **no** toca `Remuneracion.Core/` ni `Remuneracion.Infrastructure/`: el diff de producción es
vacío. Altas esperadas: **7** tests nuevos del T0 (5 hechos + `[Theory]` ×3 períodos). Rojos nuevos
esperados: **ninguno**. SHAs v1/v2 intactos antes/después.

---

## 6. Hallazgo posterior a H0 (H1, 2026-10-23) — bloqueo pre-existente ENMASCARADO

Al retirar la guarda del ancla, la corrida real ya no aborta en el oráculo y llega al validador. Ahí
aparece un **segundo bloqueo, ajeno a este plan y anterior a él**, que el `ERR-PLANTILLA` ocultaba
(el lector del oráculo corría ANTES del juicio, así que ninguna corrida 5-ASE había llegado nunca a
`VALIDACION_TOTAL`):

```
[ERR-VALIDACION] ASE n · VALIDACION_TOTAL · O25/D34/F34: el sub-bloque evaluado desde lo escrito es falso
(58210117742.11 != 58210094820.50)  ·  2026071, Δ = 22.921,61
(53002710082.05 != 53002691468.18)  ·  2026101, Δ = 18.513,87
```

- **Qué compara:** `D34` = `D33` vs `'CONSOLIDADO_TOTAL RECAUDO'!D109`, con
  `D33 = SUM(D29:D32)` = Σ `D109` de las cuatro hojas de remuneración por empresa
  (`'REMUNERACION Reciprocidad EAAB'`, `REMUNERACION_ENEL`, `REMUNERACION_ENERBIT`,
  `'REMUNERACION_OCCIDENTE_ Directa'`) y `D109 = SUM(D104:D108)` = Σ `DetRetri` de los 5 ASE.
  `O25`/`F34` arrastran la misma brecha.
- **Δ constante por período y por ASE** (las tres celdas son globales): no es un error de lectura ni
  de escritura del pipeline — el `DetRetri` cierra con el R10 al peso en los 15 cruces (§3).
- **Correlación:** `2026072` (Q2, sin conciliación por empresa por `R-RECORTE-Q2`) cierra con
  **exit 0**; los dos Q1 (`2026071`, `2026101`), que sí aplican la conciliación por empresa,
  arrastran el Δ. La brecha vive en el agregado por empresa vs el consolidado.
- **Estado:** deuda PRE-EXISTENTE, **fuera del alcance del Plan 57** (§1.3 prohíbe tocar
  `ResolvedorConciliacionEmpresa`, la conciliación y los agregados). El workbook de salida es
  **celda-idéntico** antes y después de este plan (60.159 celdas, mismo hash de `(celda, <f>, <v>)`
  en la corrida 2026101 previa y posterior), de modo que el plan no movió ningún peso.
- **Requiere** plan propio con su T0 (clase Plan 55) para cerrar `CA-A1`/`CA-A2` de los períodos Q1.

