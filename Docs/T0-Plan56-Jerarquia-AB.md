# T0-Plan 56 — Jerarquía sección/subsección de la plantilla R1 (identidad A/B)

> **Estado:** H0 cerrado en matriz **15/15** (5 bloques × julio-1 / julio-2 / octubre-1), read-only por bytes.
> **Alcance:** S1.1 a S1.5 del `plans/56 - jerarquia-seccion-subseccion-plantilla-r1.md`. **Cero código de producción.**
> **Fixture:** `Remuneracion.IntegrationTests/JerarquiaABT0Plan56Tests.cs` (6 tests, 6 verdes).
> **Fecha de corrida:** 2026-10-23. Este documento + el fixture SON la evidencia (S0 `G-H0-READONLY`).

---

## 1. Método (S1.1/S1.2/S1.3)

- **BCL puro**: `System.IO.Compression.ZipArchive` + `System.Xml.Linq`. Sin Excel/COM, sin OpenXml SDK, sin ExcelDataReader.
- **Rutas explícitas** por período (prohibido `Insumos.Raiz()` en lo nuevo):
  - plantilla: `Docs/Plantilla estandarizada/v2/Formato Remuneracion 2026xx-x Total vf.xlsx`, hoja `Reporte Componentes R1`;
  - fuentes: `Docs/Prueba Julio-1|Julio-2|Octubre-1/Insumos/<n-empresa>/Recaudoporcomponente_*.xlsx`, hoja `Sheet1`;
  - oráculo: `Docs/Prueba <período>/Insumos/R10_Remuneracion_<AAAAMMQ>.xlsx`, hoja `DetRetri<AAAAMMQ>`.
- **Lectura con `FileShare.ReadWrite | Delete`** (permite leer mientras Excel tiene el archivo abierto) y **cero escritura**: SHA256 del v2 verificado antes y después de la medición completa (§9).
- **Oráculo de cierre**: texto-`<f>` + literales; el `<v>` de una celda con `<f>` **nunca** se usa (S0 `G-WORKBOOK`). El evaluador del fixture soporta refs, `+`/`-` y `SUM(rango)`, con doble negación semántica de Excel (`--L112` ⇒ `+L112`, ver §10.3).

---

## 2. Modelo medido (definiciones congeladas)

| Nivel | Columna | Valores observados | Cierre |
|---|---|---|---|
| **Sección** | **A** | `AFaseo`, `Componente`, `Subs/Cont` (`Total` = gran total del bloque) | fila con `A=<sección>` y `B='Total'` |
| **Subsección** | **B** | `Mes`, `Aplicacion nuevos x reversion` | fila con `B=<subsección>` y `C='Total'` |

- **Identidad del cuadro** = `(sección A, subsección B, ocurrencia de B dentro de A)`.
- Dentro del cuadro: **fuente `C` → canal `D` → concepto `E`**; los valores son **componentes por encabezado** (`F`..`Z` en la plantilla, `F`..`U` en la fuente), nunca por posición.
- **Canal efectivo** (Plan 46 §2.1): si la fila no trae `D`, el canal es el `D` de la fila resumen `…/Total` que cierra su grupo (primera hacia abajo con `D` no vacía y distinta de `Total`).
- Las etiquetas se escriben **al cierre**: la fila de cierre de la subsección (`B=<sub>`) precede a la de la sección (`A=<sec>`).
- El **gran total** del bloque es la fila `A='Total'` con `B` vacía; las **tablas resumen** (`C='TOTAL'` / `D='OPORTUNO'|'EXTEMPORANEO'`) viven **después** del gran total del bloque y **antes** del bloque siguiente.

---

## 3. Lado PLANTILLA — `Reporte Componentes R1` (v2) (S1.1)

Cabecera de componentes = fila anterior al primer dato de cada bloque (`F='Total'`, `L='SERVICIO ESPECIALES'`, `Z='Rec.Intereses'`, idénticos en los 5 bloques).

| Bloque | Empresa (B de la cabecera−1) | Cab. | Datos | `AFaseo` | `Componente` | `Subs/Cont` | Gran total |
|---|---|---|---|---|---|---|---|
| 1 | PROMOAMBIENTAL | 5 | 6–72 | 6–17 · `Mes` 6–16 (cierra **B16**) | 18–48 · `Aplicacion` 18–32 (**B32**), `Mes` 33–47 (**B47**) | 49–71 · `Aplicacion` 49–59 (**B59**), `Mes` 60–70 (**B70**) | **A72** |
| 2 | LIME | 109 | 110–210 | 110–121 · `Mes` 110–120 (**B120**) | 122–172 · `Aplicacion` 122–146 (**B146**), `Mes` 147–171 (**B171**) | 173–209 · `Aplicacion` 173–190 (**B190**), `Mes` 191–208 (**B208**) | **A210** |
| 3 | CIUDAD LIMPIA | 288 | 289–355 | 289–300 · `Mes` 289–299 (**B299**) | 301–331 · `Aplicacion` 301–315 (**B315**), `Mes` 316–330 (**B330**) | 332–354 · `Aplicacion` 332–342 (**B342**), `Mes` 343–353 (**B353**) | **A355** |
| 4 | BOGOTA LIMPIA | 449 | 450–550 | 450–461 · `Mes` 450–460 (**B460**) | 462–512 · `Aplicacion` 462–486 (**B486**), `Mes` 487–511 (**B511**) | 513–549 · `Aplicacion` 513–530 (**B530**), `Mes` 531–548 (**B548**) | **A550** |
| 5 | AREA LIMPIA | 602 | 603–676 | 603–614 · `Mes` 603–613 (**B613**) | 615–649 · `Aplicacion` 615–629 (**B629**), `Mes` 630–648 (**B648**) | 650–675 · `Aplicacion` 650–660 (**B660**), `Mes` 661–674 (**B674**) | **A676** |

- **Forma idéntica en los 5 bloques** (aserción del fixture): `AFaseo:[Mes#1] | Componente:[Aplicacion#1, Mes#1] | Subs/Cont:[Aplicacion#1, Mes#1]` ⇒ 5 cuadros por bloque, 25 en la hoja.
- Cada cierre de sección escribe `A=<sección>` con `B='Total'`; cada cierre de subsección escribe `B=<subsección>` con `C='Total'`.
- Nota de lectura frente al plan §0.5: la notación «`AFaseo` 602–613 / `Componente` 615–648 / `Subs/Cont` 650–674» cita la **fila de cabecera** como extremo inferior y la **fila de cierre de la subsección** como superior. Las identidades coinciden exactamente con las medidas aquí (Bloque 5: `Mes` cierra en 613, `Aplicacion` en 629/660, `Mes` en 648/674).

---

## 4. Lado FUENTE — `Recaudoporcomponente` (S1.1)

Misma partición A/B leída del archivo del período (la etiqueta `A=` viaja en la fila de cierre; la `B=` en la fila-frontera). Rangos de **datos+cierre**; «total» = fila `A='Total'` (cierre del bloque).

| Período | ASE | Secciones (rango) | Subsecciones (rango / cierre) | Total |
|---|---|---|---|---|
| 2026071 | 1 | `Componente` 5–25 · `Subs/Cont` 26–41 | `Aplic` 5–9, `Mes` 10–24 · `Aplic` 26–29, `Mes` 30–40 | 42 |
| 2026071 | 2 | `AFaseo` 5–15 · `Componente` 16–38 · `Subs/Cont` 39–55 | `Mes` 5–14 · `Mes` 16–37 · `Mes` 39–54 | 56 |
| 2026071 | 3 | `AFaseo` 5–12 · `Componente` 13–33 · `Subs/Cont` 34–49 | `Mes` 5–11 · `Aplic` 13–17, `Mes` 18–32 · `Aplic` 34–37, `Mes` 38–48 | 50 |
| 2026071 | 4 | `AFaseo` 5–12 · `Componente` 13–46 · `Subs/Cont` 47–72 | `Mes` 5–11 · `Aplic` 13–23, `Mes` 24–45 · `Aplic` 47–55, `Mes` 56–71 | 73 |
| 2026071 | 5 | `AFaseo` 5–12 · `Componente` 13–32 · `Subs/Cont` 33–47 | `Mes` 5–11 · `Mes` 13–31 · `Mes` 33–46 | 48 |
| 2026072 | 1 | `AFaseo` 5–12 · `Componente` 13–32 · `Subs/Cont` 33–48 | `Mes` 5–11 · `Aplic` 13–16, `Mes` 17–31 · `Aplic` 33–36, `Mes` 37–47 | 49 |
| 2026072 | 2 | `AFaseo` 5–12 · `Componente` 13–50 · `Subs/Cont` 51–78 | `Mes` 5–11 · `Aplic` 13–24, `Mes` 25–49 · `Aplic` 51–59, `Mes` 60–77 | 79 |
| 2026072 | 3 | `AFaseo` 5–9 · `Componente` 10–30 · `Subs/Cont` 31–46 | `Mes` 5–8 · `Aplic` 10–14, `Mes` 15–29 · `Aplic` 31–34, `Mes` 35–45 | 47 |
| 2026072 | 4 | `AFaseo` 5–12 · `Componente` 13–50 · `Subs/Cont` 51–76 | `Mes` 5–11 · `Aplic` 13–24, `Mes` 25–49 · `Aplic` 51–57, `Mes` 58–75 | 77 |
| 2026072 | 5 | `Componente` 5–34 · `Subs/Cont` 35–55 | `Aplic` 5–14, `Mes` 15–33 · `Aplic` 35–40, `Mes` 41–54 | 56 |
| 2026101 | 1 | `AFaseo` 5–9 · `Componente` 10–34 · `Subs/Cont` 35–53 | `Mes` 5–8 · `Aplic` 10–18, `Mes` 19–33 · `Aplic` 35–41, `Mes` 42–52 | 54 |
| 2026101 | 2 | `AFaseo` 5–12 · `Componente` 13–39 · `Subs/Cont` 40–58 | `Mes` 5–11 · `Aplic` 13–23, `Mes` 24–38 · `Aplic` 40–46, `Mes` 47–57 | 59 |
| 2026101 | 3 | `AFaseo` 5–9 · `Componente` 10–25 · `Subs/Cont` 26–37 | `Mes` 5–8 · `Mes` 10–24 · `Mes` 26–36 | 38 |
| 2026101 | 4 | `AFaseo` 5–12 · `Componente` 13–47 · `Subs/Cont` 48–71 | `Mes` 5–11 · `Aplic` 13–31, `Mes` 32–46 · `Aplic` 48–59, `Mes` 60–70 | 72 |
| 2026101 | 5 | `Componente` 5–24 · `Subs/Cont` 25–39 | `Mes` 5–23 · `Mes` 25–38 | 40 |

Hechos estructurales (verificados por el fixture en 15/15):

1. La fuente **no siempre trae `AFaseo`** (`2026071`-ASE1, `2026072`-ASE5, `2026101`-ASE5) y **no siempre trae `Aplicacion`** (`2026101`-ASE3/ASE5, `2026071`-ASE2/ASE5). Toda combinación que trae está **declarada** por la plantilla (§3): no hay combinación `(A, B, ocurrencia)` huérfana.
2. Las ocurrencias son siempre **#1** dentro de su sección en los 15 casos (la plantilla declara #1 y #1 de `Aplicacion`/`Mes` por sección).
3. Los códigos de canal **dependen del emisor del reporte**, no del período: `E/H/O/T` (bloques 1, 3, 4, 5) y `1/2/5` (bloque 2). No se normalizan: la identidad los compara literales, igual que hoy.

---

## 5. Cruce `(A, B, ocurrencia, C, D, E)` → celda declarada (S1.2)

Regla aplicada (sin fallback posicional, sin listas por bloque):

1. la fila cae en el cuadro `(A, B, ocurrencia)` que la **cierra**;
2. su firma es `C | D(efectivo) | E`;
3. el destino es la fila declarada del **mismo cuadro** de la plantilla con la misma firma; si la firma se repite dentro del cuadro (p. ej. dos grupos `D=E`), se empareja por **ordinal dentro de la firma** (k-ésima ↔ k-ésima);
4. los componentes se mapean por **encabezado** (`F`..`U` fuente → `F`..`Z` plantilla);
5. los cierres también son destino declarado: `B=<sub>/C=Total` por identidad de cuadro, `A=<sec>/B=Total` por etiqueta de sección y `A=Total` (gran total) por etiqueta.

### Matriz 15/15

| Período | ASE | Secciones | Subsecciones | Filas de cuadro | Firmas resueltas | Sin destino en su cuadro | Huérfanas del bloque |
|---|---|---|---|---|---|---|---|
| 2026071 | 1 | 2 | 4 | 31 | 42 | 0 | 0 |
| 2026071 | 2 | 3 | 3 | 45 | 52 | **3** (§7) | 0 |
| 2026071 | 3 | 3 | 5 | 37 | 50 | 0 | 0 |
| 2026071 | 4 | 3 | 5 | 60 | 72 | 0 | 0 |
| 2026071 | 5 | 3 | 3 | 37 | 47 | 0 | 0 |
| 2026072 | 1 | 3 | 5 | 36 | 47 | 0 | 0 |
| 2026072 | 2 | 3 | 5 | 66 | 78 | 0 | 0 |
| 2026072 | 3 | 3 | 5 | 34 | 45 | 0 | 0 |
| 2026072 | 4 | 3 | 5 | 64 | 76 | 0 | 0 |
| 2026072 | 5 | 2 | 4 | 45 | 55 | 0 | 0 |
| 2026101 | 1 | 3 | 5 | 41 | 52 | 0 | 0 |
| 2026101 | 2 | 3 | 5 | 46 | 58 | 0 | 0 |
| 2026101 | 3 | 3 | 3 | 27 | 33 | 0 | 0 |
| 2026101 | 4 | 3 | 5 | 59 | 71 | 0 | 0 |
| 2026101 | 5 | 2 | 2 | 31 | 32 | 0 | 0 |
| **Total** | | | | **659** | **777** | **3** | **0** |

- **777 firmas resueltas** = 659 filas de datos + cierres de subsección + cierres de sección + gran total de bloque, en los 15 cruces.
- **0 huérfanas del bloque** (ninguna combinación presente en la fuente falta en todo el bloque de la plantilla).
- **0 ambigüedades** por cardinalidad inversa (ninguna firma de la fuente excede las filas declaradas por la plantilla).
- La aserción `T0_Cruce_IdentidadAB_15de15` **no admite** fallback: no existe en el fixture ningún camino que asigne por proximidad, orden físico global ni «próximo libre».
- Cardinalidades «fuente < plantilla» (ausencia ⇒ cero del template, §2-D3) son la norma: el fixture las reporta por cuadro y no son error.

---

## 6. Cierre aritmético R1 — `TOT_OPT` / `EXTEMP` (S1.3)

Anclas localizadas **por etiqueta** dentro de la tabla resumen de cada bloque (`C='TOTAL'` + `D='OPORTUNO'`; la fila `D='EXTEMPORANEO'` del mismo grupo no repite `C`):

| Bloque | `TOT_OPT` | `EXTEMP` |
|---|---|---|
| 1 | `F75` | `F77` |
| 2 | `F254` | `F256` |
| 3 | `F415` | `F417` |
| 4 | `F568` | `F570` |
| 5 | `F680` | `F682` |

Texto-`<f>` de las anclas (congelado por el fixture):

- **B1**: `F75 = F47+F70+F16-L16-L47` · `F77 = F59+F32-L32`
- **B2**: `F254 = F171+F208-L171+F120-L120` · `F256 = F190+F146-L146`
- **B3**: `F415 = F330+F353+F299-L299-L330` · `F417 = F342+F315-L315`
- **B4**: `F568 = F511+F548+F460-L460-L511` · `F570 = F530+F486-L486`
- **B5**: `F680 = F674+F648+F613-L613-L648` · `F682 = F660+F629-L629`

Doble derivación (ambas comparadas con `Assert.Equal`, **sin tolerancia**):

- **(a)** ancla evaluada con texto-`<f>` + literales del template + los valores que declara el cruce;
- **(b)** reconstrucción independiente desde la fuente: `TOT_OPT = ΣF(Mes) − ΣL(Mes sin la última)` y `EXTEMP = ΣF(Aplicación) − ΣL(Aplicación sin la primera)` (semántica T0-Plan54 §3).

**(a) ≡ (b) en 15/15 al centavo.**

### Resultado por bloque y período

| Período | ASE | `TOT_OPT` (ancla) | `EXTEMP` (ancla) | R1 = TOT_OPT+EXTEMP | R10 `RECAUDO TOTAL` (`DetRetri` D9:D13) |
|---|---|---|---|---|---|
| 2026071 | 1 | 16.704.332.434,57 (F75) | 11.673.020,00 (F77) | 16.716.005.454,57 | 16.758.167.585 |
| 2026071 | 2 | 12.157.780.441,19 (F254) | 0,00 (F256) | 12.157.780.441,19 | 12.227.292.053 |
| 2026071 | 3 | 10.101.988.514,82 (F415) | 10.198.723,07 (F417) | 10.112.187.237,89 | 10.127.195.599 |
| 2026071 | 4 | 10.552.409.503,83 (F568) | 5.419.780,00 (F570) | 10.557.829.283,83 | 10.553.776.376 |
| 2026071 | 5 | 8.519.310.329,28 (F680) | 0,00 (F682) | 8.519.310.329,28 | 8.543.663.209 |
| 2026072 | 1 | 17.339.793.559,86 (F75) | 49.580,00 (F77) | 17.339.843.139,86 | 17.450.228.673 |
| 2026072 | 2 | 20.352.268.540,72 (F254) | 56.190.070,00 (F256) | 20.408.458.610,72 | 20.516.143.970 |
| 2026072 | 3 | 15.226.209.207,13 (F415) | 1.307.680,00 (F417) | 15.227.516.887,13 | 15.221.896.467 |
| 2026072 | 4 | 7.047.492.494,73 (F568) | 4.343.101,00 (F570) | 7.051.835.595,73 | 7.179.595.396 |
| 2026072 | 5 | 12.033.011.685,71 (F680) | 441.849,00 (F682) | 12.033.453.534,71 | 12.073.344.662 |
| 2026101 | 1 | 14.815.021.974,75 (F75) | 648.870,00 (F77) | 14.815.670.844,75 | 14.842.425.136 |
| 2026101 | 2 | 11.219.101.305,09 (F254) | 9.704.974,00 (F256) | 11.228.806.279,09 | 11.302.637.605 |
| 2026101 | 3 | 8.811.472.160,34 (F415) | 0,00 (F417) | 8.811.472.160,34 | 8.823.787.300 |
| 2026101 | 4 | 10.473.967.813,42 (F568) | 21.557.177,00 (F570) | 10.495.524.990,42 | 10.483.172.762 |
| 2026101 | 5 | 7.544.253.279,94 (F680) | 0,00 (F682) | 7.544.253.279,94 | 7.550.668.665 |

### 6.1 Relación con el R10 — precisión obligatoria

El R10 **no es un oráculo de R1**: publica `RECAUDO TOTAL` (`D9:D13`), que es el oráculo del **DetRetri** consolidado
(`CONSOLIDADO!D104 = D9+D28+D47+D66+D85 = TOT_OPT + R2 + EXTEMP + R4 + AJUSTES`, `D9=D9:TOT_OPT`, `D47=EXTEMP`,
verificado en `ValidadorBasico.ValidarDetRetriContraR10`, tolerancia ±0,5). Por eso `TOT_OPT` **no** puede igualar
`D9:D13` por construcción; el plan §0.5 lo enuncia como igualdad por conflación de columnas (`D9` medido vs `D104` del R10).

La cadena completa sí cierra **al peso** con las piernas no-R1 ya congeladas por T0-Plan54 §3 (octubre-1):

| ASE | `TOT_OPT` | `EXTEMP` | `R2` (D28) | `R4` (D66) | `AJUSTES` Q1 | Σ | R10 | Δ |
|---|---|---|---|---|---|---|---|---|
| 1 | 14.815.021.974,75 | 648.870,00 | 30.588.071,02 | −3.833.780,00 | 0 | 14.842.425.135,77 | 14.842.425.136 | **0,23** |
| 2 | 11.219.101.305,09 | 9.704.974,00 | 95.634.974,32 | −21.803.648,45 | 0 | 11.302.637.604,96 | 11.302.637.605 | **0,04** |
| 3 | 8.811.472.160,34 | 0,00 | 14.574.412,03 | −2.259.272,00 | 0 | 8.823.787.300,37 | 8.823.787.300 | **0,37** |
| 4 | 10.473.967.813,42 | 21.557.177,00 | 29.206.357,07 | −41.558.585,40 | 0 | 10.483.172.762,09 | 10.483.172.762 | **0,09** |
| 5 | 7.544.253.279,94 | 0,00 | 10.379.542,05 | −3.964.157,00 | 0 | 7.550.668.664,99 | 7.550.668.665 | **0,01** |

- **5/5 dentro de ±0,5** (redondeo del R10 al peso). `R2`/`R4`/`AJUSTES` están **fuera del alcance de H0** (Planes 49/50/53); se citan de T0-Plan54 §3, no se re-derivan aquí.
- Para **julio-1 y julio-2** no existe descomposición publicada de `D28`/`D66`: la cadena `D104 = R10` de esos períodos es la que T0-Plan54 §3 ya declaró **5/5 al peso**, y los valores R10 usados aquí son los mismos que el proyecto congeló (`plans/24 - T0 Evidencia.md` para julio-2 y `Docs/T0-Residual-24329-OccidenteAse5.md` para julio-2 ASE5). El fixture **pinea** esa identidad: `R10 leído == R10 documentado` en 15/15.
- Los valores `TOT_OPT`/`EXTEMP` medidos en octubre-1 coinciden **al centavo** con `D9`/`D47` de T0-Plan54 §3 (5/5): es la verificación R10-anclada del cierre R1.

---

## 7. Combinaciones sin destino (S1.4)

**Huérfanas del bloque (destino inexistente, §2-D3): 0 en 15/15.** No hay ninguna combinación presente en la fuente y ausente en todo el bloque de la plantilla.

**Filas sin destino en SU cuadro (pero declaradas en otro cuadro del mismo bloque): 3, todas en un único cruce.**

| Período | ASE | Cuadro | Fila fuente | Firma (`C|D|E`) | ¿Dónde sí existe en el bloque? | Veredicto |
|---|---|---|---|---|---|---|
| 2026071 | 2 | (`AFaseo`, `Mes`, #1) | 8 | `·|5|·Vlr Servicio` | R132, R157 | **No se reacomoda.** Queda en el cero del template; si H1 la considera error ⇒ `ERR-SIN-SLOT` accionable (§2-D3), nunca remap. |
| 2026071 | 2 | (`AFaseo`, `Mes`, #1) | 9 | `·|5|·Total` | R134, R159, R181, R199 | ídem |
| 2026071 | 2 | (`AFaseo`, `Mes`, #1) | 10 | `NUEVO ESQUEMA|Total|·` | R135, R160, R182, R200 | ídem |

Lectura: la fuente de julio-1 ASE2 emite en su bloque `AFaseo` un grupo `C='NUEVO ESQUEMA'` con `D='5'` (F = −16.473,56 en las tres filas) que la plantilla **solo declara en los cuadros `Componente`/`Subs/Cont`**, no en `AFaseo` (cuyos canales son `E/H/O/T`). No es un defecto de la jerarquía A/B: el cuadro `(AFaseo, Mes, #1)` resuelve y su **total de cuadro** (`F14` fuente → `F120` plantilla) entra en el cierre; lo que no tiene destino son tres **filas internas** de ese cuadro.

- Impacto en el dinero: **ninguno** en el cierre R1 (el ancla consume el total del cuadro, no sus filas internas). Ver §6: `TOT_OPT` de `2026071`-ASE2 = 12.157.780.441,19 por doble derivación exacta.
- Impacto en H1: el bloque `AFaseo` es hoy la única sección que el Plan 46 dejó en la ruta heredada (`ParticipaDeSeccionConformada` excluye `AFaseo`). Decidir si esas 3 filas son error accionable o cero documentado **es una decisión del Ingeniero**, no de H0. Se deja como **condición abierta H1** (§10.4).

---

## 8. Veredicto de las dos muestras conocidas (verificación, no alcance)

| Muestra del plan §0.5 | Medición H0 | Veredicto |
|---|---|---|
| Muestra 1 — «sección `Componente`/`Subs/Cont` del bloque de Área Limpia en octubre-1», `TOT_OPT = 7.544.253.279,94` | `2026101` ASE5: (`Componente`,`Mes`,#1) = 630–648, (`Subs/Cont`,`Mes`,#1) = 661–674, (`AFaseo`,`Mes`,#1) = 603–613; `F680 = F674+F648+F613−L613−L648 = 7.544.253.279,94` | **CONFIRMADA al centavo**; la fila `C=ENERBIT, D=Total` cae en **F637** (no en la sección `AFaseo`) |
| Muestra 2 — «otro bloque», `TOT_OPT = 8.811.472.160,34` | `2026101` ASE3: `F415 = F330+F353+F299−L299−L330 = 8.811.472.160,34` | **CONFIRMADA al centavo** |
| «y sus equivalentes en los otros bloques/períodos» | §5 + §6: los 15 cruces resuelven por la misma ruta; los 15 cierres se reproducen por doble derivación exacta | **GENERALIZADO 15/15** |

El defecto `ERR-SIN-SLOT` de la corrida `--periodo 2026101 --cinco-ase` queda explicado por bytes: el resolvedor vigente
tipa la «sección» con la frontera de `B` e ignora `A`; la primera sección de la fuente de ASE5 es `(Componente, Mes)` y se
empareja con la `AFaseo` del template, donde `C=ENERBIT/D=Total` no está declarada. Con la identidad A/B, esa fila cae en
`F637`. La misma causa aparece hoy en la suite en `2026071`-ASE2 (`ERR-SIN-SLOT: ASE 2`), es decir **no es un caso de octubre**.

---

## 9. Plantillas: SHAs y estado (S0 `G-SHA`, `G-H0-READONLY`)

| Plantilla | SHA256 esperado por el plan | SHA256 medido (antes) | SHA256 medido (después) | Veredicto |
|---|---|---|---|---|
| v1 `Formato Remuneracion 2026xx-x Total.xlsx` | `3086AC31…C5214` | `3086AC317595C2F8E84ED24B1C5EE3E52565DD5193394E3062878E7FA49C5214` | ídem | **IDÉNTICA** |
| v2 `v2/Formato Remuneracion 2026xx-x Total vf.xlsx` | `0F192AD2…7F2DAB` | `EAFF76D69634E69D1BB4C7EF2DD3AC707AAF312C6E83CC6CD54F84C6D5374218` | `EAFF76D6…374218` | **H0 no escribió** (antes ≡ después), pero **no coincide con el ancla del plan** |

### 9.1 Desviación pre-existente del v2 (no introducida por H0)

- El blob **commiteado** (`HEAD:Docs/Plantilla estandarizada/v2/…vf.xlsx`) **sí** tiene el SHA del plan: `0F192AD2…7F2DAB` (693.400 bytes).
- El archivo **en disco** tiene `EAFF76D6…374218` (693.968 bytes) y `mtime = 2026-10-23 10:12` (posterior al último commit `b2e4fdf`, 08:58), con Excel abierto sobre él (PID 79844). Es un **re-guardado de Excel**, anterior a esta corrida.
- Contenido semántico **idéntico**: comparación pieza por pieza de las 162 entradas del ZIP y celda por celda de las 40 hojas (15.560 celdas en `Reporte Componentes R1`): **0 diferencias** (fórmulas, valores y textos). Sólo cambian `docProps/core.xml`, `styles.xml`, `workbook.xml`, `vmlDrawing1.vml` y la serialización de las hojas.
- Consecuencia: la medición de §3–§6 es la misma sobre el blob commiteado o sobre el archivo en disco.
- **Acción requerida:** el Ingeniero debe decidir si restaura el v2 desde git (recupera el SHA del plan) o **ratifica el re-guardado**. Mientras no se ratifique, `G-SHA` no puede declararse "ancla del plan" para el v2; el fixture sólo puede garantizar «antes ≡ después».

---

## 10. Gates, hallazgos secundarios y condiciones abiertas

### 10.1 Gates de H0

| Gate | Estado |
|---|---|
| `G-BUILD` (`dotnet build Remuneracion.WinForms/Remuneracion.WinForms.slnx`) | **0 warnings / 0 errores** |
| `G-SHA` v1 | **verde** (idéntica al ancla) |
| `G-SHA` v2 | **antes ≡ después** verde; **ancla del plan**: desviación pre-existente (§9.1) |
| `G-FORMULA` | **verde por construcción**: H0 no abre ningún workbook en escritura; cero `<f>` tocadas |
| `G-H0-READONLY` | **verde**: el diff de `Remuneracion.Core/` y `Remuneracion.Infrastructure/` es **vacío**; sólo se agrega el fixture `JerarquiaABT0Plan56Tests.cs` y este doc |
| `G-SUITE` | **PENDIENTE / no evaluable hoy**: la plantilla v2 está abierta en Excel y 291 de los 305 rojos de la corrida son `IOException: … being used by another process` en `PreparadorBasePipeline`. Conteo de la corrida: **1490 / 1185 / 305 / 0** (14 rojos no-lock, ya presentes en el mundo de deudas de los planes 49–55, p. ej. `SeccionPorOperandoR1T0Tests` con diferencias de redondeo decimal, `BasePeriodoPreflightTests…2026071` con `ERR-SIN-SLOT: ASE 2`, que es este mismo `R-JERARQUIA-AB`). **Requiere cerrar Excel y re-correr para el set-diff contra `p55-H3-suite.trx`.** |

### 10.2 Hallazgo: la firma no se forward-fillea

La `C` de la fila se compara **literal** (no se propaga desde la fila anterior). Es lo que hace el escritor vigente
(`ColumnasCCompatibles`: vacío vs no-vacío ⇒ no encaja) y es lo que permite que las filas `Vlr Servicio`/`Vlr Intereses`
sin `C` de un grupo se distingan por **canal efectivo**. Con forward-fill de `C` el cruce produce 87 falsos huérfanos: la
identidad correcta es `C literal + D efectivo + E`.

### 10.3 Hallazgo: `--` en el ancla del bloque 2

`F264` (fila `ENEL`/`OPORTUNO` de la tabla resumen del bloque 2) trae `F193+F150+F112--L112-L150`: doble negación ⇒ **suma**
`L112` en vez de restarlo. No afecta al cierre (`TOT_OPT` usa `F254`), pero es una **deuda de plantilla** que el fixture
evalúa con semántica Excel (no se "corrige"). Se reporta para no confundir un futuro rojo.

### 10.4 Condiciones abiertas para H1 (decisión del Ingeniero)

1. **Las 3 filas del §7** (`2026071`-ASE2, `AFaseo`): ¿error accionable `ERR-SIN-SLOT` o cero documentado? Hoy `AFaseo` está fuera de la alineación por sección (Plan 46). Cualquier decisión debe respetar «sin remap, sin reacomodo».
2. **Ancla v2** (§9.1): restaurar desde git o ratificar el re-guardado.
3. **Cierre `D104 = R10` en julio-1/julio-2**: si se quiere reproducir al peso *dentro* de H0, hay que autorizar lectura de `R2`/`R4`/`AJUSTES` (hoy fuera de alcance por §1.2/§1.3). H0 ya prueba lo que le corresponde: los dos operandos R1 exactos y la cadena de octubre 5/5.

---

## 11. Reproducción

```powershell
dotnet build Remuneracion.IntegrationTests/Remuneracion.IntegrationTests.csproj   # 0 warnings
dotnet test  Remuneracion.IntegrationTests/Remuneracion.IntegrationTests.csproj `
  --filter "FullyQualifiedName~JerarquiaABT0Plan56" --logger "console;verbosity=detailed"
```

Salida esperada: **6/6 verdes** — `T0_Plantillas_Intactas_AntesYDespues`, `T0_PlantillaV2_SeccionesYSubsecciones_PorBloque`,
`T0_Fuentes_SeccionesYSubsecciones_PorPeriodoAse`, `T0_Cruce_IdentidadAB_15de15`,
`T0_CierreAritmetico_TotOptYExtemp_15de15`, `T0_CombinacionesSinDestino_ListaCerradaVacia`.
