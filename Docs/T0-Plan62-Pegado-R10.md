# T0 — Plan 62: pegado R10 (DetRetri/DetValiRetri) y validación del recaudo en CONSOLIDADO 142–147

**Fecha:** 2026-10-28 · **Fase:** H0 (read-only, cero código de producción) · **Fixture:**
`Remuneracion.IntegrationTests/PegadoR10T0Plan62Tests.cs` (**12/12 verde**) · **Método:** BCL puro
(`System.IO.Compression.ZipArchive` + `System.Xml.Linq`) sobre rutas EXPLÍCITAS. Sin Excel/COM, sin
DocumentFormat.OpenXml, sin ExcelDataReader, sin `Insumos.Raiz()`. Cero escritura y cero corridas de
producción.

**Fuentes medidas (SHA256 antes y después del fixture — sin cambios):**

| Rol | Archivo | SHA256 |
|---|---|---|
| Plantilla oficial (VIGENTE) | `Docs/Plantilla estandarizada/Formato Remuneracion 2026xx-x Total.xlsx` | `3086AC317595C2F8E84ED24B1C5EE3E52565DD5193394E3062878E7FA49C5214` |
| R10 del caso | `Docs/Prueba Septiembre 1/Insumos/R10_Remuneracion_2026091.xlsx` | `DD0EAA7B6C03180F1B360F3D83EDADD7B4CE18B921478C6AAC775C9F50E91122` |
| Manual (paridad) | `Docs/Prueba Septiembre 1/Resultados/Administrativo/Remuneracion 202609-1 Total.xlsx` | `3CE1B86967CD709520272224D60686875AD1AD2EE5C510A531BE09726E420916` |

> La huella de la plantilla coincide con la citada por `Docs/T0-Plan57-Anclas-Derivadas.md` §8 ⇒ **no
> hay cambio de plantilla**; no aplica el STOP del `gate_entrada`. La base interna es 2026092 ⇒
> `sheet36` = `DetRetri2026092`, `sheet37` = `DetValiRetri2026092`, `sheet18` = `CONSOLIDADO_TOTAL RECAUDO`.

> **[Fe de erratas Plan 62 — H5]** El rol "Plantilla oficial (VIGENTE)" de la tabla anterior cita la
> **v1** (`Docs/Plantilla estandarizada/Formato Remuneracion 2026xx-x Total.xlsx`). El pipeline NO
> consume esa ruta: consume `Insumos.PlantillaOficialV4` =
> `Docs/Plantilla estandarizada/v4/Formato Remuneracion 2026xx-x Total v4.xlsx`. Verificado en H4: la
> geometría medida por los asserts de este doc (`CONSOLIDADO!E142:E146`, `L145` máster `si=68`,
> `M142` máster `si=67`, DetRetri A3:O19 / DetValiRetri A3:O14) es **idéntica en v1 y v4**, por lo que
> ningún assert T0 ni T4 se invalida. La v1 se conserva en la tabla como evidencia de la medición
> original (bytes), no como la plantilla de producción.

**Mapa de asserts:** T0-01…T0-04 (T0a, §2) · T0-05…T0-06 (T0b, §3) · T0-07…T0-09 (T0c, §4) ·
T0-10…T0-12 (T0d, §5). Las desviaciones del texto del plan frente al byte se declaran en **§6** y son
la única puerta abierta de H0.

---

## 1. T0a — R10 2026091: 2 hojas con el layout de la plantilla

### T0-01 — Exactamente 2 hojas, nombres compuestos por dominio

| Hecho | Byte |
|---|---|
| Conteo de `<sheet>` en `xl/workbook.xml` | **2** (una tercera ⇒ FAIL) |
| Hoja 0 | `DetRetri2026091` = `NombresHojaPeriodo.DetRetri(Periodo.Parse("2026091"))` |
| Hoja 1 | `DetValiRetri2026091` = `NombresHojaPeriodo.DetValiRetri(…)` |
| Partes | `xl/worksheets/sheet1.xml`, `sheet2.xml` (dimensiones `B3:O19` y `B3:O14`) |

### T0-02 — Encabezado corporativo, sello de proceso y rango

Común a las 2 hojas (texto exacto, todo `t="s"`):

| Celda | Literal |
|---|---|
| `F3` | `PROCERASEO S.A.S.` |
| `H3` | `NIT: 901 153 639 - 0` |
| `F4` | `3. ESQUEMA FINANCIERO DEL SERVICIO PUBLICO DE ASEO` |
| `F5` | `Resolución UAESP Número 27 de 2018` |
| `C6` / `D6` | `Fecha de Proceso:` / `17/09/2026` |
| `C7` / `D7` | `Hora:` / `08:44 AM` |
| `F7` / `G7` | `Fecha Desde:` / `01/09/2026` |
| `I7` / `J7` | `Fecha Hasta:` / `15/09/2026` |
| fila 8 | encabezado de tabla, 14 columnas `B..O` (`B8=ASE`, `C8=CONCESIONARIOS`) |
| `B9:B13` | `1`,`2`,`3`,`4`,`5` |
| `C14` | `Total` |
| Código de remuneración | `2026091` en **`M3`** (DetRetri) y **`K3`** (DetValiRetri) |

> El sello `D6`/`D7` viaja en el propio R10 ⇒ el retiro de `SellarProcesoEnDetRetri` (D-31SELLO) no
> pierde dato: el pegado lo trae por construcción.

### T0-03 — Montos como TEXTO monetario (no número con formato)

Los **72 montos por hoja** (`D9:O14`) son `t="s"` (shared string) con la convención es-CO del lector
(`ExcelDataReaderDetRetriR10Reader.ParsearNumeroMonetario`): `^\-?\$[0-9.]+$`.

| Muestra | Celda | Texto | Parseo |
|---|---|---|---|
| Positivo | `DetRetri2026091!D9` | `$17.450.250.543` | `17.450.250.543` |
| Negativo | `DetRetri2026091!J10` | `-$5.093.259.797` | `-5.093.259.797` |
| Negativo (DetVali) | `DetValiRetri2026091!I9` | `-$189.185.987` | `-189.185.987` |

No hay ninguna celda `t="n"` en el bloque de montos: **el R10 no usa número con formato moneda** ⇒ el
conversor (y su extracción a `ParseoMonetarioR10`, T1.2) sigue siendo imprescindible.

### T0-04 — Columna A vacía y layout idéntico a la plantilla

Matriz R10-vs-plantilla por hoja (R10 `…2026091` ↔ plantilla `…2026092`):

| Elemento | R10 | Plantilla | Veredicto |
|---|---|---|---|
| Columna `A` (filas 3–19 / 3–14) | 0 celdas | 0 celdas | **idéntico** |
| `B9:B13` | `1..5` (texto) | `1..5` (texto) | **idéntico** |
| `C9:C13` | 5 rótulos | 5 rótulos | **idéntico** (texto exacto) |
| `C11` (ASE3) | `CIUDAD LIMPIA BOGOTA S.A. ESP` (sin punto, H-59-4) | igual | **idéntico** |
| Encabezado fila 8 (`B8:O8`, 14 celdas) | DetRetri: `ASE`,`CONCESIONARIOS`,`RECAUDO TOTAL`,… ; DetVali: `…(TCS)`,`…(TLU)`,… | igual por hoja | **idéntico** |
| Bloque de encabezado filas 3–7 (`A:O`) | celdas ocupadas | mismas celdas | **idéntico** |
| Fila 14 | `C14=Total`, `D14:O14` montos | `C14=Total`, `D14:O14` en 0 | **idéntico (estructura)** |
| `D9:O14` | montos del período | literales `0` | difiere por diseño (es el destino) |
| `D6`/`D7`, `M3`/`K3` | `17/09/2026`, `08:44 AM`, `2026091` | `02/10/2026`, `10:43 AM`, `2026092` | difiere por período |
| `C16` (pie) | `* Costos de Interventoría: REMUNERACIÓN 202609-1` | `* Costos de Interventoría: remuneración 2026031` | difiere por período |

**Consecuencia para H2:** el rango `A3:O19` incluye el pie 15–19 (`C16`, `E18`, `F19`). El pegado
reemplaza el pie heredado de la plantilla (`…2026031`) por el del R10 (`…202609-1`), que es lo que
hace el manual. Es un desplazamiento **declarado** del diff, no una regresión.

---

## 2. T0b — Plantilla `DetRetri2026092` (`sheet36.xml`)

### T0-05 — Top-block values-only

- **0 celdas con `<f>` en `A3:O19`** (barrido XML completo del rango).
- Los 72 slots `D9:O14` son literales numéricos `<v>0</v>` (`t` ausente) ⇒ destinos de escritura.
- Sello y rótulos: `M3=2026092`, `C6/D6=Fecha de Proceso:/02/10/2026`, `C7/D7=Hora:/10:43 AM`,
  `F6=R10 - Recaudo Consolidado`, `B8=ASE`, `C14=Total`.

### T0-06 — Fórmulas solo BAJO la fila 19

| Familia | Rango | Celdas | Naturaleza |
|---|---|---|---|
| Consumidoras del CONSOLIDADO | `D23:O28` | 72 | `<f>` con texto |
| Auto-chequeos (incluye `L32:L36` y `O32:O36`) | `D32:O36` | 60 | `<f>` con texto |
| **Total** | — | **132** | — |

`Min(fila de <f>) = 23` y el conjunto de fórmulas es **exactamente** `D23:O28 ∪ D32:O36` ⇒ el pegado
`A3:O19` no toca ninguna `<f>` (diff antes/después = ∅, assert H2). Cadena protegida verificada:
`D23 = D9-'CONSOLIDADO_TOTAL RECAUDO'!D104` … `D28 = D14-'CONSOLIDADO_TOTAL RECAUDO'!D109`;
`L32 = INTERVENTORIA!M26+L23=L9` … (ver Plan 39/40).

---

## 3. T0c — Plantilla `DetValiRetri2026092` (`sheet37.xml`)

### T0-07 — Top-block values-only

- **0 `<f>` en `A3:O14`**; los 72 slots `D9:O14` son literales `<v>0</v>`.
- `K3=2026092` (columna del código distinta a DetRetri), `J8=AJUSTE A LA DECENA`, `C14=Total`.
- `Min(fila de <f>) = 16` ⇒ todo el bloque de fórmulas vive bajo el rango del pegado.

### T0-08 — `D16:O20`: consumidoras intactas y fuera del pegado

Forma exacta por ASE (fila 16↔9, 17↔10, 18↔11, 19↔12, 20↔13):

| Col. origen | `<f>` esperada (fila `9+i`, CONSOLIDADO `104+i`) |
|---|---|
| `D..I` | `X{9+i}-'CONSOLIDADO_TOTAL RECAUDO'!{U..Z}{104+i}` |
| `J` | `J{9+i}-'CONSOLIDADO_TOTAL RECAUDO'!AA{104+i}-'CONSOLIDADO_TOTAL RECAUDO'!AC{104+i}` |
| `K` | `K{9+i}-'CONSOLIDADO_TOTAL RECAUDO'!AB{104+i}` |
| `L` / `M` / `N` / `O` | `…!AE{104+i}` / `…!AD{104+i}` / `…!AF{104+i}` / `…!AG{104+i}` |

72 celdas con texto byte-idéntico congelado (contrato del H2 para "antes/después"). La fila `21`
(`Total`) repite la forma contra `U109:AG109` y también queda fuera del pegado. **Ninguna `<f>` de la
hoja cae en filas ≤ 14.**

### T0-09 — Veredicto `J9:J14` (AJUSTE A LA DECENA): entra al comparador

| Celda | R10 (texto `-$`/`$`) | Manual Administrativo (literal) | ≠ 0 |
|---|---|---|---|
| `J9` | `-$46.493` | `-46493` | ✔ |
| `J10` | `$4.616` | `4616` | ✔ |
| `J11` | `$46.663` | `46663` | ✔ |
| `J12` | `-$64.129` | `-64129` | ✔ |
| `J13` | `-$16.688` | `-16688` | ✔ |
| `J14` | `-$76.031` | `-76031` | ✔ |

`J9:J14` está **dentro** de `A3:O14` y R10 ≡ Manual al peso ⇒ **la exclusión `DetValiRetri!J` (SALE,
Plan 29 T0d) se retira** (D-J9 / T5.1). Nota: `Σ J9:J13 = -76.031 = J14`, serie internamente
consistente.

---

## 4. T0d — CONSOLIDADO `sheet18.xml`, filas 141–147

### T0-10 — Geometría

| Fila | Contenido |
|---|---|
| 141 | headers: `B=ASE`, `C=Concesionarios`, `D=Recaudo REMUNERADO`, `E=RECAUDO RECIBIDO`, `F..J` R2/R4/R5/REVERSAR/PERIODOS ANTERIORES |
| 142–146 | un ASE por fila: `B` = `1..5` (literal numérico), `C` = rótulo del ASE |
| 147 | `C147 = TOTAL`; `D147/E147/F147/H147 = SUM(…142:…146)` |

### T0-11 — Forma actual de `E142:E146` (el bloque que N2 debe gobernar)

```
E142 = 'Recaudo EAAB Reciprocidad'!F21+'Recaudo EAAB + Ciud Limp'!F21+'Recaudo Directa Occidente'!F21+'Recaudo ENEL'!F21+'Recaudo ENERBIT'!F21
E143 = … !F22 ×5        E144 = … !F23 ×5        E145 = … !F24 ×5        E146 = … !F25 ×5
```

5 términos, **columna `F` fija**, filas `21..25` (E142→21 … E146→25). Resto de la fila (medido):

| Celda | Forma |
|---|---|
| `D{k}` | `D{104+i}+N{104+i}+D{125+i}+E{125+i}` |
| `F{k}` | `D{28+i}+N{28+i}` (R2) |
| `G{k}` | `D{66+i}+N{66+i}` (R4) |
| `H{k}` | `D{85+i}+N{85+i}` (R5) |
| `L{k}` | `D-E-F-G-J-H-I` — `L142:L144` con texto; **`L145` es master shared** (`ref=L145:L147`, `si=68`) y `L146/L147` seguidoras sin texto |
| `M{k}` | `M142` master shared `INT(ABS(L142))=0` (`ref=M142:M147`, `si=67`); `M143:M147` seguidoras `t="b"` |

> **Riesgo vivo para H4:** `L145`/`M142` son **masters de shared-formula**. El remapeo de H4 toca
> SOLO `E142:E146` (`<f>` con texto propio, no shared) ⇒ no hay riesgo de huérfanas, pero el gate de
> H4 debe verificar `si`/`ref` intactos (doctrina Plan 33).

### T0-12 — Desalineación Q1 (N2) demostrada estáticamente

- El bloque lee **`F{21:25}`** de las 5 hojas `'Recaudo *'` (0 referencias a `D`).
- La columna del recaudo la decide el **dominio**: `PreparadorBasePipeline.ColumnaRecaudoDe(periodo)`
  → `ColumnaRecaudoPorQuincena` = **{1→`D`, 2→`F`}**; `ValidadorBasePeriodo` valida con esa misma
  fuente (`ValidadorBasePeriodo.cs:147`).
- `Periodo.Parse("2026091").NumeroQuincena = 1` ⇒ el writer escribe en **`D{21:25}`** mientras
  `E142` sigue leyendo `F21` ⇒ **bloque en 0 en Q1**. En Q2 (`F`) la columna coincide.

**Declaración de alcance:** H0 **no ejecuta corridas de producción**. La demostración **dinámica**
(una corrida Q1 real con `E142:E146` en 0 antes del fix y con `D{21:25}` después) queda para el
**Gate G.3** del plan; aquí la desalineación se demuestra por texto-`<f>` medido + columna de
dominio, que es evidencia suficiente y no depende de cachés `<v>` (G-WORKBOOK).

---

## 5. Cierre de tareas H0

| Task | AC del §3 | Estado |
|---|---|---|
| T0.1 (T0-01…T0-04: 2 hojas, encabezado, texto monetario, layout) | 4/4 verde | **4/4 verde** |
| T0.2 (T0-05…T0-09: 0 `<f>` en el pegado, consumidoras intactas, J9 con dato) | 5/5 verde | **5/5 verde** |
| T0.3 (T0-10…T0-12: geometría, forma con `F`, desalineación Q1) | 3/3 verde | **3/3 verde** |
| T0.4 (este doc: matriz R10-vs-plantilla + forma `E142:E146` + veredicto J9 + cita de los 12 asserts) | doc existe | **este documento** |

**Comando y resultado real:**

```
dotnet build Remuneracion.WinForms/Remuneracion.WinForms.slnx     → Build succeeded. 0 Warning(s), 0 Error(s)
dotnet test Remuneracion.IntegrationTests/Remuneracion.IntegrationTests.csproj \
  --filter "FullyQualifiedName~PegadoR10T0Plan62Tests"
                                                                  → Passed! Failed: 0, Passed: 12, Skipped: 0, Total: 12
```

**Delta de suite:** altas = **12** (el fixture completo). Rojos nuevos = **0**. Producción tocada =
**0 archivos** (`git status` solo muestra el fixture nuevo; los 3 .xlsx conservan su SHA256).

---

## 6. Desviaciones medidas frente al texto del §T0 (requieren ratificación — NO se ajustó el plan)

El §T0 manda re-medir y reportar cuando un número difiere. Estas son las diferencias entre el texto
del plan y el byte; **el fixture congela el byte** y el plan queda pendiente de fe de errata:

| # | §T0 dice | Byte medido | Impacto |
|---|---|---|---|
| D-1 | T0-03: negativo `-$5.093.257.797` | `DetRetri2026091!J10` = `-$5.093.259.797` | **Dígito** (257↔259). La convención (texto negativo `-$`) se confirma; solo cambia el valor citado. El assert se ata al byte. |
| D-2 | T0-02: "literales de `A3:C8`" | Los literales viven en `F3`,`H3`,`F4`,`F5`,`C6`,`D6`,`F6`,`C7`,`D7`,`F7`,`G7`,`I7`,`J7` (+`L3`/`M3` ó `J3`/`K3` según hoja). `A3:B8` y `E3:E8` están vacías | Solo `C6`/`C7` caen en `A3:C8`: el rango del plan **sub-describe** el encabezado. Los HECHOS (razón social, NIT, código, fecha, hora, rango) se confirman todos. |
| D-3 | T0-10: `C147 = Total` ("comparte s-string con `Total`") | `C147 = TOTAL` (mayúsculas), s-string **distinto** del `Total` de `DetRetri!C14` | Cosmético de rótulo; la geometría (fila 147 = cierre, `SUM` en `D/E/F/H`) se confirma. |
| D-4 | T0-05: rótulo T2 `REPORTE R10- RECAUDO CONSOLIDADO` | `F6 = R10 - Recaudo Consolidado` (idéntico al del R10) | El literal citado no existe en la plantilla vigente. El hecho (el bloque se rotula como R10) se confirma. |
| D-5 | T0-11: `M{k} = INT(ABS(L))=0` para las 5 filas | Solo `M142` trae el TEXTO de la `<f>`; `M143:M147` son seguidoras de la **misma shared** (`si=67`, `ref=M142:M147`) con texto vacío y celda `t="b"` | No cambia nada del remapeo (`E142:E146`); sí obliga a que el gate de H4 preserve `si`/`ref`. |

Ninguna desviación altera el diseño (M-A) ni la extensión E1-4: son precisiones de cita. **La puerta
H0→H1 queda condicionada a la ratificación de D-1…D-5** por el Ingeniero (regla del §T0: no se
"ajusta" el plan en silencio).

### Hallazgo adicional (para T1.2, no bloquea H0)

`ExcelDataReaderDetRetriR10Reader.ParsearNumeroMonetario` documenta "`'.'` miles, `','` decimal" pero
implementa `new string(texto.Where(char.IsDigit).ToArray())` ⇒ un monto con decimales
(`"$1.234,56"`) se leería `123456`. En 2026091 **no se dispara** (los 144 montos son enteros, T0-03),
pero al extraer `ParseoMonetarioR10` (fuente única) la implementación debe honrar el decimal o
declarar explícitamente el redondeo a entero. El fixture NO fija el bug: fija la convención.

---

## 7. Consecuencias para H1…H5 (trazabilidad)

| Hecho medido | Consume |
|---|---|
| T0-01 (2 hojas, nombres de dominio) | T1.1/T1.3 (resolución de hoja por `NombresHojaPeriodo`, nunca enumeración) y el fail-fast `ERR-PLANTILLA` si falta una |
| T0-02 (encabezado + sello en el R10) | D-31SELLO (retiro de `SellarProcesoEnDetRetri`), T2.4 conserva solo la LECTURA de `FechaProceso`/`HoraProceso` |
| T0-03 (`t="s"` monetario) | T1.2 (`ParseoMonetarioR10` única fuente) + D-1/D-hallazgo de §6 |
| T0-04 (layout idéntico, `A` vacía) | T2.2 (pegado celda-a-celda con estilos de destino) — el layout del origen ya coincide |
| T0-05/T0-06 (0 `<f>` en `A3:O19`; fórmulas bajo 19) | T2.2 (guard anti-fórmula como red, no como camino) y el diff `<f>` = ∅ |
| T0-07/T0-08 (0 `<f>` en `A3:O14`; `D16:O20` consumidoras) | T2.1 (orden: CONSOLIDADO → pegado → oráculos), T3.1 (gate per-ASE por `D16:D20`) |
| T0-09 (J9:J14 con dato) | T5.1 (retiro de la exclusión `DetValiRetri!J`), D-J9 |
| T0-10/T0-11 (geometría 141–147 y forma `E{k}` con `F`) | T4.1 (`CeldasRecaudoValidacionConsolidado142` + `AplicarVariante…` con `ColumnaRecaudoDe`) |
| T0-12 (desalineación Q1) | T4.2 (Q1 → `D{21:25}` / Q2 → `F{21:25}`) y G.3 (demostración dinámica) |
