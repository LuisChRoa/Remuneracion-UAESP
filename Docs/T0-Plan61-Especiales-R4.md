# T0 — Plan 61: Especiales R4 por encabezado (evidencia H0/F0)

**Fecha:** 2026-10-28 · **Fase:** H0/F0 (read-only, G-H0-READONLY) · **Plan:** `plans/61 - especiales-r4-por-encabezado-y-entrega-con-novedades.md` (APROBADO, §6 ratificado 8/8)
**Fixture:** `Remuneracion.IntegrationTests/EspecialesR4T0Plan61Tests.cs` — **17/17 verde**
**Naturaleza:** medición sobre el disco. Ningún archivo de producción, test existente, plantilla ni insumo se modificó para producir esta evidencia; las fuentes R4 se abren con `ExcelDataReader` solo-lectura sobre **copias** temporales y la plantilla/administrativos/R10 con BCL puro (`ZipArchive` + `System.Xml.Linq`) sobre `FileShare.ReadWrite`.

---

## 0. Veredicto

La evidencia del §0 del plan **se reproduce en el disco**, celda por celda. No hay ningún monto, fila ni celda del §0 que difiera. Hay **cuatro diferencias de citación** (no de fondo) que se declaran en §11: el nombre real de la hoja consolidada, el de las hojas `DetRetri`, un carácter de la huella semántica en el frontmatter del plan y el estado de la suite de entrada.

### 0.1 Números medidos vs §0

| Dato citado en §0 | Citado | Medido en bytes | Veredicto |
|---|---|---|---|
| 2026092 ASE4 · R4 fila de cierre (`A="Total"`, `B` vacío) | — | fila **23** (única candidata) | ✅ |
| 2026092 ASE4 · `R4!D23` | −178.619.709,53 | **−178.619.709,53** | ✅ coincide |
| 2026092 ASE4 · `R4!P23` (`SERVICIO ESPECIALES`) | −49.900,44 | **−49.900,44** (letra física **P**) | ✅ coincide |
| 2026092 ASE4 · cierre del administrativo | `CONSOLIDADO!D69 = 'Reversion Pagos R4'!D316 = D239 − P239` | `D69` fórmula **`='Reversion Pagos R4'!D316`**; `D316` fórmula **`=D239-P239`** = **−178.569.809,09**; `N69` = `'Reversion Pagos R4'!P316` = −49.900,44 | ✅ coincide |
| 2026092 ASE4 · `CONSOLIDADO!D107` | 7.675.362.501,96 | **7.675.362.501,96** (`=D12+D31+D50+D69+D88`) | ✅ coincide |
| 2026092 ASE4 · oráculo R10 | 7.675.362.502 | **7.675.362.502** (`DetRetri2026092!D12`, texto `"$7.675.362.502"`) | ✅ coincide |
| 2026092 ASE4 · calculado por la app | 7.675.312.602 | **7.675.312.601,52** ⇒ redondeo **7.675.312.602** (D107 + P23) | ✅ coincide |
| 2026091 ASE2 · R4 fila de cierre | — | fila **11** (única candidata) | ✅ |
| 2026091 ASE2 · `R4!D11` | −8.739.602,37 | **−8.739.602,37** | ✅ coincide |
| 2026091 ASE2 · `R4!P11` | −26.500,00 | **−26.500,00** (letra física **P**) | ✅ coincide |
| 2026091 ASE2 · `CONSOLIDADO!D67` | −8.713.102,37 | **−8.713.102,37** (`='Reversion Pagos R4'!D160`; `D160` = `=D97-P97`) | ✅ coincide |
| 2026091 ASE2 · `CONSOLIDADO!D105` | 14.147.357.755,39 | **14.147.357.755,39** (`=D10+D29+D48+D67+D86`) | ✅ coincide |
| 2026091 ASE2 · `R10 DetRetri!D10` | 14.147.357.755 | **14.147.357.755** (`DetRetri2026091!D10`, texto `"$14.147.357.755"`) | ✅ coincide |
| 2026091 ASE2 · calculado por la app | 14.147.331.255,39 | **14.147.331.255,39** (D105 + P67) | ✅ coincide |
| 2026091 ASE2 · mensaje del gate Σ-empresas | Σ −8.713.102,37 vs bloque −8.739.602,37 | **Δ = 26.500,00** = `abs(P11)`; la resta por empresa existe (`ResolvedorConciliacionEmpresa:302-303`), el bloque no resta | ✅ coincide |
| 2026092 ASE2 · `Ajuste Recargo Aseo` en la letra P | P | **P4 = "Ajuste Recargo Aseo"**, `Q4 = "Otros"`; **no hay** encabezado de Especiales en el archivo | ✅ coincide |
| Corpus R4 | 30 / 25 / 1 / 2 (cuatro cifras en el brief) | **30** archivos `ReversiónPorComponente*`; **2** con encabezado; **2** con valor ≠ 0 | ✅ resuelve §8-C4 (§6 de este doc) |
| Huella semántica v4 (frontmatter del plan) | `a7476ee2…524d5140b8608c55` | **`a7476ee2…526d5140b8608c55`** | ⚠️ **difiere en 1 carácter** (= constante canónica `HuellaSemanticaPlantilla.V4`). Ver §11-D4 |

**Conclusión:** el T0 confirma el §0 completo. El defecto es exactamente uno: el bloque R4 resta `Total` sin `Especiales`; el oráculo R10 y el administrativo sí restan.

---

## 1. T0.1 — Caso 2026092 · ASE4 (Bogotá Limpia)

Fuente (nombre exacto, verificado en disco):

```
Docs/Prueba Septiembre-2/Insumos/4-Bogota Limpia/ReversiónPorComponente_to_date16092026ddMMyyyy_to_date30092026ddMMyyyy___20261024531443.xlsx
```

Hoja `Sheet1`, dimensión `A1:Q23`, 20 filas con datos. Encabezados de componente en la fila **4**; fila de cierre `A=Total` con `B` vacío en la **23** (una sola candidata).

| Letra | Encabezado (fila 4) | Valor fila 23 |
|---|---|---|
| D | `Total` | **−178.619.709,53** |
| O | `Ajuste Decena` | 27,97 |
| **P** | **`SERVICIO ESPECIALES`** | **−49.900,44** |
| Q | `Otros` | 192,75 |

Derivación esperada por la plantilla: `D23 − P23 = −178.619.709,53 − (−49.900,44) = **−178.569.809,09**` (fila 23 inspeccionada también en la fuente para las filas de empresa: `ENEL` en 10, `OCCIDENTE` en 22).

### Cierre en el administrativo

`Docs/Prueba Septiembre-2/Resultados/Administrativo/Remuneracion 202609-2 Total.xlsx`, hoja **`CONSOLIDADO_TOTAL RECAUDO`** (la forma corta `CONSOLIDADO` del plan **no existe** como nombre de hoja; ver §11-D1) y hoja **`Reversion Pagos R4`**:

| Celda | Fórmula medida | Valor |
|---|---|---|
| `Reversion Pagos R4!D239` (fila `A=Total` de la fuente) | — | −178.619.709,53 |
| `Reversion Pagos R4!P239` | — | −49.900,44 |
| `Reversion Pagos R4!D316` (cierre `A=TOTAL`) | **`=D239-P239`** | −178.569.809,09 |
| `CONSOLIDADO!D69` (ASE4) | **`='Reversion Pagos R4'!D316`** | −178.569.809,09 |
| `CONSOLIDADO!N69` | `='Reversion Pagos R4'!P316` | −49.900,44 |
| `CONSOLIDADO!D107` (total del ASE4) | **`=D12+D31+D50+D69+D88`** | **7.675.362.501,96** |

Cabecera del bloque R4 en la plantilla: `P2 = "SERVICIO ESPECIALES"`, **`Q2 = "Ajuste Recargo Aseo"`**, `R2 = "Otros"`. La plantilla espera **ambos** operandos y ubica el `Ajuste Recargo Aseo` en una letra distinta a la de la fuente (§4).

### R10

`Docs/Prueba Septiembre-2/Insumos/R10_Remuneracion_2026092.xlsx`, hoja **`DetRetri2026092`** (no existe una hoja llamada `DetRetri`; ver §11-D2):

| Celda | Valor |
|---|---|
| `B12` | `4` (ASE4) |
| `D12` (`RECAUDO TOTAL`) | `"$7.675.362.502"` (texto con formato de moneda) |

**Δ del caso:** `|D107(con Especiales) − DetRetri| = |7.675.362.501,96 − 7.675.362.502| = 0,04` (redondeo). El número que hoy calcula la app, `7.675.362.501,96 + P23 = 7.675.312.601,52`, se separa del R10 en **49.900,44**, muy por encima de la tolerancia ±0,5. En 2026092 el gate Σ-empresas **no corre** (residuo `R-SIGMA-Q2`): el único detector posible era el R10.

---

## 2. T0.2 — Caso 2026091 · ASE2 (Lime)

Fuente (nombre exacto, verificado):

```
Docs/Prueba Septiembre 1/Insumos/2-Lime/ReversiónPorComponente_to_date01092026ddMMyyyy_to_date15092026ddMMyyyy___202691615358215.xlsx
```

Hoja `Sheet1`, dimensión `A1:P11`, 8 filas con datos. Encabezados en la fila **4**; cierre en la **11**.

| Letra | Encabezado | Valor fila 11 |
|---|---|---|
| D | `Total` | **−8.739.602,37** |
| **P** | **`SERVICIO ESPECIALES`** | **−26.500,00** |

`D11 − P11 = **−8.713.102,37**`.

### Cierre en el administrativo

`Docs/Prueba Septiembre 1/Resultados/Administrativo/Remuneracion 202609-1 Total.xlsx`:

| Celda | Fórmula medida | Valor |
|---|---|---|
| `Reversion Pagos R4!D97` (fila `A=Total` de la fuente) | — | −8.739.602,37 |
| `Reversion Pagos R4!P97` | — | −26.500,00 |
| `Reversion Pagos R4!D160` (cierre) | **`=D97-P97`** | −8.713.102,37 |
| `Reversion Pagos R4!P160` | — | −26.500,00 |
| `CONSOLIDADO!D67` (ASE2) | **`='Reversion Pagos R4'!D160`** | −8.713.102,37 |
| `CONSOLIDADO!N67` | `='Reversion Pagos R4'!P160` | −26.500,00 |
| `CONSOLIDADO!D105` (total del ASE2) | **`=D10+D29+D48+D67+D86`** | **14.147.357.755,39** |

### R10

`Docs/Prueba Septiembre 1/Insumos/R10_Remuneracion_2026091.xlsx`, hoja `DetRetri2026091`: `B10 = 2`, `D10 = "$14.147.357.755"` ⇒ **14.147.357.755**.

**Δ del caso:** el número de la app (`D105 + P67 = 14.147.331.255,39`) se separa del R10 en **26.500,00**. Aquí el detector fue el gate Σ-empresas de Q1 (`ValidadorBasico:884-886`): el lado por empresa resta Especiales y el lado del bloque no.

---

## 3. T0.3 — Encabezado exacto, normalización y columna intermitente

- El encabezado medido es exactamente **`SERVICIO ESPECIALES`** (sin espacios de borde, sin variante de caja) en las dos fuentes con la columna.
- Normalización aplicable: `Trim` + `OrdinalIgnoreCase` sobre la tabla **única** `ResolvedorConciliacionEmpresa.EncabezadosEspeciales = ["SERVICIO ESPECIALES", "Especiales"]` (una sola declaración en producción; el fixture lo comprueba).
- `"Ajuste Recargo Aseo"` y `"Otros"` **no** son alias de Especiales.

### Presencia/ausencia por ASE y período (10 archivos, todos medidos)

| Período | ASE1 | ASE2 | ASE3 | ASE4 | ASE5 |
|---|---|---|---|---|---|
| 2026091 | sin columna | **P = `SERVICIO ESPECIALES` (−26.500,00)** | sin columna | sin columna | sin columna |
| 2026092 | sin columna | **P = `Ajuste Recargo Aseo`** (`Q = Otros`; ≠ Especiales ⇒ 0) | sin columna | **P = `SERVICIO ESPECIALES` (−49.900,44)** | **P = `Ajuste Recargo Aseo`** (sin Q) |

La letra **no** es contrato: en 2026092 el ASE2 y el ASE5 usan la letra P para un concepto distinto. Un lector por índice habría inyectado **−10,16** (valor de `P22` en el ASE2-2026092) como si fuera Especiales; el lector por encabezado devuelve **0** y lo declara ausente.

En la plantilla (hoja `Reversion Pagos R4`), `Ajuste Recargo Aseo` vive en **Q**: la fuente lo trae corrido una letra a la izquierda. Sin resolución por encabezado, cualquier mapeo por letra se rompe el mes siguiente.

---

## 4. T0.4 — Corpus con conteo cerrado y gates respecto de la escritura

### 4.1 Corpus (§8-C4 resuelto)

Barrido exhaustivo de `Docs/**/*.xlsx` (287 libros en total), por prefijo normalizado del *stem*:

| Familia | Archivos | Con encabezado de Especiales | Con valor ≠ 0 |
|---|---|---|---|
| `ReversiónPorComponente*` (R4 del pipeline) | **30** | **2** | **2** |
| `R4-BalanceSubsidioyContribuciones*` (familia distinta: Balance) | 31 | no aplica | no aplica |

Distribución del corpus R4: **6 períodos × 5 ASE** (Julio-1, Julio-2, Agosto-2, Septiembre 1, Septiembre-2, Octubre-1). Los dos únicos con la columna son:

1. `Docs/Prueba Septiembre 1/Insumos/2-Lime/ReversiónPorComponente_to_date01092026ddMMyyyy_to_date15092026ddMMyyyy___202691615358215.xlsx`
2. `Docs/Prueba Septiembre-2/Insumos/4-Bogota Limpia/ReversiónPorComponente_to_date16092026ddMMyyyy_to_date30092026ddMMyyyy___20261024531443.xlsx`

**Resolución de C4:** el brief mezcló tres universos. «30» = total de archivos R4 del pipeline en `Docs/`; «2» = los que traen la columna con valor; «25»/«1 de 25» **no se reproduce** con ninguna definición razonable de corpus (ni por archivos, ni por períodos, ni excluyendo una quincena). La cifra vinculante es **30 archivos / 2 con encabezado / 2 con valor ≠ 0**.

### 4.2 Inventario de gates R4 y posición respecto del writer

Archivo `Remuneracion.Core/Services/ProcesadorPeriodo.cs` (líneas medidas el 2026-10-28):

| # | Gate | Ubicación | Posición vs. escritura | Efecto hoy | Clase §2-D-H3 (H3) |
|---|---|---|---|---|---|
| 1 | Preflight de insumos | `ValidadorInsumosPeriodo` (antes del cálculo) | ANTES | fallo duro | duro |
| 2 | `_validador.Validar(resultado, leafs)` — multi-ASE; incluye Σ visibles-empresa vs bloque R1/R2/R4 (Q1) | `ProcesadorPeriodo.cs:374`; aritmética en `ValidadorBasico.cs:846-888` (mensaje R4 en `:884-886`; recorte a Q1 en `:866`, residuo `R-SIGMA-Q2` en `:853-865`) | **ANTES** (línea 374 < 402) | throw `Validacion`, sin Excel | **novedad** |
| 3 | `ValidarContraResultadoMultiAse` (leaf vs agregado; R4 en `:42-44` y `:200-203`) | llamada **dentro** de `OpenXmlPlantillaWriter.GenerarWorkbook` en `OpenXmlPlantillaWriter.cs:240` | ANTES de escribir celdas, pero **dentro** de la llamada de escritura | throw `Validacion`, sin archivo | **novedad** |
| 4 | `ValidarSigmaEmpresas` (Σ por empresa vs bloque) | `OpenXmlPlantillaWriter.cs:243-246` → `WorkbookLeafCoherence.cs:95-123` | ídem (dentro del writer) | throw `Validacion`, sin archivo | **novedad** |
| 5 | `ValidarDetRetriContraR10` (ambas quincenas) | `ProcesadorPeriodo.cs:389` | **ANTES** (389 < 402) | throw `Validacion`, sin Excel | **novedad** |
| 6 | Escritura a `rutaSinValidar` | `ProcesadorPeriodo.cs:402` | — | escribe el workbook | — |
| 7 | Validación cruzada 2.7 no declarada | `ProcesadorPeriodo.cs:453-459` | **DESPUÉS** (453 > 402) | throw `Validacion`, el `(sin validar)` ya está en disco | **novedad** (D-NOVEDAD-27) |
| 8 | Promoción `rutaSinValidar → RutaSalida` | `ProcesadorPeriodo.cs:539-545` (Move en `:544`) | **DESPUÉS** | renombra al nombre limpio | — |

**Hallazgo para H3 (no bloquea H0):** los gates 3 y 4 **no** son pre-gates del procesador: viven al inicio de `GenerarWorkbook` y abortan **antes de crear el archivo**. Reordenar solo `ProcesadorPeriodo` (H3.2) no basta: si no se sacan del writer (o se degradan allí), la corrida con coherencia rota seguirá sin dejar ni siquiera el `(sin validar)`. Es la razón por la que `ProcesadorPeriodoTests:398-401` y `VerificadorRecaudo:409-436` afirman `!File.Exists(salida)` y por la que H3.3 debe preservar la intención **en esos dos fixtures** y no en el procesador (ver §8).

---

## 5. T0.5 — Supuestos de ausencia, fixtures con valores inventados y tautologías

### 5.1 Supuestos de ausencia en producción (archivo:línea, clasificados)

| Archivo:línea | Texto | Clasificación |
|---|---|---|
| `ExcelDataReaderWorkbookLeafInputReader.cs:1389-1390` | `<c>P = 0</c> (sin análogo en la fuente, patrón histórico)` | **latente = el defecto** (se corrige en H1) |
| `ExcelDataReaderWorkbookLeafInputReader.cs:1400` | `var p9 = 0m;` | **latente = el defecto** |
| `ResolvedorConciliacionEmpresa.cs:236` | «el componente Especiales (P) no tiene análogo en la fuente → ausencia = cero del template» | **latente/obsoleto**: el código de `:302-303` **sí** resta `Especiales`; el comentario miente |
| `ResolvedorConciliacionEmpresa.cs:301` | «Visible R4 por empresa = Total (P = 0 en la fuente)» | **obsoleto**: contradice a `:303` (`o.Total - o.Especiales`) |
| `WorkbookLeafCellMapQ2.cs:209-210` | «P = 0 en el golden (sin análogo en la fuente, patrón Q1)» | **latente** (mapa de celdas, no de valores) |
| `ExcelDataReaderWorkbookLeafInputReader.cs:896-899` | `if (indiceEspeciales < 0) valores["SERVICIO ESPECIALES"] = 0m;` (R1) | **legítimo**: columna ausente ⇒ 0, con regla explícita |
| `ExcelDataReaderWorkbookLeafInputReader.cs:1494-1498` | `// Columna 'Especiales' ausente en la fuente (julio) → 0 explícito por fila` | **legítimo** |
| `ExcelDataReaderWorkbookLeafInputReader.cs:1291-1293` | `F30/F10/L10 = 0m` (R1) | **legítimo** (valores que el gate usa sólo en Q1; declarado en HU) |
| `ResolutorValorValidacion.cs:191` y `:226` | `// celda ausente = 0 (semántica de Excel)` | **legítimo** (`:226` incluso advierte que no debe usarse para ocultar un ancla borrada) |
| `CalculoRemuneracion.cs:41` | `AjustesSfT = 0m` en sobrecarga sin ajustes | **legítimo** (Q1 por contrato) |

Lista **cerrada**: sólo **tres** archivos de producción contienen las frases «sin análogo», «patrón histórico» o «no tiene análogo» (`ExcelDataReaderWorkbookLeafInputReader.cs`, `ResolvedorConciliacionEmpresa.cs`, `WorkbookLeafCellMapQ2.cs`). El fixture falla si aparece una cuarta.

### 5.2 Fixtures con valores de fuente inventados (§8-C3 resuelto)

| Archivo:línea | Literal | Nota |
|---|---|---|
| `AjustesSfTTests.cs:312` | `R4 = new WorkbookLeafInputsR4 { D9 = -(300m + i), P9 = 0m }` | citado por el brief |
| `ConciliacionEmpresaSigmaPlan55Tests.cs:319` | `R4 = new WorkbookLeafInputsR4 { D9 = visibleR4, P9 = 0m }` | citado por el brief |
| `ConciliacionEmpresaTests.cs:152` | `D9 = r4, P9 = 0m,` | citado por el brief |
| **`ValidadorBasicoTests.cs:90`** | `P9 = 0m` | **omitido por el brief (C3)** |
| **`ValidadorMultiAseTests.cs:174`** | `P9 = 0m,` | **omitido por el brief (C3)** |

`E26 = 0m` (misma doctrina, misma deuda) en **4** fixtures: `AjustesSfTTests.cs:311`, `ConciliacionEmpresaSigmaPlan55Tests.cs:318`, `ConciliacionEmpresaTests.cs:147`, `ValidadorMultiAseTests.cs:167`. `K15 = 0m` acompaña en tres de ellos.

### 5.3 Tautologías

| Archivo:línea | Aserto | Por qué es tautológico |
|---|---|---|
| `DetalleR2R4Tests.cs:134` | `Assert.Equal(leaf.R4.D9 - leaf.R4.P9, leaf.R4.TotalReversionEsperada)` | es la definición de `TotalReversionEsperada`, no un valor medido |
| `DetalleR2R4Tests.cs:161` | idéntico | idéntico |
| `DetalleR2R4Tests.cs:133` y `:160` | `Assert.Equal(leaf.R2.E15 + leaf.R2.E26 - leaf.R2.K15, leaf.R2.TotalOportunoEsperado)` | misma forma (R2); el plan no la cita, se declara aquí |

---

## 6. T0.4/T0.6 (segunda mitad) — Tests que asumen ausencia de salida, nombre limpio o nombre del reporte

| Fixture | Qué asume | Estado real medido | Lectura para H3.3 |
|---|---|---|---|
| `ProcesadorPeriodoTests.cs:372-402` (`WriterMultiAse_GateDeCoherenciaRotoEnAse3_NoGeneraArchivo`) | `Assert.False(File.Exists(salida))` tras coherencia rota | **NO** pasa por `ProcesadorPeriodo`: llama a `new OpenXmlPlantillaWriter().GenerarWorkbook(...)` con un `salida` limpio. No existe en él la ruta `(sin validar)` | La receta de H3.3 («pasa a exigir no existe nombre LIMPIO + existe `(sin validar)`») **no aplica tal cual**: o el gate sigue dentro del writer (fallo duro estructural, test intacto) o se traslada al procesador y el test se re-ancla pasando por `ProcesadorPeriodo` |
| `VerificadorRecaudo/Program.cs:409-436` (T19) | `noCertificado = !File.Exists(salida)` | igual: `writer.GenerarWorkbook(...)` directo | Ídem |
| `SalidaSinValidarO4Tests.cs:20-51` | no existe nombre definitivo **y** existe el marcado | **ya contempla** el `(sin validar)` (`Assert.True(File.Exists(marcada))`, `NombresSalida.RutaSinValidar`) | intacto en espíritu |
| `AnclasDerivadasH2Plan57Tests.cs:462` | `File.Exists(salida) ? salida : RutaSinValidar(salida)` | **ya contempla** el marcado | intacto en espíritu |
| `DivergenciaDeclaradaCliTests.cs:92` | nombre exacto `Remuneración 202607-2 Total - Divergencias declaradas.txt` | depende del literal | conservar compatibilidad (D-TXT ratificada) |
| `ReporteDivergenciasDeclaradasTests.cs:83-84` | nombre contiene `Divergencias declaradas` | depende del literal | ídem |

**Ningún test asume `exit ≤ 6`.** Barrido con regex `<=\s*6\b|<\s*7\b` sobre toda la fuente de tests: **0 coincidencias**. El contrato se afirma con constantes exactas (`CatalogoErroresTests:115-120`, `CodigosSalidaCliTests`, `MicroFixesHu17Tests:199` usa `Assert.Equal(4, proceso.ExitCode)`). `CodigosSalida` declara hoy **7** constantes (`Ok = 0` … `DivergenciaDeclaradaReconocida = 6`); agregar `7` no rompe ninguna aserción existente.

---

## 7. T0.7 — Huella semántica v4

| Dato | Valor |
|---|---|
| Archivo | `Docs/Plantilla estandarizada/v4/Formato Remuneracion 2026xx-x Total v4.xlsx` |
| Huella semántica medida (`HuellaSemanticaPlantilla.Calcular`, 42.174 entradas normalizadas) | `a7476ee2e80dfa13c9e39430d1ea4d77f0c52e3ca0a3d11a**526**d5140b8608c55` |
| Constante canónica `HuellaSemanticaPlantilla.V4` | idéntica |
| Frontmatter del plan | `…a3d11a**524**d5140b8608c55` (1 carácter distinto) |
| SHA-256 del archivo (informativo, no gate) | `5fedfd316980276945df2af29784effc5fe178c5cf1aa16f3f3c5d5f6f0a4e38e5` |

El gate G-HUELLA está **intacto**: el valor medido coincide con la constante canónica que consumen los tests vigentes. La diferencia es de **citación en el plan** (§11-D4).

### Δ de suite

Ver §9 (medición base/después).

---

## 8. Disposición de las contradicciones §8-C1..C6

| ID | Estado | Evidencia y disposición |
|---|---|---|
| **C1** | **RESUELTA** | El archivo real es `ConciliacionEmpresaCierreBloqueT0Tests.cs` (sin `Plan55`), con el supuesto «`MapearR4 (D − P, con P = 0)`» en `:682` y gemelos en `:395` y `:821`. El fondo del brief se confirma; H1.3/H5.3 deben citar el nombre real. |
| **C2** | **RESUELTA** | Líneas exactas medidas: mensaje R4 del gate Σ-empresas en `ValidadorBasico.cs:884-886`; recorte a Q1 en `:866`; residuo `R-SIGMA-Q2` documentado en `:853-865`. Sin contradicción de fondo. |
| **C3** | **RESUELTA** | Hay **5** fixtures con `P9 = 0m` (3 citados + `ValidadorBasicoTests:90` + `ValidadorMultiAseTests:174`) y **4** con `E26 = 0m`. H2.3 los cubre a todos. |
| **C4** | **RESUELTA** | Corpus cerrado: **30** archivos R4 del pipeline, **2** con encabezado, **2** con valor ≠ 0 (§4.1). La cifra «1 de 25» no se reproduce. |
| **C5** | **RESUELTA CON CORRECCIÓN DE RECETA** | Ver §6: los dos fixtures que afirman ausencia de archivo (`ProcesadorPeriodoTests:401`, `VerificadorRecaudo:430`) ejercitan el **writer**, no el procesador. `SalidaSinValidarO4Tests` y `AnclasDerivadasH2:462` ya contemplan el marcado. La re-ancla de H3.3 debe decir explícitamente cuál de los dos caminos se elige (gate dentro del writer = test intacto; gate movido al procesador = test re-anclado pasando por `ProcesadorPeriodo`). |
| **C6** | **RESUELTA** | `ValidarDetRetriContraR10` es fail-fast **duro y pre-escritura** (`ProcesadorPeriodo.cs:389` antes de `:402`): no «explota» el corpus, aborta sin Excel. Confirma que D-H3 debe moverlo a novedad post-escritura. |

---

## 9. Suite de entrada y Δ de H0

La primera referencia (trx que cita el gate de entrada) quedó **desactualizada respecto del árbol**: precede a tres commits del mismo día. Se reportan las dos y se explican.

| Referencia | Trx | Total | Aprobados | Fallidos | Omitidos |
|---|---|---|---|---|---|
| Trx base citado por el gate de entrada (2026-10-27 08:16) | `p59-h5-suite.trx` | 1618 | 1584 | 34 | 0 |
| Corrida de H0 (árbol actual + fixture T0, 2026-10-28 20:43) | `p61-h0-suite.trx` | **1674** | **1649** | **25** | **0** |

Los commits `c9da3d5` (16:25), `c8ab07d` (17:23) y `472e26f` (19:47) del 2026-10-27 son posteriores al trx base: agregaron **39** tests y dejaron verdes **9** de aquellos 34 rojos. La aritmética cierra: `1618 + 39 + 17 = 1674`.

**Línea base efectiva de H0** = árbol actual sin el fixture ⇒ **1657** tests (1674 − 17), con los **25** rojos de §9.1.

**Δ de H0:**

- **+17** tests nuevos (los 17 del fixture T0), todos verdes.
- **0 rojos nuevos**: el conjunto de rojos de hoy es un **subconjunto estricto** del trx base (`hoy − base = ∅`); ningún rojo puede atribuirse a H0.
- **9 bajas declaradas** (rojos del trx base que hoy pasan, por los commits posteriores, no por H0): `MicroFixesHu17Tests.S4_Hu16_InterventoriaLecturaPorAse_Debug_HitoInformation`, `ProcesadorPeriodoTests.Ejecutar_Periodo2026071_ConLectorOracle_VeredictosValidacionesNoVacios`, `CodigosSalidaCliTests.Ejecucion_5Ase_Q1_Devuelve0_ConArchivoYResultadoOk`, `ObservabilidadTests.Niveles_HitosInformation_DetalleDebug`, `ParidadCliTests.Paridad_5Ase_Q2_CliVsDirecto_IncluyeAjustesYDetRetri`, `ObservabilidadTests.BloqueValidaciones_PropiedadValidacion_NombreReal_ConAseId`, `ObservabilidadTests.EjecucionPeriodo_RunIdCorrelacionaInicioAFin_YAseId`, `MicroFixesHu17Tests.L1_Hu12_BloqueVacio_NoSeLimpia_OutputIgualTemplate`, `MicroFixesHu17Tests.W22_Hu14_Validaciones_CulturaEsCo_LineaYEventoCoherentes`.
- **0** tests borrados, **0** `Skip`, **0** tolerancias ampliadas.

El único cambio de H0 es un archivo de test nuevo (diff de producción vacío, ningún test existente editado); la comparación de conjuntos lo confirma de forma empírica.

### 9.1 Lista cerrada de rojos del árbol (25, pre-existentes a H0)

- `AdopcionPlantillaV2T0Plan48Tests.T0P48_P3_R1_Manual_492_Formulas_DosPatrones`
- `AdopcionPlantillaV2T0Plan48Tests.T0P48_P6_TablaParidad_40Hojas_ConteosReales`
- `BalanceScPermutaTests.JulioAse5Bce_ManualDivergeDeLaFuente_ConstanteDocumentada`
- `BalanceScPermutaTests.JulioQ2_PermutaBce_SubsidioAD_ContribucionAE_VsFuente_E2E`
- `BalanceScTests.J9Consolidado_NoEsOraculoDeValoresBce`
- `BaseConsistente2026082Tests.BaseAgosto2026082_EsInternamenteConsistenteConElManual`
- `BaseConsistente2026082Tests.BaseAgosto2026082_MetadatosCanonicosEsperados`
- `BasePeriodoPreflightTests.S5_OficialDirecto_EndToEnd_L9L13PorQuincena(2026071, …)`
- `ComparadorSalidaVsManualTests.Salida5AseQ2_vs_Manual_CeroDivergenciasFueraDeBrechasDeclaradas(agosto)`
- `ComparadorSalidaVsManualTests.Salida5AseQ2_vs_Manual_CeroDivergenciasFueraDeBrechasDeclaradas(julio)`
- `DetRetriDesgloseTrazableTests.JulioQ2_DesgloseTrazable_DetRetri_JyL_Y_DetVali_IyO`
- `GoldenBalanceScTests.ProcesadorPeriodo_CapaA2_4_BceContraGolden`
- `GoldenConciliacionEmpresaTests.ProcesadorPeriodo_CapaA2_2_EmpresasYRecaudosContraGolden`
- `GoldenInterventoriaTests.CapaA_InterventoriaYLEspeciales_AmbosCanonicos`
- `GoldenLeafPathTests.Procesador_CapaA_LeafVisiblesFormulasYHashSinUsarCacheDeSalida`
- `GoldenMultiAseTests.ProcesadorPeriodo_CapaAX5_LeafVisiblesFormulasHashYSumas`
- `GoldenReporteBancoTests.ProcesadorPeriodo_CapaA2_3_BloquesYConsolidadoContraGolden`
- `GoldenResumenMesYR10Tests.GoldenQ1_S3_ResumenMes_DE_EsCopia1a1DeLosRecaudos`
- `InterventoriaTests.D3b_LEspecialesMenores_TodaLNumericaFueraDeMapas_EsCero_EnAmbosCanonicos`
- `ParidadCliTests.Paridad_5Ase_Q1_CliVsDirecto_ValoresYSnapshotsIguales`
- `Regresion2026082Tests.Ejecutar_Periodo2026082_Modo5Ase_EndToEnd_DetRetri5De5VsR10`
- `SharedFormulaR1GateTests.S1_RegenFrescaJulio_MaestrosSharedPreservados`
- `ValidadorInteriorR1GateTests.S1_Agosto2026082_FormaInteriorVerde`
- `ValidadorInteriorR1GateTests.S1_Julio2026072_FormaInteriorVerde`
- `WorkbookLeafInputsTests.GenerarWorkbook_EscribeSoloLeafYPreservaFormulas`

> Lectura para el Ingeniero: G-SUITE exige «suite verde», pero el árbol trae **25 rojos** que H0 **no** introduce ni puede arreglar (son de golden/salidas/observabilidad previas). El set-diff de H0 es exactamente cero, así que la puerta se cumple en su forma de *gate* («rojo nuevo fuera de lista ⇒ STOP»); el verde total es deuda previa, ajena a este plan (§11-D3).

---

## 10. Puerta H0 → H1

| Criterio | Estado | Evidencia |
|---|---|---|
| **S1.1** D23/P23 (2026092 ASE4) y D11/P11 (2026091 ASE2) re-medidos: fila, letra, encabezado, incluido `Ajuste Recargo Aseo` en P del ASE2-2026092 | **CERRADO** | §1, §2, §3 (`T0_1_*`, `T0_2_*`, `T0_3_*`) |
| **S1.2** forma real de `D69 = D239 − P239` (2026092) y `D67/D105` + `DetRetri!D10` (2026091) | **CERRADO** | §1, §2 (fórmulas y valores textuales) |
| **S1.3** encabezado exacto + normalización + presencia/ausencia por ASE | **CERRADO** | §3 (10/10 archivos) |
| **S1.4** barrido de corpus con conteo cerrado + inventario de gates con posición vs writer | **CERRADO** | §4 (C4 resuelta) |
| **S1.5** inventario de supuestos de ausencia + fixtures inventados (incl. omitidos) + tautologías + tests que asumen ausencia de salida/nombre | **CERRADO** | §5, §6 (C3 y C5 resueltas) |
| **S1.6** lista cerrada de Δ de suite + huella v4 antes/después | **CERRADO CON OBSERVACIÓN** | §7 y §9: huella intacta (constante canónica); Δ = +17 nuevos verdes, **0 rojos nuevos**, 9 bajas explicadas por commits posteriores al trx base; el árbol conserva 25 rojos previos |
| Fixture verde | **CERRADO** | 17/17, sin `Skip` |
| Build 0 warnings | **CERRADO** | `dotnet build Remuneracion.WinForms/Remuneracion.WinForms.slnx` → 0 warnings / 0 errores |
| G-H0-READONLY (diff de producción vacío) | **CERRADO** | `git status --porcelain`: sólo el fixture y este doc (más los `TestResults/*.trx` de la corrida) |

**Puerta H0→H1: ABIERTA**, con las cuatro observaciones de §11 para decisión del Ingeniero antes de H1.

---

## 11. Diferencias y decisiones que el Ingeniero debe conocer antes de H1

**Ninguna diferencia afecta montos, celdas ni el diseño del §2.** Son de citación o de receta.

| ID | Diferencia | Impacto | Propuesta |
|---|---|---|---|
| **D1** | El plan cita `CONSOLIDADO!D69`. La hoja real (plantilla v4 y los dos administrativos) se llama **`CONSOLIDADO_TOTAL RECAUDO`**; **no existe** ninguna hoja llamada `CONSOLIDADO`. | Ninguno sobre las celdas: `D69`, `D107`, `N69` se midieron en la hoja real. El «mapa de celdas» del proyecto usa la forma corta como apodo. | Mantener el apodo en prosa y usar el nombre real en código/tests/evidencia. |
| **D2** | El plan cita `R10 DetRetri!D10` y `DetRetri2026092`. No existe una hoja `DetRetri`: son **`DetRetri2026091`** y **`DetRetri2026092`** (nombre por período, ya gobernado por `HojasPeriodo`). | Ninguno: `D10`/`D12` son las celdas de `RECAUDO TOTAL` del ASE. | Igual que D1: nombre real en código. |
| **D3** | La suite del árbol tiene **25 rojos pre-existentes** (no es «suite verde»), y el trx que cita el gate de entrada (`p59-h5-suite.trx`, 1618/34) quedó desactualizado por tres commits posteriores del mismo día (hoy: 1674/25 con el fixture; 1657/25 sin él). | G-SUITE/G1 piden «suite verde». El set-diff de H0 es cero (subconjunto estricto). | Ratificar que las fases miden con **set-diff** contra el árbol al abrir cada H (no contra un trx viejo) y que el verde total es deuda previa, no de este plan. |
| **D4** | La huella del frontmatter del plan dice `…524d…`; la real es `…526d…` (= constante canónica `HuellaSemanticaPlantilla.V4`, verificada hoy y en el T0 del Plan 60). | Sólo el texto del plan. El gate G-HUELLA está intacto. | Corregir el carácter en el plan; el fixture afirma el valor medido. |
| **D5** | H3.3 propone re-anclar `ProcesadorPeriodoTests…_NoGeneraArchivo` y `VerificadorRecaudo` T19 a «no limpio + `(sin validar)`», pero ambos prueban el **writer** directamente (gates 3 y 4 de §4.2 viven dentro de `GenerarWorkbook`). | Afecta la receta de H3, no H0. | Decidir en H3: (a) los gates de coherencia siguen siendo duros **dentro del writer** y los tests quedan intactos; o (b) se sacan al procesador y ambos tests se reescriben pasando por `ProcesadorPeriodo`. Recomendación senior: (b), porque el requisito del Ingeniero es «nunca quedarse sin Excel», y un gate dentro del writer garantiza justo lo contrario. |

---

## 12. Artefactos

| Artefacto | Ruta | Estado |
|---|---|---|
| Fixture T0 | `Remuneracion.IntegrationTests/EspecialesR4T0Plan61Tests.cs` | nuevo, 17/17 verde |
| Evidencia | `Docs/T0-Plan61-Especiales-R4.md` (este documento) | nuevo |
| Código de producción | — | **sin cambios** (G-H0-READONLY cumplido) |
| Trx de la corrida focal | `Remuneracion.IntegrationTests/TestResults/p61-h0-foco.trx` | generado por el runner |
| Trx de la suite | `Remuneracion.IntegrationTests/TestResults/p61-h0-suite.trx` | generado por el runner |
