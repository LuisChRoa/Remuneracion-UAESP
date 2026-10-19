# T0 — Plan 47: revisión de las demás pestañas de pegado con la lente R1

> **Fase:** H0 (T0 read-only). **Fecha:** 2026-10-17. **Autor:** `implementer` (subagente de
> `remuneracion-architect`). **Plan fuente:** `plans/47 - revision-demas-pestanas-lente-r1.md`
> (APROBADO 2026-10-17, Clarification Gate §7 ratificado).
> **Fixture congelado (read-only):** `Remuneracion.IntegrationTests/RevisionPestanasLenteR1T0Plan47Tests.cs`
> → **42/42 verdes** (suite completa: 38 Failed / 1002 Passed / 1040 Total / 0 Skipped, mismos 38
> fallos que la línea base `p46-cierre.trx`: **0 altas / 0 bajas**).
> **Corrida de referencia:** `%TEMP%\opencode\p46-corrida9\Remuneración 202607-2 Total.xlsx` (corrida 9,
> frente cerrado en el Plan 46). **Oficial:** `Docs/Plantilla estandarizada/Formato Remuneracion
> 2026xx-x Total.xlsx` (SHA256 `3086AC31…C5214` intacto antes/después).
> **Metodología:** lectura por bytes (BCL ZIP+XML / `dumpxlsx.ps1` + `diffxlsx.ps1` corregido). El diff
> de **fórmulas** es RUIDO por *shared formulas* → **sólo `values_changed` es evidencia**; `<v>` nunca
> es oráculo. Este T0 **no corrige nada**: los hallazgos son candidatos a H1 (la corrección requiere la
> puerta).

---

## 1. H0.1 — Inventario oficial → corrida 9 (re-verificado por bytes)

`values_changed` = celdas presentes en ambos archivos con valor distinto. Se compara **por nombre de
hoja**; las 3 hojas de período se comparan además **por su gemela renombrada** (ver §7-P5).

| Hoja | `values_changed` | Sólo en oficial | Sólo en corrida | Qué escribe el runtime (P5) |
|---|---|---|---|---|
| `Reporte Componentes R1` | 3.496 | — | — | Frente **cerrado** en el Plan 46 (referencia) |
| `Rem. Anticipos R2` | **2.051** | — | — | Desglose-detalle por ASE (`EscribirCeldasDetalleR2R4`) |
| `Reversion Pagos R4` | **828** | — | — | Desglose-detalle por ASE (idem) |
| `SALDOS POR NOTA` | 237 | — | — | Operandos editables del ASE (`LeerSaldosNotas`) |
| `REPORTE RECAUDO x BANCO` | 30 | — | — | 4 filas de concepto por ASE × empresa (`EscribirCeldasBanco`) |
| `BCE SC POR FACT.` | 10 | — | — | `D3:D7` SUBSIDIO + `E3:E7` CONTRIBUCION (`EscribirCeldasBalanceSc`) |
| `CONSOLIDADO_TOTAL RECAUDO` | 15 | — | — | 13 tokens de período + `G7`/`K7` (fechas) |
| `AJUSTES - SF-T` | 6 | — | — | 6 tokens (`N3/AC3/N22/AC22/N41/AC41`) |
| `VALIDACION_TOTAL` | 8 | — | — | Sellado de cachés (Plan 40): `C3:C7`, `C9`, `C14`, `C15` |
| `Valida - Control Recaudo` | 2 | — | — | `F10` (total de control, literal) + `C2` (token) |
| `Valida -Remunera` | 2 | — | — | `N121`/`AD121` (tokens) |
| `Valida - Anticipos` | 0 | — | — | sin escritura |
| `GERENTES_TOTAL` / `GERENTES_Reciprocidad_ EAAB` / `GERENTES_ENEL` / `GERENTES_ENERBIT` / `GERENTES_OCCIDENTE _ Directa` / `GERENTES_EAAB-CL` | 3 c/u | — | — | `N3`/`AD3`/`N22` (tokens de período) |
| `REMUNERACION Reciprocidad EAAB` / `_ENEL` / `_ENERBIT` / `_OCCIDENTE_ Directa` / `_EAAB-CL` | 13 c/u | — | — | 13 tokens (`N3,N22,N41,N60,N79,N98,AD3,AD22,AD41,AD60,AD79,AD98,M117`) |
| `VALIDACION_RECIP` / `_ENEL` / `_ENERBIT` / `_OCCIDENTE` / `_EAAB-CL` | 0 | — | — | 100% fórmulas (protegidas) |
| `Recaudo EAAB Reciprocidad` | 18 | — | — | Pegado real (`RESUMEN MES`, cols F/G) |
| `Recaudo ENEL` | 36 | — | — | idem |
| `Recaudo ENERBIT` | 8 | — | — | idem |
| `Recaudo Directa Occidente` | 36 | — | — | idem |
| `Recaudo EAAB + Ciud Limp` | **0** | — | — | Pegado real de **ceros** (fuente en 0, §7-P5) |
| `ANT EXT-REV`, `ANTICIPOS USUARIOS`, `INTERVENTORIA`, `RETRIBUCION NEGATIVA` | 0 | — | — | sin escritura (fórmulas protegidas / fuente vacía) |
| `Informe AFaseo Recaudo 202609-2` | — | **sí** | **sí** (`…202607-2`) | Sólo **renombrada** por el preparador: par renombrado → **0** valores |
| `DetRetri2026092` | — | **sí** | **sí** (`DetRetri2026072`) | Par renombrado → **19** valores (DetRetri-Q2 + totales) |
| `DetValiRetri2026092` | — | **sí** | **sí** (`DetValiRetri2026072`) | Par renombrado → **8** valores (sello fecha/hora + `I9:I13` + token) |

**Nota metodológica vinculante (hallazgo de herramienta, no de producto).** Los conteos del §0.2 del
plan para las hojas de septiembre (**DetRetri 293 / DetValiRetri 290 / AFaseo 164**) son un **artefacto
del script viejo** `diffxlsx.ps1`: cuando una hoja existe en A y no en B, el script comparaba cada
celda contra el **último `$vb` de la hoja anterior** (`$vb = $cb[$k][0]` sobre `$cb` nulo → la variable
conserva el valor previo) y reportaba "todo cambió" (≈ el número de celdas de la hoja: 297/293/164
celdas con valor o fórmula medidas por bytes). La comparación correcta por **par renombrado** da
**19 / 8 / 0** (§7-P5). El script de este T0 (`%TEMP%\opencode\p47-diff.ps1`) compara por nombre de
hoja y clasifica sólo_A/sólo_B; el ruido queda documentado para que no reaparezca.

**Ruido de fórmulas (no es evidencia).** El diff reporta `formula_presence_changed` en `BCE SC POR
FACT.` (6), `VALIDACION_EAAB-CL` (5), `VALIDACION_ENERBIT` (4), `VALIDACION_RECIP` (2) y
`Informe AFaseo 202609-2` (72). Son los **masters/followers de *shared formulas*** (el texto del `<f>`
vive en el maestro y los seguidores lo referencian por `si`) y los **6 tokens `<f>` reescritos por el
preparador de base** (p. ej. `BCE!H3: DetRetri2026092!J9 → DetRetri2026072!J9`,
`VALIDACION_EAAB-CL!C4: 'Recaudo EAAB + Ciud Limp'!D22 → !F22`). Nada de esto se interpreta como
escritura de la lente.

---

## 2. H0.2 — P6: matrices congeladas por pestaña

### 2.1 `Rem. Anticipos R2` — (etiqueta de fila) → (slot de fila)

- **Firma/clave por fila:** `A|B|C|D` (leída de la plantilla para el destino y de la fuente para el
  origen). Bloques declarados (`WorkbookLeafCellMapDetalleR2R4.BloquesR2`) y **re-verificados por
  bytes**: ASE1 `3..33`, ASE2 `67..113`, ASE3 `164..194`, ASE4 `282..328`, ASE5 `382..419`; en los 5,
  la primera fila es `D=Vlr Servicio` y la última `A=Total` con `B/C/D` vacíos.
- **(Encabezado de componente) → (columna destino):** fila 2 `E..R` =
  `Total, Componente TDF, Componente TTL, Componente TVIAT, Aprovechamiento, CCSA Prest.Aprov.,
  Especiales, Componente TCS, Componente TLU, Componente TBL, Componente TRT, CCSA Prest. No Aprov.,
  Deb/Cred, Deb/Cred-otros` (14). La fuente trae **12** encabezados (`E..P`, sin `Especiales`) en
  ASE1/2/4 y **13** (`E..Q`, con `Especiales`) en ASE3/5; `Deb/Cred-otros` (R) **no existe** en
  ninguna fuente julio-2 (0 explícito vía `?? 0m`).
- **Regla de desempate cuando falta etiqueta:** hoy **no hay** regla de canal; el emparejamiento es
  la **LCS por firma con desempate `dp[x+1,y] >= dp[x,y+1]` (salta destino)**
  (`OpenXmlPlantillaWriter.AlinearFirmas`), de modo que las firmas **duplicadas** de las filas sin
  etiqueta (`|||Vlr Servicio`, `|||Vlr Intereses`, `|||Subsidio(-)/Contribucion(+)`) casan **por
  orden físico** (Plan 41 E1-1 prohibido). La matriz completa por ASE (firmas destino/fuente + pares +
  filas sin fuente + omitidas) está congelada en el fixture
  (`T0P47_P6_R2_MatrizFilasYEmparejamiento_Congelada`, 5 casos verdes).

### 2.2 `Reversion Pagos R4` — (etiqueta de fila) → (slot de fila)

- **Firma/clave por fila:** `A|B|C`. Bloques (`BloquesR4`) re-verificados: ASE1 `3..15`, ASE2 `93..111`,
  ASE3 `194..206`, ASE4 `231..249`, ASE5 `346..358`; primera fila `C=Vlr Servicio`, última `A=Total`.
- **(Encabezado) → (columna):** fila 2 `D..R` = `Total … Ajuste Decena, SERVICIO ESPECIALES, Ajuste
  Recargo Aseo, Otros`; la fuente trae `D..O` (12) y **ASE4 trae además `Otros`** (13).
- **Mismo defecto de clave:** LCS por firma con firmas duplicadas → emparejamiento por orden (§4).

### 2.3 `SALDOS POR NOTA`, `REPORTE RECAUDO x BANCO`, `BCE SC POR FACT.`

- **`SALDOS POR NOTA`:** bloque por concesionario; etiquetas de fila = `Vlr Servicio`, `Vlr Intereses`,
  `Componente`, `Subsidio(-)/Contribucion(+)`, `Subs/Cont`, `Total` (filas 3..8 para ASE1; cierre en la
  fila `Total`); columnas `C..O` = componentes (`Especiales` en I). Mapeo **por título y por concepto
  explícito** (`ConceptosSaldosPorAse`), con `Vlr Intereses` marcado **opcional** → **no hay
  emparejamiento por orden** (la fuente ASE1 trae 5 filas y la plantilla 6: la sexta queda en 0
  explícito).
- **`REPORTE RECAUDO x BANCO`:** filas = disposición (`1-APLICADOS A FACTURACION`,
  `2-SALDOS A FAVOR GENERADOS`, `3-APLICADOS A FINANCIACIONES NUEVAS`,
  `7-APLICADOS A RECIBOS SERVICIOS ESPECIALES`, `Total` en fórmula); columnas = **empresas bancarias**
  (`ENEL, CIUDAD LIMPIA, EAB, NUEVO ESQUEMA, ENERBIT, OCCIDENTE, Total`). La escritura usa celdas
  explícitas por ASE×empresa (`EditablesPorAse`) y la lectura valida el **nombre** del encabezado en la
  posición declarada (fail-fast si no coincide) + gate Σ conceptos = Total fuente.
- **`BCE SC POR FACT.`:** filas `3..7` = ASE `1..5` (col C), `D=SUBSIDIO`, `E=CONTRIBUCION`; `F/H/I` son
  fórmulas. Lectura de la fuente por la fila **`Total General`** (búsqueda desde el final) y columnas
  E/F/G (índices 4/5/6) del layout uniforme de las 5 fuentes.

---

## 3. H0.2 — P1: alcance de la lente (dentro/fuera con criterio citado)

| Pestaña | Veredicto | Criterio citado (bytes) |
|---|---|---|
| `Rem. Anticipos R2` | **DENTRO** | Filas con `D=concepto` + filas de cierre por empresa/canal; columnas `E..R` = componentes por encabezado; fuente `RerpoteDetalleSaldosaFavor` |
| `Reversion Pagos R4` | **DENTRO** | Filas `A/B/C` + columnas `D..O` = componentes; fuente `ReversiónPorComponente` |
| `SALDOS POR NOTA` | **DENTRO (lente parcial)** | Filas de concepto (B) × columnas `C..O` de componentes; fuente propia `SaldosaFavorAplicadosPorNotas` (el T0 **descarta** que sea derivada de R2: no comparte archivo ni valores de encabezado) |
| `REPORTE RECAUDO x BANCO` | **FUERA** | Otra lente: columnas = bancos, filas = disposición (`1-…/2-…/3-…/7-…/Total`), no componentes |
| `BCE SC POR FACT.` | **FUERA** | Agregado: 5 filas ASE × {SUBSIDIO, CONTRIBUCION}; el fix, si existiera, vive en el origen (fuente Balance) |
| `CONSOLIDADO_TOTAL RECAUDO` | **FUERA** | Arrastre: `D28 = 'Rem. Anticipos R2'!E47`, `D29..D32 = …!E144/E262/E362/E446`, `D33 = SUM(D28:D32)` |
| `AJUSTES - SF-T` | **FUERA** | Arrastre: `D47 = D9+D28` (SALDOS + RETRIBUCION) |
| `VALIDACION_TOTAL`, `Valida -*` | **FUERA** | Gate/evaluador (Plan 40): `C9 = SUM(C3:C8)`, `C14 = 'Valida - Control Recaudo'!F10`, `C15 = C14=C9` → se evalúa/sella, no se pega |
| `GERENTES_*`, `REMUNERACION_*`, `VALIDA_*` | **FUERA** | Metadatos/arrastre de período: los únicos cambios son tokens `2026092→2026072` (3 y 13 celdas) |
| `Recaudo *` | **EXPLICADO (fuera de la lente)** | Pegado real (values-only) desde `Conciliaciones/*` hoja `RESUMEN MES`, columnas F/G (2ª quincena) — §7-P5; **no** es fuente/canal/concepto/componentes |
| `DetRetri2026092` / `DetValiRetri2026092` / `Informe AFaseo 202609-2` | **EXPLICADO (fuera de la lente)** | Hojas de período **renombradas** por el preparador (Plan 30/39) — §7-P5; escritura sólo en DetRetri (19) / DetValiRetri (8); AFaseo 0 |
| `INTERVENTORIA`, `RETRIBUCION NEGATIVA`, `ANT EXT-REV`, `ANTICIPOS USUARIOS`, `Recaudo EAAB + Ciud Limp` | **FUERA (0 cambios probado)** | Fuentes vacías (Retribución: sólo la fila 1) o ausencia de escritura; `Recaudo EAAB + Ciud Limp` ← `Conjunta Otros` con F en 0 |

---

## 4. H0.2 — P2: ¿filas-fuente sin etiqueta que hoy casan por orden? **SÍ — corrimiento probado**

**Mecanismo (por qué ocurre).** Las filas sin etiqueta propia de la fuente (`Vlr Servicio`, `Vlr
Intereses`, `Subsidio(-)/Contribucion(+)`) tienen firma duplicada en la plantilla (una por grupo:
A/EAB, E/ENEL, 1, 2, 5, H, O, T…). La LCS las empareja **en orden físico**; cuando el bloque destino
tiene un grupo que la fuente **no** trae (o al revés), los valores del grupo siguiente **entran en las
filas del grupo ausente** y las filas propias de ese grupo quedan en su 0 de plantilla.

**Regla que lo tapa:** `AlinearFirmas` (LCS + desempate `>=`) + `EmparejarBloque` + la doctrina del
mapa: *"las filas-fuente sin contraparte se omiten y las filas-destino sin fuente quedan en su 0 de
plantilla"* (`WorkbookLeafCellMapDetalleR2R4`, líneas 9-12).

**Corrimientos probados (fila-fuente → celda-slot; canal efectivo = R46-1).**

| Pestaña | ASE | Fila-fuente (canal efectivo) | Slot que recibe hoy (canal efectivo del slot) | Evidencia de valor |
|---|---|---|---|---|
| `Rem. Anticipos R2` | **1** | `src#0 Vlr Servicio` (canal **E**, cierre `C=E D=Total`) | **`F3..R3`** (canal **A**, cierre `C=A D=Total`) | `E3 = 104,754,634.94` (corrida 9) = Total de src#0 |
| `Rem. Anticipos R2` | **1** | `src#1 Vlr Intereses` (canal E) | **`F4..R4`** (canal A) | `E4 = 15,816.77` |
| `Rem. Anticipos R2` | **2** | `src#0/#1` (canal E) | **`F67..R67` / `F68..R68`** (canal A) | `E67 = 71,702,408.71` |
| `Rem. Anticipos R2` | **4** | `src#0/#1` (canal E) | **`F282..R282` / `F283..R283`** (canal A) | `E282 = 76,457,865.81`, `E283 = −546,491.49` |
| `Rem. Anticipos R2` | **5** | `src#0/#1` (canal E) | **`F382..R382` / `F383..R383`** (canal A) | `E382 = 19,713,558.23`, `E383 = 18,754.01` |
| `Reversion Pagos R4` | **2** | `src#6..10` (grupo **OCCIDENTE**) | **`D99:E103`** (grupo **NUEVO ESQUEMA**, cierre fila 104) | `D99 = −377,166.80`, `D100 = −9,054.66`, `D101 = −386,221.46`, `D102/D103 = 95,471.46` |

**Identidad probada (contra-casos, sin desvío):** R2 **ASE3** (la fuente trae el grupo A/EAB:
`src#0 Vlr Servicio` → `F164`, canal A = canal A) y R4 **ASE1/3/4/5** (la fuente trae los mismos grupos
que el destino; ASE4 trae NUEVO ESQUEMA). Queda congelado en
`T0P47_P2_R2_CanalEfectivoDeLaPrimeraFila_VsSlot` y `T0P47_P2_R4_CanalEfectivoDelSegundoGrupo_VsSlot`.

**Impacto medido: NO altera la aritmética.** Ninguna `<f>` de las hojas R2/R4 consume filas sin
etiqueta: los únicos consumos son las filas de **total con etiqueta única** (`E21/E32/K21`, `E94/E112/
E70/K70/E83/K83/E104/E97/E74/K74/E93/E111/K93` para R2; `D15/P15`, `D7+D13`, `D8/P8`, `D14/P14` para
R4) — probado por bytes en `T0P47_Impacto_NingunaFormulaConsumeFilasSinEtiqueta` (2 casos verdes). El
corrimiento es de **desglose/auditoría**, no de totales.

---

## 5. H0.2 — P3: bloques mal tipados (frontera que abre vs la que cierra)

**Sin hallazgo (identidad) con evidencia.** Los 10 bloques (5 R2 + 5 R4) se tipan **por la frontera que
CIERRA**: la última fila del rango declarado es `A=Total` con `B/C/D` vacíos y la primera es
`Vlr Servicio` (`T0P47_P3_BloquesR2_*` / `T0P47_P3_BloquesR4_*`, 10 casos verdes). No hay hermano del
defecto D2 del Plan 46 en estas pestañas: la frontera de cierre es la que gobierna el rango, y el rango
coincide con el cierre real en los 5 ASE de cada hoja.

---

## 6. H0.2 — P4: ¿falso hueco? (omisiones y "queda en 0" re-probados)

**6.1 Omitidas con y sin slot** (`T0P47_P4_R2_Omitidas_SinSlot_ProbadasPorBytes`):

| Pestaña | ASE | Omitida (fila-fuente) | Total omitido | Prueba |
|---|---|---|---|---|
| R2 | 2 | `src#8 Vlr Intereses` (canal 2) | 406,42 | El destino trae **1 sola** fila entre los cierres `1` y `2` (fila 78); la fuente trae 2 |
| R2 | 2 | `src#13 Vlr Servicio` (canal 9) + `src#14 9 Total` | 831.919,29 | **No existe** ninguna fila `C=9 D=Total` en el bloque oficial |
| R2 | 4 | `src#8 Vlr Intereses` (canal 2) | −8.785,59 | idem ASE2 |
| R2 | 4 | `src#13/#14 (canal 9)` + `src#15 9 Total` | 20.150.395,25 | idem |
| R2 | 4 | `src#37 Subsidio(−)/Contribucion(+)` (canal 9) + `src#38 9 Total` | 5.663.564,75 | El destino no tiene fila `9` en la mitad Subs/Cont |
| R2 | 1/3/5, R4 1..5 | — | — | 0 omitidas (todas las filas-fuente tienen slot) |

Los **6 casos con omisión** (ASE2 ×3, ASE4 ×6 celdas-fila) son omisiones **silenciosas** hoy: el mapa
no las reporta ni falla. No hay pérdida de total (los importes viajan en las filas de total con
etiqueta única), pero **sí** falta el desglose y no hay constancia.

**6.2 Falso hueco (contraparte real que hoy queda en 0)** — el precedente del 46 (un "cero
documentado" que era un falso hueco) **se repite**:

- **R2 ASE1/2/4/5** (`T0P47_P4_R2_FalsoHueco_SlotsSinFuenteConContraparteReal`): las filas `Vlr
  Servicio`/`Vlr Intereses` del **canal E** (p. ej. `E7`/`E8` en ASE1, `E71/E72` en ASE2, `E286/E287` en
  ASE4, `E386/E387` en ASE5) quedan en el **0 de la plantilla** (celdas de VALOR, sin `<f>`) aunque su
  fila-fuente **sí existe** y hoy cayó en el grupo A/EAB (§4). No es un hueco real: es un **desvío de
  slot**.
- **R4 ASE4 — columna `Otros`** (`T0P47_P4_R4_Ase4_Otros_TieneContraparteYNoSeEscribe`): la fuente
  `4-Bogota Limpia/ReversiónPorComponente_*` trae el encabezado **`Otros`** (13.º, col P) con valores
  (`P6 = −0,29`, `P7 = −0,29`, `P10 = −0,29`, `P23 = −0,29`) y el oficial tiene la columna **R `Otros`**
  como **VALOR** (literal `0`, `R232` sin fórmula). El mapa (`ColumnasR4`, 12 entradas) **no declara**
  `Otros` → el valor se descarta. *"Prohibido mantener un cero documentado sin re-probarlo"* (plan
  §1.2-P4) → **falso hueco probado**.

**6.4 Omitidas de AGOSTO-2 (H1 — evidencia ampliada).** La corrida de agosto
(`Docs/Prueba Agosto-2/Insumos`) ejercita el MISMO contrato sobre los mismos bloques oficiales. El
análisis por bytes (mismo espejo del §6.1, `P47PER=Prueba Agosto-2`) da **8 filas-fuente sin slot**:
R2 ASE2 (`Vlr Servicio` canal 9 = 81.263,29; `||9|Total` = 81.263,29; `Subsidio(-)/Contribucion(+)`
canal 9 = 6.719,84) y R2 ASE4 (`Vlr Intereses` canal 2 = −1.647,69; `Subsidio` canal A = −12.462,26;
`||A|Total` = −12.462,26; `|EAB|Total|` = −12.462,26). Las claves coinciden con las de julio-2 salvo
las tres de la **mitad Subs/Cont** (`Subsidio` canal 9 y grupo A/EAB): el bloque oficial de ASE4 tiene
UN solo grupo A/EAB (mitad Anticipos, `||A|Total` en 284) y la mitad Subs/Cont no lo enumera. Todas
quedan en la lista documentada (`CerosDocumentadosDetalleR2R4`: 11 claves / 14 filas declaradas).

**6.3 Ausencias probadas (no son falso hueco)** — con la contraparte **inexistente** verificada por
bytes: `Deb/Cred-otros` (R) en R2 (ninguna fuente julio-2 lo trae); `Especiales`/`Vlr Intereses` de
`SALDOS POR NOTA` (la fuente ASE1 no trae `Especiales` — su header en esa posición es `Componente TCS` —
ni fila `Vlr Intereses`); `Recaudo EAAB + Ciud Limp` (la fuente `Conjunta Otros` trae F/G en 0);
`RETRIBUCION NEGATIVA` (fuente julio-2 con sólo la fila 1).

---

## 7. H0.1/P5 — Qué escribe el runtime en cada hoja y desde qué insumo

**7.1 `Recaudo *` = PEGADO REAL (no arrastre), desde `Conciliaciones`.** `ExcelDataReaderWorkbookLeafInputReader.LeerRecaudoEmpresa`
lee `Docs/Prueba Julio-2/Insumos/Conciliaciones/{PrefijoConciliacion}*.xlsx` (p. ej.
`Conjunta ENEL-072026.xlsx`, `Directa-072026.xlsx`, `Conjunta Otros -072026.xlsx`, `Conjunta ENERBIT-072026.xlsx`,
`Conjunta Recip -072026.xlsx`), hoja **`RESUMEN MES`**: detecta los 3 bloques por la corrida de etiquetas
`ASE 1..5` (col C) y pega las columnas **F/G** (2ª quincena) en las filas `OPORTUNO 3..9`,
`EXTEMP 12..18`, `TOTAL 21..27` de cada hoja `Recaudo *` (`EscribirCeldasEmpresa` → `EscribirValorNumerico`).
Verificado fila a fila: `Recaudo ENEL` `F3..F9 = 16.467.129.562 / 19.037.496.395 / 13.952.517.350 /
6.579.568.126 / 11.684.535.987` y `F9 = 67.721.247.420 = Σ F3:F7` (los 5 valores son bytes de la fuente).
`Recaudo EAAB + Ciud Limp` ← `Conjunta Otros -072026.xlsx`: F3..F9 todo 0 en la fuente → **0 cambios con
ausencia probada**. **Riesgo documentado (no defecto hoy):** el mapeo fila-fuente→fila-destino es
**posicional** (`filasInicio[b] + i`), aunque la etiqueta `ASE n` existe en la fuente; hoy coincide
porque el gate exige la corrida `1..5` en orden (verificado en la fixture).

**7.2 Hojas de período con sello de septiembre = PLACEHOLDER, no escritura cruzada.** El oficial trae
`DetRetri2026092`, `DetValiRetri2026092`, `Informe AFaseo Recaudo 202609-2`, el token literal
`CONSOLIDADO!N3 = 2026092` y las fechas `G7 = 46281` (16/09/2026) y `K7 = 46295` (30/09/2026); además
`BCE!H3:H7` referencian `DetRetri2026092!J9..J13` **por nombre**. El **preparador de base**
(Plan 30/39, `PreparadorBasePipeline` + `NombresHojaPeriodo`) **renombra** las 3 hojas y **reescribe**
los tokens/fechas en la copia de trabajo: la corrida 9 tiene `DetRetri2026072`,
`DetValiRetri2026072`, `Informe AFaseo Recaudo 202607-2`, `N3 = 2026072` y `G7/K7 = 46219/46234`
(16/07/2026 → 31/07/2026). Comparación **por par renombrado**: `DetRetri` 19 valores (DetRetri-Q2 por
ASE + `D14`), `DetValiRetri` 8 (`K3`, `I9:I13`, sello `D6` fecha 04/08/2026 y `D7` hora 10:15),
`AFaseo` **0**. El oficial nunca se escribe (SHA256 intacto).

**7.3 Resto de la cadena (una línea por hoja):**
`Recaudo *` → conciliaciones (7.1). `SALDOS POR NOTA`/`RETRIBUCION NEGATIVA` → `SaldosaFavorAplicadosPorNotas`/
`RetribuciónNegativa` (`LeerSaldosNotas`/`LeerRetribucionNegativa`). `BCE SC POR FACT.` → `R4-BalanceSubsidioyContribuciones`
(fila `Total General`, cols E/F). `REPORTE RECAUDO x BANCO` → `ReportePagosxBanco` (`LeerReporteBanco`).
`INTERVENTORIA`/`ANT EXT-REV`/`ANTICIPOS USUARIOS` → fórmulas protegidas (0 escrituras).
`CONSOLIDADO`/`AJUSTES`/`GERENTES_*`/`REMUNERACION_*`/`VALIDACION_*` → arrastre de fórmulas + tokens de
período. `Reporte Componentes R1` → frente cerrado del 46.

---

## 8. H0.3 — Medición fila a fila contra el insumo y contra la corrida 9 (S1.6)

**R2 (`Rem. Anticipos R2`), ASE1 — coincide:** filas de total con etiqueta única
(`E9=E10=104.770.451,71` = `E3+E4`; `E21=126.791.992,42`; `E32=−843.439,30`; `E33=125.948.553,12`) y
`E47 = E21+E32−K21`. **No coincide (desvío §4):** `E3/E4` reciben lo del canal E; `E7/E8` quedan en 0.
ASE2/ASE4/ASE5: idéntico patrón; ASE3: 100% identidad.

**R4 (`Reversion Pagos R4`), ASE2 — coincide:** `D93..D98` (grupo ENEL: VS −53.998.016,21; VI
−15.789.581,83; Componente −69.787.598,04; Subs/Cont −991.410,96) y `D110/D111`. **No coincide:**
`D99..D103` reciben el grupo **OCCIDENTE** de la fuente (−377.166,80 / −9.054,66 / −386.221,46 /
95.471,46 / 95.471,46), `D104 = 0` (NUEVO ESQUEMA) y `D105..D109 = 0` (las filas propias de OCCIDENTE).
ASE1/3/4/5: identidad.

**Contra el insumo (fuente):** todas las filas de **total** (canal/empresa/bloque) coinciden ±0 (no se
aplicó tolerancia: son literales copiados); los desvíos son exclusivamente las filas **sin etiqueta**
de grupo (detalle), más las 9 filas-fuente **omitidas** de ASE2/ASE4 (§6.1).

**Contra la corrida 9 (celdas escritas):** `2.051` (R2) y `828` (R4) celdas cambiadas son exactamente
las celdas de las filas **emparejadas** × 14 (R2) / 12 (R4) columnas declaradas en el mapa (los 0 de
plantilla no cuentan como cambio). No hay ninguna celda escrita fuera de las filas emparejadas
(comprobado: las filas sin fuente permanecen en `0`).

---

## 9. Veredicto de la puerta H0→H1

**(a) HAY DESVÍO PROBADO — candidato a H1**, en dos pestañas:

1. **`Rem. Anticipos R2`** — corrimiento por orden en **ASE1, ASE2, ASE4 y ASE5** (el grupo E cae en las
   filas del grupo A; las filas propias del grupo E quedan en 0), más **omisiones silenciosas** en
   ASE2 (3 filas) y ASE4 (6 filas) por ausencia de slot (`canal 9`, `Vlr Intereses` del canal 2).
   Regla a aplicar: **R46-1 (canal efectivo)**; el slot destino debe compartir el canal efectivo de la
   fila-fuente (la que CIERRA su grupo).
2. **`Reversion Pagos R4`** — corrimiento por orden en **ASE2** (grupo OCCIDENTE bajo el encabezado
   NUEVO ESQUEMA) y **falso hueco de columna** en **ASE4** (`Otros` existe en la fuente y en la
   plantilla, el mapa no la declara).
   Además, la alineación debe dejar de **casar por orden** (Plan 41 E1-1): las filas sin slot probado
   deben ser omisión **documentada/reportada** (o fail-fast accionable), no silencio.

**(b) SIN HALLAZGO CON EVIDENCIA — no se tocan:** `SALDOS POR NOTA` (mapeo por título + opcional;
ausencias probadas), `REPORTE RECAUDO x BANCO` (otra lente; hoy alineada por etiqueta ASE; se
documenta el riesgo posicional), `BCE SC POR FACT.`, `CONSOLIDADO_TOTAL RECAUDO`, `AJUSTES - SF-T`,
`VALIDACION_*`, `Valida -*`, `GERENTES_*`, `REMUNERACION_*`, `INTERVENTORIA`, `RETRIBUCION NEGATIVA`,
`ANT EXT-REV`, `ANTICIPOS USUARIOS`, `Recaudo *` y las 3 hojas de período (renombradas por el
preparador).

**No hay STOP**: la matriz se absorbe sin ramas nuevas por período/empresa (la corrección, si el
Ingeniero la autoriza, es un predicado puro de geometría — canal efectivo — más el reporte de
omisiones sin slot). **P3 = identidad** (los bloques ya se tipan por cierre).

**Alcance propuesto para H1 (mínimo, pendiente de autorización):** R2 y R4 únicamente; clave
`(canal efectivo + firma)`; conservar los totales intactos; decidir con el Ingeniero el tratamiento de
las 9 filas sin slot (omitir con reporte vs fail-fast) y de la columna `Otros` de R4 (¿se escribe en R
o se declara formalmente ausente?).

---

## 10. Gates de la fase (G-BUILD / G-SUITE / G-Q1 / G-SHA / G-PROC)

| Gate | Resultado |
|---|---|
| **G-BUILD** | `dotnet build Remuneracion.WinForms/Remuneracion.WinForms.slnx` → **0 warnings / 0 errores** |
| **G-SUITE** | `Remuneracion.IntegrationTests/TestResults/p47-h0.trx`: **38 Failed / 1002 Passed / 1040 Total / 0 Skipped**; set-diff vs `p46-cierre.trx` → **0 altas / 0 bajas** (`p47-h0-setdiff.txt`); los 42 tests nuevos entran en Passed (960 + 42 = 1002) |
| **G-Q1** | `SeccionEspejoR1Tests.T1_Ase2Q1_NuevoEsquemaEnPrimeraSeccion_ErrSinSlotSinFallback` → **Passed** (Q1 intacto) |
| **G-SHA** | Oficial `3086AC317595C2F8E84ED24B1C5EE3E52565DD5193394E3062878E7FA49C5214` intacto antes/después |
| **G-PROC** | **Cero commits** (sólo archivos nuevos sin stage); fixture y docs con **CRLF sin mezclar** |
| **G-FORMULA** | Cero código de producción tocado en H0 (`<f>`, insert/delete y allow-list Phase 0 sin cambios) |
| **Q1 / plantilla / auditoría** | Q1 fuera de alcance; plantilla intocable; auditoría de la propietaria sigue PENDIENTE (no se declara paridad/verde por ausencia) |

## 11. Hallazgos de herramienta (no de producto)

- `%TEMP%\opencode\diffxlsx.ps1` (versión previa) infla los conteos de las hojas presentes sólo en A
  (§1). Este T0 usa `%TEMP%\opencode\p47-diff.ps1` (clasifica sólo_A / sólo_B / values_changed / cambio
  de presencia de fórmula) y `%TEMP%\opencode\p47probe` (plan/pares/valores). Ambos son read-only y
  viven fuera del repo.
- `dumpxlsx.ps1` sirve para volcados por rango; recordar que el **índice de columna de ExcelDataReader
  coincide con la letra** (A=0) - la confusión previa vino de volcar un archivo distinto
  (`Recaudoporcomponente` vs `RerpoteDetalleSaldosaFavor`).

---

## 12. H1/H2 — EJECUCIÓN (2026-10-17, autorizada por el Ingeniero)

### 12.1 Alcance implementado (traza a los puntos 1–4 de la autorización)

| Punto | Qué se hizo | Archivo (dónde) |
|---|---|---|
| **1. Canal efectivo por cierre (R46-1)** | Predicado PURO nuevo: `Rotulo` (A–D), `EtiquetaDeGrupo`, `TieneEtiquetaPropia`, `CanalEfectivo` (la etiqueta de grupo de la fila `Total` que CIERRA el grupo; `null` si la fila trae etiqueta propia) | `Remuneracion.Core/Models/CanalEfectivoDetalleR2R4.cs` (L26/L47/L74/L97) |
| | Columnas de rótulo declaradas por el mapa: propia R2 `[A,B,C]` = `[0,1,2]`, grupo R2 `[C,B]` = `[2,1]`; propia/grupo R4 `[A]` = `[0]` | `Remuneracion.Infrastructure/Excel/WorkbookLeafCellMapDetalleR2R4.cs` (L37/L45/L51/L59) |
| | Emparejamiento por clave `(firma + canal efectivo)`, primer slot libre; el LCS por orden (`AlinearFirmas`), `LeerFirmaDestino` y `NormalizarFirma` se RETIRARON | `OpenXmlPlantillaWriter.cs` (L1566 `EmparejarBloque`, L1639 `LeerRotulos`, L1664 `Firma`, L1677 `MensajeSinSlot`) |
| **2. Falsos huecos** | Los slots VS/VI del canal E (`R7/R8`, `R71/R72`, `R286/R287`, `R386/R387`) reciben su valor real (antes en 0); los grupos sin fuente (A/EAB, NUEVO ESQUEMA) quedan explícitamente en el 0 de la plantilla | ver §12.3 (tabla corrida-9 → nueva) |
| **3. `Otros` en R4** | `("Otros","R")` declarada por ENCABEZADO en `ColumnasR4` (13 columnas); `SERVICIO ESPECIALES P`/`Ajuste Recargo Aseo Q` siguen fuera (ausencia probada) | `WorkbookLeafCellMapDetalleR2R4.cs` (L135) + doc en `Core/Models/DetalleR4AseInputs.cs` |
| **4. Filas sin slot** | Lista documentada en Core (11 claves / **14 filas**: 9 julio-2 + 5 agosto-2) con motivo + evidencia citada; `CantidadDocumentada` acota el consumo por corrida; fuera de la lista ⇒ `ERR-SIN-SLOT` accionable (hoja + ASE + firma + canal) | `Remuneracion.Core/Models/CerosDocumentadosDetalleR2R4.cs` (L35/L59/L91/L97) + `EmparejarBloque` (L1626-1631) |

### 12.2 Fixtures y tests actualizados (R-FIXTURE: solo con evidencia citada)

- `Remuneracion.IntegrationTests/DetalleR2R4EscrituraTests.cs`:
  - `JulioAse1_CeldasCitadas_IgualanFuenteYManual` (L51): `E3=E4=0` (grupo A/EAB sin fuente) y
    `E7=104.754.634,94` / `F7=6.975.298,69` / `K7=0` / `E9=104.770.451,71`; el manual legacy se cita en
    su propia fila 3.
  - `AgostoAse1_EspecialesEnK_IgualaFuenteYManual` (L171): `E3=0`, `E7=68.638.686,24`,
    `F7=3.508.770,06`, `K7=144.941,68`.
  - `AssertFilaVlrServicioR2/R4` (L219/L252) reescritos: localizan el slot por la clave
    `(firma + canal efectivo)` con el predicado de PRODUCCIÓN y verifican que la primera fila del
    bloque NO reciba el valor cuando su canal difiere.
  - `DetalleR2_DestinoFormula...` (L190): la celda mutada pasa de `E3` a `E7` (la que el writer escribe).
- `Remuneracion.IntegrationTests/RevisionPestanasLenteR1T0Plan47Tests.cs`: **35 tests verdes** —
  matrices H1 congeladas de R2/R4 (pares por clave + filas sin fuente + ceros documentados con canal),
  identidad de canal en TODO par, espejo local vs predicado de producción (julio + agosto),
  ceros documentados (9 julio / 8 agosto, todas cubiertas), `Otros`, P3 (cierre), P5 y el gate de
  "ninguna `<f>` consume filas sin etiqueta".

### 12.3 Tabla corrida-9 → corrida nueva (temp fresca desde el OFICIAL + `Docs/Prueba Julio-2/Insumos`)

CLI: `--periodo 2026072 --quincena 2 --cinco-ase` → **codigo=0** (sin reporte de divergencia).
Diff por bytes: **solo 2 hojas cambian**; las otras 38 dan `values_changed=0` (incluidas
`SALDOS POR NOTA`, `Recaudo *`, `BCE`, `CONSOLIDADO`, `AJUSTES - SF-T`, `VALIDACION_*`, `GERENTES_*`,
`REMUNERACION_*`, `DetRetri/DetValiRetri/AFaseo`).

| Hoja | Celdas | Filas | Cambio (corrida 9 → nueva) | Causa citada |
|---|---|---|---|---|
| `Rem. Anticipos R2` | 192 | R3/R4, R67/R68, R282/R283, R382/R383 → **0**; R7/R8, R71/R72, R286/R287, R386/R387 ← **valor del canal E** | 4 ASE (1/2/4/5) × {VS,VI} × 12 columnas (E..R menos K/R, ya 0) | Puntos 1+2: el grupo A/EAB no existe en la fuente; el par del canal E toma su slot por cierre (ASE1 testigo autorizado) |
| `Reversion Pagos R4` | 116 | R99..R103 → **0**; R105..R109 ← **grupo OCCIDENTE** | ASE2 × 5 filas × 11–12 columnas | Punto 1: la fuente no trae NUEVO ESQUEMA (destino 99..104) → OCCIDENTE conserva su slot por cierre `R110` |
| `Reversion Pagos R4` | 4 | R232/R233/R236/R249 ← **−0,29** | ASE4 (columna R `Otros`) | Punto 3: `Otros` declarada por encabezado; 4 filas con dato en la fuente |

Total: **312 celdas** (192 + 120) — todas explicadas; cambio inexplicado = 0.
`<f>`: **13.000 celdas con fórmula en ambas corridas y 0 diferencias de texto en las 40 hojas**.

### 12.4 Gates (H1→H2 y H2→H3)

| Gate | Resultado |
|---|---|
| G-BUILD | **0 warnings / 0 errores** |
| G-SUITE | `TestResults/p47-h1.trx`: **38 Failed / 995 Passed / 1033 Total / 0 Skipped**; set-diff de FALLOS vs `p46-cierre.trx` → **0 altas / 0 bajas** (mismos 38); los 35 tests nuevos entran en Passed (998 + 35 = 1033) |
| G-Q1 | `T1_Ase2Q1_NuevoEsquemaEnPrimeraSeccion_ErrSinSlotSinFallback` → **Passed** |
| G-FORMULA | `<f>` byte-idéntica (0 diferencias de texto, 40 hojas, 13.000 celdas con fórmula) |
| G-SHA | Oficial `3086AC31…C5214` intacto |
| G-PROC / D-H | Cero commits; CRLF sin mezclar; cero literales/ramas por período o empresa en el código nuevo (las etiquetas son DATOS; el único literal estructural es `Total`) |
| Muertos | `AlinearFirmas` / `LeerFirmaDestino` / `NormalizarFirma` retirados (sustituidos por la clave con canal) |

**Veredicto H1→H2:** S2.1..S2.5 CUMPLIDOS (fix por regla geométrica con T0 citado, contratos puros en
Core, frente R1 del 46 y `DivergenciasDeclaradasCatalogo` intactos, `<f>` intacta, build 0w).
**Veredicto H2→H3:** APROBADO (set-diff sin altas, Q1 verde, SHA intacto, tabla corrida-9→nueva
justificada celda a celda, corrida fresca `codigo=0`).

### 12.5 Cierre de los warnings del code-review (W-1/W-2/W-3/S-1) — 2026-10-17

El code-review independiente dictaminó **APROBADO-CON-OBSERVACIONES**; el Ingeniero ordenó cerrarlas.

- **W-1 (CRLF):** `CanalEfectivoDetalleR2R4.cs` y `CerosDocumentadosDetalleR2R4.cs` pasaron de LF puro a
  **CRLF** (139 y 124 líneas, `LF_solo = 0`; contenido intacto). Ningún otro `.cs` del diff en LF.
- **W-2 (red de seguridad del fail-fast):** `DetalleR2R4EscrituraTests` gana
  `DetalleR2_FilaSinteticaSinSlot_LanzaErrSinSlot_ConMensajeAccionable` (grupo sintético
  `Vlr Servicio` + cierre `C=Z D=Total` inyectado en la fuente de ASE2 → `ERR-SIN-SLOT` con hoja + ASE +
  nº de fila + firma + canal, sin salida parcial) y `DetalleR2_FilaDocumentada_NoFalla_Y_ConservaSuSlot`
  (contracara: la clave documentada del canal 2 no falla y el grupo conserva `E78/E79`).
- **W-3 (canal compuesto por segundo cierre):** el predicado puro compone `<canal>/<empresa>` con la
  empresa del primer cierre de empresa por debajo (R2: `E/ENEL`, `E/ENERBIT`, `A/EAB`,
  `9/NUEVO ESQUEMA`; columnas `[C]` canal + `[B]` empresa + `[A,B,C]` etiqueta propia). Las **filas de
  cierre de canal** también componen (si no, seguían empatando por orden). R4 no compone (su rótulo de
  grupo ya es la empresa). **Prueba exigida:** fuente con los dos grupos `E` de ASE5 en **orden
  invertido** → cada grupo recibe lo suyo, en dos niveles: el test del fixture (reordenando la lectura)
  y el **end-to-end con la fuente mutada** (`E386/E388` de ENEL, `E390/E392` de ENERBIT pese al orden).
  **Claves documentadas actualizadas** al canal compuesto (`9/NUEVO ESQUEMA`, `2/NUEVO ESQUEMA`,
  `A/EAB`); sin cierre de empresa debajo → canal simple (fallback declarado). Sin ambigüedad pendiente:
  no hubo STOP.
- **S-1 (evidencia + reconciliación):** `TestResults/p47-warnings-setdiff.txt` persistido; la **baja
  heredada del Plan 46** (`ProcesadorPeriodoTests.Ejecutar_Periodo2026072_ConLectorOracle_VeredictosValidacionesNoVacios`,
  ROJO→VERDE, 39→38) queda reconciliada con su causa citada (plans/46 §9.3.3) y el frente 47 **no
  introduce bajas nuevas**.

**Re-verificación:** build 0w; `p47-warnings.trx` **38 Failed / 999 Passed / 1037 Total / 0 Skipped** con
set-diff de fallos **0 altas / 0 bajas**; Q1 Passed; corrida fresca **codigo=0**; **corrida `p47-h2b` →
nueva = 0 celdas `values_changed`** (el canal compuesto no cambia ninguna cifra: el valor está en la
robustez al orden); frente a la corrida 9 se mantiene el cuadro del §12.3 (R2 192 + R4 120); `<f>`
13.000 celdas con fórmula **0 diferencias**; SHA intacto; cero commits; CRLF.
