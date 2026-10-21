# T0 — EXTEMPORÁNEO ASE5 (Área Limpia) no llega al CONSOLIDADO

**Plan:** 49 · **Tipo:** T0 READ-ONLY de causa raíz · **Fecha:** 2026-10-19
**Período analizado:** 2026082 (agosto, 2.ª quincena) · **Modo:** 5 ASE, plantilla v2
**Autor:** debug-agent (subagente de `remuneracion-architect`)

> Restricción cumplida: **no se modificó ningún `.xlsx`, plantilla, código ni test**. Todo el
> análisis es lectura de bytes (ZIP + XML) con scripts read-only en `%TEMP%\opencode\`
> (`p49-tailcmp.ps1`, `p49-fmtdiff.ps1`, `p49-refscan.ps1`) y los utilitarios previos
> (`dumpxlsx.ps1`). No se ejecutó la suite. Hash de la plantilla v2 verificado al inicio:
> `FF5CF7A819110A1E1E2CA7BDAD553DD7A25141ADD472B1904BD496F082642632` (coincide con el congelado).

---

## 1. Resumen ejecutivo (veredictos)

| Pregunta | Veredicto |
|---|---|
| **Q1** ¿Por qué no se pega el bloque EXTEMPORÁNEO del ASE5? | **(a) + defecto de plantilla.** La fila-ancla `R1!F682` (y su gemela `F702`) es un **literal 0 sin `<f>`** en la v2 — y ya lo era en la plantilla oficial legacy. Está **fuera** del contrato de pegado (`RangoDatosAse(5) = 603..676`), por diseño: es una fila-resumen, no un dato de fuente. El escritor **no** la escribe (no debe) y la plantilla **no la calcula** (no puede). **(b) es FALSO**: el cálculo y la lectura son correctos. **(c) es FALSO**: no se escribe en otra celda; `F628`/`F659` están probados huérfanos. |
| **Q2** ¿El Plan 48 tuvo un falso verde en julio-2? | **SÍ.** La corrida promovida de julio-2 tiene `CONSOLIDADO!D51 = 0` frente a `441.849` del manual, y el gate `D34/F34/O25` **no lo detectó** porque la misma pérdida entra en los dos operandos y **cancela**. El propio T0 del Plan 48 **observó** las 9 fórmulas faltantes en 682/692/702 y las **congeló como esperadas** ("el delta es del manual, no de la base"). |
| **Q3** Barrido del patrón | **Solo en `Reporte Componentes R1`, zona de anclas del ASE5: 77 celdas** (filas 682/683, 692/693, 702/703 × columnas F..AO). En `Rem. Anticipos R2`, `Reversion Pagos R4`, `CONSOLIDADO_TOTAL RECAUDO`, `VALIDACION_*`, `REMUNERACION_*`, `AJUSTES - SF-T`, `SALDOS POR NOTA`, `RETRIBUCION NEGATIVA`, `INTERVENTORIA` y `GERENTES_*`: **0 celdas**. |

**Hallazgo adicional (no pedido, pero imprescindible para el plan):** el `12.462,26` que reporta el
gate **no es** el EXTEMPORÁNEO. Es un **segundo defecto, independiente**, de la plantilla v2 en
`Rem. Anticipos R2` del ASE4 (grupo `A/EAB`, `−12.462,26`). Consecuencia de planificación:
**arreglar solo el EXTEMPORÁNEO NO cierra el gate**; el Δ seguiría siendo `12.462,26`.

---

## 2. Evidencia base (hechos verificados)

### 2.1 Los tres números

| Celda (VALIDACION_TOTAL) | Nuestra salida | Manual administrativo | Δ |
|---|---|---|---|
| `D34` (`D33='CONSOLIDADO_TOTAL RECAUDO'!D109`) | `76.856.113.817,23` | `76.857.351.164,97` | `1.237.347,74` |
| `D109` (lado CONSOLIDADO, operando derecho de D34) | `76.856.101.354,97` | `76.857.351.164,97` | `1.249.810,00` |
| `F34`, `O25` | Δ idéntico | — | `12.462,26` |

Los dos lados deberían ser el **mismo** número (el manual cierra en `D34 = TRUE`). Nuestro lado
"empresas" (`D33`) pierde `1.237.347,74`; nuestro lado "ASE" (`D109`) pierde `1.249.810,00`. La
diferencia entre ambas pérdidas es **exactamente** el Δ del gate:

```text
1.249.810,00 (pérdida lado ASE)  −  pérdida lado empresas = 12.462,26
⇒ pérdida lado empresas = 1.237.347,74
```

### 2.2 Descomposición exacta del Δ (verificada por celdas)

| Componente | Valor | Lado ASE (CONSOLIDADO) | Lado empresas (REMUNERACION_*) | Efecto en `D34` |
|---|---|---|---|---|
| EXTEMPORÁNEO ASE5 (`R1!F682` / `R1!F702` = 0) | `1.249.810,00` | pierde | pierde | **cancela** (invisible) |
| `R2` ASE4 grupo `A/EAB` (`R2!E307 = −12.462,26`) | `−12.462,26` | **no** pierde | pierde | **rompe** (visible) |

---

## 3. Q1 — Cadena completa fuente → celda

### 3.1 Fuente (bytes)

`Docs/Prueba Agosto-2/Insumos/5-Área Limpia/Recaudoporcomponente_..._202691162043877.xlsx`, `Sheet1`:

| Fila | Contenido | `F` |
|---|---|---|
| 20 | `Aplicacion nuevos x reversion` / `C=Total` | `1666506.54` |
| 48 | `Aplicacion nuevos x reversion` / `C=Total` | `-416696.54` |
| 64 | `A=Total` (total del bloque) | `12117027175` |

`1.666.506,54 + (−416.696,54) = 1.249.810,00` → es el EXTEMPORÁNEO del ASE5.

### 3.2 Lectura y pegado: correctos (hecho verificado)

Salida `Docs/Prueba Agosto-2/Resultado/7/Remuneración 202608-2 Total (sin validar).xlsx`,
`xl/worksheets/sheet10.xml` (`Reporte Componentes R1`):

```xml
<x:c r="F629" s="710" t="n"><x:v>1666506.54</x:v></x:c>     ← se pegó
<x:c r="F660" s="710" t="n"><x:v>-416696.54</x:v></x:c>     ← se pegó
<x:c r="F676" s="90"  t="n"><x:v>12117027175</x:v></x:c>    ← total de bloque, se pegó
<x:c r="F682" s="234"><x:v>0</x:v></x:c>                    ← ancla EXTEMP: literal 0
<x:c r="F702" s="236"><x:v>0</x:v></x:c>                    ← ancla EXTEMP (OCCIDENTE): literal 0
```

Además, `DetRetri2026082!D13` (ASE5) = `12.137.660.178` en **ambos** libros → el cálculo de
dominio y la lectura de fuentes son correctos. Queda descartado el escenario (b).

### 3.3 Dónde se decide escribir u omitir (archivo:línea)

| Punto | Archivo:línea | Hecho |
|---|---|---|
| Rango de slots del ASE5 | `Remuneracion.Infrastructure/Excel/EscritorValoresSlotR1.cs:89-97` (`5 => (603, 676)`) | `F682`/`F702` quedan **fuera** del contrato |
| Selección de filas-eslot | `EscritorValoresSlotR1.cs:545-552` (`r.RowIndex >= inicio && <= fin`) | solo se leen/pegan filas 603..676 |
| Escritura `values-only` | `EscritorValoresSlotR1.cs:686-703` | escribe `<v>`; **jamás** `<f>`; si la celda destino es fórmula → `ERR-PLANTILLA` (`:694-699`) |
| Invocación del escritor | `OpenXmlPlantillaWriter.cs:290-302` | pega los 5 bloques; **nunca** toca la cola 677..710 |
| Consumidor del ancla | `WorkbookLeafCellMapQ2.cs:233` y `WorkbookLeafCellMapPorAse.cs:130` (`D51 ↔ F682`) | el CONSOLIDADO depende de esa celda |

### 3.4 La plantilla no la calcula (bytes)

Plantilla v2, `xl/worksheets/sheet10.xml`:

```xml
<c r="F97"  s="236"><f>F58+F31-L31</f><v>0</v></c>               ← ASE1: fórmula
<c r="F276" s="236"><f>F189+F145-L145</f><v>0</v></c>           ← ASE2: fórmula
<c r="F437" s="236"><f>F341+F314-L314</f><v>0</v></c>           ← ASE3: fórmula
<c r="F590" s="236"><f>F529+F485-L485</f><v>0</v></c>           ← ASE4: fórmula
<c r="F680" s="100"><f>F674+F648+F613-L613-L648</f><v>0</v></c> ← ASE5 OPORTUNO: fórmula
<c r="F682" s="234"><v>0</v></c>                                 ← ASE5 EXTEMP: LITERAL
<c r="F683" s="234"><v>0</v></c>
<c r="F702" s="236"><v>0</v></c>                                 ← ASE5 EXTEMP (OCCIDENTE): LITERAL
<c r="F703" s="236"><v>0</v></c>
```

El manual del administrativo, en la misma posición lógica, **sí** trae la fórmula:

- agosto (geometría legacy): `R1!F556 = F534+F506-L506` → `1.249.810`; `R1!F576 = F533+F505-L505` → `1.249.810`.
- julio (geometría v2): `R1!F682 = F660+F629-L629` → `441.849`; `R1!F702 = F659+F628-L628` → `441.849`
  (`Docs/Prueba Julio-2/Resultado/Resultado Administrativo con plantilla estandar v2/Formato Remuneracion 202607-2 Total.xlsx`).

**El defecto NO lo introdujo el Plan 48.** La plantilla oficial legacy
(`Docs/Plantilla estandarizada/Formato Remuneracion 2026xx-x Total.xlsx`) también trae
`<c r="F682" s="234"><v>0</v></c>` y `<c r="F702" s="236"><v>0</v></c>`. La v2 lo **heredó**.

### 3.5 Veredicto Q1

- **(a) correcto, con matiz**: no hay slot para la fila-ancla *y* la plantilla perdió la fórmula que
  la calculaba. La omisión del escritor es **correcta por doctrina** (values-only, sin inventar
  agregados); el error está en la **base**.
- **(b) descartado**: el dominio calcula `1.249.810` (lo prueba `DetRetri2026082!D13`, idéntico al manual).
- **(c) descartado**: `T0ResidualOccidenteAse5Tests.cs:61-84` prueba por bytes que `F628`/`F659` no
  son referenciadas por **ninguna** fórmula de **ninguna** hoja; no hay consumo alternativo.

---

## 4. Q2 — Julio-2: ¿hubo falso verde del Plan 48?

### 4.1 Hechos (bytes)

| Artefacto | `R1!F682` | `CONSOLIDADO!D51` |
|---|---|---|
| Manual v2 julio-2 (oráculo) | `<f>F660+F629-L629</f>` cache `441849` | `441849` |
| App `Resultado\5` (2026-10-14) | literal `0` | `0` |
| App `Resultado\6` (2026-10-15) | literal `0` | `0` |
| App `Resultado\8` (2026-10-17, promovida) | literal `0` | `0` |
| App `Resultado\9` (2026-10-17) | literal `0` | `0` |
| App `Resultado\10` (2026-10-19) | literal `0` | `0` |
| `%TEMP%\opencode\p48-h2-fresh2\Remuneración 202607-2 Total.xlsx` | literal `0` | `0` |

- `p48-h2-fresh2` (SHA256 `342CFE27…`) y `Resultado\10` (SHA256 `76FEB0AE…`) **no son
  byte-idénticos** (mismo tamaño 611.366 B, distinto contenedor/IDs): son dos copias de la misma
  familia de corrida, ambas **sin** el sufijo `" (sin validar)"`, es decir, **post-gate promovidas**
  (`NombresSalida.SufijoSinValidar`, `ProcesadorPeriodo.cs:97,400,540`).
- La corrida de julio-2 **no fue un verde limpio**: `Resultado\8\…- Divergencias declaradas.txt`
  reporta `D34/F34/O25` en FALSE con Δ `417.519,62` (otra causa: la 1.ª sección de OCCIDENTE ASE5)
  y cerró por la vía de **divergencias declaradas**. Esa entrada del catálogo está hoy **retirada**
  (`DivergenciasDeclaradasCatalogo.cs:92-123`).

### 4.2 Por qué el gate no vio los `441.849`

`D34 = (D33 = CONSOLIDADO!D109)` compara **dos derivaciones del mismo workbook**:

- lado derecho `D109` ← `D51` ← `R1!F682` (= 0);
- lado izquierdo `D33` ← `REMUNERACION_OCCIDENTE!D109` ← `D51` ← `R1!F702` (= 0).

La **misma** pérdida entra en los dos operandos y **cancela**. El Δ que reporta el gate es la
**asimetría** entre las dos pérdidas, no la pérdida. Esto está probado y documentado en el propio
fixture: `T0ResidualOccidenteAse5Tests.cs:102-113`
("`O25` = TRUE: el `24.329,38` entra en AMBOS lados y cancela").

### 4.3 El Plan 48 lo vio y lo congeló

`Remuneracion.IntegrationTests/AdopcionPlantillaV2T0Plan48Tests.cs:421-450`:

- `:421` "Las 77 nuevas se concentran en las filas 682/692/702 (+ 683/693/703), zona del residual."
- `:433-441` aserta las fórmulas **del manual**: `F682 = F660+F629-L629`, `F702 = F659+F628-L628`.
- `:446-450` "En la v2 esas celdas están vacías (sin fórmula): **el delta es del manual, no de la
  base**." + `Assert.False(v2.TieneFormula(r))` para `F682/F692/F702/...`.

Y `GeometriaPlantillaEstandarizadaT0Tests.cs:230-231` lo eleva a contrato:
"`F682` (EXTEMP ASE5) debe ser VALOR (no fórmula) en el contrato T0."

`DetRetriQ2Tests.cs:38-40` deja la constancia explícita:
"el template deja `F682` en VALOR 0 (sin fórmula); el agregado del reader … = `441849` = D51 golden".
**Se validó el valor de dominio y nunca se asertó la celda del workbook.** Ese es exactamente el
hueco que produjo el falso verde.

### 4.4 Veredicto Q2

**Sí, falso verde.** El Plan 48 certificó una salida cuyo `CONSOLIDADO!D51` es `0` en vez de
`441.849`. La clasificación "el delta es del manual, no de la base" es **incorrecta**: esas 77
celdas son **anclas consumidoras** (`CONSOLIDADO!D51/E51/F51`, `DetRetri!D36`, `DetValiRetri!D28`,
`REMUNERACION_OCCIDENTE!D51`, `REMUNERACION_ENEL!D51`), no cosmética del manual.

---

## 5. Q3 — Barrido del patrón

### 5.1 Método

1. Comparación **misma dirección** manual-v2-julio ↔ plantilla-v2 (misma geometría): lista completa
   de celdas con `<f>` en el manual y sin `<f>` en la v2 → **77 celdas**, todas en `Reporte Componentes R1`.
2. Comparación **por ancla de bloque** manual-agosto (legacy) ↔ v2 (`TOTAL|OPORTUNO` como ancla):
   74 celdas con pérdida de fórmula en agosto (las 3 restantes son `F692/G692/F693` del grupo ENEL,
   que en el manual legacy son literales `0`).

### 5.2 Lista (R1, zona de anclas del ASE5)

| Fila v2 | Celdas con fórmula en el manual y literal en la v2 | Valor del manual en agosto |
|---|---|---|
| **682** | `F..AO` (36 celdas) | `F=1.249.810,00` · `G=89.451,49` · `H=50.203,36` · `I=26.464,20` · `J=371.981,73` · `K=13.370,78` · `M=117.292,30` · `N=115.277,54` · `O=591.928,60` · `P=268.663,04` · `Q=21.816,48` · `R=57,02` · `L/S..AO=0` |
| **683** | `F` | `−416.696,54` |
| **692** | `F`, `G` | `0` (hueco **latente**: si algún período trae EXTEMP de ENEL, se pierde) |
| **693** | `F` | `0` (latente) |
| **702** | `F..AO` (36 celdas) | mismos valores que 682 (son la misma ancla, familia REMUNERACION) |
| **703** | `F` | `−416.696,54` |

**Consumidores de esas celdas** (verificado en la v2 por escaneo de fórmulas):

| Celda | Consumidores |
|---|---|
| `F682` | `CONSOLIDADO!D51`; `DetRetri!D36/K36/M36/O36`; `DetValiRetri!D28/L28/M28/N28/O28` |
| `G682`, `H682` | `CONSOLIDADO!E51`, `F51`; `DetReti/DetValiRetri` |
| `M682` | `CONSOLIDADO!U51`; `DetValiRetri!D28/O28` |
| `F683` | `CONSOLIDADO!AE51`; `DetValiRetri!L28/M28/N28` |
| `F702` | `REMUNERACION_OCCIDENTE!D51` (y `E51/F51/U51/AE51`) |
| `F692` | `REMUNERACION_ENEL!D51` |
| `F703` | `REMUNERACION_OCCIDENTE!AE51` |

### 5.3 Otras hojas

`p49-fmtdiff.ps1` (dirección a dirección, manual-v2-julio ↔ v2) sobre
`Rem. Anticipos R2`, `Reversion Pagos R4`, `CONSOLIDADO_TOTAL RECAUDO`, `VALIDACION_TOTAL`,
`VALIDACION_RECIP/ENEL/ENERBIT/OCCIDENTE/EAAB-CL`, `REMUNERACION_*`, `AJUSTES - SF-T`,
`SALDOS POR NOTA`, `RETRIBUCION NEGATIVA`, `INTERVENTORIA`, `GERENTES_TOTAL`:
**0 celdas con fórmula en el manual y literal en la v2.**

### 5.4 Valores-fuente que la app debe escribir y quedan en 0 (otra clase)

`CerosDocumentadosDetalleR2R4.cs:55-64` declara como "cero documentado" **3 filas del R2 del ASE4**
(grupo `A/EAB`, total `−12.462,26`). Verificado: la fuente
`4-Bogota Limpia/RerpoteDetalleSaldosaFavor…` filas 33-35 trae `−12.462,26`, y el valor **no
existe en ninguna celda** de nuestra salida (`grep "12462.26"` = 0 coincidencias en los 40 XML de
hoja). No es una pérdida de fórmula: es una **pérdida de fila por falta de slot**. Es el segundo
defecto (ver §6.2).

---

## 6. Causa raíz

### 6.1 Defecto primario (el que motiva el Plan 49)

> **La plantilla oficial (v2, heredada de la legacy) tiene las anclas-resumen `EXTEMPORANEO` del
> ASE5 como literal `0` en lugar de fórmula. El escritor values-only no las toca (están fuera de su
> contrato por diseño), así que el `EXTEMPORÁNEO` del ASE5 nunca llega ni a `CONSOLIDADO!D51` ni a
> `REMUNERACION_OCCIDENTE!D51`.**

Cadena causal completa:

```text
fuente R1 ASE5 (F20=1.666.506,54 · F48=−416.696,54)
  → ExcelDataReaderRecaudoReader (lectura OK)
  → RecaudoComponenteR1 / espejo (dominio OK: Extemp = 1.249.810)
  → EscritorValoresSlotR1 (pega F629 y F660, filas 603..676)
  → [FALTA: la fórmula-agregado F682 = F660+F629−L629 en la plantilla]
  → CONSOLIDADO!D51 = 0  y  REMUNERACION_OCCIDENTE!D51 = 0
  → D109 y D33 pierden 1.249.810 → el gate cancela esa pérdida (no la ve)
```

### 6.2 Defecto secundario (el que el gate sí reporta)

> **La plantilla v2 no reproduce la fórmula del ancla `RECIPROCIDAD/OPORTUNO` del R2 del ASE4.**
> El manual consume la fila `A/EAB` (`E307 = −12.462,26`) en `E363 = E317+E307+E293+E279-K279-K293`;
> la v2 usa `E387 = E341+E316+E296-K296-K316`, que **no** incluye esa fila (no existe como slot).
> `REMUNERACION_OCCIDENTE!D31` apunta a `E387` (v2) / `E363` (manual) → el lado "empresas" queda
> `12.462,26` por encima del lado "ASE".

Verificado: `12462.26` aparece en **exactamente 3 celdas** (`R2!E305/E306/E307`) del manual y en
**0 celdas** de toda nuestra salida; y la referencia
`REMUNERACION_OCCIDENTE!D31 = 'Rem. Anticipos R2'!E387` (v2) vs `!E363` (manual) está en los bytes.

### 6.3 Causa raíz de la **clase** de error (la que hay que cerrar)

> **El gate `D34/F34/O25` es una comprobación de consistencia interna (dos derivaciones del mismo
> workbook), no un oráculo absoluto. Cualquier pérdida común a ambas derivaciones es invisible por
> construcción.** El `1.249.810` (el error grande) no lo ve; el `12.462,26` (el pequeño) sí lo ve
> porque es asimétrico.

---

## 7. Alcance del defecto

| Dimensión | Alcance verificado |
|---|---|
| Hojas | `Reporte Componentes R1` (77 celdas-ancla) + `Rem. Anticipos R2` (3 filas del ASE4, sin slot) |
| ASE | Solo **ASE5** en R1 (ASE1-4 tienen fórmula); **ASE4** en R2 |
| Períodos | Julio-2 (2026072) y agosto-2 (2026082). Afecta a **cualquier** período con EXTEMPORÁNEO en el ASE5 |
| Plantillas | v2 **y** la oficial legacy (heredado) |
| Cifras | Ago-2: `1.249.810,00` (EXTEMP ASE5) · Jul-2: `441.849,00` |
| Hojas consumidoras aguas abajo | `CONSOLIDADO_TOTAL RECAUDO` (D/E/F/U/AE 51), `DetRetri`, `DetValiRetri`, `REMUNERACION_OCCIDENTE`, `REMUNERACION_ENEL` |
| Celdas del gate afectadas | `D34`, `F34`, `O25` (bloqueadas por el defecto secundario; ciegas al primario) |
| Lo que **no** está afectado | Lectura de fuentes, cálculo de dominio (`DetRetri` idéntico), colocación del interior R1 (F629/F660/F676 correctos), R4, AJUSTES, BCE, banco, DetRetri-Q2 escrito |

---

## 8. Propuesta de corrección (NO implementada)

### 8.1 Restricción dura

La plantilla v2 está congelada por hash y por T0. Cualquier cambio de bytes exige **nueva versión
(v3) + nueva huella de aceptación** y el reemplazo de los siguientes T0, que hoy **asertan el
defecto**:

- `GeometriaPlantillaEstandarizadaT0Tests.cs:104,230-231` (F682 debe ser VALOR).
- `T0ResidualOccidenteAse5Tests.cs:30-48` (F682/F702 sin fórmula, valor `0`).
- `AdopcionPlantillaV2T0Plan48Tests.cs:446-450` (v2 no debe traer fórmula en 682/692/702).
- `R1FuenteCanalConceptoT0Tests.cs:209-210`.
- `DetRetriQ2Tests.cs:38-40` (comentario y expectativa).

### 8.2 Opción A — Plantilla v3 (preferida para el defecto primario)

Restaurar en `Reporte Componentes R1` las fórmulas del manual:

| Celda | Fórmula |
|---|---|
| `F682` | `F660+F629-L629` |
| `G682` | `G629` |
| `H682` | `H629` |
| `F683` | `F660` |
| `F692` | `F652+F618-L618` |
| `G692` | `G618` |
| `F693` | `F652` |
| `F702` | `F659+F628-L628` |
| `G702` | `G628` |
| `H702` | `H628` |
| `F703` | `F659` |

(`I682..AO682` e `I702..AO702` son réplicas de `H629`/`H628` en el manual; incluirlas o no es
decisión de T0, pero **`M682` sí tiene consumidor** en `DetValiRetri!D28`.)

**Riesgos**
- Cambia la huella de la plantilla → hay que re-emitir `HuellaAceptacion` y actualizar los 5 T0.
- `ValidarFormulasProtegidas` debe incorporar estas celdas a "protegidas": con la fórmula
  restaurada, el escritor fallará si alguien las convierte en valor.
- Los cachés `<v>` de la plantilla quedan en `0`; el sellado/saneamiento actual ya contempla
  `fullCalcOnLoad`, y el resolutor del oráculo evalúa fórmulas (no caché), así que el gate
  funcionaría sin Excel.
- Requiere autorización formal porque toca el insumo "oficial".

### 8.3 Opción B — Anclas por valor desde la app (si la plantilla no se puede tocar)

Nuevo contrato `AnclasResumenR1` (ASE, ancla lógica) → celdas, escrito **values-only** con guardia
anti-fórmula (reutiliza `EscribirValor`):
`F682 ← ConsolidadoAse.Extemp`; `F702` ídem; `F683/F703` ← componente `Subsidio` de la sección.

**Riesgos**
- Duplica la fuente de verdad (plantilla + app) y no arregla el workbook si se genera por otra vía.
- `G/H` (disposición final, lixiviados) **no** están en `ConsolidadoAse`; habría que añadirlos al
  dominio o limitar el fix a `F` (dejando `E51/F51` en 0 → otro hueco).
- Cambia el carácter "fórmula" de celdas que hoy son VALOR; hay que ajustar el mapa de protegidas y
  el T0 de geometría.
- Exige un T0 nuevo de geometría de anclas por ASE (hoy no existe).

### 8.4 Opción C — Endurecer el gate (independiente, recomendada en paralelo)

Añadir una comprobación **absoluta** que no pueda cancelarse:

1. **Contra el dominio**: `CONSOLIDADO!D47:D51` == `ConsolidadoAse.Extemp` por ASE (el dominio ya lo
   tiene y hoy vale `1.249.810`). Falla con el defecto actual.
2. **Contra el interior de la propia hoja**: para cada ancla EXTEMP del ASE, comparar
   `R1!F682` con la suma resuelta de sus secciones (`F660+F629−L629`) desde `ResolutorValorValidacion`.
   Falla con el defecto actual y no depende de ningún oráculo externo.

**Riesgo**: el punto 2 codifica geometría de ancla (necesita su T0); el punto 1 requiere exponer
`Extemp` por ASE en el `LeafInputs` (ya existe en el consolidado) y decide si el gate corre en Q1
también.

### 8.5 Defecto secundario (R2 ASE4) — decisión de T0

Dos caminos mutuamente excluyentes:

- **C-1**: restaurar en v3 la fórmula `E-ancla` del manual (`E317+E307+E293+E279-K279-K293`) y
  habilitar el slot `A/EAB` en el bloque R2 del ASE4 (deja de ser "cero documentado").
- **C-2**: mantener el cero documentado y **declarar** la divergencia `12.462,26` con dueño UAESP.

**Importante para el plan**: sin resolver C, el gate seguirá fallando con `12.462,26` **aunque se
arregle el EXTEMPORÁNEO** (verificado aritméticamente: `D33` pasaría a `76.857.363.627,23` y `D109`
a `76.857.351.164,97`; Δ = `12.462,26`).

---

## 9. Tests que cubrirían el defecto

| # | Test | Estado esperado |
|---|---|---|
| 1 | **T0 plantilla v3**: `R1!F682/F683/F692/F693/F702/F703` (y `G/H`) son `<f>` con el texto del manual | Falla hoy (fuerza la corrección) |
| 2 | **Regresión 2026082**: `CONSOLIDADO!D51` resuelto = `1.249.810`; `REMUNERACION_OCCIDENTE!D51` = `1.249.810`; `DetRetri2026082!D13` = `12.137.660.178` (invariante) | Falla hoy |
| 3 | **Gate 2.7 con `Δ = 0`** para `D34/F34/O25` en 2026082 | Falla hoy (`12.462,26`) |
| 4 | **Anti-cancelación**: workbook sintético donde `D33` y `D109` pierden el mismo componente → el gate nuevo debe **fallar** (hoy pasa) | Falla hoy (el gate actual es verde) |
| 5 | **R2 ASE4**: `REMUNERACION_OCCIDENTE!D31` == valor del manual (o divergencia declarada vigente con Δ `12.462,26`) | Falla hoy |
| 6 | Actualizar los 5 T0 de §8.1 | — |

---

## 10. Anexo — trazabilidad de la evidencia

### 10.1 Archivos leídos (sin modificar)

| Rol | Ruta |
|---|---|
| Salida fallida | `Docs/Prueba Agosto-2/Resultado/7/Remuneración 202608-2 Total (sin validar).xlsx` |
| Oráculo agosto-2 | `Docs/Prueba Agosto-2/Resultado/Resultado Manual por el administrativo/Remuneracion 202608-2 Total_7721.xlsx` |
| Plantilla v2 (INTOCABLE) | `Docs/Plantilla estandarizada/v2/Formato Remuneracion 2026xx-x Total vf.xlsx` · SHA256 `FF5CF7A8…2632` |
| Plantilla oficial legacy | `Docs/Plantilla estandarizada/Formato Remuneracion 2026xx-x Total.xlsx` · SHA256 `3086AC31…5214` |
| Oráculo julio-2 (v2) | `Docs/Prueba Julio-2/Resultado/Resultado Administrativo con plantilla estandar v2/Formato Remuneracion 202607-2 Total.xlsx` |
| Corridas app julio-2 | `Docs/Prueba Julio-2/Resultado/{5,6,8,9,10}/Remuneración 202607-2 Total.xlsx` |
| Corrida app julio-2 (temp) | `%TEMP%\opencode\p48-h2-fresh2\Remuneración 202607-2 Total.xlsx` |
| Fuentes agosto-2 | `Docs/Prueba Agosto-2/Insumos/5-Área Limpia/Recaudoporcomponente_…xlsx`, `…/4-Bogota Limpia/RerpoteDetalleSaldosaFavor_…xlsx` |

### 10.2 Scripts read-only usados

| Script | Uso |
|---|---|
| `%TEMP%\opencode\p49-tailcmp.ps1` | Compara el "tail" de cada bloque ASE alineando por la fila ancla `TOTAL/OPORTUNO`; reporta `PERDIDA-F` / `F-DIST`. |
| `%TEMP%\opencode\p49-fmtdiff.ps1` | Comparación dirección-a-dirección (misma geometría) de presencia de fórmula. |
| `%TEMP%\opencode\p49-refscan.ps1` | Escaneo de fórmulas que referencian una celda (consumidores). |
| `%TEMP%\opencode\dumpxlsx.ps1` | Volcado de celdas/valores/fórmulas por hoja. |

### 10.3 Qué NO se pudo probar con bytes

- Que el `.xlsx` promovido de julio-2 haya sido generado por **una** corrida concreta: solo se
  prueba que los 6 artefactos de la era v2 tienen `F682 = 0` y que están promovidos (sin sufijo).
- La propagación **celda a celda** del `−12.462,26` por toda la cadena `VALIDACION_*` (se probó el
  eslabón decisivo `R2!E363/E387 → REMUNERACION_OCCIDENTE!D31`, y la aritmética exacta, pero no se
  recorrieron todas las hojas intermedias).
- El valor que Excel recalcularía al abrir nuestra salida (los cachés están en `0`; el gate usa el
  resolutor, no Excel).

---

## 11. Lección aprendida

Un T0 que **observa** una diferencia contra el manual y la clasifica como "delta del manual" puede
estar congelando un **defecto de base**. La prueba de que era un defecto no era la diferencia de
texto: era que **la celda es un consumidor** (`CONSOLIDADO!D51`). Regla propuesta para el Plan 49:
*ninguna celda que sea referenciada por otra hoja puede congelarse como "delta del manual" sin
demostrar que el valor no cambia el resultado del consumidor.*

Y, sobre el gate: **una igualdad entre dos derivaciones del mismo workbook nunca prueba corrección;
solo prueba consistencia.** Toda pérdida común (la más probable, porque ambas derivaciones comparten
las mismas anclas) es invisible.
