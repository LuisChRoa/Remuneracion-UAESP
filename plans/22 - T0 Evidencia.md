# Plan 22 — T0 Evidencia (Fase 0, bloqueo agosto en `SALDOS POR NOTA`)

> Documento de evidencia del Discovery T0 de la HU-22 (HU-11/Q2, paso 2.5 del `ProcesadorPeriodo`).
> **Solo evidencia: cero cambios de código productivo o de tests.** Toda afirmación remite a disco
> (archivo/hoja/fila). Replica el formato/evidencia del `plans/21 - T0 Evidencia.md`.
>
> **Fecha:** 2026-10-02
> **Origen del blocker:** `Remuneracion.IntegrationTests/Regresion2026082Tests.cs` (L32-L37) — el PR4
> del Plan 21 dejó capturado por nombre que el flujo 5-ASE de agosto se detiene en el paso 2.5
> (`ExcelDataReaderWorkbookLeafInputReader.LeerSaldosNotas`), ANTES de escribir.
> **Insumos Q2-julio:** `Docs/Insumos/REMUNERACION 2026072/{1..5}-*/`
> (`SaldosaFavorAplicadosPorNotas_to_date01072026*_to_date31072026*.xlsx`, `RetribuciónNegativa_*`).
> **Insumos agosto:** `Docs/Prueba2/Insumos/{1..5}-*/` (mismos prefijos, rango `01082026`–`31082026`).
> **Destinos:** `Docs/Insumos/REMUNERACION 2026072/Plantilla_ Remuneracion 202607-2.xlsx` (plantilla
> en ceros) y los goldens `Remuneracion 202607-1 Total.xlsx` / `Remuneracion 202607-2 Total.xlsx`
> (cachés en `%TEMP%\opencode\`; el golden en `Docs` puede estar tomado por Excel — ver §9).
> **Q1-julio:** `Docs/Insumos/REMUNERACION 2026071` **NO trae** fuentes `SaldosaFavorAplicadosPorNotas`
> ni `RetribuciónNegativa` (verificado por listado; ver §2.2). Q1 = sin datos de la cadena 2.5.

---

## 0. Conclusión de una línea

La hoja destino `SALDOS POR NOTA` **NO es espejo estructural de la fuente**: es una **plantilla de
geometría FIJA por ASE** (5 o 6 filas de concepto + filas visibles formuladas + columna `Especiales`
insertada) que se puebla por **columna según TÍTULO** y **fila según etiqueta de concepto**; por eso
el patrón espejo del Plan 21 **no aplica** (no hay que mover filas ni reanclar referencias). El
blocker de agosto es que el mapa congelado `WorkbookLeafCellMapAjustesSfT.ConceptosSaldosPorAse[2]`
exige la fila `Vlr Intereses` para ASE2 y la fuente de agosto de ASE2 no la trae; la solución
correcta es **tolerancia de concepto opcional en el lector**, no mutar la geometría. `RETRIBUCION
NEGATIVA` **no bloquea agosto** (fuente vacía = 0 legítimo).

---

## 1. Causa raíz de agosto (evidencia dura)

El fail-fast de 2026082 (distinto del `ERR-VALIDACION` R1 que resolvió el Plan 21) vive en el paso
2.5 y es **reproducible desde disco**:

| Período | ASE2 fuente (`SaldosaFavorAplicadosPorNotas_*`) | Conceptos observados | `Vlr Intereses` | Resultado |
|---|---|---|---|---|
| Q2-julio | `...2-Lime\...2026849258831.xlsx` | 6 (R4..R9) | **presente (R5)** | OK (mapa casa) |
| agosto | `Docs\Prueba2\Insumos\2-Lime\...202691202714480.xlsx` | 5 (R4..R8) | **AUSENTE** | **fail-fast antes de escribir** |

Dump crudo del breaker (agosto ASE2, `Sheet1`):

```text
R1 | A=Recaudo Desde: 01/08/2026 Hasta: 31/08/2026
R3 | C=Total | D=Componente TDF | E=Componente TTL | F=Componente TVIAT | G=Aprovechamiento | H=CCSA Prest.Aprov. | I=Componente TCS | J=Componente TLU | K=Componente TBL | L=Componente TRT | M=CCSA Prest. No Aprov. | N=Deb/Cred
R4 | B=Vlr Servicio   | C=931073.72
R5 | A=Componente     | B=Total | C=931073.72        <-- falta "Vlr Intereses"
R6 | B=Subsidio(-)/Contribucion(+) | C=-424218.36
R7 | A=Subs/Cont      | B=Total | C=-424218.36
R8 | A=Total          | C=506855.36
```

Comparación con julio (misma ASE2): R4 `Vlr Servicio=448371.28`, **R5 `Vlr Intereses=172.61`**,
R6 `Componente=448543.89`, etc. → el período de agosto pierde una fila de concepto. **La ocurrencia
congelada es el defecto, no el dato.**

---

## 2. T0a — Dumps fuente↔destino `SaldosaFavorAplicadosPorNotas` ↔ `SALDOS POR NOTA`

Método: lectura OpenXML de la fuente (`Sheet1`) como secuencia de filas etiquetadas (col A/B) y del
bloque destino desde la fila del nombre del ASE (col B) hasta la fila `A='Total'`. Se leen los
headers reales (nunca posiciones fijas).

### 2.1 Firmas de columnas reales (headers leídos de disco)

| Origen | C | D | E | F | G | H | I | J | K | L | M | N | O |
|---|---|---|---|---|---|---|---|---|---|---|---|---|---|
| **Fuente** Q2 y agosto (todos los ASE) | Total | Componente TDF | Componente TTL | Componente TVIAT | Aprovechamiento | CCSA Prest.Aprov. | Componente TCS | Componente TLU | Componente TBL | Componente TRT | CCSA Prest. No Aprov. | Deb/Cred (opcional) | — |
| **Plantilla/golden** (`SALDOS POR NOTA`) | Total | Componente TDF | Componente TTL | Componente TVIAT | Aprovechamiento | CCSA Prest.Aprov. | **Especiales** | Componente TCS | Componente TLU | Componente TBL | Componente TRT | CCSA Prest. No Aprov. | Deb/Cred |

Hechos de la firma (todos verificados en los 10 dumps Q2/agosto):

- **La fuente NO trae columna `Especiales`** (ningún ASE, ningún período). El template inserta
  `Especiales` en **I** y corre las componentes: `Componente TCS` (fuente I) → template **J**;
  `CCSA Prest. No Aprov.` (fuente M) → template **N**; `Deb/Cred` (fuente N) → template **O**.
  ⇒ el mapeo **solo puede ser por TÍTULO** (posicional fallaría en I..N).
- `Deb/Cred` (col N) es **opcional y variable por período**: Q2-julio la trae **solo ASE2**; agosto
  la traen **ASE1, ASE2 y ASE4**. Se mapea a `O`; ausente = 0 (patrón ya implementado).
- La fuente lleva **una fila de fecha (R1) + R2 vacía + headers en R3**; el template tiene **headers
  en R2 y no tiene fila de fecha**. Los bloques **no están alineados por posición**.

### 2.2 Conteos de filas de concepto y delta destino−fuente

Filas de concepto = filas etiquetadas del bloque (excluye la fila de fecha y la de headers).
Destino = rango fijo de la plantilla (`WorkbookLeafCellMapAjustesSfT.FilasSaldosNotasPorAse`).

**Fuente Q2-julio:**

| ASE | Conceptos fuente | Detalle |
|---|---:|---|
| 1 PROMOAMBIENTAL | 5 | Vlr Servicio, Componente, Subsidio, Subs/Cont, Total |
| 2 LIME | 6 | Vlr Servicio, **Vlr Intereses**, Componente, Subsidio, Subs/Cont, Total |
| 3 CIUDAD LIMPIA | 5 | Vlr Servicio, Componente, Subsidio, Subs/Cont, Total |
| 4 BOGOTA LIMPIA | 6 | Vlr Servicio, **Vlr Intereses**, Componente, Subsidio, Subs/Cont, Total |
| 5 AREA LIMPIA | 0 | **solo R1** (fecha); sin headers ni datos |

**Fuente agosto:**

| ASE | Conceptos fuente | Detalle |
|---|---:|---|
| 1 PROMOAMBIENTAL | 5 | Vlr Servicio, Componente, Subsidio, Subs/Cont, Total |
| 2 LIME | 5 | Vlr Servicio, Componente, Subsidio, Subs/Cont, Total — **SIN `Vlr Intereses`** |
| 3 CIUDAD LIMPIA | 5 | Vlr Servicio, Componente, Subsidio, Subs/Cont, Total |
| 4 BOGOTA LIMPIA | 6 | Vlr Servicio, **Vlr Intereses**, Componente, Subsidio, Subs/Cont, Total |
| 5 AREA LIMPIA | 0 | **solo R1** (fecha) |

**Destino (plantilla Q2 y golden Q2, geometría idéntica):**

| ASE | Filas plantilla | `Vlr Intereses` fija | Delta destino−fuente (Q2) | Delta destino−fuente (agosto) |
|---|---:|---:|---:|---:|
| 1 | 3..7 (5) | no | 0 | 0 |
| 2 | 15..20 (6) | **sí (fila 16)** | 0 | **+1 (slot sin dato)** |
| 3 | 28..32 (5) | no | 0 | 0 |
| 4 | 40..45 (6) | sí (fila 41) | 0 | 0 |
| 5 | 53..58 (6) | sí (fila 54) | +6 (fuente vacía) | +6 (fuente vacía) |

**Destino Q1 golden** (`Remuneracion 202607-1 Total.xlsx`, `SALDOS POR NOTA`): **6 filas de concepto
en los 5 ASE** (todas con `Vlr Intereses`), valores 0. Es una geometría **distinta** de la de Q2 y
de la plantilla. Q1 no tiene fuentes de la cadena y no invoca el lector (el paso 2.5 solo corre si
`Periodo.NumeroQuincena == 2`, `ProcesadorPeriodo.cs:196`), por lo que **Q1 queda fuera** del
problema; se documenta para mostrar que **la geometría la define el archivo plantilla, no la fuente**.

**Lectura T0a:** el destino **no reproduce la cardinalidad de la fuente**; el delta de agosto en ASE2
(`+1`) es el síntoma exacto del blocker. La fuente vacía de ASE5 (solo R1) no es fallo: es 0 legítimo
(early-return del lector, `ExcelDataReaderWorkbookLeafInputReader.cs:334-344`).

### 2.3 Valores cruzados (spot-check de que el destino refleja la fuente)

| ASE | Fuente Total (col C) | Golden Q2 (col C fila `Total`) | Cierra |
|---|---:|---:|---|
| 1 | 973693.46 (R8) | 973693.46 (R7) | sí |
| 2 | 216025.77 (R9) | 216025.77 (R20) | sí |
| 3 | 104231.83 (R8) | 104231.83 (R32) | sí |
| 4 | 35954.44 (R9) | 35954.44 (R45) | sí |
| 5 | (vacía) | 0 | sí |

---

## 3. T0b — Veredicto de geometría

**VEREDICTO: la hoja destino `SALDOS POR NOTA` NO es espejo estructural 1:1 de la fuente.** Es una
plantilla de **geometría fija** escrita por **celdas fijas** (`FilasSaldosNotasPorAse` +
`MapeoColumnasPorTitulo`). Tres pruebas independientes, todas desde disco:

1. **Columnas (offset estructural).** El template inserta `Especiales` en I y desplaza las
   componentes una posición (fuente I→template J). Un espejo posicional rompería; el mapeo real es
   **por título** (`WorkbookLeafCellMapAjustesSfT.MapeoColumnasPorTitulo`, 11 pares).
2. **Filas (bloques fijos con gaps + filas de fórmula).** Cada bloque destino tiene tamaño fijo por
   ASE (5/6) y añade filas que **no existen en la fuente**: la fila visible `Cn-In` (`C9=C7-I7`,
   `C22=C20-I20`, `C34=C32-I32`, `C47=C45-I45`, `C60=C58-I58`) y la fila `C10=C6` (etc.). La fuente
   no tiene esas filas ni la columna I que restan.
3. **Encabezado/prefijo.** Fuente = R1 fecha + R3 headers; template = R2 headers (sin fila de
   fecha). Los bloques no comparten origen posicional.

**Cómo escribe hoy el lector (fijo por mapa, no espejo):** `LeerSaldosNotas` localiza la fila de
headers dinámicamente (`BuscarFilaHeaders`), mapea **columnas por título** (`BuscarIndiceColumnaPorTitulo`),
y para cada concepto del mapa **fijo** por `Ase.Id` (`ObtenerConceptosSaldos`) busca la fila en la
fuente (`BuscarFilaConcepto`) y escribe en la **fila fija** del template
(`ExcelDataReaderWorkbookLeafInputReader.cs:321-414`). Lo único defectuoso es que **todo concepto del
mapa es obligatorio**.

---

## 4. T0c — Inventario de fórmulas/referencias de la cadena

Conteo OpenXML sobre el golden Q2 (`%TEMP%\opencode\Remuneracion 202607-2 Total.xlsx`, 40 hojas);
la plantilla en ceros tiene **exactamente los mismos anclajes y conteos**.

### 4.1 Fórmulas internas (nodos de la cadena)

| Hoja | Fórmulas totales | Masters (`<f>` con texto) | Anclajes representativos |
|---|---:|---:|---|
| `SALDOS POR NOTA` | 75 | 19 | `C9=C7-I7`; `D9=D4`; `E9:P9=E4`; `O9=O4`; `C10=C6` (bloques: 9/22/34/47/60 y 10/23/35/48/61) |
| `RETRIBUCION NEGATIVA` | 70 | 15 | `C10=C8-I8`; `D10=D5`; `E10:O10=E5`; `C11=C7` (bloques: 10/23/36/49/61 y 11/24/37/50/62) |
| `AJUSTES - SF-T` | 462 | 287 | puente (ver §4.2) |

Las filas de concepto (`15..20`, etc.) son **valores** que escribe el writer (T0b/T0-0.7); las
fórmulas viven en la fila visible y en los totales, y referencian la fila `Total`/`Componente` del
bloque. **Ninguna fórmula referencia `Vlr Intereses`** (fila 16/41/54): es una fila de entrada pura.

### 4.2 Puente `AJUSTES - SF-T` (dependencia directa de las dos hojas)

| Celda AJUSTES-SF-T | Fórmula real | Apunta a |
|---|---|---|
| `D9:D13` | `'SALDOS POR NOTA'!C9/C22/C34/C47/C60` | visible `Cn-In` de SALDOS |
| `E9:I13`, `U9:Y13`, `AD9` | `'SALDOS POR NOTA'!D/E/F/G/H/I/J/N/O/P…` | fila `Componente` (17/29/42/55) y `C10/C23/…` |
| `D28:D32` | `'RETRIBUCION NEGATIVA'!C10/C23/C36/C49/C61` | visible `Cn-In` de RETRIBUCION |
| `D14` / `D33` | `SUM(D9:D13)` / `SUM(D28:D32)` | subtotales de sección |
| `D47:D51` | `D9+D28`, `D10+D29`, … (mismo patrón por ASE) | composición por ASE |
| `D52` | `SUM(D47:D51)` | total |

### 4.3 Referencias externas aguas abajo (scanner de fórmulas)

| Referenciada | Celdas-fórmula | Hojas |
|---|---:|---|
| `'SALDOS POR NOTA'` | 70 | `AJUSTES - SF-T` (1 hoja) |
| `'RETRIBUCION NEGATIVA'` | 70 | `AJUSTES - SF-T` (1 hoja) |
| `'AJUSTES - SF-T'` | 292 | `CONSOLIDADO_TOTAL RECAUDO`=75, `DetValiRetri2026072`=60, `REMUNERACION_ENEL`=54, `DetRetri2026072`=45, `REMUNERACION Reciprocidad EAAB`=31, `REMUNERACION_ENERBIT`=27 |

Cadena `CONSOLIDADO` (verificada celda por celda):

```text
CONSOLIDADO D85 = 'AJUSTES - SF-T'!D47
CONSOLIDADO D86 = 'AJUSTES - SF-T'!D48
CONSOLIDADO D87 = 'AJUSTES - SF-T'!D49
CONSOLIDADO D88 = 'AJUSTES - SF-T'!D50
CONSOLIDADO D89 = 'AJUSTES - SF-T'!D51
```

### 4.4 Mapas/celdas absolutas que quedarían en riesgo **si** se intentara mutar (hipótesis rechazada)

Si se forzara un espejo (redimensionar bloques de SALDOS/RETRIBUCION a la cardinalidad de la fuente),
habría que reanclar: las filas visibles `C9/C22/C34/C47/C60` y `C10/C23/…`; los **70** refs de SALDOS
y **70** de RETRIBUCION en `AJUSTES - SF-T`; `D9:D13/D28:D32/D47:D51/D52`; y los **292** refs aguas
abajo. Además, los mapas que fijan filas absolutas se invalidarían: `FilasSaldosNotasPorAse`,
`FilaTotalSaldosPorAse`, `ConceptosSaldosPorAse`, `FilasRetribucionNegativaPorAse`,
`FilaTotalRetribucionPorAse`, `ConceptosRetribucionPorAse`, `ColumnaTemplateDebCred` y el arreglo
`Protegidas` (incl. `SALDOS POR NOTA`/`RETRIBUCION NEGATIVA`/`AJUSTES - SF-T`/`CONSOLIDADO D85:D89`/
`INTERVENTORIA`/`ANT EXT-REV`). **Ese costo (≈432 refs + 7 mapas) es el argumento duro para NO mutar.**

---

## 5. T0d — Causa raíz exacta + `RETRIBUCION NEGATIVA` agosto

### 5.1 Traza del fail-fast SALDOS ASE2

| Paso | Archivo:línea | Hecho |
|---|---|---|
| 1 | `ProcesadorPeriodo.cs:196-215` | En Q2 (agosto) se invoca el paso 2.5: `_leafReader.LeerSaldosNotas(ase, rutaSaldosNotas)` (línea 212) |
| 2 | `WorkbookLeafCellMapAjustesSfT.cs:124-128` | `ConceptosSaldosPorAse[2]` exige `(16, "Vlr Intereses")` |
| 3 | `ExcelDataReaderWorkbookLeafInputReader.cs:366,370-377` | Para `"Vlr Intereses"`: `BuscarFilaConcepto` → **-1** → `throw CalculoInvalidoException("ASE 2: no se encontró la fila de concepto 'Vlr Intereses' en la fuente SALDOS POR NOTA.")` |
| 4 | `Regresion2026082Tests.cs:41,87` | El test captura ese mensaje por nombre → confirma que el flujo muere **antes** de `GenerarWorkbook` (`ProcesadorPeriodo.cs:305`) |

No hay ninguna otra ocurrencia del mensaje en código/tests (grep: solo `ExcelDataReaderWorkbookLeafInputReader.cs:376`
y `Regresion2026082Tests.cs:41`). La causa es **la lista de conceptos congelada**, no un header
faltante (los headers obligatorios SÍ están: dump §2.1).

### 5.2 `RETRIBUCION NEGATIVA`: **NO bloquea agosto**

Las 5 fuentes Q2-julio y las 5 de agosto traen **solo la fila R1 (fecha)** (verificado por roweado
completo y por listado de hojas: `Sheet1` única). Por lo tanto:

- `BuscarFilaHeaders` devuelve -1 → **early-return `Total=0`** (`ExcelDataReaderWorkbookLeafInputReader.cs:428-439`).
- Nunca se llega al `throw` de concepto (`...:464-468`).
- El golden confirma 0: `AJUSTES - SF-T` D28:D32 = 0.

Conclusión: `RETRIBUCION NEGATIVA` de agosto **no** es un blocker; comparte el mismo **defecto
latente** (§5.3) pero no se ejercita.

### 5.3 Defecto latente espejado (declarado, no ejercitado)

`ConceptosRetribucionPorAse[1..4]` también exige `Vlr Intereses` (`WorkbookLeafCellMapAjustesSfT.cs:152-177`).
Si una futura fuente de RETRIBUCION trajera headers pero **no** la fila `Vlr Intereses`, el lector
fallaría igual que en SALDOS. Se documenta para que el fix cubra **ambos** readers con la misma regla.

---

## 6. T0e — Veredicto de diseño

### 6.1 Veredicto

**NO espejo.** El diseño actual ya es **"columna por título + fila por concepto acotada"**: el
destino es fijo y el defecto es que la lista de conceptos del mapa es **toda obligatoria**. La
solución correcta:

- **Mantener la geometría fija** de `SALDOS POR NOTA` / `RETRIBUCION NEGATIVA` (sin insert/delete,
  sin reanclaje). Las filas visibles y el puente `AJUSTES - SF-T` quedan intactos.
- **Tolerar la ausencia de `Vlr Intereses`**: concepto **opcional evidenciado** (varía por ASE y por
  período: presente en ASE2/4/5 Q2 y ASE4 agosto; ausente en ASE1/3 siempre y ASE2 en agosto). Si la
  fuente no lo trae → **0** (o no escribir; la celda del template ya es 0), nunca fail-fast.
- **Mantener obligatorios** los conceptos que aparecen en **todas** las fuentes no vacías (10/10):
  `Vlr Servicio`, `Componente`, `Subsidio(-)/Contribucion(+)`, `Subs/Cont`, `Total`. Ausencia de
  alguno de estos → fail-fast nombrando ASE + reporte (conducta actual, se conserva).
- La aritmética cierra: agosto ASE2 fuente ya trae `Componente = Vlr Servicio` (931073.72) y
  `Total = Componente + Subsidio` (506855.36); escribir `Vlr Intereses=0` en la fila 16 del template
  no altera ninguna fórmula (ninguna referencia a esa fila) y el visible `C22=C20-I20` = 506855.36.

Alternativa de implementación (equivalentes): (a) marcar `Opcional` las entradas `Vlr Intereses` en
`ConceptosSaldosPorAse`/`ConceptosRetribucionPorAse`; o (b) hacer el bucle **dirigido por la fuente**
(cada fila de concepto hallada se mapea a su fila fija; ausencia = 0). Cualquiera evita el fail-fast
sin tocar la geometría.

### 6.2 Invenciones prohibidas

- NO forzar el patrón espejo del Plan 21 en `SALDOS POR NOTA`/`RETRIBUCION NEGATIVA` (la geometría
  destino es fija; no calca R1).
- NO insertar/eliminar filas ni reanclar fórmulas de la cadena 2.5.
- NO declarar `Vlr Intereses` obligatoria (evidencia: es la única fila de concepto que varía).
- NO inventar el valor de intereses cuando falta (falta = 0; jamás un valor fabricado).
- NO escribir 0 silencioso para un concepto **core** ausente (debe fallar).
- NO cambiar el mapeo de columnas por título (la columna `Especiales` del template sigue en 0).
- NO tratar la fuente vacía (solo R1, ASE5) como fallo (es 0 legítimo).
- NO abrir un paso 2.5 en Q1 (Q1 no trae estas fuentes; `AjustesSfT = 0`).

### 6.3 Riesgos declarados

- **R-EXTRA-CONCEPTO (latente, sin evidencia):** si una fuente trajera un concepto extra que el
  template no tiene (dirección inversa), el lector actual lo **ignoraría** en silencio (itera el
  mapa, no la fuente). No se observa en Q2/agosto; queda como working list.
- **R-GEOMETRIA-Q1:** el golden Q1 usa 6 filas/ASE (distinta de Q2). No afecta porque Q1 no lee la
  cadena; si una HU futura abriera 2.5 en Q1, la geometría debería re-derivarse del archivo.
- **R-ORACULO-AGOSTO:** UAESP no entregó plantilla/salida de agosto; la verificación es estructural
  (sin golden ±0.5). El fix se certifica contra las fuentes reales y el R10, no contra un golden.

---

## 7. Veredictos T0a..T0e (resumen)

| ID | Entregable | Veredicto | Estado |
|---|---|---|---|
| T0a | Dumps fuente↔destino por ASE (Q2-julio, agosto, Q1) | Fuente 5/6/5/6/0 (Q2) y 5/5/5/6/0 (agosto); destino fijo 5/6/5/6/6; delta agosto ASE2 = +1; headers por título (template inserta `Especiales`); Q1 sin fuentes | DONE |
| T0b | Veredicto de geometría | **NO es espejo**: columnas por título (offset `Especiales`), filas fijas + filas de fórmula, prefijo de fecha distinto | DONE |
| T0c | Inventario de fórmulas/referencias | SALDOS=75, RETRI=70, AJUSTES=462; 70+70 refs en AJUSTES; 292 refs aguas abajo (CONSOLIDADO D85:D89 incluido); 7 mapas fijan filas absolutas | DONE |
| T0d | Causa raíz agosto ASE2 + RETRIBUCION | `ConceptosSaldosPorAse[2]` exige `Vlr Intereses`; agosto ASE2 no la trae → throw en reader (paso 2.5). RETRIBUCION **no** bloquea (fuente vacía=0); defecto latente compartido | DONE |
| T0e | Veredicto de diseño | **NO espejo**: tolerar `Vlr Intereses` opcional (ausente=0), core obligatorio, geometría fija sin reanclaje; invenciones prohibidas listadas | DONE |

No hay `NEEDS_CONTEXT` de T0a/T0b.

---

## 8. Decisiones que el Ingeniero debe aprobar (antes del plan de HU-22)

1. **Regla de opcionalidad:** ¿se acepta `Vlr Intereses` como el único concepto **opcional
   evidenciado** (ausente = 0), manteniendo core obligatorios? ¿O se adopta la regla general
   "ausencia de concepto = 0" (mayor tolerancia, menor fail-fast)? Recomendación T0: la primera.
2. **Cero mutación de geometría:** confirmar que la solución NO toca `FilasSaldosNotasPorAse` /
   `FilasRetribucionNegativaPorAse` ni reancla fórmulas (sin espejo, sin insert/delete).
3. **Alcance del fix:** ¿aplica a **ambos** readers (SALDOS y RETRIBUCION) por el defecto latente
   (§5.3), aunque RETRIBUCION no esté bloqueando hoy?
4. **Plantilla canónica:** el archivo vigente es `Docs/Insumos/REMUNERACION 2026072/Plantilla_
   Remuneracion 202607-2.xlsx` (SHA256 `95825422…0E0B8C`). El nombre citado en el comentario del mapa
   ("Plantilla 8 agos 2026 …") **ya no existe** en disco; confirmar que la canónica es la actual.
5. **Criterio del test 2026082:** al resolverse el blocker, ¿la rama end-to-end completa de
   `Regresion2026082Tests` debe asertar Δ = {-3,-9,-6,+6,+8}, invariantes R1 **y** `Vlr Intereses=0`
   en SALDOS ASE2 (sin golden ±0.5 de agosto, decisión de UAESP)?
6. **Q1 intacto:** confirmar que no se abre paso 2.5 en Q1 (sin fuentes; `AjustesSfT=0`).

---

## 9. Reproducibilidad (método, todo en disco)

Scripts de medición en `%TEMP%\opencode\` (no versionados; invocados con `-File` y `Get-Location`
para evitar el encoding de acentos en `Automatización`):

- `p22dump.ps1` — dump genérico OpenXML de filas por columnas (header-driven).
- `p22run.ps1` — dumpea las fuentes SALDOS/RETRIBUCION de Q1/Q2/agosto por ASE a `%TEMP%\opencode\p22`.
- `p22refs.ps1` — scanner de fórmulas: cuenta refs por hoja a SALDOS/RETRIBUCION/AJUSTES.
- `sheetsig.ps1`, `dumpwb.ps1` — herramientas existentes (dumps y fórmulas).

Comandos (PowerShell 5.1, cwd = raíz del repo):

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File "$env:TEMP\opencode\p22run.ps1"
$env:P22WB = "$env:TEMP\opencode\Remuneracion 202607-2 Total.xlsx"
powershell -NoProfile -ExecutionPolicy Bypass -File "$env:TEMP\opencode\p22refs.ps1"
```

Archivos crudos de evidencia en `%TEMP%\opencode\p22\`: `q2_src_saldos_ase{1..5}.txt`,
`aug_src_saldos_ase{1..5}.txt`, `q2_src_retri_ase{1..5}.txt`, `aug_src_retri_ase{1..5}.txt`,
`dst_gq2_saldos.txt`, `dst_gq1_saldos.txt`, `dst_tpl_saldos.txt`, `dst_gq2_retri.txt`,
`dst_tpl_retri.txt`, `f_SALDOS_POR_NOTA.txt`, `f_RETRIBUCION_NEGATIVA.txt`, `f_AJUSTES___SF_T.txt`,
`f_tpl_saldos.txt`, `f_tpl_retri.txt`.

Insumos usados (nombres exactos):
- Q2-julio: `Docs/Insumos/REMUNERACION 2026072/{1..5}-*/SaldosaFavorAplicadosPorNotas_to_date01072026ddMMyyyy_to_date31072026ddMMyyyy___*.xlsx`
  y `RetribuciónNegativa_to_date…_*.xlsx`.
- Agosto: `Docs/Prueba2/Insumos/{1..5}-*/SaldosaFavorAplicadosPorNotas_to_date01082026ddMMyyyy_to_date31082026ddMMyyyy___*.xlsx`
  y `RetribuciónNegativa_to_date…_*.xlsx`.
- Destinos: `Docs/Insumos/REMUNERACION 2026072/Plantilla_ Remuneracion 202607-2.xlsx` (SHA256
  `95825422B32FBE9E2B60F35C9638E492455FB98CFA182976B657FAC3520E0B8C`);
  `Docs/Insumos/Remuneracion 202607-2 Total.xlsx` (SHA256 `584310105AC7223CE840833E3CB64E26F95B61907C9ECD959A0A632EEC21DCB1`);
  caché `%TEMP%\opencode\Remuneracion 202607-2 Total.xlsx` (SHA256 `33A03BA9F302DEFC71BD6FF2820F711D56C149883B094473F4E073C11301385C`,
  A/B/C idénticos al de `Docs`); `%TEMP%\opencode\Remuneracion 202607-1 Total.xlsx` (SHA256
  `0B090E9C4851D86DAC534F2965E4A38393CC5BC09350C2A824FD48F7931C096F`).
