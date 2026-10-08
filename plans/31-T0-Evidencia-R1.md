# Plan 31 — T0 Evidencia R1 (causa raíz del total de bloque R1 mal anclado)

> Investigacion de causa raiz (debug-agent). **Solo evidencia y opciones: cero cambios de codigo, tests ni Docs**
> (Docs se leyo; el unico artefacto creado es este informe en `plans/`).
> Metodo de lectura: parser **zip+XML BCL** (`System.IO.Compression` + `System.Xml`) sobre los `.xlsx`
> — **sin Excel ni COM**. Herramienta temporal: `%TEMP%\opencode\xlsx-dump.ps1` (no versionada).
>
> Archivos comparados (todos en disco, periodo 2026082 Q2 salvo donde se indica):
> - App: `Docs/Prueba2/Resultado/Resultado3/Remuneración 202608-2 Total.xlsx` (R3). Tambien R2 y root (misma formula).
> - Manual: `Docs/Prueba2/Resultado/Resultado Manual por el administrativo/Remuneracion 202608-2 Total_7721.xlsx`.
> - Base/plantilla agosto: `Docs/Prueba2/Plantilla_Remuneracion_2026082.xlsx` (= copia de la plantilla Q2 julio,
>   producida por `Herramientas/PreparadorBasePeriodo/Program.cs`; identica a `Docs/Prueba2/Plantilla_Remuneracion.xlsx`).
> - Canonico julio: `Docs/Plantilla_Remuneracion.xlsx`; golden julio: `Docs/Insumos/Remuneracion 202607-2 Total.xlsx`.
> - Fuentes: `Docs/Prueba2/Insumos/{1..5}-*/Recaudoporcomponente_*_16082026_*.xlsx`.
>
> Fecha: 2026-10-07. Hoja objeto: `Reporte Componentes R1`.

---

## 0. Causa raiz propuesta

**Una linea:** el espejo estructural R1 dimensiona cada bloque ASE por **conteo** y reancla las formulas con un
mapa de corrimiento de filas **calibrado a un borrado en el PIE del bloque**; como la fuente de agosto recorta
filas en la **CABEZA** del bloque, los datos quedan bien escritos (por orden) pero las formulas de total quedan
ancladas a las posiciones de julio, y nadie **re-deriva las filas-ancla por firma** (ni agrega/quita terminos).

Encadenamiento exacto (archivo:linea):

1. **Plantilla base agosto.** `Docs/Prueba2/Plantilla_Remuneracion_2026082.xlsx`, hoja `Reporte Componentes R1`,
   bloque ASE1: `F53 = F32+F48+F12-L12-L32` y `G53 = G32+G12`. Las anclas 12/32/48 son las filas `Mes/Total`
   de la geometria de julio. [evidencia §1.2]
2. **Mutador.** `Remuneracion.Infrastructure/Excel/OpenXmlEspejoR1Mutador.cs`, `AjustarBloque` (`:141-219`):
   dimensiona por `delta = filasFuente.Count - dataRows.Count` (`:157`); con `delta<0` borra las `m` ultimas
   filas de datos inmediatamente antes de la fila Total (`:183-199`) y reancla con
   `primeraBorrada = totalRowIdx - m` (`:210`). El mapa es
   `r >= totalRowIdx ? r-m : r >= primeraBorrada ? primeraBorrada : r` ⇒ **solo las referencias con fila
   >= primeraBorrada se desplazan**.
3. **Geometria de agosto.** La fuente ASE1 trae su zona superior 3 filas mas corta que el canonico:
   canonico filas 6..13 = [VlrServicio, D=E, C=ENEL, VlrServicio, D=O, C=OCCIDENTE, Mes, AFaseo];
   fuente filas 5..12 = [VlrServicio, D=O, C=OCCIDENTE, Mes, AFaseo, VlrServicio, D=E, C=ENEL].
   Todo el bloque se corre 3 filas hacia arriba: `Mes` 12->9, 32->29, 48->45; `Total` 50->47;
   `TOTAL OPORTUNO` 53->50. El mutador **no alinea por firma**: escribe la secuencia fuente **por orden**
   desde la primera fila de datos (`EscribirValores`, `:299-336`), asi que los DATOS quedan en las filas
   correctas, pero el mapa de reanclaje sigue siendo el del borrado inferior.
4. **Resultado observable.** App `F50 = F32+F47+F12-L12-L32` y `G50 = G32+G12` (solo 48->47).
   Manual (Excel reanclo al borrar filas arriba) `F50 = F29+F45+F9-L9-L29` y `G50 = G29+G9`.
   Los DATOS son identicos fila a fila (§1.3): la fila 29 tiene el `Mes` grande (F29=17.977.519.904,88) y la
   32 el total pequeno (F32=21.200,58) — la app suma la 32 y la 12; el manual la 29 y la 9.
5. **Cascada.** `CONSOLIDADO_TOTAL RECAUDO!D9 = 'Reporte Componentes R1'!F50` (el desplazamiento 53->50 si es
   correcto en la formula); al recalcular en Excel (`fullCalcOnLoad`, Plan 28) `D9` hereda el TOT_OPT mal
   anclado y las columnas de componentes se ven diminutas. El `DetRetri2026082` (literal de dominio) queda
   entonces inconsistente con el recalculo.

**Por que julio cerro (P3).** El golden `Remuneracion 202607-2 Total.xlsx` tiene la MISMA geometria que la
plantilla (TOT_OPT ASE1 en F53; `Mes` en 12/32/48) ⇒ no hubo borrado/insercion (delta 0 efectivo) y la formula
del canonico quedo valida. Agosto es el primer periodo con deriva de geometria en la CABEZA y delta<0.

**Hallazgo agravante (ASE5, §1.5):** la fuente de agosto trae **3 filas `Mes`** en ASE5 mientras la plantilla
(julio) declara **2 terminos** (`F558=F552+F531-L531`). Un simple corrimiento de filas NO puede corregir un
cambio de numero de terminos: la formula debe **recomponerse desde los roles de la fuente**.

---

## 1. Evidencia

### 1.1 Metodo (reproducible, sin Excel/COM)

`xlsx-dump.ps1 -Path <xlsx> -Sheet <hoja> [-From f -To f] [-OnlyCols ...]` abre el `.xlsx` como ZIP, resuelve
`xl/workbook.xml` + `xl/_rels/workbook.xml.rels`, lee `xl/sharedStrings.xml` y vuelca por celda
`ref / tipo / texto` (formula `<f>` con su texto y `<v>` cache; strings resueltos).

### 1.2 Formulas del total del bloque (misma hoja, mismo bloque)

| Celda | Plantilla base agosto | App R3 | Manual administrativo |
|---|---|---|---|
| TOT_OPT ASE1 | `F53 = F32+F48+F12-L12-L32` (v=0) | `F50 = F32+F47+F12-L12-L32` (v=0) | `F50 = F29+F45+F9-L9-L29` (v=18298670992.96) |
| G (TDF) ASE1 | `G53 = G32+G12` | `G50 = G32+G12` | `G50 = G29+G9` (v=975168318.98) |
| EXTEMP ASE1 | `F55 = F37+F17-L17` | `F52 = F37+F17-L17` | `F52 = F34+F14-L14` (v=63570) |

La App conserva la **estructura** de julio (32/48/12 con 48->47) y el Manual usa las filas nuevas (29/45/9).
Datos de la App y del Manual en las filas 6..47 son identicos (valores); solo difieren los `<f>`.

### 1.3 Los DATOS son identicos App vs Manual (prueba de que es anclaje, no lectura)

Valores F (columna "Total"), filas clave del bloque ASE1:

| Fila | Etiqueta | App R3 | Manual | Coincide |
|---|---|---|---:|---|
| 9 | B=Mes, C=Total | -9123.9 | -9123.9 | si |
| 12 | D=E, E=Total | 42369.42 | 42369.42 | si |
| 29 | B=Mes, C=Total | 17977519904.88 | 17977519904.88 | si |
| 32 | D=E, E=Total | 21200.58 | 21200.58 | si |
| 45 | B=Mes, C=Total | 358185030.54 | 358185030.54 | si |
| 47 | A=Total | 18335759381.52 | 18335759381.52 | si |

Las etiquetas A-E y los valores numericos de ambas salidas coinciden celda a celda en todo el bloque;
la unica diferencia es a que filas apunta la formula del total. La App usa etiquetas `inlineStr` (`is`) y el
Manual `sharedString` (`s`): diferencia de representacion, no de contenido.

### 1.4 Geometria: fuente de agosto vs canonico (mismo periodo, mismo bloque)

Fuente `.../1-Promoambiental/Recaudoporcomponente_*_16082026_*.xlsx` `Sheet1` (filas 5..46, 42 filas de bloque):

```
5  E=Vlr Servicio        (la canonica tiene ademas, antes de D=O: D=E, C=ENEL, E=VlrServicio)  <-- 3 filas menos
6  D=O, E=Total
7  C=OCCIDENTE, D=Total
8  B=Mes, C=Total        (Mes0)
...
28 B=Mes, C=Total        (Mes1, F=17.977.519.904,88)
...
44 B=Mes, C=Total        (Mes2, F=358.185.030,54)
45 A=Subs/Cont, B=Total
46 A=Total
```

El destino escribe la secuencia por orden desde la fila 6: `Mes0`->9, `Mes1`->29, `Mes2`->45, `Total`->47.
Coincide con la fila real del Manual. Por eso los datos "cuadran": el bug es exclusivamente de anclaje de `<f>`.

### 1.5 El mismo patron en los 5 ASE (formulas TOT_OPT, misma fila en App y Manual)

| ASE | Plantilla (julio) | App R3 | Manual | Comentario |
|---|---|---|---|---|
| 1 | `F53 = F32+F48+F12-L12-L32` | `F50 = F32+F47+F12-L12-L32` | `F50 = F29+F45+F9-L9-L29` | 3 terminos; anclas viejas |
| 2 | `F206 = F132+F160-L132+F94-L94` | `F194 = F129+F150-L129+F91-L91` | `F194 = F125+F148-L125+F93-L93` | 3 terminos; anclas viejas |
| 3 | `F343 = F265+F281+F244-L244-L265` | `F325 = F253+F265+F232-L232-L253` | `F325 = F251+F263+F235-L235-L251` | 3 terminos; anclas viejas |
| 4 | `F468 = F422+F448+F384-L384-L422` | `F456 = F404+F430+F366-L366-L404` | `F456 = F407+F436+F366-L366-L407` | 3 terminos; anclas viejas |
| 5 | `F558 = F552+F531-L531` (2 term.) | `F554 = F540+F519-L519` (2 term.) | `F554 = F548+F525+F494-L494-L525` (3 term.) | **la fuente Agosto tiene 3 Mes; falta un termino** |

En ASE2-5 (App vs Manual) se verifica que la App referencia filas con valores pequenos/de otra fila
(p. ej. ASE2: la App usa `F91` = total `D=O` = -7651.97; el Manual usa `F93` = `Mes0` = -1.443.509,98).
En ASE5 el bloque de agosto tiene 3 `Mes` (filas 494/525/548) en App y Manual; la plantilla solo trae 2
terminos ⇒ el mapa de filas, aun corregido, seria insuficiente.

### 1.6 Lectura de codigo (causa)

- `OpenXmlEspejoR1Mutador.AjustarBloque` resuelve el bloque por `B=<ASE>` y `A='Total'` (`LocalizarBloque`,
  `:225-288`), calcula `delta` (`:157`), borra/inserta en el **borde Total** (`:161-211`) y escribe por orden.
- Reanclaje `debug`: `Reanclar` recorre TODA formula del workbook y desplaza el numero de fila con `mapRow`
  (`:568-642`, `ReescribirReferencias` `:648-709`). El `mapRow` de delta<0 (`:210`) es el del borrado inferior.
  No hay ninguna lectura de firmas A-E para anclar; el mutador no conoce los roles.
- El lector SI resuelve roles por firma (`FilaEspejoR1.EsMesTotal` = `B='Mes' && C='Total'`, `FilaEspejoR1.cs:77-79`)
  y el dominio calcula `totOpt = Sum(F(todas las Mes)) - Sum(L(Mes menos la ultima))`
  (`ExcelDataReaderWorkbookLeafInputReader.cs:1083-1086`). Ese agregado es **correcto** (ver §1.7), pero
  **nunca se usa para escribir/recomponer la formula del workbook**.
- El mapa de celdas Q2 congela direcciones destino de julio: `WorkbookLeafCellMapQ2.R1Q2EditablesPorAse[1]`
  = F12=Mes0/F32=Mes1/F48=Mes2 (`:71-108`) y `R1Q2ProtectedPorAse[1]` F53=`["F32","F48","F12","L12","L32"]`
  (`:115-143`). Con espejo activo el mapa R1 no se escribe (ver §2).

### 1.7 El dominio SI acierta (aisla el bug al anclaje del workbook)

- `DetRetri2026082` de la App `D9:D14` = `18378829331 / 21599709648 / 16369059896 / 8372092112 / 12137660178 / 76857351165`
  — **identicos** al Manual (literales de dominio `ROUND(CONSOLIDADO D104:D108)`).
- `CONSOLIDADO_TOTAL RECAUDO` App vs Manual: la formula `D9 = 'Reporte Componentes R1'!F50` y
  `D104 = D9+D28+D47+D66+D85` son **identicas** en texto; difieren solo en el cache `<v>` (App 0 por no
  recalcular; Manual 18378829331.13). El `F50` del Manual = `18298670992.96` = formula correcta; el de la App,
  recalculado, da `18335782951.52`. La App calcula bien en C# y escribe mal la formula del workbook.

---

## 2. Por que los gates no lo detectan (P4)

1. **Validacion de formulas protegidas SALTEADA en agosto.** `OpenXmlPlantillaWriter.GenerarWorkbook`
   calcula `espejoDesplazado = RequiereAjuste(...)` antes de mutar (`:277-281`) y, si es `true`, **no** ejecuta
   `ValidarFormulasProtegidasMultiAse` ni antes (`:283-286`) ni despues de escribir (`:330-339`). En agosto
   `espejoDesplazado=true` (delta<0) ⇒ no corre la validacion por direccion absoluta. Y con razon: la
   validacion Q2 (`ValidarFormulasProtegidasMultiAseQ2`, `:781-808`) aserta direcciones congeladas de julio
   (`WorkbookLeafCellMapQ2.R1Q2ProtectedPorAse` F53 con fragmento `F48`), que tras la mutacion ya no aplican
   (la celda visible se movio a F50 y el fragmento `F48` no existe). Resultado: **no hay ningun gate
   estructural sobre las formulas visibles de R1 en agosto**.
2. **Los gates de coherencia son dominio-a-dominio, no workbook-a-dominio.**
   `WorkbookLeafCoherence.ValidarContraFuentes/ValidarContraResultado/ValidarSigmaEmpresas`
   (`WorkbookLeafCoherence.cs:22-58, 85-113, 180-220`) comparan `leaf.R1.F25`, `R2`, `R4`, `AjustesSfT` y
   `DetRetriQ2` (los **agregados C#**) entre si; **ninguno** lee `R1!F50/G50` ni evalua la formula escrita.
   `ValidarContraResultado` valida `leaf.R1.F25` (Mes0) contra `consolidado.Extemp`, jamas el TOT_OPT visible.
3. **`ValidacionOracleReader` no mira R1.** (`ValidacionOracleReader.cs:28-69, 98-128`) lee caches `<v>` de
   `DetValiRetri D21/D29`, `VALIDACION_*` O/P y sub-bloques, y exige `<v>` presente (S-4). Pero (a) no lee
   `R1!F50`; (b) en la salida de la App esos caches estan en 0/stale por `fullCalcOnLoad` (sin recalcular), y
   (c) el desglose DetRetri-vs-R10 compara el literal de dominio (correcto) contra el R10, no contra la
   formula recalculada del workbook. La brecha es invisible a este oraculo.
4. **El comparador de regresion EXIME toda la hoja R1 en agosto.**
   `Remuneracion.IntegrationTests/ComparadorSalidaVsManualTests.cs:172-185` declara brecha esperada para
   `Reporte Componentes R1` (y CONSOLIDADO, etc.) cuando `esAgosto`, con el motivo
   *"H1/T0b: remesh geometrico del manual de agosto (referencias de fila/periodo)"*. El comparador
   (`ComparadorSalidaVsManual.cs:17-25`) **si** compara texto de formula (`f`-vs-`f`) y habria delatado
   `F50=F32+F47+F12` vs `F29+F45+F9`; pero la lista de brechas declaradas lo tapa.
   **La premisa de la exencion es falsa**: los DATOS/geometria de App y Manual son identicos (§1.3); lo unico
   que difiere son las formulas mal ancladas. La exencion convierte un defecto real en "ruido esperado".
5. **Causa documentada y no corregida.** `plans/21 - T0 Evidencia.md:336-346` (R-DELTA-NEGATIVO/W-3) declara
   el compromiso de reanclar a `primeraBorrada` y afirma *"las visibles (TOT_OPT/EXTEMP) viven debajo del
   bloque y solo se desplazan, no caen en el rango suprimido"*. Ese supuesto es incorrecto: las formulas
   visibles **referencian filas interiores** (`Mes`) del bloque, y en ASE5 el numero de terminos cambia.

---

## 3. Opciones de fix (NO implementadas) y recomendacion

### Opcion A — Reanclaje semantico en el mutador (alinear por firma)
Que `AjustarBloque` construya el mapa `fila_plantilla -> fila_nueva` por **alineacion de firmas A-E** entre las
filas del bloque plantilla y las filas de la fuente (LCS/greedy monotono), en vez del heurismo
`primeraBorrada`; pase ese mapa a `Reanclar`.
- Pro: ataca la causa raiz (el mapa); mantiene "la formula se compone mecanicamente"; sigue siendo composicion
  de geometria (no negocio); conserva intactas las refs externas (CONSOLIDADO usa la fila visible, que cambia
  igual).
- Contra: **no cubre el cambio de numero de terminos** (ASE5: 2->3). Aun con mapa perfecto, `F554` seguiria
  con 2 terminos y faltaria sumar el tercer `Mes`. Rework considerable en el mutador; sensibilidad a firmas
  repetidas.
- Veredicto: necesario pero **insuficiente**.

### Opcion B — Recomponer las formulas visibles R1 desde los roles de la fuente (por firma)
Tras dimensionar, escribir el **texto** de las celdas visibles protegidas (`TOT_OPT`, `EXTEMP`, y las de
empresa que el mapa declare) recomponiendo la suma a partir de los roles resueltos sobre el bloque ya
dimensionado (`FilasMes()` -> Mes0..N; `FilasAplicacion()`), replicando la estructura `Sigma F(mes) - Sigma
L(mes menos la ultima)` de `MapearR1Q2` (`:1083-1090`).
- Pro: unico enfoque que **cubre el cambio de cardinalidad de terminos** (ASE1-4 3 terminos, ASE5 3 en agosto);
  aprovecha la fuente unica de firmas del Core (`FilaEspejoR1`); el dominio ya usa exactamente ese agregado.
- Contra: toca `<f>` de forma mas amplia; hay que declarar el contrato (que celdas visibles se recomponen y con
  que plantilla de formula) y actualizar/relajar `R1Q2ProtectedPorAse` (hoy congelado a julio). Riesgo de
  divergir de la "forma" exacta del manual si el admin usa otra estructura.
- Veredicto: es el que resuelve el caso ASE5 sin ambiguedad.

### Opcion C — Gate de paridad workbook-vs-dominio (red, no fix)
Agregar una verificacion post-escritura (y/o al comparador) que lea las celdas visibles R1 del workbook y
confirme que (a) la formula referencia exactamente las filas de los roles resueltos por firma, o (b) el valor
evaluado (suma simple BCL) coincide con `leaf.R1.TotalOportunoEsperadoPorAse` en ±0.5. Quitar la exencion
ciega de `Reporte Componentes R1` en `ComparadorSalidaVsManualTests` para agosto (reemplazarla por
comparacion de texto de formula de los visibles R1).
- Pro: barato; convierte el fallo en fail-fast; cierra el agujero del comparador.
- Contra: no corrige la salida por si solo.
- Veredicto: **obligatorio** como red, independientemente de A/B.

### Recomendacion
**B (+C).** El caso ASE5 (2->3 terminos) demuestra que un mapeo de filas no alcanza: el anclaje de los totales
visibles R1 debe **recomponerse por firma** desde los roles de la fuente (es composicion de geometria/firma,
excepcion legitima al "no tocar formulas"), y el gate de paridad (C) debe volver a cubrir R1 en agosto para
que ninguna corrida futura vuelva a entregar un total mal anclado en silencio. Si se quiere minimizar el
alcance de tocar `<f>`, A puede shippear como primer paso para ASE1-4, pero ASE5 exige B de todos modos.

---

## 4. Angulos muertos / NEEDS_CONTEXT

1. **Numeros del encargo no localizados.** Los valores citados "Recaudo Base 21.946.014.844 vs 15.753.100.487"
   y "Remuneracion Final 21.756.828.856 vs 15.600.939.318" **no aparecen** como literal ni como cache en
   ninguno de los tres `.xlsx` (App R3, R2, root), ni coinciden con `CONSOLIDADO D104/D109` del Manual
   (18.378.829.331,13 / 76.857.351.164,97). Probablemente provienen de una pantalla/calculo externo de la UI
   o de otro archivo. La evidencia verificada (y reproducible) es el anclaje de `R1!F50/G50` y su cascada.
   NEEDS_CONTEXT: origen exacto de esos dos pares de cifras.
2. **Plantilla realmente usada por la corrida.** Alta confianza de que fue
   `Docs/Prueba2/Plantilla_Remuneracion_2026082.xlsx` (base canonica de agosto; unica con formula
   `F53=F32+F48+F12` que, reanclada, produce la `F50=F32+F47+F12` observada). No hay log de corrida que lo
   confirme (no se hallo `remuneracion_log.txt`). NEEDS_CONTEXT: confirmar la ruta de plantilla de la corrida R3.
3. **Forma exacta esperada de las visibles.** La Opcion B debe fijar la "forma canonica" de `TOT_OPT`/`EXTEMP`
   por ASE y periodo (los visibles de empresa tambien se mueven). El Manual es la mejor fuente (F50/F554, etc.),
   pero requiere un T0 que congele la plantilla de formula por rol antes de escribir.
4. **CF/DV (W-4).** No se inspeccionaron `conditionalFormatting`/`dataValidations` del bloque mutado (fuera de
   este encargo; ya declarado en `plans/21`). Podrian anclar a celdas afectadas.
5. **Otras hojas con refs a filas interiores de R1.** `REMUNERACION_*`, `GERENTES_*`, `VALIDACION_*`,
   `DetValiRetri` referencian visibles R1 (`M50/M52/...`) y, via CONSOLIDADO, el total. Si el total de R1
   cambia, el recalculo arrastra esas hojas; merece un barrido de consistencia post-fix (no ejecutado aqui).
6. **Mapa aritmetico exacto por ASE.** El numero de referencias reancladas por bloque depende de la tecnica
   apilada 5->1 (`AjustarEnWorkbook`, `:99-120`) y de los deltas por ASE; no se reprodujo termino a termino
   para ASE2-5 (solo se verifico la divergencia App-vs-Manual). No es necesario para la causa raiz.

---

## 5. Cadena de causa (3 lineas) y recomendacion

Cadena: plantilla julio `F53=F32+F48+F12` -> `OpenXmlEspejoR1Mutador.AjustarBloque` dimensiona por conteo y
borra en el PIE del bloque (`:183-211`) dejando solo refs `>= primeraBorrada` reancladas -> la fuente de agosto
recorta en la CABEZA (Mes 12/32/48 -> 9/29/45) y los datos se escriben por orden pero la formula queda
apuntando a filas viejas (ASE5 ademas pierde un termino) -> `CONSOLIDADO!D9='R1'!F50` recalcula el total
equivocado y las columnas de componentes se ven diminutas.

Recomendacion: recomponer por firma las formulas visibles de `Reporte Componentes R1` al dimensionar el espejo
(Opcion B) y reactivar/ajustar el gate de paridad workbook-vs-dominio sobre R1 en agosto (Opcion C).