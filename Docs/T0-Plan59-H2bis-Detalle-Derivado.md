# T0 — Plan 59 (H2-bis): Detalle del período derivado

Read-only por bytes sobre la plantilla **v3** (`Docs/Plantilla estandarizada/v3/Formato Remuneracion 2026xx-x Total v3.xlsx`,
hoja plantilla `DetRetri2026092` / `DetValiRetri2026092`; en la base preparada el nombre pasa a
`DetRetri2026072` / `DetValiRetri2026072`).

Objetivo: fijar la **gramática** de las columnas protegidas por el mapa congelado
`WorkbookLeafCellMapQ2.DetRetriProtected` — qué identifica cada **fila** y cada **columna** — antes de
sustituir la lista de celdas y fragmentos por una verificación de ROL + FORMA.

## 1. Familias de filas (medido)

Cada hoja declara **tres** familias de filas por ASE, en orden físico. Una fila es «fila de familia»
si `B` trae el id de un ASE del catálogo (1..5) o si `B` está vacía y `C = 'Total'/'TOTAL'`.

| Familia | DetRetri (v3) | DetValiRetri (v3) | Identidad de fila |
|---|---|---|---|
| Principal | 9–14 (encabezado 8 con `B='ASE'`) | 9–14 (encabezado 8) | `B = id` (1..5) y cierre `C='Total'` |
| Diferencia vs CONSOLIDADO | 23–28 (encabezado 22) | **16–21 (sin encabezado propio)** | `B = id` y cierre `C='Total'` |
| Verificación de composición | 32–36 (encabezado 31) | 24–29 (encabezado 23) | `B = id`; DetVali cierra con `C='TOTAL'` |

Hallazgos vinculantes:

1. **La familia de diferencia de `DetValiRetri` no declara encabezado**: sus columnas son las mismas
   que las de la tabla principal (la hoja publica un único juego de columnas `D..O`).
2. Los encabezados de las familias derivadas **no repiten** la columna de identidad `ASE`.
3. Los rótulos no son idénticos entre hojas: `DetRetri!L8 = 'COSTO DE LA INTERVENTORÍA (*)'` y
   `CONSOLIDADO!L8 = 'Costo de la Interventoría'`; `DetVali!D8 = 'TARIFA DE COMERCIALIZACIÓN DEL
   SERVICIO. (TCS)'` y `CONSOLIDADO!U8 = 'Tarifa de Comercialización del Servicio.  ( TCS)'`. La
   comparación de rótulos normaliza acentos, espacios y puntuación (solo letras y dígitos).
4. El CONSOLIDADO declara **dos bloques de columnas en la misma fila de encabezado** (`B..O` y
   `S..AF`). El bloque que una hoja de detalle refleja se identifica por el rótulo de su **primera**
   columna de valor (`Recaudo Total` → bloque `D..O`; `Tarifa de Comercialización…` → bloque `U..AF`).

## 2. Forma de las celdas protegidas (medido)

Con `main` = celda de la familia principal (misma identidad y columna), `dif` = celda de la familia de
diferencia, `CONS` = `CONSOLIDADO_TOTAL RECAUDO`:

| Celda v3 | Fórmula medida | Forma derivada |
|---|---|---|
| `DetRetri!D23` | `D9-'CONSOLIDADO_TOTAL RECAUDO'!D104` | `+main − CONS(tabla TOTAL, misma identidad, misma columna)` |
| `DetRetri!D28` | `D14-'CONSOLIDADO_TOTAL RECAUDO'!D109` | idem con identidad = cierre |
| `DetRetri!D32` | `+D23+'Reporte Componentes R1'!F89+'Rem. Anticipos R2'!E57+'Reversion Pagos R4'!D79+'Reporte Componentes R1'!F91+'AJUSTES - SF-T'!D47=D9` | `LHS = +dif + hojas de origen de las 5 tablas = main` |
| `DetRetri!D34` | `+R1!F484+R1!F486+R2!E299+R4!D217+D25+AJUSTES!D49=D11` | idem (misma forma, otro orden) |
| `DetValiRetri!D16` | `D9-'CONSOLIDADO_TOTAL RECAUDO'!U104` | idem |
| `DetValiRetri!D24` | `+'R1'!M89+'R1'!M91+'R2'!L57+'R4'!J79+D16+AJUSTES!U47=D9` | idem |
| `DetValiRetri!D29` (cierre) | `R1!M89+M91+L57+J79 + M323+M325+L178+J180 + M484+M486+L299+J217 + M692+M694+L423+J332 + M811+M813+L510+J373 + D21 + AJUSTES!U52 = D14` | `LHS = +dif(cierre) + Σ_{ASE}(hojas de las 4 tablas) + cierre propio de la 5.ª tabla` |

Propiedades de la derivación (cero operandos declarados):

- Las **hojas de origen** se obtienen *expandiendo* la fórmula que el CONSOLIDADO declara en cada una
  de las cinco tablas de recaudo (`R1 - RECAUDO OPORTUNO`, `R2 - ANTICIPOS A USUARIOS`,
  `R3 - RECAUDO EXTEMPORANEO`, `R4 - REVERSIÓN DE PAGOS`, `R5 - REMUNERACIÓN SALDOS A FAVOR POR NOTA`):
  una referencia a otra hoja es terminal; una referencia interna o un `SUM(rango)` se expanden. Un
  cambio de cardinalidad de subsecciones (ASE1: `F89`+`F91`; ASE3: `F484`+`F486`) cambia las hojas y
  la verificación sigue siendo cierta.
- Para la identidad **cierre**, el CONSOLIDADO de la tabla R5 suma el rango de sus ASE (`U90 =
  SUM(U85:U89)`), pero la verificación usa el cierre propio de la hoja de origen (`AJUSTES!U52`):
  la regla es «cierre de la tabla → cierre de la hoja origen».

## 3. Hallazgo que acota el alcance (no es un rótulo faltante)

Las **demás columnas** de la familia de verificación **no** siguen la regla de «una hoja por tabla de
recaudo»:

- `J` (Balance de subsidio o contribuciones): `J32 = +'BCE SC POR FACT.'!F3 + J23 = J9` (hoja BCE);
- `L` (Costo de la Interventoría): `INTERVENTORIA!M26 + L23 = L9`;
- `M`, `K`, `O`: sumas anidadas con paréntesis y restas de `INTERVENTORIA`;
- `N` (Conceptos especiales): `R1!L89 + R2!K57 + R4!P79 + R1!L91 + N23 = N9` — **omite** `AJUSTES`, a
  diferencia de `D..I`.

La protección congelada cubría **solo la columna de valor principal** (`D` en las dos hojas). La
derivación implementada verifica exactamente esa columna (la primera columna de valor declarada por el
encabezado principal), que es la única con la composición uniforme. Las columnas derivadas quedan
**fuera de la verificación por forma** con causa medida (no hay rótulo que las distinga y su
composición es heterogénea); si el Ingeniero quiere cubrirlas, requiere una gramática propia por
columna (nueva tanda), no un literal.

## 4. Barrido del camino de escritura (D)

Medido en la corrida real 2026072 con la v3, el camino de escritura tenía **dos** consumidores de
geometría congelada que bloqueaban:

| Consumidor | Síntoma con la v3 | Sustitución |
|---|---|---|
| `WorkbookLeafCellMapQ2.DetRetriProtected` | `DetRetri2026072!D32` exigía `[F75, E47, D73, F77]` | `DetallePeriodoDerivado.Validar` (ROL + FORMA) |
| `WorkbookLeafCellMapDetalleR2R4.BloquesR2/R4` | `ERR-SIN-SLOT`: firmas de la fuente sin slot en el bloque `3..33` (ventana de la v2) | `DelimitadorBloquesHoja.DelimitarBloque` → `(EncabezadoColumnas+1 … GranTotal)` |

Tras ambos cambios la corrida **genera el workbook** (`… (sin validar).xlsx`, 625 KB) y el primer
error que aparece es de **negocio**, no de camino de escritura: `ERR-VALIDACION` en
`VALIDACION_TOTAL · D34/F34/O25` (deuda declarada del Plan 57/H3).

Consumidores congelados que **siguen** en el camino productivo y no bloquean esta corrida (inventario
para la siguiente tanda): `WorkbookLeafCellMapDetRetri` (filas de escritura del DetRetri),
`WorkbookLeafCellMapAjustesSfT`, `WorkbookLeafCellMapBalanceSc`, `WorkbookLeafCellMapReporteBanco`,
`WorkbookLeafCellMapValidaciones`, `WorkbookLeafCellMapInterventoria`, `WorkbookLeafCellMapPorAse`,
`WorkbookLeafCellMapPorEmpresa`.

## 5. G-HUELLA (B)

La huella semántica se implementó como helper compartido de test
(`Remuneracion.IntegrationTests/HuellaSemanticaPlantilla.cs`): entradas `hoja|nombre|dimension` y
`celda|hoja|ref|f|texto|v|valor|si|índice|ref|rango` por cada celda con `<v>`, ordenadas
ordinalmente y unidas con salto de línea; SHA-256 en hexadecimal minúscula. Valores medidos en disco:

| Plantilla | Huella semántica (gate) | SHA de archivo (informativo) |
|---|---|---|
| v3 | `92935014f3611055c15f12db92c71020f741d009634e8ec59f1e4448235d72ff` | `5ACE6F5A…FBF71F1` (709.070 B) |
| v2 | `a4ea462b57be95938d0a25a9369bfd1e19d2d1229b75de7d89fbf71fb5991879` | `86E48E47…9D3D14` (693.983 B) |
| v1 | `cf9aa03604ed72b2252876df83abd573eb5609b14a97026ef7c1b23774490d17` | `3086AC31…A49C5214` (692.598 B, sin cambios) |

**Decisión requerida (§0.6).** Las huellas publicadas en el plan (`c371a5fb…c2e0` y
`dd7f5891…2329`) provienen de una normalización no documentada; con el algoritmo canónico de este
helper los valores son los de la tabla. Ambos gates miden lo mismo (contenido, no empaquetado) y el
helper es estable frente al re-guardado (los blobs de `2002b3f^` y los vigentes producen la misma
huella). Se pide al Ingeniero ratificar los valores de esta tabla como ancla de `G-HUELLA` o indicar
la normalización exacta con la que se calcularon los del plan.
