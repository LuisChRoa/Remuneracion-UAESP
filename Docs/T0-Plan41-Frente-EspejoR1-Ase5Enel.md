# T0 (Plan 41 — Frente Espejo R1) — Diagnóstico read-only del bloque ASE5–ENEL y alcance real

> **Frente:** el sub-bloque `ASE5–ENEL` del espejo R1 (`Reporte Componentes R1`) que el defecto E1-2 del
> Plan 40 dejó abierto: la app escribe valores pero las anclas por empresa de `REMUNERACION_*!D9:D13`
> quedan cortas/cero frente al manual del administrativo.
> **Naturaleza:** T0 DIAGNÓSTICO — **cero código de producción**. Lectura SOLO con BCL
> (`System.IO.Compression.ZipArchive` + `System.Xml.Linq`), sin Excel/COM, sin OpenXml SDK.
> **Fixture read-only:** `Remuneracion.IntegrationTests/T0FrenteEspejoR1Ase5EnelTests.cs` (congela los
> bytes citados aquí; no toca producción).
> **El OFICIAL nunca se modifica.** Hashes al inicio y al cierre de esta corrida (§0).
> **NO se propone aquí el fix como implementación**: §7 da causa raíz citada; §8 da opciones con trade-offs.

## 0. Hashes (R-INPLACE)

| Archivo | SHA256 |
|---|---|
| Oficial `Docs/Plantilla estandarizada/Formato Remuneracion 2026xx-x Total.xlsx` (antes y después) | `3086AC317595C2F8E84ED24B1C5EE3E52565DD5193394E3062878E7FA49C5214` |
| Salida de la app `Docs/Prueba Julio-2/Resultado/Resultado5/Remuneración 202607-2 Total.xlsx` | `1CE66248E9963E9C0E7503DE6C8F66A7AE59189718C9203C70C4414D90A74F5E` |
| Manual oráculo `Docs/Prueba Julio-2/Resultado/Remuneracion 202607-2 Total Administrativo.xlsx` | `584310105AC7223CE840833E3CB64E26F95B61907C9ECD959A0A632EEC21DCB1` |
| Fuente R1 ASE5 `…/5-Área Limpia/Recaudoporcomponente_…_20268465657830.xlsx` | `963D8C9F8EABBAD895D29752B1EEF5C8AC377E04CA95B6CC6DB7085B31AE054F` |

Los valores de la app se **evalúan** desde las celdas realmente escritas (el writer no recalcula: los
`<v>` de celdas-fórmula son stale por diseño, `fullCalcOnLoad` los recalcula al abrir Excel). El
resolutor BCL acotado (sumas + referencias + `'Hoja'!Ref`) reproduce el recálculo de Excel.

## 1. Lectura vs escritura — VEREDICTO: la app ENCUENTRA el dato y NO lo coloca en el slot

**No es un defecto de lectura.** La fuente R1 del ASE5 (`5-Área Limpia/Recaudoporcomponente…`) trae
**2 filas `ENEL`**:

| Fuente (hoja `Sheet1`) | Fila | Sección | `F` (Total) | `L` (Servicio Especiales) |
|---|---|---|---|---|
| `Recaudo Desde 16/07/2026` | 18 | `Aplicacion nuevos x reversion` (1.ª) | `12823196468.26` | `8997077.1` |
| | 43 | `Aplicacion nuevos x reversion` (2.ª) | `-1113041546.26` | `0` |

La app **sí escribió ambas filas**, pero en slots equivocados del bloque ASE5 (`603–676`) de la
geometría oficial:

| Slot oficial (fila) | Sección del slot | `F` app | `L` app | Valor fuente | ¿Correcto? |
|---|---|---|---|---|---|
| `605` | 1.ª sección (antes del `Mes`) | `12823196468.26` | `8997077.1` | fila 18 | **NO** (debería ir a `633`) |
| `618` | `Oportuno` (tras `Mes 613`) | `-1113041546.26` | `0` | fila 43 | **NO** (debería ir a `663`) |
| **`633`** | `Aplicacion nuevos x reversion` (1.ª) | `0` | `0` | fila 18 | **NO — slot vacío** |
| `652` | `Mes` 2 (antes del `Aplicacion` interior) | `0` | `0` | — | — |
| **`663`** | `Aplicacion` (2.ª, interior) | `0` | `0` | fila 43 | **NO — slot vacío** |

La ancla oficial `REMUNERACION_ENEL!D13 = 'Reporte Componentes R1'!F690`, con
`F690 = F663+F633-L633` (oficial, byte-fiel), evalúa en la salida de la app:

```
F690 = F663(0) + F633(0) - L633(0) = 0        (app)
F690 = F663(-1.113.041.546,26) + F633(12.823.196.468,26) - L633(8.997.077,10) = 11.701.157.844,90   (colocación por sección)
manual: REMUNERACION_ENEL!D13 = 'Reporte Componentes R1'!F568 = 11.701.157.844,90
```

**Conclusión:** el defecto es de **ESCRITURA/COLOCACIÓN del slot** (dato hallado, slot mal elegido),
no de lectura ni de dominio. El par de filas fuente existe y se consume; lo que falla es **a qué
slot de la geometría fija oficial se asigna cada fila**.

## 2. Alcance real — no es solo ASE5+ENEL

Se barrió **5 empresas × 5 ASE** de la familia espejo R1 (las que consumen `REMUNERACION_*!D9:D13`).
Correspondencia de anclas (oficial ↔ manual legacy, por empresa y ASE):

| Empresa | ASE1 | ASE2 | ASE3 | ASE4 | ASE5 |
|---|---|---|---|---|---|
| `REMUNERACION Reciprocidad EAAB` | F80↔F58 | F259↔F211 | F420↔F348 | F573↔F473 | F685↔F563 |
| `REMUNERACION_ENEL` | F85↔F63 | F264↔F216 | F425↔F353 | F578↔F478 | F690↔F568 |
| `REMUNERACION_ENERBIT` | F90↔F68 | F269↔F221 | F430↔F358 | F583↔F483 | F695↔F573 |
| `REMUNERACION_OCCIDENTE_ Directa` | F95↔F73 | F274↔F226 | F435↔F363 | F588↔F488 | F700↔F578 |
| `REMUNERACION_EAAB-CL` | F100↔F78 | F279↔F231 | F440↔F368 | F593↔F493 | F705↔F583 |

> **Trampa de comparación (R-TRAMPA-ANCLA):** las anclas de fila **difieren** entre el manual legacy
> (`F63/F216/F353/F478/F568` para ENEL) y la geometría oficial (`F85/F264/F425/F578/F690`). **Nunca**
> se comparan direcciones entre plantillas distintas: la correspondencia es por (empresa, ASE), y el
> valor esperado es el del manual en SU ancla. Tampoco confundir las dos familias del oficial:
> `CONSOLIDADO_TOTAL RECAUDO!D9:D13` referencia `F75/F254/F415/F568/F680` (TOT_OPT por ASE) mientras
> `REMUNERACION_ENEL!D9:D13` referencia `F85/F264/F425/F578/F690` (subtotal por empresa).

### Tabla ancla → esperado (manual) → observado (app)

| Empresa | ASE | Ancla oficial | Fórmula oficial (`Reporte Componentes R1`) | Esperado (manual) | Observado (app) | ¿Escrito el slot? |
|---|---|---|---|---|---|---|
| Reciprocidad | 1 | F80 | valor plano | `0` | `0` | n/a |
| **Reciprocidad** | **2** | **F259** | `F200+F160-L160` | **`11.411.240,00`** | **`-3.892.093,19`** | parcial (slots 135/160) |
| Reciprocidad | 3 | F420 | valor plano | `0` | `0` | n/a |
| **Reciprocidad** | **4** | **F573** | `F540+F500-L500` | **`5.616.850,00`** | **`5.353.948,68`** | parcial (slots 475/500/522) |
| Reciprocidad | 5 | F685 | valor plano | `0` | `0` | n/a |
| ENEL | 1 | F85 | `F62+F36+F8-L8-L36` | `16.560.599.150,50` | `16.560.599.150,50` | **sí (5/5)** |
| ENEL | 2 | F264 | `F193+F150+F112--L112-L150` | `19.023.525.223,11` | `19.023.525.223,11` | **sí (5/5)** |
| **ENEL** | **3** | **F425** | `F345+F319+F291-L291-L319` | **`13.997.194.578,87`** | **`1.307.680,00`** | parcial (slots 291/304/319/334; falta 345) |
| ENEL | 4 | F578 | `F533+F490+F452-L452-L490` | `6.584.787.380,47` | `6.584.787.380,47` | **sí (5/5)** |
| **ENEL** | **5** | **F690** | `F663+F633-L633` | **`11.701.157.844,90`** | **`0`** | parcial (slots 605/618; faltan 633/663) |
| ENERBIT | 1–4 | F90/F269/F430/F583 | valor plano | `0` | `0` | n/a |
| ENERBIT | 5 | F695 | `F666+F637-L637` | `26.864.220,00` | `26.864.220,00` | **sí** |
| **OCCIDENTE** | **1** | **F95** | `F69+F46+F15-L15-L46` | **`779.194.409,36`** | **`168.020.291,14`** | parcial (slots 15/31/46) |
| OCCIDENTE | 2 | F274 | `F207+F170+F119-L119-L170` | `1.317.332.077,61` | `1.317.332.077,61` | **sí (5/5)** |
| **OCCIDENTE** | **3** | **F435** | `F352+F329+F298-L298-L329` | **`1.229.014.628,26`** | **`404.732.299,55`** | parcial (slots 298/314/329) |
| OCCIDENTE | 4 | F588 | `F547+F510+F459-L459-L510` | `457.088.264,26` | `457.088.264,26` | **sí (5/5)** |
| **OCCIDENTE** | **5** | **F700** | `F647+F673+F612-L612-L647` | **`304.989.620,81`** | **`441.849,00`** | parcial (slots 612/628/647/659) |
| EAAB-CL | 1–5 | F100/F279/F440/F593/F705 | valor plano | `0` | `0` | n/a |

**Lectura del patrón (mecanismo):**

- Cuando el nº de filas de la empresa en la fuente **igual** al nº de slots fijos del oficial, el
  cursor físico coincide con la sección y el resultado es **correcto** (ENEL ASE1/2/4: 5 filas → 5
  slots; OCCIDENTE ASE2/4; ENERBIT ASE5).
- Cuando la fuente trae **menos** filas que slots fijos, el cursor llena los **primeros** slots y deja
  vacíos los que la ancla referencia (ENEL ASE3: 4 filas vs 5 slots → falta el 5.º; ENEL ASE5: 2 vs 5;
  OCCIDENTE ASE1: 3 vs 5; OCCIDENTE ASE3: 3 vs 5; Reciprocidad ASE2: 2 vs 4; Reciprocidad ASE4: 3 vs 4).
- **Nuance de estructura de fórmula (OCCIDENTE ASE5):** el oficial `F700` tiene **3 términos**
  (`F647+F673+F612-L612-L647`) y el manual `F578` **2** (`F530+F551-L530`). Aun colocando por sección
  (`F612 = 417.519,62`), el oficial daría `302.832.620,81` ≠ `304.989.620,81` del manual. Es una
  **divergencia de estructura de la fórmula**, no de colocación → **mini-T0 propio** (§9). Los demás
  pares en rojo tienen fórmula oficial estructuralmente análoga a la del manual → son colocación pura.

**Sub-tabla de slots por sección (evidencia de la colocación por sección que sí cierra):**

| Empresa/ASE | Slots oficiales (sección) | Filas fuente | Colocación por sección | Colocación de la app |
|---|---|---|---|---|
| ENEL ASE3 | 291(1ª), 304(Oportuno), 319(Aplic1), 334(Oportuno2), 345(Aplic2) | src13→304, src18→319, src33→334, src37→345 | **F425 = 13.997.194.578,87 ✓** | 291,304,319,334 → F425 = 1.307.680,00 |
| ENEL ASE5 | 605(1ª), 618(Oportuno), 633(Aplic1), 652(Oportuno2), 663(Aplic2) | src18→633, src43→663 | **F690 = 11.701.157.844,90 ✓** | 605,618 → F690 = 0 |
| OCCIDENTE ASE1 | 15(1ª), 31(Oportuno), 46(Aplic1), 58(Oportuno2), 69(Aplic2) | src10→15, src30→46, src46→69 | **F95 = 779.194.409,36 ✓** | 15,31,46 → F95 = 168.020.291,14 |
| OCCIDENTE ASE3 | 298(1ª), 314(Oportuno), 329(Aplic1), 341(Oportuno2), 352(Aplic2) | src7→298, src28→329, src44→352 | **F435 = 1.229.014.628,26 ✓** | 298,314,329 → F435 = 404.732.299,55 |
| Reciprocidad ASE2 | 135(1ª), 160(Oportuno), 182(Aplic1), 200(Aplic2) | src38→?, src69→? | **F259 = 11.411.240,00 ✓** | 135,160 → F259 = -3.892.093,19 |

> La "colocación por sección" asigna cada fila fuente al slot cuya **sección (frontera inmediatamente
> ARRIBA de la fila)** coincide — `Mes` → Oportuno, `Aplicacion` → Aplicación, sin frontera → 1.ª
> sección — y consume dentro de la sección en orden. Con esa regla **todos** los pares en rojo
> (salvo OCCIDENTE ASE5) cierran al valor del manual.

## 3. Mecanismo en código (por qué el cursor elige el slot equivocado)

`EscritorValoresSlotR1.Escribir` resuelve cada fila fuente a un slot así:

1. `ConstruirGruposFirma(slots)` (`EscritorValoresSlotR1.cs:275-292`) agrupa los slots por **firma
   A/B/C/D/E** solamente: `grupos.Find(g => EncajaEnSlot(firmaSlot, null, g.Representante))`. El
   argumento de frontera es **`null`** y `FilaDesdeSlot` (`:307-314`) no lleva frontera → el grupo es
   **una lista por firma** (todos los `C=ENEL, D=Total` del bloque en un solo grupo, en orden físico).
2. `BuscarGrupoFirma(fila, fronteraFuente)` (`:301-305`) elige el grupo comparando contra el
   **representante** (primer slot) y, si no hay match con frontera, reintenta con `null`:
   `conFrontera ?? grupos.Find(g => EncajaEnSlot(fila, null, g.Representante))`. Como el representante
   ENEL del ASE5 es `605` (Oportuno) y las filas fuente ENEL también clasifican Oportuno, el match cae
   al grupo único.
3. `GrupoFirma.SiguienteLibre()` (`:190-191`) es un **cursor físico**: `_slots[_cursor++]`. Consume
   `605, 618, 633, …` en orden, **sin mirar la sección de cada slot**.

El discriminador de sección disponible, `R1FirmaInterior.EsAplicacionPorFrontera`
(`Remuneracion.Core/Models/R1FirmaInterior.cs:127-143`), mira la **frontera por DEBAJO** de la fila
("la próxima fila por debajo que sea Mes/Aplicacion"). Para ASE5 las dos filas fuente ENEL clasifican
`Oportuno` (su próxima frontera por debajo es `Mes`), igual que los slots `605/633/663`; por eso ni
siquiera separándolos por esa frontera se obtendría `633/663` — hace falta la **pertenencia a la
sección (frontera ARRIBA)**, que hoy no se computa para resolver el slot.

**Dónde se pierde ASE5–ENEL, en una línea:** en `GrupoFirma.SiguienteLibre`
(`EscritorValoresSlotR1.cs:190-191`) consumiendo el grupo de firma `ENEL` en orden físico
(`ConstruirGruposFirma`, `:275-292`), sin que el slot elegido respete la sección de la fila fuente.

## 4. Segundo hueco — `AJUSTES - SF-T!D47`: defecto DISTINTO (no es el espejo R1)

`AJUSTES - SF-T!D47` **no** es un bloque ENEL; es el ajuste del **ASE1**:

```
oficial  AJUSTES - SF-T!D47 = D9+D28   (D9='SALDOS POR NOTA'!C10 ; D28='RETRIBUCION NEGATIVA'!C10)
app      AJUSTES - SF-T!D47 = 0
manual   AJUSTES - SF-T!D47 = 973.693,46
```

**Cadena real (distinta del espejo R1):** `SALDOS POR NOTA` (hoja de la conciliación de saldos por
nota, Q2). La app **sí escribió el valor `973.693,46`** pero en la fila equivocada del bloque fijo de
`SALDOS POR NOTA`:

| Celda `SALDOS POR NOTA` | Oficial (etiqueta) | App | Manual (ancla legacy) |
|---|---|---|---|
| `C7` | `Subs/Cont Total` | **`973.693,46`** (valor de la fila `Total` fuente) | `176.919,99` |
| `C8` | `Total` | **`0`** | `973.693,46` |
| `C10` | `=C8-I8` (Oportuno) | **`0`** | `=C7-I7` → `973.693,46` |

La fuente `SaldosaFavorAplicadosPorNotas…` trae `Vlr Servicio / Componente / Subsidio / Subs/Cont /
Total` (sin fila `Vlr Intereses`); el bloque oficial de `SALDOS POR NOTA` tiene una fila **`Vlr
Intereses` extra**, y la app desplaza las filas fuente **+1** (coloca `Total` en `Subs/Cont`). Es la
**misma clase de defecto** (colocación por posición en una geometría fija), pero en **otra hoja y otra
familia** (saldos por nota / retribución negativa, Q2). **No es el espejo R1**; va con su propio
frente. Aplica igual a ASE1–4 (la app deja `C8/C21/C34/C47` = 0) y ASE5 trae 0 en la fuente.

## 5. `INTERVENTORIA!S26:S31` — VEREDICTO: NO es invariante; requiere su propio T0

**Fórmulas (oficial, byte-fiel):**

```
R26 = =J15                        (maestro shared ref=R26:R31 si=0 → R27=J16 … R31=J20)
Q26 = ='CONSOLIDADO_TOTAL RECAUDO'!L104     (Q27=L105 … Q31=L109)
S26 = =Q26=R26   (S27=Q27=R27, S28:S31 shared ref=S28:S31 si=2)
```

**Tabla J y variantes (oficial):**

| Celda | Oficial | Manual | Naturaleza |
|---|---|---|---|
| `J15:J19` | `189185988, 266580256, 154788536, 120391083, 128990446` | idénticos | **literales** (tabla anual declarada) |
| `M26:M30` (2.ª quincena) | mismos valores que `J15:J19` | idénticos | declarado Q2 |
| `N26:N30` (1.ª quincena) | `189185987, 266580255, 154788535, 120391083, 128990445` (`= M-1`) | idénticos | declarado Q1 |

**Veredicto motivado:** `R26:R31` referencia la **tabla J fija** (`=J15:J19`), que **coincide con la
columna `M` (2.ª quincena)**. Pero `Q26:Q31` sigue `CONSOLIDADO!L104:L109`, cuya cadena arranca en
`L9:L13 = INTERVENTORIA!M26:M30` en Q2 y **`INTERVENTORIA!N26:N30` en Q1** (regla R-INTERVENT-Q, Plan
39). Por tanto:

- **Q2:** `S26 = (L104 = J15) = (M26 = M26) = TRUE`.
- **Q1:** `S26 = (N26 = J15) = (189185987 = 189185988) = FALSE`.

⇒ **`S26:S31` debe variar por quincena; el oficial NO trae una variante `N` para la tabla J** (la tabla
J es el valor Q2; los valores Q1 viven solo en `N26:N30`). Es la **misma familia F-NEW-1** del
Plan 39/40 (auto-chequeo con columna fija `M`). **No es invariante** y **no se puede sellar TRUE**
(sería mentira): requiere **T0 propio de la tabla J** → reescritura Phase 0 de `R26:R31` por
`Periodo.NumeroQuincena` (`=J15:J19` en Q2 → `=N26:N30` en Q1), o veredicto de invariante **solo con
cita de negocio**. Evidencia cruzada: `J15:J19 == M26:M30` (no hay columna Q1 de J); manual legacy
idéntico (`R26=F15`); `CONSOLIDADO!L9:L13` ya es quincena-dependiente.

## 6. Impacto en las hojas entregadas

Las hojas que recibe el administrativo y quedan **incompletas/incorrectas** son las de la familia
por empresa (subtotales `D9:D13`), no el consolidado ni el `DetRetri`:

| Hoja | Celdas afectadas | Magnitud (esperado → observado) |
|---|---|---|
| `REMUNERACION_ENEL` | `D11` (ASE3), `D13` (ASE5) | `13.997.194.578,87 → 1.307.680,00`; `11.701.157.844,90 → 0` |
| `REMUNERACION_OCCIDENTE_ Directa` | `D9` (ASE1), `D11` (ASE3), `D13` (ASE5) | `779.194.409,36 → 168.020.291,14`; `1.229.014.628,26 → 404.732.299,55`; `304.989.620,81 → 441.849,00` |
| `REMUNERACION Reciprocidad EAAB` | `D10` (ASE2), `D12` (ASE4) | `11.411.240,00 → -3.892.093,19`; `5.616.850,00 → 5.353.948,68` |
| `GERENTES_ENEL` / `GERENTES_OCCIDENTE _ Directa` / `GERENTES_Reciprocidad_ EAAB` | `D9:D13` (espejo `=REMUNERACION_*!D104:D108`) | heredan la magnitud anterior |
| `VALIDACION_ENEL` / `VALIDACION_OCCIDENTE` / `VALIDACION_RECIP` | columnas de contraste por empresa | heredan la divergencia (contrastan `Recaudo *` vs `REMUNERACION_*`) |
| `AJUSTES - SF-T` | `D47:D51` (ajuste por ASE) | `973.693,46 → 0` (ASE1); clase de §4 (saldos por nota) |

**No afectadas** (correctas hoy): `CONSOLIDADO_TOTAL RECAUDO!D9:D13`/`D47:D51` (usan el TOT_OPT/EXTEMP
por ASE `F75/F254/F415/F568/F680`, que la app coloca bien), `DetRetri{per}`/`DetValiRetri{per}`, y las
empresas sin divergencia (`ENERBIT`, `EAAB-CL`, `Reciprocidad` ASE1/3/5, `OCCIDENTE` ASE2/4, `ENEL`
ASE1/2/4).

## 7. Causa raíz citada (solo lectura)

| # | Archivo:línea | Rol en el defecto |
|---|---|---|
| C1 | `Remuneracion.Infrastructure/Excel/EscritorValoresSlotR1.cs:275-292` (`ConstruirGruposFirma`) | Agrupa slots por firma A/B/C/D/E con frontera `null` → **un grupo por firma**, en orden físico. |
| C2 | `Remuneracion.Infrastructure/Excel/EscritorValoresSlotR1.cs:190-191` (`GrupoFirma.SiguienteLibre`) | Cursor **físico** (`_slots[_cursor++]`): consume el 1.º, 2.º, … slot de la firma sin mirar sección. |
| C3 | `Remuneracion.Infrastructure/Excel/EscritorValoresSlotR1.cs:301-305` (`BuscarGrupoFirma`) | El match de frontera se hace contra el **representante** del grupo, no contra cada slot; cae al grupo único. |
| C4 | `Remuneracion.Core/Models/R1FirmaInterior.cs:127-143` (`EsAplicacionPorFrontera`) | Clasifica por la frontera **por debajo**; no distingue "1.ª sección vs Oportuno vs Aplic1 vs Aplic2" como requiere la ancla. |
| C5 | `Remuneracion.Infrastructure/Excel/EscritorValoresSlotR1.cs:363-407` (`LeerSlots`) | Construye la lista de slots del bloque (`RangoDatosAse` 603–676 para ASE5); **no hay** resolución por sección. |
| C6 | `Remuneracion.Infrastructure/Excel/EscritorValoresSlotR1.cs:65-73` (`RangoDatosAse`) | Geometría fija del bloque (superset); el nº de slots por empresa puede **exceder** el nº de filas fuente. |

**Cadena causal:** `LeerEspejoR1` (lectura) entrega **todas** las filas → `EscritorValoresSlotR1`
agrupa por firma (C1) y consume por cursor físico (C2) → cuando la fuente tiene menos filas que slots
fijos, las filas caen en los **primeros** slots → los slots que la ancla por empresa referencia
(`F633/F663` ENEL ASE5; `F345` ENEL ASE3; …) quedan vacíos → el ancla evalúa corta/cero.

## 8. Opciones de fix con trade-offs (SIN implementar)

| Opción | Qué implica | Pro | Contra |
|---|---|---|---|
| **A. Resolución de slot por sección (recomendada)** | Agrupar por (firma + **sección de la fila**) y consumir dentro de la sección; la sección = frontera **inmediatamente arriba** (`Mes`→Oportuno, `Aplicacion`→Aplicación, sin frontera→1.ª). Un predicado nuevo en Core (p. ej. `SeccionPorFrontera(filas, i)`) consumido por `EscritorValoresSlotR1`; `EsAplicacionPorFrontera` (Plan 35-T2) **no se toca** (otro consumidor). | Cierra ENEL ASE3/5, OCCIDENTE ASE1/3, Reciprocidad ASE2/4 al valor del manual; cambio acotado a la resolución de slot; sin ramas por empresa/período (D-H / catálogo abierto). | Requiere T0 de la regla de sección sobre los 5 ASE; riesgo R-ATRIBUCION (dinero a la empresa equivocada) si el predicado falla → gate por empresa×ASE. |
| **B. Slot map congelado (ASE × empresa × sección → celda fija)** | T0 BCL que congela la celda exacta por (ASE, empresa, sección) desde el oficial; el writer la consume directo, sin cursor. | Máxima determinación; fail-fast si la fuente trae una empresa/sección sin slot. | Fixture grande (5 ASE × 5 empresas × 5 secciones); más superficie de contrato que mantener ante cambio de plantilla (nuevo T0). |
| **C. Ajustar las anclas por empresa del oficial** | Reescribir `F690`/`F425`/… para que referencien los slots que la app llena. | "Mínimo" si solo se mira el síntoma. | **PROHIBIDO** por doctrina (el oficial nunca se edita; el `RangoDatosAse`/geometría es contrato T0); además movería el problema a las anclas `F85/F264/F578` que hoy sí cierran. |
| **D. Reintroducir el espejo dimensional (Plan 21)** | Insertar/borrar filas para que el bloque oficial espeje la fuente 1:1. | Colocación trivial. | **PROHIBIDO** por slot doctrine (Plan 36: cero insert/delete en procesamiento); rompe shared-formula/calcChain. |

**Recomendación:** **A** (resolución por sección) con T0 propio por ASE, gate workbook-vs-dominio por
empresa×ASE (lección Plan 31: el gate lee el WORKBOOK), y las anclas del oficial **intactas**.

## 9. Follow-ups declarados (fuera de este T0 diagnóstico)

1. **Frente espejo R1 (este defecto):** plan hijo con T0 propio → opción A + gate por empresa×ASE.
2. **Mini-T0 OCCIDENTE ASE5:** la ancla oficial `F700` (3 términos) difiere estructuralmente del manual
   `F578` (2 términos) → decidir la verdad de negocio antes de fijar el valor esperado.
3. **Frente saldos por nota / retribución negativa (Q2):** `AJUSTES - SF-T!D47:D51` = 0 por la fila
   `Vlr Intereses` extra en `SALDOS POR NOTA` (colocación +1); T0 propio.
4. **T0 propio de la tabla J (`INTERVENTORIA!J15:J19` / `R26:R31`):** reescritura Phase 0 por quincena
   (familia F-NEW-1) o invariante con cita — **prohibido sellar TRUE**.

## 10. Método y reproducibilidad

- Lectura BCL ZIP+XML de: oficial, `Resultado5`, manual `…Administrativo.xlsx`, y las 5 fuentes R1
  `Recaudoporcomponente_*.xlsx` por ruta explícita (ASE1 `…/1-Promoambiental/…20268316364772.xlsx`,
  ASE2 `…/2-Lime/…202683163652756.xlsx`, ASE3 `…/3-Ciudad Limpia/…202683202446434.xlsx`,
  ASE4 `…/4-Bogota Limpia/…202683202831773.xlsx`, ASE5 `…/5-Área Limpia/…20268465657830.xlsx`).
- Evaluación de anclas por resolutor BCL acotado (sumas + refs + `'Hoja'!Ref`); el fixture del test
  read-only congela los textos-`<f>` y los valores citados aquí.
- **Prohibido** `Docs/Insumos` raíz; **prohibido** usar cachés `<v>` como oráculo; el oficial se
  verifica intacto por SHA256 (§0).
