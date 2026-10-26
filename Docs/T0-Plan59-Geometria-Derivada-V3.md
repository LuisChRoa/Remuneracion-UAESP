# T0 — Plan 59: geometría derivada de la plantilla v3 (H0/F0, read-only)

> **Fase:** H0/F0 — T0 read-only bloqueante (cero código de producción).
> **Plan (fuente única de verdad):** `plans/59 - eliminacion-coordenadas-quemadas-geometria-derivada-v3.md` — estado **APROBADO 2026-10-26**, Clarification Gate §6 ratificado 8/8.
> **Fixture de evidencia:** `Remuneracion.IntegrationTests/GeometriaDerivadaV3T0Plan59Tests.cs`.
> **Método (§2-D0):** lectura **por bytes** con BCL puro (`System.IO.Compression.ZipArchive` + `System.Xml.Linq`). Sin Excel/COM, sin `DocumentFormat.OpenXml`, sin ExcelDataReader. Rutas **explícitas**; prohibido `Insumos.Raiz()`. Cero escritura sobre plantillas, fuentes o artefactos.
> **Regla de oro aplicada:** si un número medido difiere de §0.4 ⇒ STOP y reporte con archivo/hoja/celda/valor esperado vs obtenido. **No hubo ninguna discrepancia con §0.4** (ver §1).

---

## 0. Veredicto global

| Frente | Criterio | Veredicto | Número medido |
|---|---|---|---|
| **T0.1** | S1.1 — SHA v3 + filas/cabeceras (§0.4-b) + superconjunto (§0.4-a) + 0 `<f>` en zonas nuevas (§0.4-c) | **PASS** | v3 `8199F864…2B1295` / 705.934 B; 40 hojas; R1 843 (4,122,356,517,725), R2 527 (1,75,196,317,441), R4 394 (1,97,198,235,350); **0 faltantes** de etiquetas; **0** celdas `<f>` en las 5 zonas nuevas |
| **T0.2** | S1.2 — vocabulario de cierres por hoja + trailer §0.4-d + R4 sin canal | **PASS** | 4 niveles en R1, 3 en R2, 2 en R4; trailer confirmado **42/42** cierres de empresa en R1 oct-1 |
| **T0.3** | S1.3 — destinos CONSOLIDADO/DetRetri derivables por etiqueta | **PASS (derivable-por-etiqueta en las 36 celdas)** | 6 tablas + 30 celdas protegidas + 5 destinos DetRetri resueltos por rótulo; 0 celdas sin rótulo |
| **T0.4** | S1.4 — mapa `canal A`/`EAB` del 2.º cuadro R2 + `E285` valor | **PASS con hallazgo** | 5 ASE × 2 grupos en v3 (vs 1 en v2); `+149.137,87` / `−18.613,87` (ASE2 oct-1); `R2!E285` **sin `<f>`** en v3 — **hallazgo H-59-1** (§5) |
| **T0.5** | S1.5 — SHAs intactos + Δ de suite cerrado | **PASS con reserva** | v1/v2/v3 intactos antes y después; **B vs C = 0 cambios de resultado (solo +23 altas verdes)**; el trx base de la puerta está *stale* por el commit `88b3acb` (insumos Julio-2, 2026-10-26 07:46, anterior a H0) ⇒ 6 rojos nuevos con causa citada y ajena a H0 (§6.4) |

**Puerta H0→H1:** los cuatro frentes cerrados y S1.1–S1.5 verdes. **No hay frente que solo cierre con un literal o un caso particular**; los puntos que requieren arbitraje/atención del Ingeniero son los hallazgos de §5 (ninguno bloquea H1/H2; tres de ellos son insumos de diseño y uno refuerza el mandato de H2).

**Tamaño de la evidencia:** `Remuneracion.IntegrationTests/GeometriaDerivadaV3T0Plan59Tests.cs` — **23 tests**, todos verdes, cero `Skip`, cero aserciones vacías.

**Alcance de rutas (§2-D0):** el fixture consume **explícitamente** la v3, las históricas v1/v2 (solo SHA/superconjunto), las 5 fuentes `Recaudoporcomponente` y `RerpoteDetalleSaldosaFavor` de octubre-1 y la fuente `Recaudoporcomponente` de julio-1. Los otros dos insumos que §2-D0 enumera (`Docs/Prueba Julio-2/Insumos/R10_Remuneracion_2026072.xlsx` y `Docs/Prueba Octubre-1/Resultados/Resultado administrativo/Remuneracion 202610-1 Total.xlsx`) **no participan de ninguna medición T0.1–T0.4** (son oráculos de los gates de corrida CA-B1/CA-B3 de H1/H3); incluirlos aquí sería una aserción de existencia sin contenido. Se declara, no se omite en silencio.

---

## 1. T0.1 — Re-medición de §0.4 en disco (S1.1)

### 1.1 SHA256 y longitud de las tres plantillas (G-SHA / CA-B6)

| Plantilla | Ruta | SHA256 | Bytes |
|---|---|---|---|
| **v3 (vigente)** | `Docs/Plantilla estandarizada/v3/Formato Remuneracion 2026xx-x Total v3.xlsx` | `8199F8647A7ACA84812AD3A61F6607956B2270E0B5C1C999C7690B62032B1295` | 705.934 |
| v2 (histórica) | `Docs/Plantilla estandarizada/v2/Formato Remuneracion 2026xx-x Total vf.xlsx` | `0F192AD2721C77986B6A8CB85D1A92F782DCD2F69989F362E066FA17027F2DAB` | 693.400 |
| v1 (histórica) | `Docs/Plantilla estandarizada/Formato Remuneracion 2026xx-x Total.xlsx` | `3086AC317595C2F8E84ED24B1C5EE3E52565DD5193394E3062878E7FA49C5214` | 692.598 |

Los tres valores coinciden **al byte** con §0.3/§0.4. Re-verificados al inicio y al final de la corrida de mediciones (`T0_5_PlantillasByteIdenticasAntesYDespues…`).

### 1.2 Inventario de hojas (§0.4-a)

- **40 hojas en v3 y 40 en v2, con el MISMO nombre y el MISMO orden** (comparación de secuencia, no de conjunto; `T0_1_CuarentaHojasConElMismoNombreYOrdenEnV2YV3`).
- **37 hojas con geometría byte-idéntica** entre v2 y v3. Las **únicas 3** con dimensión distinta son `Reporte Componentes R1`, `Rem. Anticipos R2` y `Reversion Pagos R4`:

| Hoja | v2 (`dimension`) | v3 (`dimension`) |
|---|---|---|
| `Reporte Componentes R1` | `A1:AP712` | `A1:AP843` |
| `Rem. Anticipos R2` | `A1:S485` | `A1:S527` |
| `Reversion Pagos R4` | `A1:R382` | `A1:R394` |

⇒ **columnas intactas**: R1 = `A:AP`, R2 = `A:S`, R4 = `A:R`. Cero columnas nuevas o eliminadas.

### 1.3 Filas y cabeceras de bloque por ASE (§0.4-b)

La **cabecera de bloque** es la fila donde la columna **B** trae el rótulo del ASE (`Ase` identity by label, no por índice):

| Hoja | v2 filas / cabeceras | v3 filas / cabeceras | Δ filas | Δ por ASE (v3−v2) |
|---|---|---|---|---|
| `Reporte Componentes R1` | 712 / **4, 108, 287, 448, 601** | **843** / **4, 122, 356, 517, 725** | **+131** | ASE1 +14, ASE2 +55, ASE3 0, ASE4 +55, ASE5 +7 |
| `Rem. Anticipos R2` | 485 / **1, 65, 173, 291, 402** | **527** / **1, 75, 196, 317, 441** | **+42** | ASE1 +10, ASE2 +13, ASE3 +3, ASE4 +13, ASE5 +3 |
| `Reversion Pagos R4` | 382 / **1, 91, 192, 229, 344** | **394** / **1, 97, 198, 235, 350** | **+12** | ASE1 +6, ASE2 0, ASE3 0, ASE4 0, ASE5 +6 |

Los Δ por ASE son **derivados** (no citados en el plan): R1 ASE1 `4..121` (118) vs v2 `4..107` (104); R4 ASE1/ASE5 +6 = el bloque por empresa `A=ENERBIT` que estrena la v3 (§0.4-c). Todos cuadran con el Δ total de §0.4-b, sin residuo.

### 1.4 Superconjunto estricto de etiquetas (§0.4-a)

Comparación de **multiconjunto** de la tupla `(A,B,C,D,E)` por **(bloque ASE, sección, subsección)** — sección y subsección tipadas por su **fila de cierre** (ver §2). Resultado en la v3 frente a la v2:

| Hoja | Claves v2 | Claves v3 | **Faltantes (v2 ⊄ v3)** | Claves nuevas en v3 |
|---|---|---|---|---|
| `Reporte Componentes R1` | 320 | 374 | **0** | 83 |
| `Rem. Anticipos R2` | 171 | 191 | **0** | 33 |
| `Reversion Pagos R4` | 94 | 96 | **0** | 12 |

⇒ **v3 es superconjunto estricto de v2** en las tres hojas: ninguna etiqueta que la v2 declaraba desapareció; la v3 añade capacidad (canales `3/4/6/9`, `E/ENERBIT` de ASE1, bloque `A=ENERBIT` de R4, grupo `A/EAB` del 2.º cuadro de R2).

### 1.5 Zonas nuevas: cero celdas `<f>` (§0.4-c)

| Zona medida (v3) | Celdas `<f>` |
|---|---|
| `Reporte Componentes R1` filas 122-150 (ASE2 estrena canales) | **0** |
| `Reporte Componentes R1` filas 55-72 (ASE1 estrena `E/ENERBIT`) | **0** |
| `Rem. Anticipos R2` filas 75-106 (ASE2 estrena canales) | **0** |
| `Reversion Pagos R4` filas 14-19 (ASE1 bloque `ENERBIT`) | **0** |
| `Reversion Pagos R4` filas 350-356 (ASE5 bloque `ENERBIT`) | **0** |

⇒ la regla «solo valores, nunca fórmulas» es **satisfacible** sobre la v3 sin tocar la plantilla.

---

## 2. T0.2 — Vocabulario de cierres por hoja (S1.2) — **insumo directo de H1**

Regla transversal medida: en las tres hojas **la fila de cierre es TRAILER** — cierra lo que está **arriba**, nunca encabeza. La **primera fila de datos** del bloque no lleva cierre: es el primer renglón-concepto tras la cabecera de bloque (cabecera + 2 en las tres hojas, porque la fila intermedia es el encabezado de columnas).

### 2.1 `Reporte Componentes R1` — 4 niveles (A=sección, B=subsección, C=fuente, D=canal, E=concepto)

| Nivel | Firma de la fila de cierre | Cierra | Ejemplos v3 (ASE1) |
|---|---|---|---|
| **Grupo CANAL** | `D=<canal> ∧ E='Total'` con **A,B,C vacías** | las filas-concepto del canal inmediatamente arriba | `7` (`D=E`), `10` (`H`), `12` (`O`), `14` (`T`) |
| **Grupo EMPRESA (fuente)** | `C=<empresa> ∧ D='Total'` con **A,B vacías** | **todos los canales que la preceden** (trailer, §0.4-d) | `8` `ENEL`, `15` `OCCIDENTE`, `21` `ENEL`, `25` `ENERBIT` |
| **SUBSECCIÓN** | `B=<Mes \| Aplicacion nuevos x reversion> ∧ C='Total'` con **A vacía** | el cuadro (sub-bloque) que la precede | `16` `Mes`, `36` `Aplicacion nuevos x reversion`, `55` `Mes`, `70` `Aplicacion…`, `84` `Mes` |
| **SECCIÓN** | `A=<AFaseo \| Componente \| Subs/Cont> ∧ B='Total'` con **C,D,E vacías** | las subsecciones que la preceden | `17` `AFaseo`, `56` `Componente`, `85` `Subs/Cont` |
| **BLOQUE / gran total** | `A='Total'` con **B,C,D,E vacías** | todo el bloque del ASE | `86` |
| **Primera fila de datos** | — | — | `6` (cabecera `4` + 2) |

Subsecciones por bloque en la v3 (**5 en los 5 bloques**, siempre la misma forma `Mes, Aplicacion, Mes, Aplicacion, Mes`):

| ASE | Subsecciones (fila@rótulo) | Secciones (fila@rótulo) | Gran total | 1.ª fila datos |
|---|---|---|---|---|
| 1 PROMOAMBIENTAL | 16 Mes · 36 Aplicacion · 55 Mes · 70 Aplicacion · 84 Mes | 17 AFaseo · 56 Componente · 85 Subs/Cont | 86 | 6 |
| 2 LIME | 149 · 187 · 224 · 251 · 277 | 150 · 225 · 278 | 279 | 124 |
| 3 CIUDAD LIMPIA | 368 · 384 · 399 · 411 · 422 | 369 · 400 · 423 | 424 | 358 |
| 4 BOGOTA LIMPIA | 544 · 582 · 619 · 646 · 672 | 545 · 620 · 673 | 674 | 519 |
| 5 AREA LIMPIA | 737 · 757 · 776 · 791 · 805 | 738 · 777 · 806 | 807 | 727 |

Cierres de grupo por bloque (v3, filas): **ASE1** `8,15,21,25,35,40,44,54,59,62,69,73,76,83` (14) · **ASE2** `126,141,148,154,176,186,191,213,223,228,243,250,254,269,276` (14) · **ASE3** `360,367,373,383,388,398,403,410,414,421` (10) · **ASE4** `521,536,543,549,571,581,586,608,618,623,638,645,649,664,671` (14) · **ASE5** `729,736,742,746,756,761,765,775,780,783,790,794,797,804` (14). Cierres de **canal** (rango cabecera…gran total): **24 / 55 / 20 / 55 / 24**.

### 2.2 `Rem. Anticipos R2` — 3 niveles (A=sección, B=empresa, C=canal, D=concepto, E=valor)

| Nivel | Firma de la fila de cierre | Cierra | Ejemplos v3 (ASE1) |
|---|---|---|---|
| **Grupo CANAL** | `C=<canal> ∧ D='Total'` con **A,B vacías** | las filas-concepto del canal (arriba) | `5` (`A`), `9` (`E`), `13` (`E`), `17` (`H`), `20` (`O`), `23` (`T`) |
| **Grupo EMPRESA** | `B=<empresa> ∧ C='Total'` con **A vacía** | **los canales que la preceden** (trailer) | `6` `EAB`, `10` `ENEL`, `14` `ENERBIT`, `24` `OCCIDENTE` |
| **SECCIÓN** | `A=<Componente \| Subs/Cont> ∧ B='Total'` con **C,D vacías** | el cuadro que la precede | `25` `Componente`, `42` `Subs/Cont` |
| **BLOQUE / gran total** | `A='Total'` con **B vacía** | todo el bloque | `43` |
| **Primera fila de datos** | — | — | `3` (cabecera `1` + 2) |

- **NO hay dimensión de subsección en R2.** Las dos mitades del bloque son **secciones distintas** (`Componente` = anticipos, `Subs/Cont` = subsidio/contribución), no subsecciones de una misma.
- Secciones por bloque (v3): ASE1 `25, 42` · ASE2 `117, 146` · ASE3 `216, 230` · ASE4 `359, 388` · ASE5 `465, 482`. Gran total: `43, 147, 231, 389, 483`. 1.ª fila de datos: `3, 77, 198, 319, 443`.
- Cierres de grupo por bloque: ASE1 `6,10,14,24,28,31,34,41` · ASE2 `80,84,106,116,120,123,138,145` · ASE3 `201,205,215,219,222,229` · ASE4 `322,326,348,358,362,365,380,387` · ASE5 `446,450,454,464,468,471,474,481`. Cierres de canal: 12 / 24 / 10 / 24 / 12.

### 2.3 `Reversion Pagos R4` — 2 niveles, **sin dimensión de canal** (A=empresa, B=sección, C=concepto, D..=valores)

| Nivel | Firma de la fila de cierre | Cierra | Ejemplos v3 (ASE1) |
|---|---|---|---|
| **Grupo EMPRESA** | `A=<empresa> ∧ B='Total'` | las secciones que la preceden (trailer) | `8` `ENEL`, `14` `ENERBIT`, `20` `OCCIDENTE` |
| **SECCIÓN** | `B=<Componente \| Subs/Cont> ∧ C='Total'` con **A vacía** | el cuadro concepto/subsidio que la precede | `5` `Componente`, `7` `Subs/Cont`, `11`, `13`, `17`, `19` |
| **BLOQUE / gran total** | `A='Total'` con **B vacía** | todo el bloque | `21` |
| **Primera fila de datos** | — | — | `3` (cabecera `1` + 2) |

- **El rótulo de grupo (columna A) ES la empresa**: no hay composición empresa/canal como en R2 (§0.4-d). Cero cierres de canal en los 5 bloques (0/0/0/0/0).
- Secciones por bloque: ASE1 `5,7,11,13,17,19` · ASE2 `101,103,107,109,113,115` · ASE3 `202,204,208,210` · ASE4 `239,241,245,247,251,253` · ASE5 `354,356,360,362,366,368`. Empresas: ASE1 `8,14,20` · ASE2 `104,110,116` · ASE3 `205,211` · ASE4 `242,248,254` · ASE5 `357,363,369`. Gran total: `21, 117, 212, 255, 370`. 1.ª fila de datos: `3, 99, 200, 237, 352`.

### 2.4 Confirmación aritmética del trailer (§0.4-d) con la fuente real

**Método:** en `Recaudoporcomponente` de octubre-1, para cada cierre `C=<empresa> ∧ D='Total'` se suma (columna `F`) la serie de cierres de canal `D=<canal> ∧ E='Total'` que lo preceden dentro del mismo cuadro, y se compara al centavo.

| ASE | Cierres de empresa | Verificados (Δ ≤ 0,005) |
|---|---|---|
| 1 PROMOAMBIENTAL | 8 | **8/8** |
| 2 LIME | 10 | **10/10** |
| 3 CIUDAD LIMPIA | 5 | **5/5** |
| 4 BOGOTA LIMPIA | 12 | **12/12** |
| 5 AREA LIMPIA | 6 | **6/6** |
| **Total** | **42** | **42/42** |

**Caso citado célula por célula** (ASE2 LIME, octubre-1, `Recaudoporcomponente_…_20261016174350151.xlsx`):

| Fila | A | B | C | D | E | F (valor) | Lectura |
|---|---|---|---|---|---|---|---|
| 15 | | | | `E` | `Total` | 11.087.940,08 | cierre canal `E` |
| 16 | | | `ENEL` | `Total` | | 11.087.940,08 | **cierre empresa ENEL = E** |
| 18 | | | | `H` | `Total` | 21.770,00 | cierre canal `H` |
| 21 | | | | `O` | `Total` | 644.372,54 | cierre canal `O` |
| 22 | | | `OCCIDENTE` | `Total` | | 666.142,54 | **cierre empresa OCCIDENTE = H + O** |
| 23 | | `Aplicacion nuevos x reversion` | `Total` | | | 11.754.082,62 | **cierre subsección = ENEL + OCCIDENTE** |

⇒ tres niveles de trailer al centavo (canal → empresa → subsección). Queda **confirmada** la semántica de §0.4-d y **descartado** el modelo «la fila de empresa encabeza» (un encabezado no puede ser la suma de lo que le sigue en un reporte acumulativo).

### 2.5 Trampa medida: el cuadro-resumen reutiliza el vocabulario de datos

Debajo del gran total de cada bloque hay un **cuadro-resumen** (OPORTUNO / EXTEMPORANEO por empresa) que **reutiliza parcialmente** las firmas del bloque de datos:

| Hoja | Firma del resumen | ¿Colisiona con el vocabulario de datos? |
|---|---|---|
| R1 | `D∈{OPORTUNO, EXTEMPORANEO} ∧ E='Total'` (C = empresa o `TOTAL`) | **No** (la de canal es `D=<canal> ∧ E='Total'` y el canal nunca es `OPORTUNO`) |
| R2 | `C='OPORTUNO' ∧ D='Total'` | **SÍ** — es exactamente la firma de un cierre de canal |
| R4 | `B='OPORTUNO' ∧ C='Total'` | **SÍ** — es exactamente la firma de un cierre de sección |

**Consecuencia vinculante para H1 (D-1.1):** el bloque de datos se delimita **[cabecera de bloque … fila del gran total]**; el resumen queda **fuera**. Una regla que escanee «hasta la cabecera del ASE siguiente» o «hasta el final del `dimension`» tomaría 6 resúmenes por hoja como cierres de grupo/sección y rompería la cardinalidad. El fixture lo asevera (`T0_2_ElCuadroResumenReutilizaElVocabulario…`).

**Tensión medida con la letra de D-1.1 (requiere OK del Ingeniero, no es discrepancia con §0.4):** D-1.1 dice que «el inicio del bloque es la fila siguiente al gran total del bloque anterior». Eso es cierto **solo si** ese rango excluye el cuadro-resumen; medido, la distancia entre el gran total de un bloque y la cabecera del siguiente es **36 filas en R1** (86→122), **32 en R2** (43→75) y **76 en R4** (21→97), y en ese hueco vive el resumen. En R2/R4 ese hueco contiene firmas de cierre **idénticas** a las del bloque de datos (§2.5) ⇒ si el delimitador arranca en «gran total anterior + 1» sin cortar en el gran total, leerá cierres falsos. **Refinamiento propuesto (probado por bytes):** el inicio de bloque se fija por la **cabecera de bloque** — la fila cuyo `B` es el rótulo de ASE, que es **dato leído del propio workbook**, no un literal ni un índice — y el fin por su **gran total**. Así no se congela ninguna fila y la regla funciona en las tres hojas.

### 2.6 Caso 1.ª quincena / subsecciones ausentes (el delimitador NO congela cardinalidades)

| Medición | Valor |
|---|---|
| Plantilla v3, R1: subsecciones por bloque | **5 en 5/5 bloques** (2 `Aplicacion nuevos x reversion` + 3 `Mes`) |
| Fuente **julio-1** (Q1) ASE2 LIME: cierres de subsección | **3** (`Mes`, `Mes`, `Mes`) — **0** filas `Aplicacion nuevos x reversion` |
| Fuente **octubre-1** ASE2 LIME: cierres de subsección | **5** (2 `Aplicacion` + 3 `Mes`) |
| Secciones de la fuente en ambas quincenas | **3** (`AFaseo`, `Componente`, `Subs/Cont`) |

⇒ la ausencia es de **subsección**, no de sección; es la forma documentada de Q1 (Planes 41/45/46: la fuente Q1 no trae la frontera `Aplicacion`). El delimitador debe leer **cuántas** subsecciones trae la fuente y no asumir 5/o 3; la plantilla declara 5 slots fijos y la fuente que no los trae cae en el **0 del template** (doctrina Plan 50).

---

## 3. T0.3 — ¿Los destinos CONSOLIDADO/DetRetri son derivables por etiqueta? (S1.3)

**Veredicto:** **derivable-por-etiqueta en las 36 celdas** (6 tablas × 5 ASE + fila de total + 5 destinos DetRetri). Ninguna celda exige literal y **ningún veredicto se apoya en «es estable»**.

**Regla que resuelve sin literales** (probada en el fixture):

1. **Tabla:** fila cuyo rótulo en **F** coincide con el título **y** que declara un encabezado `B='ASE' / C='Concesionarios'` **debajo**.
2. **Fila del ASE:** primera fila bajo ese encabezado con `B=<id 1..5>` y `C=<nombre largo del ASE>`.
3. **Celda:** columna de valor de la tabla (**D**).

### 3.1 Las 6 tablas del `CONSOLIDADO_TOTAL RECAUDO` (v3)

| Título (rótulo en `F`) | Fila título | Encabezado `B='ASE'` | Filas ASE (`B=1..5`) | Fila total (`C='TOTAL'`) | Celdas destino |
|---|---|---|---|---|---|
| `R1 - RECAUDO OPORTUNO` | 6 | 8 | 9-13 | 14 | `D9:D13` (+ total `D14`) |
| `R2 - ANTICIPOS A USUARIOS` | 25 | 27 | 28-32 | 33 | `D28:D32` (+ `D33`) |
| `R3 - RECAUDO EXTEMPORANEO` | 44 | 46 | 47-51 | 52 | `D47:D51` (+ `D52`) |
| `R4 - REVERSIÓN DE PAGOS` | 63 | 65 | 66-70 | 71 | `D66:D70` (+ `D71`) |
| `R5 - REMUNERACIÓN SALDOS A FAVOR POR NOTA` | 82 | 84 | 85-89 | 90 | `D85:D89` (+ `D90`) |
| `REPORTE CONSOLIDADO TOTAL DE RECAUDO` | 101 | 103 | 104-108 | 109 | `D104:D108` (+ `D109`) |

**Identidad de cada fila (columna C, v3)** — el rótulo existe y es único *dentro de su tabla*:

| `B` | `C` |
|---|---|
| 1 | `PROMOAMBIENTAL DISTRITO S.A.S. ESP.` |
| 2 | `LIMPIEZA METROPOLITANA S.A. ESP.` |
| 3 | `CIUDAD LIMPIA BOGOTA S.A. ESP.` |
| 4 | `BOGOTA LIMPIA S.A.S. ESP.` |
| 5 | `AREA LIMPIA DISTRITO CAPITAL S.A.S. ESP.` |

**Veredicto por celda:**

| Celda(s) | Rótulo que la identifica | Veredicto |
|---|---|---|
| `D9`–`D13` | tabla `R1 - RECAUDO OPORTUNO` + `B=1..5` + `C=<nombre>` (fila) | **derivable-por-etiqueta** |
| `D28`–`D32` | tabla `R2 - ANTICIPOS A USUARIOS` + `B=1..5` + `C=<nombre>` | **derivable-por-etiqueta** (⚠ el rótulo aparece 2 veces en `F`; ver §3.3) |
| `D47`–`D51` | tabla `R3 - RECAUDO EXTEMPORANEO` + `B=1..5` + `C=<nombre>` | **derivable-por-etiqueta** |
| `D66`–`D70` | tabla `R4 - REVERSIÓN DE PAGOS` + `B=1..5` + `C=<nombre>` | **derivable-por-etiqueta** |
| `D85`–`D89` | tabla `R5 - REMUNERACIÓN SALDOS A FAVOR POR NOTA` + `B=1..5` + `C=<nombre>` | **derivable-por-etiqueta** |
| `D104`–`D108` | tabla `REPORTE CONSOLIDADO TOTAL DE RECAUDO` + `B=1..5` + `C=<nombre>` | **derivable-por-etiqueta** |
| `D109` | misma tabla, fila de cierre `C='TOTAL'` | **derivable-por-etiqueta** |

⇒ La resolución por etiqueta reproduce **exactamente** las 30 celdas que hoy declara el mapa congelado `WorkbookLeafCellMapQ2.ConsolidadoProtected` (`:224`-`:257`) y las 5 de `DetRetriDestinoPorAse` (`:264`).

### 3.2 Naturaleza de las celdas destino (para que H2 verifique identidad, no coordenada)

| Celda | Naturaleza en la v3 |
|---|---|
| `D9:D13`, `D28:D32`, `D47:D51`, `D66:D70`, `D85:D89` | **`<f>`** (referencia a `R1!F…` / `R2!E…` / `R4!D…` / `AJUSTES!D…`) |
| `D104`, `D105`, `D107`, `D108` | **`<f>`** |
| `D106` | **`<f t="shared" si="47"/>`** — seguidora del maestro `D104` (`ref="D104:F106"`), texto vacío |
| `D109` | **`<f t="shared" ref="D109:N109" si="61">SUM(D104:D108)</f>`** |

⇒ el destino es **identificable por etiqueta** y su **protección** se verifica por **rol estructural** (la celda encontrada es la del cierre/tabla exigida), nunca por comparar la coordenada literal ni fragmentos congelados.

### 3.3 Colisión de rótulo medida (y cómo se desambigua)

`R2 - ANTICIPOS A USUARIOS` aparece **dos veces** en la columna `F` de la v3: **fila 25** (título de tabla) y **fila 141** (rótulo de **columna** de la tabla `VALIDACION DEL RECAUDO VERSUS REMUNERACION`). El título de tabla se distingue porque **declara el encabezado `B='ASE' / C='Concesionarios'` debajo** (`27`); la fila `141` no declara ninguno ⇒ no es tabla. Es la **única** colisión de las 6 y queda cubierta por la regla de §3.1 (no por un caso particular: la regla es estructural y el fixture asevera que las 6 resuelven a una única fila).

### 3.4 `DetRetri` — `DetRetriDestinoPorAse` (`WorkbookLeafCellMapQ2.cs:264`)

| Elemento | Medición (v3, hoja `DetRetri2026092`) |
|---|---|
| Encabezado de tabla | **fila 8**, `B='ASE'`, `C='CONCESIONARIOS'`, `D='RECAUDO TOTAL'` — **único** en la hoja |
| Filas por ASE | 9-13 (`B=1..5`, `C=<nombre largo>`) |
| Fila de total | 14 (`C='Total'`) |
| Naturaleza de `D9:D13` | **literales `<v>0</v>`** (destinos de escritura; no son fórmula) |
| Literal del mapa | ASE1..5 → `D9`, `D10`, `D11`, `D12`, `D13` = **exactamente** lo que resuelve la etiqueta |

**Identidad de cada fila (columna C, v3)** — ⚠ **ASE3 va SIN el punto final** (dato medido, **hallazgo H-59-4**):

| `B` | `C` en `DetRetri` | `C` en `CONSOLIDADO` |
|---|---|---|
| 1 | `PROMOAMBIENTAL DISTRITO S.A.S. ESP.` | idéntico |
| 2 | `LIMPIEZA METROPOLITANA S.A. ESP.` | idéntico |
| 3 | `CIUDAD LIMPIA BOGOTA S.A. ESP` ← **sin punto** | `CIUDAD LIMPIA BOGOTA S.A. ESP.` |
| 4 | `BOGOTA LIMPIA S.A.S. ESP.` | idéntico |
| 5 | `AREA LIMPIA DISTRITO CAPITAL S.A.S. ESP.` | idéntico |

**Veredicto:** `D9:D13` **derivable-por-etiqueta** (encabezado único + `B`/`C` por fila). `D14` (`DetRetriTotal`) también: fila con `C='Total'` de esa misma tabla. La identidad de fila debe resolverse por **`B` (id numérico) + rótulo leído del propio workbook**, no por igualdad exacta contra una cadena de otra hoja (H-59-4).

---

## 4. T0.4 — Grupo `canal A` / empresa `EAB` del 2.º cuadro de R2 (S1.4)

### 4.1 Mapa etiqueta→celda en la v3 (los 5 ASE, 2 grupos cada uno)

| ASE | 1.er cuadro (sección `Componente`) | 2.º cuadro (sección `Subs/Cont`) |
|---|---|---|
| 1 PROMOAMBIENTAL | cierre `C=A/D=Total` en **5** → cierre `B=EAB/C=Total` en **6** (sección cierra en 25) | **27** → **28** (sección cierra en 42) |
| 2 LIME | **79** → **80** (117) | **119** → **120** (146) |
| 3 CIUDAD LIMPIA | **200** → **201** (216) | **218** → **219** (230) |
| 4 BOGOTA LIMPIA | **321** → **322** (359) | **361** → **362** (388) |
| 5 AREA LIMPIA | **445** → **446** (465) | **467** → **468** (482) |

### 4.2 Misma ubicación en la v2 (el hueco de §0.5)

| ASE | cierre `C=A/D=Total` | cierre `B=EAB` | ¿2.º cuadro lo declara? |
|---|---|---|---|
| 1 | **5** | **6** | **NO** |
| 2 | **69** | **70** | **NO** |
| 3 | **177** | **178** | **NO** |
| 4 | **295** | **296** | **NO** |
| 5 | **406** | **407** | **NO** |

⇒ en la v2 el bloque declara **un solo** grupo `canal A / EAB` (el del 1.er cuadro). El `−18.613,87` del 2.º cuadro de ASE2 **no tenía destino**: es exactamente el pendiente **P1 del Plan 58** que la v3 resuelve declarando el grupo en los 5 ASE.

### 4.3 `R2!E285` — dejó de ser fórmula en la v3

| Versión | `<c r="E285">` cruda |
|---|---|
| v2 | `<c r="E285" s="93"><f>E192+E203-K192</f><v>0</v></c>` |
| **v3** | `<c r="E285" s="90"/>` — **sin `<f>` y sin `<v>`** |
| v2 `K285` | `<c r="K285" s="93"><f t="shared" si="9"/><v>0</v></c>` |
| **v3 `K285`** | `<c r="K285" s="90"/>` — sin `<f>` |

⇒ **confirmado:** el bloqueo `R-MAPA-V2-R2-ASE4` (celda destino fórmula) **no existe** en la v3.

### 4.4 Fuente `RerpoteDetalleSaldosaFavor` de octubre-1

| ASE | Grupos `canal A` | Filas | Valores (`E` del cierre `B=EAB`) |
|---|---|---|---|
| **2 LIME** | **2** | canal `6` → EAB `7`; canal `35` → EAB `36` | **+149.137,87** (1.er cuadro) y **−18.613,87** (2.º cuadro) |
| 1 PROMOAMBIENTAL | **0** | — | — |
| 3 CIUDAD LIMPIA | **0** | — | — |
| 4 BOGOTA LIMPIA | **0** | — | — |
| 5 AREA LIMPIA | **0** | — | — |

⇒ coincide con §0.5 (`+149.137,87` / `−18.613,87`; sólo ASE2 trae grupos `A/EAB` en ese período).

---

## 5. Hallazgos (no contradicen §0.4; se reportan como insumo de H1/H2)

### H-59-1 — En la v3, el literal `E285` ya no bloquea… pero tampoco cae en el bloque de ASE4

| Dato | Medición |
|---|---|
| Bloque ASE3 en R2 v3 | filas **196..316**; gran total **231**; cuadro-resumen desde **299** |
| Bloque ASE4 en R2 v3 | filas **317..440**; gran total **389** |
| Ubicación de la fila 285 | dentro del bloque de **ASE3**, entre el gran total (231) y el resumen (299) → celda **vacía** |
| Destino real del cierre `EAB` del 2.º cuadro de ASE4 | fila **362** (`E362`, literal `<v>0</v>`, escribible) |

**Lectura:** §0.5 dice «en la v3 esa coordenada ya no es fórmula» — **cierto** (E285 no tiene `<f>`). La precisión adicional es que el literal **tampoco es el destino del ASE4**: escribir en `E285` sería una colocación silenciosa en el bloque equivocado. **No requiere decisión**: refuerza el mandato de H2 (destino **derivado por etiqueta**, jamás el literal aunque ya no esté protegido). Se reporta porque cambia la razón del retiro: no es «ya no bloquea» sino «ya no sirve».

### H-59-2 — El cuadro-resumen reutiliza las firmas de cierre en R2 y R4

Documentado en §2.5. **Insumo vinculante para H1:** el bloque de datos termina en el **gran total** (`A='Total' ∧ B` vacío). Sin esa frontera, R2 tendría 6 «cierres de canal» falsos por hoja y R4 6 «cierres de sección» falsos. **Añadido:** la distancia gran-total→cabecera siguiente es 36 filas (R1), 32 (R2) y 76 (R4); el inicio de bloque debe fijarse por la **cabecera (B=rótulo de ASE)**, no por «gran total anterior + 1» (§2.5, refinamiento de D-1.1 sujeto a OK del Ingeniero).

### H-59-3 — El rótulo `R2 - ANTICIPOS A USUARIOS` no es único en `CONSOLIDADO`

Documentado en §3.3. **Insumo para H2:** resolver la tabla por (título **+** encabezado `B='ASE'` debajo), no por rótulo suelto.

### H-59-4 — El nombre largo del ASE3 difiere entre `DetRetri` y `CONSOLIDADO` por el punto final

| Hoja | ASE3 (columna C) |
|---|---|
| `CONSOLIDADO_TOTAL RECAUDO` | `CIUDAD LIMPIA BOGOTA S.A. ESP.` |
| `DetRetri2026092` | `CIUDAD LIMPIA BOGOTA S.A. ESP` (**sin** punto) |

**Insumo para H2:** un resolvedor que compare el rótulo por **igualdad exacta** entre hojas fallaría en ASE3; la identidad se resuelve por el **id numérico** (`B`) + el rótulo **leído de la propia hoja** (el catálogo abierto de sinónimos no aplica: son nombres de ASE, no de empresa de recaudo). El fixture fija ambos textos exactos para que la divergencia no se «corrija» en silencio.

---

## 6. T0.5 — Gates y línea base (S1.5, §3-S0)

### 6.1 G-SHA (plantillas intactas)

| Momento | v1 | v2 | v3 |
|---|---|---|---|
| Antes de H0 | `3086AC31…C5214` | `0F192AD2…7F2DAB` | `8199F864…2B1295` |
| Después de H0 | `3086AC31…C5214` | `0F192AD2…7F2DAB` | `8199F864…2B1295` |

### 6.2 G-FORMULA (cero `<f>` escritas) y G-H0-READONLY

- El fixture **solo lee**: `ZipFile.OpenRead` + `XDocument.Load`. Cero `CellFormula`, cero `InsertRow`/`DeleteRow`, cero `File.Write*`.
- `Remuneracion.Core/` y `Remuneracion.Infrastructure/` **no se tocan**: el diff de H0 es **vacío** (ver §6.4).
- Los únicos archivos nuevos de H0 son el fixture `Remuneracion.IntegrationTests/GeometriaDerivadaV3T0Plan59Tests.cs` y este documento.

### 6.3 G-BUILD

`dotnet build Remuneracion.WinForms/Remuneracion.WinForms.slnx` → **0 warnings, 0 errores**.

### 6.4 G-SUITE (set-diff de rojos)

**Aviso de base desactualizada (declarado, no silenciado):** el trx base de la puerta (`p58-final.trx`, **2026-10-23 21:10**) quedó desactualizado por el commit **`88b3acb` "Actualización de insumos Julio-2" (2026-10-26 07:46:50)**, **posterior** a esa base. Ese commit reemplazó archivos **trackeados** de `Docs/Prueba Julio-2/Insumos/` (46 en `git ls-files`): `Recaudoporcomponente_…_20268465657830.xlsx` → `…_2026831522272.xlsx`; `Conjunta ENEL-072026.xlsx` → `Conjunta ENEL-072026vf.xlsx`; `R4-BalanceSubsidioyContribuciones-Optimiz…` → `R4-BalanceSubsidioyContribuciones_to_date…`.

| Corrida | Qué mide | Total | Passed | Failed | Skipped | Archivo |
|---|---|---|---|---|---|---|
| **A** Base de la puerta (árbol 2026-10-23) | referencia de entrada | 1558 | 1523 | **35** | 0 | `TestResults/p58-final.trx` |
| **B** Antes de H0 **re-medido hoy** (`--filter FullyQualifiedName!~GeometriaDerivadaV3T0Plan59`) | mismo árbol e insumos de `88b3acb`, **sin** el fixture | 1558 | 1518 | **40** | 0 | `TestResults/p59-h0-baseline-hoy.trx` |
| **C** Después de H0 (suite completa **con** el fixture) | efecto de H0 | **1581** | **1541** | **40** | 0 | `TestResults/p59-h0-final.trx` |

| Comparación | Altas | Bajas | Cambios de resultado | Lectura |
|---|---|---|---|---|
| **B vs C** (efecto de H0) | **+23** (solo los tests de este T0, todos verdes) | 0 | **0** | **H0 no cambia ningún resultado. Δ efectivo vacío.** |
| A vs B (deriva de insumos) | 6 verdes→rojos | 1 rojo→verde | 7 | causa única `88b3acb` (insumos Julio-2), **ajena a H0** |

**Rojos nuevos de A vs B (causa citada; los 6 fallan igual en aislamiento, sin correr ningún test del fixture):**

| # | Test | Firma del fallo |
|---|---|---|
| 1 | `AjustesSfTTests.BalanceOptimizado_Q2_ResuelvePorSegundoPrefijo` | `Assert.StartsWith`: archivo real `R4-BalanceSubsidioyContribuciones_to_date…` vs esperado `…-Optimiz…` |
| 2 | `BalanceScPermutaTests.JulioAse5Bce_ManualDivergeDeLaFuente_ConstanteDocumentada` | Δ `2.477.195.172,11` (InRange −0,5..0,5) |
| 3 | `BalanceScPermutaTests.JulioQ2_PermutaBce_SubsidioAD_ContribucionAE_VsFuente_E2E` | Δ `1.210.425.628,31` |
| 4 | `DetRetriDesgloseTrazableTests.JulioQ2_DesgloseTrazable_DetRetri_JyL_Y_DetVali_IyO` | ASE5 escrito `−82.959.024` (se esperaba divergencia del manual, T0a) |
| 5 | `RevisionPestanasLenteR1T0Plan47Tests.T0P47_Recaudo_EsPegadoReal_DesdeResumenMes_PosicionalHoyAlineado` | `FileNotFoundException`: `…/Conciliaciones/Conjunta ENEL-072026.xlsx` |
| 6 | `T0FrenteEspejoR1Ase5EnelTests.T0_Fuente_Ase5_TieneDosFilasEnel` | `FileNotFoundException`: `…/5-Área Limpia/Recaudoporcomponente_…_20268465657830.xlsx` |

**Recuperado (rojo→verde, misma causa):** `RevisionPestanasLenteR1T0Plan47Tests.T0P47_P6_R2_MatrizFilasYEmparejamiento_Congelada(ase: 4)`.
**Los otros 34 rojos de A siguen rojos en B y en C** (mismo conjunto, mismos nombres, sin cambios).

**Veredicto G-SUITE:** contra el árbol actual **0 rojos nuevos por H0** (B vs C sin cambios de resultado). El trx de la puerta **está stale** por un commit de insumos anterior a este H0 ⇒ **no se reclasifica como verde**: se escala al Ingeniero (§7.3).

### 6.5 Reproducción

```bash
# Sólo la evidencia de este T0 (23 tests)
dotnet test Remuneracion.IntegrationTests/Remuneracion.IntegrationTests.csproj \
  --filter "FullyQualifiedName~GeometriaDerivadaV3T0Plan59"

# Suite completa con trx (gate G-SUITE)
dotnet test Remuneracion.IntegrationTests/Remuneracion.IntegrationTests.csproj \
  --logger "trx;LogFileName=p59-h0-final.trx"
```

---

## 7. Conclusión de la puerta H0→H1

### 7.1 Estado de los criterios

1. **S1.1** ✅ §0.4 re-medido en disco sin una sola discrepancia (SHA, filas, cabeceras, superconjunto, 0 `<f>`).
2. **S1.2** ✅ vocabulario de cierres fijado **celda por celda** para R1/R2/R4, con la semántica **trailer** confirmada por aritmética (42/42) y el caso R4 sin canal documentado.
3. **S1.3** ✅ **36 celdas** con veredicto **derivable-por-etiqueta** (rótulo citado). Ninguna celda exige literal.
4. **S1.4** ✅ mapa `canal A / EAB` de los 5 ASE (v3 y v2), valores octubre-1 ASE2 y `E285` **sin `<f>`** en la v3.
5. **S1.5** ✅/⚠ SHAs intactos; **H0 no cambia ningún resultado** (B vs C: +23 altas verdes, 0 cambios). Reserva: el trx base de la puerta está desactualizado por `88b3acb` (insumos Julio-2) ⇒ 6 rojos nuevos **con causa citada y ajena a H0** (§6.4).

### 7.2 Restricciones de diseño probadas por bytes (para H1/H2)

1. El bloque termina en el **gran total** (`A='Total' ∧ B` vacío); el cuadro-resumen de debajo **reutiliza** las firmas de cierre en R2/R4 (§2.5).
2. El inicio de bloque se fija por la **cabecera (B=rótulo de ASE)**, no por «gran total anterior + 1» (refinamiento de D-1.1: ese hueco mide 36/32/76 filas y contiene el resumen).
3. La **cardinalidad de subsecciones** la decide la **fuente** (Q1 trae 3, Q2 trae 5; la plantilla declara 5 slots fijos) (§2.6).
4. El destino se **deriva por etiqueta** incluso cuando la celda literal ya no es fórmula (`E285`, H-59-1).
5. La tabla del CONSOLIDADO se resuelve por **título + encabezado `B='ASE'` debajo** (el título `R2 - ANTICIPOS A USUARIOS` no es único) (H-59-3).
6. La identidad de fila del ASE se resuelve por **`B` (id) + rótulo leído de la propia hoja** (el nombre largo difiere por un punto entre `DetRetri` y `CONSOLIDADO`) (H-59-4).

### 7.3 Puntos que requieren decisión del Ingeniero

| # | Punto | Opciones |
|---|---|---|
| **D-1** | **Línea base de suite stale**: `p58-final.trx` (2026-10-23) precede a `88b3acb` (2026-10-26 07:46, insumos Julio-2). 6 verdes→rojos y 1 rojo→verde por esa causa; **0 por H0**. | (a) aceptar el Δ con causa citada (§6.4) y abrir H1; **o** (b) fijar la línea base en la corrida **B** de este informe (1558/1518/40/0) y re-emitir el trx de la puerta |
| **D-2** | **Refinamiento de D-1.1**: el inicio de bloque por «fila siguiente al gran total anterior» choca con el cuadro-resumen en R2/R4 (§2.5). Se propone fijarlo por la **cabecera de bloque (B=rótulo)**. | ratificar el refinamiento antes de H1.3 (delimitación R2/R4) |

**H1 puede abrirse** en cuanto D-1 y D-2 queden resueltos; no hay frente del T0 que dependa de un literal ni de un caso particular.
